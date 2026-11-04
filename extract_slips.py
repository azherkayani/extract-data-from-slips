#!/usr/bin/env python3
"""
I-16 deposit slip extractor
===========================
Reads a PDF of bank deposit slips and writes an Excel sheet in the layout of
"I-16 payment" (Name, Block, Apartment, Membership, Deposit Slip, Bank Account,
Deposit Slip Amount, DPC ... Membership Fee, Total, Difference).

Setup (once):      pip install pymupdf openpyxl
Run on everything: python extract_slips.py        (a window opens to pick the PDF)
Test on a few:     python extract_slips.py --pages 1-30
See raw text:      python extract_slips.py --dump 1
(Typing a PDF path after the script name skips the pick-a-file window.)

Outputs (saved in the same folder as the PDF you pick)
  I-16_payment_extracted.xlsx  slips that need an entry (more than one classification,
                               or the classification total differs from the slip amount)
  unresolved_slips.xlsx        slips the script could not understand (slip no., page, reason)

Rules applied
  * Description items are mapped to columns (see LABEL_TO_COL).
  * "Additional Charges (x)" is added under Installment.
  * A slip whose only description is "Additional Charges (0.00)" is a plain installment
    slip: the whole Amount goes under Installment.
  * Skipped = the description has exactly ONE item (nothing to add up) AND it equals the
    slip amount. Slips with two or more items (e.g. Installments + Additional Charges)
    are always written to the sheet.
  * "Other Charges (x)" and "Membership Fee (x)" have their own columns and are ALWAYS
    written to the sheet, even when the amount equals the slip amount.
  * Bank Account = last 3 digits of the A/C number (1490106 -> 106, 0002-2008938604 -> 604).
  * Anything with an unknown label, unreadable field or leftover text is NOT guessed;
    it goes to unresolved_slips.xlsx.
"""
import argparse
import os
import re
import sys
from collections import defaultdict
from decimal import Decimal

try:
    import pymupdf
except ImportError:  # older PyMuPDF
    import fitz as pymupdf

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side

# description label (lower case) -> column letter in the sheet
LABEL_TO_COL = {
    "dpc": "I",
    "possession charges": "J",
    "watch and ward charges": "K",
    "watch & ward charges": "K",
    "restoration fee": "L",
    "installment": "M",
    "installments": "M",
    "additional charges": "M",  # rule: Additional Charges are written under Installment
    "down payment": "N",
    "parking fee": "O",
    "other charges": "P",
    "membership fee": "Q",
}
CLASS_COLS = ["I", "J", "K", "L", "M", "N", "O", "P", "Q"]  # DPC .. Membership Fee; R = Total, S = Difference
ALWAYS_WRITE = {"P", "Q"}  # Other Charges / Membership Fee slips are never skipped

FIELD_RES = {
    "slip": re.compile(r"Slip\s*No\.?\s*([0-9][0-9-]*)", re.I),
    "account": re.compile(r"A/C\s*#\s*([0-9][0-9-]*)", re.I),
    "member": re.compile(r"Member\s*#\s*([0-9][0-9-]*)", re.I),
    "name": re.compile(r"\bName\b\s*(.+?)\s*Block\s*/\s*Category", re.I | re.S),
    "block": re.compile(r"Block\s*/\s*Category\s*(.+?)\s*Floor\s*/\s*Street\s*#", re.I | re.S),
    "unit": re.compile(r"Unit\s*#\.?\s*(\S+)", re.I),
    "amount": re.compile(r"Amount\s*:\s*([\d,]+(?:\.\d+)?)\s*(.*?)\s*In\s*Words", re.I | re.S),
}
ITEM_RE = re.compile(r"([A-Za-z][A-Za-z &/.\-]*?)\s*\(\s*([\d,]+(?:\.\d+)?)\s*\)")


class Unresolved(Exception):
    def __init__(self, reason, slip_no=None):
        super().__init__(reason)
        self.reason = reason
        self.slip_no = slip_no


def clean(s):
    return re.sub(r"\s+", " ", s).strip()


def money(s):
    return Decimal(s.replace(",", ""))


def numify(s):
    return int(s) if s.isdigit() else s


def parse_slip(chunk):
    m = FIELD_RES["slip"].search(chunk)
    slip = m.group(1) if m else None
    if not slip:
        raise Unresolved("could not read slip no.")
    found = {}
    for key in ("account", "member", "name", "block", "unit", "amount"):
        m = FIELD_RES[key].search(chunk)
        if not m:
            raise Unresolved(f"could not read {key}", slip)
        found[key] = m

    amount = money(found["amount"].group(1))
    desc = clean(found["amount"].group(2))

    # --- classification -------------------------------------------------
    items = [(clean(lbl).lower(), money(val)) for lbl, val in ITEM_RE.findall(desc)]
    leftover = re.sub(r"[\s,;]+", "", ITEM_RE.sub("", desc))
    if not items:
        raise Unresolved(f"no classification found in description: '{desc}'", slip)
    if leftover:
        raise Unresolved(f"unrecognised text in description: '{desc}'", slip)
    cols = defaultdict(Decimal)
    for label, value in items:
        col = LABEL_TO_COL.get(label)
        if col is None:
            raise Unresolved(f"unknown classification '{label}' in '{desc}'", slip)
        cols[col] += value
    if len(items) == 1 and items[0][0] == "additional charges" and items[0][1] == 0:
        cols = {"M": amount}  # plain installment slip
    total = sum(cols.values(), Decimal(0))
    skip = len(items) == 1 and total == amount and not (ALWAYS_WRITE & set(cols))

    # --- identity fields ------------------------------------------------
    member = found["member"].group(1)
    if re.fullmatch(r"\d{9}", member):  # 405000582 -> 405-000582
        member = f"{member[:3]}-{member[3:]}"
    acct = found["account"].group(1)
    bank = int(re.sub(r"\D", "", acct)[-3:])  # 1490106 -> 106, 0002-2008938604 -> 604

    return {
        "slip": slip,
        "name": clean(found["name"].group(1)),
        "block": numify(clean(found["block"].group(1))),
        "unit": numify(clean(found["unit"].group(1))),
        "member": member,
        "bank": bank,
        "amount": amount,
        "cols": dict(cols),
        "skip": skip,
        "sig": (clean(found["name"].group(1)), amount, desc),
    }


def iter_chunks(doc, first, last):
    """Yield (page_no, text_chunk, problem). One chunk per 'Slip No' on a page."""
    for pno in range(first - 1, last):
        text = doc[pno].get_text("text", sort=True)
        if not text.strip():
            yield pno + 1, None, "no text layer on page (scanned image? needs OCR)"
            continue
        chunks = [c for c in re.split(r"(?=Slip\s*No)", text, flags=re.I)
                  if re.search(r"Slip\s*No", c, re.I)]
        if not chunks:
            yield pno + 1, None, "no 'Slip No' found on page"
        for c in chunks:
            yield pno + 1, c, None


# ---------------------------------------------------------------- Excel output
ACC0 = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"??_);_(@_)'
ACC2 = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)'  # Difference: shows paisa (e.g. 0.04)


def write_main(rows, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "I-16"
    font = lambda **kw: Font(name="Arial", size=kw.pop("size", 12), **kw)
    thin = Side(style="thin")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.merge_cells("A1:R1")
    ws["A1"] = "I-16 Project – Payment Classification and Segregation"
    ws["A1"].font = font(size=16, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    headers = ["Sr.No", "Name", "Block", "Apartment", "Membership", "Deposit\n Silp",
               "Bank\n Account", "Deposit Slip Amount", "DPC", "Possession\n Charges",
               "Watch \n& Ward", "Restoration\n Fee", "Installment", "Down Payment",
               "Parking Fee", "Other Charges", "Membership Fee", "Total", "Difference"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(2, c, h)
        cell.font = font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = box
    ws.row_dimensions[2].height = 48

    for i, r in enumerate(rows, 1):
        row = 2 + i
        values = [i, r["name"], r["block"], r["unit"], r["member"], r["slip"], r["bank"],
                  float(r["amount"])]
        for c, v in enumerate(values, 1):
            ws.cell(row, c, v)
        for col in CLASS_COLS:
            v = r["cols"].get(col)
            if v:
                ws[f"{col}{row}"] = float(v)
        ws[f"R{row}"] = f"=SUM(I{row}:Q{row})"
        ws[f"S{row}"] = f"=H{row}-R{row}"
        for c in range(1, 20):
            cell = ws.cell(row, c)
            cell.font = font()
            cell.border = box
            if c in (1, 3, 4, 5, 6, 7):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c >= 8:
                cell.number_format = ACC2 if c == 19 else ACC0
        ws.cell(row, 19).font = font(bold=True)

    first, last = 3, 2 + len(rows)
    tot = last + 1
    ws.merge_cells(f"A{tot}:G{tot}")
    ws[f"A{tot}"] = "Total:"
    for col in ["H"] + CLASS_COLS:
        ws[f"{col}{tot}"] = f"=SUM({col}{first}:{col}{last})"
    ws[f"R{tot}"] = f"=SUM(I{tot}:Q{tot})"
    ws[f"S{tot}"] = f"=H{tot}-R{tot}"
    for c in range(1, 20):
        cell = ws.cell(tot, c)
        cell.font = font(bold=True)
        cell.border = box
        if c >= 8:
            cell.number_format = ACC2 if c == 19 else ACC0
    ws[f"A{tot}"].alignment = Alignment(horizontal="center")

    for col, w in {"A": 7, "B": 34, "C": 9, "D": 11, "E": 15, "F": 16, "G": 16, "H": 17,
                   "I": 14, "J": 14, "K": 13, "L": 13, "M": 15, "N": 15, "O": 14, "P": 14, "Q": 15, "R": 16, "S": 15}.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A3"
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


def write_unresolved(items, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Unresolved"
    for c, h in enumerate(["Slip No.", "PDF page", "Reason"], 1):
        cell = ws.cell(1, c, h)
        cell.font = Font(name="Arial", size=12, bold=True)
    for r, (slip, page, reason) in enumerate(items, 2):
        ws.cell(r, 1, slip or "(unreadable)")
        ws.cell(r, 2, page)
        ws.cell(r, 3, reason)
        for c in (1, 2, 3):
            ws.cell(r, c).font = Font(name="Arial", size=12)
    for col, w in {"A": 18, "B": 10, "C": 110}.items():
        ws.column_dimensions[col].width = w
    wb.save(path)


# ------------------------------------------------------------------------ main
def pick_pdf():
    """Open a file-picker window and return the chosen PDF path ('' if cancelled)."""
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError:
        return input("Type or paste the full path of the PDF: ").strip().strip('"')
    try:
        root = tk.Tk()
    except tk.TclError:  # no display available
        return input("Type or paste the full path of the PDF: ").strip().strip('"')
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.askopenfilename(
        title="Select the deposit slips PDF",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
    )
    root.destroy()
    return path


def main():
    ap = argparse.ArgumentParser(description="Extract I-16 deposit slips from a PDF into Excel")
    ap.add_argument("pdf", nargs="?", help="PDF file (optional: a file-picker window opens if omitted)")
    ap.add_argument("--pages", help="page range to process, e.g. 1-30 (default: all)")
    ap.add_argument("--dump", type=int, help="print the raw text of this page and exit")
    ap.add_argument("--out")
    ap.add_argument("--unresolved")
    a = ap.parse_args()

    if not a.pdf:
        a.pdf = pick_pdf()
        if not a.pdf:
            sys.exit("No file selected.")
    folder = os.path.dirname(os.path.abspath(a.pdf))
    a.out = a.out or os.path.join(folder, "I-16_payment_extracted.xlsx")
    a.unresolved = a.unresolved or os.path.join(folder, "unresolved_slips.xlsx")
    print(f"PDF: {a.pdf}")
    doc = pymupdf.open(a.pdf)
    if a.dump:
        print(doc[a.dump - 1].get_text("text", sort=True))
        return
    first, last = 1, len(doc)
    if a.pages:
        first, last = (int(x) for x in a.pages.split("-"))
        last = min(last, len(doc))

    rows, unresolved, seen = [], [], {}
    n_skipped = n_dupes = n_slips = 0
    for page, chunk, problem in iter_chunks(doc, first, last):
        if problem:
            unresolved.append((None, page, problem))
            continue
        n_slips += 1
        try:
            slip = parse_slip(chunk)
        except Unresolved as e:
            unresolved.append((e.slip_no, page, e.reason))
            continue
        if slip["slip"] in seen:
            if seen[slip["slip"]] == slip["sig"]:
                n_dupes += 1
            else:
                unresolved.append((slip["slip"], page, "same slip no. appears again with different data"))
            continue
        seen[slip["slip"]] = slip["sig"]
        if slip["skip"]:
            n_skipped += 1
        else:
            rows.append(slip)

    write_main(rows, a.out)
    if unresolved:
        write_unresolved(unresolved, a.unresolved)
    print(f"Pages {first}-{last}: {n_slips} slips read")
    print(f"  skipped (single classification, amount matches): {n_skipped}")
    print(f"  written to {a.out}: {len(rows)}")
    print(f"  unresolved -> {a.unresolved if unresolved else '(none)'}: {len(unresolved)}")
    if n_dupes:
        print(f"  exact duplicate slips ignored: {n_dupes}")


if __name__ == "__main__":
    try:
        main()
    finally:
        if len(sys.argv) == 1:  # started by double-click: keep the window open
            input("\nPress Enter to close...")
