# Deposit Slip Extractor

Extracts data from bank deposit slip PDFs and writes it into a formatted Excel
sheet, flagging slips that need manual review.

## Problem
Over 1,000 pages of deposit slips had to be checked by hand to classify each
payment (installment, DPC, fees, etc.) and verify the amounts.

## What it does
- Reads each slip: name, member no., account, slip no., amount, description
- Maps description items to classification columns
- Skips simple slips where the amount already matches; writes the rest to Excel
- Calculates Total and Difference columns with Excel formulas
- Sends anything unreadable or unknown to a separate "unresolved" file instead of guessing
- Detects duplicate slips

## Tools
Python, PyMuPDF, OpenPyXL, regex, Decimal

## Usage
pip install -r requirements.txt
python extract_slips.py                 # opens a file picker
python extract_slips.py --pages 1-30    # test on a few pages
python extract_slips.py --dump 1        # view raw text of a page

## Output
- extracted.xlsx: slips needing review/entry
- unresolved_slips.xlsx: slips the script could not parse, with reasons

## Note
Sample files use dummy data only.
