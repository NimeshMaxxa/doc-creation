# FFS invoice generator

Makes invoice PDFs in the same layout as INVOICE 1706 / 1707
(letterhead, item table, grand total, bank details, signature lines).

## One-time setup
Keep this whole folder together. Python must be installed (Anaconda is fine).
- Mac: the first time, right-click **Make Invoice.command** > Open > Open
  (macOS asks once because the file was downloaded).
- Windows: install Python from python.org (tick "Add Python to PATH").

## Making an invoice
1. Copy `invoice_template.xlsx` and rename it (e.g. `Ebony Oct.xlsx`).
2. Fill the yellow cells (customer, PO No, Quote No) and the item rows.
   - Leave **Invoice No** blank: the next number is used automatically
     (last used number is kept in `assets/last_invoice_no.txt`) and written
     back into your sheet, so re-running gives the same number.
   - Leave **Date** blank for today's date.
3. Close Excel. Mac: double-click **Make Invoice.command** and pick the .xlsx.
   Windows: drag the .xlsx onto **Make Invoice.bat**.
   The PDF appears in `output/` as `INVOICE <number>.pdf`.

## Item rows
| What you want | How to fill the row |
|---|---|
| Normal item | Description, Unit, Qty, Rate (Value = Qty x Rate, numbered 1, 2, 3...) |
| Bold section heading (e.g. REPLACEMENT) | Description only; numbering restarts after it |
| Lump sum (e.g. Transport Charges) | Description + Amount |
| Bold second line (e.g. "(Defective Device)") | put it in Remark |
| Force a number / no number | type it in Item, or `-` for none |

Grand Total is the sum of all values.

## Changing fixed details
Address, phone numbers and bank details are at the top of `make_invoice.py`
(the `COMPANY` block). The logo is `assets/logo.png`.
Fonts: uses Calibri if installed, otherwise the bundled Carlito (same widths).
