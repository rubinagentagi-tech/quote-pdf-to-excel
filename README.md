# 📄 Quote PDF → Excel extractor

Turn supplier quotation PDFs into a structured spreadsheet: line items, part
numbers, quantities, unit prices, line totals, and a summary of the tax /
freight / total block. **Cross-checks the arithmetic and flags anomalies**
instead of trusting the vendor's math.

Works offline. No API keys. No OCR. Two files of pure Python.

## Why it exists

Buyers get quotes as PDFs and re-key them into RFQ sheets, spreadsheets and
ERP screens. Re-keying is where errors are born. This tool does the boring
part and keeps the *"does this line actually add up?"* check that humans are
supposed to do but skip.

## What it does

- **Extracts line items**: part no, qty, UOM, description, unit price, line total
- **Detects document metadata**: supplier, quote/ref number, date, currency
  (best-effort heuristics)
- **Flags anomalies** per row:
  - `math: 2 x 45.5 != 100` — qty × unit price disagrees with the line total
  - `no unit price` — total present but no unit price
  - `no quantity` — price present but no quantity
- **Captures the totals block**: Subtotal / Freight / HST / VAT / TOTAL lines
- **Two layout engines**: ruled-table detection (PyMuPDF `find_tables`) with a
  borderless aligned-column parser as fallback
- Outputs `.xlsx` (flagged rows shaded) and optional `.csv`

## Install & run

```bash
pip install -r requirements.txt        # or: pip install pymupdf openpyxl
python quote2xls.py quote.pdf [more.pdf ...] -o out --csv
```

```text
$ python quote2xls.py sample_quotes/HL-2026-0517_Harborlight-Marine.pdf
  HL-2026-0517_Harborlight-Marine.pdf: 5 items, 2 flagged | Harborlight Marine Supply Inc. | HL-2026-0517 | 2026-09-04
      ! row Bilge pump strainer, bronze, 2 inch NPT: math: 2 x 45.5 != 100
      ! row Stuffing box packing, 12mm square, graphite: no unit price
  -> wrote to /home/you/out
```

## Sample output

From `sample_quotes/Q-88213_Nordmarin.pdf` (a clean, unruled layout):

```csv
Part No,Qty,UOM,Description,Unit Price,Line Total,Flags
"1",8.0,"EA","Hydraulic hose assembly, 3/8 inch x 2500 psi, JIC fittings both ends",38.4,307.2,""
"2",10.0,"LB","Welding electrode E7018, 1/8 inch, low hydrogen",6.15,61.5,""
...
```

The `.xlsx` "Items" sheet also carries supplier / quote no / date / currency /
notes, and appends the totals block (Subtotal, Freight, HST, TOTAL) beneath the
items. Flagged rows are shaded.

## Try it

The `sample_quotes/` folder holds two realistic synthetic quotes:

- `HL-2026-0517_Harborlight-Marine.pdf` — ruled table, **contains a deliberate
  math error and a missing unit price** (see if the tool catches them)
- `Q-88213_Nordmarin.pdf` — borderless aligned layout, clean math

> **All sample content is fictitious**: companies, contacts, addresses, email
> domains (reserved `.example.com`) and part numbers. No real supplier quote
> or vendor data appears anywhere in this repository.

## Known limits (honest list)

- Text-based PDFs only. Scanned/image quotes need OCR first (out of scope).
- Works best on table-shaped quotes. Truly free-form layouts (no column
  alignment at all) will be reported, not guessed at.
- On borderless layouts a description line that wraps far outside its column
  band can occasionally lose a fragment. Ruled tables extract descriptions in
  full.
- Metadata is heuristic: supplier, ref and date are detected, not guaranteed.
  Always eyeball the Report sheet.

## License

MIT — see [LICENSE](LICENSE).
