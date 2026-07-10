# ACFR Digest (alpha)

Extracts General Fund figures from U.S. municipal Annual Comprehensive
Financial Reports (ACFRs) — 11 figures per report from the Balance Sheet and
the Statement of Revenues, Expenditures, and Changes in Fund Balances — into
a single Excel workbook.

**Alpha software.** Coverage on the audited 100-report corpus is 95.4% of
fields. The parser deliberately returns NOT FOUND rather than guessing; read
[KNOWN_LIMITATIONS.md](KNOWN_LIMITATIONS.md) before relying on the output.

## Web UI (drag and drop)

```bash
pip install .
streamlit run app.py
```

Drop PDFs (or a `.zip` of a folder — browsers can't upload folders directly)
into the uploader and download the results workbook. Files take 1–3 minutes
each; scanned/OCR-path files take longer.

## CLI

```bash
pipx install .
acfr-digest --dir /path/to/acfrs --output results.xlsx
acfr-digest report_a.pdf report_b.pdf --output results.xlsx --log run.log
acfr-digest --version
```

`acfr-digest` is a thin wrapper around the audited `parse_cafr.py` — output is
identical to `python3 parse_cafr.py` with the same arguments. Alongside the
workbook it writes `<output>.checkpoint.csv` row by row as crash insurance.

## Optional OCR support (scanned / broken-encoding PDFs)

```bash
brew install poppler tesseract        # macOS (Debian: poppler-utils tesseract-ocr)
pipx install '.[ocr]'                 # or: pip install '.[ocr]'
```

Without the OCR toolchain, scanned or mojibake PDFs return honest NOT FOUNDs.
With it, they're OCR-extracted and flagged "verify manually". The `poppler`
package is also recommended for native-text PDFs: it provides `pdffonts`,
which the parser uses to detect scanned documents up front.

## Output format

One row per PDF: entity name, fiscal year end, reporting units, Total Assets,
Total Liabilities, the five GASB 54 fund balance categories, Total Fund
Balances, Total Revenues, Total Expenditures, Total Other Financing Sources
(Uses), and extraction notes — plus USD-normalized columns. Missing figures
appear as NOT FOUND, never as a guess.

## Development

```bash
python3 tests/smoke.py    # regression smoke test (needs the local Set2 corpus)
```

License: [MIT](LICENSE)
