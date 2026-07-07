"""
Deep-dive examination: Baltimore BS detection and RevEx false positive analysis.
"""

import pdfplumber
import re
from pathlib import Path

PDF_DIR = Path(__file__).parent

def normalize(text):
    if not text:
        return ""
    text = text.lower()
    text = re.sub(r'[–—―]', '-', text)
    text = re.sub(r'[''""]', "'", text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

# ---- Baltimore: scan every page for BS signals ----
print("="*70)
print("BALTIMORE: scanning all pages for balance sheet signals")
print("="*70)

with pdfplumber.open(PDF_DIR / "Baltimore FY25 Financial Report.pdf") as pdf:
    print(f"Total pages: {len(pdf.pages)}\n")
    for i, page in enumerate(pdf.pages):
        text = page.extract_text() or ""
        norm = normalize(text)
        has_bs = 'balance sheet' in norm
        has_fund_bal = 'fund balance' in norm or 'fund balances' in norm
        has_gf = 'governmental fund' in norm or 'general fund' in norm
        if has_bs or (has_fund_bal and has_gf):
            print(f"  Page {i+1}: bs={has_bs}, fund_bal={has_fund_bal}, gf={has_gf}")
            # Print first 400 chars
            print(f"    Text: {text[:500]!r}")
            print()

print("\n" + "="*70)
print("BALTIMORE: all pages with 'balance' in title area (first 200 chars)")
print("="*70)

with pdfplumber.open(PDF_DIR / "Baltimore FY25 Financial Report.pdf") as pdf:
    for i, page in enumerate(pdf.pages):
        text = page.extract_text() or ""
        header = text[:300]
        norm_header = normalize(header)
        if 'balance' in norm_header:
            print(f"\n  Page {i+1}:")
            print(f"  {header[:300]!r}")

print("\n" + "="*70)
print("BALTIMORE: pages 40-60 first 600 chars each")
print("="*70)

with pdfplumber.open(PDF_DIR / "Baltimore FY25 Financial Report.pdf") as pdf:
    for i in range(39, 60):
        if i < len(pdf.pages):
            text = pdf.pages[i].extract_text() or ""
            print(f"\n--- Page {i+1} ---")
            print(text[:600])

print("\n" + "="*70)
print("ANNAPOLIS / GAITHERSBURG: RevEx false positive page analysis")
print("="*70)

for fname, pages in [
    ("Annapolis FY25 Financial Report.pdf", [38, 39]),
    ("Gaithersburg FY25 Financial Report.pdf", [44, 45])
]:
    print(f"\n--- {fname} ---")
    with pdfplumber.open(PDF_DIR / fname) as pdf:
        for p in pages:
            text = pdf.pages[p-1].extract_text() or ""
            print(f"  Page {p} first 600 chars:")
            print(text[:600])
            print()
