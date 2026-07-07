"""
Quick examination script — runs pdfplumber against all 5 ACFRs
to map statement page locations, column structure, and unit disclosures.
Not the production script; used only for pre-design exploration.
"""

import pdfplumber
import re
import sys
from pathlib import Path

PDF_DIR = Path(__file__).parent

PDFS = [
    "Annapolis FY25 Financial Report.pdf",
    "Gaithersburg FY25 Financial Report.pdf",
    "Frederick FY25 Financial Report.pdf",
    "Rockville FY25 Financial Report.pdf",
    "Baltimore FY25 Financial Report.pdf",
]

def normalize(text):
    if not text:
        return ""
    text = text.lower()
    text = re.sub(r'[–—―]', '-', text)   # em/en dashes
    text = re.sub(r'[‘’“”]', "'", text)  # smart quotes
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def is_balance_sheet_page(norm_text):
    has_bs = 'balance sheet' in norm_text
    has_gf = 'governmental fund' in norm_text
    no_budget = 'budget' not in norm_text
    no_combining = 'combining' not in norm_text
    no_proprietary = 'proprietary' not in norm_text
    no_fiduciary = 'fiduciary' not in norm_text
    no_net_pos = 'net position' not in norm_text
    return has_bs and has_gf and no_budget and no_combining and no_proprietary and no_fiduciary and no_net_pos

def is_revex_page(norm_text):
    has_rev = 'revenues' in norm_text
    has_exp = 'expenditures' in norm_text
    has_changes = 'changes in fund' in norm_text
    no_budget = 'budget' not in norm_text
    no_combining = 'combining' not in norm_text
    return has_rev and has_exp and has_changes and no_budget and no_combining

def extract_words_clustered(page):
    """Return list of (y_center, x_center, text) for words on page."""
    words = page.extract_words(x_tolerance=3, y_tolerance=3)
    return [(w['top'] + (w['bottom'] - w['top']) / 2, w['x0'] + (w['x1'] - w['x0']) / 2, w['text']) for w in words]

def get_column_headers(page, header_y_range=None):
    """Try to find column header row by looking for fund-name clusters."""
    words = page.extract_words(x_tolerance=5, y_tolerance=3)
    if not words:
        return []
    # Group words by approximate y position (rows)
    rows = {}
    for w in words:
        y = round(w['top'] / 5) * 5
        rows.setdefault(y, []).append(w)

    # Find rows that have multiple numeric-ish or fund-name tokens spread across x
    result = []
    for y, row_words in sorted(rows.items()):
        if len(row_words) >= 2:
            x_spread = max(w['x1'] for w in row_words) - min(w['x0'] for w in row_words)
            row_text = ' '.join(w['text'] for w in row_words)
            result.append((y, x_spread, row_text, row_words))
    return result

def examine_pdf(filename):
    path = PDF_DIR / filename
    print(f"\n{'='*70}")
    print(f"  {filename}")
    print(f"{'='*70}")

    bs_pages = []
    revex_pages = []

    with pdfplumber.open(path) as pdf:
        total = len(pdf.pages)
        print(f"  Total pages: {total}")

        # Scan all pages for target statements
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            norm = normalize(text)

            if is_balance_sheet_page(norm):
                bs_pages.append(i + 1)
            if is_revex_page(norm):
                revex_pages.append(i + 1)

        print(f"\n  Balance Sheet candidate pages: {bs_pages}")
        print(f"  RevEx candidate pages:         {revex_pages}")

        # Examine the first balance sheet page in detail
        if bs_pages:
            target_page_num = bs_pages[0]
            page = pdf.pages[target_page_num - 1]
            text = page.extract_text() or ""
            print(f"\n  --- Balance Sheet (page {target_page_num}) raw text ---")
            print(text[:3000])

            # Check preceding page for unit notation
            if target_page_num > 1:
                prev_text = pdf.pages[target_page_num - 2].extract_text() or ""
                prev_norm = normalize(prev_text)
                if 'thousand' in prev_norm or 'million' in prev_norm or '000' in prev_norm:
                    print(f"\n  [UNIT NOTATION on preceding page {target_page_num-1}]")
                    for line in prev_text.split('\n'):
                        if any(k in normalize(line) for k in ['thousand', 'million', '000']):
                            print(f"    >> {line.strip()}")

            # Check current page for unit notation
            for line in text.split('\n'):
                ln = normalize(line)
                if any(k in ln for k in ['thousand', 'million', '000s', 'in dollars']):
                    print(f"\n  [UNIT NOTATION on BS page]: {line.strip()}")

            # Show word layout / x positions for column detection
            print(f"\n  --- Column structure (first 40 word-rows) ---")
            words = page.extract_words(x_tolerance=5, y_tolerance=3)
            rows = {}
            for w in words:
                y_key = round(w['top'] / 4) * 4
                rows.setdefault(y_key, []).append(w)
            count = 0
            for y, rw in sorted(rows.items()):
                if count > 40:
                    break
                row_str = '  '.join(f"{w['text']}@{w['x0']:.0f}" for w in sorted(rw, key=lambda x: x['x0']))
                print(f"    y={y:5.0f}  {row_str}")
                count += 1

        # Examine the first revex page in detail
        if revex_pages:
            target_page_num = revex_pages[0]
            page = pdf.pages[target_page_num - 1]
            text = page.extract_text() or ""
            print(f"\n  --- RevEx (page {target_page_num}) raw text ---")
            print(text[:3000])

if __name__ == '__main__':
    for pdf in PDFS:
        examine_pdf(pdf)
