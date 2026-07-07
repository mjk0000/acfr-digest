"""
diagnose.py — Run after parse_cafr.py to inspect any PDF with NOT FOUND results.
Usage:
    python3 diagnose.py "path/to/report.pdf"
Shows: entity/date detection, column names found, and statement page detection.
"""
import sys
import re
import pdfplumber
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from parse_cafr import (
    get_page_words, cluster_header_columns, find_statement_pages,
    extract_entity_info, identify_general_fund_col, _get_title_area,
    is_balance_sheet_page, is_revex_page,
)

if len(sys.argv) < 2:
    print("Usage: python3 diagnose.py path/to/report.pdf")
    sys.exit(1)

pdf_path = Path(sys.argv[1])
print(f"\n{'='*60}")
print(f"DIAGNOSING: {pdf_path.name}")
print(f"{'='*60}")

with pdfplumber.open(pdf_path) as pdf:
    print(f"Total pages: {len(pdf.pages)}")

    # --- Entity / FY date ---
    entity, fy = extract_entity_info(pdf)
    print(f"\nEntity: {entity}")
    print(f"FY End: {fy}")

    # Show first 10 pages' first lines to help find entity manually
    print("\n--- First line of pages 1-10 (for entity/date debugging) ---")
    for i in range(min(10, len(pdf.pages))):
        raw = pdf.pages[i].extract_text() or ""
        first_lines = [l.strip() for l in raw.splitlines()[:4] if l.strip()]
        print(f"  p{i+1}: {' | '.join(first_lines[:3])}")

    # --- Statement page detection ---
    bs_idx, revex_idx = find_statement_pages(pdf)
    print(f"\nBalance Sheet page: {bs_idx + 1 if bs_idx is not None else 'NOT FOUND'}")
    print(f"RevEx page:         {revex_idx + 1 if revex_idx is not None else 'NOT FOUND'}")

    if bs_idx is None or revex_idx is None:
        print("\n--- Scanning all pages for near-match titles ---")
        for i, page in enumerate(pdf.pages):
            raw = page.extract_text() or ""
            ta = _get_title_area(raw)
            if any(kw in ta for kw in ['balance sheet', 'revenue', 'expenditure',
                                        'governmental fund', 'changes in fund']):
                bs = is_balance_sheet_page(ta)
                rx = is_revex_page(ta)
                snippet = ta[:120].replace('\n', ' ')
                print(f"  p{i+1}: BS={bs} RevEx={rx} | {snippet!r}")

    # --- Column identification ---
    for label, idx in [('Balance Sheet', bs_idx), ('RevEx', revex_idx)]:
        if idx is None:
            continue
        print(f"\n--- {label} (p{idx+1}): column headers ---")
        page = pdf.pages[idx]
        words = get_page_words(page)
        columns = cluster_header_columns(words)
        print(f"  All columns: {[c['name'] for c in columns]}")
        gf = identify_general_fund_col(columns)
        if gf:
            print(f"  General Fund column: '{gf['name']}' x_center={gf['x_center']:.1f}")
        else:
            print("  General Fund column: NOT IDENTIFIED")
