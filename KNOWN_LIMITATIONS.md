# Known Limitations — ACFR Digest (alpha)

Read this before pasting extracted figures into anything that matters.

## What it extracts, and from where

Eleven figures per report, all from the **General Fund column** of two
governmental-fund statements:

- **Balance Sheet**: Total Assets, Total Liabilities, the five GASB 54 fund
  balance categories (Nonspendable, Restricted, Committed, Assigned,
  Unassigned), and Total Fund Balances.
- **Statement of Revenues, Expenditures, and Changes in Fund Balances**:
  Total Revenues, Total Expenditures, and Total Other Financing Sources (Uses).

Plus three metadata fields: entity name, fiscal year end, and reporting units.
The workbook also appends USD-normalized columns (raw figure × detected unit).

## Coverage is ~95%, not 100%

On the most recent audited corpus — 100 state, county, and city FY2025 ACFRs —
the parser extracted **1,049 of 1,100 possible figures (95.4%)**. Every report
yielded at least some fields, but roughly 1 in 20 figures comes back NOT FOUND.
Expect gaps.

## The parser refuses to guess

The core design principle: **a visible blank beats a plausible wrong number.**
When the parser can't establish a figure with confidence — ambiguous columns,
fused characters, values printed outside their column band — it returns
NOT FOUND rather than its best guess. A blank cell means "look this one up
yourself," never "zero."

## Rows flagged ⚠️ need manual verification

Any row whose extraction notes mention OCR, derived values, positional
inference, or "verify manually" is flagged in the UI. These figures were
recovered through fallback paths that are right most of the time but are
exactly where errors concentrate. Check them against the source PDF before use.

## PDF classes that still defeat it

- **Fused/glued text**: PDFs whose text layer merges labels and numbers into
  single tokens (e.g., Allegheny County's fund balance section). Token-repair
  handles common cases; the worst survive.
- **Out-of-band figures**: numbers printed shifted out of their column's
  alignment (e.g., Honolulu's liabilities). Widening the search band enough to
  catch them would risk grabbing neighboring-column numbers — a silent wrong
  answer — so the parser leaves them NOT FOUND.
- **Unusual Other Financing Sources layouts** (e.g., Atlanta) can miss the
  OFS total.
- **Scanned or broken-encoding PDFs** require the optional OCR toolchain
  (poppler + tesseract). Without it, these files return honest NOT FOUNDs.
  With it, OCR-sourced figures are extracted and always flagged for manual
  verification.

## Units are detected per report, not per statement

Reporting units (full dollars / thousands / millions) are detected once per
report. If the two statements print in different units, the parser picks one
and records a note — the USD-normalized columns for one statement can then be
off by a factor of 1,000. OCR-sourced figures never re-detect units. **Raw
figures are unaffected**; treat the USD columns as a convenience, not a source
of record.

## What the audit does — and doesn't — guarantee

The audited release passed: (1) a determinism check (same PDF + same toolchain
always yields the same output); (2) an accounting-identity audit across all
113 processed reports (fund balance categories sum to the reported total,
balance sheet residuals, magnitude sanity) with zero findings; (3) independent
re-derivation of a 24-figure sample directly from the PDFs, all matching.

This does **not** guarantee every figure on a PDF the parser hasn't seen.
Identity checks can't catch an error that is internally self-consistent, and
new PDF layouts can fail in new ways. For publication-grade use, verify
flagged rows and spot-check material figures against the source document.

## Scope

U.S. state and local government ACFRs, in English, with GASB-style
governmental fund statements. Out of scope: proprietary/enterprise fund
statements, government-wide statements, budgetary comparison schedules, and
non-U.S. reports.
