# ACFR Digest — Project Instructions

## Current state (2026-07-07, post-independent-audit)
The Set3 batch (100 state/county/city ACFRs) is COMPLETE and INDEPENDENTLY AUDITED.
Final deliverable: `FY25_Set3_results.xlsx` (audited run), log `set3_run_final.log`,
pre-fix baseline at `set3_run_pass1.log` / `FY25_Set3_results_pass1.xlsx`.

Coverage: **983/1100 fields (89.4%)**. 56 files complete (11/11), 44 partial, **0 zero**.
PER-PAGE OCR FALLBACKS 2026-07-07 (commit c8ca62f): San Antonio BS 0→8 fields (cross-
verified: OCR BS total FB = native RevEx ending FB exactly), Allegheny +2 (RevEx totals),
Massachusetts +2 (bonus rescue). Adoption requires strictly-more-fields AND rev/exp
plausibility (0.25–4.0 ratio) — the ratio gate exists because Honolulu's OCR attempt
adopted a wrong revenues grab (caught by identity audit, reverted).
OCR LAYER WIRED 2026-07-07: mojibake PDFs (Arkansas 0→11/11, Atlanta 0→10/11) now
extract via a pdfplumber-compatible tesseract adapter (OCRPage/OCRPdf) — strip-OCR for
page identification, full-page OCR on statement pages only; words carry an 'ocr' marker
that widens the value-band tolerance to +12 (native text stays +5 — a global +12
regressed six dense files, caught by the identity audit, reverted). Expenditures can be
derived exactly from a pure excess line when OCR drops an underlined total (noted).
Requires: tesseract (brew, 5.5.2), pdftoppm, pytesseract; degrades to honest fast-bail
notes when absent.
Orange County RESOLVED 2026-07-07 (0/11 → 10/11): MD&A typo trap ("Disccusion"),
two-page-spread layout (title on right page, GF column on left), underscore-interleaved
unit text — all three fixed generically (commit b89e715).

INDEPENDENT AUDIT (2026-07-07): (1) determinism — three fresh runs byte-identical to
references; (2) accounting-identity audit across 113 files (FB category sums, BS residual,
magnitude/unit sanity) — found 4 real defects, ALL FIXED + verified: DC & Broward units
(`($000s)` pattern gap; `(InThousands)` kerning merge → compact unit matching added),
Kentucky nonspendable (footnote-ref '1)' token poisoning the value join), Detroit committed
(dash-row single-line capture overwrote accumulated 20,000,000 — category captures now sum,
never overwrite); (3) independent PDF re-derivation via pdftotext — 24/24 sampled figures
match, 8/8 units correct, GF column correct in all 8. Post-fix re-audit: **0 findings,
FB identity 78/78 across all three sets**. Audit tooling: scratchpad audit/identity_audit.py
(note: allows small negative BS residual — NC reports fund-level deferred outflows).
Regression gates 1-11 all clean.

BACKUP: full snapshot (script + all three result workbooks + logs + baseline + audit
script + README) at `Backups/2026-07-07_audited/`. Script copy sha256-verified.

## Project facts
- `parse_cafr.py` is the production parser (CLI: `--dir <folder> --output <xlsx>`).
- Regression ground truth: Set1 (8 cities, `FY25_City_ACFRs`) + Set2 (5 files,
  `FY25_City_ACFRs_Set2`). ANY parser edit requires a clean diff of both before batch runs
  (~6 min total). Ten consecutive clean gates through the 2026-06-23/24 campaign.
- Safe-failure principle: NOT FOUND beats a plausible wrong number. Prefer guards that
  convert silent-wrong into visible-missing.
- Run batches detached: `nohup caffeinate -i python3 parse_cafr.py ... > log 2>&1 &` —
  survives session close; arm a harness watcher (`until`-loop in background Bash) for
  completion notification.
- Fix-campaign history and per-fix rationale: see NEXT_SESSION_PLAN.md and the
  2026-06-23/24 session transcripts. Key architecture: page-ID (strict pass then relaxed
  'general fund' second pass), per-page derived label gutter (lower-only), continuation-page
  column re-verification (skip other-fund pages, re-adopt shifted GF bands), adaptive word
  tolerance for tight-kerning PDFs, mojibake fast-bail, positional statement rescue,
  OFS overwrite guard.

## Remaining backlog (post-deliverable, in rough priority order)
1. Re-triage the partial files below 10/11 (NY State rotated text; Nebraska 3/11,
   Allegheny FB section resisted OCR too, Miami-Dade/Honolulu 5/11 are the big ones).
2. Atlanta OFS (10/11) and Arkansas expenditures-derivation confirmation against a
   published Arkansas summary (value is exact-identity-derived, flagged in notes).
3. Incremental CSV checkpointing during batches; per-statement unit multipliers.
