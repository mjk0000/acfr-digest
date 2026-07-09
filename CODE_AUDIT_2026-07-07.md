# Code Audit — parse_cafr.py (2026-07-07, post-campaign)

Three parallel review agents (state-machine correctness, dead-code/redundancy,
robustness/performance) over the ~2,250-line parser after the fix campaign.
Every load-bearing claim verified against the code before action. Disposition
of every finding below. Post-fix verification: gates 22 clean (Set1 8/8 +
Set2 5/5 identical), full Set3 batch **zero changed cells** vs the 1049
reference, identity audit **0 findings** — the wave is provably
behavior-neutral on the audited corpus while closing future-input hazards.

## FIXED in this wave (commit refs in git log)

Fabrication-path closures (trace-confirmed, unit-tested):
- **parse_number comma-grouping validation** — malformed fusions that escape
  token repairs ('5,3781') fabricated 10x values invisible to the band guard.
  Closes three paths at once (two-way-ambiguous fusions, fused dot-leader
  tails, three-column fusions).
- **Split-number continuation requires an INCOMPLETE rightmost token** — a
  neighbor's kerning-split fragment could graft onto a complete in-band value
  ('338,871' + ',234').
- **Two-completes guard reordered BEFORE the lone-dash rule** — the dash rule
  could return a confident 0.0 on a band the guard would have refused.
- **OFS pending consumption exclusions** — fused/wrapped net-change rows and
  beginning-of-year balances could be captured as the OFS total, and the
  no-reopen guard then locked the wrong value in.
- **Bare 'Total' with an open FB category never promotes to grand total** —
  previously did so when the category's sub-items were unreadable, fabricating
  the total AND truncating remaining categories.
- **'deferred' exclusion on both liabilities paths**; **'financing' exclusion
  on revenue/expenditure totals** (combined-line captures).
- **bs_total_pending registers in fallback_captured** (overwritable);
  **prev_lbl/prev_val reset per page**; 'organic' added to compact DQ list.

Batch survivability / robustness:
- **Incremental CSV checkpoint** (`<output>.checkpoint.csv`, flushed per file)
  — no failure mode after file 1 can forfeit a batch again. Verified live.
- **write_excel control-char sanitization + try with checkpoint fallback** —
  one mojibake character in a note could previously vaporize a finished batch.
- **tesseract timeout=120** in image_to_data — an unbounded hang would
  silently freeze a detached batch. **Image.open moved inside the guard.**
- **Per-file insurance try in the main loop**; is_scanned_pdf except widened
  (OSError under memory pressure was a whole-batch abort).
- **Toolchain version logging** at batch start (pdfplumber/poppler/tesseract)
  — determinism breaks are now attributable to upgrades.
- **Degenerate-box guards** (x1<=x0) in both token splitters.
- Error notes truncated to 300 chars.

Cleanups:
- Dead "Layer 2" pdftotext fallback deleted (never wired; superseded by the
  OCR adapter). Five identical comma-number regexes unified into
  `_COMPLETE_NUM` (the divergence trap). Redundant unit-regex alternatives
  and dead tuple members removed. Fallbacks 2/3 now log when they fire
  (reachability evidence for future cleanup decisions).

## DEFERRED (documented backlog, in priority order)

1. **Per-statement unit multipliers** — unit disagreement currently resolves
   via sorted()[0] (alphabetical: Millions beats Thousands) with only a note;
   USD columns for one statement can be silently 1000x off. Also: units are
   not re-detected when figures come from per-page OCR (San Antonio class).
   The strongest remaining safe-failure gap; raw figures unaffected.
2. **Partial-accumulator refusal at page-window exhaustion** — an FB category
   still accumulating when the 3-page window ends is finalized as a partial
   sum (needs a 4+-page statement to trigger; none in current corpus).
3. **process_pdf BS/RevEx consolidation** (~95 duplicated lines; ends the
   fix-lands-in-one-block failure mode) + shared continuation preamble.
4. **OCR result cache** across the up-to-five OCRPdf instantiations per file
   (minutes/file on hostile PDFs); OCR-mode entity extraction does 15
   full-page OCRs where strips would do; a per-file OCR time budget.
5. **Named constants for the magic-number inventory** (title zones 300/250,
   3-page window, 180 gutter — continuation check hardcodes it while page 1
   may derive lower —, 35-char name cap, 45pt refine proximity, 0.25-4.0
   ratio, conf>=30, 0.2 mojibake threshold — noting non-English ACFRs will
   false-trigger it).
6. **BS-side plausibility guard on OCR re-adoption** (RevEx has one; BS
   adoption is count-only). **page.flush_cache()** for 500+-page documents.
   Fused-subtotal item-tracker gaps (P6). DQ-list derivation from one source
   (latent 'cip'-substring vs 'principal' hazard). Requirements pinning.
7. Repo tidiness: examine2.py / examine_pdfs.py are June-22 one-offs with no
   imports from parse_cafr (archive candidates); diagnose.py IS a live
   import dependency — renames must update it.

## Explicitly reviewed and KEPT (not redundant)

is_header_tail + 'and '-prefix (distinct coverage); the three-defense TOC
chain (stripper / bad-word / dot-leader — interlocking, Bexar proves it);
column-detection fallback 4 (fired on Mecklenburg); refine_gf_band (201
activations in Set3); the OFS positional fallback (needs Set1/Set2 log
evidence before touching). Determinism review: sorted glob, insertion-order
dicts, sorted() over sets — byte-identity is sound for a fixed toolchain.
