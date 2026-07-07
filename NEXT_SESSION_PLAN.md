# Next Session Plan — Structural Fixes + Set3 Triage
_Written 2026-06-23. Context: three-agent audit complete, fix wave 1 (11 edits) applied and
regression-verified on Set1 (8/8 identical) + remaining Set2 (5/5 identical). Set3 batch of
100 state/county/city PDFs relaunched 16:47 DETACHED (nohup+caffeinate, survives session
close) on WAVE-1-FIXED code → `set3_run.log` / `FY25_Set3_results.xlsx`. The original
pre-fix control run was killed at ~4.5 files to gain detachment; its partial log is
`set3_run_attempt1_prefix.log` (Alaska/Jefferson County/Alabama figures + Arkansas entity
miss — the only true pre-fix baseline data). Consequence: Set3 first pass now measures
wave-1 code, so the Fix A/B rerun diffs against THAT, and audit-prediction grading uses
failure presence/shape rather than before/after counts for wave-1 items._

## Step 0 — Triage Set3 control-run results (do first)
- `grep -c "Processing:" set3_run.log` — confirm batch finished (100 expected).
- Bucket every failure/anomaly:
  - `No General Fund column identified` → candidate for Fix B (gutter)
  - Absurd/oversized figures (concatenation) → confirms audit F2-layout; wave-1 guard should catch on rerun
  - Wrong-but-plausible figures on multi-page statements → candidate for Fix A (continuation)
  - `Entity: NOT FOUND` → should be fixed by wave-1 entity patch (state_pat); count them
  - Unit anomalies (Millions detections, disagreement warnings won't appear in control run)
  - Hard failures (No /Root, CID) → list for pdftotext-wiring test
- Grade the 5 predictions from the pre-run audit + the 3 agents' findings against actual
  frequencies. Fix priority = observed hit count, not theoretical severity.

## Fix A — Horizontal continuation guard (deferred from audit; all 3 agents flagged)
**Problem:** `extract_bs_figures` (~line 600) and `extract_revex_figures` (~line 770) scan
start_page+3 applying PAGE 1's GF column x-band to continuation pages. States continue wide
statements horizontally: "(Continued)" page has DIFFERENT funds at the same x-positions →
silently wrong numbers. State flags (`in_fund_balance`, `fb_accumulator`, `ofs_total_pending`,
section trackers) also persist across the boundary.

**Design:** at each continuation page (pages_checked > 0), after the existing title check:
1. Run column-header detection on the continuation page (reuse `get_column_structure` or a
   lighter header-only pass).
2. Case 1 — a 'general' column found with x_center within ~25pt of page-1's: continue; prefer
   UPDATING the band to the re-detected one (handles 15–30pt re-typesetting shift).
3. Case 2 — column headers found but NO general column: horizontal continuation carrying
   other funds → `break` (NOT FOUND is the safe failure; wrong number is not). Append note.
4. Case 3 — no detectable column headers (pure vertical continuation, headers not repeated):
   keep page-1 band, current behavior.
5. Independently: reset `ofs_total_pending = False` at every page boundary in RevEx
   (state-machine audit F4 — pending must never satisfy across a page break).

**Tests:** Nashville (historic bleed onto p56 — guard should hit Case 2), any Set3 state
flagged in triage as multi-page-wrong, then full Set1+Set2 regression (now ~6 min total).

## Fix B — Dynamic label gutter (replaces hardcoded x ≥ 180)
**Problem:** three call sites assume the row-label zone ends at x=180: `find_data_start_row`
(~line 360), `cluster_header_columns` (~line 385), fallback-3 in `get_column_structure`
(~line 947). Dense portrait state statements (8–12 fund columns) compress the gutter below
180 → "General Fund" header words filtered out → GF column NOT FOUND. Landscape pages have
the opposite problem.

**Design:** derive the gutter per page:
1. Collect tokens matching a full comma-formatted number (`^\$?\(?\d{1,3}(?:,\d{3})+\)?$`,
   underscores stripped) — but ONLY from rows containing ≥2 such tokens (real data rows have
   multiple columns; this excludes centered title dates like "June 30, 2025" that merge into
   "30,2025").
2. `gutter = min(x0 of those tokens) - 12` (margin). Clamp to [100, 300]. Fallback 180 when
   no qualifying tokens (then behavior is exactly today's).
3. Thread `gutter` through the three call sites (parameter with default 180 keeps the
   function signatures compatible).

**Tests:** every Set3 file bucketed as "No General Fund column identified" in triage, then
full Set1+Set2 regression.

## Step 3 — Rerun Set3 with all fixes; diff vs control
`python3 parse_cafr.py --dir "FY25_City_ACFRs_Set3/_All_PDFs_Compiled" --output FY25_Set3_results_v2.xlsx`
(~40–60 min with early-break fix vs ~4h control). Diff v2 against control per file per field:
fixed / regressed / unchanged-bad. Spot-check 5 states manually against their actual PDFs.

## Wave 3 backlog (execute as evidence warrants, roughly in this order)
1. **Wire pdftotext fallback (audit C1 — currently 100% dead code):** wrap `pdfplumber.open`
   so `PDFSyntaxError`/'No /Root object' retries via `extract_via_pdftotext` layout mode.
   Test on Charlotte_FY2025 / Fort_Worth_FY2025 (may rescue both).
2. **Per-statement unit multipliers (audit C3):** wave 1 only added a disagreement NOTE.
   Full fix: carry `unit_bs` and `unit_revex`, apply respective multipliers per figure group
   in `write_excel`; 'Reporting Units' column shows both when they differ.
3. **Incremental checkpointing (audit H2):** append each result row to a CSV inside the batch
   loop; wrap `wb.save` in try with a timestamped fallback filename (output file open in
   Excel is a likely 4-hour-run killer).
4. **MD&A condensed-table trap (layout F4):** prefer statement pages whose FULL text contains
   'accompanying notes are an integral part' (standard footer); fall back to first candidate
   without footer if none match. Regression-check that all Set1/Set2 pages carry the footer.
5. **`x_gap=30` scaling (layout F2):** scale header-cluster gap by column count / page width
   for dense statements. Only if Set3 shows merged-header cases the wave-1 concatenation
   guard converts to NOT FOUND (guard makes it visible-safe; this fix makes it correct).
6. **Paren sign flip (state-machine F3):** if leftmost numeric token starts with a digit and a
   lone `(` token sits within 6pt left of `col_left`, include it. Only with an observed case.
7. **Deferred-outflows substring trap (F11):** require `total assets` label NOT contain
   'deferred' unless no plain match exists on the page.

## Known state (do not re-derive)
- Wave-1 fixes all applied to `parse_cafr.py` and regression-clean; unit checks for the
  value-band guard pass (two complete numbers → None; SF split-number still works).
- Set2 baseline SHRANK: Charlotte FY2023, Columbus, Denver, Fort Worth FY2024, Indianapolis,
  Jacksonville, Seattle PDFs were removed from FY25_City_ACFRs_Set2 (Columbus/Seattle in Set3
  are NEWER fiscal years, different documents). Ask Mike whether the originals still exist —
  they'd rebuild the regression corpus for the Jacksonville-typo and Charlotte-split-OFS paths.
- Charlotte_FY2025 + Fort_Worth_FY2025: 'No /Root object' hard fails (candidates for wave-3 #1).
- San Antonio BS: GF column not identified — pre-existing, unchanged; likely a Fix B beneficiary
  (BS page 156 is an anomaly worth checking too — the real BS may be earlier and misidentified).
