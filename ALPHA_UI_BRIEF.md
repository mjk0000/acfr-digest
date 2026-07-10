# Mission Brief: Alpha Package + Drag-and-Drop UI for the ACFR Parser

_Handoff doc written 2026-07-07 at the end of the audit session (context was at 77%).
Start a new session in this project directory — CLAUDE.md loads automatically and carries
the parser's full state. This file is the work order._

## Mission
Package the audited parser (`parse_cafr.py`, commit a45ad45, 1049/1100 fields, 0 audit
findings) as an alpha for external testers, fronted by a Streamlit UI where users drop
PDFs (single, multi-select, or a zipped folder) and get back an Excel workbook in the
EXISTING output format — all 11 fields per file, NOT FOUND markers included, extraction
notes surfaced. Mike has approved this plan explicitly.

## Non-negotiables
1. **`parse_cafr.py` is the audited asset — WRAP IT, do not modify it.** If any parser
   edit becomes unavoidable, the full discipline applies: Set1+Set2 regression gate AND
   the full-batch identity audit (see CLAUDE.md; audit script preserved at
   `Backups/2026-07-07_audited/identity_audit.py`). Twenty-two consecutive clean gates
   are the project's crown jewels.
2. **Output format is byte-compatible**: call the existing `write_excel(all_results,
   output_path, logger)` — do not re-implement the workbook. Same for `process_pdf(
   pdf_path: Path, logger) -> dict` (keys = `OUTPUT_COLUMNS`), which is the only
   entry point the UI needs per file.
3. **The trust surface moves INTO the UI.** External users will paste these figures into
   reports. Requirements:
   - NOT FOUND cells rendered as visible gaps with a one-line explainer ("this parser
     refuses to guess — a blank beats a plausible wrong number");
   - any row whose Extraction Notes contain "verify manually", "derived", "OCR", or
     "inferred" gets a warning badge, notes shown on hover/expand;
   - a Known Limitations page in the UI, adapted from CLAUDE.md's current-state section
     (95.4% coverage, the hostile-PDF classes: Allegheny glued tokens, Atlanta OFS,
     Honolulu right-shifted liabilities; what the identity audit does/doesn't guarantee).

## Deliverables checklist
- [ ] `pyproject.toml`: package `acfr-digest`, console entry point (`acfr-digest --dir …
      --output …` mapping to the existing `main()`), pinned deps (pdfplumber 0.11.8,
      openpyxl, streamlit; pytesseract/Pillow as an `[ocr]` extra). `--version` flag
      sourced from the package version; tag releases in git.
- [ ] `app.py` (Streamlit, keep it ~150-250 lines):
      * `st.file_uploader(accept_multiple_files=True)` for PDFs, PLUS accept a `.zip`
        (browsers cannot drag folders — unzip server-side, glob `*.pdf`). Note this
        limitation in the UI copy ("zip a folder to upload it whole").
      * Per-file progress (files take 1–3 min each; OCR-path files several more). Process
        sequentially with `st.status`/progress; never block silently.
      * Save uploads to a tempdir, run `process_pdf` per file, accumulate rows, call
        `write_excel`, offer `st.download_button` for the .xlsx (and the checkpoint CSV).
      * Results table on screen with the badge/gap styling from Non-negotiable 3.
      * OCR toolchain detection: call the existing `ocr_available()`; if absent, show a
        soft notice that scanned/mojibake PDFs will return honest NOT FOUNDs.
- [ ] `KNOWN_LIMITATIONS.md` (also rendered in-app) + `README.md` (install for CLI users:
      `pipx install .`, `brew install poppler tesseract` optional; UI users:
      `streamlit run app.py`).
- [ ] LICENSE: MIT (approved direction; confirm with Mike only if he's present).
- [ ] Smoke test: `tests/smoke.py` runs Set2 (5 local PDFs, 2 are known hard-fails —
      expected!) through `process_pdf` and diffs key fields against the audited baseline
      (`Backups/2026-07-07_audited/FY25_Set2_results.xlsx`). Fast (~2 min), proves the
      packaging didn't change behavior.
- [ ] Commit in the established style (see `git log` for voice; Co-Authored-By trailer).

## Verification before calling it done
1. Smoke test green.
2. Launch the UI, upload 2–3 Set1 PDFs (small ones: Philadelphia ~230p is quickest),
   confirm the downloaded workbook matches the corresponding rows of the audited
   baseline exactly, badges/gaps render, zip-upload path works.
3. CLI entry point produces byte-identical output to `python3 parse_cafr.py` on Set2.

## Explicitly out of scope (documented decisions)
- PyInstaller/native binaries, Electron/React — rejected for effort/audience fit.
- Hosted beta (Streamlit Community Cloud from a GitHub repo, apt packages
  `poppler-utils tesseract-ocr`) is the NEXT step after alpha feedback — design the app
  so deploy is config-only, but don't deploy yet. GitHub push awaits Mike's go-ahead.
- Parser backlog (per-statement units etc.) lives in CODE_AUDIT_2026-07-07.md — do not
  mix it into this packaging wave.

## Context pointers (read before writing code)
- `CLAUDE.md` — project state, safe-failure principle, batch/gate discipline.
- `CODE_AUDIT_2026-07-07.md` — what was just hardened and what's deferred.
- `parse_cafr.py`: `process_pdf`, `write_excel`, `OUTPUT_COLUMNS`, `ocr_available`,
  `main()` — the four seams the package/UI attach to. The main loop also writes
  `<output>.checkpoint.csv` per file — surface it in the UI as crash insurance.
- PDF corpora are gitignored and local-only; Set1/Set2 in `FY25_City_ACFRs*/`.
