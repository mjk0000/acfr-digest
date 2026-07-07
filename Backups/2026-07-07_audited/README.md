# ACFR Parser Backup — 2026-07-07 (post-independent-audit)

Snapshot taken after the Set3 campaign completed and passed a three-layer independent
audit (determinism, accounting identities, PDF re-derivation) with 0 findings.

| File | What |
|---|---|
| `parse_cafr.py` | Production parser, all fixes through audit wave (sha256 verified against working copy at backup time) |
| `FY25_Set1_results.xlsx` | Fresh audited run, 8 large-city ACFRs |
| `FY25_Set2_results.xlsx` | Fresh audited run, 5 Set2 files |
| `FY25_Set3_results.xlsx` | FINAL audited deliverable — 100 state/county/city ACFRs, 948/1100 fields |
| `FY25_Set3_results_pass1_baseline.xlsx` | Pre-fix-campaign baseline (865/1100) for comparison |
| `set3_run_final.log` / `set3_run_pass1_baseline.log` | Per-file extraction logs for both |
| `identity_audit.py` | Independent audit script (GAAP identity + magnitude checks); run against any output workbook |

Known limitations at snapshot time: Arkansas/Atlanta (unusable text layers, need OCR),
Orange County (0/11, undiagnosed), Allegheny partial (glued tokens), San Antonio BS
(broken font cmap). Full state and backlog: project CLAUDE.md.
