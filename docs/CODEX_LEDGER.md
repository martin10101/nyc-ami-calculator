\# CODEX LEDGER



\## Repo + Base Branch

\- Repo: martin10101/nyc-ami-calculator

\- Base branch: perf/api-optimize-speed-2026-02-05



\## Render (fill once, keep updated)

\- Service name:

\- Environment:

\- Service ID: srv-d37n6her433s73et1bp0

\- Auto-deploy setting (On Commit / After CI / Off): OFF

\- Deploy hook URL: \[REDACTED]



\## Fix Queue

\- \[x] Fix-01 DataReader stability + remove ribbon sheet activation (prevents wrong-sheet/AMI-col bugs)

\- \[ ] Fix-04 Ribbon stays active (no collapsing; no modal MsgBox in dropdown callbacks)

\- \[ ] Fix-03 Rent roll YEAR selector + Manage uploads + mismatch warning (MIH + UAP)

\- \[ ] Fix-01b Utilities variant breakdown (top-of-sheet; do not collapse variant labels)

\- \[ ] Fix-02 Solver dedupe identical outcomes + placement tie-break (Python solver)

\- \[ ] Fix-05 Manual working-copy always-on + two-way sync to MIH/UAP + rent roll + year (requires event modules / sheet code)



\## Work Log

\### 2026-02-17 Bootstrap

\- Repo: martin10101/nyc-ami-calculator

\- Base branch: perf/api-optimize-speed-2026-02-05

\- Work branch: n/a

\- Commit: n/a

\- PR: n/a

\- Files changed: CODEX.md, docs/CODEX\_LEDGER.md

\- Summary: Added persistent operating protocol + ledger so Codex can resume work reliably across sessions and track repo/branch/PR and Render deploy readiness.

\- Tests run + results: n/a

\- Render deploy: n/a

\- Notes / risks:

&nbsp; - Auto-deploy is OFF; deployments are manual from Render dashboard. Record the “ready-to-deploy” commit SHA in the ledger.

\- Next: Fix-01



\### 2026-02-17 Fix-01 DataReader stability + remove ribbon sheet activation

\- Repo: martin10101/nyc-ami-calculator

\- Base branch: perf/api-optimize-speed-2026-02-05

\- Work branch: fix/01-datareader-stability

\- Commit: e370726c9d1b8c7b8627cbb2cb59cdf72783103e

\- PR: \#8 https://github.com/martin10101/nyc-ami-calculator/pull/8

\- Files changed:

&nbsp; - CODEX.md

&nbsp; - docs/CODEX\_LEDGER.md

&nbsp; - docs/FIX\_REQUIREMENTS.md

&nbsp; - docs/TASKCARD\_Fix-01\_DataReader\_stability.md

&nbsp; - excel-addin/src/AMI\_Optix\_DataReader.bas

&nbsp; - excel-addin/src/AMI\_Optix\_Ribbon.bas

\- Summary:

&nbsp; - DataReader: FindDataSheet no longer treats incidental ActiveSheet as authoritative unless it's a known template/program sheet; prefers stable template names first to avoid wrong-sheet/AMI-column caching.

&nbsp; - Ribbon: rent roll dropdown no longer activates worksheets or shows modal selection UI (stores selection only; DebugLog).

\- Tests run + results: python -m pytest -q (51 passed, 13 warnings)

\- Render deploy: manual; auto-deploy OFF; ready-to-deploy commit SHA = e370726c9d1b8c7b8627cbb2cb59cdf72783103e

\- Notes / risks:

&nbsp; - Workbooks with multiple unit-like tables may resolve to the first matching template-named sheet when the active sheet is not a known program sheet.

\- Next: Fix-04



