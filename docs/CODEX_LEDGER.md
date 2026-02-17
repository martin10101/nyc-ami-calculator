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

\- \[ ] Fix-01 DataReader stability + remove ribbon sheet activation (prevents wrong-sheet/AMI-col bugs)

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



