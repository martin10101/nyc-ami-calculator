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

\- \[x] Fix-04 Ribbon stays active (no collapsing; no modal MsgBox in dropdown callbacks)

\- \[x] Fix-03 Rent roll YEAR selector + Manage uploads + mismatch warning (MIH + UAP)

\- \[x] Fix-01b Utilities variant breakdown (top-of-sheet; do not collapse variant labels)

\- \[x] Fix-02 Solver dedupe identical outcomes + placement tie-break (Python solver)

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



\### 2026-02-18 Fix-04 Ribbon stays active (no collapsing)

\- Repo: martin10101/nyc-ami-calculator

\- Base branch: perf/api-optimize-speed-2026-02-05

\- Work branch: fix/04-ribbon-stays-active

\- Commit: f6792738125956226d4088d90239eba42cbfc8ff

\- PR: \#9 https://github.com/martin10101/nyc-ami-calculator/pull/9

\- Files changed:

&nbsp; - docs/TASKCARD\_Fix-04\_Ribbon\_stays\_active.md

&nbsp; - excel-addin/customUI/customUI14.xml

&nbsp; - excel-addin/src/AMI\_Optix\_Ribbon.bas

&nbsp; - docs/CODEX\_LEDGER.md

\- Summary:

&nbsp; - Ribbon XML: add onLoad hook so VBA can capture `IRibbonUI`.

&nbsp; - Ribbon callbacks: best-effort re-activate `tabAMIOptix` after `onAction` handlers so the AMI Optix tab remains selected after actions.

\- Tests run + results: python -m pytest -q (51 passed, 13 warnings)

\- Render deploy: manual; auto-deploy OFF; ready-to-deploy commit SHA = f6792738125956226d4088d90239eba42cbfc8ff

\- Notes / risks:

&nbsp; - PR \#9 is stacked on PR \#8; merge PR \#8 first to reduce diff.

\- Next: Fix-03



\### 2026-02-18 Fix-03 Rent roll YEAR selector + Manage uploads + mismatch warning (MIH + UAP)

\- Repo: martin10101/nyc-ami-calculator

\- Base branch: perf/api-optimize-speed-2026-02-05

\- Work branch: fix/03-rent-roll-year-selector

\- Commit: e07ecb0ecb0623622d7311996e56911c8baba379

\- PR: \#10 https://github.com/martin10101/nyc-ami-calculator/pull/10

\- Files changed:

&nbsp; - CODEX.md

&nbsp; - docs/CODEX\_LEDGER.md

&nbsp; - docs/FIX\_REQUIREMENTS.md

&nbsp; - docs/TASKCARD\_Fix-03\_RentRoll\_YEAR\_selector.md

&nbsp; - excel-addin/customUI/customUI14.xml

&nbsp; - excel-addin/src/AMI\_Optix\_Ribbon.bas

&nbsp; - excel-addin/src/AMI\_Optix\_API.bas

\- Summary:

&nbsp; - Ribbon: add **Rent Roll Year** dropdown (2022â€“2026, default 2025) + **Manage Rent Roll Yearsâ€¦** upload/replace action.

&nbsp; - Local persistence: store selected year calculators under `%APPDATA%\\AMI_Optix\\RentRollYears\\<year>\\`.

&nbsp; - API storage: upload selected year calculator to `/api/rent-calculators/upload` and activate via `/api/rent-calculators/activate`.

&nbsp; - Enforcement: API calls (optimize/evaluate/manual\_calculate) ensure the selected year is active before running.

&nbsp; - Warning: best-effort mismatch warning if workbook appears to declare a different year than the selection (OK continues; not blocking).

\- Tests run + results: python -m pytest -q (51 passed, 13 warnings)

\- Render deploy: manual; auto-deploy OFF; ready-to-deploy commit SHA = e07ecb0ecb0623622d7311996e56911c8baba379

\- Notes / risks:

&nbsp; - Rent calculator activation is server-global; switching years affects other clients using the same API service.

&nbsp; - Declared-year detection is best-effort; some templates may not yield a detectable year (no warning shown).

\- Next: Fix-01b


\### 2026-02-18 Fix-01b Utilities variant breakdown (top-of-sheet; do not collapse variant labels)

\- Repo: martin10101/nyc-ami-calculator

\- Base branch: perf/api-optimize-speed-2026-02-05

\- Work branch: fix/01b-utilities-variant-breakdown

\- Commit: 954abcfd0b957258089f54b4d5e1f7bbfc2fd73a

\- PR: \#11 https://github.com/martin10101/nyc-ami-calculator/pull/11

\- Files changed:

&nbsp; - CODEX.md

&nbsp; - docs/CODEX\_LEDGER.md

&nbsp; - docs/FIX\_REQUIREMENTS.md

&nbsp; - docs/TASKCARD\_Fix-01b\_Utilities\_variant\_breakdown.md

&nbsp; - excel-addin/src/AMI\_Optix\_ResultsWriter.bas

\- Summary:

&nbsp; - AMI Scenarios (top-of-sheet): utilities block now shows the exact rent roll guideline variant labels per category (no collapsed 'Gas/Oil/Electric' labels).

&nbsp; - No scenario grid/per-scenario column changes.

\- Tests run + results: python -m pytest -q (51 passed, 13 warnings)

\- Render deploy: manual; auto-deploy OFF; ready-to-deploy commit SHA = 954abcfd0b957258089f54b4d5e1f7bbfc2fd73a

\- Notes / risks:

&nbsp; - Heat labels include guideline footnote markers ('...ccASHP)1', 'Other2'), matching the rent calculator workbook.

\- Next: Fix-02


\### 2026-02-18 Fix-02 Solver dedupe identical outcomes + placement tie-break (Python)

\- Repo: martin10101/nyc-ami-calculator

\- Base branch: perf/api-optimize-speed-2026-02-05

\- Work branch: fix/02-solver-outcome-dedupe

\- Commit: acac3a76ca5760f5f3ca8fdf09cb3fd54bbfcbb8

\- PR: \#12 https://github.com/martin10101/nyc-ami-calculator/pull/12

\- Files changed:

&nbsp; - CODEX.md

&nbsp; - docs/CODEX\_LEDGER.md

&nbsp; - docs/FIX\_REQUIREMENTS.md

&nbsp; - docs/TASKCARD\_Fix-02\_Solver\_outcome\_dedupe.md

&nbsp; - app.py

\- Summary:

&nbsp; - API (/api/optimize): post-process returned scenarios to remove outcome-identical duplicates (same band mix + same net rent totals), keeping the best floor placement (40% lower floors; higher AMI/rent higher floors).

\- Tests run + results: python -m pytest -q (51 passed, 13 warnings)

\- Render deploy: manual; auto-deploy OFF; ready-to-deploy commit SHA = acac3a76ca5760f5f3ca8fdf09cb3fd54bbfcbb8

\- Notes / risks:

&nbsp; - Dedupe is conservative and only runs when rent totals are available (rent calculator loaded).

&nbsp; - Edge check: Test.xlsx optimize returned 5 scenarios and reported 1 duplicate removed (Fix-02 note).

\- Next: Fix-05



