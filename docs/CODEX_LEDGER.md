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

\- \[x] Fix-05 Manual working-copy always-on + two-way sync to MIH/UAP + rent roll + year (requires event modules / sheet code)

\- \[x] Fix-06 Local rent calculation for Manual Working Copy (Z:\ shared + AppData fallback)

\- \[x] Fix-06b Harden local manual rent calc (fingerprint + fail-fast errors)

\- \[x] Fix-06c Table-driven local rent calc + per-user rent tables cache (CSV)

\- \[x] Fix-06d Verify Manual Rents (API) via stateless /api/evaluate



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


\### 2026-02-18 Fix-05 Manual working-copy always-on + 2-way sync (MIH/UAP + rent roll)

\- Repo: martin10101/nyc-ami-calculator

\- Base branch: perf/api-optimize-speed-2026-02-05

\- Work branch: fix/05-manual-working-copy-sync

\- Commit: 77f9b75e7d2aa6dcf7b0cb2b281373094101aa26

\- PR: \#13 https://github.com/martin10101/nyc-ami-calculator/pull/13 (draft)

\- Files changed:

&nbsp; - CODEX.md

&nbsp; - docs/CODEX\_LEDGER.md

&nbsp; - docs/FIX\_REQUIREMENTS.md

&nbsp; - docs/TASKCARD\_Fix-05\_Manual\_working\_copy\_sync.md

&nbsp; - excel-addin/customUI/customUI14.xml

&nbsp; - excel-addin/src/AMI\_Optix\_AppEvents.cls

&nbsp; - excel-addin/src/AMI\_Optix\_EventHooks.bas

\- Summary:

&nbsp; - Ribbon: removed the Manual “Live Sync” toggle; Manual Working Copy is always-on.

&nbsp; - Live sync: on any AMI edit (Manual Working Copy or MIH/UAP AMI column), refresh the manual block via `/api/manual_calculate` so invalid mixes are allowed (tradeoffs shown) and rent roll calcs update.

&nbsp; - Eventing: added stronger suppression to prevent re-entrant SheetChange loops; no auto-revert of edits.

\- Tests run + results: python -m pytest -q (51 passed, 13 warnings); edge optimize sanity: status 200, scenarios=4

\- Render deploy: manual; auto-deploy OFF; ready-to-deploy commit SHA = 77f9b75e7d2aa6dcf7b0cb2b281373094101aa26

\- Notes / risks:

&nbsp; - Manual edits now persist even when invalid by program rules (intentional); tradeoffs are shown in the manual block.

&nbsp; - Live refresh uses the Manual Calculate path; it temporarily activates AMI Scenarios internally but restores the user’s active sheet/selection with events disabled to avoid macro side effects.

\- Next: (stop; queue complete)


\### 2026-02-18 Fix-06 Local rent calculation for Manual Working Copy (Z:\ + AppData fallback)

\- Repo: martin10101/nyc-ami-calculator-all-fixes-test

\- Base branch: main

\- Work branch: fix/06-local-manual-rent-calc

\- Commit: bc75a5307007ddeb038efb9b21b138231bc957d5

\- PR: \#1 https://github.com/martin10101/nyc-ami-calculator-all-fixes-test/pull/1 (draft)

\- Files changed:

&nbsp; - docs/TASKCARD\_Fix-06\_Local\_manual\_rent\_calc.md

&nbsp; - excel-addin/src/AMI\_Optix\_ResultsWriter.bas

&nbsp; - excel-addin/src/AMI\_Optix\_AppEvents.cls

\- Summary:

&nbsp; - Manual Working Copy refresh (on AMI edits) now computes rents/totals locally from the selected year's rent workbook (Z:\ first; %APPDATA% fallback) and caches the opened workbook.

&nbsp; - Manual edits no longer call `/api/manual_calculate` (no per-edit API calls); solver scenario grid/output remains API-provided and unchanged.

\- Tests run + results: python -m pytest -q (51 passed, 13 warnings); edge optimize sanity: status 200, scenarios=4

\- Render deploy: manual; auto-deploy OFF; ready-to-deploy commit SHA = bc75a5307007ddeb038efb9b21b138231bc957d5

\- Notes / risks:

&nbsp; - If the year workbook is missing/unreadable, the manual block will still refresh but rents may be blank/0 and a warning is shown in Tradeoffs.

&nbsp; - Constraint-validity tradeoffs from `/api/manual_calculate` are no longer computed per edit; this fix focuses on local rent/totals refresh.

\- Next: (stop; Fix-06 complete)


\### 2026-02-18 Fix-06b Harden local manual rent calc (fingerprint + fail-fast errors)

\- Repo: martin10101/nyc-ami-calculator-all-fixes-test

\- Base branch: main

\- Work branch: fix/06-local-manual-rent-calc

\- Commit: fdd2cbcb6edb75524dffe30b3ae0f478026ad393

\- PR: \#1 https://github.com/martin10101/nyc-ami-calculator-all-fixes-test/pull/1 (draft)

\- Files changed:

&nbsp; - docs/TASKCARD\_Fix-06b\_Local\_rent\_calc\_hardening.md

&nbsp; - excel-addin/src/AMI\_Optix\_ResultsWriter.bas

&nbsp; - docs/CODEX\_LEDGER.md

\- Summary:

&nbsp; - Added a rent workbook layout fingerprint check (anchors on "AMI & Rent" headers + "of AMI" markers) and fail-fast behavior when the workbook layout is unexpected.

&nbsp; - Missing gross rent / utility allowance lookups now raise a hard error with unit_id + key + workbook path; manual rents/totals are not written on failure (prevents silent 0 / partial results).

&nbsp; - Solver scenario rents remain API-provided and unchanged (this only affects manual local refresh).

\- Tests run + results: python -m pytest -q (51 passed, 13 warnings); edge optimize sanity: status 200, scenarios=4

\- Render deploy: manual; auto-deploy OFF; ready-to-deploy commit SHA = fdd2cbcb6edb75524dffe30b3ae0f478026ad393

\- Notes / risks:

&nbsp; - Local manual refresh will now show a blocking error if the year workbook layout/labels drift; update the year workbook to match the expected "AMI & Rent" structure and option labels.

\- Next: (stop; Fix-06b complete)


\### 2026-02-18 Fix-06c Table-driven local rent calc + per-user rent tables cache (CSV)

\- Repo: martin10101/nyc-ami-calculator-all-fixes-test

\- Base branch: main

\- Work branch: fix/06c-table-driven-rent-tables

\- Commit: 7b264fd205933d840afd18a11166fd1f630354a4

\- PR: \#2 https://github.com/martin10101/nyc-ami-calculator-all-fixes-test/pull/2 (draft)

\- Files changed:

&nbsp; - docs/RENT\_TABLES\_PLAN.md

&nbsp; - excel-addin/customUI/customUI14.xml

&nbsp; - excel-addin/src/AMI\_Optix\_ResultsWriter.bas

&nbsp; - excel-addin/src/AMI\_Optix\_Diagnostics.bas

&nbsp; - excel-addin/src/AMI\_Optix\_Ribbon.bas

&nbsp; - excel-addin/src/AMI\_Optix\_RentCalcTables.bas

&nbsp; - excel-addin/src/AMI\_Optix\_RentTables.bas

\- Summary:

&nbsp; - Manual Working Copy local rent refresh now uses normalized, table-driven lookups from per-user CSV cache (not runtime workbook scraping).

&nbsp; - Cache source resolution: prefer Z:\ year workbook; fallback to %APPDATA% year workbook; per-user cache stored in %APPDATA%\\AMI\_Optix\\RentTablesCache\\<YEAR>\\.

&nbsp; - Added Ribbon action to force-refresh cache for selected year (and log diagnostics); missing data still hard-fails and blocks rent writes (no silent 0).

&nbsp; - Added a visible "Rent Tables Status" panel on the Diagnostics sheet (selected year, cache source path/label, source last-modified, cache built timestamp) plus a Ribbon status label/button for at-a-glance visibility.

&nbsp; - Hardened cache invalidation + concurrency safety: explicit rebuild reasons in diagnostics/logs and more robust read-only year workbook open/close behavior.

&nbsp; - Solver scenario rents remain API-provided and unchanged (this only affects manual local refresh).

\- Tests run + results: python -m pytest -q (51 passed, 13 warnings)

\- Manual test checklist (Excel):

&nbsp; - PENDING: Manual edit changes rents immediately with NO /api/manual\_calculate calls.

&nbsp; - PENDING: Switch year 2024 <-> 2025 and confirm rent changes where expected.

&nbsp; - PENDING: Flip a utility variant and confirm allowance + net rent changes.

&nbsp; - PENDING: Force Z: unavailable -> AppData year workbook used -> cache still builds/loads.

&nbsp; - PENDING: Deliberately break workbook layout -> cache build fails fast; rents not written.

\- Render deploy: manual; auto-deploy OFF; ready-to-deploy commit SHA = 7b264fd205933d840afd18a11166fd1f630354a4

\- Notes / risks:

&nbsp; - PR \#1 was auto-closed when the branch was renamed; PR \#2 supersedes it.

&nbsp; - Cache import falls back to the existing "AMI & Rent" scraper only when named tables are absent; runtime calculation never falls back to scraping.

\- Next: (stop; Fix-06c pending manual validation)

\### 2026-02-18 Fix-06d Verify Manual Rents (API) via stateless /api/evaluate

\- Repo: martin10101/nyc-ami-calculator-all-fixes-test

\- Base branch: main

\- Work branch: fix/06c-table-driven-rent-tables

\- Commit: 8a5d49096bf3b1c40e0b849e8e0073226cda683a

\- PR: \#2 https://github.com/martin10101/nyc-ami-calculator-all-fixes-test/pull/2 (draft)

\- Files changed:

&nbsp; - app.py

&nbsp; - tests/test_api_evaluate.py

&nbsp; - excel-addin/customUI/customUI14.xml

&nbsp; - excel-addin/src/AMI\_Optix\_API.bas

&nbsp; - excel-addin/src/AMI\_Optix\_Diagnostics.bas

&nbsp; - excel-addin/src/AMI\_Optix\_Ribbon.bas

&nbsp; - excel-addin/src/AMI\_Optix\_VerifyManualRents.bas

\- Summary:

&nbsp; - /api/evaluate now accepts rent_roll_year and/or calculator_id to select the rent schedule per request (stateless) and echoes back year_used/calculator_filename metadata for reproducible verification.

&nbsp; - Added a Ribbon button “Verify Manual Rents (API)” that calls /api/evaluate exactly once on demand, compares per-unit net rents + totals to the local Fix-06c cache computation with a $1 tolerance, shows a single MATCH/MISMATCH summary, and records full details in Diagnostics.

&nbsp; - Diagnostics now includes a “Verify Manual Rents (API)” panel with last verify timestamp + result and a full mismatch table (unit_id + local vs API + delta).

\- Tests run + results: python -m pytest -q (52 passed, 13 warnings)

\- Manual test checklist (Excel):

&nbsp; - PENDING: Click Verify with a normal manual state: expected MATCH; Diagnostics logs year, cache source path, and API year/calculator used.

&nbsp; - PENDING: Deliberately create a mismatch (e.g., modify local cache or select a year missing on server): expected MISMATCH; summary shows top 5 mismatched unit_ids; full list in Diagnostics.

&nbsp; - PENDING: Confirm no /api/manual\_calculate calls occur during edits (only /api/evaluate when the Verify button is clicked).

\- Render deploy: manual; auto-deploy OFF; ready-to-deploy commit SHA = 8a5d49096bf3b1c40e0b849e8e0073226cda683a


\### 2026-02-19 Release branch created: all fixes combined (Fix-01..Fix-06d)

\- Repo: martin10101/nyc-ami-calculator-all-fixes-test

\- Base branch: main

\- Work branch: release/2026-ami-optix-all-fixes

\- Commit: e5bda4d3c6dca92f5bcc4292a7214d6019eee7ef (release branch base commit containing Fix-06c + Fix-06d + ledger updates)

\- PR: \#3 https://github.com/martin10101/nyc-ami-calculator-all-fixes-test/pull/3

\- Files changed:

&nbsp; - docs/CODEX\_LEDGER.md

&nbsp; - docs/RELEASE\_NOTES\_2026.md

&nbsp; - docs/EXCEL\_INSTALL\_GUIDE.md

\- Summary:

&nbsp; - Created final combined release branch without squashing history; verified Fix-01, Fix-04, Fix-03, Fix-01b, Fix-02, Fix-05, Fix-06, Fix-06b, Fix-06c, and Fix-06d commits are all included.

&nbsp; - Added release notes and a full Excel install/update guide with explicit module/class/form import paths and troubleshooting.

&nbsp; - No solver scenario API rent behavior changes; no reintroduction of per-edit `/api/manual_calculate`.

\- Tests run + results: python -m pytest -q (52 passed, 13 warnings)

\- Render deploy: manual; auto-deploy OFF; ready-to-deploy commit SHA = e5bda4d3c6dca92f5bcc4292a7214d6019eee7ef

\- Notes / risks:

&nbsp; - Render auto-deploy remains OFF; manual deploy is required in Render UI.

&nbsp; - Release PR is open and tracks this release branch to `main`.
