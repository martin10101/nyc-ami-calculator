\# FIX REQUIREMENTS — NYC AMI Calculator



Base branch: perf/api-optimize-speed-2026-02-05  

Repo: martin10101/nyc-ami-calculator



This document defines the expected behavior for each fix.  

Codex must not reinterpret these requirements.



---



\## Fix-01 — DataReader stability + remove ribbon sheet activation

\### Problem

Scenario application and other actions sometimes write to the wrong sheet/column after UI interactions because DataReader binds to ActiveSheet and ribbon callbacks activate sheets.



\### Requirements

\- DataReader must locate MIH/UAP target sheet and AMI column using stable anchors.

\- FindDataSheet must NOT prefer ActiveSheet unless it is explicitly validated to be the MIH/UAP data sheet.

\- Ribbon dropdown callbacks must NOT `.Activate` worksheets and must not use modal MsgBoxes for normal selection.

\- Scenario apply must work even after any manual clears / switching sheets / using dropdowns.



\### Must NOT change

\- Excel sheet structure or headers.

\- Solver logic.

\- Scenario grid layout.



\### Test (manual)

\- Apply scenario #2 repeatedly while switching sheets; AMI writes always land in correct column.

\- After using rent roll dropdown, apply scenario; still correct.



---



\## Fix-04 — Ribbon stays active (no collapsing)

\### Problem

AMI ribbon tab collapses/disappears after interacting with worksheet, forcing user to click AMI tab again.



\### Requirements

\- Selecting AMI ribbon tab should behave like normal Excel tabs: it stays active while user works.

\- Dropdown selection and button clicks must not cause the tab to collapse.

\- Must work consistently for all ribbon buttons (MIH + UAP).



\### Must NOT change

\- Any calculations.

\- Any scenario output structure.



\### Test (manual)

\- Open AMI tab → click a control → click worksheet cells → AMI tab remains selected.

\- Repeat with dropdown controls and MIH/UAP buttons.



---



\## Fix-03 — Rent Roll YEAR selector + Manage uploads + mismatch warning (MIH + UAP)

\### Requirements

\- Add ribbon year selector: 2022, 2023, 2024, 2025, 2026 (default 2025).

\- Add a “Manage Rent Roll Years…” UI (form preferred) to upload/replace year files.

\- Uploaded year files must persist locally (suggested: %APPDATA%\\\\AMI\_Optix\\\\RentRollYears\\\\<year>\\\\).

\- Upload action should also transfer the file once to the API for storage (exact server storage mechanism can be decided, but client must not need to send files manually outside Excel).

\- Selected year must override any year implied inside the active workbook’s rent roll page.

\- If workbook “declared year” != selected year, show warning (OK to continue; not blocking).

\- Every MIH/UAP action that uses rent roll/utilities must use the selected year.



\### Must NOT change

\- Scenario grid structure.

\- Solver logic (except reading different year data).

\- Existing rent roll math except to swap source year data.



\### Test (manual)

\- Upload 2022 and 2024.

\- Select 2022 → run calc → confirm it uses 2022 tables.

\- Workbook declares 2025 but dropdown is 2022 → warning appears; OK continues.



---



\## Fix-01b — Utilities variant breakdown (top-of-sheet; do not collapse)

\### Problem

Utilities are displayed too generically; the sheet should show which specific variants are used.



\### Requirements

\- Do NOT alter scenario grid structure or per-scenario columns.

\- Add a top-of-sheet utility breakdown block that shows selected variants (exact names/labels as in rent roll guidelines).

\- Ensure rent roll net rent uses the correct total utility allowance based on selected variants.

\- Utilities are already variant-aware in the UI + payload (electricity/cooking/heat/hot\_water). Ensure display and mapping do not collapse variants incorrectly.



\### Must NOT change

\- Existing payload keys unless confirmed needed.

\- Scenario output structure.



\### Test (manual)

\- Pick a non-default heat/hot water variant → run → top-of-sheet shows exact variant and allowance matches.



---



\## Fix-02 — Solver dedupe identical outcomes + placement tie-break (Python)

\### Problem

Solver returns “duplicate-looking” scenarios that are outcome-identical but differ only by unit swaps.



\### Requirements

\- Only dedupe when outcomes are truly identical:

&nbsp; - same AMI band mix

&nbsp; - same rent outcome metrics used for client results

&nbsp; - same relevant mix/sqft outcome if part of outputs

\- Keep only one scenario per equivalence group.

\- Choose the kept scenario using placement tie-break:

&nbsp; - 40% units prefer lower floors

&nbsp; - higher AMI / higher paying units prefer higher floors

\- Must be surgical: do not change solver’s main scoring/constraints; apply as post-processing right before returning results.



\### Test (manual or unit test)

\- Two identical-outcome scenarios differing only by floor swap → only one returned, preferred placement chosen.

\- Two scenarios with same AMI mix but different rent totals → BOTH remain.



---



\## Fix-05 — Manual working-copy always-on + 2-way sync (MIH/UAP + rent roll + year)

\### Goal

Remove the manual scenario on/off toggle completely and make the “manual scenario” the always-present editable working copy.



\### Requirements

\- There is always an editable Manual Working Copy.

\- Stored solver scenarios (1..N) remain immutable.

\- Selecting scenario #k copies it into Manual Working Copy and applies it to MIH/UAP, but does not change scenario #k snapshot.

\- Edits to Manual Working Copy:

&nbsp; - are allowed even if violating rules (warn but allow override)

&nbsp; - immediately update MIH/UAP AMI column at matching unit

&nbsp; - immediately update rent roll calculations (including using selected rent roll YEAR)

\- Edits to MIH/UAP AMI column:

&nbsp; - immediately update Manual Working Copy

&nbsp; - update rent roll calculations

\- Must include re-entrancy guard to prevent infinite change-event loops.



\### Must NOT change

\- Scenario snapshots for 1..N.

\- Solver logic.

\- Scenario grid structure.



\### Test (manual)

\- Select scenario #2 → manual reflects it; scenario #2 snapshot unchanged.

\- Edit manual unit AMI → MIH/UAP updates instantly; rent roll updates.

\- Edit MIH/UAP AMI cell → manual updates; rent roll updates.

\- Repeat after switching sheets and switching selected rent roll year.



