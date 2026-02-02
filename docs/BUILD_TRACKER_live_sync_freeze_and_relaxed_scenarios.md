# AMI Optix Build Tracker — Live Sync Freeze + 6 Scenarios

Last updated: 2026-02-02  
Branch: `feature/results-overhaul-2026-01-28`

## Goals (user-confirmed)

### A) Always return **up to 6 scenarios** (both UAP and MIH)
- **3 strict**: follow all rules; stay as close to 60% WAAMI as possible.
- **3 relaxed**: maximize rent; WAAMI can go **below 60% down to 58%** (but should stay higher if feasible).
- **MIH relaxed**: keep the **10% share at ≤40%** rule enforced 100%; only relax WAAMI floor down toward 58%.

### B) “Freeze Live Sync” mode (Excel add-in)
- Add **Live Sync ON/OFF** toggle.
- When **OFF**:
  - Clear only the **Scenario Manual** block on `AMI Scenarios` (do not touch Scenario 1/2/3 tables).
  - Clear the **entire AMI column** (below header) on the program sheet (`UAP` or `MIH`) so the user can type their own AMIs.
  - Stop auto-refreshing the manual block on every edit (no “fighting” the user while typing).
- Add **Manual Calculate** action:
  - Reads whatever AMIs the user entered on the program sheet.
  - Computes WAAMI, bands used, band mix (units/SF/share), allowances, gross/net rents, totals.
  - Must compute even if the assignment is “non-compliant”; show warnings/tradeoffs instead of reverting.

### C) MIH sheet preference (Excel add-in)
- For MIH, prefer reading/writing the unit table from the `MIH` sheet first (fallback only if needed).

## Already implemented (before this tracker)
- MIH uses **Net Floor Area** as `mih_residential_sf` (fallback to `MIH!J21`).
- Fixed MIH AMI scaling so `120%` is written as `120%` (not `1%`).

## In progress / next steps
### Implemented
1) Backend: MIH now returns up to 6 scenarios (strict + relaxed + fallback fill).
2) Backend: added `/api/manual_calculate` for lenient “manual calculate” (returns rents + tradeoffs even if invalid).
3) VBA: added Live Sync toggle + Manual Calculate action; Live Sync OFF clears Scenario Manual + clears the full AMI column.
4) VBA: MIH workflows now prefer the `MIH` sheet first (fallback only if needed).

### Validation notes (local)
- File: `C:\Users\MLFLL\Downloads\files\Unit Schedule_2-35-11, MIHtest3.xlsb`
- `/api/optimize` (MIH Option 1) now returned **6** scenario keys:
  - `absolute_best`, `best_3_band`, `best_2_band` (strict)
  - `max_revenue`, `edge_waami_floor_590` (relaxed rent-max)
  - `client_oriented` (fallback fill when relaxed variants were not distinct)
- Example rents from that run:
  - `max_revenue`: ~$340,848 net annual (WAAMI ~60.0%)
  - `edge_waami_floor_590`: ~$360,048 net annual (WAAMI ~59.92%)

## Remaining questions / follow-ups
- Decide how we want to label “strict vs relaxed” in the UI when the 6th scenario is a strict fallback (e.g., `client_oriented`).
- Confirm the desired WAAMI floors for “relaxed” scenarios (currently we target 59% + 58% floors, but 58% may be non-distinct for some projects).

## Fixes (2026-02-02)
- Fixed Excel ribbon `getPressed` callback signature for the Live Sync toggle (was causing: “Wrong number of arguments or invalid property assignment” when clicking the AMI Optix tab).
- Forced the `AMI Scenarios` sheet to `xlSheetVisible` whenever the add-in creates/uses/activates it (prevents the tab from “disappearing” in client workbooks that hide sheets via macros).
- Preserved the user’s active sheet after applying a scenario (avoids confusion where it looks like the scenarios sheet “flew away”).
- Added a sheet-visibility guard (Application `SheetActivate` + deferred `OnTime`) to re-unhide `AMI Scenarios` / `AMI Optix Diagnostics` even if the client workbook hides them after the fact.
- Added program mismatch guard rails: block **Run UAP** on MIH workbooks and block **Run MIH** on non-MIH workbooks; MIH requires `Prog!K4` to be Option 1 or Option 4.
- Fixed a common MIH utility parsing crash (`Type mismatch`) when the `Rents & Utilities` sheet contains Excel error values in the selection cells.
- Improved unit reading for solver runs to always read from the correct program sheet (so running MIH/UAP from `AMI Scenarios` does not cause “no data”).
