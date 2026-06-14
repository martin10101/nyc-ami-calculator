# FIX: SCENARIO OVERVIEW "Monthly Rent" now updates on recalc / year change

**Branch:** `fix/overview-rent-recalc`
**Cut from:** `feature/excel-agent-foundation`
**Date:** 2026-06-14
**Risk:** Very low — VBA-only, additive. No solver/server change. Touches one recalc routine + one new private helper.

## Symptom (client, 2026-06-14)

When the user runs **manual recalculate** or picks a **different rent-roll year** from the dropdown, every detail scenario block refreshes its rents correctly — but the **SCENARIO OVERVIEW** summary table at the top (the one-shot list of all scenarios) kept showing the **Monthly Rent** from the year the sheet was first written. Everything else updated; just that column was stale.

## Root cause

The sheet is value-based, not formula-based: rents are written as plain numbers and refreshed by a VBA routine, `RecalculateSolverScenarioRents` ([excel-addin/src/AMI_Optix_ResultsWriter.bas](excel-addin/src/AMI_Optix_ResultsWriter.bas)). That routine walks each **detail** table, re-calls the rent API for the new year, and rewrites the per-unit cells plus the "Total Monthly Rent:" line in place. It **never touched the overview table**, so the overview's `Monthly Rent` column (col 5) stayed at its original value.

## Change (VBA only)

`RecalculateSolverScenarioRents`:
1. For each detail table, parse its scenario number from the `SCENARIO N:` header just above the unit table (the same number shown in the overview `#` column — both come from `BuildGroupedScenarioOrder`, so it's a reliable join key).
2. When the detail "Total Monthly Rent:" cell is refreshed, also stash the fresh `net_monthly` into a `freshRentByNum` dictionary keyed by that scenario number.
3. After all tables are processed, call the new helper to push those values into the overview.

New private helper `UpdateOverviewRentColumn(ws, freshRentByNum, lastRow)`:
- Scans the sheet for any overview table (identified by its `Monthly Rent` header in column 5) and, for each scenario row, rewrites column 5 from `freshRentByNum` keyed by the row's `#` number (handles the `> N` recommended-marker prefix). Stops at the legend (`HOW ...`) or the first detail `SCENARIO ` block, so it never strays into unit tables. Updates **all** overview tables on the sheet (main + manual block) for consistency.

No change to how the overview is first written — initial render was already correct; only the recalc path was missing the update.

## Verified

- Static checks: proc balance (16 Sub / 16 End Sub, 51 Function / 51 End Function), no duplicate `Dim`s in the routine, GoTo/label resolve, no leading-`=` cell writes.
- Logic: overview `#` and detail `SCENARIO N:` are the same `BuildGroupedScenarioOrder` index — confirmed in source.

## Deploy

VBA-only → standard `.xlam` refresh (GitHub Refresh button or the rebuild one-liner). No Render deploy needed.
