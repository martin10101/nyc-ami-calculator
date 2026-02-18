# Task Card — Fix-06: Local Rent Calculation for Manual Working Copy (Z:\ + AppData fallback)

Date: 2026-02-18  
Repo: martin10101/nyc-ami-calculator-all-fixes-test  
Base branch: main  
Work branch: fix/06-local-manual-rent-calc

## Goal
- Keep API-provided rents/totals for solver scenarios (no change to scenario outputs / grid).
- Stop calling `/api/manual_calculate` on Manual Working Copy edits (no per-edit API calls).
- Recompute Manual Working Copy rents/totals locally in Excel/VBA using the selected year’s rent calculator workbook.
- Use shared authoritative path first (`Z:\AMI_Optix\RentRollYears\<YEAR>\`), with `%APPDATA%` fallback.

## Success Criteria (from Fix-06 spec)
- Running solver still shows scenario rents/totals exactly as returned by `/api/optimize`.
- Editing Manual Working Copy AMI updates MIH/UAP AMI column and refreshes manual rents/totals locally **without** calling `/api/manual_calculate`.
- Editing MIH/UAP AMI column refreshes Manual Working Copy + manual rents/totals locally.
- Z:\ missing/unavailable → AppData fallback works.
- Changing Rent Roll Year changes which workbook is used for the next local refresh.
- Re-entrancy guard prevents infinite SheetChange loops.

## Files to change (keep minimal)
- `excel-addin/src/AMI_Optix_ResultsWriter.bas`
- `excel-addin/src/AMI_Optix_AppEvents.cls`
- (Optional) `excel-addin/customUI/customUI14.xml`, `excel-addin/src/AMI_Optix_Ribbon.bas` for “Verify Manual Rents (API)” button

## Functions / Entrypoints to change
- `AMI_Optix_AppEvents.cls`
  - `RefreshManualWorkingCopyFromProgramInputs(...)`
  - `ApplyManualAmiChangesAndRefresh(...)`
- `AMI_Optix_ResultsWriter.bas`
  - New: `RefreshManualWorkingCopyLocalRents(...)` (rebuild manual block + compute rents locally)
  - New: local rent schedule loader + cached workbook handle
  - New: gross rent + utility allowance lookups against “AMI & Rent” sheet

## Proposed Patch (logic-level)
- Implement a local rent schedule cache:
  - Resolve workbook path by year:
    - Prefer `Z:\AMI_Optix\RentRollYears\<YEAR>\RentCalculator_<YEAR>.xlsx` (or first `.xlsx/.xlsm` found in folder).
    - Fallback `%APPDATA%\AMI_Optix\RentRollYears\<YEAR>\RentCalculator_<YEAR>.xlsx`.
  - Open workbook read-only and hide its window; cache handle; re-open only when year changes.
  - Parse “AMI & Rent” into dictionaries (gross rents and allowances) similar to `ami_optix/rent_calculator.py`.
- Replace per-edit `/api/manual_calculate` refresh:
  - On any manual/program AMI edit, rebuild the manual block from current workbook units and compute:
    - gross rent (AMI + bedroom)
    - utility allowances (bedroom + utility selections)
    - net rent and annual rent
    - totals
  - Write results into the existing Manual Working Copy table/summary (no scenario grid changes).
- Keep existing event suppression (`g_AMIOptixSuppressEvents` + `Application.EnableEvents=False`) to avoid loops.

## Risks + mitigations
- **Wrong rent workbook path** → show warning/tradeoff; leave rents blank/0; do not call API.
- **Workbook visibility / focus disruption** → open with events off, hide window, restore user sheet/selection.
- **Parsing mismatch vs API** → optional verify button can compare against `/api/evaluate` once on demand.

## Test Plan
### Automated
- `python -m pytest -q`

### Manual (Excel)
1) Run solver once; scenario rents/totals still appear (from API).
2) Edit a Manual Working Copy AMI cell:
   - MIH/UAP AMI updates immediately
   - Manual rents/totals update immediately
   - Confirm no `/api/manual_calculate` call (watch API logs or temporarily disconnect network).
3) Make Z:\ unavailable (or test on a machine without Z:\), repeat: rents still update using AppData.
4) Change Rent Roll Year (e.g., 2024 ↔ 2025), edit again: rents reflect selected year workbook.
5) (Optional) Verify button: compare local totals to `/api/evaluate` for a valid assignment.

## Rollback Plan
- Revert the Fix-06 commit(s) to restore `/api/manual_calculate`-based manual refresh.

