# Task Card — Fix-01 DataReader stability + remove ribbon sheet activation

## Goal
- Prevent `ReadUnitData()` from selecting the wrong worksheet after UI interactions change `ActiveSheet`.
- Remove rent roll dropdown side effects that activate worksheets (to avoid context drift that can lead to wrong sheet/AMI-column caching).

## Success Criteria
- Running the solver from the ribbon reads unit data from the intended program sheet even if the user is currently on non-data tabs (e.g., “AMI Scenarios” / diagnostics / other sheets).
- `GetDataSheet()` and `GetAMIColumn()` remain consistent for subsequent “apply scenario” operations.
- Selecting a rent roll from the ribbon dropdown does **not** activate/switch the current worksheet.

## Files to Change
- `excel-addin/src/AMI_Optix_DataReader.bas`
- `excel-addin/src/AMI_Optix_Ribbon.bas`
- `docs/CODEX_LEDGER.md` (work log update)

## Functions / Entrypoints
- `FindDataSheet()` in `AMI_Optix_DataReader.bas`
- `Ribbon_SelectRentRoll(...)` in `AMI_Optix_Ribbon.bas`

## Proposed Patch (logic-level)
1) **DataReader sheet selection guardrails**
   - Keep supporting “active sheet” reads when the active sheet is a known program sheet name (e.g., UAP/MIH/RentRoll variants).
   - When `ActiveSheet` is *not* a known program sheet, prefer workbook sheets with known template names (UAP/MIH/RentRoll/Project Worksheet/etc) before falling back to arbitrary sheets.
   - Expand the preferred-name list to include MIH + common Rent Roll variants so MIH flows remain stable without depending on incidental sheet activation order.

2) **Ribbon rent roll dropdown**
   - Update selection state (`m_SelectedRentRoll`) only.
   - Remove `Worksheet.Activate` from `Ribbon_SelectRentRoll` and avoid modal UI (no `MsgBox`) for normal selection.

## Risk Pre-mortem
- **Risk:** Workbooks with multiple “unit-table-looking” sheets may now resolve differently (preferring template-named sheets over a custom active sheet).
  - **Mitigation:** Keep a fallback path that uses `ActiveSheet` only when no template-named sheet matches; keep the full-workbook scan as the last resort.
- **Risk:** MIH workbooks where the data sheet name deviates from expected template names could still be ambiguous.
  - **Mitigation:** Ensure fallback scan still finds a sheet with headers; avoid excluding user sheets beyond known AMI Optix output sheets.

## Test Plan
- Manual (Excel):
  1. Open a workbook that includes “UAP” (and/or “MIH”/“RentRoll”) plus the “AMI Scenarios” sheet.
  2. Activate “AMI Scenarios”, then click **AMI Optix → Run Solver**.
     - Expected: Units are read from the program sheet (UAP/MIH), and the solver completes normally.
  3. Select a rent roll in the dropdown.
     - Expected: The active sheet does not change.
  4. Apply a scenario from “AMI Scenarios”.
     - Expected: AMI values write back to the correct data sheet column (no “wrong tab” writes).

## Rollback Plan
- Revert the fix commit (single-commit revert) and redeploy the prior SHA if needed.
