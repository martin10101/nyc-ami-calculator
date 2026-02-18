# Task Card — Fix-01b Utilities variant breakdown (top-of-sheet; do not collapse)

## Goal
- Make the top-of-sheet utilities block show the **exact utility variant labels** used by the rent roll guidelines (e.g., “Gas Stove”, “Gas Heat”, “Electric Hot Water - Heat Pump”, “Electric Heat - … (ccASHP)1”).
- Prevent “collapsed”/generic labels (e.g., showing just “Gas” for multiple utilities).
- Keep scenario grid + per-scenario columns unchanged.

## Success Criteria (from `docs/FIX_REQUIREMENTS.md`)
- Scenario grid structure and per-scenario columns are unchanged.
- Top-of-sheet utility breakdown block shows selected variants using the exact names/labels as the rent roll guidelines.
- Net rent continues to reflect the correct **total** utility allowance based on the selected variants.
- Utilities display/mapping does not collapse variants incorrectly (electricity/cooking/heat/hot_water).

## Files To Change
- `excel-addin/src/AMI_Optix_ResultsWriter.bas`

## Functions / Entrypoints To Change
- `WriteUtilitySettings(ws As Worksheet, startRow As Long) As Long`
- Replace/extend helper: `FormatUtilityType(...)` (or introduce a category-aware label helper)

## Proposed Patch (logic-level)
- Update the utilities block writer to display **category-aware** labels:
  - Cooking: “Electric Stove” / “Gas Stove” / “N/A or owner pays”
  - Heat: “Electric Heat - Cold Climate Air Source Heat Pump (ccASHP)1” / “Electric Heat - Other2” / “Gas Heat” / “Oil Heat” / “N/A or owner pays”
  - Hot Water: “Electric Hot Water - Heat Pump” / “Electric Hot Water - Other” / “Gas Hot Water” / “Oil Hot Water” / “N/A or owner pays”
  - Electricity: “Tenant Pays” / “N/A or owner pays”
- Keep the existing per-bedroom utility deduction totals table intact; this fix is display/mapping only.

## Risk Pre‑Mortem
- **Risk:** Labels drift from the backend/rent workbook over time.
  - **Mitigation:** Use strings that match the documented guideline labels and current rent calculator workbook labels (including the “1”/“2” heat footnotes).
- **Risk:** Changing the number of rows in the utilities block shifts where the manual scenario table starts.
  - **Mitigation:** Only change displayed strings (not row structure) so downstream row offsets remain stable.

## Test Plan
### Automated
- `python -m pytest -q`

### Manual (Excel)
- Open a workbook, set a **non-default** Heat and/or Hot Water variant (e.g., Heat = ccASHP, Hot Water = Gas).
- Run Optimize / Evaluate.
- In the “AMI Scenarios” sheet top block:
  - Verify the utilities table shows the exact variant labels (not generic “Gas/Oil/Electric”).
  - Verify the per-bedroom utility deduction totals change appropriately when switching variants.

## Rollback Plan
- Revert commit on this branch, or restore prior `WriteUtilitySettings` label formatting in `excel-addin/src/AMI_Optix_ResultsWriter.bas`.

