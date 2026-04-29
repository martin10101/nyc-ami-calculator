# FIX: Utility selections must drive rent calculation (MIH + UAP)

**Branch:** `fix/utility-selections-respected`
**Cut from:** `feature/excel-agent-foundation` @ `bbb4076`
**Date:** 2026-04-29
**Author:** client request via remote session
**Risk:** Medium (precedence change for utility source-of-truth, plus a new auto-recalc on year-change)

## Status

| Step | When | Result |
|---|---|---|
| Commit 1 (Registry-only utility source) | 2026-04-29 — sha `1744343` | ✅ |
| Commit 2 (year-change auto-recalc) | 2026-04-29 — sha `08cf9e5` | ⏳ awaiting client test |
| Commit 3 (this doc) | 2026-04-29 — sha _pending_ | ⏳ |
| Pushed `fix/utility-selections-respected` | _pending_ | ⏳ |
| Fast-forward merged into `feature/excel-agent-foundation` | _pending_ | ⏳ |
| Pushed `feature/excel-agent-foundation` (triggers Render auto-deploy) | _pending_ | ⏳ |
| Client PC refreshed via PS agent | _pending_ | ⏳ |
| Manual test passed on client PC | _pending_ | ⏳ |
| Approved by client | _pending_ | ⏳ |

## One-line PowerShell command for client-PC update

Run on the client PC (cmd.exe, Run dialog, or PowerShell):

```
powershell -ExecutionPolicy Bypass -File C:\AMI_Optix_Agent\scripts\Refresh-AmiOptixAgentFromGitHub.ps1 -AgentRoot C:\AMI_Optix_Agent
```

Standard one-liner — same as every other fix in this project.

## Problem

The client picks utility options (electricity / cooking / heat / hot water — tenant pays vs owner pays, plus the type when tenant pays) in the **Configure Utilities** ribbon dialog. The "Utilities — Selected Variants (Affects Rent Allowances)" block on the AMI Scenarios sheet was showing values that didn't match what was just picked — typically every category rendered "YES" — and per-unit net rents looked like the chosen allowance amounts had been ignored.

Separately: changing the **Rent Roll Year** dropdown (2024 ↔ 2025) updated only the cached rent table; it did not refresh the scenarios already on the sheet, so the user had to click **Run MIH** or **Run UAP** again after every year change.

## Diagnosis (Phase 1, read-only)

End-to-end trace of the utilities path. The wire and the server were both clean; the bug lived in the VBA precedence logic.

**Server is fine:** `/api/optimize` and `/api/manual_calculate` correctly extract `data.get('utilities', {})`, default missing fields to `'na'`, and forward to `compute_rents_for_assignments`. JSON keys match between VBA and Python (`"utilities"` on both sides). Option key strings are identical (`tenant_pays`, `gas`, `oil`, `electric_ccashp`, `electric_other`, `electric_heat_pump`, `na`). Both 2024 and 2025 rent-calculator layouts use the same canonical keys.

**Root cause: form writes to one place, API reads from a different place.**
- Form `frmUtilities.btnSave_Click` and the ribbon InputBox path write only to the **Windows Registry** via `SaveUtilitySelections`.
- `GetUtilitySelectionsForProgram` (the function that builds the API payload's utility values) read the **workbook sheet first**:
  - MIH → `TryReadMIHUtilities` reads "Rents & Utilities" rows 14-16
  - UAP → `TryReadUAPUtilities` reads "Calculations" cells P3:AA3
- Workbook-read fell back to Registry **only when the workbook read failed**. On success, it then **overwrote** the Registry with the workbook values, silently invalidating whatever the user had just chosen in the form.
- The "Selected Variants" display block reads the same Registry — so it echoed back the post-overwrite workbook values, not the user's form picks.
- The "all YES" appearance was the display logic at `WriteUtilitySettings` rendering YES whenever the value was anything other than `"na"`. Property templates with "Tenant Pays" markers in the relevant cells map every category to a non-`na` key → every row reads YES.

**Year-change handler:** `Ribbon_SelectRentRollYear` saved the year and warmed the rent-table cache but did not invoke any scenario recalc.

## Code change

### Commit 1 — `1744343` — Form is the source of truth

**File:** `excel-addin/src/AMI_Optix_Main.bas`
**Function:** `GetUtilitySelectionsForProgram` (lines 617-634)

Replaced the workbook-precedence body with a Registry-only read:

```vba
Public Function GetUtilitySelectionsForProgram(program As String) As Object
    Dim utils As Object
    Set utils = CreateObject("Scripting.Dictionary")
    utils("electricity") = GetSetting("AMI_Optix", "Utilities", "electricity", "na")
    utils("cooking") = GetSetting("AMI_Optix", "Utilities", "cooking", "na")
    utils("heat") = GetSetting("AMI_Optix", "Utilities", "heat", "na")
    utils("hot_water") = GetSetting("AMI_Optix", "Utilities", "hot_water", "na")
    Set GetUtilitySelectionsForProgram = utils
End Function
```

`TryReadMIHUtilities` and `TryReadUAPUtilities` are left in place as dead helpers so a future revert just restores the function body — no other call sites depend on them being removed.

**Net effect:** form / InputBox writes go to the Registry → API picks them up → rent calc applies them → display block reads the Registry → all four touchpoints agree. Manual Calculate's per-scenario loop inherits the fix automatically (it already calls `GetUtilitySelectionsForProgram` per scenario).

### Commit 2 — `08cf9e5` — Year change triggers full recalc

**File:** `excel-addin/src/AMI_Optix_Ribbon.bas`
**Function:** `Ribbon_SelectRentRollYear` (lines 774-813) and new helper `HasExistingSolverScenarios` (lines 815-848)

After the existing year-save / cache-warm / mismatch-warning logic, added:

```vba
' Recalculate all scenario rents (5 solver scenarios + Scenario Manual) using
' the new year x current utility selections. Skip silently if scenarios are
' not yet on the sheet (first-time year change before any Run MIH/UAP) or if
' the API key is unset, so we don't fire surprise popups for what is just a
' dropdown change.
On Error Resume Next
If HasAPIKey() Then
    If HasExistingSolverScenarios() Then
        Call ManualCalculateScenario(DetectProgramFromWorkbook())
    End If
End If
On Error GoTo 0
```

Reuses `ManualCalculateScenario` — the same public function the **Manual Calculate** ribbon button already calls (`Ribbon_ManualCalculate` at Ribbon.bas:351). Recalc path is identical to a manual click. The new `HasExistingSolverScenarios` private helper scans the "AMI Scenarios" sheet for any "SCENARIO N:" header (excluding the Manual block), returning False before the user has run the optimizer at least once. The `HasAPIKey` gate prevents the API-key MsgBox from firing for a dropdown interaction.

## Files NOT touched (verified clean during diagnosis)

- `app.py` — server correctly threads `utility_selections`
- `ami_optix/rent_calculator.py` — option keys are consistent across 2024 / 2025
- `ami_optix/solver.py` — utility-agnostic
- `excel-addin/forms/frmUtilities.frm` — form already saves to Registry correctly
- `excel-addin/src/AMI_Optix_API.bas` — JSON contract is correct
- `excel-addin/src/AMI_Optix_ResultsWriter.bas` — display block already reads Registry, which is now the canonical source
- Ribbon XML — untouched
- `TryReadMIHUtilities` / `TryReadUAPUtilities` in `Main.bas` — left in place as dead helpers; revert-friendly

All edits are `.bas`-only → fully PowerShell-deployable per `Refresh-AmiOptixAgentFromGitHub.ps1`.

## Manual test (required before merging back to `feature/excel-agent-foundation`)

**Test workbook:** `230 Kent_Unit Schedule - MIH v3.xlsm` (and a UAP workbook for cross-program coverage).

### After Commit 1 — form drives

1. Open the MIH workbook. Click ribbon → **Configure Utilities** → set **all four** to "N/A or owner pays" → Save.
2. Click **Run MIH**. Wait for scenarios.
3. ✅ Scroll to "UTILITIES — Selected Variants". All four rows show **NO**. Per-unit net rent ≈ gross rent (no allowances applied — small differences only from rent-table rounding).
4. Open Utilities again → set **Heat = "Gas Heat"**, **Hot Water = "Gas Hot Water"**, others stay N/A → Save.
5. Click **Run MIH** again.
6. ✅ Selected Variants shows: Heat = YES (Gas Heat), Hot Water = YES (Gas Hot Water), Electricity = NO, Cooking = NO. Per-unit net rents drop by approximately the gas heat + gas hot water allowance amounts vs step 3.
7. Open Utilities → set **Electricity = "Tenant Pays"**, **Cooking = "Gas Stove"** (in addition to step 4's heat picks) → Save → **Run MIH**.
8. ✅ All four rows show YES with the correct types. Per-unit net rents drop further.
9. Repeat steps 1, 4, 5 on a UAP workbook → ✅ same expected behavior.

### After Commit 2 — year-change recalc

10. With the workbook from step 7 still open (utilities all tenant-pays), note total annual rent for Scenario 1.
11. Click ribbon → **Rent Roll Year** dropdown → switch to **2024**.
12. ✅ Within ~5 seconds, scenarios on the AMI Scenarios sheet recalculate. Total annual rent for Scenario 1 reflects the 2024 gross-rent table × current (all-tenant-pays) utility selection.
13. Switch back to **2025**.
14. ✅ Total annual rent reverts to the value from step 10.
15. Repeat 10-14 on a UAP workbook → ✅ same expected behavior.

### Edge cases worth checking

16. Open a fresh property workbook (no scenarios yet). Change the year dropdown → ✅ no popup, no error, no scenario sheet created. (`HasExistingSolverScenarios` returns False; recalc is skipped.)
17. Without setting up the API key, change the year on a workbook that DOES have scenarios → ✅ no API-key warning popup. (`HasAPIKey` returns False; recalc is skipped.)

### Server-side regressions

```
python -m pytest tests/test_rent_calculator.py tests/test_api_evaluate.py tests/test_api_optimize_learning.py -v
```

All existing tests pass (no Python changes in this fix).

## Behavior change worth flagging to the client

After Commit 1, the **workbook utility cells are no longer auto-imported** into the Registry. Property templates with "Tenant Pays" markers in "Rents & Utilities" rows 14-16 (MIH) or "Calculations" P3:AA3 (UAP) **no longer** seed the API call. The user picks utilities in the **Configure Utilities** dialog, which writes to the Registry, and the Registry's values are what every Run MIH / Run UAP / Manual Calculate / year-change-recalc uses.

This matches the user's stated workflow ("the client selects utilities in the ribbon box"). Any client who today relied on the workbook-read path will need to open the Utilities dialog once after the deploy.

## Rollback

```bash
git checkout feature/excel-agent-foundation
git revert 1744343           # restores workbook-precedence GetUtilitySelectionsForProgram
git revert 08cf9e5           # removes the year-change auto-recalc
```

Either commit can be reverted independently. Reverting Commit 1 alone restores the workbook-precedence behavior. Reverting Commit 2 alone keeps the form-source-of-truth fix but removes the year-change auto-recalc (user must click Run again after a year change, as before).

After revert, re-run the PS agent on the client PC to regenerate the staged `.xlam` from the reverted source.

## Deploy notes

- Render auto-deploy is **ON**. Pushing the merged feature branch triggers a Render deploy. The change is VBA-only — Render serves no VBA — so the deploy adds nothing functional, but it does happen.
- Client-PC rollout is via PowerShell agent only; no manual VBA copy / paste in the Excel VBA editor.
- Per the project's "no Render-wait" rule: the PS one-liner above can be run immediately after the push; it pulls source from GitHub directly and is independent of the Render deploy state.
