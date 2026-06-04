# FIX: VBA electricity parser + AMI-cell jump-back

**Branch:** `fix/vba-electricity-parser-and-cell-jump`
**Cut from:** `feature/excel-agent-foundation` @ `3eb1db3`
**Date:** 2026-06-04
**Author:** client PC reported repeated popup + cursor jump-back on every AMI edit
**Risk:** Low (both fixes are surgical; no logic restructuring)

## Two distinct problems, one fix

### Problem 1 — "Failed to build rent tables cache" popup on every AMI edit

Symptom: open MIH/UAP page, enter a new value in any AMI cell, press Enter — popup fires:

```
Failed to build rent tables cache.
Year: 2026
Workbook: ...\RentCalculator_2026.xlsx
Error: Utility allowances validation failed: Utility allowances missing required key.
Key: electricity|tenant_pays|studio
```

Root cause: **the same latent parser bug we fixed in Python last week, but here in VBA.** The Apartment Electricity column (col 2 in the rent calculator sheet) has no per-option label in row 16 — it's a single-option binary. On a fresh HPD template the dropdown in row 17 reads "N/A or owner pays" (default state). The VBA parser at [AMI_Optix_RentTables.bas:509-511](excel-addin/src/AMI_Optix_RentTables.bas#L509-L511) falls back to row 17, uses "N/A or owner pays" as the option label, then drops the column entirely because that label doesn't map to a paid variant. Result: `electricity|tenant_pays|<bedroom>` rows never get added → validation fails → popup.

Why it fires every edit: each AMI cell edit triggers `RefreshManualWorkingCopyLocalRents`, which loads the rent cache. With the cache build failing, every edit re-attempts the build and re-shows the popup.

Fix: same 3-line patch as the Python one. After computing `optionCat` and `optionLabel`, recognize the electricity-column special case and force the canonical "Tenant Pays" label:

```vba
If optionCat = "electricity" And LCase$(optionLabel) = "n/a or owner pays" Then
    optionLabel = "Tenant Pays"
End If
```

Now col 2's values get registered under `electricity|tenant_pays|<bedroom>`, validation passes, cache saves, popup never fires.

### Problem 2 — Cursor jumps back to previous row after pressing Enter

Symptom: edit AMI cell P42, press Enter, cursor should move to P43 — instead it jumps back to P42 (or whatever cell was just edited).

Root cause: [AMI_Optix_AppEvents.cls:139 and :185](excel-addin/src/AMI_Optix_AppEvents.cls#L139) had:

```vba
If Not contextSheet Is Nothing Then
    contextSheet.Activate
    If Not contextTarget Is Nothing Then
        contextTarget.Cells(1, 1).Select   ← THIS IS THE BUG
    End If
End If
```

Excel commits the cell edit on Enter, moves the cursor down naturally, then fires the `SheetChange` event with the OLD cell as `Target`. The handler then forces selection back to that old cell, overriding Excel's natural Enter behavior.

Fix: remove the `.Cells(1, 1).Select` line. Keep the `.Activate` (so the user's view returns to the original sheet if the refresh happened to activate another). Net effect: after editing an AMI cell and pressing Enter, the cursor stays where Excel moved it — the next row — as the user expects.

## Files modified

- `excel-addin/src/AMI_Optix_RentTables.bas` — 7 new lines in `ExtractUtilityAllowances` (the electricity parser fix)
- `excel-addin/src/AMI_Optix_AppEvents.cls` — removed two 3-line blocks (the `.Select` calls in `RefreshManualWorkingCopyFromProgramInputs` and `ApplyManualAmiChangesAndRefresh`)

No Python change. No solver change. No new tests (Python suite still 63 pass — VBA isn't exercised by pytest).

## Client-PC update

Since both fixes are in VBA `.bas` / `.cls` files (bundled into the `.xlam` during build), the client needs to refresh their local AMI Optix Agent:

```
powershell -ExecutionPolicy Bypass -File C:\AMI_Optix_Agent\scripts\Refresh-AmiOptixAgentFromGitHub.ps1 -AgentRoot C:\AMI_Optix_Agent
```

After it completes, **close Excel completely**, reopen, re-enable the add-in if prompted, then test by editing any AMI cell — should advance to the next row with no popup.

## Verification (client PC, after refresh + Excel restart)

1. Open the workbook on MIH sheet
2. Click into any AMI cell (e.g., P42)
3. Type 60% (or any AMI value)
4. Press Enter
5. **Expected:** cursor moves to P43, no popup
6. Edit another row
7. **Expected:** still no popup, cursor advances naturally
8. (If the popup ever appears once, e.g., very first edit on cold start, it's because the parser is building the cache for the first time — but after this fix it should build successfully without erroring)

## What this does NOT change

- Server-side Python (already had the parser fix)
- The solver / scenario generation
- The rent calculator XLSX files
- The 100% AMI haircut
- The `low_40_share` / `max_40_share` scenarios
- The strict WAAMI cap enforcement
- Any other event handlers (only the two `.Select` lines were touched in AppEvents)

## Risks

- **Edge case**: if a workbook has a different rent calculator with a deliberately-set "N/A or owner pays" for Apartment Electricity (genuinely owner-pays), the fix would still parse it under `electricity|tenant_pays` with the dollar values from the IF formula. That's correct behavior — the rent table stores "what the rent would be IF tenant paid", not "what the current dropdown selection is."
- **Cell jump fix**: the only callers of `RefreshManualWorkingCopyFromProgramInputs` and `ApplyManualAmiChangesAndRefresh` are the AMI-edit event handlers. Removing the `.Select` doesn't affect any other code path.
