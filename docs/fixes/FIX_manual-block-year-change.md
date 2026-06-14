# FIX: year-dropdown no longer reverts the manual block to original input

**Branch:** `fix/manual-block-year-change`
**Cut from:** `feature/excel-agent-foundation`
**Date:** 2026-06-14
**Risk:** Low — VBA-only. Year-dropdown path only; the "Manual Calculate" button is unchanged. New optional flag defaults to existing behavior, with a safe fallback.

## Symptom (client, 2026-06-14)

After Run MIH, the manual block (Scenario Manual / "best of 40", the recommended fewest-40 scenario) is pinned at the top. But selecting a **different rent-roll year** from the dropdown silently **swapped the manual block back to the original input** — the user's hand-entered AMIs — instead of just re-pricing the applied scenario. The 5 solver blocks below kept their bands correctly; only the manual block reverted.

## Root cause

Year change fires [`Ribbon_SelectRentRollYear`](excel-addin/src/AMI_Optix_Ribbon.bas) → `ManualCalculateScenario` → `ReadUnitData()`, which reads the **raw AMI column on the input sheet** (the original input) and rebuilds the manual block from it via `WriteManualScenarioBlockFromEvaluate`.

The solver blocks survive because [`RecalculateSolverScenarioRents`](excel-addin/src/AMI_Optix_ResultsWriter.bas) re-reads **each block's own displayed AMI** and only re-prices it. The manual block was the lone exception — regenerated from input rather than re-priced in place — so the recommended scenario pinned at [WriteManualScenarioBlockFromResult line ~2861](excel-addin/src/AMI_Optix_ResultsWriter.bas) (`GetBestScenarioKey`) was discarded on every year change.

## Change (VBA only)

1. `ManualCalculateScenario` gains `Optional preserveAppliedScenario As Boolean = False`. When True it sources units from the **manual block's own displayed assignments** instead of the raw input sheet; if that can't be read it falls back to `ReadUnitData()` (today's behavior).
2. New private helper `ReadUnitsFromManualBlock(ws)` reads the manual block's `Unit / Bedrooms / Net SF / AMI` table (the first such table on the sheet, above the first numbered `SCENARIO N:` header). Column mapping mirrors `RecalculateSolverScenarioRents` exactly. Returns `Nothing` when no readable table exists.
3. `Ribbon_SelectRentRollYear` now calls `ManualCalculateScenario(prog, True)`. The "Manual Calculate" ribbon button still calls it with the default `False`, so it keeps reading the input sheet — correct for that button.

Net effect: a year change re-prices the applied/recommended bands (and any manual edits) at the new year; it no longer reverts to original input. WAAMI/tradeoffs stay correct because the rent API is still called — just with the applied bands as input.

## Verified

- Static checks: proc balance (ResultsWriter 16/16 Sub, 52/52 Function; Ribbon 41/41 Sub, 23/23 Function), single `Dim units` in the handler, no duplicate `Dim`s in the new helper, both preserve-path and fallback present, ribbon passes `True`.
- Manual block uses the same `WriteScenarioSummaryAndTable` as solver blocks → identical table layout, so the read logic is the already-proven solver re-pricer mapping.

## Deploy

VBA-only → standard `.xlam` refresh (GitHub Refresh button or the from-GitHub PowerShell one-liner). No Render deploy.
