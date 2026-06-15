# FIX: MIH page and Scenario Manual now hold the SAME scenario (re-synced)

**Branch:** `fix/mih-page-manual-block-sync`
**Cut from:** `feature/excel-agent-foundation`
**Date:** 2026-06-15
**Risk:** Low–medium — two targeted VBA changes. No solver, no server. Restores synchronization between the MIH AMI column and the Scenario Manual block.

## Symptoms (client, 2026-06-15)

- After **Run MIH**, the MIH page AMI column and the **Scenario Manual** block showed **different** scenarios.
- After the prior "keep displayed" change, **Manual Calculate stopped reflecting MIH edits** — the two were fully detached.

The client's model: Scenario Manual and the MIH AMI columns must represent the same thing (Scenario Manual is the canonical view of the MIH AMI columns).

## Root cause (the real one, unifying both this and the earlier "swap")

Run MIH writes the chosen scenario to two places that picked it **differently**:

- `ApplyBestScenario` writes AMI bands to the **MIH page** ([ResultsWriter.bas:63](excel-addin/src/AMI_Optix_ResultsWriter.bas#L63)). It chose from a hardcoded list starting with **`"absolute_best"`** (max-rent, higher-40%-count).
- The **manual block** uses `GetBestScenarioKey`, which prefers **`recommended_key`** (fewest-40).

So Run MIH put `absolute_best` on the MIH page and `recommended` in the manual block → they disagreed. The earlier "Manual Calculate swaps to a higher 40% count" was the *same* mismatch surfacing: Manual Calculate read the MIH page (`absolute_best`) and overwrote the recommended in the manual block. The previous "keep displayed" patch (preserve=True on the button) over-corrected and detached the two entirely.

## Changes

1. **[ResultsWriter.bas](excel-addin/src/AMI_Optix_ResultsWriter.bas) `ApplyBestScenario`:** apply the server's `recommended_key` (the same scenario the manual block shows), falling back to the legacy priority list only when `recommended_key` is absent (e.g. UAP). Read straight from `result` because `g_AMIOptixRecommendedKey` isn't set until `CreateScenariosSheet` runs afterward. → MIH page == manual block after Run MIH.
2. **[Ribbon.bas](excel-addin/src/AMI_Optix_Ribbon.bas) `Ribbon_ManualCalculate`:** reverted to `ManualCalculateScenario(programNorm)` (preserve=False) so the button reads the MIH AMI column and refreshes the manual block from it — keeping them in sync and reflecting MIH edits. Kept the `AMI_Optix_CancelDeferredRefresh` call (Ctrl+Z protection). With change #1 the MIH page now equals the recommended, so reading it no longer swaps the displayed scenario.

The rent-roll-year change keeps its in-place `preserve=True` behavior (unchanged; the client confirmed that one works), which is consistent because the two sources are now equal.

## Why this fixes both complaints at once

- After Run MIH: both the MIH page and the manual block get `recommended_key` → identical.
- Manual Calculate: reads MIH (== recommended) → manual block stays consistent and reflects any MIH edits → synced.

## Not changed (safety)

Live Sync, the 2-second edit debounce, and AMI input-normalization (the Ctrl+Z machinery) are untouched.

## Verified

- Static checks: proc balance (ResultsWriter 16/16 Sub, 52/52 Function; Ribbon 41/41 Sub, 23/23 Function); `ApplyBestScenario` prefers `recommended_key` with legacy fallback retained; no duplicate `Dim`s; button is back to preserve=False.

## Deploy

VBA-only → standard `.xlam` refresh (GitHub Refresh button or the from-GitHub PowerShell one-liner). No Render deploy.
