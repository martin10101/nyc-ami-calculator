# FIX: MIH — Remove duplicate AMI band rows from Square Footage Summary

**Branch:** `fix/mih-remove-duplicate-band-summary`
**Cut from:** `feature/excel-agent-foundation` @ `6cab6e5`
**Date:** 2026-04-29
**Author:** client request via remote session
**Risk:** Very low (display-only change)

## Status

| Step | When | Result |
|---|---|---|
| Committed locally | 2026-04-29 — sha `1e9acec` | ✅ |
| Pushed `fix/mih-remove-duplicate-band-summary` | _pending_ | ⏳ |
| Fast-forward merged into `feature/excel-agent-foundation` | _pending_ | ⏳ |
| Pushed `feature/excel-agent-foundation` (triggers Render auto-deploy) | _pending_ | ⏳ |
| Render deploy live | _pending_ | ⏳ |
| Client PC refreshed via PS agent | _pending — user runs the one-command_ | ⏳ |
| Manual test passed on client PC | _pending — user follows the test checklist below_ | ⏳ |
| Approved by client | _pending_ | ⏳ |

This table is updated as the deploy chain progresses. Note: the change is
VBA-only — Render serves the Python backend, which is unchanged. The
Render deploy step is for keeping `feature/excel-agent-foundation`
in sync as the source of truth, not because Render's runtime behavior
changes.

## Problem

On the MIH results sheet, the **Square Footage Summary** block at the top
duplicated the per-AMI-band breakdown (40 / 60 / 80 % rows) that already
appears in the **Scenario Manual → Band Mix (by Net SF)** table below.

The client asked to remove the duplication while keeping the
building-level totals (Total Building Net SF, Affordable Net SF) at the
top of the page.

UAP is unaffected — there is no equivalent rendering function in UAP code
(grep confirmed: no `WriteUapSquareFootageSummary`).

## Visual diff

**Before:**
```
SQUARE FOOTAGE SUMMARY
Total Building Net SF:    23,887.43
Affordable Net SF:         6,025.30
AMI Band   Net SF       % of Total Building SF
40% AMI    2,145.84      8.98%   [yellow highlight]
60% AMI    1,733.62      7.26%
80% AMI    2,145.84      8.98%
                                                <-- blank
UTILITIES — Selected Variants ...
```

**After:**
```
SQUARE FOOTAGE SUMMARY
Total Building Net SF:    23,887.43
Affordable Net SF:         6,025.30
                                                <-- blank
UTILITIES — Selected Variants ...
```

The 40 / 60 / 80 breakdown still appears unchanged in the
"Band Mix (by Net SF)" table inside the Scenario Manual block lower on
the same sheet. No data is lost.

## Code change

**File:** `excel-addin/src/AMI_Optix_ResultsWriter.bas`
**Function:** `WriteMihSquareFootageSummary` (originally lines 2266-2382, now lines 2266-2333)

| Section | Action | Lines (old numbering) |
|---|---|---|
| Function signature, decl `Dim row…buildingSf` | kept | 2266-2273 |
| Decls `Dim bandVal / netSf / bandLabel` | **removed** (locals only used by deleted loop) | 2276-2278 |
| Decl `Dim idx` and `Dim bm` | kept (still used by `totalSf` accumulator loop) | 2274-2275 |
| Function comment | rewritten to reflect new behavior; explicitly notes per-band lives in Band Mix below to prevent re-introduction | 2267-2270 |
| Early returns + `totalSf` computation + `buildingSf` assignment | kept | 2280-2306 |
| `SQUARE FOOTAGE SUMMARY` section header + Total Building row + Affordable row | kept | 2308-2329 |
| `' Column headers` block (`AMI Band` / `Net SF` / `% of Total Building SF`) | **removed** | was 2334-2340 |
| `' Per-band rows` loop (writes 40 / 60 / 80 % rows + yellow highlight on 40%) | **removed** | was 2342-2378 |
| Trailing `row = row + 1` blank-row spacer + return | kept | 2380-2382 |

**Net change:** function shrinks from 117 lines to 68 lines. ~49 lines deleted, 0 functional lines added (just the corrected comment).

The function is `Private` and is called from two manual-block writers:
- `WriteManualScenarioBlockFromResult` (line 2232 in original, after Run MIH)
- `WriteManualScenarioBlockFromEvaluate` (line 2440 in original, after Manual Calculate)

Editing the function once fixes both rendering paths.

No other files touched. No solver, no API, no rent calc, no event handlers.

## Manual test (required before merging back to `feature/excel-agent-foundation`)

1. Push this branch to GitHub.
2. On the client PC: run
   `powershell -ExecutionPolicy Bypass -File C:\AMI_Optix_Agent\scripts\Refresh-AmiOptixAgentFromGitHub.ps1 -AgentRoot C:\AMI_Optix_Agent`
   to pull the updated `.bas`, then `Run-AmiOptixAutofix.ps1` to rebuild
   the staged `.xlam`.
3. Reopen Excel; load `230 Kent_Unit Schedule - MIH v3.xlsm` (the
   workbook in the screenshot that surfaced this issue).
4. Click **Run MIH**.
5. ✅ Top of the AMI OPTIMIZATION RESULTS sheet shows only:
   `SQUARE FOOTAGE SUMMARY` → `Total Building Net SF` → `Affordable Net SF` → blank → `UTILITIES`.
6. ✅ The `AMI Band / Net SF / % of Total Building SF` header row and
   the per-band 40 / 60 / 80 % rows are **gone**.
7. ✅ Lower on the same sheet, the `SCENARIO MANUAL` block's
   `Band Mix (by Net SF)` table still shows full 40 / 60 / 80 %
   breakdown with `Units`, `Net SF`, `Share of SF`, `Share of Building SF`
   columns intact.
8. Tweak a unit's AMI on the input sheet → click **Manual Calculate**.
   Confirm same layout (top stays clean, manual block recalcs).
9. Open a UAP workbook → click **Run UAP**. ✅ UAP layout is **identical**
   to before this fix.

## Rollback

```bash
git checkout feature/excel-agent-foundation
git revert <commit-sha-on-fix-branch>   # creates a new commit that restores the deleted block
# or, if the branch hasn't been merged yet, just abandon it:
git branch -D fix/mih-remove-duplicate-band-summary
```

Then re-run the PS agent on the client PC to regenerate the staged
`.xlam` from the (now-reverted) source.

## Deploy notes

- Render auto-deploy is **ON** (verified 2026-04-28). Pushing this branch
  will trigger a Render deploy. The change is VBA-only — Render serves
  no VBA — so the deploy adds nothing functional. But the deploy does
  happen and will replace the live revision on
  https://nyc-ami-calculator.onrender.com.
- Client-PC rollout is via PowerShell agent only; no manual VBA copy /
  paste in the Excel VBA editor.
