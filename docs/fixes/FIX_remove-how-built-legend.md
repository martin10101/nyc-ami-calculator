# FIX: Remove "HOW THESE OPTIONS ARE BUILT" legend from results sheet

**Branch:** `fix/remove-how-built-legend`
**Date:** 2026-09-02
**Risk:** Low. VBA-only (`AMI_Optix_ResultsWriter.bas`), display text removal. No server change, no ribbon change, no new buttons.

## Symptom (client, Building D feedback)

"Please remove the paragraph of how these options were built." The 5-row
methodology legend rendered under the SCENARIO OVERVIEW on the client-facing
AMI Scenarios sheet.

## Change

1. `WriteScenarioOverview` (was ~:904-925): legend header + 4 explanation lines
   deleted; one spacer row retained so the block below keeps its gap.
2. `UpdateOverviewRentColumn` (:3524): the overview scan used
   `Left$(c1,4) = "HOW "` as its terminator — with the legend gone that stop
   condition would never fire and the scan would run into the SF/utilities
   blocks (harmless today only thanks to the IsNumeric guard). Terminator now
   stops at `"SQUARE"` (SF summary) or `"UTILITIES"` in addition to the
   legacy `"HOW "` (kept for sheets written by OLDER add-in versions) and
   `"SCENARIO "`.

## Couplings checked (full-source audit, 15 modules + AppEvents.cls)

- `"HOW "` was read back at exactly ONE site (:3524) — updated in this fix.
- The 4 legend lines are never read back anywhere.
- No absolute row anchors below the legend; all callers consume
  `WriteScenarioOverview`'s returned row cursor, so content shifts up cleanly.
- Ribbon XML untouched; no callbacks reference the legend.

## Verification

- Grep: zero remaining references to legend text in excel-addin/.
- Full pytest suite: 87 passed (server untouched — baseline identity).
- Sandbox Excel QA before any client PC: run UAP + MIH, re-run, Manual
  Calculate, Apply, Ctrl+Z, live-sync edit (per tests/manual_qa_plan.md).

## Deploy

VBA module swap per PC (Deploy-AmiOptixFixes.ps1 pipeline), staged:
sandbox -> secondary PC -> Rachel. Rollback: re-pin previous commit.
