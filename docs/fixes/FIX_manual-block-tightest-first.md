# FIX: Manual block = Scenario 1, FEWEST group tightest-hug first

**Branch:** `fix/manual-block-tightest-first`
**Cut from:** `feature/excel-agent-foundation` @ `1e7e5a7`
**Date:** 2026-06-12
**Risk:** Low — VBA display/ordering only. No server, solver, or compliance changes.

## Symptom (143-24 94 Ave run on the new PC)

1. The manual/current working-copy block showed **ABSOLUTE BEST (8 units @ 12.19%)** as the headline, instead of the least-units option.
2. The manual block did not match **Scenario 1** in the overview/detail blocks (which was the 7-unit ALTERNATIVE).

This is the exact concept Rachel flagged: owners want the *least units* at 40% as the headline, not the more-units-for-more-income option.

## Root cause

`GetBestScenarioKey` (which the manual block uses) picked from a hard-coded name list — `fewest_40_units` first, then `absolute_best`. This building never produced a literal `fewest_40_units` key (its minimum-count options came back named `low_40_share` / `alternative`), so the list fell through to `absolute_best` — the 12% option. Meanwhile the overview/detail blocks group by *actual unit count*, so they correctly led with the 7-unit option. Two different orderings → they disagreed.

## Fix ([excel-addin/src/AMI_Optix_ResultsWriter.bas](excel-addin/src/AMI_Optix_ResultsWriter.bas))

1. **`GetBestScenarioKey` now returns Scenario 1 of the grouped order** (`BuildGroupedScenarioOrder`) — the manual block and the first numbered scenario come from one source and can never disagree. Falls back to the old name list only if grouped order is unavailable.
2. **FEWEST group ordered tightest-hug first** — within the minimum-40%-unit-count group, scenarios are sorted by ascending 40% SF share (stable insertion sort; equal/unknown shares keep base order). So Scenario 1 is the least-units option that hugs the 10% floor most tightly.

## Verified (143-24 94 Ave, 2026)

New order: **1. LOW 40 SHARE 7u @ 10.27% $35,751** → 2. ALTERNATIVE 7u @ 11.11% → 3+. MAX RENT 8u @ 12.19%. Manual block now shows `low_40_share` = Scenario 1. Min count confirmed = 7 (6 largest units = 9.89% < 10% floor → infeasible at 6).

Static VBA checks clean (65/65 procs, no dup Dims, For/Next + Do/Loop balanced). pytest 78 green (no Python changes).

## Deploy

VBA-only → standard `.xlam` refresh.

## Open follow-up (not in this fix)

Investigation found a true tightest-footprint layout at this building of **10.04%** (3,142 SF, 12 SF over the floor) that the program does not currently surface — it earns $255/mo less than the 10.27% option. Whether to add it as an additional option (and whether it should lead, since it is tighter but lower-rent) is a separate decision pending client direction.
