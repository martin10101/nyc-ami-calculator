# FIX: FEWEST group membership by actual unit count (not key name)

**Branch:** `fix/fewest-group-by-count`
**Cut from:** `feature/excel-agent-foundation` @ `47d7a85`
**Date:** 2026-06-11
**Risk:** Very low — one classifier change + one read-only helper in the shared ordering function. No server changes.

## Symptom (Building 1 v5 run)

LOW 40 SHARE (8 units @ 10.32%) appeared under MID RANGE while the FEWEST UNITS group held only the two `fewest_40_units*` options (8 @ 10.78%). Same minimum unit count, different group — the user read this as a grouping discrepancy, and he was right.

## Investigation findings (verified on the actual workbook, 2026 rents)

- **8 units is the PROVEN minimum** for this building: the 7 largest pool units sum to 5,494.72 SF = 9.50% of residential — below the 10% floor. Count 7 is mathematically impossible; the solver confirms INFEASIBLE.
- **The fewest search is exhaustive**: count-8 rent-max = $47,188 (8/8/8 at 40/60/90; share 10.78% = exactly the 8 largest units) at the production combo cap of 12, identical at caps 50 and 200.
- **10.78% is the fingerprint of the fewest-units strategy**, not over-shooting: big units at 40% = more WAAMI ballast = more 90s = more rent ($47,188 vs $47,129 for the SF-min 8-unit pick).
- The only defect was the grouping rule: G1 was key-name-based.

## Change ([excel-addin/src/AMI_Optix_ResultsWriter.bas](excel-addin/src/AMI_Optix_ResultsWriter.bas))

- New `ScenarioFortyUnitCount(scenario)` helper (mirrors `ScenarioFortyShare`; count needs no building SF, so it works for UAP too).
- `BuildGroupedScenarioOrder` pre-passes all solver scenarios (excluding `original`) to find the minimum positive 40%-unit count, then classifies:
  - `original` → YOUR INPUT
  - **40-unit count == minimum → FEWEST UNITS AT 40%** (now catches `low_40_share` and any other same-minimum option; `fewest_40_units*` keys keep a name-based safety net when counts are unknowable)
  - else share ≤ 11.5% → MID RANGE, else MAX RENT (unchanged)

Overview table, detail blocks, and the View Scenario picker all follow automatically (single ordering function).

Expected B1 v5 regrouping: FEWEST = FEWEST 40 UNITS ($47,188), FEWEST 40 UNITS 2 ($47,125), LOW 40 SHARE ($47,129) — all 8 units; MID = CLOSEST TO 60 (9 @ 11.22%); MAX = ABSOLUTE BEST / ALTERNATIVE / MAX 40 SHARE (9 units, 11.67-12.05%).

## Verification

- Static VBA checks: 65 procs/65 ends, no duplicate Dims, no leading-"=" cell strings.
- pytest: 78 passed (no Python changes).
- Manual after `.xlam` refresh: rerun B1 v5 — three 8-unit options under FEWEST; numbering consistent across overview/blocks/picker.

## Deploy

VBA-only → standard `.xlam` refresh.
