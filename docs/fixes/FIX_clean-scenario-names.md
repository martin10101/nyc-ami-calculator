# FIX: Client-facing scenario names (display only)

**Branch:** `fix/clean-scenario-names`
**Date:** 2026-06-14
**Risk:** Minimal — display-only. Only the two scenario-name formatter functions changed; no logic, ordering, solver, or compliance code touched.

## Why

Scenarios were displaying their raw internal keys (e.g. "EDGE WAAMI FLOOR 590", "TIGHT 40 FOOTPRINT 1", "BEST 3-BAND", and the recommended showing as "ALTERNATIVE"). A client should never see internal keys. Names now read like a person's option list.

## Change

Rewrote `FormatScenarioName` ([AMI_Optix_ResultsWriter.bas](excel-addin/src/AMI_Optix_ResultsWriter.bas)) and `FormatScenarioNameForPicker` ([AMI_Optix_Ribbon.bas](excel-addin/src/AMI_Optix_Ribbon.bas)) with a clean key->label map (and dynamic-suffix handling for the fewest/tight/edge families):

| Internal key | Display name |
|---|---|
| fewest_40_units* | FEWEST 40% UNITS |
| tight_40_footprint* | TIGHTER 40% FOOTPRINT |
| low_40_share | LOW 40% SHARE |
| mid_40_share | MID-RANGE 40% |
| max_40_share | MAX 40% SHARE |
| absolute_best / max_revenue | MAXIMUM RENT |
| edge_waami_floor_* | HIGHER RENT (MORE 40% UNITS) |
| edge_min/max_share_* | HIGHER RENT (RELAXED SHARE) |
| best_3_band / best_2_band | THREE-BAND / TWO-BAND MIX |
| closest_to_60 | CLOSEST TO 60% CAP |
| alternative | ALTERNATIVE MIX |
| original | YOUR ORIGINAL INPUT |
| best_rent_roll / client_oriented | BEST RENT ROLL / CLIENT ORIENTED |

The "(RECOMMENDED)" tag and the per-scenario "Why" line (which already explain the role) are unchanged. Picker labels mirror the sheet in Title Case.

## Verified

Static VBA checks clean (proc balance, no duplicate Dims, no leading-"=" cells) on both files. Diff confined to the two formatter functions. VBA-only.

## Deploy

VBA-only -> standard `.xlam` refresh.
