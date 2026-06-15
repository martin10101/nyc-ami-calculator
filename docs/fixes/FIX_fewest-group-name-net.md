# FIX: FEWEST group name-net only fires when the unit count is unknowable

**Branch:** `fix/fewest-group-name-net`
**Cut from:** `feature/excel-agent-foundation` @ `c48e157`
**Date:** 2026-06-15
**Risk:** Very low — adds one guard (`cnt <= 0`) to one `ElseIf` in the shared display-ordering function. Display/grouping only. No solver, no API, no Ctrl+Z, no manual-block/MIH-sync code touched. No server changes.

## Symptom (43-09 52nd Street, LLC — 2026)

In the Scenario Overview, scenario 6 `fewest_40_units_3` (**4 @ 12.03%**, $14,439) was listed under **FEWEST UNITS AT 40%**, even though the proven minimum for the run is **3** units (scenarios 1–5 are all 3-unit options). A 4-unit option does not have the fewest units, so it belongs under **MAX RENT / OTHER OPTIONS**. The user flagged it as a grouping mix-up, and he was right.

## Root cause ([excel-addin/src/AMI_Optix_ResultsWriter.bas](excel-addin/src/AMI_Optix_ResultsWriter.bas), `BuildGroupedScenarioOrder`)

[FIX_fewest-group-by-count.md](FIX_fewest-group-by-count.md) (2026-06-11) made group membership count-based, keeping "a name-based safety net **when counts are unknowable**" for `fewest_40_units*` keys. But the code never checked that the count was actually unknowable:

```vba
ElseIf minFortyCount > 0 And cnt = minFortyCount Then
    g = 0                                                  ' at the true minimum -> FEWEST (correct)
ElseIf LCase$(Left$(ks, Len("fewest_40_units"))) = "fewest_40_units" Then
    g = 0   ' name-based safety net when counts are unknowable   <-- fired even when cnt was known
```

So any key starting with `fewest_40_units` was forced into FEWEST by name alone. `fewest_40_units_3` has a *known* count of 4 (> the minimum of 3), yet the name net captured it before the share rule could place it.

## Change

Guard the name net with `cnt <= 0` so it only applies when the count genuinely cannot be read:

```vba
ElseIf cnt <= 0 And LCase$(Left$(ks, Len("fewest_40_units"))) = "fewest_40_units" Then
    g = 0   ' name-based safety net ONLY when the 40% unit count is unknowable
```

`ScenarioFortyUnitCount` already returns `-1` when the count can't be computed, so the safety net still covers that case. A `fewest_40_units_*` option with a *known* count above the minimum now falls through to the existing share rule (≤ 11.5% → MID, else MAX / OTHER).

## Effect (traced against the 8-scenario run above)

`minFortyCount = 3`. Only scenario 6 moves; the other 7 stay put.

| # | Key | 40% count | Before | After |
|---|-----|-----------|--------|-------|
| 1 | fewest_40_units (recommended) | 3 = min | FEWEST | FEWEST |
| 2 | fewest_40_units_2 | 3 = min | FEWEST | FEWEST |
| 3 | absolute_best | 3 = min | FEWEST | FEWEST |
| 4 | closest_to_60 | 3 = min | FEWEST | FEWEST |
| 5 | best_3_band | 3 = min | FEWEST | FEWEST |
| **6** | **fewest_40_units_3** | **4 > min** | **FEWEST** | **MAX RENT / OTHER** (12.03% > 11.5%) |
| 7 | max_40_share | 4 > min | MAX RENT / OTHER | MAX RENT / OTHER |
| 8 | original | n/a | YOUR INPUT | YOUR INPUT |

No scenario's units, rents, bands, recommended pick, or numbering source changes — only the section heading scenario 6 is listed under. Overview table, detail blocks, and the View Scenario picker all follow automatically (single ordering function).

## Verification

- Logic trace above: only `fewest_40_units_3` regroups; FEWEST still holds every true-minimum option; unknowable-count `fewest_40_units*` options still land in FEWEST via the guarded net.
- Manual after `.xlam` refresh: rerun the 43-09 52nd Street run — scenario 6 appears under MAX RENT / OTHER OPTIONS; recommended scenario 1 and the Scenario Manual / MIH page are unchanged.

## Deploy

VBA-only (one `.bas`) → standard `.xlam` refresh.
