# FIX: Minimum-count 40% frontier + RECOMMENDED option

**Branch:** `fix/min-count-40-frontier`
**Cut from:** `feature/excel-agent-foundation` @ `8544122`
**Date:** 2026-06-12
**Risk:** Medium — new server generation + WAAMI-cap float tolerance + VBA ordering. MIH-only; UAP untouched. Suite 80 green.

## Client direction (2026-06-12)

At the minimum number of apartments at 40%, surface **as many distinct options as exist** from the 10% floor up to the rent-strong layout — not just one. Among them, **name the RECOMMENDED** = the option that hugs the 10% floor most tightly while still earning essentially the best income ("closest to 10% with the best available," Rachel's way). The tighter-but-lower-rent layouts must still appear but be **demoted and clearly marked** with their exact rent gap. Must work across all building sizes; UAP unchanged.

## What it does

Building 143-24 94 Ave (2026), 7-apartment minimum:

```
FEWEST UNITS AT 40%
 ★1 LOW 40 SHARE (RECOMMENDED)   40/60/90   10.27%   $35,751
  2 ALTERNATIVE                  40/60/80   11.11%   $35,755
  3 TIGHT 40 FOOTPRINT 2         40/60/80   10.15%   $35,626   -$125/mo vs recommended; not recommended
  4 TIGHT 40 FOOTPRINT 1         40/70/80   10.04%   $35,496   -$255/mo vs recommended; not recommended
MAX RENT / OTHER OPTIONS …  (8-apartment options)
```

Manual block = the RECOMMENDED. Order = recommended first, then income descending (so the tighter footprint-minimizing options sink and are marked).

## Changes

### Server ([app.py](app.py)) — MIH-only

- **Min-count 40% frontier generator**: at the proven minimum unit count, steps the 40% share ceiling finely from just above the floor up to the tightest already-offered option's share, rent-maximizing at each, collecting every distinct **tighter** layout. Tags them `tier='tight_footprint'`, keys `tight_40_footprint_N`. Combo-check budget raised to 36 (12 missed the tightest layout).
- **RECOMMENDED computation**: among minimum-count scenarios, the tightest 40% footprint whose income is within a small tolerance (max($25, 0.12%·income)) of the best at that count. Exposed as `response['recommended_key']`.
- **Descriptions**: tight_footprint options get a "Tighter 40% footprint … −$X/mo vs the recommended option; not the recommended option" Why line; the recommended gets a "RECOMMENDED — …" prefix.
- **WAAMI-cap float tolerance (1e-9)**: a layout sitting exactly at the 60.00% cap recomputes to 60.0000000000001% in float (≈1e-16 representation noise) and was falsely dropped by the strict `> cap` checks — so the absolute-tightest option (which often lands right at the cap) vanished. Added a 1e-9 tolerance to the frontier check and the final hard-cap pass. Not a real-violation allowance: 1e-9 ≪ any share the client cares about; verified the dropped layout's exact (integer-method) WAAMI is 0.600000 (over by 0).

### VBA ([excel-addin/src/AMI_Optix_ResultsWriter.bas](excel-addin/src/AMI_Optix_ResultsWriter.bas))

- New `g_AMIOptixRecommendedKey` (set from `result("recommended_key")`).
- FEWEST group ordering changed from tightest-first to **recommended-first, then income descending** — the tight_footprint options naturally sink below the rent-strong choices.
- `GetBestScenarioKey` (manual block) prefers the recommended key; overview row gets a bold "(RECOMMENDED)" tag.
- New `ScenarioNetMonthly` helper.

## Verified

- 143-24 94 Ave: full frontier (10.04 / 10.15 / 10.27 / 11.11%), recommended = 10.27%, order + manual block correct.
- Broad sweep (Building D, B1 v4, B2 v4, B3 v5, 3320 Atlantic): recommended set to a min-count scenario every time, no WAAMI over cap, all descriptions present, tighter options generated where they exist.
- UAP (Test, 1530 Bergen, 212 West 231): `recommended_key` is None, no `tight_40_footprint` keys — completely unchanged.
- Tests: 2 new in [tests/test_low40_options.py](tests/test_low40_options.py); static VBA checks clean; full suite **80 passed**.

## Deploy

Server (Render auto) + VBA → standard `.xlam` refresh.
