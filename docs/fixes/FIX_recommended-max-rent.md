# FIX: RECOMMENDED = fewest units at 40% + maximum rent

**Branch:** `fix/recommended-max-rent`
**Cut from:** `feature/excel-agent-foundation` @ `e86552f`
**Date:** 2026-06-14
**Risk:** Very low — one server-side selection rule. No VBA change (the add-in already reads `recommended_key`). No solver/compliance change.

## Client direction (2026-06-14)

Rachel's MIH rule, restated and locked: *"the least apartments at 40% AMI while maximizing rent."* So the headline RECOMMENDED should be **fewest units at 40%, then maximum rent at that count** — with larger units / lower floors emerging naturally from rent-max + the existing floor tiebreak. The tighter, lower-rent layouts still appear as additional options below.

## Change ([app.py](app.py))

The `recommended_key` computation previously chose, among minimum-40%-unit-count scenarios, the **tightest footprint** within a small income tolerance ("closest to 10%"). Per the clarified rule it now chooses the **maximum-rent** minimum-count scenario, tie-broken toward the tighter footprint when rents are equal:

```python
recommended_key, rec_s = max(rk_min, key=lambda kv: (round(income, 2), -sf40))
```

So: fewest apartments at 40% (guaranteed), highest rent at that count, and among equal-rent layouts the one that gives away the least 40% SF. Description prefix updated to "RECOMMENDED - fewest apartments at 40%, maximum rent."

The tighter `tight_40_footprint_*` options are unchanged — they remain as the "even lower footprint, lower rent" choices, now ordered below the recommended by income descending.

## 10% floor — confirmed (no change needed)

The solver enforces the MIH 40% floor as `low_band_sf >= ceil(0.10 x residential_sf)` ([ami_optix/solver.py:278](ami_optix/solver.py#L278)) — the requirement is rounded **up**, so every returned scenario is **>= 10.00%**. A 9.99999% mix mathematically fails the constraint and can never be returned (verified: on 143-24 the tightest feasible was 10.04%; exact 10.00% was infeasible, never below).

## Verified

- 143-24 94 Ave: recommended = 7 units @ 11.11%, **$35,755** (max rent at the 7-unit minimum); low_40_share 10.27% / tight 10.15% / tight 10.04% listed below by income. No scenario < 10%.
- Building 2 v4: recommended = 5 units @ 10.23%, **$32,507** (max rent at the 5-unit minimum).
- Both: recommended confirmed = min-count AND max income. Full suite **80 passed**.

## Deploy

Server-only → Render auto-deploys. **No `.xlam` refresh needed** (the add-in already consumes `recommended_key`).
