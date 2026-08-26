# FIX: Remove the 3% haircut at 100% AMI (show true published rent)

**Branch:** `fix/remove-100-ami-haircut`
**Cut from:** `feature/excel-agent-foundation` @ `0335f4e`
**Date:** 2026-08-05
**Risk:** Low, server-only, one constant. Solver code untouched. No .xlam change, no PC updates.

## Background

From 2026-06-04 ([FIX_rent-first-objective-with-100pct-haircut.md](FIX_rent-first-objective-with-100pct-haircut.md)) the program charged 97% of the published rent at exactly the 100% AMI band, applied at the source in `rent_components` and disclosed on every results sheet ("3% Cap Applied" + Pre-Cap column). On 2026-08-05 the client directed that results show the **true published rent** at 100% AMI with no reduction. Removed per that instruction.

## Change ([ami_optix/rent_calculator.py](../../ami_optix/rent_calculator.py))

`HAIRCUT_FACTOR: 0.97 → 1.0`. The mechanism is intentionally kept as a one-line switch (with a comment recording the removal) in case a regulatory cap must ever be re-applied. Nothing else touched:

- `haircut_applied` computes to `False` for every unit, so the Excel add-in's "3% Cap Applied" / Pre-Cap display naturally shows no cap — zero VBA changes needed.
- All API response fields (`gross_pre_haircut`, `haircut_applied`) remain present, so no client parsing can break.
- The solver was not modified; its rent coefficients simply receive the true value from `rent_components` like every other consumer.

## Expected visible effects

- 100% AMI rents now equal the published table value (e.g. 2026 1BR: **$3,181**, was $3,085.57). Any total containing 100% units rises ~3% of those units' rent.
- Because 100% AMI units are now worth 3% more to the rent objective, the optimizer may legitimately shift some mixes toward the 100% band. Same solver, richer 100% input — expected, not a bug.
- Results sheets show "(no 100% AMI units; cap not applied)" style text on the cap line and an empty Pre-Cap column. Cosmetic; removable later in the batched Excel display cleanup.

## Verification

- Tests: the three June haircut tests rewritten to assert true rents (`tests/test_rent_calculator.py`); full suite **87 passed**.
- All-years check: 2024 / 2025 / 2026 / default calculators, all bedroom sizes — charged == published at 100% AMI, `haircut_applied=False`, all PASS (year tables are data; the factor is year-independent by construction).
- API end-to-end: `/api/evaluate` returns gross 3,181 (2026 1BR @ 100%), pre_cap 3,181, haircut False.

## Deploy

Server-only → push → Render auto-deploys.
