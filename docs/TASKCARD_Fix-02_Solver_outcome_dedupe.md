# Task Card — Fix-02 Solver dedupe identical outcomes + placement tie-break (Python)

## Goal
- Remove “duplicate-looking” solver scenarios from the `/api/optimize` response when they are outcome-identical (same band mix + same rent totals), even if unit-level assignments differ.
- Keep only one scenario per equivalence group.
- Choose the kept scenario using placement tie-break (40% units lower floors; higher AMI / higher-paying units higher floors).

## Success Criteria (from `docs/FIX_REQUIREMENTS.md`)
- Only dedupe when outcomes are truly identical:
  - same AMI band mix
  - same rent outcome metrics used for client results
  - same relevant mix/sqft outcome if part of outputs
- Keep only one scenario per equivalence group.
- Choose the kept scenario using placement tie-break:
  - 40% units prefer lower floors
  - higher AMI / higher paying units prefer higher floors
- Must be surgical: do not change solver’s main scoring/constraints; apply as post-processing right before returning results.

## Files To Change (target ≤ 3 total files for this fix)
- `app.py`
- `docs/CODEX_LEDGER.md` (after PR is opened)

## Functions / Entrypoints To Change
- `optimize_units()` (`/api/optimize`)

## Proposed Patch (logic-level)
- After strict scenarios + edge/top-up scenarios are fully assembled **and** rent totals are computed, dedupe scenarios by an “outcome signature”:
  - normalized `metrics.band_mix` (band, units, net_sf)
  - `rent_totals.net_monthly` + `rent_totals.net_annual`
  - total SF / WAAMI (rounded) as an additional guard
- For each equivalence group, pick the winner by “placement” score:
  - reward higher `floor * assigned_ami` alignment
  - reward higher `floor * monthly_rent` alignment (when present)
  - penalize high floors for 40% band units
  - deterministic tie-breaker (stable key ordering)
- Keep the highest-priority scenario key for the group (`absolute_best` first, then `best_3_band`, `best_2_band`, `alternative`, etc.), but store the winning scenario object under that key and remove the others.
- Be conservative: if required signature fields are missing (e.g., no rent schedule loaded → no `rent_totals`), skip dedupe.

## Risk Pre‑Mortem
- **Risk:** Over-dedupe due to float noise / rounding.
  - **Mitigation:** Use rent totals (already rounded to 2 decimals) + band mix + total SF/WAAMI guard; only dedupe when all fields are present.
- **Risk:** Returning fewer scenarios than expected (e.g., fewer than ~6).
  - **Mitigation:** This is acceptable if duplicates were previously inflating count; record counts in edge test and keep top-up logic unchanged.
- **Risk:** Accidentally removing `absolute_best` key.
  - **Mitigation:** Always keep `absolute_best` as the “keeper key” when it’s part of a duplicate group.

## Test Plan
### Automated
- `python -m pytest -q`

### Edge check (count scenarios returned)
- Run a local request via Flask test client and print:
  - number of scenario keys returned
  - which keys remain after dedupe
  - confirm `absolute_best` is present

## Rollback Plan
- Revert the Fix-02 commit(s) on the branch and re-deploy the prior ready commit SHA from the ledger (Render auto-deploy is OFF).

