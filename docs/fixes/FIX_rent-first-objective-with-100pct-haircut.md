# FIX: Rent-first objective + 100% AMI haircut

**Branch:** `fix/rent-first-objective-with-100pct-haircut`
**Cut from:** `feature/excel-agent-foundation` @ `cacecd8`
**Date:** 2026-06-04
**Author:** client gap analysis (see plan file)
**Risk:** Medium (changes which scenario is "absolute_best"; existing scenarios still accessible)

## Status

| Step | When | Result |
|---|---|---|
| Client gap analysis report | 2026-06-04 | ✅ See plan file |
| Baseline pytest (60 tests) | 2026-06-04 | ✅ |
| Apply haircut in rent_components | 2026-06-04 — sha _pending_ | ⏳ |
| Re-rank scenarios in app.py | 2026-06-04 — sha _pending_ | ⏳ |
| New haircut regression tests (3) | 2026-06-04 | ✅ |
| Local pytest (63 tests post-fix) | 2026-06-04 | ✅ all passed |
| Commit + FIX doc | 2026-06-04 — sha _pending_ | ⏳ |
| Pushed fix branch | _pending_ | ⏳ |
| Fast-forward merge to feature | _pending_ | ⏳ |
| Pushed feature branch (Render auto-deploy) | _pending_ | ⏳ |
| Approved by client | _pending_ | ⏳ |

## One-line PowerShell command for client-PC update

```
powershell -ExecutionPolicy Bypass -File C:\AMI_Optix_Agent\scripts\Refresh-AmiOptixAgentFromGitHub.ps1 -AgentRoot C:\AMI_Optix_Agent
```

**Note:** No VBA changes in this fix. The PS refresh is a no-op for the add-in (.xlam unchanged). The fix takes effect server-side once Render finishes auto-deploying — typically 1-3 minutes after the push. The client can test by re-running "Find Optimal Scenarios" on any MIH workbook; no Excel restart needed.

## Problem

Client gap analysis (2026-06-04, see plan file) showed our solver maximizes WAAMI under the 60% cap, but the client actually maximizes **rent dollars** subject to compliance. Two specific gaps:

1. **The 3% regulatory haircut at 100% AMI was never applied** in our code. Our solver modeled a 100% AMI 1BR as generating $3,180/mo when the regulatory cap means the landlord can only collect $3,180 × 0.97 = $3,084.60. That's ~$95/unit/month of phantom revenue — for Building 1 v4 with 6×100% units, ~$6.8K/year that doesn't exist.

2. **The "absolute_best" scenario was the WAAMI-max winner**, not the rent-max winner. Building 1 client example proves the gap:
   - **v4** (our solver would pick): 6×100% + 10×60%, WAAMI 59.65%, monthly rent $41,037
   - **v5** (client picked): 8×90% + 8×60%, WAAMI 59.79%, monthly rent $41,039

   The client traded 100% units for 90% units even though it moved WAAMI *further* from 60%, because she's optimizing for rent dollars and a clean rent roll. Our solver couldn't pick v5 because the math literally preferred v4.

## Solution

Two coordinated server-side changes, both apply to UAP and MIH equally. No VBA change. No solver.py change (no new constraint variables or CP-SAT model changes).

### Change 1: Apply 3% haircut at the source in `ami_optix/rent_calculator.py`

Added module-level constants at the top:

```python
HAIRCUT_BAND_AMI = 1.0
HAIRCUT_FACTOR = 0.97
```

Modified `RentSchedule.rent_components` to apply the haircut when `ami_percent == 1.0`:

```python
gross_pre_haircut = float(self._gross_rents_lookup(ami_percent, bedroom_label))
if abs(ami_percent - HAIRCUT_BAND_AMI) < 1e-6:
    gross = gross_pre_haircut * HAIRCUT_FACTOR
else:
    gross = gross_pre_haircut
```

Return dict now includes both `gross` (haircut-adjusted) and `gross_pre_haircut` (headline) plus a `haircut_applied` boolean for transparency. Backward-compatible: existing keys (`gross`, `net`, `allowances`, `allowance_total`) preserved.

**Why this is the single point of change:** every rent computation in the codebase flows through `rent_components`. Specifically:
- `rent_by_band_cents` builder at `app.py:1025-1031` calls `rent_components` per band → haircut applied
- `_apply_rents_to_scenarios` calls `compute_rents_for_assignments` → which calls `rent_components` → haircut applied
- `RentSchedule.rent_for` calls `rent_components` → haircut applied
- Edge scenarios via `find_max_revenue_scenario` use `rent_by_band_cents` → haircut applied

One change point, universal effect.

### Change 2: Re-rank scenarios by actual rent in `app.py`

After `_apply_rents_to_scenarios(scenarios)` populates `rent_totals` on every scenario (around line 991), added a re-ranking step:

```python
original_best = scenarios.get('absolute_best')
if original_best and original_best.get('rent_totals'):
    rented = [(k, v) for k, v in scenarios.items()
              if v and isinstance(v.get('rent_totals'), dict)
                 and v['rent_totals'].get('net_monthly') is not None]
    if rented:
        best_key, best_scenario = max(rented, key=lambda kv: float(kv[1]['rent_totals']['net_monthly']))
        if best_scenario is not original_best:
            scenarios['closest_to_60'] = original_best
            scenarios['absolute_best'] = best_scenario
            notes.append(f"Recommended scenario switched from WAAMI-max to rent-max ...")
```

**Why this works:** `_apply_rents_to_scenarios` runs `compute_rents_for_assignments` (which uses the now-haircut-adjusted `rent_components`) on every scenario. After that, `rent_totals.net_monthly` reflects actual collectable rent (with haircut). Re-ranking is then a simple max over scenarios.

Note: the solver's internal `revenue_score` field is `sum(net_sf × assigned_ami)` — that's a WAAMI proxy, not actual rent. So sorting on `revenue_score` inside the solver was NOT a true rent-max. The fix correctly uses actual `net_monthly` rent dollars instead.

### Why we don't add an "avoid 100%" tie-breaker

Earlier draft proposed an explicit penalty for 100% AMI units. Removed after client clarification: she does sometimes use 100% (when the rent math justifies it). With the haircut applied, the rent math itself will:
- Prefer 90% in cases like Building 1 (where the rent gap is small enough that 0.97× of 100% < 1.0× of 90% × N units of a different mix)
- Still prefer 100% in cases where it genuinely beats 90% on net rent

So the haircut handles the 100% preference organically — no hard-coded rule needed.

## What this does NOT change

- VBA / `.xlam` add-in
- Rent calculator XLSX files (haircut lives in Python, not the spreadsheet)
- 40% AMI window, band-count cap, WAAMI cap, share thresholds — all hard compliance rules unchanged
- The solver's CP-SAT model (no new variables, no new constraints)
- UAP-specific `deep_affordability_min_share` / `_max_share` rules
- Existing test assertions about WAAMI calculation, share thresholds, band counts
- Edge scenarios (`max_revenue`, `edge_max_share_*`) — these continue to use the rent-max code path they already did

## Verification

### Local tests (all green)

```
tests/test_rent_calculator.py::test_rent_allowances_reflect_utility_selection PASSED
tests/test_rent_calculator.py::test_100_pct_ami_rent_has_haircut_applied PASSED  ← NEW
tests/test_rent_calculator.py::test_no_haircut_at_bands_other_than_100 PASSED    ← NEW
tests/test_rent_calculator.py::test_compute_rents_for_assignments_reflects_haircut PASSED  ← NEW

Full suite: 63 passed (60 baseline + 3 new haircut tests).
```

### Post-deploy verification on real Building 1 workbook

1. Server check: `curl https://<render-url>/api/healthz` returns 200.
2. Open `Unit Schedule_Building 1, v4.xlsb` (the workbook where v4 is currently saved).
3. Run "Find Optimal Scenarios" with 2025 or 2026 rent calculator.
4. The `absolute_best` scenario in the API response should now have **0 units at 100% AMI** (not 6).
5. The `closest_to_60` scenario in the response should preserve the old behavior (6×100%, WAAMI 59.65%).
6. The note "Recommended scenario switched from WAAMI-max to rent-max..." should appear in the `notes` array.
7. In Excel, the `absolute_best` scenario should display ~$41,039/mo net rent (matching v5).

If steps 4-6 work, the fix is live and correct.

## Risks / open questions

- **Edge cases where `absolute_best` and the rent-max winner are the same scenario** — re-ranking is a no-op, no `closest_to_60` added. This is fine and expected.
- **Existing customers' historical scenarios may shift** — if someone re-runs an old Building 1 calculation, they'll get v5-style results instead of v4. This is intended (matches client preference), but documented here so we know to communicate it.
- **Haircut might not be exactly 3% in all programs/years** — hard-coded at 0.97 with a clear constant. If client comes back with "actually it's 4% under program X," we add a `HAIRCUT_FACTOR_OVERRIDE` env var or per-program config.
- **No UAP regression test yet** — Building 1-3 are MIH workbooks. A UAP regression test should follow once the client provides a UAP workbook with a known "right answer."
