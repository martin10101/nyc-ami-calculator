# FIX: Push rent-max into per-combo allocation (closes residual gap)

**Branch:** `fix/rent-max-per-combo`
**Cut from:** `feature/excel-agent-foundation` @ `95896ed`
**Date:** 2026-06-04
**Author:** end-to-end testing against B1 v5 workbook surfaced residual gap
**Risk:** Medium (changes solver's per-combo objective when rent data is available; backward-compat for tests/callers without rent data)

## Problem

After the previous fix (`FIX_rent-first-objective-with-100pct-haircut.md`), we re-ranked scenarios by actual rent at the API layer and applied the 100% haircut at the source. End-to-end testing against `Unit Schedule_Building 1, v5 UPDATED AMI.xlsb` with the 2026 rent calc and the client's actual tenant-paid utilities revealed a **residual $895/mo gap**:

| | Bands | Allocation | Net monthly (2026 + client utils) |
|---|---|---|---|
| Client's v5 (as saved) | 40/60/90 | 8 / 8 / 8 | $43,549.00 |
| Previous fix's `absolute_best` | 40/70/90 | 10 / 9 / 5 | $42,654.00 |
| **Gap** | — | — | **$895.00/mo** |

Root cause: the solver's per-combo allocation step (`_solve_single_scenario`) was optimizing `revenue_score = sum(net_sf × assigned_ami)` — a WAAMI **proxy**, not actual rent dollars. So within each band combo, the solver picked the WAAMI-max unit-to-band assignment. The client's hand-crafted v5 allocation was rent-better but WAAMI-worse, so the solver couldn't find it.

## Solution

Push rent-maximization down into the per-combo solve. Pass `rent_by_band_cents` (pre-computed using haircut-adjusted rent_components) into `find_optimal_scenarios`. Inside the per-combo loop, when rent coefficients are available, call `_solve_single_scenario` with `objective_mode="rent"` and the rent coefficients — same code path that `find_max_revenue_scenario` already uses for edge scenarios. Each combo's allocation now maximizes actual rent dollars (in cents, as integers for CP-SAT).

Backward-compatible: when `rent_by_band_cents` is None (no rent schedule available, or unit tests without rent data), the solver falls back to the existing WAAMI-max behavior.

### Changes

**`ami_optix/solver.py`** — `find_optimal_scenarios` accepts new optional `rent_by_band_cents` parameter. When provided and the band combo has rent coefficients for all bands, the per-combo solve uses rent objective. When not, behaves as before.

**`app.py`** — load `rent_schedule` BEFORE the solver call (moved up from line 962 to before line 821). Build `rent_by_band_cents` using `rent_components` (which applies the 100% haircut). Pass to all three `find_optimal_scenarios` call sites: MIH floor-walk loop, UAP path, baseline-compare path, and UAP widening loop.

## Verification

### Pre-fix baseline

```
Solver absolute_best (after previous fix):  $42,654.00/mo  [40, 70, 90]  10/9/5
Client's v5 (target):                       $43,549.00/mo  [40, 60, 90]   8/8/8
Gap:                                              -$895.00/mo
```

### Post-fix result

```
Solver absolute_best (this fix):            $43,802.00/mo  [40, 70, 90]   9/9/6
Client's v5 (target):                       $43,549.00/mo  [40, 60, 90]   8/8/8
Gap:                                              +$253.00/mo  (solver wins)
```

The solver now finds an allocation that beats the client by $253/mo, by trading some 60% units for 70% units within the rent-optimal mix. The previous gap is closed and the solver also surfaces better scenarios than what the client manually crafted.

### Test suite

```
63 tests pass (60 baseline + 3 from previous fix).
```

No new tests added — the existing haircut tests prove the rent flow; the per-combo behavior is already exercised by the `test_solver.py` suite which doesn't pass `rent_by_band_cents` and so exercises the backward-compat WAAMI-max path. End-to-end verification done via the probe scripts (B1 v5 workbook + 2026 rents + client utilities).

## What this does NOT change

- VBA / `.xlam` add-in (no client-PC refresh needed beyond the standard one)
- Rent calculator XLSX files
- Hard compliance constraints (WAAMI cap, 40% window, share thresholds, band caps)
- The 100% AMI haircut (from previous fix — still applied at the source)
- Backward compatibility for solver callers without rent data — they still get WAAMI-max behavior
- The scenario family returned (`absolute_best`, `best_3_band`, `client_oriented`, `alternative`) — same names, same semantics, just better internal allocations

## Risks / open questions

- **Solver runtime** — rent objective requires building rent_coeffs_int per combo (one int per unit per band). For 24 units × 3 bands × 50 combos = 3,600 ints per request. Negligible overhead.
- **Edge scenarios still use `find_max_revenue_scenario`** — unchanged, they already used rent objective.
- **The "closest_to_60" scenario from the previous fix may disappear** — when the per-combo rent-max converges with the WAAMI-max ranking (which is now likely for most workbooks), the re-ranking step finds the same scenario and skips adding `closest_to_60`. This is expected and correct.
- **WAAMI floor enforcement still applies** — the per-combo solve respects `waami_floor` regardless of objective mode.

## One-line PowerShell command for client-PC update

```
powershell -ExecutionPolicy Bypass -File C:\AMI_Optix_Agent\scripts\Refresh-AmiOptixAgentFromGitHub.ps1 -AgentRoot C:\AMI_Optix_Agent
```

(No VBA change; PS refresh is a no-op for this fix. Render auto-deploys; takes effect ~1-3 min after push.)

## Status

| Step | When | Result |
|---|---|---|
| Probed gap with B1 v5 + 2026 + client utilities | 2026-06-04 | ✅ Gap measured |
| Add rent_by_band_cents param to find_optimal_scenarios | 2026-06-04 | ✅ |
| Use rent objective per-combo when rent data available | 2026-06-04 | ✅ |
| Move rent_schedule loading + build rent_by_band_cents in app.py | 2026-06-04 | ✅ |
| Pass rent_by_band_cents to all find_optimal_scenarios call sites | 2026-06-04 | ✅ |
| Re-probe B1 v5: gap closed (+$253/mo over client) | 2026-06-04 | ✅ |
| Full test suite (63 pass) | 2026-06-04 | ✅ |
| Commit + push fix branch | _pending_ | ⏳ |
| Fast-forward merge to feature, push | _pending_ | ⏳ |
| Approved by client | _pending_ | ⏳ |
