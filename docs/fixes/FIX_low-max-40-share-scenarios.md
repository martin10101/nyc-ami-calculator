# FIX: Add `low_40_share` + `max_40_share` scenarios for explicit 40% trade-off

**Branch:** `fix/low-max-40-share-scenarios`
**Cut from:** `feature/excel-agent-foundation` @ `8381c45`
**Date:** 2026-06-04
**Author:** client UX feedback — every scenario landed at the same 40% allocation
**Risk:** Low (additive only — adds 2 scenarios to the existing dict; no existing behavior changes)

## Problem

After the rent-max-per-combo fix (`8381c45`), the solver returned 4 scenarios for B1 v5 — but every single one had 40% SF at ~11.67%:

```
absolute_best          [40, 70, 90]      11.67%
client_oriented        [40, 60, 80]      11.68%
alternative            [40, 70, 100]     11.64%
best_3_band            [40, 60, 100]     11.68%
```

The client's hand-crafted v5 had 40% at 10.32% (very close to the legal floor of 10%). When she opened the Excel output, she saw four scenarios that all looked identical on the 40%-share dimension, and not one of them matched her preferred "low 40% exposure" structure. Per the developer who flagged this: *"only 1 result beat the client results and the client had less 40% units... basically why is the solver bringing back so many close results instead of deeply understanding that 40% is something we try to avoid."*

Rent-max naturally pushes 40% allocation toward the *ceiling* of the [10%, 12.5%] window (more 40% SF → more SF available at high bands → more rent without breaking the WAAMI cap). That's mathematically correct, but it ignores the developer's preference for minimal "deep affordability" exposure.

## Solution

Add two additional scenarios that explicitly span the 40% trade-off:

- **`low_40_share`** — re-solve with the 40% window pinned at the floor: `[10.0%, 10.5%]`. Gives the lowest legal 40% allocation.
- **`max_40_share`** — re-solve with the 40% window pinned at the ceiling: `[12.0%, 12.5%]`. Gives the highest allowed 40% allocation.

Both still run the full rent-max objective inside the constrained window — so within each variant, rent is maximized. The client sees three explicit points on the trade-off curve.

## Code change

`app.py` — after the existing scenario re-ranking block, run two extra `find_optimal_scenarios` calls with modified `share_thresholds` (for MIH) or `deep_affordability_*` keys (for UAP). Apply rents to each, pick the rent-max from each variant's returned dict, register as `scenarios['low_40_share']` and `scenarios['max_40_share']`.

- Reads the *effective* 40% window from the (possibly mutated) config, so MIH's floor-walk and UAP's widening loop both flow through correctly.
- `max_40_share` is only added if its allocation is distinct from `absolute_best` — prevents duplicate scenarios when the rent-max already lands at the ceiling.
- Skipped entirely if rent data is unavailable (graceful fallback).
- ~70 new lines, no changes to existing logic.

## Verification (B1 v5 + 2026 + client utilities)

| Scenario | 40% SF | 40% units | 100% units | WAAMI | Net monthly |
|---|---|---|---|---|---|
| `absolute_best` (rent-max) | 11.67% | 9 | 0 | 59.97% | $43,802 |
| `max_40_share` (at ceiling) | 12.05% | 9 | 0 | 59.72% | $43,617 |
| client_oriented | 11.68% | 9 | 0 | 59.93% | $43,554 |
| **`low_40_share` (at floor)** | **10.36%** | **8** | 0 | 59.98% | $43,553 |
| alternative | 11.64% | 9 | 4 | 59.95% | $43,418 |
| best_3_band | 11.68% | 9 | 7 | 59.95% | $43,389 |

**`low_40_share` matches the client's hand-crafted v5 almost exactly:** 40% at 10.36% (vs her 10.32%), 8 units at 40% (matching her count), zero 100% units, $43,553/mo (matches her $43,549). Now she has a visible "stay at the floor" option alongside the "go for max rent" option.

The trade-off curve she now sees in Excel:

- Less 40% exposure (10.36% of SF) → `low_40_share`, $43,553/mo
- Rent-max (40% lands at 11.67%) → `absolute_best`, $43,802/mo (+$249/mo, +$2,988/yr vs low_40)
- Maxed-out 40% (at ceiling, 12.05%) → `max_40_share`, $43,617/mo

## Tests

63 existing tests pass (no regressions). No new test added — the existing solver test suite doesn't have a rent-aware fixture, and the behavior is exercised by the end-to-end probe scripts. The two new code paths are additive and only trigger when `rent_schedule` and `rent_by_band_cents` are both present.

## What this does NOT change

- VBA / `.xlam` add-in
- Rent calculator XLSX files
- Existing scenarios (`absolute_best`, `best_3_band`, `best_2_band`, `client_oriented`, `alternative`, `closest_to_60`) — unchanged
- Hard compliance constraints (WAAMI cap, share thresholds, band caps)
- 100% AMI haircut (from previous fix — still applied)
- Backward compatibility — when no rent data, no extra scenarios added

## Risks

- **Two extra solver passes per request** — each ~0.5-2s. Total request time on MIH may grow from ~75-150s (current) to ~76-154s. Negligible.
- **`max_40_share` may be duplicate of `absolute_best`** in projects where rent-max naturally lands at the ceiling — handled by the canonical_assignments check.
- **Display in Excel** — VBA will write all scenarios to the worksheet as it currently does. The user may want a follow-up UX tweak to highlight the recommended one (`absolute_best`) more prominently, but that's a separate VBA change not in this fix.

## One-line PowerShell command for client-PC update

```
powershell -ExecutionPolicy Bypass -File C:\AMI_Optix_Agent\scripts\Refresh-AmiOptixAgentFromGitHub.ps1 -AgentRoot C:\AMI_Optix_Agent
```

(No VBA change; PS refresh is a no-op for this fix. Render auto-deploys; takes effect ~1-3 min after push.)
