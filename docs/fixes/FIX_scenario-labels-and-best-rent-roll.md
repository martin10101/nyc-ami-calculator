# FIX: Scenario labels (G) + Best Rent Roll promotion (H)

**Branch:** `fix/scenario-labels-and-best-rent-roll`
**Cut from:** `feature/excel-agent-foundation` @ `0cbfbb8`
**Date:** 2026-04-29
**Author:** client request via remote session
**Risk:** Medium (changes scenario set composition, no solver behavior change)

## Status

| Step | When | Result |
|---|---|---|
| Committed locally | 2026-04-29 — sha `_pending_` | ⏳ |
| Full test suite passes locally (59/59) | 2026-04-29 | ✅ |
| Pushed `fix/scenario-labels-and-best-rent-roll` | _pending_ | ⏳ |
| Fast-forward merged into `feature/excel-agent-foundation` | _pending_ | ⏳ |
| Pushed `feature/excel-agent-foundation` (Render auto-deploy) | _pending_ | ⏳ |
| Render deploy live | _pending_ | ⏳ |
| Client PC refreshed via PS agent | _pending_ | ⏳ |
| Manual test passed on client PC | _pending_ | ⏳ |
| Approved by client | _pending_ | ⏳ |

## Problem

Two related issues in the scenario output, surfaced by the user's 230 Kent test on Fix D.2:

### G — "Max Revenue" label is misleading

`client_oriented` was labeled `"CLIENT ORIENTED (MAX REVENUE)"` in the Excel display, but it isn't actually the max revenue. Per [solver.py:942-961](../../ami_optix/solver.py#L942-L961), `client_oriented` is the highest-revenue 3-band scenario *not already chosen* as `absolute_best`/`best_3_band`/`best_2_band`/`alternative`. In the user's 230 Kent test, `client_oriented` returned **$205,428** while `absolute_best` had **$214,884** and the relaxed `edge_waami_floor_590` had **$216,120**.

### H — Highest-revenue scenario hidden as an "edge"

The actual highest-revenue scenario for many MIH buildings is `edge_waami_floor_590` (or similar) — found by the rent-max solver after relaxing the WAAMI floor from 60% to 59%. But it's labeled `"EDGE WAAMI FLOOR 590"` with a `"Tradeoffs: WAAMI 59.95% is below floor 60.00%"` warning. Per the client (2026-04-29): *"the client also asked that they wanna get one specific scenario that gives them the absolute best possible rent role, even if it means a lower utilization of the 60% AMI."* They want it as a first-class scenario.

## Code change

### app.py — promote best_rent_roll (post-processing)

After all strict + edge scenarios are generated, scan for the highest `net_annual` rent. If the winner is keyed `edge_waami_floor_*`, deep-copy it into `scenarios['best_rent_roll']` with cleaned tradeoffs and `tier='rent_max'`, and delete the original edge key to avoid duplicate display. If a strict scenario already has the highest revenue, no promotion needed.

Inserted right after the edge-scenarios block, before the de-dupe block (~line 1190):

```python
# --- Promote "Best Rent Roll" scenario (client request 2026-04) ---
try:
    best_rent_key = None
    best_rent_value = -1.0
    for _key, _scen in (scenarios or {}).items():
        if not _scen:
            continue
        _rt = _scen.get("rent_totals") or {}
        _annual = float(_rt.get("net_annual") or _rt.get("total_annual_rent") or 0.0)
        if _annual > best_rent_value:
            best_rent_value = _annual
            best_rent_key = _key
    if best_rent_key and best_rent_key.startswith("edge_waami_floor_"):
        if "best_rent_roll" not in scenarios:
            promoted = copy.deepcopy(scenarios[best_rent_key])
            promoted["tradeoffs"] = []
            promoted["tier"] = "rent_max"
            promoted_settings = promoted.get("edge_settings") or {}
            promoted_settings["promoted_to_best_rent_roll"] = True
            promoted["edge_settings"] = promoted_settings
            scenarios["best_rent_roll"] = promoted
            del scenarios[best_rent_key]
            ...
            notes.append(...)
except Exception as e:
    notes.append(f"Warning: best_rent_roll promotion failed: {str(e)}")
```

`scenario_priority` list (~line 1199): `"best_rent_roll"` inserted right after `"absolute_best"` so it appears as Scenario 2 in Excel.

### ResultsWriter.bas:3020 FormatScenarioName

- `"client_oriented"` → display `"CLIENT ORIENTED"` (was `"CLIENT ORIENTED (MAX REVENUE)"`).
- New `Case "best_rent_roll" → "BEST RENT ROLL"`.

### Ribbon.bas:1567 FormatScenarioNameForPicker

- `"client_oriented"` → display `"Client Oriented"` (was `"Client Oriented (Max Revenue)"`).
- New `Case "best_rent_roll" → "Best Rent Roll"`.

### Files NOT touched

- `ami_optix/solver.py` — solver logic unchanged. Promotion is pure server-side post-processing.
- Existing tests — `test_optimize_mih_option1_meets_40_band_min_share_floor` and the Option 4 variant iterate ALL scenarios. `best_rent_roll` is subject to the same 40-band floor check automatically (because it was promoted from a relaxed-floor scenario, which itself respected the share constraint).

## Expected effect on 230 Kent results

**Before:**
```
1. ABSOLUTE BEST                 $214,884   WAAMI 60.00%
2. BEST 3-BAND                   $211,524   WAAMI 60.00%
3. ALTERNATIVE                   $209,796   WAAMI 60.00%
4. CLIENT ORIENTED (MAX REVENUE) $205,428   WAAMI 60.00%   <-- wrong label
5. EDGE WAAMI FLOOR 590          $216,120   WAAMI 59.95%   <-- hidden, has highest revenue
```

**After:**
```
1. ABSOLUTE BEST     $214,884   WAAMI 60.00%
2. BEST RENT ROLL    $216,120   WAAMI 59.95%   <-- promoted, no tradeoff warning
3. BEST 3-BAND       $211,524   WAAMI 60.00%
4. ALTERNATIVE       $209,796   WAAMI 60.00%
5. CLIENT ORIENTED   $205,428   WAAMI 60.00%   <-- honest label
```

## Manual test (after deploy)

1. Push branch + ff-merge to feature + push feature.
2. Standard PS one-liner on client PC (covers both server Python change and VBA label updates).
3. Close + reopen Excel + 230 Kent_Unit Schedule - MIH v3.xlsm.
4. Click **Run MIH** → confirm preflight popup (Fix A) → click YES.
5. ✅ Scenario 2 reads `"BEST RENT ROLL"` with **no** "Tradeoffs:" warning, has the highest annual rent of all scenarios.
6. ✅ Whatever was previously called `"CLIENT ORIENTED (MAX REVENUE)"` now reads just `"CLIENT ORIENTED"` (no parenthetical).
7. ✅ All scenarios still satisfy 40 band ≥ 10% (Fix D.2 still in effect).
8. ✅ The previously displayed `"EDGE WAAMI FLOOR 590"` no longer appears (it was promoted/renamed).
9. Run UAP → same labels apply. UAP unchanged in behavior.

If a building's strict `absolute_best` already has the highest revenue (no relaxed-floor edge with higher rent), then `best_rent_roll` simply doesn't get added — only one scenario position is freed up. That's correct behavior.

## Rollback

```bash
git checkout feature/excel-agent-foundation
git revert <commit-sha-on-fix-branch>
git push origin feature/excel-agent-foundation
```

Then run the standard PS one-liner on the client PC.

## One-line PowerShell command for client-PC update

```
powershell -ExecutionPolicy Bypass -File C:\AMI_Optix_Agent\scripts\Refresh-AmiOptixAgentFromGitHub.ps1 -AgentRoot C:\AMI_Optix_Agent
```
