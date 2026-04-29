# FIX: MIH — 40% AMI band must hit ≥10% of full residential building SF

**Branch:** `fix/mih-40-ami-floor-10pct`
**Cut from:** `feature/excel-agent-foundation` @ `2510b90`
**Date:** 2026-04-29
**Author:** client request via remote session
**Risk:** Medium-High (changes solver behavior for MIH, affects scenario output)

## Status

| Step | When | Result |
|---|---|---|
| Committed locally | 2026-04-29 — sha `_pending_` | ⏳ |
| Full test suite passes locally (59/59) | 2026-04-29 11:XX | ✅ |
| Pushed `fix/mih-40-ami-floor-10pct` | _pending_ | ⏳ |
| Fast-forward merged into `feature/excel-agent-foundation` | _pending_ | ⏳ |
| Pushed `feature/excel-agent-foundation` (triggers Render auto-deploy) | _pending_ | ⏳ |
| Render deploy live (this is the FIRST deploy with a real Python change) | _pending_ | ⏳ |
| Client PC refreshed via PS agent (column-rename only, server already deployed) | _pending_ | ⏳ |
| Manual test passed on client PC | _pending_ | ⏳ |
| Approved by client | _pending_ | ⏳ |

**Unlike Fix C / Fix A, this is the first fix that actually changes Render's runtime behavior** — the new constraint is enforced on the server side. Wait for the Render deploy to be live before testing in Excel.

## Problem

Per client meeting 2026-04-29, the MIH rule for the 40% AMI band needs to be a **hard floor at 10% of full residential building SF** (not a ceiling, which is what the code did before). The optimizer should naturally land as close to 10% as possible, in small increments above. Applies to **all MIH** options (1 AND 4). Other rules unchanged.

In the screenshot the client shared (230 Kent), the 40% band currently lands at 8.98% of full building SF, violating the new rule. After this fix, it must be ≥10%.

## What changed in the code

### 1. `app.py:200-219` — share_thresholds for both MIH options

**Option 1**: `'max_share': 0.10` (CEILING) → `'min_share': 0.10` (FLOOR). One field rename.

**Option 4**: ADDED `{'band_threshold': 40, 'min_share': 0.10, 'denominator': 'residential'}` at the top of the existing list. Rule applies to all MIH per client spec (was Option-1-only before).

### 2. `app.py:811-852` — MIH floor-walking loop wrapping `find_optimal_scenarios`

Per client spec: *"if it doesn't have scenarios, then it keeps walking it a little bit more and more upper till it finds enough scenarios that it's, you know, it can bring back. While re-running all calculations on each individual, you know, increase."*

Implementation:
- Start: `min_share = 0.100` for the 40 band threshold.
- If `find_optimal_scenarios` returns no `absolute_best` scenario, bump `min_share` by `0.001` (a 0.1% step) and rerun the optimizer.
- Cap: `0.150` (15%). If we hit the cap, return the last empty result with an explanatory note.
- Stop on first non-empty result. Most buildings succeed at 0.100 on the first iteration; walking only triggers in edge cases.
- After walk settles, the same final `min_share` flows into the subsequent `find_max_revenue_scenario` call (config is mutated in place).

Mirrors the existing UAP `deep_affordability_max_share` widening pattern at `app.py:826-859`.

#### Math caveat (worth knowing for future debugging)

CP-SAT-wise, walking `min_share` UP makes the constraint *stricter* (requires more SF at 40% band), not looser. So for buildings where 10% is infeasible due to the WAAMI **floor** (achievable WAAMI < 59.1%), walking up makes it *more* infeasible and the loop will burn iterations to the cap and return empty.

The walk is genuinely helpful in **one specific case**: when the binding constraint at min=0.10 is the WAAMI **cap** (60%) being exceeded — pushing more SF into the 40% band lowers WAAMI, which can slip it under the cap. The user's spec accepted this trade-off.

#### Performance

- **Typical building** (10% feasible): 1 iteration ≈ 30s, identical to today's runtime.
- **Edge building** that needs walking: each iteration is a full CP-SAT solve cycle (~30s). Worst case = 51 iterations = ~25 min for one `/api/optimize` call. If this bites in production, coarsen the step (`MIH_FLOOR_STEP = 0.005` → 11 iterations max) or short-circuit at a smaller iteration count.

### 3. `excel-addin/src/AMI_Optix_ResultsWriter.bas:2780,2785` — column header renames

- `"Share of SF"` → `"Share of SF AMI"` (line 2780). Clarifies this is share of the affordable AMI pool.
- `"Share of Building SF"` → `"Share of Full Building SF"` (line 2785). Clarifies this uses the full residential building denominator.

Single function (`WriteScenarioSummaryAndTable` band-mix block); applies to both Scenario Manual and per-scenario blocks.

### 4. `tests/test_api_evaluate.py:55-129` — flipped Option 1 ceiling test, added Option 4

- **Renamed** `test_evaluate_mih_option1_enforces_40_band_max_share_cap` →
  `test_evaluate_mih_option1_enforces_40_band_min_share_floor`. Payload now has 0% at 40 band; asserts error contains `"below required 10.00%"` instead of `"above maximum 10.00%"`.
- **Added** `test_evaluate_mih_option1_floor_satisfied_passes_share_constraint` — assignment with exactly 10% at 40 band must NOT trip the share-floor error.
- **Added** `test_evaluate_mih_option4_enforces_40_band_min_share_floor` — covers the new Option-4 rule.

### 5. `tests/test_api_optimize_learning.py:58-126` — flipped Option 1 ceiling assertion, added Option 4

- **Renamed** `test_optimize_mih_option1_never_exceeds_40_band_max_share` →
  `test_optimize_mih_option1_meets_40_band_min_share_floor`. Asserts every scenario has `low_sf >= 0.10 * residential_sf` (was `<=`).
- **Added** `test_optimize_mih_option4_meets_40_band_min_share_floor`.

### Files NOT touched (verified)

- `ami_optix/solver.py` — already supports `min_share` correctly. Zero solver code changes needed.
- `ami_optix/rent_calculator.py` — unrelated.
- UAP code paths — completely separate from MIH constraints. UAP edge scenarios (`app.py:1053-1099`) use legacy `deep_affordability_min/max_share`, not `share_thresholds`.
- MIH edge scenarios (`app.py:1106-1119`) — only relax WAAMI floor; don't touch share rules.
- `customUI/customUI*.xml` — no ribbon changes.

## Manual test (after Render deploy is live)

1. **Wait** for Render to finish auto-deploying after the push. Verify `cfc6bbc`-equivalent SHA is live before testing.
2. Run the standard one-line PS command on the client PC (for the column-rename change in the `.xlam`).
3. Close + reopen Excel + `230 Kent_Unit Schedule - MIH v3.xlsm` (the screenshot's building, 8.98% before this fix).
4. Click **Run MIH**.
5. ✅ Expected: every returned scenario has `Share of Full Building SF` **≥ 10%** on the 40% AMI row. Most scenarios should land at 10.0%-11.5% (smallest feasible amount above the floor).
6. ✅ Confirm column headers in Band Mix table read `Share of SF AMI` and `Share of Full Building SF` (not the old labels).
7. ⚠ **Possible**: building turns out infeasible (no band combo satisfies ≥10% floor + ≥59.1% WAAMI floor + ≤60% WAAMI cap + ≤3 bands). In that case the API returns no scenarios with a clear note in the analysis_notes. **Per user spec, this is correct** — but worth knowing which buildings hit it. If the screenshot's building hits this, the floor-walking loop will try 0.101, 0.102, ... up to 0.150 before giving up. Worst-case wait ~25 min for that one optimize call.
8. Run a UAP workbook → ✅ scenarios identical to before, no impact.
9. Run an MIH Option 4 workbook → ✅ scenarios respect the new 40% band floor.

## Rollback

```bash
git checkout feature/excel-agent-foundation
git revert <commit-sha-on-fix-branch>
git push origin feature/excel-agent-foundation
```

Render auto-redeploys the reverted Python. Then run the standard PS one-liner on the client PC for the column-header revert.

## One-line PowerShell command for client-PC update

```
powershell -ExecutionPolicy Bypass -File C:\AMI_Optix_Agent\scripts\Refresh-AmiOptixAgentFromGitHub.ps1 -AgentRoot C:\AMI_Optix_Agent
```

Same one-liner for every fix in this project. Pulls latest `feature/excel-agent-foundation`, rebuilds the staged `.xlam`, runs acceptance.

## Test results

Full test suite (59 tests) passed locally on the fix branch before push:
- `test_api_evaluate.py` — 6/6 (3 new + 3 existing) ✅
- `test_api_optimize_learning.py` — 4/4 (2 new + 2 existing) ✅
- `test_solver.py` — 17/17 ✅ (incl. golden workbook + 5 sample workbook tests)
- `test_rent_calculator.py`, `test_validator.py`, `test_parser.py`, etc. — all green ✅
