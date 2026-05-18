# FIX: Cap MIH at 100% AMI band (Option 1 + Option 4)

**Branch:** `fix/mih-100-band-cap`
**Cut from:** `feature/excel-agent-foundation` @ `15e06ee`
**Date:** 2026-05-18
**Author:** client request via remote session
**Risk:** Low (3 functional lines server-side, no VBA change, UAP untouched)

## Status

| Step | When | Result |
|---|---|---|
| Commit 1 (app.py hard cap) | 2026-05-18 — sha `ca128ea` | ✅ |
| Commit 2 (regression test) | 2026-05-18 — sha `cd750a9` | ✅ |
| Local pytest (28 tests including new one) | 2026-05-18 | ✅ all passed in 17s |
| Commit 3 (this doc) | 2026-05-18 — sha _pending_ | ⏳ |
| Pushed `fix/mih-100-band-cap` | _pending_ | ⏳ |
| Fast-forward merged into `feature/excel-agent-foundation` | _pending_ | ⏳ |
| Pushed `feature/excel-agent-foundation` (triggers Render auto-deploy) | _pending_ | ⏳ |
| Render deploy live | _pending_ | ⏳ |
| Manual test passed on client PC | _pending_ | ⏳ |
| Approved by client | _pending_ | ⏳ |

## One-line PowerShell command for client-PC update

Run on the client PC (cmd.exe, Run dialog, or PowerShell):

```
powershell -ExecutionPolicy Bypass -File C:\AMI_Optix_Agent\scripts\Refresh-AmiOptixAgentFromGitHub.ps1 -AgentRoot C:\AMI_Optix_Agent
```

**Note for this fix specifically:** the code change is server-side only (Python in `app.py`). The `.xlam` add-in has NOT changed. The PS one-liner is mostly a no-op for this fix (it'll rebuild the `.xlam` from the same source it already has). The fix takes effect once Render finishes auto-deploying the push. **Wait for the Render deploy to go green** before testing — typically 1-3 minutes after the GitHub push.

You can also test simply by re-running MIH from the existing Excel session (no PS, no Excel restart needed) — as soon as Render is live, the next `/api/optimize` call hits the new code path.

## Problem

Client decision 2026-05-18: MIH must never use bands above 100% AMI. Bands 110, 120, 130, 135 are removed from the MIH eligible-band universe for both Option 1 and Option 4. Until now MIH could reach 130 (or 135 in some workbooks).

All other MIH rules stay exactly as today:
- 40% band floor at 10% of full residential SF (Fix D / D.2)
- 40% band ceiling at 12.5% of full residential SF (Fix D.2)
- 60% WAAMI cap and 59.1% WAAMI floor + edge relaxation to 58%
- 3-band max (Option 1) / 4-band max (Option 4)
- Premium-score logic for unit-to-band assignment
- `absolute_best` / `best_rent_roll` / `best_3_band` / `client_oriented` scenario set
- Floor-walking loop if the strict 10% floor is infeasible

UAP completely untouched.

**Trade-off accepted by client:** some MIH buildings that today produce scenarios using 110/120/130 bands may now produce fewer scenarios, or none in the extreme case where the higher bands were the only way to hit the 59.1% WAAMI floor under the 60% cap.

## Diagnosis (Phase 1, read-only)

The MIH band universe was already parameterized via `mih_max_band_percent`. Found in `_build_program_config` ([app.py](../../app.py) lines 234-239 pre-fix):

```python
# Band cap is enforced by limiting potential bands.
if mih_max_band_percent is None:
    mih_max_band_percent = 135
max_band = int(mih_max_band_percent)
candidate_bands = [40, 60, 70, 80, 90, 100, 110, 120, 130, 135]
rules['potential_bands'] = [b for b in candidate_bands if b <= max_band and b != 50]
```

Flow:
1. MIH workbook's `Prog` sheet has a cell that `TryReadMIHInputs` (Main.bas:801) reads as `mihMaxBandPercent`.
2. VBA includes `mih_max_band_percent` in every `/api/optimize` and `/api/manual_calculate` payload (API.bas:590, 688; VerifyManualRents.bas:638).
3. Server `_build_program_config` uses it to filter `potential_bands`.
4. `potential_bands` is the single chokepoint — both `find_optimal_scenarios` (solver.py:484) and `find_max_revenue_scenario` (solver.py:652) read from it. MIH edge scenarios only relax the WAAMI floor; they don't add bands.

So capping at 100 in `_build_program_config` propagates everywhere. UAP is in a completely separate branch of the same function (returns early at app.py:184-186) and is not touched.

## Code change

### Commit 1 — `ca128ea` — `app.py:234-246`

**Before:**
```python
# Band cap is enforced by limiting potential bands.
if mih_max_band_percent is None:
    mih_max_band_percent = 135
max_band = int(mih_max_band_percent)
candidate_bands = [40, 60, 70, 80, 90, 100, 110, 120, 130, 135]
rules['potential_bands'] = [b for b in candidate_bands if b <= max_band and b != 50]
```

**After:**
```python
# Band cap. Client decision 2026-05-18: MIH must NEVER use bands above 100% AMI
# (Option 1 AND Option 4). We hard-cap on the server regardless of what the
# client workbook sends, so we don't have to re-edit every property workbook
# in the field - existing workbooks may still set 130 or 135 in the Prog sheet
# and we silently floor that to 100.
# To intentionally allow a LOWER cap (e.g. 80) the workbook can still send a
# smaller mih_max_band_percent; only the upper bound is enforced.
MIH_HARD_BAND_CAP = 100
if mih_max_band_percent is None:
    mih_max_band_percent = MIH_HARD_BAND_CAP
max_band = min(int(mih_max_band_percent), MIH_HARD_BAND_CAP)
candidate_bands = [40, 60, 70, 80, 90, 100, 110, 120, 130, 135]
rules['potential_bands'] = [b for b in candidate_bands if b <= max_band and b != 50]
```

Net: 1 new constant (`MIH_HARD_BAND_CAP = 100`), 1 `if`-default swap, 1 `min(...)` wrap. 5 lines of comment explaining the why. `candidate_bands` left untouched so the file still reads as a complete list of conceivable AMI bands — the filter does the cap work.

**Why hard-cap server-side instead of updating workbooks:** workbooks live on multiple client PCs, multiple per property. Server is the single source. If the client ever changes their mind (e.g., raise to 110), it's a one-line edit to `MIH_HARD_BAND_CAP`.

### Commit 2 — `cd750a9` — `tests/test_api_optimize_learning.py`

Added `test_optimize_mih_caps_bands_at_100_regardless_of_payload`. Sends `mih_max_band_percent: 135` (the OLD default), iterates Option 1 and Option 4, asserts every returned assignment has `assigned_ami <= 1.00`.

Local pytest result: **28 passed in 17s** (including the new test + all pre-existing MIH/UAP/solver tests).

## Files NOT touched

- `ami_optix/solver.py` — already correctly reads `potential_bands`; no solver change
- `ami_optix/rent_calculator.py` — unrelated
- All `excel-addin/src/*.bas` — workbook can still send `mih_max_band_percent: 130/135`; server caps. **No .xlam rebuild needed.**
- `rules_config.yml` — already declares `[40, 60, 70, 80, 90, 100]` as base; MIH overrides are programmatic
- UAP code path in `_build_program_config` (line 184-186) — provably untouched
- Edge scenarios — reuse `potential_bands`, which is capped
- `find_max_revenue_scenario` / `best_rent_roll` promotion — reuses `potential_bands`, capped
- Manual Calculate path — `/api/manual_calculate` is unconstrained by design, so users can still manually type 130 into a sheet cell for what-if exploration; only the solver respects the cap

## Manual test

The change is server-side only — `.xlam` does not need rebuilding. Run after Render is live (1-3 min after push).

1. Open an MIH property workbook (e.g., `3320 Atlantic_Unit Schedule - MIH v6.xlsb`).
2. Click **Run MIH** (Option 1 or Option 4).
3. ✅ Examine returned scenarios. No assignment may have AMI > 100. Bands appearing in Band Mix should only be from {40, 60, 70, 80, 90, 100}.
4. ✅ 40% band floor still works: every scenario's 40% band SF ≥ 10% of `mih_residential_sf` (≤ 12.5% under Fix D.2, with floor-walking slack up to 17.5%).
5. ✅ WAAMI: all scenarios at ≤ 60%, ≥ 58% (edge) or ≥ 59.1% (strict).
6. ✅ Scenario set: `absolute_best`, `best_rent_roll`, `best_3_band`, `client_oriented` (or whichever apply) all present.
7. ✅ 130 band is NOT present in any scenario, even though the workbook's `Prog` sheet still says max=130 (or 135).
8. Open a UAP workbook → click **Run UAP** → ✅ scenarios identical to before this fix (UAP can still use 110/130/175 bands).

### Edge case worth confirming

3320 Atlantic (~59K SF residential) currently produces scenarios that include 130-band assignments. Possible outcomes after the cap:
- Fewer scenarios than before (some band combos including 130 are now infeasible).
- All scenarios have lower total annual rent than before (less rent ceiling).
- In the worst case, scenarios collapse to zero — floor-walking widens the 40% window upward to try; if even that fails, the building genuinely cannot satisfy all constraints simultaneously under the new cap. Accepted trade-off.

### Server-side regressions

```bash
python -m pytest tests/test_api_optimize_learning.py tests/test_api_evaluate.py tests/test_solver.py -v
```

Verified locally on 2026-05-18: 28 passed in 17s.

## Rollback

```bash
git checkout feature/excel-agent-foundation
git revert ca128ea cd750a9     # restores old 135 default + removes new test
```

Or, if you want to keep the cap but adjust the value:

```python
# In app.py _build_program_config, change one line:
MIH_HARD_BAND_CAP = 110   # or 120, 130, etc.
```

Both routes are server-side only. Re-run pytest, push, Render auto-deploys.

## Deploy notes

- Render auto-deploy is **ON**. Pushing the merged feature branch triggers a deploy. **Wait for it to go green** — this fix's behavior only takes effect once the server is live.
- Client-PC rollout via PowerShell agent is optional (no `.xlam` change), but harmless to run for workflow consistency.
- No persistent disk / Render config changes needed.
