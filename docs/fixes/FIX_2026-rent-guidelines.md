# FIX: Add 2026 HPD rent guidelines

**Branch:** `fix/2026-rent-guidelines`
**Cut from:** `feature/excel-agent-foundation` @ `3629b8a`
**Date:** 2026-05-24
**Author:** client request (2026 HPD rent guidelines released)
**Risk:** Low (binary file drops + 1-line config + surgical parser fix; 2025/2024 byte-identical)

## Status

| Step | When | Result |
|---|---|---|
| Probe 2026 file layout vs 2025 | 2026-05-24 | ✅ Identical layout, 11 utility options, only $ values differ |
| Discover latent parser bug (electricity column) | 2026-05-24 | ✅ Diagnosed |
| Snapshot 2025 baseline rents | 2026-05-24 | ✅ |
| Commit 1 (parser fix + file drops + acceptance config) | 2026-05-24 — sha _pending_ | ⏳ |
| Verify 2025 byte-identical post-fix | 2026-05-24 | ✅ |
| Verify 2026 numbers match HPD published values | 2026-05-24 | ✅ |
| Local pytest (7 tests in rent_calculator + api_evaluate) | 2026-05-24 | ✅ all passed |
| Commit 2 (this doc) | 2026-05-24 — sha _pending_ | ⏳ |
| Pushed `fix/2026-rent-guidelines` | _pending_ | ⏳ |
| Fast-forward merged into `feature/excel-agent-foundation` | _pending_ | ⏳ |
| Pushed `feature/excel-agent-foundation` (triggers Render auto-deploy) | _pending_ | ⏳ |
| Approved by client | _pending_ | ⏳ |

## One-line PowerShell command for client-PC update

```
powershell -ExecutionPolicy Bypass -File C:\AMI_Optix_Agent\scripts\Refresh-AmiOptixAgentFromGitHub.ps1 -AgentRoot C:\AMI_Optix_Agent
```

**Note:** No VBA changed in this fix — the ribbon dropdown already offered 2026 (year range is 2022–2026 in `AMI_Optix_API.bas:19` and `AMI_Optix_Ribbon.bas:20`). The PS one-liner pulls the new bundled 2026 rent calculator into the local agent cache. Once Render is live, the next API call uses 2026 if the user selects it from the ribbon.

## Problem

NYC HPD released the 2026 rent guidelines. The client needs to be able to pick "2026" from the year dropdown in the AMI Optix ribbon and have it work end-to-end for **both UAP and MIH**, alongside the existing 2024 and 2025 calculators.

During the rollout I discovered a latent parser bug that had been masked for 2025 (see "Bonus: latent parser bug" below).

## What's already in place (no changes needed)

- **VBA dropdown already supports 2026.** `RENTROLL_YEAR_MAX = 2026` in both `excel-addin/src/AMI_Optix_API.bas:19` and `excel-addin/src/AMI_Optix_Ribbon.bas:20`.
- **Server auto-seeds bundled rent calculators on startup.** `app.py:111-124` walks `tools/excel-agent/assets/rent-roll-years/{YEAR}/RentCalculator_{YEAR}.xlsx`.
- **VBA activates server-side calc on every API request.** `EnsureRentCalculatorYearActive` in `AMI_Optix_API.bas:127` flips the active rent calculator before each request, so `/api/optimize` (which uses server-global active calc) ends up using the year the user picked. **Same flow for UAP and MIH.**
- **API year validator already accepts 2026.** `_normalize_rent_roll_year` in `app.py:317` accepts any year 1900-2100.

## Changes in this fix

### 1. File drop (binary)
- `tools/excel-agent/assets/rent-roll-years/2026/RentCalculator_2026.xlsx` — bundled into the repo. On every server startup, `app.py:111-124` auto-seeds this into `rent_calculators/AMI_Optix_Rent_Calculator_2026.xlsx` (the runtime working directory). `rent_calculators/` itself is gitignored, so only the bundled path needs to be committed.

### 2. Acceptance config (`tools/excel-agent/config/acceptance.template.json`)
Added `2026` to `cacheWarmupYears`, so the client agent caches the 2026 calculator during acceptance runs.

### 3. Bonus: latent parser bug fix (`ami_optix/rent_calculator.py`)

**Symptom:** When loading the fresh 2026 HPD file, the parser silently dropped the entire `electricity` category. The Apartment Electricity column's allowance values ($96/$108/$142/...) got stored under `hot_water["N/A or owner pays"]` instead. Result: ALL 2026 rent calculations were wrong — when tenant pays electricity, allowance was undercounted by ~$108/1BR; when owner pays everything, allowance was overcounted by ~$108/1BR.

**Root cause:** The Apartment Electricity column (col B) has no per-option label row — it has only one paid option ("Tenant Pays"), so HPD encodes that intent in the dropdown at row 17 rather than as a static label. The parser was reading the dropdown VALUE as if it were the option label. In the user's 2025 file the dropdown happened to read "Tenant Pays" (because someone had set it long ago), masking the bug. The fresh 2026 file has the default "N/A or owner pays" → parser stored col B values under the wrong key.

**Fix:** Three new functional lines in `_parse_allowances` (around line 240). When `current_category == 'electricity'` and the recovered option string is the dropdown sentinel "N/A or owner pays", force the canonical label `"Tenant Pays"`. This is the only meaningful paid option for that category, so the mapping is unambiguous.

**Why it can't change existing 2025/2024 behavior:** The fix branch only fires when the option string came back as the dropdown sentinel. Your existing 2025 file has "Tenant Pays" in B17 — fix is skipped, code path unchanged. Verified byte-identical via side-by-side snapshot of 8 unit/AMI combinations across 3 utility profiles.

## Verification (local, pre-deploy)

| Check | Expected | Result |
|---|---|---|
| Probe: 2026 sheet name = "AMI & Rent" | yes | ✅ |
| Probe: 2026 has 11 utility options across 4 categories | yes | ✅ |
| Probe: 2026 layout = 2025 layout (rows 14-17 pandas) | identical | ✅ |
| Side-by-side: 2025 rent grid (24 cells) pre-fix vs post-fix | byte-identical | ✅ |
| 2026 electricity Tenant Pays 1BR | $108 (matches HPD published) | ✅ |
| 2026 owner_pays_all allowance for all 8 test units | $0 each | ✅ |
| 2026 all_tenant_pays_gas 1BR total = $108+$33+$86+$26 | $253 | ✅ |
| `pytest tests/test_rent_calculator.py tests/test_api_evaluate.py` | all pass | ✅ 7/7 |

## Post-deploy verification (client-side)

1. Server: `GET /api/rent-calculators` lists `AMI_Optix_Rent_Calculator_2026.xlsx`.
2. VBA: open ribbon, click year dropdown, confirm 2026 is selectable.
3. UAP end-to-end: pick year=2026, run Manual Calculate, confirm rents come back.
4. MIH end-to-end: pick year=2026, run Find Optimal Scenarios on `3320 Atlantic_Unit Schedule - MIH v6.xlsb`, confirm scenarios return and net rents match the official 2026 HPD calculator within $1.
5. Acceptance: run PS refresh + acceptance, confirm `cacheWarmupYears` seeds 2026 and `Verify Manual Rents (API)` passes for a 2026 scenario.
