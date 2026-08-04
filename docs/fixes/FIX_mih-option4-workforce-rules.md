# FIX: MIH Option 4 uses the real Workforce Option rules

**Branch:** `fix/mih-option4-workforce-rules`
**Cut from:** `feature/excel-agent-foundation` @ `ab78cf5`
**Date:** 2026-08-04
**Risk:** Moderate scope, server-only. Option 1 behavior covered by regression tests (87 passed). No Excel/.xlam change — no PC updates needed.

## Symptom (client, 2026-08 — "Option 4.xlsb", 701 Myrtle portfolio)

Running a real MIH **Option 4 (Workforce Option)** project produced "incorrect AMI levels": every scenario forced ~22 units to 40% AMI, capped bands at 100%, and returned $172,485/mo — while the client's own hand allocation (bands 50/70/130, zero 40s) was worth $232,277/mo. The program was ~$60K/mo below a legitimate manual allocation and would have flagged that allocation as non-compliant.

## Root cause

The server's Option 4 config was Option 1 in disguise ([app.py](../../app.py), `_build_program_config`):

1. **Option 1's [10%, 12.5%]-at-40% window was copied onto Option 4** ("same 40% window as Option 1", 2026-04). The Workforce Option has **no 40% AMI set-aside at all** (ZR 23-154(d); NYC Council MIH page; the client's own workbook Prog sheet).
2. **The 100% band hard cap** (client decision 2026-05-18, intended for Option 1) was applied to **all** MIH — but the Workforce Option's 115% *average* is unreachable without its 110–135% bands.
3. Downstream, every 40%-centric extra (floor-walk, `low_40_share`/`max_40_share`, the fewest-40 ladder, `mid_40_share`, the min-count frontier) contains a `... or 0.10` fallback that would **re-fabricate** a 10%-at-40 requirement even with the window removed.
4. The `recommended_key` picker only considers scenarios with >0 units at 40% ("fewest at 40%" model) — for Option 4 it crowned a $190K ballast scenario over the $259K rent-max, and Excel's ApplyBestScenario writes the recommended scenario to the MIH page.

## Confirmed rules (client, 2026-08-04; ZR 23-154(d))

Option 4 / Workforce: **no 40% requirement · ≥5% at ≤70% · ≥5% at ≤90% (encoded cumulatively as ≥10% at ≤90) · weighted average ≤115% · bands up to 135% · at most 4 income bands.**

## Change (all server-side, app.py)

- `_build_program_config`: Option 4 share thresholds are now only `70% ≥5%` and `90% ≥10%` (no 40 entry). `MIH_HARD_BAND_CAP = 100 if Option 1 else 135`. WAAMI cap 115 and max 4 bands unchanged (already correct).
- New `mih_forty_required` flag in `/api/optimize` (true iff a ≤40 threshold with a min share exists; always true for non-MIH). Gates: the 40%-window floor-walk, `low_40_share`/`max_40_share`, the fewest-40 ladder, `mid_40_share`, and the min-count frontier — none of which apply to a no-forty option, and all of which would otherwise re-impose 10% via fallbacks.
- `recommended_key`: when `mih_forty_required` is false, recommend the straight **rent-max** scenario instead of the fewest-40 picker.

## Result on the client's building (76-unit pool, residential 123,472 SF, her utilities, 2026 rents)

| | Monthly income | 40% units | WAAMI | Bands |
|---|---|---|---|---|
| Old server (bug) | $172,485 | 22 forced | 80% | 40/100 |
| **Client's hand allocation** | **$232,277** | 0 | ~113% | 50/70/130 |
| **Fixed server (recommended)** | **$258,934** | 0 | 115.0% | 70/90/130/135 |

The fixed program beats her manual allocation by **+$26,657/mo (~$320K/yr)**, fully compliant (set-asides met, average exactly at the 115% cap, 4 bands).

## Verification

- `tests/test_mih_option4.py` (new): config unit tests + API tests asserting no forced 40s, bands >100 used, ≤4 bands, set-asides, no 40-families, recommended = rent-max; plus an Option 1 regression test (40-window, 100 cap, families intact).
- Three legacy tests that encoded the superseded rules updated to the confirmed rules (`test_api_evaluate.py`, `test_api_optimize_learning.py` ×2).
- Full suite: **87 passed**.

## Not addressed here (separate fixes, display-level, Excel-side)

- Scenario names hard-mapped from keys (e.g. "FEWEST 40% UNITS" on a non-fewest option; "HIGHER RENT (MORE 40% UNITS)" with equal rent/count).
- Outcome-duplicate scenarios (same bands + same rent shown twice).
- Excel's 40%-centric overview headings look odd on Option 4 runs (cosmetic).

## Deploy

Server-only → push to GitHub → Render auto-deploys. No .xlam change, no per-PC steps.
