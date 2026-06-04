# FIX: Strict WAAMI cap enforcement — drop scenarios above 60.00% (even by a sliver)

**Branch:** `fix/strict-waami-cap-enforcement`
**Cut from:** `feature/excel-agent-foundation` @ `6bae896`
**Date:** 2026-06-04
**Author:** client requirement — WAAMI must not exceed the legal cap by any amount
**Risk:** Low (only drops scenarios that are already flagged as non-compliant)

## Problem

The client's Excel output was showing `SCENARIO 4: MAX REVENUE` with a tradeoff message *"WAAMI 60.00% exceeds cap 60.00%"*. Numerically (extracted from the screenshot):

```
40% units: 9, total SF 5,869.70
60% units: 5, total SF 2,763.62
80% units: 10, total SF 5,869.73
TOTAL SF: 14,503.05
WAAMI = (0.4 × 5869.70 + 0.6 × 2763.62 + 0.8 × 5869.73) / 14503.05
      = 8,701.836 / 14,503.05
      = 0.6000041378
      = 60.0000414%
```

So WAAMI is over the 60% cap by **0.0000414 percentage points**. Microscopic, but real. In display ("60.00%") it looks compliant — same as the cap — but the tradeoff message correctly says it exceeds.

The client's hard requirement: **WAAMI must never exceed 60.00% under any circumstances, even by a fraction of a percent.** A scenario above the cap should not be surfaced at all — showing it with a tradeoff warning creates the risk that a user picks it without reading the fine print and ends up with a technically non-compliant rent roll.

## Why the scenario was being returned

The `max_revenue` edge scenario sets the solver's `waami_floor = 60%` (= cap), asking the solver to find scenarios "exactly at the cap." Due to float-vs-integer rounding in the SF basis-points math, the returned scenario can land 0.0000414% above the cap. The existing validation correctly flagged this as a tradeoff but still returned the scenario.

## Fix

Two coordinated drops, both server-side, no VBA change:

### A. In `_maybe_add_edge` (app.py around line 1195)

After `compute_rents_for_assignments` populates the candidate, compute the actual float WAAMI from the assignments. If it exceeds the strict cap by any amount (strict `>`, no epsilon), `return False` and the scenario is never added to the dict. A note is appended for diagnostic traceability.

### B. Final belt-and-suspenders pass before response

After all scenario generation (main + re-ranking + low/max_40 + edges) is complete, scan every entry in `scenarios`. Drop any whose float-computed WAAMI exceeds the cap. Append a note for each dropped scenario showing its actual WAAMI to 6 decimal places.

Belt-and-suspenders structure means: even if a future code path adds another scenario source, the final pass catches anything that slipped through.

## Verification

```
63 tests pass (no regressions).
```

Screenshot's max_revenue scenario (WAAMI = 60.0000414%) would now be **dropped silently** from the response. The note `"Removed scenario 'max_revenue' from results: WAAMI 60.000041% exceeds the 60.00% cap (client requires strict compliance)."` appears in the response notes so the developer can see why it's gone.

The user will no longer see any "WAAMI exceeds cap" tradeoff in Excel. All visible scenarios will be strictly compliant at the basis-points level.

## What this does NOT change

- VBA / `.xlam` add-in
- The actual solver objective (still rent-max)
- The 100% AMI haircut (from previous fix)
- low_40_share / max_40_share (from previous fix)
- Main scenarios from the integer solver (which are already compliant by construction)
- Hard compliance constraints (40% window, share thresholds, band cap)

## Risks

- **Edge scenarios may now be empty for some workbooks.** If `max_revenue`, `edge_max_share_*`, `edge_min_share_*` all overshoot the cap, the response will have fewer scenarios. The main scenarios (`absolute_best`, `best_3_band`, `closest_to_60`, `low_40_share`, `max_40_share`, `client_oriented`, `alternative`) will always still be present (they're compliant by integer enforcement).
- **No data loss** — the dropped scenarios were already flagged as non-compliant. We're just making the filtering decisive instead of letting users see a warning they might overlook.

## One-line PowerShell command for client-PC update

```
powershell -ExecutionPolicy Bypass -File C:\AMI_Optix_Agent\scripts\Refresh-AmiOptixAgentFromGitHub.ps1 -AgentRoot C:\AMI_Optix_Agent
```

(No VBA change; PS refresh is a no-op for this fix. Render auto-deploys; takes effect ~1-3 min after push.)
