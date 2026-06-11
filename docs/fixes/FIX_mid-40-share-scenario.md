# FIX: MID 40 SHARE scenario (fills the 10.5-11.5% gap)

**Branch:** `fix/mid-40-share-scenario`
**Cut from:** `feature/excel-agent-foundation` @ `98bd837`
**Date:** 2026-06-11
**Risk:** Very low — one additive scenario via the existing `_solve_with_40_window` helper; no existing scenario's code path touched. Server-side only, no client-PC refresh.

## Client request

The 40% AMI coverage had a hole: the LOW 40 SHARE group hugs the floor ([10%, 10.5%]), the unconstrained ABSOLUTE BEST often runs toward the ceiling (~12.25% on Building D), and MAX 40 SHARE pins [12%, 12.5%]. Nothing explored 10.5-11.5%. The client wants at least one option in that middle range.

## Change ([app.py](app.py))

After the `max_40_share` block in the 40%-variants section, add `mid_40_share`: rent-maximized with the 40% window pinned to `[floor + 0.5%, min(ceiling, floor + 1.5%)]` (standard MIH: [10.5%, 11.5%]) using the same battle-tested `_solve_with_40_window` helper, with the low-floor tie-break on. De-duped against every existing scenario's canonical assignments; carries a description line ("Mid-range option: 40% AMI between 10.5% and 11.5% of residential SF, rent-maximized") rendered by the existing tradeoffs display.

Excel shows it automatically as **MID 40 SHARE** (dynamic key ordering + auto display names + band suffix + per-scenario compliance lines all apply with zero VBA changes).

## Verified

- New test in [tests/test_low40_options.py](tests/test_low40_options.py): `mid_40_share` exists on the synthetic ladder building, its 40-band share lands inside [10.5%, 11.5%], its layout is canonically distinct, and the description line is present.
- Full suite: **78 passed**.
- Building D probe (2026 rents): MID 40 SHARE = 11.01% share, $45,118/mo, bands 40/70/90 — the coverage curve now reads 10.00% ($45,057) → 10.45% ($45,246) → 11.01% ($45,118) → 12.25% ($45,376).

## Deploy

Render auto-deploys; nothing to install on client PCs.

## Revision 2026-06-11 (branch `fix/mid-40-after-edge-block`)

User feedback after the first deploy: an edge scenario (`edge_waami_floor_590`,
$45,376 on Building D) disappeared and the low-40 ladder could lose rungs.
Cause: the mid-40 block originally ran in the 40%-variants section, BEFORE
the edge block — so it consumed an edge-budget slot (`target_edge_count =
6 - len(scenarios)`) and polluted the ladder's duplicate filter. Exactly the
interference the low-40 ladder's after-the-edge-block placement was designed
to avoid.

Moved: mid_40_share now generates AFTER the edge block and the low-40
ladder (right before Fix-02 de-dupe), self-contained (recomputes its own
window, dedupes by canonical + outcome signature against everything).
Verified on Building D: the full pre-mid scenario set returns byte-identical
(edge_waami_floor_590 and the $45,247 alternative restored, all 3 low-40s),
with MID 40 SHARE purely additive — 9 scenarios total.
