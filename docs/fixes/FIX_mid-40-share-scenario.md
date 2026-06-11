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
