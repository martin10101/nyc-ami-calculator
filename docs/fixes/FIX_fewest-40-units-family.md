# FIX: Fewest-40-units scenario family (the developer's lens)

**Branch:** `fix/fewest-40-units-family`
**Cut from:** `feature/excel-agent-foundation` @ `8315eba`
**Date:** 2026-06-11
**Risk:** Low — replaces the low-40 ladder block (same position after the edge section, purely additive); one VBA ordering line. All other scenarios untouched.

## Client feedback (Rachel Hiller, 2026-06-11)

> "The results does not have a single option with the least amount possible units at 40%... Owners always want to give the least units possible at 40% AMI... we would like to have most of the options, not only one, with the least units at 40% possible. It's a no brainer that when more units are at 40% AMI, the income is more, but this is not the point."

The conceptual gap: developers minimize the **NUMBER of apartments** at 40% AMI — each one is a deeply discounted unit — not the 40% **SF share**. Our low-40 ladder minimized SF share (10-10.5% window), which forces SMALL units at 40% and therefore MORE of them. Her approach: the LARGEST units at 40% — fewer apartments sacrificed, more weighted-average ballast per sacrificed unit, upper bands stay high, income stays near-max. (Her Building D 40% tier is literally the 8 largest pool units.)

Also: 60% vs 70% AMI has **no fixed rule** — it's a neighborhood/clientele call the owner makes — so the family must offer band variety at the minimum count rather than encode a preference.

## Changes

### 1. app.py — fewest-40-units family (replaces the low-40 ladder block)

Same position (after the edge block — purely additive, no displacement). Produces:

- **`fewest_40_units`** — true minimum unit count k* over the FULL legal window (greedy largest-first lower bound, probe upward via `find_max_revenue_scenario(low_band_unit_count=k)` until feasible), rent-maximized.
- **`fewest_40_units_2/_3`** — same k*, **different band mixes**: drop one upper band from the first option's mix at a time and re-solve. (Layout-only variants tie on rent and de-dupe away — band families are the real variety, exactly like the client's own A/B options: same 8 units at 40%, different upper bands.) Falls back to a differ-layout rung, then k*+1, only when needed.

All candidates: hard WAAMI cap check, canonical + outcome de-dupe vs every existing scenario, low-floor tie-break, client-readable description lines.

`low_40_share` (SF-minimal), `mid_40_share`, `max_40_share`, `absolute_best`, edge scenarios: untouched.

### 2. VBA — fewest family leads the sheet

`BuildScenarioKeyOrder` preferred order now starts with `fewest_40_units`, `_2`, `_3` before `absolute_best`. Scenario 1 = the option owners actually want; the rent-max-at-any-cost options follow.

## Building D verification (the client's own benchmark)

| Scenario | 40% units | Bands | Monthly rent |
|---|---|---|---|
| fewest_40_units | **8** | 40/60/90 | **$45,057** |
| fewest_40_units_2 | **8** | 40/60/80 | **$44,933 = her Option B to the dollar** |
| fewest_40_units_3 | **8** | 40/70/80 | **$44,801 = her Option A to the dollar** |

The family reproduces both of her hand-built options exactly and adds a same-count, better-rent one she didn't find (+$124/mo over her best). All pre-existing scenarios still generate identically (absolute_best $45,376, edge_waami_floor_590, low_40_share, mid_40_share, etc.).

## Tests

- Rewrote the API family test in [tests/test_low40_options.py](tests/test_low40_options.py): minimum count proven (6 for the synthetic building), ≥2 family options, counts ∈ {k*, k*+1}, outcome-distinct, descriptions present. Solver-level tests (count pin, differ_from, floor tie-break) and the mid-40 test unchanged.
- Full suite: **78 passed**.

## Deploy

- Render auto-deploys the server side.
- **VBA changed → standard `.xlam` refresh** on client PCs — rides along with the planned 3-PC install.
