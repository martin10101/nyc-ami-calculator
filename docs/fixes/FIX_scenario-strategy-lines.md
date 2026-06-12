# FIX: "Why:" strategy lines + method legend

**Branch:** `fix/scenario-strategy-lines`
**Cut from:** `feature/excel-agent-foundation` @ `6205556`
**Date:** 2026-06-11
**Risk:** Very low — pure display text. Server adds a post-pipeline annotation pass (touches only `description`/`tradeoffs` text fields); VBA adds one row per scenario block and a 6-row legend, both inside writers already proven additive-safe.

## Why

Client direction: make the program's reasoning visible so a perceived discrepancy becomes a precise bug report instead of distrust — but structured, one-look readable, no info dump.

## What it looks like

Each scenario block opens with a computed, scenario-specific line:

> **Why:** *8 apartments at 40% - the minimum possible (10.78% of residential SF), using the largest units (avg 779 SF vs 604 SF pool avg); rent-maximized.*
> **Why:** *Same 8-apartment minimum at 40%, alternative band mix (40/70/80) for the neighborhood/clientele call. -$63/mo vs FEWEST 40 UNITS.*
> **Why:** *Maximum income: 9 apartments at 40% (11.67%) - pays with 1 apartment(s) above the minimum at 40%. +$190/mo vs FEWEST 40 UNITS.*

And under the overview table, a 4-line method legend ("HOW THESE OPTIONS ARE BUILT") stating the playbook: minimum apartments first (largest units), rent-max at that minimum, band variants for the clientele call, mid/max options as comparisons.

## Changes

### Server ([app.py](app.py))

Annotation pass after the Original-Scenario block, before response assembly: for every non-`original` scenario, computes 40% unit count, SF share of residential, average 40-tier unit size vs pool average, band list, and rent delta vs `fewest_40_units`, then writes a per-family `description`. Edge scenarios' real violation `tradeoffs` are untouched; the three 40-family keys stop duplicating their description into `tradeoffs` (the Why line replaces that). Wrapped in try/except — failure degrades to a note, never blocks the response.

### VBA ([excel-addin/src/AMI_Optix_ResultsWriter.bas](excel-addin/src/AMI_Optix_ResultsWriter.bas))

- `WriteScenarioSummaryAndTable`: renders `Why:` + italic description as the first row of every scenario block when `description` is a non-empty string (locally-built manual blocks have no description → line skipped gracefully).
- `WriteScenarioOverview`: appends the 4-line legend (font 9, italic) after the table. Legend lines start with digits — no leading-"=" formula risk; no collision with the scan patterns (SCENARIO n / GROUP: / UNIT).

## Surgical guarantees

- No solver, ordering, compliance, or assignment changes — text fields only.
- Ctrl+Z paths (EventHooks/AppEvents), View Scenario picker, apply, Manual Calculate, year switches: untouched code; they flow through the same shared writers and inherit the additions.
- Tests updated: strategy line now asserted in `description` (tradeoffs reserved for edge violations). Full suite 78 green. Static VBA checks clean (65/65 procs, no dup Dims, no leading-equals).

## Verified output (Building 1 v5, 2026)

All 8 solver scenarios carry computed Why lines with correct deltas (fewest = baseline; absolute_best +$190; low_40_share -$59; etc.); `original` keeps its "your saved assignments" line.

## Deploy

Server (Render auto) + VBA → standard `.xlam` refresh.
