# FIX: Compliance box ("required vs provided") + band-mix scenario titles

**Branch:** `fix/compliance-language-display`
**Cut from:** `feature/excel-agent-foundation` @ `657271c`
**Date:** 2026-06-11
**Risk:** Very low — display-only. One additive server field, two VBA display tweaks. No solver, rent, or constraint logic touched.

## Why

Analysis of the client's own allocation workbooks (Building D, 2026-06-10) showed how she validates and names her work:

- She checks compliance as **"required vs provided"** (her sheet has a box: required 40% SF, provided 40% SF, surplus) — not as bare percentages.
- She names options by **band family** ("Option A - 40, 70, & 80").

Aligning the program's output with her checklist format makes results readable as *her* compliance sheet instead of a black box.

## Changes

### 1. Server — effective 40% window in `project_summary` ([app.py](app.py))

For MIH requests, `project_summary` now carries `mih_low_band_min_share` / `mih_low_band_max_share` read from the post-floor-walk config — the window the solver actually enforced. Additive field; old add-ins ignore it.

### 2. VBA — compliance box in the SQUARE FOOTAGE SUMMARY ([AMI_Optix_ResultsWriter.bas](excel-addin/src/AMI_Optix_ResultsWriter.bas))

`WriteMihSquareFootageSummary` now appends, under "Affordable Net SF":

> **40% AMI Floor:** required 5,741.42 SF | provided 5,764.21 SF | surplus +22.79 SF
> **Affordable Share:** provided 25.17% of residential SF

- "required" = effective min share × residential SF (server-reported window; falls back to the standard 10% when talking to an older server).
- A shortfall renders as red bold "SHORTFALL n SF" — impossible to miss.
- MIH-only (requires a real building denominator); UAP output unchanged.

### 3. VBA — band mix in scenario titles ([AMI_Optix_ResultsWriter.bas](excel-addin/src/AMI_Optix_ResultsWriter.bas))

Scenario headers now read:

> SCENARIO 1: ABSOLUTE BEST - 40/60/90
> SCENARIO 2: LOW 40 SHARE 2 - 40/60/90

via a new `FormatBandsSuffix` helper (tolerates integer and fractional band values; silently omits the suffix when bands are unavailable). Plain hyphen, no special characters (VBA codepage safety).

## Verified

- New [tests/test_project_summary_compliance.py](tests/test_project_summary_compliance.py): MIH response carries the window (0.10/0.125); UAP response has no MIH window fields.
- Full suite: **77 passed** (75 + 2 new).
- VBA changes are write-only formatting in existing, exercised code paths.

## Deploy

- Render auto-deploys the server field.
- Client PCs need the standard `.xlam` PowerShell refresh (VBA changed).

## Out of scope (deliberately)

- Band-preference rules (e.g. "avoid 60% AMI") — the client's own options contradict a blanket rule; awaiting her answer to the band-choice question before encoding anything.

## Revision 2026-06-11 (branch `fix/compliance-box-per-scenario`)

User feedback after seeing it live: the compliance lines sat in the top
SQUARE FOOTAGE SUMMARY, where the "provided" 40% SF (a **per-scenario**
value — every scenario allocates a different 40% SF) read as a building
fact. Wrong altitude.

Moved: the "40% AMI Floor" + "Affordable Share" lines now render inside
**every scenario block**, directly under its Band Mix table, via a new
`WriteMihComplianceLines` helper called from `WriteScenarioSummaryAndTable`
(which serves both the manual live block and all numbered scenarios — one
call site covers everything). The top summary reverted to building-level
facts only (Total Building Net SF, Affordable Net SF). Same formatting,
same red SHORTFALL treatment, same server fields.
