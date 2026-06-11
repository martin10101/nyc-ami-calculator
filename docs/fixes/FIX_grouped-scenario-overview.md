# FIX: Grouped scenario overview + consistent Scenario 1

**Branch:** `fix/grouped-scenario-overview`
**Cut from:** `feature/excel-agent-foundation` @ `b68e4aa`
**Date:** 2026-06-11
**Risk:** Medium-low — display/ordering only, but touches the shared sheet-layout writers. No server changes.

## Why

1. After the fewest-40-units change, the manual working copy defaulted to `absolute_best` while Scenario 1 became `fewest_40_units` — they used to be the same scenario (the View Scenario picker cycles the manual block), so the mismatch confused users.
2. ~10 scenarios with no structure reads as clutter. User approved a "grouped sections" layout: an always-visible overview table at the top plus grouped detail blocks.

## Changes (all VBA)

### [AMI_Optix_ResultsWriter.bas](excel-addin/src/AMI_Optix_ResultsWriter.bas)

- **`BuildGroupedScenarioOrder`** (new, public): one ordering used everywhere — G1 FEWEST UNITS AT 40% (`fewest_40_units*`), G2 MID RANGE (`mid_40_share`), G3 MAX RENT / OTHER OPTIONS (everything else, existing preferred order), G4 YOUR INPUT (`original`) — with the canonical de-dupe built in, so the overview, the picker, and the numbered blocks always agree on what "Scenario 3" is.
- **`WriteScenarioOverview`** (new, public): compact index under the title/year line — group banners + one row per scenario (`> #` marker / name / bands / "8 @ 10.00%" / monthly rent). Reads `g_LastScenarios`, so all four manual-block writers can re-create it after clearing the top; column N stores each row's key. New global `g_AMIOptixCurrentScenarioKey` drives the `>` marker.
- **All four manual-block writers** (optimize result, evaluate/Manual Calculate, local refresh, apply-scenario) now write the overview after the Rent Roll Year line. The apply-scenario writer also **gains the year line it was missing** and sets the current-key marker, so picking a scenario moves the `>`.
- **`CreateScenariosSheet`**: detail blocks iterate the grouped order with banner rows (`=== FEWEST UNITS AT 40% ===`) between groups; numbering = grouped position.
- **`GetBestScenarioKey`**: prefers `fewest_40_units` — the working copy defaults to Scenario 1 again.
- **`ClearManualBlock`** clears through column N (the key helper column); **`FindFirstScenarioHeaderRow`** treats `===` banners as the start of the scenario area so the clear never eats the first banner.

### [AMI_Optix_Ribbon.bas](excel-addin/src/AMI_Optix_Ribbon.bas)

- **`ShowScenarioList`** (the View Scenario picker) lists scenarios in the same grouped order with group separators — picker numbers now match the sheet's scenario numbers (previously strict-then-edge, which could disagree).

## Known one-time cosmetic edge

If a workbook's results sheet was built by an older add-in (no overview), the first Manual Calculate after upgrading writes a taller top block than was cleared and can overlap the old first banner area. Self-heals on the next Find Optimal Scenarios run.

## Verification

- Static checks: proc/EndProc balance, no duplicate Dims, For/Next + If/End If balance in new functions.
- pytest: 78 passed (no server changes).
- Manual on Building D after `.xlam` refresh: overview with 4 group banners; `>` on FEWEST 40 UNITS; manual header "CURRENT: FEWEST 40 UNITS" = Scenario 1; picking scenario 5 via View Scenario updates the manual block AND moves the `>`; picker numbers = sheet numbers; Manual Calculate + year switch keep the overview.

## Deploy

VBA-only → standard `.xlam` refresh (bundle with the 3-PC install).
