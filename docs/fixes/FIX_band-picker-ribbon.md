# FIX: Band picker - ribbon "AMI Bands" menu (Fix 4 of the 2026-09-02 client feedback)

**Branch:** `fix/band-picker-ribbon`
**Date:** 2026-09-22
**Risk:** Moderate, additive. Server + VBA + ribbon XML (first ribbon change
shipped through the deploy script). No EventHooks / manual-block / Ctrl+Z
changes. Absent payload field = byte-identical legacy behavior.

## What the client asked for

Building D options 5-6 used a 60% band the owner refuses. She wants to
choose which AMI levels the program may use (60, 70, 80...), per building,
from the ribbon - not from cells.

## What shipped

**Ribbon** (`excel-addin/customUI/customUI14.xml`): new group "Bands & Floors"
with a `dynamicMenu` "AMI Bands" (`invalidateContentOnDrop="true"`, so the
menu is rebuilt by VBA every time it opens - never stale after an
Option 1 <-> Option 4 flip or a workbook switch). Nothing else in the ribbon
moved.

**VBA** (new `AMI_Optix_Bands.bas`, callbacks in `AMI_Optix_Ribbon.bas`):

- Menu per program/option: UAP and MIH Option 1 show 40 (locked, "required")
  60 70 80 90 100; MIH Option 4 shows 40-135 (no 40 lock, but at least one
  band <= 70 must stay on). A lower `Prog!I4` cap shortens the list, like
  the server.
- Rules enforced on every click: 40 cannot be unchecked where required;
  Option 4 must keep a band <= 70; at least 2 bands must stay checked. A
  refused click shows why and the box re-renders unchanged.
- Storage: one hidden defined name in the DATA workbook
  `AMI_Optix_AllowedBands = "<program/option>|<csv>"`. Checking everything
  stores nothing (= program default, payload field not sent).
- Run preflight (`AMI_Optix_Main.RunOptimizationForProgram`, Step 3D):
  re-reads program + option for THIS run; a selection stored for another
  option, or one that became unusable, STOPS the run with a clear message.
  Never silently narrows or widens.
- Payload (`BuildAPIPayloadV2`): `allowed_bands: [40, 70, 80]` only when
  narrowed.
- Results header (`WriteBandRulesLine`): "AMI Bands Allowed: 40%, 70%, 80%
  (your selection ...)" only when narrowed; survives Manual Calculate and
  year changes (falls back to the last optimize response).

**Server** (`app.py`): `_normalize_allowed_bands` + `_apply_allowed_bands`
inside `_build_program_config`. The list is intersected with the option's
own candidate bands AFTER the client caps (100 for Option 1, 135 for
Option 4), so it can only narrow. Unusable selections return
`success: false` with a plain-English `error` (shown verbatim in Excel).
Echo: `project_summary.band_rules = {source, allowed_bands, requested_bands,
ignored_bands}` and a note "Built with your band rules: ...". Every solver
call in the request reads `potential_bands` from this config, so the
40-window ladders, fewest-40 ladder and edge scenarios all obey the picker.

## Three locks

1. Ribbon/VBA validation (this module) - refuses illegal toggles, stops stale runs.
2. Server intersection - cannot widen past program/option caps; clean errors.
3. Solver band domain - built from the narrowed `potential_bands`.

## Verification

- `tests/test_band_picker.py` (21 tests): normalization; identity when the
  field is absent (UAP, O1, O4); "every band checked" == default; O1/UAP
  exclusion of 60; caps still apply; refusals (no 40 on O1/UAP, no band <= 70
  on O4, single band, all above cap); end-to-end Building D shape - no 60%
  in any optimized scenario, echo + note present, unusable selection is a
  clean error, absent field echoes default.
- Full suite green (see commit message for the count).
- Sandbox Excel QA (before any client PC): open AMI Bands on a MIH O1 file
  (40 locked, 60-100 listed), uncheck 60, run, header shows the line, no 60%
  anywhere; flip Prog!K4 to Option 4, open the menu (shows "previous
  selection was for MIH Option 1"), Run MIH -> stopped with the message;
  Allow all bands -> runs; UAP file shows 40-100; Ctrl+Z and Manual
  Calculate unaffected.

## Deploy

`tools/excel-agent/Deploy-AmiOptixFixes.ps1` now also swaps
`AMI_Optix_Bands` + `AMI_Optix_Ribbon` and replaces `customUI/customUI14.xml`
inside the .xlam (zip entry) after the COM module swap, verifying the XML
parses and carries the `mnuBands` marker before promoting. Rollback = re-pin
the previous commit (the backed-up master `.bak` also restores in one copy).
