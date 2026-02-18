# Task Card — Fix-05: Manual Working Copy always-on + 2-way sync

Date: 2026-02-18  
Repo: martin10101/nyc-ami-calculator  
Base branch: perf/api-optimize-speed-2026-02-05  
Work branch: fix/05-manual-working-copy-sync

## Goal
- Remove the manual scenario ON/OFF (live sync) toggle; manual working copy is always present/editable.
- Keep solver scenarios (1..N) immutable; selecting scenario #k copies it into Manual Working Copy and applies to MIH/UAP without mutating the snapshot.
- Allow Manual Working Copy edits even when rule-invalid (warn but allow override).
- Ensure 2-way sync between Manual Working Copy and MIH/UAP AMI column, and refresh rent roll calcs on either edit.

## Success Criteria (docs/FIX_REQUIREMENTS.md)
- Selecting scenario #2 updates manual working copy; scenario #2 snapshot remains unchanged.
- Editing manual unit AMI updates MIH/UAP AMI instantly and refreshes rent roll.
- Editing MIH/UAP AMI cell updates manual working copy and refreshes rent roll.
- No infinite change-event loops; behavior remains stable after switching sheets and changing rent roll year (if present).

## Files to change (keep minimal)
- `excel-addin/customUI/customUI14.xml`
- `excel-addin/src/AMI_Optix_EventHooks.bas`
- `excel-addin/src/AMI_Optix_AppEvents.cls`

## Functions / Entrypoints to change
- Ribbon UI: remove references to `Ribbon_ToggleLiveSync` / `Ribbon_GetLiveSync` by deleting the toggle control from XML.
- `excel-addin/src/AMI_Optix_EventHooks.bas`
  - `GetLiveSyncEnabled()`
  - `SetLiveSyncEnabled(enabled As Boolean)`
- `excel-addin/src/AMI_Optix_AppEvents.cls`
  - `App_SheetChange(...)`
  - Manual working copy handlers (single-cell + range edits)
  - MIH/UAP AMI edit handler

## Proposed Patch (logic-level)
- **Always-on manual working copy**
  - Remove the ribbon toggle control from `customUI14.xml`.
  - Make `GetLiveSyncEnabled()` always return `True` and make `SetLiveSyncEnabled(...)` a no-op (or keep state but never disable).
- **Relaxed manual edits (warn, don’t revert)**
  - In `AMI_Optix_AppEvents.cls`, replace the “evaluate + revert on invalid” behavior for Manual Working Copy edits with a “manual_calculate + refresh” path:
    - Apply the edited AMI value to the corresponding MIH/UAP AMI cell (and vice versa).
    - Call `/api/manual_calculate` (quietly) to compute rent totals + diagnostics without enforcing constraints.
    - Write diagnostics/tradeoffs into the existing manual/tradeoffs area (no modal MsgBox; no revert).
- **Re-entrancy guard**
  - Ensure all writes triggered by SheetChange run under `g_AMIOptixSuppressEvents` + `Application.EnableEvents = False` to prevent loops.
- **Must NOT touch**
  - Scenario snapshot grid (solver scenarios 1..N), solver logic, or scenario keys.

## Risk pre-mortem + mitigations
- **Infinite SheetChange loops** if we write back into the edited range.
  - Mitigation: strict guard + write suppression; early-exit if `g_AMIOptixSuppressEvents` or `Application.EnableEvents=False`.
- **Wrong sheet/AMI column** targeted during sync.
  - Mitigation: use existing anchored “data sheet” helpers (no `ActiveSheet` assumptions) and keep all lookups unchanged.
- **Performance regression** from calling API on every edit.
  - Mitigation: only call manual_calculate for edits inside manual working copy or AMI column; avoid calling for unrelated edits.
- **Accidental scenario snapshot mutation**.
  - Mitigation: only write to Manual Working Copy region and MIH/UAP AMI column; never write into scenario snapshot grid.

## Test Plan
### Automated (CLI)
- `python -m pytest -q`
- Edge scenario count sanity (API should still return up to target scenarios; record count):
  - `python -c "from app import app; c=app.test_client(); p={'program':'UAP','utilities':{'electricity':'na','cooking':'na','heat':'na','hot_water':'na'},'units':[{'unit_id':'L1','bedrooms':1,'net_sf':200,'floor':1,'balcony':False},{'unit_id':'L2','bedrooms':1,'net_sf':200,'floor':1,'balcony':False},{'unit_id':'H1','bedrooms':2,'net_sf':400,'floor':6,'balcony':True},{'unit_id':'H2','bedrooms':2,'net_sf':400,'floor':6,'balcony':True},{'unit_id':'M1','bedrooms':2,'net_sf':400,'floor':3,'balcony':False},{'unit_id':'M2','bedrooms':2,'net_sf':400,'floor':3,'balcony':False}]}; r=c.post('/api/optimize',json=p); d=r.get_json(); print('status',r.status_code,'scenarios',len((d or {}).get('scenarios') or {}))\"`

### Manual (Excel)
- Select scenario #2 → Manual Working Copy updates; scenario #2 snapshot unchanged.
- Edit a Manual Working Copy AMI value to an invalid mix:
  - It stays (no revert), but a warning/tradeoffs are shown somewhere in the manual/tradeoffs area.
  - MIH/UAP AMI column updates immediately; rent roll calcs refresh.
- Edit MIH/UAP AMI column cell → Manual Working Copy updates immediately; rent roll calcs refresh.
- Repeat after switching worksheets and changing selected rent roll year (if available).

## Rollback Plan
- `git revert <fix_commit_sha>` (or reset branch) to restore:
  - live sync toggle UI + persisted enabled/disabled state
  - strict evaluate+revert behavior for manual edits

