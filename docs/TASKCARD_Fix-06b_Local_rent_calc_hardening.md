# Task Card — Fix-06b: Harden Local Manual Rent Calc (Fingerprint + Fail-Fast Errors)

Date: 2026-02-18  
Repo: martin10101/nyc-ami-calculator-all-fixes-test  
Base branch: main  
Work branch: fix/06-local-manual-rent-calc

## Goal
- Keep API-provided rents/totals for solver scenarios unchanged.
- Harden Fix-06 local manual rent calculation so it **fails fast** when the rent workbook layout is unexpected.
- Treat missing gross/allowance lookups as **hard errors** (no silent 0), blocking rent writes and showing a clear message.

## Success Criteria
- Manual edits do **not** call `/api/manual_calculate` (unchanged from Fix-06).
- If the selected year workbook is present but layout is unexpected:
  - Local manual rent calc fails fast and shows a clear message including year + workbook path + fingerprint.
  - Manual block still refreshes AMI values, but rents/totals are not written (blank) to avoid inaccurate data.
- If any gross rent lookup is missing:
  - Hard error message includes: year workbook path, unit_id, AMI%, bedroom label, requested key, and indicates gross-table lookup.
  - No rents/totals are written.
- If any utility allowance lookup is missing for a non-“N/A or owner pays” selection:
  - Hard error message includes: year workbook path, unit_id, category, option label, bedroom label, requested key, and indicates allowance lookup.
  - No rents/totals are written.
- Solver scenarios (Optimize output) still include API-provided rents/totals and the scenario grid remains unchanged.

## Files to change (keep minimal)
- `excel-addin/src/AMI_Optix_ResultsWriter.bas`
- `docs/TASKCARD_Fix-06b_Local_rent_calc_hardening.md`
- `docs/CODEX_LEDGER.md`

## Functions / Entrypoints to change
- `AMI_Optix_ResultsWriter.bas`
  - `RefreshManualWorkingCopyLocalRents(...)` — ensure failures don’t partially write rents.
  - `EnsureLocalRentScheduleReady(...)` — add fingerprint validation.
  - Local lookup helpers — convert “missing -> 0” into “missing -> hard error” for required lookups.

## Proposed Patch (logic-level)
- Add a **workbook layout fingerprint** computed from:
  - Presence of expected allowance headers in row 15 (electricity/cooking/heat/hot water),
  - Presence of an “of AMI” marker in column D with a numeric AMI in column C,
  - Detection of bedroom labels in the gross rent table following an “of AMI” marker.
- Validate fingerprint before parsing and fail fast if missing anchors.
- During rent enrichment:
  - Missing gross rent lookup raises a hard error with unit_id + key + workbook path.
  - Missing allowance lookup raises a hard error for non-“N/A or owner pays” option labels.
  - On any failure, clear per-assignment rent fields so the manual table does not show partial results.
- Show a clear error message (and also add a Tradeoffs line) so users can self-diagnose workbook issues.

## Risk pre-mortem + mitigations
- **Noisy errors (MsgBox on each edit)**: mitigate by de-duping repeated identical errors per session.
- **Workbooks with minor formatting differences** could start failing even if they’re “close enough”:
  - Intentional: fail-fast is preferred over silently wrong values.

## Test Plan
### Automated
- `python -m pytest -q`
- Edge scenario count sanity:
  - `python -c "from app import app; c=app.test_client(); p={'program':'UAP','utilities':{'electricity':'na','cooking':'na','heat':'na','hot_water':'na'},'units':[{'unit_id':'L1','bedrooms':1,'net_sf':200,'floor':1,'balcony':False},{'unit_id':'L2','bedrooms':1,'net_sf':200,'floor':1,'balcony':False},{'unit_id':'H1','bedrooms':2,'net_sf':400,'floor':6,'balcony':True},{'unit_id':'H2','bedrooms':2,'net_sf':400,'floor':6,'balcony':True},{'unit_id':'M1','bedrooms':2,'net_sf':400,'floor':3,'balcony':False},{'unit_id':'M2','bedrooms':2,'net_sf':400,'floor':3,'balcony':False}]}; r=c.post('/api/optimize',json=p); d=r.get_json(); print('status',r.status_code,'scenarios',len((d or {}).get('scenarios') or {}))\"`

### Manual (Excel)
1) Manual edit on AMI Scenarios manual table:
   - No `/api/manual_calculate` hits.
2) Switch year 2024 ↔ 2025 and re-edit:
   - Rents change if schedules differ and workbook path used is correct (Z:\ first).
3) Force a missing lookup (e.g., temporarily rename a utility option label in the year workbook copy):
   - Clear error message; rents/totals not written.

## Rollback Plan
- Revert the Fix-06b commit(s) to restore prior “best-effort” local calc behavior.

