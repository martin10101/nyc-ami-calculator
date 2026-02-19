# 2026 AMI Optix Release Notes

- Date: 2026-02-19
- Release branch: `release/2026-ami-optix-all-fixes`
- Repository: `martin10101/nyc-ami-calculator-all-fixes-test`
- Render deploy model: Auto-deploy OFF, manual deploy from Render UI
- Ready-to-deploy code commit (Render): `e5bda4d3c6dca92f5bcc4292a7214d6019eee7ef`

## Fixes Included

- `Fix-01` DataReader stability + no ribbon-driven sheet activation.
  - Prevents wrong-sheet and wrong AMI-column targeting drift during apply/write operations.
- `Fix-04` Ribbon stays active after actions.
  - Added reliable tab re-activation behavior and removed disruptive callback patterns.
- `Fix-03` Rent Roll Year selector + Manage uploads + mismatch warning.
  - Added explicit year selection/activation flow and non-blocking workbook year mismatch warning.
- `Fix-01b` Utilities variant breakdown labels.
  - Preserves exact utility variant labels in the top-of-sheet utilities output section.
- `Fix-02` Solver outcome de-duplication + placement tie-break.
  - Removes duplicate outcome-equivalent scenarios while preserving solver scoring behavior.
- `Fix-05` Manual Working Copy always-on + two-way sync.
  - Keeps manual sync active and stable without requiring a live-sync toggle.
- `Fix-06` Manual rents computed locally.
  - Manual edit refresh no longer calls `/api/manual_calculate` on each edit.
- `Fix-06b` Local rent calc hardening.
  - Added fail-fast checks and fingerprinting to block silent/partial/zeroed rent writes on bad source data.
- `Fix-06c` Table-driven local rent tables cache.
  - Introduced per-user CSV cache with Z: primary source and `%APPDATA%` fallback, plus status visibility.
- `Fix-06d` Verify Manual Rents (API) button.
  - Added one-click `/api/evaluate` verification with MATCH/MISMATCH summary and full diagnostics details.

## Key Behavior Changes

- Manual Working Copy rents are local/table-driven (`Fix-06c` path), not per-edit API-driven.
- Manual edits no longer trigger `/api/manual_calculate`; API verification is explicit/on-demand via **Verify Manual Rents (API)**.
- Solver scenario rents and solver outcome behavior remain API-driven and unchanged.

## Known Limitations

- Verification is on-demand, not automatic. It runs only when **Verify Manual Rents (API)** is clicked.
- Local rent refresh depends on year workbook availability/layout (`Z:\AMI_Optix\RentRollYears\<YEAR>\` or `%APPDATA%\AMI_Optix\RentRollYears\<YEAR>\`).
- If rent table source layout/labels drift, local refresh intentionally fails fast instead of writing partial values.

## Quick Validation Checklist

1. Open Excel workbook and confirm **AMI Optix** ribbon appears.
2. Click **AMI Optix -> API Settings** and confirm API key is set.
3. Click **AMI Optix -> Diagnostics** and confirm the diagnostics sheet opens.
4. In **Rent Roll Year**, select target year (for example `2025`).
5. Click **Refresh Rent Tables (Selected Year)**, then **Rent Tables Status**; confirm cache/source fields populate.
6. Edit one AMI value in Manual Working Copy and confirm rent values refresh locally.
7. Click **Verify Manual Rents (API)** and confirm MATCH/MISMATCH summary appears.
8. Confirm no per-edit `/api/manual_calculate` dependency is required for manual rent refresh.
