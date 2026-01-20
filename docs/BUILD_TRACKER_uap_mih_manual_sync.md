# Build Tracker: UAP/MIH + Manual Sync + Scenarios

Branch: `feature/uap-mih-manual-sync`  
Rule: do **not** push to `main` until approved.

## Goals (from client meeting)
- Separate programs: UAP and MIH run paths must not mix rules.
- MIH residential SF denominator: `MIH!J21` (error if blank).
- Utilities: auto-read from the client workbook “matrix” selections when present; fall back to stored settings.
- Scenarios: show at least 3 scenarios when feasible.
- Add extra scenario: **Max Revenue** with WAAMI floor **58.9%**, still respecting share + caps; never auto-applied.
- Manual scenario: bi-directional sync between UAP AMI column and “Scenario Manual” on `AMI Scenarios`, with validation (invalid edits → error + revert).
- Results sheet: remove “Gross Monthly Rent” line and remove “Total Utility Allowances” total line; keep individual allowance components.
- Scenario application UX: simple “pick scenario number → apply”.
- Add unit size fields in scenarios: Bedrooms + Net SF + Floor + Balcony.
- Log “Final Selection” + “Reason/Notes” for future learning.

## Status Checklist

### Backend (Python)
- [x] Add `program` field to API request payload (`UAP` / `MIH`)
- [x] Add MIH mode config + constraints (Option 1 / Option 4)
- [x] Read MIH option from workbook-derived payload (from Excel add-in) or future file upload path
- [x] Implement **Max Revenue** scenario (WAAMI floor 58.9%) using rent-based objective
- [x] Add `/api/evaluate` endpoint to validate a manual assignment + compute rents/totals
- [x] Ensure scenario output includes: bands, assignments, metrics, rent_totals, notes

### Excel Add-in (VBA + Ribbon)
- [x] Ribbon: add separate buttons `Run UAP` and `Run MIH`
- [x] Utilities: detect + read selections from workbook matrices (UAP: `Calculations!P3:AA3`; MIH: `Rents & Utilities` matrix)
- [x] Payload: send `program`, MIH option, and MIH residential SF (`MIH!J21`)
- [ ] `AMI Scenarios` formatting updates (remove gross line + total allowance line; add individual breakdown lines)
- [ ] Include unit size columns: Bedrooms + Net SF + Floor + Balcony
- [ ] Scenario apply flow: prompt for scenario # and apply to UAP AMI column
- [x] Manual Scenario sync (UAP ↔ AMI Scenarios), “most recent edit wins”, invalid edits revert with error
- [ ] Add “Final Selection” dropdown + “Reason/Notes” box and append to a history table

### Testing
- [ ] Add unit tests for MIH constraints + program separation
- [ ] Add tests for Max Revenue scenario selection behavior
- [x] Run `pytest` locally and confirm green
- [ ] Manual QA on provided workbooks in `C:\Users\MLFLL\Downloads\files`

## Notes / Decisions (locked)
- Auto-applied scenario remains “best WAAMI” (current behavior).
- “Max Revenue” (58.9% floor) is an **extra option** only (never auto-applied).
- “Most recent edit wins” for Manual Scenario sync.
- Manual edits must validate; invalid edits show an error (do not silently accept).
