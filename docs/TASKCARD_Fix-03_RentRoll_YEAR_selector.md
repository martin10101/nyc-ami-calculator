# Task Card — Fix-03 Rent Roll YEAR selector + Manage uploads + mismatch warning (MIH + UAP)

## Goal
- Add a **Rent Roll Year** selector (2022–2026, default **2025**) in the AMI Optix ribbon.
- Provide an in-Excel **Manage Rent Roll Years…** action to upload/replace year calculator files.
- Ensure all MIH/UAP workflows that compute rents/utilities use the **selected year**.

## Acceptance Criteria (from `docs/FIX_REQUIREMENTS.md`)
- Add ribbon year selector: 2022, 2023, 2024, 2025, 2026 (default 2025).
- Add a “Manage Rent Roll Years…” UI (form preferred) to upload/replace year files.
- Uploaded year files must persist locally (suggested: `%APPDATA%\\AMI_Optix\\RentRollYears\\<year>\\`).
- Upload action should also transfer the file once to the API for storage.
- Selected year must override any year implied inside the active workbook’s rent roll page.
- If workbook “declared year” != selected year, show warning (OK to continue; not blocking).
- Every MIH/UAP action that uses rent roll/utilities must use the selected year.

## Success Criteria
- User can select a year (defaults to 2025) and the server uses that year’s rent calculator for subsequent optimize/evaluate/manual-calculate calls.
- “Manage Rent Roll Years…” lets user pick a rent calculator workbook and:
  - copies it into `%APPDATA%\\AMI_Optix\\RentRollYears\\<year>\\`
  - uploads it to the API’s rent-calculator storage (overwrite supported)
  - activates it on the API so computations use that year immediately
- If the active workbook appears to declare a different year than the selection, a warning is shown and the user can proceed.

## Files to Change
- `excel-addin/customUI/customUI14.xml`
- `excel-addin/src/AMI_Optix_Ribbon.bas`
- `excel-addin/src/AMI_Optix_API.bas`
- `docs/CODEX_LEDGER.md` (work log update)

## Functions / Entrypoints
- RibbonX:
  - Year dropdown callbacks: `Ribbon_GetRentRollYear*`, `Ribbon_SelectRentRollYear`
  - Manage button: `Ribbon_ManageRentRollYears`
- API calls (Excel → Render):
  - `POST /api/rent-calculators/upload` (multipart form-data)
  - `POST /api/rent-calculators/activate` (JSON)
  - `CallOptimizeAPI`, `CallEvaluateAPI`, `CallManualCalculateAPI` (ensure selected year activated)

## Proposed Patch (logic-level)
1) **Ribbon UI**
   - Add a “Rent Roll Year” dropdown with fixed options 2022–2026 (default 2025).
   - Add a “Manage Rent Roll Years…” button in the Rent Roll group.

2) **Local persistence**
   - Store uploaded calculator files under `%APPDATA%\\AMI_Optix\\RentRollYears\\<year>\\`.
   - Store selected year (and remote filename used) in the registry via `SaveSetting` so it survives Excel restarts.

3) **API storage + activation**
   - On “Manage…” upload: POST the chosen file to `/api/rent-calculators/upload` (with `overwrite=true`), then activate it.
   - On year selection change: activate the matching year calculator on the API (or activate default for 2025).
   - Before any optimize/evaluate/manual-calculate request: ensure the selected year is activated (so Live Sync / worksheet-driven calls also use the selected year).

4) **Mismatch warning**
   - Best-effort detect a “declared year” (from workbook name / rent roll sheet name / top-of-sheet scan).
   - If it differs from selected year, show a warning that does not block continuing.

## Risk Pre-mortem
- **Risk:** API rent calculator activation is server-global; switching years affects other concurrent users.
  - **Mitigation:** Keep behavior explicit (only switches on user year selection / Manage upload); document as known limitation.
- **Risk:** Declared-year detection may miss or mis-detect in some workbooks.
  - **Mitigation:** Best-effort only; warning shown only when a clear year is detected.
- **Risk:** Multipart upload code could fail on large files or locked paths.
  - **Mitigation:** Clear error messaging; local copy step first; use conservative timeouts.

## Test Plan
- Manual (Excel):
  1. Open a workbook with AMI Optix add-in installed.
  2. Use **Manage Rent Roll Years…**:
     - Select year **2022** and upload a valid 2022 calculator workbook.
     - Repeat for **2024**.
  3. Select **2022** in the year dropdown.
  4. Run **Run UAP** (or **Manual Calculate**) and verify outputs reflect 2022 rent tables.
  5. With a workbook that appears to declare **2025**, select **2022** → warning appears; clicking OK continues.

- Automated (repo sanity):
  - `python -m pytest -q`

## Rollback Plan
- Revert the fix commit(s) and rebuild the `.xlam` with prior ribbon/module versions.
