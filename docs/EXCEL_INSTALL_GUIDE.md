# AMI Optix Excel Add-in Install / Update Guide (Final Release)

This guide covers both a fresh build and an in-place upgrade for the final combined fixes release.

- Repo: `martin10101/nyc-ami-calculator-all-fixes-test`
- Branch: `release/2026-ami-optix-all-fixes`
- Files used:
  - VBA modules: `excel-addin/src/*.bas`
  - VBA class modules: `excel-addin/src/*.cls`
  - VBA forms: `excel-addin/forms/*.frm`
  - Ribbon XML: `excel-addin/customUI/customUI14.xml`

## Source Files Used (Exact Paths)

- Modules (`.bas`) from `excel-addin/src/`:
  - `AMI_Optix_API.bas`
  - `AMI_Optix_DataReader.bas`
  - `AMI_Optix_Diagnostic.bas`
  - `AMI_Optix_Diagnostics.bas`
  - `AMI_Optix_EventHooks.bas`
  - `AMI_Optix_Learning.bas`
  - `AMI_Optix_Main.bas`
  - `AMI_Optix_RentCalcTables.bas`
  - `AMI_Optix_RentTables.bas`
  - `AMI_Optix_ResultsWriter.bas`
  - `AMI_Optix_Ribbon.bas`
  - `AMI_Optix_Setup.bas`
  - `AMI_Optix_VerifyManualRents.bas`
- Class module (`.cls`) from `excel-addin/src/`:
  - `AMI_Optix_AppEvents.cls`
- Form (`.frm`) from `excel-addin/forms/`:
  - `frmUtilities.frm`
- Ribbon XML:
  - `excel-addin/customUI/customUI14.xml`

## PATH A (Recommended): Build Fresh `AMI_Optix.xlam` From Source

1. Close Excel completely.
   - End all `EXCEL.EXE` processes before starting.
2. Open Excel and create a new blank workbook.
3. Press `Alt + F11` to open the VBA editor.
4. Remove old components if present.
   - In the target VBA project, remove these modules:
     - `AMI_Optix_API`
     - `AMI_Optix_DataReader`
     - `AMI_Optix_Diagnostic`
     - `AMI_Optix_Diagnostics`
     - `AMI_Optix_EventHooks`
     - `AMI_Optix_Learning`
     - `AMI_Optix_Main`
     - `AMI_Optix_RentCalcTables`
     - `AMI_Optix_RentTables`
     - `AMI_Optix_ResultsWriter`
     - `AMI_Optix_Ribbon`
     - `AMI_Optix_Setup`
     - `AMI_Optix_VerifyManualRents`
   - Remove class module:
     - `AMI_Optix_AppEvents`
   - Remove form:
     - `frmUtilities`
5. Import modules from `excel-addin/src/`.
   - VBA editor -> `File -> Import File...`, then import each `.bas` listed above.
6. Import class module from `excel-addin/src/`.
   - Import `AMI_Optix_AppEvents.cls`.
7. Import forms from `excel-addin/forms/`.
   - Import `frmUtilities.frm`.
   - If VBA asks for missing `frmUtilities.frx`, use PATH B (upgrade existing add-in with existing form assets) or recreate form controls from `excel-addin/forms/frmUtilities_DESIGN.txt`.
8. Compile.
   - VBA editor -> `Debug -> Compile VBAProject`.
9. Save as add-in.
   - Excel -> `File -> Save As`.
   - Type: `Excel Add-in (*.xlam)`.
   - Name: `AMI_Optix.xlam`.
   - Save into your Excel AddIns folder (typically `%APPDATA%\Microsoft\AddIns\`).
10. Embed Ribbon XML.
    - Close Excel completely.
    - Open `AMI_Optix.xlam` in OfficeRibbonXEditor.
    - If `customUI14.xml` is missing: insert `Office 2010+ Custom UI Part`.
    - Paste XML from `excel-addin/customUI/customUI14.xml`.
    - Save and close OfficeRibbonXEditor.
11. Enable add-in in Excel and restart.
    - Excel -> `File -> Options -> Add-ins`.
    - `Manage: Excel Add-ins -> Go... -> Browse...`.
    - Select `AMI_Optix.xlam`, check it, click `OK`.
    - Restart Excel.
12. Set API Settings and confirm diagnostics.
    - Ribbon -> `AMI Optix -> API Settings`, enter API key.
    - Run one API action (`Run UAP` or `Run MIH`) with a valid workbook.
    - Click `AMI Optix -> Diagnostics`.
    - Confirm diagnostics sheet includes `API Base URL` and populated `Last API Scenarios`.
13. Set Rent Roll Year and build cache.
    - Ribbon -> `AMI Optix -> Rent Roll Year` dropdown, pick year.
    - Click `Refresh Rent Tables (Selected Year)`.
    - Click `Rent Tables Status`.
    - In `AMI Optix Diagnostics`, verify `Rent Tables Status` rows are populated.
14. Verify manual rents.
    - Open workbook with Manual Working Copy.
    - Make one AMI edit in Manual Working Copy.
    - Click `AMI Optix -> Verify Manual Rents (API)`.
    - Confirm summary dialog shows `MATCH` or `MISMATCH`, then review diagnostics mismatch table if needed.

## PATH B: Upgrade Existing `AMI_Optix.xlam` (Replace Components In Place)

1. Close Excel completely.
2. Open existing `AMI_Optix.xlam` in Excel.
3. Press `Alt + F11`.
4. Remove and re-import components.
   - Remove and re-import all modules listed in `excel-addin/src/*.bas`.
   - Remove and re-import class module `AMI_Optix_AppEvents.cls`.
   - Remove and re-import form `frmUtilities.frm` if your existing add-in includes it and form assets are available.
5. Compile.
   - `Debug -> Compile VBAProject`.
6. Save and close `AMI_Optix.xlam`.
7. Refresh ribbon XML.
   - Open add-in in OfficeRibbonXEditor.
   - Replace `customUI14.xml` with content from `excel-addin/customUI/customUI14.xml`.
   - Save.
8. Reload add-in in Excel.
   - Excel -> `File -> Options -> Add-ins -> Manage: Excel Add-ins -> Go...`.
   - Uncheck `AMI_Optix`, click `OK`.
   - Re-open Add-ins dialog, re-check `AMI_Optix`, click `OK`.
   - Restart Excel.

### Preserving User Settings (API key/base URL/year/utilities)

- API key and user settings are stored under Windows per-user registry keys (`AMI_Optix` / `VB and VBA Program Settings` entries).
- Replacing modules inside the `.xlam` does not clear those registry values.
- API base URL is code-defined (`API_BASE_URL` in `excel-addin/src/AMI_Optix_Main.bas`), not user-entered in ribbon UI.
- Rent Roll Year selection, utility choices, and log path are preserved unless explicitly changed.

## Validation Click Path (Explicit)

1. Open target workbook.
2. Click `AMI Optix`.
3. Click `API Settings`, enter API key, click `OK`.
4. Click `Run UAP` (or `Run MIH`) once to validate API round-trip.
5. Click `Diagnostics`.
6. In `Rent Roll` group:
   - Select `Rent Roll Year`.
   - Click `Refresh Rent Tables (Selected Year)`.
   - Click `Rent Tables Status`.
7. Make one manual AMI edit, then click `Verify Manual Rents (API)`.
8. Confirm expected dialog and diagnostics data.

## Troubleshooting (Most Likely Issues)

### 1) Ribbon Not Showing or Not Updating

- Confirm `customUI14.xml` exists inside the actual installed `.xlam` (not just in repo).
- Make sure Excel is fully closed before editing ribbon XML.
- Re-run `Debug -> Compile VBAProject` to catch stale mixed modules.
- Disable/re-enable add-in from `File -> Options -> Add-ins -> Excel Add-ins`.
- If still stale, rename add-in file (for example `AMI_Optix_v2026.xlam`), re-browse to it, restart Excel.

### 2) Rent Tables Cache Cannot Find Year Workbook (`Z:` or `%APPDATA%`)

- Expected source search order:
  - `Z:\AMI_Optix\RentRollYears\<YEAR>\`
  - `%APPDATA%\AMI_Optix\RentRollYears\<YEAR>\`
- Ensure selected year in ribbon matches folder year.
- Use `Manage Rent Roll Years...` to place/update the selected year workbook into local year folder.
- Click `Refresh Rent Tables (Selected Year)`.
- Open `Diagnostics` and inspect `Rent Tables Status`:
  - `Source (Resolved Now)`
  - `Source Access Error`
  - `Cache Built From`
  - `Cache Build Reason`

### 3) Verify Shows MISMATCH

- First read diagnostics `Verify Manual Rents (API)` panel:
  - `Result`, `API`, `Year`, `Total Delta`, `Tolerance`, and mismatch table rows.
- If API year/calculator differs from selected local year, align year and refresh cache, then re-run verify.
- If cache/source changed recently, click `Refresh Rent Tables (Selected Year)` and verify again.
- A mismatch means local rent output and API evaluation differ beyond tolerance; use mismatch table (`unit_id`, local/API rent, delta, reason) to pinpoint the exact source.
