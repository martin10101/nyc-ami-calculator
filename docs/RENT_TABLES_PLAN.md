# Fix-06c — Table-driven local rent calculation (Rent Tables cache)

Date: 2026-02-18  
Repo: `martin10101/nyc-ami-calculator-all-fixes-test`  
Scope: Excel add-in manual refresh path only

## Goals
- Keep solver scenario rents/totals **unchanged** (still written exactly as returned by `/api/optimize`).
- Manual Working Copy refresh (on local edits) computes **gross rent + utility allowance + net rent** locally (no `/api/manual_calculate`).
- Replace runtime workbook layout scraping with **table-driven lookups** from normalized tables.
- Normalize tables from the **selected year** rent workbook stored at:
  - Primary: `Z:\AMI_Optix\RentRollYears\<YEAR>\`
  - Fallback: `%APPDATA%\AMI_Optix\RentRollYears\<YEAR>\`
- Support a 7-user environment:
  - Z:\ is authoritative when available
  - per-user cache allows offline usage
- Accuracy-first: **missing lookups hard-fail** (no silent 0 / partial rent writes).

## Non-goals
- No changes to solver inputs, solver constraints, or API logic.
- No changes to scenario output writing from `/api/optimize`.
- No attempt to “best-effort” infer missing rents/allowances at runtime.

## Data model (normalized)

### `rent_limits.csv`
Schema (columns):
- `Year` (integer)
- `Program` (string; currently `UAP` and `MIH`)
- `Bedrooms` (string; `studio`, `1 BR`, …, `5 BR`)
- `AMI` (string/number; canonicalized to 4 decimals, e.g. `0.6000`)
- `GrossRent` (number)

Primary key (runtime): `Program|Bedrooms|AMI`

### `utility_allowances.csv`
Schema (columns):
- `Year` (integer)
- `UtilityType` (string; `electricity`, `cooking`, `heat`, `hot_water`)
- `UtilityVariant` (string; normalized to selection codes like `tenant_pays`, `gas`, `electric_ccashp`, etc.)
- `Bedrooms` (string; `studio`, `1 BR`, …, `5 BR`)
- `Allowance` (number)

Primary key (runtime): `UtilityType|UtilityVariant|Bedrooms`

## Storage strategy

### Source workbook resolution
- Resolve the selected year workbook from:
  1) `Z:\AMI_Optix\RentRollYears\<YEAR>\RentCalculator_<YEAR>.xlsx`
  2) first `*.xlsx` / `*.xlsm` in `Z:\AMI_Optix\RentRollYears\<YEAR>\`
  3) `%APPDATA%\AMI_Optix\RentRollYears\<YEAR>\RentCalculator_<YEAR>.xlsx`
  4) first `*.xlsx` / `*.xlsm` in `%APPDATA%\AMI_Optix\RentRollYears\<YEAR>\`

### Per-user normalized cache (primary)
Cache folder:
- `%APPDATA%\AMI_Optix\RentTablesCache\<YEAR>\`

Files:
- `rent_limits.csv`
- `utility_allowances.csv`
- `cache_meta.txt` (source path + fingerprint)

Cache freshness:
- Fingerprint source workbook using `last modified time + file size`.
- Rebuild cache if:
  - any cache file is missing
  - `cache_meta.txt` does not match the resolved source path + fingerprint
  - user forces refresh via Ribbon button

## Import workflow (year workbook → cache CSV)
1) Resolve year workbook (Z:\ first, AppData fallback).
2) Compute fingerprint: `mtime=…;size=…`.
3) If cache missing/stale, open the source workbook **read-only** (hidden if opened by the add-in).
4) Extract normalized rows:
   - Prefer stable sources if present:
     - Excel table named `rent_limits` (or common variants like `tbl_rent_limits`)
     - Excel table named `utility_allowances` (or common variants like `tbl_utility_allowances`)
   - Fallback: scrape the existing `"AMI & Rent"` layout **only for import**:
     - Use `ValidateLocalRentWorkbookLayout` fail-fast fingerprinting.
     - Parse:
       - gross rent table (`col C` bedrooms, `col D` marker `of AMI`, `col G` gross rent)
       - allowance table (`row 15` headers, `row 16/17` option labels, `rows 18–23` values)
5) Validate coverage (hard-fail on missing):
   - `rent_limits`: must include at least AMI bands `0.4`, `0.6`, `0.8`, `1.0` for `studio..5 BR` for both `UAP` and `MIH`.
   - `utility_allowances`: must include all supported variants for each `studio..5 BR`:
     - electricity: `tenant_pays`
     - cooking: `electric`, `gas`
     - heat: `electric_ccashp`, `electric_other`, `gas`, `oil`
     - hot_water: `electric_heat_pump`, `electric_other`, `gas`, `oil`
6) Write CSV files + `cache_meta.txt`.

## Runtime lookup algorithm (manual refresh only)
1) `EnsureRentTablesCache(selectedYear)` (silent on success; logs diagnostics).
2) Load cache CSV → dictionaries:
   - `LoadRentLimitsCacheToDict(year)` → `Program|Bedrooms|AMI` → `GrossRent`
   - `LoadUtilityAllowancesCacheToDict(year)` → `UtilityType|Variant|Bedrooms` → `Allowance`
3) For each unit assignment:
   - Normalize AMI to a fraction (`60` → `0.6`; `0.6` stays `0.6`).
   - Map bedrooms to labels (`0` → `studio`; `>=5` → `5 BR`).
   - Gross rent lookup by `Program|Bedrooms|AMI`.
   - Allowance lookup by 4 utility categories using the user’s saved selections.
   - Net rent = gross − total allowances (floored at 0).
4) Write results only when all lookups succeed.

## Error behavior (hard-fail)
- Cache build failures:
  - Missing workbook, unreadable workbook, unexpected layout, missing required option coverage → **blocking error** and **do not write rents**.
- Runtime missing keys:
  - Any missing dictionary key → **blocking error** with:
    - year
    - cache folder path
    - unit_id
    - missing key
    - which table is missing (`rent_limits.csv` or `utility_allowances.csv`)
- No silent fallback to the runtime scraper.

## Migration plan
- Tables preferred: if the year workbook is upgraded to include named Excel tables (`rent_limits`, `utility_allowances`), imports become layout-independent.
- Scraper remains as a **cache import fallback** only (not runtime).

## Diagnostics
On successful cache ensure (build or reuse), record:
- selected year
- resolved source workbook path
- cache folder
- source fingerprint
- timestamp

Recorded to existing logs:
- Debug log (`AMI_Optix_Debug.log`) via `DebugLog(..., force:=True)`
- Append-only JSONL run log via `AppendRunLog("rent_tables_cache", ...)`

## Acceptance tests

### Automated
- `python -m pytest -q` passes

### Manual (Excel)
1) Manual edit changes rents immediately with **NO** `/api/manual_calculate` calls.
2) Switch year 2024 ↔ 2025 and confirm rent changes where expected.
3) Flip a utility variant and confirm allowance + net rent changes.
4) Force Z:\ unavailable → AppData year workbook used → cache still builds/loads.
5) Deliberately break workbook layout → cache build fails fast with clear error; rents not written.

