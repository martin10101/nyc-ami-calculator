# AMI Optix — Full Project Context for AI Agent

## What This Project Is

AMI Optix is a NYC affordable housing AMI (Area Median Income) optimization tool. It assigns
AMI bands (40%, 60%, 70%, 80%, 90%, 100%) to affordable housing units to maximize revenue
while meeting HPD regulatory constraints.

Two programs exist: **UAP** and **MIH**. They have DIFFERENT rules.

## UAP vs MIH — CRITICAL DIFFERENCES

| Rule | UAP | MIH |
|------|-----|-----|
| Deep affordability (≤40% AMI) | 20-21% of affordable SF | 10-11% of TOTAL BUILDING SF |
| Share denominator | `affordable` (sum of unit SF) | `total_building` (entire building) |
| MIH Option | N/A | Option 1 or Option 4 (from Prog!K4) |
| Residential SF | N/A | Read from MIH sheet (Net Floor Area) |
| Max band percent cap | From Prog!I4 | From Prog!I4 (default 135%) |

**NEVER confuse UAP and MIH constraints.** The 20-21% rule is UAP ONLY.
MIH uses 10-11% of total building SF via `mih_40_ami_sf_constraint` override.

## Architecture Overview

```
Excel Workbook (user data)
    ↓ VBA reads unit data
AMI_Optix_Automation.bas → BuildAPIPayloadV2()
    ↓ HTTP POST JSON
Flask API (app.py on Render) → /api/optimize
    ↓ passes to solver
solver.py (OR-Tools CP-SAT) → find_optimal_scenarios()
    ↓ returns scenarios JSON
AMI_Optix_ResultsWriter.bas → writes results to Excel sheets
```

## VBA Modules — What Each One Does

| Module | Purpose |
|--------|---------|
| AMI_Optix_Main.bas | Entry points, ribbon buttons, API_BASE_URL, API keys |
| AMI_Optix_API.bas | HTTP requests, JSON parsing, rent roll year management |
| AMI_Optix_DataReader.bas | Reads unit data from workbook (fuzzy header matching) |
| AMI_Optix_Ribbon.bas | Custom ribbon callbacks, dropdown management |
| AMI_Optix_ResultsWriter.bas | Writes optimization results to Excel, creates "AMI Scenarios" sheet |
| AMI_Optix_RentTables.bas | Rent table cache normalization (CSV files) |
| AMI_Optix_RentCalcTables.bas | Local rent calculation from cached tables |
| AMI_Optix_VerifyManualRents.bas | Verifies local rents match API /api/evaluate (±$1 tolerance) |
| AMI_Optix_EventHooks.bas | Live sync between Manual block and UAP sheet edits |
| AMI_Optix_Diagnostics.bas | Diagnostic info output to sheet |
| AMI_Optix_Automation.bas | Agent automation: RunOptimizationUAP_Agent(), RunOptimizationMIH_Agent() |
| AMI_Optix_Learning.bas | Soft preference learning (premium weights ONLY, never hard constraints) |

## API Endpoints

| Endpoint | Method | Purpose |
|----------|--------|---------|
| /api/optimize | POST | Main optimization (JSON units + utilities) |
| /api/evaluate | POST | Validate user assignments, compute rents |
| /api/manual_calculate | POST | Compute rents even if invalid (live sync OFF) |
| /api/analyze | POST | File upload optimization (web dashboard) |
| /healthz | GET | Health check |

## Solver Constraint System

The solver uses `share_thresholds` — a list of constraints, each with:
- `band_threshold`: AMI band cutoff (e.g., 40 = all bands ≤40%)
- `min_share` / `max_share`: fraction of denominator SF
- `denominator`: which SF pool to use (`affordable`, `residential`, or `total_building`)

The denominators dict maps names to SF values (scaled by 100 for integer math):
- `affordable` = sum of all unit net_sf
- `residential` = residential SF from project data
- `total_building` = total building net SF (used by MIH 40% AMI constraint)

## Common VBA Pitfalls

1. **`Str$()` drops leading zeros**: `Str$(0.1)` produces `".1"` not `"0.1"`. Never use `Str$()` for JSON decimals between -1 and 1. Use hardcoded string literals or `Format$()`.
2. **`Set` required for objects**: When storing a Collection or Dictionary into another Dictionary, you MUST use `Set`: `Set dict("key") = collectionObj`. Without `Set`, VBA tries to access the default property, causing Error 450.
3. **`On Error Resume Next` scope**: Always restore error handling with `On Error GoTo ErrorHandler` after the guarded section.
4. **Module naming**: `Attribute VB_Name` on line 1 must match the filename exactly.

## File Categories — What You Can vs Cannot Edit

**SAFE TO EDIT (VBA source, rebuilt automatically):**
- `excel-addin/src/*.bas`
- `excel-addin/src/*.cls`

**SAFE TO EDIT (PowerShell agent scripts):**
- `tools/excel-agent/*.ps1`
- `tools/excel-agent/*.psm1`

**DO NOT EDIT (server-side — requires GitHub push to deploy):**
- `app.py` — Flask API server
- `ami_optix/solver.py` — optimization engine
- `ami_optix/*.py` — any Python module
- `rules_config.yml` — constraint configuration

If you identify a bug in server-side code, DO NOT try to fix it locally.
Instead, use `ask_user` to report: "I found a bug in [file] at [location]: [description].
This is a server-side file that needs to be updated on GitHub and deployed to Render."

## Acceptance Test Scenarios

1. **Diagnostics smoke** — opens UAP workbook, runs diagnostics macro
2. **Run UAP optimization** — opens UAP workbook, runs RunOptimizationUAP_Agent
3. **Run MIH optimization** — opens MIH workbook, runs RunOptimizationMIH_Agent
4. **Verify Manual Rents (API)** — opens post-optimization workbook, verifies rents match API

## Key Constants

- WAAMI cap: 60% (strict <=)
- WAAMI floor: 59.1%
- Potential bands: [40, 60, 70, 80, 90, 100]
- API timeout: 90 seconds
- Rent roll years supported: 2022-2026
- Default rent roll year: 2025
- Rent cache: %APPDATA%\AMI_Optix\RentRollYears\<YEAR>\
- Shared source: Z:\AMI_Optix\RentRollYears\<YEAR>\
