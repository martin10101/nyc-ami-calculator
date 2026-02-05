# Build Tracker — Client Feedback (2026-02-05)

This file is a “don’t lose context” checklist for the NYC AMI Optix Excel add-in + API.

## Branch / deployment notes

- Render deploy branch: `perf/api-optimize-speed-2026-02-05`
  - This branch now includes **both** API performance work **and** the Excel add-in formatting fixes.
- Validation workbook used during this feedback cycle:
  - `C:\Users\MLFLL\OneDrive\Desktop\Unit Schedule_2-35-11, MIHtest4.xlsb`

## What’s already fixed (do not regress)

### Excel add-in (VBA)

- Ribbon “Type mismatch” popup removed by making RibbonX callbacks Variant-safe.
- Rent roll dropdown + ribbon-load logic **does not scan worksheet cells** (sheet-name-only) to avoid triggering:
  - Client workbook VBA compilation, and
  - “User-defined type not defined” errors from broken workbook references (e.g. missing Microsoft Scripting Runtime).
- `AMI Scenarios` output cleanup:
  - Removed **Floor** from scenario tables.
  - Removed **Allowances** column from scenario tables.
  - Dollar columns render as **whole dollars**.
  - Added **utility deduction totals by bedroom** near the top (once per run).
  - Live Sync “Scenario Manual” AMI column detection is dynamic (no hard-coded column index).

### API / solver (Render)

- Added timing instrumentation gated by `AMI_OPTIX_TIMING_LOG=1`.
  - When enabled, server prints one JSON timing line and includes `timing` in the `/api/optimize` response.
- In-memory rent schedule caching per worker to avoid repeated `pandas.read_excel`.
- Solver speedups that **do not change feasibility/compliance constraints**:
  - Removed expensive lexicographic re-solve loop.
  - Bounded pass-2 tie-break time.
  - Configurable solver workers via `AMI_OPTIX_SOLVER_WORKERS` (default stays `1` for the $7 Render plan).

## Client “small issues” addressed next (implemented carefully)

- Alignment + sizing:
  - Unit IDs should be right-aligned.
  - Band values (band mix table) should be right-aligned.
  - Column A should be kept compact (avoid huge empty “boxes” in the utilities section).
- Utility deductions:
  - Show per-bedroom deduction **broken down by utility category** (Electricity/Cooking/Heat/Hot Water) + Total.
  - Keep it displayed **once at the top**, not repeated per scenario.
- Duplicate scenarios:
  - If two scenarios have identical assignments, the duplicate should be skipped in the Excel display.

## Debugging / timing (for future troubleshooting)

- Excel-side log file:
  - `%TEMP%\AMI_Optix_Debug.log`
  - Contains ribbon-load markers and elapsed time for the API call and main workflow steps.
- Render-side:
  - Set `AMI_OPTIX_TIMING_LOG=1` and check Render logs for per-request timing JSON.
  - Optional: temporarily lower `AMI_OPTIX_TARGET_TOTAL_SCENARIOS` if edge generation is expensive.

