# Manual QA Plan for Rent + XLSB Workflow

## Prerequisites
- Python environment prepared with project dependencies (`pip install -r requirements.txt`).
- Node 18+ with the dashboard built (`npm install` then `npm run dev` for local testing) or exported static bundle.
- (Optional) LibreOffice on a Windows host if you want to double-check XLSB conversions.
- Sample workbooks: `Unit Schedule - UAP.xlsm`, any `.xlsb` rent roll, and the `2025 AMI Rent Calculator Unlocked.xlsx` file.

## Test Matrix
Run the following scenarios end-to-end using the dashboard. For each run capture the solver notes, verify rent totals, and download the report ZIP.

| Scenario | Upload Unit File | Rent Calculator | Prefer XLSB | Expected Reports |
| --- | --- | --- | --- | --- |
| 1 | `.xlsx` unit roll | none | off | ZIP contains only `.xlsx` files |
| 2 | `.xlsm` unit roll | bundled | on | ZIP contains `.xlsx`; solver notes mention rent workbook default |
| 3 | `.xlsb` unit roll | custom 2025 workbook | on | ZIP contains `.xlsx` with rent totals + fallback note |
| 4 | `.xlsb` unit roll | custom workbook | off | ZIP contains `.xlsx` (same as before) |

## Execution Steps
1. Start the Flask API (`python app.py`) and dashboard (`npm run dev` in `dashboard/`).
2. Step through the wizard for each scenario:
   - Upload the specified unit schedule.
   - Toggle the rent workbook according to the matrix.
   - Select utilities (exercise at least one non-`N/A` option).
   - Add a premium-weight tweak and a band whitelist to confirm overrides serialize.
   - Toggle “Return downloads in .xlsb format” per scenario.
3. Run the solver and wait for the summary card to populate.
4. Confirm metrics:
   - Total monthly/annual rent display in the project summary and scenario panel.
   - 40% AMI share stays within 20–21% (inspect project summary and scenario metrics).
5. Download the ZIP and inspect contents:
   - Ensure expected file extensions per scenario.
   - Open the updated source workbook and confirm unit assignments + rent columns.
   - Spot-check the assignments tab for monthly/annual rent columns.
6. For Scenario 3, open the generated `.xlsx` in Excel and confirm row 17 utilities and column H net rents match the dashboard summary.

## Regression Checklist
- Re-run existing `pytest` suite (`pytest -q`).
- Smoke test dashboard navigation (stepper back/next, JSON save/load, preferXlsb toggle persistence).
- When the prefer XLSB toggle is on, confirm the analysis notes mention the `.xlsx` fallback message.
