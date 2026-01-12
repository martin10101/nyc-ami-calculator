# AMI Optix - Project Context Document

**Use this file to onboard a new conversation. Copy/paste the relevant sections as needed.**

---

## Project Overview

AMI Optix is an NYC affordable housing AMI optimization tool. It assigns AMI bands (40%, 60%, 70%, 80%, 90%, 100%) to affordable units to maximize revenue while meeting regulatory constraints.

**Two interfaces:**
1. **Web Dashboard** - Upload Excel file, get results (Flask app on Render)
2. **Excel Add-in** - Run optimization directly from Excel via VBA + API

---

## Architecture

```
Excel Add-in (.xlam)                    Render Server
├── AMI_Optix_Main.bas                  ├── app.py (Flask API)
├── AMI_Optix_API.bas       ──HTTPS──>  ├── ami_optix/
├── AMI_Optix_DataReader.bas            │   ├── solver.py (OR-Tools CP-SAT)
├── AMI_Optix_ResultsWriter.bas         │   ├── parser.py (Excel reader)
├── AMI_Optix_Ribbon.bas                │   ├── config_loader.py
├── frmUtilities.frm                    │   └── rent_calculator.py
└── customUI/customUI14.xml             └── rules_config.yml
```

---

## Excel Add-in Installation (IMPORTANT)

The ribbon and VBA code exist but need to be installed into Excel. Follow these steps:

### Step 1: Create the Add-in File

1. Open Excel → New blank workbook
2. Press `Alt+F11` to open VBA Editor
3. File → Import File → Import these .bas files from `excel-addin/src/`:
   - `AMI_Optix_Main.bas`
   - `AMI_Optix_API.bas`
   - `AMI_Optix_DataReader.bas`
   - `AMI_Optix_ResultsWriter.bas`
   - `AMI_Optix_Ribbon.bas`
4. File → Import File → Import the form from `excel-addin/forms/`:
   - `frmUtilities.frm`
5. Save as: `AMI_Optix.xlam` (Excel Add-in format)
   - File → Save As → Browse → Save as type: "Excel Add-in (*.xlam)"
   - Save to: `C:\Users\<YourName>\AppData\Roaming\Microsoft\AddIns\`

### Step 2: Add the Ribbon (customUI)

The ribbon won't appear automatically. You need to embed the XML:

**Option A: Use Office RibbonX Editor (Recommended)**
1. Download: https://github.com/fernandreu/office-ribbonx-editor/releases
2. Open `AMI_Optix.xlam` in RibbonX Editor
3. Insert → Office 2010+ Custom UI Part
4. Copy contents of `excel-addin/customUI/customUI14.xml` into the editor
5. Save and close

**Option B: Manual ZIP Method**
1. Rename `AMI_Optix.xlam` to `AMI_Optix.xlam.zip`
2. Extract the ZIP
3. Create folder `customUI` inside extracted folder
4. Copy `customUI14.xml` into `customUI` folder
5. Edit `[Content_Types].xml` - add this line inside `<Types>`:
   ```xml
   <Override PartName="/customUI/customUI14.xml" ContentType="application/xml"/>
   ```
6. Edit `_rels/.rels` - add this relationship:
   ```xml
   <Relationship Id="rCustomUI" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>
   ```
7. Re-zip all files and rename back to `AMI_Optix.xlam`

### Step 3: Enable the Add-in

1. Open Excel
2. File → Options → Add-ins
3. At bottom: Manage: "Excel Add-ins" → Go...
4. Click "Browse" → Navigate to your `AMI_Optix.xlam`
5. Check the box next to "AMI_Optix" → OK
6. Restart Excel

### Step 4: Set API Key

1. Click the "AMI Optix" tab in ribbon
2. Click "API Settings"
3. Enter your API key

---

## Ribbon Features

Once installed, you'll see an "AMI Optix" tab with:

| Button | Function |
|--------|----------|
| **Run Solver** | Reads unit data, calls API, writes best scenario to AMI column |
| **View Scenarios** | Shows list of all scenarios, lets you pick which one to apply |
| **Select Rent Roll** | Dropdown to choose which sheet has the unit data |
| **Refresh List** | Rescans workbook for sheets with unit data |
| **Utilities** | Opens form to configure utility payment settings |
| **API Settings** | Set/change your API key |
| **About** | Version and connection info |

---

## Utilities Configuration

The **Utilities** button opens a form with 4 dropdowns:

| Category | Options |
|----------|---------|
| **Electricity** | Tenant Pays, N/A or owner pays |
| **Cooking** | Electric Stove, Gas Stove, N/A or owner pays |
| **Heat** | Electric (ccASHP), Electric (Other), Gas, Oil, N/A or owner pays |
| **Hot Water** | Electric (Heat Pump), Electric (Other), Gas, Oil, N/A or owner pays |

Settings are saved to Windows Registry and persist across sessions.

---

## Scenario Switcher

After running the solver:
1. Click **View Scenarios**
2. See list of all scenarios with their WAAMI percentages
3. Enter the number of the scenario you want to apply
4. That scenario's AMI assignments will be written to your spreadsheet

---

## Supported File Formats

| Format | Read | Write |
|--------|------|-------|
| `.xlsx` | ✅ | ✅ |
| `.xlsm` | ✅ | ✅ |
| `.xlsb` | ✅ | ❌ |
| `.csv` | ✅ | N/A |

The add-in works with ANY open workbook regardless of format.

---

## Core Rules (rules_config.yml)

### WAAMI Constraints
| Rule | Value | Description |
|------|-------|-------------|
| `waami_cap_percent` | 60.0% | Maximum weighted average AMI |
| `waami_floor` | 59.1% | Minimum weighted average AMI (NEVER go below this) |

### Deep Affordability (40% AMI Requirement)
| Rule | Value | Description |
|------|-------|-------------|
| `deep_affordability_sf_threshold` | 0 | Applies to ALL projects (not just >10k SF) |
| `deep_affordability_min_share` | 0.20 (20%) | Minimum SF at 40% AMI |
| `deep_affordability_max_share` | 0.21 (21%) | Maximum SF at 40% AMI |

### Valid AMI Bands
```yaml
potential_bands: [40, 60, 70, 80, 90, 100]
```
**NO 50% AMI** - explicitly filtered out in solver.py line 128

### Band Limits
| Rule | Value |
|------|-------|
| `max_bands_per_scenario` | 3 |
| `max_unique_scenarios` | 25 |
| `max_band_combo_checks` | 50 |

### Premium Score Weights (for assigning higher AMI to better units)
| Factor | Weight | Description |
|--------|--------|-------------|
| Floor | 45% | Higher floor = premium unit |
| Net SF | 30% | Larger unit = premium unit |
| Bedrooms | 15% | More bedrooms = premium unit |
| Balcony | 10% | Has balcony = premium unit |

### Priority Band Combinations (tried first)
```yaml
priority_band_combos:
  - [40, 60, 90]
  - [40, 60, 100]
  - [40, 60, 80]
  - [40, 80, 100]
  - [40, 90, 100]
```

---

## How the Solver Works

1. **Input**: Units with bedrooms, net_sf, floor, client_ami (user's pre-selected AMI)
2. **Constraints**:
   - WAAMI must be 59.1% - 60%
   - 20-21% of SF must be at 40% AMI
   - Maximum 3 different AMI bands per scenario
3. **Optimization**:
   - Maximizes WAAMI (closer to 60% = more revenue)
   - Then maximizes premium alignment (best units get highest AMI)
4. **Output**: Multiple scenarios (absolute_best, best_3_band, best_2_band, alternative, client_oriented)

**The solver uses soft floor preferences (premium score), NOT hard constraints for floor placement. This makes it more flexible and ensures solutions can always be found.**

---

## Rent Calculation

The API returns rent data for each unit:

| Field | Description |
|-------|-------------|
| `gross_rent` | Rent before utility allowances |
| `monthly_rent` | Net rent after allowances |
| `annual_rent` | Net rent × 12 |
| `allowance_total` | Total utility allowances deducted |
| `allowances` | Breakdown by category (cooking, heat, etc.) |

Rent calculations require `2025 AMI Rent Calculator Unlocked.xlsx` on the server.

---

## Key Files

### Python (Server)
| File | Purpose |
|------|---------|
| `app.py` | Flask API endpoints (`/api/optimize`, `/api/analyze`) |
| `ami_optix/solver.py` | OR-Tools CP-SAT optimization engine |
| `ami_optix/parser.py` | Excel file parser, preferred sheets: UAP, PROJECT WORKSHEET, RentRoll |
| `ami_optix/rent_calculator.py` | Calculates rents based on AMI bands |
| `ami_optix/config_loader.py` | Loads rules_config.yml |
| `rules_config.yml` | All optimization rules and constraints |

### VBA (Excel Add-in)
| File | Purpose |
|------|---------|
| `AMI_Optix_Main.bas` | Entry point, `RunOptimization()`, API key management |
| `AMI_Optix_API.bas` | HTTP calls to API, JSON building/parsing |
| `AMI_Optix_DataReader.bas` | Reads units from Excel, fuzzy header matching |
| `AMI_Optix_ResultsWriter.bas` | Writes optimization results back to Excel |
| `AMI_Optix_Ribbon.bas` | Ribbon button callbacks, scenario selector, utility form launcher |
| `frmUtilities.frm` | Utility configuration form |
| `customUI/customUI14.xml` | Ribbon XML definition |

---

## VBA Header Matching

The VBA uses fuzzy matching to find columns regardless of exact header names:

| Column | Patterns Matched |
|--------|-----------------|
| Unit ID | APT, UNIT, UNIT ID, APARTMENT |
| Bedrooms | BED, BEDS, BEDROOMS, BR |
| Net SF | NET SF, SQFT, SQ FT, AREA |
| AMI | AMI FOR 35 YEARS, AMI FOR 35, AMI BAND, AMI, AFFORDABILITY, AFF %, AFF |
| Floor | FLOOR, STORY, LEVEL, CONSTRUCTION STORY, MARKETING STORY, FLR |

**Important**: "AMI AFTER 35" is EXCLUDED (that's an output column, not input)

---

## Unit Filtering

Both VBA and Python parser only process units where:
- AMI column has a positive numeric value
- Units without AMI values are skipped (they're not meant to be optimized)

---

## API Endpoint

```
POST /api/optimize
Headers: X-API-Key: <key>
Content-Type: application/json

Request:
{
  "units": [
    {"unit_id": "1A", "bedrooms": 2, "net_sf": 850, "floor": 3, "client_ami": 0.6},
    ...
  ],
  "utilities": {"electricity": "tenant_pays", "cooking": "gas", "heat": "na", "hot_water": "na"}
}

Response:
{
  "success": true,
  "scenarios": {
    "absolute_best": {
      "waami": 0.5997,
      "bands": [40, 60, 90],
      "assignments": [...],
      "rent_totals": {
        "gross_monthly": 25000.00,
        "net_monthly": 24500.00,
        "allowances_monthly": 500.00
      }
    },
    ...
  },
  "notes": [...]
}
```

---

## Recent Fixes (January 2025)

| Issue | Fix |
|-------|-----|
| Parser read wrong sheet (9 units instead of 20) | Added "UAP" to preferred sheet list |
| Solver returned only 1 scenario | Increased `max_band_combo_checks` to 50 |
| VBA sent all 83 units instead of 20 | Added AMI filtering in DataReader |
| VBA excluded "AMI FOR 35 YEARS" header | Fixed `IsAMIHeader()` pattern matching |
| `client_ami` not sent to API | Added to `BuildAPIPayload()` |
| Results wrote to ALL units | Now only writes to optimized units |
| WAAMI floor too low | Raised from 58% to 59.1% |
| Missing ribbon functions | Added `ShowUtilityForm()`, `ShowSettingsForm()`, `ShowScenarioSelector()` |

---

## Security

| Item | Status |
|------|--------|
| API key storage | ✅ Server: env var `AMI_OPTIX_API_KEY`, Client: Windows Registry |
| .env files | ✅ Not committed to git |
| HTTPS | ✅ All API calls use HTTPS |
| Authentication | ✅ X-API-Key header required |
| Sensitive files | ✅ None in repository |

---

## Known Limitations

1. **Windows Only**: VBA uses `MSXML2.XMLHTTP` (Windows-specific)
2. **Cold Start**: First API call may take 30-60s if Render instance sleeping
3. **No Offline Mode**: Requires internet connection
4. **Tight Constraints**: 20-21% at 40% AMI + 59.1-60% WAAMI limits valid configurations
5. **Ribbon Installation**: Requires manual setup with RibbonX Editor

---

## Design Decisions

1. **Soft floor preference vs hard constraint**: The solver uses premium scoring (45% floor weight) to PREFER higher AMI for upper floors, but doesn't ENFORCE it. This ensures solutions can always be found regardless of building layout.

2. **Band count optimization vs assignment strategies**: The solver optimizes by trying different band combinations and maximizing WAAMI, rather than using fixed assignment strategies (smallest units first, etc.). This leaves less money on the table.

3. **No 50% AMI**: Explicitly excluded as it's not a valid NYC affordable housing band.

4. **20-21% at 40% AMI window**: Very tight (1% range) which limits scenario diversity but meets client requirements.

---

## Environment Variables (Server)

- `AMI_OPTIX_API_KEY` - API authentication key
- Stored in Render environment settings

## Registry Keys (Excel Client)

- API key: `HKEY_CURRENT_USER\Software\VB and VBA Program Settings\AMI_Optix\Settings\APIKey`
- Utilities: `HKEY_CURRENT_USER\Software\VB and VBA Program Settings\AMI_Optix\Utilities\*`

---

## Pending Features

1. **Website PDF Rent Roll Upload** - Allow clients to upload new rent roll PDFs yearly with year selection

---

*Last updated: January 12, 2025*
