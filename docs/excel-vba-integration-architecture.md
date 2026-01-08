# Excel VBA Integration Architecture

## Overview

This document describes the complete architecture for integrating the NYC AMI Calculator with Excel via VBA. The solution allows clients to use their existing Excel workbooks without installing Python, while leveraging the powerful OR-Tools optimization engine hosted on Render.

**Key Principle**: Excel stays 100% intact. VBA reads data, calls API, writes results back. No file round-trip that destroys formatting.

---

## System Architecture

```
┌─────────────────────────────────────────────────────────────────────┐
│                        CLIENT'S EXCEL WORKBOOK                       │
│  ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────────┐  │
│  │   UAP Sheet     │  │ Utility Config  │  │  Scenario Sheets    │  │
│  │  (Unit Data)    │  │    (Form)       │  │  (Results)          │  │
│  │                 │  │                 │  │                     │  │
│  │ • Unit IDs      │  │ • Electricity   │  │ • Scenario 1        │  │
│  │ • Bedrooms      │  │ • Cooking       │  │ • Scenario 2        │  │
│  │ • Net SF        │  │ • Heat          │  │ • Scenario 3        │  │
│  │ • AMI (col E)   │  │ • Hot Water     │  │                     │  │
│  │ • Floor         │  │                 │  │ [Apply This] btn    │  │
│  └────────┬────────┘  └────────┬────────┘  └──────────┬──────────┘  │
│           │                    │                      │             │
│           └────────────────────┼──────────────────────┘             │
│                                │                                    │
│                         ┌──────┴──────┐                             │
│                         │  VBA Module │                             │
│                         │             │                             │
│                         │ • Read Data │                             │
│                         │ • Build JSON│                             │
│                         │ • HTTP Call │                             │
│                         │ • Parse Resp│                             │
│                         │ • Write Back│                             │
│                         └──────┬──────┘                             │
└────────────────────────────────┼────────────────────────────────────┘
                                 │
                                 │ HTTPS POST
                                 │ (JSON payload)
                                 ▼
┌─────────────────────────────────────────────────────────────────────┐
│                        RENDER API SERVER                            │
│                   https://nyc-ami-calculator.onrender.com           │
│                                                                     │
│  ┌─────────────────────────────────────────────────────────────┐   │
│  │                      /api/optimize                           │   │
│  │                                                              │   │
│  │  Input:                      Output:                         │   │
│  │  • units[]                   • scenarios[] (up to 3)         │   │
│  │  • utilities{}               • unit_assignments[]            │   │
│  │  • config{}                  • waami, rent_totals            │   │
│  │                              • notes[]                       │   │
│  └─────────────────────────────────────────────────────────────┘   │
│                                                                     │
│  ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐     │
│  │  OR-Tools       │  │  Rent Schedule  │  │  Rules Config   │     │
│  │  CP-SAT Solver  │  │  (Built-in)     │  │  (WAAMI caps)   │     │
│  └─────────────────┘  └─────────────────┘  └─────────────────┘     │
└─────────────────────────────────────────────────────────────────────┘
```

---

## Component Details

### 1. UAP Sheet (Unit Data Source)

The UAP sheet contains the building's unit data. The VBA module reads from this sheet.

**Expected Structure** (Row 4 = Headers):
| Column | Header | Description |
|--------|--------|-------------|
| A | FLOOR | Floor number |
| B | APT | Unit ID / Apartment number |
| C | BED | Number of bedrooms (0 = studio) |
| D | NET SF | Net square footage |
| E | AMI | AMI percentage (this is where results get written) |
| F | Balconies | Optional - balcony indicator |
| G | 485X | Optional - tax program indicator |

**AMI Column Detection**: The VBA uses fuzzy header matching to find the AMI column:
- Looks for headers containing "AMI" but NOT "AMI AFTER" or "AMI FOR 35"
- Falls back to column E if ambiguous
- Validates header row is row 4

---

### 2. Utility Configuration

Utilities affect rent calculations through allowances. There are two approaches:

#### Option A: Excel Form (Recommended)
A dedicated "Config" sheet or form in Excel where users select:

| Utility | Options |
|---------|---------|
| **Electricity** | "Tenant Pays" / "N/A or owner pays" |
| **Cooking** | "Electric Stove" / "Gas Stove" / "N/A or owner pays" |
| **Heat** | "Electric Heat - ccASHP" / "Electric Heat - Other" / "Gas Heat" / "Oil Heat" / "N/A or owner pays" |
| **Hot Water** | "Electric Hot Water - Heat Pump" / "Electric Hot Water - Other" / "Gas Hot Water" / "Oil Hot Water" / "N/A or owner pays" |

VBA reads these selections and includes them in the API payload.

#### Option B: API Default
If no utility config sheet exists, VBA sends `"na"` (owner pays) for all utilities.
The API uses default allowances from the built-in rent workbook.

**Recommendation**: Start with Option B (defaults), add Option A later as enhancement.

---

### 3. VBA Module Functions

The VBA module contains these key functions:

#### 3.1 `RunOptimization()` - Main Entry Point
```vba
' Triggered by "Optimize" button on UAP sheet
Sub RunOptimization()
    ' 1. Read unit data from UAP sheet
    ' 2. Read utility config (or use defaults)
    ' 3. Build JSON payload
    ' 4. Call API
    ' 5. Parse response
    ' 6. Create/update scenario sheets
    ' 7. Show summary message
End Sub
```

#### 3.2 `ReadUnitData()` - Data Extraction
```vba
Function ReadUnitData() As String
    ' Returns JSON array of units
    ' Example output:
    ' [
    '   {"unit_id": "1A", "bedrooms": 2, "net_sf": 850, "floor": 1},
    '   {"unit_id": "1B", "bedrooms": 1, "net_sf": 650, "floor": 1},
    '   ...
    ' ]

    ' Logic:
    ' 1. Find header row (row 4)
    ' 2. Find AMI column using fuzzy matching
    ' 3. Find last row with data
    ' 4. Loop through rows, build JSON objects
    ' 5. Skip rows where AMI is empty or zero
End Function
```

#### 3.3 `FindAMIColumn()` - Fuzzy Header Matching
```vba
Function FindAMIColumn() As Integer
    ' Scans row 4 for AMI column
    ' Returns column number (default: 5 for column E)

    ' Matching rules:
    ' - Header contains "AMI" (case insensitive)
    ' - Header does NOT contain "AFTER" or "FOR 35"
    ' - Prefers exact match to partial match

    ' Example matches:
    ' ✓ "AMI" -> matches
    ' ✓ "AMI BAND" -> matches
    ' ✓ "AFFORDABLE HOUSING UNIT AMI BAND" -> matches
    ' ✗ "AMI AFTER 35 YEARS" -> skip
    ' ✗ "AMI FOR 35 YEARS" -> skip
End Function
```

#### 3.4 `CallAPI()` - HTTP Request
```vba
Function CallAPI(jsonPayload As String) As String
    ' Makes HTTPS POST to Render API
    ' Returns JSON response string

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")

    http.Open "POST", "https://nyc-ami-calculator.onrender.com/api/optimize", False
    http.setRequestHeader "Content-Type", "application/json"
    http.send jsonPayload

    ' Handle timeout (Render cold start can take 30-60 seconds)
    ' Show progress indicator while waiting

    CallAPI = http.responseText
End Function
```

#### 3.5 `CreateScenarioSheets()` - Results Display
```vba
Sub CreateScenarioSheets(scenarios As Object)
    ' Creates/updates 3 scenario sheets
    ' Each sheet shows:
    ' - WAAMI percentage
    ' - Total monthly/annual rent
    ' - Unit-by-unit assignments
    ' - "Apply This" button

    ' Sheet naming: "Scenario 1", "Scenario 2", "Scenario 3"
    ' If sheets exist, clear and repopulate
    ' If sheets don't exist, create them
End Sub
```

#### 3.6 `ApplyScenario()` - Write Results Back
```vba
Sub ApplyScenario(scenarioNumber As Integer)
    ' Triggered by "Apply This" button on scenario sheet
    ' Writes AMI assignments back to UAP sheet column E

    ' Steps:
    ' 1. Read unit assignments from scenario sheet
    ' 2. For each unit, find matching row in UAP sheet
    ' 3. Write AMI value to column E
    ' 4. Highlight changed cells (optional)
    ' 5. Show confirmation message
End Sub
```

---

### 4. API Endpoint: `/api/optimize`

**URL**: `https://nyc-ami-calculator.onrender.com/api/optimize`
**Method**: POST
**Content-Type**: application/json

#### Request Payload
```json
{
  "units": [
    {
      "unit_id": "1A",
      "bedrooms": 2,
      "net_sf": 850,
      "floor": 1,
      "balcony": false
    },
    {
      "unit_id": "1B",
      "bedrooms": 1,
      "net_sf": 650,
      "floor": 1,
      "balcony": false
    }
  ],
  "utilities": {
    "electricity": "tenant_pays",
    "cooking": "gas",
    "heat": "gas",
    "hot_water": "gas"
  },
  "config": {
    "waami_cap_percent": 60.0,
    "max_scenarios": 3
  }
}
```

#### Response Payload
```json
{
  "success": true,
  "scenarios": [
    {
      "scenario_id": 1,
      "waami": 0.60,
      "waami_display": "60.00%",
      "bands_used": [40, 60, 80],
      "total_monthly_rent": 45230.50,
      "total_annual_rent": 542766.00,
      "unit_assignments": [
        {
          "unit_id": "1A",
          "assigned_ami": 0.60,
          "bedrooms": 2,
          "monthly_rent": 2150.00,
          "gross_rent": 2300.00,
          "allowance_total": 150.00
        }
      ]
    }
  ],
  "notes": [
    "Filtered 2 scenario(s) below 60% WAAMI threshold (better options available)."
  ],
  "processing_time_ms": 1250
}
```

---

### 5. Scenario Sheets Structure

Each scenario sheet displays optimization results in a clear format:

```
┌─────────────────────────────────────────────────────────────────┐
│                        SCENARIO 1                                │
│                                                                  │
│  ┌─────────────────────────────────────────────────────────┐    │
│  │  WAAMI: 60.00%                                          │    │
│  │  Bands Used: 40%, 60%, 80%                              │    │
│  │                                                          │    │
│  │  Total Monthly Rent: $45,230.50                         │    │
│  │  Total Annual Rent:  $542,766.00                        │    │
│  └─────────────────────────────────────────────────────────┘    │
│                                                                  │
│  ┌─────────────────────────────────────────────────────────┐    │
│  │  UNIT ASSIGNMENTS                                        │    │
│  │                                                          │    │
│  │  Unit ID │ Bedrooms │ AMI │ Monthly Rent │ Gross │ Allow │    │
│  │  ────────┼──────────┼─────┼──────────────┼───────┼───────│    │
│  │  1A      │ 2        │ 60% │ $2,150.00    │$2,300 │ $150  │    │
│  │  1B      │ 1        │ 40% │ $1,450.00    │$1,550 │ $100  │    │
│  │  2A      │ 2        │ 80% │ $2,650.00    │$2,800 │ $150  │    │
│  │  ...     │ ...      │ ... │ ...          │ ...   │ ...   │    │
│  └─────────────────────────────────────────────────────────┘    │
│                                                                  │
│                    ┌──────────────────┐                          │
│                    │  APPLY THIS      │  <-- Button              │
│                    └──────────────────┘                          │
│                                                                  │
└─────────────────────────────────────────────────────────────────┘
```

---

### 6. Button Workflow

#### "Optimize" Button (UAP Sheet)
**Location**: UAP sheet, near the data
**Action**: Runs full optimization workflow

```
User clicks "Optimize"
        │
        ▼
┌───────────────────┐
│ Show "Processing" │
│ dialog            │
└─────────┬─────────┘
          │
          ▼
┌───────────────────┐
│ Read unit data    │
│ from UAP sheet    │
└─────────┬─────────┘
          │
          ▼
┌───────────────────┐
│ Read utility      │
│ config (or use    │
│ defaults)         │
└─────────┬─────────┘
          │
          ▼
┌───────────────────┐
│ Build JSON        │
│ payload           │
└─────────┬─────────┘
          │
          ▼
┌───────────────────┐
│ Call Render API   │──────┐
│ (may take 30-60s  │      │ Cold start
│ on cold start)    │◄─────┘ retry
└─────────┬─────────┘
          │
          ▼
┌───────────────────┐
│ Parse JSON        │
│ response          │
└─────────┬─────────┘
          │
          ▼
┌───────────────────┐
│ Create/update     │
│ Scenario 1, 2, 3  │
│ sheets            │
└─────────┬─────────┘
          │
          ▼
┌───────────────────┐
│ Show summary:     │
│ "3 scenarios      │
│ generated"        │
└───────────────────┘
```

#### "Apply This" Button (Scenario Sheets)
**Location**: Each scenario sheet (Scenario 1, 2, 3)
**Action**: Writes AMI values back to UAP sheet

```
User clicks "Apply This" on Scenario 2
        │
        ▼
┌───────────────────┐
│ Confirm dialog:   │
│ "Apply Scenario 2 │
│ to UAP sheet?"    │
└─────────┬─────────┘
          │ Yes
          ▼
┌───────────────────┐
│ Read unit_id and  │
│ AMI from scenario │
│ sheet             │
└─────────┬─────────┘
          │
          ▼
┌───────────────────┐
│ For each unit:    │
│ Find row in UAP   │
│ by unit_id        │
└─────────┬─────────┘
          │
          ▼
┌───────────────────┐
│ Write AMI to      │
│ column E          │
│ (as decimal 0.60) │
└─────────┬─────────┘
          │
          ▼
┌───────────────────┐
│ Highlight changed │
│ cells (yellow bg) │
└─────────┬─────────┘
          │
          ▼
┌───────────────────┐
│ Show confirmation │
│ "15 units updated"│
└───────────────────┘
```

---

### 7. Error Handling

#### Network Errors
```vba
' Handle no internet connection
If http.Status = 0 Then
    MsgBox "Cannot connect to server. Check internet connection.", vbCritical
    Exit Function
End If
```

#### Timeout (Cold Start)
```vba
' Render free tier has cold starts
' First request may take 30-60 seconds
' Show progress and allow retry

http.setTimeoutSeconds 90  ' Give 90 seconds for cold start

If http.Status = 504 Or http.Status = 0 Then
    ' Gateway timeout or connection timeout
    response = MsgBox("Server is starting up. Try again?", vbRetryCancel)
    If response = vbRetry Then
        ' Retry the request
    End If
End If
```

#### API Errors
```vba
' Check for API-level errors
If response.success = False Then
    MsgBox "Optimization failed: " & response.error, vbExclamation
    Exit Sub
End If
```

#### Unit ID Mismatch
```vba
' When applying scenario, unit might not be found
' (user edited UAP sheet after optimization)

If unitRow = 0 Then
    missingUnits.Add unitId
    Continue For
End If

' After loop, warn about missing units
If missingUnits.Count > 0 Then
    MsgBox "Could not find these units in UAP sheet: " & _
           Join(missingUnits.ToArray(), ", "), vbWarning
End If
```

---

### 8. Data Flow Summary

```
┌──────────────────────────────────────────────────────────────────┐
│                         COMPLETE FLOW                             │
└──────────────────────────────────────────────────────────────────┘

1. USER ACTION
   └─► Click "Optimize" button on UAP sheet

2. VBA: READ DATA
   └─► FindAMIColumn() - locate column E (fuzzy match)
   └─► Loop rows 5 to last row with data
   └─► Build units array (unit_id, bedrooms, net_sf, floor)
   └─► Read utility selections (or use defaults)

3. VBA: BUILD REQUEST
   └─► Create JSON object with units[], utilities{}, config{}
   └─► Serialize to JSON string

4. VBA: CALL API
   └─► POST to https://nyc-ami-calculator.onrender.com/api/optimize
   └─► Wait for response (up to 90 seconds for cold start)
   └─► Receive JSON response

5. API: PROCESS
   └─► Parse units from request
   └─► Load rent schedule (built-in workbook)
   └─► Run OR-Tools CP-SAT solver
   └─► Generate up to 3 optimal scenarios
   └─► Apply WAAMI filtering (hide 59.x when 60% exists)
   └─► Calculate rents with utility allowances
   └─► Return scenarios with unit assignments

6. VBA: DISPLAY RESULTS
   └─► Parse JSON response
   └─► Create/update "Scenario 1", "Scenario 2", "Scenario 3" sheets
   └─► Populate WAAMI, totals, unit assignments
   └─► Add "Apply This" button to each sheet

7. USER ACTION (OPTIONAL)
   └─► Review scenarios
   └─► Click "Apply This" on preferred scenario

8. VBA: APPLY SCENARIO
   └─► Read unit assignments from scenario sheet
   └─► Find each unit in UAP sheet by unit_id
   └─► Write AMI value to column E
   └─► Highlight changes

9. DONE
   └─► UAP sheet now has optimized AMI values
   └─► User can save workbook normally
```

---

### 9. Implementation Phases

#### Phase 1: Core VBA Module
- [ ] `ReadUnitData()` function
- [ ] `FindAMIColumn()` with fuzzy matching
- [ ] `CallAPI()` with timeout handling
- [ ] `ParseResponse()` JSON parsing
- [ ] Basic error handling

#### Phase 2: Scenario Sheets
- [ ] `CreateScenarioSheets()` function
- [ ] Sheet layout and formatting
- [ ] Unit assignment table
- [ ] Summary section (WAAMI, totals)

#### Phase 3: Apply Functionality
- [ ] `ApplyScenario()` function
- [ ] Unit matching by unit_id
- [ ] Write AMI to column E
- [ ] Change highlighting
- [ ] Confirmation dialogs

#### Phase 4: Buttons & UI
- [ ] "Optimize" button on UAP sheet
- [ ] "Apply This" buttons on scenario sheets
- [ ] Progress indicator during API call
- [ ] Error message dialogs

#### Phase 5: Utility Config (Enhancement)
- [ ] Config sheet with dropdown selectors
- [ ] Read utility selections
- [ ] Include in API payload

---

### 10. Known Limitations

1. **Cold Start Delay**: First request after inactivity takes 30-60 seconds (Render free/starter tier)

2. **No Offline Mode**: Requires internet connection to reach Render API

3. **Excel Version**: VBA uses MSXML2.XMLHTTP which requires Windows. Mac Excel would need different HTTP approach.

4. **JSON Parsing**: VBA doesn't have native JSON support. Need to include a JSON parser (VBA-JSON library) or use simple string parsing.

5. **Concurrent Edits**: If user edits UAP sheet while viewing scenario results, "Apply This" may fail to find units.

---

### 11. Security Considerations

1. **HTTPS Only**: All API calls use HTTPS for encryption in transit

2. **No Sensitive Data**: Unit data (bedrooms, sq ft) is not personally identifiable

3. **No Authentication**: API is publicly accessible (consider adding API key in future)

4. **Input Validation**: API validates all input, rejects malformed requests

---

### 12. Testing Checklist

- [ ] VBA reads correct data from UAP sheet
- [ ] AMI column found correctly (including edge cases)
- [ ] API call succeeds with valid data
- [ ] Handles cold start timeout gracefully
- [ ] Scenario sheets created with correct data
- [ ] "Apply This" writes correct AMI values
- [ ] Error messages are user-friendly
- [ ] Works with different Excel versions (2016, 2019, 365)
- [ ] Works with different workbook structures (varying header rows)

---

## Appendix A: Sample VBA Code Structure

```vba
Option Explicit

' Constants
Private Const API_URL As String = "https://nyc-ami-calculator.onrender.com/api/optimize"
Private Const HEADER_ROW As Integer = 4
Private Const DATA_START_ROW As Integer = 5

' Main entry point
Public Sub RunOptimization()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.StatusBar = "Reading unit data..."

    ' Step 1: Read data
    Dim unitsJson As String
    unitsJson = ReadUnitData()

    ' Step 2: Build request
    Dim payload As String
    payload = BuildPayload(unitsJson)

    ' Step 3: Call API
    Application.StatusBar = "Calling optimization API (may take up to 60 seconds)..."
    Dim response As String
    response = CallAPI(payload)

    ' Step 4: Parse and display
    Application.StatusBar = "Creating scenario sheets..."
    CreateScenarioSheets response

    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Optimization complete! Review the Scenario sheets.", vbInformation
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' Additional functions would follow...
```

---

## Appendix B: API Contract

See main API documentation at `/docs/api.md` (to be created).

Key endpoints:
- `POST /api/optimize` - Run optimization
- `GET /api/health` - Check server status
- `GET /api/rent-schedule` - Get current rent schedule info

---

## Appendix C: Glossary

| Term | Definition |
|------|------------|
| **AMI** | Area Median Income - percentage threshold for affordable housing |
| **WAAMI** | Weighted Average AMI - average AMI across all units, weighted by some factor |
| **UAP** | Unit Assignment Plan - the main sheet with unit data |
| **Bands** | AMI tiers (40%, 60%, 80%, etc.) used in optimization |
| **Utility Allowance** | Amount deducted from gross rent when tenant pays utilities |
| **CP-SAT** | Constraint Programming - Satisfiability solver (Google OR-Tools) |

---

*Document created: 2025-01-08*
*Last updated: 2025-01-08*
*Branch: feature/excel-vba-integration*
