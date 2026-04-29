Attribute VB_Name = "AMI_Optix_ResultsWriter"
'===============================================================================
' AMI OPTIX - Results Writer Module
' Writes optimization results back to Excel
'===============================================================================
Option Explicit

Private Const MANUAL_BLOCK_START_ROW As Long = 1
Private Const MANUAL_CLEAR_FALLBACK_HEIGHT As Long = 250

' True only when the most recent /api/evaluate call returned success=false.
Public g_AMIOptixLastManualScenarioInvalid As Boolean
' Last error encountered while building the "AMI Scenarios" sheet (blank if OK).
Public g_AMIOptixLastScenariosSheetBuildError As String
' Total building SF for MIH — used by BuildBandMix to compute share of building SF.
Public g_MihTotalBuildingSf As Double

'-------------------------------------------------------------------------------
' LOCAL RENT CALC CACHE (Fix-06)
'-------------------------------------------------------------------------------

Private Const RENTROLL_YEAR_REG_SECTION As String = "RentRollYears"
Private Const RENTROLL_YEAR_REG_KEY_SELECTED As String = "SelectedYear"
Private Const RENTROLL_YEAR_MIN As Long = 2022
Private Const RENTROLL_YEAR_MAX As Long = 2026
Private Const RENTROLL_YEAR_DEFAULT As Long = 2025

Private Const RENTROLL_SHARED_BASE As String = "Z:\AMI_Optix\RentRollYears"
Private Const RENTROLL_LOCAL_SUBPATH As String = "\AMI_Optix\RentRollYears"
Private Const RENTROLL_LOCAL_FILENAME_PREFIX As String = "RentCalculator_"
Private Const RENTROLL_LOCAL_FILENAME_SUFFIX As String = ".xlsx"

Private Const LOCAL_RENT_ERR_LAYOUT As Long = vbObjectError + 606
Private Const LOCAL_RENT_ERR_GROSS As Long = vbObjectError + 607
Private Const LOCAL_RENT_ERR_ALLOWANCE As Long = vbObjectError + 608
Private Const LOCAL_RENT_ERR_UNEXPECTED As Long = vbObjectError + 609

Private m_LocalRentYear As Long
Private m_LocalRentPath As String
Private m_LocalRentWb As Workbook
Private m_LocalGrossRents As Object ' Scripting.Dictionary (key: "<amiKey>|<bedLabel>" -> gross)
Private m_LocalAllowances As Object ' Scripting.Dictionary (category -> optionLabel -> bedLabel -> amount)
Private m_LocalRentFingerprint As String
Private m_LastLocalRentErrorSig As String

'-------------------------------------------------------------------------------
' APPLY BEST SCENARIO
'-------------------------------------------------------------------------------

Public Sub ApplyBestScenario(result As Object)
    ' Applies the best scenario's AMI assignments to the source data sheet

    Dim scenarios As Object
    Dim bestScenario As Object
    Dim assignments As Object
    Dim assignment As Object
    Dim ws As Worksheet
    Dim amiCol As Long
    Dim i As Long
    Dim unitId As String
    Dim ami As Double
    Dim amiValue As Double
    Dim row As Long
    Dim updatedCount As Long
    Dim prevEnableEvents As Boolean
    Dim prevSuppress As Boolean

    prevEnableEvents = Application.EnableEvents
    prevSuppress = g_AMIOptixSuppressEvents
    Application.EnableEvents = False
    g_AMIOptixSuppressEvents = True

    On Error GoTo ErrorHandler

    ' Get scenarios
    Set scenarios = result("scenarios")

    ' Find best scenario - use correct keys from solver
    Dim bestKey As String

    ' Priority order for best scenario (matches solver output)
    Dim priorities As Variant
    priorities = Array("absolute_best", "best_3_band", "best_2_band", "alternative", "client_oriented")

    For i = LBound(priorities) To UBound(priorities)
        If scenarios.Exists(CStr(priorities(i))) Then
            bestKey = CStr(priorities(i))
            Exit For
        End If
    Next i

    ' If no priority key found, take first available
    If bestKey = "" Then
        Dim keys As Variant
        keys = scenarios.keys
        If UBound(keys) >= 0 Then
            bestKey = keys(0)
        Else
            Exit Sub  ' No scenarios
        End If
    End If

    Set bestScenario = scenarios(bestKey)
    Set assignments = bestScenario("assignments")

    ' Get data sheet and AMI column
    Set ws = GetDataSheet()
    amiCol = GetAMIColumn()

    If ws Is Nothing Or amiCol = 0 Then
        MsgBox "Cannot write results: data sheet or AMI column not found.", vbExclamation, "AMI Optix"
        GoTo Cleanup
    End If

    ' Build lookup of unit_id to row number
    Dim unitRows As Object
    Set unitRows = BuildUnitRowLookup(ws)

    ' Apply assignments
    updatedCount = 0

    For i = 1 To assignments.Count
        Set assignment = assignments(i)

        unitId = CStr(assignment("unit_id"))
        ami = CDbl(assignment("assigned_ami"))

        ' Find row for this unit
        If unitRows.Exists(unitId) Then
            row = unitRows(unitId)

            ' Write AMI value
            ' Support both styles:
            ' - Whole-percent values (e.g., 60, 120, 130) -> divide by 100
            ' - Decimal values (e.g., 0.6, 1.2, 1.3) -> keep as-is
            If ami > 2# Then
                amiValue = ami / 100#  ' Convert 60 to 0.60; 120 to 1.20
            Else
                amiValue = ami
            End If

            ws.Cells(row, amiCol).Value = amiValue
            ws.Cells(row, amiCol).NumberFormat = "0%"  ' Ensure percentage format

            ' Highlight the cell
            ws.Cells(row, amiCol).Interior.Color = RGB(255, 255, 200)  ' Light yellow

            updatedCount = updatedCount + 1
        End If
    Next i

    Debug.Print "Applied best scenario: " & bestKey & " - Updated " & updatedCount & " units"

    ' Best-effort learning audit: record what got auto-applied.
    On Error Resume Next
    Dim programNorm As String
    Dim mihOption As String
    Dim profileKey As String
    programNorm = "UAP"
    mihOption = ""
    If Not result Is Nothing Then
        If result.Exists("project_summary") Then
            Dim ps As Object
            Set ps = result("project_summary")
            If Not ps Is Nothing Then
                If ps.Exists("program") Then programNorm = UCase$(CStr(ps("program")))
                If ps.Exists("mih_option") Then mihOption = CStr(ps("mih_option"))
            End If
        End If
    End If
    profileKey = GetLearningProfileKey(programNorm, mihOption)
    Call LogScenarioApplied(profileKey, programNorm, mihOption, bestKey, "AUTO", bestScenario)
    On Error GoTo 0

    ' Ensure WAAMI/Avg AMI display shows sufficient precision (e.g., 59.96% vs 60.0%).
    On Error Resume Next
    EnsureProvidedAvgAmiPrecision
    On Error GoTo 0
    GoTo Cleanup

ErrorHandler:
    Debug.Print "ApplyBestScenario Error: " & Err.Description
Cleanup:
    Application.EnableEvents = prevEnableEvents
    g_AMIOptixSuppressEvents = prevSuppress
End Sub

Private Function BuildUnitRowLookup(ws As Worksheet) As Object
    ' Builds a dictionary mapping unit_id to row number
    Dim lookup As Object
    Set lookup = CreateObject("Scripting.Dictionary")

    Dim unitIdCol As Long
    Dim headerRow As Long
    Dim lastRow As Long
    Dim row As Long
    Dim unitId As String

    ' Get column info from DataReader module
    unitIdCol = 0
    headerRow = GetHeaderRow()

    ' Find unit_id column (reuse mapping logic)
    Dim col As Long
    For col = 1 To 20
        Dim header As String
        header = UCase(Trim(CStr(ws.Cells(headerRow, col).Value)))
        If header = "APT" Or header = "APT #" Or header = "UNIT" Or header = "UNIT ID" Then
            unitIdCol = col
            Exit For
        End If
    Next col

    If unitIdCol = 0 Then unitIdCol = 2  ' Default to column B

    ' Find last row
    lastRow = ws.Cells(ws.Rows.Count, unitIdCol).End(xlUp).row

    ' Build lookup
    For row = headerRow + 1 To lastRow
        unitId = Trim(CStr(ws.Cells(row, unitIdCol).Value))
        If unitId <> "" And Not lookup.Exists(unitId) Then
            lookup(unitId) = row
        End If
    Next row

    Set BuildUnitRowLookup = lookup
End Function

'-------------------------------------------------------------------------------
' CREATE SCENARIOS SHEET
'-------------------------------------------------------------------------------

Public Sub CreateScenariosSheet(result As Object)
    ' Creates or updates the "AMI Scenarios" sheet with all scenarios

    Dim ws As Worksheet
    Dim scenarios As Object
    Dim scenarioKey As Variant
    Dim scenario As Object
    Dim row As Long
    Dim prevEnableEvents As Boolean
    Dim prevSuppress As Boolean
    Dim hadError As Boolean
    Dim errMsg As String

    prevEnableEvents = Application.EnableEvents
    prevSuppress = g_AMIOptixSuppressEvents
    Application.EnableEvents = False
    g_AMIOptixSuppressEvents = True
    g_AMIOptixLastScenariosSheetBuildError = ""

    On Error GoTo ErrorHandler

    ' Get or create sheet (do not delete; manual scenario sync expects stability)
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("AMI Scenarios")
    On Error GoTo ErrorHandler

    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
        ws.Name = "AMI Scenarios"
    Else
        ws.Cells.Clear
    End If
    ' Some client workbooks hide sheets via macros; force visibility so the tab doesn't "disappear".
    On Error Resume Next
    ws.Visible = xlSheetVisible
    On Error GoTo ErrorHandler

    If result Is Nothing Then GoTo ErrorHandler
    If Not result.Exists("scenarios") Then GoTo ErrorHandler
    Set scenarios = result("scenarios")

    ' Extract total building SF for MIH share-of-building display
    g_MihTotalBuildingSf = 0#
    On Error Resume Next
    If result.Exists("project_summary") Then
        Dim projSummary As Object
        Set projSummary = result("project_summary")
        If Not projSummary Is Nothing Then
            If projSummary.Exists("total_building_sf") Then
                g_MihTotalBuildingSf = CDbl(projSummary("total_building_sf"))
            End If
        End If
    End If
    On Error GoTo ErrorHandler

    ' Manual block at top (live sync area)
    Dim manualEndRow As Long
    manualEndRow = WriteManualScenarioBlockFromResult(ws, result)

    ' Start scenarios immediately after the manual block (no hard jump to row 125)
    row = manualEndRow + 2

    ' Process each scenario (include the best scenario even though it's also shown in the manual block).
    ' Users expect Scenario 1 to be the "Absolute Best" result, and the manual block is a live view.
    Dim scenarioNum As Long
    scenarioNum = 1

    ' Stable ordering + dedupe (client reported duplicates like "Scenario 4" and "Scenario 5" matching exactly).
    Dim orderedKeys As Collection
    Set orderedKeys = BuildScenarioKeyOrder(scenarios)

    Dim seenCanons As Object
    Set seenCanons = CreateObject("Scripting.Dictionary") ' canonical string -> True

    Dim k As Variant
    For Each k In orderedKeys
        scenarioKey = CStr(k)
        Set scenario = scenarios(scenarioKey)

        Dim canonKey As String
        canonKey = ScenarioCanonicalKey(scenario)
        If canonKey <> "" Then
            If seenCanons.Exists(canonKey) Then
                DebugLog "CreateScenariosSheet: skipping duplicate scenario '" & scenarioKey & "'", True
                GoTo NextScenarioKey
            End If
            seenCanons(canonKey) = True
        End If

        ' Scenario header
        ws.Cells(row, 1).Value = "SCENARIO " & scenarioNum & ": " & FormatScenarioName(CStr(scenarioKey))
        ws.Cells(row, 1).Font.Bold = True
        ws.Cells(row, 1).Font.Size = 14
        ws.Range(ws.Cells(row, 1), ws.Cells(row, 8)).Interior.Color = ScenarioHeaderColor(CStr(scenarioKey))
        row = row + 1

        row = WriteScenarioSummaryAndTable(ws, row, scenario)
        row = row + 1

        ws.Range(ws.Cells(row, 1), ws.Cells(row, 8)).Interior.Color = RGB(240, 240, 240)
        row = row + 1
        scenarioNum = scenarioNum + 1

NextScenarioKey:
    Next k

    ' Client formatting: alignment is handled per-table (Unit/AMI columns are explicitly aligned when written).

    ' Column sizing:
    ' - Keep column A compact (client requested) so the utilities block and tables don't look like huge boxes.
    ' - AutoFit the numeric columns so rents/SF are still readable.
    ws.Columns("B:K").AutoFit
    ws.Columns("A:A").ColumnWidth = 22

    ' Freeze the top row and jump to the scenarios so users immediately see the scenario list.
    On Error Resume Next
    ws.Activate
    ws.Rows(2).Select
    ActiveWindow.FreezePanes = True
    ws.Cells(manualEndRow + 2, 1).Select
    ActiveWindow.ScrollRow = manualEndRow + 2
    On Error GoTo ErrorHandler

    Debug.Print "Created scenarios sheet with " & (scenarioNum - 1) & " scenarios"
    GoTo Cleanup

ErrorHandler:
    hadError = True
    errMsg = Err.Description
    g_AMIOptixLastScenariosSheetBuildError = errMsg
    Debug.Print "CreateScenariosSheet Error: " & errMsg
Cleanup:
    Application.EnableEvents = prevEnableEvents
    g_AMIOptixSuppressEvents = prevSuppress
    If hadError Then
        MsgBox "Failed to build 'AMI Scenarios' sheet: " & errMsg, vbExclamation, "AMI Optix"
    End If
End Sub

Private Function BuildScenarioKeyOrder(scenarios As Object) As Collection
    ' Builds a stable display order for scenario keys:
    ' strict keys first, then max_revenue, then any remaining edge keys sorted A→Z.
    Dim ordered As New Collection

    On Error GoTo Fail
    If scenarios Is Nothing Then
        Set BuildScenarioKeyOrder = ordered
        Exit Function
    End If

    Dim preferred As Variant
    preferred = Array("absolute_best", "best_3_band", "best_2_band", "alternative", "client_oriented", "max_revenue")

    Dim i As Long
    For i = LBound(preferred) To UBound(preferred)
        If scenarios.Exists(CStr(preferred(i))) Then
            ordered.Add CStr(preferred(i))
        End If
    Next i

    Dim otherKeys() As String
    Dim otherCount As Long
    otherCount = 0
    ReDim otherKeys(0 To Application.Max(0, scenarios.Count - 1))

    Dim k As Variant
    For Each k In scenarios.Keys
        Dim keyStr As String
        keyStr = CStr(k)
        If Not StringInVariantArray(keyStr, preferred) Then
            otherKeys(otherCount) = keyStr
            otherCount = otherCount + 1
        End If
    Next k

    If otherCount > 0 Then
        ReDim Preserve otherKeys(0 To otherCount - 1)
        SortStringArray otherKeys
        For i = LBound(otherKeys) To UBound(otherKeys)
            ordered.Add otherKeys(i)
        Next i
    End If

    Set BuildScenarioKeyOrder = ordered
    Exit Function

Fail:
    Set BuildScenarioKeyOrder = ordered
End Function

Private Function StringInVariantArray(value As String, arr As Variant) As Boolean
    On Error GoTo Fail
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If LCase$(CStr(arr(i))) = LCase$(CStr(value)) Then
            StringInVariantArray = True
            Exit Function
        End If
    Next i
    Exit Function
Fail:
    StringInVariantArray = False
End Function

Private Sub SortStringArray(ByRef arr() As String)
    ' Simple in-place sort (A→Z). Scenario key lists are small.
    On Error GoTo Fail
    Dim i As Long, j As Long
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If UCase$(arr(j)) < UCase$(arr(i)) Then
                Dim tmp As String
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i
    Exit Sub
Fail:
End Sub

Private Function ScenarioCanonicalKey(scenario As Object) As String
    ' Returns a stable canonical string for dedupe, e.g. "202:60|204:40|...".
    On Error GoTo Fail
    ScenarioCanonicalKey = ""
    If scenario Is Nothing Then Exit Function
    If Not scenario.Exists("canonical_assignments") Then Exit Function

    Dim canon As Object
    Set canon = Nothing
    On Error Resume Next
    Set canon = scenario("canonical_assignments")
    On Error GoTo Fail
    If canon Is Nothing Then Exit Function
    If TypeName(canon) <> "Collection" Then Exit Function

    Dim i As Long
    Dim s As String
    s = ""
    For i = 1 To canon.Count
        Dim pair As Object
        Set pair = canon(i)
        If pair Is Nothing Then GoTo NextPair
        If TypeName(pair) <> "Collection" Then GoTo NextPair
        If pair.Count < 2 Then GoTo NextPair
        If s <> "" Then s = s & "|"
        s = s & CStr(pair(1)) & ":" & CStr(pair(2))
NextPair:
    Next i

    ScenarioCanonicalKey = s
    Exit Function

Fail:
    ScenarioCanonicalKey = ""
End Function

Private Function ScenarioHeaderColor(scenarioKey As String) As Long
    Dim key As String
    key = LCase$(Trim$(scenarioKey))

    ' Strict scenarios: blue
    Select Case key
        Case "absolute_best", "best_3_band", "best_2_band", "alternative", "client_oriented"
            ScenarioHeaderColor = RGB(200, 220, 255)
            Exit Function
    End Select

    ' Edge / relaxed scenarios: orange
    If InStr(1, key, "edge", vbTextCompare) > 0 Or InStr(1, key, "relaxed", vbTextCompare) > 0 Or InStr(1, key, "max_revenue", vbTextCompare) > 0 Then
        ScenarioHeaderColor = RGB(255, 225, 180)
        Exit Function
    End If

    ' Default: neutral blue
    ScenarioHeaderColor = RGB(200, 220, 255)
End Function

Public Sub ApplyCanonicalAssignmentsToDataSheet(canonicalAssignments As Object, Optional highlightColor As Long = -1)
    ' Applies canonical assignments (array of [unit_id, band_percent]) to the source data sheet.
    ' Used for Learning "Shadow" mode to apply baseline results without needing full assignment objects.

    Dim ws As Worksheet
    Dim amiCol As Long
    Dim prevEnableEvents As Boolean
    Dim prevSuppress As Boolean

    prevEnableEvents = Application.EnableEvents
    prevSuppress = g_AMIOptixSuppressEvents
    Application.EnableEvents = False
    g_AMIOptixSuppressEvents = True

    On Error GoTo ErrorHandler

    If canonicalAssignments Is Nothing Then GoTo Cleanup

    ' Get data sheet and AMI column
    Set ws = GetDataSheet()
    amiCol = GetAMIColumn()

    If ws Is Nothing Or amiCol = 0 Then
        MsgBox "Cannot write results: data sheet or AMI column not found.", vbExclamation, "AMI Optix"
        GoTo Cleanup
    End If

    ' Build lookup of unit_id to row number
    Dim unitRows As Object
    Set unitRows = BuildUnitRowLookup(ws)

    Dim i As Long
    Dim updatedCount As Long
    updatedCount = 0

    For i = 1 To canonicalAssignments.Count
        Dim unitId As String
        Dim band As Double
        If Not TryGetCanonicalPair(canonicalAssignments(i), unitId, band) Then GoTo NextPair

        Dim amiValue As Double
        If band > 2# Then
            amiValue = band / 100#
        Else
            amiValue = band
        End If

        If unitRows.Exists(unitId) Then
            Dim row As Long
            row = unitRows(unitId)
            ws.Cells(row, amiCol).Value = amiValue
            ws.Cells(row, amiCol).NumberFormat = "0%"
            If highlightColor <> -1 Then
                ws.Cells(row, amiCol).Interior.Color = highlightColor
            End If
            updatedCount = updatedCount + 1
        End If

NextPair:
    Next i

    Debug.Print "Applied canonical assignments - Updated " & updatedCount & " units"
    GoTo Cleanup

ErrorHandler:
    Debug.Print "ApplyCanonicalAssignmentsToDataSheet Error: " & Err.Description
Cleanup:
    Application.EnableEvents = prevEnableEvents
    g_AMIOptixSuppressEvents = prevSuppress
End Sub

Private Function TryGetCanonicalPair(pair As Variant, ByRef unitId As String, ByRef band As Double) As Boolean
    On Error GoTo Fail

    unitId = ""
    band = 0#

    If IsObject(pair) Then
        ' VBA-JSON parses nested arrays as Collections.
        Dim c As Object
        Set c = pair
        If c Is Nothing Then GoTo Fail
        If c.Count < 2 Then GoTo Fail
        unitId = CStr(c(1))
        band = CDbl(c(2))
        TryGetCanonicalPair = (Trim$(unitId) <> "")
        Exit Function
    End If

    If IsArray(pair) Then
        unitId = CStr(pair(0))
        band = CDbl(pair(1))
        TryGetCanonicalPair = (Trim$(unitId) <> "")
        Exit Function
    End If

Fail:
    TryGetCanonicalPair = False
End Function

Public Function UpdateManualScenario(Optional undoOnInvalid As Boolean = False, Optional programOverride As String = "") As Boolean
    ' Rebuilds the top "Scenario Manual" block from the current UAP/MUH AMI values.
    On Error GoTo ErrorHandler

    g_AMIOptixLastManualScenarioInvalid = False

    If Not HasAPIKey() Then
        UpdateManualScenario = False
        Exit Function
    End If
    If ActiveWorkbook Is Nothing Then
        UpdateManualScenario = False
        Exit Function
    End If

    Dim programNorm As String
    programNorm = UCase(Trim(programOverride))
    If programNorm <> "UAP" And programNorm <> "MIH" Then
        programNorm = "UAP"
        On Error Resume Next
        Dim wsMIH As Worksheet
        Set wsMIH = ActiveWorkbook.Worksheets("MIH")
        On Error GoTo ErrorHandler
        If Not wsMIH Is Nothing Then programNorm = "MIH"
    End If

    Dim mihOption As String
    Dim mihResidentialSF As Double
    Dim mihMaxBandPercent As Long
    mihOption = ""
    mihResidentialSF = 0
    mihMaxBandPercent = 0
    If programNorm = "MIH" Then
        If Not TryReadMIHInputs(mihOption, mihResidentialSF, mihMaxBandPercent) Then
            UpdateManualScenario = False
            Exit Function
        End If
    End If

    Dim prevSheet As Worksheet
    Set prevSheet = ActiveSheet

    ' Read units from the rent roll sheet (prefer UAP/MIH templates by name).
    Dim dataWs As Worksheet
    Set dataWs = Nothing
    On Error Resume Next
    If programNorm = "MIH" Then
        ' Prefer MIH sheet first (per client request), fallback only if needed.
        Set dataWs = ActiveWorkbook.Worksheets("MIH")
        If dataWs Is Nothing Then Set dataWs = ActiveWorkbook.Worksheets("RentRoll")
    Else
        Set dataWs = ActiveWorkbook.Worksheets("UAP")
    End If
    On Error GoTo ErrorHandler

    If Not dataWs Is Nothing Then
        dataWs.Activate
    End If

    Dim units As Collection
    Set units = ReadUnitData()
    prevSheet.Activate
    If units Is Nothing Or units.Count = 0 Then
        UpdateManualScenario = False
        Exit Function
    End If

    Dim utilities As Object
    Set utilities = GetUtilitySelectionsForProgram(programNorm)

    Dim payload As String
    payload = BuildEvaluatePayloadV2(units, utilities, programNorm, mihOption, mihResidentialSF, mihMaxBandPercent)

    Dim response As String
    response = CallEvaluateAPI(payload)
    If response = "" Then
        UpdateManualScenario = False
        Exit Function
    End If

    Dim evalResult As Object
    Set evalResult = ParseJSON(response)
    If evalResult Is Nothing Then
        UpdateManualScenario = False
        Exit Function
    End If

    If evalResult.Exists("success") Then
        If evalResult("success") = False Then
            g_AMIOptixLastManualScenarioInvalid = True
            If evalResult.Exists("errors") Then
                Dim errs As Object
                Set errs = evalResult("errors")
                Dim msg As String
                msg = "Manual scenario is invalid:" & vbCrLf & vbCrLf
                Dim i As Long
                For i = 1 To errs.Count
                    msg = msg & "- " & errs(i) & vbCrLf
                Next i
                MsgBox msg, vbExclamation, "AMI Optix"
            End If
            UpdateManualScenario = False
            Exit Function
        End If
    End If

    Dim ws As Worksheet
    Set ws = GetOrCreateScenariosSheet()

    Dim prevEnableEvents As Boolean
    Dim prevScreenUpdating As Boolean
    prevEnableEvents = Application.EnableEvents
    prevScreenUpdating = Application.ScreenUpdating

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    WriteManualScenarioBlockFromEvaluate ws, evalResult
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents

    UpdateManualScenario = True
    Exit Function

ErrorHandler:
    Debug.Print "UpdateManualScenario Error: " & Err.Description
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    On Error GoTo 0
    UpdateManualScenario = False
End Function

Private Function GetOrCreateScenariosSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("AMI Scenarios")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
        ws.Name = "AMI Scenarios"
    End If
    ' Some client workbooks hide sheets via macros; force visibility so the tab doesn't "disappear".
    On Error Resume Next
    ws.Visible = xlSheetVisible
    On Error GoTo 0
    Set GetOrCreateScenariosSheet = ws
End Function

Public Sub ClearScenarioManualBlock()
    ' Clears only the Scenario Manual block (top of AMI Scenarios), leaving Scenario 1/2/3... intact.
    On Error GoTo Fail

    If ActiveWorkbook Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Set ws = GetOrCreateScenariosSheet()
    ClearManualBlock ws
    Exit Sub

Fail:
End Sub

Public Sub ClearProgramAmiColumn(programNorm As String)
    ' Clears the entire AMI input column (below header) on the active program sheet (UAP/MIH).
    On Error GoTo Fail

    If ActiveWorkbook Is Nothing Then Exit Sub

    Dim prevSheet As Worksheet
    Set prevSheet = ActiveSheet

    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    If UCase$(Trim$(programNorm)) = "MIH" Then
        ' Prefer MIH sheet first (per client request), fallback only if needed.
        Set ws = ActiveWorkbook.Worksheets("MIH")
        If ws Is Nothing Then Set ws = ActiveWorkbook.Worksheets("RentRoll")
        If ws Is Nothing Then Set ws = ActiveWorkbook.Worksheets("UAP")
        If ws Is Nothing Then Set ws = ActiveWorkbook.Worksheets("PROJECT WORKSHEET")
    Else
        Set ws = ActiveWorkbook.Worksheets("UAP")
    End If
    On Error GoTo Fail

    If ws Is Nothing Then GoTo Cleanup

    ws.Activate

    ' Force a column map so GetAMIColumn/GetHeaderRow are valid even if zero units have AMI values.
    Dim dummy As Collection
    Set dummy = ReadUnitData()

    Dim dataWs As Worksheet
    Dim headerRow As Long
    Dim amiCol As Long
    Set dataWs = GetDataSheet()
    headerRow = GetHeaderRow()
    amiCol = GetAMIColumn()

    If dataWs Is Nothing Or headerRow = 0 Or amiCol = 0 Then GoTo Cleanup

    Dim used As Range
    Set used = dataWs.UsedRange

    Dim lastRow As Long
    lastRow = used.row + used.Rows.Count - 1
    If lastRow < headerRow + 1 Then lastRow = headerRow + 1

    Dim rng As Range
    Set rng = dataWs.Range(dataWs.Cells(headerRow + 1, amiCol), dataWs.Cells(lastRow, amiCol))
    rng.ClearContents
    rng.Interior.Pattern = xlNone

Cleanup:
    If Not prevSheet Is Nothing Then prevSheet.Activate
    Exit Sub

Fail:
    On Error Resume Next
    If Not prevSheet Is Nothing Then prevSheet.Activate
End Sub

Public Function ManualCalculateScenario(Optional programOverride As String = "") As Boolean
    ' Computes the Scenario Manual results from the current sheet inputs without enforcing constraints.
    ' Uses /api/manual_calculate so we can show tradeoffs instead of reverting edits.
    On Error GoTo ErrorHandler

    ManualCalculateScenario = False

    If Not HasAPIKey() Then
        MsgBox "API key is not configured." & vbCrLf & vbCrLf & _
               "Click AMI Optix → API Settings to set your key.", _
               vbExclamation, "AMI Optix"
        Exit Function
    End If
    If ActiveWorkbook Is Nothing Then Exit Function

    Dim programNorm As String
    programNorm = UCase$(Trim$(programOverride))
    If programNorm <> "UAP" And programNorm <> "MIH" Then
        programNorm = DetectProgramFromWorkbook()
    End If

    Dim mihOption As String
    Dim mihResidentialSF As Double
    Dim mihMaxBandPercent As Long
    mihOption = ""
    mihResidentialSF = 0
    mihMaxBandPercent = 0
    If programNorm = "MIH" Then
        If Not TryReadMIHInputs(mihOption, mihResidentialSF, mihMaxBandPercent) Then Exit Function
        If mihResidentialSF > 0 Then g_MihTotalBuildingSf = mihResidentialSF
    End If

    Dim prevSheet As Worksheet
    Set prevSheet = ActiveSheet

    ' Read units from the program sheet (prefer template sheet by name).
    Dim dataWs As Worksheet
    Set dataWs = Nothing
    On Error Resume Next
    If programNorm = "MIH" Then
        ' Prefer MIH sheet first.
        Set dataWs = ActiveWorkbook.Worksheets("MIH")
        If dataWs Is Nothing Then Set dataWs = ActiveWorkbook.Worksheets("RentRoll")
        If dataWs Is Nothing Then Set dataWs = ActiveWorkbook.Worksheets("UAP")
        If dataWs Is Nothing Then Set dataWs = ActiveWorkbook.Worksheets("PROJECT WORKSHEET")
    Else
        Set dataWs = ActiveWorkbook.Worksheets("UAP")
    End If
    On Error GoTo ErrorHandler
    If dataWs Is Nothing Then Exit Function

    dataWs.Activate
    Dim units As Collection
    Set units = ReadUnitData()
    prevSheet.Activate

    If units Is Nothing Or units.Count = 0 Then
        MsgBox "Could not read units from the workbook." & vbCrLf & vbCrLf & _
               "Make sure the AMI column has numeric values for the units you want to calculate.", _
               vbExclamation, "AMI Optix"
        Exit Function
    End If

    Dim utilities As Object
    Set utilities = GetUtilitySelectionsForProgram(programNorm)

    Dim payload As String
    payload = BuildEvaluatePayloadV2(units, utilities, programNorm, mihOption, mihResidentialSF, mihMaxBandPercent)

    Dim response As String
    response = CallManualCalculateAPI(payload)
    If response = "" Then Exit Function

    Dim evalResult As Object
    Set evalResult = ParseJSON(response)
    If evalResult Is Nothing Then Exit Function

    Dim ws As Worksheet
    Set ws = GetOrCreateScenariosSheet()

    Dim prevEnableEvents As Boolean
    Dim prevScreenUpdating As Boolean
    prevEnableEvents = Application.EnableEvents
    prevScreenUpdating = Application.ScreenUpdating

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Dim headerLabel As String
    If GetLiveSyncEnabled() Then
        headerLabel = "SCENARIO MANUAL (LIVE SYNC)"
    Else
        headerLabel = "SCENARIO MANUAL (CUSTOM - LIVE SYNC OFF)"
    End If

    Call WriteManualScenarioBlockFromEvaluate(ws, evalResult, headerLabel)

    ' Recalculate rents for all solver scenario blocks with the selected year.
    RecalculateSolverScenarioRents ws, programNorm, utilities, mihOption, mihResidentialSF, mihMaxBandPercent

    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents

    ' Ensure Avg AMI display shows sufficient precision (e.g., 59.96% vs 60.0%).
    On Error Resume Next
    EnsureProvidedAvgAmiPrecision
    On Error GoTo 0

    ws.Activate
    ManualCalculateScenario = True
    Exit Function

ErrorHandler:
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ManualCalculateScenario = False
End Function

Public Function RefreshManualWorkingCopyLocalRents(Optional programOverride As String = "", Optional headerLabel As String = "SCENARIO MANUAL (WORKING COPY)") As Boolean
    ' Fix-06: Rebuild the Scenario Manual block and compute rents/totals locally from the selected year's rent workbook.
    ' IMPORTANT: This function must NOT call /api/manual_calculate.
    On Error GoTo ErrorHandler

    RefreshManualWorkingCopyLocalRents = False

    If ActiveWorkbook Is Nothing Then Exit Function

    Dim programNorm As String
    programNorm = UCase$(Trim$(programOverride))
    If programNorm <> "UAP" And programNorm <> "MIH" Then
        programNorm = DetectProgramFromWorkbook()
    End If

    Dim prevSheet As Worksheet
    Set prevSheet = ActiveSheet

    ' Read units from the program sheet (prefer template sheet by name).
    Dim dataWs As Worksheet
    Set dataWs = Nothing
    On Error Resume Next
    If programNorm = "MIH" Then
        ' Prefer MIH sheet first.
        Set dataWs = ActiveWorkbook.Worksheets("MIH")
        If dataWs Is Nothing Then Set dataWs = ActiveWorkbook.Worksheets("RentRoll")
        If dataWs Is Nothing Then Set dataWs = ActiveWorkbook.Worksheets("UAP")
        If dataWs Is Nothing Then Set dataWs = ActiveWorkbook.Worksheets("PROJECT WORKSHEET")
    Else
        Set dataWs = ActiveWorkbook.Worksheets("UAP")
    End If
    On Error GoTo ErrorHandler
    If dataWs Is Nothing Then Exit Function

    dataWs.Activate
    Dim units As Collection
    Set units = ReadUnitData()
    prevSheet.Activate

    If units Is Nothing Or units.Count = 0 Then Exit Function

    Dim utilities As Object
    Set utilities = GetUtilitySelectionsForProgram(programNorm)

    Dim tradeoffs As Collection
    Set tradeoffs = New Collection

    Dim assignments As Collection
    Set assignments = BuildAssignmentsFromUnits(units)

    Dim rentOk As Boolean
    rentOk = False

    Dim rentTotals As Object
    Set rentTotals = Nothing

    Dim rentError As String
    rentError = ""

    On Error GoTo RentFail

    Dim selectedYear As Long
    selectedYear = GetSelectedRentRollYearLocal()

    Dim sourcePath As String
    Dim cacheFolder As String
    Dim sourceFp As String
    sourcePath = ""
    cacheFolder = ""
    sourceFp = ""

    Call EnsureRentTablesCache(selectedYear, False, sourcePath, cacheFolder, sourceFp)

    ' Load normalized cache tables (CSV -> Dict).
    Call LoadRentLimitsCacheToDict(selectedYear)
    Call LoadUtilityAllowancesCacheToDict(selectedYear)

    Dim totalNet As Double
    totalNet = 0#

    Dim i As Long
    For i = 1 To assignments.Count
        Dim a As Object
        Set a = assignments(i)
        If a Is Nothing Then GoTo NextAssignment

        Dim unitId As String
        unitId = ""
        On Error Resume Next
        If a.Exists("unit_id") Then unitId = CStr(a("unit_id"))
        On Error GoTo RentFail

        Dim ami As Double
        ami = 0#
        If a.Exists("assigned_ami") Then ami = CDbl(a("assigned_ami"))
        If ami > 2# Then ami = ami / 100#
        a("assigned_ami") = ami

        Dim rentResult As Object
        Set rentResult = ComputeNetRent(selectedYear, programNorm, a("bedrooms"), ami, utilities, unitId)

        a("gross_rent") = rentResult("gross_rent")
        a("monthly_rent") = rentResult("monthly_rent")
        a("annual_rent") = rentResult("annual_rent")
        a("allowance_total") = rentResult("allowance_total")
        Set a("allowances") = rentResult("allowances")

        If IsNumeric(a("monthly_rent")) Then totalNet = totalNet + CDbl(a("monthly_rent"))

NextAssignment:
    Next i

    Set rentTotals = CreateObject("Scripting.Dictionary")
    rentTotals("net_monthly") = Round(totalNet, 2)
    rentTotals("net_annual") = Round(totalNet * 12#, 2)
    rentOk = True

    tradeoffs.Add "Local rent calc: table-driven cache OK (" & CStr(selectedYear) & ")."

    On Error GoTo ErrorHandler
    GoTo AfterRent

RentFail:
    rentError = Err.Description
    Debug.Print "RefreshManualWorkingCopyLocalRents RentFail: " & rentError
    On Error GoTo ErrorHandler
    rentOk = False
    Call ClearLocalRentComputedFields(assignments, True)

    Dim firstLine As String
    firstLine = rentError
    If InStr(firstLine, vbCrLf) > 0 Then firstLine = Left$(firstLine, InStr(firstLine, vbCrLf) - 1)
    If Trim$(firstLine) <> "" Then
        tradeoffs.Add firstLine & " (rents not updated)"
    Else
        tradeoffs.Add "Local rent calc ERROR (rents not updated)."
    End If
    Call ShowLocalRentCalcErrorOnce(rentError)

AfterRent:
    ' Hard block on local rent calc failure: do not rewrite the manual table with partial/blank rents.
    If Not rentOk Then
        RefreshManualWorkingCopyLocalRents = False
        Exit Function
    End If

    Dim waami As Double
    waami = ComputeWaami(assignments)

    Dim bands As Collection
    Set bands = ComputeBandsUsed(assignments)

    Dim metrics As Object
    Set metrics = CreateObject("Scripting.Dictionary")
    Set metrics("band_mix") = BuildBandMix(assignments, g_MihTotalBuildingSf)

    Dim scenario As Object
    Set scenario = CreateObject("Scripting.Dictionary")
    scenario("waami") = waami
    scenario("bands") = bands
    scenario("metrics") = metrics
    scenario("tradeoffs") = tradeoffs
    scenario("assignments") = assignments
    If rentOk Then
        scenario("rent_totals") = rentTotals
    End If

    Dim ws As Worksheet
    Set ws = GetOrCreateScenariosSheet()

    Dim prevEnableEvents As Boolean
    Dim prevScreenUpdating As Boolean
    prevEnableEvents = Application.EnableEvents
    prevScreenUpdating = Application.ScreenUpdating

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ClearManualBlock ws

    Dim row As Long
    row = MANUAL_BLOCK_START_ROW

    ws.Cells(row, 1).Value = "AMI OPTIMIZATION RESULTS"
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 1).Font.Size = 16
    row = row + 2

    row = WriteUtilitySettings(ws, row)
    row = WriteUtilityDeductionTotalsByBedroom(ws, row, scenario)
    row = row + 1

    ws.Cells(row, 1).Value = headerLabel
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 1).Font.Size = 14
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 13)).Interior.Color = RGB(220, 240, 220)
    row = row + 1

    row = WriteScenarioSummaryAndTable(ws, row, scenario)

    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents

    RefreshManualWorkingCopyLocalRents = True
    Exit Function

ErrorHandler:
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    RefreshManualWorkingCopyLocalRents = False
End Function

Private Sub ClearLocalRentComputedFields(assignments As Collection, Optional clearAllowances As Boolean = True)
    ' Ensures we never write partial rent results to the Manual Working Copy table.
    On Error Resume Next
    If assignments Is Nothing Then Exit Sub

    Dim i As Long
    For i = 1 To assignments.Count
        Dim a As Object
        Set a = assignments(i)
        If a Is Nothing Then GoTo NextA

        If a.Exists("gross_rent") Then a.Remove "gross_rent"
        If a.Exists("monthly_rent") Then a.Remove "monthly_rent"
        If a.Exists("annual_rent") Then a.Remove "annual_rent"
        If a.Exists("allowance_total") Then a.Remove "allowance_total"
        If clearAllowances Then
            If a.Exists("allowances") Then a.Remove "allowances"
        End If

NextA:
    Next i
End Sub

Private Function BuildAssignmentsFromUnits(units As Collection) As Collection
    Dim assignments As Collection
    Set assignments = New Collection

    Dim i As Long
    For i = 1 To units.Count
        Dim u As Object
        Set u = units(i)
        If u Is Nothing Then GoTo NextUnit

        Dim a As Object
        Set a = CreateObject("Scripting.Dictionary")

        If u.Exists("unit_id") Then a("unit_id") = CStr(u("unit_id"))
        If u.Exists("bedrooms") Then a("bedrooms") = u("bedrooms")
        If u.Exists("net_sf") Then a("net_sf") = u("net_sf")
        If u.Exists("floor") Then a("floor") = u("floor")
        If u.Exists("balcony") Then a("balcony") = u("balcony")

        Dim ami As Double
        ami = 0#
        On Error Resume Next
        If u.Exists("client_ami") Then
            If IsNumeric(u("client_ami")) Then ami = CDbl(u("client_ami"))
        End If
        On Error GoTo 0
        If ami > 2# Then ami = ami / 100#
        a("assigned_ami") = ami

        assignments.Add a

NextUnit:
    Next i

    Set BuildAssignmentsFromUnits = assignments
End Function

Private Function GetSelectedRentRollYearLocal() As Long
    Dim raw As String
    raw = GetSetting("AMI_Optix", RENTROLL_YEAR_REG_SECTION, RENTROLL_YEAR_REG_KEY_SELECTED, CStr(RENTROLL_YEAR_DEFAULT))

    Dim y As Long
    y = RENTROLL_YEAR_DEFAULT
    On Error Resume Next
    y = CLng(raw)
    On Error GoTo 0

    If y < RENTROLL_YEAR_MIN Or y > RENTROLL_YEAR_MAX Then y = RENTROLL_YEAR_DEFAULT
    GetSelectedRentRollYearLocal = y
End Function

Private Function FirstWorkbookInFolder(folderPath As String) As String
    ' Returns the first *.xlsx or *.xlsm found in a folder (alphabetical by Dir enumeration).
    On Error GoTo SafeExit

    If Trim$(folderPath) = "" Then GoTo SafeExit

    Dim f As String
    f = Dir(folderPath & "\*.xlsx")
    If Trim$(f) <> "" Then
        FirstWorkbookInFolder = folderPath & "\" & f
        Exit Function
    End If

    f = Dir(folderPath & "\*.xlsm")
    If Trim$(f) <> "" Then
        FirstWorkbookInFolder = folderPath & "\" & f
        Exit Function
    End If

SafeExit:
    FirstWorkbookInFolder = ""
End Function

Private Function ResolveRentWorkbookPath(year As Long, ByRef sourceLabel As String) As String
    ' Fix-06: Prefer shared Z:\ location, fallback to per-user %APPDATA% cache.
    Dim preferredName As String
    preferredName = RENTROLL_LOCAL_FILENAME_PREFIX & CStr(year) & RENTROLL_LOCAL_FILENAME_SUFFIX

    Dim sharedFolder As String
    sharedFolder = RENTROLL_SHARED_BASE & "\" & CStr(year)

    Dim sharedPreferred As String
    sharedPreferred = sharedFolder & "\" & preferredName
    If Dir(sharedPreferred) <> "" Then
        sourceLabel = "Z:"
        ResolveRentWorkbookPath = sharedPreferred
        Exit Function
    End If

    Dim sharedAny As String
    sharedAny = FirstWorkbookInFolder(sharedFolder)
    If Trim$(sharedAny) <> "" Then
        sourceLabel = "Z:"
        ResolveRentWorkbookPath = sharedAny
        Exit Function
    End If

    Dim appDataBase As String
    appDataBase = Environ$("APPDATA") & RENTROLL_LOCAL_SUBPATH & "\" & CStr(year)

    Dim localPreferred As String
    localPreferred = appDataBase & "\" & preferredName
    If Dir(localPreferred) <> "" Then
        sourceLabel = "AppData"
        ResolveRentWorkbookPath = localPreferred
        Exit Function
    End If

    Dim localAny As String
    localAny = FirstWorkbookInFolder(appDataBase)
    If Trim$(localAny) <> "" Then
        sourceLabel = "AppData"
        ResolveRentWorkbookPath = localAny
        Exit Function
    End If

    sourceLabel = ""
    ResolveRentWorkbookPath = ""
End Function

Private Sub CloseLocalRentCache()
    On Error Resume Next

    If Not m_LocalRentWb Is Nothing Then
        m_LocalRentWb.Close SaveChanges:=False
    End If

    Set m_LocalRentWb = Nothing
    Set m_LocalGrossRents = Nothing
    Set m_LocalAllowances = Nothing
    m_LocalRentYear = 0
    m_LocalRentPath = ""
    m_LocalRentFingerprint = ""
End Sub

Private Function EnsureLocalRentScheduleReady(ByRef tradeoffs As Collection) As Boolean
    On Error GoTo Fail

    Dim year As Long
    year = GetSelectedRentRollYearLocal()

    Dim sourceLabel As String
    Dim path As String
    sourceLabel = ""
    path = ResolveRentWorkbookPath(year, sourceLabel)

    If Trim$(path) = "" Then
        tradeoffs.Add "Local rent calc: rent workbook not found for year " & CStr(year) & " (checked Z:\\ and %APPDATA%)."
        Call CloseLocalRentCache
        EnsureLocalRentScheduleReady = False
        Exit Function
    End If

    If Not m_LocalRentWb Is Nothing Then
        If m_LocalRentYear = year And UCase$(m_LocalRentPath) = UCase$(path) Then
            EnsureLocalRentScheduleReady = True
            Exit Function
        End If
        Call CloseLocalRentCache
    End If

    Dim prevSheet As Worksheet
    Set prevSheet = ActiveSheet

    Set m_LocalRentWb = Workbooks.Open(path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, Notify:=False)
    m_LocalRentYear = year
    m_LocalRentPath = path

    ' Hide the rent workbook so it doesn't distract the user.
    On Error Resume Next
    m_LocalRentWb.Windows(1).Visible = False
    On Error GoTo Fail

    If Not prevSheet Is Nothing Then prevSheet.Activate

    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = m_LocalRentWb.Worksheets("AMI & Rent")
    On Error GoTo Fail

    If ws Is Nothing Then
        tradeoffs.Add "Local rent calc: sheet 'AMI & Rent' not found in " & path
        Call CloseLocalRentCache
        EnsureLocalRentScheduleReady = False
        Exit Function
    End If

    Dim layoutReason As String
    Dim fingerprint As String
    layoutReason = ""
    fingerprint = ""
    If Not ValidateLocalRentWorkbookLayout(ws, fingerprint, layoutReason) Then
        m_LocalRentFingerprint = fingerprint
        Call CloseLocalRentCache
        Err.Raise LOCAL_RENT_ERR_LAYOUT, "AMI_Optix_ResultsWriter.LocalRentCalc", _
                  "Local rent workbook layout not recognized." & vbCrLf & vbCrLf & _
                  "Year: " & CStr(year) & vbCrLf & _
                  "Workbook: " & path & vbCrLf & _
                  "Sheet: AMI & Rent" & vbCrLf & _
                  "Reason: " & layoutReason & vbCrLf & _
                  "Fingerprint: " & fingerprint
    End If
    m_LocalRentFingerprint = fingerprint

    If Not LoadLocalRentLookups(ws, tradeoffs) Then
        Call CloseLocalRentCache
        EnsureLocalRentScheduleReady = False
        Exit Function
    End If

    Dim missing As String
    missing = ""
    If Not ValidateLocalRentLookupOptions(missing) Then
        Call CloseLocalRentCache
        Err.Raise LOCAL_RENT_ERR_LAYOUT, "AMI_Optix_ResultsWriter.LocalRentCalc", _
                  "Local rent workbook is missing required allowance option labels." & vbCrLf & vbCrLf & _
                  "Year: " & CStr(year) & vbCrLf & _
                  "Workbook: " & path & vbCrLf & _
                  "Sheet: AMI & Rent" & vbCrLf & _
                  "Missing: " & missing & vbCrLf & _
                  "Fingerprint: " & m_LocalRentFingerprint
    End If

    tradeoffs.Add "Local rent calc: using " & sourceLabel & " rent workbook (" & CStr(year) & ")."
    EnsureLocalRentScheduleReady = True
    Exit Function

Fail:
    Dim n As Long
    Dim src As String
    Dim desc As String
    n = Err.Number
    src = Err.Source
    desc = Err.Description

    tradeoffs.Add "Local rent calc failed: " & desc
    Call CloseLocalRentCache
    EnsureLocalRentScheduleReady = False
    If n <> 0 Then Err.Raise n, src, desc
End Function

Public Function ValidateLocalRentWorkbookLayout(rentWs As Worksheet, ByRef fingerprint As String, ByRef reason As String) As Boolean
    ' Fix-06b: Fail fast if the rent workbook no longer matches the expected "AMI & Rent" layout
    ' used by the local lookup parser (layout-scraping).
    On Error GoTo Fail

    ValidateLocalRentWorkbookLayout = False
    fingerprint = ""
    reason = ""

    If rentWs Is Nothing Then
        reason = "Worksheet is missing."
        Exit Function
    End If

    ' --- Allowance headers rows 15-16 ---
    Dim catCols As Object
    Set catCols = CreateObject("Scripting.Dictionary") ' category -> first col index

    Dim col As Long
    Dim hRow As Long
    For hRow = 15 To 16
        For col = 1 To 200
            Dim hv As String
            hv = Trim$(CStr(rentWs.Cells(hRow, col).Value))
            Dim cat As String
            cat = UtilityCategoryFromHeaderValue(hv)
            If cat <> "" Then
                If Not catCols.Exists(cat) Then catCols(cat) = col
            End If
        Next col
    Next hRow

    Dim missingCats As String
    missingCats = ""
    If Not catCols.Exists("electricity") Then missingCats = missingCats & IIf(missingCats <> "", ", ", "") & "electricity"
    If Not catCols.Exists("cooking") Then missingCats = missingCats & IIf(missingCats <> "", ", ", "") & "cooking"
    If Not catCols.Exists("heat") Then missingCats = missingCats & IIf(missingCats <> "", ", ", "") & "heat"
    If Not catCols.Exists("hot_water") Then missingCats = missingCats & IIf(missingCats <> "", ", ", "") & "hot_water"
    If missingCats <> "" Then
        reason = "Missing expected allowance headers in rows 15-16: " & missingCats
    End If

    ' --- Gross table: find first "of AMI" marker in col D with numeric AMI in col C ---
    Dim firstOfAmiRow As Long
    firstOfAmiRow = 0
    Dim r As Long
    For r = 1 To 2000
        Dim marker As String
        marker = LCase$(Trim$(CStr(rentWs.Cells(r, 4).Value)))
        Dim cVal As Variant
        cVal = rentWs.Cells(r, 3).Value
        If marker = "of ami" And IsNumeric(cVal) Then
            firstOfAmiRow = r
            Exit For
        End If
    Next r
    If firstOfAmiRow = 0 Then
        If reason <> "" Then reason = reason & " | "
        reason = reason & "Could not find 'of AMI' marker in column D with numeric AMI in column C."
    End If

    ' Verify we can see bedroom labels soon after the marker.
    Dim bedFound As Long
    bedFound = 0
    If firstOfAmiRow > 0 Then
        For r = firstOfAmiRow To Application.Min(firstOfAmiRow + 60, 5000)
            Dim bedLabel As String
            bedLabel = BedroomLabelFromSheetLabel(CStr(rentWs.Cells(r, 3).Value))
            If bedLabel <> "" Then
                bedFound = bedFound + 1
                If bedFound >= 2 Then Exit For
            End If
        Next r
        If bedFound < 2 Then
            If reason <> "" Then reason = reason & " | "
            reason = reason & "Gross rent table did not show expected bedroom labels following the first 'of AMI' marker."
        End If
    End If

    fingerprint = "sheet='AMI & Rent'; allowances_rows=15-23; gross_cols=C/D/G; " & _
                  "cat_cols[electricity=" & IIf(catCols.Exists("electricity"), CStr(catCols("electricity")), "?") & _
                  ", cooking=" & IIf(catCols.Exists("cooking"), CStr(catCols("cooking")), "?") & _
                  ", heat=" & IIf(catCols.Exists("heat"), CStr(catCols("heat")), "?") & _
                  ", hot_water=" & IIf(catCols.Exists("hot_water"), CStr(catCols("hot_water")), "?") & _
                  "]; first_of_ami_row=" & IIf(firstOfAmiRow > 0, CStr(firstOfAmiRow), "?")

    If reason <> "" Then Exit Function

    ValidateLocalRentWorkbookLayout = True
    Exit Function

Fail:
    fingerprint = "sheet='AMI & Rent'; fingerprint_failed"
    reason = "Fingerprint check failed: " & Err.Description
    ValidateLocalRentWorkbookLayout = False
End Function

Private Function ValidateLocalRentLookupOptions(ByRef missing As String) As Boolean
    ' Validates that our utility variant label mappings exist in the parsed allowance table.
    missing = ""
    ValidateLocalRentLookupOptions = False

    If m_LocalAllowances Is Nothing Then
        missing = "allowances not parsed"
        Exit Function
    End If

    Dim required As Object
    Set required = CreateObject("Scripting.Dictionary") ' category -> array(optionLabel)
    required("electricity") = Array("Tenant Pays")
    required("cooking") = Array("Electric Stove", "Gas Stove")
    required("heat") = Array("Electric Heat - Cold Climate Air Source Heat Pump (ccASHP)1", "Electric Heat - Other2", "Gas Heat", "Oil Heat")
    required("hot_water") = Array("Electric Hot Water - Heat Pump", "Electric Hot Water - Other", "Gas Hot Water", "Oil Hot Water")

    Dim catKey As Variant
    For Each catKey In required.Keys
        Dim cat As String
        cat = CStr(catKey)

        If Not m_LocalAllowances.Exists(cat) Then
            missing = missing & IIf(missing <> "", "; ", "") & cat & ":<category missing>"
            GoTo NextCat
        End If

        Dim catDict As Object
        Set catDict = m_LocalAllowances(cat)
        If catDict Is Nothing Then
            missing = missing & IIf(missing <> "", "; ", "") & cat & ":<category missing>"
            GoTo NextCat
        End If

        Dim opts As Variant
        opts = required(cat)

        Dim i As Long
        For i = LBound(opts) To UBound(opts)
            Dim opt As String
            opt = CStr(opts(i))
            If Not catDict.Exists(opt) Then
                missing = missing & IIf(missing <> "", "; ", "") & cat & ":" & opt
            End If
        Next i

NextCat:
    Next catKey

    ValidateLocalRentLookupOptions = (Trim$(missing) = "")
End Function

Private Sub ShowLocalRentCalcErrorOnce(message As String)
    ' Avoid spamming a modal popup on every edit when the underlying error is the same.
    On Error Resume Next

    Dim sig As String
    sig = Trim$(CStr(message))
    If sig = "" Then Exit Sub
    If Len(sig) > 500 Then sig = Left$(sig, 500)

    If sig = m_LastLocalRentErrorSig Then Exit Sub
    m_LastLocalRentErrorSig = sig

    MsgBox message, vbCritical, "AMI Optix - Local Rent Calc"
End Sub

Private Function CanonAmiKey(ami As Double) As String
    CanonAmiKey = Format$(Round(CDbl(ami), 4), "0.0000")
End Function

Private Function BedroomLabelFromCount(bedrooms As Variant) As String
    Dim n As Long
    n = 0
    On Error Resume Next
    If IsNumeric(bedrooms) Then n = CLng(Round(CDbl(bedrooms), 0))
    On Error GoTo 0

    If n <= 0 Then
        BedroomLabelFromCount = "studio"
    ElseIf n >= 5 Then
        BedroomLabelFromCount = "5 BR"
    Else
        BedroomLabelFromCount = CStr(n) & " BR"
    End If
End Function

Private Function BedroomLabelFromSheetLabel(label As String) As String
    Dim v As String
    v = LCase$(Trim$(CStr(label)))

    Select Case v
        Case "studio": BedroomLabelFromSheetLabel = "studio"
        Case "1 br": BedroomLabelFromSheetLabel = "1 BR"
        Case "2 br": BedroomLabelFromSheetLabel = "2 BR"
        Case "3 br": BedroomLabelFromSheetLabel = "3 BR"
        Case "4 br": BedroomLabelFromSheetLabel = "4 BR"
        Case "5 br": BedroomLabelFromSheetLabel = "5 BR"
        Case Else: BedroomLabelFromSheetLabel = ""
    End Select
End Function

Private Function UtilityCategoryFromHeaderValue(headerValue As String) As String
    Dim h As String
    h = LCase$(Trim$(CStr(headerValue)))

    Select Case h
        Case "apartment electricity only": UtilityCategoryFromHeaderValue = "electricity"
        Case "cooking": UtilityCategoryFromHeaderValue = "cooking"
        Case "heat": UtilityCategoryFromHeaderValue = "heat"
        Case "hot water": UtilityCategoryFromHeaderValue = "hot_water"
        Case Else: UtilityCategoryFromHeaderValue = UtilityCategoryFromOptionLabel(headerValue)
    End Select
End Function

Private Function UtilityCategoryFromOptionLabel(optionLabel As String) As String
    Dim v As String
    v = Trim$(CStr(optionLabel))

    Select Case v
        Case "Tenant Pays": UtilityCategoryFromOptionLabel = "electricity"
        Case "Electric Stove", "Gas Stove": UtilityCategoryFromOptionLabel = "cooking"
        Case "Electric Heat - Cold Climate Air Source Heat Pump (ccASHP)1", "Electric Heat - Other2", "Gas Heat", "Oil Heat": UtilityCategoryFromOptionLabel = "heat"
        Case "Electric Hot Water - Heat Pump", "Electric Hot Water - Other", "Gas Hot Water", "Oil Hot Water": UtilityCategoryFromOptionLabel = "hot_water"
        Case Else: UtilityCategoryFromOptionLabel = "" ' N/A label is ambiguous; rely on current category from header.
    End Select
End Function

Private Function LoadLocalRentLookups(rentWs As Worksheet, ByRef tradeoffs As Collection) As Boolean
    ' Parses the "AMI & Rent" sheet similarly to ami_optix/rent_calculator.py.
    On Error GoTo Fail

    Set m_LocalGrossRents = CreateObject("Scripting.Dictionary")
    Set m_LocalAllowances = CreateObject("Scripting.Dictionary")

    ' --- Gross rent table ---
    Dim lastRow As Long
    lastRow = rentWs.Cells(rentWs.Rows.Count, 3).End(xlUp).row ' col C
    If lastRow < 1 Then lastRow = 1

    Dim currentAmi As Double
    currentAmi = -1#

    Dim r As Long
    For r = 1 To Application.Min(lastRow, 5000)
        Dim cVal As Variant
        cVal = rentWs.Cells(r, 3).Value

        Dim marker As String
        marker = LCase$(Trim$(CStr(rentWs.Cells(r, 4).Value)))

        If IsNumeric(cVal) And marker = "of ami" Then
            currentAmi = CDbl(cVal)
        ElseIf currentAmi >= 0# Then
            Dim bedLabel As String
            bedLabel = BedroomLabelFromSheetLabel(CStr(cVal))
            If bedLabel <> "" Then
                Dim grossVal As Variant
                grossVal = rentWs.Cells(r, 7).Value ' col G
                If IsNumeric(grossVal) Then
                    m_LocalGrossRents(CanonAmiKey(currentAmi) & "|" & bedLabel) = CDbl(grossVal)
                End If
            End If
        End If
    Next r

    ' --- Allowances table ---
    Dim currentCat As String
    currentCat = ""

    Dim lastCol As Long
    lastCol = rentWs.Cells(15, rentWs.Columns.Count).End(xlToLeft).Column
    Dim lastCol16 As Long
    lastCol16 = rentWs.Cells(16, rentWs.Columns.Count).End(xlToLeft).Column
    If lastCol16 > lastCol Then lastCol = lastCol16
    If lastCol < 1 Then lastCol = 1

    ' Detect layout variant (same logic as RentTables cache builder)
    Dim hasCategoryHeaders As Boolean
    hasCategoryHeaders = False
    Dim probeCol As Long
    For probeCol = 1 To Application.Min(lastCol, 200)
        Dim probeVal As String
        probeVal = LCase$(Trim$(CStr(rentWs.Cells(15, probeCol).Value)))
        Select Case probeVal
            Case "apartment electricity only", "cooking", "heat", "hot water"
                hasCategoryHeaders = True
                Exit For
        End Select
    Next probeCol

    Dim firstDataRow As Long
    If hasCategoryHeaders Then
        firstDataRow = 18
    Else
        firstDataRow = 17
    End If

    Dim col As Long
    For col = 1 To Application.Min(lastCol, 200)
        Dim headerVal As String
        headerVal = Trim$(CStr(rentWs.Cells(15, col).Value))

        Dim headerCat As String
        headerCat = UtilityCategoryFromHeaderValue(headerVal)
        If headerCat <> "" Then currentCat = headerCat

        Dim optionVal As Variant
        If (Not hasCategoryHeaders) And UtilityCategoryFromOptionLabel(headerVal) <> "" Then
            optionVal = headerVal
        Else
            optionVal = rentWs.Cells(16, col).Value
            If Trim$(CStr(optionVal)) = "" Then optionVal = rentWs.Cells(17, col).Value
        End If

        Dim optionLabel As String
        optionLabel = Trim$(CStr(optionVal))
        If optionLabel = "" Then GoTo NextCol
        If LCase$(optionLabel) = "select -->>" Then GoTo NextCol

        Dim optionCat As String
        optionCat = UtilityCategoryFromOptionLabel(optionLabel)
        If optionCat = "" Then optionCat = currentCat
        If optionCat = "" Then GoTo NextCol

        Dim catDict As Object
        Set catDict = Nothing
        If m_LocalAllowances.Exists(optionCat) Then
            Set catDict = m_LocalAllowances(optionCat)
        Else
            Set catDict = CreateObject("Scripting.Dictionary")
            m_LocalAllowances(optionCat) = catDict
        End If

        Dim bedDict As Object
        Set bedDict = CreateObject("Scripting.Dictionary")

        Dim bedLabels As Variant
        bedLabels = Array("studio", "1 BR", "2 BR", "3 BR", "4 BR", "5 BR")

        Dim i As Long
        For i = LBound(bedLabels) To UBound(bedLabels)
            Dim amtVal As Variant
            amtVal = rentWs.Cells(firstDataRow + i, col).Value

            Dim amt As Double
            amt = 0#
            If IsNumeric(amtVal) Then
                amt = CDbl(amtVal)
            ElseIf Len(Trim$(CStr(amtVal))) > 0 Then
                On Error Resume Next
                amt = CDbl(amtVal)
                On Error GoTo Fail
            End If
            bedDict(CStr(bedLabels(i))) = amt
        Next i

        catDict(optionLabel) = bedDict

NextCol:
    Next col

    If m_LocalGrossRents Is Nothing Then
        LoadLocalRentLookups = False
    Else
        LoadLocalRentLookups = (m_LocalGrossRents.Count > 0)
    End If
    If Not LoadLocalRentLookups Then
        tradeoffs.Add "Local rent calc: could not parse gross rent table from 'AMI & Rent'."
    End If
    Exit Function

Fail:
    tradeoffs.Add "Local rent parse failed: " & Err.Description
    LoadLocalRentLookups = False
End Function

Private Function LookupLocalGrossRent(ami As Double, bedroomLabel As String, ByRef gross As Double) As Boolean
    gross = 0#
    LookupLocalGrossRent = False

    If m_LocalGrossRents Is Nothing Then Exit Function

    Dim key As String
    key = CanonAmiKey(ami) & "|" & bedroomLabel

    If m_LocalGrossRents.Exists(key) Then
        gross = CDbl(m_LocalGrossRents(key))
        LookupLocalGrossRent = True
    End If
End Function

Private Function TryLookupLocalAllowance(category As String, optionLabel As String, bedroomLabel As String, ByRef amount As Double) As Boolean
    On Error GoTo SafeExit

    amount = 0#
    TryLookupLocalAllowance = False

    If m_LocalAllowances Is Nothing Then Exit Function
    If Not m_LocalAllowances.Exists(category) Then Exit Function

    Dim catDict As Object
    Set catDict = m_LocalAllowances(category)
    If catDict Is Nothing Then Exit Function
    If Not catDict.Exists(optionLabel) Then Exit Function

    Dim bedDict As Object
    Set bedDict = catDict(optionLabel)
    If bedDict Is Nothing Then Exit Function
    If Not bedDict.Exists(bedroomLabel) Then Exit Function

    amount = CDbl(bedDict(bedroomLabel))
    TryLookupLocalAllowance = True
    Exit Function

SafeExit:
    amount = 0#
    TryLookupLocalAllowance = False
End Function

Private Function EnrichAssignmentsWithLocalRents(assignments As Collection, utilities As Object, ByRef tradeoffs As Collection) As Object
    ' Returns rent_totals dict (net_monthly/net_annual).
    On Error GoTo Fail

    Dim totals As Object
    Set totals = CreateObject("Scripting.Dictionary")

    Dim totalNet As Double
    totalNet = 0#

    Dim categories As Variant
    categories = Array("electricity", "cooking", "heat", "hot_water")

    Dim i As Long
    For i = 1 To assignments.Count
        Dim a As Object
        Set a = assignments(i)
        If a Is Nothing Then GoTo NextAssignment

        Dim unitId As String
        unitId = ""
        On Error Resume Next
        If a.Exists("unit_id") Then unitId = CStr(a("unit_id"))
        On Error GoTo Fail

        Dim ami As Double
        ami = 0#
        If a.Exists("assigned_ami") Then ami = CDbl(a("assigned_ami"))
        If ami > 2# Then ami = ami / 100#
        a("assigned_ami") = ami

        Dim bedLabel As String
        bedLabel = BedroomLabelFromCount(a("bedrooms"))

        Dim gross As Double
        gross = 0#
        If Not LookupLocalGrossRent(ami, bedLabel, gross) Then
            Err.Raise LOCAL_RENT_ERR_GROSS, "AMI_Optix_ResultsWriter.LocalRentCalc", _
                      "Local rent calc blocked (missing gross rent lookup)." & vbCrLf & vbCrLf & _
                      "Year: " & CStr(m_LocalRentYear) & vbCrLf & _
                      "Workbook: " & m_LocalRentPath & vbCrLf & _
                      "Sheet: AMI & Rent" & vbCrLf & _
                      "Unit: " & unitId & vbCrLf & _
                      "AMI: " & Format$(ami, "0.00%") & " (" & CanonAmiKey(ami) & ")" & vbCrLf & _
                      "Bedrooms: " & bedLabel & vbCrLf & _
                      "Key: " & CanonAmiKey(ami) & "|" & bedLabel & vbCrLf & _
                      "Lookup: gross table (col C labels, col D marker 'of AMI', col G gross)." & vbCrLf & _
                      "Fingerprint: " & m_LocalRentFingerprint
        End If

        Dim allowancesArr As Collection
        Set allowancesArr = New Collection

        Dim totalAllowance As Double
        totalAllowance = 0#

        Dim c As Long
        For c = LBound(categories) To UBound(categories)
            Dim cat As String
            cat = CStr(categories(c))

            Dim selectionKey As String
            selectionKey = "na"
            On Error Resume Next
            If Not utilities Is Nothing Then
                If utilities.Exists(cat) Then selectionKey = CStr(utilities(cat))
            End If
            On Error GoTo Fail

            Dim optionLabel As String
            optionLabel = FormatUtilityType(selectionKey, cat)

            Dim amt As Double
            amt = 0#
            If LCase$(Trim$(optionLabel)) <> "n/a or owner pays" Then
                If Not TryLookupLocalAllowance(cat, optionLabel, bedLabel, amt) Then
                    Err.Raise LOCAL_RENT_ERR_ALLOWANCE, "AMI_Optix_ResultsWriter.LocalRentCalc", _
                              "Local rent calc blocked (missing utility allowance lookup)." & vbCrLf & vbCrLf & _
                              "Year: " & CStr(m_LocalRentYear) & vbCrLf & _
                              "Workbook: " & m_LocalRentPath & vbCrLf & _
                              "Sheet: AMI & Rent" & vbCrLf & _
                              "Unit: " & unitId & vbCrLf & _
                              "Category: " & cat & vbCrLf & _
                              "Option: " & optionLabel & vbCrLf & _
                              "Bedrooms: " & bedLabel & vbCrLf & _
                              "Key: " & cat & " | " & optionLabel & " | " & bedLabel & vbCrLf & _
                              "Lookup: allowances table (row 15 headers, row 16/17 options, rows 18-23 values)." & vbCrLf & _
                              "Fingerprint: " & m_LocalRentFingerprint
                End If
            End If

            Dim item As Object
            Set item = CreateObject("Scripting.Dictionary")
            item("category") = cat
            item("label") = optionLabel
            item("amount") = Round(amt, 2)
            allowancesArr.Add item

            totalAllowance = totalAllowance + amt
        Next c

        Dim net As Double
        net = gross - totalAllowance
        If net < 0# Then net = 0#

        a("gross_rent") = Round(gross, 2)
        a("monthly_rent") = Round(net, 2)
        a("annual_rent") = Round(net * 12#, 2)
        a("allowance_total") = Round(totalAllowance, 2)
        Set a("allowances") = allowancesArr

        totalNet = totalNet + net

NextAssignment:
    Next i

    totals("net_monthly") = Round(totalNet, 2)
    totals("net_annual") = Round(totalNet * 12#, 2)

    Set EnrichAssignmentsWithLocalRents = totals
    Exit Function

Fail:
    If Err.Number = LOCAL_RENT_ERR_LAYOUT Or Err.Number = LOCAL_RENT_ERR_GROSS Or Err.Number = LOCAL_RENT_ERR_ALLOWANCE Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

    Err.Raise LOCAL_RENT_ERR_UNEXPECTED, "AMI_Optix_ResultsWriter.LocalRentCalc", _
              "Local rent calc failed unexpectedly." & vbCrLf & vbCrLf & _
              "Year: " & CStr(m_LocalRentYear) & vbCrLf & _
              "Workbook: " & m_LocalRentPath & vbCrLf & _
              "Sheet: AMI & Rent" & vbCrLf & _
              "Fingerprint: " & m_LocalRentFingerprint & vbCrLf & _
              "Error: " & Err.Description
End Function

Private Function ComputeWaami(assignments As Collection) As Double
    On Error GoTo SafeExit

    Dim totalSf As Double
    Dim weighted As Double
    totalSf = 0#: weighted = 0#

    Dim i As Long
    For i = 1 To assignments.Count
        Dim a As Object
        Set a = assignments(i)
        If a Is Nothing Then GoTo NextA

        Dim sf As Double
        sf = 0#
        If a.Exists("net_sf") Then
            If IsNumeric(a("net_sf")) Then sf = CDbl(a("net_sf"))
        End If

        Dim ami As Double
        ami = 0#
        If a.Exists("assigned_ami") Then ami = CDbl(a("assigned_ami"))
        If ami > 2# Then ami = ami / 100#

        totalSf = totalSf + sf
        weighted = weighted + (sf * ami)

NextA:
    Next i

    If totalSf <= 0# Then
        ComputeWaami = 0#
    Else
        ComputeWaami = weighted / totalSf
    End If
    Exit Function

SafeExit:
    ComputeWaami = 0#
End Function

Private Function ComputeBandsUsed(assignments As Collection) As Collection
    Dim uniq As Object
    Set uniq = CreateObject("Scripting.Dictionary") ' band (as long percent) -> True

    Dim i As Long
    For i = 1 To assignments.Count
        Dim a As Object
        Set a = assignments(i)
        If a Is Nothing Then GoTo NextA

        Dim ami As Double
        ami = 0#
        If a.Exists("assigned_ami") Then ami = CDbl(a("assigned_ami"))
        If ami > 2# Then ami = ami / 100#

        Dim band As Long
        band = CLng(Application.Round(ami * 100#, 0))
        uniq(CStr(band)) = True

NextA:
    Next i

    Dim keys As Variant
    keys = uniq.keys

    ' Sort numeric ascending (simple bubble sort; small N)
    Dim j As Long, k As Long
    For j = LBound(keys) To UBound(keys) - 1
        For k = j + 1 To UBound(keys)
            If CLng(keys(k)) < CLng(keys(j)) Then
                Dim tmp As Variant
                tmp = keys(j)
                keys(j) = keys(k)
                keys(k) = tmp
            End If
        Next k
    Next j

    Dim bands As Collection
    Set bands = New Collection

    For j = LBound(keys) To UBound(keys)
        If Len(Trim$(CStr(keys(j)))) > 0 Then
            bands.Add CLng(keys(j))
        End If
    Next j

    Set ComputeBandsUsed = bands
End Function

Private Function BuildBandMix(assignments As Collection, Optional totalBuildingSf As Double = 0#) As Collection
    Dim totalSf As Double
    totalSf = 0#

    Dim bandAgg As Object
    Set bandAgg = CreateObject("Scripting.Dictionary") ' band -> dict(units, net_sf)

    Dim i As Long
    For i = 1 To assignments.Count
        Dim a As Object
        Set a = assignments(i)
        If a Is Nothing Then GoTo NextA

        Dim sf As Double
        sf = 0#
        If a.Exists("net_sf") Then
            If IsNumeric(a("net_sf")) Then sf = CDbl(a("net_sf"))
        End If

        Dim ami As Double
        ami = 0#
        If a.Exists("assigned_ami") Then ami = CDbl(a("assigned_ami"))
        If ami > 2# Then ami = ami / 100#

        Dim band As Long
        band = CLng(Application.Round(ami * 100#, 0))

        Dim key As String
        key = CStr(band)

        Dim agg As Object
        If bandAgg.Exists(key) Then
            Set agg = bandAgg(key)
        Else
            Set agg = CreateObject("Scripting.Dictionary")
            agg("units") = 0
            agg("net_sf") = 0#
            bandAgg(key) = agg
        End If

        agg("units") = CLng(agg("units")) + 1
        agg("net_sf") = CDbl(agg("net_sf")) + sf
        totalSf = totalSf + sf

NextA:
    Next i

    Dim keys As Variant
    keys = bandAgg.keys

    ' Sort numeric ascending
    Dim j As Long, k As Long
    For j = LBound(keys) To UBound(keys) - 1
        For k = j + 1 To UBound(keys)
            If CLng(keys(k)) < CLng(keys(j)) Then
                Dim tmp As Variant
                tmp = keys(j)
                keys(j) = keys(k)
                keys(k) = tmp
            End If
        Next k
    Next j

    Dim mix As Collection
    Set mix = New Collection

    For j = LBound(keys) To UBound(keys)
        Dim agg2 As Object
        Set agg2 = bandAgg(CStr(keys(j)))

        Dim bm As Object
        Set bm = CreateObject("Scripting.Dictionary")
        bm("band") = CLng(keys(j))
        bm("units") = CLng(agg2("units"))
        bm("net_sf") = CDbl(agg2("net_sf"))
        If totalSf > 0# Then
            bm("share_of_sf") = CDbl(agg2("net_sf")) / totalSf
        Else
            bm("share_of_sf") = 0#
        End If
        If totalBuildingSf > 0# Then
            bm("share_of_building_sf") = CDbl(agg2("net_sf")) / totalBuildingSf
        End If
        mix.Add bm
    Next j

    Set BuildBandMix = mix
End Function

Private Sub ClearManualBlock(ws As Worksheet)
    ' Clears only the top "Scenario Manual" region without wiping the scenarios below.
    Dim firstScenarioRow As Long
    firstScenarioRow = FindFirstScenarioHeaderRow(ws)

    Dim clearToRow As Long
    If firstScenarioRow > 0 Then
        clearToRow = Application.Max(1, firstScenarioRow - 1)
    Else
        clearToRow = MANUAL_CLEAR_FALLBACK_HEIGHT
    End If

    ws.Range("A1:M" & clearToRow).Clear
End Sub

Private Function FindFirstScenarioHeaderRow(ws As Worksheet) As Long
    On Error GoTo Fail

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    If lastRow < 1 Then lastRow = 1

    Dim r As Long
    For r = 1 To Application.Min(lastRow, 5000)
        Dim v As String
        v = UCase$(Trim$(CStr(ws.Cells(r, 1).Value)))
        ' Match scenario table headers like "SCENARIO 1: ..." (not "SCENARIO MANUAL ...").
        If (v Like "SCENARIO [0-9]*") Or (v Like "SCENARIO #[0-9]*") Then
            FindFirstScenarioHeaderRow = r
            Exit Function
        End If
    Next r

Fail:
    FindFirstScenarioHeaderRow = 0
End Function

Private Function WriteManualScenarioBlockFromResult(ws As Worksheet, result As Object) As Long
    ClearManualBlock ws

    Dim row As Long
    row = MANUAL_BLOCK_START_ROW

    ws.Cells(row, 1).Value = "AMI OPTIMIZATION RESULTS"
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 1).Font.Size = 16
    row = row + 2

    Dim scenarioKey As String
    scenarioKey = GetBestScenarioKey(result)

    Dim scenarios As Object
    Set scenarios = Nothing
    On Error Resume Next
    Set scenarios = result("scenarios")
    On Error GoTo 0

    Dim scenario As Object
    Set scenario = Nothing
    If Not scenarios Is Nothing Then
        On Error Resume Next
        If Trim$(scenarioKey) <> "" Then Set scenario = scenarios(scenarioKey)
        On Error GoTo 0
    End If

    row = WriteMihSquareFootageSummary(ws, row, scenario)
    row = WriteUtilitySettings(ws, row)
    row = WriteUtilityDeductionTotalsByBedroom(ws, row, scenario)
    row = row + 1

    Dim headerLabel As String
    headerLabel = "SCENARIO MANUAL (LIVE SYNC)"
    If Trim$(scenarioKey) <> "" Then
        headerLabel = headerLabel & " - CURRENT: " & UCase$(FormatScenarioName(CStr(scenarioKey)))
    End If

    ws.Cells(row, 1).Value = headerLabel
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 1).Font.Size = 14
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 13)).Interior.Color = RGB(220, 240, 220)
    row = row + 1
    If scenarioKey = "" Then
        WriteManualScenarioBlockFromResult = row
        Exit Function
    End If

    If scenarios Is Nothing Then
        WriteManualScenarioBlockFromResult = row
        Exit Function
    End If
    If scenario Is Nothing Then
        WriteManualScenarioBlockFromResult = row
        Exit Function
    End If

    row = WriteScenarioSummaryAndTable(ws, row, scenario)
    WriteManualScenarioBlockFromResult = row
End Function

Private Function WriteMihSquareFootageSummary(ws As Worksheet, startRow As Long, scenario As Object) As Long
    ' Writes total building SF and affordable SF totals at the top of the
    ' MIH results sheet. Per-AMI-band breakdown lives in the Scenario Manual
    ' "Band Mix" table below; do not duplicate it here.
    Dim row As Long
    Dim metrics As Object
    Dim bandMix As Object
    Dim totalSf As Double
    Dim buildingSf As Double
    Dim idx As Long
    Dim bm As Object

    row = startRow
    WriteMihSquareFootageSummary = row

    If scenario Is Nothing Then Exit Function

    Set metrics = Nothing
    On Error Resume Next
    Set metrics = scenario("metrics")
    On Error GoTo 0
    If metrics Is Nothing Then Exit Function

    Set bandMix = Nothing
    On Error Resume Next
    Set bandMix = metrics("band_mix")
    On Error GoTo 0
    If bandMix Is Nothing Then Exit Function

    ' Calculate total building SF from band_mix
    totalSf = 0#
    For idx = 1 To bandMix.Count
        Set bm = bandMix(idx)
        If Not bm Is Nothing Then
            If bm.Exists("net_sf") Then totalSf = totalSf + CDbl(bm("net_sf"))
        End If
    Next idx
    If totalSf <= 0# Then Exit Function

    ' Use actual building SF for denominator; fall back to affordable total.
    buildingSf = g_MihTotalBuildingSf
    If buildingSf <= 0# Then buildingSf = totalSf

    ' Section header
    ws.Cells(row, 1).Value = "SQUARE FOOTAGE SUMMARY"
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 1).Font.Size = 13
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 4)).Interior.Color = RGB(230, 245, 255)
    row = row + 1

    ' Total building SF (actual building, not just affordable units)
    ws.Cells(row, 1).Value = "Total Building Net SF:"
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 2).Value = buildingSf
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    ws.Cells(row, 2).Font.Bold = True
    row = row + 1

    ' Affordable-only SF subtotal
    ws.Cells(row, 1).Value = "Affordable Net SF:"
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 2).Value = totalSf
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    ws.Cells(row, 2).Font.Bold = True
    row = row + 1

    row = row + 1
    WriteMihSquareFootageSummary = row
End Function

Private Function WriteManualScenarioBlockFromEvaluate(ws As Worksheet, evalResult As Object, Optional headerLabel As String = "SCENARIO MANUAL (LIVE SYNC)") As Long
    ClearManualBlock ws

    Dim row As Long
    row = MANUAL_BLOCK_START_ROW

    ws.Cells(row, 1).Value = "AMI OPTIMIZATION RESULTS"
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 1).Font.Size = 16
    row = row + 2

    ' Build a minimal scenario-shaped object from /api/evaluate response.
    Dim scenario As Object
    Set scenario = CreateObject("Scripting.Dictionary")

    If evalResult.Exists("summary") Then
        Dim summary As Object
        Set summary = evalResult("summary")
        If summary.Exists("waami") Then scenario("waami") = summary("waami")
        If summary.Exists("bands_used") Then
            If IsObject(summary("bands_used")) Then
                Set scenario("bands") = summary("bands_used")
            Else
                scenario("bands") = summary("bands_used")
            End If
        End If
    End If
    If evalResult.Exists("metrics") Then
        If IsObject(evalResult("metrics")) Then
            Set scenario("metrics") = evalResult("metrics")
        Else
            scenario("metrics") = evalResult("metrics")
        End If
    End If
    If evalResult.Exists("tradeoffs") Then
        If IsObject(evalResult("tradeoffs")) Then
            Set scenario("tradeoffs") = evalResult("tradeoffs")
        Else
            scenario("tradeoffs") = evalResult("tradeoffs")
        End If
    End If
    If evalResult.Exists("assignments") Then
        If IsObject(evalResult("assignments")) Then
            Set scenario("assignments") = evalResult("assignments")
        Else
            scenario("assignments") = evalResult("assignments")
        End If
    End If
    If evalResult.Exists("rent_totals") And Not IsNull(evalResult("rent_totals")) Then
        If IsObject(evalResult("rent_totals")) Then
            Set scenario("rent_totals") = evalResult("rent_totals")
        Else
            scenario("rent_totals") = evalResult("rent_totals")
        End If
    End If

    row = WriteMihSquareFootageSummary(ws, row, scenario)
    row = WriteUtilitySettings(ws, row)
    row = WriteUtilityDeductionTotalsByBedroom(ws, row, scenario)
    row = row + 1

    ws.Cells(row, 1).Value = headerLabel
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 1).Font.Size = 14
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 13)).Interior.Color = RGB(220, 240, 220)
    row = row + 1

    row = WriteScenarioSummaryAndTable(ws, row, scenario)
    WriteManualScenarioBlockFromEvaluate = row
End Function

'-------------------------------------------------------------------------------
' SOLVER SCENARIO RENT RECALCULATION
'-------------------------------------------------------------------------------

Private Sub RecalculateSolverScenarioRents(ws As Worksheet, programNorm As String, _
    utilities As Object, mihOption As String, mihResidentialSF As Double, _
    mihMaxBandPercent As Long)
    ' Recalculates rents for all solver scenario blocks on the AMI Scenarios sheet
    ' using the currently selected rent roll year.  Keeps AMI assignments intact;
    ' only gross rent, net rent, annual rent, and totals are updated.
    On Error GoTo Cleanup

    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 10 Then Exit Sub

    ' Phase 1: Find all solver-scenario table header rows.
    ' A table header has: col1="Unit", col5="AMI", col6="Gross Rent".
    ' We skip the Manual block's table by only considering rows that appear
    ' after the first "SCENARIO N:" header that is NOT a MANUAL header.
    Dim tableHeaderRows As Collection
    Set tableHeaderRows = New Collection

    Dim r As Long
    Dim inSolverSection As Boolean
    inSolverSection = False
    Dim cellA As String

    For r = 1 To lastRow
        cellA = Trim$(CStr(ws.Cells(r, 1).Value))
        If Left$(cellA, 9) = "SCENARIO " And InStr(1, cellA, "MANUAL", vbTextCompare) = 0 Then
            inSolverSection = True
        End If
        If inSolverSection And cellA = "Unit" And _
           Trim$(CStr(ws.Cells(r, 5).Value)) = "AMI" And _
           Trim$(CStr(ws.Cells(r, 6).Value)) = "Gross Rent" Then
            tableHeaderRows.Add r
        End If
    Next r

    If tableHeaderRows.Count = 0 Then Exit Sub

    ' Phase 2: For each solver table, read assignments, call API, update rents.
    Dim tIdx As Long
    Dim headerRow As Long
    Dim dataRow As Long
    Dim units As Collection
    Dim unit As Object
    Dim unitVal As Variant
    Dim bedroomsVal As Variant
    Dim sfVal As Variant
    Dim rawAmi As Variant
    Dim amiVal As Double
    Dim payload As String
    Dim response As String
    Dim evalResult As Object
    Dim apiAssignments As Object
    Dim apiAssign As Object
    Dim rentTotals As Object
    Dim assignRow As Long
    Dim scanRow As Long
    Dim scanVal As String
    Dim maxA As Long
    Dim a As Long

    For tIdx = 1 To tableHeaderRows.Count
        headerRow = tableHeaderRows(tIdx)

        Application.StatusBar = "Recalculating rents for scenario " & tIdx & " of " & tableHeaderRows.Count & "..."

        ' -- Read assignments from the table --
        Set units = New Collection
        dataRow = headerRow + 1

        Do While dataRow <= lastRow
            unitVal = ws.Cells(dataRow, 1).Value
            If IsEmpty(unitVal) Or Trim$(CStr(unitVal)) = "" Then Exit Do

            Set unit = CreateObject("Scripting.Dictionary")
            unit("unit_id") = CStr(unitVal)

            bedroomsVal = ws.Cells(dataRow, 2).Value
            If IsNumeric(bedroomsVal) Then
                unit("bedrooms") = CDbl(bedroomsVal)
            Else
                unit("bedrooms") = 0
            End If

            sfVal = ws.Cells(dataRow, 3).Value
            If IsNumeric(sfVal) Then
                unit("net_sf") = CDbl(sfVal)
            Else
                unit("net_sf") = 0
            End If

            amiVal = 0#
            rawAmi = ws.Cells(dataRow, 5).Value
            If IsNumeric(rawAmi) Then amiVal = CDbl(rawAmi)
            unit("client_ami") = amiVal

            units.Add unit
            dataRow = dataRow + 1
        Loop

        If units.Count = 0 Then GoTo NextTable

        ' -- Call API for this scenario's assignments --
        payload = BuildEvaluatePayloadV2(units, utilities, programNorm, _
            mihOption, mihResidentialSF, mihMaxBandPercent)

        response = CallManualCalculateAPI(payload)
        If response = "" Then GoTo NextTable

        Set evalResult = ParseJSON(response)
        If evalResult Is Nothing Then GoTo NextTable

        ' -- Update rent cells in place --
        If evalResult.Exists("assignments") Then
            Set apiAssignments = evalResult("assignments")
            maxA = apiAssignments.Count
            If maxA > units.Count Then maxA = units.Count

            For a = 1 To maxA
                assignRow = headerRow + a
                Set apiAssign = apiAssignments(a)

                If apiAssign.Exists("gross_rent") Then
                    ws.Cells(assignRow, 6).Value = CDbl(apiAssign("gross_rent"))
                    ws.Cells(assignRow, 6).NumberFormat = "$#,##0"
                End If
                If apiAssign.Exists("monthly_rent") Then
                    ws.Cells(assignRow, 7).Value = CDbl(apiAssign("monthly_rent"))
                    ws.Cells(assignRow, 7).NumberFormat = "$#,##0"
                End If
                If apiAssign.Exists("annual_rent") Then
                    ws.Cells(assignRow, 8).Value = CDbl(apiAssign("annual_rent"))
                    ws.Cells(assignRow, 8).NumberFormat = "$#,##0"
                End If
            Next a
        End If

        ' -- Update rent totals (scan backward from table header) --
        If evalResult.Exists("rent_totals") Then
            Set rentTotals = Nothing
            On Error Resume Next
            Set rentTotals = evalResult("rent_totals")
            On Error GoTo Cleanup

            If Not rentTotals Is Nothing Then
                For scanRow = headerRow - 1 To Application.Max(1, headerRow - 20) Step -1
                    scanVal = Trim$(CStr(ws.Cells(scanRow, 1).Value))
                    If scanVal = "Total Monthly Rent:" Then
                        If rentTotals.Exists("net_monthly") Then
                            ws.Cells(scanRow, 2).Value = CDbl(rentTotals("net_monthly"))
                            ws.Cells(scanRow, 2).NumberFormat = "$#,##0"
                        End If
                    ElseIf scanVal = "Total Annual Rent:" Then
                        If rentTotals.Exists("net_annual") Then
                            ws.Cells(scanRow, 2).Value = CDbl(rentTotals("net_annual"))
                            ws.Cells(scanRow, 2).NumberFormat = "$#,##0"
                        End If
                    End If
                Next scanRow
            End If
        End If

NextTable:
    Next tIdx

Cleanup:
    Application.StatusBar = False
End Sub

'-------------------------------------------------------------------------------
' MANUAL BLOCK REFRESH (FROM A KNOWN SCENARIO)
'-------------------------------------------------------------------------------

Public Sub RefreshManualScenarioFromScenario(scenarioKey As String, scenario As Object)
    ' Refreshes the "Scenario Manual (LIVE SYNC)" block using a scenario object already returned by the API.
    ' This avoids an extra /api/evaluate call and guarantees the manual block matches the applied scenario.
    On Error GoTo Fail

    Dim ws As Worksheet
    Set ws = GetOrCreateScenariosSheet()

    Dim prevEnableEvents As Boolean
    Dim prevScreenUpdating As Boolean
    prevEnableEvents = Application.EnableEvents
    prevScreenUpdating = Application.ScreenUpdating

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ClearManualBlock ws

    Dim row As Long
    row = MANUAL_BLOCK_START_ROW

    ws.Cells(row, 1).Value = "AMI OPTIMIZATION RESULTS"
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 1).Font.Size = 16
    row = row + 2

    row = WriteUtilitySettings(ws, row)
    row = WriteUtilityDeductionTotalsByBedroom(ws, row, scenario)
    row = row + 1

    ws.Cells(row, 1).Value = "SCENARIO MANUAL (LIVE SYNC) - CURRENT: " & UCase$(FormatScenarioName(CStr(scenarioKey)))
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 1).Font.Size = 14
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 13)).Interior.Color = RGB(220, 240, 220)
    row = row + 1

    If scenario Is Nothing Then GoTo Cleanup
    row = WriteScenarioSummaryAndTable(ws, row, scenario)

Cleanup:
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    Exit Sub

Fail:
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Private Function GetBestScenarioKey(result As Object) As String
    Dim scenarios As Object
    If result Is Nothing Then Exit Function
    If Not result.Exists("scenarios") Then Exit Function
    Set scenarios = result("scenarios")

    Dim priorities As Variant
    priorities = Array("absolute_best", "best_3_band", "best_2_band", "alternative")

    Dim i As Long
    For i = LBound(priorities) To UBound(priorities)
        If scenarios.Exists(CStr(priorities(i))) Then
            GetBestScenarioKey = CStr(priorities(i))
            Exit Function
        End If
    Next i
End Function

Private Function WriteScenarioSummaryAndTable(ws As Worksheet, startRow As Long, scenario As Object) As Long
    Dim row As Long
    row = startRow
    Dim tableHeaderRow As Long

    If scenario Is Nothing Then
        WriteScenarioSummaryAndTable = row
        Exit Function
    End If

    If scenario.Exists("waami") Then
        ws.Cells(row, 1).Value = "WAAMI:"
        ws.Cells(row, 2).Value = Format(scenario("waami"), "0.00%")
        ws.Cells(row, 2).Font.Bold = True
        row = row + 1
    End If

    If scenario.Exists("bands") Then
        ws.Cells(row, 1).Value = "Bands Used:"
        Dim bandStr As String
        bandStr = ""

        ' IMPORTANT: scenario("bands") is usually a Collection. Never assign it to a Variant without Set,
        ' or VBA will try to call the default property (Item) and raise "Wrong number of arguments...".
        Dim bandsObj As Object
        Set bandsObj = Nothing
        On Error Resume Next
        Set bandsObj = scenario("bands")
        On Error GoTo 0

        If Not bandsObj Is Nothing Then
            Dim b As Long
            If TypeName(bandsObj) = "Collection" Then
                For b = 1 To bandsObj.Count
                    If bandStr <> "" Then bandStr = bandStr & ", "
                    bandStr = bandStr & Format(bandsObj(b), "0") & "%"
                Next b
            ElseIf TypeName(bandsObj) = "Dictionary" Then
                Dim bandKey As Variant
                For Each bandKey In bandsObj.Keys
                    If bandStr <> "" Then bandStr = bandStr & ", "
                    bandStr = bandStr & Format(bandsObj(bandKey), "0") & "%"
                Next bandKey
            End If
        Else
            ' Fallback: array or scalar
            Dim bandsVal As Variant
            bandsVal = scenario("bands")
            If IsArray(bandsVal) Then
                Dim i As Long
                For i = LBound(bandsVal) To UBound(bandsVal)
                    If bandStr <> "" Then bandStr = bandStr & ", "
                    bandStr = bandStr & Format(bandsVal(i), "0") & "%"
                Next i
            ElseIf Len(Trim(CStr(bandsVal))) > 0 Then
                bandStr = CStr(bandsVal)
            End If
        End If

        ws.Cells(row, 2).Value = bandStr
        row = row + 1
    End If

    ' Tradeoffs (edge scenarios): short, client-readable list of what was relaxed/violated.
    If scenario.Exists("tradeoffs") Then
        Dim tradeoffsObj As Object
        Set tradeoffsObj = Nothing
        On Error Resume Next
        Set tradeoffsObj = scenario("tradeoffs")
        On Error GoTo 0

        If Not tradeoffsObj Is Nothing Then
            If TypeName(tradeoffsObj) = "Collection" Then
                If tradeoffsObj.Count > 0 Then
                    ws.Cells(row, 1).Value = "Tradeoffs:"
                    ws.Cells(row, 1).Font.Bold = True
                    row = row + 1

                    Dim t As Long
                    For t = 1 To Application.Min(4, tradeoffsObj.Count)
                        ws.Cells(row, 1).Value = "- " & CStr(tradeoffsObj(t))
                        row = row + 1
                    Next t
                End If
            ElseIf TypeName(tradeoffsObj) = "Dictionary" Then
                Dim tradeKey As Variant
                Dim wrote As Long
                wrote = 0
                For Each tradeKey In tradeoffsObj.Keys
                    If wrote = 0 Then
                        ws.Cells(row, 1).Value = "Tradeoffs:"
                        ws.Cells(row, 1).Font.Bold = True
                        row = row + 1
                    End If
                    ws.Cells(row, 1).Value = "- " & CStr(tradeoffsObj(tradeKey))
                    row = row + 1
                    wrote = wrote + 1
                    If wrote >= 4 Then Exit For
                Next tradeKey
            End If
        End If
    End If

    ' Band mix breakdown (per client request)
    Dim metrics As Object
    Set metrics = Nothing
    On Error Resume Next
    Set metrics = scenario("metrics")
    On Error GoTo 0

    If Not metrics Is Nothing Then
        Dim bandMix As Object
        Set bandMix = Nothing
        On Error Resume Next
        Set bandMix = metrics("band_mix")
        On Error GoTo 0

        If Not bandMix Is Nothing Then
            row = row + 1
            ws.Cells(row, 1).Value = "Band Mix (by Net SF):"
            ws.Cells(row, 1).Font.Bold = True
            row = row + 1

            ws.Cells(row, 1).Value = "Band"
            ws.Cells(row, 2).Value = "Units"
            ws.Cells(row, 3).Value = "Net SF"
            ws.Cells(row, 4).Value = "Share of SF"
            Dim hasBuildingSfCol As Boolean
            hasBuildingSfCol = (g_MihTotalBuildingSf > 0#)
            Dim headerEndCol As Long
            If hasBuildingSfCol Then
                ws.Cells(row, 5).Value = "Share of Building SF"
                headerEndCol = 5
            Else
                headerEndCol = 4
            End If
            ws.Range(ws.Cells(row, 1), ws.Cells(row, headerEndCol)).Font.Bold = True
            ws.Range(ws.Cells(row, 1), ws.Cells(row, headerEndCol)).Interior.Color = RGB(230, 230, 230)
            ws.Cells(row, 1).HorizontalAlignment = xlRight
            row = row + 1

            Dim bmIdx As Long
            For bmIdx = 1 To bandMix.Count
                Dim bm As Object
                Set bm = bandMix(bmIdx)
                If Not bm Is Nothing Then
                    If bm.Exists("band") Then
                        If IsNumeric(bm("band")) Then
                            ws.Cells(row, 1).Value = CDbl(bm("band")) / 100#
                            ws.Cells(row, 1).NumberFormat = "0%"
                        Else
                            ws.Cells(row, 1).Value = CStr(bm("band")) & "%"
                        End If
                        ws.Cells(row, 1).HorizontalAlignment = xlRight
                    End If
                    If bm.Exists("units") Then ws.Cells(row, 2).Value = bm("units")
                    If bm.Exists("net_sf") Then
                        ws.Cells(row, 3).Value = bm("net_sf")
                        ws.Cells(row, 3).NumberFormat = "0.00"
                    End If
                    If bm.Exists("share_of_sf") Then
                        ws.Cells(row, 4).Value = bm("share_of_sf")
                        ws.Cells(row, 4).NumberFormat = "0.00%"
                    End If
                    If hasBuildingSfCol And bm.Exists("net_sf") Then
                        If g_MihTotalBuildingSf > 0# Then
                            ws.Cells(row, 5).Value = CDbl(bm("net_sf")) / g_MihTotalBuildingSf
                            ws.Cells(row, 5).NumberFormat = "0.00%"
                        End If
                    End If
                    row = row + 1
                End If
            Next bmIdx
        End If
    End If

    If scenario.Exists("rent_totals") Then
        Dim rentTotals As Object
        Set rentTotals = Nothing
        On Error Resume Next
        Set rentTotals = scenario("rent_totals")
        On Error GoTo 0

        If Not rentTotals Is Nothing Then
            ' Show net totals only (gross/total allowances removed per client request)
            If rentTotals.Exists("net_monthly") Then
                ws.Cells(row, 1).Value = "Total Monthly Rent:"
                If IsNumeric(rentTotals("net_monthly")) Then
                    ws.Cells(row, 2).Value = CDbl(rentTotals("net_monthly"))
                    ws.Cells(row, 2).NumberFormat = "$#,##0"
                Else
                    ws.Cells(row, 2).Value = rentTotals("net_monthly")
                End If
                row = row + 1
            End If
            If rentTotals.Exists("net_annual") Then
                ws.Cells(row, 1).Value = "Total Annual Rent:"
                If IsNumeric(rentTotals("net_annual")) Then
                    ws.Cells(row, 2).Value = CDbl(rentTotals("net_annual"))
                    ws.Cells(row, 2).NumberFormat = "$#,##0"
                Else
                    ws.Cells(row, 2).Value = rentTotals("net_annual")
                End If
                row = row + 1
            End If
        End If
    End If

    row = row + 1

    ' Assignment table
    tableHeaderRow = row
    ws.Cells(row, 1).Value = "Unit"
    ws.Cells(row, 2).Value = "Bedrooms"
    ws.Cells(row, 3).Value = "Net SF"
    ws.Cells(row, 4).Value = "Balcony"
    ws.Cells(row, 5).Value = "AMI"
    ws.Cells(row, 6).Value = "Gross Rent"
    ws.Cells(row, 7).Value = "Net Rent"
    ws.Cells(row, 8).Value = "Annual Rent"
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 8)).Font.Bold = True
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 8)).Interior.Color = RGB(230, 230, 230)
    ws.Cells(row, 1).HorizontalAlignment = xlRight
    ws.Cells(row, 5).HorizontalAlignment = xlRight
    row = row + 1

    If scenario.Exists("assignments") Then
        Dim assignments As Object
        Set assignments = scenario("assignments")
        Dim a As Long
        For a = 1 To assignments.Count
            Dim assignment As Object
            Set assignment = assignments(a)

            ws.Cells(row, 1).Value = assignment("unit_id")
            ws.Cells(row, 1).HorizontalAlignment = xlRight
            If assignment.Exists("bedrooms") Then ws.Cells(row, 2).Value = assignment("bedrooms")
            If assignment.Exists("net_sf") Then ws.Cells(row, 3).Value = assignment("net_sf")
            If assignment.Exists("balcony") Then ws.Cells(row, 4).Value = IIf(assignment("balcony"), "Y", "")
            Dim assignedAmi As Double
            assignedAmi = CDbl(assignment("assigned_ami"))
            If assignedAmi > 2# Then assignedAmi = assignedAmi / 100#
            ws.Cells(row, 5).Value = assignedAmi
            ws.Cells(row, 5).NumberFormat = "0%"
            ws.Cells(row, 5).HorizontalAlignment = xlRight

            If assignment.Exists("gross_rent") Then
                ws.Cells(row, 6).Value = assignment("gross_rent")
                ws.Cells(row, 6).NumberFormat = "$#,##0"
            End If

            If assignment.Exists("monthly_rent") Then
                ws.Cells(row, 7).Value = assignment("monthly_rent")
                ws.Cells(row, 7).NumberFormat = "$#,##0"
            End If

            If assignment.Exists("annual_rent") Then
                ws.Cells(row, 8).Value = assignment("annual_rent")
                ws.Cells(row, 8).NumberFormat = "$#,##0"
            End If

            row = row + 1
        Next a
    End If

    ' Alignment: keep Unit and AMI columns consistent regardless of text/numeric formatting.
    Dim firstDataRow As Long
    Dim lastDataRow As Long
    firstDataRow = tableHeaderRow + 1
    lastDataRow = row - 1
    If lastDataRow >= firstDataRow Then
        ws.Range(ws.Cells(firstDataRow, 1), ws.Cells(lastDataRow, 1)).HorizontalAlignment = xlRight
        ws.Range(ws.Cells(firstDataRow, 5), ws.Cells(lastDataRow, 5)).HorizontalAlignment = xlRight
    End If

    WriteScenarioSummaryAndTable = row
End Function

Private Function BuildAllowanceBreakdown(allowancesValue As Variant) As String
    ' API returns allowances as an ARRAY (Collection) of objects:
    '   [{amount, category, label}, ...]
    ' Older versions may return a Dictionary keyed by category.
    Dim allowanceStr As String
    allowanceStr = ""
    
    On Error GoTo Cleanup
    
    If IsObject(allowancesValue) Then
        Select Case TypeName(allowancesValue)
            Case "Collection"
                Dim allowancesArr As Collection
                Set allowancesArr = allowancesValue
                
                Dim i As Long
                For i = 1 To allowancesArr.Count
                    Dim item As Object
                    Set item = allowancesArr(i)
                    
                    If Not item Is Nothing Then
                        Dim category As String
                        Dim amount As Double
                        category = ""
                        amount = 0
                        
                        If item.Exists("category") Then category = CStr(item("category"))
                        If item.Exists("amount") And IsNumeric(item("amount")) Then amount = CDbl(item("amount"))
                        
                        If amount > 0 And category <> "" Then
                            Dim shortName As String
                            Select Case LCase(category)
                                Case "electricity": shortName = "Elec"
                                Case "cooking": shortName = "Cook"
                                Case "heat": shortName = "Heat"
                                Case "hot_water": shortName = "HW"
                                Case Else: shortName = category
                            End Select
                            
                            If allowanceStr <> "" Then allowanceStr = allowanceStr & " + "
                            allowanceStr = allowanceStr & shortName & "($" & Format(amount, "0") & ")"
                        End If
                    End If
                Next i
                
            Case "Dictionary"
                Dim allowancesDict As Object
                Set allowancesDict = allowancesValue
                
                ' Build breakdown string: Elec($X) + Cook($Y) + Heat($Z) + HW($W)
                If allowancesDict.Exists("electricity") Then
                    If IsNumeric(allowancesDict("electricity")) Then
                        If CDbl(allowancesDict("electricity")) > 0 Then
                            allowanceStr = "Elec($" & Format(allowancesDict("electricity"), "0") & ")"
                        End If
                    End If
                End If
                If allowancesDict.Exists("cooking") Then
                    If IsNumeric(allowancesDict("cooking")) Then
                        If CDbl(allowancesDict("cooking")) > 0 Then
                            If allowanceStr <> "" Then allowanceStr = allowanceStr & " + "
                            allowanceStr = allowanceStr & "Cook($" & Format(allowancesDict("cooking"), "0") & ")"
                        End If
                    End If
                End If
                If allowancesDict.Exists("heat") Then
                    If IsNumeric(allowancesDict("heat")) Then
                        If CDbl(allowancesDict("heat")) > 0 Then
                            If allowanceStr <> "" Then allowanceStr = allowanceStr & " + "
                            allowanceStr = allowanceStr & "Heat($" & Format(allowancesDict("heat"), "0") & ")"
                        End If
                    End If
                End If
                If allowancesDict.Exists("hot_water") Then
                    If IsNumeric(allowancesDict("hot_water")) Then
                        If CDbl(allowancesDict("hot_water")) > 0 Then
                            If allowanceStr <> "" Then allowanceStr = allowanceStr & " + "
                            allowanceStr = allowanceStr & "HW($" & Format(allowancesDict("hot_water"), "0") & ")"
                        End If
                    End If
                End If
        End Select
    End If
    
Cleanup:
    BuildAllowanceBreakdown = allowanceStr
End Function

Private Function FormatScenarioName(key As String) As String
    ' Formats scenario key into readable name
    Select Case key
        Case "absolute_best"
            FormatScenarioName = "ABSOLUTE BEST"
        Case "best_3_band"
            FormatScenarioName = "BEST 3-BAND"
        Case "best_2_band"
            FormatScenarioName = "BEST 2-BAND"
        Case "alternative"
            FormatScenarioName = "ALTERNATIVE"
        Case "client_oriented"
            FormatScenarioName = "CLIENT ORIENTED (MAX REVENUE)"
        Case Else
            FormatScenarioName = UCase(Replace(key, "_", " "))
    End Select
End Function

'-------------------------------------------------------------------------------
' UTILITY SETTINGS DISPLAY
'-------------------------------------------------------------------------------

Private Function WriteUtilitySettings(ws As Worksheet, startRow As Long) As Long
    ' Writes the current utility settings to the scenarios sheet
    ' Shows which utilities the TENANT pays for (affects rent calculations)

    Dim row As Long
    row = startRow

    ' Header
    ws.Cells(row, 1).Value = "UTILITIES - Selected Variants (Affects Rent Allowances)"
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 1).Font.Size = 12
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 4)).Interior.Color = RGB(255, 230, 200)
    row = row + 1

    ' Get current utility settings from registry
    Dim elec As String, cook As String, heat As String, hw As String
    elec = GetSetting("AMI_Optix", "Utilities", "electricity", "na")
    cook = GetSetting("AMI_Optix", "Utilities", "cooking", "na")
    heat = GetSetting("AMI_Optix", "Utilities", "heat", "na")
    hw = GetSetting("AMI_Optix", "Utilities", "hot_water", "na")

    ' Column headers
    ws.Cells(row, 1).Value = "Utility"
    ws.Cells(row, 2).Value = "Tenant Pays?"
    ws.Cells(row, 3).Value = "Type"
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 3)).Font.Bold = True
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 3)).Interior.Color = RGB(230, 230, 230)
    row = row + 1

    ' Electricity
    ws.Cells(row, 1).Value = "Electricity"
    ws.Cells(row, 3).Value = FormatUtilityType(elec, "electricity")
    If elec = "tenant_pays" Then
        ws.Cells(row, 2).Value = "YES"
        ws.Cells(row, 2).Font.Color = RGB(0, 128, 0)
    Else
        ws.Cells(row, 2).Value = "NO"
        ws.Cells(row, 2).Font.Color = RGB(128, 128, 128)
    End If
    row = row + 1

    ' Cooking
    ws.Cells(row, 1).Value = "Cooking"
    ws.Cells(row, 3).Value = FormatUtilityType(cook, "cooking")
    If cook <> "na" Then
        ws.Cells(row, 2).Value = "YES"
        ws.Cells(row, 2).Font.Color = RGB(0, 128, 0)
    Else
        ws.Cells(row, 2).Value = "NO"
        ws.Cells(row, 2).Font.Color = RGB(128, 128, 128)
    End If
    row = row + 1

    ' Heat
    ws.Cells(row, 1).Value = "Heat"
    ws.Cells(row, 3).Value = FormatUtilityType(heat, "heat")
    If heat <> "na" Then
        ws.Cells(row, 2).Value = "YES"
        ws.Cells(row, 2).Font.Color = RGB(0, 128, 0)
    Else
        ws.Cells(row, 2).Value = "NO"
        ws.Cells(row, 2).Font.Color = RGB(128, 128, 128)
    End If
    row = row + 1

    ' Hot Water
    ws.Cells(row, 1).Value = "Hot Water"
    ws.Cells(row, 3).Value = FormatUtilityType(hw, "hot_water")
    If hw <> "na" Then
        ws.Cells(row, 2).Value = "YES"
        ws.Cells(row, 2).Font.Color = RGB(0, 128, 0)
    Else
        ws.Cells(row, 2).Value = "NO"
        ws.Cells(row, 2).Font.Color = RGB(128, 128, 128)
    End If
    row = row + 1

    WriteUtilitySettings = row
End Function

Private Function SumAllowanceAmounts(allowancesValue As Variant) As Double
    ' Best-effort sum of allowance amounts from the API payload.
    ' Expected shapes:
    ' - Collection of objects: [{amount, category, label}, ...]
    ' - Dictionary keyed by category with numeric values
    On Error GoTo Fail

    SumAllowanceAmounts = 0#
    If Not IsObject(allowancesValue) Then Exit Function

    Select Case TypeName(allowancesValue)
        Case "Collection"
            Dim allowancesArr As Collection
            Set allowancesArr = allowancesValue

            Dim i As Long
            For i = 1 To allowancesArr.Count
                Dim item As Object
                Set item = allowancesArr(i)
                If Not item Is Nothing Then
                    If item.Exists("amount") Then
                        If IsNumeric(item("amount")) Then
                            SumAllowanceAmounts = SumAllowanceAmounts + CDbl(item("amount"))
                        End If
                    End If
                End If
            Next i

        Case "Dictionary", "Scripting.Dictionary"
            Dim d As Object
            Set d = allowancesValue

            Dim key As Variant
            For Each key In d.Keys
                Dim v As Variant
                v = d(key)
                If IsNumeric(v) Then
                    SumAllowanceAmounts = SumAllowanceAmounts + CDbl(v)
                ElseIf IsObject(v) Then
                    On Error Resume Next
                    If v.Exists("amount") Then
                        If IsNumeric(v("amount")) Then
                            SumAllowanceAmounts = SumAllowanceAmounts + CDbl(v("amount"))
                        End If
                    End If
                    On Error GoTo Fail
                End If
            Next key
    End Select
    Exit Function

Fail:
    SumAllowanceAmounts = 0#
End Function

Private Function WriteUtilityDeductionTotalsByBedroom(ws As Worksheet, startRow As Long, scenario As Object) As Long
    ' Writes per-bedroom utility deduction totals (monthly) under the utilities block.
    ' Client request: show the per-utility breakdown (Electricity/Cooking/Heat/Hot Water) once at the top.
    On Error GoTo Fail

    Dim row As Long
    row = startRow

    If scenario Is Nothing Then GoTo SafeExit
    If Not scenario.Exists("assignments") Then GoTo SafeExit

    Dim assignments As Object
    Set assignments = Nothing
    On Error Resume Next
    Set assignments = scenario("assignments")
    On Error GoTo Fail
    If assignments Is Nothing Then GoTo SafeExit

    Dim byBr As Object
    Set byBr = CreateObject("Scripting.Dictionary") ' bedrooms -> dict(electricity,cooking,heat,hot_water,total)

    Dim i As Long
    For i = 1 To assignments.Count
        Dim a As Object
        Set a = assignments(i)
        If a Is Nothing Then GoTo NextAssignment

        Dim br As Long
        br = 0
        On Error Resume Next
        If a.Exists("bedrooms") Then br = CLng(a("bedrooms"))
        On Error GoTo Fail

        Dim k As String
        k = CStr(br)
        If byBr.Exists(k) Then GoTo NextAssignment ' already captured a representative row for this bedroom type

        Dim elecAmt As Double, cookAmt As Double, heatAmt As Double, hwAmt As Double
        elecAmt = 0#: cookAmt = 0#: heatAmt = 0#: hwAmt = 0#

        If a.Exists("allowances") Then
            Dim allowancesValue As Variant
            allowancesValue = a("allowances")

            If IsObject(allowancesValue) Then
                Select Case TypeName(allowancesValue)
                    Case "Collection"
                        Dim allowancesArr As Collection
                        Set allowancesArr = allowancesValue

                        Dim j As Long
                        For j = 1 To allowancesArr.Count
                            Dim item As Object
                            Set item = allowancesArr(j)
                            If item Is Nothing Then GoTo NextItem

                            Dim cat As String
                            cat = ""
                            On Error Resume Next
                            If item.Exists("category") Then cat = LCase$(Trim$(CStr(item("category"))))
                            On Error GoTo Fail

                            Dim amt As Double
                            amt = 0#
                            On Error Resume Next
                            If item.Exists("amount") Then
                                If IsNumeric(item("amount")) Then amt = CDbl(item("amount"))
                            End If
                            On Error GoTo Fail

                            Select Case cat
                                Case "electricity": elecAmt = amt
                                Case "cooking": cookAmt = amt
                                Case "heat": heatAmt = amt
                                Case "hot_water": hwAmt = amt
                            End Select

NextItem:
                        Next j

                    Case "Dictionary", "Scripting.Dictionary"
                        ' Legacy shape: {"electricity": 123, ...} or {"electricity": {amount:123}, ...}
                        Dim d As Object
                        Set d = allowancesValue

                        elecAmt = AllowanceAmountFromDict(d, "electricity")
                        cookAmt = AllowanceAmountFromDict(d, "cooking")
                        heatAmt = AllowanceAmountFromDict(d, "heat")
                        hwAmt = AllowanceAmountFromDict(d, "hot_water")
                End Select
            End If
        End If

        Dim totalAmt As Double
        totalAmt = elecAmt + cookAmt + heatAmt + hwAmt
        If a.Exists("allowance_total") Then
            If IsNumeric(a("allowance_total")) Then totalAmt = CDbl(a("allowance_total"))
        End If

        Dim detail As Object
        Set detail = CreateObject("Scripting.Dictionary")
        detail("electricity") = elecAmt
        detail("cooking") = cookAmt
        detail("heat") = heatAmt
        detail("hot_water") = hwAmt
        detail("total") = totalAmt
        byBr(k) = detail

NextAssignment:
    Next i

    If byBr.Count = 0 Then GoTo SafeExit

    ws.Cells(row, 1).Value = "TOTAL UTILITY DEDUCTION (Monthly)"
    ws.Cells(row, 1).Font.Bold = True
    row = row + 1

    ws.Cells(row, 1).Value = "Bedroom Type"
    ws.Cells(row, 2).Value = "Electricity"
    ws.Cells(row, 3).Value = "Cooking"
    ws.Cells(row, 4).Value = "Heat"
    ws.Cells(row, 5).Value = "Hot Water"
    ws.Cells(row, 6).Value = "Total"
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 6)).Font.Bold = True
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 6)).Interior.Color = RGB(230, 230, 230)
    row = row + 1

    Dim brIdx As Long
    For brIdx = 0 To 10
        Dim brKey As String
        brKey = CStr(brIdx)
        If byBr.Exists(brKey) Then
            ws.Cells(row, 1).Value = IIf(brIdx <= 0, "Studio", CStr(brIdx) & " BR")
            Dim d2 As Object
            Set d2 = byBr(brKey)

            ws.Cells(row, 2).Value = CDbl(d2("electricity"))
            ws.Cells(row, 3).Value = CDbl(d2("cooking"))
            ws.Cells(row, 4).Value = CDbl(d2("heat"))
            ws.Cells(row, 5).Value = CDbl(d2("hot_water"))
            ws.Cells(row, 6).Value = CDbl(d2("total"))

            ws.Range(ws.Cells(row, 2), ws.Cells(row, 6)).NumberFormat = "$#,##0"
            ws.Range(ws.Cells(row, 2), ws.Cells(row, 6)).HorizontalAlignment = xlRight
            row = row + 1
        End If
    Next brIdx

SafeExit:
    WriteUtilityDeductionTotalsByBedroom = row
    Exit Function

Fail:
    WriteUtilityDeductionTotalsByBedroom = startRow
End Function

Private Function AllowanceAmountFromDict(d As Object, key As String) As Double
    On Error GoTo Fail
    AllowanceAmountFromDict = 0#
    If d Is Nothing Then Exit Function
    If Not d.Exists(key) Then Exit Function

    Dim v As Variant
    v = d(key)
    If IsNumeric(v) Then
        AllowanceAmountFromDict = CDbl(v)
        Exit Function
    End If

    If IsObject(v) Then
        On Error Resume Next
        If v.Exists("amount") Then
            If IsNumeric(v("amount")) Then
                AllowanceAmountFromDict = CDbl(v("amount"))
            End If
        End If
        On Error GoTo Fail
    End If
    Exit Function

Fail:
    AllowanceAmountFromDict = 0#
End Function

Private Function FormatUtilityType(value As String, Optional category As String = "") As String
    ' Formats utility selection codes to the exact rent roll guideline labels (do not collapse variants).
    Dim v As String
    Dim c As String
    v = LCase$(Trim$(CStr(value)))
    c = LCase$(Trim$(CStr(category)))

    Select Case c
        Case "electricity"
            If v = "tenant_pays" Then
                FormatUtilityType = "Tenant Pays"
            Else
                FormatUtilityType = "N/A or owner pays"
            End If
            Exit Function

        Case "cooking"
            Select Case v
                Case "electric", "electric_stove": FormatUtilityType = "Electric Stove"
                Case "gas": FormatUtilityType = "Gas Stove"
                Case Else: FormatUtilityType = "N/A or owner pays"
            End Select
            Exit Function

        Case "heat"
            Select Case v
                Case "electric_ccashp": FormatUtilityType = "Electric Heat - Cold Climate Air Source Heat Pump (ccASHP)1"
                Case "electric_other": FormatUtilityType = "Electric Heat - Other2"
                Case "gas": FormatUtilityType = "Gas Heat"
                Case "oil": FormatUtilityType = "Oil Heat"
                Case Else: FormatUtilityType = "N/A or owner pays"
            End Select
            Exit Function

        Case "hot_water"
            Select Case v
                Case "electric_heat_pump": FormatUtilityType = "Electric Hot Water - Heat Pump"
                Case "electric_other": FormatUtilityType = "Electric Hot Water - Other"
                Case "gas": FormatUtilityType = "Gas Hot Water"
                Case "oil": FormatUtilityType = "Oil Hot Water"
                Case Else: FormatUtilityType = "N/A or owner pays"
            End Select
            Exit Function
    End Select

    ' Backward-compatible fallback (should not be used when category is provided).
    Select Case v
        Case "electric", "electric_stove": FormatUtilityType = "Electric Stove"
        Case "gas": FormatUtilityType = "Gas"
        Case "oil": FormatUtilityType = "Oil"
        Case "electric_ccashp": FormatUtilityType = "Electric (ccASHP)"
        Case "electric_other": FormatUtilityType = "Electric (Other)"
        Case "electric_heat_pump": FormatUtilityType = "Electric (Heat Pump)"
        Case "tenant_pays": FormatUtilityType = "Tenant Pays"
        Case "na", "": FormatUtilityType = "N/A or owner pays"
        Case Else: FormatUtilityType = CStr(value)
    End Select
End Function

'-------------------------------------------------------------------------------
' APPLY SPECIFIC SCENARIO (for manual selection)
'-------------------------------------------------------------------------------

Public Sub ApplyScenarioByKey(scenarioKey As String)
    ' Applies a specific scenario by its key
    ' Called from scenario sheet buttons

    Dim prevSheet As Worksheet
    Set prevSheet = ActiveSheet

    If g_LastScenarios Is Nothing Then
        MsgBox "No scenarios available. Run Optimize first.", vbExclamation, "AMI Optix"
        Exit Sub
    End If

    Dim scenarios As Object
    Set scenarios = g_LastScenarios("scenarios")

    If Not scenarios.Exists(scenarioKey) Then
        MsgBox "Scenario '" & scenarioKey & "' not found.", vbExclamation, "AMI Optix"
        Exit Sub
    End If

    Dim scenario As Object
    Set scenario = scenarios(scenarioKey)

    Dim assignments As Object
    Set assignments = scenario("assignments")

    Dim ws As Worksheet
    Set ws = GetDataSheet()

    Dim amiCol As Long
    amiCol = GetAMIColumn()

    If ws Is Nothing Or amiCol = 0 Then
        MsgBox "Cannot write results: data sheet or AMI column not found.", vbExclamation, "AMI Optix"
        Exit Sub
    End If

    ' Build lookup
    Dim unitRows As Object
    Set unitRows = BuildUnitRowLookup(ws)

    Dim programNorm As String
    Dim mihOption As String
    programNorm = "UAP"
    mihOption = ""

    On Error Resume Next
    If Not g_LastScenarios Is Nothing Then
        If g_LastScenarios.Exists("project_summary") Then
            Dim ps As Object
            Set ps = g_LastScenarios("project_summary")
            If Not ps Is Nothing Then
                If ps.Exists("program") Then programNorm = UCase$(CStr(ps("program")))
                If ps.Exists("mih_option") Then mihOption = CStr(ps("mih_option"))
            End If
        End If
    End If
    On Error GoTo 0

    Dim prevEnableEvents As Boolean
    Dim prevSuppress As Boolean
    prevEnableEvents = Application.EnableEvents
    prevSuppress = g_AMIOptixSuppressEvents
    Application.EnableEvents = False
    g_AMIOptixSuppressEvents = True

    On Error GoTo ApplyFail

    ' Apply
    Dim i As Long
    Dim assignment As Object
    Dim unitId As String
    Dim ami As Double
    Dim amiValue As Double
    Dim row As Long
    Dim updatedCount As Long

    updatedCount = 0

    For i = 1 To assignments.Count
        Set assignment = assignments(i)

        unitId = CStr(assignment("unit_id"))
        ami = CDbl(assignment("assigned_ami"))

        If unitRows.Exists(unitId) Then
            row = unitRows(unitId)

            ' Support both whole-percent (60/120) and decimal (0.6/1.2) values.
            If ami > 2# Then
                amiValue = ami / 100#  ' Convert 60 to 0.60; 120 to 1.20
            Else
                amiValue = ami
            End If

            ws.Cells(row, amiCol).Value = amiValue
            ws.Cells(row, amiCol).NumberFormat = "0%"  ' Ensure percentage format
            ws.Cells(row, amiCol).Interior.Color = RGB(200, 255, 200)  ' Light green
            updatedCount = updatedCount + 1
        End If
    Next i

    ' Refresh Scenario Manual (live sync) to match what is now applied.
    Dim manualOk As Boolean
    manualOk = False
    On Error Resume Next
    Call RefreshManualScenarioFromScenario(scenarioKey, scenario)
    manualOk = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    ' Best-effort learning audit: record the user's chosen scenario.
    On Error Resume Next
    Dim profileKey As String
    profileKey = GetLearningProfileKey(programNorm, mihOption)
    Call LogScenarioApplied(profileKey, programNorm, mihOption, scenarioKey, "USER", scenario)
    On Error GoTo 0

    ' Ensure WAAMI/Avg AMI display shows sufficient precision (e.g., 59.96% vs 60.0%).
    On Error Resume Next
    EnsureProvidedAvgAmiPrecision
    On Error GoTo 0

    GoTo ApplyCleanup

ApplyFail:
    Debug.Print "ApplyScenarioByKey Error: " & Err.Description
ApplyCleanup:
    Application.EnableEvents = prevEnableEvents
    g_AMIOptixSuppressEvents = prevSuppress

    Dim msg As String
    msg = "Applied scenario '" & FormatScenarioName(scenarioKey) & "'" & vbCrLf & _
          "Updated " & updatedCount & " units."
    If Not manualOk Then
        msg = msg & vbCrLf & vbCrLf & _
              "Note: Scenario Manual did not refresh." & vbCrLf & _
              "Click AMI Optix → Diagnostics to see the API/error details."
    End If

    MsgBox msg, vbInformation, "AMI Optix"

    ' Preserve where the user was working (avoid the "AMI Scenarios tab disappeared" confusion).
    On Error Resume Next
    If Not prevSheet Is Nothing Then prevSheet.Activate
    On Error GoTo 0
End Sub

Private Sub EnsureProvidedAvgAmiPrecision()
    ' Fixes the common workbook display issue where the provided Avg AMI shows as 60.0%
    ' even when the computed value is e.g. 59.96% (formatting/rounding only).
    On Error GoTo Fail

    If ActiveWorkbook Is Nothing Then Exit Sub

    Dim wsCalc As Worksheet
    On Error Resume Next
    Set wsCalc = ActiveWorkbook.Worksheets("Calculations")
    On Error GoTo Fail
    If wsCalc Is Nothing Then Exit Sub

    Dim c As Range
    Set c = wsCalc.Range("C20")

    Dim fmt As String
    fmt = Replace(CStr(c.NumberFormat), " ", "")

    ' Only adjust if the workbook is using coarse percent rounding.
    If fmt = "0%" Or fmt = "0.0%" Then
        c.NumberFormat = "0.00%"
    End If

    wsCalc.Calculate

    ' Also adjust the displayed "Avg AMI" cells on the main sheet (e.g., AHFA chart).
    Dim wsData As Worksheet
    Set wsData = GetDataSheet()
    If Not wsData Is Nothing Then
        EnsureAvgAmiRowPrecision wsData
        wsData.Calculate
    End If
    Exit Sub

Fail:
End Sub

Private Sub EnsureAvgAmiRowPrecision(ws As Worksheet)
    On Error GoTo Fail

    If ws Is Nothing Then Exit Sub

    Dim found As Range
    Set found = ws.Cells.Find(What:="Avg AMI", After:=ws.Cells(1, 1), LookIn:=xlValues, LookAt:=xlPart, _
                              SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    If found Is Nothing Then Exit Sub

    Dim col As Long
    For col = found.Column + 1 To Application.Min(found.Column + 10, ws.Columns.Count)
        Dim c As Range
        Set c = ws.Cells(found.Row, col)

        If IsNumeric(c.Value) Then
            Dim v As Double
            v = CDbl(c.Value)
            If v > 0# And v < 1# Then
                Dim fmt As String
                fmt = Replace(CStr(c.NumberFormat), " ", "")
                If InStr(1, fmt, "%", vbTextCompare) > 0 Then
                    c.NumberFormat = "0.00%"
                End If
            End If
        End If
    Next col

Fail:
End Sub
