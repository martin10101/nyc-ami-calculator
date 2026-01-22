Attribute VB_Name = "AMI_Optix_ResultsWriter"
'===============================================================================
' AMI OPTIX - Results Writer Module
' Writes optimization results back to Excel
'===============================================================================
Option Explicit

Private Const MANUAL_BLOCK_HEIGHT As Long = 120
Private Const MANUAL_BLOCK_START_ROW As Long = 1
Private Const SCENARIOS_START_ROW As Long = 125

' True only when the most recent /api/evaluate call returned success=false.
Public g_AMIOptixLastManualScenarioInvalid As Boolean

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
            ' API returns integer (60, 80, 130) - convert to decimal for percentage format
            If ami > 1 Then
                amiValue = ami / 100  ' Convert 60 to 0.60
            Else
                amiValue = ami  ' Already decimal
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

    If result Is Nothing Then GoTo ErrorHandler
    If Not result.Exists("scenarios") Then GoTo ErrorHandler
    Set scenarios = result("scenarios")

    ' Start solver output below the manual block
    row = 1

    ' Manual block (reserved region at top)
    WriteManualScenarioBlockFromResult ws, result

    row = SCENARIOS_START_ROW

    ' Notes from solver
    If result.Exists("notes") Then
        ws.Cells(row, 1).Value = "Solver Notes:"
        ws.Cells(row, 1).Font.Bold = True
        row = row + 1

        Dim notes As Object
        Set notes = result("notes")
        Dim n As Long
        For n = 1 To notes.Count
            ws.Cells(row, 1).Value = "- " & notes(n)
            row = row + 1
        Next n
        row = row + 1
    End If

    ' Process each scenario
    Dim scenarioNum As Long
    scenarioNum = 1

    For Each scenarioKey In scenarios.Keys
        Set scenario = scenarios(scenarioKey)

        ' Scenario header
        ws.Cells(row, 1).Value = "SCENARIO " & scenarioNum & ": " & FormatScenarioName(CStr(scenarioKey))
        ws.Cells(row, 1).Font.Bold = True
        ws.Cells(row, 1).Font.Size = 14
        ws.Range(ws.Cells(row, 1), ws.Cells(row, 10)).Interior.Color = RGB(200, 220, 255)
        row = row + 1

        row = WriteScenarioSummaryAndTable(ws, row, scenario)
        row = row + 1

        ws.Range(ws.Cells(row, 1), ws.Cells(row, 10)).Interior.Color = RGB(240, 240, 240)
        row = row + 1
        scenarioNum = scenarioNum + 1
    Next scenarioKey

    ' Auto-fit columns
    ws.Columns("A:K").AutoFit

    ' Freeze the top row and jump to the solver output so users immediately see scenarios.
    On Error Resume Next
    ws.Activate
    ws.Rows(2).Select
    ActiveWindow.FreezePanes = True
    ws.Cells(SCENARIOS_START_ROW, 1).Select
    ActiveWindow.ScrollRow = SCENARIOS_START_ROW
    On Error GoTo ErrorHandler

    Debug.Print "Created scenarios sheet with " & (scenarioNum - 1) & " scenarios"
    GoTo Cleanup

ErrorHandler:
    hadError = True
    errMsg = Err.Description
    Debug.Print "CreateScenariosSheet Error: " & errMsg
Cleanup:
    Application.EnableEvents = prevEnableEvents
    g_AMIOptixSuppressEvents = prevSuppress
    If hadError Then
        MsgBox "Failed to build 'AMI Scenarios' sheet: " & errMsg, vbExclamation, "AMI Optix"
    End If
End Sub

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
        If band > 1 Then
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
        Set dataWs = ActiveWorkbook.Worksheets("RentRoll")
        If dataWs Is Nothing Then Set dataWs = ActiveWorkbook.Worksheets("MIH")
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
    Set GetOrCreateScenariosSheet = ws
End Function

Private Sub ClearManualBlock(ws As Worksheet)
    ws.Range("A1:M" & MANUAL_BLOCK_HEIGHT).Clear
End Sub

Private Sub WriteManualScenarioBlockFromResult(ws As Worksheet, result As Object)
    ClearManualBlock ws

    Dim row As Long
    row = MANUAL_BLOCK_START_ROW

    ws.Cells(row, 1).Value = "AMI OPTIMIZATION RESULTS"
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 1).Font.Size = 16
    row = row + 2

    row = WriteUtilitySettings(ws, row)
    row = row + 1

    ws.Cells(row, 1).Value = "SCENARIO MANUAL (LIVE SYNC)"
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 1).Font.Size = 14
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 13)).Interior.Color = RGB(220, 240, 220)
    row = row + 1

    Dim scenarioKey As String
    scenarioKey = GetBestScenarioKey(result)
    If scenarioKey = "" Then Exit Sub

    Dim scenarios As Object
    Set scenarios = result("scenarios")
    If Not scenarios.Exists(scenarioKey) Then Exit Sub

    Dim scenario As Object
    Set scenario = scenarios(scenarioKey)

    row = WriteScenarioSummaryAndTable(ws, row, scenario)
End Sub

Private Sub WriteManualScenarioBlockFromEvaluate(ws As Worksheet, evalResult As Object)
    ClearManualBlock ws

    Dim row As Long
    row = MANUAL_BLOCK_START_ROW

    ws.Cells(row, 1).Value = "AMI OPTIMIZATION RESULTS"
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 1).Font.Size = 16
    row = row + 2

    row = WriteUtilitySettings(ws, row)
    row = row + 1

    ws.Cells(row, 1).Value = "SCENARIO MANUAL (LIVE SYNC)"
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 1).Font.Size = 14
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 13)).Interior.Color = RGB(220, 240, 220)
    row = row + 1

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

    row = WriteScenarioSummaryAndTable(ws, row, scenario)
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
        Dim bands As Variant
        bands = scenario("bands")

        On Error Resume Next
        Dim bCount As Long
        bCount = bands.Count
        On Error GoTo 0

        If bCount > 0 Then
            Dim b As Long
            For b = 1 To bCount
                If bandStr <> "" Then bandStr = bandStr & ", "
                bandStr = bandStr & Format(bands(b), "0") & "%"
            Next b
        Else
            ' bands might already be a VBA array
        End If

        ws.Cells(row, 2).Value = bandStr
        row = row + 1
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
            ws.Range(ws.Cells(row, 1), ws.Cells(row, 4)).Font.Bold = True
            ws.Range(ws.Cells(row, 1), ws.Cells(row, 4)).Interior.Color = RGB(230, 230, 230)
            row = row + 1

            Dim bmIdx As Long
            For bmIdx = 1 To bandMix.Count
                Dim bm As Object
                Set bm = bandMix(bmIdx)
                If Not bm Is Nothing Then
                    If bm.Exists("band") Then ws.Cells(row, 1).Value = CStr(bm("band")) & "%"
                    If bm.Exists("units") Then ws.Cells(row, 2).Value = bm("units")
                    If bm.Exists("net_sf") Then
                        ws.Cells(row, 3).Value = bm("net_sf")
                        ws.Cells(row, 3).NumberFormat = "0.00"
                    End If
                    If bm.Exists("share_of_sf") Then
                        ws.Cells(row, 4).Value = bm("share_of_sf")
                        ws.Cells(row, 4).NumberFormat = "0.00%"
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
                ws.Cells(row, 2).Value = Format(rentTotals("net_monthly"), "$#,##0.00")
                row = row + 1
            End If
            If rentTotals.Exists("net_annual") Then
                ws.Cells(row, 1).Value = "Total Annual Rent:"
                ws.Cells(row, 2).Value = Format(rentTotals("net_annual"), "$#,##0.00")
                row = row + 1
            End If

            ' Allowance breakdown (per utility)
            Dim ab As Object
            Set ab = Nothing
            On Error Resume Next
            Set ab = rentTotals("allowances_breakdown")
            On Error GoTo 0

            If Not ab Is Nothing Then
                row = row + 1
                ws.Cells(row, 1).Value = "Utility Allowances (Monthly):"
                ws.Cells(row, 1).Font.Bold = True
                row = row + 1

                Dim key As Variant
                For Each key In ab.Keys
                    Dim entry As Object
                    Set entry = ab(key)
                    ws.Cells(row, 1).Value = UCase(CStr(key))
                    If Not entry Is Nothing Then
                        If entry.Exists("monthly") Then
                            ws.Cells(row, 2).Value = Format(entry("monthly"), "$#,##0.00")
                        End If
                    End If
                    row = row + 1
                Next key
            End If
        End If
    End If

    row = row + 1

    ' Assignment table
    ws.Cells(row, 1).Value = "Unit"
    ws.Cells(row, 2).Value = "Bedrooms"
    ws.Cells(row, 3).Value = "Net SF"
    ws.Cells(row, 4).Value = "Floor"
    ws.Cells(row, 5).Value = "Balcony"
    ws.Cells(row, 6).Value = "AMI"
    ws.Cells(row, 7).Value = "Gross Rent"
    ws.Cells(row, 8).Value = "Allowances"
    ws.Cells(row, 9).Value = "Net Rent"
    ws.Cells(row, 10).Value = "Annual Rent"
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 10)).Font.Bold = True
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 10)).Interior.Color = RGB(230, 230, 230)
    row = row + 1

    If scenario.Exists("assignments") Then
        Dim assignments As Object
        Set assignments = scenario("assignments")
        Dim a As Long
        For a = 1 To assignments.Count
            Dim assignment As Object
            Set assignment = assignments(a)

            ws.Cells(row, 1).Value = assignment("unit_id")
            If assignment.Exists("bedrooms") Then ws.Cells(row, 2).Value = assignment("bedrooms")
            If assignment.Exists("net_sf") Then ws.Cells(row, 3).Value = assignment("net_sf")
            If assignment.Exists("floor") Then ws.Cells(row, 4).Value = assignment("floor")
            If assignment.Exists("balcony") Then ws.Cells(row, 5).Value = IIf(assignment("balcony"), "Y", "")
            ws.Cells(row, 6).Value = Format(assignment("assigned_ami"), "0%")

            If assignment.Exists("gross_rent") Then
                ws.Cells(row, 7).Value = assignment("gross_rent")
                ws.Cells(row, 7).NumberFormat = "$#,##0.00"
            End If

            If assignment.Exists("allowances") Then
                ws.Cells(row, 8).Value = BuildAllowanceBreakdown(assignment("allowances"))
            ElseIf assignment.Exists("allowance_total") Then
                ws.Cells(row, 8).Value = assignment("allowance_total")
                ws.Cells(row, 8).NumberFormat = "$#,##0.00"
            End If

            If assignment.Exists("monthly_rent") Then
                ws.Cells(row, 9).Value = assignment("monthly_rent")
                ws.Cells(row, 9).NumberFormat = "$#,##0.00"
            End If

            If assignment.Exists("annual_rent") Then
                ws.Cells(row, 10).Value = assignment("annual_rent")
                ws.Cells(row, 10).NumberFormat = "$#,##0.00"
            End If

            row = row + 1
        Next a
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
    ws.Cells(row, 1).Value = "TENANT-PAID UTILITIES (Affects Rent Allowances)"
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
    If elec = "tenant_pays" Then
        ws.Cells(row, 2).Value = "YES"
        ws.Cells(row, 2).Font.Color = RGB(0, 128, 0)
        ws.Cells(row, 3).Value = "Standard"
    Else
        ws.Cells(row, 2).Value = "NO"
        ws.Cells(row, 2).Font.Color = RGB(128, 128, 128)
        ws.Cells(row, 3).Value = "Owner Pays"
    End If
    row = row + 1

    ' Cooking
    ws.Cells(row, 1).Value = "Cooking"
    If cook <> "na" Then
        ws.Cells(row, 2).Value = "YES"
        ws.Cells(row, 2).Font.Color = RGB(0, 128, 0)
        ws.Cells(row, 3).Value = FormatUtilityType(cook)
    Else
        ws.Cells(row, 2).Value = "NO"
        ws.Cells(row, 2).Font.Color = RGB(128, 128, 128)
        ws.Cells(row, 3).Value = "Owner Pays"
    End If
    row = row + 1

    ' Heat
    ws.Cells(row, 1).Value = "Heat"
    If heat <> "na" Then
        ws.Cells(row, 2).Value = "YES"
        ws.Cells(row, 2).Font.Color = RGB(0, 128, 0)
        ws.Cells(row, 3).Value = FormatUtilityType(heat)
    Else
        ws.Cells(row, 2).Value = "NO"
        ws.Cells(row, 2).Font.Color = RGB(128, 128, 128)
        ws.Cells(row, 3).Value = "Owner Pays"
    End If
    row = row + 1

    ' Hot Water
    ws.Cells(row, 1).Value = "Hot Water"
    If hw <> "na" Then
        ws.Cells(row, 2).Value = "YES"
        ws.Cells(row, 2).Font.Color = RGB(0, 128, 0)
        ws.Cells(row, 3).Value = FormatUtilityType(hw)
    Else
        ws.Cells(row, 2).Value = "NO"
        ws.Cells(row, 2).Font.Color = RGB(128, 128, 128)
        ws.Cells(row, 3).Value = "Owner Pays"
    End If
    row = row + 1

    WriteUtilitySettings = row
End Function

Private Function FormatUtilityType(value As String) As String
    ' Formats utility type code to display name
    Select Case value
        Case "electric", "electric_stove": FormatUtilityType = "Electric"
        Case "gas": FormatUtilityType = "Gas"
        Case "oil": FormatUtilityType = "Oil"
        Case "electric_ccashp": FormatUtilityType = "Electric (ccASHP)"
        Case "electric_other": FormatUtilityType = "Electric (Other)"
        Case "electric_heat_pump": FormatUtilityType = "Electric (Heat Pump)"
        Case "tenant_pays": FormatUtilityType = "Standard"
        Case Else: FormatUtilityType = value
    End Select
End Function

'-------------------------------------------------------------------------------
' APPLY SPECIFIC SCENARIO (for manual selection)
'-------------------------------------------------------------------------------

Public Sub ApplyScenarioByKey(scenarioKey As String)
    ' Applies a specific scenario by its key
    ' Called from scenario sheet buttons

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

            ' API returns integer (60, 80, 130) - convert to decimal for percentage format
            If ami > 1 Then
                amiValue = ami / 100  ' Convert 60 to 0.60
            Else
                amiValue = ami  ' Already decimal
            End If

            ws.Cells(row, amiCol).Value = amiValue
            ws.Cells(row, amiCol).NumberFormat = "0%"  ' Ensure percentage format
            ws.Cells(row, amiCol).Interior.Color = RGB(200, 255, 200)  ' Light green
            updatedCount = updatedCount + 1
        End If
    Next i

    ' Refresh Scenario Manual (live sync) to match what is now applied.
    On Error Resume Next
    Call UpdateManualScenario(False, programNorm)
    On Error GoTo 0

    ' Best-effort learning audit: record the user's chosen scenario.
    On Error Resume Next
    Dim profileKey As String
    profileKey = GetLearningProfileKey(programNorm, mihOption)
    Call LogScenarioApplied(profileKey, programNorm, mihOption, scenarioKey, "USER", scenario)
    On Error GoTo 0

    GoTo ApplyCleanup

ApplyFail:
    Debug.Print "ApplyScenarioByKey Error: " & Err.Description
ApplyCleanup:
    Application.EnableEvents = prevEnableEvents
    g_AMIOptixSuppressEvents = prevSuppress

    MsgBox "Applied scenario '" & FormatScenarioName(scenarioKey) & "'" & vbCrLf & _
           "Updated " & updatedCount & " units.", vbInformation, "AMI Optix"

    ' Switch to data sheet
    ws.Activate
End Sub
