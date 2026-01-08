Attribute VB_Name = "AMI_Optix_ResultsWriter"
'===============================================================================
' AMI OPTIX - Results Writer Module
' Writes optimization results back to Excel
'===============================================================================
Option Explicit

'-------------------------------------------------------------------------------
' APPLY BEST SCENARIO
'-------------------------------------------------------------------------------

Public Sub ApplyBestScenario(result As Object)
    ' Applies the best scenario's AMI assignments to the source data sheet

    Dim scenarios As Object
    Dim bestScenario As Object
    Dim assignments As Object
    Dim ws As Worksheet
    Dim amiCol As Long
    Dim i As Long
    Dim unitId As String
    Dim ami As Double
    Dim row As Long
    Dim updatedCount As Long

    On Error GoTo ErrorHandler

    ' Get scenarios
    Set scenarios = result("scenarios")

    ' Find best scenario (first one returned, already sorted by solver)
    ' The solver returns scenarios keyed by type: "developer_first", "client_oriented", etc.
    Dim scenarioKeys As Variant
    Dim bestKey As String

    ' Priority order for best scenario
    Dim priorities As Variant
    priorities = Array("developer_first", "highest_waami", "client_oriented")

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
        Exit Sub
    End If

    ' Build lookup of unit_id to row number
    Dim unitRows As Object
    Set unitRows = BuildUnitRowLookup(ws)

    ' Apply assignments
    updatedCount = 0

    For i = 1 To assignments.Count
        Dim assignment As Object
        Set assignment = assignments(i)

        unitId = CStr(assignment("unit_id"))
        ami = CDbl(assignment("assigned_ami"))

        ' Find row for this unit
        If unitRows.Exists(unitId) Then
            row = unitRows(unitId)

            ' Write AMI value (as decimal, e.g., 0.60 for 60%)
            ws.Cells(row, amiCol).Value = ami

            ' Highlight the cell
            ws.Cells(row, amiCol).Interior.Color = RGB(255, 255, 200)  ' Light yellow

            updatedCount = updatedCount + 1
        End If
    Next i

    Debug.Print "Applied best scenario: " & bestKey & " - Updated " & updatedCount & " units"
    Exit Sub

ErrorHandler:
    Debug.Print "ApplyBestScenario Error: " & Err.Description
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
    Dim col As Long

    On Error GoTo ErrorHandler

    ' Delete existing scenarios sheet if it exists
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Worksheets("AMI Scenarios").Delete
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = True

    ' Create new sheet
    Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
    ws.Name = "AMI Scenarios"

    Set scenarios = result("scenarios")

    row = 1

    ' Title
    ws.Cells(row, 1).Value = "AMI OPTIMIZATION RESULTS"
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 1).Font.Size = 16
    row = row + 2

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

    For Each scenarioKey In scenarios.keys
        Set scenario = scenarios(scenarioKey)

        ' Scenario header
        ws.Cells(row, 1).Value = "SCENARIO " & scenarioNum & ": " & UCase(CStr(scenarioKey))
        ws.Cells(row, 1).Font.Bold = True
        ws.Cells(row, 1).Font.Size = 14
        ws.Range(ws.Cells(row, 1), ws.Cells(row, 8)).Interior.Color = RGB(200, 220, 255)
        row = row + 1

        ' WAAMI
        ws.Cells(row, 1).Value = "WAAMI:"
        ws.Cells(row, 2).Value = Format(scenario("waami"), "0.00%")
        ws.Cells(row, 2).Font.Bold = True
        row = row + 1

        ' Bands used
        If scenario.Exists("band_combination") Then
            ws.Cells(row, 1).Value = "Bands Used:"
            Dim bands As Object
            Set bands = scenario("band_combination")
            Dim bandStr As String
            bandStr = ""
            Dim b As Long
            For b = 1 To bands.Count
                If bandStr <> "" Then bandStr = bandStr & ", "
                bandStr = bandStr & Format(bands(b) * 100, "0") & "%"
            Next b
            ws.Cells(row, 2).Value = bandStr
        End If
        row = row + 1

        ' Rent totals (if available)
        If scenario.Exists("rent_totals") Then
            Dim rentTotals As Object
            Set rentTotals = scenario("rent_totals")

            If rentTotals.Exists("net_monthly") Then
                ws.Cells(row, 1).Value = "Total Monthly Rent:"
                ws.Cells(row, 2).Value = Format(rentTotals("net_monthly"), "$#,##0.00")
            End If
            row = row + 1

            If rentTotals.Exists("net_annual") Then
                ws.Cells(row, 1).Value = "Total Annual Rent:"
                ws.Cells(row, 2).Value = Format(rentTotals("net_annual"), "$#,##0.00")
            End If
            row = row + 1

            If rentTotals.Exists("gross_monthly") Then
                ws.Cells(row, 1).Value = "Gross Monthly Rent:"
                ws.Cells(row, 2).Value = Format(rentTotals("gross_monthly"), "$#,##0.00")
            End If
            row = row + 1

            If rentTotals.Exists("allowances_monthly") Then
                ws.Cells(row, 1).Value = "Total Utility Allowances:"
                ws.Cells(row, 2).Value = Format(rentTotals("allowances_monthly"), "$#,##0.00")
            End If
            row = row + 1
        End If

        row = row + 1

        ' Unit assignments table header
        ws.Cells(row, 1).Value = "Unit"
        ws.Cells(row, 2).Value = "Bedrooms"
        ws.Cells(row, 3).Value = "AMI"
        ws.Cells(row, 4).Value = "Gross Rent"
        ws.Cells(row, 5).Value = "Allowances"
        ws.Cells(row, 6).Value = "Net Rent"
        ws.Cells(row, 7).Value = "Annual Rent"
        ws.Range(ws.Cells(row, 1), ws.Cells(row, 7)).Font.Bold = True
        ws.Range(ws.Cells(row, 1), ws.Cells(row, 7)).Interior.Color = RGB(230, 230, 230)
        row = row + 1

        ' Unit assignments
        If scenario.Exists("assignments") Then
            Dim assignments As Object
            Set assignments = scenario("assignments")

            Dim a As Long
            For a = 1 To assignments.Count
                Dim assignment As Object
                Set assignment = assignments(a)

                ws.Cells(row, 1).Value = assignment("unit_id")
                ws.Cells(row, 2).Value = assignment("bedrooms")
                ws.Cells(row, 3).Value = Format(assignment("assigned_ami"), "0%")

                If assignment.Exists("gross_rent") Then
                    ws.Cells(row, 4).Value = assignment("gross_rent")
                    ws.Cells(row, 4).NumberFormat = "$#,##0.00"
                End If

                If assignment.Exists("allowance_total") Then
                    ws.Cells(row, 5).Value = assignment("allowance_total")
                    ws.Cells(row, 5).NumberFormat = "$#,##0.00"
                End If

                If assignment.Exists("monthly_rent") Then
                    ws.Cells(row, 6).Value = assignment("monthly_rent")
                    ws.Cells(row, 6).NumberFormat = "$#,##0.00"
                End If

                If assignment.Exists("annual_rent") Then
                    ws.Cells(row, 7).Value = assignment("annual_rent")
                    ws.Cells(row, 7).NumberFormat = "$#,##0.00"
                End If

                row = row + 1
            Next a
        End If

        row = row + 2
        scenarioNum = scenarioNum + 1
    Next scenarioKey

    ' Auto-fit columns
    ws.Columns("A:H").AutoFit

    ' Freeze top row
    ws.Rows(2).Select
    ActiveWindow.FreezePanes = True

    ws.Cells(1, 1).Select

    Debug.Print "Created scenarios sheet with " & (scenarioNum - 1) & " scenarios"
    Exit Sub

ErrorHandler:
    Debug.Print "CreateScenariosSheet Error: " & Err.Description
End Sub

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

    ' Apply
    Dim i As Long
    Dim unitId As String
    Dim ami As Double
    Dim row As Long
    Dim updatedCount As Long

    updatedCount = 0

    For i = 1 To assignments.Count
        Dim assignment As Object
        Set assignment = assignments(i)

        unitId = CStr(assignment("unit_id"))
        ami = CDbl(assignment("assigned_ami"))

        If unitRows.Exists(unitId) Then
            row = unitRows(unitId)
            ws.Cells(row, amiCol).Value = ami
            ws.Cells(row, amiCol).Interior.Color = RGB(200, 255, 200)  ' Light green
            updatedCount = updatedCount + 1
        End If
    Next i

    MsgBox "Applied scenario '" & scenarioKey & "'" & vbCrLf & _
           "Updated " & updatedCount & " units.", vbInformation, "AMI Optix"

    ' Switch to data sheet
    ws.Activate
End Sub
