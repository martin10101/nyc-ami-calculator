Attribute VB_Name = "AMI_Optix_Diagnostic"
'===============================================================================
' AMI OPTIX - DIAGNOSTIC MODULE
' Run DiagnoseAllIssues() to find problems
' This module has NO dependencies on forms or other modules
'===============================================================================
Option Explicit

Public Sub DiagnoseAllIssues()
    ' Master diagnostic - run this to find all problems
    ' Writes results to a new sheet called "AMI_Diagnostic"

    Dim ws As Worksheet
    Dim targetWb As Workbook
    Dim targetSheet As Worksheet
    Dim row As Long
    Dim prevEnableEvents As Boolean
    Dim prevScreenUpdating As Boolean
    Dim prevDisplayAlerts As Boolean

    ' Capture user's context BEFORE we add/activate any sheets.
    Set targetWb = ActiveWorkbook
    Set targetSheet = ActiveSheet

    prevEnableEvents = Application.EnableEvents
    prevScreenUpdating = Application.ScreenUpdating
    prevDisplayAlerts = Application.DisplayAlerts

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    On Error GoTo ErrorHandler

    ' Prefer reusing the existing sheet (avoids 1004 errors when workbook structure is protected).
    On Error Resume Next
    Set ws = targetWb.Worksheets("AMI_Diagnostic")
    On Error GoTo ErrorHandler

    If Not ws Is Nothing Then
        ws.Cells.Clear
    Else
        Set ws = targetWb.Worksheets.Add(After:=targetWb.Worksheets(targetWb.Worksheets.Count))
        ws.Name = "AMI_Diagnostic"
    End If

    row = 1
    ws.Cells(row, 1).Value = "AMI OPTIX DIAGNOSTIC REPORT"
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 1).Font.Size = 16
    row = row + 2

    ' Test 1: Check VBA Project for forms
    ws.Cells(row, 1).Value = "=== TEST 1: VBA PROJECT FORMS ==="
    ws.Cells(row, 1).Font.Bold = True
    row = row + 1
    row = CheckVBAForms(ws, row)
    row = row + 1

    ' Test 2: Check data sheet detection
    ws.Cells(row, 1).Value = "=== TEST 2: DATA SHEET DETECTION ==="
    ws.Cells(row, 1).Font.Bold = True
    row = row + 1
    row = CheckDataSheet(ws, row, targetWb, targetSheet)
    row = row + 1

    ' Test 3: Check column mapping
    ws.Cells(row, 1).Value = "=== TEST 3: COLUMN MAPPING ==="
    ws.Cells(row, 1).Font.Bold = True
    row = row + 1
    row = CheckColumnMapping(ws, row, targetSheet)
    row = row + 1

    ' Test 4: Check AMI values
    ws.Cells(row, 1).Value = "=== TEST 4: AMI VALUE ANALYSIS ==="
    ws.Cells(row, 1).Font.Bold = True
    row = row + 1
    row = CheckAMIValues(ws, row, targetSheet)
    row = row + 1

    ' Test 5: Unit count
    ws.Cells(row, 1).Value = "=== TEST 5: UNIT COUNT ==="
    ws.Cells(row, 1).Font.Bold = True
    row = row + 1
    row = CheckUnitCount(ws, row, targetWb, targetSheet)

    On Error Resume Next
    ws.Columns("A:D").AutoFit
    On Error GoTo ErrorHandler
    ws.Activate

    MsgBox "Diagnostic complete! Check the 'AMI_Diagnostic' sheet for results.", vbInformation, "AMI Optix"
    GoTo Cleanup

ErrorHandler:
    Dim extra As String
    extra = ""
    If Err.Number = 1004 Then
        extra = vbCrLf & vbCrLf & "If this workbook has protected structure, unprotect it first:" & vbCrLf & _
                "Review > Protect Workbook > Unprotect."
    End If
    MsgBox "Diagnostic failed: " & Err.Description & extra, vbCritical, "AMI Optix"

Cleanup:
    Application.DisplayAlerts = prevDisplayAlerts
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
End Sub

Private Function CheckVBAForms(ws As Worksheet, row As Long) As Long
    ' Check for forms in the VBA project that might cause type mismatch
    Dim vbProj As Object
    Dim vbComp As Object
    Dim formCount As Long

    On Error Resume Next
    Set vbProj = ThisWorkbook.VBProject

    If vbProj Is Nothing Then
        ws.Cells(row, 1).Value = "ERROR: Cannot access VBA Project"
        ws.Cells(row, 2).Value = "Enable: Trust Center > Macro Settings > Trust access to VBA project"
        ws.Cells(row, 1).Font.Color = RGB(255, 0, 0)
        CheckVBAForms = row + 1
        Exit Function
    End If
    On Error GoTo 0

    formCount = 0
    For Each vbComp In vbProj.VBComponents
        ' Type 3 = UserForm
        If vbComp.Type = 3 Then
            formCount = formCount + 1
            ws.Cells(row, 1).Value = "FORM FOUND: " & vbComp.Name
            ws.Cells(row, 2).Value = "*** THIS MAY CAUSE TYPE MISMATCH ***"
            ws.Cells(row, 1).Font.Color = RGB(255, 0, 0)
            ws.Cells(row, 2).Font.Color = RGB(255, 0, 0)
            row = row + 1
        End If
    Next vbComp

    If formCount = 0 Then
        ws.Cells(row, 1).Value = "No UserForms found in project"
        ws.Cells(row, 2).Value = "GOOD - Forms are not the cause"
        ws.Cells(row, 1).Font.Color = RGB(0, 128, 0)
        row = row + 1
    Else
        ws.Cells(row, 1).Value = "TOTAL FORMS: " & formCount
        ws.Cells(row, 2).Value = "DELETE ALL FORMS TO FIX TYPE MISMATCH"
        ws.Cells(row, 1).Font.Bold = True
        ws.Cells(row, 2).Font.Bold = True
        ws.Cells(row, 2).Font.Color = RGB(255, 0, 0)
        row = row + 1
    End If

    ' Also list all modules for reference
    ws.Cells(row, 1).Value = "All VBA Components:"
    row = row + 1
    For Each vbComp In vbProj.VBComponents
        ws.Cells(row, 1).Value = "  " & vbComp.Name
        Select Case vbComp.Type
            Case 1: ws.Cells(row, 2).Value = "Standard Module"
            Case 2: ws.Cells(row, 2).Value = "Class Module"
            Case 3: ws.Cells(row, 2).Value = "UserForm"
            Case 100: ws.Cells(row, 2).Value = "Document Module"
            Case Else: ws.Cells(row, 2).Value = "Type " & vbComp.Type
        End Select
        row = row + 1
    Next vbComp

    CheckVBAForms = row
End Function

Private Function CheckDataSheet(ws As Worksheet, row As Long, targetWb As Workbook, targetSheet As Worksheet) As Long
    ' Check which sheet would be used for data
    Dim dataSheet As Worksheet
    Dim preferredNames As Variant
    Dim i As Long
    Dim found As Boolean

    preferredNames = Array("UAP", "PROJECT WORKSHEET", "RentRoll", "Units", "Sheet1", "Data")

    ws.Cells(row, 1).Value = "Active Sheet: " & targetSheet.Name
    row = row + 1

    ' Check preferred sheets
    For i = LBound(preferredNames) To UBound(preferredNames)
        On Error Resume Next
        Set dataSheet = targetWb.Worksheets(CStr(preferredNames(i)))
        On Error GoTo 0

        If Not dataSheet Is Nothing Then
            ws.Cells(row, 1).Value = "Found preferred sheet: " & dataSheet.Name
            ws.Cells(row, 1).Font.Color = RGB(0, 128, 0)
            row = row + 1
            Set dataSheet = Nothing
        End If
    Next i

    CheckDataSheet = row
End Function

Private Function CheckColumnMapping(ws As Worksheet, row As Long, targetSheet As Worksheet) As Long
    ' Check column headers on active sheet
    Dim dataSheet As Worksheet
    Dim col As Long
    Dim headerRow As Long
    Dim cellValue As String
    Dim maxCol As Long

    Set dataSheet = targetSheet

    ' Find header row (first row with multiple non-empty cells)
    headerRow = FindHeaderRowDiag(dataSheet)

    If headerRow = 0 Then
        ws.Cells(row, 1).Value = "ERROR: Could not find header row"
        ws.Cells(row, 1).Font.Color = RGB(255, 0, 0)
        CheckColumnMapping = row + 1
        Exit Function
    End If

    ws.Cells(row, 1).Value = "Header row found at: " & headerRow
    row = row + 1

    maxCol = Application.Min(20, dataSheet.UsedRange.Columns.Count)

    ws.Cells(row, 1).Value = "Column"
    ws.Cells(row, 2).Value = "Header Value"
    ws.Cells(row, 3).Value = "Detected As"
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 3)).Font.Bold = True
    row = row + 1

    For col = 1 To maxCol
        cellValue = Trim(CStr(dataSheet.Cells(headerRow, col).Value))
        If cellValue <> "" Then
            ws.Cells(row, 1).Value = col
            ws.Cells(row, 2).Value = cellValue
            ws.Cells(row, 3).Value = IdentifyColumn(UCase(cellValue))

            ' Highlight AMI column
            If InStr(UCase(cellValue), "AMI") > 0 Then
                ws.Cells(row, 2).Font.Bold = True
                ws.Cells(row, 2).Font.Color = RGB(0, 0, 255)
            End If
            row = row + 1
        End If
    Next col

    CheckColumnMapping = row
End Function

Private Function FindHeaderRowDiag(dataSheet As Worksheet) As Long
    Dim row As Long
    Dim col As Long
    Dim nonEmptyCount As Long

    For row = 1 To 20
        nonEmptyCount = 0
        For col = 1 To 10
            If Trim(CStr(dataSheet.Cells(row, col).Value)) <> "" Then
                nonEmptyCount = nonEmptyCount + 1
            End If
        Next col

        If nonEmptyCount >= 3 Then
            FindHeaderRowDiag = row
            Exit Function
        End If
    Next row

    FindHeaderRowDiag = 0
End Function

Private Function IdentifyColumn(header As String) As String
    If InStr(header, "UNIT") > 0 Or InStr(header, "APT") > 0 Then
        IdentifyColumn = "UNIT ID"
    ElseIf InStr(header, "BED") > 0 Or header = "BR" Then
        IdentifyColumn = "BEDROOMS"
    ElseIf InStr(header, "SF") > 0 Or InStr(header, "SQFT") > 0 Then
        IdentifyColumn = "NET SF"
    ElseIf InStr(header, "FLOOR") > 0 Or InStr(header, "FLR") > 0 Then
        IdentifyColumn = "FLOOR"
    ElseIf InStr(header, "AMI") > 0 Then
        If InStr(header, "AFTER") > 0 Then
            IdentifyColumn = "AMI AFTER (EXCLUDED)"
        Else
            IdentifyColumn = "AMI (INPUT)"
        End If
    ElseIf InStr(header, "BALCON") > 0 Then
        IdentifyColumn = "BALCONY"
    Else
        IdentifyColumn = "-"
    End If
End Function

Private Function CheckAMIValues(ws As Worksheet, row As Long, targetSheet As Worksheet) As Long
    ' Analyze AMI column values
    Dim dataSheet As Worksheet
    Dim amiCol As Long
    Dim headerRow As Long
    Dim lastRow As Long
    Dim r As Long
    Dim cellValue As Variant
    Dim validCount As Long
    Dim emptyCount As Long
    Dim textCount As Long
    Dim zeroCount As Long
    Dim col As Long

    Set dataSheet = targetSheet
    headerRow = FindHeaderRowDiag(dataSheet)

    If headerRow = 0 Then
        ws.Cells(row, 1).Value = "Cannot analyze - no header row"
        CheckAMIValues = row + 1
        Exit Function
    End If

    ' Find AMI column
    amiCol = 0
    For col = 1 To 20
        Dim hdr As String
        hdr = UCase(Trim(CStr(dataSheet.Cells(headerRow, col).Value)))
        If InStr(hdr, "AMI") > 0 And InStr(hdr, "AFTER") = 0 Then
            amiCol = col
            Exit For
        End If
    Next col

    If amiCol = 0 Then
        ws.Cells(row, 1).Value = "ERROR: No AMI column found!"
        ws.Cells(row, 1).Font.Color = RGB(255, 0, 0)
        CheckAMIValues = row + 1
        Exit Function
    End If

    ws.Cells(row, 1).Value = "AMI Column: " & amiCol & " (Header: " & dataSheet.Cells(headerRow, amiCol).Value & ")"
    row = row + 1

    ' Find last row
    lastRow = dataSheet.Cells(dataSheet.Rows.Count, amiCol).End(xlUp).row
    ws.Cells(row, 1).Value = "Data rows: " & headerRow + 1 & " to " & lastRow
    row = row + 1

    ' Analyze values
    validCount = 0
    emptyCount = 0
    textCount = 0
    zeroCount = 0

    ws.Cells(row, 1).Value = "Row"
    ws.Cells(row, 2).Value = "Raw Value"
    ws.Cells(row, 3).Value = "Type"
    ws.Cells(row, 4).Value = "Status"
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 4)).Font.Bold = True
    row = row + 1

    For r = headerRow + 1 To Application.Min(lastRow, headerRow + 30)  ' First 30 rows
        cellValue = dataSheet.Cells(r, amiCol).Value

        ws.Cells(row, 1).Value = r
        ws.Cells(row, 2).Value = "'" & CStr(cellValue)
        ws.Cells(row, 3).Value = TypeName(cellValue)

        If IsEmpty(cellValue) Or Trim(CStr(cellValue)) = "" Then
            ws.Cells(row, 4).Value = "EMPTY - SKIPPED"
            ws.Cells(row, 4).Font.Color = RGB(255, 0, 0)
            emptyCount = emptyCount + 1
        ElseIf Not IsNumeric(cellValue) Then
            ' Check if it's a percentage string like "60%"
            Dim strVal As String
            strVal = Trim(CStr(cellValue))
            If Right(strVal, 1) = "%" Then
                strVal = Left(strVal, Len(strVal) - 1)
            End If
            If IsNumeric(strVal) Then
                ws.Cells(row, 4).Value = "VALID (text %)"
                ws.Cells(row, 4).Font.Color = RGB(0, 128, 0)
                validCount = validCount + 1
            Else
                ws.Cells(row, 4).Value = "TEXT - SKIPPED"
                ws.Cells(row, 4).Font.Color = RGB(255, 0, 0)
                textCount = textCount + 1
            End If
        ElseIf CDbl(cellValue) <= 0 Then
            ws.Cells(row, 4).Value = "ZERO/NEG - SKIPPED"
            ws.Cells(row, 4).Font.Color = RGB(255, 0, 0)
            zeroCount = zeroCount + 1
        Else
            ws.Cells(row, 4).Value = "VALID"
            ws.Cells(row, 4).Font.Color = RGB(0, 128, 0)
            validCount = validCount + 1
        End If
        row = row + 1
    Next r

    row = row + 1
    ws.Cells(row, 1).Value = "SUMMARY:"
    ws.Cells(row, 1).Font.Bold = True
    row = row + 1
    ws.Cells(row, 1).Value = "Valid AMI values: " & validCount
    If validCount < 5 Then
        ws.Cells(row, 2).Value = "*** TOO FEW - THIS IS THE PROBLEM ***"
        ws.Cells(row, 2).Font.Color = RGB(255, 0, 0)
        ws.Cells(row, 2).Font.Bold = True
    End If
    row = row + 1
    ws.Cells(row, 1).Value = "Empty cells: " & emptyCount
    row = row + 1
    ws.Cells(row, 1).Value = "Text (non-numeric): " & textCount
    row = row + 1
    ws.Cells(row, 1).Value = "Zero/Negative: " & zeroCount
    row = row + 1

    CheckAMIValues = row
End Function

Private Function CheckUnitCount(ws As Worksheet, row As Long, targetWb As Workbook, targetSheet As Worksheet) As Long
    ' Try to read units using the actual DataReader logic
    Dim units As Collection
    Dim prevSheet As Worksheet

    On Error Resume Next
    Set prevSheet = ActiveSheet
    targetWb.Activate
    targetSheet.Activate
    Set units = ReadUnitData()
    If Not prevSheet Is Nothing Then prevSheet.Activate
    On Error GoTo 0

    If units Is Nothing Then
        ws.Cells(row, 1).Value = "ReadUnitData() returned Nothing"
        ws.Cells(row, 1).Font.Color = RGB(255, 0, 0)
    ElseIf units.Count = 0 Then
        ws.Cells(row, 1).Value = "ReadUnitData() returned 0 units"
        ws.Cells(row, 1).Font.Color = RGB(255, 0, 0)
    Else
        ws.Cells(row, 1).Value = "ReadUnitData() found " & units.Count & " units"
        If units.Count < 10 Then
            ws.Cells(row, 1).Font.Color = RGB(255, 165, 0)  ' Orange warning
        Else
            ws.Cells(row, 1).Font.Color = RGB(0, 128, 0)
        End If
        row = row + 1

        ' List first 10 units
        Dim i As Long
        Dim unit As Object
        For i = 1 To Application.Min(10, units.Count)
            Set unit = units(i)
            ws.Cells(row, 1).Value = "  Unit " & i & ":"
            ws.Cells(row, 2).Value = "ID=" & unit("unit_id")
            ws.Cells(row, 3).Value = "BR=" & unit("bedrooms")
            ws.Cells(row, 4).Value = "SF=" & unit("net_sf")
            row = row + 1
        Next i
    End If

    CheckUnitCount = row + 1
End Function
