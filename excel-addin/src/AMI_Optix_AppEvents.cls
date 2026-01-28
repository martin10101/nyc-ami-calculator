VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AMI_Optix_AppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' AMI OPTIX - Application Events
'===============================================================================
Option Explicit

Public WithEvents App As Application

Private m_LastSelectionSheet As String
Private m_LastSelectionAddress As String
Private m_LastSelectionValue As Variant

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo SafeExit

    If g_AMIOptixSuppressEvents Then Exit Sub
    If Target Is Nothing Then Exit Sub
    If TypeName(Sh) <> "Worksheet" Then Exit Sub

    m_LastSelectionSheet = CStr(Sh.Name)
    m_LastSelectionAddress = Target.Address(False, False)

    ' Store old values for revert support (best effort).
    ' Avoid huge selections (can be slow / memory-heavy).
    If Target.CountLarge <= 5000 Then
        m_LastSelectionValue = Target.Value
    Else
        m_LastSelectionValue = Empty
    End If

SafeExit:
End Sub

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo SafeExit

    If g_AMIOptixSuppressEvents Then Exit Sub
    If Target Is Nothing Then Exit Sub
    If TypeName(Sh) <> "Worksheet" Then Exit Sub

    Dim ws As Worksheet
    Set ws = Sh

    Dim programNorm As String
    programNorm = DetectProgramFromWorkbook()

    If UCase(ws.Name) = UCase("AMI Scenarios") Then
        If ChangeIsManualScenarioAMI(ws, Target) Then
            HandleManualScenarioEdit ws, Target
        End If
        Exit Sub
    End If

    If programNorm = "UAP" Then
        If UCase(ws.Name) = "UAP" Then
            If ChangeIsInAMIColumn(ws, Target) Then
                HandleDataSheetAmiChange ws, Target, "UAP"
            End If
            Exit Sub
        End If
    ElseIf programNorm = "MIH" Then
        ' MIH: the editable unit table is typically on "RentRoll" (or another sheet with Unit/Bed/SF/AMI).
        If ChangeIsInAMIColumn(ws, Target) Then
            HandleDataSheetAmiChange ws, Target, "MIH"
            Exit Sub
        End If
    End If

SafeExit:
End Sub

Private Function DetectProgramFromWorkbook() As String
    ' Best-effort program detection:
    ' - If a sheet named "MIH" exists, treat workbook as MIH.
    ' - Otherwise, treat as UAP.
    On Error GoTo Fail

    DetectProgramFromWorkbook = "UAP"

    If ActiveWorkbook Is Nothing Then Exit Function

    Dim wsMIH As Worksheet
    Set wsMIH = Nothing
    On Error Resume Next
    Set wsMIH = ActiveWorkbook.Worksheets("MIH")
    On Error GoTo Fail
    If Not wsMIH Is Nothing Then DetectProgramFromWorkbook = "MIH"

    Exit Function

Fail:
    DetectProgramFromWorkbook = "UAP"
End Function

Private Sub HandleDataSheetAmiChange(ws As Worksheet, target As Range, programNorm As String)
    ' Sync: when the user edits the AMI column on the data sheet, update the Scenario Manual block.
    On Error GoTo SafeExit

    g_AMIOptixSuppressEvents = True

    Dim oldValue As Variant
    Dim hasOldValue As Boolean
    hasOldValue = (UCase(m_LastSelectionSheet) = UCase(ws.Name) And m_LastSelectionAddress = target.Address(False, False))
    If hasOldValue Then oldValue = m_LastSelectionValue

    Dim ok As Boolean
    ok = UpdateManualScenario(False, programNorm)

    If (Not ok) And g_AMIOptixLastManualScenarioInvalid Then
        ' Revert the invalid edit without relying on Application.Undo (unreliable in event handlers).
        If hasOldValue Then
            Application.EnableEvents = False
            target.Value = oldValue

            ' Best-effort: keep AMI column formatted as percent.
            Dim headerRow As Long
            Dim amiCol As Long
            headerRow = FindHeaderRow(ws)
            amiCol = 0
            If headerRow > 0 Then amiCol = FindAMIColumn(ws, headerRow)
            If amiCol > 0 Then
                Dim amiCells As Range
                Set amiCells = Intersect(target, ws.Columns(amiCol))
                If Not amiCells Is Nothing Then
                    amiCells.NumberFormat = "0%"
                End If
            End If
            Application.EnableEvents = True

            ' Refresh the manual block to match the reverted values.
            Call UpdateManualScenario(False, programNorm)

            ' Keep last selection cache consistent for subsequent edits.
            m_LastSelectionValue = oldValue
        End If
    ElseIf ok Then
        ' Keep last selection cache consistent for subsequent edits.
        m_LastSelectionSheet = CStr(ws.Name)
        m_LastSelectionAddress = target.Address(False, False)
        m_LastSelectionValue = target.Value
    End If

SafeExit:
    g_AMIOptixSuppressEvents = False
End Sub

Private Function ChangeIsInAMIColumn(ws As Worksheet, target As Range) As Boolean
    Dim headerRow As Long
    Dim amiCol As Long
    headerRow = FindHeaderRow(ws)
    If headerRow = 0 Then Exit Function
    amiCol = FindAMIColumn(ws, headerRow)
    If amiCol = 0 Then Exit Function

    Dim amiData As Range
    Set amiData = ws.Range(ws.Cells(headerRow + 1, amiCol), ws.Cells(ws.Rows.Count, amiCol))
    If Not Intersect(target, amiData) Is Nothing Then
        ChangeIsInAMIColumn = True
    End If
End Function

Private Function ChangeIsManualScenarioAMI(ws As Worksheet, target As Range) As Boolean
    Dim amiCells As Range
    Set amiCells = Intersect(target, ws.Columns(6))
    If amiCells Is Nothing Then Exit Function

    Dim headerRow As Long
    headerRow = FindManualTableHeaderRow(ws)
    If headerRow = 0 Then Exit Function

    Dim lastRow As Long
    lastRow = FindManualTableLastRow(ws, headerRow)
    If lastRow = 0 Then Exit Function

    Dim tableAmi As Range
    Set tableAmi = ws.Range(ws.Cells(headerRow + 1, 6), ws.Cells(lastRow, 6))

    Dim inTable As Range
    Set inTable = Intersect(amiCells, tableAmi)
    If inTable Is Nothing Then Exit Function

    ' At least one changed cell must correspond to a real unit row.
    Dim c As Range
    For Each c In inTable.Cells
        If Trim(CStr(ws.Cells(c.row, 1).Value)) <> "" Then
            ChangeIsManualScenarioAMI = True
            Exit Function
        End If
    Next c

End Function

Private Sub HandleManualScenarioEdit(ws As Worksheet, target As Range)
    On Error GoTo Fail

    ' Support paste/fill-down by handling all changed AMI cells in the manual table at once.
    If target.CountLarge > 1 Then
        HandleManualScenarioEditRange ws, target
        Exit Sub
    End If

    Dim unitId As String
    unitId = Trim(CStr(ws.Cells(target.row, 1).Value))
    If unitId = "" Then Exit Sub

    Dim newAmi As Double
    newAmi = ParseAmiValue(target.Value)
    If newAmi <= 0 Then Exit Sub

    ' Validate via /api/evaluate before applying to the data sheet.
    Dim prevSheet As Worksheet
    Set prevSheet = ActiveSheet

    Dim programNorm As String
    programNorm = DetectProgramFromWorkbook()

    Dim mihOption As String
    Dim mihResidentialSF As Double
    Dim mihMaxBandPercent As Long
    mihOption = ""
    mihResidentialSF = 0
    mihMaxBandPercent = 0
    If programNorm = "MIH" Then
        If Not TryReadMIHInputs(mihOption, mihResidentialSF, mihMaxBandPercent) Then Exit Sub
    End If

    Dim dataWs As Worksheet
    Set dataWs = Nothing
    On Error Resume Next
    If programNorm = "MIH" Then
        Set dataWs = ActiveWorkbook.Worksheets("RentRoll")
        If dataWs Is Nothing Then Set dataWs = ActiveWorkbook.Worksheets("MIH")
    Else
        Set dataWs = ActiveWorkbook.Worksheets("UAP")
    End If
    On Error GoTo Fail
    If dataWs Is Nothing Then Exit Sub

    dataWs.Activate
    Dim units As Collection
    Set units = ReadUnitData()
    prevSheet.Activate

    If units Is Nothing Or units.Count = 0 Then Exit Sub

    Dim i As Long
    Dim u As Object
    For i = 1 To units.Count
        Set u = units(i)
        If CStr(u("unit_id")) = unitId Then
            u("client_ami") = newAmi
            Exit For
        End If
    Next i

    Dim utilities As Object
    Set utilities = GetUtilitySelectionsForProgram(programNorm)

    Dim payload As String
    payload = BuildEvaluatePayloadV2(units, utilities, programNorm, mihOption, mihResidentialSF, mihMaxBandPercent)

    Dim response As String
    response = CallEvaluateAPI(payload)
    If response = "" Then Exit Sub

    Dim evalResult As Object
    Set evalResult = ParseJSON(response)
    If evalResult Is Nothing Then Exit Sub

    If evalResult.Exists("success") Then
        If evalResult("success") = False Then
            Dim oldValue As Variant
            Dim hasOldValue As Boolean
            hasOldValue = (UCase(m_LastSelectionSheet) = UCase(ws.Name) And m_LastSelectionAddress = target.Address(False, False))
            If hasOldValue Then oldValue = m_LastSelectionValue

            Dim fallbackValue As Variant
            Dim hasFallback As Boolean
            hasFallback = False
            If Not hasOldValue Then
                hasFallback = TryGetDataSheetAmiValue(programNorm, unitId, fallbackValue)
            End If

            g_AMIOptixSuppressEvents = True
            Application.EnableEvents = False
            If hasOldValue Then
                target.Value = oldValue
                If IsNumeric(oldValue) Then target.NumberFormat = "0%"
                m_LastSelectionValue = oldValue
            ElseIf hasFallback Then
                target.Value = fallbackValue
                If IsNumeric(fallbackValue) Then target.NumberFormat = "0%"
                m_LastSelectionValue = fallbackValue
            End If
            Application.EnableEvents = True
            g_AMIOptixSuppressEvents = False

            Dim msg As String
            msg = "Invalid manual change. Reverted." & vbCrLf & vbCrLf
            If evalResult.Exists("errors") Then
                Dim errs As Object
                Set errs = evalResult("errors")
                For i = 1 To errs.Count
                    msg = msg & "- " & errs(i) & vbCrLf
                Next i
            End If
            MsgBox msg, vbExclamation, "AMI Optix"
            Exit Sub
        End If
    End If

    ' Apply to the data sheet (last-edit wins) and refresh manual scenario block.
    g_AMIOptixSuppressEvents = True
    ApplyAmiToDataSheet programNorm, unitId, newAmi
    Call UpdateManualScenario(False, programNorm)
    g_AMIOptixSuppressEvents = False
    Exit Sub

Fail:
    g_AMIOptixSuppressEvents = False
End Sub

Private Sub HandleManualScenarioEditRange(ws As Worksheet, target As Range)
    ' Handles multi-cell changes in the manual scenario AMI column (e.g., paste/fill-down).
    On Error GoTo Fail

    Dim headerRow As Long
    headerRow = FindManualTableHeaderRow(ws)
    If headerRow = 0 Then Exit Sub

    Dim lastRow As Long
    lastRow = FindManualTableLastRow(ws, headerRow)
    If lastRow = 0 Then Exit Sub

    Dim amiCells As Range
    Set amiCells = Intersect(target, ws.Columns(6))
    If amiCells Is Nothing Then Exit Sub

    Dim tableAmi As Range
    Set tableAmi = ws.Range(ws.Cells(headerRow + 1, 6), ws.Cells(lastRow, 6))

    Dim inTable As Range
    Set inTable = Intersect(amiCells, tableAmi)
    If inTable Is Nothing Then Exit Sub

    Dim changes As Object
    Set changes = CreateObject("Scripting.Dictionary") ' unit_id -> newAmi

    Dim changeRows As Object
    Set changeRows = CreateObject("Scripting.Dictionary") ' unit_id -> row

    Dim c As Range
    For Each c In inTable.Cells
        Dim unitId As String
        unitId = Trim(CStr(ws.Cells(c.row, 1).Value))
        If unitId = "" Then GoTo NextCell

        Dim newAmi As Double
        newAmi = ParseAmiValue(c.Value)
        If newAmi <= 0 Then GoTo NextCell

        changes(unitId) = newAmi
        changeRows(unitId) = CLng(c.row)

NextCell:
    Next c

    If changes.Count = 0 Then Exit Sub

    Dim prevSheet As Worksheet
    Set prevSheet = ActiveSheet

    Dim programNorm As String
    programNorm = DetectProgramFromWorkbook()

    Dim mihOption As String
    Dim mihResidentialSF As Double
    Dim mihMaxBandPercent As Long
    mihOption = ""
    mihResidentialSF = 0
    mihMaxBandPercent = 0
    If programNorm = "MIH" Then
        If Not TryReadMIHInputs(mihOption, mihResidentialSF, mihMaxBandPercent) Then Exit Sub
    End If

    ' Activate a likely data sheet so ReadUnitData doesn't accidentally read from AMI Scenarios.
    Dim dataWs As Worksheet
    Set dataWs = Nothing
    On Error Resume Next
    If programNorm = "MIH" Then
        Set dataWs = ActiveWorkbook.Worksheets("RentRoll")
        If dataWs Is Nothing Then Set dataWs = ActiveWorkbook.Worksheets("UAP")
        If dataWs Is Nothing Then Set dataWs = ActiveWorkbook.Worksheets("PROJECT WORKSHEET")
        If dataWs Is Nothing Then Set dataWs = ActiveWorkbook.Worksheets("MIH")
    Else
        Set dataWs = ActiveWorkbook.Worksheets("UAP")
    End If
    On Error GoTo Fail
    If Not dataWs Is Nothing Then dataWs.Activate

    Dim units As Collection
    Set units = ReadUnitData()

    ' Capture which sheet actually contained the unit table.
    On Error Resume Next
    Set dataWs = GetDataSheet()
    On Error GoTo Fail

    prevSheet.Activate

    If units Is Nothing Or units.Count = 0 Then Exit Sub

    Dim i As Long
    Dim u As Object
    For i = 1 To units.Count
        Set u = units(i)
        Dim uid As String
        uid = CStr(u("unit_id"))
        If changes.Exists(uid) Then
            u("client_ami") = CDbl(changes(uid))
        End If
    Next i

    Dim utilities As Object
    Set utilities = GetUtilitySelectionsForProgram(programNorm)

    Dim payload As String
    payload = BuildEvaluatePayloadV2(units, utilities, programNorm, mihOption, mihResidentialSF, mihMaxBandPercent)

    Dim response As String
    response = CallEvaluateAPI(payload)
    If response = "" Then Exit Sub

    Dim evalResult As Object
    Set evalResult = ParseJSON(response)
    If evalResult Is Nothing Then Exit Sub

    If evalResult.Exists("success") Then
        If evalResult("success") = False Then
            ' Revert all changed cells to the current data sheet value (best-effort).
            g_AMIOptixSuppressEvents = True
            Application.EnableEvents = False

            Dim key As Variant
            For Each key In changes.Keys
                Dim fallbackValue As Variant
                Dim hasFallback As Boolean
                hasFallback = False
                If Not dataWs Is Nothing Then
                    hasFallback = TryGetAmiValueFromWorksheet(dataWs, CStr(key), fallbackValue)
                End If
                If hasFallback Then
                    ws.Cells(CLng(changeRows(key)), 6).Value = fallbackValue
                    ws.Cells(CLng(changeRows(key)), 6).NumberFormat = "0%"
                End If
            Next key

            Application.EnableEvents = True
            g_AMIOptixSuppressEvents = False

            Dim msg As String
            msg = "Invalid manual change. Reverted." & vbCrLf & vbCrLf
            If evalResult.Exists("errors") Then
                Dim errs As Object
                Set errs = evalResult("errors")
                For i = 1 To errs.Count
                    msg = msg & "- " & errs(i) & vbCrLf
                Next i
            End If
            MsgBox msg, vbExclamation, "AMI Optix"
            Exit Sub
        End If
    End If

    ' Apply all changes to the data sheet and refresh the manual scenario.
    g_AMIOptixSuppressEvents = True
    Application.EnableEvents = False

    If Not dataWs Is Nothing Then
        Dim k2 As Variant
        For Each k2 In changes.Keys
            ApplyAmiToWorksheet dataWs, CStr(k2), CDbl(changes(k2))
        Next k2
    End If

    Application.EnableEvents = True
    g_AMIOptixSuppressEvents = False

    Call UpdateManualScenario(False, programNorm)
    Exit Sub

Fail:
    g_AMIOptixSuppressEvents = False
    On Error Resume Next
    Application.EnableEvents = True
End Sub

Private Function TryGetDataSheetAmiValue(programNorm As String, unitId As String, ByRef amiValue As Variant) As Boolean
    On Error GoTo Fail

    Dim ws As Worksheet
    Set ws = GetProgramDataSheet(programNorm)
    If ws Is Nothing Then Exit Function

    Dim headerRow As Long
    headerRow = FindHeaderRow(ws)
    If headerRow = 0 Then Exit Function

    Dim unitCol As Long
    unitCol = FindUnitIdColumn(ws, headerRow)
    If unitCol = 0 Then unitCol = 2

    Dim amiCol As Long
    amiCol = FindAMIColumn(ws, headerRow)
    If amiCol = 0 Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, unitCol).End(xlUp).Row

    Dim r As Long
    For r = headerRow + 1 To lastRow
        If Trim(CStr(ws.Cells(r, unitCol).Value)) = unitId Then
            amiValue = ws.Cells(r, amiCol).Value
            TryGetDataSheetAmiValue = True
            Exit Function
        End If
    Next r

Fail:
    TryGetDataSheetAmiValue = False
End Function

Private Sub ApplyAmiToDataSheet(programNorm As String, unitId As String, newAmi As Double)
    On Error GoTo Fail

    Dim ws As Worksheet
    Set ws = GetProgramDataSheet(programNorm)
    If ws Is Nothing Then Exit Sub

    Dim headerRow As Long
    headerRow = FindHeaderRow(ws)
    If headerRow = 0 Then Exit Sub

    Dim unitCol As Long
    unitCol = FindUnitIdColumn(ws, headerRow)
    If unitCol = 0 Then unitCol = 2

    Dim amiCol As Long
    amiCol = FindAMIColumn(ws, headerRow)
    If amiCol = 0 Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, unitCol).End(xlUp).Row

    Dim r As Long
    For r = headerRow + 1 To lastRow
        If Trim(CStr(ws.Cells(r, unitCol).Value)) = unitId Then
            ws.Cells(r, amiCol).Value = newAmi
            ws.Cells(r, amiCol).NumberFormat = "0%"
            Exit Sub
        End If
    Next r

Fail:
End Sub

Private Function TryGetAmiValueFromWorksheet(ws As Worksheet, unitId As String, ByRef amiValue As Variant) As Boolean
    On Error GoTo Fail

    If ws Is Nothing Then Exit Function

    Dim headerRow As Long
    headerRow = FindHeaderRow(ws)
    If headerRow = 0 Then Exit Function

    Dim unitCol As Long
    unitCol = FindUnitIdColumn(ws, headerRow)
    If unitCol = 0 Then unitCol = 2

    Dim amiCol As Long
    amiCol = FindAMIColumn(ws, headerRow)
    If amiCol = 0 Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, unitCol).End(xlUp).Row

    Dim r As Long
    For r = headerRow + 1 To lastRow
        If Trim(CStr(ws.Cells(r, unitCol).Value)) = unitId Then
            amiValue = ws.Cells(r, amiCol).Value
            TryGetAmiValueFromWorksheet = True
            Exit Function
        End If
    Next r

Fail:
    TryGetAmiValueFromWorksheet = False
End Function

Private Sub ApplyAmiToWorksheet(ws As Worksheet, unitId As String, newAmi As Double)
    On Error GoTo Fail

    If ws Is Nothing Then Exit Sub

    Dim headerRow As Long
    headerRow = FindHeaderRow(ws)
    If headerRow = 0 Then Exit Sub

    Dim unitCol As Long
    unitCol = FindUnitIdColumn(ws, headerRow)
    If unitCol = 0 Then unitCol = 2

    Dim amiCol As Long
    amiCol = FindAMIColumn(ws, headerRow)
    If amiCol = 0 Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, unitCol).End(xlUp).Row

    Dim r As Long
    For r = headerRow + 1 To lastRow
        If Trim(CStr(ws.Cells(r, unitCol).Value)) = unitId Then
            ws.Cells(r, amiCol).Value = newAmi
            ws.Cells(r, amiCol).NumberFormat = "0%"
            Exit Sub
        End If
    Next r

Fail:
End Sub

Private Function GetProgramDataSheet(programNorm As String) As Worksheet
    On Error GoTo Fail

    If ActiveWorkbook Is Nothing Then Exit Function

    Dim ws As Worksheet
    Set ws = Nothing

    On Error Resume Next
    If UCase$(Trim$(programNorm)) = "MIH" Then
        Set ws = ActiveWorkbook.Worksheets("RentRoll")
        If ws Is Nothing Then Set ws = ActiveWorkbook.Worksheets("UAP")
        If ws Is Nothing Then Set ws = ActiveWorkbook.Worksheets("PROJECT WORKSHEET")
        If ws Is Nothing Then Set ws = ActiveWorkbook.Worksheets("MIH")
    Else
        Set ws = ActiveWorkbook.Worksheets("UAP")
    End If
    On Error GoTo Fail

    Set GetProgramDataSheet = ws
    Exit Function

Fail:
    Set GetProgramDataSheet = Nothing
End Function

Private Function ParseAmiValue(value As Variant) As Double
    On Error GoTo Fail
    If IsNumeric(value) Then
        ParseAmiValue = CDbl(value)
        If ParseAmiValue > 1 Then ParseAmiValue = ParseAmiValue / 100
        Exit Function
    End If
    Dim s As String
    s = Trim(CStr(value))
    If Right(s, 1) = "%" Then
        s = Left(s, Len(s) - 1)
        If IsNumeric(s) Then
            ParseAmiValue = CDbl(s) / 100
            Exit Function
        End If
    End If
Fail:
    ParseAmiValue = 0
End Function

Private Function FindManualTableHeaderRow(ws As Worksheet) As Long
    ' Finds the header row for the "Scenario Manual (Live Sync)" unit table.
    ' We intentionally anchor on the "SCENARIO MANUAL (LIVE SYNC)" label to avoid
    ' accidentally matching the scenario tables lower on the sheet.
    Dim r As Long
    Dim labelRow As Long
    labelRow = 0

    For r = 1 To 300
        If UCase$(Trim$(CStr(ws.Cells(r, 1).Value))) = "SCENARIO MANUAL (LIVE SYNC)" Then
            labelRow = r
            Exit For
        End If
    Next r

    If labelRow > 0 Then
        For r = labelRow + 1 To labelRow + 80
            If UCase$(Trim$(CStr(ws.Cells(r, 1).Value))) = "UNIT" And UCase$(Trim$(CStr(ws.Cells(r, 6).Value))) = "AMI" Then
                FindManualTableHeaderRow = r
                Exit Function
            End If
        Next r
    End If

    ' Fallback: search near the top of the sheet.
    For r = 1 To 200
        If UCase$(Trim$(CStr(ws.Cells(r, 1).Value))) = "UNIT" And UCase$(Trim$(CStr(ws.Cells(r, 6).Value))) = "AMI" Then
            FindManualTableHeaderRow = r
            Exit Function
        End If
    Next r

    FindManualTableHeaderRow = 0
End Function

Private Function FindManualTableLastRow(ws As Worksheet, headerRow As Long) As Long
    On Error GoTo Fail

    Dim r As Long
    For r = headerRow + 1 To headerRow + 20000
        Dim unitId As String
        unitId = Trim$(CStr(ws.Cells(r, 1).Value))

        If unitId = "" Then
            FindManualTableLastRow = r - 1
            Exit Function
        End If

        Dim v As String
        v = UCase$(unitId)
        If v Like "SCENARIO #*" Then
            FindManualTableLastRow = r - 1
            Exit Function
        End If
    Next r

Fail:
    FindManualTableLastRow = 0
End Function

Private Function FindHeaderRow(ws As Worksheet) As Long
    Dim r As Long, c As Long
    For r = 1 To 30
        Dim hasApt As Boolean, hasBed As Boolean, hasSf As Boolean
        hasApt = False: hasBed = False: hasSf = False
        For c = 1 To 20
            Dim v As String
            v = UCase(Trim(CStr(ws.Cells(r, c).Value)))
            If v = "APT" Or v = "APT #" Or v = "UNIT" Or v = "UNIT ID" Then hasApt = True
            If InStr(v, "BED") > 0 Or v = "BR" Then hasBed = True
            If InStr(v, "NET") > 0 And InStr(v, "SF") > 0 Then hasSf = True
            If hasApt And hasBed And hasSf Then
                FindHeaderRow = r
                Exit Function
            End If
        Next c
    Next r
    FindHeaderRow = 0
End Function

Private Function FindUnitIdColumn(ws As Worksheet, headerRow As Long) As Long
    Dim c As Long
    For c = 1 To 30
        Dim v As String
        v = UCase(Trim(CStr(ws.Cells(headerRow, c).Value)))
        If v = "APT" Or v = "APT #" Or v = "UNIT" Or v = "UNIT ID" Then
            FindUnitIdColumn = c
            Exit Function
        End If
    Next c
    FindUnitIdColumn = 0
End Function

Private Function FindAMIColumn(ws As Worksheet, headerRow As Long) As Long
    Dim c As Long
    For c = 1 To 30
        Dim v As String
        v = UCase(Trim(CStr(ws.Cells(headerRow, c).Value)))
        If InStr(v, "AMI") > 0 And InStr(v, "AFTER") = 0 Then
            FindAMIColumn = c
            Exit Function
        End If
    Next c
    FindAMIColumn = 0
End Function
