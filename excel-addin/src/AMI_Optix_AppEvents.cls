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

Private Sub App_SheetActivate(ByVal Sh As Object)
    ' Keep AMI Optix sheets visible even in client workbooks that hide sheets via macros.
    On Error GoTo SafeExit

    Dim wbName As String
    wbName = ""
    On Error Resume Next
    wbName = CStr(Sh.Parent.Name)
    On Error GoTo SafeExit

    If Trim$(wbName) <> "" Then
        g_AMIOptixVisibilityWorkbookName = wbName
    End If

    Call AMI_Optix_EnsureSheetsVisible
    Call AMI_Optix_ScheduleEnsureSheetsVisible(wbName)

SafeExit:
End Sub

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
    If Not GetLiveSyncEnabled() Then Exit Sub

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
    Dim mihOption As String
    mihOption = ""

    Dim kind As String
    kind = DetectWorkbookKind(mihOption)

    If kind = "MIH" Or kind = "MIH_INVALID" Then
        DetectProgramFromWorkbook = "MIH"
    Else
        DetectProgramFromWorkbook = "UAP"
    End If
End Function

Private Sub RefreshManualWorkingCopyFromProgramInputs(programNorm As String, contextSheet As Worksheet, contextTarget As Range)
    ' Fix-05: "Manual Working Copy" is always present and should refresh after any AMI edit,
    ' even if the current assignment violates rules (warn via tradeoffs; do not revert).
    Dim prevEnableEvents As Boolean
    Dim prevScreenUpdating As Boolean
    Dim prevSuppress As Boolean
    prevEnableEvents = Application.EnableEvents
    prevScreenUpdating = Application.ScreenUpdating
    prevSuppress = g_AMIOptixSuppressEvents

    On Error GoTo SafeExit

    If Not HasAPIKey() Then GoTo SafeExit
    If ActiveWorkbook Is Nothing Then GoTo SafeExit

    g_AMIOptixSuppressEvents = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Call ManualCalculateScenario(programNorm)

    ' ManualCalculateScenario activates AMI Scenarios at the end; restore user context.
    If Not contextSheet Is Nothing Then
        contextSheet.Activate
        If Not contextTarget Is Nothing Then
            contextTarget.Cells(1, 1).Select
        End If
    End If

SafeExit:
    On Error Resume Next
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    g_AMIOptixSuppressEvents = prevSuppress
End Sub

Private Sub ApplyManualAmiChangesAndRefresh(programNorm As String, changes As Object, contextSheet As Worksheet, contextTarget As Range)
    ' Applies unit_id -> AMI updates to the program sheet, then refreshes the manual working copy.
    Dim prevEnableEvents As Boolean
    Dim prevScreenUpdating As Boolean
    Dim prevSuppress As Boolean
    prevEnableEvents = Application.EnableEvents
    prevScreenUpdating = Application.ScreenUpdating
    prevSuppress = g_AMIOptixSuppressEvents

    On Error GoTo SafeExit

    If changes Is Nothing Then GoTo SafeExit
    If changes.Count = 0 Then GoTo SafeExit
    If Not HasAPIKey() Then GoTo SafeExit
    If ActiveWorkbook Is Nothing Then GoTo SafeExit

    g_AMIOptixSuppressEvents = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Dim dataWs As Worksheet
    Set dataWs = GetProgramDataSheet(programNorm)

    If Not dataWs Is Nothing Then
        Dim key As Variant
        For Each key In changes.Keys
            ApplyAmiToWorksheet dataWs, CStr(key), CDbl(changes(key))
        Next key
    End If

    Call ManualCalculateScenario(programNorm)

    ' ManualCalculateScenario activates AMI Scenarios at the end; restore user context.
    If Not contextSheet Is Nothing Then
        contextSheet.Activate
        If Not contextTarget Is Nothing Then
            contextTarget.Cells(1, 1).Select
        End If
    End If

SafeExit:
    On Error Resume Next
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    g_AMIOptixSuppressEvents = prevSuppress
End Sub

Private Sub HandleDataSheetAmiChange(ws As Worksheet, target As Range, programNorm As String)
    ' Sync: when the user edits the AMI column on the data sheet, update the Scenario Manual block.
    On Error GoTo SafeExit

    Call RefreshManualWorkingCopyFromProgramInputs(programNorm, ws, target)

    ' Keep last selection cache consistent for subsequent edits.
    m_LastSelectionSheet = CStr(ws.Name)
    m_LastSelectionAddress = target.Address(False, False)
    m_LastSelectionValue = target.Value

SafeExit:
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
    Dim headerRow As Long
    headerRow = FindManualTableHeaderRow(ws)
    If headerRow = 0 Then Exit Function

    Dim amiCol As Long
    amiCol = FindAMIColumn(ws, headerRow)
    If amiCol = 0 Then Exit Function

    Dim amiCells As Range
    Set amiCells = Intersect(target, ws.Columns(amiCol))
    If amiCells Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = FindManualTableLastRow(ws, headerRow)
    If lastRow = 0 Then Exit Function

    Dim tableAmi As Range
    Set tableAmi = ws.Range(ws.Cells(headerRow + 1, amiCol), ws.Cells(lastRow, amiCol))

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

    Dim programNorm As String
    programNorm = DetectProgramFromWorkbook()

    Dim changes As Object
    Set changes = CreateObject("Scripting.Dictionary") ' unit_id -> newAmi
    changes(unitId) = newAmi

    Call ApplyManualAmiChangesAndRefresh(programNorm, changes, ws, target)
    Exit Sub

Fail:
End Sub

Private Sub HandleManualScenarioEditRange(ws As Worksheet, target As Range)
    ' Handles multi-cell changes in the manual scenario AMI column (e.g., paste/fill-down).
    On Error GoTo Fail

    Dim headerRow As Long
    headerRow = FindManualTableHeaderRow(ws)
    If headerRow = 0 Then Exit Sub

    Dim amiCol As Long
    amiCol = FindAMIColumn(ws, headerRow)
    If amiCol = 0 Then Exit Sub

    Dim lastRow As Long
    lastRow = FindManualTableLastRow(ws, headerRow)
    If lastRow = 0 Then Exit Sub

    Dim amiCells As Range
    Set amiCells = Intersect(target, ws.Columns(amiCol))
    If amiCells Is Nothing Then Exit Sub

    Dim tableAmi As Range
    Set tableAmi = ws.Range(ws.Cells(headerRow + 1, amiCol), ws.Cells(lastRow, amiCol))

    Dim inTable As Range
    Set inTable = Intersect(amiCells, tableAmi)
    If inTable Is Nothing Then Exit Sub

    Dim changes As Object
    Set changes = CreateObject("Scripting.Dictionary") ' unit_id -> newAmi

    Dim c As Range
    For Each c In inTable.Cells
        Dim unitId As String
        unitId = Trim(CStr(ws.Cells(c.row, 1).Value))
        If unitId = "" Then GoTo NextCell

        Dim newAmi As Double
        newAmi = ParseAmiValue(c.Value)
        If newAmi <= 0 Then GoTo NextCell

        changes(unitId) = newAmi

NextCell:
    Next c

    If changes.Count = 0 Then Exit Sub

    Dim programNorm As String
    programNorm = DetectProgramFromWorkbook()

    Call ApplyManualAmiChangesAndRefresh(programNorm, changes, ws, target)
    Exit Sub

Fail:
    On Error Resume Next
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
        ' Prefer MIH sheet first (per client request), fallback only if needed.
        Set ws = ActiveWorkbook.Worksheets("MIH")
        If ws Is Nothing Then Set ws = ActiveWorkbook.Worksheets("RentRoll")
        If ws Is Nothing Then Set ws = ActiveWorkbook.Worksheets("UAP")
        If ws Is Nothing Then Set ws = ActiveWorkbook.Worksheets("PROJECT WORKSHEET")
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
        ' Only divide whole-percent entries like 60/120; MIH can legitimately be 1.2 (120%).
        If ParseAmiValue > 2# Then ParseAmiValue = ParseAmiValue / 100#
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
    ' Finds the header row for the Manual Working Copy unit table.
    ' We intentionally anchor on the manual-block label to avoid
    ' accidentally matching the scenario tables lower on the sheet.
    Dim r As Long
    Dim labelRow As Long
    labelRow = 0

    For r = 1 To 300
        ' Allow suffixes like "- CURRENT: ABSOLUTE BEST".
        Dim label As String
        label = UCase$(Trim$(CStr(ws.Cells(r, 1).Value)))
        If Left$(label, Len("SCENARIO MANUAL")) = "SCENARIO MANUAL" Or Left$(label, Len("MANUAL WORKING COPY")) = "MANUAL WORKING COPY" Then
            labelRow = r
            Exit For
        End If
    Next r

    If labelRow > 0 Then
        For r = labelRow + 1 To labelRow + 80
            If UCase$(Trim$(CStr(ws.Cells(r, 1).Value))) = "UNIT" Then
                If FindAMIColumn(ws, r) > 0 Then
                    FindManualTableHeaderRow = r
                    Exit Function
                End If
            End If
        Next r
    End If

    ' Fallback: search near the top of the sheet.
    For r = 1 To 200
        If UCase$(Trim$(CStr(ws.Cells(r, 1).Value))) = "UNIT" Then
            If FindAMIColumn(ws, r) > 0 Then
                FindManualTableHeaderRow = r
                Exit Function
            End If
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
        ' Stop when the scenario list begins below the manual block.
        If (v Like "SCENARIO [0-9]*") Or (v Like "SCENARIO #[0-9]*") Then
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
