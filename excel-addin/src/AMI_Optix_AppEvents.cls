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

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo SafeExit

    If g_AMIOptixSuppressEvents Then Exit Sub
    If Target Is Nothing Then Exit Sub
    If Target.CountLarge <> 1 Then Exit Sub
    If TypeName(Sh) <> "Worksheet" Then Exit Sub

    Dim ws As Worksheet
    Set ws = Sh

    If UCase(ws.Name) = "UAP" Then
        If ChangeIsInAMIColumn(ws, Target) Then
            g_AMIOptixSuppressEvents = True
            UpdateManualScenario True
            g_AMIOptixSuppressEvents = False
        End If
        Exit Sub
    End If

    If UCase(ws.Name) = UCase("AMI Scenarios") Then
        If ChangeIsManualScenarioAMI(ws, Target) Then
            HandleManualScenarioEdit ws, Target
        End If
        Exit Sub
    End If

SafeExit:
End Sub

Private Function ChangeIsInAMIColumn(ws As Worksheet, target As Range) As Boolean
    Dim headerRow As Long
    Dim amiCol As Long
    headerRow = FindHeaderRow(ws)
    If headerRow = 0 Then Exit Function
    amiCol = FindAMIColumn(ws, headerRow)
    If amiCol = 0 Then Exit Function

    If target.Column = amiCol And target.row > headerRow Then
        ChangeIsInAMIColumn = True
    End If
End Function

Private Function ChangeIsManualScenarioAMI(ws As Worksheet, target As Range) As Boolean
    ' Manual block is in the top region (rows 1-120). AMI column is col 6.
    If target.row > 120 Then Exit Function
    If target.Column <> 6 Then Exit Function

    Dim headerRow As Long
    headerRow = FindManualTableHeaderRow(ws)
    If headerRow = 0 Then Exit Function

    If target.row <= headerRow Then Exit Function
    If Trim(CStr(ws.Cells(target.row, 1).Value)) = "" Then Exit Function

    ChangeIsManualScenarioAMI = True
End Function

Private Sub HandleManualScenarioEdit(ws As Worksheet, target As Range)
    On Error GoTo Fail

    Dim unitId As String
    unitId = Trim(CStr(ws.Cells(target.row, 1).Value))
    If unitId = "" Then Exit Sub

    Dim newAmi As Double
    newAmi = ParseAmiValue(target.Value)
    If newAmi <= 0 Then Exit Sub

    ' Validate via /api/evaluate before applying to UAP.
    Dim prevSheet As Worksheet
    Set prevSheet = ActiveSheet

    Dim uapWs As Worksheet
    Set uapWs = Nothing
    On Error Resume Next
    Set uapWs = ActiveWorkbook.Worksheets("UAP")
    On Error GoTo Fail
    If uapWs Is Nothing Then Exit Sub

    uapWs.Activate
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
    Set utilities = GetUtilitySelectionsForProgram("UAP")

    Dim payload As String
    payload = BuildEvaluatePayloadV2(units, utilities, "UAP", "", 0, 0)

    Dim response As String
    response = CallEvaluateAPI(payload)
    If response = "" Then Exit Sub

    Dim evalResult As Object
    Set evalResult = ParseJSON(response)
    If evalResult Is Nothing Then Exit Sub

    If evalResult.Exists("success") Then
        If evalResult("success") = False Then
            g_AMIOptixSuppressEvents = True
            Application.EnableEvents = False
            Application.Undo
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

    ' Apply to UAP (last-edit wins) and refresh manual scenario block.
    g_AMIOptixSuppressEvents = True
    ApplyAmiToUap unitId, newAmi
    UpdateManualScenario
    g_AMIOptixSuppressEvents = False
    Exit Sub

Fail:
    g_AMIOptixSuppressEvents = False
End Sub

Private Sub ApplyAmiToUap(unitId As String, newAmi As Double)
    On Error GoTo Fail

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("UAP")

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
    lastRow = ws.Cells(ws.Rows.Count, unitCol).End(xlUp).row

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
    Dim r As Long
    For r = 1 To 120
        If UCase(Trim(CStr(ws.Cells(r, 1).Value))) = "UNIT" And UCase(Trim(CStr(ws.Cells(r, 6).Value))) = "AMI" Then
            FindManualTableHeaderRow = r
            Exit Function
        End If
    Next r
    FindManualTableHeaderRow = 0
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
