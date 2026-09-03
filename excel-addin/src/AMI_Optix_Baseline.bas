Attribute VB_Name = "AMI_Optix_Baseline"
Option Explicit

' Baseline snapshot of the user's own AMI inputs ("YOUR ORIGINAL INPUT" truth).
'
' Problem this solves: ApplyBestScenario / ApplyScenarioByKey write solver
' output into the data sheet's AMI column. A later run then read that column
' as the user's "original input", so the Original Scenario showed the PREVIOUS
' RUN'S OUTPUT instead of what the user typed (Building D client report,
' 2026-09-02).
'
' How it works:
'   - A very-hidden sheet (BASELINE_SHEET) in the DATA workbook stores
'     unit_id -> AMI as last typed by the user, plus a signature of the last
'     PROGRAM write to the AMI column.
'   - Before each optimize run, EnsureBaselineAndTag compares the current AMI
'     column's signature to the last program-write signature:
'       equal     -> the user changed nothing since the program wrote; keep
'                    the stored baseline.
'       different -> the user hand-edited (or first run); the current column
'                    becomes the new baseline.
'     Each unit dict then gets unit("original_ami") from the baseline, which
'     BuildAPIPayloadV2 sends as units[].original_ami (server prefers it when
'     building the Original Scenario; absent field = legacy behavior).
'   - Every program write path calls RecordProgramWrite afterward.
'
' All entry points swallow their own errors: baseline failures must never
' block an optimize run or an apply.

Private Const BASELINE_SHEET As String = "AMI_Optix_Baseline"
Private Const MARKER As String = "BASELINE_V1"
Private Const FIRST_DATA_ROW As Long = 3

Public Sub EnsureBaselineAndTag(units As Collection)
    ' Called by Main after ReadUnitData, before the payload is built.
    On Error GoTo Done
    If units Is Nothing Then Exit Sub

    Dim sh As Worksheet
    Set sh = GetOrCreateBaselineSheet()
    If sh Is Nothing Then Exit Sub

    Dim curSig As String
    curSig = SignatureOfUnits(units)

    Dim lastProgSig As String
    lastProgSig = Trim$(CStr(sh.Range("B1").Value))

    Dim hasBaseline As Boolean
    hasBaseline = (Trim$(CStr(sh.Range("A1").Value)) = MARKER) And _
                  (Trim$(CStr(sh.Cells(FIRST_DATA_ROW, 1).Value)) <> "")

    If (Not hasBaseline) Or (curSig <> lastProgSig) Then
        ' First capture, or the user hand-edited since the program last wrote:
        ' the current column IS the user's input.
        WriteBaseline sh, units
    End If

    InjectOriginalAmi sh, units
Done:
End Sub

Public Sub RecordProgramWrite()
    ' Called by ResultsWriter immediately after any program write to the AMI
    ' column (apply / clear), so the next run can tell program writes apart
    ' from the user's own edits.
    On Error GoTo Done
    Dim units As Collection
    Set units = ReadUnitData()

    Dim sh As Worksheet
    Set sh = GetOrCreateBaselineSheet()
    If sh Is Nothing Then Exit Sub

    sh.Range("B1").Value = SignatureOfUnits(units)
Done:
End Sub

Private Function GetOrCreateBaselineSheet() As Worksheet
    On Error GoTo Fail
    Dim dataWs As Worksheet
    Set dataWs = GetDataSheet()
    If dataWs Is Nothing Then Exit Function

    Dim wb As Workbook
    Set wb = dataWs.Parent

    Dim sh As Worksheet
    On Error Resume Next
    Set sh = wb.Worksheets(BASELINE_SHEET)
    On Error GoTo Fail

    If sh Is Nothing Then
        Dim prevActive As Object
        Dim prevEvents As Boolean
        Dim prevUpdating As Boolean
        Set prevActive = ActiveSheet
        prevEvents = Application.EnableEvents
        prevUpdating = Application.ScreenUpdating
        Application.EnableEvents = False
        Application.ScreenUpdating = False

        Set sh = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        sh.Name = BASELINE_SHEET
        sh.Range("A1").Value = MARKER
        sh.Visible = xlSheetVeryHidden

        On Error Resume Next
        If Not prevActive Is Nothing Then prevActive.Activate
        On Error GoTo Fail
        Application.EnableEvents = prevEvents
        Application.ScreenUpdating = prevUpdating
    End If

    Set GetOrCreateBaselineSheet = sh
    Exit Function
Fail:
    Set GetOrCreateBaselineSheet = Nothing
End Function

Private Sub WriteBaseline(sh As Worksheet, units As Collection)
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Done

    sh.Range("A1").Value = MARKER
    sh.Range(sh.Cells(FIRST_DATA_ROW, 1), sh.Cells(sh.Rows.Count, 2)).ClearContents

    Dim r As Long
    r = FIRST_DATA_ROW
    Dim i As Long
    Dim unit As Object
    For i = 1 To units.Count
        Set unit = units(i)
        If unit.Exists("client_ami") Then
            sh.Cells(r, 1).Value = CStr(unit("unit_id"))
            sh.Cells(r, 2).Value = CDbl(unit("client_ami"))
            r = r + 1
        End If
    Next i
Done:
    Application.EnableEvents = prevEvents
End Sub

Private Sub InjectOriginalAmi(sh As Worksheet, units As Collection)
    On Error GoTo Done
    Dim lookup As Object
    Set lookup = CreateObject("Scripting.Dictionary")

    Dim r As Long
    r = FIRST_DATA_ROW
    Do While Trim$(CStr(sh.Cells(r, 1).Value)) <> ""
        lookup(CStr(sh.Cells(r, 1).Value)) = CDbl(sh.Cells(r, 2).Value)
        r = r + 1
        If r > 100000 Then Exit Do
    Loop

    Dim i As Long
    Dim unit As Object
    Dim id As String
    For i = 1 To units.Count
        Set unit = units(i)
        id = CStr(unit("unit_id"))
        If lookup.Exists(id) Then
            unit("original_ami") = CDbl(lookup(id))
        End If
    Next i
Done:
End Sub

Private Function SignatureOfUnits(units As Collection) As String
    ' Stable signature of (unit_id, AMI) pairs. Format$("0.####") absorbs
    ' float noise; the polynomial hash keeps the stored value short.
    On Error GoTo Fail
    Dim s As String
    Dim i As Long
    Dim unit As Object
    If Not units Is Nothing Then
        For i = 1 To units.Count
            Set unit = units(i)
            If unit.Exists("client_ami") Then
                s = s & "|" & CStr(unit("unit_id")) & "=" & Format$(CDbl(unit("client_ami")), "0.####")
            End If
        Next i
    End If

    Dim h As Double
    Dim p As Long
    h = 0#
    For p = 1 To Len(s)
        h = (h * 31# + CDbl(Asc(Mid$(s, p, 1)))) - Int((h * 31# + CDbl(Asc(Mid$(s, p, 1)))) / 999999937#) * 999999937#
    Next p
    SignatureOfUnits = CStr(h) & ":" & CStr(Len(s))
    Exit Function
Fail:
    SignatureOfUnits = "ERR"
End Function
