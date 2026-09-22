Attribute VB_Name = "AMI_Optix_FortyRules"
Option Explicit

' AMI_OPTIX_FORTY_RULES_V1
'
' 40% Rules - ribbon "40% Rules" menu (owner request 2026-09-22, Fix 6).
'
' Optional. The owner decides WHICH apartments carry the 40% label; the
' optimizer still finds the best rent within those decisions. Four rules,
' each independent, all clearable:
'   - Pin selected units at 40%      (select unit rows on the sheet, click)
'   - Keep selected units OUT of 40% (select rows, click)
'   - Bedroom types allowed at 40%   (Studio / 1 / 2 / 3 / 4+; default all)
'   - Max 40% units per floor        (no limit / 1 / 2 / 3 / 4)
' "Clear all 40% rules" returns the workbook to "the program decides".
'
' Storage: hidden defined names in the DATA workbook (each building keeps
' its own rules):
'   AMI_Optix_Forty_Pins      "2A|3A"        unit ids, "|"-separated
'   AMI_Optix_Forty_Excludes  "5B"
'   AMI_Optix_Forty_Bedrooms  "1,2"          "" = all bedroom types
'   AMI_Optix_Forty_PerFloor  "2"            "" = no limit
' No stored rule = payload field not sent = legacy behavior.
'
' Server side (app.py) turns pins/exclusions/bedroom filter into per-unit
' band rules and the per-floor cap into a solver constraint, checks the
' rules against the 40% share window BEFORE solving, and echoes them in
' project_summary.forty_rules + the results header. Every entry point here
' swallows its own errors - a rules failure must never block a run.

Private Const PINS_NAME As String = "AMI_Optix_Forty_Pins"
Private Const EXCL_NAME As String = "AMI_Optix_Forty_Excludes"
Private Const BEDS_NAME As String = "AMI_Optix_Forty_Bedrooms"
Private Const FLOOR_NAME As String = "AMI_Optix_Forty_PerFloor"
Private Const SEP As String = "|"
Private Const ALL_BEDS As String = "0,1,2,3,4"
Private Const CUSTOMUI_NS As String = "http://schemas.microsoft.com/office/2009/07/customui"

'-------------------------------------------------------------------------------
' Storage
'-------------------------------------------------------------------------------

Private Function TargetWorkbook() As Workbook
    On Error Resume Next
    Set TargetWorkbook = ActiveWorkbook
End Function

Private Function ReadName(nameKey As String) As String
    On Error GoTo Fail
    Dim wb As Workbook
    Set wb = TargetWorkbook()
    If wb Is Nothing Then Exit Function
    Dim nm As Name
    Set nm = Nothing
    On Error Resume Next
    Set nm = wb.Names(nameKey)
    On Error GoTo Fail
    If nm Is Nothing Then Exit Function
    Dim raw As String
    raw = CStr(nm.RefersTo)
    If Left$(raw, 1) = "=" Then raw = Mid$(raw, 2)
    If Len(raw) >= 2 Then
        If Left$(raw, 1) = """" And Right$(raw, 1) = """" Then raw = Mid$(raw, 2, Len(raw) - 2)
    End If
    ReadName = Replace(raw, """""", """")
    Exit Function
Fail:
    ReadName = ""
End Function

Private Sub WriteName(nameKey As String, value As String)
    On Error GoTo Done
    Dim wb As Workbook
    Set wb = TargetWorkbook()
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    wb.Names(nameKey).Delete
    On Error GoTo Done
    If Trim$(value) <> "" Then
        wb.Names.Add Name:=nameKey, RefersTo:="=""" & Replace(value, """", """""") & """", Visible:=False
    End If
Done:
End Sub

Public Function GetPins() As String
    GetPins = ReadName(PINS_NAME)
End Function

Public Function GetExcludes() As String
    GetExcludes = ReadName(EXCL_NAME)
End Function

Public Function GetBedroomsCsv() As String
    ' "" = all bedroom types allowed at 40%.
    GetBedroomsCsv = ReadName(BEDS_NAME)
End Function

Public Function GetPerFloor() As Long
    ' 0 = no limit.
    Dim s As String
    s = ReadName(FLOOR_NAME)
    If IsNumeric(s) Then GetPerFloor = CLng(Val(s)) Else GetPerFloor = 0
    If GetPerFloor < 0 Then GetPerFloor = 0
End Function

Public Sub SetPerFloor(n As Long)
    If n <= 0 Then WriteName FLOOR_NAME, "" Else WriteName FLOOR_NAME, CStr(n)
End Sub

Public Function HasAnyRule() As Boolean
    HasAnyRule = (GetPins() <> "") Or (GetExcludes() <> "") Or (GetBedroomsCsv() <> "") Or (GetPerFloor() > 0)
End Function

Public Sub ClearAllRules()
    WriteName PINS_NAME, ""
    WriteName EXCL_NAME, ""
    WriteName BEDS_NAME, ""
    WriteName FLOOR_NAME, ""
End Sub

'-------------------------------------------------------------------------------
' List helpers ("|"-separated unit ids)
'-------------------------------------------------------------------------------

Private Function ListContains(lst As String, id As String) As Boolean
    Dim parts() As String
    Dim i As Long
    If Trim$(lst) = "" Then Exit Function
    parts = Split(lst, SEP)
    For i = LBound(parts) To UBound(parts)
        If StrComp(Trim$(parts(i)), Trim$(id), vbTextCompare) = 0 Then
            ListContains = True
            Exit Function
        End If
    Next i
End Function

Private Function ListAdd(lst As String, id As String) As String
    If ListContains(lst, id) Then
        ListAdd = lst
    ElseIf Trim$(lst) = "" Then
        ListAdd = Trim$(id)
    Else
        ListAdd = lst & SEP & Trim$(id)
    End If
End Function

Private Function ListRemove(lst As String, id As String) As String
    Dim parts() As String
    Dim i As Long
    Dim out As String
    If Trim$(lst) = "" Then Exit Function
    parts = Split(lst, SEP)
    For i = LBound(parts) To UBound(parts)
        If Trim$(parts(i)) <> "" Then
            If StrComp(Trim$(parts(i)), Trim$(id), vbTextCompare) <> 0 Then
                If out <> "" Then out = out & SEP
                out = out & Trim$(parts(i))
            End If
        End If
    Next i
    ListRemove = out
End Function

Private Function ListCount(lst As String) As Long
    If Trim$(lst) = "" Then Exit Function
    ListCount = UBound(Split(lst, SEP)) - LBound(Split(lst, SEP)) + 1
End Function

Private Function ListDisplay(lst As String) As String
    ListDisplay = Replace(lst, SEP, ", ")
End Function

'-------------------------------------------------------------------------------
' Selection -> unit ids
'-------------------------------------------------------------------------------

Public Function SelectedUnitIds(ByRef msg As String) As String
    ' Maps the user's current selection (rows on the program's unit sheet)
    ' to affordable unit ids. Returns "" with msg on any problem.
    On Error GoTo Fail
    msg = ""
    If TypeName(Selection) <> "Range" Then
        msg = "Select the unit rows on the MIH / UAP sheet first, then click again."
        Exit Function
    End If
    Dim sel As Range
    Set sel = Selection
    Dim selWs As Worksheet
    Set selWs = sel.Parent

    Dim units As Collection
    Set units = ReadUnitData()
    If units Is Nothing Then
        msg = "No unit data found in this workbook."
        Exit Function
    End If
    Dim dataWs As Worksheet
    Set dataWs = GetDataSheet()
    If dataWs Is Nothing Then
        msg = "No unit data sheet found in this workbook."
        Exit Function
    End If
    If StrComp(selWs.Name, dataWs.Name, vbTextCompare) <> 0 Then
        msg = "Select the unit rows on the '" & dataWs.Name & "' sheet (your selection is on '" & selWs.Name & "')."
        Exit Function
    End If

    Dim rows As Object
    Set rows = CreateObject("Scripting.Dictionary")
    Dim a As Range
    Dim r As Long
    For Each a In sel.Areas
        For r = a.Row To a.Row + a.Rows.Count - 1
            rows(CStr(r)) = True
        Next r
    Next a

    Dim out As String
    Dim i As Long
    Dim unit As Object
    For i = 1 To units.Count
        Set unit = units(i)
        If unit.Exists("row") Then
            If rows.Exists(CStr(unit("row"))) Then
                out = ListAdd(out, CStr(unit("unit_id")))
            End If
        End If
    Next i
    If out = "" Then
        msg = "None of the selected rows is an affordable unit (a unit needs a numeric AMI value to count)."
        Exit Function
    End If
    SelectedUnitIds = out
    Exit Function
Fail:
    msg = "Could not read the selection: " & Err.Description
    SelectedUnitIds = ""
End Function

Public Function PinSelected(ByRef msg As String) As Boolean
    Dim ids As String
    ids = SelectedUnitIds(msg)
    If ids = "" Then Exit Function
    Dim pins As String, excl As String
    pins = GetPins()
    excl = GetExcludes()
    Dim parts() As String
    Dim i As Long
    parts = Split(ids, SEP)
    For i = LBound(parts) To UBound(parts)
        pins = ListAdd(pins, parts(i))
        excl = ListRemove(excl, parts(i))
    Next i
    WriteName PINS_NAME, pins
    WriteName EXCL_NAME, excl
    msg = "Pinned at 40%: " & ListDisplay(ids) & vbCrLf & vbCrLf & "Now pinned: " & ListDisplay(pins)
    PinSelected = True
End Function

Public Function ExcludeSelected(ByRef msg As String) As Boolean
    Dim ids As String
    ids = SelectedUnitIds(msg)
    If ids = "" Then Exit Function
    Dim pins As String, excl As String
    pins = GetPins()
    excl = GetExcludes()
    Dim parts() As String
    Dim i As Long
    parts = Split(ids, SEP)
    For i = LBound(parts) To UBound(parts)
        excl = ListAdd(excl, parts(i))
        pins = ListRemove(pins, parts(i))
    Next i
    WriteName PINS_NAME, pins
    WriteName EXCL_NAME, excl
    msg = "Kept out of 40%: " & ListDisplay(ids) & vbCrLf & vbCrLf & "Now kept out: " & ListDisplay(excl)
    ExcludeSelected = True
End Function

Public Function UnruleSelected(ByRef msg As String) As Boolean
    Dim ids As String
    ids = SelectedUnitIds(msg)
    If ids = "" Then Exit Function
    Dim pins As String, excl As String
    pins = GetPins()
    excl = GetExcludes()
    Dim parts() As String
    Dim i As Long
    parts = Split(ids, SEP)
    For i = LBound(parts) To UBound(parts)
        pins = ListRemove(pins, parts(i))
        excl = ListRemove(excl, parts(i))
    Next i
    WriteName PINS_NAME, pins
    WriteName EXCL_NAME, excl
    msg = "Rule removed from: " & ListDisplay(ids) & vbCrLf & vbCrLf & "The program decides these units again."
    UnruleSelected = True
End Function

'-------------------------------------------------------------------------------
' Bedroom filter + per-floor cap
'-------------------------------------------------------------------------------

Private Function CsvContains(csv As String, v As Long) As Boolean
    Dim parts() As String
    Dim i As Long
    If Trim$(csv) = "" Then Exit Function
    parts = Split(csv, ",")
    For i = LBound(parts) To UBound(parts)
        If CLng(Val(Trim$(parts(i)))) = v And Trim$(parts(i)) <> "" Then
            CsvContains = True
            Exit Function
        End If
    Next i
End Function

Private Function NormalizeBedsCsv(csv As String) As String
    Dim allBeds() As String
    allBeds = Split(ALL_BEDS, ",")
    Dim i As Long
    Dim out As String
    For i = LBound(allBeds) To UBound(allBeds)
        If CsvContains(csv, CLng(allBeds(i))) Then
            If out <> "" Then out = out & ","
            out = out & allBeds(i)
        End If
    Next i
    NormalizeBedsCsv = out
End Function

Public Function IsBedroomAllowed(bed As Long) As Boolean
    Dim csv As String
    csv = GetBedroomsCsv()
    If csv = "" Then IsBedroomAllowed = True Else IsBedroomAllowed = CsvContains(csv, bed)
End Function

Public Function ToggleBedroom(bed As Long, pressed As Boolean) As String
    ' Returns "" on success, else the reason the change was refused.
    On Error GoTo Fail
    Dim csv As String
    csv = GetBedroomsCsv()
    If csv = "" Then csv = ALL_BEDS
    Dim nextCsv As String
    If pressed Then
        If CsvContains(csv, bed) Then nextCsv = csv Else nextCsv = csv & "," & CStr(bed)
    Else
        Dim parts() As String
        Dim i As Long
        parts = Split(csv, ",")
        nextCsv = ""
        For i = LBound(parts) To UBound(parts)
            If CLng(Val(Trim$(parts(i)))) <> bed Then
                If nextCsv <> "" Then nextCsv = nextCsv & ","
                nextCsv = nextCsv & Trim$(parts(i))
            End If
        Next i
    End If
    nextCsv = NormalizeBedsCsv(nextCsv)
    If nextCsv = "" Then
        ToggleBedroom = "At least one bedroom type must stay allowed at 40%."
        Exit Function
    End If
    If nextCsv = ALL_BEDS Then WriteName BEDS_NAME, "" Else WriteName BEDS_NAME, nextCsv
    ToggleBedroom = ""
    Exit Function
Fail:
    ToggleBedroom = "Could not save the bedroom rule: " & Err.Description
End Function

Public Function BedroomLabel(bed As Long) As String
    If bed <= 0 Then
        BedroomLabel = "Studio"
    ElseIf bed >= 4 Then
        BedroomLabel = "4+ BR"
    Else
        BedroomLabel = CStr(bed) & " BR"
    End If
End Function

Private Function BedsDisplay(csv As String) As String
    Dim parts() As String
    Dim i As Long
    Dim out As String
    If Trim$(csv) = "" Then Exit Function
    parts = Split(csv, ",")
    For i = LBound(parts) To UBound(parts)
        If out <> "" Then out = out & ", "
        out = out & BedroomLabel(CLng(Val(Trim$(parts(i)))))
    Next i
    BedsDisplay = out
End Function

'-------------------------------------------------------------------------------
' Describe / payload / preflight
'-------------------------------------------------------------------------------

Public Function DescribeRules() As String
    On Error GoTo Fail
    Dim parts As String
    If GetPins() <> "" Then parts = parts & "Pinned at 40%: " & ListDisplay(GetPins()) & vbCrLf
    If GetExcludes() <> "" Then parts = parts & "Kept out of 40%: " & ListDisplay(GetExcludes()) & vbCrLf
    If GetBedroomsCsv() <> "" Then parts = parts & "40% only for: " & BedsDisplay(GetBedroomsCsv()) & vbCrLf
    If GetPerFloor() > 0 Then parts = parts & "Max " & GetPerFloor() & " unit(s) at 40% per floor" & vbCrLf
    If parts = "" Then
        DescribeRules = "None - the program decides which apartments are 40%."
    Else
        DescribeRules = parts
    End If
    Exit Function
Fail:
    DescribeRules = "(could not read the 40% rules)"
End Function

Public Function DescribeRulesShort() As String
    On Error GoTo Fail
    Dim s As String
    If GetPins() <> "" Then s = s & ListCount(GetPins()) & " pinned"
    If GetExcludes() <> "" Then s = s & IIf(s <> "", ", ", "") & ListCount(GetExcludes()) & " kept out"
    If GetBedroomsCsv() <> "" Then s = s & IIf(s <> "", ", ", "") & BedsDisplay(GetBedroomsCsv()) & " only"
    If GetPerFloor() > 0 Then s = s & IIf(s <> "", ", ", "") & "max " & GetPerFloor() & "/floor"
    If s = "" Then s = "none (program decides)"
    DescribeRulesShort = s
    Exit Function
Fail:
    DescribeRulesShort = "none (program decides)"
End Function

Private Function JsonStr(s As String) As String
    Dim t As String
    t = Replace(s, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")
    t = Replace(t, vbTab, " ")
    JsonStr = """" & t & """"
End Function

Private Function ListToJsonArray(lst As String) As String
    Dim parts() As String
    Dim i As Long
    Dim out As String
    If Trim$(lst) = "" Then
        ListToJsonArray = "[]"
        Exit Function
    End If
    parts = Split(lst, SEP)
    For i = LBound(parts) To UBound(parts)
        If Trim$(parts(i)) <> "" Then
            If out <> "" Then out = out & ", "
            out = out & JsonStr(Trim$(parts(i)))
        End If
    Next i
    ListToJsonArray = "[" & out & "]"
End Function

Public Function FortyRulesJson() As String
    ' JSON object for the payload, or "" when no rule is set (field not sent).
    On Error GoTo Fail
    If Not HasAnyRule() Then Exit Function
    Dim j As String
    j = "{"
    j = j & """pin_units"": " & ListToJsonArray(GetPins())
    j = j & ", ""exclude_units"": " & ListToJsonArray(GetExcludes())
    If GetBedroomsCsv() <> "" Then
        j = j & ", ""bedrooms_allowed"": [" & Replace(GetBedroomsCsv(), ",", ", ") & "]"
    End If
    If GetPerFloor() > 0 Then
        j = j & ", ""max_per_floor"": " & CStr(GetPerFloor())
    End If
    j = j & "}"
    FortyRulesJson = j
    Exit Function
Fail:
    FortyRulesJson = ""
End Function

Public Function ValidateForRun(units As Collection, ByRef msg As String) As Boolean
    ' Run preflight: pinned / excluded ids must exist in THIS run's units.
    On Error GoTo Fail
    msg = ""
    ValidateForRun = True
    If Not HasAnyRule() Then Exit Function
    If units Is Nothing Then Exit Function

    Dim known As Object
    Set known = CreateObject("Scripting.Dictionary")
    known.CompareMode = vbTextCompare
    Dim i As Long
    For i = 1 To units.Count
        known(Trim$(CStr(units(i)("unit_id")))) = True
    Next i

    Dim missing As String
    Dim both As String
    Dim parts() As String
    Dim lst As String
    lst = GetPins()
    If lst <> "" Then
        parts = Split(lst, SEP)
        For i = LBound(parts) To UBound(parts)
            If Not known.Exists(Trim$(parts(i))) Then missing = ListAdd(missing, parts(i))
            If ListContains(GetExcludes(), parts(i)) Then both = ListAdd(both, parts(i))
        Next i
    End If
    lst = GetExcludes()
    If lst <> "" Then
        parts = Split(lst, SEP)
        For i = LBound(parts) To UBound(parts)
            If Not known.Exists(Trim$(parts(i))) Then missing = ListAdd(missing, parts(i))
        Next i
    End If

    If missing <> "" Then
        msg = "Your 40% rules mention unit(s) that are not in this run's affordable list: " & ListDisplay(missing) & "." & vbCrLf & vbCrLf & _
              "Open AMI Optix > 40% Rules and remove or re-pin them (or 'Clear all 40% rules'), then run again."
        ValidateForRun = False
        Exit Function
    End If
    If both <> "" Then
        msg = "Unit(s) " & ListDisplay(both) & " are both pinned at 40% and kept out of 40%." & vbCrLf & vbCrLf & _
              "Open AMI Optix > 40% Rules and fix the selection, then run again."
        ValidateForRun = False
        Exit Function
    End If
    Exit Function
Fail:
    msg = ""
    ValidateForRun = True
End Function

'-------------------------------------------------------------------------------
' Ribbon menu XML (dynamicMenu getContent)
'-------------------------------------------------------------------------------

Private Function EscapeXml(s As String) As String
    Dim t As String
    t = Replace(s, "&", "&amp;")
    t = Replace(t, "<", "&lt;")
    t = Replace(t, ">", "&gt;")
    t = Replace(t, """", "&quot;")
    EscapeXml = t
End Function

Public Function BuildFortyMenuXml() As String
    On Error GoTo Fail
    Dim x As String
    x = "<menu xmlns=""" & CUSTOMUI_NS & """ itemSize=""normal"">"
    x = x & "<button id=""btnFortyInfo"" enabled=""false"" imageMso=""Info"" label=""" & _
        EscapeXml("Current 40% rules: " & DescribeRulesShort()) & """/>"
    x = x & "<menuSeparator id=""sepForty1""/>"
    x = x & "<button id=""btnFortyPin"" label=""Pin selected units at 40%"" imageMso=""Lock"" onAction=""Ribbon_FortyPinSelected"" " & _
            "supertip=""Select the unit rows on the MIH / UAP sheet first. These apartments will always carry the 40% label.""/>"
    x = x & "<button id=""btnFortyExclude"" label=""Keep selected units OUT of 40%"" imageMso=""Cancel"" onAction=""Ribbon_FortyExcludeSelected"" " & _
            "supertip=""Select the unit rows first. These apartments will never carry the 40% label.""/>"
    x = x & "<button id=""btnFortyUnrule"" label=""Remove rule from selected units"" imageMso=""ClearFormatting"" onAction=""Ribbon_FortyUnruleSelected""/>"
    x = x & "<menuSeparator id=""sepForty2""/>"

    x = x & "<menu id=""mnuFortyBeds"" label=""Bedroom types allowed at 40%"" imageMso=""PropertySheet"">"
    Dim b As Long
    For b = 0 To 4
        x = x & "<checkBox id=""chkFortyBed" & b & """ tag=""" & b & """ label=""" & EscapeXml(BedroomLabel(b)) & """" & _
                " getPressed=""Ribbon_GetFortyBedroomPressed"" onAction=""Ribbon_ToggleFortyBedroom""/>"
    Next b
    x = x & "</menu>"

    x = x & "<menu id=""mnuFortyFloor"" label=""Max 40% units per floor"" imageMso=""TableRowsDistribute"">"
    Dim n As Long
    For n = 0 To 4
        x = x & "<checkBox id=""chkFortyFloor" & n & """ tag=""" & n & """ label=""" & IIf(n = 0, "No limit", CStr(n) & " per floor") & """" & _
                " getPressed=""Ribbon_GetFortyPerFloorPressed"" onAction=""Ribbon_SelectFortyPerFloor""/>"
    Next n
    x = x & "</menu>"

    x = x & "<menuSeparator id=""sepForty3""/>"
    x = x & "<button id=""btnFortyShow"" label=""Show current 40% rules"" imageMso=""ZoomPrintPreview"" onAction=""Ribbon_FortyShowRules""/>"
    x = x & "<button id=""btnFortyClear"" label=""Clear all 40% rules (program decides)"" imageMso=""Delete"" onAction=""Ribbon_FortyClear""/>"
    x = x & "</menu>"
    BuildFortyMenuXml = x
    Exit Function
Fail:
    BuildFortyMenuXml = BuildFortyMenuXmlFallback()
End Function

Public Function BuildFortyMenuXmlFallback() As String
    BuildFortyMenuXmlFallback = "<menu xmlns=""" & CUSTOMUI_NS & """>" & _
        "<button id=""btnFortyUnavailable"" enabled=""false"" label=""40% Rules unavailable (open a workbook)""/>" & _
        "<button id=""btnFortyClear"" label=""Clear all 40% rules (program decides)"" imageMso=""Delete"" onAction=""Ribbon_FortyClear""/>" & _
        "</menu>"
End Function
