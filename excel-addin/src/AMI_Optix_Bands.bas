Attribute VB_Name = "AMI_Optix_Bands"
Option Explicit

' AMI_OPTIX_BANDS_V1
'
' Band picker - ribbon "AMI Bands" menu (client feedback 2026-09-02, Fix 4).
'
' The owner decides which AMI bands the optimizer may use (e.g. "no 60%").
' The ribbon menu is generated here per program/option every time it opens
' (dynamicMenu + invalidateContentOnDrop), so it can never show stale items:
'   - UAP:          40 (required), 60, 70, 80, 90, 100
'   - MIH Option 1: 40 (required), 60, 70, 80, 90, 100      (client 100 cap)
'   - MIH Option 4: 40, 60, 70, 80, 90, 100, 110, 120, 130, 135 (no 40 rule,
'                   but at least one band <= 70 must stay on)
' A lower workbook cap (Prog!I4) shortens the list further.
'
' Storage: one hidden defined name in the DATA workbook,
'   AMI_Optix_AllowedBands = "<context>|<csv>"   e.g. "MIH Option 1|40,70,80"
' so each building remembers its own choice. No stored name = all bands
' (byte-identical legacy behavior; the payload field is simply not sent).
'
' Three locks: (1) this module refuses illegal toggles and the run preflight
' stops on a stale/invalid selection; (2) the server intersects the list with
' its own per-option candidates (can only narrow); (3) the solver's band
' domain is built from that list. Every entry point here swallows its own
' errors - a picker failure must never block a run.

Private Const BANDS_NAME As String = "AMI_Optix_AllowedBands"
Private Const ALL_BANDS As String = "40,60,70,80,90,100,110,120,130,135"
Private Const CUSTOMUI_NS As String = "http://schemas.microsoft.com/office/2009/07/customui"

' Floor-spread rule (Fix 5, ribbon toggle "Spread Across Floors"). Stored per
' workbook as AMI_Optix_FloorSpread = "1"/"0". Unset = ON for MIH workbooks
' (the HPD reviewer test that rejected Building D options 1-4), OFF for UAP.
Private Const FLOOR_SPREAD_NAME As String = "AMI_Optix_FloorSpread"

Public Function GetFloorSpreadEnabled() As Boolean
    On Error GoTo Fallback
    Dim wb As Workbook
    Set wb = TargetWorkbook()
    If wb Is Nothing Then GoTo Fallback

    Dim nm As Name
    Set nm = Nothing
    On Error Resume Next
    Set nm = wb.Names(FLOOR_SPREAD_NAME)
    On Error GoTo Fallback
    If nm Is Nothing Then GoTo Fallback

    Dim raw As String
    raw = CStr(nm.RefersTo)
    raw = Replace(raw, "=", "")
    raw = Replace(raw, """", "")
    GetFloorSpreadEnabled = (Trim$(raw) = "1")
    Exit Function
Fallback:
    On Error Resume Next
    GetFloorSpreadEnabled = (DetectProgramFromWorkbook() = "MIH")
End Function

Public Sub SetFloorSpreadEnabled(enabled As Boolean)
    On Error GoTo Done
    Dim wb As Workbook
    Set wb = TargetWorkbook()
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    wb.Names(FLOOR_SPREAD_NAME).Delete
    On Error GoTo Done
    wb.Names.Add Name:=FLOOR_SPREAD_NAME, RefersTo:="=""" & IIf(enabled, "1", "0") & """", Visible:=False
Done:
End Sub

'-------------------------------------------------------------------------------
' Context: what program/option is this workbook, and what bands does it allow?
'-------------------------------------------------------------------------------

Public Sub GetBandContext(ByRef programNorm As String, ByRef mihOption As String, _
                          ByRef capPct As Long, ByRef fortyRequired As Boolean)
    On Error GoTo Fallback
    programNorm = "UAP"
    mihOption = ""
    capPct = 100
    fortyRequired = True

    Dim kind As String
    Dim opt As String
    opt = ""
    kind = DetectWorkbookKind(opt)
    If kind = "MIH" Then
        programNorm = "MIH"
        mihOption = opt
        If opt = "Option 4" Then
            capPct = 135
            fortyRequired = False
        Else
            capPct = 100
            fortyRequired = True
        End If
        ' Honor a LOWER workbook cap (Prog!I4 factor, e.g. 0.8 = 80%), like the server does.
        Dim wbCap As Long
        wbCap = ReadWorkbookBandCap()
        If wbCap >= 60 And wbCap < capPct Then capPct = wbCap
    ElseIf kind = "MIH_INVALID" Then
        programNorm = "MIH"
        mihOption = opt
    End If
    Exit Sub
Fallback:
    programNorm = "UAP"
    mihOption = ""
    capPct = 100
    fortyRequired = True
End Sub

Private Function ReadWorkbookBandCap() As Long
    On Error GoTo Fail
    Dim v As Variant
    v = ActiveWorkbook.Worksheets("Prog").Range("I4").Value
    If IsNumeric(v) Then
        If CDbl(v) > 0 Then ReadWorkbookBandCap = CLng(CDbl(v) * 100)
    End If
    Exit Function
Fail:
    ReadWorkbookBandCap = 0
End Function

Public Function ProgramLabel(programNorm As String, mihOption As String) As String
    If UCase$(Trim$(programNorm)) = "MIH" Then
        If Trim$(mihOption) <> "" Then
            ProgramLabel = "MIH " & Trim$(mihOption)
        Else
            ProgramLabel = "MIH"
        End If
    Else
        ProgramLabel = "UAP"
    End If
End Function

Private Function CurrentContextLabel() As String
    Dim p As String, o As String, c As Long, f As Boolean
    GetBandContext p, o, c, f
    CurrentContextLabel = ProgramLabel(p, o)
End Function

'-------------------------------------------------------------------------------
' Storage (hidden defined name in the data workbook)
'-------------------------------------------------------------------------------

Private Function TargetWorkbook() As Workbook
    On Error Resume Next
    Set TargetWorkbook = ActiveWorkbook
End Function

Private Sub ReadStored(ByRef ctx As String, ByRef csv As String)
    ' ctx/csv are "" when nothing is stored.
    ctx = ""
    csv = ""
    On Error GoTo Fail
    Dim wb As Workbook
    Set wb = TargetWorkbook()
    If wb Is Nothing Then Exit Sub

    Dim nm As Name
    Set nm = Nothing
    On Error Resume Next
    Set nm = wb.Names(BANDS_NAME)
    On Error GoTo Fail
    If nm Is Nothing Then Exit Sub

    Dim raw As String
    raw = CStr(nm.RefersTo)              ' ="MIH Option 1|40,70,80"
    raw = Replace(raw, "=", "")
    raw = Replace(raw, """", "")
    raw = Trim$(raw)
    If raw = "" Then Exit Sub

    Dim bar As Long
    bar = InStr(1, raw, "|")
    If bar > 0 Then
        ctx = Trim$(Left$(raw, bar - 1))
        csv = NormalizeCsv(Mid$(raw, bar + 1))
    Else
        ctx = ""
        csv = NormalizeCsv(raw)
    End If
    If csv = "" Then ctx = ""
    Exit Sub
Fail:
    ctx = ""
    csv = ""
End Sub

Private Sub WriteStored(ctx As String, csv As String)
    On Error GoTo Done
    Dim wb As Workbook
    Set wb = TargetWorkbook()
    If wb Is Nothing Then Exit Sub

    Dim clean As String
    clean = NormalizeCsv(csv)

    On Error Resume Next
    wb.Names(BANDS_NAME).Delete
    On Error GoTo Done

    If clean <> "" Then
        wb.Names.Add Name:=BANDS_NAME, RefersTo:="=""" & ctx & "|" & clean & """", Visible:=False
    End If
Done:
End Sub

Public Sub ClearBandSelection()
    WriteStored "", ""
End Sub

Public Function HasBandSelection() As Boolean
    Dim ctx As String, csv As String
    ReadStored ctx, csv
    HasBandSelection = (csv <> "")
End Function

'-------------------------------------------------------------------------------
' CSV helpers
'-------------------------------------------------------------------------------

Private Function NormalizeCsv(csv As String) As String
    ' Keeps only known bands, de-duplicates, sorts ascending.
    Dim allBands() As String
    allBands = Split(ALL_BANDS, ",")
    Dim out As String
    Dim i As Long
    For i = LBound(allBands) To UBound(allBands)
        If CsvContains(csv, CLng(allBands(i))) Then
            If out <> "" Then out = out & ","
            out = out & allBands(i)
        End If
    Next i
    NormalizeCsv = out
End Function

Private Function CsvContains(csv As String, band As Long) As Boolean
    Dim parts() As String
    Dim i As Long
    If Trim$(csv) = "" Then Exit Function
    parts = Split(csv, ",")
    For i = LBound(parts) To UBound(parts)
        If CLng(Val(Trim$(parts(i)))) = band Then
            CsvContains = True
            Exit Function
        End If
    Next i
End Function

Private Function CsvAdd(csv As String, band As Long) As String
    If CsvContains(csv, band) Then
        CsvAdd = NormalizeCsv(csv)
    ElseIf Trim$(csv) = "" Then
        CsvAdd = CStr(band)
    Else
        CsvAdd = NormalizeCsv(csv & "," & CStr(band))
    End If
End Function

Private Function CsvRemove(csv As String, band As Long) As String
    Dim parts() As String
    Dim i As Long
    Dim out As String
    If Trim$(csv) = "" Then Exit Function
    parts = Split(csv, ",")
    For i = LBound(parts) To UBound(parts)
        If CLng(Val(Trim$(parts(i)))) <> band And Trim$(parts(i)) <> "" Then
            If out <> "" Then out = out & ","
            out = out & Trim$(parts(i))
        End If
    Next i
    CsvRemove = NormalizeCsv(out)
End Function

Private Function CsvWithinCap(csv As String, capPct As Long) As String
    Dim parts() As String
    Dim i As Long
    Dim out As String
    If Trim$(csv) = "" Then Exit Function
    parts = Split(csv, ",")
    For i = LBound(parts) To UBound(parts)
        If CLng(Val(Trim$(parts(i)))) > 0 And CLng(Val(Trim$(parts(i)))) <= capPct Then
            If out <> "" Then out = out & ","
            out = out & Trim$(parts(i))
        End If
    Next i
    CsvWithinCap = NormalizeCsv(out)
End Function

Private Function CsvCount(csv As String) As Long
    If Trim$(csv) = "" Then Exit Function
    CsvCount = UBound(Split(csv, ",")) - LBound(Split(csv, ",")) + 1
End Function

Private Function CsvMin(csv As String) As Long
    Dim parts() As String
    Dim i As Long
    Dim v As Long
    CsvMin = 0
    If Trim$(csv) = "" Then Exit Function
    parts = Split(csv, ",")
    For i = LBound(parts) To UBound(parts)
        v = CLng(Val(Trim$(parts(i))))
        If v > 0 Then
            If CsvMin = 0 Or v < CsvMin Then CsvMin = v
        End If
    Next i
End Function

Public Function MenuBandsCsv(capPct As Long) As String
    MenuBandsCsv = CsvWithinCap(ALL_BANDS, capPct)
End Function

Private Function CsvToDisplay(csv As String) As String
    Dim parts() As String
    Dim i As Long
    Dim out As String
    If Trim$(csv) = "" Then Exit Function
    parts = Split(csv, ",")
    For i = LBound(parts) To UBound(parts)
        If out <> "" Then out = out & ", "
        out = out & Trim$(parts(i)) & "%"
    Next i
    CsvToDisplay = out
End Function

'-------------------------------------------------------------------------------
' Rules (lock 1)
'-------------------------------------------------------------------------------

Private Function ValidateCsv(csv As String, programNorm As String, mihOption As String, _
                             capPct As Long, fortyRequired As Boolean) As String
    ' Returns "" when the selection can produce a legal scenario, else the reason.
    Dim eff As String
    eff = CsvWithinCap(csv, capPct)
    Dim label As String
    label = ProgramLabel(programNorm, mihOption)

    If fortyRequired And Not CsvContains(eff, 40) Then
        ValidateCsv = "40% AMI is required for " & label & " and cannot be turned off."
        Exit Function
    End If
    If UCase$(programNorm) = "MIH" And mihOption = "Option 4" Then
        If CsvMin(eff) = 0 Or CsvMin(eff) > 70 Then
            ValidateCsv = "MIH Option 4 needs at least one band at or below 70% AMI (the 5% set-aside). Keep 40%, 60% or 70% checked."
            Exit Function
        End If
    End If
    If CsvCount(eff) < 2 Then
        ValidateCsv = "At least 2 bands must stay checked for " & label & " (every scenario uses 2 or more bands)."
        Exit Function
    End If
    ValidateCsv = ""
End Function

Public Function IsBandAllowed(band As Long) As Boolean
    ' Used by the menu's getPressed. A selection stored for a DIFFERENT
    ' program/option is treated as "all bands" (the menu says so).
    On Error GoTo Fail
    Dim ctx As String, csv As String
    ReadStored ctx, csv
    If csv = "" Then IsBandAllowed = True: Exit Function
    If ctx <> "" And ctx <> CurrentContextLabel() Then IsBandAllowed = True: Exit Function
    IsBandAllowed = CsvContains(csv, band)
    Exit Function
Fail:
    IsBandAllowed = True
End Function

Public Function ToggleBand(band As Long, pressed As Boolean) As String
    ' Applies a checkbox change. Returns "" on success, otherwise the reason
    ' the change was refused (the checkbox re-renders in its old state).
    On Error GoTo Fail
    Dim programNorm As String, mihOption As String, capPct As Long, fortyReq As Boolean
    GetBandContext programNorm, mihOption, capPct, fortyReq
    Dim curCtx As String
    curCtx = ProgramLabel(programNorm, mihOption)

    Dim ctx As String, csv As String
    ReadStored ctx, csv
    ' No selection yet, or one made for another option: start from "all bands for this option".
    If csv = "" Or (ctx <> "" And ctx <> curCtx) Then csv = MenuBandsCsv(capPct)
    csv = CsvWithinCap(csv, capPct)

    Dim nextCsv As String
    If pressed Then
        nextCsv = CsvAdd(csv, band)
    Else
        If band = 40 And fortyReq Then
            ToggleBand = "40% AMI is required for " & curCtx & " and cannot be turned off."
            Exit Function
        End If
        nextCsv = CsvRemove(csv, band)
    End If

    Dim why As String
    why = ValidateCsv(nextCsv, programNorm, mihOption, capPct, fortyReq)
    If why <> "" Then
        ToggleBand = why
        Exit Function
    End If

    If nextCsv = MenuBandsCsv(capPct) Then
        ' Everything checked = program default. Store nothing so the payload
        ' field is not sent and the server runs exactly as before the picker.
        WriteStored "", ""
    Else
        WriteStored curCtx, nextCsv
    End If
    ToggleBand = ""
    Exit Function
Fail:
    ToggleBand = "Could not save the band selection: " & Err.Description
End Function

Public Function ValidateSelectionForRun(programNorm As String, mihOption As String, ByRef msg As String) As Boolean
    ' Run preflight. True = go. False = stop, with msg for the user.
    ' Re-reads program + option on EVERY run so a stale selection (made for
    ' another option) can never silently narrow or widen a run.
    On Error GoTo Fail
    msg = ""
    ValidateSelectionForRun = True

    Dim ctx As String, csv As String
    ReadStored ctx, csv
    If csv = "" Then Exit Function

    Dim p As String, o As String, capPct As Long, fortyReq As Boolean
    GetBandContext p, o, capPct, fortyReq
    Dim curCtx As String
    curCtx = ProgramLabel(UCase$(Trim$(programNorm)), mihOption)

    If ctx <> "" And ctx <> curCtx Then
        msg = "Your AMI Bands selection (" & CsvToDisplay(csv) & ") was made for " & ctx & "." & vbCrLf & _
              "This run is " & curCtx & "." & vbCrLf & vbCrLf & _
              "Open AMI Optix > AMI Bands, check the bands you want for " & curCtx & _
              " (or click 'Allow all bands'), then run again."
        ValidateSelectionForRun = False
        Exit Function
    End If

    Dim why As String
    why = ValidateCsv(csv, UCase$(Trim$(programNorm)), mihOption, capPct, fortyReq)
    If why <> "" Then
        msg = "Your AMI Bands selection (" & CsvToDisplay(csv) & ") cannot be used: " & why & vbCrLf & vbCrLf & _
              "Open AMI Optix > AMI Bands and adjust it (or click 'Allow all bands'), then run again."
        ValidateSelectionForRun = False
        Exit Function
    End If
    Exit Function
Fail:
    ' Never block a run because the preflight itself failed.
    msg = ""
    ValidateSelectionForRun = True
End Function

'-------------------------------------------------------------------------------
' Payload + display
'-------------------------------------------------------------------------------

Public Function AllowedBandsJsonArray() As String
    ' "[40, 70, 80]" when the user narrowed the bands for the current
    ' program/option; "" otherwise (field not sent = legacy behavior).
    On Error GoTo Fail
    Dim ctx As String, csv As String
    ReadStored ctx, csv
    If csv = "" Then Exit Function
    Dim p As String, o As String, capPct As Long, fortyReq As Boolean
    GetBandContext p, o, capPct, fortyReq
    If ctx <> "" And ctx <> ProgramLabel(p, o) Then Exit Function
    Dim eff As String
    eff = CsvWithinCap(csv, capPct)
    If eff = "" Then Exit Function
    AllowedBandsJsonArray = "[" & Replace(eff, ",", ", ") & "]"
    Exit Function
Fail:
    AllowedBandsJsonArray = ""
End Function

Public Function DescribeSelection() As String
    On Error GoTo Fail
    Dim ctx As String, csv As String
    ReadStored ctx, csv
    If csv = "" Then
        DescribeSelection = "All bands"
    Else
        DescribeSelection = CsvToDisplay(csv)
    End If
    Exit Function
Fail:
    DescribeSelection = "All bands"
End Function

Private Function EscapeXml(s As String) As String
    Dim t As String
    t = Replace(s, "&", "&amp;")
    t = Replace(t, "<", "&lt;")
    t = Replace(t, ">", "&gt;")
    t = Replace(t, """", "&quot;")
    EscapeXml = t
End Function

Public Function BuildBandsMenuXml() As String
    ' getContent for the ribbon dynamicMenu. Rebuilt on every drop.
    On Error GoTo Fail
    Dim programNorm As String, mihOption As String, capPct As Long, fortyReq As Boolean
    GetBandContext programNorm, mihOption, capPct, fortyReq
    Dim curCtx As String
    curCtx = ProgramLabel(programNorm, mihOption)

    Dim ctx As String, csv As String
    ReadStored ctx, csv

    Dim x As String
    x = "<menu xmlns=""" & CUSTOMUI_NS & """ itemSize=""normal"">"
    x = x & "<button id=""btnBandsInfo"" enabled=""false"" imageMso=""Info"" label=""" & _
        EscapeXml(curCtx & " - bands up to " & CStr(capPct) & "%") & """/>"
    If csv <> "" And ctx <> "" And ctx <> curCtx Then
        x = x & "<button id=""btnBandsStale"" enabled=""false"" imageMso=""AlertsView"" label=""" & _
            EscapeXml("Previous selection was for " & ctx & " - showing all bands") & """/>"
    ElseIf csv <> "" Then
        x = x & "<button id=""btnBandsCurrent"" enabled=""false"" imageMso=""FilterBySelection"" label=""" & _
            EscapeXml("Using: " & CsvToDisplay(csv)) & """/>"
    Else
        x = x & "<button id=""btnBandsCurrent"" enabled=""false"" imageMso=""Filter"" label=""Using: all bands""/>"
    End If
    x = x & "<menuSeparator id=""sepBands1""/>"

    Dim allBands() As String
    allBands = Split(ALL_BANDS, ",")
    Dim i As Long
    Dim b As Long
    Dim lbl As String
    For i = LBound(allBands) To UBound(allBands)
        b = CLng(allBands(i))
        If b <= capPct Then
            lbl = CStr(b) & "% AMI"
            If b = 40 And fortyReq Then lbl = "40% AMI (required)"
            x = x & "<checkBox id=""chkBand" & CStr(b) & """ tag=""" & CStr(b) & """ label=""" & EscapeXml(lbl) & """" & _
                    " getPressed=""Ribbon_GetBandPressed"" onAction=""Ribbon_ToggleBand"""
            If b = 40 And fortyReq Then x = x & " enabled=""false"""
            x = x & "/>"
        End If
    Next i

    x = x & "<menuSeparator id=""sepBands2""/>"
    x = x & "<button id=""btnBandsAllowAll"" label=""Allow all bands"" imageMso=""Refresh"" onAction=""Ribbon_BandsAllowAll""/>"
    x = x & "</menu>"
    BuildBandsMenuXml = x
    Exit Function
Fail:
    BuildBandsMenuXml = BuildBandsMenuXmlFallback()
End Function

Public Function BuildBandsMenuXmlFallback() As String
    ' Minimal valid menu so the ribbon never shows a broken control.
    BuildBandsMenuXmlFallback = "<menu xmlns=""" & CUSTOMUI_NS & """>" & _
        "<button id=""btnBandsUnavailable"" enabled=""false"" label=""AMI Bands unavailable (open a workbook)""/>" & _
        "<button id=""btnBandsAllowAll"" label=""Allow all bands"" imageMso=""Refresh"" onAction=""Ribbon_BandsAllowAll""/>" & _
        "</menu>"
End Function
