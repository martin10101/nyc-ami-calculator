Attribute VB_Name = "AMI_Optix_VerifyManualRents"
'===============================================================================
' AMI OPTIX - Verify Manual Rents (API) (Fix-06d)
'
' Manual Working Copy rents are computed locally via Fix-06c cache tables.
' This feature provides an on-demand, one-click verification that local manual
' rents match API evaluation using /api/evaluate (NOT /api/manual_calculate).
'===============================================================================
Option Explicit

Public Const AMI_OPTIX_VERIFY_TOLERANCE_UNIT_DOLLARS As Double = 1#
Public Const AMI_OPTIX_VERIFY_TOLERANCE_TOTAL_DOLLARS As Double = 1#

' Last verify result for Diagnostics rendering (non-persistent; current Excel session).
' Keys:
'  - timestamp, result (MATCH/MISMATCH)
'  - program, rent_roll_year
'  - local_cache_* (status/label/source_path/generated_at)
'  - api_* (rent_roll_year_used/calculator_filename/rent_schedule_source/rent_schedule_warning)
'  - compare_* (per_unit/totals)
'  - totals (local_net_monthly/api_net_monthly/delta)
'  - mismatches (Collection of Dictionaries)
'  - errors (string)
'  - workbook_name, workbook_path
Public g_AMIOptixLastVerify As Object

Public Sub VerifyManualRentsAPI()
    VerifyManualRentsAPICore False
End Sub

Public Sub VerifyManualRentsAPI_Agent()
    VerifyManualRentsAPICore True
End Sub

Private Sub VerifyManualRentsAPICore(Optional ByVal automationMode As Boolean = False)
    ' One-click verification: compute local rents from cache, then compare with /api/evaluate.
    On Error GoTo Fail

    Dim startedAt As String
    startedAt = Format$(Now, "yyyy-mm-dd hh:nn:ss")

    Dim wbName As String
    Dim wbPath As String
    wbName = ""
    wbPath = ""
    If Not ActiveWorkbook Is Nothing Then
        wbName = ActiveWorkbook.Name
        wbPath = ActiveWorkbook.FullName
    End If

    Dim programNorm As String
    programNorm = UCase$(Trim$(DetectProgramFromWorkbook()))
    If programNorm <> "UAP" And programNorm <> "MIH" Then programNorm = "UAP"

    Dim selectedYear As Long
    selectedYear = GetSelectedRentRollYearSettingLocal()

    ' Ensure local cache exists/valid before verification.
    ' This gives a clear blocking message with expected source paths if the year workbook is missing.
    Dim ensuredSourcePath As String
    Dim ensuredCacheFolder As String
    Dim ensuredSourceFp As String
    ensuredSourcePath = ""
    ensuredCacheFolder = ""
    ensuredSourceFp = ""
    Call EnsureRentTablesCache(selectedYear, False, ensuredSourcePath, ensuredCacheFolder, ensuredSourceFp)

    Dim rentStatus As Object
    Set rentStatus = Nothing
    On Error Resume Next
    Set rentStatus = GetRentTablesStatus(selectedYear)
    On Error GoTo Fail

    Dim localCacheStatus As String
    Dim localCacheLabel As String
    Dim localCacheSourcePath As String
    Dim localCacheGeneratedAt As String
    localCacheStatus = ""
    localCacheLabel = ""
    localCacheSourcePath = ""
    localCacheGeneratedAt = ""

    If Not rentStatus Is Nothing Then
        On Error Resume Next
        localCacheStatus = CStr(rentStatus("cache_status"))
        localCacheLabel = CStr(rentStatus("cache_source_label"))
        localCacheSourcePath = CStr(rentStatus("cache_source_path"))
        localCacheGeneratedAt = CStr(rentStatus("cache_generated_at"))
        On Error GoTo Fail
    End If

    Dim units As Collection
    Set units = ReadProgramUnits(programNorm)
    If units Is Nothing Or units.Count = 0 Then
        Err.Raise vbObjectError + 650, "AMI_Optix_VerifyManualRents.VerifyManualRentsAPI", _
                  "No units found for verification." & vbCrLf & vbCrLf & _
                  "Make sure the program sheet has a populated AMI column with numeric values."
    End If

    Dim utilities As Object
    Set utilities = GetUtilitySelectionsForProgram(programNorm)

    ' Local rent calc (Fix-06c) - DO NOT rebuild cache here.
    Call LoadRentLimitsCacheToDict(selectedYear)
    Call LoadUtilityAllowancesCacheToDict(selectedYear)

    Dim localNetByUnitId As Object
    Set localNetByUnitId = CreateObject("Scripting.Dictionary")
    localNetByUnitId.CompareMode = vbTextCompare

    Dim localTotalNet As Double
    localTotalNet = 0#

    Dim i As Long
    For i = 1 To units.Count
        Dim u As Object
        Set u = units(i)
        If u Is Nothing Then GoTo NextUnit

        Dim unitId As String
        unitId = ""
        On Error Resume Next
        If u.Exists("unit_id") Then unitId = CStr(u("unit_id"))
        On Error GoTo Fail
        If Trim$(unitId) = "" Then GoTo NextUnit

        Dim ami As Double
        ami = 0#
        If u.Exists("client_ami") Then
            If IsNumeric(u("client_ami")) Then ami = CDbl(u("client_ami"))
        End If
        If ami > 2# Then ami = ami / 100#
        If ami <= 0# Then GoTo NextUnit

        Dim rentResult As Object
        Set rentResult = ComputeNetRent(selectedYear, programNorm, u("bedrooms"), ami, utilities, unitId)

        Dim net As Double
        net = 0#
        If Not rentResult Is Nothing Then
            net = CDbl(rentResult("monthly_rent"))
        End If

        localNetByUnitId(unitId) = net
        localTotalNet = localTotalNet + net

NextUnit:
    Next i

    Dim mihOption As String
    Dim mihResidentialSF As Double
    Dim mihMaxBandPercent As Long
    mihOption = ""
    mihResidentialSF = 0#
    mihMaxBandPercent = 0

    If programNorm = "MIH" Then
        Dim mihErr As String
        mihErr = ""
        If Not TryReadMIHInputsQuiet(mihOption, mihResidentialSF, mihMaxBandPercent, mihErr) Then
            RecordVerifyState startedAt, "MISMATCH", programNorm, selectedYear, _
                              localCacheStatus, localCacheLabel, localCacheSourcePath, localCacheGeneratedAt, _
                              "", "", "", "", _
                              localTotalNet, Empty, False, False, Nothing, _
                              mihErr, wbName, wbPath
            ShowVerifySummary automationMode, "MISMATCH", _
                              BuildVerifySummaryMessage("MISMATCH", programNorm, selectedYear, localCacheLabel, localCacheSourcePath, localCacheGeneratedAt, _
                                                        "", "", "", "", localTotalNet, Empty, False, False, Nothing, _
                                                        "MIH inputs are missing/invalid:" & vbCrLf & mihErr)
            Exit Sub
        End If
    End If

    Dim payload As String
    payload = BuildEvaluatePayloadV2(units, utilities, programNorm, mihOption, mihResidentialSF, mihMaxBandPercent)

    Dim apiErr As String
    apiErr = ""
    Dim responseText As String
    responseText = CallEvaluateAPIStateless(payload, apiErr)
    If responseText = "" Then
        Err.Raise vbObjectError + 651, "AMI_Optix_VerifyManualRents.VerifyManualRentsAPI", _
                  "API evaluate failed." & vbCrLf & vbCrLf & apiErr
    End If

    Dim apiResult As Object
    Set apiResult = ParseJSON(responseText)
    If apiResult Is Nothing Then
        Err.Raise vbObjectError + 652, "AMI_Optix_VerifyManualRents.VerifyManualRentsAPI", "Could not parse /api/evaluate JSON response."
    End If

    Dim apiSuccess As Boolean
    apiSuccess = True
    On Error Resume Next
    If apiResult.Exists("success") Then apiSuccess = CBool(apiResult("success"))
    On Error GoTo Fail

    Dim apiYearUsed As String
    Dim apiCalcFilename As String
    Dim apiSource As String
    Dim apiWarn As String
    apiYearUsed = DictGetString(apiResult, "rent_roll_year_used", "")
    apiCalcFilename = DictGetString(apiResult, "calculator_filename", "")
    apiSource = DictGetString(apiResult, "rent_schedule_source", "")
    apiWarn = DictGetString(apiResult, "rent_schedule_warning", "")

    If Not apiSuccess Then
        Dim errs As String
        errs = ExtractErrorsList(apiResult)

        RecordVerifyState startedAt, "MISMATCH", programNorm, selectedYear, _
                          localCacheStatus, localCacheLabel, localCacheSourcePath, localCacheGeneratedAt, _
                          apiYearUsed, apiCalcFilename, apiSource, apiWarn, _
                          localTotalNet, Empty, False, False, Nothing, _
                          "API evaluate reported assignment invalid." & vbCrLf & vbCrLf & errs, wbName, wbPath

        ShowVerifySummary automationMode, "MISMATCH", _
                          BuildVerifySummaryMessage("MISMATCH", programNorm, selectedYear, localCacheLabel, localCacheSourcePath, localCacheGeneratedAt, _
                                                    apiYearUsed, apiCalcFilename, apiSource, apiWarn, _
                                                    localTotalNet, Empty, False, False, Nothing, _
                                                    "API evaluate reported assignment invalid:" & vbCrLf & vbCrLf & errs)
        Exit Sub
    End If

    Dim apiNetByUnitId As Object
    Set apiNetByUnitId = CreateObject("Scripting.Dictionary")
    apiNetByUnitId.CompareMode = vbTextCompare

    Dim perUnitAvailable As Boolean
    perUnitAvailable = False

    Dim apiAssignments As Object
    Set apiAssignments = Nothing
    On Error Resume Next
    Set apiAssignments = apiResult("assignments")
    On Error GoTo Fail

    If Not apiAssignments Is Nothing Then
        Dim a As Variant
        For Each a In apiAssignments
            If Not a Is Nothing Then
                Dim aid As String
                aid = ""
                On Error Resume Next
                If a.Exists("unit_id") Then aid = CStr(a("unit_id"))
                On Error GoTo Fail
                If Trim$(aid) = "" Then GoTo NextA

                Dim netVal As Variant
                netVal = Empty
                On Error Resume Next
                If a.Exists("monthly_rent") Then netVal = a("monthly_rent")
                On Error GoTo Fail

                If IsNumeric(netVal) Then
                    apiNetByUnitId(aid) = CDbl(netVal)
                    perUnitAvailable = True
                End If
            End If
NextA:
        Next a
    End If

    Dim apiTotalNet As Variant
    apiTotalNet = Empty

    Dim rentTotals As Object
    Set rentTotals = Nothing
    On Error Resume Next
    Set rentTotals = apiResult("rent_totals")
    On Error GoTo Fail

    If Not rentTotals Is Nothing Then
        Dim tval As Variant
        tval = Empty
        On Error Resume Next
        If rentTotals.Exists("net_monthly") Then tval = rentTotals("net_monthly")
        On Error GoTo Fail
        If IsNumeric(tval) Then apiTotalNet = CDbl(tval)
    End If

    If IsEmpty(apiTotalNet) And perUnitAvailable Then
        apiTotalNet = SumDictValues(apiNetByUnitId)
    End If

    Dim comparedPerUnit As Boolean
    comparedPerUnit = perUnitAvailable

    Dim comparedTotals As Boolean
    comparedTotals = Not IsEmpty(apiTotalNet)

    Dim mismatches As Collection
    Set mismatches = New Collection

    If comparedPerUnit Then
        Dim key As Variant
        For Each key In localNetByUnitId.Keys
            Dim lid As String
            lid = CStr(key)

            Dim lnet As Double
            lnet = CDbl(localNetByUnitId(lid))

            If Not apiNetByUnitId.Exists(lid) Then
                mismatches.Add MakeMismatch(lid, lnet, Empty, Empty, "missing_in_api")
            Else
                Dim anet As Double
                anet = CDbl(apiNetByUnitId(lid))
                Dim delta As Double
                delta = lnet - anet
                If Abs(delta) > AMI_OPTIX_VERIFY_TOLERANCE_UNIT_DOLLARS + 0.0001 Then
                    mismatches.Add MakeMismatch(lid, lnet, anet, delta, "delta_exceeds_tolerance")
                End If
            End If
        Next key

        For Each key In apiNetByUnitId.Keys
            Dim rid As String
            rid = CStr(key)
            If Not localNetByUnitId.Exists(rid) Then
                mismatches.Add MakeMismatch(rid, Empty, CDbl(apiNetByUnitId(rid)), Empty, "missing_in_local")
            End If
        Next key
    End If

    Dim totalsMismatch As Boolean
    Dim totalDelta As Double
    totalsMismatch = False
    totalDelta = 0#
    If comparedTotals Then
        totalDelta = localTotalNet - CDbl(apiTotalNet)
        If Abs(totalDelta) > AMI_OPTIX_VERIFY_TOLERANCE_TOTAL_DOLLARS + 0.0001 Then totalsMismatch = True
    End If

    Dim isMatch As Boolean
    If comparedPerUnit Then
        isMatch = (mismatches.Count = 0) And (Not totalsMismatch)
    ElseIf comparedTotals Then
        isMatch = (Not totalsMismatch)
    Else
        isMatch = False
    End If

    Dim result As String
    If isMatch Then
        result = "MATCH"
    Else
        result = "MISMATCH"
    End If

    RecordVerifyState startedAt, result, programNorm, selectedYear, _
                      localCacheStatus, localCacheLabel, localCacheSourcePath, localCacheGeneratedAt, _
                      apiYearUsed, apiCalcFilename, apiSource, apiWarn, _
                      localTotalNet, apiTotalNet, comparedPerUnit, comparedTotals, mismatches, _
                      "", wbName, wbPath

    Dim summary As String
    summary = BuildVerifySummaryMessage(result, programNorm, selectedYear, localCacheLabel, localCacheSourcePath, localCacheGeneratedAt, _
                                        apiYearUsed, apiCalcFilename, apiSource, apiWarn, _
                                        localTotalNet, apiTotalNet, comparedPerUnit, comparedTotals, mismatches, _
                                        "")

    ShowVerifySummary automationMode, result, summary

    Exit Sub

Fail:
    Dim errMsg As String
    errMsg = Err.Description

    On Error Resume Next
    RecordVerifyState startedAt, "MISMATCH", programNorm, selectedYear, _
                      localCacheStatus, localCacheLabel, localCacheSourcePath, localCacheGeneratedAt, _
                      "", "", "", "", _
                      localTotalNet, Empty, False, False, Nothing, _
                      errMsg, wbName, wbPath
    On Error GoTo 0

    ShowVerifySummary automationMode, "MISMATCH", "MISMATCH — Verify Manual Rents (API) failed." & vbCrLf & vbCrLf & errMsg
End Sub

Private Sub ShowVerifySummary(ByVal automationMode As Boolean, ByVal result As String, ByVal summary As String)
    If automationMode Then
        On Error Resume Next
        ShowAMIOptixDiagnostics
        On Error GoTo 0
    Else
        If UCase$(Trim$(result)) = "MATCH" Then
            MsgBox summary, vbInformation, "AMI Optix - Verify Manual Rents (API)"
        Else
            MsgBox summary, vbExclamation, "AMI Optix - Verify Manual Rents (API)"
        End If
    End If
End Sub

'-------------------------------------------------------------------------------
' INTERNALS
'-------------------------------------------------------------------------------

Private Function ReadProgramUnits(programNorm As String) As Collection
    On Error GoTo Fail

    If ActiveWorkbook Is Nothing Then Exit Function

    Dim prevSheet As Worksheet
    Set prevSheet = Nothing
    On Error Resume Next
    Set prevSheet = ActiveSheet
    On Error GoTo Fail

    Dim prevAddress As String
    prevAddress = ""
    On Error Resume Next
    prevAddress = ActiveCell.Address(False, False)
    On Error GoTo Fail

    Dim dataWs As Worksheet
    Set dataWs = Nothing
    On Error Resume Next
    If UCase$(Trim$(programNorm)) = "MIH" Then
        Set dataWs = ActiveWorkbook.Worksheets("MIH")
        If dataWs Is Nothing Then Set dataWs = ActiveWorkbook.Worksheets("RentRoll")
        If dataWs Is Nothing Then Set dataWs = ActiveWorkbook.Worksheets("UAP")
        If dataWs Is Nothing Then Set dataWs = ActiveWorkbook.Worksheets("PROJECT WORKSHEET")
    Else
        Set dataWs = ActiveWorkbook.Worksheets("UAP")
    End If
    On Error GoTo Fail

    If Not dataWs Is Nothing Then dataWs.Activate
    Set ReadProgramUnits = ReadUnitData()

    ' Best-effort restore user context.
    If Not prevSheet Is Nothing Then
        prevSheet.Activate
        If prevAddress <> "" Then
            On Error Resume Next
            prevSheet.Range(prevAddress).Select
            On Error GoTo 0
        End If
    End If
    Exit Function

Fail:
    Set ReadProgramUnits = Nothing
End Function

Private Function GetSelectedRentRollYearSettingLocal() As Long
    Dim raw As String
    raw = GetSetting("AMI_Optix", "RentRollYears", "SelectedYear", "2025")

    Dim y As Long
    y = 2025
    On Error Resume Next
    y = CLng(raw)
    On Error GoTo 0

    If y < 2022 Or y > 2026 Then y = 2025
    GetSelectedRentRollYearSettingLocal = y
End Function

Private Function TryReadMIHInputsQuiet(ByRef mihOption As String, ByRef residentialSF As Double, ByRef maxBandPercent As Long, ByRef errMsg As String) As Boolean
    On Error GoTo Fail

    TryReadMIHInputsQuiet = False
    errMsg = ""

    If ActiveWorkbook Is Nothing Then
        errMsg = "No workbook is open."
        Exit Function
    End If

    Dim wsMIH As Worksheet
    Dim wsProg As Worksheet
    Set wsMIH = Nothing
    Set wsProg = Nothing

    On Error Resume Next
    Set wsMIH = ActiveWorkbook.Worksheets("MIH")
    Set wsProg = ActiveWorkbook.Worksheets("Prog")
    On Error GoTo Fail

    If wsMIH Is Nothing Then
        errMsg = "MIH run requires a sheet named 'MIH'."
        Exit Function
    End If
    If wsProg Is Nothing Then
        errMsg = "MIH run requires a sheet named 'Prog' (for OptionSelected)."
        Exit Function
    End If

    residentialSF = 0#
    If Not TryFindNetFloorAreaQuiet(wsMIH, residentialSF) Then
        Dim v As Variant
        v = wsMIH.Range("J21").Value
        If Not IsNumeric(v) Or CDbl(v) <= 0 Then
            errMsg = "MIH residential SF is missing (expected Net Floor Area value or MIH!J21)."
            Exit Function
        End If
        residentialSF = CDbl(v)
    End If

    mihOption = Trim$(CStr(wsProg.Range("K4").Value))
    If mihOption = "" Then
        errMsg = "MIH option is missing (expected 'Option 1' or 'Option 4' in Prog!K4)."
        Exit Function
    End If

    Dim capFactor As Variant
    capFactor = wsProg.Range("I4").Value
    If IsNumeric(capFactor) Then
        maxBandPercent = CLng(CDbl(capFactor) * 100)
    Else
        maxBandPercent = 135
    End If

    TryReadMIHInputsQuiet = True
    Exit Function

Fail:
    errMsg = Err.Description
    TryReadMIHInputsQuiet = False
End Function

Private Function TryFindNetFloorAreaQuiet(wsMIH As Worksheet, ByRef netFloorArea As Double) As Boolean
    On Error GoTo Fail

    TryFindNetFloorAreaQuiet = False
    netFloorArea = 0#
    If wsMIH Is Nothing Then Exit Function

    Dim found As Range
    Set found = wsMIH.Cells.Find(What:="Net Floor Area", After:=wsMIH.Cells(1, 1), LookIn:=xlValues, LookAt:=xlPart, _
                                 SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    If found Is Nothing Then Exit Function

    Dim i As Long
    For i = 1 To 6
        Dim v As Variant
        v = found.Offset(0, i).Value
        If IsNumeric(v) And CDbl(v) > 0 Then
            netFloorArea = CDbl(v)
            TryFindNetFloorAreaQuiet = True
            Exit Function
        End If
    Next i
    Exit Function

Fail:
    TryFindNetFloorAreaQuiet = False
End Function

Private Function DictGetString(d As Object, key As String, Optional defaultValue As String = "") As String
    On Error GoTo SafeExit
    DictGetString = defaultValue
    If d Is Nothing Then Exit Function
    If d.Exists(key) Then DictGetString = CStr(d(key))
    Exit Function
SafeExit:
    DictGetString = defaultValue
End Function

Private Function ExtractErrorsList(apiObj As Object) As String
    On Error GoTo SafeExit
    ExtractErrorsList = ""
    If apiObj Is Nothing Then Exit Function
    If Not apiObj.Exists("errors") Then Exit Function

    Dim errs As Object
    Set errs = Nothing
    Set errs = apiObj("errors")

    Dim buf As String
    buf = ""

    Dim i As Long
    If Not errs Is Nothing Then
        If TypeName(errs) = "Collection" Then
        For i = 1 To errs.Count
            buf = buf & "- " & CStr(errs(i)) & vbCrLf
        Next i
        End If
    End If

    ExtractErrorsList = Trim$(buf)
    Exit Function

SafeExit:
    ExtractErrorsList = ""
End Function

Private Function SumDictValues(d As Object) As Double
    On Error GoTo SafeExit
    SumDictValues = 0#
    If d Is Nothing Then Exit Function

    Dim k As Variant
    For Each k In d.Keys
        If IsNumeric(d(k)) Then SumDictValues = SumDictValues + CDbl(d(k))
    Next k
    Exit Function

SafeExit:
    SumDictValues = 0#
End Function

Private Function MakeMismatch(unitId As String, localVal As Variant, apiVal As Variant, deltaVal As Variant, reason As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("unit_id") = CStr(unitId)
    d("local_monthly_rent") = localVal
    d("api_monthly_rent") = apiVal
    d("delta") = deltaVal
    d("reason") = CStr(reason)
    Set MakeMismatch = d
End Function

Private Sub RecordVerifyState( _
    startedAt As String, _
    result As String, _
    programNorm As String, _
    rentRollYear As Long, _
    localCacheStatus As String, _
    localCacheLabel As String, _
    localCacheSourcePath As String, _
    localCacheGeneratedAt As String, _
    apiYearUsed As String, _
    apiCalcFilename As String, _
    apiSource As String, _
    apiWarning As String, _
    localNetMonthly As Double, _
    apiNetMonthly As Variant, _
    comparedPerUnit As Boolean, _
    comparedTotals As Boolean, _
    mismatches As Collection, _
    errorsText As String, _
    workbookName As String, _
    workbookPath As String _
)
    On Error Resume Next

    Dim s As Object
    Set s = CreateObject("Scripting.Dictionary")
    s("timestamp") = startedAt
    s("result") = CStr(result)
    s("program") = CStr(programNorm)
    s("rent_roll_year") = CLng(rentRollYear)

    s("workbook_name") = CStr(workbookName)
    s("workbook_path") = CStr(workbookPath)

    s("local_cache_status") = CStr(localCacheStatus)
    s("local_cache_source_label") = CStr(localCacheLabel)
    s("local_cache_source_path") = CStr(localCacheSourcePath)
    s("local_cache_generated_at") = CStr(localCacheGeneratedAt)

    s("api_rent_roll_year_used") = CStr(apiYearUsed)
    s("api_calculator_filename") = CStr(apiCalcFilename)
    s("api_rent_schedule_source") = CStr(apiSource)
    s("api_rent_schedule_warning") = CStr(apiWarning)

    s("compare_per_unit") = comparedPerUnit
    s("compare_totals") = comparedTotals

    s("tolerance_unit_dollars") = AMI_OPTIX_VERIFY_TOLERANCE_UNIT_DOLLARS
    s("tolerance_total_dollars") = AMI_OPTIX_VERIFY_TOLERANCE_TOTAL_DOLLARS

    s("local_net_monthly") = Round(localNetMonthly, 2)

    If Not IsEmpty(apiNetMonthly) And IsNumeric(apiNetMonthly) Then
        s("api_net_monthly") = Round(CDbl(apiNetMonthly), 2)
        s("total_delta") = Round(localNetMonthly - CDbl(apiNetMonthly), 2)
    Else
        s("api_net_monthly") = ""
        s("total_delta") = ""
    End If

    If Not mismatches Is Nothing Then
        Set s("mismatches") = mismatches
    End If

    s("errors") = CStr(errorsText)

    Set g_AMIOptixLastVerify = s
End Sub

Private Function BuildVerifySummaryMessage( _
    result As String, _
    programNorm As String, _
    rentRollYear As Long, _
    localCacheLabel As String, _
    localCacheSourcePath As String, _
    localCacheGeneratedAt As String, _
    apiYearUsed As String, _
    apiCalcFilename As String, _
    apiSource As String, _
    apiWarning As String, _
    localTotalNet As Double, _
    apiTotalNet As Variant, _
    comparedPerUnit As Boolean, _
    comparedTotals As Boolean, _
    mismatches As Collection, _
    extraNote As String _
) As String
    Dim buf As String
    buf = ""

    buf = buf & CStr(result) & " — Verify Manual Rents (API)" & vbCrLf & vbCrLf

    buf = buf & "Year: " & CStr(rentRollYear) & vbCrLf
    buf = buf & "Program: " & CStr(programNorm) & vbCrLf

    Dim localLine As String
    localLine = Trim$(CStr(localCacheLabel) & " | " & CStr(localCacheSourcePath))
    If localLine = "|" Or localLine = "" Then localLine = "(unknown)"
    buf = buf & "Local cache: " & localLine & vbCrLf
    If Trim$(localCacheGeneratedAt) <> "" Then
        buf = buf & "Local cache built: " & localCacheGeneratedAt & vbCrLf
    End If

    Dim apiLine As String
    apiLine = ""
    If Trim$(apiYearUsed) <> "" Then apiLine = apiLine & "year_used=" & apiYearUsed
    If Trim$(apiCalcFilename) <> "" Then
        If apiLine <> "" Then apiLine = apiLine & " | "
        apiLine = apiLine & "calculator=" & apiCalcFilename
    End If
    If Trim$(apiSource) <> "" Then
        If apiLine <> "" Then apiLine = apiLine & " | "
        apiLine = apiLine & "source=" & apiSource
    End If
    If apiLine = "" Then apiLine = "(unknown)"
    buf = buf & "API: " & apiLine & vbCrLf

    If Trim$(apiWarning) <> "" Then
        buf = buf & "API warning: " & apiWarning & vbCrLf
    End If

    buf = buf & vbCrLf

    If comparedTotals Then
        buf = buf & "Totals (net monthly): local=" & FormatCurrency(localTotalNet, 0) & " | api=" & FormatCurrency(CDbl(apiTotalNet), 0) & _
              " | Δ=" & FormatCurrency(localTotalNet - CDbl(apiTotalNet), 0) & vbCrLf
    Else
        buf = buf & "Totals: (not available from API response)" & vbCrLf
    End If

    If Not comparedPerUnit And comparedTotals Then
        buf = buf & "Per-unit rents: (not available from API response; compared totals only)" & vbCrLf
    End If

    buf = buf & "Tolerance: " & FormatCurrency(AMI_OPTIX_VERIFY_TOLERANCE_UNIT_DOLLARS, 0) & "/unit, " & _
          FormatCurrency(AMI_OPTIX_VERIFY_TOLERANCE_TOTAL_DOLLARS, 0) & " total" & vbCrLf

    If result <> "MATCH" Then
        If Not mismatches Is Nothing Then
            Dim top As Collection
            Set top = GetTopMismatches(mismatches, 5)

            If Not top Is Nothing And top.Count > 0 Then
                buf = buf & vbCrLf & "Top mismatches (unit_id | local | api | Δ):" & vbCrLf

                Dim i As Long
                For i = 1 To top.Count
                    Dim m As Object
                    Set m = top(i)
                    If Not m Is Nothing Then
                        buf = buf & "- " & CStr(m("unit_id")) & " | " & _
                              FormatMaybeCurrency(m("local_monthly_rent")) & " | " & _
                              FormatMaybeCurrency(m("api_monthly_rent")) & " | " & _
                              FormatMaybeCurrency(m("delta")) & vbCrLf
                    End If
                Next i
            End If
        End If

        buf = buf & vbCrLf & "See Diagnostics → Rent Tables Status → Verify Manual Rents (API) for details."
    End If

    If Trim$(extraNote) <> "" Then
        buf = buf & vbCrLf & vbCrLf & extraNote
    End If

    BuildVerifySummaryMessage = buf
End Function

Private Function GetTopMismatches(mismatches As Collection, topN As Long) As Collection
    On Error GoTo SafeExit

    Set GetTopMismatches = Nothing
    If mismatches Is Nothing Then Exit Function
    If mismatches.Count = 0 Then Exit Function
    If topN <= 0 Then Exit Function

    Dim n As Long
    n = mismatches.Count

    Dim arr() As Variant
    Dim score() As Double
    ReDim arr(1 To n)
    ReDim score(1 To n)

    Dim i As Long
    For i = 1 To n
        Set arr(i) = mismatches(i)
        score(i) = MismatchScore(arr(i))
    Next i

    ' Insertion sort by score DESC (n is small in practice).
    Dim j As Long
    For i = 2 To n
        Dim cur As Variant
        Dim curScore As Double
        Set cur = arr(i)
        curScore = score(i)

        j = i - 1
        Do While j >= 1 And score(j) < curScore
            Set arr(j + 1) = arr(j)
            score(j + 1) = score(j)
            j = j - 1
        Loop
        Set arr(j + 1) = cur
        score(j + 1) = curScore
    Next i

    Dim out As Collection
    Set out = New Collection

    Dim limit As Long
    limit = topN
    If limit > n Then limit = n

    For i = 1 To limit
        out.Add arr(i)
    Next i

    Set GetTopMismatches = out
    Exit Function

SafeExit:
    Set GetTopMismatches = Nothing
End Function

Private Function MismatchScore(ByVal m As Variant) As Double
    On Error GoTo SafeExit
    MismatchScore = 0#

    Dim obj As Object
    Set obj = Nothing
    If Not IsObject(m) Then Exit Function
    Set obj = m
    If obj Is Nothing Then Exit Function

    Dim d As Variant
    d = Empty
    If obj.Exists("delta") Then d = obj("delta")
    If IsNumeric(d) Then
        MismatchScore = Abs(CDbl(d))
    Else
        ' Missing values should float to the top.
        MismatchScore = 1E+99
    End If
    Exit Function

SafeExit:
    MismatchScore = 0#
End Function

Private Function FormatMaybeCurrency(v As Variant) As String
    On Error GoTo SafeExit
    If IsEmpty(v) Then
        FormatMaybeCurrency = "(missing)"
    ElseIf IsNumeric(v) Then
        FormatMaybeCurrency = FormatCurrency(CDbl(v), 0)
    Else
        FormatMaybeCurrency = CStr(v)
    End If
    Exit Function

SafeExit:
    FormatMaybeCurrency = "(unknown)"
End Function
