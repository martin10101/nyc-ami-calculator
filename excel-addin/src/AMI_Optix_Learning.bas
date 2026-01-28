Attribute VB_Name = "AMI_Optix_Learning"
'===============================================================================
' AMI OPTIX - AI Learning (Soft Preferences Only)
'
' Learning is intentionally constrained to "soft preferences" (premium weights).
' It must NOT change any hard compliance constraints.
'===============================================================================
Option Explicit

Private Const LEARNING_REGISTRY_PATH As String = "AMI_Optix"
Private Const LEARNING_REGISTRY_SECTION As String = "Learning"

Private Const DEFAULT_BASE_FLOOR_WEIGHT As Double = 0.45
Private Const DEFAULT_BASE_SF_WEIGHT As Double = 0.3
Private Const DEFAULT_BASE_BEDROOM_WEIGHT As Double = 0.15
Private Const DEFAULT_BASE_BALCONY_WEIGHT As Double = 0.1

' Append-only run log (single file; one JSON object per line).
Private Const RUN_LOG_FILE_NAME As String = "ami_optix_runs.jsonl"

' Learning modes (stored as strings in registry)
Public Const LEARNING_MODE_OFF As String = "OFF"
Public Const LEARNING_MODE_SHADOW As String = "SHADOW"
Public Const LEARNING_MODE_ON As String = "ON"

Public Function GetLearningProfileKey(program As String, mihOption As String) As String
    Dim programNorm As String
    programNorm = UCase(Trim(program))
    If programNorm = "" Then programNorm = "UAP"

    If programNorm <> "MIH" Then
        GetLearningProfileKey = "UAP"
        Exit Function
    End If

    Dim optNorm As String
    optNorm = UCase(Trim(mihOption))
    If InStr(optNorm, "4") > 0 Then
        GetLearningProfileKey = "MIH_OPTION_4"
    Else
        GetLearningProfileKey = "MIH_OPTION_1"
    End If
End Function

Public Function GetLearningMode(profileKey As String) As String
    Dim key As String
    key = "Mode_" & UCase(Trim(profileKey))

    Dim value As String
    value = UCase(Trim(GetSetting(LEARNING_REGISTRY_PATH, LEARNING_REGISTRY_SECTION, key, LEARNING_MODE_OFF)))

    Select Case value
        Case LEARNING_MODE_OFF, LEARNING_MODE_SHADOW, LEARNING_MODE_ON
            GetLearningMode = value
        Case Else
            GetLearningMode = LEARNING_MODE_OFF
    End Select
End Function

Public Sub SetLearningMode(profileKey As String, mode As String)
    Dim key As String
    key = "Mode_" & UCase(Trim(profileKey))
    SaveSetting LEARNING_REGISTRY_PATH, LEARNING_REGISTRY_SECTION, key, UCase(Trim(mode))
End Sub

Public Function GetLearningCompareBaseline(profileKey As String) As Boolean
    Dim key As String
    key = "CompareBaseline_" & UCase(Trim(profileKey))
    GetLearningCompareBaseline = (GetSetting(LEARNING_REGISTRY_PATH, LEARNING_REGISTRY_SECTION, key, "1") = "1")
End Function

Public Sub SetLearningCompareBaseline(profileKey As String, enabled As Boolean)
    Dim key As String
    key = "CompareBaseline_" & UCase(Trim(profileKey))
    SaveSetting LEARNING_REGISTRY_PATH, LEARNING_REGISTRY_SECTION, key, IIf(enabled, "1", "0")
End Sub

Public Function GetLearningLogRootPath() As String
    Dim defaultPath As String
    defaultPath = Environ$("USERPROFILE") & "\Documents\AMI_Optix_Learning"
    GetLearningLogRootPath = GetSetting(LEARNING_REGISTRY_PATH, LEARNING_REGISTRY_SECTION, "LogRootPath", defaultPath)
End Function

Public Sub SetLearningLogRootPath(path As String)
    SaveSetting LEARNING_REGISTRY_PATH, LEARNING_REGISTRY_SECTION, "LogRootPath", CStr(path)
End Sub

Public Function GetLearningProfileFolder(profileKey As String) As String
    Dim rootPath As String
    rootPath = GetLearningLogRootPath()
    GetLearningProfileFolder = rootPath & "\" & UCase(Trim(profileKey))
End Function

Public Sub EnsureFolderExists(path As String)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Sub

    If Not fso.FolderExists(path) Then
        fso.CreateFolder path
    End If
End Sub

Private Function NewGuidString() As String
    On Error GoTo Fail
    Dim guid As String
    guid = CreateObject("Scriptlet.TypeLib").GUID
    guid = Replace(guid, "{", "")
    guid = Replace(guid, "}", "")
    NewGuidString = guid
    Exit Function
Fail:
    NewGuidString = CStr(Int((99999999# * Rnd) + 1))
End Function

Private Function EscapeJSONString(str As String) As String
    Dim result As String
    result = str
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCr, "\r")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")
    EscapeJSONString = result
End Function

Private Function InvariantNumber(value As Double) As String
    Dim s As String
    s = Trim$(CStr(value))
    s = Replace(s, ",", ".")
    InvariantNumber = s
End Function

Public Function BuildProjectOverridesJson(premiumWeights As Object, notes As String) As String
    If premiumWeights Is Nothing Then
        BuildProjectOverridesJson = ""
        Exit Function
    End If

    Dim json As String
    json = "{"
    json = json & """premiumWeights"": {"
    json = json & """floor"": " & InvariantNumber(CDbl(premiumWeights("floor"))) & ", "
    json = json & """net_sf"": " & InvariantNumber(CDbl(premiumWeights("net_sf"))) & ", "
    json = json & """bedrooms"": " & InvariantNumber(CDbl(premiumWeights("bedrooms"))) & ", "
    json = json & """balcony"": " & InvariantNumber(CDbl(premiumWeights("balcony")))
    json = json & "}"

    If Trim$(notes) <> "" Then
        json = json & ", ""notes"": [""" & EscapeJSONString(notes) & """]"
    End If

    json = json & "}"
    BuildProjectOverridesJson = json
End Function

Public Sub LogLearningEvent(profileKey As String, eventName As String, jsonBody As String)
    ' Backward-compatible wrapper. We no longer create per-run JSON files (file spam).
    ' Instead, append one JSON object per line to the shared JSONL log file.
    On Error GoTo Fail

    If Trim$(jsonBody) = "" Then jsonBody = "{}"
    Call AppendRunLog(eventName, jsonBody)
    Exit Sub

Fail:
End Sub

Public Sub LogSolverRun(profileKey As String, programNorm As String, mihOption As String, learningMode As String, _
                        compareBaseline As Boolean, premiumWeights As Object, apiResult As Object)
    On Error GoTo Fail

    Dim wbName As String
    Dim wbPath As String
    wbName = ""
    wbPath = ""
    If Not ActiveWorkbook Is Nothing Then
        wbName = ActiveWorkbook.Name
        wbPath = ActiveWorkbook.FullName
    End If

    Dim weightsJson As String
    weightsJson = "null"
    If Not premiumWeights Is Nothing Then
        weightsJson = "{""floor"": " & InvariantNumber(CDbl(premiumWeights("floor"))) & _
                      ", ""net_sf"": " & InvariantNumber(CDbl(premiumWeights("net_sf"))) & _
                      ", ""bedrooms"": " & InvariantNumber(CDbl(premiumWeights("bedrooms"))) & _
                      ", ""balcony"": " & InvariantNumber(CDbl(premiumWeights("balcony"))) & "}"
    End If

    Dim scenarioCount As Long
    scenarioCount = 0
    If Not apiResult Is Nothing Then
        If apiResult.Exists("scenarios") Then
            scenarioCount = apiResult("scenarios").Count
        End If
    End If

    Dim hasLearningCompare As Boolean
    hasLearningCompare = False

    Dim changedUnitCount As Long
    changedUnitCount = 0

    Dim baselineWaamiPercent As Double
    baselineWaamiPercent = 0#

    Dim learnedWaamiPercent As Double
    learnedWaamiPercent = 0#

    If Not apiResult Is Nothing Then
        If apiResult.Exists("learning") Then
            hasLearningCompare = True
            Dim learningObj As Object
            Set learningObj = apiResult("learning")

            If Not learningObj Is Nothing Then
                If learningObj.Exists("diff") Then
                    Dim diffObj As Object
                    Set diffObj = learningObj("diff")
                    If Not diffObj Is Nothing Then
                        If diffObj.Exists("changed_unit_count") Then
                            changedUnitCount = CLng(diffObj("changed_unit_count"))
                        End If
                    End If
                End If

                If learningObj.Exists("baseline") Then
                    Dim baseObj As Object
                    Set baseObj = learningObj("baseline")
                    If Not baseObj Is Nothing Then
                        If baseObj.Exists("absolute_best") Then
                            Dim baseBest As Object
                            Set baseBest = baseObj("absolute_best")
                            If Not baseBest Is Nothing Then
                                If baseBest.Exists("waami_percent") Then baselineWaamiPercent = CDbl(baseBest("waami_percent"))
                            End If
                        End If
                    End If
                End If

                If learningObj.Exists("learned") Then
                    Dim learnedObj As Object
                    Set learnedObj = learningObj("learned")
                    If Not learnedObj Is Nothing Then
                        If learnedObj.Exists("absolute_best") Then
                            Dim learnedBest As Object
                            Set learnedBest = learnedObj("absolute_best")
                            If Not learnedBest Is Nothing Then
                                If learnedBest.Exists("waami_percent") Then learnedWaamiPercent = CDbl(learnedBest("waami_percent"))
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    ' Scenario keys + solver notes (kept out of client-facing sheets).
    Dim keysJson As String
    keysJson = "[]"
    On Error Resume Next
    If Not apiResult Is Nothing Then
        If apiResult.Exists("scenarios") Then
            Dim scenarios As Object
            Set scenarios = apiResult("scenarios")
            If Not scenarios Is Nothing Then
                Dim scenarioKey As Variant
                Dim buf As String
                buf = "["
                Dim first As Boolean
                first = True
                For Each scenarioKey In scenarios.Keys
                    If Not first Then buf = buf & ","
                    buf = buf & """" & EscapeJSONString(CStr(scenarioKey)) & """"
                    first = False
                Next scenarioKey
                buf = buf & "]"
                keysJson = buf
            End If
        End If
    End If
    On Error GoTo Fail

    Dim notesJson As String
    notesJson = "[]"
    On Error Resume Next
    If Not apiResult Is Nothing Then
        If apiResult.Exists("notes") Then
            Dim notes As Object
            Set notes = apiResult("notes")
            If Not notes Is Nothing Then
                Dim n As Long
                Dim nbuf As String
                nbuf = "["
                For n = 1 To notes.Count
                    If n > 1 Then nbuf = nbuf & ","
                    nbuf = nbuf & """" & EscapeJSONString(CStr(notes(n))) & """"
                Next n
                nbuf = nbuf & "]"
                notesJson = nbuf
            End If
        End If
    End If
    On Error GoTo Fail

    Dim payload As String
    payload = "{"
    payload = payload & """schema_version"":1,"
    payload = payload & """created_local"":""" & EscapeJSONString(Format$(Now, "yyyy-mm-dd hh:nn:ss")) & ""","
    payload = payload & """profile_key"":""" & EscapeJSONString(profileKey) & ""","
    payload = payload & """program"":""" & EscapeJSONString(programNorm) & ""","
    payload = payload & """mih_option"":""" & EscapeJSONString(mihOption) & ""","
    payload = payload & """workbook_name"":""" & EscapeJSONString(wbName) & ""","
    payload = payload & """workbook_path"":""" & EscapeJSONString(wbPath) & ""","
    payload = payload & """learning_mode"":""" & EscapeJSONString(learningMode) & ""","
    payload = payload & """compare_baseline"":" & IIf(compareBaseline, "true", "false") & ","
    payload = payload & """premium_weights_sent"":" & weightsJson & ","
    payload = payload & """scenario_count"":" & scenarioCount & ","
    payload = payload & """scenario_keys"":" & keysJson & ","
    payload = payload & """notes"":" & notesJson & ","
    payload = payload & """has_learning_compare"":" & IIf(hasLearningCompare, "true", "false") & ","
    payload = payload & """baseline_waami_percent"":" & InvariantNumber(baselineWaamiPercent) & ","
    payload = payload & """learned_waami_percent"":" & InvariantNumber(learnedWaamiPercent) & ","
    payload = payload & """changed_unit_count"":" & changedUnitCount
    payload = payload & "}"

    Call AppendRunLog("solver_run", payload)
    Exit Sub

Fail:
End Sub

'-------------------------------------------------------------------------------
' RUN LOG (APPEND-ONLY JSONL)
'-------------------------------------------------------------------------------

Public Function GetRunLogFilePath() As String
    Dim rootPath As String
    rootPath = GetLearningLogRootPath()
    GetRunLogFilePath = rootPath & "\" & RUN_LOG_FILE_NAME
End Function

Public Sub AppendRunLog(eventName As String, payloadJson As String)
    ' Appends one JSON line to the shared run log file.
    ' The file is reused across runs (no file spam).
    On Error GoTo Fail

    Dim rootPath As String
    rootPath = GetLearningLogRootPath()
    Call EnsureFolderExists(rootPath)

    Dim filePath As String
    filePath = GetRunLogFilePath()

    Dim ts As String
    ts = Format$(Now, "yyyy-mm-ddThh:nn:ss")

    If Trim$(payloadJson) = "" Then payloadJson = "{}"

    Dim line As String
    line = "{""timestamp"":""" & ts & """,""event"":""" & EscapeJSONString(CStr(eventName)) & """,""payload"":" & payloadJson & "}"

    Dim ff As Integer
    ff = FreeFile
    Open filePath For Append As #ff
    Print #ff, line
    Close #ff
    Exit Sub

Fail:
    On Error Resume Next
    If ff <> 0 Then Close #ff
End Sub

Public Sub AppendSolverNotesToRunLog(apiResult As Object)
    ' Writes solver notes + scenario keys to the shared run log file.
    ' This replaces dumping verbose "Solver Notes" onto client-facing sheets.
    On Error GoTo Fail

    If apiResult Is Nothing Then Exit Sub

    Dim wbName As String
    Dim wbPath As String
    wbName = ""
    wbPath = ""
    On Error Resume Next
    If Not ActiveWorkbook Is Nothing Then
        wbName = ActiveWorkbook.Name
        wbPath = ActiveWorkbook.FullName
    End If
    On Error GoTo 0

    Dim programNorm As String
    Dim mihOption As String
    programNorm = ""
    mihOption = ""
    On Error Resume Next
    If apiResult.Exists("project_summary") Then
        Dim ps As Object
        Set ps = apiResult("project_summary")
        If Not ps Is Nothing Then
            If ps.Exists("program") Then programNorm = CStr(ps("program"))
            If ps.Exists("mih_option") Then mihOption = CStr(ps("mih_option"))
        End If
    End If
    On Error GoTo 0

    Dim keysJson As String
    keysJson = "[]"
    On Error Resume Next
    If apiResult.Exists("scenarios") Then
        Dim scenarios As Object
        Set scenarios = apiResult("scenarios")
        If Not scenarios Is Nothing Then
            Dim scenarioKey As Variant
            Dim buf As String
            buf = "["
            Dim first As Boolean
            first = True
            For Each scenarioKey In scenarios.Keys
                If Not first Then buf = buf & ","
                buf = buf & """" & EscapeJSONString(CStr(scenarioKey)) & """"
                first = False
            Next scenarioKey
            buf = buf & "]"
            keysJson = buf
        End If
    End If
    On Error GoTo 0

    Dim notesJson As String
    notesJson = "[]"
    On Error Resume Next
    If apiResult.Exists("notes") Then
        Dim notes As Object
        Set notes = apiResult("notes")
        If Not notes Is Nothing Then
            Dim n As Long
            Dim nbuf As String
            nbuf = "["
            For n = 1 To notes.Count
                If n > 1 Then nbuf = nbuf & ","
                nbuf = nbuf & """" & EscapeJSONString(CStr(notes(n))) & """"
            Next n
            nbuf = nbuf & "]"
            notesJson = nbuf
        End If
    End If
    On Error GoTo 0

    Dim payload As String
    payload = "{"
    payload = payload & """workbook_name"":""" & EscapeJSONString(wbName) & ""","
    payload = payload & """workbook_path"":""" & EscapeJSONString(wbPath) & ""","
    payload = payload & """program"":""" & EscapeJSONString(programNorm) & ""","
    payload = payload & """mih_option"":""" & EscapeJSONString(mihOption) & ""","
    payload = payload & """scenario_keys"":" & keysJson & ","
    payload = payload & """notes"":" & notesJson
    payload = payload & "}"

    Call AppendRunLog("solver_run", payload)
    Exit Sub

Fail:
    ' Swallow logging errors (network drives can be flaky).
End Sub

Public Sub LogScenarioChoiceToRunLog(profileKey As String, programNorm As String, mihOption As String, _
                                     scenarioNumber As Long, scenarioKey As String, scenario As Object)
    ' Logs which scenario the user/client chose, without making any API calls.
    ' Appends a single JSONL entry to the shared run log file.
    On Error GoTo Fail

    Dim wbName As String
    Dim wbPath As String
    wbName = ""
    wbPath = ""
    On Error Resume Next
    If Not ActiveWorkbook Is Nothing Then
        wbName = ActiveWorkbook.Name
        wbPath = ActiveWorkbook.FullName
    End If
    On Error GoTo 0

    Dim scenarioWaami As Double
    scenarioWaami = 0#
    On Error Resume Next
    If Not scenario Is Nothing Then
        If scenario.Exists("waami") Then scenarioWaami = CDbl(scenario("waami"))
    End If
    On Error GoTo 0

    Dim bandsJson As String
    bandsJson = "[]"
    On Error Resume Next
    If Not scenario Is Nothing Then
        If scenario.Exists("bands") Then
            Dim bandsObj As Object
            Set bandsObj = scenario("bands")
            If Not bandsObj Is Nothing And TypeName(bandsObj) = "Collection" Then
                Dim b As Long
                Dim buf As String
                buf = "["
                For b = 1 To bandsObj.Count
                    If b > 1 Then buf = buf & ","
                    buf = buf & CStr(bandsObj(b))
                Next b
                buf = buf & "]"
                bandsJson = buf
            End If
        End If
    End If
    On Error GoTo 0

    Dim tradeoffsJson As String
    tradeoffsJson = "[]"
    On Error Resume Next
    If Not scenario Is Nothing Then
        If scenario.Exists("tradeoffs") Then
            Dim tObj As Object
            Set tObj = scenario("tradeoffs")
            If Not tObj Is Nothing And TypeName(tObj) = "Collection" Then
                Dim t As Long
                Dim tbuf As String
                tbuf = "["
                For t = 1 To tObj.Count
                    If t > 1 Then tbuf = tbuf & ","
                    tbuf = tbuf & """" & EscapeJSONString(CStr(tObj(t))) & """"
                Next t
                tbuf = tbuf & "]"
                tradeoffsJson = tbuf
            End If
        End If
    End If
    On Error GoTo 0

    Dim rentTotalsJson As String
    rentTotalsJson = "null"
    On Error Resume Next
    If Not scenario Is Nothing Then
        If scenario.Exists("rent_totals") Then
            Dim rt As Object
            Set rt = scenario("rent_totals")
            If Not rt Is Nothing Then
                Dim netMonthly As String
                Dim netAnnual As String
                netMonthly = "null"
                netAnnual = "null"
                If rt.Exists("net_monthly") Then netMonthly = InvariantNumber(CDbl(rt("net_monthly")))
                If rt.Exists("net_annual") Then netAnnual = InvariantNumber(CDbl(rt("net_annual")))
                rentTotalsJson = "{""net_monthly"":" & netMonthly & ",""net_annual"":" & netAnnual & "}"
            End If
        End If
    End If
    On Error GoTo 0

    ' Capture the current workbook unit assignments from the sheet (client_ami values).
    Dim unitsJson As String
    unitsJson = "null"
    On Error Resume Next
    Dim prevSheet As Worksheet
    Set prevSheet = ActiveSheet
    Dim dataWs As Worksheet
    Set dataWs = Nothing
    If UCase$(Trim$(programNorm)) = "MIH" Then
        Set dataWs = ActiveWorkbook.Worksheets("RentRoll")
        If dataWs Is Nothing Then Set dataWs = ActiveWorkbook.Worksheets("MIH")
    Else
        Set dataWs = ActiveWorkbook.Worksheets("UAP")
    End If
    If Not dataWs Is Nothing Then dataWs.Activate
    Dim units As Collection
    Set units = ReadUnitData()
    If Not prevSheet Is Nothing Then prevSheet.Activate
    If Not units Is Nothing Then
        Dim i As Long
        Dim ubuf As String
        ubuf = "["
        For i = 1 To units.Count
            Dim u As Object
            Set u = units(i)
            If i > 1 Then ubuf = ubuf & ","
            ubuf = ubuf & "{"
            ubuf = ubuf & """" & "unit_id" & """:""" & EscapeJSONString(CStr(u("unit_id"))) & """"
            If u.Exists("bedrooms") Then ubuf = ubuf & ",""bedrooms"":" & CLng(u("bedrooms"))
            If u.Exists("net_sf") Then ubuf = ubuf & ",""net_sf"":" & InvariantNumber(CDbl(u("net_sf")))
            If u.Exists("floor") Then ubuf = ubuf & ",""floor"":" & CLng(u("floor"))
            If u.Exists("balcony") Then ubuf = ubuf & ",""balcony"":" & IIf(CBool(u("balcony")), "true", "false")
            If u.Exists("client_ami") Then ubuf = ubuf & ",""client_ami"":" & InvariantNumber(CDbl(u("client_ami")))
            ubuf = ubuf & "}"
        Next i
        ubuf = ubuf & "]"
        unitsJson = ubuf
    End If
    On Error GoTo 0

    Dim payload As String
    payload = "{"
    payload = payload & """profile_key"":""" & EscapeJSONString(profileKey) & ""","
    payload = payload & """workbook_name"":""" & EscapeJSONString(wbName) & ""","
    payload = payload & """workbook_path"":""" & EscapeJSONString(wbPath) & ""","
    payload = payload & """program"":""" & EscapeJSONString(programNorm) & ""","
    payload = payload & """mih_option"":""" & EscapeJSONString(mihOption) & ""","
    payload = payload & """scenario_number"":" & CLng(scenarioNumber) & ","
    payload = payload & """scenario_key"":""" & EscapeJSONString(CStr(scenarioKey)) & ""","
    payload = payload & """scenario_waami"":" & InvariantNumber(scenarioWaami) & ","
    payload = payload & """scenario_bands"":" & bandsJson & ","
    payload = payload & """scenario_tradeoffs"":" & tradeoffsJson & ","
    payload = payload & """scenario_rent_totals"":" & rentTotalsJson & ","
    payload = payload & """workbook_units"":" & unitsJson
    payload = payload & "}"

    Call AppendRunLog("scenario_choice", payload)
    Exit Sub

Fail:
    ' Swallow logging errors.
End Sub

Public Sub LogScenarioApplied(profileKey As String, programNorm As String, mihOption As String, scenarioKey As String, _
                             source As String, scenario As Object)
    On Error GoTo Fail

    Dim wbName As String
    Dim wbPath As String
    wbName = ""
    wbPath = ""
    If Not ActiveWorkbook Is Nothing Then
        wbName = ActiveWorkbook.Name
        wbPath = ActiveWorkbook.FullName
    End If

    Dim unitsJson As String
    unitsJson = "[]"
    If Not scenario Is Nothing Then
        If scenario.Exists("assignments") Then
            Dim assignments As Object
            Set assignments = scenario("assignments")
            Dim i As Long

            unitsJson = "["
            For i = 1 To assignments.Count
                Dim a As Object
                Set a = assignments(i)
                If i > 1 Then unitsJson = unitsJson & ","

                Dim unitId As String
                unitId = CStr(a("unit_id"))

                Dim assignedAmi As Double
                assignedAmi = CDbl(a("assigned_ami"))
                If assignedAmi > 2# Then assignedAmi = assignedAmi / 100#

                unitsJson = unitsJson & "{"
                unitsJson = unitsJson & """unit_id"": """ & EscapeJSONString(unitId) & """"
                If a.Exists("floor") Then unitsJson = unitsJson & ", ""floor"": " & InvariantNumber(CDbl(a("floor")))
                If a.Exists("net_sf") Then unitsJson = unitsJson & ", ""net_sf"": " & InvariantNumber(CDbl(a("net_sf")))
                If a.Exists("bedrooms") Then unitsJson = unitsJson & ", ""bedrooms"": " & CLng(a("bedrooms"))
                If a.Exists("balcony") Then unitsJson = unitsJson & ", ""balcony"": " & IIf(CBool(a("balcony")), "true", "false")
                unitsJson = unitsJson & ", ""assigned_ami"": " & InvariantNumber(assignedAmi)
                unitsJson = unitsJson & "}"
            Next i
            unitsJson = unitsJson & "]"
        End If
    End If

    Dim payload As String
    payload = "{"
    payload = payload & """schema_version"":1,"
    payload = payload & """created_local"":""" & EscapeJSONString(Format$(Now, "yyyy-mm-dd hh:nn:ss")) & ""","
    payload = payload & """profile_key"":""" & EscapeJSONString(profileKey) & ""","
    payload = payload & """program"":""" & EscapeJSONString(programNorm) & ""","
    payload = payload & """mih_option"":""" & EscapeJSONString(mihOption) & ""","
    payload = payload & """workbook_name"":""" & EscapeJSONString(wbName) & ""","
    payload = payload & """workbook_path"":""" & EscapeJSONString(wbPath) & ""","
    payload = payload & """scenario_key"":""" & EscapeJSONString(scenarioKey) & ""","
    payload = payload & """source"":""" & EscapeJSONString(source) & ""","
    payload = payload & """units"":" & unitsJson
    payload = payload & "}"

    Call AppendRunLog("scenario_applied", payload)
    Exit Sub

Fail:
End Sub

Public Function GetBaselinePremiumWeights() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("floor") = DEFAULT_BASE_FLOOR_WEIGHT
    d("net_sf") = DEFAULT_BASE_SF_WEIGHT
    d("bedrooms") = DEFAULT_BASE_BEDROOM_WEIGHT
    d("balcony") = DEFAULT_BASE_BALCONY_WEIGHT
    Set GetBaselinePremiumWeights = d
End Function

Public Function ComputeLearnedPremiumWeights(profileKey As String, Optional maxTrainingEvents As Long = 250) As Object
    On Error GoTo Fail

    Dim baseline As Object
    Set baseline = GetBaselinePremiumWeights()

    Dim folderPath As String
    folderPath = GetLearningProfileFolder(profileKey)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then
        Set ComputeLearnedPremiumWeights = baseline
        Exit Function
    End If
    If Not fso.FolderExists(folderPath) Then
        Set ComputeLearnedPremiumWeights = baseline
        Exit Function
    End If

    Dim nFloor As Long, sumXFloor As Double, sumX2Floor As Double, sumYFloor As Double, sumY2Floor As Double, sumXYFloor As Double
    Dim nSf As Long, sumXSf As Double, sumX2Sf As Double, sumYSf As Double, sumY2Sf As Double, sumXYSf As Double
    Dim nBeds As Long, sumXBeds As Double, sumX2Beds As Double, sumYBeds As Double, sumY2Beds As Double, sumXYBeds As Double
    Dim nBal As Long, sumXBal As Double, sumX2Bal As Double, sumYBal As Double, sumY2Bal As Double, sumXYBal As Double

    Dim trainingEvents As Long
    trainingEvents = 0

    Dim folder As Object
    Set folder = fso.GetFolder(folderPath)

    Dim file As Object
    For Each file In folder.Files
        If LCase$(fso.GetExtensionName(file.Name)) <> "json" Then GoTo NextFile
        If trainingEvents >= maxTrainingEvents Then Exit For

        Dim content As String
        content = ""
        content = file.OpenAsTextStream(1, -2).ReadAll
        If Trim$(content) = "" Then GoTo NextFile

        Dim evt As Object
        Set evt = ParseJSON(content)
        If evt Is Nothing Then GoTo NextFile

        If Not evt.Exists("event") Then GoTo NextFile
        If CStr(evt("event")) <> "scenario_applied" Then GoTo NextFile

        If evt.Exists("source") Then
            If UCase$(CStr(evt("source"))) <> "USER" Then GoTo NextFile
        Else
            GoTo NextFile
        End If

        If Not evt.Exists("units") Then GoTo NextFile

        Dim units As Object
        Set units = evt("units")
        If units Is Nothing Then GoTo NextFile
        If units.Count = 0 Then GoTo NextFile

        Dim maxFloor As Double, maxSf As Double, maxBeds As Double
        maxFloor = 0#: maxSf = 0#: maxBeds = 0#

        Dim i As Long
        For i = 1 To units.Count
            Dim u As Object
            Set u = units(i)
            If u.Exists("floor") Then maxFloor = Application.Max(maxFloor, CDbl(u("floor")))
            If u.Exists("net_sf") Then maxSf = Application.Max(maxSf, CDbl(u("net_sf")))
            If u.Exists("bedrooms") Then maxBeds = Application.Max(maxBeds, CDbl(u("bedrooms")))
        Next i

        Dim eventHasVariance As Boolean
        eventHasVariance = False
        Dim minY As Double, maxY As Double
        minY = 999999#: maxY = -999999#
        For i = 1 To units.Count
            Dim u2 As Object
            Set u2 = units(i)
            If Not u2.Exists("assigned_ami") Then GoTo NextUnitVar
            Dim yv As Double
            yv = CDbl(u2("assigned_ami"))
            If yv > 2# Then yv = yv / 100#
            minY = Application.Min(minY, yv)
            maxY = Application.Max(maxY, yv)
NextUnitVar:
        Next i
        If maxY - minY > 0.0000001 Then eventHasVariance = True
        If Not eventHasVariance Then GoTo NextFile

        trainingEvents = trainingEvents + 1

        For i = 1 To units.Count
            Dim u3 As Object
            Set u3 = units(i)

            If Not u3.Exists("assigned_ami") Then GoTo NextUnit
            Dim y As Double
            y = CDbl(u3("assigned_ami"))
            If y > 2# Then y = y / 100#

            If u3.Exists("floor") And maxFloor > 0 Then
                Dim xF As Double
                xF = CDbl(u3("floor")) / maxFloor
                nFloor = nFloor + 1
                sumXFloor = sumXFloor + xF
                sumX2Floor = sumX2Floor + xF * xF
                sumYFloor = sumYFloor + y
                sumY2Floor = sumY2Floor + y * y
                sumXYFloor = sumXYFloor + xF * y
            End If

            If u3.Exists("net_sf") And maxSf > 0 Then
                Dim xS As Double
                xS = CDbl(u3("net_sf")) / maxSf
                nSf = nSf + 1
                sumXSf = sumXSf + xS
                sumX2Sf = sumX2Sf + xS * xS
                sumYSf = sumYSf + y
                sumY2Sf = sumY2Sf + y * y
                sumXYSf = sumXYSf + xS * y
            End If

            If u3.Exists("bedrooms") And maxBeds > 0 Then
                Dim xB As Double
                xB = CDbl(u3("bedrooms")) / maxBeds
                nBeds = nBeds + 1
                sumXBeds = sumXBeds + xB
                sumX2Beds = sumX2Beds + xB * xB
                sumYBeds = sumYBeds + y
                sumY2Beds = sumY2Beds + y * y
                sumXYBeds = sumXYBeds + xB * y
            End If

            Dim xBal As Double
            xBal = 0#
            If u3.Exists("balcony") Then
                xBal = IIf(CBool(u3("balcony")), 1#, 0#)
            End If
            nBal = nBal + 1
            sumXBal = sumXBal + xBal
            sumX2Bal = sumX2Bal + xBal * xBal
            sumYBal = sumYBal + y
            sumY2Bal = sumY2Bal + y * y
            sumXYBal = sumXYBal + xBal * y

NextUnit:
        Next i

NextFile:
    Next file

    If trainingEvents < 3 Then
        Set ComputeLearnedPremiumWeights = baseline
        Exit Function
    End If

    Dim alpha As Double
    alpha = 0.25
    If trainingEvents < 20 Then
        alpha = alpha * (trainingEvents / 20#)
    End If

    Dim rFloor As Double, rSf As Double, rBeds As Double, rBal As Double
    rFloor = CorrFromSums(nFloor, sumXFloor, sumX2Floor, sumYFloor, sumY2Floor, sumXYFloor)
    rSf = CorrFromSums(nSf, sumXSf, sumX2Sf, sumYSf, sumY2Sf, sumXYSf)
    rBeds = CorrFromSums(nBeds, sumXBeds, sumX2Beds, sumYBeds, sumY2Beds, sumXYBeds)
    rBal = CorrFromSums(nBal, sumXBal, sumX2Bal, sumYBal, sumY2Bal, sumXYBal)

    Dim learned As Object
    Set learned = CreateObject("Scripting.Dictionary")

    learned("floor") = ClampWeight(CDbl(baseline("floor")) * (1# + alpha * rFloor))
    learned("net_sf") = ClampWeight(CDbl(baseline("net_sf")) * (1# + alpha * rSf))
    learned("bedrooms") = ClampWeight(CDbl(baseline("bedrooms")) * (1# + alpha * rBeds))
    learned("balcony") = ClampWeight(CDbl(baseline("balcony")) * (1# + alpha * rBal))

    NormalizeWeights learned
    Set ComputeLearnedPremiumWeights = learned
    Exit Function

Fail:
    Set ComputeLearnedPremiumWeights = GetBaselinePremiumWeights()
End Function

Private Function CorrFromSums(n As Long, sumX As Double, sumX2 As Double, sumY As Double, sumY2 As Double, sumXY As Double) As Double
    On Error GoTo Fail
    If n < 2 Then
        CorrFromSums = 0#
        Exit Function
    End If

    Dim num As Double
    num = (n * sumXY) - (sumX * sumY)

    Dim denX As Double
    denX = (n * sumX2) - (sumX * sumX)

    Dim denY As Double
    denY = (n * sumY2) - (sumY * sumY)

    If denX <= 0# Or denY <= 0# Then
        CorrFromSums = 0#
        Exit Function
    End If

    CorrFromSums = num / Sqr(denX * denY)
    Exit Function

Fail:
    CorrFromSums = 0#
End Function

Private Function ClampWeight(value As Double) As Double
    Dim minW As Double, maxW As Double
    minW = 0.05
    maxW = 0.8
    If value < minW Then value = minW
    If value > maxW Then value = maxW
    ClampWeight = value
End Function

Private Sub NormalizeWeights(weights As Object)
    On Error GoTo Fail
    Dim total As Double
    total = CDbl(weights("floor")) + CDbl(weights("net_sf")) + CDbl(weights("bedrooms")) + CDbl(weights("balcony"))
    If total <= 0# Then Exit Sub

    weights("floor") = CDbl(weights("floor")) / total
    weights("net_sf") = CDbl(weights("net_sf")) / total
    weights("bedrooms") = CDbl(weights("bedrooms")) / total
    weights("balcony") = CDbl(weights("balcony")) / total
    Exit Sub
Fail:
End Sub
