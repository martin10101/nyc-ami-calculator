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
    On Error GoTo Fail

    Dim folderPath As String
    folderPath = GetLearningProfileFolder(profileKey)

    EnsureFolderExists GetLearningLogRootPath()
    EnsureFolderExists folderPath

    Dim fileName As String
    fileName = Format$(Now, "yyyymmdd_hhnnss") & "_" & NewGuidString() & "_" & LCase$(eventName) & ".json"

    Dim fullPath As String
    fullPath = folderPath & "\" & fileName

    Dim ff As Integer
    ff = FreeFile
    Open fullPath For Output As #ff
    Print #ff, jsonBody
    Close #ff
    Exit Sub

Fail:
    On Error Resume Next
    If ff <> 0 Then Close #ff
End Sub

Public Sub LogSolverRun(profileKey As String, programNorm As String, mihOption As String, learningMode As String, _
                        compareBaseline As Boolean, premiumWeights As Object, apiResult As Object)
    On Error GoTo Fail

    Dim wbName As String
    wbName = ""
    If Not ActiveWorkbook Is Nothing Then
        wbName = ActiveWorkbook.Name
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

    Dim json As String
    json = "{"
    json = json & """schema_version"": 1,"
    json = json & """event"": ""solver_run"","
    json = json & """created_local"": """ & EscapeJSONString(Format$(Now, "yyyy-mm-dd hh:nn:ss")) & ""","
    json = json & """profile"": """ & EscapeJSONString(profileKey) & ""","
    json = json & """program"": """ & EscapeJSONString(programNorm) & ""","
    json = json & """mih_option"": """ & EscapeJSONString(mihOption) & ""","
    json = json & """workbook"": """ & EscapeJSONString(wbName) & ""","
    json = json & """learning_mode"": """ & EscapeJSONString(learningMode) & ""","
    json = json & """compare_baseline"": " & IIf(compareBaseline, "true", "false") & ","
    json = json & """premium_weights_sent"": " & weightsJson & ","
    json = json & """scenario_count"": " & scenarioCount & ","
    json = json & """has_learning_compare"": " & IIf(hasLearningCompare, "true", "false") & ","
    json = json & """baseline_waami_percent"": " & InvariantNumber(baselineWaamiPercent) & ","
    json = json & """learned_waami_percent"": " & InvariantNumber(learnedWaamiPercent) & ","
    json = json & """changed_unit_count"": " & changedUnitCount
    json = json & "}"

    LogLearningEvent profileKey, "solver_run", json
    Exit Sub

Fail:
End Sub

Public Sub LogScenarioApplied(profileKey As String, programNorm As String, mihOption As String, scenarioKey As String, _
                             source As String, scenario As Object)
    On Error GoTo Fail

    Dim wbName As String
    wbName = ""
    If Not ActiveWorkbook Is Nothing Then
        wbName = ActiveWorkbook.Name
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

    Dim json As String
    json = "{"
    json = json & """schema_version"": 1,"
    json = json & """event"": ""scenario_applied"","
    json = json & """created_local"": """ & EscapeJSONString(Format$(Now, "yyyy-mm-dd hh:nn:ss")) & ""","
    json = json & """profile"": """ & EscapeJSONString(profileKey) & ""","
    json = json & """program"": """ & EscapeJSONString(programNorm) & ""","
    json = json & """mih_option"": """ & EscapeJSONString(mihOption) & ""","
    json = json & """workbook"": """ & EscapeJSONString(wbName) & ""","
    json = json & """scenario_key"": """ & EscapeJSONString(scenarioKey) & ""","
    json = json & """source"": """ & EscapeJSONString(source) & ""","
    json = json & """units"": " & unitsJson
    json = json & "}"

    LogLearningEvent profileKey, "scenario_applied", json
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
