Attribute VB_Name = "AMI_Optix_Automation"
Option Explicit

Public Sub RunOptimizationUAP_Agent()
    RunOptimizationAgentCore "UAP"
End Sub

Public Sub RunOptimizationMIH_Agent()
    RunOptimizationAgentCore "MIH"
End Sub

Public Sub RefreshRentTablesCache_Agent(Optional ByVal rentRollYear As Long = 0)
    Dim targetYear As Long
    Dim sourcePath As String
    Dim cacheFolder As String
    Dim fingerprint As String

    On Error GoTo ErrorHandler

    targetYear = rentRollYear
    If targetYear <= 0 Then
        targetYear = CLng(GetSetting("AMI_Optix", "RentRollYears", "SelectedYear", "2025"))
    End If

    sourcePath = ""
    cacheFolder = ""
    fingerprint = ""
    Call EnsureRentTablesCache(targetYear, True, sourcePath, cacheFolder, fingerprint)
    Exit Sub

ErrorHandler:
    Err.Raise vbObjectError + 797, "AMI_Optix_Automation.RefreshRentTablesCache_Agent", Err.Description
End Sub

Public Sub RefreshBundledRentTablesCaches_Agent()
    On Error GoTo ErrorHandler

    Call RefreshRentTablesCache_Agent(2024)
    Call RefreshRentTablesCache_Agent(2025)
    Exit Sub

ErrorHandler:
    Err.Raise vbObjectError + 798, "AMI_Optix_Automation.RefreshBundledRentTablesCaches_Agent", Err.Description
End Sub

Private Sub RunOptimizationAgentCore(program As String)
    Dim units As Collection
    Dim utilities As Object
    Dim payload As String
    Dim response As String
    Dim result As Object
    Dim programNorm As String
    Dim workbookKind As String
    Dim detectedMihOption As String
    Dim mihOption As String
    Dim mihResidentialSF As Double
    Dim mihMaxBandPercent As Long
    Dim runStart As Double

    On Error GoTo ErrorHandler

    If ActiveWorkbook Is Nothing Then
        Err.Raise vbObjectError + 780, "AMI_Optix_Automation.RunOptimizationAgentCore", "No workbook is open."
    End If

    programNorm = UCase$(Trim$(program))
    If programNorm = "" Then programNorm = "UAP"

    runStart = Timer
    DebugLog "Automation: start program=" & programNorm & ", workbook=" & ActiveWorkbook.Name, True

    detectedMihOption = ""
    workbookKind = DetectWorkbookKind(detectedMihOption)
    If programNorm = "UAP" Then
        If workbookKind = "MIH" Then
            Err.Raise vbObjectError + 781, "AMI_Optix_Automation.RunOptimizationAgentCore", _
                      "This workbook is an MIH file (" & detectedMihOption & "). Please run MIH automation instead."
        ElseIf workbookKind = "MIH_INVALID" Then
            Err.Raise vbObjectError + 782, "AMI_Optix_Automation.RunOptimizationAgentCore", _
                      "This workbook appears to be an MIH file, but the MIH option is not supported."
        End If
    ElseIf programNorm = "MIH" Then
        If workbookKind = "UAP" Then
            Err.Raise vbObjectError + 783, "AMI_Optix_Automation.RunOptimizationAgentCore", _
                      "This workbook does not look like an MIH file."
        ElseIf workbookKind = "MIH_INVALID" Then
            Err.Raise vbObjectError + 784, "AMI_Optix_Automation.RunOptimizationAgentCore", _
                      "MIH file detected, but the MIH option is not supported."
        End If
    End If

    If Not HasAPIKey() Then
        Err.Raise vbObjectError + 785, "AMI_Optix_Automation.RunOptimizationAgentCore", "API key is not configured."
    End If

    Application.StatusBar = "AMI Optix: Reading unit data..."
    Application.ScreenUpdating = False

    Dim prevSheet As Worksheet
    Dim prevEnableEvents As Boolean
    Dim dataWs As Worksheet
    Set prevSheet = ActiveSheet
    prevEnableEvents = Application.EnableEvents

    Application.EnableEvents = False
    Set dataWs = Nothing
    On Error Resume Next
    If programNorm = "MIH" Then
        Set dataWs = ActiveWorkbook.Worksheets("MIH")
        If dataWs Is Nothing Then Set dataWs = ActiveWorkbook.Worksheets("RentRoll")
        If dataWs Is Nothing Then Set dataWs = ActiveWorkbook.Worksheets("UAP")
        If dataWs Is Nothing Then Set dataWs = ActiveWorkbook.Worksheets("PROJECT WORKSHEET")
    Else
        Set dataWs = ActiveWorkbook.Worksheets("UAP")
    End If
    On Error GoTo ErrorHandler

    If Not dataWs Is Nothing Then dataWs.Activate
    Set units = ReadUnitData()
    If Not prevSheet Is Nothing Then prevSheet.Activate
    Application.EnableEvents = prevEnableEvents

    If units Is Nothing Or units.Count = 0 Then
        Err.Raise vbObjectError + 786, "AMI_Optix_Automation.RunOptimizationAgentCore", _
                  "No unit data found in the active workbook."
    End If

    If units.Count <= 3 Then
        DebugLog "Automation: continuing despite low unit count (" & units.Count & ")", True
    End If

    Application.StatusBar = "AMI Optix: Loading utility settings..."
    Set utilities = GetUtilitySelectionsForProgram(programNorm)

    If programNorm = "MIH" Then
        Dim mihErr As String
        mihErr = ""
        If Not TryReadMIHInputsQuietForAutomation(mihOption, mihResidentialSF, mihMaxBandPercent, mihErr) Then
            If Trim$(mihErr) = "" Then mihErr = "MIH inputs are missing or invalid."
            Err.Raise vbObjectError + 787, "AMI_Optix_Automation.RunOptimizationAgentCore", mihErr
        End If
    End If

    Dim profileKey As String
    profileKey = GetLearningProfileKey(programNorm, mihOption)

    Dim learningMode As String
    learningMode = LEARNING_MODE_OFF

    Dim compareBaseline As Boolean
    compareBaseline = False

    Dim premiumWeights As Object
    Set premiumWeights = Nothing

    Dim projectOverridesJson As String
    projectOverridesJson = ""

    Application.StatusBar = "AMI Optix: Building request..."
    payload = BuildAPIPayloadV2(units, utilities, programNorm, mihOption, mihResidentialSF, mihMaxBandPercent, projectOverridesJson, compareBaseline)

    Application.StatusBar = "AMI Optix: Calling optimization API..."
    response = CallOptimizeAPI(payload)
    If response = "" Then
        Err.Raise vbObjectError + 788, "AMI_Optix_Automation.RunOptimizationAgentCore", _
                  "Failed to connect to the optimization server."
    End If

    Application.StatusBar = "AMI Optix: Processing results..."
    Set result = ParseJSON(response)
    If result Is Nothing Then
        Err.Raise vbObjectError + 789, "AMI_Optix_Automation.RunOptimizationAgentCore", _
                  "Invalid response from server."
    End If

    Set g_LastScenarios = result

    On Error Resume Next
    Call LogSolverRun(profileKey, programNorm, mihOption, learningMode, compareBaseline, premiumWeights, result)
    On Error GoTo ErrorHandler

    If result.Exists("error") Then
        Err.Raise vbObjectError + 790, "AMI_Optix_Automation.RunOptimizationAgentCore", "API Error: " & CStr(result("error"))
    End If

    If result.Exists("success") Then
        If result("success") = False Then
            Dim errorMsg As String
            errorMsg = "No optimal solution found."
            If result.Exists("error") Then errorMsg = CStr(result("error"))

            Dim notes As String
            notes = ""
            If result.Exists("notes") Then
                Dim i As Long
                For i = 1 To result("notes").Count
                    notes = notes & "- " & CStr(result("notes")(i)) & vbCrLf
                Next i
            End If

            Err.Raise vbObjectError + 791, "AMI_Optix_Automation.RunOptimizationAgentCore", _
                      Trim$(errorMsg & vbCrLf & vbCrLf & "Notes from solver:" & vbCrLf & notes)
        End If
    End If

    If Not result.Exists("scenarios") Then
        Err.Raise vbObjectError + 792, "AMI_Optix_Automation.RunOptimizationAgentCore", _
                  "Invalid response: no scenarios returned."
    End If

    Dim scenariosObj As Object
    Set scenariosObj = result("scenarios")
    If scenariosObj Is Nothing Or scenariosObj.Count = 0 Then
        Err.Raise vbObjectError + 793, "AMI_Optix_Automation.RunOptimizationAgentCore", _
                  "No scenarios returned from solver."
    End If

    Application.StatusBar = "AMI Optix: Writing results..."
    ApplyBestScenario result
    CreateScenariosSheet result

    If Len(Trim$(g_AMIOptixLastScenariosSheetBuildError)) > 0 Then
        Err.Raise vbObjectError + 794, "AMI_Optix_Automation.RunOptimizationAgentCore", _
                  "Scenarios sheet build failed: " & g_AMIOptixLastScenariosSheetBuildError
    End If

    Dim wsScenarios As Worksheet
    Set wsScenarios = Nothing
    On Error Resume Next
    Set wsScenarios = ActiveWorkbook.Worksheets("AMI Scenarios")
    On Error GoTo ErrorHandler
    If wsScenarios Is Nothing Then
        Err.Raise vbObjectError + 795, "AMI_Optix_Automation.RunOptimizationAgentCore", _
                  "AMI Scenarios sheet was not created."
    End If

    wsScenarios.Visible = xlSheetVisible
    wsScenarios.Activate

Cleanup:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    Dim errDescription As String
    errDescription = Err.Description
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    DebugLog "Automation optimization failed: " & errDescription & " (elapsed " & Format$(ElapsedSeconds(runStart), "0.00") & "s)", True
    On Error GoTo 0
    Err.Raise vbObjectError + 796, "AMI_Optix_Automation.RunOptimizationAgentCore", errDescription
End Sub

Private Function TryReadMIHInputsQuietForAutomation(ByRef mihOption As String, ByRef residentialSF As Double, ByRef maxBandPercent As Long, ByRef errMsg As String) As Boolean
    On Error GoTo Fail

    TryReadMIHInputsQuietForAutomation = False
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
    If Not TryFindNetFloorAreaQuietForAutomation(wsMIH, residentialSF) Then
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

    TryReadMIHInputsQuietForAutomation = True
    Exit Function

Fail:
    errMsg = Err.Description
    TryReadMIHInputsQuietForAutomation = False
End Function

Private Function TryFindNetFloorAreaQuietForAutomation(wsMIH As Worksheet, ByRef netFloorArea As Double) As Boolean
    On Error GoTo Fail

    TryFindNetFloorAreaQuietForAutomation = False
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
            TryFindNetFloorAreaQuietForAutomation = True
            Exit Function
        End If
    Next i

    Exit Function

Fail:
    TryFindNetFloorAreaQuietForAutomation = False
End Function
