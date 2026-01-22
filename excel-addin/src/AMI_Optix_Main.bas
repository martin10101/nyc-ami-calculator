Attribute VB_Name = "AMI_Optix_Main"
'===============================================================================
' AMI OPTIX - Main Module
' Entry points for ribbon buttons and main workflow coordination
'
' NOTE: The web dashboard continues to work as a backup option.
' This Excel add-in is an additional interface to the same API.
'===============================================================================
Option Explicit

' API Configuration - Same API as website
Public Const API_BASE_URL As String = "https://nyc-ami-calculator.onrender.com"
Public Const API_TIMEOUT_MS As Long = 90000  ' 90 seconds for cold start

' API Key - stored in Windows Registry for security
' User sets this once via the Settings form
Private Const API_KEY_REGISTRY_PATH As String = "AMI_Optix"
Private Const API_KEY_REGISTRY_KEY As String = "APIKey"

' Global state
Public g_LastScenarios As Object  ' Stores last API response for viewing

'-------------------------------------------------------------------------------
' API KEY MANAGEMENT
'-------------------------------------------------------------------------------

Public Function GetAPIKey() As String
    ' Retrieves API key from Windows Registry
    GetAPIKey = GetSetting(API_KEY_REGISTRY_PATH, "Settings", API_KEY_REGISTRY_KEY, "")
End Function

Public Sub SetAPIKey(apiKey As String)
    ' Saves API key to Windows Registry
    SaveSetting API_KEY_REGISTRY_PATH, "Settings", API_KEY_REGISTRY_KEY, apiKey
End Sub

Public Function HasAPIKey() As Boolean
    ' Checks if API key is configured
    HasAPIKey = (Len(GetAPIKey()) > 0)
End Function

'-------------------------------------------------------------------------------
' RIBBON BUTTON HANDLERS
'-------------------------------------------------------------------------------

Public Sub OnOptimizeClick(control As IRibbonControl)
    ' Main "Optimize" button click handler
    On Error GoTo ErrorHandler

    ' Check if a workbook is open
    If ActiveWorkbook Is Nothing Then
        MsgBox "Please open a workbook with unit data first.", vbExclamation, "AMI Optix"
        Exit Sub
    End If

    ' Run the optimization (UAP default)
    RunOptimizationForProgram "UAP"
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "AMI Optix"
End Sub

Public Sub OnOptimizeMIHClick(control As IRibbonControl)
    ' MIH optimize handler (for the alternate Ribbon XML)
    On Error GoTo ErrorHandler

    If ActiveWorkbook Is Nothing Then
        MsgBox "Please open a MIH workbook first.", vbExclamation, "AMI Optix"
        Exit Sub
    End If

    RunOptimizationForProgram "MIH"
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "AMI Optix"
End Sub

Public Sub OnUtilitiesClick(control As IRibbonControl)
    ' Show utilities configuration using InputBox prompts
    ShowUtilityForm
End Sub

Public Sub OnViewScenariosClick(control As IRibbonControl)
    ' Navigate to scenarios sheet
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("AMI Scenarios")
    On Error GoTo ErrorHandler

    If ws Is Nothing Then
        MsgBox "No scenarios available. Run Optimize first.", vbInformation, "AMI Optix"
    Else
        ws.Activate
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "AMI Optix"
End Sub

Public Sub OnAboutClick(control As IRibbonControl)
    Dim keyStatus As String
    If HasAPIKey() Then
        keyStatus = "Configured"
    Else
        keyStatus = "NOT SET - Click Settings to configure"
    End If

    MsgBox "AMI Optix - NYC Affordable Housing Optimizer" & vbCrLf & vbCrLf & _
           "Version 1.0" & vbCrLf & _
           "API: " & API_BASE_URL & vbCrLf & _
           "API Key: " & keyStatus & vbCrLf & vbCrLf & _
           "This add-in connects to the AMI optimization API." & vbCrLf & _
           "The web dashboard is also available as a backup at:" & vbCrLf & _
           API_BASE_URL, _
           vbInformation, "About AMI Optix"
End Sub

Public Sub OnSettingsClick(control As IRibbonControl)
    ' Show settings dialog for API key configuration
    On Error GoTo ErrorHandler

    Dim currentKey As String
    Dim newKey As String
    Dim maskedKey As String

    currentKey = GetAPIKey()

    ' Mask the current key for display
    If Len(currentKey) > 8 Then
        maskedKey = Left(currentKey, 4) & "..." & Right(currentKey, 4)
    ElseIf Len(currentKey) > 0 Then
        maskedKey = "****"
    Else
        maskedKey = "(not set)"
    End If

    newKey = InputBox("Enter your API Key:" & vbCrLf & vbCrLf & _
                      "Current: " & maskedKey & vbCrLf & vbCrLf & _
                      "Leave blank to keep current key." & vbCrLf & _
                      "Enter 'CLEAR' to remove the key.", _
                      "AMI Optix - API Key Settings", "")

    If newKey = "" Then
        ' User cancelled or left blank - keep current
        Exit Sub
    ElseIf UCase(newKey) = "CLEAR" Then
        SetAPIKey ""
        MsgBox "API key has been cleared.", vbInformation, "AMI Optix"
    Else
        SetAPIKey newKey
        MsgBox "API key has been saved.", vbInformation, "AMI Optix"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "AMI Optix"
End Sub

'-------------------------------------------------------------------------------
' MAIN OPTIMIZATION WORKFLOW
'-------------------------------------------------------------------------------

Public Sub RunOptimization()
    ' Backward-compatible entrypoint (defaults to UAP)
    RunOptimizationForProgram "UAP"
End Sub

Public Sub RunOptimizationForProgram(program As String)
    Dim units As Collection
    Dim utilities As Object
    Dim payload As String
    Dim response As String
    Dim result As Object
    Dim programNorm As String
    Dim mihOption As String
    Dim mihResidentialSF As Double
    Dim mihMaxBandPercent As Long

    On Error GoTo ErrorHandler

    programNorm = UCase(Trim(program))
    If programNorm = "" Then programNorm = "UAP"

    ' Step 0: Check API key is configured
    If Not HasAPIKey() Then
        Dim setupNow As VbMsgBoxResult
        setupNow = MsgBox("API key is not configured." & vbCrLf & vbCrLf & _
                          "You need an API key to use the optimization service." & vbCrLf & _
                          "Would you like to enter it now?", _
                          vbYesNo + vbQuestion, "AMI Optix - Setup Required")
        If setupNow = vbYes Then
            OnSettingsClick Nothing
        End If
        If Not HasAPIKey() Then
            Exit Sub
        End If
    End If

    ' Step 1: Show progress
    Application.StatusBar = "AMI Optix: Reading unit data..."
    Application.ScreenUpdating = False

    ' Step 2: Read unit data from active workbook
    Set units = ReadUnitData()
    If units Is Nothing Then GoTo NoUnitsFound
    If units.Count = 0 Then GoTo NoUnitsFound

    ' DEBUG: Show unit count - warn if very few
    Debug.Print "=== UNITS FOUND ==="
    Debug.Print "Total units with valid AMI values: " & units.Count

    If units.Count <= 3 Then
        Dim continueAnyway As VbMsgBoxResult
        continueAnyway = MsgBox("WARNING: Only " & units.Count & " unit(s) found!" & vbCrLf & vbCrLf & _
                                "Units are only included if the AMI column has a" & vbCrLf & _
                                "NUMERIC value (like 60, 0.6, or 60%)." & vbCrLf & vbCrLf & _
                                "Check your AMI column - empty or text values are skipped." & vbCrLf & vbCrLf & _
                                "Continue with " & units.Count & " unit(s)?", _
                                vbYesNo + vbExclamation, "AMI Optix - Low Unit Count")
        If continueAnyway = vbNo Then GoTo Cleanup
    End If

    ' Step 3: Get utility selections
    Application.StatusBar = "AMI Optix: Loading utility settings..."
    Set utilities = GetUtilitySelectionsForProgram(programNorm)

    ' Step 3B: MIH inputs
    If programNorm = "MIH" Then
        If Not TryReadMIHInputs(mihOption, mihResidentialSF, mihMaxBandPercent) Then
            GoTo Cleanup
        End If
    End If

    ' Step 3C: AI Learning (soft preferences only)
    Dim profileKey As String
    profileKey = GetLearningProfileKey(programNorm, mihOption)

    Dim learningMode As String
    learningMode = GetLearningMode(profileKey)

    Dim compareBaseline As Boolean
    compareBaseline = False

    Dim premiumWeights As Object
    Set premiumWeights = Nothing

    Dim projectOverridesJson As String
    projectOverridesJson = ""

    If learningMode <> LEARNING_MODE_OFF Then
        If learningMode = LEARNING_MODE_SHADOW Then
            ' Shadow mode must compare baseline so we can apply baseline results while still logging diffs.
            compareBaseline = True
        Else
            compareBaseline = GetLearningCompareBaseline(profileKey)
        End If
        Set premiumWeights = ComputeLearnedPremiumWeights(profileKey)
        projectOverridesJson = BuildProjectOverridesJson(premiumWeights, "excel_learning_v1")
    End If

    ' Step 4: Build API payload
    Application.StatusBar = "AMI Optix: Building request..."
    payload = BuildAPIPayloadV2(units, utilities, programNorm, mihOption, mihResidentialSF, mihMaxBandPercent, projectOverridesJson, compareBaseline)

    ' DEBUG: Print payload being sent
    Debug.Print "=== UNITS READ FROM WORKBOOK ==="
    Debug.Print "Total units: " & units.Count
    Dim debugUnit As Object
    Dim debugIdx As Long
    For debugIdx = 1 To Application.Min(5, units.Count)  ' Show first 5 units
        Set debugUnit = units(debugIdx)
        Debug.Print "Unit " & debugIdx & ": ID=" & debugUnit("unit_id") & _
                    ", BR=" & debugUnit("bedrooms") & _
                    ", SF=" & debugUnit("net_sf")
    Next debugIdx
    Debug.Print "=== PAYLOAD BEING SENT (first 1000 chars) ==="
    Debug.Print Left(payload, 1000)
    Debug.Print "=== END PAYLOAD ==="

    ' Step 5: Call API
    Application.StatusBar = "AMI Optix: Calling optimization API (this may take up to 60 seconds)..."
    response = CallOptimizeAPI(payload)

    If response = "" Then
        MsgBox "Failed to connect to the optimization server." & vbCrLf & vbCrLf & _
               "Please check your internet connection." & vbCrLf & _
               "You can also use the web dashboard as backup:" & vbCrLf & _
               API_BASE_URL, _
               vbCritical, "AMI Optix"
        GoTo Cleanup
    End If

    ' Step 6: Parse response
    Application.StatusBar = "AMI Optix: Processing results..."

    ' DEBUG: Print raw response to Immediate window
    Debug.Print "=== RAW API RESPONSE (length: " & Len(response) & ") ==="
    Debug.Print response
    Debug.Print "=== END RAW RESPONSE ==="

    Set result = ParseJSON(response)

    If result Is Nothing Then
        MsgBox "Invalid response from server." & vbCrLf & vbCrLf & _
               "Raw response (" & Len(response) & " chars):" & vbCrLf & _
               Left(response, 500), vbCritical, "AMI Optix"
        GoTo Cleanup
    End If

    ' Store for later viewing
    Set g_LastScenarios = result

    ' Best-effort learning audit log
    On Error Resume Next
    Call LogSolverRun(profileKey, programNorm, mihOption, learningMode, compareBaseline, premiumWeights, result)
    On Error GoTo ErrorHandler

    ' Step 7: Check for success and scenarios
    ' First check if API returned an error
    If result.Exists("error") Then
        MsgBox "API Error: " & result("error"), vbExclamation, "AMI Optix"
        GoTo Cleanup
    End If

    ' Check for success flag (API returns success: false if no solution)
    If result.Exists("success") Then
        If result("success") = False Then
            Dim errorMsg As String
            errorMsg = "No optimal solution found."
            If result.Exists("error") Then
                errorMsg = result("error")
            End If

            Dim notes As String
            notes = ""
            If result.Exists("notes") Then
                Dim i As Long
                For i = 1 To result("notes").Count
                    notes = notes & "- " & result("notes")(i) & vbCrLf
                Next i
            End If

            MsgBox errorMsg & vbCrLf & vbCrLf & _
                   "Notes from solver:" & vbCrLf & notes, _
                   vbInformation, "AMI Optix"
            GoTo Cleanup
        End If
    End If

    ' Check scenarios exist
    If Not result.Exists("scenarios") Then
        MsgBox "Invalid response: no scenarios returned.", vbExclamation, "AMI Optix"
        GoTo Cleanup
    End If

    Dim scenariosObj As Object
    Set scenariosObj = result("scenarios")

    If scenariosObj.Count = 0 Then
        MsgBox "No scenarios returned from solver.", vbInformation, "AMI Optix"
        GoTo Cleanup
    End If

    ' Step 8: Write results
    Application.StatusBar = "AMI Optix: Writing results..."

    Dim appliedScenarioLabel As String
    appliedScenarioLabel = "best scenario"

    ' Apply to source sheet
    If learningMode = LEARNING_MODE_SHADOW And compareBaseline Then
        Dim appliedBaseline As Boolean
        appliedBaseline = False

        On Error Resume Next
        If result.Exists("learning") Then
            Dim learningObj As Object
            Set learningObj = result("learning")
            If Not learningObj Is Nothing Then
                If learningObj.Exists("baseline") Then
                    Dim baseObj As Object
                    Set baseObj = learningObj("baseline")
                    If Not baseObj Is Nothing Then
                        If baseObj.Exists("absolute_best") Then
                            Dim baseBest As Object
                            Set baseBest = baseObj("absolute_best")
                            If Not baseBest Is Nothing Then
                                If baseBest.Exists("canonical_assignments") Then
                                    Dim canon As Object
                                    Set canon = baseBest("canonical_assignments")
                                    If Not canon Is Nothing Then
                                        ApplyCanonicalAssignmentsToDataSheet canon, RGB(200, 220, 255)
                                        appliedBaseline = True
                                        appliedScenarioLabel = "baseline scenario (shadow mode)"
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        On Error GoTo ErrorHandler

        If Not appliedBaseline Then
            ApplyBestScenario result
        End If
    Else
        ApplyBestScenario result
    End If

    ' Create scenarios comparison sheet
    CreateScenariosSheet result
    Dim scenariosSheetOk As Boolean
    scenariosSheetOk = (Len(Trim$(g_AMIOptixLastScenariosSheetBuildError)) = 0)

    ' In Shadow mode, the scenarios sheet was built from the learned result, but the sheet data
    ' reflects baseline (by design). Refresh the Scenario Manual block to match current sheet values.
    If learningMode = LEARNING_MODE_SHADOW Then
        On Error Resume Next
        Call UpdateManualScenario(False, programNorm)
        On Error GoTo ErrorHandler
    End If

    ' Done
    Application.StatusBar = False
    Application.ScreenUpdating = True

    MsgBox "Optimization complete!" & vbCrLf & vbCrLf & _
           "- " & appliedScenarioLabel & " applied to your data" & vbCrLf & _
           IIf(scenariosSheetOk, "- All scenarios available on 'AMI Scenarios' sheet", "- Scenarios sheet failed to build (see error popup)"), _
           vbInformation, "AMI Optix"

    ' Navigate to scenarios sheet
    On Error Resume Next
    ActiveWorkbook.Worksheets("AMI Scenarios").Activate
    On Error GoTo 0

    Exit Sub

NoUnitsFound:
    MsgBox "No unit data found in the active workbook." & vbCrLf & vbCrLf & _
           "Please ensure your workbook has columns for:" & vbCrLf & _
           "- Unit ID (APT, UNIT, etc.)" & vbCrLf & _
           "- Bedrooms (BED, BEDS, etc.)" & vbCrLf & _
           "- Net SF (NET SF, SQFT, etc.)" & vbCrLf & _
           "- Floor (optional)" & vbCrLf & _
           "- AMI (must have NUMERIC value to be included)", _
           vbExclamation, "AMI Optix"
    GoTo Cleanup

Cleanup:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Optimization failed: " & Err.Description & vbCrLf & vbCrLf & _
           "Try the web dashboard as backup: " & API_BASE_URL, _
           vbCritical, "AMI Optix"
End Sub

'-------------------------------------------------------------------------------
' UTILITY SELECTIONS (stored in Windows Registry for persistence)
'-------------------------------------------------------------------------------

Private Function GetUtilitySelections() As Object
    ' Backward-compatible: defaults to stored settings
    Set GetUtilitySelections = GetUtilitySelectionsForProgram("")
End Function

Public Function GetUtilitySelectionsForProgram(program As String) As Object
    ' Returns utility selections from workbook (if present) or stored settings.
    Dim utils As Object
    Set utils = CreateObject("Scripting.Dictionary")

    Dim programNorm As String
    programNorm = UCase(Trim(program))

    Dim readFromWorkbook As Boolean
    readFromWorkbook = False

    If programNorm = "MIH" Then
        readFromWorkbook = TryReadMIHUtilities(utils)
    Else
        readFromWorkbook = TryReadUAPUtilities(utils)
    End If

    If Not readFromWorkbook Then
        utils("electricity") = GetSetting("AMI_Optix", "Utilities", "electricity", "na")
        utils("cooking") = GetSetting("AMI_Optix", "Utilities", "cooking", "na")
        utils("heat") = GetSetting("AMI_Optix", "Utilities", "heat", "na")
        utils("hot_water") = GetSetting("AMI_Optix", "Utilities", "hot_water", "na")
    Else
        ' Persist so the user sees consistent choices if workbook detection isn't available next time.
        SaveUtilitySelections CStr(utils("electricity")), CStr(utils("cooking")), CStr(utils("heat")), CStr(utils("hot_water"))
    End If

    Set GetUtilitySelectionsForProgram = utils
End Function

Private Function IsTenantPays(value As Variant) As Boolean
    Dim s As String
    s = UCase(Trim(CStr(value)))
    IsTenantPays = (s = "TENANT PAYS" Or InStr(s, "TENANT") > 0)
End Function

Private Function TryReadUAPUtilities(ByRef utils As Object) As Boolean
    On Error GoTo Fail

    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("Calculations")
    On Error GoTo Fail
    If ws Is Nothing Then GoTo Fail

    ' Electricity: P3
    utils("electricity") = IIf(IsTenantPays(ws.Range("P3").Value), "tenant_pays", "na")

    ' Cooking: Q3 electric stove, R3 gas stove
    If IsTenantPays(ws.Range("Q3").Value) Then
        utils("cooking") = "electric"
    ElseIf IsTenantPays(ws.Range("R3").Value) Then
        utils("cooking") = "gas"
    Else
        utils("cooking") = "na"
    End If

    ' Heat: S3 ccASHP, T3 electric other, U3 gas, V3 oil
    If IsTenantPays(ws.Range("S3").Value) Then
        utils("heat") = "electric_ccashp"
    ElseIf IsTenantPays(ws.Range("T3").Value) Then
        utils("heat") = "electric_other"
    ElseIf IsTenantPays(ws.Range("U3").Value) Then
        utils("heat") = "gas"
    ElseIf IsTenantPays(ws.Range("V3").Value) Then
        utils("heat") = "oil"
    Else
        utils("heat") = "na"
    End If

    ' Hot water: X3 heat pump, Y3 electric other, Z3 gas, AA3 oil
    If IsTenantPays(ws.Range("X3").Value) Then
        utils("hot_water") = "electric_heat_pump"
    ElseIf IsTenantPays(ws.Range("Y3").Value) Then
        utils("hot_water") = "electric_other"
    ElseIf IsTenantPays(ws.Range("Z3").Value) Then
        utils("hot_water") = "gas"
    ElseIf IsTenantPays(ws.Range("AA3").Value) Then
        utils("hot_water") = "oil"
    Else
        utils("hot_water") = "na"
    End If

    TryReadUAPUtilities = True
    Exit Function

Fail:
    TryReadUAPUtilities = False
End Function

Private Function TryReadMIHUtilities(ByRef utils As Object) As Boolean
    On Error GoTo Fail

    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("Rents & Utilities")
    On Error GoTo Fail
    If ws Is Nothing Then GoTo Fail

    Dim col As Long
    Dim lastCol As Long
    lastCol = ws.Cells(14, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 2 Then GoTo Fail

    ' Defaults
    utils("electricity") = "na"
    utils("cooking") = "na"
    utils("heat") = "na"
    utils("hot_water") = "na"

    For col = 1 To lastCol
        Dim header As String
        Dim optionLabel As String
        Dim selected As Variant
        header = UCase(Trim(CStr(ws.Cells(14, col).Value)))
        optionLabel = UCase(Trim(CStr(ws.Cells(15, col).Value)))
        selected = ws.Cells(16, col).Value

        If Not IsTenantPays(selected) Then
            GoTo NextCol
        End If

        If InStr(header, "APARTMENT ELECTRICITY") > 0 Then
            utils("electricity") = "tenant_pays"
        ElseIf header = "COOKING" Then
            If InStr(optionLabel, "ELECTRIC") > 0 Then
                utils("cooking") = "electric"
            ElseIf InStr(optionLabel, "GAS") > 0 Then
                utils("cooking") = "gas"
            End If
        ElseIf header = "HEAT" Then
            If InStr(optionLabel, "CCASHP") > 0 Then
                utils("heat") = "electric_ccashp"
            ElseIf InStr(optionLabel, "ELECTRIC") > 0 Then
                utils("heat") = "electric_other"
            ElseIf InStr(optionLabel, "GAS") > 0 Then
                utils("heat") = "gas"
            ElseIf InStr(optionLabel, "OIL") > 0 Then
                utils("heat") = "oil"
            End If
        ElseIf InStr(header, "HOT WATER") > 0 Then
            If InStr(optionLabel, "HEAT PUMP") > 0 Then
                utils("hot_water") = "electric_heat_pump"
            ElseIf InStr(optionLabel, "ELECTRIC") > 0 Then
                utils("hot_water") = "electric_other"
            ElseIf InStr(optionLabel, "GAS") > 0 Then
                utils("hot_water") = "gas"
            ElseIf InStr(optionLabel, "OIL") > 0 Then
                utils("hot_water") = "oil"
            End If
        End If

NextCol:
    Next col

    TryReadMIHUtilities = True
    Exit Function

Fail:
    TryReadMIHUtilities = False
End Function

Public Function TryReadMIHInputs(ByRef mihOption As String, ByRef residentialSF As Double, ByRef maxBandPercent As Long) As Boolean
    On Error GoTo Fail

    Dim wsMIH As Worksheet
    Dim wsProg As Worksheet
    Set wsMIH = Nothing
    Set wsProg = Nothing

    On Error Resume Next
    Set wsMIH = ActiveWorkbook.Worksheets("MIH")
    Set wsProg = ActiveWorkbook.Worksheets("Prog")
    On Error GoTo Fail

    If wsMIH Is Nothing Then
        MsgBox "MIH run requires a sheet named 'MIH'.", vbExclamation, "AMI Optix"
        TryReadMIHInputs = False
        Exit Function
    End If
    If wsProg Is Nothing Then
        MsgBox "MIH run requires a sheet named 'Prog' (for OptionSelected).", vbExclamation, "AMI Optix"
        TryReadMIHInputs = False
        Exit Function
    End If

    Dim v As Variant
    v = wsMIH.Range("J21").Value
    If Not IsNumeric(v) Or CDbl(v) <= 0 Then
        MsgBox "MIH residential SF is missing." & vbCrLf & vbCrLf & _
               "Expected a numeric value in MIH!J21.", vbExclamation, "AMI Optix"
        TryReadMIHInputs = False
        Exit Function
    End If
    residentialSF = CDbl(v)

    mihOption = Trim(CStr(wsProg.Range("K4").Value))
    If mihOption = "" Then
        MsgBox "MIH option is missing." & vbCrLf & vbCrLf & _
               "Expected 'Option 1' or 'Option 4' in Prog!K4.", vbExclamation, "AMI Optix"
        TryReadMIHInputs = False
        Exit Function
    End If

    Dim capFactor As Variant
    capFactor = wsProg.Range("I4").Value
    If IsNumeric(capFactor) Then
        maxBandPercent = CLng(CDbl(capFactor) * 100)
    Else
        maxBandPercent = 135
    End If

    TryReadMIHInputs = True
    Exit Function

Fail:
    TryReadMIHInputs = False
End Function

Public Sub SaveUtilitySelections(electricity As String, cooking As String, _
                                  heat As String, hot_water As String)
    ' Save utility selections to registry for persistence across sessions
    SaveSetting "AMI_Optix", "Utilities", "electricity", electricity
    SaveSetting "AMI_Optix", "Utilities", "cooking", cooking
    SaveSetting "AMI_Optix", "Utilities", "heat", heat
    SaveSetting "AMI_Optix", "Utilities", "hot_water", hot_water
End Sub
