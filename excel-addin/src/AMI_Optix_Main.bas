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

    ' Run the optimization
    RunOptimization
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
    Dim units As Collection
    Dim utilities As Object
    Dim payload As String
    Dim response As String
    Dim result As Object

    On Error GoTo ErrorHandler

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
    If units Is Nothing Or units.Count = 0 Then
        MsgBox "No unit data found in the active workbook." & vbCrLf & vbCrLf & _
               "Please ensure your workbook has columns for:" & vbCrLf & _
               "- Unit ID (APT, UNIT, etc.)" & vbCrLf & _
               "- Bedrooms (BED, BEDS, etc.)" & vbCrLf & _
               "- Net SF (NET SF, SQFT, etc.)" & vbCrLf & _
               "- Floor (optional)" & vbCrLf & _
               "- AMI (must have NUMERIC value to be included)", _
               vbExclamation, "AMI Optix"
        GoTo Cleanup
    End If

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
    Set utilities = GetUtilitySelections()

    ' Step 4: Build API payload
    Application.StatusBar = "AMI Optix: Building request..."
    payload = BuildAPIPayload(units, utilities)

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

    ' Apply best scenario to source sheet
    ApplyBestScenario result

    ' Create scenarios comparison sheet
    CreateScenariosSheet result

    ' Done
    Application.StatusBar = False
    Application.ScreenUpdating = True

    MsgBox "Optimization complete!" & vbCrLf & vbCrLf & _
           "- Best scenario applied to your data" & vbCrLf & _
           "- All scenarios available on 'AMI Scenarios' sheet", _
           vbInformation, "AMI Optix"

    ' Navigate to scenarios sheet
    On Error Resume Next
    ActiveWorkbook.Worksheets("AMI Scenarios").Activate
    On Error GoTo 0

    Exit Sub

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
    ' Returns utility selections from stored settings or defaults
    Dim utils As Object
    Set utils = CreateObject("Scripting.Dictionary")

    ' Read from registry or use defaults (owner pays all = "na")
    utils("electricity") = GetSetting("AMI_Optix", "Utilities", "electricity", "na")
    utils("cooking") = GetSetting("AMI_Optix", "Utilities", "cooking", "na")
    utils("heat") = GetSetting("AMI_Optix", "Utilities", "heat", "na")
    utils("hot_water") = GetSetting("AMI_Optix", "Utilities", "hot_water", "na")

    Set GetUtilitySelections = utils
End Function

Public Sub SaveUtilitySelections(electricity As String, cooking As String, _
                                  heat As String, hot_water As String)
    ' Save utility selections to registry for persistence across sessions
    SaveSetting "AMI_Optix", "Utilities", "electricity", electricity
    SaveSetting "AMI_Optix", "Utilities", "cooking", cooking
    SaveSetting "AMI_Optix", "Utilities", "heat", heat
    SaveSetting "AMI_Optix", "Utilities", "hot_water", hot_water
End Sub
