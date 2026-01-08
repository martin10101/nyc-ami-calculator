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

' Global state
Public g_LastScenarios As Object  ' Stores last API response for viewing

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
    ' Show utilities configuration form
    On Error GoTo ErrorHandler

    frmUtilities.Show
    Exit Sub

ErrorHandler:
    MsgBox "Error opening utilities form: " & Err.Description, vbCritical, "AMI Optix"
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
    MsgBox "AMI Optix - NYC Affordable Housing Optimizer" & vbCrLf & vbCrLf & _
           "Version 1.0" & vbCrLf & _
           "API: " & API_BASE_URL & vbCrLf & vbCrLf & _
           "This add-in connects to the AMI optimization API." & vbCrLf & _
           "The web dashboard is also available as a backup at:" & vbCrLf & _
           API_BASE_URL, _
           vbInformation, "About AMI Optix"
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
               "- AMI (for writing results)", _
               vbExclamation, "AMI Optix"
        GoTo Cleanup
    End If

    ' Step 3: Get utility selections
    Application.StatusBar = "AMI Optix: Loading utility settings..."
    Set utilities = GetUtilitySelections()

    ' Step 4: Build API payload
    Application.StatusBar = "AMI Optix: Building request..."
    payload = BuildAPIPayload(units, utilities)

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
    Set result = ParseJSON(response)

    If result Is Nothing Then
        MsgBox "Invalid response from server.", vbCritical, "AMI Optix"
        GoTo Cleanup
    End If

    ' Store for later viewing
    Set g_LastScenarios = result

    ' Step 7: Check for scenarios
    If Not result.Exists("scenarios") Or result("scenarios").Count = 0 Then
        Dim notes As String
        notes = ""
        If result.Exists("notes") Then
            Dim i As Long
            For i = 1 To result("notes").Count
                notes = notes & "- " & result("notes")(i) & vbCrLf
            Next i
        End If
        MsgBox "No optimal scenarios found." & vbCrLf & vbCrLf & _
               "Notes from solver:" & vbCrLf & notes, _
               vbInformation, "AMI Optix"
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
