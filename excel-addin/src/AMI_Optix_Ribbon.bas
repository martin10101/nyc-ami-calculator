Attribute VB_Name = "AMI_Optix_Ribbon"
'===============================================================================
' AMI OPTIX - Ribbon Callback Module
' Handles custom ribbon button clicks and dropdown interactions
'===============================================================================
Option Explicit

' Store available rent roll sheets
Private m_RentRollSheets() As String
Private m_RentRollCount As Long
Private m_SelectedRentRoll As String

'-------------------------------------------------------------------------------
' RIBBON CALLBACKS - SOLVER GROUP
'-------------------------------------------------------------------------------

Public Sub Ribbon_RunSolver(control As IRibbonControl)
    ' Called when "Run Solver" button is clicked
    RunOptimizationForProgram "UAP"
End Sub

Public Sub Ribbon_RunSolverUAP(control As IRibbonControl)
    RunOptimizationForProgram "UAP"
End Sub

Public Sub Ribbon_RunSolverMIH(control As IRibbonControl)
    RunOptimizationForProgram "MIH"
End Sub

Public Sub Ribbon_ViewScenarios(control As IRibbonControl)
    ' Called when "View Scenarios" button is clicked
    ShowScenarioSelector
End Sub

'-------------------------------------------------------------------------------
' RIBBON CALLBACKS - RENT ROLL GROUP
'-------------------------------------------------------------------------------

Public Sub Ribbon_SelectRentRoll(control As IRibbonControl, id As String, index As Integer)
    ' Called when user selects a rent roll from dropdown
    If index >= 0 And index < m_RentRollCount Then
        m_SelectedRentRoll = m_RentRollSheets(index)

        ' Activate the selected sheet
        On Error Resume Next
        ActiveWorkbook.Worksheets(m_SelectedRentRoll).Activate
        On Error GoTo 0

        MsgBox "Selected rent roll: " & m_SelectedRentRoll, vbInformation, "AMI Optix"
    End If
End Sub

Public Sub Ribbon_GetRentRollCount(control As IRibbonControl, ByRef returnedVal)
    ' Returns the number of rent roll sheets
    ' NOTE: RibbonX passes this ByRef as a Variant; keep it untyped to avoid "Type mismatch".
    On Error GoTo Fail

    RefreshRentRollList
    returnedVal = m_RentRollCount
    Exit Sub

Fail:
    returnedVal = 0
End Sub

Public Sub Ribbon_GetRentRollLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)
    ' Returns the label for each rent roll item
    ' NOTE: RibbonX passes returnedVal ByRef as a Variant.
    On Error GoTo Fail

    If index >= 0 And index < m_RentRollCount Then
        returnedVal = m_RentRollSheets(index)
    Else
        returnedVal = ""
    End If
    Exit Sub

Fail:
    returnedVal = ""
End Sub

Public Sub Ribbon_GetRentRollID(control As IRibbonControl, index As Integer, ByRef returnedVal)
    ' Returns unique ID for each rent roll item
    ' NOTE: RibbonX passes returnedVal ByRef as a Variant.
    On Error GoTo Fail

    returnedVal = "rentroll_" & index
    Exit Sub

Fail:
    returnedVal = ""
End Sub

Public Sub Ribbon_RefreshRentRolls(control As IRibbonControl)
    ' Refresh the rent roll list
    RefreshRentRollList
    MsgBox "Found " & m_RentRollCount & " potential rent roll sheets.", vbInformation, "AMI Optix"
End Sub

Private Sub RefreshRentRollList()
    ' Scans workbook for sheets that could be rent rolls
    Dim ws As Worksheet
    Dim tempSheets() As String
    Dim sheetCount As Long
    Dim preferredNames As Variant
    Dim i As Long
    Dim isPreferred As Boolean

    If ActiveWorkbook Is Nothing Then
        m_RentRollCount = 1
        ReDim m_RentRollSheets(0 To 0)
        m_RentRollSheets(0) = "(No workbook open)"
        Exit Sub
    End If

    sheetCount = 0
    ReDim tempSheets(0 To 50)  ' Max 50 sheets

    preferredNames = Array("UAP", "PROJECT WORKSHEET", "RentRoll", "Rent Roll", "Units", "Unit Schedule", "Sheet1", "Data")

    ' First add preferred sheets in order
    For i = LBound(preferredNames) To UBound(preferredNames)
        On Error Resume Next
        Set ws = ActiveWorkbook.Worksheets(CStr(preferredNames(i)))
        On Error GoTo 0

        If Not ws Is Nothing Then
            If SheetHasUnitData(ws) Then
                tempSheets(sheetCount) = ws.Name
                sheetCount = sheetCount + 1
            End If
            Set ws = Nothing
        End If
    Next i

    ' Then add any other sheets with unit data
    For Each ws In ActiveWorkbook.Worksheets
        If Not IsSheetInArray(ws.Name, tempSheets, sheetCount) Then
            If SheetHasUnitData(ws) Then
                tempSheets(sheetCount) = ws.Name
                sheetCount = sheetCount + 1
            End If
        End If
    Next ws

    ' Store results
    m_RentRollCount = sheetCount
    If sheetCount > 0 Then
        ReDim m_RentRollSheets(0 To sheetCount - 1)
        For i = 0 To sheetCount - 1
            m_RentRollSheets(i) = tempSheets(i)
        Next i
    Else
        ReDim m_RentRollSheets(0 To 0)
        m_RentRollSheets(0) = "(No rent rolls found)"
        m_RentRollCount = 1
    End If
End Sub

Private Function SheetHasUnitData(ws As Worksheet) As Boolean
    ' Check if sheet has recognizable unit data columns
    Dim col As Long
    Dim cellValue As String
    Dim hasUnitId As Boolean
    Dim hasBedrooms As Boolean
    Dim hasNetSF As Boolean
    Dim maxCol As Long
    Dim maxRow As Long

    On Error Resume Next
    maxCol = Application.Min(20, ws.UsedRange.Columns.count)
    maxRow = Application.Min(30, ws.UsedRange.Rows.count)
    On Error GoTo 0

    If maxCol = 0 Or maxRow = 0 Then
        SheetHasUnitData = False
        Exit Function
    End If

    ' Check first 30 rows for headers
    Dim row As Long
    For row = 1 To maxRow
        hasUnitId = False
        hasBedrooms = False
        hasNetSF = False

        For col = 1 To maxCol
            cellValue = UCase(Trim(CStr(ws.Cells(row, col).Value)))

            ' Unit ID patterns
            If InStr(cellValue, "UNIT") > 0 Or InStr(cellValue, "APT") > 0 Then
                hasUnitId = True
            End If

            ' Bedrooms patterns
            If InStr(cellValue, "BED") > 0 Or cellValue = "BR" Then
                hasBedrooms = True
            End If

            ' Net SF patterns
            If InStr(cellValue, "SF") > 0 Or InStr(cellValue, "SQFT") > 0 Or InStr(cellValue, "AREA") > 0 Then
                hasNetSF = True
            End If
        Next col

        ' Need at least 2 of 3 key columns
        If (hasUnitId And hasBedrooms) Or (hasUnitId And hasNetSF) Or (hasBedrooms And hasNetSF) Then
            SheetHasUnitData = True
            Exit Function
        End If
    Next row

    SheetHasUnitData = False
End Function

Private Function IsSheetInArray(sheetName As String, arr() As String, count As Long) As Boolean
    Dim i As Long
    For i = 0 To count - 1
        If UCase(arr(i)) = UCase(sheetName) Then
            IsSheetInArray = True
            Exit Function
        End If
    Next i
    IsSheetInArray = False
End Function

'-------------------------------------------------------------------------------
' RIBBON CALLBACKS - SETTINGS GROUP
'-------------------------------------------------------------------------------

Public Sub Ribbon_OpenUtilities(control As IRibbonControl)
    ' Open utility settings dialog
    ShowUtilityForm
End Sub

Public Sub Ribbon_OpenAPISettings(control As IRibbonControl)
    ' Open API settings dialog
    ShowSettingsForm
End Sub

'-------------------------------------------------------------------------------
' RIBBON CALLBACKS - LEARNING GROUP
'-------------------------------------------------------------------------------

Public Sub Ribbon_OpenLearningSettings(control As IRibbonControl)
    ' Configure AI learning (soft preferences only) for the current program profile.
    ' Uses InputBox/MsgBox (no UserForms) to avoid type-mismatch issues.

    On Error GoTo Fail

    If ActiveWorkbook Is Nothing Then
        MsgBox "Open a workbook first.", vbExclamation, "AMI Optix"
        Exit Sub
    End If

    Dim programNorm As String
    programNorm = "UAP"

    On Error Resume Next
    Dim wsMIH As Worksheet
    Set wsMIH = ActiveWorkbook.Worksheets("MIH")
    On Error GoTo Fail
    If Not wsMIH Is Nothing Then programNorm = "MIH"

    Dim mihOption As String
    mihOption = ""

    If programNorm = "MIH" Then
        Dim residentialSF As Double
        Dim maxBandPercent As Long
        On Error Resume Next
        Call TryReadMIHInputs(mihOption, residentialSF, maxBandPercent)
        On Error GoTo Fail
    End If

    Dim profileKey As String
    profileKey = GetLearningProfileKey(programNorm, mihOption)

    Dim currentMode As String
    currentMode = GetLearningMode(profileKey)

    Dim currentCompare As Boolean
    currentCompare = GetLearningCompareBaseline(profileKey)

    Dim currentRoot As String
    currentRoot = GetLearningLogRootPath()

    Dim summary As String
    summary = "Profile: " & profileKey & vbCrLf & _
              "Program: " & programNorm & vbCrLf & _
              IIf(programNorm = "MIH", "MIH Option: " & mihOption & vbCrLf, "") & _
              vbCrLf & _
              "Learning Mode: " & currentMode & vbCrLf & _
              "Compare Baseline: " & IIf(currentCompare, "YES", "NO") & vbCrLf & _
              "Log Root: " & currentRoot

    MsgBox summary, vbInformation, "AMI Optix - Learning Settings"

    Dim newMode As String
    newMode = InputBox("Enter learning mode for " & profileKey & " (OFF / SHADOW / ON):", "AMI Optix - Learning Mode", currentMode)
    If Trim$(newMode) <> "" Then
        newMode = UCase$(Trim$(newMode))
        If newMode <> LEARNING_MODE_OFF And newMode <> LEARNING_MODE_SHADOW And newMode <> LEARNING_MODE_ON Then
            MsgBox "Invalid mode. Use OFF, SHADOW, or ON.", vbExclamation, "AMI Optix"
            Exit Sub
        End If
        Call SetLearningMode(profileKey, newMode)
        currentMode = newMode
    End If

    Dim cmpResp As VbMsgBoxResult
    cmpResp = MsgBox("Compare baseline each run for " & profileKey & "?" & vbCrLf & vbCrLf & _
                     "YES = run baseline+learned and log diff" & vbCrLf & _
                     "NO = only run the selected mode", _
                     vbYesNoCancel + vbQuestion, "AMI Optix - Compare Baseline")
    If cmpResp <> vbCancel Then
        Call SetLearningCompareBaseline(profileKey, (cmpResp = vbYes))
    End If

    Dim pathResp As VbMsgBoxResult
    pathResp = MsgBox("Change learning log folder root?" & vbCrLf & vbCrLf & currentRoot, vbYesNo + vbQuestion, "AMI Optix - Learning Logs")
    If pathResp = vbYes Then
        Dim newRoot As String
        newRoot = InputBox("Enter log root folder path (e.g. Z:\AMI_Optix_Learning or \\server\share\AMI_Optix_Learning):", _
                           "AMI Optix - Learning Log Folder", currentRoot)
        If Trim$(newRoot) <> "" Then
            Call SetLearningLogRootPath(newRoot)
        End If
    End If

    MsgBox "Saved." & vbCrLf & vbCrLf & _
           "Mode: " & GetLearningMode(profileKey) & vbCrLf & _
           "Compare Baseline: " & IIf(GetLearningCompareBaseline(profileKey), "YES", "NO") & vbCrLf & _
           "Log Root: " & GetLearningLogRootPath(), _
           vbInformation, "AMI Optix"
    Exit Sub

Fail:
    MsgBox "Learning settings failed: " & Err.Description, vbExclamation, "AMI Optix"
End Sub

Public Sub Ribbon_OpenLearningLogs(control As IRibbonControl)
    ' Opens the log folder for the current program profile in Explorer.
    On Error GoTo Fail

    If ActiveWorkbook Is Nothing Then Exit Sub

    Dim programNorm As String
    programNorm = "UAP"

    On Error Resume Next
    Dim wsMIH As Worksheet
    Set wsMIH = ActiveWorkbook.Worksheets("MIH")
    On Error GoTo Fail
    If Not wsMIH Is Nothing Then programNorm = "MIH"

    Dim mihOption As String
    mihOption = ""
    If programNorm = "MIH" Then
        Dim residentialSF As Double
        Dim maxBandPercent As Long
        On Error Resume Next
        Call TryReadMIHInputs(mihOption, residentialSF, maxBandPercent)
        On Error GoTo Fail
    End If

    Dim profileKey As String
    profileKey = GetLearningProfileKey(programNorm, mihOption)

    Dim folderPath As String
    folderPath = GetLearningProfileFolder(profileKey)
    Call EnsureFolderExists(GetLearningLogRootPath())
    Call EnsureFolderExists(folderPath)

    Shell "explorer.exe """ & folderPath & """", vbNormalFocus
    Exit Sub

Fail:
End Sub

'-------------------------------------------------------------------------------
' RIBBON CALLBACKS - HELP GROUP
'-------------------------------------------------------------------------------

Public Sub Ribbon_ShowAbout(control As IRibbonControl)
    ' Show about dialog
    MsgBox "AMI Optix Excel Add-in" & vbCrLf & vbCrLf & _
           "Version 1.0" & vbCrLf & _
           "NYC Affordable Housing AMI Optimizer" & vbCrLf & vbCrLf & _
           "Optimizes AMI band assignments for affordable housing " & _
           "projects to maximize revenue while meeting regulatory requirements." & vbCrLf & vbCrLf & _
           "API: " & API_BASE_URL, _
           vbInformation, "About AMI Optix"
End Sub

'-------------------------------------------------------------------------------
' PUBLIC ACCESSORS
'-------------------------------------------------------------------------------

Public Function GetSelectedRentRoll() As String
    GetSelectedRentRoll = m_SelectedRentRoll
End Function

Public Sub SetSelectedRentRoll(sheetName As String)
    m_SelectedRentRoll = sheetName
End Sub

'-------------------------------------------------------------------------------
' HELPER FUNCTIONS (called by ribbon callbacks)
'-------------------------------------------------------------------------------

Public Sub ShowUtilityForm()
    ' Utility configuration - 4 simple prompts with CLEAR options
    Dim electricity As String
    Dim cooking As String
    Dim heat As String
    Dim hotWater As String
    Dim response As String

    ' ELECTRICITY - simple Y/N
    response = MsgBox("Does the TENANT pay for ELECTRICITY?" & vbCrLf & vbCrLf & _
                      "Click YES if tenant pays" & vbCrLf & _
                      "Click NO if owner pays or N/A", _
                      vbYesNoCancel + vbQuestion, "AMI Optix - Electricity")
    If response = vbCancel Then Exit Sub
    electricity = IIf(response = vbYes, "tenant_pays", "na")

    ' COOKING - choose type
    response = InputBox("COOKING - What type?" & vbCrLf & vbCrLf & _
                        "E = Electric Stove (tenant pays)" & vbCrLf & _
                        "G = Gas Stove (tenant pays)" & vbCrLf & _
                        "N = Owner pays / N/A" & vbCrLf & vbCrLf & _
                        "Enter E, G, or N:", _
                        "AMI Optix - Cooking", "G")
    If response = "" Then Exit Sub
    Select Case UCase(Trim(response))
        Case "E": cooking = "electric"
        Case "G": cooking = "gas"
        Case "N": cooking = "na"
        Case Else: cooking = "gas"
    End Select

    ' HEAT - choose type
    response = InputBox("HEAT - What type does TENANT pay?" & vbCrLf & vbCrLf & _
                        "1 = Electric (ccASHP)" & vbCrLf & _
                        "2 = Electric (Other)" & vbCrLf & _
                        "3 = Gas" & vbCrLf & _
                        "4 = Oil" & vbCrLf & _
                        "N = Owner pays / N/A" & vbCrLf & vbCrLf & _
                        "Enter 1, 2, 3, 4, or N:", _
                        "AMI Optix - Heat", "3")
    If response = "" Then Exit Sub
    Select Case UCase(Trim(response))
        Case "1": heat = "electric_ccashp"
        Case "2": heat = "electric_other"
        Case "3": heat = "gas"
        Case "4": heat = "oil"
        Case "N": heat = "na"
        Case Else: heat = "gas"
    End Select

    ' HOT WATER - choose type
    response = InputBox("HOT WATER - What type does TENANT pay?" & vbCrLf & vbCrLf & _
                        "1 = Electric (Heat Pump)" & vbCrLf & _
                        "2 = Electric (Other)" & vbCrLf & _
                        "3 = Gas" & vbCrLf & _
                        "4 = Oil" & vbCrLf & _
                        "N = Owner pays / N/A" & vbCrLf & vbCrLf & _
                        "Enter 1, 2, 3, 4, or N:", _
                        "AMI Optix - Hot Water", "3")
    If response = "" Then Exit Sub
    Select Case UCase(Trim(response))
        Case "1": hotWater = "electric_heat_pump"
        Case "2": hotWater = "electric_other"
        Case "3": hotWater = "gas"
        Case "4": hotWater = "oil"
        Case "N": hotWater = "na"
        Case Else: hotWater = "gas"
    End Select

    ' Save selections
    SaveUtilitySelections electricity, cooking, heat, hotWater

    ' Show confirmation with friendly names
    MsgBox "Utility settings saved:" & vbCrLf & vbCrLf & _
           "Electricity: " & UtilityDisplayName(electricity) & vbCrLf & _
           "Cooking: " & UtilityDisplayName(cooking) & vbCrLf & _
           "Heat: " & UtilityDisplayName(heat) & vbCrLf & _
           "Hot Water: " & UtilityDisplayName(hotWater), vbInformation, "AMI Optix"
End Sub

Private Function UtilityDisplayName(value As String) As String
    ' Convert utility code to friendly display name
    Select Case value
        Case "tenant_pays": UtilityDisplayName = "Tenant Pays"
        Case "na": UtilityDisplayName = "Owner Pays / N/A"
        Case "electric": UtilityDisplayName = "Electric Stove"
        Case "gas": UtilityDisplayName = "Gas"
        Case "oil": UtilityDisplayName = "Oil"
        Case "electric_ccashp": UtilityDisplayName = "Electric (ccASHP)"
        Case "electric_other": UtilityDisplayName = "Electric (Other)"
        Case "electric_heat_pump": UtilityDisplayName = "Electric (Heat Pump)"
        Case Else: UtilityDisplayName = value
    End Select
End Function

Public Sub ShowSettingsForm()
    ' Shows API settings dialog (uses InputBox in Main module)
    OnSettingsClick Nothing
End Sub

Public Sub ShowScenarioSelector()
    ' Shows scenario selection dialog to apply different scenarios
    ' Uses InputBox-based selection (no forms needed)

    If g_LastScenarios Is Nothing Then
        MsgBox "No scenarios available." & vbCrLf & vbCrLf & _
               "Run the solver first to generate scenarios.", _
               vbInformation, "AMI Optix"
        Exit Sub
    End If

    ' Check if scenarios sheet exists
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("AMI Scenarios")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "No scenarios sheet found." & vbCrLf & vbCrLf & _
               "Run the solver first to generate scenarios.", _
               vbInformation, "AMI Optix"
        Exit Sub
    End If

    ' Show simple scenario list (InputBox-based)
    ShowScenarioList
End Sub

Private Sub ShowScenarioList()
    ' Fallback: Shows a simple list of scenarios to choose from
    Dim scenarios As Object
    Dim scenarioKey As Variant
    Dim msg As String
    Dim choice As String
    Dim i As Long
    Dim keys() As String

    If g_LastScenarios Is Nothing Then
        MsgBox "No scenarios available. Run the solver first.", vbInformation, "AMI Optix"
        Exit Sub
    End If

    Set scenarios = g_LastScenarios("scenarios")

    If scenarios.Count = 0 Then
        MsgBox "No scenarios available.", vbInformation, "AMI Optix"
        Exit Sub
    End If

    ' Build list of scenarios
    ReDim keys(1 To scenarios.Count)
    msg = "Available Scenarios:" & vbCrLf & vbCrLf
    i = 1

    For Each scenarioKey In scenarios.keys
        keys(i) = CStr(scenarioKey)
        Dim scenario As Object
        Set scenario = scenarios(scenarioKey)

        Dim waami As String
        waami = Format(scenario("waami"), "0.00%")

        msg = msg & i & ". " & FormatScenarioNameForPicker(CStr(scenarioKey)) & _
              " (WAAMI: " & waami & ")" & vbCrLf
        i = i + 1
    Next scenarioKey

    msg = msg & vbCrLf & "Enter scenario number (1-" & scenarios.Count & "):"

    choice = InputBox(msg, "Select Scenario", "1")

    If choice = "" Then Exit Sub  ' Cancelled

    Dim choiceNum As Long
    On Error Resume Next
    choiceNum = CLng(choice)
    On Error GoTo 0

    If choiceNum < 1 Or choiceNum > scenarios.Count Then
        MsgBox "Invalid selection.", vbExclamation, "AMI Optix"
        Exit Sub
    End If

    ' Apply selected scenario
    ApplyScenarioByKey keys(choiceNum)
End Sub

Private Function FormatScenarioNameForPicker(key As String) As String
    Select Case key
        Case "absolute_best"
            FormatScenarioNameForPicker = "Absolute Best"
        Case "best_3_band"
            FormatScenarioNameForPicker = "Best 3-Band"
        Case "best_2_band"
            FormatScenarioNameForPicker = "Best 2-Band"
        Case "alternative"
            FormatScenarioNameForPicker = "Alternative"
        Case "client_oriented"
            FormatScenarioNameForPicker = "Client Oriented (Max Revenue)"
        Case Else
            FormatScenarioNameForPicker = Replace(key, "_", " ")
    End Select
End Function
