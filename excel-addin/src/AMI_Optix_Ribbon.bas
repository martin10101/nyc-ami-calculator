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
' RIBBON CALLBACKS - LEARNING / LOGS GROUP
'-------------------------------------------------------------------------------

Public Sub Ribbon_RecordScenarioChoice(control As IRibbonControl)
    ' Records which scenario the client chose (no API calls).
    ' Appends a single entry to the shared JSONL run log.

    On Error GoTo Fail

    If g_LastScenarios Is Nothing Then
        MsgBox "No scenarios available." & vbCrLf & vbCrLf & _
               "Run the solver first to generate scenarios.", _
               vbInformation, "AMI Optix"
        Exit Sub
    End If

    Dim scenarios As Object
    On Error Resume Next
    If g_LastScenarios.Exists("scenarios") Then Set scenarios = g_LastScenarios("scenarios")
    On Error GoTo Fail

    If scenarios Is Nothing Then
        MsgBox "No scenarios available." & vbCrLf & vbCrLf & _
               "Run the solver first to generate scenarios.", _
               vbInformation, "AMI Optix"
        Exit Sub
    End If

    If scenarios.Count = 0 Then
        MsgBox "No scenarios available.", vbInformation, "AMI Optix"
        Exit Sub
    End If

    Dim programNorm As String
    Dim mihOption As String
    programNorm = "UAP"
    mihOption = ""

    On Error Resume Next
    If g_LastScenarios.Exists("project_summary") Then
        Dim ps As Object
        Set ps = g_LastScenarios("project_summary")
        If Not ps Is Nothing Then
            If ps.Exists("program") Then programNorm = CStr(ps("program"))
            If ps.Exists("mih_option") Then mihOption = CStr(ps("mih_option"))
        End If
    End If
    On Error GoTo Fail

    Dim profileKey As String
    profileKey = GetLearningProfileKey(programNorm, mihOption)

    ' Build an ordered list: strict first, then edge.
    Dim strictKeys As Collection
    Dim edgeKeys As Collection
    Set strictKeys = New Collection
    Set edgeKeys = New Collection

    Dim scenarioKey As Variant
    For Each scenarioKey In scenarios.keys
        Dim tier As String
        tier = ""

        On Error Resume Next
        Dim s As Object
        Set s = scenarios(scenarioKey)
        If Not s Is Nothing Then
            If s.Exists("tier") Then tier = CStr(s("tier"))
        End If
        On Error GoTo Fail

        If LCase$(Trim$(tier)) = "edge" Then
            edgeKeys.Add CStr(scenarioKey)
        Else
            strictKeys.Add CStr(scenarioKey)
        End If
    Next scenarioKey

    Dim total As Long
    total = strictKeys.Count + edgeKeys.Count

    Dim keys() As String
    ReDim keys(1 To total)

    Dim msg As String
    msg = "Available Scenarios:" & vbCrLf & vbCrLf

    Dim i As Long
    i = 1

    Dim k As Variant
    For Each k In strictKeys
        keys(i) = CStr(k)
        msg = msg & ScenarioPickerLine(i, keys(i), scenarios(keys(i))) & vbCrLf
        i = i + 1
    Next k

    For Each k In edgeKeys
        keys(i) = CStr(k)
        msg = msg & ScenarioPickerLine(i, keys(i), scenarios(keys(i))) & vbCrLf
        i = i + 1
    Next k

    msg = msg & vbCrLf & "Enter chosen scenario number (1-" & total & "):"

    Dim choice As String
    choice = InputBox(msg, "Record Chosen Scenario", "1")
    If choice = "" Then Exit Sub

    Dim choiceNum As Long
    On Error Resume Next
    choiceNum = CLng(choice)
    On Error GoTo Fail

    If choiceNum < 1 Or choiceNum > total Then
        MsgBox "Invalid selection.", vbExclamation, "AMI Optix"
        Exit Sub
    End If

    Dim chosenKey As String
    chosenKey = keys(choiceNum)

    Dim chosenScenario As Object
    Set chosenScenario = scenarios(chosenKey)

    ' Ask for a short explanation of why this scenario was chosen.
    Dim wsScenarios As Worksheet
    Set wsScenarios = Nothing
    On Error Resume Next
    Set wsScenarios = ActiveWorkbook.Worksheets("AMI Scenarios")
    On Error GoTo Fail

    Dim existingReason As String
    existingReason = ""
    If Not wsScenarios Is Nothing Then
        existingReason = ReadFinalChoiceReason(wsScenarios)
    End If

    Dim reasonVar As Variant
    reasonVar = Application.InputBox( _
        "Why did you choose this scenario? (optional)" & vbCrLf & vbCrLf & _
        "Example: best rent roll / closest to 60% / client preference / unit mix / compliance tradeoff.", _
        "AMI Optix - Choice Reason", _
        existingReason, _
        Type:=2 _
    )
    If reasonVar = False Then Exit Sub ' Cancel

    Dim choiceReason As String
    choiceReason = CStr(reasonVar)

    ' Write a visible "Final Selection" box to the scenarios sheet (so it can be shared with the client).
    If Not wsScenarios Is Nothing Then
        WriteFinalChoiceBox wsScenarios, chosenKey, FormatScenarioNameForPicker(chosenKey), choiceReason
    End If

    Call LogScenarioChoiceToRunLog(profileKey, programNorm, mihOption, choiceNum, chosenKey, chosenScenario, choiceReason)

    MsgBox "Recorded choice:" & vbCrLf & _
           "Scenario: " & FormatScenarioNameForPicker(chosenKey) & vbCrLf & _
           "Log file: " & GetRunLogFilePath(), _
           vbInformation, "AMI Optix"
    Exit Sub

Fail:
    MsgBox "Could not record choice: " & Err.Description, vbExclamation, "AMI Optix"
End Sub

Private Function ReadFinalChoiceReason(ws As Worksheet) As String
    On Error GoTo Fail
    If ws Is Nothing Then Exit Function

    Dim c As Range
    Set c = ws.Range("P5")
    ReadFinalChoiceReason = CStr(c.Value)
    Exit Function
Fail:
    ReadFinalChoiceReason = ""
End Function

Private Sub WriteFinalChoiceBox(ws As Worksheet, scenarioKey As String, scenarioLabel As String, choiceReason As String)
    On Error GoTo Fail
    If ws Is Nothing Then Exit Sub

    ' Place the final choice box to the right of the main scenario tables so it doesn't get cleared.
    ' (Manual/scenarios use columns A-M; this uses O-U.)
    Dim header As Range
    Set header = ws.Range("O1:U1")
    On Error Resume Next
    header.UnMerge
    On Error GoTo Fail
    header.Merge
    header.Value = "FINAL SELECTION (CLIENT)"
    header.Font.Bold = True
    header.Font.Size = 12
    header.Interior.Color = RGB(255, 242, 204) ' light yellow
    header.HorizontalAlignment = xlCenter

    ws.Range("O2").Value = "Selected Scenario:"
    ws.Range("O3").Value = "Scenario Key:"
    ws.Range("O4").Value = "Selected On:"
    ws.Range("O5").Value = "Why Chosen (Notes):"
    ws.Range("O10").Value = "Run Log File:"

    ws.Range("O2:O10").Font.Bold = True

    Dim v1 As Range
    Set v1 = ws.Range("P2:U2")
    On Error Resume Next
    v1.UnMerge
    On Error GoTo Fail
    v1.Merge
    v1.Value = scenarioLabel

    Dim v2 As Range
    Set v2 = ws.Range("P3:U3")
    On Error Resume Next
    v2.UnMerge
    On Error GoTo Fail
    v2.Merge
    v2.Value = scenarioKey

    Dim v3 As Range
    Set v3 = ws.Range("P4:U4")
    On Error Resume Next
    v3.UnMerge
    On Error GoTo Fail
    v3.Merge
    v3.Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")

    Dim notes As Range
    Set notes = ws.Range("P5:U9")
    On Error Resume Next
    notes.UnMerge
    On Error GoTo Fail
    notes.Merge
    notes.Value = choiceReason
    notes.WrapText = True
    notes.VerticalAlignment = xlTop
    notes.RowHeight = 70

    Dim logCell As Range
    Set logCell = ws.Range("P10:U10")
    On Error Resume Next
    logCell.UnMerge
    On Error GoTo Fail
    logCell.Merge
    logCell.Value = GetRunLogFilePath()
    logCell.WrapText = True

    ws.Range("O1:U10").Borders.LineStyle = xlContinuous
    ws.Columns("O:U").ColumnWidth = 16
    Exit Sub

Fail:
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
    ' Configure local logging (no API learning).
    ' Uses InputBox/MsgBox (no UserForms).

    On Error GoTo Fail

    If ActiveWorkbook Is Nothing Then
        MsgBox "Open a workbook first.", vbExclamation, "AMI Optix"
        Exit Sub
    End If

    Dim currentRoot As String
    currentRoot = GetLearningLogRootPath()

    Dim summary As String
    summary = "Logging is local-only (no API learning)." & vbCrLf & vbCrLf & _
              "Log Root: " & currentRoot & vbCrLf & _
              "Run Log File: " & GetRunLogFilePath()

    Dim pathResp As VbMsgBoxResult
    pathResp = MsgBox(summary & vbCrLf & vbCrLf & "Change log folder root?", vbYesNo + vbQuestion, "AMI Optix - Log Settings")
    If pathResp = vbYes Then
        Dim newRoot As String
        newRoot = InputBox("Enter log root folder path (e.g. Z:\AMI_Optix_Learning or \\server\share\AMI_Optix_Learning):", _
                           "AMI Optix - Log Folder", currentRoot)
        If Trim$(newRoot) <> "" Then
            Call SetLearningLogRootPath(newRoot)
        End If
    End If

    MsgBox "Saved." & vbCrLf & vbCrLf & _
           "Log Root: " & GetLearningLogRootPath() & vbCrLf & _
           "Run Log File: " & GetRunLogFilePath(), _
           vbInformation, "AMI Optix"
    Exit Sub

Fail:
    MsgBox "Log settings failed: " & Err.Description, vbExclamation, "AMI Optix"
End Sub

Public Sub Ribbon_OpenLearningLogs(control As IRibbonControl)
    ' Opens the log folder in Explorer.
    On Error GoTo Fail

    If ActiveWorkbook Is Nothing Then Exit Sub

    Dim folderPath As String
    folderPath = GetLearningLogRootPath()
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

    ' Build list of scenarios (strict first, then edge)
    Dim strictKeys As Collection
    Dim edgeKeys As Collection
    Set strictKeys = New Collection
    Set edgeKeys = New Collection

    For Each scenarioKey In scenarios.keys
        Dim tier As String
        tier = ""
        On Error Resume Next
        Dim s As Object
        Set s = scenarios(scenarioKey)
        If Not s Is Nothing Then
            If s.Exists("tier") Then tier = CStr(s("tier"))
        End If
        On Error GoTo 0

        If LCase$(Trim$(tier)) = "edge" Then
            edgeKeys.Add CStr(scenarioKey)
        Else
            strictKeys.Add CStr(scenarioKey)
        End If
    Next scenarioKey

    ReDim keys(1 To scenarios.Count)
    msg = "Available Scenarios:" & vbCrLf & vbCrLf
    i = 1

    Dim k As Variant
    For Each k In strictKeys
        keys(i) = CStr(k)
        msg = msg & ScenarioPickerLine(i, keys(i), scenarios(keys(i))) & vbCrLf
        i = i + 1
    Next k

    For Each k In edgeKeys
        keys(i) = CStr(k)
        msg = msg & ScenarioPickerLine(i, keys(i), scenarios(keys(i))) & vbCrLf
        i = i + 1
    Next k

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
        Case "max_revenue"
            FormatScenarioNameForPicker = "Max Revenue (Rent-Max)"
        Case "best_3_band"
            FormatScenarioNameForPicker = "Best 3-Band"
        Case "best_2_band"
            FormatScenarioNameForPicker = "Best 2-Band"
        Case "alternative"
            FormatScenarioNameForPicker = "Alternative"
        Case "client_oriented"
            FormatScenarioNameForPicker = "Client Oriented (Max Revenue)"
        Case Else
            If Left$(key, 15) = "edge_max_share_" Then
                FormatScenarioNameForPicker = "Edge: 40% Max Share " & Mid$(key, 16) & "%"
            ElseIf Left$(key, 15) = "edge_min_share_" Then
                Dim v As Long
                On Error Resume Next
                v = CLng(Mid$(key, 16))
                On Error GoTo 0
                If v > 0 Then
                    FormatScenarioNameForPicker = "Edge: 40% Min Share " & Format$(v / 10#, "0.0") & "%"
                Else
                    FormatScenarioNameForPicker = Replace(key, "_", " ")
                End If
            ElseIf Left$(key, 17) = "edge_waami_floor_" Then
                Dim w As Long
                On Error Resume Next
                w = CLng(Mid$(key, 18))
                On Error GoTo 0
                If w > 0 Then
                    FormatScenarioNameForPicker = "Edge: WAAMI Floor " & Format$(w / 10#, "0.0") & "%"
                Else
                    FormatScenarioNameForPicker = Replace(key, "_", " ")
                End If
            Else
                FormatScenarioNameForPicker = Replace(key, "_", " ")
            End If
    End Select
End Function

Private Function ScenarioPickerLine(index As Long, scenarioKey As String, scenario As Object) As String
    Dim waami As String
    waami = "n/a"

    On Error Resume Next
    If Not scenario Is Nothing Then
        If scenario.Exists("waami") Then waami = Format(CDbl(scenario("waami")), "0.00%")
    End If
    On Error GoTo 0

    Dim tierLabel As String
    tierLabel = "STRICT"
    On Error Resume Next
    If Not scenario Is Nothing Then
        If scenario.Exists("tier") Then
            If LCase$(Trim$(CStr(scenario("tier")))) = "edge" Then tierLabel = "EDGE"
        End If
    End If
    On Error GoTo 0

    ScenarioPickerLine = index & ". " & FormatScenarioNameForPicker(CStr(scenarioKey)) & _
                         " [" & tierLabel & "] (WAAMI: " & waami & ")"
End Function
