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
    RunOptimization
End Sub

Public Sub Ribbon_ViewScenarios(control As IRibbonControl)
    ' Called when "View Scenarios" button is clicked
    ShowScenarioSelector
End Sub

'-------------------------------------------------------------------------------
' RIBBON CALLBACKS - RENT ROLL GROUP
'-------------------------------------------------------------------------------

Public Sub Ribbon_SelectRentRoll(control As IRibbonControl, id As String, index As Long)
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

Public Sub Ribbon_GetRentRollCount(control As IRibbonControl, ByRef count As Long)
    ' Returns the number of rent roll sheets
    RefreshRentRollList
    count = m_RentRollCount
End Sub

Public Sub Ribbon_GetRentRollLabel(control As IRibbonControl, index As Long, ByRef label As String)
    ' Returns the label for each rent roll item
    If index >= 0 And index < m_RentRollCount Then
        label = m_RentRollSheets(index)
    Else
        label = ""
    End If
End Sub

Public Sub Ribbon_GetRentRollID(control As IRibbonControl, index As Long, ByRef id As String)
    ' Returns unique ID for each rent roll item
    id = "rentroll_" & index
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
    ' Simple utility configuration using InputBox (bypasses broken form)
    Dim msg As String
    Dim choice As String

    msg = "Utility Settings:" & vbCrLf & vbCrLf & _
          "For most NYC affordable housing, select:" & vbCrLf & _
          "  - Electricity: Tenant pays" & vbCrLf & _
          "  - Cooking: Gas" & vbCrLf & _
          "  - Heat: Gas" & vbCrLf & _
          "  - Hot Water: Gas" & vbCrLf & vbCrLf & _
          "Use default settings? (Yes/No)"

    choice = InputBox(msg, "AMI Optix - Utilities", "Yes")

    If UCase(Trim(choice)) = "YES" Or choice = "" Then
        SaveUtilitySelections "tenant_pays", "gas", "gas", "gas"
        MsgBox "Default utility settings saved.", vbInformation, "AMI Optix"
    Else
        MsgBox "To customize utilities, edit the registry or contact support.", vbInformation, "AMI Optix"
    End If
End Sub

Public Sub ShowSettingsForm()
    ' Shows API settings dialog (uses InputBox in Main module)
    OnSettingsClick Nothing
End Sub

Public Sub ShowScenarioSelector()
    ' Shows scenario selection dialog to apply different scenarios
    On Error GoTo ErrorHandler

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
    On Error GoTo ErrorHandler

    If ws Is Nothing Then
        MsgBox "No scenarios sheet found." & vbCrLf & vbCrLf & _
               "Run the solver first to generate scenarios.", _
               vbInformation, "AMI Optix"
        Exit Sub
    End If

    ' Show scenario picker
    frmScenarioPicker.Show
    Exit Sub

ErrorHandler:
    ' If form doesn't exist, show simple list
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
