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

' Best-effort: keep the AMI Optix ribbon tab selected after ribbon actions.
Private m_RibbonUI As IRibbonUI
Private Const AMI_OPTIX_TAB_ID As String = "tabAMIOptix"

' Rent roll year selector (server-side rent calculator)
Private Const AMI_OPTIX_REGISTRY_PATH As String = "AMI_Optix"
Private Const RENTROLL_YEAR_MIN As Long = 2022
Private Const RENTROLL_YEAR_MAX As Long = 2026
Private Const RENTROLL_YEAR_DEFAULT As Long = 2025
Private Const RENTROLL_YEAR_REG_SECTION As String = "RentRollYears"
Private Const RENTROLL_YEAR_REG_KEY_SELECTED As String = "SelectedYear"
Private Const RENTROLL_YEAR_REG_KEY_REMOTE_PREFIX As String = "RemoteFilename_"

Private m_SelectedRentRollYear As Long
Private m_RentRollYearInitialized As Boolean

Public Sub Ribbon_OnLoad(ribbon As IRibbonUI)
    Set m_RibbonUI = ribbon
End Sub

Private Sub EnsureAMIOptixTabActive()
    On Error Resume Next
    If Not m_RibbonUI Is Nothing Then
        m_RibbonUI.ActivateTab AMI_OPTIX_TAB_ID
    End If
    On Error GoTo 0
End Sub

Private Sub InvalidateRibbonControl(controlId As String)
    On Error Resume Next
    If Not m_RibbonUI Is Nothing Then
        m_RibbonUI.InvalidateControl controlId
    End If
    On Error GoTo 0
End Sub

Private Function GetDictString(d As Object, key As String, Optional defaultValue As String = "") As String
    On Error GoTo SafeExit
    GetDictString = defaultValue
    If d Is Nothing Then Exit Function
    Dim k As String
    k = LCase$(Trim$(CStr(key)))
    If k = "" Then Exit Function
    If d.Exists(k) Then GetDictString = CStr(d(k))
    Exit Function
SafeExit:
    GetDictString = defaultValue
End Function

Private Function InferSourceLabelFromPath(path As String) As String
    Dim p As String
    p = UCase$(Trim$(CStr(path)))
    If p = "" Then Exit Function
    If Left$(p, 3) = "Z:\" Then
        InferSourceLabelFromPath = "Z:"
    ElseIf InStr(1, p, "\AMI_OPTIX\RENTROLLYEARS\", vbTextCompare) > 0 Then
        InferSourceLabelFromPath = "AppData"
    Else
        InferSourceLabelFromPath = ""
    End If
End Function

'-------------------------------------------------------------------------------
' RIBBON CALLBACKS - SOLVER GROUP
'-------------------------------------------------------------------------------

Public Sub Ribbon_RunSolver(control As IRibbonControl)
    ' Called when "Run Solver" button is clicked
    RunOptimizationForProgram "UAP"
    EnsureAMIOptixTabActive
End Sub

Public Sub Ribbon_RunSolverUAP(control As IRibbonControl)
    ' Pre-flight reminder before running UAP: confirm Utilities filled in.
    ' Pure reminder - does not auto-detect; the client confirms manually
    ' so they get the prompt every time. Mirrors the MIH preflight (Fix A)
    ' but only checks Utilities (UAP has no Option 1/4 selection).
    Dim msg As String
    msg = "Before running UAP, please confirm:" & vbCrLf & vbCrLf & _
          "  [ ]  Utilities are filled in (Settings > Utilities)" & vbCrLf & vbCrLf & _
          "Click YES if done - UAP will run." & vbCrLf & _
          "Click NO to cancel and complete the missing item."
    If MsgBox(msg, vbYesNo + vbInformation, "Run UAP - Pre-flight") = vbNo Then
        MsgBox "Please fill in Utilities, then click Run UAP again.", _
               vbInformation, "Run UAP cancelled"
        EnsureAMIOptixTabActive
        Exit Sub
    End If

    RunOptimizationForProgram "UAP"
    EnsureAMIOptixTabActive
End Sub

Public Sub Ribbon_RunSolverMIH(control As IRibbonControl)
    ' Pre-flight reminder before running MIH: confirm Option 1/4 selected
    ' and Utilities filled in. Pure reminder - does not auto-detect; the
    ' client confirms manually so they get the prompt every time.
    Dim msg As String
    msg = "Before running MIH, please confirm:" & vbCrLf & vbCrLf & _
          "  [ ]  Option 1 or Option 4 is selected on the MIH sheet" & vbCrLf & _
          "  [ ]  Utilities are filled in (Settings > Utilities)" & vbCrLf & vbCrLf & _
          "Click YES if both are done - MIH will run." & vbCrLf & _
          "Click NO to cancel and complete the missing item(s)."
    If MsgBox(msg, vbYesNo + vbInformation, "Run MIH - Pre-flight") = vbNo Then
        MsgBox "Please select Option 1 or 4 and fill in Utilities, then click Run MIH again.", _
               vbInformation, "Run MIH cancelled"
        EnsureAMIOptixTabActive
        Exit Sub
    End If

    RunOptimizationForProgram "MIH"
    EnsureAMIOptixTabActive
End Sub

Public Sub Ribbon_ViewScenarios(control As IRibbonControl)
    ' Called when "View Scenarios" button is clicked
    ShowScenarioSelector
    EnsureAMIOptixTabActive
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
        EnsureAMIOptixTabActive
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
        EnsureAMIOptixTabActive
        Exit Sub
    End If

    If scenarios.Count = 0 Then
        MsgBox "No scenarios available.", vbInformation, "AMI Optix"
        EnsureAMIOptixTabActive
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

    ' Determine which scenario is currently applied by comparing the live sheet assignments
    ' to the last solver scenarios. If none match exactly, we record "Scenario Manual (custom)".
    Dim liveUnits As Collection
    Set liveUnits = ReadCurrentProgramUnits(programNorm)
    If liveUnits Is Nothing Or liveUnits.Count = 0 Then
        MsgBox "Could not read units from the workbook.", vbExclamation, "AMI Optix"
        EnsureAMIOptixTabActive
        Exit Sub
    End If

    Dim chosenKey As String
    Dim chosenScenario As Object
    Dim isExactMatch As Boolean
    isExactMatch = False

    chosenKey = FindMatchingScenarioKeyForUnits(liveUnits, scenarios)
    If chosenKey <> "" Then
        isExactMatch = True
        Set chosenScenario = scenarios(chosenKey)
    Else
        chosenKey = "scenario_manual"
        Set chosenScenario = BuildScenarioFromUnits(liveUnits)
    End If

    Dim chosenLabel As String
    If isExactMatch Then
        chosenLabel = FormatScenarioNameForPicker(chosenKey)
    Else
        chosenLabel = "Scenario Manual (custom)"
    End If

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
    If reasonVar = False Then
        EnsureAMIOptixTabActive
        Exit Sub ' Cancel
    End If

    Dim choiceReason As String
    choiceReason = CStr(reasonVar)

    ' Write a visible "Final Selection" box to the scenarios sheet (so it can be shared with the client).
    If Not wsScenarios Is Nothing Then
        WriteFinalChoiceBox wsScenarios, chosenKey, chosenLabel, choiceReason
    End If

    Dim choiceNum As Long
    choiceNum = IIf(isExactMatch, FindScenarioIndexForKey(chosenKey, scenarios), 0)

    Call LogScenarioChoiceToRunLog(profileKey, programNorm, mihOption, choiceNum, chosenKey, chosenScenario, choiceReason)

    Dim logPath As String
    logPath = GetRunLogFilePath()

    Dim logFileName As String
    logFileName = Dir$(logPath)

    If Trim$(logFileName) <> "" Then
        MsgBox "Recorded choice:" & vbCrLf & _
               "Scenario: " & chosenLabel & vbCrLf & _
               "Log file: " & logPath, _
               vbInformation, "AMI Optix"
    Else
        MsgBox "Recorded choice in the workbook, but could not write the run log file." & vbCrLf & vbCrLf & _
               "Scenario: " & chosenLabel & vbCrLf & _
               "Log file (expected): " & logPath & vbCrLf & _
               "Error: " & GetLastRunLogError() & vbCrLf & vbCrLf & _
               "Try: Settings → Log Settings → set Log Root to a local folder like C:\Temp\AMI_Optix_Learning, then click Record Choice again.", _
               vbExclamation, "AMI Optix"
    End If
    EnsureAMIOptixTabActive
    Exit Sub

Fail:
    MsgBox "Could not record choice: " & Err.Description, vbExclamation, "AMI Optix"
    EnsureAMIOptixTabActive
End Sub

'-------------------------------------------------------------------------------
' RIBBON CALLBACKS - MANUAL GROUP
'-------------------------------------------------------------------------------

Public Sub Ribbon_GetLiveSync(control As IRibbonControl, ByRef returnedVal)
    ' NOTE: RibbonX passes returnedVal ByRef as a Variant; keep it untyped to avoid "Type mismatch".
    On Error GoTo Fail
    Static didLog As Boolean
    If Not didLog Then
        DebugLog "Ribbon_GetLiveSync: first call", True
        didLog = True
    End If
    returnedVal = CBool(GetLiveSyncEnabled())
    Exit Sub

Fail:
    DebugLogError "Ribbon_GetLiveSync"
    returnedVal = False
End Sub

Public Sub Ribbon_ToggleLiveSync(control As IRibbonControl, pressed As Boolean)
    ' Toggle Live Sync ON/OFF.
    ' When OFF: clear Scenario Manual block + clear the program AMI column so the user can type custom values.
    On Error GoTo Fail

    Dim programNorm As String
    programNorm = DetectProgramFromWorkbook()

    Call SetLiveSyncEnabled(pressed)

    If Not pressed Then
        Dim prevEnableEvents As Boolean
        Dim prevScreenUpdating As Boolean
        prevEnableEvents = Application.EnableEvents
        prevScreenUpdating = Application.ScreenUpdating

        Application.EnableEvents = False
        Application.ScreenUpdating = False
        g_AMIOptixSuppressEvents = True

        Call ClearScenarioManualBlock
        Call ClearProgramAmiColumn(programNorm)

        g_AMIOptixSuppressEvents = False
        Application.ScreenUpdating = prevScreenUpdating
        Application.EnableEvents = prevEnableEvents

        MsgBox "Live Sync is now OFF." & vbCrLf & vbCrLf & _
               "- Scenario Manual was cleared." & vbCrLf & _
               "- The program AMI column was cleared so you can type custom values." & vbCrLf & vbCrLf & _
               "When you're ready, click AMI Optix → Manual Calculate to compute rents/band mix.", _
               vbInformation, "AMI Optix"
    Else
        MsgBox "Live Sync is now ON." & vbCrLf & vbCrLf & _
               "Edits in the program AMI column will refresh Scenario Manual automatically.", _
               vbInformation, "AMI Optix"
    End If
    EnsureAMIOptixTabActive
    Exit Sub

Fail:
    MsgBox "Could not toggle Live Sync: " & Err.Description, vbExclamation, "AMI Optix"
    EnsureAMIOptixTabActive
End Sub

Public Sub Ribbon_ManualCalculate(control As IRibbonControl)
    ' Compute Scenario Manual from the current sheet inputs (even if non-compliant).
    On Error GoTo Fail

    Dim programNorm As String
    programNorm = DetectProgramFromWorkbook()

    If Not ManualCalculateScenario(programNorm) Then
        EnsureAMIOptixTabActive
        Exit Sub
    End If
    EnsureAMIOptixTabActive
    Exit Sub

Fail:
    MsgBox "Manual Calculate failed: " & Err.Description, vbExclamation, "AMI Optix"
    EnsureAMIOptixTabActive
End Sub

Private Function ReadCurrentProgramUnits(programNorm As String) As Collection
    ' Reads the live unit table for the program so we can detect the currently applied scenario.
    On Error GoTo Fail

    If ActiveWorkbook Is Nothing Then Exit Function

    Dim prevSheet As Worksheet
    Set prevSheet = ActiveSheet

    Dim ws As Worksheet
    Set ws = Nothing

    On Error Resume Next
    If UCase$(Trim$(programNorm)) = "MIH" Then
        ' Prefer MIH sheet first (per client request), fallback only if needed.
        Set ws = ActiveWorkbook.Worksheets("MIH")
        If ws Is Nothing Then Set ws = ActiveWorkbook.Worksheets("RentRoll")
        If ws Is Nothing Then Set ws = ActiveWorkbook.Worksheets("UAP")
        If ws Is Nothing Then Set ws = ActiveWorkbook.Worksheets("PROJECT WORKSHEET")
    Else
        Set ws = ActiveWorkbook.Worksheets("UAP")
    End If
    On Error GoTo Fail

    If Not ws Is Nothing Then ws.Activate
    Set ReadCurrentProgramUnits = ReadUnitData()
    If Not prevSheet Is Nothing Then prevSheet.Activate
    Exit Function

Fail:
    Set ReadCurrentProgramUnits = Nothing
End Function

Private Function FindMatchingScenarioKeyForUnits(units As Collection, scenarios As Object) As String
    ' Returns the scenario key whose canonical assignments match the current live units.
    On Error GoTo Fail

    If units Is Nothing Or scenarios Is Nothing Then Exit Function

    Dim liveMap As Object
    Set liveMap = BuildCanonicalMapFromUnits(units)
    If liveMap Is Nothing Then Exit Function

    Dim scenarioKey As Variant
    For Each scenarioKey In scenarios.Keys
        Dim s As Object
        Set s = scenarios(scenarioKey)
        If s Is Nothing Then GoTo NextKey

        Dim scenMap As Object
        Set scenMap = BuildCanonicalMapFromScenario(s)
        If scenMap Is Nothing Then GoTo NextKey

        If CanonicalMapsEqual(liveMap, scenMap) Then
            FindMatchingScenarioKeyForUnits = CStr(scenarioKey)
            Exit Function
        End If

NextKey:
    Next scenarioKey

Fail:
    FindMatchingScenarioKeyForUnits = ""
End Function

Private Function BuildCanonicalMapFromUnits(units As Collection) As Object
    On Error GoTo Fail

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary") ' unit_id -> band%

    Dim i As Long
    For i = 1 To units.Count
        Dim u As Object
        Set u = units(i)
        If u Is Nothing Then GoTo NextUnit
        If Not u.Exists("unit_id") Then GoTo NextUnit
        If Not u.Exists("client_ami") Then GoTo NextUnit

        Dim unitId As String
        unitId = CStr(u("unit_id"))
        If Trim$(unitId) = "" Then GoTo NextUnit

        Dim band As Long
        band = BandPercentFromAmi(u("client_ami"))
        If band <= 0 Then GoTo NextUnit

        d(unitId) = band

NextUnit:
    Next i

    Set BuildCanonicalMapFromUnits = d
    Exit Function

Fail:
    Set BuildCanonicalMapFromUnits = Nothing
End Function

Private Function BuildCanonicalMapFromScenario(scenario As Object) As Object
    On Error GoTo Fail

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary") ' unit_id -> band%

    If scenario.Exists("canonical_assignments") Then
        Dim canon As Object
        Set canon = Nothing
        On Error Resume Next
        Set canon = scenario("canonical_assignments")
        On Error GoTo Fail

        If Not canon Is Nothing Then
            If TypeName(canon) = "Collection" Then
                Dim i As Long
                For i = 1 To canon.Count
                    Dim pair As Object
                    Set pair = canon(i)
                    If pair Is Nothing Then GoTo NextPair
                    If TypeName(pair) <> "Collection" Then GoTo NextPair
                    If pair.Count < 2 Then GoTo NextPair
                    d(CStr(pair(1))) = CLng(pair(2))
NextPair:
                Next i
                Set BuildCanonicalMapFromScenario = d
                Exit Function
            End If
        End If
    End If

    If scenario.Exists("assignments") Then
        Dim assigns As Object
        Set assigns = Nothing
        On Error Resume Next
        Set assigns = scenario("assignments")
        On Error GoTo Fail

        If Not assigns Is Nothing And TypeName(assigns) = "Collection" Then
            Dim j As Long
            For j = 1 To assigns.Count
                Dim a As Object
                Set a = assigns(j)
                If a Is Nothing Then GoTo NextAssign
                If Not a.Exists("unit_id") Then GoTo NextAssign
                If Not a.Exists("assigned_ami") Then GoTo NextAssign
                d(CStr(a("unit_id"))) = BandPercentFromAmi(a("assigned_ami"))
NextAssign:
            Next j
        End If
    End If

    Set BuildCanonicalMapFromScenario = d
    Exit Function

Fail:
    Set BuildCanonicalMapFromScenario = Nothing
End Function

Private Function CanonicalMapsEqual(a As Object, b As Object) As Boolean
    On Error GoTo Fail

    If a Is Nothing Or b Is Nothing Then Exit Function
    If a.Count <> b.Count Then Exit Function

    Dim k As Variant
    For Each k In a.Keys
        If Not b.Exists(k) Then Exit Function
        If CLng(a(k)) <> CLng(b(k)) Then Exit Function
    Next k

    CanonicalMapsEqual = True
    Exit Function

Fail:
    CanonicalMapsEqual = False
End Function

Private Function BandPercentFromAmi(value As Variant) As Long
    On Error GoTo Fail

    If Not IsNumeric(value) Then Exit Function

    Dim v As Double
    v = CDbl(value)
    If v > 2# Then
        BandPercentFromAmi = CLng(Application.WorksheetFunction.Round(v, 0))
    Else
        BandPercentFromAmi = CLng(Application.WorksheetFunction.Round(v * 100#, 0))
    End If
    Exit Function

Fail:
    BandPercentFromAmi = 0
End Function

Private Function BuildScenarioFromUnits(units As Collection) As Object
    ' Create a minimal scenario-like object for logging when the live sheet doesn't match
    ' any scenario snapshot exactly (i.e., user customized the manual scenario).
    On Error GoTo Fail

    Dim scenario As Object
    Set scenario = CreateObject("Scripting.Dictionary")

    Dim denom As Double
    Dim numer As Double
    denom = 0#
    numer = 0#

    Dim bands As Object
    Set bands = CreateObject("Scripting.Dictionary") ' band% -> True

    Dim i As Long
    For i = 1 To units.Count
        Dim u As Object
        Set u = units(i)
        If u Is Nothing Then GoTo NextUnit
        If Not u.Exists("net_sf") Then GoTo NextUnit
        If Not u.Exists("client_ami") Then GoTo NextUnit
        If Not IsNumeric(u("net_sf")) Then GoTo NextUnit
        If Not IsNumeric(u("client_ami")) Then GoTo NextUnit

        Dim sf As Double
        sf = CDbl(u("net_sf"))
        If sf <= 0 Then GoTo NextUnit

        Dim ami As Double
        ami = CDbl(u("client_ami"))
        If ami > 2# Then ami = ami / 100#
        If ami <= 0 Then GoTo NextUnit

        denom = denom + sf
        numer = numer + (sf * ami)

        Dim b As Long
        b = BandPercentFromAmi(ami)
        If b > 0 Then bands(CStr(b)) = True

NextUnit:
    Next i

    If denom > 0 Then
        scenario("waami") = (numer / denom)
    End If

    Dim bandList As New Collection
    Dim key As Variant
    For Each key In bands.Keys
        bandList.Add CLng(key)
    Next key
    Set scenario("bands") = bandList

    Set BuildScenarioFromUnits = scenario
    Exit Function

Fail:
    Set BuildScenarioFromUnits = Nothing
End Function

Private Function FindScenarioIndexForKey(scenarioKey As String, scenarios As Object) As Long
    ' Returns a stable 1-based index: strict first, then edge. 0 if not found.
    On Error GoTo Fail

    If scenarios Is Nothing Then Exit Function

    Dim strictKeys As Collection
    Dim edgeKeys As Collection
    Set strictKeys = New Collection
    Set edgeKeys = New Collection

    Dim k As Variant
    For Each k In scenarios.Keys
        Dim tier As String
        tier = ""
        On Error Resume Next
        Dim s As Object
        Set s = scenarios(k)
        If Not s Is Nothing Then
            If s.Exists("tier") Then tier = CStr(s("tier"))
        End If
        On Error GoTo Fail

        If LCase$(Trim$(tier)) = "edge" Then
            edgeKeys.Add CStr(k)
        Else
            strictKeys.Add CStr(k)
        End If
    Next k

    Dim idx As Long
    idx = 1

    For Each k In strictKeys
        If UCase$(CStr(k)) = UCase$(CStr(scenarioKey)) Then
            FindScenarioIndexForKey = idx
            Exit Function
        End If
        idx = idx + 1
    Next k

    For Each k In edgeKeys
        If UCase$(CStr(k)) = UCase$(CStr(scenarioKey)) Then
            FindScenarioIndexForKey = idx
            Exit Function
        End If
        idx = idx + 1
    Next k

Fail:
    FindScenarioIndexForKey = 0
End Function

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

Public Sub Ribbon_SelectRentRollYear(control As IRibbonControl, id As String, index As Integer)
    ' Called when user selects a year from dropdown (2022-2026).
    InitRentRollYearState

    Dim year As Long
    year = RENTROLL_YEAR_MIN + CLng(index)
    If year < RENTROLL_YEAR_MIN Or year > RENTROLL_YEAR_MAX Then Exit Sub

    m_SelectedRentRollYear = year
    SaveSetting AMI_OPTIX_REGISTRY_PATH, RENTROLL_YEAR_REG_SECTION, RENTROLL_YEAR_REG_KEY_SELECTED, CStr(year)
    DebugLog "Ribbon_SelectRentRollYear: selected=" & year, True

    ' Best-effort: ensure the API uses the selected year (rent calculator activation is server-global).
    Call EnsureSelectedRentRollYearActive(True)

    ' Warm local rent tables cache for the selected year so live sync can compute
    ' rents locally without requiring "Refresh Rent Tables" first.
    On Error Resume Next
    Dim warmSrc As String, warmCf As String, warmFp As String
    EnsureRentTablesCache year, False, warmSrc, warmCf, warmFp
    On Error GoTo 0

    Call MaybeWarnRentRollYearMismatch(year)

    ' Recalculate all scenario rents (5 solver scenarios + Scenario Manual) using
    ' the new year x current utility selections. Skip silently if scenarios are
    ' not yet on the sheet (first-time year change before any Run MIH/UAP) or if
    ' the API key is unset, so we don't fire surprise popups for what is just a
    ' dropdown change.
    On Error Resume Next
    If HasAPIKey() Then
        If HasExistingSolverScenarios() Then
            ' preserveAppliedScenario:=True keeps the scenario currently pinned
            ' to the manual block (recommended/applied) and only re-prices it at
            ' the new year — a year change must not revert it to the raw input.
            Call ManualCalculateScenario(DetectProgramFromWorkbook(), True)
        End If
    End If
    On Error GoTo 0

    ' Update the "Rent Tables Status" label to reflect the selected year.
    Call InvalidateRibbonControl("lblRentTablesStatus")
End Sub

Private Function HasExistingSolverScenarios() As Boolean
    ' Returns True if the "AMI Scenarios" sheet exists and contains at least one
    ' "SCENARIO N:" header that is not the Manual block. Used to gate the
    ' year-change auto-recalc so we don't trigger setup popups before the user
    ' has run the optimizer at least once.
    On Error GoTo Fail

    HasExistingSolverScenarios = False
    If ActiveWorkbook Is Nothing Then Exit Function

    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("AMI Scenarios")
    On Error GoTo Fail
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 1 Then Exit Function

    Dim r As Long
    Dim cellA As String
    For r = 1 To lastRow
        cellA = Trim$(CStr(ws.Cells(r, 1).Value))
        If Left$(cellA, 9) = "SCENARIO " And InStr(1, cellA, "MANUAL", vbTextCompare) = 0 Then
            HasExistingSolverScenarios = True
            Exit Function
        End If
    Next r
    Exit Function

Fail:
    HasExistingSolverScenarios = False
End Function

Public Sub Ribbon_GetRentRollYearCount(control As IRibbonControl, ByRef returnedVal)
    ' NOTE: RibbonX passes this ByRef as a Variant; keep it untyped to avoid "Type mismatch".
    returnedVal = (RENTROLL_YEAR_MAX - RENTROLL_YEAR_MIN + 1)
End Sub

Public Sub Ribbon_GetRentRollYearLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)
    ' NOTE: RibbonX passes returnedVal ByRef as a Variant.
    Dim year As Long
    year = RENTROLL_YEAR_MIN + CLng(index)
    If year < RENTROLL_YEAR_MIN Or year > RENTROLL_YEAR_MAX Then
        returnedVal = ""
    Else
        returnedVal = CStr(year)
    End If
End Sub

Public Sub Ribbon_GetRentRollYearID(control As IRibbonControl, index As Integer, ByRef returnedVal)
    ' NOTE: RibbonX passes returnedVal ByRef as a Variant.
    returnedVal = "rryear_" & index
End Sub

Public Sub Ribbon_GetRentRollYearSelectedIndex(control As IRibbonControl, ByRef returnedVal)
    ' NOTE: RibbonX passes returnedVal ByRef as a Variant.
    InitRentRollYearState

    Dim idx As Long
    idx = CLng(m_SelectedRentRollYear - RENTROLL_YEAR_MIN)
    If idx < 0 Or idx > (RENTROLL_YEAR_MAX - RENTROLL_YEAR_MIN) Then
        idx = CLng(RENTROLL_YEAR_DEFAULT - RENTROLL_YEAR_MIN)
    End If
    returnedVal = idx
End Sub

Public Sub Ribbon_GetRentTablesStatusLabel(control As IRibbonControl, ByRef returnedVal)
    ' Non-invasive: show what the per-user cache was last built from (meta-only; no network calls).
    On Error GoTo SafeExit

    InitRentRollYearState

    Dim year As Long
    year = m_SelectedRentRollYear

    Dim meta As Object
    Set meta = Nothing

    Dim label As String
    label = "Rent Tables: " & CStr(year) & " | cache: (missing)"

    If TryReadRentTablesCacheMeta(year, meta) Then
        Dim sourceLabel As String
        sourceLabel = GetDictString(meta, "source_label", "")
        If sourceLabel = "" Then sourceLabel = InferSourceLabelFromPath(GetDictString(meta, "source_path", ""))

        Dim builtAt As String
        builtAt = GetDictString(meta, "generated_at", "")

        If sourceLabel = "" Then sourceLabel = "?"
        If builtAt = "" Then builtAt = "?"

        label = "Rent Tables: " & CStr(year) & " | " & sourceLabel & " | built " & builtAt
    End If

    returnedVal = label
    Exit Sub

SafeExit:
    returnedVal = "Rent Tables: (status unavailable)"
End Sub

Public Sub Ribbon_ManageRentRollYears(control As IRibbonControl)
    ' Upload/replace year calculator files and activate on the API.
    ' UI is intentionally lightweight (FileDialog) to avoid adding new UserForm files.
    On Error GoTo Fail

    InitRentRollYearState

    Dim year As Long
    year = m_SelectedRentRollYear

    If ActiveWorkbook Is Nothing Then
        MsgBox "Open a workbook first.", vbExclamation, "AMI Optix"
        Exit Sub
    End If

    Dim fd As Object
    Set fd = Application.FileDialog(3) ' msoFileDialogFilePicker
    If fd Is Nothing Then Exit Sub

    With fd
        .Title = "Select rent calculator workbook for year " & CStr(year)
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Workbooks", "*.xlsx; *.xlsm"
        If .Show <> -1 Then Exit Sub
    End With

    Dim selectedPath As String
    selectedPath = CStr(fd.SelectedItems(1))
    If Trim$(selectedPath) = "" Then Exit Sub

    Dim yearFolder As String
    yearFolder = EnsureRentRollYearFolder(year)
    If Trim$(yearFolder) = "" Then
        MsgBox "Could not create Rent Roll Years folder under %APPDATA%.", vbExclamation, "AMI Optix"
        Exit Sub
    End If

    Dim localPath As String
    localPath = yearFolder & "\RentCalculator_" & CStr(year) & ".xlsx"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then
        MsgBox "File system access is unavailable.", vbExclamation, "AMI Optix"
        Exit Sub
    End If

    fso.CopyFile selectedPath, localPath, True

    Dim remoteName As String
    remoteName = RentCalculatorRemoteNameForYear(year)

    If Not UploadRentCalculatorFile(localPath, remoteName, True, True) Then Exit Sub
    If Not ActivateRentCalculatorByName(remoteName, True) Then Exit Sub

    MsgBox "Uploaded and activated rent calculator for year " & CStr(year) & ".", vbInformation, "AMI Optix"
    Exit Sub

Fail:
    MsgBox "Manage Rent Roll Years failed: " & Err.Description, vbExclamation, "AMI Optix"
End Sub

Public Sub Ribbon_ShowRentTablesStatus(control As IRibbonControl)
    ' Shows the Rent Tables section in Diagnostics (non-modal; no MsgBox on success).
    On Error GoTo Fail

    Call ShowAMIOptixDiagnostics

    If Not ActiveWorkbook Is Nothing Then
        Dim ws As Worksheet
        Set ws = Nothing
        On Error Resume Next
        Set ws = ActiveWorkbook.Worksheets("AMI Optix Diagnostics")
        On Error GoTo Fail

        If Not ws Is Nothing Then
            Dim found As Range
            Set found = ws.Columns(1).Find(What:="Rent Tables Status", After:=ws.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, _
                                           SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
            If Not found Is Nothing Then
                ws.Activate
                found.Select
            End If
        End If
    End If

    EnsureAMIOptixTabActive
    Exit Sub

Fail:
    MsgBox "Rent Tables Status failed: " & Err.Description, vbExclamation, "AMI Optix"
    EnsureAMIOptixTabActive
End Sub

Public Sub Ribbon_RefreshRentTablesCache(control As IRibbonControl)
    ' Fix-06c: Force-refresh the per-user normalized rent tables cache (CSV) for the selected year.
    On Error GoTo Fail

    InitRentRollYearState

    Dim year As Long
    year = m_SelectedRentRollYear

    Dim sourcePath As String
    Dim cacheFolder As String
    Dim fingerprint As String
    sourcePath = ""
    cacheFolder = ""
    fingerprint = ""

    Call EnsureRentTablesCache(year, True, sourcePath, cacheFolder, fingerprint)

    MsgBox "Rent tables cache refreshed for " & CStr(year) & " from " & sourcePath & " -> " & cacheFolder, vbInformation, "AMI Optix"
    Call InvalidateRibbonControl("lblRentTablesStatus")
    Exit Sub

Fail:
    MsgBox Err.Description, vbCritical, "AMI Optix - Refresh Rent Tables"
End Sub

Public Sub Ribbon_VerifyManualRentsAPI(control As IRibbonControl)
    ' Fix-06d: One-click verification that locally-computed manual rents match /api/evaluate.
    ' IMPORTANT: This must NOT trigger any automatic API calls on edits; only runs on button click.
    On Error GoTo Fail

    Call VerifyManualRentsAPI
    EnsureAMIOptixTabActive
    Exit Sub

Fail:
    MsgBox "Verify Manual Rents (API) failed: " & Err.Description, vbExclamation, "AMI Optix"
    EnsureAMIOptixTabActive
End Sub

Public Sub Ribbon_SelectRentRoll(control As IRibbonControl, id As String, index As Integer)
    ' Called when user selects a rent roll from dropdown
    If index >= 0 And index < m_RentRollCount Then
        m_SelectedRentRoll = m_RentRollSheets(index)
        DebugLog "Ribbon_SelectRentRoll: selected=" & m_SelectedRentRoll, True
    End If
    EnsureAMIOptixTabActive
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
    EnsureAMIOptixTabActive
End Sub

Private Sub RefreshRentRollList()
    ' Scans workbook for sheets that could be rent rolls
    Dim ws As Worksheet
    Dim tempSheets() As String
    Dim sheetCount As Long
    Dim preferredNames As Variant
    Dim i As Long

    Dim wbName As String
    wbName = "(none)"
    If Not ActiveWorkbook Is Nothing Then wbName = ActiveWorkbook.Name
    DebugLog "RefreshRentRollList: start workbook=" & wbName, True

    If ActiveWorkbook Is Nothing Then
        m_RentRollCount = 1
        ReDim m_RentRollSheets(0 To 0)
        m_RentRollSheets(0) = "(No workbook open)"
        DebugLog "RefreshRentRollList: no workbook", True
        Exit Sub
    End If

    sheetCount = 0
    ReDim tempSheets(0 To Application.Max(0, ActiveWorkbook.Worksheets.Count - 1))

    ' IMPORTANT: Do NOT scan sheet cell contents here.
    ' Some client workbooks contain formulas that reference VBA UDFs, and those UDFs can fail to compile
    ' (e.g., missing "Microsoft Scripting Runtime" reference). Reading lots of cell values on ribbon-load
    ' can trigger that compile and show a confusing error unrelated to AMI Optix.
    '
    ' We therefore build this list using SHEET NAMES ONLY.
    preferredNames = Array("UAP", "MIH", "PROJECT WORKSHEET", "RentRoll", "Rent Roll", "Units", "Unit Schedule", "Sheet1", "Data")

    ' First add preferred sheets in order
    For i = LBound(preferredNames) To UBound(preferredNames)
        On Error Resume Next
        Set ws = ActiveWorkbook.Worksheets(CStr(preferredNames(i)))
        On Error GoTo 0

        If Not ws Is Nothing Then
            tempSheets(sheetCount) = ws.Name
            sheetCount = sheetCount + 1
            Set ws = Nothing
        End If
    Next i

    ' Then add other likely sheets by NAME pattern.
    For Each ws In ActiveWorkbook.Worksheets
        If Not IsSheetInArray(ws.Name, tempSheets, sheetCount) Then
            If IsLikelyRentRollSheetName(ws.Name) Then
                tempSheets(sheetCount) = ws.Name
                sheetCount = sheetCount + 1
            End If
        End If
    Next ws

    ' Fallback: if nothing matched, include all sheets (still name-only).
    If sheetCount = 0 Then
        For Each ws In ActiveWorkbook.Worksheets
            tempSheets(sheetCount) = ws.Name
            sheetCount = sheetCount + 1
        Next ws
    End If

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

    DebugLog "RefreshRentRollList: done count=" & m_RentRollCount, True
End Sub

Private Function IsLikelyRentRollSheetName(sheetName As String) As Boolean
    Dim s As String
    s = UCase$(Trim$(sheetName))

    ' Never include our output sheets in the rent roll picker.
    If s = "AMI SCENARIOS" Or s = "AMI OPTIX DIAGNOSTICS" Then Exit Function

    ' Prefer "data-like" sheet names.
    If InStr(1, s, "RENT", vbTextCompare) > 0 Then IsLikelyRentRollSheetName = True: Exit Function
    If InStr(1, s, "ROLL", vbTextCompare) > 0 Then IsLikelyRentRollSheetName = True: Exit Function
    If InStr(1, s, "UNIT", vbTextCompare) > 0 Then IsLikelyRentRollSheetName = True: Exit Function
    If InStr(1, s, "SCHEDULE", vbTextCompare) > 0 Then IsLikelyRentRollSheetName = True: Exit Function
    If InStr(1, s, "PROJECT", vbTextCompare) > 0 Then IsLikelyRentRollSheetName = True: Exit Function
    If s = "UAP" Or s = "MIH" Then IsLikelyRentRollSheetName = True: Exit Function
End Function

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
    EnsureAMIOptixTabActive
End Sub

Public Sub Ribbon_OpenAPISettings(control As IRibbonControl)
    ' Open API settings dialog
    ShowSettingsForm
    EnsureAMIOptixTabActive
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
        EnsureAMIOptixTabActive
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
        newRoot = InputBox("Enter log root folder path (no quotes) (e.g. Z:\AMI_Optix_Learning or \\server\share\AMI_Optix_Learning):", _
                           "AMI Optix - Log Folder", currentRoot)
        If Trim$(newRoot) <> "" Then
            Call SetLearningLogRootPath(newRoot)
        End If
    End If

    MsgBox "Saved." & vbCrLf & vbCrLf & _
           "Log Root: " & GetLearningLogRootPath() & vbCrLf & _
           "Run Log File: " & GetRunLogFilePath(), _
           vbInformation, "AMI Optix"
    EnsureAMIOptixTabActive
    Exit Sub

Fail:
    MsgBox "Log settings failed: " & Err.Description, vbExclamation, "AMI Optix"
    EnsureAMIOptixTabActive
End Sub

Public Sub Ribbon_OpenLearningLogs(control As IRibbonControl)
    ' Opens the log folder in Explorer.
    On Error GoTo Fail

    If ActiveWorkbook Is Nothing Then
        EnsureAMIOptixTabActive
        Exit Sub
    End If

    Dim folderPath As String
    folderPath = GetLearningLogRootPath()
    Call EnsureFolderExists(folderPath)

    Shell "explorer.exe """ & folderPath & """", vbNormalFocus
    EnsureAMIOptixTabActive
    Exit Sub

Fail:
    EnsureAMIOptixTabActive
End Sub

'-------------------------------------------------------------------------------
' RIBBON CALLBACKS - HELP GROUP
'-------------------------------------------------------------------------------

Public Sub Ribbon_ShowDiagnostics(control As IRibbonControl)
    On Error GoTo Fail
    Call ShowAMIOptixDiagnostics
    EnsureAMIOptixTabActive
    Exit Sub
Fail:
    MsgBox "Diagnostics failed: " & Err.Description, vbExclamation, "AMI Optix"
    EnsureAMIOptixTabActive
End Sub

Public Sub Ribbon_ShowAbout(control As IRibbonControl)
    ' Show about dialog
    MsgBox "AMI Optix Excel Add-in" & vbCrLf & vbCrLf & _
           "Version 1.0" & vbCrLf & _
           "NYC Affordable Housing AMI Optimizer" & vbCrLf & vbCrLf & _
           "Optimizes AMI band assignments for affordable housing " & _
           "projects to maximize revenue while meeting regulatory requirements." & vbCrLf & vbCrLf & _
           "API: " & API_BASE_URL, _
           vbInformation, "About AMI Optix"
    EnsureAMIOptixTabActive
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

    ' Some client workbooks hide sheets via macros; force visibility so the tab doesn't "disappear".
    On Error Resume Next
    ws.Visible = xlSheetVisible
    ws.Activate
    On Error GoTo 0

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

    ' Same grouped, de-duped order as the sheet's overview table and the
    ' numbered detail blocks — so "3" in this picker is "SCENARIO 3" on the
    ' sheet. (Was strict-then-edge, which could disagree with the sheet.)
    Dim groupLabels As Collection
    Dim orderedKeys As Collection
    Set orderedKeys = BuildGroupedScenarioOrder(scenarios, groupLabels)

    If orderedKeys.Count = 0 Then
        MsgBox "No scenarios available.", vbInformation, "AMI Optix"
        Exit Sub
    End If

    ReDim keys(1 To orderedKeys.Count)
    msg = "Available Scenarios:" & vbCrLf & vbCrLf
    Dim lastGroup As String
    lastGroup = ""

    For i = 1 To orderedKeys.Count
        keys(i) = CStr(orderedKeys(i))
        Dim grpLabel As String
        grpLabel = CStr(groupLabels(i))
        If grpLabel <> lastGroup Then
            msg = msg & "--- " & grpLabel & " ---" & vbCrLf
            lastGroup = grpLabel
        End If
        msg = msg & ScenarioPickerLine(i, keys(i), scenarios(keys(i))) & vbCrLf
    Next i

    msg = msg & vbCrLf & "Enter scenario number (1-" & orderedKeys.Count & "):"

    choice = InputBox(msg, "Select Scenario", "1")

    If choice = "" Then Exit Sub  ' Cancelled

    Dim choiceNum As Long
    On Error Resume Next
    choiceNum = CLng(choice)
    On Error GoTo 0

    If choiceNum < 1 Or choiceNum > orderedKeys.Count Then
        MsgBox "Invalid selection.", vbExclamation, "AMI Optix"
        Exit Sub
    End If

    ' Apply selected scenario
    ApplyScenarioByKey keys(choiceNum)
End Sub

Private Function FormatScenarioNameForPicker(key As String) As String
    ' Client-facing names for the View Scenario picker. Mirrors the sheet's
    ' FormatScenarioName labels so the picker and the sheet agree, in Title
    ' Case for the list dialog.
    Dim k As String
    k = LCase$(Trim$(key))

    If Left$(k, Len("fewest_40_units")) = "fewest_40_units" Then
        FormatScenarioNameForPicker = "Fewest 40% Units"
        Exit Function
    End If
    If Left$(k, Len("tight_40_footprint")) = "tight_40_footprint" Then
        FormatScenarioNameForPicker = "Tighter 40% Footprint"
        Exit Function
    End If
    If Left$(k, Len("edge_waami_floor")) = "edge_waami_floor" Then
        FormatScenarioNameForPicker = "Higher Rent (More 40% Units)"
        Exit Function
    End If
    If Left$(k, Len("edge_min_share")) = "edge_min_share" Or Left$(k, Len("edge_max_share")) = "edge_max_share" Then
        FormatScenarioNameForPicker = "Higher Rent (Relaxed Share)"
        Exit Function
    End If

    Select Case k
        Case "low_40_share"
            FormatScenarioNameForPicker = "Low 40% Share"
        Case "mid_40_share"
            FormatScenarioNameForPicker = "Mid-Range 40%"
        Case "max_40_share"
            FormatScenarioNameForPicker = "Max 40% Share"
        Case "absolute_best"
            FormatScenarioNameForPicker = "Maximum Rent"
        Case "best_rent_roll"
            FormatScenarioNameForPicker = "Best Rent Roll"
        Case "max_revenue"
            FormatScenarioNameForPicker = "Maximum Rent"
        Case "best_3_band"
            FormatScenarioNameForPicker = "Three-Band Mix"
        Case "best_2_band"
            FormatScenarioNameForPicker = "Two-Band Mix"
        Case "closest_to_60"
            FormatScenarioNameForPicker = "Closest to 60% Cap"
        Case "alternative"
            FormatScenarioNameForPicker = "Alternative Mix"
        Case "client_oriented"
            FormatScenarioNameForPicker = "Client Oriented"
        Case "original"
            FormatScenarioNameForPicker = "Your Original Input"
        Case Else
            FormatScenarioNameForPicker = Replace(key, "_", " ")
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

Private Sub InitRentRollYearState()
    If m_RentRollYearInitialized Then Exit Sub

    Dim raw As String
    raw = GetSetting(AMI_OPTIX_REGISTRY_PATH, RENTROLL_YEAR_REG_SECTION, RENTROLL_YEAR_REG_KEY_SELECTED, CStr(RENTROLL_YEAR_DEFAULT))

    Dim y As Long
    y = RENTROLL_YEAR_DEFAULT
    On Error Resume Next
    y = CLng(raw)
    On Error GoTo 0

    If y < RENTROLL_YEAR_MIN Or y > RENTROLL_YEAR_MAX Then y = RENTROLL_YEAR_DEFAULT

    m_SelectedRentRollYear = y
    m_RentRollYearInitialized = True
End Sub

Private Function EnsureRentRollYearFolder(year As Long) As String
    ' Ensure: %APPDATA%\AMI_Optix\RentRollYears\<year>\
    On Error GoTo Fail

    Dim appData As String
    appData = Environ$("APPDATA")
    If Trim$(appData) = "" Then Exit Function

    Dim basePath As String
    basePath = appData & "\AMI_Optix"
    Call EnsureFolderExists(basePath)

    Dim yearsPath As String
    yearsPath = basePath & "\RentRollYears"
    Call EnsureFolderExists(yearsPath)

    Dim yearPath As String
    yearPath = yearsPath & "\" & CStr(year)
    Call EnsureFolderExists(yearPath)

    If Dir(yearPath, vbDirectory) = "" Then Exit Function

    EnsureRentRollYearFolder = yearPath
    Exit Function

Fail:
End Function

Private Function RentCalculatorRemoteNameForYear(year As Long) As String
    RentCalculatorRemoteNameForYear = "AMI_Optix_Rent_Calculator_" & CStr(year) & ".xlsx"
End Function

Private Sub MaybeWarnRentRollYearMismatch(selectedYear As Long)
    ' Best-effort warning only; OK continues (not blocking).
    On Error GoTo Fail

    Dim declaredYear As Long
    declaredYear = DetectWorkbookDeclaredRentRollYear()
    If declaredYear <= 0 Then Exit Sub
    If declaredYear = selectedYear Then Exit Sub

    MsgBox "Warning: workbook appears to declare Rent Roll year " & CStr(declaredYear) & _
           ", but Rent Roll Year is set to " & CStr(selectedYear) & "." & vbCrLf & vbCrLf & _
           "Click OK to continue (the selected year will be used).", _
           vbExclamation, "AMI Optix"
    Exit Sub

Fail:
End Sub

Public Function DetectWorkbookDeclaredRentRollYear() As Long
    ' Best-effort detection: workbook name -> rent roll sheet name -> small top-of-sheet scan.
    On Error GoTo Fail

    If ActiveWorkbook Is Nothing Then Exit Function

    Dim y As Long
    y = FindYearInText(ActiveWorkbook.Name)
    If y > 0 Then DetectWorkbookDeclaredRentRollYear = y: Exit Function

    Dim ws As Worksheet
    Set ws = Nothing

    If Trim$(m_SelectedRentRoll) <> "" Then
        On Error Resume Next
        Set ws = ActiveWorkbook.Worksheets(m_SelectedRentRoll)
        On Error GoTo Fail
    End If

    If ws Is Nothing Then
        On Error Resume Next
        Set ws = ActiveWorkbook.Worksheets("RentRoll")
        If ws Is Nothing Then Set ws = ActiveWorkbook.Worksheets("UAP")
        If ws Is Nothing Then Set ws = ActiveWorkbook.Worksheets("MIH")
        On Error GoTo Fail
    End If

    If ws Is Nothing Then Exit Function

    y = FindYearInText(ws.Name)
    If y > 0 Then DetectWorkbookDeclaredRentRollYear = y: Exit Function

    Dim r As Long, c As Long
    For r = 1 To 10
        For c = 1 To 8
            Dim v As Variant
            v = ws.Cells(r, c).Value

            If IsNumeric(v) Then
                On Error Resume Next
                y = CLng(v)
                On Error GoTo Fail
                If y >= RENTROLL_YEAR_MIN And y <= RENTROLL_YEAR_MAX Then
                    DetectWorkbookDeclaredRentRollYear = y
                    Exit Function
                End If
            ElseIf VarType(v) = vbString Then
                y = FindYearInText(CStr(v))
                If y > 0 Then DetectWorkbookDeclaredRentRollYear = y: Exit Function
            End If
        Next c
    Next r

    Exit Function

Fail:
End Function

Private Function FindYearInText(text As String) As Long
    Dim y As Long
    For y = RENTROLL_YEAR_MAX To RENTROLL_YEAR_MIN Step -1
        If InStr(1, text, CStr(y), vbTextCompare) > 0 Then
            FindYearInText = y
            Exit Function
        End If
    Next y
End Function
