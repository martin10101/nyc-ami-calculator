Attribute VB_Name = "AMI_Optix_EventHooks"
'===============================================================================
' AMI OPTIX - Application Event Hooks
' Keeps the "Scenario Manual" block synced with UAP edits.
'===============================================================================
Option Explicit

Public g_AMIOptixAppEvents As AMI_Optix_AppEvents
Public g_AMIOptixSuppressEvents As Boolean
Public g_AMIOptixLiveSyncEnabled As Boolean
Public g_AMIOptixVisibilityWorkbookName As String

' Custom-undo state for AMI cell edits. The live-sync refresh that runs after
' each AMI edit writes to other cells, which always clears Excel's native undo
' stack. We register a custom Application.OnUndo handler so Ctrl+Z can still
' restore the most recently edited AMI cell.
Public g_AMIOptixUndoWorkbookName As String
Public g_AMIOptixUndoSheetName As String
Public g_AMIOptixUndoAddress As String
Public g_AMIOptixUndoOldValue As Variant
Public g_AMIOptixUndoProgramNorm As String
Public g_AMIOptixUndoArmed As Boolean

Private m_LiveSyncInitialized As Boolean
Private m_EnsureVisibleScheduled As Boolean
Private m_EnsureVisibleAt As Date

Public Sub EnsureLiveSyncInitialized()
    If m_LiveSyncInitialized Then Exit Sub

    ' Fix-05: Manual Working Copy is always-on; live sync cannot be disabled.
    g_AMIOptixLiveSyncEnabled = True

    m_LiveSyncInitialized = True
End Sub

Public Function GetLiveSyncEnabled() As Boolean
    ' Fix-05: Manual Working Copy is always-on; live sync cannot be disabled.
    GetLiveSyncEnabled = True
End Function

Public Sub SetLiveSyncEnabled(enabled As Boolean)
    EnsureLiveSyncInitialized
    ' Fix-05: Manual Working Copy is always-on; ignore requested state.
    g_AMIOptixLiveSyncEnabled = True
End Sub

Public Sub Auto_Open()
    StartAMIOptixEventHooks
End Sub

Public Sub Auto_Close()
    StopAMIOptixEventHooks
End Sub

Public Sub StartAMIOptixEventHooks()
    On Error Resume Next
    EnsureLiveSyncInitialized
    g_AMIOptixVisibilityWorkbookName = ""
    Set g_AMIOptixAppEvents = New AMI_Optix_AppEvents
    Set g_AMIOptixAppEvents.App = Application
    On Error GoTo 0
End Sub

Public Sub StopAMIOptixEventHooks()
    On Error Resume Next
    Call AMI_Optix_CancelEnsureSheetsVisible
    If Not g_AMIOptixAppEvents Is Nothing Then
        Set g_AMIOptixAppEvents.App = Nothing
    End If
    Set g_AMIOptixAppEvents = Nothing
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' SHEET VISIBILITY GUARD
'-------------------------------------------------------------------------------

Public Sub AMI_Optix_EnsureSheetsVisible()
    ' Some client workbooks hide sheets via macros. Re-assert that our sheets remain visible.
    On Error GoTo SafeExit

    Dim wb As Workbook
    Set wb = Nothing

    If Trim$(g_AMIOptixVisibilityWorkbookName) <> "" Then
        On Error Resume Next
        Set wb = Application.Workbooks(g_AMIOptixVisibilityWorkbookName)
        On Error GoTo SafeExit
    End If

    If wb Is Nothing Then
        Set wb = ActiveWorkbook
    End If
    If wb Is Nothing Then GoTo SafeExit

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets("AMI Scenarios")
    If Not ws Is Nothing Then ws.Visible = xlSheetVisible
    Set ws = Nothing
    Set ws = wb.Worksheets("AMI Optix Diagnostics")
    If Not ws Is Nothing Then ws.Visible = xlSheetVisible
    On Error GoTo SafeExit

SafeExit:
End Sub

Public Sub AMI_Optix_ScheduleEnsureSheetsVisible(Optional workbookName As String = "")
    ' Schedule a second-pass visibility restore after other workbook events run.
    On Error GoTo SafeExit

    If Trim$(workbookName) <> "" Then
        g_AMIOptixVisibilityWorkbookName = workbookName
    End If

    Call AMI_Optix_CancelEnsureSheetsVisible

    ' Delay by ~1 second to run after client Workbook/Sheet_Activate macros.
    m_EnsureVisibleAt = Now + TimeSerial(0, 0, 1)
    Application.OnTime EarliestTime:=m_EnsureVisibleAt, Procedure:="AMI_Optix_EnsureSheetsVisible", Schedule:=True
    m_EnsureVisibleScheduled = True

SafeExit:
End Sub

Public Sub AMI_Optix_CancelEnsureSheetsVisible()
    On Error Resume Next
    If m_EnsureVisibleScheduled Then
        Application.OnTime EarliestTime:=m_EnsureVisibleAt, Procedure:="AMI_Optix_EnsureSheetsVisible", Schedule:=False
        m_EnsureVisibleScheduled = False
    End If
    On Error GoTo 0
End Sub

Public Sub AMI_Optix_ArmUndoForAmiEdit(target As Range, oldValue As Variant, programNorm As String)
    ' Capture the pre-edit AMI cell state so Ctrl+Z can restore it. Excel's
    ' native undo stack gets cleared by the live-sync refresh that runs after
    ' an AMI edit, so we register a custom Application.OnUndo handler that
    ' fires before the user's next macro-clearing action.
    On Error GoTo SafeExit
    If target Is Nothing Then Exit Sub

    Dim wb As String
    Dim ws As String
    On Error Resume Next
    wb = CStr(target.Worksheet.Parent.Name)
    ws = CStr(target.Worksheet.Name)
    On Error GoTo SafeExit

    g_AMIOptixUndoWorkbookName = wb
    g_AMIOptixUndoSheetName = ws
    g_AMIOptixUndoAddress = target.Address(False, False)
    g_AMIOptixUndoOldValue = oldValue
    g_AMIOptixUndoProgramNorm = CStr(programNorm)
    g_AMIOptixUndoArmed = True

    ' Application.OnUndo requires the procedure name fully qualified when the
    ' code lives in an add-in (.xlam) and the active workbook is a different
    ' file. Without the "'AMI_Optix.xlam'!" prefix, Excel silently fails to
    ' resolve the procedure when the user presses Ctrl+Z.
    Dim addinName As String
    addinName = ThisWorkbook.Name
    Dim procRef As String
    procRef = "'" & addinName & "'!AMI_Optix_UndoLastAmiEdit"

    On Error Resume Next
    Application.OnUndo "Undo AMI change", procRef
    ' Log for diagnostics so we can confirm the registration in the debug log.
    On Error Resume Next
    DebugLog "OnUndo armed: " & procRef & " for " & wb & "!" & ws & "!" & target.Address(False, False) & " (old=" & CStr(oldValue) & ")", True

SafeExit:
End Sub

Public Sub AMI_Optix_UndoLastAmiEdit()
    ' Restores the most recently edited AMI cell to its pre-edit value. Called
    ' by Excel via Application.OnUndo when the user presses Ctrl+Z.
    On Error Resume Next
    DebugLog "OnUndo fired: armed=" & CStr(g_AMIOptixUndoArmed) & " sheet=" & g_AMIOptixUndoSheetName & " addr=" & g_AMIOptixUndoAddress & " old=" & CStr(g_AMIOptixUndoOldValue), True
    On Error GoTo SafeExit
    If Not g_AMIOptixUndoArmed Then Exit Sub
    If g_AMIOptixUndoSheetName = "" Or g_AMIOptixUndoAddress = "" Then Exit Sub

    Dim wb As Workbook
    Set wb = Nothing
    On Error Resume Next
    Set wb = Application.Workbooks(g_AMIOptixUndoWorkbookName)
    On Error GoTo SafeExit
    If wb Is Nothing Then
        On Error Resume Next
        Set wb = ActiveWorkbook
        On Error GoTo SafeExit
    End If
    If wb Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets(g_AMIOptixUndoSheetName)
    On Error GoTo SafeExit
    If ws Is Nothing Then Exit Sub

    Dim prevEnableEvents As Boolean
    Dim prevSuppress As Boolean
    prevEnableEvents = Application.EnableEvents
    prevSuppress = g_AMIOptixSuppressEvents
    Application.EnableEvents = False
    g_AMIOptixSuppressEvents = True

    On Error Resume Next
    ws.Range(g_AMIOptixUndoAddress).Value = g_AMIOptixUndoOldValue
    ws.Range(g_AMIOptixUndoAddress).NumberFormat = "0%"
    On Error GoTo SafeExit

    Application.EnableEvents = prevEnableEvents
    g_AMIOptixSuppressEvents = prevSuppress

    ' One-shot — clear so subsequent Ctrl+Z presses don't re-fire the same undo.
    g_AMIOptixUndoArmed = False

    ' Refresh the Manual Working Copy with the restored AMI value so rents
    ' reflect the undone state. Best-effort; ignore errors.
    On Error Resume Next
    If g_AMIOptixUndoProgramNorm <> "" Then
        Call RefreshManualWorkingCopyLocalRents(g_AMIOptixUndoProgramNorm)
    End If
    On Error GoTo SafeExit

SafeExit:
    On Error Resume Next
    Application.EnableEvents = True
End Sub
