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

' Deferred Manual Working Copy refresh. We delay the refresh by ~2 seconds
' after the last AMI edit so Excel's native Ctrl+Z keeps working during that
' window. If the user makes another edit before the timer fires, we cancel
' and reschedule. The refresh also self-cancels if the AMI workbook is no
' longer active when the timer fires — otherwise our writes would land in an
' unrelated workbook and wipe its undo stack.
Public g_AMIOptixDeferredRefreshAt As Date
Public g_AMIOptixDeferredRefreshProgramNorm As String
Public g_AMIOptixDeferredRefreshWorkbook As String

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
    ' Self-heal: release any stale Ctrl+Z hijack left over from prior buggy
    ' add-in versions. Calling OnKey with no second argument resets the key
    ' to Excel's default behavior. Idempotent and safe on fresh sessions —
    ' must run BEFORE anything else so installs upgrading from the broken
    ' version get their global Ctrl+Z back immediately.
    On Error Resume Next
    Application.OnKey "^z"
    Application.OnKey "^+z"
    On Error GoTo 0

    On Error Resume Next
    EnsureLiveSyncInitialized
    g_AMIOptixVisibilityWorkbookName = ""
    Set g_AMIOptixAppEvents = New AMI_Optix_AppEvents
    Set g_AMIOptixAppEvents.App = Application
    On Error GoTo 0
End Sub

Public Sub StopAMIOptixEventHooks()
    ' Release any pending timers and event sinks. We deliberately do NOT wrap
    ' the OnKey releases in On Error Resume Next — if those fail to release,
    ' Excel's Ctrl+Z stays hijacked for the rest of the session and breaks
    ' undo in EVERY workbook. We need to know if that ever happens.
    Call AMI_Optix_CancelEnsureSheetsVisible
    Call AMI_Optix_CancelDeferredRefresh

    ' Unconditional release so Excel's default shortcuts are restored even if
    ' the AppEvents object was already torn down or never created.
    Application.OnKey "^z"
    Application.OnKey "^+z"

    On Error Resume Next
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

Public Sub AMI_Optix_ScheduleDeferredRefresh(programNorm As String)
    ' Debounced Manual Working Copy refresh: cancel any pending refresh,
    ' schedule a new one ~2 seconds out. The delay gives the user a window
    ' to press Ctrl+Z natively (Excel's undo stack stays intact as long as
    ' we haven't written to any cell).
    On Error Resume Next

    Call AMI_Optix_CancelDeferredRefresh

    ' Capture which workbook initiated this refresh. If the user switches to
    ' a different workbook before the timer fires, AMI_Optix_DoDeferredRefresh
    ' will bail out rather than write into the wrong workbook (which would
    ' clear that workbook's Ctrl+Z undo stack).
    g_AMIOptixDeferredRefreshWorkbook = ""
    If Not ActiveWorkbook Is Nothing Then
        g_AMIOptixDeferredRefreshWorkbook = CStr(ActiveWorkbook.Name)
    End If

    g_AMIOptixDeferredRefreshProgramNorm = CStr(programNorm)
    g_AMIOptixDeferredRefreshAt = Now + TimeSerial(0, 0, 2)
    Application.OnTime EarliestTime:=g_AMIOptixDeferredRefreshAt, _
                        Procedure:="'" & ThisWorkbook.Name & "'!AMI_Optix_DoDeferredRefresh"
End Sub

Public Sub AMI_Optix_CancelDeferredRefresh()
    ' Cancel any pending deferred refresh. Called when the user switches to
    ' a different workbook (so our timer doesn't fire into the wrong workbook
    ' and clear THAT workbook's undo stack) or when the add-in unloads.
    On Error Resume Next
    If g_AMIOptixDeferredRefreshAt <> 0 Then
        Application.OnTime EarliestTime:=g_AMIOptixDeferredRefreshAt, _
                            Procedure:="'" & ThisWorkbook.Name & "'!AMI_Optix_DoDeferredRefresh", _
                            Schedule:=False
        g_AMIOptixDeferredRefreshAt = 0
    End If
End Sub

Public Sub AMI_Optix_DoDeferredRefresh()
    ' Fires (via OnTime) ~2 seconds after the last AMI edit. Performs the
    ' Manual Working Copy refresh that was previously immediate.
    On Error Resume Next
    g_AMIOptixDeferredRefreshAt = 0

    Dim prog As String
    Dim scheduledWb As String
    prog = g_AMIOptixDeferredRefreshProgramNorm
    scheduledWb = g_AMIOptixDeferredRefreshWorkbook
    If prog = "" Then Exit Sub

    ' GUARD: only refresh when the workbook that scheduled the timer is still
    ' the active workbook. If the user switched to a different workbook (or
    ' closed the AMI workbook entirely), writing 100+ cells now would clear
    ' Excel's session-wide undo stack and break Ctrl+Z in their other docs.
    ' On return, App_WorkbookActivate reschedules the refresh.
    If ActiveWorkbook Is Nothing Then Exit Sub
    If scheduledWb <> "" Then
        If CStr(ActiveWorkbook.Name) <> scheduledWb Then Exit Sub
    End If
    ' Belt-and-suspenders: must have an "AMI Scenarios" sheet to be a real AMI
    ' workbook. Prevents accidental fire-through on a random workbook that
    ' happens to share a name (rare, but cheap to guard).
    If Not WorkbookHasAmiScenariosSheet(ActiveWorkbook) Then Exit Sub

    Dim prevEnableEvents As Boolean
    Dim prevSuppress As Boolean
    prevEnableEvents = Application.EnableEvents
    prevSuppress = g_AMIOptixSuppressEvents
    Application.EnableEvents = False
    g_AMIOptixSuppressEvents = True

    Call RefreshManualWorkingCopyLocalRents(prog)

    Application.EnableEvents = prevEnableEvents
    g_AMIOptixSuppressEvents = prevSuppress
End Sub

Private Function WorkbookHasAmiScenariosSheet(wb As Workbook) As Boolean
    On Error GoTo Fail
    If wb Is Nothing Then Exit Function
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets("AMI Scenarios")
    On Error GoTo Fail
    WorkbookHasAmiScenariosSheet = (Not ws Is Nothing)
    Exit Function
Fail:
    WorkbookHasAmiScenariosSheet = False
End Function

Public Sub AMI_Optix_OnWorkbookActivate(ByVal Wb As Workbook)
    ' Called by App_WorkbookActivate. If we have a pending refresh for THIS
    ' workbook and the timer was canceled while the user was elsewhere,
    ' reschedule so the refresh fires on their return.
    On Error Resume Next
    If Wb Is Nothing Then Exit Sub
    If g_AMIOptixDeferredRefreshWorkbook = "" Then Exit Sub
    If CStr(Wb.Name) <> g_AMIOptixDeferredRefreshWorkbook Then Exit Sub
    If g_AMIOptixDeferredRefreshProgramNorm = "" Then Exit Sub

    ' Only reschedule if no timer is currently armed (avoid double-scheduling
    ' the same refresh).
    If g_AMIOptixDeferredRefreshAt <> 0 Then Exit Sub

    g_AMIOptixDeferredRefreshAt = Now + TimeSerial(0, 0, 2)
    Application.OnTime EarliestTime:=g_AMIOptixDeferredRefreshAt, _
                        Procedure:="'" & ThisWorkbook.Name & "'!AMI_Optix_DoDeferredRefresh"
End Sub
