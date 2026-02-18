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
