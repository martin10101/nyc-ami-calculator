Attribute VB_Name = "AMI_Optix_EventHooks"
'===============================================================================
' AMI OPTIX - Application Event Hooks
' Keeps the "Scenario Manual" block synced with UAP edits.
'===============================================================================
Option Explicit

Public g_AMIOptixAppEvents As AMI_Optix_AppEvents
Public g_AMIOptixSuppressEvents As Boolean
Public g_AMIOptixLiveSyncEnabled As Boolean

Private m_LiveSyncInitialized As Boolean

Public Sub EnsureLiveSyncInitialized()
    If m_LiveSyncInitialized Then Exit Sub

    Dim v As String
    v = GetSetting("AMI_Optix", "Settings", "LiveSyncEnabled", "1")
    g_AMIOptixLiveSyncEnabled = (Trim$(v) <> "0")

    m_LiveSyncInitialized = True
End Sub

Public Function GetLiveSyncEnabled() As Boolean
    EnsureLiveSyncInitialized
    GetLiveSyncEnabled = g_AMIOptixLiveSyncEnabled
End Function

Public Sub SetLiveSyncEnabled(enabled As Boolean)
    EnsureLiveSyncInitialized
    g_AMIOptixLiveSyncEnabled = enabled
    SaveSetting "AMI_Optix", "Settings", "LiveSyncEnabled", IIf(enabled, "1", "0")
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
    Set g_AMIOptixAppEvents = New AMI_Optix_AppEvents
    Set g_AMIOptixAppEvents.App = Application
    On Error GoTo 0
End Sub

Public Sub StopAMIOptixEventHooks()
    On Error Resume Next
    If Not g_AMIOptixAppEvents Is Nothing Then
        Set g_AMIOptixAppEvents.App = Nothing
    End If
    Set g_AMIOptixAppEvents = Nothing
    On Error GoTo 0
End Sub
