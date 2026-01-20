Attribute VB_Name = "AMI_Optix_EventHooks"
'===============================================================================
' AMI OPTIX - Application Event Hooks
' Keeps the "Scenario Manual" block synced with UAP edits.
'===============================================================================
Option Explicit

Public g_AMIOptixAppEvents As AMI_Optix_AppEvents
Public g_AMIOptixSuppressEvents As Boolean

Public Sub Auto_Open()
    StartAMIOptixEventHooks
End Sub

Public Sub Auto_Close()
    StopAMIOptixEventHooks
End Sub

Public Sub StartAMIOptixEventHooks()
    On Error Resume Next
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

