Attribute VB_Name = "AMI_Optix_Setup"
'===============================================================================
' AMI OPTIX - Setup Module (OPTIONAL)
' This module is NOT required for the add-in to work.
' It only contains helper functions for initial configuration.
'
' NOTE: You can safely DELETE this module from your VBA project.
' The add-in will work without it.
'===============================================================================
Option Explicit

Public Sub ShowSetupInstructions()
    ' Shows setup instructions for the add-in
    ' No form references - safe to compile

    MsgBox "AMI Optix Setup Instructions:" & vbCrLf & vbCrLf & _
           "1. Click 'API Settings' to enter your API key" & vbCrLf & _
           "2. Click 'Utilities' to configure utility settings" & vbCrLf & _
           "3. Open a workbook with unit data" & vbCrLf & _
           "4. Click 'Run Solver' to optimize" & vbCrLf & vbCrLf & _
           "For help, visit the web dashboard at:" & vbCrLf & _
           "https://nyc-ami-calculator.onrender.com", _
           vbInformation, "AMI Optix Setup"
End Sub

Public Sub ClearAllSettings()
    ' Clears all AMI Optix settings from registry
    ' Useful for troubleshooting

    On Error Resume Next

    ' Clear API key
    DeleteSetting "AMI_Optix", "Settings"

    ' Clear utility settings
    DeleteSetting "AMI_Optix", "Utilities"

    On Error GoTo 0

    MsgBox "All AMI Optix settings have been cleared." & vbCrLf & vbCrLf & _
           "You will need to re-enter your API key.", _
           vbInformation, "AMI Optix"
End Sub

Public Sub ShowDebugInfo()
    ' Shows current configuration for debugging
    Dim info As String

    info = "AMI Optix Debug Information:" & vbCrLf & vbCrLf
    info = info & "API Key: " & IIf(HasAPIKey(), "Configured", "NOT SET") & vbCrLf
    info = info & "API URL: " & API_BASE_URL & vbCrLf & vbCrLf

    info = info & "Utility Settings:" & vbCrLf
    info = info & "  Electricity: " & GetSetting("AMI_Optix", "Utilities", "electricity", "(not set)") & vbCrLf
    info = info & "  Cooking: " & GetSetting("AMI_Optix", "Utilities", "cooking", "(not set)") & vbCrLf
    info = info & "  Heat: " & GetSetting("AMI_Optix", "Utilities", "heat", "(not set)") & vbCrLf
    info = info & "  Hot Water: " & GetSetting("AMI_Optix", "Utilities", "hot_water", "(not set)") & vbCrLf

    MsgBox info, vbInformation, "AMI Optix Debug"
End Sub
