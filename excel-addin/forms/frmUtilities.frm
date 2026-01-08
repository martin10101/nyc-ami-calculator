VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUtilities
   Caption         =   "Utility Configuration"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5400
   OleObjectBlob   =   "frmUtilities.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmUtilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' AMI OPTIX - Utility Configuration Form
' Allows user to select utility payment responsibilities
'===============================================================================
Option Explicit

'-------------------------------------------------------------------------------
' FORM EVENTS
'-------------------------------------------------------------------------------

Private Sub UserForm_Initialize()
    ' Populate combo boxes with options

    ' Electricity options
    With cboElectricity
        .Clear
        .AddItem "Tenant Pays"
        .AddItem "N/A or owner pays"
        .List(0, 1) = "tenant_pays"
        .List(1, 1) = "na"
    End With

    ' Cooking options
    With cboCooking
        .Clear
        .AddItem "Electric Stove"
        .AddItem "Gas Stove"
        .AddItem "N/A or owner pays"
        .List(0, 1) = "electric"
        .List(1, 1) = "gas"
        .List(2, 1) = "na"
    End With

    ' Heat options
    With cboHeat
        .Clear
        .AddItem "Electric Heat - Cold Climate Air Source Heat Pump (ccASHP)"
        .AddItem "Electric Heat - Other"
        .AddItem "Gas Heat"
        .AddItem "Oil Heat"
        .AddItem "N/A or owner pays"
        .List(0, 1) = "electric_ccashp"
        .List(1, 1) = "electric_other"
        .List(2, 1) = "gas"
        .List(3, 1) = "oil"
        .List(4, 1) = "na"
    End With

    ' Hot Water options
    With cboHotWater
        .Clear
        .AddItem "Electric Hot Water - Heat Pump"
        .AddItem "Electric Hot Water - Other"
        .AddItem "Gas Hot Water"
        .AddItem "Oil Hot Water"
        .AddItem "N/A or owner pays"
        .List(0, 1) = "electric_heat_pump"
        .List(1, 1) = "electric_other"
        .List(2, 1) = "gas"
        .List(3, 1) = "oil"
        .List(4, 1) = "na"
    End With

    ' Load saved values
    LoadSavedValues
End Sub

Private Sub LoadSavedValues()
    ' Load previously saved utility selections from registry

    Dim electricity As String
    Dim cooking As String
    Dim heat As String
    Dim hot_water As String

    electricity = GetSetting("AMI_Optix", "Utilities", "electricity", "na")
    cooking = GetSetting("AMI_Optix", "Utilities", "cooking", "na")
    heat = GetSetting("AMI_Optix", "Utilities", "heat", "na")
    hot_water = GetSetting("AMI_Optix", "Utilities", "hot_water", "na")

    ' Select matching items
    SelectByValue cboElectricity, electricity
    SelectByValue cboCooking, cooking
    SelectByValue cboHeat, heat
    SelectByValue cboHotWater, hot_water
End Sub

Private Sub SelectByValue(cbo As ComboBox, value As String)
    ' Select combo box item by its hidden value
    Dim i As Long

    For i = 0 To cbo.ListCount - 1
        If cbo.Column(1, i) = value Then
            cbo.ListIndex = i
            Exit Sub
        End If
    Next i

    ' Default to last item (N/A) if not found
    cbo.ListIndex = cbo.ListCount - 1
End Sub

'-------------------------------------------------------------------------------
' BUTTON EVENTS
'-------------------------------------------------------------------------------

Private Sub btnSave_Click()
    ' Save selections and close

    Dim electricity As String
    Dim cooking As String
    Dim heat As String
    Dim hot_water As String

    ' Get selected values
    If cboElectricity.ListIndex >= 0 Then
        electricity = cboElectricity.Column(1, cboElectricity.ListIndex)
    Else
        electricity = "na"
    End If

    If cboCooking.ListIndex >= 0 Then
        cooking = cboCooking.Column(1, cboCooking.ListIndex)
    Else
        cooking = "na"
    End If

    If cboHeat.ListIndex >= 0 Then
        heat = cboHeat.Column(1, cboHeat.ListIndex)
    Else
        heat = "na"
    End If

    If cboHotWater.ListIndex >= 0 Then
        hot_water = cboHotWater.Column(1, cboHotWater.ListIndex)
    Else
        hot_water = "na"
    End If

    ' Save to registry
    SaveUtilitySelections electricity, cooking, heat, hot_water

    MsgBox "Utility settings saved." & vbCrLf & vbCrLf & _
           "These settings will be used for all future optimizations.", _
           vbInformation, "AMI Optix"

    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnReset_Click()
    ' Reset all to N/A (owner pays)
    cboElectricity.ListIndex = cboElectricity.ListCount - 1
    cboCooking.ListIndex = cboCooking.ListCount - 1
    cboHeat.ListIndex = cboHeat.ListCount - 1
    cboHotWater.ListIndex = cboHotWater.ListCount - 1
End Sub
