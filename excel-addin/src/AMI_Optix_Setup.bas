Attribute VB_Name = "AMI_Optix_Setup"
'===============================================================================
' AMI OPTIX - One-Time Setup Module
' Run SetupUtilitiesForm() once to add controls to frmUtilities
'===============================================================================
Option Explicit

Public Sub SetupUtilitiesForm()
    ' Adds the required controls to frmUtilities
    ' Run this ONCE after importing the add-in

    Dim frm As Object
    Dim ctrl As MSForms.Control
    Dim topPos As Single

    On Error GoTo ErrorHandler

    ' Get the form
    Set frm = frmUtilities

    ' Check if controls already exist
    On Error Resume Next
    Set ctrl = frm.Controls("cboElectricity")
    On Error GoTo ErrorHandler

    If Not ctrl Is Nothing Then
        MsgBox "Form controls already exist. Setup not needed.", vbInformation, "AMI Optix Setup"
        Exit Sub
    End If

    ' We can't add controls at runtime to a UserForm that's already compiled
    ' So instead, show manual instructions

    MsgBox "Please add the following controls manually in the VBA Editor:" & vbCrLf & vbCrLf & _
           "1. Open frmUtilities in design mode (double-click it)" & vbCrLf & _
           "2. From Toolbox, add 4 Labels and 4 ComboBoxes" & vbCrLf & _
           "3. Name the ComboBoxes:" & vbCrLf & _
           "   - cboElectricity" & vbCrLf & _
           "   - cboCooking" & vbCrLf & _
           "   - cboHeat" & vbCrLf & _
           "   - cboHotWater" & vbCrLf & _
           "4. Add 2 buttons named btnSave and btnCancel" & vbCrLf & vbCrLf & _
           "Or use the alternative: Run CreateSimpleUtilityForm", _
           vbInformation, "AMI Optix Setup"
    Exit Sub

ErrorHandler:
    MsgBox "Setup error: " & Err.Description, vbCritical, "AMI Optix Setup"
End Sub

Public Sub CreateSimpleUtilityForm()
    ' Creates a brand new form with all controls
    ' This uses VBE extensibility - requires trust access to VBA project

    Dim vbProj As Object
    Dim vbComp As Object
    Dim frm As Object
    Dim code As String

    On Error GoTo ErrorHandler

    ' Check for VBE access
    On Error Resume Next
    Set vbProj = ThisWorkbook.VBProject
    If vbProj Is Nothing Then
        MsgBox "Cannot access VBA Project. Please enable:" & vbCrLf & vbCrLf & _
               "File > Options > Trust Center > Trust Center Settings > " & vbCrLf & _
               "Macro Settings > Trust access to the VBA project object model", _
               vbExclamation, "AMI Optix Setup"
        Exit Sub
    End If
    On Error GoTo ErrorHandler

    ' Check if form exists and remove it
    On Error Resume Next
    vbProj.VBComponents.Remove vbProj.VBComponents("frmUtilities")
    On Error GoTo ErrorHandler

    ' Add new UserForm
    Set vbComp = vbProj.VBComponents.Add(3) ' vbext_ct_MSForm = 3
    vbComp.Name = "frmUtilities"

    Set frm = vbComp.Designer

    ' Set form properties
    frm.Caption = "Utility Configuration"
    frm.Width = 360
    frm.Height = 300

    ' Add controls
    AddLabel frm, "lblTitle", "Select Utility Payment Responsibilities", 12, 12, 330, 18, True

    AddLabel frm, "lblElectricity", "Electricity:", 12, 48, 100, 18, False
    AddComboBox frm, "cboElectricity", 120, 45, 220

    AddLabel frm, "lblCooking", "Cooking:", 12, 78, 100, 18, False
    AddComboBox frm, "cboCooking", 120, 75, 220

    AddLabel frm, "lblHeat", "Heat:", 12, 108, 100, 18, False
    AddComboBox frm, "cboHeat", 120, 105, 220

    AddLabel frm, "lblHotWater", "Hot Water:", 12, 138, 100, 18, False
    AddComboBox frm, "cboHotWater", 120, 135, 220

    AddButton frm, "btnSave", "Save", 120, 200, 72, 24
    AddButton frm, "btnCancel", "Cancel", 200, 200, 72, 24
    AddButton frm, "btnReset", "Reset", 12, 200, 96, 24

    ' Add the code
    code = GetFormCode()
    vbComp.CodeModule.DeleteLines 1, vbComp.CodeModule.CountOfLines
    vbComp.CodeModule.AddFromString code

    MsgBox "frmUtilities has been created successfully!" & vbCrLf & vbCrLf & _
           "Please save the add-in (Ctrl+S) and restart Excel.", _
           vbInformation, "AMI Optix Setup"

    Exit Sub

ErrorHandler:
    MsgBox "Error creating form: " & Err.Description & vbCrLf & vbCrLf & _
           "You may need to manually add controls to frmUtilities.", _
           vbCritical, "AMI Optix Setup"
End Sub

Private Sub AddLabel(frm As Object, ctrlName As String, caption As String, _
                     Left As Single, Top As Single, Width As Single, Height As Single, _
                     isBold As Boolean)
    Dim lbl As MSForms.Label
    Set lbl = frm.Controls.Add("Forms.Label.1", ctrlName)
    lbl.caption = caption
    lbl.Left = Left
    lbl.Top = Top
    lbl.Width = Width
    lbl.Height = Height
    If isBold Then lbl.Font.Bold = True
End Sub

Private Sub AddComboBox(frm As Object, ctrlName As String, _
                        Left As Single, Top As Single, Width As Single)
    Dim cbo As MSForms.ComboBox
    Set cbo = frm.Controls.Add("Forms.ComboBox.1", ctrlName)
    cbo.Left = Left
    cbo.Top = Top
    cbo.Width = Width
    cbo.Height = 20
    cbo.ColumnCount = 2
    cbo.BoundColumn = 1
    cbo.ColumnWidths = "200;0"
    cbo.Style = fmStyleDropDownList
End Sub

Private Sub AddButton(frm As Object, ctrlName As String, caption As String, _
                      Left As Single, Top As Single, Width As Single, Height As Single)
    Dim btn As MSForms.CommandButton
    Set btn = frm.Controls.Add("Forms.CommandButton.1", ctrlName)
    btn.caption = caption
    btn.Left = Left
    btn.Top = Top
    btn.Width = Width
    btn.Height = Height
End Sub

Private Function GetFormCode() As String
    Dim c As String
    c = "Option Explicit" & vbCrLf & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    With cboElectricity" & vbCrLf
    c = c & "        .Clear: .AddItem ""Tenant Pays"": .AddItem ""N/A or owner pays""" & vbCrLf
    c = c & "        .List(0, 1) = ""tenant_pays"": .List(1, 1) = ""na""" & vbCrLf
    c = c & "    End With" & vbCrLf
    c = c & "    With cboCooking" & vbCrLf
    c = c & "        .Clear: .AddItem ""Electric Stove"": .AddItem ""Gas Stove"": .AddItem ""N/A or owner pays""" & vbCrLf
    c = c & "        .List(0, 1) = ""electric"": .List(1, 1) = ""gas"": .List(2, 1) = ""na""" & vbCrLf
    c = c & "    End With" & vbCrLf
    c = c & "    With cboHeat" & vbCrLf
    c = c & "        .Clear: .AddItem ""Electric ccASHP"": .AddItem ""Electric Other"": .AddItem ""Gas Heat"": .AddItem ""Oil Heat"": .AddItem ""N/A or owner pays""" & vbCrLf
    c = c & "        .List(0, 1) = ""electric_ccashp"": .List(1, 1) = ""electric_other"": .List(2, 1) = ""gas"": .List(3, 1) = ""oil"": .List(4, 1) = ""na""" & vbCrLf
    c = c & "    End With" & vbCrLf
    c = c & "    With cboHotWater" & vbCrLf
    c = c & "        .Clear: .AddItem ""Electric Heat Pump"": .AddItem ""Electric Other"": .AddItem ""Gas"": .AddItem ""Oil"": .AddItem ""N/A or owner pays""" & vbCrLf
    c = c & "        .List(0, 1) = ""electric_heat_pump"": .List(1, 1) = ""electric_other"": .List(2, 1) = ""gas"": .List(3, 1) = ""oil"": .List(4, 1) = ""na""" & vbCrLf
    c = c & "    End With" & vbCrLf
    c = c & "    LoadSavedValues" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf

    c = c & "Private Sub LoadSavedValues()" & vbCrLf
    c = c & "    SelectByValue cboElectricity, GetSetting(""AMI_Optix"", ""Utilities"", ""electricity"", ""na"")" & vbCrLf
    c = c & "    SelectByValue cboCooking, GetSetting(""AMI_Optix"", ""Utilities"", ""cooking"", ""na"")" & vbCrLf
    c = c & "    SelectByValue cboHeat, GetSetting(""AMI_Optix"", ""Utilities"", ""heat"", ""na"")" & vbCrLf
    c = c & "    SelectByValue cboHotWater, GetSetting(""AMI_Optix"", ""Utilities"", ""hot_water"", ""na"")" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf

    c = c & "Private Sub SelectByValue(cbo As ComboBox, value As String)" & vbCrLf
    c = c & "    Dim i As Long" & vbCrLf
    c = c & "    For i = 0 To cbo.ListCount - 1" & vbCrLf
    c = c & "        If cbo.Column(1, i) = value Then cbo.ListIndex = i: Exit Sub" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "    cbo.ListIndex = cbo.ListCount - 1" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf

    c = c & "Private Sub btnSave_Click()" & vbCrLf
    c = c & "    Dim e$, c$, h$, w$" & vbCrLf
    c = c & "    If cboElectricity.ListIndex >= 0 Then e = cboElectricity.Column(1) Else e = ""na""" & vbCrLf
    c = c & "    If cboCooking.ListIndex >= 0 Then c = cboCooking.Column(1) Else c = ""na""" & vbCrLf
    c = c & "    If cboHeat.ListIndex >= 0 Then h = cboHeat.Column(1) Else h = ""na""" & vbCrLf
    c = c & "    If cboHotWater.ListIndex >= 0 Then w = cboHotWater.Column(1) Else w = ""na""" & vbCrLf
    c = c & "    SaveUtilitySelections e, c, h, w" & vbCrLf
    c = c & "    MsgBox ""Utility settings saved."", vbInformation, ""AMI Optix""" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf

    c = c & "Private Sub btnCancel_Click(): Unload Me: End Sub" & vbCrLf & vbCrLf

    c = c & "Private Sub btnReset_Click()" & vbCrLf
    c = c & "    cboElectricity.ListIndex = cboElectricity.ListCount - 1" & vbCrLf
    c = c & "    cboCooking.ListIndex = cboCooking.ListCount - 1" & vbCrLf
    c = c & "    cboHeat.ListIndex = cboHeat.ListCount - 1" & vbCrLf
    c = c & "    cboHotWater.ListIndex = cboHotWater.ListCount - 1" & vbCrLf
    c = c & "End Sub" & vbCrLf

    GetFormCode = c
End Function
