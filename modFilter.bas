Attribute VB_Name = "modFilter"
Option Explicit

Public Sub GenerateFilterCode()
  DoEvents
  frmWizard.lblProcess.Caption = "Generating code for filter form..."
  DoEvents
  Dim i As Integer, j As Integer, k As Integer
  strFileName = fldr & "\frmFilter.frm"
  
  frmProcess.rtfLap1.Text = ""
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "Version 5.00" & vbCrLf & _
    "Begin VB.Form frmFilter" & vbCrLf & _
    "   BorderStyle = 1        'Fixed Single" & vbCrLf & _
    "   Caption = ""Filter""" & vbCrLf & _
    "   ClientHeight = 1740" & vbCrLf & _
    "   ClientLeft = 3300" & vbCrLf & _
    "   ClientTop = 6255" & vbCrLf & _
    "   ClientWidth = 5670" & vbCrLf & _
    "   LinkTopic = ""Form1""" & vbCrLf & _
    "   LockControls = -1       'True" & vbCrLf & _
    "   MaxButton = 0           'False" & vbCrLf & _
    "   MinButton = 0           'False" & vbCrLf & _
    "   ScaleHeight = 1740" & vbCrLf & _
    "   ScaleWidth = 5670" & vbCrLf & _
    "   StartUpPosition = 2    'CenterScreen" & vbCrLf & _
    "   Begin VB.CheckBox chkMatch" & vbCrLf & _
    "      Caption = ""&Match whole word only""" & vbCrLf & _
    "      Height = 255" & vbCrLf & _
    "      Left = 120" & vbCrLf & _
    "      TabIndex = 6" & vbCrLf & _
    "      Top = 1320" & vbCrLf & _
    "      Width = 2175" & vbCrLf & _
    "   End" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "   Begin VB.ComboBox cboField" & vbCrLf & _
    "      Height = 315" & vbCrLf & _
    "      Left = 1200" & vbCrLf & _
    "      Style = 2              'Dropdown List" & vbCrLf & _
    "      TabIndex = 0" & vbCrLf & _
    "      Top = 360" & vbCrLf & _
    "      Width = 3015" & vbCrLf & _
    "   End" & vbCrLf & _
    "   Begin VB.ComboBox cboFilter" & vbCrLf & _
    "      Height = 315" & vbCrLf & _
    "      Left = 1200" & vbCrLf & _
    "      TabIndex = 1" & vbCrLf & _
    "      Top = 840" & vbCrLf & _
    "      Width = 3015" & vbCrLf & _
    "   End" & vbCrLf & _
    "   Begin VB.CommandButton cmdFilter" & vbCrLf & _
    "      Caption = ""&Filter""" & vbCrLf & _
    "      Height = 375" & vbCrLf & _
    "      Left = 4320" & vbCrLf & _
    "      TabIndex = 2" & vbCrLf & _
    "      Top = 120" & vbCrLf & _
    "      Width = 1215" & vbCrLf & _
    "   End" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "   Begin VB.CommandButton cmdCancel" & vbCrLf & _
    "      Caption = ""&Cancel""" & vbCrLf & _
    "      Height = 375" & vbCrLf & _
    "      Left = 4320" & vbCrLf & _
    "      TabIndex = 3" & vbCrLf & _
    "      Top = 600" & vbCrLf & _
    "      Width = 1215" & vbCrLf & _
    "   End" & vbCrLf & _
    "   Begin VB.Label Label1" & vbCrLf & _
    "      BackStyle = 0          'Transparent" & vbCrLf & _
    "      Caption = ""Filter in Field:""" & vbCrLf & _
    "      Height = 255" & vbCrLf & _
    "      Left = 120" & vbCrLf & _
    "      TabIndex = 5" & vbCrLf & _
    "      Top = 360" & vbCrLf & _
    "      Width = 975" & vbCrLf & _
    "   End" & vbCrLf & _
    "   Begin VB.Label Label2" & vbCrLf & _
    "      BackStyle = 0          'Transparent" & vbCrLf & _
    "      Caption = ""Filter What:""" & vbCrLf & _
    "      Height = 255" & vbCrLf & _
    "      Left = 120" & vbCrLf & _
    "      TabIndex = 4" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "      Top = 840" & vbCrLf & _
    "      Width = 975" & vbCrLf & _
    "   End" & vbCrLf & _
    "End" & vbCrLf & _
    "Attribute VB_Name = ""frmFilter""" & vbCrLf & _
    "Attribute VB_GlobalNameSpace = False" & vbCrLf & _
    "Attribute VB_Creatable = False" & vbCrLf & _
    "Attribute VB_PredeclaredId = True" & vbCrLf & _
    "Attribute VB_Exposed = False" & vbCrLf & _
    "'File Name  : frmFilter.frm" & vbCrLf & _
    "'Description: Filter data that you specified" & vbCrLf & _
    "'             based on a selected fields" & vbCrLf & _
    "'             or based on (All Fields)." & vbCrLf & _
    "'Copyrights : Masino Sinaga (masino_sinaga@yahoo.com)" & vbCrLf & _
    "'             http://www30.brinkster.com/masinosinaga/" & vbCrLf & _
    "'             http://www.geocities.com/masino_sinaga/" & vbCrLf & _
    "'             PLEASE DO NOT REMOVE THE COPYRIGHTS." & vbCrLf & _
    "'Author     : (put your name here)" & vbCrLf & _
    "'Web Site   : http://" & vbCrLf & _
    "'Created on : " & GetMyDateTime & "" & vbCrLf & _
    "'Modified   : ......." & vbCrLf & _
    "'Location   : (put your location here)" & vbCrLf & _
    "'------------------------------------------------" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "Option Explicit        'All variables that we use" & vbCrLf & _
    "                       'must be declared" & vbCrLf
    
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "" & vbCrLf & _
    "Private Sub cboField_Click()" & vbCrLf & _
    "  If cboField.Text = ""(All Fields)"" Then" & vbCrLf & _
    "     chkMatch.Value = 0" & vbCrLf & _
    "     chkMatch.Enabled = False" & vbCrLf & _
    "  Else" & vbCrLf & _
    "     chkMatch.Enabled = True" & vbCrLf & _
    "  End If" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "" & vbCrLf & _
    "Private Sub cboField_KeyPress(KeyAscii As Integer)" & vbCrLf & _
    "  If KeyAscii = 13 Then SendKeys ""{Tab}""" & vbCrLf & _
    "End Sub" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "Private Sub cboFilter_KeyPress(KeyAscii As Integer)" & vbCrLf & _
    "  If KeyAscii = 13 Then SendKeys ""{Tab}""" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "" & vbCrLf & _
    "'If there is a change in cboFilter..." & vbCrLf & _
    "Private Sub cboFilter_Change()" & vbCrLf & _
    "  If Len(Trim(cboFilter.Text)) > 0 Then" & vbCrLf & _
    "     'cmdFilter will be active and ready" & vbCrLf & _
    "     cmdFilter.Enabled = True" & vbCrLf & _
    "     cmdFilter.Default = True" & vbCrLf & _
    "  Else 'Still empty" & vbCrLf & _
    "     cmdFilter.Enabled = False 'We can't use it" & vbCrLf & _
    "  End If" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "" & vbCrLf & _
    "Private Sub cmdFilter_Click()" & vbCrLf & _
    "On Error GoTo Message" & vbCrLf & _
    " 'Assign recordset variable to new recordset" & vbCrLf & _
    "  Set adoFilter = New ADODB.Recordset" & vbCrLf & _
    " 'Filter recordset based on paramter in SQL Statement" & vbCrLf & _
    " AddCriteriaToCombo" & vbCrLf & _
    " If cboField.Text <> ""(All Fields)"" Then" & vbCrLf & _
    "   If chkMatch.Value = 0 Then 'Not match whole criteria word" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    If frmWizard.lstFormLayout.Text = frmWizard.lstFormLayout.List(0) Then
     frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
     "     adoFilter.Open ""SHAPE "" & _" & vbCrLf & _
     "     ""{SELECT * FROM "" & m_RecordSource1 & "" "" & _" & vbCrLf & _
     "     ""WHERE "" & Trim(cboField.Text) & "" "" & _" & vbCrLf & _
     "     ""LIKE '%"" & cboFilter.Text & ""%' "" & _" & vbCrLf & _
     "     ""ORDER BY "" & m_FieldKey1 & ""} AS ParentCMD APPEND "" & _" & vbCrLf & _
     "     ""({SELECT * FROM "" & m_RecordSource1 & "" "" & _" & vbCrLf & _
     "     ""WHERE "" & Trim(cboField.Text) & "" "" & _" & vbCrLf & _
     "     ""LIKE '%"" & cboFilter.Text & ""%' "" & _" & vbCrLf & _
     "     ""ORDER BY "" & m_FieldKey1 & ""} AS ChildCMD "" & _" & vbCrLf & _
     "     ""RELATE "" & m_FieldKey1 & "" TO "" & m_FieldKey1 & "") "" & _" & vbCrLf & _
     "     ""AS ChildCMD"", cnn, adOpenStatic, adLockOptimistic" & vbCrLf
    Else 'Master/Detail
     frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
     "     adoFilter.Open ""SHAPE "" & _" & vbCrLf & _
     "     ""{SELECT * FROM "" & m_RecordSource1 & "" "" & _" & vbCrLf & _
     "     ""WHERE "" & Trim(cboField.Text) & "" "" & _" & vbCrLf & _
     "     ""LIKE '%"" & cboFilter.Text & ""%' "" & _" & vbCrLf & _
     "     ""ORDER BY "" & m_FieldKey1 & ""} AS ParentCMD APPEND "" & _" & vbCrLf & _
     "     ""({SELECT * FROM "" & m_RecordSource2 & "" "" & _" & vbCrLf & _
     "     ""ORDER BY "" & m_FieldKey2 & ""} AS ChildCMD "" & _" & vbCrLf & _
     "     ""RELATE "" & m_FieldKey1 & "" TO "" & m_FieldKey2 & "") "" & _" & vbCrLf & _
     "     ""AS ChildCMD"", cnn, adOpenStatic, adLockOptimistic" & vbCrLf
     '     "     ""WHERE "" & Trim(cboField.Text) & "" "" & _" & vbCrLf & _
     "     ""LIKE '%"" & cboFilter.Text & ""%' "" & _" & vbCrLf & _

    End If
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    If frmWizard.lstFormLayout.Text = frmWizard.lstFormLayout.List(0) Then
     frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
     "   Else 'Match whole criteria word only" & vbCrLf & _
     "     adoFilter.Open ""SHAPE "" & _" & vbCrLf & _
     "     ""{SELECT * FROM "" & m_RecordSource1 & "" "" & _" & vbCrLf & _
     "     ""WHERE "" & Trim(cboField.Text) & "" "" & _" & vbCrLf & _
     "     ""= '"" & cboFilter.Text & ""' "" & _" & vbCrLf & _
     "     ""ORDER BY "" & m_FieldKey1 & ""} AS ParentCMD APPEND "" & _" & vbCrLf & _
     "     ""({SELECT * FROM "" & m_RecordSource1 & "" "" & _" & vbCrLf & _
     "     ""WHERE "" & Trim(cboField.Text) & "" "" & _" & vbCrLf & _
     "     ""= '"" & cboFilter.Text & ""' "" & _" & vbCrLf & _
     "     ""ORDER BY "" & m_FieldKey1 & ""} AS ChildCMD "" & _" & vbCrLf & _
     "     ""RELATE "" & m_FieldKey1 & "" TO "" & m_FieldKey1 & "") "" & _" & vbCrLf & _
     "     ""AS ChildCMD"", cnn, adOpenStatic, adLockOptimistic" & vbCrLf & _
     "   End If" & vbCrLf
    Else 'Master/Detail
     frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
     "   Else 'Match whole criteria word only" & vbCrLf & _
     "     adoFilter.Open ""SHAPE "" & _" & vbCrLf & _
     "     ""{SELECT * FROM "" & m_RecordSource1 & "" "" & _" & vbCrLf & _
     "     ""WHERE "" & Trim(cboField.Text) & "" "" & _" & vbCrLf & _
     "     ""= '"" & cboFilter.Text & ""' "" & _" & vbCrLf & _
     "     ""ORDER BY "" & m_FieldKey1 & ""} AS ParentCMD APPEND "" & _" & vbCrLf & _
     "     ""({SELECT * FROM "" & m_RecordSource2 & "" "" & _" & vbCrLf & _
     "     ""ORDER BY "" & m_FieldKey2 & ""} AS ChildCMD "" & _" & vbCrLf & _
     "     ""RELATE "" & m_FieldKey1 & "" TO "" & m_FieldKey2 & "") "" & _" & vbCrLf & _
     "     ""AS ChildCMD"", cnn, adOpenStatic, adLockOptimistic" & vbCrLf & _
     "   End If" & vbCrLf
     '     "     ""WHERE "" & Trim(cboField.Text) & "" "" & _" & vbCrLf & _
     "     ""= '"" & cboFilter.Text & ""' "" & _" & vbCrLf & _

    End If
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "   With " & glo.strFormName & "" & vbCrLf & _
    "     'If recordset is not empty" & vbCrLf & _
    "     If adoFilter.RecordCount > 0 Then" & vbCrLf & _
    "       'Display the result to datagrid" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    'If complete ADO Code.....
    If frmWizard.lstFormLayout.Text = frmWizard.lstFormLayout.List(0) Then
      frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
      "       'This will update the status label in" & vbCrLf & _
      "       'middle of navigation button" & vbCrLf & _
      "       Set .adoPrimaryRS = adoFilter" & vbCrLf & _
      "       Set .grdDataGrid.DataSource = adoFilter.DataSource" & vbCrLf
    Else 'Master/Detail
      frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
      "       'This will update the status label in" & vbCrLf & _
      "       'middle of navigation button" & vbCrLf & _
      "       Set .adoPrimaryRS = adoFilter" & vbCrLf & _
      "       Set .grdDataGrid.DataSource = .adoPrimaryRS(""ChildCMD"").UnderlyingValue" & vbCrLf
    End If
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "       'Bind the data to textbox" & vbCrLf & _
    "       Dim oTextData As TextBox" & vbCrLf & _
    "       For Each oTextData In .txtFields" & vbCrLf & _
    "           Set oTextData.DataSource = adoFilter.DataSource" & vbCrLf & _
    "       Next" & vbCrLf & _
    "       'Go to the first record" & vbCrLf & _
    "       .cmdFirst.Value = True" & vbCrLf & _
    "       .cmdBookmark.Enabled = False" & vbCrLf & _
    "       Set .adoPrimaryRS = adoFilter" & vbCrLf & _
    "     Else" & vbCrLf & _
    "       .cmdRefresh.Value = True" & vbCrLf & _
    "       MsgBox ""'"" & cboFilter.Text & ""' not found "" & _" & vbCrLf & _
    "              ""in field "" & cboField.Text & ""."", _" & vbCrLf & _
    "              vbExclamation, ""No Result""" & vbCrLf & _
    "     End If" & vbCrLf & _
    "   End With" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "   Exit Sub" & vbCrLf & _
    " Else" & vbCrLf & _
    "   FilterInAllFields" & vbCrLf & _
    "   Exit Sub" & vbCrLf & _
    " End If" & vbCrLf & _
    "Message:" & vbCrLf & _
    "  MsgBox ""'"" & cboFilter.Text & ""' not found "" & _" & vbCrLf & _
    "         ""in field '"" & cboField.Text & ""'."", _" & vbCrLf & _
    "         vbExclamation, ""No Result""" & vbCrLf & _
    "End Sub" & vbCrLf
    
    If frmWizard.lstFormLayout.Text = frmWizard.lstFormLayout.List(0) Then
     frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
     "Private Sub cmdCancel_Click()" & vbCrLf & _
     "  'Clear memory from object variable" & vbCrLf & _
     "  Set adoField1 = Nothing" & vbCrLf & _
     "  Set rs1 = Nothing" & vbCrLf & _
     "  Unload Me" & vbCrLf & _
     "End Sub" & vbCrLf & vbCrLf
    Else
     frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
     "Private Sub cmdCancel_Click()" & vbCrLf & _
     "  'Clear memory from object variable" & vbCrLf & _
     "  Set adoField1 = Nothing" & vbCrLf & _
     "  Set rs1 = Nothing" & vbCrLf & _
     "  Set adoField2 = Nothing" & vbCrLf & _
     "  Set rs2 = Nothing" & vbCrLf & _
     "  Unload Me" & vbCrLf & _
     "End Sub" & vbCrLf & vbCrLf
    End If
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "Private Sub Form_Load()" & vbCrLf & _
    "On Error Resume Next" & vbCrLf & _
    "  If cboFilter.Text = """" Then" & vbCrLf & _
    "     cmdFilter.Enabled = False" & vbCrLf & _
    "  Else 'If cboFilter is not empty" & vbCrLf & _
    "     cmdFilter.Enabled = True 'cmdFilter ready!" & vbCrLf & _
    "  End If" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    If frmWizard.lstFormLayout.Text = frmWizard.lstFormLayout.List(0) Then
     frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
     "  Set rs1 = New ADODB.Recordset" & vbCrLf & _
     "  rs1.Open m_SQLRS1, cnn, adOpenKeyset, adLockOptimistic" & vbCrLf & _
     "  cboField.Clear" & vbCrLf & _
     "  cboField.AddItem ""(All Fields)""" & vbCrLf & _
     "  For Each adoField1 In rs1.Fields" & vbCrLf & _
     "      cboField.AddItem adoField1.Name" & vbCrLf & _
     "  Next" & vbCrLf & _
     "  cboField.Text = cboField.List(0)" & vbCrLf & _
     "  'Get setting for this form from INI File" & vbCrLf & _
     "  Call ReadFromINIToControls(frmFilter, ""Filter"")" & vbCrLf & _
     "End Sub" & vbCrLf & vbCrLf & _
     "Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)" & vbCrLf & _
     "  'Save setting this form to INI File" & vbCrLf & _
     "  Call SaveFromControlsToINI(frmFilter, ""Filter"")" & vbCrLf & _
     "  'Clear memory" & vbCrLf & _
     "  Set adoFilter = Nothing" & vbCrLf & _
     "  Set adoField1 = Nothing" & vbCrLf & _
     "  Screen.MousePointer = vbDefault" & vbCrLf & _
     "  Unload Me" & vbCrLf & _
     "End Sub" & vbCrLf
    Else  'Master/Detail
     frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
     "  Set rs1 = New ADODB.Recordset" & vbCrLf & _
     "  Set rs2 = New ADODB.Recordset" & vbCrLf & _
     "  rs1.Open m_SQLRS1, cnn, adOpenKeyset, adLockOptimistic" & vbCrLf & _
     "  rs2.Open m_SQLRS2, cnn, adOpenKeyset, adLockOptimistic" & vbCrLf & _
     "  cboField.Clear" & vbCrLf & _
     "  cboField.AddItem ""(All Fields)""" & vbCrLf & _
     "  For Each adoField1 In rs1.Fields" & vbCrLf & _
     "      cboField.AddItem adoField1.Name" & vbCrLf & _
     "  Next" & vbCrLf & _
     "  cboField.Text = cboField.List(0)" & vbCrLf & _
     "  'Get setting for this form from INI File" & vbCrLf & _
     "  Call ReadFromINIToControls(frmFilter, ""Filter"")" & vbCrLf & _
     "End Sub" & vbCrLf & vbCrLf & _
     "Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)" & vbCrLf & _
     "  'Save setting this form to INI File" & vbCrLf & _
     "  Call SaveFromControlsToINI(frmFilter, ""Filter"")" & vbCrLf & _
     "  'Clear memory" & vbCrLf & _
     "  Set adoFilter = Nothing" & vbCrLf & _
     "  Set adoField1 = Nothing" & vbCrLf & _
     "  Set adoField2 = Nothing" & vbCrLf & _
     "  Screen.MousePointer = vbDefault" & vbCrLf & _
     "  Unload Me" & vbCrLf & _
     "End Sub" & vbCrLf
    End If
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    If frmWizard.lstFormLayout.Text = frmWizard.lstFormLayout.List(0) Then
     frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
     "Private Sub FilterInAllFields() " & vbCrLf & _
     "Dim strCriteria As String, strField As String" & vbCrLf & _
     "Dim intField As Integer, i As Integer, j As Integer" & vbCrLf & _
     "Dim tabField() As String " & vbCrLf & _
     "  rs1.MoveFirst " & vbCrLf & _
     "  strCriteria = "" " & vbCrLf & _
     "  intField = rs1.Fields.Count " & vbCrLf & _
     "  ReDim tabField(intField) " & vbCrLf & _
     "  intField = rs1.Fields.Count " & vbCrLf & _
     "  i = 0 " & vbCrLf & _
     "  For Each adoField1 In rs1.Fields " & vbCrLf & _
     "      tabField(i) = adoField1.Name " & vbCrLf & _
     "      i = i + 1 " & vbCrLf & _
     "  Next " & vbCrLf & _
     "  For i = 0 To intField - 1 " & vbCrLf & _
     "    If chkMatch.Value = 0 Then 'Not match whole criteria word " & vbCrLf & _
     "     If i <> intField - 1 Then " & vbCrLf & _
     "        strField = strField & tabField(i) & "","" " & vbCrLf & _
     "        strCriteria = strCriteria & _ " & vbCrLf & _
     "           tabField(i) & "" LIKE '%"" & cboFilter.Text & ""%' Or "" " & vbCrLf & _
     "     Else " & vbCrLf
    Else 'Master/Detail
     frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
     "Private Sub FilterInAllFields() " & vbCrLf & _
     "Dim strCriteria1 As String, strField1 As String" & vbCrLf & _
     "Dim strCriteria2 As String, strField2 As String" & vbCrLf & _
     "Dim intField1 As Integer, intField2 As String" & vbCrLf & _
     "Dim tabField1() As String, i As Integer, j As Integer" & vbCrLf & _
     "Dim tabField2() As String" & vbCrLf & _
     "  rs1.MoveFirst " & vbCrLf & _
     "  rs2.MoveFirst " & vbCrLf & _
     "  strCriteria1 = "" " & vbCrLf & _
     "  strCriteria2 = "" " & vbCrLf & _
     "  intField1 = rs1.Fields.Count " & vbCrLf & _
     "  intField2 = rs2.Fields.Count " & vbCrLf & _
     "  ReDim tabField1(intField1) " & vbCrLf & _
     "  ReDim tabField2(intField2) " & vbCrLf & _
     "  i = 0 " & vbCrLf
     frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
     "  For Each adoField1 In rs1.Fields " & vbCrLf & _
     "      tabField1(i) = adoField1.Name " & vbCrLf & _
     "      i = i + 1 " & vbCrLf & _
     "  Next" & vbCrLf & _
     "  i = 0 " & vbCrLf & _
     "  For Each adoField2 In rs2.Fields " & vbCrLf & _
     "      tabField2(i) = adoField2.Name " & vbCrLf & _
     "      i = i + 1 " & vbCrLf & _
     "  Next" & vbCrLf
     'master
     frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
     "  For i = 0 To intField1 - 1 " & vbCrLf & _
     "    If chkMatch.Value = 0 Then 'Not match whole criteria word " & vbCrLf & _
     "     If i <> intField1 - 1 Then " & vbCrLf & _
     "        strField1 = strField1 & tabField1(i) & "","" " & vbCrLf & _
     "        strCriteria1 = strCriteria1 & _ " & vbCrLf & _
     "           tabField1(i) & "" LIKE '%"" & cboFilter.Text & ""%' Or "" " & vbCrLf & _
     "     Else" & vbCrLf & _
     "        strField1 = strField1 & tabField1(i) & """ & vbCrLf & _
     "        strCriteria1 = strCriteria1 & tabField1(i) & "" LIKE '%"" & cboFilter.Text & ""%' """ & vbCrLf & _
     "     End If " & vbCrLf & _
     "    Else  'Match whole criteria word only " & vbCrLf & _
     "     If i <> intField1 - 1 Then " & vbCrLf & _
     "        strField1 = strField1 & tabField1(i) & "","" " & vbCrLf & _
     "        strCriteria1 = strCriteria1 & _ " & vbCrLf & _
     "           tabField1(i) & "" = '%"" & cboFilter.Text & ""%' Or " & vbCrLf & _
     "     Else " & vbCrLf & _
     "        strField1 = strField1 & tabField1(i) & """ & vbCrLf & _
     "        strCriteria1 = strCriteria1 & tabField1(i) & "" = '%"" & cboFilter.Text & ""%' """ & vbCrLf & _
     "     End If " & vbCrLf & _
     "  End If " & vbCrLf & _
     "  Next i " & vbCrLf
     'detail
     frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
     "  For i = 0 To intField2 - 1 " & vbCrLf & _
     "    If chkMatch.Value = 0 Then 'Not match whole criteria word " & vbCrLf & _
     "     If i <> intField2 - 1 Then " & vbCrLf & _
     "        strField2 = strField2 & tabField2(i) & "","" " & vbCrLf & _
     "        strCriteria2 = strCriteria2 & _ " & vbCrLf & _
     "           tabField2(i) & "" LIKE '%"" & cboFilter.Text & ""%' Or "" " & vbCrLf & _
     "     Else" & vbCrLf & _
     "        strField2 = strField2 & tabField2(i) & """ & vbCrLf & _
     "        strCriteria2 = strCriteria2 & tabField2(i) & "" LIKE '%"" & cboFilter.Text & ""%' """ & vbCrLf & _
     "     End If " & vbCrLf & _
     "    Else  'Match whole criteria word only " & vbCrLf & _
     "     If i <> intField2 - 1 Then " & vbCrLf & _
     "        strField2 = strField2 & tabField2(i) & "","" " & vbCrLf & _
     "        strCriteria2 = strCriteria2 & _ " & vbCrLf & _
     "           tabField2(i) & "" = '%"" & cboFilter.Text & ""%' Or " & vbCrLf & _
     "     Else " & vbCrLf & _
     "        strField2 = strField2 & tabField2(i) & """ & vbCrLf & _
     "        strCriteria2 = strCriteria2 & tabField2(i) & "" = '%"" & cboFilter.Text & ""%' """ & vbCrLf & _
     "     End If " & vbCrLf & _
     "  End If " & vbCrLf & _
     "  Next i " & vbCrLf
     frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
     "  Set adoFilter = New ADODB.Recordset " & vbCrLf & _
     "     adoFilter.Open _" & vbCrLf & _
     "     ""SHAPE "" & _ " & vbCrLf & _
     "     ""{SELECT "" & strField1 & "" FROM "" & m_RecordSource1 & "" "" & _" & vbCrLf & _
     "     ""WHERE "" & strCriteria1 & "" ORDER BY "" & m_FieldKey1 & ""} "" & _" & vbCrLf & _
     "     ""AS ParentCMD APPEND "" & _" & vbCrLf & _
     "     ""({SELECT "" & strField2 & "" FROM "" & m_RecordSource2 & "" "" & _" & vbCrLf & _
     "     "" ORDER BY "" & m_FieldKey2 & ""} "" & _" & vbCrLf & _
     "     ""AS ChildCMD RELATE "" & m_FieldKey1 & "" TO "" & m_FieldKey2 & "") "" & _" & vbCrLf & _
     "     ""AS ChildCMD"", cnn, adOpenStatic, adLockOptimistic" & vbCrLf
     '     "     ""WHERE "" & strCriteria2 & "" ORDER BY "" & m_FieldKey2 & ""} "" & _" & vbCrLf & _

    End If
    'WHERE "" & strCriteria2 & ""
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    If frmWizard.lstFormLayout.Text = frmWizard.lstFormLayout.List(0) Then
     frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
     "        strField = strField & tabField(i) & """ & vbCrLf & _
     "        strCriteria = strCriteria & tabField(i) & "" LIKE '%"" & cboFilter.Text & ""%' """ & vbCrLf & _
     "     End If " & vbCrLf & _
     "  Else  'Match whole criteria word only " & vbCrLf & _
     "     If i <> intField - 1 Then " & vbCrLf & _
     "        strField = strField & tabField(i) & "","" " & vbCrLf & _
     "        strCriteria = strCriteria & _ " & vbCrLf & _
     "           tabField(i) & "" = '%"" & cboFilter.Text & ""%' Or " & vbCrLf & _
     "     Else " & vbCrLf & _
     "        strField = strField & tabField(i) & """ & vbCrLf & _
     "        strCriteria = strCriteria & tabField(i) & "" = '%"" & cboFilter.Text & ""%' """ & vbCrLf & _
     "     End If " & vbCrLf & _
     "  End If " & vbCrLf & _
     "  Next i " & vbCrLf & _
     "  Set adoFilter = New ADODB.Recordset " & vbCrLf & _
     "     adoFilter.Open _" & vbCrLf & _
     "     ""SHAPE "" & _ " & vbCrLf & _
     "     ""{SELECT "" & strField & "" FROM "" & m_RecordSource1 & "" "" & _" & vbCrLf & _
     "     ""WHERE "" & strCriteria & "" ORDER BY "" & m_FieldKey1 & ""} "" & _" & vbCrLf & _
     "     ""AS ParentCMD APPEND "" & _" & vbCrLf & _
     "     ""({SELECT "" & strField & "" FROM "" & m_RecordSource1 & "" "" & _" & vbCrLf & _
     "     ""WHERE "" & strCriteria & "" ORDER BY "" & m_FieldKey1 & ""} "" & _" & vbCrLf & _
     "     ""AS ChildCMD RELATE "" & m_FieldKey1 & "" TO "" & m_FieldKey1 & "") "" & _" & vbCrLf & _
     "     ""AS ChildCMD"", cnn, adOpenStatic, adLockOptimistic" & vbCrLf
    Else 'Master/Detail
     'look at above...
    End If
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    If frmWizard.lstFormLayout.Text = frmWizard.lstFormLayout.List(0) Then
     frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
     "  With " & glo.strFormName & "" & vbCrLf & _
     "  If adoFilter.RecordCount > 0 Then" & vbCrLf & _
     "     Set .adoPrimaryRS = adoFilter" & vbCrLf & _
     "     Set .grdDataGrid.DataSource = adoFilter.DataSource" & vbCrLf & _
     "     Dim oTextData As TextBox" & vbCrLf & _
     "     For Each oTextData In .txtFields" & vbCrLf & _
     "         Set oTextData.DataSource = adoFilter.DataSource" & vbCrLf & _
     "     Next" & vbCrLf & _
     "     .cmdFirst.Value = True" & vbCrLf & _
     "     .cmdBookmark.Enabled = False" & vbCrLf & _
     "  Else" & vbCrLf & _
     "     .cmdRefresh.Value = True" & vbCrLf & _
     "     MsgBox ""'"" & cboFilter.Text & ""' not found "" & _" & vbCrLf & _
     "            ""in field '"" & cboField.Text & ""'."", _" & vbCrLf & _
     "            vbExclamation, ""No Result""" & vbCrLf & _
     "  End If" & vbCrLf & _
     "  End With" & vbCrLf & _
     "  Exit Sub" & vbCrLf & _
     "Message:" & vbCrLf & _
     "  'MsgBox Err.Number & "" - "" & Err.Description" & vbCrLf & _
     "  MsgBox ""'"" & cboFilter.Text & ""' not found "" & vbCrLf  & _" & vbCrLf & _
     "         ""in field '"" & cboField.Text & ""'."", _" & vbCrLf & _
     "         vbExclamation, ""No Result""" & vbCrLf & _
     "End Sub" & vbCrLf
    Else 'Master/Detail
     frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
     "  With " & glo.strFormName & "" & vbCrLf & _
     "  If adoFilter.RecordCount > 0 Then" & vbCrLf & _
     "     Set .adoPrimaryRS = adoFilter" & vbCrLf & _
     "     Set .grdDataGrid.DataSource = .adoPrimaryRS(""ChildCMD"").UnderlyingValue" & vbCrLf & _
     "     Dim oTextData As TextBox" & vbCrLf & _
     "     For Each oTextData In .txtFields" & vbCrLf & _
     "         Set oTextData.DataSource = adoFilter.DataSource" & vbCrLf & _
     "     Next" & vbCrLf & _
     "     .cmdFirst.Value = True" & vbCrLf & _
     "     .cmdBookmark.Enabled = False" & vbCrLf & _
     "  Else" & vbCrLf & _
     "     .cmdRefresh.Value = True" & vbCrLf & _
     "     MsgBox ""'"" & cboFilter.Text & ""' not found "" & _" & vbCrLf & _
     "            ""in field '"" & cboField.Text & ""'."", _" & vbCrLf & _
     "            vbExclamation, ""No Result""" & vbCrLf & _
     "  End If" & vbCrLf & _
     "  End With" & vbCrLf & _
     "  Exit Sub" & vbCrLf & _
     "Message:" & vbCrLf & _
     "  'MsgBox Err.Number & "" - "" & Err.Description" & vbCrLf & _
     "  MsgBox ""'"" & cboFilter.Text & ""' not found "" & vbCrLf  & _" & vbCrLf & _
     "         ""in field '"" & cboField.Text & ""'."", _" & vbCrLf & _
     "         vbExclamation, ""No Result""" & vbCrLf & _
     "End Sub" & vbCrLf
    End If
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "Private Sub AddCriteriaToCombo()" & vbCrLf & _
    "Dim i As Integer" & vbCrLf & _
    "  If cboFilter.Text = """" Then" & vbCrLf & _
    "     MsgBox ""Data is empty!"", _" & vbCrLf & _
    "            vbExclamation, ""Empty""" & vbCrLf & _
    "     cboFilter.SetFocus" & vbCrLf & _
    "     Exit Sub" & vbCrLf & _
    "  End If" & vbCrLf & _
    "  For i = 0 To cboFilter.ListCount - 1" & vbCrLf & _
    "    If cboFilter.List(i) = cboFilter.Text Then" & vbCrLf & _
    "       cboFilter.SetFocus" & vbCrLf & _
    "       SendKeys ""{Home}+{End}""" & vbCrLf & _
    "       Exit Sub" & vbCrLf & _
    "    End If" & vbCrLf & _
    "  Next i" & vbCrLf & _
    "  cboFilter.AddItem cboFilter.Text" & vbCrLf & _
    "  cboFilter.Text = cboFilter.List(cboFilter.ListCount - 1)" & vbCrLf & _
    "End Sub" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1

End Sub
