Attribute VB_Name = "modSort"
Option Explicit

Public Sub GenerateSortCode()
  DoEvents
  frmWizard.lblProcess.Caption = "Generating code for sort form..."
  DoEvents
  
  Dim i As Integer, j As Integer, k As Integer
  strFileName = fldr & "\frmSort.frm"
  
  frmProcess.rtfLap1.Text = ""
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "Version 5.00" & vbCrLf & _
    "Begin VB.Form frmSort" & vbCrLf & _
    "   BorderStyle = 1        'Fixed Single" & vbCrLf & _
    "   Caption = ""Sort""" & vbCrLf & _
    "   ClientHeight = 1740" & vbCrLf & _
    "   ClientLeft = 45" & vbCrLf & _
    "   ClientTop = 330" & vbCrLf & _
    "   ClientWidth = 5670" & vbCrLf & _
    "   LinkTopic = ""Form1""" & vbCrLf & _
    "   LockControls = -1       'True" & vbCrLf & _
    "   MaxButton = 0           'False" & vbCrLf & _
    "   MinButton = 0           'False" & vbCrLf & _
    "   ScaleHeight = 1740" & vbCrLf & _
    "   ScaleWidth = 5670" & vbCrLf & _
    "   StartUpPosition = 2    'CenterScreen" & vbCrLf & _
    "   Begin VB.CommandButton cmdCancel" & vbCrLf & _
    "      Caption = ""&Cancel""" & vbCrLf & _
    "      Height = 375" & vbCrLf & _
    "      Left = 4320" & vbCrLf & _
    "      TabIndex = 3" & vbCrLf & _
    "      Top = 600" & vbCrLf & _
    "      Width = 1215" & vbCrLf & _
    "   End" & vbCrLf & _
    "   Begin VB.CommandButton cmdSort" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "      Caption = ""&Sort""" & vbCrLf & _
    "      Height = 375" & vbCrLf & _
    "      Left = 4320" & vbCrLf & _
    "      TabIndex = 2" & vbCrLf & _
    "      Top = 120" & vbCrLf & _
    "      Width = 1215" & vbCrLf & _
    "   End" & vbCrLf & _
    "   Begin VB.ComboBox cboSort" & vbCrLf & _
    "      Height = 315" & vbCrLf & _
    "      Left = 1200" & vbCrLf & _
    "      Style = 2              'Dropdown List" & vbCrLf & _
    "      TabIndex = 1" & vbCrLf & _
    "      Top = 840" & vbCrLf & _
    "      Width = 3015" & vbCrLf & _
    "   End" & vbCrLf & _
    "   Begin VB.ComboBox cboField" & vbCrLf & _
    "      Height = 315" & vbCrLf & _
    "      Left = 1200" & vbCrLf & _
    "      Style = 2              'Dropdown List" & vbCrLf & _
    "      TabIndex = 0" & vbCrLf & _
    "      Top = 360" & vbCrLf & _
    "      Width = 3015" & vbCrLf & _
    "   End" & vbCrLf & _
    "   Begin VB.Label Label2" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "      BackStyle = 0          'Transparent" & vbCrLf & _
    "      Caption = ""Sort Type:""" & vbCrLf & _
    "      Height = 255" & vbCrLf & _
    "      Left = 120" & vbCrLf & _
    "      TabIndex = 5" & vbCrLf & _
    "      Top = 840" & vbCrLf & _
    "      Width = 975" & vbCrLf & _
    "   End" & vbCrLf & _
    "   Begin VB.Label Label1" & vbCrLf & _
    "      BackStyle = 0          'Transparent" & vbCrLf & _
    "      Caption = ""Sort in Field:""" & vbCrLf & _
    "      Height = 255" & vbCrLf & _
    "      Left = 120" & vbCrLf & _
    "      TabIndex = 4" & vbCrLf & _
    "      Top = 360" & vbCrLf & _
    "      Width = 975" & vbCrLf & _
    "   End" & vbCrLf & _
    "End" & vbCrLf & _
    "Attribute VB_Name = ""frmSort""" & vbCrLf & _
    "Attribute VB_GlobalNameSpace = False" & vbCrLf & _
    "Attribute VB_Creatable = False" & vbCrLf & _
    "Attribute VB_PredeclaredId = True" & vbCrLf & _
    "Attribute VB_Exposed = False" & vbCrLf & _
    "'File Name  : frmSort.frm"
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "'Description: Sort recordset that you specified" & vbCrLf & _
    "'             based on a selected fields or" & vbCrLf & _
    "'             based on (All Fields)." & vbCrLf & _
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
    "Private Sub cboField_KeyPress(KeyAscii As Integer)" & vbCrLf & _
    "  If KeyAscii = 13 Then SendKeys ""{Tab}""" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "Private Sub cboSort_KeyPress(KeyAscii As Integer)" & vbCrLf & _
    "  If KeyAscii = 13 Then SendKeys ""{Tab}""" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "Private Sub cmdCancel_Click()" & vbCrLf & _
    "  'Clear memory from object variable" & vbCrLf & _
    "  Set adoField1 = Nothing" & vbCrLf & _
    "  Set rs1 = Nothing" & vbCrLf & _
    "  Unload Me" & vbCrLf & _
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
    "Private Sub cmdSort_Click()" & vbCrLf & _
    "Dim TipeSort As String" & vbCrLf & _
    "   If cboSort.Text = cboSort.List(0) Then" & vbCrLf & _
    "      TipeSort = ""ASC""" & vbCrLf & _
    "   Else" & vbCrLf & _
    "      TipeSort = ""DESC""" & vbCrLf & _
    "   End If" & vbCrLf
    
    If frmWizard.lstFormLayout.Text = frmWizard.lstFormLayout.List(0) Then
     frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
     "   Set adoSort = New ADODB.Recordset" & vbCrLf & _
     "   adoSort.Open ""SHAPE "" & _" & vbCrLf & _
     "     ""{SELECT * FROM "" & m_RecordSource1 & "" "" & _" & vbCrLf & _
     "     ""ORDER BY "" & cboField.Text & "" "" & TipeSort & ""} "" & _" & vbCrLf & _
     "     ""AS ParentCMD APPEND "" & _" & vbCrLf & _
     "     ""({SELECT * FROM "" & m_RecordSource1 & "" "" & _" & vbCrLf & _
     "     ""ORDER BY "" & cboField.Text & "" "" & TipeSort & ""} "" & _" & vbCrLf & _
     "     ""AS ChildCMD RELATE "" & m_FieldKey1 & "" TO "" & m_FieldKey1 & "" ) "" & _" & vbCrLf & _
     "     ""AS ChildCMD"", _" & vbCrLf & _
     "     cnn, adOpenStatic, adLockOptimistic" & vbCrLf
    
    Else 'Master/Detail
    
     frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
     "Dim strCriteria1 As String, strField1 As String" & vbCrLf & _
     "Dim strCriteria2 As String, strField2 As String" & vbCrLf & _
     "Dim intField1 As Integer, intField2 As String" & vbCrLf & _
     "Dim tabField1() As String, i As Integer, j As Integer" & vbCrLf & _
     "Dim tabField2() As String" & vbCrLf & _
     "  Set rs1 = New ADODB.Recordset" & vbCrLf & _
     "  Set rs2 = New ADODB.Recordset" & vbCrLf & _
     "  rs1.Open m_SQLRS1, cnn, adOpenKeyset, adLockOptimistic" & vbCrLf & _
     "  rs2.Open m_SQLRS2, cnn, adOpenKeyset, adLockOptimistic" & vbCrLf & _
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
     "     If i <> intField1 - 1 Then " & vbCrLf & _
     "        strField1 = strField1 & tabField1(i) & "","" " & vbCrLf & _
     "     Else" & vbCrLf & _
     "        strField1 = strField1 & tabField1(i) & """ & vbCrLf & _
     "     End If " & vbCrLf & _
     "  Next i " & vbCrLf
     'detail
     frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
     "  For i = 0 To intField2 - 1 " & vbCrLf & _
     "     If i <> intField2 - 1 Then " & vbCrLf & _
     "        strField2 = strField2 & tabField2(i) & "","" " & vbCrLf & _
     "     Else" & vbCrLf & _
     "        strField2 = strField2 & tabField2(i) & """ & vbCrLf & _
     "     End If " & vbCrLf & _
     "  Next i " & vbCrLf
          
     frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
     "  Set adoSort = New ADODB.Recordset " & vbCrLf & _
     "     adoSort.Open _" & vbCrLf & _
     "     ""SHAPE "" & _ " & vbCrLf & _
     "     ""{SELECT "" & strField1 & "" FROM "" & m_RecordSource1 & "" "" & _" & vbCrLf & _
     "     ""ORDER BY "" & cboField.Text & "" "" & TipeSort & ""} "" & _" & vbCrLf & _
     "     ""AS ParentCMD APPEND "" & _" & vbCrLf & _
     "     ""({SELECT "" & strField2 & "" FROM "" & m_RecordSource2 & "" "" & _" & vbCrLf & _
     "     ""ORDER BY "" & m_FieldKey2 & "" "" & TipeSort & ""} "" & _" & vbCrLf & _
     "     ""AS ChildCMD RELATE "" & m_FieldKey1 & "" TO "" & m_FieldKey2 & "") "" & _" & vbCrLf & _
     "     ""AS ChildCMD"", cnn, adOpenStatic, adLockOptimistic" & vbCrLf
    End If
    
    
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "   With " & glo.strFormName & "" & vbCrLf & _
    "     On Error Resume Next" & vbCrLf & _
    "     If adoSort.RecordCount > 0 Then" & vbCrLf
    
    If frmWizard.lstFormLayout.Text = frmWizard.lstFormLayout.List(0) Then
     frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
     "        Set .grdDataGrid.DataSource = adoSort.DataSource" & vbCrLf
    Else 'Master/Detail
     frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
     "        Set .grdDataGrid.DataSource = adoSort(""ChildCMD"").UnderlyingValue" & vbCrLf
     '"        Set .grdDataGrid.DataSource = adoSort.DataSource" & vbCrLf
    End If
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "        Set .rsstrFindData = adoSort.DataSource" & vbCrLf & _
    "        Dim oTextData As TextBox" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "        For Each oTextData In .txtFields" & vbCrLf & _
    "            Set oTextData.DataSource = adoSort.DataSource" & vbCrLf & _
    "        Next" & vbCrLf & _
    "        Set .adoPrimaryRS = adoSort" & vbCrLf & _
    "        .cmdFirst.Value = True" & vbCrLf & _
    "     End If" & vbCrLf & _
    "   End With" & vbCrLf & _
    "   Exit Sub" & vbCrLf & _
    "Message:" & vbCrLf & _
    "     MsgBox Err.Number & "" - "" & _" & vbCrLf & _
    "            Err.Description, _" & vbCrLf & _
    "            vbExclamation, ""No Result""" & vbCrLf & _
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
    "Private Sub Form_Load()" & vbCrLf & _
    "On Error Resume Next" & vbCrLf & _
    "  Set rs1 = New ADODB.Recordset" & vbCrLf & _
    "  rs1.Open m_SQLRS1, cnn, adOpenKeyset, adLockOptimistic" & vbCrLf & _
    "  cboField.Clear" & vbCrLf & _
    "  For Each adoField1 In rs1.Fields" & vbCrLf & _
    "      cboField.AddItem adoField1.Name" & vbCrLf & _
    "  Next" & vbCrLf & _
    "  rs1.Close" & vbCrLf & _
    "  cboField.Text = cboField.List(0)" & vbCrLf & _
    "  cboSort.AddItem ""Ascending (ASC)""" & vbCrLf & _
    "  cboSort.AddItem ""Descending (DESC)""" & vbCrLf & _
    "  cboSort.Text = cboSort.List(0)" & vbCrLf & _
    "  'Get setting for this form from INI File" & vbCrLf & _
    "  Call ReadFromINIToControls(frmSort, ""Sort"")" & vbCrLf & _
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
    "Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)" & vbCrLf & _
    "  'Save setting this form to INI File" & vbCrLf & _
    "  Call SaveFromControlsToINI(frmSort, ""Sort"")" & vbCrLf & _
    "  'Clear memory" & vbCrLf & _
    "  Set adoSort = Nothing" & vbCrLf & _
    "  Set adoField1 = Nothing" & vbCrLf & _
    "  Screen.MousePointer = vbDefault" & vbCrLf & _
    "  Unload Me" & vbCrLf & _
    "End Sub" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1

End Sub

Public Function GetMyDateTime() As String
Dim aDay As Variant, sDay As String, sTime As String
  aDay = Array("Sunday", "Monday", "Tuesday", "Wednesday", _
               "Thursday", "Friday", "Saturday")
  sDay = aDay(Abs(Weekday(Date) - 1))
  GetMyDateTime = "" & sDay & ", " & _
                  Format(Date, "dd mmmm yyyy") & "; " & _
                  Format(Time, "hh:mm:ss")
End Function
