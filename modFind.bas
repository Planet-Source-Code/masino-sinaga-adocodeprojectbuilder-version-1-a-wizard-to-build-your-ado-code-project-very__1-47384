Attribute VB_Name = "modFind"
Option Explicit

Public Sub GenerateFindCode()
  DoEvents
  frmWizard.lblProcess.Caption = "Generating code for find form..."
  DoEvents
  
  Dim i As Integer, j As Integer, k As Integer
  strFileName = fldr & "\frmFind.frm"
  
  frmProcess.rtfLap1.Text = ""
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "Version 5.00" & vbCrLf & _
    "Begin VB.Form frmFind" & vbCrLf & _
    "   BorderStyle = 1        'Fixed Single" & vbCrLf & _
    "   Caption = ""Find""" & vbCrLf & _
    "   ClientHeight = 2055" & vbCrLf & _
    "   ClientLeft = 3090" & vbCrLf & _
    "   ClientTop = 6150" & vbCrLf & _
    "   ClientWidth = 5700" & vbCrLf & _
    "   LinkTopic = ""Form1""" & vbCrLf & _
    "   LockControls = -1       'True" & vbCrLf & _
    "   MaxButton = 0           'False" & vbCrLf & _
    "   MinButton = 0           'False" & vbCrLf & _
    "   ScaleHeight = 2055" & vbCrLf & _
    "   ScaleWidth = 5700" & vbCrLf & _
    "   StartUpPosition = 2    'CenterScreen" & vbCrLf & _
    "   Begin VB.CheckBox chkMatch" & vbCrLf & _
    "      Caption = ""&Match whole word only""" & vbCrLf & _
    "      Height = 255" & vbCrLf & _
    "      Left = 120" & vbCrLf & _
    "      TabIndex = 8" & vbCrLf & _
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
    "Begin VB.CheckBox chkKonfirmasi" & vbCrLf & _
    "   Caption = ""&Display the complete data in found record""" & vbCrLf & _
    "   Height = 255" & vbCrLf & _
    "   Left = 120" & vbCrLf & _
    "   TabIndex = 7" & vbCrLf & _
    "   Top = 1680" & vbCrLf & _
    "   Value = 1              'Checked" & vbCrLf & _
    "   Width = 3855" & vbCrLf & _
    "End" & vbCrLf & _
    "Begin VB.CommandButton cmdCancel" & vbCrLf & _
    "   Caption = ""&Cancel""" & vbCrLf & _
    "   Height = 375" & vbCrLf & _
    "   Left = 4320" & vbCrLf & _
    "   TabIndex = 6" & vbCrLf & _
    "   Top = 1200" & vbCrLf & _
    "   Width = 1215" & vbCrLf & _
    "End" & vbCrLf & _
    "Begin VB.CommandButton cmdFindNext" & vbCrLf & _
    "   Caption = ""Find &Next""" & vbCrLf & _
    "   Height = 375" & vbCrLf & _
    "   Left = 4320" & vbCrLf & _
    "   TabIndex = 5" & vbCrLf & _
    "   Top = 600" & vbCrLf & _
    "   Width = 1215" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "    End" & vbCrLf & _
    "    Begin VB.CommandButton cmdFindFirst" & vbCrLf & _
    "      Caption = ""Find &First""" & vbCrLf & _
    "      Height = 375" & vbCrLf & _
    "      Left = 4320" & vbCrLf & _
    "      TabIndex = 4" & vbCrLf & _
    "      Top = 120" & vbCrLf & _
    "      Width = 1215" & vbCrLf & _
    "   End" & vbCrLf & _
    "   Begin VB.ComboBox cboFind" & vbCrLf & _
    "      Height = 315" & vbCrLf & _
    "      Left = 1200" & vbCrLf & _
    "      TabIndex = 3" & vbCrLf & _
    "      Top = 840" & vbCrLf & _
    "      Width = 3015" & vbCrLf & _
    "   End" & vbCrLf & _
    "   Begin VB.ComboBox cboField" & vbCrLf & _
    "      Height = 315" & vbCrLf & _
    "      Left = 1200" & vbCrLf & _
    "      Style = 2              'Dropdown List" & vbCrLf & _
    "      TabIndex = 1" & vbCrLf & _
    "      Top = 360" & vbCrLf & _
    "      Width = 3015" & vbCrLf & _
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
    "   Begin VB.Label Label2" & vbCrLf & _
    "      BackStyle = 0          'Transparent" & vbCrLf & _
    "      Caption = ""Find what:""" & vbCrLf & _
    "      Height = 255" & vbCrLf & _
    "      Left = 120" & vbCrLf & _
    "      TabIndex = 2" & vbCrLf & _
    "      Top = 840" & vbCrLf & _
    "      Width = 975" & vbCrLf & _
    "   End" & vbCrLf & _
    "   Begin VB.Label Label1" & vbCrLf & _
    "      BackStyle = 0          'Transparent" & vbCrLf & _
    "      Caption = ""Find in Field:""" & vbCrLf & _
    "      Height = 255" & vbCrLf & _
    "      Left = 120" & vbCrLf & _
    "      TabIndex = 0" & vbCrLf & _
    "      Top = 360" & vbCrLf & _
    "      Width = 975" & vbCrLf & _
    "   End" & vbCrLf & _
    "End" & vbCrLf & _
    "Attribute VB_Name = ""frmFind""" & vbCrLf & _
    "Attribute VB_GlobalNameSpace = False" & vbCrLf & _
    "Attribute VB_Creatable = False" & vbCrLf & _
    "Attribute VB_PredeclaredId = True" & vbCrLf & _
    "Attribute VB_Exposed = False"
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "'File Name  : frmFind.frm" & vbCrLf & _
    "'Description: Find first and find next data that you " & vbCrLf & _
    "'             specified; based on a selected fields" & vbCrLf & _
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
    "'------------------------------------------------" & vbCrLf & _
    "" & vbCrLf & _
    "Option Explicit" & vbCrLf & _
    "" & vbCrLf & _
    "Dim rs As ADODB.Recordset" & vbCrLf & _
    "Dim adoField1 As ADODB.Field" & vbCrLf & _
    "Dim mark As Variant, intCount As Integer, intPosition As Integer" & vbCrLf & _
    "Dim bFound As Boolean, bCancel As Boolean" & vbCrLf & _
    "Dim strFind As String, strFindNext As String, strResult As String" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "Private Sub cboField_Click()" & vbCrLf & _
    "  If cboField.Text = ""(All Fields)"" Then" & vbCrLf & _
    "     chkMatch.Value = 0" & vbCrLf & _
    "     chkMatch.Enabled = False" & vbCrLf & _
    "  Else" & vbCrLf & _
    "     chkMatch.Enabled = True" & vbCrLf & _
    "  End If" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf & _
    "Private Sub cboFind_Change()" & vbCrLf & _
    "  If Len(Trim(cboFind.Text)) > 0 Then" & vbCrLf & _
    "     cmdFindFirst.Enabled = True" & vbCrLf & _
    "     cmdFindFirst.Default = True" & vbCrLf & _
    "  Else" & vbCrLf & _
    "     cmdFindFirst.Enabled = False" & vbCrLf & _
    "     cmdFindNext.Enabled = False" & vbCrLf & _
    "  End If" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "Private Sub cboFind_Click()" & vbCrLf & _
    "  If Len(Trim(cboFind.Text)) > 0 Then" & vbCrLf & _
    "     cmdFindFirst.Enabled = True" & vbCrLf & _
    "     cmdFindFirst.Default = True" & vbCrLf & _
    "  Else" & vbCrLf & _
    "     cmdFindFirst.Enabled = False" & vbCrLf & _
    "  End If" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "Private Sub cboField_KeyPress(KeyAscii As Integer)" & vbCrLf & _
    "  If KeyAscii = 13 Then" & vbCrLf & _
    "     cboFind.SetFocus" & vbCrLf & _
    "     SendKeys ""{Home}+{End}""" & vbCrLf & _
    "  End If" & vbCrLf & _
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
    "Private Sub cmdFindFirst_Click()" & vbCrLf & _
    "Dim strFound As String" & vbCrLf & _
    "Dim i As Integer" & vbCrLf & _
    "'If criteria is not (All Fields)" & vbCrLf & _
    "If Trim(cboField.Text) <> ""(All Fields)"" Then" & vbCrLf & _
    "  On Error GoTo Message" & vbCrLf & _
    "  intCount = 0" & vbCrLf & _
    "  CheckDouble" & vbCrLf & _
    "  adoFind.MoveFirst" & vbCrLf & _
    "  bFound = False 'Not found yet" & vbCrLf & _
    "  Do While adoFind.EOF <> True" & vbCrLf & _
    "     DoEvents" & vbCrLf & _
    "     If bCancel = True Then 'If use interrupt by clicking" & vbCrLf & _
    "                            'Cancel button..." & vbCrLf & _
    "        Exit Sub            '... exit from this procedure" & vbCrLf & _
    "     End If" & vbCrLf & _
    "     If chkMatch.Value = 0 Then  'Not match whole word" & vbCrLf & _
    "       If InStr(UCase(adoFind.Fields(cboField.Text)), UCase(cboFind.Text)) > 0 Then" & vbCrLf & _
    "          DoEvents" & vbCrLf & _
    "          intCount = intCount + 1" & vbCrLf & _
    "          DoEvents" & vbCrLf & _
    "          'Get the absolute position" & vbCrLf & _
    "          intPosition = adoFind.AbsolutePosition" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "          'We found it, update bFound now" & vbCrLf & _
    "          bFound = True" & vbCrLf & _
    "       End If" & vbCrLf & _
    "     Else 'Match whole word only" & vbCrLf & _
    "       If UCase(adoFind.Fields(cboField.Text)) = UCase(cboFind.Text) Then" & vbCrLf & _
    "          DoEvents" & vbCrLf & _
    "          intCount = intCount + 1" & vbCrLf & _
    "          DoEvents" & vbCrLf & _
    "          'Get the absolute position" & vbCrLf & _
    "          intPosition = adoFind.AbsolutePosition" & vbCrLf & _
    "          'We found it, update bFound now" & vbCrLf & _
    "          bFound = True" & vbCrLf & _
    "       End If" & vbCrLf & _
    "     End If" & vbCrLf & _
    "     If intCount = 1 Then 'If this is the first found" & vbCrLf & _
    "        bFound = True 'Update bFound" & vbCrLf & _
    "        Exit Do       'Exit from this looping, because" & vbCrLf & _
    "                      'this is only the first time" & vbCrLf & _
    "     End If" & vbCrLf & _
    "     DoEvents" & vbCrLf & _
    "     adoFind.MoveNext" & vbCrLf & _
    "  Loop" & vbCrLf & _
    "  'If we found and intCount <> 0" & vbCrLf & _
    "  If bFound = True And intCount <> 0 Then" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "     cmdFindNext.Enabled = True" & vbCrLf & _
    "     'Display what position we found..." & vbCrLf & _
    "     strFound = ""Found '"" & cboFind.Text & ""' in record number "" & adoFind.AbsolutePosition  " & vbCrLf & _
    "     'This will get the name of field" & vbCrLf & _
    "     For i = 0 To adoFind.Fields.Count - 1" & vbCrLf & _
    "       'Get just field name that we need, but ""ChildCMD""" & vbCrLf & _
    "       If adoFind.Fields(i).Name = ""ChildCMD"" Then" & vbCrLf & _
    "          Exit For" & vbCrLf & _
    "       End If" & vbCrLf & _
    "       'Get all data in record we found" & vbCrLf & _
    "       strFound = strFound & vbCrLf & _" & vbCrLf & _
    "            adoFind.Fields(i).Name & "": "" & _" & vbCrLf & _
    "            vbTab & adoFind.Fields(i).Value" & vbCrLf & _
    "     Next i" & vbCrLf & _
    "  End If" & vbCrLf & _
    "  'If chkKonfirmasi was checked by user and data found" & vbCrLf & _
    "  If chkKonfirmasi.Value = 1 And bFound = True Then" & vbCrLf & _
    "     'Display in messagebox" & vbCrLf & _
    "     MsgBox strFound, vbInformation, ""Found""" & vbCrLf & _
    "  End If" & vbCrLf & _
    "  If (adoFind.EOF) Then  'If pointer in end of recordset" & vbCrLf & _
    "     adoFind.MoveLast    'move to the last record" & vbCrLf & _
    "     bFound = False      'so, we haven't found it yet" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "     'Display messagebox we haven't found it" & vbCrLf & _
    "     MsgBox """" & cboFind.Text & "" not found "" & _ " & vbCrLf & _
    "            ""in field '"" & cboField.Text & ""'."", _ " & vbCrLf & _
    "            vbExclamation, ""Finished Searching""" & vbCrLf & _
    "     'cmdFindNext is not active because we haven't found" & vbCrLf & _
    "     'in cmdFindFirst" & vbCrLf & _
    "     cmdFindNext.Enabled = False" & vbCrLf & _
    "     Exit Sub" & vbCrLf & _
    "  End If" & vbCrLf & _
    "  Exit Sub" & vbCrLf & _
    "Else 'If user select (All Fields)" & vbCrLf & _
    "  FindFirstInAllFields '<-- call this procedure" & vbCrLf & _
    "  Exit Sub" & vbCrLf & _
    "End If" & vbCrLf & _
    "Message:" & vbCrLf & _
    "  MsgBox Err.Number & "" - "" & Err.Description" & vbCrLf & _
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
    "" & vbCrLf & _
    "Private Sub cmdFindNext_Click()" & vbCrLf & _
    "Dim strFound As String" & vbCrLf & _
    "Dim i As Integer" & vbCrLf & _
    "'If user select criteria: (All Fields)" & vbCrLf & _
    "If Trim(cboField.Text) <> ""(All Fields)"" Then" & vbCrLf & _
    "  On Error GoTo Message" & vbCrLf & _
    "  'First of all, we haven't found it, yet..." & vbCrLf & _
    "  bFound = False" & vbCrLf & _
    "  Do While adoFind.EOF <> True" & vbCrLf & _
    "     DoEvents" & vbCrLf & _
    "     If bCancel = True Then 'If use interrupt by clicking" & vbCrLf & _
    "                            'Cancel button..." & vbCrLf & _
    "        Exit Sub            '... exit from this procedure" & vbCrLf & _
    "     End If" & vbCrLf & _
    "" & vbCrLf & _
    "     If chkMatch.Value = 0 Then  'Not match whole word" & vbCrLf & _
    "       'In FindNext, we compare the intPosition variable" & vbCrLf & _
    "       'with AbsolutePosition. If they are not same" & vbCrLf & _
    "        'then we found it" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "       If (InStr(UCase(adoFind.Fields(cboField.Text)), _" & vbCrLf & _
    "              UCase(cboFind.Text)) > 0) And _" & vbCrLf & _
    "              intPosition <> adoFind.AbsolutePosition Then" & vbCrLf & _
    "          DoEvents" & vbCrLf & _
    "          'Update counter position" & vbCrLf & _
    "          intCount = intCount + 1" & vbCrLf & _
    "          DoEvents" & vbCrLf & _
    "          'Get the absolute position" & vbCrLf & _
    "          intPosition = adoFind.AbsolutePosition" & vbCrLf & _
    "          'We found it, update bFound now" & vbCrLf & _
    "          bFound = True" & vbCrLf & _
    "       End If" & vbCrLf & _
    "     Else 'Match whole word only" & vbCrLf & _
    "        If UCase(adoFind.Fields(cboField.Text)) = _" & vbCrLf & _
    "              UCase(cboFind.Text) And _" & vbCrLf & _
    "              intPosition <> adoFind.AbsolutePosition Then" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "          DoEvents " & vbCrLf & _
    "          'Update counter position" & vbCrLf & _
    "          intCount = intCount + 1" & vbCrLf & _
    "          DoEvents" & vbCrLf & _
    "          'Get the absolute position" & vbCrLf & _
    "          intPosition = adoFind.AbsolutePosition" & vbCrLf & _
    "          'We found it, update bFound now" & vbCrLf & _
    "          bFound = True" & vbCrLf & _
    "       End If" & vbCrLf & _
    "     End If" & vbCrLf & _
    "" & vbCrLf & _
    "     If bFound = True Then 'If we found it then" & vbCrLf & _
    "        Exit Do            'exit from this looping" & vbCrLf & _
    "     End If" & vbCrLf & _
    "     " & vbCrLf & _
    "     adoFind.MoveNext      'Process to next record" & vbCrLf & _
    "     DoEvents" & vbCrLf & _
    "     " & vbCrLf & _
    "     If adoFind.EOF Then   'If we are in EOF" & vbCrLf & _
    "        adoFind.MoveLast   'move to last record" & vbCrLf & _
    "        'Display message if we don't find it in looping" & vbCrLf & _
    "        MsgBox ""'"" & cboFind.Text & ""' not found "" & _" & vbCrLf & _
    "            ""in field '"" & cboField.Text & "" '."", _" & vbCrLf & _
    "            vbExclamation, ""Finished Searching""" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "        cmdFindNext.Enabled = False" & vbCrLf & _
    "        Exit Do" & vbCrLf & _
    "     End If" & vbCrLf & _
    "  Loop" & vbCrLf & _
    "  " & vbCrLf & _
    "  'If user check this checkbox" & vbCrLf & _
    "  If chkKonfirmasi.Value = 1 And _" & vbCrLf & _
    "     bFound = True And intCount <> 0 Then" & vbCrLf & _
    "     strFound = ""Found '"" & cboFind.Text & _" & vbCrLf & _
    "                ""' in record number "" & adoFind.AbsolutePosition " & vbCrLf & _
    "     'This iteration will get the name of all fields in" & vbCrLf & _
    "     'recordset, in order that we will display all data" & vbCrLf & _
    "     'in that record we found" & vbCrLf & _
    "     For i = 0 To adoFind.Fields.Count - 1" & vbCrLf & _
    "       'Check if the name contain ""ChildCMD"", exit from" & vbCrLf & _
    "       'iteration, we will not display this one." & vbCrLf & _
    "       If adoFind.Fields(i).Name = ""ChildCMD"" Then" & vbCrLf & _
    "          Exit For" & vbCrLf & _
    "       End If" & vbCrLf & _
    "       'This will keep all data in record we found" & vbCrLf & _
    "       strFound = strFound & vbCrLf & _" & vbCrLf & _
    "            adoFind.Fields(i).Name & "": "" & _ " & vbCrLf & _
    "            vbTab & adoFind.Fields(i).Value" & vbCrLf & _
    "     Next i" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "     'Show the complete data in messagebox" & vbCrLf & _
    "     MsgBox strFound, vbInformation, ""Found""" & vbCrLf & _
    "  End If" & vbCrLf & _
    "" & vbCrLf & _
    "  If (adoFind.EOF) Then" & vbCrLf & _
    "     adoFind.MoveLast" & vbCrLf & _
    "     bFound = False 'We haven't found it, yet" & vbCrLf & _
    "     'Show messagebox" & vbCrLf & _
    "     MsgBox ""'"" & cboFind.Text & ""' not found "" & _ " & vbCrLf & _
    "            ""in field '"" & cboField.Text & ""'."", _" & vbCrLf & _
    "            vbExclamation, ""Finished Searching""" & vbCrLf & _
    "     cmdFindNext.Enabled = False" & vbCrLf & _
    "     Exit Sub" & vbCrLf & _
    "  End If" & vbCrLf & _
    "  Exit Sub" & vbCrLf & _
    "Else 'If user select (All Fields)" & vbCrLf & _
    "  FindNextInAllFields  '<-- Call this procedure" & vbCrLf & _
    "  Exit Sub" & vbCrLf & _
    "End If" & vbCrLf & _
    "Message:" & vbCrLf & _
    "     adoFind.MoveLast" & vbCrLf & _
    "     bFound = False" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "     MsgBox ""'"" & cboFind.Text & ""' not found "" & _" & vbCrLf & _
    "            ""in field '"" & cboField.Text & ""'."", _" & vbCrLf & _
    "            vbExclamation, ""Finished Searching""" & vbCrLf & _
    "     cmdFindNext.Enabled = False" & vbCrLf & _
    "     Exit Sub" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "" & vbCrLf & _
    "Private Sub cmdCancel_Click()" & vbCrLf & _
    "  bCancel = True" & vbCrLf & _
    "  bFound = False" & vbCrLf & _
    "  " & vbCrLf & _
    "  Set adoField1 = Nothing" & vbCrLf & _
    "  Set rs1 = Nothing" & vbCrLf & _
    "  Unload Me" & vbCrLf & _
    "  'Me.Hide  'Just hide this form, in order that we still" & vbCrLf & _
    "           'need the data later" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "" & vbCrLf
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
    "  bCancel = False" & vbCrLf & _
    "  If cboFind.Text = """" Then" & vbCrLf & _
    "     cmdFindFirst.Enabled = False" & vbCrLf & _
    "     cmdFindNext.Enabled = False" & vbCrLf & _
    "  Else" & vbCrLf & _
    "     cmdFindFirst.Enabled = True" & vbCrLf & _
    "     cmdFindNext.Enabled = False" & vbCrLf & _
    "  End If" & vbCrLf & _
    "  Set rs1 = New ADODB.Recordset" & vbCrLf & _
    "  rs1.Open m_SQLRS1, cnn, adOpenKeyset, adLockOptimistic" & vbCrLf & _
    "  cboField.Clear" & vbCrLf & _
    "  cboField.AddItem ""(All Fields)""" & vbCrLf & _
    "  'This will get field name" & vbCrLf & _
    "  For Each adoField1 In rs1.Fields" & vbCrLf & _
    "      cboField.AddItem adoField1.Name" & vbCrLf & _
    "  Next" & vbCrLf & _
    "  rs1.Close" & vbCrLf & _
    "  'Highlight the first item in combobox" & vbCrLf & _
    "  cboField.Text = cboField.List(0)" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "  'Get setting for this form from INI File" & vbCrLf & _
    "  Call ReadFromINIToControls(frmFind, ""Find"")" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "" & vbCrLf & _
    "Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)" & vbCrLf & _
    "  'Save setting this form to INI File" & vbCrLf & _
    "  Call SaveFromControlsToINI(frmFind, ""Find"")" & vbCrLf & _
    "  'Clear memory" & vbCrLf & _
    "  Set adoFind = Nothing" & vbCrLf & _
    "  Set adoField1 = Nothing" & vbCrLf & _
    "  Screen.MousePointer = vbDefault" & vbCrLf & _
    "  Unload Me" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "" & vbCrLf & _
    "Private Sub CheckDouble()" & vbCrLf & _
    "Dim i As Integer" & vbCrLf & _
    "  If cboFind.Text = """" Then" & vbCrLf & _
    "     MsgBox ""It can't not be a empty string!"", _" & vbCrLf & _
    "            vbExclamation, ""Invalid""" & vbCrLf & _
    "     cboFind.SetFocus" & vbCrLf & _
    "     Exit Sub" & vbCrLf & _
    "  End If" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "  For i = 0 To cboFind.ListCount - 1" & vbCrLf & _
    "    If cboFind.List(i) = cboFind.Text Then" & vbCrLf & _
    "       cboFind.SetFocus" & vbCrLf & _
    "       SendKeys ""{Home}+{End}""" & vbCrLf & _
    "       Exit Sub" & vbCrLf & _
    "    End If" & vbCrLf & _
    "  Next i" & vbCrLf & _
    "  cboFind.AddItem cboFind.Text" & vbCrLf & _
    "  cboFind.Text = cboFind.List(cboFind.ListCount - 1)" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "" & vbCrLf & _
    "'This will search data in all fields for the very first time" & vbCrLf & _
    "Private Sub FindFirstInAllFields()" & vbCrLf & _
    "Dim strstrResult As String, strFound As String" & vbCrLf & _
    "Dim i As Integer, j As Integer, k As Integer" & vbCrLf & _
    "  'Always start from first record" & vbCrLf & _
    "  adoFind.MoveFirst" & vbCrLf & _
    "  strFind = cboFind.Text" & vbCrLf & _
    "  CheckDouble" & vbCrLf & _
    "Ulang:" & vbCrLf & _
    "  If adoFind.EOF And adoFind.RecordCount > 0 Then" & vbCrLf & _
    "     adoFind.MoveLast" & vbCrLf & _
    "     MsgBox """" & strFind & "" not found in '"" & cboField.Text & ""'."", _" & vbCrLf & _
    "            vbExclamation, ""Finished Searching""" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "     cmdFindNext.Enabled = False" & vbCrLf & _
    "     Exit Sub" & vbCrLf & _
    "  End If" & vbCrLf & _
    "  strstrResult = """":  strFound = """"" & vbCrLf & _
    "  With " & glo.strFormName & "" & vbCrLf & _
    "  For i = 0 To " & glo.intNumOfFields - 1 & "  'This iteration for data in textbox" & vbCrLf & _
    "      strResult = UCase(.txtFields(i).Text)" & vbCrLf & _
    "      If InStr(1, UCase(.txtFields(i).Text), UCase(strFind)) > 0 Then" & vbCrLf & _
    "         strstrResult = """" & strstrResult & "" Found '"" & strFind & ""' at:"" & vbCrLf & _" & vbCrLf & _
    "                      """ & vbCrLf & _
    "       For j = 0 To " & glo.intNumOfFields - 1 & " 'This iteration for data in datagrid" & vbCrLf & _
    "          strResult = UCase(.txtFields(j).Text)" & vbCrLf & _
    "          If InStr(1, UCase(.txtFields(j).Text), UCase(strFind)) > 0 Then" & vbCrLf & _
    "             strFindNext = strFind" & vbCrLf & _
    "             'If we found it, tell user which position" & vbCrLf & _
    "             'it is..." & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "              strstrResult = strstrResult & vbCrLf & _" & vbCrLf & _
    "                 ""  Record number "" & CStr(adoFind.AbsolutePosition) & """" & vbCrLf & _" & vbCrLf & _
    "                 ""  - Field name: "" & .txtFields(j).DataField & """" & vbCrLf & _" & vbCrLf & _
    "                 ""  - Contains: "" & .txtFields(j).Text & """" & vbCrLf & _" & vbCrLf & _
    "                 ""  - Column number: "" & j + 1 & "" in DataGrid.""" & vbCrLf & _
    "             For k = 0 To adoFind.Fields.Count - 1" & vbCrLf & _
    "                 If adoFind.Fields(k).Name = ""ChildCMD"" Then" & vbCrLf & _
    "                    Exit For" & vbCrLf & _
    "                 End If" & vbCrLf & _
    "                 strFound = strFound & vbCrLf & _" & vbCrLf & _
    "                         adoFind.Fields(k).Name & "": "" & _" & vbCrLf & _
    "                         vbTab & adoFind.Fields(k).Value" & vbCrLf & _
    "             Next k" & vbCrLf & _
    "             'Because we found, make cmdFindNext active..." & vbCrLf & _
    "             cmdFindNext.Enabled = True" & vbCrLf & _
    "             'If chkKonfirmasi was checked by user" & vbCrLf & _
    "             If chkKonfirmasi.Value = 1 Then" & vbCrLf & _
    "                'Display data" & vbCrLf & _
    "                 MsgBox strstrResult & vbCrLf & _" & vbCrLf & _
    "                        strFound, _" & vbCrLf & _
    "                        vbInformation, ""Found""" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "             End If" & vbCrLf & _
    "          Else" & vbCrLf & _
    "          End If" & vbCrLf & _
    "       Next j  'End of iteration in datagrid" & vbCrLf & _
    "       Exit Sub" & vbCrLf & _
    "    Else" & vbCrLf & _
    "    End If" & vbCrLf & _
    "  Next i  'End of iteration in textBox" & vbCrLf & _
    "  End With" & vbCrLf & _
    "  'If we don't find in first record, move to next record" & vbCrLf & _
    "  adoFind.MoveNext" & vbCrLf & _
    "  GoTo Ulang" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "" & vbCrLf & _
    "" & vbCrLf & _
    "'This will search data from the record position" & vbCrLf & _
    "'we found in FindFirstInAllFields procedure above." & vbCrLf & _
    "'" & vbCrLf & _
    "Private Sub FindNextInAllFields()" & vbCrLf & _
    "Dim m As Integer, n As Integer, k As Integer" & vbCrLf & _
    "Dim strstrResult As String, strFound As String" & vbCrLf & _
    "strFindNext = strFind" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'SELALU BERMASALAH........
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "If Len(Trim(strResult)) = 0 Then" & vbCrLf & _
    "   FindFirstInAllFields" & vbCrLf & _
    "   Exit Sub" & vbCrLf & _
    "End If" & vbCrLf & _
    "'Start from record position we found in FindFirstInAllFields" & vbCrLf & _
    "adoFind.MoveNext" & vbCrLf & _
    "strFound = """": strstrResult = """"" & vbCrLf & _
    "Ulang:" & vbCrLf & _
    "  'If we don't find it" & vbCrLf & _
    "  If adoFind.EOF And adoFind.RecordCount > 0 Then" & vbCrLf & _
    "     adoFind.MoveLast" & vbCrLf & _
    "     MsgBox """" & strFindNext & "" not found in '"" & cboField.Text & ""'."", _" & vbCrLf & _
    "            vbExclamation, ""Finished Searching""" & vbCrLf & _
    "     Exit Sub" & vbCrLf & _
    "  End If" & vbCrLf & _
    "  With " & glo.strFormName & "" & vbCrLf & _
    "  For n = 0 To " & glo.intNumOfFields - 1 & "  'This iteration for textbox" & vbCrLf & _
    "    strResult = UCase(cboFind.Text)" & vbCrLf & _
    "    'If we found it, all or similiar to it" & vbCrLf & _
    "    If InStr(1, UCase(.txtFields(n).Text), UCase(strFindNext)) > 0 Then" & vbCrLf & _
    "       strstrResult = strstrResult & ""Found '"" & strFindNext & ""' at:" & vbCrLf & _
    "                      " & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "       For m = 0 To " & glo.intNumOfFields - 1 & " 'This iteration for datagrid" & vbCrLf & _
    "          strResult = UCase(cboFind.Text)" & vbCrLf & _
    "          If InStr(1, UCase(.txtFields(m).Text), UCase(strFindNext)) > 0 Then" & vbCrLf & _
    "             'If we found, tell user which record position" & vbCrLf & _
    "             'it is.." & vbCrLf & _
    "             strstrResult = strstrResult & vbCrLf & _ " & vbCrLf & _
    "                 ""  Record number "" & CStr(adoFind.AbsolutePosition) & """" & vbCrLf & _" & vbCrLf & _
    "                 ""  - Field name: "" & .txtFields(m).DataField & """" & vbCrLf & _" & vbCrLf & _
    "                 ""  - Contains: "" & .txtFields(m).Text & """" & vbCrLf & _" & vbCrLf & _
    "                 ""  - Column number: "" & m + 1 & "" in DataGrid.""" & vbCrLf & _
    "             For k = 0 To adoFind.Fields.Count - 1" & vbCrLf & _
    "                If adoFind.Fields(k).Name = ""ChildCMD"" Then" & vbCrLf & _
    "                  Exit For" & vbCrLf & _
    "               End If" & vbCrLf & _
    "               'Get all data we found in that record" & vbCrLf & _
    "               strFound = strFound & vbCrLf & _" & vbCrLf & _
    "                         adoFind.Fields(k).Name & "": "" & _" & vbCrLf & _
    "                         vbTab & adoFind.Fields(k).Value" & vbCrLf & _
    "             Next k" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "             'If chkKonfirmasi was checked by user" & vbCrLf & _
    "             If chkKonfirmasi.Value = 1 Then" & vbCrLf & _
    "                 'Display all data in that record we found" & vbCrLf & _
    "                 MsgBox strstrResult & vbCrLf & _" & vbCrLf & _
    "                        strFound, _" & vbCrLf & _
    "                        vbInformation, ""Found""" & vbCrLf & _
    "                 cmdFindNext.Enabled = True" & vbCrLf & _
    "             End If" & vbCrLf & _
    "             Exit Sub" & vbCrLf & _
    "          Else" & vbCrLf & _
    "          End If" & vbCrLf & _
    "       Next m  'End of iteration in DataGrid" & vbCrLf & _
    "       Exit Sub" & vbCrLf & _
    "    Else" & vbCrLf & _
    "    End If" & vbCrLf & _
    "  Next n  'End of iteration in TextBox" & vbCrLf & _
    "  End With" & vbCrLf & _
    "  adoFind.MoveNext" & vbCrLf & _
    "  GoTo Ulang" & vbCrLf & _
    "" & vbCrLf & _
    "End Sub" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
  
 End Sub

'After we save it to file, put it to richtexbox control
Public Sub PutCodeToRichTextBox()
  Open strFileName For Input As #1
    frmProcess.rtfLap1.Text = Input(LOF(1), 1)
  Close #1
End Sub

