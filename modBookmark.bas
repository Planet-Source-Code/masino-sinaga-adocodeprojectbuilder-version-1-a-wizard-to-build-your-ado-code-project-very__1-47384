Attribute VB_Name = "modBookmark"
Option Explicit

Public Sub GenerateBookmarkCode()
  DoEvents
  frmWizard.lblProcess.Caption = "Generating code for bookmark form..."
  DoEvents
  
  Dim i As Integer, j As Integer, k As Integer
  strFileName = fldr & "\frmBookmark.frm"
  
  frmProcess.rtfLap1.Text = ""
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "Version 5.00" & vbCrLf & _
    "Begin VB.Form frmBookmark" & vbCrLf & _
    "   BorderStyle = 1        'Fixed Single" & vbCrLf & _
    "   Caption = ""Bookmark""" & vbCrLf & _
    "   ClientHeight = 4215" & vbCrLf & _
    "   ClientLeft = 8175" & vbCrLf & _
    "   ClientTop = 3600" & vbCrLf & _
    "   ClientWidth = 3210" & vbCrLf & _
    "   LinkTopic = ""Form1""" & vbCrLf & _
    "   LockControls = -1       'True" & vbCrLf & _
    "   MaxButton = 0           'False" & vbCrLf & _
    "   MinButton = 0           'False" & vbCrLf & _
    "   ScaleHeight = 4215" & vbCrLf & _
    "   ScaleWidth = 3210" & vbCrLf & _
    "   Begin VB.CommandButton cmdButton" & vbCrLf & _
    "      Caption = ""&Help""" & vbCrLf & _
    "      Height = 375" & vbCrLf & _
    "      Index = 4" & vbCrLf & _
    "      Left = 1080" & vbCrLf & _
    "      TabIndex = 7" & vbCrLf & _
    "      ToolTipText = ""What is bookmark? How do I use it?""" & vbCrLf & _
    "      Top = 3720" & vbCrLf & _
    "      Width = 975" & vbCrLf & _
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
    "   Begin VB.CommandButton cmdButton" & vbCrLf & _
    "      Caption = ""&Cancel""" & vbCrLf & _
    "      Height = 375" & vbCrLf & _
    "      Index = 3" & vbCrLf & _
    "      Left = 2160" & vbCrLf & _
    "      TabIndex = 6" & vbCrLf & _
    "      ToolTipText = ""Finish with bookmark""" & vbCrLf & _
    "      Top = 3720" & vbCrLf & _
    "      Width = 975" & vbCrLf & _
    "   End" & vbCrLf & _
    "   Begin VB.CommandButton cmdButton" & vbCrLf & _
    "      Caption = ""&Jump""" & vbCrLf & _
    "      Height = 375" & vbCrLf & _
    "      Index = 2" & vbCrLf & _
    "      Left = 2160" & vbCrLf & _
    "      TabIndex = 5" & vbCrLf & _
    "      ToolTipText = ""Go to record which bookmark name selected""" & vbCrLf & _
    "      Top = 3240" & vbCrLf & _
    "      Width = 975" & vbCrLf & _
    "   End" & vbCrLf & _
    "   Begin VB.CommandButton cmdButton" & vbCrLf & _
    "      Caption = ""&Delete""" & vbCrLf & _
    "      Height = 375" & vbCrLf & _
    "      Index = 1" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "      Left = 1080" & vbCrLf & _
    "      TabIndex = 4" & vbCrLf & _
    "      ToolTipText = ""Delete the selected bookmark""" & vbCrLf & _
    "      Top = 3240" & vbCrLf & _
    "      Width = 975" & vbCrLf & _
    "   End" & vbCrLf & _
    "   Begin VB.CommandButton cmdButton" & vbCrLf & _
    "      Caption = ""&Add""" & vbCrLf & _
    "      Height = 375" & vbCrLf & _
    "      Index = 0" & vbCrLf & _
    "      Left = 50" & vbCrLf & _
    "      TabIndex = 3" & vbCrLf & _
    "      ToolTipText = ""Add new bookmark""" & vbCrLf & _
    "      Top = 3240" & vbCrLf & _
    "      Width = 975" & vbCrLf & _
    "   End" & vbCrLf & _
    "   Begin VB.ListBox lstBookmark" & vbCrLf & _
    "      Height = 2400" & vbCrLf & _
    "      Left = 50" & vbCrLf & _
    "      TabIndex = 2" & vbCrLf & _
    "      ToolTipText = ""Double click the name to go to its record""" & vbCrLf & _
    "      Top = 720" & vbCrLf & _
    "      Width = 3135" & vbCrLf & _
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
    "   Begin VB.TextBox txtBookmark" & vbCrLf & _
    "      Height = 285" & vbCrLf & _
    "      Left = 50" & vbCrLf & _
    "      MaxLength = 30" & vbCrLf & _
    "      TabIndex = 0" & vbCrLf & _
    "      ToolTipText = ""Enter bookmark name here""" & vbCrLf & _
    "      Top = 360" & vbCrLf & _
    "      Width = 3135" & vbCrLf & _
    "   End" & vbCrLf & _
    "   Begin VB.Label Label1" & vbCrLf & _
    "      Caption = ""Bookmark name:""" & vbCrLf & _
    "      Height = 255" & vbCrLf & _
    "      Left = 50" & vbCrLf & _
    "      TabIndex = 1" & vbCrLf & _
    "      Top = 120" & vbCrLf & _
    "      Width = 2415" & vbCrLf & _
    "   End" & vbCrLf & _
    "End" & vbCrLf & _
    "Attribute VB_Name = ""frmBookmark""" & vbCrLf & _
    "Attribute VB_GlobalNameSpace = False" & vbCrLf & _
    "Attribute VB_Creatable = False" & vbCrLf & _
    "Attribute VB_PredeclaredId = True" & vbCrLf & _
    "Attribute VB_Exposed = False" & vbCrLf & _
    "'File Name  : frmBookmark.frm"
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "'Description: Mark a record in recordset so we can" & vbCrLf & _
    "'             go back to that record later without" & vbCrLf & _
    "'             remember the position of record.." & vbCrLf & _
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
    "'User Defined Type (UDT) for bookmark" & vbCrLf & _
    "Private Type arrMark" & vbCrLf & _
    "   AbsolutePosition As Double" & vbCrLf & _
    "   BookmarkName As String * 30" & vbCrLf & _
    "   BookmarkNumber As Variant" & vbCrLf & _
    "End Type" & vbCrLf & _
    "" & vbCrLf & _
    "'Declare dynamic array with arrMark types" & vbCrLf & _
    "Dim tabMark() As arrMark" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "'This is procedure to give mark to a record" & vbCrLf & _
    "Private Sub GiveMark()" & vbCrLf & _
    "On Error GoTo Message" & vbCrLf & _
    "'Static, so we can increase this variable as long as" & vbCrLf & _
    "'program stay in memory even we declare them in procedure" & vbCrLf & _
    "Static intNumber As Integer" & vbCrLf & _
    "'Add to array each time user add a new bookmark" & vbCrLf & _
    "ReDim Preserve tabMark(UBound(tabMark) + 1)" & vbCrLf & _
    "'Update counter variable intNumber" & vbCrLf & _
    "intNumber = intNumber + 1" & vbCrLf & _
    "  'Get information for this bookmark we added" & vbCrLf & _
    "  tabMark(intNumber).AbsolutePosition = adoBookMark.AbsolutePosition" & vbCrLf & _
    "  tabMark(intNumber).BookmarkNumber = adoBookMark.Bookmark" & vbCrLf & _
    "  tabMark(intNumber).BookmarkName = txtBookmark.Text" & vbCrLf & _
    "  Exit Sub" & vbCrLf & _
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
    "Private Sub cmdButton_Click(Index As Integer)" & vbCrLf & _
    "  Select Case Index" & vbCrLf & _
    "         Case 0  'Add button clicked" & vbCrLf & _
    "              Dim i As Integer" & vbCrLf & _
    "              For i = 0 To lstBookmark.ListCount - 1" & vbCrLf & _
    "                If lstBookmark.List(i) = txtBookmark.Text Then" & vbCrLf & _
    "                   MsgBox ""This bookmark name already exist in the list!"" & vbCrLf & _" & vbCrLf & _
    "                          ""Please change to another name..."", _" & vbCrLf & _
    "                          vbExclamation, ""Bookmark Name""" & vbCrLf & _
    "                   txtBookmark.SetFocus" & vbCrLf & _
    "                   SendKeys ""{Home}+{End}""" & vbCrLf & _
    "                   Exit Sub" & vbCrLf & _
    "                 End If" & vbCrLf & _
    "              Next i" & vbCrLf & _
    "              lstBookmark.AddItem txtBookmark.Text" & vbCrLf & _
    "              GiveMark" & vbCrLf & _
    "              txtBookmark.Text = """"" & vbCrLf & _
    "              cmdButton(0).Enabled = False" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "         Case 1 'Delete button clicked" & vbCrLf & _
    "              If lstBookmark.ListCount > 0 Then" & vbCrLf & _
    "                 lstBookmark.RemoveItem lstBookmark.ListIndex" & vbCrLf & _
    "                 If lstBookmark.ListCount > 0 Then" & vbCrLf & _
    "                    lstBookmark.Selected(0) = True" & vbCrLf & _
    "                    cmdButton(1).Enabled = True" & vbCrLf & _
    "                 Else" & vbCrLf & _
    "                    cmdButton(1).Enabled = False" & vbCrLf & _
    "                    cmdButton(2).Enabled = False" & vbCrLf & _
    "                 End If" & vbCrLf & _
    "              Else" & vbCrLf & _
    "                 cmdButton(1).Enabled = False" & vbCrLf & _
    "                 cmdButton(2).Enabled = False" & vbCrLf & _
    "              End If" & vbCrLf & _
    "         Case 2 'Jump button clicked" & vbCrLf & _
    "              Dim strTemp As String" & vbCrLf & _
    "              Dim Location As Double" & vbCrLf & _
    "              strTemp = Trim(lstBookmark.List(lstBookmark.ListIndex))" & vbCrLf & _
    "              Location = CekintPosition(strTemp)" & vbCrLf & _
    "              'Here is the essential of bookmark we added," & vbCrLf & _
    "              'we can jump direct to position we bookmark" & vbCrLf & _
    "              'before..." & vbCrLf & _
    "              adoBookMark.MoveFirst" & vbCrLf & _
    "              adoBookMark.Move Location - 1" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "         Case 3 'Cancel button clicked" & vbCrLf & _
    "              Me.Hide 'Just hide this form in order that" & vbCrLf & _
    "                      'we can use it later as long as program" & vbCrLf & _
    "                      'stay in memory" & vbCrLf & _
    "         Case 4 'Help button clicked, display how to use" & vbCrLf & _
    "                'bookmark..." & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "              MsgBox ""1. Bookmark is a way to mark a record in a recordset "" & vbCrLf & _" & vbCrLf & _
    "                     ""   so you can go back to the record later quickly"" & vbCrLf & _" & vbCrLf & _
    "                     ""   without remember the position of that record."" & vbCrLf & _" & vbCrLf & _
    "                     ""   Let program keep the position of record."" & vbCrLf & _" & vbCrLf & _
    "                     """" & vbCrLf & _" & vbCrLf & _
    "                     ""2. Select the record you want to bookmark by"" & vbCrLf & _" & vbCrLf & _
    "                     ""   clicking it in DataGrid or through Navigation "" & vbCrLf & _" & vbCrLf & _
    "                     ""   button on your form, then enter the bookmark "" & vbCrLf & _" & vbCrLf & _
    "                     ""   name in the textbox above, and press Enter or "" & vbCrLf & _" & vbCrLf & _
    "                     ""   click 'Add' button to add this name to the listbox "" & vbCrLf & _" & vbCrLf & _
    "                     ""   below. This will keep/save your bookmark."" & vbCrLf & _" & vbCrLf & _
    "                     """" & vbCrLf & _" & vbCrLf & _
    "                     ""3. If you want to go back to record that you have"" & vbCrLf & _" & vbCrLf & _
    "                     ""   marked, click bookmark name in the listbox"" & vbCrLf & _" & vbCrLf & _
    "                     ""   then click 'Jump' button, or you can double-click"" & vbCrLf & _" & vbCrLf & _
    "                     ""   the bookmark name in the listbox. "" & vbCrLf & _" & vbCrLf & _
    "                     """" & vbCrLf & _" & vbCrLf & _
    "                     ""4. If you want to delete the bookmark name, click"" & vbCrLf & _" & vbCrLf & _
    "                     ""   bookmark name in the listbox, then click"" & vbCrLf & _" & vbCrLf & _
    "                     ""   'Delete' button."" & vbCrLf & _" & vbCrLf & _
    "                     """", vbInformation, ""About Bookmark and How To Use It""" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "  End Select" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "" & vbCrLf & _
    "'This will check and take the position of bookmark" & vbCrLf & _
    "Function CekintPosition(Name As String) As Double" & vbCrLf & _
    "Dim i As Integer" & vbCrLf & _
    "  For i = 0 To UBound(tabMark)" & vbCrLf & _
    "    If Name = Trim(tabMark(i).BookmarkName) Then" & vbCrLf & _
    "       CekintPosition = tabMark(i).AbsolutePosition" & vbCrLf & _
    "       Exit For" & vbCrLf & _
    "    End If" & vbCrLf & _
    "  Next i" & vbCrLf & _
    "End Function" & vbCrLf & _
    "'In order that there is no double bookmark name saved in the listbox" & vbCrLf & _
    "Private Sub CheckDouble()" & vbCrLf & _
    "Dim i As Integer" & vbCrLf & _
    "  For i = 0 To lstBookmark.ListCount - 1" & vbCrLf & _
    "    If lstBookmark.List(i) = txtBookmark.Text Then" & vbCrLf & _
    "       MsgBox ""This bookmark name already exist in the list!"" & vbCrLf & _" & vbCrLf & _
    "              ""You can not save the same bookmark name."" & vbCrLf & _" & vbCrLf & _
    "              ""Please change to another name..."", _" & vbCrLf & _
    "              vbExclamation, ""Bookmark Name""" & vbCrLf & _
    "              txtBookmark.SetFocus" & vbCrLf & _
    "       SendKeys ""{Home}+{End}""" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "       Exit Sub" & vbCrLf & _
    "    End If" & vbCrLf & _
    "  Next i" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "Private Sub Form_Load()" & vbCrLf & _
    "  LockTheFormButton" & vbCrLf & _
    "  ReDim tabMark(lstBookmark.ListCount)" & vbCrLf & _
    "  'Get setting for this form from INI File" & vbCrLf & _
    "  Call ReadFromINIToControls(frmBookmark, ""Bookmark"")" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)" & vbCrLf & _
    "  Cancel = -1  'We don't unload this form" & vbCrLf & _
    "  Me.Hide      'Just hide them in order that we can" & vbCrLf & _
    "               'use this form later as long as program" & vbCrLf & _
    "               'stays in memory" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "'If a bookmark name was selected or clicked..." & vbCrLf & _
    "Private Sub lstBookmark_Click()" & vbCrLf & _
    "  If lstBookmark.ListCount > 0 Then" & vbCrLf & _
    "     'Unlock the button we will use" & vbCrLf & _
    "     UnlockTheFormButton" & vbCrLf & _
    "  End If" & vbCrLf & _
    "  'If textbox is empty, Add button is not active" & vbCrLf & _
    "  'Add button is not active in order that there is" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    frmProcess.rtfLap1.SelStart = Len(frmProcess.rtfLap1.Text)
    frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
    "  'no bookmark name contains a empty string" & vbCrLf & _
    "  If Len(Trim(txtBookmark.Text)) = 0 Then _" & vbCrLf & _
    "     cmdButton(0).Enabled = False" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "'Alternative way to go to the record we mark," & vbCrLf & _
    "'by double click the bookmark name in listbox" & vbCrLf & _
    "Private Sub lstBookmark_DblClick()" & vbCrLf & _
    "  cmdButton(2).Enabled = True" & vbCrLf & _
    "  cmdButton_Click (2)" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "'If there is a change in textbox" & vbCrLf & _
    "Private Sub txtBookmark_Change()" & vbCrLf & _
    "  'If textbox is not empty" & vbCrLf & _
    "  If Len(Trim(txtBookmark.Text)) > 0 Then" & vbCrLf & _
    "     'Add button is active and ready now" & vbCrLf & _
    "     cmdButton(0).Enabled = True" & vbCrLf & _
    "     cmdButton(0).Default = True" & vbCrLf & _
    "  Else 'If textbox is empty" & vbCrLf & _
    "     cmdButton(0).Enabled = False" & vbCrLf & _
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
    "'Lock the button that we don't use it" & vbCrLf & _
    "Private Sub LockTheFormButton()" & vbCrLf & _
    "  For i = 0 To 2" & vbCrLf & _
    "    cmdButton(i).Enabled = False" & vbCrLf & _
    "  Next i" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "'Unlock the button that we need" & vbCrLf & _
    "Private Sub UnlockTheFormButton()" & vbCrLf & _
    "  For i = 0 To 2" & vbCrLf & _
    "    cmdButton(i).Enabled = True" & vbCrLf & _
    "  Next i" & vbCrLf & _
    "End Sub" & vbCrLf
    Print #1, frmProcess.rtfLap1.Text
  Close #1

End Sub
