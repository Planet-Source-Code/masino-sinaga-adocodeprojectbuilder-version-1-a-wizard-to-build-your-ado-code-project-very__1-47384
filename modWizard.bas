Attribute VB_Name = "modWizard"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpszOp As _
    String, ByVal lpszFile As String, ByVal lpszParams As String, _
    ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long

Public Const SW_SHOWNORMAL = 1

'WinHelp Commands
Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Public Const HELP_QUIT = &H2              '  Terminate help
Public Const HELP_CONTENTS = &H3&         '  Display index/contents
Public Const HELP_CONTEXT = &H1           '  Display topic in ulTopic
Public Const HELP_INDEX = &H3             '  Display index

Public Const WIZARD_NAME = "WizardTemplate"

Public strOrdTB As String, strOrdDG As String
Public strFieldChild As String, strFieldParent As String

Public Const APP_CATEGORY = "Wizards"
Public Const CONFIRM_KEY = "ConfirmScreen"
Public Const DONTSHOW_CONFIRM = "DontShow"

'BASE VALUE FOR HELP FILE FOR THIS WIZARD:
Public Const HELP_BASE = 1000
Public Const HELP_FILE = "ADOCODEPROJECTBUILDER1.HLP"

Public Type tInfo
  strDBFileName As String
  strDBPassword As String
  strDSNName As String
  strDSNPassword As String
  strDSNUserID As String
  strDSNDatabase As String
  strDSNDriver As String
  strDSNServer As String
  strFormLayout As String
  strRSTextBox As String
  strRSDataGrid As String
  strOrderTextBox As String
  strOrderDataGrid As String
  strRelationTextBox As String
  strRelationDataGrid As String
  strFormName As String
  intNumOfFields As Byte
End Type
Public glo As tInfo

Public cnn As ADODB.Connection
Public rs As ADODB.Recordset
Public fldr As String
Public strFileName As String
Public strSQLOpen As String
Public mbHelpStarted As Boolean

Function GetResString(nRes As Integer) As String
    Dim sTmp As String
    Dim sRetStr As String
  
    Do
        sTmp = LoadResString(nRes)
        If Right(sTmp, 1) = "_" Then
            sRetStr = sRetStr + VBA.Left(sTmp, Len(sTmp) - 1)
        Else
            sRetStr = sRetStr + sTmp
        End If
        nRes = nRes + 1
    Loop Until Right(sTmp, 1) <> "_"
    GetResString = sRetStr
  
End Function

Sub LoadResStrings(frm As Form)
    On Error Resume Next
    
    Dim ctl As Control
    Dim obj As Object
    
    'set the form's caption
    If IsNumeric(frm.Tag) Then
       frm.Caption = LoadResString(CInt(frm.Tag))
    End If
    
    'set the controls' captions using the caption
    'property for menu items and the Tag property
    'for all other controls
    For Each ctl In frm.Controls
        If TypeName(ctl) = "Menu" Then
            If IsNumeric(ctl.Caption) Then
                If Err = 0 Then
                    ctl.Caption = LoadResString(CInt(ctl.Caption))
                Else
                    Err = 0
                End If
            End If
        ElseIf TypeName(ctl) = "TabStrip" Then
            For Each obj In ctl.Tabs
                If IsNumeric(obj.Tag) Then
                    obj.Caption = LoadResString(CInt(obj.Tag))
                End If
                'check for a tooltip
                If IsNumeric(obj.ToolTipText) Then
                    If Err = 0 Then
                        obj.ToolTipText = LoadResString(CInt(obj.ToolTipText))
                    Else
                        Err = 0
                    End If
                End If
            Next
        ElseIf TypeName(ctl) = "Toolbar" Then
            For Each obj In ctl.Buttons
                If IsNumeric(obj.Tag) Then
                    obj.ToolTipText = LoadResString(CInt(obj.Tag))
                End If
            Next
        ElseIf TypeName(ctl) = "ListView" Then
            For Each obj In ctl.ColumnHeaders
                If IsNumeric(obj.Tag) Then
                    obj.Text = LoadResString(CInt(obj.Tag))
                End If
            Next
        Else
            If IsNumeric(ctl.Tag) Then
                If Err = 0 Then
                    ctl.Caption = GetResString(CInt(ctl.Tag))
                Else
                    Err = 0
                End If
            End If
            'check for a tooltip
            If IsNumeric(ctl.ToolTipText) Then
                If Err = 0 Then
                    ctl.ToolTipText = LoadResString(CInt(ctl.ToolTipText))
                Else
                    Err = 0
                End If
            End If
        End If
    Next

End Sub

'Close all forms in this project
Public Sub CloseAllForms()
Dim Form As Form
   For Each Form In Forms
       Unload Form
       Set Form = Nothing
   Next Form
   End
End Sub


Public Function OpenDocument(ByVal DocName As String) As Long
   OpenDocument = ShellExecute(frmProcess.hwnd, "Open", DocName, _
          "", "C:\", SW_SHOWNORMAL)
End Function

'This will create sub directory in the app's location
'So, in this sub directory I put all files (.frm, .bas,
'and .vbp). This will make you easy to manage your project
'that has been generated by this program.
Public Sub CreateProjectDirectory()
    Dim fso As FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    fldr = App.Path & "\" & glo.strFormName
    If Not (fso.FolderExists(fldr)) Then
        fso.CreateFolder (fldr)
    End If
End Sub


