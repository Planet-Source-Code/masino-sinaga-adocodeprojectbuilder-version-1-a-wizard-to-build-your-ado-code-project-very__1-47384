VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmProcess 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generating..."
   ClientHeight    =   645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   645
   ScaleWidth      =   1800
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox rtfLap1 
      Height          =   255
      Left            =   3600
      TabIndex        =   0
      Top             =   0
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   450
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmProcess.frx":0000
   End
End
Attribute VB_Name = "frmProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
     
Dim tabFieldParent() As String
Dim tabFieldChild() As String
      

Public Sub GenerateADOCodeToModule()
  DoEvents
  frmWizard.lblProcess.Caption = "Generating code to module..."
  DoEvents
  'This will create a module file with extension .bas
  strFileName = fldr & "\modGeneral.bas"
  DoEvents
  'For your information, I open and close the file
  'in order that to prevent an error from IDE VB:
  '"Too many line continuations" if we assign the code
  'all together to richtextbox control...
  'So, I make them separate in several "open and save"!
  'I don't use additional file to put this template of
  'code because the code contains some variable that
  'we get from the wizard.
  
  'You'll find this method below...
  DoEvents
  frmWizard.lblProcess.Caption = "Generating header in module..."
  DoEvents

  'Start with empty richtextbox
  rtfLap1.Text = ""
  Open strFileName For Output As #1
    rtfLap1.Text = _
    "Attribute VB_Name = ""modGeneral""" & vbCrLf & _
    "'File Name  : modGeneral.bas" & vbCrLf & _
    "'Description: - Global variable declaration" & vbCrLf & _
    "'             - Global procedure/function" & vbCrLf & _
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
    Print #1, rtfLap1.Text
  Close #1
  
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "" & vbCrLf & _
    "Option Explicit" & vbCrLf & _
    "" & vbCrLf & _
    "Public cnn As ADODB.Connection" & vbCrLf
    If frmWizard.chkButton(6).Value = 1 Then _
       rtfLap1.Text = rtfLap1.Text & _
       "Public adoFind As ADODB.Recordset" & vbCrLf
    If frmWizard.chkButton(7).Value = 1 Then _
       rtfLap1.Text = rtfLap1.Text & _
       "Public adoFilter As ADODB.Recordset" & vbCrLf
    If frmWizard.chkButton(8).Value = 1 Then _
       rtfLap1.Text = rtfLap1.Text & _
    "Public adoSort As ADODB.Recordset" & vbCrLf
    If frmWizard.chkButton(9).Value = 1 Then _
       rtfLap1.Text = rtfLap1.Text & _
    "Public adoBookmark As ADODB.Recordset" & vbCrLf
    
    If frmWizard.lstFormLayout.Text = frmWizard.lstFormLayout.List(0) Then
      rtfLap1.Text = rtfLap1.Text & _
      "Public m_ConnectionString As String" & vbCrLf & _
      "Public m_RecordSource1 As String" & vbCrLf & _
      "Public m_SQLRS1 As String" & vbCrLf & _
      "Public m_FieldKey1 As String" & vbCrLf & _
      "Public strSQL As String" & vbCrLf & _
      "Public intMax As Integer" & vbCrLf
    Else 'Master/Detail
      rtfLap1.Text = rtfLap1.Text & _
      "Public m_ConnectionString As String" & vbCrLf & _
      "Public m_SQLRS1 As String" & vbCrLf & _
      "Public m_SQLRS2 As String" & vbCrLf & _
      "Public m_RecordSource1 As String" & vbCrLf & _
      "Public m_RecordSource2 As String" & vbCrLf & _
      "Public m_FieldKey1 As String" & vbCrLf & _
      "Public m_FieldKey2 As String" & vbCrLf & _
      "Public strSQL As String" & vbCrLf & _
      "Public intMax As Integer" & vbCrLf
    End If
    
    If frmWizard.lstFormLayout.Text = frmWizard.lstFormLayout.List(0) Then
     frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
     "Public rs1 As ADODB.Recordset" & vbCrLf & _
     "Public adoField1 As ADODB.Field" & vbCrLf
    Else 'MasterDetail
     frmProcess.rtfLap1.Text = frmProcess.rtfLap1.Text & _
     "Public rs1 As ADODB.Recordset" & vbCrLf & _
     "Public adoField1 As ADODB.Field" & vbCrLf & _
     "Public rs2 As ADODB.Recordset" & vbCrLf & _
     "Public adoField2 As ADODB.Field" & vbCrLf
    End If
    
    Print #1, rtfLap1.Text
  Close #1
  
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox

  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Public Declare Function GetPrivateProfileString Lib ""kernel32"" Alias _" & vbCrLf & _
    """GetPrivateProfileStringA"" (ByVal lpApplicationName As String, ByVal lpKeyName As _" & vbCrLf & _
    "Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, _" & vbCrLf & _
    "ByVal lpFileName As String) As Long" & vbCrLf & _
    "" & vbCrLf & _
    "Public INIFileName As String" & vbCrLf & _
    "" & vbCrLf & _
    "Public Declare Function WritePrivateProfileString Lib ""kernel32"" Alias _" & vbCrLf & _
    """WritePrivateProfileStringA"" (ByVal lpApplicationName As String, ByVal lpKeyName _" & vbCrLf & _
    "As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
    
  'Copy the Access database file to project directory,
  'so you don't need to remember where did you get
  'this file if you want to copy this project to
  'another machine or another location, because
  'the application will access database file by
  'using App.Path &"\" FileName
  Dim OldLocation As String
  If frmWizard.chkCopyDBFile.Value = 1 Then
     cnn.Close
     Set cnn = Nothing  'clear memory of database
     'so we can copy the database later
     OldLocation = glo.strDBFileName
     glo.strDBFileName = """ & App.Path & ""\" & StripPath(glo.strDBFileName) & ""
  End If
       
  If frmWizard.lstDBFormat.Text = frmWizard.lstDBFormat.List(0) Then 'Using Access database
     Dim strSQLCon As String
     strSQLCon = _
     "PROVIDER=MSDataShape;Data PROVIDER="" & _" & vbCrLf & _
     "           ""Microsoft.Jet.OLEDB.4.0;Data Source="" & _" & vbCrLf & _
     "           """ & glo.strDBFileName & ";Jet OLEDB:"" & _ " & vbCrLf & _
     "           ""Database Password=" & glo.strDBPassword & ";"
    'Save again to temporary file
    Open strFileName For Output As #1
    
     'Get SQL statement from master/textbox recordsource
     Dim intField1 As Integer, i As Integer
     Dim tabField1() As String, strField1 As String
     Dim strSQLMaster As String
     ReDim tabField1(frmWizard.lstSelectedFields(0).ListCount - 1)
     intField1 = frmWizard.lstSelectedFields(0).ListCount
     For i = 0 To frmWizard.lstSelectedFields(0).ListCount - 1
        tabField1(i) = frmWizard.lstSelectedFields(0).List(i)
        If i <> intField1 - 1 Then
           strField1 = strField1 & tabField1(i) & ","
        Else
           strField1 = strField1 & tabField1(i) & ""
        End If
     Next i
     strSQLMaster = "SELECT " & strField1 & " FROM " & frmWizard.cboRS(0).Text
     
     'Get SQL statement from detail/datagrid recordsource
     Dim intField2 As Integer, strSQLDetail As String
     Dim tabField2() As String, strField2 As String
     ReDim tabField2(frmWizard.lstSelectedFields(1).ListCount - 1)
     intField2 = frmWizard.lstSelectedFields(1).ListCount
     For i = 0 To frmWizard.lstSelectedFields(1).ListCount - 1
        tabField2(i) = frmWizard.lstSelectedFields(1).List(i)
        If i <> intField2 - 1 Then
           strField2 = strField2 & tabField2(i) & ","
        Else
           strField2 = strField2 & tabField2(i) & ""
        End If
     Next i
     strSQLDetail = "SELECT " & strField2 & " FROM " & frmWizard.cboRS(1).Text

      'Get the new start of next record in richtexbox control
      rtfLap1.SelStart = Len(rtfLap1.Text)
      If frmWizard.lstFormLayout.Text = frmWizard.lstFormLayout.List(0) Then
       rtfLap1.Text = rtfLap1.Text & _
       "Public Sub OpenConnection()" & vbCrLf & _
       "  Set cnn = New ADODB.Connection" & vbCrLf & _
       "  cnn.CursorLocation = adUseClient" & vbCrLf & _
       "  m_ConnectionString = """ & strSQLCon & """" & vbCrLf & _
       "  cnn.Open m_ConnectionString " & vbCrLf & _
       "  m_SQLRS1 = """ & strSQLMaster & """" & vbCrLf & _
       "  m_RecordSource1 = """ & frmWizard.cboRS(0).Text & """" & vbCrLf & _
       "  m_FieldKey1 = """ & frmWizard.lstTextBox.Text & """" & vbCrLf & _
       "" & vbCrLf & _
       "End Sub" & vbCrLf
      Else 'Master/Detail
       rtfLap1.Text = rtfLap1.Text & _
       "Public Sub OpenConnection()" & vbCrLf & _
       "  Set cnn = New ADODB.Connection" & vbCrLf & _
       "  cnn.CursorLocation = adUseClient" & vbCrLf & _
       "  m_ConnectionString = """ & strSQLCon & """" & vbCrLf & _
       "  cnn.Open m_ConnectionString " & vbCrLf & _
       "  m_SQLRS1 = """ & strSQLMaster & """" & vbCrLf & _
       "  m_RecordSource1 = """ & frmWizard.cboRS(0).Text & """" & vbCrLf & _
       "  m_SQLRS2 = """ & strSQLDetail & """" & vbCrLf & _
       "  m_RecordSource2 = """ & frmWizard.cboRS(1).Text & """" & vbCrLf & _
       "  m_FieldKey1 = """ & frmWizard.lstTextBox.Text & """" & vbCrLf & _
       "  m_FieldKey2 = """ & frmWizard.lstDataGrid.Text & """" & vbCrLf & _
       "" & vbCrLf & _
       "End Sub" & vbCrLf
      End If
      Print #1, rtfLap1.Text
    Close #1
  Else  'using DSN (ODBC)
    'Check again, whether using DSN or using parameter
    'to connect to the server (e.g. MySQL, SQL Server, etc)
    If frmWizard.cboDSNList.Text <> _
       frmWizard.cboDSNList.List(0) Then  'DSN was choosen
      Open strFileName For Output As #1
        'Get the new start of next record in richtexbox control
        rtfLap1.SelStart = Len(rtfLap1.Text)
        If frmWizard.lstFormLayout.Text = frmWizard.lstFormLayout.List(0) Then
         rtfLap1.Text = rtfLap1.Text & _
         "Public Sub OpenConnection()" & vbCrLf & _
         "  Set cnn = New ADODB.Connection" & vbCrLf & _
         "  cnn.CursorLocation = adUseClient" & vbCrLf & _
         "  cnn.Open ""PROVIDER=MSDataShape;"" & _ " & vbCrLf & _
         "  ""Data PROVIDER=MSDASQL;"" & _ " & vbCrLf & _
         "  ""DSN=" & glo.strDSNName & ";"" & _ " & vbCrLf & _
         "  ""UID=" & glo.strDSNUserID & ";"" & _ " & vbCrLf & _
         "  ""PWD=" & glo.strDSNPassword & ";""" & vbCrLf & _
         "  m_SQLRS1 = """ & strSQLMaster & """" & vbCrLf & _
         "  m_RecordSource1 = """ & frmWizard.cboRS(0).Text & """" & vbCrLf & _
         "End Sub" & vbCrLf
        Else 'Master/Detail
         rtfLap1.Text = rtfLap1.Text & _
         "Public Sub OpenConnection()" & vbCrLf & _
         "  Set cnn = New ADODB.Connection" & vbCrLf & _
         "  cnn.CursorLocation = adUseClient" & vbCrLf & _
         "  cnn.Open ""PROVIDER=MSDataShape;"" & _ " & vbCrLf & _
         "  ""Data PROVIDER=MSDASQL;"" & _ " & vbCrLf & _
         "  ""DSN=" & glo.strDSNName & ";"" & _ " & vbCrLf & _
         "  ""UID=" & glo.strDSNUserID & ";"" & _ " & vbCrLf & _
         "  ""PWD=" & glo.strDSNPassword & ";""" & vbCrLf & _
         "  m_SQLRS1 = """ & strSQLMaster & """" & vbCrLf & _
         "  m_RecordSource1 = """ & frmWizard.cboRS(0).Text & """" & vbCrLf & _
         "  m_SQLRS2 = """ & strSQLDetail & """" & vbCrLf & _
         "  m_RecordSource2 = """ & frmWizard.cboRS(1).Text & """" & vbCrLf & _
         "  m_FieldKey1 = """ & frmWizard.lstTextBox.Text & """" & vbCrLf & _
         "  m_FieldKey2 = """ & frmWizard.lstDataGrid.Text & """" & vbCrLf & _
         "End Sub" & vbCrLf
        End If
        Print #1, rtfLap1.Text
      Close #1
    Else  'DSN was not choosen
      'Check again, whether MySQL or SQL Server ?
      'This is for MySQL:
      If InStr(1, UCase(frmWizard.cboDrivers.Text), "MYSQL") > 0 Then
        Open strFileName For Output As #1
          'Get the new start of next record in richtexbox control
          rtfLap1.SelStart = Len(rtfLap1.Text)
          If frmWizard.lstFormLayout.Text = frmWizard.lstFormLayout.List(0) Then
           rtfLap1.Text = rtfLap1.Text & _
           "Public Sub OpenConnection()" & vbCrLf & _
           "  Set cnn = New ADODB.Connection" & vbCrLf & _
           "  cnn.CursorLocation = adUseClient" & vbCrLf & _
           "  cnn.Open ""PROVIDER=MSDataShape;"" & _ " & vbCrLf & _
           "  ""Driver={" & glo.strDSNDriver & "};"" & _ " & vbCrLf & _
           "  ""Server=" & glo.strDSNServer & ";"" & _ " & vbCrLf & _
           "  ""Port=3306;"" & _ " & vbCrLf & _
           "  ""Option=147458;"" & _ " & vbCrLf & _
           "  ""Stmt=;"" & _ " & vbCrLf & _
           "  ""Database=" & glo.strDSNDatabase & ";"" & _ " & vbCrLf & _
           "  ""User=" & glo.strDSNUserID & ";"" & _ " & vbCrLf & _
           "  ""Password=" & glo.strDSNPassword & """" & vbCrLf & _
           "  m_SQLRS1 = """ & strSQLMaster & """" & vbCrLf & _
           "  m_RecordSource1 = """ & frmWizard.cboRS(0).Text & """" & vbCrLf & _
           "End Sub" & vbCrLf
          Else 'Master/Detail
           rtfLap1.Text = rtfLap1.Text & _
           "Public Sub OpenConnection()" & vbCrLf & _
           "  Set cnn = New ADODB.Connection" & vbCrLf & _
           "  cnn.CursorLocation = adUseClient" & vbCrLf & _
           "  cnn.Open ""PROVIDER=MSDataShape;"" & _ " & vbCrLf & _
           "  ""Driver={" & glo.strDSNDriver & "};"" & _ " & vbCrLf & _
           "  ""Server=" & glo.strDSNServer & ";"" & _ " & vbCrLf & _
           "  ""Port=3306;"" & _ " & vbCrLf & _
           "  ""Option=147458;"" & _ " & vbCrLf & _
           "  ""Stmt=;"" & _ " & vbCrLf & _
           "  ""Database=" & glo.strDSNDatabase & ";"" & _ " & vbCrLf & _
           "  ""User=" & glo.strDSNUserID & ";"" & _ " & vbCrLf & _
           "  ""Password=" & glo.strDSNPassword & """" & vbCrLf & _
           "  m_SQLRS1 = """ & strSQLMaster & """" & vbCrLf & _
           "  m_RecordSource1 = """ & frmWizard.cboRS(0).Text & """" & vbCrLf & _
           "  m_SQLRS2 = """ & strSQLDetail & """" & vbCrLf & _
           "  m_RecordSource2 = """ & frmWizard.cboRS(1).Text & """" & vbCrLf & _
           "  m_FieldKey1 = """ & frmWizard.lstTextBox.Text & """" & vbCrLf & _
           "  m_FieldKey2 = """ & frmWizard.lstDataGrid.Text & """" & vbCrLf & _
           "End Sub" & vbCrLf
          End If
          Print #1, rtfLap1.Text
        Close #1
      Else '(e.g. SQL Server; I don't try this,
           'just my general assumption... ;P )
        Open strFileName For Output As #1
          'Get the new start of next record in richtexbox control
          rtfLap1.SelStart = Len(rtfLap1.Text)
          If frmWizard.lstFormLayout.Text = frmWizard.lstFormLayout.List(0) Then
           rtfLap1.Text = rtfLap1.Text & _
           "Public Sub OpenConnection()" & vbCrLf & _
           "  Set cnn = New ADODB.Connection" & vbCrLf & _
           "  cnn.CursorLocation = adUseClient" & vbCrLf & _
           "  cnn.Open ""PROVIDER=MSDataShape;"" & _ " & vbCrLf & _
           "  ""Driver={" & glo.strDSNDriver & "};"" & _ " & vbCrLf & _
           "  ""Server=" & glo.strDSNServer & ";"" & _ " & vbCrLf & _
           "  ""Database=" & glo.strDSNDatabase & ";"" & _ " & vbCrLf & _
           "  ""User=" & glo.strDSNUserID & ";"" & _ " & vbCrLf & _
           "  ""Password=" & glo.strDSNPassword & """" & vbCrLf & _
           "  m_SQLRS1 = """ & strSQLMaster & """" & vbCrLf & _
           "  m_RecordSource1 = """ & frmWizard.cboRS(0).Text & """" & vbCrLf & _
           "End Sub" & vbCrLf
          Else 'Master/Detail
           rtfLap1.Text = rtfLap1.Text & _
           "Public Sub OpenConnection()" & vbCrLf & _
           "  Set cnn = New ADODB.Connection" & vbCrLf & _
           "  cnn.CursorLocation = adUseClient" & vbCrLf & _
           "  cnn.Open ""PROVIDER=MSDataShape;"" & _ " & vbCrLf & _
           "  ""Driver={" & glo.strDSNDriver & "};"" & _ " & vbCrLf & _
           "  ""Server=" & glo.strDSNServer & ";"" & _ " & vbCrLf & _
           "  ""Database=" & glo.strDSNDatabase & ";"" & _ " & vbCrLf & _
           "  ""User=" & glo.strDSNUserID & ";"" & _ " & vbCrLf & _
           "  ""Password=" & glo.strDSNPassword & """" & vbCrLf & _
           "  m_SQLRS1 = """ & strSQLMaster & """" & vbCrLf & _
           "  m_RecordSource1 = """ & frmWizard.cboRS(0).Text & """" & vbCrLf & _
           "  m_SQLRS2 = """ & strSQLDetail & """" & vbCrLf & _
           "  m_RecordSource2 = """ & frmWizard.cboRS(1).Text & """" & vbCrLf & _
           "  m_FieldKey1 = """ & frmWizard.lstTextBox.Text & """" & vbCrLf & _
           "  m_FieldKey2 = """ & frmWizard.lstDataGrid.Text & """" & vbCrLf & _
           "End Sub" & vbCrLf
          End If
          Print #1, rtfLap1.Text
        Close #1
      End If
    End If
  End If
      
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
            
  If frmWizard.chkCopyDBFile.Value = 1 Then
     Call FileCopy(OldLocation, fldr & "\" & StripPath(OldLocation))
  End If
  
 'This is for cmdDataGrid button
 If frmWizard.chkButton(10).Value = 1 Then
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Public Sub AdjustDataGridColumnWidth _" & vbCrLf & _
    "           (DG As DataGrid, _" & vbCrLf & _
    "           adoData As ADODB.Recordset, _" & vbCrLf & _
    "           intRecord As Integer, _" & vbCrLf & _
    "           intField As Integer, _" & vbCrLf & _
    "           Optional AccForHeaders As Boolean)" & vbCrLf & _
    "" & vbCrLf & _
    "'This procedure will adjust DataGrids column width" & vbCrLf & _
    "'based on longest field in underlying source" & vbCrLf & _
    "" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
    
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "    Dim row As Long, col As Long" & vbCrLf & _
    "    Dim width As Single, maxWidth As Single" & vbCrLf & _
    "    Dim saveFont As StdFont, saveScaleMode As Integer" & vbCrLf & _
    "    Dim cellText As String" & vbCrLf & _
    "    Dim i As Integer" & vbCrLf & _
    "    'If number of records = 0 then exit from the sub" & vbCrLf & _
    "    If intRecord = 0 Then Exit Sub" & vbCrLf & _
    "    'Save the form's font for DataGrid's font" & vbCrLf & _
    "    'We need this for form's TextWidth method" & vbCrLf & _
    "    Set saveFont = DG.Parent.Font" & vbCrLf & _
    "    Set DG.Parent.Font = DG.Font" & vbCrLf & _
    "    'Adjust ScaleMode to vbTwips for the form (parent)." & vbCrLf & _
    "    saveScaleMode = DG.Parent.ScaleMode" & vbCrLf & _
    "    DG.Parent.ScaleMode = vbTwips" & vbCrLf & _
    "    'Always from first record..." & vbCrLf & _
    "    adoData.MoveFirst" & vbCrLf & _
    "    maxWidth = 0" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "    'Get maximal value for progressbar control" & vbCrLf & _
    "    intMax = intField * intRecord" & vbCrLf & _
    "    " & glo.strFormName & ".prgBar1.Visible = True" & vbCrLf & _
    "    " & glo.strFormName & ".prgBar1.Max = intMax" & vbCrLf & _
    "        " & vbCrLf & _
    "    'We begin from the first column until the last column" & vbCrLf & _
    "    For col = 0 To intField - 1" & vbCrLf & _
    "        'Tampilkan nama field/kolom yg sedang diproses" & vbCrLf & _
    "        " & glo.strFormName & ".lblField.Caption = _" & vbCrLf & _
    "           ""Column: "" & DG.Columns(col).DataField & """ & vbCrLf & _
    "        adoData.MoveFirst" & vbCrLf & _
    "        'Optional param, if true, set maxWidth to" & vbCrLf & _
    "        'width of DG.Parent" & vbCrLf & _
    "        If AccForHeaders Then" & vbCrLf & _
    "            maxWidth = DG.Parent.TextWidth(DG.Columns(col).Text) + 200" & vbCrLf & _
    "        End If" & vbCrLf & _
    "        'Repeat from first record again after we have" & vbCrLf & _
    "        'finished process the last record in" & vbCrLf & _
    "        'former column..." & vbCrLf & _
    "        adoData.MoveFirst" & vbCrLf & _
    "        For row = 0 To intRecord - 1" & vbCrLf & _
    "            'Get the text from the DataGrid's cell" & vbCrLf & _
    "            If intField = 1 Then" & vbCrLf & _
    "            Else  'If number of field more than one" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "               cellText = DG.Columns(col).Text" & vbCrLf & _
    "            End If" & vbCrLf & _
    "            width = DG.Parent.TextWidth(cellText) + 200" & vbCrLf & _
    "            If width > maxWidth Then" & vbCrLf & _
    "               maxWidth = width" & vbCrLf & _
    "               DG.Columns(col).width = maxWidth" & vbCrLf & _
    "            End If" & vbCrLf & _
    "            adoData.MoveNext" & vbCrLf & _
    "            DoEvents" & vbCrLf & _
    "            i = i + 1" & vbCrLf & _
    "            " & glo.strFormName & ".lblAngka.Caption = _" & vbCrLf & _
    "              ""Finished "" & Format((i / intMax) * 100, ""0"") & ""%""" & vbCrLf & _
    "             DoEvents" & vbCrLf & _
    "            " & glo.strFormName & ".prgBar1.Value = i" & vbCrLf & _
    "            DoEvents" & vbCrLf & _
    "        Next row" & vbCrLf & _
    "        DG.Columns(col).width = maxWidth " & vbCrLf & _
    "    Next col" & vbCrLf & _
    "    'Change the DataGrid's parent property" & vbCrLf & _
    "    Set DG.Parent.Font = saveFont" & vbCrLf & _
    "    DG.Parent.ScaleMode = saveScaleMode" & vbCrLf & _
    "    adoData.MoveFirst" & vbCrLf & _
    "    ResetProgressBar" & vbCrLf & _
    "End Sub  'End of AdjustDataGridColumnWidth" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  
  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Public Sub ResetProgressBar()" & vbCrLf & _
    "  With " & glo.strFormName & "" & vbCrLf & _
    "    .prgBar1.Value = 0" & vbCrLf & _
    "    .lblAngka.Caption = """ & vbCrLf & _
    "    .lblField.Caption = """ & vbCrLf & _
    "  End With" & vbCrLf & _
    "End Sub" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1

  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
 End If  'End of cmdDataGrid button
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Public Function SaveFromControlsToINI(Objek, MyAppName As String)" & vbCrLf & _
    "Dim Contrl As Control, Result As Long" & vbCrLf & _
    "Dim TempControlName As String, TempControlValue As String" & vbCrLf & _
    "On Error Resume Next" & vbCrLf & _
    "For Each Contrl In Objek" & vbCrLf & _
    "  If (TypeOf Contrl Is CheckBox) Or (TypeOf Contrl Is ComboBox) Then" & vbCrLf & _
    "    TempControlName = Contrl.Name" & vbCrLf & _
    "    TempControlValue = Contrl.Value" & vbCrLf & _
    "    If (TypeOf Contrl Is ComboBox) Then" & vbCrLf & _
    "      TempControlValue = Contrl.Text" & vbCrLf & _
    "      If TempControlValue = """" Then TempControlValue = 1" & vbCrLf & _
    "    End If" & vbCrLf & _
    "    Result = WritePrivateProfileString(MyAppName, TempControlName, _" & vbCrLf & _
    "    TempControlValue, INIFileName)" & vbCrLf & _
    "  End If" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1

  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "  If (TypeOf Contrl Is TextBox) Then" & vbCrLf & _
    "    TempControlName = Contrl.Name" & vbCrLf & _
    "    TempControlValue = Contrl.Text" & vbCrLf & _
    "    Result = WritePrivateProfileString(MyAppName, TempControlName, _" & vbCrLf & _
    "    TempControlValue, INIFileName)" & vbCrLf & _
    "  End If" & vbCrLf & _
    "  If (TypeOf Contrl Is OptionButton) Then" & vbCrLf & _
    "    TempControlValue = Contrl.Value" & vbCrLf & _
    "    If TempControlValue = True Then" & vbCrLf & _
    "      TempControlName = Contrl.Name" & vbCrLf & _
    "      TempControlValue = Contrl.Index" & vbCrLf & _
    "      Result = WritePrivateProfileString(MyAppName, TempControlName, _" & vbCrLf & _
    "      TempControlValue, INIFileName)" & vbCrLf & _
    "    End If" & vbCrLf & _
    "  End If" & vbCrLf & _
    "Next" & vbCrLf & _
    "End Function" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1

  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Public Function ReadFromINIToControls(Objek, MyAppName As String)" & vbCrLf & _
    "Dim Contrl As Control, Result As Long" & vbCrLf & _
    "Dim TempControlName As String * 101, TempControlValue As String * 101" & vbCrLf & _
    "On Error Resume Next" & vbCrLf & _
    "For Each Contrl In Objek" & vbCrLf & _
    "If (TypeOf Contrl Is CheckBox) Or (TypeOf Contrl Is ComboBox) Or (TypeOf _" & vbCrLf & _
    "Contrl Is OptionButton) Or (TypeOf Contrl Is TextBox) Or (TypeOf Contrl Is CheckBox) Then" & vbCrLf & _
    "TempControlName = Contrl.Name" & vbCrLf & _
    "If (TypeOf Contrl Is TextBox) Or (TypeOf Contrl Is ComboBox) Then 'Or _" & vbCrLf & _
    "   '(TypeOf Contrl Is MaskEdBox) Then" & vbCrLf & _
    "   Result = GetPrivateProfileString(MyAppName, TempControlName, """", _" & vbCrLf & _
    "   TempControlValue, Len(TempControlValue), INIFileName)" & vbCrLf & _
    "Else 'If (TypeOf Contrl Is CheckBox) Then" & vbCrLf & _
    "   Result = GetPrivateProfileString(MyAppName, TempControlName, ""0"", _" & vbCrLf & _
    "   TempControlValue, Len(TempControlValue), INIFileName)" & vbCrLf & _
    "End If" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1

  'Put the saved code to richtextbox control again
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "If (TypeOf Contrl Is OptionButton) Then" & vbCrLf & _
    "   If Contrl.Index = Val(TempControlValue) Then Contrl = True" & vbCrLf & _
    "Else" & vbCrLf & _
    "    Contrl = TempControlValue" & vbCrLf & _
    "   If (TypeOf Contrl Is ComboBox) Then" & vbCrLf & _
    "      If Len(Contrl.Text) = 0 Then Contrl.ListIndex = 0" & vbCrLf & _
    "      End If" & vbCrLf & _
    "   End If" & vbCrLf & _
    "End If" & vbCrLf & _
    "Next" & vbCrLf & _
    "End Function" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  
End Sub

Public Sub RunProject()
  Call OpenDocument(fldr & "\prj" & Trim(glo.strFormName) & ".vbp")
End Sub

'------ This will generate reference and form to project -----
Public Sub GenerateADOCodeToProject()
  DoEvents
  frmWizard.lblProcess.Caption = "Generating code for project..."
  DoEvents
  'This will create a project file with extension .vbp
  strFileName = fldr & "\prj" & Trim(glo.strFormName) & ".vbp"
  'Start with empty richtextbox
  rtfLap1.Text = ""
  Open strFileName For Output As #1
    rtfLap1.Text = _
    "Type=Exe" & vbCrLf & _
    "Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\WINDOWS\SYSTEM\stdole2.tlb#OLE Automation" & vbCrLf & _
    "Reference=*\G{00000200-0000-0010-8000-00AA006D2EA4}#2.0#0#C:\PROGRAM FILES\COMMON FILES\SYSTEM\ADO\msado20.tlb#Microsoft ActiveX Data Objects 2.0 Library" & vbCrLf & _
    "Object={CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0; MSDATGRD.OCX" & vbCrLf & _
    "Object={6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0; COMCTL32.OCX" & vbCrLf & _
    "Reference=*\G{56BF9020-7A2F-11D0-9482-00A0C91110ED}#1.0#0#C:\WINDOWS\SYSTEM\MSBIND.DLL#Microsoft Data Binding Collection VB 6.0 (SP4)" & vbCrLf & _
    "Form=" & glo.strFormName & ".frm" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1

  'After we save it to file, put it to richtexbox control
  PutCodeToRichTextBox
  
  'Check, which form will be included to your project
  Dim x As Byte
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    'Looping to include the valid form based on button you specified
    For x = 6 To 9  'start from index of 6 to 9
    If frmWizard.chkButton(x).Value = 1 Then
      rtfLap1.Text = rtfLap1.Text & _
      "Form=frm" & Trim(Right(frmWizard.chkButton(x).Tag, Len(frmWizard.chkButton(x).Tag) - 3)) & ".frm" & vbCrLf
    End If
    Next x
    Print #1, rtfLap1.Text
  Close #1
  
  'After we save it to file, put it to richtexbox control
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Module=modGeneral; modGeneral.bas" & vbCrLf & _
    "IconForm=""" & glo.strFormName & """" & vbCrLf & _
    "Startup=""" & glo.strFormName & """" & vbCrLf & _
    "HelpFile=""""" & vbCrLf & _
    "Title=""prj" & glo.strFormName & """" & vbCrLf & _
    "ExeName32=""" & glo.strFormName & ".exe""" & vbCrLf & _
    "Command32=""""" & vbCrLf & _
    "Name=""prj" & glo.strFormName & """" & vbCrLf & _
    "HelpContextID=""0""" & vbCrLf & _
    "CompatibleMode=""0""" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1

  'After we save it to file, put it to richtexbox control
  PutCodeToRichTextBox

  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "MajorVer=1" & vbCrLf & _
    "MinorVer=0" & vbCrLf & _
    "RevisionVer=0" & vbCrLf & _
    "AutoIncrementVer=0" & vbCrLf & _
    "ServerSupportFiles=0" & vbCrLf & _
    "VersionCompanyName=" & vbCrLf & _
    "CompilationType=0" & vbCrLf & _
    "OptimizationType=0" & vbCrLf & _
    "FavorPentiumPro(tm)=0" & vbCrLf & _
    "CodeViewDebugInfo=0" & vbCrLf & _
    "NoAliasing=0" & vbCrLf & _
    "BoundsCheck=0" & vbCrLf & _
    "OverflowCheck=0" & vbCrLf & _
    "FlPointCheck=0" & vbCrLf & _
    "FDIVCheck=0" & vbCrLf & _
    "UnroundedFP=0" & vbCrLf & _
    "StartMode=0" & vbCrLf & _
    "Unattended=0" & vbCrLf & _
    "Retained=0" & vbCrLf & _
    "ThreadPerObject=0" & vbCrLf & _
    "MaxNumberOfThreads=1" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
End Sub
'------ End of generate reference and form to project -----

'----- This will generate ADO Code to the form -----
Public Sub GenerateADOCodeToForm()
  DoEvents
  frmWizard.lblProcess.Caption = "Generating " & Trim(glo.strFormName) & ".frm form..."
  DoEvents
  Dim i As Integer
  Dim nTop As Long, nTabIndex As Byte
  Dim intSpace As Long
  intSpace = 300 * 1.04   'lblLabels.Height = 300
  nTabIndex = 0
  
  CreateProjectDirectory
  GenerateSQLOpenRecordset
   
  strFileName = fldr & "\" & Trim(glo.strFormName) & ".frm"
  Open strFileName For Output As #1
     
  'First of all, sum the height of form. This will
  'provide the place for all textbox controls, datagrid,
  'including another controls under datagrid control.
  Dim ct As Integer, FormHeight As Long
  Dim AllControlsHeight As Long, TextBoxesHeight As Long
    
    'Sum of all textboxes height
    For ct = 0 To frmWizard.lstSelectedFields(0).ListCount - 1
      TextBoxesHeight = TextBoxesHeight + intSpace
    Next ct
    
    'Sum of the height of form........
    FormHeight = TextBoxesHeight + 2000 + 255 + 180 + 300 + 600 + 600
    
    'Compare FormHeight to the height of picButtons where
    'the main buttons were placed.
    If FormHeight > 4665 Then   '4665 is height of picButtons (fix)
       FormHeight = FormHeight
    Else  'formheight is shorter than height of picButtons
       FormHeight = 5350    '<-- pembulatan dari 4665 + 650
    End If
            
    DoEvents
    frmWizard.lblProcess.Caption = "Generating controls to form..."
    DoEvents
            
    'GENERATING FORM START HERE...............
    rtfLap1.Text = _
    "VERSION 5.00" & vbCrLf & _
    "Object = ""{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0""; ""MSDATGRD.OCX""" & vbCrLf & _
    "Object = ""{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0""; ""COMCTL32.OCX""" & vbCrLf & _
    "Begin VB.Form " & glo.strFormName & "" & vbCrLf & _
    "  BorderStyle = 1        'Fixed Single" & vbCrLf & _
    "  Caption = ""This project was created by ADOCodeProjectBuilder ver. 1, (c) Masino Sinaga""" & vbCrLf & _
    "  ClientHeight = " & FormHeight & "" & vbCrLf & _
    "  ClientLeft = 1095" & vbCrLf & _
    "  ClientTop = 330" & vbCrLf & _
    "  ClientWidth = 7995" & vbCrLf & _
    "  KeyPreview = -1         'True" & vbCrLf & _
    "  LinkTopic = ""Form1""" & vbCrLf & _
    "  LockControls = -1       'True" & vbCrLf & _
    "  MaxButton = 0           'False" & vbCrLf & _
    "  ScaleHeight = 6510" & vbCrLf & _
    "  ScaleWidth = 7995" & vbCrLf & _
    "  StartUpPosition = 2    'CenterScreen" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
    
  PutCodeToRichTextBox
  
  DoEvents
  frmWizard.lblProcess.Caption = "Generating statusbar control..."
  DoEvents
  
  '-------- START GENERATE STATUSBAR -------------
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "   Begin ComctlLib.StatusBar StatusBar1" & vbCrLf & _
    "      Align = 2              'Align Bottom" & vbCrLf & _
    "      Height = 270" & vbCrLf & _
    "      Left = 0" & vbCrLf & _
    "      TabIndex = 30" & vbCrLf & _
    "      Top = 6240" & vbCrLf & _
    "      Width = 7995" & vbCrLf & _
    "      _ExtentX        =   14102" & vbCrLf & _
    "      _ExtentY        =   476" & vbCrLf & _
    "      SimpleText = """"" & vbCrLf & _
    "      _Version        =   327682" & vbCrLf & _
    "      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7}" & vbCrLf & _
    "         NumPanels = 4" & vbCrLf & _
    "         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7}" & vbCrLf & _
    "            AutoSize = 1" & vbCrLf & _
    "            Object.Width = 8625" & vbCrLf & _
    "            MinWidth = 7408" & vbCrLf & _
    "            Key = """ & vbCrLf & _
    "            Object.Tag = """"" & vbCrLf & _
    "            Object.ToolTipText = ""(C) Masino Sinaga (masino_sinaga@yahoo.com)""" & vbCrLf & _
    "         EndProperty" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7}" & vbCrLf & _
    "            Object.Width = 1764" & vbCrLf & _
    "            MinWidth = 1764" & vbCrLf & _
    "            Key = """"" & vbCrLf & _
    "            Object.Tag = """"" & vbCrLf & _
    "            Object.ToolTipText = ""It's up to you...""" & vbCrLf & _
    "         EndProperty" & vbCrLf & _
    "         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7}" & vbCrLf & _
    "            Style = 6" & vbCrLf & _
    "            Object.Width = 2117" & vbCrLf & _
    "            MinWidth = 2117" & vbCrLf & _
    "            TextSave = ""07/07/2003""" & vbCrLf & _
    "            Key = """"" & vbCrLf & _
    "            Object.Tag = """"" & vbCrLf & _
    "            Object.ToolTipText = ""Date today""" & vbCrLf & _
    "         EndProperty" & vbCrLf & _
    "         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7}" & vbCrLf & _
    "            Style = 5" & vbCrLf & _
    "            Object.Width = 1464" & vbCrLf & _
    "            MinWidth = 1464" & vbCrLf & _
    "            TextSave = ""06:05""" & vbCrLf & _
    "            Key = """"" & vbCrLf & _
    "            Object.Tag = """"" & vbCrLf & _
    "           Object.ToolTipText = ""Time right now""" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "         EndProperty" & vbCrLf & _
    "      EndProperty" & vbCrLf & _
    "   End" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
  '----------- END OF GENERATE STATUSBAR ------
  
  '----------- START GENERATE LABEL -----------
  'Save again to temporary file
  Open strFileName For Output As #1
    nTop = 180
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    'Looping to generate Label control on the form
    For i = 0 To frmWizard.lstSelectedFields(0).ListCount - 1
      rtfLap1.Text = rtfLap1.Text & _
      "      Begin VB.Label lblLabels" & vbCrLf & _
      "         BackStyle = 0          'Transparent" & vbCrLf & _
      "         Caption = """ & frmWizard.lstSelectedFields(0).List(i) & ":""" & vbCrLf & _
      "         Height = 255" & vbCrLf & _
      "         Index = " & i & "" & vbCrLf & _
      "         Left = 240" & vbCrLf & _
      "         TabIndex = " & nTabIndex & "" & vbCrLf & _
      "         Top = " & nTop & "" & vbCrLf & _
      "         Width = 1700" & vbCrLf & _
      "      End" & vbCrLf
      nTop = nTop + intSpace
      nTabIndex = nTabIndex + 1
    Next i
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
  '----------- END OF GENERATE LABEL ------------
  
  DoEvents
  frmWizard.lblProcess.Caption = "Generating textbox control..."
  DoEvents
  
  '----------- START GENERATE TEXTBOX -----------
  'Save again to temporary file
  Open strFileName For Output As #1
    nTop = 180
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    'Looping to generate Textbox control on the form
    For i = 0 To frmWizard.lstSelectedFields(0).ListCount - 1
      rtfLap1.Text = rtfLap1.Text & _
      "      Begin VB.TextBox txtFields" & vbCrLf & _
      "         DataField = """ & frmWizard.lstSelectedFields(0).List(i) & """" & vbCrLf & _
      "         Height = 285" & vbCrLf & _
      "         Index = " & i & "" & vbCrLf & _
      "         Left = 2040" & vbCrLf & _
      "         TabIndex = " & nTabIndex & "" & vbCrLf & _
      "         Top = " & nTop & "" & vbCrLf & _
      "         Width = 4230" & vbCrLf & _
      "      End" & vbCrLf
      nTop = nTop + intSpace
      nTabIndex = nTabIndex + 1
    Next i
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
    
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)  '3225
    rtfLap1.Text = rtfLap1.Text & _
    "   Begin VB.PictureBox picButtons" & vbCrLf & _
    "      Height = 4665" & vbCrLf & _
    "      Left = 6480" & vbCrLf & _
    "      ScaleHeight = 3165" & vbCrLf & _
    "      ScaleWidth = 1335" & vbCrLf & _
    "      TabIndex = 100" & vbCrLf & _
    "      Top = 180" & vbCrLf & _
    "      Width = 1395" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  
  PutCodeToRichTextBox
  '----------- END OF GENERATE TEXTBOX --------
  
  DoEvents
  frmWizard.lblProcess.Caption = "Generating button control..."
  DoEvents
  
  '----------- START GENERATE MAIN BUTTON ------
  Dim nTopBtn As Long
  Dim intSpaceBtn As Long
  intSpaceBtn = 350 * 1.03   'cmdButton.Height = 300
  'Save again to temporary file
  Open strFileName For Output As #1
    nTopBtn = 120
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    'Looping to generate Commanbutton Control
    For i = 0 To frmWizard.chkButton.UBound
      If frmWizard.chkButton(i).Value = 1 Then
      rtfLap1.Text = rtfLap1.Text & _
      "      Begin VB.CommandButton " & frmWizard.chkButton(i).Tag & "" & vbCrLf & _
      "         Caption = ""&" & Right(frmWizard.chkButton(i).Tag, Len(frmWizard.chkButton(i).Tag) - 3) & """" & vbCrLf & _
      "         Height = 350" & vbCrLf & _
      "         Left = 120" & vbCrLf & _
      "         TabIndex = " & nTabIndex + 1 & "" & vbCrLf & _
      "         Top = " & nTopBtn & "" & vbCrLf & _
      "         Width = 1095" & vbCrLf & _
      "      End" & vbCrLf
      nTopBtn = nTopBtn + intSpaceBtn
      nTabIndex = nTabIndex + 1
      End If
    Next i
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
  
  Open strFileName For Output As #1
    nTopBtn = 120
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "   End" & vbCrLf      '<----- akhir button
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
  '-------- END OF GENERATE MAIN BUTTON ---------
    
  nTop = nTop + 50
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "   Begin MSDataGridLib.DataGrid grdDataGrid" & vbCrLf & _
    "      Height = 2000" & vbCrLf & _
    "      Left = 240" & vbCrLf & _
    "      TabIndex = " & nTabIndex + 1 & "" & vbCrLf & _
    "      TabStop = 0             'False" & vbCrLf & _
    "      Top = " & nTop & "" & vbCrLf & _
    "      Width = 6015" & vbCrLf & _
    "      _ExtentX        =   10610" & vbCrLf & _
    "      _ExtentY        =   3519" & vbCrLf & _
    "      _Version        =   393216" & vbCrLf & _
    "      AllowUpdate = 0         'False" & vbCrLf & _
    "      HeadLines = 1" & vbCrLf & _
    "      RowHeight = 15" & vbCrLf & _
    "      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851}" & vbCrLf & _
    "         Name = ""MS Sans Serif""" & vbCrLf & _
    "         Size = 8.25" & vbCrLf & _
    "         Charset = 0" & vbCrLf & _
    "         Weight = 400" & vbCrLf & _
    "         Underline = 0           'False" & vbCrLf & _
    "         Italic = 0              'False" & vbCrLf & _
    "         Strikethrough = 0       'False" & vbCrLf & _
    "      EndProperty" & vbCrLf & _
    "      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}" & vbCrLf & _
    "         Name = ""MS Sans Serif""" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "         Size = 8.25" & vbCrLf & _
    "         Charset = 0" & vbCrLf & _
    "         Weight = 400" & vbCrLf & _
    "         Underline = 0           'False" & vbCrLf & _
    "         Italic = 0              'False" & vbCrLf & _
    "         Strikethrough = 0       'False" & vbCrLf & _
    "      EndProperty" & vbCrLf & _
    "      ColumnCount = 2" & vbCrLf & _
    "      BeginProperty Column00" & vbCrLf & _
    "         DataField = """ & vbCrLf & _
    "         Caption = """ & vbCrLf & _
    "         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED}" & vbCrLf & _
    "            Type            =   0" & vbCrLf & _
    "            Format = """ & vbCrLf & _
    "            HaveTrueFalseNull = 0" & vbCrLf & _
    "            FirstDayOfWeek = 0" & vbCrLf & _
    "            FirstWeekOfYear = 0" & vbCrLf & _
    "            LCID = 1057" & vbCrLf & _
    "            SubFormatType = 0" & vbCrLf & _
    "         EndProperty" & vbCrLf & _
    "      EndProperty" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "      BeginProperty Column01" & vbCrLf & _
    "         DataField = """"" & vbCrLf & _
    "         Caption = """"" & vbCrLf & _
    "         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED}" & vbCrLf & _
    "            Type            =   0" & vbCrLf & _
    "            Format = """ & vbCrLf & _
    "            HaveTrueFalseNull = 0" & vbCrLf & _
    "            FirstDayOfWeek = 0" & vbCrLf & _
    "            FirstWeekOfYear = 0" & vbCrLf & _
    "            LCID = 1057" & vbCrLf & _
    "            SubFormatType = 0" & vbCrLf & _
    "         EndProperty" & vbCrLf & _
    "      EndProperty" & vbCrLf & _
    "      SplitCount = 1" & vbCrLf & _
    "      BeginProperty Split0" & vbCrLf & _
    "         BeginProperty Column00" & vbCrLf & _
    "         EndProperty" & vbCrLf & _
    "         BeginProperty Column01" & vbCrLf & _
    "         EndProperty" & vbCrLf & _
    "      EndProperty" & vbCrLf & _
    "   End" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  
  PutCodeToRichTextBox
  
  nTop = nTop + 50 + 2000  '<-- 2000 = height of DataGrid
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "   Begin VB.Label lblAngka" & vbCrLf & _
    "      Alignment = 1          'Right Justify" & vbCrLf & _
    "      BackStyle = 0          'Transparent" & vbCrLf & _
    "      Height = 255" & vbCrLf & _
    "      Left = 4800" & vbCrLf & _
    "      TabIndex = 41" & vbCrLf & _
    "      Top = " & nTop & "" & vbCrLf & _
    "      Width = 1455" & vbCrLf & _
    "   End" & vbCrLf & _
    "   Begin VB.Label lblField" & vbCrLf & _
    "      BackStyle = 0          'Transparent" & vbCrLf & _
    "      Height = 255" & vbCrLf & _
    "      Left = 240" & vbCrLf & _
    "      TabIndex = 42" & vbCrLf & _
    "      Top = " & nTop & "" & vbCrLf & _
    "      Width = 2655" & vbCrLf & _
    "   End" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
  
  nTop = nTop + 100 + 255  '<--- 255 = Height of lblAngka
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "   Begin ComctlLib.ProgressBar prgBar1" & vbCrLf & _
    "      Height = 180" & vbCrLf & _
    "      Left = 240" & vbCrLf & _
    "      TabIndex = " & nTabIndex + 1 & "" & vbCrLf & _
    "      Top = " & nTop & "" & vbCrLf & _
    "      Width = 6015" & vbCrLf & _
    "      _ExtentX        =   10610" & vbCrLf & _
    "      _ExtentY        =   318" & vbCrLf & _
    "      _Version        =   327682" & vbCrLf & _
    "      Appearance = 1" & vbCrLf & _
    "   End" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
     
  nTop = nTop + 150 + 180  '<--- 180 = Height of progressbar
  'This will generate navigation button
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "   Begin VB.PictureBox picStatBox" & vbCrLf & _
    "      Height = 600" & vbCrLf & _
    "      Left = 240" & vbCrLf & _
    "      ScaleHeight = 540" & vbCrLf & _
    "      ScaleWidth = 6075" & vbCrLf & _
    "      TabIndex = 44" & vbCrLf & _
    "      Top = " & nTop & "" & vbCrLf & _
    "      Width = 6015" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "      Begin VB.CommandButton cmdFirst" & vbCrLf & _
    "         Caption = ""First""" & vbCrLf & _
    "         Height = 350" & vbCrLf & _
    "         Left = 120" & vbCrLf & _
    "         TabIndex = " & nTabIndex + 1 & "" & vbCrLf & _
    "         Top = 100" & vbCrLf & _
    "         UseMaskColor = -1       'True" & vbCrLf & _
    "         Width = 705" & vbCrLf & _
    "      End" & vbCrLf & _
    "      Begin VB.CommandButton cmdPrevious" & vbCrLf & _
    "         Caption = ""Prev""" & vbCrLf & _
    "         Height = 350" & vbCrLf & _
    "         Left = 840" & vbCrLf & _
    "         TabIndex = " & nTabIndex + 1 & "" & vbCrLf & _
    "         Top = 100" & vbCrLf & _
    "         UseMaskColor = -1       'True" & vbCrLf & _
    "         Width = 705" & vbCrLf & _
    "      End" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "      Begin VB.CommandButton cmdNext" & vbCrLf & _
    "         Caption = ""Next""" & vbCrLf & _
    "         Height = 350" & vbCrLf & _
    "         Left = 4440" & vbCrLf & _
    "         TabIndex = " & nTabIndex + 1 & "" & vbCrLf & _
    "         Top = 100" & vbCrLf & _
    "         UseMaskColor = -1       'True" & vbCrLf & _
    "         Width = 705" & vbCrLf & _
    "      End" & vbCrLf & _
    "      Begin VB.CommandButton cmdLast" & vbCrLf & _
    "         Caption = ""Last""" & vbCrLf & _
    "         Height = 350" & vbCrLf & _
    "         Left = 5160" & vbCrLf & _
    "         TabIndex = " & nTabIndex + 1 & "" & vbCrLf & _
    "         Top = 100" & vbCrLf & _
    "         UseMaskColor = -1       'True" & vbCrLf & _
    "         Width = 705" & vbCrLf & _
    "      End" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
    
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "      Begin VB.Label lblStatus" & vbCrLf & _
    "         Alignment = 2          'Center" & vbCrLf & _
    "         BackColor = &HFFFFFF" & vbCrLf & _
    "         BorderStyle = 1        'Fixed Single" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "         Height = 285" & vbCrLf & _
    "         Left = 1440" & vbCrLf & _
    "         TabIndex = " & nTabIndex + 1 & "" & vbCrLf & _
    "         Top = 120" & vbCrLf & _
    "         Width = 3120" & vbCrLf & _
    "      End" & vbCrLf & _
    "   End"
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
   
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "End"
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Attribute VB_Name = " & glo.strFormName & "" & vbCrLf & _
    "Attribute VB_GlobalNameSpace = False" & vbCrLf & _
    "Attribute VB_Creatable = False" & vbCrLf & _
    "Attribute VB_PredeclaredId = True" & vbCrLf & _
    "Attribute VB_Exposed = False"
     Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
  '------------- END OF GENERATE FORM ------------
    
  DoEvents
  frmWizard.lblProcess.Caption = "Generating code to form..."
  DoEvents
    
  DoEvents
  frmWizard.lblProcess.Caption = "Generating form header..."
  DoEvents
    
  '------ START GENERATE CODING IN THE FORM -----
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "'File Name  : " & glo.strFormName & ".frm" & vbCrLf & _
    "'Description: (put your description here)..." & vbCrLf & _
    "'             .............................." & vbCrLf & _
    "'             .............................." & vbCrLf & _
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
    "" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
  
  DoEvents
  frmWizard.lblProcess.Caption = "Generating variable declaration in form..."
  DoEvents
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Option Explicit" & vbCrLf & _
    "" & vbCrLf & _
    "'General variable for this module" & vbCrLf & _
    "Public WithEvents adoPrimaryRS As ADODB.Recordset" & vbCrLf & _
    "Attribute adoPrimaryRS.VB_VarHelpID = -1" & vbCrLf & _
    "Public WithEvents rsstrFindData As Recordset" & vbCrLf & _
    "Attribute rsstrFindData.VB_VarHelpID = -1" & vbCrLf & _
    "Dim mbChangedByCode As Boolean" & vbCrLf & _
    "Dim mvBookMark As Variant" & vbCrLf & _
    "Dim mbEditFlag As Boolean" & vbCrLf & _
    "Dim mbAddNewFlag As Boolean" & vbCrLf & _
    "Dim mbDataChanged As Boolean" & vbCrLf & _
    "Dim blnCancel As Boolean" & vbCrLf & _
    "Dim NumData As Integer" & vbCrLf & _
    "Dim intRecord As Integer" & vbCrLf & _
    "Dim intField As Integer" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
  
  DoEvents
  frmWizard.lblProcess.Caption = "Generating Form_Load procedure..."
  DoEvents
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Private Sub Form_Load()" & vbCrLf & _
    "On Error GoTo Message" & vbCrLf & _
    "INIFileName = App.Path & ""\Setting" & glo.strFormName & ".ini" & vbCrLf & _
    "  blnCancel = False" & vbCrLf & _
    "  OpenConnection" & vbCrLf & _
    "  Set adoPrimaryRS = New Recordset" & vbCrLf & _
    "  'We display all data in a datagrid below and underlying" & vbCrLf & _
    "  'source (the selected record in datagrid) above." & vbCrLf & _
    "  strSQL = """ & strSQLOpen & """" & vbCrLf & _
    "  adoPrimaryRS.Open strSQL, cnn, adOpenDynamic,  adLockOptimistic " & vbCrLf & _
    "  Dim oText As TextBox " & vbCrLf & _
    "  'Bind textbox to recordset " & vbCrLf & _
    "  For Each oText In Me.txtFields " & vbCrLf & _
    "    Set oText.DataSource = adoPrimaryRS " & vbCrLf & _
    "  Next" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
  
  If glo.strFormLayout = 0 Then 'ADO Code complete
   'Save again to temporary file
    Open strFileName For Output As #1
      'Get the new start of next record in richtexbox control
      rtfLap1.SelStart = Len(rtfLap1.Text)
      rtfLap1.Text = rtfLap1.Text & _
      "  'Bind recordset to datagrid " & vbCrLf & _
      "  Set grdDataGrid.DataSource = adoPrimaryRS.DataSource " & vbCrLf & _
      "  mbDataChanged = False" & vbCrLf
      Print #1, rtfLap1.Text
    Close #1
  Else  '<--- Real Master/Detail
   'Save again to temporary file
    Open strFileName For Output As #1
      'Get the new start of next record in richtexbox control
      rtfLap1.SelStart = Len(rtfLap1.Text)
      rtfLap1.Text = rtfLap1.Text & _
      "  'Bind recordset to datagrid " & vbCrLf & _
      "  Set grdDataGrid.DataSource = adoPrimaryRS(""ChildCMD"").UnderlyingValue " & vbCrLf & _
      "  mbDataChanged = False" & vbCrLf
      Print #1, rtfLap1.Text
    Close #1
  End If
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "  LockTheForm  'Lock textbox, and make datagrid enable" & vbCrLf & _
    "  grdDataGrid.Enabled = True" & vbCrLf & _
    "  'If we have no data in recordset" & vbCrLf & _
    "  If adoPrimaryRS.RecordCount < 1 Then " & vbCrLf & _
    "     MsgBox ""Recordset is empty. Please click Add button to add new record!"", vbExclamation, ""Empty Recordset""" & vbCrLf & _
    "     Exit Sub " & vbCrLf & _
    "  End If " & vbCrLf & _
    "  LockTheForm 'Lock textbox, combobox, and optionbutton" & vbCrLf & _
    "  'Except Datagrid...." & vbCrLf & _
    "  grdDataGrid.Enabled = True" & vbCrLf & _
    "  grdDataGrid.TabStop = False" & vbCrLf & _
    "  SetButtons True" & vbCrLf & _
    "  Exit Sub" & vbCrLf & _
    "Message:" & vbCrLf & _
    "  MsgBox Err.Number & "" - "" & Err.Description" & vbCrLf & _
    "  End" & vbCrLf & _
    "End Sub" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Private Sub Message(strMessage As String)" & vbCrLf & _
    "  StatusBar1.Panels(1).Text = strMessage" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
      
  If frmWizard.chkButton(10).Value = 1 Then 'cmdDataGrid
  DoEvents
  frmWizard.lblProcess.Caption = "Generating code for cmdDataGrid_Click procedure..."
  DoEvents
  
  If frmWizard.lstFormLayout.Text = frmWizard.lstFormLayout.List(0) Then
    'Save again to temporary file
    Open strFileName For Output As #1
      'Get the new start of next record in richtexbox control
      rtfLap1.SelStart = Len(rtfLap1.Text)
      rtfLap1.Text = rtfLap1.Text & _
      "Private Sub cmdDataGrid_Click()" & vbCrLf & _
      "  intRecord = adoPrimaryRS.RecordCount" & vbCrLf & _
      "  intField = adoPrimaryRS.Fields.Count - 1" & vbCrLf & _
      "  Call AdjustDataGridColumnWidth(grdDataGrid, adoPrimaryRS, _" & vbCrLf & _
      "                              intRecord, intField, True)" & vbCrLf & _
      "End Sub" & vbCrLf & _
      "Private Sub cmdDataGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf & _
      "  Call Message(""Adjust datagrid columns based on the longest field."")" & vbCrLf & _
      "End Sub" & vbCrLf & vbCrLf
    Print #1, rtfLap1.Text
    Close #1
  Else  'Master Detail, special treatment
    'Save again to temporary file
    Open strFileName For Output As #1
      'Get the new start of next record in richtexbox control
      rtfLap1.SelStart = Len(rtfLap1.Text)
      rtfLap1.Text = rtfLap1.Text & _
      "Private Sub cmdDataGrid_Click()" & vbCrLf & _
      "Dim rsChild As New ADODB.Recordset" & vbCrLf & _
      "Dim strSQLChild As String" & vbCrLf & _
      "  strSQLChild = ""SELECT " & strFieldChild & """ & vbCrLf & _ " & vbCrLf & _
      "  ""FROM " & glo.strRSDataGrid & """" & vbCrLf & _
      "  rsChild.Open strSQLChild, cnn " & vbCrLf & _
      "  intRecord = rsChild.RecordCount" & vbCrLf & _
      "  intField = rsChild.Fields.Count" & vbCrLf & _
      "  Call AdjustDataGridColumnWidth(grdDataGrid, rsChild, _" & vbCrLf & _
      "                              intRecord, intField, True)" & vbCrLf & _
      "End Sub" & vbCrLf & _
      "Private Sub cmdDataGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf & _
      "  Call Message(""Adjust datagrid columns based on the longest field."")" & vbCrLf & _
      "End Sub" & vbCrLf & vbCrLf
      Print #1, rtfLap1.Text
    Close #1
  End If
  End If
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Private Sub cmdFirst_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf & _
    "  Call Message(""Go to the first record."")" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf & _
    "Private Sub cmdLast_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf & _
    "  Call Message(""Go to the last record."")" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf & _
    "Private Sub cmdNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf & _
    "  Call Message(""Go to the next record."")" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf & _
    "Private Sub cmdPrevious_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf & _
    "  Call Message(""Go to the previous record."")" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Private Sub grdDataGrid_Error(ByVal DataError As Integer, Response As Integer)" & vbCrLf & _
    "  Response = -1" & vbCrLf & _
    "  'DataError = -1" & vbCrLf & _
    "End Sub" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
  
  If frmWizard.chkButton(1).Value = 1 And _
     frmWizard.chkButton(2).Value = 1 Then
  DoEvents
  frmWizard.lblProcess.Caption = "Generating code for Form_QueryUnload procedure..."
  DoEvents
    
    'Save again to temporary file
    Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)" & vbCrLf & _
    "  If cmdUpdate.Enabled = True And cmdCancel.Enabled = True Then" & vbCrLf & _
    "     MsgBox ""You have to save or cancel the changes "" & vbCrLf & _" & vbCrLf & _
    "            ""that you have just made before quit!"", _" & vbCrLf & _
    "            vbExclamation, ""Warning""" & vbCrLf & _
    "     cmdUpdate.SetFocus" & vbCrLf & _
    "     Cancel = -1" & vbCrLf & _
    "     Exit Sub" & vbCrLf & _
    "  End If" & vbCrLf
    Print #1, rtfLap1.Text
    Close #1
    PutCodeToRichTextBox
  Else
    'Save again to temporary file
    Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)" & vbCrLf
    Print #1, rtfLap1.Text
    Close #1
    PutCodeToRichTextBox
  End If
    
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "  If Not adoPrimaryRS Is Nothing Then _" & vbCrLf & _
    "    Set adoPrimaryRS = Nothing  'Clear memory from recordset" & vbCrLf & _
    "  'In order that prevent error from DataGrid...!" & vbCrLf & _
    "  If grdDataGrid.TabStop = True Then" & vbCrLf & _
    "     txtFields(0).SetFocus" & vbCrLf & _
    "  End If" & vbCrLf & _
    "  cnn.Close 'Close database" & vbCrLf & _
    "  Set cnn = Nothing  'Clear memory from database" & vbCrLf & _
    "  End" & vbCrLf & _
    "End Sub" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
    
  DoEvents
  frmWizard.lblProcess.Caption = "Generating code for adoPrimaryRS_MoveComplete procedure..."
  DoEvents
    
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Private Sub Form_Unload(Cancel As Integer)" & vbCrLf & _
    "  Screen.MousePointer = vbDefault 'Mouse pointer back to normal" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "" & vbCrLf & _
    "'Display the selected record in datagrid" & vbCrLf & _
    "Public Sub adoPrimaryRS_MoveComplete(ByVal adReason As _" & vbCrLf & _
    "            ADODB.EventReasonEnum, ByVal pError As _" & vbCrLf & _
    "            ADODB.Error, adStatus As ADODB.EventStatusEnum, _" & vbCrLf & _
    "            ByVal pRecordset As ADODB.Recordset)" & vbCrLf & _
    "  NumData = adoPrimaryRS.AbsolutePosition" & vbCrLf & _
    "  lblStatus.Caption = ""Record number "" & CStr(NumData) & "" from "" _" & vbCrLf & _
    "                      & adoPrimaryRS.RecordCount" & vbCrLf & _
    "  CheckNavigation" & vbCrLf & _
    "End Sub" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
  
  DoEvents
  frmWizard.lblProcess.Caption = "Generating code for CheckNavigation procedure..."
  DoEvents
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Private Sub CheckNavigation()" & vbCrLf & _
    "  'This will check which navigation button can be" & vbCrLf & _
    "  'accessed when you navigate the recordset through" & vbCrLf & _
    "  'Datagrid control or navigation button itself" & vbCrLf & _
    "  With adoPrimaryRS" & vbCrLf & _
    "   'If we have at least two record..." & vbCrLf & _
    "   If (.RecordCount > 1) Then" & vbCrLf & _
    "      'BOF = Begin Of Recordset" & vbCrLf & _
    "      If (.BOF) Or _" & vbCrLf & _
    "         (.AbsolutePosition = 1) Then" & vbCrLf & _
    "          cmdFirst.Enabled = False" & vbCrLf & _
    "          cmdPrevious.Enabled = False" & vbCrLf & _
    "          cmdNext.Enabled = True" & vbCrLf & _
    "          cmdLast.Enabled = True" & vbCrLf & _
    "      'EOF = End Of Recordset" & vbCrLf & _
    "      ElseIf (.EOF) Or _" & vbCrLf & _
    "          (.AbsolutePosition = .RecordCount) Then" & vbCrLf & _
    "          cmdNext.Enabled = False" & vbCrLf & _
    "          cmdLast.Enabled = False" & vbCrLf & _
    "          cmdFirst.Enabled = True" & vbCrLf & _
    "          cmdPrevious.Enabled = True" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "      Else" & vbCrLf & _
    "          cmdFirst.Enabled = True" & vbCrLf & _
    "          cmdPrevious.Enabled = True" & vbCrLf & _
    "          cmdNext.Enabled = True" & vbCrLf & _
    "          cmdLast.Enabled = True" & vbCrLf & _
    "      End If" & vbCrLf & _
    "   Else" & vbCrLf & _
    "      cmdFirst.Enabled = False" & vbCrLf & _
    "      cmdPrevious.Enabled = False" & vbCrLf & _
    "      cmdNext.Enabled = False" & vbCrLf & _
    "      cmdLast.Enabled = False" & vbCrLf & _
    "   End If" & vbCrLf & _
    " End With" & vbCrLf & _
    "End Sub" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
  
  If frmWizard.chkButton(0).Value = 1 Then 'cmdAdd
  DoEvents
  frmWizard.lblProcess.Caption = "Generating code for cmdAdd_Click procedure..."
  DoEvents
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Private Sub cmdAdd_Click()" & vbCrLf & _
    "  On Error GoTo AddErr" & vbCrLf & _
    "  With adoPrimaryRS" & vbCrLf & _
    "    If Not (.BOF And .EOF) Then" & vbCrLf & _
    "      mvBookMark = .Bookmark" & vbCrLf & _
    "    End If" & vbCrLf & _
    "    UnlockTheForm" & vbCrLf & _
    "    .AddNew" & vbCrLf & _
    "    lblStatus.Caption = ""Add record""" & vbCrLf & _
    "    mbAddNewFlag = True" & vbCrLf & _
    "    SetButtons False" & vbCrLf & _
    "  End With" & vbCrLf & _
    "  grdDataGrid.Enabled = False  'In order that prevent error" & vbCrLf & _
    "  On Error Resume Next" & vbCrLf & _
    "  txtFields(0).SetFocus" & vbCrLf & _
    "  Exit Sub" & vbCrLf & _
    "AddErr:" & vbCrLf & _
    "  MsgBox Err.Description" & vbCrLf & _
    "End Sub" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf & _
    "  Call Message(""Add new record."")" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
  End If  'End of checking event procedure cmdAdd
  
  If frmWizard.chkButton(4).Value = 1 Then 'cmdDelete
  DoEvents
  frmWizard.lblProcess.Caption = "Generating code for cmdDelete_Click procedure..."
  DoEvents
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Private Sub cmdDelete_Click()" & vbCrLf & _
    "  On Error GoTo DeleteErr" & vbCrLf & _
    "  If adoPrimaryRS.RecordCount < 1 Then " & vbCrLf & _
    "     MsgBox ""Recordset is empty. Please click Add button to add new record!"", vbExclamation, ""Empty Recordset""" & vbCrLf & _
    "     Exit Sub " & vbCrLf & _
    "  End If " & vbCrLf & _
    "  If MsgBox(""Are you sure you want to delete this record?"", _" & vbCrLf & _
    "            vbQuestion + vbYesNo + vbDefaultButton2, _" & vbCrLf & _
    "            ""Delete Record"") _" & vbCrLf & _
    "            <> vbYes Then" & vbCrLf & _
    "     Exit Sub" & vbCrLf & _
    "  End If" & vbCrLf & _
    "  With adoPrimaryRS" & vbCrLf & _
    "    .Delete" & vbCrLf & _
    "    .MoveNext" & vbCrLf & _
    "    If .EOF Then .MoveLast" & vbCrLf & _
    "  End With" & vbCrLf & _
    "  Exit Sub" & vbCrLf & _
    "DeleteErr:" & vbCrLf & _
    "  MsgBox Err.Description" & vbCrLf & _
    "End Sub" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Private Sub cmdDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf & _
    "  Call Message(""Delete the selected record."")" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
  End If 'End of checking cmdDelete
  
If frmWizard.chkButton(5).Value = 1 Or _
   frmWizard.chkButton(2).Value = 1 Then
  DoEvents
  frmWizard.lblProcess.Caption = "Generating code for cmdRefresh_Click procedure..."
  DoEvents
   
   'cmdRefresh or cmdCancel true, display Refresh!
  If glo.strFormLayout = 0 Then 'ADO Code complete
    'Save again to temporary file
    Open strFileName For Output As #1
      'Get the new start of next record in richtexbox control
      rtfLap1.SelStart = Len(rtfLap1.Text)
      rtfLap1.Text = rtfLap1.Text & _
      "Private Sub cmdRefresh_Click()" & vbCrLf & _
      "  'Refresh is very important in multiuser app" & vbCrLf & _
      "  On Error GoTo RefreshErr" & vbCrLf & _
      "  If blnCancel = True Then" & vbCrLf & _
      "     SetButtons True" & vbCrLf & _
      "     blnCancel = False" & vbCrLf & _
      "  End If" & vbCrLf & _
      "  LockTheForm" & vbCrLf & _
      "  Set grdDataGrid.DataSource = Nothing" & vbCrLf & _
      "  Set adoPrimaryRS = New Recordset" & vbCrLf & _
      "  strSQL = """ & strSQLOpen & """" & vbCrLf & _
      "  adoPrimaryRS.Open strSQL, cnn, adOpenDynamic,  adLockOptimistic " & vbCrLf & _
      "  Dim oText As TextBox" & vbCrLf & _
      "  For Each oText In Me.txtFields" & vbCrLf & _
      "    Set oText.DataSource = adoPrimaryRS" & vbCrLf & _
      "  Next" & vbCrLf & _
      "  cmdBookmark.Enabled = True" & vbCrLf
      Print #1, rtfLap1.Text
    Close #1
  Else  '<-- Real Master/Detail
    'Save again to temporary file
    Open strFileName For Output As #1
      'Get the new start of next record in richtexbox control
      rtfLap1.SelStart = Len(rtfLap1.Text)
      rtfLap1.Text = rtfLap1.Text & _
      "Private Sub cmdRefresh_Click()" & vbCrLf & _
      "  'This is only needed for multi user apps" & vbCrLf & _
      "  On Error GoTo RefreshErr" & vbCrLf & _
      "  Set grdDataGrid.DataSource = Nothing" & vbCrLf & _
      "  Set adoPrimaryRS = New Recordset" & vbCrLf & _
      "  strSQL = """ & strSQLOpen & """" & vbCrLf & _
      "  adoPrimaryRS.Open strSQL, cnn, adOpenDynamic,  adLockOptimistic " & vbCrLf & _
      "  Dim oText As TextBox" & vbCrLf & _
      "  For Each oText In Me.txtFields" & vbCrLf & _
      "    Set oText.DataSource = adoPrimaryRS" & vbCrLf & _
      "  Next" & vbCrLf & _
      "  'Bind recordset to datagrid " & vbCrLf & _
      "  Set grdDataGrid.DataSource = adoPrimaryRS(""ChildCMD"").UnderlyingValue " & vbCrLf & _
      "  grdDataGrid.Enabled = True" & vbCrLf & _
      "  cmdBookmark.Enabled = True" & vbCrLf & _
      "  Exit Sub" & vbCrLf & _
      "RefreshErr:" & vbCrLf & _
      "  MsgBox Err.Description" & vbCrLf & _
      "End Sub" & vbCrLf
      Print #1, rtfLap1.Text
    Close #1
  End If
  PutCodeToRichTextBox
  
  If glo.strFormLayout = 0 Then 'ADO Code complete
    'Save again to temporary file
    Open strFileName For Output As #1
      'Get the new start of next record in richtexbox control
      rtfLap1.SelStart = Len(rtfLap1.Text)
      rtfLap1.Text = rtfLap1.Text & _
      "  'Bind recordset to datagrid " & vbCrLf & _
      "  Set grdDataGrid.DataSource = adoPrimaryRS.DataSource " & vbCrLf & _
      "  grdDataGrid.Enabled = True" & vbCrLf & _
      "  Exit Sub" & vbCrLf
    Print #1, rtfLap1.Text
    Close #1
    PutCodeToRichTextBox

    'Save again to temporary file
    Open strFileName For Output As #1
     'Get the new start of next record in richtexbox control
      rtfLap1.SelStart = Len(rtfLap1.Text)
      rtfLap1.Text = rtfLap1.Text & _
      "RefreshErr:" & vbCrLf & _
      "  mbEditFlag = False" & vbCrLf & _
      "  mbAddNewFlag = False" & vbCrLf & _
      "  adoPrimaryRS.CancelUpdate" & vbCrLf & _
      "  If mvBookMark <> 0 Then" & vbCrLf & _
      "      adoPrimaryRS.Bookmark = mvBookMark" & vbCrLf & _
      "  Else" & vbCrLf & _
      "      adoPrimaryRS.MoveFirst" & vbCrLf & _
      "  End If" & vbCrLf & _
      "  mbDataChanged = False" & vbCrLf & _
      "  blnCancel = True" & vbCrLf & _
      "  cmdRefresh_Click  'Automatically refresh" & vbCrLf & _
      "  Exit Sub" & vbCrLf & _
      "End Sub" & vbCrLf
    Print #1, rtfLap1.Text
    Close #1
  End If
  PutCodeToRichTextBox
    
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Private Sub cmdRefresh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf & _
    "  Call Message(""Retrieve all records from database."")" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
End If  'End of checking cmdRefresh
  
If frmWizard.chkButton(3).Value = 1 Then 'cmdEdit
  DoEvents
  frmWizard.lblProcess.Caption = "Generating code for cmdEdit_Click procedure..."
  DoEvents
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Private Sub cmdEdit_Click()" & vbCrLf & _
    "  On Error GoTo EditErr" & vbCrLf & _
    "  If adoPrimaryRS.RecordCount < 1 Then " & vbCrLf & _
    "     MsgBox ""Recordset is empty. Please click Add button to add new record!"", vbExclamation, ""Empty Recordset""" & vbCrLf & _
    "     Exit Sub " & vbCrLf & _
    "  End If " & vbCrLf & _
    "  lblStatus.Caption = ""Edit record""" & vbCrLf & _
    "  mbEditFlag = True" & vbCrLf & _
    "  SetButtons False" & vbCrLf & _
    "  UnlockTheForm 'Unlock textbox; we can edit data" & vbCrLf & _
    "  txtFields(0).SetFocus: SendKeys ""{Home}+{End}""" & vbCrLf & _
    "  Exit Sub" & vbCrLf & _
    "EditErr:" & vbCrLf & _
    "  MsgBox Err.Description" & vbCrLf & _
    "End Sub" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Private Sub cmdEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf & _
    "  Call Message(""Edit the selected record."")" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
End If  'End of checking cmdEdit
  
  
If frmWizard.chkButton(2).Value = 1 Then 'cmdCancel
  DoEvents
  frmWizard.lblProcess.Caption = "Generating code for cmdCancel_Click procedure..."
  DoEvents
  
  If glo.strFormLayout = 0 Then 'ADO Code complete
    'Save again to temporary file
    Open strFileName For Output As #1
      'Get the new start of next record in richtexbox control
      rtfLap1.SelStart = Len(rtfLap1.Text)
      rtfLap1.Text = rtfLap1.Text & _
      "Private Sub cmdCancel_Click()" & vbCrLf & _
      "  On Error Resume Next" & vbCrLf & _
      "  LockTheForm" & vbCrLf & _
      "  cmdRefresh_Click" & vbCrLf & _
      "  grdDataGrid.Enabled = True" & vbCrLf & _
      "  If blnCancel = True Then" & vbCrLf & _
      "     Exit Sub" & vbCrLf & _
      "  End If" & vbCrLf & _
      "  SetButtons True" & vbCrLf & _
      "  mbEditFlag = False" & vbCrLf & _
      "  mbAddNewFlag = False" & vbCrLf & _
      "  adoPrimaryRS.CancelUpdate" & vbCrLf & _
      "  If mvBookMark > 0 Then" & vbCrLf & _
      "    adoPrimaryRS.Bookmark = mvBookMark" & vbCrLf & _
      "  Else" & vbCrLf & _
      "    adoPrimaryRS.MoveFirst" & vbCrLf & _
      "  End If" & vbCrLf & _
      "  LockTheForm    'Lock textbox" & vbCrLf & _
      "  grdDataGrid.Enabled = True" & vbCrLf & _
      "  mbDataChanged = False" & vbCrLf & _
      "End Sub" & vbCrLf
      Print #1, rtfLap1.Text
    Close #1
  Else '<--- Real Master/Detail
    'Save again to temporary file
    Open strFileName For Output As #1
      'Get the new start of next record in richtexbox control
      rtfLap1.SelStart = Len(rtfLap1.Text)
      rtfLap1.Text = rtfLap1.Text & _
      "Private Sub cmdCancel_Click()" & vbCrLf & _
      "  On Error Resume Next" & vbCrLf & _
      "  LockTheForm" & vbCrLf & _
      "  If blnCancel = True Then" & vbCrLf & _
      "     Exit Sub" & vbCrLf & _
      "  End If" & vbCrLf & _
      "  SetButtons True" & vbCrLf & _
      "  mbEditFlag = False" & vbCrLf & _
      "  mbAddNewFlag = False" & vbCrLf & _
      "  grdDataGrid.Enabled = True" & vbCrLf & _
      "  adoPrimaryRS.CancelUpdate" & vbCrLf & _
      "  If mvBookMark > 0 Then" & vbCrLf & _
      "    adoPrimaryRS.Bookmark = mvBookMark" & vbCrLf & _
      "  Else" & vbCrLf & _
      "    adoPrimaryRS.MoveFirst" & vbCrLf & _
      "  End If" & vbCrLf & _
      "  LockTheForm    'Lock textbox" & vbCrLf & _
      "  grdDataGrid.Enabled = True" & vbCrLf & _
      "  mbDataChanged = False" & vbCrLf & _
      "End Sub" & vbCrLf
      Print #1, rtfLap1.Text
    Close #1
  End If
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf & _
    "  Call Message(""Cancel the change or new record that have not been saved."")" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf
    Print #1, rtfLap1.Text
    Close #1
    PutCodeToRichTextBox
End If 'End of checking cmdCancel
  

If frmWizard.chkButton(1).Value = 1 Then 'cmdUpdate
  DoEvents
  frmWizard.lblProcess.Caption = "Generating code for cmdUpdate_Click procedure..."
  DoEvents
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Private Sub cmdUpdate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf & _
    "  Call Message(""Save the change or new record."")" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox

  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Private Sub cmdUpdate_Click()" & vbCrLf & _
    "Dim i As Integer" & vbCrLf & _
    "  On Error GoTo UpdateErr" & vbCrLf & _
    "  For i = 0 To " & glo.intNumOfFields - 1 & "" & vbCrLf & _
    "    If txtFields(i).Text = """" Then" & vbCrLf & _
    "       MsgBox ""You have to fill in all textbox!"", _" & vbCrLf & _
    "              vbExclamation, ""Validation""" & vbCrLf & _
    "       txtFields(i).SetFocus" & vbCrLf & _
    "       Exit Sub" & vbCrLf & _
    "     End If" & vbCrLf & _
    "  Next i" & vbCrLf & _
    "  'Update by using UpdateBatch. UpdateBatch will" & vbCrLf & _
    "  'automatically update all data in various fields type." & vbCrLf & _
    "  adoPrimaryRS.UpdateBatch adAffectAll" & vbCrLf & _
    "  'Move pointer to last record if we just added data" & vbCrLf & _
    "  If mbAddNewFlag Then" & vbCrLf & _
    "    adoPrimaryRS.MoveLast" & vbCrLf & _
    "  End If" & vbCrLf & _
    "  If mbEditFlag Then" & vbCrLf & _
    "    adoPrimaryRS.MoveNext" & vbCrLf & _
    "    adoPrimaryRS.MovePrevious" & vbCrLf & _
    "  End If" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox

  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "  'Update all status" & vbCrLf & _
    "  mbEditFlag = False" & vbCrLf & _
    "  mbAddNewFlag = False" & vbCrLf & _
    "  SetButtons True" & vbCrLf & _
    "  mbDataChanged = False" & vbCrLf & _
    "  LockTheForm  'Lock textbox" & vbCrLf & _
    "  grdDataGrid.Enabled = True" & vbCrLf & _
    "  'Display the record position" & vbCrLf & _
    "  NumData = adoPrimaryRS.AbsolutePosition" & vbCrLf & _
    "  lblStatus.Caption = ""Record number "" & CStr(NumData) & "" from "" _" & vbCrLf & _
    "                      & adoPrimaryRS.RecordCount" & vbCrLf & _
    "  Exit Sub" & vbCrLf & _
    "UpdateErr:" & vbCrLf & _
    "  MsgBox Err.Number & "" - "" & _ " & vbCrLf & _
    "         Err.Description, vbCritical, ""Error Occured""" & vbCrLf & _
    "End Sub" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
End If 'End of checking cmdUpdate

If frmWizard.chkButton(11).Value = 1 Then 'cmdClose
  DoEvents
  frmWizard.lblProcess.Caption = "Generating code for cmdClose_Click procedure..."
  DoEvents
    
    'Save again to temporary file
    Open strFileName For Output As #1
      'Get the new start of next record in richtexbox control
      rtfLap1.SelStart = Len(rtfLap1.Text)
      rtfLap1.Text = rtfLap1.Text & _
    "Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf & _
    "  Call Message(""Quit from this program now."")" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Private Sub cmdClose_Click()" & vbCrLf & _
    "  Unload Me" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
End If  'End of Checking cmdClose
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Private Sub cmdFirst_Click()" & vbCrLf & _
    "  On Error GoTo GoFirstError" & vbCrLf
    If frmWizard.chkButton(7).Value = 1 Then
     rtfLap1.Text = rtfLap1.Text & _
     "If adoFilter Is Nothing Then" & vbCrLf & _
     "   adoPrimaryRS.MoveFirst" & vbCrLf & _
     "Else" & vbCrLf & _
     "   adoFilter.MoveFirst" & vbCrLf & _
     "End If" & vbCrLf
    Else
     rtfLap1.Text = rtfLap1.Text & _
     "   adoPrimaryRS.MoveFirst" & vbCrLf
    End If
    rtfLap1.Text = rtfLap1.Text & _
    "  mbDataChanged = False" & vbCrLf & _
    "  Exit Sub" & vbCrLf & _
    "GoFirstError:" & vbCrLf & _
    "  MsgBox Err.Description" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "" & vbCrLf & _
    "Private Sub cmdLast_Click()" & vbCrLf
    If frmWizard.chkButton(7).Value = 1 Then
     rtfLap1.Text = rtfLap1.Text & _
     "If adoFilter Is Nothing Then" & vbCrLf & _
     "   adoPrimaryRS.MoveLast" & vbCrLf & _
     "Else" & vbCrLf & _
     "   adoFilter.MoveLast" & vbCrLf & _
     "End If" & vbCrLf
    Else
     rtfLap1.Text = rtfLap1.Text & _
     "   adoPrimaryRS.MoveLast" & vbCrLf
    End If
     rtfLap1.Text = rtfLap1.Text & _
    "  mbDataChanged = False" & vbCrLf & _
    "  Exit Sub" & vbCrLf & _
    "GoLastError:" & vbCrLf & _
    "  MsgBox Err.Description" & vbCrLf & _
    "End Sub" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1

  PutCodeToRichTextBox

  DoEvents
  frmWizard.lblProcess.Caption = "Generating code for navigation procedure..."
  DoEvents

  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Private Sub cmdNext_Click()" & vbCrLf & _
    "  On Error GoTo GoNextError" & vbCrLf
    If frmWizard.chkButton(7).Value = 1 Then
     rtfLap1.Text = rtfLap1.Text & _
     "If adoFilter Is Nothing Then" & vbCrLf & _
     "   If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext" & vbCrLf & _
     "   If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then" & vbCrLf & _
     "      Beep" & vbCrLf & _
     "      adoPrimaryRS.MoveLast" & vbCrLf & _
     "      MsgBox ""This is the last record."", _" & vbCrLf & _
     "             vbInformation , ""Last Record""" & vbCrLf & _
     "   End If" & vbCrLf & _
     "Else" & vbCrLf & _
     "   If Not adoFilter.EOF Then adoFilter.MoveNext" & vbCrLf & _
     "   If adoFilter.EOF And adoFilter.RecordCount > 0 Then" & vbCrLf & _
     "      Beep" & vbCrLf & _
     "      adoFilter.MoveLast" & vbCrLf & _
     "      MsgBox ""This is the last record."", _" & vbCrLf & _
     "             vbInformation, ""Last Record""" & vbCrLf & _
     "   End If" & vbCrLf & _
     "End If" & vbCrLf
    Else
     rtfLap1.Text = rtfLap1.Text & _
     "   If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext" & vbCrLf & _
     "   If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then" & vbCrLf & _
     "      Beep" & vbCrLf & _
     "      adoPrimaryRS.MoveLast" & vbCrLf & _
     "      MsgBox ""This is the last record."", _" & vbCrLf & _
     "             vbInformation , ""Last Record""" & vbCrLf & _
     "   End If" & vbCrLf
    End If
    rtfLap1.Text = rtfLap1.Text & _
    "  mbDataChanged = False" & vbCrLf & _
    "  Exit Sub" & vbCrLf & _
    "GoNextError:" & vbCrLf & _
    "  MsgBox Err.Description" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf
    Print #1, rtfLap1.Text
  Close #1

  PutCodeToRichTextBox

  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Private Sub cmdPrevious_Click()" & vbCrLf & _
    "  On Error GoTo GoPrevError" & vbCrLf
    If frmWizard.chkButton(7).Value = 1 Then
     rtfLap1.Text = rtfLap1.Text & _
     "If adoFilter Is Nothing Then" & vbCrLf & _
     "   If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious" & vbCrLf & _
     "   If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then" & vbCrLf & _
     "      Beep" & vbCrLf & _
     "      adoPrimaryRS.MoveFirst" & vbCrLf & _
     "      MsgBox ""This is the first record."", _" & vbCrLf & _
     "             vbInformation, ""First Record""" & vbCrLf & _
     "   End If" & vbCrLf & _
     "Else" & vbCrLf & _
     "   If Not adoFilter.BOF Then adoFilter.MovePrevious" & vbCrLf & _
     "   If adoFilter.BOF And adoFilter.RecordCount > 0 Then" & vbCrLf & _
     "      Beep" & vbCrLf & _
     "      adoFilter.MoveFirst" & vbCrLf & _
     "      MsgBox ""This is the first record."", _" & vbCrLf & _
     "             vbInformation, ""First Record""" & vbCrLf & _
     "   End If" & vbCrLf & _
     "End If" & vbCrLf
    Else
     rtfLap1.Text = rtfLap1.Text & _
     "   If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious" & vbCrLf & _
     "   If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then" & vbCrLf & _
     "      Beep" & vbCrLf & _
     "      adoPrimaryRS.MoveFirst" & vbCrLf & _
     "      MsgBox ""This is the first record."", _" & vbCrLf & _
     "             vbInformation, ""First Record""" & vbCrLf & _
     "   End If" & vbCrLf
    End If
    rtfLap1.Text = rtfLap1.Text & _
    "  mbDataChanged = False" & vbCrLf & _
    "  Exit Sub" & vbCrLf & _
    "GoPrevError:" & vbCrLf & _
    "  MsgBox Err.Description" & vbCrLf & _
    "End Sub" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox

  DoEvents
  frmWizard.lblProcess.Caption = "Generating code for SetButtons procedure..."
  DoEvents

  'STARTING for SETBUTTONS PROCEDURE
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Private Sub SetButtons(bVal As Boolean)" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    DoEvents
    frmWizard.lblProcess.Caption = "Checking main buttons..."
    DoEvents
    
    'LOOPING FOR CHECKING COMMANDBUTTON ACCESSING
    For i = 0 To frmWizard.chkButton.UBound
      If frmWizard.chkButton(i).Value = 1 Then
        If frmWizard.chkButton(i).Tag = "cmdUpdate" Or _
           frmWizard.chkButton(i).Tag = "cmdCancel" Then
           rtfLap1.Text = rtfLap1.Text & _
           "  " & frmWizard.chkButton(i).Tag & ".Enabled = Not bVal" & vbCrLf
        Else
           rtfLap1.Text = rtfLap1.Text & _
           "  " & frmWizard.chkButton(i).Tag & ".Enabled = bVal" & vbCrLf
        End If
      End If
    Next i
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
  'END OF CHECKING COMMANDBUTTON ACCESSING

  'These 4 buttons are fixed...!
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "  cmdNext.Enabled = bVal" & vbCrLf & _
    "  cmdFirst.Enabled = bVal" & vbCrLf & _
    "  cmdLast.Enabled = bVal" & vbCrLf & _
    "  cmdPrevious.Enabled = bVal" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf & _
    "Private Sub picButtons_KeyPress(KeyAscii As Integer)" & vbCrLf & _
    "   If KeyAscii = 13 Then SendKeys ""{Tab}""" & vbCrLf & _
    "End Sub" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox

  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)" & vbCrLf & _
    "  Select Case Index  'If we hit Enter, jump to next textbox" & vbCrLf & _
    "         Case 0 To " & glo.intNumOfFields - 1 & "" & vbCrLf & _
    "              If KeyAscii = 13 Then SendKeys ""{Tab}""" & vbCrLf & _
    "  End Select" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "" & vbCrLf & _
    "'Lock textbox in order that we can't edit data" & vbCrLf & _
    "Private Sub LockTheForm()" & vbCrLf & _
    "Dim i As Integer" & vbCrLf & _
    "  For i = 0 To " & glo.intNumOfFields - 1 & "" & vbCrLf & _
    "    txtFields(i).Locked = True" & vbCrLf & _
    "  Next i" & vbCrLf & _
    "  grdDataGrid.Enabled = False" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "" & vbCrLf & _
    "'Unlock textbox in order that we can edit data" & vbCrLf & _
    "Sub UnlockTheForm()" & vbCrLf & _
    "Dim i As Integer" & vbCrLf & _
    "  For i = 0 To " & glo.intNumOfFields - 1 & "" & vbCrLf & _
    "    txtFields(i).Locked = False" & vbCrLf & _
    "  Next i" & vbCrLf & _
    "  grdDataGrid.Enabled = False" & vbCrLf & _
    "End Sub" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox

If frmWizard.chkButton(6).Value = 1 Then 'cmdFind
  DoEvents
  frmWizard.lblProcess.Caption = "Generating code for cmdFind_Click procedure..."
  DoEvents
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Private Sub cmdFind_Click()" & vbCrLf & _
    "On Error Resume Next" & vbCrLf & _
    "  Screen.MousePointer = vbHourglass" & vbCrLf
    If frmWizard.chkButton(9).Value = 1 Then
     rtfLap1.Text = rtfLap1.Text & _
     "  Set adoBookMark = Nothing" & vbCrLf
    End If
    If frmWizard.chkButton(7).Value = 1 Then
     rtfLap1.Text = rtfLap1.Text & _
     "  Set adoFilter = Nothing" & vbCrLf
    End If
    If frmWizard.chkButton(8).Value = 1 Then
     rtfLap1.Text = rtfLap1.Text & _
     "  Set adoSort = Nothing" & vbCrLf
    End If
    rtfLap1.Text = rtfLap1.Text & _
    "  Set adoFind = New ADODB.Recordset" & vbCrLf & _
    "  Set adoFind = adoPrimaryRS" & vbCrLf & _
    "  frmFind.Show , " & glo.strFormName & "" & vbCrLf & _
    "  Screen.MousePointer = vbDefault" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "Private Sub cmdFind_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf & _
    "  Call Message(""Find record (find first and find next)."")" & vbCrLf & _
    "End Sub" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
End If

If frmWizard.chkButton(7).Value = 1 Then 'cmdFilter
  DoEvents
  frmWizard.lblProcess.Caption = "Generating code for cmdFilter_Click procedure..."
  DoEvents
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Private Sub cmdFilter_Click()" & vbCrLf & _
    "On Error Resume Next" & vbCrLf & _
    "  Screen.MousePointer = vbHourglass" & vbCrLf
    If frmWizard.chkButton(9).Value = 1 Then
     rtfLap1.Text = rtfLap1.Text & _
     "  Set adoBookMark = Nothing" & vbCrLf
    End If
    If frmWizard.chkButton(6).Value = 1 Then
     rtfLap1.Text = rtfLap1.Text & _
     "  Set adoFind = Nothing" & vbCrLf
    End If
    If frmWizard.chkButton(8).Value = 1 Then
     rtfLap1.Text = rtfLap1.Text & _
     "  Set adoSort = Nothing" & vbCrLf
    End If
    rtfLap1.Text = rtfLap1.Text & _
    "  Set adoFilter = New ADODB.Recordset" & vbCrLf & _
    "  Set adoFilter = adoPrimaryRS" & vbCrLf & _
    "  frmFilter.Show , " & glo.strFormName & "" & vbCrLf & _
    "  Screen.MousePointer = vbDefault" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "Private Sub cmdFilter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf & _
    "  Call Message(""Filter recordset."")" & vbCrLf & _
    "End Sub" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
End If

  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Public Sub rsstrFindData_MoveComplete(ByVal adReason As _" & vbCrLf & _
    "            ADODB.EventReasonEnum, ByVal pError As _" & vbCrLf & _
    "            ADODB.Error, adStatus As ADODB.EventStatusEnum, _" & vbCrLf & _
    "            ByVal pRecordset As ADODB.Recordset)" & vbCrLf & _
    "    NumData = rsstrFindData.AbsolutePosition" & vbCrLf & _
    "    lblStatus.Caption = ""Record number "" & CStr(NumData) & "" from "" _" & vbCrLf & _
    "                      & rsstrFindData.RecordCount" & vbCrLf & _
    "End Sub" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox

If frmWizard.chkButton(8).Value = 1 Then 'cmdSort
  DoEvents
  frmWizard.lblProcess.Caption = "Generating code for cmdSort_Click procedure..."
  DoEvents
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Private Sub cmdSort_Click()" & vbCrLf & _
    "On Error Resume Next" & vbCrLf & _
    "  Screen.MousePointer = vbHourglass" & vbCrLf
    If frmWizard.chkButton(9).Value = 1 Then
     rtfLap1.Text = rtfLap1.Text & _
     "  Set adoBookMark = Nothing" & vbCrLf
    End If
    If frmWizard.chkButton(6).Value = 1 Then
     rtfLap1.Text = rtfLap1.Text & _
     "  Set adoFind = Nothing" & vbCrLf
    End If
    If frmWizard.chkButton(7).Value = 1 Then
     rtfLap1.Text = rtfLap1.Text & _
     "  Set adoFilter = Nothing" & vbCrLf
    End If
    rtfLap1.Text = rtfLap1.Text & _
    "  Set adoSort = New ADODB.Recordset" & vbCrLf & _
    "  Set adoSort = adoPrimaryRS" & vbCrLf & _
    "  frmSort.Show , " & glo.strFormName & "" & vbCrLf & _
    "  Screen.MousePointer = vbDefault" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "Private Sub cmdSort_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf & _
    "  Call Message(""Sort recordset."")" & vbCrLf & _
    "End Sub" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
  PutCodeToRichTextBox
End If

If frmWizard.chkButton(9).Value = 1 Then 'cmdBookmark
  DoEvents
  frmWizard.lblProcess.Caption = "Generating code for cmdBookmark_Click procedure..."
  DoEvents
  
  'Save again to temporary file
  Open strFileName For Output As #1
    'Get the new start of next record in richtexbox control
    rtfLap1.SelStart = Len(rtfLap1.Text)
    rtfLap1.Text = rtfLap1.Text & _
    "Private Sub cmdBookmark_Click()" & vbCrLf & _
    "On Error Resume Next" & vbCrLf & _
    "  Screen.MousePointer = vbHourglass" & vbCrLf
    If frmWizard.chkButton(8).Value = 1 Then
     rtfLap1.Text = rtfLap1.Text & _
     "  Set adoSort = Nothing" & vbCrLf
    End If
    If frmWizard.chkButton(6).Value = 1 Then
     rtfLap1.Text = rtfLap1.Text & _
     "  Set adoFind = Nothing" & vbCrLf
    End If
    If frmWizard.chkButton(7).Value = 1 Then
     rtfLap1.Text = rtfLap1.Text & _
     "  Set adoFilter = Nothing" & vbCrLf
    End If
    rtfLap1.Text = rtfLap1.Text & _
    "  Set adoBookmark = New ADODB.Recordset" & vbCrLf & _
    "  Set adoBookmark = adoPrimaryRS" & vbCrLf & _
    "  frmBookmark.Show , " & glo.strFormName & "" & vbCrLf & _
    "  Screen.MousePointer = vbDefault" & vbCrLf & _
    "End Sub" & vbCrLf & _
    "Private Sub cmdBookmark_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf & _
    "  Call Message(""Bookmark record so you can go back easily."")" & vbCrLf & _
    "End Sub" & vbCrLf
    Print #1, rtfLap1.Text
  Close #1
End If

End Sub
'----- End of generate ADO Code to the form -----

Private Sub GenerateSQLOpenRecordset()
  DoEvents
  frmWizard.lblProcess.Caption = "Generating connection code..."
  DoEvents
  Dim intFieldParent As Integer, intFieldChild As Integer
  Dim i As Integer, j As Integer
  
  'To get the criteria and to make SQL Statement
  'shorter, we can use this way...
  intFieldParent = frmWizard.lstSelectedFields(0).ListCount
  intFieldChild = frmWizard.lstSelectedFields(1).ListCount
  ReDim tabFieldParent(intFieldParent)
  ReDim tabFieldChild(intFieldChild)
  
  'Check Field Parent
  For i = 0 To frmWizard.lstSelectedFields(0).ListCount - 1
      tabFieldParent(i) = frmWizard.lstSelectedFields(0).List(i)
      'Check whether a field contains space
      'so, add the "[...]"
      If InStr(1, tabFieldParent(i), " ") > 0 Then
         tabFieldParent(i) = "[" & tabFieldParent(i) & "]"
      End If
  Next

  'Check Field Child
  For i = 0 To frmWizard.lstSelectedFields(1).ListCount - 1
      tabFieldChild(i) = frmWizard.lstSelectedFields(1).List(i)
      'Check whether a field contains space
      'so, add the "[...]"
      If InStr(1, tabFieldChild(i), " ") > 0 Then
         tabFieldChild(i) = "[" & tabFieldChild(i) & "]"
      End If
  Next

  strFieldParent = ""
  For i = 0 To intFieldParent - 1
     If i <> intFieldParent - 1 Then
        strFieldParent = strFieldParent & tabFieldParent(i) & ","
     Else
        strFieldParent = strFieldParent & tabFieldParent(i) & " "
     End If
  Next i
  
  strFieldChild = ""
  For i = 0 To intFieldChild - 1
     If i <> intFieldChild - 1 Then
        strFieldChild = strFieldChild & tabFieldChild(i) & ","
     Else
        strFieldChild = strFieldChild & tabFieldChild(i) & " "
     End If
  Next i
  
  
  If glo.strOrderTextBox = "(None)" Then
     strOrdTB = ""
  Else
     'Check whether a select field contains space
     'so, add the "[...]"
     If InStr(1, glo.strOrderTextBox, " ") > 0 Then
        glo.strOrderTextBox = "[" & glo.strOrderTextBox & "]"
     End If
     strOrdTB = "ORDER BY " & glo.strOrderTextBox
  End If
  
  If glo.strOrderDataGrid = "(None)" Then
     strOrdDG = ""
  Else
     'Check whether an order field contains space
     'so, add the "[...]"
     If InStr(1, glo.strOrderDataGrid, " ") > 0 Then
        glo.strOrderDataGrid = "[" & glo.strOrderDataGrid & "]"
     End If
     strOrdDG = "ORDER BY " & glo.strOrderDataGrid
  End If
  
  'Check whether a relation textbox field contains space
  'If yes, add the "[...]"
  If InStr(1, glo.strRelationTextBox, " ") > 0 Then
     glo.strRelationTextBox = "[" & glo.strRelationTextBox & "]"
  End If
  'Check whether a relation datagrid field contains space
  'If yes, add the "[...]"
  If InStr(1, glo.strRelationDataGrid, " ") > 0 Then
     glo.strRelationDataGrid = "[" & glo.strRelationDataGrid & "]"
  End If
     
  strSQLOpen = _
     "SHAPE " & _
     "{SELECT " & strFieldParent & """ & vbCrLf & _ " & vbCrLf & _
     "  ""FROM " & glo.strRSTextBox & " " & _
     "  " & strOrdTB & " } AS ParentCMD APPEND "" & vbCrLf & _ " & vbCrLf & _
     "  ""({SELECT " & strFieldChild & """ & vbCrLf & _ " & vbCrLf & _
     "  ""FROM " & glo.strRSDataGrid & " " & _
     "" & strOrdDG & " } AS ChildCMD " & _
     "RELATE " & glo.strRelationTextBox & _
     "  TO " & glo.strRelationDataGrid & ") " & _
     "AS ChildCMD"
   
  Exit Sub
Message:
  MsgBox Err.Number & " - " & Err.Description
End Sub

'This will get only name of file from a complete
'file name with its directory
Function StripPath(T$) As String
Dim x%, ct%
  StripPath$ = T$
  x% = InStr(T$, "\")
  Do While x%
     ct% = x%
     x% = InStr(ct% + 1, T$, "\")
  Loop
  If ct% > 0 Then StripPath$ = Mid$(T$, ct% + 1)
End Function

