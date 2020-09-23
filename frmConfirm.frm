VERSION 5.00
Begin VB.Form frmConfirm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Successful"
   ClientHeight    =   2595
   ClientLeft      =   3540
   ClientTop       =   5310
   ClientWidth     =   5325
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfirm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   5325
   StartUpPosition =   2  'CenterScreen
   Tag             =   "10000"
   Begin VB.CheckBox chkDontShowAgain 
      Caption         =   "Don't show this confirm next time."
      Height          =   270
      Left            =   120
      MaskColor       =   &H00000000&
      TabIndex        =   1
      Tag             =   "10002"
      Top             =   1560
      Width           =   5055
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   1680
      MaskColor       =   &H00000000&
      TabIndex        =   0
      Tag             =   "10003"
      Top             =   1935
      Width           =   1875
   End
   Begin VB.Label lblLocation 
      BackStyle       =   0  'Transparent
      Height          =   855
      Left            =   1560
      TabIndex        =   3
      Tag             =   "10001"
      Top             =   600
      Width           =   3480
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   930
      Left            =   195
      Picture         =   "frmConfirm.frx":000C
      Stretch         =   -1  'True
      Top             =   210
      Width           =   975
   End
   Begin VB.Label lblConfirm 
      BackStyle       =   0  'Transparent
      Caption         =   "Congratulations! A project has been created in your directory:"
      Height          =   450
      Left            =   1560
      TabIndex        =   2
      Tag             =   "10001"
      Top             =   165
      Width           =   3555
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkDontShowAgain_Click()
    On Error Resume Next
  
    If chkDontShowAgain.Value = vbChecked Then
        SaveSetting APP_CATEGORY, WIZARD_NAME, CONFIRM_KEY, DONTSHOW_CONFIRM
    Else
        SaveSetting APP_CATEGORY, WIZARD_NAME, CONFIRM_KEY, vbNullString
    End If

End Sub

Private Sub cmdOK_Click()
   On Error Resume Next
   If frmWizard.chkRunProject.Value = 1 Then
      frmProcess.RunProject
   End If
   Dim rc As Long
   If mbHelpStarted Then rc = WinHelp(Me.hwnd, HELP_FILE, HELP_QUIT, 0)
   Unload Me
   End
End Sub

Private Sub Form_Load()
    LoadResStrings Me
    lblLocation.Caption = fldr
End Sub

