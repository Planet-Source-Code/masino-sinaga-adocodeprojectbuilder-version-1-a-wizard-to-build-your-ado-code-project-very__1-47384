VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmWizard 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ADO Code Project Builder ver 1.0 (c) Masino Sinaga (masino_sinaga@yahoo.com)"
   ClientHeight    =   5100
   ClientLeft      =   -9960
   ClientTop       =   1815
   ClientWidth     =   7155
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
   Icon            =   "frmWizard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Tag             =   "10"
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Finished!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Index           =   8
      Left            =   -10000
      TabIndex        =   30
      Tag             =   "3000"
      Top             =   0
      Width           =   7155
      Begin VB.CheckBox chkRunProject 
         Caption         =   "&Run project after it has been created."
         Height          =   255
         Left            =   3000
         TabIndex        =   93
         Tag             =   "3002"
         Top             =   2640
         Value           =   1  'Checked
         Width           =   3495
      End
      Begin VB.Label lblProcess 
         Height          =   975
         Left            =   3000
         TabIndex        =   102
         Top             =   1320
         Width           =   3975
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   3075
         Index           =   9
         Left            =   210
         Picture         =   "frmWizard.frx":0442
         Stretch         =   -1  'True
         Top             =   210
         Width           =   2430
      End
      Begin VB.Label lblStep 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ADO Code Maker Wizard is finished collecting information. Press Finish to create the project."
         ForeColor       =   &H80000008&
         Height          =   990
         Index           =   9
         Left            =   3000
         TabIndex        =   31
         Tag             =   "3001"
         Top             =   210
         Width           =   3960
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Step 7"
      Enabled         =   0   'False
      Height          =   4425
      Index           =   7
      Left            =   -10000
      TabIndex        =   16
      Tag             =   "2010"
      Top             =   0
      Width           =   7245
      Begin VB.Frame fraControls 
         Caption         =   "Available Controls:"
         Height          =   2415
         Left            =   240
         TabIndex        =   64
         Tag             =   "115"
         Top             =   1920
         Width           =   6735
         Begin VB.CheckBox chkButton 
            Caption         =   "Cl&ose Button"
            Height          =   255
            Index           =   11
            Left            =   4320
            TabIndex        =   101
            Tag             =   "cmdClose"
            Top             =   1440
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkButton 
            Caption         =   "Data&Grid Button"
            Height          =   255
            Index           =   10
            Left            =   4320
            TabIndex        =   100
            Tag             =   "cmdDataGrid"
            Top             =   1080
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkButton 
            Caption         =   "&Bookmark Button"
            Height          =   255
            Index           =   9
            Left            =   4320
            TabIndex        =   99
            Tag             =   "cmdBookmark"
            Top             =   720
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkButton 
            Caption         =   "&Sort Button"
            Height          =   255
            Index           =   8
            Left            =   4320
            TabIndex        =   98
            Tag             =   "cmdSort"
            Top             =   360
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkButton 
            Caption         =   "&Filter Button"
            Height          =   255
            Index           =   7
            Left            =   2280
            TabIndex        =   96
            Tag             =   "cmdFilter"
            Top             =   1440
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkButton 
            Caption         =   "Fin&d Button"
            Height          =   255
            Index           =   6
            Left            =   2280
            TabIndex        =   95
            Tag             =   "cmdFind"
            Top             =   1080
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkButton 
            Caption         =   "&Refresh Button"
            Height          =   255
            Index           =   5
            Left            =   2280
            TabIndex        =   94
            Tag             =   "cmdRefresh"
            Top             =   720
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkButton 
            Caption         =   "&Delete Button"
            Height          =   255
            Index           =   4
            Left            =   2280
            TabIndex        =   71
            Tag             =   "cmdDelete"
            Top             =   360
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkButton 
            Caption         =   "&Edit Button"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   70
            Tag             =   "cmdEdit"
            Top             =   1440
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkButton 
            Caption         =   "&Cancel Button"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   69
            Tag             =   "cmdCancel"
            Top             =   1080
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkButton 
            Caption         =   "&Update Button"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   68
            Tag             =   "cmdUpdate"
            Top             =   720
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CommandButton cmdClearAll 
            Caption         =   "&Clear All"
            Height          =   312
            Left            =   3240
            TabIndex        =   67
            Tag             =   "116"
            Top             =   1920
            Width           =   1575
         End
         Begin VB.CommandButton cmdSelectControls 
            Caption         =   "&Select All"
            Height          =   312
            Left            =   4965
            TabIndex        =   66
            Tag             =   "117"
            Top             =   1920
            Width           =   1575
         End
         Begin VB.CheckBox chkButton 
            Caption         =   "&Add Button"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   65
            Tag             =   "cmdAdd"
            Top             =   360
            Value           =   1  'Checked
            Width           =   1695
         End
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   1600
         Index           =   7
         Left            =   210
         Picture         =   "frmWizard.frx":8624
         Stretch         =   -1  'True
         Top             =   210
         Width           =   2070
      End
      Begin VB.Label lblStep 
         Caption         =   "Select the desired controls to place on the form."
         Height          =   990
         Index           =   8
         Left            =   2700
         TabIndex        =   17
         Tag             =   "2012"
         Top             =   210
         Width           =   3960
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Step 6"
      Enabled         =   0   'False
      Height          =   4425
      Index           =   6
      Left            =   -10000
      TabIndex        =   14
      Tag             =   "2009"
      Top             =   0
      Width           =   7245
      Begin VB.ListBox lstDataGrid 
         Height          =   1620
         Left            =   4800
         TabIndex        =   61
         Top             =   2280
         Width           =   2055
      End
      Begin VB.ListBox lstTextBox 
         Height          =   1620
         Left            =   2520
         TabIndex        =   60
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   1600
         Index           =   6
         Left            =   210
         Picture         =   "frmWizard.frx":CEFA
         Stretch         =   -1  'True
         Top             =   210
         Width           =   2070
      End
      Begin VB.Label Label2 
         Caption         =   "&DataGrid"
         Height          =   255
         Left            =   4800
         TabIndex        =   63
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label lblTextBox 
         Caption         =   "&TextBox"
         Height          =   255
         Left            =   2520
         TabIndex        =   62
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label lblStep 
         Caption         =   "Select a field from each record source that links the two sources together."
         Height          =   1470
         Index           =   7
         Left            =   2700
         TabIndex        =   15
         Tag             =   "2011"
         Top             =   210
         Width           =   3960
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Step 5"
      Enabled         =   0   'False
      Height          =   4425
      Index           =   5
      Left            =   0
      TabIndex        =   12
      Tag             =   "2008"
      Top             =   0
      Width           =   7245
      Begin VB.ComboBox cboRS 
         Height          =   315
         Index           =   1
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   1560
         Width           =   3975
      End
      Begin VB.ListBox lstSelectedFields 
         Height          =   1425
         Index           =   1
         Left            =   3480
         TabIndex        =   46
         Top             =   2280
         Width           =   2415
      End
      Begin VB.CommandButton cmdMoveAR 
         Caption         =   ">>"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2865
         TabIndex        =   53
         Top             =   2640
         Width           =   400
      End
      Begin VB.CommandButton cmdMove1L 
         Caption         =   "<"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2865
         TabIndex        =   52
         Top             =   3000
         Width           =   400
      End
      Begin VB.CommandButton cmdMoveAL 
         Caption         =   "<<"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2865
         TabIndex        =   51
         Top             =   3360
         Width           =   400
      End
      Begin VB.CommandButton cmdMoveUp 
         Caption         =   "^"
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   6120
         TabIndex        =   50
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton cmdMoveDown 
         Caption         =   "V"
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   6120
         TabIndex        =   49
         Top             =   2640
         Width           =   375
      End
      Begin VB.ListBox lstAvailableFields 
         Height          =   1425
         Index           =   1
         Left            =   240
         TabIndex        =   48
         Top             =   2280
         Width           =   2415
      End
      Begin VB.CommandButton cmdMove1R 
         Caption         =   ">"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2865
         TabIndex        =   47
         Top             =   2280
         Width           =   400
      End
      Begin VB.ComboBox cboColumnSort 
         Height          =   315
         Index           =   1
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   4080
         Width           =   2415
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   1600
         Index           =   5
         Left            =   210
         Picture         =   "frmWizard.frx":117D0
         Stretch         =   -1  'True
         Top             =   210
         Width           =   2070
      End
      Begin VB.Label lblRecordSource 
         Caption         =   "&Record Source:"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   59
         Tag             =   "111"
         Top             =   1300
         Width           =   2775
      End
      Begin VB.Label lblColumnSort 
         Caption         =   "Column To Sort By:"
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   58
         Tag             =   "114"
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label lblAvFields 
         Caption         =   "Available Fields:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   57
         Tag             =   "112"
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label lblSelFields 
         Caption         =   "Selected Fields:"
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   56
         Tag             =   "113"
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label lblStep 
         Caption         =   "Select the record source and fields for the selected record that would be displayed in DataGrid control."
         Height          =   630
         Index           =   6
         Left            =   2520
         TabIndex        =   13
         Tag             =   "2013"
         Top             =   210
         Width           =   3960
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Introduction Screen"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Index           =   0
      Left            =   -10000
      TabIndex        =   6
      Tag             =   "1000"
      Top             =   0
      Width           =   7155
      Begin VB.CheckBox chkShowIntro 
         Caption         =   "Skip this welcome screen next time I use this wizard."
         Height          =   315
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   18
         Tag             =   "1002"
         Top             =   3960
         Width           =   5610
      End
      Begin VB.Label lblStep 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmWizard.frx":160A6
         ForeColor       =   &H80000008&
         Height          =   1000
         Index           =   0
         Left            =   2520
         TabIndex        =   19
         Tag             =   "1001"
         Top             =   240
         Width           =   4260
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   1600
         Index           =   0
         Left            =   210
         Picture         =   "frmWizard.frx":1619C
         Stretch         =   -1  'True
         Top             =   210
         Width           =   2070
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Step 1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Index           =   1
      Left            =   -10000
      TabIndex        =   7
      Tag             =   "2000"
      Top             =   0
      Width           =   7155
      Begin VB.ListBox lstDBFormat 
         Height          =   2010
         ItemData        =   "frmWizard.frx":1AA72
         Left            =   2640
         List            =   "frmWizard.frx":1AA7C
         TabIndex        =   20
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label lblStep 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Select a database format from the list."
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   2640
         TabIndex        =   21
         Tag             =   "2001"
         Top             =   240
         Width           =   3960
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   1600
         Index           =   1
         Left            =   210
         Picture         =   "frmWizard.frx":1AA97
         Stretch         =   -1  'True
         Top             =   210
         Width           =   2070
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Step 2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Index           =   2
      Left            =   -10000
      TabIndex        =   8
      Tag             =   "2002"
      Top             =   0
      Width           =   7155
      Begin VB.Frame fraMDB 
         Caption         =   "Access Database:"
         Height          =   2175
         Left            =   2520
         TabIndex        =   72
         Top             =   2160
         Width           =   4455
         Begin VB.CheckBox chkCopyDBFile 
            Caption         =   "&Copy database file to the project directory"
            Height          =   375
            Left            =   360
            TabIndex        =   97
            Tag             =   "106"
            ToolTipText     =   "Your database file will be copied to the project directory (App.Path)."
            Top             =   1680
            Width           =   3975
         End
         Begin VB.TextBox txtDBFileName 
            Height          =   285
            Left            =   360
            TabIndex        =   75
            Top             =   600
            Width           =   3855
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "Bro&wse..."
            Height          =   312
            Left            =   3120
            TabIndex        =   74
            Top             =   960
            Width           =   1092
         End
         Begin VB.TextBox txtPass 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   360
            PasswordChar    =   "*"
            TabIndex        =   73
            Top             =   1320
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.Label lblDBName 
            Caption         =   "&Database Name:"
            Height          =   255
            Left            =   360
            TabIndex        =   77
            Tag             =   "105"
            Top             =   360
            Width           =   3855
         End
         Begin VB.Label lblPass 
            Caption         =   "&Password:"
            Height          =   255
            Left            =   360
            TabIndex        =   76
            Top             =   1080
            Visible         =   0   'False
            Width           =   1335
         End
      End
      Begin MSComDlg.CommonDialog dlgOpen 
         Left            =   6480
         Top             =   1440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame fraODBC 
         Caption         =   "ODBC Connect:"
         Height          =   2415
         Left            =   2520
         TabIndex        =   84
         Top             =   1920
         Width           =   4455
         Begin VB.ComboBox cboDrivers 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1110
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   82
            Top             =   1590
            Width           =   3015
         End
         Begin VB.TextBox txtServer 
            Enabled         =   0   'False
            Height          =   330
            Left            =   1110
            TabIndex        =   83
            Top             =   1935
            Width           =   3015
         End
         Begin VB.ComboBox cboDSNList 
            Height          =   315
            ItemData        =   "frmWizard.frx":1F36D
            Left            =   1110
            List            =   "frmWizard.frx":1F36F
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   78
            Top             =   240
            Width           =   3000
         End
         Begin VB.TextBox txtDatabase 
            Height          =   300
            Left            =   1110
            TabIndex        =   81
            Top             =   1260
            Width           =   3015
         End
         Begin VB.TextBox txtPWD 
            Height          =   300
            Left            =   1110
            TabIndex        =   80
            Top             =   930
            Width           =   3015
         End
         Begin VB.TextBox txtUID 
            Height          =   300
            Left            =   1110
            TabIndex        =   79
            Top             =   600
            Width           =   3015
         End
         Begin VB.Label lblStep3 
            AutoSize        =   -1  'True
            Caption         =   "&Server:"
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   90
            Top             =   2010
            Width           =   510
         End
         Begin VB.Label lblStep3 
            AutoSize        =   -1  'True
            Caption         =   "Dri&ver:"
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   89
            Top             =   1665
            Width           =   465
         End
         Begin VB.Label lblStep3 
            AutoSize        =   -1  'True
            Caption         =   "Data&base:"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   88
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label lblStep3 
            AutoSize        =   -1  'True
            Caption         =   "&PWD:"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   87
            Top             =   975
            Width           =   435
         End
         Begin VB.Label lblStep3 
            AutoSize        =   -1  'True
            Caption         =   "&UID:"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   86
            Top             =   630
            Width           =   330
         End
         Begin VB.Label lblStep3 
            AutoSize        =   -1  'True
            Caption         =   "&DSN:"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   85
            Top             =   285
            Width           =   390
         End
      End
      Begin VB.Label lblStep 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Click the browse button to select a database file."
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   2
         Left            =   2520
         TabIndex        =   22
         Tag             =   "2003"
         Top             =   240
         Width           =   4200
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   1600
         Index           =   2
         Left            =   210
         Picture         =   "frmWizard.frx":1F371
         Stretch         =   -1  'True
         Top             =   210
         Width           =   2070
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Step 3"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Index           =   3
      Left            =   -10000
      TabIndex        =   9
      Tag             =   "2004"
      Top             =   0
      Width           =   7155
      Begin VB.Frame fraBinding 
         Caption         =   "Binding Type"
         Height          =   1575
         Left            =   4560
         TabIndex        =   28
         Tag             =   "110"
         Top             =   2790
         Width           =   2415
         Begin VB.OptionButton optBindingType 
            Caption         =   "ADO &Code"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   29
            Top             =   360
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.Label lblADO 
            BackStyle       =   0  'Transparent
            Caption         =   "Only ADO Code is available now. The others binding type is not available."
            Height          =   735
            Left            =   250
            TabIndex        =   91
            Tag             =   "109"
            Top             =   720
            Width           =   2055
         End
      End
      Begin VB.ListBox lstFormLayout 
         Height          =   1230
         ItemData        =   "frmWizard.frx":23C47
         Left            =   2160
         List            =   "frmWizard.frx":23C51
         TabIndex        =   24
         Tag             =   "118"
         Top             =   3120
         Width           =   2295
      End
      Begin VB.TextBox txtFormName 
         Height          =   350
         Left            =   2400
         TabIndex        =   23
         Top             =   1890
         Width           =   4575
      End
      Begin VB.Label lblDescLayout 
         Height          =   1215
         Left            =   120
         TabIndex        =   92
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label lblFormLayout 
         Caption         =   "&Form Layout"
         Height          =   255
         Left            =   2160
         TabIndex        =   27
         Tag             =   "108"
         Top             =   2790
         Width           =   2295
      End
      Begin VB.Label lblStep 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Select the desired form type and a data binding type to use to access the data."
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   3
         Left            =   2400
         TabIndex        =   26
         Tag             =   "2005"
         Top             =   120
         Width           =   4080
      End
      Begin VB.Label lblFormName 
         Caption         =   "What name do you want for the form?"
         Height          =   255
         Left            =   2400
         TabIndex        =   25
         Tag             =   "107"
         Top             =   1590
         Width           =   3615
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   1600
         Index           =   3
         Left            =   210
         Picture         =   "frmWizard.frx":23C77
         Stretch         =   -1  'True
         Top             =   210
         Width           =   2070
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Step 4"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Index           =   4
      Left            =   -10000
      TabIndex        =   10
      Tag             =   "2006"
      Top             =   0
      Width           =   7155
      Begin VB.ComboBox cboColumnSort 
         Height          =   315
         Index           =   0
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   4080
         Width           =   2415
      End
      Begin VB.CommandButton cmdMove1R 
         Caption         =   ">"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   2860
         TabIndex        =   41
         Top             =   2280
         Width           =   400
      End
      Begin VB.ListBox lstAvailableFields 
         Height          =   1425
         Index           =   0
         Left            =   240
         TabIndex        =   40
         Top             =   2280
         Width           =   2415
      End
      Begin VB.CommandButton cmdMoveDown 
         Caption         =   "V"
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   6120
         TabIndex        =   39
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton cmdMoveUp 
         Caption         =   "^"
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   6120
         TabIndex        =   38
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton cmdMoveAL 
         Caption         =   "<<"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   2860
         TabIndex        =   37
         Top             =   3360
         Width           =   400
      End
      Begin VB.CommandButton cmdMove1L 
         Caption         =   "<"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   2860
         TabIndex        =   36
         Top             =   3000
         Width           =   400
      End
      Begin VB.CommandButton cmdMoveAR 
         Caption         =   ">>"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   2860
         TabIndex        =   35
         Top             =   2640
         Width           =   400
      End
      Begin VB.ListBox lstSelectedFields 
         Height          =   1425
         Index           =   0
         Left            =   3480
         TabIndex        =   34
         Top             =   2280
         Width           =   2415
      End
      Begin VB.ComboBox cboRS 
         Height          =   315
         Index           =   0
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1560
         Width           =   3975
      End
      Begin VB.Label lblSelFields 
         Caption         =   "Selected Fields:"
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   44
         Tag             =   "113"
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label lblAvFields 
         Caption         =   "Available Fields:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   43
         Tag             =   "112"
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label lblColumnSort 
         Caption         =   "Column To Sort By:"
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   42
         Tag             =   "114"
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label lblRecordSource 
         Caption         =   "&Record Source:"
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   32
         Tag             =   "111"
         Top             =   1300
         Width           =   2775
      End
      Begin VB.Label lblStep 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Select the record source and fields for the selected record that would be displayed in textbox control."
         ForeColor       =   &H80000008&
         Height          =   630
         Index           =   4
         Left            =   2520
         TabIndex        =   11
         Tag             =   "2007"
         Top             =   210
         Width           =   4140
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   1600
         Index           =   4
         Left            =   210
         Picture         =   "frmWizard.frx":2854D
         Stretch         =   -1  'True
         Top             =   210
         Width           =   2070
      End
   End
   Begin VB.PictureBox picNav 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   7155
      TabIndex        =   0
      Top             =   4530
      Width           =   7155
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Finish"
         Height          =   312
         Index           =   4
         Left            =   5910
         MaskColor       =   &H00000000&
         TabIndex        =   5
         Tag             =   "104"
         ToolTipText     =   "Build your project now!"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Next >"
         Default         =   -1  'True
         Height          =   312
         Index           =   3
         Left            =   4545
         MaskColor       =   &H00000000&
         TabIndex        =   4
         Tag             =   "103"
         ToolTipText     =   "Move to the next step"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "< &Back"
         Height          =   312
         Index           =   2
         Left            =   3435
         MaskColor       =   &H00000000&
         TabIndex        =   3
         Tag             =   "102"
         ToolTipText     =   "Move back one step"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   312
         Index           =   1
         Left            =   2250
         MaskColor       =   &H00000000&
         TabIndex        =   2
         Tag             =   "101"
         ToolTipText     =   "Close the wizard without building anything"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "Help"
         Height          =   312
         Index           =   0
         Left            =   108
         MaskColor       =   &H00000000&
         TabIndex        =   1
         Tag             =   "100"
         ToolTipText     =   "Get help on using the wizard"
         Top             =   120
         Width           =   1092
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   108
         X2              =   7012
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   108
         X2              =   7012
         Y1              =   24
         Y2              =   24
      End
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----LOCALEID
Private Declare Function GetUserDefaultLangID _
        Lib "kernel32" () As Integer
'------------

'----- ODBC Connect -----
Private Declare Function SQLDataSources Lib "ODBC32.DLL" (ByVal henv&, ByVal fDirection%, ByVal szDSN$, ByVal cbDSNMax%, pcbDSN%, ByVal szDescription$, ByVal cbDescriptionMax%, pcbDescription%) As Integer
Private Declare Function SQLAllocEnv% Lib "ODBC32.DLL" (env&)
Const SQL_SUCCESS As Long = 0
Const SQL_FETCH_NEXT As Long = 1
'----- ODBC Connect -----

Const NUM_STEPS = 9 'number of all steps

'Error Base in .RES file...
Const RES_ERROR_MSG = 30000

Const BTN_HELP = 0
Const BTN_CANCEL = 1
Const BTN_BACK = 2
Const BTN_NEXT = 3
Const BTN_FINISH = 4

Const STEP_INTRO = 0
Const STEP_1 = 1
Const STEP_2 = 2
Const STEP_3 = 3
Const STEP_4 = 4
Const STEP_5 = 5
Const STEP_6 = 6
Const STEP_7 = 7
Const STEP_FINISH = 8

Const DIR_NONE = 0
Const DIR_BACK = 1
Const DIR_NEXT = 2

Const FRM_TITLE = "ADO Code Project Builder ver. 1"
Const INTRO_KEY = "IntroductionScreen"
Const SHOW_INTRO = "ShowIntro"
Const TOPIC_TEXT = "<TOPIC_TEXT>"

'module level vars
Dim mnCurStep       As Integer
Dim mbStepSeven     As Boolean
Public VBInst       As VBIDE.VBE
Dim mbFinishOK      As Boolean

'When a record source was selected in cboRS, list all
'fields in that record source to the lstAvailableFields
Private Sub cboRS_Click(Index As Integer)
Select Case Index
       Case 0
            If cboRS(0).Text <> "" Then
               DoEvents
               lstSelectedFields(0).Clear
               Call AllFieldsToListBox(cboRS(0).Text, lstAvailableFields(0))
               DoEvents
               lstAvailableFields(0).Text = lstAvailableFields(0).List(0)
               Call AllFieldsToComboBox(cboRS(0).Text, cboColumnSort(0))
               DoEvents
            End If
            If lstAvailableFields(0).ListCount > 0 Then
               cmdMove1R(0).Enabled = True
               cmdMoveAR(0).Enabled = True
            Else
               cmdMove1R(0).Enabled = False
               cmdMoveAR(0).Enabled = False
            End If
            cmdMove1L(0).Enabled = False
            cmdMoveAL(0).Enabled = False
            CheckMoveButton
       Case 1
            If cboRS(1).Text <> "" Then
               DoEvents
               lstSelectedFields(1).Clear
               Call AllFieldsToListBox(cboRS(1).Text, lstAvailableFields(1))
               DoEvents
               lstAvailableFields(1).Text = lstAvailableFields(1).List(0)
               Call AllFieldsToComboBox(cboRS(1).Text, cboColumnSort(1))
               DoEvents
            End If
            If lstAvailableFields(1).ListCount > 0 Then
               cmdMove1R(1).Enabled = True
               cmdMoveAR(1).Enabled = True
            Else
               cmdMove1R(1).Enabled = False
               cmdMoveAR(1).Enabled = False
            End If
            cmdMove1L(1).Enabled = False
            cmdMoveAL(1).Enabled = False
            CheckMoveButton
End Select
End Sub

Private Sub cboRS_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
       Case 0
            If cmdMove1R(0).Enabled = True Then
               cmdMove1R(0).SetFocus
            End If
       Case 1
            If cmdMove1R(1).Enabled = True Then
               cmdMove1R(1).SetFocus
            End If
End Select
End Sub

Private Sub chkButton_Click(Index As Integer)
Select Case Index
       Case 0
            If chkButton(0).Value = 1 Then
               chkButton(1).Value = 1
               chkButton(2).Value = 1
            End If
       Case 1
            If chkButton(1).Value = 1 Then
               chkButton(2).Value = 1
            End If
       Case 2
            If chkButton(2).Value = 1 Then
               chkButton(1).Value = 1
               chkButton(5).Value = 1
            End If
       Case 4
            If chkButton(4).Value = 1 Then
               chkButton(1).Value = 1
               chkButton(2).Value = 1
            End If
End Select
End Sub

'The welcome screen; Display it or not next time you
'use this wizard
Private Sub chkShowIntro_Click()
    If chkShowIntro.Value Then
        SaveSetting APP_CATEGORY, WIZARD_NAME, INTRO_KEY, SHOW_INTRO
    Else
        SaveSetting APP_CATEGORY, WIZARD_NAME, INTRO_KEY, vbNullString
    End If
End Sub

'Browse Access database file
Private Sub cmdBrowse_Click()
On Error GoTo Cancel
   With dlgOpen
      .CancelError = True
      .Filter = "*.mdb (Microsoft Access)|*.mdb"
      .ShowOpen
      txtDBFileName.Text = .FileName
      cmdNav(BTN_NEXT).SetFocus
      cmdNav(BTN_NEXT).Default = True
   End With
   Exit Sub
Cancel:
   Exit Sub
End Sub

'This for button
Private Sub cmdClearAll_Click()
  SetCheckBox 0
End Sub

'This will move an item in lstSelectedFields to
'lstAvailableFields .
Private Sub cmdMove1L_Click(Index As Integer)
Select Case Index
       Case 0
            lstAvailableFields(0).AddItem _
                 lstSelectedFields(0).Text
            lstSelectedFields(0).RemoveItem _
                 lstSelectedFields(0).ListIndex
            If lstSelectedFields(0).ListCount > 0 Then
               lstSelectedFields(0).Selected(0) = True
               cmdMove1L(0).Enabled = True
               cmdMoveAL(0).Enabled = True
            Else
               cmdMove1L(0).Enabled = False
               cmdMove1R(0).Enabled = True
               cmdMoveAL(0).Enabled = False
               cmdMoveAR(0).Enabled = True
            End If
            If lstAvailableFields(0).ListCount > 0 Then
               lstAvailableFields(0).Selected(0) = True
               cmdMove1R(0).Enabled = True
               cmdMoveAR(0).Enabled = True
            Else
               cmdMove1R(0).Enabled = False
               cmdMove1L(0).Enabled = True
               cmdMoveAR(0).Enabled = False
               cmdMoveAL(0).Enabled = True
            End If
       Case 1
            lstAvailableFields(1).AddItem _
                 lstSelectedFields(1).Text
            lstSelectedFields(1).RemoveItem _
                 lstSelectedFields(1).ListIndex
            If lstSelectedFields(1).ListCount > 0 Then
               lstSelectedFields(1).Selected(0) = True
               cmdMove1L(1).Enabled = True
               cmdMoveAL(1).Enabled = True
            Else
               cmdMove1L(1).Enabled = False
               cmdMove1R(1).Enabled = True
               cmdMoveAL(1).Enabled = False
               cmdMoveAR(1).Enabled = True
            End If
            If lstAvailableFields(1).ListCount > 0 Then
               lstAvailableFields(1).Selected(0) = True
               cmdMove1R(1).Enabled = True
               cmdMoveAR(1).Enabled = True
            Else
               cmdMove1R(1).Enabled = False
               cmdMove1L(1).Enabled = True
               cmdMoveAR(1).Enabled = False
               cmdMoveAL(1).Enabled = True
            End If
End Select
CheckMoveButton
End Sub

'This will move an item in lstAvailableFields to
'lstSelectedFields.
Private Sub cmdMove1R_Click(Index As Integer)
Select Case Index
       Case 0
            lstSelectedFields(0).AddItem _
                 lstAvailableFields(0).Text
            lstAvailableFields(0).RemoveItem _
                 lstAvailableFields(0).ListIndex
            If lstAvailableFields(0).ListCount > 0 Then
               lstAvailableFields(0).Selected(0) = True
               cmdMove1R(0).Enabled = True
               cmdMoveAR(0).Enabled = True
            Else
               cmdMove1R(0).Enabled = False
               cmdMove1L(0).Enabled = True
               cmdMoveAR(0).Enabled = False
               cmdMoveAL(0).Enabled = True
            End If
            If lstSelectedFields(0).ListCount > 0 Then
               lstSelectedFields(0).Selected(0) = True
               cmdMove1L(0).Enabled = True
               cmdMoveAL(0).Enabled = True
            Else
               cmdMove1L(0).Enabled = False
               cmdMove1R(0).Enabled = True
               cmdMoveAL(0).Enabled = False
               cmdMoveAR(0).Enabled = True
            End If
       
       Case 1
            lstSelectedFields(1).AddItem _
                 lstAvailableFields(1).Text
            lstAvailableFields(1).RemoveItem _
                 lstAvailableFields(1).ListIndex
            If lstAvailableFields(1).ListCount > 0 Then
               lstAvailableFields(1).Selected(0) = True
               cmdMove1R(1).Enabled = True
               cmdMoveAR(1).Enabled = True
            Else
               cmdMove1R(1).Enabled = False
               cmdMove1L(1).Enabled = True
               cmdMoveAR(1).Enabled = False
               cmdMoveAL(1).Enabled = True
            End If
            If lstSelectedFields(1).ListCount > 0 Then
               lstSelectedFields(1).Selected(0) = True
               cmdMove1L(1).Enabled = True
               cmdMoveAL(1).Enabled = True
            Else
               cmdMove1L(1).Enabled = False
               cmdMove1R(1).Enabled = True
               cmdMoveAL(1).Enabled = False
               cmdMoveAR(1).Enabled = True
            End If
            If lstAvailableFields(1).ListCount < 1 Then
               cboColumnSort(1).SetFocus
            End If
End Select
CheckMoveButton
End Sub

'This will move all items in lstSelectedFields to
'lstAvailableFields .
Private Sub cmdMoveAL_Click(Index As Integer)
Dim i As Byte
Select Case Index
       Case 0
            For i = 0 To lstSelectedFields(0).ListCount - 1
                lstAvailableFields(0).AddItem _
                   lstSelectedFields(0).List(i)
            Next i
            lstSelectedFields(0).Clear
            lstAvailableFields(0).Selected(0) = True
            cmdMove1R(0).Enabled = True
            cmdMoveAR(0).Enabled = True
            cmdMove1L(0).Enabled = False
            cmdMoveAL(0).Enabled = False
       Case 1
            For i = 0 To lstSelectedFields(1).ListCount - 1
                lstAvailableFields(1).AddItem _
                   lstSelectedFields(1).List(i)
            Next i
            lstSelectedFields(1).Clear
            lstAvailableFields(1).Selected(0) = True
            cmdMove1R(1).Enabled = True
            cmdMoveAR(1).Enabled = True
            cmdMove1L(1).Enabled = False
            cmdMoveAL(1).Enabled = False
End Select
CheckMoveButton
End Sub

'This will move all items in lstAvailableFields to
'lstSelectedFields.
Private Sub cmdMoveAR_Click(Index As Integer)
Dim i As Byte
Select Case Index
       Case 0
            For i = 0 To lstAvailableFields(0).ListCount - 1
                lstSelectedFields(0).AddItem _
                   lstAvailableFields(0).List(i)
            Next i
            lstAvailableFields(0).Clear
            lstSelectedFields(0).Selected(0) = True
            cmdMove1R(0).Enabled = False
            cmdMoveAR(0).Enabled = False
            cmdMove1L(0).Enabled = True
            cmdMoveAL(0).Enabled = True
       Case 1
            For i = 0 To lstAvailableFields(1).ListCount - 1
                lstSelectedFields(1).AddItem _
                   lstAvailableFields(1).List(i)
            Next i
            lstAvailableFields(1).Clear
            lstSelectedFields(1).Selected(0) = True
            cmdMove1R(1).Enabled = False
            cmdMoveAR(1).Enabled = False
            cmdMove1L(1).Enabled = True
            cmdMoveAL(1).Enabled = True
End Select
CheckMoveButton
End Sub

'The item in lstSelectedFields can be moved to another
'lower position, and will adjust the other position.
Private Sub cmdMoveDown_Click(Index As Integer)
Select Case Index
       Case 0
            Call MoveDown(lstSelectedFields(0))
       Case 1
            Call MoveDown(lstSelectedFields(1))
End Select
End Sub

'The item in lstSelectedFields can be moved to another
'higher position, and will adjust the other position.
Private Sub cmdMoveUp_Click(Index As Integer)
Select Case Index
       Case 0
            Call MoveUp(lstSelectedFields(0))
       Case 1
            Call MoveUp(lstSelectedFields(1))
End Select
End Sub

'Check navigation process.
Private Sub cmdNav_Click(Index As Integer)
    Dim nAltStep As Integer
    Dim lHelpTopic As Long
    Dim rc As Long
    
    Select Case Index
        Case BTN_HELP
            mbHelpStarted = True
            lHelpTopic = mnCurStep
            rc = WinHelp(Me.hwnd, HELP_FILE, HELP_CONTEXT, lHelpTopic)
        Case BTN_CANCEL
            Unload Me
          
        Case BTN_BACK
            'place special cases here to jump
            'to alternate steps
            nAltStep = mnCurStep - 1
            SetStep nAltStep, DIR_BACK
          
        Case BTN_NEXT
            'place special cases here to jump
            'to alternate steps
            nAltStep = mnCurStep + 1
            SetStep nAltStep, DIR_NEXT
          
        Case BTN_FINISH
            'wizard creation code goes here
            'Then, generate the code...!
            DoEvents
            lblStep(9).Caption = LoadResString(133)
            DoEvents
            chkRunProject.Enabled = False
            Dim k As Byte
            For k = 0 To cmdNav.UBound
               cmdNav(k).Enabled = False
            Next k
            DoEvents
            With frmProcess
              Screen.MousePointer = vbHourglass
              .GenerateADOCodeToForm
              .GenerateADOCodeToModule
              .GenerateADOCodeToProject
              If chkButton(6).Value = 1 Then _
                 GenerateFindCode
              If chkButton(7).Value = 1 Then _
                 GenerateFilterCode
              If chkButton(8).Value = 1 Then _
                 GenerateSortCode
              If chkButton(9).Value = 1 Then _
                 GenerateBookmarkCode
              Screen.MousePointer = vbDefault
            End With
                        
            DoEvents
            lblProcess.Caption = "Finished!"
            DoEvents
                        
            If GetSetting(APP_CATEGORY, WIZARD_NAME, CONFIRM_KEY, vbNullString) = vbNullString Then
                frmConfirm.Show vbModal
            End If
                        
            If chkRunProject.Value = 1 Then
               frmProcess.RunProject
            End If
                        
            Unload Me
            CloseAllForms
    End Select
End Sub

'This is for displaying all button in the result-form
Private Sub cmdSelectControls_Click()
  SetCheckBox 1
End Sub

Private Sub SetCheckBox(intVal As Byte)
  Dim i As Byte
  For i = 0 To chkButton.UBound
     chkButton(i).Value = intVal
  Next i
End Sub

'Display Help (if any) if user press F1 on keyboard
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        cmdNav_Click BTN_HELP
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    'init all vars
    mbFinishOK = False
    
    For i = 0 To NUM_STEPS - 1
      fraStep(i).Left = -10000
    Next
    
    'Load All string info for Form
    LoadResStrings Me
    
    'Determine 1st Step:
    If GetSetting(APP_CATEGORY, WIZARD_NAME, INTRO_KEY, vbNullString) = SHOW_INTRO Then
        chkShowIntro.Value = vbChecked
        SetStep 1, DIR_NEXT
    Else
        SetStep 0, DIR_NONE
    End If
        
    InitLanguage
End Sub

Private Sub SetStep(nStep As Integer, nDirection As Integer)
Dim i As Byte
    Select Case nStep
        Case STEP_INTRO
      
        Case STEP_1
             'Default value is Access database
             lstDBFormat.Text = lstDBFormat.List(0)
                          
        Case STEP_2
             'If Access database Access was selected
             If lstDBFormat.Text = lstDBFormat.List(0) Then
               fraODBC.Visible = False
               fraODBC.Move -10000
               fraMDB.Visible = True
               fraMDB.Move 2520, 1920
             Else 'Remote (ODBC) was selected
               fraMDB.Visible = False
               fraMDB.Move -10000
               fraODBC.Visible = True
               fraODBC.Move 2520, 1920
               GetDSNsAndDrivers
               lblStep(2).Caption = LoadResString(134)
             End If
             
        Case STEP_3
             
             'If Access was selected
             If fraMDB.Left <> -10000 Then
               If txtDBFileName.Text = "" Then
                  Call IncompleteData(1)
                  'MsgBox "Please enter or browse a .mdb file!", _
                         vbExclamation, "Database File"
                  txtDBFileName.SetFocus
                  Exit Sub
               End If
               'Check database filename
               If Dir(txtDBFileName.Text) = "" Then
                  Call IncompleteData(2)
                  'MsgBox "Invalid Access database file. Try again!", _
                         vbExclamation, "Invalid File"
                  txtDBFileName.SetFocus
                  SendKeys "{Home}+{End}"
                  Exit Sub
               End If
                              
               'Open Access database
               Set cnn = New ADODB.Connection
               cnn.CursorLocation = adUseClient
               On Error GoTo STEP_3_MDB_ERR
               cnn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
                   "Data Source=" & txtDBFileName.Text & ";Jet OLEDB:" & _
                   "Database Password=" & txtPass.Text & ";"
               lblPass.Visible = False
               txtPass.Visible = False
               
             Else 'Remote (ODBC) was selected
               'ODBC Connect:
               'special case... assign the label from code
               lblStep(2).Caption = LoadResString(134)
               Screen.MousePointer = vbHourglass
               Set cnn = New ADODB.Connection
               cnn.CursorLocation = adUseClient
               On Error GoTo STEP_3_ODBC_ERR
               If cboDSNList.Text <> cboDSNList.List(0) Then
                  cnn.Open "PROVIDER=MSDataShape;" & _
                           "Data PROVIDER=MSDASQL;" & _
                           "dsn=" & cboDSNList.Text & ";" & _
                           "uid=" & txtUID.Text & ";" & _
                           "pwd=" & txtPWD.Text & ";"
                  lblPass.Visible = False
                  txtPass.Visible = False
               Else
               '---------
                  Screen.MousePointer = vbHourglass
                  'Check again, whether MySQL or SQL Server ?
                  If InStr(1, UCase(cboDrivers.Text), "MYSQL") > 0 Then
                     cnn.Open "PROVIDER=MSDataShape;" & vbCrLf & _
                     "Driver={" & cboDrivers.Text & "};" & vbCrLf & _
                     "Server=" & txtServer.Text & ";" & vbCrLf & _
                     "Port=3306;" & vbCrLf & _
                     "Option=147458;" & vbCrLf & _
                     "Stmt=;" & vbCrLf & _
                     "Database=" & txtDatabase.Text & ";" & vbCrLf & _
                     "User=" & txtUID.Text & ";" & vbCrLf & _
                     "Password=" & txtPWD.Text & ""
                  Else  '(e.g. SQL Server)
                     cnn.Open "PROVIDER=MSDataShape;" & vbCrLf & _
                     "Driver={" & cboDrivers.Text & "};" & vbCrLf & _
                     "Server=" & txtServer.Text & ";" & vbCrLf & _
                     "Database=" & txtDatabase.Text & ";" & vbCrLf & _
                     "User=" & txtUID.Text & ";" & vbCrLf & _
                     "Password=" & txtPWD.Text & ""
                  End If
               '---------
               End If
             End If
             
             'Open and list table name to listbox and combobox
             Set rs = cnn.OpenSchema(adSchemaTables)
             cboRS(0).Clear
             cboRS(1).Clear
             cboColumnSort(0).Clear
             cboColumnSort(1).Clear
             lstAvailableFields(0).Clear
             lstAvailableFields(1).Clear
             lstSelectedFields(0).Clear
             lstSelectedFields(1).Clear
             cmdMove1R(0).Enabled = False
             cmdMove1R(1).Enabled = False
             cmdMove1L(0).Enabled = False
             cmdMove1L(1).Enabled = False
             cmdMoveAR(0).Enabled = False
             cmdMoveAR(1).Enabled = False
             cmdMoveAL(0).Enabled = False
             cmdMoveAL(1).Enabled = False
             'List all table name except system table,
             'and some invalid tables name
             While rs.EOF <> True
              If rs.Fields("TABLE_TYPE").Value = "TABLE" Then
               If UCase(Left(Trim(rs.Fields("TABLE_NAME").Value), 3)) <> "SYS" And _
                 UCase(Left(Trim(rs.Fields("TABLE_NAME").Value), 4)) <> "MSYS" And _
                 UCase(Left(Trim(rs.Fields("TABLE_NAME").Value), 4)) <> "FROM" And _
                 UCase(Left(Trim(rs.Fields("TABLE_NAME").Value), 5)) <> "WHERE" And _
                 UCase(Left(Trim(rs.Fields("TABLE_NAME").Value), 6)) <> "SELECT" Then
                 cboRS(0).AddItem rs.Fields("TABLE_NAME").Value
                 cboRS(1).AddItem rs.Fields("TABLE_NAME").Value
               End If
              End If
              rs.MoveNext
             Wend
             lstFormLayout.Text = lstFormLayout.List(0)
             Screen.MousePointer = vbDefault
 
        Case STEP_4
                          
             'Check form name, it can't be an empty string
             If txtFormName.Text = "" Then
                'MsgBox "Please enter the name for this form!", _
                       vbExclamation, "Form Name"
                Call IncompleteData(3)
                txtFormName.SetFocus
                Exit Sub
             End If
             'First character must be not a numeric!
             If IsNumeric(Left(txtFormName.Text, 1)) = True Then
                'Call IncompleteData(30000)
                'MsgBox LoadResString(CInt(Me.Tag))
                Call IncompleteData(4)
                'MsgBox "Not a legal object name. Please correct it!", _
                       vbExclamation, "Invalid Form Name"
                txtFormName.SetFocus
                SendKeys "{Home}+{Right}"
                Exit Sub
             End If
            
        Case STEP_5
                          
             'Check record source
             If cboRS(0).Text = "" Then
                Call IncompleteData(5)
                'MsgBox "Please select a record source before you can continue!", _
                       vbExclamation, "Record Source"
                cboRS(0).SetFocus
                Exit Sub
             End If
             'Check selected fields, at least one field
             If lstSelectedFields(0).ListCount = 0 Then
                Call IncompleteData(6)
                'MsgBox "You must select at least one field before to continue!", _
                       vbExclamation, "Selected Fields"
                If cmdMove1R(0).Enabled = True Then
                   cmdMove1R(0).SetFocus
                End If
                Exit Sub
             End If
                                       
             'Add fields to first listbox (left)
             lstTextBox.Clear
             For i = 0 To lstSelectedFields(0).ListCount - 1
                lstTextBox.AddItem _
                   lstSelectedFields(0).List(i)
             Next i
             
             Dim Mundur As Boolean
             Mundur = False
             
             'SKIP Next step is ADO Code Complete
             If lstFormLayout.Text = lstFormLayout.List(0) Then
                Dim FoundTextbox As Boolean
                Dim FoundDataGrid As Boolean
                                         
                cboRS(1).Text = cboRS(0).Text
                                         
                For i = 0 To lstSelectedFields(0).ListCount - 1
                  lstSelectedFields(1).AddItem lstSelectedFields(0).List(i)
                Next i
                
                'Add fields to second listbox (right)
                lstDataGrid.Clear
                For i = 0 To lstSelectedFields(1).ListCount - 1
                   lstDataGrid.AddItem _
                      lstSelectedFields(1).List(i)
                Next i
                 
                
                FoundTextbox = True
                FoundDataGrid = True
                
    
                If nStep > mnCurStep Then
                  nStep = nStep + 1
                  
                  SetCaption nStep + 1
                  SetNavBtns nStep + 1
             
                  fraStep(nStep).Left = -10000
                  mnCurStep = mnCurStep + 1
                 
                  Mundur = False
                
                Else

                  Mundur = True
                  nStep = nStep - 1
                  fraStep(nStep).Left = -10000
                  mnCurStep = mnCurStep - 2
                  SetStep 6, DIR_NONE
                  GoTo GoAhead
                  
                  Exit Sub
                End If
                
                SetStep 6, DIR_NONE
                                
                GoTo GoAhead
                Exit Sub
             End If
             
                                                  
        Case STEP_6
                          
             'Check record source
             If cboRS(1).Text = "" Then
                Call IncompleteData(5)
                'MsgBox "Please select a record source before you can continue!", _
                       vbExclamation, "Record Source"
                cboRS(1).SetFocus
                Exit Sub
             End If
                          
             'At least, one field must be selected
             If lstSelectedFields(1).ListCount = 0 Then
                Call IncompleteData(6)
                'MsgBox "You must select at least one field before to continue!", _
                       vbExclamation, "Selected Fields"
                cmdMove1R(1).SetFocus
                Exit Sub
             End If
             
             'Add fields to second listbox (right)
             lstDataGrid.Clear
             For i = 0 To lstSelectedFields(1).ListCount - 1
                lstDataGrid.AddItem _
                   lstSelectedFields(1).List(i)
             Next i
                                       
        Case STEP_7
             mbFinishOK = False
             mbStepSeven = True
                          
             'Dim FoundTextbox As Boolean
             'Dim FoundDataGrid As Boolean
             FoundTextbox = False
             FoundDataGrid = False
                                       
             For i = 0 To lstTextBox.ListCount - 1
                If lstTextBox.Selected(i) = True Then
                   FoundTextbox = True
                   Exit For
                Else
                End If
             Next i
             
             'SKIP Next step is ADO Code Complete
             'If lstFormLayout.Text <> lstFormLayout.List(0) Then

             'This for relation between textbox and datagrid
             If FoundTextbox = False Then
                   Call IncompleteData(7)
                   'MsgBox "You must select a primary link before you can continue!", _
                          vbExclamation, "Primary Link-1"
'                   lstTextBox.SetFocus
                   Exit Sub
             End If
             For i = 0 To lstDataGrid.ListCount - 1
                If lstDataGrid.Selected(i) = True Then
                   FoundDataGrid = True
                   Exit For
                Else
                End If
             Next i
             If FoundDataGrid = False Then
                   Call IncompleteData(8)
                   'MsgBox "You must select a primary link before you can continue!", _
                          vbExclamation, "Primary Link-2"
                   lstDataGrid.SetFocus
                   Exit Sub
             End If
             
             'End If
        Case STEP_FINISH
             'this is special case!
             'Me.Caption = "ADO Code Project Builder ver. 1"
             'Me.Caption = Me.Caption & " - Finished!"
             
             mbFinishOK = True
                    
             'Get all information that was needed
             'and have been collected before!
             With glo
               .strDBFileName = txtDBFileName.Text
               .strDBPassword = txtPass.Text
               .strDSNName = cboDSNList.Text
               .strDSNUserID = txtUID.Text
               .strDSNPassword = txtPWD.Text
               .strDSNDatabase = txtDatabase.Text
               .strDSNDriver = cboDrivers.Text
               .strDSNServer = txtServer.Text
               .strFormName = txtFormName.Text
               .strFormLayout = lstFormLayout.ListIndex
               .strRSTextBox = cboRS(0).Text
               .strRSDataGrid = cboRS(1).Text
               .intNumOfFields = lstSelectedFields(0).ListCount
               .strOrderTextBox = cboColumnSort(0).Text
               .strOrderDataGrid = cboColumnSort(1).Text
               .strRelationTextBox = lstTextBox.Text
               .strRelationDataGrid = lstDataGrid.Text
             End With
              
    End Select
    
GoAhead:

    'move to new step
    fraStep(mnCurStep).Enabled = False
    fraStep(nStep).Left = 0
    If nStep <> mnCurStep Then
        fraStep(mnCurStep).Left = -10000
    End If
    fraStep(nStep).Enabled = True
  
    SetCaption nStep
    SetNavBtns nStep
    
    'Special case...!
    'If Mundur = True Then
      'Me.Caption = "ADO Code Project Builder ver. 1"
      'Me.Caption = Me.Caption & " - Master Record Source"
    'End If
    
    If fraMDB.Left = 0 And txtDBFileName.Text = "" Then
       cmdBrowse.SetFocus
    End If
    If fraStep(2).Left = 0 And fraMDB.Visible = True Then
       txtDBFileName.SetFocus
    End If
    If fraStep(2).Left = 0 And fraODBC.Visible = True Then
       cboDSNList.SetFocus
    End If
    If fraStep(3).Left = 0 And txtFormName.Text = "" Then
       txtFormName.SetFocus
    End If
    If fraStep(4).Left = 0 And cboRS(0).Text = "" Then
       cboRS(0).SetFocus
    End If
    If fraStep(5).Left = 0 And cboRS(1).Text = "" Then
       cboRS(1).SetFocus
    End If
    
    Exit Sub
    
STEP_3_MDB_ERR:  'This is error handling for Access database
    If cnn.State <> 1 Then
       Call IncompleteData(9)
       'MsgBox "This .mdb file was password protected." & vbCrLf & _
              "Please enter a valid password!", _
              vbExclamation, "Password Protected"
       lblPass.Visible = True
       txtPass.Visible = True
       txtPass.SetFocus
       SendKeys "{Home}+{End}"
       Exit Sub
    Else
       txtDBFileName.Text = dlgOpen.FileName
    End If
    Exit Sub
STEP_3_ODBC_ERR: 'This is error handling for Remote (ODBC)
    Screen.MousePointer = vbDefault
    If cnn.State <> 1 Then
       lblPass.Visible = True
       txtPass.Visible = True
       MsgBox Err.Number & " - " & Err.Description, _
              vbExclamation, "Error ODBC Connect"
       cboDSNList.SetFocus
       Exit Sub
    Else
    End If
End Sub

Private Sub SetNavBtns(nStep As Integer)
    mnCurStep = nStep
    If mnCurStep = 0 Then
        cmdNav(BTN_BACK).Enabled = False
        cmdNav(BTN_NEXT).Enabled = True
    ElseIf mnCurStep = NUM_STEPS - 1 Then
        cmdNav(BTN_NEXT).Enabled = False
        cmdNav(BTN_BACK).Enabled = True
    Else
        cmdNav(BTN_BACK).Enabled = True
        cmdNav(BTN_NEXT).Enabled = True
    End If
    If mbFinishOK Then
        cmdNav(BTN_FINISH).Enabled = True
        cmdNav(BTN_FINISH).Default = True
    Else
        cmdNav(BTN_FINISH).Enabled = False
    End If
End Sub

Private Sub SetCaption(nStep As Integer)
    On Error Resume Next
    Me.Caption = FRM_TITLE & " - " & LoadResString(fraStep(nStep).Tag)
End Sub

'=========================================================
'this sub displays an error message when the user has
'not entered enough data to continue
'=========================================================
Sub IncompleteData(nIndex As Integer)
    On Error Resume Next
    Dim sTmp As String
      
    'get the base error message
    sTmp = LoadResString(RES_ERROR_MSG)
    'get the specific message
    sTmp = sTmp & vbCrLf & LoadResString(RES_ERROR_MSG + nIndex)
    Beep
    MsgBox sTmp, vbExclamation, LoadResString(RES_ERROR_MSG)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Dim rc As Long
    'see if we need to save the settings
'    If chkSaveSettings.Value = vbChecked Then
'        SaveSetting APP_CATEGORY, WIZARD_NAME, "OptionName", Option Value
'    End If
    If mbHelpStarted Then rc = WinHelp(Me.hwnd, HELP_FILE, HELP_QUIT, 0)
End Sub

'This procedure will list all fields to listbox
Private Sub AllFieldsToListBox(sTableName As String, cList As ListBox)
Dim Adofl As ADODB.Field
On Error GoTo ListError
    Screen.MousePointer = vbHourglass
    Set rs = New ADODB.Recordset
    rs.Open "[" & sTableName & "]", cnn, _
       adOpenKeyset, adLockOptimistic, adCmdTable
    cList.Clear
    For Each Adofl In rs.Fields
        cList.AddItem Adofl.Name
    Next
    Screen.MousePointer = vbDefault
    Exit Sub
ListError:
    On Error GoTo Err
    rs.Open sTableName, cnn, _
       adOpenKeyset, adLockOptimistic, adCmdTable
    cList.Clear
    For Each Adofl In rs.Fields
        cList.AddItem Adofl.Name
    Next
    Screen.MousePointer = vbDefault
    Exit Sub
Err:
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " - " & _
           Err.Description, vbExclamation, _
           "Error Listbox List"
End Sub

'This procedure will list all fields to combobox
Private Sub AllFieldsToComboBox(sTableName As String, cCbo As ComboBox)
Dim Adofl As ADODB.Field
On Error GoTo ListError
    Screen.MousePointer = vbHourglass
    Set rs = New ADODB.Recordset
    rs.Open "[" & sTableName & "]", cnn, _
       adOpenKeyset, adLockOptimistic, adCmdTable
    cCbo.Clear
    cCbo.AddItem "(None)"  '<-- default
    For Each Adofl In rs.Fields
        cCbo.AddItem Adofl.Name
    Next
    cCbo.Text = cCbo.List(0)
    Screen.MousePointer = vbDefault
    Exit Sub
ListError:
    On Error GoTo Err
    rs.Open sTableName, cnn, _
       adOpenKeyset, adLockOptimistic, adCmdTable
    cCbo.Clear
    cCbo.AddItem "(None)"  '<-- default
    For Each Adofl In rs.Fields
        cCbo.AddItem Adofl.Name
    Next
    cCbo.Text = cCbo.List(0)
    Screen.MousePointer = vbDefault
    Exit Sub
Err:
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " - " & _
           Err.Description, vbExclamation, _
           "Error Combobox List"
End Sub

'Move down an item in listbox one line, (c) W. Matos
'Thanks to W. Matos for this code
Public Sub MoveUp(lb As Object)
    Dim tmpField As String
    Dim i As Integer
    i = lb.ListIndex
    If lb.ListCount < 1 Then Exit Sub

    If i > 0 And i < lb.ListCount Then
        tmpField = lb.List(i - 1)
        lb.List(i - 1) = lb.List(i)
        lb.List(i) = tmpField
        lb.ListIndex = i - 1
        lb.Selected(i - 1) = True
        lb.Selected(i) = False
    End If
    CheckMoveButton
End Sub

'Move down an item in listbox one line, (c) W. Matos
'Thanks to W. Matos for this code
Public Sub MoveDown(lb As Object)
    Dim tmpField As String
    Dim i As Integer
    i = lb.ListIndex
    If lb.ListCount < 1 Then Exit Sub

    If i > -1 And i < lb.ListCount - 1 Then
        tmpField = lb.List(i + 1)
        lb.List(i + 1) = lb.List(i)
        lb.List(i) = tmpField
        lb.ListIndex = i + 1
        lb.Selected(i + 1) = True
        lb.Selected(i) = False
    End If
    CheckMoveButton
End Sub

'Check, which button can be accessed, up or down...
Private Sub CheckMoveButton()
  If lstSelectedFields(0).Text = lstSelectedFields(0).List(0) Then
     cmdMoveUp(0).Enabled = False
     cmdMoveDown(0).Enabled = True
  ElseIf lstSelectedFields(0).Text = lstSelectedFields(0).List(lstSelectedFields(0).ListCount - 1) Then
     cmdMoveUp(0).Enabled = True
     cmdMoveDown(0).Enabled = False
  Else
     cmdMoveUp(0).Enabled = True
     cmdMoveDown(0).Enabled = True
  End If

  If lstSelectedFields(1).Text = lstSelectedFields(1).List(0) Then
     cmdMoveUp(1).Enabled = False
     cmdMoveDown(1).Enabled = True
  ElseIf lstSelectedFields(1).Text = lstSelectedFields(1).List(lstSelectedFields(1).ListCount - 1) Then
     cmdMoveUp(1).Enabled = True
     cmdMoveDown(1).Enabled = False
  Else
     cmdMoveUp(1).Enabled = True
     cmdMoveDown(1).Enabled = True
  End If
  
  If lstSelectedFields(0).ListCount = 0 Or _
     lstSelectedFields(0).ListCount = 1 Then
     cmdMoveUp(0).Enabled = False
     cmdMoveDown(0).Enabled = False
  End If
  
  If lstSelectedFields(1).ListCount = 0 Or _
     lstSelectedFields(1).ListCount = 1 Then
     cmdMoveUp(1).Enabled = False
     cmdMoveDown(1).Enabled = False
  End If
    
End Sub

'If an item in lstAvailableFields was double-clicked,
'move the item to the lstSelectedFields.
Private Sub lstAvailableFields_DblClick(Index As Integer)
Select Case Index
       Case 0: Call cmdMove1R_Click(0)
       Case 1: Call cmdMove1R_Click(1)
End Select
End Sub

Private Sub lstFormLayout_Click()
  If lstFormLayout.Selected(0) = True Then
     lblDescLayout.Caption = LoadResString(118)
     '"The selected record " & _
     "that was pointed in DataGrid below will be displayed " & _
     "in TextBox above. You can see all records in DataGrid."
  Else
     lblDescLayout.Caption = LoadResString(119)
     '"Master record will be " & _
     "displayed in TextBox above, and its Detail in DataGrid below. " & _
     "Records in DataGrid belong to Master record in TextBox."
  End If
End Sub

Private Sub lstSelectedFields_Click(Index As Integer)
  CheckMoveButton
End Sub

'If an item in lstSelectedFields was double-clicked,
'move the item to the lstAvailableFields.
Private Sub lstSelectedFields_DblClick(Index As Integer)
Select Case Index
       Case 0: Call cmdMove1L_Click(0)
       Case 1: Call cmdMove1L_Click(1)
End Select
End Sub

'If an item in lstDBFormat was double-clicked,
'click Next button (continue to next step)
Private Sub lstDBFormat_DblClick()
  cmdNav(BTN_NEXT).Value = True
End Sub



'----------------- ODBC Connect -------------------
Private Sub cboDSNList_Click()
    On Error Resume Next
    If cboDSNList.Text = "(None)" Then
        txtServer.Enabled = True
        txtServer.BackColor = vbWhite
        cboDrivers.Enabled = True
        cboDrivers.BackColor = vbWhite
    Else
        txtServer.Enabled = False
        txtServer.BackColor = &H8000000F
        cboDrivers.Enabled = False
        cboDrivers.BackColor = &H8000000F
    End If
End Sub

Sub GetDSNsAndDrivers()
    Dim i As Integer
    Dim sDSNItem As String * 1024
    Dim sDRVItem As String * 1024
    Dim sDSN As String
    Dim sDRV As String
    Dim iDSNLen As Integer
    Dim iDRVLen As Integer
    Dim lHenv As Long         'handle to the environment

    On Error Resume Next
    cboDSNList.Clear
    cboDSNList.AddItem "(None)"

    'get the DSNs
    If SQLAllocEnv(lHenv) <> -1 Then
        Do Until i <> SQL_SUCCESS
            sDSNItem = Space$(1024)
            sDRVItem = Space$(1024)
            i = SQLDataSources(lHenv, SQL_FETCH_NEXT, sDSNItem, 1024, iDSNLen, sDRVItem, 1024, iDRVLen)
            sDSN = Left$(sDSNItem, iDSNLen)
            sDRV = Left$(sDRVItem, iDRVLen)
                
            If sDSN <> Space(iDSNLen) Then
                cboDSNList.AddItem sDSN
                cboDrivers.AddItem sDRV
            End If
        Loop
    End If
    'remove the dupes
    If cboDSNList.ListCount > 0 Then
        With cboDrivers
            If .ListCount > 1 Then
                i = 0
                While i < .ListCount
                    If .List(i) = .List(i + 1) Then
                        .RemoveItem (i)
                    Else
                        i = i + 1
                    End If
                Wend
            End If
        End With
    End If
    cboDSNList.ListIndex = 0
End Sub


Private Sub txtFormName_KeyPress(KeyAscii As Integer)
'Validate every character that user type in txtUserID
Dim strValid As String
'This is the valid string user can type to this textbox
'It's up to you, if you want to add another character...
strValid = "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz0123456789_"
  If InStr(strValid, Chr(KeyAscii)) = 0 _
     And KeyAscii <> vbKeyBack _
     And KeyAscii <> vbKeyDelete _
     And KeyAscii <> 13 Then
     KeyAscii = 0  'do nothing
  End If
End Sub

Sub InitLanguage()
  Dim LangId As Long, ProgID As String
  ' Get the default language.
  LangId = GetUserDefaultLangID()
  'LangId = 1040
  
  'MsgBox LangId  '1057
  ' Build the complete class name.
  ProgID = App.EXEName & Hex$(LangId) & ".Resources"
  ' Try to create the object, but ignore errors. If this statement
  ' fails, the RS variable will point to the default DLL (English).
  On Error Resume Next
  'Set rs = CreateObject(ProgID)
  'MsgBox ProgID '1421
End Sub

