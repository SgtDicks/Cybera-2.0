VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Object = "{01646141-065C-11D4-8ED3-00E07D815373}#1.0#0"; "mbbrowse.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11535
   ControlBox      =   0   'False
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList imlLocked 
      Left            =   6540
      Top             =   6420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":059C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraLocked 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   60
      TabIndex        =   225
      Top             =   6660
      Width           =   4215
      Begin VB.Frame fraAdminPassword 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   420
         TabIndex        =   226
         Top             =   0
         Visible         =   0   'False
         Width           =   3495
         Begin VB.CommandButton cmdUnlockOptions 
            Height          =   315
            Left            =   3120
            Picture         =   "frmOptions.frx":06F6
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   0
            Width           =   375
         End
         Begin VB.TextBox txtLockedPassword 
            BorderStyle     =   0  'None
            Height          =   195
            IMEMode         =   3  'DISABLE
            Left            =   1320
            PasswordChar    =   "*"
            TabIndex        =   4
            Top             =   60
            Width           =   1695
         End
         Begin VB.Label lblLockedPassword 
            Alignment       =   1  'Right Justify
            Caption         =   "_"
            Height          =   195
            Left            =   0
            TabIndex        =   227
            Top             =   60
            Width           =   1215
         End
         Begin VB.Shape shp 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000010&
            Height          =   315
            Index           =   29
            Left            =   1260
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.CheckBox chkLocked 
         Height          =   315
         Left            =   0
         Picture         =   "frmOptions.frx":0840
         Style           =   1  'Graphical
         TabIndex        =   159
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblUnlocked 
         Caption         =   "_"
         Height          =   195
         Left            =   420
         TabIndex        =   229
         Top             =   60
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label lblLocked 
         Caption         =   "Les options sont verouillés"
         Height          =   195
         Left            =   420
         TabIndex        =   228
         Top             =   60
         Width           =   3495
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "_"
      Height          =   375
      Left            =   10140
      TabIndex        =   161
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "_"
      Default         =   -1  'True
      Height          =   375
      Left            =   8700
      TabIndex        =   160
      Top             =   6600
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog cdlPath 
      Left            =   8160
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Base de donnée CybAdmin"
      Filter          =   "Base de donnée Microsoft Acces (*.mdb)|*.mdb"
      Orientation     =   2
   End
   Begin MSComctlLib.ImageList imlTabs 
      Left            =   11040
      Top             =   -120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":098A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":0AE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":0C3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":0D98
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":0EF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":104C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":11A6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tbsTabs 
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   661
      TabWidthStyle   =   1
      Style           =   2
      HotTracking     =   -1  'True
      ImageList       =   "imlTabs"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   7
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "_"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "_"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "_"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "_"
            ImageVarType    =   2
            ImageIndex      =   4
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "_"
            ImageVarType    =   2
            ImageIndex      =   5
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "_"
            ImageVarType    =   2
            ImageIndex      =   6
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "_"
            ImageVarType    =   2
            ImageIndex      =   7
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraOptions 
      BorderStyle     =   0  'None
      Height          =   6495
      Index           =   3
      Left            =   60
      TabIndex        =   171
      Top             =   480
      Visible         =   0   'False
      Width           =   11415
      Begin VB.Frame fra 
         BorderStyle     =   0  'None
         Height          =   1155
         Index           =   0
         Left            =   60
         TabIndex        =   286
         Top             =   4860
         Width           =   1875
         Begin VB.OptionButton optPostpaidSsMode 
            Caption         =   "_"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   106
            Top             =   480
            Width           =   1875
         End
         Begin VB.OptionButton optPostpaidSsMode 
            Caption         =   "_"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   105
            Top             =   120
            Width           =   1875
         End
      End
      Begin VB.ComboBox cmbSessBatchStyle 
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   95
         Top             =   3720
         Width           =   3015
      End
      Begin VB.CheckBox chkSessBatchEnd 
         Caption         =   "_"
         Height          =   315
         Left            =   120
         TabIndex        =   91
         Top             =   3360
         Width           =   3315
      End
      Begin VB.CommandButton cmdSessBachEndOpen 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6960
         Picture         =   "frmOptions.frx":1300
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   3360
         Width           =   375
      End
      Begin VB.CommandButton cmdSessBachEndPath 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6540
         TabIndex        =   93
         Top             =   3360
         Width           =   375
      End
      Begin VB.TextBox txtSessBatchEnd 
         BorderStyle     =   0  'None
         Height          =   195
         IMEMode         =   3  'DISABLE
         Left            =   3540
         TabIndex        =   92
         Top             =   3420
         Width           =   2895
      End
      Begin VB.TextBox txtSessBatchStart 
         BorderStyle     =   0  'None
         Height          =   195
         IMEMode         =   3  'DISABLE
         Left            =   3540
         TabIndex        =   88
         Top             =   3060
         Width           =   2895
      End
      Begin VB.CommandButton cmdSessBachStartPath 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6540
         TabIndex        =   89
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton cmdSessBachStartOpen 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6960
         Picture         =   "frmOptions.frx":144A
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   3000
         Width           =   375
      End
      Begin VB.CheckBox chkSessClnRestoreIMs 
         Caption         =   "_"
         Height          =   195
         Left            =   120
         TabIndex        =   79
         Top             =   1440
         Width           =   7275
      End
      Begin VB.TextBox txtAppName 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   7740
         TabIndex        =   100
         Top             =   2460
         Width           =   2655
      End
      Begin VB.CommandButton cmdForbidAppDelete 
         Enabled         =   0   'False
         Height          =   315
         Left            =   10920
         Picture         =   "frmOptions.frx":1594
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   2760
         Width           =   375
      End
      Begin VB.CommandButton cmdForbidAppAdd 
         Height          =   315
         Left            =   10920
         Picture         =   "frmOptions.frx":16DE
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   2400
         Width           =   375
      End
      Begin VB.ListBox lstKilledApps 
         Height          =   1245
         IntegralHeight  =   0   'False
         ItemData        =   "frmOptions.frx":1828
         Left            =   7680
         List            =   "frmOptions.frx":182A
         Sorted          =   -1  'True
         TabIndex        =   103
         Top             =   2760
         Width           =   3195
      End
      Begin VB.CommandButton cmdPathForbidApps 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10500
         TabIndex        =   101
         Top             =   2400
         Width           =   375
      End
      Begin VB.CheckBox chkForbidIE 
         Caption         =   "_"
         Height          =   195
         Left            =   7680
         TabIndex        =   96
         Top             =   780
         Width           =   3675
      End
      Begin VB.CheckBox chkForbidNetscape 
         Caption         =   "_"
         Height          =   195
         Left            =   7680
         TabIndex        =   97
         Top             =   1140
         Width           =   3675
      End
      Begin VB.CheckBox chkForbidMozilla 
         Caption         =   "_"
         Height          =   195
         Left            =   7680
         TabIndex        =   98
         Top             =   1500
         Width           =   3675
      End
      Begin VB.CheckBox chkForbidOE 
         Caption         =   "_"
         Height          =   195
         Left            =   7680
         TabIndex        =   99
         Top             =   1860
         Width           =   3675
      End
      Begin VB.CheckBox chkSessClnRestoreBackground 
         Caption         =   "_"
         Height          =   315
         Left            =   120
         TabIndex        =   83
         Top             =   2100
         Width           =   3555
      End
      Begin VB.CheckBox chkSessClnAlignDesktopIcons 
         Caption         =   "_"
         Height          =   195
         Left            =   120
         TabIndex        =   76
         Top             =   360
         Width           =   7275
      End
      Begin VB.CheckBox chkSessClnDeleteDocHistory 
         Caption         =   "_"
         Height          =   195
         Left            =   120
         TabIndex        =   77
         Top             =   720
         Width           =   7275
      End
      Begin VB.CommandButton cmdRetreaveCurrentBackground 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6960
         Picture         =   "frmOptions.frx":182C
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   2100
         Width           =   375
      End
      Begin VB.CommandButton cmdPathBackground 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6540
         TabIndex        =   85
         Top             =   2100
         Width           =   375
      End
      Begin VB.TextBox txtSessClnRestoreBackgroundPath 
         BorderStyle     =   0  'None
         Height          =   195
         IMEMode         =   3  'DISABLE
         Left            =   3780
         TabIndex        =   84
         Top             =   2160
         Width           =   2655
      End
      Begin VB.CommandButton cmdBlankIEHomePage 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6960
         Picture         =   "frmOptions.frx":1976
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   1740
         Width           =   375
      End
      Begin VB.TextBox txtSessClnRestoreIEHomePagePath 
         BorderStyle     =   0  'None
         Height          =   195
         IMEMode         =   3  'DISABLE
         Left            =   3780
         TabIndex        =   81
         Top             =   1800
         Width           =   3075
      End
      Begin VB.CheckBox chkSessClnRestoreIEHomePage 
         Caption         =   "_"
         Height          =   315
         Left            =   120
         TabIndex        =   80
         Top             =   1740
         Width           =   3555
      End
      Begin VB.CheckBox chkSessClnRestoreIESettings 
         Caption         =   "_"
         Height          =   195
         Left            =   120
         TabIndex        =   78
         Top             =   1080
         Width           =   7275
      End
      Begin VB.CheckBox chkSessBatchStart 
         Caption         =   "_"
         Height          =   315
         Left            =   120
         TabIndex        =   87
         Top             =   3000
         Width           =   3315
      End
      Begin MSComCtl2.DTPicker dtpPostpaidSsGraceTime 
         Height          =   315
         Left            =   10320
         TabIndex        =   114
         Top             =   4620
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:mm:ss"
         Format          =   21561347
         UpDown          =   -1  'True
         CurrentDate     =   38353
         MaxDate         =   47848
         MinDate         =   36526
      End
      Begin VB.Frame fraPostpaidSsMode 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1155
         Index           =   2
         Left            =   2100
         TabIndex        =   287
         Top             =   4860
         Width           =   5895
         Begin MSComCtl2.UpDown udPostpaidSsRecharges 
            Height          =   555
            Left            =   2340
            TabIndex        =   108
            Top             =   420
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   979
            _Version        =   393216
            Value           =   -1
            Max             =   -1
            Min             =   -1000
            Enabled         =   -1  'True
         End
         Begin VB.ListBox lstPostpaidSsRecharges 
            Enabled         =   0   'False
            Height          =   885
            IntegralHeight  =   0   'False
            ItemData        =   "frmOptions.frx":1AC0
            Left            =   0
            List            =   "frmOptions.frx":1AC2
            Style           =   1  'Checkbox
            TabIndex        =   107
            Top             =   240
            Width           =   2295
         End
         Begin MSComCtl2.DTPicker dtpPostpaidSsBaseIncr 
            Height          =   315
            Left            =   4920
            TabIndex        =   109
            Top             =   60
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:mm:ss"
            Format          =   21561347
            UpDown          =   -1  'True
            CurrentDate     =   38353
            MaxDate         =   47848
            MinDate         =   36526
         End
         Begin MSComCtl2.DTPicker dtpPostpaidSsMinTime 
            Height          =   315
            Left            =   4920
            TabIndex        =   110
            Top             =   420
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:mm:ss"
            Format          =   21561347
            UpDown          =   -1  'True
            CurrentDate     =   38353
            MaxDate         =   38717
            MinDate         =   38353
         End
         Begin VB.Line lin 
            BorderColor     =   &H80000010&
            Index           =   0
            X1              =   2640
            X2              =   2640
            Y1              =   0
            Y2              =   1140
         End
         Begin VB.Label lbl 
            Caption         =   "_"
            Height          =   195
            Index           =   85
            Left            =   0
            TabIndex        =   290
            Top             =   0
            Width           =   2295
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "_"
            Height          =   195
            Index           =   87
            Left            =   2700
            TabIndex        =   289
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "_"
            Height          =   195
            Index           =   86
            Left            =   2700
            TabIndex        =   288
            Top             =   120
            Width           =   2175
         End
      End
      Begin VB.Frame fraPostpaidSsMode 
         BorderStyle     =   0  'None
         Height          =   1155
         Index           =   1
         Left            =   2100
         TabIndex        =   292
         Top             =   4860
         Width           =   5895
         Begin MSComCtl2.DTPicker dtpPostpaidSsEndIncr 
            Height          =   315
            Left            =   4920
            TabIndex        =   112
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:mm:ss"
            Format          =   21561347
            UpDown          =   -1  'True
            CurrentDate     =   38353
            MaxDate         =   47848
            MinDate         =   36526
         End
         Begin VB.OptionButton optPostpaidSsModeRechSubMode 
            Caption         =   "_"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   111
            Top             =   0
            Width           =   5835
         End
         Begin VB.OptionButton optPostpaidSsModeRechSubMode 
            Caption         =   "_"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   113
            Top             =   600
            Width           =   5835
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "_"
            Height          =   195
            Index           =   89
            Left            =   60
            TabIndex        =   293
            Top             =   300
            Width           =   4815
         End
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   90
         Left            =   8100
         TabIndex        =   295
         Top             =   4680
         Width           =   2175
      End
      Begin VB.Line lin 
         BorderColor     =   &H80000010&
         Index           =   9
         X1              =   8040
         X2              =   8040
         Y1              =   4620
         Y2              =   6000
      End
      Begin VB.Label lbl 
         BackColor       =   &H80000014&
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   84
         Left            =   60
         TabIndex        =   294
         Top             =   4620
         Width           =   7935
      End
      Begin VB.Line lin 
         BorderColor     =   &H80000010&
         Index           =   8
         X1              =   2040
         X2              =   2040
         Y1              =   4860
         Y2              =   6000
      End
      Begin VB.Shape shp 
         BorderColor     =   &H80000010&
         Height          =   1815
         Index           =   50
         Left            =   0
         Top             =   4260
         Width           =   11415
      End
      Begin VB.Label lbl 
         Caption         =   "_"
         Height          =   195
         Index           =   29
         Left            =   120
         TabIndex        =   278
         Top             =   3780
         Width           =   3315
      End
      Begin VB.Shape shp 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   48
         Left            =   3480
         Top             =   3360
         Width           =   3015
      End
      Begin VB.Shape shp 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   47
         Left            =   3480
         Top             =   3000
         Width           =   3015
      End
      Begin VB.Shape shp 
         BorderColor     =   &H80000010&
         Height          =   1515
         Index           =   13
         Left            =   0
         Top             =   2640
         Width           =   7455
      End
      Begin VB.Label lbl 
         Caption         =   "_"
         Height          =   375
         Index           =   82
         Left            =   7680
         TabIndex        =   258
         Top             =   300
         Width           =   3675
      End
      Begin VB.Shape shp 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   43
         Left            =   7680
         Top             =   2400
         Width           =   2775
      End
      Begin VB.Shape shp 
         BorderColor     =   &H80000010&
         Height          =   4155
         Index           =   42
         Left            =   7560
         Top             =   0
         Width           =   3855
      End
      Begin VB.Label lbl 
         Caption         =   "_"
         Height          =   195
         Index           =   80
         Left            =   7680
         TabIndex        =   256
         Top             =   2160
         Width           =   3675
      End
      Begin VB.Shape shp 
         BorderColor     =   &H80000010&
         Height          =   2535
         Index           =   41
         Left            =   0
         Top             =   0
         Width           =   7455
      End
      Begin VB.Shape shp 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   40
         Left            =   3720
         Top             =   1740
         Width           =   3195
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   79
         Left            =   0
         TabIndex        =   255
         Top             =   0
         Width           =   7455
      End
      Begin VB.Shape shp 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   44
         Left            =   3720
         Top             =   2100
         Width           =   2775
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   81
         Left            =   7560
         TabIndex        =   257
         Top             =   0
         Width           =   3855
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   0
         TabIndex        =   277
         Top             =   2640
         Width           =   7455
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   285
         Top             =   4260
         Width           =   11415
      End
   End
   Begin VB.Frame fraOptions 
      BorderStyle     =   0  'None
      Height          =   6495
      Index           =   6
      Left            =   60
      TabIndex        =   172
      Top             =   480
      Visible         =   0   'False
      Width           =   11415
      Begin VB.TextBox txtServiceStock 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   147
         Top             =   2220
         Width           =   2835
      End
      Begin VB.OptionButton optServiceStockMode 
         Caption         =   "_"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   8340
         TabIndex        =   145
         Top             =   1500
         Width           =   2955
      End
      Begin VB.OptionButton optServiceStockMode 
         Caption         =   "_"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   8340
         TabIndex        =   146
         Top             =   1800
         Width           =   2955
      End
      Begin VB.TextBox txtServicePrice 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   144
         Top             =   1080
         Width           =   2835
      End
      Begin VB.TextBox txtServiceLabel 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   8400
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   143
         Top             =   660
         Width           =   2835
      End
      Begin VB.TextBox txtServiceStockAlert 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   148
         Top             =   2640
         Width           =   2835
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fxgServices 
         Height          =   2295
         Left            =   180
         TabIndex        =   142
         Top             =   360
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   4048
         _Version        =   393216
         Rows            =   10
         Cols            =   5
         FixedCols       =   0
         BackColorFixed  =   -2147483628
         BackColorBkg    =   -2147483643
         GridColor       =   12632256
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
         GridLinesUnpopulated=   1
         ScrollBars      =   2
         SelectionMode   =   1
         BorderStyle     =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.CommandButton cmdServiceAddCancel 
         Caption         =   "_"
         Height          =   375
         Left            =   6240
         TabIndex        =   196
         Top             =   5580
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdServiceNew 
         Caption         =   "_"
         Height          =   375
         Left            =   10020
         TabIndex        =   149
         Top             =   5580
         Width           =   1335
      End
      Begin VB.CommandButton cmdServiceDelete 
         Caption         =   "_"
         Height          =   375
         Left            =   6240
         TabIndex        =   152
         Top             =   5580
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdServiceModify 
         Caption         =   "_"
         Height          =   375
         Left            =   8640
         TabIndex        =   150
         Top             =   5580
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdServiceAdd 
         Caption         =   "_"
         Height          =   375
         Left            =   10020
         TabIndex        =   151
         Top             =   5580
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   56
         Left            =   6240
         TabIndex        =   194
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Shape shp 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   28
         Left            =   8340
         Top             =   2160
         Width           =   2955
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   55
         Left            =   6240
         TabIndex        =   193
         Top             =   2220
         Width           =   2055
      End
      Begin VB.Shape shp 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   27
         Left            =   8340
         Top             =   1020
         Width           =   2955
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   51
         Left            =   6240
         TabIndex        =   190
         Top             =   1500
         Width           =   2055
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   50
         Left            =   6240
         TabIndex        =   189
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   49
         Left            =   6240
         TabIndex        =   188
         Top             =   660
         Width           =   2055
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "_"
         Height          =   195
         Index           =   48
         Left            =   0
         TabIndex        =   187
         Top             =   0
         Width           =   11415
      End
      Begin VB.Shape shpServiceGrid 
         BorderColor     =   &H80000010&
         Height          =   5775
         Left            =   0
         Top             =   240
         Width           =   6075
      End
      Begin VB.Shape shp 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   24
         Left            =   8340
         Top             =   600
         Width           =   2955
      End
      Begin VB.Shape shp 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   22
         Left            =   8340
         Top             =   2580
         Width           =   2955
      End
      Begin VB.Shape shp 
         BorderColor     =   &H80000010&
         Height          =   5775
         Index           =   4
         Left            =   6180
         Top             =   240
         Width           =   5235
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   6180
         TabIndex        =   195
         Top             =   240
         Width           =   5235
      End
   End
   Begin VB.Frame fraOptions 
      BorderStyle     =   0  'None
      Height          =   6495
      Index           =   1
      Left            =   60
      TabIndex        =   162
      Top             =   480
      Width           =   11415
      Begin VB.Frame fra 
         BorderStyle     =   0  'None
         Height          =   1215
         Index           =   1
         Left            =   10020
         TabIndex        =   291
         Top             =   360
         Width           =   1335
         Begin VB.OptionButton optUsageDefault 
            Height          =   195
            Index           =   4
            Left            =   60
            TabIndex        =   24
            Top             =   960
            Width           =   1515
         End
         Begin VB.OptionButton optUsageDefault 
            Height          =   195
            Index           =   2
            Left            =   60
            TabIndex        =   23
            Top             =   660
            Width           =   1515
         End
         Begin VB.OptionButton optUsageDefault 
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   22
            Top             =   360
            Width           =   1515
         End
         Begin VB.OptionButton optUsageDefault 
            Height          =   195
            Index           =   5
            Left            =   60
            TabIndex        =   21
            Top             =   60
            Visible         =   0   'False
            Width           =   1515
         End
      End
      Begin VB.CheckBox chkUsageSelfServicePrePaid 
         Caption         =   "_"
         Height          =   195
         Left            =   5880
         TabIndex        =   17
         Top             =   420
         Width           =   4455
      End
      Begin MBBrowse.BrowseFF bffBackup 
         Left            =   6240
         Top             =   5880
         _ExtentX        =   1085
         _ExtentY        =   1085
         ReturnOnlyFSDirs=   -1  'True
         StartUpPosition =   1
         HasButtons      =   0   'False
         HasLines        =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmdDBBackupPath 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5160
         TabIndex        =   281
         Top             =   3060
         Width           =   375
      End
      Begin VB.TextBox txtDBBackupPath 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   180
         TabIndex        =   280
         Top             =   3120
         Width           =   4875
      End
      Begin VB.ComboBox cmbBindInterface 
         Height          =   315
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   3435
      End
      Begin VB.CheckBox chkUsageClients 
         Caption         =   "_"
         Height          =   195
         Left            =   5880
         TabIndex        =   20
         Top             =   1320
         Width           =   4455
      End
      Begin VB.CheckBox chkUsageCards 
         Caption         =   "_"
         Height          =   195
         Left            =   5880
         TabIndex        =   19
         Top             =   1020
         Width           =   4455
      End
      Begin VB.CheckBox chkUsageSelfServicePostPaid 
         Caption         =   "_"
         Height          =   195
         Left            =   5880
         TabIndex        =   18
         Top             =   720
         Width           =   4455
      End
      Begin VB.OptionButton optWksnCnxManual 
         Caption         =   "_"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   1140
         Width           =   3855
      End
      Begin VB.OptionButton optWksnCnxAuto 
         Caption         =   "_"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   780
         Width           =   5415
      End
      Begin MSComCtl2.DTPicker dtpNumDaysInMonth 
         Height          =   315
         Left            =   7980
         TabIndex        =   27
         Top             =   3000
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd"
         Format          =   21561347
         UpDown          =   -1  'True
         CurrentDate     =   38348
         MaxDate         =   38352
         MinDate         =   38322
      End
      Begin VB.CheckBox chkBackupDBSendToWksn 
         Caption         =   "_"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   3480
         Width           =   5475
      End
      Begin VB.TextBox txtPortNum 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4920
         TabIndex        =   6
         Top             =   1080
         Width           =   555
      End
      Begin VB.CheckBox chkCompactDB 
         Caption         =   "_"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   2520
         Width           =   5475
      End
      Begin VB.TextBox txtDBPath 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   180
         TabIndex        =   7
         Top             =   2160
         Width           =   3615
      End
      Begin VB.CommandButton cmdDBPath 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3900
         TabIndex        =   8
         Top             =   2100
         Width           =   375
      End
      Begin VB.CheckBox chkBackupDB 
         Caption         =   "_"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2760
         Width           =   5475
      End
      Begin VB.TextBox txtDBPasswd 
         BorderStyle     =   0  'None
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   4440
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   2160
         Width           =   1035
      End
      Begin VB.CheckBox chkCnxTesting 
         Caption         =   "_"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   4260
         Width           =   5475
      End
      Begin VB.TextBox txtTestServerName 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   960
         TabIndex        =   14
         Top             =   4800
         Width           =   3015
      End
      Begin VB.TextBox txtTestServerPort 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4920
         TabIndex        =   15
         Top             =   4800
         Width           =   555
      End
      Begin VB.CheckBox chkWksnNotification 
         Caption         =   "_"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   5160
         Width           =   5475
      End
      Begin VB.TextBox txtMonSymbol 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   8040
         MaxLength       =   4
         TabIndex        =   25
         Top             =   2160
         Width           =   375
      End
      Begin VB.ComboBox cmbLanguage 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmOptions.frx":1AC4
         Left            =   7980
         List            =   "frmOptions.frx":1AC6
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   3480
         Width           =   3315
      End
      Begin VB.TextBox txtMonFormat 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   8040
         TabIndex        =   26
         Top             =   2580
         Width           =   3195
      End
      Begin VB.Frame fra 
         BorderStyle     =   0  'None
         Height          =   1695
         Index           =   3
         Left            =   5760
         TabIndex        =   245
         Top             =   4080
         Width           =   5895
         Begin VB.TextBox txtOffPeak 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   2280
            MaxLength       =   4
            TabIndex        =   30
            Top             =   840
            Width           =   315
         End
         Begin VB.TextBox txtDisconConRatio 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   2280
            MaxLength       =   4
            TabIndex        =   29
            Top             =   420
            Width           =   315
         End
         Begin MSComCtl2.UpDown udDisconConRatio 
            Height          =   315
            Left            =   2670
            TabIndex        =   246
            TabStop         =   0   'False
            Top             =   360
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Value           =   1
            OrigLeft        =   1920
            OrigTop         =   300
            OrigRight       =   2160
            OrigBottom      =   555
            Max             =   100
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udOffPeak 
            Height          =   315
            Left            =   2670
            TabIndex        =   247
            TabStop         =   0   'False
            Top             =   780
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Value           =   1
            OrigLeft        =   1920
            OrigTop         =   300
            OrigRight       =   2160
            OrigBottom      =   555
            Max             =   100
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtpOffPeakHourStop 
            Height          =   315
            Left            =   2220
            TabIndex        =   31
            Top             =   1200
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:mm:ss"
            Format          =   21561347
            UpDown          =   -1  'True
            CurrentDate     =   38348
         End
         Begin MSComCtl2.DTPicker dtpOffPeakHourStart 
            Height          =   315
            Left            =   4500
            TabIndex        =   32
            Top             =   1200
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:mm:ss"
            Format          =   21561347
            UpDown          =   -1  'True
            CurrentDate     =   38348
         End
         Begin VB.Shape shp 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000010&
            Height          =   315
            Index           =   46
            Left            =   2220
            Top             =   780
            Width           =   435
         End
         Begin VB.Shape shp 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000010&
            Height          =   315
            Index           =   45
            Left            =   2220
            Top             =   360
            Width           =   435
         End
         Begin VB.Shape shp 
            BorderColor     =   &H80000010&
            Height          =   1635
            Index           =   8
            Left            =   0
            Top             =   0
            Width           =   5655
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "_"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   36
            Left            =   0
            TabIndex        =   254
            Top             =   0
            Width           =   5655
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   22
            Left            =   2940
            TabIndex        =   253
            Top             =   810
            Width           =   195
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   21
            Left            =   2940
            TabIndex        =   252
            Top             =   390
            Width           =   195
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "_"
            Height          =   195
            Index           =   17
            Left            =   60
            TabIndex        =   251
            Top             =   1260
            Width           =   2115
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "_"
            Height          =   195
            Index           =   13
            Left            =   3300
            TabIndex        =   250
            Top             =   1260
            Width           =   1155
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "_"
            Height          =   195
            Index           =   12
            Left            =   60
            TabIndex        =   249
            Top             =   840
            Width           =   2115
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "_"
            Height          =   195
            Index           =   11
            Left            =   60
            TabIndex        =   248
            Top             =   420
            Width           =   2115
         End
      End
      Begin VB.Shape shp 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   49
         Left            =   120
         Top             =   3060
         Width           =   4995
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   30
         Left            =   120
         TabIndex        =   279
         Top             =   420
         Width           =   1935
      End
      Begin VB.Shape shp 
         BorderColor     =   &H80000010&
         Height          =   1635
         Index           =   11
         Left            =   5760
         Top             =   0
         Width           =   5655
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   3
         Left            =   4080
         TabIndex        =   186
         Top             =   1080
         Width           =   735
      End
      Begin VB.Shape shp 
         BorderColor     =   &H80000010&
         Height          =   1455
         Index           =   3
         Left            =   0
         Top             =   0
         Width           =   5655
      End
      Begin VB.Shape shp 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   5
         Left            =   4860
         Top             =   1020
         Width           =   675
      End
      Begin VB.Label lbl 
         Caption         =   "_"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   184
         Top             =   1860
         Width           =   3735
      End
      Begin VB.Shape shp 
         BorderColor     =   &H80000010&
         Height          =   2235
         Index           =   6
         Left            =   0
         Top             =   1560
         Width           =   5655
      End
      Begin VB.Shape shp 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   7
         Left            =   120
         Top             =   2100
         Width           =   3735
      End
      Begin VB.Shape shp 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   19
         Left            =   4380
         Top             =   2100
         Width           =   1155
      End
      Begin VB.Label lbl 
         Caption         =   "_"
         Height          =   195
         Index           =   43
         Left            =   4380
         TabIndex        =   182
         Top             =   1860
         Width           =   1095
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   40
         Left            =   60
         TabIndex        =   181
         Top             =   4800
         Width           =   795
      End
      Begin VB.Shape shp 
         BorderColor     =   &H80000010&
         Height          =   1635
         Index           =   16
         Left            =   0
         Top             =   3900
         Width           =   5655
      End
      Begin VB.Shape shp 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   17
         Left            =   900
         Top             =   4740
         Width           =   3135
      End
      Begin VB.Label lbl 
         Caption         =   "_"
         Height          =   195
         Index           =   41
         Left            =   120
         TabIndex        =   179
         Top             =   4500
         Width           =   3075
      End
      Begin VB.Shape shp 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   18
         Left            =   4860
         Top             =   4740
         Width           =   675
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   42
         Left            =   4080
         TabIndex        =   178
         Top             =   4800
         Width           =   735
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   9
         Left            =   5820
         TabIndex        =   177
         Top             =   2160
         Width           =   2115
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   10
         Left            =   5820
         TabIndex        =   176
         Top             =   3060
         Width           =   2115
      End
      Begin VB.Shape shp 
         BorderColor     =   &H80000010&
         Height          =   2235
         Index           =   9
         Left            =   5760
         Top             =   1740
         Width           =   5655
      End
      Begin VB.Shape shp 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   10
         Left            =   7980
         Top             =   2100
         Width           =   495
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   44
         Left            =   5820
         TabIndex        =   174
         Top             =   3540
         Width           =   2115
      End
      Begin VB.Shape shp 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   21
         Left            =   7980
         Top             =   2520
         Width           =   3315
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   47
         Left            =   5820
         TabIndex        =   173
         Top             =   2580
         Width           =   2115
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   37
         Left            =   5760
         TabIndex        =   175
         Top             =   1740
         Width           =   5655
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   39
         Left            =   0
         TabIndex        =   180
         Top             =   3900
         Width           =   5655
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   35
         Left            =   0
         TabIndex        =   183
         Top             =   1560
         Width           =   5655
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   34
         Left            =   0
         TabIndex        =   185
         Top             =   0
         Width           =   5655
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   5760
         TabIndex        =   259
         Top             =   0
         Width           =   5655
      End
   End
   Begin VB.Frame fraOptions 
      BorderStyle     =   0  'None
      Height          =   6495
      Index           =   5
      Left            =   60
      TabIndex        =   164
      Top             =   480
      Visible         =   0   'False
      Width           =   11415
      Begin VB.CheckBox chkRechargeTimeUnlimited 
         Caption         =   "_"
         Height          =   315
         Left            =   9480
         TabIndex        =   130
         Top             =   600
         Width           =   1875
      End
      Begin VB.ComboBox cmbRechargeValidityType 
         Height          =   315
         Left            =   8940
         Style           =   2  'Dropdown List
         TabIndex        =   133
         Top             =   1440
         Width           =   2355
      End
      Begin VB.CheckBox chkRechargeCrd 
         Caption         =   "_"
         Height          =   255
         Left            =   8340
         TabIndex        =   135
         Top             =   2340
         Width           =   2955
      End
      Begin VB.CheckBox chkRechargeCli 
         Caption         =   "_"
         Height          =   255
         Left            =   8340
         TabIndex        =   136
         Top             =   2640
         Width           =   2955
      End
      Begin VB.CheckBox chkRechargePreSs 
         Caption         =   "_"
         Height          =   255
         Left            =   8340
         TabIndex        =   137
         Top             =   2940
         Width           =   2955
      End
      Begin VB.CheckBox chkRechargePostSs 
         Caption         =   "_"
         Height          =   255
         Left            =   8340
         TabIndex        =   138
         Top             =   3240
         Width           =   2955
      End
      Begin VB.TextBox txtRechargeCaption 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   8400
         MaxLength       =   50
         TabIndex        =   134
         Top             =   1920
         Width           =   2835
      End
      Begin VB.TextBox txtRechargePrice 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   8400
         MaxLength       =   7
         TabIndex        =   131
         Top             =   1080
         Width           =   2835
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fxgRecharges 
         Height          =   2295
         Left            =   60
         TabIndex        =   128
         Top             =   300
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   4048
         _Version        =   393216
         Rows            =   10
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   -2147483628
         BackColorBkg    =   -2147483643
         GridColor       =   12632256
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
         GridLinesUnpopulated=   1
         ScrollBars      =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         BorderStyle     =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSComCtl2.DTPicker dtpRechargeTime 
         Height          =   315
         Left            =   8340
         TabIndex        =   129
         Top             =   600
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:mm:ss"
         Format          =   21561347
         UpDown          =   -1  'True
         CurrentDate     =   38353
         MaxDate         =   73415
         MinDate         =   -73046
      End
      Begin MSComCtl2.DTPicker dtpRechargeValidity 
         Height          =   315
         Left            =   8340
         TabIndex        =   132
         Top             =   1440
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yy"
         Format          =   21561347
         UpDown          =   -1  'True
         CurrentDate     =   36892
         MaxDate         =   47484
         MinDate         =   36892
      End
      Begin VB.CommandButton cmdRechargeDelete 
         Caption         =   "_"
         Height          =   375
         Left            =   6240
         TabIndex        =   141
         Top             =   5580
         Width           =   1335
      End
      Begin VB.CommandButton cmdRechargeAdd 
         Caption         =   "_"
         Height          =   375
         Left            =   10020
         TabIndex        =   140
         Top             =   5580
         Width           =   1335
      End
      Begin VB.CommandButton cmdRechargeModify 
         Caption         =   "_"
         Height          =   375
         Left            =   8640
         TabIndex        =   139
         Top             =   5580
         Width           =   1335
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   53
         Left            =   6240
         TabIndex        =   192
         Top             =   2340
         Width           =   2055
      End
      Begin VB.Shape shp 
         BorderColor     =   &H80000010&
         Height          =   5775
         Index           =   26
         Left            =   6180
         Top             =   240
         Width           =   5235
      End
      Begin VB.Shape shp 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   14
         Left            =   8340
         Top             =   1860
         Width           =   2955
      End
      Begin VB.Shape shp 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   12
         Left            =   8340
         Top             =   1020
         Width           =   2955
      End
      Begin VB.Shape shpRechargeGrid 
         BorderColor     =   &H80000010&
         Height          =   5775
         Left            =   0
         Top             =   240
         Width           =   6075
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   8
         Left            =   6240
         TabIndex        =   169
         Top             =   1500
         Width           =   2055
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "_"
         Height          =   195
         Index           =   33
         Left            =   0
         TabIndex        =   168
         Top             =   0
         Width           =   11415
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   7
         Left            =   6240
         TabIndex        =   167
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   6
         Left            =   6240
         TabIndex        =   166
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   5
         Left            =   6240
         TabIndex        =   165
         Top             =   660
         Width           =   2055
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         Caption         =   "Résumé"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   52
         Left            =   6180
         TabIndex        =   191
         Top             =   240
         Width           =   5235
      End
   End
   Begin VB.Frame fraOptions 
      BorderStyle     =   0  'None
      Height          =   6495
      Index           =   4
      Left            =   60
      TabIndex        =   170
      Top             =   480
      Visible         =   0   'False
      Width           =   11415
      Begin VB.CheckBox chkCliEnableExpiry 
         Caption         =   "_"
         Height          =   255
         Left            =   5850
         TabIndex        =   276
         Top             =   1560
         Width           =   5475
      End
      Begin VB.CheckBox chkHideCardNum 
         Caption         =   "_"
         Height          =   255
         Left            =   120
         TabIndex        =   118
         Top             =   2280
         Width           =   4515
      End
      Begin VB.CheckBox chkCliBonus 
         Caption         =   "_"
         Height          =   255
         Left            =   5850
         TabIndex        =   124
         Top             =   780
         Width           =   5475
      End
      Begin VB.TextBox txtCliBonusThreshold 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   7290
         TabIndex        =   125
         Top             =   1140
         Width           =   675
      End
      Begin VB.CheckBox chkCliAutoPurge 
         Caption         =   "_"
         Height          =   255
         Left            =   5850
         TabIndex        =   127
         Top             =   1860
         Width           =   5415
      End
      Begin VB.CheckBox chkCrdAutoPurge 
         Caption         =   "_"
         Height          =   255
         Left            =   120
         TabIndex        =   119
         Top             =   1560
         Width           =   4455
      End
      Begin VB.CommandButton cmdCrdTypeAdd 
         Caption         =   "_"
         Height          =   315
         Left            =   4380
         TabIndex        =   121
         Top             =   2880
         Width           =   1155
      End
      Begin VB.CommandButton cmdCrdTypeDelete 
         Caption         =   "_"
         Enabled         =   0   'False
         Height          =   315
         Left            =   4380
         TabIndex        =   123
         Top             =   3300
         Width           =   1155
      End
      Begin VB.ListBox lstCardTypes 
         Height          =   1140
         IntegralHeight  =   0   'False
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   122
         Top             =   3300
         Width           =   4215
      End
      Begin VB.TextBox txtCrdTypeNew 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   180
         MaxLength       =   50
         TabIndex        =   120
         Top             =   2940
         Width           =   4095
      End
      Begin MSComCtl2.DTPicker dtpCrdBonusTime 
         Height          =   315
         Left            =   4500
         TabIndex        =   117
         Top             =   1080
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:mm:ss"
         Format          =   21561347
         UpDown          =   -1  'True
         CurrentDate     =   38237
      End
      Begin VB.TextBox txtCrdBonusThreshold 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1500
         TabIndex        =   116
         Top             =   1140
         Width           =   675
      End
      Begin VB.CheckBox chkCrdBonus 
         Caption         =   "_"
         Height          =   255
         Left            =   120
         TabIndex        =   115
         Top             =   780
         Width           =   4515
      End
      Begin MSComCtl2.UpDown udCrdListFillInterval 
         Height          =   255
         Left            =   5280
         TabIndex        =   260
         TabStop         =   0   'False
         Top             =   360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "lblCrdListFillInterval"
         BuddyDispid     =   196721
         OrigLeft        =   2820
         OrigTop         =   420
         OrigRight       =   3060
         OrigBottom      =   675
         Increment       =   10
         Max             =   60
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown udCrdPurgeDelta 
         Height          =   255
         Left            =   5280
         TabIndex        =   265
         TabStop         =   0   'False
         Top             =   1860
         Width           =   240
         _ExtentX        =   370
         _ExtentY        =   450
         _Version        =   327681
         BuddyControl    =   "lblCrdPurgeDelta"
         BuddyDispid     =   196720
         OrigLeft        =   2820
         OrigTop         =   1500
         OrigRight       =   3060
         OrigBottom      =   1755
         Max             =   24
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtpCliBonusTime 
         Height          =   315
         Left            =   10260
         TabIndex        =   126
         Top             =   1080
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:mm:ss"
         Format          =   21561347
         UpDown          =   -1  'True
         CurrentDate     =   38237
      End
      Begin MSComCtl2.UpDown udCliListFillInterval 
         Height          =   255
         Left            =   11010
         TabIndex        =   268
         TabStop         =   0   'False
         Top             =   360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "lblCliListFillInterval"
         BuddyDispid     =   196718
         OrigLeft        =   2820
         OrigTop         =   420
         OrigRight       =   3060
         OrigBottom      =   675
         Increment       =   10
         Max             =   60
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown udCliPurgeDelta 
         Height          =   255
         Left            =   11010
         TabIndex        =   269
         TabStop         =   0   'False
         Top             =   2160
         Width           =   240
         _ExtentX        =   370
         _ExtentY        =   450
         _Version        =   327681
         BuddyControl    =   "lblCliPurgeDelta"
         BuddyDispid     =   196719
         OrigLeft        =   2820
         OrigTop         =   1500
         OrigRight       =   3060
         OrigBottom      =   1755
         Max             =   24
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Line lin 
         BorderColor     =   &H80000010&
         Index           =   7
         X1              =   0
         X2              =   5640
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Line lin 
         BorderColor     =   &H80000010&
         Index           =   6
         X1              =   5760
         X2              =   11370
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Line lin 
         BorderColor     =   &H80000010&
         Index           =   5
         X1              =   5760
         X2              =   11370
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Shape shp 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   0
         Left            =   7230
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   19
         Left            =   5850
         TabIndex        =   275
         Top             =   1140
         Width           =   1335
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   18
         Left            =   8070
         TabIndex        =   274
         Top             =   1140
         Width           =   2115
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   255
         Index           =   16
         Left            =   5850
         TabIndex        =   273
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label lblCliListFillInterval 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10710
         TabIndex        =   272
         Top             =   360
         Width           =   285
      End
      Begin VB.Label lblCliPurgeDelta 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10710
         TabIndex        =   271
         Top             =   2160
         Width           =   285
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   255
         Index           =   15
         Left            =   5790
         TabIndex        =   270
         Top             =   2160
         Width           =   4875
      End
      Begin VB.Line lin 
         BorderColor     =   &H80000010&
         Index           =   4
         X1              =   0
         X2              =   5640
         Y1              =   2220
         Y2              =   2220
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   255
         Index           =   28
         Left            =   120
         TabIndex        =   267
         Top             =   1860
         Width           =   4815
      End
      Begin VB.Label lblCrdPurgeDelta 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4980
         TabIndex        =   266
         Top             =   1860
         Width           =   285
      End
      Begin VB.Line lin 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   0
         X2              =   5640
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Line lin 
         BorderColor     =   &H80000010&
         Index           =   1
         X1              =   0
         X2              =   5640
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label lblCrdListFillInterval 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4980
         TabIndex        =   264
         Top             =   360
         Width           =   285
      End
      Begin VB.Shape shp 
         BorderColor     =   &H80000010&
         Height          =   4575
         Index           =   20
         Left            =   0
         Top             =   0
         Width           =   5655
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   255
         Index           =   45
         Left            =   120
         TabIndex        =   262
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label lbl 
         Caption         =   "_"
         Height          =   195
         Index           =   65
         Left            =   120
         TabIndex        =   261
         Top             =   2640
         Width           =   5415
      End
      Begin VB.Shape shpCrdComment 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Left            =   120
         Top             =   2880
         Width           =   4215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   68
         Left            =   2280
         TabIndex        =   232
         Top             =   1140
         Width           =   2175
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   67
         Left            =   60
         TabIndex        =   231
         Top             =   1140
         Width           =   1335
      End
      Begin VB.Shape shp 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   31
         Left            =   1440
         Top             =   1080
         Width           =   795
      End
      Begin VB.Shape shp 
         BorderColor     =   &H80000010&
         Height          =   2535
         Index           =   30
         Left            =   5760
         Top             =   0
         Width           =   5625
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   66
         Left            =   5760
         TabIndex        =   230
         Top             =   0
         Width           =   5625
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   46
         Left            =   0
         TabIndex        =   263
         Top             =   0
         Width           =   5655
      End
   End
   Begin VB.Frame fraOptions 
      BorderStyle     =   0  'None
      Height          =   6495
      Index           =   7
      Left            =   60
      TabIndex        =   211
      Top             =   480
      Visible         =   0   'False
      Width           =   11415
      Begin VB.OptionButton optStatsGraphType 
         Caption         =   "_"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   158
         Top             =   3180
         Width           =   2955
      End
      Begin VB.OptionButton optStatsGraphType 
         Caption         =   "_"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   157
         Top             =   2940
         Value           =   -1  'True
         Width           =   2955
      End
      Begin VB.CommandButton cmdStatsTimeAdd 
         Caption         =   "_"
         Height          =   315
         Left            =   1860
         TabIndex        =   154
         Top             =   540
         Width           =   1155
      End
      Begin VB.CommandButton cmdStatsTimeDelete 
         Caption         =   "_"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1860
         TabIndex        =   156
         Top             =   960
         Width           =   1155
      End
      Begin MSComCtl2.DTPicker dtpStatsTime 
         Height          =   315
         Left            =   120
         TabIndex        =   153
         Top             =   540
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:mm:ss"
         Format          =   21561347
         UpDown          =   -1  'True
         CurrentDate     =   38217
      End
      Begin VB.ListBox lstStatsTimes 
         Height          =   1215
         IntegralHeight  =   0   'False
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   155
         Top             =   960
         Width           =   1635
      End
      Begin ComCtl2.UpDown udStatsGraphStart 
         Height          =   255
         Left            =   1980
         TabIndex        =   216
         TabStop         =   0   'False
         Top             =   3780
         Width           =   240
         _ExtentX        =   370
         _ExtentY        =   450
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "lblStatsGraphStart"
         BuddyDispid     =   196728
         OrigLeft        =   2100
         OrigTop         =   1500
         OrigRight       =   2340
         OrigBottom      =   1755
         Max             =   23
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown udStatsGraphStop 
         Height          =   255
         Left            =   1980
         TabIndex        =   219
         TabStop         =   0   'False
         Top             =   4140
         Width           =   240
         _ExtentX        =   370
         _ExtentY        =   450
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "lblStatsGraphStop"
         BuddyDispid     =   196727
         OrigLeft        =   2100
         OrigTop         =   1500
         OrigRight       =   2340
         OrigBottom      =   1755
         Max             =   24
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lbl 
         Caption         =   "_"
         Height          =   195
         Index           =   64
         Left            =   120
         TabIndex        =   224
         Top             =   300
         Width           =   2895
      End
      Begin VB.Label lbl 
         Caption         =   "_"
         Height          =   195
         Index           =   62
         Left            =   120
         TabIndex        =   223
         Top             =   3540
         Width           =   2955
      End
      Begin VB.Label lbl 
         Caption         =   "_"
         Height          =   195
         Index           =   61
         Left            =   120
         TabIndex        =   222
         Top             =   2700
         Width           =   2955
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "h"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   60
         Left            =   2220
         TabIndex        =   221
         Top             =   4140
         Width           =   135
      End
      Begin VB.Label lblStatsGraphStop 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   220
         Top             =   4140
         Width           =   285
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "h"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   59
         Left            =   2220
         TabIndex        =   218
         Top             =   3780
         Width           =   135
      End
      Begin VB.Label lblStatsGraphStart 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   217
         Top             =   3780
         Width           =   285
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   58
         Left            =   240
         TabIndex        =   215
         Top             =   4140
         Width           =   1395
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   57
         Left            =   240
         TabIndex        =   214
         Top             =   3780
         Width           =   1395
      End
      Begin VB.Shape shp 
         BorderColor     =   &H80000010&
         Height          =   2115
         Index           =   25
         Left            =   0
         Top             =   2400
         Width           =   3135
      End
      Begin VB.Shape shp 
         BorderColor     =   &H80000010&
         Height          =   2295
         Index           =   23
         Left            =   0
         Top             =   0
         Width           =   3135
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   63
         Left            =   0
         TabIndex        =   212
         Top             =   0
         Width           =   3135
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   54
         Left            =   0
         TabIndex        =   213
         Top             =   2400
         Width           =   3135
      End
   End
   Begin VB.Frame fraOptions 
      BorderStyle     =   0  'None
      Height          =   6495
      Index           =   2
      Left            =   60
      TabIndex        =   163
      Top             =   480
      Visible         =   0   'False
      Width           =   11415
      Begin VB.ComboBox cmbWksnDefaultMode 
         Height          =   315
         ItemData        =   "frmOptions.frx":1AC8
         Left            =   3420
         List            =   "frmOptions.frx":1ACA
         Style           =   2  'Dropdown List
         TabIndex        =   283
         Top             =   1620
         Width           =   1395
      End
      Begin VB.ComboBox cmbWksnDefaultState 
         Height          =   315
         ItemData        =   "frmOptions.frx":1ACC
         Left            =   3420
         List            =   "frmOptions.frx":1ACE
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   1260
         Width           =   1395
      End
      Begin VB.CommandButton cmdWksnLockDrives 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4440
         TabIndex        =   44
         Top             =   4680
         Width           =   375
      End
      Begin VB.ComboBox cmbWksnLockDrives 
         Height          =   315
         Left            =   540
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   4680
         Width           =   3855
      End
      Begin VB.ComboBox cmbURLtimespan 
         Height          =   315
         ItemData        =   "frmOptions.frx":1AD0
         Left            =   3420
         List            =   "frmOptions.frx":1AE3
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   900
         Width           =   1395
      End
      Begin VB.CheckBox chkEnablePause 
         Caption         =   "_"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   660
         Width           =   4755
      End
      Begin VB.CheckBox chkWksnMsgLogoUpload 
         Caption         =   "_"
         Height          =   375
         Left            =   5220
         TabIndex        =   51
         Top             =   1740
         Width           =   6135
      End
      Begin VB.TextBox txtPathIntro 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   6660
         TabIndex        =   52
         Top             =   2220
         Width           =   3315
      End
      Begin VB.CommandButton cmdPathIntro 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10080
         TabIndex        =   53
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox txtPathBlocked 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   6660
         TabIndex        =   56
         Top             =   2640
         Width           =   3315
      End
      Begin VB.CommandButton cmdPathBlocked 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10080
         TabIndex        =   57
         Top             =   2580
         Width           =   375
      End
      Begin VB.TextBox txtPathNoCredit 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   6660
         TabIndex        =   60
         Top             =   3060
         Width           =   3315
      End
      Begin VB.CommandButton cmdPathNoCredit 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10080
         TabIndex        =   61
         Top             =   3000
         Width           =   375
      End
      Begin VB.TextBox txtPathSessionEnd 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   6660
         TabIndex        =   64
         Top             =   3480
         Width           =   3315
      End
      Begin VB.CommandButton cmdPathSessionEnd 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10080
         TabIndex        =   65
         Top             =   3420
         Width           =   375
      End
      Begin VB.TextBox txtPathSessionPaused 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   6660
         TabIndex        =   68
         Top             =   3900
         Width           =   3315
      End
      Begin VB.CommandButton cmdPathSessionPaused 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10080
         TabIndex        =   69
         Top             =   3840
         Width           =   375
      End
      Begin VB.TextBox txtPathLogo 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   6660
         TabIndex        =   72
         Top             =   4320
         Width           =   3315
      End
      Begin VB.CommandButton cmdPathLogo 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10080
         TabIndex        =   73
         Top             =   4260
         Width           =   375
      End
      Begin VB.CommandButton cmdShellBlocked 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10500
         Picture         =   "frmOptions.frx":1B12
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   2580
         Width           =   375
      End
      Begin VB.CommandButton cmdShellIntro 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10500
         Picture         =   "frmOptions.frx":1C5C
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton cmdShellSessionEnd 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10500
         Picture         =   "frmOptions.frx":1DA6
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   3420
         Width           =   375
      End
      Begin VB.CommandButton cmdShellNoCredit 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10500
         Picture         =   "frmOptions.frx":1EF0
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton cmdShellLogo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10500
         Picture         =   "frmOptions.frx":203A
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   4260
         Width           =   375
      End
      Begin VB.CommandButton cmdShellSessionPaused 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10500
         Picture         =   "frmOptions.frx":2184
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   3840
         Width           =   375
      End
      Begin VB.CommandButton cmdPrevIntro 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10920
         Picture         =   "frmOptions.frx":22CE
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton cmdPrevBlocked 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10920
         Picture         =   "frmOptions.frx":2418
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   2580
         Width           =   375
      End
      Begin VB.CommandButton cmdPrevNoCredit 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10920
         Picture         =   "frmOptions.frx":2562
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton cmdPrevSessionEnd 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10920
         Picture         =   "frmOptions.frx":26AC
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   3420
         Width           =   375
      End
      Begin VB.CommandButton cmdPrevLogo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10920
         Picture         =   "frmOptions.frx":27F6
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   4260
         Width           =   375
      End
      Begin VB.CommandButton cmdPrevSessionPaused 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10920
         Picture         =   "frmOptions.frx":2940
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   3840
         Width           =   375
      End
      Begin VB.CheckBox chkWksnLock 
         Caption         =   "_"
         Height          =   195
         Index           =   6
         Left            =   300
         TabIndex        =   45
         Top             =   5160
         Width           =   4575
      End
      Begin VB.CheckBox chkWksnLock 
         Caption         =   "_"
         Height          =   195
         Index           =   5
         Left            =   300
         TabIndex        =   47
         Top             =   5760
         Width           =   4575
      End
      Begin VB.CheckBox chkWksnLock 
         Caption         =   "_"
         Height          =   195
         Index           =   4
         Left            =   300
         TabIndex        =   42
         Top             =   4440
         Width           =   4575
      End
      Begin VB.CheckBox chkWksnLock 
         Caption         =   "_"
         Height          =   195
         Index           =   3
         Left            =   300
         TabIndex        =   46
         Top             =   5460
         Width           =   4575
      End
      Begin VB.CheckBox chkWksnLock 
         Caption         =   "_"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   41
         Top             =   4140
         Width           =   4575
      End
      Begin VB.CheckBox chkWksnLock 
         Caption         =   "_"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   39
         Top             =   3540
         Width           =   4575
      End
      Begin VB.CheckBox chkWksnLock 
         Caption         =   "_"
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   40
         Top             =   3840
         Width           =   4575
      End
      Begin VB.CheckBox chkWksnLockdown 
         Caption         =   "_"
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   3120
         Width           =   4755
      End
      Begin VB.CheckBox chkMsgCard 
         Caption         =   "_"
         Height          =   195
         Left            =   5160
         TabIndex        =   48
         Top             =   360
         Width           =   3015
      End
      Begin VB.CheckBox chkMsgSS 
         Caption         =   "_"
         Height          =   195
         Left            =   8340
         TabIndex        =   49
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox txtMsgSSAmount 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   10260
         TabIndex        =   50
         Top             =   660
         Width           =   975
      End
      Begin VB.CheckBox chkStartForceWksnsConnected 
         Caption         =   "_"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1980
         Width           =   4755
      End
      Begin MSComCtl2.UpDown udNberWorkstations 
         Height          =   255
         Left            =   4560
         TabIndex        =   197
         TabStop         =   0   'False
         Top             =   360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "lblNberWorkstations"
         BuddyDispid     =   196770
         OrigLeft        =   2460
         OrigTop         =   360
         OrigRight       =   2700
         OrigBottom      =   615
         Max             =   14
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65537
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown udMsg1CardMinutes 
         Height          =   255
         Left            =   7260
         TabIndex        =   201
         TabStop         =   0   'False
         Top             =   600
         Width           =   240
         _ExtentX        =   370
         _ExtentY        =   450
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "lblMsg1CardMinutes"
         BuddyDispid     =   196768
         OrigLeft        =   6000
         OrigTop         =   540
         OrigRight       =   6240
         OrigBottom      =   795
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65537
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown udMsg2CardMinutes 
         Height          =   255
         Left            =   7260
         TabIndex        =   202
         TabStop         =   0   'False
         Top             =   900
         Width           =   240
         _ExtentX        =   370
         _ExtentY        =   450
         _Version        =   327681
         BuddyControl    =   "lblMsg2CardMinutes"
         BuddyDispid     =   196769
         OrigLeft        =   2820
         OrigTop         =   1500
         OrigRight       =   3060
         OrigBottom      =   1755
         Max             =   9
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65537
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udFollowInterval 
         Height          =   255
         Left            =   4260
         TabIndex        =   241
         TabStop         =   0   'False
         Top             =   2340
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   393216
         Value           =   5
         BuddyControl    =   "lblFollowInterval"
         BuddyDispid     =   196767
         OrigLeft        =   2820
         OrigTop         =   420
         OrigRight       =   3060
         OrigBottom      =   675
         Max             =   30
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.CheckBox chkEnableLoggingURLs 
         Caption         =   "_"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   83
         Left            =   60
         TabIndex        =   284
         Top             =   1680
         Width           =   3315
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   31
         Left            =   60
         TabIndex        =   282
         Top             =   1320
         Width           =   3315
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   78
         Left            =   4500
         TabIndex        =   244
         Top             =   2340
         Width           =   315
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   77
         Left            =   120
         TabIndex        =   243
         Top             =   2340
         Width           =   3795
      End
      Begin VB.Label lblFollowInterval 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   242
         Top             =   2340
         Width           =   285
      End
      Begin VB.Shape shp 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   39
         Left            =   6600
         Top             =   2160
         Width           =   3435
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   75
         Left            =   5100
         TabIndex        =   239
         Top             =   2220
         Width           =   1455
      End
      Begin VB.Shape shp 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   38
         Left            =   6600
         Top             =   2580
         Width           =   3435
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   74
         Left            =   5100
         TabIndex        =   238
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Shape shp 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   37
         Left            =   6600
         Top             =   3000
         Width           =   3435
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   73
         Left            =   5100
         TabIndex        =   237
         Top             =   3060
         Width           =   1455
      End
      Begin VB.Shape shp 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   36
         Left            =   6600
         Top             =   3420
         Width           =   3435
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   72
         Left            =   5100
         TabIndex        =   236
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Shape shp 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   35
         Left            =   6600
         Top             =   3840
         Width           =   3435
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   71
         Left            =   5100
         TabIndex        =   235
         Top             =   3900
         Width           =   1455
      End
      Begin VB.Shape shp 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   34
         Left            =   6600
         Top             =   4260
         Width           =   3435
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   70
         Left            =   5100
         TabIndex        =   234
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Shape shp 
         BorderColor     =   &H80000010&
         Height          =   3315
         Index           =   33
         Left            =   5040
         Top             =   1380
         Width           =   6375
      End
      Begin VB.Shape shp 
         BorderColor     =   &H80000010&
         Height          =   3255
         Index           =   32
         Left            =   0
         Top             =   2820
         Width           =   4935
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   23
         Left            =   7500
         TabIndex        =   210
         Top             =   600
         Width           =   315
      End
      Begin VB.Label lblMsg1CardMinutes 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6960
         TabIndex        =   209
         Top             =   600
         Width           =   285
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   24
         Left            =   5160
         TabIndex        =   208
         Top             =   630
         Width           =   1755
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   25
         Left            =   7500
         TabIndex        =   207
         Top             =   900
         Width           =   315
      End
      Begin VB.Label lblMsg2CardMinutes 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6960
         TabIndex        =   206
         Top             =   900
         Width           =   285
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   26
         Left            =   5220
         TabIndex        =   205
         Top             =   900
         Width           =   1695
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   27
         Left            =   8460
         TabIndex        =   204
         Top             =   660
         Width           =   1635
      End
      Begin VB.Line lin 
         BorderColor     =   &H80000010&
         Index           =   2
         X1              =   8220
         X2              =   8220
         Y1              =   360
         Y2              =   1140
      End
      Begin VB.Shape shp 
         BorderColor     =   &H80000010&
         Height          =   1275
         Index           =   1
         Left            =   5040
         Top             =   0
         Width           =   6375
      End
      Begin VB.Shape shp 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   2
         Left            =   10200
         Top             =   600
         Width           =   1095
      End
      Begin VB.Shape shp 
         BorderColor     =   &H80000010&
         Height          =   2715
         Index           =   15
         Left            =   0
         Top             =   0
         Width           =   4935
      End
      Begin VB.Label lblNberWorkstations 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   199
         Top             =   360
         Width           =   345
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   198
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   38
         Left            =   0
         TabIndex        =   200
         Top             =   0
         Width           =   4935
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   32
         Left            =   5040
         TabIndex        =   203
         Top             =   0
         Width           =   6375
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   69
         Left            =   0
         TabIndex        =   233
         Top             =   2820
         Width           =   4935
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   76
         Left            =   5040
         TabIndex        =   240
         Top             =   1380
         Width           =   6375
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ OPTIONS FORM _______________________________|
'|  Name:          frmOptions.frm                                              |
'|  Type:          VB6 Form                                                    |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
'|  Description:   Options form. contains all user settable settings           |
'|  Copyright:     (C) 2004-2006 Samuel Monsarrat (smonsarr@users.sf.net)      |
'|                                                                             |
'| This file is part of Cybera.                                                |
'|                                                                             |
'| Cybera is free software; you can redistribute it and/or modify it under the |
'| terms of the GNU General Public License as published by the Free Software   |
'| Foundation; either version 2 of the License, or (at your option) any later  |
'| version.                                                                    |
'|                                                                             |
'| Cybera is distributed in the hope that it will be useful, but WITHOUT ANY   |
'| WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS   |
'| FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more      |
'| details.                                                                    |
'|                                                                             |
'| You should have received a copy of the GNU General Public License along     |
'| with Cybera; if not, write to the Free Software Foundation, Inc.,           |
'| 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA                     |
'|                                                                             |
'|_____________________________________________________________________________|

' $Id: frmOptions.frm 24 2006-01-18 17:35:36Z sam $

Option Explicit

Dim mblnSettingsChanged As Boolean
Dim mblnWorkstationLockDownChanged As Boolean
Dim mblnDBPasswdChanged As Boolean
Dim mblnDBPathChanged As Boolean
Dim mlngSelectedRechargeIndex As Long
Dim mlngSelectedServiceIndex As Long
Dim mblnServiceChanged As Boolean
Dim mstrLangPacks() As String
Dim muAdapters As typAdapters
Dim mblnPostPaidRechargeFill As Boolean
Dim mblnPostPaidRechargeChanged As Boolean
Dim mlngSelectedPostPaidRecharge As Long
Dim mlngSelectedRechargeOrder As Long
Dim mdtePostPaidRechargeMaxBaseIncr As Date
Dim mdtePostPaidRechargeMaxMinTime As Date


'*******************************************************************************
Private Sub chkBackupDB_Click()

    chkBackupDBSendToWksn.enabled = chkBackupDB.value = vbChecked
    txtDBBackupPath.enabled = chkBackupDB.value = vbChecked
    cmdDBBackupPath.enabled = chkBackupDB.value = vbChecked
    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub chkCliAutoPurge_Click()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub chkCliBonus_Click()
On Error Resume Next

    mblnSettingsChanged = True
    dtpCliBonusTime.enabled = chkCliBonus.value = vbChecked
    txtCliBonusThreshold.enabled = chkCliBonus.value = vbChecked
    lbl(18).enabled = chkCliBonus.value = vbChecked
    lbl(19).enabled = chkCliBonus.value = vbChecked
End Sub


'*******************************************************************************
Private Sub chkCliEnableExpiry_Click()

    mblnSettingsChanged = True
    chkCliAutoPurge.enabled = chkCliEnableExpiry.value = vbChecked
    udCliPurgeDelta.enabled = chkCliEnableExpiry.value = vbChecked
    lblCliPurgeDelta.enabled = chkCliEnableExpiry.value = vbChecked
    lbl(15).enabled = chkCliEnableExpiry.value = vbChecked
End Sub


'*******************************************************************************
Private Sub chkCrdAutoPurge_Click()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub chkCrdBonus_Click()
On Error Resume Next

    mblnSettingsChanged = True
    dtpCrdBonusTime.enabled = chkCrdBonus.value = vbChecked
    txtCrdBonusThreshold.enabled = chkCrdBonus.value = vbChecked
    lbl(67).enabled = chkCrdBonus.value = vbChecked
    lbl(68).enabled = chkCrdBonus.value = vbChecked
End Sub


'*******************************************************************************
Private Sub chkForbidIE_Click()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub chkForbidMozilla_Click()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub chkForbidNetscape_Click()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub chkForbidOE_Click()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub chkRechargeTimeUnlimited_Click()

    dtpRechargeTime.enabled = chkRechargeTimeUnlimited.value = vbUnchecked
    chkRechargePostSs.enabled = chkRechargeTimeUnlimited.value = vbUnchecked
    chkRechargePreSs.enabled = chkRechargeTimeUnlimited.value = vbUnchecked
    If chkRechargeTimeUnlimited.value = vbChecked Then
        chkRechargePostSs.value = vbUnchecked
        chkRechargePreSs.value = vbUnchecked
        dtpRechargeTime.value = NULL_DATE
    End If
End Sub


'*******************************************************************************
Private Sub chkSessBatchEnd_Click()

    mblnSettingsChanged = True
End Sub

'*******************************************************************************
Private Sub chkSessBatchStart_Click()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub chkSessClnAlignDesktopIcons_Click()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub chkSessClnDeleteDocHistory_Click()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub chkSessClnRestoreBackground_Click()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub chkSessClnRestoreIEHomePage_Click()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub chkSessClnRestoreIESettings_Click()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub chkSessClnRestoreIMs_Click()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub chkUsageCards_Click()
On Error Resume Next

    mblnSettingsChanged = True
    If chkUsageSelfServicePostPaid.value = vbUnchecked And chkUsageClients.value = vbUnchecked And chkUsageSelfServicePrePaid.value = vbUnchecked Then chkUsageCards.value = vbChecked
    optUsageDefault(enuSessionType.Card).enabled = chkUsageCards.value = vbChecked
    If optUsageDefault(enuSessionType.Card).value Or Not (optUsageDefault(enuSessionType.Selfservice).value Or optUsageDefault(enuSessionType.Card).value Or optUsageDefault(enuSessionType.Client).value) Then subRedefineUsageDefault
End Sub


'*******************************************************************************
Private Sub chkUsageClients_Click()
On Error Resume Next

    mblnSettingsChanged = True
    If chkUsageCards.value = vbUnchecked And chkUsageSelfServicePostPaid.value = vbUnchecked And chkUsageSelfServicePrePaid.value = vbUnchecked Then chkUsageClients.value = vbChecked
    optUsageDefault(enuSessionType.Client).enabled = chkUsageClients.value = vbChecked
    If optUsageDefault(enuSessionType.Client).value Or Not (optUsageDefault(enuSessionType.Selfservice).value Or optUsageDefault(enuSessionType.Card).value Or optUsageDefault(enuSessionType.Client).value) Then subRedefineUsageDefault
End Sub


'*******************************************************************************
Private Sub chkUsageSelfServicePostPaid_Click()
On Error Resume Next

    mblnSettingsChanged = True
    If chkUsageCards.value = vbUnchecked And chkUsageClients.value = vbUnchecked And chkUsageSelfServicePrePaid.value = vbUnchecked Then chkUsageSelfServicePostPaid.value = vbChecked
    optUsageDefault(enuSessionType.Selfservice).enabled = chkUsageSelfServicePostPaid.value = vbChecked
    If optUsageDefault(enuSessionType.Selfservice).value Or Not (optUsageDefault(enuSessionType.Selfservice).value Or optUsageDefault(enuSessionType.Card).value Or optUsageDefault(enuSessionType.Client).value) Then subRedefineUsageDefault
End Sub


'*******************************************************************************
Private Sub chkUsageSelfServicePrePaid_Click()
On Error Resume Next
    
    mblnSettingsChanged = True
    If chkUsageCards.value = vbUnchecked And chkUsageClients.value = vbUnchecked And chkUsageSelfServicePostPaid.value = vbUnchecked Then chkUsageSelfServicePrePaid.value = vbChecked
    optUsageDefault(enuSessionType.SelfServicePrePaid).enabled = chkUsageSelfServicePrePaid.value = vbChecked
'    If optUsageDefault(enuSessionType.SelfServicePrePaid).value Then subRedefineUsageDefault
End Sub


'*******************************************************************************
Private Sub chkWksnLock_Click(Index As Integer)

    mblnSettingsChanged = True
    mblnWorkstationLockDownChanged = True
    If Index = 4 Then
        cmbWksnLockDrives.enabled = chkWksnLock(4).value = vbChecked
        cmdWksnLockDrives.enabled = chkWksnLock(4).value = vbChecked And cmbWksnLockDrives.ListIndex = vbChecked
    End If
End Sub

'*******************************************************************************
Private Sub chkWksnLockdown_Click()
Dim i As Integer

    For i = 0 To chkWksnLock.Count - 1
        chkWksnLock(i).enabled = chkWksnLockdown.value = vbChecked
    Next i
    cmbWksnLockDrives.enabled = chkWksnLockdown.value = vbChecked And chkWksnLock(4).value = vbChecked
    cmdWksnLockDrives.enabled = chkWksnLockdown.value = vbChecked And chkWksnLock(4).value = vbChecked And cmbWksnLockDrives.ListIndex = vbChecked
    mblnSettingsChanged = True
    mblnWorkstationLockDownChanged = True
End Sub


'*******************************************************************************
Private Sub chkWksnMsgLogoUpload_Click()

    mblnSettingsChanged = True
    lbl(70).enabled = chkWksnMsgLogoUpload.value = vbChecked
    lbl(71).enabled = chkWksnMsgLogoUpload.value = vbChecked
    lbl(72).enabled = chkWksnMsgLogoUpload.value = vbChecked
    lbl(73).enabled = chkWksnMsgLogoUpload.value = vbChecked
    lbl(74).enabled = chkWksnMsgLogoUpload.value = vbChecked
    lbl(75).enabled = chkWksnMsgLogoUpload.value = vbChecked
    txtPathIntro.enabled = chkWksnMsgLogoUpload.value = vbChecked
    txtPathBlocked.enabled = chkWksnMsgLogoUpload.value = vbChecked
    txtPathLogo.enabled = chkWksnMsgLogoUpload.value = vbChecked
    txtPathNoCredit.enabled = chkWksnMsgLogoUpload.value = vbChecked
    txtPathSessionEnd.enabled = chkWksnMsgLogoUpload.value = vbChecked
    txtPathSessionPaused.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdPathIntro.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdPathBlocked.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdPathLogo.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdPathNoCredit.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdPathSessionEnd.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdPathSessionPaused.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdShellIntro.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdShellBlocked.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdShellLogo.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdShellNoCredit.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdShellSessionEnd.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdShellSessionPaused.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdPrevIntro.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdPrevBlocked.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdPrevLogo.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdPrevNoCredit.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdPrevSessionEnd.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdPrevSessionPaused.enabled = chkWksnMsgLogoUpload.value = vbChecked
End Sub


'*******************************************************************************
Private Sub chkWksnNotification_Click()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub chkCnxTesting_Click()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub chkEnableLoggingURLs_Click()

    mblnSettingsChanged = True
    cmbURLtimespan.enabled = chkEnableLoggingURLs = vbChecked
End Sub


'*******************************************************************************
Private Sub chkEnablePause_Click()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub chkHideCardNum_Click()

    mblnSettingsChanged = True
End Sub

'*******************************************************************************
Private Sub chkLocked_Click()
On Error Resume Next

    fraAdminPassword.Visible = chkLocked.value = vbChecked
    lblUnlocked.Visible = chkLocked.value = vbChecked
    lblLocked.Visible = chkLocked.value = vbUnchecked
    cmdUnlockOptions.Default = chkLocked.value = vbChecked
    cmdSave.Default = chkLocked.value = vbUnchecked
    
    If chkLocked.value = vbChecked Then chkLocked.ToolTipText = funLoadResString(10525)
    If chkLocked.value = vbChecked Then Call funSetFocus(txtLockedPassword)
    If chkLocked.value = vbUnchecked Then Call subDisableAllControls
    If chkLocked.value = vbUnchecked Then Call funSetFocus(tbsTabs)
    If chkLocked.value = vbChecked Then txtLockedPassword.Text = ""
    If chkLocked.value = vbChecked And guServOpts.uConfig.strPasswd = "" Then Call cmdUnlockOptions_Click
End Sub

'*******************************************************************************
Private Sub chkMsgCard_Click()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub chkMsgSS_Click()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub chkStartForceWksnsConnected_Click()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub cmbBindInterface_Click()

    mblnSettingsChanged = True
    cmbBindInterface.ToolTipText = cmbBindInterface.List(cmbBindInterface.ListIndex)
                    
    cmbBindInterface.ToolTipText = funLoadResString(10884) & muAdapters.Item(cmbBindInterface.ItemData(cmbBindInterface.ListIndex)).IpAddress(0) & funLoadResString(10885) & muAdapters.Item(cmbBindInterface.ItemData(cmbBindInterface.ListIndex)).Description

End Sub

'*******************************************************************************
Private Sub cmbLanguage_Change()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub cmbSessBatchStyle_Change()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub cmbURLtimespan_Click()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub cmbWksnDefaultMode_Click()

    mblnSettingsChanged = True
    If cmbWksnDefaultMode.ItemData(cmbWksnDefaultMode.ListIndex) = enuWksnMode.Disconnected Then chkStartForceWksnsConnected.value = vbUnchecked
    chkStartForceWksnsConnected.enabled = cmbWksnDefaultMode.ItemData(cmbWksnDefaultMode.ListIndex) = enuWksnMode.Connected
End Sub


'*******************************************************************************
Private Sub cmbWksnDefaultState_Click()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub cmbWksnLockDrives_Click()

    cmdWksnLockDrives.enabled = cmbWksnLockDrives.ListIndex = vbChecked
    mblnSettingsChanged = True
    mblnWorkstationLockDownChanged = True
End Sub

Private Sub cmdBlankIEHomePage_Click()

    txtSessClnRestoreIEHomePagePath.Text = URL_EMPTY_IE
End Sub


'*******************************************************************************
Private Sub cmdCancel_Click()
On Error GoTo cmdCancel_ClickError

    If mblnSettingsChanged Then
        If funMsg(enuMsg.SaveSettings) = vbYes Then Call Unload(Me)
    Else
        Call Unload(Me)
    End If
    Exit Sub
    
cmdCancel_ClickError:
    Call funErrorManager("frmOptions.cmdCancel_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdCrdTypeAdd_Click()
Dim i As Integer
On Error GoTo cmdCrdTypeAdd_ClickError

    If txtCrdTypeNew.Text = "" Then Exit Sub
    For i = 0 To lstCardTypes.ListCount - 1
        If LCase(lstCardTypes.List(i)) = LCase(txtCrdTypeNew.Text) Then Exit Sub
    Next i
    Call lstCardTypes.AddItem(txtCrdTypeNew.Text)
    txtCrdTypeNew.Text = ""
    mblnSettingsChanged = True
    Exit Sub
    
cmdCrdTypeAdd_ClickError:
    Call funErrorManager("frmOptions.cmdCrdTypeAdd_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdCrdTypeDelete_Click()
On Error GoTo cmdCrdTypeDelete_ClickError

    Call lstCardTypes.RemoveItem(lstCardTypes.ListIndex)
    mblnSettingsChanged = True
    cmdCrdTypeDelete.enabled = False
    Exit Sub
    
cmdCrdTypeDelete_ClickError:
    Call funErrorManager("frmOptions.cmdCrdTypeDelete_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdDBBackupPath_Click()
On Error GoTo cmdDBBackupPath_ClickError
    
    bffBackup.InitialFolder = txtDBBackupPath.Text
    If bffBackup.Browse Then txtDBBackupPath.Text = bffBackup.SelectedItem
    Exit Sub

cmdDBBackupPath_ClickError:
    Call funErrorManager("frmOptions.cmdDBBackupPath_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdDBPath_Click()
On Error GoTo cmdDBPath_ClickError

    Call subCdlOpenFile(cdlPath, txtDBPath, guServOpts.uDB.strDBPath, funGetFolderPath(CSIDL_DESKTOP), funLoadResString(10209), funLoadResString(10000) & " - " & funLoadResString(10208))
    Exit Sub

cmdDBPath_ClickError:
    Call funErrorManager("frmOptions.cmdDBPath_ClickError", True)
End Sub


Private Sub cmdForbidAppAdd_Click()
On Error GoTo cmdForbidAppAdd_ClickError
    
    If txtAppName.Text <> "" Then
        Call lstKilledApps.AddItem(txtAppName.Text)
        txtAppName.Text = ""
        Call funSetFocus(txtAppName)
        mblnSettingsChanged = True
    End If
    Exit Sub

cmdForbidAppAdd_ClickError:
    Call funErrorManager("frmOptions.cmdForbidAppAdd_ClickError", True)
End Sub

Private Sub cmdForbidAppDelete_Click()
On Error GoTo cmdForbidAppDelete_ClickError
    
    If lstKilledApps.ListIndex = -1 Then
        Call funMsg(enuMsg.NoElementSelected)
    Else
        lstKilledApps.RemoveItem (lstKilledApps.ListIndex)
        mblnSettingsChanged = True
        cmdForbidAppDelete.enabled = False
    End If
    Exit Sub

cmdForbidAppDelete_ClickError:
    Call funErrorManager("frmOptions.cmdForbidAppDelete_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdPathBackground_Click()
On Error GoTo cmdPathBackground_ClickError
    
    Call subCdlOpenFile(cdlPath, txtSessClnRestoreBackgroundPath, guServOpts.uSessionCleanup.strBackgroundPath, funGetFolderPath(CSIDL_WINDOWS), funLoadResString(10676), funLoadResString(10677))
    Exit Sub

cmdPathBackground_ClickError:
    Call funErrorManager("frmOptions.cmdPathBackground_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdPathForbidApps_Click()
On Error GoTo cmdPathForbidApps_ClickError
    
    Call subCdlOpenFile(cdlPath, txtAppName, funGetFolderPath(CSIDL_DESKTOP), funGetFolderPath(CSIDL_DESKTOP), funLoadResString(10678), funLoadResString(10679), False)
    Exit Sub

cmdPathForbidApps_ClickError:
    Call funErrorManager("frmOptions.cmdPathForbidApps_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdPrevBlocked_Click()
On Error GoTo cmdPrevBlocked_ClickError

    Call Load(frmTextPrevew)
    frmTextPrevew.rtbMessage.filename = txtPathBlocked.Text
    Call frmTextPrevew.Show(vbModal, Me)
    Exit Sub

cmdPrevBlocked_ClickError:
Dim lngErrNumber As Long
Dim strErrSource As String
Dim strErrDesc As String

    lngErrNumber = Err.Number
    strErrSource = Err.source
    strErrDesc = Err.Description
    Select Case lngErrNumber
    Case 53, 75
        Call funMsg(FileNotFound, txtPathBlocked.Text)
    Case Else
        Call subDebugPrint(strErrSource & ": #" & lngErrNumber & " | " & strErrDesc & " | " & strErrSource)
        Call subWriteToLogFile(strErrSource & ": #" & lngErrNumber & " | " & strErrDesc & " | " & strErrSource)
        Call funMsg(UntrappedError, "frmOptions.cmdPrevBlocked_ClickError")
    End Select
End Sub


'*******************************************************************************
Private Sub cmdPrevIntro_Click()
On Error GoTo cmdPrevIntro_ClickError

    Call Load(frmTextPrevew)
    frmTextPrevew.rtbMessage.filename = txtPathIntro.Text
    Call frmTextPrevew.Show(vbModal, Me)
    Exit Sub

cmdPrevIntro_ClickError:
Dim lngErrNumber As Long
Dim strErrSource As String
Dim strErrDesc As String

    lngErrNumber = Err.Number
    strErrSource = Err.source
    strErrDesc = Err.Description
    Select Case lngErrNumber
    Case 53, 75
        Call funMsg(FileNotFound, txtPathIntro.Text)
    Case Else
        Call subDebugPrint(strErrSource & ": #" & lngErrNumber & " | " & strErrDesc & " | " & strErrSource)
        Call subWriteToLogFile(strErrSource & ": #" & lngErrNumber & " | " & strErrDesc & " | " & strErrSource)
        Call funMsg(UntrappedError, "frmOptions.cmdPrevIntro_ClickError")
    End Select
End Sub


'*******************************************************************************
Private Sub cmdPrevLogo_Click()
On Error GoTo cmdPrevLogo_ClickError

    Call Load(frmImgPrevew)
    frmImgPrevew.imgLogo.Picture = LoadPicture(txtPathLogo.Text)
    Call frmImgPrevew.Show(vbModal, Me)
    Exit Sub

cmdPrevLogo_ClickError:
Dim lngErrNumber As Long
Dim strErrSource As String
Dim strErrDesc As String

    lngErrNumber = Err.Number
    strErrSource = Err.source
    strErrDesc = Err.Description
    Select Case lngErrNumber
    Case 53, 75
        Call funMsg(FileNotFound, txtPathLogo.Text)
    Case Else
        Call subDebugPrint(strErrSource & ": #" & lngErrNumber & " | " & strErrDesc & " | " & strErrSource)
        Call subWriteToLogFile(strErrSource & ": #" & lngErrNumber & " | " & strErrDesc & " | " & strErrSource)
        Call funMsg(UntrappedError, "frmOptions.cmdPrevLogo_ClickError")
    End Select
End Sub


'*******************************************************************************
Private Sub cmdPrevNoCredit_Click()
On Error GoTo cmdPrevNoCredit_ClickError

    Call Load(frmTextPrevew)
    frmTextPrevew.rtbMessage.filename = txtPathNoCredit.Text
    Call frmTextPrevew.Show(vbModal, Me)
    Exit Sub

cmdPrevNoCredit_ClickError:
Dim lngErrNumber As Long
Dim strErrSource As String
Dim strErrDesc As String

    lngErrNumber = Err.Number
    strErrSource = Err.source
    strErrDesc = Err.Description
    Select Case lngErrNumber
    Case 53, 75
        Call funMsg(FileNotFound, txtPathNoCredit.Text)
    Case Else
        Call subDebugPrint(strErrSource & ": #" & lngErrNumber & " | " & strErrDesc & " | " & strErrSource)
        Call subWriteToLogFile(strErrSource & ": #" & lngErrNumber & " | " & strErrDesc & " | " & strErrSource)
        Call funMsg(UntrappedError, "frmOptions.cmdPrevNoCredit_ClickError")
    End Select
End Sub


'*******************************************************************************
Private Sub cmdPrevSessionEnd_Click()
On Error GoTo cmdPrevSessionEnd_ClickError

    Call Load(frmTextPrevew)
    frmTextPrevew.rtbMessage.filename = txtPathSessionEnd.Text
    Call frmTextPrevew.Show(vbModal, Me)
    Exit Sub

cmdPrevSessionEnd_ClickError:
Dim lngErrNumber As Long
Dim strErrSource As String
Dim strErrDesc As String

    lngErrNumber = Err.Number
    strErrSource = Err.source
    strErrDesc = Err.Description
    Select Case lngErrNumber
    Case 53, 75
        Call funMsg(FileNotFound, txtPathSessionEnd.Text)
    Case Else
        Call subDebugPrint(strErrSource & ": #" & lngErrNumber & " | " & strErrDesc & " | " & strErrSource)
        Call subWriteToLogFile(strErrSource & ": #" & lngErrNumber & " | " & strErrDesc & " | " & strErrSource)
        Call funMsg(UntrappedError, "frmOptions.cmdPrevSessionEnd_ClickError")
    End Select
End Sub


'*******************************************************************************
Private Sub cmdPrevSessionPaused_Click()
On Error GoTo cmdPrevSessionPaused_ClickError

    Call Load(frmTextPrevew)
    frmTextPrevew.rtbMessage.filename = txtPathSessionPaused.Text
    Call frmTextPrevew.Show(vbModal, Me)
    Exit Sub

cmdPrevSessionPaused_ClickError:
Dim lngErrNumber As Long
Dim strErrSource As String
Dim strErrDesc As String

    lngErrNumber = Err.Number
    strErrSource = Err.source
    strErrDesc = Err.Description
    Select Case lngErrNumber
    Case 53, 75
        Call funMsg(FileNotFound, txtPathSessionPaused.Text)
    Case Else
        Call subDebugPrint(strErrSource & ": #" & lngErrNumber & " | " & strErrDesc & " | " & strErrSource)
        Call subWriteToLogFile(strErrSource & ": #" & lngErrNumber & " | " & strErrDesc & " | " & strErrSource)
        Call funMsg(UntrappedError, "frmOptions.cmdPrevSessionPaused_ClickError")
    End Select
End Sub


'*******************************************************************************
Private Sub cmdRechargeAdd_Click()
On Error GoTo cmdRechargeAdd_ClickError

    If txtRechargePrice.Text = "" Then txtRechargePrice.Text = 0
'    If txtRechargeValidity.Text = "" Then txtRechargeValidity.Text = 0
'    If Not IsDate(txtRechargeTime.Text) Or Not IsNumeric(txtRechargePrice.Text) Or Not IsNumeric(txtRechargeValidity.Text) Then
'        Call funMsg(enuMsg.RechargeNotCorrectFormat)
'    Else
        Call grstRecharges.AddNew
        grstRecharges!Time = TimeValue(dtpRechargeTime.value)
        grstRecharges!unlimited_play = chkRechargeTimeUnlimited.value = vbChecked
        
        grstRecharges!Price = txtRechargePrice.Text
        grstRecharges!validity = CInt(format(dtpRechargeValidity.value, "yy"))
        grstRecharges!validity_type = cmbRechargeValidityType.ItemData(cmbRechargeValidityType.ListIndex)
        
        grstRecharges!Caption = txtRechargeCaption.Text
        grstRecharges!type_card = chkRechargeCrd.value = vbChecked
        grstRecharges!type_client = chkRechargeCli.value = vbChecked
        grstRecharges!type_prepaid_ss = chkRechargePreSs.value = vbChecked
        grstRecharges!type_postpaid_ss = chkRechargePostSs.value = vbChecked
        grstRecharges!usage_order = 0
        Call grstRecharges.Update
        Call subFillRecharges
'    End If
    Call fxgRecharges_RowColChange
    fxgRecharges.enabled = True
    Exit Sub
    
cmdRechargeAdd_ClickError:
    Call funErrorManager("frmOptions.cmdRechargeAdd_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdRechargeDelete_Click()
Dim strQuery As String
Dim rstRechargesToUpdate As DAO.Recordset, rstRecharges As DAO.Recordset
On Error GoTo cmdRechargeDelete_ClickError
    
    If mlngSelectedRechargeIndex = 0 Then
        Call funMsg(enuMsg.RechargeNotSelected)
    Else
        If funMsg(enuMsg.RechargeDeleteConfirm) = vbYes Then
            grstRecharges.Index = "PrimaryKey"
            Call grstRecharges.Seek("=", fxgRecharges.TextMatrix(mlngSelectedRechargeIndex, 5))
            If grstRecharges.NoMatch Then
                Call funMsg(enuMsg.RechargeNotFound)
                Exit Sub
            End If
            
            ' If the deleted recharge has a usage_order > 0 then decrement all subsequent recharges
            If grstRecharges!usage_order > 0 Then
                strQuery = "SELECT id FROM recharges WHERE usage_order > " & grstRecharges!usage_order & " ORDER BY usage_order"
                Set rstRechargesToUpdate = gdbsCybAdmin.OpenRecordset(strQuery, dbOpenSnapshot)
                If Not (rstRechargesToUpdate.EOF And rstRechargesToUpdate.BOF) Then
                    Set rstRecharges = gdbsCybAdmin.OpenRecordset("recharges", dbOpenTable)
                    rstRecharges.Index = "PrimaryKey"
                    Call rstRechargesToUpdate.MoveFirst
                    Do While Not rstRechargesToUpdate.EOF
                        Call rstRecharges.Seek("=", rstRechargesToUpdate!ID)
                        If Not rstRecharges.NoMatch Then
                            Call rstRecharges.Edit
                            rstRecharges!usage_order = rstRecharges!usage_order - 1
                            Call rstRecharges.Update
                        End If
                        Call rstRechargesToUpdate.MoveNext
                    Loop
                End If
            End If
            Call grstRecharges.Delete
            Call subFillRecharges
            Call fxgRecharges_RowColChange
        End If
    End If
    Exit Sub
    
cmdRechargeDelete_ClickError:
    Call funErrorManager("frmOptions.cmdRechargeDelete_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdRechargeModify_Click()
Dim strTmpTime As String
Dim blnTypePostPaidSSChanged As Boolean
Dim strQuery As String
Dim rstRechargesToUpdate As DAO.Recordset, rstRecharges As DAO.Recordset
On Error GoTo cmdRechargeModify_ClickError
    
    If mlngSelectedRechargeIndex = 0 Then
        Call funMsg(enuMsg.RechargeNotSelected)
    Else
        If funMsg(enuMsg.RechargeModifyConfirm) = vbYes Then
            grstRecharges.Index = "PrimaryKey"
            Call grstRecharges.Seek("=", fxgRecharges.TextMatrix(mlngSelectedRechargeIndex, 5))
            If grstRecharges.NoMatch Then
                Call funMsg(enuMsg.RechargeNotFound)
                Exit Sub
            End If
            Call grstRecharges.Edit
            grstRecharges!Time = TimeValue(dtpRechargeTime.value)
            grstRecharges!unlimited_play = chkRechargeTimeUnlimited.value = vbChecked
            
            grstRecharges!Price = txtRechargePrice.Text
            grstRecharges!validity = CInt(format(dtpRechargeValidity.value, "yy"))
            grstRecharges!validity_type = cmbRechargeValidityType.ItemData(cmbRechargeValidityType.ListIndex)
            
            grstRecharges!Caption = txtRechargeCaption.Text
            grstRecharges!type_card = chkRechargeCrd.value = vbChecked
            grstRecharges!type_client = chkRechargeCli.value = vbChecked
            grstRecharges!type_prepaid_ss = chkRechargePreSs.value = vbChecked
            blnTypePostPaidSSChanged = grstRecharges!type_postpaid_ss And chkRechargePostSs.value = vbUnchecked
            grstRecharges!type_postpaid_ss = chkRechargePostSs.value = vbChecked
            
            ' If the deleted recharge has a usage_order > 0 then decrement all subsequent recharges
            If blnTypePostPaidSSChanged And grstRecharges!usage_order > 0 Then
                strQuery = "SELECT id FROM recharges WHERE usage_order > " & grstRecharges!usage_order & " ORDER BY usage_order"
                Set rstRechargesToUpdate = gdbsCybAdmin.OpenRecordset(strQuery, dbOpenSnapshot)
                If Not (rstRechargesToUpdate.EOF And rstRechargesToUpdate.BOF) Then
                    Set rstRecharges = gdbsCybAdmin.OpenRecordset("recharges", dbOpenTable)
                    rstRecharges.Index = "PrimaryKey"
                    Call rstRechargesToUpdate.MoveFirst
                    Do While Not rstRechargesToUpdate.EOF
                        Call rstRecharges.Seek("=", rstRechargesToUpdate!ID)
                        If Not rstRecharges.NoMatch Then
                            Call rstRecharges.Edit
                            rstRecharges!usage_order = rstRecharges!usage_order - 1
                            Call rstRecharges.Update
                        End If
                        Call rstRechargesToUpdate.MoveNext
                    Loop
                End If
                grstRecharges!usage_order = 0
            End If
            Call grstRecharges.Update
            Call subFillRecharges
            Call fxgRecharges_RowColChange
        End If
    End If
    Exit Sub
    
cmdRechargeModify_ClickError:
    Call funErrorManager("frmOptions.cmdRechargeModify_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdRetreaveCurrentBackground_Click()
Dim hKey As Long, nBytes As Long
Dim strValue As String
Dim hResult As Long
On Error GoTo cmdRetreaveCurrentBackground_ClickError

    hResult = RegOpenKeyEx(HKEY_CURRENT_USER, "Control panel\Desktop", 0, KEY_ALL_ACCESS, hKey)
    hResult = RegQueryValueEx(hKey, "Wallpaper", 0&, REG_SZ, ByVal strValue, nBytes)
    strValue = Space$(nBytes)
    hResult = RegQueryValueEx(hKey, "Wallpaper", 0&, REG_SZ, ByVal strValue, Len(strValue))
    If nBytes > 0 Then
        txtSessClnRestoreBackgroundPath.Text = left$(strValue, nBytes - 1)
        txtSessClnRestoreBackgroundPath.ToolTipText = left$(strValue, nBytes - 1)
        txtSessClnRestoreBackgroundPath.SelStart = Len(txtSessClnRestoreBackgroundPath.Text)
    Else
        txtSessClnRestoreBackgroundPath.Text = ""
        txtSessClnRestoreBackgroundPath.ToolTipText = ""
    End If
    Call RegCloseKey(hKey)
    Exit Sub

cmdRetreaveCurrentBackground_ClickError:
    Call funErrorManager("frmOptions.cmdRetreaveCurrentBackground_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdSave_Click()
Dim i As Integer, j As Integer
Dim uWorkstation As New clsWorkstation
Dim frmTemp As Form
Dim blnTmpValue As Boolean
On Error GoTo cmdSave_ClickError

    guServOpts.uWorkstations.lngNumWorkstations = udNberWorkstations.value
    guServOpts.uConfig.intNumDaysInMonth = DatePart("d", dtpNumDaysInMonth.value)
    
    guServOpts.uTimeRatios.intDisconCon = txtDisconConRatio.Text
    guServOpts.uTimeRatios.intOffPeak = txtOffPeak.Text
    guServOpts.uTimeRatios.dteOffPeakStart = dtpOffPeakHourStart.value
    guServOpts.uTimeRatios.dteOffPeakStop = dtpOffPeakHourStop.value

    guServOpts.uDB.strDBPath = txtDBPath.Text
    guServOpts.uDB.strDBBackupPath = txtDBBackupPath.Text
    guServOpts.uDB.strDBPasswd = txtDBPasswd.Text
    guServOpts.uDB.intDBBackup = chkBackupDB.value
    guServOpts.uDB.intDBBackupSendToWorkstation = chkBackupDBSendToWksn.value
    guServOpts.uDB.intDBCompact = chkCompactDB.value
    
    guServOpts.uServer.blnAutoConnect = optWksnCnxAuto.value
    guServOpts.uServer.lngPort = txtPortNum.Text
    If muAdapters.Count > 0 Then guServOpts.uServer.strServerMac = muAdapters.Item(cmbBindInterface.ItemData(cmbBindInterface.ListIndex)).MacAdress
    
    guServOpts.uConfig.strMonSymbol = txtMonSymbol.Text
    guServOpts.uConfig.strMonFormat = txtMonFormat.Text
    guServOpts.uConfig.strInterfaceLang = mstrLangPacks(0, cmbLanguage.ItemData(cmbLanguage.ListIndex))
    guServOpts.uConfig.intInterfaceLangCharset = mstrLangPacks(3, cmbLanguage.ItemData(cmbLanguage.ListIndex))
    
    guServOpts.uWorkstations.uWksnMsg.intMsgCard = chkMsgCard.value
    guServOpts.uWorkstations.uWksnMsg.intMsgSelfService = chkMsgSS.value
    guServOpts.uWorkstations.uWksnMsg.intMsgCardAlert1 = udMsg1CardMinutes.value
    guServOpts.uWorkstations.uWksnMsg.intMsgCardAlert2 = udMsg2CardMinutes.value
    guServOpts.uWorkstations.uWksnMsg.sngMsgSelfServiceBalance = txtMsgSSAmount.Text
    
    guServOpts.uWorkstations.intStartForceWorkstationsConnected = chkStartForceWksnsConnected.value
    guServOpts.uWorkstations.intActivateSelfServicePostPaid = chkUsageSelfServicePostPaid.value
    guServOpts.uWorkstations.intActivateSelfServicePrePaid = chkUsageSelfServicePrePaid.value
    guServOpts.uWorkstations.intActivateCards = chkUsageCards.value
    guServOpts.uWorkstations.intActivateClients = chkUsageClients.value
    
    On Error Resume Next
    guServOpts.uWorkstations.intUsageDefault = enuSessionType.Unknown
    For i = optUsageDefault.LBound To optUsageDefault.UBound
        blnTmpValue = False
        blnTmpValue = optUsageDefault(i).value
        If blnTmpValue Then
            guServOpts.uWorkstations.intUsageDefault = i
            Exit For
        End If
    Next i
    On Error GoTo cmdSave_ClickError
    
    guServOpts.uWorkstations.intHideCardNum = chkHideCardNum.value
    guServOpts.uWorkstations.intLogURLs = chkEnableLoggingURLs.value
    guServOpts.uWorkstations.intLoggedURLTimeSpan = cmbURLtimespan.ItemData(cmbURLtimespan.ListIndex)
    guServOpts.uWorkstations.intEnablePause = chkEnablePause.value
    guServOpts.uWorkstations.intDefaultState = cmbWksnDefaultState.ItemData(cmbWksnDefaultState.ListIndex)
    guServOpts.uWorkstations.intDefaultMode = cmbWksnDefaultMode.ItemData(cmbWksnDefaultMode.ListIndex)
    
    guServOpts.uWorkstations.intModifyLockdown = chkWksnLockdown.value
    guServOpts.uWorkstations.intFollowInterval = udFollowInterval.value
    guServOpts.uWorkstations.uLocks.intNoClose = chkWksnLock(0).value
    guServOpts.uWorkstations.uLocks.intNoAdminTools = chkWksnLock(1).value
    guServOpts.uWorkstations.uLocks.intNoCPL = chkWksnLock(2).value
    guServOpts.uWorkstations.uLocks.intLockDesktop = chkWksnLock(3).value
    guServOpts.uWorkstations.uLocks.intNoFileSystemAccess = chkWksnLock(4).value
    guServOpts.uWorkstations.uLocks.intLockIE = chkWksnLock(5).value
    guServOpts.uWorkstations.uLocks.intNoCommonGroup = chkWksnLock(6).value
    guServOpts.uWorkstations.uLocks.intNoFileSystemAccessType = cmbWksnLockDrives.ListIndex
    guServOpts.uWorkstations.uLocks.lngLockedDrives = guServOpts.uWorkstations.uLocks.lngLockedDrivesTmp

    
    guServOpts.uCards.intCardPurgeDelta = udCrdPurgeDelta.value
    guServOpts.uCards.intAutoCardPurge = chkCrdAutoPurge.value
    guServOpts.uCards.intCardListUpdateInterval = udCrdListFillInterval.value
    guServOpts.uCards.intBonusActive = chkCrdBonus.value
    guServOpts.uCards.sngBonusThreshold = txtCrdBonusThreshold.Text
    guServOpts.uCards.dteBonusTime = dtpCrdBonusTime.value
    
    guServOpts.uClients.intEnableExpiry = chkCliEnableExpiry.value
    guServOpts.uClients.intAutoClientPurge = chkCliAutoPurge.value
    guServOpts.uClients.intClientListUpdateInterval = udCliListFillInterval.value
    guServOpts.uClients.intClientPurgeDelta = udCliPurgeDelta.value
    guServOpts.uClients.intBonusActive = chkCliBonus.value
    guServOpts.uClients.sngBonusThreshold = txtCliBonusThreshold.Text
    guServOpts.uClients.dteBonusTime = dtpCliBonusTime.value
    
    guServOpts.uCnxTest.intCnxTesting = chkCnxTesting.value
    guServOpts.uCnxTest.intWorkstationNotification = chkWksnNotification.value
    guServOpts.uCnxTest.strTestServerName = txtTestServerName.Text
    guServOpts.uCnxTest.strTestServerPort = txtTestServerPort.Text
    
    guServOpts.uWorkstations.intModifyMsgLogoFiles = chkWksnMsgLogoUpload.value
    guServOpts.uWorkstations.uPaths.strBlockedPath = txtPathBlocked.Text
    guServOpts.uWorkstations.uPaths.strIntroPath = txtPathIntro.Text
    guServOpts.uWorkstations.uPaths.strLogoPath = txtPathLogo.Text
    guServOpts.uWorkstations.uPaths.strNoCreditPath = txtPathNoCredit.Text
    guServOpts.uWorkstations.uPaths.strSessionEndPath = txtPathSessionEnd.Text
    guServOpts.uWorkstations.uPaths.strSessionPausedPath = txtPathSessionPaused.Text
    
    On Error Resume Next
    ReDim guServOpts.uStats.strStatsTimeInt(lstStatsTimes.ListCount - 1)
    On Error GoTo cmdSave_ClickError
    For i = 0 To lstStatsTimes.ListCount - 1
        guServOpts.uStats.strStatsTimeInt(i) = funDateTimeToSend(lstStatsTimes.List(i))
    Next i
    On Error Resume Next
    ReDim guServOpts.uCards.strCardTypes(lstCardTypes.ListCount - 2)
    On Error GoTo cmdSave_ClickError
    j = 0
    For i = 0 To lstCardTypes.ListCount - 1
        If lstCardTypes.ItemData(i) <> 1 Then
            guServOpts.uCards.strCardTypes(j) = lstCardTypes.List(i)
            j = j + 1
        End If
    Next i
    
    guServOpts.uStats.intStatsGraphStartTime = udStatsGraphStart.value
    guServOpts.uStats.intStatsGraphStopTime = udStatsGraphStop.value
    
    guServOpts.uSessionCleanup.intAlignDesktopIcons = chkSessClnAlignDesktopIcons.value
    guServOpts.uSessionCleanup.intDeleteDocHistory = chkSessClnDeleteDocHistory.value
    guServOpts.uSessionCleanup.intRestoreIEHomePage = chkSessClnRestoreIEHomePage.value
    guServOpts.uSessionCleanup.strBackgroundPath = txtSessClnRestoreBackgroundPath.Text
    guServOpts.uSessionCleanup.intRestoreBackground = IIf(guServOpts.uSessionCleanup.strBackgroundPath = "", vbUnchecked, chkSessClnRestoreBackground.value)
    guServOpts.uSessionCleanup.strIEHomePagePath = txtSessClnRestoreIEHomePagePath.Text
    guServOpts.uSessionCleanup.intRestoreIESettings = chkSessClnRestoreIESettings.value
    guServOpts.uSessionCleanup.intRestoreIMs = chkSessClnRestoreIMs.value
    
    guServOpts.uSessionBatch.intEnableStart = chkSessBatchStart.value
    guServOpts.uSessionBatch.strStartPath = txtSessBatchStart.Text
    guServOpts.uSessionBatch.intEnableEnd = chkSessBatchEnd.value
    guServOpts.uSessionBatch.strEndPath = txtSessBatchEnd.Text
    guServOpts.uSessionBatch.intWindowStyle = cmbSessBatchStyle.ItemData(cmbSessBatchStyle.ListIndex)
    
    If lstKilledApps.ListCount > 0 Then
        ReDim guServOpts.uKilledApps.strKilledApps(lstKilledApps.ListCount - 1)
        For i = 0 To lstKilledApps.ListCount - 1
            guServOpts.uKilledApps.strKilledApps(i) = lstKilledApps.List(i)
        Next i
    Else
        ReDim guServOpts.uKilledApps.strKilledApps(0)
    End If
    guServOpts.uKilledApps.intForbidIE = chkForbidIE.value
    guServOpts.uKilledApps.intForbidMozilla = chkForbidMozilla.value
    guServOpts.uKilledApps.intForbidNetscape = chkForbidNetscape.value
    guServOpts.uKilledApps.intForbidOE = chkForbidOE.value
    
    For i = optPostpaidSsMode.LBound To optPostpaidSsMode.UBound
        If optPostpaidSsMode(i).value Then
            guServOpts.uPostPaidSs.intMode = i
            Exit For
        End If
    Next i
    For i = optPostpaidSsModeRechSubMode.LBound To optPostpaidSsModeRechSubMode.UBound
        If optPostpaidSsModeRechSubMode(i).value Then
            guServOpts.uPostPaidSs.intSubMode = i
            Exit For
        End If
    Next i
    guServOpts.uPostPaidSs.dteEndTimeIncrement = dtpPostpaidSsEndIncr.value
    guServOpts.uPostPaidSs.dteBaseTimeIncrement = dtpPostpaidSsBaseIncr.value
    guServOpts.uPostPaidSs.dteMinBilledTime = dtpPostpaidSsMinTime.value
    guServOpts.uPostPaidSs.dteGraceTime = dtpPostpaidSsGraceTime.value
    
    ' Update server and workstations......
    Call subSetSystemSettings
    
    Call subInitCards
    Call subInitClients
    Call subInitWorkstations
    Call subInitSessions
    
    Call subInitFrmMain
    
    If gblnDBOK Then
        Call subInitServer
        Call subInitCnxTesting
        For Each uWorkstation In guWorkstations
            Call subWorkstationInit(uWorkstation, False)
            DoEvents
        Next uWorkstation
    End If
    
    frmMain.ucWorkstations(1).Active = True
    frmMain.shpWksnConsoleBackground.BackColor = vb3DFace
    frmMain.fraWksnControls.Visible = True
    frmMain.fraWksnConsoleControls.Visible = False

    If mblnWorkstationLockDownChanged Then
        For Each uWorkstation In guWorkstations
            Call subWorkstationLockdown(uWorkstation)
        Next uWorkstation
        If guWorkstations.Count > 0 Then
            If funMsg(WorkstationLockdownConfigHasChanged) = vbYes Then
                Call Load(frmWksnShutdown)
                frmWksnShutdown.optWho(1).value = True
                frmWksnShutdown.optShutdown(2).value = True
                Call frmWksnShutdown.Show(vbModal, Me)
            End If
        End If
    End If
    
    If mblnDBPasswdChanged Then
        guServOpts.uDB.intDBBackup = vbUnchecked
        guServOpts.uDB.intDBCompact = vbUnchecked
        If funMsg(AskEndAppOnDBPasswdChange) = vbYes Then Call Unload(frmMain)
    End If
    If mblnDBPathChanged Then
        guServOpts.uDB.intDBBackup = vbUnchecked
        guServOpts.uDB.intDBCompact = vbUnchecked
        If funMsg(AskEndAppOnDBPathChange) = vbYes Then Call Unload(frmMain)
    End If
    
    For Each frmTemp In Forms
        Call frmTemp.subLocalizeInterface
    Next frmTemp
    
    Call Unload(Me)
    Exit Sub
    
cmdSave_ClickError:
    Call funErrorManager("frmOptions.cmdSave_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdServiceAdd_Click()
Dim uNewService As New clsService
On Error GoTo cmdServiceAdd_ClickError

    If txtServiceLabel.Text = "" Then
        Call funMsg(ServiceAddNoLabel)
        Call funSetFocus(txtServiceLabel)
    Else
        uNewService.Label = txtServiceLabel.Text
        uNewService.Price = CSng(IIf(txtServicePrice.Text = "", 0, Replace(txtServicePrice.Text, " ", "")))
        uNewService.StockMode = IIf(optServiceStockMode(0).value = True, enuServiceStockMode.Unmanaged, enuServiceStockMode.Managed)
        uNewService.Stock = CLng(IIf(txtServiceStock.Text = "", 0, txtServiceStock.Text))
        uNewService.StockMinAlert = CLng(IIf(txtServiceStockAlert.Text = "", 0, txtServiceStockAlert.Text))
        Call subDBAddService(uNewService)
        Call subFillServices
        Call subInitSelectedService
        Call subFillLocalServices
        
        cmdServiceNew.Visible = True
        cmdServiceAdd.Visible = False
        cmdServiceAddCancel.Visible = False
        cmdServiceDelete.Visible = False
        cmdServiceModify.Visible = False
        lbl(1).Caption = funLoadResString(10348)
        Call funSetFocus(fxgServices)
    End If
    Exit Sub
    
cmdServiceAdd_ClickError:
    Call funErrorManager("frmOptions.cmdServiceAdd_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdServiceAddCancel_Click()
On Error GoTo cmdServiceAddCancel_ClickError

    cmdServiceNew.Visible = True
    cmdServiceAdd.Visible = False
    cmdServiceAddCancel.Visible = False
    cmdServiceDelete.Visible = False
    cmdServiceModify.Visible = False
    
    Call subInitSelectedService
    
    fxgServices.enabled = True
    lbl(1).Caption = funLoadResString(10348)
    Call funSetFocus(fxgServices)
    Exit Sub
    
cmdServiceAddCancel_ClickError:
    Call funErrorManager("frmOptions.cmdServiceAddCancel_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdServiceDelete_Click()
Dim uCurrentService As New clsService
On Error GoTo cmdServiceDelete_ClickError

    If funMsg(enuMsg.ServiceDeleteConfirm) = vbYes Then
        Set uCurrentService = guServices(mlngSelectedServiceIndex)
        Call subDBDeleteService(uCurrentService)
        Call subFillServices
        Call subInitSelectedService
        Call subFillLocalServices
        
        cmdServiceNew.Visible = True
        cmdServiceAdd.Visible = False
        cmdServiceAddCancel.Visible = False
        cmdServiceDelete.Visible = False
        cmdServiceModify.Visible = False
        If fxgServices.enabled Then Call funSetFocus(fxgServices)
    End If
    Exit Sub
    
cmdServiceDelete_ClickError:
    Call funErrorManager("frmOptions.cmdServiceDelete_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdServiceModify_Click()
Dim uCurrentService As New clsService
On Error GoTo cmdServiceModify_ClickError

    cmdServiceNew.Visible = True
    cmdServiceAdd.Visible = False
    cmdServiceAddCancel.Visible = False
    cmdServiceDelete.Visible = True
    cmdServiceModify.Visible = True

    Set uCurrentService = guServices(mlngSelectedServiceIndex)
    
    uCurrentService.Label = txtServiceLabel.Text
    uCurrentService.Price = CSng(IIf(txtServicePrice.Text = "", 0, Replace(txtServicePrice.Text, " ", "")))
    uCurrentService.StockMode = IIf(optServiceStockMode(0).value = True, enuServiceStockMode.Unmanaged, enuServiceStockMode.Managed)
    uCurrentService.Stock = CLng(IIf(txtServiceStock.Text = "", 0, txtServiceStock.Text))
    uCurrentService.StockMinAlert = CLng(IIf(txtServiceStockAlert.Text = "", 0, txtServiceStockAlert.Text))
    Call subDBUpdateService(uCurrentService)
    Call subFillServices
    Call subFillLocalServices
    Exit Sub
    
cmdServiceModify_ClickError:
    Call funErrorManager("frmOptions.cmdServiceModify_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdServiceNew_Click()
On Error GoTo cmdServiceNew_ClickError

    cmdServiceNew.Visible = False
    cmdServiceAdd.Visible = True
    cmdServiceAddCancel.Visible = True
    cmdServiceDelete.Visible = False
    cmdServiceModify.Visible = False
    
    Call subInitSelectedService
    txtServiceLabel.Locked = False
    txtServicePrice.Locked = False
    optServiceStockMode(0).enabled = True
    optServiceStockMode(1).enabled = True
    txtServiceStock.Locked = False
    txtServiceStockAlert.Locked = False
    optServiceStockMode(0).value = True
    
    Call funSetFocus(txtServiceLabel)
    Exit Sub
    
cmdServiceNew_ClickError:
    Call funErrorManager("frmOptions.cmdServiceNew_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdSessBachEndOpen_Click()
On Error GoTo cmdSessBachEndOpen_ClickError

    Call Shell("notepad.exe " & txtSessBatchEnd.Text, vbNormalFocus)
    Exit Sub

cmdSessBachEndOpen_ClickError:
    Call funErrorManager("frmOptions.cmdSessBachEndOpen_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdSessBachEndPath_Click()
On Error GoTo cmdSessBachEndPath_ClickError
    
    Call subCdlOpenFile(cdlPath, txtSessBatchEnd, guServOpts.uSessionBatch.strEndPath, funGetFolderPath(CSIDL_DESKTOP), funLoadResString(10869), funLoadResString(10871))
    Exit Sub

cmdSessBachEndPath_ClickError:
    Call funErrorManager("frmOptions.cmdSessBachEndPath_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdSessBachStartOpen_Click()
On Error GoTo cmdSessBachStartOpen_ClickError

    Call Shell("notepad.exe " & txtSessBatchStart.Text, vbNormalFocus)
    Exit Sub

cmdSessBachStartOpen_ClickError:
    Call funErrorManager("frmOptions.cmdSessBachStartOpen_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdSessBachStartPath_Click()
On Error GoTo cmdSessBachStartPath_ClickError
    
    Call subCdlOpenFile(cdlPath, txtSessBatchStart, guServOpts.uSessionBatch.strStartPath, funGetFolderPath(CSIDL_DESKTOP), funLoadResString(10869), funLoadResString(10870))
    Exit Sub

cmdSessBachStartPath_ClickError:
    Call funErrorManager("frmOptions.cmdSessBachStartPath_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdShellBlocked_Click()
On Error GoTo cmdShellBlocked_ClickError

    Call funStartDoc(txtPathBlocked.Text)
    Exit Sub

cmdShellBlocked_ClickError:
    Call funErrorManager("frmOptions.cmdShellBlocked_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdShellIntro_Click()
On Error GoTo cmdShellIntro_ClickError

    Call funStartDoc(txtPathIntro.Text)
    Exit Sub

cmdShellIntro_ClickError:
    Call funErrorManager("frmOptions.cmdShellIntro_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdShellLogo_Click()
On Error GoTo cmdShellLogo_ClickError

    Call funStartDoc(txtPathLogo.Text)
    Exit Sub

cmdShellLogo_ClickError:
    Call funErrorManager("frmOptions.cmdShellLogo_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdShellNoCredit_Click()
On Error GoTo cmdShellNoCredit_ClickError

    Call funStartDoc(txtPathNoCredit.Text)
    Exit Sub

cmdShellNoCredit_ClickError:
    Call funErrorManager("frmOptions.cmdShellNoCredit_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdShellSessionEnd_Click()
On Error GoTo cmdShellSessionEnd_ClickError

    Call funStartDoc(txtPathSessionEnd.Text)
    Exit Sub

cmdShellSessionEnd_ClickError:
    Call funErrorManager("frmOptions.cmdShellSessionEnd_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdShellSessionPaused_Click()
On Error GoTo cmdShellSessionPaused_ClickError

    Call funStartDoc(txtPathSessionPaused.Text)
    Exit Sub

cmdShellSessionPaused_ClickError:
    Call funErrorManager("frmOptions.cmdShellSessionPaused_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdStatsTimeAdd_Click()
Dim i As Integer
On Error GoTo cmdStatsTimeAdd_ClickError

    For i = 0 To lstStatsTimes.ListCount
        If lstStatsTimes.List(i) = funFormatTime(dtpStatsTime.value) Then Exit Sub
    Next i
    Call lstStatsTimes.AddItem(funFormatTime(dtpStatsTime.value))
    mblnSettingsChanged = True
    Exit Sub
    
cmdStatsTimeAdd_ClickError:
    Call funErrorManager("frmOptions.cmdStatsTimeAdd_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdStatsTimeDelete_Click()
On Error GoTo cmdStatsTimeDelete_ClickError

    Call lstStatsTimes.RemoveItem(lstStatsTimes.ListIndex)
    mblnSettingsChanged = True
    cmdStatsTimeDelete.enabled = False
    Exit Sub
    
cmdStatsTimeDelete_ClickError:
    Call funErrorManager("frmOptions.cmdStatsTimeDelete_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdUnlockOptions_Click()
On Error Resume Next

    If txtLockedPassword.Text = guServOpts.uConfig.strPasswd Then
        Call subEnableAllControls
        fraAdminPassword.Visible = False
        lblUnlocked.Visible = True
        lblLocked.Visible = False
        cmdSave.Default = True
        Call funSetFocus(tbsTabs)
    Else
        Call funMsg(enuMsg.BadPasswd)
        txtLockedPassword.Text = ""
        Call funSetFocus(txtLockedPassword)
    End If
    Exit Sub
End Sub


'*******************************************************************************
Private Sub cmdWksnLockDrives_Click()

    Call Load(frmOptionsWksnLockDriveSelect)
    Call frmOptionsWksnLockDriveSelect.Show(vbModal)
    cmbWksnLockDrives.List(1) = funLoadResString(10903, funWksnLockDrivesStr(guServOpts.uWorkstations.uLocks.lngLockedDrivesTmp))
    If Not mblnSettingsChanged Then mblnSettingsChanged = guServOpts.uWorkstations.uLocks.lngLockedDrivesTmp <> guServOpts.uWorkstations.uLocks.lngLockedDrives
    If Not mblnWorkstationLockDownChanged Then mblnWorkstationLockDownChanged = guServOpts.uWorkstations.uLocks.lngLockedDrivesTmp <> guServOpts.uWorkstations.uLocks.lngLockedDrives
End Sub


'*******************************************************************************
Private Sub dtpCrdBonusTime_Change()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub dtpPostpaidSsBaseIncr_Change()
On Error Resume Next

    mblnSettingsChanged = True
    If dtpPostpaidSsBaseIncr.value > mdtePostPaidRechargeMaxBaseIncr Then
        dtpPostpaidSsBaseIncr.value = mdtePostPaidRechargeMaxBaseIncr
    ElseIf dtpPostpaidSsBaseIncr.value < MIN_PPSS_TIME_INCR Then
        dtpPostpaidSsBaseIncr.value = MIN_PPSS_TIME_INCR
    End If
    If dtpPostpaidSsMinTime.value < dtpPostpaidSsBaseIncr.value Then
        dtpPostpaidSsMinTime.value = dtpPostpaidSsBaseIncr.value
    End If
End Sub


'*******************************************************************************
Private Sub dtpPostpaidSsEndIncr_Change()
On Error Resume Next

    mblnSettingsChanged = True
    Call subInitPostpaidSsEndIncrLbl
End Sub


'*******************************************************************************
Private Sub dtpPostpaidSsGraceTime_Change()
On Error Resume Next

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub dtpPostpaidSsMinTime_Change()
On Error Resume Next

    mblnSettingsChanged = True
    If dtpPostpaidSsMinTime.value > mdtePostPaidRechargeMaxMinTime Then
        dtpPostpaidSsMinTime.value = mdtePostPaidRechargeMaxMinTime
    ElseIf dtpPostpaidSsMinTime.value < dtpPostpaidSsBaseIncr.value Then
        dtpPostpaidSsMinTime.value = dtpPostpaidSsBaseIncr.value
    ElseIf dtpPostpaidSsMinTime.value < MIN_PPSS_TIME_INCR Then
        dtpPostpaidSsMinTime.value = MIN_PPSS_TIME_INCR
    End If
End Sub

'*******************************************************************************
Private Sub Form_Activate()
On Error Resume Next
    
    Call funSetFocus(tbsTabs)
End Sub


'*******************************************************************************
Private Sub Form_Load()
Dim i As Integer
On Error GoTo Form_LoadError

    gblnfrmOptions = True
    Call subSetDefaultFont(Me)
    Call subLocalizeInterface
    mstrLangPacks = funFindLangPacks
    muAdapters = funGetIPAdapters
    
    Call tbsTabs.ZOrder(1)
    
    Call cmbBindInterface.Clear
    If muAdapters.Count > 0 Then
        For i = 0 To muAdapters.Count - 1
            If muAdapters.Item(i).AdapterTypeID = MIB_IF_TYPE_ETHERNET Then
                Call cmbBindInterface.AddItem(muAdapters.Item(i).IpAddress(0) & " - " & muAdapters.Item(i).Description)
                cmbBindInterface.ItemData(cmbBindInterface.NewIndex) = i
                If guServOpts.uServer.strServerMac = muAdapters.Item(i).MacAdress Then
                    cmbBindInterface.ListIndex = i
                    cmbBindInterface.ToolTipText = funLoadResString(10884) & muAdapters.Item(i).IpAddress(0) & funLoadResString(10885) & muAdapters.Item(i).Description
                End If
            End If
        Next i
        If cmbBindInterface.ListIndex = -1 Then cmbBindInterface.ListIndex = 0
    End If
    
    udNberWorkstations.Max = guServOpts.uWorkstations.lngMaxNumWorkstations
    udNberWorkstations.value = guServOpts.uWorkstations.lngNumWorkstations
    udFollowInterval.value = guServOpts.uWorkstations.intFollowInterval
    dtpNumDaysInMonth.value = DateAdd("d", guServOpts.uConfig.intNumDaysInMonth - 1, dtpNumDaysInMonth.MinDate)
    udDisconConRatio.value = guServOpts.uTimeRatios.intDisconCon
    udOffPeak.value = guServOpts.uTimeRatios.intOffPeak
    dtpOffPeakHourStart.value = guServOpts.uTimeRatios.dteOffPeakStart
    dtpOffPeakHourStop.value = guServOpts.uTimeRatios.dteOffPeakStop
        
    chkMsgCard.value = guServOpts.uWorkstations.uWksnMsg.intMsgCard
    chkMsgSS.value = guServOpts.uWorkstations.uWksnMsg.intMsgSelfService
    udMsg1CardMinutes.value = guServOpts.uWorkstations.uWksnMsg.intMsgCardAlert1
    udMsg2CardMinutes.value = guServOpts.uWorkstations.uWksnMsg.intMsgCardAlert2
    txtMsgSSAmount.Text = guServOpts.uWorkstations.uWksnMsg.sngMsgSelfServiceBalance
    
    txtDBPath.Text = guServOpts.uDB.strDBPath
    txtDBPath.ToolTipText = guServOpts.uDB.strDBPath
    txtDBPath.SelStart = Len(txtDBPath.Text)
    txtDBBackupPath.Text = guServOpts.uDB.strDBBackupPath
    txtDBBackupPath.ToolTipText = guServOpts.uDB.strDBBackupPath
    txtDBBackupPath.SelStart = Len(txtDBBackupPath.Text)
    txtDBPasswd.Text = guServOpts.uDB.strDBPasswd
    chkBackupDB.value = guServOpts.uDB.intDBBackup
    chkBackupDBSendToWksn.value = guServOpts.uDB.intDBBackupSendToWorkstation
    chkCompactDB.value = guServOpts.uDB.intDBCompact
    
    cdlPath.filename = guServOpts.uDB.strDBPath
    cdlPath.InitDir = guServOpts.uDB.strDBPath
    cdlPath.flags = cdlOFNHideReadOnly + cdlOFNFileMustExist

    optWksnCnxAuto.value = guServOpts.uServer.blnAutoConnect
    optWksnCnxManual.value = Not guServOpts.uServer.blnAutoConnect
    txtPortNum.Text = guServOpts.uServer.lngPort
    txtMonSymbol.Text = guServOpts.uConfig.strMonSymbol
    txtMonFormat.Text = guServOpts.uConfig.strMonFormat
    
    chkUsageSelfServicePostPaid.value = guServOpts.uWorkstations.intActivateSelfServicePostPaid
    optUsageDefault(enuSessionType.Selfservice).enabled = chkUsageSelfServicePostPaid.value = vbChecked
    chkUsageSelfServicePrePaid.value = guServOpts.uWorkstations.intActivateSelfServicePrePaid
    optUsageDefault(enuSessionType.SelfServicePrePaid).enabled = chkUsageSelfServicePrePaid.value = vbChecked
    chkUsageCards.value = guServOpts.uWorkstations.intActivateCards
    optUsageDefault(enuSessionType.Card).enabled = chkUsageCards.value = vbChecked
    chkUsageClients.value = guServOpts.uWorkstations.intActivateClients
    optUsageDefault(enuSessionType.Client).enabled = chkUsageClients.value = vbChecked
    If guServOpts.uWorkstations.intUsageDefault <> enuSessionType.Unknown Then optUsageDefault(guServOpts.uWorkstations.intUsageDefault).value = True
    
    chkStartForceWksnsConnected.value = guServOpts.uWorkstations.intStartForceWorkstationsConnected
    chkHideCardNum.value = guServOpts.uWorkstations.intHideCardNum
    chkEnableLoggingURLs.value = guServOpts.uWorkstations.intLogURLs
    chkEnablePause.value = guServOpts.uWorkstations.intEnablePause
    
    Call cmbWksnDefaultState.Clear
    Call cmbWksnDefaultState.AddItem(funLoadResString(10911), 0)
    cmbWksnDefaultState.ItemData(0) = enuWksnState.Ready
    Call cmbWksnDefaultState.AddItem(funLoadResString(10912), 1)
    cmbWksnDefaultState.ItemData(1) = enuWksnState.Blocked_Deactivated
    If guServOpts.uWorkstations.intDefaultState = enuWksnState.Ready Then
        cmbWksnDefaultState.ListIndex = 0
    Else
        cmbWksnDefaultState.ListIndex = 1
    End If
    
    chkWksnLockdown.value = guServOpts.uWorkstations.intModifyLockdown
    chkWksnLock(0).value = guServOpts.uWorkstations.uLocks.intNoClose
    chkWksnLock(1).value = guServOpts.uWorkstations.uLocks.intNoAdminTools
    chkWksnLock(2).value = guServOpts.uWorkstations.uLocks.intNoCPL
    chkWksnLock(3).value = guServOpts.uWorkstations.uLocks.intLockDesktop
    chkWksnLock(4).value = guServOpts.uWorkstations.uLocks.intNoFileSystemAccess
    chkWksnLock(5).value = guServOpts.uWorkstations.uLocks.intLockIE
    chkWksnLock(6).value = guServOpts.uWorkstations.uLocks.intNoCommonGroup
        
    guServOpts.uWorkstations.uLocks.lngLockedDrivesTmp = guServOpts.uWorkstations.uLocks.lngLockedDrives
    Call cmbWksnLockDrives.Clear
    Call cmbWksnLockDrives.AddItem(funLoadResString(10902), 0)
    Call cmbWksnLockDrives.AddItem(funLoadResString(10903, funWksnLockDrivesStr(guServOpts.uWorkstations.uLocks.lngLockedDrivesTmp)), 1)
    cmbWksnLockDrives.ListIndex = guServOpts.uWorkstations.uLocks.intNoFileSystemAccessType
            
    udCrdPurgeDelta.value = guServOpts.uCards.intCardPurgeDelta
    chkCrdAutoPurge.value = guServOpts.uCards.intAutoCardPurge
    udCrdListFillInterval.value = guServOpts.uCards.intCardListUpdateInterval
    
    chkCrdBonus.value = guServOpts.uCards.intBonusActive
    txtCrdBonusThreshold.Text = guServOpts.uCards.sngBonusThreshold
    dtpCrdBonusTime.value = guServOpts.uCards.dteBonusTime
    
    chkCliEnableExpiry.value = guServOpts.uClients.intEnableExpiry
    chkCliBonus.value = guServOpts.uClients.intBonusActive
    txtCliBonusThreshold.Text = guServOpts.uClients.sngBonusThreshold
    dtpCliBonusTime.value = guServOpts.uClients.dteBonusTime
    
    udCliPurgeDelta.value = guServOpts.uClients.intClientPurgeDelta
    chkCliAutoPurge.value = guServOpts.uClients.intAutoClientPurge
    udCliListFillInterval.value = guServOpts.uClients.intClientListUpdateInterval
        
    chkCnxTesting.value = guServOpts.uCnxTest.intCnxTesting
    chkWksnNotification.value = guServOpts.uCnxTest.intWorkstationNotification
    txtTestServerName.Text = guServOpts.uCnxTest.strTestServerName
    txtTestServerPort.Text = guServOpts.uCnxTest.strTestServerPort
        
    For i = 0 To UBound(mstrLangPacks, 2)
        If mstrLangPacks(5, i) = True Then
            Call cmbLanguage.AddItem(mstrLangPacks(4, i))
            cmbLanguage.ItemData(cmbLanguage.NewIndex) = i
        End If
    Next i
    For i = 0 To cmbLanguage.ListCount - 1
        If guServOpts.uConfig.strInterfaceLang = mstrLangPacks(0, cmbLanguage.ItemData(i)) Then cmbLanguage.ListIndex = i
    Next i
    
    If cmbLanguage.ListIndex = -1 Then cmbLanguage.ListIndex = 0
    
    With fxgRecharges
        .top = shpRechargeGrid.top + Screen.TwipsPerPixelY
        .left = shpRechargeGrid.left + Screen.TwipsPerPixelX
        .height = shpRechargeGrid.height - Screen.TwipsPerPixelY * 2
        .width = shpRechargeGrid.width - Screen.TwipsPerPixelX * 2
        .Cols = 6
        .Rows = 18
        .Row = 0
        .ColWidth(0) = 750
        .ColAlignment(0) = flexAlignCenterTop
        .ColWidth(1) = 700
        .ColAlignment(1) = flexAlignRightTop
        .ColWidth(2) = 1000
        .ColAlignment(2) = flexAlignLeftTop
        .ColWidth(3) = 1600
        .ColAlignment(3) = flexAlignLeftTop
        .ColWidth(4) = 4000
        .ColAlignment(4) = flexAlignLeftTop
        .enabled = gblnDBOK
        Call subFillRecharges
    End With
    With cmbRechargeValidityType
'        Call .AddItem(funLoadResString(10995))
'        .ItemData(.NewIndex) = enuRechargeValidityType.Hours
        Call .AddItem(funLoadResString(10996))
        .ItemData(.NewIndex) = enuRechargeValidityType.Days
        Call .AddItem(funLoadResString(10997))
        .ItemData(.NewIndex) = enuRechargeValidityType.Weeks
        Call .AddItem(funLoadResString(10998))
        .ItemData(.NewIndex) = enuRechargeValidityType.Months
        Call .AddItem(funLoadResString(10999))
        .ItemData(.NewIndex) = enuRechargeValidityType.Years
        .ListIndex = 2
    End With
    
    
    mlngSelectedServiceIndex = 0
    With fxgServices
        .top = shpServiceGrid.top + Screen.TwipsPerPixelY
        .left = shpServiceGrid.left + Screen.TwipsPerPixelX
        .height = shpServiceGrid.height - Screen.TwipsPerPixelY * 2
        .width = shpServiceGrid.width - Screen.TwipsPerPixelX * 2
        .Cols = 5
        .Rows = 18
        .Row = 0
        .ColWidth(0) = 2500
        .ColAlignment(0) = flexAlignLeftTop
        .ColWidth(1) = 1000
        .ColAlignment(1) = flexAlignRightTop
        .ColWidth(2) = 600
        .ColAlignment(2) = flexAlignLeftTop
        .ColWidth(3) = 9000
        .ColAlignment(3) = flexAlignLeftTop
        .enabled = gblnDBOK
        Call subFillLocalServices
    End With
    
    cmdRechargeAdd.Visible = gblnDBOK
    cmdRechargeDelete.Visible = gblnDBOK
    cmdRechargeModify.Visible = gblnDBOK
    
    cmdServiceAdd.Visible = False
    cmdServiceAddCancel.Visible = False
    cmdServiceDelete.Visible = False
    cmdServiceModify.Visible = False
    cmdServiceNew.Visible = gblnDBOK
    
    chkSessBatchStart.value = guServOpts.uSessionBatch.intEnableStart
    txtSessBatchStart.Text = guServOpts.uSessionBatch.strStartPath
    txtSessBatchStart.ToolTipText = guServOpts.uSessionBatch.strStartPath
    txtSessBatchStart.SelStart = Len(txtSessBatchStart.Text)
    chkSessBatchEnd.value = guServOpts.uSessionBatch.intEnableEnd
    txtSessBatchEnd.Text = guServOpts.uSessionBatch.strEndPath
    txtSessBatchEnd.ToolTipText = guServOpts.uSessionBatch.strEndPath
    txtSessBatchEnd.SelStart = Len(txtSessBatchEnd.Text)
    
    cmbSessBatchStyle.AddItem (funLoadResString(10877))
    cmbSessBatchStyle.ItemData(cmbSessBatchStyle.NewIndex) = vbNormalFocus
    cmbSessBatchStyle.AddItem (funLoadResString(10878))
    cmbSessBatchStyle.ItemData(cmbSessBatchStyle.NewIndex) = vbMinimizedNoFocus
    cmbSessBatchStyle.AddItem (funLoadResString(10879))
    cmbSessBatchStyle.ItemData(cmbSessBatchStyle.NewIndex) = vbHide
    For i = 0 To cmbSessBatchStyle.ListCount - 1
        If cmbSessBatchStyle.ItemData(i) = guServOpts.uSessionBatch.intWindowStyle Then cmbSessBatchStyle.ListIndex = i
    Next i
    If cmbSessBatchStyle.ListIndex = -1 Then cmbSessBatchStyle.ListIndex = 0
    
    udStatsGraphStart.value = guServOpts.uStats.intStatsGraphStartTime
    udStatsGraphStop.value = guServOpts.uStats.intStatsGraphStopTime
    
    Call lstStatsTimes.Clear
    i = funArraySize(guServOpts.uStats.strStatsTimeInt)
    If i > 0 Then
        For i = 0 To funArraySize(guServOpts.uStats.strStatsTimeInt) - 1
            Call lstStatsTimes.AddItem(funFormatTime(funDateTimefromNet(guServOpts.uStats.strStatsTimeInt(i))))
        Next i
    End If
    
    Call lstCardTypes.Clear
    Call lstCardTypes.AddItem(funLoadResString(10526))
    lstCardTypes.ItemData(0) = 1
    i = funArraySize(guServOpts.uCards.strCardTypes)
    If i > 0 Then
        For i = 0 To funArraySize(guServOpts.uCards.strCardTypes) - 1
            Call lstCardTypes.AddItem(guServOpts.uCards.strCardTypes(i))
        Next i
    End If
    
    chkWksnMsgLogoUpload.value = guServOpts.uWorkstations.intModifyMsgLogoFiles
    txtPathBlocked.Text = guServOpts.uWorkstations.uPaths.strBlockedPath
    txtPathBlocked.ToolTipText = guServOpts.uWorkstations.uPaths.strBlockedPath
    txtPathBlocked.SelStart = Len(txtPathBlocked.Text)
    txtPathIntro.Text = guServOpts.uWorkstations.uPaths.strIntroPath
    txtPathIntro.ToolTipText = guServOpts.uWorkstations.uPaths.strIntroPath
    txtPathIntro.SelStart = Len(txtPathIntro.Text)
    txtPathLogo.Text = guServOpts.uWorkstations.uPaths.strLogoPath
    txtPathLogo.ToolTipText = guServOpts.uWorkstations.uPaths.strLogoPath
    txtPathLogo.SelStart = Len(txtPathLogo.Text)
    txtPathNoCredit.Text = guServOpts.uWorkstations.uPaths.strNoCreditPath
    txtPathNoCredit.ToolTipText = guServOpts.uWorkstations.uPaths.strNoCreditPath
    txtPathNoCredit.SelStart = Len(txtPathNoCredit.Text)
    txtPathSessionEnd.Text = guServOpts.uWorkstations.uPaths.strSessionEndPath
    txtPathSessionEnd.ToolTipText = guServOpts.uWorkstations.uPaths.strSessionEndPath
    txtPathSessionEnd.SelStart = Len(txtPathSessionEnd.Text)
    txtPathSessionPaused.Text = guServOpts.uWorkstations.uPaths.strSessionPausedPath
    txtPathSessionPaused.ToolTipText = guServOpts.uWorkstations.uPaths.strSessionPausedPath
    txtPathSessionPaused.SelStart = Len(txtPathSessionPaused.Text)
    
    chkSessClnAlignDesktopIcons.value = guServOpts.uSessionCleanup.intAlignDesktopIcons
    chkSessClnDeleteDocHistory.value = guServOpts.uSessionCleanup.intDeleteDocHistory
    chkSessClnRestoreBackground.value = guServOpts.uSessionCleanup.intRestoreBackground
    chkSessClnRestoreIEHomePage.value = guServOpts.uSessionCleanup.intRestoreIEHomePage
    chkSessClnRestoreIESettings.value = guServOpts.uSessionCleanup.intRestoreIESettings
    chkSessClnRestoreIMs.value = guServOpts.uSessionCleanup.intRestoreIMs
    txtSessClnRestoreBackgroundPath.Text = guServOpts.uSessionCleanup.strBackgroundPath
    txtSessClnRestoreBackgroundPath.ToolTipText = guServOpts.uSessionCleanup.strBackgroundPath
    txtSessClnRestoreBackgroundPath.SelStart = Len(txtSessClnRestoreBackgroundPath.Text)
    txtSessClnRestoreIEHomePagePath.Text = guServOpts.uSessionCleanup.strIEHomePagePath
    txtSessClnRestoreIEHomePagePath.ToolTipText = guServOpts.uSessionCleanup.strIEHomePagePath
    txtSessClnRestoreIEHomePagePath.SelStart = Len(txtSessClnRestoreIEHomePagePath.Text)
    
    i = funArraySize(guServOpts.uKilledApps.strKilledApps)
    If i > 0 Then
        For i = 0 To funArraySize(guServOpts.uKilledApps.strKilledApps) - 1
            If guServOpts.uKilledApps.strKilledApps(i) <> "" Then Call lstKilledApps.AddItem(guServOpts.uKilledApps.strKilledApps(i))
        Next i
    End If
    
    chkForbidIE.value = guServOpts.uKilledApps.intForbidIE
    chkForbidMozilla.value = guServOpts.uKilledApps.intForbidMozilla
    chkForbidNetscape.value = guServOpts.uKilledApps.intForbidNetscape
    chkForbidOE.value = guServOpts.uKilledApps.intForbidOE
    
    Call cmbURLtimespan.Clear
    Call cmbURLtimespan.AddItem("1 " & funLoadResString(10740))
    cmbURLtimespan.ItemData(cmbURLtimespan.NewIndex) = 1
    Call cmbURLtimespan.AddItem("2 " & funLoadResString(10741))
    cmbURLtimespan.ItemData(cmbURLtimespan.NewIndex) = 2
    Call cmbURLtimespan.AddItem("6 " & funLoadResString(10741))
    cmbURLtimespan.ItemData(cmbURLtimespan.NewIndex) = 6
    Call cmbURLtimespan.AddItem("1 " & funLoadResString(10742))
    cmbURLtimespan.ItemData(cmbURLtimespan.NewIndex) = 12
    Call cmbURLtimespan.AddItem(funLoadResString(10743))
    cmbURLtimespan.ItemData(cmbURLtimespan.NewIndex) = 0
    cmbURLtimespan.ListIndex = 1
    For i = 0 To cmbURLtimespan.ListCount
        If cmbURLtimespan.ItemData(i) = guServOpts.uWorkstations.intLoggedURLTimeSpan Then
            cmbURLtimespan.ListIndex = i
            Exit For
        End If
    Next i
    If cmbURLtimespan.ListIndex = -1 Then cmbURLtimespan.ListIndex = 0
    
    Call cmbWksnDefaultMode.Clear
    Call cmbWksnDefaultMode.AddItem(funLoadResString(10969))
    cmbWksnDefaultMode.ItemData(cmbWksnDefaultMode.NewIndex) = enuWksnMode.Connected
    Call cmbWksnDefaultMode.AddItem(funLoadResString(10970))
    cmbWksnDefaultMode.ItemData(cmbWksnDefaultMode.NewIndex) = enuWksnMode.Disconnected
    For i = 0 To cmbWksnDefaultMode.ListCount
        If cmbWksnDefaultMode.ItemData(i) = guServOpts.uWorkstations.intDefaultMode Then
            cmbWksnDefaultMode.ListIndex = i
            Exit For
        End If
    Next i
    If cmbWksnDefaultMode.ListIndex = -1 Then cmbWksnDefaultMode.ListIndex = 0
    
    optPostpaidSsMode(guServOpts.uPostPaidSs.intMode).value = True
    optPostpaidSsModeRechSubMode(guServOpts.uPostPaidSs.intSubMode).value = True
    dtpPostpaidSsEndIncr.MinDate = MIN_PPSS_TIME_INCR
    dtpPostpaidSsEndIncr.value = guServOpts.uPostPaidSs.dteEndTimeIncrement
    dtpPostpaidSsBaseIncr.MinDate = MIN_PPSS_TIME_INCR
    dtpPostpaidSsBaseIncr.value = guServOpts.uPostPaidSs.dteBaseTimeIncrement
    dtpPostpaidSsMinTime.MinDate = MIN_PPSS_TIME_INCR
    dtpPostpaidSsMinTime.value = guServOpts.uPostPaidSs.dteMinBilledTime
    dtpPostpaidSsGraceTime.MinDate = NULL_DATE
    dtpPostpaidSsGraceTime.value = guServOpts.uPostPaidSs.dteGraceTime
    dtpPostpaidSsGraceTime.MinDate = NULL_DATE
    
    'Cleanup and exit.
    mblnSettingsChanged = False
    mblnDBPasswdChanged = False
    mblnDBPathChanged = False
    mblnWorkstationLockDownChanged = False

    If guServOpts.uConfig.strPasswd = "" Then
        chkLocked.value = vbChecked
        Call cmdUnlockOptions_Click
    Else
        Call subDisableAllControls
    End If
    
    Exit Sub
    
Form_LoadError:
    Call funErrorManager("frmOptions.Form_LoadError", True)
End Sub


'*******************************************************************************
Private Sub subDisableAllControls()
Dim uTmpCtrl As Control
On Error Resume Next
    
    chkLocked.Picture = imlLocked.ListImages(1).Picture
    chkLocked.ToolTipText = funLoadResString(10523)
    txtLockedPassword.Text = ""
    For Each uTmpCtrl In frmOptions.Controls
        If uTmpCtrl.Name <> "tbsTabs" And uTmpCtrl.Name <> "chkLocked" And uTmpCtrl.Name <> "txtLockedPassword" And uTmpCtrl.Name <> "cmdUnlockOptions" And uTmpCtrl.Name <> "cmdCancel" And uTmpCtrl.Name <> "fraAdminPassword" And uTmpCtrl.Name <> "fraLocked" And uTmpCtrl.Name <> "lblLockedPassword" And uTmpCtrl.Name <> "lblLocked" Then uTmpCtrl.enabled = False
    Next uTmpCtrl
End Sub


'*******************************************************************************
Private Sub subEnableAllControls()
Dim uTmpCtrl As Control
Dim i As Integer
On Error Resume Next
    
    chkLocked.Picture = imlLocked.ListImages(2).Picture
    chkLocked.ToolTipText = funLoadResString(10525)
    For Each uTmpCtrl In frmOptions.Controls
        If uTmpCtrl.Name <> "optServiceStockMode" And Not (uTmpCtrl.Name = "optStatsGraphType" And uTmpCtrl.Index = 1) Then uTmpCtrl.enabled = True
    Next uTmpCtrl
    cmdStatsTimeDelete.enabled = False
    cmdCrdTypeDelete.enabled = False
    For i = 0 To chkWksnLock.Count - 1
        chkWksnLock(i).enabled = chkWksnLockdown.value = vbChecked
    Next i
    fxgServices.enabled = guServices.Count <> 0
    fxgRecharges.enabled = Not (grstRecharges.BOF = True And grstRecharges.EOF = True Or grstRecharges.RecordCount = 0)
    chkBackupDBSendToWksn.enabled = chkBackupDB.value = vbChecked
    txtDBBackupPath.enabled = chkBackupDB.value = vbChecked
    cmdDBBackupPath.enabled = chkBackupDB.value = vbChecked
    dtpCrdBonusTime.enabled = chkCrdBonus.value = vbChecked
    txtCrdBonusThreshold.enabled = chkCrdBonus.value = vbChecked
    lbl(67).enabled = chkCrdBonus.value = vbChecked
    lbl(68).enabled = chkCrdBonus.value = vbChecked
    dtpCliBonusTime.enabled = chkCliBonus.value = vbChecked
    txtCliBonusThreshold.enabled = chkCliBonus.value = vbChecked
    lbl(18).enabled = chkCliBonus.value = vbChecked
    lbl(19).enabled = chkCliBonus.value = vbChecked
    chkCliAutoPurge.enabled = chkCliEnableExpiry.value = vbChecked
    udCliPurgeDelta.enabled = chkCliEnableExpiry.value = vbChecked
    lblCliPurgeDelta.enabled = chkCliEnableExpiry.value = vbChecked
    lbl(15).enabled = chkCliEnableExpiry.value = vbChecked
    txtPortNum.enabled = optWksnCnxManual.value
    lbl(3).enabled = optWksnCnxManual.value
    cmbWksnLockDrives.enabled = chkWksnLock(4).value = vbChecked And chkWksnLockdown.value = vbChecked
    cmdWksnLockDrives.enabled = chkWksnLock(4).value = vbChecked And chkWksnLockdown.value = vbChecked And cmbWksnLockDrives.ListIndex = vbChecked
    cmbURLtimespan.enabled = chkEnableLoggingURLs.value = vbChecked
    
    lbl(70).enabled = chkWksnMsgLogoUpload.value = vbChecked
    lbl(71).enabled = chkWksnMsgLogoUpload.value = vbChecked
    lbl(72).enabled = chkWksnMsgLogoUpload.value = vbChecked
    lbl(73).enabled = chkWksnMsgLogoUpload.value = vbChecked
    lbl(74).enabled = chkWksnMsgLogoUpload.value = vbChecked
    lbl(75).enabled = chkWksnMsgLogoUpload.value = vbChecked
    txtPathIntro.enabled = chkWksnMsgLogoUpload.value = vbChecked
    txtPathBlocked.enabled = chkWksnMsgLogoUpload.value = vbChecked
    txtPathLogo.enabled = chkWksnMsgLogoUpload.value = vbChecked
    txtPathNoCredit.enabled = chkWksnMsgLogoUpload.value = vbChecked
    txtPathSessionEnd.enabled = chkWksnMsgLogoUpload.value = vbChecked
    txtPathSessionPaused.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdPathIntro.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdPathBlocked.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdPathLogo.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdPathNoCredit.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdPathSessionEnd.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdPathSessionPaused.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdShellIntro.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdShellBlocked.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdShellLogo.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdShellNoCredit.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdShellSessionEnd.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdShellSessionPaused.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdPrevIntro.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdPrevBlocked.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdPrevLogo.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdPrevNoCredit.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdPrevSessionEnd.enabled = chkWksnMsgLogoUpload.value = vbChecked
    cmdPrevSessionPaused.enabled = chkWksnMsgLogoUpload.value = vbChecked
    chkStartForceWksnsConnected.enabled = cmbWksnDefaultMode.ItemData(cmbWksnDefaultMode.ListIndex) = enuWksnMode.Connected
    
    optUsageDefault(enuSessionType.Selfservice).enabled = chkUsageSelfServicePostPaid.value = vbChecked
    optUsageDefault(enuSessionType.SelfServicePrePaid).enabled = chkUsageSelfServicePrePaid.value = vbChecked
    optUsageDefault(enuSessionType.Card).enabled = chkUsageCards.value = vbChecked
    optUsageDefault(enuSessionType.Client).enabled = chkUsageClients.value = vbChecked
    
    dtpPostpaidSsEndIncr.enabled = optPostpaidSsModeRechSubMode(enuPostpaidSsSubMode.RechSimple).value
    lbl(89).enabled = optPostpaidSsModeRechSubMode(enuPostpaidSsSubMode.RechSimple).value

End Sub


'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    gblnfrmOptions = False
End Sub


'*******************************************************************************
Private Sub fxgRecharges_Click()
On Error GoTo fxgRecharges_ClickError
    
    If mlngSelectedRechargeIndex = 0 Then Call fxgRecharges_RowColChange
    Exit Sub
    
fxgRecharges_ClickError:
    Call funErrorManager("frmOptions.fxgRecharges_ClickError", True)
End Sub


'*******************************************************************************
Private Sub fxgRecharges_DblClick()
On Error GoTo fxgRecharges_DblClickError
    
    With fxgRecharges
        If .MouseRow = 0 Then
            If mlngSelectedRechargeIndex > 0 Then
                .Row = mlngSelectedRechargeIndex
                .ColSel = .Cols - 1
                .CellForeColor = vbWindowText
                .CellBackColor = vbWindowBackground
            End If
            
            Select Case .MouseCol
            Case 0
                mlngSelectedRechargeOrder = enuRechargeOrder.orderByTime
            Case 1
                mlngSelectedRechargeOrder = enuRechargeOrder.orderByPrice
            Case 2
                mlngSelectedRechargeOrder = enuRechargeOrder.orderByValidity
            Case 3
                mlngSelectedRechargeOrder = enuRechargeOrder.orderByType
            Case 4
                mlngSelectedRechargeOrder = enuRechargeOrder.RechargesIndex
            End Select
            Call subFillRecharges
            .Row = 0
            mlngSelectedRechargeIndex = 0
            txtRechargeCaption.Text = ""
            txtRechargePrice.Text = ""
            dtpRechargeTime.value = NULL_DATE
            dtpRechargeValidity.value = dtpRechargeValidity.MinDate
            chkRechargeCli.value = vbUnchecked
            chkRechargeCrd.value = vbUnchecked
            chkRechargePostSs.value = vbUnchecked
            chkRechargePreSs.value = vbUnchecked
            chkRechargeTimeUnlimited.value = vbUnchecked
            cmbRechargeValidityType.ListIndex = 2
        End If
    End With
    Exit Sub
    
fxgRecharges_DblClickError:
    Call funErrorManager("frmOptions.fxgRecharges_DblClickError", True)
End Sub


'*******************************************************************************
Private Sub fxgRecharges_RowColChange()
Dim lngNewSelectedRowIndex As Long
Dim i As Integer
Dim strBuffer As String
On Error GoTo fxgRecharges_RowColChangeError
    
'    subDebugPrint fxgRecharges.Col
    With fxgRecharges
        lngNewSelectedRowIndex = .Row
        If lngNewSelectedRowIndex < 1 Then Exit Sub
        If mlngSelectedRechargeIndex > .Rows - 1 Then mlngSelectedRechargeIndex = .Rows - 1
        
        .Col = 0
        If mlngSelectedRechargeIndex > 0 Then
            .Row = mlngSelectedRechargeIndex
            .ColSel = .Cols - 1
            .CellForeColor = vbWindowText
            .CellBackColor = vbWindowBackground
        End If
        
        mlngSelectedRechargeIndex = lngNewSelectedRowIndex
        
        .Row = lngNewSelectedRowIndex
        .ColSel = .Cols - 1
        .CellForeColor = vbHighlightText
        .CellBackColor = vbHighlight
        
        If .TextMatrix(lngNewSelectedRowIndex, 0) = funLoadResString(11000) Then
            dtpRechargeTime.value = 0
            chkRechargeTimeUnlimited.value = vbChecked
        Else
            dtpRechargeTime.value = CDate(IIf(IsNull(.TextMatrix(lngNewSelectedRowIndex, 0)), 0, .TextMatrix(lngNewSelectedRowIndex, 0)))
            chkRechargeTimeUnlimited.value = vbUnchecked
        End If
        For i = 1 To Len(.TextMatrix(lngNewSelectedRowIndex, 1))
            If IsNumeric(Mid(.TextMatrix(lngNewSelectedRowIndex, 1), i, 1)) Or funGetDecimalSeparator = Mid(.TextMatrix(lngNewSelectedRowIndex, 1), i, 1) Then strBuffer = strBuffer & Mid(.TextMatrix(lngNewSelectedRowIndex, 1), i, 1)
        Next i
        txtRechargePrice.Text = strBuffer
        
        dtpRechargeValidity.value = IIf(.TextMatrix(lngNewSelectedRowIndex, 2) <> "0", DateAdd("yyyy", Trim(left(.TextMatrix(lngNewSelectedRowIndex, 2), InStr(.TextMatrix(lngNewSelectedRowIndex, 2), " "))), DateAdd("yyyy", -1, dtpRechargeValidity.MinDate)), dtpRechargeValidity.MinDate)
        
        txtRechargeCaption.Text = .TextMatrix(lngNewSelectedRowIndex, 4)
        
        chkRechargeCrd.value = IIf(InStr(.TextMatrix(lngNewSelectedRowIndex, 3), funLoadResString(10991)) > 0, vbChecked, vbUnchecked)
        chkRechargeCli.value = IIf(InStr(.TextMatrix(lngNewSelectedRowIndex, 3), funLoadResString(10992)) > 0, vbChecked, vbUnchecked)
        chkRechargePostSs.value = IIf(InStr(.TextMatrix(lngNewSelectedRowIndex, 3), funLoadResString(10994)) > 0, vbChecked, vbUnchecked)
        chkRechargePreSs.value = IIf(InStr(.TextMatrix(lngNewSelectedRowIndex, 3), funLoadResString(10993)) > 0, vbChecked, vbUnchecked)
        
        For i = 0 To cmbRechargeValidityType.ListCount - 1
            If InStr(.TextMatrix(lngNewSelectedRowIndex, 2), cmbRechargeValidityType.List(i)) <> 0 Then
                cmbRechargeValidityType.ListIndex = i
                Exit For
            End If
        Next i
    
    End With
    Exit Sub
    
fxgRecharges_RowColChangeError:
    Call funErrorManager("frmOptions.fxgRecharges_RowColChangeError", True)
End Sub


'*******************************************************************************
Private Sub fxgServices_Click()
On Error GoTo fxgServices_ClickError
    
    If mlngSelectedServiceIndex = 0 Then Call fxgServices_RowColChange
    Exit Sub
    
fxgServices_ClickError:
    Call funErrorManager("frmOptions.fxgServices_ClickError", True)
End Sub


'*******************************************************************************
Private Sub fxgServices_RowColChange()
Dim lngNewSelectedRowIndex As Long
Dim uOldService, uCurrentService As New clsService
On Error GoTo fxgServices_RowColChangeError
    
    With fxgServices
        lngNewSelectedRowIndex = .Row
        If lngNewSelectedRowIndex < 1 Then Exit Sub
        If mlngSelectedServiceIndex > .Rows - 1 Then mlngSelectedServiceIndex = .Rows - 1
        
        Set uCurrentService = guServices(lngNewSelectedRowIndex)
        If mlngSelectedServiceIndex > 0 Then
            Set uOldService = guServices(mlngSelectedServiceIndex)
            .Col = 0
            .Row = mlngSelectedServiceIndex
            .ColSel = fxgServices.Cols - 1
            .CellForeColor = vbWindowText
            .CellBackColor = vbWindowBackground
            .Col = 2
            .CellForeColor = IIf(uOldService.StockMode = enuServiceStockMode.Managed And uOldService.Stock <= uOldService.StockMinAlert, vbRed, vbWindowText)
        End If
        
        mlngSelectedServiceIndex = lngNewSelectedRowIndex
        
        .Col = 0
        .Row = lngNewSelectedRowIndex
        .ColSel = .Cols - 1
        .CellForeColor = vbHighlightText
        .CellBackColor = vbHighlight
        
    End With
    lbl(1).Caption = funLoadResString(10348)
    optServiceStockMode(uCurrentService.StockMode).value = True
    txtServiceLabel.Text = uCurrentService.Label
    txtServicePrice.Text = uCurrentService.Price
    txtServiceStock.Text = IIf(uCurrentService.StockMode = enuServiceStockMode.Managed, uCurrentService.Stock, "")
    txtServiceStockAlert.Text = IIf(uCurrentService.StockMode = enuServiceStockMode.Managed, uCurrentService.StockMinAlert, "")
    txtServiceLabel.Locked = False
    txtServicePrice.Locked = False
    optServiceStockMode(0).enabled = True
    optServiceStockMode(1).enabled = True
    txtServiceStock.Locked = False
    txtServiceStockAlert.Locked = False
    cmdServiceDelete.Visible = True
    cmdServiceModify.Visible = True
    Exit Sub
    
fxgServices_RowColChangeError:
    Call funErrorManager("frmOptions.fxgServices_RowColChangeError", True)
End Sub


'*******************************************************************************
Private Sub lstCardTypes_Click()
On Error GoTo lstCardTypes_ClickError

    cmdCrdTypeDelete.enabled = lstCardTypes.ItemData(lstCardTypes.ListIndex) <> 1
    Exit Sub
    
lstCardTypes_ClickError:
    Call funErrorManager("frmOptions.lstCardTypes_ClickError", True)
End Sub


'*******************************************************************************
Private Sub lstKilledApps_Click()

    cmdForbidAppDelete.enabled = True
End Sub


'*******************************************************************************
Private Sub lstPostpaidSsRecharges_Click()
Dim i As Integer
Dim intNumSelected As Integer
On Error GoTo lstPostpaidSsRecharges_ClickError

    If mblnPostPaidRechargeFill Or lstPostpaidSsRecharges.ListIndex < 0 Then Exit Sub
    subDebugPrint "lstPostpaidSsRecharges_Click"
    udPostpaidSsRecharges.enabled = lstPostpaidSsRecharges.Selected(lstPostpaidSsRecharges.ListIndex)
    mlngSelectedPostPaidRecharge = lstPostpaidSsRecharges.ItemData(lstPostpaidSsRecharges.ListIndex)
    If mblnPostPaidRechargeChanged Then Call subFillPostpaidRecharges
    If udPostpaidSsRecharges.enabled Then
        For i = 0 To lstPostpaidSsRecharges.ListCount - 1
            If lstPostpaidSsRecharges.Selected(i) Then intNumSelected = intNumSelected + 1
        Next i
        udPostpaidSsRecharges.Min = -intNumSelected
'        udPostpaidSsRecharges.Max = -1
        udPostpaidSsRecharges.value = -(lstPostpaidSsRecharges.ListIndex + 1)
    End If
    Exit Sub


lstPostpaidSsRecharges_ClickError:
    Call funErrorManager("frmOptions.lstPostpaidSsRecharges_ClickError", True)
End Sub


'*******************************************************************************
Private Sub lstPostpaidSsRecharges_ItemCheck(Item As Integer)
Dim i As Integer
Dim intNumSelected As Integer
On Error GoTo lstPostpaidSsRecharges_ItemCheckError

    If mblnPostPaidRechargeFill Then Exit Sub
    subDebugPrint "lstPostpaidSsRecharges_ItemCheck"
    mblnPostPaidRechargeChanged = True
    If lstPostpaidSsRecharges.Selected(Item) Then
        For i = 0 To lstPostpaidSsRecharges.ListCount - 1
            If lstPostpaidSsRecharges.Selected(i) Then intNumSelected = intNumSelected + 1
        Next i
        grstRecharges.Index = "PrimaryKey"
        Call grstRecharges.Seek("=", lstPostpaidSsRecharges.ItemData(Item))
        If Not grstRecharges.NoMatch Then
            Call grstRecharges.Edit
            grstRecharges!usage_order = intNumSelected '+ 1
            Call grstRecharges.Update
        End If
    Else
        grstRecharges.Index = "PrimaryKey"
        Call grstRecharges.Seek("=", lstPostpaidSsRecharges.ItemData(Item))
        If Not grstRecharges.NoMatch Then
            Call grstRecharges.Edit
            grstRecharges!usage_order = 0
            Call grstRecharges.Update
        End If
        For i = Item To lstPostpaidSsRecharges.ListCount - 1
            If lstPostpaidSsRecharges.Selected(i) Then
                Call grstRecharges.Seek("=", lstPostpaidSsRecharges.ItemData(i))
                If Not grstRecharges.NoMatch Then
                    Call grstRecharges.Edit
                    grstRecharges!usage_order = grstRecharges!usage_order - 1
                    Call grstRecharges.Update
                End If
            End If
        Next i
    End If
    Exit Sub

lstPostpaidSsRecharges_ItemCheckError:
    Call funErrorManager("frmOptions.lstPostpaidSsRecharges_ItemCheckError", True)
End Sub


'*******************************************************************************
Private Sub lstPostpaidSsRecharges_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo lstPostpaidSsRecharges_KeyUpError

    If KeyCode = vbKeySpace Then Call lstPostpaidSsRecharges_Click
    Exit Sub

lstPostpaidSsRecharges_KeyUpError:
    Call funErrorManager("frmOptions.lstPostpaidSsRecharges_KeyUpError", True)
End Sub


'*******************************************************************************
Private Sub lstStatsTimes_Click()
On Error GoTo lstStatsTimes_ClickError

    cmdStatsTimeDelete.enabled = True
    Exit Sub
    
lstStatsTimes_ClickError:
    Call funErrorManager("frmOptions.lstStatsTimes_ClickError", True)
End Sub


'*******************************************************************************
Private Sub optPostpaidSsMode_Click(Index As Integer)
Dim i As Integer
On Error Resume Next

    mblnSettingsChanged = True
    For i = optPostpaidSsMode.LBound To optPostpaidSsMode.UBound
        fraPostpaidSsMode(i).Visible = i = Index
    Next i
End Sub


'*******************************************************************************
Private Sub optPostpaidSsModeRechSubMode_Click(Index As Integer)
On Error Resume Next

    mblnSettingsChanged = True
    dtpPostpaidSsEndIncr.enabled = Index = enuPostpaidSsSubMode.RechSimple
    lbl(89).enabled = Index = enuPostpaidSsSubMode.RechSimple
End Sub


'*******************************************************************************
Private Sub optUsageDefault_Click(Index As Integer)
Dim i As Integer
On Error Resume Next
    
    mblnSettingsChanged = True
    For i = optUsageDefault.LBound To optUsageDefault.UBound
        optUsageDefault(i).Caption = IIf(optUsageDefault(i).value, funLoadResString(11013), "")
    Next i
End Sub


'*******************************************************************************
Private Sub optWksnCnxAuto_Click()
On Error Resume Next
    
    mblnSettingsChanged = True
    txtPortNum.enabled = False
    lbl(3).enabled = False
End Sub


'*******************************************************************************
Private Sub optWksnCnxManual_Click()
On Error Resume Next
    
    mblnSettingsChanged = True
    txtPortNum.enabled = True
    lbl(3).enabled = True
End Sub


'*******************************************************************************
Private Sub optServiceStockMode_Click(Index As Integer)
On Error GoTo optServiceStockMode_ClickError
    
    txtServiceStockAlert.Locked = Index = enuServiceStockMode.Unmanaged
    txtServiceStock.Locked = Index = enuServiceStockMode.Unmanaged
    lbl(55).enabled = Index = enuServiceStockMode.Managed
    lbl(56).enabled = Index = enuServiceStockMode.Managed
    txtServiceStock.Text = IIf(Index = enuServiceStockMode.Unmanaged, "", txtServiceStock.Text)
    txtServiceStockAlert.Text = IIf(Index = enuServiceStockMode.Unmanaged, "", txtServiceStockAlert.Text)
    Exit Sub
    
optServiceStockMode_ClickError:
    Call funErrorManager("frmOptions.optServiceStockMode_ClickError", True)
End Sub


'*******************************************************************************
Private Sub optServiceStockMode_KeyPress(Index As Integer, KeyAscii As Integer)

    mblnServiceChanged = True
End Sub


'*******************************************************************************
Private Sub optServiceStockMode_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    mblnServiceChanged = True
End Sub


'*******************************************************************************
Private Sub tbsTabs_Click()
Dim i As Integer
On Error GoTo tbsTabs_ClickError
    
    For i = 1 To fraOptions.Count
        fraOptions(i).Visible = i = tbsTabs.SelectedItem.Index
    Next i
    If tbsTabs.SelectedItem.Index = 3 Then Call subInitPostpaidSs
    Exit Sub
    
tbsTabs_ClickError:
    Call funErrorManager("frmOptions.tbsTabs_ClickError", True)
End Sub


'*******************************************************************************
Private Sub txtAppName_Change()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub txtAppName_GotFocus()
On Error Resume Next
    
    txtAppName.SelStart = 0
    txtAppName.SelLength = Len(txtAppName.Text)
End Sub


'*******************************************************************************
Private Sub txtCrdBonusThreshold_Change()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub txtCrdBonusThreshold_GotFocus()
On Error Resume Next
    
    txtCrdBonusThreshold.SelStart = 0
    txtCrdBonusThreshold.SelLength = Len(txtCrdBonusThreshold.Text)
End Sub


'*******************************************************************************
Private Sub txtCrdBonusThreshold_LostFocus()
Dim sngTest As Single
On Error GoTo txtCrdBonusThreshold_LostFocusError

    If Not IsNumeric(txtCrdBonusThreshold.Text) Then
        Call funMsg(GeneralError, funLoadResString(10648))
        txtCrdBonusThreshold.SelStart = 0
        txtCrdBonusThreshold.SelLength = Len(txtCrdBonusThreshold.Text)
        Call funSetFocus(txtCrdBonusThreshold)
    End If
    Exit Sub

txtCrdBonusThreshold_LostFocusError:
    Call funErrorManager("frmOptions.txtCrdBonusThreshold_LostFocusError", True)
End Sub

'*******************************************************************************
Private Sub txtCrdTypeNew_Change()
Dim strCardType As String
Dim strChar As String
Dim i As Integer
On Error GoTo txtCrdTypeNew_ChangeError
    
    For i = 1 To Len(txtCrdTypeNew.Text)
        strChar = Mid(txtCrdTypeNew.Text, i, 1)
        If strChar <> CARD_TYPE_TOKEN Then strCardType = strCardType + strChar
    Next i
    txtCrdTypeNew.Text = strCardType
    txtCrdTypeNew.SelStart = Len(txtCrdTypeNew.Text)
    Exit Sub

txtCrdTypeNew_ChangeError:
    Call funErrorManager("frmOptions.txtCrdTypeNew_ChangeError", True)
End Sub


'*******************************************************************************
Private Sub txtCrdTypeNew_GotFocus()
On Error Resume Next
    
    txtCrdTypeNew.SelStart = 0
    txtCrdTypeNew.SelLength = Len(txtCrdTypeNew.Text)
End Sub


'*******************************************************************************
Private Sub txtDBBackupPath_Change()

    mblnSettingsChanged = True
    txtDBBackupPath.ToolTipText = txtDBBackupPath.Text
End Sub


'*******************************************************************************
Private Sub txtDBPasswd_Change()

    mblnSettingsChanged = True
    mblnDBPasswdChanged = True
End Sub


'*******************************************************************************
Private Sub txtDBPasswd_GotFocus()
On Error Resume Next
    
    txtDBPasswd.SelStart = 0
    txtDBPasswd.SelLength = Len(txtDBPasswd.Text)
End Sub


'*******************************************************************************
Private Sub txtDBPath_Change()
On Error Resume Next

    mblnSettingsChanged = True
    mblnDBPathChanged = True
    txtDBPath.ToolTipText = txtDBPath.Text
End Sub


'*******************************************************************************
Private Sub txtDisconConRatio_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    
    Select Case KeyCode
    Case 40
        If udDisconConRatio.value > udDisconConRatio.Min Then udDisconConRatio.value = udDisconConRatio.value - udDisconConRatio.Increment
    Case 38
        If udDisconConRatio.value < udDisconConRatio.Max Then udDisconConRatio.value = udDisconConRatio.value + udDisconConRatio.Increment
    End Select
End Sub


'*******************************************************************************
Private Sub txtDisconConRatio_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub


'*******************************************************************************
Private Sub txtDisconConRatio_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    Select Case KeyCode
    Case 38, 40
        txtDisconConRatio.SelStart = 0
        txtDisconConRatio.SelLength = Len(txtDisconConRatio.Text)
    End Select
End Sub


'*******************************************************************************
Private Sub txtDisconConRatio_LostFocus()
On Error Resume Next
    
    If CInt(txtDisconConRatio.Text) >= udDisconConRatio.Min And CInt(txtDisconConRatio.Text) <= udDisconConRatio.Max Then
        udDisconConRatio.value = CInt(txtDisconConRatio.Text)
    Else
        txtDisconConRatio.Text = udDisconConRatio.value
    End If
End Sub


'*******************************************************************************
Private Sub txtOffPeak_Change()
On Error Resume Next

'    If CInt(txtOffPeak.Text) >= udOffPeak.Min And CInt(txtOffPeak.Text) <= udOffPeak.Max Then
'        udOffPeak.value = CInt(txtOffPeak.Text)
'    Else
'        txtOffPeak.Text = udOffPeak.value
'    End If
End Sub


'*******************************************************************************
Private Sub txtOffPeak_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    Select Case KeyCode
    Case 40
        If udOffPeak.value > udOffPeak.Min Then udOffPeak.value = udOffPeak.value - udOffPeak.Increment
    Case 38
        If udOffPeak.value < udOffPeak.Max Then udOffPeak.value = udOffPeak.value + udOffPeak.Increment
    End Select
End Sub


'*******************************************************************************
Private Sub txtMonFormat_Change()
    
    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub txtMonFormat_GotFocus()
On Error Resume Next
    
    txtMonFormat.SelStart = 0
    txtMonFormat.SelLength = Len(txtMonFormat.Text)
End Sub


'*******************************************************************************
Private Sub txtMonSymbol_Change()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub txtMonSymbol_GotFocus()
On Error Resume Next
    
    txtMonSymbol.SelStart = 0
    txtMonSymbol.SelLength = Len(txtMonSymbol.Text)
End Sub


'*******************************************************************************
Private Sub txtMsgSSAmount_Change()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub txtMsgSSAmount_GotFocus()
On Error Resume Next
    
    txtMsgSSAmount.SelStart = 0
    txtMsgSSAmount.SelLength = Len(txtMsgSSAmount.Text)
End Sub


'*******************************************************************************
Private Sub txtMsgSSAmount_LostFocus()
On Error GoTo txtMsgSSAmount_LostFocusError

    If txtMsgSSAmount.Text = "" Then txtMsgSSAmount.Text = 0
    If Not IsNumeric(txtMsgSSAmount.Text) Then
        Call funMsg(GeneralError, funLoadResString(10648))
        txtMsgSSAmount.SelStart = 0
        txtMsgSSAmount.SelLength = Len(txtMsgSSAmount.Text)
        Call funSetFocus(txtMsgSSAmount)
    End If
    Exit Sub

txtMsgSSAmount_LostFocusError:
    Call funErrorManager("frmOptions.txtMsgSSAmount_LostFocusError", True)
End Sub


'*******************************************************************************
Private Sub txtOffPeak_KeyPress(KeyAscii As Integer)

    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub


'*******************************************************************************
Private Sub txtOffPeak_KeyUp(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    Case 38, 40
        txtOffPeak.SelStart = 0
        txtOffPeak.SelLength = Len(txtOffPeak.Text)
    End Select
End Sub


'*******************************************************************************
Private Sub txtOffPeak_LostFocus()

    If CInt(txtOffPeak.Text) >= udOffPeak.Min And CInt(txtOffPeak.Text) <= udOffPeak.Max Then
        udOffPeak.value = CInt(txtOffPeak.Text)
    Else
        txtOffPeak.Text = udOffPeak.value
    End If
End Sub


'*******************************************************************************
Private Sub txtPathBlocked_Change()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub txtPathIntro_Change()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub txtPathLogo_Change()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub txtPathNoCredit_Change()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub txtPathSessionEnd_Change()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub txtPathSessionPaused_Change()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub txtPortNum_Change()
Dim strPort As String
Dim strChar As String
Dim i As Integer
On Error GoTo txtPortNum_ChangeError
    
    txtPortNum.Text = Trim(txtPortNum.Text)
    For i = 1 To Len(txtPortNum.Text)
        strChar = Mid(txtPortNum.Text, i, 1)
        If IsNumeric(strChar) Then strPort = strPort + strChar
    Next i
    txtPortNum.Text = strPort
    txtPortNum.SelStart = Len(txtPortNum.Text)
    mblnSettingsChanged = True
    Exit Sub

txtPortNum_ChangeError:
    Call funErrorManager("frmOptions.txtPortNum_ChangeError", True)
End Sub


'*******************************************************************************
Private Sub txtPortNum_GotFocus()
On Error Resume Next
    
    txtPortNum.SelStart = 0
    txtPortNum.SelLength = Len(txtPortNum.Text)
End Sub


'*******************************************************************************
Private Sub txtPortNum_LostFocus()

    If CLng(txtPortNum.Text) > 65535 Then txtPortNum.Text = 65535
End Sub


'*******************************************************************************
Private Sub txtRechargeCaption_GotFocus()
On Error Resume Next
    
    txtRechargeCaption.SelStart = 0
    txtRechargeCaption.SelLength = Len(txtRechargeCaption.Text)
End Sub

'*******************************************************************************
Private Sub txtRechargePrice_GotFocus()
On Error Resume Next
    
    txtRechargePrice.SelStart = 0
    txtRechargePrice.SelLength = Len(txtRechargePrice.Text)
End Sub


'*******************************************************************************
Private Sub txtRechargePrice_LostFocus()
On Error GoTo txtRechargePrice_LostFocusError

    If txtRechargePrice.Text = "" Then txtRechargePrice.Text = 0
    If Not IsNumeric(txtRechargePrice.Text) Then
        Call funMsg(GeneralError, funLoadResString(10648))
        txtRechargePrice.SelStart = 0
        txtRechargePrice.SelLength = Len(txtRechargePrice.Text)
        Call funSetFocus(txtRechargePrice)
    End If
    Exit Sub

txtRechargePrice_LostFocusError:
    Call funErrorManager("frmSetupWizard.txtRechargePrice_LostFocusError", True)
End Sub


'*******************************************************************************
Private Sub txtServiceLabel_GotFocus()
On Error Resume Next
    
    txtServiceLabel.SelStart = 0
    txtServiceLabel.SelLength = Len(txtServiceLabel.Text)
End Sub


'*******************************************************************************
Private Sub txtServiceLabel_KeyPress(KeyAscii As Integer)
    
    mblnServiceChanged = True
End Sub


'*******************************************************************************
Private Sub txtServicePrice_GotFocus()
On Error Resume Next
    
    txtServicePrice.SelStart = 0
    txtServicePrice.SelLength = Len(txtServicePrice.Text)
End Sub


'*******************************************************************************
Private Sub txtServicePrice_KeyPress(KeyAscii As Integer)

    mblnServiceChanged = True
End Sub


'*******************************************************************************
Private Sub txtServicePrice_LostFocus()
On Error GoTo txtServicePrice_LostFocusError

    If txtServicePrice.Text = "" Then txtServicePrice.Text = 0
    If Not IsNumeric(txtServicePrice.Text) Then
        Call funMsg(GeneralError, funLoadResString(10648))
        txtServicePrice.SelStart = 0
        txtServicePrice.SelLength = Len(txtServicePrice.Text)
        Call funSetFocus(txtServicePrice)
    End If
    Exit Sub

txtServicePrice_LostFocusError:
    Call funErrorManager("frmOptions.txtServicePrice_LostFocusError", True)
End Sub


'*******************************************************************************
Private Sub txtServiceStock_Change()
Dim strQuantity As String
Dim strChar As String
Dim i As Integer
On Error GoTo txtServiceStock_ChangeError
    
    txtServiceStock.Text = Trim(txtServiceStock.Text)
    For i = 1 To Len(txtServiceStock.Text)
        strChar = Mid(txtServiceStock.Text, i, 1)
        If IsNumeric(strChar) Or i = 1 And strChar = "-" Then strQuantity = strQuantity + strChar
    Next i
    If optServiceStockMode(enuServiceStockMode.Managed).value Then strQuantity = IIf(strQuantity = "", 0, strQuantity)
    txtServiceStock.Text = strQuantity
    txtServiceStock.SelStart = Len(txtServiceStock.Text)
    Exit Sub

txtServiceStock_ChangeError:
    Call funErrorManager("frmOptions.txtServiceStock_ChangeError", True)
End Sub


'*******************************************************************************
Private Sub txtServiceStock_GotFocus()
On Error Resume Next
    
    txtServiceStock.SelStart = 0
    txtServiceStock.SelLength = Len(txtServiceStock.Text)
End Sub


'*******************************************************************************
Private Sub txtServiceStock_KeyPress(KeyAscii As Integer)

    mblnServiceChanged = True
End Sub


'*******************************************************************************
Private Sub txtServiceStockAlert_Change()
Dim strQuantity As String
Dim strChar As String
Dim i As Integer
On Error GoTo txtServiceStockAlert_ChangeError
    
    txtServiceStockAlert.Text = Trim(txtServiceStockAlert.Text)
    For i = 1 To Len(txtServiceStockAlert.Text)
        strChar = Mid(txtServiceStockAlert.Text, i, 1)
        If IsNumeric(strChar) Or i = 1 And strChar = "-" Then strQuantity = strQuantity + strChar
    Next i
    If optServiceStockMode(enuServiceStockMode.Managed).value Then strQuantity = IIf(strQuantity = "", 0, strQuantity)
    txtServiceStockAlert.Text = strQuantity
    txtServiceStockAlert.SelStart = Len(txtServiceStockAlert.Text)
    Exit Sub

txtServiceStockAlert_ChangeError:
    Call funErrorManager("frmOptions.txtServiceStockAlert_ChangeError", True)
End Sub


'*******************************************************************************
Private Sub txtServiceStockAlert_GotFocus()
On Error Resume Next
    
    txtServiceStockAlert.SelStart = 0
    txtServiceStockAlert.SelLength = Len(txtServiceStockAlert.Text)
End Sub


'*******************************************************************************
Private Sub txtServiceStockAlert_KeyPress(KeyAscii As Integer)

    mblnServiceChanged = True
End Sub


'*******************************************************************************
Private Sub txtSessBatchEnd_Change()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub txtSessBatchStart_Change()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub txtSessClnRestoreBackgroundPath_Change()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub txtSessClnRestoreIEHomePagePath_Change()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub txtSessClnRestoreIEHomePagePath_GotFocus()
On Error Resume Next
    
    txtSessClnRestoreIEHomePagePath.SelStart = 0
    txtSessClnRestoreIEHomePagePath.SelLength = Len(txtSessClnRestoreIEHomePagePath.Text)
End Sub


'*******************************************************************************
Private Sub txtTestServerName_Change()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub txtTestServerName_GotFocus()
On Error Resume Next
    
    txtTestServerName.SelStart = 0
    txtTestServerName.SelLength = Len(txtTestServerName.Text)
End Sub


'*******************************************************************************
Private Sub txtTestServerPort_Change()
Dim strPort As String
Dim strChar As String
Dim i As Integer
On Error GoTo txtTestServerPort_ChangeError
    
    txtTestServerPort.Text = Trim(txtTestServerPort.Text)
    For i = 1 To Len(txtTestServerPort.Text)
        strChar = Mid(txtTestServerPort.Text, i, 1)
        If IsNumeric(strChar) Then strPort = strPort + strChar
    Next i
    txtTestServerPort.Text = strPort
    txtTestServerPort.SelStart = Len(txtTestServerPort.Text)
    mblnSettingsChanged = True
    Exit Sub

txtTestServerPort_ChangeError:
    Call funErrorManager("frmOptions.txtTestServerPort_ChangeError", True)
End Sub


'*******************************************************************************
Private Sub txtTestServerPort_GotFocus()
On Error Resume Next
    
    txtTestServerPort.SelStart = 0
    txtTestServerPort.SelLength = Len(txtTestServerPort.Text)
End Sub


'*******************************************************************************
Private Sub txtTestServerPort_LostFocus()

    If CLng(txtPortNum.Text) > 65535 Then txtPortNum.Text = 65535
End Sub


'*******************************************************************************
Private Sub udCliListFillInterval_Change()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub udCliPurgeDelta_Change()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub udCrdListFillInterval_Change()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub udCrdPurgeDelta_Change()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub udDisconConRatio_Change()
On Error Resume Next

    txtDisconConRatio.Text = udDisconConRatio.value
    txtDisconConRatio.SelStart = 0
    txtDisconConRatio.SelLength = Len(txtDisconConRatio.Text)
    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub udDisconConRatio_DownClick()
On Error Resume Next
    
    Call funSetFocus(txtDisconConRatio)
End Sub


'*******************************************************************************
Private Sub udDisconConRatio_UpClick()
On Error Resume Next

    Call funSetFocus(txtDisconConRatio)
End Sub

'*******************************************************************************
Private Sub udMsg1CardMinutes_Change()
On Error GoTo udMsg1CardMinutes_ChangeError

    mblnSettingsChanged = True
    If udMsg2CardMinutes.value >= udMsg1CardMinutes.value Then udMsg2CardMinutes.value = udMsg1CardMinutes.value - 1
    Exit Sub

udMsg1CardMinutes_ChangeError:
    Call funErrorManager("frmOptions.udMsg1CardMinutes_ChangeError", True)
End Sub


'*******************************************************************************
Private Sub udMsg2CardMinutes_Change()
On Error GoTo udMsg2CardMinutes_ChangeError

    mblnSettingsChanged = True
    If udMsg1CardMinutes.value <= udMsg2CardMinutes.value Then udMsg1CardMinutes.value = udMsg2CardMinutes.value + 1
    Exit Sub

udMsg2CardMinutes_ChangeError:
    Call funErrorManager("frmOptions.udMsg2CardMinutes_ChangeError", True)
End Sub


'*******************************************************************************
Private Sub udNberWorkstations_Change()

    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub subInitPostpaidSs()
On Error GoTo subInitPostpaidSsError

    Call subInitPostpaidSsEndIncrLbl
    Call subFillPostpaidRecharges
    mlngSelectedPostPaidRecharge = -1
    lstPostpaidSsRecharges.ListIndex = -1
    udPostpaidSsRecharges.enabled = False
    Exit Sub

subInitPostpaidSsError:
    Call funErrorManager("frmOptions.subInitPostpaidSsError")
End Sub


'*******************************************************************************
Private Sub subInitPostpaidSsEndIncrLbl()
Dim rstRecharges As DAO.Recordset
On Error GoTo subInitPostpaidSsEndIncrLblError


    Set rstRecharges = gdbsCybAdmin.OpenRecordset("SELECT time, price FROM recharges WHERE type_postpaid_ss=true ORDER BY time", dbOpenDynaset)
    If Not rstRecharges.EOF And Not rstRecharges.BOF Then
        Call rstRecharges.MoveLast
        lbl(89).Caption = funLoadResString(11016, funFormatTime(rstRecharges!Time, enuTimeFormat.Hh_Nmn_Ssec_dyn), funFormatCurrencyString(rstRecharges!Price / rstRecharges!Time * dtpPostpaidSsEndIncr.value, guServOpts.uConfig.strMonFormat, guServOpts.uConfig.strMonSymbol))
    End If
    Exit Sub

subInitPostpaidSsEndIncrLblError:
    Call funErrorManager("frmOptions.subInitPostpaidSsEndIncrLblError")
End Sub


'*******************************************************************************
Private Sub subFillPostpaidRecharges()
Dim i As Long
Dim strBuffer As String
Dim rstRechargesSelected As DAO.Recordset
Dim rstRechargesUnselected As DAO.Recordset
Dim dteMinTime As Date
On Error GoTo subFillPostpaidRechargesError

    subDebugPrint "subFillPostpaidRecharges"
    mblnPostPaidRechargeFill = True
    Set rstRechargesSelected = gdbsCybAdmin.OpenRecordset("SELECT id,usage_order,time,price FROM recharges WHERE type_postpaid_ss=true AND NOT usage_order=0 ORDER BY usage_order", dbOpenDynaset)
    Set rstRechargesUnselected = gdbsCybAdmin.OpenRecordset("SELECT id,usage_order,time,price FROM recharges WHERE type_postpaid_ss=true AND usage_order=0 ORDER BY time", dbOpenDynaset)
    If Not rstRechargesSelected.EOF And Not rstRechargesSelected.BOF Then
        mdtePostPaidRechargeMaxMinTime = rstRechargesSelected!Time
        dteMinTime = rstRechargesSelected!Time
    End If
    Call lstPostpaidSsRecharges.Clear
    Do While Not rstRechargesSelected.EOF
        strBuffer = funFormatTime(rstRechargesSelected!Time) & " = " & funFormatCurrencyString(rstRechargesSelected!Price, guServOpts.uConfig.strMonFormat, guServOpts.uConfig.strMonSymbol)
        Call lstPostpaidSsRecharges.AddItem(strBuffer)
        lstPostpaidSsRecharges.ItemData(lstPostpaidSsRecharges.NewIndex) = rstRechargesSelected!ID
        lstPostpaidSsRecharges.Selected(lstPostpaidSsRecharges.NewIndex) = True
        If dteMinTime > rstRechargesSelected!Time Then dteMinTime = rstRechargesSelected!Time
        Call rstRechargesSelected.MoveNext
    Loop
    mdtePostPaidRechargeMaxBaseIncr = dteMinTime
    Do While Not rstRechargesUnselected.EOF
        strBuffer = funFormatTime(rstRechargesUnselected!Time) & " = " & funFormatCurrencyString(rstRechargesUnselected!Price, guServOpts.uConfig.strMonFormat, guServOpts.uConfig.strMonSymbol)
        Call lstPostpaidSsRecharges.AddItem(strBuffer)
        lstPostpaidSsRecharges.ItemData(lstPostpaidSsRecharges.NewIndex) = rstRechargesUnselected!ID
        rstRechargesUnselected.MoveNext
    Loop
    For i = 0 To lstPostpaidSsRecharges.ListCount - 1
        If lstPostpaidSsRecharges.ItemData(i) = mlngSelectedPostPaidRecharge Then
            lstPostpaidSsRecharges.ListIndex = i
            Exit For
        End If
    Next i
    If dtpPostpaidSsBaseIncr.value > mdtePostPaidRechargeMaxBaseIncr Then
        dtpPostpaidSsBaseIncr.value = mdtePostPaidRechargeMaxBaseIncr
    ElseIf dtpPostpaidSsBaseIncr.value < MIN_PPSS_TIME_INCR Then
        dtpPostpaidSsBaseIncr.value = MIN_PPSS_TIME_INCR
    End If
    If dtpPostpaidSsMinTime.value > mdtePostPaidRechargeMaxMinTime Then
        dtpPostpaidSsMinTime.value = mdtePostPaidRechargeMaxMinTime
    ElseIf dtpPostpaidSsMinTime.value < dtpPostpaidSsBaseIncr.value Then
        dtpPostpaidSsMinTime.value = dtpPostpaidSsBaseIncr.value
    ElseIf dtpPostpaidSsMinTime.value < MIN_PPSS_TIME_INCR Then
        dtpPostpaidSsMinTime.value = MIN_PPSS_TIME_INCR
    End If
    mblnPostPaidRechargeFill = False
    mblnPostPaidRechargeChanged = False
    Exit Sub
    
subFillPostpaidRechargesError:
    mblnPostPaidRechargeFill = False
    Call funErrorManager("frmOptions.subFillPostpaidRechargesError")
End Sub


'*******************************************************************************
Private Sub subFillRecharges()
Dim i As Long
Dim strType As String
On Error GoTo subFillRechargesError

    If Not gblnDBOK Then Exit Sub
    With fxgRecharges
        If grstRecharges.BOF = True And grstRecharges.EOF = True Or grstRecharges.RecordCount = 0 Then
            mlngSelectedRechargeIndex = 0
            .Rows = 22
            .enabled = False
            .Font = ""
            For i = 1 To .Rows - 1
                .TextMatrix(i, 0) = ""
                .TextMatrix(i, 1) = ""
                .TextMatrix(i, 2) = ""
                .TextMatrix(i, 3) = ""
                .TextMatrix(i, 4) = ""
                .TextMatrix(i, 5) = ""
                .Col = 0
                .Row = i
                .ColSel = .Cols - 1
                .CellForeColor = vbWindowText
                .CellBackColor = vbWindowBackground
            Next i
            .Row = 0
            txtRechargeCaption.Text = ""
            txtRechargePrice.Text = ""
            chkRechargeCli.value = vbUnchecked
            chkRechargeCrd.value = vbUnchecked
            chkRechargePostSs.value = vbUnchecked
            chkRechargePreSs.value = vbUnchecked
            Exit Sub
        End If
        
        .Row = 0
        For i = 0 To .Cols - 1
            .Col = i
            .CellFontBold = False
        Next i
        Select Case mlngSelectedRechargeOrder
        Case enuRechargeOrder.orderByTime
            grstRecharges.Index = "orderByTime"
            .Col = 0
        Case enuRechargeOrder.orderByPrice
            grstRecharges.Index = "orderByPrice"
            .Col = 1
        Case enuRechargeOrder.orderByValidity
            .Col = 2
            grstRecharges.Index = "orderByValidity"
        Case enuRechargeOrder.orderByType
            .Col = 3
            grstRecharges.Index = "orderByType"
        Case Else
            .Col = 4
            grstRecharges.Index = "RechargesIndex"
        End Select
        .CellFontBold = True

        Call grstRecharges.MoveFirst
        .Rows = grstRecharges.RecordCount + 1
        For i = 1 To grstRecharges.RecordCount
            If grstRecharges!unlimited_play Then
                .TextMatrix(i, 0) = funLoadResString(11000)
            Else
                .TextMatrix(i, 0) = funFormatTime(grstRecharges!Time)
            End If
            
            .TextMatrix(i, 1) = funFormatCurrencyString(grstRecharges!Price, guServOpts.uConfig.strMonFormat)
            Select Case grstRecharges!validity_type
            Case enuRechargeValidityType.Hours
                .TextMatrix(i, 2) = grstRecharges!validity & " " & funLoadResString(10995)
            Case enuRechargeValidityType.Days
                .TextMatrix(i, 2) = grstRecharges!validity & " " & funLoadResString(10996)
            Case enuRechargeValidityType.Weeks
                .TextMatrix(i, 2) = grstRecharges!validity & " " & funLoadResString(10997)
            Case enuRechargeValidityType.Months
                .TextMatrix(i, 2) = grstRecharges!validity & " " & funLoadResString(10998)
            Case enuRechargeValidityType.Years
                .TextMatrix(i, 2) = grstRecharges!validity & " " & funLoadResString(10999)
            End Select
            
            strType = ""
            If grstRecharges!type_card Then strType = strType & " " & funLoadResString(10991)
            If grstRecharges!type_client Then strType = strType & " " & funLoadResString(10992)
            If grstRecharges!type_prepaid_ss Then strType = strType & " " & funLoadResString(10993)
            If grstRecharges!type_postpaid_ss Then strType = strType & " " & funLoadResString(10994)
            .TextMatrix(i, 3) = Trim(strType)
'            .TextMatrix(i, 3) = IIf(grstRecharges!type_postpaid_ss, funLoadResString(10313), "")
            .TextMatrix(i, 4) = IIf(IsNull(grstRecharges!Caption), "", grstRecharges!Caption)
            .TextMatrix(i, 5) = grstRecharges!ID
            Call grstRecharges.MoveNext
            
        Next i
    End With
    Exit Sub
    
subFillRechargesError:
    Call funErrorManager("frmOptions.subFillRechargesError", True)
End Sub


'*******************************************************************************
Private Sub subFillLocalServices()
Dim i As Long
On Error GoTo subFillLocalServicesError

    If Not gblnDBOK Then Exit Sub
    With fxgServices
        .Rows = IIf(guServices.Count = 0, 22, guServices.Count + 1)
        .enabled = guServices.Count <> 0
        .Font = ""
        For i = 1 To .Rows - 1
            .TextMatrix(i, 0) = ""
            .TextMatrix(i, 1) = ""
            .TextMatrix(i, 2) = ""
            .TextMatrix(i, 3) = ""
            .TextMatrix(i, 4) = ""
            .Col = 0
            .Row = i
            .ColSel = .Cols - 1
            .CellForeColor = vbWindowText
            .CellBackColor = vbWindowBackground
        Next i
        For i = 1 To guServices.Count
            .TextMatrix(i, 0) = guServices(i).Label
            .TextMatrix(i, 1) = funFormatCurrencyString(guServices(i).Price, guServOpts.uConfig.strMonFormat)
            .Col = 2
            .Row = i
            If guServices(i).StockMode = enuServiceStockMode.Managed Then
                .TextMatrix(i, 2) = guServices(i).Stock
                .TextMatrix(i, 3) = guServices(i).StockMinAlert
                .CellForeColor = IIf(i = mlngSelectedServiceIndex, vbHighlightText, IIf(guServices(i).Stock <= guServices(i).StockMinAlert, vbRed, vbWindowText))
            Else
                .TextMatrix(i, 2) = funLoadResString(10332)
                .TextMatrix(i, 3) = funLoadResString(10332)
                .CellForeColor = IIf(i = mlngSelectedServiceIndex, vbHighlightText, vbWindowText)
            End If
        Next i
    End With
    Exit Sub
    
subFillLocalServicesError:
    Call funErrorManager("frmOptions.subFillLocalServicesError", True)
End Sub


'*******************************************************************************
Private Sub udOffPeak_Change()

    txtOffPeak.Text = udOffPeak.value
    txtOffPeak.SelStart = 0
    txtOffPeak.SelLength = Len(txtOffPeak.Text)
    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Public Sub subLocalizeInterface()
On Error GoTo subLocalizeInterfaceError

    Call subSetCharset(Me, guServOpts.uConfig.intInterfaceLangCharset)
    
    Me.Caption = funLoadResString(10000) & " - " & funLoadResString(10124)
    
    tbsTabs.Tabs(1).Caption = funLoadResString(10255)
    tbsTabs.Tabs(1).ToolTipText = funLoadResString(10474)
    tbsTabs.Tabs(2).Caption = funLoadResString(10149)
    tbsTabs.Tabs(2).ToolTipText = funLoadResString(10475)
    tbsTabs.Tabs(3).Caption = funLoadResString(10258)
    tbsTabs.Tabs(3).ToolTipText = funLoadResString(10476)
    tbsTabs.Tabs(4).Caption = funLoadResString(10841)
    tbsTabs.Tabs(4).ToolTipText = funLoadResString(10477)
    tbsTabs.Tabs(5).Caption = funLoadResString(10259)
    tbsTabs.Tabs(5).ToolTipText = funLoadResString(10478)
    tbsTabs.Tabs(6).Caption = funLoadResString(10326)
    tbsTabs.Tabs(6).ToolTipText = funLoadResString(10479)
    tbsTabs.Tabs(7).Caption = funLoadResString(10420)
    tbsTabs.Tabs(7).ToolTipText = funLoadResString(10480)
    
    optWksnCnxAuto.Caption = funLoadResString(10733)
    optWksnCnxAuto.ToolTipText = funLoadResString(10735)
    optWksnCnxManual.Caption = funLoadResString(10734)
    optWksnCnxManual.ToolTipText = funLoadResString(10736)
    
    txtPortNum.ToolTipText = funLoadResString(10482)
    txtDBPath.ToolTipText = funLoadResString(10483)
    txtDBPasswd.ToolTipText = funLoadResString(10484)
    txtTestServerName.ToolTipText = funLoadResString(10487)
    txtTestServerPort.ToolTipText = funLoadResString(10488)
    txtMonFormat.ToolTipText = funLoadResString(10491)
    txtMonSymbol.ToolTipText = funLoadResString(10490)
    dtpNumDaysInMonth.ToolTipText = funLoadResString(10492)
    cmbLanguage.ToolTipText = funLoadResString(10493)
    lblNberWorkstations.ToolTipText = funLoadResString(10499)
    udNberWorkstations.ToolTipText = funLoadResString(10499)
    txtDisconConRatio.ToolTipText = funLoadResString(10504)
    udDisconConRatio.ToolTipText = funLoadResString(10504)
    txtOffPeak.ToolTipText = funLoadResString(10505)
    udOffPeak.ToolTipText = funLoadResString(10505)
    lblFollowInterval.ToolTipText = funLoadResString(10651)
    udFollowInterval.ToolTipText = funLoadResString(10651)
    
    chkUsageSelfServicePrePaid.Caption = funLoadResString(10210)
    chkUsageSelfServicePrePaid.ToolTipText = funLoadResString(10500)
    chkUsageSelfServicePostPaid.Caption = funLoadResString(10978)
    chkUsageSelfServicePostPaid.ToolTipText = funLoadResString(10979)
    chkUsageClients.Caption = funLoadResString(10836)
    chkUsageClients.ToolTipText = funLoadResString(10838)
    chkUsageCards.Caption = funLoadResString(10835)
    chkUsageCards.ToolTipText = funLoadResString(10837)
    optUsageDefault(1).ToolTipText = funLoadResString(11012)
    optUsageDefault(2).ToolTipText = funLoadResString(11012)
    optUsageDefault(4).ToolTipText = funLoadResString(11012)
    optUsageDefault(5).ToolTipText = funLoadResString(11012)
    
    chkHideCardNum.Caption = funLoadResString(10684)
    chkHideCardNum.ToolTipText = funLoadResString(10685)
    chkEnableLoggingURLs.Caption = funLoadResString(10704)
    chkEnableLoggingURLs.ToolTipText = funLoadResString(10705)
    cmbURLtimespan.ToolTipText = funLoadResString(10705)
    chkEnablePause.Caption = funLoadResString(10738)
    chkEnablePause.ToolTipText = funLoadResString(10739)
    chkBackupDB.Caption = funLoadResString(10211)
    chkBackupDB.ToolTipText = funLoadResString(10485)
    chkBackupDBSendToWksn.Caption = funLoadResString(10709)
    chkBackupDBSendToWksn.ToolTipText = funLoadResString(10710)
    chkCompactDB.Caption = funLoadResString(10518)
    chkCompactDB.ToolTipText = funLoadResString(10519)
    chkCrdAutoPurge.Caption = funLoadResString(10212)
    chkCliAutoPurge.Caption = funLoadResString(10212)
    chkWksnNotification.Caption = funLoadResString(10213)
    chkWksnNotification.ToolTipText = funLoadResString(10489)
    chkCnxTesting.Caption = funLoadResString(10214)
    chkCnxTesting.ToolTipText = funLoadResString(10486)
    chkMsgCard.Caption = funLoadResString(10215)
    chkMsgCard.ToolTipText = funLoadResString(10502)
    chkMsgSS.Caption = funLoadResString(10216)
    chkMsgSS.ToolTipText = funLoadResString(10503)
    
    chkRechargeCrd.Caption = funLoadResString(10314)
    chkRechargeCli.Caption = funLoadResString(10831)
    chkRechargePreSs.Caption = funLoadResString(10984)
    chkRechargePostSs.Caption = funLoadResString(10217)
    
    chkRechargeTimeUnlimited.Caption = funLoadResString(11000)
    
    chkStartForceWksnsConnected.Caption = funLoadResString(10218)
    chkStartForceWksnsConnected.ToolTipText = funLoadResString(10501)
    chkCrdBonus.Caption = funLoadResString(10550)
    chkCliBonus.Caption = funLoadResString(10550)
    chkCliEnableExpiry.Caption = funLoadResString(10863)
    chkCliEnableExpiry.ToolTipText = funLoadResString(10864)
    
    chkWksnMsgLogoUpload.Caption = funLoadResString(10606)
    chkWksnLockdown.Caption = funLoadResString(10584)
    chkWksnLockdown.ToolTipText = funLoadResString(10737)
    chkWksnLock(0).Caption = funLoadResString(10577)
    chkWksnLock(1).Caption = funLoadResString(10578)
    chkWksnLock(2).Caption = funLoadResString(10579)
    chkWksnLock(3).Caption = funLoadResString(10580)
    chkWksnLock(4).Caption = funLoadResString(10581)
    chkWksnLock(5).Caption = funLoadResString(10582)
    chkWksnLock(6).Caption = funLoadResString(10583)
    
    cmdCancel.Caption = funLoadResString(10123)
    cmdCancel.ToolTipText = funLoadResString(10495)
    cmdDBPath.Caption = funLoadResString(10219)
    cmdRechargeAdd.Caption = funLoadResString(10220)
    cmdRechargeDelete.Caption = funLoadResString(10221)
    cmdRechargeModify.Caption = funLoadResString(10222)
    cmdServiceNew.Caption = funLoadResString(10356)
    cmdServiceAdd.Caption = funLoadResString(10220)
    cmdServiceAddCancel.Caption = funLoadResString(10123)
    cmdServiceDelete.Caption = funLoadResString(10221)
    cmdServiceModify.Caption = funLoadResString(10222)
    cmdStatsTimeAdd.Caption = funLoadResString(10220)
    cmdStatsTimeDelete.Caption = funLoadResString(10221)
    cmdCrdTypeAdd.Caption = funLoadResString(10220)
    cmdCrdTypeDelete.Caption = funLoadResString(10221)
    
    cmdShellBlocked.ToolTipText = funLoadResString(10609)
    cmdShellIntro.ToolTipText = funLoadResString(10609)
    cmdShellLogo.ToolTipText = funLoadResString(10609)
    cmdShellNoCredit.ToolTipText = funLoadResString(10609)
    cmdShellSessionEnd.ToolTipText = funLoadResString(10609)
    cmdShellSessionPaused.ToolTipText = funLoadResString(10609)
    
    cmdPrevIntro.ToolTipText = funLoadResString(10013)
    cmdPrevBlocked.ToolTipText = funLoadResString(10013)
    cmdPrevNoCredit.ToolTipText = funLoadResString(10013)
    cmdPrevSessionEnd.ToolTipText = funLoadResString(10013)
    cmdPrevSessionPaused.ToolTipText = funLoadResString(10013)
    cmdPrevLogo.ToolTipText = funLoadResString(10013)
    
    cmdSave.Caption = funLoadResString(10223)
    cmdSave.ToolTipText = funLoadResString(10494)
    
    With fxgRecharges
        .TextMatrix(0, 0) = funLoadResString(10143)
        .TextMatrix(0, 1) = funLoadResString(10144)
        .TextMatrix(0, 2) = funLoadResString(10145)
        .TextMatrix(0, 3) = funLoadResString(10649)
        .TextMatrix(0, 4) = funLoadResString(10146)
    End With
    
    With fxgServices
        .TextMatrix(0, 0) = funLoadResString(10328)
        .TextMatrix(0, 1) = funLoadResString(10329)
        .TextMatrix(0, 2) = funLoadResString(10330)
        .TextMatrix(0, 3) = funLoadResString(10347)
    End With
    
    optServiceStockMode(0).Caption = funLoadResString(10355)
    optServiceStockMode(1).Caption = funLoadResString(10354)
    
    optStatsGraphType(0).Caption = funLoadResString(10440)
    optStatsGraphType(1).Caption = funLoadResString(10441)
    
    optPostpaidSsMode(1).Caption = funLoadResString(11002)
    optPostpaidSsMode(1).ToolTipText = funLoadResString(11008)
    optPostpaidSsMode(2).Caption = funLoadResString(11003)
    optPostpaidSsMode(2).ToolTipText = funLoadResString(11009)
    optPostpaidSsModeRechSubMode(1).Caption = funLoadResString(11017)
    optPostpaidSsModeRechSubMode(1).ToolTipText = funLoadResString(11015)
    optPostpaidSsModeRechSubMode(2).Caption = funLoadResString(11018)
    optPostpaidSsModeRechSubMode(2).ToolTipText = funLoadResString(11014)
    dtpPostpaidSsEndIncr.ToolTipText = funLoadResString(11019)
    dtpPostpaidSsGraceTime.ToolTipText = funLoadResString(11020)
    
    lbl(0).Caption = funLoadResString(10224)
    lbl(1).Caption = funLoadResString(10348)
    lbl(2).Caption = funLoadResString(11001)
    lbl(3).Caption = funLoadResString(10226)
    lbl(4).Caption = funLoadResString(10227)
    lbl(5).Caption = funLoadResString(10228)
    lbl(6).Caption = funLoadResString(10229)
    lbl(7).Caption = funLoadResString(10230)
    lbl(8).Caption = funLoadResString(10231)
    lbl(9).Caption = funLoadResString(10232)
    lbl(10).Caption = funLoadResString(10233)
    lbl(11).Caption = funLoadResString(10234)
    lbl(12).Caption = funLoadResString(10235)
    lbl(13).Caption = funLoadResString(10236)
    lbl(14).Caption = funLoadResString(10834)
    lbl(15).Caption = funLoadResString(10843)
    lbl(16).Caption = funLoadResString(10842)
    lbl(17).Caption = funLoadResString(10240)
    lbl(18).Caption = funLoadResString(10549)
    lbl(19).Caption = funLoadResString(10548)
    lbl(20).Caption = funLoadResString(10866)
    lbl(21).Caption = funLoadResString(10241)
    lbl(22).Caption = funLoadResString(10241)
    lbl(23).Caption = funLoadResString(10238)
    lbl(24).Caption = funLoadResString(10242)
    lbl(25).Caption = funLoadResString(10238)
    lbl(26).Caption = funLoadResString(10243)
    lbl(27).Caption = funLoadResString(10244)
    lbl(28).Caption = funLoadResString(10245)
    lbl(29).Caption = funLoadResString(10876)
    lbl(30).Caption = funLoadResString(10882)
    lbl(30).ToolTipText = funLoadResString(10883)
    lbl(31).Caption = funLoadResString(10913)
    lbl(32).Caption = funLoadResString(10249)
    lbl(33).Caption = funLoadResString(10250)
    lbl(34).Caption = funLoadResString(10251)
    lbl(35).Caption = funLoadResString(10252)
    lbl(36).Caption = funLoadResString(10253)
    lbl(37).Caption = funLoadResString(10254)
    lbl(38).Caption = funLoadResString(10255)
    lbl(39).Caption = funLoadResString(10256)
    lbl(40).Caption = funLoadResString(10225)
    lbl(41).Caption = funLoadResString(10257)
    lbl(42).Caption = funLoadResString(10226)
    lbl(43).Caption = funLoadResString(10116)
    lbl(44).Caption = funLoadResString(10437)
    lbl(45).Caption = funLoadResString(10292)
    lbl(46).Caption = funLoadResString(10839)
    lbl(47).Caption = funLoadResString(10323)
    lbl(48).Caption = funLoadResString(10346)
    lbl(49).Caption = funLoadResString(10349)
    lbl(50).Caption = funLoadResString(10350)
    lbl(51).Caption = funLoadResString(10351)
    lbl(52).Caption = funLoadResString(10344)
    lbl(53).Caption = funLoadResString(10345)
    lbl(54).Caption = funLoadResString(10425)
    lbl(55).Caption = funLoadResString(10352)
    lbl(56).Caption = funLoadResString(10353)
    lbl(57).Caption = funLoadResString(10423)
    lbl(58).Caption = funLoadResString(10424)
    lbl(59).Caption = funLoadResString(10001)
    lbl(60).Caption = funLoadResString(10001)
    lbl(61).Caption = funLoadResString(10438)
    lbl(62).Caption = funLoadResString(10439)
    lbl(63).Caption = funLoadResString(10419)
    lbl(64).Caption = funLoadResString(10442)
    lbl(65).Caption = funLoadResString(10529)
    lbl(66).Caption = funLoadResString(10840)
    lbl(67).Caption = funLoadResString(10548)
    lbl(68).Caption = funLoadResString(10549)
    lbl(69).Caption = funLoadResString(10576)
    lbl(70).Caption = funLoadResString(10599)
    lbl(71).Caption = funLoadResString(10600)
    lbl(72).Caption = funLoadResString(10601)
    lbl(73).Caption = funLoadResString(10602)
    lbl(74).Caption = funLoadResString(10603)
    lbl(75).Caption = funLoadResString(10604)
    lbl(76).Caption = funLoadResString(10605)
    lbl(77).Caption = funLoadResString(10650)
    lbl(78).Caption = funLoadResString(10239)
    lbl(79).Caption = funLoadResString(10672)
    lbl(80).Caption = funLoadResString(10670)
    lbl(81).Caption = funLoadResString(10671)
    lbl(82).Caption = funLoadResString(10680)
    lbl(83).Caption = funLoadResString(10968)
    lbl(84).Caption = " " & funLoadResString(11004)
    lbl(85).Caption = funLoadResString(11005)
    lbl(86).Caption = funLoadResString(11006)
    lbl(87).Caption = funLoadResString(11007)
    
    lbl(89).Caption = funLoadResString(11016)
    lbl(89).ToolTipText = funLoadResString(11019)
    lbl(90).Caption = funLoadResString(11021)
    lbl(90).ToolTipText = funLoadResString(11020)
    
    
    lblLocked.Caption = funLoadResString(10521)
    lblLockedPassword.Caption = funLoadResString(10116)
    txtLockedPassword.ToolTipText = funLoadResString(10524)
    lblUnlocked.Caption = funLoadResString(10522)
    cmdUnlockOptions.ToolTipText = funLoadResString(10523)
    
    chkSessClnAlignDesktopIcons.Caption = funLoadResString(10657)
    chkSessClnDeleteDocHistory.Caption = funLoadResString(10658)
    chkSessClnRestoreBackground.Caption = funLoadResString(10659)
    chkSessClnRestoreIEHomePage.Caption = funLoadResString(10660)
    chkSessClnRestoreIESettings.Caption = funLoadResString(10667)
    chkSessClnRestoreIMs.Caption = funLoadResString(10732)

    chkSessBatchStart.Caption = funLoadResString(10867)
    chkSessBatchEnd.Caption = funLoadResString(10868)
    cmdSessBachEndOpen.ToolTipText = funLoadResString(10872)
    cmdSessBachStartOpen.ToolTipText = funLoadResString(10872)

    chkForbidIE.Caption = funLoadResString(10663)
    chkForbidMozilla.Caption = funLoadResString(10665)
    chkForbidNetscape.Caption = funLoadResString(10664)
    chkForbidOE.Caption = funLoadResString(10666)
    
    cmdRetreaveCurrentBackground.ToolTipText = funLoadResString(10661)
    cmdBlankIEHomePage.ToolTipText = funLoadResString(10662)
    
    cmdForbidAppAdd.ToolTipText = funLoadResString(10668)
    cmdForbidAppDelete.ToolTipText = funLoadResString(10669)
    
    txtAppName.ToolTipText = funLoadResString(10673)
    lstKilledApps.ToolTipText = funLoadResString(10674)
    
    bffBackup.Caption = funLoadResString(10909)
    bffBackup.Prompt = funLoadResString(10910)

    Exit Sub

subLocalizeInterfaceError:
    Call funErrorManager("frmOptions.subLocalizeInterfaceError", True)
End Sub


'*******************************************************************************
Private Sub subInitSelectedService()
Dim uCurrentService As New clsService
On Error GoTo subInitSelectedServiceError

    With fxgServices
        If mlngSelectedServiceIndex > 0 And mlngSelectedServiceIndex <= guServices.Count Then
            Set uCurrentService = guServices(mlngSelectedServiceIndex)
            .Col = 0
            .Row = mlngSelectedServiceIndex
            .ColSel = fxgServices.Cols - 1
            .CellForeColor = vbWindowText
            .CellBackColor = vbWindowBackground
            .Col = 2
            .CellForeColor = IIf(uCurrentService.StockMode = enuServiceStockMode.Managed And uCurrentService.Stock <= uCurrentService.StockMinAlert, vbRed, vbWindowText)
        End If
        .enabled = False
    End With
    mlngSelectedServiceIndex = 0
    lbl(1).Caption = funLoadResString(10357)
    txtServiceLabel.Text = ""
    txtServiceLabel.Locked = True
    txtServicePrice.Text = ""
    txtServicePrice.Locked = True
    optServiceStockMode(0).value = False
    optServiceStockMode(0).enabled = False
    optServiceStockMode(1).value = False
    optServiceStockMode(1).enabled = False
    txtServiceStock.Text = ""
    txtServiceStock.Locked = True
    txtServiceStockAlert.Text = ""
    txtServiceStockAlert.Locked = True
    lbl(55).enabled = True
    lbl(56).enabled = True
    Exit Sub
    
subInitSelectedServiceError:
    Call funErrorManager("frmOptions.subInitSelectedServiceError", True)
End Sub


'*******************************************************************************
Private Sub udOffPeak_DownClick()
On Error Resume Next
    
    Call funSetFocus(txtOffPeak)
End Sub


'*******************************************************************************
Private Sub udOffPeak_UpClick()
On Error Resume Next
    
    Call funSetFocus(txtOffPeak)
End Sub


'*******************************************************************************
Private Sub udPostpaidSsRecharges_Change()
Dim intCurrentPosition As Integer
Dim i As Integer
On Error GoTo udPostpaidSsRecharges_ChangeError

    If mblnPostPaidRechargeFill Then Exit Sub
    subDebugPrint "udPostpaidSsRecharges_Change " & udPostpaidSsRecharges.value
    intCurrentPosition = lstPostpaidSsRecharges.ListIndex + 1
    
    If intCurrentPosition > -udPostpaidSsRecharges.value Then ' Go up (-1)
        grstRecharges.Index = "PrimaryKey"
        Call grstRecharges.Seek("=", lstPostpaidSsRecharges.ItemData(intCurrentPosition - 1))
        If Not grstRecharges.NoMatch Then
            Call grstRecharges.Edit
            grstRecharges!usage_order = grstRecharges!usage_order - 1
            Call grstRecharges.Update
        End If
        Call grstRecharges.Seek("=", lstPostpaidSsRecharges.ItemData(-udPostpaidSsRecharges.value - 1))
        If Not grstRecharges.NoMatch Then
            Call grstRecharges.Edit
            grstRecharges!usage_order = grstRecharges!usage_order + 1
            Call grstRecharges.Update
        End If
    ElseIf intCurrentPosition < -udPostpaidSsRecharges.value Then 'Go down +1
        grstRecharges.Index = "PrimaryKey"
        Call grstRecharges.Seek("=", lstPostpaidSsRecharges.ItemData(intCurrentPosition - 1))
        If Not grstRecharges.NoMatch Then
            Call grstRecharges.Edit
            grstRecharges!usage_order = grstRecharges!usage_order + 1
            Call grstRecharges.Update
        End If
        Call grstRecharges.Seek("=", lstPostpaidSsRecharges.ItemData(-udPostpaidSsRecharges.value - 1))
        If Not grstRecharges.NoMatch Then
            Call grstRecharges.Edit
            grstRecharges!usage_order = grstRecharges!usage_order - 1
            Call grstRecharges.Update
        End If
    End If
    If intCurrentPosition <> -udPostpaidSsRecharges.value Then Call subFillPostpaidRecharges
    Exit Sub
    
udPostpaidSsRecharges_ChangeError:
    Call funErrorManager("frmOptions.udPostpaidSsRecharges_ChangeError", True)
End Sub


'*******************************************************************************
Private Sub udPostpaidSsRecharges_DownClick()
On Error GoTo udPostpaidSsRecharges_DownClickError

    Exit Sub
    
udPostpaidSsRecharges_DownClickError:
    Call funErrorManager("frmOptions.udPostpaidSsRecharges_DownClickError", True)
End Sub


'*******************************************************************************
Private Sub udPostpaidSsRecharges_UpClick()
On Error GoTo udPostpaidSsRecharges_UpClickError

    Exit Sub
    
udPostpaidSsRecharges_UpClickError:
    Call funErrorManager("frmOptions.udPostpaidSsRecharges_UpClickError", True)
End Sub


'*******************************************************************************
Private Sub udStatsGraphStart_Change()
On Error GoTo udStatsGraphStart_ChangeError
    
    If udStatsGraphStart.value = udStatsGraphStop.value Then udStatsGraphStop.value = udStatsGraphStart.value + 1
    mblnSettingsChanged = True
    Exit Sub
    
udStatsGraphStart_ChangeError:
    Call funErrorManager("frmOptions.udStatsGraphStart_ChangeError", True)
End Sub


'*******************************************************************************
Private Sub udStatsGraphStop_Change()
On Error GoTo udStatsGraphStop_ChangeError

    If udStatsGraphStart.value = udStatsGraphStop.value Then udStatsGraphStart.value = udStatsGraphStop.value - 1
    mblnSettingsChanged = True
    Exit Sub
    
udStatsGraphStop_ChangeError:
    Call funErrorManager("frmOptions.udStatsGraphStop_ChangeError", True)
End Sub


'*******************************************************************************
Private Sub cmdPathIntro_Click()
On Error GoTo cmdPathIntro_ClickError
    
    Call subCdlOpenFile(cdlPath, txtPathIntro, guServOpts.uWorkstations.uPaths.strIntroPath, funGetFolderPath(CSIDL_DESKTOP), funLoadResString(10607), funLoadResString(10604))
    Exit Sub
    
cmdPathIntro_ClickError:
    Call funErrorManager("frmOptions.cmdPathIntro_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdPathLogo_Click()
On Error GoTo cmdPathLogo_ClickError

    Call subCdlOpenFile(cdlPath, txtPathLogo, guServOpts.uWorkstations.uPaths.strLogoPath, funGetFolderPath(CSIDL_DESKTOP), funLoadResString(10608), funLoadResString(10599))
    Exit Sub
    
cmdPathLogo_ClickError:
    Call funErrorManager("frmOptions.cmdPathLogo_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdPathNoCredit_Click()
On Error GoTo cmdPathNoCredit_ClickError

    Call subCdlOpenFile(cdlPath, txtPathNoCredit, guServOpts.uWorkstations.uPaths.strNoCreditPath, funGetFolderPath(CSIDL_DESKTOP), funLoadResString(10607), funLoadResString(10602))
    Exit Sub
    
cmdPathNoCredit_ClickError:
    Call funErrorManager("frmOptions.cmdPathNoCredit_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdPathSessionEnd_Click()
On Error GoTo cmdPathSessionEnd_ClickError

    Call subCdlOpenFile(cdlPath, txtPathSessionEnd, guServOpts.uWorkstations.uPaths.strSessionEndPath, funGetFolderPath(CSIDL_DESKTOP), funLoadResString(10607), funLoadResString(10601))
    Exit Sub
    
cmdPathSessionEnd_ClickError:
    Call funErrorManager("frmOptions.cmdPathSessionEnd_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdPathSessionPaused_Click()
On Error GoTo cmdPathSessionPaused_ClickError

    Call subCdlOpenFile(cdlPath, txtPathSessionPaused, guServOpts.uWorkstations.uPaths.strSessionPausedPath, funGetFolderPath(CSIDL_DESKTOP), funLoadResString(10607), funLoadResString(10600))
    Exit Sub
    
cmdPathSessionPaused_ClickError:
    Call funErrorManager("frmOptions.cmdPathSessionPaused_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdPathBlocked_Click()
On Error GoTo cmdPathBlocked_ClickError

    Call subCdlOpenFile(cdlPath, txtPathBlocked, guServOpts.uWorkstations.uPaths.strBlockedPath, funGetFolderPath(CSIDL_DESKTOP), funLoadResString(10607), funLoadResString(10603))
    Exit Sub
    
cmdPathBlocked_ClickError:
    Call funErrorManager("frmOptions.cmdPathBlocked_ClickError", True)
End Sub


'*******************************************************************************
Private Function funWksnLockDrivesStr(ByVal lngDriveFlags As Long) As String
Dim i As Integer
On Error GoTo funWksnLockDrivesStrError
    
    For i = 25 To 0 Step -1
        If lngDriveFlags - 2 ^ i >= 0 Then
            lngDriveFlags = lngDriveFlags - 2 ^ i
            funWksnLockDrivesStr = Chr(65 + i) & "," & funWksnLockDrivesStr
        End If
    Next i
    If Len(funWksnLockDrivesStr) > 0 Then
        funWksnLockDrivesStr = left(funWksnLockDrivesStr, Len(funWksnLockDrivesStr) - 1)
    Else
        funWksnLockDrivesStr = funLoadResString(10904)
    End If
    Exit Function
    
funWksnLockDrivesStrError:
    Call funErrorManager("frmOptions.funWksnLockDrivesStrError")
End Function


'*******************************************************************************
Private Sub subRedefineUsageDefault()
On Error Resume Next

    If optUsageDefault(enuSessionType.Card).enabled Then
        optUsageDefault(enuSessionType.Card).value = True
    ElseIf optUsageDefault(enuSessionType.Client).enabled Then
        optUsageDefault(enuSessionType.Client).value = True
'    ElseIf optUsageDefault(enuSessionType.SelfServicePrePaid).enabled Then
'        optUsageDefault(enuSessionType.SelfServicePrePaid).value = vbChecked
    ElseIf optUsageDefault(enuSessionType.Selfservice).enabled Then
        optUsageDefault(enuSessionType.Selfservice).value = True
    Else
        optUsageDefault(enuSessionType.Card).value = False
        optUsageDefault(enuSessionType.Card).Caption = ""
        optUsageDefault(enuSessionType.Client).value = False
        optUsageDefault(enuSessionType.Client).Caption = ""
        optUsageDefault(enuSessionType.Selfservice).value = False
        optUsageDefault(enuSessionType.Selfservice).Caption = ""
    End If
End Sub
