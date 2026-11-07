VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "sysinfo.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmMain 
   Caption         =   "_"
   ClientHeight    =   8340
   ClientLeft      =   165
   ClientTop       =   705
   ClientWidth     =   15240
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrSystTrayUpdate 
      Interval        =   1000
      Left            =   14640
      Top             =   7080
   End
   Begin VB.Timer tmrSystTrayLogoff 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   14640
      Top             =   7500
   End
   Begin MSComctlLib.ImageList imgSystTrayIcons 
      Left            =   13980
      Top             =   7380
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1454
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1708
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrCliRefresh 
      Interval        =   20000
      Left            =   3600
      Top             =   7440
   End
   Begin VB.Timer tmrCrdRefresh 
      Interval        =   20000
      Left            =   3120
      Top             =   7440
   End
   Begin MSWinsockLib.Winsock wskWakeOnLan 
      Left            =   13920
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Timer tmrActivateServer 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   12180
      Top             =   0
   End
   Begin MSWinsockLib.Winsock wskHandshakeWorkstation 
      Left            =   14760
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock wskFindWorkstation 
      Left            =   14340
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      LocalPort       =   4646
   End
   Begin VB.Timer tmrBackupDBToWorkstation 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   14760
      Top             =   3900
   End
   Begin VB.Timer tmrScreenCapture 
      Interval        =   10000
      Left            =   14760
      Top             =   2580
   End
   Begin VB.Timer tmrCnxTest 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   14760
      Top             =   1200
   End
   Begin MSWinsockLib.Winsock wskCnxTest 
      Left            =   14760
      Top             =   780
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "www.fasonet.bf"
      RemotePort      =   80
   End
   Begin MSComctlLib.ImageList imlSatusBar 
      Left            =   600
      Top             =   7380
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1862
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B56
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E8A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrStatusBarUpdate 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   0
      Top             =   7500
   End
   Begin SysInfoLib.SysInfo sysInfo 
      Left            =   11400
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wskServ 
      Index           =   0
      Left            =   13500
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   5000
   End
   Begin MSComctlLib.ImageList imlTabs 
      Left            =   10200
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2024
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":217E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2432
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":258C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":26E6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrStateUpdade 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   12900
      Top             =   60
   End
   Begin MSComctlLib.TabStrip tbsFunctions 
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   661
      Style           =   2
      HotTracking     =   -1  'True
      ImageList       =   "imlTabs"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "_"
            Key             =   "k1"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "_"
            Key             =   "k4"
            ImageVarType    =   2
            ImageIndex      =   4
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "_"
            Key             =   "k5"
            ImageVarType    =   2
            ImageIndex      =   5
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "_"
            Key             =   "k6"
            ImageVarType    =   2
            ImageIndex      =   6
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
   Begin MSComctlLib.StatusBar stbStatus 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   8040
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   26238
            MinWidth        =   18
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   159
            MinWidth        =   18
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraFunctions 
      BorderStyle     =   0  'None
      Height          =   6795
      Index           =   3
      Left            =   60
      TabIndex        =   192
      Top             =   480
      Visible         =   0   'False
      Width           =   11775
      Begin VB.Frame fraCliControls 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   615
         Left            =   2760
         TabIndex        =   198
         Top             =   5580
         Width           =   8955
         Begin VB.CommandButton cmdCliButtons 
            Height          =   615
            Index           =   8
            Left            =   4080
            Picture         =   "frmMain.frx":2840
            Style           =   1  'Graphical
            TabIndex        =   230
            TabStop         =   0   'False
            Top             =   0
            Width           =   195
         End
         Begin VB.CommandButton cmdCliButtons 
            Height          =   615
            Index           =   7
            Left            =   8400
            Picture         =   "frmMain.frx":28EA
            Style           =   1  'Graphical
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   0
            Width           =   195
         End
         Begin VB.CommandButton cmdCliButtons 
            Caption         =   "_"
            Height          =   615
            Index           =   0
            Left            =   0
            Picture         =   "frmMain.frx":2994
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   0
            Width           =   1395
         End
         Begin VB.CommandButton cmdCliButtons 
            Caption         =   "_"
            Height          =   615
            Index           =   1
            Left            =   1440
            Picture         =   "frmMain.frx":2ADE
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   0
            Width           =   1395
         End
         Begin VB.CommandButton cmdCliButtons 
            Caption         =   "_"
            Height          =   615
            Index           =   6
            Left            =   7200
            Picture         =   "frmMain.frx":2C28
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   0
            Width           =   1215
         End
         Begin VB.CommandButton cmdCliButtons 
            Caption         =   "_"
            Height          =   615
            Index           =   2
            Left            =   2880
            Picture         =   "frmMain.frx":2D72
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   0
            Width           =   1215
         End
         Begin VB.CommandButton cmdCliButtons 
            Caption         =   "_"
            DisabledPicture =   "frmMain.frx":2EBC
            Height          =   615
            Index           =   4
            Left            =   4320
            Picture         =   "frmMain.frx":3006
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   0
            Width           =   1395
         End
         Begin VB.CommandButton cmdCliButtons 
            Caption         =   "_"
            Height          =   615
            Index           =   5
            Left            =   5760
            Picture         =   "frmMain.frx":3150
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   0
            Width           =   1395
         End
         Begin VB.CommandButton cmdCliButtons 
            Caption         =   "_"
            DisabledPicture =   "frmMain.frx":329A
            Height          =   615
            Index           =   3
            Left            =   4320
            Picture         =   "frmMain.frx":33E4
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   0
            Visible         =   0   'False
            Width           =   1395
         End
      End
      Begin VB.Frame fraCliListControls 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   315
         Left            =   0
         TabIndex        =   196
         Top             =   0
         Width           =   2715
         Begin VB.TextBox txtCliFind 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            HideSelection   =   0   'False
            Left            =   1020
            MaxLength       =   15
            TabIndex        =   47
            Top             =   40
            Width           =   1215
         End
         Begin VB.CommandButton cmdCliFind 
            Height          =   315
            Left            =   2340
            Picture         =   "frmMain.frx":352E
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   0
            Width           =   375
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "_"
            Height          =   195
            Index           =   35
            Left            =   0
            TabIndex        =   197
            Top             =   60
            Width           =   915
         End
         Begin VB.Shape shpLine 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000010&
            Height          =   315
            Index           =   4
            Left            =   960
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.Frame fracliListTotal 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   0
         TabIndex        =   193
         Top             =   3420
         Width           =   2715
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "_"
            Height          =   195
            Index           =   34
            Left            =   60
            TabIndex        =   195
            Top             =   60
            Width           =   1635
         End
         Begin VB.Label lblCliListTotal 
            BackColor       =   &H80000018&
            Height          =   195
            Left            =   1740
            TabIndex        =   194
            Top             =   60
            Width           =   915
         End
         Begin VB.Shape shpLine 
            BorderColor     =   &H80000010&
            BorderStyle     =   6  'Inside Solid
            Height          =   315
            Index           =   3
            Left            =   0
            Top             =   0
            Width           =   2715
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fxgClients 
         Height          =   2895
         Left            =   60
         TabIndex        =   49
         Top             =   420
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   5106
         _Version        =   393216
         Rows            =   100
         Cols            =   4
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
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Shape shpCliControlsBorder 
         BorderColor     =   &H80000010&
         Height          =   855
         Left            =   7740
         Top             =   660
         Width           =   2235
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
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
         Index           =   38
         Left            =   2820
         TabIndex        =   212
         Top             =   2220
         Width           =   2355
      End
      Begin VB.Label lblCliSel 
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
         Height          =   3015
         Index           =   7
         Left            =   5280
         TabIndex        =   211
         Top             =   2220
         Width           =   2655
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
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
         Index           =   36
         Left            =   2820
         TabIndex        =   210
         Top             =   1860
         Width           =   2355
      End
      Begin VB.Label lblCliSel 
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
         Index           =   6
         Left            =   5220
         TabIndex        =   209
         Top             =   1860
         Width           =   2715
      End
      Begin VB.Label lblCliSel 
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
         Index           =   1
         Left            =   5220
         TabIndex        =   206
         Top             =   420
         Width           =   2715
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
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
         Index           =   42
         Left            =   2820
         TabIndex        =   205
         Top             =   420
         Width           =   2355
      End
      Begin VB.Label lblCliSel 
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
         Index           =   2
         Left            =   5220
         TabIndex        =   204
         Top             =   780
         Width           =   2715
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
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
         Index           =   41
         Left            =   2820
         TabIndex        =   203
         Top             =   780
         Width           =   2355
      End
      Begin VB.Label lblCliSel 
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
         Index           =   3
         Left            =   5220
         TabIndex        =   202
         Top             =   1140
         Width           =   2715
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
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
         Index           =   40
         Left            =   2820
         TabIndex        =   201
         Top             =   1140
         Width           =   2355
      End
      Begin VB.Label lblCliSel 
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
         Index           =   4
         Left            =   5520
         TabIndex        =   199
         Top             =   1500
         Width           =   2415
      End
      Begin VB.Shape shpCliListBorder 
         BorderColor     =   &H80000010&
         Height          =   3075
         Left            =   0
         Top             =   360
         Width           =   2715
      End
      Begin VB.Label lblCliSel 
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
         Height          =   3015
         Index           =   8
         Left            =   5220
         TabIndex        =   213
         Top             =   2220
         Width           =   75
      End
      Begin VB.Label lblCliSel 
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   3000
         TabIndex        =   207
         Top             =   0
         Width           =   6375
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
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
         Index           =   39
         Left            =   2820
         TabIndex        =   200
         Top             =   1500
         Width           =   2355
      End
      Begin VB.Image imgCliDisabled 
         Height          =   240
         Left            =   5280
         Picture         =   "frmMain.frx":3678
         Top             =   1500
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgCliEnabled 
         Height          =   240
         Left            =   5280
         Picture         =   "frmMain.frx":37C2
         Top             =   1500
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblCliSel 
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
         Index           =   5
         Left            =   5220
         TabIndex        =   208
         Top             =   1500
         Width           =   315
      End
   End
   Begin VB.Frame fraFunctions 
      BorderStyle     =   0  'None
      Height          =   6795
      Index           =   2
      Left            =   60
      TabIndex        =   100
      Top             =   480
      Visible         =   0   'False
      Width           =   11595
      Begin VB.Frame fraCrdPending 
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   0
         TabIndex        =   237
         Top             =   4080
         Width           =   2715
         Begin VB.TextBox txtCrdPending 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1440
            MaxLength       =   6
            TabIndex        =   34
            Top             =   400
            Width           =   735
         End
         Begin VB.CommandButton cmdCrdPendingEnable 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2280
            Picture         =   "frmMain.frx":390C
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   360
            Width           =   375
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid fxgCrdPending 
            Height          =   1095
            Left            =   360
            TabIndex        =   36
            Top             =   840
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   1931
            _Version        =   393216
            Rows            =   100
            Cols            =   4
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
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   4
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "_"
            Height          =   195
            Index           =   52
            Left            =   60
            TabIndex        =   238
            Top             =   420
            Width           =   1275
         End
         Begin VB.Shape shpLine 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000010&
            Height          =   315
            Index           =   7
            Left            =   1380
            Top             =   360
            Width           =   855
         End
         Begin VB.Shape shpLine 
            BorderColor     =   &H80000010&
            BorderStyle     =   6  'Inside Solid
            Height          =   735
            Index           =   6
            Left            =   0
            Top             =   0
            Width           =   2715
         End
         Begin VB.Shape shpCrdPending 
            BorderColor     =   &H80000010&
            Height          =   1335
            Left            =   0
            Top             =   720
            Width           =   2715
         End
         Begin VB.Label lbl 
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
            Index           =   53
            Left            =   0
            TabIndex        =   239
            Top             =   0
            Width           =   2715
         End
      End
      Begin VB.TextBox txtCrdComment 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         Left            =   5280
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   38
         Top             =   2340
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ComboBox cmbCrdType 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmMain.frx":3A56
         Left            =   5220
         List            =   "frmMain.frx":3A58
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   1860
         Width           =   2895
      End
      Begin VB.Frame fraCrdListTotal 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   0
         TabIndex        =   160
         Top             =   3420
         Width           =   2715
         Begin VB.Shape shpLine 
            BorderColor     =   &H80000010&
            BorderStyle     =   6  'Inside Solid
            Height          =   315
            Index           =   1
            Left            =   0
            Top             =   0
            Width           =   2715
         End
         Begin VB.Label lblCrdListTotal 
            BackColor       =   &H80000018&
            Height          =   195
            Left            =   1740
            TabIndex        =   162
            Top             =   60
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "_"
            Height          =   195
            Index           =   18
            Left            =   0
            TabIndex        =   161
            Top             =   60
            Width           =   1695
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fxgCards 
         Height          =   2895
         Left            =   60
         TabIndex        =   33
         Top             =   420
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   5106
         _Version        =   393216
         Rows            =   100
         Cols            =   4
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
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Frame fraCrdListControls 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   315
         Left            =   0
         TabIndex        =   121
         Top             =   0
         Width           =   2715
         Begin VB.CommandButton cmdCrdFind 
            Height          =   315
            Left            =   2340
            Picture         =   "frmMain.frx":3A5A
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   0
            Width           =   375
         End
         Begin VB.TextBox txtCrdFindNum 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            HideSelection   =   0   'False
            Left            =   1020
            MaxLength       =   15
            TabIndex        =   31
            Top             =   40
            Width           =   1215
         End
         Begin VB.Shape shpLine 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000010&
            Height          =   315
            Index           =   0
            Left            =   960
            Top             =   0
            Width           =   1335
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "_"
            Height          =   195
            Index           =   5
            Left            =   0
            TabIndex        =   122
            Top             =   60
            Width           =   915
         End
      End
      Begin VB.Frame fraCrdControls 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   615
         Left            =   2820
         TabIndex        =   120
         Top             =   3900
         Width           =   8775
         Begin VB.CommandButton cmdCrdButtons 
            Height          =   615
            Index           =   7
            Left            =   8400
            Picture         =   "frmMain.frx":3BA4
            Style           =   1  'Graphical
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   0
            Width           =   195
         End
         Begin VB.CommandButton cmdCrdButtons 
            Caption         =   "_"
            Height          =   615
            Index           =   6
            Left            =   7200
            Picture         =   "frmMain.frx":3C4E
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   0
            Width           =   1215
         End
         Begin VB.CommandButton cmdCrdButtons 
            Caption         =   "_"
            Height          =   615
            Index           =   2
            Left            =   2880
            Picture         =   "frmMain.frx":3D98
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   0
            Width           =   1395
         End
         Begin VB.CommandButton cmdCrdButtons 
            Caption         =   "_"
            Height          =   615
            Index           =   1
            Left            =   1440
            Picture         =   "frmMain.frx":3EE2
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   0
            Width           =   1395
         End
         Begin VB.CommandButton cmdCrdButtons 
            Caption         =   "_"
            Height          =   615
            Index           =   5
            Left            =   5760
            Picture         =   "frmMain.frx":402C
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   0
            Width           =   1395
         End
         Begin VB.CommandButton cmdCrdButtons 
            Caption         =   "_"
            Height          =   615
            Index           =   0
            Left            =   0
            Picture         =   "frmMain.frx":4176
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   0
            Width           =   1395
         End
         Begin VB.CommandButton cmdCrdDelete 
            Caption         =   "_"
            DisabledPicture =   "frmMain.frx":42C0
            Height          =   615
            Left            =   4320
            Picture         =   "frmMain.frx":440A
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   0
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.CommandButton cmdCrdButtons 
            Caption         =   "_"
            DisabledPicture =   "frmMain.frx":4554
            Height          =   615
            Index           =   3
            Left            =   4320
            Picture         =   "frmMain.frx":469E
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   0
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.CommandButton cmdCrdButtons 
            Caption         =   "_"
            DisabledPicture =   "frmMain.frx":47E8
            Height          =   615
            Index           =   4
            Left            =   4320
            Picture         =   "frmMain.frx":4932
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   0
            Width           =   1395
         End
      End
      Begin VB.Image imgCrdStateInactive 
         Height          =   240
         Left            =   5280
         Picture         =   "frmMain.frx":4A7C
         Top             =   1500
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgCrdStateActive 
         Height          =   240
         Left            =   5280
         Picture         =   "frmMain.frx":4BC6
         Top             =   1500
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Shape shpCrdComment 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   375
         Left            =   5220
         Top             =   2280
         Width           =   3735
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
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
         Index           =   20
         Left            =   2820
         TabIndex        =   164
         Top             =   2340
         Width           =   2355
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
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
         Index           =   19
         Left            =   2820
         TabIndex        =   163
         Top             =   1920
         Width           =   2355
      End
      Begin VB.Shape shpCrdListBorder 
         BorderColor     =   &H80000010&
         Height          =   3075
         Left            =   0
         Top             =   360
         Width           =   2715
      End
      Begin VB.Label lblCrdSel 
         BackColor       =   &H80000018&
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
         Index           =   4
         Left            =   5520
         TabIndex        =   132
         Top             =   1500
         Width           =   2415
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
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
         Index           =   10
         Left            =   2820
         TabIndex        =   130
         Top             =   1140
         Width           =   2355
      End
      Begin VB.Label lblCrdSel 
         BackColor       =   &H80000018&
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
         Index           =   3
         Left            =   5220
         TabIndex        =   129
         Top             =   1140
         Width           =   2715
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
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
         Index           =   9
         Left            =   2820
         TabIndex        =   128
         Top             =   780
         Width           =   2355
      End
      Begin VB.Label lblCrdSel 
         BackColor       =   &H80000018&
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
         Index           =   2
         Left            =   5220
         TabIndex        =   127
         Top             =   780
         Width           =   2715
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
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
         Index           =   7
         Left            =   2820
         TabIndex        =   125
         Top             =   420
         Width           =   2355
      End
      Begin VB.Label lblCrdSel 
         BackColor       =   &H80000018&
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
         Index           =   1
         Left            =   5220
         TabIndex        =   124
         Top             =   420
         Width           =   2715
      End
      Begin VB.Shape shpCrdControlsBorder 
         BorderColor     =   &H80000010&
         Height          =   855
         Left            =   7740
         Top             =   660
         Width           =   2235
      End
      Begin VB.Label lblCrdSel 
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   3300
         TabIndex        =   123
         Top             =   0
         Width           =   6375
      End
      Begin VB.Label lblCrdSel 
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
         Index           =   5
         Left            =   5220
         TabIndex        =   187
         Top             =   1500
         Width           =   315
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
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
         Index           =   11
         Left            =   2820
         TabIndex        =   131
         Top             =   1500
         Width           =   2355
      End
   End
   Begin VB.Frame fraFunctions 
      BorderStyle     =   0  'None
      Height          =   6795
      Index           =   6
      Left            =   60
      TabIndex        =   119
      Top             =   480
      Visible         =   0   'False
      Width           =   11355
      Begin VB.CheckBox chkTypeIncome 
         Caption         =   "_"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   251
         Top             =   3600
         Value           =   1  'Checked
         Width           =   2595
      End
      Begin VB.CheckBox chkTypeIncome 
         Caption         =   "_"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   214
         Top             =   4080
         Value           =   1  'Checked
         Width           =   2595
      End
      Begin VB.OptionButton optIncomeFilter 
         Caption         =   "_"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   84
         Top             =   1020
         Width           =   2595
      End
      Begin VB.OptionButton optIncomeFilter 
         Caption         =   "_"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   86
         Top             =   1500
         Width           =   2595
      End
      Begin VB.OptionButton optIncomeFilter 
         Caption         =   "_"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   87
         Top             =   1740
         Width           =   2595
      End
      Begin VB.OptionButton optIncomeFilter 
         Caption         =   "_"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   85
         Top             =   1260
         Width           =   2595
      End
      Begin VB.OptionButton optIncomeFilter 
         Caption         =   "_"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   88
         Top             =   1980
         Width           =   2595
      End
      Begin VB.OptionButton optIncomeFilter 
         Caption         =   "_"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   90
         Top             =   2220
         Width           =   1275
      End
      Begin VB.OptionButton optIncomeFilter 
         Caption         =   "_"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   82
         Top             =   540
         Value           =   -1  'True
         Width           =   2595
      End
      Begin VB.OptionButton optIncomeFilter 
         Caption         =   "_"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   83
         Top             =   780
         Width           =   2595
      End
      Begin MSComDlg.CommonDialog cdlIncomeExport 
         Left            =   180
         Top             =   5400
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame fraIncomeControls 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1275
         Left            =   60
         TabIndex        =   138
         Top             =   4680
         Width           =   2655
         Begin VB.CommandButton cmdIncomePrint 
            Caption         =   "_"
            Enabled         =   0   'False
            Height          =   615
            Left            =   1380
            Picture         =   "frmMain.frx":4D10
            Style           =   1  'Graphical
            TabIndex        =   246
            Top             =   660
            Width           =   1275
         End
         Begin VB.CommandButton cmdIncomeFilter 
            Caption         =   "_"
            Height          =   615
            Left            =   0
            Picture         =   "frmMain.frx":4E5A
            Style           =   1  'Graphical
            TabIndex        =   96
            Top             =   0
            Width           =   2655
         End
         Begin VB.CommandButton cmdIncomeExport 
            Caption         =   "_"
            Height          =   615
            Left            =   0
            Picture         =   "frmMain.frx":4FA4
            Style           =   1  'Graphical
            TabIndex        =   97
            Top             =   660
            Width           =   1275
         End
      End
      Begin VB.CheckBox chkTypeIncome 
         Caption         =   "_"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   95
         Top             =   4320
         Value           =   1  'Checked
         Width           =   2595
      End
      Begin VB.CheckBox chkTypeIncome 
         Caption         =   "_"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   94
         Top             =   3840
         Value           =   1  'Checked
         Width           =   2595
      End
      Begin VB.CheckBox chkTypeIncome 
         Caption         =   "_"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   93
         Top             =   3360
         Value           =   1  'Checked
         Width           =   2595
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   315
         Left            =   1500
         TabIndex        =   91
         Top             =   2400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CalendarTitleBackColor=   -2147483628
         Format          =   49807361
         CurrentDate     =   37622
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fxgIncome 
         Height          =   2895
         Left            =   2880
         TabIndex        =   98
         Top             =   60
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   5106
         _Version        =   393216
         Rows            =   10
         Cols            =   10
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
         SelectionMode   =   1
         AllowUserResizing=   1
         BorderStyle     =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   10
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   315
         Left            =   1500
         TabIndex        =   92
         Top             =   2760
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CalendarTitleBackColor=   -2147483628
         Format          =   49807361
         CurrentDate     =   38353
      End
      Begin VB.OptionButton optIncomeFilter 
         Caption         =   "_"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   89
         Top             =   2100
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.Label lbl 
         Caption         =   "_"
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   159
         Top             =   300
         Width           =   2595
      End
      Begin VB.Shape shpIncomeListBorder 
         BorderColor     =   &H80000010&
         Height          =   4875
         Left            =   2820
         Top             =   0
         Width           =   4275
      End
      Begin VB.Label lbl 
         Caption         =   "_"
         Height          =   195
         Index           =   17
         Left            =   120
         TabIndex        =   137
         Top             =   3120
         Width           =   2595
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   16
         Left            =   180
         TabIndex        =   136
         Top             =   2820
         Width           =   1275
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   15
         Left            =   120
         TabIndex        =   135
         Top             =   2460
         Width           =   1335
      End
      Begin VB.Shape shpIncomeFilters 
         BorderColor     =   &H80000010&
         Height          =   4800
         Left            =   0
         Top             =   0
         Width           =   2775
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         Caption         =   "Filtres"
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
         Left            =   0
         TabIndex        =   134
         Top             =   0
         Width           =   2775
      End
   End
   Begin VB.Frame fraFunctions 
      BorderStyle     =   0  'None
      Height          =   6795
      Index           =   4
      Left            =   60
      TabIndex        =   139
      Top             =   480
      Width           =   10695
      Begin VB.Frame fraServicesControls 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   615
         Left            =   4080
         TabIndex        =   149
         Top             =   3360
         Width           =   3615
         Begin VB.CommandButton cmdServicesCollect 
            Caption         =   "_"
            Enabled         =   0   'False
            Height          =   615
            Left            =   0
            Picture         =   "frmMain.frx":50EE
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   0
            Width           =   1275
         End
      End
      Begin VB.TextBox txtServicesControlsQuantity 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   5940
         TabIndex        =   60
         Text            =   "0"
         Top             =   1200
         Width           =   1155
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fxgServices 
         Height          =   2895
         Left            =   120
         TabIndex        =   59
         Top             =   420
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   5106
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
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSComCtl2.UpDown udServicesControlsQuantity 
         Height          =   315
         Left            =   7200
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   1140
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         BuddyControl    =   "txtServicesControlsQuantity"
         BuddyDispid     =   196664
         OrigLeft        =   2040
         OrigTop         =   480
         OrigRight       =   2280
         OrigBottom      =   735
         Max             =   1000
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   0   'False
      End
      Begin VB.Label lblServSel 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   5880
         TabIndex        =   148
         Top             =   1560
         Width           =   1875
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   23
         Left            =   4140
         TabIndex        =   147
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
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
         Index           =   12
         Left            =   4080
         TabIndex        =   146
         Top             =   1200
         Width           =   1755
      End
      Begin VB.Shape shpLine 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   5
         Left            =   5880
         Top             =   1140
         Width           =   1275
      End
      Begin VB.Shape shpServicesControlsBorder 
         BorderColor     =   &H80000010&
         Height          =   4035
         Left            =   4020
         Top             =   0
         Width           =   5655
      End
      Begin VB.Shape shpServicesListBorder 
         BorderColor     =   &H80000010&
         Height          =   4875
         Left            =   0
         Top             =   0
         Width           =   3915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
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
         Index           =   22
         Left            =   4140
         TabIndex        =   143
         Top             =   420
         Width           =   1695
      End
      Begin VB.Label lblServSel 
         BackColor       =   &H80000018&
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
         Left            =   5880
         TabIndex        =   142
         Top             =   420
         Width           =   1875
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
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
         Index           =   21
         Left            =   4080
         TabIndex        =   141
         Top             =   780
         Width           =   1755
      End
      Begin VB.Label lblServSel 
         BackColor       =   &H80000018&
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
         Index           =   2
         Left            =   5880
         TabIndex        =   140
         Top             =   780
         Width           =   1875
      End
      Begin VB.Label lblServSel 
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   4020
         TabIndex        =   144
         Top             =   0
         Width           =   3135
      End
      Begin VB.Label lblServicesTitle 
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   145
         Top             =   0
         Width           =   3135
      End
   End
   Begin VB.Frame fraFunctions 
      BorderStyle     =   0  'None
      Height          =   6795
      Index           =   1
      Left            =   60
      TabIndex        =   99
      Top             =   480
      Width           =   12495
      Begin VB.Frame fraWksnConsole 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   0
         TabIndex        =   172
         Top             =   0
         Width           =   2715
         Begin VB.Label lblConsoleWksnControl 
            BackStyle       =   0  'Transparent
            Caption         =   "_"
            ForeColor       =   &H80000011&
            Height          =   195
            Left            =   600
            TabIndex        =   183
            Top             =   300
            Width           =   2055
         End
         Begin VB.Label lblConsoleName 
            BackStyle       =   0  'Transparent
            Caption         =   "_"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   600
            TabIndex        =   173
            Top             =   60
            Width           =   2055
         End
         Begin VB.Shape shpWksnConsoleBorder 
            BorderColor     =   &H80000010&
            Height          =   615
            Left            =   0
            Top             =   0
            Width           =   2715
         End
         Begin VB.Image imgConsole 
            Height          =   480
            Left            =   60
            Picture         =   "frmMain.frx":5238
            Top             =   60
            Width           =   480
         End
         Begin VB.Shape shpWksnConsoleBackground 
            BackColor       =   &H80000014&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   615
            Left            =   0
            Top             =   0
            Width           =   2715
         End
      End
      Begin VB.Frame fraWksnContainer 
         BorderStyle     =   0  'None
         Height          =   3015
         Left            =   0
         TabIndex        =   170
         Top             =   660
         Width           =   2715
         Begin VB.Frame fraWksn 
            BorderStyle     =   0  'None
            Height          =   1875
            Left            =   0
            TabIndex        =   171
            Top             =   0
            Width           =   2475
            Begin cybserv.ucWorkstation ucWorkstations 
               Height          =   615
               Index           =   1
               Left            =   0
               TabIndex        =   1
               TabStop         =   0   'False
               Top             =   0
               Width           =   2430
               _extentx        =   4286
               _extenty        =   1085
               caption         =   ""
               status          =   3
               statetext       =   "Not present"
               active          =   -1  'True
               connectionindex =   ""
               statetext       =   "Not present"
            End
         End
      End
      Begin MSComCtl2.FlatScrollBar fsbWksn 
         Height          =   4755
         Left            =   3000
         TabIndex        =   169
         Top             =   660
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   8387
         _Version        =   393216
         Appearance      =   0
         Max             =   13
         Orientation     =   1245184
      End
      Begin VB.Frame fraWksnControls 
         BorderStyle     =   0  'None
         Height          =   6435
         Left            =   3420
         TabIndex        =   101
         Top             =   0
         Visible         =   0   'False
         Width           =   8955
         Begin VB.Frame fraWksnButtons 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   615
            Left            =   0
            TabIndex        =   133
            Top             =   5640
            Width           =   8835
            Begin VB.CommandButton cmdWksnButtons 
               Height          =   615
               Index           =   6
               Left            =   4080
               Picture         =   "frmMain.frx":5542
               Style           =   1  'Graphical
               TabIndex        =   27
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   195
            End
            Begin MSComctlLib.ImageList imlWksnDisable 
               Left            =   6120
               Top             =   360
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
                     Picture         =   "frmMain.frx":55EC
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmMain.frx":5746
                     Key             =   ""
                  EndProperty
               EndProperty
            End
            Begin MSComctlLib.ImageList imlWksnRecharge 
               Left            =   4740
               Top             =   360
               _ExtentX        =   1005
               _ExtentY        =   1005
               BackColor       =   -2147483643
               ImageWidth      =   16
               ImageHeight     =   16
               MaskColor       =   12632256
               _Version        =   393216
               BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
                  NumListImages   =   3
                  BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmMain.frx":58A0
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmMain.frx":59FA
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmMain.frx":5B54
                     Key             =   ""
                  EndProperty
               EndProperty
            End
            Begin VB.CommandButton cmdWksnButtons 
               Caption         =   "_"
               DisabledPicture =   "frmMain.frx":5CAE
               Height          =   615
               Index           =   10
               Left            =   5760
               Picture         =   "frmMain.frx":5DF8
               Style           =   1  'Graphical
               TabIndex        =   29
               Top             =   0
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.CheckBox chkWksnConnect 
               Caption         =   "_"
               Height          =   615
               Left            =   7200
               Picture         =   "frmMain.frx":5F42
               Style           =   1  'Graphical
               TabIndex        =   30
               Top             =   0
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.CommandButton cmdWksnButtons 
               Caption         =   "_"
               Height          =   615
               Index           =   0
               Left            =   0
               Picture         =   "frmMain.frx":608C
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   0
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.CommandButton cmdWksnButtons 
               Caption         =   "_"
               DisabledPicture =   "frmMain.frx":61D6
               Height          =   615
               Index           =   1
               Left            =   0
               Picture         =   "frmMain.frx":6320
               Style           =   1  'Graphical
               TabIndex        =   17
               Top             =   0
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.CommandButton cmdWksnButtons 
               Caption         =   "_"
               Height          =   615
               Index           =   2
               Left            =   0
               Picture         =   "frmMain.frx":646A
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   0
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.CommandButton cmdWksnButtons 
               Caption         =   "_"
               Height          =   615
               Index           =   15
               Left            =   1440
               Picture         =   "frmMain.frx":65B4
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   0
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.CommandButton cmdWksnButtons 
               Caption         =   "_"
               Height          =   615
               Index           =   3
               Left            =   1440
               Picture         =   "frmMain.frx":66FE
               Style           =   1  'Graphical
               TabIndex        =   22
               Top             =   0
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.CommandButton cmdWksnButtons 
               Caption         =   "_"
               DisabledPicture =   "frmMain.frx":6848
               Height          =   615
               Index           =   4
               Left            =   1440
               Picture         =   "frmMain.frx":6992
               Style           =   1  'Graphical
               TabIndex        =   21
               Top             =   0
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.CommandButton cmdWksnButtons 
               Caption         =   "_"
               Height          =   615
               Index           =   5
               Left            =   2880
               Picture         =   "frmMain.frx":6ADC
               Style           =   1  'Graphical
               TabIndex        =   26
               Top             =   0
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.CommandButton cmdWksnButtons 
               Caption         =   "_"
               Enabled         =   0   'False
               Height          =   615
               Index           =   7
               Left            =   2880
               Picture         =   "frmMain.frx":6C26
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   0
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.CommandButton cmdWksnButtons 
               Caption         =   "_"
               Height          =   615
               Index           =   8
               Left            =   2880
               Picture         =   "frmMain.frx":6D70
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   0
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.CommandButton cmdWksnButtons 
               Caption         =   "_"
               Height          =   615
               Index           =   9
               Left            =   4320
               Picture         =   "frmMain.frx":6EBA
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   0
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.CommandButton cmdWksnButtons 
               Caption         =   "_"
               DisabledPicture =   "frmMain.frx":7004
               Height          =   615
               Index           =   14
               Left            =   2880
               Picture         =   "frmMain.frx":714E
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   0
               Visible         =   0   'False
               Width           =   1395
            End
         End
         Begin VB.Frame fraWksnSessionInfo 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   4095
            Left            =   60
            TabIndex        =   105
            Top             =   780
            Visible         =   0   'False
            Width           =   6375
            Begin VB.Frame fraWksnChat 
               BorderStyle     =   0  'None
               Height          =   1335
               Left            =   0
               TabIndex        =   240
               Top             =   2580
               Width           =   6015
               Begin VB.TextBox txtWksnChatHist 
                  BackColor       =   &H80000018&
                  BorderStyle     =   0  'None
                  Height          =   615
                  Left            =   2400
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   242
                  Top             =   120
                  Width           =   2655
               End
               Begin VB.Frame fraWksnChatEntry 
                  BorderStyle     =   0  'None
                  Height          =   315
                  Left            =   0
                  TabIndex        =   244
                  Top             =   780
                  Width           =   5475
                  Begin VB.CommandButton cmdWksnButtons 
                     Height          =   315
                     Index           =   13
                     Left            =   4680
                     Picture         =   "frmMain.frx":7298
                     Style           =   1  'Graphical
                     TabIndex        =   15
                     Top             =   0
                     Width           =   375
                  End
                  Begin VB.TextBox txtWksnChat 
                     BorderStyle     =   0  'None
                     Height          =   195
                     Left            =   2460
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   14
                     Top             =   60
                     Width           =   2115
                  End
                  Begin VB.Shape shpWksnChatHist 
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H80000010&
                     Height          =   315
                     Left            =   2400
                     Top             =   0
                     Width           =   2235
                  End
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
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
                  Index           =   55
                  Left            =   60
                  TabIndex        =   243
                  Top             =   120
                  Width           =   2295
               End
               Begin VB.Label lbl 
                  Caption         =   "fraWksnChat"
                  Height          =   195
                  Index           =   54
                  Left            =   480
                  TabIndex        =   241
                  Top             =   480
                  Visible         =   0   'False
                  Width           =   1335
               End
            End
            Begin VB.Frame fraWksnSelfserveInfo 
               BorderStyle     =   0  'None
               Height          =   375
               Left            =   0
               TabIndex        =   115
               Top             =   1080
               Width           =   5235
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
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
                  Index           =   6
                  Left            =   60
                  TabIndex        =   117
                  Top             =   120
                  Width           =   2295
               End
               Begin VB.Label lbl 
                  Caption         =   "fraWksnSelfserveInfo"
                  Height          =   195
                  Index           =   49
                  Left            =   2460
                  TabIndex        =   216
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   1935
               End
               Begin VB.Label lblWksnSel 
                  BackColor       =   &H80000018&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   5
                  Left            =   2400
                  TabIndex        =   116
                  Top             =   60
                  Width           =   2715
               End
            End
            Begin VB.Frame fraWksnCliInfo 
               BorderStyle     =   0  'None
               Height          =   1515
               Left            =   0
               TabIndex        =   217
               Top             =   1080
               Width           =   5115
               Begin VB.CommandButton cmdWksnButtons 
                  Height          =   315
                  Index           =   12
                  Left            =   4740
                  Picture         =   "frmMain.frx":73E2
                  Style           =   1  'Graphical
                  TabIndex        =   218
                  Top             =   60
                  Width           =   375
               End
               Begin VB.Label lbl 
                  Caption         =   "fraWksnCliInfo"
                  Height          =   195
                  Index           =   44
                  Left            =   2460
                  TabIndex        =   227
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   1335
               End
               Begin VB.Label lblWksnSel 
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
                  Index           =   10
                  Left            =   2400
                  TabIndex        =   226
                  Top             =   1260
                  Width           =   2715
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
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
                  Index           =   45
                  Left            =   60
                  TabIndex        =   225
                  Top             =   1260
                  Width           =   2295
               End
               Begin VB.Label lblWksnSel 
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
                  Index           =   11
                  Left            =   2400
                  TabIndex        =   224
                  Top             =   900
                  Width           =   2715
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
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
                  Index           =   46
                  Left            =   60
                  TabIndex        =   223
                  Top             =   900
                  Width           =   2295
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
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
                  Index           =   47
                  Left            =   60
                  TabIndex        =   222
                  Top             =   510
                  Width           =   2295
               End
               Begin VB.Label lblWksnSel 
                  BackColor       =   &H80000018&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   12
                  Left            =   2400
                  TabIndex        =   221
                  Top             =   480
                  Width           =   2715
               End
               Begin VB.Label lblWksnSel 
                  BackColor       =   &H80000018&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   13
                  Left            =   2400
                  TabIndex        =   220
                  Top             =   60
                  Width           =   2295
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
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
                  Index           =   48
                  Left            =   60
                  TabIndex        =   219
                  Top             =   60
                  Width           =   2295
               End
            End
            Begin VB.Frame fraWksnCrdInfo 
               BorderStyle     =   0  'None
               Height          =   1515
               Left            =   0
               TabIndex        =   110
               Top             =   1080
               Width           =   5115
               Begin VB.CommandButton cmdWksnButtons 
                  Height          =   315
                  Index           =   11
                  Left            =   4740
                  Picture         =   "frmMain.frx":752C
                  Style           =   1  'Graphical
                  TabIndex        =   20
                  Top             =   60
                  Width           =   375
               End
               Begin VB.Label lbl 
                  Caption         =   "fraWksnCrdInfo"
                  Height          =   195
                  Index           =   43
                  Left            =   2460
                  TabIndex        =   215
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   1335
               End
               Begin VB.Label lblWksnSel 
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
                  Index           =   9
                  Left            =   2400
                  TabIndex        =   168
                  Top             =   1260
                  Width           =   2715
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
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
                  Index           =   29
                  Left            =   0
                  TabIndex        =   167
                  Top             =   1260
                  Width           =   2355
               End
               Begin VB.Label lblWksnSel 
                  BackColor       =   &H80000018&
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
                  Index           =   8
                  Left            =   2400
                  TabIndex        =   166
                  Top             =   900
                  Width           =   2715
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
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
                  Index           =   28
                  Left            =   0
                  TabIndex        =   165
                  Top             =   900
                  Width           =   2355
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
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
                  Index           =   4
                  Left            =   0
                  TabIndex        =   114
                  Top             =   510
                  Width           =   2355
               End
               Begin VB.Label lblWksnSel 
                  BackColor       =   &H80000018&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   7
                  Left            =   2400
                  TabIndex        =   113
                  Top             =   480
                  Width           =   2715
               End
               Begin VB.Label lblWksnSel 
                  BackColor       =   &H80000018&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   6
                  Left            =   2400
                  TabIndex        =   112
                  Top             =   60
                  Width           =   2295
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
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
                  Index           =   3
                  Left            =   60
                  TabIndex        =   111
                  Top             =   90
                  Width           =   2295
               End
            End
            Begin VB.Frame fraWksnSelfservePrepaidInfo 
               BorderStyle     =   0  'None
               Height          =   375
               Left            =   0
               TabIndex        =   247
               Top             =   1080
               Width           =   5235
               Begin VB.Label lbl 
                  Caption         =   "fraWksnSelfservePrepaidInfo"
                  Height          =   195
                  Index           =   57
                  Left            =   2460
                  TabIndex        =   249
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   2415
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
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
                  Index           =   56
                  Left            =   60
                  TabIndex        =   248
                  Top             =   120
                  Width           =   2295
               End
               Begin VB.Label lblWksnSel 
                  BackColor       =   &H80000018&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   14
                  Left            =   2400
                  TabIndex        =   250
                  Top             =   60
                  Width           =   2715
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
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
               Index           =   32
               Left            =   60
               TabIndex        =   185
               Top             =   780
               Width           =   2295
            End
            Begin VB.Label lblWksnSel 
               BackColor       =   &H80000018&
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
               Index           =   4
               Left            =   2400
               TabIndex        =   184
               Top             =   780
               Width           =   2715
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
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
               Index           =   1
               Left            =   60
               TabIndex        =   109
               Top             =   60
               Width           =   2295
            End
            Begin VB.Label lblWksnSel 
               BackColor       =   &H80000018&
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
               Index           =   2
               Left            =   2400
               TabIndex        =   108
               Top             =   60
               Width           =   2715
            End
            Begin VB.Label lblWksnSel 
               BackColor       =   &H80000018&
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
               Index           =   3
               Left            =   2400
               TabIndex        =   107
               Top             =   420
               Width           =   2715
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
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
               Index           =   2
               Left            =   60
               TabIndex        =   106
               Top             =   420
               Width           =   2295
            End
         End
         Begin VB.Shape shpWksnControlsBorder 
            BorderColor     =   &H80000010&
            Height          =   855
            Left            =   7080
            Top             =   5520
            Width           =   1815
         End
         Begin VB.Label lblWksnSel 
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   2460
            TabIndex        =   104
            Top             =   420
            Width           =   2715
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
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
            Index           =   0
            Left            =   60
            TabIndex        =   103
            Top             =   420
            Width           =   2355
         End
         Begin VB.Label lblWksnSel 
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   102
            Top             =   0
            Width           =   6495
         End
      End
      Begin VB.Frame fraWksnConsoleControls 
         BorderStyle     =   0  'None
         Height          =   6435
         Left            =   3420
         TabIndex        =   174
         Top             =   0
         Visible         =   0   'False
         Width           =   8595
         Begin VB.Frame fraWksnConsoleButtons 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   1695
            Left            =   0
            TabIndex        =   175
            Top             =   3720
            Width           =   8595
            Begin VB.CommandButton cmdConsoleButtons 
               Enabled         =   0   'False
               Height          =   315
               Index           =   8
               Left            =   8160
               Picture         =   "frmMain.frx":7676
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   0
               Width           =   375
            End
            Begin VB.OptionButton optConsoleWorkstationsSome 
               Caption         =   "Contrôler certains clients"
               Height          =   315
               Left            =   4320
               Style           =   1  'Graphical
               TabIndex        =   4
               Top             =   0
               Width           =   3795
            End
            Begin VB.OptionButton optConsoleWorkstationsAll 
               Caption         =   "Contrôler tous les clients"
               Height          =   315
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   3
               Top             =   0
               Value           =   -1  'True
               Width           =   4215
            End
            Begin VB.CommandButton cmdConsoleButtons 
               Caption         =   "Verrouiller"
               Height          =   615
               Index           =   6
               Left            =   5160
               Picture         =   "frmMain.frx":77C0
               Style           =   1  'Graphical
               TabIndex        =   12
               Top             =   1080
               Width           =   1695
            End
            Begin VB.CommandButton cmdConsoleButtons 
               Caption         =   "Suivi"
               Height          =   615
               Index           =   7
               Left            =   6900
               Picture         =   "frmMain.frx":790A
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   1080
               Width           =   1695
            End
            Begin VB.CommandButton cmdConsoleButtons 
               Caption         =   "Débloquer"
               Height          =   615
               Index           =   2
               Left            =   1680
               Picture         =   "frmMain.frx":7F14
               Style           =   1  'Graphical
               TabIndex        =   8
               Top             =   420
               Width           =   1695
            End
            Begin VB.CommandButton cmdConsoleButtons 
               Caption         =   "&Fermer"
               DisabledPicture =   "frmMain.frx":805E
               Height          =   615
               Index           =   3
               Left            =   1680
               Picture         =   "frmMain.frx":81A8
               Style           =   1  'Graphical
               TabIndex        =   9
               Top             =   1080
               Width           =   1695
            End
            Begin VB.CommandButton cmdConsoleButtons 
               Caption         =   "&Activer"
               Height          =   615
               Index           =   0
               Left            =   0
               Picture         =   "frmMain.frx":82F2
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   420
               Width           =   1635
            End
            Begin VB.CommandButton cmdConsoleButtons 
               Caption         =   "&Bloquer"
               Height          =   615
               Index           =   1
               Left            =   0
               Picture         =   "frmMain.frx":843C
               Style           =   1  'Graphical
               TabIndex        =   7
               Top             =   1080
               Width           =   1635
            End
            Begin VB.CommandButton cmdConsoleButtons 
               Caption         =   "&Reprendre"
               Height          =   615
               Index           =   4
               Left            =   3420
               Picture         =   "frmMain.frx":8586
               Style           =   1  'Graphical
               TabIndex        =   10
               Top             =   420
               Width           =   5175
            End
            Begin VB.CommandButton cmdConsoleButtons 
               Caption         =   "&Suspendre"
               Height          =   615
               Index           =   5
               Left            =   3420
               Picture         =   "frmMain.frx":86D0
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   1080
               Width           =   1695
            End
         End
         Begin VB.Frame fraWksnConsoleWaitList 
            BorderStyle     =   0  'None
            Height          =   2055
            Left            =   60
            TabIndex        =   228
            Top             =   1500
            Visible         =   0   'False
            Width           =   5295
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
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
               Index           =   50
               Left            =   0
               TabIndex        =   229
               Top             =   0
               Width           =   2355
            End
         End
         Begin VB.Label lblConsoleSel 
            BackColor       =   &H80000018&
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
            Index           =   1
            Left            =   2460
            TabIndex        =   182
            Top             =   420
            Width           =   2715
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
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
            Index           =   31
            Left            =   60
            TabIndex        =   181
            Top             =   420
            Width           =   2355
         End
         Begin VB.Label lblConsoleSel 
            BackColor       =   &H80000018&
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
            Index           =   3
            Left            =   2460
            TabIndex        =   180
            Top             =   1140
            Width           =   2715
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
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
            Index           =   30
            Left            =   60
            TabIndex        =   179
            Top             =   1140
            Width           =   2355
         End
         Begin VB.Shape shpWksnConsoleControlsBorder 
            BorderColor     =   &H80000010&
            Height          =   855
            Left            =   4380
            Top             =   5460
            Width           =   1815
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
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
            Index           =   37
            Left            =   60
            TabIndex        =   177
            Top             =   780
            Width           =   2355
         End
         Begin VB.Label lblConsoleSel 
            BackColor       =   &H80000018&
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
            Index           =   2
            Left            =   2460
            TabIndex        =   176
            Top             =   780
            Width           =   2715
         End
         Begin VB.Label lblConsoleSel 
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   178
            Top             =   0
            Width           =   6495
         End
      End
   End
   Begin VB.Frame fraFunctions 
      BorderStyle     =   0  'None
      Height          =   6795
      Index           =   5
      Left            =   60
      TabIndex        =   118
      Top             =   480
      Width           =   12255
      Begin VB.CheckBox chkSessStatType 
         Caption         =   "_"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   236
         Top             =   4320
         Value           =   1  'Checked
         Width           =   2595
      End
      Begin VB.CheckBox chkSessStatType 
         Caption         =   "_"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   235
         Top             =   4080
         Value           =   1  'Checked
         Width           =   2595
      End
      Begin VB.CheckBox chkSessStatType 
         Caption         =   "_"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   234
         Top             =   3840
         Value           =   1  'Checked
         Width           =   2595
      End
      Begin VB.CheckBox chkSessStatType 
         Caption         =   "_"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   233
         Top             =   3600
         Value           =   1  'Checked
         Width           =   2595
      End
      Begin VB.CheckBox chkSessStatType 
         Caption         =   "_"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   231
         Top             =   3360
         Value           =   1  'Checked
         Width           =   2595
      End
      Begin VB.Frame fraSessionResults 
         BorderStyle     =   0  'None
         Height          =   1515
         Index           =   4
         Left            =   3060
         TabIndex        =   186
         Top             =   3360
         Visible         =   0   'False
         Width           =   5115
         Begin VB.Frame fraSessionResultsURLSearch 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   60
            TabIndex        =   188
            Top             =   1020
            Width           =   4275
            Begin VB.TextBox txtSessionURLSearch 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               HideSelection   =   0   'False
               Left            =   1020
               MaxLength       =   50
               TabIndex        =   190
               Top             =   45
               Width           =   2355
            End
            Begin VB.CommandButton cmdSessionURLSearch 
               Default         =   -1  'True
               Height          =   315
               Left            =   3480
               Picture         =   "frmMain.frx":881A
               Style           =   1  'Graphical
               TabIndex        =   189
               Top             =   0
               Width           =   375
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "_"
               Height          =   195
               Index           =   33
               Left            =   0
               TabIndex        =   191
               Top             =   60
               Width           =   915
            End
            Begin VB.Shape shpLine 
               BackStyle       =   1  'Opaque
               BorderColor     =   &H80000010&
               Height          =   315
               Index           =   2
               Left            =   960
               Top             =   0
               Width           =   2475
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid fxgSession 
            Height          =   675
            Index           =   4
            Left            =   180
            TabIndex        =   80
            Top             =   120
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   1191
            _Version        =   393216
            Rows            =   10
            Cols            =   1000
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
            SelectionMode   =   1
            AllowUserResizing=   1
            BorderStyle     =   0
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   1000
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Line linSessionResultsURLs 
            BorderColor     =   &H80000010&
            X1              =   0
            X2              =   4860
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Shape shpSessionResultsURLsBorder 
            BorderColor     =   &H80000010&
            Height          =   1440
            Left            =   0
            Top             =   0
            Width           =   4875
         End
      End
      Begin VB.Frame fraSessionResults 
         BorderStyle     =   0  'None
         Height          =   915
         Index           =   5
         Left            =   3060
         TabIndex        =   158
         Top             =   4920
         Visible         =   0   'False
         Width           =   5115
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid fxgSession 
            Height          =   675
            Index           =   5
            Left            =   180
            TabIndex        =   81
            Top             =   120
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   1191
            _Version        =   393216
            Rows            =   10
            Cols            =   288
            BackColorFixed  =   -2147483628
            BackColorBkg    =   -2147483643
            GridColor       =   -2147483643
            GridColorFixed  =   -2147483643
            GridColorUnpopulated=   -2147483643
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   0
            FillStyle       =   1
            GridLinesFixed  =   0
            SelectionMode   =   1
            MergeCells      =   1
            BorderStyle     =   0
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   288
            _Band(0).GridLinesBand=   0
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Shape shpSessionLines 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000005&
            Height          =   495
            Left            =   4380
            Top             =   120
            Width           =   495
         End
         Begin VB.Shape shpSessionResultsLinesBorder 
            BorderColor     =   &H80000010&
            Height          =   600
            Left            =   0
            Top             =   0
            Width           =   1155
         End
      End
      Begin VB.Frame fraSessionResults 
         BorderStyle     =   0  'None
         Height          =   915
         Index           =   3
         Left            =   3060
         TabIndex        =   157
         Top             =   2580
         Visible         =   0   'False
         Width           =   5115
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid fxgSession 
            Height          =   675
            Index           =   3
            Left            =   180
            TabIndex        =   79
            Top             =   60
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   1191
            _Version        =   393216
            Rows            =   10
            Cols            =   1000
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
            SelectionMode   =   1
            AllowUserResizing=   1
            BorderStyle     =   0
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   1000
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Shape shpSessionResultsTimesBorder 
            BorderColor     =   &H80000010&
            Height          =   660
            Left            =   0
            Top             =   0
            Width           =   1155
         End
      End
      Begin VB.OptionButton optSessionFilter 
         Caption         =   "_"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   64
         Top             =   780
         Width           =   2595
      End
      Begin VB.OptionButton optSessionFilter 
         Caption         =   "_"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   63
         Top             =   540
         Value           =   -1  'True
         Width           =   2595
      End
      Begin MSComctlLib.ImageList imlSessionTabs 
         Left            =   10920
         Top             =   60
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483633
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":8964
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":8ABE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":8C18
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":8D72
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":8ECC
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.OptionButton optSessionFilter 
         Caption         =   "_"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   71
         Top             =   2220
         Width           =   1275
      End
      Begin VB.OptionButton optSessionFilter 
         Caption         =   "_"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   69
         Top             =   1980
         Width           =   2595
      End
      Begin VB.OptionButton optSessionFilter 
         Caption         =   "_"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   66
         Top             =   1260
         Width           =   2595
      End
      Begin VB.OptionButton optSessionFilter 
         Caption         =   "_"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   68
         Top             =   1740
         Width           =   2595
      End
      Begin VB.OptionButton optSessionFilter 
         Caption         =   "_"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   67
         Top             =   1500
         Width           =   2595
      End
      Begin VB.OptionButton optSessionFilter 
         Caption         =   "_"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   65
         Top             =   1020
         Width           =   2595
      End
      Begin VB.Frame fraSessionResults 
         BorderStyle     =   0  'None
         Height          =   1035
         Index           =   2
         Left            =   3060
         TabIndex        =   155
         Top             =   1620
         Visible         =   0   'False
         Width           =   5115
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid fxgSession 
            Height          =   735
            Index           =   2
            Left            =   180
            TabIndex        =   78
            Top             =   120
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   1296
            _Version        =   393216
            Rows            =   10
            Cols            =   10
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
            SelectionMode   =   1
            AllowUserResizing=   1
            BorderStyle     =   0
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   10
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Shape shpSessionResultsListBorder 
            BorderColor     =   &H80000010&
            Height          =   600
            Left            =   0
            Top             =   0
            Width           =   1155
         End
      End
      Begin VB.Frame fraSessionResults 
         BorderStyle     =   0  'None
         Height          =   1215
         Index           =   1
         Left            =   3060
         TabIndex        =   154
         Top             =   420
         Width           =   5835
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid fxgSession 
            Height          =   735
            Index           =   1
            Left            =   120
            TabIndex        =   77
            Top             =   180
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   1296
            _Version        =   393216
            Rows            =   10
            Cols            =   100
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
            SelectionMode   =   1
            AllowUserResizing=   1
            BorderStyle     =   0
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   100
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Shape shpSessionResultsStatsBorder 
            BorderColor     =   &H80000010&
            Height          =   840
            Left            =   0
            Top             =   0
            Width           =   1155
         End
      End
      Begin MSComctlLib.TabStrip tbsSessionResultType 
         Height          =   375
         Left            =   2820
         TabIndex        =   76
         Top             =   0
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   661
         Style           =   2
         HotTracking     =   -1  'True
         ImageList       =   "imlSessionTabs"
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   5
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
         EndProperty
      End
      Begin VB.Frame fraSessionControls 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1275
         Left            =   60
         TabIndex        =   153
         Top             =   4680
         Width           =   2655
         Begin VB.CommandButton cmdSessionPrint 
            Caption         =   "_"
            Enabled         =   0   'False
            Height          =   615
            Left            =   1380
            Picture         =   "frmMain.frx":9026
            Style           =   1  'Graphical
            TabIndex        =   245
            Top             =   660
            Width           =   1275
         End
         Begin MSComDlg.CommonDialog cdlSessionExport 
            Left            =   60
            Top             =   720
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CommandButton cmdSessionFilter 
            Caption         =   "_"
            Height          =   615
            Left            =   0
            Picture         =   "frmMain.frx":9170
            Style           =   1  'Graphical
            TabIndex        =   74
            Top             =   0
            Width           =   2655
         End
         Begin VB.CommandButton cmdSessionExport 
            Caption         =   "_"
            Enabled         =   0   'False
            Height          =   615
            Left            =   0
            Picture         =   "frmMain.frx":92BA
            Style           =   1  'Graphical
            TabIndex        =   75
            Top             =   660
            Width           =   1275
         End
      End
      Begin MSComCtl2.DTPicker dtpSessionEndDate 
         Height          =   315
         Left            =   1500
         TabIndex        =   73
         Top             =   2760
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   49807361
         CurrentDate     =   38196
      End
      Begin MSComCtl2.DTPicker dtpSessionStartDate 
         Height          =   315
         Left            =   1500
         TabIndex        =   72
         Top             =   2400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   49807361
         CurrentDate     =   38196
      End
      Begin VB.OptionButton optSessionFilter 
         Caption         =   "_"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   70
         Top             =   2100
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.Label lbl 
         Caption         =   "_"
         Height          =   195
         Index           =   51
         Left            =   120
         TabIndex        =   232
         Top             =   3120
         Width           =   2595
      End
      Begin VB.Label lbl 
         Caption         =   "_"
         Height          =   195
         Index           =   27
         Left            =   120
         TabIndex        =   156
         Top             =   300
         Width           =   2595
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   26
         Left            =   60
         TabIndex        =   152
         Top             =   2460
         Width           =   1395
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   25
         Left            =   120
         TabIndex        =   151
         Top             =   2820
         Width           =   1335
      End
      Begin VB.Shape shpSessionControlsBorder 
         BorderColor     =   &H80000010&
         Height          =   6000
         Left            =   0
         Top             =   0
         Width           =   2775
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         Caption         =   "Filtres"
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
         Index           =   24
         Left            =   0
         TabIndex        =   150
         Top             =   0
         Width           =   2775
      End
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Etat:"
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
      Index           =   8
      Left            =   3780
      TabIndex        =   126
      Top             =   1200
      Width           =   1875
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLogoff 
         Caption         =   "Logoff"
         Shortcut        =   +{F4}
      End
      Begin VB.Menu mnuMinimizeToSysTray 
         Caption         =   "MinimizeToSysTray"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuMode 
      Caption         =   "Mode"
      Begin VB.Menu mnuFunctions 
         Caption         =   "Functions1"
         Checked         =   -1  'True
         Index           =   1
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuFunctions 
         Caption         =   "Functions2"
         Index           =   2
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuFunctions 
         Caption         =   "Functions3"
         Index           =   3
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuFunctions 
         Caption         =   "Functions4"
         Index           =   4
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuFunctions 
         Caption         =   "Functions5"
         Index           =   5
      End
      Begin VB.Menu mnuFunctions 
         Caption         =   "Functions6"
         Index           =   6
      End
   End
   Begin VB.Menu mnuSubFunctions 
      Caption         =   "SubFunctions1"
      Index           =   1
      Begin VB.Menu mnuWksnShowConsole 
         Caption         =   "mnuWksnShowConsole"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuWksnShowPrev 
         Caption         =   "mnuWksnShowPrev"
         Shortcut        =   +{F2}
      End
      Begin VB.Menu mnuWksnShowNext 
         Caption         =   "mnuWksnShowNext"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnu6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWksnActivate 
         Caption         =   "mnuWksnActivate"
      End
      Begin VB.Menu mnuWksnBloc 
         Caption         =   "mnuWksnBloc"
      End
      Begin VB.Menu mnu0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWksnPaySession 
         Caption         =   "mnuWksnPaySession"
      End
      Begin VB.Menu mnuWksnCancelPaySession 
         Caption         =   "mnuWksnCancelPaySession"
      End
      Begin VB.Menu mnuWksnUnblock 
         Caption         =   "mnuWksnUnblock"
      End
      Begin VB.Menu mnuWksnStartSelfServPrePaid 
         Caption         =   "mnuWksnStartSelfServPrePaid"
      End
      Begin VB.Menu mnuWksnClose 
         Caption         =   "mnuWksnClose"
      End
      Begin VB.Menu mnuWksnPause 
         Caption         =   "mnuWksnPause"
      End
      Begin VB.Menu mnuWksnLock 
         Caption         =   "mnuWksnLock"
      End
      Begin VB.Menu mnuWksnFollow 
         Caption         =   "mnuWksnFollow"
      End
      Begin VB.Menu mnuWksnPlay 
         Caption         =   "mnuWksnPlay"
      End
      Begin VB.Menu mnuWksnCrdRecharge 
         Caption         =   "mnuWksnCrdRecharge"
      End
      Begin VB.Menu mnuWksnCrdDisable 
         Caption         =   "mnuWksnCrdDisable"
      End
      Begin VB.Menu mnuWksnCrdHistory 
         Caption         =   "mnuWksnCrdHistory"
      End
      Begin VB.Menu mnuWksnConnect 
         Caption         =   "mnuWksnConnect"
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWksnActivateConsole 
         Caption         =   "mnuWksnActivateConsole"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuWksnSetSize 
         Caption         =   "mnuWksnSetSize"
         Begin VB.Menu mnuWksnSize 
            Caption         =   "mnuWksnSize"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnuWksnSize 
            Caption         =   "mnuWksnSize"
            Index           =   2
         End
      End
      Begin VB.Menu mnuWksnShutdown 
         Caption         =   "mnuWksnShutdown"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mnuWksnBoot 
         Caption         =   "mnuWksnBoot"
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu mnuWksnUserRelease 
         Caption         =   "mnuWksnUserRelease"
      End
   End
   Begin VB.Menu mnuSubFunctions 
      Caption         =   "SubFunctions2"
      Index           =   2
      Begin VB.Menu mnuCrd 
         Caption         =   "CrdNew"
         Index           =   0
      End
      Begin VB.Menu mnuCrd 
         Caption         =   "CrdRecharge"
         Index           =   1
      End
      Begin VB.Menu mnuCrd 
         Caption         =   "CrdSave"
         Index           =   2
      End
      Begin VB.Menu mnuCrd 
         Caption         =   "CrdEnable"
         Index           =   3
      End
      Begin VB.Menu mnuCrd 
         Caption         =   "CrdDisable"
         Index           =   4
      End
      Begin VB.Menu mnuCrd 
         Caption         =   "CrdHistory"
         Index           =   5
      End
      Begin VB.Menu mnuCrd 
         Caption         =   "CrdStartSession"
         Index           =   6
      End
      Begin VB.Menu mnuCrd 
         Caption         =   "CrdDelete"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCrd 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuCrd 
         Caption         =   "CrdFillList"
         Index           =   9
      End
      Begin VB.Menu mnuCrd 
         Caption         =   "CrdMassAdd"
         Index           =   10
      End
      Begin VB.Menu mnuCrd 
         Caption         =   "CrdPurge"
         Index           =   11
      End
   End
   Begin VB.Menu mnuSubFunctions 
      Caption         =   "SubFunctions3"
      Index           =   3
      Begin VB.Menu mnuCli 
         Caption         =   "CliNew"
         Index           =   0
      End
      Begin VB.Menu mnuCli 
         Caption         =   "CliRecharge"
         Index           =   1
      End
      Begin VB.Menu mnuCli 
         Caption         =   "CliModify"
         Index           =   2
      End
      Begin VB.Menu mnuCli 
         Caption         =   "CliDisable"
         Index           =   3
      End
      Begin VB.Menu mnuCli 
         Caption         =   "CliEnable"
         Index           =   4
      End
      Begin VB.Menu mnuCli 
         Caption         =   "CliHistory"
         Index           =   5
      End
      Begin VB.Menu mnuCli 
         Caption         =   "CliStartSession"
         Index           =   6
      End
      Begin VB.Menu mnuCli 
         Caption         =   "CliDelete"
         Index           =   7
      End
      Begin VB.Menu mnuCli 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuCli 
         Caption         =   "CliFillList"
         Index           =   9
      End
      Begin VB.Menu mnuCli 
         Caption         =   "CliPurge"
         Index           =   10
      End
   End
   Begin VB.Menu mnuSubFunctions 
      Caption         =   "SubFunctions4"
      Index           =   4
      Begin VB.Menu mnuServicesCollect 
         Caption         =   "mnuServicesCollect"
      End
   End
   Begin VB.Menu mnuSubFunctions 
      Caption         =   "SubFunctions5"
      Index           =   5
      Begin VB.Menu mnuSessionApply 
         Caption         =   "mnuSessionApply"
      End
      Begin VB.Menu mnu3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSessionPrint 
         Caption         =   "mnuSessionPrint"
      End
      Begin VB.Menu mnuSessionExport 
         Caption         =   "mnuSessionExport"
      End
   End
   Begin VB.Menu mnuSubFunctions 
      Caption         =   "SubFunctions6"
      Index           =   6
      Begin VB.Menu mnuIncomeApply 
         Caption         =   "mnuIncomeApply"
      End
      Begin VB.Menu mnu4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIncomePrint 
         Caption         =   "mnuIncomePrint"
      End
      Begin VB.Menu mnuIncomeExport 
         Caption         =   "mnuIncomeExport"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuChangePasswdUser 
         Caption         =   "ChangePasswdUser"
      End
      Begin VB.Menu mnu8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuUserAdmin 
         Caption         =   "UserAdmin"
      End
      Begin VB.Menu mnuChangePasswd 
         Caption         =   "ChangePasswd"
      End
      Begin VB.Menu mnuDBMaintenance 
         Caption         =   "DBMaintenance"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuCrdStartSessionDropdown 
      Caption         =   "CrdStartSessionDropdown"
      Visible         =   0   'False
      Begin VB.Menu mnuCrdStartSessionFirstCilent 
         Caption         =   "CrdStartSessionFirstCilent"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCrdStartSessionWorkstation 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuCrdStartSessionWorkstation 
         Caption         =   "CrdStartSessionWorkstation1"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mnuCrdStartSessionWorkstation 
         Caption         =   "_"
         Enabled         =   0   'False
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCrdStartSessionWorkstation 
         Caption         =   "_"
         Enabled         =   0   'False
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCrdStartSessionWorkstation 
         Caption         =   "_"
         Enabled         =   0   'False
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCrdStartSessionWorkstation 
         Caption         =   "_"
         Enabled         =   0   'False
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCrdStartSessionWorkstation 
         Caption         =   "_"
         Enabled         =   0   'False
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCrdStartSessionWorkstation 
         Caption         =   "_"
         Enabled         =   0   'False
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCrdStartSessionWorkstation 
         Caption         =   "_"
         Enabled         =   0   'False
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCrdStartSessionWorkstation 
         Caption         =   "_"
         Enabled         =   0   'False
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCrdStartSessionWorkstation 
         Caption         =   "_"
         Enabled         =   0   'False
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCrdStartSessionWorkstation 
         Caption         =   "_"
         Enabled         =   0   'False
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCrdStartSessionWorkstation 
         Caption         =   "_"
         Enabled         =   0   'False
         Index           =   12
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCrdStartSessionWorkstation 
         Caption         =   "_"
         Enabled         =   0   'False
         Index           =   13
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCrdStartSessionWorkstation 
         Caption         =   "_"
         Enabled         =   0   'False
         Index           =   14
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCrdStartSessionWorkstation 
         Caption         =   "_"
         Enabled         =   0   'False
         Index           =   15
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCrdStartSessionWorkstation 
         Caption         =   "_"
         Enabled         =   0   'False
         Index           =   16
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCrdStartSessionWorkstation 
         Caption         =   "_"
         Enabled         =   0   'False
         Index           =   17
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCrdStartSessionWorkstation 
         Caption         =   "_"
         Enabled         =   0   'False
         Index           =   18
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCrdStartSessionWorkstation 
         Caption         =   "_"
         Enabled         =   0   'False
         Index           =   19
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCrdStartSessionWorkstation 
         Caption         =   "_"
         Enabled         =   0   'False
         Index           =   20
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuWksnPauseDropDown 
      Caption         =   "WksnPauseDropDown"
      Visible         =   0   'False
      Begin VB.Menu mnuWksnPauseOne 
         Caption         =   "WksnPauseOne"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuWksnLockOne 
         Caption         =   "WksnLockOne"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuWksnFollowOne 
         Caption         =   "WksnFollowOne"
      End
   End
   Begin VB.Menu mnuSessionURLDropDown 
      Caption         =   "SessionURLDropDown"
      Visible         =   0   'False
      Begin VB.Menu mnuSessionURLCopy 
         Caption         =   "SessionURLCopy"
      End
      Begin VB.Menu mnuSessionURLOpen 
         Caption         =   "SessionURLOpen"
      End
      Begin VB.Menu mnu7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSessionCardNumCopy 
         Caption         =   "SessionCardNumCopy"
      End
   End
   Begin VB.Menu mnuWksnDropDown 
      Caption         =   "WksnDropDown"
      Visible         =   0   'False
      Begin VB.Menu mnuWksnActivateDD 
         Caption         =   "WksnActivateDD"
      End
      Begin VB.Menu mnuWksnBlocDD 
         Caption         =   "WksnBlocDD"
      End
      Begin VB.Menu mnu9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWksnPaySessionDD 
         Caption         =   "WksnPaySessionDD"
      End
      Begin VB.Menu mnuWksnUnblockDD 
         Caption         =   "WksnUnblockDD"
      End
      Begin VB.Menu mnuWksnStartSelfServPrePaidDD 
         Caption         =   "mnuWksnStartSelfServPrePaidDD"
      End
      Begin VB.Menu mnuWksnCloseDD 
         Caption         =   "WksnCloseDD"
      End
      Begin VB.Menu mnuWksnPauseDD 
         Caption         =   "WksnPauseDD"
      End
      Begin VB.Menu mnuWksnLockDD 
         Caption         =   "WksnLockDD"
      End
      Begin VB.Menu mnuWksnFollowDD 
         Caption         =   "WksnFollowDD"
      End
      Begin VB.Menu mnuWksnPlayDD 
         Caption         =   "WksnPlayDD"
      End
      Begin VB.Menu mnuWksnCrdRechargeDD 
         Caption         =   "WksnCrdRechargeDD"
      End
      Begin VB.Menu mnuWksnCrdDisableDD 
         Caption         =   "WksnCrdDisableDD"
      End
      Begin VB.Menu mnuWksnCrdHistoryDD 
         Caption         =   "WksnCrdHistoryDD"
      End
      Begin VB.Menu mnu10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWksnBootDD 
         Caption         =   "WksnBootDD"
      End
      Begin VB.Menu mnuWksnShutdownDD 
         Caption         =   "WksnShutdownDD"
      End
      Begin VB.Menu mnuWksnUserReleaseDD 
         Caption         =   "mnuWksnUserReleaseDD"
      End
   End
   Begin VB.Menu mnuCrdDropDown 
      Caption         =   "CrdDropDown"
      Visible         =   0   'False
      Begin VB.Menu mnuCrdDD 
         Caption         =   "CrdDD0"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCrdDD 
         Caption         =   "CrdRechargeDD"
         Index           =   1
      End
      Begin VB.Menu mnuCrdDD 
         Caption         =   "CrdDD2"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCrdDD 
         Caption         =   "CrdEnableDD"
         Index           =   3
      End
      Begin VB.Menu mnuCrdDD 
         Caption         =   "CrdDisableDD"
         Index           =   4
      End
      Begin VB.Menu mnuCrdDD 
         Caption         =   "CrdHistoryDD"
         Index           =   5
      End
      Begin VB.Menu mnuCrdDD 
         Caption         =   "CrdStartSessionDD"
         Index           =   6
      End
      Begin VB.Menu mnuCrdDD 
         Caption         =   "CrdDeleteDD"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCrdDD 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuCrdDD 
         Caption         =   "CrdFillListDD"
         Index           =   9
      End
   End
   Begin VB.Menu mnuCliDropDown 
      Caption         =   "CliDropDown"
      Visible         =   0   'False
      Begin VB.Menu mnuCliDD 
         Caption         =   "CliDD0"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCliDD 
         Caption         =   "CliRechargeDD"
         Index           =   1
      End
      Begin VB.Menu mnuCliDD 
         Caption         =   "CliModifyDD"
         Index           =   2
      End
      Begin VB.Menu mnuCliDD 
         Caption         =   "CliEnableDD"
         Index           =   3
      End
      Begin VB.Menu mnuCliDD 
         Caption         =   "CliDisableDD"
         Index           =   4
      End
      Begin VB.Menu mnuCliDD 
         Caption         =   "CliHistoryDD"
         Index           =   5
      End
      Begin VB.Menu mnuCliDD 
         Caption         =   "CliStartSessionDD"
         Index           =   6
      End
      Begin VB.Menu mnuCliDD 
         Caption         =   "CliDeleteDD"
         Index           =   7
      End
      Begin VB.Menu mnuCliDD 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuCliDD 
         Caption         =   "CliFillListDD"
         Index           =   9
      End
   End
   Begin VB.Menu mnuCliModifyDropDown 
      Caption         =   "CliModifyDropDown"
      Visible         =   0   'False
      Begin VB.Menu mnuCliModifyModifyDD 
         Caption         =   "CliModifyModifyDD"
      End
      Begin VB.Menu mnuCliModifyDeleteDD 
         Caption         =   "CliModifyDeleteDD"
      End
   End
   Begin VB.Menu mnuSysTray 
      Caption         =   "SysTray"
      Visible         =   0   'False
      Begin VB.Menu mnuSysTrayShow 
         Caption         =   "SysTrayShow"
      End
      Begin VB.Menu mnuSysTrayHide 
         Caption         =   "SysTrayHide"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ MAIN PROGRAM FORM __________________________|
'|  Name:          frmMain.frm                                                 |
'|  Type:          VB6 Form                                                    |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
'|  Description:   Main form of Cybera Server                                  |
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

' $Id: frmMain.frm 41 2006-02-21 18:14:29Z sam $

Option Explicit
Dim mblnSessionFiltered As Boolean
Dim mblnLogoffOnShow As Boolean
Dim mblntxtCrdPendingHasFocus As Boolean
Dim mblnfxgCrdPendingHasFocus As Boolean

'*******************************************************************************
Private Sub chkWksnConnect_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo chkWksnConnect_ClickError

    Call subToggleForceConnected
    Exit Sub

chkWksnConnect_ClickError:
    Call funErrorManager("frmMain.chkWksnConnect_MouseUp", True)
End Sub


'*******************************************************************************
Private Sub chkTypeIncome_Click(Index As Integer)
Dim chkTmp As CheckBox
Dim blnFActivateFilter As Boolean
On Error GoTo chkTypeIncome_ClickError

    For Each chkTmp In chkTypeIncome
        If chkTmp.value = vbChecked Then blnFActivateFilter = True
    Next
    cmdIncomeFilter.enabled = blnFActivateFilter
    mnuIncomeApply.enabled = blnFActivateFilter
'    mnuIncomeExport.enabled = blnFActivateFilter
'    mnuIncomePrint.enabled = blnFActivateFilter
    Exit Sub
    
chkTypeIncome_ClickError:
    Call funErrorManager("frmMain.chkTypeIncome_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmbCrdType_Change()
On Error GoTo cmbCrdType_ChangeError

    tmrCrdRefresh.enabled = False
    Exit Sub
    
cmbCrdType_ChangeError:
    Call funErrorManager("frmMain.cmbCrdType_ChangeError", True)
End Sub


'*******************************************************************************
Private Sub cmbCrdType_GotFocus()
On Error GoTo cmbCrdType_GotFocusError

    tmrCrdRefresh.enabled = False
    Exit Sub
    
cmbCrdType_GotFocusError:
    Call funErrorManager("frmMain.cmbCrdType_GotFocusError", True)
End Sub


'*******************************************************************************
Private Sub cmdCliButtons_Click(Index As Integer)
Dim i As Integer
Dim intWorkstationIndex As Integer
Dim uSelectedClient As clsClient
Dim uClient As New clsClient
Dim uWorkstation As clsWorkstation
Dim strBuffer As String
On Error GoTo cmdCliButtons_ClickError

    Select Case Index
    Case 0
        Call funSetFocus(fxgClients)
        cmdCliButtons(0).enabled = False
        mnuCli(0).enabled = False
        Call Load(frmClientConfig)
        Call frmClientConfig.subSetModify(False)
        Call frmClientConfig.Show(vbModal, frmMain)
        cmdCliButtons(0).enabled = True
        mnuCli(0).enabled = True
    Case 1
        Call funSetFocus(fxgClients)
        cmdCliButtons(1).enabled = False
        mnuCli(1).enabled = False
        If guServOpts.lngSelectedClientIndex = 0 Then
            Call funMsg(enuMsg.ClientNotSelected)
        Else
            intWorkstationIndex = funClientInUse(guClients(guServOpts.lngSelectedClientIndex), True)
            If intWorkstationIndex = 0 Then
                Call subFillClients
                Set guSelectedClient = guClients(guServOpts.lngSelectedClientIndex)
            Else
                Set guSelectedClient = guWorkstations(intWorkstationIndex).CurrentClient
            End If
            If guSelectedClient.Active Then
                Call Load(frmRecharge)
                Call frmRecharge.subSetRechargeType(enuCustomerType.Client)
                Call frmRecharge.Show(vbModal, frmMain)
                Call subFillClients
                Call subUpdateUISelectedClient
            Else
                Call funMsg(enuMsg.ClientDisabled)
            End If
        End If
        cmdCliButtons(1).enabled = True 'Not guSelectedClient.UnlimitedPlay
        mnuCli(1).enabled = True 'Not guSelectedClient.UnlimitedPlay
    Case 2
        Call funSetFocus(fxgClients)
        cmdCliButtons(2).enabled = False
        If guServOpts.lngSelectedClientIndex = 0 Then
            Call funMsg(enuMsg.ClientNotSelected)
        Else
            intWorkstationIndex = funClientInUse(guClients(guServOpts.lngSelectedClientIndex))
            If intWorkstationIndex = 0 Then
                Call subFillClients
                Set guSelectedClient = guClients(guServOpts.lngSelectedClientIndex)
            Else
                Set guSelectedClient = guWorkstations(intWorkstationIndex).CurrentClient
            End If
            If guSelectedClient.Active Then
                Call Load(frmClientConfig)
                Call frmClientConfig.subSetModify(True)
                Call frmClientConfig.Show(vbModal, frmMain)
                Call subFillClients
                Call subUpdateUISelectedClient
            Else
                Call funMsg(enuMsg.ClientDisabled)
            End If
        End If
        cmdCliButtons(2).enabled = True
    Case 3
        If gblnDBOK Then
            If guServOpts.lngSelectedClientIndex = 0 Then
                Call funMsg(enuMsg.ClientNotSelected)
            Else
                Set uSelectedClient = guClients(guServOpts.lngSelectedClientIndex)
                If funMsg(AskCanActivateClient, uSelectedClient.login) = vbYes Then
                    grstClients.Index = "PrimaryKey"
                    Call grstClients.Seek("=", uSelectedClient.login)
                    If grstClients.NoMatch Then
                        Exit Sub
                    End If
                    Call grstClients.Edit
                    uSelectedClient.Active = True
                    grstClients!Active = uSelectedClient.Active
                    Call grstClients.Update
                    Call subUpdateUISelectedClient
                End If
            End If
        End If
        Call funSetFocus(fxgClients)
    Case 4
        If gblnDBOK Then
            If guServOpts.lngSelectedClientIndex = 0 Then
                Call funMsg(enuMsg.ClientNotSelected)
            Else
                Set uSelectedClient = guClients(guServOpts.lngSelectedClientIndex)
                If funMsg(AskCanDeactivateClient, uSelectedClient.login) = vbYes Then
                    grstClients.Index = "PrimaryKey"
                    Call grstClients.Seek("=", uSelectedClient.login)
                    If grstClients.NoMatch Then
                        Exit Sub
                    End If
                    Call grstClients.Edit
                    uSelectedClient.Active = False
                    grstClients!Active = uSelectedClient.Active
                    Call grstClients.Update
                    Call subUpdateUISelectedClient
                    intWorkstationIndex = funClientInUse(uSelectedClient)
                    If intWorkstationIndex <> 0 Then
                        If funMsg(enuMsg.AskCanCloseSession) = vbYes Then
                            strBuffer = CNX_SESSION_CLOSE & CNX_MSG_END
                            Call funSendData(guWorkstations(intWorkstationIndex).Cnx, strBuffer)
                        End If
                    End If
                End If
            End If
        End If
        Call funSetFocus(fxgClients)
    Case 5
        Call funSetFocus(fxgClients)
        cmdCliButtons(5).enabled = False
        mnuCli(5).enabled = False
        If guServOpts.lngSelectedClientIndex = 0 Then
            Call funMsg(enuMsg.ClientNotSelected)
        Else
            Call subFillClients
            Set guSelectedClient = guClients(guServOpts.lngSelectedClientIndex)
            Call Load(frmCardHistory)
            Call frmCardHistory.subSetHistoryType(enuCustomerType.Client)
            Call frmCardHistory.Show(vbModal, frmMain)
        End If
        cmdCliButtons(5).enabled = True
        mnuCli(5).enabled = True
    Case 6
        If guServOpts.lngSelectedClientIndex = 0 Then
            Call funMsg(enuMsg.ClientNotSelected)
        Else
            Set uClient = guClients(guServOpts.lngSelectedClientIndex)
            If funClientInUse(uClient) <> 0 Then
                Call funMsg(enuMsg.NoRechargeCardInUse)
            Else
                If uClient.Active And (uClient.TimeCredit > NULL_DATE Or uClient.UnlimitedPlay) Then
                    For i = 1 To guMacAddresses.Count
                        Set uWorkstation = funFindWorkstationByMac(guMacAddresses(i))
                        If uWorkstation.State = cliReady Then
                            Call subWorkstationCliSessionOpen(uWorkstation, uClient)
                            Exit For
                        End If
                    Next i
                Else
                    Call funMsg(NoCredit)
                End If
            End If
        End If
        DoEvents
        Call subUpdateUISelectedCard
        Call funSetFocus(fxgClients)
    End Select
    Exit Sub
    
cmdCliButtons_ClickError:
    Select Case Index
    Case 0
        cmdCliButtons(0).enabled = True
        mnuCli(0).enabled = True
    Case 1
        cmdCliButtons(1).enabled = True
        mnuCli(1).enabled = True
    Case 2
        cmdCliButtons(2).enabled = True
    Case 5
        cmdCliButtons(5).enabled = True
        mnuCli(5).enabled = True

    End Select
    Call funErrorManager("frmMain.cmdCliButtons_ClickError - Index: " & Index, True)
End Sub


'*******************************************************************************
Private Sub cmdCliButtons_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo cmdCliButtons_MouseDownError

    Select Case Index
    Case 7
        cmdCliButtons(Index).value = True
        Call funSetFocus(tbsFunctions)
        Call PopupMenu(mnuCrdStartSessionDropdown, vbPopupMenuRightAlign, cmdCliButtons(7).left + cmdCliButtons(7).width + fraCrdControls.left + fraFunctions(2).left, cmdCliButtons(7).top + cmdCliButtons(7).height + fraCrdControls.top + fraFunctions(1).top, mnuCrdStartSessionFirstCilent)
    Case 8
        cmdCliButtons(Index).value = True
        Call funSetFocus(tbsFunctions)
        Call PopupMenu(mnuCliModifyDropDown, vbPopupMenuRightAlign, cmdCliButtons(8).left + cmdCliButtons(8).width + fraCrdControls.left + fraFunctions(2).left, cmdCliButtons(8).top + cmdCliButtons(8).height + fraCrdControls.top + fraFunctions(1).top, mnuCliModifyModifyDD)
    End Select
    Exit Sub

cmdCliButtons_MouseDownError:
    Call funErrorManager("frmMain.cmdCliButtons_MouseDownError - Index: " & Index, True)
End Sub


'*******************************************************************************
Private Sub cmdCliFind_Click()
On Error GoTo cmdCliFind_ClickError
    
    Call funSetFocus(cmdCliFind)
    Call subFindClient(txtCliFind.Text)
    Call funSetFocus(txtCliFind)
    Exit Sub
    
cmdCliFind_ClickError:
    Call funErrorManager("frmMain.cmdCliFind_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdConsoleButtons_Click(Index As Integer)
Dim uWorkstation As New clsWorkstation
Dim strBuffer As String
On Error GoTo cmdConsoleButtons_ClickError

    Call funSetFocus(tbsFunctions)
    Select Case Index
    Case 0
        For Each uWorkstation In guWorkstations
            If uWorkstation.AllowGroupControl Then
                Select Case uWorkstation.State
                Case enuWksnState.Blocked_Deactivated, enuWksnState.NotActive
                    strBuffer = CNX_ACTIVATE_WORKSTATION & CNX_MSG_END
                    Call funSendData(uWorkstation.Cnx, strBuffer)
                End Select
            End If
        Next uWorkstation
    Case 1
        If funMsg(AskCanLockAllWorkstations) = vbYes Then
            For Each uWorkstation In guWorkstations
                If uWorkstation.AllowGroupControl Then
                    Select Case uWorkstation.State
                    Case enuWksnState.NotPresent, enuWksnState.Blocked_Deactivated, enuWksnState.NotActive
                        
                    Case Else
                        strBuffer = CNX_LOCK_WORKSTATION & CNX_MSG_END
                        Call funSendData(uWorkstation.Cnx, strBuffer)
                    End Select
                End If
            Next uWorkstation
        End If
    Case 2
        For Each uWorkstation In guWorkstations
            If uWorkstation.AllowGroupControl Then
                If uWorkstation.State = enuWksnState.Ready Then
                    strBuffer = CNX_SESSION_FREE & CNX_MSG_END
                    Call funSendData(uWorkstation.Cnx, strBuffer)
                    uWorkstation.StartTime = Now
                    uWorkstation.State = enuWksnState.UsedFree
                    Call subSessionOpen(uWorkstation)
                End If
            End If
        Next uWorkstation
    Case 3
        If funMsg(AskCanCloseAllSessions) = vbYes Then
            For Each uWorkstation In guWorkstations
                If uWorkstation.AllowGroupControl Then
                    If uWorkstation.State > enuWksnState.Blocked_SessionEnd Then
                        strBuffer = CNX_SESSION_CLOSE & CNX_MSG_END
                        Call funSendData(uWorkstation.Cnx, strBuffer)
                    End If
                End If
            Next uWorkstation
        End If
    Case 4
        For Each uWorkstation In guWorkstations
            If uWorkstation.AllowGroupControl Then
                Select Case uWorkstation.State
                Case enuWksnState.Blocked_Paused, enuWksnState.Blocked_Locked, enuWksnState.Blocked_follow
                    strBuffer = CNX_UNPAUSE_WORKSTATION & CNX_MSG_END
                    Call funSendData(uWorkstation.Cnx, strBuffer)
                End Select
            End If
        Next uWorkstation
    Case 5
        For Each uWorkstation In guWorkstations
            If uWorkstation.AllowGroupControl Then
                Select Case uWorkstation.State
                Case enuWksnState.UsedConnected, enuWksnState.UsedDisconnected, enuWksnState.Blocked_follow, enuWksnState.Blocked_Locked, enuWksnState.UsedFree
                    If uWorkstation.State = enuWksnState.Blocked_follow Or uWorkstation.State = enuWksnState.Blocked_Locked Or uWorkstation.State = enuWksnState.Blocked_Paused Then
                        strBuffer = CNX_UNPAUSE_WORKSTATION & CNX_MSG_END
                        Call funSendData(uWorkstation.Cnx, strBuffer)
                        DoEvents
                    End If
                    strBuffer = CNX_PAUSE_WORKSTATION & CNX_MSG_END
                    Call funSendData(uWorkstation.Cnx, strBuffer)
                End Select
            End If
        Next uWorkstation
    Case 6
        For Each uWorkstation In guWorkstations
            If uWorkstation.AllowGroupControl Then
                Select Case uWorkstation.State
                Case enuWksnState.UsedConnected, enuWksnState.UsedDisconnected, enuWksnState.UsedFree, enuWksnState.Blocked_follow, enuWksnState.Blocked_Paused
                    If uWorkstation.State = enuWksnState.Blocked_follow Or uWorkstation.State = enuWksnState.Blocked_Locked Or uWorkstation.State = enuWksnState.Blocked_Paused Then
                        strBuffer = CNX_UNPAUSE_WORKSTATION & CNX_MSG_END
                        Call funSendData(uWorkstation.Cnx, strBuffer)
                        DoEvents
                    End If
                    strBuffer = CNX_PAUSE_LOCK_WORKSTATION & CNX_MSG_END
                    Call funSendData(uWorkstation.Cnx, strBuffer)
                End Select
            End If
        Next uWorkstation
    Case 7
        For Each uWorkstation In guWorkstations
            If uWorkstation.AllowGroupControl Then
                Select Case uWorkstation.State
                Case enuWksnState.UsedConnected, enuWksnState.UsedDisconnected, enuWksnState.UsedFree, enuWksnState.Blocked_Paused, enuWksnState.Blocked_Locked
                    If uWorkstation.State = enuWksnState.Blocked_follow Or uWorkstation.State = enuWksnState.Blocked_Locked Or uWorkstation.State = enuWksnState.Blocked_Paused Then
                        strBuffer = CNX_UNPAUSE_WORKSTATION & CNX_MSG_END
                        Call funSendData(uWorkstation.Cnx, strBuffer)
                        DoEvents
                    End If
                    strBuffer = CNX_PAUSE_FOLLOW_WORKSTATION & CNX_MSG_END
                    Call funSendData(uWorkstation.Cnx, strBuffer)
                End Select
            End If
        Next uWorkstation
    Case 8
        Call Load(frmWksnSelectionPopup)
        frmWksnSelectionPopup.top = frmMain.top + (frmMain.height - frmMain.ScaleHeight) - (frmMain.width - frmMain.ScaleWidth) / 2 + fraFunctions(1).top + fraWksnConsoleControls.top + fraWksnConsoleButtons.top + cmdConsoleButtons(8).top + cmdConsoleButtons(8).height
        frmWksnSelectionPopup.left = frmMain.left + (frmMain.width - frmMain.ScaleWidth) / 2 + fraFunctions(1).left + fraWksnConsoleControls.left + fraWksnConsoleButtons.left + cmdConsoleButtons(8).left + cmdConsoleButtons(8).width - frmWksnSelectionPopup.width
        Call frmWksnSelectionPopup.Show
    End Select
    Exit Sub
    
cmdConsoleButtons_ClickError:
    Call funErrorManager("frmMain.cmdConsoleButtons_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdCrdButtons_Click(Index As Integer)
Dim i As Integer
Dim lngNewCardID As Long
Dim intWorkstationIndex As Integer
Dim uSelectedCard As clsCard
Dim uCard As New clsCard
Dim uWorkstation As clsWorkstation
Dim strBuffer As String
Dim strErrorcontext As String
On Error GoTo cmdCrdButtons_ClickError

    Select Case Index
    Case 0
        Call funSetFocus(fxgCards)
        cmdCrdButtons(0).enabled = False
        mnuCrd(0).enabled = False
        lngNewCardID = funNewCardNumber
        Call grstCards.AddNew
        grstCards!ID = lngNewCardID
        grstCards!Active = True
        Call grstCards.Update
        Call subDBAddEvent(enuEvents.CardNew, 0, lngNewCardID, "", 0, funLoadResString(10341), All)
        DoEvents
        Call subFillCards
        Call subFindCard(CStr(lngNewCardID), False)
        Set guSelectedCard = guCards(guServOpts.lngSelectedCardIndex)
        Call Load(frmRecharge)
        Call frmRecharge.subSetRechargeType(enuCustomerType.Card)
        Call frmRecharge.Show(vbModal, frmMain)
        Call subFillCards
        Call subUpdateUISelectedCard
        cmdCrdButtons(0).enabled = True
        mnuCrd(0).enabled = True
    Case 1
        Call funSetFocus(fxgCards)
        cmdCrdButtons(1).enabled = False
        mnuCrd(1).enabled = False
        If guServOpts.lngSelectedCardIndex = 0 Then
            Call funMsg(enuMsg.CardNotSelected)
        Else
            intWorkstationIndex = funCardInUse(guCards(guServOpts.lngSelectedCardIndex), True)
            If intWorkstationIndex = 0 Then
                Call subFillCards
                Set guSelectedCard = guCards(guServOpts.lngSelectedCardIndex)
            Else
                Set guSelectedCard = guWorkstations(intWorkstationIndex).CurrentCard
            End If
            If guSelectedCard.Active Then
                Call Load(frmRecharge)
                Call frmRecharge.subSetRechargeType(enuCustomerType.Card)
                Call frmRecharge.Show(vbModal, frmMain)
                Call subFillCards
                Call subUpdateUISelectedCard
            Else
                Call funMsg(enuMsg.CardDisabled)
            End If
        End If
        cmdCrdButtons(1).enabled = True 'Not guSelectedCard.UnlimitedPlay
        mnuCrd(1).enabled = True 'Not guSelectedCard.UnlimitedPlay
    Case 2
        If gblnDBOK Then
            If guServOpts.lngSelectedCardIndex = 0 Then
                Call funMsg(enuMsg.CardNotSelected)
            Else
                Set uSelectedCard = guCards(guServOpts.lngSelectedCardIndex)
                grstCards.Index = "PrimaryKey"
                Call grstCards.Seek("=", uSelectedCard.ID)
                If grstCards.NoMatch Then
                    Exit Sub
                End If
                Call grstCards.Edit
                uSelectedCard.Comment = txtCrdComment.Text
                uSelectedCard.CardType = IIf(cmbCrdType.ItemData(cmbCrdType.ListIndex) = 1, "", cmbCrdType.List(cmbCrdType.ListIndex))
                grstCards!Comment = uSelectedCard.Comment
                grstCards!Type = uSelectedCard.CardType
                Call grstCards.Update
                intWorkstationIndex = funCardInUse(uSelectedCard)
                If intWorkstationIndex <> 0 Then
                    guWorkstations(intWorkstationIndex).CurrentCard.Comment = uSelectedCard.Comment
                    guWorkstations(intWorkstationIndex).CurrentCard.CardType = uSelectedCard.CardType
                End If
            End If
        End If
        tmrCrdRefresh.enabled = True
        Call funSetFocus(fxgCards)
    Case 3
        If gblnDBOK Then
            If guServOpts.lngSelectedCardIndex = 0 Then
                Call funMsg(enuMsg.CardNotSelected)
            Else
                Set uSelectedCard = guCards(guServOpts.lngSelectedCardIndex)
                If funMsg(AskCanActivateCard, uSelectedCard.ID) = vbYes Then
                    grstCards.Index = "PrimaryKey"
                    Call grstCards.Seek("=", uSelectedCard.ID)
                    If grstCards.NoMatch Then
                        Exit Sub
                    End If
                    Call grstCards.Edit
                    uSelectedCard.Active = True
                    grstCards!Active = uSelectedCard.Active
                    Call grstCards.Update
                    Call subUpdateUISelectedCard
                End If
            End If
        End If
        Call funSetFocus(fxgCards)
    Case 4
        If gblnDBOK Then
            If guServOpts.lngSelectedCardIndex = 0 Then
                Call funMsg(enuMsg.CardNotSelected)
            Else
                Set uSelectedCard = guCards(guServOpts.lngSelectedCardIndex)
                If funMsg(AskCanDeactivateCard, uSelectedCard.ID) = vbYes Then
                    grstCards.Index = "PrimaryKey"
                    Call grstCards.Seek("=", uSelectedCard.ID)
                    If grstCards.NoMatch Then
                        Exit Sub
                    End If
                    Call grstCards.Edit
                    uSelectedCard.Active = False
                    grstCards!Active = uSelectedCard.Active
                    Call grstCards.Update
                    Call subUpdateUISelectedCard
                    intWorkstationIndex = funCardInUse(uSelectedCard)
                    If intWorkstationIndex <> 0 Then
                        If funMsg(enuMsg.AskCanCloseSession) = vbYes Then
                            strBuffer = CNX_SESSION_CLOSE & CNX_MSG_END
                            Call funSendData(guWorkstations(intWorkstationIndex).Cnx, strBuffer)
                        End If
                    End If
                End If
            End If
        End If
        Call funSetFocus(fxgCards)
    Case 5
        Call funSetFocus(fxgCards)
        cmdCrdButtons(5).enabled = False
        mnuCrd(5).enabled = False
        If guServOpts.lngSelectedCardIndex = 0 Then
            Call funMsg(enuMsg.CardNotSelected)
        Else
            Call subFillCards
            Set guSelectedCard = guCards(guServOpts.lngSelectedCardIndex)
            Call Load(frmCardHistory)
            Call frmCardHistory.subSetHistoryType(enuCustomerType.Card)
            Call frmCardHistory.Show(vbModal, frmMain)
        End If
        cmdCrdButtons(5).enabled = True
        mnuCrd(5).enabled = True
    Case 6
        If guServOpts.lngSelectedCardIndex = 0 Then
            Call funMsg(enuMsg.CardNotSelected)
        Else
            Set uCard = guCards(guServOpts.lngSelectedCardIndex)
            If funCardInUse(uCard) <> 0 Then
                Call funMsg(enuMsg.NoRechargeCardInUse)
            Else
                If uCard.Active And (uCard.TimeCredit > NULL_DATE Or uCard.UnlimitedPlay) Then
                    For i = 1 To guMacAddresses.Count
                        Set uWorkstation = funFindWorkstationByMac(guMacAddresses(i))
                        If uWorkstation.State = cliReady Then
                            Call subWorkstationCrdSessionOpen(uWorkstation, uCard)
                            Exit For
                        End If
                    Next i
                Else
                    Call funMsg(NoCredit)
                End If
            End If
        End If
        DoEvents
        Call subUpdateUISelectedCard
        Call funSetFocus(fxgCards)
    End Select
    Exit Sub
    
cmdCrdButtons_ClickError:
    Select Case Index
    Case 0
        cmdCrdButtons(0).enabled = True
        mnuCrd(0).enabled = True
        strErrorcontext = lngNewCardID
    Case 1
        cmdCrdButtons(1).enabled = True
        mnuCrd(1).enabled = True
    Case 5
        cmdCrdButtons(5).enabled = True
        mnuCrd(5).enabled = True

    End Select
    Call funErrorManager("frmMain.cmdCrdButtons_ClickError - Index: " & Index & IIf(strErrorcontext <> "", " Context: " & strErrorcontext, ""), True)
End Sub


'*******************************************************************************
Private Sub cmdCrdButtons_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo cmdCrdButtons_MouseDownError

    Select Case Index
    Case 7
        cmdCrdButtons(Index).value = True
        Call funSetFocus(tbsFunctions)
        Call PopupMenu(mnuCrdStartSessionDropdown, vbPopupMenuRightAlign, cmdCrdButtons(7).left + cmdCrdButtons(7).width + fraCrdControls.left + fraFunctions(2).left, cmdCrdButtons(7).top + cmdCrdButtons(7).height + fraCrdControls.top + fraFunctions(1).top, mnuCrdStartSessionFirstCilent)
    End Select
    Exit Sub
    
cmdCrdButtons_MouseDownError:
    Call funErrorManager("frmMain.cmdCrdButtons_MouseDownError - Index: " & Index, True)
End Sub


'*******************************************************************************
Private Sub cmdCrdDelete_Click()
On Error GoTo cmdCrdDelete_ClickError

    If guServOpts.lngSelectedCardIndex = 0 Then
        Call funMsg(enuMsg.CardNotSelected)
    Else
        Call subCardDelete(guCards(guServOpts.lngSelectedCardIndex))
        Call frmMain.fxgCards_RowColChange
    End If
    Call funSetFocus(fxgCards)
    Exit Sub
    
cmdCrdDelete_ClickError:
    Call funErrorManager("frmMain.cmdCrdDelete_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdCrdFind_Click()
On Error GoTo cmdCrdFind_ClickError
    
    Call funSetFocus(cmdCrdFind)
    Call subFindCard(txtCrdFindNum.Text)
    Call funSetFocus(txtCrdFindNum)
    Exit Sub
    
cmdCrdFind_ClickError:
    Call funErrorManager("frmMain.cmdCrdFind_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdCrdPendingEnable_Click()
Dim uNewCard As New clsCard
Dim dteNewCardDate As Date
Dim lngCardNum As Long
Dim dteExpiryDateNew As Long
Dim strRechargeTime As String
On Error GoTo cmdCrdPendingEnable_ClickError

    If fxgCrdPending.TextMatrix(guServOpts.lngSelectedCardPendingIndex, 1) = funLoadResString(11000) Then
        strRechargeTime = funLoadResString(11000)
    Else
        strRechargeTime = funFormatTime(fxgCrdPending.TextMatrix(guServOpts.lngSelectedCardPendingIndex, 1), enuTimeFormat.Hh_Nmn_Ssec_dyn)
    End If
    If funMsg(AskCanEnableCard, _
              funLoadResString(10949, _
                               fxgCrdPending.TextMatrix(guServOpts.lngSelectedCardPendingIndex, 0), _
                               strRechargeTime, _
                               fxgCrdPending.TextMatrix(guServOpts.lngSelectedCardPendingIndex, 2))) = vbOK Then
        
        grstCardsPending.Index = "PrimaryKey"
        Call grstCardsPending.Seek("=", fxgCrdPending.TextMatrix(guServOpts.lngSelectedCardPendingIndex, 0))
        If Not grstCardsPending.NoMatch Then
            txtCrdPending.Text = ""
            fxgCrdPending.Row = guServOpts.lngSelectedCardPendingIndex
            fxgCrdPending.ColSel = fxgCrdPending.Cols - 1
            fxgCrdPending.CellForeColor = vbWindowText
            fxgCrdPending.CellBackColor = vbWindowBackground
            guServOpts.lngSelectedCardPendingIndex = 0
            
            Call grstCards.AddNew
            dteNewCardDate = Date
            lngCardNum = grstCardsPending!ID
            grstCards!ID = lngCardNum
            grstCards!Active = True
            grstCards!date_creation = dteNewCardDate
            grstCards!Comment = grstCardsPending!Comment
            grstCards!Type = grstCardsPending!Type
            Call grstCards.Update
            Call subDBAddEvent(enuEvents.CardNew, 0, grstCardsPending!ID, "", 0, funLoadResString(10341), All)
            If grstCardsPending!unlimited_play Or grstCardsPending!time_credit <> NULL_DATE Then
                uNewCard.ID = lngCardNum
                uNewCard.DateCreation = dteNewCardDate
                uNewCard.TimeCredit = grstCardsPending!time_credit
                uNewCard.UnlimitedPlay = grstCardsPending!unlimited_play
                Select Case grstCardsPending!validity_type
                Case enuRechargeValidityType.Hours
                    dteExpiryDateNew = DateAdd("h", grstCardsPending!validity, Date)
                Case enuRechargeValidityType.Days
                    dteExpiryDateNew = DateAdd("d", grstCardsPending!validity, Date)
                Case enuRechargeValidityType.Weeks
                    dteExpiryDateNew = DateAdd("ww", grstCardsPending!validity, Date)
                Case enuRechargeValidityType.Months
                    dteExpiryDateNew = DateAdd("d", grstCardsPending!validity * guServOpts.uConfig.intNumDaysInMonth, Date)
                Case enuRechargeValidityType.Years
                    dteExpiryDateNew = DateAdd("d", grstCardsPending!validity * NUM_DAYS_IN_YEAR, Date)
                End Select
                uNewCard.DateExpiry = dteExpiryDateNew
                Call subCardRecharge(uNewCard, grstCardsPending!Price, grstCardsPending!time_credit, funLoadResString(10289) & " " & funFormatTime(grstCardsPending!time_credit) & " " & grstCardsPending!Caption, funCardInUse(uNewCard))
            End If
            Call grstCardsPending.Delete
            Call subFillCardsPending
            Call subFillCards
            Call subResizeCards
            Call subFindCard(CStr(lngCardNum), False)
            Call funSetFocus(txtCrdPending)
        End If
    End If
    Exit Sub
    
cmdCrdPendingEnable_ClickError:
    Call funErrorManager("frmMain.cmdCrdPendingEnable_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdIncomeExport_Click()
On Error GoTo cmdIncomeExport_ClickError
        
    Call subExportIncome
    Exit Sub
    
cmdIncomeExport_ClickError:
    Call funErrorManager("frmMain.cmdIncomeExport_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdIncomeFilter_Click()
Dim i As Integer
Dim intSelectedFilter As Integer
Dim dteIncomeStart As Date
Dim dteIncomeEnd As Date
Dim chkTmp As CheckBox
Dim strTypes As String
On Error GoTo cmdIncomeFilter_ClickError
    
    Screen.MousePointer = vbArrowHourglass
    fxgIncome.MousePointer = vbArrowHourglass
    cmdIncomeExport.enabled = False
    cmdIncomePrint.enabled = False
    cmdIncomeFilter.enabled = False
    mnuIncomeApply.enabled = False
    For i = optIncomeFilter.LBound To optIncomeFilter.UBound
        If Not optIncomeFilter(i).value Then optIncomeFilter(i).enabled = False
    Next i
    For i = optIncomeFilter.LBound To optIncomeFilter.UBound
        If optIncomeFilter(i).value Then optIncomeFilter(i).enabled = False
    Next i
    mnuIncomeExport.enabled = False
    mnuIncomePrint.enabled = False
    For i = 0 To optIncomeFilter.UBound
        If optIncomeFilter(i).value Then
            intSelectedFilter = i
            Exit For
        End If
    Next i
    Select Case intSelectedFilter
    Case 0
        dteIncomeStart = DateAdd("d", -6, Now)
        dteIncomeEnd = Now
    
    Case 1
        dteIncomeEnd = Now
        Do While Weekday(dteIncomeEnd) <> vbSunday
            dteIncomeEnd = DateAdd("d", -1, dteIncomeEnd)
        Loop
        dteIncomeStart = dteIncomeEnd
        Do While Weekday(dteIncomeStart) <> vbMonday
            dteIncomeStart = DateAdd("d", -1, dteIncomeStart)
        Loop
    
    Case 2
        dteIncomeStart = DateAdd("d", -29, Now)
        dteIncomeEnd = Now
    
    Case 3
        dteIncomeEnd = Now
        Do While DatePart("m", dteIncomeEnd) = DatePart("m", Now)
            dteIncomeEnd = DateAdd("d", -1, dteIncomeEnd)
        Loop
        dteIncomeStart = DateAdd("d", 1 - DatePart("d", dteIncomeEnd), dteIncomeEnd)
    
    Case 4
        dteIncomeStart = DateAdd("w", -364, Now)
        dteIncomeEnd = Now
    
    Case 5
        dteIncomeEnd = Now
        dteIncomeStart = DateAdd("d", 1 - DatePart("y", Now), dteIncomeEnd)
    
    Case 6
        dteIncomeStart = dtpStartDate.value
        dteIncomeEnd = dtpEndDate.value
    
    Case 7
        dteIncomeEnd = Now
        dteIncomeStart = dteIncomeEnd
    
    Case 8
        dteIncomeEnd = DateAdd("d", -1, Now)
        dteIncomeStart = dteIncomeEnd
    
    End Select
    For Each chkTmp In chkTypeIncome
        If chkTmp.value = vbChecked Then strTypes = strTypes & IIf(strTypes = "", "", ",") & chkTmp.Index
    Next
        
    Call subFillIncome(dteIncomeStart, dteIncomeEnd, strTypes)
    Screen.MousePointer = vbNormal
    fxgIncome.MousePointer = vbNormal
    cmdIncomeExport.enabled = True
    cmdIncomePrint.enabled = guServOpts.uConfig.strDefaultPrinterName <> ""
    cmdIncomeFilter.enabled = True
    For i = optIncomeFilter.LBound To optIncomeFilter.UBound
        optIncomeFilter(i).enabled = True
    Next i
    mnuIncomeApply.enabled = True
    mnuIncomeExport.enabled = True
    mnuIncomePrint.enabled = True
    Exit Sub
    
cmdIncomeFilter_ClickError:
    Screen.MousePointer = vbNormal
    fxgIncome.MousePointer = vbNormal
    cmdIncomeExport.enabled = True
    cmdIncomePrint.enabled = guServOpts.uConfig.strDefaultPrinterName <> ""
    cmdIncomeFilter.enabled = True
    For i = optIncomeFilter.LBound To optIncomeFilter.UBound
        optIncomeFilter(i).enabled = True
    Next i
    mnuIncomeApply.enabled = True
    mnuIncomeExport.enabled = True
    mnuIncomePrint.enabled = True
    Call funErrorManager("frmMain.cmdIncomeFilter_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdIncomePrint_Click()
On Error GoTo cmdIncomePrint_ClickError

    Call subStatsPrint(fxgIncome, tbsFunctions.SelectedItem.Caption)
    Exit Sub
    
cmdIncomePrint_ClickError:
    Call funErrorManager("frmMain.cmdIncomePrint_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdServicesCollect_Click()
Dim sngPrice As Single
Dim lngQuantity As Long
Dim strLabel As String
Dim uCurrentService As New clsService
Dim uMsg As New clsStatusBarMessage
On Error GoTo cmdServicesCollect_ClickError
    
    Set uCurrentService = guServices(guServOpts.lngSelectedServiceIndex)
    lngQuantity = CLng(txtServicesControlsQuantity.Text)
    sngPrice = lngQuantity * uCurrentService.Price
    strLabel = uCurrentService.Label & " x " & lngQuantity
    Call grstJournIncome.AddNew
    grstJournIncome!Type = enuIncomeType.Other
    grstJournIncome!income = sngPrice
    grstJournIncome!date_stamp = Now
    grstJournIncome!Label = left(strLabel, 50)
    grstJournIncome!paid = True
    grstJournIncome!login = guServOpts.uConnectedUser.login
    Call subDBAddEvent(enuEvents.ServicePay, 0, 0, "", grstJournIncome!ID, funLoadResString(10343) & strLabel, All)
    Call grstJournIncome.Update
    
    If uCurrentService.StockMode = enuServiceStockMode.Managed Then
        uCurrentService.Stock = uCurrentService.Stock - lngQuantity
        grstServices.Index = "PrimaryKey"
        Call grstServices.Seek("=", uCurrentService.ID)
        If Not grstServices.NoMatch Then
            Call grstServices.Edit
            grstServices!Stock = uCurrentService.Stock
            Call grstServices.Update
        End If
        If uCurrentService.StockMode = enuServiceStockMode.Managed And uCurrentService.Stock <= uCurrentService.StockMinAlert Then
            uMsg.ImageIndex = enuStatus.cliInfo
            uMsg.Text = funLoadResString(10373) & uCurrentService.Label & funLoadResString(10374)
            Call funAddStatusBarMsg(uMsg, "SERV_" & uCurrentService.ID)
        End If
        fxgServices.TextMatrix(guServOpts.lngSelectedServiceIndex, 2) = uCurrentService.Stock
        lblServSel(2).Caption = " " & uCurrentService.Stock
        lblServSel(2).ForeColor = IIf(uCurrentService.Stock <= uCurrentService.StockMinAlert, vbRed, vbWindowText)
    End If
    udServicesControlsQuantity.value = 0
    Exit Sub

cmdServicesCollect_ClickError:
    Call funErrorManager("frmMain.cmdServicesCollect_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdSessionExport_Click()
On Error GoTo cmdSessionExport_ClickError

    Screen.MousePointer = vbArrowHourglass
    fxgSession(2).MousePointer = vbArrowHourglass
    fxgSession(1).MousePointer = vbArrowHourglass
    cmdSessionExport.enabled = False
    cmdSessionPrint.enabled = False
    cmdSessionFilter.enabled = False
    mnuSessionApply.enabled = False
    mnuSessionExport.enabled = False
    mnuSessionPrint.enabled = False
    
    Select Case tbsSessionResultType.SelectedItem.Index
    
    Case 1
        Call subExportSessionStats
        
    Case 2
        Call subExportSessionTimes
        
    Case 3
        Call subExportSessionList

    Case 4
        Call subExportSessionURLs

    End Select
    cmdSessionExport.enabled = Not tbsSessionResultType.SelectedItem.Index = 5
    cmdSessionPrint.enabled = Not tbsSessionResultType.SelectedItem.Index = 5 And guServOpts.uConfig.strDefaultPrinterName <> ""
    cmdSessionFilter.enabled = True
    mnuSessionApply.enabled = True
    mnuSessionExport.enabled = Not tbsSessionResultType.SelectedItem.Index = 5
    mnuSessionPrint.enabled = Not tbsSessionResultType.SelectedItem.Index = 5 And guServOpts.uConfig.strDefaultPrinterName <> ""
    Screen.MousePointer = vbNormal
    fxgSession(2).MousePointer = vbNormal
    fxgSession(1).MousePointer = vbNormal
    Exit Sub

cmdSessionExport_ClickError:
    cmdSessionExport.enabled = Not tbsSessionResultType.SelectedItem.Index = 5
    cmdSessionPrint.enabled = Not tbsSessionResultType.SelectedItem.Index = 5 And guServOpts.uConfig.strDefaultPrinterName <> ""
    cmdSessionFilter.enabled = True
    mnuSessionApply.enabled = True
    mnuSessionExport.enabled = Not tbsSessionResultType.SelectedItem.Index = 5
    mnuSessionPrint.enabled = Not tbsSessionResultType.SelectedItem.Index = 5 And guServOpts.uConfig.strDefaultPrinterName <> ""
    Screen.MousePointer = vbNormal
    fxgSession(2).MousePointer = vbNormal
    fxgSession(1).MousePointer = vbNormal
    Call funErrorManager("frmMain.cmdSessionExport_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdSessionFilter_Click()
Dim i As Integer
Dim intSelectedFilter As Integer
Dim dteSessionStart As Date
Dim dteSessionEnd As Date
On Error GoTo cmdSessionFilter_ClickError

    Call Load(frmStatsProgress)
    Screen.MousePointer = vbArrowHourglass
    fxgSession(2).MousePointer = vbArrowHourglass
    fxgSession(1).MousePointer = vbArrowHourglass
    fxgSession(4).MousePointer = vbArrowHourglass
    fxgSession(3).MousePointer = vbArrowHourglass
    cmdSessionExport.enabled = False
    cmdSessionFilter.enabled = False
    cmdSessionPrint.enabled = False
    For i = optSessionFilter.LBound To optSessionFilter.UBound
        If Not optSessionFilter(i).value Then optSessionFilter(i).enabled = False
    Next i
    For i = optSessionFilter.LBound To optSessionFilter.UBound
        If optSessionFilter(i).value Then optSessionFilter(i).enabled = False
    Next i
    For i = chkSessStatType.LBound To chkSessStatType.UBound
        chkSessStatType(i).enabled = False
    Next i
    mnuSessionApply.enabled = False
    mnuSessionExport.enabled = False
    mnuSessionPrint.enabled = False
    mblnSessionFiltered = True
    gblnStopStats = False
    
    For i = 0 To optSessionFilter.UBound
        If optSessionFilter(i).value Then
            intSelectedFilter = i
            Exit For
        End If
    Next i
    Select Case intSelectedFilter
    Case 0
        dteSessionStart = DateAdd("d", -6, Now)
        dteSessionEnd = Now
    
    Case 1
        dteSessionEnd = Now
        Do While Weekday(dteSessionEnd) <> vbSunday
            dteSessionEnd = DateAdd("d", -1, dteSessionEnd)
        Loop
        dteSessionStart = dteSessionEnd
        Do While Weekday(dteSessionStart) <> vbMonday
            dteSessionStart = DateAdd("d", -1, dteSessionStart)
        Loop
    
    Case 2
        dteSessionStart = DateAdd("d", -29, Now)
        dteSessionEnd = Now
    
    Case 3
        dteSessionEnd = Now
        Do While DatePart("m", dteSessionEnd) = DatePart("m", Now)
            dteSessionEnd = DateAdd("d", -1, dteSessionEnd)
        Loop
        dteSessionStart = DateAdd("d", 1 - DatePart("d", dteSessionEnd), dteSessionEnd)
    
    Case 4
        dteSessionStart = DateAdd("w", -364, Now)
        dteSessionEnd = Now
    
    Case 5
        dteSessionEnd = Now
        dteSessionStart = DateAdd("d", 1 - DatePart("y", Now), dteSessionEnd)
    
    Case 6
        dteSessionStart = dtpSessionStartDate.value
        dteSessionEnd = dtpSessionEndDate.value
    
    Case 7
        dteSessionEnd = Now
        dteSessionStart = dteSessionEnd
    
    Case 8
        dteSessionEnd = DateAdd("d", -1, Now)
        dteSessionStart = dteSessionEnd
    
    End Select
    dteSessionStart = CDate(DateValue(dteSessionStart) & " 00:00:00")
    dteSessionEnd = CDate(DateValue(dteSessionEnd) & " 23:59:59")
    
    Call subInitSessions(False)

    If chkSessStatType(1).value = vbChecked Then
        frmStatsProgress.lblType.Caption = funLoadResString(10430)
        If Not gblnStopStats Then Call subProcessSessionStats(dteSessionStart, dteSessionEnd)
        If Not gblnStopStats Then Call subFillSessionStats
    End If
    frmStatsProgress.pgbStatType.value = 1
    frmStatsProgress.pgbStatProgress.value = 0
    
    If chkSessStatType(2).value = vbChecked Then
        frmStatsProgress.lblType.Caption = funLoadResString(10431)
        If Not gblnStopStats Then Call subProcessSessionTimes(dteSessionStart, dteSessionEnd)
        If Not gblnStopStats Then Call subFillSessionTimes
    End If
    frmStatsProgress.pgbStatType.value = 2
    frmStatsProgress.pgbStatProgress.value = 0
    
    If chkSessStatType(3).value = vbChecked Then
        frmStatsProgress.lblType.Caption = funLoadResString(10432)
        If Not gblnStopStats Then Call subQuerySession(dteSessionStart, dteSessionEnd)
        If Not gblnStopStats Then Call subFillSessionList
    End If
    frmStatsProgress.pgbStatType.value = 3
    frmStatsProgress.pgbStatProgress.value = 0
    
    If chkSessStatType(4).value = vbChecked Then
        frmStatsProgress.lblType.Caption = funLoadResString(10700)
        If Not gblnStopStats Then Call subQueryURLs(dteSessionStart, dteSessionEnd)
        If Not gblnStopStats Then Call subFillURLList
    End If
    frmStatsProgress.pgbStatType.value = 4
    frmStatsProgress.pgbStatProgress.value = 0
    
    If chkSessStatType(5).value = vbChecked Then
        frmStatsProgress.lblType.Caption = funLoadResString(10433)
        If Not gblnStopStats And chkSessStatType(3).value <> vbChecked Then Call subQuerySession(dteSessionStart, dteSessionEnd)
        If Not gblnStopStats Then Call subFillSessionLines
    End If
    frmStatsProgress.pgbStatType.value = 5
    
    cmdSessionExport.enabled = Not tbsSessionResultType.SelectedItem.Index = 4
    cmdSessionPrint.enabled = Not tbsSessionResultType.SelectedItem.Index = 4 And guServOpts.uConfig.strDefaultPrinterName <> ""
    cmdSessionFilter.enabled = True
    mnuSessionApply.enabled = True
    For i = optSessionFilter.LBound To optSessionFilter.UBound
        optSessionFilter(i).enabled = True
    Next i
    For i = chkSessStatType.LBound To chkSessStatType.UBound
        chkSessStatType(i).enabled = True
    Next i
    mnuSessionExport.enabled = Not tbsSessionResultType.SelectedItem.Index = 4
    mnuSessionPrint.enabled = Not tbsSessionResultType.SelectedItem.Index = 4 And guServOpts.uConfig.strDefaultPrinterName <> ""
    Screen.MousePointer = vbNormal
    fxgSession(2).MousePointer = vbNormal
    fxgSession(1).MousePointer = vbNormal
    fxgSession(4).MousePointer = vbNormal
    fxgSession(3).MousePointer = vbNormal
    
    Call Unload(frmStatsProgress)
    Exit Sub

cmdSessionFilter_ClickError:
    cmdSessionExport.enabled = Not tbsSessionResultType.SelectedItem.Index = 4
    cmdSessionPrint.enabled = Not tbsSessionResultType.SelectedItem.Index = 4 And guServOpts.uConfig.strDefaultPrinterName <> ""
    cmdSessionFilter.enabled = True
    mnuSessionApply.enabled = True
    For i = optSessionFilter.LBound To optSessionFilter.UBound
        optSessionFilter(i).enabled = True
    Next i
    For i = chkSessStatType.LBound To chkSessStatType.UBound
        chkSessStatType(i).enabled = True
    Next i
    mnuSessionExport.enabled = Not tbsSessionResultType.SelectedItem.Index = 4
    mnuSessionPrint.enabled = Not tbsSessionResultType.SelectedItem.Index = 4
    Screen.MousePointer = vbNormal
    fxgSession(2).MousePointer = vbNormal
    fxgSession(1).MousePointer = vbNormal
    fxgSession(4).MousePointer = vbNormal
    fxgSession(3).MousePointer = vbNormal
    Call Unload(frmStatsProgress)
    Call funErrorManager("frmMain.cmdSessionFilter_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdSessionPrint_Click()
On Error GoTo cmdSessionPrint_ClickError

    Call subStatsPrint(fxgSession(tbsSessionResultType.SelectedItem.Index), tbsFunctions.SelectedItem.Caption & " - " & tbsSessionResultType.SelectedItem.Caption)
    Exit Sub
    
cmdSessionPrint_ClickError:
    Call funErrorManager("frmMain.cmdSessionPrint_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdSessionURLSearch_Click()
Dim i As Long
Dim strSearchString As String
Dim blnFound As Boolean
On Error GoTo cmdSessionURLSearch_ClickError
    
    strSearchString = txtSessionURLSearch.Text
    If Len(strSearchString) = 0 Then Exit Sub
    Call funSetFocus(cmdSessionURLSearch)
    With fxgSession(4)
        For i = IIf(.Row <> .Rows, .Row + 1, 1) To .Rows - 1
            blnFound = InStr(LCase(.TextMatrix(i, 3) & " " & .TextMatrix(i, 4)), LCase(strSearchString)) <> 0
            If blnFound Then Exit For
        Next i
        If Not blnFound And .Row <> 0 Then
            For i = 1 To .Row
                blnFound = InStr(LCase(.TextMatrix(i, 3) & " " & .TextMatrix(i, 4)), LCase(strSearchString)) <> 0
                If blnFound Then Exit For
            Next i
        End If
        If blnFound Then
            .Row = i
            Call fxgSession_RowColChange(4)
        Else
            Call funMsg(enuMsg.URLNotFound)
        End If
    End With
    Call funSetFocus(txtSessionURLSearch)
    Exit Sub
    
cmdSessionURLSearch_ClickError:
    Call funErrorManager("frmMain.cmdSessionURLSearch_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdWksnButtons_Click(Index As Integer)
Dim strBuffer As String
Dim uSelectedCard As clsCard
Dim uSelectedClient As clsClient
On Error GoTo cmdWksnButtons_ClickError

    Select Case Index
    Case 0
        strBuffer = CNX_ACTIVATE_WORKSTATION & CNX_MSG_END
        Call funSendData(guSelectedWorkstation.Cnx, strBuffer)
        Call funSetFocus(ucWorkstations(1))
    Case 1 ' Pay PostpaidSs
        Call subSessionPay(guSelectedWorkstation)
        strBuffer = IIf(guServOpts.uWorkstations.intDefaultState = enuWksnState.Ready, CNX_ACTIVATE_WORKSTATION, CNX_LOCK_WORKSTATION) & CNX_MSG_END & CNX_EMPTY_SERIALIZED_USER & CNX_MSG_END
        Call funSendData(guSelectedWorkstation.Cnx, strBuffer)
        Call funSetFocus(ucWorkstations(1))
    Case 2
        If funMsg(enuMsg.AskCanLockWorkstation) = vbYes Then
            strBuffer = CNX_LOCK_WORKSTATION & CNX_MSG_END
            Call funSendData(guSelectedWorkstation.Cnx, strBuffer)
        End If
        Call funSetFocus(ucWorkstations(1))
    Case 3
        strBuffer = CNX_SESSION_FREE & CNX_MSG_END
        Call funSendData(guSelectedWorkstation.Cnx, strBuffer)
        guSelectedWorkstation.StartTime = Now
        guSelectedWorkstation.State = enuWksnState.UsedFree
        Call subSessionOpen(guSelectedWorkstation)
    Case 4
        If funMsg(enuMsg.AskCanCloseSession) = vbYes Then
            strBuffer = CNX_SESSION_CLOSE & CNX_MSG_END
            Call funSendData(guSelectedWorkstation.Cnx, strBuffer)
        End If
        Call funSetFocus(ucWorkstations(1))
    Case 5
        strBuffer = CNX_PAUSE_WORKSTATION & CNX_MSG_END
        Call funSendData(guSelectedWorkstation.Cnx, strBuffer)
        Call funSetFocus(ucWorkstations(1))
    Case 7
        strBuffer = CNX_RESUME_WORKSTATION & CNX_MSG_END
        Call funSendData(guSelectedWorkstation.Cnx, strBuffer)
        cmdWksnButtons(7).enabled = False
        Call funSetFocus(ucWorkstations(1))
    Case 8
        strBuffer = CNX_UNPAUSE_WORKSTATION & CNX_MSG_END
        Call funSendData(guSelectedWorkstation.Cnx, strBuffer)
        Call funSetFocus(ucWorkstations(1))
    Case 9 ' recharge...
        Call funSetFocus(ucWorkstations(1))
        cmdWksnButtons(9).enabled = False
        mnuWksnCrdRecharge.enabled = False
        Call Load(frmRecharge)
        Select Case guSelectedWorkstation.SessionType
        Case enuSessionType.Card
            Set guSelectedCard = guSelectedWorkstation.CurrentCard
            Call frmRecharge.subSetRechargeType(enuCustomerType.Card)
        Case enuSessionType.Client
            Set guSelectedClient = guSelectedWorkstation.CurrentClient
            Call frmRecharge.subSetRechargeType(enuCustomerType.Client)
        Case enuSessionType.SelfServicePrePaid
            Set guSelectedClient = guSelectedWorkstation.SelfServicePrePaid
            Call frmRecharge.subSetRechargeType(enuCustomerType.SelfServicePrePaid)
        End Select
        Call frmRecharge.Show(vbModal, frmMain)
        cmdWksnButtons(9).enabled = True
        mnuWksnCrdRecharge.enabled = True
    Case 10 ' deactivate account
        If gblnDBOK Then
            Select Case guSelectedWorkstation.SessionType
            Case enuSessionType.Card
                Set uSelectedCard = guSelectedWorkstation.CurrentCard
                grstCards.Index = "PrimaryKey"
                Call grstCards.Seek("=", uSelectedCard.ID)
                If grstCards.NoMatch Then
                    Exit Sub
                End If
                Call grstCards.Edit
                uSelectedCard.Active = False
                grstCards!Active = uSelectedCard.Active
                Call grstCards.Update
                Call subFillCards
            Case enuSessionType.Client
                Set uSelectedClient = guSelectedWorkstation.CurrentClient
                grstClients.Index = "PrimaryKey"
                Call grstClients.Seek("=", uSelectedClient.login)
                If grstClients.NoMatch Then
                    Exit Sub
                End If
                Call grstClients.Edit
                uSelectedClient.Active = False
                grstClients!Active = uSelectedClient.Active
                Call grstClients.Update
                Call subFillClients
            End Select
            If funMsg(enuMsg.AskCanCloseSession) = vbYes Then
                strBuffer = CNX_SESSION_CLOSE & CNX_MSG_END
                Call funSendData(guSelectedWorkstation.Cnx, strBuffer)
            End If
        End If
        Call funSetFocus(ucWorkstations(1))
    Case 11
        Call funSetFocus(ucWorkstations(1))
        cmdWksnButtons(11).enabled = False
        mnuWksnCrdHistory.enabled = False
        Set guSelectedCard = guSelectedWorkstation.CurrentCard
        Call Load(frmCardHistory)
        Call frmCardHistory.subSetHistoryType(enuCustomerType.Card)
        Call frmCardHistory.Show(vbModal, frmMain)
        cmdWksnButtons(11).enabled = True
        mnuWksnCrdHistory.enabled = True
    Case 12
        Call funSetFocus(ucWorkstations(1))
        cmdWksnButtons(12).enabled = False
        mnuWksnCrdHistory.enabled = False
        Set guSelectedClient = guSelectedWorkstation.CurrentClient
        Call Load(frmCardHistory)
        Call frmCardHistory.subSetHistoryType(enuCustomerType.Client)
        Call frmCardHistory.Show(vbModal, frmMain)
        cmdWksnButtons(12).enabled = True
        mnuWksnCrdHistory.enabled = True
    Case 13
        Call subWorkstationChat(guSelectedWorkstation, txtWksnChat)
        txtWksnChat.Text = ""
        Call funSetFocus(txtWksnChat)
    
    Case 14 'New prepaid self service session
        cmdWksnButtons(14).enabled = False
        mnuWksnStartSelfServPrePaid.enabled = False
        Call funSendData(guSelectedWorkstation.Cnx, CNX_LOCK_WORKSTATION & CNX_MSG_END)
        DoEvents
        Set guSelectedWorkstation.SelfServicePrePaid = New clsClient
        guSelectedWorkstation.SelfServicePrePaid.DateCreation = Now
        guSelectedWorkstation.SelfServicePrePaid.Active = True
        Set guSelectedClient = guSelectedWorkstation.SelfServicePrePaid
        Call frmRecharge.subSetRechargeType(enuCustomerType.SelfServicePrePaid)
        Call frmRecharge.Show(vbModal, frmMain)
        If guSelectedWorkstation.SelfServicePrePaid.TimeCredit <> NULL_DATE Then
            Call subWorkstationSelfServicePrePaidSessionOpen(guSelectedWorkstation)
        Else
            Call funSendData(guSelectedWorkstation.Cnx, CNX_ACTIVATE_WORKSTATION & CNX_MSG_END)
        End If
        cmdWksnButtons(14).enabled = True
        mnuWksnStartSelfServPrePaid.enabled = True
    
    Case 15 ' Cancel postpaidSs payment (time < grace period)
        Call subSessionCancelPay(guSelectedWorkstation)
        strBuffer = IIf(guServOpts.uWorkstations.intDefaultState = enuWksnState.Ready, CNX_ACTIVATE_WORKSTATION, CNX_LOCK_WORKSTATION) & CNX_MSG_END & CNX_EMPTY_SERIALIZED_USER & CNX_MSG_END
        Call funSendData(guSelectedWorkstation.Cnx, strBuffer)
        Call funSetFocus(ucWorkstations(1))
    
    End Select
    Exit Sub

cmdWksnButtons_ClickError:
    Select Case Index
    Case 9
        cmdWksnButtons(9).enabled = True
        mnuWksnCrdRecharge.enabled = True
    Case 11
        cmdWksnButtons(11).enabled = True
        mnuWksnCrdHistory.enabled = True
    Case 12
        cmdWksnButtons(12).enabled = True
        mnuWksnCrdHistory.enabled = True
    Case 14
        cmdWksnButtons(14).enabled = True
        mnuWksnStartSelfServPrePaid.enabled = True
    
    End Select
    Call funErrorManager("frmMain.cmdWksnButtons_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdWksnButtons_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo cmdWksnButtons_MouseDownError

    Select Case Index
    Case 6
        Call funSetFocus(ucWorkstations(1))
        Call PopupMenu(mnuWksnPauseDropDown, vbPopupMenuRightAlign, cmdWksnButtons(6).left + cmdWksnButtons(6).width + fraWksnButtons.left + fraWksnControls.left + fraFunctions(1).left, cmdWksnButtons(6).top + cmdWksnButtons(6).height + fraWksnButtons.top + fraFunctions(1).top, mnuWksnPauseOne)
    End Select
    Exit Sub

cmdWksnButtons_MouseDownError:
    Call funErrorManager("frmMain.cmdWksnButtons_MouseDownError", True)
End Sub


'*******************************************************************************
Private Sub Form_GotFocus()
On Error Resume Next

    gblnfrmMain = True
End Sub


'*******************************************************************************
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    Call subDebugPrint("frmMain.Form_KeyUp: KeyCode: " & KeyCode & " | Shift: " & Shift)
    Select Case tbsFunctions.SelectedItem.Key
    Case "k1"
    
    Case "k2"
        If Shift = 0 Then
            Select Case KeyCode
            Case vbKeyF3
                If txtCrdFindNum.Text = "" Then
                    Call funSetFocus(txtCrdFindNum)
                Else
                    cmdCrdFind.value = True
                End If
            Case vbKeyF4
                Call mnuCrd_Click(9)
            End Select
        
        ElseIf Shift = vbCtrlMask Then
            Select Case KeyCode
            Case vbKeyE
                Call funSetFocus(txtCrdPending)
            Case vbKeyF
                Call funSetFocus(txtCrdFindNum)
            Case vbKeyH
                Call mnuCrd_Click(5)
            Case vbKeyN
                Call mnuCrd_Click(0)
            Case vbKeyR
                Call mnuCrd_Click(1)
            Case vbKeyS
                Call mnuCrd_Click(2)
            End Select
        End If
    Case "k3"
        If Shift = 0 Then
            Select Case KeyCode
            Case vbKeyF3
                If txtCliFind.Text = "" Then
                    Call funSetFocus(txtCliFind)
                Else
                    cmdCliFind.value = True
                End If
            Case vbKeyF4
                Call mnuCli_Click(9)
            End Select
        
        ElseIf Shift = vbCtrlMask Then
            Select Case KeyCode
            Case vbKeyF
                Call funSetFocus(txtCliFind)
            Case vbKeyH
                Call mnuCli_Click(5)
            Case vbKeyM
                Call mnuCli_Click(2)
            Case vbKeyN
                Call mnuCli_Click(0)
            Case vbKeyR
                Call mnuCli_Click(1)
            End Select
        End If
    Case "k4"
    
    Case "k5"
        If Shift = 0 Then
        ElseIf Shift = vbCtrlMask Then
            Select Case KeyCode
            Case vbKeyP
                If cmdSessionPrint.enabled Then cmdSessionPrint.value = True
            End Select
        End If
    Case "k6"
        If Shift = 0 Then
        ElseIf Shift = vbCtrlMask Then
            Select Case KeyCode
            Case vbKeyP
                If cmdIncomePrint.enabled Then cmdIncomePrint.value = True
            End Select
        End If
    End Select
End Sub


'*******************************************************************************
Private Sub Form_Load()
On Error GoTo Form_LoadError

    gHW = Me.hWnd
    If Not gblnDebug Then Call subHook
    Call subSetDefaultFont(Me)
    
    Call subLocalizeInterface
    If guServOpts.uWinPos.intTop + frmMain.height > Screen.height Then guServOpts.uWinPos.intTop = Screen.height - frmMain.height
    If guServOpts.uWinPos.intLeft + frmMain.width > Screen.width Then guServOpts.uWinPos.intLeft = Screen.width - frmMain.width
    frmMain.top = guServOpts.uWinPos.intTop
    frmMain.left = guServOpts.uWinPos.intLeft
    frmMain.height = guServOpts.uWinPos.intHeight
    frmMain.width = guServOpts.uWinPos.intWidth
    frmMain.WindowState = guServOpts.uWinPos.intWinState
    
    With guIconData
        .cbSize = Len(guIconData)                       ' The length of the NOTIFYICONDATA type
        .hIcon = imgSystTrayIcons.ListImages(1).Picture                 ' A reference to the form's icon
        .hWnd = Me.hWnd                                 ' hWnd of the form
        .szTip = funLoadResString(10000) & Chr(0)                     ' Tooltip string delimited with a null character
        .uCallbackMessage = WM_MOUSEMOVE                ' The icon we're placing will send messages to the MouseMove event
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE    ' It will have message handling and a tooltip
        .uID = vbNull                                   ' uID is not used by VB, so it's set to a Null value
    End With
    Exit Sub

Form_LoadError:
    Call funErrorManager("frmMain.Form_LoadError", True)
End Sub


'*******************************************************************************
Private Sub Form_LostFocus()
On Error Resume Next

    gblnfrmMain = False
End Sub


'*******************************************************************************
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lngMsg As Long
On Error GoTo Form_MouseMoveError

    lngMsg = x
    If lngMsg = WM_LBUTTONDBLCLK Or lngMsg = 7725 Then
        If Me.Visible Then
            Call mnuSysTrayHide_Click
        Else
            Call mnuSysTrayShow_Click
        End If
    ElseIf lngMsg = WM_RBUTTONDOWN Or lngMsg = 7740 Then '7740, go figure...
        mnuSysTrayShow.Visible = True
        mnuSysTrayHide.Visible = True
                
        mnuSysTrayShow.Visible = Not Me.Visible
        mnuSysTrayHide.Visible = Me.Visible
        mnuSysTrayHide.enabled = Not (gblnfrmCardHistory Or gblnfrmCardMassAdd Or gblnfrmLogon Or gblnfrmRecharge Or gblnfrmSetupWizard Or gblnfrmOptions)
        Call PopupMenu(mnuSysTray, , , , IIf(Me.Visible, mnuSysTrayHide, mnuSysTrayShow))
    End If
    Exit Sub

Form_MouseMoveError:
    Call funErrorManager("frmMain.Form_MouseMoveError", True)
End Sub


'*******************************************************************************
Private Sub Form_Resize()
Dim i As Integer
On Error GoTo Form_ResizeError

    If frmMain.WindowState = vbNormal Then
        If gblnDebug And frmMain.height < MAX_WIN_HEIGHT * Screen.TwipsPerPixelY Then frmMain.height = MAX_WIN_HEIGHT * Screen.TwipsPerPixelY
        If gblnDebug And frmMain.width < MAX_WIN_WIDTH * Screen.TwipsPerPixelX Then frmMain.width = MAX_WIN_WIDTH * Screen.TwipsPerPixelX
    End If
    If frmMain.WindowState <> vbMinimized Then
        For i = 1 To fraFunctions.Count
            fraFunctions(i).width = frmMain.width - 240
            fraFunctions(i).height = frmMain.ScaleHeight - frmMain.tbsFunctions.height - frmMain.stbStatus.height - 60 * 2
        Next i
        tbsFunctions.width = frmMain.width - 240
        
        Call subResizeWorkstations
        Call subResizeCards
        Call subResizeClients
        Call subResizeIncome
        Call subResizeServices
        Call subResizeSession
    
    End If
    Exit Sub
    
Form_ResizeError:
    Call funErrorManager("frmMain.Form_ResizeError", True)
End Sub


'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
Dim blnShutdownApplication As Boolean
On Error GoTo Form_UnloadError
    
    blnShutdownApplication = True
    Cancel = DO_NOT_SHUTDOWN
    If guWorkstations.Count >= 1 Then
        If vbYes = funMsg(ConfirmServerShutdown, guWorkstations.Count) Then
            If guServOpts.uConnectedUser.Passwd = "" Then
                Call funShutdownApplication
            Else
                Call subQuitPassword(frmMain)
            End If
        End If
    Else
        Call funShutdownApplication
    End If
    Exit Sub

Form_UnloadError:
    Call funErrorManager("frmMain.Form_UnloadError", True)
End Sub


'*******************************************************************************
Public Sub fraWksnConsole_Click()
Dim i As Integer
Dim uEmptyWorkstation As New clsWorkstation
On Error GoTo fraWksnConsole_ClickError
    
    For i = 1 To ucWorkstations.Count
        ucWorkstations(i).Active = False
    Next i
    shpWksnConsoleBackground.BackColor = vb3DHighlight
    fraWksnControls.Visible = False
    fraWksnConsoleControls.Visible = True
    
    lblConsoleSel(0).Caption = " " & funLoadResString(10610)
    mnuWksnCrdRecharge.Visible = False
    mnuWksnConnect.Visible = False
    mnuWksnPaySession.Visible = False
    mnuWksnCancelPaySession.Visible = False
    mnuWksnCrdDisable.Visible = False
    mnuWksnCrdHistory.Visible = False
    mnuWksnStartSelfServPrePaid.Visible = False
    mnuWksnClose.enabled = True
    mnuWksnPause.enabled = True
    mnuWksnLock.enabled = True
    mnuWksnFollow.enabled = True
    mnuWksnPlay.enabled = True
    mnuWksnActivate.enabled = True
    mnuWksnBloc.enabled = True
    mnuWksnUnblock.enabled = guServOpts.uConnectedUser.AccessUnblock
    mnuWksnUserRelease.Visible = False
    
    mnuWksnActivateDD.Visible = mnuWksnActivate.Visible
    mnuWksnBlocDD.Visible = mnuWksnBloc.Visible
    mnuWksnCrdRechargeDD.Visible = mnuWksnCrdRecharge.Visible
    mnuWksnCrdDisableDD.Visible = mnuWksnCrdDisable.Visible
    mnuWksnCrdHistoryDD.Visible = mnuWksnCrdHistory.Visible
    mnuWksnCloseDD.Visible = mnuWksnClose.Visible
    mnuWksnFollowDD.Visible = mnuWksnFollow.Visible
    mnuWksnLockDD.Visible = mnuWksnLock.Visible
    mnuWksnPauseDD.Visible = mnuWksnPause.Visible
    mnuWksnPaySessionDD.Visible = mnuWksnPaySession.Visible
    mnuWksnPlayDD.Visible = mnuWksnPlay.Visible
    mnuWksnShutdownDD.Visible = mnuWksnShutdown.Visible
    mnuWksnUnblockDD.Visible = mnuWksnUnblock.Visible
    mnuWksnStartSelfServPrePaidDD.Visible = mnuWksnStartSelfServPrePaid.Visible
    mnuWksnBootDD.Visible = mnuWksnBootDD.enabled
    mnuWksnUserReleaseDD.Visible = mnuWksnUserRelease.Visible
    mnu10.Visible = mnuWksnUnblockDD.Visible Or mnuWksnPlayDD.Visible Or mnuWksnPaySessionDD.Visible Or mnuWksnPauseDD.Visible Or mnuWksnLockDD.Visible Or mnuWksnFollowDD.Visible Or mnuWksnCloseDD.Visible Or mnuWksnCrdHistoryDD.Visible Or mnuWksnCrdDisableDD.Visible Or mnuWksnCrdRechargeDD.Visible
    mnu9.Visible = mnuWksnBootDD.Visible
    Call subUpdateUIConsole
    Exit Sub
    
    
fraWksnConsole_ClickError:
    Call funErrorManager("frmMain.fraWksnConsole_ClickError", True)
End Sub


'*******************************************************************************
Private Sub fraWksnConsole_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

    If Not fraWksnConsoleControls.Visible Then Call fraWksnConsole_Click
    If Button = vbRightButton Then Call PopupMenu(mnuWksnDropDown, vbPopupMenuLeftButton)
End Sub

'*******************************************************************************
Public Sub fsbWksn_Change()
On Error GoTo fsbWksn_ChangeError


    fraWksn.top = -fsbWksn.value * (ucWorkstations(1).height + 60)
    Exit Sub
    
fsbWksn_ChangeError:
    Call funErrorManager("frmMain.fsbWksn_ChangeError", True)
End Sub


'*******************************************************************************
Private Sub fxgCards_Click()
On Error GoTo fxgCards_ClickError
    
    If fxgCards.MouseRow = 0 Then Exit Sub
    If guServOpts.lngSelectedCardIndex = 0 Then Call fxgCards_RowColChange
    Exit Sub
    
fxgCards_ClickError:
    Call funErrorManager("frmMain.fxgCards_ClickError", True)
End Sub


'*******************************************************************************
Private Sub fxgCards_DblClick()
On Error GoTo fxgCards_DblClickError
    
    If fxgCards.MouseRow = 0 Then
        Call tmrCrdRefresh_Timer
    Else
'        Call cmdCrdButtons(1)_Click
    End If
    Exit Sub
    
fxgCards_DblClickError:
    Call funErrorManager("frmMain.fxgCards_DblClickError", True)
End Sub


'*******************************************************************************
Private Sub fxgCards_GotFocus()
On Error Resume Next

    cmdCrdFind.Default = True
End Sub


'*******************************************************************************
Private Sub fxgCards_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lngNewRow As Long
On Error GoTo fxgCards_MouseDownError

    
    lngNewRow = fxgCards.TopRow + Int(y / fxgCards.RowHeight(0)) - 1
    If lngNewRow <> fxgCards.Row And lngNewRow < fxgCards.Rows Then
        fxgCards.Row = lngNewRow
        Call fxgCards_RowColChange
    End If
    If Button = vbRightButton And y >= fxgCards.RowHeight(0) And lngNewRow < fxgCards.Rows Then
        Call PopupMenu(mnuCrdDropDown, vbPopupMenuLeftButton)
    ElseIf Shift = vbAltMask And Button = vbLeftButton Then
        cmdCrdButtons(5).value = True
    ElseIf Shift = vbCtrlMask And Button = vbLeftButton Then
        cmdCrdButtons(1).value = True
    End If
    Exit Sub
    
fxgCards_MouseDownError:
    Call funErrorManager("frmMain.fxgCards_MouseDownError", True)
End Sub


'*******************************************************************************
Public Sub fxgCards_RowColChange()
Dim lngNewSelectedRowIndex As Long
Dim i As Integer
On Error GoTo fxgCards_RowColChangeError
    
    With fxgCards
        lngNewSelectedRowIndex = .Row
        If lngNewSelectedRowIndex < 1 Then Exit Sub
        If guServOpts.lngSelectedCardIndex > .Rows - 1 Then guServOpts.lngSelectedCardIndex = .Rows - 1
    
        .Col = 0
        If guServOpts.lngSelectedCardIndex > 0 Then
            .Row = guServOpts.lngSelectedCardIndex
            .ColSel = .Cols - 1
            .CellForeColor = IIf(i = guServOpts.lngSelectedCardIndex, vbHighlightText, IIf(guCards(guServOpts.lngSelectedCardIndex).Active, vbWindowText, vbGrayText))
            .CellBackColor = vbWindowBackground
        End If
    
        guServOpts.lngSelectedCardIndex = lngNewSelectedRowIndex
        
        .Row = lngNewSelectedRowIndex
        .ColSel = .Cols - 1
        .CellForeColor = vbHighlightText
        .CellBackColor = vbHighlight
        
'        If .TopRow > lngNewSelectedRowIndex Then
'            .TopRow = lngNewSelectedRowIndex
'        ElseIf .TopRow + Int(.height / .RowHeight(0) - 2) - 1 < lngNewSelectedRowIndex Then
'            .TopRow = lngNewSelectedRowIndex - Int(.height / .RowHeight(0) - 2) + 1
'        End If
        
        If .TopRow > lngNewSelectedRowIndex Then
            .TopRow = lngNewSelectedRowIndex
        ElseIf .TopRow + Int(.height / .RowHeight(0) - 2) < lngNewSelectedRowIndex Then  '- 1
            .TopRow = lngNewSelectedRowIndex - Int(.height / .RowHeight(0) - 2) '+ 1
        End If
        
        
        Call cmbCrdType.Clear
        Call cmbCrdType.AddItem(funLoadResString(10526))
        cmbCrdType.ItemData(0) = 1
        i = funArraySize(guServOpts.uCards.strCardTypes)
        If i > 0 Then
            For i = 0 To funArraySize(guServOpts.uCards.strCardTypes) - 1
                Call cmbCrdType.AddItem(guServOpts.uCards.strCardTypes(i))
            Next i
        End If
        Call subUpdateUISelectedCard
    End With
    Exit Sub
    
fxgCards_RowColChangeError:
    Call funErrorManager("frmMain.fxgCards_RowColChangeError", True)
End Sub


'*******************************************************************************
Private Sub fxgClients_Click()
On Error GoTo fxgClients_ClickError
    
    If fxgClients.MouseRow = 0 Then Exit Sub
    If guServOpts.lngSelectedClientIndex = 0 Then Call fxgClients_RowColChange
    Exit Sub
    
fxgClients_ClickError:
    Call funErrorManager("frmMain.fxgClients_ClickError", True)
End Sub


'*******************************************************************************
Private Sub fxgClients_DblClick()
On Error GoTo fxgClients_DblClickError
    
    If fxgClients.MouseRow = 0 Then
        Call tmrCliRefresh_Timer
    End If
    Exit Sub
    
fxgClients_DblClickError:
    Call funErrorManager("frmMain.fxgClients_DblClickError", True)
End Sub


'*******************************************************************************
Private Sub fxgClients_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lngNewRow As Long
On Error GoTo fxgClients_MouseDownError

    
    lngNewRow = fxgClients.TopRow + Int(y / fxgClients.RowHeight(0)) - 1
    If lngNewRow <> fxgClients.Row And lngNewRow < fxgClients.Rows Then
        fxgClients.Row = lngNewRow
        Call fxgClients_RowColChange
    End If
    If Button = vbRightButton And y >= fxgClients.RowHeight(0) And lngNewRow < fxgClients.Rows Then
        Call PopupMenu(mnuCliDropDown, vbPopupMenuLeftButton)
    ElseIf Shift = vbAltMask And Button = vbLeftButton Then
        cmdCliButtons(5).value = True
    ElseIf Shift = vbCtrlMask And Button = vbLeftButton Then
        cmdCliButtons(1).value = True
    End If
    Exit Sub
    
fxgClients_MouseDownError:
    Call funErrorManager("frmMain.fxgClients_MouseDownError", True)
End Sub


'*******************************************************************************
Public Sub fxgClients_RowColChange()
Dim lngNewSelectedRowIndex As Long
Dim i As Integer
On Error GoTo fxgClients_RowColChangeError
    
    With fxgClients
        lngNewSelectedRowIndex = .Row
        If lngNewSelectedRowIndex < 1 Then Exit Sub
        If guServOpts.lngSelectedClientIndex > .Rows - 1 Then guServOpts.lngSelectedClientIndex = .Rows - 1
    
        .Col = 0
        If guServOpts.lngSelectedClientIndex > 0 Then
            .Row = guServOpts.lngSelectedClientIndex
            .ColSel = .Cols - 1
            .CellForeColor = IIf(i = guServOpts.lngSelectedClientIndex, vbHighlightText, IIf(guClients(guServOpts.lngSelectedClientIndex).Active, vbWindowText, vbGrayText))
            .CellBackColor = vbWindowBackground
        End If
    
        guServOpts.lngSelectedClientIndex = lngNewSelectedRowIndex
        
        .Row = lngNewSelectedRowIndex
        .ColSel = .Cols - 1
        .CellForeColor = vbHighlightText
        .CellBackColor = vbHighlight
        
        If .TopRow > lngNewSelectedRowIndex Then
            .TopRow = lngNewSelectedRowIndex
        ElseIf .TopRow + Int(.height / .RowHeight(0) - 2) - 1 < lngNewSelectedRowIndex Then
            .TopRow = lngNewSelectedRowIndex - Int(.height / .RowHeight(0) - 2) + 1
        End If
        
        Call subUpdateUISelectedClient
    End With
    Exit Sub
    
fxgClients_RowColChangeError:
    Call funErrorManager("frmMain.fxgClients_RowColChangeError", True)
End Sub


'*******************************************************************************
Private Sub fxgCrdPending_Click()
On Error GoTo fxgCrdPending_ClickError
    
    If fxgCrdPending.MouseRow = 0 Then Exit Sub
    If guServOpts.lngSelectedCardPendingIndex = 0 Then
        Call fxgCrdPending_RowColChange
    Else
        txtCrdPending.Text = fxgCrdPending.TextMatrix(guServOpts.lngSelectedCardPendingIndex, 0)
        txtCrdPending.SelStart = 0
        txtCrdPending.SelLength = Len(txtCrdPending.Text)
        cmdCrdPendingEnable.enabled = True
        cmdCrdPendingEnable.Default = True
    End If
    Exit Sub
    
fxgCrdPending_ClickError:
    Call funErrorManager("frmMain.fxgCrdPending_ClickError", True)
End Sub


'*******************************************************************************
Private Sub fxgCrdPending_DblClick()
On Error GoTo fxgCrdPending_DblClickError

    If guServOpts.lngSelectedCardPendingIndex <> 0 Then
        cmdCrdPendingEnable.value = True
    End If
    Exit Sub
    
fxgCrdPending_DblClickError:
    Call funErrorManager("frmMain.fxgCrdPending_DblClickError", True)
End Sub


'*******************************************************************************
Private Sub fxgCrdPending_GotFocus()

    mblnfxgCrdPendingHasFocus = True
End Sub


'*******************************************************************************
Private Sub fxgCrdPending_LostFocus()

    mblnfxgCrdPendingHasFocus = False
End Sub


'*******************************************************************************
Private Sub fxgCrdPending_RowColChange()
Dim lngNewSelectedRowIndex As Long
On Error GoTo fxgCrdPending_RowColChangeError

    With fxgCrdPending
        lngNewSelectedRowIndex = .Row
        If lngNewSelectedRowIndex = .Rows Or lngNewSelectedRowIndex < 1 Then
            .Row = guServOpts.lngSelectedCardPendingIndex
            Exit Sub
        End If
        If guServOpts.lngSelectedCardPendingIndex > .Rows - 1 Then guServOpts.lngSelectedCardPendingIndex = .Rows - 1
        
        .Col = 0
        If guServOpts.lngSelectedCardPendingIndex > 0 Then
            .Row = guServOpts.lngSelectedCardPendingIndex
            .ColSel = .Cols - 1
            .CellForeColor = vbWindowText
            .CellBackColor = vbWindowBackground
        End If
        
        guServOpts.lngSelectedCardPendingIndex = lngNewSelectedRowIndex
        
        .Row = lngNewSelectedRowIndex
        .ColSel = .Cols - 1
        .CellForeColor = vbHighlightText
        .CellBackColor = vbHighlight
        
        If .TopRow > lngNewSelectedRowIndex Then
            .TopRow = lngNewSelectedRowIndex
        ElseIf .TopRow + Int(.height / .RowHeight(0) - 2) < lngNewSelectedRowIndex Then  '- 1
            .TopRow = lngNewSelectedRowIndex - Int(.height / .RowHeight(0) - 2) '+ 1
        End If
        
        If mblnfxgCrdPendingHasFocus Then
            txtCrdPending.Text = .TextMatrix(guServOpts.lngSelectedCardPendingIndex, 0)
            txtCrdPending.SelStart = 0
            txtCrdPending.SelLength = Len(txtCrdPending.Text)
            cmdCrdPendingEnable.enabled = True
            cmdCrdPendingEnable.Default = True
        End If

    End With
    Exit Sub
    
fxgCrdPending_RowColChangeError:
    Call funErrorManager("frmCardHistory.fxgCrdPending_RowColChangeError", True)
End Sub


'*******************************************************************************
Private Sub fxgServices_Click()
On Error GoTo fxgServices_ClickError
    
    If fxgServices.MouseRow = 0 Then Exit Sub
    If guServOpts.lngSelectedServiceIndex = 0 Then Call fxgServices_RowColChange
    Exit Sub
    
fxgServices_ClickError:
    Call funErrorManager("frmMain.fxgServices_ClickError", True)
End Sub


'*******************************************************************************
Private Sub fxgServices_RowColChange()
Dim lngNewSelectedRowIndex As Long
Dim uOldService, uCurrentService As New clsService
On Error GoTo fxgServices_RowColChangeError
    
    lngNewSelectedRowIndex = fxgServices.Row
    If lngNewSelectedRowIndex < 1 Then Exit Sub
    If guServOpts.lngSelectedServiceIndex > fxgServices.Rows - 1 Then guServOpts.lngSelectedServiceIndex = fxgServices.Rows - 1

    Set uCurrentService = guServices(lngNewSelectedRowIndex)
    
    If guServOpts.lngSelectedServiceIndex > 0 Then
        Set uOldService = guServices(guServOpts.lngSelectedServiceIndex)
        fxgServices.Col = 0
        fxgServices.Row = guServOpts.lngSelectedServiceIndex
        fxgServices.ColSel = fxgServices.Cols - 1
        fxgServices.CellForeColor = vbWindowText
        fxgServices.CellBackColor = vbWindowBackground
        fxgServices.Col = 2
        fxgServices.CellForeColor = IIf(uOldService.StockMode = enuServiceStockMode.Managed And uOldService.Stock <= uOldService.StockMinAlert, vbRed, vbWindowText)
    End If

    guServOpts.lngSelectedServiceIndex = lngNewSelectedRowIndex
    
    fxgServices.Col = 0
    fxgServices.Row = lngNewSelectedRowIndex
    fxgServices.ColSel = fxgServices.Cols - 1
    fxgServices.CellForeColor = vbHighlightText
    fxgServices.CellBackColor = vbHighlight
    
    If fxgServices.TopRow > lngNewSelectedRowIndex Then
        fxgServices.TopRow = lngNewSelectedRowIndex
    ElseIf fxgServices.TopRow + guServOpts.intNumVisRows - 3 < lngNewSelectedRowIndex Then
        fxgServices.TopRow = lngNewSelectedRowIndex - guServOpts.intNumVisRows + 3
    End If
    
    lblServSel(0).Caption = " " & uCurrentService.Label
    lblServSel(1).Caption = " " & funFormatCurrencyString(uCurrentService.Price, guServOpts.uConfig.strMonFormat, guServOpts.uConfig.strMonSymbol)
    lblServSel(2).Caption = " " & IIf(uCurrentService.StockMode = enuServiceStockMode.Managed, uCurrentService.Stock, funLoadResString(10332))
    lblServSel(2).ForeColor = IIf(uCurrentService.StockMode = enuServiceStockMode.Managed And uCurrentService.Stock <= uCurrentService.StockMinAlert, vbRed, vbWindowText)
    udServicesControlsQuantity.value = 0
    udServicesControlsQuantity.enabled = True
    Exit Sub
    
fxgServices_RowColChangeError:
    Call funErrorManager("frmMain.fxgServices_RowColChangeError", True)
End Sub


'*******************************************************************************
Private Sub fxgSession_Click(Index As Integer)
On Error GoTo fxgSession_ClickError

    Select Case Index
    Case 4
        If guServOpts.lngSelectedURLIndex = 0 Then Call fxgSession_RowColChange(Index)
    End Select
    Exit Sub
    
fxgSession_ClickError:
    Call funErrorManager("frmCardHistory.fxgSession_ClickError", True)
End Sub


'*******************************************************************************
Private Sub fxgSession_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lngNewRow As Long
On Error GoTo fxgSession_MouseDownError

    Select Case Index
    Case 4
        lngNewRow = fxgSession(Index).TopRow + Int(y / fxgSession(Index).RowHeight(0)) - 1
        If lngNewRow <> fxgSession(Index).Row And lngNewRow < fxgSession(Index).Rows Then
            fxgSession(Index).Row = lngNewRow
            Call fxgSession_RowColChange(Index)
        End If
        If Button = vbRightButton And y >= fxgSession(Index).RowHeight(0) And lngNewRow < fxgSession(Index).Rows Then
            Call PopupMenu(mnuSessionURLDropDown)
        End If
    End Select
    Exit Sub
    
fxgSession_MouseDownError:
    Call funErrorManager("frmCardHistory.fxgSession_MouseDownError", True)
End Sub


'*******************************************************************************
Private Sub fxgSession_RowColChange(Index As Integer)
Dim lngNewSelectedRowIndex As Long
On Error GoTo fxgSession_RowColChangeError

    Select Case Index
    Case 4
        With fxgSession(4)
            lngNewSelectedRowIndex = .Row
            If lngNewSelectedRowIndex = .Rows Or lngNewSelectedRowIndex < 1 Then
                .Row = guServOpts.lngSelectedURLIndex
                Exit Sub
            End If
            If guServOpts.lngSelectedURLIndex > .Rows - 1 Then guServOpts.lngSelectedURLIndex = .Rows - 1
            
            .Col = 0
            If guServOpts.lngSelectedURLIndex > 0 Then
                .Row = guServOpts.lngSelectedURLIndex
                .ColSel = .Cols - 1
                .CellForeColor = vbWindowText
                .CellBackColor = vbWindowBackground
            End If
            
            guServOpts.lngSelectedURLIndex = lngNewSelectedRowIndex
            
            .Row = lngNewSelectedRowIndex
            .ColSel = .Cols - 1
            .CellForeColor = vbHighlightText
            .CellBackColor = vbHighlight
            
            If .TopRow > lngNewSelectedRowIndex Then
                .TopRow = lngNewSelectedRowIndex
            ElseIf .TopRow + Int(.height / .RowHeight(0) - 2) - 1 < lngNewSelectedRowIndex Then
                .TopRow = lngNewSelectedRowIndex - Int(.height / .RowHeight(0) - 2) + 1
            End If
        End With
    End Select
    Exit Sub
    
fxgSession_RowColChangeError:
    Call funErrorManager("frmCardHistory.fxgSession_RowColChangeError", True)
End Sub


'*******************************************************************************
Private Sub imgConsole_Click()
On Error GoTo imgConsole_ClickError

    Call fraWksnConsole_Click
    Exit Sub
    
imgConsole_ClickError:
    Call funErrorManager("frmMain.imgConsole_ClickError", True)
End Sub


'*******************************************************************************
Private Sub imgConsole_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

    If Not fraWksnConsoleControls.Visible Then Call fraWksnConsole_Click
    If Button = vbRightButton Then Call PopupMenu(mnuWksnDropDown, vbPopupMenuLeftButton)
End Sub

'*******************************************************************************
Private Sub lblConsoleName_Click()
On Error GoTo lblConsoleName_ClickError

    Call fraWksnConsole_Click
    Exit Sub
    
lblConsoleName_ClickError:
    Call funErrorManager("frmMain.lblConsoleName_ClickError", True)
End Sub


'*******************************************************************************
Private Sub lblConsoleName_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

    If Not fraWksnConsoleControls.Visible Then Call fraWksnConsole_Click
    If Button = vbRightButton Then Call PopupMenu(mnuWksnDropDown, vbPopupMenuLeftButton)
End Sub


'*******************************************************************************
Private Sub lblConsoleWksnControl_Click()
On Error GoTo lblConsoleWksnControl_ClickError

    Call fraWksnConsole_Click
    Exit Sub
    
lblConsoleWksnControl_ClickError:
    Call funErrorManager("frmMain.lblConsoleWksnControl_ClickError", True)
End Sub


'*******************************************************************************
Private Sub lblConsoleWksnControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

    If Not fraWksnConsoleControls.Visible Then Call fraWksnConsole_Click
    If Button = vbRightButton Then Call PopupMenu(mnuWksnDropDown, vbPopupMenuLeftButton)
End Sub


'*******************************************************************************
Private Sub mnuAbout_Click()
On Error GoTo mnuAbout_ClickError
    
    Call frmAbout.Show(vbModal, frmMain)
    Exit Sub
    
mnuAbout_ClickError:
    Call funErrorManager("frmMain.mnuAbout_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuCli_Click(Index As Integer)
Dim intWorkstationIndex As Integer
Dim strBuffer As String
On Error GoTo mnuCli_ClickError

    Select Case Index
    Case Is <= 6
        cmdCliButtons(Index).value = True
    Case 7
        Call funSetFocus(fxgClients)
        If guServOpts.lngSelectedClientIndex = 0 Then
            Call funMsg(enuMsg.ClientNotSelected)
        Else
            If funMsg(AskCanDeleteClient, guClients(guServOpts.lngSelectedClientIndex).login) = vbYes Then
                intWorkstationIndex = funClientInUse(guClients(guServOpts.lngSelectedClientIndex))
                If intWorkstationIndex <> 0 Then
                    strBuffer = CNX_SESSION_CLOSE & CNX_MSG_END
                    Call funSendData(guWorkstations(intWorkstationIndex).Cnx, strBuffer)
                    DoEvents
                End If
                Call subDBDeleteClient(guClients(guServOpts.lngSelectedClientIndex))
                Call subFillClients
                If guServOpts.lngSelectedClientIndex > guClients.Count Then
                    guServOpts.lngSelectedClientIndex = guClients.Count
                    Call fxgClients_RowColChange
                End If
                Call subUpdateUISelectedClient
            End If
        End If
    Case 9
        Call tmrCliRefresh_Timer
    Case 10
        If funMsg(enuMsg.AskCanPurgeClients) = vbYes Then
            Call subClientsPurge
            Call subFillClients
        End If
    End Select
    Exit Sub
    
mnuCli_ClickError:
    Call funErrorManager("frmMain.mnuCli_ClickError - " & Index, True)
End Sub


'*******************************************************************************
Private Sub mnuCliDD_Click(Index As Integer)
On Error Resume Next

    Call mnuCli_Click(Index)
End Sub


'*******************************************************************************
Private Sub mnuCliModifyDeleteDD_Click()
On Error Resume Next

    Call mnuCli_Click(7)
End Sub


'*******************************************************************************
Private Sub mnuCliModifyModifyDD_Click()
On Error Resume Next

    Call mnuCli_Click(2)
End Sub


'*******************************************************************************
Private Sub mnuCrd_Click(Index As Integer)
On Error GoTo mnuCrd_ClickError

    Select Case Index
    Case Is <= 6
        cmdCrdButtons(Index).value = True
    Case 9
        Call tmrCrdRefresh_Timer
    Case 10
        Call frmCardMassAdd.Show(vbModal, Me)
    Case 11
        If funMsg(enuMsg.AskCanPurgeCards) = vbYes Then
            Call subCardsPurge
            Call subFillCards
        End If
    End Select
    Exit Sub

mnuCrd_ClickError:
    Call funErrorManager("frmMain.mnuCrd_ClickError - " & Index, True)
End Sub


'*******************************************************************************
Private Sub mnuChangePasswd_Click()
On Error GoTo mnuChangePasswd_ClickError

    Call frmPasswdChange.Show(vbModal, frmMain)
    Exit Sub

mnuChangePasswd_ClickError:
    Call funErrorManager("frmMain.mnuChangePasswd_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuChangePasswdUser_Click()
On Error GoTo mnuChangePasswdUser_ClickError

    Call frmPasswdChangeUser.Show(vbModal, frmMain)
    Exit Sub

mnuChangePasswdUser_ClickError:
    Call funErrorManager("frmMain.mnuChangePasswdUser_ClickError", True)
End Sub


Private Sub mnuCrdDD_Click(Index As Integer)
On Error Resume Next

    Call mnuCrd_Click(Index)
End Sub


'*******************************************************************************
Private Sub mnuIncomePrint_Click()
On Error GoTo mnuIncomePrint_ClickError
    
    cmdIncomePrint.value = True
    Exit Sub
    
mnuIncomePrint_ClickError:
    Call funErrorManager("frmMain.mnuIncomePrint_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuMinimizeToSysTray_Click()
On Error GoTo mnuMinimizeToSysTray_ClickError

    Call mnuSysTrayHide_Click
    Exit Sub

mnuMinimizeToSysTray_ClickError:
    Call funErrorManager("frmMain.mnuMinimizeToSysTray_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuSessionPrint_Click()
On Error GoTo mnuSessionPrint_ClickError
    
    cmdSessionPrint.value = True
    Exit Sub
    
mnuSessionPrint_ClickError:
    Call funErrorManager("frmMain.mnuSessionPrint_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuSysTrayHide_Click()
On Error GoTo mnuSysTrayHide_ClickError

    If gblnfrmCardHistory Or gblnfrmCardMassAdd Or gblnfrmLogon Or gblnfrmRecharge Or gblnfrmSetupWizard Or gblnfrmOptions Then Exit Sub
    Call frmMain.Hide '.Visible = False
    tmrSystTrayLogoff.enabled = True
    tmrSystTrayUpdate.enabled = True
    mblnLogoffOnShow = False
    Exit Sub

mnuSysTrayHide_ClickError:
    Call funErrorManager("frmMain.mnuSysTrayHide_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuSysTrayShow_Click()
On Error GoTo mnuSysTrayShow_ClickError

    Call frmMain.Show '.Visible = True
    tmrSystTrayLogoff.enabled = False
    tmrSystTrayUpdate.enabled = False
    guIconData.hIcon = imgSystTrayIcons.ListImages(1).Picture
    guIconData.szTip = funLoadResString(10000) & Chr(0)
    Call Shell_NotifyIcon(NIM_MODIFY, guIconData)
    If mblnLogoffOnShow Then
        Call funUserLogoff(guServOpts.uConnectedUser, False)
    End If
    Exit Sub

mnuSysTrayShow_ClickError:
    Call funErrorManager("frmMain.mnuSysTrayShow_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuWksnActivate_Click()
On Error GoTo mnuWksnActivate_ClickError

    If funAnyActiveWorkstation Then
        cmdWksnButtons(0).value = True
    Else
        cmdConsoleButtons(0).value = True
    End If
    Exit Sub

mnuWksnActivate_ClickError:
    Call funErrorManager("frmMain.mnuWksnActivate_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuWksnActivateConsole_Click()
On Error GoTo mnuWksnActivateConsole_ClickError

    guServOpts.uFlags.blnConsoleActive = Not guServOpts.uFlags.blnConsoleActive
    mnuWksnActivateConsole.Checked = guServOpts.uFlags.blnConsoleActive
    mnuWksnShowConsole.Visible = guServOpts.uFlags.blnConsoleActive
    fraWksnConsole.Visible = guServOpts.uFlags.blnConsoleActive
    If funAnyActiveWorkstation And guServOpts.uFlags.blnConsoleActive Then
        Call fraWksnConsole_Click
    Else
        Call ucWorkstations_Click(1)
    End If
    Call subResizeWorkstations
    Exit Sub

mnuWksnActivateConsole_ClickError:
    Call funErrorManager("frmMain.mnuWksnActivateConsole_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuWksnActivateDD_Click()
On Error Resume Next
    
    Call mnuWksnActivate_Click
End Sub


'*******************************************************************************
Private Sub mnuWksnBloc_Click()
On Error GoTo mnuWksnBloc_ClickError

    If funAnyActiveWorkstation Then
        cmdWksnButtons(2).value = True
    Else
        cmdConsoleButtons(1).value = True
    End If
    Exit Sub

mnuWksnBloc_ClickError:
    Call funErrorManager("frmMain.mnuWksnBloc_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuWksnBlocDD_Click()
On Error Resume Next

    Call mnuWksnBloc_Click
End Sub


'*******************************************************************************
Private Sub mnuWksnBoot_Click()
On Error GoTo mnuWksnBoot_ClickError

    Call frmWksnBoot.Show(vbModal, frmMain)
    Exit Sub

mnuWksnBoot_ClickError:
    Call funErrorManager("frmMain.mnuWksnBoot_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuWksnBootDD_Click()
Dim uWOL As clsMagicPacket
On Error GoTo mnuWksnBootDD_ClickError
    
    If funAnyActiveWorkstation Then
        Set uWOL = New clsMagicPacket
        uWOL.WinsockControl = wskWakeOnLan
        uWOL.NICAddress = guSelectedWorkstation.ComputerMacAddress
        Call uWOL.WakeUp
        Set uWOL = Nothing
    Else
        Call mnuWksnBoot_Click
    End If
    Exit Sub

mnuWksnBootDD_ClickError:
    Call funErrorManager("frmMain.mnuWksnBootDD_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuWksnCancelPaySession_Click()
On Error GoTo mnuWksnCancelPaySession_ClickError

    cmdWksnButtons(15).value = True
    Exit Sub

mnuWksnCancelPaySession_ClickError:
    Call funErrorManager("frmMain.mnuWksnCancelPaySession_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuWksnCrdRecharge_Click()
On Error GoTo mnuWksnCrdRecharge_ClickError

    cmdWksnButtons(9).value = True
    Exit Sub

mnuWksnCrdRecharge_ClickError:
    Call funErrorManager("frmMain.mnuWksnCrdRecharge_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuWksnCrdRechargeDD_Click()
On Error Resume Next
    
    Call mnuWksnCrdRecharge_Click
End Sub


'*******************************************************************************
Private Sub mnuWksnCrdDisable_Click()
On Error GoTo mnuWksnCrdDisable_ClickError

    cmdWksnButtons(10).value = True
    Exit Sub

mnuWksnCrdDisable_ClickError:
    Call funErrorManager("frmMain.mnuWksnCrdDisable_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuWksnCrdDisableDD_Click()
On Error Resume Next

    Call mnuWksnCrdDisable_Click
End Sub


'*******************************************************************************
Private Sub mnuWksnCrdHistory_Click()
On Error GoTo mnuWksnCrdHistory_ClickError

    Select Case guSelectedWorkstation.SessionType
    Case enuSessionType.Card
        cmdWksnButtons(11).value = True
    Case enuSessionType.Client
        cmdWksnButtons(12).value = True
    End Select
    Exit Sub

mnuWksnCrdHistory_ClickError:
    Call funErrorManager("frmMain.mnuWksnCrdHistory_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuWksnCrdHistoryDD_Click()
On Error Resume Next

    Call mnuWksnCrdHistory_Click
End Sub


'*******************************************************************************
Private Sub mnuWksnClose_Click()
On Error GoTo mnuWksnClose_ClickError

    If funAnyActiveWorkstation Then
        cmdWksnButtons(4).value = True
    Else
        cmdConsoleButtons(3).value = True
    End If
    Exit Sub

mnuWksnClose_ClickError:
    Call funErrorManager("frmMain.mnuWksnClose_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuWksnCloseDD_Click()
On Error Resume Next

    Call mnuWksnClose_Click
End Sub


'*******************************************************************************
Private Sub mnuWksnConnect_Click()
On Error GoTo mnuWksnConnect_ClickError

    Call subToggleForceConnected
    Exit Sub

mnuWksnConnect_ClickError:
    Call funErrorManager("frmMain.mnuWksnConnect_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuWksnFollow_Click()
On Error GoTo mnuWksnFollow_ClickError

    If funAnyActiveWorkstation Then
        Call mnuWksnFollowOne_Click
    Else
        cmdConsoleButtons(7).value = True
    End If
    Exit Sub

mnuWksnFollow_ClickError:
    Call funErrorManager("frmMain.mnuWksnFollow_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuWksnFollowDD_Click()
On Error Resume Next

    Call mnuWksnFollow_Click
End Sub


'*******************************************************************************
Private Sub mnuWksnFollowOne_Click()
Dim strBuffer As String
On Error GoTo mnuWksnFollowOne_ClickError

    strBuffer = CNX_PAUSE_FOLLOW_WORKSTATION & CNX_MSG_END
    Call funSendData(guSelectedWorkstation.Cnx, strBuffer)
    Call funSetFocus(ucWorkstations(1))
    Exit Sub

mnuWksnFollowOne_ClickError:
    Call funErrorManager("frmMain.mnuWksnFollowOne_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuWksnLock_Click()
On Error GoTo mnuWksnLock_ClickError

    If funAnyActiveWorkstation Then
        Call mnuWksnLockOne_Click
    Else
        cmdConsoleButtons(6).value = True
    End If
    Exit Sub

mnuWksnLock_ClickError:
    Call funErrorManager("frmMain.mnuWksnLock_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuWksnLockDD_Click()
On Error Resume Next

    Call mnuWksnLock_Click
End Sub


'*******************************************************************************
Private Sub mnuWksnLockOne_Click()
Dim strBuffer As String
On Error GoTo mnuWksnLockOne_ClickError

    strBuffer = CNX_PAUSE_LOCK_WORKSTATION & CNX_MSG_END
    Call funSendData(guSelectedWorkstation.Cnx, strBuffer)
    Call funSetFocus(ucWorkstations(1))
    Exit Sub
    
mnuWksnLockOne_ClickError:
    Call funErrorManager("frmMain.mnuWksnLockOne_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuWksnPause_Click()
On Error GoTo mnuWksnPause_ClickError

    If funAnyActiveWorkstation Then
        cmdWksnButtons(5).value = True
    Else
        cmdConsoleButtons(5).value = True
    End If
    Exit Sub

mnuWksnPause_ClickError:
    Call funErrorManager("frmMain.mnuWksnPause_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuWksnPauseDD_Click()
On Error Resume Next

    Call mnuWksnPause_Click
End Sub


'*******************************************************************************
Private Sub mnuWksnPauseOne_Click()
On Error GoTo mnuWksnPauseOne_ClickError

    cmdWksnButtons(5).value = True
    Exit Sub

mnuWksnPauseOne_ClickError:
    Call funErrorManager("frmMain.mnuWksnPauseOne_ClickError", True)
End Sub

'*******************************************************************************
Private Sub mnuWksnPaySession_Click()
On Error GoTo mnuWksnPaySession_ClickError

    cmdWksnButtons(1).value = True
    Exit Sub

mnuWksnPaySession_ClickError:
    Call funErrorManager("frmMain.mnuWksnPaySession_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuWksnPaySessionDD_Click()
On Error Resume Next

    Call mnuWksnPaySession_Click
End Sub


'*******************************************************************************
Private Sub mnuWksnPlay_Click()
On Error GoTo mnuWksnPlay_ClickError

    If funAnyActiveWorkstation Then
        cmdWksnButtons(8).value = True
    Else
        cmdConsoleButtons(4).value = True
    End If
    Exit Sub

mnuWksnPlay_ClickError:
    Call funErrorManager("frmMain.mnuWksnPlay_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuWksnPlayDD_Click()
On Error Resume Next

    Call mnuWksnPlay_Click
End Sub


'*******************************************************************************
Private Sub mnuWksnShowConsole_Click()
On Error GoTo mnuWksnShowConsole_ClickError
    
    Call fraWksnConsole_Click
    Exit Sub

mnuWksnShowConsole_ClickError:
    Call funErrorManager("frmMain.mnuWksnShowConsole_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuWksnShowNext_Click()
Dim i As Integer
Dim intActiveWorkstation As Integer
On Error GoTo mnuWksnShowNext_ClickError
    
    For i = 1 To ucWorkstations.Count
        If ucWorkstations(i).Active Then intActiveWorkstation = i
    Next i
    
    If intActiveWorkstation = ucWorkstations.UBound Then
        Call ucWorkstations_Click(1)
    Else
        Call ucWorkstations_Click(intActiveWorkstation + 1)
    End If
    Exit Sub

mnuWksnShowNext_ClickError:
    Call funErrorManager("frmMain.mnuWksnShowNext_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuWksnShowPrev_Click()
Dim i As Integer
Dim intActiveWorkstation As Integer
On Error GoTo mnuWksnShowPrev_ClickError
    
    For i = 1 To ucWorkstations.Count
        If ucWorkstations(i).Active Then intActiveWorkstation = i
    Next i
    
    If intActiveWorkstation = ucWorkstations.LBound Then
        Call ucWorkstations_Click(ucWorkstations.UBound)
    ElseIf intActiveWorkstation = 0 Then
        Call ucWorkstations_Click(1)
    Else
        Call ucWorkstations_Click(intActiveWorkstation - 1)
    End If
    Exit Sub

mnuWksnShowPrev_ClickError:
    Call funErrorManager("frmMain.mnuWksnShowPrev_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuWksnShutdown_Click()
On Error GoTo mnuWksnShutdown_ClickError

    Call frmWksnShutdown.Show(vbModal, frmMain)
    Exit Sub

mnuWksnShutdown_ClickError:
    Call funErrorManager("frmMain.mnuWksnShutdown_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuWksnShutdownDD_Click()
On Error Resume Next

    Call mnuWksnShutdown_Click
End Sub


'*******************************************************************************
Private Sub mnuWksnSize_Click(Index As Integer)
Dim i As Integer
On Error GoTo mnuWksnSize_ClickError

    guServOpts.uFlags.intWksnSize = Index
    For i = 1 To mnuWksnSize.UBound
        mnuWksnSize(i).Checked = i = Index
    Next i
    Call subInitWorkstations
    Exit Sub

mnuWksnSize_ClickError:
    Call funErrorManager("frmMain.mnuWksnSize_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuWksnStartSelfServPrePaid_Click()
On Error GoTo mnuWksnStartSelfServPrePaid_ClickError

    If funAnyActiveWorkstation Then
        cmdWksnButtons(14).value = True
    Else
    
    End If
    Exit Sub

mnuWksnStartSelfServPrePaid_ClickError:
    Call funErrorManager("frmMain.mnuWksnStartSelfServPrePaid_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuWksnStartSelfServPrePaidDD_Click()
On Error Resume Next

    Call mnuWksnStartSelfServPrePaid_Click
End Sub


'*******************************************************************************
Private Sub mnuWksnUnblock_Click()
On Error GoTo mnuWksnUnblock_ClickError

    If Not guServOpts.uConnectedUser.AccessUnblock Then Exit Sub
    If funAnyActiveWorkstation Then
        cmdWksnButtons(3).value = True
    Else
        cmdConsoleButtons(2).value = True
    End If
    Exit Sub

mnuWksnUnblock_ClickError:
    Call funErrorManager("frmMain.mnuWksnUnblock_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuWksnUnblockDD_Click()
On Error Resume Next

    Call mnuWksnUnblock_Click
End Sub


'*******************************************************************************
Private Sub mnuCrdStartSessionWorkstation_Click(Index As Integer)
'Dim i As Integer
Dim uWorkstation As New clsWorkstation
Dim uCard As New clsCard
Dim uClient As New clsClient
On Error GoTo mnuCrdStartSessionWorkstation_ClickError

    Select Case tbsFunctions.SelectedItem.Key
    Case "k2"
        If guServOpts.lngSelectedCardIndex = 0 Then
            Call funMsg(enuMsg.CardNotSelected)
        Else
            Set uCard = guCards(guServOpts.lngSelectedCardIndex)
            If funCardInUse(uCard) <> 0 Then
                Call funMsg(enuMsg.NoRechargeCardInUse)
            Else
                If uCard.TimeCredit > NULL_DATE And uCard.Active Then
                    Set uWorkstation = funFindWorkstationByMac(guMacAddresses(Index))
                    Call subWorkstationCrdSessionOpen(uWorkstation, uCard)
                Else
                    Call funMsg(NoCredit)
                End If
            End If
        End If
        Call funSetFocus(fxgCards)
    Case "k3"
        If guServOpts.lngSelectedClientIndex = 0 Then
            Call funMsg(enuMsg.ClientNotSelected)
        Else
            Set uClient = guClients(guServOpts.lngSelectedClientIndex)
            If funClientInUse(uClient) <> 0 Then
                Call funMsg(enuMsg.NoRechargeCardInUse)
            Else
                If uClient.TimeCredit > NULL_DATE And uClient.Active Then
                    Set uWorkstation = funFindWorkstationByMac(guMacAddresses(Index))
                    Call subWorkstationCliSessionOpen(uWorkstation, uClient)
                Else
                    Call funMsg(NoCredit)
                End If
            End If
        End If
        Call funSetFocus(fxgClients)
    End Select
    
    Exit Sub
    
mnuCrdStartSessionWorkstation_ClickError:
    Call funErrorManager("frmMain.mnuCrdStartSessionWorkstation_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuCrdStartSessionFirstCilent_Click()
On Error GoTo mnuCrdStartSessionFirstCilent_ClickError

    Select Case tbsFunctions.SelectedItem.Image
    Case 2
        cmdCrdButtons(6).value = True
    Case 3
        cmdCliButtons(6).value = True
    End Select
    Exit Sub
    
mnuCrdStartSessionFirstCilent_ClickError:
    Call funErrorManager("frmMain.mnuCrdStartSessionFirstCilent_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuDBMaintenance_Click()
On Error GoTo mnuDBMaintenance_ClickError

    Call frmDBMaintenance.Show(vbModal, frmMain)
    Exit Sub
    
mnuDBMaintenance_ClickError:
    Call funErrorManager("frmMain.mnuDBMaintenance_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuFunctions_Click(Index As Integer)
'Dim i As Integer
On Error GoTo mnuFunctions_ClickError
    
    tbsFunctions.Tabs("k" & Index).Selected = True
    
    Exit Sub
    
mnuFunctions_ClickError:
    Call funErrorManager("frmMain.mnuFunctions_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuIncomeApply_Click()
On Error GoTo mnuIncomeApply_ClickError
    
    Call cmdIncomeFilter_Click
    Exit Sub
    
mnuIncomeApply_ClickError:
    Call funErrorManager("frmMain.mnuIncomeApply_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuIncomeExport_Click()
On Error GoTo mnuIncomeExport_ClickError
    
    cmdIncomeExport.value = True
    Exit Sub
    
mnuIncomeExport_ClickError:
    Call funErrorManager("frmMain.mnuIncomeExport_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuLogoff_Click()
On Error GoTo mnuLogoff_ClickError

    Call funUserLogoff(guServOpts.uConnectedUser)
    Exit Sub
    
mnuLogoff_ClickError:
    Call funErrorManager("frmMain.mnuLogoff_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuOptions_Click()
On Error GoTo mnuOptions_ClickError

    Call frmOptions.Show(vbModal, frmMain)
    Exit Sub
    
mnuOptions_ClickError:
    Call funErrorManager("frmMain.mnuOptions_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuQuit_Click()
On Error GoTo mnuQuit_ClickError
    
    Call Unload(Me)
    Exit Sub
    
mnuQuit_ClickError:
    Call funErrorManager("frmMain.mnuQuit_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuServicesCollect_Click()
On Error GoTo mnuServicesCollect_ClickError

    Call cmdServicesCollect_Click
    Exit Sub
    
mnuServicesCollect_ClickError:
    Call funErrorManager("frmMain.mnuServicesCollect_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuSessionApply_Click()
On Error GoTo mnuSessionApply_ClickError
    
    Call cmdSessionFilter_Click
    Exit Sub
    
mnuSessionApply_ClickError:
    Call funErrorManager("frmMain.mnuSessionApply_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuSessionCardNumCopy_Click()
On Error GoTo mnuSessionCardNumCopy_ClickError

    If gblnfrmCardHistory Then
    Else
        If guServOpts.lngSelectedURLIndex > 0 Then
            Call Clipboard.Clear
            Call Clipboard.SetText(fxgSession(4).TextMatrix(guServOpts.lngSelectedURLIndex, 2))
        End If
    End If
    Exit Sub
    
mnuSessionCardNumCopy_ClickError:
    Call funErrorManager("frmMain.mnuSessionCardNumCopy_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuSessionExport_Click()
On Error GoTo mnuSessionExport_ClickError
    
    cmdSessionExport.value = True
    Exit Sub
    
mnuSessionExport_ClickError:
    Call funErrorManager("frmMain.mnuSessionExport_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuSessionURLCopy_Click()
On Error GoTo mnuSessionURLCopy_ClickError

    If gblnfrmCardHistory Then
        If glngHistorySelectedURLIndex > 0 Then
            Call Clipboard.Clear
            Call Clipboard.SetText(frmCardHistory.fxgURLs.TextMatrix(glngHistorySelectedURLIndex, 3))
        End If
    Else
        If guServOpts.lngSelectedURLIndex > 0 Then
            Call Clipboard.Clear
            Call Clipboard.SetText(fxgSession(4).TextMatrix(guServOpts.lngSelectedURLIndex, 4))
        End If
    End If
    Exit Sub
    
mnuSessionURLCopy_ClickError:
    Call funErrorManager("frmMain.mnuSessionURLCopy_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuSessionURLDropDown_Click()
On Error GoTo mnuSessionURLOpen_ClickError

    mnuSessionCardNumCopy.enabled = Not gblnfrmCardHistory 'And IsNumeric(fxgSession(4).TextMatrix(guServOpts.lngSelectedURLIndex, 2))
    Exit Sub
    
mnuSessionURLOpen_ClickError:
    Call funErrorManager("frmMain.mnuSessionURLOpen_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuSessionURLOpen_Click()
On Error GoTo mnuSessionURLOpen_ClickError

    If gblnfrmCardHistory Then
        If glngHistorySelectedURLIndex > 0 Then Call subGotoWeb(frmCardHistory.fxgURLs.TextMatrix(glngHistorySelectedURLIndex, 3))
    Else
        If guServOpts.lngSelectedURLIndex > 0 Then Call subGotoWeb(fxgSession(4).TextMatrix(guServOpts.lngSelectedURLIndex, 4))
    End If
    Exit Sub
    
mnuSessionURLOpen_ClickError:
    Call funErrorManager("frmMain.mnuSessionURLOpen_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuUserAdmin_Click()
On Error GoTo mnuUserAdmin_ClickError

    Call frmUserAdmin.Show(vbModal, frmMain)
    Exit Sub

mnuUserAdmin_ClickError:
    Call funErrorManager("frmMain.mnuUserAdmin_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuWksnUserRelease_Click()
On Error GoTo mnuWksnUserRelease_ClickError

    Call subWorkstationRelease(guSelectedWorkstation)
    Exit Sub

mnuWksnUserRelease_ClickError:
    Call funErrorManager("frmMain.mnuWksnUserRelease_ClickError", True)
End Sub


'*******************************************************************************
Private Sub mnuWksnUserReleaseDD_Click()
On Error Resume Next

    Call mnuWksnUserRelease_Click
End Sub


'*******************************************************************************
Private Sub optConsoleWorkstationsAll_Click()
Dim uWorkstation As New clsWorkstation
On Error GoTo optConsoleWorkstationsAll_ClickError

    For Each uWorkstation In guWorkstations
        uWorkstation.AllowGroupControl = True
    Next uWorkstation
    lblConsoleWksnControl.Caption = funLoadResString(10634)
    cmdConsoleButtons(8).enabled = False
    Call funSetFocus(tbsFunctions)
    Exit Sub
    
optConsoleWorkstationsAll_ClickError:
    Call funErrorManager("frmMain.optConsoleWorkstationsAll_ClickError", True)
End Sub


'*******************************************************************************
Private Sub optConsoleWorkstationsSome_Click()
On Error GoTo optConsoleWorkstationsSome_ClickError

    cmdConsoleButtons(8).enabled = True
    lblConsoleWksnControl.Caption = funLoadResString(10635)
    Call funSetFocus(tbsFunctions)
    cmdConsoleButtons(8).value = True
    Exit Sub
    
optConsoleWorkstationsSome_ClickError:
    Call funErrorManager("frmMain.optConsoleWorkstationsSome_ClickError", True)
End Sub


'*******************************************************************************
Private Sub optIncomeFilter_Click(Index As Integer)
On Error GoTo optIncomeFilter_ClickError

    dtpStartDate.enabled = Index = 6
    dtpEndDate.enabled = Index = 6
    Exit Sub
    
optIncomeFilter_ClickError:
    Call funErrorManager("frmMain.optIncomeFilter_ClickError", True)
End Sub


'*******************************************************************************
Private Sub optIncomeFilter_DblClick(Index As Integer)
On Error GoTo optIncomeFilter_DblClickError
     
    dtpStartDate.enabled = Index = 6
    dtpEndDate.enabled = Index = 6
    If gblnDBOK Then Call cmdIncomeFilter_Click
    Exit Sub
    
optIncomeFilter_DblClickError:
    Call funErrorManager("frmMain.optIncomeFilter_DblClickError", True)
End Sub


'*******************************************************************************
Private Sub optSessionFilter_Click(Index As Integer)
On Error GoTo optSessionFilter_ClickError

    dtpSessionStartDate.enabled = Index = 6
    dtpSessionEndDate.enabled = Index = 6
    Exit Sub
    
optSessionFilter_ClickError:
    Call funErrorManager("frmMain.optSessionFilter_ClickError", True)
End Sub


'*******************************************************************************
Private Sub optSessionFilter_DblClick(Index As Integer)
On Error GoTo optSessionFilter_DblClickError
     
    dtpSessionStartDate.enabled = Index = 6
    dtpSessionEndDate.enabled = Index = 6
    If gblnDBOK Then Call cmdSessionFilter_Click
    Exit Sub
    
optSessionFilter_DblClickError:
    Call funErrorManager("frmMain.optSessionFilter_DblClickError", True)
End Sub


'*******************************************************************************
Private Sub stbStatus_PanelClick(ByVal Panel As MSComctlLib.Panel)
On Error GoTo stbStatus_PanelClickError

    If Panel.Index = 2 Then
        Panel.Bevel = sbrInset
        DoEvents
        Call funUserLogoff(guServOpts.uConnectedUser)
        Panel.Bevel = sbrRaised
    End If
    Exit Sub
    
stbStatus_PanelClickError:
    Call funErrorManager("frmMain.stbStatus_PanelClickError", True)
End Sub


'*******************************************************************************
Private Sub sysInfo_SettingChanged(ByVal Item As Integer)
Dim frmTemp As Form
On Error GoTo sysInfo_SettingChangedError

    For Each frmTemp In Forms
        Call frmTemp.subLocalizeInterface
    Next frmTemp
    Exit Sub
    
sysInfo_SettingChangedError:
    Call funErrorManager("frmMain.sysInfo_SettingChangedError", True)
End Sub


'*******************************************************************************
Private Sub tbsFunctions_Click()
Dim i As Integer
On Error GoTo tbsFunctions_ClickError
     
    For i = 1 To fraFunctions.Count
        fraFunctions(i).Visible = "k" & i = tbsFunctions.SelectedItem.Key
        mnuFunctions(i).Checked = "k" & i = tbsFunctions.SelectedItem.Key
        mnuSubFunctions(i).Visible = "k" & i = tbsFunctions.SelectedItem.Key
    Next i
    cmdCrdFind.Default = tbsFunctions.SelectedItem.Key = "k2"
    cmdCliFind.Default = tbsFunctions.SelectedItem.Key = "k3"
    cmdSessionURLSearch.Default = tbsFunctions.SelectedItem.Key = "k5"
    Exit Sub
    
tbsFunctions_ClickError:
    Call funErrorManager("frmMain.tbsFunctions_ClickError", True)
End Sub


'*******************************************************************************
Private Sub tbsSessionResultType_Click()
Dim i As Integer
On Error GoTo tbsSessionResultType_ClickError

    For i = 1 To fraSessionResults.Count
        fraSessionResults(i).Visible = i = tbsSessionResultType.SelectedItem.Index
    Next i
    If mblnSessionFiltered Then
        cmdSessionExport.enabled = Not tbsSessionResultType.SelectedItem.Index = 5
        cmdSessionPrint.enabled = Not tbsSessionResultType.SelectedItem.Index = 5 And guServOpts.uConfig.strDefaultPrinterName <> ""
    End If
    Exit Sub
    
tbsSessionResultType_ClickError:
    Call funErrorManager("frmMain.tbsSessionResultType_ClickError", True)
End Sub


Private Sub tmrActivateServer_Timer()
On Error GoTo tmrActivateServer_TimerError

    Call subInitServer
    Call subInitCnxTesting
    tmrActivateServer.enabled = False
    Exit Sub
    
tmrActivateServer_TimerError:
    Call funErrorManager("frmMain.tmrActivateServer_TimerError", True)
End Sub

'*******************************************************************************
Private Sub tmrBackupDBToWorkstation_Timer()
Dim intFileNumber As Integer
Dim intWksnNumber As Integer
Dim strBuffer As String
Dim i As Integer
Dim dteOldestBackup As Date
On Error GoTo tmrBackupDBToWorkstation_TimerError

    If guWorkstations.Count > 0 Then
        If Dir(guServOpts.uDB.strDBPath & ".bak1") <> "" And gblnDBOK Then
            intWksnNumber = 1
            For i = 1 To guWorkstations.Count
                If guWorkstations(i).LastDBBackupDate = NULL_DATE Then
                    intWksnNumber = i
                    Exit For
                Else
                    If dteOldestBackup = NULL_DATE Or dteOldestBackup > guWorkstations(i).LastDBBackupDate Then
                        dteOldestBackup = guWorkstations(i).LastDBBackupDate
                        intWksnNumber = i
                    End If
                End If
            Next i
            Call funSendFile(guWorkstations(intWksnNumber).Cnx, guServOpts.uDB.strDBPath & ".bak1", SaveDB)
            intFileNumber = FreeFile
            Open guServOpts.uDB.strDBPath & FILENAME_SUFFIX_BD_NET_BACKUP For Output As intFileNumber
            strBuffer = funLoadResString(10000) & " - " & funLoadResString(10706) & vbNewLine
            strBuffer = strBuffer & funLoadResString(10707) & "\\" & guWorkstations(intWksnNumber).ComputerName & vbNewLine
            strBuffer = strBuffer & funLoadResString(10708) & format(FileDateTime(guServOpts.uDB.strDBPath & ".bak1"), "General Date")
            Print #intFileNumber, strBuffer
            Close intFileNumber
            Call SetRegSettings(funLoadResString(10006), "DB", "Last workstation location", "\\" & guWorkstations(intWksnNumber).ComputerName)
            Call subDBAddEvent(AppDBNetworkBackup, 0, 0, "", 0, funLoadResString(10706) & " - " & funLoadResString(10707) & "\\" & guWorkstations(intWksnNumber).ComputerName, All)
        End If
        tmrBackupDBToWorkstation.enabled = False
    End If
    Exit Sub
    
tmrBackupDBToWorkstation_TimerError:
    Call funErrorManager("frmMain.tmrBackupDBToWorkstation_TimerError", True)
End Sub


'*******************************************************************************
Private Sub tmrCliRefresh_Timer()
On Error GoTo tmrCliRefresh_TimerError

    If guServOpts.uWorkstations.intActivateClients = vbChecked Then
        Call subFillClients
        Call subUpdateUISelectedClient
    End If
    Exit Sub
    
tmrCliRefresh_TimerError:
    Call funErrorManager("frmMain.tmrCliRefresh_TimerError", True)
End Sub


'*******************************************************************************
Private Sub tmrCrdRefresh_Timer()
On Error GoTo tmrCrdRefresh_TimerError

    If guServOpts.uWorkstations.intActivateCards = vbChecked Then
        Call subFillCards
        Call subUpdateUISelectedCard
    End If
    Exit Sub
    
tmrCrdRefresh_TimerError:
    Call funErrorManager("frmMain.tmrCrdRefresh_TimerError", True)
End Sub


'*******************************************************************************
Private Sub tmrCnxTest_Timer()
On Error GoTo tmrCnxTest_TimerError

    If wskCnxTest.State <> sckConnected Then
        Call wskCnxTest.Close
        Call wskCnxTest.Connect
    End If
    Exit Sub
    
tmrCnxTest_TimerError:
    tmrCnxTest.enabled = False
    Call funErrorManager("frmMain.tmrCnxTest_TimerError", True)
End Sub


'*******************************************************************************
Private Sub tmrScreenCapture_Timer()
Dim uWorkstation As New clsWorkstation
Dim blnAnyWorkstationFollowing As Boolean
Dim dib As Long
Dim bOK As Long
On Error GoTo tmrScreenCapture_TimerError
    
    For Each uWorkstation In guWorkstations
        blnAnyWorkstationFollowing = uWorkstation.State = enuWksnState.Blocked_follow
        If blnAnyWorkstationFollowing Then Exit For
    Next uWorkstation
    
    If blnAnyWorkstationFollowing Then
        Screen.MousePointer = vbHourglass
        
        Call SavePicture(CaptureScreen, guServOpts.uConfig.strAppConfigPath & FILENAME_SCREENCAP_BMP)
        DoEvents
        
        dib = FreeImage_Load(FIF_BMP, guServOpts.uConfig.strAppConfigPath & FILENAME_SCREENCAP_BMP, 0)
        DoEvents
        
'        dib = FreeImage_Rescale(dib, lngTest1, lngTest2, FILTER_BILINEAR)
'        DoEvents
        
        bOK = FreeImage_Save(FIF_JPEG, dib, guServOpts.uConfig.strAppConfigPath & FILENAME_SCREENCAP_JPG, 0)
        DoEvents
        
        ' Unload the dib
        Call FreeImage_Unload(dib)
        DoEvents

        For Each uWorkstation In guWorkstations
            If uWorkstation.State = enuWksnState.Blocked_follow And uWorkstation.ReadyForServerScreen Then
                uWorkstation.ReadyForServerScreen = False
                Call funSendFile(uWorkstation.Cnx, guServOpts.uConfig.strAppConfigPath & FILENAME_SCREENCAP_JPG, LoadFollowImage)
            End If
            DoEvents
        Next uWorkstation
        Screen.MousePointer = vbNormal
    End If
    Exit Sub
    
tmrScreenCapture_TimerError:
    tmrStateUpdade.enabled = False
    Call funErrorManager("frmMain.tmrScreenCapture_TimerError", True)
End Sub


'*******************************************************************************
Private Sub tmrStateUpdade_Timer()
On Error GoTo tmrStateUpdade_TimerError
    
    DoEvents
    Call subUpdateUIWorkstations
    Call subUpdateSessions
    mnuMinimizeToSysTray.enabled = Not (gblnfrmCardHistory Or gblnfrmCardMassAdd Or gblnfrmLogon Or gblnfrmRecharge Or gblnfrmSetupWizard)
    Exit Sub
    
tmrStateUpdade_TimerError:
    tmrStateUpdade.enabled = False
    Call funErrorManager("frmMain.tmrStateUpdade_TimerError", True)
End Sub


'*******************************************************************************
Private Sub tmrStatusBarUpdate_Timer()
Static i As Integer
On Error GoTo tmrStatusBarUpdate_TimerError


'    If Me.Visible Then
'        guIconData.hIcon = imgSystTrayIcons.ListImages(1).Picture
'        guIconData.szTip = funLoadResString(10000) & Chr(0)
'    End If
    If guStatusBarMessages.Count = 0 Then
        stbStatus.Panels(1).Picture = imlSatusBar.ListImages(1).Picture
        stbStatus.Panels(1).Text = ""
'        If Not Me.Visible Then
'            guIconData.hIcon = imgSystTrayIcons.ListImages(1).Picture
'            guIconData.szTip = funLoadResString(10000) & Chr(0)
'        End If
    Else
        i = i + 1
        If i > guStatusBarMessages.Count Then i = 1
        stbStatus.Panels(1).Picture = imlSatusBar.ListImages(guStatusBarMessages(i).ImageIndex).Picture
        stbStatus.Panels(1).Text = guStatusBarMessages(i).Text & "."
'        If Not Me.Visible Then
'            guIconData.hIcon = imgSystTrayIcons.ListImages(guStatusBarMessages(i).ImageIndex).Picture
'            guIconData.szTip = funLoadResString(10000) & " - " & guStatusBarMessages(i).Text & Chr(0)
'        End If
    End If
'    Call Shell_NotifyIcon(NIM_MODIFY, guIconData)
    Exit Sub
    
tmrStatusBarUpdate_TimerError:
    tmrStatusBarUpdate.enabled = False
    Call funErrorManager("frmMain.tmrStatusBarUpdate_TimerError", True)
End Sub


'*******************************************************************************
Private Sub tmrSystTrayLogoff_Timer()
Static intCounter As Integer
On Error GoTo tmrSystTrayLogoff_TimerError

'    If intCounter > 2 Then intCounter = 0
    intCounter = intCounter + 1
    
    If intCounter = HIDDEN_LOGOFF_TIMEOUT Then
        tmrSystTrayLogoff.enabled = False
        mblnLogoffOnShow = True
        intCounter = 0
    End If
    Exit Sub
    
tmrSystTrayLogoff_TimerError:
    tmrSystTrayLogoff.enabled = False
    Call funErrorManager("frmMain.tmrSystTrayLogoff_TimerError", True)
End Sub


'*******************************************************************************
Private Sub tmrSystTrayUpdate_Timer()
Static i As Integer
Static blnFlip As Boolean
On Error GoTo tmrSystTrayUpdate_TimerError


    If Me.Visible Or Not (blnFlip Or Me.Visible) Then
        guIconData.hIcon = imgSystTrayIcons.ListImages(1).Picture
        If Me.Visible Then guIconData.szTip = funLoadResString(10000) & Chr(0)
    ElseIf blnFlip Then
        If guStatusBarMessages.Count = 0 Then
            guIconData.hIcon = imgSystTrayIcons.ListImages(1).Picture
            guIconData.szTip = funLoadResString(10000) & Chr(0)
        Else
            i = i + 1
            If i > guStatusBarMessages.Count Then i = 1
            guIconData.hIcon = imgSystTrayIcons.ListImages(guStatusBarMessages(i).ImageIndex).Picture
            guIconData.szTip = funLoadResString(10000) & " - " & guStatusBarMessages(i).Text & Chr(0)
        End If
    End If
    blnFlip = Not blnFlip
    Call Shell_NotifyIcon(NIM_MODIFY, guIconData)
    Exit Sub
    
tmrSystTrayUpdate_TimerError:
    tmrSystTrayUpdate.enabled = False
    Call funErrorManager("frmMain.tmrSystTrayUpdate_TimerError", True)
End Sub


'*******************************************************************************
Private Sub txtCliFind_GotFocus()
On Error Resume Next

    txtCliFind.SelStart = 0
    txtCliFind.SelLength = Len(txtCliFind.Text)
End Sub


'*******************************************************************************
Private Sub txtCrdFindNum_GotFocus()
On Error GoTo txtCrdFindNum_GotFocusError
    
    txtCrdFindNum.SelStart = 0
    txtCrdFindNum.SelLength = Len(txtCrdFindNum.Text)
    cmdCrdFind.Default = True
    Exit Sub
    
txtCrdFindNum_GotFocusError:
    Call funErrorManager("frmMain.txtCrdFindNum_GotFocusError", True)
End Sub


'*******************************************************************************
Private Sub txtCrdComment_Change()
On Error GoTo txtCrdComment_ChangeError
    
    tmrCrdRefresh.enabled = False
    Exit Sub
    
txtCrdComment_ChangeError:
    Call funErrorManager("frmMain.txtCrdComment_ChangeError", True)
End Sub


'*******************************************************************************
Private Sub txtCrdComment_GotFocus()
On Error GoTo txtCrdComment_GotFocusError

    tmrCrdRefresh.enabled = False
    Exit Sub
    
txtCrdComment_GotFocusError:
    Call funErrorManager("frmMain.txtCrdComment_GotFocusError", True)
End Sub


'*******************************************************************************
Private Sub txtCrdPending_Change()
Dim i As Long
Dim intCurrentIndex As Integer
Static blnAnalizing As Boolean
Dim blnFound As Boolean
On Error GoTo txtCrdPending_ChangeError

    If Not mblntxtCrdPendingHasFocus Or blnAnalizing Or Len(txtCrdPending.Text) = 0 Then
        cmdCrdPendingEnable.enabled = Len(txtCrdPending.Text) <> 0
        Exit Sub
    End If
    For i = 1 To fxgCrdPending.Rows - 1
        If InStr(1, fxgCrdPending.TextMatrix(i, 0), txtCrdPending.Text) = 1 Then
            blnAnalizing = True
            blnFound = True
            fxgCrdPending.TopRow = IIf(i = 1, i, i - 1)
            fxgCrdPending.Row = i
            Call fxgCrdPending_RowColChange
            intCurrentIndex = Len(txtCrdPending.Text)
            txtCrdPending.Text = fxgCrdPending.TextMatrix(i, 0)
            txtCrdPending.SelStart = intCurrentIndex
            txtCrdPending.SelLength = Len(txtCrdPending.Text) - intCurrentIndex
            blnAnalizing = False
           Exit For
        End If
    Next i
    cmdCrdPendingEnable.enabled = blnFound
    If Not blnFound Then
        fxgCrdPending.Row = guServOpts.lngSelectedCardPendingIndex
        fxgCrdPending.ColSel = fxgCrdPending.Cols - 1
        fxgCrdPending.CellForeColor = vbWindowText
        fxgCrdPending.CellBackColor = vbWindowBackground
        guServOpts.lngSelectedCardPendingIndex = 0
    End If
    Exit Sub
    
txtCrdPending_ChangeError:
    Call funErrorManager("frmMain.txtCrdPending_ChangeError", True)
End Sub


'*******************************************************************************
Private Sub txtCrdPending_GotFocus()
On Error Resume Next

    cmdCrdPendingEnable.Default = True
    mblntxtCrdPendingHasFocus = True
End Sub


'*******************************************************************************
Private Sub txtCrdPending_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    
    Select Case KeyCode
    Case 40
        If fxgCrdPending.Row < fxgCrdPending.Rows - 1 Then
            fxgCrdPending.TopRow = fxgCrdPending.TopRow + 1
            fxgCrdPending.Row = fxgCrdPending.Row + 1
            Call fxgCrdPending_RowColChange
        End If
    Case 38
        If fxgCrdPending.Row > 1 Then
            fxgCrdPending.TopRow = fxgCrdPending.TopRow - 1
            fxgCrdPending.Row = fxgCrdPending.Row - 1
            Call fxgCrdPending_RowColChange
        End If
        
    Case 33
        If fxgCrdPending.Row > Int(fxgCrdPending.height / fxgCrdPending.RowHeight(0) - 2) Then
            fxgCrdPending.TopRow = fxgCrdPending.TopRow - Int(fxgCrdPending.height / fxgCrdPending.RowHeight(0) - 2)
            fxgCrdPending.Row = fxgCrdPending.Row - Int(fxgCrdPending.height / fxgCrdPending.RowHeight(0) - 2)
            Call fxgCrdPending_RowColChange
        Else
            fxgCrdPending.TopRow = 1
            fxgCrdPending.Row = 1
            Call fxgCrdPending_RowColChange
        End If
    Case 34
        If fxgCrdPending.Row < fxgCrdPending.Rows - Int(fxgCrdPending.height / fxgCrdPending.RowHeight(0) - 2) Then
            fxgCrdPending.TopRow = fxgCrdPending.TopRow + Int(fxgCrdPending.height / fxgCrdPending.RowHeight(0) - 2)
            fxgCrdPending.Row = fxgCrdPending.Row + Int(fxgCrdPending.height / fxgCrdPending.RowHeight(0) - 2)
            Call fxgCrdPending_RowColChange
        Else
            fxgCrdPending.TopRow = fxgCrdPending.Rows - 1
            fxgCrdPending.Row = fxgCrdPending.Rows - 1
            Call fxgCrdPending_RowColChange
        End If
    
    
    
    End Select
End Sub


'*******************************************************************************
Private Sub txtCrdPending_KeyPress(KeyAscii As Integer)
Dim i As Long
Dim intCurrentIndex As Integer
On Error GoTo txtCrdPending_KeyPressError

    If ((KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8) Then
        KeyAscii = 0
        Exit Sub
    End If
    Exit Sub
    
txtCrdPending_KeyPressError:
    Call funErrorManager("frmMain.txtCrdPending_KeyPressError", True)
End Sub


Private Sub txtCrdPending_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    If (KeyCode = 38 Or KeyCode = 40) And guServOpts.lngSelectedCardPendingIndex > 0 Then
        txtCrdPending.Text = fxgCrdPending.TextMatrix(guServOpts.lngSelectedCardPendingIndex, 0)
        txtCrdPending.SelStart = 0
        txtCrdPending.SelLength = Len(txtCrdPending.Text)
        cmdCrdPendingEnable.enabled = True
        cmdCrdPendingEnable.Default = True
    End If
End Sub

'*******************************************************************************
Private Sub txtCrdPending_LostFocus()
On Error Resume Next

    cmdCrdFind.Default = True
    mblntxtCrdPendingHasFocus = False
End Sub

'*******************************************************************************
Private Sub txtServicesControlsQuantity_Change()
Dim strQuantity As String
Dim strChar As String
Dim i As Integer
On Error GoTo txtServicesControlsQuantity_ChangeError
    
    For i = 1 To Len(txtServicesControlsQuantity.Text)
        strChar = Mid(txtServicesControlsQuantity.Text, i, 1)
        If IsNumeric(strChar) Then strQuantity = strQuantity + strChar
    Next i
    strQuantity = IIf(strQuantity = "", 0, strQuantity)
    txtServicesControlsQuantity.Text = strQuantity
    txtServicesControlsQuantity.SelStart = Len(txtServicesControlsQuantity.Text)
    
    If guServOpts.lngSelectedServiceIndex <> 0 Then
        lblServSel(3).Caption = " " & funFormatCurrencyString(CLng(strQuantity) * guServices(guServOpts.lngSelectedServiceIndex).Price, guServOpts.uConfig.strMonFormat, guServOpts.uConfig.strMonSymbol)
        cmdServicesCollect.enabled = strQuantity <> 0
        mnuServicesCollect.enabled = strQuantity <> 0
    End If
    Exit Sub

txtServicesControlsQuantity_ChangeError:
    Call funErrorManager("frmMain.txtServicesControlsQuantity_ChangeError", True)
End Sub


'*******************************************************************************
Private Sub txtSessionURLSearch_GotFocus()
On Error GoTo txtSessionURLSearch_GotFocusError
    
    txtSessionURLSearch.SelStart = 0
    txtSessionURLSearch.SelLength = Len(txtSessionURLSearch.Text)
    Exit Sub
    
txtSessionURLSearch_GotFocusError:
    Call funErrorManager("frmMain.txtSessionURLSearch_GotFocusError", True)
End Sub


'*******************************************************************************
Private Sub txtWksnChat_GotFocus()
On Error Resume Next

    cmdWksnButtons(13).Default = True
End Sub


'*******************************************************************************
Private Sub ucWorkstations_Click(Index As Integer)
Dim i As Integer, intRowIndex As Integer
Dim uEmptyWorkstation As New clsWorkstation
Dim intNumWorkstationRows As Integer
On Error GoTo ucWorkstations_ClickError

    intNumWorkstationRows = Round((frmMain.fraFunctions(1).height - frmMain.fraWksnContainer.top) / (frmMain.ucWorkstations(1).height + 60))
    If guServOpts.uFlags.intWksnSize = enuWksnSize.WksnSizeNormal Then
        intRowIndex = Index
    Else
        intRowIndex = IIf(CSng(Index / 2) > Int(Index / 2), Int(Index / 2) + 1, Int(Index / 2))
    End If
    For i = 1 To ucWorkstations.Count
        ucWorkstations(i).Active = i = Index
    Next i
    shpWksnConsoleBackground.BackColor = vb3DFace
    frmMain.fraWksnControls.Visible = True
    frmMain.fraWksnConsoleControls.Visible = False
    
    If intRowIndex > intNumWorkstationRows + fsbWksn.value Then
        fsbWksn.value = intRowIndex - intNumWorkstationRows
    ElseIf intRowIndex <= fsbWksn.value Then
        fsbWksn.value = intRowIndex - 1
    End If
    If Index <= guMacAddresses.Count Then
        Call subUpdateUISelectedWorkstation(funFindWorkstationByMac(guMacAddresses(Index)))
    Else
        Call subUpdateUISelectedWorkstation(uEmptyWorkstation)
    End If
    Exit Sub
    
ucWorkstations_ClickError:
    Call funErrorManager("frmMain.ucWorkstations_ClickError", True)
End Sub


'*******************************************************************************
Private Sub ucWorkstations_RightClick(Index As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer

    If Index <= guMacAddresses.Count Then
        If Not ucWorkstations(Index).Active Then Call ucWorkstations_Click(Index)
        Call PopupMenu(mnuWksnDropDown, vbPopupMenuLeftButton, , y + fraFunctions(1).top + fraWksnContainer.top + ucWorkstations(Index).top - fraWksn.top)
    End If
End Sub


'*******************************************************************************
Private Sub udServicesControlsQuantity_Change()
On Error GoTo udServicesControlsQuantity_ChangeError
On Error Resume Next

    cmdServicesCollect.enabled = udServicesControlsQuantity.value <> 0
    mnuServicesCollect.enabled = udServicesControlsQuantity.value <> 0
    txtServicesControlsQuantity.SelStart = 0
    txtServicesControlsQuantity.SelLength = Len(txtServicesControlsQuantity.Text)
    Call funSetFocus(txtServicesControlsQuantity)
    Exit Sub
    
udServicesControlsQuantity_ChangeError:
    Call funErrorManager("frmMain.udServicesControlsQuantity_ChangeError", True)
End Sub


'*******************************************************************************
Private Sub wskCnxTest_Close()
On Error GoTo wskCnxTest_CloseError

    tmrCnxTest.enabled = True
    Exit Sub
    
wskCnxTest_CloseError:
    Call funErrorManager("frmMain.wskCnxTest_CloseError", True)
End Sub


'*******************************************************************************
Private Sub wskCnxTest_Connect()
Dim strBuffer As String
Dim uWorkstation As New clsWorkstation
On Error GoTo wskCnxTest_ConnectError

    tmrCnxTest.enabled = False
    
    If Not guServOpts.uFlags.blnInetCnx Then
        guServOpts.uFlags.blnInetCnx = True
        Call subRemoveStatusBarMsg("NO_CNX")
        Call subDBAddEvent(InternetConnected, 0, 0, "", 0, App.Title & funLoadResString(10377))
        If guServOpts.uCnxTest.intWorkstationNotification = vbChecked Then
            strBuffer = CNX_NOTIFY_WKSN_INETCNX & CNX_MSG_END
            Call subBroadcastData(strBuffer)
        End If
    End If
    Exit Sub
    
wskCnxTest_ConnectError:
    Call funErrorManager("frmMain.wskCnxTest_ConnectError", True)
End Sub


'*******************************************************************************
Private Sub wskCnxTest_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim uMsg As New clsStatusBarMessage
Dim strBuffer As String
Dim uWorkstation As New clsWorkstation
On Error GoTo wskCnxTest_ErrorError

    tmrCnxTest.enabled = True
    If Number <> sckConnectAborted And Number <> sckConnectionReset Then
        guServOpts.uFlags.blnInetCnx = False
        uMsg.ImageIndex = enuStatus.cliWarn
        uMsg.Text = App.Title & funLoadResString(10207)
        If funAddStatusBarMsg(uMsg, "NO_CNX") Then Call subDBAddEvent(InternetDiconnected, 0, 0, "", 0, App.Title & funLoadResString(10207))
        If guServOpts.uCnxTest.intWorkstationNotification = vbChecked Then
            strBuffer = CNX_NOTIFY_WKSN_NO_INETCNX & CNX_MSG_END
            Call subBroadcastData(strBuffer)
        End If
    End If
    Exit Sub
    
wskCnxTest_ErrorError:
    Call funErrorManager("frmMain.wskCnxTest_ErrorError", True)
End Sub


'*******************************************************************************
Private Sub wskFindWorkstation_DataArrival(ByVal bytesTotal As Long)
Dim strNewData As String
On Error GoTo wskFindWorkstation_DataArrivalError

    Call wskFindWorkstation.GetData(strNewData)
    If strNewData = CNX_AUTO_CONFIG_HANDSHAKE_WKSN Then
        Call wskHandshakeWorkstation.Close
        wskHandshakeWorkstation.Protocol = sckUDPProtocol
        wskHandshakeWorkstation.RemoteHost = wskFindWorkstation.RemoteHostIP
        wskHandshakeWorkstation.RemotePort = DEFAULT_UDP_PORT_WORKSTATION
        Call frmMain.wskHandshakeWorkstation.SendData(CNX_AUTO_CONFIG_HANDSHAKE_SERV & guServOpts.uServer.strServerName)
    End If
    Exit Sub
    
wskFindWorkstation_DataArrivalError:
'    Call funErrorManager("frmMain.wskFindWorkstation_DataArrivalError RemoteHostIP: " & wskFindWorkstation.RemoteHostIP, True)
End Sub


'*******************************************************************************
Private Sub wskServ_Close(Index As Integer)
Dim uWorkstation As New clsWorkstation
On Error GoTo wskServ_CloseError

    Set uWorkstation = guWorkstations("Key: " & Index)
    Call subSessionClose(uWorkstation)
    Call subDBAddEvent(WorkstationNotPresent, 0, 0, "", 0, funLoadResString(10562) & uWorkstation.WorkstationName, All)
'    Call subWriteToLogFile("frmMain.wskServ_Close: " & funLoadResString(10562) & uWorkstation.WorkstationName)
    Call wskServ(Index).Close
    Call Unload(wskServ(Index))
    If uWorkstation.ComputerMacAddress <> "" Then guMacAddresses(uWorkstation.ComputerMacAddress).CurrentlyConnected = False
    Call guWorkstations.Remove("Key: " & Index)
    Exit Sub
    
wskServ_CloseError:
    Call funErrorManager("frmMain.wskServ_CloseError", True)
End Sub


'*******************************************************************************
Private Sub wskServ_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim uNewWorkstation As New clsWorkstation
Static lngLastrequestID As Long
Static lngNumConcurrentCnx As Long
On Error GoTo wskServ_ConnectionRequestError

        
    If requestID = lngLastrequestID Then
        Exit Sub
    End If
    If Index = 0 Then
        If guWorkstations.Count < guServOpts.uWorkstations.lngNumWorkstations Then
            lngNumConcurrentCnx = lngNumConcurrentCnx + 1
            gblnWorkstationConnexionInProgress = True
            Call subDebugPrint("wskServ_ConnectionRequest requestID:" & requestID)
            Call Load(wskServ(wskServ.UBound + 1))
            wskServ(wskServ.UBound).LocalPort = 0
            Call wskServ(wskServ.UBound).Accept(requestID)
            uNewWorkstation.Key = "Key: " & wskServ.UBound
            Set uNewWorkstation.Cnx = wskServ(wskServ.UBound)
            Set uNewWorkstation.CurrentCard = New clsCard
            Set uNewWorkstation.CurrentClient = New clsClient
            Set uNewWorkstation.SelfServicePrePaid = New clsClient
            lngLastrequestID = requestID
            Call subWorkstationAdd(uNewWorkstation)
            lngNumConcurrentCnx = lngNumConcurrentCnx - 1
            gblnWorkstationConnexionInProgress = lngNumConcurrentCnx <> 0
        End If
    End If
    Exit Sub
    
wskServ_ConnectionRequestError:
    Call funErrorManager("frmMain.wskServ_ConnectionRequestError", False, True)
    Call wskServ(Index).Close
    Call Unload(wskServ(Index))
    Call guWorkstations.Remove("Key: " & Index)
    lngNumConcurrentCnx = lngNumConcurrentCnx - 1
    gblnWorkstationConnexionInProgress = lngNumConcurrentCnx <> 0
End Sub


'*******************************************************************************
Private Sub wskServ_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strNewData As String
Dim strDatagram As String
Dim lngEndIndex As Long
Dim uWorkstation As clsWorkstation
On Error GoTo wskServ_DataArrivalError

    Call wskServ(Index).GetData(strNewData, vbString)
    Set uWorkstation = guWorkstations("Key: " & Index)
    uWorkstation.CnxIncomingBuffer = uWorkstation.CnxIncomingBuffer & strNewData
    Call subWorkstationProcessData(uWorkstation)
    Exit Sub
    
wskServ_DataArrivalError:
    Call funErrorManager("frmMain.wskServ_DataArrivalError", True)
End Sub


'*******************************************************************************
Private Sub wskServ_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error GoTo wskServ_ErrorError
    
'    If Index <> 0 And Number <> sckConnectionReset And Number <> sckConnectAborted Then Call wskServ_Close(Index)
    If Index <> 0 Then
        Call wskServ_Close(Index)
    End If
    Exit Sub
    
wskServ_ErrorError:
    Call funErrorManager("frmMain.wskServ_ErrorError", True)
End Sub


'*******************************************************************************
Public Sub subLocalizeInterface()
Dim i As Integer
Dim intNumIntervals As Integer
Dim strTmpLabel As String
On Error GoTo subLocalizeInterfaceError

    Call subSetCharset(Me, guServOpts.uConfig.intInterfaceLangCharset)
    
    Me.Caption = funLoadResString(10000) & IIf(BETA_VERSION, " - " & App.Major & "." & App.Minor & "." & App.Revision & " BETA", "") & " - " & guServOpts.uServer.strServerName
    On Error Resume Next
    tbsFunctions.Tabs("k1").Caption = funLoadResString(10149)
    tbsFunctions.Tabs("k1").ToolTipText = funLoadResString(10444)
    If guServOpts.uWorkstations.intActivateCards = vbChecked Then
        tbsFunctions.Tabs("k2").Caption = funLoadResString(10150)
        tbsFunctions.Tabs("k2").ToolTipText = funLoadResString(10445)
    End If
    If guServOpts.uWorkstations.intActivateClients = vbChecked Then
        tbsFunctions.Tabs("k3").Caption = funLoadResString(10796)
        tbsFunctions.Tabs("k3").ToolTipText = funLoadResString(10797)
    End If
    tbsFunctions.Tabs("k4").Caption = funLoadResString(10324)
    tbsFunctions.Tabs("k4").ToolTipText = funLoadResString(10446)
    tbsFunctions.Tabs("k5").Caption = funLoadResString(10151)
    tbsFunctions.Tabs("k5").ToolTipText = funLoadResString(10447)
    tbsFunctions.Tabs("k6").Caption = funLoadResString(10152)
    tbsFunctions.Tabs("k6").ToolTipText = funLoadResString(10448)
    On Error GoTo subLocalizeInterfaceError
    
    tbsSessionResultType.Tabs(1).Caption = funLoadResString(10401)
    tbsSessionResultType.Tabs(1).ToolTipText = funLoadResString(10449)
    tbsSessionResultType.Tabs(2).Caption = funLoadResString(10421)
    tbsSessionResultType.Tabs(2).ToolTipText = funLoadResString(10450)
    tbsSessionResultType.Tabs(3).Caption = funLoadResString(10402)
    tbsSessionResultType.Tabs(3).ToolTipText = funLoadResString(10451)
    tbsSessionResultType.Tabs(4).Caption = funLoadResString(10698)
    tbsSessionResultType.Tabs(4).ToolTipText = funLoadResString(10699)
    tbsSessionResultType.Tabs(5).Caption = funLoadResString(10422)
    tbsSessionResultType.Tabs(5).ToolTipText = funLoadResString(10452)
    
    stbStatus.Panels(1).ToolTipText = funLoadResString(10508)
    
    'Localize Menus
    mnuAbout.Caption = funLoadResString(10153) & funLoadResString(10000) & "..."
    
    mnuCrd(0).Caption = funLoadResString(10154) & "..."
    mnuCrd(1).Caption = funLoadResString(10141) & "..."
    mnuCrdDD(1).Caption = funLoadResString(10141) & "..."
    mnuCrd(2).Caption = funLoadResString(10530)
    mnuCrd(5).Caption = funLoadResString(10532) & "..."
    mnuCrdDD(5).Caption = funLoadResString(10532) & "..."
    mnuCrd(6).Caption = funLoadResString(10595)
    mnuCrdDD(6).Caption = funLoadResString(10595)
    mnuCrd(3).Caption = funLoadResString(10688)
    mnuCrdDD(3).Caption = funLoadResString(10688)
    mnuCrd(4).Caption = funLoadResString(10690)
    mnuCrdDD(4).Caption = funLoadResString(10690)
    
    mnuCrd(9).Caption = funLoadResString(10291)
    mnuCrdDD(9).Caption = funLoadResString(10291)
    mnuCrd(10).Caption = funLoadResString(10918) & "..."
    mnuCrd(11).Caption = funLoadResString(10183) & "..."
    
    mnuChangePasswd.Caption = funLoadResString(10125) & "..."
    mnuChangePasswdUser.Caption = funLoadResString(10784) & "..."
    mnuUserAdmin.Caption = funLoadResString(10752) & "..."
    
    mnuCli(0).Caption = funLoadResString(10803) & "..."
    mnuCli(1).Caption = funLoadResString(10141) & "..."
    mnuCliDD(1).Caption = funLoadResString(10141) & "..."
    mnuCli(10).Caption = funLoadResString(10846) & "..."
    mnuCli(2).Caption = funLoadResString(10848) & "..."
    mnuCliDD(2).Caption = funLoadResString(10848) & "..."
    mnuCli(7).Caption = funLoadResString(10873)
    mnuCliDD(7).Caption = funLoadResString(10873)
    mnuCli(5).Caption = funLoadResString(10844) & "..."
    mnuCliDD(5).Caption = funLoadResString(10844) & "..."
    mnuCli(9).Caption = funLoadResString(10845)
    mnuCliDD(9).Caption = funLoadResString(10845)
    mnuCli(6).Caption = funLoadResString(10595)
    mnuCliDD(6).Caption = funLoadResString(10595)
    mnuCli(3).Caption = funLoadResString(10688)
    mnuCliDD(3).Caption = funLoadResString(10688)
    mnuCli(4).Caption = funLoadResString(10690)
    mnuCliDD(4).Caption = funLoadResString(10690)
    
    mnuCliModifyModifyDD.Caption = funLoadResString(10848) & "..."
    mnuCliModifyDeleteDD.Caption = funLoadResString(10873)
    
    mnuWksnActivate.Caption = funLoadResString(10155)
    mnuWksnActivateDD.Caption = funLoadResString(10155)
    mnuWksnBloc.Caption = funLoadResString(10156)
    mnuWksnBlocDD.Caption = funLoadResString(10156)
    mnuWksnUnblock.Caption = funLoadResString(10614)
    mnuWksnUnblockDD.Caption = funLoadResString(10614)
    mnuWksnCrdRecharge.Caption = funLoadResString(10791) & "..."
    mnuWksnCrdRechargeDD.Caption = funLoadResString(10791) & "..."
    mnuWksnCrdHistory.Caption = funLoadResString(10793) & "..."
    mnuWksnCrdHistoryDD.Caption = funLoadResString(10793) & "..."
    mnuWksnCrdDisable.Caption = funLoadResString(10792)
    mnuWksnCrdDisableDD.Caption = funLoadResString(10792)
    
    mnuWksnClose.Caption = funLoadResString(10157)
    mnuWksnCloseDD.Caption = funLoadResString(10157)
    mnuWksnConnect.Caption = funLoadResString(10158)
    mnuWksnPause.Caption = funLoadResString(10159)
    mnuWksnPauseDD.Caption = funLoadResString(10159)
    mnuWksnLock.Caption = funLoadResString(10597)
    mnuWksnLockDD.Caption = funLoadResString(10597)
    mnuWksnFollow.Caption = funLoadResString(10623)
    mnuWksnFollowDD.Caption = funLoadResString(10623)
    mnuWksnPauseOne.Caption = funLoadResString(10159)
    mnuWksnLockOne.Caption = funLoadResString(10597)
    mnuWksnFollowOne.Caption = funLoadResString(10623)
    mnuWksnPaySession.Caption = funLoadResString(10160)
    mnuWksnPaySessionDD.Caption = funLoadResString(10160)
    mnuWksnCancelPaySession.Caption = funLoadResString(11023)
    mnuWksnPlay.Caption = funLoadResString(10161)
    mnuWksnPlayDD.Caption = funLoadResString(10161)
    mnuWksnShutdown.Caption = funLoadResString(10162) & "..."
    mnuWksnShutdownDD.Caption = funLoadResString(10162) & "..."
    mnuWksnShowConsole.Caption = funLoadResString(10611)
    mnuWksnShowPrev.Caption = funLoadResString(10612)
    mnuWksnShowNext.Caption = funLoadResString(10613)
    mnuWksnActivateConsole.Caption = funLoadResString(10624)
    mnuWksnBoot.Caption = funLoadResString(10794) & "..."
    mnuWksnBootDD.Caption = funLoadResString(10794) & "..."
    mnuWksnSetSize.Caption = funLoadResString(10899)
    mnuWksnSize(1).Caption = funLoadResString(10900)
    mnuWksnSize(2).Caption = funLoadResString(10901)
    mnuWksnUserRelease.Caption = funLoadResString(10974)
    mnuWksnUserReleaseDD.Caption = funLoadResString(10974)
    mnuWksnStartSelfServPrePaid.Caption = funLoadResString(10980) & "..."
    mnuWksnStartSelfServPrePaidDD.Caption = funLoadResString(10980) & "..."
    
    mnuFile.Caption = funLoadResString(10163)
    If Not guServOpts.uConnectedUser Is Nothing Then mnuLogoff.Caption = funLoadResString(10778) & guServOpts.uConnectedUser.login & "..."
    mnuFunctions(1).Caption = funLoadResString(10149)
    mnuFunctions(2).Caption = funLoadResString(10150)
    mnuFunctions(3).Caption = funLoadResString(10796)
    mnuFunctions(4).Caption = funLoadResString(10324)
    mnuFunctions(5).Caption = funLoadResString(10151)
    mnuFunctions(6).Caption = funLoadResString(10152)
    mnuHelp.Caption = funLoadResString(10164)
    mnuMode.Caption = funLoadResString(10165)
    mnuOptions.Caption = funLoadResString(10372)
    mnuQuit.Caption = funLoadResString(10130) & " " & funLoadResString(10000)
    mnuTools.Caption = funLoadResString(10166)
    mnuSubFunctions(1).Caption = funLoadResString(10167)
    mnuSubFunctions(2).Caption = funLoadResString(10168)
    mnuSubFunctions(3).Caption = funLoadResString(10796)
    mnuSubFunctions(4).Caption = funLoadResString(10325)
    mnuSubFunctions(5).Caption = funLoadResString(10169)
    mnuSubFunctions(6).Caption = funLoadResString(10170)
    mnuIncomeApply.Caption = funLoadResString(10318)
    mnuIncomeExport.Caption = funLoadResString(10320)
    mnuIncomePrint.Caption = funLoadResString(10965)
    mnuServicesCollect.Caption = funLoadResString(10160)
    mnuSessionApply.Caption = funLoadResString(10318)
    mnuSessionExport.Caption = funLoadResString(10426)
    mnuSessionPrint.Caption = funLoadResString(10965)
    mnuDBMaintenance.Caption = funLoadResString(10514)
    mnuCrdStartSessionFirstCilent.Caption = funLoadResString(10592)
    
    mnuSessionURLOpen.Caption = funLoadResString(10716) & "..."
    mnuSessionURLCopy.Caption = funLoadResString(10717)
    mnuSessionCardNumCopy.Caption = funLoadResString(10730)
    
    mnuMinimizeToSysTray.Caption = funLoadResString(10919)
    mnuSysTrayHide.Caption = funLoadResString(10919)
    mnuSysTrayShow.Caption = funLoadResString(10920)
    
    lblServicesTitle.Caption = funLoadResString(10331)
    lblConsoleName.Caption = funLoadResString(10610)
    lblConsoleWksnControl.Caption = funLoadResString(10634)

    lbl(0).Caption = funLoadResString(10184)
    lbl(1).Caption = funLoadResString(10185)
    lbl(2).Caption = funLoadResString(10186)
    lbl(3).Caption = funLoadResString(10655)
    lbl(4).Caption = funLoadResString(10188)
    lbl(5).Caption = funLoadResString(10187)
    lbl(6).Caption = funLoadResString(10189)
    lbl(7).Caption = funLoadResString(10133)
    lbl(8).Caption = funLoadResString(10184)
    lbl(9).Caption = funLoadResString(10134)
    lbl(10).Caption = funLoadResString(10135)
    lbl(11).Caption = funLoadResString(10184)
    lbl(12).Caption = funLoadResString(10333)
    lbl(13).Caption = funLoadResString(10436)
    lbl(14).Caption = funLoadResString(10298)
    lbl(15).Caption = funLoadResString(10299)
    lbl(16).Caption = funLoadResString(10300)
    lbl(17).Caption = funLoadResString(10301)
    lbl(18).Caption = funLoadResString(10506)
    lbl(19).Caption = funLoadResString(10527)
    lbl(20).Caption = funLoadResString(10528)
    lbl(21).Caption = funLoadResString(10334)
    lbl(22).Caption = funLoadResString(10335)
    lbl(23).Caption = funLoadResString(10336)
    lbl(24).Caption = funLoadResString(10390)
    lbl(25).Caption = funLoadResString(10391)
    lbl(26).Caption = funLoadResString(10392)
    lbl(27).Caption = funLoadResString(10393)
    lbl(28).Caption = funLoadResString(10527)
    lbl(29).Caption = funLoadResString(10528)
    lbl(30).Caption = funLoadResString(10638)
    lbl(31).Caption = funLoadResString(10636)
    lbl(32).Caption = funLoadResString(10694)
    lbl(33).Caption = funLoadResString(10187)
    lbl(34).Caption = funLoadResString(10506)
    lbl(35).Caption = funLoadResString(10187)
    lbl(36).Caption = funLoadResString(10801)
    lbl(37).Caption = funLoadResString(10637)
    lbl(38).Caption = funLoadResString(10802)
    lbl(39).Caption = funLoadResString(10184)
    lbl(40).Caption = funLoadResString(10135)
    lbl(41).Caption = funLoadResString(10134)
    lbl(42).Caption = funLoadResString(10133)
    lbl(45).Caption = funLoadResString(10802)
    lbl(46).Caption = funLoadResString(10801)
    lbl(47).Caption = funLoadResString(10188)
    lbl(48).Caption = funLoadResString(10833)
    lbl(51).Caption = funLoadResString(10956)
    lbl(52).Caption = funLoadResString(10921)
    lbl(53).Caption = " " & UCase(left(funLoadResString(10944), 1)) & right(funLoadResString(10944), Len(funLoadResString(10944)) - 1)
    lbl(55).Caption = funLoadResString(10957)
    lbl(56).Caption = funLoadResString(10188)
    
    optSessionFilter(0).Caption = funLoadResString(10394)
    optSessionFilter(1).Caption = funLoadResString(10395)
    optSessionFilter(2).Caption = funLoadResString(10396)
    optSessionFilter(3).Caption = funLoadResString(10397)
    optSessionFilter(4).Caption = funLoadResString(10398)
    'optSessionFilter(5).Caption = funLoadResString(10399)
    optSessionFilter(6).Caption = funLoadResString(10400)
    optSessionFilter(7).Caption = funLoadResString(10411)
    optSessionFilter(8).Caption = funLoadResString(10412)
    
    optIncomeFilter(0).Caption = funLoadResString(10394)
    optIncomeFilter(1).Caption = funLoadResString(10395)
    optIncomeFilter(2).Caption = funLoadResString(10396)
    optIncomeFilter(3).Caption = funLoadResString(10397)
    optIncomeFilter(4).Caption = funLoadResString(10398)
    'optIncomeFilter(5).Caption = funLoadResString(10399)
    optIncomeFilter(6).Caption = funLoadResString(10400)
    optIncomeFilter(7).Caption = funLoadResString(10411)
    optIncomeFilter(8).Caption = funLoadResString(10412)
    
    cmdCrdDelete.Caption = funLoadResString(10190)
    cmdCrdButtons(0).Caption = funLoadResString(10154) & "..."
    cmdCrdButtons(0).ToolTipText = funLoadResString(10457)
    cmdCrdButtons(1).Caption = funLoadResString(10141) & "..."
    cmdCrdButtons(1).ToolTipText = funLoadResString(10458)
    cmdCrdFind.ToolTipText = funLoadResString(10459)
    cmdCrdPendingEnable.ToolTipText = funLoadResString(10955)
    cmdCrdButtons(2).Caption = funLoadResString(10530)
    cmdCrdButtons(2).ToolTipText = funLoadResString(10531)
    txtCrdFindNum.ToolTipText = funLoadResString(10460)
    txtCrdPending.ToolTipText = funLoadResString(10954)
    cmdCrdButtons(5).Caption = funLoadResString(10192) & "..."
    cmdCrdButtons(5).ToolTipText = funLoadResString(10542)
    cmdCrdButtons(6).Caption = funLoadResString(10590)
    cmdCrdButtons(6).ToolTipText = funLoadResString(10591)
    cmdCrdButtons(7).ToolTipText = funLoadResString(10687)
    cmdCrdButtons(3).Caption = funLoadResString(10688)
    cmdCrdButtons(3).ToolTipText = funLoadResString(10689)
    cmdCrdButtons(4).Caption = funLoadResString(10690)
    cmdCrdButtons(4).ToolTipText = funLoadResString(10691)
    
    cmdCliButtons(0).Caption = funLoadResString(10803) & "..."
    cmdCliButtons(0).ToolTipText = funLoadResString(10806)
    cmdCliButtons(1).Caption = funLoadResString(10141) & "..."
    cmdCliButtons(1).ToolTipText = funLoadResString(10807)
    cmdCliFind.ToolTipText = funLoadResString(10808)
    cmdCliButtons(2).Caption = funLoadResString(10805) & "..."
    cmdCliButtons(2).ToolTipText = funLoadResString(10809)
    txtCliFind.ToolTipText = funLoadResString(10810)
    cmdCliButtons(5).Caption = funLoadResString(10192) & "..."
    cmdCliButtons(5).ToolTipText = funLoadResString(10811)
    cmdCliButtons(6).Caption = funLoadResString(10590)
    cmdCliButtons(6).ToolTipText = funLoadResString(10814)
    cmdCliButtons(7).ToolTipText = funLoadResString(10687)
    cmdCliButtons(3).Caption = funLoadResString(10688)
    cmdCliButtons(3).ToolTipText = funLoadResString(10812)
    cmdCliButtons(4).Caption = funLoadResString(10690)
    cmdCliButtons(4).ToolTipText = funLoadResString(10813)
    
    cmdWksnButtons(10).Caption = funLoadResString(10690)
    
    cmdWksnButtons(0).Caption = funLoadResString(10155)
    cmdWksnButtons(0).ToolTipText = funLoadResString(10463)
    cmdConsoleButtons(0).Caption = funLoadResString(10155)
    cmdConsoleButtons(0).ToolTipText = funLoadResString(10639)
    
    cmdWksnButtons(2).Caption = funLoadResString(10156)
    cmdWksnButtons(2).ToolTipText = funLoadResString(10464)
    cmdWksnButtons(3).Caption = funLoadResString(10614)
    cmdWksnButtons(3).ToolTipText = funLoadResString(10982)
    cmdWksnButtons(6).ToolTipText = funLoadResString(10687)
    cmdConsoleButtons(1).Caption = funLoadResString(10156)
    cmdConsoleButtons(1).ToolTipText = funLoadResString(10640)
    cmdConsoleButtons(2).Caption = funLoadResString(10614)
    cmdConsoleButtons(2).ToolTipText = funLoadResString(10642)
    
    cmdWksnButtons(9).Caption = funLoadResString(10141) & "..."
    
    cmdWksnButtons(4).Caption = funLoadResString(10157)
    cmdWksnButtons(4).ToolTipText = funLoadResString(10466)
    cmdConsoleButtons(3).Caption = funLoadResString(10157)
    cmdConsoleButtons(3).ToolTipText = funLoadResString(10641)
    
    cmdWksnButtons(5).Caption = funLoadResString(10159)
    cmdWksnButtons(5).ToolTipText = funLoadResString(10467)
    cmdWksnButtons(13).ToolTipText = funLoadResString(10962)
    
    cmdConsoleButtons(5).Caption = funLoadResString(10159)
    cmdConsoleButtons(5).ToolTipText = funLoadResString(10643)
    cmdConsoleButtons(6).Caption = funLoadResString(10597)
    cmdConsoleButtons(6).ToolTipText = funLoadResString(10644)
    cmdConsoleButtons(7).Caption = funLoadResString(10615)
    cmdConsoleButtons(7).ToolTipText = funLoadResString(10645)
    
    cmdWksnButtons(1).Caption = funLoadResString(10160)
    cmdWksnButtons(1).ToolTipText = funLoadResString(10469)
    
    cmdWksnButtons(7).Caption = funLoadResString(10161)
    cmdWksnButtons(7).ToolTipText = funLoadResString(10468)
    
    cmdWksnButtons(8).Caption = funLoadResString(10161)
    cmdWksnButtons(8).ToolTipText = funLoadResString(10468)
    cmdConsoleButtons(4).Caption = funLoadResString(10161)
    cmdConsoleButtons(4).ToolTipText = funLoadResString(10646)
    
    cmdConsoleButtons(8).ToolTipText = funLoadResString(10647)
    
    cmdWksnButtons(11).ToolTipText = funLoadResString(10572)
    
    cmdWksnButtons(14).Caption = funLoadResString(10980) & "..."
    cmdWksnButtons(14).ToolTipText = funLoadResString(10981)
    
    cmdWksnButtons(15).Caption = funLoadResString(11023)
    cmdWksnButtons(15).ToolTipText = funLoadResString(11024)
    
    cmdIncomeFilter.Caption = funLoadResString(10306)
    cmdIncomeFilter.ToolTipText = funLoadResString(10455)
    cmdIncomeExport.Caption = funLoadResString(10307)
    cmdIncomeExport.ToolTipText = funLoadResString(10454)
    cmdIncomePrint.Caption = funLoadResString(10965)
    cmdIncomePrint.ToolTipText = funLoadResString(10966, guServOpts.uConfig.strDefaultPrinterName)
    cmdServicesCollect.Caption = funLoadResString(10160)
    cmdServicesCollect.ToolTipText = funLoadResString(10462)

    cmdSessionFilter.Caption = funLoadResString(10306)
    cmdSessionFilter.ToolTipText = funLoadResString(10453)
    cmdSessionExport.Caption = funLoadResString(10307)
    cmdSessionExport.ToolTipText = funLoadResString(10454)
    cmdSessionPrint.Caption = funLoadResString(10965)
    cmdSessionPrint.ToolTipText = funLoadResString(10966, guServOpts.uConfig.strDefaultPrinterName)

    chkWksnConnect.Caption = funLoadResString(10193)
    chkWksnConnect.ToolTipText = funLoadResString(10470)

    chkTypeIncome(1).Caption = funLoadResString(10308)
    chkTypeIncome(2).Caption = funLoadResString(10309)
    chkTypeIncome(3).Caption = funLoadResString(10310)
    chkTypeIncome(4).Caption = funLoadResString(10831)
    chkTypeIncome(5).Caption = funLoadResString(10985)

    optConsoleWorkstationsAll.Caption = funLoadResString(10616)
    optConsoleWorkstationsSome.Caption = funLoadResString(10617)

    cmdSessionURLSearch.ToolTipText = funLoadResString(10718)

    txtWksnChat.ToolTipText = funLoadResString(10961)
    txtWksnChatHist.ToolTipText = funLoadResString(10960)

    With fxgCards
        .Cols = 4
        .TextMatrix(0, 0) = funLoadResString(10194)
        .TextMatrix(0, 1) = funLoadResString(10195)
        .TextMatrix(0, 2) = funLoadResString(10196)
        .TextMatrix(0, 3) = funLoadResString(10197)
        .ToolTipText = funLoadResString(10456)
    End With

    With fxgCrdPending
        .Cols = 3
        .TextMatrix(0, 0) = funLoadResString(10194)
        .TextMatrix(0, 1) = funLoadResString(10195)
        .TextMatrix(0, 2) = funLoadResString(10144)
        If gblnDBOK Then .ToolTipText = funLoadResString(10953, grstCardsPending.RecordCount)
    End With

    With fxgClients
        .Cols = 2
        .TextMatrix(0, 0) = funLoadResString(10798)
        .TextMatrix(0, 1) = funLoadResString(10195)
        .ToolTipText = funLoadResString(10799)
    End With

    With fxgIncome
        .Cols = 8
        .TextMatrix(0, 0) = funLoadResString(10293)
        .TextMatrix(0, 1) = funLoadResString(10294)
        .TextMatrix(0, 2) = funLoadResString(10297)
        .TextMatrix(0, 3) = funLoadResString(10295)
        .TextMatrix(0, 4) = funLoadResString(10404)
        .TextMatrix(0, 5) = funLoadResString(10652)
        .TextMatrix(0, 6) = funLoadResString(10296)
        .TextMatrix(0, 7) = funLoadResString(10754)
    End With

    With fxgServices
        .Cols = 3
        .TextMatrix(0, 0) = funLoadResString(10328)
        .TextMatrix(0, 1) = funLoadResString(10329)
        .TextMatrix(0, 2) = funLoadResString(10330)
        .ToolTipText = funLoadResString(10461)
    End With

    With fxgSession(1)
        .Cols = 9 + guServOpts.uWorkstations.lngNumWorkstations * 2
        .TextMatrix(0, 0) = funLoadResString(10378)
        .TextMatrix(0, 1) = funLoadResString(10379)
        .TextMatrix(0, 2) = funLoadResString(10380)
        .TextMatrix(0, 3) = funLoadResString(10381)
        .TextMatrix(0, 4) = funLoadResString(10382)
        .TextMatrix(0, 5) = funLoadResString(10386)
        .TextMatrix(0, 6) = funLoadResString(10653)
        .TextMatrix(0, 7) = " " & funLoadResString(10383) & funLoadResString(10385)
        .TextMatrix(0, 8) = funLoadResString(10386)
        .TextMatrix(0, 9) = " " & funLoadResString(10384)
        .TextMatrix(0, 10) = funLoadResString(10386)
        For i = 2 To (guServOpts.uWorkstations.lngNumWorkstations - 1) * 2 Step 2
            .TextMatrix(0, i + 9) = " " & i / 2 + 1 & funLoadResString(10385)
            .TextMatrix(0, i + 10) = funLoadResString(10386)
        Next i
    End With

    With fxgSession(3)
        .Cols = 11
        .TextMatrix(0, 0) = funLoadResString(10403)
        .TextMatrix(0, 1) = funLoadResString(10404)
        .TextMatrix(0, 2) = funLoadResString(10405)
        .TextMatrix(0, 3) = funLoadResString(10409)
        .TextMatrix(0, 4) = funLoadResString(10406)
        .TextMatrix(0, 5) = funLoadResString(10386)
        .TextMatrix(0, 6) = funLoadResString(10407)
        .TextMatrix(0, 7) = funLoadResString(10386)
        .TextMatrix(0, 8) = funLoadResString(10408)
        .TextMatrix(0, 9) = funLoadResString(10386)
        .TextMatrix(0, 10) = funLoadResString(10754)
    End With

    With fxgSession(4)
        .Cols = 5
        .TextMatrix(0, 0) = funLoadResString(10403)
        .TextMatrix(0, 1) = funLoadResString(10405)
        .TextMatrix(0, 2) = funLoadResString(10404)
        .TextMatrix(0, 3) = funLoadResString(10695)
        .TextMatrix(0, 4) = funLoadResString(10696)

    End With

    chkSessStatType(1).Caption = funLoadResString(10401)
    chkSessStatType(2).Caption = funLoadResString(10421)
    chkSessStatType(3).Caption = funLoadResString(10402)
    chkSessStatType(4).Caption = funLoadResString(10698)
    chkSessStatType(5).Caption = funLoadResString(10422)


    intNumIntervals = funArraySize(guServOpts.uStats.strStatsTimeInt)
    With fxgSession(2)
        .Cols = 2 + 2 * (intNumIntervals + 1)
        .TextMatrix(0, 0) = funLoadResString(10378)
        .TextMatrix(0, 1) = funLoadResString(10380)
        strTmpLabel = " 00:00"
        For i = 0 To intNumIntervals * 2 Step 2
            If intNumIntervals = 0 Or i = intNumIntervals * 2 Then
                strTmpLabel = strTmpLabel & "-24:00"
                .TextMatrix(0, i + 2) = strTmpLabel
                .TextMatrix(0, i + 3) = funLoadResString(10386)
            Else
                strTmpLabel = strTmpLabel & "-" & left(funFormatTime(funDateTimefromNet(guServOpts.uStats.strStatsTimeInt(i / 2))), 5)
                .TextMatrix(0, i + 2) = strTmpLabel
                .TextMatrix(0, i + 3) = funLoadResString(10386)
                strTmpLabel = " " & left(funFormatTime(funDateTimefromNet(guServOpts.uStats.strStatsTimeInt(i / 2))), 5)
            End If
        Next i
    End With

    With fxgSession(5)
        .TextMatrix(0, 0) = funLoadResString(10378)
    End With
    Exit Sub
    
subLocalizeInterfaceError:
    Call funErrorManager("frmMain.subLocalizeInterfaceError", True)
End Sub

