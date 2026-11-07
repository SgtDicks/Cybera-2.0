VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmBlock 
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   ClientHeight    =   10740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12855
   Icon            =   "frmBlock.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   716
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   857
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrUnblock 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   12240
      Top             =   2100
   End
   Begin VB.Frame fraFooter 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   735
      Left            =   120
      TabIndex        =   48
      Top             =   8160
      Width           =   11775
      Begin VB.Label lblAppVersion 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Version info"
         ForeColor       =   &H80000011&
         Height          =   615
         Left            =   6960
         TabIndex        =   50
         Top             =   60
         Width           =   4755
      End
      Begin VB.Label lblGNU 
         BackStyle       =   0  'Transparent
         Caption         =   "GNU"
         ForeColor       =   &H80000011&
         Height          =   615
         Left            =   60
         TabIndex        =   49
         Top             =   60
         Width           =   3075
      End
      Begin VB.Shape shpFooter 
         BorderColor     =   &H80000010&
         Height          =   195
         Left            =   4920
         Top             =   300
         Width           =   195
      End
   End
   Begin VB.Frame fraHeader 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   600
      Left            =   120
      TabIndex        =   47
      Top             =   120
      Width           =   11775
      Begin VB.Timer tmrDefaultLanguage 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   3540
         Top             =   120
      End
      Begin VB.Frame fraNoServCnx 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   10020
         TabIndex        =   52
         Top             =   60
         Width           =   1695
         Begin VB.Timer tmrNoServCnx 
            Enabled         =   0   'False
            Interval        =   500
            Left            =   120
            Top             =   0
         End
         Begin MSComctlLib.ImageList imNoServCnx 
            Left            =   600
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   3
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmBlock.frx":0442
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmBlock.frx":174C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmBlock.frx":2A56
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Image imgNoServCnx 
            Height          =   480
            Left            =   1200
            Picture         =   "frmBlock.frx":3D60
            Top             =   0
            Width           =   480
         End
      End
      Begin VB.ComboBox cmbLanguage 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         ItemData        =   "frmBlock.frx":41A2
         Left            =   600
         List            =   "frmBlock.frx":41A4
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   80
         Width           =   2895
      End
      Begin VB.Image imgFlags 
         Height          =   480
         Left            =   60
         Picture         =   "frmBlock.frx":41A6
         Top             =   60
         Width           =   480
      End
      Begin VB.Label lblWksnName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2820
         TabIndex        =   51
         Top             =   0
         Width           =   6015
      End
      Begin VB.Shape shpHeader 
         BorderColor     =   &H80000010&
         Height          =   195
         Left            =   9360
         Top             =   360
         Width           =   195
      End
   End
   Begin VB.Timer tmrActivateTimeout 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   12240
      Top             =   2640
   End
   Begin VB.Frame fraFollowServer 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   300
      TabIndex        =   38
      Top             =   9240
      Visible         =   0   'False
      Width           =   2295
      Begin VB.Image imgFollowServer 
         Height          =   1275
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   915
      End
   End
   Begin VB.Timer tmrTopMost 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   12240
      Top             =   1500
   End
   Begin VB.Timer tmrCloseSession 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   12240
      Top             =   960
   End
   Begin VB.Frame fraForm 
      BorderStyle     =   0  'None
      Height          =   7005
      Left            =   660
      TabIndex        =   12
      Top             =   960
      Width           =   10740
      Begin VB.TextBox txtLoginToolTip 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   525
         Left            =   4620
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   46
         Top             =   6000
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txtCardNumToolTip 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3780
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   45
         Top             =   6240
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Frame fraMessage 
         BackColor       =   &H80000014&
         BorderStyle     =   0  'None
         Height          =   5115
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   5115
         Begin RichTextLib.RichTextBox rtbMessage 
            Height          =   4995
            Left            =   60
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   60
            Width           =   4995
            _ExtentX        =   8811
            _ExtentY        =   8811
            _Version        =   393217
            BackColor       =   -2147483628
            BorderStyle     =   0
            Enabled         =   0   'False
            ReadOnly        =   -1  'True
            Appearance      =   0
            TextRTF         =   $"frmBlock.frx":44B0
         End
         Begin VB.Shape shpFrame 
            BorderColor     =   &H80000010&
            Height          =   5115
            Index           =   1
            Left            =   0
            Top             =   0
            Width           =   5115
         End
      End
      Begin VB.Frame fraControlsActive 
         BorderStyle     =   0  'None
         Height          =   1275
         Left            =   120
         TabIndex        =   15
         Top             =   5340
         Width           =   10395
         Begin VB.Frame fra 
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
            Height          =   1215
            Index           =   0
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   3135
            Begin VB.OptionButton OptClient 
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
               Left            =   120
               TabIndex        =   3
               ToolTipText     =   "Cliquez ici si vous utilisez une carte Cyber"
               Top             =   600
               Width           =   2895
            End
            Begin VB.OptionButton optCard 
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
               Left            =   120
               TabIndex        =   2
               ToolTipText     =   "Cliquez ici si vous utilisez une carte Cyber"
               Top             =   300
               Value           =   -1  'True
               Width           =   2895
            End
            Begin VB.OptionButton optSelfService 
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
               Left            =   120
               TabIndex        =   4
               Top             =   900
               Width           =   2895
            End
            Begin VB.Shape shpFrame 
               BorderColor     =   &H80000010&
               Height          =   1215
               Index           =   4
               Left            =   0
               Top             =   0
               Width           =   3075
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
               TabIndex        =   21
               Top             =   0
               Width           =   3075
            End
         End
         Begin VB.Frame fra 
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
            Height          =   1215
            Index           =   2
            Left            =   6000
            TabIndex        =   16
            Top             =   0
            Width           =   2595
            Begin VB.OptionButton optConnected 
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
               Height          =   195
               Left            =   120
               TabIndex        =   5
               ToolTipText     =   "Cliquez ici pour utiliser le mode connecté à internet"
               Top             =   480
               Value           =   -1  'True
               Width           =   2355
            End
            Begin VB.OptionButton optDisconnected 
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
               Height          =   195
               Left            =   120
               TabIndex        =   6
               ToolTipText     =   "Cliquez ici pour utiliser le mode déconnecté. Vous n'aurez pas accès à internet"
               Top             =   840
               Width           =   2355
            End
            Begin VB.Shape shpFrame 
               BorderColor     =   &H80000010&
               Height          =   1215
               Index           =   5
               Left            =   0
               Top             =   0
               Width           =   2535
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
               Index           =   3
               Left            =   0
               TabIndex        =   17
               Top             =   0
               Width           =   2535
            End
         End
         Begin VB.CommandButton cmdCxn 
            Caption         =   "_"
            Default         =   -1  'True
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   8640
            Picture         =   "frmBlock.frx":4532
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Cliquez ici pour démarrer votre session"
            Top             =   0
            Width           =   1695
         End
         Begin VB.Frame fraClient 
            BorderStyle     =   0  'None
            Caption         =   "Numéro de carte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   3180
            TabIndex        =   39
            Top             =   0
            Visible         =   0   'False
            Width           =   2775
            Begin VB.TextBox txtPassword 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               IMEMode         =   3  'DISABLE
               Left            =   1320
               MaxLength       =   20
               PasswordChar    =   "*"
               TabIndex        =   44
               ToolTipText     =   "Entrer votre numéro de carte ici"
               Top             =   900
               Width           =   1215
            End
            Begin VB.TextBox txtLogin 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   1320
               MaxLength       =   255
               TabIndex        =   40
               ToolTipText     =   "Entrer votre numéro de carte ici"
               Top             =   420
               Width           =   1215
            End
            Begin VB.Shape shpFrame 
               BorderColor     =   &H80000010&
               Height          =   1215
               Index           =   12
               Left            =   0
               Top             =   0
               Width           =   2715
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "_"
               Height          =   195
               Index           =   15
               Left            =   60
               TabIndex        =   43
               Top             =   840
               Width           =   1155
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   14
               Left            =   60
               TabIndex        =   42
               Top             =   420
               Width           =   1155
            End
            Begin VB.Shape shpFrame 
               BackStyle       =   1  'Opaque
               BorderColor     =   &H80000010&
               Height          =   285
               Index           =   13
               Left            =   1260
               Top             =   840
               Width           =   1335
            End
            Begin VB.Shape shpFrame 
               BackStyle       =   1  'Opaque
               BorderColor     =   &H80000010&
               Height          =   375
               Index           =   11
               Left            =   1260
               Top             =   360
               Width           =   1335
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
               Index           =   13
               Left            =   0
               TabIndex        =   41
               Top             =   0
               Width           =   2715
            End
         End
         Begin VB.Frame fraCardNumber 
            BorderStyle     =   0  'None
            Caption         =   "Numéro de carte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   3180
            TabIndex        =   18
            Top             =   0
            Width           =   2775
            Begin VB.TextBox txtCardNum 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   480
               MaxLength       =   6
               TabIndex        =   1
               ToolTipText     =   "Entrer votre numéro de carte ici"
               Top             =   600
               Width           =   1755
            End
            Begin VB.Shape shpFrame 
               BorderColor     =   &H80000010&
               Height          =   1215
               Index           =   3
               Left            =   0
               Top             =   0
               Width           =   2715
            End
            Begin VB.Shape shpFrame 
               BackStyle       =   1  'Opaque
               BorderColor     =   &H80000010&
               Height          =   435
               Index           =   2
               Left            =   420
               Top             =   540
               Width           =   1875
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
               Left            =   0
               TabIndex        =   19
               Top             =   0
               Width           =   2715
            End
         End
      End
      Begin VB.Frame fraControlsSessionEnd 
         BorderStyle     =   0  'None
         Height          =   1275
         Left            =   120
         TabIndex        =   28
         Top             =   5340
         Width           =   10395
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
            Height          =   375
            Index           =   9
            Left            =   60
            TabIndex        =   30
            Top             =   660
            Width           =   10215
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
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
            Height          =   375
            Index           =   8
            Left            =   60
            TabIndex        =   29
            Top             =   120
            Width           =   10215
         End
         Begin VB.Shape shpFrame 
            BackColor       =   &H80000014&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000010&
            Height          =   1215
            Index           =   8
            Left            =   0
            Top             =   0
            Width           =   10335
         End
      End
      Begin VB.Frame fraControlsLocked 
         BorderStyle     =   0  'None
         Height          =   1275
         Left            =   120
         TabIndex        =   35
         Top             =   5340
         Width           =   10395
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
            Height          =   375
            Index           =   12
            Left            =   60
            TabIndex        =   37
            Top             =   660
            Width           =   10215
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
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
            Height          =   375
            Index           =   0
            Left            =   60
            TabIndex        =   36
            Top             =   120
            Width           =   10215
         End
         Begin VB.Shape shpFrame 
            BackColor       =   &H80000014&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000010&
            Height          =   1215
            Index           =   10
            Left            =   0
            Top             =   0
            Width           =   10335
         End
      End
      Begin VB.Frame fraControlsBlocked 
         BorderStyle     =   0  'None
         Height          =   1275
         Left            =   120
         TabIndex        =   25
         Top             =   5340
         Width           =   10395
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
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
            Height          =   375
            Index           =   7
            Left            =   60
            TabIndex        =   27
            Top             =   120
            Width           =   10215
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
            Height          =   375
            Index           =   6
            Left            =   60
            TabIndex        =   26
            Top             =   660
            Width           =   10215
         End
         Begin VB.Shape shpFrame 
            BackColor       =   &H80000014&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000010&
            Height          =   1215
            Index           =   7
            Left            =   0
            Top             =   0
            Width           =   10335
         End
      End
      Begin VB.Frame fraControlsNoMoreCredit 
         BorderStyle     =   0  'None
         Height          =   1275
         Left            =   120
         TabIndex        =   22
         Top             =   5340
         Width           =   10395
         Begin VB.CommandButton cmdQuit 
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
            Height          =   1215
            Index           =   0
            Left            =   7320
            Picture         =   "frmBlock.frx":467C
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Cliquez ici pour fermer votre session"
            Top             =   0
            Width           =   1395
         End
         Begin VB.CommandButton cmdReCnx 
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
            Height          =   1215
            Left            =   8760
            Picture         =   "frmBlock.frx":47C6
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Cliquez ici pour reprendre votre session"
            Top             =   0
            Width           =   1575
         End
         Begin VB.Label lblCloseSessionWarning 
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
            Left            =   60
            TabIndex        =   34
            Top             =   480
            Width           =   7095
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
            Index           =   5
            Left            =   60
            TabIndex        =   24
            Top             =   780
            Width           =   7095
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
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
            Height          =   375
            Index           =   4
            Left            =   60
            TabIndex        =   23
            Top             =   60
            Width           =   7095
         End
         Begin VB.Shape shpFrame 
            BackColor       =   &H80000014&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000010&
            Height          =   1215
            Index           =   6
            Left            =   0
            Top             =   0
            Width           =   7215
         End
      End
      Begin VB.Frame fraControlsPaused 
         BorderStyle     =   0  'None
         Height          =   1275
         Left            =   120
         TabIndex        =   31
         Top             =   5340
         Width           =   10395
         Begin VB.CommandButton cmdUnPause 
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
            Height          =   1215
            Left            =   8760
            Picture         =   "frmBlock.frx":4910
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Cliquez ici pour reprendre votre session"
            Top             =   0
            Width           =   1575
         End
         Begin VB.CommandButton cmdQuit 
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
            Height          =   1215
            Index           =   1
            Left            =   7320
            Picture         =   "frmBlock.frx":4A5A
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Cliquez ici pour fermer votre session"
            Top             =   0
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
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
            Height          =   375
            Index           =   11
            Left            =   60
            TabIndex        =   33
            Top             =   120
            Width           =   7095
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
            Height          =   375
            Index           =   10
            Left            =   60
            TabIndex        =   32
            Top             =   660
            Width           =   7095
         End
         Begin VB.Shape shpFrame 
            BackColor       =   &H80000014&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000010&
            Height          =   1215
            Index           =   9
            Left            =   0
            Top             =   0
            Width           =   7215
         End
      End
      Begin VB.Image imgLogo 
         Height          =   4995
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   180
         Width           =   4995
      End
      Begin VB.Shape shpFrame 
         BorderColor     =   &H80000010&
         Height          =   5115
         Index           =   0
         Left            =   5340
         Top             =   120
         Width           =   5115
      End
      Begin VB.Shape shpFormFrame 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   6675
         Left            =   0
         Top             =   0
         Width           =   10575
      End
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   0
      X2              =   800
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   800
      X2              =   800
      Y1              =   0
      Y2              =   600
   End
End
Attribute VB_Name = "frmBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ SCREEN BLOCK FORM __________________________|
'|  Name:          frmBlock.frm                                                |
'|  Type:          VB6 Form                                                    |
'|  Project:       Cybera Workstation (http://cybera.sf.net)                   |
'|  Description:   The Screen block form.                                      |
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

' $Id: frmBlock.frm 15 2006-01-01 16:12:35Z sam $

Option Explicit
Dim mstrLangPacks() As String


'*******************************************************************************
Private Sub cmbLanguage_Click()
Dim frmTemp As Form
Dim strDefaultFilePrefix As String
On Error GoTo cmbLanguage_ClickError

    guCliOpts.strInterfaceLang = mstrLangPacks(0, cmbLanguage.ItemData(cmbLanguage.ListIndex))
    guCliOpts.intInterfaceLangCharset = mstrLangPacks(3, cmbLanguage.ItemData(cmbLanguage.ListIndex))
    For Each frmTemp In Forms
        Call frmTemp.subLocalizeInterface
    Next frmTemp
    strDefaultFilePrefix = FILENAME_PREFIX & Right(guCliOpts.strInterfaceLang, Len(guCliOpts.strInterfaceLang) - Len(App.EXEName) - 1) & "_"
    Select Case guUser.intWksnState
    Case enuWksnState.Blocked_Deactivated
        frmBlock.rtbMessage.FileName = IIf(guCliOpts.intMsgLogoFilesType = vbChecked, guCliOpts.uPaths.strBlockedPath, guCliOpts.strAppConfigPath & strDefaultFilePrefix & FILENAME_BLOCKED)
    
    Case enuWksnState.Blocked_SessionEnd
        frmBlock.rtbMessage.FileName = IIf(guCliOpts.intMsgLogoFilesType = vbChecked, guCliOpts.uPaths.strSessionEndPath, guCliOpts.strAppConfigPath & strDefaultFilePrefix & FILENAME_SESSIONEND)
    
    Case enuWksnState.Blocked_NoCredit
        frmBlock.rtbMessage.FileName = IIf(guCliOpts.intMsgLogoFilesType = vbChecked, guCliOpts.uPaths.strNoCreditPath, guCliOpts.strAppConfigPath & strDefaultFilePrefix & FILENAME_NOCREDIT)
    
    Case enuWksnState.Blocked_Paused, enuWksnState.Blocked_Locked
        frmBlock.rtbMessage.FileName = IIf(guCliOpts.intMsgLogoFilesType = vbChecked, guCliOpts.uPaths.strSessionPausedPath, guCliOpts.strAppConfigPath & strDefaultFilePrefix & FILENAME_SESSIONPAUSED)
    
    Case enuWksnState.Ready
        frmBlock.rtbMessage.FileName = IIf(guCliOpts.intMsgLogoFilesType = vbChecked, guCliOpts.uPaths.strIntroPath, guCliOpts.strAppConfigPath & strDefaultFilePrefix & FILENAME_INTRO)
    
    End Select
'    tmrTopMost.Enabled = True
    Exit Sub

cmbLanguage_ClickError:
    Call funErrorManager("frmBlock.cmbLanguage_ClickError", , True)
End Sub


'*******************************************************************************
Private Sub cmbLanguage_DropDown()
On Error Resume Next
    
    tmrTopMost.Enabled = False
End Sub


'*******************************************************************************
Private Sub cmbLanguage_LostFocus()
On Error Resume Next
    
'    tmrTopMost.Enabled = True
End Sub


'*******************************************************************************
Private Sub cmdCxn_Click()
Dim intSessionType As enuSessionType
Dim intWksnState As enuWksnState
Dim strID As String, strPasswd As String
On Error GoTo cmdCxn_ClickError
    
    If optCard.Value = True Then
        If txtCardNum.Text = "" Then Exit Sub
        intSessionType = enuSessionType.Card
        
        If Not IsNumeric(txtCardNum.Text) Or Len(CStr(CLng(txtCardNum.Text))) <> 6 Then
            Call funSetFocus(txtCardNum)
            Call subShowMsgBox(funLoadResString(10104), , , , Me, vbModal)
            Exit Sub
        Else
            strID = txtCardNum.Text
        End If
    ElseIf OptClient.Value Then
        If txtLogin.Text = "" Then Exit Sub
        intSessionType = enuSessionType.Client
        strID = LCase(txtLogin.Text)
        strPasswd = txtPassword.Text
    ElseIf optSelfService.Value Then
        intSessionType = enuSessionType.Selfservice
    End If
    If optConnected.Value = True Then
        intWksnState = enuWksnState.UsedConnected
    Else
        intWksnState = enuWksnState.UsedDisconnected
    End If

    Call subStartConnexion(intSessionType, intWksnState, strID, strPasswd)
    Exit Sub

cmdCxn_ClickError:
    Call subDebugPrint("frmBlock.cmdCxn_ClickError - CardNum: " & txtCardNum.Text)
    Call funErrorManager("frmBlock.cmdCxn_ClickError", , True)
End Sub


'*******************************************************************************
Private Sub cmdQuit_Click(Index As Integer)
On Error GoTo cmdQuit_ClickError

    Select Case Index
    Case 0
        Call funSetFocus(cmdReCnx)
    Case 1
        Call funSetFocus(cmdUnPause)
    End Select
    Call subShowMsgBox(funLoadResString(10184), YesNo, Button2, New clsActionCloseSession, Me, vbModal)
    Exit Sub

cmdQuit_ClickError:
    Call funErrorManager("frmBlock.cmdQuit_ClickError", , True)
End Sub


'*******************************************************************************
Private Sub cmdReCnx_Click()
On Error GoTo cmdReCnx_ClickError
    
    Call subCheckBalance
Exit Sub

cmdReCnx_ClickError:
    Call funErrorManager("frmBlock.cmdReCnx_ClickError", , True)
End Sub


'*******************************************************************************
Private Sub cmdUnPause_Click()
On Error GoTo cmdUnPause_ClickError

    Call funUnPauseSession
    Exit Sub

cmdUnPause_ClickError:
    Call funErrorManager("frmBlock.cmdUnPause_ClickError", , True)
End Sub


'*******************************************************************************
Private Sub Form_Activate()
On Error Resume Next
    
    gblnfrmBlock = True
End Sub


'*******************************************************************************
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Form_KeyDownError
    
    If KeyCode = vbKeyReturn And Shift = vbCtrlMask + vbAltMask Then
        If Not gblnNoTopmost Then Call SetWindowPos(frmBlock.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
        If guCliOpts.strPasswd = "" And guCliOpts.uConnectedManager.Passwd = "" Then
            Call subDeactivate
        Else
            Call subDeactivationPassword(Me)
        End If
        If Not gblnNoTopmost Then Call SetWindowPos(frmBlock.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    End If
    Exit Sub

Form_KeyDownError:
    Call funErrorManager("frmBlock.Form_KeyDownError", , True)
End Sub

'*******************************************************************************
Private Sub Form_Load()
Dim i As Integer
On Error Resume Next


    Call subSetDefaultFont(Me)
    
    mstrLangPacks = funFindLangPacks
    For i = 0 To UBound(mstrLangPacks, 2)
        If mstrLangPacks(5, i) = True Then
            Call cmbLanguage.AddItem(mstrLangPacks(4, i))
            cmbLanguage.ItemData(cmbLanguage.NewIndex) = i
        End If
    Next i
    
    Call subLocalizeInterface
    
    frmBlock.WindowState = vbMaximized
    If Not gblnNoTopmost Then Call SetWindowPos(frmBlock.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    fraForm.Top = Screen.Height / Screen.TwipsPerPixelY / 2 - fraForm.Height / 2
    fraForm.Left = Screen.Width / Screen.TwipsPerPixelX / 2 - fraForm.Width / 2
    fraForm.Width = shpFormFrame.Width / Screen.TwipsPerPixelX
    fraForm.Height = shpFormFrame.Height / Screen.TwipsPerPixelY
    fraFollowServer.Top = 0
    fraFollowServer.Left = 0
    fraFollowServer.Height = Screen.Height / Screen.TwipsPerPixelY
    fraFollowServer.Width = Screen.Width / Screen.TwipsPerPixelX
    imgFollowServer.Top = 0
    imgFollowServer.Left = 0
    imgFollowServer.Height = Screen.Height
    imgFollowServer.Width = Screen.Width
    
    fraHeader.Top = 8
    fraHeader.Left = fraHeader.Top
    fraHeader.Width = Screen.Width / Screen.TwipsPerPixelX - fraHeader.Left * 2
    shpHeader.Top = 0
    shpHeader.Left = 0
    shpHeader.Width = fraHeader.Width * Screen.TwipsPerPixelX
    shpHeader.Height = fraHeader.Height * Screen.TwipsPerPixelY
    
    fraNoServCnx.Left = shpHeader.Width - fraNoServCnx.Width - 4 * Screen.TwipsPerPixelY
    
    fraFooter.Top = Screen.Height / Screen.TwipsPerPixelY - fraFooter.Height - fraHeader.Top
    fraFooter.Left = fraHeader.Top
    fraFooter.Width = Screen.Width / Screen.TwipsPerPixelX - fraFooter.Left * 2
    shpFooter.Top = 0
    shpFooter.Left = 0
    shpFooter.Width = fraFooter.Width * Screen.TwipsPerPixelX
    shpFooter.Height = fraFooter.Height * Screen.TwipsPerPixelY
    
'    lblGNU.Top = 4 * Screen.TwipsPerPixelY
    lblAppVersion.Left = shpFooter.Width - lblAppVersion.Width - 4 * Screen.TwipsPerPixelY

    lblWksnName.Left = (shpHeader.Width - lblWksnName.Width) / 2
    
End Sub


'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)

    gblnfrmBlock = False
End Sub


'*******************************************************************************
Private Sub imgFollowServer_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo imgFollowServer_MouseDownError

    If Button = vbLeftButton And Shift = vbCtrlMask + vbAltMask Then
        If Not gblnNoTopmost Then Call SetWindowPos(frmBlock.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
        If guCliOpts.strPasswd <> "" Then
            Call subDeactivationPassword(Me)
        Else
            Call subDeactivate
        End If
        If Not gblnNoTopmost Then Call SetWindowPos(frmBlock.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    End If
    Exit Sub

imgFollowServer_MouseDownError:
    Call funErrorManager("frmBlock.imgFollowServer_MouseDownError", , True)
End Sub


'*******************************************************************************
Private Sub imgLogo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo imgLogo_MouseDownError

    If Button = vbLeftButton And Shift = vbCtrlMask + vbAltMask Then
        If Not gblnNoTopmost Then Call SetWindowPos(frmBlock.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
        If guCliOpts.strPasswd = "" And guCliOpts.uConnectedManager.Passwd = "" Then
            Call subDeactivate
        Else
            Call subDeactivationPassword(Me)
        End If
        If Not gblnNoTopmost Then Call SetWindowPos(frmBlock.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    End If
    Exit Sub

imgLogo_MouseDownError:
    Call funErrorManager("frmBlock.imgLogo_MouseDownError", , True)
End Sub


'*******************************************************************************
Private Sub optCard_Click()
On Error Resume Next

    fraCardNumber.Visible = True
    fraClient.Visible = False
    cmdCxn.Enabled = txtCardNum.Text <> "" And guCliOpts.uServer.intCnxState = enuSrvCxn.Connected
    Call funSetFocus(txtCardNum)
End Sub


'*******************************************************************************
Private Sub optCard_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

    Call funSetFocus(txtCardNum)
End Sub


'*******************************************************************************
Private Sub OptClient_Click()
On Error Resume Next

    fraCardNumber.Visible = False
    fraClient.Visible = True
    cmdCxn.Enabled = txtLogin.Text <> "" And guCliOpts.uServer.intCnxState = enuSrvCxn.Connected
    Call funSetFocus(txtLogin)
End Sub


'*******************************************************************************
Private Sub OptClient_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

    Call funSetFocus(txtLogin)
End Sub

'*******************************************************************************
Private Sub optSelfService_Click()
On Error Resume Next
    
    fraCardNumber.Visible = False
    fraClient.Visible = False
    cmdCxn.Enabled = guCliOpts.uServer.intCnxState = enuSrvCxn.Connected
End Sub


'*******************************************************************************
Private Sub tmrActivateTimeout_Timer()
On Error Resume Next

    cmdCxn.Enabled = ((fraCardNumber.Visible And txtCardNum.Text <> "") Or (fraClient.Visible And txtLogin.Text <> "")) And guCliOpts.uServer.intCnxState = enuSrvCxn.Connected
    frmBlock.txtCardNum.Enabled = True
    frmBlock.txtLogin.Enabled = True
    frmBlock.txtPassword.Enabled = True
    tmrActivateTimeout.Enabled = True
    If fraCardNumber.Visible Then
        Call funSetFocus(txtCardNum)
    ElseIf fraClient.Visible Then
        Call funSetFocus(txtLogin)
    End If
    frmBlock.MousePointer = vbNormal
    tmrActivateTimeout.Enabled = False
End Sub


'*******************************************************************************
Private Sub tmrCloseSession_Timer()
On Error GoTo tmrCloseSession_TimerError

    gdteExpireryTime = DateAdd("s", -1, gdteExpireryTime)
    lblCloseSessionWarning.Caption = funLoadResString(10219) & Format(gdteExpireryTime, "Nn:Ss")
    If Format(gdteExpireryTime, "Nn:Ss") = Format(NULL_DATE, "Nn:Ss") Then
        tmrCloseSession.Enabled = False
        Call funStopSession
        Call subBlock
    End If
    Exit Sub

tmrCloseSession_TimerError:
    Call funErrorManager("frmBlock.tmrCloseSession_TimerError", , True)
End Sub


'*******************************************************************************
Private Sub tmrDefaultLanguage_Timer()
Static intNumMins As Integer
On Error GoTo tmrDefaultLanguage_TimerError

'TODO
        
    Exit Sub

tmrDefaultLanguage_TimerError:
    Call funErrorManager("frmBlock.tmrDefaultLanguage_TimerError", , True)
End Sub


'*******************************************************************************
Private Sub tmrNoServCnx_Timer()
Static blnImageOn As Boolean
On Error GoTo tmrNoServCnx_TimerError
    
    If blnImageOn Then
        imgNoServCnx.Picture = imNoServCnx.ListImages(1).Picture
    Else
        imgNoServCnx.Picture = imNoServCnx.ListImages(2).Picture
    End If
    blnImageOn = Not blnImageOn
    Exit Sub

tmrNoServCnx_TimerError:
    Call funErrorManager("frmBlock.tmrNoServCnx_TimerError", Not gblnfrmBlock, True)
    tmrNoServCnx.Enabled = False
End Sub


'*******************************************************************************
Private Sub tmrTopMost_Timer()
On Error Resume Next

    If Not gblnNoTopmost And Me.Visible And Not gblnfrmMsgBox And Not gblnfrmAbout And Not funScreensaverActive Then
        Call SetWindowPos(frmBlock.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    End If
End Sub


'*******************************************************************************
Private Sub tmrUnblock_Timer()
On Error GoTo tmrUnblock_TimerError
        
    gdteExpireryTime = DateAdd("s", -1, gdteExpireryTime)
    If Format(gdteExpireryTime, "Nn:Ss") = Format(NULL_DATE, "Nn:Ss") Then
        tmrUnblock.Enabled = False
        Call subActivate
        Call subBlock
    End If
    Exit Sub

tmrUnblock_TimerError:
    Call funErrorManager("frmBlock.tmrUnblock_TimerError", , True)
End Sub


'*******************************************************************************
Private Sub txtCardNum_Change()
Dim strCardNum As String
Dim strChar As String
Dim i As Integer
On Error GoTo txtCardNum_ChangeError
    
    txtCardNumToolTip.Visible = Not funStripNonNumericChars(txtCardNum)
    cmdCxn.Enabled = txtCardNum.Text <> "" And guCliOpts.uServer.intCnxState = enuSrvCxn.Connected
    Exit Sub

txtCardNum_ChangeError:
    Call funErrorManager("frmBlock.txtCardNum_ChangeError", , True)
End Sub


'*******************************************************************************
Private Sub txtCardNum_GotFocus()
On Error Resume Next
    
    txtCardNum.SelStart = 0
    txtCardNum.SelLength = Len(txtCardNum.Text)
End Sub


'*******************************************************************************
Public Sub subLocalizeInterface()
Dim i As Integer
On Error GoTo subLocalizeInterfaceError

    Call subSetCharset(Me, guCliOpts.intInterfaceLangCharset)
    
    cmdCxn.Caption = funLoadResString(10113)
    cmdCxn.ToolTipText = funLoadResString(10135)
    cmdQuit(0).Caption = funLoadResString(10114)
    cmdQuit(0).ToolTipText = funLoadResString(10136)
    cmdQuit(1).Caption = funLoadResString(10114)
    cmdQuit(1).ToolTipText = funLoadResString(10136)
    cmdReCnx.Caption = funLoadResString(10113)
    cmdReCnx.ToolTipText = funLoadResString(10137)
    cmdUnPause.Caption = funLoadResString(10113)
    cmdUnPause.ToolTipText = funLoadResString(10138)
    fraCardNumber.Caption = funLoadResString(10115)
    lbl(0).Caption = funLoadResString(10298)
    lbl(1).Caption = funLoadResString(10115)
    lbl(2).Caption = funLoadResString(10116)
    lbl(3).Caption = funLoadResString(10117)
    lbl(4).Caption = funLoadResString(10118)
    lbl(5).Caption = funLoadResString(10119)
    lbl(6).Caption = funLoadResString(10120)
    lbl(7).Caption = funLoadResString(10121)
    lbl(8).Caption = funLoadResString(10122)
    lbl(9).Caption = funLoadResString(10123)
    lbl(10).Caption = funLoadResString(10124)
    lbl(11).Caption = funLoadResString(10125)
    lbl(12).Caption = funLoadResString(10123)
    lbl(13).Caption = funLoadResString(10319)
    lbl(14).Caption = funLoadResString(10320)
    lbl(15).Caption = funLoadResString(10321)
'    lbl(16).Caption = funLoadResString(10338)
    'lblConnexionStatus = funLoadResString(10337)
    imgNoServCnx.ToolTipText = funLoadResString(10338) & " - " & funLoadResString(IIf(guCliOpts.uServer.intCnxState = enuSrvCxn.Connected, 10213, 10336))
    lblCloseSessionWarning.Caption = funLoadResString(10219) & Format(SESSION_EXPIRERY_TIME, "Nn:Ss")
    lblAppVersion.Caption = funLoadResString(10000) & " " & App.Major & "." & App.Minor & IIf(BETA_VERSION, "." & App.Revision & " BETA", IIf(RC_VERSION <> "", "." & RC_VERSION, "")) & vbNewLine & MY_COPYRIGHT & " " & BUILD_YEAR & " " & MY_NAME & vbNewLine & MY_URL
    lblGNU.Caption = funLoadResString(10010)
    lblWksnName.Caption = guCliOpts.strWksnName
    lblWksnName.ToolTipText = IIf(guCliOpts.strComputerName <> guCliOpts.strWksnName, "\\" & guCliOpts.strComputerName, "")
    optCard.Caption = funLoadResString(10126)
    optCard.ToolTipText = funLoadResString(10127)
    optSelfService.Caption = funLoadResString(10128)
    optSelfService.ToolTipText = funLoadResString(10129)
    OptClient.Caption = funLoadResString(10318)
    OptClient.ToolTipText = funLoadResString(10330)
    
    optConnected.Caption = funLoadResString(10130)
    optConnected.ToolTipText = funLoadResString(10131)
    optDisconnected.Caption = funLoadResString(10132)
    optDisconnected.ToolTipText = funLoadResString(10133)
    txtCardNum.ToolTipText = funLoadResString(10134)
    txtCardNumToolTip.Text = funLoadResString(10332)
    txtLogin.ToolTipText = funLoadResString(10334)
    txtPassword.ToolTipText = funLoadResString(10335)
    txtLoginToolTip.Text = funLoadResString(10333, CLIENT_LOGIN_FORBIDDEN_CHARS)
    imgLogo.ToolTipText = funLoadResString(10339)
    
    For i = 0 To cmbLanguage.ListCount - 1
        If guCliOpts.strInterfaceLang = mstrLangPacks(0, cmbLanguage.ItemData(i)) Then cmbLanguage.ListIndex = i
    Next i
    If cmbLanguage.ListIndex = -1 Then cmbLanguage.ListIndex = 0

    Exit Sub

subLocalizeInterfaceError:
    Call funErrorManager("frmBlock.subLocalizeInterfaceError", , True)
End Sub


'*******************************************************************************
Private Sub txtCardNum_LostFocus()
On Error Resume Next

    txtCardNumToolTip.Visible = False
End Sub


'*******************************************************************************
Private Sub txtLogin_Change()
On Error Resume Next

    txtLoginToolTip.Visible = Not funStripChars(txtLogin, CLIENT_LOGIN_FORBIDDEN_CHARS)
    cmdCxn.Enabled = txtLogin.Text <> "" And guCliOpts.uServer.intCnxState = enuSrvCxn.Connected
End Sub


'*******************************************************************************
Private Sub txtLogin_GotFocus()
On Error Resume Next
    
    txtLogin.SelStart = 0
    txtLogin.SelLength = Len(txtLogin.Text)
End Sub


'*******************************************************************************
Private Sub txtLogin_LostFocus()
On Error Resume Next

    txtLoginToolTip.Visible = False
End Sub


'*******************************************************************************
Private Sub txtPassword_GotFocus()
On Error Resume Next
    
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
End Sub
