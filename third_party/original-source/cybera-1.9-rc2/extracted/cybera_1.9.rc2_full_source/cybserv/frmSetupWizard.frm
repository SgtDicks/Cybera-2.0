VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSetupWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "_"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   Icon            =   "frmSetupWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   7455
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "_"
      Height          =   315
      Left            =   6120
      TabIndex        =   2
      Top             =   4920
      Width           =   1155
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "_"
      Default         =   -1  'True
      Height          =   315
      Left            =   4860
      TabIndex        =   0
      Top             =   4920
      Width           =   1155
   End
   Begin VB.CommandButton cmdBack 
      Appearance      =   0  'Flat
      Caption         =   "_"
      Height          =   315
      Left            =   3660
      TabIndex        =   1
      Top             =   4920
      Width           =   1155
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "_"
      Height          =   315
      Left            =   4860
      TabIndex        =   17
      Top             =   4920
      Width           =   1155
   End
   Begin VB.Frame fraWizardContainer 
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   7455
      Begin VB.Frame fraStep 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3675
         Index           =   2
         Left            =   60
         TabIndex        =   25
         Top             =   960
         Width           =   7335
         Begin VB.ComboBox cmbRechargeValidityType 
            Height          =   315
            Left            =   4260
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1740
            Width           =   2715
         End
         Begin VB.TextBox txtRechargePrice 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   3720
            MaxLength       =   7
            TabIndex        =   6
            Top             =   1440
            Width           =   915
         End
         Begin VB.TextBox txtRechargeCaption 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   3720
            MaxLength       =   50
            TabIndex        =   9
            Top             =   2160
            Width           =   3195
         End
         Begin MSComCtl2.DTPicker dtpRechargeTime 
            Height          =   315
            Left            =   3660
            TabIndex        =   5
            Top             =   1020
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:mm:ss"
            Format          =   52363267
            UpDown          =   -1  'True
            CurrentDate     =   38353
            MaxDate         =   38717
            MinDate         =   38353
         End
         Begin MSComCtl2.DTPicker dtpRechargeValidity 
            Height          =   315
            Left            =   3660
            TabIndex        =   7
            Top             =   1740
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yy"
            Format          =   52363267
            UpDown          =   -1  'True
            CurrentDate     =   36892
            MaxDate         =   47484
            MinDate         =   36892
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "_"
            Height          =   195
            Index           =   8
            Left            =   60
            TabIndex        =   39
            Top             =   1800
            Width           =   3555
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "_"
            Height          =   195
            Index           =   6
            Left            =   60
            TabIndex        =   38
            Top             =   1440
            Width           =   3555
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "_"
            Height          =   195
            Index           =   7
            Left            =   60
            TabIndex        =   37
            Top             =   2160
            Width           =   3555
         End
         Begin VB.Shape shp 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000010&
            Height          =   315
            Index           =   12
            Left            =   3660
            Top             =   1380
            Width           =   1035
         End
         Begin VB.Shape shp 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000010&
            Height          =   315
            Index           =   14
            Left            =   3660
            Top             =   2100
            Width           =   3315
         End
         Begin VB.Label lblStepIntro 
            Caption         =   "_"
            Height          =   855
            Index           =   2
            Left            =   120
            TabIndex        =   29
            Top             =   60
            Width           =   7155
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "_"
            Height          =   195
            Index           =   2
            Left            =   60
            TabIndex        =   28
            Top             =   1080
            Width           =   3555
         End
      End
      Begin VB.Frame fraStep 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3675
         Index           =   1
         Left            =   60
         TabIndex        =   26
         Top             =   960
         Width           =   7335
         Begin MSComCtl2.DTPicker dtpNumWorkstations 
            Height          =   315
            Left            =   3480
            TabIndex        =   3
            Top             =   960
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "y"
            Format          =   52363267
            UpDown          =   -1  'True
            CurrentDate     =   36892
            MaxDate         =   55153
            MinDate         =   36892
         End
         Begin VB.CheckBox chkWksnLockdown 
            Caption         =   "_"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   1440
            Width           =   7155
         End
         Begin VB.Label lbl 
            Caption         =   "_"
            Height          =   615
            Index           =   13
            Left            =   300
            TabIndex        =   47
            Top             =   3000
            Width           =   6975
         End
         Begin VB.Image img 
            Height          =   240
            Index           =   0
            Left            =   0
            Picture         =   "frmSetupWizard.frx":014A
            Top             =   3000
            Width           =   300
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "_"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   36
            Top             =   1020
            Width           =   3255
         End
         Begin VB.Label lblStepIntro 
            Caption         =   "_"
            Height          =   855
            Index           =   1
            Left            =   120
            TabIndex        =   27
            Top             =   60
            Width           =   7155
         End
      End
      Begin VB.Frame fraStep 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3675
         Index           =   5
         Left            =   60
         TabIndex        =   34
         Top             =   960
         Width           =   7335
         Begin VB.CheckBox chkOptionsNow 
            Caption         =   "_"
            Height          =   435
            Left            =   120
            TabIndex        =   16
            Top             =   960
            Width           =   7155
         End
         Begin VB.Label lblStepIntro 
            Caption         =   "_"
            Height          =   855
            Index           =   5
            Left            =   120
            TabIndex        =   35
            Top             =   60
            Width           =   7155
         End
      End
      Begin VB.Frame fraStep 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3675
         Index           =   4
         Left            =   60
         TabIndex        =   32
         Top             =   960
         Width           =   7335
         Begin VB.TextBox txtUserPasswdCheck 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   195
            IMEMode         =   3  'DISABLE
            Left            =   3660
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   15
            Top             =   2160
            Width           =   1755
         End
         Begin VB.TextBox txtUserPasswd 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   195
            IMEMode         =   3  'DISABLE
            Left            =   3660
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   14
            Top             =   1800
            Width           =   1755
         End
         Begin VB.TextBox txtLogin 
            BorderStyle     =   0  'None
            Height          =   195
            IMEMode         =   3  'DISABLE
            Left            =   3660
            MaxLength       =   20
            TabIndex        =   12
            Top             =   1080
            Width           =   1755
         End
         Begin VB.TextBox txtFullName 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   195
            IMEMode         =   3  'DISABLE
            Left            =   3660
            MaxLength       =   50
            TabIndex        =   13
            Top             =   1440
            Width           =   1755
         End
         Begin VB.Shape shp 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000010&
            Height          =   315
            Index           =   7
            Left            =   3600
            Top             =   2100
            Width           =   1875
         End
         Begin VB.Shape shp 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000010&
            Height          =   315
            Index           =   6
            Left            =   3600
            Top             =   1740
            Width           =   1875
         End
         Begin VB.Shape shp 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000010&
            Height          =   315
            Index           =   5
            Left            =   3600
            Top             =   1020
            Width           =   1875
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "_"
            Enabled         =   0   'False
            Height          =   195
            Index           =   12
            Left            =   60
            TabIndex        =   46
            Top             =   2160
            Width           =   3495
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "_"
            Enabled         =   0   'False
            Height          =   195
            Index           =   11
            Left            =   60
            TabIndex        =   45
            Top             =   1800
            Width           =   3495
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "_"
            Height          =   195
            Index           =   5
            Left            =   60
            TabIndex        =   44
            Top             =   1080
            Width           =   3495
         End
         Begin VB.Shape shp 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000010&
            Height          =   315
            Index           =   4
            Left            =   3600
            Top             =   1380
            Width           =   1875
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "_"
            Enabled         =   0   'False
            Height          =   195
            Index           =   4
            Left            =   60
            TabIndex        =   43
            Top             =   1440
            Width           =   3495
         End
         Begin VB.Label lblStepIntro 
            Caption         =   "_"
            Height          =   855
            Index           =   4
            Left            =   120
            TabIndex        =   33
            Top             =   60
            Width           =   7155
         End
      End
      Begin VB.Frame fraStep 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3675
         Index           =   3
         Left            =   60
         TabIndex        =   30
         Top             =   960
         Width           =   7335
         Begin VB.TextBox txtNewPassword 
            BorderStyle     =   0  'None
            Height          =   195
            IMEMode         =   3  'DISABLE
            Left            =   3660
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   10
            Top             =   1080
            Width           =   1755
         End
         Begin VB.TextBox txtCheckNewPassword 
            BorderStyle     =   0  'None
            Height          =   195
            IMEMode         =   3  'DISABLE
            Left            =   3660
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   11
            Top             =   1440
            Width           =   1755
         End
         Begin VB.Image img 
            Height          =   240
            Index           =   1
            Left            =   0
            Picture         =   "frmSetupWizard.frx":02D4
            Top             =   3000
            Width           =   300
         End
         Begin VB.Label lbl 
            Caption         =   "_"
            Height          =   615
            Index           =   10
            Left            =   300
            TabIndex        =   42
            Top             =   3000
            Width           =   6975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "_"
            Height          =   195
            Index           =   9
            Left            =   60
            TabIndex        =   41
            Top             =   1080
            Width           =   3495
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "_"
            Height          =   195
            Index           =   3
            Left            =   60
            TabIndex        =   40
            Top             =   1440
            Width           =   3495
         End
         Begin VB.Shape shp 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000010&
            Height          =   315
            Index           =   3
            Left            =   3600
            Top             =   1020
            Width           =   1875
         End
         Begin VB.Shape shp 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000010&
            Height          =   315
            Index           =   2
            Left            =   3600
            Top             =   1380
            Width           =   1875
         End
         Begin VB.Label lblStepIntro 
            Caption         =   "_"
            Height          =   855
            Index           =   3
            Left            =   120
            TabIndex        =   31
            Top             =   60
            Width           =   7155
         End
      End
      Begin VB.Label lblConfigStepDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "_"
         Height          =   375
         Left            =   180
         TabIndex        =   24
         Top             =   360
         Width           =   6495
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   0
         Left            =   6780
         Picture         =   "frmSetupWizard.frx":045E
         Top             =   180
         Width           =   480
      End
      Begin VB.Label lblConfigStepTitle 
         BackStyle       =   0  'Transparent
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
         TabIndex        =   23
         Top             =   60
         Width           =   6615
      End
      Begin VB.Shape shp 
         BackColor       =   &H80000014&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000014&
         Height          =   855
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   7455
      End
   End
   Begin VB.Frame fraWelcome 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   7455
      Begin VB.Label lblWelcomeHelp 
         BackStyle       =   0  'Transparent
         Caption         =   "_"
         Height          =   195
         Left            =   2640
         TabIndex        =   21
         Top             =   4320
         Width           =   4755
      End
      Begin VB.Label lblWelcomeText 
         BackStyle       =   0  'Transparent
         Caption         =   "_"
         Height          =   1095
         Left            =   2640
         TabIndex        =   20
         Top             =   1320
         Width           =   4755
      End
      Begin VB.Label lblWelcomeTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   2640
         TabIndex        =   19
         Top             =   60
         Width           =   4755
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   1
         Left            =   1740
         Picture         =   "frmSetupWizard.frx":0768
         Top             =   240
         Width           =   480
      End
      Begin VB.Shape shp 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000002&
         Height          =   4695
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   2475
      End
   End
End
Attribute VB_Name = "frmSetupWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ SETUP WIZARD _______________________________|
'|  Name:          frmSetupWizard.frm                                          |
'|  Type:          VB6 Form                                                    |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
'|  Description:   Settings wizard used on first use of app.                   |
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
'|_____________________________________________________________________________|

' $Id: frmSetupWizard.frm 15 2006-01-01 16:12:35Z sam $

Option Explicit
Dim mintConfigStep As Integer
Dim mblnForceQuit As Boolean


'*******************************************************************************
Private Sub cmdBack_Click()
On Error GoTo cmdBack_ClickError

    mintConfigStep = mintConfigStep - 1
    mintConfigStep = IIf(mintConfigStep = 4 And Not (grstUsers.BOF And grstUsers.EOF), mintConfigStep - 1, mintConfigStep)
    mintConfigStep = IIf(mintConfigStep = 2 And Not (grstRecharges.BOF And grstRecharges.EOF), mintConfigStep - 1, mintConfigStep)
    Call subUpdateSteps
    Exit Sub

cmdBack_ClickError:
    Call funErrorManager("frmSetupWizard.cmdBack_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdCancel_Click()
On Error GoTo cmdCancel_ClickError

    Call Unload(Me)
    Exit Sub

cmdCancel_ClickError:
    Call funErrorManager("frmSetupWizard.cmdCancel_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdFinish_Click()
On Error GoTo cmdFinish_ClickError

    Call subFinish
    Exit Sub

cmdFinish_ClickError:
    Call funErrorManager("frmSetupWizard.cmdFinish_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdNext_Click()
On Error GoTo cmdNext_ClickError

    mintConfigStep = mintConfigStep + 1
    mintConfigStep = IIf(mintConfigStep = 2 And Not (grstRecharges.BOF And grstRecharges.EOF), mintConfigStep + 1, mintConfigStep)
    mintConfigStep = IIf(mintConfigStep = 4 And Not (grstUsers.BOF And grstUsers.EOF), mintConfigStep + 1, mintConfigStep)
    Call subUpdateSteps
    Exit Sub

cmdNext_ClickError:
    Call funErrorManager("frmSetupWizard.cmdNext_ClickError", True)
End Sub


'*******************************************************************************
Private Sub dtpRechargeTime_Change()
On Error Resume Next

    cmdNext.enabled = dtpRechargeTime.value <> NULL_DATE
End Sub


'*******************************************************************************
Private Sub Form_Activate()
On Error Resume Next

    Call cmdNext.SetFocus
End Sub


'*******************************************************************************
Private Sub Form_Load()
On Error GoTo Form_LoadError

    gblnfrmSetupWizard = True
    Call subSetDefaultFont(Me)
    Call subLocalizeInterface
    Call subInitSteps
    cmdBack.Visible = False
    cmdFinish.Visible = False
    fraWelcome.Visible = True
    fraWizardContainer.Visible = False
    mintConfigStep = 0
    Exit Sub

Form_LoadError:
    Call funErrorManager("frmSetupWizard.Form_LoadError", True)
End Sub


'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    If mblnForceQuit Then
        gblnfrmSetupWizard = False
    ElseIf funMsg(CancelWizard) = vbNo Then
        Cancel = DO_NOT_SHUTDOWN
    Else
        gblnfrmSetupWizard = False
    End If
End Sub


'*******************************************************************************
Public Sub subLocalizeInterface()
Dim i As Integer
On Error GoTo subLocalizeInterfaceError

    Call subSetCharset(Me, guServOpts.uConfig.intInterfaceLangCharset)
    
    Me.Caption = funLoadResString(10000) & " - " & funLoadResString(10015)
    lblWelcomeTitle.Caption = funLoadResString(10016)
    lblWelcomeText.Caption = funLoadResString(10017)
    lblWelcomeHelp.Caption = funLoadResString(10018)
    cmdBack.Caption = funLoadResString(10019)
    cmdNext.Caption = funLoadResString(10020)
    cmdCancel.Caption = funLoadResString(10021)
    cmdFinish.Caption = funLoadResString(10035)
    
    lblStepIntro(1).Caption = funLoadResString(10025)
    lblStepIntro(2).Caption = funLoadResString(10028)
    lblStepIntro(3).Caption = funLoadResString(10031)
    lblStepIntro(4).Caption = funLoadResString(10034)
    lblStepIntro(5).Caption = funLoadResString(10038)
    
    Exit Sub

subLocalizeInterfaceError:
    Call funErrorManager("frmSetupWizard.subLocalizeInterfaceError", True)
End Sub


'*******************************************************************************
Private Sub subInitSteps()
On Error GoTo subInitStepsError

'Step 1
    lbl(0).Caption = funLoadResString(10224)
    lbl(13).Caption = funLoadResString(10790)
    dtpNumWorkstations.ToolTipText = funLoadResString(10499)
    chkWksnLockdown.Caption = funLoadResString(10584)
    chkWksnLockdown.ToolTipText = funLoadResString(10737)
    dtpNumWorkstations.value = DateAdd("yyyy", guServOpts.uWorkstations.lngNumWorkstations - 1, dtpNumWorkstations.MinDate)
    chkWksnLockdown.value = guServOpts.uWorkstations.intModifyLockdown

'Step 2
    lbl(2).Caption = funLoadResString(10228)
    lbl(6).Caption = funLoadResString(10229)
    lbl(8).Caption = funLoadResString(10231)
    lbl(7).Caption = funLoadResString(10230)
    dtpRechargeTime.MinDate = NULL_DATE
    dtpRechargeTime.value = NULL_DATE
    txtRechargePrice.Text = 0
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
    
'Step 3
    lbl(10).Caption = funLoadResString(10785)
    lbl(9).Caption = funLoadResString(10786)
    lbl(3).Caption = funLoadResString(10787)
    txtNewPassword.Text = guServOpts.uConfig.strPasswd
    txtCheckNewPassword.Text = guServOpts.uConfig.strPasswd

'Step 4
    lbl(5).Caption = funLoadResString(10766)
    lbl(4).Caption = funLoadResString(10767)
    lbl(11).Caption = funLoadResString(10786)
    lbl(12).Caption = funLoadResString(10787)
    
'Step 5
    chkOptionsNow.Caption = funLoadResString(10788)

    Exit Sub

subInitStepsError:
    Call funErrorManager("frmSetupWizard.subInitStepsError", True)
End Sub


'*******************************************************************************
Private Sub subUpdateSteps()
Dim i As Integer
On Error GoTo subUpdateStepsError

    fraWelcome.Visible = mintConfigStep = 0
    cmdBack.Visible = mintConfigStep <> 0
    cmdNext.Visible = mintConfigStep <> 5
    cmdNext.Default = cmdNext.Visible
    cmdNext.enabled = True
    cmdFinish.Visible = mintConfigStep = 5
    cmdFinish.Default = cmdFinish.Visible
    fraWizardContainer.Visible = mintConfigStep <> 0
    For i = 1 To fraStep.UBound
        fraStep(i).Visible = i = mintConfigStep
    Next i
    Select Case mintConfigStep
    Case 0
        Call cmdNext.SetFocus
    
    Case 1
        lblConfigStepTitle = funLoadResString(10023)
        lblConfigStepDescription = funLoadResString(10024)
        Call funSetFocus(chkWksnLockdown)
        DoEvents
        Call funSetFocus(dtpNumWorkstations)
    
    Case 2
        lblConfigStepTitle = funLoadResString(10026)
        lblConfigStepDescription = funLoadResString(10027)
        cmdNext.enabled = dtpRechargeTime.value <> NULL_DATE
        Call funSetFocus(dtpRechargeTime)
    
    Case 3
        lblConfigStepTitle = funLoadResString(10029)
        lblConfigStepDescription = funLoadResString(10030)
        Call funSetFocus(txtNewPassword)
        
    Case 4
        lblConfigStepTitle = funLoadResString(10032)
        lblConfigStepDescription = funLoadResString(10033)
        Call funSetFocus(txtLogin)
    
    Case 5
        lblConfigStepTitle = funLoadResString(10036)
        lblConfigStepDescription = funLoadResString(10037)
        Call funSetFocus(chkOptionsNow)
    
    End Select
    Exit Sub

subUpdateStepsError:
    Call funErrorManager("frmSetupWizard.subUpdateStepsError", True)
End Sub



'*******************************************************************************
Private Sub subFinish()
Dim blnOptionsNow As Boolean
Dim uUser As New clsUser
Dim uWorkstation As clsWorkstation
On Error GoTo subFinishError
    
    guServOpts.uWorkstations.lngNumWorkstations = CLng(format(dtpNumWorkstations.value, "yy"))
    guServOpts.uWorkstations.intModifyLockdown = chkWksnLockdown.value
    If chkWksnLockdown.value = vbChecked Then
        guServOpts.uWorkstations.uLocks.intLockDesktop = vbChecked
        guServOpts.uWorkstations.uLocks.intLockIE = vbChecked
        guServOpts.uWorkstations.uLocks.intNoAdminTools = vbChecked
        guServOpts.uWorkstations.uLocks.intNoClose = vbChecked
        guServOpts.uWorkstations.uLocks.intNoCommonGroup = vbUnchecked
        guServOpts.uWorkstations.uLocks.intNoCPL = vbChecked
        guServOpts.uWorkstations.uLocks.intNoFileSystemAccess = vbChecked
    End If
    
    If dtpRechargeTime.value <> NULL_DATE Then
        Call grstRecharges.AddNew
        grstRecharges!Time = TimeValue(dtpRechargeTime.value)
        grstRecharges!Price = txtRechargePrice.Text
        grstRecharges!validity = CInt(format(dtpRechargeValidity.value, "yy"))
        grstRecharges!validity_type = cmbRechargeValidityType.ItemData(cmbRechargeValidityType.ListIndex)
        grstRecharges!Caption = txtRechargeCaption.Text
        grstRecharges!type_postpaid_ss = True
        grstRecharges!type_prepaid_ss = True
        grstRecharges!type_card = True
        grstRecharges!type_client = True
        grstRecharges!usage_order = 1
        Call grstRecharges.Update
    End If
    
    guServOpts.uConfig.strPasswd = txtNewPassword.Text
    Call SetRegSettings(funLoadResString(10006), "General", "Password", funDoXOR(guServOpts.uConfig.strPasswd, funLoadResString(10006)))
    
    If Len(txtLogin.Text) > 0 Then
        uUser.login = txtLogin.Text
        uUser.FullName = txtFullName.Text
        uUser.Passwd = txtUserPasswd.Text
        uUser.AccessJournals = True
        uUser.AccessUnblock = True
        uUser.AccessDelRecharge = True
        Call funDBAddUser(uUser)
        Call subRemoveStatusBarMsg("NO_USERS")
        Call SetRegUserSettings(funLoadResString(10006), "General", "Last user", uUser.login)
    End If
    
    Call subSetSystemSettings
    
    Call subInitWorkstations
    Call subInitSessions
    
    For Each uWorkstation In guWorkstations
        Call subWorkstationInit(uWorkstation, False)
        Call subWorkstationLockdown(uWorkstation)
        DoEvents
    Next uWorkstation
    
    blnOptionsNow = chkOptionsNow.value = vbChecked
    guServOpts.uConfigWizard.intSystem = CONFIG_WIZARD_SERV_SYSTEM
    Call SetRegSettings(funLoadResString(10006), "General", "Config Wizard", guServOpts.uConfigWizard.intSystem)
    guServOpts.uConfigWizard.intUser = CONFIG_WIZARD_SERV_USER
    Call SetRegUserSettings(funLoadResString(10006), "General", "Config Wizard", guServOpts.uConfigWizard.intUser)
    mblnForceQuit = True
    Call Unload(Me)
    If blnOptionsNow Then
        Call frmOptions.Show(vbModal, frmMain)
    Else
        Call funUserLogoff(uUser, False)
    End If
    
    Exit Sub

subFinishError:
    Call funErrorManager("frmSetupWizard.subFinishError", True)
End Sub


'*******************************************************************************
Private Sub txtCheckNewPassword_Change()
On Error Resume Next

    cmdNext.enabled = txtNewPassword.Text = txtCheckNewPassword.Text
    cmdBack.enabled = txtNewPassword.Text = txtCheckNewPassword.Text
End Sub


'*******************************************************************************
Private Sub txtCheckNewPassword_GotFocus()
On Error Resume Next
    
    txtCheckNewPassword.SelStart = 0
    txtCheckNewPassword.SelLength = Len(txtCheckNewPassword.Text)
End Sub


'*******************************************************************************
Private Sub txtFullName_GotFocus()
On Error Resume Next
    
    txtFullName.SelStart = 0
    txtFullName.SelLength = Len(txtFullName.Text)
End Sub


'*******************************************************************************
Private Sub txtLogin_Change()
On Error Resume Next

    txtFullName.enabled = Len(txtLogin.Text) > 0
    txtUserPasswd.enabled = Len(txtLogin.Text) > 0
    txtUserPasswdCheck.enabled = Len(txtLogin.Text) > 0
    lbl(4).enabled = Len(txtLogin.Text) > 0
    lbl(11).enabled = Len(txtLogin.Text) > 0
    lbl(12).enabled = Len(txtLogin.Text) > 0
End Sub


'*******************************************************************************
Private Sub txtLogin_GotFocus()
On Error Resume Next
    
    txtLogin.SelStart = 0
    txtLogin.SelLength = Len(txtLogin.Text)
End Sub


'*******************************************************************************
Private Sub txtNewPassword_Change()
On Error Resume Next

    cmdNext.enabled = txtNewPassword.Text = txtCheckNewPassword.Text
    cmdBack.enabled = txtNewPassword.Text = txtCheckNewPassword.Text
    img(1).Visible = Len(txtNewPassword.Text) = 0
    lbl(10).Visible = Len(txtNewPassword.Text) = 0
End Sub


'*******************************************************************************
Private Sub txtNewPassword_GotFocus()
On Error Resume Next
    
    txtNewPassword.SelStart = 0
    txtNewPassword.SelLength = Len(txtNewPassword.Text)
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
Private Sub txtUserPasswd_Change()
On Error Resume Next

    cmdNext.enabled = txtUserPasswd.Text = txtUserPasswdCheck.Text
    cmdBack.enabled = txtUserPasswd.Text = txtUserPasswdCheck.Text
End Sub


'*******************************************************************************
Private Sub txtUserPasswd_GotFocus()
On Error Resume Next
    
    txtUserPasswd.SelStart = 0
    txtUserPasswd.SelLength = Len(txtUserPasswd.Text)
End Sub


'*******************************************************************************
Private Sub txtUserPasswdCheck_Change()
On Error Resume Next

    cmdNext.enabled = txtUserPasswd.Text = txtUserPasswdCheck.Text
    cmdBack.enabled = txtUserPasswd.Text = txtUserPasswdCheck.Text
End Sub


'*******************************************************************************
Private Sub txtUserPasswdCheck_GotFocus()
On Error Resume Next
    
    txtUserPasswdCheck.SelStart = 0
    txtUserPasswdCheck.SelLength = Len(txtUserPasswdCheck.Text)
End Sub
