VERSION 5.00
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
   Begin VB.CommandButton cmdBack 
      Appearance      =   0  'Flat
      Caption         =   "_"
      Height          =   315
      Left            =   3660
      TabIndex        =   0
      Top             =   4920
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "_"
      Height          =   315
      Left            =   6120
      TabIndex        =   3
      Top             =   4920
      Width           =   1155
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "_"
      Height          =   315
      Left            =   4860
      TabIndex        =   2
      Top             =   4920
      Width           =   1155
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "_"
      Default         =   -1  'True
      Height          =   315
      Left            =   4860
      TabIndex        =   1
      Top             =   4920
      Width           =   1155
   End
   Begin VB.Frame fraWizardContainer 
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   7455
      Begin VB.Frame fraStep 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3675
         Index           =   3
         Left            =   60
         TabIndex        =   27
         Top             =   960
         Width           =   7335
         Begin VB.CheckBox chkStartActive 
            Caption         =   "_"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   1380
            Width           =   7155
         End
         Begin VB.CheckBox chkStartAtLogon 
            Caption         =   "_"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   1020
            Width           =   7155
         End
         Begin VB.Label lbl 
            Caption         =   "_"
            Height          =   435
            Index           =   3
            Left            =   420
            TabIndex        =   31
            Top             =   3180
            Width           =   6855
         End
         Begin VB.Image img 
            Height          =   240
            Index           =   1
            Left            =   120
            Picture         =   "frmSetupWizard.frx":014A
            Top             =   3180
            Width           =   300
         End
         Begin VB.Label lblStepIntro 
            Caption         =   "_"
            Height          =   855
            Index           =   3
            Left            =   120
            TabIndex        =   28
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
         TabIndex        =   29
         Top             =   960
         Width           =   7335
         Begin VB.CheckBox chkAllowLockdown 
            Caption         =   "_"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   1020
            Width           =   7155
         End
         Begin VB.Label lbl 
            Caption         =   "_"
            Height          =   615
            Index           =   5
            Left            =   420
            TabIndex        =   33
            Top             =   2520
            Width           =   6855
         End
         Begin VB.Image img 
            Height          =   240
            Index           =   2
            Left            =   120
            Picture         =   "frmSetupWizard.frx":02D4
            Top             =   2520
            Width           =   300
         End
         Begin VB.Image img 
            Height          =   240
            Index           =   0
            Left            =   120
            Picture         =   "frmSetupWizard.frx":045E
            Top             =   3180
            Width           =   300
         End
         Begin VB.Label lbl 
            Caption         =   "_"
            Height          =   435
            Index           =   4
            Left            =   420
            TabIndex        =   32
            Top             =   3180
            Width           =   6855
         End
         Begin VB.Label lblStepIntro 
            Caption         =   "_"
            Height          =   855
            Index           =   4
            Left            =   120
            TabIndex        =   30
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
         Begin VB.CheckBox chkActivateNow 
            Caption         =   "_"
            Height          =   315
            Left            =   120
            TabIndex        =   12
            Top             =   960
            Width           =   7155
         End
         Begin VB.Label lbl 
            Caption         =   "_"
            Height          =   1155
            Index           =   6
            Left            =   120
            TabIndex        =   36
            Top             =   1320
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
         Index           =   2
         Left            =   60
         TabIndex        =   20
         Top             =   960
         Width           =   7335
         Begin VB.TextBox txtWksnName 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   720
            MaxLength       =   20
            TabIndex        =   8
            Top             =   1020
            Width           =   2895
         End
         Begin VB.Label lblStepIntro 
            Caption         =   "_"
            Height          =   855
            Index           =   2
            Left            =   120
            TabIndex        =   26
            Top             =   60
            Width           =   7155
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "_"
            Height          =   195
            Index           =   2
            Left            =   60
            TabIndex        =   25
            Top             =   1020
            Width           =   555
         End
         Begin VB.Shape shp 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000010&
            Height          =   315
            Index           =   3
            Left            =   660
            Top             =   960
            Width           =   3015
         End
      End
      Begin VB.Frame fraStep 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3675
         Index           =   1
         Left            =   60
         TabIndex        =   21
         Top             =   960
         Width           =   7335
         Begin VB.TextBox txtServerName 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   1200
            TabIndex        =   6
            Top             =   1740
            Width           =   2775
         End
         Begin VB.TextBox txtPortNum 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   5820
            TabIndex        =   7
            Top             =   1740
            Width           =   615
         End
         Begin VB.OptionButton optServerAutoConnect 
            Caption         =   "_"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   1020
            Width           =   7155
         End
         Begin VB.OptionButton optServerManualConnect 
            Caption         =   "_"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   1380
            Width           =   7155
         End
         Begin VB.Label lblStepIntro 
            Caption         =   "_"
            Height          =   855
            Index           =   1
            Left            =   120
            TabIndex        =   24
            Top             =   60
            Width           =   6975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "_"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   1740
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "_"
            Height          =   195
            Index           =   1
            Left            =   4200
            TabIndex        =   22
            Top             =   1740
            Width           =   1515
         End
         Begin VB.Shape shp 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000010&
            Height          =   315
            Index           =   4
            Left            =   1140
            Top             =   1680
            Width           =   2895
         End
         Begin VB.Shape shp 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000010&
            Height          =   315
            Index           =   2
            Left            =   5760
            Top             =   1680
            Width           =   735
         End
      End
      Begin VB.Label lblConfigStepDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "_"
         Height          =   375
         Left            =   180
         TabIndex        =   19
         Top             =   360
         Width           =   6495
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   0
         Left            =   6780
         Picture         =   "frmSetupWizard.frx":05E8
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
         TabIndex        =   18
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
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   7455
      Begin VB.Label lblWelcomeHelp 
         BackStyle       =   0  'Transparent
         Caption         =   "_"
         Height          =   195
         Left            =   2640
         TabIndex        =   16
         Top             =   4320
         Width           =   4755
      End
      Begin VB.Label lblWelcomeText 
         BackStyle       =   0  'Transparent
         Caption         =   "_"
         Height          =   1095
         Left            =   2640
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   60
         Width           =   4755
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   1
         Left            =   1740
         Picture         =   "frmSetupWizard.frx":08F2
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
'|  Project:       Cybera Workstation (http://cybera.sf.net)                   |
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

    mintConfigStep = mintConfigStep - 1
    mintConfigStep = IIf(mintConfigStep <= 2 And guCliOpts.uConfigWizard.intSystem >= CONFIG_WIZARD_CLI_SYSTEM, 0, mintConfigStep)
    Call subUpdateSteps
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
    Call subUpdateSteps
    Exit Sub

cmdNext_ClickError:
    Call funErrorManager("frmSetupWizard.cmdNext_ClickError", True)
End Sub


'*******************************************************************************
Private Sub Form_Load()
On Error GoTo Form_LoadError

    gblnfrmSetupWizard = True
    mintConfigStep = IIf(guCliOpts.uConfigWizard.intSystem >= CONFIG_WIZARD_CLI_SYSTEM, 2, 0)
    Call subSetDefaultFont(Me)
    Call subLocalizeInterface
    Call subInitSteps
    Call SetWindowPos(frmMain.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
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
        Call SetWindowPos(frmMain.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    ElseIf funMsg(CancelWizard) = vbNo Then
        Cancel = DO_NOT_SHUTDOWN
    Else
        gblnfrmSetupWizard = False
        Call SetWindowPos(frmMain.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    End If
End Sub


'*******************************************************************************
Private Sub optServerAutoConnect_Click()

    txtPortNum.Enabled = False
    txtServerName.Enabled = False
    lbl(0).Enabled = False
    lbl(1).Enabled = False
End Sub


'*******************************************************************************
Private Sub optServerManualConnect_Click()
On Error Resume Next

    txtPortNum.Enabled = True
    txtServerName.Enabled = True
    lbl(0).Enabled = True
    lbl(1).Enabled = True
End Sub


'*******************************************************************************
Private Sub txtWksnName_GotFocus()
On Error Resume Next
    
    txtWksnName.SelStart = 0
    txtWksnName.SelLength = Len(txtWksnName.Text)
End Sub


'*******************************************************************************
Private Sub txtWksnName_LostFocus()
On Error Resume Next
    
    If InStr(txtWksnName.Text, TIME_TOKEN) > 0 Then
        Call funMsg(InvalidWksnName)
        Call funSetFocus(txtWksnName)
    End If
End Sub


'*******************************************************************************
Private Sub txtPortNum_GotFocus()
On Error Resume Next
    
    txtPortNum.SelStart = 0
    txtPortNum.SelLength = Len(txtPortNum.Text)
End Sub


'*******************************************************************************
Private Sub txtServerName_GotFocus()
On Error Resume Next
    
    txtServerName.SelStart = 0
    txtServerName.SelLength = Len(txtServerName.Text)
End Sub


'*******************************************************************************
Public Sub subLocalizeInterface()
Dim i As Integer
On Error GoTo subLocalizeInterfaceError

    Call subSetCharset(Me, guCliOpts.intInterfaceLangCharset)
    
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
    optServerAutoConnect.Caption = funLoadResString(10308)
    optServerAutoConnect.ToolTipText = funLoadResString(10311)
    optServerManualConnect.Caption = funLoadResString(10309)
    optServerManualConnect.ToolTipText = funLoadResString(10312)
    lbl(0).Caption = funLoadResString(10234)
    lbl(1).Caption = funLoadResString(10235)
    optServerAutoConnect.Value = guCliOpts.uServer.blnAutoConnect
    optServerManualConnect.Value = Not guCliOpts.uServer.blnAutoConnect
    txtServerName.Text = guCliOpts.uServer.strServerName
    txtPortNum.Text = guCliOpts.uServer.lngPort
    
'Step 2
    lbl(2).Caption = funLoadResString(10234)
    txtWksnName.Text = guCliOpts.strWksnName
    
'Step 3
    chkStartActive.Caption = funLoadResString(10245)
    chkStartActive.ToolTipText = funLoadResString(10246)
    chkStartAtLogon.Caption = funLoadResString(10269)
    chkStartAtLogon.ToolTipText = funLoadResString(10270)
    If guCliOpts.blnStartActive Then chkStartActive.Value = vbChecked
    chkStartAtLogon.Value = guCliOpts.intStartAtLogon
    lbl(3).Caption = funLoadResString(10313) & vbNewLine & funLoadResString(10314) & guCliOpts.strUserName
    
'Step 4
    chkAllowLockdown.Caption = funLoadResString(10273)
    chkAllowLockdown.ToolTipText = funLoadResString(10274)
    chkAllowLockdown.Value = guCliOpts.intAllowLockdown
    lbl(4).Caption = funLoadResString(10313) & vbNewLine & funLoadResString(10314) & guCliOpts.strUserName
    lbl(5).Caption = funLoadResString(10315)
    
'Step 5
    chkActivateNow.Caption = funLoadResString(10316)
    lbl(6).Caption = funLoadResString(10331)
    
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
    cmdFinish.Visible = mintConfigStep = 5
    cmdFinish.Default = cmdFinish.Visible
    fraWizardContainer.Visible = mintConfigStep <> 0
    For i = 1 To fraStep.UBound
        fraStep(i).Visible = i = mintConfigStep
    Next i
    Select Case mintConfigStep
    Case 0
        mintConfigStep = IIf(guCliOpts.uConfigWizard.intSystem >= CONFIG_WIZARD_CLI_SYSTEM, 2, 0)
    
    Case 1
        lblConfigStepTitle = funLoadResString(10023)
        lblConfigStepDescription = funLoadResString(10024)
        Call optServerAutoConnect.SetFocus
    
    Case 2
        lblConfigStepTitle = funLoadResString(10026)
        lblConfigStepDescription = funLoadResString(10027)
        Call txtWksnName.SetFocus
    
    Case 3
        lblConfigStepTitle = funLoadResString(10029)
        lblConfigStepDescription = funLoadResString(10030)
        Call chkStartAtLogon.SetFocus
        
    Case 4
        lblConfigStepTitle = funLoadResString(10032)
        lblConfigStepDescription = funLoadResString(10033)
        Call chkAllowLockdown.SetFocus
    
    Case 5
        lblConfigStepTitle = funLoadResString(10036)
        lblConfigStepDescription = funLoadResString(10037)
        chkActivateNow.Value = chkStartActive.Value
        Call chkActivateNow.SetFocus
    
    End Select
    Exit Sub

subUpdateStepsError:
    Call funErrorManager("frmSetupWizard.subUpdateStepsError", True)
End Sub


'*******************************************************************************
Private Sub subFinish()
Dim blnActivateNow As Boolean
On Error GoTo subFinishError
    
    guCliOpts.uServer.blnAutoConnect = optServerAutoConnect.Value
    guCliOpts.uServer.strServerName = LCase(txtServerName.Text)
    guCliOpts.uServer.lngPort = txtPortNum.Text
    guCliOpts.strWksnName = txtWksnName.Text
    If chkStartActive.Value = vbChecked Then
        guCliOpts.blnStartActive = True
    Else
        guCliOpts.blnStartActive = False
    End If
    guCliOpts.intStartAtLogon = chkStartAtLogon.Value
    guCliOpts.intAllowLockdown = chkAllowLockdown.Value
    
    Call subSetSystemSettings
    Call frmMain.wskCli.Close
    frmMain.tmrSrvCnxStatus.Enabled = True
    frmMain.lblServerName = IIf(guCliOpts.uServer.blnAutoConnect, funLoadResString(10310), guCliOpts.uServer.strServerName)
    blnActivateNow = chkActivateNow.Value = vbChecked
    guCliOpts.uConfigWizard.intSystem = CONFIG_WIZARD_CLI_SYSTEM
    Call SetRegSettings(funLoadResString(10006), "General", "Config Wizard", guCliOpts.uConfigWizard.intSystem)
    guCliOpts.uConfigWizard.intUser = CONFIG_WIZARD_CLI_USER
    Call SetRegUserSettings(funLoadResString(10006), "General", "Config Wizard", guCliOpts.uConfigWizard.intUser)
    mblnForceQuit = True
    Call Unload(Me)
    If blnActivateNow Then Call subActivate
    Exit Sub

subFinishError:
    Call funErrorManager("frmSetupWizard.subFinishError", True)
End Sub

