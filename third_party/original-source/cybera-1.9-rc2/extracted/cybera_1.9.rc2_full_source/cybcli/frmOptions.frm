VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CybCli - Options"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   ControlBox      =   0   'False
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSave 
      Caption         =   "_"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   3840
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "_"
      Height          =   375
      Left            =   5340
      TabIndex        =   10
      Top             =   3840
      Width           =   1155
   End
   Begin MSComctlLib.ImageList imlLocked 
      Left            =   600
      Top             =   3840
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
      TabIndex        =   17
      Top             =   3900
      Width           =   4215
      Begin VB.CheckBox chkLocked 
         Height          =   315
         Left            =   0
         Picture         =   "frmOptions.frx":06F6
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   375
      End
      Begin VB.Frame fraAdminPassword 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   420
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   3495
         Begin VB.TextBox txtLockedPassword 
            BorderStyle     =   0  'None
            Height          =   195
            IMEMode         =   3  'DISABLE
            Left            =   1320
            PasswordChar    =   "*"
            TabIndex        =   7
            Top             =   60
            Width           =   1695
         End
         Begin VB.CommandButton cmdUnlockOptions 
            Height          =   315
            Left            =   3120
            Picture         =   "frmOptions.frx":0840
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   0
            Width           =   375
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
         Begin VB.Label lblLockedPassword 
            Alignment       =   1  'Right Justify
            Caption         =   "_"
            Height          =   195
            Left            =   0
            TabIndex        =   19
            Top             =   60
            Width           =   1215
         End
      End
      Begin VB.Label lblLocked 
         Caption         =   "_"
         Height          =   195
         Left            =   420
         TabIndex        =   21
         Top             =   60
         Width           =   3495
      End
      Begin VB.Label lblUnlocked 
         Caption         =   "Les options sont déverouillés"
         Height          =   195
         Left            =   420
         TabIndex        =   20
         Top             =   60
         Visible         =   0   'False
         Width           =   3495
      End
   End
   Begin VB.Frame fraOptions 
      BorderStyle     =   0  'None
      Height          =   4155
      Index           =   1
      Left            =   60
      TabIndex        =   11
      Top             =   60
      Width           =   6435
      Begin VB.Frame fra 
         BorderStyle     =   0  'None
         Height          =   555
         Index           =   1
         Left            =   660
         TabIndex        =   26
         Top             =   2280
         Width           =   5715
         Begin VB.OptionButton optWksnTypeDskless 
            Caption         =   "_"
            Height          =   195
            Left            =   0
            TabIndex        =   28
            Top             =   300
            Width           =   5655
         End
         Begin VB.OptionButton optWksnTypeSndAln 
            Caption         =   "_"
            Height          =   195
            Left            =   0
            TabIndex        =   27
            Top             =   60
            Width           =   5655
         End
      End
      Begin VB.Frame fra 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   6255
         Begin VB.OptionButton optServerAutoConnect 
            Caption         =   "_"
            Height          =   195
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Width           =   6195
         End
         Begin VB.OptionButton optServerManualConnect 
            Caption         =   "_"
            Height          =   195
            Left            =   0
            TabIndex        =   24
            Top             =   300
            Width           =   6195
         End
      End
      Begin VB.CheckBox chkAllowLockdown 
         Caption         =   "_"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   3360
         Width           =   6255
      End
      Begin VB.CheckBox chkStartAtLogon 
         Caption         =   "_"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   2880
         Width           =   6255
      End
      Begin VB.CheckBox chkStartActive 
         Caption         =   "_"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   3120
         Width           =   6255
      End
      Begin VB.TextBox txtWksnName 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   720
         MaxLength       =   20
         TabIndex        =   2
         Top             =   1920
         Width           =   2895
      End
      Begin VB.TextBox txtPortNum 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   5640
         TabIndex        =   1
         Top             =   1020
         Width           =   615
      End
      Begin VB.TextBox txtServerName 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   720
         TabIndex        =   0
         Top             =   1020
         Width           =   2895
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   4
         Left            =   60
         TabIndex        =   22
         Top             =   2340
         Width           =   555
      End
      Begin VB.Shape shp 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   2
         Left            =   660
         Top             =   1860
         Width           =   3015
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   16
         Top             =   1920
         Width           =   555
      End
      Begin VB.Shape shp 
         BorderColor     =   &H80000010&
         Height          =   2175
         Index           =   1
         Left            =   0
         Top             =   1500
         Width           =   6435
      End
      Begin VB.Shape shp 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   0
         Left            =   5580
         Top             =   960
         Width           =   735
      End
      Begin VB.Shape shp 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Index           =   4
         Left            =   660
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   1
         Left            =   4200
         TabIndex        =   14
         Top             =   1020
         Width           =   1335
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   13
         Top             =   1020
         Width           =   555
      End
      Begin VB.Shape shp 
         BorderColor     =   &H80000010&
         Height          =   1395
         Index           =   3
         Left            =   0
         Top             =   0
         Width           =   6435
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         Caption         =   "Serveur"
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
         Index           =   6
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   6435
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         Caption         =   "Client"
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
         TabIndex        =   15
         Top             =   1500
         Width           =   6435
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
'|  Project:       Cybera Workstation (http://cybera.sf.net)                   |
'|  Description:   Options form.                                               |
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

' $Id: frmOptions.frm 15 2006-01-01 16:12:35Z sam $

Option Explicit
Dim mblnSettingsChanged As Boolean
Dim mblnWksnTypeSettingsChanged As Boolean


'*******************************************************************************
Private Sub chkAllowLockdown_Click()

    mblnSettingsChanged = True
End Sub

Private Sub chkLocked_Click()
On Error Resume Next

    fraAdminPassword.Visible = chkLocked.Value = vbChecked
    lblUnlocked.Visible = chkLocked.Value = vbChecked
    lblLocked.Visible = chkLocked.Value = vbUnchecked
    cmdUnlockOptions.Default = chkLocked.Value = vbChecked
    cmdSave.Default = chkLocked.Value = vbUnchecked
    
    If chkLocked.Value = vbChecked Then chkLocked.ToolTipText = funLoadResString(10279)
    If chkLocked.Value = vbChecked Then Call funSetFocus(txtLockedPassword)
    If chkLocked.Value = vbUnchecked Then Call subDisableAllControls
    If chkLocked.Value = vbChecked Then txtLockedPassword.Text = ""
    If chkLocked.Value = vbChecked And guCliOpts.strPasswd = "" Then Call cmdUnlockOptions_Click
End Sub


'*******************************************************************************
Private Sub chkStartActive_Click()
On Error GoTo chkStartActive_ClickError

    mblnSettingsChanged = True
    Exit Sub

chkStartActive_ClickError:
    Call funErrorManager("frmOptions.chkStartActive_ClickError", True)
End Sub


'*******************************************************************************
Private Sub chkStartAtLogon_Click()

    mblnSettingsChanged = True
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
Private Sub cmdSave_Click()
Dim frmTemp As Form
Dim i As Integer
On Error GoTo cmdSave_ClickError

    guCliOpts.uServer.blnAutoConnect = optServerAutoConnect.Value
    guCliOpts.uServer.strServerName = LCase(txtServerName.Text)
    guCliOpts.uServer.lngPort = txtPortNum.Text
    guCliOpts.strWksnName = txtWksnName.Text
    guCliOpts.blnStartActive = chkStartActive.Value = vbChecked
    guCliOpts.blnWksnTypeSndAln = optWksnTypeSndAln.Value
    
    guCliOpts.intStartAtLogon = chkStartAtLogon.Value
    guCliOpts.intAllowLockdown = chkAllowLockdown.Value
    
    Call subSetSystemSettings
    Call frmMain.wskCli.Close
    frmMain.tmrSrvCnxStatus.Enabled = True
    App.Title = funLoadResString(10000)
    
    frmMain.lblServerName = IIf(guCliOpts.uServer.blnAutoConnect, funLoadResString(10310), guCliOpts.uServer.strServerName)
    For Each frmTemp In Forms
        Call frmTemp.subLocalizeInterface
    Next frmTemp
    
    If mblnWksnTypeSettingsChanged Then
        Call funMsg(ShutdownWksnTypeChanged)
        Call Unload(frmMain)
    End If
    Call Unload(Me)
    Exit Sub

cmdSave_ClickError:
    Call funErrorManager("frmOptions.cmdSave_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdUnlockOptions_Click()
On Error Resume Next

    If txtLockedPassword.Text = guCliOpts.strPasswd Then
        Call subEnableAllControls
        fraAdminPassword.Visible = False
        lblUnlocked.Visible = True
        lblLocked.Visible = False
        cmdSave.Default = True
        Call funSetFocus(txtServerName)
    Else
        Call funMsg(enuMsg.BadPasswd)
        txtLockedPassword.Text = ""
        Call funSetFocus(txtLockedPassword)
    End If
    Exit Sub
End Sub


'*******************************************************************************
Private Sub Form_Activate()
On Error Resume Next

    If guCliOpts.strPasswd = "" Then
        
        If guCliOpts.uServer.blnAutoConnect Then
            Call funSetFocus(optServerAutoConnect)
        Else
            Call funSetFocus(txtServerName)
        End If
    Else
        Call funSetFocus(chkLocked)
    End If
End Sub


'*******************************************************************************
Private Sub Form_Load()
Dim i As Integer
On Error GoTo Form_LoadError

    gblnfrmOptions = True
    Call subSetDefaultFont(Me)
    Call subLocalizeInterface
    txtServerName.Text = guCliOpts.uServer.strServerName
    txtPortNum.Text = guCliOpts.uServer.lngPort
    txtWksnName.Text = guCliOpts.strWksnName
    If guCliOpts.blnStartActive Then chkStartActive.Value = vbChecked
    chkStartAtLogon.Value = guCliOpts.intStartAtLogon
    chkAllowLockdown.Value = guCliOpts.intAllowLockdown
    optServerAutoConnect.Value = guCliOpts.uServer.blnAutoConnect
    optServerManualConnect.Value = Not guCliOpts.uServer.blnAutoConnect
    optWksnTypeSndAln.Value = guCliOpts.blnWksnTypeSndAln
    optWksnTypeDskless.Value = Not guCliOpts.blnWksnTypeSndAln
    
    mblnSettingsChanged = False
    mblnWksnTypeSettingsChanged = False
    If guCliOpts.strPasswd = "" Then
        chkLocked.Value = vbChecked
        Call cmdUnlockOptions_Click
    Else
        Call subDisableAllControls
    End If
    If gblnfrmBlock Then
        frmBlock.rtbMessage.FileName = ""
        frmBlock.imgLogo.Picture = LoadPicture
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
    chkLocked.ToolTipText = funLoadResString(10279)
    txtLockedPassword.Text = ""
    For Each uTmpCtrl In frmOptions.Controls
        If uTmpCtrl.Name <> "tbsTabs" And uTmpCtrl.Name <> "chkLocked" And uTmpCtrl.Name <> "txtLockedPassword" And uTmpCtrl.Name <> "cmdUnlockOptions" And uTmpCtrl.Name <> "cmdCancel" And uTmpCtrl.Name <> "fraAdminPassword" And uTmpCtrl.Name <> "fraLocked" And uTmpCtrl.Name <> "lblLockedPassword" And uTmpCtrl.Name <> "lblLocked" Then uTmpCtrl.Enabled = False
    Next uTmpCtrl
End Sub


'*******************************************************************************
Private Sub subEnableAllControls()
Dim uTmpCtrl As Control
On Error Resume Next
    
    chkLocked.Picture = imlLocked.ListImages(2).Picture
    chkLocked.ToolTipText = funLoadResString(10276)
    For Each uTmpCtrl In frmOptions.Controls
        uTmpCtrl.Enabled = True
    Next uTmpCtrl
    
    txtPortNum.Enabled = optServerManualConnect.Value
    txtServerName.Enabled = optServerManualConnect.Value
    lbl(0).Enabled = optServerManualConnect.Value
    lbl(1).Enabled = optServerManualConnect.Value
End Sub


'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)

    gblnfrmOptions = False
End Sub


'*******************************************************************************
Private Sub optServerAutoConnect_Click()
On Error Resume Next
    
    mblnSettingsChanged = True
    txtPortNum.Enabled = False
    txtServerName.Enabled = False
    lbl(0).Enabled = False
    lbl(1).Enabled = False
End Sub


'*******************************************************************************
Private Sub optServerManualConnect_Click()
On Error Resume Next

    mblnSettingsChanged = True
    txtPortNum.Enabled = True
    txtServerName.Enabled = True
    lbl(0).Enabled = True
    lbl(1).Enabled = True
End Sub


'*******************************************************************************
Private Sub optWksnTypeDskless_Click()

    mblnSettingsChanged = True
    mblnWksnTypeSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub optWksnTypeSndAln_Click()

    mblnSettingsChanged = True
    mblnWksnTypeSettingsChanged = True
End Sub


'*******************************************************************************
Private Sub txtWksnName_Change()

    mblnSettingsChanged = True
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
Private Sub txtServerName_Change()
    
    mblnSettingsChanged = True
End Sub


'*******************************************************************************
Public Sub subLocalizeInterface()
On Error GoTo subLocalizeInterfaceError

    Me.Caption = funLoadResString(10000) & " - " & funLoadResString(10107)

    lbl(0).Caption = funLoadResString(10234)
    lbl(1).Caption = funLoadResString(10235)
    lbl(2).Caption = funLoadResString(10234)
    lbl(3).Caption = funLoadResString(10236)
    lbl(4).Caption = funLoadResString(10340)
    lbl(6).Caption = funLoadResString(10239)
'    lbl(13).Caption = funLoadResString(10259)
    
    txtWksnName.ToolTipText = funLoadResString(10247)
    txtServerName.ToolTipText = funLoadResString(10248)
    txtPortNum.ToolTipText = funLoadResString(10249)
    
    chkStartActive.Caption = funLoadResString(10245)
    chkStartActive.ToolTipText = funLoadResString(10246)
    chkStartAtLogon.Caption = funLoadResString(10269)
    chkStartAtLogon.ToolTipText = funLoadResString(10270)
    chkAllowLockdown.Caption = funLoadResString(10273)
    chkAllowLockdown.ToolTipText = funLoadResString(10274)
    
    cmdCancel.Caption = funLoadResString(10241)
    cmdCancel.ToolTipText = funLoadResString(10255)
    cmdSave.Caption = funLoadResString(10243)
    cmdSave.ToolTipText = funLoadResString(10254)
    
    lblLocked.Caption = funLoadResString(10277)
    lblLockedPassword.Caption = funLoadResString(10152)
    txtLockedPassword.ToolTipText = funLoadResString(10278)
    lblUnlocked.Caption = funLoadResString(10275)
    cmdUnlockOptions.ToolTipText = funLoadResString(10279)
    
    optServerAutoConnect.Caption = funLoadResString(10308)
    optServerAutoConnect.ToolTipText = funLoadResString(10311)
    optServerManualConnect.Caption = funLoadResString(10309)
    optServerManualConnect.ToolTipText = funLoadResString(10312)
    
    optWksnTypeSndAln.Caption = funLoadResString(10341)
    optWksnTypeDskless.Caption = funLoadResString(10342)
    optWksnTypeSndAln.ToolTipText = funLoadResString(10343)
    optWksnTypeDskless.ToolTipText = funLoadResString(10344)
    
    Exit Sub
    
subLocalizeInterfaceError:
    Call funErrorManager("frmOptions.subLocalizeInterfaceError", True)
End Sub


'*******************************************************************************
Private Sub txtServerName_GotFocus()
On Error Resume Next
    
    txtServerName.SelStart = 0
    txtServerName.SelLength = Len(txtServerName.Text)
End Sub
