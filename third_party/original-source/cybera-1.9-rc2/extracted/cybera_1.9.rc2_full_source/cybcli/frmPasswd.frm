VERSION 5.00
Begin VB.Form frmPasswd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "_"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   ControlBox      =   0   'False
   Icon            =   "frmPasswd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtLogin 
      BorderStyle     =   0  'None
      Height          =   195
      IMEMode         =   3  'DISABLE
      Left            =   1620
      MaxLength       =   255
      TabIndex        =   0
      Top             =   1020
      Width           =   1635
   End
   Begin VB.Timer tmrQuit 
      Interval        =   1000
      Left            =   2700
      Top             =   60
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "_"
      Height          =   375
      Left            =   1980
      TabIndex        =   3
      Top             =   1740
      Width           =   1335
   End
   Begin VB.CommandButton cmdDeactivate 
      Caption         =   "_"
      Default         =   -1  'True
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   1740
      Width           =   1875
   End
   Begin VB.TextBox txtPassword 
      BorderStyle     =   0  'None
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   1620
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1380
      Width           =   1635
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
      Height          =   195
      Index           =   2
      Left            =   0
      TabIndex        =   6
      Top             =   1020
      Width           =   1515
   End
   Begin VB.Shape shp 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   315
      Index           =   0
      Left            =   1560
      Top             =   960
      Width           =   1755
   End
   Begin VB.Shape shp 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   315
      Index           =   1
      Left            =   1560
      Top             =   1320
      Width           =   1755
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   5
      Top             =   1380
      Width           =   1455
   End
   Begin VB.Label lbl 
      Caption         =   "_"
      Height          =   855
      Index           =   0
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   3255
   End
End
Attribute VB_Name = "frmPasswd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ PASSWORD ENTRY FORM ________________________|
'|  Name:          frmPasswd.frm                                               |
'|  Type:          VB6 Form                                                    |
'|  Project:       Cybera Workstation (http://cybera.sf.net)                   |
'|  Description:   Password entry form - Workstation deactivation              |
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

' $Id: frmPasswd.frm 15 2006-01-01 16:12:35Z sam $

Option Explicit
Private mintTimer As Integer


'*******************************************************************************
Private Sub cmdCancel_Click()

    Call Unload(Me)
End Sub


'*******************************************************************************
Private Sub cmdDeactivate_Click()
On Error GoTo cmdDeactivate_ClickError

    If txtLogin.Text <> "" Then
        If (LCase(txtLogin.Text) = LCase(ADMIN_LOGON) And txtPassword.Text = guCliOpts.strPasswd) Or ((LCase(txtLogin.Text) <> LCase(ADMIN_LOGON) And LCase(txtLogin.Text) = LCase(guCliOpts.uConnectedManager.login)) And txtPassword.Text = guCliOpts.uConnectedManager.Passwd) Then
            guCliOpts.uFlags.intNumDeactAttempts = 0
            Call Unload(Me)
            Call subDeactivate
            Call funSetFocus(frmMain)
        Else
            lbl(0).Caption = funLoadResString(10347) & vbNewLine & funLoadResString(10151) & funLoadResString(10000)
            txtPassword.Text = ""
            If guCliOpts.uConnectedManager.login = "" Or guCliOpts.uConnectedManager.login = ADMIN_LOGON Then
                txtLogin.Text = ADMIN_LOGON
                txtLogin.Enabled = False
                Call funSetFocus(txtPassword)
            Else
                txtLogin.Text = ""
                txtLogin.Enabled = True
                Call funSetFocus(txtLogin)
            End If
            guCliOpts.uFlags.intNumDeactAttempts = guCliOpts.uFlags.intNumDeactAttempts + 1
            If guCliOpts.uFlags.intNumDeactAttempts >= 3 Then
                Call funSendData(frmMain.wskCli, CNX_MSG_SEND & funLoadResStringDefault(10304) & guCliOpts.strWksnName & CNX_MSG_END)
                Call Unload(Me)
                Call subLock
                gdteExpireryTime = ACTIVATION_DELAY_TIME
                frmBlock.tmrUnblock.Enabled = True
            Else
                mintTimer = 20
            End If
        End If
    Else
        Call funSetFocus(txtLogin)
    End If
    Exit Sub

cmdDeactivate_ClickError:
    Call funErrorManager("frmPassword.cmdDeactivate_ClickError", True)
End Sub


'*******************************************************************************
Private Sub Form_Activate()
On Error Resume Next

    Call funSetFocus(txtLogin)
End Sub


'*******************************************************************************
Private Sub Form_Load()
On Error GoTo Form_LoadError

    gblnfrmPasswd = True
    mintTimer = 20
    Call subSetDefaultFont(Me)
    Call subLocalizeInterface
    If guCliOpts.uConnectedManager.login = "" Or guCliOpts.uConnectedManager.login = ADMIN_LOGON Then
        txtLogin.Text = ADMIN_LOGON
        txtLogin.Enabled = False
    Else
        txtLogin.Text = ""
        txtLogin.Enabled = True
    End If
    Exit Sub

Form_LoadError:
    Call funErrorManager("frmPassword.Form_LoadError", True)
End Sub


'*******************************************************************************
Public Sub subLocalizeInterface()
On Error GoTo subLocalizeInterfaceError

    Call subSetCharset(Me, guCliOpts.intInterfaceLangCharset)
    
    Me.Caption = funLoadResString(10000) & " - " & funLoadResString(10108)
    cmdCancel.Caption = funLoadResString(10150) & " (" & mintTimer & ")"
    cmdDeactivate.Caption = funLoadResString(10143)
    lbl(0).Caption = funLoadResString(10151, funLoadResString(10000))
    lbl(1).Caption = funLoadResString(10152)
    lbl(2).Caption = funLoadResString(10346)
    Exit Sub

subLocalizeInterfaceError:
    Call funErrorManager("frmPassword.subLocalizeInterfaceError", True)
End Sub


'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    gblnfrmPasswd = False
End Sub

Private Sub tmrQuit_Timer()
On Error Resume Next

    mintTimer = mintTimer - 1
    cmdCancel.Caption = funLoadResString(10150) & " (" & mintTimer & ")"
    If mintTimer = 0 Then cmdCancel.Value = True
End Sub

Private Sub txtLogin_GotFocus()
On Error Resume Next
    
    txtLogin.SelStart = 0
    txtLogin.SelLength = Len(txtLogin.Text)
End Sub

Private Sub txtPassword_GotFocus()
On Error Resume Next
    
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
End Sub
