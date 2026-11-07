VERSION 5.00
Begin VB.Form frmPasswdChangeUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "_"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   ControlBox      =   0   'False
   Icon            =   "frmPasswdChangeUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdChange 
      Caption         =   "_"
      Default         =   -1  'True
      Height          =   375
      Left            =   60
      TabIndex        =   3
      Top             =   1200
      Width           =   2595
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "_"
      Height          =   375
      Left            =   2700
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtCheckNewPassword 
      BorderStyle     =   0  'None
      Height          =   195
      IMEMode         =   3  'DISABLE
      Left            =   2220
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   840
      Width           =   1755
   End
   Begin VB.TextBox txtNewPassword 
      BorderStyle     =   0  'None
      Height          =   195
      IMEMode         =   3  'DISABLE
      Left            =   2220
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   480
      Width           =   1755
   End
   Begin VB.TextBox txtOldPassword 
      BorderStyle     =   0  'None
      Height          =   195
      IMEMode         =   3  'DISABLE
      Left            =   2220
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   120
      Width           =   1755
   End
   Begin VB.Shape shp 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   315
      Index           =   2
      Left            =   2160
      Top             =   780
      Width           =   1875
   End
   Begin VB.Shape shp 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   315
      Index           =   1
      Left            =   2160
      Top             =   420
      Width           =   1875
   End
   Begin VB.Shape shp 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   315
      Index           =   0
      Left            =   2160
      Top             =   60
      Width           =   1875
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
      Height          =   195
      Index           =   2
      Left            =   60
      TabIndex        =   7
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   6
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmPasswdChangeUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ PASSWORD CHANGE FORM _______________________|
'|  Name:          frmPasswdChangeUser.frm                                     |
'|  Type:          VB6 Form                                                    |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
'|  Description:   User password modification                                  |
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

' $Id: frmPasswdChangeUser.frm 15 2006-01-01 16:12:35Z sam $

Option Explicit


'*******************************************************************************
Private Sub cmdCancel_Click()
On Error GoTo cmdCancel_ClickError
    
    Call Unload(Me)
    Exit Sub

cmdCancel_ClickError:
    Call funErrorManager("frmPasswdChangeUser.cmdCancel_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdChange_Click()
Dim uMsg As clsStatusBarMessage
Dim strBuffer As String
On Error GoTo cmdChange_ClickError

    If txtOldPassword.Text = guServOpts.uConnectedUser.Passwd Then
        If txtNewPassword.Text = txtCheckNewPassword.Text Then
            guServOpts.uConnectedUser.Passwd = txtNewPassword.Text
            Call subDBUpdateUser(guServOpts.uConnectedUser)
            If guServOpts.uConnectedUser.AccessUnblock Then
                strBuffer = CNX_PASSWD_MANGR & guServOpts.uConnectedUser.login & CNX_PAYLOAD_TOKEN & guServOpts.uConnectedUser.Passwd & CNX_MSG_END
            End If
            Call subBroadcastData(strBuffer)
            Call Unload(Me)
        Else
            Call funMsg(enuMsg.BadNewPasswd)
            txtNewPassword.Text = ""
            txtCheckNewPassword.Text = ""
            Call funSetFocus(txtNewPassword)
        End If
    Else
        Call funMsg(enuMsg.BadPasswd)
        txtOldPassword.Text = ""
        Call funSetFocus(txtOldPassword)
    End If
    Exit Sub

cmdChange_ClickError:
    Call funErrorManager("frmPasswdChangeUser.cmdChange_ClickError", True)
End Sub


'*******************************************************************************
Private Sub Form_Load()
On Error GoTo Form_LoadError

    Call subSetDefaultFont(Me)
    Call subLocalizeInterface
    Exit Sub

Form_LoadError:
    Call funErrorManager("frmPasswdChangeUser.Form_LoadError", True)
End Sub


'*******************************************************************************
Public Sub subLocalizeInterface()
On Error GoTo subLocalizeInterfaceError

    Call subSetCharset(Me, guServOpts.uConfig.intInterfaceLangCharset)
    
    Me.Caption = funLoadResString(10784)
    lbl(0).Caption = funLoadResString(10126)
    lbl(1).Caption = funLoadResString(10127)
    lbl(2).Caption = funLoadResString(10128)
    cmdChange.Caption = funLoadResString(10129)
    cmdCancel.Caption = funLoadResString(10123)
    Exit Sub

subLocalizeInterfaceError:
    Call funErrorManager("frmPasswdChangeUser.subLocalizeInterfaceError", True)
End Sub

'*******************************************************************************
Private Sub txtCheckNewPassword_GotFocus()
On Error Resume Next
    
    txtCheckNewPassword.SelStart = 0
    txtCheckNewPassword.SelLength = Len(txtCheckNewPassword.Text)
End Sub


'*******************************************************************************
Private Sub txtNewPassword_GotFocus()
On Error Resume Next
    
    txtNewPassword.SelStart = 0
    txtNewPassword.SelLength = Len(txtNewPassword.Text)
End Sub


'*******************************************************************************
Private Sub txtOldPassword_GotFocus()
On Error Resume Next
    
    txtOldPassword.SelStart = 0
    txtOldPassword.SelLength = Len(txtOldPassword.Text)
End Sub
