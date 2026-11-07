VERSION 5.00
Begin VB.Form frmPasswdClose 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "_"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   ControlBox      =   0   'False
   Icon            =   "frmPasswdClose.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "_"
      Height          =   375
      Left            =   1980
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "_"
      Default         =   -1  'True
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Top             =   960
      Width           =   1875
   End
   Begin VB.TextBox txtPassword 
      BorderStyle     =   0  'None
      Height          =   195
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   4
      Top             =   600
      Width           =   1395
   End
   Begin VB.Shape shp 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   315
      Left            =   1500
      Top             =   540
      Width           =   1815
   End
   Begin VB.Label lbl 
      Caption         =   "_"
      Height          =   435
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   3255
   End
End
Attribute VB_Name = "frmPasswdClose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ PASSWORD APP CLOSE FORM ____________________|
'|  Name:          frmPasswdClose.frm                                          |
'|  Type:          VB6 Form                                                    |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
'|  Description:   Admin password entry form for App shutdown                  |
'|  Copyright:     (C) 2004-2005 Samuel Monsarrat (smonsarr.dev@free.fr)       |
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

' $Id: frmPasswdClose.frm 4 2005-12-18 17:47:52Z sam $

Option Explicit


'*******************************************************************************
Private Sub cmdCancel_Click()
On Error GoTo cmdCancel_ClickError

    Call Unload(Me)
    Exit Sub

cmdCancel_ClickError:
    Call funErrorManager("frmPasswdClose.cmdCancel_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdQuit_Click()
On Error GoTo cmdQuit_ClickError

    If txtPassword.Text = guServOpts.uConnectedUser.Passwd Then
        Call Unload(Me)
        Call funShutdownApplication
    Else
        Call funMsg(enuMsg.BadPasswd)
        txtPassword.Text = ""
        Call funSetFocus(txtPassword)
    End If
    Exit Sub

cmdQuit_ClickError:
    Call funErrorManager("frmPasswdClose.cmdQuit_ClickError", True)
End Sub


'*******************************************************************************
Private Sub Form_Load()
On Error GoTo Form_LoadError

    Call subSetDefaultFont(Me)
    Call subLocalizeInterface
    Exit Sub

Form_LoadError:
    Call funErrorManager("frmPasswdClose.Form_LoadError", True)
End Sub


'*******************************************************************************
Public Sub subLocalizeInterface()
On Error GoTo subLocalizeInterfaceError

    Call subSetCharset(Me, guServOpts.uConfig.intInterfaceLangCharset)
    
    Me.Caption = funLoadResString(10130)
    lbl(0).Caption = funLoadResString(10117) & LCase(funLoadResString(10130)) & " " & funLoadResString(10000)
    lbl(1).Caption = funLoadResString(10116)
    cmdCancel.Caption = funLoadResString(10123)
    cmdQuit.Caption = funLoadResString(10130)
    Exit Sub

subLocalizeInterfaceError:
    Call funErrorManager("frmPasswdClose.subLocalizeInterfaceError", True)
End Sub

'*******************************************************************************
Private Sub txtPassword_GotFocus()
On Error Resume Next
    
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
End Sub
