VERSION 5.00
Begin VB.Form frmLogon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdLogon 
      Caption         =   "_"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   675
      Left            =   3900
      Picture         =   "frmLogon.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtPasswd 
      BorderStyle     =   0  'None
      Height          =   195
      IMEMode         =   3  'DISABLE
      Left            =   1620
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1140
      Width           =   2115
   End
   Begin VB.TextBox txtLogin 
      BorderStyle     =   0  'None
      Height          =   195
      IMEMode         =   3  'DISABLE
      Left            =   1620
      MaxLength       =   255
      TabIndex        =   0
      Top             =   780
      Width           =   2115
   End
   Begin VB.Label lbl 
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
      Index           =   2
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   4455
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      Height          =   195
      Index           =   3
      Left            =   60
      TabIndex        =   6
      Top             =   360
      Width           =   4455
   End
   Begin VB.Image imgCybServ 
      Height          =   480
      Left            =   4500
      Picture         =   "frmLogon.frx":014A
      Top             =   90
      Width           =   480
   End
   Begin VB.Shape shp 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   315
      Index           =   1
      Left            =   1560
      Top             =   1080
      Width           =   2235
   End
   Begin VB.Shape shp 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   315
      Index           =   0
      Left            =   1560
      Top             =   720
      Width           =   2235
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
      Height          =   195
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   1140
      Width           =   1515
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   780
      Width           =   1515
   End
   Begin VB.Shape shp 
      BackColor       =   &H80000014&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000014&
      Height          =   615
      Index           =   2
      Left            =   0
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "frmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ USER LOGON _________________________________|
'|  Name:          frmLogon.frm                                                |
'|  Type:          VB6 Form                                                    |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
'|  Description:   User logon form.                                            |
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

' $Id: frmLogon.frm 15 2006-01-01 16:12:35Z sam $

Option Explicit


'*******************************************************************************
Private Sub cmdLogon_Click()
Dim uUser As New clsUser
On Error GoTo cmdLogon_ClickError

    uUser.login = LCase(txtLogin.Text)
    uUser.Passwd = txtPasswd.Text
    If funUserLogin(uUser) Then
        Call Unload(Me)
    Else
        Call funMsg(BadPasswd)
        txtPasswd.Text = ""
        Call funSetFocus(txtPasswd)
    End If
    Exit Sub

cmdLogon_ClickError:
    Call funErrorManager("frmLogon.cmdLogon_ClickError", True)
End Sub


'*******************************************************************************
Private Sub Form_Activate()
On Error Resume Next
    
    gblnfrmLogon = True
    If Len(txtLogin.Text) > 0 Then
        Call funSetFocus(txtPasswd)
    Else
        Call funSetFocus(txtLogin)
    End If
End Sub


'*******************************************************************************
Private Sub Form_Load()
On Error GoTo Form_LoadError

    Call subSetDefaultFont(Me)
    Call subLocalizeInterface
    Exit Sub

Form_LoadError:
    Call funErrorManager("frmLogon.Form_LoadError", True)
End Sub


'*******************************************************************************
Public Sub subLocalizeInterface()
On Error GoTo subLocalizeInterfaceError

    Call subSetCharset(Me, guServOpts.uConfig.intInterfaceLangCharset)
    
    Me.Caption = funLoadResString(10000) & " - " & funLoadResString(10774)
    
    cmdLogon.Caption = funLoadResString(10774)
    
    lbl(0).Caption = funLoadResString(10775)
    lbl(1).Caption = funLoadResString(10776)
    lbl(2).Caption = funLoadResString(10782)
    lbl(3).Caption = funLoadResString(10783)
    
    Exit Sub

subLocalizeInterfaceError:
    Call funErrorManager("frmUserAdminConfig.subLocalizeInterfaceError", True)
End Sub


'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    gblnfrmLogon = False
End Sub


'*******************************************************************************
Private Sub txtLogin_Change()
On Error Resume Next
    
    cmdLogon.enabled = Len(txtLogin.Text) > 0
End Sub


'*******************************************************************************
Private Sub txtLogin_GotFocus()
On Error Resume Next
    
    txtLogin.SelStart = 0
    txtLogin.SelLength = Len(txtLogin.Text)
End Sub


'*******************************************************************************
Private Sub txtPasswd_GotFocus()
On Error Resume Next
    
    txtPasswd.SelStart = 0
    txtPasswd.SelLength = Len(txtPasswd.Text)
End Sub

