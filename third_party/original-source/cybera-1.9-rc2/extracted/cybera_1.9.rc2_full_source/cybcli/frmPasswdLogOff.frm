VERSION 5.00
Begin VB.Form frmPasswdLogOff 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "_"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   ControlBox      =   0   'False
   Icon            =   "frmPasswdLogOff.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optShutdown 
      Caption         =   "_"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   3135
   End
   Begin VB.OptionButton optShutdown 
      Caption         =   "_"
      Height          =   195
      Index           =   8
      Left            =   240
      TabIndex        =   2
      Top             =   660
      Width           =   3135
   End
   Begin VB.OptionButton optShutdown 
      Caption         =   "_"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin VB.TextBox txtPassword 
      BorderStyle     =   0  'None
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   1620
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   2460
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "_"
      Height          =   375
      Left            =   2100
      TabIndex        =   5
      Top             =   1380
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "_"
      Default         =   -1  'True
      Height          =   375
      Left            =   60
      TabIndex        =   4
      Top             =   1380
      Width           =   1995
   End
   Begin VB.Shape shp 
      BorderColor     =   &H80000010&
      Height          =   1215
      Index           =   1
      Left            =   60
      Top             =   60
      Width           =   3375
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
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
      Index           =   2
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   3375
   End
   Begin VB.Shape shp 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   315
      Index           =   0
      Left            =   1560
      Top             =   2400
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   7
      Top             =   2460
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblHelp 
      Caption         =   "_"
      Height          =   435
      Left            =   60
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   3375
   End
End
Attribute VB_Name = "frmPasswdLogOff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ PASSWORD ENTRY FORM ________________________|
'|  Name:          frmPasswdLogOff.frm                                         |
'|  Type:          VB6 Form                                                    |
'|  Project:       Cybera Workstation (http://cybera.sf.net)                   |
'|  Description:   Password entry form - Workstation logff/reboot.             |
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

' $Id: frmPasswdLogOff.frm 15 2006-01-01 16:12:35Z sam $

Option Explicit
Private mintSelectedOption As Integer
Private mstrSelectedOptionText As String


'*******************************************************************************
Private Sub cmdCancel_Click()
On Error GoTo cmdCancel_ClickError

    Call Unload(Me)
    Exit Sub

cmdCancel_ClickError:
    Call funErrorManager("frmPasswdLogOff.cmdCancel_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdQuit_Click()
On Error GoTo cmdQuit_ClickError

    
    Call Unload(Me)
    Call subShutdown(mintSelectedOption)
    
'    If txtPassword.Text = guCliOpts.strPasswd Then
'        Call Unload(Me)
'        Call subShutdown(mintSelectedOption)
''        Call subShutdown(mintSelectedOption, LCase(mstrSelectedOptionText))
'    Else
'        Call funMsg(enuMsg.BadPasswd)
'        txtPassword.Text = ""
'        Call funSetFocus(txtPassword)
'    End If
    Exit Sub

cmdQuit_ClickError:
    Call funErrorManager("frmPasswdLogOff.cmdQuit_ClickError", True)
End Sub


'*******************************************************************************
Private Sub Form_Activate()
On Error Resume Next

    Call funSetFocus(txtPassword)
End Sub


'*******************************************************************************
Private Sub Form_Load()
On Error GoTo Form_LoadError
    
    gblnfrmPasswdLogOff = True
    Call subSetDefaultFont(Me)
    Call subLocalizeInterface
    optShutdown(0).Value = True
    Exit Sub

Form_LoadError:
    Call funErrorManager("frmPasswdLogOff.Form_LoadError", True)
End Sub


'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)

    gblnfrmPasswdLogOff = False
End Sub


'*******************************************************************************
Private Sub optShutdown_Click(Index As Integer)
On Error Resume Next

    mintSelectedOption = Index
    mstrSelectedOptionText = optShutdown(Index).Caption
    cmdQuit.Caption = mstrSelectedOptionText
    lblHelp.Caption = funLoadResString(10161) & LCase(mstrSelectedOptionText)
    Call funSetFocus(txtPassword)
End Sub


'*******************************************************************************
Public Sub subLocalizeInterface()
On Error GoTo subLocalizeInterfaceError

    Call subSetCharset(Me, guCliOpts.intInterfaceLangCharset)
    
    Me.Caption = funLoadResString(10000) & " - " & funLoadResString(10111)
    cmdCancel.Caption = funLoadResString(10150)
    lbl(2).Caption = funLoadResString(10158)
    lbl(1).Caption = funLoadResString(10152)
    optShutdown(0).Caption = funLoadResString(10148)
    optShutdown(2).Caption = funLoadResString(10159)
    optShutdown(8).Caption = funLoadResString(10160)
    Exit Sub

subLocalizeInterfaceError:
    Call funErrorManager("frmPasswdLogOff.subLocalizeInterfaceError", True)
End Sub
