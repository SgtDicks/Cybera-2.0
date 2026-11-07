VERSION 5.00
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "sysinfo.ocx"
Begin VB.Form frmMsgBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "_"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   ControlBox      =   0   'False
   Icon            =   "frmMsgBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTimeOut 
      Interval        =   65535
      Left            =   4320
      Top             =   900
   End
   Begin SysInfoLib.SysInfo sysInfo 
      Left            =   60
      Top             =   900
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame fraYesNo 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton cmdYes 
         Caption         =   "_"
         Height          =   315
         Left            =   1260
         TabIndex        =   5
         Top             =   0
         Width           =   1155
      End
      Begin VB.CommandButton cmdNo 
         Caption         =   "_"
         Default         =   -1  'True
         Height          =   315
         Left            =   2520
         TabIndex        =   4
         Top             =   0
         Width           =   1155
      End
   End
   Begin VB.Frame fraOK 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton cmdOK 
         Caption         =   "_"
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.Shape shp 
      BorderColor     =   &H80000010&
      Height          =   915
      Index           =   1
      Left            =   60
      Top             =   60
      Width           =   4815
   End
   Begin VB.Label lblText 
      Height          =   795
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   120
      Picture         =   "frmMsgBox.frx":0442
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ USER MESSAGE BOX ___________________________|
'|  Name:          frmMsgBox.frm                                               |
'|  Type:          VB6 Form                                                    |
'|  Project:       Cybera Workstation (http://cybera.sf.net)                   |
'|  Description:   Small message box used to warn the user.                    |
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

' $Id: frmMsgBox.frm 26 2006-02-01 13:41:15Z sam $

Option Explicit
Public uAction As Object
Public mfrmParent As Form


'*******************************************************************************
Private Sub cmdNo_Click()
On Error GoTo cmdNo_ClickError
    
    Call Unload(Me)
    If Not uAction Is Nothing Then Call uAction.subExecuteCommandButton2
    Exit Sub

cmdNo_ClickError:
    Call funErrorManager("frmMsgBox.cmdNo_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdOK_Click()
On Error GoTo cmdOK_ClickError

    Call Unload(Me)
    Exit Sub

cmdOK_ClickError:
    Call funErrorManager("frmMsgBox.cmdOK_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdYes_Click()
On Error GoTo cmdYes_ClickError

    Call Unload(Me)
    If Not uAction Is Nothing Then Call uAction.subExecuteCommandButton1
    Exit Sub

cmdYes_ClickError:
    Call funErrorManager("frmMsgBox.cmdYes_ClickError", True)
End Sub


'*******************************************************************************
Private Sub Form_Activate()
On Error GoTo Form_ActivateError

    Select Case cmdOK.Visible
    Case True
        Call funSetFocus(cmdOK)
    Case Else
        If cmdYes.Default Then
            Call funSetFocus(cmdYes)
        Else
            Call funSetFocus(cmdNo)
        End If
    End Select
    Exit Sub

Form_ActivateError:
    Call funErrorManager("frmMsgBox.Form_ActivateError", True)
End Sub


'*******************************************************************************
Private Sub Form_Load()
On Error GoTo Form_LoadError
    
    gblnfrmMsgBox = True
    Call subSetDefaultFont(Me)
    Call subLocalizeInterface
    If Not gblnNoTopmost Then
        Call SetWindowPos(frmMsgBox.hWnd, HWND_TOPMOST, (sysInfo.WorkAreaWidth - Me.Width) / 2 / Screen.TwipsPerPixelX, (sysInfo.WorkAreaHeight - Me.Height) / 2 / Screen.TwipsPerPixelY, 0, 0, SWP_NOSIZE)
    Else
        Call SetWindowPos(frmMsgBox.hWnd, HWND_NOTOPMOST, (sysInfo.WorkAreaWidth - Me.Width) / 2 / Screen.TwipsPerPixelX, (sysInfo.WorkAreaHeight - Me.Height) / 2 / Screen.TwipsPerPixelY, 0, 0, SWP_NOSIZE)
    End If
    Exit Sub

Form_LoadError:
    Call funErrorManager("frmMsgBox.Form_LoadError", True)
End Sub


'*******************************************************************************
Public Sub subLocalizeInterface()
On Error GoTo subLocalizeInterfaceError

    Call subSetCharset(Me, guCliOpts.intInterfaceLangCharset)
    
    Me.Caption = funLoadResString(10000) & " - " & funLoadResString(10106)
    cmdOK.Caption = funLoadResString(10266)
    cmdYes.Caption = funLoadResString(10267)
    cmdNo.Caption = funLoadResString(10268)
    Exit Sub

subLocalizeInterfaceError:
    Call funErrorManager("frmMsgBox.subLocalizeInterfaceError", True)
End Sub


'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)

    gblnfrmMsgBox = False
End Sub


'*******************************************************************************
Private Sub tmrTimeOut_Timer()
On Error GoTo tmrTimeOut_TimerError

    If Not Me.Visible Then Exit Sub
    If fraOK.Visible Then
        cmdOK.Value = True
    ElseIf cmdNo.Default Then
        cmdNo.Value = True
    Else
        cmdYes.Value = True
    End If
    Exit Sub

tmrTimeOut_TimerError:
    Call funErrorManager("frmMsgBox.tmrTimeOut_TimerError", True)
End Sub
