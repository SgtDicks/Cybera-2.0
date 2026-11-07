VERSION 5.00
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "sysinfo.ocx"
Begin VB.Form frmMicroStatus 
   BorderStyle     =   0  'None
   ClientHeight    =   870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2070
   Icon            =   "frmMicroStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   2070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "_"
      Height          =   615
      Left            =   60
      Picture         =   "frmMicroStatus.frx":12FA
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Cliquez ici pour fermer votre session"
      Top             =   60
      Width           =   1155
   End
   Begin SysInfoLib.SysInfo sysInfo 
      Left            =   1380
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Shape shpFormFrame 
      BorderColor     =   &H80000010&
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   1275
   End
End
Attribute VB_Name = "frmMicroStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ MICRO WORKSTATION STATUS FORM ______________|
'|  Name:          frmMicroStatus.frm                                          |
'|  Type:          VB6 Form                                                    |
'|  Project:       Cybera Workstation (http://cybera.sf.net)                   |
'|  Description:   Micro wksn status form.                                     |
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

' $Id: frmMicroStatus.frm 26 2006-02-01 13:41:15Z sam $

Option Explicit


'*******************************************************************************
Private Sub cmdQuit_Click()
On Error GoTo cmdQuit_ClickError

    Call funSetFocus(cmdQuit)
    Call subShowMsgBox(funLoadResString(10184), YesNo, Button2, New clsActionCloseSession, Me)
    Call funSetFocus(cmdQuit)
    Exit Sub

cmdQuit_ClickError:
    Call funErrorManager("frmMicroStatus.cmdQuit_ClickError", True)
End Sub


'*******************************************************************************
Private Sub Form_Load()
On Error GoTo Form_LoadError
    
    gblnfrmMicroStatus = True
    Call subSetDefaultFont(Me)
    Call subLocalizeInterface
    Me.Width = shpFormFrame.Width
    Me.Height = shpFormFrame.Height
    Call SetWindowPos(frmMicroStatus.hWnd, HWND_BOTTOM, (sysInfo.WorkAreaWidth - shpFormFrame.Width - 60) / Screen.TwipsPerPixelX, (sysInfo.WorkAreaHeight - shpFormFrame.Height - 60) / Screen.TwipsPerPixelY, 0, 0, SWP_NOSIZE)
    cmdQuit.Enabled = Not gblnNoCloseSession
    Exit Sub

Form_LoadError:
    Call funErrorManager("frmMicroStatus.Form_LoadError", True)
End Sub


'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Form_UnloadError
    
    gblnfrmMicroStatus = False
    Cancel = 1
    Exit Sub

Form_UnloadError:
    Call funErrorManager("frmMicroStatus.Form_UnloadError", True)
End Sub


'*******************************************************************************
Public Sub subLocalizeInterface()
On Error GoTo subLocalizeInterfaceError

    Call subSetCharset(Me, guCliOpts.intInterfaceLangCharset)
    
    Me.Caption = funLoadResString(10112)
    cmdQuit.Caption = funLoadResString(10114)
    cmdQuit.ToolTipText = funLoadResString(10167)
    Exit Sub

subLocalizeInterfaceError:
    Call funErrorManager("frmMicroStatus.subLocalizeInterfaceError", True)
End Sub

