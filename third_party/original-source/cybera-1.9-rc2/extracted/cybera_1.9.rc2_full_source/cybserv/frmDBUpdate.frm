VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDBUpdate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "_"
   ClientHeight    =   915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pgbDBUpgrade 
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   660
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Min             =   1e-4
      Max             =   5
      Scrolling       =   1
   End
   Begin VB.Shape shp 
      BackColor       =   &H80000014&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   195
      Index           =   1
      Left            =   60
      Top             =   660
      Width           =   4635
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   60
      Picture         =   "frmDBUpdate.frx":0000
      Top             =   60
      Width           =   480
   End
   Begin VB.Label lbl 
      Caption         =   "_"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   60
      Width           =   4095
   End
End
Attribute VB_Name = "frmDBUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ DB UPDATE PROGRESS FORM ____________________|
'|  Name:          frmDBUpdate.frm                                             |
'|  Type:          VB6 Form                                                    |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
'|  Description:   Database update progress.                                   |
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

' $Id: frmDBUpdate.frm 15 2006-01-01 16:12:35Z sam $

Option Explicit


'*******************************************************************************
Private Sub Form_Load()
On Error GoTo Form_LoadError

    Call subSetDefaultFont(Me)
    Call subLocalizeInterface
    pgbDBUpgrade.top = shp(1).top + Screen.TwipsPerPixelY
    pgbDBUpgrade.left = shp(1).left + Screen.TwipsPerPixelX
    pgbDBUpgrade.width = shp(1).width - Screen.TwipsPerPixelX * 2
    pgbDBUpgrade.height = shp(1).height - Screen.TwipsPerPixelY * 2
    Exit Sub

Form_LoadError:
    Call funErrorManager("frmDBUpdate.Form_LoadError", True)
End Sub


'*******************************************************************************
Public Sub subLocalizeInterface()
On Error GoTo subLocalizeInterfaceError

    Call subSetCharset(Me, guServOpts.uConfig.intInterfaceLangCharset)
    
    Me.Caption = funLoadResString(10000) & " - " & funLoadResString(10971)
    lbl.Caption = funLoadResString(10972)
    Exit Sub

subLocalizeInterfaceError:
    Call funErrorManager("frmDBUpdate.subLocalizeInterfaceError", True)
End Sub


