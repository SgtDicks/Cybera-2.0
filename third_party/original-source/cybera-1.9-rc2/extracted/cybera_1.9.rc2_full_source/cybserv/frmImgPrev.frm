VERSION 5.00
Begin VB.Form frmImgPrevew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   ControlBox      =   0   'False
   Icon            =   "frmImgPrev.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "_"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Shape shp 
      BorderColor     =   &H80000010&
      Height          =   5115
      Left            =   60
      Top             =   60
      Width           =   5115
   End
   Begin VB.Image imgLogo 
      Height          =   4995
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4995
   End
End
Attribute VB_Name = "frmImgPrevew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ IMAGE PREVIEW FORM _________________________|
'|  Name:          frmImgPrevew.frm                                            |
'|  Type:          VB6 Form                                                    |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
'|  Description:   Logo preview form.                                          |
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

' $Id: frmImgPrev.frm 15 2006-01-01 16:12:35Z sam $

Option Explicit


'*******************************************************************************
Private Sub cmdClose_Click()
On Error GoTo cmdClose_ClickError

    Call Unload(Me)
    Exit Sub

cmdClose_ClickError:
    Call funErrorManager("frmImgPrevew.cmdClose_ClickError", True)
End Sub


'*******************************************************************************
Private Sub Form_Load()
On Error GoTo Form_LoadError

    Call subSetDefaultFont(Me)
    Call subLocalizeInterface
    Exit Sub

Form_LoadError:
    Call funErrorManager("frmImgPrevew.Form_LoadError", True)
End Sub


'*******************************************************************************
Public Sub subLocalizeInterface()
On Error GoTo subLocalizeInterfaceError

    Call subSetCharset(Me, guServOpts.uConfig.intInterfaceLangCharset)
    
    Me.Caption = funLoadResString(10013)
    cmdClose.Caption = funLoadResString(10012)
    Exit Sub

subLocalizeInterfaceError:
    Call funErrorManager("frmImgPrevew.subLocalizeInterfaceError", True)
End Sub

