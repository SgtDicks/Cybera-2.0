VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmTextPrevew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   ControlBox      =   0   'False
   Icon            =   "frmTextPrev.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraMessage 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   5115
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   4695
      Begin RichTextLib.RichTextBox rtbMessage 
         Height          =   4995
         Left            =   60
         TabIndex        =   2
         Top             =   60
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   8811
         _Version        =   393217
         BackColor       =   -2147483628
         BorderStyle     =   0
         Enabled         =   0   'False
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmTextPrev.frx":0442
      End
      Begin VB.Shape shp 
         BorderColor     =   &H80000010&
         Height          =   5115
         Left            =   0
         Top             =   0
         Width           =   4695
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "_"
      Default         =   -1  'True
      Height          =   375
      Left            =   3420
      TabIndex        =   0
      Top             =   5280
      Width           =   1335
   End
End
Attribute VB_Name = "frmTextPrevew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ RTF PREVIEW FORM ___________________________|
'|  Name:          frmImgPrevew.frm                                            |
'|  Type:          VB6 Form                                                    |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
'|  Description:   Message text preview form.                                  |
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

' $Id: frmTextPrev.frm 15 2006-01-01 16:12:35Z sam $

Option Explicit


'*******************************************************************************
Private Sub cmdClose_Click()
On Error GoTo cmdClose_ClickError

    Call Unload(Me)
    Exit Sub

cmdClose_ClickError:
    Call funErrorManager("frmTextPrevew.cmdClose_ClickError", True)
End Sub


'*******************************************************************************
Private Sub Form_Load()
On Error GoTo Form_LoadError

    Call subSetDefaultFont(Me)
    Call subLocalizeInterface
    Exit Sub

Form_LoadError:
    Call funErrorManager("frmTextPrevew.Form_LoadError", True)
End Sub


'*******************************************************************************
Public Sub subLocalizeInterface()
On Error GoTo subLocalizeInterfaceError

    Call subSetCharset(Me, guServOpts.uConfig.intInterfaceLangCharset)
    
    Me.Caption = funLoadResString(10013)
    cmdClose.Caption = funLoadResString(10012)
    Exit Sub

subLocalizeInterfaceError:
    Call funErrorManager("frmTextPrevew.subLocalizeInterfaceError", True)
End Sub

