VERSION 5.00
Begin VB.Form frmOptionsWksnLockDriveSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Drives"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSave 
      Caption         =   "_"
      Default         =   -1  'True
      Height          =   375
      Left            =   60
      TabIndex        =   26
      Top             =   2700
      Width           =   1875
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "_"
      Height          =   375
      Left            =   1980
      TabIndex        =   27
      Top             =   2700
      Width           =   1335
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   60
      TabIndex        =   28
      Top             =   780
      Width           =   3255
      Begin VB.CheckBox chkDrives 
         Caption         =   "_"
         Height          =   195
         Index           =   25
         Left            =   2520
         TabIndex        =   25
         Top             =   540
         Width           =   495
      End
      Begin VB.CheckBox chkDrives 
         Caption         =   "_"
         Height          =   195
         Index           =   24
         Left            =   2520
         TabIndex        =   24
         Top             =   300
         Width           =   495
      End
      Begin VB.CheckBox chkDrives 
         Caption         =   "_"
         Height          =   195
         Index           =   23
         Left            =   1920
         TabIndex        =   23
         Top             =   1500
         Width           =   495
      End
      Begin VB.CheckBox chkDrives 
         Caption         =   "_"
         Height          =   195
         Index           =   22
         Left            =   1920
         TabIndex        =   22
         Top             =   1260
         Width           =   495
      End
      Begin VB.CheckBox chkDrives 
         Caption         =   "_"
         Height          =   195
         Index           =   21
         Left            =   1920
         TabIndex        =   21
         Top             =   1020
         Width           =   495
      End
      Begin VB.CheckBox chkDrives 
         Caption         =   "_"
         Height          =   195
         Index           =   20
         Left            =   1920
         TabIndex        =   20
         Top             =   780
         Width           =   495
      End
      Begin VB.CheckBox chkDrives 
         Caption         =   "_"
         Height          =   195
         Index           =   19
         Left            =   1920
         TabIndex        =   19
         Top             =   540
         Width           =   495
      End
      Begin VB.CheckBox chkDrives 
         Caption         =   "_"
         Height          =   195
         Index           =   18
         Left            =   1920
         TabIndex        =   18
         Top             =   300
         Width           =   495
      End
      Begin VB.CheckBox chkDrives 
         Caption         =   "_"
         Height          =   195
         Index           =   17
         Left            =   1320
         TabIndex        =   17
         Top             =   1500
         Width           =   495
      End
      Begin VB.CheckBox chkDrives 
         Caption         =   "_"
         Height          =   195
         Index           =   16
         Left            =   1320
         TabIndex        =   16
         Top             =   1260
         Width           =   495
      End
      Begin VB.CheckBox chkDrives 
         Caption         =   "_"
         Height          =   195
         Index           =   15
         Left            =   1320
         TabIndex        =   15
         Top             =   1020
         Width           =   495
      End
      Begin VB.CheckBox chkDrives 
         Caption         =   "_"
         Height          =   195
         Index           =   14
         Left            =   1320
         TabIndex        =   14
         Top             =   780
         Width           =   495
      End
      Begin VB.CheckBox chkDrives 
         Caption         =   "_"
         Height          =   195
         Index           =   13
         Left            =   1320
         TabIndex        =   13
         Top             =   540
         Width           =   495
      End
      Begin VB.CheckBox chkDrives 
         Caption         =   "_"
         Height          =   195
         Index           =   12
         Left            =   1320
         TabIndex        =   12
         Top             =   300
         Width           =   495
      End
      Begin VB.CheckBox chkDrives 
         Caption         =   "_"
         Height          =   195
         Index           =   11
         Left            =   720
         TabIndex        =   11
         Top             =   1500
         Width           =   495
      End
      Begin VB.CheckBox chkDrives 
         Caption         =   "_"
         Height          =   195
         Index           =   10
         Left            =   720
         TabIndex        =   10
         Top             =   1260
         Width           =   495
      End
      Begin VB.CheckBox chkDrives 
         Caption         =   "_"
         Height          =   195
         Index           =   9
         Left            =   720
         TabIndex        =   9
         Top             =   1020
         Width           =   495
      End
      Begin VB.CheckBox chkDrives 
         Caption         =   "_"
         Height          =   195
         Index           =   8
         Left            =   720
         TabIndex        =   8
         Top             =   780
         Width           =   495
      End
      Begin VB.CheckBox chkDrives 
         Caption         =   "_"
         Height          =   195
         Index           =   7
         Left            =   720
         TabIndex        =   7
         Top             =   540
         Width           =   495
      End
      Begin VB.CheckBox chkDrives 
         Caption         =   "_"
         Height          =   195
         Index           =   6
         Left            =   720
         TabIndex        =   6
         Top             =   300
         Width           =   495
      End
      Begin VB.CheckBox chkDrives 
         Caption         =   "_"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   5
         Top             =   1500
         Width           =   495
      End
      Begin VB.CheckBox chkDrives 
         Caption         =   "_"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   1260
         Width           =   495
      End
      Begin VB.CheckBox chkDrives 
         Caption         =   "_"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Top             =   1020
         Width           =   495
      End
      Begin VB.CheckBox chkDrives 
         Caption         =   "_"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   780
         Width           =   495
      End
      Begin VB.CheckBox chkDrives 
         Caption         =   "_"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   540
         Width           =   495
      End
      Begin VB.CheckBox chkDrives 
         Caption         =   "_"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   495
      End
      Begin VB.Shape shp 
         BorderColor     =   &H80000010&
         Height          =   1815
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   3255
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
         Index           =   1
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   3255
      End
   End
   Begin VB.Label lbl 
      Caption         =   "_"
      Height          =   675
      Index           =   0
      Left            =   60
      TabIndex        =   30
      Top             =   60
      Width           =   3255
   End
End
Attribute VB_Name = "frmOptionsWksnLockDriveSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ DRIVE SELECTION POPUP ______________________|
'|  Name:          frmOptionsWksnLockDriveSelect.frm                           |
'|  Type:          VB6 Form                                                    |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
'|  Description:   Drive selection Popup used in the options                   |
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

' $Id: frmOptionsWksnLockDriveSelect.frm 15 2006-01-01 16:12:35Z sam $

Option Explicit

'*******************************************************************************
Private Sub cmdCancel_Click()
On Error GoTo cmdCancel_ClickError

    Call Unload(Me)
    Exit Sub
    
cmdCancel_ClickError:
    Call funErrorManager("frmOptionsWksnLockDriveSelect.cmdCancel_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdSave_Click()
Dim i As Integer
Dim lngDriveFlags As Long
On Error GoTo cmdSave_ClickError
    
    For i = 0 To 25
        If chkDrives(i).value = vbChecked Then
            lngDriveFlags = lngDriveFlags + 2 ^ i
        End If
    Next i
    guServOpts.uWorkstations.uLocks.lngLockedDrivesTmp = lngDriveFlags
    Call Unload(Me)
    Exit Sub

cmdSave_ClickError:
    Call funErrorManager("frmOptionsWksnLockDriveSelect.cmdSave_ClickError", True)
    Call Unload(Me)
End Sub


'*******************************************************************************
Private Sub Form_Load()
Dim i As Integer
Dim lngDriveFlags As Long
On Error GoTo Form_LoadError

    Call subSetDefaultFont(Me)
    Call subLocalizeInterface
    lngDriveFlags = guServOpts.uWorkstations.uLocks.lngLockedDrivesTmp
    For i = 25 To 0 Step -1
        chkDrives(i).Caption = Chr(65 + i) & ":"
        If lngDriveFlags - 2 ^ i >= 0 Then
            lngDriveFlags = lngDriveFlags - 2 ^ i
            chkDrives(i).value = vbChecked
        End If
    Next i
    Exit Sub

Form_LoadError:
    Call funErrorManager("frmOptionsWksnLockDriveSelect.Form_LoadError", True)
End Sub


'*******************************************************************************
Public Sub subLocalizeInterface()
On Error GoTo subLocalizeInterfaceError

    Call subSetCharset(Me, guServOpts.uConfig.intInterfaceLangCharset)
    
    Me.Caption = funLoadResString(10906)
    lbl(0).Caption = funLoadResString(10905)
    lbl(1).Caption = funLoadResString(10906)
    cmdSave.Caption = funLoadResString(10908)
    cmdCancel.Caption = funLoadResString(10907)
    Exit Sub

subLocalizeInterfaceError:
    Call funErrorManager("frmOptionsWksnLockDriveSelect.subLocalizeInterfaceError", True)
End Sub

