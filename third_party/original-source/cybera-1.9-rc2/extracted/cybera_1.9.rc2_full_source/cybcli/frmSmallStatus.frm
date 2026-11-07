VERSION 5.00
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "sysinfo.ocx"
Begin VB.Form frmSmallStatus 
   BorderStyle     =   0  'None
   Caption         =   "CybAdmin Client - Etat de la session"
   ClientHeight    =   3105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4770
   ControlBox      =   0   'False
   Icon            =   "frmSmallStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "_"
      Height          =   615
      Left            =   2460
      Picture         =   "frmSmallStatus.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cliquez ici pour fermer votre session"
      Top             =   540
      Width           =   1155
   End
   Begin VB.CommandButton cmdStatus 
      Caption         =   "_"
      Default         =   -1  'True
      Height          =   615
      Left            =   60
      Picture         =   "frmSmallStatus.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   540
      Width           =   1155
   End
   Begin SysInfoLib.SysInfo sysInfo 
      Left            =   4200
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "_"
      Height          =   615
      Left            =   1260
      Picture         =   "frmSmallStatus.frx":06D6
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Cliquez ici pour suspendre votre session"
      Top             =   540
      Width           =   1155
   End
   Begin VB.Frame fraCard 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Double-cliquez ici pour faire apparaître la fenêtre d'état complète"
      Top             =   60
      Width           =   3495
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   12
         ToolTipText     =   "Double-cliquez ici pour faire apparaître la fenêtre d'état complète"
         Top             =   0
         Width           =   2175
      End
      Begin VB.Label lblRemainTime 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
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
         Left            =   2220
         TabIndex        =   11
         ToolTipText     =   "Double-cliquez ici pour faire apparaître la fenêtre d'état complète"
         Top             =   0
         Width           =   1275
      End
      Begin VB.Label lblElapsedTimeCard 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Caption         =   "00:00:00 "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         ToolTipText     =   "Double-cliquez ici pour faire apparaître la fenêtre d'état complète"
         Top             =   1500
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Durée de la session :"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   9
         ToolTipText     =   "Double-cliquez ici pour faire apparaître la fenêtre d'état complète"
         Top             =   1560
         Visible         =   0   'False
         Width           =   1515
      End
   End
   Begin VB.Frame fraSelfService 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   60
      TabIndex        =   3
      ToolTipText     =   "Double-cliquez ici pour faire apparaître la fenêtre d'état complète"
      Top             =   60
      Width           =   3555
      Begin VB.Label lblTimeToNextIncrement 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   150
         Left            =   2280
         TabIndex        =   13
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Durée de la session :"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Double-cliquez ici pour faire apparaître la fenêtre d'état complète"
         Top             =   720
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label lblElapsedTimeSelfService 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Caption         =   "00:00:00 "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         ToolTipText     =   "Double-cliquez ici pour faire apparaître la fenêtre d'état complète"
         Top             =   660
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label lblAmountDueLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   5
         ToolTipText     =   "Double-cliquez ici pour faire apparaître la fenêtre d'état complète"
         Top             =   60
         Width           =   2235
      End
      Begin VB.Label lblAmountDue 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
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
         Left            =   2280
         TabIndex        =   4
         ToolTipText     =   "Double-cliquez ici pour faire apparaître la fenêtre d'état complète"
         Top             =   0
         Width           =   1275
      End
   End
   Begin VB.Shape shpFormFrame 
      BorderColor     =   &H80000010&
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   3675
   End
End
Attribute VB_Name = "frmSmallStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ SMALL WORKSTATION STATUS FORM ______________|
'|  Name:          frmSmallStatus.frm                                          |
'|  Type:          VB6 Form                                                    |
'|  Project:       Cybera Workstation (http://cybera.sf.net)                   |
'|  Description:   Small wksn status form - shows the user time left etc.      |
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

' $Id: frmSmallStatus.frm 26 2006-02-01 13:41:15Z sam $

Option Explicit


'*******************************************************************************
Private Sub cmdPause_Click()
On Error GoTo cmdPause_ClickError

    Call funSetFocus(cmdStatus)
    Call subShowMsgBox(funLoadResString(10204), YesNo, Button2, New clsActionPauseSession, Me)
    Call funSetFocus(cmdStatus)
    Exit Sub

cmdPause_ClickError:
    Call funErrorManager("frmSmallStatus.cmdPause_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdQuit_Click()
On Error GoTo cmdQuit_ClickError

    Call funSetFocus(cmdStatus)
    Call subShowMsgBox(funLoadResString(10184), YesNo, Button2, New clsActionCloseSession, Me)
    Call funSetFocus(cmdStatus)
    Exit Sub

cmdQuit_ClickError:
    Call funErrorManager("frmSmallStatus.cmdQuit_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdStatus_Click()
On Error GoTo cmdStatus_ClickError

    Call subShowStatus
    Exit Sub

cmdStatus_ClickError:
    Call funErrorManager("frmSmallStatus.cmdStatus_ClickError", True)
End Sub


'*******************************************************************************
Private Sub Form_DblClick()
On Error GoTo Form_DblClickError
    
    Call subShowStatus
    Exit Sub

Form_DblClickError:
    Call funErrorManager("frmSmallStatus.Form_DblClickError", True)
End Sub


'*******************************************************************************
Private Sub Form_Load()
On Error GoTo Form_LoadError
    
    gblnfrmSmallStatus = True
    Call subSetDefaultFont(Me)
    lblTimeToNextIncrement.FontName = "Arial"
    Call subLocalizeInterface
    Me.Width = shpFormFrame.Width
    Me.Height = shpFormFrame.Height
    Call SetWindowPos(frmSmallStatus.hWnd, HWND_BOTTOM, (sysInfo.WorkAreaWidth - shpFormFrame.Width - 60) / Screen.TwipsPerPixelX, (sysInfo.WorkAreaHeight - shpFormFrame.Height - 60) / Screen.TwipsPerPixelY, 0, 0, SWP_NOSIZE)
    cmdQuit.Enabled = Not gblnNoCloseSession
    Exit Sub

Form_LoadError:
    Call funErrorManager("frmSmallStatus.Form_LoadError", True)
End Sub


'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Form_UnloadError
    
    gblnfrmSmallStatus = False
    Cancel = 1
    Exit Sub

Form_UnloadError:
    Call funErrorManager("frmSmallStatus.Form_UnloadError", True)
End Sub


'*******************************************************************************
Private Sub fraCard_DblClick()
On Error GoTo fraCard_DblClickError
    
    Call subShowStatus
    Exit Sub

fraCard_DblClickError:
    Call funErrorManager("frmSmallStatus.fraCard_DblClickError", True)
End Sub


'*******************************************************************************
Private Sub fraSelfService_DblClick()
On Error GoTo fraSelfService_DblClickError
    
    Call subShowStatus
    Exit Sub

fraSelfService_DblClickError:
    Call funErrorManager("frmSmallStatus.fraSelfService_DblClickError", True)
End Sub


'*******************************************************************************
Private Sub lbl_DblClick(Index As Integer)
On Error GoTo lbl_DblClickError
    
    Call subShowStatus
    Exit Sub

lbl_DblClickError:
    Call funErrorManager("frmSmallStatus.lbl_DblClickError", True)
End Sub


'*******************************************************************************
Private Sub lblAmountDue_DblClick()
On Error GoTo lblAmountDue_DblClickError
    
    Call subShowStatus
    Exit Sub

lblAmountDue_DblClickError:
    Call funErrorManager("frmSmallStatus.lblAmountDue_DblClickError", True)
End Sub


'*******************************************************************************
Private Sub lblAmountDueLabel_DblClick()
On Error GoTo lblAmountDueLabel_DblClickError
    
    Call subShowStatus
    Exit Sub

lblAmountDueLabel_DblClickError:
    Call funErrorManager("frmSmallStatus.lblAmountDueLabel_DblClickError", True)
End Sub


'*******************************************************************************
Private Sub lblElapsedTimeCard_DblClick()
On Error GoTo lblElapsedTimeCard_DblClickError
    
    Call subShowStatus
    Exit Sub

lblElapsedTimeCard_DblClickError:
    Call funErrorManager("frmSmallStatus.lblElapsedTimeCard_DblClickError", True)
End Sub


'*******************************************************************************
Private Sub lblElapsedTimeSelfService_DblClick()
On Error GoTo lblElapsedTimeSelfService_DblClickError
    
    Call subShowStatus
    Exit Sub

lblElapsedTimeSelfService_DblClickError:
    Call funErrorManager("frmSmallStatus.lblElapsedTimeSelfService_DblClickError", True)
End Sub


'*******************************************************************************
Private Sub lblRemainTime_DblClick()
On Error GoTo lblRemainTime_DblClickError
    
    Call subShowStatus
    Exit Sub

lblRemainTime_DblClickError:
    Call funErrorManager("frmSmallStatus.lblRemainTime_DblClickError", True)
End Sub


'*******************************************************************************
Public Sub subLocalizeInterface()
On Error GoTo subLocalizeInterfaceError

    Call subSetCharset(Me, guCliOpts.intInterfaceLangCharset)
    
    Me.Caption = funLoadResString(10112)
    cmdPause.Caption = funLoadResString(10162)
    cmdPause.ToolTipText = funLoadResString(10166)
    cmdQuit.Caption = funLoadResString(10114)
    cmdQuit.ToolTipText = funLoadResString(10167)
    cmdStatus.Caption = funLoadResString(10176)
    cmdStatus.ToolTipText = funLoadResString(10177)
    lbl(1).Caption = funLoadResString(10163)
    lbl(1).ToolTipText = funLoadResString(10168)
    lbl(3).Caption = funLoadResString(10164)
    lbl(3).ToolTipText = funLoadResString(10168)
    lbl(5).Caption = funLoadResString(10163)
    lbl(5).ToolTipText = funLoadResString(10168)
    lblAmountDueLabel.Caption = funLoadResString(10165) '& guCliOpts.strMonSymbol & funLoadResString(10175)
    fraCard.ToolTipText = funLoadResString(10168)
    fraSelfService.ToolTipText = funLoadResString(10168)
    lblAmountDue.ToolTipText = funLoadResString(10168)
    lblAmountDueLabel.ToolTipText = funLoadResString(10168)
    lblElapsedTimeCard.ToolTipText = funLoadResString(10168)
    lblElapsedTimeSelfService.ToolTipText = funLoadResString(10168)
    lblRemainTime.ToolTipText = funLoadResString(10168)
    lblAmountDue.ToolTipText = funLoadResString(10168)
    lblAmountDue.ToolTipText = funLoadResString(10168)
    lblAmountDue.ToolTipText = funLoadResString(10168)
    lblTimeToNextIncrement.ToolTipText = funLoadResString(10359)
    Exit Sub

subLocalizeInterfaceError:
    Call funErrorManager("frmSmallStatus.subLocalizeInterfaceError", True)
End Sub


'*******************************************************************************
Private Sub lblTimeToNextIncrement_Click()
On Error GoTo lblTimeToNextIncrement_ClickError
    
    Call subShowStatus
    Exit Sub

lblTimeToNextIncrement_ClickError:
    Call funErrorManager("frmSmallStatus.lblTimeToNextIncrement_ClickError", True)
End Sub
