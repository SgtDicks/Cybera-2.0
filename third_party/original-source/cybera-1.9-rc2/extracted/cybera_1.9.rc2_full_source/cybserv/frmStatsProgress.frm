VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmStatsProgress 
   BorderStyle     =   0  'None
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4485
   Icon            =   "frmStatsProgress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdStop 
      Caption         =   "_"
      Height          =   375
      Left            =   3180
      TabIndex        =   0
      Top             =   1080
      Width           =   1155
   End
   Begin MSComctlLib.ProgressBar pgbStatProgress 
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   300
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pgbStatType 
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   780
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Min             =   1e-4
      Max             =   5
      Scrolling       =   1
   End
   Begin VB.Label lblType 
      Caption         =   "_"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   4275
   End
   Begin VB.Label lbl 
      Caption         =   "_"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   540
      Width           =   4275
   End
   Begin VB.Shape shp 
      BackColor       =   &H80000014&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   195
      Index           =   0
      Left            =   60
      Top             =   300
      Width           =   4275
   End
   Begin VB.Shape shp 
      BackColor       =   &H80000014&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   195
      Index           =   1
      Left            =   60
      Top             =   780
      Width           =   4275
   End
   Begin VB.Shape shpFormOutline 
      BorderColor     =   &H80000010&
      Height          =   1515
      Left            =   0
      Top             =   0
      Width           =   4395
   End
End
Attribute VB_Name = "frmStatsProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ STATS PROGRESS FORM ________________________|
'|  Name:          frmStatsProgress.frm                                        |
'|  Type:          VB6 Form                                                    |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
'|  Description:   Statistics calculation and display progress monitor         |
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

' $Id: frmStatsProgress.frm 15 2006-01-01 16:12:35Z sam $

Option Explicit

 '*******************************************************************************
Private Sub cmdStop_Click()
On Error GoTo cmdStop_ClickError
    
    gblnStopStats = True
    Exit Sub
    
cmdStop_ClickError:
    Call funErrorManager("frmStatsProgress.cmdStop_ClickError", True)
End Sub


'*******************************************************************************
Private Sub Form_Load()
On Error GoTo Form_LoadError
    
    Call subSetDefaultFont(Me)
    Me.width = shpFormOutline.width
    Me.height = shpFormOutline.height
    
    pgbStatType.top = shp(1).top + Screen.TwipsPerPixelY
    pgbStatType.left = shp(1).left + Screen.TwipsPerPixelX
    pgbStatType.width = shp(1).width - Screen.TwipsPerPixelX * 2
    pgbStatType.height = shp(1).height - Screen.TwipsPerPixelY * 2
    
    pgbStatProgress.top = shp(0).top + Screen.TwipsPerPixelY
    pgbStatProgress.left = shp(0).left + Screen.TwipsPerPixelX
    pgbStatProgress.width = shp(0).width - Screen.TwipsPerPixelX * 2
    pgbStatProgress.height = shp(0).height - Screen.TwipsPerPixelY * 2
    
    pgbStatProgress.value = pgbStatProgress.Min
    pgbStatType.value = pgbStatType.Min
    Call subLocalizeInterface
    Call Me.Show(vbModeless, frmMain)
    Call funSetFocus(cmdStop)
    Exit Sub
    
Form_LoadError:
    Call funErrorManager("frmStatsProgress.Form_LoadError", True)
End Sub


'*******************************************************************************
Public Sub subLocalizeInterface()
On Error GoTo subLocalizeInterfaceError

    Call subSetCharset(Me, guServOpts.uConfig.intInterfaceLangCharset)
    
    lbl(0).Caption = funLoadResString(10434)
    cmdStop.Caption = funLoadResString(10435)
    Exit Sub
    
subLocalizeInterfaceError:
    Call funErrorManager("frmStatsProgress.subLocalizeInterfaceError", True)
End Sub

