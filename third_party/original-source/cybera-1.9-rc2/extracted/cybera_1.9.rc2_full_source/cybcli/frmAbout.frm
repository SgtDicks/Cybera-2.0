VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   ClientHeight    =   4380
   ClientLeft      =   5190
   ClientTop       =   4590
   ClientWidth     =   8070
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrQuit 
      Enabled         =   0   'False
      Interval        =   12000
      Left            =   6840
      Top             =   60
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   4035
      Left            =   0
      ScaleHeight     =   4035
      ScaleWidth      =   7395
      TabIndex        =   0
      Top             =   0
      Width           =   7395
      Begin VB.Frame fraAckn 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   1140
         TabIndex        =   15
         Top             =   2580
         Width           =   3375
         Begin VB.Label lblAckn 
            Caption         =   "_"
            Height          =   195
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Width           =   3315
         End
      End
      Begin VB.Timer tmrAcknScroller 
         Interval        =   100
         Left            =   6840
         Top             =   600
      End
      Begin VB.Shape shp 
         BorderColor     =   &H80000010&
         Height          =   915
         Index           =   1
         Left            =   1080
         Top             =   2280
         Width           =   3495
      End
      Begin VB.Label lblRelease 
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
         Left            =   2280
         TabIndex        =   13
         Top             =   540
         Width           =   2355
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Release"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   12
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Web:"
         ForeColor       =   &H80000011&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   11
         Top             =   3660
         Width           =   975
      End
      Begin VB.Label lblWebSite 
         Caption         =   "_"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1080
         MouseIcon       =   "frmAbout.frx":0442
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   3660
         Width           =   3555
      End
      Begin VB.Label lblAcknowledgements 
         BackColor       =   &H80000014&
         Caption         =   "_"
         ForeColor       =   &H80000011&
         Height          =   225
         Left            =   1080
         TabIndex        =   9
         Top             =   2280
         Width           =   3495
      End
      Begin VB.Label lblGpl 
         Caption         =   "_"
         ForeColor       =   &H80000011&
         Height          =   675
         Left            =   1080
         TabIndex        =   8
         Top             =   1500
         Width           =   3195
      End
      Begin VB.Label lblCopyright 
         Caption         =   "_"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   1080
         TabIndex        =   7
         Top             =   1200
         Width           =   3555
      End
      Begin VB.Shape shp 
         BorderColor     =   &H80000010&
         Height          =   3915
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   4695
      End
      Begin VB.Label lblEmail 
         Caption         =   "_"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1080
         MouseIcon       =   "frmAbout.frx":074C
         TabIndex        =   6
         Top             =   3420
         Width           =   3555
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Contact:"
         ForeColor       =   &H80000011&
         Height          =   195
         Index           =   4
         Left            =   60
         TabIndex        =   5
         Top             =   3420
         Width           =   975
      End
      Begin VB.Label lblVersion 
         Caption         =   "_"
         Height          =   195
         Left            =   2280
         TabIndex        =   4
         Top             =   900
         Width           =   2355
      End
      Begin VB.Label lblMonthYear 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "mm/yyyy"
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
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lbl 
         Caption         =   "Version"
         ForeColor       =   &H80000011&
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   2
         Top             =   900
         Width           =   1215
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Left            =   300
         Picture         =   "frmAbout.frx":0A56
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   1050
         TabIndex        =   1
         Top             =   60
         Width           =   3615
      End
      Begin VB.Label lbl 
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   3
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   4695
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ ABOUT FORM _________________________________|
'|  Name:          frmAbout.frm                                                |
'|  Type:          VB6 Form                                                    |
'|  Project:       Cybera Workstation (http://cybera.sf.net)                   |
'|  Description:   The About form.                                             |
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

' $Id: frmAbout.frm 15 2006-01-01 16:12:35Z sam $

Option Explicit
Private mintWait As Integer


'*******************************************************************************
Private Sub Form_Click()

    Call Unload(Me)
End Sub


'*******************************************************************************
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Unload(Me)
End Sub


'*******************************************************************************
Private Sub Form_Load()
Dim i As Integer
Dim strtemp As String
On Error GoTo Form_LoadError

    gblnfrmAbout = True
    Call subDefineHelpers
    Call subSetDefaultFont(Me)
    Call subLocalizeInterface
    pic.Width = shp(0).Width
    pic.Height = shp(0).Height
    Me.Width = shp(0).Width
    Me.Height = shp(0).Height
    lblAckn.Height = 255 * UBound(gstrHelpers)
    lblAckn.Top = 0 'fraAckn.Height
    For i = 0 To funArraySize(gstrHelpers) - 1
        strtemp = strtemp & gstrHelpers(i) & vbNewLine
    Next i
    lblAckn.Caption = strtemp
    
    lblVersion = App.Major & "." & App.Minor & "." & App.Revision
    If BETA_VERSION Then lblVersion = lblVersion & " beta"
    If gblnDebug Then lblVersion = lblVersion & " Debug"
    lblCopyright = MY_COPYRIGHT & " " & BUILD_YEAR & " " & MY_NAME
    lblEmail = MY_EMAIL
    lblWebSite = MY_URL
    lblRelease = App.Major & "." & App.Minor & IIf(RC_VERSION <> "", "." & RC_VERSION, "")
    lblMonthYear = BUILD_MONTH_YEAR
    lblGpl.Caption = funLoadResString(10010)
    mintWait = 0
    Exit Sub
    
Form_LoadError:
    Call funErrorManager("frmAbout.Form_LoadError", True)
End Sub


'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)

    gblnfrmAbout = False
End Sub


'*******************************************************************************
Private Sub fraAckn_Click()

    Call Unload(Me)
End Sub


'*******************************************************************************
Private Sub imgIcon_Click()

    Call Unload(Me)
End Sub


'*******************************************************************************
Private Sub lbl_Click(Index As Integer)
    
    Call Unload(Me)
End Sub


'*******************************************************************************
Private Sub lblAckn_Click()

    Call Unload(Me)
End Sub


'*******************************************************************************
Private Sub lblAcknowledgements_Click()

    Call Unload(Me)
End Sub


'*******************************************************************************
Private Sub lblCopyright_Click()

    Call Unload(Me)
End Sub

'*******************************************************************************
Private Sub lblEmail_Click()
On Error Resume Next

'    Call subSendEmail(MY_EMAIL)
End Sub


'*******************************************************************************
Private Sub lblGpl_Click()

    Call Unload(Me)
End Sub


'*******************************************************************************
Private Sub lblMonthYear_Click()

    Call Unload(Me)
End Sub


'*******************************************************************************
Private Sub lblRelease_Click()
On Error GoTo lblRelease_ClickError
    
    Call Unload(Me)
    Exit Sub
    
lblRelease_ClickError:
    Call funErrorManager("frmAbout.lblRelease_ClickError", True)
End Sub


'*******************************************************************************
Private Sub lblTitle_Click()

    Call Unload(Me)
End Sub


'*******************************************************************************
Private Sub lblVersion_Click()
    
    Call Unload(Me)
End Sub


'*******************************************************************************
Private Sub lblWebSite_Click()
On Error Resume Next

    Call subGotoWeb(MY_URL)
End Sub


'*******************************************************************************
Private Sub pic_Click()
    
    Call Unload(Me)
End Sub


'*******************************************************************************
Private Sub pic_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Unload(Me)
End Sub


'*******************************************************************************
Public Sub subLocalizeInterface()

    Call subSetCharset(Me, guCliOpts.intInterfaceLangCharset)
    
    lblTitle = funLoadResString(10000)
    lblAcknowledgements.Caption = " " & funLoadResString(10011)
    lbl(0).Caption = funLoadResString(10041)
    lbl(1).Caption = funLoadResString(10042)
    lbl(2).Caption = funLoadResString(10043)
    lbl(4).Caption = funLoadResString(10040)
End Sub


'*******************************************************************************
Private Sub tmrAcknScroller_Timer()
On Error GoTo tmrAcknScroller_TimerError

    If mintWait <= 10 Then mintWait = mintWait + 1
    If mintWait <= 10 Then Exit Sub
    If lblAckn.Top < -lblAckn.Height Then lblAckn.Top = fraAckn.Height
    lblAckn.Top = lblAckn.Top - Screen.TwipsPerPixelY
    Exit Sub
    
tmrAcknScroller_TimerError:
    Call funErrorManager("frmAbout.tmrAcknScroller_TimerError", True)
End Sub


'*******************************************************************************
Private Sub tmrQuit_Timer()
    
    Call Unload(Me)
End Sub
