VERSION 5.00
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "sysinfo.ocx"
Begin VB.Form frmStatus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "_"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3675
   ControlBox      =   0   'False
   Icon            =   "frmStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   3675
   Begin VB.Frame fraChatClosed 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   60
      TabIndex        =   37
      Top             =   3180
      Visible         =   0   'False
      Width           =   3915
      Begin VB.CommandButton cmdChatOpen 
         Height          =   195
         Left            =   3270
         Picture         =   "frmStatus.frx":12FA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   30
         Width           =   255
      End
      Begin VB.Shape shpLine 
         BackColor       =   &H80000014&
         BorderColor     =   &H80000010&
         Height          =   255
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   3555
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
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
         Index           =   10
         Left            =   0
         TabIndex        =   38
         Top             =   0
         Width           =   3555
      End
   End
   Begin VB.Frame fraChat 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   60
      TabIndex        =   34
      Top             =   3180
      Width           =   3855
      Begin VB.CommandButton cmdChatClose 
         Height          =   195
         Left            =   3270
         Picture         =   "frmStatus.frx":13A4
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   30
         Width           =   255
      End
      Begin VB.TextBox txtWksnChat 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   180
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   1320
         Width           =   2775
      End
      Begin VB.CommandButton cmdChat 
         Height          =   315
         Left            =   3060
         Picture         =   "frmStatus.frx":144E
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1260
         Width           =   375
      End
      Begin VB.TextBox txtWksnChatHist 
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   36
         Top             =   360
         Width           =   3315
      End
      Begin VB.Shape shpWksnChatHist 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   315
         Left            =   120
         Top             =   1260
         Width           =   2895
      End
      Begin VB.Shape shpLine 
         BorderColor     =   &H80000010&
         Height          =   1695
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   3555
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
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
         Index           =   7
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Width           =   3555
      End
   End
   Begin VB.Timer tmrHideStatus 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   1680
      Top             =   5400
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "_"
      Height          =   615
      Left            =   2460
      Picture         =   "frmStatus.frx":1598
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cliquez ici pour fermer votre session"
      Top             =   5280
      Width           =   1155
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "_"
      Default         =   -1  'True
      Height          =   615
      Left            =   60
      Picture         =   "frmStatus.frx":16E2
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Cliquez ici pour cacher cette fenêtre"
      Top             =   5280
      Width           =   1155
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "_"
      Height          =   615
      Left            =   1260
      Picture         =   "frmStatus.frx":182C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Cliquez ici pour suspendre votre session"
      Top             =   5280
      Width           =   1155
   End
   Begin SysInfoLib.SysInfo sysInfo 
      Left            =   120
      Top             =   1140
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   60
      TabIndex        =   11
      Top             =   1140
      Width           =   3795
      Begin VB.Frame fraSelfServicePrePaid 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   3375
         Begin VB.Label lblSelfServRemainTime 
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
            Left            =   1980
            TabIndex        =   41
            Top             =   0
            Width           =   1335
         End
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
            Index           =   11
            Left            =   0
            TabIndex        =   40
            Top             =   0
            Width           =   1935
         End
      End
      Begin VB.Frame fraSelfService 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   3375
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
            Left            =   1980
            TabIndex        =   42
            Top             =   240
            Width           =   1335
         End
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
            Index           =   5
            Left            =   0
            TabIndex        =   21
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label lblElapsedTimeSelfService 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
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
            Left            =   1980
            TabIndex        =   20
            Top             =   480
            Width           =   1335
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
            TabIndex        =   19
            Top             =   60
            Width           =   1935
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
            Left            =   1980
            TabIndex        =   18
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.Frame fraClient 
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   3375
         Begin VB.CommandButton cmdCliPassword 
            Caption         =   "_"
            Height          =   375
            Left            =   0
            TabIndex        =   5
            Top             =   1080
            Width           =   3315
         End
         Begin VB.Label lblCliLogin 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
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
            Left            =   1980
            TabIndex        =   33
            Top             =   720
            Width           =   1335
         End
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
            Index           =   1
            Left            =   0
            TabIndex        =   32
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label lblCliValidDate 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
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
            Left            =   1980
            TabIndex        =   31
            Top             =   360
            Width           =   1335
         End
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
            Index           =   9
            Left            =   0
            TabIndex        =   30
            Top             =   360
            Width           =   1935
         End
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
            Index           =   8
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   1935
         End
         Begin VB.Label lblCliRemainTime 
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
            Left            =   1980
            TabIndex        =   28
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.Frame fraCard 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   3375
         Begin VB.Frame fraCardNum 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   0
            TabIndex        =   24
            Top             =   720
            Width           =   3315
            Begin VB.Label lblNumCard 
               Alignment       =   2  'Center
               BackColor       =   &H80000018&
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
               Left            =   1980
               TabIndex        =   26
               Top             =   0
               Width           =   1335
            End
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
               Index           =   0
               Left            =   0
               TabIndex        =   25
               Top             =   0
               Width           =   1935
            End
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
            Left            =   1980
            TabIndex        =   16
            Top             =   0
            Width           =   1335
         End
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
            TabIndex        =   15
            Top             =   0
            Width           =   1935
         End
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
            Index           =   2
            Left            =   0
            TabIndex        =   13
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label lblValidDate 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
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
            Left            =   1980
            TabIndex        =   14
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Shape shpLine 
         BorderColor     =   &H80000010&
         Height          =   1935
         Index           =   2
         Left            =   0
         Top             =   0
         Width           =   3555
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
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
         Index           =   4
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   3555
      End
   End
   Begin VB.Frame fraModeCnx 
      BorderStyle     =   0  'None
      Height          =   1035
      Left            =   60
      TabIndex        =   10
      Top             =   60
      Width           =   3855
      Begin VB.OptionButton optConnected 
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
         Height          =   195
         Left            =   60
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Cliquez ici pour utiliser le mode connecté à internet"
         Top             =   360
         Value           =   -1  'True
         Width           =   3435
      End
      Begin VB.OptionButton optDisconnected 
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
         Height          =   195
         Left            =   60
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Cliquez ici pour utiliser le mode déconnecté. Vous n'aurez pas accès à internet"
         Top             =   660
         Width           =   3435
      End
      Begin VB.Shape shpLine 
         BorderColor     =   &H80000010&
         Height          =   975
         Index           =   7
         Left            =   0
         Top             =   0
         Width           =   3555
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
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
         Index           =   6
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   3555
      End
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ MAIN WORKSTATION STATUS FORM _______________|
'|  Name:          frmStatus.frm                                               |
'|  Type:          VB6 Form                                                    |
'|  Project:       Cybera Workstation (http://cybera.sf.net)                   |
'|  Description:   Main wksn. status form - shows the user time left etc.      |
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

' $Id: frmStatus.frm 26 2006-02-01 13:41:15Z sam $

Option Explicit


'*******************************************************************************
Private Sub cmdChat_Click()
Dim strText As String
Dim strLogin As String
On Error GoTo cmdChat_ClickError

    If txtWksnChat.Text = "" Then Exit Sub
    strLogin = funLoadResString(10349)
    Select Case guUser.intSessionType
    Case enuSessionType.Card, enuSessionType.Client
        strLogin = strLogin & " (" & IIf(Len(guUser.strID) > 10, Left(guUser.strID, 10) & "...", guUser.strID) & ")"
    End Select
    
    strText = txtWksnChat.Text
    strText = Replace(strText, CNX_MSG_END, "")
    
    txtWksnChatHist.Text = txtWksnChatHist.Text & IIf(txtWksnChatHist.Text <> "", vbNewLine, "") & " " & strLogin & ": " & strText
    txtWksnChatHist.SelStart = Len(txtWksnChatHist.Text)
    Call funSendData(frmMain.wskCli, CNX_CHAT_MSG & strText & CNX_MSG_END)
    txtWksnChat.Text = ""
    Exit Sub

cmdChat_ClickError:
    Call funErrorManager("frmStatus.cmdChat_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdChatClose_Click()
Dim lngNewHeight As Long
On Error Resume Next

    fraChatClosed.Visible = True
    fraChat.Visible = False
    cmdClose.Top = fra.Top + shpLine(2).Height + fraChatClosed.Height + 240
    cmdPause.Top = cmdClose.Top
    cmdQuit.Top = cmdClose.Top
    lngNewHeight = cmdClose.Top + cmdClose.Height + Me.Height - Me.ScaleHeight + 60
    If Me.Height + Me.Top = sysInfo.WorkAreaHeight Then Me.Top = sysInfo.WorkAreaHeight - lngNewHeight
    Me.Height = lngNewHeight
    cmdClose.Default = True
End Sub


'*******************************************************************************
Public Sub cmdChatOpen_Click()
On Error Resume Next

    fraChat.Visible = True
    fraChatClosed.Visible = False
    cmdClose.Top = fra.Top + shpLine(2).Height + fraChat.Height + 240
    cmdPause.Top = cmdClose.Top
    cmdQuit.Top = cmdClose.Top
    Me.Height = cmdClose.Top + cmdClose.Height + Me.Height - Me.ScaleHeight + 60
    If Me.Height + Me.Top > sysInfo.WorkAreaHeight Then Me.Top = sysInfo.WorkAreaHeight - Me.Height
    cmdChat.Default = True
    Call funSetFocus(txtWksnChat)
End Sub


'*******************************************************************************
Private Sub cmdCliPassword_Click()
On Error GoTo cmdCliPassword_ClickError

    tmrHideStatus.Enabled = False
    Call frmPasswdChange.Show(vbModal, Me)
    tmrHideStatus.Enabled = True
    Exit Sub

cmdCliPassword_ClickError:
    Call funErrorManager("frmStatus.cmdCliPassword_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdClose_Click()
On Error GoTo cmdClose_ClickError

    Call subHideStatus
    Exit Sub

cmdClose_ClickError:
    Call funErrorManager("frmStatus.cmdClose_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdPause_Click()
On Error GoTo cmdPause_ClickError
    
    Call subShowMsgBox(funLoadResString(10204), YesNo, Button2, New clsActionPauseSession, Me)
    Call funSetFocus(cmdClose)
    Exit Sub

cmdPause_ClickError:
    Call funErrorManager("frmStatus.cmdPause_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdQuit_Click()
On Error GoTo cmdQuit_ClickError
    
    Call funSetFocus(cmdClose)
    Call subShowMsgBox(funLoadResString(10184), YesNo, Button2, New clsActionCloseSession, Me)
    Exit Sub

cmdQuit_ClickError:
    Call funErrorManager("frmStatus.cmdQuit_ClickError", True)
End Sub


'*******************************************************************************
Private Sub Form_Load()
On Error GoTo Form_LoadError
    
    gblnfrmStatus = True
    Call subSetDefaultFont(Me)
    lblTimeToNextIncrement.FontName = "Arial"
    Call subLocalizeInterface
    If Not gblnNoTopmost Then
        Call SetWindowPos(frmStatus.hWnd, HWND_TOPMOST, (sysInfo.WorkAreaWidth - frmStatus.Width) / Screen.TwipsPerPixelX, (sysInfo.WorkAreaHeight - frmStatus.Height) / Screen.TwipsPerPixelY, 0, 0, SWP_NOSIZE)
    Else
        Call SetWindowPos(frmStatus.hWnd, HWND_NOTOPMOST, (sysInfo.WorkAreaWidth - frmStatus.Width) / Screen.TwipsPerPixelX, (sysInfo.WorkAreaHeight - frmStatus.Height) / Screen.TwipsPerPixelY, 0, 0, SWP_NOSIZE)
    End If
    cmdQuit.Enabled = Not gblnNoCloseSession
    Exit Sub

Form_LoadError:
    Call funErrorManager("frmStatus.Form_LoadError", True)
End Sub


'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    gblnfrmStatus = False
End Sub


'*******************************************************************************
Private Sub lbl_DblClick(Index As Integer)
On Error Resume Next

    If Index = 10 Then
        If cmdChatOpen.Visible Then
            cmdChatOpen.Value = True
        Else
            cmdChatClose.Value = True
        End If
    End If
End Sub


'*******************************************************************************
Private Sub optConnected_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo optConnected_MouseDownError
    
    If guUser.intWksnState <> enuWksnState.UsedConnected Then
        optConnected.Value = True
        If optDisconnected.Enabled = True Then
            Call subShowMsgBox(funLoadResString(10190), YesNo, Button1, New clsActionConnected, Me, , True)
        End If
    End If
    Exit Sub

optConnected_MouseDownError:
    Call funErrorManager("frmStatus.optConnected_MouseDownError", True)
End Sub


'*******************************************************************************
Private Sub optDisconnected_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo optDisconnected_MouseDownError
    
    If guUser.intWksnState <> enuWksnState.UsedDisconnected Then
        optDisconnected.Value = True
        If Me.Visible Then
            Call subShowMsgBox(funLoadResString(10188) & vbNewLine & vbNewLine & funLoadResString(10189), YesNo, Button2, New clsActionDisconnected, Me, , True)
        Else
            frmStatus.optConnected.Value = True
        End If
    End If
    Exit Sub

optDisconnected_MouseDownError:
    Call funErrorManager("frmStatus.optDisconnected_MouseDownError", True)
End Sub


'*******************************************************************************
Public Sub subLocalizeInterface()
On Error GoTo subLocalizeInterfaceError

    Call subSetCharset(Me, guCliOpts.intInterfaceLangCharset)
    
    Me.Caption = funLoadResString(10112)
    cmdClose.Caption = funLoadResString(10169)
    cmdClose.ToolTipText = funLoadResString(10170)
    cmdPause.Caption = funLoadResString(10162)
    cmdPause.ToolTipText = funLoadResString(10166)
    cmdQuit.Caption = funLoadResString(10114)
    cmdQuit.ToolTipText = funLoadResString(10167)
    cmdCliPassword.Caption = funLoadResString(10326)
    cmdCliPassword.ToolTipText = funLoadResString(10327)
    cmdChatClose.ToolTipText = funLoadResString(10352)
    cmdChatOpen.ToolTipText = funLoadResString(10351)
    txtWksnChatHist.ToolTipText = funLoadResString(10353)
    txtWksnChat.ToolTipText = funLoadResString(10354)
    cmdChat.ToolTipText = funLoadResString(10355)
    lbl(0).Caption = funLoadResString(10171)
    lbl(1).Caption = funLoadResString(10328)
    lbl(2).Caption = funLoadResString(10172)
    lbl(3).Caption = funLoadResString(10164)
    lbl(4).Caption = funLoadResString(10173)
    lbl(5).Caption = funLoadResString(10163)
    lbl(6).Caption = funLoadResString(10174)
    lbl(7).Caption = funLoadResString(10348)
    lbl(8).Caption = funLoadResString(10164)
    lbl(9).Caption = funLoadResString(10172)
    lbl(10).Caption = funLoadResString(10348)
    lbl(11).Caption = funLoadResString(10164)
    lblAmountDueLabel.Caption = funLoadResString(10165) '& guCliOpts.strMonSymbol & funLoadResString(10175)
    lblTimeToNextIncrement.ToolTipText = funLoadResString(10359)
    optConnected.Caption = funLoadResString(10130)
    optConnected.ToolTipText = funLoadResString(10131)
    optDisconnected.Caption = funLoadResString(10132)
    optDisconnected.ToolTipText = funLoadResString(10133)
    Exit Sub

subLocalizeInterfaceError:
    Call funErrorManager("frmStatus.subLocalizeInterfaceError", True)
End Sub


'*******************************************************************************
Private Sub tmrHideStatus_Timer()
On Error GoTo tmrHideStatus_TimerError

    If Me.Visible Then
        Call subHideStatus
    Else
        tmrHideStatus.Enabled = False
    End If
    Exit Sub

tmrHideStatus_TimerError:
    Call funErrorManager("frmStatus.tmrHideStatus_TimerError", True)
End Sub


'*******************************************************************************
Private Sub txtWksnChat_Change()
On Error Resume Next
    
    tmrHideStatus.Enabled = False
    tmrHideStatus.Enabled = True
End Sub


'*******************************************************************************
Private Sub txtWksnChat_GotFocus()
On Error Resume Next
    
    cmdChat.Default = True
End Sub


'*******************************************************************************
Private Sub txtWksnChat_LostFocus()
On Error Resume Next
    
    cmdClose.Default = True
End Sub
