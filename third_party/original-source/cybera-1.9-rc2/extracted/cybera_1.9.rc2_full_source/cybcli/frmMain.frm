VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "sysinfo.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CybAdmin Client"
   ClientHeight    =   2505
   ClientLeft      =   150
   ClientTop       =   690
   ClientWidth     =   5055
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrActivateClient 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   660
      Top             =   1380
   End
   Begin VB.Frame fraControls 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.CommandButton cmdQuit 
         Caption         =   "_"
         Height          =   375
         Left            =   3420
         TabIndex        =   3
         Top             =   780
         Width           =   1515
      End
      Begin VB.CommandButton cmdShutdown 
         Caption         =   "_"
         Height          =   375
         Left            =   1860
         TabIndex        =   2
         Top             =   780
         Width           =   1515
      End
      Begin VB.CommandButton cmdActiver 
         Caption         =   "_"
         Default         =   -1  'True
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Top             =   780
         Width           =   1815
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   7
         Top             =   60
         Width           =   1935
      End
      Begin VB.Label lblServerName 
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
         TabIndex        =   6
         Top             =   60
         Width           =   2955
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "_"
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   5
         Top             =   420
         Width           =   1935
      End
      Begin VB.Image imgConnexionStatus 
         Height          =   240
         Left            =   4680
         Picture         =   "frmMain.frx":12FA
         Top             =   420
         Width           =   240
      End
      Begin VB.Label lblConnexionStatus 
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
         TabIndex        =   4
         Top             =   420
         Width           =   2655
      End
   End
   Begin MSWinsockLib.Winsock wskFindServer 
      Left            =   1140
      Top             =   1860
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "255.255.255.255"
      RemotePort      =   4646
   End
   Begin VB.Timer tmrURLspy 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4560
      Top             =   1380
   End
   Begin VB.Timer tmrWksnStatus 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2580
      Top             =   1800
   End
   Begin VB.Timer tmrAppKiller 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   4080
      Top             =   1380
   End
   Begin VB.Timer tmrSessionStatus 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2580
      Top             =   1380
   End
   Begin VB.Timer tmrActiveTrayIcon 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4560
      Top             =   2040
   End
   Begin VB.Timer tmrSrvCnxStatus 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1620
      Top             =   1380
   End
   Begin MSWinsockLib.Winsock wskCli 
      Left            =   1140
      Top             =   1380
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin SysInfoLib.SysInfo sysInfo 
      Left            =   60
      Top             =   1860
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Image imgCybActiveSmall 
      Height          =   240
      Left            =   60
      Picture         =   "frmMain.frx":1444
      Top             =   1620
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgCybInactive 
      Height          =   240
      Left            =   360
      Picture         =   "frmMain.frx":158E
      Top             =   1620
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgCybServSmall 
      Height          =   240
      Left            =   60
      Picture         =   "frmMain.frx":16D8
      Top             =   1320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgCybCliSmall 
      Height          =   240
      Left            =   360
      Picture         =   "frmMain.frx":1822
      Top             =   1320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLogOff 
         Caption         =   "LogOff"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuActivate 
         Caption         =   "Activate"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuStatus 
         Caption         =   "Status"
      End
      Begin VB.Menu mnuStopSession 
         Caption         =   "StopSession"
      End
      Begin VB.Menu mnuDeactivate 
         Caption         =   "Deactivate"
      End
      Begin VB.Menu mnuAboutPopup 
         Caption         =   "AboutPopup"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|__TIN TUA_______________________ MAIN FORM __________________________________|
'|  Name:          frmMain.frm                                                 |
'|  Type:          VB6 Form                                                    |
'|  Project:       Cybera Workstation (http://cybera.sf.net)                   |
'|  Description:   The Apps main form.                                         |
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

' $Id: frmMain.frm 26 2006-02-01 13:41:15Z sam $

Option Explicit


'*******************************************************************************
Private Sub cmdActiver_Click()
On Error GoTo cmdActiver_ClickError
    
    Call subDefaultState
    Exit Sub

cmdActiver_ClickError:
    Call funErrorManager("frmMain.cmdActiver_ClickError", Not gblnfrmBlock, True)
End Sub


'*******************************************************************************
Private Sub cmdQuit_Click()
On Error GoTo cmdQuit_ClickError
    
    If guCliOpts.strPasswd <> "" Then
        Call subQuitPassword(frmMain)
    Else
        Call subCloseCybCli
    End If
    
    Exit Sub
    Exit Sub
    
cmdQuit_ClickError:
    Call funErrorManager("frmMain.cmdQuit_ClickError", Not gblnfrmBlock, True)
End Sub


'*******************************************************************************
Private Sub cmdShutdown_Click()
On Error GoTo cmdShutdown_ClickError
    
    Call subLogOffPassword(frmMain)
    Exit Sub

cmdShutdown_ClickError:
    Call funErrorManager("frmMain.cmdShutdown_ClickError", Not gblnfrmBlock, True)
End Sub


'*******************************************************************************
Private Sub Form_Activate()
On Error Resume Next
    
    If Not gblnfrmAbout Then
        Call Me.ZOrder(0)
        Call Me.SetFocus
    End If
End Sub


'*******************************************************************************
Private Sub Form_Load()
On Error GoTo Form_LoadError
    
    Call subSetDefaultFont(Me)
    Call subLocalizeInterface
    lblConnexionStatus = funLoadResString(10337)
    If guCliOpts.uWinPos.intTop + frmMain.Height > Screen.Height Then guCliOpts.uWinPos.intTop = Screen.Height - frmMain.Height
    If guCliOpts.uWinPos.intLeft + frmMain.Width > Screen.Width Then guCliOpts.uWinPos.intLeft = Screen.Width - frmMain.Width
    frmMain.Top = guCliOpts.uWinPos.intTop
    frmMain.Left = guCliOpts.uWinPos.intLeft
    frmMain.lblServerName = IIf(guCliOpts.uServer.blnAutoConnect, funLoadResString(10310), guCliOpts.uServer.strServerName)
    
    With guIconData
        .cbSize = Len(guIconData)                       ' The length of the NOTIFYICONDATA type
        .hIcon = imgCybInactive.Picture                 ' A reference to the form's icon
        .hWnd = Me.hWnd                                 ' hWnd of the form
        .szTip = funLoadResString(10000) & Chr(0)                     ' Tooltip string delimited with a null character
        .uCallbackMessage = WM_MOUSEMOVE                ' The icon we're placing will send messages to the MouseMove event
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE    ' It will have message handling and a tooltip
        .uID = vbNull                                   ' uID is not used by VB, so it's set to a Null value
    End With
    mnuStopSession.Enabled = Not gblnNoCloseSession
    Exit Sub

Form_LoadError:
    Call funErrorManager("frmMain.Form_LoadError", Not gblnfrmBlock, True)
End Sub


'*******************************************************************************
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lngMsg As Long
On Error GoTo Form_MouseMoveError

    If guUser.intWksnState <= enuWksnState.Ready Or guUser.intWksnState > enuWksnState.Blocked_NoCredit Then Exit Sub
    lngMsg = x
    If lngMsg = WM_LBUTTONDBLCLK Or lngMsg = 7725 Then
        If frmSmallStatus.Visible Then
            Call subShowStatus
        Else
            Call subHideStatus
        End If
    ElseIf lngMsg = WM_RBUTTONDOWN Or lngMsg = 7740 Then '7740, go figure...
        Call PopupMenu(mnuPopup, , , , mnuStatus)
    End If
    Exit Sub

Form_MouseMoveError:
    Call funErrorManager("frmMain.Form_MouseMoveError", Not gblnfrmBlock, True)
End Sub


'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Form_UnloadError
    
    guCliOpts.uWinPos.intTop = frmMain.Top
    guCliOpts.uWinPos.intLeft = frmMain.Left
     
     '*Start clean shutdown.
    If funShutdownApplication = False Then Cancel = DO_NOT_SHUTDOWN
    Call Shell_NotifyIcon(NIM_DELETE, guIconData)
    Call subDebugPrint("App shutdown")
    End
    Exit Sub

Form_UnloadError:
    Call funErrorManager("frmMain.Form_UnloadError", Not gblnfrmBlock, True)
End Sub


'*******************************************************************************
Private Sub mnuAbout_Click()
On Error GoTo mnuAbout_ClickError

    Call Load(frmAbout)
    Call SetWindowPos(frmAbout.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    Call frmAbout.Show(vbModal, frmMain)
    Exit Sub

mnuAbout_ClickError:
    Call funErrorManager("frmMain.mnuAbout_ClickError", Not gblnfrmBlock, True)
End Sub


'*******************************************************************************
Private Sub mnuAboutPopup_Click()
On Error GoTo mnuAboutPopup_ClickError

    Call Load(frmAbout)
    frmAbout.Top = (Screen.Height - frmAbout.Height) / 2
    frmAbout.Left = (Screen.Width - frmAbout.Width) / 2
    Call frmAbout.Show
    Exit Sub

mnuAboutPopup_ClickError:
    Call funErrorManager("frmMain.mnuAboutPopup_ClickError", Not gblnfrmBlock, True)
End Sub


'*******************************************************************************
Private Sub mnuActivate_Click()
On Error GoTo mnuActivate_ClickError

    Call cmdActiver_Click
    Exit Sub

mnuActivate_ClickError:
    Call funErrorManager("frmMain.mnuActivate_ClickError", Not gblnfrmBlock, True)
End Sub


'*******************************************************************************
Private Sub mnuDeactivate_Click()
On Error GoTo mnuDeactivate_ClickError
    
    If guCliOpts.strPasswd <> "" Then
        Call subDeactivationPassword(Me)
    Else
        Call subDeactivate
    End If
    Exit Sub

mnuDeactivate_ClickError:
    Call funErrorManager("frmMain.mnuDeactivate_ClickError", Not gblnfrmBlock, True)
End Sub


'*******************************************************************************
Private Sub mnuLogOff_Click()
On Error GoTo mnuLogOff_ClickError

    Call subLogOffPassword(frmMain)
    Exit Sub

mnuLogOff_ClickError:
    Call funErrorManager("frmMain.mnuLogOff_ClickError", Not gblnfrmBlock, True)
End Sub

'*******************************************************************************
Private Sub mnuOptions_Click()
On Error GoTo mnuOptions_ClickError

    Call Load(frmOptions)
    Call SetWindowPos(frmOptions.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    Call frmOptions.Show(vbModal, frmMain)
    Exit Sub

mnuOptions_ClickError:
    Call funErrorManager("frmMain.mnuOptions_ClickError", Not gblnfrmBlock, True)
End Sub


'*******************************************************************************
Private Sub mnuQuit_Click()
On Error GoTo mnuQuit_ClickError

    If guCliOpts.strPasswd <> "" Then
        Call subQuitPassword(frmMain)
    Else
        Call subCloseCybCli
    End If
    Exit Sub

mnuQuit_ClickError:
    Call funErrorManager("frmMain.mnuQuit_ClickError", Not gblnfrmBlock, True)
End Sub


'*******************************************************************************
Private Sub mnuStatus_Click()
On Error GoTo mnuStatus_ClickError
    
    If frmSmallStatus.Visible Then
        Call subShowStatus
    Else
        Call subHideStatus
    End If
    Exit Sub

mnuStatus_ClickError:
    Call funErrorManager("frmMain.mnuStatus_ClickError", Not gblnfrmBlock, True)
End Sub


'*******************************************************************************
Private Sub mnuStopSession_Click()
On Error GoTo mnuStopSession_ClickError
    
    
    Call subShowMsgBox(funLoadResString(10184), YesNo, Button2, New clsActionCloseSession)
    Exit Sub

mnuStopSession_ClickError:
    Call funErrorManager("frmMain.mnuStopSession_ClickError", Not gblnfrmBlock, True)
End Sub


'*******************************************************************************
Private Sub sysInfo_TimeChanged()
Dim strBuffer As String
On Error GoTo sysInfo_TimeChangedError

    If guUser.intWksnState > enuWksnState.Blocked_SessionEnd Then
        If guUser.intWksnState < enuWksnState.Blocked_NoCredit Then
            Call funLockSession
            Call subBlock
        End If
        guCliOpts.uFlags.dteTimer = Now
        guCliOpts.uFlags.blnWaitingForServerDate = True
        DoEvents
    End If
    strBuffer = CNX_SERVER_DATE & CNX_MSG_END
    Call funSendData(frmMain.wskCli, strBuffer)
    
    Exit Sub

sysInfo_TimeChangedError:
    Call funErrorManager("frmMain.sysInfo_TimeChangedError", Not gblnfrmBlock, True)
End Sub


'*******************************************************************************
Private Sub tmrActivateClient_Timer()
On Error GoTo tmrActivateClient_TimerError

    
    gblnCnxOK = True
'    frmStatus.cmdCliPassword.Enabled = True
'    frmStatus.cmdChat.Enabled = True
    tmrSrvCnxStatus.Enabled = True
    tmrActivateClient.Enabled = False
    Exit Sub
    
tmrActivateClient_TimerError:
    Call funErrorManager("frmMain.tmrActivateClient_TimerError", True)
End Sub


'*******************************************************************************
Private Sub tmrActiveTrayIcon_Timer()
Static blnImageOn As Boolean
On Error GoTo tmrActiveTrayIcon_TimerError
    
    If blnImageOn Then
        guIconData.hIcon = imgCybInactive.Picture
    Else
        guIconData.hIcon = imgCybCliSmall.Picture
    End If
    Call Shell_NotifyIcon(NIM_MODIFY, guIconData)
    blnImageOn = Not blnImageOn
    Exit Sub

tmrActiveTrayIcon_TimerError:
    Call funErrorManager("frmMain.tmrActiveTrayIcon_TimerError", Not gblnfrmBlock, True)
    tmrActiveTrayIcon.Enabled = False
End Sub


'*******************************************************************************
Private Sub tmrAppKiller_Timer()
On Error GoTo tmrAppKiller_TimerError
    
    Call subKillApps
    Exit Sub

tmrAppKiller_TimerError:
    Call funErrorManager("frmMain.tmrAppKiller_TimerError", Not gblnfrmBlock, True)
    tmrAppKiller.Enabled = False
End Sub


'*******************************************************************************
Public Sub tmrWksnStatus_Timer()
Static blnSerializationOngoing As Boolean
On Error GoTo tmrWksnStatus_TimerError

    If wskCli.State = sckConnected Then
        blnSerializationOngoing = False
        If guCliOpts.uFlags.blnServerInitDone Then Call subUpdateServer
        If guUser.intSessionType = enuSessionType.Selfservice Then Call subSerializeUser
    Else
        If Not blnSerializationOngoing Then
            blnSerializationOngoing = True
            guUser.dteInitRemainingTimeSerialized = guUser.dteRemainingTime
        End If
        Call subSerializeUser
    End If
    Exit Sub

tmrWksnStatus_TimerError:
    Call funErrorManager("frmMain.tmrWksnStatus_TimerError", Not gblnfrmBlock, True)
    tmrWksnStatus.Enabled = False
End Sub


'*******************************************************************************
Public Sub tmrSessionStatus_Timer()
Dim dteTime As Date, dteCurrentTime As Date
Dim dteTimeDelta As Date
Dim dteTimeToNextIncrement As Date
Dim strBuffer As String
Dim blnDisplayRemainingTime As Boolean
On Error GoTo tmrSessionStatus_TimerError
    
    dteTime = Now
    dteTimeDelta = dteTime - guCliOpts.uFlags.dteTimer
    guCliOpts.uFlags.dteTimer = dteTime
    guUser.dteElapsedTime = guUser.dteElapsedTime + dteTimeDelta
    guUser.dteElapsedTimeAparent = guUser.dteElapsedTime - guUser.dteBlockedTime
    dteCurrentTime = CDate(Format(Now + guCliOpts.uFlags.dteServerDelta, "Hh:Nn:Ss"))
    Select Case guUser.intWksnState
    Case enuWksnState.UsedConnected, enuWksnState.UsedFree
        If funIsOffPeak(dteCurrentTime) Then
            guUser.dteTimeConnectedOffPeak = guUser.dteTimeConnectedOffPeak + dteTimeDelta
        Else
            guUser.dteTimeConnected = guUser.dteTimeConnected + dteTimeDelta
        End If
    Case enuWksnState.UsedDisconnected
        guUser.dteTimeDisconnected = guUser.dteTimeDisconnected + dteTimeDelta
    Case enuWksnState.Blocked_NoCredit
        guUser.dteBlockedTime = guUser.dteBlockedTime + dteTimeDelta
    End Select
    
    Select Case guUser.intSessionType
    Case enuSessionType.Card, enuSessionType.Client, enuSessionType.SelfServicePrePaid
        
        If guUser.blnUnlimitedPlay Then
            blnDisplayRemainingTime = False
        Else
            guUser.dteRemainingTime = guUser.dteInitRemainingTime - guUser.dteTimeConnected - guUser.dteTimeConnectedOffPeak * (guCliOpts.uTimeRatios.intOffPeak / 100) - guUser.dteTimeDisconnected * (guCliOpts.uTimeRatios.intDisconCon / 100)
            
            Select Case guUser.intWksnState
            Case enuWksnState.UsedConnected
                
                If guCliOpts.uTimeRatios.intOffPeak = 0 Or Not funIsOffPeak(dteCurrentTime) Then
                    guUser.dteRemainingTimeAparent = guUser.dteRemainingTime
                    blnDisplayRemainingTime = guCliOpts.uTimeRatios.intOffPeak <> 0
                Else
                    guUser.dteRemainingTimeAparent = guUser.dteRemainingTime / (guCliOpts.uTimeRatios.intOffPeak / 100)
                    blnDisplayRemainingTime = True
                End If
                
                
    '            If (dteCurrentTime < guCliOpts.uTimeRatios.dteOffPeakStop Or dteCurrentTime > guCliOpts.uTimeRatios.dteOffPeakStart) And guCliOpts.uTimeRatios.intOffPeak > 0 Then
    '                guUser.dteRemainingTimeAparent = guUser.dteRemainingTime / (guCliOpts.uTimeRatios.intOffPeak / 100)
    '            Else
    '                guUser.dteRemainingTimeAparent = guUser.dteRemainingTime
    '            End If
            
            Case enuWksnState.UsedDisconnected
                If guCliOpts.uTimeRatios.intDisconCon = 0 Then
                    guUser.dteRemainingTimeAparent = guUser.dteRemainingTime
                    blnDisplayRemainingTime = False
                Else
                    guUser.dteRemainingTimeAparent = guUser.dteRemainingTime / (guCliOpts.uTimeRatios.intDisconCon / 100)
                    blnDisplayRemainingTime = True
                End If
                
            Case enuWksnState.Blocked_Locked, enuWksnState.Blocked_Paused, enuWksnState.Blocked_follow
                Select Case guUser.intWksnStateOld
                Case enuWksnState.UsedConnected
                    If guCliOpts.uTimeRatios.intOffPeak = 0 Or Not funIsOffPeak(dteCurrentTime) Then
                        guUser.dteRemainingTimeAparent = guUser.dteRemainingTime
                        blnDisplayRemainingTime = guCliOpts.uTimeRatios.intOffPeak <> 0
                    Else
                        guUser.dteRemainingTimeAparent = guUser.dteRemainingTime / (guCliOpts.uTimeRatios.intOffPeak / 100)
                        blnDisplayRemainingTime = True
                    End If
    '                If dteCurrentTime < guCliOpts.uTimeRatios.dteOffPeakStop Or dteCurrentTime > guCliOpts.uTimeRatios.dteOffPeakStart Then
    '                    guUser.dteRemainingTimeAparent = guUser.dteRemainingTime / (guCliOpts.uTimeRatios.intOffPeak / 100)
    '                Else
    '                    guUser.dteRemainingTimeAparent = guUser.dteRemainingTime
    '                End If
                
                Case enuWksnState.UsedDisconnected
                    If guCliOpts.uTimeRatios.intDisconCon = 0 Then
                        guUser.dteRemainingTimeAparent = guUser.dteRemainingTime
                        blnDisplayRemainingTime = False
                    Else
                        guUser.dteRemainingTimeAparent = guUser.dteRemainingTime / (guCliOpts.uTimeRatios.intDisconCon / 100)
                        blnDisplayRemainingTime = True
                    End If
    '                guUser.dteRemainingTimeAparent = guUser.dteRemainingTime / (guCliOpts.uTimeRatios.intDisconCon / 100)
                
                End Select
            
            End Select
            
            If guUser.intWksnState = enuWksnState.UsedConnected Or guUser.intWksnState = enuWksnState.UsedDisconnected Then
                
                If guCliOpts.uWksnMsg.intMsgCard = vbChecked And _
                   (funFormatTime(guUser.dteRemainingTimeAparent) = "00:" & Format(guCliOpts.uWksnMsg.intMsgCardAlert1, "00") & ":00" _
                 Or funFormatTime(guUser.dteRemainingTimeAparent) = "00:" & Format(guCliOpts.uWksnMsg.intMsgCardAlert2, "00") & ":00") Then
                    Call subShowMsgBox(funLoadResString(10198) & DatePart("n", guUser.dteRemainingTimeAparent) & funLoadResString(10199))
                End If
                
                If guUser.dteRemainingTime <= NULL_DATE Then
                    guUser.dteRemainingTime = NULL_DATE
                    guUser.dteRemainingTimeAparent = NULL_DATE
                    Call subForceUpdate
                    guUser.intWksnStateOld = guUser.intWksnState
                    guUser.intWksnState = enuWksnState.Blocked_NoCredit
                    Call subBlock
                End If
            End If
        End If
        
        If blnDisplayRemainingTime Then
            frmStatus.lblRemainTime = funFormatTime(guUser.dteRemainingTimeAparent)
            frmStatus.lblCliRemainTime = funFormatTime(guUser.dteRemainingTimeAparent)
            frmStatus.lblSelfServRemainTime = funFormatTime(guUser.dteRemainingTimeAparent)
            frmSmallStatus.lblRemainTime = funFormatTime(guUser.dteRemainingTimeAparent)
        Else
            frmStatus.lblRemainTime = funLoadResString(10358)
            frmStatus.lblCliRemainTime = funLoadResString(10358)
            frmStatus.lblSelfServRemainTime = funLoadResString(10358)
            frmSmallStatus.lblRemainTime = funLoadResString(10358)
        End If
        frmStatus.lblValidDate = Format(guUser.dteCustomerExpiryDate, "Short Date")
        frmStatus.lblCliValidDate = Format(guUser.dteCustomerExpiryDate, "Short Date")
        
    Case enuSessionType.Selfservice
        
        guUser.sngBalance = funBalanceDue(guUser.dteTimeConnected + guUser.dteTimeConnectedOffPeak * (guCliOpts.uTimeRatios.intOffPeak / 100) + guUser.dteTimeDisconnected * (guCliOpts.uTimeRatios.intDisconCon / 100), dteTimeToNextIncrement)
        
'        frmStatus.lblElapsedTimeSelfService = funFormatTime(guUser.dteElapsedTimeAparent) & " "
'        frmSmallStatus.lblElapsedTimeSelfService = funFormatTime(guUser.dteElapsedTimeAparent) & " "
        frmStatus.lblElapsedTimeSelfService = funFormatTime(guUser.dteTimeConnected + guUser.dteTimeConnectedOffPeak + guUser.dteTimeDisconnected)
        frmSmallStatus.lblElapsedTimeSelfService = funFormatTime(guUser.dteTimeConnected + guUser.dteTimeConnectedOffPeak + guUser.dteTimeDisconnected)
        frmStatus.lblElapsedTimeSelfService.ToolTipText = funLoadResString(10356, funFormatTime(guUser.dteElapsedTimeAparent, enuTimeFormat.Hh_Nmn_Ssec))
        frmSmallStatus.lblElapsedTimeSelfService.ToolTipText = funLoadResString(10356, funFormatTime(guUser.dteElapsedTimeAparent, enuTimeFormat.Hh_Nmn_Ssec))
        
        frmStatus.lblAmountDue = funFormatCurrencyString(guUser.sngBalance, guCliOpts.strMonFormat, guCliOpts.strMonSymbol)
        frmStatus.lblTimeToNextIncrement = IIf(guCliOpts.uPostPaidSs.intMode = enuPostpaidSsMode.TimeBased Or guCliOpts.uPostPaidSs.intSubMode = enuPostpaidSsSubMode.RechSimple, funFormatTime(dteTimeToNextIncrement), "")
        frmSmallStatus.lblAmountDue = funFormatCurrencyString(guUser.sngBalance, guCliOpts.strMonFormat, guCliOpts.strMonSymbol)
        frmSmallStatus.lblTimeToNextIncrement = IIf(guCliOpts.uPostPaidSs.intMode = enuPostpaidSsMode.TimeBased Or guCliOpts.uPostPaidSs.intSubMode = enuPostpaidSsSubMode.RechSimple, funFormatTime(dteTimeToNextIncrement), "")
        
        If guCliOpts.uWksnMsg.intMsgSelfService = vbChecked Then
            If guUser.sngBalance >= guCliOpts.uWksnMsg.sngMsgSelfServiceBalance Then
                If guCliOpts.uWksnMsg.lngMsgSelfServiceBalDeltaRatio <> Int(guUser.sngBalance / guCliOpts.uWksnMsg.sngMsgSelfServiceBalance) Then
                    guCliOpts.uWksnMsg.lngMsgSelfServiceBalDeltaRatio = Int(guUser.sngBalance / guCliOpts.uWksnMsg.sngMsgSelfServiceBalance)
                    Call subShowMsgBox(funLoadResString(10200) & funFormatCurrencyString(guUser.sngBalance, guCliOpts.strMonFormat, guCliOpts.strMonSymbol) & funLoadResString(10201))
                End If
            End If
        End If
    
    Case enuSessionType.Free
    
    End Select
    Exit Sub

tmrSessionStatus_TimerError:
    Call funErrorManager("frmMain.tmrSessionStatus_TimerError", Not gblnfrmBlock, True)
    tmrSessionStatus.Enabled = False
End Sub


'*******************************************************************************
Private Sub tmrSrvCnxStatus_Timer()
Dim intCnxState As Integer
Static intOldCnxState As Integer
Dim strCnxMsg As String
Dim strBuffer As String

Static blnImageState As Boolean
On Error GoTo tmrSrvCnxStatus_TimerError

    intCnxState = frmMain.wskCli.State
    If gblnDebug Then strCnxMsg = intCnxState & " - "
    blnImageState = Not blnImageState
    Select Case intCnxState
    Case sckError, sckClosed, sckClosing ', sckConnecting
        If guCliOpts.uServer.blnAutoConnect Then
            Call subInitAutoConnection
        Else
            Call subConnectToServer
        End If
    End Select
    If Not gblnCnxOK Then Exit Sub
    Select Case intCnxState
    Case sckError, sckClosed, sckClosing, sckResolvingHost, sckConnecting
        frmMain.lblConnexionStatus = strCnxMsg & funLoadResString(10212)
'        frmBlock.lblConnexionStatus = funLoadResString(10336)
        frmBlock.imgNoServCnx.ToolTipText = funLoadResString(10338) & " - " & funLoadResString(10336)
        frmMain.lblServerName = IIf(guCliOpts.uServer.blnAutoConnect, funLoadResString(10310), guCliOpts.uServer.strServerName)
        If blnImageState Then
            imgConnexionStatus.Picture = imgCybCliSmall
        Else
            imgConnexionStatus.Picture = imgCybInactive
        End If
        guCliOpts.uServer.intCnxState = enuSrvCxn.Disconnected
        frmBlock.cmdCxn.Enabled = False
        frmBlock.cmdReCnx.Enabled = False
        frmBlock.tmrNoServCnx.Enabled = True
'        frmBlock.fraNoServCnx.Visible = True
        guCliOpts.uFlags.blnServerInitDone = False
        tmrWksnStatus.Enabled = True
    
    Case sckConnected
        Call subInitServer
        frmStatus.cmdCliPassword.Enabled = True
        frmStatus.cmdChat.Enabled = True
        frmMain.lblServerName = guCliOpts.uServer.strServerName
        frmMain.lblConnexionStatus = strCnxMsg & funLoadResString(10213)
        'frmBlock.lblConnexionStatus = ""
        frmBlock.imgNoServCnx.ToolTipText = funLoadResString(10338) & " - " & funLoadResString(10213)
        imgConnexionStatus.Picture = imgCybActiveSmall
        guCliOpts.uServer.intCnxState = enuSrvCxn.Connected
        frmBlock.cmdCxn.Enabled = (frmBlock.fraCardNumber.Visible And frmBlock.txtCardNum.Text <> "") Or (frmBlock.fraClient.Visible And frmBlock.txtLogin.Text <> "") Or Not (frmBlock.fraClient.Visible Or frmBlock.fraCardNumber.Visible)
        frmBlock.cmdReCnx.Enabled = guUser.intWksnState = enuWksnState.Blocked_NoCredit
        frmBlock.tmrNoServCnx.Enabled = False
        frmBlock.imgNoServCnx.Picture = frmBlock.imNoServCnx.ListImages(3).Picture
'        frmBlock.fraNoServCnx.Visible = False
        tmrSrvCnxStatus.Enabled = False
        tmrWksnStatus.Enabled = True
        Call wskFindServer.Close
        If guCliOpts.blnStartActive And guUser.intWksnState < enuWksnState.Ready Then Call subActivate
    
    End Select
    Exit Sub

tmrSrvCnxStatus_TimerError:
    Call funErrorManager("frmMain.tmrSrvCnxStatus_TimerError", Not gblnfrmBlock, True)
    tmrSrvCnxStatus.Enabled = False
    frmMain.lblConnexionStatus = funLoadResString(10260)
End Sub


'*******************************************************************************
Private Sub tmrURLspy_Timer()
Dim uInet As clsInetexplorer
Dim uNewURL As clsIEURL
Dim uTempURL As clsIEURL
Dim i As Integer
Dim blnOldWindow As Boolean
Dim strBuffer As String
On Error GoTo tmrURLspy_TimerError

    Set uInet = New clsInetexplorer
    
    For Each uTempURL In guCurrentURLs
        uTempURL.StillOpen = False
    Next uTempURL
    With uInet
        .Refresh
        For i = 1 To .Count
            If .ItemURL(i) <> URL_EMPTY_IE And .ItemCaption(i) <> URL_EMPTY_IE And Len(.ItemURL(i)) >= 5 And (InStr(.ItemURL(i), "/") > 0 Or Left(.ItemURL(i), 2) = "\\") Then
                Set uNewURL = New clsIEURL
                uNewURL.WindowhWnd = .ItemhWnd(i)
                uNewURL.URL = .ItemURL(i)
                uNewURL.Caption = .ItemCaption(i)
                For Each uTempURL In guCurrentURLs
                    If uTempURL.WindowhWnd = uNewURL.WindowhWnd Then
                        blnOldWindow = True
                        uTempURL.StillOpen = True
                        If uTempURL.URL <> uNewURL.URL And uTempURL.Caption <> uNewURL.Caption Then
                            uTempURL.Caption = uNewURL.Caption
                            uTempURL.URL = uNewURL.URL
                            strBuffer = CNX_SESSION_URL & uTempURL.WindowhWnd & CNX_PAYLOAD_TOKEN & uNewURL.Caption & CNX_PAYLOAD_TOKEN & uNewURL.URL & CNX_PAYLOAD_TOKEN & CNX_MSG_END
                            Call funSendData(frmMain.wskCli, strBuffer)
                        End If
                        Exit For
                    End If
                Next uTempURL
                If Not blnOldWindow Then
                    uNewURL.StillOpen = True
                    Call guCurrentURLs.Add(uNewURL, "key: " & uNewURL.WindowhWnd)
                    strBuffer = CNX_SESSION_URL & uNewURL.WindowhWnd & CNX_PAYLOAD_TOKEN & uNewURL.Caption & CNX_PAYLOAD_TOKEN & uNewURL.URL & CNX_PAYLOAD_TOKEN & CNX_MSG_END
                    Call funSendData(frmMain.wskCli, strBuffer)
                End If
            End If
        Next i
        For Each uTempURL In guCurrentURLs
            If Not uTempURL.StillOpen Then Call guCurrentURLs.Remove("key: " & uTempURL.WindowhWnd)
        Next uTempURL
    End With
    Set uInet = Nothing
    Exit Sub

tmrURLspy_TimerError:
    Set uInet = Nothing
    Call funErrorManager("frmMain.tmrURLspy_TimerError", Not gblnfrmBlock, True)
End Sub


'*******************************************************************************
Private Sub wskCli_Close()
On Error GoTo wskCli_CloseError
    
    tmrSrvCnxStatus.Enabled = True
    frmStatus.cmdCliPassword.Enabled = False
    frmStatus.cmdChat.Enabled = False
    If gblnfrmPasswdChange Then Call Unload(frmPasswdChange)
    guCliOpts.uFlags.blnServerInitDone = False
    Set guCliOpts.uConnectedManager = New clsUser
    Call wskCli.Close
'    If guUser.intWksnState >= enuWksnState.Ready And guUser.intWksnState <> enuWksnState.Blocked_Deactivated Then Call subLock
    Exit Sub

wskCli_CloseError:
    Call funErrorManager("frmMain.wskCli_CloseError", Not gblnfrmBlock, True)
End Sub


'*******************************************************************************
Private Sub wskCli_DataArrival(ByVal bytesTotal As Long)
Dim strNewData As String
Dim strTmpData As String, strEmptyData As String
Static strLastData As String
Static strDataBuffer As String
Static blnTempFileStorage As Boolean
Static intfreefile As Integer
Static strHeader As String
Dim strDatagram As String
Dim lngEndIndex As Long, lngTmpIndex As Long, lngFileHandle As Long, lngFileTemplate As Long
Dim intDebugStep As Integer
Dim uTmpSecAttrib As SECURITY_ATTRIBUTES
On Error GoTo wskCli_DataArrivalError

'intDebugStep = 0
    Call wskCli.GetData(strNewData, vbString)
    If Len(strDataBuffer) > MIN_SIZE_TEMP_FILE_STORAGE And Not blnTempFileStorage Then
        If gblnDebug Then Call subDebugPrint("wskCli_DataArrival: BEGIN TempFileStorage")
        blnTempFileStorage = True
        intfreefile = FreeFile
        On Error Resume Next
        Call Kill(guCliOpts.strAppConfigPath & FILENAME_RECEIVED)
        On Error GoTo wskCli_DataArrivalError
        Open guCliOpts.strAppConfigPath & FILENAME_RECEIVED For Binary Access Write As intfreefile
    End If
    
    If blnTempFileStorage Then
        If Len(strDataBuffer) > 0 Then
            strHeader = Left(strDataBuffer, 20)
            If strHeader = CNX_FILE_SEND Then
                lngEndIndex = InStr(strDataBuffer, CNX_PAYLOAD_TOKEN)
                strHeader = Left(strDataBuffer, lngEndIndex + 19)
                strDataBuffer = Mid(strDataBuffer, lngEndIndex + 20)
            Else
                strDataBuffer = Mid(strDataBuffer, 20)
            End If
            Put intfreefile, Seek(intfreefile), strDataBuffer
            strDataBuffer = ""
        End If
        lngEndIndex = InStr(strLastData & strNewData, CNX_MSG_END)
        If lngEndIndex <> 0 Then
            If lngEndIndex < Len(strLastData) Then
                Close intfreefile
                Open guCliOpts.strAppConfigPath & FILENAME_RECEIVED For Binary Access Read Write As intfreefile
                lngTmpIndex = FileLen(guCliOpts.strAppConfigPath & FILENAME_RECEIVED) - (Len(strLastData) - lngEndIndex)
                strTmpData = String(Len(strLastData) - lngEndIndex, " ")
                Get intfreefile, lngTmpIndex, strTmpData
                Close intfreefile
'                lngFileHandle = CreateFile(guCliOpts.strAppConfigPath & FILENAME_RECEIVED, GENERIC_WRITE, 0, uTmpSecAttrib, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, lngFileTemplate)
'                Call SetFilePointer(lngFileHandle, lngTmpIndex, Null, FILE_BEGIN)
'                Call SetEndOfFile(lngFileHandle)
'                Call CloseHandle(lngFileHandle)
'                strEmptyData = ""
'                Put intfreefile, lngTmpIndex, strEmptyData
            Else
                strTmpData = Left(strNewData, lngEndIndex - Len(strLastData) - 1) 'PB when message end is spaned over two
                Put intfreefile, Seek(intfreefile), strTmpData
    '            Put intfreefile, Seek(intfreefile), Left(strNewData, lngEndIndex - Len(strLastData) - 1) 'What the hell was that about?
                Close intfreefile
            End If
            If gblnDebug Then Call subDebugPrint("wskCli_DataArrival: END TempFileStorage")
            blnTempFileStorage = False
            strDataBuffer = Mid(strNewData, lngEndIndex - Len(strLastData) + 20)
            strDatagram = strHeader & CNX_FILE_TEMP & CNX_MSG_END
            Call subProcessServerData(strDatagram)
        Else
            Put intfreefile, Seek(intfreefile), strNewData
        End If
    Else
        strDataBuffer = strDataBuffer & strNewData
        Do Until strDataBuffer = ""
            lngEndIndex = InStr(strDataBuffer, CNX_MSG_END)
            If lngEndIndex = 0 Then Exit Do
            strDatagram = Left(strDataBuffer, lngEndIndex + 19)
            strDataBuffer = Mid(strDataBuffer, lngEndIndex + 20)
            If gblnDebug Then Call subDebugPrint("wskCli_DataArrival.ProcessServerData: " & Left(strDatagram, 200))
            Call subProcessServerData(strDatagram)
        Loop
    End If
        
    strLastData = strNewData
    Exit Sub

wskCli_DataArrivalError:
    Dim lngErrNum As Long
    Dim strErrDesc As String
    lngErrNum = Err.Number
    strErrDesc = Err.Description
    If lngErrNum <> sckBadState Then
        Call funErrorManager("frmMain.wskCli_DataArrivalError", Not gblnfrmBlock, True)
        Call wskCli.Close
        tmrSrvCnxStatus.Enabled = True
        guCliOpts.uFlags.blnServerInitDone = False
    End If
End Sub


'*******************************************************************************
Private Sub wskCli_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error GoTo wskCli_ErrorError
    
    Call wskCli_Close
    Exit Sub

wskCli_ErrorError:
    Call funErrorManager("frmMain.wskCli_ErrorError", Not gblnfrmBlock, True)
End Sub


'*******************************************************************************
Private Sub wskFindServer_DataArrival(ByVal bytesTotal As Long)
Dim strNewData As String
Dim strServerName As String
On Error GoTo wskFindServer_DataArrivalError

    Call wskFindServer.GetData(strNewData, vbString)
    If Left(strNewData, 20) = CNX_AUTO_CONFIG_HANDSHAKE_SERV Then
        strServerName = Mid(strNewData, 21)
        Call subDebugPrint("wskFindServer.RemoteHost: " & strServerName)
        Call subDebugPrint("wskFindServer.RemoteHostIP: " & wskFindServer.RemoteHostIP)
        guCliOpts.uServer.strServerName = IIf(strServerName = "", wskFindServer.RemoteHostIP, strServerName)
        guCliOpts.uServer.lngPort = DEFAULT_TCP_PORT
        Call subConnectToServer

    End If
    Exit Sub

wskFindServer_DataArrivalError:
    Call funErrorManager("frmMain.wskFindServer_DataArrivalError", Not gblnfrmBlock, True)
End Sub

'*******************************************************************************
Public Sub subLocalizeInterface()
On Error GoTo subLocalizeInterfaceError

    Call subSetCharset(Me, guCliOpts.intInterfaceLangCharset)
    
    Me.Caption = funLoadResString(10000) & IIf(BETA_VERSION, " - " & App.Major & "." & App.Minor & "." & App.Revision & " BETA", "") & " - " & IIf(guCliOpts.blnWksnTypeSndAln, guCliOpts.strComputerName, guCliOpts.strUserName)
    cmdActiver.Caption = funLoadResString(10113)
    cmdActiver.ToolTipText = funLoadResString(10301)
    cmdShutdown.Caption = funLoadResString(10300) & "..."
    cmdShutdown.ToolTipText = funLoadResString(10302)
    cmdQuit.Caption = funLoadResString(10114)
    cmdQuit.ToolTipText = funLoadResString(10303)
    lbl(0).Caption = funLoadResString(10139)
    lbl(1).Caption = funLoadResString(10140)
    mnuAbout.Caption = funLoadResString(10141) & funLoadResString(10000) & "..."
    mnuAboutPopup.Caption = funLoadResString(10141) & funLoadResString(10000) & "..."
    mnuDeactivate.Caption = funLoadResString(10143) & funLoadResString(10000)
    mnuFile.Caption = funLoadResString(10144)
    mnuHelp.Caption = funLoadResString(10145)
    mnuLogOff.Caption = funLoadResString(10146) & "..."
    mnuOptions.Caption = funLoadResString(10147) & "..."
    mnuQuit.Caption = funLoadResString(10110)
    mnuActivate.Caption = funLoadResString(10113)
    mnuStatus.Caption = funLoadResString(10112)
    mnuStopSession.Caption = funLoadResString(10148)
    mnuTools.Caption = funLoadResString(10149)
    Exit Sub

subLocalizeInterfaceError:
    Call funErrorManager("frmMain.subLocalizeInterfaceError", Not gblnfrmBlock, True)
End Sub

