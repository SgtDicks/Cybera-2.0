Attribute VB_Name = "modFunctions"
' _____________________________________________________________________________
'|________________________________ GENERAL FUNCTION MODULE ____________________|
'|  Name:          modFunctions.bas                                            |
'|  Type:          VB6 Module                                                  |
'|  Project:       Cybera Workstation (http://cybera.sf.net)                   |
'|  Description:   Contains general functions ans subs                         |
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

' $Id: modFunctions.bas 36 2006-02-08 03:11:24Z sam $

Option Explicit


'*******************************************************************************
Public Sub Main()
Dim blnMultiInstance As Boolean
Dim strCommandLine As String
Dim strStep As String
Dim lngTickCount As String
On Error GoTo MainError
    
    Call subDebugPrint("App startup")
    strStep = "CommandLine processing"
    strCommandLine = Replace(LCase(Command), "/", "-")
    gblnDebug = InStr(strCommandLine, "-d") <> 0
    gblnNoTopmost = InStr(strCommandLine, "-t") <> 0
    gblnNoKillApp = InStr(strCommandLine, "-k") <> 0
    blnMultiInstance = InStr(strCommandLine, "-m") <> 0
    gblnNoLockdown = InStr(strCommandLine, "-l") <> 0
    gblnNoStatusWindows = InStr(strCommandLine, "-s") <> 0
    gblnNoCloseSession = InStr(strCommandLine, "-c") <> 0

    strStep = "InitApp"
    Call subInitApp
    
    ' Delay app startup if the computer has just booted
    ' (in XP some fast computers can have problems with accessing GUI resources too soon)
    Call subDelay(IIf((guCliOpts.intStartDelay - GetTickCount / 1000) < 0, 0, guCliOpts.intStartDelay - GetTickCount / 1000))

    strStep = "Load forms"
    Call Load(frmMain)
    Call Load(frmStatus)
    Call Load(frmSmallStatus)
    Call Load(frmMicroStatus)
    
    strStep = "Check platform"
    If frmMain.sysInfo.OSPlatform <> 2 Then
        Call funMsg(OSNotNT)
        Call funShutdownApplication
    End If
    
    If Not blnMultiInstance Then Call subCheckForRunningWksn
    
    lngTickCount = GetTickCount
    strStep = "wksn-Server delay setup: " & lngTickCount
    ' Delay wksn-Server connection when the computer has just booted
    ' (the remote registry service is sometimes not yet loaded before the server starts user rights lockdown
    frmMain.tmrActivateClient.Interval = IIf((WKSN_BOOT_CNX_DELAY - lngTickCount) < 0, WKSN_MIN_CNX_DELAY, WKSN_BOOT_CNX_DELAY - lngTickCount)
    frmMain.tmrActivateClient.Enabled = True
    
    frmMain.tmrWksnStatus.Enabled = True
    If guCliOpts.blnStartActive And guCliOpts.uConfigWizard.intSystem >= CONFIG_WIZARD_CLI_SYSTEM And guCliOpts.uConfigWizard.intUser >= CONFIG_WIZARD_CLI_USER Then
        Call subDefaultState
    Else
        gblnNoKillApp = True
        Call subDeactivate
        gblnNoKillApp = InStr(strCommandLine, "-k") <> 0
    End If
    
    strStep = "frmMain/frmAbout GUI init"
    frmMain.Height = frmMain.Height - frmMain.ScaleHeight + frmMain.fraControls.Height
    If Not gblnDebug Then
        Call Load(frmAbout)
        Call frmAbout.Show
    
        If guCliOpts.blnStartActive And guCliOpts.uConfigWizard.intSystem >= CONFIG_WIZARD_CLI_SYSTEM And guCliOpts.uConfigWizard.intUser >= CONFIG_WIZARD_CLI_USER Then
            frmAbout.Top = (Screen.Height - frmAbout.Height) / 2
            frmAbout.Left = (Screen.Width - frmAbout.Width) / 2
        Else
            frmAbout.Top = guCliOpts.uWinPos.intTop + (frmMain.Height - frmAbout.Height) / 2
            frmAbout.Left = guCliOpts.uWinPos.intLeft + (frmMain.Width - frmAbout.Width) / 2
        End If
        Call SetWindowPos(frmAbout.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
        frmAbout.tmrQuit.Enabled = True
    End If
    
    strStep = "ConfigWizard"
    If guCliOpts.uConfigWizard.intSystem < CONFIG_WIZARD_CLI_SYSTEM Or guCliOpts.uConfigWizard.intUser < CONFIG_WIZARD_CLI_USER Then Call frmSetupWizard.Show(vbModal, frmMain)
    Exit Sub

MainError:
    Call funErrorManager("modFunctions.MainError(" & strStep & ")", True)
    Call subDeactivate
End Sub


'*******************************************************************************
Public Sub subBlock()
Dim strDefaultFilePrefix As String
On Error Resume Next

    If gblnfrmMsgBox Then Call Unload(frmMsgBox)
    If gblnfrmPasswdChange Then Call Unload(frmPasswdChange)
    If gblnfrmPasswd Then Call Unload(frmPasswd)
    Call frmStatus.Hide
    Call frmSmallStatus.Hide
    strDefaultFilePrefix = FILENAME_PREFIX & Right(guCliOpts.strInterfaceLang, Len(guCliOpts.strInterfaceLang) - Len(App.EXEName) - 1) & "_"
    With frmBlock
        Select Case guUser.intWksnState
        Case enuWksnState.Blocked_Deactivated
            .fraForm.Visible = True
            .fraHeader.Visible = True
            .fraFooter.Visible = True
            .fraFollowServer.Visible = False
            .fraControlsBlocked.Visible = True
            .fraControlsSessionEnd.Visible = False
            .fraControlsActive.Visible = False
            .fraControlsNoMoreCredit.Visible = False
            .fraControlsPaused.Visible = False
            .fraControlsLocked.Visible = False
            .rtbMessage.FileName = IIf(guCliOpts.intMsgLogoFilesType = vbChecked, guCliOpts.uPaths.strBlockedPath, guCliOpts.strAppConfigPath & strDefaultFilePrefix & FILENAME_BLOCKED)
        
        Case enuWksnState.Blocked_SessionEnd
            .fraForm.Visible = True
            .fraHeader.Visible = True
            .fraFooter.Visible = True
            .fraFollowServer.Visible = False
            .fraControlsBlocked.Visible = False
            .fraControlsSessionEnd.Visible = True
            .fraControlsActive.Visible = False
            .fraControlsNoMoreCredit.Visible = False
            .fraControlsPaused.Visible = False
            .fraControlsLocked.Visible = False
            .rtbMessage.FileName = IIf(guCliOpts.intMsgLogoFilesType = vbChecked, guCliOpts.uPaths.strSessionEndPath, guCliOpts.strAppConfigPath & strDefaultFilePrefix & FILENAME_SESSIONEND)
        
        Case enuWksnState.Blocked_NoCredit
            .fraForm.Visible = True
            .fraHeader.Visible = True
            .fraFooter.Visible = True
            .fraFollowServer.Visible = False
            .fraControlsBlocked.Visible = False
            .fraControlsSessionEnd.Visible = False
            .fraControlsActive.Visible = False
            .fraControlsNoMoreCredit.Visible = True
            .fraControlsPaused.Visible = False
            .fraControlsLocked.Visible = False
            .cmdReCnx.Default = True
            .cmdReCnx.Enabled = guCliOpts.uServer.intCnxState = enuSrvCxn.Connected
            .rtbMessage.FileName = IIf(guCliOpts.intMsgLogoFilesType = vbChecked, guCliOpts.uPaths.strNoCreditPath, guCliOpts.strAppConfigPath & strDefaultFilePrefix & FILENAME_NOCREDIT)
        
        Case enuWksnState.Blocked_Paused
            .fraForm.Visible = True
            .fraHeader.Visible = True
            .fraFooter.Visible = True
            .fraFollowServer.Visible = False
            .fraControlsBlocked.Visible = False
            .fraControlsSessionEnd.Visible = False
            .fraControlsActive.Visible = False
            .fraControlsNoMoreCredit.Visible = False
            .fraControlsPaused.Visible = True
            .fraControlsLocked.Visible = False
            .cmdUnPause.Default = True
            .rtbMessage.FileName = IIf(guCliOpts.intMsgLogoFilesType = vbChecked, guCliOpts.uPaths.strSessionPausedPath, guCliOpts.strAppConfigPath & strDefaultFilePrefix & FILENAME_SESSIONPAUSED)
        
        Case enuWksnState.Blocked_Locked
            .fraForm.Visible = True
            .fraHeader.Visible = True
            .fraFooter.Visible = True
            .fraFollowServer.Visible = False
            .fraControlsBlocked.Visible = False
            .fraControlsSessionEnd.Visible = False
            .fraControlsActive.Visible = False
            .fraControlsNoMoreCredit.Visible = False
            .fraControlsPaused.Visible = False
            .fraControlsLocked.Visible = True
            .rtbMessage.FileName = IIf(guCliOpts.intMsgLogoFilesType = vbChecked, guCliOpts.uPaths.strSessionPausedPath, guCliOpts.strAppConfigPath & strDefaultFilePrefix & FILENAME_SESSIONPAUSED)
        
        Case enuWksnState.Blocked_follow
            .fraForm.Visible = False
            .fraHeader.Visible = False
            .fraFooter.Visible = False
            .fraFollowServer.Visible = True
            
        Case enuWksnState.Ready
            Call subResetGUILanguage
            .fraForm.Visible = True
            .fraHeader.Visible = True
            .fraFooter.Visible = True
            .fraFollowServer.Visible = False
            .fraControlsBlocked.Visible = False
            .fraControlsSessionEnd.Visible = False
            .fraControlsActive.Visible = True
            .fraControlsNoMoreCredit.Visible = False
            .fraControlsPaused.Visible = False
            .fraControlsLocked.Visible = False
            .cmdCxn.Default = True
            .rtbMessage.FileName = IIf(guCliOpts.intMsgLogoFilesType = vbChecked, guCliOpts.uPaths.strIntroPath, guCliOpts.strAppConfigPath & strDefaultFilePrefix & FILENAME_INTRO)
            
            Call subSetDefaultUsage
            
            If guCliOpts.intDefaultMode = enuWksnMode.Connected Then
                .optConnected.Value = True
            Else
                .optDisconnected.Value = True
            End If
            .txtCardNum.Text = ""
            .txtLogin.Text = ""
            .txtPassword.Text = ""
            
        Case Else
            Exit Sub
            
        End Select
        .txtCardNum.PasswordChar = IIf(guCliOpts.intHideCardNum, PASSWORD_CHAR, "")
        .imgLogo.Picture = LoadPicture(guCliOpts.uPaths.strLogoPath)
        Call .Show
'        .tmrTopMost.Enabled = True
            
        If Not gblnNoTopmost Then
            Call funToggleTaskSwitching(False)
            Call funToggleCtrlAltDel(False)
            Call funToggleDesktop(False)
            Call funToggleTaskbar(False)
        End If
        
        Select Case guUser.intWksnState
        Case enuWksnState.Ready
            If .optCard.Value Then
                Call funSetFocus(.txtCardNum)
            ElseIf .OptClient.Value Then
                Call funSetFocus(.txtLogin)
            Else
                Call funSetFocus(.optConnected)
            End If
        
        Case enuWksnState.Blocked_Paused
            Call funSetFocus(.cmdUnPause)
        
        Case enuWksnState.Blocked_NoCredit
            Call funSetFocus(.cmdReCnx)
            gdteExpireryTime = SESSION_EXPIRERY_TIME
            .tmrCloseSession.Enabled = True
            
        End Select
    
    End With
End Sub


'*******************************************************************************
Public Sub subUnBlock()
On Error Resume Next

    Call frmBlock.Hide
    If Not gblnNoTopmost Then
        Call funToggleTaskSwitching(True)
        Call funToggleCtrlAltDel(True)
        Call funToggleDesktop(True)
        Call funToggleTaskbar(True)
    End If
End Sub


'*******************************************************************************
Public Sub subShowStatus()
On Error GoTo subShowStatusError

    Call frmMain.Hide
    Call subUnBlock
    frmBlock.tmrTopMost.Enabled = False
    If guUser.intSessionType <> enuSessionType.Free Then
        If Not gblnNoTopmost Then
            Call SetWindowPos(frmStatus.hWnd, HWND_TOPMOST, (frmStatus.sysInfo.WorkAreaWidth - frmStatus.Width) / Screen.TwipsPerPixelX, (frmStatus.sysInfo.WorkAreaHeight - frmStatus.Height) / Screen.TwipsPerPixelY, 0, 0, SWP_NOSIZE)
        Else
            Call SetWindowPos(frmStatus.hWnd, HWND_NOTOPMOST, (frmStatus.sysInfo.WorkAreaWidth - frmStatus.Width) / Screen.TwipsPerPixelX, (frmStatus.sysInfo.WorkAreaHeight - frmStatus.Height) / Screen.TwipsPerPixelY, 0, 0, SWP_NOSIZE)
        End If
        If Not gblnNoStatusWindows Then frmStatus.Show
        'Call funSetFocus(frmStatus.cmdClose)
        Call frmSmallStatus.Hide
        Call frmMicroStatus.Hide
        frmStatus.tmrHideStatus.Enabled = False
        frmStatus.tmrHideStatus.Enabled = True
        frmStatus.cmdCliPassword.Enabled = frmMain.wskCli.State = sckConnected
        frmStatus.cmdChat.Enabled = frmMain.wskCli.State = sckConnected
    Else
        If Not gblnNoStatusWindows Then Call frmMicroStatus.Show
    End If
    Exit Sub

subShowStatusError:
    Call funErrorManager("modFunctions.subShowStatusError")
End Sub


'*******************************************************************************
Public Sub subHideStatus()
On Error GoTo subHideStatusError

    If gblnfrmAbout Or gblnfrmMsgBox Or gblnfrmPasswd Then Exit Sub
    Call frmMain.Hide
    Call frmStatus.Hide
    frmStatus.tmrHideStatus.Enabled = False
    If Not gblnNoStatusWindows Then Call frmSmallStatus.Show
    Exit Sub

subHideStatusError:
    Call funErrorManager("modFunctions.subHideStatusError")
End Sub


'*******************************************************************************
Public Sub subLogOff()
Dim lngTmp As Long, lngRet As Long
On Error GoTo subLogOffError
    
    If funMsg(enuMsg.Logoff) = vbYes Then
        lngRet = ExitWindowsEx(EWX_LOGOFF Or EWX_FORCE, lngTmp)
    End If
    Exit Sub

subLogOffError:
    Call funErrorManager("modFunctions.subLogOffError")
End Sub


'*******************************************************************************
Public Sub subShutdown(intFlag As Integer, Optional strText As String = "")
Dim lngTmp As Long, lngRet As Long
On Error GoTo subShutdownError
    
    If strText = "" Then
        Call subAdjustToken
        lngRet = ExitWindowsEx(intFlag Or EWX_FORCE, lngTmp)
    ElseIf funMsg(enuMsg.Logoff, strText) = vbYes Then
        Call subAdjustToken
        lngRet = ExitWindowsEx(intFlag Or EWX_FORCE, lngTmp)
    End If
    Exit Sub

subShutdownError:
    Call funErrorManager("modFunctions.subShutdownError")
End Sub


'*******************************************************************************
Public Function funStartSession() As Boolean
On Error GoTo funStartSessionError
    
    funStartSession = True
    guCliOpts.uFlags.intNumCnxAttempts = 0
    Select Case guUser.intSessionType
    Case enuSessionType.Card
        frmStatus.fraCard.Visible = True
        frmStatus.fraSelfService.Visible = False
        frmStatus.fraSelfServicePrePaid.Visible = False
        frmStatus.fraClient.Visible = False
        frmSmallStatus.fraCard.Visible = True
        frmSmallStatus.fraSelfService.Visible = False
        frmStatus.lblNumCard = guUser.strID
        frmStatus.fraCard.Height = IIf(guCliOpts.intHideCardNum = vbUnchecked, 975, 615)
        frmStatus.shpLine(2).Height = 480 + frmStatus.fraCard.Height
    
    Case enuSessionType.Client
        frmStatus.fraCard.Visible = False
        frmStatus.fraClient.Visible = True
        frmStatus.fraSelfService.Visible = False
        frmStatus.fraSelfServicePrePaid.Visible = False
        frmSmallStatus.fraCard.Visible = True
        frmSmallStatus.fraSelfService.Visible = False
        frmStatus.lblCliLogin = guUser.strID
        frmStatus.shpLine(2).Height = 480 + frmStatus.fraClient.Height
    
    Case enuSessionType.Selfservice
        frmStatus.fraCard.Visible = False
        frmStatus.fraClient.Visible = False
        frmStatus.fraSelfService.Visible = True
        frmSmallStatus.fraCard.Visible = False
        frmSmallStatus.fraSelfService.Visible = True
        frmStatus.fraSelfServicePrePaid.Visible = False
        frmStatus.lblElapsedTimeSelfService = "00:00:00 "
        frmSmallStatus.lblElapsedTimeSelfService = "00:00:00 "
        frmStatus.lblAmountDue = "0 "
        frmSmallStatus.lblAmountDue = "0 "
'        frmStatus.lblAmountDueLabel = funLoadResString(10165) & guCliOpts.strMonSymbol & funLoadResString(10175)
'        frmSmallStatus.lblAmountDueLabel = funLoadResString(10165) & guCliOpts.strMonSymbol & funLoadResString(10175)
        frmStatus.shpLine(2).Height = 480 + frmStatus.fraSelfService.Height
        
    Case enuSessionType.SelfServicePrePaid
        frmStatus.fraCard.Visible = False
        frmStatus.fraSelfService.Visible = False
        frmStatus.fraSelfServicePrePaid.Visible = True
        frmStatus.fraClient.Visible = False
        
        frmSmallStatus.fraCard.Visible = True
        frmSmallStatus.fraSelfService.Visible = False
        
        frmStatus.shpLine(2).Height = 480 + frmStatus.fraSelfServicePrePaid.Height
        
    Case enuSessionType.Free
    
    End Select
    frmStatus.fraChat.Top = frmStatus.fra.Top + frmStatus.shpLine(2).Height + 120
    frmStatus.fraChatClosed.Top = frmStatus.fraChat.Top
    
    frmStatus.fraChatClosed.Visible = True
    frmStatus.fraChat.Visible = False
    frmStatus.cmdClose.Top = frmStatus.fra.Top + frmStatus.shpLine(2).Height + frmStatus.fraChatClosed.Height + 240
    frmStatus.cmdPause.Top = frmStatus.cmdClose.Top
    frmStatus.cmdQuit.Top = frmStatus.cmdClose.Top
    frmStatus.Height = frmStatus.cmdClose.Top + frmStatus.cmdClose.Height + frmStatus.Height - frmStatus.ScaleHeight + 60

    frmStatus.txtWksnChatHist.Text = ""
    
    Select Case guUser.intWksnState
    Case enuWksnState.UsedConnected, enuWksnState.UsedFree
        frmStatus.optConnected.Value = True
        frmMain.tmrAppKiller.Enabled = False
    
    Case enuWksnState.UsedDisconnected
        frmStatus.optDisconnected.Value = True
        frmMain.tmrAppKiller.Enabled = True
    
    End Select
    
    Set guCurrentURLs = New Collection
    
    Call frmMain.tmrSessionStatus_Timer
    frmMain.tmrSessionStatus.Enabled = True
    frmMain.tmrURLspy.Enabled = True
    Call subSessionCleanup
    Call funCloseAllApps
    
    If guCliOpts.uSessionBatch.intEnableStart = vbChecked Then subRunSessionStartBatch
    
    Call funToggleTaskSwitching(True)
    Call funToggleCtrlAltDel(True)
    
    Call frmMain.Hide
    Call subUnBlock
    frmBlock.tmrTopMost.Enabled = False
    Call subShowStatus
    Exit Function

funStartSessionError:
    Call funErrorManager("modFunctions.funStartSessionError")
End Function


'*******************************************************************************
Public Function funHideStatusWindows() As Boolean
On Error GoTo funHideStatusWindowsError

    funHideStatusWindows = True
    If gblnfrmPasswdChange Then Call Unload(frmPasswdChange)
    Call frmSmallStatus.Hide
    Call frmStatus.Hide
    Call frmMicroStatus.Hide
    frmStatus.tmrHideStatus.Enabled = False
    Exit Function

funHideStatusWindowsError:
    Call funErrorManager("modFunctions.funHideStatusWindowsError")
    funHideStatusWindows = False
End Function


'*******************************************************************************
Public Function funStopSession() As Boolean
On Error GoTo funStopSessionError

    Call frmMain.tmrSessionStatus_Timer
    Call frmMain.tmrWksnStatus_Timer
    frmMain.tmrSessionStatus.Enabled = False
    frmMain.tmrURLspy.Enabled = False
    frmMain.tmrAppKiller.Enabled = False
    
    Call funHideStatusWindows
    
    Call funCloseAllApps
    
    Select Case guUser.intSessionType
    Case enuSessionType.Selfservice
        guUser.intWksnStateOld = guUser.intWksnState
        guUser.intWksnState = enuWksnState.Blocked_SessionEnd
    
    Case Else
        If guCliOpts.uSessionBatch.intEnableEnd = vbChecked Then subRunSessionEndBatch
        Call subUserInit
    
    End Select
    
    Exit Function

funStopSessionError:
    Call funErrorManager("modFunctions.funStopSessionError")
End Function


'*******************************************************************************
Public Sub funCloseAllApps()
Dim uWinList() As typProc
Dim i As Integer
On Error GoTo funCloseAllAppsError

    If gblnfrmMsgBox Then Call Unload(frmMsgBox)
    uWinList = funGetWindowList(frmMain)
    For i = 0 To UBound(uWinList)
        If Not gblnNoKillApp And Not (gblnDebug And UCase(uWinList(i).strName) = "VB6.EXE") Then Call subCloseApp(uWinList(i))
    Next i
    Exit Sub

funCloseAllAppsError:
    Call funErrorManager("modFunctions.funCloseAllAppsError")
End Sub


'*******************************************************************************
Public Function funPauseSession() As Boolean
On Error GoTo funPauseSessionError

    If guUser.intWksnState = enuWksnState.Blocked_Paused Then Exit Function
    funPauseSession = True
    Call funHideStatusWindows
    guUser.intWksnStateOld = guUser.intWksnState
    guUser.intWksnState = enuWksnState.Blocked_Paused
    Exit Function

funPauseSessionError:
    funPauseSession = False
    Call funErrorManager("modFunctions.funPauseSessionError")
End Function


'*******************************************************************************
Public Function funLockSession() As Boolean
On Error GoTo funLockSessionError

    If guUser.intWksnState = enuWksnState.Blocked_Locked Then Exit Function
    funLockSession = True
    Call funHideStatusWindows
    guUser.intWksnStateOld = guUser.intWksnState
    guUser.intWksnState = enuWksnState.Blocked_Locked
    Exit Function

funLockSessionError:
    Call funErrorManager("modFunctions.funLockSessionError")
End Function


'*******************************************************************************
Public Function funFollowSession() As Boolean
On Error GoTo funFollowSessionError

    If guUser.intWksnState = enuWksnState.Blocked_follow Then Exit Function
    funFollowSession = True
    Call funHideStatusWindows
    guUser.intWksnStateOld = guUser.intWksnState
    guUser.intWksnState = enuWksnState.Blocked_follow
    Exit Function

funFollowSessionError:
    Call funErrorManager("modFunctions.funFollowSessionError")
End Function


'*******************************************************************************
Public Function funUnPauseSession() As Boolean
On Error GoTo funUnPauseSession

    funUnPauseSession = True
    guUser.intWksnState = guUser.intWksnStateOld
    Call frmMain.Hide
    Call subUnBlock
    frmBlock.tmrTopMost.Enabled = False
    Call subShowStatus
'    If guUser.intSessionType <> enuSessionType.Free Then Call subShowStatus
    Exit Function

funUnPauseSession:
    Call funErrorManager("modFunctions.funUnPauseSession")
End Function


'*******************************************************************************
Public Sub subInitApp()
On Error GoTo subInitAppError
    
    App.Title = funLoadResString(10000)
    Call subDefineConstants
    
    On Error Resume Next
    guCliOpts.strAppConfigPath = funGetFolderPath(CSIDL_COMMON_APPDATA) & "\" & funLoadResString(10006)
    Call MkDir(guCliOpts.strAppConfigPath)
    guCliOpts.strAppConfigINIPath = guCliOpts.strAppConfigPath & "\" & App.EXEName & ".ini"
    On Error GoTo subInitAppError
    
    Call subGetSystemSettings
    Call subGetUserSettings
    gstrLogfilePath = guCliOpts.uPaths.strLogfilePath
    Call subUserInit
    App.TaskVisible = False
    Set guCliOpts.uConnectedManager = New clsUser
    Exit Sub

subInitAppError:
    Call funErrorManager("modFunctions.subInitAppError")
End Sub


'*******************************************************************************
Public Sub subActivate()
On Error GoTo subActivateError

    If gblnfrmOptions Or gblnfrmPasswdChange Or gblnfrmPasswdClose Or gblnfrmSetupWizard Then Exit Sub
    Call Shell_NotifyIcon(NIM_ADD, guIconData)
    frmMain.tmrActiveTrayIcon.Enabled = True
    guUser.intWksnState = enuWksnState.Ready
    guCliOpts.uFlags.intNumCnxAttempts = 0
    Call subBlock
    frmMain.Hide
    Exit Sub

subActivateError:
    Call funErrorManager("modFunctions.subActivateError")
End Sub


'*******************************************************************************
Public Sub subDefaultState()
On Error GoTo subDefaultStateError

    If gblnfrmOptions Or gblnfrmPasswdChange Or gblnfrmPasswdClose Or gblnfrmSetupWizard Then Exit Sub
    Call Shell_NotifyIcon(NIM_ADD, guIconData)
    frmMain.tmrActiveTrayIcon.Enabled = True
    guUser.intWksnState = guCliOpts.intDefaultState
    Call subBlock
    frmMain.Hide
    Exit Sub

subDefaultStateError:
    Call funErrorManager("modFunctions.subDefaultStateError")
End Sub


'*******************************************************************************
Public Sub subDeactivationPassword(frmParent As Form)
On Error GoTo subDeactivationPasswordError

    Call frmPasswd.Show(vbModal, frmParent)
    Exit Sub

subDeactivationPasswordError:
    Call funErrorManager("modFunctions.subDeactivationPasswordError")
End Sub


'*******************************************************************************
Public Sub subQuitPassword(frmParent As Form)
On Error GoTo subQuitPasswordError

    Call frmStatus.Hide
    Call Load(frmPasswdClose)
    Call SetWindowPos(frmPasswdClose.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    Call frmPasswdClose.Show(vbModal, frmParent)
    Exit Sub

subQuitPasswordError:
    Call funErrorManager("modFunctions.subQuitPasswordError")
End Sub


'*******************************************************************************
Public Sub subLogOffPassword(frmParent As Form)
On Error GoTo subLogOffPasswordError

    Call Load(frmPasswdLogOff)
    Call SetWindowPos(frmPasswdLogOff.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    Call frmPasswdLogOff.Show(vbModal, frmParent)
    Exit Sub

subLogOffPasswordError:
    Call funErrorManager("modFunctions.subLogOffPasswordError")
End Sub


'*******************************************************************************
Public Sub subLock()
On Error GoTo subLockError

    Call funStopSession
    guUser.intWksnState = enuWksnState.Blocked_Deactivated
    Call subBlock
    Exit Sub

subLockError:
    Call funErrorManager("modFunctions.subLockError")
End Sub


'*******************************************************************************
Public Sub subDeactivate()
On Error GoTo subDeactivateError

    Call Unload(frmMsgBox)
    If guUser.intWksnState > enuWksnState.Ready Then Call funStopSession
    frmMain.WindowState = vbNormal
    Call Shell_NotifyIcon(NIM_DELETE, guIconData)
    Call subUnBlock
    Call subResetGUILanguage
    frmBlock.tmrTopMost.Enabled = False
    
    Call frmMain.Show
    frmMain.tmrActiveTrayIcon.Enabled = False
    guUser.intWksnState = enuWksnState.NotActive
    Call SetWindowPos(frmMain.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    Call funSetFocus(frmMain)
    Exit Sub

subDeactivateError:
    Call funErrorManager("modFunctions.subDeactivateError")
End Sub


'*******************************************************************************
Public Sub subConnectToServer()
Dim intErrNum As Integer
On Error GoTo subConnectToServerError

    Call frmMain.wskCli.Close
    frmMain.wskCli.RemoteHost = guCliOpts.uServer.strServerName
    frmMain.wskCli.RemotePort = guCliOpts.uServer.lngPort
    Call frmMain.wskCli.Connect
    gblnCnxOK = True
    frmMain.tmrSrvCnxStatus.Enabled = True
    Exit Sub

subConnectToServerError:
Dim strErrDest As String
    intErrNum = Err.Number
    strErrDest = funErrorManager("modFunctions.subConnectToServerError", False, True)
    If intErrNum >= 10041 And intErrNum <= 10047 Then
        Call funMsg(NetworkConfigError)
        frmMain.tmrSrvCnxStatus.Enabled = False
        frmMain.cmdActiver.Enabled = False
        frmStatus.cmdCliPassword.Enabled = False
        frmStatus.cmdChat.Enabled = False
        If gblnfrmPasswdChange Then Call Unload(frmPasswdChange)
        frmMain.lblConnexionStatus.Caption = funLoadResString(10218)
        guCliOpts.blnStartActive = False
        gblnCnxOK = False
    Else
        Call Err.Raise(CYBERA_ERROR, "modFunctions.subConnectToServerError", strErrDest)
    End If
End Sub


'*******************************************************************************
Public Function funShutdownApplication() As Boolean
On Error GoTo funShutdownApplicationError
    
    funShutdownApplication = True
    If gblnfrmMsgBox Then Call Unload(frmMsgBox)
    If gblnfrmSmallStatus Then Call Unload(frmSmallStatus)
    If gblnfrmStatus Then Call Unload(frmStatus)
    If gblnfrmBlock Then Call Unload(frmBlock)
        
'    Close guCliOpts.intSerializerFile

    Call subSetUserSettings
    Exit Function

funShutdownApplicationError:
    Call funErrorManager("modFunctions.funShutdownApplicationError")
End Function


'*******************************************************************************
Private Sub subGetUserSettings()
'Retreave user specific settings (from registry)
On Error GoTo subGetUserSettingsError
     
     '*Get Main window position
    guCliOpts.uWinPos.intTop = GetRegUserSettings(funLoadResString(10006), "Main window position", "Top", (Screen.Height - 1935) / 2)
    guCliOpts.uWinPos.intLeft = GetRegUserSettings(funLoadResString(10006), "Main window position", "Left", (Screen.Width - 4740) / 2)
    Exit Sub

subGetUserSettingsError:
    Call funErrorManager("modFunctions.subGetUserSettingsError")
End Sub


'*******************************************************************************
Private Sub subSetUserSettings()
On Error GoTo subSetUserSettingsError

    Call SetRegUserSettings(funLoadResString(10006), "Main window position", "Top", guCliOpts.uWinPos.intTop)
    Call SetRegUserSettings(funLoadResString(10006), "Main window position", "Left", guCliOpts.uWinPos.intLeft)
    Exit Sub

subSetUserSettingsError:
    Call funErrorManager("modFunctions.subSetUserSettingsError")
End Sub


'*******************************************************************************
Private Sub subGetSystemSettings()
Dim i As Integer
Dim lngSize As Long
Dim strName As String
Dim strUserName As String * 255
Dim strAppName As String
Dim lngDefaultLangID As Long
Dim strDefaultLang As String
Dim intDefaultLangCharset As Integer
Dim strDefaultFilePrefix As String
Dim uAdapters As typAdapters
On Error GoTo subGetSystemSettingsError
    
    strAppName = funLoadResString(10006)
    strName = Space$(16)
    lngSize = Len(strName)
    Call GetUserName(strUserName, 255)
    guCliOpts.strUserName = LCase(Left$(strUserName, InStr(strUserName, Chr$(0)) - 1))
    guCliOpts.blnWksnTypeSndAln = GetRegSettings(strAppName, "General", "Standalone Workstation", vbChecked, REG_DWORD) = vbChecked
    
    If guCliOpts.blnWksnTypeSndAln Then
        Call GetComputerName(strName, lngSize)
        guCliOpts.strComputerName = LCase(Left(strName, lngSize))
        uAdapters = funGetIPAdapters
        If uAdapters.Count > 0 Then
            For i = 0 To uAdapters.Count - 1
                If uAdapters.Item(i).AdapterTypeID = MIB_IF_TYPE_ETHERNET And uAdapters.Item(i).IpAddressCount > 0 Then
                    If uAdapters.Item(i).IpAddress(0) <> EMPTY_ADDRESS Then
                        guCliOpts.strComputerMacAddress = uAdapters.Item(i).MacAdress
                        Exit For
                    End If
                End If
            Next i
        End If
        'If the Mac address was not found, use the NETBIOS name instead (so that the server will be able to index the workstations)
        If guCliOpts.strComputerMacAddress = "" Then guCliOpts.strComputerMacAddress = guCliOpts.strComputerName
        guCliOpts.strWksnName = Replace(GetRegSettings(strAppName, "General", "Workstation Name", guCliOpts.strComputerName), TIME_TOKEN, "")
    
        guCliOpts.uPaths.strLogfilePath = guCliOpts.strAppConfigPath & "\" & App.EXEName & ".log"
    
    
    Else
        guCliOpts.strComputerName = guCliOpts.strUserName
        guCliOpts.strComputerMacAddress = guCliOpts.strUserName
        guCliOpts.strWksnName = Replace(GetRegUserSettings(strAppName, "General", "Workstation Name", guCliOpts.strComputerName), TIME_TOKEN, "")
    
        guCliOpts.uPaths.strLogfilePath = guCliOpts.strAppConfigPath & "\" & guCliOpts.strUserName & "-" & App.EXEName & ".log"
    
    End If
        
    guCliOpts.uServer.strServerName = GetRegSettings(strAppName, "Server", "Name", "")
    guCliOpts.uServer.lngPort = GetRegSettings(strAppName, "Server", "Port", DEFAULT_TCP_PORT)
    guCliOpts.uServer.intTimeOut = GetRegSettings(strAppName, "Server", "TimeOut", "10")
    guCliOpts.uServer.blnAutoConnect = GetRegSettings(strAppName, "Server", "Automatic Server Connection", vbChecked, REG_DWORD) = vbChecked
    
    guCliOpts.strPasswd = funDoXOR(GetRegSettings(strAppName, "General", "Password", ""), strAppName)
    guCliOpts.uConfigWizard.intSystem = GetRegSettings(strAppName, "General", "Config Wizard", 0)
    guCliOpts.uConfigWizard.intUser = GetRegUserSettings(strAppName, "General", "Config Wizard", 0)
    guCliOpts.intHideCardNum = GetRegSettings(strAppName, "General", "Hide Card Number", vbUnchecked, REG_DWORD)
    guCliOpts.intDefaultState = GetRegSettings(strAppName, "General", "Default State", enuWksnState.Ready)
    guCliOpts.intDefaultMode = GetRegSettings(strAppName, "General", "Default Mode", enuWksnMode.Connected)
    guCliOpts.intDefaultUsage = GetRegSettings(strAppName, "General", "Default Usage", enuSessionType.Card)
    
    guCliOpts.intStartDelay = CInt(GetRegSettings(strAppName, "General", "Start Delay", WKSN_DEFAULT_START_DELAY))
    
    guCliOpts.uPaths.strLogoPath = GetRegSettings(strAppName, "General", "Logo path", guCliOpts.strAppConfigPath & FILENAME_PREFIX & FILENAME_LOGO)
    
    intDefaultLangCharset = 0
    Select Case GetThreadLocale
    Case 2060, 3084, 1036, 5132, 6156, 4108
        strDefaultLang = App.EXEName & LANG_RES_FR
    
    Case 1031, 3079, 5127, 4103, 2055
        strDefaultLang = App.EXEName & LANG_RES_DE
    
    Case 3082, 11274, 16394, 13322, 9226, 5130, 7178, 12298, 17418, 4106, 18442, 2058, 19466, 6154, 15370, 10250, 20490, 14346, 8202
        strDefaultLang = App.EXEName & LANG_RES_ES
    
    Case 1032
        strDefaultLang = App.EXEName & LANG_RES_EL
        intDefaultLangCharset = 160
    
    Case 1040
        strDefaultLang = App.EXEName & LANG_RES_IT
    
    Case 1046
        strDefaultLang = App.EXEName & LANG_RES_BR
    
    Case 1038
        strDefaultLang = App.EXEName & LANG_RES_HU
    
    Case Else
        strDefaultLang = App.EXEName & LANG_RES_EN
    
    End Select
    guCliOpts.strInterfaceLang = GetRegSettings(strAppName, "General", "Interface Language", strDefaultLang)
    guCliOpts.strInterfaceLangDefault = GetRegSettings(strAppName, "General", "Interface Language", strDefaultLang)
    guCliOpts.intInterfaceLangCharset = GetRegSettings(strAppName, "General", "Interface Language Charset", intDefaultLangCharset, REG_DWORD)
    guCliOpts.intInterfaceLangCharsetDefault = GetRegSettings(strAppName, "General", "Interface Language Charset", intDefaultLangCharset, REG_DWORD)
    Select Case guCliOpts.strInterfaceLang
    Case App.EXEName & LANG_RES_EN
        strDefaultFilePrefix = FILENAME_PREFIX_EN
    
    Case App.EXEName & LANG_RES_FR
        strDefaultFilePrefix = FILENAME_PREFIX_FR
    
    Case App.EXEName & LANG_RES_DE
        strDefaultFilePrefix = FILENAME_PREFIX_DE
    
    Case App.EXEName & LANG_RES_ES
        strDefaultFilePrefix = FILENAME_PREFIX_ES
    
    Case App.EXEName & LANG_RES_EL
        strDefaultFilePrefix = FILENAME_PREFIX_EL
    
    Case App.EXEName & LANG_RES_IT
        strDefaultFilePrefix = FILENAME_PREFIX_IT
    
    Case App.EXEName & LANG_RES_BR
        strDefaultFilePrefix = FILENAME_PREFIX_BR
    
    Case App.EXEName & LANG_RES_HU
        strDefaultFilePrefix = FILENAME_PREFIX_HU
    
    End Select
    guCliOpts.uPaths.strIntroPath = GetRegSettings(strAppName, "Message paths", "Intro", guCliOpts.strAppConfigPath & strDefaultFilePrefix & FILENAME_INTRO)
    guCliOpts.uPaths.strBlockedPath = GetRegSettings(strAppName, "Message paths", "Blocked", guCliOpts.strAppConfigPath & strDefaultFilePrefix & FILENAME_BLOCKED)
    guCliOpts.uPaths.strNoCreditPath = GetRegSettings(strAppName, "Message paths", "No credit", guCliOpts.strAppConfigPath & strDefaultFilePrefix & FILENAME_NOCREDIT)
    guCliOpts.uPaths.strSessionEndPath = GetRegSettings(strAppName, "Message paths", "Session end", guCliOpts.strAppConfigPath & strDefaultFilePrefix & FILENAME_SESSIONEND)
    guCliOpts.uPaths.strSessionPausedPath = GetRegSettings(strAppName, "Message paths", "Session paused", guCliOpts.strAppConfigPath & strDefaultFilePrefix & FILENAME_SESSIONPAUSED)
    
    guCliOpts.intMsgLogoFilesType = GetRegSettings(strAppName, "General", "Message & Logo Files Type", vbUnchecked, REG_DWORD)
    
    guCliOpts.blnStartActive = GetRegUserSettings(strAppName, "General", "Start active", vbChecked, REG_DWORD) = vbChecked
    guCliOpts.intAllowLockdown = GetRegUserSettings(strAppName, "General", "Allow Lockdown", vbChecked, REG_DWORD)
    guCliOpts.intStartAtLogon = IIf("" <> GetRegUserSettings("Microsoft\Windows\CurrentVersion", "Run", strAppName, "") Or guCliOpts.uConfigWizard.intSystem = 0, vbChecked, vbUnchecked)
    Exit Sub

subGetSystemSettingsError:
    Call funErrorManager("modFunctions.subGetSystemSettingsError")
End Sub


'*******************************************************************************
Public Sub subSetSystemSettings()
Dim i As Integer
Dim strAppName As String
On Error GoTo subSetSystemSettingsError

    strAppName = funLoadResString(10006)
    
    Call SetRegSettings(strAppName, "General", "Standalone Workstation", IIf(guCliOpts.blnWksnTypeSndAln, vbChecked, vbUnchecked), REG_DWORD)
    If guCliOpts.blnWksnTypeSndAln Then
        Call SetRegSettings(strAppName, "General", "Workstation Name", guCliOpts.strWksnName)
    Else
        Call SetRegUserSettings(strAppName, "General", "Workstation Name", guCliOpts.strWksnName)
    End If
    Call SetRegSettings(strAppName, "Server", "Name", guCliOpts.uServer.strServerName)
    Call SetRegSettings(strAppName, "Server", "Port", guCliOpts.uServer.lngPort)
    Call SetRegSettings(strAppName, "Server", "TimeOut", guCliOpts.uServer.intTimeOut)
    Call SetRegSettings(strAppName, "Server", "Automatic Server Connection", IIf(guCliOpts.uServer.blnAutoConnect, vbChecked, vbUnchecked), REG_DWORD)
    Call SetRegSettings(strAppName, "General", "Interface Language", guCliOpts.strInterfaceLang)
    Call SetRegSettings(strAppName, "General", "Interface Language Charset", guCliOpts.intInterfaceLangCharset, REG_DWORD)
    Call SetRegSettings(strAppName, "General", "Start Delay", guCliOpts.intStartDelay)
    
    Call SetRegUserSettings(strAppName, "General", "Start active", IIf(guCliOpts.blnStartActive, vbChecked, vbUnchecked), REG_DWORD)
    Call SetRegUserSettings(strAppName, "General", "Allow Lockdown", guCliOpts.intAllowLockdown, REG_DWORD)
    
    Select Case guCliOpts.intStartAtLogon
    Case vbChecked
        Call SetRegUserSettings("Microsoft\Windows\CurrentVersion", "Run", strAppName, Chr(34) & App.Path & "\" & App.EXEName & ".exe" & Chr(34))
    Case Else
        Call DelRegUserSettings("Microsoft\Windows\CurrentVersion", "Run", strAppName)
    End Select
    Exit Sub

subSetSystemSettingsError:
    Call funErrorManager("modFunctions.subSetSystemSettingsError")
End Sub


'*******************************************************************************
Public Sub subKillApps()
Dim uProList()  As typProc
Dim i As Long, j As Long, lngNumAppsToKill As Long
Dim k As Integer
Dim blnFoundApp As Boolean
Dim uInet       As clsInetexplorer
Dim uProc       As typProc
On Error GoTo subKillAppsError

    blnFoundApp = False
    uProList = funGetPoccessList
    For i = 0 To UBound(uProList)
        If guCliOpts.uKilledApps.intForbidIE = vbChecked And LCase(uProList(i).strName) = APP_NAME_IE Then
            blnFoundApp = True
            Call subCloseApp(uProList(i))
        End If
        If guCliOpts.uKilledApps.intForbidNetscape = vbChecked And LCase(uProList(i).strName) = APP_NAME_NETSCAPE Then
            blnFoundApp = True
            Call subCloseApp(uProList(i))
        End If
        If guCliOpts.uKilledApps.intForbidMozilla = vbChecked And LCase(uProList(i).strName) = APP_NAME_MOZILLA Then
            blnFoundApp = True
            Call subCloseApp(uProList(i))
        End If
        If guCliOpts.uKilledApps.intForbidMozilla = vbChecked And LCase(uProList(i).strName) = APP_NAME_FIREFOX Then
            blnFoundApp = True
            Call subCloseApp(uProList(i))
        End If
        If guCliOpts.uKilledApps.intForbidMozilla = vbChecked And LCase(uProList(i).strName) = APP_NAME_THUNDERBIRD Then
            blnFoundApp = True
            Call subCloseApp(uProList(i))
        End If
        If guCliOpts.uKilledApps.intForbidOE = vbChecked And LCase(uProList(i).strName) = APP_NAME_OE Then
            blnFoundApp = True
            Call subCloseApp(uProList(i))
        End If
        lngNumAppsToKill = 0
        On Error Resume Next
        lngNumAppsToKill = UBound(guCliOpts.uKilledApps.strKilledApps)
        On Error GoTo subKillAppsError
        If lngNumAppsToKill > 0 Then
            For j = 0 To lngNumAppsToKill
                If LCase(uProList(i).strName) = LCase(guCliOpts.uKilledApps.strKilledApps(j)) Then
                    blnFoundApp = True
                    Call subCloseApp(uProList(i))
                End If
            Next j
        End If
    Next i
    DoEvents
    If guCliOpts.uKilledApps.intForbidIE = vbChecked Then
        Set uInet = New clsInetexplorer
        With uInet
            .Refresh
            For k = 0 To .Count
                If .ItemURL(k) = URL_EMPTY_IE Or .ItemCaption(k) = URL_EMPTY_IE Or InStr(.ItemURL(k), "/") > 0 Or LCase(Left(.ItemURL(k), 4)) = "http" Or LCase(Left(.ItemURL(k), 3)) = "ftp" Then
                    blnFoundApp = True
                    uProc.lngWinHandle = .ItemhWnd(k)
                    Call subCloseApp(uProc)
                End If
            Next k
        End With
        Set uInet = Nothing
    End If
    
    If blnFoundApp Then Call subShowMsgBox(funLoadResString(10000) & funLoadResString(10186) & vbNewLine & funLoadResString(10187))
    Exit Sub

subKillAppsError:
    Call funErrorManager("modFunctions.subKillAppsError")
End Sub


'*******************************************************************************
Public Sub subCheckForRunningWksn()
On Error GoTo subCheckForRunningWksnError

    If App.PrevInstance Then
        Call funMsg(AppAlreadyRunning)
        End
    End If
    Exit Sub

subCheckForRunningWksnError:
    Call funErrorManager("modFunctions.subCheckForRunningWksnError")
End Sub


'*******************************************************************************
Public Sub subStartConnexion(intSessionType As enuSessionType, intWksnState As enuWksnState, Optional strID As String = "", Optional strPasswd As String = "")
Dim lngCardID As Long
Dim strBuffer As String
Dim dteStartTime As Date
Dim dteStartTimer As Date
On Error GoTo subStartConnexionError

    frmBlock.cmdCxn.Enabled = False
    frmBlock.txtCardNum.Enabled = False
    frmBlock.txtLogin.Enabled = False
    frmBlock.txtPassword.Enabled = False
    frmBlock.MousePointer = vbHourglass
    
    
    Call subUserInit
    dteStartTime = Now + guCliOpts.uFlags.dteServerDelta
    If guCliOpts.uServer.intCnxState = enuSrvCxn.Disconnected Then
        frmBlock.tmrActivateTimeout.Enabled = True
        Call subShowMsgBox(funLoadResString(10178) & vbNewLine & funLoadResString(10179), , , , frmBlock, vbModal)
        Exit Sub
    End If
    guCliOpts.uFlags.intNumCnxAttempts = guCliOpts.uFlags.intNumCnxAttempts + 1
    strBuffer = CNX_SESSION_OPEN & intSessionType & CNX_PAYLOAD_TOKEN & strPasswd & CNX_PAYLOAD_TOKEN & strID & CNX_PAYLOAD_TOKEN & intWksnState & CNX_PAYLOAD_TOKEN & CNX_MSG_END
    guCliOpts.uFlags.blnWaitingForReply = True
    dteStartTimer = Now
    If Not funSendData(frmMain.wskCli, strBuffer) Then
        frmBlock.tmrActivateTimeout.Enabled = True
        Call subShowMsgBox(funLoadResString(10178) & vbNewLine & funLoadResString(10179), , , , frmBlock, vbModal)
        Exit Sub
    End If
    Do While guCliOpts.uFlags.blnWaitingForReply
        DoEvents
        If DatePart("s", Now - dteStartTimer) > guCliOpts.uServer.intTimeOut Then
            frmBlock.tmrActivateTimeout.Enabled = True
            Call subShowMsgBox(funLoadResString(10191) & vbNewLine & funLoadResString(10179), , , , frmBlock, vbModal)
            Call subUserInit
            Exit Sub
        End If
    Loop
    frmBlock.tmrActivateTimeout.Enabled = True
    If guCliOpts.uFlags.intServerResponse <> enuOpenSessionReply.OK Then
        Select Case intSessionType
        Case enuSessionType.Card
            frmBlock.txtCardNum.Text = ""
            Call funSetFocus(frmBlock.txtCardNum)
        Case enuSessionType.Client
            frmBlock.txtLogin.Text = ""
            frmBlock.txtPassword.Text = ""
            Call funSetFocus(frmBlock.txtLogin)
        End Select
        
        If guCliOpts.uFlags.intNumCnxAttempts >= MAX_NUM_CNX_ATTEMPTS Then
            Call funSendData(frmMain.wskCli, CNX_MSG_SEND & funLoadResStringDefault(10305) & MAX_NUM_CNX_ATTEMPTS & funLoadResStringDefault(10306) & guCliOpts.strWksnName & vbNewLine & funLoadResStringDefault(10307) & CNX_MSG_END)
            Call subLock
            gdteExpireryTime = ACTIVATION_DELAY_TIME
            frmBlock.tmrUnblock.Enabled = True
        End If
    End If
    Select Case guCliOpts.uFlags.intServerResponse
    Case enuOpenSessionReply.NOK_NoCard
        Call subShowMsgBox(funLoadResString(10104), , , , frmBlock, vbModal)
        Exit Sub
        
    Case enuOpenSessionReply.NOK_CardBalanceEmpty
        Call subShowMsgBox(funLoadResString(10192) & vbNewLine & funLoadResString(10193), , , , frmBlock, vbModal)
        Exit Sub
        
    Case enuOpenSessionReply.NOK_CardDeactivated
        Call subShowMsgBox(funLoadResString(10194) & vbNewLine & funLoadResString(10179), , , , frmBlock, vbModal)
        Exit Sub
        
    Case enuOpenSessionReply.NOK_CardInUse
        Call subShowMsgBox(funLoadResString(10195) & vbNewLine & funLoadResString(10179), , , , frmBlock, vbModal)
        Exit Sub
        
    Case enuOpenSessionReply.NOK_NoClient
        Call subShowMsgBox(funLoadResString(10322), , , , frmBlock, vbModal)
        Exit Sub
        
    Case enuOpenSessionReply.NOK_ClientBalanceEmpty
        Call subShowMsgBox(funLoadResString(10192) & vbNewLine & funLoadResString(10323), , , , frmBlock, vbModal)
        Exit Sub
        
    Case enuOpenSessionReply.NOK_ClientDeactivated
        Call subShowMsgBox(funLoadResString(10324) & vbNewLine & funLoadResString(10179), , , , frmBlock, vbModal)
        Exit Sub
        
    Case enuOpenSessionReply.NOK_ClientInUse
        Call subShowMsgBox(funLoadResString(10325) & vbNewLine & funLoadResString(10179), , , , frmBlock, vbModal)
        Exit Sub
        
    End Select
    
    guUser.intSessionType = intSessionType
    guUser.intWksnState = intWksnState
    guUser.dteStartTime = dteStartTime
    guUser.strID = strID
    guUser.strPassword = strPasswd
    Call subUpdateServer
    Call funStartSession
    Exit Sub

subStartConnexionError:
    Call funErrorManager("modFunctions.subStartConnexionError")
End Sub


'*******************************************************************************
Public Sub subCheckBalance()
Dim strBuffer As String
Dim dteStartTimer As Date
On Error GoTo subCheckBalanceError

    If guCliOpts.uServer.intCnxState = enuSrvCxn.Disconnected Then
        Call subShowMsgBox(funLoadResString(10178) & vbNewLine & funLoadResString(10179), , , , frmBlock, vbModal)
        Exit Sub
    End If
    
    strBuffer = CNX_CHECK_BALANCE & guUser.strID & CNX_MSG_END
    guCliOpts.uFlags.blnWaitingForReply = True
    dteStartTimer = Now
    If Not funSendData(frmMain.wskCli, strBuffer) Then
        Call subShowMsgBox(funLoadResString(10178) & vbNewLine & funLoadResString(10179), , , , frmBlock, vbModal)
        Exit Sub
    End If
    frmBlock.cmdReCnx.Enabled = False
    Screen.MousePointer = vbHourglass
    Do While guCliOpts.uFlags.blnWaitingForReply
        DoEvents
        If DatePart("s", Now - dteStartTimer) > guCliOpts.uServer.intTimeOut Then
            Screen.MousePointer = vbNormal
            Call subShowMsgBox(funLoadResString(10191) & vbNewLine & funLoadResString(10179), , , , frmBlock, vbModal)
            frmBlock.cmdReCnx.Enabled = guCliOpts.uServer.intCnxState = enuSrvCxn.Connected
            Exit Sub
        End If
    Loop
    frmBlock.cmdReCnx.Enabled = guCliOpts.uServer.intCnxState = enuSrvCxn.Connected
    Screen.MousePointer = vbNormal
    Select Case guCliOpts.uFlags.intServerResponse
    Case enuOpenSessionReply.NOK_NoCard
        Call subShowMsgBox(funLoadResString(10104), , , , frmBlock, vbModal)
        Exit Sub
        
    Case enuOpenSessionReply.NOK_CardBalanceEmpty
        Call subShowMsgBox(funLoadResString(10192) & vbNewLine & funLoadResString(10193), , , , frmBlock, vbModal)
        Exit Sub
        
    Case enuOpenSessionReply.NOK_CardDeactivated
        Call subShowMsgBox(funLoadResString(10194) & vbNewLine & funLoadResString(10179), , , , frmBlock, vbModal)
        Exit Sub
    
    Case enuOpenSessionReply.NOK_NoClient
        Call subShowMsgBox(funLoadResString(10322), , , , frmBlock, vbModal)
        Exit Sub
        
    Case enuOpenSessionReply.NOK_ClientBalanceEmpty
        Call subShowMsgBox(funLoadResString(10192) & vbNewLine & funLoadResString(10323), , , , frmBlock, vbModal)
        Exit Sub
        
    Case enuOpenSessionReply.NOK_ClientDeactivated
        Call subShowMsgBox(funLoadResString(10324) & vbNewLine & funLoadResString(10179), , , , frmBlock, vbModal)
        Exit Sub
        
    Case enuOpenSessionReply.NOK_ClientInUse
        Call subShowMsgBox(funLoadResString(10325) & vbNewLine & funLoadResString(10179), , , , frmBlock, vbModal)
        Exit Sub
        
    Case enuOpenSessionReply.NOK_SelfServicePrePaidBalanceEmpty
        Call subShowMsgBox(funLoadResString(10192) & vbNewLine & funLoadResString(10357), , , , frmBlock, vbModal)
        Exit Sub
        
    End Select
    
    frmBlock.tmrCloseSession.Enabled = False
    guUser.intWksnState = guUser.intWksnStateOld
    Call subShowStatus
    Exit Sub

subCheckBalanceError:
    Call funErrorManager("modFunctions.subCheckBalanceError")
End Sub


'*******************************************************************************
Public Sub subCloseCybCli()
On Error GoTo subCloseCybCliError
    
    Call Unload(frmMain)
    End
    Exit Sub

subCloseCybCliError:
    Call funErrorManager("modFunctions.subCloseCybCliError")
End Sub


'*******************************************************************************
Public Sub subUserInit()
On Error GoTo subUserInitError
    
    guUser.strID = ""
    guUser.strPassword = ""
    guUser.lngSessionID = -1
    guUser.sngBalance = 0
    guUser.dteStartTime = NULL_DATE
    guUser.dteBlockedTime = NULL_DATE
    guUser.dteElapsedTime = NULL_DATE
    guUser.dteElapsedTimeAparent = NULL_DATE
    guUser.dteTimeConnected = NULL_DATE
    guUser.dteTimeConnectedOffPeak = NULL_DATE
    guUser.dteTimeDisconnected = NULL_DATE
    guUser.dteRemainingTime = NULL_DATE
    guUser.dteRemainingTimeAparent = NULL_DATE
    guUser.dteInitRemainingTime = NULL_DATE
    guUser.dteInitRemainingTimeSerialized = NULL_DATE
    guUser.dteCustomerExpiryDate = NULL_DATE
    guUser.intSessionType = enuSessionType.Unknown
    guUser.intWksnState = guCliOpts.intDefaultState
    guUser.intWksnStateOld = guCliOpts.intDefaultState
    guUser.blnUnlimitedPlay = False
    
    guCliOpts.uFlags.dteTimer = Now
    guCliOpts.uWksnMsg.lngMsgSelfServiceBalDeltaRatio = 0
    guCliOpts.uFlags.blnNoInetCnxNotified = False
    Exit Sub

subUserInitError:
    Call funErrorManager("modFunctions.subUserInitError")
End Sub


'*******************************************************************************
Public Sub subProcessServerData(strData As String)
Dim strDataBuffer As String
Dim strDatagram As String
Dim strHeader As String
Dim strPayload As String
Dim strTokenizedPayload() As String
Dim lngEndIndex As Long
Dim i As Integer
Dim frmTemp As Form
Dim strDefaultFilePrefix As String
On Error GoTo subProcessServerDataError


    strDataBuffer = strData
    
    Do Until strDataBuffer = ""
        lngEndIndex = InStr(strDataBuffer, CNX_MSG_END)
        If lngEndIndex = 0 Then Exit Do
        strDatagram = Left(strDataBuffer, lngEndIndex - 1)
        strDataBuffer = Mid(strDataBuffer, lngEndIndex + 20)
        strHeader = Left(strDatagram, 20)
        strPayload = Mid(strDatagram, 21)
        Select Case strHeader
        Case CNX_APP_VERSION
            If CLng(strPayload) < SERVER_MIN_VERSION Then
                Call subWriteToLogFile("modFunctions.subProcessServerDataError: Context: " & strHeader & " Payload: " & Left(strPayload, 30))
                Call subDisableWksn(funLoadResString(10263))
                Exit Sub
            End If
        
        Case CNX_SESSION_OK
            strTokenizedPayload = funTokenize(strPayload, CNX_PAYLOAD_TOKEN)
            Select Case CInt(strTokenizedPayload(0))
            Case enuSessionType.Card, enuSessionType.Client
                guUser.dteRemainingTime = funDateTimefromNet(strTokenizedPayload(1))
                guUser.dteCustomerExpiryDate = funDateTimefromNet(strTokenizedPayload(2))
                guUser.blnUnlimitedPlay = CBool(CInt(strTokenizedPayload(3)))
                guUser.dteInitRemainingTime = guUser.dteInitRemainingTime + guUser.dteRemainingTime
            Case enuSessionType.SelfServicePrePaid
                guUser.dteRemainingTime = funDateTimefromNet(strTokenizedPayload(1))
                guUser.dteInitRemainingTime = guUser.dteInitRemainingTime + guUser.dteRemainingTime
            End Select
            guCliOpts.uFlags.intServerResponse = enuOpenSessionReply.OK
            guCliOpts.uFlags.blnWaitingForReply = False
        
        Case CNX_SESSION_NOK
            guCliOpts.uFlags.intServerResponse = CInt(strPayload)
            guCliOpts.uFlags.blnWaitingForReply = False
        
        Case CNX_POSTPAID_MODE
            strTokenizedPayload = funTokenize(strPayload, CNX_PAYLOAD_TOKEN)
            guCliOpts.uPostPaidSs.intMode = CInt(strTokenizedPayload(0))
            guCliOpts.uPostPaidSs.dteBaseTimeIncrement = funDateTimefromNet(strTokenizedPayload(1))
            guCliOpts.uPostPaidSs.dteMinBilledTime = funDateTimefromNet(strTokenizedPayload(2))
            guCliOpts.uPostPaidSs.intSubMode = CInt(strTokenizedPayload(3))
            guCliOpts.uPostPaidSs.dteEndTimeIncrement = funDateTimefromNet(strTokenizedPayload(4))
        
        Case CNX_DEFAULT_RECHARGES
            strTokenizedPayload = funTokenize(strPayload, CNX_PAYLOAD_TOKEN)
            On Error Resume Next
            ReDim guDefaultRecharges(Int(UBound(strTokenizedPayload) / 2))
            If Err.Number <> 0 Then
'                frmBlock.optCard.Value = True
                frmBlock.optSelfService.Enabled = False
            Else
                For i = 0 To UBound(guDefaultRecharges)
                    guDefaultRecharges(i).dteTimeCredit = funDateTimefromNet(strTokenizedPayload(2 * i))
                    guDefaultRecharges(i).sngPrice = funSinglefromNet(strTokenizedPayload(2 * i + 1))
                Next i
            End If
            On Error GoTo subProcessServerDataError
            
        Case CNX_PASSWD_ADMIN
            guCliOpts.strPasswd = strPayload
            Call SetRegSettings(funLoadResString(10006), "General", "Password", funDoXOR(guCliOpts.strPasswd, funLoadResString(10006)))
        
        Case CNX_RATIO_DISCON_CON
            guCliOpts.uTimeRatios.intDisconCon = CInt(strPayload)
        
        Case CNX_CURRENCY_SYMBOL
            strTokenizedPayload = funTokenize(strPayload, CNX_PAYLOAD_TOKEN)
            guCliOpts.strMonSymbol = CStr(strTokenizedPayload(0))
            guCliOpts.strMonFormat = CStr(strTokenizedPayload(1))
'            frmStatus.lblAmountDueLabel = funLoadResString(10165) & guCliOpts.strMonSymbol & funLoadResString(10175)
'            frmSmallStatus.lblAmountDueLabel = funLoadResString(10165) & guCliOpts.strMonSymbol & funLoadResString(10175)
        
        Case CNX_OFFPEAK_PARAMS
            strTokenizedPayload = funTokenize(strPayload, CNX_PAYLOAD_TOKEN)
            guCliOpts.uTimeRatios.intOffPeak = CInt(strTokenizedPayload(0))
            guCliOpts.uTimeRatios.dteOffPeakStop = funDateTimefromNet(strTokenizedPayload(1))
            guCliOpts.uTimeRatios.dteOffPeakStart = funDateTimefromNet(strTokenizedPayload(2))
        
        Case CNX_ACTIVATE_WORKSTATION
            If guUser.intWksnState = enuWksnState.Blocked_SessionEnd And guCliOpts.uSessionBatch.intEnableEnd = vbChecked Then subRunSessionEndBatch
            Call subActivate
        
        Case CNX_LOCK_WORKSTATION
            If guUser.intWksnState = enuWksnState.Blocked_SessionEnd And guCliOpts.uSessionBatch.intEnableEnd = vbChecked Then subRunSessionEndBatch
            Call subLock
        
        Case CNX_SESSION_CLOSE
            Call funStopSession
            Call subBlock
        
        Case CNX_FORCE_CONNECTED
            Select Case strPayload
            Case vbChecked
                frmBlock.optConnected.Value = True
                frmStatus.optConnected.Value = True
                frmBlock.optDisconnected.Enabled = False
                frmStatus.optDisconnected.Enabled = False
                If guUser.intWksnState = enuWksnState.UsedDisconnected Then
                    Call subShowMsgBox(funLoadResString(10196) & vbNewLine & funLoadResString(10197))
                    guUser.intWksnState = enuWksnState.UsedConnected
                    frmMain.tmrAppKiller.Enabled = False
                End If
                
            Case Else
                frmBlock.optDisconnected.Enabled = True
                frmStatus.optDisconnected.Enabled = True
            
            End Select
        
        Case CNX_WORKSTATION_MSG_OPTIONS
            strTokenizedPayload = funTokenize(strPayload, CNX_PAYLOAD_TOKEN)
            guCliOpts.uWksnMsg.intMsgCard = CInt(strTokenizedPayload(0))
            guCliOpts.uWksnMsg.intMsgCardAlert1 = CInt(strTokenizedPayload(1))
            guCliOpts.uWksnMsg.intMsgCardAlert2 = CInt(strTokenizedPayload(2))
            guCliOpts.uWksnMsg.sngMsgSelfServiceBalance = funSinglefromNet(strTokenizedPayload(4))
            guCliOpts.uWksnMsg.intMsgSelfService = IIf(guCliOpts.uWksnMsg.sngMsgSelfServiceBalance <> 0, CInt(strTokenizedPayload(3)), vbUnchecked)
        
        Case CNX_ACTIVATE_SELFSERVICE
            frmBlock.optSelfService.Enabled = strPayload = vbChecked
            Call subSetDefaultUsage
            If guUser.intSessionType = enuSessionType.Selfservice And strPayload = vbUnchecked Then
                Call funStopSession
                Call subBlock
            End If
        
        Case CNX_ACTIVATE_CARDS
            frmBlock.optCard.Enabled = strPayload = vbChecked
            Call subSetDefaultUsage
            If guUser.intSessionType = enuSessionType.Card And strPayload = vbUnchecked Then
                Call funStopSession
                Call subBlock
            ElseIf strPayload = vbUnchecked Then
               frmBlock.fraCardNumber.Visible = False
            End If
        
        Case CNX_ACTIVATE_CLIENTS
            frmBlock.OptClient.Enabled = strPayload = vbChecked
            Call subSetDefaultUsage
            If guUser.intSessionType = enuSessionType.Client And strPayload = vbUnchecked Then
                Call funStopSession
                Call subBlock
            ElseIf strPayload = vbUnchecked Then
               frmBlock.fraClient.Visible = False
            End If
        
        Case CNX_SHUTDOWN_WORKSTATION
            Call subShutdown(CInt(strPayload))
        
        Case CNX_RECHARGE_CARD
            strTokenizedPayload = funTokenize(strPayload, CNX_PAYLOAD_TOKEN)
            guUser.dteRemainingTime = guUser.dteRemainingTime + funDateTimefromNet(strTokenizedPayload(0))
            guUser.dteCustomerExpiryDate = funDateTimefromNet(strTokenizedPayload(1))
            guUser.dteInitRemainingTime = guUser.dteInitRemainingTime + funDateTimefromNet(strTokenizedPayload(0))
        
        Case CNX_PAUSE_WORKSTATION
            Call funPauseSession
            Call subBlock
        
        Case CNX_PAUSE_LOCK_WORKSTATION
            Call funLockSession
            Call subBlock
        
        Case CNX_PAUSE_FOLLOW_WORKSTATION
            Call funFollowSession
            Call subBlock
        
        Case CNX_UNPAUSE_WORKSTATION
            Call funUnPauseSession
        
        Case CNX_RESUME_WORKSTATION
            Call subCheckBalance
        
        Case CNX_NOTIFY_WKSN_INETCNX
            If guUser.intWksnState = enuWksnState.Blocked_Paused And Not gblnfrmMsgBox Then
                Call subShowMsgBox(funLoadResString(10207) & vbNewLine & funLoadResString(10208), YesNo, Button1, New clsActionUnPauseSession)
            End If
            guCliOpts.uFlags.blnNoInetCnxNotified = False
        
        Case CNX_NOTIFY_WKSN_NO_INETCNX
            If guUser.intWksnState = enuWksnState.UsedConnected And Not guCliOpts.uFlags.blnNoInetCnxNotified And Not gblnfrmMsgBox Then
                guCliOpts.uFlags.blnNoInetCnxNotified = True
                Call subShowMsgBox(funLoadResString(10205) & vbNewLine & funLoadResString(10206), YesNo, Button2, New clsActionPauseSession)
            End If
        
        Case CNX_DISABLE_WORKSTATION
            Select Case CInt(strPayload)
            Case enuWksnDisableReason.VersionNOK
                Call subDisableWksn(funLoadResString(10264))
            Case enuWksnDisableReason.InternalError
                Call subDisableWksn(funLoadResString(10265))
            End Select
        
        Case CNX_SESSION_OPEN
            strTokenizedPayload = funTokenize(strPayload, CNX_PAYLOAD_TOKEN)
            Call subStartConnexion(CInt(strTokenizedPayload(0)), IIf(guCliOpts.intDefaultMode = enuWksnMode.Connected, enuWksnState.UsedConnected, enuWksnState.UsedDisconnected), strTokenizedPayload(2), strTokenizedPayload(1))
                
        Case CNX_FILE_SEND
            strTokenizedPayload = funTokenize(strPayload, CNX_PAYLOAD_TOKEN)
            Call subProcessReceivedFile(CInt(strTokenizedPayload(0)), strTokenizedPayload(1))
                
        Case CNX_SESSION_FREE
            Call subUserInit
            guUser.intSessionType = enuSessionType.Free
            guUser.intWksnState = enuWksnState.UsedFree
            Call funStartSession
                
        Case CNX_MSG_SEND
            Call subShowMsgBox(strPayload, OKOnly, , New clsActionEmpty)
        
        Case CNX_KILLED_APPS
            strTokenizedPayload = funTokenize(strPayload, CNX_PAYLOAD_TOKEN)
            guCliOpts.uKilledApps.intForbidIE = CInt(strTokenizedPayload(0))
            guCliOpts.uKilledApps.intForbidMozilla = CInt(strTokenizedPayload(1))
            guCliOpts.uKilledApps.intForbidNetscape = CInt(strTokenizedPayload(2))
            guCliOpts.uKilledApps.intForbidOE = CInt(strTokenizedPayload(3))
            guCliOpts.uKilledApps.strKilledApps = funTokenize(strTokenizedPayload(3), FILENAME_TOKEN)
        
        Case CNX_SESSION_CLEANUP
            strTokenizedPayload = funTokenize(strPayload, CNX_PAYLOAD_TOKEN)
            guCliOpts.uSessionCleanup.intAlignDesktopIcons = CInt(strTokenizedPayload(0))
            guCliOpts.uSessionCleanup.intDeleteDocHistory = CInt(strTokenizedPayload(1))
            guCliOpts.uSessionCleanup.intRestoreBackground = CInt(strTokenizedPayload(2))
            guCliOpts.uSessionCleanup.intRestoreIEHomePage = CInt(strTokenizedPayload(3))
            guCliOpts.uSessionCleanup.intRestoreIESettings = CInt(strTokenizedPayload(4))
            guCliOpts.uSessionCleanup.strIEHomePagePath = strTokenizedPayload(5)
            guCliOpts.uSessionCleanup.intRestoreIMs = CInt(strTokenizedPayload(6))
        
        Case CNX_SESSION_ID
            guUser.lngSessionID = CLng(strPayload)
        
        Case CNX_HIDE_CARD_NUM
            guCliOpts.intHideCardNum = CInt(strPayload)
            Call SetRegSettings(funLoadResString(10006), "General", "Hide Card Number", guCliOpts.intHideCardNum)
            frmBlock.txtCardNum.PasswordChar = IIf(guCliOpts.intHideCardNum = vbChecked, PASSWORD_CHAR, "")
            frmStatus.fraCardNum.Visible = guCliOpts.intHideCardNum = vbUnchecked
        
        Case CNX_ENABLE_PAUSE
            frmSmallStatus.cmdPause.Visible = CInt(strPayload) = vbChecked
            frmStatus.cmdPause.Visible = CInt(strPayload) = vbChecked
            If CInt(strPayload) = vbChecked Then
                frmSmallStatus.cmdQuit.Width = 1155
                frmSmallStatus.cmdQuit.Left = 2460
                frmSmallStatus.cmdStatus.Width = 1155
                
                frmStatus.cmdQuit.Width = 1155
                frmStatus.cmdQuit.Left = 2460
                frmStatus.cmdClose.Width = 1155
            Else
                frmSmallStatus.cmdQuit.Width = 1755
                frmSmallStatus.cmdQuit.Left = 1860
                frmSmallStatus.cmdStatus.Width = 1755
                
                frmStatus.cmdQuit.Width = 1755
                frmStatus.cmdQuit.Left = 1860
                frmStatus.cmdClose.Width = 1755
            End If
        
        Case CNX_SERVER_DATE
            guCliOpts.uFlags.dteServerDelta = funDateTimefromNet(strPayload) - Now
            If guCliOpts.uFlags.blnWaitingForServerDate Then
                If guUser.intWksnState > enuWksnState.Ready And guUser.intWksnState <> enuWksnState.Blocked_NoCredit Then Call funUnPauseSession
                guCliOpts.uFlags.blnWaitingForServerDate = False
            End If
        
        Case CNX_SESSION_BATCH_START
            guCliOpts.uSessionBatch.intEnableStart = CInt(strPayload)
        
        Case CNX_SESSION_BATCH_END
            guCliOpts.uSessionBatch.intEnableEnd = CInt(strPayload)
        
        Case CNX_SESSION_BATCH_STYLE
            guCliOpts.uSessionBatch.intWindowStyle = CInt(strPayload)
        
        Case CNX_WKSN_DEFAULT_STATE
            guCliOpts.intDefaultState = CInt(strPayload)
            Call SetRegSettings(funLoadResString(10006), "General", "Default State", guCliOpts.intDefaultState)
        
        Case CNX_WKSN_DEFAULT_USAGE
            guCliOpts.intDefaultUsage = CInt(strPayload)
            Call SetRegSettings(funLoadResString(10006), "General", "Default Usage", guCliOpts.intDefaultUsage)
            Call subSetDefaultUsage
        
        Case CNX_WKSN_DEFAULT_MODE
            guCliOpts.intDefaultMode = CInt(strPayload)
            Call SetRegSettings(funLoadResString(10006), "General", "Default Mode", guCliOpts.intDefaultMode)
            If guCliOpts.intDefaultMode = enuWksnMode.Connected Then
                frmBlock.optConnected.Value = True
            Else
                frmBlock.optDisconnected.Value = True
            End If
        
        Case CNX_DEFAULT_LANG
            strTokenizedPayload = funTokenize(strPayload, CNX_PAYLOAD_TOKEN)
            guCliOpts.strInterfaceLang = App.EXEName & CStr(strTokenizedPayload(0))
            guCliOpts.strInterfaceLangDefault = App.EXEName & CStr(strTokenizedPayload(0))
            guCliOpts.intInterfaceLangCharset = CInt(strTokenizedPayload(1))
            guCliOpts.intInterfaceLangCharsetDefault = CInt(strTokenizedPayload(1))
            For Each frmTemp In Forms
                Call frmTemp.subLocalizeInterface
            Next frmTemp
            Call SetRegSettings(funLoadResString(10006), "General", "Interface Language", guCliOpts.strInterfaceLang)
            Call SetRegSettings(funLoadResString(10006), "General", "Interface Language Charset", guCliOpts.intInterfaceLangCharset, REG_DWORD)
        
        
        Case CNX_MSG_LOGO_FILES_TYPE
            guCliOpts.intMsgLogoFilesType = CInt(strPayload)
            Call SetRegSettings(funLoadResString(10006), "General", "Message & Logo Files Type", guCliOpts.intMsgLogoFilesType)
            If gblnfrmBlock And frmBlock.Visible Then
                strDefaultFilePrefix = FILENAME_PREFIX & Right(guCliOpts.strInterfaceLang, Len(guCliOpts.strInterfaceLang) - Len(App.EXEName) - 1) & "_"
                Select Case guUser.intWksnState
                Case enuWksnState.Blocked_Deactivated
                    frmBlock.rtbMessage.FileName = IIf(guCliOpts.intMsgLogoFilesType = vbChecked, guCliOpts.uPaths.strBlockedPath, guCliOpts.strAppConfigPath & strDefaultFilePrefix & FILENAME_BLOCKED)
                
                Case enuWksnState.Blocked_SessionEnd
                    frmBlock.rtbMessage.FileName = IIf(guCliOpts.intMsgLogoFilesType = vbChecked, guCliOpts.uPaths.strSessionEndPath, guCliOpts.strAppConfigPath & strDefaultFilePrefix & FILENAME_SESSIONEND)
                
                Case enuWksnState.Blocked_NoCredit
                    frmBlock.rtbMessage.FileName = IIf(guCliOpts.intMsgLogoFilesType = vbChecked, guCliOpts.uPaths.strNoCreditPath, guCliOpts.strAppConfigPath & strDefaultFilePrefix & FILENAME_NOCREDIT)
                
                Case enuWksnState.Blocked_Paused, enuWksnState.Blocked_Locked
                    frmBlock.rtbMessage.FileName = IIf(guCliOpts.intMsgLogoFilesType = vbChecked, guCliOpts.uPaths.strSessionPausedPath, guCliOpts.strAppConfigPath & strDefaultFilePrefix & FILENAME_SESSIONPAUSED)
                
                Case enuWksnState.Ready
                    frmBlock.rtbMessage.FileName = IIf(guCliOpts.intMsgLogoFilesType = vbChecked, guCliOpts.uPaths.strIntroPath, guCliOpts.strAppConfigPath & strDefaultFilePrefix & FILENAME_INTRO)
                
                End Select
            End If
        
        Case CNX_PASSWD_MANGR
            strTokenizedPayload = funTokenize(strPayload, CNX_PAYLOAD_TOKEN)
            guCliOpts.uConnectedManager.login = strTokenizedPayload(0)
            If UBound(strTokenizedPayload) > 0 Then
                guCliOpts.uConnectedManager.Passwd = strTokenizedPayload(1)
            Else
                guCliOpts.uConnectedManager.Passwd = ""
            End If
                
        Case CNX_CHAT_MSG
            If guUser.intWksnState >= enuWksnState.UsedDisconnected Then
                frmStatus.txtWksnChatHist.Text = frmStatus.txtWksnChatHist.Text & IIf(frmStatus.txtWksnChatHist.Text <> "", vbNewLine, "") & " " & funLoadResString(10350) & ": " & strPayload
                Call frmStatus.cmdChatOpen_Click
                If Not frmStatus.Visible Then
                    Call subShowStatus
                Else
                    frmStatus.tmrHideStatus.Enabled = False
                    frmStatus.tmrHideStatus.Enabled = True
                End If
                Call funSetFocus(frmStatus.txtWksnChat)
                frmStatus.txtWksnChatHist.SelStart = Len(frmStatus.txtWksnChatHist.Text)
            End If
        
        Case CNX_EMPTY_SERIALIZED_USER
            Call subEmptySerializedUser
        
        End Select
    Loop
    Exit Sub

subProcessServerDataError:
Dim strErrDesc As String
    strErrDesc = funErrorManager("modFunctions.subProcessServerDataError", , True)
    Call subDebugPrint("modFunctions.subProcessServerDataError: Context: " & strHeader & " Payload: " & Left(strPayload, 30))
    Call subWriteToLogFile("modFunctions.subProcessServerDataError: Context: " & strHeader & " Payload: " & Left(strPayload, 30))
    Call subDisableWksn(funLoadResString(10265))
    Call Err.Raise(CYBERA_ERROR, "modFunctions.subProcessServerDataError", strErrDesc)
End Sub


'*******************************************************************************
Public Function funBalanceDue(dteSessionTime As Date, ByRef dteTimeToNextIncrement As Date) As Single
On Error GoTo funBalanceDueError

    If guCliOpts.uPostPaidSs.intMode = enuPostpaidSsMode.RechargeBased Then
        If guCliOpts.uPostPaidSs.intSubMode = enuPostpaidSsSubMode.RechAutoOptim Then
            funBalanceDue = funBalanceDueRechargeBasedAutoOptim(dteSessionTime, dteTimeToNextIncrement)
        Else
            funBalanceDue = funBalanceDueRechargeBasedSimple(dteSessionTime, dteTimeToNextIncrement)
        End If
    Else
        funBalanceDue = funBalanceDueTimeBased(dteSessionTime, dteTimeToNextIncrement)
    End If
    Exit Function

funBalanceDueError:
    Call funErrorManager("modFunctions.funBalanceDueError")
End Function


'*******************************************************************************
Private Function funBalanceDueRechargeBasedAutoOptim(dteSessionTime As Date, ByRef dteTimeToNextIncrement As Date) As Single
Dim i As Integer, j As Integer
Dim intNumRecharges() As Integer      ' Array of the quantity of each recharge to be billed
Dim dteRemainder As Date              ' Remaining time to be billed
Dim sngBalanceDue As Single
Dim sngBalanceRemainning As Single
On Error GoTo funBalanceDueRechargeBasedAutoOptimError

    ReDim intNumRecharges(UBound(guDefaultRecharges))
    dteRemainder = dteSessionTime
    
    ' First parse the list or recharges from largest to smallest to determine how many
    ' recharges are needed to fit in the SessionTime
    For i = UBound(guDefaultRecharges) To LBound(guDefaultRecharges) Step -1
        intNumRecharges(i) = Int(dteRemainder / guDefaultRecharges(i).dteTimeCredit)
        dteRemainder = dteRemainder - intNumRecharges(i) * guDefaultRecharges(i).dteTimeCredit
    Next i
    
    ' If there is still time left to be billed add one more or the smallest recharge
    If dteRemainder <> 0 Then
        intNumRecharges(0) = intNumRecharges(0) + 1
    End If
    
    ' Calculate the balace due
    For i = UBound(intNumRecharges) To LBound(intNumRecharges) Step -1
        sngBalanceDue = sngBalanceDue + intNumRecharges(i) * guDefaultRecharges(i).sngPrice
        sngBalanceRemainning = 0
        For j = i - 1 To LBound(intNumRecharges) Step -1
            sngBalanceRemainning = sngBalanceRemainning + intNumRecharges(j) * guDefaultRecharges(j).sngPrice
        Next j
        If sngBalanceRemainning > guDefaultRecharges(i).sngPrice Then
            sngBalanceDue = sngBalanceDue + guDefaultRecharges(i).sngPrice
            Exit For
        End If
    Next i
    funBalanceDueRechargeBasedAutoOptim = sngBalanceDue
    Exit Function

funBalanceDueRechargeBasedAutoOptimError:
    Call funErrorManager("modFunctions.funBalanceDueRechargeBasedAutoOptimError")
End Function


'*******************************************************************************
Private Function funBalanceDueRechargeBasedSimple(dteSessionTime As Date, ByRef dteTimeToNextIncrement As Date) As Single
Dim i As Integer
Dim dteRemainder As Date              ' Remaining time to be billed
Dim dteCurrentTime As Date
Dim sngBalanceDue As Single
On Error GoTo funBalanceDueRechargeBasedSimpleError
    
    If dteSessionTime < guDefaultRecharges(UBound(guDefaultRecharges)).dteTimeCredit Then
        For i = LBound(guDefaultRecharges) To UBound(guDefaultRecharges)
            If dteSessionTime <= guDefaultRecharges(i).dteTimeCredit Then Exit For
        Next i
        sngBalanceDue = guDefaultRecharges(i).sngPrice
        dteTimeToNextIncrement = guDefaultRecharges(i).dteTimeCredit - dteSessionTime
    Else
        sngBalanceDue = guDefaultRecharges(UBound(guDefaultRecharges)).sngPrice
        dteRemainder = dteSessionTime - guDefaultRecharges(UBound(guDefaultRecharges)).dteTimeCredit
        If dteRemainder <> NULL_DATE Then
            sngBalanceDue = sngBalanceDue + -Int(-(dteRemainder / guCliOpts.uPostPaidSs.dteBaseTimeIncrement)) * guDefaultRecharges(UBound(guDefaultRecharges)).sngPrice / (guDefaultRecharges(UBound(guDefaultRecharges)).dteTimeCredit / guCliOpts.uPostPaidSs.dteBaseTimeIncrement)
            dteRemainder = dteRemainder - guDefaultRecharges(UBound(guDefaultRecharges)).dteTimeCredit / (guDefaultRecharges(UBound(guDefaultRecharges)).dteTimeCredit / guCliOpts.uPostPaidSs.dteBaseTimeIncrement) * Int(dteRemainder / guCliOpts.uPostPaidSs.dteBaseTimeIncrement)
            dteTimeToNextIncrement = guCliOpts.uPostPaidSs.dteBaseTimeIncrement - dteRemainder
        End If
    End If
    
    Select Case guUser.intWksnState
    Case enuWksnState.UsedConnected, enuWksnState.UsedFree
        If funIsOffPeak(dteCurrentTime) Then
            dteTimeToNextIncrement = dteTimeToNextIncrement / (guCliOpts.uTimeRatios.intOffPeak / 100)
        End If
    Case enuWksnState.UsedDisconnected
        dteTimeToNextIncrement = dteTimeToNextIncrement / (guCliOpts.uTimeRatios.intDisconCon / 100)
    End Select
    
    funBalanceDueRechargeBasedSimple = sngBalanceDue
    Exit Function

funBalanceDueRechargeBasedSimpleError:
    Call funErrorManager("modFunctions.funBalanceDueRechargeBasedSimpleError")
End Function


'*******************************************************************************
Private Function funBalanceDueTimeBased(dteSessionTime As Date, ByRef dteTimeToNextIncrement As Date) As Single
Dim i As Integer
Dim dteRemainder As Date              ' Remaining time to be billed
Dim sngBalanceDue As Single
Dim dteCurrentTime As Date
On Error GoTo funBalanceDueTimeBasedError

    dteRemainder = dteSessionTime
    
    If dteSessionTime <= guCliOpts.uPostPaidSs.dteMinBilledTime Then
        sngBalanceDue = guDefaultRecharges(i).sngPrice / (guDefaultRecharges(i).dteTimeCredit / guCliOpts.uPostPaidSs.dteMinBilledTime)
        dteTimeToNextIncrement = guCliOpts.uPostPaidSs.dteMinBilledTime - dteRemainder
    Else
        For i = LBound(guDefaultRecharges) To UBound(guDefaultRecharges)
            If dteRemainder > guDefaultRecharges(i).dteTimeCredit Then
                sngBalanceDue = sngBalanceDue + guDefaultRecharges(i).sngPrice
                dteRemainder = dteRemainder - guDefaultRecharges(i).dteTimeCredit
            Else
                Exit For
            End If
        Next i
        If i > UBound(guDefaultRecharges) Then i = UBound(guDefaultRecharges)
        If dteRemainder <> NULL_DATE Then
            sngBalanceDue = sngBalanceDue + -Int(-(dteRemainder / guCliOpts.uPostPaidSs.dteBaseTimeIncrement)) * guDefaultRecharges(i).sngPrice / (guDefaultRecharges(i).dteTimeCredit / guCliOpts.uPostPaidSs.dteBaseTimeIncrement)
            dteRemainder = dteRemainder - guDefaultRecharges(i).dteTimeCredit / (guDefaultRecharges(i).dteTimeCredit / guCliOpts.uPostPaidSs.dteBaseTimeIncrement) * Int(dteRemainder / guCliOpts.uPostPaidSs.dteBaseTimeIncrement)
            dteTimeToNextIncrement = guCliOpts.uPostPaidSs.dteBaseTimeIncrement - dteRemainder
        End If
    End If
    
    dteCurrentTime = CDate(Format(Now + guCliOpts.uFlags.dteServerDelta, "Hh:Nn:Ss"))
    Select Case guUser.intWksnState
    Case enuWksnState.UsedConnected, enuWksnState.UsedFree
        If funIsOffPeak(dteCurrentTime) Then
            dteTimeToNextIncrement = dteTimeToNextIncrement / (guCliOpts.uTimeRatios.intOffPeak / 100)
        End If
    Case enuWksnState.UsedDisconnected
        dteTimeToNextIncrement = dteTimeToNextIncrement / (guCliOpts.uTimeRatios.intDisconCon / 100)
    End Select
    
    funBalanceDueTimeBased = sngBalanceDue
    Exit Function

funBalanceDueTimeBasedError:
    Call funErrorManager("modFunctions.funBalanceDueTimeBasedError")
End Function


'*******************************************************************************
Public Sub subShowMsgBox(strData As String, Optional intType As enuMessageBoxType = OKOnly, Optional intDefaultButton As enuMessageBoxDefaultButton = Button1, Optional uAction As Object, Optional frmParent As Form, Optional intMode As FormShowConstants = vbModeless, Optional blnPriority As Boolean = False)
Static blnLastPriority As Boolean
Static strLastData As String
On Error GoTo subShowMsgBoxError

    If Not gblnfrmMsgBox Or strLastData <> strData Then Beep
    If gblnfrmPasswdChange Then
        Call Unload(frmPasswdChange)
        DoEvents
    End If
    If gblnfrmPasswd Then
        Call Unload(frmPasswd)
        DoEvents
    End If
    Call Load(frmMsgBox)
    If frmMsgBox.Visible And blnLastPriority Then Exit Sub
    blnLastPriority = blnPriority
    strLastData = strData
    frmMsgBox.tmrTimeOut.Enabled = False
    frmMsgBox.tmrTimeOut.Enabled = True
    frmMsgBox.lblText.Caption = strData
    frmMsgBox.fraOK.Visible = intType = OKOnly
    frmMsgBox.fraYesNo.Visible = intType = YesNo
    If Not IsEmpty(frmParent) Then Set frmMsgBox.mfrmParent = frmParent
    Select Case intType
    Case OKOnly
        frmMsgBox.cmdOK.Default = True
        frmMsgBox.cmdOK.Cancel = True
    Case YesNo
        frmMsgBox.cmdYes.Default = intDefaultButton = Button1
        frmMsgBox.cmdYes.Cancel = intDefaultButton = Button1
        frmMsgBox.cmdNo.Default = intDefaultButton = Button2
        frmMsgBox.cmdNo.Cancel = intDefaultButton = Button2
        Set frmMsgBox.uAction = uAction
    End Select
    If Not frmParent Is Nothing Then
        Call frmMsgBox.Show(intMode, frmParent)
    Else
        Call frmMsgBox.Show
        Select Case intType
        Case OKOnly
            Call funSetFocus(frmMsgBox.cmdOK)
        Case YesNo
            If frmMsgBox.cmdYes.Default Then
                Call funSetFocus(frmMsgBox.cmdYes)
            Else
                Call funSetFocus(frmMsgBox.cmdNo)
            End If
        End Select
    End If
    Exit Sub

subShowMsgBoxError:
    Call funErrorManager("modFunctions.subShowMsgBoxError")
End Sub


'*******************************************************************************
Public Sub subUpdateServer()
Dim strBuffer As String
On Error GoTo subUpdateServerError

    If frmMain.wskCli.State <> sckConnected Then Exit Sub
    
    strBuffer = CNX_WORKSTATION_STATE & guUser.intWksnState & CNX_MSG_END
    strBuffer = strBuffer & CNX_SESSION_ID & guUser.lngSessionID & CNX_MSG_END
    strBuffer = strBuffer & CNX_SESSION_TYPE & guUser.intSessionType & CNX_MSG_END
    strBuffer = strBuffer & CNX_ELAPSED_TIME & funDateTimeToSend(guUser.dteElapsedTimeAparent) & CNX_MSG_END
    strBuffer = strBuffer & CNX_USER_ID & guUser.strID & CNX_MSG_END
    Select Case guUser.intSessionType
    Case enuSessionType.Card, enuSessionType.Client, enuSessionType.SelfServicePrePaid
        strBuffer = strBuffer & CNX_REMAINING_TIME & funDateTimeToSend(guUser.dteRemainingTime) & CNX_MSG_END
        strBuffer = strBuffer & CNX_REMAINING_TIME_APARENT & funDateTimeToSend(guUser.dteRemainingTimeAparent) & CNX_MSG_END
    
    Case enuSessionType.Selfservice
        strBuffer = strBuffer & CNX_USER_BALANCE & funSingleToSend(guUser.sngBalance) & CNX_MSG_END
    
    End Select
    strBuffer = strBuffer & CNX_CONNECTED_TIME & funDateTimeToSend(guUser.dteTimeConnected) & CNX_MSG_END
    strBuffer = strBuffer & CNX_CONNECTED_TIME_OFFPEAK & funDateTimeToSend(guUser.dteTimeConnectedOffPeak) & CNX_MSG_END
    strBuffer = strBuffer & CNX_DISCONNECTED_TIME & funDateTimeToSend(guUser.dteTimeDisconnected) & CNX_MSG_END
    Call funSendData(frmMain.wskCli, strBuffer)
    Exit Sub

subUpdateServerError:
    Call funErrorManager("modFunctions.subUpdateServerError")
End Sub


'*******************************************************************************
Public Sub subInitServer()
Dim strBuffer As String
Dim intAllowLockdown As Integer
On Error GoTo subInitServerError

    If frmMain.wskCli.State <> sckConnected Then Exit Sub
    intAllowLockdown = guCliOpts.intAllowLockdown
    intAllowLockdown = IIf(guCliOpts.uConfigWizard.intSystem < CONFIG_WIZARD_CLI_SYSTEM Or guCliOpts.uConfigWizard.intUser < CONFIG_WIZARD_CLI_USER Or gblnNoLockdown, vbUnchecked, intAllowLockdown)
    strBuffer = CNX_APP_VERSION & CLng(App.Major) * 1000000 + CLng(App.Minor) * 1000 + CLng(App.Revision) & CNX_MSG_END
    strBuffer = strBuffer & CNX_WORKSTATION_NAME & guCliOpts.strWksnName & CNX_MSG_END
    strBuffer = strBuffer & CNX_COMPUTER_NAME & guCliOpts.strComputerName & CNX_PAYLOAD_TOKEN & guCliOpts.strUserName & CNX_PAYLOAD_TOKEN & intAllowLockdown & CNX_PAYLOAD_TOKEN & guCliOpts.strComputerMacAddress & CNX_MSG_END
    strBuffer = strBuffer & CNX_READY_FOR_SERVER_SCREEN & CStr(CInt(True)) & CNX_MSG_END
    If Dir(guCliOpts.strAppConfigPath & FILENAME_DB_BACKUP) <> "" Then
        strBuffer = strBuffer & CNX_DB_BACKUP_DTE & funDateTimeToSend(FileDateTime(guCliOpts.strAppConfigPath & FILENAME_DB_BACKUP)) & CNX_MSG_END
    End If
    Call funSendData(frmMain.wskCli, strBuffer)
    Call subInitSerializedUser
    guCliOpts.uFlags.blnServerInitDone = True
    Exit Sub

subInitServerError:
    Call funErrorManager("modFunctions.subInitServerError")
End Sub


'*******************************************************************************
Public Sub subForceUpdate()
Dim strBuffer As String
On Error GoTo subForceUpdateError

    
    If frmMain.wskCli.State = sckConnected Then
        strBuffer = CNX_FORCE_UPDATE & funDateTimeToSend(guUser.dteRemainingTime) & CNX_PAYLOAD_TOKEN & funDateTimeToSend(guUser.dteRemainingTimeAparent) & CNX_MSG_END
        Call funSendData(frmMain.wskCli, strBuffer)
    Else
        Call subSerializeUser
    End If
    Exit Sub

subForceUpdateError:
    Call funErrorManager("modFunctions.subForceUpdateError")
End Sub



'*******************************************************************************
Public Sub subDisableWksn(strReason As String)
Dim strBuffer As String
On Error GoTo subDisableWksnError

    Call subDeactivate
    Call frmMain.wskCli.Close
    frmMain.tmrSrvCnxStatus.Enabled = False
    frmMain.tmrAppKiller.Enabled = False
    frmMain.tmrWksnStatus.Enabled = False
    frmMain.tmrSessionStatus.Enabled = False
    frmMain.cmdActiver.Enabled = False
    frmMain.lblConnexionStatus.Caption = funLoadResString(10260)
    Call funMsg(WksnDisabled, strReason)
    Exit Sub

subDisableWksnError:
    Call funErrorManager("modFunctions.subDisableWksnError")
End Sub


'*******************************************************************************
Public Sub subSessionCleanup()
Dim hWnd1 As Long
Dim hWnd2 As Long
Dim hKey As Long
Dim lngRet As Long
Dim i As Integer
Dim bytData() As Byte
Dim strIEWidth As String
Dim strIEHeight As String
Dim strIELeft As String
Dim strIETop As String
On Error GoTo subSessionCleanupError

    If guCliOpts.intAllowLockdown = vbUnchecked Then Exit Sub
    If guCliOpts.uSessionCleanup.intAlignDesktopIcons = vbChecked Then
        hWnd1 = FindWindow("Progman", vbNullString)
        hWnd2 = GetWindow(hWnd1, GW_CHILD)
        hWnd1 = GetWindow(hWnd2, GW_CHILD)
        lngRet = SendMessage(hWnd1, LVM_ARRANGE, LVA_ALIGNLEFT, 0)
    End If
    If guCliOpts.uSessionCleanup.intDeleteDocHistory = vbChecked Then
        Call SHAddToRecentDocs(0, 0)
    End If
    If guCliOpts.uSessionCleanup.intRestoreBackground = vbChecked Then
        lngRet = RegCreateKeyEx(HKEY_CURRENT_USER, "Control panel\Desktop", 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, lngRet)
        lngRet = RegSetValueEx(hKey, "Wallpaper", 0&, REG_SZ, ByVal guCliOpts.uSessionCleanup.strBackgroundPath, Len(guCliOpts.uSessionCleanup.strBackgroundPath))
        lngRet = RegCloseKey(hKey)
        lngRet = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, guCliOpts.uSessionCleanup.strBackgroundPath, 0)
    End If
    If guCliOpts.uSessionCleanup.intRestoreIEHomePage = vbChecked Then
        Call SetRegUserSettings("Microsoft\Internet Explorer", "Main", "Start Page", guCliOpts.uSessionCleanup.strIEHomePagePath)
    End If
    If guCliOpts.uSessionCleanup.intRestoreIESettings = vbChecked Then
        Call SetRegUserSettings("Microsoft\Internet Explorer\International\Scripts", "3", "IEFontSize", 2, REG_BINARY)
    
        Call RegCreateKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, lngRet)
        strIEWidth = funLpad(Hex(Screen.Width / Screen.TwipsPerPixelX * 0.95), "0", 8)
        strIEHeight = funLpad(Hex(Screen.Height / Screen.TwipsPerPixelY * 0.95), "0", 8)
        strIELeft = funLpad(Hex(Screen.Width / Screen.TwipsPerPixelX * 0.05), "0", 8)
        strIETop = funLpad(Hex(Screen.Height / Screen.TwipsPerPixelY * 0.05), "0", 8)
        ReDim bytData(44)
        bytData(0) = 44
        bytData(8) = 1
        For i = 12 To 27
            bytData(i) = 255
        Next i
        bytData(28) = funHexToLong(Mid(strIELeft, 7, 2))
        bytData(29) = funHexToLong(Mid(strIELeft, 5, 2))
        bytData(30) = funHexToLong(Mid(strIELeft, 3, 2))
        bytData(31) = funHexToLong(Mid(strIELeft, 1, 2))
        bytData(32) = funHexToLong(Mid(strIETop, 7, 2))
        bytData(33) = funHexToLong(Mid(strIETop, 5, 2))
        bytData(34) = funHexToLong(Mid(strIETop, 3, 2))
        bytData(35) = funHexToLong(Mid(strIETop, 1, 2))
        bytData(36) = funHexToLong(Mid(strIEWidth, 7, 2))
        bytData(37) = funHexToLong(Mid(strIEWidth, 5, 2))
        bytData(38) = funHexToLong(Mid(strIEWidth, 3, 2))
        bytData(39) = funHexToLong(Mid(strIEWidth, 1, 2))
        bytData(40) = funHexToLong(Mid(strIEHeight, 7, 2))
        bytData(41) = funHexToLong(Mid(strIEHeight, 5, 2))
        bytData(42) = funHexToLong(Mid(strIEHeight, 3, 2))
        bytData(43) = funHexToLong(Mid(strIEHeight, 1, 2))
        Call RegSetValueEx(hKey, "Window_Placement", 0&, REG_BINARY, bytData(0), 44)
        Call RegCloseKey(hKey)
        
        Call subIEDeleteCache
        Call funIEDeleteHistory
    End If
    If guCliOpts.uSessionCleanup.intRestoreIMs = vbChecked Then
        If Dir(funGetFolderPath(CSIDL_APPDATA) & "\" & APP_DIR_GAIM & "\*") <> "" Then Call Kill(funGetFolderPath(CSIDL_APPDATA) & "\" & APP_DIR_GAIM & "\*")
        
        If Dir(funGetFolderPath(CSIDL_APPDATA) & "\" & APP_DIR_YAHOO) <> "" Then
            Call SetRegUserSettings("Yahoo", "Pager", "Yahoo! User ID", "")
        End If
    End If
        
    
    Exit Sub

subSessionCleanupError:
    Call funErrorManager("modFunctions.subSessionCleanupError")
End Sub


'*******************************************************************************
Public Sub subProcessReceivedFile(intFileAction As Integer, strFileContents As String)
Dim intfreefile As Integer
Dim strFileName As String
Dim blnModifyBlockFiles As Boolean
On Error GoTo subProcessReceivedFileError
            
    intfreefile = FreeFile
    strFileName = guCliOpts.strAppConfigPath & FILENAME_RECEIVED
    If strFileContents <> CNX_FILE_TEMP Then
        On Error Resume Next
        Call Kill(strFileName)
        On Error GoTo subProcessReceivedFileError
        Open strFileName For Binary Access Write As intfreefile
        Put intfreefile, , strFileContents
        Close intfreefile
    End If
    blnModifyBlockFiles = intFileAction <= enuFileActions.LoadLogo
    If blnModifyBlockFiles And gblnfrmBlock And frmBlock.Visible Then
        frmBlock.rtbMessage.FileName = ""
        frmBlock.imgLogo.Picture = LoadPicture
    End If
    
    Select Case intFileAction
    Case enuFileActions.LoadBlock
        Call FileCopy(strFileName, guCliOpts.strAppConfigPath & FILENAME_PREFIX & FILENAME_BLOCKED)
        guCliOpts.uPaths.strBlockedPath = guCliOpts.strAppConfigPath & FILENAME_PREFIX & FILENAME_BLOCKED
        Call SetRegSettings(funLoadResString(10006), "Message paths", "Blocked", guCliOpts.uPaths.strBlockedPath)
        guCliOpts.intMsgLogoFilesType = vbChecked
        Call SetRegSettings(funLoadResString(10006), "General", "Message & Logo Files Type", guCliOpts.intMsgLogoFilesType)
    
    Case enuFileActions.LoadIntro
        Call FileCopy(strFileName, guCliOpts.strAppConfigPath & FILENAME_PREFIX & FILENAME_INTRO)
        guCliOpts.uPaths.strIntroPath = guCliOpts.strAppConfigPath & FILENAME_PREFIX & FILENAME_INTRO
        Call SetRegSettings(funLoadResString(10006), "Message paths", "Intro", guCliOpts.uPaths.strIntroPath)
        guCliOpts.intMsgLogoFilesType = vbChecked
        Call SetRegSettings(funLoadResString(10006), "General", "Message & Logo Files Type", guCliOpts.intMsgLogoFilesType)
    
    Case enuFileActions.LoadLogo
        Call FileCopy(strFileName, guCliOpts.strAppConfigPath & FILENAME_PREFIX & FILENAME_LOGO)
        guCliOpts.uPaths.strLogoPath = guCliOpts.strAppConfigPath & FILENAME_PREFIX & FILENAME_LOGO
        Call SetRegSettings(funLoadResString(10006), "General", "Logo path", guCliOpts.uPaths.strLogoPath)
        guCliOpts.intMsgLogoFilesType = vbChecked
        Call SetRegSettings(funLoadResString(10006), "General", "Message & Logo Files Type", guCliOpts.intMsgLogoFilesType)
    
    Case enuFileActions.LoadNoCredit
        Call FileCopy(strFileName, guCliOpts.strAppConfigPath & FILENAME_PREFIX & FILENAME_NOCREDIT)
        guCliOpts.uPaths.strNoCreditPath = guCliOpts.strAppConfigPath & FILENAME_PREFIX & FILENAME_NOCREDIT
        Call SetRegSettings(funLoadResString(10006), "Message paths", "No credit", guCliOpts.uPaths.strNoCreditPath)
        guCliOpts.intMsgLogoFilesType = vbChecked
        Call SetRegSettings(funLoadResString(10006), "General", "Message & Logo Files Type", guCliOpts.intMsgLogoFilesType)
    
    Case enuFileActions.LoadSessionEnd
        Call FileCopy(strFileName, guCliOpts.strAppConfigPath & FILENAME_PREFIX & FILENAME_SESSIONEND)
        guCliOpts.uPaths.strSessionEndPath = guCliOpts.strAppConfigPath & FILENAME_PREFIX & FILENAME_SESSIONEND
        Call SetRegSettings(funLoadResString(10006), "Message paths", "Session end", guCliOpts.uPaths.strSessionEndPath)
        guCliOpts.intMsgLogoFilesType = vbChecked
        Call SetRegSettings(funLoadResString(10006), "General", "Message & Logo Files Type", guCliOpts.intMsgLogoFilesType)
    
    Case enuFileActions.LoadSessionPaused
        Call FileCopy(strFileName, guCliOpts.strAppConfigPath & FILENAME_PREFIX & FILENAME_SESSIONPAUSED)
        guCliOpts.uPaths.strSessionPausedPath = guCliOpts.strAppConfigPath & FILENAME_PREFIX & FILENAME_SESSIONPAUSED
        Call SetRegSettings(funLoadResString(10006), "Message paths", "Session paused", guCliOpts.uPaths.strSessionPausedPath)
        guCliOpts.intMsgLogoFilesType = vbChecked
        Call SetRegSettings(funLoadResString(10006), "General", "Message & Logo Files Type", guCliOpts.intMsgLogoFilesType)
    
    Case enuFileActions.LoadFollowImage
        Call FileCopy(strFileName, guCliOpts.strAppConfigPath & FILENAME_FOLLOW_IMAGE)
        frmBlock.imgFollowServer.Picture = LoadPicture(guCliOpts.strAppConfigPath & FILENAME_FOLLOW_IMAGE)
        Call funSendData(frmMain.wskCli, CNX_READY_FOR_SERVER_SCREEN & CStr(CInt(True)) & CNX_MSG_END)
    
    Case enuFileActions.LoadBackground
        Call FileCopy(strFileName, guCliOpts.strAppConfigPath & FILENAME_BACKGROUND)
        guCliOpts.uSessionCleanup.strBackgroundPath = guCliOpts.strAppConfigPath & FILENAME_BACKGROUND
    
    Case enuFileActions.SaveDB
        Call FileCopy(strFileName, guCliOpts.strAppConfigPath & FILENAME_DB_BACKUP)
    
    Case enuFileActions.SessionBatchStart
        Call FileCopy(strFileName, guCliOpts.strAppConfigPath & FILENAME_SESSION_BATCH_START)
    
    Case enuFileActions.SessionBatchEnd
        Call FileCopy(strFileName, guCliOpts.strAppConfigPath & FILENAME_SESSION_BATCH_END)
    
    End Select
    
    Call Kill(strFileName)
    
    If blnModifyBlockFiles And gblnfrmBlock And frmBlock.Visible Then
        Select Case guUser.intWksnState
        Case enuWksnState.Blocked_Deactivated
            frmBlock.rtbMessage.FileName = guCliOpts.uPaths.strBlockedPath
        
        Case enuWksnState.Blocked_SessionEnd
            frmBlock.rtbMessage.FileName = guCliOpts.uPaths.strSessionEndPath
        
        Case enuWksnState.Blocked_NoCredit
            frmBlock.rtbMessage.FileName = guCliOpts.uPaths.strNoCreditPath
        
        Case enuWksnState.Blocked_Paused, enuWksnState.Blocked_Locked
            frmBlock.rtbMessage.FileName = guCliOpts.uPaths.strSessionPausedPath
        
        Case enuWksnState.Ready
            frmBlock.rtbMessage.FileName = guCliOpts.uPaths.strIntroPath
        
        Case Else
        
        End Select
        frmBlock.imgLogo.Picture = LoadPicture(guCliOpts.uPaths.strLogoPath)
    End If
    
    Exit Sub

subProcessReceivedFileError:
    Call funErrorManager("modFunctions.subProcessReceivedFileError", Not gblnfrmBlock, True)
End Sub


'*******************************************************************************
Public Sub subSerializeUser()
Dim strBuffer As String * 200
Dim strPreBuffer As String
On Error GoTo subSerializeUserError

    If guUser.intWksnState > enuWksnState.Ready And guUser.intWksnState <> enuWksnState.Blocked_Deactivated Then
        strPreBuffer = guUser.lngSessionID & SERIALIZED_SESSION_TOKEN & guUser.intSessionType
        strPreBuffer = strPreBuffer & SERIALIZED_SESSION_TOKEN & funDateTimeToSend(guUser.dteElapsedTimeAparent)
        strPreBuffer = strPreBuffer & SERIALIZED_SESSION_TOKEN & funDateTimeToSend(guUser.dteTimeConnected)
        strPreBuffer = strPreBuffer & SERIALIZED_SESSION_TOKEN & funDateTimeToSend(guUser.dteTimeConnectedOffPeak)
        strPreBuffer = strPreBuffer & SERIALIZED_SESSION_TOKEN & funDateTimeToSend(guUser.dteTimeDisconnected)
        strPreBuffer = strPreBuffer & SERIALIZED_SESSION_TOKEN & funDateTimeToSend(Now + guCliOpts.uFlags.dteServerDelta)
        Select Case guUser.intSessionType
        Case enuSessionType.Card, enuSessionType.Client
            strPreBuffer = strPreBuffer & SERIALIZED_SESSION_TOKEN & guUser.strID
            strPreBuffer = strPreBuffer & SERIALIZED_SESSION_TOKEN & funDateTimeToSend(guUser.dteInitRemainingTimeSerialized - guUser.dteRemainingTime)
        Case enuSessionType.Selfservice
            strPreBuffer = strPreBuffer & SERIALIZED_SESSION_TOKEN & funSingleToSend(guUser.sngBalance)
        End Select
        strBuffer = strPreBuffer
'        Put guCliOpts.intSerializerFile, 1, strBuffer
        If guCliOpts.blnWksnTypeSndAln Then
            Call SetRegSettings(funLoadResString(10006), "General", "Serialized User", strBuffer)
        Else
            Call SetRegUserSettings(funLoadResString(10006), "General", "Serialized User", strBuffer)
        End If
    Else
        If guCliOpts.uServer.intCnxState = enuSrvCxn.Connected Then
            strBuffer = strPreBuffer
'            Put guCliOpts.intSerializerFile, 1, strBuffer
            If guCliOpts.blnWksnTypeSndAln Then
                Call SetRegSettings(funLoadResString(10006), "General", "Serialized User", strBuffer)
            Else
                Call SetRegUserSettings(funLoadResString(10006), "General", "Serialized User", strBuffer)
            End If
        End If
    End If
    Exit Sub

subSerializeUserError:
    Call funErrorManager("modFunctions.subSerializeUserError")
End Sub


'*******************************************************************************
Public Sub subInitSerializedUser()
Dim strBuffer As String * 200
Dim strPostBuffer As String
Dim strTokenizedBuffer() As String
Dim lngSessionID As Long
On Error GoTo subInitSerializedUserError

    
    If guCliOpts.blnWksnTypeSndAln Then
        strBuffer = GetRegSettings(funLoadResString(10006), "General", "Serialized User", "")
    Else
        strBuffer = GetRegUserSettings(funLoadResString(10006), "General", "Serialized User", "")
    End If
    strPostBuffer = Trim(strBuffer)
    If strPostBuffer <> "" Then
        Call subDebugPrint("subInitSerializedUser: " & strPostBuffer)
        strTokenizedBuffer = funTokenize(strPostBuffer, SERIALIZED_SESSION_TOKEN)
        If IsNumeric(strTokenizedBuffer(0)) Then
            lngSessionID = CLng(strTokenizedBuffer(0))
            strPostBuffer = guUser.intWksnState & SERIALIZED_SESSION_TOKEN & IIf(lngSessionID = guUser.lngSessionID, enuSerializedSession.Ongoing, enuSerializedSession.Ended) & SERIALIZED_SESSION_TOKEN & strPostBuffer
            Call funSendData(frmMain.wskCli, CNX_SESSION_RESUME & strPostBuffer & CNX_MSG_END)
            
            If CLng(strTokenizedBuffer(1)) = enuSessionType.Selfservice And lngSessionID <> guUser.lngSessionID Then
            ' This is an unpayed self service session! we must block the client!
                Call subDefaultState
                guUser.lngSessionID = lngSessionID
                guUser.intSessionType = CLng(strTokenizedBuffer(1))
                guUser.intWksnState = enuWksnState.Blocked_SessionEnd
                guUser.dteElapsedTimeAparent = funDateTimefromNet(strTokenizedBuffer(2))
                guUser.dteTimeConnected = funDateTimefromNet(strTokenizedBuffer(3))
                guUser.dteTimeConnectedOffPeak = funDateTimefromNet(strTokenizedBuffer(4))
                guUser.dteTimeDisconnected = funDateTimefromNet(strTokenizedBuffer(5))
                guUser.sngBalance = funSinglefromNet(strTokenizedBuffer(7))
                Call subBlock
            End If
        Else
'TEMP
            Call subDebugPrint("subInitSerializedUser strPostBuffer: '" & strPostBuffer & "'")
            Call subWriteToLogFile("subInitSerializedUser strPostBuffer: '" & strPostBuffer & "'")
        End If
    End If
    strBuffer = ""
    If guCliOpts.blnWksnTypeSndAln Then
        Call SetRegSettings(funLoadResString(10006), "General", "Serialized User", strBuffer)
    Else
        Call SetRegUserSettings(funLoadResString(10006), "General", "Serialized User", strBuffer)
    End If
    Exit Sub

subInitSerializedUserError:
    Call funErrorManager("modFunctions.subInitSerializedUserError")
'TEMP
    Call subDebugPrint("subInitSerializedUser strPostBuffer: '" & strPostBuffer & "'")
    Call subWriteToLogFile("subInitSerializedUser strPostBuffer: '" & strPostBuffer & "'")
End Sub


'*******************************************************************************
Public Sub subEmptySerializedUser()
On Error GoTo subEmptySerializedUserError

    If guCliOpts.blnWksnTypeSndAln Then
        Call SetRegSettings(funLoadResString(10006), "General", "Serialized User", "")
    Else
        Call SetRegUserSettings(funLoadResString(10006), "General", "Serialized User", "")
    End If
    Exit Sub

subEmptySerializedUserError:
    Call funErrorManager("modFunctions.subEmptySerializedUserError")
End Sub


'*******************************************************************************
Public Function funLoadResString(lngResID As Long, Optional strInsert1 As String = "", Optional strInsert2 As String = "", Optional strInsert3 As String = "", Optional strInsert4 As String = "") As String
Dim uResfile As Object
Dim strResString As String
On Error GoTo funLoadResStringError

    On Error Resume Next
    Set uResfile = CreateObject(guCliOpts.strInterfaceLang & ".clsResources")
    If Not uResfile Is Nothing Then strResString = uResfile.GetResStr(lngResID)
    On Error GoTo funLoadResStringError
    If strResString = "" Then strResString = LoadResString(lngResID)
    strResString = Replace(strResString, RES_NEW_LINE, vbNewLine)
    strResString = Replace(strResString, RES_TAB, vbTab)
    strResString = Replace(strResString, RES_INSERT_1, strInsert1)
    strResString = Replace(strResString, RES_INSERT_2, strInsert2)
    strResString = Replace(strResString, RES_INSERT_3, strInsert3)
    strResString = Replace(strResString, RES_INSERT_4, strInsert4)
    funLoadResString = strResString
    Exit Function
    
funLoadResStringError:
    Call funErrorManager("modFunctions.funLoadResStringError")
End Function


'*******************************************************************************
Public Function funLoadResStringDefault(lngResID As Long, Optional strInsert1 As String = "", Optional strInsert2 As String = "", Optional strInsert3 As String = "", Optional strInsert4 As String = "") As String
Dim uResfile As Object
Dim strResString As String
On Error GoTo funLoadResStringDefaultError

    On Error Resume Next
    Set uResfile = CreateObject(guCliOpts.strInterfaceLangDefault & ".clsResources")
    If Not uResfile Is Nothing Then strResString = uResfile.GetResStr(lngResID)
    On Error GoTo funLoadResStringDefaultError
    If strResString = "" Then strResString = LoadResString(lngResID)
    strResString = Replace(strResString, RES_NEW_LINE, vbNewLine)
    strResString = Replace(strResString, RES_TAB, vbTab)
    strResString = Replace(strResString, RES_INSERT_1, strInsert1)
    strResString = Replace(strResString, RES_INSERT_2, strInsert2)
    strResString = Replace(strResString, RES_INSERT_3, strInsert3)
    strResString = Replace(strResString, RES_INSERT_4, strInsert4)
    funLoadResStringDefault = strResString
    Exit Function
    
funLoadResStringDefaultError:
    Call funErrorManager("modFunctions.funLoadResStringDefaultError")
End Function


'*******************************************************************************
Public Sub subInitAutoConnection()
On Error GoTo subInitAutoConnectionError

    'Send out Auto config handshake
    Call frmMain.wskFindServer.Close
    frmMain.wskFindServer.Protocol = sckUDPProtocol
    frmMain.wskFindServer.RemoteHost = BROADCAST_ADDRESS
    frmMain.wskFindServer.RemotePort = DEFAULT_UDP_PORT_SERVER
    Call frmMain.wskFindServer.SendData(CNX_AUTO_CONFIG_HANDSHAKE_WKSN)
    
    'Set the winsock to listen for a reply
    Call frmMain.wskFindServer.Close
    frmMain.wskFindServer.RemoteHost = ""
    frmMain.wskFindServer.RemotePort = 0
    Call frmMain.wskFindServer.Bind(DEFAULT_UDP_PORT_WORKSTATION)
    Exit Sub

subInitAutoConnectionError:
Dim strErrDest As String
Dim lngErrNum As Long

    lngErrNum = Err.Number
    strErrDest = Err.Description
'    strErrDest = funErrorManager("modFunctions.subInitAutoConnectionError", False, True)
    If lngErrNum <> 10065 Then
        Call funMsg(NetworkConfigError)
        frmMain.tmrSrvCnxStatus.Enabled = False
        frmMain.cmdActiver.Enabled = False
        frmMain.lblConnexionStatus.Caption = funLoadResString(10218)
        guCliOpts.blnStartActive = False
        gblnCnxOK = False
        frmStatus.cmdCliPassword.Enabled = False
        frmStatus.cmdChat.Enabled = False
        If gblnfrmPasswdChange Then Call Unload(frmPasswdChange)
    Else
        'Call Err.Raise(CYBERA_ERROR, "modFunctions.subInitAutoConnectionError", strErrDest)
    End If
End Sub


'*******************************************************************************
Public Sub subRunSessionStartBatch()
Dim dblReturn As Double
On Error GoTo subRunSessionStartBatchError

    dblReturn = Shell(Chr(34) & guCliOpts.strAppConfigPath & FILENAME_SESSION_BATCH_START & Chr(34) & " " & IIf(guUser.strID <> "", guUser.strID, guCliOpts.strComputerName) & " " & guCliOpts.uServer.strServerName & " " & guCliOpts.strComputerName, guCliOpts.uSessionBatch.intWindowStyle)
    Exit Sub

subRunSessionStartBatchError:
    Call funErrorManager("modFunctions.subRunSessionStartBatchError", False, True)
End Sub


'*******************************************************************************
Public Sub subRunSessionEndBatch()
Dim dblReturn As Double
On Error GoTo subRunSessionEndBatchError

    dblReturn = Shell(Chr(34) & guCliOpts.strAppConfigPath & FILENAME_SESSION_BATCH_END & Chr(34) & " " & IIf(guUser.strID <> "", guUser.strID, guCliOpts.strComputerName) & " " & guCliOpts.uServer.strServerName & " " & guCliOpts.strComputerName, guCliOpts.uSessionBatch.intWindowStyle)
    Exit Sub
    
subRunSessionEndBatchError:
    Call funErrorManager("modFunctions.subRunSessionEndBatchError", False, True)
End Sub


'*******************************************************************************
Public Sub subResetGUILanguage()
Dim frmTemp As Form
On Error GoTo subResetGUILanguageError

    guCliOpts.strInterfaceLang = guCliOpts.strInterfaceLangDefault
    guCliOpts.intInterfaceLangCharset = guCliOpts.intInterfaceLangCharsetDefault
    For Each frmTemp In Forms
        Call frmTemp.subLocalizeInterface
    Next frmTemp
    Exit Sub
    
subResetGUILanguageError:
    Call funErrorManager("modFunctions.subResetGUILanguageError", False, True)
End Sub


'*******************************************************************************
Public Function funDelTree(strPath As String, intAttributes As Integer, Optional blnDeleteRoot As Boolean = True)
Dim strPaths() As String
Dim strCurrentPath As String
On Error GoTo funDelTreeError

'TODO DelTree Code
        
    Exit Function
    
funDelTreeError:
    Call funErrorManager("modFunctions.funDelTreeError", False, True)
End Function


'*******************************************************************************
Public Function funIsOffPeak(dteCurrentTime As Date) As Boolean
On Error GoTo funIsOffPeakError

    If guCliOpts.uTimeRatios.dteOffPeakStart > guCliOpts.uTimeRatios.dteOffPeakStop Then
        funIsOffPeak = dteCurrentTime < guCliOpts.uTimeRatios.dteOffPeakStop Or dteCurrentTime > guCliOpts.uTimeRatios.dteOffPeakStart
    Else
        funIsOffPeak = dteCurrentTime < guCliOpts.uTimeRatios.dteOffPeakStop And dteCurrentTime > guCliOpts.uTimeRatios.dteOffPeakStart
    End If
    Exit Function
    
funIsOffPeakError:
    funIsOffPeak = False
    Call funErrorManager("modFunctions.funIsOffPeakError", False, True)
End Function


'*******************************************************************************
Public Sub subSetDefaultUsage()
On Error GoTo subSetDefaultUsageError

    With frmBlock
        If guCliOpts.intDefaultUsage = enuSessionType.Client Then
            .OptClient.Value = True
        ElseIf guCliOpts.intDefaultUsage = enuSessionType.Selfservice Then
            If .optSelfService.Enabled Then
                .optSelfService.Value = True
                .cmdCxn.Enabled = True
            End If
        Else
            If .optCard.Enabled Then
                .optCard.Value = True
            Else
                .OptClient.Value = False
                .optSelfService.Value = False
                .optCard.Value = False
            End If
        End If
    End With
    Exit Sub
    
subSetDefaultUsageError:
    Call funErrorManager("modFunctions.subSetDefaultUsageError", False, True)
End Sub

