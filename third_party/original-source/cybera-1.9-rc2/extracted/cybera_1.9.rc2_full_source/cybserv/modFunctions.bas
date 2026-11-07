Attribute VB_Name = "modFunctions"

' _____________________________________________________________________________
'|________________________________ GENERAL FUNCTION MODULE ____________________|
'|  Name:          modFunctions.bas                                            |
'|  Type:          VB6 Module                                                  |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
'|  Description:   Contains general functions and subs                         |
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

' $Id: modFunctions.bas 41 2006-02-21 18:14:29Z sam $

Option Explicit


'*******************************************************************************
Public Sub Main()
Dim blnMultiInstance As Boolean
Dim strCommandLine As String
On Error GoTo MainError

    strCommandLine = Replace(LCase(Command), "/", "-")
    If InStr(strCommandLine, "-d") Then gblnDebug = True
    If InStr(strCommandLine, "-e:a") Then
        gintEventLevel = enuEventLevels.All
    ElseIf InStr(strCommandLine, "-e:v") Then
        gintEventLevel = enuEventLevels.Verbose
    End If
    If InStr(strCommandLine, "-m") <> 0 Then blnMultiInstance = True
    
    Call subDebugPrint("App startup")
    
    guServOpts.uFlags.dteOnTime = Now
    Call funInitApp
    Call Load(frmMain)
    
    If frmMain.sysInfo.OSPlatform <> 2 Then
        Call funMsg(OSNotNT)
        Call funShutdownApplication
    End If
    If Not blnMultiInstance Then Call subCheckForRunningServer
    
    If gblnDBOK Then Call subDBAddEvent(AppStart, 0, 0, "", 0, App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision & " - " & funLoadResString(10007))
    
    Call subInitWorkstations
    Call subInitCards
    Call subInitClients
    Call subFillCards
    Call subFillClients
    Call subInitIncome
    Call subInitServices
    Call subFillServices
    Call subInitSessions
    Call subInitFrmMain
    
    frmMain.tmrActivateServer.enabled = gblnDBOK
    Call frmMain.Show
    
    Call funUserLogoff(guServOpts.uConnectedUser, False)
    If gblnDBOK And (guServOpts.uConfigWizard.intSystem < CONFIG_WIZARD_SERV_SYSTEM Or guServOpts.uConfigWizard.intUser < CONFIG_WIZARD_SERV_USER) Then Call frmSetupWizard.Show(vbModal, frmMain)
    
    Exit Sub
    
MainError:
    Call funErrorManager("modFunctions.MainError", True)
    End
End Sub

'*******************************************************************************
Public Function funInitApp() As Boolean
Dim uMsg As clsStatusBarMessage
On Error GoTo funInitAppError
    
    App.Title = funLoadResString(10000)
    Call subDefineConstants
    
    On Error Resume Next
    guServOpts.uConfig.strAppConfigPath = funGetFolderPath(CSIDL_COMMON_APPDATA) & "\" & funLoadResString(10006)
    Call MkDir(guServOpts.uConfig.strAppConfigPath)
    guServOpts.uConfig.strAppConfigINIPath = guServOpts.uConfig.strAppConfigPath & "\" & App.EXEName & ".ini"
    On Error GoTo funInitAppError
    
    Call subGetSystemSettings
    Call subGetUserSettings
    
    App.Title = funLoadResString(10000)
    
    gblnDBOK = funInitDB
    If gblnDBOK Then
        Call subCardsMaintenance
        If guServOpts.uClients.intEnableExpiry = vbChecked Then Call subClientsMaintenance
    
        If grstUsers.EOF And grstUsers.BOF Then
            Set uMsg = New clsStatusBarMessage
            uMsg.ImageIndex = enuStatus.cliCritical
            uMsg.Text = funLoadResString(10773)
            Call funAddStatusBarMsg(uMsg, "NO_USERS")
        End If
        
        Set guCards = New Collection
        Set guClients = New Collection
    
    Else
        Set uMsg = New clsStatusBarMessage
        uMsg.ImageIndex = enuStatus.cliCritical
        uMsg.Text = funLoadResString(10369) & App.Title & funLoadResString(10370) & "(" & funLoadResString(10371) & ")"
        Call funAddStatusBarMsg(uMsg, "NO_DB")
        funInitApp = False
    End If
        
    Set guServOpts.uConnectedUser = New clsUser
    
    If guServOpts.uConfig.strPasswd = "" Then
        Set uMsg = New clsStatusBarMessage
        uMsg.ImageIndex = enuStatus.cliCritical
        uMsg.Text = funLoadResString(10719)
        Call funAddStatusBarMsg(uMsg, "NO_PWD")
    End If
    
    guServOpts.intNumVisRows = 19
    
    Exit Function
    
funInitAppError:
    Call funErrorManager("modFunctions.funInitAppError")
End Function


'*******************************************************************************
Public Sub subInitServer()
Dim uMsg As New clsStatusBarMessage
On Error GoTo subInitServerError

    Call frmMain.wskServ(0).Close
    Call frmMain.wskFindWorkstation.Close
    frmMain.wskServ(0).Protocol = sckTCPProtocol
    guServOpts.uServer.strServerIP = funGetIPfromMac(guServOpts.uServer.strServerMac)
    If guServOpts.uServer.blnAutoConnect Then
'        frmMain.wskServ(0).LocalPort = DEFAULT_TCP_PORT
        Call frmMain.wskServ(0).Bind(DEFAULT_TCP_PORT, guServOpts.uServer.strServerIP)
        Call frmMain.wskFindWorkstation.Bind(DEFAULT_UDP_PORT_SERVER, guServOpts.uServer.strServerIP)
    Else
        Call frmMain.wskServ(0).Bind(guServOpts.uServer.lngPort, guServOpts.uServer.strServerIP)
'        frmMain.wskServ(0).LocalPort = guServOpts.uServer.lngPort
    End If
    Call frmMain.wskServ(0).Listen
    gblnNETOK = True
    Exit Sub
    
subInitServerError:
    Select Case Err.Number
    Case sckAddressInUse
        Call frmMain.wskServ(0).Close
        If funMsg(enuMsg.ServerAlreadyRunning) = vbYes Then Call funShutdownApplication
    Case Else
        gblnNETOK = False
        Call funErrorManager("modFunctions.subInitServerError", False, True)
        Call funMsg(NetError)
    End Select
    uMsg.ImageIndex = enuStatus.cliCritical
    uMsg.Text = funLoadResString(10369) & App.Title & funLoadResString(10370) & "(" & funLoadResString(10496) & ")"
    Call funAddStatusBarMsg(uMsg, "NO_NET")
End Sub


'*******************************************************************************
Public Sub subInitCnxTesting()
On Error GoTo subInitCnxTestingError
    
    If gblnNETOK Then
        Call frmMain.wskCnxTest.Close
        frmMain.wskCnxTest.Protocol = sckTCPProtocol
        frmMain.wskCnxTest.RemoteHost = guServOpts.uCnxTest.strTestServerName
        frmMain.wskCnxTest.RemotePort = guServOpts.uCnxTest.strTestServerPort
        If gblnDBOK Then frmMain.tmrCnxTest.enabled = guServOpts.uCnxTest.intCnxTesting = vbChecked
    End If
    Exit Sub
    
subInitCnxTestingError:
    frmMain.tmrCnxTest.enabled = False
    Call funErrorManager("modFunctions.subInitCnxTestingError", False, True)
End Sub


'*******************************************************************************
Public Sub subInitFrmMain()
Dim i As Integer
On Error GoTo subInitFrmMainError
    
    With frmMain
        On Error Resume Next
        Call .tbsFunctions.Tabs.Remove("k2")
        Call .tbsFunctions.Tabs.Remove("k3")
        On Error GoTo subInitFrmMainError
        If guServOpts.uWorkstations.intActivateCards = vbChecked Then
            Call .tbsFunctions.Tabs.Add(2, "k2", funLoadResString(10150), 2)
            .tbsFunctions.Tabs("k2").ToolTipText = funLoadResString(10445)
        End If
        If guServOpts.uWorkstations.intActivateClients = vbChecked Then
            Call .tbsFunctions.Tabs.Add(IIf(guServOpts.uWorkstations.intActivateCards = vbChecked, 3, 2), "k3", funLoadResString(10796), 3)
            .tbsFunctions.Tabs("k3").ToolTipText = funLoadResString(10797)
        End If
        .tbsFunctions.Tabs(1).Selected = True
        
        .mnuFunctions(2).Visible = guServOpts.uWorkstations.intActivateCards = vbChecked
        .mnuFunctions(3).Visible = guServOpts.uWorkstations.intActivateClients = vbChecked
    
        .mnuWksnActivateConsole.Checked = guServOpts.uFlags.blnConsoleActive
        .mnuWksnShowConsole.Visible = guServOpts.uFlags.blnConsoleActive
        .fraWksnConsole.Visible = guServOpts.uFlags.blnConsoleActive
        If guServOpts.uFlags.blnConsoleActive Then
            .shpWksnConsoleBackground.BackColor = vb3DHighlight
            .fraWksnControls.Visible = False
            .fraWksnConsoleControls.Visible = True
            Call .fraWksnConsole_Click
            .mnuWksnUserRelease.Visible = False
        Else
            .ucWorkstations(1).Active = True
            .shpWksnConsoleBackground.BackColor = vb3DFace
            .fraWksnControls.Visible = True
            .fraWksnConsoleControls.Visible = False
        End If
        .tmrStateUpdade.enabled = True
        .tmrStatusBarUpdate.enabled = True
        .tmrScreenCapture.Interval = guServOpts.uWorkstations.intFollowInterval * 1000
        .tmrScreenCapture.enabled = True
        .mnuUserAdmin.enabled = gblnDBOK
        For i = 1 To .mnuWksnSize.UBound
            .mnuWksnSize(i).Checked = i = guServOpts.uFlags.intWksnSize
        Next i
        .tmrBackupDBToWorkstation.enabled = gblnDBOK And guServOpts.uDB.intDBBackupSendToWorkstation = vbChecked
'        .mnuSessionPrint.enabled = guServOpts.uConfig.strDefaultPrinterName <> ""
'        .mnuIncomePrint.enabled = guServOpts.uConfig.strDefaultPrinterName <> ""
'        .cmdSessionPrint.enabled = guServOpts.uConfig.strDefaultPrinterName <> ""
'        .cmdIncomePrint.enabled = guServOpts.uConfig.strDefaultPrinterName <> ""
    End With
    Call Shell_NotifyIcon(NIM_ADD, guIconData)
    Exit Sub
    
subInitFrmMainError:
    Call funErrorManager("modFunctions.subInitFrmMainError", False, True)
End Sub


'*******************************************************************************
Public Sub subInitCards()
Dim i As Integer
On Error GoTo subInitCardsError

    With frmMain.fxgCards
        .width = frmMain.shpCrdListBorder.width - 2 * Screen.TwipsPerPixelX
        .top = frmMain.shpCrdListBorder.top + Screen.TwipsPerPixelY
        .left = Screen.TwipsPerPixelX
        .Cols = 4
        .Row = 0
        .ColWidth(0) = 670
        .ColAlignment(0) = flexAlignCenterTop
        .ColWidth(1) = 800
        .ColAlignment(1) = flexAlignRightTop
        .ColWidth(2) = 1500
        .ColAlignment(2) = flexAlignLeftTop
        .ColWidth(3) = 1500
        .ColAlignment(3) = flexAlignLeftTop
        .enabled = gblnDBOK
        .Font = ""
    End With
    With frmMain.fxgCrdPending
        .width = frmMain.shpCrdPending.width - 2 * Screen.TwipsPerPixelX
        .top = frmMain.shpCrdPending.top + Screen.TwipsPerPixelY
        .left = frmMain.shpCrdPending.left + Screen.TwipsPerPixelX
        .height = frmMain.shpCrdPending.height - 2 * Screen.TwipsPerPixelX
        .Cols = 3
        .Row = 0
        .ColWidth(0) = 670
        .ColAlignment(0) = flexAlignCenterTop
        .ColWidth(1) = 800
        .ColAlignment(1) = flexAlignRightTop
        .ColWidth(2) = 1500
        .ColAlignment(2) = flexAlignLeftTop
        .enabled = gblnDBOK
        .Font = ""
    End With
    With frmMain
        
        .cmdCrdDelete.enabled = gblnDBOK
        .cmdCrdFind.enabled = gblnDBOK
        For i = .cmdCrdButtons.LBound To .cmdCrdButtons.UBound
            .cmdCrdButtons(i).enabled = gblnDBOK
        Next i
        For i = .mnuCrd.LBound To .mnuCrd.UBound
            If .mnuCrd(i).Caption <> "-" Then .mnuCrd(i).enabled = gblnDBOK
        Next i
        For i = .mnuCrdDD.LBound To .mnuCrdDD.UBound
            If .mnuCrdDD(i).Caption <> "-" Then .mnuCrdDD(i).enabled = gblnDBOK
        Next i
        
        .txtCrdFindNum.enabled = gblnDBOK
        .txtCrdComment.enabled = gblnDBOK
        .cmbCrdType.enabled = gblnDBOK
        
        .tmrCrdRefresh.Interval = 1000 * CLng(guServOpts.uCards.intCardListUpdateInterval)
    End With
    
    Call subFillCardsPending
    
    Exit Sub
    
subInitCardsError:
    Call funErrorManager("modFunctions.subInitCardsError")
End Sub


'*******************************************************************************
Public Sub subInitClients()
On Error GoTo subInitClientsError

    With frmMain.fxgClients
        .width = frmMain.shpCliListBorder.width - 2 * Screen.TwipsPerPixelX
        .top = frmMain.shpCliListBorder.top + Screen.TwipsPerPixelY
        .left = Screen.TwipsPerPixelX
        .Cols = 2
        .Row = 0
        .ColWidth(0) = 1600
        .ColAlignment(0) = flexAlignLeftTop
        .ColWidth(1) = 2000
        .ColAlignment(1) = flexAlignLeftTop
        .enabled = gblnDBOK
        .Font = ""
    End With
    frmMain.cmdCliFind.enabled = gblnDBOK
    frmMain.cmdCliButtons(0).enabled = gblnDBOK
    frmMain.cmdCliButtons(1).enabled = gblnDBOK
    frmMain.cmdCliButtons(2).enabled = gblnDBOK
    frmMain.cmdCliButtons(5).enabled = gblnDBOK
    frmMain.cmdCliButtons(6).enabled = gblnDBOK
    frmMain.cmdCliButtons(7).enabled = gblnDBOK
    frmMain.cmdCliButtons(3).enabled = gblnDBOK
    frmMain.cmdCliButtons(4).enabled = gblnDBOK
    
    frmMain.lbl(40).enabled = guServOpts.uClients.intEnableExpiry = vbChecked
    frmMain.cmdCliButtons(8).Visible = guServOpts.uClients.intEnableExpiry = vbUnchecked
    frmMain.cmdCliButtons(2).width = IIf(guServOpts.uClients.intEnableExpiry = vbChecked, 1395, 1215)
    frmMain.lblCliSel(3).enabled = guServOpts.uClients.intEnableExpiry = vbChecked
    frmMain.lblCliSel(3).Caption = IIf(guServOpts.uClients.intEnableExpiry = vbChecked, "", funLoadResString(10865))
    frmMain.mnuCli(10).Visible = guServOpts.uClients.intEnableExpiry = vbChecked
    frmMain.mnuCli(7).Visible = guServOpts.uClients.intEnableExpiry = vbUnchecked
    frmMain.mnuCliDD(7).Visible = guServOpts.uClients.intEnableExpiry = vbUnchecked
    
    frmMain.mnuCli(9).enabled = gblnDBOK
    frmMain.mnuCliDD(9).enabled = gblnDBOK
    frmMain.mnuCli(0).enabled = gblnDBOK
    frmMain.mnuCli(10).enabled = gblnDBOK
    frmMain.mnuCli(1).enabled = gblnDBOK
    frmMain.mnuCliDD(1).enabled = gblnDBOK
    frmMain.mnuCli(2).enabled = gblnDBOK
    frmMain.mnuCliDD(2).enabled = gblnDBOK
    frmMain.mnuCli(5).enabled = gblnDBOK
    frmMain.mnuCliDD(5).enabled = gblnDBOK
    frmMain.mnuCli(3).enabled = gblnDBOK
    frmMain.mnuCliDD(3).enabled = gblnDBOK
    frmMain.mnuCli(4).enabled = gblnDBOK
    frmMain.mnuCliDD(4).enabled = gblnDBOK
    frmMain.mnuCli(6).enabled = gblnDBOK
    frmMain.mnuCliDD(6).enabled = gblnDBOK
    
    frmMain.txtCliFind.enabled = gblnDBOK
    frmMain.tmrCliRefresh.Interval = 1000 * CLng(guServOpts.uClients.intClientListUpdateInterval)
    Exit Sub
    
subInitClientsError:
    Call funErrorManager("modFunctions.subInitClientsError")
End Sub


'*******************************************************************************
Public Sub subInitIncome()
Dim i As Integer, j As Integer
On Error GoTo subInitIncomeError

    With frmMain.fxgIncome
        .Cols = 8
        .Rows = 2
        
        .Row = 0
        .ColWidth(0) = 1900
        .ColAlignment(0) = flexAlignLeftTop
        
        .ColWidth(1) = 900
        .ColAlignment(1) = flexAlignRightTop
        
        .ColWidth(2) = 2500
        .ColAlignment(2) = flexAlignLeftTop
        
        .ColWidth(3) = 1200
        .ColAlignment(3) = flexAlignLeftTop
        
        .ColWidth(4) = 1500
        .ColAlignment(4) = flexAlignLeftTop
        
        .ColWidth(5) = 1000
        .ColAlignment(5) = flexAlignRightTop
        
        .ColWidth(6) = 500
        .ColAlignment(6) = flexAlignLeftTop
        
        .ColWidth(7) = 1500
        .ColAlignment(7) = flexAlignLeftTop
        
        For i = 1 To .Rows - 1
            For j = 0 To .Cols - 1
                .TextMatrix(i, j) = ""
            Next j
        Next i
        .enabled = gblnDBOK
        .Font = ""
    End With
    frmMain.dtpStartDate.value = Now
    frmMain.dtpEndDate.value = Now
    If gblnDBOK Then
        If Not (grstJournIncome.BOF = True And grstJournIncome.EOF = True) Then
            Call grstJournIncome.MoveFirst
            frmMain.dtpStartDate.value = grstJournIncome!date_stamp
            Call grstJournIncome.MoveLast
            frmMain.dtpEndDate.value = grstJournIncome!date_stamp
        End If
    End If
    frmMain.cmdIncomeExport.enabled = False
    frmMain.cmdIncomePrint.enabled = False
    frmMain.mnuIncomeExport.enabled = False
    frmMain.mnuIncomePrint.enabled = False
    frmMain.cmdIncomeFilter.enabled = gblnDBOK And guServOpts.uConnectedUser.AccessJournals
    frmMain.mnuIncomeApply.enabled = gblnDBOK And guServOpts.uConnectedUser.AccessJournals
    Exit Sub
    
subInitIncomeError:
    funErrorManager ("modFunctions.subInitIncomeError")
End Sub


'*******************************************************************************
Public Sub subInitSessions(Optional blnInitDatePickers As Boolean = True)
Dim i As Integer, j As Integer, k As Integer
Dim intNumHours As Integer
Dim intNumIntervals As Integer
On Error GoTo subInitSessionsError

    With frmMain.fxgSession(1)
        .Cols = 9 + guServOpts.uWorkstations.lngNumWorkstations * 2
        .Rows = 2
        
        .Row = 0
        .ColWidth(0) = 1300
        .ColAlignment(0) = flexAlignLeftTop
        
        .ColWidth(1) = 800
        .ColAlignment(1) = flexAlignRightTop
        
        .ColWidth(2) = 900
        .ColAlignment(2) = flexAlignRightTop
        
        .ColWidth(3) = 1000
        .ColAlignment(3) = flexAlignRightTop
        
        .ColWidth(4) = 900
        .ColAlignment(4) = flexAlignRightTop
        
        .ColWidth(5) = 400
        .ColAlignment(5) = flexAlignLeftTop
        
        .ColWidth(6) = 1200
        .ColAlignment(6) = flexAlignRightTop
        
        For i = 7 To .Cols - 2 Step 2
            .ColWidth(i) = 800
            .ColAlignment(i) = flexAlignRightTop
            .ColWidth(i + 1) = 400
            .ColAlignment(i + 1) = flexAlignLeftTop
        Next i
        
        For i = 1 To .Rows - 1
            For j = 0 To .Cols - 1
                .TextMatrix(i, j) = ""
            Next j
        Next i
        .enabled = gblnDBOK
        .Font = ""
    End With
    
    With frmMain.fxgSession(3)
        .Cols = 11
        .Rows = 2
        
        .Row = 0
        .ColWidth(0) = 1900
        .ColAlignment(0) = flexAlignLeftTop
        
        .ColWidth(1) = 1500
        .ColAlignment(1) = flexAlignLeftTop
        
        .ColWidth(2) = 1200
        .ColAlignment(2) = flexAlignLeftTop
        
        .ColWidth(3) = 1200
        .ColAlignment(3) = flexAlignLeftTop
        
        .ColWidth(4) = 1200
        .ColAlignment(4) = flexAlignLeftTop
        
        .ColWidth(5) = 400
        .ColAlignment(5) = flexAlignLeftTop
        
        .ColWidth(6) = 1200
        .ColAlignment(6) = flexAlignLeftTop
        
        .ColWidth(7) = 400
        .ColAlignment(7) = flexAlignLeftTop
        
        .ColWidth(8) = 1200
        .ColAlignment(8) = flexAlignLeftTop
        
        .ColWidth(9) = 400
        .ColAlignment(9) = flexAlignLeftTop
        
        .ColWidth(10) = 1500
        .ColAlignment(10) = flexAlignLeftTop
        
        For i = 1 To .Rows - 1
            For j = 0 To .Cols - 1
                .TextMatrix(i, j) = ""
            Next j
        Next i
        .enabled = gblnDBOK
        .Font = ""
    End With
    
    With frmMain.fxgSession(4)
        .Cols = 5
        .Rows = 2
        
        .Row = 0
        .ColWidth(0) = 1900
        .ColAlignment(0) = flexAlignLeftTop
        
        .ColWidth(1) = 1200
        .ColAlignment(1) = flexAlignLeftTop
        
        .ColWidth(2) = 1500
        .ColAlignment(2) = flexAlignLeftTop
        
        .ColWidth(3) = 2500
        .ColAlignment(3) = flexAlignLeftTop
        
        .ColWidth(4) = 7000
        .ColAlignment(4) = flexAlignLeftTop
        
        For i = 1 To .Rows - 1
            For j = 0 To .Cols - 1
                .TextMatrix(i, j) = ""
            Next j
        Next i
        .enabled = gblnDBOK
        .Font = ""
    End With
    
    intNumIntervals = funArraySize(guServOpts.uStats.strStatsTimeInt)
    With frmMain.fxgSession(2)
        .Cols = 2 + 2 * (intNumIntervals + 1)
        .Rows = 2
        
        .Row = 0
        .ColWidth(0) = 1300
        .ColAlignment(0) = flexAlignLeftTop
        
        .ColWidth(1) = 900
        .ColAlignment(1) = flexAlignRightTop
        
        
        For i = 2 To .Cols - 2 Step 2
            .ColWidth(i) = 1000
            .ColWidth(i + 1) = 400
            .ColAlignment(i) = flexAlignCenterTop
            .ColAlignment(i + 1) = flexAlignLeftTop
        Next i
        
        For i = 1 To .Rows - 1
            For j = 0 To .Cols - 1
                .TextMatrix(i, j) = ""
            Next j
        Next i
        .enabled = gblnDBOK
        .Font = ""
    End With
    
    intNumHours = guServOpts.uStats.intStatsGraphStopTime - guServOpts.uStats.intStatsGraphStartTime
    With frmMain.fxgSession(5)
        .Cols = 12 * (intNumHours) + intNumHours + 1 + 1
        For i = 1 To .Cols - 1
            .ColWidth(i) = 60
        Next i
        .Rows = 2
        .TextMatrix(1, 0) = ""
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
        .CellBackColor = vbWindowBackground
        .ColWidth(0) = 1000
        .Row = 0
        For i = 0 To intNumHours - 1
            j = i * 12 + i + 1
            For k = j To j + 11
                .TextMatrix(0, k + 1) = " " & format(i + guServOpts.uStats.intStatsGraphStartTime, "00") & ":00"
            Next k
            .ColWidth(j) = Screen.TwipsPerPixelX
            .Col = j
            .ColSel = .Col
            .CellBackColor = &HC0C0C0
        Next i
        .Col = .Cols - 1
        .ColWidth(.Col) = 15
        .ColSel = .Col
        .CellBackColor = &HC0C0C0
        .Text = " "
        
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .MergeCol(0) = True
        .Font = ""
    End With
    
    If blnInitDatePickers Then
        frmMain.dtpSessionStartDate.value = Now
        frmMain.dtpSessionEndDate.value = Now
        If gblnDBOK Then
            If Not (grstJournSessions.BOF = True And grstJournSessions.EOF = True) Then
                Call grstJournSessions.MoveFirst
                frmMain.dtpSessionStartDate.value = grstJournSessions!date_start
                Call grstJournSessions.MoveLast
                frmMain.dtpSessionEndDate.value = grstJournSessions!date_start
            End If
            If Not (grstEvents.BOF = True And grstEvents.EOF = True) Then
                Call grstEvents.MoveFirst
                frmMain.dtpSessionStartDate.value = IIf(grstEvents!date_stamp < frmMain.dtpSessionStartDate.value, grstEvents!date_stamp, frmMain.dtpSessionStartDate.value)
                Call grstEvents.MoveLast
                frmMain.dtpSessionEndDate.value = IIf(grstEvents!date_stamp > frmMain.dtpSessionEndDate.value, grstEvents!date_stamp, frmMain.dtpSessionEndDate.value)
            End If
        End If
    End If
    frmMain.cmdSessionExport.enabled = False
    frmMain.mnuSessionExport.enabled = False
    frmMain.cmdSessionPrint.enabled = False
    frmMain.mnuSessionPrint.enabled = False
    frmMain.cmdSessionFilter.enabled = gblnDBOK And guServOpts.uConnectedUser.AccessJournals
    frmMain.mnuSessionApply.enabled = gblnDBOK And guServOpts.uConnectedUser.AccessJournals
    Exit Sub
    
subInitSessionsError:
        Call funErrorManager("modFunctions.subInitSessionsError")
End Sub


'*******************************************************************************
Public Sub subInitServices()
Dim i As Integer, j As Integer
On Error GoTo subInitServicesError

    With frmMain.fxgServices
        .Cols = 3
        .Rows = 100
        
        .Row = 0
        .ColWidth(0) = 2000
        .ColAlignment(0) = flexAlignLeftTop
        
        .ColWidth(1) = 1000
        .ColAlignment(1) = flexAlignRightTop
        
        .ColWidth(2) = 6000
        .ColAlignment(2) = flexAlignLeftTop
        
        For i = 1 To .Rows - 1
            For j = 0 To .Cols - 1
                .TextMatrix(i, j) = ""
            Next j
            .Col = 0
            .Row = i
            .ColSel = .Cols - 1
            .CellForeColor = vbWindowText
            .CellBackColor = vbWindowBackground
        Next i
        .enabled = gblnDBOK
        .Font = ""
    End With
    frmMain.cmdServicesCollect.enabled = False
    frmMain.mnuServicesCollect.enabled = False
    Exit Sub
    
subInitServicesError:
        Call funErrorManager("modFunctions.subInitServicesError")
End Sub


'*******************************************************************************
Public Sub subFillCards()
Dim i As Long, j As Long
Dim uNewCard As clsCard
Dim dteTotalRemainingTime As Date
Dim lngNumActiveCards As Long
Dim uCards  As New Collection
Dim blnForceRefresh As Boolean
On Error GoTo subFillCardsError

    
    If Not gblnDBOK Then Exit Sub
    Screen.MousePointer = vbArrowHourglass
    frmMain.fxgCards.MousePointer = vbArrowHourglass
    If grstCards.BOF = True And grstCards.EOF = True Then
        frmMain.fxgCards.Rows = guServOpts.intNumVisRows
        frmMain.fxgCards.enabled = False
        frmMain.lblCrdListTotal.Caption = " " & funFormatTime(NULL_DATE)
        frmMain.lblCrdListTotal.ToolTipText = "0/0" & funLoadResString(10507)
        frmMain.lbl(18).ToolTipText = "0/0" & funLoadResString(10507)
    Else
        Call grstCards.MoveFirst
        For i = 1 To grstCards.RecordCount
            If grstCards.EOF Then Exit For
            Set uNewCard = New clsCard
            uNewCard.Index = i
            uNewCard.ID = grstCards!ID
            uNewCard.TimeCredit = grstCards!time_credit
            uNewCard.UnlimitedPlay = grstCards!unlimited_play
            uNewCard.DateCreation = grstCards!date_creation
            uNewCard.DateExpiry = grstCards!date_expiry
            uNewCard.Active = grstCards!Active
            uNewCard.CardType = IIf(IsNull(grstCards!Type), "", grstCards!Type)
            uNewCard.Comment = IIf(IsNull(grstCards!Comment), "", grstCards!Comment)
            dteTotalRemainingTime = dteTotalRemainingTime + uNewCard.TimeCredit
            If uNewCard.TimeCredit <> NULL_DATE Then lngNumActiveCards = lngNumActiveCards + 1
            If blnForceRefresh Then
                uNewCard.ForceRefresh = True
            Else
                If guCards.Count < i Then
                    blnForceRefresh = True
                    uNewCard.ForceRefresh = True
                ElseIf uNewCard.ID <> guCards(i).ID Or guCards(i).ForceRefresh Then
                    blnForceRefresh = True
                    uNewCard.ForceRefresh = True
                Else
                    uNewCard.ForceRefresh = uNewCard.TimeCredit <> guCards(i).TimeCredit Or uNewCard.DateExpiry <> guCards(i).DateExpiry Or uNewCard.Active <> guCards(i).Active
                End If
            End If
            Call uCards.Add(uNewCard, CStr(uNewCard.ID))
            Call grstCards.MoveNext
        Next i
        Set guCards = uCards
        With frmMain.fxgCards
            .Rows = guCards.Count + 1
            .enabled = True
            .Col = 0
            For i = 1 To guCards.Count
                If guCards(i).ForceRefresh Then
                    .Row = i
                    .ColSel = .Cols - 1
                    .CellForeColor = IIf(i = guServOpts.lngSelectedCardIndex, vbHighlightText, IIf(guCards(i).Active, vbWindowText, vbGrayText))
                    .TextMatrix(i, 0) = guCards(i).ID
                    .TextMatrix(i, 1) = IIf(guCards(i).UnlimitedPlay, funLoadResString(11000), funFormatTime(guCards(i).TimeCredit))
                    .TextMatrix(i, 2) = guCards(i).DateExpiry
                    If guCards(i).Active Then
                        .TextMatrix(i, 3) = funLoadResString(10205)
                    Else
                        .TextMatrix(i, 3) = funLoadResString(10206)
                    End If
                    guCards(i).ForceRefresh = False
                End If
                DoEvents
            Next i
        End With
        frmMain.lblCrdListTotal.Caption = " " & funFormatTime(dteTotalRemainingTime)
        frmMain.lblCrdListTotal.ToolTipText = lngNumActiveCards & "/" & grstCards.RecordCount & funLoadResString(10507)
        frmMain.lbl(18).ToolTipText = lngNumActiveCards & "/" & grstCards.RecordCount & funLoadResString(10507)
    End If
    Screen.MousePointer = vbNormal
    frmMain.fxgCards.MousePointer = vbNormal
    Exit Sub
    
subFillCardsError:
    Screen.MousePointer = vbNormal
    frmMain.fxgCards.MousePointer = vbNormal
        Call funErrorManager("modFunctions.subFillCardsError")
End Sub


'*******************************************************************************
Public Sub subFillCardsPending()
Dim i As Long, j As Long
On Error GoTo subFillCardsPendingError

    If Not gblnDBOK Then Exit Sub
    If grstCardsPending.BOF = True And grstCardsPending.EOF = True Then
        frmMain.fxgCrdPending.Rows = 5
        frmMain.fxgCrdPending.enabled = False
    Else
        Call grstCardsPending.MoveFirst
        With frmMain.fxgCrdPending
            .Rows = 2
            For i = 1 To .Rows - 1
                For j = 0 To .Cols - 1
                    .TextMatrix(i, j) = ""
                Next j
            Next i
            .Rows = grstCardsPending.RecordCount + 1
            .enabled = True
            .Col = 0
            For i = 1 To grstCardsPending.RecordCount
                If grstCardsPending.EOF Then Exit For
                .Row = i
                .TextMatrix(i, 0) = grstCardsPending!ID
                .TextMatrix(i, 1) = IIf(grstCardsPending!unlimited_play, funLoadResString(11000), funFormatTime(grstCardsPending!time_credit))
                .TextMatrix(i, 2) = funFormatCurrencyString(grstCardsPending!Price, guServOpts.uConfig.strMonFormat, guServOpts.uConfig.strMonSymbol)
                Call grstCardsPending.MoveNext
            Next i
            .ToolTipText = funLoadResString(10953, grstCardsPending.RecordCount)
            .Row = 0
        End With
    End If
    Exit Sub
    
subFillCardsPendingError:
    Call funErrorManager("modFunctions.subFillCardsPendingError")
End Sub


'*******************************************************************************
Public Sub subFillClients()
Dim i As Long, j As Long
Dim varClients As Variant
Dim uNewClient As clsClient
Dim uTmpClient As New clsClient
Dim dteTotalRemainingTime As Date
Dim lngNumActiveClients As Long
Dim uClients  As New Collection
Dim blnForceRefresh As Boolean
On Error GoTo subFillClientsError

    
    If Not gblnDBOK Then Exit Sub
    Screen.MousePointer = vbArrowHourglass
    frmMain.fxgClients.MousePointer = vbArrowHourglass
    If grstClients.BOF = True And grstClients.EOF = True Then
        frmMain.fxgClients.Rows = guServOpts.intNumVisRows
        frmMain.fxgClients.enabled = False
        frmMain.lblCliListTotal.Caption = " " & funFormatTime(NULL_DATE)
        frmMain.lblCliListTotal.ToolTipText = "0/0" & funLoadResString(10507)
        frmMain.lbl(34).ToolTipText = "0/0" & funLoadResString(10507)
    Else
        Call grstClients.MoveFirst
        For i = 1 To grstClients.RecordCount
            If grstClients.EOF Then Exit For
            Set uNewClient = New clsClient
            uNewClient.Index = i
            uNewClient.login = grstClients!login
            uNewClient.Password = grstClients!Passwd
            uNewClient.TimeCredit = grstClients!time_credit
            uNewClient.UnlimitedPlay = grstClients!unlimited_play
            uNewClient.DateCreation = grstClients!date_creation
            uNewClient.DateExpiry = grstClients!date_expiry
            uNewClient.Active = grstClients!Active
            uNewClient.ClientType = IIf(IsNull(grstClients!Type), 0, grstClients!Type)
            uNewClient.FullName = IIf(IsNull(grstClients!full_name), "", grstClients!full_name)
            uNewClient.Comment = IIf(IsNull(grstClients!Comment), "", grstClients!Comment)
            dteTotalRemainingTime = dteTotalRemainingTime + uNewClient.TimeCredit
            If uNewClient.TimeCredit <> NULL_DATE Then lngNumActiveClients = lngNumActiveClients + 1
            If blnForceRefresh Then
                uNewClient.ForceRefresh = True
            Else
                If guClients.Count < i Then
                    blnForceRefresh = True
                    uNewClient.ForceRefresh = True
                ElseIf uNewClient.login <> guClients(i).login Or guClients(i).ForceRefresh Then
                    blnForceRefresh = True
                    uNewClient.ForceRefresh = True
                Else
                    uNewClient.ForceRefresh = uNewClient.TimeCredit <> guClients(i).TimeCredit Or uNewClient.DateExpiry <> guClients(i).DateExpiry Or uNewClient.Active <> guClients(i).Active
                End If
            End If
            Call uClients.Add(uNewClient, CStr(uNewClient.login))
            Call grstClients.MoveNext
        Next i
        Set guClients = uClients
        With frmMain.fxgClients
            .Rows = guClients.Count + 1
            .enabled = True
            .Col = 0
            For i = 1 To guClients.Count
                If guClients(i).ForceRefresh Then
                    .Row = i
                    .ColSel = .Cols - 1
                    .CellForeColor = IIf(i = guServOpts.lngSelectedClientIndex, vbHighlightText, IIf(guClients(i).Active, vbWindowText, vbGrayText))
                    .TextMatrix(i, 0) = guClients(i).login
                    .TextMatrix(i, 1) = IIf(guClients(i).UnlimitedPlay, funLoadResString(11000), funFormatTime(guClients(i).TimeCredit))
                    guClients(i).ForceRefresh = False
                End If
                DoEvents
            Next i
        End With
        frmMain.lblCliListTotal.Caption = " " & funFormatTime(dteTotalRemainingTime)
        frmMain.lblCliListTotal.ToolTipText = lngNumActiveClients & "/" & grstClients.RecordCount & funLoadResString(10824)
        frmMain.lbl(34).ToolTipText = lngNumActiveClients & "/" & grstClients.RecordCount & funLoadResString(10824)
    End If
    Screen.MousePointer = vbNormal
    frmMain.fxgClients.MousePointer = vbNormal
    Exit Sub
    
subFillClientsError:
    Screen.MousePointer = vbNormal
    frmMain.fxgClients.MousePointer = vbNormal
        Call funErrorManager("modFunctions.subFillClientsError")
End Sub


'*******************************************************************************
Public Sub subFillServices()
Dim i, j As Long
Dim varCards As Variant
Dim uNewService As clsService
Dim uTmpService As New clsService
Dim uMsg As New clsStatusBarMessage
On Error GoTo subFillServicesError

    
    If Not gblnDBOK Then Exit Sub
    If (grstServices.BOF = True And grstServices.EOF = True) Or grstServices.RecordCount = 0 Then
        Call subInitServices
        frmMain.fxgServices.enabled = False
        Set guServices = New Collection
    Else
                
        grstServices.Index = "ServicesIndex"
        Call grstServices.MoveFirst
        Set guServices = New Collection
        For i = 1 To grstServices.RecordCount
            Set uNewService = New clsService
            uNewService.ID = grstServices!ID
            uNewService.Index = i
            uNewService.Label = grstServices!Label
            uNewService.Price = grstServices!Price
            uNewService.Stock = grstServices!Stock
            uNewService.StockMode = grstServices!stock_mode
            uNewService.StockMinAlert = grstServices!stock_min_alert
            
            Call guServices.Add(uNewService, CStr(uNewService.Index))
            Call subRemoveStatusBarMsg("SERV_" & uNewService.ID)
            If uNewService.StockMode = enuServiceStockMode.Managed And uNewService.Stock <= uNewService.StockMinAlert Then
                uMsg.ImageIndex = enuStatus.cliInfo
                uMsg.Text = funLoadResString(10373) & uNewService.Label & funLoadResString(10374)
                Call funAddStatusBarMsg(uMsg, "SERV_" & uNewService.ID)
            End If
            Call grstServices.MoveNext
        Next i
    
        guServOpts.lngSelectedServiceIndex = 0
        With frmMain.fxgServices
            .Rows = 1
            .Rows = guServices.Count + 1
            .Col = 0
            .Row = 0
            .ColSel = .Cols - 1
            .CellBackColor = vb3DHighlight
            .FixedRows = 1
            
            .enabled = True
            For i = 1 To guServices.Count
                .TextMatrix(i, 0) = guServices(i).Label
                .TextMatrix(i, 1) = funFormatCurrencyString(guServices(i).Price, guServOpts.uConfig.strMonFormat)
                .Col = 2
                .Row = i
                If guServices(i).StockMode = enuServiceStockMode.Managed Then
                    .TextMatrix(i, 2) = guServices(i).Stock
                    .CellForeColor = IIf(i = guServOpts.lngSelectedServiceIndex, vbHighlightText, IIf(guServices(i).Stock <= guServices(i).StockMinAlert, vbRed, vbWindowText))
                Else
                    .TextMatrix(i, 2) = funLoadResString(10332)
                    .CellForeColor = IIf(i = guServOpts.lngSelectedServiceIndex, vbHighlightText, vbWindowText)
                End If
            Next i
        End With
    End If
    frmMain.lblServSel(0).Caption = ""
    frmMain.lblServSel(2).Caption = ""
    frmMain.lblServSel(1).Caption = ""
    frmMain.lblServSel(3).Caption = ""
    frmMain.udServicesControlsQuantity.value = 0
    frmMain.udServicesControlsQuantity.enabled = False
    Exit Sub
    
subFillServicesError:
    Call funErrorManager("modFunctions.subFillServicesError")
End Sub


'*******************************************************************************
Public Function funShutdownApplication() As Boolean
Dim uWorkstation As New clsWorkstation
On Error GoTo funShutdownApplicationError
    
    funShutdownApplication = True
    Screen.MousePointer = vbHourglass
    frmMain.tmrStateUpdade.enabled = False
    Call frmMain.wskFindWorkstation.Close
    frmMain.tmrStatusBarUpdate.enabled = False
    frmMain.stbStatus.Panels(1).Picture = frmMain.imlSatusBar.ListImages(2).Picture
    frmMain.stbStatus.Panels(1).Text = funLoadResString(10729)
    DoEvents
    For Each uWorkstation In guWorkstations
        If gblnDBOK Then Call subSessionClose(uWorkstation)
        Call uWorkstation.Cnx.Close
        DoEvents
    Next uWorkstation
    
    Call frmMain.wskCnxTest.Close
    
    If frmMain.WindowState = vbMinimized Then
        guServOpts.uWinPos.intWinState = vbNormal
    Else
        guServOpts.uWinPos.intWinState = frmMain.WindowState
    End If
    If frmMain.WindowState = vbNormal Then
        guServOpts.uWinPos.intTop = frmMain.top
        guServOpts.uWinPos.intLeft = frmMain.left
        guServOpts.uWinPos.intHeight = frmMain.height
        guServOpts.uWinPos.intWidth = frmMain.width
    End If
    Call subSetUserSettings
    
    If gblnDBOK Then
        If guServOpts.uWorkstations.intLoggedURLTimeSpan <> 0 Then Call subDBPurgeURLs
        Call subDBAddEvent(AppStop, 0, 0, "", 0, App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision & " - " & funLoadResString(10008))
        Call grstUsers.Close
        If IsEmpty(grstUsersQry) Then Call grstUsersQry.Close
        Call grstCards.Close
        Call grstCardsPending.Close
        Call grstClients.Close
        Call grstServices.Close
        Call grstJournIncome.Close
        If IsEmpty(grstJournIncomeQry) Then Call grstJournIncomeQry.Close
        Call grstJournSessions.Close
        If IsEmpty(grstJournSessionsQry) Then Call grstJournSessionsQry.Close
        Call grstEvents.Close
        If IsEmpty(grstEventsQry) Then Call grstEventsQry.Close
        Call grstUrls.Close
        If IsEmpty(grstUrlsQry) Then Call grstUrlsQry.Close
        Call grstRecharges.Close
        If IsEmpty(grstRechargesQry) Then Call grstRechargesQry.Close
        Call grstConfig.Close
        
        Call gdbsCybAdmin.Close
        Call gwspCybAdmin.Close
        If guServOpts.uDB.intDBCompact = vbChecked Then Call subDBCompact
        If guServOpts.uDB.intDBBackup = vbChecked Then Call subBackupDB
    End If
    
    Call subDebugPrint("App shutdown")
    Screen.MousePointer = vbNormal
    
    Call Shell_NotifyIcon(NIM_DELETE, guIconData)

    If Not gblnDebug Then Call subUnhook
    End
    Exit Function
    
funShutdownApplicationError:
    Call funErrorManager("modFunctions.funShutdownApplicationError", False, True)
    End
End Function


'*******************************************************************************
Private Sub subGetSystemSettings()
'Retreave system specific settings (from ini file and api calls)
Dim lngSize As Long
Dim strName As String
Dim strAppName As String
Dim strMaxNumWorkstations As String
Dim lngDefaultLangID As Long
Dim strDefaultLang As String
Dim intDefaultLangCharset As Integer
Dim strDefaultFilePrefix As String
Dim i As Integer
Dim uMac As clsMacAddress
Dim strComputerMacAddress As String
Dim uAdapters As typAdapters
On Error GoTo subGetSystemSettingsError
    
    'App system settings
    strAppName = funLoadResString(10006)
    strName = Space$(16)
    lngSize = Len(strName)
    Call GetComputerName(strName, lngSize)
    strName = LCase(left(strName, lngSize))

    guServOpts.uConfig.strPasswd = funDoXOR(GetRegSettings(strAppName, "General", "Password", ""), strAppName)
    guServOpts.uConfig.strMonSymbol = GetRegSettings(strAppName, "General", "Monetary symbol", funGetCurrencySymbol)
    guServOpts.uConfig.strMonFormat = GetRegSettings(strAppName, "General", "Monetary format", "# ### ##0.#0")
    If InStr(guServOpts.uConfig.strMonFormat, CURRENCY_SYMBOL_INSERT) = 0 Then guServOpts.uConfig.strMonFormat = guServOpts.uConfig.strMonFormat & " " & CURRENCY_SYMBOL_INSERT
    
    guServOpts.uConfig.intNumDaysInMonth = GetRegSettings(strAppName, "General", "Number of days in month", 30)
    gstrLogfilePath = GetRegSettings(strAppName, "General", "Log file path", guServOpts.uConfig.strAppConfigPath & "\" & App.EXEName & ".log")
    intDefaultLangCharset = 0
    Select Case GetThreadLocale
    Case 2060, 3084, 1036, 5132, 6156, 4108
        strDefaultLang = App.EXEName & LANG_RES_FR
    
    Case 1031, 3079, 5127, 4103, 2055
        strDefaultLang = App.EXEName & LANG_RES_DE
    
    Case 3082, 11274, 16394, 13322, 9226, 5130, 7178, 12298, 17418, 4106, 18442, 2058, 19466, 6154, 15370, 10250, 20490, 14346, 8202
        strDefaultLang = App.EXEName & LANG_RES_ES
    
    Case 1040
        strDefaultLang = App.EXEName & LANG_RES_IT
    
    Case 1032
        strDefaultLang = App.EXEName & LANG_RES_EL
        intDefaultLangCharset = 160
    
    Case 1046
        strDefaultLang = App.EXEName & LANG_RES_BR
    
    Case 1038
        strDefaultLang = App.EXEName & LANG_RES_HU
    
    Case 1055
        strDefaultLang = App.EXEName & LANG_RES_TR
    
    Case Else
        strDefaultLang = App.EXEName & LANG_RES_EN
    
    End Select
    guServOpts.uConfig.strInterfaceLang = GetRegSettings(strAppName, "General", "Interface Language", strDefaultLang)
    guServOpts.uConfig.intInterfaceLangCharset = GetRegSettings(strAppName, "General", "Interface Language Charset", intDefaultLangCharset, REG_DWORD)
    
    Select Case guServOpts.uConfig.strInterfaceLang
    Case App.EXEName & LANG_RES_EN
        strDefaultFilePrefix = FILENAME_PREFIX_EN
    
    Case App.EXEName & LANG_RES_FR
        strDefaultFilePrefix = FILENAME_PREFIX_FR
    
    Case App.EXEName & LANG_RES_DE
        strDefaultFilePrefix = FILENAME_PREFIX_DE
    
    Case App.EXEName & LANG_RES_ES
        strDefaultFilePrefix = FILENAME_PREFIX_ES
    
    Case App.EXEName & LANG_RES_IT
        strDefaultFilePrefix = FILENAME_PREFIX_IT
    
    Case App.EXEName & LANG_RES_BR
        strDefaultFilePrefix = FILENAME_PREFIX_BR
    
    Case App.EXEName & LANG_RES_EL
        strDefaultFilePrefix = FILENAME_PREFIX_EL
    
    Case App.EXEName & LANG_RES_HU
        strDefaultFilePrefix = FILENAME_PREFIX_HU
    
    Case App.EXEName & LANG_RES_TR
        strDefaultFilePrefix = FILENAME_PREFIX_TR
    
    End Select
    
    guServOpts.uWorkstations.lngNumWorkstations = GetRegSettings(strAppName, "Workstations", "Number of workstations", 1)
    strMaxNumWorkstations = GetRegSettings(strAppName, "Workstations", "Max Number of workstations", DELAULT_MAX_NUM_WORKSTATIONS)
    guServOpts.uWorkstations.lngMaxNumWorkstations = IIf(IsNumeric(strMaxNumWorkstations) And strMaxNumWorkstations >= DELAULT_MAX_NUM_WORKSTATIONS, strMaxNumWorkstations, DELAULT_MAX_NUM_WORKSTATIONS)
    If guServOpts.uWorkstations.lngMaxNumWorkstations > MAX_NUM_WORKSTATIONS Then guServOpts.uWorkstations.lngMaxNumWorkstations = MAX_NUM_WORKSTATIONS
    guServOpts.uWorkstations.intStartForceWorkstationsConnected = GetRegSettings(strAppName, "Workstations", "Start Force Connected", vbUnchecked, REG_DWORD)
    guServOpts.uWorkstations.intActivateSelfServicePostPaid = GetRegSettings(strAppName, "Workstations", "Activate Self Service", vbChecked, REG_DWORD)
    guServOpts.uWorkstations.intActivateSelfServicePrePaid = GetRegSettings(strAppName, "Workstations", "Activate Self Service PrePaid", vbChecked, REG_DWORD)
    guServOpts.uWorkstations.intActivateCards = GetRegSettings(strAppName, "Workstations", "Activate Cards", vbChecked, REG_DWORD)
    guServOpts.uWorkstations.intActivateClients = GetRegSettings(strAppName, "Workstations", "Activate workstations", vbChecked, REG_DWORD)
    guServOpts.uWorkstations.intUsageDefault = GetRegSettings(strAppName, "Workstations", "Usage Default", enuSessionType.Card, REG_DWORD)
    guServOpts.uWorkstations.intFollowInterval = GetRegSettings(strAppName, "Workstations", "Follow Interval", 5)
    guServOpts.uWorkstations.intHideCardNum = GetRegSettings(strAppName, "Workstations", "Hide Card Number", vbUnchecked)
    guServOpts.uWorkstations.intLogURLs = GetRegSettings(strAppName, "Workstations", "Log URLs", vbChecked)
    guServOpts.uWorkstations.intLoggedURLTimeSpan = GetRegSettings(strAppName, "Workstations", "Logged URL time span", 1)
    guServOpts.uWorkstations.intEnablePause = GetRegSettings(strAppName, "Workstations", "Enable Pause", vbUnchecked)
    guServOpts.uWorkstations.intDefaultState = GetRegSettings(strAppName, "Workstations", "Default state", enuWksnState.Ready)
    guServOpts.uWorkstations.intDefaultMode = GetRegSettings(strAppName, "Workstations", "Default mode", enuWksnMode.Connected)
    guServOpts.uWorkstations.intModifyLockdown = GetRegSettings(strAppName, "Workstations", "Modify Lockdown", vbChecked, REG_DWORD)
    guServOpts.uWorkstations.strMacAddresses = funTokenize(GetRegSettings(strAppName, "Workstations", "Mac Addresses", ""), TIME_TOKEN)
    
    guServOpts.uWorkstations.uLocks.intNoClose = GetRegSettings(strAppName, "Workstations\Locks", "No Close", vbChecked, REG_DWORD)
    guServOpts.uWorkstations.uLocks.intNoAdminTools = GetRegSettings(strAppName, "Workstations\Locks", "No Admin tools", vbChecked, REG_DWORD)
    guServOpts.uWorkstations.uLocks.intNoCPL = GetRegSettings(strAppName, "Workstations\Locks", "No CPL", vbChecked, REG_DWORD)
    guServOpts.uWorkstations.uLocks.intLockDesktop = GetRegSettings(strAppName, "Workstations\Locks", "Lock Desktop", vbChecked, REG_DWORD)
    guServOpts.uWorkstations.uLocks.intLockIE = GetRegSettings(strAppName, "Workstations\Locks", "Lock IE", vbChecked, REG_DWORD)
    guServOpts.uWorkstations.uLocks.intNoCommonGroup = GetRegSettings(strAppName, "Workstations\Locks", "No Common Group", vbUnchecked, REG_DWORD)
    guServOpts.uWorkstations.uLocks.intNoFileSystemAccess = GetRegSettings(strAppName, "Workstations\Locks", "No File System Access", vbChecked, REG_DWORD)
    guServOpts.uWorkstations.uLocks.intNoFileSystemAccessType = GetRegSettings(strAppName, "Workstations\Locks", "No File System Access Type", vbUnchecked, REG_DWORD)
    guServOpts.uWorkstations.uLocks.lngLockedDrives = GetRegSettings(strAppName, "Workstations\Locks", "Locked Drives", WKSN_LOCK_DRIVES_C_ONLY, REG_DWORD)
    
    guServOpts.uDB.strDBPath = GetRegSettings(strAppName, "DB", "Path", guServOpts.uConfig.strAppConfigPath & "\cybera.mdb")
    guServOpts.uDB.strDBBackupPath = GetRegSettings(strAppName, "DB", "Backup Path", guServOpts.uConfig.strAppConfigPath)
    guServOpts.uDB.strDBPasswd = funDoXOR(GetRegSettings(strAppName, "DB", "Password", funDoXOR("cybera", strAppName)), strAppName)
    guServOpts.uDB.intDBBackup = GetRegSettings(strAppName, "DB", "Backup", vbChecked, REG_DWORD)
    guServOpts.uDB.intDBMaxBackups = GetRegSettings(strAppName, "DB", "Max Backups", DELAULT_MAX_NUM_DB_BACKUPS)
    guServOpts.uDB.intDBMaxBackups = IIf(guServOpts.uDB.intDBMaxBackups > MAX_NUM_DB_BACKUPS, MAX_NUM_DB_BACKUPS, guServOpts.uDB.intDBMaxBackups)
    guServOpts.uDB.intDBBackupSendToWorkstation = GetRegSettings(strAppName, "DB", "Backup to workstation", vbChecked, REG_DWORD)
    guServOpts.uDB.intDBCompact = GetRegSettings(strAppName, "DB", "Compact as shutdown", vbChecked, REG_DWORD)
    
    guServOpts.uServer.strServerName = strName
    
    uAdapters = funGetIPAdapters
    If uAdapters.Count > 0 Then
        For i = 0 To uAdapters.Count - 1
            If uAdapters.Item(i).AdapterTypeID = MIB_IF_TYPE_ETHERNET And uAdapters.Item(i).IpAddressCount > 0 Then
                If uAdapters.Item(i).IpAddress(0) <> EMPTY_ADDRESS Then
                    strComputerMacAddress = uAdapters.Item(i).MacAdress
                    Exit For
                End If
            End If
        Next i
    End If
    guServOpts.uServer.strServerMac = GetRegSettings(strAppName, "Server", "Mac address", strComputerMacAddress)
    guServOpts.uServer.strServerIP = funGetIPfromMac(guServOpts.uServer.strServerMac)
    
    guServOpts.uServer.lngPort = GetRegSettings(strAppName, "Server", "Port", DEFAULT_TCP_PORT)
    guServOpts.uServer.blnAutoConnect = GetRegSettings(strAppName, "Server", "Automatic workstation connection", vbChecked, REG_DWORD) = vbChecked
    
    guServOpts.uConfigWizard.intSystem = GetRegSettings(strAppName, "General", "Config Wizard", 0)
    guServOpts.uConfigWizard.intUser = GetRegUserSettings(strAppName, "General", "Config Wizard", CONFIG_WIZARD_SERV_USER)
    
    guServOpts.uCards.intCardPurgeDelta = GetRegSettings(strAppName, "Cards", "Card purge delta", 1)
    guServOpts.uCards.intAutoCardPurge = GetRegSettings(strAppName, "Cards", "Auto Card purge", vbChecked, REG_DWORD)
    guServOpts.uCards.intCardListUpdateInterval = GetRegSettings(strAppName, "Cards", "List Update Interval", 20)
    guServOpts.uCards.strCardTypes = funTokenize(GetRegSettings(strAppName, "Cards", "Card Types", ""), CARD_TYPE_TOKEN)
    guServOpts.uCards.intBonusActive = GetRegSettings(strAppName, "Cards", "Bonus active", vbUnchecked, REG_DWORD)
    guServOpts.uCards.sngBonusThreshold = funSinglefromNet(GetRegSettings(strAppName, "Cards", "Bonus threshold", 0))
    guServOpts.uCards.dteBonusTime = funDateTimefromNet(GetRegSettings(strAppName, "Cards", "Bonus time", funDateTimeToSend(#12:00:00 AM#)))
    
    guServOpts.uClients.intEnableExpiry = GetRegSettings(strAppName, "Clients", "Enable Expiry", vbChecked, REG_DWORD)
    guServOpts.uClients.intClientPurgeDelta = GetRegSettings(strAppName, "Clients", "Client purge delta", 1)
    guServOpts.uClients.intAutoClientPurge = GetRegSettings(strAppName, "Clients", "Auto Client purge", vbChecked, REG_DWORD)
    guServOpts.uClients.intClientListUpdateInterval = GetRegSettings(strAppName, "Clients", "List Update Interval", 20)
    guServOpts.uClients.intBonusActive = GetRegSettings(strAppName, "Clients", "Bonus active", vbUnchecked, REG_DWORD)
    guServOpts.uClients.sngBonusThreshold = funSinglefromNet(GetRegSettings(strAppName, "Clients", "Bonus threshold", 0))
    guServOpts.uClients.dteBonusTime = funDateTimefromNet(GetRegSettings(strAppName, "Clients", "Bonus time", funDateTimeToSend(#12:00:00 AM#)))
    
    guServOpts.uPostPaidSs.intMode = GetRegSettings(strAppName, "Postpaid Selfservice", "Mode", enuPostpaidSsMode.RechargeBased, REG_DWORD)
    guServOpts.uPostPaidSs.intSubMode = GetRegSettings(strAppName, "Postpaid Selfservice", "Sub Mode", IIf(guServOpts.uConfigWizard.intSystem < CONFIG_WIZARD_SERV_SYSTEM, enuPostpaidSsSubMode.RechSimple, enuPostpaidSsSubMode.RechAutoOptim), REG_DWORD)
    guServOpts.uPostPaidSs.dteEndTimeIncrement = funDateTimefromNet(GetRegSettings(strAppName, "Postpaid Selfservice", "End Time Increment", funDateTimeToSend(DEFAULT_PPSS_TIME_INCR)))
    guServOpts.uPostPaidSs.dteBaseTimeIncrement = funDateTimefromNet(GetRegSettings(strAppName, "Postpaid Selfservice", "Base Time Increment", funDateTimeToSend(DEFAULT_PPSS_TIME_INCR)))
    guServOpts.uPostPaidSs.dteMinBilledTime = funDateTimefromNet(GetRegSettings(strAppName, "Postpaid Selfservice", "Min Billed Time", funDateTimeToSend(DEFAULT_PPSS_TIME_INCR)))
    guServOpts.uPostPaidSs.dteGraceTime = funDateTimefromNet(GetRegSettings(strAppName, "Postpaid Selfservice", "Grace Time", funDateTimeToSend(NULL_DATE)))
    
    guServOpts.uTimeRatios.intDisconCon = GetRegSettings(strAppName, "Time ratios", "Disconnected/connected", 100)
    guServOpts.uTimeRatios.intOffPeak = GetRegSettings(strAppName, "Time ratios", "Off peak ratio", 100)
    guServOpts.uTimeRatios.dteOffPeakStart = funDateTimefromNet(GetRegSettings(strAppName, "Time ratios", "Off peak Start time", funDateTimeToSend(#8:00:00 PM#)))
    guServOpts.uTimeRatios.dteOffPeakStop = funDateTimefromNet(GetRegSettings(strAppName, "Time ratios", "Off peak Stop time", funDateTimeToSend(#7:00:00 AM#)))
    
    guServOpts.uWorkstations.uWksnMsg.intMsgCard = GetRegSettings(strAppName, "Workstations\Messages", "Activate Card messages", vbChecked, REG_DWORD)
    guServOpts.uWorkstations.uWksnMsg.intMsgCardAlert1 = GetRegSettings(strAppName, "Workstations\Messages", "1st Card message delay", 5)
    guServOpts.uWorkstations.uWksnMsg.intMsgCardAlert2 = GetRegSettings(strAppName, "Workstations\Messages", "2st Card message delay", 2)
    guServOpts.uWorkstations.uWksnMsg.intMsgSelfService = GetRegSettings(strAppName, "Workstations\Messages", "Activate Self service messages", vbUnchecked, REG_DWORD)
    guServOpts.uWorkstations.uWksnMsg.sngMsgSelfServiceBalance = funSinglefromNet(GetRegSettings(strAppName, "Workstations\Messages", "Self service message balance delta", 0))
    
    guServOpts.uCnxTest.intCnxTesting = GetRegSettings(strAppName, "Internet connection testing", "Activate testing", vbChecked, REG_DWORD)
    guServOpts.uCnxTest.strTestServerName = GetRegSettings(strAppName, "Internet connection testing", "Test server name", "www.microsoft.com")
    guServOpts.uCnxTest.strTestServerPort = GetRegSettings(strAppName, "Internet connection testing", "Test server port", "80")
    guServOpts.uCnxTest.intWorkstationNotification = GetRegSettings(strAppName, "Internet connection testing", "Notify workstations", vbChecked, REG_DWORD)
    
    guServOpts.uStats.strStatsTimeInt = funTokenize(GetRegSettings(strAppName, "Stats", "Time intervals", funDateTimeToSend(CDate("12:00:00"))), TIME_TOKEN)
    guServOpts.uStats.intStatsGraphStartTime = GetRegSettings(strAppName, "Stats", "Graph Start Time", 0)
    guServOpts.uStats.intStatsGraphStopTime = GetRegSettings(strAppName, "Stats", "Graph Stop Time", 24)
    
    guServOpts.uWorkstations.intModifyMsgLogoFiles = GetRegSettings(strAppName, "Workstations", "Modify Msg and Logo Files", vbUnchecked, REG_DWORD)
    guServOpts.uWorkstations.uPaths.strLogoPath = GetRegSettings(strAppName, "Workstations", "Logo path", guServOpts.uConfig.strAppConfigPath & FILENAME_PREFIX & FILENAME_LOGO)

On Error Resume Next
    If Dir(guServOpts.uConfig.strAppConfigPath & FILENAME_PREFIX & FILENAME_INTRO) = "" Then Call FileCopy(guServOpts.uConfig.strAppConfigPath & strDefaultFilePrefix & FILENAME_INTRO, guServOpts.uConfig.strAppConfigPath & FILENAME_PREFIX & FILENAME_INTRO)
    If Dir(guServOpts.uConfig.strAppConfigPath & FILENAME_PREFIX & FILENAME_BLOCKED) = "" Then Call FileCopy(guServOpts.uConfig.strAppConfigPath & strDefaultFilePrefix & FILENAME_BLOCKED, guServOpts.uConfig.strAppConfigPath & FILENAME_PREFIX & FILENAME_BLOCKED)
    If Dir(guServOpts.uConfig.strAppConfigPath & FILENAME_PREFIX & FILENAME_NOCREDIT) = "" Then Call FileCopy(guServOpts.uConfig.strAppConfigPath & strDefaultFilePrefix & FILENAME_NOCREDIT, guServOpts.uConfig.strAppConfigPath & FILENAME_PREFIX & FILENAME_NOCREDIT)
    If Dir(guServOpts.uConfig.strAppConfigPath & FILENAME_PREFIX & FILENAME_SESSIONEND) = "" Then Call FileCopy(guServOpts.uConfig.strAppConfigPath & strDefaultFilePrefix & FILENAME_SESSIONEND, guServOpts.uConfig.strAppConfigPath & FILENAME_PREFIX & FILENAME_SESSIONEND)
    If Dir(guServOpts.uConfig.strAppConfigPath & FILENAME_PREFIX & FILENAME_SESSIONPAUSED) = "" Then Call FileCopy(guServOpts.uConfig.strAppConfigPath & strDefaultFilePrefix & FILENAME_SESSIONPAUSED, guServOpts.uConfig.strAppConfigPath & FILENAME_PREFIX & FILENAME_SESSIONPAUSED)

    If Dir(guServOpts.uConfig.strAppConfigPath & FILENAME_CARD_TEMPLATE_FOLDER & FILENAME_CARD_TEMPLATE) = "" Then Call FileCopy(guServOpts.uConfig.strAppConfigPath & strDefaultFilePrefix & FILENAME_CARD_TEMPLATE, guServOpts.uConfig.strAppConfigPath & FILENAME_CARD_TEMPLATE_FOLDER & FILENAME_CARD_TEMPLATE)
On Error GoTo subGetSystemSettingsError

    guServOpts.uWorkstations.uPaths.strIntroPath = GetRegSettings(strAppName, "Workstations\Message paths", "Intro", guServOpts.uConfig.strAppConfigPath & FILENAME_PREFIX & FILENAME_INTRO)
    guServOpts.uWorkstations.uPaths.strBlockedPath = GetRegSettings(strAppName, "Workstations\Message paths", "Blocked", guServOpts.uConfig.strAppConfigPath & FILENAME_PREFIX & FILENAME_BLOCKED)
    guServOpts.uWorkstations.uPaths.strNoCreditPath = GetRegSettings(strAppName, "Workstations\Message paths", "No credit", guServOpts.uConfig.strAppConfigPath & FILENAME_PREFIX & FILENAME_NOCREDIT)
    guServOpts.uWorkstations.uPaths.strSessionEndPath = GetRegSettings(strAppName, "Workstations\Message paths", "Session end", guServOpts.uConfig.strAppConfigPath & FILENAME_PREFIX & FILENAME_SESSIONEND)
    guServOpts.uWorkstations.uPaths.strSessionPausedPath = GetRegSettings(strAppName, "Workstations\Message paths", "Session paused", guServOpts.uConfig.strAppConfigPath & FILENAME_PREFIX & FILENAME_SESSIONPAUSED)
    
    guServOpts.uSessionCleanup.intAlignDesktopIcons = GetRegSettings(strAppName, "Session Cleanup", "Align Desktop Icons", vbChecked, REG_DWORD)
    guServOpts.uSessionCleanup.intDeleteDocHistory = GetRegSettings(strAppName, "Session Cleanup", "Delete Document History", vbChecked, REG_DWORD)
    guServOpts.uSessionCleanup.intRestoreBackground = GetRegSettings(strAppName, "Session Cleanup", "Restore Background", vbUnchecked, REG_DWORD)
    guServOpts.uSessionCleanup.strBackgroundPath = GetRegSettings(strAppName, "Session Cleanup", "Restore Background Path", "")
    guServOpts.uSessionCleanup.intRestoreIESettings = GetRegSettings(strAppName, "Session Cleanup", "Restore IE Settings", vbChecked, REG_DWORD)
    guServOpts.uSessionCleanup.intRestoreIEHomePage = GetRegSettings(strAppName, "Session Cleanup", "Restore IE Home Page", vbChecked, REG_DWORD)
    guServOpts.uSessionCleanup.strIEHomePagePath = GetRegSettings(strAppName, "Session Cleanup", "Restore IE Home Page Path", "about:blank")
    guServOpts.uSessionCleanup.intRestoreIMs = GetRegSettings(strAppName, "Session Cleanup", "Restore IM Settings", vbChecked, REG_DWORD)
    
    guServOpts.uSessionBatch.intEnableStart = GetRegSettings(strAppName, "Session Batch", "Enable Start Batch", vbUnchecked, REG_DWORD)
    guServOpts.uSessionBatch.intEnableEnd = GetRegSettings(strAppName, "Session Batch", "Enable End Batch", vbUnchecked, REG_DWORD)
    guServOpts.uSessionBatch.strStartPath = GetRegSettings(strAppName, "Session Batch", "Start Batch Path", guServOpts.uConfig.strAppConfigPath & FILENAME_SESSION_BATCH_START)
    guServOpts.uSessionBatch.strEndPath = GetRegSettings(strAppName, "Session Batch", "End Batch Path", guServOpts.uConfig.strAppConfigPath & FILENAME_SESSION_BATCH_END)
    guServOpts.uSessionBatch.intWindowStyle = GetRegSettings(strAppName, "Session Batch", "Window Style", vbNormalFocus)
    
    guServOpts.uKilledApps.intForbidIE = GetRegSettings(strAppName, "Disconnected Killed Apps", "Forbid IE", vbChecked)
    guServOpts.uKilledApps.intForbidMozilla = GetRegSettings(strAppName, "Disconnected Killed Apps", "Forbid Mozilla", vbChecked)
    guServOpts.uKilledApps.intForbidNetscape = GetRegSettings(strAppName, "Disconnected Killed Apps", "Forbid Netscqpe", vbChecked)
    guServOpts.uKilledApps.intForbidOE = GetRegSettings(strAppName, "Disconnected Killed Apps", "Forbid OE", vbChecked)
    guServOpts.uKilledApps.strKilledApps = funTokenize(GetRegSettings(strAppName, "Disconnected Killed Apps", "Fobidden Apps", ""), FILENAME_TOKEN)
    
    Exit Sub
    
subGetSystemSettingsError:
    Call funErrorManager("modFunctions.subGetSystemSettingsError")
End Sub


'*******************************************************************************
Public Sub subSetSystemSettings()
Dim strAppName As String
On Error GoTo subSetSystemSettingsError
    
    strAppName = funLoadResString(10006)
    
    Call SetRegSettings(strAppName, "Workstations", "Number of workstations", guServOpts.uWorkstations.lngNumWorkstations)
    Call SetRegSettings(strAppName, "Workstations", "Max Number of workstations", guServOpts.uWorkstations.lngMaxNumWorkstations)
    Call SetRegSettings(strAppName, "Workstations", "Start Force Connected", guServOpts.uWorkstations.intStartForceWorkstationsConnected, REG_DWORD)
    Call SetRegSettings(strAppName, "Workstations", "Activate Self Service", guServOpts.uWorkstations.intActivateSelfServicePostPaid, REG_DWORD)
    Call SetRegSettings(strAppName, "Workstations", "Activate Self Service PrePaid", guServOpts.uWorkstations.intActivateSelfServicePrePaid, REG_DWORD)
    Call SetRegSettings(strAppName, "Workstations", "Activate Cards", guServOpts.uWorkstations.intActivateCards, REG_DWORD)
    Call SetRegSettings(strAppName, "Workstations", "Usage Default", guServOpts.uWorkstations.intUsageDefault, REG_DWORD)
    Call SetRegSettings(strAppName, "Workstations", "Activate workstations", guServOpts.uWorkstations.intActivateClients, REG_DWORD)
    Call SetRegSettings(strAppName, "Workstations", "Modify Lockdown", guServOpts.uWorkstations.intModifyLockdown, REG_DWORD)
    Call SetRegSettings(strAppName, "Workstations", "Follow Interval", guServOpts.uWorkstations.intFollowInterval)
    Call SetRegSettings(strAppName, "Workstations", "Hide Card Number", guServOpts.uWorkstations.intHideCardNum)
    Call SetRegSettings(strAppName, "Workstations", "Default state", guServOpts.uWorkstations.intDefaultState)
    Call SetRegSettings(strAppName, "Workstations", "Default mode", guServOpts.uWorkstations.intDefaultMode)
    Call SetRegSettings(strAppName, "Workstations", "Log URLs", guServOpts.uWorkstations.intLogURLs)
    Call SetRegSettings(strAppName, "Workstations", "Logged URL time span", guServOpts.uWorkstations.intLoggedURLTimeSpan)
    Call SetRegSettings(strAppName, "Workstations", "Enable Pause", guServOpts.uWorkstations.intEnablePause)
    Call SetRegSettings(strAppName, "Workstations", "Modify Msg and Logo Files", guServOpts.uWorkstations.intModifyMsgLogoFiles, REG_DWORD)
    Call SetRegSettings(strAppName, "Workstations", "Logo path", guServOpts.uWorkstations.uPaths.strLogoPath)
    
    Call SetRegSettings(strAppName, "Workstations\Message paths", "Intro", guServOpts.uWorkstations.uPaths.strIntroPath)
    Call SetRegSettings(strAppName, "Workstations\Message paths", "Blocked", guServOpts.uWorkstations.uPaths.strBlockedPath)
    Call SetRegSettings(strAppName, "Workstations\Message paths", "No credit", guServOpts.uWorkstations.uPaths.strNoCreditPath)
    Call SetRegSettings(strAppName, "Workstations\Message paths", "Session end", guServOpts.uWorkstations.uPaths.strSessionEndPath)
    Call SetRegSettings(strAppName, "Workstations\Message paths", "Session paused", guServOpts.uWorkstations.uPaths.strSessionPausedPath)

    Call SetRegSettings(strAppName, "Workstations\Locks", "No Close", guServOpts.uWorkstations.uLocks.intNoClose, REG_DWORD)
    Call SetRegSettings(strAppName, "Workstations\Locks", "No Admin tools", guServOpts.uWorkstations.uLocks.intNoAdminTools, REG_DWORD)
    Call SetRegSettings(strAppName, "Workstations\Locks", "No CPL", guServOpts.uWorkstations.uLocks.intNoCPL, REG_DWORD)
    Call SetRegSettings(strAppName, "Workstations\Locks", "Lock Desktop", guServOpts.uWorkstations.uLocks.intLockDesktop, REG_DWORD)
    Call SetRegSettings(strAppName, "Workstations\Locks", "Lock IE", guServOpts.uWorkstations.uLocks.intLockIE, REG_DWORD)
    Call SetRegSettings(strAppName, "Workstations\Locks", "No Common Group", guServOpts.uWorkstations.uLocks.intNoCommonGroup, REG_DWORD)
    Call SetRegSettings(strAppName, "Workstations\Locks", "No File System Access", guServOpts.uWorkstations.uLocks.intNoFileSystemAccess, REG_DWORD)
    Call SetRegSettings(strAppName, "Workstations\Locks", "No File System Access Type", guServOpts.uWorkstations.uLocks.intNoFileSystemAccessType, REG_DWORD)
    Call SetRegSettings(strAppName, "Workstations\Locks", "Locked Drives", guServOpts.uWorkstations.uLocks.lngLockedDrives, REG_DWORD)
    
    Call SetRegSettings(strAppName, "Workstations\Messages", "Activate Card messages", guServOpts.uWorkstations.uWksnMsg.intMsgCard, REG_DWORD)
    Call SetRegSettings(strAppName, "Workstations\Messages", "1st Card message delay", guServOpts.uWorkstations.uWksnMsg.intMsgCardAlert1)
    Call SetRegSettings(strAppName, "Workstations\Messages", "2st Card message delay", guServOpts.uWorkstations.uWksnMsg.intMsgCardAlert2)
    Call SetRegSettings(strAppName, "Workstations\Messages", "Activate Self service messages", guServOpts.uWorkstations.uWksnMsg.intMsgSelfService, REG_DWORD)
    Call SetRegSettings(strAppName, "Workstations\Messages", "Self service message balance delta", funSingleToSend(guServOpts.uWorkstations.uWksnMsg.sngMsgSelfServiceBalance))
    
    Call SetRegSettings(strAppName, "DB", "Path", guServOpts.uDB.strDBPath)
    Call SetRegSettings(strAppName, "DB", "Backup Path", guServOpts.uDB.strDBBackupPath)
    Call SetRegSettings(strAppName, "DB", "Password", funDoXOR(guServOpts.uDB.strDBPasswd, strAppName))
    Call SetRegSettings(strAppName, "DB", "Backup", guServOpts.uDB.intDBBackup, REG_DWORD)
    Call SetRegSettings(strAppName, "DB", "Backup to workstation", guServOpts.uDB.intDBBackupSendToWorkstation, REG_DWORD)
    Call SetRegSettings(strAppName, "DB", "Compact as shutdown", guServOpts.uDB.intDBCompact, REG_DWORD)
    Call SetRegSettings(strAppName, "DB", "Max Backups", guServOpts.uDB.intDBMaxBackups)
    
    Call SetRegSettings(strAppName, "Server", "Mac address", guServOpts.uServer.strServerMac)
    Call SetRegSettings(strAppName, "Server", "Port", guServOpts.uServer.lngPort)
    Call SetRegSettings(strAppName, "Server", "Automatic workstation connection", IIf(guServOpts.uServer.blnAutoConnect, vbChecked, vbUnchecked), REG_DWORD)
    
    Call SetRegSettings(strAppName, "Cards", "Card purge delta", guServOpts.uCards.intCardPurgeDelta)
    Call SetRegSettings(strAppName, "Cards", "Auto Card purge", guServOpts.uCards.intAutoCardPurge, REG_DWORD)
    Call SetRegSettings(strAppName, "Cards", "List Update Interval", guServOpts.uCards.intCardListUpdateInterval)
    Call SetRegSettings(strAppName, "Cards", "Card Types", funConcatToken(guServOpts.uCards.strCardTypes, CARD_TYPE_TOKEN))
    Call SetRegSettings(strAppName, "Cards", "Bonus active", guServOpts.uCards.intBonusActive, REG_DWORD)
    Call SetRegSettings(strAppName, "Cards", "Bonus threshold", funSingleToSend(guServOpts.uCards.sngBonusThreshold))
    Call SetRegSettings(strAppName, "Cards", "Bonus time", funDateTimeToSend(guServOpts.uCards.dteBonusTime))
    
    Call SetRegSettings(strAppName, "Clients", "Enable Expiry", guServOpts.uClients.intEnableExpiry, REG_DWORD)
    Call SetRegSettings(strAppName, "Clients", "Client purge delta", guServOpts.uClients.intClientPurgeDelta)
    Call SetRegSettings(strAppName, "Clients", "Auto Client purge", guServOpts.uClients.intAutoClientPurge, REG_DWORD)
    Call SetRegSettings(strAppName, "Clients", "List Update Interval", guServOpts.uClients.intClientListUpdateInterval)
    Call SetRegSettings(strAppName, "Clients", "Bonus active", guServOpts.uClients.intBonusActive, REG_DWORD)
    Call SetRegSettings(strAppName, "Clients", "Bonus threshold", funSingleToSend(guServOpts.uClients.sngBonusThreshold))
    Call SetRegSettings(strAppName, "Clients", "Bonus time", funDateTimeToSend(guServOpts.uClients.dteBonusTime))
    
    Call SetRegSettings(strAppName, "Postpaid Selfservice", "Mode", guServOpts.uPostPaidSs.intMode, REG_DWORD)
    Call SetRegSettings(strAppName, "Postpaid Selfservice", "Sub Mode", guServOpts.uPostPaidSs.intSubMode, REG_DWORD)
    Call SetRegSettings(strAppName, "Postpaid Selfservice", "End Time Increment", funDateTimeToSend(guServOpts.uPostPaidSs.dteEndTimeIncrement))
    Call SetRegSettings(strAppName, "Postpaid Selfservice", "Base Time Increment", funDateTimeToSend(guServOpts.uPostPaidSs.dteBaseTimeIncrement))
    Call SetRegSettings(strAppName, "Postpaid Selfservice", "Min Billed Time", funDateTimeToSend(guServOpts.uPostPaidSs.dteMinBilledTime))
    Call SetRegSettings(strAppName, "Postpaid Selfservice", "Grace Time", funDateTimeToSend(guServOpts.uPostPaidSs.dteGraceTime))
    
    Call SetRegSettings(strAppName, "General", "Monetary symbol", guServOpts.uConfig.strMonSymbol)
    Call SetRegSettings(strAppName, "General", "Monetary format", guServOpts.uConfig.strMonFormat)
    Call SetRegSettings(strAppName, "General", "Number of days in month", guServOpts.uConfig.intNumDaysInMonth)
    Call SetRegSettings(strAppName, "General", "Log file path", gstrLogfilePath)
    Call SetRegSettings(strAppName, "General", "Interface Language", guServOpts.uConfig.strInterfaceLang)
    Call SetRegSettings(strAppName, "General", "Interface Language Charset", guServOpts.uConfig.intInterfaceLangCharset, REG_DWORD)
    
    Call SetRegSettings(strAppName, "Time ratios", "Disconnected/connected", guServOpts.uTimeRatios.intDisconCon)
    Call SetRegSettings(strAppName, "Time ratios", "Off peak ratio", guServOpts.uTimeRatios.intOffPeak)
    Call SetRegSettings(strAppName, "Time ratios", "Off peak Start time", funDateTimeToSend(guServOpts.uTimeRatios.dteOffPeakStart))
    Call SetRegSettings(strAppName, "Time ratios", "Off peak Stop time", funDateTimeToSend(guServOpts.uTimeRatios.dteOffPeakStop))
    
    Call SetRegSettings(strAppName, "Internet connection testing", "Activate testing", guServOpts.uCnxTest.intCnxTesting, REG_DWORD)
    Call SetRegSettings(strAppName, "Internet connection testing", "Test server name", guServOpts.uCnxTest.strTestServerName)
    Call SetRegSettings(strAppName, "Internet connection testing", "Test server port", guServOpts.uCnxTest.strTestServerPort)
    Call SetRegSettings(strAppName, "Internet connection testing", "Notify workstations", guServOpts.uCnxTest.intWorkstationNotification, REG_DWORD)
    
    Call SetRegSettings(strAppName, "Stats", "Time intervals", funConcatToken(guServOpts.uStats.strStatsTimeInt, TIME_TOKEN))
    Call SetRegSettings(strAppName, "Stats", "Graph Start Time", guServOpts.uStats.intStatsGraphStartTime)
    Call SetRegSettings(strAppName, "Stats", "Graph Stop Time", guServOpts.uStats.intStatsGraphStopTime)
    
    Call SetRegSettings(strAppName, "Session Cleanup", "Align Desktop Icons", guServOpts.uSessionCleanup.intAlignDesktopIcons, REG_DWORD)
    Call SetRegSettings(strAppName, "Session Cleanup", "Delete Document History", guServOpts.uSessionCleanup.intDeleteDocHistory, REG_DWORD)
    Call SetRegSettings(strAppName, "Session Cleanup", "Restore Background", guServOpts.uSessionCleanup.intRestoreBackground, REG_DWORD)
    Call SetRegSettings(strAppName, "Session Cleanup", "Restore Background Path", guServOpts.uSessionCleanup.strBackgroundPath)
    Call SetRegSettings(strAppName, "Session Cleanup", "Restore IE Settings", guServOpts.uSessionCleanup.intRestoreIESettings, REG_DWORD)
    Call SetRegSettings(strAppName, "Session Cleanup", "Restore IM Settings", guServOpts.uSessionCleanup.intRestoreIMs, REG_DWORD)
    Call SetRegSettings(strAppName, "Session Cleanup", "Restore IE Home Page", guServOpts.uSessionCleanup.intRestoreIEHomePage, REG_DWORD)
    Call SetRegSettings(strAppName, "Session Cleanup", "Restore IE Home Page Path", guServOpts.uSessionCleanup.strIEHomePagePath)
        
    Call SetRegSettings(strAppName, "Session Batch", "Enable Start Batch", guServOpts.uSessionBatch.intEnableStart, REG_DWORD)
    Call SetRegSettings(strAppName, "Session Batch", "Start Batch Path", guServOpts.uSessionBatch.strStartPath)
    Call SetRegSettings(strAppName, "Session Batch", "Enable End Batch", guServOpts.uSessionBatch.intEnableEnd, REG_DWORD)
    Call SetRegSettings(strAppName, "Session Batch", "End Batch Path", guServOpts.uSessionBatch.strEndPath)
    Call SetRegSettings(strAppName, "Session Batch", "Window Style", guServOpts.uSessionBatch.intWindowStyle)
    
    Call SetRegSettings(strAppName, "Disconnected Killed Apps", "Forbid IE", guServOpts.uKilledApps.intForbidIE)
    Call SetRegSettings(strAppName, "Disconnected Killed Apps", "Forbid Mozilla", guServOpts.uKilledApps.intForbidMozilla)
    Call SetRegSettings(strAppName, "Disconnected Killed Apps", "Forbid Netscqpe", guServOpts.uKilledApps.intForbidNetscape)
    Call SetRegSettings(strAppName, "Disconnected Killed Apps", "Forbid OE", guServOpts.uKilledApps.intForbidOE)
    Call SetRegSettings(strAppName, "Disconnected Killed Apps", "Fobidden Apps", funConcatToken(guServOpts.uKilledApps.strKilledApps, FILENAME_TOKEN))
    
    Exit Sub
    
subSetSystemSettingsError:
    Call funErrorManager("modFunctions.subSetSystemSettingsError")
End Sub


'*******************************************************************************
Private Sub subGetUserSettings()
'Retreave user specific settings (from registry)
Dim strAppName As String
On Error GoTo subGetUserSettingsError
     
    guServOpts.uConfig.strDefaultPrinterName = funGetPrinerName
    
    strAppName = funLoadResString(10006)
     '*Get Main window position
    guServOpts.uWinPos.intHeight = GetRegUserSettings(strAppName, "Main window position", "Height", 455 * Screen.TwipsPerPixelY)
    guServOpts.uWinPos.intWidth = GetRegUserSettings(strAppName, "Main window position", "Width", 654 * Screen.TwipsPerPixelX)
    guServOpts.uWinPos.intTop = GetRegUserSettings(strAppName, "Main window position", "Top", (Screen.height - guServOpts.uWinPos.intHeight) / 2)
    guServOpts.uWinPos.intLeft = GetRegUserSettings(strAppName, "Main window position", "Left", (Screen.width - guServOpts.uWinPos.intWidth) / 2)
    guServOpts.uWinPos.intWinState = GetRegUserSettings(strAppName, "Main window position", "Window state", vbNormal)
    guServOpts.uConfig.strLastExportPath = GetRegUserSettings(strAppName, "General", "Last Export path", funGetFolderPath(CSIDL_DESKTOP))
    guServOpts.uFlags.blnConsoleActive = GetRegUserSettings(strAppName, "General", "Console Active", vbChecked, REG_DWORD) = vbChecked
    guServOpts.uFlags.intWksnSize = GetRegUserSettings(strAppName, "General", "Workstation Size", enuWksnSize.WksnSizeNormal)
    If right(guServOpts.uConfig.strLastExportPath, 1) <> "\" Then guServOpts.uConfig.strLastExportPath = guServOpts.uConfig.strLastExportPath & "\"
    Exit Sub
    
subGetUserSettingsError:
    Call funErrorManager("modFunctions.subGetUserSettingsError")
End Sub


'*******************************************************************************
Private Sub subSetUserSettings()
Dim strAppName As String
On Error GoTo subSetUserSettingsError

    strAppName = funLoadResString(10006)
    Call SetRegUserSettings(strAppName, "Main window position", "Top", guServOpts.uWinPos.intTop)
    Call SetRegUserSettings(strAppName, "Main window position", "Left", guServOpts.uWinPos.intLeft)
    Call SetRegUserSettings(strAppName, "Main window position", "Height", guServOpts.uWinPos.intHeight)
    Call SetRegUserSettings(strAppName, "Main window position", "Width", guServOpts.uWinPos.intWidth)
    Call SetRegUserSettings(strAppName, "Main window position", "Window state", guServOpts.uWinPos.intWinState)
    Call SetRegUserSettings(strAppName, "General", "Last Export path", guServOpts.uConfig.strLastExportPath)
    Call SetRegUserSettings(strAppName, "General", "Console Active", IIf(guServOpts.uFlags.blnConsoleActive, vbChecked, vbUnchecked), REG_DWORD)
    Call SetRegUserSettings(strAppName, "General", "Workstation Size", guServOpts.uFlags.intWksnSize)
    Exit Sub
    
subSetUserSettingsError:
    Call funErrorManager("modFunctions.subSetUserSettingsError")
End Sub


'*******************************************************************************
Public Sub subInitWorkstations()
Dim i As Integer, intMaxMacs As Integer
Dim lngNewTop As Long, lngNewLeft As Long
Dim sgnNumSmallRows As Single
On Error GoTo subInitWorkstationsError
    
    ' Init stuff
    Screen.MousePointer = vbHourglass
    frmMain.ucWorkstations(1).State = cliNotPresent
    frmMain.mnuCrdStartSessionWorkstation(1).enabled = False
    frmMain.mnuCrdStartSessionWorkstation(1).Caption = funLoadResString(10588) & funLoadResString(10589) & 1
    For i = 2 To 20
        frmMain.mnuCrdStartSessionWorkstation(i).Visible = False
    Next i
    For i = 2 To frmMain.ucWorkstations.Count
        Unload frmMain.ucWorkstations(i)
    Next i
    
    For i = 2 To guServOpts.uWorkstations.lngNumWorkstations
        Load frmMain.ucWorkstations(i)
        frmMain.ucWorkstations(i).State = cliNotPresent
        frmMain.ucWorkstations(i).Active = False
        If i <= 20 Then
            frmMain.mnuCrdStartSessionWorkstation(i).Visible = True
            frmMain.mnuCrdStartSessionWorkstation(i).enabled = False
            frmMain.mnuCrdStartSessionWorkstation(i).Caption = funLoadResString(10588) & funLoadResString(10589) & i
        End If
    Next i
    
    Select Case guServOpts.uFlags.intWksnSize
    Case enuWksnSize.WksnSizeNormal
        frmMain.fsbWksn.Max = guServOpts.uWorkstations.lngNumWorkstations
        frmMain.fraWksn.height = frmMain.ucWorkstations(1).height * guServOpts.uWorkstations.lngNumWorkstations + 60 * (guServOpts.uWorkstations.lngNumWorkstations - 1)
        For i = 1 To frmMain.ucWorkstations.Count
            lngNewTop = (i - 1) * (frmMain.ucWorkstations(1).height + 60)
            frmMain.ucWorkstations(i).top = lngNewTop
            frmMain.ucWorkstations(i).width = WKSN_GUI_SIZE_NORMAL
            frmMain.ucWorkstations(i).left = 0
            frmMain.ucWorkstations(i).size = enuWksnSize.WksnSizeNormal
            frmMain.ucWorkstations(i).Visible = True
        Next i
    
    Case enuWksnSize.WksnSizeSmall
        sgnNumSmallRows = guServOpts.uWorkstations.lngNumWorkstations / 2
        sgnNumSmallRows = IIf(sgnNumSmallRows > Int(guServOpts.uWorkstations.lngNumWorkstations / 2), Int(guServOpts.uWorkstations.lngNumWorkstations / 2) + 1, sgnNumSmallRows)
        frmMain.fsbWksn.Max = sgnNumSmallRows
        frmMain.fraWksn.height = frmMain.ucWorkstations(1).height * sgnNumSmallRows + 60 * (sgnNumSmallRows - 1)
        For i = 1 To frmMain.ucWorkstations.Count
            If CSng(i / 2) = Int(i / 2) Then
                lngNewTop = frmMain.ucWorkstations(i - 1).top
                lngNewLeft = WKSN_GUI_SIZE_SMALL + 60
            Else
                lngNewTop = Int((i - 1) / 2) * (frmMain.ucWorkstations(1).height + 60)
                lngNewLeft = 0
            End If
            
            frmMain.ucWorkstations(i).top = lngNewTop
            frmMain.ucWorkstations(i).width = WKSN_GUI_SIZE_SMALL
            frmMain.ucWorkstations(i).left = lngNewLeft
            frmMain.ucWorkstations(i).size = enuWksnSize.WksnSizeSmall
            frmMain.ucWorkstations(i).Visible = True
        Next i
    
    End Select
    
'    On Error Resume Next
    intMaxMacs = funArraySize(guServOpts.uWorkstations.strMacAddresses)
    intMaxMacs = IIf(intMaxMacs / 3 < guServOpts.uWorkstations.lngNumWorkstations, intMaxMacs, guServOpts.uWorkstations.lngNumWorkstations * 3)
    Set guMacAddresses = New Collection
    If intMaxMacs > 0 Then
        For i = 1 To intMaxMacs Step 3
            Call subUpdateWksnList(guServOpts.uWorkstations.strMacAddresses(i - 1), guServOpts.uWorkstations.strMacAddresses(i), funDateTimefromNet(guServOpts.uWorkstations.strMacAddresses(i + 1)), False, False)
        Next i
    End If
    
    Call subResizeWorkstations
    Screen.MousePointer = vbNormal
    Exit Sub
    
subInitWorkstationsError:
    Screen.MousePointer = vbNormal
    Call funErrorManager("modFunctions.subInitWorkstationsError")
End Sub


'*******************************************************************************
Public Sub subUpdateUIWorkstations()
Dim i As Integer, j As Integer
Dim uTmpWorkstation As New clsWorkstation
Dim uEmptyWorkstation As New clsWorkstation
Dim uEmptyCard As New clsCard
Dim uEmptyClient As New clsClient
Dim uMsg As New clsStatusBarMessage
Dim blnAnyActiveWorkstation As Boolean
Dim blnAnySelectedWorkstation As Boolean
On Error GoTo subUpdateUIWorkstationsError

    Set uEmptyWorkstation.CurrentCard = uEmptyCard
    Set uEmptyWorkstation.CurrentClient = uEmptyClient

'    Set guWorkstationsGUI = New Collection
'    For i = 1 To frmMain.ucWorkstations.count
'        Set uWksnGUI = New clsWorkstationGUI
'        Call guWorkstationsGUI.Add(uWksnGUI)
'    Next i
'    For i = 1 To guMacAddresses.count
'        guWorkstationsGUI(i).WorkstationName = guMacAddresses(i).WorkstationName
'        guWorkstationsGUI(i).MacAddress = guMacAddresses(i).WorkstationName
'    Next i



    'Sort workstations by Name
'    If guWorkstations.count > 1 And Not gblnWorkstationConnexionInProgress Then
'        For i = 1 To guWorkstations.count
'            For j = 1 To guWorkstations.count - 1
'                If guWorkstations(j).WorkstationName > guWorkstations(j + 1).WorkstationName Then
'                    Set uTmpWorkstation = guWorkstations(j)
'                    Call guWorkstations.Remove(j)
'                    Call guWorkstations.Add(uTmpWorkstation, uTmpWorkstation.Key, , j)
'                End If
'            Next j
'        Next i
'    End If
    
    blnAnyActiveWorkstation = False
    For i = 1 To guMacAddresses.Count
        Set uTmpWorkstation = funFindWorkstationByMac(guMacAddresses(i))
        If Not uTmpWorkstation.AuthenticationInProgress Then
            With frmMain.ucWorkstations(i)
                .WorkstationIndex = i
                .Caption = uTmpWorkstation.WorkstationName
                .Status = uTmpWorkstation.Status
                If i <= 20 Then frmMain.mnuCrdStartSessionWorkstation(i).Caption = funLoadResString(10588) & uTmpWorkstation.WorkstationName
                .ConnectionIndex = uTmpWorkstation.Key
                .ComputerName = IIf(Len(uTmpWorkstation.ComputerName) > 0, IIf(guServOpts.uFlags.intWksnSize = enuWksnSize.WksnSizeSmall, uTmpWorkstation.WorkstationName & " - ", "") & "\\" & uTmpWorkstation.ComputerName, "")
                Select Case uTmpWorkstation.SessionType
                Case enuSessionType.Card
                    .CustomerID = funLoadResString(10850) & uTmpWorkstation.CurrentCard.ID
                    .CustomerInfo = funLoadResString(10188) & " " & IIf(uTmpWorkstation.UnlimitedPlay, funLoadResString(11000), funFormatTime(uTmpWorkstation.RemainingTimeAparent))
                    If uTmpWorkstation.State <> enuWksnState.Blocked_follow Then
                        If uTmpWorkstation.CurrentCard.CardType <> "" Then
                            .SessionType = cliCardWarn
                        Else
                            .SessionType = cliCard
                        End If
                    Else
                        .SessionType = cliFollow
                    End If
                
                Case enuSessionType.Client
                    .CustomerID = funLoadResString(10849) & uTmpWorkstation.CurrentClient.login
                    .CustomerInfo = funLoadResString(10188) & " " & IIf(uTmpWorkstation.UnlimitedPlay, funLoadResString(11000), funFormatTime(uTmpWorkstation.RemainingTimeAparent))
                    If uTmpWorkstation.State <> enuWksnState.Blocked_follow Then
                        .SessionType = cliClient
                    Else
                        .SessionType = cliFollow
                    End If
                
                Case enuSessionType.Selfservice
                    .CustomerID = funLoadResString(10203)
                    .CustomerInfo = funLoadResString(10189) & " " & funFormatCurrencyString(uTmpWorkstation.UserBalance, guServOpts.uConfig.strMonFormat, guServOpts.uConfig.strMonSymbol)
                    If uTmpWorkstation.State <> enuWksnState.Blocked_follow Then
                        .SessionType = cliSelfservice
                    Else
                        .SessionType = cliFollow
                    End If
                
                Case enuSessionType.SelfServicePrePaid
                    .CustomerID = funLoadResString(10984)
                    .CustomerInfo = funLoadResString(10188) & " " & funFormatTime(uTmpWorkstation.RemainingTimeAparent)
                    If uTmpWorkstation.State <> enuWksnState.Blocked_follow Then
                        .SessionType = cliSelfservice
                    Else
                        .SessionType = cliFollow
                    End If
                
                Case enuSessionType.Free
                    .CustomerID = ""
                    .CustomerInfo = ""
                    If uTmpWorkstation.State <> enuWksnState.Blocked_follow Then
                        .SessionType = cliFree
                    Else
                        .SessionType = cliFollow
                    End If
                
                Case enuSessionType.Unknown
                    .CustomerID = ""
                    .CustomerInfo = ""
                    .SessionType = cliSelfservice
                
                End Select
                
                Select Case uTmpWorkstation.State
                Case enuWksnState.NotPresent
                    .State = cliNotPresent
                    If i <= 20 Then frmMain.mnuCrdStartSessionWorkstation(i).enabled = False
                
                Case enuWksnState.NotActive
                    .State = cliNonActive
                    If i <= 20 Then frmMain.mnuCrdStartSessionWorkstation(i).enabled = False
                
                Case enuWksnState.Ready
                    .State = cliReady
                    If i <= 20 Then frmMain.mnuCrdStartSessionWorkstation(i).enabled = True
                    blnAnyActiveWorkstation = True
        
                Case enuWksnState.UsedConnected
                    .State = cliUsedConnected
                    If i <= 20 Then frmMain.mnuCrdStartSessionWorkstation(i).enabled = False
                
                Case enuWksnState.UsedDisconnected
                    .State = cliUsedDisconnected
                    If i <= 20 Then frmMain.mnuCrdStartSessionWorkstation(i).enabled = False
                
                Case enuWksnState.UsedFree
                    .State = cliUsedFree
                    If i <= 20 Then frmMain.mnuCrdStartSessionWorkstation(i).enabled = False
                
                Case enuWksnState.Blocked_NoCredit
                    .State = cliBlocked_NoCredit
                    If i <= 20 Then frmMain.mnuCrdStartSessionWorkstation(i).enabled = False
                
                Case enuWksnState.Blocked_Paused
                    .State = cliBlocked_Paused
                    If i <= 20 Then frmMain.mnuCrdStartSessionWorkstation(i).enabled = False
                
                Case enuWksnState.Blocked_Locked
                    .State = cliBlocked_Locked
                    If i <= 20 Then frmMain.mnuCrdStartSessionWorkstation(i).enabled = False
                
                Case enuWksnState.Blocked_follow
                    .State = cliBlocked_follow
                    If i <= 20 Then frmMain.mnuCrdStartSessionWorkstation(i).enabled = False
                
                Case enuWksnState.Blocked_SessionEnd
                    .State = cliBlocked_SessionEnd
                    If i <= 20 Then frmMain.mnuCrdStartSessionWorkstation(i).enabled = False
                
                Case enuWksnState.Blocked_Deactivated
                    .State = cliBlocked_Deactivated
                    If i <= 20 Then frmMain.mnuCrdStartSessionWorkstation(i).enabled = False
                
                End Select
                
            End With
        End If
        If frmMain.ucWorkstations(i).Active Then
            If frmMain.mnuFunctions(1).Checked Then
                uTmpWorkstation.Status = enuStatus.cliNormal
                frmMain.ucWorkstations(i).Status = uTmpWorkstation.Status
                Call subRemoveStatusBarMsg("WKSN_MSG_" & uTmpWorkstation.ComputerMacAddress)
            End If
            Set guSelectedWorkstation = uTmpWorkstation
            If Not uTmpWorkstation.AuthenticationInProgress Then Call subUpdateUISelectedWorkstation(guSelectedWorkstation)
            blnAnySelectedWorkstation = True
        End If
    Next i
    frmMain.mnuCrdStartSessionFirstCilent.enabled = blnAnyActiveWorkstation
    frmMain.mnuCrd(6).enabled = blnAnyActiveWorkstation
    frmMain.mnuCrdDD(6).Visible = blnAnyActiveWorkstation
    frmMain.mnuCli(6).enabled = blnAnyActiveWorkstation
    frmMain.mnuCliDD(6).Visible = blnAnyActiveWorkstation
    
    frmMain.cmdCrdButtons(6).enabled = blnAnyActiveWorkstation
    frmMain.cmdCrdButtons(7).enabled = blnAnyActiveWorkstation
    frmMain.cmdCliButtons(6).enabled = blnAnyActiveWorkstation
    frmMain.cmdCliButtons(7).enabled = blnAnyActiveWorkstation
    
    If guWorkstations.Count < frmMain.ucWorkstations.Count Then
        uMsg.ImageIndex = enuStatus.cliCritical
        uMsg.Text = funLoadResString(10204)
        Call funAddStatusBarMsg(uMsg, "NO_WORKSTATION")
    Else
        Call subRemoveStatusBarMsg("NO_WORKSTATION")
    End If
    If guMacAddresses.Count < frmMain.ucWorkstations.Count Then
        For i = guMacAddresses.Count + 1 To frmMain.ucWorkstations.Count
                If i <= 20 Then
                    frmMain.mnuCrdStartSessionWorkstation(i).enabled = False
                    frmMain.mnuCrdStartSessionWorkstation(i).Caption = funLoadResString(10588) & funLoadResString(10589) & i
                End If
            With frmMain.ucWorkstations(i)
                .State = cliNotPresent
                .Caption = "" 'funLoadResString(10598) & i
                .ComputerName = ""
            End With
            If frmMain.ucWorkstations(i).Active Then
                Set guSelectedWorkstation = uEmptyWorkstation
                Call subUpdateUISelectedWorkstation(guSelectedWorkstation)
                blnAnySelectedWorkstation = True
            End If
        Next i
    End If
    If Not blnAnySelectedWorkstation Then Call subUpdateUIConsole
    Exit Sub
    
subUpdateUIWorkstationsError:
    Call funErrorManager("modFunctions.subUpdateUIWorkstationsError")
End Sub


'*******************************************************************************
Public Sub subUpdateUIConsole()
Dim dteLargestRemaining As Date
Dim strLargestRemaining As String
Dim dteSmallestRemaining As Date
Dim strSmallestRemaining As String
Dim blnAnyCustomer As Boolean
Dim uWorkstation As New clsWorkstation
Dim intNumWksnPresent As Integer
Dim intNumWksnUsed As Integer
On Error GoTo subUpdateUIConsoleError
    
    With frmMain
        For Each uWorkstation In guWorkstations
            If Not uWorkstation.UnlimitedPlay Then
                Select Case uWorkstation.SessionType
                Case enuSessionType.Card, enuSessionType.Client, enuSessionType.SelfServicePrePaid
                    blnAnyCustomer = True
                    If uWorkstation.RemainingTimeAparent > dteLargestRemaining Then
                        dteLargestRemaining = uWorkstation.RemainingTimeAparent
                        strLargestRemaining = uWorkstation.WorkstationName
                    End If
    '            Case enuSessionType.Client
    '                blnAnyCustomer = True
    '                If uWorkstation.CurrentClient.TimeCredit > dteLargestRemaining Then
    '                    dteLargestRemaining = uWorkstation.CurrentClient.TimeCredit
    '                    strLargestRemaining = uWorkstation.WorkstationName
    '                End If
                End Select
            End If
        Next uWorkstation
        dteSmallestRemaining = dteLargestRemaining
        strSmallestRemaining = strLargestRemaining
        For Each uWorkstation In guWorkstations
            If uWorkstation.State > enuWksnState.NotPresent Then intNumWksnPresent = intNumWksnPresent + 1
            If uWorkstation.State >= enuWksnState.UsedDisconnected Then
                intNumWksnUsed = intNumWksnUsed + 1
                If Not uWorkstation.UnlimitedPlay Then
                    Select Case uWorkstation.SessionType
                    Case enuSessionType.Card, enuSessionType.Client, enuSessionType.SelfServicePrePaid
                        If uWorkstation.RemainingTimeAparent < dteSmallestRemaining Then
                            dteSmallestRemaining = uWorkstation.RemainingTimeAparent
                            strSmallestRemaining = uWorkstation.WorkstationName
                        End If
        '                Case enuSessionType.Client
        '                    If uWorkstation.CurrentClient.TimeCredit < dteSmallestRemaining Then
        '                        dteSmallestRemaining = uWorkstation.CurrentClient.TimeCredit
        '                        strSmallestRemaining = uWorkstation.WorkstationName
        '                    End If
                    End Select
                End If
            End If
        Next uWorkstation
        .lblConsoleSel(1).Caption = " " & funLoadResString(10629) & funFormatTime(Now - guServOpts.uFlags.dteOnTime, enuTimeFormat.HH_NN_SS) & " - " & funLoadResString(10633) & intNumWksnUsed & funLoadResString(10631) & " / " & intNumWksnPresent & funLoadResString(10632)
        .lblConsoleSel(2).Caption = " " & IIf(guServOpts.uFlags.blnInetCnx, funLoadResString(10627), funLoadResString(10628))
        .lblConsoleSel(3) = IIf(blnAnyCustomer, " " & strSmallestRemaining & " - " & funFormatTime(dteSmallestRemaining, enuTimeFormat.HH_NN_SS), "")
    End With
    Exit Sub
    
subUpdateUIConsoleError:
    Call funErrorManager("modFunctions.subUpdateUIConsoleError")
End Sub


'*******************************************************************************
Public Sub subUpdateUISelectedWorkstation(uWorkstation As clsWorkstation)
Dim uWOL As New clsMagicPacket
Dim varTmpHist As Variant, strHist As String
Static lngHistLen As Long
Static intLastSessionType As Integer
Static blnWasVisible As Boolean
On Error GoTo subUpdateUISelectedWorkstationError
    
    With frmMain
        .lblWksnSel(0).Caption = " " & uWorkstation.WorkstationName
        Select Case uWorkstation.State
        Case enuWksnState.NotPresent
            .lblWksnSel(1).Caption = " " & funLoadResString(10171)
            .fraWksnSessionInfo.Visible = False
            .cmdWksnButtons(0).Visible = False
            .cmdWksnButtons(2).Visible = False
            .cmdWksnButtons(3).Visible = False
            .cmdWksnButtons(9).Visible = False
            .cmdWksnButtons(11).Visible = False
            .cmdWksnButtons(12).Visible = False
            .cmdWksnButtons(4).Visible = False
            .cmdWksnButtons(1).Visible = False
            .cmdWksnButtons(5).Visible = False
            .cmdWksnButtons(6).Visible = False
            .cmdWksnButtons(8).Visible = False
            .cmdWksnButtons(7).Visible = False
            .cmdWksnButtons(10).Visible = False
            .chkWksnConnect.Visible = False
            .cmdWksnButtons(13).enabled = False
            .txtWksnChat.enabled = False
            .cmdWksnButtons(14).Visible = False
            .cmdWksnButtons(15).Visible = False
            
        Case enuWksnState.NotActive
            .lblWksnSel(1).Caption = " " & funLoadResString(10172)
            .fraWksnSessionInfo.Visible = False
            .cmdWksnButtons(0).Visible = True
            .cmdWksnButtons(2).Visible = False
            .cmdWksnButtons(3).Visible = False
            .cmdWksnButtons(9).Visible = False
            .cmdWksnButtons(10).Visible = False
            .cmdWksnButtons(11).Visible = False
            .cmdWksnButtons(12).Visible = False
            .cmdWksnButtons(4).Visible = False
            .cmdWksnButtons(1).Visible = False
            .cmdWksnButtons(5).Visible = False
            .cmdWksnButtons(6).Visible = False
            .cmdWksnButtons(8).Visible = False
            .cmdWksnButtons(7).Visible = False
            .chkWksnConnect.Visible = False
            .cmdWksnButtons(13).enabled = False
            .txtWksnChat.enabled = False
            .cmdWksnButtons(14).Visible = False
            .cmdWksnButtons(15).Visible = False
            
        Case enuWksnState.Ready
            .lblWksnSel(1).Caption = " " & funLoadResString(10173)
            .fraWksnSessionInfo.Visible = False
            .cmdWksnButtons(0).Visible = False
            .cmdWksnButtons(2).Visible = True
            .cmdWksnButtons(2).enabled = True
            .cmdWksnButtons(3).Visible = True
            .cmdWksnButtons(9).Visible = False
            .cmdWksnButtons(10).Visible = False
            .cmdWksnButtons(11).Visible = False
            .cmdWksnButtons(12).Visible = False
            .cmdWksnButtons(4).Visible = False
            .cmdWksnButtons(1).Visible = False
            .cmdWksnButtons(5).Visible = False
            .cmdWksnButtons(6).Visible = False
            .cmdWksnButtons(8).Visible = False
            .cmdWksnButtons(7).Visible = False
            .chkWksnConnect.Visible = True
            .chkWksnConnect.value = uWorkstation.ConnectionForced
            .cmdWksnButtons(13).enabled = False
            .txtWksnChat.enabled = False
            .cmdWksnButtons(14).Visible = guServOpts.uWorkstations.intActivateSelfServicePrePaid = vbChecked
            .cmdWksnButtons(15).Visible = False
            
        Case enuWksnState.UsedConnected
            .lblWksnSel(1).Caption = " " & funLoadResString(10174)
            .fraWksnSessionInfo.Visible = True
            .cmdWksnButtons(0).Visible = False
            .cmdWksnButtons(2).Visible = True
            .cmdWksnButtons(2).enabled = uWorkstation.SessionType = enuSessionType.Card Or uWorkstation.SessionType = enuSessionType.Client Or uWorkstation.SessionType = enuSessionType.SelfServicePrePaid
            .cmdWksnButtons(3).Visible = False
            .cmdWksnButtons(9).Visible = uWorkstation.SessionType = enuSessionType.Card Or uWorkstation.SessionType = enuSessionType.Client Or uWorkstation.SessionType = enuSessionType.SelfServicePrePaid
            .cmdWksnButtons(10).Visible = uWorkstation.SessionType = enuSessionType.Card Or uWorkstation.SessionType = enuSessionType.Client
            .cmdWksnButtons(11).Visible = uWorkstation.SessionType = enuSessionType.Card
            .cmdWksnButtons(12).Visible = uWorkstation.SessionType = enuSessionType.Client
            .cmdWksnButtons(4).Visible = True
            .cmdWksnButtons(1).Visible = False
            .cmdWksnButtons(5).Visible = True
            .cmdWksnButtons(6).Visible = True
            .cmdWksnButtons(8).Visible = False
            .cmdWksnButtons(7).Visible = False
            .chkWksnConnect.Visible = True
            .chkWksnConnect.value = uWorkstation.ConnectionForced
            .cmdWksnButtons(13).enabled = True
            .txtWksnChat.enabled = True
            .cmdWksnButtons(14).Visible = False
            .cmdWksnButtons(15).Visible = False
            
        Case enuWksnState.UsedDisconnected
            .lblWksnSel(1).Caption = " " & funLoadResString(10175)
            .fraWksnSessionInfo.Visible = True
            .cmdWksnButtons(0).Visible = False
            .cmdWksnButtons(2).Visible = True
            .cmdWksnButtons(2).enabled = uWorkstation.SessionType = enuSessionType.Card Or uWorkstation.SessionType = enuSessionType.Client Or uWorkstation.SessionType = enuSessionType.SelfServicePrePaid
            .cmdWksnButtons(3).Visible = False
            .cmdWksnButtons(9).Visible = uWorkstation.SessionType = enuSessionType.Card Or uWorkstation.SessionType = enuSessionType.Client Or uWorkstation.SessionType = enuSessionType.SelfServicePrePaid
            .cmdWksnButtons(10).Visible = uWorkstation.SessionType = enuSessionType.Card Or uWorkstation.SessionType = enuSessionType.Client
            .cmdWksnButtons(11).Visible = uWorkstation.SessionType = enuSessionType.Card
            .cmdWksnButtons(12).Visible = uWorkstation.SessionType = enuSessionType.Client
            .cmdWksnButtons(4).Visible = True
            .cmdWksnButtons(1).Visible = False
            .cmdWksnButtons(5).Visible = True
            .cmdWksnButtons(6).Visible = True
            .cmdWksnButtons(8).Visible = False
            .cmdWksnButtons(7).Visible = False
            .chkWksnConnect.Visible = True
            .chkWksnConnect.value = uWorkstation.ConnectionForced
            .cmdWksnButtons(13).enabled = True
            .txtWksnChat.enabled = True
            .cmdWksnButtons(14).Visible = False
            .cmdWksnButtons(15).Visible = False
        
        Case enuWksnState.UsedFree
            .lblWksnSel(1).Caption = " " & funLoadResString(10620)
            .fraWksnSessionInfo.Visible = True
            .cmdWksnButtons(0).Visible = False
            .cmdWksnButtons(2).Visible = True
            .cmdWksnButtons(2).enabled = True
            .cmdWksnButtons(3).Visible = False
            .cmdWksnButtons(9).Visible = False
            .cmdWksnButtons(11).Visible = False
            .cmdWksnButtons(12).Visible = False
            .cmdWksnButtons(10).Visible = False
            .cmdWksnButtons(4).Visible = True
            .cmdWksnButtons(1).Visible = False
            .cmdWksnButtons(5).Visible = True
            .cmdWksnButtons(6).Visible = True
            .cmdWksnButtons(8).Visible = False
            .cmdWksnButtons(7).Visible = False
            .chkWksnConnect.Visible = False
            .chkWksnConnect.value = uWorkstation.ConnectionForced
            .cmdWksnButtons(13).enabled = False
            .txtWksnChat.enabled = False
            .cmdWksnButtons(14).Visible = False
            .cmdWksnButtons(15).Visible = False
            
        Case enuWksnState.Blocked_NoCredit
            .lblWksnSel(1).Caption = " " & funLoadResString(10176)
            .fraWksnSessionInfo.Visible = True
            .cmdWksnButtons(0).Visible = False
            .cmdWksnButtons(2).Visible = True
            .cmdWksnButtons(2).enabled = True
            .cmdWksnButtons(3).Visible = False
            .cmdWksnButtons(9).Visible = True
            .cmdWksnButtons(10).Visible = True
            .cmdWksnButtons(11).Visible = True
            .cmdWksnButtons(12).Visible = True
            .cmdWksnButtons(4).Visible = True
            .cmdWksnButtons(1).Visible = False
            .cmdWksnButtons(5).Visible = False
            .cmdWksnButtons(6).Visible = False
            .cmdWksnButtons(8).Visible = False
            .cmdWksnButtons(7).Visible = True
            .chkWksnConnect.Visible = True
            .cmdWksnButtons(13).enabled = False
            .txtWksnChat.enabled = False
            .cmdWksnButtons(14).Visible = False
            .cmdWksnButtons(15).Visible = False
        
        Case enuWksnState.Blocked_Paused
            .lblWksnSel(1).Caption = " " & funLoadResString(10177)
            .fraWksnSessionInfo.Visible = True
            .cmdWksnButtons(0).Visible = False
            .cmdWksnButtons(2).Visible = True
            .cmdWksnButtons(2).enabled = uWorkstation.SessionType = enuSessionType.Card Or uWorkstation.SessionType = enuSessionType.Client Or uWorkstation.SessionType = enuSessionType.SelfServicePrePaid
            .cmdWksnButtons(3).Visible = False
            .cmdWksnButtons(9).Visible = uWorkstation.SessionType = enuSessionType.Card Or uWorkstation.SessionType = enuSessionType.Client Or uWorkstation.SessionType = enuSessionType.SelfServicePrePaid
            .cmdWksnButtons(10).Visible = uWorkstation.SessionType = enuSessionType.Card Or uWorkstation.SessionType = enuSessionType.Client
            .cmdWksnButtons(11).Visible = uWorkstation.SessionType = enuSessionType.Card
            .cmdWksnButtons(12).Visible = uWorkstation.SessionType = enuSessionType.Client
            .cmdWksnButtons(4).Visible = True
            .cmdWksnButtons(1).Visible = False
            .cmdWksnButtons(5).Visible = False
            .cmdWksnButtons(6).Visible = False
            .cmdWksnButtons(8).Visible = True
            .cmdWksnButtons(7).Visible = False
            .chkWksnConnect.Visible = uWorkstation.SessionType <> enuSessionType.Free
            .cmdWksnButtons(13).enabled = False
            .txtWksnChat.enabled = False
            .cmdWksnButtons(14).Visible = False
            .cmdWksnButtons(15).Visible = False
        
        Case enuWksnState.Blocked_Locked
            .lblWksnSel(1).Caption = " " & funLoadResString(10596)
            .fraWksnSessionInfo.Visible = True
            .cmdWksnButtons(0).Visible = False
            .cmdWksnButtons(2).Visible = True
            .cmdWksnButtons(2).enabled = uWorkstation.SessionType = enuSessionType.Card Or uWorkstation.SessionType = enuSessionType.Client Or uWorkstation.SessionType = enuSessionType.SelfServicePrePaid
            .cmdWksnButtons(3).Visible = False
            .cmdWksnButtons(9).Visible = uWorkstation.SessionType = enuSessionType.Card Or uWorkstation.SessionType = enuSessionType.Client Or uWorkstation.SessionType = enuSessionType.SelfServicePrePaid
            .cmdWksnButtons(10).Visible = uWorkstation.SessionType = enuSessionType.Card Or uWorkstation.SessionType = enuSessionType.Client
            .cmdWksnButtons(11).Visible = uWorkstation.SessionType = enuSessionType.Card
            .cmdWksnButtons(12).Visible = uWorkstation.SessionType = enuSessionType.Client
            .cmdWksnButtons(4).Visible = True
            .cmdWksnButtons(1).Visible = False
            .cmdWksnButtons(5).Visible = False
            .cmdWksnButtons(6).Visible = False
            .cmdWksnButtons(8).Visible = True
            .cmdWksnButtons(7).Visible = False
            .chkWksnConnect.Visible = uWorkstation.SessionType <> enuSessionType.Free
            .cmdWksnButtons(13).enabled = False
            .txtWksnChat.enabled = False
            .cmdWksnButtons(14).Visible = False
            .cmdWksnButtons(15).Visible = False
        
        Case enuWksnState.Blocked_follow
            .lblWksnSel(1).Caption = " " & funLoadResString(10626)
            .fraWksnSessionInfo.Visible = True
            .cmdWksnButtons(0).Visible = False
            .cmdWksnButtons(2).Visible = True
            .cmdWksnButtons(2).enabled = uWorkstation.SessionType = enuSessionType.Card Or uWorkstation.SessionType = enuSessionType.Client Or uWorkstation.SessionType = enuSessionType.SelfServicePrePaid
            .cmdWksnButtons(3).Visible = False
            .cmdWksnButtons(9).Visible = uWorkstation.SessionType = enuSessionType.Card Or uWorkstation.SessionType = enuSessionType.Client Or uWorkstation.SessionType = enuSessionType.SelfServicePrePaid
            .cmdWksnButtons(10).Visible = uWorkstation.SessionType = enuSessionType.Card Or uWorkstation.SessionType = enuSessionType.Client
            .cmdWksnButtons(11).Visible = uWorkstation.SessionType = enuSessionType.Card
            .cmdWksnButtons(12).Visible = uWorkstation.SessionType = enuSessionType.Client
            .cmdWksnButtons(4).Visible = True
            .cmdWksnButtons(1).Visible = False
            .cmdWksnButtons(5).Visible = False
            .cmdWksnButtons(6).Visible = False
            .cmdWksnButtons(8).Visible = True
            .cmdWksnButtons(7).Visible = False
            .chkWksnConnect.Visible = uWorkstation.SessionType <> enuSessionType.Free
            .cmdWksnButtons(13).enabled = False
            .txtWksnChat.enabled = False
            .cmdWksnButtons(14).Visible = False
            .cmdWksnButtons(15).Visible = False
        
        Case enuWksnState.Blocked_SessionEnd
            .lblWksnSel(1).Caption = " " & funLoadResString(10178)
            .fraWksnSessionInfo.Visible = True
            .cmdWksnButtons(0).Visible = False
            .cmdWksnButtons(2).Visible = False
            .cmdWksnButtons(3).Visible = False
            .cmdWksnButtons(9).Visible = False
            .cmdWksnButtons(10).Visible = False
            .cmdWksnButtons(11).Visible = False
            .cmdWksnButtons(12).Visible = False
            .cmdWksnButtons(4).Visible = False
            .cmdWksnButtons(1).Visible = True
            .cmdWksnButtons(5).Visible = False
            .cmdWksnButtons(6).Visible = False
            .cmdWksnButtons(8).Visible = False
            .cmdWksnButtons(7).Visible = False
            .chkWksnConnect.Visible = False
            .cmdWksnButtons(13).enabled = False
            .txtWksnChat.enabled = False
            .cmdWksnButtons(14).Visible = False
            .cmdWksnButtons(15).Visible = (uWorkstation.ConnectedTime + uWorkstation.ConnectedTimeOffPeak + uWorkstation.DisconnectedTime) < guServOpts.uPostPaidSs.dteGraceTime
        
        Case enuWksnState.Blocked_Deactivated
            .lblWksnSel(1).Caption = " " & funLoadResString(10179)
            .fraWksnSessionInfo.Visible = False
            .cmdWksnButtons(0).Visible = True
            .cmdWksnButtons(2).Visible = False
            .cmdWksnButtons(3).Visible = False
            .cmdWksnButtons(9).Visible = False
            .cmdWksnButtons(10).Visible = False
            .cmdWksnButtons(11).Visible = False
            .cmdWksnButtons(12).Visible = False
            .cmdWksnButtons(4).Visible = False
            .cmdWksnButtons(1).Visible = False
            .cmdWksnButtons(5).Visible = False
            .cmdWksnButtons(6).Visible = False
            .cmdWksnButtons(8).Visible = False
            .cmdWksnButtons(7).Visible = False
            .chkWksnConnect.Visible = False
            .cmdWksnButtons(13).enabled = False
            .txtWksnChat.enabled = False
            .cmdWksnButtons(14).Visible = False
            .cmdWksnButtons(15).Visible = False
        
        End Select
        If uWorkstation.State >= enuWksnState.UsedDisconnected Or uWorkstation.State = enuWksnState.Blocked_SessionEnd Then
            .lblWksnSel(3).Caption = " " & funFormatTime(uWorkstation.ConnectedTime + uWorkstation.ConnectedTimeOffPeak + uWorkstation.DisconnectedTime)
            .lblWksnSel(3).ToolTipText = funLoadResString(10977, funFormatTime(uWorkstation.ElapsedTime, enuTimeFormat.Hh_Nmn_Ssec))
            .lblWksnSel(4).Caption = " " & uWorkstation.LastVisitedTitle
            .lblWksnSel(4).ToolTipText = uWorkstation.LastVisitedURL
            For Each varTmpHist In uWorkstation.ChatHistory
                If CStr(varTmpHist) <> "" Then strHist = strHist & IIf(strHist <> "", vbNewLine, "") & " " & varTmpHist
            Next varTmpHist
            .txtWksnChatHist.Text = strHist
            If lngHistLen <> Len(strHist) Or Not blnWasVisible Then .txtWksnChatHist.SelStart = Len(.txtWksnChatHist.Text)
            lngHistLen = Len(strHist)
            blnWasVisible = .fraFunctions(1).Visible And .fraWksnControls.Visible
            
            Select Case uWorkstation.SessionType
            Case enuSessionType.Client
                .lblWksnSel(2).Caption = " " & funLoadResString(10832)
                .fraWksnCliInfo.Visible = True
                .fraWksnCrdInfo.Visible = False
                .fraWksnSelfserveInfo.Visible = False
                .fraWksnSelfservePrepaidInfo.Visible = False
                .fraWksnChat.Visible = True
                .fraWksnChat.top = .fraWksnCliInfo.top + .fraWksnCliInfo.height
                .lblWksnSel(13).Caption = " " & uWorkstation.CurrentClient.login
                .lblWksnSel(12) = " " & IIf(uWorkstation.UnlimitedPlay, funLoadResString(11000), funFormatTime(uWorkstation.RemainingTimeAparent))
                .lblWksnSel(11) = " " & uWorkstation.CurrentClient.FullName
                .lbl(45).enabled = "" <> uWorkstation.CurrentClient.Comment
                .lblWksnSel(10) = " " & funLeftNewLines(uWorkstation.CurrentClient.Comment, 0)
                .cmdWksnButtons(9).Picture = .imlWksnRecharge.ListImages(enuCustomerType.Client).Picture
                .cmdWksnButtons(9).ToolTipText = funLoadResString(10989)
                .cmdWksnButtons(10).Picture = .imlWksnDisable.ListImages(enuCustomerType.Client).Picture
                .cmdWksnButtons(10).ToolTipText = funLoadResString(10813)

            Case enuSessionType.Card
                .lblWksnSel(2).Caption = " " & funLoadResString(10202)
                .fraWksnCliInfo.Visible = False
                .fraWksnCrdInfo.Visible = True
                .fraWksnSelfserveInfo.Visible = False
                .fraWksnSelfservePrepaidInfo.Visible = False
                .fraWksnChat.Visible = True
                .fraWksnChat.top = .fraWksnCrdInfo.top + .fraWksnCrdInfo.height
                .lblWksnSel(6).Caption = " " & uWorkstation.CurrentCard.ID
                .lblWksnSel(7) = " " & IIf(uWorkstation.UnlimitedPlay, funLoadResString(11000), funFormatTime(uWorkstation.RemainingTimeAparent))
                .lbl(29).enabled = "" <> uWorkstation.CurrentCard.Comment
                .lblWksnSel(9) = " " & funLeftNewLines(uWorkstation.CurrentCard.Comment, 0)
                .lblWksnSel(8) = " " & IIf(uWorkstation.CurrentCard.CardType = "", funLoadResString(10526), uWorkstation.CurrentCard.CardType)
                .cmdWksnButtons(9).Picture = .imlWksnRecharge.ListImages(enuCustomerType.Card).Picture
                .cmdWksnButtons(9).ToolTipText = funLoadResString(10465)
                .cmdWksnButtons(10).Picture = .imlWksnDisable.ListImages(enuCustomerType.Card).Picture
                .cmdWksnButtons(10).ToolTipText = funLoadResString(10692)
                
            Case enuSessionType.Selfservice
                .lblWksnSel(2).Caption = " " & funLoadResString(10203)
                .fraWksnCliInfo.Visible = False
                .fraWksnCrdInfo.Visible = False
                .fraWksnSelfserveInfo.Visible = True
                .fraWksnSelfservePrepaidInfo.Visible = False
                .fraWksnChat.Visible = True
                .fraWksnChat.top = .fraWksnSelfserveInfo.top + .fraWksnSelfserveInfo.height
                .lblWksnSel(5) = " " & funFormatCurrencyString(uWorkstation.UserBalance, guServOpts.uConfig.strMonFormat, guServOpts.uConfig.strMonSymbol)
            
            Case enuSessionType.SelfServicePrePaid
                .lblWksnSel(2).Caption = " " & funLoadResString(10984)
                .fraWksnCliInfo.Visible = False
                .fraWksnCrdInfo.Visible = False
                .fraWksnSelfserveInfo.Visible = False
                .fraWksnSelfservePrepaidInfo.Visible = True
                .fraWksnChat.Visible = True
                .fraWksnChat.top = .fraWksnSelfserveInfo.top + .fraWksnSelfserveInfo.height
                .lblWksnSel(14) = " " & funFormatTime(uWorkstation.RemainingTimeAparent)
                .cmdWksnButtons(9).Picture = .imlWksnRecharge.ListImages(enuCustomerType.SelfServicePrePaid).Picture
                .cmdWksnButtons(9).ToolTipText = funLoadResString(10990)
            
            Case enuSessionType.Free
                .lblWksnSel(2).Caption = " " & funLoadResString(10621)
                .fraWksnCliInfo.Visible = False
                .fraWksnCrdInfo.Visible = False
                .fraWksnSelfserveInfo.Visible = False
                .fraWksnSelfservePrepaidInfo.Visible = False
                .fraWksnChat.Visible = False
            
            End Select
            If intLastSessionType <> uWorkstation.SessionType Then Call subResizeWorkstations
            intLastSessionType = uWorkstation.SessionType
        End If
        .mnuWksnCrdRecharge.Visible = True
        .mnuWksnConnect.Visible = True
        .mnuWksnPaySession.Visible = True
        .mnuWksnCancelPaySession.Visible = True
        .mnuWksnCrdDisable.Visible = True
        .mnuWksnCrdHistory.Visible = True
        .mnuWksnStartSelfServPrePaid.Visible = True
        .mnuWksnUserRelease.Visible = LCase(guServOpts.uConnectedUser.login) = ADMIN_LOGON
        
        .mnuWksnActivate.enabled = .cmdWksnButtons(0).Visible
        .mnuWksnBloc.enabled = .cmdWksnButtons(2).enabled And .cmdWksnButtons(2).Visible
        .mnuWksnStartSelfServPrePaid.enabled = .cmdWksnButtons(14).Visible
        .mnuWksnCrdRecharge.enabled = .cmdWksnButtons(9).Visible
        .mnuWksnUnblock.enabled = .cmdWksnButtons(3).Visible And guServOpts.uConnectedUser.AccessUnblock
        .mnuWksnClose.enabled = .cmdWksnButtons(4).Visible
        .mnuWksnConnect.enabled = .chkWksnConnect.Visible
        .mnuWksnConnect.Checked = .chkWksnConnect.value = vbChecked
        .mnuWksnPaySession.enabled = .cmdWksnButtons(1).Visible
        .mnuWksnCancelPaySession.enabled = .cmdWksnButtons(15).Visible
        .mnuWksnPause.enabled = .cmdWksnButtons(5).Visible
        .mnuWksnLock.enabled = .cmdWksnButtons(5).Visible
        .mnuWksnFollow.enabled = .cmdWksnButtons(5).Visible
        .mnuWksnPauseOne.enabled = .cmdWksnButtons(5).Visible
        .mnuWksnLockOne.enabled = .cmdWksnButtons(5).Visible
        .mnuWksnPlay.enabled = .cmdWksnButtons(8).Visible
        .mnuWksnCrdDisable.enabled = .cmdWksnButtons(10).Visible
        .mnuWksnCrdHistory.enabled = .cmdWksnButtons(11).Visible Or .cmdWksnButtons(12).Visible
        .mnuWksnUserRelease.enabled = uWorkstation.State > enuWksnState.NotPresent
    
        .mnuWksnActivateDD.Visible = .mnuWksnActivate.enabled
        .mnuWksnBlocDD.Visible = .mnuWksnBloc.enabled
        .mnuWksnCrdRechargeDD.Visible = .mnuWksnCrdRecharge.enabled
        .mnuWksnCrdDisableDD.Visible = .mnuWksnCrdDisable.enabled
        .mnuWksnCrdHistoryDD.Visible = .mnuWksnCrdHistory.enabled
        .mnuWksnCloseDD.Visible = .mnuWksnClose.enabled
        .mnuWksnFollowDD.Visible = .mnuWksnFollow.enabled
        .mnuWksnLockDD.Visible = .mnuWksnLock.enabled
        .mnuWksnPauseDD.Visible = .mnuWksnPause.enabled
        .mnuWksnPaySessionDD.Visible = .mnuWksnPaySession.enabled
        .mnuWksnPlayDD.Visible = .mnuWksnPlay.enabled
        .mnuWksnUnblockDD.Visible = .mnuWksnUnblock.enabled
        .mnuWksnStartSelfServPrePaidDD.Visible = .mnuWksnStartSelfServPrePaid.enabled
        .mnu10.Visible = .mnuWksnUnblockDD.Visible Or .mnuWksnPlayDD.Visible Or .mnuWksnPaySessionDD.Visible Or .mnuWksnPauseDD.Visible Or .mnuWksnLockDD.Visible Or .mnuWksnFollowDD.Visible Or .mnuWksnCloseDD.Visible Or .mnuWksnCrdHistoryDD.Visible Or .mnuWksnCrdDisableDD.Visible Or .mnuWksnCrdRechargeDD.Visible
        .mnuWksnUserReleaseDD.Visible = .mnuWksnUserRelease.enabled And .mnuWksnUserRelease.Visible
        
        If .mnuWksnBootDD.Visible Then
            .mnuWksnShutdownDD.Visible = uWorkstation.State > enuWksnState.NotPresent
            .mnuWksnBootDD.Visible = uWorkstation.State = enuWksnState.NotPresent
        Else
            .mnuWksnBootDD.Visible = uWorkstation.State = enuWksnState.NotPresent
            .mnuWksnShutdownDD.Visible = uWorkstation.State > enuWksnState.NotPresent
        End If
        .mnuWksnBootDD.enabled = "" <> uWOL.ValidateNICAddress(uWorkstation.ComputerMacAddress)

        .mnu9.Visible = (.mnuWksnActivateDD.Visible Or .mnuWksnBlocDD.Visible) And Not .mnuWksnBootDD.Visible
    End With
    Exit Sub
    
subUpdateUISelectedWorkstationError:
    Call funErrorManager("modFunctions.subUpdateUISelectedWorkstationError")
End Sub


'*******************************************************************************
Public Sub subForceCardUpdate(uWorkstation As clsWorkstation)
On Error GoTo subForceCardUpdateError
                        
    grstCards.Index = "PrimaryKey"
    Call grstCards.Seek("=", uWorkstation.CurrentCard.ID)
    If Not grstCards.NoMatch Then
        Call grstCards.Edit
        grstCards!time_credit = uWorkstation.CurrentCard.TimeCredit
        Call grstCards.Update
    End If
    Exit Sub
    
subForceCardUpdateError:
    Call funErrorManager("modFunctions.subForceCardUpdateError")
End Sub


'*******************************************************************************
Public Sub subForceClientUpdate(uWorkstation As clsWorkstation)
On Error GoTo subForceClientUpdateError
                        
    grstClients.Index = "PrimaryKey"
    Call grstClients.Seek("=", uWorkstation.CurrentClient.login)
    If Not grstClients.NoMatch Then
        Call grstClients.Edit
        grstClients!time_credit = uWorkstation.CurrentClient.TimeCredit
        Call grstClients.Update
    End If
    Exit Sub
    
subForceClientUpdateError:
    Call funErrorManager("modFunctions.subForceClientUpdateError")
End Sub


'*******************************************************************************
Public Sub subResizeWorkstations()
Dim lngNewTop As Long
Dim lngNewLeft As Long
Dim lngCliWidth As Long
Dim lngCliHeight As Long
Dim intNumWorkstationRows As Integer
Dim sgnTotalNumWorkstationRows As Single
On Error Resume Next

    lngCliWidth = WKSN_GUI_SIZE_NORMAL + Screen.TwipsPerPixelX '* 2
    lngCliHeight = frmMain.ucWorkstations(1).height
    Select Case guServOpts.uFlags.intWksnSize
    Case enuWksnSize.WksnSizeNormal
        sgnTotalNumWorkstationRows = guServOpts.uWorkstations.lngNumWorkstations
    Case enuWksnSize.WksnSizeSmall
        sgnTotalNumWorkstationRows = guServOpts.uWorkstations.lngNumWorkstations / 2
        sgnTotalNumWorkstationRows = IIf(sgnTotalNumWorkstationRows > Int(guServOpts.uWorkstations.lngNumWorkstations / 2), Int(guServOpts.uWorkstations.lngNumWorkstations / 2) + 1, sgnTotalNumWorkstationRows)
    End Select
    intNumWorkstationRows = Int((frmMain.fraFunctions(1).height - frmMain.fraWksnContainer.top) / (lngCliHeight + 60))
    frmMain.fraWksnContainer.width = lngCliWidth
    If guServOpts.uFlags.blnConsoleActive Then
        frmMain.fraWksnContainer.top = frmMain.fraWksnConsole.height + 60
    Else
        frmMain.fraWksnContainer.top = 0
    End If
    frmMain.fraWksnContainer.height = frmMain.fraFunctions(1).height - frmMain.fraWksnContainer.top
    frmMain.fsbWksn.top = frmMain.fraWksnContainer.top
    frmMain.fsbWksn.enabled = frmMain.fraWksn.height > frmMain.fraWksnContainer.height
    frmMain.fsbWksn.Max = IIf(sgnTotalNumWorkstationRows - intNumWorkstationRows >= frmMain.fsbWksn.Min, sgnTotalNumWorkstationRows - intNumWorkstationRows, frmMain.fsbWksn.Min)
    If frmMain.fsbWksn.value = frmMain.fsbWksn.Max Then
        Call frmMain.fsbWksn_Change
    End If
    frmMain.fraWksn.width = lngCliWidth
    frmMain.fsbWksn.left = lngCliWidth + 30
    frmMain.fsbWksn.height = frmMain.fraWksnContainer.height
    frmMain.fraWksnControls.height = frmMain.fraFunctions(1).height
    frmMain.fraWksnControls.left = frmMain.fsbWksn.left + frmMain.fsbWksn.width + 60
    frmMain.fraWksnControls.width = frmMain.fraFunctions(1).width - frmMain.fraWksn.width - frmMain.fsbWksn.width - 60 * 2 + 30
    frmMain.shpWksnControlsBorder.top = 0
    frmMain.shpWksnControlsBorder.left = 0
    frmMain.shpWksnControlsBorder.height = frmMain.fraFunctions(1).height
    frmMain.shpWksnControlsBorder.width = frmMain.fraWksnControls.width
    frmMain.fraWksnConsoleControls.height = frmMain.fraFunctions(1).height
    frmMain.fraWksnConsoleControls.left = frmMain.fsbWksn.left + frmMain.fsbWksn.width + 60
    frmMain.fraWksnConsoleControls.width = frmMain.fraFunctions(1).width - frmMain.fraWksn.width - frmMain.fsbWksn.width - 60 * 2 + 30
    frmMain.shpWksnConsoleControlsBorder.top = 0
    frmMain.shpWksnConsoleControlsBorder.left = 0
    frmMain.shpWksnConsoleControlsBorder.height = frmMain.fraFunctions(1).height
    frmMain.shpWksnConsoleControlsBorder.width = frmMain.fraWksnControls.width
    frmMain.lblConsoleSel(1).width = frmMain.fraWksnConsoleControls.width - frmMain.lbl(31).width - 60 * 3 - 20
    frmMain.lblConsoleSel(2).width = frmMain.lblConsoleSel(1).width
    frmMain.lblConsoleSel(3).width = frmMain.lblConsoleSel(1).width
    frmMain.lblWksnSel(0).width = frmMain.fraWksnControls.width
    frmMain.lblWksnSel(1).width = frmMain.fraWksnControls.width - frmMain.lbl(0).width - 60 * 3 - 20
    frmMain.lblWksnSel(2).width = frmMain.lblWksnSel(1).width
    frmMain.lblWksnSel(3).width = frmMain.lblWksnSel(1).width
    frmMain.lblWksnSel(4).width = frmMain.lblWksnSel(1).width
    frmMain.lblWksnSel(6).width = frmMain.lblWksnSel(1).width - frmMain.cmdWksnButtons(11).width - 60
    frmMain.cmdWksnButtons(11).left = frmMain.lblWksnSel(1).width + frmMain.lbl(3).width - frmMain.cmdWksnButtons(11).width + 100
    frmMain.lblWksnSel(7).width = frmMain.lblWksnSel(1).width
    frmMain.lblWksnSel(5).width = frmMain.lblWksnSel(1).width
    frmMain.lblWksnSel(9).width = frmMain.lblWksnSel(1).width
    frmMain.lblWksnSel(8).width = frmMain.lblWksnSel(1).width
    frmMain.lblWksnSel(13).width = frmMain.lblWksnSel(1).width - frmMain.cmdWksnButtons(12).width - 60
    frmMain.cmdWksnButtons(12).left = frmMain.lblWksnSel(1).width + frmMain.lbl(48).width - frmMain.cmdWksnButtons(12).width + 100
    frmMain.lblWksnSel(12).width = frmMain.lblWksnSel(1).width
    frmMain.lblWksnSel(11).width = frmMain.lblWksnSel(1).width
    frmMain.lblWksnSel(10).width = frmMain.lblWksnSel(1).width
    frmMain.lblWksnSel(14).width = frmMain.lblWksnSel(1).width
    frmMain.fraWksnButtons.top = frmMain.shpWksnControlsBorder.height - frmMain.fraWksnButtons.height - 60
    frmMain.fraWksnButtons.left = frmMain.shpWksnControlsBorder.left + 60
    frmMain.fraWksnButtons.width = frmMain.shpWksnControlsBorder.width - 120
    frmMain.fraWksnSessionInfo.height = frmMain.fraWksnButtons.top - frmMain.fraWksnSessionInfo.top - 60
    frmMain.fraWksnSessionInfo.width = frmMain.fraWksnControls.width - 120
    frmMain.fraWksnCrdInfo.width = frmMain.fraWksnSessionInfo.width
    frmMain.fraWksnCliInfo.width = frmMain.fraWksnSessionInfo.width
    frmMain.fraWksnSelfserveInfo.width = frmMain.fraWksnSessionInfo.width
    frmMain.fraWksnSelfservePrepaidInfo.width = frmMain.fraWksnSessionInfo.width
    frmMain.fraWksnChat.width = frmMain.fraWksnSessionInfo.width
    frmMain.fraWksnChat.height = frmMain.fraWksnSessionInfo.height - frmMain.fraWksnChat.top - 60
    frmMain.fraWksnChatEntry.width = frmMain.fraWksnChat.width
    frmMain.fraWksnChatEntry.top = frmMain.fraWksnChat.height - frmMain.fraWksnChatEntry.height
    frmMain.txtWksnChatHist.width = frmMain.lblWksnSel(1).width
    frmMain.txtWksnChatHist.height = frmMain.fraWksnChat.height - frmMain.fraWksnChatEntry.height - frmMain.txtWksnChatHist.top - 60
    frmMain.txtWksnChatHist.SelStart = Len(frmMain.txtWksnChatHist.Text)
    frmMain.txtWksnChat.width = frmMain.fraWksnChat.width - frmMain.txtWksnChat.left - frmMain.cmdWksnButtons(13).width - 120
    frmMain.cmdWksnButtons(13).left = frmMain.txtWksnChatHist.width + frmMain.txtWksnChatHist.left - frmMain.cmdWksnButtons(13).width 'frmMain.lblWksnSel(1).width + frmMain.lbl(55).width - frmMain.cmdWksnButtons(14).width + 100
    frmMain.shpWksnChatHist.width = frmMain.txtWksnChatHist.width - frmMain.cmdWksnButtons(13).width - 60
    frmMain.lblConsoleSel(0).width = frmMain.fraWksnControls.width
    frmMain.fraWksnConsoleButtons.top = frmMain.shpWksnConsoleControlsBorder.height - frmMain.fraWksnConsoleButtons.height - 60
    frmMain.fraWksnConsoleButtons.left = frmMain.shpWksnConsoleControlsBorder.left + 60
    frmMain.fraWksnConsoleButtons.width = frmMain.shpWksnConsoleControlsBorder.width - 120
End Sub


'*******************************************************************************
Public Sub subResizeCards()
On Error GoTo subResizeCardsError

    frmMain.fraCrdPending.top = frmMain.fraFunctions(2).height - frmMain.fraCrdPending.height
    If gblnDBOK Then
        frmMain.fraCrdPending.Visible = Not grstCardsPending.BOF Or Not grstCardsPending.EOF
        frmMain.shpCrdListBorder.height = -IIf(grstCardsPending.BOF And grstCardsPending.EOF, 0, frmMain.fraCrdPending.height + 60) + frmMain.fraFunctions(2).height - frmMain.shpCrdListBorder.top - frmMain.fraCrdListTotal.height + Screen.TwipsPerPixelY
    Else
        frmMain.fraCrdPending.Visible = False
        frmMain.shpCrdListBorder.height = frmMain.fraFunctions(2).height - frmMain.shpCrdListBorder.top - frmMain.fraCrdListTotal.height + Screen.TwipsPerPixelY
    End If
    frmMain.fxgCards.height = frmMain.shpCrdListBorder.height - 2 * Screen.TwipsPerPixelY
    
    frmMain.fraCrdListTotal.top = frmMain.shpCrdListBorder.top + frmMain.shpCrdListBorder.height - Screen.TwipsPerPixelY ' frmMain.fraFunctions(2).height - frmMain.fraCrdListTotal.height '
    
    frmMain.shpCrdControlsBorder.left = frmMain.shpCrdListBorder.width + 60
    frmMain.shpCrdControlsBorder.top = 0
    frmMain.shpCrdControlsBorder.height = frmMain.fraFunctions(2).height
    frmMain.shpCrdControlsBorder.width = frmMain.fraFunctions(2).width - frmMain.shpCrdControlsBorder.left
    frmMain.lblCrdSel(0).left = frmMain.shpCrdControlsBorder.left
    frmMain.lblCrdSel(0).width = frmMain.shpCrdControlsBorder.width
    frmMain.fraCrdControls.top = frmMain.shpCrdControlsBorder.height - frmMain.fraCrdControls.height - 60
    frmMain.fraCrdControls.left = frmMain.shpCrdControlsBorder.left + 60
    frmMain.fraCrdControls.width = frmMain.shpCrdControlsBorder.width - 120
    frmMain.lblCrdSel(1).width = frmMain.lblCrdSel(0).width - frmMain.lbl(7).width - 3 * 60
    frmMain.lblCrdSel(2).width = frmMain.lblCrdSel(1).width
    frmMain.lblCrdSel(3).width = frmMain.lblCrdSel(1).width
    frmMain.lblCrdSel(4).width = frmMain.lblCrdSel(1).width - frmMain.lblCrdSel(5).width + Screen.TwipsPerPixelX
    frmMain.cmbCrdType.width = frmMain.lblCrdSel(1).width
    frmMain.shpCrdComment.width = frmMain.lblCrdSel(1).width
    frmMain.txtCrdComment.width = frmMain.lblCrdSel(1).width - 120
    guServOpts.intNumVisRows = (frmMain.fxgCards.height - 255) / 240 - 1
    Exit Sub
    
subResizeCardsError:
    Call funErrorManager("modFunctions.subResizeCardsError")
End Sub


'*******************************************************************************
Public Sub subResizeClients()
On Error GoTo subResizeClientsError

    frmMain.shpCliListBorder.height = frmMain.fraFunctions(3).height - frmMain.shpCliListBorder.top - frmMain.fracliListTotal.height + Screen.TwipsPerPixelY
    frmMain.fxgClients.height = frmMain.shpCliListBorder.height - 2 * Screen.TwipsPerPixelY
    frmMain.fracliListTotal.top = frmMain.fraFunctions(3).height - frmMain.fracliListTotal.height '- Screen.TwipsPerPixelY
    frmMain.shpCliControlsBorder.left = frmMain.shpCliListBorder.width + 60
    frmMain.shpCliControlsBorder.top = 0
    frmMain.shpCliControlsBorder.height = frmMain.fraFunctions(3).height
    frmMain.shpCliControlsBorder.width = frmMain.fraFunctions(3).width - frmMain.shpCliControlsBorder.left
    frmMain.lblCliSel(0).left = frmMain.shpCliControlsBorder.left
    frmMain.lblCliSel(0).width = frmMain.shpCliControlsBorder.width
    frmMain.fraCliControls.top = frmMain.shpCliControlsBorder.height - frmMain.fraCliControls.height - 60
    frmMain.fraCliControls.left = frmMain.shpCliControlsBorder.left + 60
    frmMain.fraCliControls.width = frmMain.shpCliControlsBorder.width - 120
    
    frmMain.lblCliSel(1).width = frmMain.lblCliSel(0).width - frmMain.lbl(42).width - 3 * 60
    frmMain.lblCliSel(2).width = frmMain.lblCliSel(1).width
    frmMain.lblCliSel(3).width = frmMain.lblCliSel(1).width
    frmMain.lblCliSel(4).width = frmMain.lblCliSel(1).width - frmMain.lblCliSel(5).width + Screen.TwipsPerPixelX
    frmMain.lblCliSel(6).width = frmMain.lblCliSel(1).width
    frmMain.lblCliSel(7).width = frmMain.lblCliSel(1).width - frmMain.lblCliSel(8).width + Screen.TwipsPerPixelX
'    frmMain.cmbCrdType.width = frmMain.lblCrdSel(1).width
'    frmMain.shpCrdComment.width = frmMain.lblCrdSel(1).width
'    frmMain.txtCrdComment.width = frmMain.lblCrdSel(1).width - 120
    
'    guServOpts.intNumVisRows = (frmMain.fxgCards.height - 255) / 240 - 1
    Exit Sub
    
subResizeClientsError:
    Call funErrorManager("modFunctions.subResizeClientsError")
End Sub

'*******************************************************************************
Public Sub subResizeIncome()
On Error GoTo subResizeIncomeError

    frmMain.fraIncomeControls.top = frmMain.fraFunctions(5).height - frmMain.fraIncomeControls.height - 60
    frmMain.shpIncomeFilters.height = frmMain.fraFunctions(5).height
    frmMain.shpIncomeListBorder.left = frmMain.shpIncomeFilters.width + 60
    frmMain.shpIncomeListBorder.width = frmMain.fraFunctions(5).width - frmMain.shpIncomeFilters.width - 60
    frmMain.shpIncomeListBorder.height = frmMain.fraFunctions(5).height
    frmMain.fxgIncome.top = frmMain.shpIncomeListBorder.top + Screen.TwipsPerPixelY
    frmMain.fxgIncome.left = frmMain.shpIncomeListBorder.left + Screen.TwipsPerPixelX
    frmMain.fxgIncome.height = frmMain.shpIncomeListBorder.height - 2 * Screen.TwipsPerPixelY
    frmMain.fxgIncome.width = frmMain.shpIncomeListBorder.width - 2 * Screen.TwipsPerPixelX
    Exit Sub
    
subResizeIncomeError:
    Call funErrorManager("modFunctions.subResizeIncomeError")
End Sub


'*******************************************************************************
Public Sub subResizeSession()
On Error GoTo subResizeSessionError

    frmMain.fraSessionControls.top = frmMain.fraFunctions(4).height - frmMain.fraSessionControls.height - 60
    frmMain.shpSessionControlsBorder.height = frmMain.fraFunctions(4).height
    
    frmMain.tbsSessionResultType.left = frmMain.shpSessionControlsBorder.width + 60
    frmMain.fraSessionResults(3).top = frmMain.tbsSessionResultType.height '+ 60
    frmMain.fraSessionResults(3).left = frmMain.shpSessionControlsBorder.width + 60
    frmMain.fraSessionResults(3).width = frmMain.fraFunctions(4).width - frmMain.fraSessionResults(3).left
    frmMain.fraSessionResults(3).height = frmMain.fraFunctions(4).height - frmMain.fraSessionResults(3).top
    frmMain.shpSessionResultsListBorder.width = frmMain.fraSessionResults(3).width
    frmMain.shpSessionResultsListBorder.height = frmMain.fraSessionResults(3).height
    frmMain.fraSessionResults(1).top = frmMain.fraSessionResults(3).top
    frmMain.fraSessionResults(1).left = frmMain.fraSessionResults(3).left
    frmMain.fraSessionResults(1).width = frmMain.fraSessionResults(3).width
    frmMain.fraSessionResults(1).height = frmMain.fraSessionResults(3).height
    frmMain.shpSessionResultsStatsBorder.width = frmMain.fraSessionResults(1).width
    frmMain.shpSessionResultsStatsBorder.height = frmMain.fraSessionResults(1).height
    frmMain.fraSessionResults(2).top = frmMain.fraSessionResults(3).top
    frmMain.fraSessionResults(2).left = frmMain.fraSessionResults(3).left
    frmMain.fraSessionResults(2).width = frmMain.fraSessionResults(3).width
    frmMain.fraSessionResults(2).height = frmMain.fraSessionResults(3).height
    frmMain.shpSessionResultsTimesBorder.width = frmMain.fraSessionResults(1).width
    frmMain.shpSessionResultsTimesBorder.height = frmMain.fraSessionResults(1).height
    
    frmMain.fraSessionResults(4).top = frmMain.fraSessionResults(3).top
    frmMain.fraSessionResults(4).left = frmMain.fraSessionResults(3).left
    frmMain.fraSessionResults(4).width = frmMain.fraSessionResults(3).width
    frmMain.fraSessionResults(4).height = frmMain.fraSessionResults(3).height
    frmMain.fraSessionResultsURLSearch.top = frmMain.fraSessionResults(3).height - frmMain.fraSessionResultsURLSearch.height - 60
    frmMain.shpSessionResultsURLsBorder.width = frmMain.fraSessionResults(1).width
    frmMain.shpSessionResultsURLsBorder.height = frmMain.fraSessionResults(1).height
    frmMain.linSessionResultsURLs.X1 = 0
    frmMain.linSessionResultsURLs.X2 = frmMain.fraSessionResults(4).width
    frmMain.linSessionResultsURLs.Y1 = frmMain.fraSessionResults(4).height - frmMain.fraSessionResultsURLSearch.height - 2 * 60
    frmMain.linSessionResultsURLs.Y2 = frmMain.linSessionResultsURLs.Y1
    
    frmMain.fraSessionResults(5).top = frmMain.fraSessionResults(3).top
    frmMain.fraSessionResults(5).left = frmMain.fraSessionResults(3).left
    frmMain.fraSessionResults(5).width = frmMain.fraSessionResults(3).width
    frmMain.fraSessionResults(5).height = frmMain.fraSessionResults(3).height
    frmMain.shpSessionResultsLinesBorder.width = frmMain.fraSessionResults(1).width
    frmMain.shpSessionResultsLinesBorder.height = frmMain.fraSessionResults(1).height
    frmMain.fxgSession(2).top = Screen.TwipsPerPixelY
    frmMain.fxgSession(2).left = Screen.TwipsPerPixelX
    frmMain.fxgSession(2).width = frmMain.fraSessionResults(3).width - 2 * Screen.TwipsPerPixelX
    frmMain.fxgSession(2).height = frmMain.fraSessionResults(3).height - 2 * Screen.TwipsPerPixelY
    frmMain.fxgSession(1).top = frmMain.fxgSession(2).top
    frmMain.fxgSession(1).left = frmMain.fxgSession(2).left
    frmMain.fxgSession(1).width = frmMain.fxgSession(2).width
    frmMain.fxgSession(1).height = frmMain.fxgSession(2).height
    frmMain.fxgSession(3).top = frmMain.fxgSession(2).top
    frmMain.fxgSession(3).left = frmMain.fxgSession(2).left
    frmMain.fxgSession(3).width = frmMain.fxgSession(2).width
    frmMain.fxgSession(3).height = frmMain.fxgSession(2).height
    frmMain.fxgSession(5).top = frmMain.fxgSession(2).top
    frmMain.fxgSession(5).left = frmMain.fxgSession(2).left
    frmMain.fxgSession(5).width = frmMain.fxgSession(2).width
    frmMain.fxgSession(5).height = frmMain.fxgSession(2).height
    frmMain.shpSessionLines.top = frmMain.fxgSession(2).top
    frmMain.shpSessionLines.left = frmMain.fxgSession(2).left
    frmMain.shpSessionLines.width = frmMain.fxgSession(2).width
    frmMain.shpSessionLines.height = frmMain.fxgSession(2).height
    
    frmMain.fxgSession(4).top = frmMain.fxgSession(2).top
    frmMain.fxgSession(4).left = frmMain.fxgSession(2).left
    frmMain.fxgSession(4).width = frmMain.fxgSession(2).width
    frmMain.fxgSession(4).height = frmMain.linSessionResultsURLs.Y1 - Screen.TwipsPerPixelY
    
    frmMain.tbsSessionResultType.width = frmMain.fraSessionResults(3).width
    Exit Sub
    
subResizeSessionError:
    Call funErrorManager("modFunctions.subResizeSessionError")
End Sub


'*******************************************************************************
Public Sub subResizeServices()
On Error GoTo subResizeServicesError

    frmMain.shpServicesListBorder.height = frmMain.fraFunctions(1).height
    frmMain.shpServicesControlsBorder.left = frmMain.shpServicesListBorder.width + 60
    frmMain.shpServicesControlsBorder.height = frmMain.fraFunctions(1).height
    frmMain.shpServicesControlsBorder.width = frmMain.fraFunctions(1).width - frmMain.shpServicesControlsBorder.left
    frmMain.lblServSel(0).left = frmMain.shpServicesControlsBorder.left
    frmMain.lblServSel(0).width = frmMain.shpServicesControlsBorder.width
    frmMain.lblServSel(1).width = frmMain.shpServicesControlsBorder.width - 2000
    frmMain.lblServSel(2).width = frmMain.lblServSel(1).width
    frmMain.lblServSel(3).width = frmMain.lblServSel(1).width
    frmMain.fraServicesControls.left = frmMain.shpServicesControlsBorder.left + 1 * 60
    frmMain.fraServicesControls.top = frmMain.shpServicesControlsBorder.height - frmMain.fraServicesControls.height - 1 * 60
    
    frmMain.lblServicesTitle.width = frmMain.shpServicesListBorder.width
    frmMain.fxgServices.top = frmMain.shpServicesListBorder.top + Screen.TwipsPerPixelY + frmMain.lblServicesTitle.height
    frmMain.fxgServices.left = frmMain.shpServicesListBorder.left + Screen.TwipsPerPixelX
    frmMain.fxgServices.height = frmMain.shpServicesListBorder.height - 2 * Screen.TwipsPerPixelY - frmMain.lblServicesTitle.height
    frmMain.fxgServices.width = frmMain.shpServicesListBorder.width - 2 * Screen.TwipsPerPixelX
    
    Exit Sub
    
subResizeServicesError:
    Call funErrorManager("modFunctions.subResizeServicesError")
End Sub


'*******************************************************************************
Public Sub subUpdateUISelectedCard()
Dim uSelectedCard As New clsCard
Dim i As Integer
Dim blnCardTypeExists As Boolean
Dim intWorkstationIndex As Integer
On Error GoTo subUpdateUISelectedCardError

    If guServOpts.lngSelectedCardIndex > 0 Then
        With frmMain
            Set uSelectedCard = guCards(guServOpts.lngSelectedCardIndex)
            
            .lblCrdSel(0).Caption = funLoadResString(10198) & uSelectedCard.ID
            .lblCrdSel(1).Caption = " " & IIf(uSelectedCard.UnlimitedPlay, funLoadResString(11000), funFormatTime(uSelectedCard.TimeCredit, enuTimeFormat.Hh_Nmn_Ssec_dyn))
            .lblCrdSel(2).Caption = " " & format(uSelectedCard.DateCreation, "Long Date")
            .lblCrdSel(3).Caption = " " & format(uSelectedCard.DateExpiry, "Long Date")
            
            If uSelectedCard.CardType = "" Then
                For i = 0 To frmMain.cmbCrdType.ListCount - 1
                    If .cmbCrdType.ItemData(i) = 1 Then
                        .cmbCrdType.ListIndex = i
                        Exit For
                    End If
                Next i
            Else
                For i = 0 To frmMain.cmbCrdType.ListCount - 1
                    If LCase(frmMain.cmbCrdType.List(i)) = LCase(uSelectedCard.CardType) Then
                        blnCardTypeExists = True
                        Exit For
                    End If
                Next i
                If Not blnCardTypeExists Then Call frmMain.cmbCrdType.AddItem(uSelectedCard.CardType)
                For i = 0 To .cmbCrdType.ListCount - 1
                    If LCase(.cmbCrdType.List(i)) = LCase(uSelectedCard.CardType) Then
                        .cmbCrdType.ListIndex = i
                        Exit For
                    End If
                Next i
            End If
            
'            .cmdCrdButtons(1).enabled = Not uSelectedCard.UnlimitedPlay
'            .mnuCrd(1).enabled = Not uSelectedCard.UnlimitedPlay
'            .mnuCrdDD(1).enabled = Not uSelectedCard.UnlimitedPlay

            .cmbCrdType.Locked = False
            .cmbCrdType.enabled = True
            .txtCrdComment.Text = uSelectedCard.Comment
            .txtCrdComment.enabled = True
            .txtCrdComment.Locked = False
            .txtCrdComment.Visible = True
            If uSelectedCard.Active Then
                intWorkstationIndex = funCardInUse(uSelectedCard, True)
                If intWorkstationIndex <> 0 Then
                    .lblCrdSel(4).Caption = "" & funLoadResString(10199) & " -" & funLoadResString(10686) & guWorkstations(intWorkstationIndex).WorkstationName
                Else
                    .lblCrdSel(4).Caption = "" & funLoadResString(10199)
                End If
                .cmdCrdButtons(4).Visible = True
                .cmdCrdButtons(3).Visible = False
                .mnuCrd(4).enabled = True
                .mnuCrdDD(4).Visible = True
                .mnuCrd(3).enabled = False
                .mnuCrdDD(3).Visible = False
                .imgCrdStateActive.Visible = True
                .imgCrdStateInactive.Visible = False
            Else
                .lblCrdSel(4).Caption = "" & funLoadResString(10200)
                .cmdCrdButtons(4).Visible = False
                .cmdCrdButtons(3).Visible = True
                .mnuCrd(4).enabled = False
                .mnuCrdDD(4).Visible = False
                .mnuCrd(3).enabled = True
                .mnuCrdDD(3).Visible = True
                .imgCrdStateActive.Visible = False
                .imgCrdStateInactive.Visible = True
            End If
            .tmrCrdRefresh.enabled = True
        End With
    End If
    Exit Sub
    
subUpdateUISelectedCardError:
    Call funErrorManager("modFunctions.subUpdateUISelectedCardError")
End Sub


'*******************************************************************************
Public Sub subUpdateUISelectedClient()
Dim uSelectedClient As New clsClient
Dim i As Integer
Dim blnClientTypeExists As Boolean
Dim intWorkstationIndex As Integer
On Error GoTo subUpdateUISelectedClientError

    If guServOpts.lngSelectedClientIndex > 0 Then
        With frmMain
            Set uSelectedClient = guClients(guServOpts.lngSelectedClientIndex)
            
            .lblCliSel(0).Caption = funLoadResString(10800) & uSelectedClient.login
            .lblCliSel(1).Caption = " " & IIf(uSelectedClient.UnlimitedPlay, funLoadResString(11000), funFormatTime(uSelectedClient.TimeCredit, enuTimeFormat.Hh_Nmn_Ssec_dyn))
            .lblCliSel(2).Caption = " " & format(uSelectedClient.DateCreation, "Long Date")
            If guServOpts.uClients.intEnableExpiry = vbChecked Then .lblCliSel(3).Caption = " " & format(uSelectedClient.DateExpiry, "Long Date")
            .lblCliSel(6).Caption = " " & uSelectedClient.FullName
            .lblCliSel(7).Caption = funLeftNewLines(uSelectedClient.Comment, 12)
            
'            .cmdCliButtons(1).enabled = Not uSelectedClient.UnlimitedPlay
'            .mnuCli(1).enabled = Not uSelectedClient.UnlimitedPlay
'            .mnuCliDD(1).enabled = Not uSelectedClient.UnlimitedPlay
            
            If uSelectedClient.Active Then
                intWorkstationIndex = funClientInUse(uSelectedClient, True)
                If intWorkstationIndex <> 0 Then
                    .lblCliSel(4).Caption = "" & funLoadResString(10199) & " -" & funLoadResString(10686) & guWorkstations(intWorkstationIndex).WorkstationName
                Else
                    .lblCliSel(4).Caption = "" & funLoadResString(10199)
                End If
                .cmdCliButtons(4).Visible = True
                .cmdCliButtons(3).Visible = False
                .mnuCli(4).enabled = True
                .mnuCliDD(4).Visible = True
                .mnuCli(3).enabled = False
                .mnuCliDD(3).Visible = False
                .imgCliEnabled.Visible = True
                .imgCliDisabled.Visible = False
            Else
                .lblCliSel(4).Caption = "" & funLoadResString(10200)
                .cmdCliButtons(4).Visible = False
                .cmdCliButtons(3).Visible = True
                .mnuCli(4).enabled = False
                .mnuCliDD(4).Visible = False
                .mnuCli(3).enabled = True
                .mnuCliDD(3).Visible = True
                .imgCliEnabled.Visible = False
                .imgCliDisabled.Visible = True
            End If
            .tmrCliRefresh.enabled = True
        End With
    End If
    Exit Sub
    
subUpdateUISelectedClientError:
    Call funErrorManager("modFunctions.subUpdateUISelectedClientError")
End Sub


'*******************************************************************************
'Public Sub subWorkstationProcessData(ByVal strData As String, ByVal strKey As String)
Public Sub subWorkstationProcessData(uWorkstation As clsWorkstation)
'Dim uWorkstation As New clsWorkstation
'Dim uCCard As New clsCard
'Dim strDataBuffer As String
Dim strDatagram As String
Dim strPayload As String
Dim strTokenizedPayload() As String
Dim strHeader As String
Dim lngEndIndex As Long
Dim intNewState As Integer
Dim intOldState As Integer
Dim strOldComputerName As String
Dim uMsg As New clsStatusBarMessage
On Error GoTo subWorkstationProcessDataError

'    Set uWorkstation = guWorkstations(strKey)
'    strDataBuffer = uWorkstation.CnxIncomingBuffer
    
    Do Until uWorkstation.CnxIncomingBuffer = ""
        lngEndIndex = InStr(uWorkstation.CnxIncomingBuffer, CNX_MSG_END)
        If lngEndIndex = 0 Then Exit Do
        strDatagram = left(uWorkstation.CnxIncomingBuffer, lngEndIndex - 1)
        uWorkstation.CnxIncomingBuffer = Mid(uWorkstation.CnxIncomingBuffer, lngEndIndex + 20)
        strHeader = left(strDatagram, 20)
        strPayload = Mid(strDatagram, 21)
        
        Select Case strHeader
        Case CNX_APP_VERSION
            If CLng(strPayload) < WORKSTATION_MIN_VERSION Then
                Call subWorkstationDisable(uWorkstation, VersionNOK)
                Exit Do
            End If
        
        Case CNX_WORKSTATION_NAME
            uWorkstation.WorkstationName = strPayload
        
        Case CNX_COMPUTER_NAME
            strOldComputerName = uWorkstation.ComputerName
            strTokenizedPayload = funTokenize(strPayload, CNX_PAYLOAD_TOKEN)
            uWorkstation.ComputerName = CStr(strTokenizedPayload(0))
            uWorkstation.UserName = CStr(strTokenizedPayload(1))
            uWorkstation.AllowLockdown = CInt(strTokenizedPayload(2))
            uWorkstation.ComputerMacAddress = CStr(strTokenizedPayload(3))
            If uWorkstation.ComputerName <> strOldComputerName Then
                Call subWorkstationLockdown(uWorkstation)
            End If
            Call subUpdateWksnList(uWorkstation.ComputerMacAddress, uWorkstation.WorkstationName, Now)
            
        Case CNX_WORKSTATION_STATE
            intOldState = uWorkstation.State
            intNewState = CInt(strPayload)
            If uWorkstation.State <> intNewState Then uWorkstation.OldState = uWorkstation.State
            uWorkstation.State = intNewState
            Select Case uWorkstation.State
            Case enuWksnState.NotActive
                If intOldState <> enuWksnState.NotActive Then Call subDBAddEvent(WorkstationNotActive, 0, 0, "", 0, funLoadResString(10563) & uWorkstation.WorkstationName, All)
                
            Case enuWksnState.Ready
                If intOldState <> enuWksnState.Ready Then Call subDBAddEvent(WorkstationReady, 0, 0, "", 0, funLoadResString(10564) & uWorkstation.WorkstationName, All)
                
            Case enuWksnState.Blocked_Deactivated
                If intOldState <> enuWksnState.Blocked_Deactivated Then Call subDBAddEvent(WorkstationDeactivated, 0, 0, "", 0, funLoadResString(10565) & uWorkstation.WorkstationName, All)
            
            Case enuWksnState.Blocked_SessionEnd
            
            Case enuWksnState.UsedConnected
                If intOldState <> enuWksnState.UsedConnected And intOldState <> enuWksnState.Blocked_Paused And intOldState <> enuWksnState.Blocked_Locked And intOldState <> enuWksnState.Blocked_follow Then Call subDBAddEvent(SessionConnected, uWorkstation.SessionID, uWorkstation.CurrentCard.ID, uWorkstation.CurrentClient.login, 0, funLoadResString(10570) & uWorkstation.WorkstationName, Verbose)
                If intOldState = enuWksnState.Blocked_Paused Or intOldState = enuWksnState.Blocked_Locked Or intOldState = enuWksnState.Blocked_follow Then Call subDBAddEvent(SessionUnPaused, uWorkstation.SessionID, uWorkstation.CurrentCard.ID, uWorkstation.CurrentClient.login, 0, funLoadResString(10566) & uWorkstation.WorkstationName, All)
            
            Case enuWksnState.UsedFree
                If intOldState <> enuWksnState.UsedFree And intOldState <> enuWksnState.Blocked_Paused And intOldState <> enuWksnState.Blocked_Locked And intOldState <> enuWksnState.Blocked_follow Then Call subDBAddEvent(SessionFree, uWorkstation.SessionID, uWorkstation.CurrentCard.ID, uWorkstation.CurrentClient.login, 0, funLoadResString(10625) & uWorkstation.WorkstationName, Verbose)
                If intOldState = enuWksnState.Blocked_Paused Or intOldState = enuWksnState.Blocked_Locked Or intOldState = enuWksnState.Blocked_follow Then Call subDBAddEvent(SessionUnPaused, uWorkstation.SessionID, uWorkstation.CurrentCard.ID, uWorkstation.CurrentClient.login, 0, funLoadResString(10566) & uWorkstation.WorkstationName, All)
            
            Case enuWksnState.UsedDisconnected
                If intOldState <> enuWksnState.UsedDisconnected And intOldState <> enuWksnState.Blocked_Paused And intOldState <> enuWksnState.Blocked_Locked And intOldState <> enuWksnState.Blocked_follow Then Call subDBAddEvent(SessionDisconnected, uWorkstation.SessionID, uWorkstation.CurrentCard.ID, uWorkstation.CurrentClient.login, 0, funLoadResString(10571) & uWorkstation.WorkstationName, Verbose)
                If intOldState = enuWksnState.Blocked_Paused Or intOldState = enuWksnState.Blocked_Locked Or intOldState = enuWksnState.Blocked_follow Then Call subDBAddEvent(SessionUnPaused, uWorkstation.SessionID, uWorkstation.CurrentCard.ID, uWorkstation.CurrentClient.login, 0, funLoadResString(10566) & uWorkstation.WorkstationName, All)
            
            Case enuWksnState.Blocked_NoCredit
                If intOldState <> enuWksnState.Blocked_NoCredit Then Call subDBAddEvent(SessionNoCredit, uWorkstation.SessionID, uWorkstation.CurrentCard.ID, "", 0, funLoadResString(10567) & uWorkstation.WorkstationName, All)
                
            Case enuWksnState.Blocked_Paused, enuWksnState.Blocked_Locked, enuWksnState.Blocked_follow
                If intOldState <> enuWksnState.Blocked_Paused And intOldState <> enuWksnState.Blocked_Locked And intOldState <> enuWksnState.Blocked_follow Then Call subDBAddEvent(SessionPaused, uWorkstation.SessionID, uWorkstation.CurrentCard.ID, uWorkstation.CurrentClient.login, 0, funLoadResString(10568) & uWorkstation.WorkstationName, All)
                
            End Select
        
        Case CNX_SESSION_ID
'            uWorkstation.SessionID = CLng(strPayload)
        
        Case CNX_SESSION_TYPE
            uWorkstation.SessionType = CInt(strPayload)
        
        Case CNX_USER_ID
            Select Case uWorkstation.SessionType
            Case enuSessionType.Card
                If IsNumeric(strPayload) Then uWorkstation.CurrentCard.ID = CLng(strPayload)
            Case enuSessionType.Client
                uWorkstation.CurrentClient.login = CStr(strPayload)
            Case Else
                uWorkstation.CurrentCard.ID = 0
                uWorkstation.CurrentClient.login = ""
            End Select
        
        Case CNX_USER_BALANCE
            uWorkstation.UserBalance = funSinglefromNet(strPayload)
        
        Case CNX_ELAPSED_TIME
            uWorkstation.ElapsedTime = funDateTimefromNet(strPayload)
            
        Case CNX_CONNECTED_TIME
            uWorkstation.ConnectedTime = funDateTimefromNet(strPayload)
            
        Case CNX_CONNECTED_TIME_OFFPEAK
            uWorkstation.ConnectedTimeOffPeak = funDateTimefromNet(strPayload)
            
        Case CNX_DISCONNECTED_TIME
            uWorkstation.DisconnectedTime = funDateTimefromNet(strPayload)
            
        Case CNX_REMAINING_TIME
            If uWorkstation.State <> enuWksnState.Blocked_NoCredit Then
                Select Case uWorkstation.SessionType
                Case enuSessionType.Card
                    uWorkstation.CurrentCard.TimeCredit = funDateTimefromNet(strPayload)
                    If uWorkstation.CurrentCard.TimeCredit < NULL_DATE Then uWorkstation.CurrentCard.TimeCredit = NULL_DATE
                Case enuSessionType.Client
                    uWorkstation.CurrentClient.TimeCredit = funDateTimefromNet(strPayload)
                    If uWorkstation.CurrentClient.TimeCredit < NULL_DATE Then uWorkstation.CurrentClient.TimeCredit = NULL_DATE
                Case enuSessionType.SelfServicePrePaid
                    uWorkstation.SelfServicePrePaid.TimeCredit = funDateTimefromNet(strPayload)
                    If uWorkstation.SelfServicePrePaid.TimeCredit < NULL_DATE Then uWorkstation.SelfServicePrePaid.TimeCredit = NULL_DATE
                End Select
            End If
            
        Case CNX_REMAINING_TIME_APARENT
            If uWorkstation.State <> enuWksnState.Blocked_NoCredit Then
                uWorkstation.RemainingTimeAparent = funDateTimefromNet(strPayload)
                If uWorkstation.RemainingTimeAparent < NULL_DATE Then uWorkstation.RemainingTimeAparent = NULL_DATE
            End If
            
        Case CNX_SESSION_OPEN
            strTokenizedPayload = funTokenize(strPayload, CNX_PAYLOAD_TOKEN)
            uWorkstation.SessionType = CInt(strTokenizedPayload(0))
            Select Case uWorkstation.SessionType
            Case enuSessionType.Card
                uWorkstation.CurrentCard.ID = CLng(strTokenizedPayload(2))
            Case enuSessionType.Client
                uWorkstation.CurrentClient.login = CStr(strTokenizedPayload(2))
                uWorkstation.CurrentClient.Password = CStr(strTokenizedPayload(1))
            End Select
            uWorkstation.State = CInt(strTokenizedPayload(3))
            uWorkstation.StartTime = Now
            uWorkstation.AuthenticationInProgress = True
            Call subDBAddEvent(WorkstationUsageAttempt, 0, uWorkstation.CurrentCard.ID, uWorkstation.CurrentClient.login, 0, funLoadResString(10656) & uWorkstation.WorkstationName, Verbose)
            Call subSessionOpenRequest(uWorkstation)
        
        Case CNX_FORCE_UPDATE
            strTokenizedPayload = funTokenize(strPayload, CNX_PAYLOAD_TOKEN)
            uWorkstation.RemainingTimeAparent = funDateTimefromNet(strTokenizedPayload(1))
            Select Case uWorkstation.SessionType
            Case enuSessionType.Card
                uWorkstation.CurrentCard.TimeCredit = funDateTimefromNet(strTokenizedPayload(0))
                Call subForceCardUpdate(uWorkstation)
            Case enuSessionType.Client
                uWorkstation.CurrentClient.TimeCredit = funDateTimefromNet(strTokenizedPayload(0))
                Call subForceClientUpdate(uWorkstation)
            Case enuSessionType.SelfServicePrePaid
                uWorkstation.SelfServicePrePaid.TimeCredit = funDateTimefromNet(strTokenizedPayload(0))
            End Select
            
        Case CNX_CHECK_BALANCE
            Select Case uWorkstation.SessionType
            Case enuSessionType.Card
                uWorkstation.CurrentCard.ID = CLng(strPayload)
                Call funCheckCardBalance(uWorkstation)
            Case enuSessionType.Client
                uWorkstation.CurrentClient.login = CStr(strPayload)
                Call funCheckClientBalance(uWorkstation)
            Case enuSessionType.SelfServicePrePaid
                Call funCheckSelfServicePrePaidBalance(uWorkstation)
            End Select
            
        Case CNX_MSG_SEND
            Call funMsg(Info, strPayload)
            Call subDBAddEvent(AppError, 0, 0, "", 0, strPayload, Verbose)
        
        Case CNX_SESSION_RESUME
            strTokenizedPayload = funTokenize(strPayload, SERIALIZED_SESSION_TOKEN)
            Call subSessionResume(uWorkstation, strTokenizedPayload)
        
        Case CNX_SESSION_URL
            strTokenizedPayload = funTokenize(strPayload, CNX_PAYLOAD_TOKEN)
            Call subProcessURL(uWorkstation, strTokenizedPayload)
        
        Case CNX_DB_BACKUP_DTE
            uWorkstation.LastDBBackupDate = funDateTimefromNet(strPayload)
            
        Case CNX_PASSWD_CHANGE
            uWorkstation.CurrentClient.Password = strPayload
            Call subDBUpdateClient(uWorkstation.CurrentClient)
        
        Case CNX_SERVER_DATE
            Call funSendData(uWorkstation.Cnx, CNX_SERVER_DATE & funDateTimeToSend(Now) & CNX_MSG_END)
        
        Case CNX_CHAT_MSG
            Call uWorkstation.ChatHistory.Add(funFormatTime(Time) & " " & funLoadResString(10958) & ": " & strPayload)
            If Not frmMain.mnuFunctions(1).Checked Or (frmMain.mnuFunctions(1).Checked And (uWorkstation.ComputerMacAddress <> guSelectedWorkstation.ComputerMacAddress Or frmMain.fraWksnConsoleControls.Visible)) Then
                uWorkstation.Status = enuStatus.cliInfo
                uMsg.ImageIndex = enuStatus.cliInfo
                uMsg.Text = funLoadResString(10959, uWorkstation.WorkstationName)
                Call funAddStatusBarMsg(uMsg, "WKSN_MSG_" & uWorkstation.ComputerMacAddress)
            End If
        
        Case CNX_READY_FOR_SERVER_SCREEN
            uWorkstation.ReadyForServerScreen = CBool(strPayload)
        
        End Select
    Loop
    Exit Sub
    
subWorkstationProcessDataError:
Dim strErrDesc As String
    strErrDesc = funErrorManager("modFunctions.subWorkstationProcessDataError", , True)
    Call subWriteToLogFile("modFunctions.subWorkstationProcessDataError: Wksn Key: " & uWorkstation.Key & " Name: " & uWorkstation.ComputerName & " Mac: " & uWorkstation.ComputerMacAddress)
    Call subWriteToLogFile("modFunctions.subWorkstationProcessDataError: Context: " & strHeader & " Payload: " & left(strPayload, 50))
    Call subWorkstationDisable(uWorkstation, InternalError)
    Call Err.Raise(CYBERA_ERROR, "modFunctions.subWorkstationProcessDataError", strErrDesc)
End Sub


'*******************************************************************************
Public Sub subFindCard(strSearchString As String, Optional blnIncludeAllFields As Boolean = True)
Dim uTmpCard As New clsCard
Dim blnFound As Boolean
Dim i As Long
On Error GoTo subFindCardError

    If guCards.Count <> 0 Then
        strSearchString = Trim(strSearchString)
        For i = IIf(frmMain.fxgCards.Row <> frmMain.fxgCards.Rows, frmMain.fxgCards.Row + 1, 1) To guCards.Count
            Set uTmpCard = guCards(i)
            If blnIncludeAllFields Then
                blnFound = (Len(strSearchString) > 0 And (InStr(LCase(uTmpCard.Comment), LCase(strSearchString)) <> 0 Or InStr(LCase(uTmpCard.ID), LCase(strSearchString)) <> 0))
            Else
                blnFound = (Len(strSearchString) > 0 And (InStr(LCase(uTmpCard.ID), LCase(strSearchString)) <> 0))
            End If
            If blnFound Then Exit For
        Next i
        If Not blnFound And frmMain.fxgCards.Row <> 0 Then
            For i = 1 To frmMain.fxgCards.Row
                Set uTmpCard = guCards(i)
                If blnIncludeAllFields Then
                    blnFound = (Len(strSearchString) > 0 And (InStr(LCase(uTmpCard.Comment), LCase(strSearchString)) <> 0 Or InStr(LCase(uTmpCard.ID), LCase(strSearchString)) <> 0))
                Else
                    blnFound = (Len(strSearchString) > 0 And (InStr(LCase(uTmpCard.ID), LCase(strSearchString)) <> 0))
                End If
                If blnFound Then Exit For
            Next i
        End If
    End If
    If blnFound Then
        frmMain.fxgCards.Row = uTmpCard.Index
        Call frmMain.fxgCards_RowColChange
    Else
        Call funMsg(enuMsg.CardNotFound)
    End If
    Exit Sub
    
subFindCardError:
    Call funErrorManager("modFunctions.subFindCardError")
End Sub


'*******************************************************************************
Public Sub subFindClient(strSearchString As String, Optional blnIncludeAllFields As Boolean = True)
Dim uTmpClient As New clsClient
Dim blnFound As Boolean
Dim i As Long
On Error GoTo subFindClientError

    If guClients.Count <> 0 Then
        strSearchString = Trim(strSearchString)
        For i = IIf(frmMain.fxgClients.Row <> frmMain.fxgClients.Rows, frmMain.fxgClients.Row + 1, 1) To guClients.Count
            Set uTmpClient = guClients(i)
            If blnIncludeAllFields Then
                blnFound = (Len(strSearchString) > 0 And (InStr(LCase(uTmpClient.Comment), LCase(strSearchString)) <> 0 Or InStr(LCase(uTmpClient.login), LCase(strSearchString)) <> 0 Or InStr(LCase(uTmpClient.FullName), LCase(strSearchString)) <> 0))
            Else
                blnFound = (Len(strSearchString) > 0 And (InStr(LCase(uTmpClient.login), LCase(strSearchString)) <> 0))
            End If
            If blnFound Then Exit For
        Next i
        If Not blnFound And frmMain.fxgClients.Row <> 0 Then
            For i = 1 To frmMain.fxgClients.Row
                Set uTmpClient = guClients(i)
                If blnIncludeAllFields Then
                    blnFound = (Len(strSearchString) > 0 And (InStr(LCase(uTmpClient.Comment), LCase(strSearchString)) <> 0 Or InStr(LCase(uTmpClient.login), LCase(strSearchString)) <> 0 Or InStr(LCase(uTmpClient.FullName), LCase(strSearchString)) <> 0))
                Else
                    blnFound = (Len(strSearchString) > 0 And (InStr(LCase(uTmpClient.login), LCase(strSearchString)) <> 0))
                End If
                If blnFound Then Exit For
            Next i
        End If
    End If
    If blnFound Then
        frmMain.fxgClients.Row = uTmpClient.Index
        Call frmMain.fxgClients_RowColChange
    Else
        Call funMsg(enuMsg.ClientNotFound)
    End If
    Exit Sub
    
subFindClientError:
    Call funErrorManager("modFunctions.subFindClientError")
End Sub


'*******************************************************************************
Public Sub subCardDelete(uCard As clsCard)
On Error GoTo subCardDeleteError

    If funMsg(enuMsg.CardDeleteConfirm) = vbYes Then
        grstCards.Index = "PrimaryKey"
        Call grstCards.Seek("=", uCard.ID)
        If grstCards.NoMatch Then
            Call funMsg(enuMsg.CardNotFound)
            Exit Sub
        End If
        grstCards.Delete
        Call subDBAddEvent(CardDelete, 0, uCard.ID, "", 0, funLoadResString(10342), Verbose)
        Call subFillCards
    End If
    Exit Sub

subCardDeleteError:
    Call funErrorManager("modFunctions.subCardDeleteError")
End Sub


'*******************************************************************************
Public Function funCardInUse(ByVal uCard As clsCard, Optional blnIncludeNoCredit As Boolean = False) As Integer
Dim i As Integer
Dim uTmpWorkstation As New clsWorkstation
On Error GoTo funCardInUseError

    funCardInUse = 0
    For i = 1 To guWorkstations.Count
        Set uTmpWorkstation = guWorkstations(i)
        If uTmpWorkstation.CurrentCard.ID = uCard.ID And Not uTmpWorkstation.AuthenticationInProgress And uTmpWorkstation.SessionID <> -1 Then
            If blnIncludeNoCredit Or uTmpWorkstation.State <> enuWksnState.Blocked_NoCredit Then
                funCardInUse = i
                Exit Function
            End If
        End If
    Next i
    Exit Function

funCardInUseError:
    Call funErrorManager("modFunctions.funCardInUseError")
End Function


'*******************************************************************************
Public Function funSelfServicePrePaidRunning(ByVal uSelfServ As clsClient, Optional blnIncludeNoCredit As Boolean = False) As Integer
Dim i As Integer
Dim uTmpWorkstation As New clsWorkstation
On Error GoTo funSelfServicePrePaidRunningError

    funSelfServicePrePaidRunning = 0
    For i = 1 To guWorkstations.Count
        Set uTmpWorkstation = guWorkstations(i)
        If uTmpWorkstation.Key = uSelfServ.login And Not uTmpWorkstation.AuthenticationInProgress And uTmpWorkstation.SessionID <> -1 Then
            If blnIncludeNoCredit Or uTmpWorkstation.State <> enuWksnState.Blocked_NoCredit Then
                funSelfServicePrePaidRunning = i
                Exit Function
            End If
        End If
    Next i
    Exit Function

funSelfServicePrePaidRunningError:
    Call funErrorManager("modFunctions.funSelfServicePrePaidRunningError")
End Function


'*******************************************************************************
Public Function funClientInUse(ByVal uClient As clsClient, Optional blnIncludeNoCredit As Boolean = False) As Integer
Dim i As Integer
Dim uTmpWorkstation As New clsWorkstation
On Error GoTo funClientInUseError

    funClientInUse = 0
    For i = 1 To guWorkstations.Count
        Set uTmpWorkstation = guWorkstations(i)
        If uTmpWorkstation.CurrentClient.login = uClient.login And Not uTmpWorkstation.AuthenticationInProgress And uTmpWorkstation.SessionID <> -1 Then
            If blnIncludeNoCredit Or uTmpWorkstation.State <> enuWksnState.Blocked_NoCredit Then
                funClientInUse = i
                Exit Function
            End If
        End If
    Next i
    Exit Function

funClientInUseError:
    Call funErrorManager("modFunctions.funClientInUseError")
End Function


'*******************************************************************************
Public Sub subSessionPay(uWorkstation As clsWorkstation)
On Error GoTo subSessionPayError

    grstJournIncome.Index = "session_id"
    Call grstJournIncome.Seek("=", uWorkstation.StickySessionID)
    If Not grstJournIncome.NoMatch Then
        Call grstJournIncome.Edit
        grstJournIncome!paid = True
        grstJournIncome!date_stamp = Now
        grstJournIncome!Label = left(funLoadResString(10201), 50)
        Call grstJournIncome.Update
        Call subDBAddEvent(SessionPay, uWorkstation.StickySessionID, 0, "", grstJournIncome!ID, funLoadResString(10339), Verbose)
        uWorkstation.StickySessionID = -1
    End If
    Exit Sub

subSessionPayError:
    Call funErrorManager("modFunctions.subSessionPayError")
End Sub


'*******************************************************************************
Public Sub subSessionCancelPay(uWorkstation As clsWorkstation)
On Error GoTo subSessionCancelPayError

    grstJournIncome.Index = "session_id"
    Call grstJournIncome.Seek("=", uWorkstation.StickySessionID)
    If Not grstJournIncome.NoMatch Then
        Call grstJournIncome.Edit
        grstJournIncome!paid = False
        grstJournIncome!date_stamp = Now
        grstJournIncome!income = 0
        grstJournIncome!Label = left(funLoadResString(11022) & " - " & funLoadResString(10201), 50)
        Call grstJournIncome.Update
        Call subDBAddEvent(SessionPay, uWorkstation.StickySessionID, 0, "", grstJournIncome!ID, funLoadResString(11022) & " - " & funLoadResString(10339), Verbose)
        uWorkstation.StickySessionID = -1
    End If
    Exit Sub

subSessionCancelPayError:
    Call funErrorManager("modFunctions.subSessionCancelPayError")
End Sub


'*******************************************************************************
Public Sub subCardRecharge(uCard As clsCard, ByVal sngPrice As Single, ByVal dteRechargeTime As Date, ByVal strLabel As String, ByVal intWorkstationIndex As Integer)
Dim strBuffer As String
Dim intCliIndex As Integer
On Error GoTo subCardRechargeError

    grstCards.Index = "PrimaryKey"
    Call grstCards.Seek("=", uCard.ID)
    If grstCards.NoMatch Then
        Call funMsg(enuMsg.CardNotFound)
        Exit Sub
    End If
    Call grstJournIncome.AddNew
    grstJournIncome!card_id = uCard.ID
    grstJournIncome!Type = enuIncomeType.Card
    grstJournIncome!income = sngPrice
    grstJournIncome!date_stamp = Now
    grstJournIncome!Time = dteRechargeTime
    grstJournIncome!Label = left(strLabel, 50)
    grstJournIncome!paid = True
    grstJournIncome!login = guServOpts.uConnectedUser.login
    grstJournIncome!unlimited_play = uCard.UnlimitedPlay
    grstCards.Index = "PrimaryKey"
    Call subDBAddEvent(CardRecharge, 0, uCard.ID, "", grstJournIncome!ID, funLoadResString(10340), All)
    Call grstJournIncome.Update
    
    grstCards.Index = "PrimaryKey"
    Call grstCards.Seek("=", uCard.ID)
    Call grstCards.Edit
    grstCards!date_expiry = uCard.DateExpiry
    grstCards!unlimited_play = uCard.UnlimitedPlay
    If intWorkstationIndex = 0 Then
        grstCards!time_credit = uCard.TimeCredit
'        grstCards!Active = uCard.Active
        Call grstCards.Update
'        DoEvents
        intCliIndex = funCardInUse(uCard, True)
        If intCliIndex <> 0 Then
'            Set guWorkstations(intCliIndex).CurrentCard = uCard
            guWorkstations(intCliIndex).CurrentCard.TimeCredit = uCard.TimeCredit
            guWorkstations(intCliIndex).RemainingTimeAparent = IIf(guWorkstations(intCliIndex).OldState = enuWksnState.UsedDisconnected, uCard.TimeCredit / (guServOpts.uTimeRatios.intDisconCon / 100), uCard.TimeCredit)
            frmMain.cmdWksnButtons(7).enabled = True
        End If
    Else
        Call grstCards.Update
        strBuffer = CNX_RECHARGE_CARD & funDateTimeToSend(dteRechargeTime) & CNX_PAYLOAD_TOKEN & funDateTimeToSend(uCard.DateExpiry) & CNX_PAYLOAD_TOKEN & CInt(uCard.UnlimitedPlay) & CNX_MSG_END
        Call funSendData(guWorkstations(intWorkstationIndex).Cnx, strBuffer)
    End If
    Exit Sub

subCardRechargeError:
    Call funErrorManager("modFunctions.subCardRechargeError")
End Sub


'*******************************************************************************
Public Sub subSelfservicePrePaidRecharge(uSelfServ As clsClient, ByVal sngPrice As Single, ByVal dteRechargeTime As Date, ByVal strLabel As String, ByVal intWorkstationIndex As Integer)
Dim strBuffer As String
Dim intCliIndex As Integer
On Error GoTo subSelfservicePrePaidRechargeError

    Call grstJournIncome.AddNew
    grstJournIncome!card_id = 0
    grstJournIncome!Type = enuIncomeType.SelfServicePrePaid
    grstJournIncome!income = sngPrice
    grstJournIncome!date_stamp = Now
    grstJournIncome!Time = dteRechargeTime
    grstJournIncome!Label = left(strLabel, 50)
    grstJournIncome!paid = True
    grstJournIncome!login = guServOpts.uConnectedUser.login
    Call subDBAddEvent(SelfServiceRecharge, 0, 0, "", grstJournIncome!ID, funLoadResString(10340), All)
    Call grstJournIncome.Update
    
    If intWorkstationIndex = 0 Then
        intCliIndex = funSelfServicePrePaidRunning(uSelfServ, True)
        If intCliIndex <> 0 Then
            guWorkstations(intCliIndex).SelfServicePrePaid.TimeCredit = uSelfServ.TimeCredit
            guWorkstations(intCliIndex).RemainingTimeAparent = IIf(guWorkstations(intCliIndex).OldState = enuWksnState.UsedDisconnected, uSelfServ.TimeCredit / (guServOpts.uTimeRatios.intDisconCon / 100), uSelfServ.TimeCredit)
            frmMain.cmdWksnButtons(7).enabled = True
        End If
    Else
        strBuffer = CNX_RECHARGE_CARD & funDateTimeToSend(dteRechargeTime) & CNX_PAYLOAD_TOKEN & CNX_PAYLOAD_TOKEN & CNX_MSG_END
        Call funSendData(guWorkstations(intWorkstationIndex).Cnx, strBuffer)
    End If
    Exit Sub

subSelfservicePrePaidRechargeError:
    Call funErrorManager("modFunctions.subSelfservicePrePaidRechargeError")
End Sub


'*******************************************************************************
Public Sub subClientRecharge(uClient As clsClient, ByVal sngPrice As Single, ByVal dteRechargeTime As Date, ByVal strLabel As String, ByVal intWorkstationIndex As Integer)
Dim strBuffer As String
Dim intCliIndex As Integer
On Error GoTo subClientRechargeError

    grstClients.Index = "PrimaryKey"
    Call grstClients.Seek("=", uClient.login)
    If grstClients.NoMatch Then
        Call funMsg(enuMsg.CardNotFound)
        Exit Sub
    End If
    Call grstJournIncome.AddNew
    grstJournIncome!client_login = uClient.login
    grstJournIncome!Type = enuIncomeType.Client
    grstJournIncome!income = sngPrice
    grstJournIncome!date_stamp = Now
    grstJournIncome!Time = dteRechargeTime
    grstJournIncome!Label = left(strLabel, 50)
    grstJournIncome!paid = True
    grstJournIncome!login = guServOpts.uConnectedUser.login
    grstJournIncome!unlimited_play = uClient.UnlimitedPlay
    grstClients.Index = "PrimaryKey"
    Call subDBAddEvent(ClientRecharge, 0, 0, uClient.login, grstJournIncome!ID, funLoadResString(10340), All)
    Call grstJournIncome.Update
    
    grstClients.Index = "PrimaryKey"
    Call grstClients.Seek("=", uClient.login)
    Call grstClients.Edit
    grstClients!date_expiry = uClient.DateExpiry
    grstClients!unlimited_play = uClient.UnlimitedPlay
    If intWorkstationIndex = 0 Then
        grstClients!time_credit = uClient.TimeCredit
        Call grstClients.Update
        intCliIndex = funClientInUse(uClient, True)
        If intCliIndex <> 0 Then
'            Set guWorkstations(intCliIndex).CurrentCard = uclient
            guWorkstations(intCliIndex).CurrentClient.TimeCredit = uClient.TimeCredit
            guWorkstations(intCliIndex).RemainingTimeAparent = IIf(guWorkstations(intCliIndex).OldState = enuWksnState.UsedDisconnected, uClient.TimeCredit / (guServOpts.uTimeRatios.intDisconCon / 100), uClient.TimeCredit)
            frmMain.cmdWksnButtons(7).enabled = True
        End If
    Else
        Call grstClients.Update
        strBuffer = CNX_RECHARGE_CARD & funDateTimeToSend(dteRechargeTime) & CNX_PAYLOAD_TOKEN & funDateTimeToSend(uClient.DateExpiry) & CNX_PAYLOAD_TOKEN & CInt(uClient.UnlimitedPlay) & CNX_MSG_END
        Call funSendData(guWorkstations(intWorkstationIndex).Cnx, strBuffer)
    End If
    Exit Sub

subClientRechargeError:
    Call funErrorManager("modFunctions.subClientRechargeError")
End Sub


'*******************************************************************************
Public Function funCheckCardBalance(uWorkstation As clsWorkstation) As Boolean
Dim strReply As String
Dim blnReturnVal As Boolean
On Error GoTo funCheckCardBalanceError

    grstCards.Index = "PrimaryKey"
    Call grstCards.Seek("=", uWorkstation.CurrentCard.ID)
    If grstCards.NoMatch Then
        blnReturnVal = False
        strReply = CNX_SESSION_NOK & enuOpenSessionReply.NOK_NoCard & CNX_MSG_END
    Else
        If grstCards!time_credit = NULL_DATE And Not grstCards!unlimited_play Then
            blnReturnVal = False
            strReply = CNX_SESSION_NOK & enuOpenSessionReply.NOK_CardBalanceEmpty & CNX_MSG_END
        ElseIf grstCards!Active = False Then
            blnReturnVal = False
            strReply = CNX_SESSION_NOK & enuOpenSessionReply.NOK_CardDeactivated & CNX_MSG_END
        Else
            blnReturnVal = True
            uWorkstation.CurrentCard.Active = grstCards!Active
            uWorkstation.CurrentCard.TimeCredit = grstCards!time_credit
            uWorkstation.CurrentCard.UnlimitedPlay = grstCards!unlimited_play
            uWorkstation.CurrentCard.DateExpiry = grstCards!date_expiry
            uWorkstation.CurrentCard.DateCreation = grstCards!date_creation
            uWorkstation.CurrentCard.Comment = IIf(IsNull(grstCards!Comment), "", grstCards!Comment)
            uWorkstation.CurrentCard.CardType = IIf(IsNull(grstCards!Type), "", grstCards!Type)
            uWorkstation.CurrentCard.UnlimitedPlay = grstCards!unlimited_play
            uWorkstation.UnlimitedPlay = grstCards!unlimited_play
            strReply = CNX_SESSION_OK & uWorkstation.SessionType & CNX_PAYLOAD_TOKEN & funDateTimeToSend(uWorkstation.CurrentCard.TimeCredit) & CNX_PAYLOAD_TOKEN & funDateTimeToSend(uWorkstation.CurrentCard.DateExpiry) & CNX_PAYLOAD_TOKEN & CInt(uWorkstation.CurrentCard.UnlimitedPlay) & CNX_MSG_END
        End If
    End If
    If Not blnReturnVal Then uWorkstation.CurrentCard.ID = 0
    Call funSendData(uWorkstation.Cnx, strReply)
    funCheckCardBalance = blnReturnVal
    Exit Function

funCheckCardBalanceError:
    Call funErrorManager("modFunctions.funCheckCardBalanceError")
End Function


'*******************************************************************************
Public Function funCheckClientBalance(uWorkstation As clsWorkstation) As Boolean
Dim strReply As String
Dim blnReturnVal As Boolean
On Error GoTo funCheckClientBalanceError

    grstClients.Index = "PrimaryKey"
    Call grstClients.Seek("=", LCase(uWorkstation.CurrentClient.login))
    If grstClients.NoMatch Then
        blnReturnVal = False
        strReply = CNX_SESSION_NOK & enuOpenSessionReply.NOK_NoClient & CNX_MSG_END
    Else
        If grstClients!Passwd <> uWorkstation.CurrentClient.Password Then
            blnReturnVal = False
            strReply = CNX_SESSION_NOK & enuOpenSessionReply.NOK_NoClient & CNX_MSG_END
        ElseIf grstClients!time_credit = NULL_DATE And Not grstClients!unlimited_play Then
            blnReturnVal = False
            strReply = CNX_SESSION_NOK & enuOpenSessionReply.NOK_ClientBalanceEmpty & CNX_MSG_END
        ElseIf grstClients!Active = False Then
            blnReturnVal = False
            strReply = CNX_SESSION_NOK & enuOpenSessionReply.NOK_ClientDeactivated & CNX_MSG_END
        Else
            blnReturnVal = True
            uWorkstation.CurrentClient.Active = grstClients!Active
            uWorkstation.CurrentClient.TimeCredit = grstClients!time_credit
            uWorkstation.CurrentClient.UnlimitedPlay = grstClients!unlimited_play
            uWorkstation.CurrentClient.DateExpiry = grstClients!date_expiry
            uWorkstation.CurrentClient.DateCreation = grstClients!date_creation
            uWorkstation.CurrentClient.FullName = IIf(IsNull(grstClients!full_name), "", grstClients!full_name)
            uWorkstation.CurrentClient.Comment = IIf(IsNull(grstClients!Comment), "", grstClients!Comment)
            uWorkstation.CurrentClient.UnlimitedPlay = grstClients!unlimited_play
            uWorkstation.UnlimitedPlay = grstClients!unlimited_play
            strReply = CNX_SESSION_OK & uWorkstation.SessionType & CNX_PAYLOAD_TOKEN & funDateTimeToSend(uWorkstation.CurrentClient.TimeCredit) & CNX_PAYLOAD_TOKEN & funDateTimeToSend(uWorkstation.CurrentClient.DateExpiry) & CNX_PAYLOAD_TOKEN & CInt(uWorkstation.CurrentClient.UnlimitedPlay) & CNX_MSG_END
        End If
    End If
    If Not blnReturnVal Then
        uWorkstation.CurrentClient.login = ""
        uWorkstation.CurrentClient.Password = ""
    End If
    Call funSendData(uWorkstation.Cnx, strReply)
    funCheckClientBalance = blnReturnVal
    Exit Function

funCheckClientBalanceError:
    Call funErrorManager("modFunctions.funCheckClientBalanceError")
End Function


'*******************************************************************************
Public Function funCheckSelfServicePrePaidBalance(uWorkstation As clsWorkstation) As Boolean
Dim strReply As String
Dim blnReturnVal As Boolean
On Error GoTo funCheckSelfServicePrePaidBalanceError

    
    blnReturnVal = uWorkstation.SelfServicePrePaid.TimeCredit <> NULL_DATE
    If blnReturnVal Then
        uWorkstation.SelfServicePrePaid.login = uWorkstation.Key
        strReply = CNX_SESSION_OK & uWorkstation.SessionType & CNX_PAYLOAD_TOKEN & funDateTimeToSend(uWorkstation.SelfServicePrePaid.TimeCredit) & CNX_PAYLOAD_TOKEN & CNX_MSG_END
    Else
        uWorkstation.SelfServicePrePaid.login = ""
        strReply = CNX_SESSION_NOK & enuOpenSessionReply.NOK_SelfServicePrePaidBalanceEmpty & CNX_MSG_END
    End If
    Call funSendData(uWorkstation.Cnx, strReply)
    funCheckSelfServicePrePaidBalance = blnReturnVal
    Exit Function

funCheckSelfServicePrePaidBalanceError:
    Call funErrorManager("modFunctions.funCheckSelfServicePrePaidBalanceError")
End Function


'*******************************************************************************
Public Sub subSessionOpenRequest(uWorkstation As clsWorkstation)
Dim strReply As String
Dim blnOpenSession As Boolean
Dim uTmpWorkstation As New clsWorkstation
On Error GoTo subSessionOpenRequestError
    
    Select Case uWorkstation.SessionType
    Case enuSessionType.Card
        For Each uTmpWorkstation In guWorkstations
            If uWorkstation.Key <> uTmpWorkstation.Key And uWorkstation.CurrentCard.ID = uTmpWorkstation.CurrentCard.ID Then
                blnOpenSession = False
                strReply = CNX_SESSION_NOK & enuOpenSessionReply.NOK_CardInUse & CNX_MSG_END
                Call funSendData(uWorkstation.Cnx, strReply)
                Exit Sub
            End If
        Next
        blnOpenSession = funCheckCardBalance(uWorkstation)
    
    Case enuSessionType.Client
        For Each uTmpWorkstation In guWorkstations
            If uWorkstation.Key <> uTmpWorkstation.Key And LCase(uWorkstation.CurrentClient.login) = LCase(uTmpWorkstation.CurrentClient.login) Then
                blnOpenSession = False
                strReply = CNX_SESSION_NOK & enuOpenSessionReply.NOK_ClientInUse & CNX_MSG_END
                Call funSendData(uWorkstation.Cnx, strReply)
                Exit Sub
            End If
        Next
        blnOpenSession = funCheckClientBalance(uWorkstation)
    
    Case enuSessionType.Selfservice
        blnOpenSession = True
        strReply = CNX_SESSION_OK & enuSessionType.Selfservice & CNX_PAYLOAD_TOKEN & enuOpenSessionReply.OK & CNX_MSG_END
        Call funSendData(uWorkstation.Cnx, strReply)
    
    Case enuSessionType.SelfServicePrePaid
        blnOpenSession = funCheckSelfServicePrePaidBalance(uWorkstation)
    
    End Select
'    DoEvents
    If blnOpenSession Then Call subSessionOpen(uWorkstation)
    uWorkstation.AuthenticationInProgress = False
    Exit Sub

subSessionOpenRequestError:
    Call funErrorManager("modFunctions.subSessionOpenRequestError")
End Sub


'*******************************************************************************
Public Sub subSessionOpen(uWorkstation As clsWorkstation)
Dim strReply As String
Dim lngEventType As Long
Dim strEventLabel As String
On Error GoTo subSessionOpenError

    Call grstJournSessions.AddNew
    uWorkstation.SessionID = grstJournSessions!ID
    uWorkstation.StickySessionID = grstJournSessions!ID
    grstJournSessions!card_id = uWorkstation.CurrentCard.ID
    grstJournSessions!client_login = uWorkstation.CurrentClient.login
    grstJournSessions!date_start = uWorkstation.StartTime
    grstJournSessions!Type = uWorkstation.SessionType
    grstJournSessions!login = guServOpts.uConnectedUser.login
    If uWorkstation.SessionType = enuSessionType.Selfservice Then
        Call grstJournIncome.AddNew
        grstJournIncome!session_id = grstJournSessions!ID
        grstJournIncome!Type = enuSessionType.Selfservice
        grstJournIncome!date_stamp = uWorkstation.StartTime
        grstJournIncome!login = guServOpts.uConnectedUser.login
        Call grstJournIncome.Update
    End If
    Call grstJournSessions.Update
    
    strReply = CNX_SESSION_ID & uWorkstation.SessionID & CNX_MSG_END
    Call funSendData(uWorkstation.Cnx, strReply)
    
    uWorkstation.LastVisitedTitle = ""
    uWorkstation.LastVisitedURL = ""
    uWorkstation.Status = enuStatus.cliNormal
    
    
    Set uWorkstation.ChatHistory = New Collection
        
    Call subDBAddEvent(SessionOpen, uWorkstation.SessionID, uWorkstation.CurrentCard.ID, uWorkstation.CurrentClient.login, 0, funLoadResString(10337) & uWorkstation.WorkstationName)
    Select Case uWorkstation.State
    Case enuWksnState.UsedConnected
        lngEventType = enuEvents.SessionConnected
        strEventLabel = funLoadResString(10570)
        
    Case enuWksnState.UsedDisconnected
        lngEventType = enuEvents.SessionDisconnected
        strEventLabel = funLoadResString(10571)
        
    Case enuWksnState.UsedFree
        lngEventType = enuEvents.SessionFree
        strEventLabel = funLoadResString(10625)
        
    End Select
    Call subDBAddEvent(lngEventType, uWorkstation.SessionID, uWorkstation.CurrentCard.ID, uWorkstation.CurrentClient.login, 0, strEventLabel & uWorkstation.WorkstationName, Verbose)
    Exit Sub

subSessionOpenError:
    Call funErrorManager("modFunctions.subSessionOpenError")
End Sub


'*******************************************************************************
Public Sub subSessionClose(uWorkstation As clsWorkstation)
On Error GoTo subSessionCloseError
    
    If gblnDBOK Then
        grstJournSessions.Index = "PrimaryKey"
        Call grstJournSessions.Seek("=", uWorkstation.SessionID)
        If Not grstJournSessions.NoMatch Then
            Call grstJournSessions.Edit
            grstJournSessions!date_end = Now
            uWorkstation.SessionID = -1
            uWorkstation.Status = enuStatus.cliNormal
            uWorkstation.UnlimitedPlay = False
            Call grstJournSessions.Update
            Call subDBAddEvent(enuEvents.SessionClose, uWorkstation.StickySessionID, 0, "", 0, funLoadResString(10338) & uWorkstation.WorkstationName)
            If gblnfrmRecharge And (guSelectedCard.ID = uWorkstation.CurrentCard.ID Or guSelectedClient.login = uWorkstation.CurrentClient.login) Then Call Unload(frmRecharge)
            Call subRemoveStatusBarMsg("WKSN_MSG_" & uWorkstation.ComputerMacAddress)
        End If
    End If
    Exit Sub

subSessionCloseError:
    Call funErrorManager("modFunctions.subSessionCloseError")
End Sub


'*******************************************************************************
Public Sub subSessionResume(uWorkstation As clsWorkstation, strSessionInfo() As String)
Dim strID As String
Dim dteTimeCredit As Date
On Error GoTo subSessionResumeError
    
    If gblnDBOK Then
        grstJournSessions.Index = "PrimaryKey"
        Call grstJournSessions.Seek("=", CLng(strSessionInfo(2)))
        If Not grstJournSessions.NoMatch Then
            Select Case strSessionInfo(1)
            Case enuSerializedSession.Ongoing
                uWorkstation.SessionID = grstJournSessions!ID
                uWorkstation.StickySessionID = grstJournSessions!ID
                uWorkstation.CurrentCard.ID = grstJournSessions!card_id
                uWorkstation.CurrentClient.login = grstJournSessions!client_login
                uWorkstation.StartTime = grstJournSessions!date_start
                uWorkstation.SessionType = grstJournSessions!Type
                uWorkstation.State = CInt(strSessionInfo(0))
                Select Case strSessionInfo(3)
                Case enuSessionType.Card
                    strID = uWorkstation.CurrentCard.ID
                    grstCards.Index = "PrimaryKey"
                    Call grstCards.Seek("=", strID)
                    If Not grstCards.NoMatch Then
                        uWorkstation.CurrentCard.Active = grstCards!Active
                        uWorkstation.CurrentCard.DateExpiry = grstCards!date_expiry
                        uWorkstation.CurrentCard.DateCreation = grstCards!date_creation
                        uWorkstation.CurrentCard.Comment = IIf(IsNull(grstCards!Comment), "", grstCards!Comment)
                        uWorkstation.CurrentCard.CardType = IIf(IsNull(grstCards!Type), "", grstCards!Type)
                        uWorkstation.CurrentCard.UnlimitedPlay = grstCards!unlimited_play
                        uWorkstation.UnlimitedPlay = grstCards!unlimited_play
                    End If
                Case enuSessionType.Client
                    strID = uWorkstation.CurrentClient.login
                    grstClients.Index = "PrimaryKey"
                    Call grstClients.Seek("=", strID)
                    If Not grstClients.NoMatch Then
                        uWorkstation.CurrentClient.Active = grstClients!Active
                        uWorkstation.CurrentClient.DateExpiry = grstClients!date_expiry
                        uWorkstation.CurrentClient.DateCreation = grstClients!date_creation
                        uWorkstation.CurrentClient.Comment = IIf(IsNull(grstClients!Comment), "", grstClients!Comment)
                        uWorkstation.CurrentClient.FullName = IIf(IsNull(grstClients!full_name), "", grstClients!full_name)
                        uWorkstation.CurrentClient.UnlimitedPlay = grstClients!unlimited_play
                        uWorkstation.UnlimitedPlay = grstClients!unlimited_play
                    End If
                Case enuSessionType.SelfServicePrePaid
                    uWorkstation.SelfServicePrePaid.login = uWorkstation.Key
                End Select
            
            Case enuSerializedSession.Ended
                Call grstJournSessions.Edit
                grstJournSessions!time_connected = funDateTimefromNet(strSessionInfo(5))
                grstJournSessions!time_connected_offpeak = funDateTimefromNet(strSessionInfo(6))
                grstJournSessions!time_disconnected = funDateTimefromNet(strSessionInfo(7))
                grstJournSessions!date_end = funDateTimefromNet(strSessionInfo(8))
                Select Case strSessionInfo(3)
                Case enuSessionType.Card
                    strID = strSessionInfo(9)
                    grstCards.Index = "PrimaryKey"
                    Call grstCards.Seek("=", strID)
                    If Not grstCards.NoMatch Then
                        If Not grstCards!unlimited_play Then
                            Call grstCards.Edit
                            dteTimeCredit = grstCards!time_credit - funDateTimefromNet(strSessionInfo(10))
                            grstCards!time_credit = IIf(dteTimeCredit < NULL_DATE, NULL_DATE, dteTimeCredit)
                            Call grstCards.Update
                        End If
                    End If
                
                Case enuSessionType.Client
                    strID = strSessionInfo(9)
                    grstClients.Index = "PrimaryKey"
                    Call grstClients.Seek("=", strID)
                    If Not grstClients.NoMatch Then
                        If Not grstClients!unlimited_play Then
                            Call grstClients.Edit
                            dteTimeCredit = grstClients!time_credit - funDateTimefromNet(strSessionInfo(10))
                            grstClients!time_credit = IIf(dteTimeCredit < NULL_DATE, NULL_DATE, dteTimeCredit)
                            Call grstClients.Update
                        End If
                    End If
                
                Case enuSessionType.Selfservice
                    uWorkstation.SessionID = grstJournSessions!ID
                    uWorkstation.StickySessionID = grstJournSessions!ID
                    uWorkstation.SessionType = enuSessionType.Selfservice
                    uWorkstation.StartTime = grstJournSessions!date_start
                    uWorkstation.UserBalance = funSinglefromNet(strSessionInfo(9))
                    
                    grstJournIncome.Index = "session_id"
                    Call grstJournIncome.Seek("=", CLng(strSessionInfo(2)))
                    If Not grstJournIncome.NoMatch Then
                        Call grstJournIncome.Edit
                        grstJournIncome!income = funSinglefromNet(strSessionInfo(9))
                        grstJournIncome!Time = grstJournSessions!time_connected + grstJournSessions!time_connected_offpeak * (guServOpts.uTimeRatios.intOffPeak / 100) + grstJournSessions!time_disconnected * (guServOpts.uTimeRatios.intDisconCon / 100)
                        Call grstJournIncome.Update
                    End If
                
                End Select
                Call grstJournSessions.Update
            
            End Select
            Call subDBAddEvent(enuEvents.SessionResume, CLng(strSessionInfo(2)), CLng(IIf(strSessionInfo(3) = enuSessionType.Card And IsNumeric(strID), strID, 0)), CStr(IIf(strSessionInfo(3) = enuSessionType.Client, strID, "")), 0, funLoadResString(10683) & uWorkstation.WorkstationName)
        End If
    End If
    Exit Sub

subSessionResumeError:
    Call funErrorManager("modFunctions.subSessionResumeError")
End Sub


'*******************************************************************************
Public Sub subWorkstationAdd(uNewWorkstation As clsWorkstation)
Dim strEmpry As String
On Error GoTo subWorkstationAddError
    
    
    uNewWorkstation.ConnectionForced = guServOpts.uWorkstations.intStartForceWorkstationsConnected
    uNewWorkstation.AllowGroupControl = True
    Set uNewWorkstation.ChatHistory = New Collection
    uNewWorkstation.Status = enuStatus.cliNormal
    Call uNewWorkstation.ChatHistory.Add(strEmpry)
    
    'Add workstation to the global collection
    Call guWorkstations.Add(uNewWorkstation, uNewWorkstation.Key)
    DoEvents

    'Send init params to new workstation
    Call subWorkstationInit(uNewWorkstation, False)
    Exit Sub

subWorkstationAddError:
    Call funErrorManager("modFunctions.subWorkstationAddError")
End Sub


'*******************************************************************************
Public Sub subWorkstationInit(uNewWorkstation As clsWorkstation, Optional blnActivateWorkstation As Boolean = True)
Dim rstRecharges As DAO.Recordset
Dim strBuffer As String, strRechargeBuffer As String
Dim intActivateSelfServicePostPaid As Integer
On Error GoTo subWorkstationInitError

    'Send init params to workstation
    
    If guServOpts.uPostPaidSs.intMode = enuPostpaidSsMode.RechargeBased Then
        Set rstRecharges = gdbsCybAdmin.OpenRecordset("SELECT time,price FROM recharges WHERE type_postpaid_ss=true ORDER BY time", dbOpenDynaset)
    Else
        Set rstRecharges = gdbsCybAdmin.OpenRecordset("SELECT time,price FROM recharges WHERE type_postpaid_ss=true AND usage_order>0 ORDER BY usage_order", dbOpenDynaset)
    End If
    If rstRecharges.RecordCount = 0 Then
        intActivateSelfServicePostPaid = vbUnchecked
    Else
        intActivateSelfServicePostPaid = guServOpts.uWorkstations.intActivateSelfServicePostPaid
    End If
    Do While Not rstRecharges.EOF
        strRechargeBuffer = strRechargeBuffer & funDateTimeToSend(rstRecharges!Time) & CNX_PAYLOAD_TOKEN & funSingleToSend(rstRecharges!Price) & CNX_PAYLOAD_TOKEN
        rstRecharges.MoveNext
    Loop
    Call rstRecharges.Close
    strBuffer = CNX_APP_VERSION & CLng(App.Major) * 1000000 + CLng(App.Minor) * 1000 + CLng(App.Revision) & CNX_MSG_END
    strBuffer = strBuffer & CNX_SERVER_DATE & funDateTimeToSend(Now) & CNX_MSG_END
    strBuffer = strBuffer & CNX_CURRENCY_SYMBOL & guServOpts.uConfig.strMonSymbol & CNX_PAYLOAD_TOKEN & guServOpts.uConfig.strMonFormat & CNX_MSG_END
    strBuffer = strBuffer & CNX_RATIO_DISCON_CON & guServOpts.uTimeRatios.intDisconCon & CNX_MSG_END
    strBuffer = strBuffer & CNX_OFFPEAK_PARAMS & guServOpts.uTimeRatios.intOffPeak & CNX_PAYLOAD_TOKEN & funDateTimeToSend(guServOpts.uTimeRatios.dteOffPeakStop) & CNX_PAYLOAD_TOKEN & funDateTimeToSend(guServOpts.uTimeRatios.dteOffPeakStart) & CNX_MSG_END
    strBuffer = strBuffer & CNX_WORKSTATION_MSG_OPTIONS & guServOpts.uWorkstations.uWksnMsg.intMsgCard & CNX_PAYLOAD_TOKEN & guServOpts.uWorkstations.uWksnMsg.intMsgCardAlert1 & CNX_PAYLOAD_TOKEN & guServOpts.uWorkstations.uWksnMsg.intMsgCardAlert2 & CNX_PAYLOAD_TOKEN & guServOpts.uWorkstations.uWksnMsg.intMsgSelfService & CNX_PAYLOAD_TOKEN & funSingleToSend(guServOpts.uWorkstations.uWksnMsg.sngMsgSelfServiceBalance) & CNX_MSG_END
    strBuffer = strBuffer & CNX_FORCE_CONNECTED & guServOpts.uWorkstations.intStartForceWorkstationsConnected & CNX_MSG_END
    strBuffer = strBuffer & CNX_ACTIVATE_SELFSERVICE & intActivateSelfServicePostPaid & CNX_MSG_END
    strBuffer = strBuffer & CNX_ACTIVATE_CARDS & guServOpts.uWorkstations.intActivateCards & CNX_MSG_END
    strBuffer = strBuffer & CNX_ACTIVATE_CLIENTS & guServOpts.uWorkstations.intActivateClients & CNX_MSG_END
    strBuffer = strBuffer & CNX_PASSWD_ADMIN & guServOpts.uConfig.strPasswd & CNX_MSG_END
    strBuffer = strBuffer & CNX_PASSWD_MANGR & guServOpts.uConnectedUser.login & CNX_PAYLOAD_TOKEN & guServOpts.uConnectedUser.Passwd & CNX_MSG_END
    strBuffer = strBuffer & CNX_HIDE_CARD_NUM & guServOpts.uWorkstations.intHideCardNum & CNX_MSG_END
    strBuffer = strBuffer & CNX_SESSION_BATCH_START & guServOpts.uSessionBatch.intEnableStart & CNX_MSG_END
    strBuffer = strBuffer & CNX_SESSION_BATCH_END & guServOpts.uSessionBatch.intEnableEnd & CNX_MSG_END
    strBuffer = strBuffer & CNX_SESSION_BATCH_STYLE & guServOpts.uSessionBatch.intWindowStyle & CNX_MSG_END
    strBuffer = strBuffer & CNX_ENABLE_PAUSE & guServOpts.uWorkstations.intEnablePause & CNX_MSG_END
    strBuffer = strBuffer & CNX_WKSN_DEFAULT_STATE & guServOpts.uWorkstations.intDefaultState & CNX_MSG_END
    strBuffer = strBuffer & CNX_WKSN_DEFAULT_MODE & guServOpts.uWorkstations.intDefaultMode & CNX_MSG_END
    strBuffer = strBuffer & CNX_WKSN_DEFAULT_USAGE & guServOpts.uWorkstations.intUsageDefault & CNX_MSG_END
    strBuffer = strBuffer & CNX_DEFAULT_LANG & right(guServOpts.uConfig.strInterfaceLang, Len(guServOpts.uConfig.strInterfaceLang) - Len(App.EXEName)) & CNX_PAYLOAD_TOKEN & guServOpts.uConfig.intInterfaceLangCharset & CNX_MSG_END
    strBuffer = strBuffer & CNX_MSG_LOGO_FILES_TYPE & guServOpts.uWorkstations.intModifyMsgLogoFiles & CNX_MSG_END
    strBuffer = strBuffer & CNX_KILLED_APPS & guServOpts.uKilledApps.intForbidIE & CNX_PAYLOAD_TOKEN & guServOpts.uKilledApps.intForbidMozilla & CNX_PAYLOAD_TOKEN & guServOpts.uKilledApps.intForbidNetscape & CNX_PAYLOAD_TOKEN & guServOpts.uKilledApps.intForbidOE & CNX_PAYLOAD_TOKEN & funConcatToken(guServOpts.uKilledApps.strKilledApps, FILENAME_TOKEN) & CNX_MSG_END
    strBuffer = strBuffer & CNX_SESSION_CLEANUP & guServOpts.uSessionCleanup.intAlignDesktopIcons & CNX_PAYLOAD_TOKEN & guServOpts.uSessionCleanup.intDeleteDocHistory & CNX_PAYLOAD_TOKEN & guServOpts.uSessionCleanup.intRestoreBackground & CNX_PAYLOAD_TOKEN & guServOpts.uSessionCleanup.intRestoreIEHomePage & CNX_PAYLOAD_TOKEN & guServOpts.uSessionCleanup.intRestoreIESettings & CNX_PAYLOAD_TOKEN & guServOpts.uSessionCleanup.strIEHomePagePath & CNX_PAYLOAD_TOKEN & guServOpts.uSessionCleanup.intRestoreIMs & CNX_MSG_END
    strBuffer = strBuffer & CNX_POSTPAID_MODE & guServOpts.uPostPaidSs.intMode & CNX_PAYLOAD_TOKEN & funDateTimeToSend(guServOpts.uPostPaidSs.dteBaseTimeIncrement) & CNX_PAYLOAD_TOKEN & funDateTimeToSend(guServOpts.uPostPaidSs.dteMinBilledTime) & CNX_PAYLOAD_TOKEN & guServOpts.uPostPaidSs.intSubMode & CNX_PAYLOAD_TOKEN & funDateTimeToSend(guServOpts.uPostPaidSs.dteEndTimeIncrement) & CNX_MSG_END
    strBuffer = strBuffer & CNX_DEFAULT_RECHARGES & strRechargeBuffer & CNX_MSG_END
    If blnActivateWorkstation Then strBuffer = strBuffer & CNX_ACTIVATE_WORKSTATION & CNX_MSG_END
    If Not funSendData(uNewWorkstation.Cnx, strBuffer) Then Exit Sub
    DoEvents
    
    If guServOpts.uWorkstations.intModifyMsgLogoFiles = vbChecked Then
        If Not funSendFile(uNewWorkstation.Cnx, guServOpts.uWorkstations.uPaths.strBlockedPath, LoadBlock) Then Exit Sub
        DoEvents
        If Not funSendFile(uNewWorkstation.Cnx, guServOpts.uWorkstations.uPaths.strIntroPath, LoadIntro) Then Exit Sub
        DoEvents
        If Not funSendFile(uNewWorkstation.Cnx, guServOpts.uWorkstations.uPaths.strLogoPath, LoadLogo) Then Exit Sub
        DoEvents
        If Not funSendFile(uNewWorkstation.Cnx, guServOpts.uWorkstations.uPaths.strNoCreditPath, LoadNoCredit) Then Exit Sub
        DoEvents
        If Not funSendFile(uNewWorkstation.Cnx, guServOpts.uWorkstations.uPaths.strSessionEndPath, LoadSessionEnd) Then Exit Sub
        DoEvents
        If Not funSendFile(uNewWorkstation.Cnx, guServOpts.uWorkstations.uPaths.strSessionPausedPath, LoadSessionPaused) Then Exit Sub
        DoEvents
    End If
    If guServOpts.uSessionCleanup.intRestoreBackground = vbChecked Then
        If Not funSendFile(uNewWorkstation.Cnx, guServOpts.uSessionCleanup.strBackgroundPath, LoadBackground) Then Exit Sub
        DoEvents
    End If
    If guServOpts.uSessionBatch.intEnableStart = vbChecked Then
        If Not funSendFile(uNewWorkstation.Cnx, guServOpts.uSessionBatch.strStartPath, SessionBatchStart) Then Exit Sub
        DoEvents
    End If
    If guServOpts.uSessionBatch.intEnableEnd = vbChecked Then
        If Not funSendFile(uNewWorkstation.Cnx, guServOpts.uSessionBatch.strEndPath, SessionBatchEnd) Then Exit Sub
        DoEvents
    End If
    Exit Sub

subWorkstationInitError:
     Call funErrorManager("modFunctions.subWorkstationInitError")
End Sub


'*******************************************************************************
Public Sub subQuitPassword(frmParent As Form)
On Error GoTo subQuitPasswordError

    Call frmPasswdClose.Show(vbModal, frmParent)
    Exit Sub

subQuitPasswordError:
    Call funErrorManager("modFunctions.subQuitPasswordError")
End Sub



'*******************************************************************************
Public Sub subCheckForRunningServer()
On Error GoTo subCheckForRunningServerError

    If App.PrevInstance Then
        Call funMsg(AppAlreadyRunning)
        End
    End If
    Exit Sub

subCheckForRunningServerError:
    Call funErrorManager("modFunctions.subCheckForRunningServerError")
End Sub


'*******************************************************************************
Public Sub subRemoveStatusBarMsg(strKey As String)
On Error Resume Next

    Do
        Call guStatusBarMessages.Remove(strKey)
'        DoEvents
    Loop While Err.Number = 0
End Sub



'*******************************************************************************
Public Function funAddStatusBarMsg(uMsg As clsStatusBarMessage, strKey As String) As Boolean
On Error GoTo funAddStatusBarMsgError

    Call guStatusBarMessages.Add(uMsg, strKey)
    funAddStatusBarMsg = True
    Exit Function
    
funAddStatusBarMsgError:
    funAddStatusBarMsg = False
End Function


'*******************************************************************************
Public Sub subToggleForceConnected()
On Error GoTo subToggleForceConnectedError
Dim strBuffer As String

    If guSelectedWorkstation.ConnectionForced = vbUnchecked Then
        If funMsg(enuMsg.AskCanForceConnected) = vbYes Then
            strBuffer = CNX_FORCE_CONNECTED & vbChecked & CNX_MSG_END
            Call funSendData(guSelectedWorkstation.Cnx, strBuffer)
            guSelectedWorkstation.ConnectionForced = vbChecked
        End If
    Else
        strBuffer = CNX_FORCE_CONNECTED & vbUnchecked & CNX_MSG_END
        Call funSendData(guSelectedWorkstation.Cnx, strBuffer)
        guSelectedWorkstation.ConnectionForced = vbUnchecked
    End If
    frmMain.chkWksnConnect.value = guSelectedWorkstation.ConnectionForced
    frmMain.mnuWksnConnect.Checked = guSelectedWorkstation.ConnectionForced = vbChecked
    Call funSetFocus(frmMain.ucWorkstations(1))
    Exit Sub
    
subToggleForceConnectedError:
    Call funErrorManager("modFunctions.subToggleForceConnectedError")
End Sub


'*******************************************************************************
Public Sub subWorkstationShutdown(uWorkstation As clsWorkstation, intFlag As Integer)
Dim strBuffer As String
On Error GoTo subWorkstationShutdownError

    If uWorkstation.State <> enuWksnState.NotPresent Then
        strBuffer = CNX_SHUTDOWN_WORKSTATION & intFlag & CNX_MSG_END
        Call funSendData(uWorkstation.Cnx, strBuffer)
    End If
    Exit Sub
    
subWorkstationShutdownError:
    Call funErrorManager("modFunctions.subWorkstationShutdownError")
End Sub


'*******************************************************************************
Public Sub subWorkstationDisable(uWorkstation As clsWorkstation, intReason As enuWksnDisableReason)
Dim strBuffer As String
On Error Resume Next

    strBuffer = CNX_DISABLE_WORKSTATION & intReason & CNX_MSG_END
    Call funSendData(uWorkstation.Cnx, strBuffer)
End Sub


'*******************************************************************************
Public Sub subProcessURL(uWorkstation As clsWorkstation, strNewURL() As String)
Dim lngIEWinh As Long
Dim strIETitle As String
Dim strIEURL As String
On Error GoTo subProcessURLError

    lngIEWinh = CLng(strNewURL(0))
    strIETitle = strNewURL(1)
    strIEURL = strNewURL(2)
    
    uWorkstation.LastVisitedTitle = IIf(strIETitle <> "", strIETitle, funLoadResString(10693))
    uWorkstation.LastVisitedURL = strIEURL
    If guServOpts.uWorkstations.intLogURLs = vbChecked Then Call subDBAddURL(uWorkstation.SessionID, uWorkstation.CurrentCard.ID, uWorkstation.CurrentClient.login, uWorkstation.WorkstationName, lngIEWinh, strIETitle, strIEURL)
    Exit Sub
    
subProcessURLError:
    Call funErrorManager("modFunctions.subProcessURLError")
End Sub


'*******************************************************************************
Public Sub subWorkstationLockdown(uWorkstation As clsWorkstation)
Dim lngResult As Long
Dim hKey As Long, hEnumKey As Long, nResult As Long
Dim hRemoteKey As Long, nBytes As Long
Dim strUserName As String, strUserSID As String
Dim lngIndex As Long
Dim lngDriveFlags As Long
Dim strObject As String
Dim uLastWriteTime As FILETIME
Dim blnUsernameFound As Boolean
Static blnNoAccess As Boolean
On Error GoTo subWorkstationLockdownError
    
    ' Make shure no lockdown is done if it is explicitly refused by the server or the wksn
    If guServOpts.uWorkstations.intModifyLockdown = vbUnchecked Or uWorkstation.AllowLockdown = vbUnchecked Then Exit Sub
    
    blnNoAccess = GetRegSettings(funLoadResString(10006), "Workstations", "Warn on no access", 0, REG_DWORD) = 1
    
    ' Attempt to connect to the remote registry
    lngResult = RegConnectRegistry("\\" & uWorkstation.ComputerName, HKEY_USERS, hRemoteKey)
    If lngResult = RPC_S_CALL_FAILED Then lngResult = RegConnectRegistry("\\" & uWorkstation.ComputerName, HKEY_USERS, hRemoteKey)
    If lngResult <> ERROR_SUCCESS Then ' NOK then inform the user (This would be because the wksn is runnig XPHome for ex)
        If Not blnNoAccess Then
            blnNoAccess = True
            Call funMsg(GeneralError, funLoadResString(10573) & vbNewLine & funLoadResString(10000) & funLoadResString(10574) & vbNewLine & funLoadResString(10575) & vbNewLine & funLoadResString(10881))
        End If
        Exit Sub
    End If
    
    ' Attempt to open the remote registry
    lngResult = RegOpenKeyEx(hRemoteKey, "", 0, KEY_READ, hEnumKey)
    If lngResult <> ERROR_SUCCESS Then ' NOK then inform the user (This would be because the server user is not admin on the wksn)
        If Not blnNoAccess Then Call funMsg(GeneralError, funLoadResString(10573) & vbNewLine & funLoadResString(10000) & funLoadResString(10574) & vbNewLine & funLoadResString(10575) & vbNewLine & funLoadResString(10881))
        blnNoAccess = True
        Exit Sub
    End If
    
    ' Find the current user's registry tree among all HKAU
    strObject = Space$(MAX_PATH)
    nBytes = MAX_PATH
    Do While RegEnumKeyEx(hEnumKey, lngIndex, strObject, nBytes, 0&, vbNullString, 0&, uLastWriteTime) = ERROR_SUCCESS And Not blnUsernameFound
        strUserSID = left$(strObject, nBytes)
        lngResult = RegOpenKeyEx(hRemoteKey, strUserSID & "\Software\Microsoft\Windows\CurrentVersion\Explorer", 0, KEY_ALL_ACCESS, hKey)
        
        If lngResult = ERROR_ACCESS_DENIED Then
            If Not blnNoAccess Then Call funMsg(GeneralError, funLoadResString(10573) & vbNewLine & funLoadResString(10000) & funLoadResString(10574) & vbNewLine & funLoadResString(10575))
            blnNoAccess = True
            Exit Sub
        End If
        lngResult = RegQueryValueEx(hKey, "Logon User Name", 0&, REG_SZ, ByVal strUserName, nBytes)
        strUserName = Space$(nBytes)
        lngResult = RegQueryValueEx(hKey, "Logon User Name", 0&, REG_SZ, ByVal strUserName, Len(strUserName))
        If nBytes > 0 Then
            strUserName = left$(strUserName, nBytes - 1)
            blnUsernameFound = LCase(strUserName) = LCase(uWorkstation.UserName)
        End If
        lngIndex = lngIndex + 1
        nBytes = MAX_PATH
        strObject = Space$(MAX_PATH)
    Loop
    lngResult = RegCloseKey(hKey)
    lngResult = RegCloseKey(hEnumKey)
    
    ' If the user is found the lock him down!
    If blnUsernameFound Then
        lngResult = RegCreateKeyEx(hRemoteKey, strUserSID & "\Software\Microsoft\Windows\CurrentVersion\Policies\System", 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, nResult)
        lngResult = RegSetValueEx(hKey, "DisableRegistryTools", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intNoAdminTools), 4)
        lngResult = RegSetValueEx(hKey, "DisableTaskMgr", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intNoAdminTools), 4)
        lngResult = RegSetValueEx(hKey, "NoDispCPL", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intNoCPL), 4)
        lngResult = RegSetValueEx(hKey, "DisableLockWorkstation", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intNoClose), 4)
        lngResult = RegSetValueEx(hKey, "DisableChangePassword", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intNoAdminTools), 4)
        lngResult = RegCloseKey(hKey)
   
        lngResult = RegCreateKeyEx(hRemoteKey, strUserSID & "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, nResult)
        lngResult = RegSetValueEx(hKey, "NoClose", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intNoClose), 4)
        lngResult = RegSetValueEx(hKey, "NoLogoff", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intNoClose), 4)
        lngResult = RegSetValueEx(hKey, "NoSimpleStartMenu", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intNoClose), 4)
        lngResult = RegSetValueEx(hKey, "NoRun", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intNoAdminTools), 4)
        lngResult = RegSetValueEx(hKey, "NoActiveDesktop", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intLockDesktop), 4)
        lngResult = RegSetValueEx(hKey, "NoSetTaskbar", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intLockDesktop), 4)
        lngResult = RegSetValueEx(hKey, "NoSetFolders", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intLockDesktop), 4)
        lngResult = RegSetValueEx(hKey, "NoFolderOptions", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intLockDesktop), 4)
        lngResult = RegSetValueEx(hKey, "NoTrayContextMenu", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intLockDesktop), 4)
        lngResult = RegSetValueEx(hKey, "NoViewContextMenu", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intLockDesktop), 4)
        lngResult = RegSetValueEx(hKey, "NoNetConnectDisconnect", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intLockDesktop), 4)
        lngResult = RegSetValueEx(hKey, "NoFind", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intLockDesktop), 4)
        lngResult = RegSetValueEx(hKey, "NoSaveSettings", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intLockDesktop), 4)
        lngResult = RegSetValueEx(hKey, "NoGoTo", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intLockDesktop), 4)
        lngResult = RegSetValueEx(hKey, "NoDisconnect", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intLockDesktop), 4)
        lngResult = RegSetValueEx(hKey, "NoFileAssociate", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intLockDesktop), 4)
        lngResult = RegSetValueEx(hKey, "NoRecentDocsMenu", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intLockDesktop), 4)
        lngResult = RegSetValueEx(hKey, "NoFavoritesMenu", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intLockDesktop), 4)
        lngResult = RegSetValueEx(hKey, "NoOptions", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intLockDesktop), 4)
        lngResult = RegSetValueEx(hKey, "NoNTSecurity", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intLockDesktop), 4)
        lngResult = RegSetValueEx(hKey, "NoCommonGroups", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intNoCommonGroup), 4)
        lngResult = RegSetValueEx(hKey, "NoNetHood", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intNoFileSystemAccess), 4)
        If guServOpts.uWorkstations.uLocks.intNoFileSystemAccess = vbChecked Then
            lngDriveFlags = IIf(guServOpts.uWorkstations.uLocks.intNoFileSystemAccessType = 0, WKSN_LOCK_DRIVES_ALL, guServOpts.uWorkstations.uLocks.lngLockedDrives)
        Else
            lngDriveFlags = WKSN_LOCK_DRIVES_NONE
        End If
        lngResult = RegSetValueEx(hKey, "NoDrives", 0&, REG_BINARY, CLng(lngDriveFlags), 4)
        lngResult = RegCloseKey(hKey)
    
        lngResult = RegCreateKeyEx(hRemoteKey, strUserSID & "\Software\Policies\Microsoft\Internet Explorer\Restrictions", 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, nResult)
        lngResult = RegSetValueEx(hKey, "NoBrowserOptions", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intLockIE), 4)
        lngResult = RegSetValueEx(hKey, "NoFileNew", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intLockIE), 4)
        lngResult = RegSetValueEx(hKey, "NoFindFiles", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intLockIE), 4)
        lngResult = RegSetValueEx(hKey, "NoSelectDownloadDir", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intLockIE), 4)
        lngResult = RegSetValueEx(hKey, "NoTheaterMode", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intLockIE), 4)
        lngResult = RegCloseKey(hKey)
        
        lngResult = RegCreateKeyEx(hRemoteKey, strUserSID & "\Software\Policies\Microsoft\Internet Explorer\Infodelivery\Restrictions", 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, nResult)
        lngResult = RegSetValueEx(hKey, "NoAddingChannels", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intLockIE), 4)
        lngResult = RegSetValueEx(hKey, "NoAddingSubscriptions", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intLockIE), 4)
        lngResult = RegSetValueEx(hKey, "NoSearchCustomization", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intLockIE), 4)
        lngResult = RegSetValueEx(hKey, "NoRemovingChannels", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intLockIE), 4)
        lngResult = RegSetValueEx(hKey, "NoRemovingSubscriptions", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intLockIE), 4)
        lngResult = RegCloseKey(hKey)
        
        lngResult = RegCreateKeyEx(hRemoteKey, strUserSID & "\Software\Policies\Microsoft\Internet Explorer\Control Panel", 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, nResult)
        lngResult = RegSetValueEx(hKey, "Connwiz Admin Lock", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intLockIE), 4)
        lngResult = RegCloseKey(hKey)
        
        lngResult = RegCreateKeyEx(hRemoteKey, strUserSID & "\Software\Policies\Microsoft\Internet Explorer\Main", 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, nResult)
        lngResult = RegSetValueEx(hKey, "Disable Script Debugger", 0&, REG_SZ, CStr(IIf(guServOpts.uWorkstations.uLocks.intLockIE = vbChecked, "yes", "no")), 4)
        lngResult = RegSetValueEx(hKey, "Use FormSuggest", 0&, REG_SZ, CStr(IIf(guServOpts.uWorkstations.uLocks.intLockIE = vbUnchecked, "yes", "no")), 4)
        lngResult = RegSetValueEx(hKey, "FormSuggest Passwords", 0&, REG_SZ, CStr(IIf(guServOpts.uWorkstations.uLocks.intLockIE = vbUnchecked, "yes", "no")), 4)
        lngResult = RegSetValueEx(hKey, "Err Dlg Displayed On Every Error", 0&, REG_SZ, CStr(IIf(guServOpts.uWorkstations.uLocks.intLockIE = vbUnchecked, "yes", "no")), 4)
        lngResult = RegSetValueEx(hKey, "Use Custom Search URL", 0&, REG_DWORD, CLng(IIf(guServOpts.uWorkstations.uLocks.intLockIE = vbChecked, 0, 1)), 4)
        lngResult = RegCloseKey(hKey)
        
        lngResult = RegCreateKeyEx(hRemoteKey, strUserSID & "\Software\Policies\Microsoft\WindowsMediaPlayer", 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, nResult)
        lngResult = RegSetValueEx(hKey, "NoFindNewStations", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intLockIE), 4)
        lngResult = RegSetValueEx(hKey, "NoMediaFavorites", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intLockIE), 4)
        lngResult = RegSetValueEx(hKey, "NoRadioBar", 0&, REG_DWORD, CLng(guServOpts.uWorkstations.uLocks.intLockIE), 4)
        lngResult = RegCloseKey(hKey)
        
        lngResult = RegCreateKeyEx(hRemoteKey, strUserSID & "\Software\Microsoft\Internet Connection Wizard", 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, nResult)
        lngResult = RegSetValueEx(hKey, "CanInstallISPKit5", 0&, REG_SZ, CStr(IIf(guServOpts.uWorkstations.uLocks.intLockIE = vbUnchecked, "yes", "no")), 4)
        lngResult = RegCloseKey(hKey)
    
    End If
    lngResult = RegCloseKey(hRemoteKey)
    Exit Sub
    
subWorkstationLockdownError:
    Call funErrorManager("modFunctions.subWorkstationLockdownError")
End Sub


'*******************************************************************************
Public Sub subWorkstationRelease(uWorkstation As clsWorkstation)
Dim lngResult As Long
Dim hKey As Long, hEnumKey As Long, nResult As Long
Dim hRemoteKey As Long, nBytes As Long
Dim strUserName As String, strUserSID As String
Dim lngIndex As Long
Dim lngDriveFlags As Long
Dim strObject As String
Dim uLastWriteTime As FILETIME
Dim blnUsernameFound As Boolean
Static blnNoAccess As Boolean
On Error GoTo subWorkstationReleaseError
    
    ' Attempt to connect to the remote registry
    lngResult = RegConnectRegistry("\\" & uWorkstation.ComputerName, HKEY_USERS, hRemoteKey)
    If lngResult = RPC_S_CALL_FAILED Then lngResult = RegConnectRegistry("\\" & uWorkstation.ComputerName, HKEY_USERS, hRemoteKey)
    If lngResult <> ERROR_SUCCESS Then ' NOK then inform the user (This would be because the wksn is runnig XPHome for ex)
        If Not blnNoAccess Then
            blnNoAccess = True
            Call funMsg(GeneralError, funLoadResString(10573) & vbNewLine & funLoadResString(10000) & funLoadResString(10975) & vbNewLine & funLoadResString(10575))
        End If
        Exit Sub
    End If
    
    ' Attempt to open the remote registry
    lngResult = RegOpenKeyEx(hRemoteKey, "", 0, KEY_READ, hEnumKey)
    If lngResult <> ERROR_SUCCESS Then ' NOK then inform the user (This would be because the server user is not admin on the wksn)
        If Not blnNoAccess Then Call funMsg(GeneralError, funLoadResString(10573) & vbNewLine & funLoadResString(10000) & funLoadResString(10975) & vbNewLine & funLoadResString(10575))
        blnNoAccess = True
        Exit Sub
    End If
    
    ' Find the current user's registry tree among all HKAU
    strObject = Space$(MAX_PATH)
    nBytes = MAX_PATH
    Do While RegEnumKeyEx(hEnumKey, lngIndex, strObject, nBytes, 0&, vbNullString, 0&, uLastWriteTime) = ERROR_SUCCESS And Not blnUsernameFound
        strUserSID = left$(strObject, nBytes)
        lngResult = RegOpenKeyEx(hRemoteKey, strUserSID & "\Software\Microsoft\Windows\CurrentVersion\Explorer", 0, KEY_ALL_ACCESS, hKey)
        
        If lngResult = ERROR_ACCESS_DENIED Then
            If Not blnNoAccess Then Call funMsg(GeneralError, funLoadResString(10573) & vbNewLine & funLoadResString(10000) & funLoadResString(10975) & vbNewLine & funLoadResString(10575))
            blnNoAccess = True
            Exit Sub
        End If
        lngResult = RegQueryValueEx(hKey, "Logon User Name", 0&, REG_SZ, ByVal strUserName, nBytes)
        strUserName = Space$(nBytes)
        lngResult = RegQueryValueEx(hKey, "Logon User Name", 0&, REG_SZ, ByVal strUserName, Len(strUserName))
        If nBytes > 0 Then
            strUserName = left$(strUserName, nBytes - 1)
            blnUsernameFound = LCase(strUserName) = LCase(uWorkstation.UserName)
        End If
        lngIndex = lngIndex + 1
        nBytes = MAX_PATH
        strObject = Space$(MAX_PATH)
    Loop
    lngResult = RegCloseKey(hKey)
    lngResult = RegCloseKey(hEnumKey)
    
    ' If the user is found the lock him down!
    If blnUsernameFound Then
        lngResult = RegCreateKeyEx(hRemoteKey, strUserSID & "\Software\Microsoft\Windows\CurrentVersion\Policies\System", 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, nResult)
        lngResult = RegSetValueEx(hKey, "DisableRegistryTools", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "DisableTaskMgr", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoDispCPL", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "DisableLockWorkstation", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "DisableChangePassword", 0&, REG_DWORD, 0, 4)
        lngResult = RegCloseKey(hKey)
   
        lngResult = RegCreateKeyEx(hRemoteKey, strUserSID & "\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, nResult)
        lngResult = RegSetValueEx(hKey, "NoClose", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoLogoff", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoSimpleStartMenu", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoRun", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoActiveDesktop", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoSetTaskbar", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoSetFolders", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoFolderOptions", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoTrayContextMenu", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoViewContextMenu", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoNetConnectDisconnect", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoFind", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoSaveSettings", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoGoTo", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoDisconnect", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoFileAssociate", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoRecentDocsMenu", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoFavoritesMenu", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoOptions", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoNTSecurity", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoCommonGroups", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoNetHood", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoDrives", 0&, REG_BINARY, WKSN_LOCK_DRIVES_NONE, 4)
        lngResult = RegCloseKey(hKey)
    
        lngResult = RegCreateKeyEx(hRemoteKey, strUserSID & "\Software\Policies\Microsoft\Internet Explorer\Restrictions", 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, nResult)
        lngResult = RegSetValueEx(hKey, "NoBrowserOptions", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoFileNew", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoFindFiles", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoSelectDownloadDir", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoTheaterMode", 0&, REG_DWORD, 0, 4)
        lngResult = RegCloseKey(hKey)
        
        lngResult = RegCreateKeyEx(hRemoteKey, strUserSID & "\Software\Policies\Microsoft\Internet Explorer\Infodelivery\Restrictions", 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, nResult)
        lngResult = RegSetValueEx(hKey, "NoAddingChannels", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoAddingSubscriptions", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoSearchCustomization", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoRemovingChannels", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoRemovingSubscriptions", 0&, REG_DWORD, 0, 4)
        lngResult = RegCloseKey(hKey)
        
        lngResult = RegCreateKeyEx(hRemoteKey, strUserSID & "\Software\Policies\Microsoft\Internet Explorer\Control Panel", 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, nResult)
        lngResult = RegSetValueEx(hKey, "Connwiz Admin Lock", 0&, REG_DWORD, 0, 4)
        lngResult = RegCloseKey(hKey)
        
        lngResult = RegCreateKeyEx(hRemoteKey, strUserSID & "\Software\Policies\Microsoft\Internet Explorer\Main", 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, nResult)
        lngResult = RegSetValueEx(hKey, "Disable Script Debugger", 0&, REG_SZ, "no", 4)
        lngResult = RegSetValueEx(hKey, "Use FormSuggest", 0&, REG_SZ, "yes", 4)
        lngResult = RegSetValueEx(hKey, "FormSuggest Passwords", 0&, REG_SZ, "yes", 4)
        lngResult = RegSetValueEx(hKey, "Err Dlg Displayed On Every Error", 0&, REG_SZ, "yes", 4)
        lngResult = RegSetValueEx(hKey, "Use Custom Search URL", 0&, REG_DWORD, 1, 4)
        lngResult = RegCloseKey(hKey)
        
        lngResult = RegCreateKeyEx(hRemoteKey, strUserSID & "\Software\Policies\Microsoft\WindowsMediaPlayer", 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, nResult)
        lngResult = RegSetValueEx(hKey, "NoFindNewStations", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoMediaFavorites", 0&, REG_DWORD, 0, 4)
        lngResult = RegSetValueEx(hKey, "NoRadioBar", 0&, REG_DWORD, 0, 4)
        lngResult = RegCloseKey(hKey)
        
        lngResult = RegCreateKeyEx(hRemoteKey, strUserSID & "\Software\Microsoft\Internet Connection Wizard", 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, nResult)
        lngResult = RegSetValueEx(hKey, "CanInstallISPKit5", 0&, REG_SZ, "yes", 4)
        lngResult = RegCloseKey(hKey)
    
        Call funMsg(Info, funLoadResString(10976, uWorkstation.WorkstationName))
    
    End If
    lngResult = RegCloseKey(hRemoteKey)
    Exit Sub
    
subWorkstationReleaseError:
    Call funErrorManager("modFunctions.subWorkstationReleaseError")
End Sub


'*******************************************************************************
Public Sub subWorkstationShutdownAll(intFlag As Integer)
Dim strBuffer As String
On Error GoTo subWorkstationShutdownAllError

    strBuffer = CNX_SHUTDOWN_WORKSTATION & intFlag & CNX_MSG_END
    Call subBroadcastData(strBuffer)
    Exit Sub
    
subWorkstationShutdownAllError:
    Call funErrorManager("modFunctions.subWorkstationShutdownAllError")
End Sub


'*******************************************************************************
Public Sub subWorkstationShutdownUnused(intFlag As Integer)
Dim uWorkstation As clsWorkstation
On Error GoTo subWorkstationShutdownUnusedError

    For Each uWorkstation In guWorkstations
        If uWorkstation.State = enuWksnState.Ready Then Call subWorkstationShutdown(uWorkstation, intFlag)
    Next uWorkstation
    Exit Sub
    
subWorkstationShutdownUnusedError:
    Call funErrorManager("modFunctions.subWorkstationShutdownUnusedError")
End Sub


'*******************************************************************************
Public Function funNewCardNumber() As Long
Dim lngNewCardID As Long
Dim rstCards As DAO.Recordset
Dim rstCardsPending As DAO.Recordset
Dim rstClients As DAO.Recordset
On Error GoTo funNewCardNumberError

    Set rstCards = gdbsCybAdmin.OpenRecordset("cards")
    rstCards.Index = "PrimaryKey"
    Set rstCardsPending = gdbsCybAdmin.OpenRecordset("cards_pending")
    rstCardsPending.Index = "PrimaryKey"
    Set rstClients = gdbsCybAdmin.OpenRecordset("clients")
    rstClients.Index = "PrimaryKey"
    Do
        Call Randomize
        lngNewCardID = Int((999999# - 100000 + 1) * Rnd + 100000)
        Call rstCards.Seek("=", lngNewCardID)
        Call rstCardsPending.Seek("=", lngNewCardID)
        Call rstClients.Seek("=", lngNewCardID)
    Loop While Not rstCards.NoMatch Or Not rstCardsPending.NoMatch Or Not rstClients.NoMatch
    funNewCardNumber = lngNewCardID
    Exit Function
    
funNewCardNumberError:
    Call funErrorManager("modFunctions.funNewCardNumberError")
End Function


'*******************************************************************************
Public Sub subBroadcastData(strData As String)
Dim uWorkstation As New clsWorkstation
On Error GoTo subBroadcastDataError

    For Each uWorkstation In guWorkstations
        Call funSendData(uWorkstation.Cnx, strData)
'        DoEvents
    Next uWorkstation
    Exit Sub
    
subBroadcastDataError:
    Call funErrorManager("modFunctions.subBroadcastDataError")
End Sub


'*******************************************************************************
Public Sub subWorkstationCrdSessionOpen(uWorkstation As clsWorkstation, uCard As clsCard)
Dim strBuffer As String
On Error GoTo subWorkstationCrdSessionOpenError

    strBuffer = CNX_SESSION_OPEN & enuSessionType.Card & CNX_PAYLOAD_TOKEN & "" & CNX_PAYLOAD_TOKEN & uCard.ID & CNX_MSG_END
    Call funSendData(uWorkstation.Cnx, strBuffer)
    DoEvents
    Call subUpdateUISelectedCard
    Call funMsg(SessionStarted, uWorkstation.WorkstationName)
    Exit Sub
    
subWorkstationCrdSessionOpenError:
    Call funErrorManager("modFunctions.subWorkstationCrdSessionOpenError")
End Sub


'*******************************************************************************
Public Sub subWorkstationCliSessionOpen(uWorkstation As clsWorkstation, uClient As clsClient)
Dim strBuffer As String
On Error GoTo subWorkstationCliSessionOpenError

    strBuffer = CNX_SESSION_OPEN & enuSessionType.Client & CNX_PAYLOAD_TOKEN & uClient.Password & CNX_PAYLOAD_TOKEN & uClient.login & CNX_MSG_END
    Call funSendData(uWorkstation.Cnx, strBuffer)
    DoEvents
    Call subUpdateUISelectedClient
    Call funMsg(SessionStarted, uWorkstation.WorkstationName)
    Exit Sub
    
subWorkstationCliSessionOpenError:
    Call funErrorManager("modFunctions.subWorkstationCliSessionOpenError")
End Sub


'*******************************************************************************
Public Sub subWorkstationSelfServicePrePaidSessionOpen(uWorkstation As clsWorkstation)
Dim strBuffer As String
On Error GoTo subWorkstationSelfServicePrePaidSessionOpenError

    strBuffer = CNX_SESSION_OPEN & enuSessionType.SelfServicePrePaid & CNX_PAYLOAD_TOKEN & "" & CNX_PAYLOAD_TOKEN & "" & CNX_PAYLOAD_TOKEN & "" & CNX_MSG_END
    Call funSendData(uWorkstation.Cnx, strBuffer)
    DoEvents
'    Call funMsg(SessionStarted, uWorkstation.WorkstationName)
    Exit Sub
    
subWorkstationSelfServicePrePaidSessionOpenError:
    Call funErrorManager("modFunctions.subWorkstationSelfServicePrePaidSessionOpenError")
End Sub


'*******************************************************************************
Public Function funAnyActiveWorkstation() As Boolean
Dim i As Integer
On Error GoTo funAnyActiveWorkstationError

    For i = frmMain.ucWorkstations.LBound To frmMain.ucWorkstations.UBound
        If frmMain.ucWorkstations(i).Active Then funAnyActiveWorkstation = True
    Next i
    Exit Function
    
funAnyActiveWorkstationError:
    Call funErrorManager("modFunctions.funAnyActiveWorkstationError")
End Function


'*******************************************************************************
Public Function funLoadResString(lngResID As Long, Optional strInsert1 As String = "", Optional strInsert2 As String = "", Optional strInsert3 As String = "", Optional strInsert4 As String = "") As String
Dim uResfile As Object
Dim strResString As String
On Error GoTo funLoadResStringError

    On Error Resume Next
    Set uResfile = CreateObject(guServOpts.uConfig.strInterfaceLang & ".clsResources")
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
Public Sub subSetAccessJournal(uUser As clsUser)
Dim i As Integer
On Error GoTo subSetAccessJournalError

    frmMain.cmdIncomeFilter.enabled = uUser.AccessJournals
    frmMain.cmdIncomeExport.enabled = uUser.AccessJournals And frmMain.cmdIncomeExport.enabled
    frmMain.cmdIncomePrint.enabled = uUser.AccessJournals And frmMain.cmdIncomePrint.enabled And guServOpts.uConfig.strDefaultPrinterName <> ""
    frmMain.mnuIncomeApply.enabled = uUser.AccessJournals
    frmMain.mnuIncomeExport.enabled = frmMain.cmdIncomeExport.enabled
    frmMain.mnuIncomePrint.enabled = frmMain.cmdIncomePrint.enabled
    frmMain.cmdSessionFilter.enabled = uUser.AccessJournals
    frmMain.cmdSessionExport.enabled = uUser.AccessJournals And frmMain.cmdSessionExport.enabled
    frmMain.cmdSessionPrint.enabled = uUser.AccessJournals And frmMain.cmdSessionPrint.enabled And guServOpts.uConfig.strDefaultPrinterName <> ""
    frmMain.mnuSessionApply.enabled = uUser.AccessJournals
    frmMain.mnuSessionExport.enabled = frmMain.cmdSessionExport.enabled
    frmMain.mnuSessionPrint.enabled = frmMain.cmdSessionPrint.enabled
    For i = 0 To frmMain.optIncomeFilter.UBound
        frmMain.optIncomeFilter(i).enabled = uUser.AccessJournals
    Next i
    For i = 0 To frmMain.optSessionFilter.UBound
        frmMain.optSessionFilter(i).enabled = uUser.AccessJournals
    Next i
    For i = 1 To frmMain.chkTypeIncome.UBound
        frmMain.chkTypeIncome(i).enabled = uUser.AccessJournals
    Next i
    For i = 1 To frmMain.chkSessStatType.UBound
        frmMain.chkSessStatType(i).enabled = uUser.AccessJournals
    Next i
    If Not uUser.AccessJournals Then Call subInitIncome
    If Not uUser.AccessJournals Then Call subInitSessions
    Exit Sub
    
subSetAccessJournalError:
    Call funErrorManager("modFunctions.subSetAccessJournalError")
End Sub

    
'*******************************************************************************
Public Sub subSetAccessUnblock(uUser As clsUser)
Dim strBuffer As String
On Error GoTo subSetAccessUnblockError

    frmMain.cmdWksnButtons(3).enabled = uUser.AccessUnblock
    frmMain.cmdConsoleButtons(2).enabled = uUser.AccessUnblock
    frmMain.mnuWksnUnblock.enabled = uUser.AccessUnblock
    frmMain.mnuWksnUnblockDD.enabled = uUser.AccessUnblock
    If uUser.AccessUnblock Then
        strBuffer = CNX_PASSWD_MANGR & uUser.login & CNX_PAYLOAD_TOKEN & uUser.Passwd & CNX_MSG_END
    Else
        strBuffer = CNX_PASSWD_MANGR & ADMIN_LOGON & CNX_PAYLOAD_TOKEN & guServOpts.uConfig.strPasswd & CNX_MSG_END
    End If
    Call subBroadcastData(strBuffer)
    Exit Sub
    
subSetAccessUnblockError:
    Call funErrorManager("modFunctions.subSetAccessUnblockError")
End Sub

    
'*******************************************************************************
Public Sub subSetAccessMassCardCreation(uUser As clsUser)
On Error GoTo subSetAccessMassCardCreationError

    frmMain.mnuCrd(10).enabled = uUser.AccessMassCardCreation And gblnDBOK
    Exit Sub
    
subSetAccessMassCardCreationError:
    Call funErrorManager("modFunctions.subSetAccessMassCardCreationError")
End Sub

    
'*******************************************************************************
Public Sub subSetAccessCardCreation(uUser As clsUser)
On Error GoTo subSetAccessCardCreationError

    frmMain.mnuCrd(0).enabled = uUser.AccessCardCreation And gblnDBOK
    frmMain.cmdCrdButtons(0).enabled = uUser.AccessCardCreation And gblnDBOK
    Exit Sub
    
subSetAccessCardCreationError:
    Call funErrorManager("modFunctions.subSetAccessCardCreationError")
End Sub

    
'*******************************************************************************
Public Sub subSetAccessClientCreation(uUser As clsUser)
On Error GoTo subSetAccessClientCreationError

    frmMain.mnuCli(0).enabled = uUser.AccessClientCreation And gblnDBOK
    frmMain.cmdCliButtons(0).enabled = uUser.AccessClientCreation And gblnDBOK
    Exit Sub
    
subSetAccessClientCreationError:
    Call funErrorManager("modFunctions.subSetAccessClientCreationError")
End Sub

    
'*******************************************************************************
Public Function funUserLogin(uUser As clsUser) As Boolean
On Error GoTo subUserLoginError

    funUserLogin = False
    If LCase(uUser.login) = ADMIN_LOGON Then
        If uUser.Passwd = guServOpts.uConfig.strPasswd Then
            uUser.login = ADMIN_LOGON
            uUser.Passwd = guServOpts.uConfig.strPasswd
            uUser.FullName = ADMIN_LOGON
            uUser.AccessJournals = True
            uUser.AccessUnblock = True
            uUser.AccessDelRecharge = True
            uUser.AccessMassCardCreation = True
            uUser.AccessCardCreation = True
            uUser.AccessClientCreation = True
            uUser.DateLastLogin = NULL_DATE
            funUserLogin = True
            frmMain.mnuChangePasswdUser.enabled = False
            frmMain.mnuWksnUserRelease.Visible = True
            frmMain.mnuWksnUserReleaseDD.Visible = True
        End If
    Else
        If gblnDBOK Then
            grstUsers.Index = "PrimaryKey"
            Call grstUsers.Seek("=", LCase(uUser.login))
            If Not grstUsers.NoMatch Then
                If uUser.Passwd = funDoXOR(grstUsers!Passwd, funLoadResString(10006)) Then
                    uUser.login = grstUsers!login
                    uUser.Passwd = funDoXOR(grstUsers!Passwd, funLoadResString(10006))
                    uUser.FullName = grstUsers!full_name
                    uUser.AccessJournals = grstUsers!access_journals
                    uUser.AccessUnblock = grstUsers!access_unblock
                    uUser.AccessDelRecharge = grstUsers!access_del_recharge
                    uUser.AccessMassCardCreation = grstUsers!access_mass_card_creation
                    uUser.AccessCardCreation = grstUsers!access_card_creation
                    uUser.AccessClientCreation = grstUsers!access_client_creation
                    uUser.DateLastLogin = IIf(IsNull(grstUsers!last_login), NULL_DATE, grstUsers!last_login)
                    Call grstUsers.Edit
                    grstUsers!last_login = Now
                    Call grstUsers.Update
                    funUserLogin = True
                    frmMain.mnuChangePasswdUser.enabled = True
                    frmMain.mnuWksnUserRelease.Visible = False
                    frmMain.mnuWksnUserReleaseDD.Visible = False
                End If
            End If
        End If
    End If
    If funUserLogin Then
        Call subSetAccessJournal(uUser)
        Call subSetAccessUnblock(uUser)
        Call subSetAccessMassCardCreation(uUser)
        Call subSetAccessCardCreation(uUser)
        Call subSetAccessClientCreation(uUser)
        frmMain.stbStatus.Panels(2).Picture = frmMain.imlSatusBar.ListImages(5).Picture
        frmMain.stbStatus.Panels(2).Text = uUser.login & "  "
        frmMain.stbStatus.Panels(2).Bevel = sbrRaised
        If uUser.DateLastLogin = NULL_DATE Then
            frmMain.stbStatus.Panels(2).ToolTipText = ""
        Else
            frmMain.stbStatus.Panels(2).ToolTipText = funLoadResString(10777) & format(uUser.DateLastLogin, "Short Date") & " " & format(uUser.DateLastLogin, "Hh:Nn")
        End If
        frmMain.mnuLogoff.Caption = funLoadResString(10778) & uUser.login & "..."
        If gblnDBOK Then Call subDBAddEvent(AppLogon, 0, 0, "", 0, funLoadResString(10779) & uUser.login)
        Call SetRegUserSettings(funLoadResString(10006), "General", "Last user", uUser.login)
        Set guServOpts.uConnectedUser = uUser
        
    End If
    Exit Function

subUserLoginError:
    Call funErrorManager("modFunctions.subUserLoginError")
End Function

    
'*******************************************************************************
Public Function funUserLogoff(uUser As clsUser, Optional blnAsk As Boolean = True) As Boolean
Dim blnGo As Boolean
On Error GoTo funUserLogoffError

    funUserLogoff = True
    blnGo = True
    If blnAsk Then
        blnGo = funMsg(AskCanLogoff) = vbYes
    End If
    If blnGo Then
        If gblnDBOK And Len(uUser.login) > 0 Then Call subDBAddEvent(AppLogoff, 0, 0, "", 0, funLoadResString(10780) & uUser.login)
        uUser.login = ""
        uUser.Passwd = ""
        uUser.FullName = ""
        uUser.DateLastLogin = NULL_DATE
        uUser.AccessJournals = False
        uUser.AccessUnblock = False
        uUser.AccessDelRecharge = False
        uUser.AccessMassCardCreation = False
        uUser.AccessCardCreation = False
        uUser.AccessClientCreation = False
        frmMain.stbStatus.Panels(2).Picture = frmMain.imlSatusBar.ListImages(1).Picture
        frmMain.stbStatus.Panels(2).Text = uUser.login
        frmMain.stbStatus.Panels(2).ToolTipText = ""
        frmMain.stbStatus.Panels(2).Bevel = sbrNoBevel
        Call subSetAccessJournal(uUser)
        Call subSetAccessUnblock(uUser)
        Call subSetAccessMassCardCreation(uUser)
        Call Load(frmLogon)
        If gblnDBOK Then
            If grstUsers.EOF And grstUsers.BOF Or grstUsers.RecordCount = 0 Then
                frmLogon.txtLogin.Text = ADMIN_LOGON
            End If
        Else
            frmLogon.txtLogin.Text = ADMIN_LOGON
        End If
        If frmLogon.txtLogin.Text = ADMIN_LOGON And guServOpts.uConfig.strPasswd = "" Then
            uUser.login = ADMIN_LOGON
            Call funUserLogin(uUser)
            Call Unload(frmLogon)
        ElseIf frmLogon.txtLogin.Text = ADMIN_LOGON Then
            frmLogon.txtLogin.enabled = False
            Call frmLogon.Show(vbModal, frmMain)
        Else
            frmLogon.txtLogin.Text = GetRegUserSettings(funLoadResString(10006), "General", "Last user", "")
            frmLogon.txtLogin.enabled = True
            frmLogon.txtPasswd.Text = ""
            Call frmLogon.Show(vbModal, frmMain)
        End If
   End If
    Exit Function
    
funUserLogoffError:
    funUserLogoff = False
    Call funErrorManager("modFunctions.funUserLogoffError")
End Function


    
'*******************************************************************************
Public Sub subUpdateWksnList(strNewMac As String, strWorkstationName As String, dteDateLastConnection As Date, Optional blnCurrentyConnected As Boolean = True, Optional blnUpdateRegistry As Boolean = True)
Dim uNewMac As New clsMacAddress
Dim uTmpMac As clsMacAddress
Dim uOldMac As New clsMacAddress
Dim i As Integer, j As Integer
On Error GoTo subUpdateWksnListError

    
    If funIsInCollection(guMacAddresses, strNewMac) Then
        With guMacAddresses
            ' Update the mac addess entry
            .Item(strNewMac).DateLastConnection = dteDateLastConnection
            .Item(strNewMac).WorkstationName = strWorkstationName
            .Item(strNewMac).CurrentlyConnected = blnCurrentyConnected
            
            'Purge all other entries that have the same workstation name and are not currently connected
            For Each uTmpMac In guMacAddresses
                If uTmpMac.WorkstationName = strWorkstationName And Not uTmpMac.CurrentlyConnected Then Call guMacAddresses.Remove(uTmpMac.MacAddress)
            Next uTmpMac
        End With
    Else 'Add a new entry...
        If guMacAddresses.Count >= guServOpts.uWorkstations.lngNumWorkstations Then
            
            'If no new slots are available take the oldes unconnected slot
            uOldMac.DateLastConnection = dteDateLastConnection
            For Each uTmpMac In guMacAddresses
                If uTmpMac.DateLastConnection < uOldMac.DateLastConnection And Not uTmpMac.CurrentlyConnected Then Set uOldMac = uTmpMac
            Next uTmpMac
            Call guMacAddresses.Remove(uOldMac.MacAddress)
        End If
        uNewMac.MacAddress = strNewMac
        uNewMac.DateLastConnection = dteDateLastConnection
        uNewMac.WorkstationName = strWorkstationName
        uNewMac.CurrentlyConnected = blnCurrentyConnected
        Call guMacAddresses.Add(uNewMac, strNewMac)
    End If
    If blnUpdateRegistry Then
        ReDim guServOpts.uWorkstations.strMacAddresses(0 To guMacAddresses.Count * 3 - 1)
        For i = 1 To guMacAddresses.Count
            guServOpts.uWorkstations.strMacAddresses(i * 3 - 3) = guMacAddresses(i).MacAddress
            guServOpts.uWorkstations.strMacAddresses(i * 3 - 2) = guMacAddresses(i).WorkstationName
            guServOpts.uWorkstations.strMacAddresses(i * 3 - 1) = funDateTimeToSend(guMacAddresses(i).DateLastConnection)
        Next i
        Call SetRegSettings(funLoadResString(10006), "Workstations", "Mac Addresses", funConcatToken(guServOpts.uWorkstations.strMacAddresses, TIME_TOKEN))
    End If
    
    'Sort MACs by WorkstationName
    If guMacAddresses.Count > 1 Then
        For i = 1 To guMacAddresses.Count
            For j = 1 To guMacAddresses.Count - 1
                If LCase(guMacAddresses(j).WorkstationName) > LCase(guMacAddresses(j + 1).WorkstationName) Then
                    Set uTmpMac = guMacAddresses(j)
                    Call guMacAddresses.Remove(j)
                    Call guMacAddresses.Add(uTmpMac, uTmpMac.MacAddress, , j)
                End If
            Next j
        Next i
    End If
    Exit Sub
    
subUpdateWksnListError:
    Call funErrorManager("modFunctions.subUpdateWksnListError")
End Sub

    
'*******************************************************************************
Public Function funFindWorkstationByMac(uMac As clsMacAddress) As clsWorkstation
Dim i As Integer
On Error GoTo funFindWorkstationByMacError

    Set funFindWorkstationByMac = New clsWorkstation
    funFindWorkstationByMac.ComputerMacAddress = uMac.MacAddress
    funFindWorkstationByMac.WorkstationName = uMac.WorkstationName
    If guWorkstations.Count > 0 Then
        For i = 1 To guWorkstations.Count
            If guWorkstations(i).ComputerMacAddress = uMac.MacAddress Then
                Set funFindWorkstationByMac = guWorkstations(i)
                Exit For
            End If
        Next i
    End If
    Exit Function
    
funFindWorkstationByMacError:
    Call funErrorManager("modFunctions.funFindWorkstationByMacError")
End Function

    
'*******************************************************************************
Public Sub subWorkstationChat(uWorkstation As clsWorkstation, strText As String)
On Error GoTo subWorkstationChatError

    If strText = "" Then Exit Sub
    strText = Replace(strText, CNX_MSG_END, "")
    Call uWorkstation.ChatHistory.Add(funFormatTime(Time) & " " & guServOpts.uConnectedUser.login & ": " & strText)
    Call funSendData(uWorkstation.Cnx, CNX_CHAT_MSG & strText & CNX_MSG_END)
    
    Exit Sub
    
subWorkstationChatError:
    Call funErrorManager("modFunctions.subWorkstationChatError")
End Sub

    
'*******************************************************************************
Public Function funExportFlexGrigToText(uFxg As MSHFlexGrid) As String
Dim i As Long, j As Long
Dim strOutput As String
On Error GoTo funExportFlexGrigToTextError


    With uFxg
        If .Cols = 0 Or .Rows = 0 Then Exit Function
        For i = 0 To .Rows - 1
            For j = 0 To .Cols - 1
                strOutput = strOutput & IIf(j = 0, "", vbTab) & .TextMatrix(i, j)
            Next j
            strOutput = strOutput & IIf(i = .Rows - 1, "", vbNewLine)
        Next i
    End With
    funExportFlexGrigToText = strOutput
    Exit Function
    
funExportFlexGrigToTextError:
    Call funErrorManager("modFunctions.funExportFlexGrigToTextError")
End Function

    
'*******************************************************************************
Public Function funGetPrinerName() As String
On Error GoTo funGetPrinerNameError

    funGetPrinerName = Printer.DeviceName
    Exit Function
    
funGetPrinerNameError:
    funGetPrinerName = ""
End Function

    
'*******************************************************************************
Public Function funLeftNewLines(strString As String, intNumNewLines As Integer) As String
Dim i As Integer, lngIndex As Long
On Error GoTo funLeftNewLinesError

    lngIndex = InStr(1, strString, vbNewLine)
    Do While lngIndex <> 0 And i < intNumNewLines - 1
        lngIndex = InStr(lngIndex + 1, strString, vbNewLine) 'right(strString, Len(strString) - (lngIndex + 1))
        i = i + 1
    Loop
    If lngIndex > 0 Then
        funLeftNewLines = left(strString, lngIndex - 1) & " ..."
    Else
        funLeftNewLines = strString
    End If
    Exit Function
    
funLeftNewLinesError:
    funLeftNewLines = ""
End Function

