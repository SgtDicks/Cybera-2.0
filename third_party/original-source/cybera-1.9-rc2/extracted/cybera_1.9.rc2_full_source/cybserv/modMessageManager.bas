Attribute VB_Name = "modMessageManager"
' _____________________________________________________________________________
'|________________________________ MESSAGE GENERATION _________________________|
'|  Name:          modMessageManager.bas                                       |
'|  Type:          VB6 Module                                                  |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
'|  Description:   Contains Message deffinitions.                              |
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

' $Id: modMessageManager.bas 15 2006-01-01 16:12:35Z sam $

Option Explicit


'=== Constant deffinitions =====================================================

'--- Message type enumeration --------------------------------------------------
Public Enum enuMsg
    Info = 0                    'General info msg
    GeneralError = 1            'General Error msg
    BadPasswd = 3
    BadNewPasswd = 4
    SaveSettings = 5
    ServerAlreadyRunning = 16
    CardNotFound = 17
    CardDeleteConfirm = 18
    CardNotSelected = 19
    RechargeNotSelected = 20
    GeneralDBError = 21
    RechargeNotFound = 22
    RechargeDeleteConfirm = 23
    RechargeModifyConfirm = 24
    RechargeNotCorrectFormat = 25
    AskCanLockWorkstation = 29
    AskCanCloseSession = 31
    AskCanForceConnected = 32
    NoRechargeCardInUse = 34
    ConfirmServerShutdown = 37
    AppAlreadyRunning = 38
    AskCanPurgeCards = 39
    AskEndAppOnDBPasswdChange = 43
    UntrappedError = 44
    DBVersionError = 45
    ServiceDeleteConfirm = 46
    ServiceAddNoLabel = 47
    ExportImpossible = 48
    AskEndAppOnDBPathChange = 49
    DBPasswordError = 50
    DBPathError = 51
    OSNotNT = 52
    NetError = 53
    AskCanPurgeHistory = 54
    CardBonusReached = 55
    AskCanDeleteRechargeFromCard = 56
    WorkstationLockdownConfigHasChanged = 57
    NoCredit = 58
    SessionStarted = 59
    AskCanLockAllWorkstations = 60
    AskCanCloseAllSessions = 61
    NoElementSelected = 62
    FileNotFound = 63
    CardDisabled = 64
    AskCanDeactivateCard = 65
    AskCanActivateCard = 66
    URLNotFound = 67
    DBUpdrageWarning = 68
    AskCanRestoreDB = 69
    AddUserError = 70
    AskCanDeleteUser = 71
    AskCanLogoff = 72
    CancelWizard = 73
    AskCanBootWorkstations = 74
    AskCanDeactivateClient = 75
    AskCanActivateClient = 76
    ClientNotSelected = 77
    ClientBonusReached = 78
    ClientInUseNoModify = 79
    AskCanDeleteRechargeFromClient = 80
    ClientNotFound = 81
    AddClientError = 82
    ClientDisabled = 83
    AskCanPurgeClients = 84
    AskCanDeleteClient = 85
    AskCanGenerateCards = 86
    CardTemplateNOK = 87
    AskCanEnableCard = 88
    AskCanViewGenCards = 89
    AskCanPrint = 90
End Enum


'*******************************************************************************
Public Function funMsg(ByVal intMsgType As enuMsg, Optional strText As Variant = "", Optional strTitle As Variant, Optional strMoreText As Variant = "") As Integer
Dim strPrompt As String     'Message body
Dim strMsgTitle As String
Dim lngButtons As Long      'Message box type (buttons & icon to display)
On Error GoTo funMsgError

    If Not IsMissing(strTitle) Then strMsgTitle = strTitle
    
    Select Case intMsgType
    
    Case enuMsg.Info
        lngButtons = vbOKOnly + vbInformation
        If Not IsMissing(strText) Then strPrompt = strText
    
    Case enuMsg.GeneralError
        lngButtons = vbOKOnly + vbCritical
        If Not IsMissing(strText) Then strPrompt = strText
    
    Case enuMsg.BadPasswd
        lngButtons = vbOKOnly + vbInformation
        strPrompt = funLoadResString(10101)
    
    Case enuMsg.BadNewPasswd
        lngButtons = vbOKOnly + vbInformation
        strPrompt = funLoadResString(10105)
    
    Case enuMsg.SaveSettings
        lngButtons = vbYesNo + vbQuestion + vbDefaultButton2
        strPrompt = funLoadResString(10103)
    
    Case enuMsg.ServerAlreadyRunning
        lngButtons = vbYesNo + vbCritical + vbMsgBoxSetForeground + vbSystemModal
        strPrompt = funLoadResString(10106) & guServOpts.uServer.lngPort & "." & vbNewLine & funLoadResString(10107) & App.Title & funLoadResString(10108)
    
    Case enuMsg.CardNotFound
        lngButtons = vbOKOnly + vbExclamation
        strPrompt = funLoadResString(10109)
    
    Case enuMsg.CardDeleteConfirm
        lngButtons = vbYesNo + vbExclamation
        strPrompt = funLoadResString(10110)
    
    Case enuMsg.CardNotSelected
        lngButtons = vbOKOnly + vbExclamation
        strPrompt = funLoadResString(10111)
    
    Case enuMsg.RechargeNotSelected
        lngButtons = vbOKOnly + vbExclamation
        strPrompt = funLoadResString(10112)
    
    Case enuMsg.GeneralDBError
        lngButtons = vbYesNo + vbCritical + vbMsgBoxSetForeground + vbSystemModal
        strPrompt = funLoadResString(10260) & vbNewLine & strText & vbNewLine & funLoadResString(10261) & App.Title & funLoadResString(10262)
    
    Case enuMsg.RechargeNotFound
        lngButtons = vbOKOnly + vbExclamation
        strPrompt = funLoadResString(10263)
    
    Case enuMsg.RechargeDeleteConfirm
        lngButtons = vbYesNo + vbExclamation
        strPrompt = funLoadResString(10264)
    
    Case enuMsg.RechargeModifyConfirm
        lngButtons = vbYesNo + vbExclamation
        strPrompt = funLoadResString(10265)
    
    Case enuMsg.RechargeNotCorrectFormat
        lngButtons = vbOKOnly + vbExclamation
        strPrompt = funLoadResString(10266) & vbNewLine & vbNewLine & funLoadResString(10267) & vbNewLine & funLoadResString(10268)
    
    Case enuMsg.AskCanLockWorkstation
        lngButtons = vbYesNo + vbDefaultButton2 + vbExclamation + vbMsgBoxSetForeground
        strPrompt = funLoadResString(10269) & vbNewLine & funLoadResString(10270)
    
    Case enuMsg.AskCanCloseSession
        lngButtons = vbYesNo + vbDefaultButton2 + vbExclamation + vbMsgBoxSetForeground
        strPrompt = funLoadResString(10271)
    
    Case enuMsg.AskCanForceConnected
        lngButtons = vbYesNo + vbDefaultButton2 + vbExclamation + vbMsgBoxSetForeground
        strPrompt = funLoadResString(10272)
    
    Case enuMsg.NoRechargeCardInUse
        lngButtons = vbOKOnly + vbInformation
        strPrompt = funLoadResString(10273) & vbNewLine & funLoadResString(10274)
    
    Case enuMsg.ConfirmServerShutdown
        lngButtons = vbYesNo + vbQuestion
        strPrompt = funLoadResString(10275) & strText & funLoadResString(10276) & vbNewLine & vbNewLine & funLoadResString(10277)
    
    Case enuMsg.AppAlreadyRunning
        lngButtons = vbOKOnly + vbCritical + vbMsgBoxSetForeground + vbSystemModal
        strPrompt = App.Title & funLoadResString(10278) & vbNewLine & funLoadResString(10279)
    
    Case enuMsg.AskCanPurgeCards
        lngButtons = vbYesNo + vbExclamation
        strPrompt = funLoadResString(10280) & vbNewLine & vbNewLine & funLoadResString(10281) & vbNewLine & funLoadResString(10282) & format(DateAdd("d", -guServOpts.uCards.intCardPurgeDelta * guServOpts.uConfig.intNumDaysInMonth, Date), "Long Date") & funLoadResString(10283)
    
    Case enuMsg.AskEndAppOnDBPasswdChange
        lngButtons = vbYesNo + vbCritical + vbMsgBoxSetForeground + vbSystemModal
        strPrompt = funLoadResString(10284) & vbNewLine & funLoadResString(10285) & App.Title & funLoadResString(10262)
    
    Case enuMsg.UntrappedError
        lngButtons = vbOKOnly + vbCritical + vbMsgBoxSetForeground + vbSystemModal
        If strText = "" Then
            strPrompt = funLoadResString(10286) & vbNewLine & funLoadResString(10287) & MY_EMAIL & funLoadResString(10288) & vbNewLine & gstrLogfilePath
        Else
            strPrompt = funLoadResString(10443) & strText & vbNewLine & funLoadResString(10287) & MY_EMAIL & funLoadResString(10288) & vbNewLine & gstrLogfilePath
        End If
    
    Case enuMsg.DBVersionError
        lngButtons = vbYesNo + vbCritical + vbMsgBoxSetForeground + vbSystemModal
        strPrompt = funLoadResString(10260) & vbNewLine & funLoadResString(10317) & App.Title & vbNewLine & funLoadResString(10261) & App.Title & funLoadResString(10262)
    
    Case enuMsg.ServiceDeleteConfirm
        lngButtons = vbYesNo + vbExclamation
        strPrompt = funLoadResString(10358)
    
    Case enuMsg.ServiceAddNoLabel
        lngButtons = vbOKOnly + vbInformation
        strPrompt = funLoadResString(10360)
    
    Case enuMsg.ExportImpossible
        lngButtons = vbOKOnly + vbExclamation
        strPrompt = funLoadResString(10361) & vbNewLine & funLoadResString(10362) & strText & funLoadResString(10363) & vbNewLine & funLoadResString(10364)
    
    Case enuMsg.AskEndAppOnDBPathChange
        lngButtons = vbYesNo + vbCritical + vbMsgBoxSetForeground + vbSystemModal
        strPrompt = funLoadResString(10365) & vbNewLine & funLoadResString(10285) & App.Title & funLoadResString(10262)
    
    Case enuMsg.DBPasswordError
        lngButtons = vbOKOnly + vbCritical + vbMsgBoxSetForeground + vbSystemModal
        strPrompt = funLoadResString(10366) & vbNewLine & funLoadResString(10367) & App.Title
    
    Case enuMsg.DBPathError
        lngButtons = vbOKOnly + vbCritical + vbMsgBoxSetForeground + vbSystemModal
        strPrompt = funLoadResString(10368) & vbNewLine & strText & vbNewLine & funLoadResString(10367) & App.Title
    
    Case enuMsg.OSNotNT
        lngButtons = vbOKOnly + vbCritical
        strPrompt = funLoadResString(10375) & vbNewLine & funLoadResString(10376) & funLoadResString(10000)
    
    Case enuMsg.NetError
        lngButtons = vbOKOnly + vbCritical + vbMsgBoxSetForeground + vbSystemModal
        strPrompt = funLoadResString(10497) & vbNewLine & funLoadResString(10498)
    
    Case enuMsg.AskCanPurgeHistory
        lngButtons = vbYesNo + vbExclamation + vbDefaultButton2
        strPrompt = funLoadResString(10517) & vbNewLine & vbNewLine & funLoadResString(10515) & strText & funLoadResString(10283) & vbNewLine & vbNewLine & funLoadResString(10751)
    
    Case enuMsg.CardBonusReached
        lngButtons = vbOKOnly + vbInformation
        strPrompt = funLoadResString(10543) & strText & vbNewLine & funLoadResString(10544) & strMoreText & funLoadResString(10545)
    
    Case enuMsg.AskCanDeleteRechargeFromCard
        lngButtons = vbYesNo + vbQuestion
        strPrompt = funLoadResString(10554) & strText & vbNewLine & funLoadResString(10555) & strMoreText
    
    Case enuMsg.WorkstationLockdownConfigHasChanged
        lngButtons = vbYesNo + vbQuestion + vbDefaultButton2
        strPrompt = funLoadResString(10585) & vbNewLine & funLoadResString(10586) & vbNewLine & funLoadResString(10587)
    
    Case enuMsg.NoCredit
        lngButtons = vbOKOnly + vbInformation
        strPrompt = funLoadResString(10593)
    
    Case enuMsg.SessionStarted
        lngButtons = vbOKOnly + vbInformation
        strPrompt = funLoadResString(10594) & strText
    
    Case enuMsg.AskCanLockAllWorkstations
        lngButtons = vbYesNo + vbDefaultButton2 + vbExclamation + vbMsgBoxSetForeground
        strPrompt = funLoadResString(10618) & vbNewLine & funLoadResString(10270)
    
    Case enuMsg.AskCanCloseAllSessions
        lngButtons = vbYesNo + vbDefaultButton2 + vbExclamation + vbMsgBoxSetForeground
        strPrompt = funLoadResString(10619)
    
    Case enuMsg.NoElementSelected
        lngButtons = vbOKOnly + vbExclamation
        strPrompt = funLoadResString(10675)
    
    Case enuMsg.FileNotFound
        lngButtons = vbOKOnly + vbExclamation
        strPrompt = funLoadResString(10681) & strText & funLoadResString(10682)
    
    Case enuMsg.CardDisabled
        lngButtons = vbOKOnly + vbInformation
        strPrompt = funLoadResString(10711) & vbNewLine & funLoadResString(10274)
    
    Case enuMsg.AskCanDeactivateCard
        lngButtons = vbYesNo + vbQuestion + vbDefaultButton2
        strPrompt = funLoadResString(10712) & strText & funLoadResString(10713)
    
    Case enuMsg.AskCanActivateCard
        lngButtons = vbYesNo + vbQuestion + vbDefaultButton2
        strPrompt = funLoadResString(10714) & strText & funLoadResString(10713)
    
    Case enuMsg.URLNotFound
        lngButtons = vbOKOnly + vbExclamation
        strPrompt = funLoadResString(10715)
    
    Case enuMsg.DBUpdrageWarning
        lngButtons = vbOKOnly + vbExclamation
        strPrompt = funLoadResString(10720) & vbNewLine & funLoadResString(10721)
    
    Case enuMsg.AskCanRestoreDB
        lngButtons = vbYesNo + vbExclamation
        strPrompt = funLoadResString(10723) & vbNewLine & vbNewLine & funLoadResString(10724) & strText & funLoadResString(10725)
    
    Case enuMsg.AddUserError
        lngButtons = vbOKOnly + vbExclamation
        strPrompt = funLoadResString(10771)
    
    Case enuMsg.AskCanDeleteUser
        lngButtons = vbYesNo + vbExclamation + vbDefaultButton2
        strPrompt = funLoadResString(10772)
    
    Case enuMsg.AskCanLogoff
        lngButtons = vbYesNo + vbQuestion + vbDefaultButton1
        strPrompt = funLoadResString(10781)
    
    Case enuMsg.CancelWizard
        lngButtons = vbYesNo + vbQuestion + vbMsgBoxSetForeground + vbDefaultButton2
        strPrompt = funLoadResString(10022)
    
    Case enuMsg.AskCanBootWorkstations
        lngButtons = vbYesNo + vbQuestion + vbMsgBoxSetForeground + vbDefaultButton1
        strPrompt = funLoadResString(10795)
    
    Case enuMsg.AskCanDeactivateClient
        lngButtons = vbYesNo + vbQuestion + vbDefaultButton2
        strPrompt = funLoadResString(10815, CStr(strText))
    
    Case enuMsg.AskCanActivateClient
        lngButtons = vbYesNo + vbQuestion + vbDefaultButton2
        strPrompt = funLoadResString(10816, CStr(strText))
    
    Case enuMsg.ClientNotSelected
        lngButtons = vbOKOnly + vbExclamation
        strPrompt = funLoadResString(10817)
    
    Case enuMsg.ClientBonusReached
        lngButtons = vbOKOnly + vbInformation
        strPrompt = funLoadResString(10820, CStr(strText), CStr(strMoreText))
    
    Case enuMsg.ClientInUseNoModify
        lngButtons = vbOKOnly + vbInformation
        strPrompt = funLoadResString(10821)
    
    Case enuMsg.AskCanDeleteRechargeFromClient
        lngButtons = vbYesNo + vbQuestion
        strPrompt = funLoadResString(10823, CStr(strText), CStr(strMoreText))
    
    Case enuMsg.ClientNotFound
        lngButtons = vbOKOnly + vbExclamation
        strPrompt = funLoadResString(10827)
    
    Case enuMsg.AddClientError
        lngButtons = vbOKOnly + vbExclamation
        strPrompt = funLoadResString(10828, CStr(strText))
    
    Case enuMsg.ClientDisabled
        lngButtons = vbOKOnly + vbInformation
        strPrompt = funLoadResString(10830)
    
    Case enuMsg.AskCanPurgeClients
        lngButtons = vbYesNo + vbExclamation
        strPrompt = funLoadResString(10847, format(DateAdd("d", -guServOpts.uCards.intCardPurgeDelta * guServOpts.uConfig.intNumDaysInMonth, Date), "Long Date"))
    
    Case enuMsg.AskCanDeleteClient
        lngButtons = vbYesNo + vbExclamation + vbDefaultButton2
        strPrompt = funLoadResString(10874, CStr(strText))
    
    Case enuMsg.AskCanGenerateCards
        lngButtons = vbOKCancel + vbQuestion + vbDefaultButton2
        strPrompt = funLoadResString(10939, CStr(strText))
    
    Case enuMsg.CardTemplateNOK
        lngButtons = vbOKOnly + vbExclamation
        strPrompt = funLoadResString(10945, CARD_GEN_ID)
    
    Case enuMsg.AskCanEnableCard
        lngButtons = vbOKCancel + vbQuestion + vbDefaultButton1
        strPrompt = funLoadResString(10948, CStr(strText))
        
    Case enuMsg.AskCanViewGenCards
        lngButtons = vbOKCancel + vbQuestion + vbDefaultButton1
        strPrompt = funLoadResString(10951, CStr(strText))
        
    Case enuMsg.AskCanPrint
        lngButtons = vbYesNo + vbQuestion + vbDefaultButton2
        strPrompt = funLoadResString(10967, CStr(strText))
    

    
    End Select
    
    If strMsgTitle = "" Then
        strMsgTitle = App.Title
    Else
        strMsgTitle = App.Title & " - " & strMsgTitle
    End If
    If BETA_VERSION Then strMsgTitle = strMsgTitle & " - Msg# " & intMsgType
    
    funMsg = MsgBox(strPrompt, lngButtons, strMsgTitle)
    Exit Function

funMsgError:
    Call funErrorManager("modMessageManager.funMsgError")
End Function

