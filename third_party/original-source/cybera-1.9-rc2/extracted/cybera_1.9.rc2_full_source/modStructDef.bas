Attribute VB_Name = "modGlobalStructDef"
' _____________________________________________________________________________
'|________________________________ STRUCTURE DEFS _____________________________|
'|  Name:          modStructDef.bas                                            |
'|  Type:          VB6 Module                                                  |
'|  Project:       Cybera Server & Workstation (http://cybera.sf.net)          |
'|  Description:   Contains structure declarations.                            |
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

' $Id: modStructDef.bas 22 2006-01-13 18:04:34Z sam $

Option Explicit


'=== Enumeration defs ==========================================================

'--- General App Enumeratons ---------------------------------------------------
Public Enum enuSessionType
    Unknown = 0
    Selfservice = 1
    Card = 2
    Free = 3
    Client = 4
    SelfServicePrePaid = 5
End Enum

Public Enum enuIncomeType
    Selfservice = 1
    Card = 2
    Other = 3
    Client = 4
    SelfServicePrePaid = 5
End Enum

Public Enum enuWksnState
    NotPresent = 0
    NotActive = 1
    Ready = 2
    Blocked_Deactivated = 3
    Blocked_SessionEnd = 4
    UsedDisconnected = 5
    UsedConnected = 6
    UsedFree = 7
    Blocked_NoCredit = 8
    Blocked_Paused = 9
    Blocked_Locked = 10
    Blocked_follow = 11
End Enum

Public Enum enuWksnMode
    Disconnected = 0
    Connected = 1
End Enum

Public Enum enuWksnSize
    WksnSizeNormal = 1
    WksnSizeSmall = 2
End Enum

Public Enum enuCustomerType
    Card = 1
    Client = 2
    SelfServicePrePaid = 3
End Enum

Public Enum enuOpenSessionReply
    OK = 0
    NOK_NoCard = 1
    NOK_CardDeactivated = 2
    NOK_CardBalanceEmpty = 3
    NOK_CardInUse = 4
    NOK_NoClient = 5
    NOK_ClientDeactivated = 6
    NOK_ClientBalanceEmpty = 7
    NOK_ClientInUse = 8
    NOK_SelfServicePrePaidBalanceEmpty = 9
End Enum

Public Enum enuSrvCxn
    Disconnected = 0
    Connected = 1
End Enum

Public Enum enuTimeFormat
    HH_NN_SS = 0
    Hh_Nmn = 1
    Hh_Nmn_Ssec = 2
    Hh_Nmn_Ssec_dyn = 3
End Enum

Public Enum enuServiceStockMode
    Unmanaged = 0
    Managed = 1
End Enum

Public Enum enuWksnDisableReason
    VersionNOK = 0
    InternalError = 1
End Enum

Public Enum enuMessageBoxType
    OKOnly = 0
    YesNo = 1
End Enum

Public Enum enuMessageBoxDefaultButton
    Button1 = 1
    Button2 = 2
End Enum

Public Enum enuEventLevels
    Production = 0
    Verbose = 5
    All = 10
End Enum

Public Enum enuEvents
    AppStart = 0
    AppStop = 1
    AppError = 2
    AppDBNetworkBackup = 3
    AppLogon = 4
    AppLogoff = 5
    SessionOpen = 10
    SessionClose = 11
    SessionPay = 12
    SessionNoCredit = 13
    SessionPaused = 14
    SessionUnPaused = 15
    SessionConnected = 16
    SessionDisconnected = 17
    SessionFree = 18
    SessionResume = 19
    CardRecharge = 20
    CardNew = 21
    CardDelete = 22
    CardUnRecharge = 23
    ServicePay = 30
    InternetConnected = 40
    InternetDiconnected = 41
    WorkstationNotPresent = 50
    WorkstationNotActive = 51
    WorkstationReady = 52
    WorkstationDeactivated = 53
    WorkstationUsageAttempt = 54
    ClientRecharge = 60
    ClientNew = 61
    ClientDelete = 62
    ClientUnRecharge = 63
    SelfServiceRecharge = 70
End Enum

Public Enum enuFileActions
    LoadIntro = 0
    LoadBlock = 1
    LoadNoCredit = 2
    LoadSessionEnd = 3
    LoadSessionPaused = 4
    LoadLogo = 5
    LoadFollowImage = 6
    LoadBackground = 7
    SaveDB = 8
    SessionBatchStart = 9
    SessionBatchEnd = 10
End Enum

Public Enum enuSerializedSession
    Ended = 0
    Ongoing = 1
End Enum

Public Enum enuMsgType
    Default = 0
    User = 1
End Enum

Public Enum enuRechargeValidityType
    Hours = 0
    Days = 1
    Weeks = 2
    Months = 3
    Years = 4
    Unlimited = 5
End Enum

Public Enum enuPostpaidSsMode
    RechargeBased = 1
    TimeBased = 2
End Enum

Public Enum enuPostpaidSsSubMode
    RechAutoOptim = 1
    RechSimple = 2
End Enum

Public Enum enuRechargeOrder
    PrimaryKey = 0
    RechargesIndex = 1
    orderByTime = 2
    orderByPrice = 3
    orderByValidity = 4
    orderByType = 5
End Enum


'=== Structure defs ============================================================

'--- General options ------------------------------------------------------------
Type typPos
    intTop As Integer                       '-Window position.
    intLeft As Integer                      '
    intHeight As Integer                    '
    intWidth As Integer
    intWinState As Integer '
End Type

Type typCnx
    blnAutoConnect As Boolean
    strServerName As String
    strServerMac As String
    strServerIP As String
    lngPort As Long
    intCnxState As Integer
    intTimeOut As Integer
End Type

Private Type typTimeRatios
    intDisconCon As Integer
    intOffPeak As Integer
    dteOffPeakStart As Date
    dteOffPeakStop As Date
End Type

Private Type typWorkstationMsg
    intMsgCard As Integer
    intMsgCardAlert1 As Integer
    intMsgCardAlert2 As Integer
    intMsgSelfService As Integer
    sngMsgSelfServiceBalance As Single
    lngMsgSelfServiceBalDeltaRatio As Long
End Type

Type typConfigWizard
    intSystem As Integer
    intUser As Integer
End Type


'--- Workstation options ----------------------------------------------------------

Private Type typPostPaidSs
    intMode As Integer
    dteBaseTimeIncrement As Date
    dteMinBilledTime As Date
    intSubMode As Integer
    dteEndTimeIncrement As Date
    dteGraceTime As Date
End Type

Private Type typKilledApps
    intForbidIE As Integer
    intForbidNetscape As Integer
    intForbidMozilla As Integer
    intForbidOE As Integer
    strKilledApps() As String
End Type

Private Type typCliFlags
    blnWaitingForReply As Boolean
    intServerResponse As Integer
    dteTimer As Date
    blnWaitingForServerDate As Boolean
    dteServerDelta As Date
    blnNoInetCnxNotified As Boolean
    intNumCnxAttempts As Integer
    intNumDeactAttempts As Integer
    blnServerInitDone As Boolean
End Type

Private Type typPaths
    strIntroPath As String
    strBlockedPath As String
    strNoCreditPath As String
    strSessionEndPath As String
    strSessionPausedPath As String
    strLogoPath As String
    strLogfilePath As String
End Type

Private Type typSessionClean
    intAlignDesktopIcons As Integer
    intDeleteDocHistory As Integer
    intRestoreBackground As Integer
    strBackgroundPath As String
    intRestoreIEHomePage As Integer
    intRestoreIESettings As Integer
    strIEHomePagePath As String
    intRestoreIMs As Integer
End Type

Private Type typSessionBatch
    intEnableStart As Integer
    intEnableEnd As Integer
    strStartPath As String
    strEndPath As String
    intWindowStyle As Integer
End Type

Public Type typCybAdminCli
    uWinPos As typPos
    uFlags As typCliFlags
    uTimeRatios As typTimeRatios
    uPaths As typPaths
    uWksnMsg As typWorkstationMsg
    uSessionCleanup As typSessionClean
    uSessionBatch As typSessionBatch
    uKilledApps As typKilledApps
    uConfigWizard As typConfigWizard
    uConnectedManager As clsUser
    uPostPaidSs As typPostPaidSs
    strWksnName As String
    strComputerName As String
    strComputerMacAddress As String
    strUserName As String
    strPasswd As String
    uServer As typCnx
    blnStartActive As Boolean
    blnWksnTypeSndAln As Boolean
    strKilledApps() As String
    strMonSymbol As String
    strMonFormat As String
    strAppConfigPath As String
    strAppConfigINIPath As String
    lngInterfaceLangID As Long
    strInterfaceLang As String
    strInterfaceLangDefault As String
    intInterfaceLangCharset As Integer
    intInterfaceLangCharsetDefault As Integer
    intStartAtLogon As Integer
    intDefaultState As Integer
    intDefaultMode As Integer
    intDefaultUsage As Integer
    intAllowLockdown As Integer
    intSerializerFile As Integer
    intHideCardNum As Integer
    intMsgLogoFilesType As Integer
    intStartDelay As Integer
    blnReadyForServerScreen As Boolean
End Type


'--- Server options -------------------------------------------------------------
Private Type typCnxTests
    intCnxTesting As Integer
    strTestServerName As String
    strTestServerPort As String
    intWorkstationNotification As Integer
End Type

Private Type typServFlags
    blnInetCnx As Boolean
    blnConsoleActive As Boolean
    dteOnTime As Date
    intWksnSize As enuWksnSize
End Type

Private Type typDBconfig
    strDBPath As String
    strDBBackupPath As String
    strDBPasswd As String
    intDBBackup As Integer
    intDBMaxBackups As Integer
    intDBBackupSendToWorkstation As Integer
    intDBCompact As Integer
End Type

Private Type typStats
    strStatsTimeInt() As String
    intStatsGraphStartTime As Integer
    intStatsGraphStopTime As Integer
End Type

Private Type typCards
    strCardTypes() As String
    intCardListUpdateInterval As Integer
    intCardPurgeDelta As Integer
    intAutoCardPurge As Integer
    intBonusActive As Integer
    sngBonusThreshold As Single
    dteBonusTime As Date
End Type

Private Type typClients
    intClientListUpdateInterval As Integer
    intClientPurgeDelta As Integer
    intAutoClientPurge As Integer
    intBonusActive As Integer
    sngBonusThreshold As Single
    dteBonusTime As Date
    intEnableExpiry As Integer
End Type

Private Type typLockdown
    intNoClose As Integer
    intNoAdminTools As Integer
    intNoCPL As Integer
    intLockDesktop As Integer
    intNoFileSystemAccess As Integer
    intNoFileSystemAccessType As Integer
    lngLockedDrives As Long
    lngLockedDrivesTmp As Long
    intNoCommonGroup As Integer
    intLockIE As Integer
End Type

Private Type typWorkstations
    uLocks As typLockdown
    uWksnMsg As typWorkstationMsg
    uPaths As typPaths
    intStartForceWorkstationsConnected As Integer
    intActivateSelfServicePostPaid As Integer
    intActivateSelfServicePrePaid As Integer
    intActivateCards As Integer
    intActivateClients As Integer
    intUsageDefault As Integer
    lngMaxNumWorkstations As Long
    lngNumWorkstations As Long
    intFollowInterval As Integer
    intModifyLockdown As Integer
    intModifyMsgLogoFiles As Integer
    intHideCardNum As Integer
    intEnablePause As Integer
    intLogURLs As Integer
    intLoggedURLTimeSpan As Integer
    strMacAddresses() As String
    intDefaultState As Integer
    intDefaultMode As Integer
End Type

Private Type typConfig
    strPasswd As String
    strMonSymbol As String
    strMonFormat As String
    intNumDaysInMonth As Integer
    strLastExportPath As String
    lngInterfaceLangID As Long
    strInterfaceLang As String
    intInterfaceLangCharset As Integer
    strAppConfigPath As String
    strAppConfigINIPath As String
    strDefaultPrinterName As String
End Type

Public Type typCybAdminServ
    uWinPos As typPos
    uDB As typDBconfig
    uFlags As typServFlags
    uTimeRatios As typTimeRatios
    uCnxTest As typCnxTests
    uStats As typStats
    uCards As typCards
    uClients As typClients
    uPostPaidSs As typPostPaidSs
    uServer As typCnx
    uWorkstations As typWorkstations
    uConfig As typConfig
    uKilledApps As typKilledApps
    uSessionCleanup As typSessionClean
    uSessionBatch As typSessionBatch
    uConfigWizard As typConfigWizard
    uConnectedUser As clsUser
    lngSelectedCardIndex As Long
    lngSelectedCardPendingIndex As Long
    lngSelectedClientIndex As Long
    lngSelectedServiceIndex As Long
    lngSelectedURLIndex As Long
    intNumVisRows As Integer
End Type

'--- Tarif options --------------------------------------------------------------
Public Type typRechargeType
    dteTimeCredit As Date
    sngPrice As Single
End Type

'--- User options --------------------------------------------------------------
Public Type typUsers
    strID As String
    strPassword As String
    lngSessionID As Long
    sngBalance As Single
    dteStartTime As Date
    dteBlockedTime As Date
    dteElapsedTime As Date
    dteElapsedTimeAparent As Date
    dteTimeConnected As Date
    dteTimeConnectedOffPeak As Date
    dteTimeDisconnected As Date
    dteRemainingTime As Date
    dteRemainingTimeAparent As Date
    dteInitRemainingTime As Date
    dteInitRemainingTimeSerialized As Date
    dteCustomerExpiryDate As Date
    intSessionType As Integer
    intWksnState As Integer
    intWksnStateOld As Integer
    blnUnlimitedPlay As Boolean
End Type

'--- Event querys --------------------------------------------------------------
Public Type typTimeElement
    dteStart As Date
    dteTotal As Date
End Type

Public Type typDayStats
    dteDate As Date
    uTimeOpen As typTimeElement
    uTimeInetCnx As typTimeElement
    uTimeWs(MAX_NUM_WORKSTATIONS) As typTimeElement
    dteTimeWsTotal As Date
    lngSessionNum As Long
    dteSessionTotal As Date
End Type

'--- Session querys --------------------------------------------------------------
Public Type typTimeStats
    dteDate As Date
    lngSessionNum As Long
    lngSessionCount As Long
    lngTimeInt() As Long
End Type

