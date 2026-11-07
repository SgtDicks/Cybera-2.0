Attribute VB_Name = "modCnxMsg"
' _____________________________________________________________________________
'|________________________________ SERVER-WORKSTATION MESSAGE DEFS ____________|
'|  Name:          modCnxMsg.bas                                               |
'|  Type:          VB6 Module                                                  |
'|  Project:       Cybera Server & Workstation (http://cybera.sf.net)          |
'|  Description:   Contains the Server - Workstation message headers.          |
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

' $Id: modCnxMsg.bas 15 2006-01-01 16:12:35Z sam $

Option Explicit


' "%$&#################"
' "%$&#################"

Public Const CNX_AUTO_CONFIG_HANDSHAKE = "%$&#CYBERA@!*£µ¤?§;#"
Public Const CNX_AUTO_CONFIG_HANDSHAKE_WKSN = "%$&#CYBERA@!*£µ¤?§;#"
Public Const CNX_AUTO_CONFIG_HANDSHAKE_SERV = "%$&#CYBERA{[°£!^}%~#"


Public Const CNX_MSG_END = "%$&######MessageEnd#"
Public Const CNX_PAYLOAD_TOKEN = "%$&####PayloadToken#"
Public Const CNX_APP_VERSION = "%$&######AppVersion#"
Public Const CNX_WORKSTATION_NAME = "%$&#WorkstationName#"
Public Const CNX_COMPUTER_NAME = "%$&####ComputerName#"
Public Const CNX_WORKSTATION_STATE = "%$&WorkstationState#"
Public Const CNX_SESSION_TYPE = "%$&#####SessionType#"
Public Const CNX_USER_ID = "%$&##########UserID#"
Public Const CNX_USER_BALANCE = "%$&#####UserBalance#"
'Public Const CNX_START_TIME = "%$&#######StartTime#"
Public Const CNX_ELAPSED_TIME = "%$&#####ElapsedTime#"
Public Const CNX_CONNECTED_TIME = "%$&###ConnectedTime#"
Public Const CNX_CONNECTED_TIME_OFFPEAK = "%$&#CnxtTimeOffPeak#"
Public Const CNX_DISCONNECTED_TIME = "%$&DisconnectedTime#"
Public Const CNX_REMAINING_TIME = "%$&###RemainingTime#"
Public Const CNX_REMAINING_TIME_APARENT = "%$&##RemTimeAparent#"
Public Const CNX_SESSION_OPEN = "%$&#####SessionOpen#"
Public Const CNX_SESSION_CLOSE = "%$&####SessionClose#"
Public Const CNX_SESSION_OK = "%$&#######SessionOK#"
Public Const CNX_SESSION_NOK = "%$&######SessionNOK#"
Public Const CNX_SESSION_URL = "%$&######SessionURL#"
Public Const CNX_SESSION_FREE = "%$&#####SessionFree#"
Public Const CNX_SESSION_ID = "%$&#######SessionID#"
Public Const CNX_SESSION_RESUME = "%$&###SessionResume#"
Public Const CNX_DEFAULT_RECHARGES = "%$&DefaultRecharges#"
Public Const CNX_RATIO_DISCON_CON = "%$&##RatioDisconCon#"
Public Const CNX_CURRENCY_SYMBOL = "%$&##CurrencySymbol#"
Public Const CNX_OFFPEAK_PARAMS = "%$&###OffPeakParams#"
Public Const CNX_CHECK_BALANCE = "%$&CheckCardBalance#"
Public Const CNX_FORCE_UPDATE = "%$&#ForceCardUpdate#"
Public Const CNX_LOCK_WORKSTATION = "%$&########LockWksn#"
Public Const CNX_ACTIVATE_WORKSTATION = "%$&###ActivateWksn#"
Public Const CNX_DEACTIVATE_WORKSTATION = "%$&#DeactivateWksn#"
Public Const CNX_DISABLE_WORKSTATION = "%$&####DisableWksn#"
Public Const CNX_FORCE_CONNECTED = "%$&##ForceConnected#"
Public Const CNX_WORKSTATION_MSG_OPTIONS = "%$&##WksnMsgOptions#"
Public Const CNX_ACTIVATE_SELFSERVICE = "%$&ActivateSelfServ#"
Public Const CNX_ACTIVATE_CARDS = "%$&###ActivateCards#"
Public Const CNX_ACTIVATE_CLIENTS = "%$&#ActivateClients#"
Public Const CNX_SHUTDOWN_WORKSTATION = "%$&####ShutdownWksn#"
Public Const CNX_RECHARGE_CARD = "%$&####RechargeCard#"
Public Const CNX_PAUSE_WORKSTATION = "%$&#######PauseWksn#"
Public Const CNX_PAUSE_LOCK_WORKSTATION = "%$&###PauseLocKWksn#"
Public Const CNX_PAUSE_FOLLOW_WORKSTATION = "%$&#PauseFollowWksn#"
Public Const CNX_RESUME_WORKSTATION = "%$&#####ResumeWksn#"
Public Const CNX_UNPAUSE_WORKSTATION = "%$&####UnPauseWksn#"
Public Const CNX_NOTIFY_WKSN_NO_INETCNX = "%$&#NotifyNoInetCnx#"
Public Const CNX_NOTIFY_WKSN_INETCNX = "%$&###NotifyInetCnx#"
Public Const CNX_FILE_SEND = "%$&########FileSend#"
Public Const CNX_FILE_TEMP = "%$&########FileTemp#"
Public Const CNX_MSG_SEND = "%$&#####MessageSend#"
Public Const CNX_PASSWD_MANGR = "%$&###PasswordMangr#"
Public Const CNX_PASSWD_ADMIN = "%$&###PasswordAdmin#"
Public Const CNX_KILLED_APPS = "%$&######KilledApps#"
Public Const CNX_SESSION_CLEANUP = "%$&##SEssionCleanup#"
Public Const CNX_HIDE_CARD_NUM = "%$&#####HideCardNum#"
Public Const CNX_DB_BACKUP_DTE = "%$&#####DBBackupDte#"
Public Const CNX_ENABLE_PAUSE = "%$&#####EnablePause#"
Public Const CNX_PASSWD_CHANGE = "%$&##PasswordChange#"
Public Const CNX_SERVER_DATE = "%$&######ServerDate#"
Public Const CNX_SESSION_BATCH_START = "%$&##SessBatchStart#"
Public Const CNX_SESSION_BATCH_END = "%$&####SessBatchEnd#"
Public Const CNX_SESSION_BATCH_STYLE = "%$&##SessBatchStyle#"
Public Const CNX_WKSN_DEFAULT_STATE = "%$&WksnDefaultState#"
Public Const CNX_WKSN_DEFAULT_MODE = "%$&#WksnDefaultMode#"
Public Const CNX_DEFAULT_LANG = "%$&#DefaultLanguage#"
Public Const CNX_MSG_LOGO_FILES_TYPE = "%$&MsgLogoFilesType#"
Public Const CNX_CHAT_MSG = "%$&#####ChatMessage#"
Public Const CNX_EMPTY_SERIALIZED_USER = "%$&EmptySerialzdUsr#"
Public Const CNX_READY_FOR_SERVER_SCREEN = "%$&RdyForServScreen#"
Public Const CNX_POSTPAID_MODE = "%$&####PostPaidMode#"
Public Const CNX_WKSN_DEFAULT_USAGE = "%$&####UsageDefault#"

