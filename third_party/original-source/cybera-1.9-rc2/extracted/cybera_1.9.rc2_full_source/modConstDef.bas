Attribute VB_Name = "modConstDef"
' _____________________________________________________________________________
'|________________________________ CONSTANT DEFS ______________________________|
'|  Name:          modConstDef.bas                                             |
'|  Type:          VB6 Module                                                  |
'|  Project:       Cybera Server & Workstation (http://cybera.sf.net)          |
'|  Description:   Contains Constants declarations.                            |
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

' $Id: modConstDef.bas 41 2006-02-21 18:14:29Z sam $

Option Explicit


'=== Constant deffinitions =====================================================

'--- General App constants -----------------------------------------------------
Public Const BETA_VERSION = False              'Display beta in the "About" window
Public Const RC_VERSION = "rc2"                'Release candidate or beta
Public Const BUILD_MONTH_YEAR = "02/2006"       'Month and year of present build
Public Const BUILD_YEAR = "2004-2006"
Public Const DB_MIN_VERSION = 37
Public Const SERVER_MIN_VERSION = 1009000
Public Const WORKSTATION_MIN_VERSION = 1009000
Public Const DEFAULT_TCP_PORT = 4646
Public Const DEFAULT_UDP_PORT = 4646
Public Const DEFAULT_UDP_PORT_SERVER = 4647
Public Const DEFAULT_UDP_PORT_WORKSTATION = 4648
Public Const BROADCAST_ADDRESS = "255.255.255.255"
Public Const EMPTY_ADDRESS = "0.0.0.0"

Public Const CYBERA_ERROR = 4646

'--- Personal contact information constants ------------------------------------
Public Const MY_URL = "http://cybera.sourceforge.net/"
Public Const MY_DONATE_URL = "https://sourceforge.net/donate/index.php?group_id=121074"
Public Const MY_NAME = "Samuel Monsarrat"
Public Const MY_EMAIL = "smonsarr@users.sourceforge.net"
Public Const MY_COPYRIGHT = "Copyright ©"

Public Const DO_NOT_SHUTDOWN = -1         '-Form_Unload management constant.

Public Const NULL_DATE = #12:00:00 AM#
Public Const SESSION_EXPIRERY_TIME = #12:05:00 AM# ' Used in wksn: delay before a blocked (no credit) session will be terminated
Public Const ACTIVATION_DELAY_TIME = #12:02:00 AM# ' Used in wksn: delay wksn becomes available after an incorrect auth.
Public Const NUM_DAYS_IN_YEAR = 365
Public Const LANG_EN = 2057
Public Const LANG_FR = 1036
Public Const LANG_DE = 1031
Public Const LANG_IT = 1040
Public Const LANG_ES = 3082
Public Const LANG_EL = 1032
Public Const LANG_BR = 1046
Public Const LANG_HU = 1038
Public Const LANG_TR = 1055

Public Const LANG_RES_EN = "_en"
Public Const LANG_RES_FR = "_fr"
Public Const LANG_RES_DE = "_de"
Public Const LANG_RES_IT = "_it"
Public Const LANG_RES_ES = "_es"
Public Const LANG_RES_EL = "_el"
Public Const LANG_RES_BR = "_br"
Public Const LANG_RES_HU = "_hu"
Public Const LANG_RES_TR = "_tr"

Public Const DELAULT_MAX_NUM_WORKSTATIONS = 50
Public Const MAX_NUM_WORKSTATIONS = 100
Public Const DELAULT_MAX_NUM_DB_BACKUPS = 4
Public Const MAX_NUM_DB_BACKUPS = 20

Public Const TIME_TOKEN = "#"
Public Const CARD_TYPE_TOKEN = "#"
Public Const FILENAME_TOKEN = "*"
Public Const SERIALIZED_SESSION_TOKEN = "#"

Public Const PASSWORD_CHAR = "¤"

Public Const CSV_SEPARATOR = ","

Public Const APP_NAME_IE = "iexplore.exe"
Public Const APP_NAME_NETSCAPE = "netscp.exe"
Public Const APP_NAME_MOZILLA = "mozilla.exe"
Public Const APP_NAME_FIREFOX = "firefox.exe"
Public Const APP_NAME_THUNDERBIRD = "thunderbird.exe"
Public Const APP_NAME_OE = "msimn.exe"

Public Const APP_DIR_GAIM = ".gaim"
Public Const APP_DIR_YAHOO = "Yahoo! Messenger"

Public Const FILENAME_PREFIX = "\"
Public Const FILENAME_PREFIX_EN = "\en_"
Public Const FILENAME_PREFIX_FR = "\fr_"
Public Const FILENAME_PREFIX_DE = "\de_"
Public Const FILENAME_PREFIX_IT = "\it_"
Public Const FILENAME_PREFIX_ES = "\es_"
Public Const FILENAME_PREFIX_EL = "\el_"
Public Const FILENAME_PREFIX_BR = "\br_"
Public Const FILENAME_PREFIX_HU = "\hu_"
Public Const FILENAME_PREFIX_TR = "\tr_"

Public Const FILENAME_SUFFIX_BD_NET_BACKUP = "_network_backup_location.txt"

Public Const FILENAME_SCREENCAP_BMP = "\screen_cap.bmp"
Public Const FILENAME_SCREENCAP_JPG = "\screen_cap.jpg"
Public Const FILENAME_RECEIVED = "\received_file.tmp"
Public Const FILENAME_SERIALIZED = "\serialized_session.txt"
Public Const FILENAME_LOGO = "logo.gif"
Public Const FILENAME_INTRO = "intro.rtf"
Public Const FILENAME_BLOCKED = "blocked.rtf"
Public Const FILENAME_NOCREDIT = "nocredit.rtf"
Public Const FILENAME_SESSIONEND = "sessionend.rtf"
Public Const FILENAME_SESSIONPAUSED = "sessionpaused.rtf"
Public Const FILENAME_BACKGROUND = "\background_image"
Public Const FILENAME_DB_BACKUP = "\server_db_backup.mdb"
Public Const FILENAME_FOLLOW_IMAGE = "\follow_image.jpg"

Public Const SESS_CMD_ID = "</cmdUserID>"
Public Const SESS_CMD_SERVER = "</cmdServerName>"
Public Const SESS_CMD_WKSN = "</cmdWksnName>"

Public Const FILENAME_SESSION_BATCH_START = "\session_batch_start.bat" ' " & SESS_CMD_ID & " " & SESS_CMD_SERVER & " " & SESS_CMD_WKSN
Public Const FILENAME_SESSION_BATCH_END = "\session_batch_end.bat" ' " & SESS_CMD_ID & " " & SESS_CMD_SERVER & " " & SESS_CMD_WKSN

Public Const FILENAME_CARD_TEMPLATE = "card_template.html"
Public Const FILENAME_CARD_TEMPLATE_FOLDER = "\Generated cards\"

Public Const MAX_NUM_CNX_ATTEMPTS = 3

Public Const SCREENCAP_COMPRESSION_RATIO = 90

Public Const MIN_SIZE_TEMP_FILE_STORAGE = 50000

Public Const MAX_SIZE_DB_INFO = 30
Public Const MAX_SIZE_DB_WARN = 60
Public Const MAX_SIZE_DB_CRIT = 120

Public Const URL_EMPTY_IE = "about:blank"

Public Const CONFIG_WIZARD_CLI_SYSTEM = 1
Public Const CONFIG_WIZARD_CLI_USER = 1
Public Const CONFIG_WIZARD_SERV_SYSTEM = 1
Public Const CONFIG_WIZARD_SERV_USER = 1

Public Const ADMIN_LOGON = "admin"

Public Const CLIENT_LOGIN_FORBIDDEN_CHARS = "\?/:*<>|"" '#~¨^`%$"


Public Const RES_NEW_LINE = "</resNewLine>"
Public Const RES_TAB = "</resTab>"
Public Const RES_INSERT_1 = "</resInsert1>"
Public Const RES_INSERT_2 = "</resInsert2>"
Public Const RES_INSERT_3 = "</resInsert3>"
Public Const RES_INSERT_4 = "</resInsert4>"

Public Const CURRENCY_SYMBOL_INSERT = "</currSymbol>"


Public Const WKSN_GUI_SIZE_NORMAL = 2415
Public Const WKSN_GUI_SIZE_SMALL = 1185

Public Const WKSN_BOOT_CNX_DELAY = 90000
Public Const WKSN_MIN_CNX_DELAY = 2000
Public Const WKSN_DEFAULT_START_DELAY = 30

Public Const WKSN_LOCK_DRIVES_ALL = 67108863
Public Const WKSN_LOCK_DRIVES_NONE = 0
Public Const WKSN_LOCK_DRIVES_C_ONLY = 4

Public Const HIDDEN_LOGOFF_TIMEOUT = 5

Public Const CARD_GEN_ID = "</crdNumber>"
Public Const CARD_GEN_TITLE = "</crdPageTitle>"
Public Const CARD_GEN_HEADER = "</crdPageHeader>"
Public Const CARD_GEN_FOOTER = "</crdPageFooter>"

Public Const NUM_ROWS_STATS_PRINT = 50

Public Const PAPER_LONG_SIDE_A4 = 29.7
Public Const PAPER_SHORT_SIDE_A4 = 21
Public Const PAPER_LONG_SIDE_LR = 27.94
Public Const PAPER_SHORT_SIDE_LR = 21.59
Public Const TWIPS_PER_CM = 567
Public Const PRINT_MARGIN = 1.5
Public Const PRINT_MAX_PAGES_BEFORE_WARNING = 10

Public Const DEFAULT_PPSS_TIME_INCR = #12:05:00 AM#
Public Const MIN_PPSS_TIME_INCR = #12:00:01 AM#

Public Const DB_UPGRADE_SUFFIX = ".tmp_upgd.mdb"


Public gstrHelpers() As String


'*******************************************************************************
Public Sub subDefineConstants()
On Error Resume Next

End Sub


'*******************************************************************************
Public Sub subDefineHelpers()
Dim strLangPacks() As String
Dim i As Integer
On Error Resume Next

    strLangPacks = funFindLangPacks
    ReDim gstrHelpers(0)
    gstrHelpers(0) = funLoadResString(10039)
    Call subAddHelper("Anne Berthet")
    Call subAddHelper("Karen De Cokere")
    Call subAddHelper("Fabian Goreth")
    Call subAddHelper("Daniel Lankoandé")
    Call subAddHelper("Biba Nadjoari")
    Call subAddHelper("Bipin Patel")
    Call subAddHelper("Teninyala Tankoano")
    Call subAddHelper("Natama Thiombiano")
    Call subAddHelper("Vasilis Charamidis")
    Call subAddHelper("")
    Call subAddHelper(funLoadResString(10044))
    For i = 0 To UBound(strLangPacks, 2)
        If strLangPacks(2, i) <> MY_NAME Then Call subAddHelper(strLangPacks(4, i) & " - " & strLangPacks(2, i))
    Next i
End Sub


'*******************************************************************************
Private Sub subAddHelper(strHelperLine As String)
On Error Resume Next

    ReDim Preserve gstrHelpers(UBound(gstrHelpers) + 1)
    gstrHelpers(UBound(gstrHelpers)) = strHelperLine
End Sub


