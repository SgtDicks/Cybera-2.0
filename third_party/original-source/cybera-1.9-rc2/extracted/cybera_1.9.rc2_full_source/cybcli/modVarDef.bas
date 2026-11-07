Attribute VB_Name = "modVarDef"
' _____________________________________________________________________________
'|________________________________ VAR DECLARATION ____________________________|
'|  Name:          modVarDef.bas                                               |
'|  Type:          VB6 Module                                                  |
'|  Project:       Cybera Workstation (http://cybera.sf.net)                   |
'|  Description:   Contains variabe and Structure declarations                 |
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

' $Id: modVarDef.bas 26 2006-02-01 13:41:15Z sam $

Option Explicit

'=== ===========================================================================

Public guCliOpts As typCybAdminCli

Public guUser As typUsers

Public guDefaultRecharges() As typRechargeType

Public gblnNoTopmost As Boolean

Public gblnNoKillApp As Boolean

Public gblnNoLockdown As Boolean

Public gblnCnxOK As Boolean

Public gdteExpireryTime As Date

Public gblnNoStatusWindows As Boolean

Public gblnNoCloseSession As Boolean

Public gblnfrmOptions As Boolean
Public gblnfrmAbout As Boolean
Public gblnfrmBlock As Boolean
Public gblnfrmMsgBox As Boolean
Public gblnfrmSmallStatus As Boolean
Public gblnfrmMicroStatus As Boolean
Public gblnfrmStatus As Boolean
Public gblnfrmPasswd As Boolean
Public gblnfrmPasswdChange As Boolean
Public gblnfrmPasswdClose As Boolean
Public gblnfrmPasswdLogOff As Boolean
Public gblnfrmSetupWizard As Boolean

Public guCurrentURLs As New Collection

