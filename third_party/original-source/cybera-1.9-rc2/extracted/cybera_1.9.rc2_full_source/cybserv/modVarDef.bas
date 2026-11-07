Attribute VB_Name = "modVarDef"
' _____________________________________________________________________________
'|________________________________ VAR DECLARATION ____________________________|
'|  Name:          modVarDef.bas                                               |
'|  Type:          VB6 Module                                                  |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
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

' $Id: modVarDef.bas 15 2006-01-01 16:12:35Z sam $

Option Explicit


'=== ===========================================================================

'Public Const MAX_WIN_WIDTH = 662
'
'Public Const MAX_WIN_HEIGHT = 455

Public Const MAX_WIN_WIDTH = 782

Public Const MAX_WIN_HEIGHT = 500


'Public guLang As New Collection

Public guServOpts As typCybAdminServ

Public guWorkstations As New Collection

Public guCards As New Collection

Public guCardsPending As New Collection

Public guClients As New Collection

Public guServices As New Collection

Public guMacAddresses As New Collection

Public gwspCybAdmin As DAO.Workspace

Public gdbsCybAdmin As DAO.Database

Public grstConfig As DAO.Recordset

Public grstEvents As DAO.Recordset
Public grstEventsQry As DAO.Recordset
Public gtdfEvents As DAO.TableDef

Public grstCards As DAO.Recordset
Public gtdfCards As DAO.TableDef

Public grstCardsPending As DAO.Recordset
Public gtdfCardsPending As DAO.TableDef

Public grstClients As DAO.Recordset
Public gtdfClients As DAO.TableDef

Public grstServices As DAO.Recordset
Public gtdfServices As DAO.TableDef

Public grstRecharges As DAO.Recordset
Public grstRechargesQry As DAO.Recordset
Public gtdfRecharges As DAO.TableDef

Public grstJournSessions As DAO.Recordset
Public grstJournSessionsQry As DAO.Recordset
Public gtdfJournSessions As DAO.TableDef

Public grstJournIncome As DAO.Recordset
Public grstJournIncomeQry As DAO.Recordset
Public gtdfJournIncome As DAO.TableDef

Public grstUrls As DAO.Recordset
Public grstUrlsQry As DAO.Recordset
Public gtdfUrls As DAO.TableDef

Public grstUsers As DAO.Recordset
Public grstUsersQry As DAO.Recordset
Public gtdfUsers As DAO.TableDef

Public guSelectedWorkstation As New clsWorkstation

Public guSelectedCard As New clsCard

Public guSelectedClient As New clsClient

Public guStatusBarMessages As New Collection

Public gblnDBOK As Boolean

Public gblnNETOK As Boolean

Public guStatsDay() As typDayStats
Public guStatsTimes() As typTimeStats

Public gblnfrmRecharge As Boolean
Public gblnfrmCardHistory As Boolean
Public gblnfrmSetupWizard As Boolean
Public gblnfrmLogon As Boolean
Public gblnfrmCardMassAdd As Boolean
Public gblnfrmOptions As Boolean
Public gblnfrmMain As Boolean

Public glngHistorySelectedURLIndex As Long

Public gblnWorkstationConnexionInProgress As Boolean


