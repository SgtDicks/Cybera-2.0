Attribute VB_Name = "modFunctionsDB"
' _____________________________________________________________________________
'|________________________________ GENERAL DB FUNCTION MODULE _________________|
'|  Name:          modFunctionsDB.bas                                          |
'|  Type:          VB6 Module                                                  |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
'|  Description:   Contains general Database functions and subs                |
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

' $Id: modFunctionsDB.bas 35 2006-02-08 02:20:12Z sam $

Option Explicit


'*******************************************************************************
Public Function funInitDB() As Boolean
Dim strDBVersion As String
Dim dblDBSize As Double
Dim uMsg As clsStatusBarMessage
Dim idxRecharges As DAO.Index
Dim strStep As String
On Error GoTo subInitDBError

    Call subRestoreDB
    
    dblDBSize = FileLen(guServOpts.uDB.strDBPath) / 1024 ^ 2
    
    Set gwspCybAdmin = DAO.CreateWorkspace("Cyb Admin", "admin", "", dbUseJet)
    Set gdbsCybAdmin = gwspCybAdmin.OpenDatabase(guServOpts.uDB.strDBPath, False, False, ";pwd=" & guServOpts.uDB.strDBPasswd)
    
    Set grstConfig = gdbsCybAdmin.OpenRecordset("config", dbOpenTable)
    Call grstConfig.MoveFirst
    If grstConfig!Version < DB_MIN_VERSION Then
        Call grstConfig.Close
        Call gdbsCybAdmin.Close
        If Dir(guServOpts.uDB.strDBPath & DB_UPGRADE_SUFFIX) <> "" Then Call Kill(guServOpts.uDB.strDBPath & DB_UPGRADE_SUFFIX)
        Call FileCopy(guServOpts.uDB.strDBPath, guServOpts.uDB.strDBPath & DB_UPGRADE_SUFFIX)
        Set gdbsCybAdmin = gwspCybAdmin.OpenDatabase(guServOpts.uDB.strDBPath & DB_UPGRADE_SUFFIX, False, False, ";pwd=" & guServOpts.uDB.strDBPasswd)
        Set grstConfig = gdbsCybAdmin.OpenRecordset("config", dbOpenTable)
        funInitDB = funDBUpgrade
        Call grstConfig.Close
        Call gdbsCybAdmin.Close
        If funInitDB Then
            If Dir(guServOpts.uDB.strDBPath) <> "" Then Call Kill(guServOpts.uDB.strDBPath)
            Name guServOpts.uDB.strDBPath & DB_UPGRADE_SUFFIX As guServOpts.uDB.strDBPath
        End If
        Set gdbsCybAdmin = gwspCybAdmin.OpenDatabase(guServOpts.uDB.strDBPath, False, False, ";pwd=" & guServOpts.uDB.strDBPasswd)
        Set grstConfig = gdbsCybAdmin.OpenRecordset("config", dbOpenTable)
    End If
    
    strStep = "table: events"
    Set gtdfEvents = gdbsCybAdmin!events
    Set grstEvents = gdbsCybAdmin.OpenRecordset("events", dbOpenTable)
    grstEvents.Index = "PrimaryKey"
    
    strStep = "table: cards"
    Set gtdfCards = gdbsCybAdmin!cards
    Set grstCards = gdbsCybAdmin.OpenRecordset("cards", dbOpenTable)
    grstCards.Index = "PrimaryKey"
    
    strStep = "table: cards_pending"
    Set gtdfCardsPending = gdbsCybAdmin!cards_pending
    Set grstCardsPending = gdbsCybAdmin.OpenRecordset("cards_pending", dbOpenTable)
    grstCardsPending.Index = "PrimaryKey"
    
    strStep = "table: clients"
    Set gtdfClients = gdbsCybAdmin!clients
    Set grstClients = gdbsCybAdmin.OpenRecordset("clients", dbOpenTable)
    grstClients.Index = "PrimaryKey"
    
    strStep = "table: recharges"
    Set gtdfRecharges = gdbsCybAdmin!recharges
    Set grstRecharges = gdbsCybAdmin.OpenRecordset("recharges", dbOpenTable)
    grstRecharges.Index = "RechargesIndex"
    
    strStep = "table: services"
    Set gtdfServices = gdbsCybAdmin!recharges
    Set grstServices = gdbsCybAdmin.OpenRecordset("services", dbOpenTable)
    grstServices.Index = "ServicesIndex"
    
    strStep = "table: journ_sessions"
    Set gtdfJournSessions = gdbsCybAdmin!journ_sessions
    Set grstJournSessions = gdbsCybAdmin.OpenRecordset("journ_sessions", dbOpenTable)
    grstJournSessions.Index = "PrimaryKey"
    
    strStep = "table: journ_income"
    Set gtdfJournIncome = gdbsCybAdmin!journ_income
    Set grstJournIncome = gdbsCybAdmin.OpenRecordset("journ_income", dbOpenTable)
    grstJournIncome.Index = "PrimaryKey"
    
    strStep = "table: urls"
    Set gtdfUrls = gdbsCybAdmin!urls
    Set grstUrls = gdbsCybAdmin.OpenRecordset("urls", dbOpenTable)
    grstUrls.Index = "PrimaryKey"
    
    strStep = "table: users"
    Set gtdfUsers = gdbsCybAdmin!Users
    Set grstUsers = gdbsCybAdmin.OpenRecordset("users", dbOpenTable)
    grstUsers.Index = "PrimaryKey"
    
    If dblDBSize > MAX_SIZE_DB_INFO Then Set uMsg = New clsStatusBarMessage
    If dblDBSize > MAX_SIZE_DB_CRIT Then
        uMsg.ImageIndex = enuStatus.cliCritical
        uMsg.Text = funLoadResString(10748, MAX_SIZE_DB_CRIT)
    ElseIf dblDBSize > MAX_SIZE_DB_WARN Then
        uMsg.ImageIndex = enuStatus.cliWarn
        uMsg.Text = funLoadResString(10749, MAX_SIZE_DB_WARN)
    ElseIf dblDBSize > MAX_SIZE_DB_INFO Then
        uMsg.ImageIndex = enuStatus.cliInfo
        uMsg.Text = funLoadResString(10750, MAX_SIZE_DB_INFO)
    End If
    If dblDBSize > MAX_SIZE_DB_INFO Then Call funAddStatusBarMsg(uMsg, "DB_SIZE_WARNING")
    
    funInitDB = True
    Exit Function
    
subInitDBError:
Dim strDBErrorMsg As String
    funInitDB = False
    strDBErrorMsg = Err.Description
    Select Case Err.Number
    Case 3024
        Call funMsg(enuMsg.DBPathError, strDBErrorMsg)
    Case 3031
        Call funMsg(enuMsg.DBPasswordError)
    Case 3011
        If funMsg(enuMsg.DBVersionError) = vbYes Then Call funShutdownApplication
    Case Else
        Call funErrorManager("modFunctionsDB.subInitDBError(" & strStep & ")", False, True)
        If funMsg(enuMsg.GeneralDBError, strDBErrorMsg) = vbYes Then Call funShutdownApplication
    End Select
End Function


'*******************************************************************************
Public Sub subBackupDB()
Dim i As Integer
Dim strOldName As String, strNewName As String, strDBName As String
On Error GoTo subBackupDBError

    
    strDBName = right(guServOpts.uDB.strDBPath, Len(guServOpts.uDB.strDBPath) - InStrRev(guServOpts.uDB.strDBPath, "\"))
    
    For i = guServOpts.uDB.intDBMaxBackups - 1 To 1 Step -1
        strOldName = guServOpts.uDB.strDBBackupPath & "\" & strDBName & ".bak" & i
        strNewName = guServOpts.uDB.strDBBackupPath & "\" & strDBName & ".bak" & i + 1
        If Dir(strOldName) <> "" Then
            If Dir(strNewName) <> "" Then Call Kill(strNewName)
            Name strOldName As strNewName
        End If
    Next i
    Call FileCopy(guServOpts.uDB.strDBPath, guServOpts.uDB.strDBBackupPath & "\" & strDBName & ".bak1")
    Exit Sub

subBackupDBError:
    Call funErrorManager("modFunctionsDB.subBackupDBError", False, True)
End Sub


'*******************************************************************************
Public Sub subRestoreDB()
Dim intBackupNumber As Integer
Dim strBackupPath As String
Dim strDBName As String
Dim wspCybTest As DAO.Workspace
Dim dbsCybTest As DAO.Database
On Error GoTo subRestoreDBError

    intBackupNumber = CInt(GetRegSettings(funLoadResString(10006), "DB", "Restore backup", 0))
    Call SetRegSettings(funLoadResString(10006), "DB", "Restore backup", 0)
    If intBackupNumber > 0 Then
        strDBName = right(guServOpts.uDB.strDBPath, Len(guServOpts.uDB.strDBPath) - InStrRev(guServOpts.uDB.strDBPath, "\"))
        strBackupPath = guServOpts.uDB.strDBBackupPath & "\" & strDBName & ".bak" & intBackupNumber
        Set wspCybTest = CreateWorkspace("Cyb Admin", "admin", "", dbUseJet)
        Set dbsCybTest = funDBTestOpen(wspCybTest, strBackupPath, guServOpts.uDB.strDBPasswd)
        If Not dbsCybTest Is Nothing Then Call dbsCybTest.Close
        Call wspCybTest.Close
        If Dir(guServOpts.uDB.strDBPath) <> "" Then Call Kill(guServOpts.uDB.strDBPath)
        Call FileCopy(strBackupPath, guServOpts.uDB.strDBPath)
        Call funMsg(enuMsg.Info, App.Title & funLoadResString(10731))
    End If
    Exit Sub

subRestoreDBError:
    Call funErrorManager("modFunctionsDB.subRestoreDBError", False, True)
    If funMsg(enuMsg.GeneralDBError, App.Title & funLoadResString(10726)) = vbYes Then Call funShutdownApplication
End Sub


'*******************************************************************************
Public Function funDBTestOpen(wspWorkspace As DAO.Workspace, strPath As String, strPasswd As String) As DAO.Database
On Error GoTo funDBTestOpenError

    Set funDBTestOpen = wspWorkspace.OpenDatabase(strPath, False, False, ";pwd=" & strPasswd)
    Exit Function

funDBTestOpenError:
Dim strDBErrorMsg As String
    strDBErrorMsg = Err.Description
    Select Case Err.Number
    Case 3024
        Call funMsg(enuMsg.DBPathError, strDBErrorMsg)
        Call Err.Raise(CYBERA_ERROR, "modFunctions.funDBTestOpenError", strDBErrorMsg)
    
    Case 3031
'        Call funMsg(enuMsg.DBPasswordError)

'    Case 3011
'        If funMsg(enuMsg.DBVersionError) = vbYes Then call funShutdownApplication
    
    Case Else
        Call funErrorManager("modFunctionsDB.funDBTestOpenError", False, True)
        Call Err.Raise(CYBERA_ERROR, "modFunctions.funDBTestOpenError", strDBErrorMsg)
    End Select
End Function


'*******************************************************************************
Public Sub subQueryIncome(dteMinDate As Date, dteMaxDate As Date, strTypes As String, Optional strCardID As String = "", Optional strID As String = "", Optional strClientLogin As String = "")
Dim strQuery As String
Dim intMinQueryLength As Long
On Error GoTo subQueryIncomeError
        
    strQuery = "SELECT * FROM journ_income WHERE"
    intMinQueryLength = Len(strQuery)
    If dteMinDate <> NULL_DATE Then strQuery = strQuery & " journ_income.date_stamp >= CDate('" & DateValue(dteMinDate) & " 00:00:00')"
    If dteMaxDate <> NULL_DATE Then
        If Len(strQuery) > intMinQueryLength Then strQuery = strQuery & " AND"
        strQuery = strQuery & " journ_income.date_stamp <= CDate('" & DateValue(dteMaxDate) & " 23:59:59')"
    End If
    If strID <> "" Then
        If Len(strQuery) > intMinQueryLength Then strQuery = strQuery & " AND"
        strQuery = strQuery & " journ_income.ID  = " & strID
    End If
    If strTypes <> "" Then
        If Len(strQuery) > intMinQueryLength Then strQuery = strQuery & " AND"
        strQuery = strQuery & " journ_income.type IN (" & strTypes & ")"
    End If
    If strCardID <> "" Then
        If Len(strQuery) > intMinQueryLength Then strQuery = strQuery & " AND"
        strQuery = strQuery & " journ_income.card_id = " & strCardID
    End If
    If strClientLogin <> "" Then
        If Len(strQuery) > intMinQueryLength Then strQuery = strQuery & " AND"
        strQuery = strQuery & " journ_income.client_login = '" & strClientLogin & "'"
    End If
    strQuery = strQuery & " ORDER BY journ_income.ID"
    Set grstJournIncomeQry = gdbsCybAdmin.OpenRecordset(strQuery, dbOpenSnapshot)
    Exit Sub
    
subQueryIncomeError:
    Call funErrorManager("modFunctionsDB.subQueryIncomeError")
End Sub


'*******************************************************************************
Public Sub subQueryJournIncome(dteMinDate As Date, dteMaxDate As Date, strTypes As String, rstJournIncomeQry As DAO.Recordset)
Dim strQuery As String
Dim intMinQueryLength As Long
On Error GoTo subQueryJournIncomeError
        
    strQuery = "SELECT * FROM journ_income WHERE"
    intMinQueryLength = Len(strQuery)
    If dteMinDate <> NULL_DATE Then strQuery = strQuery & " journ_income.date_stamp >= CDate('" & DateValue(dteMinDate) & " 00:00:00')"
    If dteMaxDate <> NULL_DATE Then
        If Len(strQuery) > intMinQueryLength Then strQuery = strQuery & " AND"
        strQuery = strQuery & " journ_income.date_stamp <= CDate('" & DateValue(dteMaxDate) & " 23:59:59')"
    End If
    If strTypes <> "" Then
        If Len(strQuery) > intMinQueryLength Then strQuery = strQuery & " AND"
        strQuery = strQuery & " journ_income.type IN (" & strTypes & ")"
    End If
    strQuery = strQuery & " ORDER BY journ_income.ID"
    Set rstJournIncomeQry = gdbsCybAdmin.OpenRecordset(strQuery, dbOpenSnapshot)
    Exit Sub
    
subQueryJournIncomeError:
    Call funErrorManager("modFunctionsDB.subQueryJournIncomeError")
End Sub


'*******************************************************************************
Public Sub subQuerySession(dteMinDate As Date, dteMaxDate As Date, Optional strCardID As String = "", Optional strClientLogin As String = "")
Dim strQuery As String
Dim intMinQueryLength As Long
On Error GoTo subQuerySessionError
        
    strQuery = "SELECT * FROM journ_sessions WHERE"
    intMinQueryLength = Len(strQuery)
    If dteMinDate <> NULL_DATE Then strQuery = strQuery & " journ_sessions.date_start >= CDate('" & DateValue(dteMinDate) & " 00:00:00')"
    If dteMaxDate <> NULL_DATE Then
        If Len(strQuery) > intMinQueryLength Then strQuery = strQuery & " AND"
        strQuery = strQuery & " journ_sessions.date_start <= CDate('" & DateValue(dteMaxDate) & " 23:59:59')"
    End If
    If strCardID <> "" Then
        If Len(strQuery) > intMinQueryLength Then strQuery = strQuery & " AND"
        strQuery = strQuery & " journ_sessions.card_id = " & strCardID
    End If
    If strClientLogin <> "" Then
        If Len(strQuery) > intMinQueryLength Then strQuery = strQuery & " AND"
        strQuery = strQuery & " journ_sessions.client_login = '" & strClientLogin & "'"
    End If
    strQuery = strQuery & " ORDER BY journ_sessions.date_start"
    Set grstJournSessionsQry = gdbsCybAdmin.OpenRecordset(strQuery, dbOpenSnapshot)
    Exit Sub
    
subQuerySessionError:
    Call funErrorManager("modFunctionsDB.subQuerySessionError")
End Sub


'*******************************************************************************
Public Sub subQueryRecharges(blnTypeCard As Boolean, blnTypeClient As Boolean, blnTypePrepaidSs As Boolean, blnTypePostpaidSs As Boolean, Optional varOrderMode As Variant, Optional varUnlimitedPlay As Variant)
Dim strQuery As String
Dim strOrderBy As String
Dim intMinQueryLength As Long
Dim intOrderMode As Integer
On Error GoTo subQueryRechargesError
        
    strQuery = "SELECT * FROM recharges WHERE"
    intMinQueryLength = Len(strQuery)
    If blnTypeCard Then strQuery = strQuery & " recharges.type_card = True"
    If blnTypeClient Then
        If Len(strQuery) > intMinQueryLength Then strQuery = strQuery & " AND"
        strQuery = strQuery & " recharges.type_client = True"
    End If
    If blnTypePrepaidSs Then
        If Len(strQuery) > intMinQueryLength Then strQuery = strQuery & " AND"
        strQuery = strQuery & " recharges.type_prepaid_ss = True"
    End If
    If blnTypePostpaidSs Then
        If Len(strQuery) > intMinQueryLength Then strQuery = strQuery & " AND"
        strQuery = strQuery & " recharges.type_postpaid_ss = True"
    End If
    If Not IsMissing(varUnlimitedPlay) Then
        If Len(strQuery) > intMinQueryLength Then strQuery = strQuery & " AND"
        strQuery = strQuery & " recharges.unlimited_play = " & CBool(varUnlimitedPlay)
    End If
    If IsMissing(varOrderMode) Then
        intOrderMode = enuRechargeOrder.RechargesIndex
    Else
        intOrderMode = varOrderMode
    End If
    Select Case intOrderMode
    Case enuRechargeOrder.orderByTime
        strOrderBy = " ORDER BY recharges.time, recharges.price, recharges.caption"
    Case enuRechargeOrder.orderByPrice
        strOrderBy = " ORDER BY recharges.price, recharges.time, recharges.caption"
    Case enuRechargeOrder.orderByValidity
        strOrderBy = " ORDER BY recharges.validity_type, recharges.validity, recharges.caption, recharges.time, recharges.price"
    Case enuRechargeOrder.orderByType
        strOrderBy = " ORDER BY recharges.type_card, recharges.type_client, recharges.type_prepaid_ss, recharges.type_postpaid_ss, recharges.caption, recharges.time, recharges.price"
    Case enuRechargeOrder.RechargesIndex
        strOrderBy = " ORDER BY recharges.caption, recharges.time, recharges.price"
    Case enuRechargeOrder.PrimaryKey
        strOrderBy = " ORDER BY recharges.id"
    End Select
    strQuery = strQuery & strOrderBy
    Set grstRechargesQry = gdbsCybAdmin.OpenRecordset(strQuery, dbOpenSnapshot)
    
    If grstRechargesQry.EOF And grstRechargesQry.BOF Then Exit Sub
    Call grstRechargesQry.MoveLast
    Call grstRechargesQry.MoveFirst
    
    Exit Sub
    
subQueryRechargesError:
    Call funErrorManager("modFunctionsDB.subQueryRechargesError")
End Sub


'*******************************************************************************
Public Sub subQueryURLs(dteMinDate As Date, dteMaxDate As Date, Optional strCardID As String = "", Optional strSessionID As String = "", Optional strClientLogin As String = "")
Dim strQuery As String
Dim intMinQueryLength As Long
On Error GoTo subQueryURLsError
        
    strQuery = "SELECT * FROM urls WHERE"
    intMinQueryLength = Len(strQuery)
    If dteMinDate <> NULL_DATE Then strQuery = strQuery & " urls.date_stamp >= CDate('" & DateValue(dteMinDate) & " 00:00:00')"
    If dteMaxDate <> NULL_DATE Then
        If Len(strQuery) > intMinQueryLength Then strQuery = strQuery & " AND"
        strQuery = strQuery & " urls.date_stamp <= CDate('" & DateValue(dteMaxDate) & " 23:59:59')"
    End If
    If strCardID <> "" Then
        If Len(strQuery) > intMinQueryLength Then strQuery = strQuery & " AND"
        strQuery = strQuery & " urls.card_id = " & strCardID
    End If
    If strClientLogin <> "" Then
        If Len(strQuery) > intMinQueryLength Then strQuery = strQuery & " AND"
        strQuery = strQuery & " urls.client_login = '" & strClientLogin & "'"
    End If
    strQuery = strQuery & " ORDER BY urls.date_stamp"
    Set grstUrlsQry = gdbsCybAdmin.OpenRecordset(strQuery, dbOpenSnapshot)
    Exit Sub
    
subQueryURLsError:
    Call funErrorManager("modFunctionsDB.subQueryURLsError")
End Sub


'*******************************************************************************
Public Sub subQueryEventExact(dteMinDate As Date, dteMaxDate As Date)
Dim strQuery As String
On Error GoTo subQueryEventExactError
        
    strQuery = "SELECT * FROM events WHERE"
    strQuery = strQuery & " events.date_stamp >= CDate('" & DateValue(dteMinDate) & " " & TimeValue(dteMinDate) & "')"
    strQuery = strQuery & " AND events.date_stamp <= CDate('" & DateValue(dteMaxDate) & " " & TimeValue(dteMaxDate) & "')"
    strQuery = strQuery & " ORDER BY events.date_stamp"
    Set grstEventsQry = gdbsCybAdmin.OpenRecordset(strQuery, dbOpenSnapshot)
    Exit Sub
    
subQueryEventExactError:
    Call funErrorManager("modFunctionsDB.subQueryEventExactError")
End Sub


'*******************************************************************************
Public Sub subQueryEvent(dteMinDate As Date, dteMaxDate As Date)
Dim strQuery As String
On Error GoTo subQueryEventError
        
    strQuery = "SELECT * FROM events WHERE"
    strQuery = strQuery & " events.date_stamp >= CDate('" & DateValue(dteMinDate) & " 00:00:00')"
    strQuery = strQuery & " AND events.date_stamp <= CDate('" & DateValue(dteMaxDate) & " 23:59:59')"
    strQuery = strQuery & " ORDER BY events.date_stamp"
    Set grstEventsQry = gdbsCybAdmin.OpenRecordset(strQuery, dbOpenSnapshot)
    Exit Sub
    
subQueryEventError:
    Call funErrorManager("modFunctionsDB.subQueryEventError")
End Sub


'*******************************************************************************
Public Sub subUpdateSessions()
Dim i As Long
Dim uWorkstation As New clsWorkstation
On Error GoTo subUpdateSessionsError
    
    For i = 1 To guWorkstations.Count
        Set uWorkstation = guWorkstations(i)
        If uWorkstation.SessionID <> -1 Then 'If the session is still open...
            grstJournSessions.Index = "PrimaryKey"
            Call grstJournSessions.Seek("=", uWorkstation.SessionID)
            If Not grstJournSessions.NoMatch Then
                If uWorkstation.State >= enuWksnState.UsedDisconnected Then 'If the workstation is still being used...
                    Call grstJournSessions.Edit
                    grstJournSessions!card_id = uWorkstation.CurrentCard.ID
                    grstJournSessions!client_login = uWorkstation.CurrentClient.login
'                    grstJournSessions!date_start = uWorkstation.StartTime
                    grstJournSessions!Type = uWorkstation.SessionType
                    grstJournSessions!computer_name = uWorkstation.ComputerName
                    grstJournSessions!workstation_name = uWorkstation.WorkstationName
                    grstJournSessions!time_connected = uWorkstation.ConnectedTime
                    grstJournSessions!time_connected_offpeak = uWorkstation.ConnectedTimeOffPeak
                    grstJournSessions!time_disconnected = uWorkstation.DisconnectedTime
                    
                    If uWorkstation.SessionType = enuSessionType.Card And uWorkstation.State <= enuWksnState.Blocked_NoCredit Then
                        grstCards.Index = "PrimaryKey"
                        Call grstCards.Seek("=", uWorkstation.CurrentCard.ID)
                        If Not grstCards.NoMatch Then
                            Call grstCards.Edit
                            grstCards!time_credit = uWorkstation.CurrentCard.TimeCredit
                            Call grstCards.Update
                        End If
                    ElseIf uWorkstation.SessionType = enuSessionType.Client And uWorkstation.State <= enuWksnState.Blocked_NoCredit Then
                        grstClients.Index = "PrimaryKey"
                        Call grstClients.Seek("=", uWorkstation.CurrentClient.login)
                        If Not grstClients.NoMatch Then
                            Call grstClients.Edit
                            grstClients!time_credit = uWorkstation.CurrentClient.TimeCredit
                            Call grstClients.Update
                        End If
                    ElseIf uWorkstation.SessionType = enuSessionType.Selfservice Then
                        grstJournIncome.Index = "session_id"
                        Call grstJournIncome.Seek("=", uWorkstation.SessionID)
                        If Not grstJournIncome.NoMatch Then
                            Call grstJournIncome.Edit
                            grstJournIncome!income = uWorkstation.UserBalance
                            grstJournIncome!Time = uWorkstation.ConnectedTime + uWorkstation.ConnectedTimeOffPeak * (guServOpts.uTimeRatios.intOffPeak / 100) + uWorkstation.DisconnectedTime * (guServOpts.uTimeRatios.intDisconCon / 100)
                            Call grstJournIncome.Update
                        End If
                    End If
                    Call grstJournSessions.Update
                ElseIf grstJournSessions!computer_name <> "" Then  'End the session
                    Call subSessionClose(uWorkstation)
                End If
            End If
        End If
    Next i
    Exit Sub
    
subUpdateSessionsError:
    Call funErrorManager("modFunctionsDB.subUpdateSessionsError")
End Sub


'*******************************************************************************
Public Sub subDBAddEvent(lngType As enuEvents, lngSessionID As Long, lngCardID As Long, strClientLogin As String, lngIncomeID As Long, strLabel As String, Optional intEventLevel As enuEventLevels = enuEventLevels.Production)
On Error GoTo subDBAddEventError
        
    If intEventLevel > gintEventLevel Then Exit Sub
    Call grstEvents.AddNew
    grstEvents!Type = lngType
    grstEvents!session_id = lngSessionID
    grstEvents!card_id = lngCardID
    grstEvents!client_login = strClientLogin
    grstEvents!income_id = lngIncomeID
    grstEvents!date_stamp = Now
    grstEvents!Label = left(strLabel, 200)
    Call grstEvents.Update
    If gblnDebug Then Call subDebugPrint("Event: Type " & lngType & " | Label: " & strLabel & " | SessionID: " & lngSessionID & " | CardID: " & lngCardID & " | ClientLogin: " & strClientLogin)
    Exit Sub
    
subDBAddEventError:
    Call funErrorManager("modFunctionsDB.subDBAddEventError")
End Sub


'*******************************************************************************
Public Sub subDBAddCard()
On Error GoTo subDBAddCardError
        
    Exit Sub
    
subDBAddCardError:
    Call funErrorManager("modFunctionsDB.subDBAddCardError")
End Sub


'*******************************************************************************
Public Sub subDBAddURL(lngSessionID As Long, lngCardID As Long, strClientLogin As String, strWorkstationName As String, lngHWnd As Long, strTitle As String, strURL As String)
On Error GoTo subDBAddURLError
        
    Call grstUrls.AddNew
    grstUrls!session_id = lngSessionID
    grstUrls!card_id = lngCardID
    grstUrls!client_login = strClientLogin
    grstUrls!workstation_name = strWorkstationName
    grstUrls!hWnd = lngHWnd
    grstUrls!date_stamp = Now
    grstUrls!Title = left(strTitle, 200)
    grstUrls!url = left(strURL, 200)
    Call grstUrls.Update
    Exit Sub
    
subDBAddURLError:
    Call funErrorManager("modFunctionsDB.subDBAddURLError")
End Sub


'*******************************************************************************
Public Sub subDBAddSession()
On Error GoTo subDBAddSessionError
        
    Exit Sub
    
subDBAddSessionError:
    Call funErrorManager("modFunctionsDB.subDBAddSessionError")
End Sub


'*******************************************************************************
Public Sub subDBAddIncome()
On Error GoTo subDBAddIncomeError
        
    Exit Sub
    
subDBAddIncomeError:
    Call funErrorManager("modFunctionsDB.subDBAddIncomeError")
End Sub


'*******************************************************************************
Public Sub subDBAddRecharge()
On Error GoTo subDBAddRechargeError
        
    Exit Sub
    
subDBAddRechargeError:
    Call funErrorManager("modFunctionsDB.subDBAddRechargeError")
End Sub


'*******************************************************************************
Public Sub subDBAddService(uService As clsService)
On Error GoTo subDBAddServiceError
        
    Call grstServices.AddNew
    grstServices!Label = uService.Label
    grstServices!Price = uService.Price
    grstServices!Stock = uService.Stock
    grstServices!stock_mode = uService.StockMode
    grstServices!stock_min_alert = uService.StockMinAlert
    Call grstServices.Update
    Exit Sub
    
subDBAddServiceError:
    Call funErrorManager("modFunctionsDB.subDBAddServiceError")
End Sub


'*******************************************************************************
Public Function funDBAddUser(uUser As clsUser) As Boolean
On Error GoTo subDBAddUserError
        
    Call grstUsers.AddNew
    grstUsers!login = uUser.login
    grstUsers!full_name = uUser.FullName
    grstUsers!Passwd = funDoXOR(uUser.Passwd, funLoadResString(10006))
    grstUsers!access_journals = uUser.AccessJournals
    grstUsers!access_unblock = uUser.AccessUnblock
    grstUsers!access_del_recharge = uUser.AccessDelRecharge
    grstUsers!access_mass_card_creation = uUser.AccessMassCardCreation
    grstUsers!access_card_creation = uUser.AccessCardCreation
    grstUsers!access_client_creation = uUser.AccessClientCreation
    Call grstUsers.Update
    funDBAddUser = True
    Exit Function
    
subDBAddUserError:
    funDBAddUser = False
End Function


'*******************************************************************************
Public Sub subDBUpdateService(uService As clsService)
On Error GoTo subDBUpdateServiceError

    grstServices.Index = "PrimaryKey"
    Call grstServices.Seek("=", uService.ID)
    If Not grstServices.NoMatch Then
        Call grstServices.Edit
        grstServices!Label = uService.Label
        grstServices!Price = uService.Price
        grstServices!Stock = uService.Stock
        grstServices!stock_mode = uService.StockMode
        grstServices!stock_min_alert = uService.StockMinAlert
        Call grstServices.Update
    End If
    Exit Sub
    
subDBUpdateServiceError:
    Call funErrorManager("modFunctionsDB.subDBUpdateServiceError")
End Sub


'*******************************************************************************
Public Sub subDBUpdateUser(uUser As clsUser)
On Error GoTo subDBUpdateUserError

    grstUsers.Index = "PrimaryKey"
    Call grstUsers.Seek("=", uUser.login)
    If Not grstUsers.NoMatch Then
        Call grstUsers.Edit
        grstUsers!full_name = uUser.FullName
        grstUsers!Passwd = funDoXOR(uUser.Passwd, funLoadResString(10006))
        grstUsers!access_journals = uUser.AccessJournals
        grstUsers!access_unblock = uUser.AccessUnblock
        grstUsers!access_del_recharge = uUser.AccessDelRecharge
        grstUsers!access_mass_card_creation = uUser.AccessMassCardCreation
        grstUsers!access_card_creation = uUser.AccessCardCreation
        grstUsers!access_client_creation = uUser.AccessClientCreation
        Call grstUsers.Update
    End If
    Exit Sub
    
subDBUpdateUserError:
    Call funErrorManager("modFunctionsDB.subDBUpdateUserError")
End Sub


'*******************************************************************************
Public Sub subDBUpdateClient(uClient As clsClient)
On Error GoTo subDBUpdateClientError

    grstClients.Index = "PrimaryKey"
    Call grstClients.Seek("=", uClient.login)
    If Not grstUsers.NoMatch Then
        Call grstClients.Edit
        grstClients!full_name = uClient.FullName
        grstClients!Passwd = uClient.Password
        grstClients!Comment = uClient.Comment
        grstClients!Active = uClient.Active
        Call grstClients.Update
    End If
    Exit Sub
    
subDBUpdateClientError:
    Call funErrorManager("modFunctionsDB.subDBUpdateClientError")
End Sub


'*******************************************************************************
Public Sub subDBUpdateCard(uCard As clsCard)
On Error GoTo subDBUpdateCardError

    grstCards.Index = "PrimaryKey"
    Call grstCards.Seek("=", uCard.ID)
    If Not grstCards.NoMatch Then
        Call grstCards.Edit
        Call grstCards.Update
    End If
    Exit Sub
    
subDBUpdateCardError:
    Call funErrorManager("modFunctionsDB.subDBUpdateCardError")
End Sub


'*******************************************************************************
Public Sub subDBDeleteService(uService As clsService)
On Error GoTo subDBDeleteServiceError
        
    grstServices.Index = "PrimaryKey"
    Call grstServices.Seek("=", uService.ID)
    If Not grstServices.NoMatch Then
        grstServices.Delete
    End If
    Exit Sub
    
subDBDeleteServiceError:
    Call funErrorManager("modFunctionsDB.subDBDeleteServiceError")
End Sub


'*******************************************************************************
Public Sub subDBDeleteUser(uUser As clsUser)
On Error GoTo subDBDeleteUserError
        
    grstUsers.Index = "PrimaryKey"
    Call grstUsers.Seek("=", uUser.login)
    If Not grstUsers.NoMatch Then
        grstUsers.Delete
    End If
    Exit Sub
    
subDBDeleteUserError:
    Call funErrorManager("modFunctionsDB.subDBDeleteUserError")
End Sub


'*******************************************************************************
Public Sub subDBDeleteClient(uClient As clsClient)
On Error GoTo subDBDeleteClientError
        
    grstClients.Index = "PrimaryKey"
    Call grstClients.Seek("=", uClient.login)
    If Not grstClients.NoMatch Then
        grstClients.Delete
        Call subDBAddEvent(enuEvents.ClientDelete, 0, 0, uClient.login, 0, funLoadResString(10875), Verbose)
    End If
    Exit Sub
    
subDBDeleteClientError:
    Call funErrorManager("modFunctionsDB.subDBDeleteClientError")
End Sub


'*******************************************************************************
Public Sub subDBDeleteIncome(lngID As Long)
On Error GoTo subDBDeleteIncomeError
        
    grstJournIncome.Index = "PrimaryKey"
    Call grstJournIncome.Seek("=", CStr(lngID))
    If Not grstJournIncome.NoMatch Then
        grstJournIncome.Delete
    End If
    Exit Sub
    
subDBDeleteIncomeError:
    Call funErrorManager("modFunctionsDB.subDBDeleteIncomeError")
End Sub


'*******************************************************************************
Public Sub subCardsPurge()
Dim dteLimit As Date
On Error GoTo subCardsPurgeError

    dteLimit = DateAdd("d", -guServOpts.uCards.intCardPurgeDelta * guServOpts.uConfig.intNumDaysInMonth, Now)
    grstCards.Index = "PrimaryKey"
    Call grstCards.MoveFirst
    Do While Not grstCards.EOF
        If grstCards!date_expiry < dteLimit Then grstCards.Delete
        Call grstCards.MoveNext
    Loop
    Exit Sub
    
subCardsPurgeError:
    Call funErrorManager("modFunctionsDB.subCardsPurgeError")
End Sub


'*******************************************************************************
Public Sub subClientsPurge()
Dim dteLimit As Date
On Error GoTo subClientsPurgeError

    dteLimit = DateAdd("d", -guServOpts.uClients.intClientPurgeDelta * guServOpts.uConfig.intNumDaysInMonth, Now)
    grstClients.Index = "PrimaryKey"
    Call grstClients.MoveFirst
    Do While Not grstClients.EOF
        If grstClients!date_expiry < dteLimit Then grstClients.Delete
        Call grstClients.MoveNext
    Loop
    Exit Sub
    
subClientsPurgeError:
    Call funErrorManager("modFunctionsDB.subClientsPurgeError")
End Sub


'*******************************************************************************
Public Sub subCardsMaintenance()
On Error GoTo subCardsMaintenanceError

    If grstCards.BOF = True And grstCards.EOF = True Then Exit Sub
    
    If guServOpts.uCards.intAutoCardPurge = vbChecked Then Call subCardsPurge
    grstCards.Index = "PrimaryKey"
    Call grstCards.MoveFirst
    Do While Not grstCards.EOF
        If grstCards!date_expiry < Now Then
            Call grstCards.Edit
            If grstCards!unlimited_play Then
                grstCards!unlimited_play = False
            End If
            grstCards!time_credit = NULL_DATE
            Call grstCards.Update
        End If
        Call grstCards.MoveNext
        DoEvents
    Loop
    Exit Sub
    
subCardsMaintenanceError:
    Call funErrorManager("modFunctionsDB.subCardsMaintenanceError")
End Sub


'*******************************************************************************
Public Sub subClientsMaintenance()
On Error GoTo subClientsMaintenanceError

    If grstClients.BOF = True And grstClients.EOF = True Then Exit Sub
    
    If guServOpts.uClients.intAutoClientPurge = vbChecked Then Call subClientsPurge
    grstClients.Index = "PrimaryKey"
    Call grstClients.MoveFirst
    Do While Not grstClients.EOF
        If grstClients!date_expiry < Now Then
            Call grstClients.Edit
            If grstClients!unlimited_play Then
                grstClients!unlimited_play = False
            End If
            grstClients!time_credit = NULL_DATE
            Call grstClients.Update
        End If
        Call grstClients.MoveNext
        DoEvents
    Loop
    Exit Sub
    
subClientsMaintenanceError:
    Call funErrorManager("modFunctionsDB.subClientsMaintenanceError")
End Sub


'*******************************************************************************
Public Sub subDBMaintenance(dtePurgeLimit As Date, blnIncludeIncome As Boolean)
Dim rstEvents As DAO.Recordset
Dim rstJournSessions As DAO.Recordset
Dim rstJournIncome As DAO.Recordset
Dim rstUrls As DAO.Recordset
On Error GoTo subDBMaintenanceError

    Set rstEvents = gdbsCybAdmin.OpenRecordset("events", dbOpenTable)
    rstEvents.Index = "PrimaryKey"
    Call rstEvents.MoveFirst
    Do While Not rstEvents.EOF
        If rstEvents!date_stamp < dtePurgeLimit Then rstEvents.Delete
        Call rstEvents.MoveNext
        DoEvents
        Screen.MousePointer = vbHourglass
    Loop
    Call rstEvents.Close
    If blnIncludeIncome Then
        Set rstJournIncome = gdbsCybAdmin.OpenRecordset("journ_income", dbOpenTable)
        rstJournIncome.Index = "PrimaryKey"
        Call rstJournIncome.MoveFirst
        Do While Not rstJournIncome.EOF
            If rstJournIncome!date_stamp < dtePurgeLimit Then rstJournIncome.Delete
            Call rstJournIncome.MoveNext
            DoEvents
            Screen.MousePointer = vbHourglass
        Loop
        Call rstJournIncome.Close
    End If
    Set rstJournSessions = gdbsCybAdmin.OpenRecordset("journ_sessions", dbOpenTable)
    rstJournSessions.Index = "PrimaryKey"
    Call rstJournSessions.MoveFirst
    Do While Not rstJournSessions.EOF
        If rstJournSessions!date_start < dtePurgeLimit Then rstJournSessions.Delete
        Call rstJournSessions.MoveNext
        DoEvents
        Screen.MousePointer = vbHourglass
    Loop
    Call rstJournSessions.Close
    Set rstUrls = gdbsCybAdmin.OpenRecordset("urls", dbOpenTable)
    rstUrls.Index = "PrimaryKey"
    Call rstUrls.MoveFirst
    Do While Not rstUrls.EOF
        If rstUrls!date_stamp < dtePurgeLimit Then rstUrls.Delete
        Call rstUrls.MoveNext
        DoEvents
        Screen.MousePointer = vbHourglass
    Loop
    Call rstUrls.Close
    Exit Sub
    
subDBMaintenanceError:
    Call funErrorManager("modFunctionsDB.subDBMaintenanceError")
End Sub


'*******************************************************************************
Public Sub subDBPurgeURLs()
Dim dtePurgeLimit As Date
On Error GoTo subDBPurgeURLsError

    If guServOpts.uWorkstations.intLoggedURLTimeSpan = 0 Or (grstUrls.EOF And grstUrls.BOF) Then Exit Sub
    
    dtePurgeLimit = DateAdd("d", -guServOpts.uWorkstations.intLoggedURLTimeSpan * guServOpts.uConfig.intNumDaysInMonth, CDate(DateValue(Now) & " 00:00:00"))
    
    grstUrls.Index = "PrimaryKey"
    Call grstUrls.MoveFirst
    Do While Not grstUrls.EOF
        If grstUrls!date_stamp < dtePurgeLimit Then grstUrls.Delete
        Call grstUrls.MoveNext
'        DoEvents
    Loop
    Exit Sub
    
subDBPurgeURLsError:
    Call funErrorManager("modFunctionsDB.subDBPurgeURLsError")
End Sub


'*******************************************************************************
Public Sub subDBCompact()
Dim strTempDB As String
On Error Resume Next

    If Not gblnDBOK Then Exit Sub
    strTempDB = guServOpts.uDB.strDBPath & ".tmp"
    If Dir(strTempDB) <> "" Then Call Kill(strTempDB)
    Call DBEngine.CompactDatabase(guServOpts.uDB.strDBPath, strTempDB, ";pwd=" & guServOpts.uDB.strDBPasswd, , ";pwd=" & guServOpts.uDB.strDBPasswd)
    Call FileCopy(strTempDB, guServOpts.uDB.strDBPath)
    Call Kill(strTempDB)
End Sub


'*******************************************************************************
Public Function funDBUpgrade() As Boolean
Dim intcurrentDBVersion As Integer
Dim tdfTemp As DAO.TableDef
Dim fldTemp As DAO.Field
Dim idxTemp As DAO.Index
On Error GoTo funDBUpgradeError

    Screen.MousePointer = vbHourglass
    funDBUpgrade = True
    Call grstConfig.MoveFirst
    intcurrentDBVersion = grstConfig!Version
    Call funMsg(DBUpdrageWarning)
    Call Load(frmDBUpdate)
    frmDBUpdate.pgbDBUpgrade.Max = DB_MIN_VERSION
    frmDBUpdate.pgbDBUpgrade.Min = intcurrentDBVersion
    frmDBUpdate.pgbDBUpgrade.value = frmDBUpdate.pgbDBUpgrade.Min
    Call frmDBUpdate.Show
    DoEvents
    Call SetWindowPos(frmDBUpdate.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    Do While intcurrentDBVersion < DB_MIN_VERSION
        intcurrentDBVersion = funDBUpgradeSteps(intcurrentDBVersion)
        frmDBUpdate.pgbDBUpgrade.value = intcurrentDBVersion
        DoEvents
        If intcurrentDBVersion = 0 Then
            funDBUpgrade = False
            Exit Do
        End If
    Loop
    frmDBUpdate.lbl.Caption = frmDBUpdate.lbl.Caption & " " & funLoadResString(10973)
    DoEvents
    Call subDelay(1)
    Call Unload(frmDBUpdate)
    Screen.MousePointer = vbNormal
    Exit Function
    
funDBUpgradeError:
    Call Unload(frmDBUpdate)
    Screen.MousePointer = vbNormal
    funDBUpgrade = False
End Function


'*******************************************************************************
Public Function funDBUpgradeSteps(intBaseVersion As Integer) As Integer
Dim tdfTemp As DAO.TableDef
Dim fldTemp As DAO.Field
Dim idxTemp As DAO.Index
On Error GoTo funDBUpgradeStepsError

    Select Case intBaseVersion
    Case Is < 6
        Call Err.Raise(3011)
    Case 6 '6 -> 7
        ' Add uls table
        Set tdfTemp = gdbsCybAdmin.CreateTableDef("urls")
        Call tdfTemp.Fields.Append(tdfTemp.CreateField("ID", dbLong))
        tdfTemp.Fields("ID").Attributes = dbAutoIncrField
        Call tdfTemp.Fields.Append(tdfTemp.CreateField("session_id", dbLong))
        Call tdfTemp.Fields.Append(tdfTemp.CreateField("card_id", dbLong))
        Call tdfTemp.Fields.Append(tdfTemp.CreateField("client_name", dbText, 50))
        Call tdfTemp.Fields.Append(tdfTemp.CreateField("hWnd", dbLong))
        Call tdfTemp.Fields.Append(tdfTemp.CreateField("date_stamp", dbDate))
        Call tdfTemp.Fields.Append(tdfTemp.CreateField("title", dbText, 200))
        Call tdfTemp.Fields.Append(tdfTemp.CreateField("url", dbText, 200))
        Call gdbsCybAdmin.TableDefs.Append(tdfTemp)
        Set idxTemp = tdfTemp.CreateIndex("PrimaryKey")
        idxTemp.Fields.Append idxTemp.CreateField("ID")
        idxTemp.Primary = True
        Call tdfTemp.Indexes.Append(idxTemp)
            
    Case 7 '7 -> 8
        ' Add users table
        Set tdfTemp = gdbsCybAdmin.CreateTableDef("users")
        Call tdfTemp.Fields.Append(tdfTemp.CreateField("login", dbText, 20))
        tdfTemp.Fields("login").AllowZeroLength = False
        Call tdfTemp.Fields.Append(tdfTemp.CreateField("passwd", dbText, 20))
        tdfTemp.Fields("passwd").AllowZeroLength = True
        Call tdfTemp.Fields.Append(tdfTemp.CreateField("full_name", dbText, 50))
        tdfTemp.Fields("full_name").AllowZeroLength = True
        Call tdfTemp.Fields.Append(tdfTemp.CreateField("access_journals", dbBoolean))
        Call tdfTemp.Fields.Append(tdfTemp.CreateField("access_unblock", dbBoolean))
        Call tdfTemp.Fields.Append(tdfTemp.CreateField("last_login", dbDate))
        Call gdbsCybAdmin.TableDefs.Append(tdfTemp)
        Set idxTemp = tdfTemp.CreateIndex("PrimaryKey")
        idxTemp.Fields.Append idxTemp.CreateField("login")
        idxTemp.Primary = True
        Call tdfTemp.Indexes.Append(idxTemp)
        
        ' Add login col to Income journal
        Set gtdfJournIncome = gdbsCybAdmin!journ_income
        Call gtdfJournIncome.Fields.Append(gtdfJournIncome.CreateField("login", dbText, 20))
        gtdfJournIncome.Fields("login").AllowZeroLength = True
        
        ' Add login col to session journal
        Set gtdfJournSessions = gdbsCybAdmin!journ_sessions
        Call gtdfJournSessions.Fields.Append(gtdfJournSessions.CreateField("login", dbText, 20))
        gtdfJournIncome.Fields("login").AllowZeroLength = True
    
    Case 8 ' 8 -> 9
        ' Add clients table
        Set tdfTemp = gdbsCybAdmin.CreateTableDef("clients")
        Call tdfTemp.Fields.Append(tdfTemp.CreateField("login", dbText, 20))
        tdfTemp.Fields("login").AllowZeroLength = False
        Call tdfTemp.Fields.Append(tdfTemp.CreateField("passwd", dbText, 20))
        tdfTemp.Fields("passwd").AllowZeroLength = True
        Call tdfTemp.Fields.Append(tdfTemp.CreateField("full_name", dbText, 50))
        tdfTemp.Fields("full_name").AllowZeroLength = True
        Call tdfTemp.Fields.Append(tdfTemp.CreateField("date_creation", dbDate))
        Call tdfTemp.Fields.Append(tdfTemp.CreateField("date_expiry", dbDate))
        Call tdfTemp.Fields.Append(tdfTemp.CreateField("time_credit", dbDate))
        Call tdfTemp.Fields.Append(tdfTemp.CreateField("balance", dbDouble))
        Call tdfTemp.Fields.Append(tdfTemp.CreateField("active", dbBoolean))
        Call tdfTemp.Fields.Append(tdfTemp.CreateField("type", dbInteger))
        Call tdfTemp.Fields.Append(tdfTemp.CreateField("comment", dbText, 100))
        tdfTemp.Fields("comment").AllowZeroLength = True
        Call gdbsCybAdmin.TableDefs.Append(tdfTemp)
        Set idxTemp = tdfTemp.CreateIndex("PrimaryKey")
        Call idxTemp.Fields.Append(idxTemp.CreateField("login"))
        idxTemp.Primary = True
        Call tdfTemp.Indexes.Append(idxTemp)
        
        ' Add client_login col to Income journal
        Set gtdfJournIncome = gdbsCybAdmin!journ_income
        Call gtdfJournIncome.Fields.Append(gtdfJournIncome.CreateField("client_login", dbText, 20))
        gtdfJournIncome.Fields("client_login").AllowZeroLength = True
        Set idxTemp = gtdfJournIncome.CreateIndex("client_login")
        idxTemp.Unique = False
        Call idxTemp.Fields.Append(idxTemp.CreateField("client_login"))
        Call gtdfJournIncome.Indexes.Append(idxTemp)
        
        ' Updates to session journal
        Set gtdfJournSessions = gdbsCybAdmin!journ_sessions
        gtdfJournSessions.Fields("client_name").Name = "workstation_name"
        Call gtdfJournSessions.Fields.Append(gtdfJournSessions.CreateField("client_login", dbText, 20))
        gtdfJournSessions.Fields("client_login").AllowZeroLength = True
        Set idxTemp = gtdfJournSessions.CreateIndex("client_login")
        idxTemp.Unique = False
        Call idxTemp.Fields.Append(idxTemp.CreateField("client_login"))
        Call gtdfJournSessions.Indexes.Append(idxTemp)
    
        ' Updates to event table
        Set gtdfEvents = gdbsCybAdmin!events
        Call gtdfEvents.Fields.Append(gtdfEvents.CreateField("client_login", dbText, 20))
        gtdfEvents.Fields("client_login").AllowZeroLength = True
        Set idxTemp = gtdfEvents.CreateIndex("client_login")
        idxTemp.Unique = False
        Call idxTemp.Fields.Append(idxTemp.CreateField("client_login"))
        Call gtdfEvents.Indexes.Append(idxTemp)
    
        ' Updates to urls table
        Set gtdfUrls = gdbsCybAdmin!urls
        gtdfUrls.Fields("client_name").Name = "workstation_name"
        Call gtdfUrls.Fields.Append(gtdfUrls.CreateField("client_login", dbText, 20))
        gtdfUrls.Fields("client_login").AllowZeroLength = True
        Set idxTemp = gtdfUrls.CreateIndex("client_login")
        idxTemp.Unique = False
        Call idxTemp.Fields.Append(idxTemp.CreateField("client_login"))
        Call gtdfUrls.Indexes.Append(idxTemp)
    
    Case 9 '9 -> 10
        ' Add access_del_recharge col to users and set default to True
        Set gtdfUsers = gdbsCybAdmin!Users
        Call gtdfUsers.Fields.Append(gtdfUsers.CreateField("access_del_recharge", dbBoolean))
        
        Set grstUsers = gdbsCybAdmin.OpenRecordset("users", dbOpenTable)
        grstUsers.Index = "PrimaryKey"
        If Not (grstUsers.BOF And grstUsers.EOF) Then
            Call grstUsers.MoveFirst
            Do While Not grstUsers.EOF
                Call grstUsers.Edit
                grstUsers!access_del_recharge = True
                Call grstUsers.Update
                Call grstUsers.MoveNext
            Loop
        End If
        Call grstUsers.Close
    
    Case 10 '10 -> 11
        ' Add cards_pending table
        Set tdfTemp = gdbsCybAdmin.CreateTableDef("cards_pending")
        Call tdfTemp.Fields.Append(tdfTemp.CreateField("ID", dbLong))
        tdfTemp.Fields("ID").AllowZeroLength = False
        Call tdfTemp.Fields.Append(tdfTemp.CreateField("time_credit", dbDate))
        Call tdfTemp.Fields.Append(tdfTemp.CreateField("price", dbCurrency))
        Call tdfTemp.Fields.Append(tdfTemp.CreateField("validity", dbLong))
        Call tdfTemp.Fields.Append(tdfTemp.CreateField("caption", dbText, 50))
        tdfTemp.Fields("caption").AllowZeroLength = True
        Call gdbsCybAdmin.TableDefs.Append(tdfTemp)
        Set idxTemp = tdfTemp.CreateIndex("PrimaryKey")
        Call idxTemp.Fields.Append(idxTemp.CreateField("ID"))
        idxTemp.Primary = True
        Call tdfTemp.Indexes.Append(idxTemp)
    
    Case 11 '11 -> 12
        ' Add access_mass_card_creation col to users and set default to True
        Set gtdfUsers = gdbsCybAdmin!Users
        Call gtdfUsers.Fields.Append(gtdfUsers.CreateField("access_mass_card_creation", dbBoolean))
        
        Set grstUsers = gdbsCybAdmin.OpenRecordset("users", dbOpenTable)
        grstUsers.Index = "PrimaryKey"
        If Not (grstUsers.BOF And grstUsers.EOF) Then
            Call grstUsers.MoveFirst
            Do While Not grstUsers.EOF
                DoEvents
                Call grstUsers.Edit
                grstUsers!access_mass_card_creation = True
                Call grstUsers.Update
                Call grstUsers.MoveNext
            Loop
        End If
        Call grstUsers.Close
    
    Case 12 '12 -> 13
        ' Add type and comment cols to cards_pending table
        Set gtdfCardsPending = gdbsCybAdmin!cards_pending
        Call gtdfCardsPending.Fields.Append(gtdfCardsPending.CreateField("type", dbText, 50))
        Call gtdfCardsPending.Fields.Append(gtdfCardsPending.CreateField("comment", dbText, 100))
    
    Case 13 '13 -> 14
        ' Change the size of the login col of the users table to 255 chars
        Set gtdfUsers = gdbsCybAdmin!Users
        Call gtdfUsers.Fields.Append(gtdfUsers.CreateField("temp_login", dbText, 20))
        Set grstUsers = gdbsCybAdmin.OpenRecordset("users", dbOpenTable)
        If Not (grstUsers.BOF And grstUsers.EOF) Then
            Call grstUsers.MoveFirst
            Do While Not grstUsers.EOF
                DoEvents
                Call grstUsers.Edit
                grstUsers!temp_login = grstUsers!login
                Call grstUsers.Update
                Call grstUsers.MoveNext
            Loop
        End If
        Call grstUsers.Close
        Call gtdfUsers.Indexes.Delete("PrimaryKey")
        Call gtdfUsers.Fields.Delete("login")
        Call gtdfUsers.Fields.Append(gtdfUsers.CreateField("login", dbText, 0))
        Set grstUsers = gdbsCybAdmin.OpenRecordset("users", dbOpenTable)
        If Not (grstUsers.BOF And grstUsers.EOF) Then
            Call grstUsers.MoveFirst
            Do While Not grstUsers.EOF
                DoEvents
                Call grstUsers.Edit
                grstUsers!login = grstUsers!temp_login
                Call grstUsers.Update
                Call grstUsers.MoveNext
            Loop
        End If
        Call grstUsers.Close
        Set idxTemp = gtdfUsers.CreateIndex("PrimaryKey")
        Call idxTemp.Fields.Append(idxTemp.CreateField("login"))
        idxTemp.Primary = True
        Call gtdfUsers.Indexes.Append(idxTemp)
        
    Case 14 '14 -> 15
        ' Fixed big bug in mass card creation
        Set gtdfCardsPending = gdbsCybAdmin!cards_pending
        gtdfCardsPending.Fields("type").AllowZeroLength = True
        gtdfCardsPending.Fields("comment").AllowZeroLength = True
    
    Case 15 '15 -> 16
        ' Fixed bug in users table
        Set gtdfUsers = gdbsCybAdmin!Users
        Call gtdfUsers.Fields.Delete("temp_login")
    
    Case 16 '16 -> 17
        ' Change the size of the login col of the journ_sessions table to 255 chars
        Set gtdfJournSessions = gdbsCybAdmin!journ_sessions
        With gtdfJournSessions
            Call .Fields.Append(.CreateField("temp_login", dbText, 20))
            .Fields("temp_login").AllowZeroLength = True
            Set grstJournSessions = gdbsCybAdmin.OpenRecordset("journ_sessions", dbOpenTable)
            With grstJournSessions
                If Not (.BOF And .EOF) Then
                    Call .MoveFirst
                    Do While Not .EOF
                        DoEvents
                        Call .Edit
                        .Fields("temp_login").value = .Fields("login").value
                        Call .Update
                        Call .MoveNext
                    Loop
                End If
                Call .Close
            End With
'            Call .Indexes.Delete("PrimaryKey")
            Call .Fields.Delete("login")
            Call .Fields.Append(.CreateField("login", dbText, 0))
            .Fields("login").AllowZeroLength = True
            Set grstJournSessions = gdbsCybAdmin.OpenRecordset("journ_sessions", dbOpenTable)
            With grstJournSessions
                If Not (.BOF And .EOF) Then
                    Call .MoveFirst
                    Do While Not .EOF
                        DoEvents
                        Call .Edit
                        .Fields("login").value = .Fields("temp_login").value
                        Call .Update
                        Call .MoveNext
                    Loop
                End If
                Call .Close
            End With
            Call .Fields.Delete("temp_login")
'            Set idxTemp = .CreateIndex("PrimaryKey")
'            Call idxTemp.Fields.Append(idxTemp.CreateField("ID"))
'            idxTemp.Primary = True
'            Call .Indexes.Append(idxTemp)
        End With
        
        ' Change the size of the login col of the journ_income table to 255 chars
        Set gtdfJournIncome = gdbsCybAdmin!journ_income
        With gtdfJournIncome
            Call .Fields.Append(.CreateField("temp_login", dbText, 20))
            .Fields("temp_login").AllowZeroLength = True
            Set grstJournIncome = gdbsCybAdmin.OpenRecordset("journ_income", dbOpenTable)
            With grstJournIncome
                If Not (.BOF And .EOF) Then
                    Call .MoveFirst
                    Do While Not .EOF
                        DoEvents
                        Call .Edit
                        .Fields("temp_login").value = .Fields("login").value
                        Call .Update
                        Call .MoveNext
                    Loop
                End If
                Call .Close
            End With
'            Call .Indexes.Delete("PrimaryKey")
            Call .Fields.Delete("login")
            Call .Fields.Append(.CreateField("login", dbText, 0))
            .Fields("login").AllowZeroLength = True
            Set grstJournIncome = gdbsCybAdmin.OpenRecordset("journ_income", dbOpenTable)
            With grstJournIncome
                If Not (.BOF And .EOF) Then
                    Call .MoveFirst
                    Do While Not .EOF
                        DoEvents
                        Call .Edit
                        .Fields("login").value = .Fields("temp_login").value
                        Call .Update
                        Call .MoveNext
                    Loop
                End If
                Call .Close
            End With
            Call .Fields.Delete("temp_login")
'            Set idxTemp = .CreateIndex("PrimaryKey")
'            Call idxTemp.Fields.Append(idxTemp.CreateField("ID"))
'            idxTemp.Primary = True
'            Call .Indexes.Append(idxTemp)
        End With
    
    Case 17 '17 -> 18
        ' Change the size of the client_login col of the events table to 255 chars
        Set gtdfEvents = gdbsCybAdmin!events
        With gtdfEvents
            Call .Fields.Append(.CreateField("temp_client_login", dbText, 20))
            .Fields("temp_client_login").AllowZeroLength = True
            Set grstEvents = gdbsCybAdmin.OpenRecordset("events", dbOpenTable)
            With grstEvents
                If Not (.BOF And .EOF) Then
                    Call .MoveFirst
                    Do While Not .EOF
                        DoEvents
                        Call .Edit
                        .Fields("temp_client_login").value = .Fields("client_login").value
                        Call .Update
                        Call .MoveNext
                    Loop
                End If
                Call .Close
            End With
            Call .Indexes.Delete("client_login")
            Call .Fields.Delete("client_login")
            Call .Fields.Append(.CreateField("client_login", dbText, 0))
            .Fields("client_login").AllowZeroLength = True
            Set grstEvents = gdbsCybAdmin.OpenRecordset("events", dbOpenTable)
            With grstEvents
                If Not (.BOF And .EOF) Then
                    Call .MoveFirst
                    Do While Not .EOF
                        DoEvents
                        Call .Edit
                        .Fields("client_login").value = .Fields("temp_client_login").value
                        Call .Update
                        Call .MoveNext
                    Loop
                End If
                Call .Close
            End With
            Call .Fields.Delete("temp_client_login")
            Set idxTemp = .CreateIndex("client_login")
            Call idxTemp.Fields.Append(idxTemp.CreateField("client_login"))
            Call .Indexes.Append(idxTemp)
        End With
    
    Case 18
        ' Change the size of the client_login col of the urls table to 255 chars
        Set gtdfUrls = gdbsCybAdmin!urls
        With gtdfUrls
            Call .Fields.Append(.CreateField("temp_client_login", dbText, 20))
            .Fields("temp_client_login").AllowZeroLength = True
            Set grstUrls = gdbsCybAdmin.OpenRecordset("urls", dbOpenTable)
            With grstUrls
                If Not (.BOF And .EOF) Then
                    Call .MoveFirst
                    Do While Not .EOF
                        DoEvents
                        Call .Edit
                        .Fields("temp_client_login").value = .Fields("client_login").value
                        Call .Update 'TOTO
                        Call .MoveNext
                    Loop
                End If
                Call .Close
            End With
            Call .Indexes.Delete("client_login")
            Call .Fields.Delete("client_login")
            Call .Fields.Append(.CreateField("client_login", dbText, 0))
            .Fields("client_login").AllowZeroLength = True
            Set grstUrls = gdbsCybAdmin.OpenRecordset("urls", dbOpenTable)
            With grstUrls
                If Not (.BOF And .EOF) Then
                    Call .MoveFirst
                    Do While Not .EOF
                        DoEvents
                        Call .Edit
                        .Fields("client_login").value = .Fields("temp_client_login").value
                        Call .Update
                        Call .MoveNext
                    Loop
                End If
                Call .Close
            End With
            Call .Fields.Delete("temp_client_login")
            Set idxTemp = .CreateIndex("client_login")
            Call idxTemp.Fields.Append(idxTemp.CreateField("client_login"))
            Call .Indexes.Append(idxTemp)
        End With
    
    Case 19
        ' Change the size of the client_login col of the journ_income table to 255 chars
        Set gtdfJournIncome = gdbsCybAdmin!journ_income
        With gtdfJournIncome
            Call .Fields.Append(.CreateField("temp_client_login", dbText, 20))
            .Fields("temp_client_login").AllowZeroLength = True
            Set grstJournIncome = gdbsCybAdmin.OpenRecordset("journ_income", dbOpenTable)
            With grstJournIncome
                If Not (.BOF And .EOF) Then
                    Call .MoveFirst
                    Do While Not .EOF
                        DoEvents
                        Call .Edit
                        .Fields("temp_client_login").value = .Fields("client_login").value
                        Call .Update
                        Call .MoveNext
                    Loop
                End If
                Call .Close
            End With
            Call .Indexes.Delete("client_login")
            Call .Fields.Delete("client_login")
            Call .Fields.Append(.CreateField("client_login", dbText, 0))
            .Fields("client_login").AllowZeroLength = True
            Set grstJournIncome = gdbsCybAdmin.OpenRecordset("journ_income", dbOpenTable)
            With grstJournIncome
                If Not (.BOF And .EOF) Then
                    Call .MoveFirst
                    Do While Not .EOF
                        DoEvents
                        Call .Edit
                        .Fields("client_login").value = .Fields("temp_client_login").value
                        Call .Update
                        Call .MoveNext
                    Loop
                End If
                Call .Close
            End With
            Call .Fields.Delete("temp_client_login")
            Set idxTemp = .CreateIndex("client_login")
            Call idxTemp.Fields.Append(idxTemp.CreateField("client_login"))
            Call .Indexes.Append(idxTemp)
        End With
    
        ' Change the size of the client_login col of the journ_sessions table to 255 chars
        Set gtdfJournSessions = gdbsCybAdmin!journ_sessions
        With gtdfJournSessions
            Call .Fields.Append(.CreateField("temp_client_login", dbText, 20))
            .Fields("temp_client_login").AllowZeroLength = True
            Set grstJournSessions = gdbsCybAdmin.OpenRecordset("journ_sessions", dbOpenTable)
            With grstJournSessions
                If Not (.BOF And .EOF) Then
                    Call .MoveFirst
                    Do While Not .EOF
                        DoEvents
                        Call .Edit
                        .Fields("temp_client_login").value = .Fields("client_login").value
                        Call .Update
                        Call .MoveNext
                    Loop
                End If
                Call .Close
            End With
            Call .Indexes.Delete("client_login")
            Call .Fields.Delete("client_login")
            Call .Fields.Append(.CreateField("client_login", dbText, 0))
            .Fields("client_login").AllowZeroLength = True
            Set grstJournSessions = gdbsCybAdmin.OpenRecordset("journ_sessions", dbOpenTable)
            With grstJournSessions
                If Not (.BOF And .EOF) Then
                    Call .MoveFirst
                    Do While Not .EOF
                        DoEvents
                        Call .Edit
                        .Fields("client_login").value = .Fields("temp_client_login").value
                        Call .Update
                        Call .MoveNext
                    Loop
                End If
                Call .Close
            End With
            Call .Fields.Delete("temp_client_login")
            Set idxTemp = .CreateIndex("client_login")
            Call idxTemp.Fields.Append(idxTemp.CreateField("client_login"))
            Call .Indexes.Append(idxTemp)
        End With
    
    Case 20
        ' Change the size of the login col of the clients table to 255 chars
        Set gtdfClients = gdbsCybAdmin!clients
        With gtdfClients
            Call .Fields.Append(.CreateField("temp_login", dbText, 20))
            .Fields("temp_login").AllowZeroLength = True
            Set grstClients = gdbsCybAdmin.OpenRecordset("clients", dbOpenTable)
            With grstClients
                If Not (.BOF And .EOF) Then
                    Call .MoveFirst
                    Do While Not .EOF
                        DoEvents
                        Call .Edit
                        .Fields("temp_login").value = .Fields("login").value
                        Call .Update
                        Call .MoveNext
                    Loop
                End If
                Call .Close
            End With
            Call .Indexes.Delete("PrimaryKey")
            Call .Fields.Delete("login")
            Call .Fields.Append(.CreateField("login", dbText, 0))
            .Fields("login").AllowZeroLength = True
            Set grstClients = gdbsCybAdmin.OpenRecordset("clients", dbOpenTable)
            With grstClients
                If Not (.BOF And .EOF) Then
                    Call .MoveFirst
                    Do While Not .EOF
                        DoEvents
                        Call .Edit
                        .Fields("login").value = .Fields("temp_login").value
                        Call .Update
                        Call .MoveNext
                    Loop
                End If
                Call .Close
            End With
            Call .Fields.Delete("temp_login")
            Set idxTemp = .CreateIndex("PrimaryKey")
            Call idxTemp.Fields.Append(idxTemp.CreateField("login"))
            idxTemp.Primary = True
            Call .Indexes.Append(idxTemp)
        End With
        
    Case 21
        ' Add type_card type_client type_prepaid_ss type_postpaid_ss to the recharge table and init existing recharges
        Set gtdfRecharges = gdbsCybAdmin!recharges
        With gtdfRecharges
            Call .Fields.Append(.CreateField("type_card", dbBoolean))
            Call .Fields.Append(.CreateField("type_client", dbBoolean))
            Call .Fields.Append(.CreateField("type_prepaid_ss", dbBoolean))
            Call .Fields.Append(.CreateField("type_postpaid_ss", dbBoolean))
            Set grstRecharges = gdbsCybAdmin.OpenRecordset("recharges", dbOpenTable)
            With grstRecharges
                If Not (.BOF And .EOF) Then
                    Call .MoveFirst
                    Do While Not .EOF
                        DoEvents
                        Call .Edit
                        .Fields("type_postpaid_ss").value = .Fields("default").value
                        .Fields("type_card").value = True
                        .Fields("type_client").value = True
                        .Fields("type_prepaid_ss").value = True
                        Call .Update
                        Call .MoveNext
                    Loop
                End If
                Call .Close
            End With
            Call .Fields.Delete("default")
        End With
    
    Case 22
        ' Add validity_type col to the recharges table and set existing recharges to month validity
        Set gtdfRecharges = gdbsCybAdmin!recharges
        With gtdfRecharges
            Call .Fields.Append(.CreateField("validity_type", dbInteger))
            Set grstRecharges = gdbsCybAdmin.OpenRecordset("recharges", dbOpenTable)
            With grstRecharges
                If Not (.BOF And .EOF) Then
                    Call .MoveFirst
                    Do While Not .EOF
                        DoEvents
                        Call .Edit
                        .Fields("validity_type").value = enuRechargeValidityType.Months
                        Call .Update
                        Call .MoveNext
                    Loop
                End If
                Call .Close
            End With
        End With
    
    Case 23
        ' Add validity_type col to the cards_pending table and set existing pending cards to month validity
        Set gtdfCardsPending = gdbsCybAdmin!cards_pending
        With gtdfCardsPending
            Call .Fields.Append(.CreateField("validity_type", dbInteger))
            Set grstCardsPending = gdbsCybAdmin.OpenRecordset("cards_pending", dbOpenTable)
            With grstCardsPending
                If Not (.BOF And .EOF) Then
                    Call .MoveFirst
                    Do While Not .EOF
                        DoEvents
                        Call .Edit
                        .Fields("validity_type").value = enuRechargeValidityType.Months
                        Call .Update
                        Call .MoveNext
                    Loop
                End If
                Call .Close
            End With
        End With
    
    Case 24
        ' Add unlimited_play col to cards and set to false for existing cards
        Set gtdfCards = gdbsCybAdmin!cards
        With gtdfCards
            Call .Fields.Append(.CreateField("unlimited_play", dbBoolean))
            Set grstCards = gdbsCybAdmin.OpenRecordset("cards", dbOpenTable)
            With grstCards
                If Not (.BOF And .EOF) Then
                    Call .MoveFirst
                    Do While Not .EOF
                        DoEvents
                        Call .Edit
                        .Fields("unlimited_play").value = False
                        Call .Update
                        Call .MoveNext
                    Loop
                End If
                Call .Close
            End With
        End With
    
    Case 25
        ' Add unlimited_play col to clients and set to false for existing clients
        Set gtdfClients = gdbsCybAdmin!clients
        With gtdfClients
            Call .Fields.Append(.CreateField("unlimited_play", dbBoolean))
            Set grstClients = gdbsCybAdmin.OpenRecordset("clients", dbOpenTable)
            With grstClients
                If Not (.BOF And .EOF) Then
                    Call .MoveFirst
                    Do While Not .EOF
                        DoEvents
                        Call .Edit
                        .Fields("unlimited_play").value = False
                        Call .Update
                        Call .MoveNext
                    Loop
                End If
                Call .Close
            End With
        End With
    
    Case 26
        ' Add unlimited_play col to recharge and set to false for existing recharges
        Set gtdfRecharges = gdbsCybAdmin!recharges
        With gtdfRecharges
            Call .Fields.Append(.CreateField("unlimited_play", dbBoolean))
            Set grstRecharges = gdbsCybAdmin.OpenRecordset("recharges", dbOpenTable)
            With grstRecharges
                If Not (.BOF And .EOF) Then
                    Call .MoveFirst
                    Do While Not .EOF
                        DoEvents
                        Call .Edit
                        .Fields("unlimited_play").value = False
                        Call .Update
                        Call .MoveNext
                    Loop
                End If
                Call .Close
            End With
        End With
    
    Case 27
        ' Add unlimited_play col to cards_pending and set to false for existing pending cards
        Set gtdfCardsPending = gdbsCybAdmin!cards_pending
        With gtdfCardsPending
            Call .Fields.Append(.CreateField("unlimited_play", dbBoolean))
            Set grstCardsPending = gdbsCybAdmin.OpenRecordset("cards_pending", dbOpenTable)
            With grstCardsPending
                If Not (.BOF And .EOF) Then
                    Call .MoveFirst
                    Do While Not .EOF
                        DoEvents
                        Call .Edit
                        .Fields("unlimited_play").value = False
                        Call .Update
                        Call .MoveNext
                    Loop
                End If
                Call .Close
            End With
        End With
    
    Case 28
        ' Add unlimited_play col to journ_income and set to false for existing income entries
        Set gtdfJournIncome = gdbsCybAdmin!journ_income
        With gtdfJournIncome
            Call .Fields.Append(.CreateField("unlimited_play", dbBoolean))
            Set grstJournIncome = gdbsCybAdmin.OpenRecordset("journ_income", dbOpenTable)
            With grstJournIncome
                If Not (.BOF And .EOF) Then
                    Call .MoveFirst
                    Do While Not .EOF
                        DoEvents
                        Call .Edit
                        .Fields("unlimited_play").value = False
                        Call .Update
                        Call .MoveNext
                    Loop
                End If
                Call .Close
            End With
        End With
    
    Case 29
        ' Add usage_order col to recharge and set to 0 for existing recharges
        Set gtdfRecharges = gdbsCybAdmin!recharges
        With gtdfRecharges
            Call .Fields.Append(.CreateField("usage_order", dbInteger))
            Set grstRecharges = gdbsCybAdmin.OpenRecordset("recharges", dbOpenTable)
            With grstRecharges
                If Not (.BOF And .EOF) Then
                    Call .MoveFirst
                    Do While Not .EOF
                        DoEvents
                        Call .Edit
                        .Fields("usage_order").value = 0
                        Call .Update
                        Call .MoveNext
                    Loop
                End If
                Call .Close
            End With
        End With
    
    Case 30
        ' Add id col to recharge (autoIncrement)
        Set gtdfRecharges = gdbsCybAdmin!recharges
        With gtdfRecharges
            Set fldTemp = .CreateField("id", dbLong)
            fldTemp.Attributes = dbAutoIncrField
            Call .Fields.Append(fldTemp)
        End With
    
    Case 31
        ' set id col to primary key in recharges
        Set gtdfRecharges = gdbsCybAdmin!recharges
        With gtdfRecharges
            Set idxTemp = .CreateIndex("PrimaryKey")
            Call idxTemp.Fields.Append(idxTemp.CreateField("id"))
            idxTemp.Primary = True
            Call .Indexes.Append(idxTemp)
        End With
    
    Case 32
        ' Create extra indexes on recharges
        Set gtdfRecharges = gdbsCybAdmin!recharges
        With gtdfRecharges
            Set idxTemp = .CreateIndex("orderByTime")
            Call idxTemp.Fields.Append(idxTemp.CreateField("time"))
            Call idxTemp.Fields.Append(idxTemp.CreateField("price"))
            Call idxTemp.Fields.Append(idxTemp.CreateField("caption"))
            Call .Indexes.Append(idxTemp)
        
            Set idxTemp = .CreateIndex("orderByPrice")
            Call idxTemp.Fields.Append(idxTemp.CreateField("price"))
            Call idxTemp.Fields.Append(idxTemp.CreateField("time"))
            Call idxTemp.Fields.Append(idxTemp.CreateField("caption"))
            Call .Indexes.Append(idxTemp)
        
            Set idxTemp = .CreateIndex("orderByValidity")
            Call idxTemp.Fields.Append(idxTemp.CreateField("validity_type"))
            Call idxTemp.Fields.Append(idxTemp.CreateField("validity"))
            Call idxTemp.Fields.Append(idxTemp.CreateField("caption"))
            Call idxTemp.Fields.Append(idxTemp.CreateField("time"))
            Call idxTemp.Fields.Append(idxTemp.CreateField("price"))
            Call .Indexes.Append(idxTemp)
        
            Set idxTemp = .CreateIndex("orderByType")
            Call idxTemp.Fields.Append(idxTemp.CreateField("type_card"))
            Call idxTemp.Fields.Append(idxTemp.CreateField("type_client"))
            Call idxTemp.Fields.Append(idxTemp.CreateField("type_prepaid_ss"))
            Call idxTemp.Fields.Append(idxTemp.CreateField("type_postpaid_ss"))
            Call idxTemp.Fields.Append(idxTemp.CreateField("caption"))
            Call idxTemp.Fields.Append(idxTemp.CreateField("time"))
            Call idxTemp.Fields.Append(idxTemp.CreateField("price"))
            Call .Indexes.Append(idxTemp)
        End With
    
    Case 33
        ' Clean up recharges no zeros in validity => 1 day
        Set grstRecharges = gdbsCybAdmin.OpenRecordset("recharges", dbOpenTable)
        With grstRecharges
            If Not (.BOF And .EOF) Then
                Call .MoveFirst
                Do While Not .EOF
                    DoEvents
                    Call .Edit
                    If .Fields("validity").value = 0 Then
                        .Fields("validity").value = 1
                        .Fields("validity_type").value = enuRechargeValidityType.Days
                    End If
                    Call .Update
                    Call .MoveNext
                Loop
            End If
            Call .Close
        End With
    
    Case 34
        ' Add access_card_creation col to users and set default to True
        Set gtdfUsers = gdbsCybAdmin!Users
        Call gtdfUsers.Fields.Append(gtdfUsers.CreateField("access_card_creation", dbBoolean))
        
        Set grstUsers = gdbsCybAdmin.OpenRecordset("users", dbOpenTable)
        grstUsers.Index = "PrimaryKey"
        If Not (grstUsers.BOF And grstUsers.EOF) Then
            Call grstUsers.MoveFirst
            Do While Not grstUsers.EOF
                DoEvents
                Call grstUsers.Edit
                grstUsers!access_card_creation = True
                Call grstUsers.Update
                Call grstUsers.MoveNext
            Loop
        End If
        Call grstUsers.Close
    
    Case 35
        ' Add access_client_creation col to users and set default to True
        Set gtdfUsers = gdbsCybAdmin!Users
        Call gtdfUsers.Fields.Append(gtdfUsers.CreateField("access_client_creation", dbBoolean))
        
        Set grstUsers = gdbsCybAdmin.OpenRecordset("users", dbOpenTable)
        grstUsers.Index = "PrimaryKey"
        If Not (grstUsers.BOF And grstUsers.EOF) Then
            Call grstUsers.MoveFirst
            Do While Not grstUsers.EOF
                DoEvents
                Call grstUsers.Edit
                grstUsers!access_client_creation = True
                Call grstUsers.Update
                Call grstUsers.MoveNext
            Loop
        End If
        Call grstUsers.Close
    
    Case 36
        ' Change the size of the comment col of the clients table to 255 chars
        Set gtdfClients = gdbsCybAdmin!clients
        With gtdfClients
            Call .Fields.Append(.CreateField("temp_comment", dbText, 100))
            .Fields("temp_comment").AllowZeroLength = True
            Set grstClients = gdbsCybAdmin.OpenRecordset("clients", dbOpenTable)
            With grstClients
                If Not (.BOF And .EOF) Then
                    Call .MoveFirst
                    Do While Not .EOF
                        DoEvents
                        Call .Edit
                        .Fields("temp_comment").value = .Fields("comment").value
                        Call .Update
                        Call .MoveNext
                    Loop
                End If
                Call .Close
            End With
            Call .Fields.Delete("comment")
            Call .Fields.Append(.CreateField("comment", dbText, 0))
            .Fields("comment").AllowZeroLength = True
            Set grstClients = gdbsCybAdmin.OpenRecordset("clients", dbOpenTable)
            With grstClients
                If Not (.BOF And .EOF) Then
                    Call .MoveFirst
                    Do While Not .EOF
                        DoEvents
                        Call .Edit
                        .Fields("comment").value = .Fields("temp_comment").value
                        Call .Update
                        Call .MoveNext
                    Loop
                End If
                Call .Close
            End With
            Call .Fields.Delete("temp_comment")
        End With
    
    End Select
    funDBUpgradeSteps = intBaseVersion + 1
    Call grstConfig.MoveFirst
    Call grstConfig.Edit
    grstConfig!Version = funDBUpgradeSteps
    Call grstConfig.Update
    Exit Function
    
funDBUpgradeStepsError:
Dim strDBErrorMsg As String
    funDBUpgradeSteps = 0
    strDBErrorMsg = Err.Description
    Select Case Err.Number
    Case 3011
        If funMsg(enuMsg.DBVersionError) = vbYes Then Call funShutdownApplication
    Case Else
        Call funErrorManager("modFunctionsDB.funDBUpgradeStepsError", False, True)
        If funMsg(enuMsg.GeneralDBError, strDBErrorMsg) = vbYes Then Call funShutdownApplication
    End Select
End Function



