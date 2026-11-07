Attribute VB_Name = "modStats"
' _____________________________________________________________________________
'|________________________________ STATISTICS FUNCTION MODULE _________________|
'|  Name:          modStats.bas                                                |
'|  Type:          VB6 Module                                                  |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
'|  Description:   Contains statistical functions ans subs                     |
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

' $Id: modStats.bas 15 2006-01-01 16:12:35Z sam $

Option Explicit
Public gblnStopStats As Boolean


'*******************************************************************************
Public Sub subProcessSessionStats(dteMinDate As Date, dteMaxDate As Date)
Dim dteMinRealDate As Date
Dim dteMaxRealDate As Date
Dim dteMinSessionDate As Date
Dim dteMaxSessionDate As Date
Dim dteMinEventDate As Date
Dim dteMaxEventDate As Date
Dim lngNumDays As Long
Dim dteCurrentDate As Date
Dim dteLastDateStamp As Date
Dim dteTodayDate As Date
Dim dteTodayDateTime As Date
Dim lngCurrentNumWs As Long
Dim blnInetConnected As Boolean
Dim blnAppStopped As Boolean
Dim blnDayStarted As Boolean
Dim i As Long, j As Long, k As Long
On Error GoTo subProcessSessionStatsError

    If Not (grstJournSessions.BOF = True And grstJournSessions.EOF = True) Then
        Call grstJournSessions.MoveFirst
        dteMinSessionDate = grstJournSessions!date_start
        Call grstJournSessions.MoveLast
        dteMaxSessionDate = grstJournSessions!date_start
    Else
        Exit Sub
    End If
    If Not (grstEvents.BOF = True And grstEvents.EOF = True) Then
        Call grstEvents.MoveFirst
        dteMinEventDate = grstEvents!date_stamp
        Call grstEvents.MoveLast
        dteMaxEventDate = grstEvents!date_stamp
    End If
    dteMinRealDate = IIf(dteMinSessionDate > dteMinEventDate, dteMinEventDate, dteMinSessionDate)
    dteMaxRealDate = IIf(dteMaxSessionDate < dteMaxEventDate, dteMaxEventDate, dteMaxSessionDate)
    
    dteMinRealDate = IIf(dteMinSessionDate > dteMinDate, dteMinSessionDate, dteMinDate)
    dteMaxRealDate = IIf(dteMaxSessionDate < dteMaxDate, dteMaxSessionDate, dteMaxDate)
    dteMaxRealDate = IIf(dteMaxRealDate < dteMinRealDate, dteMinRealDate, dteMaxRealDate)
    
    lngNumDays = DateDiff("d", dteMinRealDate, dteMaxRealDate) + 1
    ReDim guStatsDay(1 To lngNumDays)
    dteTodayDateTime = Now
    dteTodayDate = CDate(DateValue(dteTodayDateTime) & " 00:00:00")
    dteCurrentDate = CDate(DateValue(dteMinRealDate) & " 00:00:00")
    blnAppStopped = True
    
    DoEvents
    Call subQueryEvent(dteMinRealDate, DateAdd("d", 1, dteMaxRealDate))
    If grstEventsQry.RecordCount > 0 Then dteLastDateStamp = grstEventsQry!date_stamp
    
    For i = 1 To lngNumDays
        With guStatsDay(i)
        
            Call subQuerySession(dteCurrentDate, dteCurrentDate)
            .lngSessionNum = grstJournSessionsQry.RecordCount
            For j = 1 To grstJournSessionsQry.RecordCount
                .dteSessionTotal = .dteSessionTotal + IIf(IsNull(grstJournSessionsQry!time_connected), NULL_DATE, grstJournSessionsQry!time_connected) + IIf(IsNull(grstJournSessionsQry!time_connected_offpeak), NULL_DATE, grstJournSessionsQry!time_connected_offpeak) + IIf(IsNull(grstJournSessionsQry!time_disconnected), NULL_DATE, grstJournSessionsQry!time_disconnected)
                grstJournSessionsQry.MoveNext
            Next j
        
            blnDayStarted = False
            .dteDate = dteCurrentDate
            .uTimeOpen.dteStart = dteCurrentDate
            If blnInetConnected Then .uTimeInetCnx.dteStart = dteCurrentDate
            For k = 0 To lngCurrentNumWs
                .uTimeWs(k).dteStart = dteCurrentDate
            Next k
            If Not grstEventsQry.EOF Then
                Do While grstEventsQry!date_stamp < DateAdd("d", 1, dteCurrentDate)
                
                    If Not blnDayStarted Then
                        If grstEventsQry!Type <> enuEvents.AppStart Then
                            .uTimeOpen.dteStart = dteCurrentDate
                            .uTimeWs(lngCurrentNumWs).dteStart = dteCurrentDate
                            If blnInetConnected Then .uTimeInetCnx.dteStart = dteCurrentDate
                        Else
                            blnAppStopped = True
                            blnInetConnected = False
                        End If
                    End If
                    blnDayStarted = True
                    Select Case grstEventsQry!Type
                    Case enuEvents.AppStart
                        If grstEventsQry!session_id = 0 And grstEventsQry!income_id = 0 And grstEventsQry!card_id = 0 Then
                            If Not blnAppStopped Then
                                .uTimeOpen.dteTotal = .uTimeOpen.dteTotal + .uTimeOpen.dteStart - dteLastDateStamp
                                .uTimeWs(lngCurrentNumWs).dteTotal = .uTimeWs(lngCurrentNumWs).dteTotal + .uTimeWs(lngCurrentNumWs).dteStart - dteLastDateStamp
                                If blnInetConnected Then .uTimeInetCnx.dteTotal = .uTimeInetCnx.dteTotal + .uTimeInetCnx.dteStart - dteLastDateStamp
                            End If
                            
                            lngCurrentNumWs = 0
                            .uTimeOpen.dteStart = grstEventsQry!date_stamp
                            .uTimeWs(lngCurrentNumWs).dteStart = grstEventsQry!date_stamp
                            blnInetConnected = False
                            blnAppStopped = False
                        Else
                        Debug.Print "test"
                        
                        End If
                        
                    Case enuEvents.AppStop
                        .uTimeOpen.dteTotal = .uTimeOpen.dteTotal + .uTimeOpen.dteStart - grstEventsQry!date_stamp
                        .uTimeWs(lngCurrentNumWs).dteTotal = .uTimeWs(lngCurrentNumWs).dteTotal + .uTimeWs(lngCurrentNumWs).dteStart - grstEventsQry!date_stamp
                        If blnInetConnected Then .uTimeInetCnx.dteTotal = .uTimeInetCnx.dteTotal + .uTimeInetCnx.dteStart - grstEventsQry!date_stamp
                        blnInetConnected = False
                        blnAppStopped = True
                    
                    Case enuEvents.AppLogon
                        If blnAppStopped Then
                            .uTimeOpen.dteStart = grstEventsQry!date_stamp
                            .uTimeWs(lngCurrentNumWs).dteStart = grstEventsQry!date_stamp
                            blnAppStopped = False
                        End If
                    
                    Case enuEvents.InternetConnected
                        If Not blnInetConnected Then .uTimeInetCnx.dteStart = grstEventsQry!date_stamp
                        blnInetConnected = True
                        
                    Case enuEvents.InternetDiconnected
                        If blnInetConnected Then .uTimeInetCnx.dteTotal = .uTimeInetCnx.dteTotal + .uTimeInetCnx.dteStart - grstEventsQry!date_stamp
                        blnInetConnected = False
                        
                    Case enuEvents.SessionOpen, enuEvents.SessionResume
                        If blnAppStopped Then
                            lngCurrentNumWs = 0
                            .uTimeOpen.dteStart = grstEventsQry!date_stamp
                            .uTimeWs(lngCurrentNumWs).dteStart = grstEventsQry!date_stamp
                            blnInetConnected = False
                            blnAppStopped = False
                        End If
                        .uTimeWs(lngCurrentNumWs).dteTotal = .uTimeWs(lngCurrentNumWs).dteTotal + .uTimeWs(lngCurrentNumWs).dteStart - grstEventsQry!date_stamp
                        lngCurrentNumWs = lngCurrentNumWs + 1
                        If lngCurrentNumWs > MAX_NUM_WORKSTATIONS - 1 Then
                            Exit For
                        End If
                        .uTimeWs(lngCurrentNumWs).dteStart = grstEventsQry!date_stamp
                    
                    Case enuEvents.SessionClose
                        If lngCurrentNumWs > 0 Then
                            .uTimeWs(lngCurrentNumWs).dteTotal = .uTimeWs(lngCurrentNumWs).dteTotal + .uTimeWs(lngCurrentNumWs).dteStart - grstEventsQry!date_stamp
                            lngCurrentNumWs = lngCurrentNumWs - 1
                            .uTimeWs(lngCurrentNumWs).dteStart = grstEventsQry!date_stamp
                        End If
                        
                    End Select
                                   
                    dteLastDateStamp = grstEventsQry!date_stamp
                    grstEventsQry.MoveNext
                
                    If grstEventsQry.EOF Then Exit Do
                    DoEvents
                Loop
                
                If Not grstEventsQry.EOF Then
                    If grstEventsQry!Type <> enuEvents.AppStart Then
                        .uTimeOpen.dteTotal = .uTimeOpen.dteTotal + .uTimeOpen.dteStart - CDate(DateValue(dteCurrentDate) & " 23:59:59")
                        .uTimeWs(lngCurrentNumWs).dteTotal = .uTimeWs(lngCurrentNumWs).dteTotal + .uTimeWs(lngCurrentNumWs).dteStart - CDate(DateValue(dteCurrentDate) & " 23:59:59")
                        If blnInetConnected Then .uTimeInetCnx.dteTotal = .uTimeInetCnx.dteTotal + .uTimeInetCnx.dteStart - CDate(DateValue(dteCurrentDate) & " 23:59:59")
                    End If
                End If
            End If
            
            If Not blnAppStopped And dteCurrentDate = dteTodayDate Then
                .uTimeOpen.dteTotal = .uTimeOpen.dteTotal + .uTimeOpen.dteStart - dteTodayDateTime
                If blnInetConnected Then .uTimeInetCnx.dteTotal = .uTimeInetCnx.dteTotal + .uTimeInetCnx.dteStart - dteTodayDateTime
                .uTimeWs(lngCurrentNumWs).dteTotal = .uTimeWs(lngCurrentNumWs).dteTotal + .uTimeWs(lngCurrentNumWs).dteStart - dteTodayDateTime
            End If
            For j = 0 To guServOpts.uWorkstations.lngMaxNumWorkstations
                .dteTimeWsTotal = .dteTimeWsTotal + .uTimeWs(j).dteTotal
            Next j
        
        End With
        dteCurrentDate = DateAdd("d", 1, dteCurrentDate)
        DoEvents
        If gblnStopStats Then Exit Sub
        frmStatsProgress.pgbStatProgress.value = frmStatsProgress.pgbStatProgress.Max * 0.8 / lngNumDays * i
        frmStatsProgress.pgbStatType.value = frmStatsProgress.pgbStatType.Max * 0.2 * 0.8 / lngNumDays * i
    Next i
    Exit Sub
    
subProcessSessionStatsError:
    Call funErrorManager("modStats.subProcessSessionStatsError")
End Sub


'*******************************************************************************
Public Sub subProcessSessionTimes(dteMinDate As Date, dteMaxDate As Date)
Dim dteMinRealDate As Date
Dim dteMaxRealDate As Date
Dim dteMinSessionDate As Date
Dim dteMaxSessionDate As Date
Dim dteMinEventDate As Date
Dim dteMaxEventDate As Date
Dim lngNumDays As Long
Dim dteCurrentDate As Date
Dim dteTimeIntervals() As Date
Dim intNumIntervals As Integer
Dim dteStartTime As Date
Dim dteStopTime As Date
Dim i As Long, j As Long, k As Long
On Error GoTo subProcessSessionTimesError

    If Not (grstJournSessions.BOF = True And grstJournSessions.EOF = True) Then
        Call grstJournSessions.MoveFirst
        dteMinSessionDate = grstJournSessions!date_start
        Call grstJournSessions.MoveLast
        dteMaxSessionDate = grstJournSessions!date_start
    Else
        Exit Sub
    End If
    If Not (grstEvents.BOF = True And grstEvents.EOF = True) Then
        Call grstEvents.MoveFirst
        dteMinEventDate = grstEvents!date_stamp
        Call grstEvents.MoveLast
        dteMaxEventDate = grstEvents!date_stamp
    End If
    dteMinRealDate = IIf(dteMinSessionDate > dteMinEventDate, dteMinEventDate, dteMinSessionDate)
    dteMaxRealDate = IIf(dteMaxSessionDate < dteMaxEventDate, dteMaxEventDate, dteMaxSessionDate)
    
    dteMinRealDate = IIf(dteMinSessionDate > dteMinDate, dteMinSessionDate, dteMinDate)
    dteMaxRealDate = IIf(dteMaxSessionDate < dteMaxDate, dteMaxSessionDate, dteMaxDate)
    dteMaxRealDate = IIf(dteMaxRealDate < dteMinRealDate, dteMinRealDate, dteMaxRealDate)
    
    lngNumDays = DateDiff("d", dteMinRealDate, dteMaxRealDate) + 1
    ReDim guStatsTimes(1 To lngNumDays)
    intNumIntervals = funArraySize(guServOpts.uStats.strStatsTimeInt)
    ReDim dteTimeIntervals(intNumIntervals + 1)
    dteTimeIntervals(0) = NULL_DATE
    For i = 1 To intNumIntervals
        dteTimeIntervals(i) = funDateTimefromNet(guServOpts.uStats.strStatsTimeInt(i - 1))
    Next i
    dteTimeIntervals(intNumIntervals + 1) = CDate("23:59:59")
    
    dteCurrentDate = CDate(DateValue(dteMinRealDate) & " 00:00:00")
    For i = 1 To lngNumDays
        With guStatsTimes(i)
            ReDim .lngTimeInt(intNumIntervals + 1)
            .dteDate = dteCurrentDate
            .lngSessionCount = 0
            For k = 0 To intNumIntervals
                .lngTimeInt(k) = 0
            Next k
            Call subQuerySession(dteCurrentDate, dteCurrentDate)
            .lngSessionNum = grstJournSessionsQry.RecordCount
            For j = 1 To grstJournSessionsQry.RecordCount
                dteStartTime = grstJournSessionsQry!date_start
                dteStartTime = CDate(format(dteStartTime, "Hh:Nn:Ss"))
                dteStopTime = IIf(IsNull(grstJournSessionsQry!date_end), IIf(IsNull(grstJournSessionsQry!date_start), NULL_DATE, grstJournSessionsQry!date_start) + IIf(IsNull(grstJournSessionsQry!time_connected), NULL_DATE, grstJournSessionsQry!time_connected) + IIf(IsNull(grstJournSessionsQry!time_disconnected), NULL_DATE, grstJournSessionsQry!time_disconnected) + IIf(IsNull(grstJournSessionsQry!time_connected_offpeak), NULL_DATE, grstJournSessionsQry!time_connected_offpeak), grstJournSessionsQry!date_end)
                dteStopTime = CDate(format(dteStopTime, "Hh:Nn:Ss"))
                
                For k = 0 To intNumIntervals
                    If dteStartTime >= dteTimeIntervals(k) And dteStartTime <= dteTimeIntervals(k + 1) Or dteStopTime >= dteTimeIntervals(k) And dteStopTime <= dteTimeIntervals(k + 1) Or dteStartTime <= dteTimeIntervals(k) And dteStopTime >= dteTimeIntervals(k + 1) Then
                        .lngTimeInt(k) = .lngTimeInt(k) + 1
                        .lngSessionCount = .lngSessionCount + 1
                    End If
                Next k
                grstJournSessionsQry.MoveNext
            Next j
        End With
        dteCurrentDate = DateAdd("d", 1, dteCurrentDate)
        DoEvents
        If gblnStopStats Then Exit Sub
        frmStatsProgress.pgbStatProgress.value = frmStatsProgress.pgbStatProgress.Max * 0.8 / lngNumDays * i
        frmStatsProgress.pgbStatType.value = frmStatsProgress.pgbStatType.Max * 0.2 + frmStatsProgress.pgbStatType.Max * 0.2 * 0.8 / lngNumDays * i
    Next i
    Exit Sub
    
subProcessSessionTimesError:
    Call funErrorManager("modStats.subProcessSessionTimesError")
End Sub



'*******************************************************************************
Public Sub subExportIncome()
Dim intLogFileNumber As Integer
Dim strFilePath As String
Dim strExportLine As String
Dim i As Long, j As Long
On Error GoTo subExportIncomeError

    frmMain.cdlIncomeExport.flags = cdlOFNOverwritePrompt Or cdlOFNExplorer
    frmMain.cdlIncomeExport.DialogTitle = funLoadResString(10321)
    frmMain.cdlIncomeExport.filename = guServOpts.uConfig.strLastExportPath & format(Now, "yyyymmdd") & "_" & LCase(funLoadResString(10005) & "-" & funLoadResString(10294)) & ".csv"
    frmMain.cdlIncomeExport.filter = funLoadResString(10322) & " (*.csv)|*.csv"
    frmMain.cdlIncomeExport.CancelError = True
    Call frmMain.cdlIncomeExport.ShowSave
    
    strFilePath = frmMain.cdlIncomeExport.filename
    guServOpts.uConfig.strLastExportPath = left(strFilePath, Len(strFilePath) - Len(frmMain.cdlIncomeExport.FileTitle))
    intLogFileNumber = FreeFile
    Open strFilePath For Output As intLogFileNumber
    With frmMain.fxgIncome
        strExportLine = funLoadResString(10000) & CSV_SEPARATOR & funLoadResString(10327) & vbNewLine
        strExportLine = strExportLine & funLoadResString(10299) & CSV_SEPARATOR & .TextMatrix(1, 0) & vbNewLine
        strExportLine = strExportLine & funLoadResString(10300) & CSV_SEPARATOR & .TextMatrix(.Rows - 3, 0) & vbNewLine
        Print #intLogFileNumber, strExportLine
        Close intLogFileNumber
        Open strFilePath For Append As intLogFileNumber
        For i = 0 To .Rows - 1
            strExportLine = .TextMatrix(i, 0) & CSV_SEPARATOR & Replace(.TextMatrix(i, 1), " ", "")
            For j = 2 To .Cols - 1
                strExportLine = strExportLine & CSV_SEPARATOR & .TextMatrix(i, j)
            Next j
            Print #intLogFileNumber, strExportLine
        Next i
    End With
    Close intLogFileNumber
    Exit Sub
    
subExportIncomeError:
    Select Case Err.Number
    Case cdlCancel
        Exit Sub
    
    Case 70
        Call funMsg(ExportImpossible, strFilePath)
        Call subExportIncome
    
    Case Else
        Call funErrorManager("modStats.subExportIncomeError")
    
    End Select
End Sub


'*******************************************************************************
Public Sub subExportSessionStats()
Dim intLogFileNumber As Integer
Dim strFilePath As String
Dim strExportLine As String
Dim i As Long, j As Long
On Error GoTo subExportSessionStatsError

    With frmMain.cdlSessionExport
        .flags = cdlOFNOverwritePrompt Or cdlOFNExplorer
        .DialogTitle = funLoadResString(10413)
        .filename = guServOpts.uConfig.strLastExportPath & format(Now, "yyyymmdd") & "_" & LCase(funLoadResString(10005) & "-" & funLoadResString(10414)) & ".csv"
        .filter = funLoadResString(10322) & " (*.csv)|*.csv"
        .CancelError = True
        Call .ShowSave
        strFilePath = .filename
    
        guServOpts.uConfig.strLastExportPath = left(strFilePath, Len(strFilePath) - Len(.FileTitle))
    End With
    intLogFileNumber = FreeFile
    Open strFilePath For Output As intLogFileNumber
    
    With frmMain.fxgSession(1)
    
        strExportLine = funLoadResString(10000) & CSV_SEPARATOR & funLoadResString(10415) & vbNewLine
        strExportLine = strExportLine & funLoadResString(10299) & CSV_SEPARATOR & .TextMatrix(1, 0) & vbNewLine
        strExportLine = strExportLine & funLoadResString(10300) & CSV_SEPARATOR & .TextMatrix(.Rows - 3, 0) & vbNewLine
        Print #intLogFileNumber, strExportLine
        Close intLogFileNumber
        Open strFilePath For Append As intLogFileNumber
                
        
        For i = 0 To .Rows - 1
            strExportLine = .TextMatrix(i, 0)
            For j = 1 To 7
                strExportLine = strExportLine & CSV_SEPARATOR & .TextMatrix(i, j)
            Next j
            For j = 1 To guServOpts.uWorkstations.lngNumWorkstations * 2 Step 2
                strExportLine = strExportLine & CSV_SEPARATOR & .TextMatrix(i, j + 7)
                strExportLine = strExportLine & CSV_SEPARATOR & .TextMatrix(i, j + 8)
            Next j
            Print #intLogFileNumber, strExportLine
        Next i
    End With
    Close intLogFileNumber
    Exit Sub
    
subExportSessionStatsError:
    Select Case Err.Number
    Case cdlCancel
        Exit Sub
    
    Case 70
        Call funMsg(ExportImpossible, strFilePath)
        Call subExportSessionStats
    
    Case Else
        Call funErrorManager("modStats.subExportSessionStatsError")
    
    End Select
End Sub


'*******************************************************************************
Public Sub subExportSessionList()
Dim intLogFileNumber As Integer
Dim strFilePath As String
Dim strExportLine As String
Dim i As Long, j As Long
On Error GoTo subExportSessionListError

    With frmMain.cdlSessionExport
        .flags = cdlOFNOverwritePrompt Or cdlOFNExplorer
        .DialogTitle = funLoadResString(10416)
        .filename = guServOpts.uConfig.strLastExportPath & format(Now, "yyyymmdd") & "_" & LCase(funLoadResString(10005) & "-" & funLoadResString(10417)) & ".csv"
        .filter = funLoadResString(10322) & " (*.csv)|*.csv"
        .CancelError = True
        Call .ShowSave
        strFilePath = .filename
    
        guServOpts.uConfig.strLastExportPath = left(strFilePath, Len(strFilePath) - Len(.FileTitle))
    End With
    intLogFileNumber = FreeFile
    Open strFilePath For Output As intLogFileNumber
    
    With frmMain.fxgSession(3)
    
        strExportLine = funLoadResString(10000) & CSV_SEPARATOR & funLoadResString(10418) & vbNewLine
        strExportLine = strExportLine & funLoadResString(10299) & CSV_SEPARATOR & .TextMatrix(1, 0) & vbNewLine
        strExportLine = strExportLine & funLoadResString(10300) & CSV_SEPARATOR & .TextMatrix(.Rows - 3, 0) & vbNewLine
        Print #intLogFileNumber, strExportLine
        Close intLogFileNumber
        Open strFilePath For Append As intLogFileNumber
                
        
        For i = 0 To .Rows - 1
            strExportLine = .TextMatrix(i, 0)
            For j = 1 To .Cols - 1
                strExportLine = strExportLine & CSV_SEPARATOR & .TextMatrix(i, j)
            Next j
            Print #intLogFileNumber, strExportLine
        Next i
    End With
    Close intLogFileNumber
    Exit Sub
    
subExportSessionListError:
    Select Case Err.Number
    Case cdlCancel
        Exit Sub
    
    Case 70
        Call funMsg(ExportImpossible, strFilePath)
        Call subExportSessionList
    
    Case Else
        Call funErrorManager("modStats.subExportSessionListError")
    
    End Select
End Sub


'*******************************************************************************
Public Sub subExportSessionURLs()
Dim intLogFileNumber As Integer
Dim strFilePath As String
Dim strExportLine As String
Dim i As Long, j As Long
On Error GoTo subExportSessionListError

    With frmMain.cdlSessionExport
        .flags = cdlOFNOverwritePrompt Or cdlOFNExplorer
        .DialogTitle = funLoadResString(10701)
        .filename = guServOpts.uConfig.strLastExportPath & format(Now, "yyyymmdd") & "_" & LCase(funLoadResString(10005) & "-" & funLoadResString(10702)) & ".csv"
        .filter = funLoadResString(10322) & " (*.csv)|*.csv"
        .CancelError = True
        Call .ShowSave
        strFilePath = .filename
    
        guServOpts.uConfig.strLastExportPath = left(strFilePath, Len(strFilePath) - Len(.FileTitle))
    End With
    intLogFileNumber = FreeFile
    Open strFilePath For Output As intLogFileNumber
    
    With frmMain.fxgSession(4)
    
        strExportLine = funLoadResString(10000) & CSV_SEPARATOR & funLoadResString(10703) & vbNewLine
        strExportLine = strExportLine & funLoadResString(10299) & CSV_SEPARATOR & .TextMatrix(1, 0) & vbNewLine
        strExportLine = strExportLine & funLoadResString(10300) & CSV_SEPARATOR & .TextMatrix(.Rows - 1, 0) & vbNewLine
        Print #intLogFileNumber, strExportLine
        Close intLogFileNumber
        Open strFilePath For Append As intLogFileNumber
                
        
        For i = 0 To .Rows - 1
            strExportLine = .TextMatrix(i, 0)
            For j = 1 To .Cols - 1
                strExportLine = strExportLine & CSV_SEPARATOR & .TextMatrix(i, j)
            Next j
            Print #intLogFileNumber, strExportLine
        Next i
    End With
    Close intLogFileNumber
    Exit Sub
    
subExportSessionListError:
    Select Case Err.Number
    Case cdlCancel
        Exit Sub
    
    Case 70
        Call funMsg(ExportImpossible, strFilePath)
        Call subExportSessionList
    
    Case Else
        Call funErrorManager("modStats.subExportSessionListError")
    
    End Select
End Sub


'*******************************************************************************
Public Sub subExportSessionTimes()
Dim intLogFileNumber As Integer
Dim strFilePath As String
Dim strExportLine As String
Dim i As Long, j As Long
On Error GoTo subExportSessionTimesError

    With frmMain.cdlSessionExport
        .flags = cdlOFNOverwritePrompt Or cdlOFNExplorer
        .DialogTitle = funLoadResString(10427)
        .filename = guServOpts.uConfig.strLastExportPath & format(Now, "yyyymmdd") & "_" & LCase(funLoadResString(10005) & "-" & funLoadResString(10428)) & ".csv"
        .filter = funLoadResString(10322) & " (*.csv)|*.csv"
        .CancelError = True
        Call .ShowSave
        strFilePath = .filename
    
        guServOpts.uConfig.strLastExportPath = left(strFilePath, Len(strFilePath) - Len(.FileTitle))
    End With
    intLogFileNumber = FreeFile
    Open strFilePath For Output As intLogFileNumber
    
    With frmMain.fxgSession(2)
    
        strExportLine = funLoadResString(10000) & CSV_SEPARATOR & funLoadResString(10429) & vbNewLine
        strExportLine = strExportLine & funLoadResString(10299) & CSV_SEPARATOR & .TextMatrix(1, 0) & vbNewLine
        strExportLine = strExportLine & funLoadResString(10300) & CSV_SEPARATOR & .TextMatrix(.Rows - 3, 0) & vbNewLine
        Print #intLogFileNumber, strExportLine
        Close intLogFileNumber
        Open strFilePath For Append As intLogFileNumber
                
        
        For i = 0 To .Rows - 1
            strExportLine = .TextMatrix(i, 0)
            For j = 1 To .Cols - 1
                strExportLine = strExportLine & CSV_SEPARATOR & .TextMatrix(i, j)
            Next j
            Print #intLogFileNumber, strExportLine
        Next i
    End With
    Close intLogFileNumber
    Exit Sub
    
subExportSessionTimesError:
    Select Case Err.Number
    Case cdlCancel
        Exit Sub
    
    Case 70
        Call funMsg(ExportImpossible, strFilePath)
        Call subExportSessionTimes
    
    Case Else
        Call funErrorManager("modStats.subExportSessionTimesError")
    
    End Select
End Sub


'*******************************************************************************
Public Sub subFillIncome(dteIncomeStart As Date, dteIncomeEnd As Date, strTypes As String)
Dim i As Long, j As Long
Dim sngIncomeTotal As Single
Dim dteTimeTotal As Date
Dim strTmp As String
Dim rstJournIncomeQry As DAO.Recordset
On Error GoTo subFillIncomeError

    If Not gblnDBOK Then Exit Sub
    Call subQueryJournIncome(dteIncomeStart, dteIncomeEnd, strTypes, rstJournIncomeQry)
    With frmMain.fxgIncome
        .Rows = 2
        For i = 1 To .Rows - 1
            For j = 0 To .Cols - 1
                .TextMatrix(i, j) = ""
            Next j
        Next i
        If rstJournIncomeQry.BOF = True And rstJournIncomeQry.EOF = True Then
            .Rows = 2
            For i = 1 To .Rows - 1
                .TextMatrix(i, 0) = ""
                .TextMatrix(i, 1) = ""
                .TextMatrix(i, 2) = ""
                .TextMatrix(i, 3) = ""
                .TextMatrix(i, 4) = ""
                .TextMatrix(i, 5) = ""
                .TextMatrix(i, 6) = ""
                .TextMatrix(i, 7) = ""
            Next i
            frmMain.cmdIncomeExport.enabled = False
            frmMain.cmdIncomePrint.enabled = False
            frmMain.mnuIncomeExport.enabled = False
            frmMain.mnuIncomePrint.enabled = False
        Else
            Call rstJournIncomeQry.MoveFirst
            .Rows = rstJournIncomeQry.RecordCount + 2 + IIf(frmMain.chkTypeIncome(3).value = vbUnchecked, 1, 0)
            .enabled = True
            For i = 1 To rstJournIncomeQry.RecordCount
                sngIncomeTotal = sngIncomeTotal + rstJournIncomeQry!income
                .TextMatrix(i, 0) = rstJournIncomeQry!date_stamp
                .TextMatrix(i, 1) = funFormatCurrencyString(rstJournIncomeQry!income, guServOpts.uConfig.strMonFormat)
                .TextMatrix(i, 2) = IIf(IsNull(rstJournIncomeQry!Label), "", rstJournIncomeQry!Label)
                Select Case rstJournIncomeQry!Type
                Case enuIncomeType.Selfservice
                    .TextMatrix(i, 3) = funLoadResString(10313)
                Case enuIncomeType.Card
                    .TextMatrix(i, 3) = funLoadResString(10314)
                Case enuIncomeType.Other
                    .TextMatrix(i, 3) = funLoadResString(10315)
                Case enuIncomeType.Client
                    .TextMatrix(i, 3) = funLoadResString(10831)
                Case enuIncomeType.SelfServicePrePaid
                    .TextMatrix(i, 3) = funLoadResString(10986)
                Case Else
                    .TextMatrix(i, 3) = funLoadResString(10316)
                End Select
                .TextMatrix(i, 4) = IIf(rstJournIncomeQry!Type = enuIncomeType.Card, IIf(IsNull(rstJournIncomeQry!card_id), "", rstJournIncomeQry!card_id), IIf(IsNull(rstJournIncomeQry!client_login), "", rstJournIncomeQry!client_login))
                dteTimeTotal = dteTimeTotal + IIf(IsNull(rstJournIncomeQry!Time), NULL_DATE, rstJournIncomeQry!Time)
                strTmp = IIf(rstJournIncomeQry!unlimited_play, funLoadResString(11000), funFormatTime(IIf(IsNull(rstJournIncomeQry!Time), NULL_DATE, rstJournIncomeQry!Time)))
                .TextMatrix(i, 5) = IIf(strTmp <> "00:00:00", strTmp, "")
                If rstJournIncomeQry!paid Then
                    .TextMatrix(i, 6) = funLoadResString(10311)
                Else
                    .TextMatrix(i, 6) = funLoadResString(10312)
                End If
                .TextMatrix(i, 7) = IIf(IsNull(rstJournIncomeQry!login), "", rstJournIncomeQry!login)
                Call rstJournIncomeQry.MoveNext
                DoEvents
            Next i
            i = rstJournIncomeQry.RecordCount + 1
            .Row = i
            .Col = 0
            .ColSel = .Cols - 1
            .CellBackColor = vbButtonShadow
            .TextMatrix(i, 0) = funLoadResString(10387) & " (" & rstJournIncomeQry.RecordCount & funLoadResString(10303) & ")"
            .TextMatrix(i, 1) = funFormatCurrencyString(sngIncomeTotal, guServOpts.uConfig.strMonFormat)
            .TextMatrix(i, 5) = funFormatTime(dteTimeTotal)
                        
            If frmMain.chkTypeIncome(3).value = vbUnchecked Then
                i = rstJournIncomeQry.RecordCount + 2
                .Row = i
                .Col = 0
                .ColSel = .Cols - 1
                .CellBackColor = vbButtonFace
                .TextMatrix(i, 0) = funLoadResString(10654) ' & " (" & rstJournIncomeQry.RecordCount & funLoadResString(10303) & ")"
                .TextMatrix(i, 1) = funFormatCurrencyString(sngIncomeTotal / funNumHours(dteTimeTotal), guServOpts.uConfig.strMonFormat)
            
            
            End If
            
            frmMain.cmdIncomeExport.enabled = True
            frmMain.cmdIncomePrint.enabled = guServOpts.uConfig.strDefaultPrinterName <> ""
            frmMain.mnuIncomeExport.enabled = True
            frmMain.mnuIncomePrint.enabled = True
        End If
        .TopRow = .Rows - 1
    End With
    Exit Sub
    
subFillIncomeError:
    Call funErrorManager("modStats.subFillIncomeError")
End Sub


'*******************************************************************************
Public Sub subFillSessionList()
Dim i As Long, j As Long
Dim dteSessionTime As Date
Dim dteTotalSessionTime As Date
Dim dteTotalCon As Date
Dim dteTotalDiscon As Date
Dim dteTotalConOffpeak As Date
Dim lngNumTimeConnected As Long
Dim lngNumTimeDisconnected As Long
Dim lngNumTimeConnectedOffPeak As Long
On Error GoTo subFillSessionListError

    If Not gblnDBOK Then Exit Sub
    With frmMain.fxgSession(3)
        .Rows = 2
        For i = 1 To .Rows - 1
            For j = 0 To .Cols - 1
                .TextMatrix(i, j) = ""
            Next j
        Next i
        If grstJournSessionsQry.BOF = True And grstJournSessionsQry.EOF = True Then
            .Rows = 2
        Else
            Call grstJournSessionsQry.MoveFirst
            .Rows = grstJournSessionsQry.RecordCount + 3
            .enabled = True
            For i = 1 To grstJournSessionsQry.RecordCount
                dteSessionTime = IIf(IsNull(grstJournSessionsQry!time_connected), NULL_DATE, grstJournSessionsQry!time_connected) + IIf(IsNull(grstJournSessionsQry!time_disconnected), NULL_DATE, grstJournSessionsQry!time_disconnected) + IIf(IsNull(grstJournSessionsQry!time_connected_offpeak), NULL_DATE, grstJournSessionsQry!time_connected_offpeak)
                .TextMatrix(i, 0) = grstJournSessionsQry!date_start
                Select Case grstJournSessionsQry!Type
                Case enuSessionType.Card
                    .TextMatrix(i, 1) = grstJournSessionsQry!card_id
                Case enuSessionType.Client
                    .TextMatrix(i, 1) = IIf(IsNull(grstJournSessionsQry!client_login), "", grstJournSessionsQry!client_login)
                Case enuSessionType.Free
                    .TextMatrix(i, 1) = funLoadResString(10622)
                Case enuSessionType.Selfservice
                    .TextMatrix(i, 1) = funLoadResString(10313)
                Case enuSessionType.SelfServicePrePaid
                    .TextMatrix(i, 1) = funLoadResString(10986)
                End Select
                .TextMatrix(i, 2) = IIf(IsNull(grstJournSessionsQry!workstation_name), "", grstJournSessionsQry!workstation_name)
                .TextMatrix(i, 3) = funFormatTime(dteSessionTime)
                dteTotalSessionTime = dteTotalSessionTime + dteSessionTime
                
                If grstJournSessionsQry!time_connected <> NULL_DATE Then
                    .TextMatrix(i, 4) = funFormatTime(grstJournSessionsQry!time_connected)
                    dteTotalCon = dteTotalCon + grstJournSessionsQry!time_connected
                    lngNumTimeConnected = lngNumTimeConnected + 1
                End If
                If dteSessionTime <> NULL_DATE And grstJournSessionsQry!time_connected <> NULL_DATE Then .TextMatrix(i, 5) = format(grstJournSessionsQry!time_connected / dteSessionTime * 100, "0")
                
                If grstJournSessionsQry!time_disconnected <> NULL_DATE Then
                    .TextMatrix(i, 6) = funFormatTime(grstJournSessionsQry!time_disconnected)
                    dteTotalDiscon = dteTotalDiscon + grstJournSessionsQry!time_disconnected
                    lngNumTimeDisconnected = lngNumTimeDisconnected + 1
                End If
                If dteSessionTime <> NULL_DATE And grstJournSessionsQry!time_disconnected <> NULL_DATE Then .TextMatrix(i, 7) = format(grstJournSessionsQry!time_disconnected / dteSessionTime * 100, "0")
                
                If grstJournSessionsQry!time_connected_offpeak <> NULL_DATE Then
                    .TextMatrix(i, 8) = funFormatTime(grstJournSessionsQry!time_connected_offpeak)
                    dteTotalConOffpeak = dteTotalConOffpeak + grstJournSessionsQry!time_connected_offpeak
                    lngNumTimeConnectedOffPeak = lngNumTimeConnectedOffPeak + 1
                End If
                If dteSessionTime <> NULL_DATE And grstJournSessionsQry!time_connected_offpeak <> NULL_DATE Then .TextMatrix(i, 9) = format(grstJournSessionsQry!time_connected_offpeak / dteSessionTime * 100, "0")
                .TextMatrix(i, 10) = IIf(IsNull(grstJournSessionsQry!login), "", grstJournSessionsQry!login)
                
                Call grstJournSessionsQry.MoveNext
                DoEvents
                If gblnStopStats Then Exit Sub
                frmStatsProgress.pgbStatProgress.value = frmStatsProgress.pgbStatProgress.Max / grstJournSessionsQry.RecordCount * i
                frmStatsProgress.pgbStatType.value = frmStatsProgress.pgbStatType.Max * 0.4 + frmStatsProgress.pgbStatType.Max * 0.2 / grstJournSessionsQry.RecordCount * i
            Next i
            i = grstJournSessionsQry.RecordCount + 1
            .Row = i
            .Col = 0
            .ColSel = .Cols - 1
            .CellBackColor = vbButtonShadow
            
            .TextMatrix(i, 0) = funLoadResString(10387) & " (" & grstJournSessionsQry.RecordCount & funLoadResString(10410) & ")"
            .TextMatrix(i, 3) = funFormatTime(dteTotalSessionTime)
            .TextMatrix(i, 4) = funFormatTime(dteTotalCon)
            If dteTotalSessionTime <> NULL_DATE Then .TextMatrix(i, 5) = format(dteTotalCon / dteTotalSessionTime * 100, "0")
            .TextMatrix(i, 6) = funFormatTime(dteTotalDiscon)
            If dteTotalSessionTime <> NULL_DATE Then .TextMatrix(i, 7) = format(dteTotalDiscon / dteTotalSessionTime * 100, "0")
            .TextMatrix(i, 8) = funFormatTime(dteTotalConOffpeak)
            If dteTotalSessionTime <> NULL_DATE Then .TextMatrix(i, 9) = format(dteTotalConOffpeak / dteTotalSessionTime * 100, "0")
            
            i = i + 1
            .Row = i
            .Col = 0
            .ColSel = .Cols - 1
            .CellBackColor = vbButtonFace
            
            .TextMatrix(i, 0) = funLoadResString(10388)
            .TextMatrix(i, 3) = funFormatTime(dteTotalSessionTime / grstJournSessionsQry.RecordCount)
            If lngNumTimeConnected <> 0 Then .TextMatrix(i, 4) = funFormatTime(dteTotalCon / lngNumTimeConnected)
            If lngNumTimeDisconnected <> 0 Then .TextMatrix(i, 6) = funFormatTime(dteTotalDiscon / lngNumTimeDisconnected)
            If lngNumTimeConnectedOffPeak <> 0 Then .TextMatrix(i, 8) = funFormatTime(dteTotalConOffpeak / lngNumTimeConnectedOffPeak)
                    
        End If
        .TopRow = .Rows - 1
    End With
    Exit Sub
    
subFillSessionListError:
    Call funErrorManager("modStats.subFillSessionListError")
End Sub


'*******************************************************************************
Public Sub subFillURLList()
Dim i As Long, j As Long
On Error GoTo subFillURLListError

    If Not gblnDBOK Then Exit Sub
    With frmMain.fxgSession(4)
        .Rows = 2
        For i = 1 To .Rows - 1
            For j = 0 To .Cols - 1
                .TextMatrix(i, j) = ""
            Next j
        Next i
        If grstUrlsQry.BOF = True And grstUrlsQry.EOF = True Then
            .Rows = 2
        Else
            Call grstUrlsQry.MoveFirst
            .Rows = grstUrlsQry.RecordCount + 1
            .enabled = True
            For i = 1 To grstUrlsQry.RecordCount
                .TextMatrix(i, 0) = grstUrlsQry!date_stamp
                .TextMatrix(i, 1) = IIf(IsNull(grstUrlsQry!workstation_name), "", grstUrlsQry!workstation_name)
                .TextMatrix(i, 2) = IIf(grstUrlsQry!card_id = 0, IIf(IsNull(grstUrlsQry!client_login), "", grstUrlsQry!client_login), grstUrlsQry!card_id)
                .TextMatrix(i, 3) = IIf(IsNull(grstUrlsQry!Title), "", grstUrlsQry!Title)
                .TextMatrix(i, 4) = IIf(IsNull(grstUrlsQry!url), "", grstUrlsQry!url)
                
                Call grstUrlsQry.MoveNext
                DoEvents
                If gblnStopStats Then Exit Sub
                frmStatsProgress.pgbStatProgress.value = frmStatsProgress.pgbStatProgress.Max / grstUrlsQry.RecordCount * i
                frmStatsProgress.pgbStatType.value = frmStatsProgress.pgbStatType.Max * 0.6 + frmStatsProgress.pgbStatType.Max * 0.2 / grstUrlsQry.RecordCount * i
            Next i
        End If
        .ToolTipText = grstUrlsQry.RecordCount & funLoadResString(10303)
        .TopRow = .Rows - 1
    End With
    Exit Sub
    
subFillURLListError:
    Call funErrorManager("modStats.subFillURLListError")
End Sub


'*******************************************************************************
Public Sub subFillSessionTimes()
Dim i As Long, j As Long, k As Long
Dim lngNumDays As Long
Dim intNumIntervals As Integer
Dim lngTotalTimeInt() As Long
Dim lngTotalCount As Long
Dim lngTotalSession As Long
Dim lngNumOpenDays As Long
On Error GoTo subFillSessionTimesError

    If Not gblnDBOK Then Exit Sub
    
    intNumIntervals = funArraySize(guServOpts.uStats.strStatsTimeInt)
    ReDim lngTotalTimeInt(intNumIntervals)
    
    With frmMain.fxgSession(2)
        .Rows = 2
        For i = 1 To .Rows - 1
            For j = 0 To .Cols - 1
                .TextMatrix(i, j) = ""
            Next j
        Next i
        On Error Resume Next
        lngNumDays = UBound(guStatsDay)
        If Err.Number <> 0 Then Exit Sub
        On Error GoTo subFillSessionTimesError
        .Rows = lngNumDays + 3
        For i = 1 To lngNumDays
            .TextMatrix(i, 0) = format(guStatsTimes(i).dteDate, "Short Date")
            
            If guStatsTimes(i).lngSessionNum <> 0 Then .TextMatrix(i, 1) = guStatsTimes(i).lngSessionNum
            lngTotalSession = lngTotalSession + guStatsTimes(i).lngSessionNum
            
            For j = 0 To intNumIntervals
                If guStatsTimes(i).lngTimeInt(j) <> 0 Then .TextMatrix(i, j * 2 + 2) = guStatsTimes(i).lngTimeInt(j)
                lngTotalTimeInt(j) = lngTotalTimeInt(j) + guStatsTimes(i).lngTimeInt(j)
                lngTotalCount = lngTotalCount + guStatsTimes(i).lngTimeInt(j)
                If guStatsTimes(i).lngTimeInt(j) <> 0 And guStatsTimes(i).lngSessionCount <> 0 Then .TextMatrix(i, j * 2 + 3) = format(guStatsTimes(i).lngTimeInt(j) / guStatsTimes(i).lngSessionCount * 100, "0")
            
            Next j
            If guStatsTimes(i).lngSessionNum <> 0 Then lngNumOpenDays = lngNumOpenDays + 1
            DoEvents
            If gblnStopStats Then Exit Sub
            frmStatsProgress.pgbStatProgress.value = frmStatsProgress.pgbStatProgress.Max * 0.8 + frmStatsProgress.pgbStatProgress.Max * 0.2 / lngNumDays * i
            frmStatsProgress.pgbStatType.value = frmStatsProgress.pgbStatType.Max * 0.2 + frmStatsProgress.pgbStatType.Max * 0.2 * 0.8 + frmStatsProgress.pgbStatType.Max * 0.25 * 0.2 / lngNumDays * i
        Next i
        
        i = lngNumDays + 1
        .Row = i
        .Col = 0
        .ColSel = .Cols - 1
        .CellBackColor = vbButtonShadow

        .TextMatrix(i, 0) = funLoadResString(10387) & " (" & lngNumOpenDays & funLoadResString(10389) & ")"
        .TextMatrix(i, 1) = lngTotalSession

        For j = 0 To intNumIntervals
            If lngTotalTimeInt(j) <> 0 Then .TextMatrix(i, j * 2 + 2) = lngTotalTimeInt(j)
            If lngTotalTimeInt(j) <> 0 And lngTotalCount <> 0 Then .TextMatrix(i, j * 2 + 3) = format(lngTotalTimeInt(j) / lngTotalCount * 100, "0")
        Next j

        i = lngNumDays + 2
        .Row = i
        .Col = 0
        .ColSel = .Cols - 1
        .CellBackColor = vbButtonFace

        .TextMatrix(i, 0) = funLoadResString(10388)
        If lngNumOpenDays <> 0 Then
            .TextMatrix(i, 1) = Int(lngTotalSession / lngNumOpenDays)
            For j = 0 To intNumIntervals
                If lngTotalTimeInt(j) <> 0 Then .TextMatrix(i, j * 2 + 2) = Round(lngTotalTimeInt(j) / lngNumOpenDays)
            Next j
        End If
        .TopRow = .Rows - 1
        
    End With
    Exit Sub
    
subFillSessionTimesError:
    Call funErrorManager("modStats.subFillSessionTimesError")
End Sub


'*******************************************************************************
Public Sub subFillSessionLines()
Dim i As Long, j As Long, k As Long
Dim blnToggleBack As Boolean
Dim intNumHours As Integer
Dim dteStartMidnight As Date
Dim dteLastStartMidnight As Date
Dim dteStartTime As Date
Dim intStartIndex As Integer
Dim dteStopTime As Date
Dim dteCurrentTime As Date
Dim dteNextCurrentTime As Date
Dim blnInSession As Boolean
On Error GoTo subFillSessionLinesError

    If Not gblnDBOK Then Exit Sub
    intNumHours = guServOpts.uStats.intStatsGraphStopTime - guServOpts.uStats.intStatsGraphStartTime
    With frmMain.fxgSession(5)
        .Visible = False
        .Rows = 2
        For i = 1 To .Rows - 1
            .TextMatrix(i, 0) = ""
            For j = 1 To .Cols - 1
                .TextMatrix(i, j) = ""
                .Col = 1
                .Row = i
                .ColSel = .Cols - 1
                .CellBackColor = vbWindowBackground
            Next j
        Next i
        If grstJournSessionsQry.BOF = True And grstJournSessionsQry.EOF = True Then
'            .Rows = 100
        Else
            Call grstJournSessionsQry.MoveFirst
            .Rows = grstJournSessionsQry.RecordCount + 1
            .enabled = True
            For i = 1 To grstJournSessionsQry.RecordCount
                If Not IsNull(grstJournSessionsQry!date_start) Then
                    .Row = i
                    .RowHeight(i) = 50
                    dteStartTime = grstJournSessionsQry!date_start
                    dteLastStartMidnight = dteStartMidnight
                    dteStartMidnight = CDate(DateValue(dteStartTime) & " " & format(guServOpts.uStats.intStatsGraphStartTime, "00") & ":00:00")
                    If dteLastStartMidnight <> dteStartMidnight Then blnToggleBack = Not blnToggleBack
                    .Col = 0
                    .ColSel = .Cols - 1
                    .CellBackColor = IIf(blnToggleBack, vbWindowBackground, vbButtonFace)
                    dteStopTime = IIf(IsNull(grstJournSessionsQry!date_end), IIf(IsNull(grstJournSessionsQry!date_start), NULL_DATE, grstJournSessionsQry!date_start) + IIf(IsNull(grstJournSessionsQry!time_connected), NULL_DATE, grstJournSessionsQry!time_connected) + IIf(IsNull(grstJournSessionsQry!time_disconnected), NULL_DATE, grstJournSessionsQry!time_disconnected) + IIf(IsNull(grstJournSessionsQry!time_connected_offpeak), NULL_DATE, grstJournSessionsQry!time_connected_offpeak), grstJournSessionsQry!date_end)
                    .TextMatrix(i, 0) = format(dteStartTime, "Short Date")
                    For j = 0 To intNumHours - 1
                        k = j * 12 + j + 1
                        .Col = k
                        .ColSel = .Col
                        .CellBackColor = &HC0C0C0
                    Next j
                    .Col = .Cols - 1
                    .ColSel = .Col
                    .CellBackColor = &HC0C0C0
                    blnInSession = False
                    For j = 0 To (intNumHours) * 12
                        k = j + Int(j / 12) + 1 + 1
                        If k < .Cols Then
                            dteCurrentTime = DateAdd("n", j * 5, dteStartMidnight)
                            dteNextCurrentTime = DateAdd("n", (j + 1) * 5, dteStartMidnight)
                            If dteStartTime <= dteNextCurrentTime And Not blnInSession Then
                                intStartIndex = k
                                blnInSession = True
                            End If
                            If blnInSession And dteStopTime <= dteCurrentTime Then
                                .Col = intStartIndex
                                .ColSel = k
                                .CellBackColor = vbButtonShadow
                                Exit For
                            End If
                        End If
                    Next j
                Else
                    .Row = i
                    For j = 0 To intNumHours - 1
                        k = j * 12 + j + 1
                        .Col = k
                        .ColSel = .Col
                        .CellBackColor = &HC0C0C0
                    Next j
                End If
                Call grstJournSessionsQry.MoveNext
                DoEvents
                If gblnStopStats Then
                    .Visible = True
                    Exit Sub
                End If
                frmStatsProgress.pgbStatProgress.value = frmStatsProgress.pgbStatProgress.Max / grstJournSessionsQry.RecordCount * i
                frmStatsProgress.pgbStatType.value = frmStatsProgress.pgbStatType.Max * 0.8 + frmStatsProgress.pgbStatType.Max * 0.2 / grstJournSessionsQry.RecordCount * i
            Next i
        End If
        .Visible = True
    End With
    Exit Sub
    
subFillSessionLinesError:
    Call funErrorManager("modStats.subFillSessionLinesError")
End Sub


'*******************************************************************************
Public Sub subFillSessionStats()
Dim i As Long, j As Long, k As Long
Dim lngNumDays As Long
Dim lngNumOpenDays As Long
Dim strTmp As String
Dim dteTotalOpen As Date
Dim dteTotalSession As Date
Dim dteTotalInetCnx As Date
Dim dteTotalWs() As Date
Dim lngTotalSession As Long
On Error GoTo subFillSessionStatsError

    If Not gblnDBOK Then Exit Sub
    ReDim dteTotalWs(guServOpts.uWorkstations.lngNumWorkstations)
    
    With frmMain.fxgSession(1)
        .Rows = 2
        For i = 1 To .Rows - 1
            For j = 0 To .Cols - 1
                .TextMatrix(i, j) = ""
            Next j
        Next i
        On Error Resume Next
        lngNumDays = UBound(guStatsDay)
        If Err.Number <> 0 Then Exit Sub
        On Error GoTo subFillSessionStatsError
        .Rows = lngNumDays + 3
        For i = 1 To lngNumDays
            .TextMatrix(i, 0) = format(guStatsDay(i).dteDate, "Short Date")
            
            If guStatsDay(i).uTimeOpen.dteTotal <> NULL_DATE Then .TextMatrix(i, 1) = funFormatTime(guStatsDay(i).uTimeOpen.dteTotal)
            dteTotalOpen = dteTotalOpen + guStatsDay(i).uTimeOpen.dteTotal
            
            If guStatsDay(i).lngSessionNum <> 0 Then .TextMatrix(i, 2) = guStatsDay(i).lngSessionNum
            lngTotalSession = lngTotalSession + guStatsDay(i).lngSessionNum
            
            If guStatsDay(i).dteSessionTotal <> NULL_DATE Then .TextMatrix(i, 3) = funFormatTime(guStatsDay(i).dteSessionTotal)
            dteTotalSession = dteTotalSession + guStatsDay(i).dteSessionTotal
            
            If guStatsDay(i).uTimeInetCnx.dteTotal <> NULL_DATE Then .TextMatrix(i, 4) = funFormatTime(guStatsDay(i).uTimeInetCnx.dteTotal)
            If guStatsDay(i).uTimeInetCnx.dteTotal <> NULL_DATE And guStatsDay(i).uTimeOpen.dteTotal <> NULL_DATE Then .TextMatrix(i, 5) = format(guStatsDay(i).uTimeInetCnx.dteTotal / guStatsDay(i).uTimeOpen.dteTotal * 100, "0")
            dteTotalInetCnx = dteTotalInetCnx + guStatsDay(i).uTimeInetCnx.dteTotal
            
            If guStatsDay(i).uTimeInetCnx.dteTotal <> NULL_DATE And guStatsDay(i).dteSessionTotal <> NULL_DATE Then .TextMatrix(i, 6) = format(Round(Abs(guStatsDay(i).dteSessionTotal / guStatsDay(i).uTimeInetCnx.dteTotal), 2), "0.00")
            
            For j = 6 To .Cols - 2 Step 2
                k = j / 2 - 3
                If guStatsDay(i).uTimeWs(k).dteTotal <> NULL_DATE Then .TextMatrix(i, j + 1) = funFormatTime(guStatsDay(i).uTimeWs(k).dteTotal)
                If guStatsDay(i).dteTimeWsTotal <> NULL_DATE And guStatsDay(i).uTimeWs(k).dteTotal <> NULL_DATE Then .TextMatrix(i, j + 2) = format(guStatsDay(i).uTimeWs(k).dteTotal / guStatsDay(i).dteTimeWsTotal * 100, "0")
                dteTotalWs(k) = dteTotalWs(k) + guStatsDay(i).uTimeWs(k).dteTotal
            Next j
            If guStatsDay(i).uTimeOpen.dteTotal <> NULL_DATE Then lngNumOpenDays = lngNumOpenDays + 1
            DoEvents
            If gblnStopStats Then Exit Sub
            frmStatsProgress.pgbStatProgress.value = frmStatsProgress.pgbStatProgress.Max * 0.8 + frmStatsProgress.pgbStatProgress.Max * 0.2 / lngNumDays * i
            frmStatsProgress.pgbStatType.value = frmStatsProgress.pgbStatType.Max * 0.2 * 0.8 + frmStatsProgress.pgbStatType.Max * 0.2 * 0.2 / lngNumDays * i
        Next i
        
        i = lngNumDays + 1
        .Row = i
        .Col = 0
        .ColSel = .Cols - 1
        .CellBackColor = vbButtonShadow
        
        .TextMatrix(i, 0) = funLoadResString(10387) & " (" & lngNumOpenDays & funLoadResString(10389) & ")"
        .TextMatrix(i, 1) = funFormatTime(dteTotalOpen)
        .TextMatrix(i, 2) = lngTotalSession
        .TextMatrix(i, 3) = funFormatTime(dteTotalSession)
        .TextMatrix(i, 4) = funFormatTime(dteTotalInetCnx)
        If dteTotalOpen <> NULL_DATE Then .TextMatrix(i, 5) = format(dteTotalInetCnx / dteTotalOpen * 100, "0")
        For j = 6 To .Cols - 2 Step 2
            k = j / 2 - 3
            .TextMatrix(i, j + 1) = funFormatTime(dteTotalWs(k))
            If dteTotalOpen <> NULL_DATE Then .TextMatrix(i, j + 2) = format(dteTotalWs(k) / dteTotalOpen * 100, "0")
        Next j
        
        i = lngNumDays + 2
        .Row = i
        .Col = 0
        .ColSel = .Cols - 1
        .CellBackColor = vbButtonFace
        
        .TextMatrix(i, 0) = funLoadResString(10388)
        If lngNumOpenDays <> 0 Then
            .TextMatrix(i, 1) = funFormatTime(dteTotalOpen / lngNumOpenDays)
            .TextMatrix(i, 2) = Int(lngTotalSession / lngNumOpenDays)
            .TextMatrix(i, 3) = funFormatTime(dteTotalSession / lngNumOpenDays)
            .TextMatrix(i, 4) = funFormatTime(dteTotalInetCnx / lngNumOpenDays)
            If dteTotalInetCnx <> NULL_DATE Then .TextMatrix(i, 6) = format(Round(Abs(dteTotalSession / dteTotalInetCnx), 2), "0.00")
            For j = 6 To .Cols - 2 Step 2
                k = j / 2 - 3
                .TextMatrix(i, j + 1) = funFormatTime(dteTotalWs(k) / lngNumOpenDays)
            Next j
        End If
        .TopRow = .Rows - 1
        
    End With
    Exit Sub
    
subFillSessionStatsError:
    Call funErrorManager("modStats.subFillSessionStatsError")
End Sub


'*******************************************************************************
Public Sub subStatsPrint(fxgSource As MSHFlexGrid, strTitle As String)
Dim uPrinter As Printer
Dim fxgDest As MSHFlexGrid
Dim strDateTime As String
Dim lngNumRowsRemaining As Long
Dim lngNumRowsToPrint As Long
Dim intNumVisibleRows As Integer
Dim lngNumPages As Long
Dim i As Long, j As Long, k As Long
On Error GoTo subStatsPrintError

    If guServOpts.uConfig.strDefaultPrinterName = "" Then Exit Sub
    
    fxgSource.MousePointer = vbHourglass
    Screen.MousePointer = vbHourglass
    ' Set up printer
    Printer.Orientation = vbPRORLandscape
'    Printer.PaperSize = vbPRPSA4
    Printer.PrintQuality = vbPRPQMedium
    
    ' Set up printing form (size & layout)
    Call Load(frmStatsPrint)
    With frmStatsPrint
        .width = IIf(Printer.PaperSize = vbPRPSLetter, PAPER_LONG_SIDE_LR, PAPER_LONG_SIDE_A4) * TWIPS_PER_CM
        .height = IIf(Printer.PaperSize = vbPRPSLetter, PAPER_SHORT_SIDE_LR, PAPER_SHORT_SIDE_A4) * TWIPS_PER_CM
        .imgLogo.top = PRINT_MARGIN
        .fxgContent.top = PRINT_MARGIN + 1
        .fxgContent.left = PRINT_MARGIN
        .fxgContent.width = .ScaleWidth - 2 * PRINT_MARGIN
        .fxgContent.height = .ScaleHeight - .fxgContent.top - PRINT_MARGIN
        .imgLogo.left = PRINT_MARGIN
        .lblDtePage.top = PRINT_MARGIN
        .lblDtePage.left = PRINT_MARGIN
        .lblDtePage.width = .fxgContent.width
        .lblTitle.top = PRINT_MARGIN
        .lblTitle.left = PRINT_MARGIN + .imgLogo.width
        .lblTitle.width = .fxgContent.width - .imgLogo.width
        .fxgContent.Rows = NUM_ROWS_STATS_PRINT
        For i = 0 To .fxgContent.Rows - 1
            If Not .fxgContent.RowIsVisible(i) Then Exit For
        Next i
        intNumVisibleRows = i - 1
    End With
    
    ' Define print job variables and fill out constant data
    strDateTime = format(Date, "General Date") & " " & format(Time, "General Date")
    Set fxgDest = frmStatsPrint.fxgContent
    fxgDest.Cols = fxgSource.Cols
    For i = 0 To fxgSource.Cols - 1
        fxgDest.ColWidth(i) = fxgSource.ColWidth(i)
        fxgDest.TextMatrix(0, i) = fxgSource.TextMatrix(0, i)
        fxgDest.ColAlignment(i) = fxgSource.ColAlignment(i)
    Next i
    frmStatsPrint.lblTitle.Caption = " " & funLoadResString(10005) & " - " & strTitle
    lngNumPages = -Int(-(fxgSource.Rows - 1) / (intNumVisibleRows - 1))
    lngNumRowsRemaining = fxgSource.Rows
    
    If lngNumPages > PRINT_MAX_PAGES_BEFORE_WARNING Then
        If funMsg(AskCanPrint, lngNumPages) = vbNo Then
            Call Unload(frmStatsPrint)
            Screen.MousePointer = vbNormal
            fxgSource.MousePointer = vbNormal
            Exit Sub
        End If
    End If
    
    ' Print all pages that make up the print job
    For k = 1 To lngNumPages
        frmStatsPrint.lblDtePage.Caption = funLoadResString(10964, strDateTime, CStr(k), CStr(lngNumPages)) & " "
        
        ' Define data grid size
        If lngNumRowsRemaining > intNumVisibleRows Then
            fxgDest.Rows = intNumVisibleRows
            lngNumRowsToPrint = intNumVisibleRows - 1
        Else
            fxgDest.Rows = lngNumRowsRemaining
            lngNumRowsToPrint = lngNumRowsRemaining - 1
        End If
        
        ' Fill out the grid with data
        For j = 1 To lngNumRowsToPrint
            For i = 0 To fxgSource.Cols - 1
                DoEvents
                fxgDest.TextMatrix(j, i) = fxgSource.TextMatrix(fxgSource.Rows - lngNumRowsRemaining + j, i)
            Next i
        Next j
                    
        Call frmStatsPrint.PrintForm
        
        lngNumRowsRemaining = lngNumRowsRemaining - intNumVisibleRows - 1
        For j = 1 To fxgDest.Rows - 1
            For i = 0 To fxgDest.Cols - 1
                DoEvents
                fxgDest.TextMatrix(j, i) = ""
            Next i
        Next j
    Next k
        
    Call Unload(frmStatsPrint)
    Screen.MousePointer = vbNormal
    fxgSource.MousePointer = vbNormal
    Exit Sub
    
subStatsPrintError:
    Screen.MousePointer = vbNormal
    fxgSource.MousePointer = vbNormal
    Call funErrorManager("modStats.subStatsPrintError")
End Sub
