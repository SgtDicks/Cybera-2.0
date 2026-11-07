Attribute VB_Name = "modMessageManager"
' _____________________________________________________________________________
'|________________________________ MESSAGE GENERATION _________________________|
'|  Name:          modMessageManager.bas                                       |
'|  Type:          VB6 Module                                                  |
'|  Project:       Cybera Workstation (http://cybera.sf.net)                   |
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
    BadPasswd = 2
    BadNewPasswd = 3
    SaveSettings = 4
    BadCardNum = 5
    NoServerCnx = 6
    Logoff = 7
    CloseSession = 8
    NoElementSelected = 9
    FoundKilledApp = 10
    ChangeToConnected = 11
    ChangeToDisconnected = 12
    NoServerResponse = 13
    CardEmpty = 14
    CardDeactivated = 15
    CardInUse = 16
    ConnectionForced = 17
    WksnMsgCard = 18
    WksnMsgSelfService = 19
    AppAlreadyRunning = 20
    PauseSession = 21
    AskPauseSession = 22
    AskUnPauseSession = 23
    UntrappedError = 24
    OSNotNT = 25
    NetworkConfigError = 26
    WksnDisabled = 27
    FileNotFound = 28
    CancelWizard = 29
    InvalidWksnName = 30
    ShutdownWksnTypeChanged = 31
End Enum


'*******************************************************************************
Public Function funMsg(ByVal intMsgType As enuMsg, Optional strText As Variant = "", Optional strTitle As Variant) As Integer
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
    
    Case enuMsg.BadCardNum
        lngButtons = vbOKOnly + vbExclamation + vbMsgBoxSetForeground + vbSystemModal
        strPrompt = funLoadResString(10104)
    
    Case enuMsg.NoServerCnx
        lngButtons = vbOKOnly + vbCritical + vbMsgBoxSetForeground + vbSystemModal
        strPrompt = funLoadResString(10178) & vbNewLine & funLoadResString(10179)
    
    Case enuMsg.Logoff
        lngButtons = vbYesNo + vbDefaultButton2 + vbExclamation + vbMsgBoxSetForeground + vbSystemModal
        If strText = "" Then
            strPrompt = funLoadResString(10180)
        Else
            strPrompt = strText & funLoadResString(10181)
        End If
        strPrompt = funLoadResString(10182) & strPrompt & vbNewLine & vbNewLine & funLoadResString(10183)
    
    Case enuMsg.CloseSession
        lngButtons = vbYesNo + vbDefaultButton2 + vbQuestion + vbMsgBoxSetForeground + vbSystemModal
        strPrompt = funLoadResString(10184)
    
    Case enuMsg.NoElementSelected
        lngButtons = vbOKOnly + vbExclamation
        strPrompt = funLoadResString(10185)
    
    Case enuMsg.FoundKilledApp
        lngButtons = vbOKOnly + vbExclamation
        strPrompt = funLoadResString(10000) & funLoadResString(10186) & vbNewLine & funLoadResString(10187)
    
    Case enuMsg.ChangeToDisconnected
        lngButtons = vbYesNo + vbDefaultButton2 + vbQuestion + vbMsgBoxSetForeground + vbSystemModal
        strPrompt = funLoadResString(10188) & vbNewLine & vbNewLine & funLoadResString(10189)
    
    Case enuMsg.ChangeToConnected
        lngButtons = vbYesNo + vbDefaultButton2 + vbQuestion + vbMsgBoxSetForeground + vbSystemModal
        strPrompt = funLoadResString(10190)
    
    Case enuMsg.NoServerResponse
        lngButtons = vbOKOnly + vbCritical + vbMsgBoxSetForeground + vbSystemModal
        strPrompt = funLoadResString(10191) & vbNewLine & funLoadResString(10179)
    
    Case enuMsg.CardEmpty
        lngButtons = vbOKOnly + vbExclamation + vbMsgBoxSetForeground + vbSystemModal
        strPrompt = funLoadResString(10192) & vbNewLine & funLoadResString(10193)
    
    Case enuMsg.CardDeactivated
        lngButtons = vbOKOnly + vbExclamation + vbMsgBoxSetForeground + vbSystemModal
        strPrompt = funLoadResString(10194) & vbNewLine & funLoadResString(10179)
    
    Case enuMsg.CardInUse
        lngButtons = vbOKOnly + vbExclamation + vbMsgBoxSetForeground + vbSystemModal
        strPrompt = funLoadResString(10195) & vbNewLine & funLoadResString(10179)
    
    Case enuMsg.ConnectionForced
        lngButtons = vbOKOnly + vbExclamation + vbMsgBoxSetForeground + vbSystemModal
        strPrompt = funLoadResString(10196) & vbNewLine & funLoadResString(10197)
    
    Case enuMsg.WksnMsgCard
        lngButtons = vbOKOnly + vbInformation + vbMsgBoxSetForeground + vbSystemModal
        strPrompt = funLoadResString(10198) & strText & funLoadResString(10199)
    
    Case enuMsg.WksnMsgSelfService
        lngButtons = vbOKOnly + vbInformation + vbMsgBoxSetForeground + vbSystemModal
        strPrompt = funLoadResString(10200) & strText & funLoadResString(10201)
    
    Case enuMsg.AppAlreadyRunning
        lngButtons = vbOKOnly + vbCritical + vbMsgBoxSetForeground + vbSystemModal
        strPrompt = funLoadResString(10000) & funLoadResString(10202) & vbNewLine & funLoadResString(10203)
    
    Case enuMsg.PauseSession
        lngButtons = vbYesNo + vbDefaultButton2 + vbQuestion + vbMsgBoxSetForeground + vbSystemModal
        strPrompt = funLoadResString(10204)
    
    Case enuMsg.AskPauseSession
        lngButtons = vbYesNo + vbDefaultButton2 + vbQuestion + vbMsgBoxSetForeground + vbSystemModal
        strPrompt = funLoadResString(10205) & vbNewLine & funLoadResString(10206)
    
    Case enuMsg.AskUnPauseSession
        lngButtons = vbYesNo + vbQuestion + vbMsgBoxSetForeground + vbSystemModal
        strPrompt = funLoadResString(10207) & vbNewLine & funLoadResString(10208)
    
    Case enuMsg.UntrappedError
        lngButtons = vbOKOnly + vbCritical + vbMsgBoxSetForeground + vbSystemModal
        strPrompt = funLoadResString(10209) & strText & vbNewLine & funLoadResString(10210) & MY_EMAIL & funLoadResString(10211) & vbNewLine & gstrLogfilePath
    
    Case enuMsg.OSNotNT
        lngButtons = vbOKOnly + vbCritical
        strPrompt = funLoadResString(10214) & vbNewLine & funLoadResString(10215) & funLoadResString(10000)
    
    Case enuMsg.NetworkConfigError
        lngButtons = vbOKOnly + vbCritical
        strPrompt = funLoadResString(10216) & vbNewLine & funLoadResString(10217) & funLoadResString(10000)
    
    Case enuMsg.WksnDisabled
        lngButtons = vbOKOnly + vbCritical
        strPrompt = funLoadResString(10000) & funLoadResString(10261) & vbNewLine & strText & vbNewLine & funLoadResString(10262) & funLoadResString(10000)
    
    Case enuMsg.CancelWizard
        lngButtons = vbYesNo + vbQuestion + vbMsgBoxSetForeground + vbDefaultButton2
        strPrompt = funLoadResString(10022)
    
    Case enuMsg.InvalidWksnName
        lngButtons = vbOKOnly + vbExclamation + vbMsgBoxSetForeground + vbSystemModal
        strPrompt = funLoadResString(10317, TIME_TOKEN)
    
    Case enuMsg.ShutdownWksnTypeChanged
        lngButtons = vbOKOnly + vbExclamation + vbMsgBoxSetForeground + vbSystemModal
        strPrompt = funLoadResString(10345)
    

    End Select
    
    If strMsgTitle = "" Then
        strMsgTitle = funLoadResString(10000)
    Else
        strMsgTitle = funLoadResString(10000) & " - " & strMsgTitle
    End If
    If BETA_VERSION Then strMsgTitle = strMsgTitle & " - Msg# " & intMsgType
    
    funMsg = MsgBox(strPrompt, lngButtons, strMsgTitle)
    Exit Function

funMsgError:
    Call funErrorManager("modMessageManager.funMsgError")
End Function

