Attribute VB_Name = "modGlobalFunctions"
' _____________________________________________________________________________
'|________________________________ GLOBAL FUNCTION MODULE _____________________|
'|  Name:          modGlobalFunctions.bas                                      |
'|  Type:          VB6 Module                                                  |
'|  Project:       Cybera Server & Workstation (http://cybera.sf.net)          |
'|  Description:   Contains global functions ans subs                          |
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

' $Id: modGlobalFunctions.bas 19 2006-01-08 21:00:47Z sam $

Option Explicit
      
'____________________________________________________________________________________________________
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Const SW_SHOWNORMAL = 1
Const SE_ERR_FNF = 2&
Const SE_ERR_PNF = 3&
Const SE_ERR_ACCESSDENIED = 5&
Const SE_ERR_OOM = 8&
Const SE_ERR_DLLNOTFOUND = 32&
Const SE_ERR_SHARE = 26&
Const SE_ERR_ASSOCINCOMPLETE = 27&
Const SE_ERR_DDETIMEOUT = 28&
Const SE_ERR_DDEFAIL = 29&
Const SE_ERR_DDEBUSY = 30&
Const SE_ERR_NOASSOC = 31&
Const ERROR_BAD_FORMAT = 11&

'____________________________________________________________________________________________________
Private Declare Function UuidCreateSequential Lib "rpcrt4.dll" (firstbyte As Byte) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (firstbyte As Byte, ByVal lpsz As Long, ByVal cbMax As Long) As Long

'____________________________________________________________________________________________________
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


'*******************************************************************************
Public Function funFormatTime(dteInput As Date, Optional varFormat As Variant) As String
Dim intFormat As Integer
Dim strTmp As String
On Error GoTo funFormatTimeError

    intFormat = enuTimeFormat.HH_NN_SS
    If Not IsMissing(varFormat) Then intFormat = varFormat
    Select Case intFormat
    Case enuTimeFormat.Hh_Nmn
        funFormatTime = Int(CSng(dteInput * 24)) & funLoadResString(10001) & " " & Format(dteInput, "n") & funLoadResString(10002)
    
    Case enuTimeFormat.Hh_Nmn_Ssec
        funFormatTime = Int(CSng(dteInput * 24)) & funLoadResString(10001) & " " & Format(dteInput, "n") & funLoadResString(10002) & " " & Format(dteInput, "s") & funLoadResString(10003)
    
    Case enuTimeFormat.Hh_Nmn_Ssec_dyn
        strTmp = Int(CSng(dteInput * 24)) & funLoadResString(10001)
        If Format(dteInput, "n") <> "0" Or Format(dteInput, "s") <> "0" Then strTmp = strTmp & " " & Format(dteInput, "n") & funLoadResString(10002)
        If Format(dteInput, "s") <> "0" Then strTmp = strTmp & " " & Format(dteInput, "s") & funLoadResString(10003)
        funFormatTime = strTmp
    Case Else
        funFormatTime = Format(Int(Abs(CSng(dteInput * 24))), "00") & ":" & Format(dteInput, "Nn:Ss")
    
    End Select
    Exit Function

funFormatTimeError:
    Call funErrorManager("modGlobalFunctions.funFormatTimeError")
End Function


'*******************************************************************************
Public Sub subDebugPrint(strBuffer As String)
Dim strOutput As String
Dim uSystTime As SYSTEMTIME
On Error Resume Next

    Call GetSystemTime(uSystTime)
    strOutput = App.EXEName & " " & funSystemTime
    strOutput = strOutput & " - " & strBuffer
    Debug.Print (strOutput)
End Sub


'*******************************************************************************
Public Function funSystemTime() As String
Dim uSystTime As SYSTEMTIME
On Error Resume Next

    Call GetSystemTime(uSystTime)
    funSystemTime = Right(uSystTime.wYear, 2) & Format(uSystTime.wMonth, "0#") & Format(uSystTime.wDay, "0#") & "." & Format(uSystTime.wHour, "0#") & Format(uSystTime.wMinute, "0#") & Format(uSystTime.wSecond, "0#") & "." & Format(uSystTime.wMilliseconds, "00#")
End Function


'*******************************************************************************
Public Function funTokenize(strData As String, strToken As String) As String()
Dim strElements() As String
Dim strDataBuffer As String
Dim lngEndIndex As Long
Dim strDatagram As String
On Error Resume Next

    
    If strToken = "" Or strData = "" Then Exit Function
    strDataBuffer = strData
    Do Until strDataBuffer = ""
        lngEndIndex = InStr(strDataBuffer, strToken)
        If lngEndIndex = 0 Then Exit Do
        strDatagram = Left(strDataBuffer, lngEndIndex - 1)
        strDataBuffer = Mid(strDataBuffer, lngEndIndex + Len(strToken))
        ReDim Preserve strElements(funArraySize(strElements)) ' + 1)
'        If Err.Number = 9 Then
'            ReDim strElements(0)
'            Call Err.Clear
'        End If
        strElements(funArraySize(strElements) - 1) = strDatagram
    Loop
    If Len(strDataBuffer) > 0 Then
        ReDim Preserve strElements(funArraySize(strElements)) ' + 1)
        strElements(funArraySize(strElements) - 1) = strDataBuffer
    End If
    funTokenize = strElements
End Function


'*******************************************************************************
Public Function funConcatToken(strData() As String, strToken As String) As String
On Error GoTo funConcatTokenError
Dim strDataBuffer As String
Dim i As Long

    If strToken = "" Then Exit Function
    For i = LBound(strData) To funArraySize(strData) - 1
        strDataBuffer = strDataBuffer & strData(i) & strToken
    Next i
    strDataBuffer = Left(strDataBuffer, Len(strDataBuffer) - Len(strToken))
    funConcatToken = strDataBuffer

funConcatTokenError:
    
End Function


'*******************************************************************************
Public Function funSendData(wsCnx As Winsock, strData As String, Optional blnSendInProgress As Boolean = False, Optional blnTestSendInProgress As Boolean = False) As Boolean
On Error GoTo funSendDataError

    If blnTestSendInProgress Then
'        Do While blnSendInProgress
'            DoEvents
'        Loop
    End If
    If wsCnx.State = sckConnected Then
        Call wsCnx.SendData(CStr(strData))
        funSendData = True
    Else
        funSendData = False
    End If
    Exit Function
    
funSendDataError:
    funSendData = False
    Call funErrorManager("modGlobalFunctions.funSendDataError - Data: " & Left(strData, 20))
End Function


'*******************************************************************************
Public Sub subWriteToLogFile(strOperator As String, Optional strLogfilePath As String = "")
'Append Argument to Backuper.log file.
Dim intLogFileNumber As Integer             '-Contains Backuper.log file number.
Dim strFilePath As String                   '-Contains full path of Backuper.log.
Dim strLogLine As String                    '-Compiled text line to be appended to log file.
On Error GoTo subWriteToLogFileError

     '*Init Vars.
    strFilePath = strLogfilePath
    If strFilePath = "" Then strFilePath = gstrLogfilePath
    If strFilePath = "" Then strFilePath = App.Path & "\" & App.EXEName & ".log"
    
     '*If file does not exist then add header to strLogLine.
    If Dir(strFilePath) = "" Then strLogLine = App.EXEName & " Log file." & vbCrLf & "Creation date: " & Now & vbCrLf & vbCrLf
    
     '*Compile Log lines.
    If strOperator <> "" Then
        strLogLine = strLogLine & App.EXEName & " " & funSystemTime & " - " & strOperator
    End If
    
     '*Get free file number, append log lines to file and close file.
    intLogFileNumber = FreeFile
    Open strFilePath For Append As intLogFileNumber
    Print #intLogFileNumber, strLogLine
    Close intLogFileNumber
    Exit Sub

subWriteToLogFileError:
    Call MsgBox(funLoadResString(10004) & vbNewLine & strFilePath, vbCritical + vbOKOnly, funLoadResString(10000))
End Sub


'*******************************************************************************
Public Sub subAdjustToken()
On Error GoTo subAdjustTokenError
    
    Const TOKEN_ADJUST_PRIVILEGES = &H20
    Const TOKEN_QUERY = &H8
    Const SE_PRIVILEGE_ENABLED = &H2
    Dim hdlProcessHandle As Long
    Dim hdlTokenHandle As Long
    Dim tmpLuid As LUID
    Dim tkp As TOKEN_PRIVILEGES
    Dim tkpNewButIgnored As TOKEN_PRIVILEGES
    Dim lBufferNeeded As Long
    
    hdlProcessHandle = GetCurrentProcess()
    OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
       TOKEN_QUERY), hdlTokenHandle
    
    ' Get the LUID for shutdown privilege.
    LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
    
    tkp.PrivilegeCount = 1    ' One privilege to set
    tkp.TheLuid = tmpLuid
    tkp.Attributes = SE_PRIVILEGE_ENABLED
    
    ' Enable the shutdown privilege in the access token of this process.
    AdjustTokenPrivileges hdlTokenHandle, False, _
    tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
    Exit Sub
    
subAdjustTokenError:
    Call funErrorManager("modGlobalFunctions.subAdjustTokenError")
End Sub


'*******************************************************************************
Public Function funDateTimeToSend(dteDate As Date) As String
On Error GoTo funDateTimeToSendError

    funDateTimeToSend = Replace(CStr(CDbl(dteDate)), funGetDecimalSeparator, ".")
    Exit Function
    
funDateTimeToSendError:
    funDateTimeToSend = ""
End Function


'*******************************************************************************
Public Function funDateTimefromNet(strDate As String) As Date
On Error GoTo funDateTimefromNetError

    funDateTimefromNet = CDate(CDbl(Replace(strDate, ".", funGetDecimalSeparator)))
    Exit Function
    
funDateTimefromNetError:
    funDateTimefromNet = NULL_DATE
End Function


'*******************************************************************************
Public Function funSingleToSend(sngNumber As Single) As String
On Error GoTo ffunSingleToSendError

    funSingleToSend = Replace(CStr(sngNumber), funGetDecimalSeparator, ".")
    Exit Function
    
ffunSingleToSendError:
    funSingleToSend = ""
End Function


'*******************************************************************************
Public Function funSinglefromNet(strNumber As String) As Single
On Error GoTo funSinglefromNetError

    funSinglefromNet = CSng(Replace(strNumber, ".", funGetDecimalSeparator))
    Exit Function
    
funSinglefromNetError:
    funSinglefromNet = 0
End Function


'*******************************************************************************
Public Function funNumHours(dteInput As Date) As Single
On Error GoTo funNumHoursError


    funNumHours = Int(Abs(CSng(dteInput * 24))) + DatePart("n", dteInput) / 60 + DatePart("s", dteInput) / 3600
    Exit Function
    
funNumHoursError:
    funNumHours = 0
End Function


'*******************************************************************************
Public Function funGetDecimalSeparator() As String
Dim strDecSep As String
Dim lngRet1 As Long
Dim lngRet2 As Long
Dim strlpLCDataVar As String
Dim intPos As Integer
Dim lngLocale As Long
On Error GoTo funGetDecimalSeparatorError

    lngLocale = GetUserDefaultLCID()
    lngRet1 = GetLocaleInfo(lngLocale, LOCALE_SDECIMAL, strlpLCDataVar, 0)
    
    strDecSep = String$(lngRet1, 0)
    lngRet2 = GetLocaleInfo(lngLocale, LOCALE_SDECIMAL, strDecSep, lngRet1)
    intPos = InStr(strDecSep, Chr$(0))
    If intPos > 0 Then strDecSep = Left$(strDecSep, intPos - 1)
    funGetDecimalSeparator = strDecSep
    Exit Function

funGetDecimalSeparatorError:
    Call funErrorManager("modGlobalFunctions.funGetDecimalSeparatorError")
End Function


'*******************************************************************************
Public Function funGetCurrencySymbol() As String
Dim strDecSep As String
Dim lngRet1 As Long
Dim lngRet2 As Long
Dim strlpLCDataVar As String
Dim intPos As Integer
Dim lngLocale As Long
On Error GoTo funGetCurrencySymbolError

    lngLocale = GetUserDefaultLCID()
    lngRet1 = GetLocaleInfo(lngLocale, LOCALE_SCURRENCY, strlpLCDataVar, 0)
    
    strDecSep = String$(lngRet1, 0)
    lngRet2 = GetLocaleInfo(lngLocale, LOCALE_SCURRENCY, strDecSep, lngRet1)
    intPos = InStr(strDecSep, Chr$(0))
    If intPos > 0 Then strDecSep = Left$(strDecSep, intPos - 1)
    funGetCurrencySymbol = strDecSep
    Exit Function

funGetCurrencySymbolError:
    Call funErrorManager("modGlobalFunctions.funGetCurrencySymbolError")
End Function


'*******************************************************************************
Public Function funDoXOR(strInputData As String, strKey As String) As String
On Error GoTo funDoXORError

  Dim lngCtr As Long, strtemp As String, strData As String
  Dim ByteArray(3) As Byte, fl As Long, lngKey As Long
  
    If (Len(strInputData) = 0) Or (Len(strKey) = 0) Then
        Exit Function '>---> Bottom
    End If
    
    lngKey = 0
    For lngCtr = 1 To Len(strKey)
        lngKey = lngKey + Asc(Mid$(strKey, lngCtr, 1))
    Next lngCtr
    
    ByteArray(0) = lngKey And 255
'    ByteArray(0) = (lngKey \ 2 ^ 1) And 255
    ByteArray(1) = (lngKey \ 2 ^ 8) And 255
    ByteArray(2) = (lngKey \ 2 ^ 16) And 255
    ByteArray(3) = (lngKey \ 2 ^ 24) And 255
    strData = strInputData
    fl = Len(strData)
    For lngCtr = 1 To fl
        strtemp = Chr$(Asc(Mid$(strData, lngCtr, 1)) Xor ByteArray((lngCtr - 1) Mod UBound(ByteArray)))
        Mid$(strData, lngCtr, 1) = strtemp
    Next lngCtr
    funDoXOR = strData
    Exit Function

funDoXORError:
    Call funErrorManager("modGlobalFunctions.funDoXORError")
    funDoXOR = ""
End Function

      
'*******************************************************************************
Function funStartDoc(strDocName As String) As Long
Dim lngScr_hDC As Long
On Error GoTo funStartDocError
    
    lngScr_hDC = GetDesktopWindow()
    funStartDoc = ShellExecute(lngScr_hDC, "Open", strDocName, "", "C:\", SW_SHOWNORMAL)
    Exit Function

funStartDocError:
    Call funErrorManager("modGlobalFunctions.funStartDoc")
End Function


'*******************************************************************************
Function funArraySize(varArray() As String) As Long
On Error Resume Next
    
    funArraySize = 0
    funArraySize = UBound(varArray) + 1
End Function


'*******************************************************************************
Sub subCdlOpenFile(cdlPath As CommonDialog, txtBox As TextBox, strDefaultPath As String, strFailsafePath As String, strFilter As String, strTitle As String, Optional blnFullPath As Boolean = True)
Dim strPath As String
On Error Resume Next
    
    strPath = Dir(strDefaultPath)
    If strPath <> "" And strDefaultPath <> "" Then
        cdlPath.FileName = strDefaultPath
        cdlPath.InitDir = strDefaultPath
    Else
        cdlPath.FileName = ""
        cdlPath.InitDir = strFailsafePath
    End If
    On Error GoTo canceled
    cdlPath.Filter = strFilter
    cdlPath.DialogTitle = strTitle
    Call cdlPath.ShowOpen
    If blnFullPath Then
        txtBox.Text = cdlPath.FileName
        txtBox.ToolTipText = cdlPath.FileName
        txtBox.SelStart = Len(txtBox.Text)
    Else
        txtBox.Text = Right(cdlPath.FileName, Len(cdlPath.FileName) - InStrRev(cdlPath.FileName, "\"))
        txtBox.SelStart = Len(txtBox.Text)
    End If

canceled:
End Sub


'*******************************************************************************
Function funSendFile(wsCnx As Winsock, strSourceFilePath As String, intActionType As enuFileActions, Optional blnWarnOnError As Boolean = True) As Boolean
Dim strBuf As String
Dim intfreefile As Integer
Dim lngRead As Long, lngLen As Long, lngThisRead As Long, lngLastRead As Long
Static strMissingFiles() As String
Static intNumMissingFiles As Integer
Static i As Integer
On Error GoTo funSendFileError
    
    If intNumMissingFiles > 0 Then
        For i = 0 To intNumMissingFiles - 1
            If strMissingFiles(i) = strSourceFilePath Then
                Exit Function
            End If
        Next i
    End If
    If Dir(strSourceFilePath) = "" Then Call Err.Raise(CYBERA_ERROR, "modGlobalFunctions.funSendFileError", "File not found")
    
    intfreefile = FreeFile
    
    ' Open file for binary access:
    Open strSourceFilePath For Binary Access Read As #intfreefile
    lngLen = LOF(intfreefile)
    
    funSendFile = funSendData(wsCnx, CNX_FILE_SEND & intActionType & CNX_PAYLOAD_TOKEN)
    
    ' Loop through the file, loading it up in chunks of 64k:
    Do While lngRead < lngLen
        lngThisRead = 65536
'        lngThisRead = 655360
        If lngThisRead + lngRead > lngLen Then
            lngThisRead = lngLen - lngRead
        End If
        If Not lngThisRead = lngLastRead Then
            strBuf = Space$(lngThisRead)
        End If
        Get #intfreefile, , strBuf
        lngRead = lngRead + lngThisRead
        lngLastRead = lngThisRead
        funSendFile = funSendData(wsCnx, strBuf)
'        strSend = strSend & strBuf
    Loop
    Close intfreefile
    
    funSendFile = funSendData(wsCnx, CNX_MSG_END)
    Exit Function

funSendFileError:
    Call funErrorManager("modGlobalFunctions.funSendFileError", , True)
    intNumMissingFiles = intNumMissingFiles + 1
    ReDim Preserve strMissingFiles(intNumMissingFiles - 1)
    strMissingFiles(UBound(strMissingFiles)) = strSourceFilePath
    If blnWarnOnError Then Call funMsg(FileNotFound, strSourceFilePath)
End Function


'*******************************************************************************
Function funLpad(strMyValue, strMyPadCharacter, lngMyPaddedLength) As String
Dim lngPadlength As Long
Dim strPadString As String
Dim x As Long
On Error GoTo funLpadError

   lngPadlength = lngMyPaddedLength - Len(strMyValue)
   For x = 1 To lngPadlength
      strPadString = strPadString & strMyPadCharacter
   Next
   funLpad = strPadString + strMyValue
    Exit Function

funLpadError:
    Call funErrorManager("modGlobalFunctions.funLpadError")
End Function


'*******************************************************************************
Function funHexToLong(ByVal strHex As String) As Long
On Error GoTo funHexToLongError
    
    funHexToLong = Val("&H" & strHex & "&")
    Exit Function

funHexToLongError:
    Call funErrorManager("modGlobalFunctions.funHexToLongError")
End Function


'*******************************************************************************
Function funErrorManager(strSource As String, Optional blnDisplayMsg As Boolean = False, Optional blnForbbidRaise As Boolean = False) As String
Dim lngErrNumber As Long
Dim strErrSource As String
Dim strErrDesc As String

    lngErrNumber = Err.Number
    strErrSource = Err.source
    strErrDesc = Err.Description
    Call subDebugPrint(strSource & ": #" & lngErrNumber & " | " & strErrDesc & " | " & strErrSource)
    Call subWriteToLogFile(strSource & ": #" & lngErrNumber & " | " & strErrDesc & " | " & strErrSource)
    
    If blnDisplayMsg Then
        Call funMsg(UntrappedError, strSource)
    Else
        If Not blnForbbidRaise Then Call Err.Raise(CYBERA_ERROR, strSource, strErrDesc)
    End If
    funErrorManager = lngErrNumber & " - " & strErrDesc
End Function


'*******************************************************************************
Sub subSetDefaultFont(frmTmp As Form)
Dim uTmpControl As Control
On Error Resume Next
    
    For Each uTmpControl In frmTmp.Controls
        uTmpControl.Font = ""
        uTmpControl.Font.Name = ""
    Next uTmpControl
End Sub


'*******************************************************************************
Sub subSetCharset(frmTmp As Form, Optional intCharset As Integer = 0)
Dim uTmpControl As Control
On Error Resume Next
    
    For Each uTmpControl In frmTmp.Controls
        uTmpControl.Font.Charset = intCharset
    Next uTmpControl
End Sub


'*******************************************************************************
Public Function funFindLangPacks() As String()
Dim strLags() As String
Dim strNewFileName  As String
Dim strNewResName  As String
Dim strNewLangName  As String
Dim strNewLangNameWestern  As String
Dim strNewTradName  As String
Dim strNewLangOK  As String
Dim strNewLangCharset  As String
Dim uResfile As Object
On Error GoTo funFindLangPacksError

    ReDim strLags(5, 0)
    strLags(0, 0) = App.EXEName & LANG_RES_EN
    strLags(1, 0) = LoadResString(10014)
    strLags(2, 0) = LoadResString(10045)
    strLags(3, 0) = CStr(0)
    strLags(4, 0) = LoadResString(1)
    strLags(5, 0) = True
    strNewFileName = Dir(App.Path & "\" & App.EXEName & "*.dll")
    Do While strNewFileName <> ""
        strNewResName = Left(strNewFileName, InStr(LCase(strNewFileName), LCase(".dll")) - 1)
        On Error Resume Next
        Set uResfile = CreateObject(strNewResName & ".clsResources")
        If Not IsEmpty(uResfile) Then
            strNewLangName = ""
            strNewTradName = ""
            strNewLangCharset = CStr(0)
            strNewLangOK = True
            
            strNewLangName = uResfile.GetLangName
            strNewLangNameWestern = uResfile.GetLangNameWestern
            strNewTradName = uResfile.GetTradName
            strNewLangCharset = CStr(uResfile.GetCharset)
            
            Select Case LCase(Right(strNewResName, 2))
            Case "el"
                If Asc(Left(strNewLangName, 1)) = 63 Then strNewLangOK = False
            End Select
            
            If strNewLangName <> "" Then
                If strLags(0, UBound(strLags, 2)) <> "" Then ReDim Preserve strLags(5, UBound(strLags, 2) + 1)
                strLags(0, UBound(strLags, 2)) = strNewResName
                strLags(1, UBound(strLags, 2)) = strNewLangName
                strLags(2, UBound(strLags, 2)) = strNewTradName
                strLags(3, UBound(strLags, 2)) = strNewLangCharset
                strLags(4, UBound(strLags, 2)) = strNewLangNameWestern
                strLags(5, UBound(strLags, 2)) = strNewLangOK
            End If
        End If
        On Error GoTo funFindLangPacksError
        strNewFileName = Dir
    Loop
    funFindLangPacks = strLags
    Exit Function
    
funFindLangPacksError:
    Call funErrorManager("modGlobalFunctions.funFindLangPacksError")
End Function


'*******************************************************************************
Function funSetFocus(objControl As Object) As Boolean
On Error GoTo funFindLangPacksError
    
    funSetFocus = True
    Call objControl.SetFocus
    Exit Function
    
funFindLangPacksError:
    funSetFocus = False
End Function


'*******************************************************************************
Public Function funGetMacAddress() As String
Dim strGuidbuffer(15) As Byte
Dim strGuidstring As String
On Error GoTo funGetMacAddressError
    
    strGuidstring = Space(39)
    Call UuidCreateSequential(strGuidbuffer(0))
    Call StringFromGUID2(strGuidbuffer(0), StrPtr(strGuidstring), 39)
    funGetMacAddress = Mid(strGuidstring, 26, 12)
    Exit Function
    
funGetMacAddressError:
    Call funErrorManager("modGlobalFunctions.funGetMacAddressError")
End Function


'*******************************************************************************
Function funIsInCollection(objCollection As Collection, strKey As String) As Boolean
On Error GoTo funFindInCollectionError
    
    Call objCollection.Item(strKey)
    funIsInCollection = True
    Exit Function
    
funFindInCollectionError:
    funIsInCollection = False
End Function


'*******************************************************************************
Function funStripChars(txtBox As TextBox, strForbiddenChars As String, Optional strReplace As String = "") As Boolean
Dim intSelStart As Integer, intBadIndex As Integer, i As Integer
Dim strSearch As String
On Error GoTo funStripCharsError

    If Len(strForbiddenChars) = 0 Then Exit Function
    funStripChars = True
    For i = 1 To Len(strForbiddenChars)
        strSearch = Mid(strForbiddenChars, i, 1)
        intSelStart = txtBox.SelStart
        intBadIndex = InStr(txtBox.Text, strSearch)
        If intBadIndex > 0 Then
            txtBox.Text = Replace(txtBox.Text, strSearch, strReplace)
            txtBox.SelStart = intSelStart - 1
            funStripChars = False
        End If
    Next i
    
    Exit Function
    
funStripCharsError:
    funStripChars = False
End Function


'*******************************************************************************
Function funStripNonNumericChars(txtBox As TextBox, Optional strReplace As String = "") As Boolean
Dim intSelStart As Integer, i As Integer
Dim strBox As String
Dim strChar As String
Dim blnIsNumeric As Boolean
On Error GoTo funStripCharsError

    funStripNonNumericChars = True
    intSelStart = txtBox.SelStart
    strBox = txtBox.Text
    For i = 1 To Len(strBox)
        strChar = Mid(strBox, i, 1)
        blnIsNumeric = IsNumeric(strChar)
        If Not blnIsNumeric Then
            funStripNonNumericChars = False
            strBox = Left(strBox, i - 1) & Replace(strBox, strChar, strReplace, i, 1)
            If intSelStart >= i Then intSelStart = intSelStart - 1 + Len(strReplace)
        End If
    Next i
    txtBox.Text = strBox
    txtBox.SelStart = intSelStart
    Exit Function
    
funStripCharsError:
    funStripNonNumericChars = False
End Function


'*******************************************************************************
Public Sub subDelay(pdblSeconds As Double)
Const OneSecond As Double = 1# / (1440# * 60#)
Dim dblWaitUntil As Date
On Error GoTo subDelayError

    dblWaitUntil = Now + OneSecond * pdblSeconds
    Do Until Now > dblWaitUntil
        Sleep 100
        DoEvents ' Allow windows message to be processed
    Loop
    Exit Sub
    
subDelayError:
End Sub


'*******************************************************************************
Function funFormatCurrencyString(sngAmount As Single, strCurrencyFormat As String, Optional strCurrencySymbol As String = "") As String
Dim strTmpCurrFormat As String
On Error GoTo funFormatCurrencyStringError
    
    If strCurrencySymbol = "" Then
        strTmpCurrFormat = Trim(Replace(strCurrencyFormat, CURRENCY_SYMBOL_INSERT, ""))
    Else
        If InStr(strCurrencyFormat, CURRENCY_SYMBOL_INSERT) = 0 Then strCurrencyFormat = strCurrencyFormat & " " & CURRENCY_SYMBOL_INSERT
        strTmpCurrFormat = Replace(strCurrencyFormat, CURRENCY_SYMBOL_INSERT, Chr(34) & strCurrencySymbol & Chr(34))
    End If
    funFormatCurrencyString = Trim(Format(sngAmount, strTmpCurrFormat))
    Do While InStr(funFormatCurrencyString, "  ") > 0
        funFormatCurrencyString = Replace(funFormatCurrencyString, "  ", " ")
    Loop
    Exit Function
    
funFormatCurrencyStringError:
    funFormatCurrencyString = ""
End Function



