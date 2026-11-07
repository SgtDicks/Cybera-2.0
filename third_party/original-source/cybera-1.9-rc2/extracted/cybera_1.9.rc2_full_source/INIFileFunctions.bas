Attribute VB_Name = "modINIFileFunctions"
' _____________________________________________________________________________
'|________________________________ INI FILE FUNCTIONS _________________________|
'|  Name:          INIFileFunctions.bas                                        |
'|  Type:          VB6 Module                                                  |
'|  Project:       Cybera Server & Workstation (http://cybera.sf.net)          |
'|  Description:   Contains functions that Emulate GetSetting and SaveSetting  |
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

' $Id: INIFileFunctions.bas 15 2006-01-01 16:12:35Z sam $

Option Explicit


'=== Constant deffinitions =====================================================
Private Const mstrINIIntroString = "Cybera - Cyber Centre Administration"


'*******************************************************************************
Public Function GetINISettings(strAppFile As String, strHeadding As String, strKey As String, Optional Operator As Variant) As Variant
'-Retreaves data from INI file under a given Headding and Key.
'If the Headding / Key does not exist, a delault value is returned.
Dim strFilePath As String
Dim intFileNumber As Integer
Dim strNextLine As String
Dim blHeaddingFound As Boolean
Dim blKeyFound As Boolean
On Error GoTo GetINISettingsError

    strFilePath = strAppFile
    If Dir(strFilePath, vbNormal) <> "" Then
        intFileNumber = FreeFile
        Open strFilePath For Input As intFileNumber
        Do Until EOF(intFileNumber) Or blHeaddingFound
            Line Input #intFileNumber, strNextLine
            If strNextLine = "[" & strHeadding & "]" Then blHeaddingFound = True
        Loop
        strNextLine = ""
        If blHeaddingFound Then
            Do Until EOF(intFileNumber) Or left(strNextLine, 1) = "[" Or blKeyFound
                Line Input #intFileNumber, strNextLine
                If left(strNextLine, Len(strKey)) = strKey Then
                    blKeyFound = True
                    GetINISettings = right(strNextLine, Len(strNextLine) - Len(strKey) - 1)
                End If
            Loop
        End If
    End If
    If Not blHeaddingFound Or Not blKeyFound Or GetINISettings = "" Then
        If IsMissing(Operator) Then
            GetINISettings = 0
        Else
            GetINISettings = Operator
        End If
    End If
    Close intFileNumber
    Exit Function

GetINISettingsError:
    Call funErrorManager("modINIFileFunctions.GetINISettingsError")
End Function


'*******************************************************************************
Public Function SetINISettings(strAppFile As String, strHeadding As String, strKey As String, Operator As Variant) As Boolean
'-Writes data to an INI file under a given Headding and Key.
'If the file / Headding / Key does not exist it gets created.
Dim strFilePath As String
Dim intFileNumber As Integer
Dim strNextLine As String
Dim blHeaddingFound As Boolean
Dim blKeyFound As Boolean
Dim strOutputFile As String
Dim i As Integer
On Error GoTo SetINISettingsError

    SetINISettings = True
    strOutputFile = ""
    strFilePath = strAppFile
    If Dir(strFilePath, vbNormal) <> "" Then
        intFileNumber = FreeFile
        Open strFilePath For Input As intFileNumber
        Do Until EOF(intFileNumber)
            Line Input #intFileNumber, strNextLine
            If strNextLine <> "" Then
                If left(strNextLine, 1) = "[" Then
                    strOutputFile = _
                        strOutputFile & vbCrLf & strNextLine & vbCrLf
                Else
                    strOutputFile = strOutputFile & strNextLine & vbCrLf
                End If
                If Not blHeaddingFound And _
                    strNextLine = "[" & strHeadding & "]" Then
                    blHeaddingFound = True
                    strNextLine = ""
                    Do Until EOF(intFileNumber) Or _
                        left(strNextLine, 1) = "[" Or blKeyFound
                        Line Input #intFileNumber, strNextLine
                        If strNextLine <> "" Then
                            If (left(strNextLine, Len(strKey)) = strKey) Then
                                blKeyFound = True
                                strOutputFile = strOutputFile & strKey & "=" & Operator & vbCrLf
                            ElseIf Not left(strNextLine, 1) = "[" Then
                                strOutputFile = _
                                    strOutputFile & strNextLine & vbCrLf
                            End If
                        End If
                    Loop
                    If Not blKeyFound Then
                        If EOF(intFileNumber) Then
                            strOutputFile = _
                                strOutputFile & strKey & "=" & Operator & vbCrLf
                        ElseIf left(strNextLine, 1) = "[" Then
                            strOutputFile = strOutputFile & strKey & "=" & _
                                Operator & vbCrLf & vbCrLf & strNextLine & _
                                vbCrLf
                        End If
                    End If
                End If
            End If
        Loop
        If Not blHeaddingFound Then
            SetINISettings = False
            strOutputFile = _
                strOutputFile & vbCrLf & "[" & strHeadding & "]" & vbCrLf & _
                strKey & "=" & Operator & vbCrLf
        End If
    Else
        SetINISettings = False
        strOutputFile = _
         ";" & mstrINIIntroString & vbCrLf & ";File: " & strAppFile & vbCrLf & vbCrLf & "[" & strHeadding & _
         "]" & vbCrLf & strKey & "=" & Operator & vbCrLf
    End If
    Close #intFileNumber
    intFileNumber = FreeFile
    Open strFilePath For Output As intFileNumber
    Print #intFileNumber, left(strOutputFile, Len(strOutputFile) - 2);
    Close #intFileNumber
    Exit Function

SetINISettingsError:
    Call funErrorManager("modINIFileFunctions.SetINISettingsError")
End Function


'*******************************************************************************
Public Function CanWriteToFile(strAppFile) As Boolean
Dim strFilePath As String
Dim intFileNumber As Integer
On Error GoTo CanWriteToFileError

    CanWriteToFile = True
    If Dir(strAppFile) <> "" Then
        strFilePath = strAppFile
    Else
        strFilePath = App.path & "\" & strAppFile
    End If
    If Dir(strFilePath, vbNormal) <> "" Then
        intFileNumber = FreeFile
'        Open strFilePath For Output Lock Write As intFileNumber
        Open strFilePath For Append As intFileNumber

        Close #intFileNumber
    End If
    Exit Function
    
CanWriteToFileError:
    CanWriteToFile = False
End Function


'*******************************************************************************
Public Sub SetRegSettings(ByVal strAppKeyName As String, ByVal strHeadding As String, ByVal strKey As String, ByVal strValue As String, Optional ByVal lngKeyType As Long = REG_SZ, Optional ByVal lngNumDigits As Long = 4)
Dim hKey As Long, nResult As Long, lngValue As Long
On Error GoTo SetRegSettingsError
    
    If IsNumeric(strValue) And lngKeyType <> REG_SZ Then lngValue = CLng(strValue)
    
    Call RegCreateKeyEx(HKEY_LOCAL_MACHINE, "Software\" & strAppKeyName & "\" & strHeadding, 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, nResult)
    Select Case lngKeyType
    Case REG_SZ
        Call RegSetValueEx(hKey, strKey, 0&, REG_SZ, ByVal strValue, Len(strValue))
    
    Case REG_DWORD
        Call RegSetValueEx(hKey, strKey, 0&, REG_DWORD, lngValue, lngNumDigits)
    
    Case REG_BINARY
        Call RegSetValueEx(hKey, strKey, 0&, REG_BINARY, lngValue, lngNumDigits)
    
    End Select
    Call RegCloseKey(hKey)
    Exit Sub

SetRegSettingsError:
    Call funErrorManager("modINIFileFunctions.SetRegSettingsError: Headding: " & strHeadding & " - Key: " & strKey & " - Value: " & strValue)
End Sub


'*******************************************************************************
Public Sub DelRegSettings(ByVal strAppKeyName As String, ByVal strHeadding As String, ByVal strKey As String)
Dim hKey As Long, nResult As Long, lngValue As Long
On Error GoTo DelRegSettingsError
    
    Call RegCreateKeyEx(HKEY_LOCAL_MACHINE, "Software\" & strAppKeyName & "\" & strHeadding, 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, nResult)
    Call RegDeleteValue(hKey, strKey)
    Call RegCloseKey(hKey)
    Exit Sub

DelRegSettingsError:
    Call funErrorManager("modINIFileFunctions.DelRegSettingsError")
End Sub


'*******************************************************************************
Public Sub DelRegUserSettings(ByVal strAppKeyName As String, ByVal strHeadding As String, ByVal strKey As String)
Dim hKey As Long, nResult As Long, lngValue As Long
On Error GoTo DelRegUserSettingsError
    
    Call RegCreateKeyEx(HKEY_CURRENT_USER, "Software\" & strAppKeyName & "\" & strHeadding, 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, nResult)
    Call RegDeleteValue(hKey, strKey)
    Call RegCloseKey(hKey)
    Exit Sub

DelRegUserSettingsError:
    Call funErrorManager("modINIFileFunctions.DelRegUserSettingsError")
End Sub


'*******************************************************************************
Public Sub SetRegUserSettings(ByVal strAppKeyName As String, ByVal strHeadding As String, ByVal strKey As String, ByVal strValue As String, Optional ByVal lngKeyType As Long = REG_SZ, Optional ByVal lngNumDigits As Long = 4)
Dim hKey As Long, nResult As Long, lngValue As Long
On Error GoTo SetRegUserSettingsError
    
    If IsNumeric(strValue) And lngKeyType <> REG_SZ Then lngValue = CLng(strValue)
    
    Call RegCreateKeyEx(HKEY_CURRENT_USER, "Software\" & strAppKeyName & "\" & strHeadding, 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, nResult)
    Select Case lngKeyType
    Case REG_SZ
        Call RegSetValueEx(hKey, strKey, 0&, REG_SZ, ByVal strValue, Len(strValue))
    
    Case REG_DWORD
        Call RegSetValueEx(hKey, strKey, 0&, REG_DWORD, lngValue, lngNumDigits)
    
    Case REG_BINARY
        Call RegSetValueEx(hKey, strKey, 0&, REG_BINARY, lngValue, lngNumDigits)
    
    End Select
    Call RegCloseKey(hKey)
    Exit Sub

SetRegUserSettingsError:
    Call funErrorManager("modINIFileFunctions.SetRegUserSettingsError: Headding: " & strHeadding & " - Key: " & strKey & " - Value: " & strValue)
End Sub


'*******************************************************************************
Public Function GetRegSettings(ByVal strAppKeyName As String, ByVal strHeadding As String, ByVal strKey As String, ByVal varDefaultValue As Variant, Optional ByVal lngKeyType As Long = REG_SZ) As Variant
Dim hKey As Long, nBytes As Long
Dim strValue As String
Dim lngValue As Long
Dim hResult As Long
On Error GoTo GetRegSettingsError

    hResult = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\" & strAppKeyName & "\" & strHeadding, 0, KEY_ALL_ACCESS, hKey)
        
    Select Case lngKeyType
    Case REG_SZ
        hResult = RegQueryValueEx(hKey, strKey, 0&, REG_SZ, ByVal strValue, nBytes)
        strValue = Space$(nBytes)
        hResult = RegQueryValueEx(hKey, strKey, 0&, REG_SZ, ByVal strValue, Len(strValue))
        
        If nBytes > 0 Then
            GetRegSettings = left$(strValue, nBytes - 1)
        Else
            GetRegSettings = varDefaultValue
        End If
    
    Case REG_DWORD
        hResult = RegQueryValueEx(hKey, strKey, 0&, REG_DWORD, lngValue, 4)
        If hResult = ERROR_SUCCESS Then
            GetRegSettings = lngValue
        Else
            GetRegSettings = varDefaultValue
        End If
    End Select
    Call RegCloseKey(hKey)
    Exit Function

GetRegSettingsError:
    Call funErrorManager("modINIFileFunctions.GetRegSettingsError")
End Function


'*******************************************************************************
Public Function GetRegUserSettings(ByVal strAppKeyName As String, ByVal strHeadding As String, ByVal strKey As String, ByVal varDefaultValue As Variant, Optional ByVal lngKeyType As Long = REG_SZ) As Variant
Dim hKey As Long, nBytes As Long
Dim strValue As String
Dim lngValue As Long
Dim hResult As Long
On Error GoTo GetRegUserSettingsError

    hResult = RegOpenKeyEx(HKEY_CURRENT_USER, "Software\" & strAppKeyName & "\" & strHeadding, 0, KEY_ALL_ACCESS, hKey)
    Select Case lngKeyType
    Case REG_SZ
        hResult = RegQueryValueEx(hKey, strKey, 0&, REG_SZ, ByVal strValue, nBytes)
        strValue = Space$(nBytes)
        hResult = RegQueryValueEx(hKey, strKey, 0&, REG_SZ, ByVal strValue, Len(strValue))
        If nBytes > 0 Then
            GetRegUserSettings = left$(strValue, nBytes - 1)
        Else
            GetRegUserSettings = varDefaultValue
        End If
    
    Case REG_DWORD
        hResult = RegQueryValueEx(hKey, strKey, 0&, REG_DWORD, lngValue, 4)
        If hResult = ERROR_SUCCESS Then
            GetRegUserSettings = lngValue
        Else
            GetRegUserSettings = varDefaultValue
        End If
    
    End Select
    Call RegCloseKey(hKey)
    Exit Function

GetRegUserSettingsError:
    Call funErrorManager("modINIFileFunctions.GetRegUserSettingsError")
End Function

