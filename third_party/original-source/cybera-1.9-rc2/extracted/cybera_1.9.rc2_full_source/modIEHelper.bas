Attribute VB_Name = "modIEHelper"
' _____________________________________________________________________________
'|________________________________ IE HELPER FUNCTION MODULE __________________|
'|  Name:          modIEHelper.bas                                             |
'|  Type:          VB6 Module                                                  |
'|  Project:       Cybera Server                                               |
'|  Description:   Contains functions that interact with Internet Explorer     |
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

' $Id: modIEHelper.bas 15 2006-01-01 16:12:35Z sam $

Option Explicit

Private Const CLSID_CUrlHistory = "{3C374A40-BAE4-11CF-BF7D-00AA006946EE}"
Private Const CLSID_IUrlHistoryStg2 = "{AFA0DC11-C313-11D0-831A-00C04FD5AE38}"
Private Const IUrlHistoryStg2_Release As Long = 8&
Private Const IUrlHistoryStg2_ClearHistory As Long = 36&
Private Const CLSCTX_INPROC_SERVER As Long = 1&
Private Const CC_STDCALL As Long = 4&
Private Const S_OK As Long = 0&

Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Private Declare Function CoCreateInstance Lib "ole32" (rclsid As Any, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, riid As Any, pvarResult As Long) As Long
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvarResult As Long, ByVal oVft As Long, ByVal cc As Long, ByVal vtReturn As VbVarType, ByVal nParams As Long, pVarTypes As Long, pVarArgs As Long, pvarResult As Variant) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszGuid As Long, pGuid As Any) As Long


Private Const ERROR_CACHE_FIND_FAIL As Long = 0
Private Const ERROR_CACHE_FIND_SUCCESS As Long = 1
Private Const ERROR_FILE_NOT_FOUND As Long = 2
Private Const ERROR_ACCESS_DENIED As Long = 5
Private Const ERROR_INSUFFICIENT_BUFFER As Long = 122
Private Const MAX_PATH As Long = 260
Private Const MAX_CACHE_ENTRY_INFO_SIZE As Long = 4096

Private Const LMEM_FIXED As Long = &H0
Private Const LMEM_ZEROINIT As Long = &H40
Private Const LPTR As Long = (LMEM_FIXED Or LMEM_ZEROINIT)

Private Const NORMAL_CACHE_ENTRY As Long = &H1
Private Const EDITED_CACHE_ENTRY As Long = &H8
Private Const TRACK_OFFLINE_CACHE_ENTRY As Long = &H10
Private Const TRACK_ONLINE_CACHE_ENTRY As Long = &H20
Private Const STICKY_CACHE_ENTRY As Long = &H40
Private Const SPARSE_CACHE_ENTRY As Long = &H10000
Private Const COOKIE_CACHE_ENTRY As Long = &H100000
Private Const URLHISTORY_CACHE_ENTRY As Long = &H200000
Private Const URLCACHE_FIND_DEFAULT_FILTER As Long = NORMAL_CACHE_ENTRY Or COOKIE_CACHE_ENTRY Or URLHISTORY_CACHE_ENTRY Or TRACK_OFFLINE_CACHE_ENTRY Or TRACK_ONLINE_CACHE_ENTRY Or STICKY_CACHE_ENTRY
Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Type INTERNET_CACHE_ENTRY_INFO
   dwStructSize As Long
   lpszSourceUrlName As Long
   lpszLocalFileName As Long
   CacheEntryType  As Long
   dwUseCount As Long
   dwHitRate As Long
   dwSizeLow As Long
   dwSizeHigh As Long
   LastModifiedTime As FILETIME
   ExpireTime As FILETIME
   LastAccessTime As FILETIME
   LastSyncTime As FILETIME
   lpHeaderInfo As Long
   dwHeaderInfoSize As Long
   lpszFileExtension As Long
   dwExemptDelta  As Long
End Type

Private Declare Function FindFirstUrlCacheEntry Lib "wininet" Alias "FindFirstUrlCacheEntryA" (ByVal lpszUrlSearchPattern As String, lpFirstCacheEntryInfo As Any, lpdwFirstCacheEntryInfoBufferSize As Long) As Long
Private Declare Function FindNextUrlCacheEntry Lib "wininet" Alias "FindNextUrlCacheEntryA" (ByVal hEnumHandle As Long, lpNextCacheEntryInfo As Any, lpdwNextCacheEntryInfoBufferSize As Long) As Long
Private Declare Function FindCloseUrlCache Lib "wininet" (ByVal hEnumHandle As Long) As Long
Private Declare Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Function lstrcpyA Lib "kernel32" (ByVal RetVal As String, ByVal Ptr As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal Ptr As Any) As Long
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long


'*******************************************************************************
Public Function funIEDeleteHistory() As Boolean
Dim objClsid As GUID
Dim idClsid As GUID
Dim lngResult  As Long
Dim lngRet As Long
On Error GoTo funIEDeleteHistoryError
   
    Call CLSIDFromString(StrPtr(CLSID_CUrlHistory), objClsid)
    Call CLSIDFromString(StrPtr(CLSID_IUrlHistoryStg2), idClsid)
    If CoCreateInstance(objClsid, 0&, CLSCTX_INPROC_SERVER, idClsid, lngResult) = S_OK Then
        If DispCallFunc(lngResult, IUrlHistoryStg2_ClearHistory, CC_STDCALL, vbLong, 0&, 0&, 0&, lngRet) = S_OK Then
            If DispCallFunc(lngResult, IUrlHistoryStg2_Release, CC_STDCALL, vbLong, 0&, 0&, 0&, lngRet) = S_OK Then
                funIEDeleteHistory = True
            End If
        End If
    End If
    Exit Function

funIEDeleteHistoryError:
    Call funErrorManager("modIEHelper.funIEDeleteHistoryError")
End Function


'*******************************************************************************
Public Sub subIEDeleteCache()
Dim icei As INTERNET_CACHE_ENTRY_INFO
Dim hFile As Long
Dim cachefile As String
Dim posUrl As Long
Dim posEnd As Long
Dim dwBuffer As Long
Dim pntrICE As Long
On Error GoTo subIEDeleteCacheError
   
    hFile = FindFirstUrlCacheEntry(0&, ByVal 0, dwBuffer)
    If (hFile = ERROR_CACHE_FIND_FAIL) And (Err.LastDllError = ERROR_INSUFFICIENT_BUFFER) Then
        pntrICE = LocalAlloc(LMEM_FIXED, dwBuffer)
        If pntrICE <> 0 Then
            CopyMemory ByVal pntrICE, dwBuffer, 4
            hFile = FindFirstUrlCacheEntry(vbNullString, ByVal pntrICE, dwBuffer)
            If hFile <> ERROR_CACHE_FIND_FAIL Then
                Do
                    Call CopyMemory(icei, ByVal pntrICE, Len(icei))
                    If (icei.CacheEntryType And NORMAL_CACHE_ENTRY) = NORMAL_CACHE_ENTRY Then
                        cachefile = GetStrFromPtrA(icei.lpszSourceUrlName)
'                        If InStr(cachefile, "Cookie") = 0 Then
                            Call DeleteUrlCacheEntry(cachefile)
'                        End If
                    End If
                    Call LocalFree(pntrICE)
                    dwBuffer = 0
                    Call FindNextUrlCacheEntry(hFile, ByVal 0, dwBuffer)
                    pntrICE = LocalAlloc(LMEM_FIXED, dwBuffer)
                    Call CopyMemory(ByVal pntrICE, dwBuffer, 4)
                Loop While FindNextUrlCacheEntry(hFile, ByVal pntrICE, dwBuffer)
            End If
        End If
    End If
    Call LocalFree(pntrICE)
    Call FindCloseUrlCache(hFile)
    Exit Sub

subIEDeleteCacheError:
    Call funErrorManager("modIEHelper.subIEDeleteCacheError")
End Sub


'*******************************************************************************
Private Function GetStrFromPtrA(ByVal lpszA As Long) As String
On Error GoTo GetStrFromPtrAError

   GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
   Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)
    Exit Function

GetStrFromPtrAError:
    Call funErrorManager("modIEHelper.GetStrFromPtrAError")
End Function



