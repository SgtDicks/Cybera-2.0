Attribute VB_Name = "modProcManager"
' _____________________________________________________________________________
'|________________________________ PROCESS MANAGEMENT FUNCTIONS _______________|
'|  Name:          modProcManager.bas                                          |
'|  Type:          VB6 Module                                                  |
'|  Project:       Cybera Server & Workstation (http://cybera.sf.net)          |
'|  Description:   Contains process management api calls and functions.        |
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

' $Id: modProcManager.bas 15 2006-01-01 16:12:35Z sam $

Option Explicit

Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal handle As Long) As Long
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function GetWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal wCmd As Long) As Integer
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpSting As String, ByVal nMaxCount As Long) As Integer
Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Integer
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Private Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindowVisible Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hWnd As Long, lpdwProcessId As Long) As Long


Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassname As String, ByVal lpWindowName As String) As Long

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long

Const WM_CLOSE = &H10
Const INFINITE = &HFFFFFFFF


Const WS_EX_TOOLWINDOW = &H80
Const WS_EX_APPWINDOW = &H40000


Const PROCESS_QUERY_INFORMATION = 1024
Const PROCESS_VM_READ = 16
Const MAX_PATH = 260
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const SYNCHRONIZE = &H100000
Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const ANYSIZE_ARRAY = 1
Const PROCESS_ALL_ACCESS = &H1F0FFF
Const SE_DEBUG_NAME = "SeDebugPrivilege"
Const SE_PRIVILEGE_ENABLED = &H2
Const GW_CHILD = 5
Const GW_HWNDFIRST = 0
Const GW_HWNDLAST = 1
Const GW_HWNDNEXT = 2
Const GW_HWNDPREV = 3
Const GW_OWNER = 4
Const GWL_EXSTYLE = (-20)


Private Type LARGE_INTEGER
   lowpart As Long
   highpart As Long
End Type

Private Type LUID
   lowpart As Long
   highpart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
   pLuid As LUID
   Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
   PrivilegeCount As Long
   Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type

Public Type typProc
    lngID As Long
    strName As String
    strPath As String
    strWinName As String
    lngWinHandle As Long
End Type

Private muProcList() As typProc
Private muWinList() As typProc


'*******************************************************************************
Public Function funGetPoccessList(Optional blnRefresh As Boolean = True) As typProc()
On Error GoTo funGetPoccessListError
    
    If blnRefresh Then Call subPopulatePoccessList
    funGetPoccessList = muProcList
    Exit Function

funGetPoccessListError:
    Call funErrorManager("modProcManager.funGetPoccessListError")
End Function


'*******************************************************************************
Public Function funGetWindowList(frmStart As Form, Optional blnRefresh As Boolean = True) As typProc()
On Error GoTo funGetWindowListError
    
    If blnRefresh Then Call subPopulateWindowList(frmStart)
    funGetWindowList = muWinList
    Exit Function

funGetWindowListError:
    Call funErrorManager("modProcManager.funGetWindowListError")
End Function


'*******************************************************************************
Private Sub subPopulatePoccessList()
Dim lngCb                   As Long
Dim lngCbNeeded             As Long
Dim lngNumElements          As Long
Dim lngProcessIDs()         As Long
Dim lngRet                  As Long
Dim intProcNum              As Integer
Dim i                       As Long
Dim uProc                   As typProc
On Error GoTo subPopulatePoccessListError
    
    intProcNum = 0
    ReDim muProcList(intProcNum)
    
    'Get the array containing the process id's for each process object
    lngCb = 8
    lngCbNeeded = 96
    Do While lngCb <= lngCbNeeded
       lngCb = lngCb * 2
       ReDim lngProcessIDs(lngCb / 4) As Long
       lngRet = EnumProcesses(lngProcessIDs(1), lngCb, lngCbNeeded)
    Loop
    
    'calculate how many process IDs were returned
    lngNumElements = lngCbNeeded / 4
    
    For i = 1 To lngNumElements
        uProc = funGetProcInfo(lngProcessIDs(i))
        If uProc.lngID > 0 Then
            ReDim Preserve muProcList(intProcNum)
            muProcList(intProcNum) = uProc
            intProcNum = intProcNum + 1
        End If
    Next
    Exit Sub

subPopulatePoccessListError:
    Call funErrorManager("modProcManager.subPopulatePoccessListError")
End Sub


'*******************************************************************************
Private Function funGetProcInfo(lngProcessID) As typProc
Dim lngCbNeededMod          As Long
Dim lngHModule(1 To 1024)   As Long
Dim strModuleName           As String
Dim lngHProcess             As Long
Dim lngRet                  As Long
Dim lngSize                 As Long
Dim uProc                   As typProc
On Error GoTo funGetProcInfoError
    
    'Get a handle to the Process
    lngHProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lngProcessID)
    
    ' Iterate through each process with an ID that <> 0
    If lngHProcess Then
        
        'Get an array of the module handles for the specified process
        lngRet = EnumProcessModules(lngHProcess, lngHModule(1), 1024, lngCbNeededMod)
        lngRet = EnumProcessModules(lngHProcess, lngHModule(1), lngCbNeededMod, lngCbNeededMod)
        
        'If the Module Array is retrieved, Get the ModuleFileName
        If lngRet <> 0 Then
            
            'Fill the ModuleName buffer with spaces
            strModuleName = Space(MAX_PATH)
    
            'Preset buffer size
            lngSize = 500
    
            'Get the module file name
    '                lngRet = GetModuleFileNameExA(lngHProcess, Modules(1), ModuleName, nSize)
            lngRet = GetModuleFileNameExA(lngHProcess, lngHModule(1), strModuleName, lngSize)
    
            'Get the module file name out of the buffer, lRet is how
            'many characters the string is, the rest of the buffer is spaces
            strModuleName = left(strModuleName, lngRet)
    
            'Update the proccess info
            uProc.lngID = lngProcessID
            uProc.strName = right(strModuleName, Len(strModuleName) - InStrRev(strModuleName, "\"))
            uProc.strPath = strModuleName
        Else
            uProc.lngID = -1
        End If
    End If
    'Close the handle to the process
    lngRet = CloseHandle(lngHProcess)

    funGetProcInfo = uProc
    Exit Function

funGetProcInfoError:
    Call funErrorManager("modProcManager.funGetProcInfoError")
End Function

' SeDebugSample shows how to grab the Process Handle of any running
' Application, and then use the SeDebug privilege to shut it down.

'*******************************************************************************
Public Sub subKillProc(ApplicationPID As Long)
Dim hProcessID As Long         ' Handle to your sample
Dim hProcess As Long           ' Handle to your current process
Dim hToken As Long             ' Handle to your process token.
Dim lPrivilege As Long         ' Privilege to enable/disable
Dim iPrivilegeflag As Boolean  ' Flag whether to enable/disable
                               ' the privilege of concern.
Dim lResult As Long            ' Result call of various APIs.
On Error Resume Next


   ' set the incoming PID to our internal variable
   hProcessID = ApplicationPID

   ' get our current process handle
   hProcess = GetCurrentProcess

   ' open the tokens for this process
   lResult = OpenProcessToken(hProcess, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken)

   ' if OpenProcessToken fails, the return result is zero, test for
   ' success here

   If (lResult = 0) Then
      CloseHandle (hToken)
      Exit Sub
   End If

   ' Now that you have the token for this process, you want to set
   ' the SE_DEBUG_NAME privilege.
   lResult = SetPrivilege(hToken, SE_DEBUG_NAME, True)

   ' Make sure you could set the privilege on this token.
   If (lResult = False) Then
      CloseHandle (hToken)
      Exit Sub
   End If

   ' Now that you have changed the privileges on the token,
   ' have some fun. You can now get a process handle to the
   ' process ID that you passed into this program, and
   ' demand whatever access you want on it!
   hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, hProcessID)

   ' Make sure you opened the process so you can do stuff with it
   If (hProcess = Null) Then
      CloseHandle (hToken)
      Exit Sub
   End If

   ' Now turn the SE_DEBUG_PRIV back off,
   lResult = SetPrivilege(hToken, SE_DEBUG_NAME, False)

   ' Make sure you succeeded in reversing the privilege!
   If (lResult = False) Then
      CloseHandle (hProcess)
      CloseHandle (hToken)
      Exit Sub
   End If

   ' Now you want to kill the application, which you can do since
   ' your process handle to the application includes full access to
   ' romp and roam - you got the process handle when you had the
   ' SE_DEBUG_NAME privilege enabled!
   lResult = TerminateProcess(hProcess, 0)

   ' Let's see the result, and go from there.
   If (lResult = 0) Then
      CloseHandle (hProcess)
      CloseHandle (hToken)
      Exit Sub
   End If

   ' Close our handles and get out of here.
   CloseHandle (hProcess)
   CloseHandle (hToken)
End Sub

' The SetPrivilege function will accept a handle to a token, a
' privilege, and a flag to either enable/disable that privilege. The
' function will attempt to perform the desired action upon the token
' returning TRUE if it succeeded, or FALSE if it failed.

'*******************************************************************************
Private Function SetPrivilege(hToken As Long, Privilege As String, bSetFlag As Boolean) As Boolean
Dim TP As TOKEN_PRIVILEGES          ' Used in getting the current
                                    ' token privileges
Dim TPPrevious As TOKEN_PRIVILEGES  ' Used in setting the new
                                    ' token privileges
Dim LUID As LUID                    ' Stores the Local Unique
                                    ' Identifier - refer to MSDN
Dim cbPrevious As Long              ' Previous size of the
                                    ' TOKEN_PRIVILEGES structure
Dim lResult As Long                 ' Result of various API calls
On Error GoTo SetPrivilegeError

   ' Grab the size of the TOKEN_PRIVILEGES structure,
   ' used in making the API calls.
   cbPrevious = Len(TP)

   ' Grab the LUID for the request privilege.
   lResult = LookupPrivilegeValue("", Privilege, LUID)

   ' If LoopupPrivilegeValue fails, the return result will be zero.
   ' Test to make sure that the call succeeded.
   If (lResult = 0) Then
      SetPrivilege = False
   End If

   ' Set up basic information for a call.
   ' You want to retrieve the current privileges
   ' of the token under concern before you can modify them.
   TP.PrivilegeCount = 1
   TP.Privileges(0).pLuid = LUID
   TP.Privileges(0).Attributes = 0
   SetPrivilege = lResult

   ' You need to acquire the current privileges first
   lResult = AdjustTokenPrivileges(hToken, -1, TP, Len(TP), _
                                  TPPrevious, cbPrevious)

   ' If AdjustTokenPrivileges fails, the return result is zero,
   ' test for success.
   If (lResult = 0) Then
      SetPrivilege = False
   End If

   ' Now you can set the token privilege information
   ' to what the user is requesting.
   TPPrevious.PrivilegeCount = 1
   TPPrevious.Privileges(0).pLuid = LUID

   ' either enable or disable the privilege,
   ' depending on what the user wants.
   Select Case bSetFlag
      Case True: TPPrevious.Privileges(0).Attributes = _
                 TPPrevious.Privileges(0).Attributes Or _
                 (SE_PRIVILEGE_ENABLED)
      Case False: TPPrevious.Privileges(0).Attributes = _
                  TPPrevious.Privileges(0).Attributes Xor _
                  (SE_PRIVILEGE_ENABLED And _
                  TPPrevious.Privileges(0).Attributes)
   End Select

   ' Call adjust the token privilege information.
   lResult = AdjustTokenPrivileges(hToken, -1, TPPrevious, _
                                  cbPrevious, TP, cbPrevious)

   ' Determine your final result of this function.
   If (lResult = 0) Then
      ' You were not able to set the privilege on this token.
      SetPrivilege = False
   Else
      ' You managed to modify the token privilege
      SetPrivilege = True
   End If
   Exit Function

SetPrivilegeError:
    Call funErrorManager("modProcManager.SetPrivilegeError")
End Function


'*******************************************************************************
Private Sub subPopulateWindowList(frmStart As Form)
On Error GoTo subPopulateWindowListError

    ReDim muWinList(0)
    Call EnumWindows(AddressOf funEnumWindowsCallBack, frmStart.hWnd)
   Exit Sub

subPopulateWindowListError:
    Call funErrorManager("modProcManager.subPopulateWindowListError")
End Sub


'*******************************************************************************
'
' This callback function is called by Windows (from
' the EnumWindows API call) for EVERY window that exists.
' It populates the listbox with a list of windows that we
' are interested in.
'
' Windows to display are those that:
'   -   are not this app's
'   -   are visible
'   -   do not have a parent
'   -   have no owner and are not Tool windows OR
'       have an owner and are App windows
'
Private Function funEnumWindowsCallBack(ByVal hWnd As Long, ByVal lParam As Long) As Long
Dim lReturn                 As Long
Dim lExStyle                As Long
Dim lpdwProcessId           As Long
Dim bNoOwner                As Boolean
Dim sWindowText             As String
Dim uProc                   As typProc
On Error GoTo funEnumWindowsCallBackError

    If hWnd <> lParam Then
        If IsWindowVisible(hWnd) Then
            If GetParent(hWnd) = 0 Then
                bNoOwner = (GetWindow(hWnd, GW_OWNER) = 0)
                lExStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
                
                If (((lExStyle And WS_EX_TOOLWINDOW) = 0) And bNoOwner) Or ((lExStyle And WS_EX_APPWINDOW) And Not bNoOwner) Then
                    '
                    ' Get the window's caption.
                    '
                    sWindowText = Space$(256)
                    lReturn = GetWindowText(hWnd, sWindowText, Len(sWindowText))
                    If lReturn Then
                        lReturn = GetWindowThreadProcessId(hWnd, lpdwProcessId)
                        uProc = funGetProcInfo(lpdwProcessId)
                        If uProc.lngID > 0 Then
                            If muWinList(UBound(muWinList)).lngID <> 0 Then ReDim Preserve muWinList(UBound(muWinList) + 1)
                            muWinList(UBound(muWinList)) = uProc
                            muWinList(UBound(muWinList)).strWinName = left(sWindowText, lReturn)
                            muWinList(UBound(muWinList)).lngWinHandle = hWnd
                        End If
                    End If
                End If
            End If
        End If
    End If
    funEnumWindowsCallBack = True
   Exit Function

funEnumWindowsCallBackError:
    Call funErrorManager("modProcManager.funEnumWindowsCallBackError")
End Function


'*******************************************************************************
Public Sub subCloseApp(uProc As typProc)
Dim lngResult As Long
Dim lngReturnValue As Long
Dim lpdwProcessId As Long
Dim lngWinHandle As Long
On Error GoTo subCloseAppError

    If uProc.lngID = 0 And uProc.lngWinHandle = 0 Then Exit Sub
    If uProc.lngID = 0 Then
        lngWinHandle = uProc.lngWinHandle
        lngResult = GetWindowThreadProcessId(uProc.lngWinHandle, lpdwProcessId)
        uProc = funGetProcInfo(lpdwProcessId)
        uProc.lngWinHandle = lngWinHandle
    End If
    
    If uProc.lngWinHandle > 0 Then
         lngReturnValue = PostMessage(uProc.lngWinHandle, WM_CLOSE, vbNull, vbNull)
         lngResult = WaitForSingleObject(uProc.lngWinHandle, INFINITE)

         'Does the handle still exist?
         DoEvents
         If IsWindow(uProc.lngWinHandle) = 1 And LCase(uProc.strName) <> "explorer.exe" Then
            'The handle still exists and it is not the main explorer.exe. Use the TerminateProcess function
            'to close all related processes to this handle.
            Call subKillProc(uProc.lngID)
         End If
    Else
        Call subKillProc(uProc.lngID)
    End If
   Exit Sub

subCloseAppError:
    Call funErrorManager("modProcManager.subCloseAppError")
End Sub


