Attribute VB_Name = "modApiDecl"
' _____________________________________________________________________________
'|________________________________ API CALL DEFS ______________________________|
'|  Name:          modApiDecl.bas                                              |
'|  Type:          VB6 Module                                                  |
'|  Project:       Cybera Server & Workstation (http://cybera.sf.net)          |
'|  Description:   Contains general WinAPI calls and their constants and types.|
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

' $Id: modApiDecl.bas 15 2006-01-01 16:12:35Z sam $

Option Explicit

'____________________________________________________________________________________________________
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

Public Const EWX_LOGOFF As Long = 0
Public Const EWX_SHUTDOWN As Long = 1
Public Const EWX_REBOOT As Long = 2
Public Const EWX_FORCE As Long = 4
Public Const EWX_POWEROFF As Long = 8

'____________________________________________________________________________________________________
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long

'____________________________________________________________________________________________________
Public Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long

'____________________________________________________________________________________________________
Public Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long

Public Type LUID
   UsedPart As Long
   IgnoredForNowHigh32BitPart As Long
End Type

'____________________________________________________________________________________________________
Public Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
      
Public Type TOKEN_PRIVILEGES
  PrivilegeCount As Long
  TheLuid As LUID
  Attributes As Long
End Type

'____________________________________________________________________________________________________
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

' SetWindowPos Flags
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering

Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

' SetWindowPos() hwndInsertAfter values
Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2


'____________________________________________________________________________________________________
Declare Function Shell_NotifyIcon Lib "shell32.dll" (ByVal dwMessage As Long, lpData As typNotifyIconData) As Long

'These three constants specify what you want to do
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIM_MODIFY = &H1

Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4

Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOUSEMOVE = &H200
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

Type typNotifyIconData
        cbSize As Long
        hWnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

'____________________________________________________________________________________________________
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'____________________________________________________________________________________________________
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'____________________________________________________________________________________________________
Declare Function SetThreadLocale Lib "kernel32" (ByVal Locale As Long) As Long

'____________________________________________________________________________________________________
Declare Function GetThreadLocale Lib "kernel32" () As Long

'____________________________________________________________________________________________________
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Public Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As Long, phkResult As Long) As Long

Public Const ERROR_SUCCESS = 0&
Public Const ERROR_ACCESS_DENIED = 5&
Public Const RPC_S_CALL_FAILED = 1726&

Public Const HKEY_CLASSES_ROOT As Long = &H80000000
Public Const HKEY_LOCAL_MACHINE As Long = &H80000002
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_USERS = &H80000003

Public Const MAX_PATH = 260
Public Const KEY_QUERY_VALUE As Long = &H1
Public Const KEY_NOTIFY As Long = &H10
Public Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Public Const READ_CONTROL As Long = &H20000
Public Const KEY_SET_VALUE As Long = &H2
Public Const KEY_CREATE_SUB_KEY As Long = &H4
Public Const STANDARD_RIGHTS_READ As Long = (READ_CONTROL)
Public Const STANDARD_RIGHTS_ALL As Long = &H1F0000
Public Const REG_OPTION_NON_VOLATILE As Long = 0       ' Key is preserved when system is rebooted
Public Const SYNCHRONIZE As Long = &H100000
Public Const KEY_CREATE_LINK As Long = &H20
Public Const KEY_READ As Long = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_ALL_ACCESS As Long = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Const REG_SZ As Long = 1                         ' Unicode nul terminated string
Public Const REG_DWORD As Long = 4                      ' 32-bit number
Public Const REG_BINARY As Long = 3                             ' Free form binary

Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type




'____________________________________________________________________________________________________
Public Declare Function GetUserDefaultLCID% Lib "kernel32" ()

'____________________________________________________________________________________________________
Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Public Const LOCALE_SDECIMAL = &HE         '  decimal separator
Public Const LOCALE_SCURRENCY = &H14        '  local monetary symbol

'____________________________________________________________________________________________________
Declare Function SHAddToRecentDocs Lib "Shell32" (ByVal lFlags As Long, ByVal lPv As Long) As Long

'____________________________________________________________________________________________________
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Long) As Long

Public Const SB_TOP = 6
Public Const SB_LINEDOWN = 1
Public Const EM_SCROLL = &HB5

Public Const LVA_ALIGNLEFT = &H1
Public Const LVM_ARRANGE = &H1016

'____________________________________________________________________________________________________
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long

'____________________________________________________________________________________________________
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassname As String, ByVal lpWindowName As String) As Long

Public Const GW_CHILD = 5

'____________________________________________________________________________________________________
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Const SPI_SETDESKWALLPAPER = 20


'____________________________________________________________________________________________________
Public Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type


'____________________________________________________________________________________________________
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const CREATE_NEW = 1
Public Const CREATE_ALWAYS = 2
Public Const OPEN_EXISTING = 3
Public Const OPEN_ALWAYS = 4
Public Const TRUNCATE_EXISTING = 5
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const FILE_FLAG_WRITE_THROUGH = &H80000000
Public Const FILE_FLAG_RANDOM_ACCESS = &H10000000
Public Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000

Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

'____________________________________________________________________________________________________
Public Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long

Public Const FILE_BEGIN = 0
Public Const FILE_CURRENT = 1
Public Const FILE_END = 2

'____________________________________________________________________________________________________
Public Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long

'____________________________________________________________________________________________________
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'____________________________________________________________________________________________________
Public Declare Function GetTickCount Lib "kernel32" () As Long
