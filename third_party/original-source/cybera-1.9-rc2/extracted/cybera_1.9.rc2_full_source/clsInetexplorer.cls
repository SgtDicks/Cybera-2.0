VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInetexplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ INTERNET EXPLORER CLASS ____________________|
'|  Name:          clsInetexplorer.cls                                         |
'|  Type:          VB6 class                                                   |
'|  Project:       Cybera Workstation (http://cybera.sf.net)                   |
'|  Description:   This class enables certain access to stated instances of    |
'|                 Internet Explorer                                           |
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

' $Id: clsInetexplorer.cls 15 2006-01-01 16:12:35Z sam $

Option Explicit
Option Base 1

'   Internal UDT to group info on an IE Instance
Private Type URL_INFO
    hWnd As Long            '  handle to the IE Window
    hWndEdit As Long      '  handle to the EditBox in the IE Window
    Caption As String       '  Window Title
    URL As String             '  Current URL being browsed
End Type

'   structure contains operating system version information
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

'   retrieves the name of the class to which the specified window belongs.
Private Declare Function apiGetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassname As String, ByVal nMaxCount As Long) As Long

'   retrieves a handle to a window that has the specified relationship (Z
'   order or owner) to the specified window.
Private Declare Function apiGetWindow Lib "user32" Alias "GetWindow" (ByVal hWnd As Long, ByVal wCmd As Long) As Long

'   copies the text of the specified window's title bar (if it has one) into a buffer.
'   If the specified window is a control, the text of the control is copied.
Private Declare Function apiGetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

'   returns a handle to the desktop window.
Private Declare Function apiGetDesktopWindow Lib "user32" Alias "GetDesktopWindow" () As Long

'   retrieves information about the specified window. The function
'   also retrieves the 32-bit (long) value at the specified offset into
'   the extra window memory of a window.
Private Declare Function apiGetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

'   retrieves a handle to a window whose class name and window name
'   match the specified strings. The function searches child windows,
'   beginning with the one following the specified child window.
Private Declare Function apiFindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

'   sends the specified message to a window or windows. The function
'   calls the window procedure for the specified window and does not
'   return until the window procedure has processed the message.
Private Declare Function apiSendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'   copies the status of the 256 virtual keys to the specified buffer.
Private Declare Function apiGetKeyboardState Lib "user32" Alias "GetKeyboardState" (pbKeyState As Byte) As Long

'   obtains extended information about the version of the operating
'   system that is currently running.
Private Declare Function apiGetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

'  GetWindow - The retrieved handle identifies the child window at the top of the
'   Z order if the specified window is a parent window; otherwise, the retrieved
'   handle is NULL. The function examines only child windows of the specified
'   window. It does not examine descendant windows.
Private Const GW_CHILD = 5

'   GetWindow - The retrieved handle identifies the window below the specified
'   window in the Z order. If the specified window is a topmost window,
'   the handle identifies the topmost window below the specified window.
'   If the specified window is a top-level window, the handle identifies the
'   top-level window below the specified window. If the specified window is
'   a child window, the handle identifies the sibling window below the
'   specified window.
Private Const GW_HWNDNEXT = 2

'   GetWindowLong - Retrieves the window styles.
Private Const GWL_STYLE = (-16)

'   //Window Styles
Private Const WS_VISIBLE = &H10000000   'Visible

'   //Window messages
'   An application sends a WM_GETTEXT message to copy the text that
'   corresponds to a window into a buffer provided by the caller.
Private Const WM_GETTEXT = &HD

'   An application sends a WM_GETTEXTLENGTH message to determine
'   the length, in characters, of the text associated with a window.
Private Const WM_GETTEXTLENGTH = &HE

'   An application sends a WM_SETTEXT message to set the text of a window.
Private Const WM_SETTEXT = &HC

'   The WM_KEYDOWN message is posted to the window with the keyboard
'   focus when a nonsystem key is pressed. A nonsystem key is a key
'   that is pressed when the ALT key is not pressed.
Private Const WM_KEYDOWN = &H100

'   //Keyboard constants
Private Const VK_RETURN = &HD

'   //GetVersionEx constants
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Const MAX_LEN = 255

'   various IE Windows (controls), mconIE_CLASS
'   is the main IE window, the rest are it's children
Private Const mconIE_CLASS = "CabinetWClass"
Private Const mconIE_CLASS_EXPLORE = "ExploreWClass"
Private Const IEWND_CLASS_FRAME = "IEFrame"
Private Const mconIE_EDIT = "Edit"
Private Const mconIE_WORKERW = "WorkerW"
Private Const mconIE_WORKERA = "WorkerA"
Private Const mconIE_REBAR = "ReBarWindow32"
Private Const mconIE_COMBO = "ComboBox"
Private Const mconIE_COMBOEx = "ComboBoxEx32"
Private Const mconCAPTION = " - Microsoft Internet Explorer"

'   internal array to hold IE information
Private matURLs() As URL_INFO

Public Sub NavigateTo(Index As Integer, NewURL As String)
'   Uses SendMessage to force an IE window with the hWnd
'   at dimension Index, navigate to a new URL
'
On Error GoTo ErrHandler
Dim hWnd As Long
Dim abytkeys(0 To 255) As Byte
    
    '   Retrieve the handle to the EditBox
    hWnd = CLng(fGetURLInfo(Index, 3))
    If hWnd > 0 Then
        '   Get the current state of the keyboard
        Call apiGetKeyboardState(abytkeys(0))
        '   send the new URL as the text for the editbox
        Call apiSendMessage(hWnd, WM_SETTEXT, 0, _
                        ByVal CStr(NewURL & vbNullChar))
        '   simulate Enter keypress to force IE to
        '   navigate to the new url
        Call apiSendMessage(hWnd, WM_KEYDOWN, VK_RETURN, abytkeys(VK_RETURN))
    End If
ExitHere:
    Exit Sub
ErrHandler:
    With Err
        Call .Raise(.Number, "InetExplorer.NavigateTo", .Description, .HelpFile, .HelpContext)
    End With
    Resume ExitHere
End Sub


'*******************************************************************************
Public Property Get ItemhWnd(Index As Integer) As Long
'   Returns the Handle of a Window from dimension Index
'   of local array
'
On Error GoTo ErrHandler
    ItemhWnd = CLng(fGetURLInfo(Index, 0))
ExitHere:
    Exit Property
ErrHandler:
    With Err
        .Raise .Number, "InetExplorer::ItemhWnd", .Description, .HelpFile, .HelpContext
    End With
    Resume ExitHere
End Property


'*******************************************************************************
Public Property Get ItemCaption(Index As Integer) As String
'   Returns the caption of a Window from dimension Index
'   of local array
'
On Error GoTo ErrHandler
    ItemCaption = fGetURLInfo(Index, 1)
ExitHere:
    Exit Property
ErrHandler:
    With Err
        .Raise .Number, "InetExplorer::ItemCaption", .Description, .HelpFile, .HelpContext
    End With
    Resume ExitHere
End Property


'*******************************************************************************
Public Property Get ItemURL(Index As Integer) As String
'   Returns the URL of a Window from dimension Index
'   of local array
'
On Error GoTo ErrHandler
    ItemURL = fGetURLInfo(Index, 2)
ExitHere:
    Exit Property
ErrHandler:
    With Err
        .Raise .Number, "InetExplorer::ItemURL", .Description, _
                    .HelpFile, .HelpContext
    End With
    Resume ExitHere
End Property


'*******************************************************************************
Public Sub Refresh()
'   Populates the internal array with IE Window information
'
On Error GoTo ErrHandler
Dim strCaption As String, lngLen As Long
Dim strClass As String, strURL As String
Dim hWnd As Long, lngStyle As Long
Dim hWndChild As Long, intHits As Integer
Dim intPos As Integer

    '   Since we are using ReDim Preserve afterwards,
    '   erase the array
    Erase matURLs
    '   Option Base is set to 1
    intHits = 1
    
    '   Start with the first child of the Desktop window
    hWnd = apiGetWindow(apiGetDesktopWindow(), _
                    GW_CHILD)
    '   Enumerate all open windows
    Do While Not hWnd = 0
        lngStyle = apiGetWindowLong(hWnd, GWL_STYLE)
        '   if the window is visible
        If lngStyle And WS_VISIBLE Then
            strClass = fGetClassName(hWnd)
            '   if the class name belongs to IE
            If strClass = mconIE_CLASS Or strClass = mconIE_CLASS_EXPLORE Or strClass = IEWND_CLASS_FRAME Then
                strCaption = fGetCaption(hWnd)
                '   trim out the trailing mconCaption
                intPos = InStr(1, strCaption, mconCAPTION)
                If intPos > 0 Then
                    strCaption = Left$(strCaption, intPos - 1)
                End If
                '   Find the first worker class child of the IE window
                '   For NT, use Unicode version
                If fIsNT() Then
                    hWndChild = apiFindWindowEx(hWnd, 0, mconIE_WORKERW, vbNullString)
                Else
                    hWndChild = apiFindWindowEx(hWnd, 0, mconIE_WORKERA, vbNullString)
                End If
                
                If hWndChild > 0 Then
                    '   Rebar is child of Worker  window
                    hWndChild = apiFindWindowEx(hWndChild, 0, mconIE_REBAR, vbNullString)
                End If
                If hWndChild > 0 Then
                    '   ComboboxEx is child of Rebar window
                    hWndChild = apiFindWindowEx(hWndChild, 0, mconIE_COMBOEx, vbNullString)
                End If
                If hWndChild > 0 Then
                    '   ComboBox is child of ComboBoxEx  Window
                    hWndChild = apiFindWindowEx(hWndChild, 0, mconIE_COMBO, vbNullString)
                End If
                If hWndChild > 0 Then
                    '   Edit class is child of ComboBox  window
                    hWndChild = apiFindWindowEx(hWndChild, 0, mconIE_EDIT, vbNullString)
                End If
                If hWndChild > 0 Then
                    '   Get the length of the URL in Editbox
                    lngLen = apiSendMessage(hWndChild, WM_GETTEXTLENGTH, 0, ByVal 0&)
                    strURL = Space$(lngLen + 1)
                    '   Get the URL itself
                    lngLen = apiSendMessage(hWndChild, WM_GETTEXT, lngLen + 1, ByVal strURL)
                    '   store the entry
                    ReDim Preserve matURLs(intHits)
                    With matURLs(intHits)
                        .Caption = strCaption
                        .hWnd = hWnd
                        .URL = Left$(strURL, lngLen)
                        .hWndEdit = hWndChild
                    End With
                    intHits = intHits + 1
                End If
            End If
        End If
        '   move on to the next window
        hWnd = apiGetWindow(hWnd, GW_HWNDNEXT)
    Loop
        
ExitHere:
    Exit Sub
ErrHandler:
    With Err
        .Raise .Number, "InetExplorer::Refresh", .Description, .HelpFile, .HelpContext
    End With
    Resume ExitHere
End Sub


'*******************************************************************************
Public Property Get Count() As Integer
'   Returns count of open IE windows
'
Dim intCount As Integer
    On Error Resume Next
    intCount = UBound(matURLs)
    If Err Then
        intCount = 0
    Else
        Count = intCount
    End If
End Property


'*******************************************************************************
Private Function fGetURLInfo(intIndex As Integer, intType As Integer) As String
'   Returns specific information about an IE window
'
On Error GoTo ErrHandler
Dim strOut As String

    Select Case intType
        Case 0:     'hWnd
            strOut = CStr(matURLs(intIndex).hWnd)
        Case 1:     'Caption
            strOut = matURLs(intIndex).Caption
        Case 2:     'URL
            strOut = matURLs(intIndex).URL
        Case 3:     'EditBox's hWnd
            strOut = CStr(matURLs(intIndex).hWndEdit)
    End Select
    fGetURLInfo = strOut
ExitHere:
    Exit Function
ErrHandler:
    fGetURLInfo = vbNullString
    Resume ExitHere
End Function


'*******************************************************************************
Private Function fGetClassName(hWnd As Long) As String
'   Returns the classname of a Window
'
Dim strBuffer As String
Dim lngCount As Long
    strBuffer = String$(MAX_LEN + 1, 0)
    lngCount = apiGetClassName(hWnd, strBuffer, MAX_LEN)
    If lngCount > 0 Then fGetClassName = Left$(strBuffer, lngCount)
End Function


'*******************************************************************************
Private Function fGetCaption(hWnd As Long) As String
'   Returns the caption of a Window
'
Dim strBuffer As String
Dim lngCount As Long
    strBuffer = String$(MAX_LEN + 1, 0)
    lngCount = apiGetWindowText(hWnd, strBuffer, MAX_LEN)
    If lngCount > 0 Then fGetCaption = Left$(strBuffer, lngCount)
End Function


'*******************************************************************************
Private Function fIsNT() As Boolean
'   Returns true if current platform is WiNNT
'
Dim tOSInfo  As OSVERSIONINFO
    tOSInfo.dwOSVersionInfoSize = Len(tOSInfo)
    Call apiGetVersionEx(tOSInfo)
    fIsNT = (tOSInfo.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function


'*******************************************************************************
Private Function fIsWin9x() As Boolean
'   Returns true if current platform is Win95 or Win98
'
Dim tOSInfo  As OSVERSIONINFO
    tOSInfo.dwOSVersionInfoSize = Len(tOSInfo)
    Call apiGetVersionEx(tOSInfo)
    fIsWin9x = (tOSInfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS)
End Function


'*******************************************************************************
Private Sub Class_Terminate()
    '   do cleanup here
    On Error Resume Next
    Erase matURLs
End Sub




