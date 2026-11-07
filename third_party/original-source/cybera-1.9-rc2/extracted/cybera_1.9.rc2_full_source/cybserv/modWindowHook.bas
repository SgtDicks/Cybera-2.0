Attribute VB_Name = "modWindowHook"
' _____________________________________________________________________________
'|________________________________ VAR DECLARATION ____________________________|
'|  Name:          modWindowHook.bas                                           |
'|  Type:          VB6 Module                                                  |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
'|  Description:   Contains Main window hooks code                             |
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

' $Id: modWindowHook.bas 15 2006-01-01 16:12:35Z sam $

Option Explicit

Global lpPrevWndProc As Long
Global gHW As Long

Private Const GWL_WNDPROC = -4
Private Const WM_GETMINMAXINFO = &H24

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub CopyMemoryToMinMaxInfo Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As MINMAXINFO, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Declare Sub CopyMemoryFromMinMaxInfo Lib "kernel32" Alias "RtlMoveMemory" (ByVal hpvDest As Long, hpvSource As MINMAXINFO, ByVal cbCopy As Long)


'*******************************************************************************
Public Sub subHook()
On Error Resume Next
    
    'Start subclassing.
    lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf funWindowProc)
End Sub


'*******************************************************************************
Public Sub subUnhook()
Dim temp As Long
On Error Resume Next

    'Cease subclassing.
    temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
End Sub


'*******************************************************************************
Function funWindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim MinMax As MINMAXINFO
On Error Resume Next

    'Check for request for min/max window sizes.
    If uMsg = WM_GETMINMAXINFO Then
        'Retrieve default MinMax settings
        Call CopyMemoryToMinMaxInfo(MinMax, lParam, Len(MinMax))

        'Specify new minimum size for window.
        MinMax.ptMinTrackSize.x = MAX_WIN_WIDTH
        MinMax.ptMinTrackSize.y = MAX_WIN_HEIGHT
        
        'Specify new maximum size for window.
'        MinMax.ptMaxTrackSize.x = 500
'        MinMax.ptMaxTrackSize.y = 500

        'Copy local structure back.
        Call CopyMemoryFromMinMaxInfo(lParam, MinMax, Len(MinMax))

        funWindowProc = DefWindowProc(hw, uMsg, wParam, lParam)
    Else
        funWindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
    End If
End Function



