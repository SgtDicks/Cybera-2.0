Attribute VB_Name = "modWinLock"
' _____________________________________________________________________________
'|________________________________ WINDOWS LOCKS ______________________________|
'|  Name:          modWinLock.bas                                              |
'|  Type:          VB6 Module                                                  |
'|  Project:       Cybera Workstation (http://cybera.sf.net)                   |
'|  Description:   Contains windows locking functions.                         |
'|                 This file uses CybWinLock.dll based on work by A. Miguel    |
'|                 Feijão that can  be found at:                               |
'|                 http://www.codeproject.com/win32/AntonioWinLock.asp         |
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

' $Id: modWinLock.bas 15 2006-01-01 16:12:35Z sam $

Option Explicit

Private Declare Function Desktop_Show_Hide Lib "CybWinLock.dll" (ByVal bShowHide As Boolean) As Integer
Private Declare Function Taskbar_Show_Hide Lib "CybWinLock.dll" (ByVal bShowHide As Boolean) As Integer
Private Declare Function TaskSwitching_Enable_Disable Lib "CybWinLock.dll" (ByVal bEnableDisable As Boolean) As Integer
Private Declare Function CtrlAltDel_Enable_Disable Lib "CybWinLock.dll" (ByVal bEnableDisable As Boolean) As Integer


'*******************************************************************************
Public Function funToggleTaskbar(blnToggle As Boolean) As Integer
On Error GoTo funToggleTaskbarError

    If blnToggle Then
        funToggleTaskbar = Taskbar_Show_Hide(True)
    Else
        funToggleTaskbar = Taskbar_Show_Hide(False)
    End If
    Exit Function
    
funToggleTaskbarError:
    Call funErrorManager("modWinLock.funToggleTaskbarError")
End Function


'*******************************************************************************
Public Function funToggleDesktop(blnToggle As Boolean) As Integer
On Error GoTo funToggleDesktopError

    If blnToggle Then
        funToggleDesktop = Desktop_Show_Hide(True)
    Else
        funToggleDesktop = Desktop_Show_Hide(False)
    End If
    Exit Function
    
funToggleDesktopError:
    Call funErrorManager("modWinLock.funToggleDesktopError")
End Function


'*******************************************************************************
Public Function funToggleTaskSwitching(blnToggle As Boolean) As Integer
On Error GoTo funToggleTaskSwitchingError

    If blnToggle Then
        funToggleTaskSwitching = TaskSwitching_Enable_Disable(True)
    Else
        funToggleTaskSwitching = TaskSwitching_Enable_Disable(False)
    End If
    Exit Function
    
funToggleTaskSwitchingError:
    Call funErrorManager("modWinLock.funToggleTaskSwitchingError")
End Function


'*******************************************************************************
Public Function funToggleCtrlAltDel(blnToggle As Boolean) As Integer
On Error GoTo funToggleCtrlAltDelError

    If blnToggle Then
        funToggleCtrlAltDel = CtrlAltDel_Enable_Disable(True)
    Else
        funToggleCtrlAltDel = CtrlAltDel_Enable_Disable(False)
    End If
    Exit Function
    
funToggleCtrlAltDelError:
    Call funErrorManager("modWinLock.funToggleCtrlAltDelError")
End Function


