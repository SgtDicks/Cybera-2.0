Attribute VB_Name = "modWebAccess"
' _____________________________________________________________________________
'|________________________________ WEB ACCESS _________________________________|
'|  Name:          modWebAccess.bas                                            |
'|  Type:          VB6 Module                                                  |
'|  Project:       Cybera Server & Workstation (http://cybera.sf.net)          |
'|  Description:   Contains Web access functions.                              |
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

' $Id: modWebAccess.bas 15 2006-01-01 16:12:35Z sam $

Option Explicit


Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1


'*******************************************************************************
Public Sub subGotoWeb(strURL)
Dim lngSuccess As Long
On Error GoTo subGotoWebError

    lngSuccess = ShellExecute(0&, vbNullString, strURL, vbNullString, "C:\", SW_SHOWNORMAL)
    Exit Sub
    
subGotoWebError:
    Call funErrorManager("modWebAccess.subGotoWebError")
End Sub


'*******************************************************************************
Public Sub subSendEmail(strEmail)
Dim lngSuccess As Long
On Error GoTo subSendEmailError

    lngSuccess = ShellExecute(0&, vbNullString, "mailto:" & strEmail, vbNullString, "C:\", SW_SHOWNORMAL)
    Exit Sub
    
subSendEmailError:
    Call funErrorManager("modWebAccess.subSendEmailError")
End Sub

