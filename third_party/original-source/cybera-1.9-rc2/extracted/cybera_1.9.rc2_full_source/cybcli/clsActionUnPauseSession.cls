VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsActionUnPauseSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ WORKSTATION MSG ACTION CLASS _______________|
'|  Name:          clsActionUnPauseSession.cls                                 |
'|  Type:          VB6 class                                                   |
'|  Project:       Cybera Workstation (http://cybera.sf.net)                   |
'|  Description:                                                               |
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

' $Id: clsActionUnPauseSession.cls 15 2006-01-01 16:12:35Z sam $

Option Explicit


'*******************************************************************************
Public Sub subExecuteCommandButton1()
On Error GoTo subExecuteCommandButton1Error
    
    Call funUnPauseSession
    Exit Sub

subExecuteCommandButton1Error:
    Call funErrorManager("clsActionUnPauseSession.subExecuteCommandButton1Error")
End Sub


'*******************************************************************************
Public Sub subExecuteCommandButton2()
    
End Sub


