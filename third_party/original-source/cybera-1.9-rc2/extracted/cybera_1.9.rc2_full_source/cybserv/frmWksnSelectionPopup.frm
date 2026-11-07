VERSION 5.00
Begin VB.Form frmWksnSelectionPopup 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4155
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstWorkstations 
      Height          =   1635
      IntegralHeight  =   0   'False
      ItemData        =   "frmWksnSelectionPopup.frx":0000
      Left            =   0
      List            =   "frmWksnSelectionPopup.frx":0007
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   0
      Width           =   4155
   End
End
Attribute VB_Name = "frmWksnSelectionPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ WORKSTATION SELECTION POPUP ________________|
'|  Name:          frmWksnSelectionPopup.frm                                   |
'|  Type:          VB6 Form                                                    |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
'|  Description:   Workstation selection Popup used in the console             |
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

' $Id: frmWksnSelectionPopup.frm 15 2006-01-01 16:12:35Z sam $

Option Explicit


'*******************************************************************************
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Form_KeyDownError

    If KeyCode = 27 Or KeyCode = 13 Then Call Unload(Me)
    Exit Sub

Form_KeyDownError:
    Call funErrorManager("frmWksnSelectionPopup.Form_KeyDownError", True)
End Sub


'*******************************************************************************
Private Sub Form_Load()
Dim uWorkstation As New clsWorkstation
On Error GoTo Form_LoadError

    Call subSetDefaultFont(Me)
    Call subSetCharset(Me, guServOpts.uConfig.intInterfaceLangCharset)
    
    Call lstWorkstations.Clear
    For Each uWorkstation In guWorkstations
        Call lstWorkstations.AddItem(uWorkstation.WorkstationName)
        lstWorkstations.Selected(lstWorkstations.ListCount - 1) = uWorkstation.AllowGroupControl
    Next uWorkstation
    Exit Sub

Form_LoadError:
    Call funErrorManager("frmWksnSelectionPopup.Form_LoadError", True)
End Sub


'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
Dim blnAllActive As Boolean
On Error GoTo Form_UnloadError

    blnAllActive = True
    For i = 0 To lstWorkstations.ListCount - 1
        If lstWorkstations.Selected(i) = False Then blnAllActive = False
    Next i
    If blnAllActive Then frmMain.optConsoleWorkstationsAll.value = True
    Exit Sub

Form_UnloadError:
    Call funErrorManager("frmWksnSelectionPopup.Form_UnloadError", True)
End Sub


'*******************************************************************************
Private Sub lstWorkstations_ItemCheck(Item As Integer)
Dim uWorkstation As New clsWorkstation
On Error GoTo lstWorkstations_ItemCheckError

    For Each uWorkstation In guWorkstations
        If uWorkstation.WorkstationName = lstWorkstations.List(Item) Then uWorkstation.AllowGroupControl = lstWorkstations.Selected(Item)
    Next uWorkstation
    Exit Sub

lstWorkstations_ItemCheckError:
    Call funErrorManager("frmWksnSelectionPopup.lstWorkstations_ItemCheckError", True)
End Sub


'*******************************************************************************
Private Sub lstWorkstations_LostFocus()
On Error GoTo lstWorkstations_LostFocusError
    
    Call Unload(Me)
    Exit Sub

lstWorkstations_LostFocusError:
    Call funErrorManager("frmWksnSelectionPopup.lstWorkstations_LostFocusError", True)
End Sub
