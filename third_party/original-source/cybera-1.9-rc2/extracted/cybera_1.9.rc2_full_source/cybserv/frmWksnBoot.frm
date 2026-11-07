VERSION 5.00
Begin VB.Form frmWksnBoot 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "_"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3195
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   3195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optWho 
      Caption         =   "_"
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   0
      Top             =   360
      Width           =   2835
   End
   Begin VB.TextBox txtPassword 
      BorderStyle     =   0  'None
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   4620
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.OptionButton optWho 
      Caption         =   "_"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   960
      Width           =   2835
   End
   Begin VB.OptionButton optWho 
      Caption         =   "_"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   660
      Value           =   -1  'True
      Width           =   2895
   End
   Begin VB.ListBox lstWksns 
      Enabled         =   0   'False
      Height          =   2175
      IntegralHeight  =   0   'False
      ItemData        =   "frmWksnBoot.frx":0000
      Left            =   180
      List            =   "frmWksnBoot.frx":0007
      MultiSelect     =   1  'Simple
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   2835
   End
   Begin VB.CommandButton cmdBoot 
      Caption         =   "_"
      Default         =   -1  'True
      Height          =   375
      Left            =   60
      TabIndex        =   5
      Top             =   3600
      Width           =   1995
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "_"
      Height          =   375
      Left            =   2100
      TabIndex        =   6
      Top             =   3600
      Width           =   1035
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   9
      Top             =   4620
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Shape shp 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   315
      Index           =   0
      Left            =   1260
      Top             =   4560
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label lblHelp 
      Caption         =   "_"
      Height          =   435
      Left            =   60
      TabIndex        =   8
      Top             =   4080
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Shape shp 
      BorderColor     =   &H80000010&
      Height          =   3435
      Index           =   2
      Left            =   60
      Top             =   60
      Width           =   3075
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   3075
   End
End
Attribute VB_Name = "frmWksnBoot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ WORKSTATION BOOT ___________________________|
'|  Name:          frmWksnBoot.frm                                             |
'|  Type:          VB6 Form                                                    |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
'|  Description:   Wksn. boot options                                          |
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

' $Id: frmWksnBoot.frm 15 2006-01-01 16:12:35Z sam $

Option Explicit
Private mintWho As Integer


Private Sub cmdBoot_Click()
Dim uWOL As clsMagicPacket
Dim uMac As clsMacAddress
Dim i As Integer
On Error GoTo cmdBoot_ClickError

'    If txtPassword.Text = guServOpts.uConnectedUser.Passwd Then
    If True Then
        Set uWOL = New clsMagicPacket
        With uWOL
            .WinsockControl = frmMain.wskWakeOnLan
            Select Case mintWho
            Case 0
                For Each uMac In guMacAddresses
                    If "" <> uWOL.ValidateNICAddress(uMac.MacAddress) Then
                        .NICAddress = uMac.MacAddress
                        Call .WakeUp
                    End If
                Next uMac
            Case 1
                If guMacAddresses.Count > 0 Then
                    For i = 0 To lstWksns.ListCount - 1
                        If lstWksns.Selected(i) And "" <> uWOL.ValidateNICAddress(guMacAddresses(lstWksns.ItemData(i)).MacAddress) Then
                            .NICAddress = guMacAddresses(lstWksns.ItemData(i)).MacAddress
                            Call .WakeUp
                        End If
                    Next i
                End If
            Case 2
                If "" <> uWOL.ValidateNICAddress(guSelectedWorkstation.ComputerMacAddress) Then
                    .NICAddress = guSelectedWorkstation.ComputerMacAddress
                    Call .WakeUp
                End If
            End Select
        End With
        Set uWOL = Nothing
        Call Unload(Me)
    Else
        Call funMsg(enuMsg.BadPasswd)
        txtPassword.Text = ""
        Call funSetFocus(txtPassword)
    End If
    Exit Sub

cmdBoot_ClickError:
    Call funErrorManager("frmWksnBoot.cmdBoot_ClickError", True)
End Sub

'*******************************************************************************
Private Sub cmdCancel_Click()
On Error GoTo cmdCancel_ClickError

    Call Unload(Me)
    Exit Sub

cmdCancel_ClickError:
    Call funErrorManager("frmWksnBoot.cmdCancel_ClickError", True)
End Sub


'*******************************************************************************
Private Sub Form_Activate()
On Error Resume Next

    Call funSetFocus(txtPassword)
End Sub


'*******************************************************************************
Private Sub Form_Load()
Dim i As Integer
Dim uWOL As New clsMagicPacket
On Error Resume Next
    
    Call subSetDefaultFont(Me)
    Call subLocalizeInterface
    Call lstWksns.Clear
    optWho(2).enabled = funAnyActiveWorkstation And guSelectedWorkstation.ComputerMacAddress <> "" And guSelectedWorkstation.State = enuWksnState.NotPresent
    optWho(2).value = funAnyActiveWorkstation And guSelectedWorkstation.ComputerMacAddress <> "" And guSelectedWorkstation.State = enuWksnState.NotPresent
    If guMacAddresses.Count > 0 Then
        For i = 1 To guMacAddresses.Count
            If funFindWorkstationByMac(guMacAddresses(i)).Key = "" And "" <> uWOL.ValidateNICAddress(guMacAddresses(i).MacAddress) Then
                Call lstWksns.AddItem(guMacAddresses(i).WorkstationName)
                lstWksns.ItemData(lstWksns.NewIndex) = i
                
            End If
        Next i
    End If
    optWho(1).enabled = lstWksns.ListCount > 0
End Sub


'*******************************************************************************
Public Sub subLocalizeInterface()
On Error GoTo subLocalizeInterfaceError

    Call subSetCharset(Me, guServOpts.uConfig.intInterfaceLangCharset)
    
    Me.Caption = funLoadResString(10853)
    lbl(0).Caption = funLoadResString(10855)
    lbl(1).Caption = funLoadResString(10116)
    lblHelp.Caption = funLoadResString(10858)
    lblHelp.ToolTipText = funLoadResString(10859)
    optWho(0).Caption = funLoadResString(10856)
    optWho(1).Caption = funLoadResString(10857)
    optWho(2).Caption = funLoadResString(10860)
    cmdCancel.Caption = funLoadResString(10123)
    cmdBoot.Caption = funLoadResString(10854)
    cmdBoot.ToolTipText = funLoadResString(10859)
    Exit Sub

subLocalizeInterfaceError:
    Call funErrorManager("frmWksnBoot.subLocalizeInterfaceError", True)
End Sub


'*******************************************************************************
Private Sub optWho_Click(Index As Integer)
Dim i As Integer
On Error Resume Next

    mintWho = Index
    lstWksns.enabled = Index = 1
    If Index = 1 Then
        Call funSetFocus(lstWksns)
    Else
        For i = 0 To lstWksns.ListCount - 1
            lstWksns.Selected(i) = False
        Next i
        Call funSetFocus(txtPassword)
    End If
End Sub
