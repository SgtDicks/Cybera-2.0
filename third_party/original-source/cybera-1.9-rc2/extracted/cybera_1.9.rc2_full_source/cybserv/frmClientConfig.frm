VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClientConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "_"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtLoginToolTip 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   525
      Left            =   2160
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   360
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.CommandButton cmdGenCode 
      Height          =   375
      Left            =   4740
      Picture         =   "frmClientConfig.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   60
      Width           =   375
   End
   Begin VB.TextBox txtComment 
      BorderStyle     =   0  'None
      Height          =   1155
      Left            =   2160
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   900
      Width           =   2895
   End
   Begin VB.TextBox txtFullName 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   2
      Top             =   540
      Width           =   2895
   End
   Begin VB.TextBox txtLogin 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      MaxLength       =   255
      TabIndex        =   0
      Top             =   120
      Width           =   2475
   End
   Begin VB.TextBox txtCheckNewPassword 
      BorderStyle     =   0  'None
      Height          =   195
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2580
      Width           =   2895
   End
   Begin VB.TextBox txtNewPassword 
      BorderStyle     =   0  'None
      Height          =   195
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2220
      Width           =   2895
   End
   Begin VB.Frame fraControls 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   60
      TabIndex        =   8
      Top             =   2940
      Width           =   5055
      Begin MSComctlLib.ImageList imlRecharge 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClientConfig.frx":014A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClientConfig.frx":02A4
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "_"
         Default         =   -1  'True
         Height          =   615
         Left            =   0
         Picture         =   "frmClientConfig.frx":03FE
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   3495
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "_"
         Height          =   615
         Left            =   3540
         Picture         =   "frmClientConfig.frx":0548
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   1515
      End
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
      Height          =   195
      Index           =   4
      Left            =   60
      TabIndex        =   13
      Top             =   900
      Width           =   1995
   End
   Begin VB.Shape shp 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   1275
      Index           =   1
      Left            =   2100
      Top             =   840
      Width           =   3015
   End
   Begin VB.Shape shp 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   315
      Index           =   4
      Left            =   2100
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Shape shp 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   315
      Index           =   2
      Left            =   2100
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
      Height          =   195
      Index           =   2
      Left            =   60
      TabIndex        =   12
      Top             =   2580
      Width           =   1995
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   11
      Top             =   2220
      Width           =   1995
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   10
      Top             =   120
      Width           =   1995
   End
   Begin VB.Shape shp 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   315
      Index           =   0
      Left            =   2100
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
      Height          =   195
      Index           =   3
      Left            =   60
      TabIndex        =   9
      Top             =   540
      Width           =   1995
   End
   Begin VB.Shape shpCrdComment 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   375
      Left            =   2100
      Top             =   60
      Width           =   2595
   End
End
Attribute VB_Name = "frmClientConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ CLIENT CONFIGURATION _______________________|
'|  Name:          frmUserAdminConfig.frm                                      |
'|  Type:          VB6 Form                                                    |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
'|  Description:   Cybera client add, modification form.                       |
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

' $Id: frmClientConfig.frm 22 2006-01-13 18:04:34Z sam $

Option Explicit
Dim mblnModify As Boolean


'*******************************************************************************
Private Sub cmdCancel_Click()

    Call Unload(Me)
End Sub


'*******************************************************************************
Private Sub cmdGenCode_Click()
On Error GoTo cmdGenCode_ClickError

    txtLogin.Text = funNewCardNumber
    Call funSetFocus(txtLogin)
    Exit Sub

cmdGenCode_ClickError:
    Call funErrorManager("frmClientConfig.cmdGenCode_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdOK_Click()
Dim blnFound As Boolean
Dim i As Long
Dim intWorkstationIndex As Integer
On Error GoTo cmdOK_ClickError

    If mblnModify Then
        grstClients.Index = "PrimaryKey"
        Call grstClients.Seek("=", guSelectedClient.login)
        If grstClients.NoMatch Then
            Exit Sub
        End If
        Call grstClients.Edit
        guSelectedClient.Comment = txtComment.Text
        guSelectedClient.FullName = txtFullName.Text
        guSelectedClient.Password = txtNewPassword.Text
        grstClients!Comment = guSelectedClient.Comment
        grstClients!Passwd = guSelectedClient.Password
        grstClients!full_name = guSelectedClient.FullName
        Call grstClients.Update
        intWorkstationIndex = funClientInUse(guSelectedClient)
        If intWorkstationIndex <> 0 Then
            guWorkstations(intWorkstationIndex).CurrentClient.Comment = guSelectedClient.Comment
            guWorkstations(intWorkstationIndex).CurrentClient.FullName = guSelectedClient.FullName
            guWorkstations(intWorkstationIndex).CurrentClient.Password = guSelectedClient.Password
        End If
        Call subFillClients
        Call subUpdateUISelectedClient
        Call Unload(Me)
    Else
        If guCards.Count <> 0 Then
            For i = 1 To guCards.Count
                blnFound = guCards(i).ID = txtLogin.Text
                If blnFound Then Exit For
            Next i
        End If
        If guClients.Count <> 0 And Not blnFound Then
            For i = 1 To guClients.Count
                blnFound = guClients(i).login = txtLogin.Text
                If blnFound Then Exit For
            Next i
        End If
        If blnFound Then
            Call funMsg(AddClientError, txtLogin.Text)
            Call funSetFocus(txtLogin)
        Else
            Call grstClients.AddNew
            grstClients!login = LCase(txtLogin.Text)
            grstClients!Passwd = txtNewPassword.Text
            grstClients!full_name = txtFullName.Text
            grstClients!Comment = txtComment.Text
            grstClients!date_creation = Now
            grstClients!date_expiry = Now
            grstClients!time_credit = NULL_DATE
            grstClients!Active = True
            Call grstClients.Update
            Call subDBAddEvent(enuEvents.ClientNew, 0, 0, txtLogin.Text, 0, funLoadResString(10829), All)
            DoEvents
            Call subFillClients
            Call subFindClient(CStr(txtLogin.Text), False)
            Set guSelectedClient = guClients(guServOpts.lngSelectedClientIndex)
            Call Unload(Me)
            Call Load(frmRecharge)
            Call frmRecharge.subSetRechargeType(enuCustomerType.Client)
            Call frmRecharge.Show(vbModal, frmMain)
            Call subFillClients
            Call subUpdateUISelectedClient
        End If
    End If
    Exit Sub

cmdOK_ClickError:
    Call funErrorManager("frmClientConfig.cmdOK_ClickError", True)
End Sub


'*******************************************************************************
Private Sub Form_Activate()
On Error GoTo Form_ActivateError

    Call subLocalizeInterface
    cmdOK.enabled = mblnModify
    txtLogin.enabled = Not mblnModify
    cmdGenCode.enabled = Not mblnModify
    If mblnModify Then
        txtLogin.Text = guSelectedClient.login
        txtFullName.Text = guSelectedClient.FullName
        txtComment.Text = guSelectedClient.Comment
        txtNewPassword.Text = guSelectedClient.Password
        txtCheckNewPassword.Text = guSelectedClient.Password
    End If
    Exit Sub

Form_ActivateError:
    Call funErrorManager("frmClientConfig.Form_ActivateError", True)
End Sub


'*******************************************************************************
Private Sub Form_Load()
On Error GoTo Form_LoadError

    Call subSetDefaultFont(Me)
    Exit Sub

Form_LoadError:
    Call funErrorManager("frmClientConfig.Form_LoadError", True)
End Sub


'*******************************************************************************
Public Sub subLocalizeInterface()
On Error GoTo subLocalizeInterfaceError

    Call subSetCharset(Me, guServOpts.uConfig.intInterfaceLangCharset)
    
    Me.Caption = funLoadResString(10000) & " - " & IIf(mblnModify, funLoadResString(10826), funLoadResString(10825))
    cmdCancel.Caption = funLoadResString(10765)
    cmdOK.Caption = IIf(mblnModify, funLoadResString(10759), funLoadResString(10757))
    cmdOK.Picture = imlRecharge.ListImages(IIf(mblnModify, 2, 1)).Picture
    
    lbl(0).Caption = funLoadResString(10851)
    lbl(3).Caption = funLoadResString(10767)
    lbl(1).Caption = funLoadResString(10127)
    lbl(2).Caption = funLoadResString(10128)
    lbl(4).Caption = funLoadResString(10802)
    
    txtLoginToolTip.Text = funLoadResString(10880, CLIENT_LOGIN_FORBIDDEN_CHARS)
    Exit Sub

subLocalizeInterfaceError:
    Call funErrorManager("frmClientConfig.subLocalizeInterfaceError", True)
End Sub


'*******************************************************************************
Public Sub subSetModify(blnModify As Boolean)
On Error GoTo subSetModifyError
    
    mblnModify = blnModify
    Exit Sub
    
subSetModifyError:
    Call funErrorManager("frmClientConfig.subSetModifyError")
End Sub

Private Sub Text1_Change()

End Sub


'*******************************************************************************
Private Sub txtCheckNewPassword_Change()
On Error Resume Next

    cmdOK.enabled = txtCheckNewPassword.Text = txtNewPassword.Text
End Sub


'*******************************************************************************
Private Sub txtCheckNewPassword_GotFocus()
On Error Resume Next

    txtCheckNewPassword.SelStart = 0
    txtCheckNewPassword.SelLength = Len(txtCheckNewPassword.Text)
End Sub

'*******************************************************************************
Private Sub txtComment_GotFocus()
On Error Resume Next

    txtComment.SelStart = 0
    txtComment.SelLength = Len(txtComment.Text)
    cmdOK.Default = False
End Sub


'*******************************************************************************
Private Sub txtComment_LostFocus()
On Error Resume Next

    cmdOK.Default = True
End Sub


'*******************************************************************************
Private Sub txtFullName_GotFocus()
On Error Resume Next

    txtFullName.SelStart = 0
    txtFullName.SelLength = Len(txtFullName.Text)
End Sub

'*******************************************************************************
Private Sub txtLogin_Change()
On Error Resume Next

    txtLoginToolTip.Visible = Not funStripChars(txtLogin, CLIENT_LOGIN_FORBIDDEN_CHARS)
    cmdOK.enabled = Len(txtLogin.Text) > 0
End Sub
    

'*******************************************************************************
Private Sub txtLogin_GotFocus()
On Error Resume Next

    txtLogin.SelStart = 0
    txtLogin.SelLength = Len(txtLogin.Text)
End Sub


'*******************************************************************************
Private Sub txtLogin_LostFocus()
Dim blnFound As Boolean
Dim i As Long
On Error GoTo txtLogin_LostFocusError

    If mblnModify Then Exit Sub
    If guCards.Count <> 0 Then
        For i = 1 To guCards.Count
            blnFound = guCards(i).ID = txtLogin.Text
            If blnFound Then Exit For
        Next i
    End If
    If guClients.Count <> 0 And Not blnFound Then
        For i = 1 To guClients.Count
            blnFound = guClients(i).login = txtLogin.Text
            If blnFound Then Exit For
        Next i
    End If
    If blnFound Then
        Call funMsg(AddClientError, txtLogin.Text)
        Call funSetFocus(txtLogin)
    End If
    txtLoginToolTip.Visible = False
    Exit Sub
    
txtLogin_LostFocusError:
    Call funErrorManager("frmClientConfig.txtLogin_LostFocusError", True)
End Sub


'*******************************************************************************
Private Sub txtNewPassword_Change()
On Error Resume Next

    cmdOK.enabled = txtCheckNewPassword.Text = txtNewPassword.Text
End Sub


'*******************************************************************************
Private Sub txtNewPassword_GotFocus()
On Error Resume Next

    txtNewPassword.SelStart = 0
    txtNewPassword.SelLength = Len(txtNewPassword.Text)
End Sub
