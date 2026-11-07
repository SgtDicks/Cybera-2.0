VERSION 5.00
Begin VB.Form frmUserAdminConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "_"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkAccessCardCreation 
      Caption         =   "_"
      Height          =   195
      Left            =   60
      TabIndex        =   8
      Top             =   2520
      Value           =   1  'Checked
      Width           =   4695
   End
   Begin VB.CheckBox chkAccessClientCreation 
      Caption         =   "_"
      Height          =   195
      Left            =   60
      TabIndex        =   9
      Top             =   2760
      Value           =   1  'Checked
      Width           =   4695
   End
   Begin VB.CheckBox chkAccessMassCardCreation 
      Caption         =   "_"
      Height          =   195
      Left            =   60
      TabIndex        =   7
      Top             =   2280
      Value           =   1  'Checked
      Width           =   4695
   End
   Begin VB.CheckBox chkAccessDelRecharge 
      Caption         =   "_"
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   2040
      Value           =   1  'Checked
      Width           =   4695
   End
   Begin VB.TextBox txtLoginToolTip 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   525
      Left            =   2220
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   300
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CheckBox chkAccessUnblock 
      Caption         =   "_"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   1800
      Value           =   1  'Checked
      Width           =   4695
   End
   Begin VB.CheckBox chkAccessJournals 
      Caption         =   "_"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   1560
      Value           =   1  'Checked
      Width           =   4695
   End
   Begin VB.TextBox txtFullName 
      BorderStyle     =   0  'None
      Height          =   195
      IMEMode         =   3  'DISABLE
      Left            =   2220
      MaxLength       =   50
      TabIndex        =   1
      Top             =   480
      Width           =   2475
   End
   Begin VB.TextBox txtLogin 
      BorderStyle     =   0  'None
      Height          =   195
      IMEMode         =   3  'DISABLE
      Left            =   2220
      MaxLength       =   255
      TabIndex        =   0
      Top             =   120
      Width           =   2475
   End
   Begin VB.TextBox txtNewPassword 
      BorderStyle     =   0  'None
      Height          =   195
      IMEMode         =   3  'DISABLE
      Left            =   2220
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   840
      Width           =   2475
   End
   Begin VB.TextBox txtCheckNewPassword 
      BorderStyle     =   0  'None
      Height          =   195
      IMEMode         =   3  'DISABLE
      Left            =   2220
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1200
      Width           =   2475
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "_"
      Height          =   375
      Left            =   3420
      TabIndex        =   11
      Top             =   3060
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "_"
      Enabled         =   0   'False
      Height          =   375
      Left            =   60
      TabIndex        =   10
      Top             =   3060
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "_"
      Height          =   375
      Left            =   60
      TabIndex        =   12
      Top             =   3060
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
      Height          =   195
      Index           =   3
      Left            =   60
      TabIndex        =   16
      Top             =   480
      Width           =   2055
   End
   Begin VB.Shape shp 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   315
      Index           =   3
      Left            =   2160
      Top             =   420
      Width           =   2595
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   15
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   14
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
      Height          =   195
      Index           =   2
      Left            =   60
      TabIndex        =   13
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Shape shp 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   315
      Index           =   0
      Left            =   2160
      Top             =   60
      Width           =   2595
   End
   Begin VB.Shape shp 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   315
      Index           =   1
      Left            =   2160
      Top             =   780
      Width           =   2595
   End
   Begin VB.Shape shp 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   315
      Index           =   2
      Left            =   2160
      Top             =   1140
      Width           =   2595
   End
End
Attribute VB_Name = "frmUserAdminConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ USER CONFIGURATION _________________________|
'|  Name:          frmUserAdminConfig.frm                                      |
'|  Type:          VB6 Form                                                    |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
'|  Description:   Cybera user add, modification form.                         |
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

' $Id: frmUserAdminConfig.frm 15 2006-01-01 16:12:35Z sam $

Option Explicit


'*******************************************************************************
Private Sub cmdAdd_Click()
Dim uNewUser As New clsUser
On Error GoTo cmdAdd_ClickError

    If txtNewPassword.Text = txtCheckNewPassword.Text Then
        If LCase(txtLogin.Text) = ADMIN_LOGON Then
            Call funMsg(AddUserError)
            Call funSetFocus(txtLogin)
            Exit Sub
        End If
        uNewUser.login = LCase(txtLogin.Text)
        uNewUser.FullName = txtFullName.Text
        uNewUser.AccessJournals = chkAccessJournals.value = vbChecked
        uNewUser.AccessUnblock = chkAccessUnblock.value = vbChecked
        uNewUser.AccessDelRecharge = chkAccessDelRecharge.value = vbChecked
        uNewUser.AccessMassCardCreation = chkAccessMassCardCreation.value = vbChecked
        uNewUser.AccessCardCreation = chkAccessCardCreation.value = vbChecked
        uNewUser.AccessClientCreation = chkAccessClientCreation.value = vbChecked
        uNewUser.Passwd = txtNewPassword.Text
        If funDBAddUser(uNewUser) Then
            Call Unload(Me)
            Call subRemoveStatusBarMsg("NO_USERS")
        Else
            Call funMsg(AddUserError)
            Call funSetFocus(txtLogin)
        End If
    Else
        Call funMsg(enuMsg.BadNewPasswd)
        txtNewPassword.Text = ""
        txtCheckNewPassword.Text = ""
        Call funSetFocus(txtNewPassword)
    End If
    Exit Sub
    
cmdAdd_ClickError:
    Call funErrorManager("frmUserAdminConfig.cmdAdd_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdCancel_Click()

    Call Unload(Me)
End Sub


'*******************************************************************************
Private Sub cmdModify_Click()
Dim uUser As New clsUser
On Error GoTo cmdModify_ClickError

    If txtNewPassword.Text = txtCheckNewPassword.Text Then
        uUser.login = LCase(txtLogin.Text)
        uUser.FullName = txtFullName.Text
        uUser.AccessJournals = chkAccessJournals = vbChecked
        uUser.AccessUnblock = chkAccessUnblock = vbChecked
        uUser.AccessDelRecharge = chkAccessDelRecharge = vbChecked
        uUser.AccessMassCardCreation = chkAccessMassCardCreation = vbChecked
        uUser.AccessCardCreation = chkAccessCardCreation.value = vbChecked
        uUser.AccessClientCreation = chkAccessClientCreation.value = vbChecked
        uUser.Passwd = txtNewPassword.Text
        Call subDBUpdateUser(uUser)
        Call Unload(Me)
    Else
        Call funMsg(enuMsg.BadNewPasswd)
        txtNewPassword.Text = ""
        txtCheckNewPassword.Text = ""
        Call funSetFocus(txtNewPassword)
    End If
    Exit Sub
    
cmdModify_ClickError:
    Call funErrorManager("frmUserAdminConfig.cmdModify_ClickError", True)
End Sub

'*******************************************************************************
Private Sub Form_Activate()
On Error Resume Next

    cmdAdd.Default = cmdAdd.Visible
    cmdModify.Default = cmdModify.Visible
    txtLogin.enabled = cmdAdd.Visible
    Me.Caption = IIf(cmdAdd.Visible, funLoadResString(10764), funLoadResString(10770))
    Call funSetFocus(txtLogin)
End Sub


'*******************************************************************************
Private Sub Form_Load()
On Error GoTo Form_LoadError

    Call subSetDefaultFont(Me)
    Call subLocalizeInterface
    Exit Sub

Form_LoadError:
    Call funErrorManager("frmUserAdminConfig.Form_LoadError", True)
End Sub


'*******************************************************************************
Public Sub subLocalizeInterface()
On Error GoTo subLocalizeInterfaceError
    
    Call subSetCharset(Me, guServOpts.uConfig.intInterfaceLangCharset)
    
    cmdCancel.Caption = funLoadResString(10765)
    cmdAdd.Caption = funLoadResString(10757)
    cmdModify.Caption = funLoadResString(10759)
    
    lbl(0).Caption = funLoadResString(10766)
    lbl(3).Caption = funLoadResString(10767)
    lbl(1).Caption = funLoadResString(10127)
    lbl(2).Caption = funLoadResString(10128)
    
    chkAccessJournals.Caption = funLoadResString(10768)
    chkAccessUnblock.Caption = funLoadResString(10769)
    chkAccessDelRecharge.Caption = funLoadResString(10916)
    chkAccessMassCardCreation.Caption = funLoadResString(10952)
    chkAccessCardCreation.Caption = funLoadResString(11010)
    chkAccessClientCreation.Caption = funLoadResString(11011)

    txtLoginToolTip.Text = funLoadResString(10880, CLIENT_LOGIN_FORBIDDEN_CHARS)
    Exit Sub

subLocalizeInterfaceError:
    Call funErrorManager("frmUserAdminConfig.subLocalizeInterfaceError", True)
End Sub


'*******************************************************************************
Private Sub txtLogin_Change()
On Error Resume Next

    txtLoginToolTip.Visible = Not funStripChars(txtLogin, CLIENT_LOGIN_FORBIDDEN_CHARS)
    cmdAdd.enabled = Len(txtLogin.Text) > 0
    cmdModify.enabled = Len(txtLogin.Text) > 0
End Sub


'*******************************************************************************
Private Sub txtLogin_GotFocus()
On Error Resume Next
    
    txtLogin.SelStart = 0
    txtLogin.SelLength = Len(txtLogin.Text)
End Sub


'*******************************************************************************
Private Sub txtFullName_GotFocus()
On Error Resume Next
    
    txtFullName.SelStart = 0
    txtFullName.SelLength = Len(txtFullName.Text)
End Sub


'*******************************************************************************
Private Sub txtLogin_LostFocus()
On Error Resume Next

    txtLoginToolTip.Visible = False
End Sub


'*******************************************************************************
Private Sub txtNewPassword_GotFocus()
On Error Resume Next
    
    txtNewPassword.SelStart = 0
    txtNewPassword.SelLength = Len(txtNewPassword.Text)
End Sub


'*******************************************************************************
Private Sub txtCheckNewPassword_GotFocus()
On Error Resume Next
    
    txtCheckNewPassword.SelStart = 0
    txtCheckNewPassword.SelLength = Len(txtCheckNewPassword.Text)
End Sub

