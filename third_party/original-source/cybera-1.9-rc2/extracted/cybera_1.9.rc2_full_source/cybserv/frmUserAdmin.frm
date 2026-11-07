VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmUserAdmin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "_"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDelete 
      Caption         =   "_"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   900
      Width           =   1335
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "_"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "_"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   60
      Width           =   1335
   End
   Begin VB.Frame fraLocked 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   60
      TabIndex        =   8
      Top             =   3360
      Width           =   4095
      Begin VB.CheckBox chkLocked 
         Height          =   315
         Left            =   0
         Picture         =   "frmUserAdmin.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   375
      End
      Begin VB.Frame fraAdminPassword 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   420
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   3615
         Begin VB.CommandButton cmdUnlockOptions 
            Height          =   315
            Left            =   3240
            Picture         =   "frmUserAdmin.frx":014A
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   0
            Width           =   375
         End
         Begin VB.TextBox txtLockedPassword 
            BorderStyle     =   0  'None
            Height          =   195
            IMEMode         =   3  'DISABLE
            Left            =   1320
            PasswordChar    =   "*"
            TabIndex        =   5
            Top             =   60
            Width           =   1815
         End
         Begin VB.Label lblLockedPassword 
            Alignment       =   1  'Right Justify
            Caption         =   "_"
            Height          =   195
            Left            =   0
            TabIndex        =   10
            Top             =   60
            Width           =   1215
         End
         Begin VB.Shape shp 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000010&
            Height          =   315
            Index           =   29
            Left            =   1260
            Top             =   0
            Width           =   1935
         End
      End
      Begin VB.Label lblUnlocked 
         Caption         =   "_"
         Height          =   195
         Left            =   420
         TabIndex        =   12
         Top             =   60
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label lblLocked 
         Caption         =   "Les recharges sont verouillés"
         Height          =   195
         Left            =   420
         TabIndex        =   11
         Top             =   60
         Width           =   3495
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "_"
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Top             =   3300
      Width           =   1335
   End
   Begin MSComctlLib.ImageList imlLocked 
      Left            =   4740
      Top             =   1320
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
            Picture         =   "frmUserAdmin.frx":0294
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAdmin.frx":03EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fxgUsers 
      Height          =   1965
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   3466
      _Version        =   393216
      Rows            =   10
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   -2147483628
      BackColorBkg    =   -2147483643
      GridColor       =   12632256
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      GridLinesUnpopulated=   1
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Shape shp 
      BorderColor     =   &H80000010&
      Height          =   3195
      Index           =   0
      Left            =   60
      Top             =   60
      Width           =   4035
   End
End
Attribute VB_Name = "frmUserAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ USER ADMINISTRATION ________________________|
'|  Name:          frmUserAdmin.frm                                            |
'|  Type:          VB6 Form                                                    |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
'|  Description:   Cybera user management form.                                |
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

' $Id: frmUserAdmin.frm 41 2006-02-21 18:14:29Z sam $

Option Explicit
Dim mlngSelectedUserIndex As Long


'*******************************************************************************
Private Sub chkLocked_Click()
On Error Resume Next

    fraAdminPassword.Visible = chkLocked.value = vbChecked
    lblUnlocked.Visible = chkLocked.value = vbChecked
    lblLocked.Visible = chkLocked.value = vbUnchecked
    cmdUnlockOptions.Default = chkLocked.value = vbChecked
    cmdClose.Default = chkLocked.value = vbUnchecked
    
    chkLocked.ToolTipText = IIf(chkLocked.value = vbChecked, funLoadResString(10560), funLoadResString(10559))
    If chkLocked.value = vbChecked Then
        txtLockedPassword.Text = ""
        Call funSetFocus(txtLockedPassword)
        If guServOpts.uConfig.strPasswd = "" Then Call cmdUnlockOptions_Click
    Else
        chkLocked.Picture = imlLocked.ListImages(1).Picture
        cmdAdd.enabled = False
        cmdDelete.enabled = False
        cmdModify.enabled = False
        Call funSetFocus(cmdClose)
    End If
End Sub


'*******************************************************************************
Private Sub cmdAdd_Click()
On Error GoTo cmdAdd_ClickError

    Call Load(frmUserAdminConfig)
    frmUserAdminConfig.cmdAdd.Visible = True
    Call frmUserAdminConfig.Show(vbModal, Me)
    Call subInitLocalUsers
    fxgUsers.Row = mlngSelectedUserIndex
    Call fxgUsers_RowColChange
    Exit Sub

cmdAdd_ClickError:
    Call funErrorManager("frmUserAdmin.cmdAdd_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdClose_Click()

    Call Unload(Me)
End Sub


'*******************************************************************************
Private Sub cmdDelete_Click()
Dim uUser As New clsUser
Dim uMsg As New clsStatusBarMessage
On Error GoTo cmdDelete_ClickError
    
    If mlngSelectedUserIndex = 0 Then
        Call funMsg(enuMsg.NoElementSelected)
    Else
        If funMsg(AskCanDeleteUser) = vbYes Then
            uUser.login = fxgUsers.TextMatrix(mlngSelectedUserIndex, 0)
            Call subDBDeleteUser(uUser)
            Call subInitLocalUsers
            mlngSelectedUserIndex = IIf(mlngSelectedUserIndex >= fxgUsers.Rows, fxgUsers.Rows - 1, mlngSelectedUserIndex)
            fxgUsers.Row = mlngSelectedUserIndex
            Call fxgUsers_RowColChange
            If grstUsers.EOF And grstUsers.BOF Or grstUsers.RecordCount = 0 Then
                uMsg.ImageIndex = enuStatus.cliCritical
                uMsg.Text = funLoadResString(10773)
                Call funAddStatusBarMsg(uMsg, "NO_USERS")
            End If
        
        End If
    End If
    Exit Sub

cmdDelete_ClickError:
    Call funErrorManager("frmUserAdmin.cmdDelete_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdModify_Click()
On Error GoTo cmdModify_ClickError

    If mlngSelectedUserIndex = 0 Then
        Call funMsg(enuMsg.NoElementSelected)
    Else
        Call Load(frmUserAdminConfig)
        frmUserAdminConfig.cmdModify.Visible = True
            
        grstUsers.Index = "PrimaryKey"
        Call grstUsers.Seek("=", fxgUsers.TextMatrix(mlngSelectedUserIndex, 0))
        If Not grstUsers.NoMatch Then
            frmUserAdminConfig.txtLogin.Text = grstUsers!login
            frmUserAdminConfig.txtFullName.Text = grstUsers!full_name
            frmUserAdminConfig.txtNewPassword.Text = funDoXOR(grstUsers!Passwd, funLoadResString(10006))
            frmUserAdminConfig.txtCheckNewPassword.Text = funDoXOR(grstUsers!Passwd, funLoadResString(10006))
            frmUserAdminConfig.chkAccessJournals.value = IIf(grstUsers!access_journals, vbChecked, vbUnchecked)
            frmUserAdminConfig.chkAccessUnblock.value = IIf(grstUsers!access_unblock, vbChecked, vbUnchecked)
            frmUserAdminConfig.chkAccessDelRecharge.value = IIf(grstUsers!access_del_recharge, vbChecked, vbUnchecked)
            frmUserAdminConfig.chkAccessMassCardCreation.value = IIf(grstUsers!access_mass_card_creation, vbChecked, vbUnchecked)
            frmUserAdminConfig.chkAccessCardCreation.value = IIf(grstUsers!access_card_creation, vbChecked, vbUnchecked)
            frmUserAdminConfig.chkAccessClientCreation.value = IIf(grstUsers!access_client_creation, vbChecked, vbUnchecked)
        End If
        Call frmUserAdminConfig.Show(vbModal, Me)
        Call subInitLocalUsers
        fxgUsers.Row = mlngSelectedUserIndex
        Call fxgUsers_RowColChange
    End If
    Exit Sub

cmdModify_ClickError:
    Call funErrorManager("frmUserAdmin.cmdModify_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdUnlockOptions_Click()
On Error Resume Next

    If txtLockedPassword.Text = guServOpts.uConfig.strPasswd Then
        cmdAdd.enabled = True
        cmdDelete.enabled = True
        cmdModify.enabled = True
        fraAdminPassword.Visible = False
        lblUnlocked.Visible = True
        lblLocked.Visible = False
        cmdClose.Default = True
        Call funSetFocus(cmdClose)
        chkLocked.Picture = imlLocked.ListImages(2).Picture
    Else
        chkLocked.Picture = imlLocked.ListImages(1).Picture
        Call funMsg(enuMsg.BadPasswd)
        txtLockedPassword.Text = ""
        Call funSetFocus(txtLockedPassword)
    End If
End Sub


'*******************************************************************************
Private Sub Form_Activate()
On Error Resume Next

    Call funSetFocus(fxgUsers)
End Sub


'*******************************************************************************
Private Sub Form_Load()
On Error GoTo Form_LoadError

    Call subSetDefaultFont(Me)
    Call subLocalizeInterface
    Call subInitLocalUsers
    If guServOpts.uConfig.strPasswd = "" Then chkLocked.value = vbChecked
    mlngSelectedUserIndex = 1
    Call fxgUsers_RowColChange
    Exit Sub

Form_LoadError:
    Call funErrorManager("frmUserAdmin.Form_LoadError", True)
End Sub


'*******************************************************************************
Public Sub subLocalizeInterface()
On Error GoTo subLocalizeInterfaceError

    Call subSetCharset(Me, guServOpts.uConfig.intInterfaceLangCharset)
    
    Me.Caption = funLoadResString(10000) & " - " & funLoadResString(10752)
    
    cmdClose.Caption = funLoadResString(10753)
    cmdDelete.Caption = funLoadResString(10758)
    cmdAdd.Caption = funLoadResString(10757) & "..."
    cmdModify.Caption = funLoadResString(10759) & "..."
    With fxgUsers
        .TextMatrix(0, 0) = funLoadResString(10754)
        .TextMatrix(0, 1) = funLoadResString(10755)
        .TextMatrix(0, 2) = funLoadResString(10756)
    End With
    
    lblLocked.Caption = funLoadResString(10760)
    lblLockedPassword.Caption = funLoadResString(10116)
    txtLockedPassword.ToolTipText = funLoadResString(10761)
    lblUnlocked.Caption = funLoadResString(10762)
    cmdUnlockOptions.ToolTipText = funLoadResString(10763)
    chkLocked.ToolTipText = funLoadResString(10763)
    
    Exit Sub

subLocalizeInterfaceError:
    Call funErrorManager("frmUserAdmin.subLocalizeInterfaceError", True)
End Sub


'*******************************************************************************
Private Sub subInitLocalUsers()
Dim i As Long
On Error GoTo subInitLocalUsersError

    With fxgUsers
        .top = shp(0).top + Screen.TwipsPerPixelY
        .left = shp(0).left + Screen.TwipsPerPixelX
        .width = shp(0).width - 2 * Screen.TwipsPerPixelX
        .height = shp(0).height - 2 * Screen.TwipsPerPixelY
        .Cols = 3
    
        .Row = 0
        .ColWidth(0) = 1000
        .ColAlignment(0) = flexAlignLeftTop
        
        .ColWidth(1) = 1400
        .ColAlignment(1) = flexAlignLeftTop
        
        .ColWidth(2) = 1700
        .ColAlignment(2) = flexAlignLeftTop
        
        .Font = ""
        
        Set grstUsersQry = gdbsCybAdmin.OpenRecordset("SELECT * FROM users ORDER BY login", dbOpenSnapshot)
        If Not (grstUsersQry.BOF = True And grstUsersQry.EOF = True) Then
            Call grstUsersQry.MoveLast
            Call grstUsersQry.MoveFirst
            .Rows = 2
            .Rows = grstUsersQry.RecordCount + 1
            For i = 1 To grstUsersQry.RecordCount
                .TextMatrix(i, 0) = grstUsersQry!login
                .TextMatrix(i, 1) = IIf(IsNull(grstUsersQry!full_name), "", grstUsersQry!full_name)
                .TextMatrix(i, 2) = IIf(IsNull(grstUsersQry!last_login), "", format(grstUsersQry!last_login, "Short Date") & " " & format(grstUsersQry!last_login, "Hh:Nn"))
                Call grstUsersQry.MoveNext
                DoEvents
            Next i
            .TopRow = .Rows - 1
       Else
            cmdDelete.enabled = False
            cmdModify.enabled = False
            .Rows = 2
            .TextMatrix(1, 0) = ""
            .TextMatrix(1, 1) = ""
            .TextMatrix(1, 2) = ""
        End If
    End With
    Exit Sub
    
subInitLocalUsersError:
    Call funErrorManager("frmUserAdmin.subInitLocalUsersError")
End Sub


'*******************************************************************************
Private Sub fxgUsers_Click()
On Error GoTo fxgUsers_ClickError
    
    If mlngSelectedUserIndex = 0 Then Call fxgUsers_RowColChange
    Exit Sub
    
fxgUsers_ClickError:
    Call funErrorManager("frmUserAdmin.fxgUsers_ClickError", True)
End Sub


'*******************************************************************************
Private Sub fxgUsers_DblClick()
On Error GoTo fxgUsers_DblClickError

    If cmdModify.enabled Then cmdModify.value = True
    Exit Sub
    
fxgUsers_DblClickError:
    Call funErrorManager("frmUserAdmin.fxgUsers_DblClickError", True)
End Sub


'*******************************************************************************
Private Sub fxgUsers_RowColChange()
Dim lngNewSelectedRowIndex As Long
On Error GoTo fxgUsers_RowColChangeError
    
    With fxgUsers
        lngNewSelectedRowIndex = .Row
        If lngNewSelectedRowIndex = .Rows Or lngNewSelectedRowIndex < 1 Then
            .Row = mlngSelectedUserIndex
            Exit Sub
        End If
        If mlngSelectedUserIndex > .Rows Then mlngSelectedUserIndex = .Rows
        
        .Col = 0
        If mlngSelectedUserIndex > 0 Then
            .Row = mlngSelectedUserIndex
            .ColSel = .Cols - 1
            .CellForeColor = vbWindowText
            .CellBackColor = vbWindowBackground
        End If
        
        mlngSelectedUserIndex = lngNewSelectedRowIndex
        
        .Row = lngNewSelectedRowIndex
        .ColSel = .Cols - 1
        .CellForeColor = vbHighlightText
        .CellBackColor = vbHighlight
        If .TopRow > lngNewSelectedRowIndex Then
            .TopRow = lngNewSelectedRowIndex
        ElseIf .TopRow + Int(.height / .RowHeight(0) - 2) - 1 < lngNewSelectedRowIndex Then
            .TopRow = lngNewSelectedRowIndex - Int(.height / .RowHeight(0) - 2) + 1
        End If
    End With
    Exit Sub
    
fxgUsers_RowColChangeError:
    Call funErrorManager("frmUserAdmin.fxgUsers_RowColChangeError", True)
End Sub
