VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmCardHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historique de carte"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9615
   ControlBox      =   0   'False
   Icon            =   "frmCardHistory.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   9615
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPrint 
      Height          =   375
      Left            =   7260
      Picture         =   "frmCardHistory.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5940
      Width           =   435
   End
   Begin VB.CommandButton cmdCopyExport 
      Height          =   375
      Left            =   7740
      Picture         =   "frmCardHistory.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5940
      Width           =   435
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "_"
      Default         =   -1  'True
      Height          =   375
      Left            =   8220
      TabIndex        =   8
      Top             =   5940
      Width           =   1335
   End
   Begin MSComctlLib.ImageList imlTabs 
      Left            =   6960
      Top             =   1260
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCardHistory.frx":06D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCardHistory.frx":0830
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCardHistory.frx":098A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tbsTabs 
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   1320
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   661
      TabWidthStyle   =   1
      MultiRow        =   -1  'True
      Style           =   2
      HotTracking     =   -1  'True
      ImageList       =   "imlTabs"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "_"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "_"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "_"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraHistory 
      BorderStyle     =   0  'None
      Height          =   4575
      Index           =   3
      Left            =   60
      TabIndex        =   25
      Top             =   1740
      Visible         =   0   'False
      Width           =   9495
      Begin VB.Frame fraSessionResultsURLSearch 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   60
         TabIndex        =   26
         Top             =   3720
         Width           =   4275
         Begin VB.CommandButton cmdSessionURLSearch 
            Height          =   315
            Left            =   3480
            Picture         =   "frmCardHistory.frx":0AE4
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   0
            Width           =   375
         End
         Begin VB.TextBox txtSessionURLSearch 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            HideSelection   =   0   'False
            Left            =   1020
            MaxLength       =   50
            TabIndex        =   27
            Top             =   45
            Width           =   2355
         End
         Begin VB.Shape shpLine 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000010&
            Height          =   315
            Index           =   2
            Left            =   960
            Top             =   0
            Width           =   2475
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "Recherche :"
            Height          =   195
            Index           =   33
            Left            =   0
            TabIndex        =   29
            Top             =   60
            Width           =   915
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fxgURLs 
         Height          =   2910
         Left            =   120
         TabIndex        =   7
         Top             =   60
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   5133
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
         SelectionMode   =   1
         AllowUserResizing=   1
         BorderStyle     =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Line linSessionResultsURLs 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   9480
         Y1              =   3660
         Y2              =   3660
      End
      Begin VB.Shape shp 
         BorderColor     =   &H80000010&
         Height          =   4095
         Index           =   3
         Left            =   0
         Top             =   0
         Width           =   9495
      End
   End
   Begin VB.Frame fraHistory 
      BorderStyle     =   0  'None
      Height          =   4575
      Index           =   2
      Left            =   60
      TabIndex        =   24
      Top             =   1740
      Visible         =   0   'False
      Width           =   9495
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fxgSessions 
         Height          =   2910
         Left            =   120
         TabIndex        =   6
         Top             =   60
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   5133
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
         Height          =   4095
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   9495
      End
   End
   Begin VB.Frame fraHistory 
      BorderStyle     =   0  'None
      Height          =   4575
      Index           =   1
      Left            =   60
      TabIndex        =   18
      Top             =   1740
      Width           =   9495
      Begin VB.Frame fraLocked 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   60
         TabIndex        =   19
         Top             =   3720
         Width           =   3975
         Begin VB.CheckBox chkLocked 
            Height          =   315
            Left            =   0
            Picture         =   "frmCardHistory.frx":0C2E
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   0
            Width           =   375
         End
         Begin VB.Frame fraAdminPassword 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   420
            TabIndex        =   20
            Top             =   0
            Visible         =   0   'False
            Width           =   3495
            Begin VB.CommandButton cmdUnlockOptions 
               Height          =   315
               Left            =   3120
               Picture         =   "frmCardHistory.frx":0D78
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   0
               Width           =   375
            End
            Begin VB.TextBox txtLockedPassword 
               BorderStyle     =   0  'None
               Height          =   195
               IMEMode         =   3  'DISABLE
               Left            =   1320
               PasswordChar    =   "*"
               TabIndex        =   4
               Top             =   60
               Width           =   1695
            End
            Begin VB.Label lblLockedPassword 
               Alignment       =   1  'Right Justify
               Caption         =   "_"
               Height          =   195
               Left            =   0
               TabIndex        =   21
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
               Width           =   1815
            End
         End
         Begin VB.Label lblUnlocked 
            Caption         =   "_"
            Height          =   195
            Left            =   420
            TabIndex        =   23
            Top             =   60
            Visible         =   0   'False
            Width           =   3495
         End
         Begin VB.Label lblLocked 
            Caption         =   "Les recharges sont verouillés"
            Height          =   195
            Left            =   420
            TabIndex        =   22
            Top             =   60
            Width           =   3495
         End
      End
      Begin VB.CommandButton cmdDeleteRecharge 
         Height          =   315
         Left            =   9060
         Picture         =   "frmCardHistory.frx":0EC2
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3720
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComctlLib.ImageList imlLocked 
         Left            =   4020
         Top             =   3960
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
               Picture         =   "frmCardHistory.frx":100C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCardHistory.frx":1166
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fxgRecharges 
         Height          =   1965
         Left            =   60
         TabIndex        =   1
         Top             =   60
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
      Begin VB.Line lin 
         BorderColor     =   &H80000010&
         Index           =   2
         X1              =   0
         X2              =   9480
         Y1              =   3660
         Y2              =   3660
      End
      Begin VB.Shape shp 
         BorderColor     =   &H80000010&
         Height          =   4095
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   9495
      End
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblCrdState 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1860
      TabIndex        =   16
      Top             =   840
      Width           =   2655
   End
   Begin VB.Shape shp 
      BorderColor     =   &H80000010&
      Height          =   1155
      Index           =   2
      Left            =   60
      Top             =   60
      Width           =   9495
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4620
      TabIndex        =   14
      Top             =   840
      Width           =   2115
   End
   Begin VB.Label lblCrdDteExpiry 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6780
      TabIndex        =   13
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   4620
      TabIndex        =   12
      Top             =   480
      Width           =   2115
   End
   Begin VB.Label lblCrdDteCreation 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6780
      TabIndex        =   11
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblCrdCredit 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1860
      TabIndex        =   9
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label lblCrdID 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   15
      Top             =   60
      Width           =   9495
   End
End
Attribute VB_Name = "frmCardHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ CARD HISTORY FORM __________________________|
'|  Name:          frmCardHistory.frm                                          |
'|  Type:          VB6 Form                                                    |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
'|  Description:   Card history and stat display from.                         |
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

' $Id: frmCardHistory.frm 39 2006-02-20 19:21:20Z sam $

Option Explicit
Dim muSelectedCard As New clsCard
Dim muSelectedClient As New clsClient
Dim mlngSelectedRechargeIndex As Long
Dim mlngRchargeIDs() As Long
Dim mintHistoryType As enuCustomerType



'*******************************************************************************
Private Sub chkLocked_Click()
Dim strPasswd As String
On Error Resume Next

    strPasswd = IIf(guServOpts.uConnectedUser.AccessDelRecharge, guServOpts.uConnectedUser.Passwd, guServOpts.uConfig.strPasswd)
    
    fraAdminPassword.Visible = chkLocked.value = vbChecked
    lblUnlocked.Visible = chkLocked.value = vbChecked
    lblLocked.Visible = chkLocked.value = vbUnchecked
    cmdUnlockOptions.Default = chkLocked.value = vbChecked
    cmdClose.Default = chkLocked.value = vbUnchecked
    
    chkLocked.ToolTipText = IIf(chkLocked.value = vbChecked, funLoadResString(10560), funLoadResString(10559))
    If chkLocked.value = vbChecked Then
        txtLockedPassword.Text = ""
        Call funSetFocus(txtLockedPassword)
        If strPasswd = "" Then cmdUnlockOptions.value = True
    Else
        chkLocked.Picture = imlLocked.ListImages(1).Picture
        cmdDeleteRecharge.Visible = False
        Call funSetFocus(cmdClose)
    End If
End Sub


'*******************************************************************************
Private Sub cmdClose_Click()
On Error GoTo cmdClose_ClickError
    
    Call Unload(Me)
    Exit Sub
    
cmdClose_ClickError:
    Call funErrorManager("frmCardHistory.cmdClose_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdCopyExport_Click()
Dim strOutput As String
Dim strHeader As String
On Error GoTo cmdCopyExport_ClickError
        
    Call Clipboard.Clear
    Select Case tbsTabs.SelectedItem.Index
    Case 1
        strOutput = funExportFlexGrigToText(fxgRecharges)
    Case 2
        strOutput = funExportFlexGrigToText(fxgSessions)
    Case 3
        strOutput = funExportFlexGrigToText(fxgURLs)
    End Select
    strHeader = LTrim(lblCrdID.Caption)
    strHeader = strHeader & vbNewLine & lbl(7).Caption & " " & lblCrdCredit.Caption
    strHeader = strHeader & vbNewLine & lbl(9).Caption & " " & lblCrdDteCreation.Caption
    strHeader = strHeader & vbNewLine & lbl(1).Caption & " " & lblCrdDteExpiry.Caption
    strHeader = strHeader & vbNewLine & vbNewLine & tbsTabs.SelectedItem.Caption
    strOutput = strHeader & vbNewLine & strOutput
    Call Clipboard.SetText(strOutput)
    Call funSetFocus(tbsTabs)
    Exit Sub
    
cmdCopyExport_ClickError:
    Call funErrorManager("frmCardHistory.cmdCopyExport_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdDeleteRecharge_Click()
Dim dteNewTimeCredit As Date
On Error GoTo cmdDeleteRecharge_ClickError
    
    If mlngSelectedRechargeIndex = 0 Then
        Call funMsg(enuMsg.RechargeNotSelected)
    Else
        Select Case mintHistoryType
        Case enuCustomerType.Card
            If funCardInUse(muSelectedCard) <> 0 Then
               Call funMsg(enuMsg.NoRechargeCardInUse)
            Else
                Call subQueryIncome(NULL_DATE, NULL_DATE, enuIncomeType.Card, muSelectedCard.ID, CStr(mlngRchargeIDs(mlngSelectedRechargeIndex)))
                If Not (grstJournIncomeQry.BOF = True And grstJournIncomeQry.EOF = True) Then
                    Call grstJournIncomeQry.MoveFirst
                    If grstJournIncomeQry!unlimited_play Then
                        dteNewTimeCredit = muSelectedCard.TimeCredit
                        If funMsg(AskCanDeleteRechargeFromCard, muSelectedCard.ID, "", funFormatTime(dteNewTimeCredit, enuTimeFormat.Hh_Nmn_Ssec_dyn)) = vbYes Then
                            Call subDBAddEvent(CardUnRecharge, 0, muSelectedCard.ID, "", 0, funLoadResString(10569) & funLoadResString(11000) & " | " & grstJournIncomeQry!Label, Verbose)
                            Call subDBDeleteIncome(mlngRchargeIDs(mlngSelectedRechargeIndex))
                            grstCards.Index = "PrimaryKey"
                            Call grstCards.Seek("=", muSelectedCard.ID)
                            If Not grstCards.NoMatch Then
                                Call grstCards.Edit
                                grstCards!unlimited_play = False
                                Call grstCards.Update
                            End If
                            
                            Call subFillCards
                            Call subUpdateUISelectedCard
                            Call subInitLocalRecharges
                            lblCrdCredit.Caption = " " & funFormatTime(muSelectedCard.TimeCredit, enuTimeFormat.Hh_Nmn_Ssec_dyn)
                        End If
                    Else
                        If IsNull(grstJournIncomeQry!Time) Then
                            Call funMsg(enuMsg.GeneralError, funLoadResString(10552))
                        ElseIf muSelectedCard.TimeCredit < grstJournIncomeQry!Time Then
                            Call funMsg(enuMsg.GeneralError, funLoadResString(10553))
                        Else
                            dteNewTimeCredit = muSelectedCard.TimeCredit - grstJournIncomeQry!Time
                            If funMsg(AskCanDeleteRechargeFromCard, muSelectedCard.ID, "", funFormatTime(dteNewTimeCredit, enuTimeFormat.Hh_Nmn_Ssec_dyn)) = vbYes Then
                                muSelectedCard.TimeCredit = dteNewTimeCredit
                                Call subDBAddEvent(CardUnRecharge, 0, muSelectedCard.ID, "", 0, funLoadResString(10569) & grstJournIncomeQry!income & " | " & grstJournIncomeQry!Label, Verbose)
                                Call subDBDeleteIncome(mlngRchargeIDs(mlngSelectedRechargeIndex))
                                grstCards.Index = "PrimaryKey"
                                Call grstCards.Seek("=", muSelectedCard.ID)
                                If Not grstCards.NoMatch Then
                                    Call grstCards.Edit
                                    grstCards!time_credit = muSelectedCard.TimeCredit
                                    Call grstCards.Update
                                End If
                                
                                Call subFillCards
                                Call subUpdateUISelectedCard
                                Call subInitLocalRecharges
                                lblCrdCredit.Caption = " " & IIf(muSelectedCard.UnlimitedPlay, funLoadResString(11000), funFormatTime(muSelectedCard.TimeCredit, enuTimeFormat.Hh_Nmn_Ssec_dyn))
                            End If
                        End If
                    End If
                End If
            End If
        Case enuCustomerType.Client
            If funClientInUse(muSelectedClient) <> 0 Then
               Call funMsg(enuMsg.ClientInUseNoModify)
            Else
                Call subQueryIncome(NULL_DATE, NULL_DATE, enuIncomeType.Client, , CStr(mlngRchargeIDs(mlngSelectedRechargeIndex)), muSelectedClient.login)
                If Not (grstJournIncomeQry.BOF = True And grstJournIncomeQry.EOF = True) Then
                    Call grstJournIncomeQry.MoveFirst
                    If grstJournIncomeQry!unlimited_play Then
                        dteNewTimeCredit = muSelectedClient.TimeCredit
                        If funMsg(AskCanDeleteRechargeFromClient, muSelectedClient.login, "", funFormatTime(dteNewTimeCredit, enuTimeFormat.Hh_Nmn_Ssec_dyn)) = vbYes Then
                            Call subDBAddEvent(ClientUnRecharge, 0, 0, muSelectedClient.login, 0, funLoadResString(10569) & funLoadResString(11000) & " | " & grstJournIncomeQry!Label, Verbose)
                            Call subDBDeleteIncome(mlngRchargeIDs(mlngSelectedRechargeIndex))
                            grstCards.Index = "PrimaryKey"
                            Call grstClients.Seek("=", muSelectedClient.login)
                            If Not grstClients.NoMatch Then
                                Call grstClients.Edit
                                grstClients!unlimited_play = False
                                Call grstClients.Update
                            End If
                            
                            Call subFillClients
                            Call subUpdateUISelectedClient
                            Call subInitLocalRecharges
                            lblCrdCredit.Caption = " " & funFormatTime(muSelectedClient.TimeCredit, enuTimeFormat.Hh_Nmn_Ssec_dyn)
                        End If
                    Else
                        If IsNull(grstJournIncomeQry!Time) Then
                            Call funMsg(enuMsg.GeneralError, funLoadResString(10552))
                        ElseIf muSelectedClient.TimeCredit < grstJournIncomeQry!Time Then
                            Call funMsg(enuMsg.GeneralError, funLoadResString(10822))
                        Else
                            dteNewTimeCredit = muSelectedClient.TimeCredit - grstJournIncomeQry!Time
                            If funMsg(AskCanDeleteRechargeFromClient, muSelectedClient.login, "", funFormatTime(dteNewTimeCredit, enuTimeFormat.Hh_Nmn_Ssec_dyn)) = vbYes Then
                                muSelectedClient.TimeCredit = dteNewTimeCredit
                                Call subDBAddEvent(ClientUnRecharge, 0, 0, muSelectedClient.login, 0, funLoadResString(10569) & grstJournIncomeQry!income & " | " & grstJournIncomeQry!Label, Verbose)
                                Call subDBDeleteIncome(mlngRchargeIDs(mlngSelectedRechargeIndex))
                                grstClients.Index = "PrimaryKey"
                                Call grstClients.Seek("=", muSelectedClient.login)
                                If Not grstClients.NoMatch Then
                                    Call grstClients.Edit
                                    grstClients!time_credit = muSelectedClient.TimeCredit
                                    Call grstClients.Update
                                End If
                                
                                Call subFillClients
                                Call subUpdateUISelectedClient
                                Call subInitLocalRecharges
                                lblCrdCredit.Caption = " " & IIf(muSelectedClient.UnlimitedPlay, funLoadResString(11000), funFormatTime(muSelectedClient.TimeCredit, enuTimeFormat.Hh_Nmn_Ssec_dyn))
                            End If
                        End If
                    End If
                End If
            End If
        End Select
    End If
    Call funSetFocus(fxgRecharges)
    Exit Sub
    
cmdDeleteRecharge_ClickError:
    Call funErrorManager("frmCardHistory.cmdDeleteRecharge_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdPrint_Click()
Dim fxgSource As MSHFlexGrid
On Error GoTo cmdPrint_ClickError

    Select Case tbsTabs.SelectedItem.Index
    Case 1
        Set fxgSource = fxgRecharges
    Case 2
        Set fxgSource = fxgSessions
    Case 3
        Set fxgSource = fxgURLs
    End Select
    Call subStatsPrint(fxgSource, lblCrdID.Caption & " - " & tbsTabs.SelectedItem.Caption)
    Exit Sub
    
cmdPrint_ClickError:
    Call funErrorManager("frmCardHistory.cmdPrint_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdSessionURLSearch_Click()
Dim i As Long
Dim strSearchString As String
Dim blnFound As Boolean
On Error GoTo cmdSessionURLSearch_ClickError
    
    strSearchString = txtSessionURLSearch.Text
    If Len(strSearchString) = 0 Then Exit Sub
    Call funSetFocus(cmdSessionURLSearch)
    With fxgURLs
        For i = IIf(.Row <> .Rows, .Row + 1, 1) To .Rows - 1
            blnFound = InStr(LCase(.TextMatrix(i, 2) & " " & .TextMatrix(i, 3)), LCase(strSearchString)) <> 0
            If blnFound Then Exit For
        Next i
        If Not blnFound And .Row <> 0 Then
            For i = 1 To .Row
                blnFound = InStr(LCase(.TextMatrix(i, 2) & " " & .TextMatrix(i, 3)), LCase(strSearchString)) <> 0
                If blnFound Then Exit For
            Next i
        End If
        If blnFound Then
            .Row = i
            Call fxgURLs_RowColChange
        Else
            Call funMsg(enuMsg.URLNotFound)
        End If
    End With
    Call funSetFocus(txtSessionURLSearch)
    Exit Sub
    
cmdSessionURLSearch_ClickError:
    Call funErrorManager("frmCardHistory.cmdSessionURLSearch_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdUnlockOptions_Click()
Dim strPasswd As String
On Error Resume Next

    strPasswd = IIf(guServOpts.uConnectedUser.AccessDelRecharge, guServOpts.uConnectedUser.Passwd, guServOpts.uConfig.strPasswd)
    If txtLockedPassword.Text = strPasswd Then
        cmdDeleteRecharge.Visible = True
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
Dim strPasswd As String
On Error GoTo Form_ActivateError

    Set muSelectedCard = guSelectedCard
    Set muSelectedClient = guSelectedClient
    mlngSelectedRechargeIndex = 0
    Select Case mintHistoryType
    Case enuCustomerType.Card
        lblCrdID.Caption = funLoadResString(10147) & muSelectedCard.ID
        lblCrdCredit.Caption = " " & IIf(muSelectedCard.UnlimitedPlay, funLoadResString(11000), funFormatTime(muSelectedCard.TimeCredit, enuTimeFormat.Hh_Nmn_Ssec_dyn))
        lblCrdDteCreation.Caption = " " & format(muSelectedCard.DateCreation, "Long Date")
        lblCrdDteExpiry.Caption = " " & format(muSelectedCard.DateExpiry, "Long Date")
        If muSelectedCard.Active Then
            lblCrdState.Caption = funLoadResString(10199)
        Else
            lblCrdState.Caption = funLoadResString(10200)
        End If
    Case enuCustomerType.Client
        lblCrdID.Caption = funLoadResString(10800) & muSelectedClient.login
        lblCrdCredit.Caption = " " & IIf(muSelectedClient.UnlimitedPlay, funLoadResString(11000), funFormatTime(muSelectedClient.TimeCredit, enuTimeFormat.Hh_Nmn_Ssec_dyn))
        lblCrdDteCreation.Caption = " " & format(muSelectedClient.DateCreation, "Long Date")
        lbl(1).enabled = guServOpts.uClients.intEnableExpiry = vbChecked
        lblCrdDteExpiry.enabled = guServOpts.uClients.intEnableExpiry = vbChecked
        lblCrdDteExpiry.Caption = IIf(guServOpts.uClients.intEnableExpiry = vbChecked, " " & format(muSelectedClient.DateExpiry, "Long Date"), funLoadResString(10865))
        If muSelectedClient.Active Then
            lblCrdState.Caption = funLoadResString(10199)
        Else
            lblCrdState.Caption = funLoadResString(10200)
        End If
    End Select
    Call subInitLocalRecharges
    Call subInitLocalSessions
    Call subInitLocalURLs
    Call subLocalizeInterface
    strPasswd = IIf(guServOpts.uConnectedUser.AccessDelRecharge, guServOpts.uConnectedUser.Passwd, guServOpts.uConfig.strPasswd)
    If strPasswd = "" Then chkLocked.value = vbChecked
    cmdPrint.enabled = guServOpts.uConfig.strDefaultPrinterName <> ""
    Exit Sub

Form_ActivateError:
    Call funErrorManager("frmCardHistory.Form_ActivateError", True)
End Sub


'*******************************************************************************
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Form_KeyUpUpError
    
    If KeyCode = vbKeyC And Shift = vbCtrlMask Then cmdCopyExport.value = True
    If KeyCode = vbKeyP And Shift = vbCtrlMask Then cmdPrint.value = True
    Exit Sub
    
Form_KeyUpUpError:
    Call funErrorManager("frmCardHistory.Form_KeyUpUpError", True)
End Sub


'*******************************************************************************
Private Sub Form_Load()
On Error GoTo Form_LoadError
    
    gblnfrmCardHistory = True
    Call subSetDefaultFont(Me)
    Exit Sub

Form_LoadError:
    Call funErrorManager("frmCardHistory.Form_LoadError", True)
End Sub


'*******************************************************************************
Public Sub subLocalizeInterface()
On Error GoTo subLocalizeInterfaceError

    Call subSetCharset(Me, guServOpts.uConfig.intInterfaceLangCharset)
    
    Me.Caption = funLoadResString(10000) & " - " & IIf(mintHistoryType = enuCustomerType.Card, funLoadResString(10532), funLoadResString(10818))
    
    tbsTabs.Tabs(1).Caption = funLoadResString(10536)
'    tbsTabs.Tabs(1).ToolTipText = funLoadResString()
    tbsTabs.Tabs(2).Caption = funLoadResString(10534)
'    tbsTabs.Tabs(2).ToolTipText = funLoadResString()
    tbsTabs.Tabs(3).Caption = funLoadResString(10697)
'    tbsTabs.Tabs(3).ToolTipText = funLoadResString()
    
    'lbl(0).Caption = funLoadResString(10534)
    lbl(1).Caption = funLoadResString(10135)
    'lbl(3).Caption = funLoadResString(10536)
    lbl(7).Caption = funLoadResString(10133)
    lbl(9).Caption = funLoadResString(10134)
    lbl(11).Caption = funLoadResString(10184)
    lbl(33).Caption = funLoadResString(10187)
    cmdClose.Caption = funLoadResString(10533)
    cmdDeleteRecharge.ToolTipText = funLoadResString(10561)
    
    cmdSessionURLSearch.ToolTipText = funLoadResString(10718)

    cmdCopyExport.ToolTipText = funLoadResString(10963)
    cmdPrint.ToolTipText = funLoadResString(10966, guServOpts.uConfig.strDefaultPrinterName)

    lblLocked.Caption = funLoadResString(10556)
    lblLockedPassword.Caption = funLoadResString(10116)
    txtLockedPassword.ToolTipText = funLoadResString(IIf(guServOpts.uConnectedUser.AccessDelRecharge, 10558, 10917))
    lblUnlocked.Caption = funLoadResString(10557)
    cmdUnlockOptions.ToolTipText = funLoadResString(10559)
    chkLocked.ToolTipText = funLoadResString(10559)

    With fxgSessions
        .TextMatrix(0, 0) = funLoadResString(10403)
        .TextMatrix(0, 1) = funLoadResString(10405)
        .TextMatrix(0, 2) = funLoadResString(10409)
        .TextMatrix(0, 3) = funLoadResString(10406)
        .TextMatrix(0, 4) = funLoadResString(10407)
        .TextMatrix(0, 5) = funLoadResString(10408)
    End With
    With fxgRecharges
        .TextMatrix(0, 0) = funLoadResString(10293)
        .TextMatrix(0, 1) = funLoadResString(10294)
        .TextMatrix(0, 2) = funLoadResString(10297)
    End With
    With fxgURLs
        .TextMatrix(0, 0) = funLoadResString(10403)
        .TextMatrix(0, 1) = funLoadResString(10405)
        .TextMatrix(0, 2) = funLoadResString(10695)
        .TextMatrix(0, 3) = funLoadResString(10696)
    End With
    Exit Sub

subLocalizeInterfaceError:
    Call funErrorManager("frmCardHistory.subLocalizeInterfaceError", True)
End Sub


'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)

    gblnfrmCardHistory = False
End Sub


'*******************************************************************************
Private Sub fxgRecharges_Click()
On Error GoTo fxgRecharges_ClickError
    
    If mlngSelectedRechargeIndex = 0 Then Call fxgRecharges_RowColChange
    Exit Sub
    
fxgRecharges_ClickError:
    Call funErrorManager("frmCardHistory.fxgRecharges_ClickError", True)
End Sub


'*******************************************************************************
Private Sub fxgRecharges_KeyUp(KeyCode As Integer, Shift As Integer)
End Sub


'*******************************************************************************
Private Sub fxgRecharges_RowColChange()
Dim lngNewSelectedRowIndex As Long
On Error GoTo fxgRecharges_RowColChangeError
    
    With fxgRecharges
        lngNewSelectedRowIndex = .Row
        If lngNewSelectedRowIndex = .Rows - 1 Or lngNewSelectedRowIndex < 1 Then
            .Row = mlngSelectedRechargeIndex
            Exit Sub
        End If
        If mlngSelectedRechargeIndex > .Rows - 1 Then mlngSelectedRechargeIndex = .Rows - 1
        
        .Col = 0
        If mlngSelectedRechargeIndex > 0 Then
            .Row = mlngSelectedRechargeIndex
            .ColSel = .Cols - 1
            .CellForeColor = vbWindowText
            .CellBackColor = vbWindowBackground
        End If
        
        mlngSelectedRechargeIndex = lngNewSelectedRowIndex
        
        .Row = lngNewSelectedRowIndex
        .ColSel = .Cols - 1
        .CellForeColor = vbHighlightText
        .CellBackColor = vbHighlight
    End With
    Exit Sub
    
fxgRecharges_RowColChangeError:
    Call funErrorManager("frmCardHistory.fxgRecharges_RowColChangeError", True)
End Sub


'*******************************************************************************
Private Sub subInitLocalRecharges()
Dim i As Long
Dim sngIncomeTotal As Single
On Error GoTo subInitRechargesError

    With fxgRecharges
        .top = shp(1).top + Screen.TwipsPerPixelY '+ lbl(0).height
        .left = shp(1).left + Screen.TwipsPerPixelX
        .width = shp(1).width - 2 * Screen.TwipsPerPixelX
        .height = lin(2).Y1 - Screen.TwipsPerPixelY  '- lbl(0).height
        .Cols = 3
        
        .Row = 0
        .ColWidth(0) = 1900
        .ColAlignment(0) = flexAlignLeftTop
        
        .ColWidth(1) = 1000
        .ColAlignment(1) = flexAlignRightTop
        
        .ColWidth(2) = 6600
        .ColAlignment(2) = flexAlignLeftTop
        
        .Font = ""
        
        If mlngSelectedRechargeIndex > 0 Then
            .Row = mlngSelectedRechargeIndex
            .ColSel = .Cols - 1
            .CellForeColor = vbWindowText
            .CellBackColor = vbWindowBackground
        End If
        Select Case mintHistoryType
        Case enuCustomerType.Client
            Call subQueryIncome(muSelectedClient.DateCreation, NULL_DATE, enuIncomeType.Client, , , muSelectedClient.login)
        Case enuCustomerType.Card
            Call subQueryIncome(muSelectedCard.DateCreation, NULL_DATE, enuIncomeType.Card, muSelectedCard.ID)
        End Select
        If Not (grstJournIncomeQry.BOF = True And grstJournIncomeQry.EOF = True) Then
            .Rows = grstJournIncomeQry.RecordCount + 2
            ReDim mlngRchargeIDs(grstJournIncomeQry.RecordCount)
            Call grstJournIncomeQry.MoveFirst
            For i = 1 To grstJournIncomeQry.RecordCount
                sngIncomeTotal = sngIncomeTotal + CSng(grstJournIncomeQry!income)
                .TextMatrix(i, 0) = grstJournIncomeQry!date_stamp
                .TextMatrix(i, 1) = funFormatCurrencyString(CSng(grstJournIncomeQry!income), guServOpts.uConfig.strMonFormat)
                .TextMatrix(i, 2) = grstJournIncomeQry!Label
                mlngRchargeIDs(i) = grstJournIncomeQry!ID
                Call grstJournIncomeQry.MoveNext
                
                DoEvents
            Next i
            i = grstJournIncomeQry.RecordCount + 1
            .Row = i
            .Col = 0
            .ColSel = .Cols - 1
            .CellBackColor = vbButtonShadow
            .TextMatrix(i, 0) = funLoadResString(10387) & " (" & grstJournIncomeQry.RecordCount & funLoadResString(10551) & ")"
            .TextMatrix(i, 1) = funFormatCurrencyString(sngIncomeTotal, guServOpts.uConfig.strMonFormat)
            .TextMatrix(i, 2) = ""
            .TopRow = .Rows - 1
        Else
            .Rows = 2
            .TextMatrix(1, 0) = ""
            .TextMatrix(1, 1) = ""
            .TextMatrix(1, 2) = ""
        End If
    End With
    mlngSelectedRechargeIndex = 0
    Exit Sub
    
subInitRechargesError:
    Call funErrorManager("frmCardHistory.subInitRechargesError")
End Sub


'*******************************************************************************
Private Sub subInitLocalSessions()
Dim i As Long
Dim dteSessionTime As Date
Dim dteTotalSessionTime As Date
Dim dteTotalCon As Date
Dim dteTotalDiscon As Date
Dim dteTotalConOffpeak As Date
On Error GoTo subInitSessionsError

    With fxgSessions
        .top = shp(0).top + Screen.TwipsPerPixelY '+ lbl(0).height
        .left = shp(0).left + Screen.TwipsPerPixelX
        .width = shp(0).width - 2 * Screen.TwipsPerPixelX
        .height = shp(0).height - 2 * Screen.TwipsPerPixelY '- lbl(0).height
        .Cols = 6
    
        .Row = 0
        .ColWidth(0) = 1900
        .ColAlignment(0) = flexAlignLeftTop
        
        .ColWidth(1) = 1500
        .ColAlignment(1) = flexAlignLeftTop
        
        .ColWidth(2) = 1450
        .ColAlignment(2) = flexAlignLeftTop
        
        .ColWidth(3) = 1450
        .ColAlignment(3) = flexAlignLeftTop
        
        .ColWidth(4) = 1450
        .ColAlignment(4) = flexAlignLeftTop
        
        .ColWidth(5) = 1800
        .ColAlignment(5) = flexAlignLeftTop
    
        .Font = ""
        
        Select Case mintHistoryType
        Case enuCustomerType.Client
            Call subQuerySession(muSelectedClient.DateCreation, NULL_DATE, , muSelectedClient.login)
        Case enuCustomerType.Card
            Call subQuerySession(muSelectedCard.DateCreation, NULL_DATE, muSelectedCard.ID)
        End Select
        If Not (grstJournSessionsQry.BOF = True And grstJournSessionsQry.EOF = True) Then
            .Rows = grstJournSessionsQry.RecordCount + 3
            Call grstJournSessionsQry.MoveFirst
            For i = 1 To grstJournSessionsQry.RecordCount
                dteSessionTime = IIf(IsNull(grstJournSessionsQry!time_connected), NULL_DATE, grstJournSessionsQry!time_connected) + IIf(IsNull(grstJournSessionsQry!time_disconnected), NULL_DATE, grstJournSessionsQry!time_disconnected) + IIf(IsNull(grstJournSessionsQry!time_connected_offpeak), NULL_DATE, grstJournSessionsQry!time_connected_offpeak)
                .TextMatrix(i, 0) = grstJournSessionsQry!date_start
                .TextMatrix(i, 1) = IIf(IsNull(grstJournSessionsQry!workstation_name), "", grstJournSessionsQry!workstation_name)
                .TextMatrix(i, 2) = funFormatTime(dteSessionTime)
                dteTotalSessionTime = dteTotalSessionTime + dteSessionTime
                If grstJournSessionsQry!time_connected <> NULL_DATE Then .TextMatrix(i, 3) = funFormatTime(grstJournSessionsQry!time_connected)
                If grstJournSessionsQry!time_connected <> NULL_DATE Then dteTotalCon = dteTotalCon + grstJournSessionsQry!time_connected
                
                If grstJournSessionsQry!time_disconnected <> NULL_DATE Then .TextMatrix(i, 4) = funFormatTime(grstJournSessionsQry!time_disconnected)
                If grstJournSessionsQry!time_disconnected <> NULL_DATE Then dteTotalDiscon = dteTotalDiscon + grstJournSessionsQry!time_disconnected
                
                If grstJournSessionsQry!time_connected_offpeak <> NULL_DATE Then .TextMatrix(i, 5) = funFormatTime(grstJournSessionsQry!time_connected_offpeak)
                If grstJournSessionsQry!time_connected_offpeak <> NULL_DATE Then dteTotalConOffpeak = dteTotalConOffpeak + grstJournSessionsQry!time_connected_offpeak
                Call grstJournSessionsQry.MoveNext
                DoEvents
            Next i
            i = grstJournSessionsQry.RecordCount + 1
            .Row = i
            .Col = 0
            .ColSel = .Cols - 1
            .CellBackColor = vbButtonShadow
            
            .TextMatrix(i, 0) = funLoadResString(10387) & " (" & grstJournSessionsQry.RecordCount & funLoadResString(10410) & ")"
            .TextMatrix(i, 1) = ""
            .TextMatrix(i, 2) = funFormatTime(dteTotalSessionTime)
            .TextMatrix(i, 3) = funFormatTime(dteTotalCon)
            .TextMatrix(i, 4) = funFormatTime(dteTotalDiscon)
            .TextMatrix(i, 5) = funFormatTime(dteTotalConOffpeak)
            
            i = i + 1
            .Row = i
            .Col = 0
            .ColSel = .Cols - 1
            .CellBackColor = vbButtonFace
            
            .TextMatrix(i, 0) = funLoadResString(10388)
            .TextMatrix(i, 2) = funFormatTime(dteTotalSessionTime / grstJournSessionsQry.RecordCount)
            .TextMatrix(i, 3) = funFormatTime(dteTotalCon / grstJournSessionsQry.RecordCount)
            .TextMatrix(i, 4) = funFormatTime(dteTotalDiscon / grstJournSessionsQry.RecordCount)
            .TextMatrix(i, 5) = funFormatTime(dteTotalConOffpeak / grstJournSessionsQry.RecordCount)
            .TopRow = .Rows - 1
        Else
            .Rows = 2
            .TextMatrix(1, 0) = ""
            .TextMatrix(1, 1) = ""
            .TextMatrix(1, 2) = ""
            .TextMatrix(1, 3) = ""
            .TextMatrix(1, 4) = ""
            .TextMatrix(1, 5) = ""
        End If
    End With
    Exit Sub
    
subInitSessionsError:
    Call funErrorManager("frmCardHistory.subInitSessionsError")
End Sub


'*******************************************************************************
Private Sub subInitLocalURLs()
Dim i As Long
On Error GoTo subInitLocalURLsError

    With fxgURLs
        .top = shp(3).top + Screen.TwipsPerPixelY
        .left = shp(3).left + Screen.TwipsPerPixelX
        .width = shp(3).width - 2 * Screen.TwipsPerPixelX
        .height = linSessionResultsURLs.Y1 - Screen.TwipsPerPixelY
        .Cols = 4
    
        .Row = 0
        .ColWidth(0) = 1900
        .ColAlignment(0) = flexAlignLeftTop
        
        .ColWidth(1) = 1200
        .ColAlignment(1) = flexAlignLeftTop
        
        .ColWidth(2) = 2000
        .ColAlignment(2) = flexAlignLeftTop
        
        .ColWidth(3) = 7000
        .ColAlignment(3) = flexAlignLeftTop
    
        .Font = ""
        
        Select Case mintHistoryType
        Case enuCustomerType.Client
            Call subQueryURLs(muSelectedClient.DateCreation, NULL_DATE, , , muSelectedClient.login)
        Case enuCustomerType.Card
            Call subQueryURLs(muSelectedCard.DateCreation, NULL_DATE, muSelectedCard.ID)
        End Select
        If Not (grstUrlsQry.BOF = True And grstUrlsQry.EOF = True) Then
            .Rows = grstUrlsQry.RecordCount + 1
            Call grstUrlsQry.MoveFirst
            For i = 1 To grstUrlsQry.RecordCount
                .TextMatrix(i, 0) = grstUrlsQry!date_stamp
                .TextMatrix(i, 1) = IIf(IsNull(grstUrlsQry!workstation_name), "", grstUrlsQry!workstation_name)
                .TextMatrix(i, 2) = IIf(IsNull(grstUrlsQry!Title), "", grstUrlsQry!Title)
                .TextMatrix(i, 3) = IIf(IsNull(grstUrlsQry!url), "", grstUrlsQry!url)
                Call grstUrlsQry.MoveNext
                DoEvents
            Next i
            .TopRow = .Rows - 1
            .ToolTipText = grstUrlsQry.RecordCount & funLoadResString(10303)
        Else
            .Rows = 2
            .TextMatrix(1, 0) = ""
            .TextMatrix(1, 1) = ""
            .TextMatrix(1, 2) = ""
            .TextMatrix(1, 3) = ""
        End If
    End With
    Exit Sub
    
subInitLocalURLsError:
    Call funErrorManager("frmCardHistory.subInitLocalURLsError")
End Sub


'*******************************************************************************
Private Sub fxgURLs_Click()
On Error GoTo fxgURLs_ClickError
    
    If glngHistorySelectedURLIndex = 0 Then Call fxgURLs_RowColChange
    Exit Sub
    
fxgURLs_ClickError:
    Call funErrorManager("frmCardHistory.fxgURLs_ClickError", True)
End Sub


'*******************************************************************************
Private Sub fxgURLs_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lngNewRow As Long
On Error GoTo fxgURLs_MouseDownError

    lngNewRow = fxgURLs.TopRow + Int(y / fxgURLs.RowHeight(0)) - 1
    If lngNewRow <> fxgURLs.Row And lngNewRow < fxgURLs.Rows Then
        fxgURLs.Row = lngNewRow
        Call fxgURLs_RowColChange
    End If
    If Button = vbRightButton And y >= fxgURLs.RowHeight(0) And lngNewRow < fxgURLs.Rows Then
        Call PopupMenu(frmMain.mnuSessionURLDropDown)
    End If
    Exit Sub
    
fxgURLs_MouseDownError:
    Call funErrorManager("frmCardHistory.fxgURLs_MouseDownError", True)
End Sub


'*******************************************************************************
Private Sub fxgURLs_RowColChange()
Dim lngNewSelectedRowIndex As Long
On Error GoTo fxgURLs_RowColChangeError
    
    With fxgURLs
        lngNewSelectedRowIndex = .Row
        If lngNewSelectedRowIndex = .Rows Or lngNewSelectedRowIndex < 1 Then
            .Row = glngHistorySelectedURLIndex
            Exit Sub
        End If
        If glngHistorySelectedURLIndex > .Rows - 1 Then glngHistorySelectedURLIndex = .Rows - 1
        
        .Col = 0
        If glngHistorySelectedURLIndex > 0 Then
            .Row = glngHistorySelectedURLIndex
            .ColSel = .Cols - 1
            .CellForeColor = vbWindowText
            .CellBackColor = vbWindowBackground
        End If
        
        glngHistorySelectedURLIndex = lngNewSelectedRowIndex
        
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
    
fxgURLs_RowColChangeError:
    Call funErrorManager("frmCardHistory.fxgURLs_RowColChangeError", True)
End Sub


'*******************************************************************************
Private Sub tbsTabs_Click()
Dim i As Integer
On Error GoTo tbsTabs_ClickError
    
    For i = 1 To fraHistory.Count
        fraHistory(i).Visible = i = tbsTabs.SelectedItem.Index
    Next i
    cmdSessionURLSearch.Default = tbsTabs.SelectedItem.Index = 3
    cmdClose.Default = tbsTabs.SelectedItem.Index = 2
    cmdUnlockOptions.Default = tbsTabs.SelectedItem.Index = 1
    Exit Sub
    
tbsTabs_ClickError:
    Call funErrorManager("frmCardHistory.tbsTabs_ClickError", True)
End Sub


'*******************************************************************************
Private Sub txtSessionURLSearch_GotFocus()
On Error GoTo txtSessionURLSearch_GotFocusError
    
    txtSessionURLSearch.SelStart = 0
    txtSessionURLSearch.SelLength = Len(txtSessionURLSearch.Text)
    Exit Sub
    
txtSessionURLSearch_GotFocusError:
    Call funErrorManager("frmCardHistory.txtSessionURLSearch_GotFocusError", True)
End Sub


'*******************************************************************************
Public Sub subSetHistoryType(intType As enuCustomerType)
On Error GoTo subSetHistoryTypeError
    
    mintHistoryType = intType
    Exit Sub
    
subSetHistoryTypeError:
    Call funErrorManager("frmCardHistory.subSetHistoryTypeError")
End Sub
