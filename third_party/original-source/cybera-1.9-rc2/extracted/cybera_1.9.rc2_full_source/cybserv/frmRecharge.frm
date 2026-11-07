VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmRecharge 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "_"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11115
   ControlBox      =   0   'False
   Icon            =   "frmRecharge.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   425
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   741
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   5040
      TabIndex        =   21
      Top             =   60
      Width           =   6015
      Begin VB.Timer tmrUpdateCard 
         Interval        =   500
         Left            =   4620
         Top             =   0
      End
      Begin VB.Shape shp 
         BorderColor     =   &H80000010&
         Height          =   1515
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   6015
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
         Index           =   10
         Left            =   60
         TabIndex        =   27
         Top             =   1140
         Width           =   2775
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
         Left            =   2880
         TabIndex        =   26
         Top             =   1140
         Width           =   3015
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
         Left            =   60
         TabIndex        =   25
         Top             =   780
         Width           =   2775
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
         Left            =   2880
         TabIndex        =   24
         Top             =   780
         Width           =   3015
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
         Left            =   60
         TabIndex        =   23
         Top             =   420
         Width           =   2775
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
         Left            =   2880
         TabIndex        =   22
         Top             =   420
         Width           =   3015
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
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   13695
      End
   End
   Begin VB.TextBox txtQuantity 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7980
      MaxLength       =   4
      TabIndex        =   1
      Text            =   "1"
      Top             =   3240
      Width           =   315
   End
   Begin MSComctlLib.ImageList imlCancel 
      Left            =   9540
      Top             =   5700
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecharge.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecharge.frx":059C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecharge.frx":06F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlRecharge 
      Left            =   5040
      Top             =   5700
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecharge.frx":0850
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecharge.frx":09AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecharge.frx":0B04
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fxgRecharges 
      Height          =   5130
      Left            =   180
      TabIndex        =   0
      Top             =   360
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   9049
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
   Begin VB.CommandButton cmdCrdRecharge 
      Caption         =   "_"
      Default         =   -1  'True
      Height          =   615
      Left            =   5040
      Picture         =   "frmRecharge.frx":0C5E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5700
      Width           =   4395
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "_"
      Height          =   615
      Left            =   9540
      Picture         =   "frmRecharge.frx":0DA8
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5700
      Width           =   1515
   End
   Begin MSComCtl2.UpDown udQuantity 
      Height          =   375
      Left            =   8370
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3180
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Value           =   1
      OrigLeft        =   1920
      OrigTop         =   300
      OrigRight       =   2160
      OrigBottom      =   555
      Max             =   99
      Min             =   1
      Enabled         =   0   'False
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
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
      Index           =   12
      Left            =   5100
      TabIndex        =   20
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Shape shp 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   375
      Index           =   45
      Left            =   7920
      Top             =   3180
      Width           =   435
   End
   Begin VB.Shape shp 
      BorderColor     =   &H80000010&
      Height          =   6255
      Index           =   3
      Left            =   60
      Top             =   60
      Width           =   4875
   End
   Begin VB.Line lin 
      BorderColor     =   &H00C0C0C0&
      X1              =   4
      X2              =   328
      Y1              =   20
      Y2              =   20
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
      Index           =   8
      Left            =   5100
      TabIndex        =   17
      Top             =   2820
      Width           =   2775
   End
   Begin VB.Label lblRechargeCaption 
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
      Left            =   7920
      TabIndex        =   16
      Top             =   2820
      Width           =   3015
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
      Index           =   6
      Left            =   5100
      TabIndex        =   15
      Top             =   5220
      Width           =   2775
   End
   Begin VB.Label lblCrdDteExpiryNew 
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
      Left            =   7920
      TabIndex        =   14
      Top             =   5220
      Width           =   3015
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
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
      Index           =   5
      Left            =   5100
      TabIndex        =   13
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Label lblRechargeCreditNew 
      BackColor       =   &H80000018&
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
      Left            =   7920
      TabIndex        =   12
      Top             =   4800
      Width           =   3015
   End
   Begin VB.Shape shp 
      BorderColor     =   &H80000010&
      Height          =   1155
      Index           =   2
      Left            =   5040
      Top             =   4440
      Width           =   6015
   End
   Begin VB.Label lblRechargePrice 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   10
      Top             =   3840
      Width           =   3015
   End
   Begin VB.Label lblRechargeValidity 
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
      Left            =   7920
      TabIndex        =   9
      Top             =   2460
      Width           =   3015
   End
   Begin VB.Label lblRechargeCredit 
      BackColor       =   &H80000018&
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
      Left            =   7920
      TabIndex        =   8
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   5100
      TabIndex        =   7
      Top             =   3840
      Width           =   2775
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
      Left            =   5100
      TabIndex        =   6
      Top             =   2460
      Width           =   2775
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
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
      Index           =   0
      Left            =   5100
      TabIndex        =   5
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Shape shp 
      BorderColor     =   &H80000010&
      Height          =   2655
      Index           =   1
      Left            =   5040
      Top             =   1680
      Width           =   6015
   End
   Begin VB.Label lbl 
      BackColor       =   &H80000014&
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
      Index           =   3
      Left            =   5040
      TabIndex        =   4
      Top             =   1680
      Width           =   6015
   End
   Begin VB.Label lbl 
      BackColor       =   &H80000014&
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
      Index           =   4
      Left            =   5040
      TabIndex        =   11
      Top             =   4440
      Width           =   6015
   End
   Begin VB.Label lbl 
      BackColor       =   &H80000014&
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
      Index           =   11
      Left            =   60
      TabIndex        =   18
      Top             =   60
      Width           =   4875
   End
End
Attribute VB_Name = "frmRecharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ CARD RECHARGE FORM _________________________|
'|  Name:          frmRecharge.frm                                             |
'|  Type:          VB6 Form                                                    |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
'|  Description:   Cyber Card recharge form                                    |
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

' $Id: frmRecharge.frm 15 2006-01-01 16:12:35Z sam $

Option Explicit

Dim mintRechargeType As enuCustomerType

Dim msngRechargePrices() As Single
Dim mdteRechargeTimes() As Date
Dim mblnRechargeUnlimited() As Boolean
Dim mlngRechargeValidities() As Single
Dim mlngRechargeValidityTypes() As Integer

Dim mlngSelectedRechargeIndex As Long
Dim mstrSelectedRechargeLabel As String
Dim mdteSelectedRechargeTime As Date
Dim mblnSelectedRechargeUnlimited As Boolean
Dim mlngSelectedRechargeValidity As Long
Dim mintSelectedRechargeValidityType As Integer
Dim muSelectedCard As New clsCard
Dim muSelectedClient As New clsClient
Dim msngSelectedRechargePrice As Single
Dim msngSelectedCustomerIncomeTotal As Single
Dim mdteTimeCreditNew, mdteExpiryDateNew As Date
Dim mlngSelectedRechargeOrder As Long


'*******************************************************************************
Private Sub cmdCancel_Click()
On Error GoTo cmdCancel_ClickError

    Call Unload(Me)
    Exit Sub
    
cmdCancel_ClickError:
    Call funErrorManager("frmRecharge.cmdCancel_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdCrdRecharge_Click()
Dim sngNewIncomeTotal As Single
Dim intNumBonuses As Integer
Dim dteBonusTime As Date
Dim i As Integer
Dim strSelectedRechargeLabel As String
On Error GoTo cmdCrdRecharge_ClickError
    
    If mlngSelectedRechargeIndex = 0 Then
        Call funMsg(enuMsg.RechargeNotSelected)
    Else
        tmrUpdateCard.enabled = False
        For i = 1 To CInt(txtQuantity.Text)
            Select Case mintRechargeType
            Case enuCustomerType.Card
                If guServOpts.uCards.intBonusActive = vbChecked And guServOpts.uCards.sngBonusThreshold > 0 And Not mblnRechargeUnlimited Then
                    sngNewIncomeTotal = msngSelectedRechargePrice + msngSelectedCustomerIncomeTotal
                    If Int(msngSelectedCustomerIncomeTotal / guServOpts.uCards.sngBonusThreshold) < Int(sngNewIncomeTotal / guServOpts.uCards.sngBonusThreshold) Then
                        intNumBonuses = Int(sngNewIncomeTotal / guServOpts.uCards.sngBonusThreshold) - Int(msngSelectedCustomerIncomeTotal / guServOpts.uCards.sngBonusThreshold)
                        dteBonusTime = guServOpts.uCards.dteBonusTime * intNumBonuses
                        mdteTimeCreditNew = mdteTimeCreditNew + dteBonusTime
                        strSelectedRechargeLabel = mstrSelectedRechargeLabel & funLoadResString(10546) & funFormatTime(dteBonusTime)
                        Call funMsg(CardBonusReached, funFormatCurrencyString(sngNewIncomeTotal, guServOpts.uConfig.strMonFormat, guServOpts.uConfig.strMonSymbol), "", funFormatTime(dteBonusTime, enuTimeFormat.Hh_Nmn_Ssec_dyn))
                    Else
                        strSelectedRechargeLabel = mstrSelectedRechargeLabel
                    End If
                Else
                    strSelectedRechargeLabel = mstrSelectedRechargeLabel
                End If
                muSelectedCard.DateExpiry = mdteExpiryDateNew
                muSelectedCard.TimeCredit = mdteTimeCreditNew
                muSelectedCard.UnlimitedPlay = mblnSelectedRechargeUnlimited
                muSelectedCard.ForceRefresh = True
                msngSelectedCustomerIncomeTotal = msngSelectedCustomerIncomeTotal + msngSelectedRechargePrice
                Call subCardRecharge(muSelectedCard, msngSelectedRechargePrice, mdteSelectedRechargeTime + dteBonusTime, funLoadResString(10289) & " " & IIf(mblnSelectedRechargeUnlimited, funLoadResString(11000), funFormatTime(mdteSelectedRechargeTime)) & " " & strSelectedRechargeLabel, funCardInUse(muSelectedCard))
            
            Case enuCustomerType.Client
                If guServOpts.uClients.intBonusActive = vbChecked And guServOpts.uClients.sngBonusThreshold > 0 And Not mblnRechargeUnlimited Then
                    sngNewIncomeTotal = msngSelectedRechargePrice + msngSelectedCustomerIncomeTotal
                    If Int(msngSelectedCustomerIncomeTotal / guServOpts.uClients.sngBonusThreshold) < Int(sngNewIncomeTotal / guServOpts.uClients.sngBonusThreshold) Then
                        intNumBonuses = Int(sngNewIncomeTotal / guServOpts.uClients.sngBonusThreshold) - Int(msngSelectedCustomerIncomeTotal / guServOpts.uClients.sngBonusThreshold)
                        dteBonusTime = guServOpts.uClients.dteBonusTime * intNumBonuses
                        mdteTimeCreditNew = mdteTimeCreditNew + dteBonusTime
                        strSelectedRechargeLabel = mstrSelectedRechargeLabel & funLoadResString(10546) & funFormatTime(dteBonusTime)
                        Call funMsg(ClientBonusReached, funFormatCurrencyString(sngNewIncomeTotal, guServOpts.uConfig.strMonFormat, guServOpts.uConfig.strMonSymbol), "", funFormatTime(dteBonusTime, enuTimeFormat.Hh_Nmn_Ssec_dyn))
                    Else
                        strSelectedRechargeLabel = mstrSelectedRechargeLabel
                    End If
                Else
                    strSelectedRechargeLabel = mstrSelectedRechargeLabel
                End If
                muSelectedClient.DateExpiry = mdteExpiryDateNew
                muSelectedClient.TimeCredit = mdteTimeCreditNew
                muSelectedClient.UnlimitedPlay = mblnSelectedRechargeUnlimited
                muSelectedClient.ForceRefresh = True
                msngSelectedCustomerIncomeTotal = msngSelectedCustomerIncomeTotal + msngSelectedRechargePrice
                Call subClientRecharge(muSelectedClient, msngSelectedRechargePrice, mdteSelectedRechargeTime + dteBonusTime, funLoadResString(10289) & " " & IIf(mblnSelectedRechargeUnlimited, funLoadResString(11000), funFormatTime(mdteSelectedRechargeTime)) & " " & strSelectedRechargeLabel, funClientInUse(muSelectedClient))
            
            Case enuCustomerType.SelfServicePrePaid
                strSelectedRechargeLabel = mstrSelectedRechargeLabel
                muSelectedClient.TimeCredit = mdteTimeCreditNew
                 Call subSelfservicePrePaidRecharge(muSelectedClient, msngSelectedRechargePrice, mdteSelectedRechargeTime, funLoadResString(10289) & " " & funFormatTime(mdteSelectedRechargeTime) & " " & strSelectedRechargeLabel, funSelfServicePrePaidRunning(muSelectedClient))
            
            End Select
        Next i
        Call Unload(Me)
    End If
    Exit Sub
    
cmdCrdRecharge_ClickError:
    Call funErrorManager("cmdCrdRecharge_ClickError", True)
End Sub


'*******************************************************************************
Private Sub Form_Activate()
Dim i As Long
On Error GoTo Form_ActivateError

    Set muSelectedCard = guSelectedCard
    Set muSelectedClient = guSelectedClient
    Call subLocalizeInterface
    cmdCrdRecharge.Picture = imlRecharge.ListImages(mintRechargeType).Picture
    cmdCancel.Picture = imlCancel.ListImages(mintRechargeType).Picture
    Select Case mintRechargeType
    Case enuCustomerType.Card
        Call subQueryIncome(muSelectedCard.DateCreation, NULL_DATE, enuIncomeType.Card, muSelectedCard.ID)
        lblCrdID.Caption = funLoadResString(10147) & muSelectedCard.ID
        lblCrdCredit.Caption = " " & funFormatTime(muSelectedCard.TimeCredit, enuTimeFormat.Hh_Nmn_Ssec_dyn)
        lblCrdDteCreation.Caption = " " & format(muSelectedCard.DateCreation, "Long Date")
        lblCrdDteExpiry.Caption = " " & format(muSelectedCard.DateExpiry, "Long Date")
        cmdCrdRecharge.ToolTipText = funLoadResString(10472)
    
    Case enuCustomerType.Client
        Call subQueryIncome(muSelectedClient.DateCreation, NULL_DATE, enuIncomeType.Client, , , muSelectedClient.login)
        lblCrdID.Caption = funLoadResString(10800) & muSelectedClient.login
        lblCrdCredit.Caption = " " & funFormatTime(muSelectedClient.TimeCredit, enuTimeFormat.Hh_Nmn_Ssec_dyn)
        lblCrdDteCreation.Caption = " " & format(muSelectedClient.DateCreation, "Long Date")
        lbl(1).enabled = guServOpts.uClients.intEnableExpiry = vbChecked
        lbl(6).enabled = guServOpts.uClients.intEnableExpiry = vbChecked
        lbl(10).enabled = guServOpts.uClients.intEnableExpiry = vbChecked
        lblCrdDteExpiry.enabled = guServOpts.uClients.intEnableExpiry = vbChecked
        lblCrdDteExpiry.Caption = IIf(guServOpts.uClients.intEnableExpiry = vbChecked, " " & format(muSelectedClient.DateExpiry, "Long Date"), funLoadResString(10865))
        lblRechargeValidity.enabled = guServOpts.uClients.intEnableExpiry = vbChecked
        lblCrdDteExpiryNew.enabled = guServOpts.uClients.intEnableExpiry = vbChecked
        cmdCrdRecharge.ToolTipText = funLoadResString(10987)
    
    Case enuCustomerType.SelfServicePrePaid
        Call subQueryIncome(muSelectedClient.DateCreation, NULL_DATE, enuIncomeType.SelfServicePrePaid, , , muSelectedClient.login)
        lblCrdID.Caption = " " & funLoadResString(10983)
        lblCrdCredit.Caption = " " & funFormatTime(muSelectedClient.TimeCredit, enuTimeFormat.Hh_Nmn_Ssec_dyn)
        lblCrdDteCreation.Caption = " " & format(muSelectedClient.DateCreation, "Long Date")
        lblCrdDteExpiry.Caption = " " & funLoadResString(10865)
        lblCrdDteExpiry.enabled = False
        lblRechargeValidity.enabled = False
        lblCrdDteExpiryNew.enabled = False
        lbl(1).enabled = False
        lbl(6).enabled = False
        lbl(10).enabled = False
        cmdCrdRecharge.ToolTipText = funLoadResString(10988)
        If muSelectedClient.login = "" Then cmdCrdRecharge.Caption = funLoadResString(10980)
    
    End Select
    mlngSelectedRechargeOrder = enuRechargeOrder.RechargesIndex
    Call subFillRecharges
    msngSelectedCustomerIncomeTotal = 0
    If Not (grstJournIncomeQry.BOF = True And grstJournIncomeQry.EOF = True) Then
        Call grstJournIncomeQry.MoveFirst
        For i = 1 To grstJournIncomeQry.RecordCount
            msngSelectedCustomerIncomeTotal = msngSelectedCustomerIncomeTotal + CSng(grstJournIncomeQry!income)
            Call grstJournIncomeQry.MoveNext
            DoEvents
        Next i
    End If
    mlngSelectedRechargeIndex = 0
    Exit Sub
    
Form_ActivateError:
    Call funErrorManager("frmRecharge.Form_ActivateError", True)
End Sub


'*******************************************************************************
Private Sub Form_Load()
On Error GoTo Form_LoadError
    
    gblnfrmRecharge = True
    Call subSetDefaultFont(Me)
    With fxgRecharges
        .left = shp(3).left + 1
        .width = shp(3).width - 2
        .Cols = 4
        .top = shp(3).top + lbl(11).height
        .height = shp(3).height - lbl(11).height - 1
        
        .Row = 0
        .Col = 0
        .ColWidth(0) = 1000
        .ColAlignment(0) = flexAlignCenterTop
        
        .Col = 1
        .ColWidth(.Col) = 650
        .ColAlignment(.Col) = flexAlignRightTop
        
        .Col = 2
        .ColWidth(.Col) = 1000
        .ColAlignment(.Col) = flexAlignLeftTop
        
        .Col = 3
        .ColWidth(.Col) = 40000
        .ColAlignment(.Col) = flexAlignLeftTop
        
    End With
    
    Exit Sub
    
Form_LoadError:
    Call funErrorManager("frmRecharge.Form_LoadError", True)
End Sub


'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)

    gblnfrmRecharge = False
End Sub


'*******************************************************************************
Private Sub fxgRecharges_Click()
On Error GoTo fxgRecharges_ClickError
    
    If mlngSelectedRechargeIndex = 0 Then Call fxgRecharges_RowColChange
    Exit Sub
    
fxgRecharges_ClickError:
    Call funErrorManager("frmRecharge.fxgRecharges_ClickError", True)
End Sub


'*******************************************************************************
Private Sub fxgRecharges_DblClick()
On Error GoTo fxgRecharges_DblClickError
    
    With fxgRecharges
        If .MouseRow = 0 Then
            If mlngSelectedRechargeIndex > 0 Then
                .Row = mlngSelectedRechargeIndex
                .ColSel = .Cols - 1
                .CellForeColor = vbWindowText
                .CellBackColor = vbWindowBackground
            End If
            
            Select Case .MouseCol
            Case 0
                mlngSelectedRechargeOrder = enuRechargeOrder.orderByTime
            Case 1
                mlngSelectedRechargeOrder = enuRechargeOrder.orderByPrice
            Case 2
                mlngSelectedRechargeOrder = enuRechargeOrder.orderByValidity
            Case 3
                mlngSelectedRechargeOrder = enuRechargeOrder.RechargesIndex
            End Select
            Call subFillRecharges
            .Row = 0
            mlngSelectedRechargeIndex = 0
            
            lblRechargeCaption.Caption = ""
            lblRechargeCredit.Caption = ""
            lblRechargePrice.Caption = ""
            lblRechargeValidity.Caption = ""
            lblRechargeCreditNew.Caption = ""
        
        End If
    End With
    Exit Sub
    
fxgRecharges_DblClickError:
    Call funErrorManager("frmOptions.fxgRecharges_DblClickError", True)
End Sub


'*******************************************************************************
Private Sub fxgRecharges_RowColChange()
Dim lngNewSelectedRowIndex As Long
On Error GoTo fxgRecharges_RowColChangeError
    
    With fxgRecharges
        lngNewSelectedRowIndex = .Row
        If lngNewSelectedRowIndex < 1 Then Exit Sub
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
        
        lblRechargeValidity.Caption = IIf(mintRechargeType = enuCustomerType.Card Or (mintRechargeType = enuCustomerType.Client And guServOpts.uClients.intEnableExpiry = vbChecked), " " & .TextMatrix(lngNewSelectedRowIndex, 2), funLoadResString(10865))
        lblRechargeCaption.Caption = " " & .TextMatrix(lngNewSelectedRowIndex, 3)
        lblRechargePrice.Caption = " " & funFormatCurrencyString(msngRechargePrices(lngNewSelectedRowIndex) * CInt(txtQuantity.Text), guServOpts.uConfig.strMonFormat, guServOpts.uConfig.strMonSymbol)
        
        mdteSelectedRechargeTime = mdteRechargeTimes(lngNewSelectedRowIndex)
        msngSelectedRechargePrice = msngRechargePrices(lngNewSelectedRowIndex)
        mlngSelectedRechargeValidity = mlngRechargeValidities(lngNewSelectedRowIndex)
        mintSelectedRechargeValidityType = mlngRechargeValidityTypes(lngNewSelectedRowIndex)
        mstrSelectedRechargeLabel = CStr(.TextMatrix(lngNewSelectedRowIndex, 3))
        mblnSelectedRechargeUnlimited = mblnRechargeUnlimited(lngNewSelectedRowIndex)
    End With
    txtQuantity.enabled = True
    udQuantity.enabled = True
    Call subUpdateUI
    Exit Sub
    
fxgRecharges_RowColChangeError:
    Call funErrorManager("frmRecharge.fxgRecharges_RowColChangeError", True)
End Sub


'*******************************************************************************
Private Sub subFillRecharges()
Dim i As Long
On Error GoTo subFillRechargesError

    Select Case mintRechargeType
    Case enuCustomerType.Card
        If muSelectedCard.UnlimitedPlay Then
            Call subQueryRecharges(True, False, False, False, mlngSelectedRechargeOrder, True)
        Else
            Call subQueryRecharges(True, False, False, False, mlngSelectedRechargeOrder)
        End If
    Case enuCustomerType.Client
        If muSelectedClient.UnlimitedPlay Then
            Call subQueryRecharges(False, True, False, False, mlngSelectedRechargeOrder, True)
        Else
            Call subQueryRecharges(False, True, False, False, mlngSelectedRechargeOrder)
        End If
    Case enuCustomerType.SelfServicePrePaid
        Call subQueryRecharges(False, False, True, False, mlngSelectedRechargeOrder, False)
    End Select
    
    With fxgRecharges
        If grstRechargesQry.BOF = True And grstRechargesQry.EOF = True Then
            .Rows = 21
            .enabled = False
            Exit Sub
        End If
        .Font = ""
        .Row = 0
        For i = 0 To .Cols - 1
            .Col = i
            .CellFontBold = False
        Next i
        Select Case mlngSelectedRechargeOrder
        Case enuRechargeOrder.orderByTime
            .Col = 0
        Case enuRechargeOrder.orderByPrice
            .Col = 1
        Case enuRechargeOrder.orderByValidity
            .Col = 2
        Case Else
            .Col = 3
        End Select
        .CellFontBold = True
        
        .Rows = grstRechargesQry.RecordCount + 1
        ReDim msngRechargePrices(1 To grstRechargesQry.RecordCount)
        ReDim mdteRechargeTimes(1 To grstRechargesQry.RecordCount)
        ReDim mblnRechargeUnlimited(1 To grstRechargesQry.RecordCount)
        ReDim mlngRechargeValidities(1 To grstRechargesQry.RecordCount)
        ReDim mlngRechargeValidityTypes(1 To grstRechargesQry.RecordCount)
        For i = 1 To grstRechargesQry.RecordCount
            .TextMatrix(i, 0) = IIf(grstRechargesQry!unlimited_play, funLoadResString(11000), funFormatTime(grstRechargesQry!Time))
            .TextMatrix(i, 1) = funFormatCurrencyString(grstRechargesQry!Price, guServOpts.uConfig.strMonFormat)
            Select Case grstRechargesQry!validity_type
            Case enuRechargeValidityType.Hours
                .TextMatrix(i, 2) = grstRechargesQry!validity & " " & funLoadResString(10995)
            Case enuRechargeValidityType.Days
                .TextMatrix(i, 2) = grstRechargesQry!validity & " " & funLoadResString(10996)
            Case enuRechargeValidityType.Weeks
                .TextMatrix(i, 2) = grstRechargesQry!validity & " " & funLoadResString(10997)
            Case enuRechargeValidityType.Months
                .TextMatrix(i, 2) = grstRechargesQry!validity & " " & funLoadResString(10998)
            Case enuRechargeValidityType.Years
                .TextMatrix(i, 2) = grstRechargesQry!validity & " " & funLoadResString(10999)
            End Select
            .TextMatrix(i, 3) = IIf(IsNull(grstRechargesQry!Caption), "", grstRechargesQry!Caption)
            msngRechargePrices(i) = grstRechargesQry!Price
            mdteRechargeTimes(i) = grstRechargesQry!Time
            mblnRechargeUnlimited(i) = grstRechargesQry!unlimited_play
            mlngRechargeValidities(i) = grstRechargesQry!validity
            mlngRechargeValidityTypes(i) = grstRechargesQry!validity_type
            Call grstRechargesQry.MoveNext
        Next i
    End With
    Exit Sub
    
subFillRechargesError:
    Call funErrorManager("frmRecharge.subFillRechargesError")
End Sub


'*******************************************************************************
Public Sub subLocalizeInterface()
On Error GoTo subLocalizeInterfaceError

    Call subSetCharset(Me, guServOpts.uConfig.intInterfaceLangCharset)
    
    Me.Caption = funLoadResString(10000) & " - " & IIf(mintRechargeType = enuCustomerType.Card, funLoadResString(10131), IIf(mintRechargeType = enuCustomerType.Client, funLoadResString(10819), funLoadResString(10983)))
    lbl(7).Caption = funLoadResString(10133)
    lbl(0).Caption = funLoadResString(10133)
    lbl(5).Caption = funLoadResString(10133)
    lbl(9).Caption = funLoadResString(10134)
    lbl(6).Caption = funLoadResString(10135)
    lbl(3).Caption = funLoadResString(10136)
    lbl(4).Caption = funLoadResString(10137)
    lbl(1).Caption = funLoadResString(10138)
    lbl(8).Caption = funLoadResString(10139)
    lbl(2).Caption = funLoadResString(10140)
    lbl(10).Caption = funLoadResString(10135)
    lbl(11).Caption = funLoadResString(10132)
    lbl(12).Caption = funLoadResString(10861)
    txtQuantity.ToolTipText = funLoadResString(10862)
    udQuantity.ToolTipText = funLoadResString(10862)
    cmdCrdRecharge.Caption = funLoadResString(10141)
    cmdCancel.Caption = funLoadResString(10123)
    cmdCancel.ToolTipText = funLoadResString(10473)
    
    With fxgRecharges
        .TextMatrix(0, 0) = funLoadResString(10143)
        .TextMatrix(0, 1) = funLoadResString(10144)
        .TextMatrix(0, 2) = funLoadResString(10145)
        .TextMatrix(0, 3) = funLoadResString(10146)
        .ToolTipText = funLoadResString(10471) & funLoadResString(10141)
    End With
    Exit Sub
    
subLocalizeInterfaceError:
    Call funErrorManager("frmRecharge.subLocalizeInterfaceError", True)
End Sub


Private Sub lblRechargeCaption_Click()
169
End Sub

'*******************************************************************************
Private Sub tmrUpdateCard_Timer()
On Error GoTo tmrUpdateCard_TimerError

    lblCrdCredit.Caption = " " & funFormatTime(IIf(mintRechargeType = enuCustomerType.Card, muSelectedCard.TimeCredit, muSelectedClient.TimeCredit), enuTimeFormat.Hh_Nmn_Ssec_dyn)
    If mlngSelectedRechargeIndex >= 1 Then Call subUpdateUI
    Exit Sub
    
tmrUpdateCard_TimerError:
    Call funErrorManager("frmRecharge.tmrUpdateCard_TimerError", True)
End Sub


'*******************************************************************************
Public Sub subUpdateUI()
Dim dteExpiryDateNew As Date
On Error GoTo subUpdateUIError

    If mblnSelectedRechargeUnlimited Then
        txtQuantity.Text = 1
        txtQuantity.enabled = False
        udQuantity.enabled = False
    Else
        txtQuantity.enabled = True
        udQuantity.enabled = True
    End If
        
    lblRechargeCredit.Caption = " " & IIf(mblnSelectedRechargeUnlimited, funLoadResString(11000), funFormatTime(CDate(mdteSelectedRechargeTime), enuTimeFormat.Hh_Nmn_Ssec_dyn))

    mdteTimeCreditNew = mdteSelectedRechargeTime * CInt(txtQuantity.Text) + IIf(mintRechargeType = enuCustomerType.Card, muSelectedCard.TimeCredit, muSelectedClient.TimeCredit)
    lblRechargeCreditNew.Caption = " " & IIf(mblnSelectedRechargeUnlimited, funLoadResString(11000), funFormatTime(CDate(mdteTimeCreditNew), enuTimeFormat.Hh_Nmn_Ssec_dyn))
    
    Select Case mintSelectedRechargeValidityType
    Case enuRechargeValidityType.Hours
        dteExpiryDateNew = DateAdd("h", mlngSelectedRechargeValidity, Date)
    Case enuRechargeValidityType.Days
        dteExpiryDateNew = DateAdd("d", mlngSelectedRechargeValidity, Date)
    Case enuRechargeValidityType.Weeks
        dteExpiryDateNew = DateAdd("ww", mlngSelectedRechargeValidity, Date)
    Case enuRechargeValidityType.Months
        dteExpiryDateNew = DateAdd("d", mlngSelectedRechargeValidity * guServOpts.uConfig.intNumDaysInMonth, Date)
    Case enuRechargeValidityType.Years
        dteExpiryDateNew = DateAdd("d", mlngSelectedRechargeValidity * NUM_DAYS_IN_YEAR, Date)
    End Select
    
    If mblnSelectedRechargeUnlimited Or dteExpiryDateNew > IIf(mintRechargeType = enuCustomerType.Card, muSelectedCard.DateExpiry, muSelectedClient.DateExpiry) Then
        mdteExpiryDateNew = dteExpiryDateNew
    Else
        mdteExpiryDateNew = IIf(mintRechargeType = enuCustomerType.Card, muSelectedCard.DateExpiry, muSelectedClient.DateExpiry)
    End If
    lblCrdDteExpiryNew.Caption = IIf(mintRechargeType = enuCustomerType.Card Or (mintRechargeType = enuCustomerType.Client And guServOpts.uClients.intEnableExpiry = vbChecked), " " & format(mdteExpiryDateNew, "Long Date"), funLoadResString(10865))
    
    Exit Sub
    
subUpdateUIError:
    Call funErrorManager("frmRecharge.subUpdateUIError")
End Sub


'*******************************************************************************
Public Sub subSetRechargeType(intType As enuCustomerType)
On Error GoTo subSetRechargeTypeError
    
    mintRechargeType = intType
    Exit Sub
    
subSetRechargeTypeError:
    Call funErrorManager("frmRecharge.subSetRechargeTypeError")
End Sub


'*******************************************************************************
Private Sub txtQuantity_Change()
On Error Resume Next

    udQuantity.value = txtQuantity.Text
    lblRechargePrice = " " & funFormatCurrencyString(msngRechargePrices(mlngSelectedRechargeIndex) * CInt(txtQuantity.Text), guServOpts.uConfig.strMonFormat, guServOpts.uConfig.strMonSymbol)
    Call subUpdateUI
End Sub


'*******************************************************************************
Private Sub txtQuantity_GotFocus()
On Error Resume Next

    txtQuantity.SelStart = 0
    txtQuantity.SelLength = Len(txtQuantity.Text)
End Sub


'*******************************************************************************
Private Sub txtQuantity_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    
    Select Case KeyCode
    Case 40
        If udQuantity.value > udQuantity.Min Then udQuantity.value = udQuantity.value - udQuantity.Increment
    Case 38
        If udQuantity.value < udQuantity.Max Then udQuantity.value = udQuantity.value + udQuantity.Increment
    End Select
End Sub


'*******************************************************************************
Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub


'*******************************************************************************
Private Sub txtQuantity_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    Select Case KeyCode
    Case 38, 40
        txtQuantity.SelStart = 0
        txtQuantity.SelLength = Len(txtQuantity.Text)
    End Select
End Sub


'*******************************************************************************
Private Sub txtQuantity_LostFocus()
On Error Resume Next
    
    If CInt(txtQuantity.Text) >= udQuantity.Min And CInt(txtQuantity.Text) <= udQuantity.Max Then
        udQuantity.value = CInt(txtQuantity.Text)
    Else
        txtQuantity.Text = udQuantity.value
    End If
End Sub


'*******************************************************************************
Private Sub udQuantity_Change()
On Error Resume Next

    txtQuantity.Text = udQuantity.value
End Sub


'*******************************************************************************
Private Sub udQuantity_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

    Call funSetFocus(txtQuantity)
    txtQuantity.SelStart = 0
    txtQuantity.SelLength = Len(txtQuantity.Text)
End Sub
