VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{01646141-065C-11D4-8ED3-00E07D815373}#1.0#0"; "mbbrowse.ocx"
Begin VB.Form frmCardMassAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "_"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10335
   ControlBox      =   0   'False
   Icon            =   "frmCardMassAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbCrdType 
      Height          =   315
      ItemData        =   "frmCardMassAdd.frx":12FA
      Left            =   7380
      List            =   "frmCardMassAdd.frx":12FC
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   3180
      Width           =   2775
   End
   Begin VB.TextBox txtCrdComment 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   7440
      MaxLength       =   100
      TabIndex        =   23
      Top             =   3660
      Width           =   2655
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   1
      Left            =   180
      TabIndex        =   20
      Top             =   420
      Width           =   4875
      Begin VB.OptionButton optCrdCreateEnabled 
         Caption         =   "_"
         Height          =   195
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   5475
      End
      Begin VB.OptionButton optCrdCreatePending 
         Caption         =   "_"
         Height          =   195
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   5475
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   0
      Left            =   180
      TabIndex        =   19
      Top             =   1440
      Width           =   4875
      Begin VB.OptionButton optRecharge 
         Caption         =   "_"
         Height          =   195
         Left            =   0
         TabIndex        =   3
         Top             =   240
         Width           =   5415
      End
      Begin VB.OptionButton optNoRecharge 
         Caption         =   "_"
         Height          =   195
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   5415
      End
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "_"
      Default         =   -1  'True
      Height          =   375
      Left            =   5280
      TabIndex        =   12
      Top             =   4140
      Width           =   3555
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "_"
      Height          =   375
      Left            =   8940
      TabIndex        =   13
      Top             =   4140
      Width           =   1335
   End
   Begin MBBrowse.BrowseFF bffPath 
      Left            =   5820
      Top             =   4080
      _ExtentX        =   1085
      _ExtentY        =   1085
      StartUpPosition =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdCrdTemplatePath 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9780
      TabIndex        =   8
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox txtCrdTemplatePath 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   5460
      TabIndex        =   7
      Top             =   1740
      Width           =   4215
   End
   Begin VB.CommandButton cmdCrdListPath 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9780
      TabIndex        =   11
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox txtCrdListPath 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   5460
      TabIndex        =   10
      Top             =   2340
      Width           =   4215
   End
   Begin VB.OptionButton optCrdNumTemplate 
      Caption         =   "_"
      Height          =   195
      Left            =   5400
      TabIndex        =   6
      Top             =   1440
      Width           =   4815
   End
   Begin VB.OptionButton optCrdNumList 
      Caption         =   "_"
      Height          =   195
      Left            =   5400
      TabIndex        =   9
      Top             =   2040
      Width           =   4815
   End
   Begin VB.TextBox txtCardQuantity 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   8940
      MaxLength       =   4
      TabIndex        =   5
      Top             =   540
      Width           =   795
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fxgRecharges 
      Height          =   810
      Left            =   360
      TabIndex        =   4
      Top             =   2040
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   1429
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
   Begin MSComCtl2.UpDown udCardQuantity 
      Height          =   315
      Left            =   9840
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   480
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Value           =   1
      OrigLeft        =   1920
      OrigTop         =   300
      OrigRight       =   2160
      OrigBottom      =   555
      Max             =   10000
      Min             =   1
      Enabled         =   -1  'True
   End
   Begin MSComDlg.CommonDialog cdlPath 
      Left            =   6960
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
      Height          =   255
      Index           =   19
      Left            =   5340
      TabIndex        =   26
      Top             =   3240
      Width           =   1995
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
      Height          =   255
      Index           =   20
      Left            =   5340
      TabIndex        =   25
      Top             =   3660
      Width           =   1995
   End
   Begin VB.Shape shpCrdComment 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   315
      Left            =   7380
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Shape shp 
      BorderColor     =   &H80000010&
      Height          =   1215
      Index           =   8
      Left            =   5280
      Top             =   2820
      Width           =   4995
   End
   Begin VB.Shape shp 
      BorderColor     =   &H80000010&
      Height          =   915
      Index           =   6
      Left            =   60
      Top             =   60
      Width           =   5115
   End
   Begin VB.Shape shp 
      BorderColor     =   &H80000010&
      Height          =   3435
      Index           =   4
      Left            =   60
      Top             =   1080
      Width           =   5115
   End
   Begin VB.Shape shp 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   315
      Index           =   2
      Left            =   5400
      Top             =   1680
      Width           =   4335
   End
   Begin VB.Shape shp 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   315
      Index           =   7
      Left            =   5400
      Top             =   2280
      Width           =   4335
   End
   Begin VB.Shape shp 
      BorderColor     =   &H80000010&
      Height          =   1635
      Index           =   1
      Left            =   5280
      Top             =   1080
      Width           =   4995
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
      Left            =   5280
      TabIndex        =   18
      Top             =   1080
      Width           =   4995
   End
   Begin VB.Shape shp 
      BorderColor     =   &H80000010&
      Height          =   915
      Index           =   0
      Left            =   5280
      Top             =   60
      Width           =   4995
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
      Index           =   2
      Left            =   5280
      TabIndex        =   17
      Top             =   60
      Width           =   4995
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
      Height          =   195
      Index           =   0
      Left            =   5400
      TabIndex        =   16
      Top             =   540
      Width           =   3435
   End
   Begin VB.Shape shp 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   315
      Index           =   5
      Left            =   8880
      Top             =   480
      Width           =   915
   End
   Begin VB.Shape shp 
      BorderColor     =   &H80000010&
      Height          =   2475
      Index           =   3
      Left            =   180
      Top             =   1920
      Width           =   4875
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
      Index           =   1
      Left            =   60
      TabIndex        =   14
      Top             =   1080
      Width           =   5115
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
      Left            =   60
      TabIndex        =   21
      Top             =   60
      Width           =   5115
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
      Index           =   5
      Left            =   5280
      TabIndex        =   22
      Top             =   2820
      Width           =   4995
   End
End
Attribute VB_Name = "frmCardMassAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ CARD MASS CREATION FORM ____________________|
'|  Name:          frmCardMassAdd.frm                                          |
'|  Type:          VB6 Form                                                    |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
'|  Description:   Card mass creation from.                                    |
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

' $Id: frmCardMassAdd.frm 24 2006-01-18 17:35:36Z sam $

Option Explicit
Dim msngRechargePrices() As Single
Dim mdteRechargeTimes() As Date
Dim mblnRechargeUnlimited() As Boolean
Dim mlngRechargeValidities() As Long
Dim mintRechargeValidityTypes() As Integer
Dim mlngSelectedRechargeIndex As Long
Dim mstrSelectedRechargeLabel As String
Dim mdteSelectedRechargeTime As Date
Dim mblnSelectedRechargeUnlimited As Boolean
Dim mlngSelectedRechargeValidity As Long
Dim mintSelectedRechargeValidityType As Integer
Dim msngSelectedRechargePrice As Single


'*******************************************************************************
Public Sub subLocalizeInterface()
On Error GoTo subLocalizeInterfaceError

    Call subSetCharset(Me, guServOpts.uConfig.intInterfaceLangCharset)
    
    Me.Caption = funLoadResString(10000) & " - " & funLoadResString(10918)

    lbl(0).Caption = " " & funLoadResString(10925)
    lbl(1).Caption = " " & funLoadResString(10922)
    lbl(2).Caption = " " & funLoadResString(10923)
    lbl(3).Caption = " " & funLoadResString(10924)
    lbl(4).Caption = " " & funLoadResString(10932)
    lbl(5).Caption = " " & funLoadResString(10372)

    lbl(19).Caption = funLoadResString(10527)
    lbl(20).Caption = funLoadResString(10528)


    optCrdNumList.Caption = funLoadResString(10926)
    optCrdNumTemplate.Caption = funLoadResString(10927)
    optNoRecharge.Caption = funLoadResString(10930)
    optRecharge.Caption = funLoadResString(10931)
    optCrdCreatePending.Caption = funLoadResString(10934)
    optCrdCreateEnabled.Caption = funLoadResString(10933)
    
    With fxgRecharges
        .TextMatrix(0, 0) = funLoadResString(10143)
        .TextMatrix(0, 1) = funLoadResString(10144)
        .TextMatrix(0, 2) = funLoadResString(10145)
        .TextMatrix(0, 3) = funLoadResString(10146)
        .ToolTipText = funLoadResString(10471) '& funLoadResString(10141)
    End With

    cmdCancel.Caption = funLoadResString(10928)
    cmdGenerate.Caption = funLoadResString(10929)
    bffPath.Caption = funLoadResString(10936)
    bffPath.Prompt = funLoadResString(10935)


'    cmdClose.Caption = funLoadResString(10012)
    Exit Sub

subLocalizeInterfaceError:
    Call funErrorManager("frmCardMassAdd.subLocalizeInterfaceError", True)
End Sub


'*******************************************************************************
Private Sub cmdCancel_Click()
On Error GoTo cmdCancel_ClickError

    Call Unload(Me)
    Exit Sub
    
cmdCancel_ClickError:
    Call funErrorManager("frmCardMassAdd.cmdCancel_ClickError", True)
End Sub


Private Sub cmdCrdListPath_Click()
On Error GoTo cmdCrdListPath_ClickError
    
    bffPath.InitialFolder = txtCrdListPath.Text
    If bffPath.Browse Then txtCrdListPath.Text = bffPath.SelectedItem
    Exit Sub

cmdCrdListPath_ClickError:
    Call funErrorManager("frmCardMassAdd.cmdCrdListPath_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdCrdTemplatePath_Click()
On Error GoTo cmdCrdTemplatePath_ClickError

    Call subCdlOpenFile(cdlPath, txtCrdTemplatePath, txtCrdTemplatePath.Text, funGetFolderPath(CSIDL_DESKTOP), funLoadResString(10937), funLoadResString(10000) & " - " & funLoadResString(10938))
    Exit Sub

cmdCrdTemplatePath_ClickError:
    Call funErrorManager("frmCardMassAdd.cmdCrdTemplatePath_ClickError", True)
End Sub


'*******************************************************************************
Private Sub cmdGenerate_Click()
Dim i As Long
Dim lngNewCardID As Long
Dim dteNewCardTimeStamp As Date
Dim dteNewCardDate As Date
Dim dteExpiryDateNew As Date
Dim uNewCard As New clsCard
Dim intInputFileNumber As Integer
Dim intOutputFileNumber As Integer
Dim strFilePath As String
Dim strFileSuffix As String
Dim strTemplateFile As String
Dim strOutputFile As String
Dim lngQtyCrdNum As Long
Dim lngCrdTemplateIndex As Long
Dim lngQtyCrdTemplate As Long
Dim lngCrdNumIndex As Long
Dim strRechargeTime As String
On Error GoTo cmdGenerate_ClickError

    If mblnSelectedRechargeUnlimited Then
        strRechargeTime = funLoadResString(11000)
    Else
        strRechargeTime = funFormatTime(mdteSelectedRechargeTime, enuTimeFormat.Hh_Nmn_Ssec_dyn)
    End If
    
    If funMsg(AskCanGenerateCards, _
              funLoadResString(10950, _
                               IIf(optCrdCreatePending, funLoadResString(10944), funLoadResString(10943)), _
                               txtCardQuantity.Text, _
                               IIf(optRecharge.value, strRechargeTime, "-"), _
                               IIf(optRecharge.value, funFormatCurrencyString(msngSelectedRechargePrice, guServOpts.uConfig.strMonFormat, guServOpts.uConfig.strMonSymbol), "-"))) = vbCancel Then Exit Sub
    Screen.MousePointer = vbHourglass
    dteNewCardTimeStamp = Now
    dteNewCardDate = Date
    intInputFileNumber = FreeFile
    intOutputFileNumber = FreeFile
    
    ' Init card number output
    If optCrdNumList.value Then
        strFilePath = txtCrdListPath & "\" & format(dteNewCardTimeStamp, "yymmdd-HhNnSs") & "_" & funLoadResString(10940) & ".txt"
        Open strFilePath For Append As intOutputFileNumber
        Print #intOutputFileNumber, funLoadResString(10941, format(dteNewCardTimeStamp, "yyyy.mm.dd-Hh:Nn:Ss"), udCardQuantity.value & " " & funLoadResString(IIf(optCrdCreateEnabled.value, 10943, 10944)), CStr(mdteSelectedRechargeTime), funFormatCurrencyString(msngSelectedRechargePrice, guServOpts.uConfig.strMonFormat, guServOpts.uConfig.strMonSymbol))
    Else
        intInputFileNumber = FreeFile
        strFilePath = txtCrdTemplatePath
        Open strFilePath For Binary Access Read As intInputFileNumber
        strTemplateFile = String(FileLen(strFilePath), " ")
        Get intInputFileNumber, , strTemplateFile
        Close intInputFileNumber
        
        strFileSuffix = right(strFilePath, Len(strFilePath) - InStrRev(strFilePath, "."))
        strFilePath = left(strFilePath, InStrRev(strFilePath, "\"))
        
        strOutputFile = strTemplateFile
        lngQtyCrdNum = 0
        lngCrdTemplateIndex = 1
        lngCrdNumIndex = 1 - Len(CARD_GEN_ID)
        Do
            lngCrdNumIndex = InStr(lngCrdNumIndex + Len(CARD_GEN_ID), strOutputFile, CARD_GEN_ID)
            If lngCrdNumIndex > 0 Then lngQtyCrdNum = lngQtyCrdNum + 1
        Loop While lngCrdNumIndex > 0
        If lngQtyCrdNum = 0 Then
            Call funMsg(CardTemplateNOK)
            Exit Sub
        End If
        lngQtyCrdTemplate = Int(udCardQuantity.value / lngQtyCrdNum)
        If lngQtyCrdTemplate < udCardQuantity.value / lngQtyCrdNum Then lngQtyCrdTemplate = lngQtyCrdTemplate + 1
        
        strOutputFile = Replace(strOutputFile, CARD_GEN_TITLE, funLoadResString(10946, format(dteNewCardTimeStamp, "yyyy.mm.dd-Hh:Nn:Ss")), , 1)
        strOutputFile = Replace(strOutputFile, CARD_GEN_HEADER, funLoadResString(10941, format(dteNewCardTimeStamp, "yyyy.mm.dd-Hh:Nn:Ss"), udCardQuantity.value & " " & funLoadResString(IIf(optCrdCreateEnabled.value, 10943, 10944)), strRechargeTime, funFormatCurrencyString(msngSelectedRechargePrice, guServOpts.uConfig.strMonFormat, guServOpts.uConfig.strMonSymbol)) & " - " & funLoadResString(10947, CStr(lngCrdTemplateIndex), CStr(lngQtyCrdTemplate)), , 1)
        strOutputFile = Replace(strOutputFile, CARD_GEN_FOOTER, funLoadResString(10941, format(dteNewCardTimeStamp, "yyyy.mm.dd-Hh:Nn:Ss"), udCardQuantity.value & " " & funLoadResString(IIf(optCrdCreateEnabled.value, 10943, 10944)), strRechargeTime, funFormatCurrencyString(msngSelectedRechargePrice, guServOpts.uConfig.strMonFormat, guServOpts.uConfig.strMonSymbol)) & " - " & funLoadResString(10947, CStr(lngCrdTemplateIndex), CStr(lngQtyCrdTemplate)), , 1)
    End If
    
    For i = 1 To udCardQuantity.value
        lngNewCardID = funNewCardNumber
        If optCrdCreateEnabled.value Then
            Call grstCards.AddNew
            grstCards!ID = lngNewCardID
            grstCards!Active = True
            grstCards!date_creation = dteNewCardDate
            grstCards!Comment = txtCrdComment.Text
            grstCards!Type = IIf(cmbCrdType.ItemData(cmbCrdType.ListIndex) = 1, "", cmbCrdType.List(cmbCrdType.ListIndex))
            Call grstCards.Update
            Call subDBAddEvent(enuEvents.CardNew, 0, lngNewCardID, "", 0, funLoadResString(10341), All)
            If optRecharge.value Then
                uNewCard.ID = lngNewCardID
                uNewCard.DateCreation = dteNewCardDate
                uNewCard.TimeCredit = mdteSelectedRechargeTime
                uNewCard.UnlimitedPlay = mblnSelectedRechargeUnlimited
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
                uNewCard.DateExpiry = dteExpiryDateNew
                Call subCardRecharge(uNewCard, msngSelectedRechargePrice, mdteSelectedRechargeTime, funLoadResString(10289) & " " & funFormatTime(mdteSelectedRechargeTime) & " " & mstrSelectedRechargeLabel, funCardInUse(uNewCard))
            End If
        Else
            Call grstCardsPending.AddNew
            grstCardsPending!ID = lngNewCardID
            grstCardsPending!time_credit = mdteSelectedRechargeTime
            grstCardsPending!unlimited_play = mblnSelectedRechargeUnlimited
            grstCardsPending!Price = msngSelectedRechargePrice
            grstCardsPending!validity = mlngSelectedRechargeValidity
            grstCardsPending!validity_type = mintSelectedRechargeValidityType
            grstCardsPending!Caption = mstrSelectedRechargeLabel
            grstCardsPending!Comment = txtCrdComment.Text
            grstCardsPending!Type = IIf(cmbCrdType.ItemData(cmbCrdType.ListIndex) = 1, "", cmbCrdType.List(cmbCrdType.ListIndex))
            Call grstCardsPending.Update
        End If
        If optCrdNumList.value Then
            Print #intOutputFileNumber, CStr(lngNewCardID)
        Else
            lngCrdNumIndex = InStr(1, strOutputFile, CARD_GEN_ID)
            If lngCrdNumIndex > 0 Then
                strOutputFile = Replace(strOutputFile, CARD_GEN_ID, lngNewCardID, , 1)
            Else
                intOutputFileNumber = FreeFile
                Open strFilePath & format(dteNewCardTimeStamp, "yymmdd-HhNnSs") & "_" & funLoadResString(10940) & "_" & format(lngCrdTemplateIndex, "0#") & "-" & format(lngQtyCrdTemplate, "0#") & "." & strFileSuffix For Binary Access Write As intOutputFileNumber
                Put intOutputFileNumber, , strOutputFile
                Close intOutputFileNumber
                strOutputFile = strTemplateFile
                lngCrdTemplateIndex = lngCrdTemplateIndex + 1
                strOutputFile = Replace(strOutputFile, CARD_GEN_TITLE, funLoadResString(10946, format(dteNewCardTimeStamp, "yyyy.mm.dd-Hh:Nn:Ss")), , 1)
                strOutputFile = Replace(strOutputFile, CARD_GEN_HEADER, funLoadResString(10941, format(dteNewCardTimeStamp, "yyyy.mm.dd-Hh:Nn:Ss"), udCardQuantity.value & " " & funLoadResString(IIf(optCrdCreateEnabled.value, 10943, 10944)), strRechargeTime, funFormatCurrencyString(msngSelectedRechargePrice, guServOpts.uConfig.strMonFormat, guServOpts.uConfig.strMonSymbol)) & " - " & funLoadResString(10947, CStr(lngCrdTemplateIndex), CStr(lngQtyCrdTemplate)), , 1)
                strOutputFile = Replace(strOutputFile, CARD_GEN_FOOTER, funLoadResString(10941, format(dteNewCardTimeStamp, "yyyy.mm.dd-Hh:Nn:Ss"), udCardQuantity.value & " " & funLoadResString(IIf(optCrdCreateEnabled.value, 10943, 10944)), strRechargeTime, funFormatCurrencyString(msngSelectedRechargePrice, guServOpts.uConfig.strMonFormat, guServOpts.uConfig.strMonSymbol)) & " - " & funLoadResString(10947, CStr(lngCrdTemplateIndex), CStr(lngQtyCrdTemplate)), , 1)
                strOutputFile = Replace(strOutputFile, CARD_GEN_ID, lngNewCardID, , 1)
            
            End If
        End If
        DoEvents
    Next i
    If optCrdNumList.value Then
        Print #intOutputFileNumber, funLoadResString(10942)
    Else
        intOutputFileNumber = FreeFile
        Open strFilePath & format(dteNewCardTimeStamp, "yymmdd-HhNnSs") & "_" & funLoadResString(10940) & "_" & format(lngCrdTemplateIndex, "0#") & "-" & format(lngQtyCrdTemplate, "0#") & "." & strFileSuffix For Binary Access Write As intOutputFileNumber
        Put intOutputFileNumber, , strOutputFile
    End If
    
    Close intOutputFileNumber
    
    If optCrdCreateEnabled.value Then
        Call subFillCards
    Else
        Call subResizeCards
        Call subFillCardsPending
    End If
    
    Call SetRegUserSettings(funLoadResString(10006), "Card Generation", "Card Pending", IIf(optCrdCreatePending.value, vbChecked, vbUnchecked), REG_DWORD)
    Call SetRegUserSettings(funLoadResString(10006), "Card Generation", "Card Recharged", IIf(optRecharge.value, vbChecked, vbUnchecked), REG_DWORD)
    Call SetRegUserSettings(funLoadResString(10006), "Card Generation", "Recharge Index", mlngSelectedRechargeIndex)
    Call SetRegUserSettings(funLoadResString(10006), "Card Generation", "Card Quantity", txtCardQuantity.Text)
    Call SetRegUserSettings(funLoadResString(10006), "Card Generation", "Output Template", IIf(optCrdNumTemplate.value, vbChecked, vbUnchecked), REG_DWORD)
    Call SetRegUserSettings(funLoadResString(10006), "Card Generation", "Card Template Path", txtCrdTemplatePath.Text)
    Call SetRegUserSettings(funLoadResString(10006), "Card Generation", "Card List Path", txtCrdListPath.Text)
    Call SetRegUserSettings(funLoadResString(10006), "Card Generation", "Card Comment", txtCrdComment.Text)
    Call SetRegUserSettings(funLoadResString(10006), "Card Generation", "Card Type", IIf(cmbCrdType.ListIndex = 0, "", cmbCrdType.List(cmbCrdType.ListIndex)))
    Screen.MousePointer = vbNormal
    
    If funMsg(AskCanViewGenCards, txtCardQuantity.Text) = vbOK Then
        Call funStartDoc(strFilePath)
    End If
    Call Unload(Me)
    Exit Sub

cmdGenerate_ClickError:
    Screen.MousePointer = vbNormal
    Call funErrorManager("frmCardMassAdd.cmdGenerate_ClickError", True)
End Sub


'*******************************************************************************
Private Sub Form_Activate()
On Error GoTo Form_ActivateError

    Call subLocalizeInterface
    Exit Sub

Form_ActivateError:
    Call funErrorManager("frmCardMassAdd.Form_ActivateError", True)
End Sub


'*******************************************************************************
Private Sub Form_Load()
Dim i As Integer
On Error GoTo Form_LoadError
    
    gblnfrmCardMassAdd = True
    Call subSetDefaultFont(Me)
    With fxgRecharges
        .left = shp(3).left + 1 * Screen.TwipsPerPixelX
        .width = shp(3).width - 2 * Screen.TwipsPerPixelX
        .Cols = 4
        .top = shp(3).top + 1 * Screen.TwipsPerPixelY
        .height = shp(3).height - 2 * Screen.TwipsPerPixelY
        
        .Row = 0
        .Col = 0
        .ColWidth(0) = 800
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
        
        Call subFillRecharges
    End With
    optCrdCreateEnabled.value = vbUnchecked = GetRegUserSettings(funLoadResString(10006), "Card Generation", "Card Pending", vbUnchecked, REG_DWORD)
    optCrdCreatePending.value = vbChecked = GetRegUserSettings(funLoadResString(10006), "Card Generation", "Card Pending", vbChecked, REG_DWORD)
    
    optNoRecharge.value = vbUnchecked = GetRegUserSettings(funLoadResString(10006), "Card Generation", "Card Recharged", vbUnchecked, REG_DWORD)
    optRecharge.value = vbChecked = GetRegUserSettings(funLoadResString(10006), "Card Generation", "Card Recharged", vbChecked, REG_DWORD)
    
    mlngSelectedRechargeIndex = GetRegUserSettings(funLoadResString(10006), "Card Generation", "Recharge Index", 0)
    If mlngSelectedRechargeIndex > 0 And optRecharge.value Then
        fxgRecharges.Row = mlngSelectedRechargeIndex
        Call fxgRecharges_RowColChange
    End If
    
    udCardQuantity.value = GetRegUserSettings(funLoadResString(10006), "Card Generation", "Card Quantity", 1)
    
    optCrdNumList.value = vbUnchecked = GetRegUserSettings(funLoadResString(10006), "Card Generation", "Output Template", vbUnchecked, REG_DWORD)
    optCrdNumTemplate.value = vbChecked = GetRegUserSettings(funLoadResString(10006), "Card Generation", "Output Template", vbChecked, REG_DWORD)
    
    txtCrdListPath.Text = GetRegUserSettings(funLoadResString(10006), "Card Generation", "Card List Path", guServOpts.uConfig.strLastExportPath)
    txtCrdListPath.SelStart = Len(txtCrdListPath.Text)

    txtCrdTemplatePath.Text = GetRegUserSettings(funLoadResString(10006), "Card Generation", "Card Template Path", guServOpts.uConfig.strAppConfigPath & FILENAME_CARD_TEMPLATE_FOLDER & FILENAME_CARD_TEMPLATE)
    txtCrdTemplatePath.SelStart = Len(txtCrdTemplatePath.Text)
    
    txtCrdComment.Text = GetRegUserSettings(funLoadResString(10006), "Card Generation", "Card Comment", "")
    txtCrdComment.SelStart = Len(txtCrdComment.Text)
    
    Call cmbCrdType.Clear
    Call cmbCrdType.AddItem(funLoadResString(10526))
    cmbCrdType.ItemData(0) = 1
    i = funArraySize(guServOpts.uCards.strCardTypes)
    If i > 0 Then
        For i = 0 To funArraySize(guServOpts.uCards.strCardTypes) - 1
            Call cmbCrdType.AddItem(guServOpts.uCards.strCardTypes(i))
        Next i
    End If
    cmbCrdType.ListIndex = 0
    For i = 0 To cmbCrdType.ListCount - 1
        If cmbCrdType.List(i) = GetRegUserSettings(funLoadResString(10006), "Card Generation", "Card Type", "") Then
            cmbCrdType.ListIndex = i
            Exit For
        End If
    Next i
    Exit Sub

Form_LoadError:
    Call funErrorManager("frmCardMassAdd.Form_LoadError", True)
End Sub


'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)

    gblnfrmCardMassAdd = False
End Sub


'*******************************************************************************
Private Sub subFillRecharges()
Dim i As Long
On Error GoTo subFillRechargesError

    Call subQueryRecharges(True, False, False, False)
    With fxgRecharges
        If grstRecharges.BOF = True And grstRecharges.EOF = True Then
            .Rows = 21
            .enabled = False
            Exit Sub
        End If
        .Font = ""
        Call grstRechargesQry.MoveLast
        Call grstRechargesQry.MoveFirst
        .Rows = grstRechargesQry.RecordCount + 1
        ReDim msngRechargePrices(1 To grstRechargesQry.RecordCount)
        ReDim mdteRechargeTimes(1 To grstRechargesQry.RecordCount)
        ReDim mblnRechargeUnlimited(1 To grstRechargesQry.RecordCount)
        ReDim mlngRechargeValidities(1 To grstRechargesQry.RecordCount)
        ReDim mintRechargeValidityTypes(1 To grstRechargesQry.RecordCount)
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
            mintRechargeValidityTypes(i) = grstRechargesQry!validity_type
            Call grstRechargesQry.MoveNext
        Next i
    End With
    Exit Sub
    
subFillRechargesError:
    Call funErrorManager("frmCardMassAdd.subFillRechargesError")
End Sub


Private Sub fxgRecharges_Click()
On Error GoTo fxgRecharges_ClickError
    
    If mlngSelectedRechargeIndex = 0 Then Call fxgRecharges_RowColChange
    Exit Sub
    
fxgRecharges_ClickError:
    Call funErrorManager("frmCardMassAdd.fxgRecharges_ClickError", True)
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
        
        If .TopRow > lngNewSelectedRowIndex Then
            .TopRow = lngNewSelectedRowIndex
        ElseIf .TopRow + Int(.height / .RowHeight(0) - 2) < lngNewSelectedRowIndex Then  '- 1
            .TopRow = lngNewSelectedRowIndex - Int(.height / .RowHeight(0) - 2) '+ 1
        End If
        
        mdteSelectedRechargeTime = mdteRechargeTimes(lngNewSelectedRowIndex)
        mblnSelectedRechargeUnlimited = mblnRechargeUnlimited(lngNewSelectedRowIndex)
        msngSelectedRechargePrice = msngRechargePrices(lngNewSelectedRowIndex)
        mlngSelectedRechargeValidity = mlngRechargeValidities(lngNewSelectedRowIndex)
        mintSelectedRechargeValidityType = mintRechargeValidityTypes(lngNewSelectedRowIndex)
        mstrSelectedRechargeLabel = CStr(.TextMatrix(lngNewSelectedRowIndex, 3))
        cmdGenerate.enabled = mlngSelectedRechargeIndex > 0
    End With
    Exit Sub
    
fxgRecharges_RowColChangeError:
    Call funErrorManager("frmCardMassAdd.fxgRecharges_RowColChangeError", True)
End Sub

'*******************************************************************************
Private Sub optCrdNumList_Click()
On Error Resume Next

    txtCrdTemplatePath.enabled = False
    cmdCrdTemplatePath.enabled = False
    txtCrdListPath.enabled = True
    cmdCrdListPath.enabled = True
End Sub


'*******************************************************************************
Private Sub optCrdNumTemplate_Click()
On Error Resume Next

    txtCrdTemplatePath.enabled = True
    cmdCrdTemplatePath.enabled = True
    txtCrdListPath.enabled = False
    cmdCrdListPath.enabled = False
End Sub


'*******************************************************************************
Private Sub optNoRecharge_Click()
On Error Resume Next

    fxgRecharges.enabled = False
    If mlngSelectedRechargeIndex > 0 Then
        fxgRecharges.Row = mlngSelectedRechargeIndex
        fxgRecharges.ColSel = fxgRecharges.Cols - 1
        fxgRecharges.CellForeColor = vbWindowText
        fxgRecharges.CellBackColor = vbWindowBackground
    End If
    mdteSelectedRechargeTime = NULL_DATE
    msngSelectedRechargePrice = 0
    mlngSelectedRechargeValidity = 0
    mintSelectedRechargeValidityType = 0
    mstrSelectedRechargeLabel = ""
    cmdGenerate.enabled = True
    mlngSelectedRechargeIndex = 0
End Sub


'*******************************************************************************
Private Sub optRecharge_Click()
On Error Resume Next

    fxgRecharges.enabled = True
    cmdGenerate.enabled = mlngSelectedRechargeIndex > 0
End Sub


'*******************************************************************************
Private Sub txtCardQuantity_Change()
On Error Resume Next
    
'    If CInt(txtCardQuantity.Text) >= udCardQuantity.Min And CInt(txtCardQuantity.Text) <= udCardQuantity.Max Then
'        udCardQuantity.value = CInt(txtCardQuantity.Text)
'    Else
''        txtCardQuantity.Text = udCardQuantity.value
'    End If
End Sub


'*******************************************************************************
Private Sub txtCardQuantity_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    
    If txtCardQuantity.Text <> udCardQuantity.value And CInt(txtCardQuantity.Text) >= udCardQuantity.Min And CInt(txtCardQuantity.Text) <= udCardQuantity.Max Then
        udCardQuantity.value = CInt(txtCardQuantity.Text)
    Else
'        txtCardQuantity.Text = udCardQuantity.value
    End If
    
    Select Case KeyCode
    Case 40
        If udCardQuantity.value > udCardQuantity.Min Then udCardQuantity.value = udCardQuantity.value - udCardQuantity.Increment
    Case 38
        If udCardQuantity.value < udCardQuantity.Max Then udCardQuantity.value = udCardQuantity.value + udCardQuantity.Increment
    End Select
End Sub


'*******************************************************************************
Private Sub txtCardQuantity_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub


'*******************************************************************************
Private Sub txtCardQuantity_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    Select Case KeyCode
    Case 38, 40
        txtCardQuantity.SelStart = 0
        txtCardQuantity.SelLength = Len(txtCardQuantity.Text)
    End Select
End Sub


'*******************************************************************************
Private Sub txtCardQuantity_LostFocus()
On Error Resume Next
    
    If CInt(txtCardQuantity.Text) >= udCardQuantity.Min And CInt(txtCardQuantity.Text) <= udCardQuantity.Max Then
        udCardQuantity.value = CInt(txtCardQuantity.Text)
    ElseIf CInt(txtCardQuantity.Text) < udCardQuantity.Min Then
        txtCardQuantity.Text = udCardQuantity.Min
    ElseIf CInt(txtCardQuantity.Text) > udCardQuantity.Max Then
        txtCardQuantity.Text = udCardQuantity.Max
    End If
End Sub


'*******************************************************************************
Private Sub txtCrdListPath_GotFocus()
On Error Resume Next

    txtCrdListPath.SelStart = 0
    txtCrdListPath.SelLength = Len(txtCrdListPath.Text)
End Sub

'*******************************************************************************
Private Sub txtCrdTemplatePath_GotFocus()
On Error Resume Next

    txtCrdTemplatePath.SelStart = 0
    txtCrdTemplatePath.SelLength = Len(txtCrdTemplatePath.Text)
End Sub


'*******************************************************************************
Private Sub udCardQuantity_Change()
On Error Resume Next

    txtCardQuantity.Text = udCardQuantity.value
End Sub


'*******************************************************************************
Private Sub udCardQuantity_DownClick()
On Error Resume Next
    
    Call funSetFocus(txtCardQuantity)
    txtCardQuantity.SelStart = 0
    txtCardQuantity.SelLength = Len(txtCardQuantity.Text)
End Sub


'*******************************************************************************
Private Sub udCardQuantity_UpClick()
On Error Resume Next

    Call funSetFocus(txtCardQuantity)
    txtCardQuantity.SelStart = 0
    txtCardQuantity.SelLength = Len(txtCardQuantity.Text)
End Sub
