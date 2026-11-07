VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDBMaintenance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maintenance de base de données"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4035
   ControlBox      =   0   'False
   Icon            =   "frmDBMaintenance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   321
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   269
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkIncludeIncome 
      Caption         =   "_"
      Enabled         =   0   'False
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   720
      Width           =   3675
   End
   Begin VB.ListBox lstRestore 
      Enabled         =   0   'False
      Height          =   1125
      IntegralHeight  =   0   'False
      ItemData        =   "frmDBMaintenance.frx":0442
      Left            =   180
      List            =   "frmDBMaintenance.frx":0444
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   1740
      Width           =   3675
   End
   Begin VB.OptionButton optRestore 
      BackColor       =   &H80000014&
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   1200
      Width           =   3915
   End
   Begin VB.TextBox txtPassword 
      BorderStyle     =   0  'None
      Height          =   195
      IMEMode         =   3  'DISABLE
      Left            =   2220
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   4020
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "_"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   4380
      Width           =   1335
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "_"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   60
      TabIndex        =   4
      Top             =   4380
      Width           =   2535
   End
   Begin MSComCtl2.UpDown udPurgeCutoff 
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   420
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   450
      _Version        =   393216
      Value           =   6
      BuddyControl    =   "lblPurgeCutoff"
      BuddyDispid     =   196619
      OrigLeft        =   3360
      OrigTop         =   420
      OrigRight       =   3600
      OrigBottom      =   675
      Max             =   48
      Min             =   1
      SyncBuddy       =   -1  'True
      Wrap            =   -1  'True
      BuddyProperty   =   65537
      Enabled         =   0   'False
   End
   Begin VB.OptionButton optPurge 
      BackColor       =   &H80000014&
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3915
   End
   Begin VB.Shape shp 
      BorderColor     =   &H80000010&
      Height          =   1035
      Index           =   0
      Left            =   60
      Top             =   60
      Width           =   3915
   End
   Begin VB.Shape shp 
      BorderColor     =   &H80000010&
      Height          =   2235
      Index           =   2
      Left            =   60
      Top             =   1200
      Width           =   3915
   End
   Begin VB.Label lblLastNetworkBackup 
      Caption         =   "_"
      Height          =   435
      Left            =   180
      TabIndex        =   13
      Top             =   2940
      Width           =   3675
   End
   Begin VB.Label lbl 
      Caption         =   "_"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   12
      Top             =   1500
      Width           =   3675
   End
   Begin VB.Label lbl 
      Caption         =   "_"
      Height          =   435
      Index           =   4
      Left            =   60
      TabIndex        =   11
      Top             =   3480
      Width           =   3915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
      Height          =   195
      Index           =   3
      Left            =   900
      TabIndex        =   10
      Top             =   4020
      Width           =   1215
   End
   Begin VB.Shape shp 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   315
      Index           =   1
      Left            =   2160
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "_"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   450
      Width           =   2535
   End
   Begin VB.Label lblPurgeCutoff 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "6"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2700
      TabIndex        =   8
      Top             =   420
      Width           =   255
   End
   Begin VB.Label lbl 
      Caption         =   "_"
      Height          =   195
      Index           =   1
      Left            =   3300
      TabIndex        =   7
      Top             =   450
      Width           =   615
   End
   Begin VB.Shape shp 
      BackColor       =   &H80000014&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   285
      Index           =   3
      Left            =   60
      Top             =   60
      Width           =   3915
   End
   Begin VB.Shape shp 
      BackColor       =   &H80000014&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   285
      Index           =   4
      Left            =   60
      Top             =   1200
      Width           =   3915
   End
End
Attribute VB_Name = "frmDBMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ DATABASE MANTENANCE FORM ___________________|
'|  Name:          frmDBMaintenance.frm                                        |
'|  Type:          VB6 Form                                                    |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
'|  Description:   Database maitenance operations (purge etc)                  |
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

' $Id: frmDBMaintenance.frm 15 2006-01-01 16:12:35Z sam $

Option Explicit


'*******************************************************************************
Private Sub cmdApply_Click()
Dim dtePurgeLimit As Date
On Error GoTo cmdApply_ClickError

    
    If txtPassword.Text = guServOpts.uConfig.strPasswd Then
        dtePurgeLimit = DateAdd("d", -udPurgeCutoff.value * guServOpts.uConfig.intNumDaysInMonth, CDate(DateValue(Now) & " 00:00:00"))
        If optPurge.value Then
            If funMsg(enuMsg.AskCanPurgeHistory, format(dtePurgeLimit, "Long Date")) = vbYes Then
                Screen.MousePointer = vbHourglass
                cmdApply.enabled = False
                cmdCancel.enabled = False
                optPurge.enabled = False
                optRestore.enabled = False
                udPurgeCutoff.enabled = False
                Call subDBMaintenance(dtePurgeLimit, chkIncludeIncome.value = vbChecked)
                Screen.MousePointer = vbNormal
                guServOpts.uDB.intDBCompact = vbChecked
                Call Unload(Me)
            End If
        ElseIf optRestore.value Then
            If lstRestore.ListIndex = -1 Then
                Call funMsg(enuMsg.NoElementSelected)
            ElseIf funMsg(enuMsg.AskCanRestoreDB, lstRestore.List(lstRestore.ListIndex)) = vbYes Then
                Call SetRegSettings(funLoadResString(10006), "DB", "Restore backup", lstRestore.ItemData(lstRestore.ListIndex))
                guServOpts.uDB.intDBCompact = vbUnchecked
                guServOpts.uDB.intDBBackup = vbUnchecked
                Call Unload(Me)
                Call Unload(frmMain)
            End If
        End If
    Else
        Call funMsg(enuMsg.BadPasswd)
        txtPassword.Text = ""
        Call funSetFocus(txtPassword)
    End If
    Exit Sub
    
cmdApply_ClickError:
    Call funErrorManager("frmDBMaintenance.cmdApply_ClickError", True)
    cmdCancel.enabled = True
End Sub


'*******************************************************************************
Private Sub cmdCancel_Click()
On Error GoTo cmdCancel_ClickError

    Call Unload(Me)
    Exit Sub

cmdCancel_ClickError:
    Call funErrorManager("frmDBMaintenance.cmdCancel_ClickError", True)
End Sub


'*******************************************************************************
Private Sub Form_Activate()
On Error Resume Next

    Call funSetFocus(optPurge)
    lstRestore.ListIndex = -1
    optPurge.value = False
    optRestore.value = False
    cmdApply.enabled = False
    txtPassword.enabled = False
End Sub


'*******************************************************************************
Private Sub Form_Load()
Dim i As Integer
Dim dteFileDate As Date
Dim strLastNetBackupLocation As String
Dim strDBName As String
On Error GoTo Form_LoadError

    Call subSetDefaultFont(Me)
    Call subLocalizeInterface
    optPurge.top = optPurge.top + 1
    optPurge.left = optPurge.left + 2
    optPurge.width = optPurge.width - 3
    
    optRestore.top = optRestore.top + 1
    optRestore.left = optRestore.left + 2
    optRestore.width = optRestore.width - 3
    
    strDBName = right(guServOpts.uDB.strDBPath, Len(guServOpts.uDB.strDBPath) - InStrRev(guServOpts.uDB.strDBPath, "\"))
    For i = 1 To guServOpts.uDB.intDBMaxBackups
        If Dir(guServOpts.uDB.strDBBackupPath & "\" & strDBName & ".bak" & i) <> "" Then
            dteFileDate = FileDateTime(guServOpts.uDB.strDBBackupPath & "\" & strDBName & ".bak" & i)
            Call lstRestore.AddItem(format(dteFileDate, "Long Date") & funLoadResString(10722) & format(dteFileDate, "Long Time")) ', i - 1)
            lstRestore.ItemData(lstRestore.NewIndex) = i
        End If
    Next i
    strLastNetBackupLocation = GetRegSettings(funLoadResString(10006), "DB", "Last workstation location", "")
    If strLastNetBackupLocation = "" Then
        lblLastNetworkBackup.Caption = funLoadResString(10745)
    Else
        lblLastNetworkBackup.Caption = funLoadResString(10744) & strLastNetBackupLocation & "."
        lblLastNetworkBackup.ToolTipText = funLoadResString(10746, strLastNetBackupLocation) & funGetFolderPath(CSIDL_COMMON_APPDATA) & "\" & funLoadResString(10747) & FILENAME_DB_BACKUP & "."
    End If
    Exit Sub

Form_LoadError:
    Call funErrorManager("frmDBMaintenance.Form_LoadError", True)
End Sub


'*******************************************************************************
Public Sub subLocalizeInterface()
On Error GoTo subLocalizeInterfaceError

    Call subSetCharset(Me, guServOpts.uConfig.intInterfaceLangCharset)
    
    Me.Caption = funLoadResString(10509)
    optPurge.Caption = funLoadResString(10510)
    optRestore.Caption = funLoadResString(10727)
    lbl(0).Caption = funLoadResString(10728)
    lbl(1).Caption = funLoadResString(10511)
    lbl(2).Caption = funLoadResString(10512)
    lbl(3).Caption = funLoadResString(10116)
    lbl(4).Caption = funLoadResString(10520)
    cmdApply.Caption = funLoadResString(10513)
    cmdCancel.Caption = funLoadResString(10123)
    chkIncludeIncome.Caption = funLoadResString(10789)
    Exit Sub

subLocalizeInterfaceError:
    Call funErrorManager("frmDBMaintenance.subLocalizeInterfaceError", True)
End Sub


'*******************************************************************************
Private Sub lstRestore_Click()
Dim i As Integer
On Error GoTo lstRestore_ClickError
    
    For i = 0 To lstRestore.ListCount - 1
        lstRestore.Selected(i) = i = lstRestore.ListIndex
    Next i
    Exit Sub

lstRestore_ClickError:
    Call funErrorManager("frmDBMaintenance.lstRestore_ClickError", True)
End Sub


'*******************************************************************************
Private Sub optPurge_Click()
On Error Resume Next

    lstRestore.enabled = False
    udPurgeCutoff.enabled = True
    lblPurgeCutoff.enabled = True
    cmdApply.enabled = True
    txtPassword.enabled = True
    chkIncludeIncome.enabled = True
End Sub


'*******************************************************************************
Private Sub optRestore_Click()
On Error Resume Next

    lstRestore.enabled = True
    udPurgeCutoff.enabled = False
    lblPurgeCutoff.enabled = False
    cmdApply.enabled = True
    txtPassword.enabled = True
    chkIncludeIncome.enabled = False
End Sub


'*******************************************************************************
Private Sub txtPassword_GotFocus()
On Error Resume Next
    
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
End Sub
