VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl ucWorkstation 
   ClientHeight    =   1440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5055
   ScaleHeight     =   1440
   ScaleWidth      =   5055
   Begin MSComctlLib.ImageList imlStatusIcons 
      Left            =   2760
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucClient.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucClient.ctx":031A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucClient.ctx":0634
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucClient.ctx":094E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrStatus 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   60
      Top             =   660
   End
   Begin MSComctlLib.ImageList imlStateIcons 
      Left            =   3300
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucClient.ctx":0C68
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucClient.ctx":0F82
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucClient.ctx":129C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSessionTypeIcons 
      Left            =   3900
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucClient.ctx":15B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucClient.ctx":18D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucClient.ctx":1BEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucClient.ctx":1F04
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucClient.ctx":221E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucClient.ctx":2538
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblState 
      BackStyle       =   0  'Transparent
      Caption         =   "Client State"
      ForeColor       =   &H80000011&
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   300
      Width           =   4395
   End
   Begin VB.Image imgStatus 
      Height          =   480
      Left            =   60
      Top             =   60
      Width           =   480
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H80000010&
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   2355
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Client Name"
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
      Left            =   600
      TabIndex        =   1
      Top             =   60
      Width           =   4335
   End
   Begin VB.Image imgSessionType 
      Height          =   480
      Left            =   60
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgWorkstations 
      Height          =   480
      Left            =   60
      Picture         =   "ucClient.ctx":2852
      Top             =   60
      Width           =   480
   End
   Begin VB.Shape shpBack 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   2355
   End
End
Attribute VB_Name = "ucWorkstation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ WORKSTATION USER CONTROL ___________________|
'|  Name:          ucWorkstation.ctl                                           |
'|  Type:          VB6 User Control                                            |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
'|  Description:   Workstation User Control                                    |
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

' $Id: ucClient.ctl 79 2005-10-19 21:36:57Z sam $

Option Explicit

'Default Property Values:
Const m_def_WorkstationName = ""
Const m_def_StateText = ""
Const m_def_Active = 0
Const m_def_State = 0
Const m_def_Status = 1
Const m_def_Size = 1
Const m_def_SessionType = 1
Const m_def_ConnectionIndex = "Key: 0"
Const m_def_WorkstationIndex = 0

'Property Variables:
Dim m_WorkstationName As String
Dim m_StateText As String
Dim m_Active As Boolean
Dim m_State As Integer
Dim m_Status As Integer
Dim m_StatusGUI As Integer
Dim m_Size As Integer
Dim m_SessionType As Integer
Dim m_ConnectionIndex As String
Dim m_WorkstationIndex As Integer

'Event Declarations:
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event RightClick(Shift As Integer, x As Single, y As Single)

Public Enum enuState
    cliNotPresent = 0
    cliNonActive = 1
    cliReady = 2
    cliBlocked_Deactivated = 3
    cliBlocked_SessionEnd = 4
    cliUsedDisconnected = 5
    cliUsedConnected = 6
    cliUsedFree = 7
    cliBlocked_NoCredit = 8
    cliBlocked_Paused = 9
    cliBlocked_Locked = 10
    cliBlocked_follow = 11
End Enum


Public Enum enuStatus
    cliNormal = 1
    cliInfo = 2
    cliWarn = 3
    cliCritical = 4
End Enum

Public Enum enuCliSize
    cliSizeNormal = 1
    cliSizeSmall = 2
End Enum

Public Enum enuWorkstationSessionType
    cliSelfservice = 1
    cliCard = 2
    cliCardWarn = 3
    cliFree = 4
    cliFollow = 5
    cliClient = 6
End Enum


'*******************************************************************************
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
On Error GoTo CaptionError
    
    Caption = m_WorkstationName
    Exit Property

CaptionError:
    Call funErrorManager("ucWorkstation.CaptionError", True)
End Property


'*******************************************************************************
Public Property Let Caption(ByVal New_Caption As String)
On Error GoTo CaptionError
    
    m_WorkstationName = New_Caption
    If m_Size = enuCliSize.cliSizeNormal Then
        lblName.Caption = m_WorkstationName
    Else
        lblName.Caption = IIf(m_WorkstationName <> "", funLoadResString(10886, CStr(m_WorkstationIndex)), "")
    End If
    PropertyChanged "Caption"
    Exit Property

CaptionError:
    Call funErrorManager("ucWorkstation.CaptionError", True)
End Property


'*******************************************************************************
Private Sub imgWorkstations_Click()
On Error GoTo imgWorkstations_ClickError
    
    Call UserControl_Click
    Exit Sub

imgWorkstations_ClickError:
    Call funErrorManager("ucWorkstation.imgWorkstations_ClickError", True)
End Sub


'*******************************************************************************
Private Sub imgCrit_Click()
On Error GoTo imgCrit_ClickError
    
    Call UserControl_Click
    Exit Sub

imgCrit_ClickError:
    Call funErrorManager("ucWorkstation.imgCrit_ClickError", True)
End Sub


'*******************************************************************************
Private Sub imgWarn_Click()
On Error GoTo imgWarn_ClickError
    
    Call UserControl_Click
    Exit Sub

imgWarn_ClickError:
    Call funErrorManager("ucWorkstation.imgWarn_ClickError", True)
End Sub


'*******************************************************************************
Private Sub imgWorkstations_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo imgWorkstations_MouseDownError
    
    Call UserControl_MouseDown(Button, Shift, imgWorkstations.left + x, imgWorkstations.top + y)
    Exit Sub

imgWorkstations_MouseDownError:
    Call funErrorManager("ucWorkstation.imgWorkstations_MouseDownError", True)
End Sub

'*******************************************************************************
Private Sub imgStatus_Click()
On Error GoTo imgStatus_ClickError
    
    Call UserControl_Click
    Exit Sub

imgStatus_ClickError:
    Call funErrorManager("ucWorkstation.imgStatus_ClickError", True)
End Sub


'*******************************************************************************
Private Sub imgStatus_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo imgStatus_MouseDownError
    
    Call UserControl_MouseDown(Button, Shift, imgStatus.left + x, imgStatus.top + y)
    Exit Sub

imgStatus_MouseDownError:
    Call funErrorManager("ucWorkstation.imgStatus_MouseDownError", True)
End Sub


'*******************************************************************************
Private Sub lblName_Click()
On Error GoTo lblName_ClickError
    
    Call UserControl_Click
    Exit Sub

lblName_ClickError:
    Call funErrorManager("ucWorkstation.lblName_ClickError", True)
End Sub


'*******************************************************************************
Private Sub lblName_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo lblName_MouseDownError
    
    Call UserControl_MouseDown(Button, Shift, lblName.left + x, lblName.top + y)
    Exit Sub

lblName_MouseDownError:
    Call funErrorManager("ucWorkstation.lblName_MouseDownError", True)
End Sub


'*******************************************************************************
Private Sub lblState_Click()
On Error GoTo lblState_ClickError
    
    Call UserControl_Click
    Exit Sub

lblState_ClickError:
    Call funErrorManager("ucWorkstation.lblState_ClickError", True)
End Sub


'*******************************************************************************
Private Sub lblState_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo lblState_MouseDownError
    
    Call UserControl_MouseDown(Button, Shift, lblState.left + x, lblState.top + y)
    Exit Sub

lblState_MouseDownError:
    Call funErrorManager("ucWorkstation.lblState_MouseDownError", True)
End Sub


'*******************************************************************************
Private Sub tmrStatus_Timer()
Static blnStatusBlink As Boolean
On Error GoTo tmrStatus_TimerError

    blnStatusBlink = Not blnStatusBlink
    imgStatus.Visible = blnStatusBlink
    Exit Sub

tmrStatus_TimerError:
    Call funErrorManager("ucWorkstation.tmrStatus_TimerError", True)
End Sub


'*******************************************************************************
Private Sub UserControl_DblClick()
On Error GoTo UserControl_DblClickError
    
    RaiseEvent DblClick
    Exit Sub

UserControl_DblClickError:
    Call funErrorManager("ucWorkstation.UserControl_DblClickError", True)
End Sub


'*******************************************************************************
Private Sub UserControl_Click()
On Error GoTo UserControl_ClickError
    
    RaiseEvent Click
    Exit Sub

UserControl_ClickError:
    Call funErrorManager("ucWorkstation.UserControl_ClickError", True)
End Sub


'*******************************************************************************
Public Property Get State() As enuState
On Error GoTo StateError
    
    State = m_State
    Exit Property

StateError:
    Call funErrorManager("ucWorkstation.StateError", True)
End Property


'*******************************************************************************
Public Property Let State(ByVal New_State As enuState)
Dim blnStateChanged As Boolean
On Error GoTo StateError

    blnStateChanged = m_State <> New_State
    If blnStateChanged Then Call subRemoveStatusBarMsg(m_ConnectionIndex)
    m_State = New_State
    Call subUpdateState(blnStateChanged)
    PropertyChanged "State"
    Exit Property

StateError:
    Call funErrorManager("ucWorkstation.StateError", True)
End Property


'*******************************************************************************
Public Property Get ConnectionIndex() As String
On Error GoTo ConnectionIndexError
    
    ConnectionIndex = m_ConnectionIndex
    Exit Property

ConnectionIndexError:
    Call funErrorManager("ucWorkstation.ConnectionIndexError", True)
End Property


'*******************************************************************************
Public Property Let ConnectionIndex(ByVal New_ConnectionIndex As String)
On Error GoTo ConnectionIndexError
        
    If New_ConnectionIndex <> m_ConnectionIndex Then Call subRemoveStatusBarMsg(m_ConnectionIndex)
    m_ConnectionIndex = New_ConnectionIndex
    PropertyChanged "ConnectionIndex"
    Exit Property

ConnectionIndexError:
    Call funErrorManager("ucWorkstation.ConnectionIndexError", True)
End Property


'*******************************************************************************
Public Property Get Status() As enuStatus
On Error GoTo StatusError
    
    Status = m_Status
    Exit Property

StatusError:
    Call funErrorManager("ucWorkstation.StatusError", True)
End Property


'*******************************************************************************
Public Property Get StatusGUI() As enuStatus
On Error GoTo StatusGUIError
    
    StatusGUI = m_StatusGUI
    Exit Property

StatusGUIError:
    Call funErrorManager("ucWorkstation.StatusGUIError", True)
End Property


'*******************************************************************************
Public Property Let size(ByVal New_Size As enuCliSize)
On Error GoTo SizeError
    
    m_Size = New_Size
    Select Case m_Size
    Case enuCliSize.cliSizeNormal
        shpBorder.width = WKSN_GUI_SIZE_NORMAL
        shpBack.width = WKSN_GUI_SIZE_NORMAL
        lblName.width = WKSN_GUI_SIZE_NORMAL - lblName.left
        lblState.width = WKSN_GUI_SIZE_NORMAL - lblState.left
        lblName.Caption = m_WorkstationName
    
    Case enuCliSize.cliSizeSmall
        shpBorder.width = WKSN_GUI_SIZE_SMALL
        shpBack.width = WKSN_GUI_SIZE_SMALL
        lblName.width = WKSN_GUI_SIZE_SMALL - lblName.left
        lblState.width = WKSN_GUI_SIZE_SMALL - lblState.left
        lblName.Caption = IIf(m_WorkstationName <> "", funLoadResString(10886, CStr(m_WorkstationIndex)), "")
    
    End Select
    Call subUpdateState(False)
    PropertyChanged "Size"
    Exit Property

SizeError:
    Call funErrorManager("ucWorkstation.SizeError", True)
End Property


'*******************************************************************************
Public Property Get size() As enuCliSize
On Error GoTo SizeError
    
    size = m_Size
    Exit Property

SizeError:
    Call funErrorManager("ucWorkstation.SizeError", True)
End Property


'*******************************************************************************
Public Property Let Status(ByVal New_Status As enuStatus)
On Error GoTo StatusError
    
    If New_Status = 0 Then Exit Property
    m_Status = New_Status
    If m_Status > m_StatusGUI Then Me.StatusGUI = m_Status
'    imgStatus.Picture = imlStatusIcons.ListImages(m_Status).Picture
    PropertyChanged "Status"
    Exit Property

StatusError:
    Call funErrorManager("ucWorkstation.StatusError", True)
End Property


'*******************************************************************************
Public Property Let StatusGUI(ByVal New_StatusGUI As enuStatus)
On Error GoTo StatusGUIError
    
    If New_StatusGUI = 0 Then Exit Property
    m_StatusGUI = New_StatusGUI
    If m_StatusGUI >= m_Status Then imgStatus.Picture = imlStatusIcons.ListImages(m_StatusGUI).Picture
    PropertyChanged "StatusGUI"
    Exit Property

StatusGUIError:
    Call funErrorManager("ucWorkstation.StatusGUIError", True)
End Property


'*******************************************************************************
Public Property Get SessionType() As enuWorkstationSessionType
On Error GoTo SessionTypeError
    
    SessionType = m_SessionType
    Exit Property

SessionTypeError:
    Call funErrorManager("ucWorkstation.SessionTypeError", True)
End Property


'*******************************************************************************
Public Property Let SessionType(ByVal New_SessionType As enuWorkstationSessionType)
On Error GoTo SessionTypeError
    
    m_SessionType = New_SessionType
    imgSessionType.Picture = imlSessionTypeIcons.ListImages(m_SessionType).Picture
    PropertyChanged "SessionType"
    Exit Property

SessionTypeError:
    Call funErrorManager("ucWorkstation.SessionTypeError", True)
End Property


'*******************************************************************************
Public Property Get WorkstationIndex() As Integer
Attribute WorkstationIndex.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
On Error GoTo WorkstationIndexError
    
    WorkstationIndex = m_WorkstationIndex
    Exit Property

WorkstationIndexError:
    Call funErrorManager("ucWorkstation.WorkstationIndexError", True)
End Property


'*******************************************************************************
Public Property Let WorkstationIndex(ByVal New_WorkstationIndex As Integer)
On Error GoTo WorkstationIndexError
    
    m_WorkstationIndex = New_WorkstationIndex
    PropertyChanged "WorkstationIndex"
    Exit Property

WorkstationIndexError:
    Call funErrorManager("ucWorkstation.WorkstationIndexError", True)
End Property


'*******************************************************************************
Public Property Get StateText() As String
On Error GoTo StateTextError
    
    StateText = lblState.Caption
    Exit Property

StateTextError:
    Call funErrorManager("ucWorkstation.StateTextError", True)
End Property


'*******************************************************************************
Public Property Let StateText(ByVal New_StateText As String)
On Error GoTo StateTextError
    
    lblState.Caption() = New_StateText
    PropertyChanged "StateText"
    Exit Property

StateTextError:
    Call funErrorManager("ucWorkstation.StateTextError", True)
End Property


Private Sub UserControl_Initialize()
Dim uTmpControl As Control
On Error Resume Next
    
    For Each uTmpControl In Controls
        uTmpControl.Font = ""
    Next uTmpControl
End Sub

'*******************************************************************************
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
On Error GoTo UserControl_InitPropertiesError
    
    
    m_Size = m_def_Size
    Me.size = m_Size
    
    m_Status = m_def_Status
    Me.Status = m_Status
    
    m_State = m_def_State
    Me.State = m_State
    
    m_Active = m_def_Active
    
    m_SessionType = m_def_SessionType
    Me.SessionType = enuWorkstationSessionType.cliSelfservice
    m_StateText = m_def_StateText
        
    Exit Sub

UserControl_InitPropertiesError:
    Call funErrorManager("ucWorkstation.UserControl_InitPropertiesError", True)
End Sub


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo UserControl_MouseDownError
    
    If Button = vbRightButton Then
'        RaiseEvent Click
        RaiseEvent RightClick(Shift, x, y)
    End If
    Exit Sub

UserControl_MouseDownError:
    Call funErrorManager("ucWorkstation.UserControl_MouseDownError", True)
End Sub

'*******************************************************************************
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error GoTo UserControl_ReadPropertiesError

    lblName.Caption = PropBag.ReadProperty("Caption", "Workstation Name")
    m_State = PropBag.ReadProperty("State", m_def_State)
    m_Status = PropBag.ReadProperty("Status", m_def_Status)
    m_StatusGUI = PropBag.ReadProperty("StatusGUI", m_def_Status)
    m_Size = PropBag.ReadProperty("GUISize", m_def_Size)
    lblState.Caption = PropBag.ReadProperty("StateText", "Workstation State")
    tmrStatus.enabled = PropBag.ReadProperty("StatusBlink", False)
    tmrStatus.Interval = PropBag.ReadProperty("StatusBlinkSpeed", 500)
    m_Active = PropBag.ReadProperty("Active", m_def_Active)
    m_ConnectionIndex = PropBag.ReadProperty("ConnectionIndex", m_def_ConnectionIndex)
    m_SessionType = PropBag.ReadProperty("SessionType", m_def_SessionType)
    lblName.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    lblName.ToolTipText = PropBag.ReadProperty("ComputerName", "")
    imgSessionType.ToolTipText = PropBag.ReadProperty("CustomerID", "")
    lblState.ToolTipText = PropBag.ReadProperty("CustomerInfo", "")
    m_StateText = PropBag.ReadProperty("StateText", m_def_StateText)
    m_WorkstationIndex = PropBag.ReadProperty("WorkstationIndex", m_def_WorkstationIndex)
    Exit Sub

UserControl_ReadPropertiesError:
    Call funErrorManager("ucWorkstation.UserControl_ReadPropertiesError", True)
End Sub


'*******************************************************************************
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error GoTo UserControl_WritePropertiesError

    Call PropBag.WriteProperty("Caption", lblName.Caption, "Workstation Name")
    Call PropBag.WriteProperty("State", m_State, m_def_State)
    Call PropBag.WriteProperty("Status", m_Status, m_def_Status)
    Call PropBag.WriteProperty("StatusGUI", m_StatusGUI, m_def_Status)
    Call PropBag.WriteProperty("GUISize", m_Size, m_def_Size)
    Call PropBag.WriteProperty("StateText", lblState.Caption, "Workstation State")
    Call PropBag.WriteProperty("StatusBlink", tmrStatus.enabled, False)
    Call PropBag.WriteProperty("StatusBlinkSpeed", tmrStatus.Interval, 500)
    Call PropBag.WriteProperty("Active", m_Active, m_def_Active)
    Call PropBag.WriteProperty("ConnectionIndex", m_ConnectionIndex, m_def_ConnectionIndex)
    Call PropBag.WriteProperty("SessionType", m_SessionType, m_def_SessionType)
    Call PropBag.WriteProperty("ToolTipText", lblName.ToolTipText, "")
    Call PropBag.WriteProperty("ComputerName", lblName.ToolTipText, "")
    Call PropBag.WriteProperty("CustomerID", imgSessionType.ToolTipText, "")
    Call PropBag.WriteProperty("CustomerInfo", lblState.ToolTipText, "")
    Call PropBag.WriteProperty("StateText", m_StateText, m_def_StateText)
    Call PropBag.WriteProperty("WorkstationIndex", m_WorkstationIndex, m_def_WorkstationIndex)
    Exit Sub

UserControl_WritePropertiesError:
    Call funErrorManager("ucWorkstation.UserControl_WritePropertiesError", True)
End Sub


'*******************************************************************************
Public Property Get StatusBlink() As Boolean
Attribute StatusBlink.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
On Error GoTo StatusBlinkError
    
    StatusBlink = tmrStatus.enabled
    Exit Property

StatusBlinkError:
    Call funErrorManager("ucWorkstation.StatusBlinkError", True)
End Property


'*******************************************************************************
Public Property Let StatusBlink(ByVal New_StatusBlink As Boolean)
On Error GoTo StatusBlinkError
    
    tmrStatus.enabled() = New_StatusBlink
    imgStatus.Visible = True
    PropertyChanged "StatusBlink"
    Exit Property

StatusBlinkError:
    Call funErrorManager("ucWorkstation.StatusBlinkError", True)
End Property


'*******************************************************************************
Public Property Get StatusBlinkSpeed() As Long
Attribute StatusBlinkSpeed.VB_Description = "Returns/sets the number of milliseconds between calls to a Timer control's Timer event."
On Error GoTo StatusBlinkSpeedError
    
    StatusBlinkSpeed = tmrStatus.Interval
    Exit Property

StatusBlinkSpeedError:
    Call funErrorManager("ucWorkstation.StatusBlinkSpeedError", True)
End Property


'*******************************************************************************
Public Property Let StatusBlinkSpeed(ByVal New_StatusBlinkSpeed As Long)
On Error GoTo StatusBlinkSpeedError
    
    tmrStatus.Interval() = New_StatusBlinkSpeed
    PropertyChanged "StatusBlinkSpeed"
    Exit Property

StatusBlinkSpeedError:
    Call funErrorManager("ucWorkstation.StatusBlinkSpeedError", True)
End Property


'*******************************************************************************
Public Property Get Active() As Boolean
On Error GoTo ActiveError
    
    Active = m_Active
    Exit Property

ActiveError:
    Call funErrorManager("ucWorkstation.ActiveError", True)
End Property


'*******************************************************************************
Public Property Let Active(ByVal New_Active As Boolean)
On Error GoTo ActiveError
    
    m_Active = New_Active
    If m_Active Then
        shpBack.BackColor = vb3DHighlight
    Else
        shpBack.BackColor = vb3DFace
    End If
    PropertyChanged "Active"
    Exit Property

ActiveError:
    Call funErrorManager("ucWorkstation.ActiveError", True)
End Property


'*******************************************************************************
Public Property Get ComputerName() As String
On Error GoTo ComputerNameError
    
    ComputerName = lblName.ToolTipText
    Exit Property

ComputerNameError:
    Call funErrorManager("ucWorkstation.ComputerNameError", True)
End Property


'*******************************************************************************
Public Property Let ComputerName(ByVal New_ComputerName As String)
On Error GoTo ComputerNameError
    
    lblName.ToolTipText = New_ComputerName
    PropertyChanged "ComputerName"
    Exit Property

ComputerNameError:
    Call funErrorManager("ucWorkstation.ComputerNameError", True)
End Property


'*******************************************************************************
Public Property Get CustomerID() As String
On Error GoTo CustomerIDError
    
    CustomerID = imgSessionType.ToolTipText
    Exit Property

CustomerIDError:
    Call funErrorManager("ucWorkstation.CustomerIDError", True)
End Property


'*******************************************************************************
Public Property Let CustomerID(ByVal New_CustomerID As String)
On Error GoTo CustomerIDError
    
    imgSessionType.ToolTipText = New_CustomerID
    imgStatus.ToolTipText = New_CustomerID
    imgWorkstations.ToolTipText = New_CustomerID
    PropertyChanged "CustomerID"
    Exit Property

CustomerIDError:
    Call funErrorManager("ucWorkstation.CustomerIDError", True)
End Property


'*******************************************************************************
Public Property Get CustomerInfo() As String
On Error GoTo CustomerInfoError
    
    CustomerInfo = lblState.ToolTipText
    Exit Property

CustomerInfoError:
    Call funErrorManager("ucWorkstation.CustomerInfoError", True)
End Property


'*******************************************************************************
Public Property Let CustomerInfo(ByVal New_CustomerInfo As String)
On Error GoTo CustomerInfoError
    
    lblState.ToolTipText = New_CustomerInfo
    PropertyChanged "CustomerInfo"
    Exit Property

CustomerInfoError:
    Call funErrorManager("ucWorkstation.CustomerInfoError", True)
End Property



'*******************************************************************************
Private Sub subUpdateState(blnStateChanged As Boolean)
Dim uMsg As New clsStatusBarMessage
On Error GoTo subUpdateStateError

    Select Case m_State
    Case enuState.cliNotPresent
        imgWorkstations.Picture = imlStateIcons.ListImages(1).Picture
        Me.StatusGUI = cliCritical
        lblState.Caption = IIf(m_Size = enuCliSize.cliSizeNormal, funLoadResString(10171), funLoadResString(10887))
        lblState.ToolTipText = ""
'        lblState.ToolTipText = IIf(m_Size = enuCliSize.cliSizeNormal, lblState.ToolTipText, funLoadResString(10171))
        Me.SessionType = cliSelfservice
    
    Case enuState.cliNonActive
        imgWorkstations.Picture = imlStateIcons.ListImages(1).Picture
        Me.StatusGUI = cliCritical
        Me.SessionType = cliSelfservice
        lblState.Caption = IIf(m_Size = enuCliSize.cliSizeNormal, funLoadResString(10172), funLoadResString(10888))
        lblState.ToolTipText = ""
'        lblState.ToolTipText = IIf(m_Size = enuCliSize.cliSizeNormal, lblState.ToolTipText, funLoadResString(10172))
        uMsg.ImageIndex = Me.StatusGUI
        uMsg.Text = lblName & " - " & lblState.Caption
        If blnStateChanged Then Call funAddStatusBarMsg(uMsg, m_ConnectionIndex)
        
    Case enuState.cliReady
        imgWorkstations.Picture = imlStateIcons.ListImages(1).Picture
        Me.StatusGUI = cliNormal
        Me.SessionType = cliSelfservice
        lblState.Caption = IIf(m_Size = enuCliSize.cliSizeNormal, funLoadResString(10173), funLoadResString(10889))
        lblState.ToolTipText = ""
'        lblState.ToolTipText = IIf(m_Size = enuCliSize.cliSizeNormal, lblState.ToolTipText, funLoadResString(10173))
    
    Case enuState.cliUsedConnected
        imgWorkstations.Picture = imlStateIcons.ListImages(3).Picture
        Me.StatusGUI = cliNormal
        lblState.Caption = IIf(m_Size = enuCliSize.cliSizeNormal, funLoadResString(10174), funLoadResString(10890))
'        lblState.ToolTipText = IIf(m_Size = enuCliSize.cliSizeNormal, lblState.ToolTipText, funLoadResString(10174))
    
    Case enuState.cliUsedDisconnected
        imgWorkstations.Picture = imlStateIcons.ListImages(2).Picture
        Me.StatusGUI = cliNormal
        lblState.Caption = IIf(m_Size = enuCliSize.cliSizeNormal, funLoadResString(10175), funLoadResString(10891))
'        lblState.ToolTipText = IIf(m_Size = enuCliSize.cliSizeNormal, lblState.ToolTipText, funLoadResString(10175))
    
    Case enuState.cliUsedFree
        imgWorkstations.Picture = imlStateIcons.ListImages(3).Picture
        Me.StatusGUI = cliNormal
        lblState.Caption = IIf(m_Size = enuCliSize.cliSizeNormal, funLoadResString(10620), funLoadResString(10896))
'        lblState.ToolTipText = IIf(m_Size = enuCliSize.cliSizeNormal, lblState.ToolTipText, funLoadResString(10620))
    
    Case enuState.cliBlocked_NoCredit
        imgWorkstations.Picture = imlStateIcons.ListImages(2).Picture
        Me.StatusGUI = cliWarn
        lblState.Caption = IIf(m_Size = enuCliSize.cliSizeNormal, funLoadResString(10176), funLoadResString(10892))
'        lblState.ToolTipText = IIf(m_Size = enuCliSize.cliSizeNormal, lblState.ToolTipText, funLoadResString(10176))
        uMsg.ImageIndex = Me.StatusGUI
        uMsg.Text = lblName & " - " & lblState.Caption
        If blnStateChanged Then Call funAddStatusBarMsg(uMsg, m_ConnectionIndex)
    
    Case enuState.cliBlocked_Paused
        imgWorkstations.Picture = imlStateIcons.ListImages(2).Picture
        Me.StatusGUI = cliWarn
        lblState.Caption = IIf(m_Size = enuCliSize.cliSizeNormal, funLoadResString(10177), funLoadResString(10893))
'        lblState.ToolTipText = IIf(m_Size = enuCliSize.cliSizeNormal, lblState.ToolTipText, funLoadResString(10177))
        uMsg.ImageIndex = Me.StatusGUI
        uMsg.Text = lblName & " - " & lblState.Caption
        If blnStateChanged Then Call funAddStatusBarMsg(uMsg, m_ConnectionIndex)
    
    Case enuState.cliBlocked_Locked
        imgWorkstations.Picture = imlStateIcons.ListImages(2).Picture
        Me.StatusGUI = cliWarn
        lblState.Caption = IIf(m_Size = enuCliSize.cliSizeNormal, funLoadResString(10596), funLoadResString(10897))
'        lblState.ToolTipText = IIf(m_Size = enuCliSize.cliSizeNormal, lblState.ToolTipText, funLoadResString(10596))
        uMsg.ImageIndex = Me.StatusGUI
        uMsg.Text = lblName & " - " & lblState.Caption
        If blnStateChanged Then Call funAddStatusBarMsg(uMsg, m_ConnectionIndex)
    
    Case enuState.cliBlocked_follow
        imgWorkstations.Picture = imlStateIcons.ListImages(2).Picture
        Me.StatusGUI = cliNormal
        lblState.Caption = IIf(m_Size = enuCliSize.cliSizeNormal, funLoadResString(10626), funLoadResString(10898))
'        lblState.ToolTipText = IIf(m_Size = enuCliSize.cliSizeNormal, lblState.ToolTipText, funLoadResString(10626))
    
    Case enuState.cliBlocked_SessionEnd
        imgWorkstations.Picture = imlStateIcons.ListImages(1).Picture
        Me.StatusGUI = cliWarn
        lblState.Caption = IIf(m_Size = enuCliSize.cliSizeNormal, funLoadResString(10178), funLoadResString(10894))
'        lblState.ToolTipText = IIf(m_Size = enuCliSize.cliSizeNormal, lblState.ToolTipText, funLoadResString(10178))
        uMsg.ImageIndex = Me.StatusGUI
        uMsg.Text = lblName & " - " & lblState.Caption
        If blnStateChanged Then Call funAddStatusBarMsg(uMsg, m_ConnectionIndex)
    
    Case enuState.cliBlocked_Deactivated
        imgWorkstations.Picture = imlStateIcons.ListImages(1).Picture
        Me.StatusGUI = cliWarn
        lblState.Caption = IIf(m_Size = enuCliSize.cliSizeNormal, funLoadResString(10179), funLoadResString(10895))
        lblState.ToolTipText = ""
'        lblState.ToolTipText = IIf(m_Size = enuCliSize.cliSizeNormal, lblState.ToolTipText, funLoadResString(10179))
        uMsg.ImageIndex = Me.StatusGUI
        uMsg.Text = lblName & " - " & lblState.Caption
        If blnStateChanged Then Call funAddStatusBarMsg(uMsg, m_ConnectionIndex)
    
    End Select
    Exit Sub

subUpdateStateError:
    Call funErrorManager("ucWorkstation.subUpdateStateError", True)
End Sub
