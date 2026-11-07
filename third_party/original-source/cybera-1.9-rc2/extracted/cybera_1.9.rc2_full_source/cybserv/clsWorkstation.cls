VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWorkstation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ WORKSTATION CLASS __________________________|
'|  Name:          clsWorkstation.cls                                          |
'|  Type:          VB6 class                                                   |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
'|  Description:   Workstation model                                           |
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

' $Id: clsWorkstation.cls 15 2006-01-01 16:12:35Z sam $

Option Explicit

Public WorkstationName As String             'Name of workstation

Public ComputerName As String           'Real computername of the workstation

Public ComputerMacAddress As String

Public UserName As String

Public Key As String                    'Collection key of the workstation

Public Cnx As Object                    'Winsock instance for the workstation

Public State As Integer                 'Gives the state of the workstation

Public OldState As Integer                 'Gives the old state of the workstation

Public CnxIncomingBuffer As String


Public Status As Integer                'Alarm status

Public ElapsedTime As Date

Public UserBalance As Single              'Amount due for a selfservice session

Public RemainingTimeAparent As Date     'Apprent remaining time

Public ConnectionForced As Integer

Public CurrentSession As clsSession

Public SessionID As Long

Public StickySessionID As Long

Public ConnectedTime As Date            'Amount of time connected

Public ConnectedTimeOffPeak As Date            'Amount of time connected

Public DisconnectedTime As Date         'Amount of time disconnected

Public SessionType As Integer           'Session type (card or selfservice)

Public UnlimitedPlay As Boolean

Public StartTime As Date

Public CurrentCard As clsCard

Public CurrentClient As clsClient

Public SelfServicePrePaid As clsClient

Public LastVisitedURL As String

Public LastVisitedTitle As String

Public LastDBBackupDate As Date

Public ChatHistory As Collection



'Flags
Public AllowLockdown As Integer

Public AllowGroupControl As Boolean

Public AuthenticationInProgress As Boolean

Public CnxDataSendInProgress As Boolean

Public ReadyForServerScreen As Boolean
