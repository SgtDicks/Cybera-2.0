VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMagicPacket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ SERVICE CLASS ______________________________|
'|  Name:          clsMagicPacket.cls                                          |
'|  Type:          VB6 class                                                   |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
'|  Description:   Wake on LAN magic packet class                              |
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

' $Id: clsMagicPacket.cls 15 2006-01-01 16:12:35Z sam $


'************************************************************************
'
' CMagicPacket - AMD Magic Packet (Wake-On-Lan) Class Module (v1.1)
' By Richard Larson, 1999/2000 - Requires VB 6.0 or better.
'
' This class encapsulates the functionality necessary to send Wake-on-Lan
' packets to WFM (Wired-For-Management) enabled computers that reside on
' an IP network.
'
'************************************************************************
'
' Version History:
'
' 1.1 - Fixed a bug that prevented IP addresses with individual octet
'       values of zero from being properly validated in the
'       ValidateIPAddress() function.
'
' 1.0 - Initial release to Internet.
'
'************************************************************************
'
' External Class Dependencies:
'
' This class depends on the existance of a winsock control on a form
' somewhere in the project.  Each instance of this object must have a
' reference set to such a winsock control before it can be used to
' wake-up a workstation.  This is accomplished via the WinsockControl
' property of the class.
'
'************************************************************************
'
' To create this class, open a new VB6 project and add a class module to
' it.  Name the class module "CMagicPacket" and then paste this code
' snippet into it.   Then save the class and add it to any project that
' needs this functionality.
'
' To use an object of this class, you must provide an MS Winsock control
' on one of your forms and then set the WinSockControl property to it as
' a reference.  Then you simply set the NICAddress, IPSubnetAddress and
' IPSubnetMask properties of the new object and call the wakeup method.
'
' Example:
'
'   Dim WOL_WS as CMagicPacket
'   Set WOL_WS = New CMagicPacket
'   With WOL_WS
'     .WinsockCtrl=Winsock1
'     .NICAddress="0008d7aa4e24"
'     .IPSubnetAddress="172.17.1.0"
'     .IPSubnetMask="255.255.255.0"
'     .WakeUp
'   End With
'   Set WOL_WS=Nothing
'
' If the IPSubnetAddress or IPSubnetMask properties are left blank
' (or are determined to be invalid), the packet will be broadcast on
' the IP subnet local to the workstation that the object is running on.
'
'************************************************************************
Option Explicit

'Object-level constants and enums...
Private Const ERROR_BASE = vbObjectError + 2000
Private Const ERROR_INVALID_NIC_ADDRESS = ERROR_BASE + 1
Private Const ERROR_INVALID_IP_SUBNET_ADDRESS = ERROR_BASE + 2
Private Const ERROR_INVALID_IP_SUBNET_MASK = ERROR_BASE + 3
Private Const ERROR_INVALID_WINSOCK_REFERENCE = ERROR_BASE + 4
Private Const ERROR_NO_WINSOCK_REFERENCE_AVAILABLE = ERROR_BASE + 5

'Object-level variables...
Private m_WinsockControl As Control
Private m_NICAddress As String
Private m_IPSubnetAddress As String
Private m_IPSubnetMask As String


'************************************************************************
' Property WinsockControl()
'
' This is the property assignment function (LET only) for the Winsock
' Control property.   It validates that the winsock control being
' assigned is really a winsock control, and that it is not set to
' "nothing".  If either of these checks fails, an appropriate error
' message is raised back to the caller.
'
'************************************************************************
Public Property Let WinsockControl(ByRef RefWinsock As Control)
  
  'Check for uninitialized controls...
  If RefWinsock Is Nothing Then
    Err.Raise ERROR_NO_WINSOCK_REFERENCE_AVAILABLE, "CMagicPacket.WinsockControl(LET)", "An attempt was made to set the WinsockControl property " & "to an uninitialized control"
  End If
  
  'Check for invalid control types...
  If TypeOf RefWinsock Is Winsock Then
    Set m_WinsockControl = RefWinsock
  Else
    Err.Raise ERROR_INVALID_WINSOCK_REFERENCE, "CMagicPacket.WinsockControl(LET)", "An attempt was made to set the WinsockControl property " & "to a non-winsock control"
  End If
  
  'It passed the test, so set our reference to it...
  Set m_WinsockControl = RefWinsock
  
End Property


'************************************************************************
' Property IPSubnetAddress()
'
' These are the assignment functions for the IP Address property. LET
' validates the IP subnet address being assigned and raises an error back
' to the calling function if it was not a valid 32-bit IP address. GET
' returns the current address in dotted-decimal notation (www.xxx.yyy.zzz)
'
'************************************************************************
Public Property Let IPSubnetAddress(ByVal NewAddress As String)
  
  'Set the new IP subnet address...
  m_IPSubnetAddress = ValidateIPAddress(NewAddress)
  
  'And raise an error if it is not valid...
  If m_IPSubnetAddress = "" Then
    Err.Raise ERROR_INVALID_IP_SUBNET_ADDRESS, "CMagicPacket.IPSubnetAddress(LET)", "Invalid IP Subnet Address"
  End If
  
End Property

Public Property Get IPSubnetAddress() As String
  IPSubnetAddress = m_IPSubnetAddress
End Property


'************************************************************************
' Property IPSubnetMask()
'
' These are the assignment functions for the IP Subnet Mask property. LET
' validates the IP subnet mask being assigned and raises an error back to
' the calling function if it was not a valid 32-bit IP mask address. GET
' returns the current address in dotted-decimal notation (www.xxx.yyy.zzz)
'
'************************************************************************
Public Property Let IPSubnetMask(ByVal NewMask As String)
  
  'Set the new IP subnet address...
  m_IPSubnetMask = ValidateIPAddress(NewMask)
  
  'And raise an error if it is not valid...
  If m_IPSubnetAddress = "" Then
    Err.Raise ERROR_INVALID_IP_SUBNET_MASK, "CMagicPacket.IPSubnetMask(LET)", "Invalid IP Subnet Address"
  End If
  
End Property

Public Property Get IPSubnetMask() As String
  IPSubnetMask = m_IPSubnetMask
End Property


'************************************************************************
' Property NICAddress()
'
' These are the assignment functions for the NIC Address property. LET
' validates the NIC Address being assigned and raises an error back to
' the calling function if it was not a valid 48-bit NIC address. GET
' returns the 48-bit unformatted hex address in the form "1A2B3C4D5E6F"
'
'************************************************************************
Public Property Let NICAddress(ByVal NewNicAddress As String)
  
  'Set the new NIC address...
  m_NICAddress = ValidateNICAddress(NewNicAddress)
  
  'And raise an error if it was not valid...
  If m_NICAddress = "" Then
    Err.Raise ERROR_INVALID_NIC_ADDRESS, "CMagicPacket.NicAddress(Let)", "Invalid Nic Address"
  End If
  
End Property

Public Property Get NICAddress() As String
  NICAddress = m_NICAddress
End Property


'************************************************************************
' Method WakeUp()
'
' This method is called in order to send the wake-on-lan magic packet to
' the machine with the 48-bit hardware address specified in NICAdress
' residing on the IP subnet specified by the IPSubnetAddress/IPSubnetMask
' properties.  It uses the winsock control that was set in the
' WinSockCtrl property to do the actual packet transfer.  An error flag
' will be raised back to the caller if any of these properties are set
' incorrectly.
'
'************************************************************************
Public Sub WakeUp()

Dim ActualNICAddress As String

  'Setup error handler...
  On Error GoTo Error_Handler1
  
  
  
  'Check for valid properties before sending the packet...
  ActualNICAddress = ValidateNICAddress(m_NICAddress)
  
  If ActualNICAddress = "" Then
    Err.Raise ERROR_INVALID_NIC_ADDRESS, "CMagicPacket.Wakeup()", "The NICAddress property does not contain a valid NIC address"
  End If
  
  If m_WinsockControl Is Nothing Then
    Err.Raise ERROR_NO_WINSOCK_REFERENCE_AVAILABLE, "CMagicPacket.Wakeup()", "The WinsockCtrl property has not been set to a valid MS-Winsock control"
  End If
  
  'Everything is OK, so send the packet on its way...
  SendMagicPacketTo m_IPSubnetAddress, m_IPSubnetMask, m_NICAddress, m_WinsockControl
  Exit Sub
  
Error_Handler1:

  'For now, just return the error to the caller with a modified source...
'  Err.Raise Err.Number, Err.source & " <Raised From CMagicPacket.WakeUp()>", Err.description, Err.HelpFile, Err.HelpContext
End Sub


'************************************************************************
' SendMagicPacketTo()
'
' Send a "Magic-packet" to a workstation that supports the "wake-on-lan"
' specification.  This routine takes the IP address of the subnet on
' which the workstation to be "awakened" resides, and the 48-bit ethernet
' hardware address of the workstation's NIC.  It then builds and sends a
' "magic packet" to that workstation which will wake it up (as though
' someone switched on the power).
'
'************************************************************************
Private Sub SendMagicPacketTo(IPSubnetAddress As String, IPSubnetMask As String, HWAddress As String, WinsockCtrl As Winsock)

'Local Variables...
Dim MagicPacketData(0 To 101) As Byte
Dim HWAddressByteValues(0 To 5) As Byte
Dim ByteIndex As Byte
Dim NICAddressIndex As Byte

  'Setup error handler...
  On Error GoTo Error_Handler
 
  'First convert the Ethernet HWAddress to a byte array...
  HWAddressStrToByteArray HWAddress, HWAddressByteValues()
 
  'Now get the IP broadcast address for the IP subnet we want to use...
  IPSubnetAddress = IPSubnetAddressToBroadcastAddress(IPSubnetAddress, IPSubnetMask)

  'Setup the Magic-Packet synchronization byte sequence (0xFFFFFFFFFFFF)...
  For ByteIndex = 0 To 5
    MagicPacketData(ByteIndex) = 255
  Next
 
  'Now add the Magic-Packet HWAddress sequence (16 HW Addresses in a row)...
  For NICAddressIndex = 0 To 15
    For ByteIndex = 0 To 5
      MagicPacketData((NICAddressIndex * 6) + ByteIndex + 6) = HWAddressByteValues(ByteIndex)
    Next
  Next
 
  'Send a datagram containing the magic-packet data to the appropriate subnet...
  m_WinsockControl.RemoteHost = IPSubnetAddress
  m_WinsockControl.Protocol = sckUDPProtocol
  m_WinsockControl.SendData MagicPacketData
  m_WinsockControl.Close
  Exit Sub
 
Error_Handler:

  'For now, just return the error back to the caller with a modified source...
'    Call funErrorManager("clsMagicPacket.SendMagicPacketTo")
  
  Err.Raise Err.Number, Err.source & " <Raised From CMagicPacket.SendMagicPacketTo()>", Err.Description, Err.HelpFile, Err.HelpContext

End Sub


'************************************************************************
' ValidateNICAddress()
'
' This function takes a string and determines if it is a valid NIC
' address by checking it's length (6-hex bytes, 48-bits) and that it is
' a valid hexadecimal value (each nibble is 0-9, or A-F).  If all is OK,
' the function returns a cleaned version of the address (with no
' whitespace characters and all upper case values), otherwise it returns
' a blank zero-length string.
'
' Based on this routine, all of the following addresses are valid:
'
'  "1a2b3c4d5e6f" returns "1A2B3C4D5E6F"
'  "1a.2b.3c.4d.5e.6f" returns "1A2B3C4D5E6F"
'  "1a 2b 3c 4d 5e 6f" returns "1A2B3C4D5E6F"
'  "1a2b 3c4d 5e6f" returns "1A2B3C4D5E6F"      Etc....
'
'************************************************************************
Public Function ValidateNICAddress(AddressToValidate As String) As String

'Local Constants...
Const VALID_NIC_ADDRESS_LENGTH = 12

'Local Variables...
Dim HexCharIndex As Long
Dim HexChar As String
Dim InitialAddress As String
Dim ReducedAddress As String

  'Init local variables...
  ReducedAddress = ""
  InitialAddress = UCase$(Trim$(AddressToValidate))
  
  'Eliminate all non-hexadecimal characters from the address...
  For HexCharIndex = 1 To Len(InitialAddress)
    HexChar = Mid$(InitialAddress, HexCharIndex, 1)
    Select Case HexChar
      Case "0" To "9", "A" To "F": ReducedAddress = ReducedAddress & HexChar
    End Select
  Next

  'Return the appropriate value...
  If Len(ReducedAddress) = VALID_NIC_ADDRESS_LENGTH Then
    ValidateNICAddress = UCase$(ReducedAddress)
  Else
    ValidateNICAddress = ""
  End If

End Function


'************************************************************************
' ValidateIPAddress()
'
' This function takes an IP address in dotted decimal format and verifies
' that there are 4 octets of data and that the value of each octet is
' between zero and 255.  This function does NOT determine whether the
' given address is legal for any particular subnetting scheme. If the
' address is valid, it will return the address itself, otherwise it will
' return an empty (zero-length) string.
'
'************************************************************************
Private Function ValidateIPAddress(AddressToValidate As String) As String

'Local Variables...
Dim TempStr As String
Dim ReducedAddress As String
Dim SplitAddress() As String
Dim OctetIndex As Long
Dim OctetValue As Long
  
  'First remove all non-decimal, non-period characters from the address...
  For OctetIndex = 1 To Len(AddressToValidate)
    TempStr = Mid$(AddressToValidate, OctetIndex, 1)
    Select Case TempStr
      Case "0" To "9", ".": ReducedAddress = ReducedAddress & TempStr
    End Select
  Next
  
  'Split the octets into a string array for inspection...
  SplitAddress = Split(ReducedAddress, ".")
  ReducedAddress = ""
  
  'If there are not exactly 4 octets (0->3), return empty...
  If UBound(SplitAddress()) <> 3 Then
    ValidateIPAddress = ""
    Exit Function
  End If

  'Verify that the value of each octet is between 0 and 255...
  For OctetIndex = 0 To 3
    TempStr = SplitAddress(OctetIndex)
    OctetValue = Val(TempStr)
    If Len(TempStr) = 0 Or OctetValue < 0 Or OctetValue > 255 Then
      ValidateIPAddress = ""
      Exit Function
    End If
    ReducedAddress = ReducedAddress & format$(OctetValue, "##0")
    If OctetIndex < 3 Then ReducedAddress = ReducedAddress & "."
  Next
  
  'Return the validated result...
  ValidateIPAddress = ReducedAddress
  
End Function


'************************************************************************
' HWAddressStrToByteArray()
'
' This routine takes a hexadecimal ethernet NIC address in string format
' and converts it into a 6-element (0 to 5) array of byte values
' representing that same address.  This routine is used when building
' wake-on-lan "Magic-Packets".
'
'************************************************************************
Private Sub HWAddressStrToByteArray(HexHWAddress As String, ByRef HWAddressByteArray() As Byte)

'Local Variables...
Dim ByteIndex As Long

  'Clear out the array to start...
  For ByteIndex = 0 To 5
    HWAddressByteArray(ByteIndex) = 0
  Next

  'Convert each hex double-digit to a byte in the array...
  For ByteIndex = 1 To 11 Step 2
    HWAddressByteArray((ByteIndex - 1) / 2) = Val("&H" & Mid$(HexHWAddress, ByteIndex, 2))
  Next
  
End Sub


'************************************************************************
' IPSubnetAddressToBCastAddress()
'
' This routine takes an IP address and it's associated subnet mask and
' returns the directed-broadcast address for that subnet. If either
' parameter is invalid the local broadcast address (255.255.255.255)
' is returned.
'
'************************************************************************
Private Function IPSubnetAddressToBroadcastAddress(IPSubnetAddress As String, IPSubnetMask As String) As String

'Local Variables...
Dim IPStrArray() As String
Dim MaskStrArray() As String
Dim BCastAddressByte As Byte
Dim BCastAddress As String
Dim ByteIndex As Long

  'Check for valid input...
  If ValidateIPAddress(IPSubnetAddress) = "" Or ValidateIPAddress(IPSubnetMask) = "" Then
       IPSubnetAddressToBroadcastAddress = "255.255.255.255"
       Exit Function
  End If

  'Initialize the strings...
  BCastAddress = ""
  IPStrArray = Split(IPSubnetAddress, ".")
  MaskStrArray = Split(IPSubnetMask, ".")
  
  'Build the new broadcast address byte-by-byte...
  For ByteIndex = 0 To 3
    
    BCastAddressByte = Val(IPStrArray(ByteIndex)) Or (Val(MaskStrArray(ByteIndex)) Xor 255)
    
    BCastAddress = BCastAddress & CStr(BCastAddressByte)
    If ByteIndex < 3 Then BCastAddress = BCastAddress & "."
  
  Next
  
  'And finally return it to the calling routine...
  IPSubnetAddressToBroadcastAddress = BCastAddress
  
End Function



