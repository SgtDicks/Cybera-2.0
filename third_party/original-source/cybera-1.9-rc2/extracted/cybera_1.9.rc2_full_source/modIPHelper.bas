Attribute VB_Name = "modIPHelper"
' _____________________________________________________________________________
'|________________________________ IP HELPER FUNCTION MODULE __________________|
'|  Name:          modIPHelper.bas                                             |
'|  Type:          VB6 Module                                                  |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
'|  Description:   Contains Low level IP system functions such as adapter      |
'|                 listing etc.                                                |
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

' $Id: modIPHelper.bas 15 2006-01-01 16:12:35Z sam $

Option Explicit
Option Base 0

Private Declare Function GetAdaptersInfo Lib "IPHlpApi" (IpuAdapterInfo As Any, pOutBufLen As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As Long)

Private Const ERROR_BUFFER_OVERFLOW = 111&
Private Const MAX_ADAPTER_NAME_LENGTH As Long = 260
Private Const MAX_ADAPTER_ADDRESS_LENGTH As Long = 8
Private Const MAX_ADAPTER_DESCRIPTION_LENGTH As Long = 132

Public Const MIB_IF_TYPE_OTHER As Byte = 1
Public Const MIB_IF_TYPE_ETHERNET As Byte = 6
Public Const MIB_IF_TYPE_TOKENRING As Byte = 9
Public Const MIB_IF_TYPE_FDDI As Byte = 15
Public Const MIB_IF_TYPE_PPP As Byte = 23
Public Const MIB_IF_TYPE_LOOPBACK As Byte = 24
Public Const MIB_IF_TYPE_SLIP As Byte = 28

Private Type typIPStrs
    tNext As Long
    IpAddress As String * 16
    IpMask As String * 16
    Context As Long
End Type

Private Type typIpAdapterInfo
    tNext As Long
    ComboIndex As Long
    AdapterName As String * MAX_ADAPTER_NAME_LENGTH
    Description As String * MAX_ADAPTER_DESCRIPTION_LENGTH
    AddressLength As Long
    Address(MAX_ADAPTER_ADDRESS_LENGTH - 1) As Byte
    Index As Long
    tType As Long
    DhcpEnabled As Long
    CurrentIpAddress As Long
    IpAddressList As typIPStrs
    GatewayList As typIPStrs
    DhcpServer As typIPStrs
    HaveWins As Boolean
    PrimaryWinsServer As typIPStrs
    SecondaryWinsServer As typIPStrs
    LeaseObtained As Long
    LeaseExpires As Long
End Type

Public Type typIPuAdapter
    AdapterName As String
    AdapterType As String
    AdapterTypeID As Long
    Description As String
    MacAdress As String
    IpAddressCount As Long
    IpAddress() As String
End Type

Public Type typAdapters
    Count As Long
    Item() As typIPuAdapter
End Type


'*******************************************************************************
Public Function funGetIPAdapters() As typAdapters
Dim uReturnAdp As typAdapters
Dim lngError As Long
Dim lngAdapterInfoSize As Long
Dim i As Long
Dim strPhysicalAddress  As String
Dim strPhysicalAddressOctet  As String
Dim uAdapterInfo As typIpAdapterInfo
Dim uAdapt As typIpAdapterInfo
Dim uAddrStr As typIPStrs
Dim uBuffer As typIPStrs
Dim lngAddrStr As Long
Dim lngAdapt As Long
Dim uBuffer2 As typIpAdapterInfo
Dim bytAdapterInfoBuffer() As Byte
On Error GoTo funGetIPAdaptersError

    'Enumerate all of the uAdapter specific information using the IP_uAdaptER_INFO structure.
    'Note:  IP_uAdaptER_INFO contains a linked list of uAdapter entries.
    
    lngAdapterInfoSize = 0
    lngError = GetAdaptersInfo(ByVal 0&, lngAdapterInfoSize)
    If lngError <> 0 Then
        If lngError <> ERROR_BUFFER_OVERFLOW Then
            Debug.Print "GetuAdaptersInfo sizing failed with lngError " & lngError
            Exit Function '>---> Bottom
        End If
    End If
    ReDim bytAdapterInfoBuffer(lngAdapterInfoSize - 1)
 
    ' Get actual uAdapter information
    lngError = GetAdaptersInfo(bytAdapterInfoBuffer(0), lngAdapterInfoSize)
    If lngError <> 0 Then
        Debug.Print "GetuAdaptersInfo failed with lngError " & lngError
        Exit Function '>---> Bottom
    End If
    CopyMemory uAdapterInfo, bytAdapterInfoBuffer(0), Len(uAdapterInfo)
    lngAdapt = -1
    uReturnAdp.Count = 0
   
    Do While lngAdapt <> 0
        uReturnAdp.Count = uReturnAdp.Count + 1
        ReDim Preserve uReturnAdp.Item(uReturnAdp.Count)
        
        With uReturnAdp.Item(uReturnAdp.Count - 1)
        
            lngAdapt = uAdapterInfo.tNext
            CopyMemory uBuffer2, uAdapterInfo, Len(uBuffer2)
            .AdapterTypeID = uBuffer2.tType
            Select Case uBuffer2.tType
              Case MIB_IF_TYPE_ETHERNET
                .AdapterType = "Ethernet adapter"
              Case MIB_IF_TYPE_TOKENRING
                .AdapterType = "Token Ring adapter"
              Case MIB_IF_TYPE_FDDI
                .AdapterType = "FDDI adapter"
              Case MIB_IF_TYPE_PPP
                .AdapterType = "PPP adapter"
              Case MIB_IF_TYPE_LOOPBACK
                .AdapterType = "Loopback adapter"
              Case MIB_IF_TYPE_SLIP
                .AdapterType = "Slip adapter"
              Case Else
                .AdapterType = "Other adapter"
            End Select
       
            .AdapterName = uBuffer2.AdapterName
            .Description = uBuffer2.Description
            strPhysicalAddress = ""
            For i = 0 To uBuffer2.AddressLength - 1
                strPhysicalAddressOctet = Hex$(uBuffer2.Address(i))
                If Len(strPhysicalAddressOctet) = 2 Then
                    strPhysicalAddress = strPhysicalAddress & strPhysicalAddressOctet
                Else
                    strPhysicalAddress = strPhysicalAddress & "0" & strPhysicalAddressOctet
                End If
'                If i < uBuffer2.AddressLength - 1 Then
'                    strPhysicalAddress = strPhysicalAddress & "-"
'                End If

            Next i
            .MacAdress = strPhysicalAddress

            .IpAddressCount = 0
            lngAddrStr = -1
            Do While lngAddrStr <> 0
                .IpAddressCount = .IpAddressCount + 1
                lngAddrStr = uBuffer2.IpAddressList.tNext
                '' 24.Aug.2002:
                '' added preserve for servers with multiple ip-addresses
                '' edited by Sandeep Mandviya, india
                ReDim Preserve .IpAddress(.IpAddressCount - 1)
           
                CopyMemory uBuffer, uBuffer2.IpAddressList, LenB(uBuffer)
                .IpAddress(.IpAddressCount - 1) = left$(uBuffer.IpAddress, InStr(uBuffer.IpAddress, Chr$(0)) - 1)
                'Debug.Print .IpAddress(.IpAddressCount - 1) & .Description
                'Debug.Print "Subnet Mask: " & uBuffer.IpMask
                lngAddrStr = uBuffer.tNext
                If lngAddrStr <> 0 Then
                    CopyMemory uBuffer2.IpAddressList, ByVal lngAddrStr, Len(uBuffer2.IpAddressList)
                End If
            Loop
        End With 'uReturnAdp.item(uReturnAdp.count
    
        lngAdapt = uBuffer2.tNext
        If lngAdapt <> 0 Then
            CopyMemory uAdapterInfo, ByVal lngAdapt, Len(uAdapterInfo)
        End If
    
    Loop
    funGetIPAdapters = uReturnAdp

    Exit Function
    
funGetIPAdaptersError:
    Call funErrorManager("modIPHelper.funGetIPAdaptersError")
End Function


'*******************************************************************************
Public Function funGetIPfromMac(strMac As String) As String
Dim uAdapters As typAdapters
Dim i As Integer
On Error GoTo funGetIPfromMacError
    
    uAdapters = funGetIPAdapters
    If uAdapters.Count = 0 Then Exit Function
    For i = 0 To uAdapters.Count - 1
        If uAdapters.Item(i).MacAdress = strMac And uAdapters.Item(i).IpAddressCount > 0 Then
            funGetIPfromMac = uAdapters.Item(i).IpAddress(0)
            Exit For
        End If
    Next i
    Exit Function
    
funGetIPfromMacError:
    Call funErrorManager("modIPHelper.funGetIPfromMacError")
End Function

