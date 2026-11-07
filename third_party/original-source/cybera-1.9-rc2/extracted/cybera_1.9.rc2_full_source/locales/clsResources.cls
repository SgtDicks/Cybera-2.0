VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsResources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' _____________________________________________________________________________
'|________________________________ SERVICE CLASS ______________________________|
'|  Name:          clsResources.cls                                            |
'|  Type:          VB6 class                                                   |
'|  Project:       Cybera Server & Client                                      |
'|  Description:   Language resource class                                     |
'|  Copyright:     (C) 2004-2006 Samuel Monsarrat (smonsarr@users.sf.net)      |
'|                                                                             |
'| This file is part of Cybera.                                                |
'|                                                                             |
'| Cybera is free software; you can redistribute it and/or modify              |
'| it under the terms of the GNU General Public License as published by        |
'| the Free Software Foundation; either version 2 of the License, or           |
'| (at your option) any later version.                                         |
'|                                                                             |
'| Cybera is distributed in the hope that it will be useful,                   |
'| but WITHOUT ANY WARRANTY; without even the implied warranty of              |
'| MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the               |
'| GNU General Public License for more details.                                |
'|                                                                             |
'| You should have received a copy of the GNU General Public License           |
'| along with Cybera Server; if not, write to the Free Software                |
'| Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA   |
'|                                                                             |
'|_____________________________________________________________________________|

' $Id: clsResources.cls 39 2006-02-20 19:21:20Z sam $

Option Explicit


'*******************************************************************************
Public Function GetResStr(ResourceIndex As Long) As String
   
   GetResStr = LoadResData(ResourceIndex, 6)
End Function


'*******************************************************************************
Public Function GetLangName() As String
   
   GetLangName = LoadResData(10014, 6)
End Function


'*******************************************************************************
Public Function GetLangNameWestern() As String
   
   GetLangNameWestern = LoadResData(1, 6)
End Function


'*******************************************************************************
Public Function GetTradName() As String
   
   GetTradName = LoadResData(10045, 6)
End Function


'*******************************************************************************
Public Function GetCharset() As Integer
   
   GetCharset = CInt(LoadResData(0, 6))
End Function


