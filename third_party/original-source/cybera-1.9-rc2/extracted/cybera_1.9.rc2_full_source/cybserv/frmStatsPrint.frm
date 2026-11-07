VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmStatsPrint 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   11880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16815
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   20.955
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   29.66
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fxgContent 
      Height          =   10095
      Left            =   705
      TabIndex        =   0
      Top             =   960
      Width           =   15390
      _ExtentX        =   27146
      _ExtentY        =   17806
      _Version        =   393216
      Rows            =   42
      Cols            =   30
      FixedCols       =   0
      BackColorFixed  =   12632256
      BackColorBkg    =   -2147483643
      GridColor       =   0
      GridColorFixed  =   0
      GridColorUnpopulated=   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      GridLinesUnpopulated=   1
      ScrollBars      =   0
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   30
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   720
      Picture         =   "frmStatsPrint.frx":0000
      Top             =   360
      Width           =   480
   End
   Begin VB.Label lblDtePage 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   420
      Width           =   14295
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1260
      TabIndex        =   1
      Top             =   420
      Width           =   13875
   End
End
Attribute VB_Name = "frmStatsPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' _____________________________________________________________________________
'|________________________________ STATS PRINTING FORM ________________________|
'|  Name:          frmStatsPrint.frm                                           |
'|  Type:          VB6 Form                                                    |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
'|  Description:   Statistics printing form                                    |
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

' $Id: frmStatsPrint.frm 15 2006-01-01 16:12:35Z sam $

Option Explicit

Private Sub Form_Load()
On Error GoTo Form_LoadError

'    Call subSetDefaultFont(Me)
    Exit Sub

Form_LoadError:
    Call funErrorManager("frmStatsPrint.Form_LoadError", True)
End Sub

