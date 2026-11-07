Attribute VB_Name = "modSystemParametersInfo"
' _____________________________________________________________________________
'|________________________________ SPECIAL FOLDER RETREAVAL ___________________|
'|  Name:          modSystemParametersInfo.bas                                 |
'|  Type:          VB6 Module                                                  |
'|  Project:       Cybera Server & Workstation                                 |
'|  Description:   Contains Windows System parameter Information functions.    |
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

' $Id: modSystemParametersInfo.bas 15 2006-01-01 16:12:35Z sam $

Option Explicit

' Parameters for SystemParametersInfo()
Const SPI_GETBEEP = 1
Const SPI_SETBEEP = 2
Const SPI_GETMOUSE = 3
Const SPI_SETMOUSE = 4
Const SPI_GETBORDER = 5
Const SPI_SETBORDER = 6
Const SPI_GETKEYBOARDSPEED = 10
Const SPI_SETKEYBOARDSPEED = 11
Const SPI_LANGDRIVER = 12
Const SPI_ICONHORIZONTALSPACING = 13
Const SPI_GETSCREENSAVETIMEOUT = 14
Const SPI_SETSCREENSAVETIMEOUT = 15
Const SPI_GETSCREENSAVEACTIVE = 16
Const SPI_SETSCREENSAVEACTIVE = 17
Const SPI_GETGRIDGRANULARITY = 18
Const SPI_SETGRIDGRANULARITY = 19
Const SPI_SETDESKWALLPAPER = 20
Const SPI_SETDESKPATTERN = 21
Const SPI_GETKEYBOARDDELAY = 22
Const SPI_SETKEYBOARDDELAY = 23
Const SPI_ICONVERTICALSPACING = 24
Const SPI_GETICONTITLEWRAP = 25
Const SPI_SETICONTITLEWRAP = 26
Const SPI_GETMENUDROPALIGNMENT = 27
Const SPI_SETMENUDROPALIGNMENT = 28
Const SPI_SETDOUBLECLKWIDTH = 29
Const SPI_SETDOUBLECLKHEIGHT = 30
Const SPI_GETICONTITLELOGFONT = 31
Const SPI_SETDOUBLECLICKTIME = 32
Const SPI_SETMOUSEBUTTONSWAP = 33
Const SPI_SETICONTITLELOGFONT = 34
Const SPI_GETFASTTASKSWITCH = 35
Const SPI_SETFASTTASKSWITCH = 36
Const SPI_SETDRAGFULLWINDOWS = 37
Const SPI_GETDRAGFULLWINDOWS = 38
Const SPI_GETNONCLIENTMETRICS = 41
Const SPI_SETNONCLIENTMETRICS = 42
Const SPI_GETMINIMIZEDMETRICS = 43
Const SPI_SETMINIMIZEDMETRICS = 44
Const SPI_GETICONMETRICS = 45
Const SPI_SETICONMETRICS = 46
Const SPI_SETWORKAREA = 47
Const SPI_GETWORKAREA = 48
Const SPI_SETPENWINDOWS = 49
Const SPI_GETFILTERKEYS = 50
Const SPI_SETFILTERKEYS = 51
Const SPI_GETTOGGLEKEYS = 52
Const SPI_SETTOGGLEKEYS = 53
Const SPI_GETMOUSEKEYS = 54
Const SPI_SETMOUSEKEYS = 55
Const SPI_GETSHOWSOUNDS = 56
Const SPI_SETSHOWSOUNDS = 57
Const SPI_GETSTICKYKEYS = 58
Const SPI_SETSTICKYKEYS = 59
Const SPI_GETACCESSTIMEOUT = 60
Const SPI_SETACCESSTIMEOUT = 61
Const SPI_GETSERIALKEYS = 62
Const SPI_SETSERIALKEYS = 63
Const SPI_GETSOUNDSENTRY = 64
Const SPI_SETSOUNDSENTRY = 65
Const SPI_GETHIGHCONTRAST = 66
Const SPI_SETHIGHCONTRAST = 67
Const SPI_GETKEYBOARDPREF = 68
Const SPI_SETKEYBOARDPREF = 69
Const SPI_GETSCREENREADER = 70
Const SPI_SETSCREENREADER = 71
Const SPI_GETANIMATION = 72
Const SPI_SETANIMATION = 73
Const SPI_GETFONTSMOOTHING = 74
Const SPI_SETFONTSMOOTHING = 75
Const SPI_SETDRAGWIDTH = 76
Const SPI_SETDRAGHEIGHT = 77
Const SPI_SETHANDHELD = 78
Const SPI_GETLOWPOWERTIMEOUT = 79
Const SPI_GETPOWEROFFTIMEOUT = 80
Const SPI_SETLOWPOWERTIMEOUT = 81
Const SPI_SETPOWEROFFTIMEOUT = 82
Const SPI_GETLOWPOWERACTIVE = 83
Const SPI_GETPOWEROFFACTIVE = 84
Const SPI_SETLOWPOWERACTIVE = 85
Const SPI_SETPOWEROFFACTIVE = 86
Const SPI_SETCURSORS = 87
Const SPI_SETICONS = 88
Const SPI_GETDEFAULTINPUTLANG = 89
Const SPI_SETDEFAULTINPUTLANG = 90
Const SPI_SETLANGTOGGLE = 91
Const SPI_GETWINDOWSEXTENSION = 92
Const SPI_SETMOUSETRAILS = 93
Const SPI_GETMOUSETRAILS = 94
Const SPI_SCREENSAVERRUNNING = 97

' SystemParametersInfo flags
Const SPIF_UPDATEINIFILE = &H1
Const SPIF_SENDWININICHANGE = &H2

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long


'*******************************************************************************
Public Function funScreensaverActive() As Boolean
Dim lngResult As Long
Dim lpvParam As Long
On Error GoTo funScreensaverActiveError

    lngResult = SystemParametersInfo(SPI_GETSCREENSAVEACTIVE, 0, lpvParam, 0)
    
    If lngResult = 0 Then
        funScreensaverActive = False
    Else
        funScreensaverActive = Not lpvParam <> 0
    End If
    Exit Function

funScreensaverActiveError:
    funScreensaverActive = False
    Call funErrorManager("modFunctions.funScreensaverActiveError")
End Function

