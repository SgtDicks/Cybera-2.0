Attribute VB_Name = "modScreenCapture"
' _____________________________________________________________________________
'|________________________________ SCREEN & WINDOW CAPTURE ROUTINES ___________|
'|  Name:          modScreenCapture.bas                                        |
'|  Type:          VB6 Module                                                  |
'|  Project:       Cybera Server (http://cybera.sf.net)                        |
'|  Description:   Contains routines for screen and window capture             |
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

' $Id: modScreenCapture.bas 15 2006-01-01 16:12:35Z sam $

Option Explicit
Option Base 0

Private Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type

Private Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(255) As PALETTEENTRY  ' Enough for 256 colors.
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Const RASTERCAPS As Long = 38
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104

Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "GDI32" (ByVal hDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "GDI32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SelectPalette Lib "GDI32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Type PicBmp
    size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long


'*******************************************************************************
' - Creates a bitmap type Picture object from a bitmap and palette.
'
' hBmp - Handle to a bitmap.
' hPal - Handle to a Palette. Can be null if the bitmap doesn't use a palette.
'
' Returns - Returns a Picture object containing the bitmap.
Public Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
Dim r As Long
Dim pic As PicBmp
' IPicture requires a reference to "Standard OLE Types."
Dim IPic As IPicture
Dim IID_IDispatch As GUID
On Error GoTo CreateBitmapPictureError
    
    ' Fill in with IDispatch Interface ID.
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    
    ' Fill Pic with necessary parts.
    With pic
        .size = Len(pic)          ' Length of structure.
        .Type = vbPicTypeBitmap   ' Type of Picture (bitmap).
        .hBmp = hBmp              ' Handle to bitmap.
        .hPal = hPal              ' Handle to palette (may be null).
    End With
    
    ' Create Picture object.
    r = OleCreatePictureIndirect(pic, IID_IDispatch, 1, IPic)
    
    ' Return the new Picture object.
    Set CreateBitmapPicture = IPic
    Exit Function

CreateBitmapPictureError:
    Call funErrorManager("modScreenCapture.CreateBitmapPictureError")
End Function


'*******************************************************************************
' - Captures any portion of a window.
'
' hWndSrc - Handle to the window to be captured.
' Client - If True CaptureWindow captures from the client area of the window. If False CaptureWindow captures from the entire window.
' LeftSrc, TopSrc, WidthSrc, HeightSrc
'    - Specify the portion of the window to capture.
'    - Dimensions need to be specified in pixels.
'
' Returns - Returns a Picture object containing a bitmap of the specified portion of the window that was captured.
Public Function CaptureWindow(ByVal hWndSrc As Long, ByVal Client As Boolean, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture
Dim hDCMemory As Long
Dim hBmp As Long
Dim hBmpPrev As Long
Dim r As Long
Dim hDCSrc As Long
Dim hPal As Long
Dim hPalPrev As Long
Dim RasterCapsScrn As Long
Dim HasPaletteScrn As Long
Dim PaletteSizeScrn As Long
Dim LogPal As LOGPALETTE
On Error GoTo CaptureWindowError
    
    ' Depending on the value of Client get the proper device context.
    If Client Then
        hDCSrc = GetDC(hWndSrc) ' Get device context for client area.
    Else
        hDCSrc = GetWindowDC(hWndSrc) ' Get device context for entire window.
    End If
    
    ' Create a memory device context for the copy process.
    hDCMemory = CreateCompatibleDC(hDCSrc)
    ' Create a bitmap and place it in the memory DC.
    hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
    hBmpPrev = SelectObject(hDCMemory, hBmp)
    
    ' Get screen properties.
    RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster
                                          ' capabilities.
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE       ' Palette
                                            ' support.
    PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of
                                            ' palette.
    
    ' If the screen has a palette make a copy and realize it.
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        ' Create a copy of the system palette.
        LogPal.palVersion = &H300
        LogPal.palNumEntries = 256
        r = GetSystemPaletteEntries(hDCSrc, 0, 256, _
        LogPal.palPalEntry(0))
        hPal = CreatePalette(LogPal)
        ' Select the new palette into the memory DC and realize it.
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        r = RealizePalette(hDCMemory)
    End If
    
    ' Copy the on-screen image into the memory DC.
    r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, _
    LeftSrc, TopSrc, vbSrcCopy)
    
    ' Remove the new copy of the  on-screen image.
    hBmp = SelectObject(hDCMemory, hBmpPrev)
    
    ' If the screen has a palette get back the palette that was
    ' selected in previously.
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If
    
    ' Release the device context resources back to the system.
    r = DeleteDC(hDCMemory)
    r = ReleaseDC(hWndSrc, hDCSrc)
    
    ' Call CreateBitmapPicture to create a picture object from the
    ' bitmap and palette handles. Then return the resulting picture
    ' object.
    Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
    Exit Function

CaptureWindowError:
    Call funErrorManager("modScreenCapture.CaptureWindowError")
End Function


'*******************************************************************************
' CaptureScreen
' - Captures the entire screen.
'
' Returns - Returns a Picture object containing a bitmap of the screen.
Public Function CaptureScreen() As Picture
Dim hWndScreen As Long
On Error GoTo CaptureScreenError
    
    ' Get a handle to the desktop window.
    hWndScreen = GetDesktopWindow()
    
    ' Call CaptureWindow to capture the entire desktop give the handle
    ' and return the resulting Picture object.
    
    Set CaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, (Screen.width \ Screen.TwipsPerPixelX), (Screen.height \ Screen.TwipsPerPixelY))
    Exit Function

CaptureScreenError:
    Call funErrorManager("modScreenCapture.CaptureScreenError")
End Function


'*******************************************************************************
' CaptureForm - Captures an entire form including title bar and border.
' frmSrc - The Form object to capture.
' Returns - Returns a Picture object containing a bitmap of the entire form.
Public Function CaptureForm(frmSrc As Form) As Picture
On Error GoTo CaptureFormError
    
    ' Call CaptureWindow to capture the entire form given its window
    ' handle and then return the resulting Picture object.
    Set CaptureForm = CaptureWindow(frmSrc.hWnd, False, 0, 0, frmSrc.ScaleX(frmSrc.width, vbTwips, vbPixels), frmSrc.ScaleY(frmSrc.height, vbTwips, vbPixels))
    Exit Function

CaptureFormError:
    Call funErrorManager("modScreenCapture.CaptureFormError")
End Function


'*******************************************************************************
' CaptureClient - Captures the client area of a form.
' frmSrc - The Form object to capture.
' Returns - Returns a Picture object containing a bitmap of the form's client area.
Public Function CaptureClient(frmSrc As Form) As Picture
On Error GoTo CaptureClientError
    
    ' Call CaptureWindow to capture the client area of the form given
    ' its window handle and return the resulting Picture object.
    Set CaptureClient = CaptureWindow(frmSrc.hWnd, True, 0, 0, frmSrc.ScaleX(frmSrc.ScaleWidth, frmSrc.ScaleMode, vbPixels), frmSrc.ScaleY(frmSrc.ScaleHeight, frmSrc.ScaleMode, vbPixels))
    Exit Function

CaptureClientError:
    Call funErrorManager("modScreenCapture.CaptureClientError")
End Function


'*******************************************************************************
' CaptureActiveWindow - Captures the currently active window on the screen.
' Returns - Returns a Picture object containing a bitmap of the activewindow.
Public Function CaptureActiveWindow() As Picture
Dim hWndActive As Long
Dim r As Long
Dim RectActive As RECT
On Error GoTo CaptureActiveWindowError
    
    ' Get a handle to the active/foreground window.
    hWndActive = GetForegroundWindow()
    
    ' Get the dimensions of the window.
    r = GetWindowRect(hWndActive, RectActive)
    
    ' Call CaptureWindow to capture the active window given its
    ' handle and return the Resulting Picture object.
    Set CaptureActiveWindow = CaptureWindow(hWndActive, False, 0, 0, RectActive.right - RectActive.left, RectActive.bottom - RectActive.top)
    Exit Function

CaptureActiveWindowError:
    Call funErrorManager("modScreenCapture.CaptureActiveWindowError")
End Function


'*******************************************************************************
' PrintPictureToFitPage - Prints a Picture object as big as possible.
' Prn - Destination Printer object.
' Pic - Source Picture object.
Public Sub PrintPictureToFitPage(Prn As Printer, pic As Picture)
Const vbHiMetric As Integer = 8
Dim PicRatio As Double
Dim PrnWidth As Double
Dim PrnHeight As Double
Dim PrnRatio As Double
Dim PrnPicWidth As Double
Dim PrnPicHeight As Double
On Error GoTo PrintPictureToFitPageError
    
    ' Determine if picture should be printed in landscape or portrait
    ' and set the orientation.
    If pic.height >= pic.width Then
        Prn.Orientation = vbPRORPortrait   ' Taller than wide.
    Else
        Prn.Orientation = vbPRORLandscape  ' Wider than tall.
    End If
    
    ' Calculate device independent Width-to-Height ratio for picture.
    PicRatio = pic.width / pic.height
    
    ' Calculate the dimentions of the printable area in HiMetric.
    PrnWidth = Prn.ScaleX(Prn.ScaleWidth, Prn.ScaleMode, vbHiMetric)
    PrnHeight = Prn.ScaleY(Prn.ScaleHeight, Prn.ScaleMode, vbHiMetric)
    ' Calculate device independent Width to Height ratio for printer.
    PrnRatio = PrnWidth / PrnHeight
    
    ' Scale the output to the printable area.
    If PicRatio >= PrnRatio Then
        ' Scale picture to fit full width of printable area.
        PrnPicWidth = Prn.ScaleX(PrnWidth, vbHiMetric, Prn.ScaleMode)
        PrnPicHeight = Prn.ScaleY(PrnWidth / PicRatio, vbHiMetric, _
        Prn.ScaleMode)
    Else
        ' Scale picture to fit full height of printable area.
        PrnPicHeight = Prn.ScaleY(PrnHeight, vbHiMetric, Prn.ScaleMode)
        PrnPicWidth = Prn.ScaleX(PrnHeight * PicRatio, vbHiMetric, _
        Prn.ScaleMode)
    End If
    
    ' Print the picture using the PaintPicture method.
    Prn.PaintPicture pic, 0, 0, PrnPicWidth, PrnPicHeight
    Exit Sub

PrintPictureToFitPageError:
    Call funErrorManager("modScreenCapture.PrintPictureToFitPageError")
End Sub

