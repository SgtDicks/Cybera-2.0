Attribute VB_Name = "modSpecialFolders"
' _____________________________________________________________________________
'|________________________________ SPECIAL FOLDER RETREAVAL ___________________|
'|  Name:          modSpecialFolders.bas                                       |
'|  Type:          VB6 Module                                                  |
'|  Project:       Cybera Server & Workstation (http://cybera.sf.net)          |
'|  Description:   Contains Windows special folder retreaval api calls & funcs.|
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

' $Id: modSpecialFolders.bas 15 2006-01-01 16:12:35Z sam $

Option Explicit

' Version 5.0. The file system directory that is used to store administrative tools for an individual user.
'   The Microsoft Management Console (MMC) will save customized consoles to this directory, and it will roam with the user.
Public Const CSIDL_ADMINTOOLS = &H30
' The file system directory that corresponds to the user's nonlocalized Startup program group.
Public Const CSIDL_ALTSTARTUP = &H1D
' Version 4.71. The file system directory that serves as a common repository for application-specific data.
'   A typical path is C:\Documents and Settings\username\Application Data. This CSIDL is supported by the redistributable
'   Shfolder.dll for systems that do not have the Microsoft® Internet Explorer 4.0 integrated Shell installed.
Public Const CSIDL_APPDATA = &H1A
' The virtual folder containing the objects in the user's Recycle Bin.
Public Const CSIDL_BITBUCKET = &HA
' Version 6.0. The file system directory acting as a staging area for files waiting to be written to CD. A typical path
'   is C:\Documents and Settings\username\Local Settings\Application Data\Microsoft\CD Burning.
Public Const CSIDL_CDBURN_AREA = &H3B
'  Version 5.0. The file system directory containing administrative tools for all users of the computer.
Public Const CSIDL_COMMON_ADMINTOOLS = &H2F
' The file system directory that corresponds to the nonlocalized Startup program group for all users. Valid only for Microsoft Windows NT® systems.
Public Const CSIDL_COMMON_ALTSTARTUP = &H1E
' Version 5.0. The file system directory containing application data for all users. A typical path is C:\Documents and Settings\All Users\Application Data.
Public Const CSIDL_COMMON_APPDATA = &H23
' The file system directory that contains files and folders that appear on the desktop for all users. A typical path is C:\Documents and Settings\All Users\Desktop.
'   Valid only for Windows NT systems.
Public Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19
' The file system directory that contains documents that are common to all users. A typical paths is C:\Documents and Settings\All Users\Documents.
'   Valid for Windows NT systems and Microsoft Windows® 95 and Windows 98 systems with Shfolder.dll installed.
Public Const CSIDL_COMMON_DOCUMENTS = &H2E
' The file system directory that serves as a common repository for favorite items common to all users. Valid only for Windows NT systems.
Public Const CSIDL_COMMON_FAVORITES = &H1F
' Version 6.0. The file system directory that serves as a repository for music files common to all users. A typical path is C:\Documents and Settings\All Users\Documents\My Music.
Public Const CSIDL_COMMON_MUSIC = &H35
' Version 6.0. The file system directory that serves as a repository for image files common to all users. A typical path is C:\Documents and Settings\All Users\Documents\My Pictures.
Public Const CSIDL_COMMON_PICTURES = &H36
' The file system directory that contains the directories for the common program groups that appear on the Start menu for all users. A typical path is C:\Documents and Settings\All Users\Start Menu\Programs.
'   Valid only for Windows NT systems.
Public Const CSIDL_COMMON_PROGRAMS = &H17
' The file system directory that contains the programs and folders that appear on the Start menu for all users. A typical path is C:\Documents and Settings\All Users\Start Menu. Valid only for Windows NT systems.
Public Const CSIDL_COMMON_STARTMENU = &H16
' The file system directory that contains the programs that appear in the Startup folder for all users. A typical path is C:\Documents and Settings\All Users\Start Menu\Programs\Startup. Valid only for Windows NT systems.
Public Const CSIDL_COMMON_STARTUP = &H18
' The file system directory that contains the templates that are available to all users. A typical path is C:\Documents and Settings\All Users\Templates. Valid only for Windows NT systems.
Public Const CSIDL_COMMON_TEMPLATES = &H2D
' Version 6.0. The file system directory that serves as a repository for video files common to all users. A typical path is C:\Documents and Settings\All Users\Documents\My Videos.
Public Const CSIDL_COMMON_VIDEO = &H37
' The virtual folder containing icons for the Control Panel applications.
Public Const CSIDL_CONTROLS = &H3
' The file system directory that serves as a common repository for Internet cookies. A typical path is C:\Documents and Settings\username\Cookies.
Public Const CSIDL_COOKIES = &H21
' The virtual folder representing the Windows desktop, the root of the namespace.
Public Const CSIDL_DESKTOP = &H0
' The file system directory used to physically store file objects on the desktop (not to be confused with the desktop folder itself). A typical path is C:\Documents and Settings\username\Desktop.
Public Const CSIDL_DESKTOPDIRECTORY = &H10
' The virtual folder representing My Computer, containing everything on the local computer: storage devices, printers, and Control Panel. The folder may also contain mapped network drives.
Public Const CSIDL_DRIVES = &H11
' The file system directory that serves as a common repository for the user's favorite items. A typical path is C:\Documents and Settings\username\Favorites.
Public Const CSIDL_FAVORITES = &H6
' A virtual folder containing fonts. A typical path is C:\Windows\Fonts.
Public Const CSIDL_FONTS = &H14
' The file system directory that serves as a common repository for Internet history items.
Public Const CSIDL_HISTORY = &H22
' A virtual folder representing the Internet.
Public Const CSIDL_INTERNET = &H1
' Version 4.72. The file system directory that serves as a common repository for temporary Internet files. A typical path is C:\Documents and Settings\username\Local Settings\Temporary Internet Files.
Public Const CSIDL_INTERNET_CACHE = &H20
' Version 5.0. The file system directory that serves as a data repository for local (nonroaming) applications. A typical path is C:\Documents and Settings\username\Local Settings\Application Data.
Public Const CSIDL_LOCAL_APPDATA = &H1C
' Version 6.0. The virtual folder representing the My Documents desktop item. This should not be confused with CSIDL_PERSONAL, which represents the file system folder that physically stores the documents.
Public Const CSIDL_MYDOCUMENTS = &HC
' The file system directory that serves as a common repository for music files. A typical path is C:\Documents and Settings\User\My Documents\My Music.
Public Const CSIDL_MYMUSIC = &HD
' Version 5.0. The file system directory that serves as a common repository for image files. A typical path is C:\Documents and Settings\username\My Documents\My Pictures.
Public Const CSIDL_MYPICTURES = &H27
' Version 6.0. The file system directory that serves as a common repository for video files. A typical path is C:\Documents and Settings\username\My Documents\My Videos.
Public Const CSIDL_MYVIDEO = &HE
' A file system directory containing the link objects that may exist in the My Network Places virtual folder. It is not the same as CSIDL_NETWORK, which represents the network namespace root. A typical path is C:\Documents and Settings\username\NetHood.
Public Const CSIDL_NETHOOD = &H13
' A virtual folder representing Network Neighborhood, the root of the network namespace hierarchy.
Public Const CSIDL_NETWORK = &H12
' The file system directory used to physically store a user's common repository of documents. A typical path is C:\Documents and Settings\username\My Documents. This should be distinguished from the virtual My Documents folder in the namespace, identified by CSIDL_MYDOCUMENTS.
'   To access that virtual folder, use SHGetFolderLocation, which returns the ITEMIDLIST for the virtual location, or refer to the technique described in Managing the File System.
Public Const CSIDL_PERSONAL = &H5
' The virtual folder containing installed printers.
Public Const CSIDL_PRINTERS = &H4
' The file system directory that contains the link objects that can exist in the Printers virtual folder. A typical path is C:\Documents and Settings\username\PrintHood.
Public Const CSIDL_PRINTHOOD = &H1B
' Version 5.0. The user's profile folder. A typical path is C:\Documents and Settings\username. Applications should not create files or folders at this level; they should put their data under the locations referred to by CSIDL_APPDATA or CSIDL_LOCAL_APPDATA.
Public Const CSIDL_PROFILE = &H28
' Version 6.0. The file system directory containing user profile folders. A typical path is C:\Documents and Settings.
Public Const CSIDL_PROFILES = &H3E
' Version 5.0. The Program Files folder. A typical path is C:\Program Files.
Public Const CSIDL_PROGRAM_FILES = &H26
' Version 5.0. A folder for components that are shared across applications. A typical path is C:\Program Files\Common. Valid only for Windows NT, Windows 2000, and Windows XP systems. Not valid for Windows Millennium Edition (Windows Me).
Public Const CSIDL_PROGRAM_FILES_COMMON = &H2B
' The file system directory that contains the user's program groups (which are themselves file system directories). A typical path is C:\Documents and Settings\username\Start Menu\Programs.
Public Const CSIDL_PROGRAMS = &H2
' The file system directory that contains shortcuts to the user's most recently used documents. A typical path is C:\Documents and Settings\username\My Recent Documents. To create a shortcut in this folder, use SHAddToRecentDocs. In addition to creating the shortcut, this function updates the Shell's list of recent documents and adds the shortcut to the My Recent Documents submenu of the Start menu.
Public Const CSIDL_RECENT = &H8
' The file system directory that contains Send To menu items. A typical path is C:\Documents and Settings\username\SendTo.
Public Const CSIDL_SENDTO = &H9
' The file system directory containing Start menu items. A typical path is C:\Documents and Settings\username\Start Menu.
Public Const CSIDL_STARTMENU = &HB
' The file system directory that corresponds to the user's Startup program group. The system starts these programs whenever any user logs onto Windows NT or starts Windows 95. A typical path is C:\Documents and Settings\username\Start Menu\Programs\Startup.
Public Const CSIDL_STARTUP = &H7
' Version 5.0. The Windows System folder. A typical path is C:\Windows\System32.
Public Const CSIDL_SYSTEM = &H25
' The file system directory that serves as a common repository for document templates. A typical path is C:\Documents and Settings\username\Templates.
Public Const CSIDL_TEMPLATES = &H15
' Version 5.0. The Windows directory or SYSROOT. This corresponds to the %windir% or %SYSTEMROOT% environment variables. A typical path is C:\Windows.
Public Const CSIDL_WINDOWS = &H24

' Defines an item identifier.
Public Type typSHITEMID
    cb As Long
    abID As Byte
End Type

' Contains a list of item identifiers.
Public Type typITEMIDLIST
    mkid As typSHITEMID
End Type

' The ShellExecute function opens or prints a specified file. The file can be an executable file or a document file.
'Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
' Converts an item identifier list to a file system path.
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
' Retrieves a pointer to the ITEMIDLIST structure of a special folder.
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As typITEMIDLIST) As Long


'*******************************************************************************
Public Function funGetFolderPath(CSIDL As Long) As String
Dim uIDL As typITEMIDLIST
Dim strtemp As String * 512
On Error Resume Next

    Call SHGetSpecialFolderLocation(0, CSIDL, uIDL)
    Call SHGetPathFromIDList(uIDL.mkid.cb, strtemp)
    funGetFolderPath = left$(strtemp, InStr(strtemp, Chr$(0)) - 1)
End Function
