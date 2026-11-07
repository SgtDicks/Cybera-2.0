# _____________________________________________________________________________
#|________________________________WORKSTATION INSTALLER SCRIPT ________________|
#|  Name:          cybcli.nsi                                                  |
#|  Type:          Nulsoft Instaler Script                                     |
#|  Project:       Cybera (http://cybera.sf.net)                               |
#|  Description:   Contains the Workstation installer generation script.       |
#|  Copyright:     (C) 2004-2006 Samuel Monsarrat (smonsarr@users.sf.net)      |
#|                                                                             |
#| This file is part of Cybera.                                                |
#|                                                                             |
#| Cybera is free software; you can redistribute it and/or modify it under the |
#| terms of the GNU General Public License as published by the Free Software   |
#| Foundation; either version 2 of the License, or (at your option) any later  |
#| version.                                                                    |
#|                                                                             |
#| Cybera is distributed in the hope that it will be useful, but WITHOUT ANY   |
#| WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS   |
#| FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more      |
#| details.                                                                    |
#|                                                                             |
#| You should have received a copy of the GNU General Public License along     |
#| with Cybera; if not, write to the Free Software Foundation, Inc.,           |
#| 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA                     |
#|                                                                             |
#!_____________________________________________________________________________!

# $Id: cybcli.nsi 41 2006-02-21 18:14:29Z sam $

SetCompressor /SOLID lzma
XPStyle on

!include "Library.nsh"
Var ALREADY_INSTALLED

!define /date NOW "%y%m%d-%H%M"
!define PRODUCT_NAME "Cybera Client"
!define PRODUCT_EXE_NAME "cybera-client"
!define INSTALLER_VERSION_NUM "1.9.0.122"
!define PRODUCT_VERSION "1.9.rc2"
!define PRODUCT_PUBLISHER "Samuel Monsarrat"
!define PRODUCT_COPYRIGHT_YEAR "© 2004-2006"
!define PRODUCT_INSTALLER_FILENAME ${PRODUCT_EXE_NAME}_${PRODUCT_VERSION}_${NOW}.exe

!define PRODUCT_DIR_REGKEY "Software\Microsoft\Windows\CurrentVersion\App Paths\cybcli.exe"
!define PRODUCT_UNINST_KEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}"
!define PRODUCT_UNINST_ROOT_KEY "HKLM"
!define PRODUCT_STARTMENU_REGVAL "NSIS:StartMenuDir"

!define VBFILESDIR "..\system" ;Location of VB6 runtime files
!define BINFILESDIR "..\..\..\Release\CybCli" ;Location of Cybera binary files
!define RESFILESDIR "..\resources" ;Location of Cybera resource files

; MUI 1.67 compatible ------
!include "MUI.nsh"

; MUI Settings
!define MUI_ABORTWARNING
!define MUI_ICON "${RESFILESDIR}\cybcli_inst.ico"
!define MUI_UNICON "${RESFILESDIR}\cybcli_uninst.ico"
!define MUI_WELCOMEFINISHPAGE_BITMAP "${RESFILESDIR}\cybcli_inst.bmp"

; Language Selection Dialog Settings
!define MUI_LANGDLL_REGISTRY_ROOT "${PRODUCT_UNINST_ROOT_KEY}"
!define MUI_LANGDLL_REGISTRY_KEY "${PRODUCT_UNINST_KEY}"
!define MUI_LANGDLL_REGISTRY_VALUENAME "NSIS:Language"

; Welcome page
!insertmacro MUI_PAGE_WELCOME

; License page
!insertmacro MUI_PAGE_LICENSE "${RESFILESDIR}\end_user_license.txt"

; Directory page
!insertmacro MUI_PAGE_DIRECTORY

; Start menu page
var ICONS_GROUP
!define MUI_STARTMENUPAGE_NODISABLE
!define MUI_STARTMENUPAGE_DEFAULTFOLDER "${PRODUCT_NAME}"
!define MUI_STARTMENUPAGE_REGISTRY_ROOT "${PRODUCT_UNINST_ROOT_KEY}"
!define MUI_STARTMENUPAGE_REGISTRY_KEY "${PRODUCT_UNINST_KEY}"
!define MUI_STARTMENUPAGE_REGISTRY_VALUENAME "${PRODUCT_STARTMENU_REGVAL}"
!insertmacro MUI_PAGE_STARTMENU Application $ICONS_GROUP

; Instfiles page
!insertmacro MUI_PAGE_INSTFILES

; Finish page
!insertmacro MUI_PAGE_FINISH

; Uninstaller pages
!insertmacro MUI_UNPAGE_INSTFILES

; Language files
!insertmacro MUI_LANGUAGE "English"
!insertmacro MUI_LANGUAGE "French"
!insertmacro MUI_LANGUAGE "German"
!insertmacro MUI_LANGUAGE "Greek"
!insertmacro MUI_LANGUAGE "Hungarian"
!insertmacro MUI_LANGUAGE "Italian"
!insertmacro MUI_LANGUAGE "PortugueseBR"
!insertmacro MUI_LANGUAGE "Spanish"
;!insertmacro MUI_LANGUAGE "Turkish"

; Reserve files
!insertmacro MUI_RESERVEFILE_INSTALLOPTIONS
!insertmacro MUI_RESERVEFILE_LANGDLL
; MUI end ------

Name "${PRODUCT_NAME}"
OutFile "..\..\..\Setup\${PRODUCT_INSTALLER_FILENAME}"
InstallDir "$PROGRAMFILES\${PRODUCT_NAME}"
InstallDirRegKey HKLM "${PRODUCT_DIR_REGKEY}" ""
ShowInstDetails hide
ShowUnInstDetails show
BrandingText "${PRODUCT_NAME} ${PRODUCT_VERSION} - built on ${__DATE__} at ${__TIME__}"

VIAddVersionKey /LANG=${LANG_ENGLISH} "ProductName"      "${PRODUCT_NAME}"
VIAddVersionKey /LANG=${LANG_ENGLISH} "ProductVersion"   "${PRODUCT_VERSION}"
VIAddVersionKey /LANG=${LANG_ENGLISH} "LegalCopyright"   "${PRODUCT_COPYRIGHT_YEAR} ${PRODUCT_PUBLISHER}"
VIAddVersionKey /LANG=${LANG_ENGLISH} "FileDescription"  "${PRODUCT_NAME} ${PRODUCT_VERSION} installer"
VIAddVersionKey /LANG=${LANG_ENGLISH} "FileVersion"      "${INSTALLER_VERSION_NUM}"
VIAddVersionKey /LANG=${LANG_ENGLISH} "OriginalFilename" "${PRODUCT_INSTALLER_FILENAME}"
VIProductVersion "${INSTALLER_VERSION_NUM}"

Function .onInit
  !insertmacro MUI_LANGDLL_DISPLAY
FunctionEnd

Section "MainSection" SEC01
  SetShellVarContext all
  SetOverwrite ifnewer

  IfFileExists "$INSTDIR\cybcli.exe" 0 new_installation
    StrCpy $ALREADY_INSTALLED 1
  new_installation:
  !insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_NOTPROTECTED "${VBFILESDIR}\msvbvm60.dll" "$SYSDIR\msvbvm60.dll" "$SYSDIR"
  !insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED    "${VBFILESDIR}\oleaut32.dll" "$SYSDIR\oleaut32.dll" "$SYSDIR"
  !insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED    "${VBFILESDIR}\olepro32.dll" "$SYSDIR\olepro32.dll" "$SYSDIR"
  !insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED    "${VBFILESDIR}\comcat.dll"   "$SYSDIR\comcat.dll"   "$SYSDIR"
  !insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED    "${VBFILESDIR}\comdlg32.ocx" "$SYSDIR\comdlg32.ocx" "$SYSDIR"
  !insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED    "${VBFILESDIR}\mscomctl.ocx" "$SYSDIR\mscomctl.ocx" "$SYSDIR"
  !insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED    "${VBFILESDIR}\mswinsck.ocx" "$SYSDIR\mswinsck.ocx" "$SYSDIR"
  !insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED    "${VBFILESDIR}\richtx32.ocx" "$SYSDIR\richtx32.ocx" "$SYSDIR"
  !insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED    "${VBFILESDIR}\sysinfo.ocx"  "$SYSDIR\sysinfo.ocx"  "$SYSDIR"
  !insertmacro InstallLib DLL    $ALREADY_INSTALLED REBOOT_PROTECTED    "${VBFILESDIR}\asycfilt.dll" "$SYSDIR\asycfilt.dll" "$SYSDIR"
  !insertmacro InstallLib DLL    $ALREADY_INSTALLED REBOOT_PROTECTED    "${VBFILESDIR}\psapi.dll"    "$SYSDIR\psapi.dll"    "$SYSDIR"
  !insertmacro InstallLib TLB    $ALREADY_INSTALLED REBOOT_PROTECTED    "${VBFILESDIR}\stdole2.tlb"  "$SYSDIR\stdole2.tlb"  "$SYSDIR"

  SetOutPath "$INSTDIR"

  !insertmacro InstallLib REGDLL $ALREADY_INSTALLED NOREBOOT_NOTPROTECTED "${BINFILESDIR}\cybcli_fr.dll" "$INSTDIR\cybcli_fr.dll" "$INSTDIR"
  !insertmacro InstallLib REGDLL $ALREADY_INSTALLED NOREBOOT_NOTPROTECTED "${BINFILESDIR}\cybcli_es.dll" "$INSTDIR\cybcli_es.dll" "$INSTDIR"
  !insertmacro InstallLib REGDLL $ALREADY_INSTALLED NOREBOOT_NOTPROTECTED "${BINFILESDIR}\cybcli_de.dll" "$INSTDIR\cybcli_de.dll" "$INSTDIR"
  !insertmacro InstallLib REGDLL $ALREADY_INSTALLED NOREBOOT_NOTPROTECTED "${BINFILESDIR}\cybcli_el.dll" "$INSTDIR\cybcli_el.dll" "$INSTDIR"
  !insertmacro InstallLib REGDLL $ALREADY_INSTALLED NOREBOOT_NOTPROTECTED "${BINFILESDIR}\cybcli_it.dll" "$INSTDIR\cybcli_it.dll" "$INSTDIR"
  !insertmacro InstallLib REGDLL $ALREADY_INSTALLED NOREBOOT_NOTPROTECTED "${BINFILESDIR}\cybcli_br.dll" "$INSTDIR\cybcli_br.dll" "$INSTDIR"
  !insertmacro InstallLib REGDLL $ALREADY_INSTALLED NOREBOOT_NOTPROTECTED "${BINFILESDIR}\cybcli_hu.dll" "$INSTDIR\cybcli_hu.dll" "$INSTDIR"

  !insertmacro InstallLib DLL $ALREADY_INSTALLED NOREBOOT_NOTPROTECTED "${VBFILESDIR}\CybWinLock.dll" "$INSTDIR\CybWinLock.dll" "$INSTDIR"
  File ${VBFILESDIR}\CybWinLock.dll

  File "${BINFILESDIR}\cybcli.exe"
  File "/oname=license.txt" "${RESFILESDIR}\end_user_license.txt"

  CreateDirectory "$SMPROGRAMS\$ICONS_GROUP"
  CreateShortCut "$SMPROGRAMS\$ICONS_GROUP\${PRODUCT_NAME}.lnk" "$INSTDIR\cybcli.exe"
  CreateShortCut "$SMPROGRAMS\$ICONS_GROUP\License.lnk" "$INSTDIR\license.txt"

  SetOverwrite off
  SetOutPath "$APPDATA\${PRODUCT_NAME}"

  File "${RESFILESDIR}\cybcli.log"
  AccessControl::GrantOnFile "$APPDATA\${PRODUCT_NAME}\cybcli.log" "(S-1-5-32-545)" "FullAccess"
  AccessControl::GrantOnFile "$APPDATA\${PRODUCT_NAME}\" "(S-1-5-32-545)" "FullAccess"

  File "${RESFILESDIR}\logo.gif"

  File "${RESFILESDIR}\de_blocked.rtf"
  File "${RESFILESDIR}\de_intro.rtf"
  File "${RESFILESDIR}\de_nocredit.rtf"
  File "${RESFILESDIR}\de_sessionend.rtf"
  File "${RESFILESDIR}\de_sessionpaused.rtf"

  File "${RESFILESDIR}\el_blocked.rtf"
  File "${RESFILESDIR}\el_intro.rtf"
  File "${RESFILESDIR}\el_nocredit.rtf"
  File "${RESFILESDIR}\el_sessionend.rtf"
  File "${RESFILESDIR}\el_sessionpaused.rtf"

  File "${RESFILESDIR}\en_blocked.rtf"
  File "${RESFILESDIR}\en_intro.rtf"
  File "${RESFILESDIR}\en_nocredit.rtf"
  File "${RESFILESDIR}\en_sessionend.rtf"
  File "${RESFILESDIR}\en_sessionpaused.rtf"

  File "${RESFILESDIR}\es_blocked.rtf"
  File "${RESFILESDIR}\es_intro.rtf"
  File "${RESFILESDIR}\es_nocredit.rtf"
  File "${RESFILESDIR}\es_sessionend.rtf"
  File "${RESFILESDIR}\es_sessionpaused.rtf"

  File "${RESFILESDIR}\fr_blocked.rtf"
  File "${RESFILESDIR}\fr_intro.rtf"
  File "${RESFILESDIR}\fr_nocredit.rtf"
  File "${RESFILESDIR}\fr_sessionend.rtf"
  File "${RESFILESDIR}\fr_sessionpaused.rtf"

  File "${RESFILESDIR}\it_blocked.rtf"
  File "${RESFILESDIR}\it_intro.rtf"
  File "${RESFILESDIR}\it_nocredit.rtf"
  File "${RESFILESDIR}\it_sessionend.rtf"
  File "${RESFILESDIR}\it_sessionpaused.rtf"

  File "${RESFILESDIR}\br_blocked.rtf"
  File "${RESFILESDIR}\br_intro.rtf"
  File "${RESFILESDIR}\br_nocredit.rtf"
  File "${RESFILESDIR}\br_sessionend.rtf"
  File "${RESFILESDIR}\br_sessionpaused.rtf"

  File "${RESFILESDIR}\hu_blocked.rtf"
  File "${RESFILESDIR}\hu_intro.rtf"
  File "${RESFILESDIR}\hu_nocredit.rtf"
  File "${RESFILESDIR}\hu_sessionend.rtf"
  File "${RESFILESDIR}\hu_sessionpaused.rtf"

  WriteRegStr HKLM "Software\${PRODUCT_NAME}" "" ""
  AccessControl::GrantOnRegKey HKLM "Software\${PRODUCT_NAME}" "(S-1-5-32-545)" "FullAccess"

  #Dissable the XP Welcome screen (back to a w2k login type)
  WriteRegDWORD HKLM "Software\Microsoft\Windows NT\CurrentVersion\Winlogon" "LogonType" "0"

SectionEnd

Section -AdditionalIcons
  SetShellVarContext current
  CreateDirectory "$SMPROGRAMS\$ICONS_GROUP"
  CreateShortCut "$SMPROGRAMS\$ICONS_GROUP\Uninstall.lnk" "$INSTDIR\uninst.exe"
  SetShellVarContext all
SectionEnd

Section -Post
  WriteUninstaller "$INSTDIR\uninst.exe"
  WriteRegStr HKLM "${PRODUCT_DIR_REGKEY}" "" "$INSTDIR\cybcli.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayName" "$(^Name)"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "UninstallString" "$INSTDIR\uninst.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayIcon" "$INSTDIR\cybcli.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayVersion" "${PRODUCT_VERSION}"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "${PRODUCT_STARTMENU_REGVAL}" "$ICONS_GROUP"
SectionEnd

Function un.onUninstSuccess
  HideWindow
  MessageBox MB_ICONINFORMATION|MB_OK "$(^Name) was successfully removed from your computer."
FunctionEnd

Function un.onInit
!insertmacro MUI_UNGETLANGUAGE
  MessageBox MB_ICONQUESTION|MB_YESNO|MB_DEFBUTTON2 "Are you sure you want to completely remove $(^Name) and all of its components?" IDYES +2
  Abort
FunctionEnd

Section Uninstall
  SetShellVarContext all

  !insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\msvbvm60.dll"
  !insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\oleaut32.dll"
  !insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\olepro32.dll"
  !insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\comcat.dll"
  !insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\comdlg32.ocx"
  !insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\mscomctl.ocx"
  !insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\mswinsck.ocx"
  !insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\richtx32.ocx"
  !insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\sysinfo.ocx"
  !insertmacro UnInstallLib DLL    SHARED NOREMOVE "$SYSDIR\asycfilt.dll"
  !insertmacro UnInstallLib DLL    SHARED NOREMOVE "$SYSDIR\psapi.dll"
  !insertmacro UnInstallLib TLB    SHARED NOREMOVE "$SYSDIR\stdole2.tlb"

  !insertmacro UnInstallLib DLL NOTSHARED NOREBOOT_NOTPROTECTED "$INSTDIR\cybcli_fr.dll"
  !insertmacro UnInstallLib DLL NOTSHARED NOREBOOT_NOTPROTECTED "$INSTDIR\cybcli_es.dll"
  !insertmacro UnInstallLib DLL NOTSHARED NOREBOOT_NOTPROTECTED "$INSTDIR\cybcli_de.dll"
  !insertmacro UnInstallLib DLL NOTSHARED NOREBOOT_NOTPROTECTED "$INSTDIR\cybcli_el.dll"
  !insertmacro UnInstallLib DLL NOTSHARED NOREBOOT_NOTPROTECTED "$INSTDIR\cybcli_it.dll"
  !insertmacro UnInstallLib DLL NOTSHARED NOREBOOT_NOTPROTECTED "$INSTDIR\cybcli_br.dll"
  !insertmacro UnInstallLib DLL NOTSHARED NOREBOOT_NOTPROTECTED "$INSTDIR\cybcli_hu.dll"

  ReadRegStr $ICONS_GROUP ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "${PRODUCT_STARTMENU_REGVAL}"
  Delete "$INSTDIR\uninst.exe"
  Delete "$INSTDIR\cybcli.exe"
  Delete "$INSTDIR\license.txt"
  Delete "$INSTDIR\CybWinLock.dll"
  RMDir "$INSTDIR"

  RMDir /r "$SMPROGRAMS\$ICONS_GROUP"
  SetShellVarContext current
  RMDir /r "$SMPROGRAMS\$ICONS_GROUP"
  SetShellVarContext all

  RMDir /r "$APPDATA\${PRODUCT_NAME}"

  DeleteRegKey HKLM "Software\${PRODUCT_NAME}"
  DeleteRegKey HKCU "Software\${PRODUCT_NAME}"

  #Enable the XP Welcome screen again
  WriteRegDWORD HKLM "Software\Microsoft\Windows NT\CurrentVersion\Winlogon" "LogonType" "1"

  DeleteRegKey ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}"
  DeleteRegKey HKLM "${PRODUCT_DIR_REGKEY}"
  SetAutoClose true
SectionEnd