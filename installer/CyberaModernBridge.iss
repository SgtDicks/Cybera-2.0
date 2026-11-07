; Cybera Modern Bridge installer source.
; Requires Inno Setup 6 or later. This package installs only the modern .NET bridge.
; It never deploys software to XP workstations or includes the archived Cybera source.

#ifndef MyAppVersion
  #define MyAppVersion "0.1.5"
#endif
#ifndef PublishDir
  #error "Pass /DPublishDir=<published bridge directory> to ISCC.exe."
#endif

#define MyAppName "Cybera Modern Bridge"
#define MyAppPublisher "LAN Legacy Lounge"
#define MyAppExeName "Cybera.ModernBridge.exe"

[Setup]
AppId={{23F44DC3-60F6-4515-AD7A-A57D6DB568C8}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\Cybera Modern Bridge
DefaultGroupName=Cybera Modern Bridge
DisableProgramGroupPage=yes
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=dialog
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=..\artifacts\installer
OutputBaseFilename=CyberaModernBridge-{#MyAppVersion}-win-x64-setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayName={#MyAppName}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Dirs]
Name: "{commonappdata}\CyberaModernBridge"

[Files]
Source: "{#PublishDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "appsettings.production.template.json"; DestDir: "{commonappdata}\CyberaModernBridge"; DestName: "appsettings.json"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "run-bridge.cmd"; DestDir: "{app}"; Flags: ignoreversion
Source: "Start-Cybera-Lounge.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\modern-bridge\README.md"; DestDir: "{app}\docs"; Flags: ignoreversion
Source: "..\modern-bridge\THIRD-PARTY-NOTICES.md"; DestDir: "{app}\docs"; Flags: ignoreversion
Source: "..\docs\FEATURE-PARITY.md"; DestDir: "{app}\docs"; Flags: ignoreversion
Source: "..\adapter\README.md"; DestDir: "{app}\docs"; Flags: ignoreversion
Source: "..\adapter\Publish-CyberaSnapshot.ps1"; DestDir: "{app}\adapter"; Flags: ignoreversion
Source: "..\adapter\Export-CyberaJetReadOnly.ps1"; DestDir: "{app}\adapter"; Flags: ignoreversion
Source: "INSTALLER.md"; DestDir: "{app}\docs"; Flags: ignoreversion

[Registry]
; The bridge reads this non-secret path at startup. The API key itself is never saved by the installer.
Root: HKLM; Subkey: "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"; ValueType: string; ValueName: "CYBERA_BRIDGE_CONFIG"; ValueData: "{commonappdata}\CyberaModernBridge\appsettings.json"; Flags: uninsdeletevalue

[Icons]
Name: "{autoprograms}\Cybera Modern Bridge\Start bridge"; Filename: "{app}\run-bridge.cmd"; WorkingDir: "{app}"
Name: "{autoprograms}\Cybera Modern Bridge\Start legacy Cybera + bridge"; Filename: "{sys}\WindowsPowerShell\v1.0\powershell.exe"; Parameters: "-NoProfile -File ""{app}\Start-Cybera-Lounge.ps1"""; WorkingDir: "{app}"
Name: "{autoprograms}\Cybera Modern Bridge\Bridge setup guide"; Filename: "{app}\docs\INSTALLER.md"
Name: "{autoprograms}\Cybera Modern Bridge\Web console (after start)"; Filename: "http://127.0.0.1:5088/"
Name: "{autoprograms}\Cybera Modern Bridge\Uninstall"; Filename: "{uninstallexe}"

[Run]
Filename: "{cmd}"; Parameters: "/C start """" ""{app}\docs\INSTALLER.md"""; Description: "Open the Bridge setup guide"; Flags: postinstall shellexec skipifsilent
