# Cybera Modern Bridge — Windows installer

This installer packages **only** the modern .NET 8 bridge and its web console for a current Windows server. It does not modify, deploy, or run anything on Windows XP workstation clients. It also excludes `third_party/original-source`.

## Build the installer

Prerequisites on the build machine:

- .NET 8 SDK;
- [Inno Setup 6 or later](https://jrsoftware.org/isinfo.php). The build script detects the standard Inno Setup 6 and 7 installation paths as well as `PATH`.

From the repository root:

```powershell
.\installer\build-installer.ps1 -Version 0.1.0
```

The script publishes a framework-dependent `win-x64` bridge to `artifacts\publish\win-x64`, then writes the installer to `artifacts\installer\CyberaModernBridge-0.1.0-win-x64-setup.exe`.

## Server prerequisites and installation

The target server needs the **.NET 8 Windows Server Hosting Bundle** (or compatible .NET 8 runtime). Run the generated installer as an administrator. It installs program files to:

```
C:\Program Files\Cybera Modern Bridge
```

Version 0.1.0 installers built before the 64-bit install-mode fix may have placed the bridge in `C:\Program Files (x86)\Cybera Modern Bridge`. Use that actual location to launch an existing installation, or uninstall and reinstall using the rebuilt setup EXE to place a fresh installation under `C:\Program Files`.

It provisions, without overwriting an existing configuration:

```
C:\ProgramData\CyberaModernBridge\appsettings.json
```

Edit that file before first start:

- replace the sample `192.168.1.0/24` with the actual management VLAN CIDR(s);
- leave `CommandRelay.Enabled` as `false` until a reviewed host-local relay has been tested;
- configure `LegacyStatusFile` only for an externally generated, read-only JSON snapshot—never point it at Cybera's Jet database.

The installer registers the non-secret `CYBERA_BRIDGE_CONFIG` path in the system environment. Open a new administrator console after installation so it sees the value.

Set the API key separately; it is intentionally not requested, generated, or stored by the installer. Provision a high-entropy plaintext API key through the dedicated service account's approved secret-management procedure, then expose it to that account as `CYBERA_BRIDGE_API_KEY`. This variable name is fixed; do not add an API key to `appsettings.json`. The bridge must not be started until that variable is available to its process.

For an interactive validation, set a temporary process-scoped value only:

```powershell
$env:CYBERA_BRIDGE_API_KEY = 'generate-and-paste-a-32-plus-character-random-value'
& 'C:\Program Files\Cybera Modern Bridge\Cybera.ModernBridge.exe' --urls http://127.0.0.1:5088
```

Use the Start Menu’s **Start bridge** entry only after configuring the key for the launching account. Browse to `http://127.0.0.1:5088/` or the server’s management-LAN address after firewalling the chosen port.

### Start the legacy Cybera server and bridge together

The installer adds **Cybera Modern Bridge → Start legacy Cybera + bridge**. Before using it, edit `C:\ProgramData\CyberaModernBridge\appsettings.json`; the default legacy server paths are empty on purpose:

```json
"LegacyServer": {
  "ExecutablePath": "C:\\Path\\To\\Installed\\cybserv.exe",
  "WorkingDirectory": "C:\\Path\\To\\Installed",
  "StartupWaitSeconds": 8
}
```

Set `ExecutablePath` to the already-installed legacy server executable and `WorkingDirectory` to its real installation directory. The launcher rejects missing paths, non-`.exe` paths, missing directories, invalid process names, and waits outside 1–60 seconds. It first checks whether the configured `LegacyProcessName` (`cybserv` by default) is already running. It never runs an executable from `third_party` and never installs anything on XP clients.

Start it from a PowerShell session that has the key for that session:

```powershell
$env:CYBERA_BRIDGE_API_KEY = [System.Net.NetworkCredential]::new('', (Read-Host -AsSecureString 'Bridge API key')).Password
& 'C:\Program Files\Cybera Modern Bridge\Start-Cybera-Lounge.ps1'
```

For a preview that starts nothing, append `-WhatIf`. The launcher starts the legacy GUI visibly, waits the configured time, confirms its process is still running, and starts the bridge hidden at `http://0.0.0.0:5088`. LAN access remains constrained by `AllowedNetworks`.

## Service deployment

The installer does not create a Windows service because the bridge requires a per-service-account secret and its own LAN binding decision. Run it under a dedicated, non-administrator service account using your approved service wrapper or Windows service deployment standard. That account must receive `CYBERA_BRIDGE_API_KEY` and read access to the ProgramData configuration; do not run it as LocalSystem by default.

The service command is:

```
"C:\Program Files\Cybera Modern Bridge\Cybera.ModernBridge.exe" --urls http://0.0.0.0:5088
```

Restrict the Windows Firewall rule to the management VLAN. Do not expose the bridge, the bridge console, or Cybera's legacy TCP port to the internet.

## Uninstall

Use **Apps & features** or the Start Menu uninstall entry. Program files, shortcuts, and the `CYBERA_BRIDGE_CONFIG` environment setting are removed. `C:\ProgramData\CyberaModernBridge\appsettings.json` is deliberately retained so upgrades do not erase operator configuration. Remove that directory and the service account’s API-key secret manually when decommissioning.
