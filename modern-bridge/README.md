# Cybera Modern Bridge

This is a .NET 8 HTTP API that runs on a current Windows server. It is not deployed to, or used by, Windows XP workstations. It deliberately does **not** connect to Cybera's XP workstation TCP listener.

## Why the relay boundary exists

Cybera 1.9.rc2's server accepts a proprietary, plaintext Winsock protocol intended for its VB6 workstations. The inspected server sends sensitive configuration values over that connection and treats connecting clients as workstations. A web service must not impersonate one: doing so could consume a workstation slot or change state.

The bridge therefore has two intentionally narrow integration paths:

1. `GET /api/v1/status` reports whether `cybserv.exe` is running on the same host and can include an operator-provided, read-only JSON snapshot. It never opens a connection to the legacy client port.
2. Session commands go only to a **separate loopback relay** on `127.0.0.1`. The relay is disabled by default. A future reviewed, host-local adapter may translate the small allowlisted command set into a safe automation/integration path. It must not expose the VB6 protocol to the LAN.

This is the safe phase-in point: run the bridge beside the existing VB6 server, first use health and status, then enable a local relay only after testing it against a non-production Cybera installation. No XP client needs a browser, TLS stack, or .NET runtime.

## Combined legacy-server and bridge launch

The installed `Start-Cybera-Lounge.ps1` launcher can start the legacy server and then the bridge in one action, but only after the external configuration explicitly identifies the installed legacy executable. It never searches for, guesses, or launches a legacy executable from `third_party`.

Add the real installed paths to the ProgramData configuration; the defaults are intentionally empty:

```json
"LegacyServer": {
  "ExecutablePath": "C:\\Path\\To\\Installed\\cybserv.exe",
  "WorkingDirectory": "C:\\Path\\To\\Installed",
  "StartupWaitSeconds": 8
}
```

The launcher checks for a running `LegacyProcessName` first. If it is absent, it validates that the configured executable exists, is an `.exe`, and has a valid configured working directory; it starts the legacy application visibly, waits 1–60 seconds, confirms the process is still running, and then starts the bridge hidden. It never connects to the legacy workstation protocol. It also refuses to start the bridge without `CYBERA_BRIDGE_API_KEY` in the launcher process environment.

## Web console

Opening the bridge root (`http://server:port/`) serves the responsive operator console. It is subject to the same LAN CIDR filter as the API, but the protected status and command calls still require an API key.

Enter the key in the console's **Connect API key** dialog. It is kept only in that browser tab's `sessionStorage`, is sent in the `X-Api-Key` header, and is never embedded in the page, configuration, or repository. Close the tab or use **Forget key** to remove it.

The dashboard never invents occupancy, sessions, or money values. It reports the local `cybserv` process state and labels workstation details as unavailable unless `LegacyStatusFile` contains a trusted exporter snapshot. An optional snapshot can expose a `workstations` array such as:

```json
{
  "observedAtUtc": "2026-11-07T00:00:00Z",
  "workstations": [
    { "id": "WS-01", "name": "Front desk 1", "state": "in-use", "session": { "id": "123", "state": "active" } }
  ]
}
```

The file must be generated outside this project, read-only to the bridge, and atomically replaced by a trusted local exporter. An absent `workstations` array is shown as **not integrated**, not zero.

Session controls require a valid operator key, a workstation ID, and an explicitly enabled loopback relay. `pause` and `end` are the currently supported allowlisted commands; **Extend** remains visibly unavailable until a safe, tested relay contract exists.

## Run locally

PowerShell:

```powershell
$env:CYBERA_BRIDGE_API_KEY = 'replace-with-a-long-random-secret'
dotnet run --project .\modern-bridge --urls http://127.0.0.1:5088
```

For a machine-specific non-secret configuration file, set `CYBERA_BRIDGE_CONFIG` to an absolute JSON path before starting. That file is loaded after the bundled `appsettings.json`; the installer uses this mechanism to load the ProgramData configuration. See [the installer guide](../installer/INSTALLER.md) for packaging and Windows Server deployment.

The sample configuration permits loopback and `192.168.1.0/24`; edit `Bridge:AllowedNetworks` for the actual management subnet before binding to a LAN address. The bridge rejects all other remote IPs and does not trust forwarded headers. Put it behind TLS/reverse-proxy protection if it must be reached beyond a protected management VLAN.

## API

`GET /health` is network-restricted but unauthenticated and returns only bridge liveness.

All `/api/v1` endpoints require `X-Api-Key` exactly matching `CYBERA_BRIDGE_API_KEY` in the bridge process environment. The API-key variable name is fixed and is not configurable in JSON, preventing a key value from accidentally being treated as an environment-variable name.

| Method | Endpoint | Result |
| --- | --- | --- |
| `GET` | `/api/v1/status` | Local `cybserv` process state plus an optional JSON snapshot. |
| `POST` | `/api/v1/sessions/{workstationId}/commands` | Allowlisted `end`, `pause`, or `resume` request to the local relay. Requires `Idempotency-Key` (16–128 chars). |

Example status request:

```powershell
Invoke-RestMethod http://127.0.0.1:5088/api/v1/status -Headers @{ 'X-Api-Key' = $env:CYBERA_BRIDGE_API_KEY }
```

Example command request (returns `503` until an explicit local relay is enabled):

```powershell
Invoke-RestMethod http://127.0.0.1:5088/api/v1/sessions/WS-01/commands -Method Post -ContentType application/json -Headers @{ 'X-Api-Key' = $env:CYBERA_BRIDGE_API_KEY; 'Idempotency-Key' = 'a-long-unique-request-key-0001' } -Body '{"action":"pause"}'
```

## Constraints and deployment rules

- Bind only to a management VLAN; firewall the port to that VLAN. Do not expose this API or Cybera's legacy TCP port to the internet.
- Use a 32+ character randomly generated API key stored outside the repository. Rotate it on operator changes.
- Do not point `LegacyStatusFile` at the Jet database or write to the `.mdb` while the VB6 server runs. A future exporter should atomically replace a JSON file in a bridge-owned directory.
- The loopback relay is a deliberate capability boundary. It accepts only newline-delimited JSON `RelayCommand` records and must independently authenticate/authorize and audit commands before touching Cybera.
- The original Cybera source remains under `third_party/original-source/` and is unmodified. Its GPLv2 notice is at `third_party/original-source/cybera-1.9-rc2/extracted/cybera_1.9.rc2_full_source/license.txt`.

## Inspection notes

The design follows the original VB6 source: `cybserv/frmMain.frm` owns the Winsock listener; `cybserv/modFunctions.bas` parses `CNX_*` messages and issues workstation commands. This bridge copies no VB6 source and treats the legacy protocol as untrusted/unsafe for direct web exposure.
