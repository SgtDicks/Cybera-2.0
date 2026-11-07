# Host-local snapshot adapter

Publish-CyberaSnapshot.ps1 is a working safe boundary: it validates a JSON export and atomically publishes it for the bridge. It never connects to the XP protocol, starts a client, or writes to the Jet database.

An exporter running beside the legacy server may produce a JSON object containing any of: workstations, sessions, accounts, cards, pricing, operators, journal, reports, and configuration. Publish it with:

    .\Publish-CyberaSnapshot.ps1 -InputPath C:\SecureExporter\snapshot.json -OutputPath C:\ProgramData\CyberaModernBridge\snapshot.json

Set Bridge:LegacyStatusFile to the output path. Build a future exporter against a copied/disposable Cybera database only, with read-only access; do not point this tool at the MDB.

Screen capture/following needs a separate isolated capture broker with explicit operator/session consent. XP registry lockdown remains a legacy-only function and is intentionally excluded from this adapter.

## Real Jet read-only exporter

Export-CyberaJetReadOnly.ps1 uses the registered Jet OLE DB provider with Mode=Read. It exports cards, clients as accounts, recharge pricing, users, journals, events, and configuration without modifying the MDB:

    $password = Read-Host -AsSecureString 'Cybera database password'
    .\Export-CyberaJetReadOnly.ps1 -DatabasePath C:\Path\To\cybera.mdb -DatabasePassword $password -OutputPath C:\ProgramData\CyberaModernBridge\snapshot.json

The database path is normally the legacy server's configured registry DB path; do not guess it and do not run this against a production copy until access is verified. Sessions/workstation connection state is in memory, not these tables, and still needs a host-local runtime exporter.
