public static class FeatureCatalog
{
    public static readonly object[] All =
    [
        new { area = "server-status", legacyEvidence = "cybserv/frmMain.frm", status = "implemented", mode = "local process + snapshot" },
        new { area = "workstations-sessions", legacyEvidence = "clsWorkstation, clsSession", status = "adapter-needed", mode = "read-only snapshot; allowlisted relay actions" },
        new { area = "accounts-clients", legacyEvidence = "clsClient, frmClientConfig", status = "adapter-needed", mode = "snapshot + local relay" },
        new { area = "cards-recharges", legacyEvidence = "clsCards, frmRecharge, frmCardMassAdd", status = "adapter-needed", mode = "snapshot + local relay" },
        new { area = "pricing-self-service", legacyEvidence = "frmOptions, CNX_POSTPAID_MODE", status = "adapter-needed", mode = "snapshot + local relay" },
        new { area = "operators-roles", legacyEvidence = "frmUserAdmin, clsUser", status = "adapter-needed", mode = "snapshot only until role adapter exists" },
        new { area = "journals-reports-stats", legacyEvidence = "frmStatsPrint, database events", status = "adapter-needed", mode = "snapshot/export adapter" },
        new { area = "configuration-backup", legacyEvidence = "frmOptions, frmDBMaintenance", status = "safely-deferred", mode = "no direct Jet DB writes" },
        new { area = "workstation-actions", legacyEvidence = "CNX_LOCK/PAUSE/RESUME/SHUTDOWN", status = "adapter-needed", mode = "allowlisted local relay only" },
        new { area = "screen-replication-follow", legacyEvidence = "screen replication protocol", status = "safely-deferred", mode = "requires isolated capture adapter and consent review" },
        new { area = "xp-lockdown-remote-registry", legacyEvidence = "modFunctions.bas remote registry", status = "safely-deferred", mode = "legacy-only; never exposed through browser" }
    ];

    public static readonly HashSet<string> Collections = new(StringComparer.OrdinalIgnoreCase)
    {
        "workstations", "sessions", "accounts", "cards", "pricing", "operators", "journal", "reports", "configuration"
    };

    public static readonly HashSet<string> Commands = new(StringComparer.OrdinalIgnoreCase)
    {
        "session.pause", "session.resume", "session.end", "workstation.lock", "workstation.unlock",
        "workstation.shutdown", "workstation.message", "card.recharge", "card.deactivate",
        "account.disable", "account.enable"
    };
}
