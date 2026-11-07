using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

var builder = WebApplication.CreateBuilder(new WebApplicationOptions
{
    Args = args,
    ContentRootPath = AppContext.BaseDirectory
});
var externalConfigPath = Environment.GetEnvironmentVariable("CYBERA_BRIDGE_CONFIG");
if (!string.IsNullOrWhiteSpace(externalConfigPath))
{
    builder.Configuration.AddJsonFile(externalConfigPath, optional: false, reloadOnChange: true);
}
var settings = builder.Configuration.GetSection("Bridge").Get<BridgeSettings>() ?? new BridgeSettings();
settings.Validate();

const string ApiKeyEnvironmentVariable = "CYBERA_BRIDGE_API_KEY";
var apiKey = Environment.GetEnvironmentVariable(ApiKeyEnvironmentVariable);
if (string.IsNullOrWhiteSpace(apiKey))
{
    throw new InvalidOperationException("CYBERA_BRIDGE_API_KEY is missing from the bridge process environment.");
}

builder.Services.AddSingleton(settings);
builder.Services.AddSingleton(new ApiKeyValidator(apiKey));
builder.Services.AddSingleton<LegacyStatusReader>();
builder.Services.AddSingleton<LegacyCommandRelay>();
builder.Services.AddSingleton<AuditLog>();

var app = builder.Build();
app.Use(async (context, next) =>
{
    var remoteIp = context.Connection.RemoteIpAddress;
    if (remoteIp is null || !settings.IsAllowed(remoteIp))
    {
        context.Response.StatusCode = StatusCodes.Status403Forbidden;
        await context.Response.WriteAsJsonAsync(new { error = "network_not_allowed" });
        return;
    }

    context.Response.Headers.CacheControl = "no-store";
    await next();
});

app.UseDefaultFiles();
app.UseStaticFiles();

app.MapGet("/health", () => Results.Ok(new
{
    status = "ok",
    service = "cybera-modern-bridge",
    utc = DateTimeOffset.UtcNow
}));

var api = app.MapGroup("/api/v1");
api.AddEndpointFilter(async (context, next) =>
{
    var validator = context.HttpContext.RequestServices.GetRequiredService<ApiKeyValidator>();
    var presented = context.HttpContext.Request.Headers["X-Api-Key"].ToString();
    if (!validator.IsValid(presented))
    {
        return Results.Json(new { error = "unauthorized" }, statusCode: StatusCodes.Status401Unauthorized);
    }
    return await next(context);
});

api.MapGet("/status", (LegacyStatusReader reader) => Results.Ok(reader.Read()));
api.MapGet("/features", () => Results.Ok(FeatureCatalog.All));
api.MapGet("/data/{collection}", (string collection, LegacyStatusReader reader) =>
    FeatureCatalog.Collections.Contains(collection)
        ? Results.Ok(reader.GetCollection(collection))
        : Results.NotFound(new { error = "unknown_collection" }));

api.MapPost("/sessions/{workstationId}/commands", async (
    string workstationId,
    SessionCommandRequest request,
    HttpContext context,
    LegacyCommandRelay relay,
    CancellationToken cancellationToken) =>
{
    if (!WorkstationId.IsValid(workstationId))
    {
        return Results.BadRequest(new { error = "invalid_workstation_id" });
    }

    if (!SessionAction.TryParse(request.Action, out var action))
    {
        return Results.BadRequest(new { error = "unsupported_action", supportedActions = SessionAction.Names });
    }

    var idempotencyKey = context.Request.Headers["Idempotency-Key"].ToString();
    if (idempotencyKey.Length is < 16 or > 128)
    {
        return Results.BadRequest(new { error = "idempotency_key_required" });
    }

    var command = new RelayCommand(Guid.NewGuid(), workstationId, action.Name, idempotencyKey, DateTimeOffset.UtcNow);
    var result = await relay.SendAsync(command, cancellationToken);
    return result.Success
        ? Results.Accepted($"/api/v1/commands/{command.Id}", new { command.Id, state = "accepted" })
        : Results.Json(new { error = result.Error }, statusCode: result.StatusCode);
});

api.MapPost("/legacy/commands", async (
    LegacyCommandRequest request,
    LegacyCommandRelay relay,
    AuditLog audit,
    CancellationToken cancellationToken) =>
{
    if (!FeatureCatalog.Commands.Contains(request.Action ?? "") || !WorkstationId.IsValid(request.TargetId ?? ""))
        return Results.BadRequest(new { error = "unsupported_or_invalid_command" });
    if (string.IsNullOrWhiteSpace(request.IdempotencyKey) || request.IdempotencyKey.Length is < 16 or > 128)
        return Results.BadRequest(new { error = "idempotency_key_required" });
    if (!audit.IsConfigured)
        return Results.Json(new { error = "audit_log_not_configured" }, statusCode: StatusCodes.Status503ServiceUnavailable);

    var command = new RelayCommand(Guid.NewGuid(), request.TargetId!, request.Action!, request.IdempotencyKey!, DateTimeOffset.UtcNow);
    var result = await relay.SendAsync(command, cancellationToken);
    await audit.WriteAsync(command, result, cancellationToken);
    return result.Success ? Results.Accepted($"/api/v1/commands/{command.Id}", new { command.Id, state = "accepted" })
        : Results.Json(new { error = result.Error }, statusCode: result.StatusCode);
});

app.MapFallbackToFile("index.html");

app.Run();

public sealed class BridgeSettings
{
    public string[] AllowedNetworks { get; init; } = ["127.0.0.1/32", "::1/128"];
    public string LegacyProcessName { get; init; } = "cybserv";
    public string LegacyStatusFile { get; init; } = "";
    public string AuditLogPath { get; init; } = "";
    public RelaySettings CommandRelay { get; init; } = new();

    public void Validate()
    {
        if (AllowedNetworks.Length == 0 || AllowedNetworks.Any(value => !Cidr.TryParse(value, out _)))
            throw new InvalidOperationException("Bridge:AllowedNetworks must contain valid CIDR ranges.");
        if (CommandRelay.Enabled)
        {
            if (!IPAddress.TryParse(CommandRelay.Host, out var address) || !IPAddress.IsLoopback(address))
                throw new InvalidOperationException("The legacy command relay must bind to a loopback IP address.");
            if (CommandRelay.Port is < 1 or > 65535)
                throw new InvalidOperationException("The legacy command relay port must be between 1 and 65535.");
        }
    }

    public bool IsAllowed(IPAddress address) => AllowedNetworks
        .Select(value => Cidr.TryParse(value, out var range) ? range : null)
        .Any(range => range is not null && range.Contains(address));
}

public sealed class RelaySettings
{
    public bool Enabled { get; init; }
    public string Host { get; init; } = "127.0.0.1";
    public int Port { get; init; } = 49190;
}

public sealed class Cidr(IPAddress network, int prefixLength)
{
    public bool Contains(IPAddress address)
    {
        if (address.IsIPv4MappedToIPv6) address = address.MapToIPv4();
        if (network.IsIPv4MappedToIPv6) network = network.MapToIPv4();
        var candidate = address.GetAddressBytes();
        var baseAddress = network.GetAddressBytes();
        if (candidate.Length != baseAddress.Length) return false;
        var wholeBytes = prefixLength / 8;
        var remainingBits = prefixLength % 8;
        for (var index = 0; index < wholeBytes; index++) if (candidate[index] != baseAddress[index]) return false;
        return remainingBits == 0 || (candidate[wholeBytes] & (byte)(0xFF << (8 - remainingBits))) == (baseAddress[wholeBytes] & (byte)(0xFF << (8 - remainingBits)));
    }

    public static bool TryParse(string value, out Cidr? range)
    {
        range = null;
        var fields = value.Split('/', 2);
        if (fields.Length != 2 || !IPAddress.TryParse(fields[0], out var address) || !int.TryParse(fields[1], out var prefix)) return false;
        if (address.IsIPv4MappedToIPv6) address = address.MapToIPv4();
        var maximum = address.AddressFamily == AddressFamily.InterNetwork ? 32 : 128;
        if (prefix < 0 || prefix > maximum) return false;
        range = new Cidr(address, prefix);
        return true;
    }
}

public sealed class ApiKeyValidator(string expected)
{
    public bool IsValid(string presented) => !string.IsNullOrEmpty(presented) &&
        CryptographicOperations.FixedTimeEquals(Encoding.UTF8.GetBytes(expected), Encoding.UTF8.GetBytes(presented));
}

public sealed class LegacyStatusReader(BridgeSettings settings)
{
    private JsonElement? Snapshot()
    {
        if (string.IsNullOrWhiteSpace(settings.LegacyStatusFile) || !File.Exists(settings.LegacyStatusFile)) return null;
        using var document = JsonDocument.Parse(File.ReadAllText(settings.LegacyStatusFile));
        return document.RootElement.Clone();
    }
    public object Read()
    {
        var snapshot = Snapshot();
        return new
        {
            integration = new
            {
                mode = "local-process-and-optional-snapshot",
                directLegacyTcp = false,
                commandRelayEnabled = settings.CommandRelay.Enabled
            },
            legacyServer = new
            {
                processName = settings.LegacyProcessName,
                processRunning = Process.GetProcessesByName(settings.LegacyProcessName).Length > 0,
                snapshot
            },
            utc = DateTimeOffset.UtcNow
        };
    }
    public object GetCollection(string collection)
    {
        var snapshot = Snapshot();
        if (snapshot is null || snapshot.Value.ValueKind != JsonValueKind.Object || !snapshot.Value.TryGetProperty(collection, out var data))
            return new { available = false, collection, reason = "No trusted adapter snapshot supplies this collection.", data = (object?)null };
        return new { available = true, collection, reason = (string?)null, data = data.Clone() };
    }
}

public sealed class AuditLog(BridgeSettings settings)
{
    public bool IsConfigured => !string.IsNullOrWhiteSpace(settings.AuditLogPath);
    public async Task WriteAsync(RelayCommand command, RelayResult result, CancellationToken cancellationToken)
    {
        if (!IsConfigured) return;
        var path = Environment.ExpandEnvironmentVariables(settings.AuditLogPath);
        var directory = Path.GetDirectoryName(path);
        if (string.IsNullOrWhiteSpace(directory)) throw new InvalidOperationException("Bridge:AuditLogPath must include a directory.");
        Directory.CreateDirectory(directory);
        var record = System.Text.Json.JsonSerializer.Serialize(new { utc = DateTimeOffset.UtcNow, command.Id, command.WorkstationId, command.Action, result.Success, result.StatusCode, result.Error });
        await File.AppendAllTextAsync(path, record + Environment.NewLine, cancellationToken);
    }
}

public sealed class LegacyCommandRelay(BridgeSettings settings, ILogger<LegacyCommandRelay> logger)
{
    public async Task<RelayResult> SendAsync(RelayCommand command, CancellationToken cancellationToken)
    {
        if (!settings.CommandRelay.Enabled)
            return RelayResult.Failure(StatusCodes.Status503ServiceUnavailable, "legacy_command_relay_disabled");

        try
        {
            using var client = new TcpClient();
            using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeout.CancelAfter(TimeSpan.FromSeconds(3));
            await client.ConnectAsync(IPAddress.Parse(settings.CommandRelay.Host), settings.CommandRelay.Port, timeout.Token);
            await using var stream = client.GetStream();
            var payload = JsonSerializer.Serialize(command) + "\n";
            await stream.WriteAsync(Encoding.UTF8.GetBytes(payload), timeout.Token);
            await stream.FlushAsync(timeout.Token);
            logger.LogInformation("Relayed session command {CommandId} for workstation {WorkstationId}", command.Id, command.WorkstationId);
            return RelayResult.Successful;
        }
        catch (OperationCanceledException)
        {
            return RelayResult.Failure(StatusCodes.Status504GatewayTimeout, "legacy_command_relay_timeout");
        }
        catch (SocketException)
        {
            return RelayResult.Failure(StatusCodes.Status502BadGateway, "legacy_command_relay_unavailable");
        }
    }
}

public sealed record SessionCommandRequest(string? Action);
public sealed record LegacyCommandRequest(string? Action, string? TargetId, string? IdempotencyKey);
public sealed record RelayCommand(Guid Id, string WorkstationId, string Action, string IdempotencyKey, DateTimeOffset RequestedAtUtc);
public sealed record RelayResult(bool Success, int StatusCode, string? Error)
{
    public static RelayResult Successful { get; } = new(true, StatusCodes.Status202Accepted, null);
    public static RelayResult Failure(int statusCode, string error) => new(false, statusCode, error);
}

public sealed class SessionAction(string name)
{
    private static readonly HashSet<string> Allowed = new(StringComparer.OrdinalIgnoreCase) { "end", "pause", "resume" };
    public static string[] Names => Allowed.Order().ToArray();
    public string Name { get; } = name;
    public static bool TryParse(string? value, out SessionAction action)
    {
        action = new SessionAction(value?.ToLowerInvariant() ?? "");
        return value is not null && Allowed.Contains(value);
    }
}

public static class WorkstationId
{
    public static bool IsValid(string value) => value.Length is >= 1 and <= 64 && value.All(character => char.IsLetterOrDigit(character) || character is '-' or '_' or '.');
}
