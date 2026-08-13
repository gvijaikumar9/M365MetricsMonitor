using System.Diagnostics;
using System.Net.Http.Headers;
using System.Text;
using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;

// M365 Metrics Monitor - local dashboard on localhost, data from Microsoft Graph.
// Auth is app-only (client credentials): no user sign-in.

var builder = WebApplication.CreateBuilder(args);
var app = builder.Build();

const string url = "http://localhost:5137";
app.Urls.Add(url);
app.UseDefaultFiles();
app.UseStaticFiles();

var tenantId = builder.Configuration["Graph:TenantId"];
var clientId = builder.Configuration["Graph:ClientId"];
var clientSecret = builder.Configuration["Graph:ClientSecret"];

// Free / self-service SKUs (provisioned with huge seat counts) to keep out of the
// "unused licenses" total.
var freeSkus = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
{
    "FLOW_FREE", "POWER_BI_STANDARD", "POWERAPPS_VIRAL", "POWERAPPS_DEV", "RIGHTSMANAGEMENT_ADHOC",
    "TEAMS_EXPLORATORY", "CCIBOTS_PRIVPREV_VIRAL", "STREAM", "WINDOWS_STORE",
    "POWERAPPS_INDIVIDUAL_USER", "FORMS_PRO", "PROJECT_MADEIRA_PREVIEW_IW_SKU",
    "MICROSOFT_BUSINESS_CENTER", "DYN365_ENTERPRISE_P1_IW"
};

ClientSecretCredential? credential = null;
GraphServiceClient? graph = null;
ClientSecretCredential Cred() => credential ??= new ClientSecretCredential(tenantId, clientId, clientSecret);
GraphServiceClient G() => graph ??= new GraphServiceClient(Cred(), new[] { "https://graph.microsoft.com/.default" });
bool Configured() => !NotConfigured(tenantId) && !NotConfigured(clientId) && !NotConfigured(clientSecret);

// Save connection settings from the UI: update runtime, reset the Graph client, persist to appsettings.json.
var appsettingsPath = Path.Combine(app.Environment.ContentRootPath, "appsettings.json");
void SaveConfig(string t, string c, string s)
{
    tenantId = t; clientId = c; clientSecret = s;
    credential = null; graph = null;
    var json = System.Text.Json.JsonSerializer.Serialize(new
    {
        Graph = new { TenantId = t, ClientId = c, ClientSecret = s },
        Logging = new { LogLevel = new Dictionary<string, string> { ["Default"] = "Warning", ["Microsoft.AspNetCore"] = "Warning" } }
    }, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
    File.WriteAllText(appsettingsPath, json);
}

// ---- Multi-tenant profiles (saved to tenants.json next to the app) -----------
var tenantsPath = Path.Combine(app.Environment.ContentRootPath, "tenants.json");
var tenantJsonOpts = new System.Text.Json.JsonSerializerOptions { WriteIndented = true };
var tenants = new List<TenantProfile>();
string? activeTenantId = null;
void SaveTenants() => File.WriteAllText(tenantsPath,
    System.Text.Json.JsonSerializer.Serialize(new TenantStore(activeTenantId, tenants), tenantJsonOpts));
void ApplyActive()
{
    var p = tenants.FirstOrDefault(t => t.Id == activeTenantId);
    if (p is not null) { tenantId = p.TenantId; clientId = p.ClientId; clientSecret = p.ClientSecret; }
    else { tenantId = clientId = clientSecret = null; }
    credential = null; graph = null;
}
void LoadTenants()
{
    if (File.Exists(tenantsPath))
    {
        try
        {
            var store = System.Text.Json.JsonSerializer.Deserialize<TenantStore>(File.ReadAllText(tenantsPath));
            if (store is not null) { tenants = store.Tenants ?? new(); activeTenantId = store.ActiveId; }
        }
        catch { /* corrupt file: fall back to appsettings seed below */ }
    }
    // First run with only appsettings.json filled in: seed a profile from it.
    if (tenants.Count == 0 && Configured())
    {
        var seed = new TenantProfile(Guid.NewGuid().ToString("N"), "My tenant", tenantId!, clientId!, clientSecret!);
        tenants.Add(seed); activeTenantId = seed.Id; SaveTenants();
    }
    if (activeTenantId is null && tenants.Count > 0) activeTenantId = tenants[0].Id;
    if (tenants.Count > 0) ApplyActive();
}
LoadTenants();

// ---- Daily metric snapshots (local history that powers trend deltas) --------
var snapshotsPath = Path.Combine(app.Environment.ContentRootPath, "snapshots.json");
var snaps = new Dictionary<string, List<Snapshot>>();
void LoadSnaps()
{
    if (File.Exists(snapshotsPath))
    {
        try { snaps = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, List<Snapshot>>>(File.ReadAllText(snapshotsPath)) ?? new(); }
        catch { snaps = new(); }
    }
}
void SaveSnaps() => File.WriteAllText(snapshotsPath, System.Text.Json.JsonSerializer.Serialize(snaps, tenantJsonOpts));
LoadSnaps();

// ---- Config + permission validation -----------------------------------------
app.MapGet("/api/config", () => Results.Json(new
{
    configured = Configured(),
    tenantId = NotConfigured(tenantId) ? "" : tenantId,
    clientId = NotConfigured(clientId) ? "" : clientId,
    hasSecret = !NotConfigured(clientSecret)
}));

app.MapPost("/api/config", async (ConfigDto body) =>
{
    // blank secret means "keep the existing one"
    var secret = (body is null || string.IsNullOrWhiteSpace(body.clientSecret)) ? clientSecret : body.clientSecret.Trim();
    if (body is null || string.IsNullOrWhiteSpace(body.tenantId) || string.IsNullOrWhiteSpace(body.clientId) || NotConfigured(secret))
        return Results.Json(new { ok = false, message = "Tenant ID, Client ID and Client Secret are all required." });
    SaveConfig(body.tenantId.Trim(), body.clientId.Trim(), secret);
    try
    {
        await Cred().GetTokenAsync(new TokenRequestContext(new[] { "https://graph.microsoft.com/.default" }));
        return Results.Json(new { ok = true, message = "Connected. Live data will load." });
    }
    catch (Exception ex) { return Results.Json(new { ok = false, message = Describe(ex) }); }
});

// ---- Multi-tenant: list / add-or-update / switch / delete --------------------
app.MapGet("/api/tenants", () => Results.Json(new
{
    activeId = activeTenantId,
    tenants = tenants.Select(t => new
    {
        id = t.Id,
        label = t.Label,
        tenantId = t.TenantId,
        clientId = t.ClientId,
        hasSecret = !NotConfigured(t.ClientSecret),
        active = t.Id == activeTenantId
    })
}));

app.MapPost("/api/tenants", async (TenantDto body) =>
{
    if (body is null || string.IsNullOrWhiteSpace(body.tenantId) || string.IsNullOrWhiteSpace(body.clientId))
        return Results.Json(new { ok = false, message = "Tenant ID and Client ID are required." });
    var existing = tenants.FirstOrDefault(t => t.Id == body.id);
    // blank secret on an existing tenant means "keep the saved one"
    var secret = string.IsNullOrWhiteSpace(body.clientSecret) ? (existing?.ClientSecret ?? "") : body.clientSecret.Trim();
    if (NotConfigured(secret))
        return Results.Json(new { ok = false, message = "Client Secret is required." });
    var label = string.IsNullOrWhiteSpace(body.label) ? body.tenantId.Trim() : body.label.Trim();
    var profile = new TenantProfile(existing?.Id ?? Guid.NewGuid().ToString("N"), label, body.tenantId.Trim(), body.clientId.Trim(), secret);
    if (existing is not null) tenants[tenants.IndexOf(existing)] = profile; else tenants.Add(profile);
    activeTenantId = profile.Id; ApplyActive(); SaveTenants();
    try
    {
        await Cred().GetTokenAsync(new TokenRequestContext(new[] { "https://graph.microsoft.com/.default" }));
        return Results.Json(new { ok = true, message = "Connected. Live data will load.", id = profile.Id });
    }
    catch (Exception ex) { return Results.Json(new { ok = false, message = Describe(ex), id = profile.Id }); }
});

app.MapPost("/api/tenants/{id}/activate", async (string id) =>
{
    var p = tenants.FirstOrDefault(t => t.Id == id);
    if (p is null) return Results.Json(new { ok = false, message = "Tenant not found." });
    activeTenantId = id; ApplyActive(); SaveTenants();
    try
    {
        await Cred().GetTokenAsync(new TokenRequestContext(new[] { "https://graph.microsoft.com/.default" }));
        return Results.Json(new { ok = true, message = "Switched to " + p.Label });
    }
    catch (Exception ex) { return Results.Json(new { ok = false, message = Describe(ex) }); }
});

app.MapDelete("/api/tenants/{id}", (string id) =>
{
    var p = tenants.FirstOrDefault(t => t.Id == id);
    if (p is null) return Results.Json(new { ok = false, message = "Tenant not found." });
    tenants.Remove(p);
    if (activeTenantId == id) { activeTenantId = tenants.FirstOrDefault()?.Id; ApplyActive(); }
    SaveTenants();
    return Results.Json(new { ok = true });
});

// Record today's metrics for the active tenant (upsert one row per UTC day).
app.MapPost("/api/snapshot", (SnapshotDto body) =>
{
    if (activeTenantId is null || body?.metrics is null || body.metrics.Count == 0)
        return Results.Json(new { ok = false });
    var today = DateTimeOffset.UtcNow.ToString("yyyy-MM-dd");
    if (!snaps.TryGetValue(activeTenantId, out var list)) { list = new(); snaps[activeTenantId] = list; }
    var existing = list.FirstOrDefault(s => s.date == today);
    if (existing is not null) list[list.IndexOf(existing)] = new Snapshot(today, body.metrics);
    else list.Add(new Snapshot(today, body.metrics));
    if (list.Count > 400) list.RemoveRange(0, list.Count - 400);   // keep ~1 year
    SaveSnaps();
    return Results.Json(new { ok = true, days = list.Count });
});

// Full daily series for the active tenant (front-end computes the deltas).
app.MapGet("/api/trends", () =>
{
    if (activeTenantId is null || !snaps.TryGetValue(activeTenantId, out var list) || list.Count == 0)
        return Results.Json(new { series = Array.Empty<object>() });
    var series = list.OrderBy(s => s.date).Select(s =>
    {
        var row = new Dictionary<string, object?> { ["date"] = s.date };
        foreach (var kv in s.metrics) row[kv.Key] = kv.Value;
        return row;
    });
    return Results.Json(new { series });
});

app.MapGet("/api/permissions/check", async () =>
{
    if (!Configured()) return Results.Json(new { error = "not_configured" });
    try { await Cred().GetTokenAsync(new TokenRequestContext(new[] { "https://graph.microsoft.com/.default" })); }
    catch (Exception ex) { return Results.Json(new { connected = false, message = Describe(ex) }); }

    async Task<bool> Probe(Func<Task> call) { try { await call(); return true; } catch { return false; } }
    var checks = new List<object>();
    async Task Add(string perm, string[] tiles, Func<Task> call)
    {
        checks.Add(new { permission = perm, granted = await Probe(call), tiles });
    }

    await Add("Organization.Read.All", new[] { "Licenses", "Unused licenses" }, async () => await G().Organization.GetAsync());
    await Add("User.Read.All", new[] { "Users", "Guest users" }, async () => await G().Users.GetAsync(rc => rc.QueryParameters.Top = 1));
    await Add("Reports.Read.All", new[] { "Storage", "Sites", "Workload", "Copilot usage" }, async () => await ReportCsv("getOffice365ServicesUserCounts(period='D7')"));
    await Add("Group.Read.All", new[] { "Governance" }, async () => await G().Groups.GetAsync(rc => rc.QueryParameters.Top = 1));
    await Add("Directory.Read.All", new[] { "App secrets", "Admin roles" }, async () => await G().Applications.GetAsync(rc => rc.QueryParameters.Top = 1));
    await Add("ServiceHealth.Read.All", new[] { "Service health", "Services" }, async () => await G().Admin.ServiceAnnouncement.HealthOverviews.GetAsync());
    await Add("SecurityEvents.Read.All", new[] { "Secure Score" }, async () => await G().Security.SecureScores.GetAsync(rc => rc.QueryParameters.Top = 1));
    return Results.Json(new { connected = true, checks });
});

// Usage reports come back as CSV, fetched with a raw token.
async Task<string> ReportCsv(string func, string version = "v1.0")
{
    var token = (await Cred().GetTokenAsync(new TokenRequestContext(new[] { "https://graph.microsoft.com/.default" }))).Token;
    using var http = new HttpClient();
    http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
    return await http.GetStringAsync($"https://graph.microsoft.com/{version}/reports/{func}");
}

// ---- Licenses (live) --------------------------------------------------------
app.MapGet("/api/licenses", async () =>
{
    if (!Configured())
        return Results.Json(new { error = "not_configured", message = "Set Graph:TenantId, Graph:ClientId and Graph:ClientSecret in appsettings.json, then restart." });
    try
    {
        var resp = await G().SubscribedSkus.GetAsync();
        var names = FriendlyNames();
        var skus = (resp?.Value ?? new())
            .Select(s =>
            {
                var owned = s.PrepaidUnits?.Enabled ?? 0;
                var used = s.ConsumedUnits ?? 0;
                var part = s.SkuPartNumber ?? "";
                return new
                {
                    sku = part,
                    name = names.TryGetValue(part, out var n) ? n : part,
                    used,
                    owned,
                    free = owned - used,
                    isFree = freeSkus.Contains(part) || owned >= 1_000_000
                };
            })
            .OrderByDescending(x => x.owned)
            .ToList();
        return Results.Json(new { skus });
    }
    catch (Exception ex) { return Results.Json(new { error = "graph_error", message = Describe(ex) }); }
});

// ---- Service health (live) --------------------------------------------------
app.MapGet("/api/servicehealth", async () =>
{
    if (!Configured()) return Results.Json(new { error = "not_configured" });
    try
    {
        var resp = await G().Admin.ServiceAnnouncement.HealthOverviews.GetAsync();
        var svcs = resp?.Value ?? new();
        var issues = svcs
            .Where(s => (s.Status?.ToString() ?? "") != "ServiceOperational")
            .Select(s => new { service = s.Service, status = s.Status?.ToString() })
            .ToList();

        // message center: posts updated in the last 30 days
        int messageCenter = 0;
        try
        {
            var msgs = await G().Admin.ServiceAnnouncement.Messages.GetAsync(rc => rc.QueryParameters.Top = 100);
            var cutoff = DateTimeOffset.UtcNow.AddDays(-30);
            messageCenter = (msgs?.Value ?? new())
                .Count(m => (m.LastModifiedDateTime ?? m.StartDateTime ?? DateTimeOffset.MinValue) >= cutoff);
        }
        catch { /* messages unavailable, leave 0 */ }

        var all = svcs
            .Select(s => new { service = s.Service, status = s.Status?.ToString() ?? "" })
            .OrderBy(s => s.status == "ServiceOperational" ? 1 : 0)
            .ThenBy(s => s.service)
            .ToList();
        return Results.Json(new { total = svcs.Count, healthy = issues.Count == 0, issues, messageCenter, all });
    }
    catch (Exception ex) { return Results.Json(new { error = "graph_error", message = Describe(ex) }); }
});

// ---- Organization / tenant name (live) --------------------------------------
app.MapGet("/api/org", async () =>
{
    if (!Configured()) return Results.Json(new { error = "not_configured" });
    try
    {
        var org = (await G().Organization.GetAsync())?.Value?.FirstOrDefault();
        var domain = org?.VerifiedDomains?.FirstOrDefault(d => d.IsDefault == true)?.Name
                     ?? org?.VerifiedDomains?.FirstOrDefault()?.Name;
        return Results.Json(new { name = org?.DisplayName, domain });
    }
    catch (Exception ex) { return Results.Json(new { error = "graph_error", message = Describe(ex) }); }
});

// ---- App secrets & certificates expiring (live) -----------------------------
app.MapGet("/api/secrets", async () =>
{
    if (!Configured()) return Results.Json(new { error = "not_configured" });
    try
    {
        var resp = await G().Applications.GetAsync(rc =>
        {
            rc.QueryParameters.Select = new[] { "displayName", "appId", "passwordCredentials", "keyCredentials" };
            rc.QueryParameters.Top = 999;
        });
        var now = DateTimeOffset.UtcNow;
        var raw = new List<(string app, string appId, string type, int days)>();
        foreach (var a in resp?.Value ?? new())
        {
            void Consider(DateTimeOffset? end, string type)
            {
                if (end is null) return;
                var days = (int)Math.Floor((end.Value - now).TotalDays);
                raw.Add((a.DisplayName ?? "(unnamed app)", a.AppId ?? "", type, days));
            }
            foreach (var p in a.PasswordCredentials ?? new()) Consider(p.EndDateTime, "Secret");
            foreach (var k in a.KeyCredentials ?? new()) Consider(k.EndDateTime, "Certificate");
        }
        var expiring30 = raw.Count(x => x.days <= 30);
        // full list, sorted soonest first (the tile slices; the "View all" page shows everything)
        var items = raw.OrderBy(x => x.days)
            .Select(x => new { app = x.app, appId = x.appId, type = x.type, days = x.days }).ToList();
        return Results.Json(new { expiring30, items });
    }
    catch (Exception ex) { return Results.Json(new { error = "graph_error", message = Describe(ex) }); }
});

// ---- Users by domain + guests (live) ----------------------------------------
app.MapGet("/api/users", async () =>
{
    if (!Configured()) return Results.Json(new { error = "not_configured" });
    try
    {
        var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        int total = 0, guests = 0;
        var list = new List<object>();
        var resp = await G().Users.GetAsync(rc =>
        {
            rc.QueryParameters.Select = new[] { "displayName", "userPrincipalName", "userType", "accountEnabled" };
            rc.QueryParameters.Top = 999;
        });
        int pages = 0;
        while (resp is not null)
        {
            foreach (var u in resp.Value ?? new())
            {
                total++;
                var upn = u.UserPrincipalName ?? "";
                var at = upn.LastIndexOf('@');
                var dom = at >= 0 ? upn[(at + 1)..] : "(none)";
                var isGuest = string.Equals(u.UserType, "Guest", StringComparison.OrdinalIgnoreCase);
                if (isGuest) guests++;
                else counts[dom] = counts.GetValueOrDefault(dom) + 1;
                list.Add(new
                {
                    name = u.DisplayName ?? "",
                    upn,
                    type = isGuest ? "Guest" : "Member",
                    domain = dom,
                    enabled = u.AccountEnabled ?? true
                });
            }
            var next = resp.OdataNextLink;
            if (string.IsNullOrEmpty(next) || ++pages >= 200) break;  // ~200k-user safety ceiling
            resp = await G().Users.WithUrl(next).GetAsync();
        }
        var domains = counts.OrderByDescending(kv => kv.Value).Take(6)
            .Select(kv => new { domain = kv.Key, count = kv.Value }).ToList();
        return Results.Json(new { total, guests, domains, list });
    }
    catch (Exception ex) { return Results.Json(new { error = "graph_error", message = Describe(ex) }); }
});

// ---- Storage: sites by template, top sites, total (reported ~2 days) ---------
app.MapGet("/api/storage", async () =>
{
    if (!Configured()) return Results.Json(new { error = "not_configured" });
    try
    {
        var csv = await ReportCsv("getSharePointSiteUsageDetail(period='D7')");
        var rows = ParseCsv(csv);
        if (rows.Count < 2) return Results.Json(new { error = "no_data" });
        var header = rows[0];
        int Col(string name) => Array.FindIndex(header, h => h.Trim().Equals(name, StringComparison.OrdinalIgnoreCase));
        int cUrl = Col("Site URL"), cStore = Col("Storage Used (Byte)"), cTmpl = Col("Root Web Template"),
            cFiles = Col("File Count"), cLast = Col("Last Activity Date");

        long totalBytes = 0;
        var byTemplate = new Dictionary<string, int>();
        var sites = new List<(string url, string tmpl, long bytes, long files, string last)>();
        foreach (var r in rows.Skip(1))
        {
            if (cStore < 0 || r.Length <= cStore) continue;
            long bytes = long.TryParse(r[cStore], out var b) ? b : 0;
            totalBytes += bytes;
            var tmpl = cTmpl >= 0 && cTmpl < r.Length ? r[cTmpl] : "";
            byTemplate[tmpl] = byTemplate.GetValueOrDefault(tmpl) + 1;
            var siteUrl = cUrl >= 0 && cUrl < r.Length ? r[cUrl] : "";
            long files = cFiles >= 0 && cFiles < r.Length && long.TryParse(r[cFiles], out var f) ? f : 0;
            var last = cLast >= 0 && cLast < r.Length ? r[cLast] : "";
            sites.Add((siteUrl, tmpl, bytes, files, last));
        }
        var top = sites.OrderByDescending(s => s.bytes).Take(5)
            .Select(s => new
            {
                url = SiteName(s.url),
                template = FriendlyTemplate(s.tmpl),
                files = s.files,
                gb = Math.Round(s.bytes / 1073741824.0, 1),
                last = s.last
            }).ToList();
        var all = sites.OrderByDescending(s => s.bytes)
            .Select(s => new
            {
                url = SiteName(s.url),
                template = FriendlyTemplate(s.tmpl),
                files = s.files,
                gb = Math.Round(s.bytes / 1073741824.0, 2),
                last = s.last
            }).ToList();
        var templates = byTemplate.OrderByDescending(kv => kv.Value).Take(6)
            .Select(kv => new { template = FriendlyTemplate(kv.Key), count = kv.Value }).ToList();
        double totalGB = Math.Round(totalBytes / 1073741824.0, 1);
        return Results.Json(new { totalSites = sites.Count, totalGB, totalTB = Math.Round(totalGB / 1024.0, 2), templates, top, all });
    }
    catch (Exception ex) { return Results.Json(new { error = "graph_error", message = Describe(ex) }); }
});

// ---- Copilot adoption (reported ~2 days) ------------------------------------
app.MapGet("/api/copilot", async () =>
{
    if (!Configured()) return Results.Json(new { error = "not_configured" });
    try
    {
        var skus = await G().SubscribedSkus.GetAsync();
        var cop = (skus?.Value ?? new()).FirstOrDefault(s => (s.SkuPartNumber ?? "").Contains("Copilot", StringComparison.OrdinalIgnoreCase));
        int owned = cop?.PrepaidUnits?.Enabled ?? 0;
        int assigned = cop?.ConsumedUnits ?? 0;
        int active = 0;
        try
        {
            var csv = await ReportCsv("getMicrosoft365CopilotUsageUserDetail(period='D7')", "beta");
            var rows = ParseCsv(csv);
            if (rows.Count > 1)
            {
                var header = rows[0];
                int cLast = Array.FindIndex(header, h => h.Trim().Equals("Last activity date", StringComparison.OrdinalIgnoreCase));
                foreach (var r in rows.Skip(1))
                    if (cLast >= 0 && cLast < r.Length && !string.IsNullOrWhiteSpace(r[cLast])) active++;
            }
        }
        catch { /* Copilot report not available (e.g. no Copilot in tenant) */ }
        int pct = assigned > 0 ? (int)Math.Round(active * 100.0 / assigned) : 0;
        return Results.Json(new { hasCopilot = owned > 0, owned, assigned, active, pct });
    }
    catch (Exception ex) { return Results.Json(new { error = "graph_error", message = Describe(ex) }); }
});

// ---- Workload activity: active users per service (reported ~2 days) ----------
app.MapGet("/api/workload", async () =>
{
    if (!Configured()) return Results.Json(new { error = "not_configured" });
    try
    {
        var csv = await ReportCsv("getOffice365ServicesUserCounts(period='D30')");
        var rows = ParseCsv(csv);
        if (rows.Count < 2) return Results.Json(new { error = "no_data" });
        var header = rows[0];
        var data = rows[1];
        int Get(string name)
        {
            var i = Array.FindIndex(header, h => h.Trim().Equals(name, StringComparison.OrdinalIgnoreCase));
            return i >= 0 && i < data.Length && int.TryParse(data[i], out var v) ? v : 0;
        }
        var items = new[]
        {
            new { service = "Teams",      active = Get("Teams Active") },
            new { service = "SharePoint", active = Get("SharePoint Active") },
            new { service = "Exchange",   active = Get("Exchange Active") },
            new { service = "OneDrive",   active = Get("OneDrive Active") },
        };
        return Results.Json(new { items });
    }
    catch (Exception ex) { return Results.Json(new { error = "graph_error", message = Describe(ex) }); }
});

// ---- Governance: ownerless / empty groups, admin holders (live) -------------
app.MapGet("/api/governance", async () =>
{
    if (!Configured()) return Results.Json(new { error = "not_configured" });
    try
    {
        var groups = await G().Groups.GetAsync(rc =>
        {
            rc.QueryParameters.Select = new[] { "id", "displayName", "groupTypes" };
            rc.QueryParameters.Top = 999;
        });
        var unified = (groups?.Value ?? new())
            .Where(g => g.GroupTypes != null && g.GroupTypes.Contains("Unified"))
            .Take(300).ToList();

        var checks = await Task.WhenAll(unified.Select(async g =>
        {
            var owners = await G().Groups[g.Id].Owners.GetAsync(rc => rc.QueryParameters.Top = 1);
            var members = await G().Groups[g.Id].Members.GetAsync(rc => rc.QueryParameters.Top = 1);
            return (ownerless: (owners?.Value?.Count ?? 0) == 0, empty: (members?.Value?.Count ?? 0) == 0);
        }));
        int ownerless = checks.Count(c => c.ownerless);
        int empty = checks.Count(c => c.empty);

        var roles = await G().DirectoryRoles.GetAsync();
        var admins = new HashSet<string>();
        foreach (var role in roles?.Value ?? new())
        {
            var mem = await G().DirectoryRoles[role.Id].Members.GetAsync();
            foreach (var m in mem?.Value ?? new()) if (m.Id != null) admins.Add(m.Id);
        }

        return Results.Json(new { groups = unified.Count, ownerless, empty, admins = admins.Count });
    }
    catch (Exception ex) { return Results.Json(new { error = "graph_error", message = Describe(ex) }); }
});

// ---- Secure Score (reported ~daily) -----------------------------------------
app.MapGet("/api/securescore", async () =>
{
    if (!Configured()) return Results.Json(new { error = "not_configured" });
    try
    {
        var resp = await G().Security.SecureScores.GetAsync(rc => rc.QueryParameters.Top = 5);
        var s = (resp?.Value ?? new()).OrderByDescending(x => x.CreatedDateTime).FirstOrDefault();
        if (s is null) return Results.Json(new { error = "no_data" });
        double cur = s.CurrentScore ?? 0, max = s.MaxScore ?? 0;
        int pct = max > 0 ? (int)Math.Round(cur * 100 / max) : 0;
        var peer = s.AverageComparativeScores?.FirstOrDefault(a => a.Basis == "AllTenants")?.AverageScore;
        int? peerPct = (max > 0 && peer.HasValue) ? (int)Math.Round(peer.Value * 100 / max) : (int?)null;
        return Results.Json(new { pct, current = (int)Math.Round(cur), max = (int)Math.Round(max), peerPct });
    }
    catch (Exception ex) { return Results.Json(new { error = "graph_error", message = Describe(ex) }); }
});

// ---- Startup ----------------------------------------------------------------
var openBrowser = !args.Contains("--no-open");
app.Lifetime.ApplicationStarted.Register(() =>
{
    if (openBrowser)
    {
        try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); } catch { }
    }
});

Console.WriteLine($"M365 Metrics Monitor running at {url}");
app.Run();

// ---- helpers ----------------------------------------------------------------
static bool NotConfigured(string? v) => string.IsNullOrWhiteSpace(v) || v.StartsWith("YOUR_");

static string Describe(Exception ex) => ex switch
{
    Microsoft.Graph.Models.ODataErrors.ODataError oe => $"{oe.Error?.Code}: {oe.Error?.Message}".Trim(':', ' '),
    _ => string.IsNullOrEmpty(ex.Message) ? ex.GetType().Name : ex.Message
};

static string SiteName(string url)
{
    if (string.IsNullOrEmpty(url)) return "(root)";
    var i = url.TrimEnd('/').LastIndexOf('/');
    return i >= 0 ? url.TrimEnd('/')[(i + 1)..] : url;
}

static string FriendlyTemplate(string code) => code switch
{
    "GROUP#0" => "Team site (Group)",
    "TEAMCHANNEL#0" or "TEAMCHANNEL#1" => "Channel site",
    "SITEPAGEPUBLISHING#0" => "Communication",
    "STS#3" => "Team site",
    "STS#0" => "Team site (classic)",
    "SPSPERS#10" or "SPSPERS#0" => "OneDrive",
    "BLANKINTERNETCONTAINER#0" => "Publishing",
    "" => "(unknown)",
    _ => code
};

// Minimal CSV parser (handles quoted fields). Report CSVs may carry a BOM.
static List<string[]> ParseCsv(string text)
{
    if (text.Length > 0 && text[0] == '﻿') text = text[1..];
    var rows = new List<string[]>();
    using var reader = new StringReader(text);
    string? line;
    while ((line = reader.ReadLine()) != null)
    {
        if (line.Length == 0) continue;
        var fields = new List<string>();
        var sb = new StringBuilder();
        bool inQ = false;
        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];
            if (inQ)
            {
                if (c == '"')
                {
                    if (i + 1 < line.Length && line[i + 1] == '"') { sb.Append('"'); i++; }
                    else inQ = false;
                }
                else sb.Append(c);
            }
            else if (c == '"') inQ = true;
            else if (c == ',') { fields.Add(sb.ToString()); sb.Clear(); }
            else sb.Append(c);
        }
        fields.Add(sb.ToString());
        rows.Add(fields.ToArray());
    }
    return rows;
}

static Dictionary<string, string> FriendlyNames() => new()
{
    ["SPE_E3"] = "Microsoft 365 E3",
    ["SPE_E5"] = "Microsoft 365 E5",
    ["SPE_F1"] = "Microsoft 365 F3",
    ["ENTERPRISEPACK"] = "Office 365 E3",
    ["ENTERPRISEPREMIUM"] = "Office 365 E5",
    ["Microsoft_365_Copilot"] = "Microsoft 365 Copilot",
    ["EXCHANGESTANDARD"] = "Exchange Online Plan 1",
    ["EXCHANGEENTERPRISE"] = "Exchange Online Plan 2",
    ["POWER_BI_STANDARD"] = "Power BI (free)",
    ["FLOW_FREE"] = "Power Automate Free",
    ["POWERAPPS_DEV"] = "Power Apps for Developer",
    ["TEAMS_EXPLORATORY"] = "Teams Exploratory",
    ["VISIOCLIENT"] = "Visio Plan 2",
    ["PROJECTPROFESSIONAL"] = "Project Plan 3",
    ["DEVELOPERPACK_E5"] = "Microsoft 365 E5 Developer",
    ["DEVELOPERPACK"] = "Office 365 E3 Developer",
};

record ConfigDto(string tenantId, string clientId, string clientSecret);
record TenantDto(string? id, string label, string tenantId, string clientId, string clientSecret);
record TenantProfile(string Id, string Label, string TenantId, string ClientId, string ClientSecret);
record TenantStore(string? ActiveId, List<TenantProfile> Tenants);
record SnapshotDto(Dictionary<string, double?> metrics);
record Snapshot(string date, Dictionary<string, double?> metrics);
