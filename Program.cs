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

// This server is localhost-only and unauthenticated. Guard against DNS-rebinding (reject any
// Host header that isn't loopback) and CSRF (reject cross-origin calls to /api, which could
// otherwise silently add/delete tenants or read tenant data from a page the user is visiting).
static bool IsLoopback(string? h) => h is "localhost" or "127.0.0.1" or "::1";
app.Use(async (ctx, next) =>
{
    if (!IsLoopback(ctx.Request.Host.Host)) { ctx.Response.StatusCode = 403; return; }
    var origin = ctx.Request.Headers.Origin.ToString();
    if (!string.IsNullOrEmpty(origin) && ctx.Request.Path.StartsWithSegments("/api")
        && (!Uri.TryCreate(origin, UriKind.Absolute, out var o) || !IsLoopback(o.Host)))
    { ctx.Response.StatusCode = 403; return; }
    await next();
});

app.UseDefaultFiles();
app.UseStaticFiles();

var tenantId = builder.Configuration["Graph:TenantId"];
var clientId = builder.Configuration["Graph:ClientId"];
var clientSecret = builder.Configuration["Graph:ClientSecret"];
var appVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "0.2.0";
var sharedHttp = new HttpClient();   // one client reused for raw HTTP (usage reports, GitHub) to avoid socket exhaustion

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

// Return a saved tenant's secret for editing. Only reachable from the local UI (the Host/Origin
// guard blocks cross-origin/remote callers); the secret already lives in plaintext in tenants.json.
app.MapGet("/api/tenants/{id}/secret", (string id) =>
{
    var p = tenants.FirstOrDefault(t => t.Id == id);
    return p is null ? Results.Json(new { ok = false }) : Results.Json(new { ok = true, clientSecret = p.ClientSecret });
});

// Record today's metrics for the active tenant (upsert one row per UTC day).
app.MapPost("/api/snapshot", (SnapshotDto body) =>
{
    if (activeTenantId is null || body?.metrics is null || body.metrics.Count == 0)
        return Results.Json(new { ok = false });
    var today = DateTimeOffset.UtcNow.ToString("yyyy-MM-dd");
    if (!snaps.TryGetValue(activeTenantId, out var list)) { list = new(); snaps[activeTenantId] = list; }
    var existing = list.FirstOrDefault(s => s.date == today);
    if (existing is not null)
    {
        // Merge per-metric rather than replacing the whole row. A partial refresh (some
        // tiles missing a permission, or a slow Graph call) then still contributes the
        // metrics that landed without blanking ones captured earlier the same day.
        var merged = new Dictionary<string, double?>(existing.metrics);
        foreach (var kv in body.metrics) merged[kv.Key] = kv.Value;
        list[list.IndexOf(existing)] = new Snapshot(today, merged);
    }
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

// GET /api/update-check -> compares this build to the latest GitHub release so the UI can
// flag when a newer version is available. Runs server-side to avoid browser CORS and to send
// the User-Agent header the GitHub API requires. Bounded so a slow GitHub can't hang the app.
app.MapGet("/api/update-check", async () =>
{
    const string repo = "gvijaikumar9/M365MetricsMonitor";
    const string repoUrl = "https://github.com/gvijaikumar9/M365MetricsMonitor";
    try
    {
        using var msg = new HttpRequestMessage(HttpMethod.Get, $"https://api.github.com/repos/{repo}/releases/latest");
        msg.Headers.UserAgent.ParseAdd($"M365MetricsMonitor/{appVersion}");
        msg.Headers.Accept.ParseAdd("application/vnd.github+json");
        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(8));
        var resp = await sharedHttp.SendAsync(msg, cts.Token);
        if (!resp.IsSuccessStatusCode)   // no releases yet (404) or rate-limited -> "no update"
            return Results.Json(new { current = appVersion, latest = (string?)null, updateAvailable = false, url = repoUrl });
        var doc = System.Text.Json.JsonSerializer.Deserialize<System.Text.Json.JsonElement>(await resp.Content.ReadAsStringAsync());
        var tag = doc.TryGetProperty("tag_name", out var t) ? t.GetString() : null;
        var html = doc.TryGetProperty("html_url", out var h) ? h.GetString() : repoUrl;
        var newer = tag is not null && CompareVersions(tag, appVersion) > 0;
        return Results.Json(new { current = appVersion, latest = tag, updateAvailable = newer, url = html ?? repoUrl });
    }
    catch (Exception ex)
    {
        return Results.Json(new { current = appVersion, latest = (string?)null, updateAvailable = false, url = repoUrl, error = ex.Message });
    }
});

// Usage reports come back as CSV, fetched with a raw token.
async Task<string> ReportCsv(string func, string version = "v1.0")
{
    var token = (await Cred().GetTokenAsync(new TokenRequestContext(new[] { "https://graph.microsoft.com/.default" }))).Token;
    var uri = $"https://graph.microsoft.com/{version}/reports/{func}";
    for (int attempt = 0; ; attempt++)
    {
        using var req = new HttpRequestMessage(HttpMethod.Get, uri);   // per-request auth (shared client is thread-safe this way)
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        var resp = await sharedHttp.SendAsync(req);
        if (resp.IsSuccessStatusCode) return await resp.Content.ReadAsStringAsync();
        if (attempt < 2 && ((int)resp.StatusCode == 429 || (int)resp.StatusCode == 503))   // back off on throttling
        {
            await Task.Delay(resp.Headers.RetryAfter?.Delta ?? TimeSpan.FromSeconds(2));
            continue;
        }
        resp.EnsureSuccessStatusCode();   // surface other failures to the caller's catch
        return "";
    }
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
        // Statuses that are NOT an active problem (operational, or an incident that's already resolved).
        var healthyStatuses = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        { "ServiceOperational", "ServiceRestored", "FalsePositive", "PostIncidentReviewPublished" };
        bool IsHealthy(string? st) => healthyStatuses.Contains(st ?? "ServiceOperational");
        var issues = svcs
            .Where(s => !IsHealthy(s.Status?.ToString()))
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
            .OrderBy(s => IsHealthy(s.status) ? 1 : 0)
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
        var now = DateTimeOffset.UtcNow;
        var raw = new List<(string app, string appId, string type, int days)>();
        var resp = await G().Applications.GetAsync(rc =>
        {
            rc.QueryParameters.Select = new[] { "displayName", "appId", "passwordCredentials", "keyCredentials" };
            rc.QueryParameters.Top = 999;
        });
        int pages = 0;
        while (resp is not null)   // follow nextLink so tenants with >999 app registrations aren't missed
        {
            foreach (var a in resp.Value ?? new())
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
            var next = resp.OdataNextLink;
            if (string.IsNullOrEmpty(next) || ++pages >= 200) break;
            resp = await G().Applications.WithUrl(next).GetAsync();
        }
        // Keep "expiring soon" (0-30 days out) distinct from "already expired" (past due) so
        // a credential that lapsed years ago doesn't inflate the "expiring ≤ 30d" urgency count.
        var expiring30 = raw.Count(x => x.days >= 0 && x.days <= 30);
        var expired = raw.Count(x => x.days < 0);
        // full list, sorted soonest first (the tile slices; the "View all" page shows everything)
        var items = raw.OrderBy(x => x.days)
            .Select(x => new { app = x.app, appId = x.appId, type = x.type, days = x.days }).ToList();
        return Results.Json(new { expiring30, expired, items });
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
            cFiles = Col("File Count"), cLast = Col("Last Activity Date"), cOwner = Col("Owner Display Name");

        long totalBytes = 0;
        var byTemplate = new Dictionary<string, int>();
        var byTemplateBytes = new Dictionary<string, long>();
        var sites = new List<(string url, string owner, string tmpl, long bytes, long files, string last)>();
        foreach (var r in rows.Skip(1))
        {
            if (cStore < 0 || r.Length <= cStore) continue;
            long bytes = long.TryParse(r[cStore], out var b) ? b : 0;
            totalBytes += bytes;
            var tmpl = cTmpl >= 0 && cTmpl < r.Length ? r[cTmpl] : "";
            byTemplate[tmpl] = byTemplate.GetValueOrDefault(tmpl) + 1;
            byTemplateBytes[tmpl] = byTemplateBytes.GetValueOrDefault(tmpl) + bytes;
            var siteUrl = cUrl >= 0 && cUrl < r.Length ? r[cUrl] : "";
            var owner = cOwner >= 0 && cOwner < r.Length ? r[cOwner] : "";
            long files = cFiles >= 0 && cFiles < r.Length && long.TryParse(r[cFiles], out var f) ? f : 0;
            var last = cLast >= 0 && cLast < r.Length ? r[cLast] : "";
            sites.Add((siteUrl, owner, tmpl, bytes, files, last));
        }
        var top = sites.OrderByDescending(s => s.bytes).Take(5)
            .Select(s => new
            {
                url = SiteName(s.url),
                webUrl = s.url,
                owner = s.owner,
                template = FriendlyTemplate(s.tmpl),
                files = s.files,
                gb = Math.Round(s.bytes / 1073741824.0, 1),
                last = s.last
            }).ToList();
        var all = sites.OrderByDescending(s => s.bytes)
            .Select(s => new
            {
                url = SiteName(s.url),
                webUrl = s.url,
                owner = s.owner,
                template = FriendlyTemplate(s.tmpl),
                files = s.files,
                gb = Math.Round(s.bytes / 1073741824.0, 2),
                last = s.last
            }).ToList();
        var templates = byTemplate.OrderByDescending(kv => kv.Value).Take(6)
            .Select(kv => new { template = FriendlyTemplate(kv.Key), count = kv.Value }).ToList();
        var templateStorage = byTemplateBytes.OrderByDescending(kv => kv.Value)
            .Select(kv => new { template = FriendlyTemplate(kv.Key), gb = Math.Round(kv.Value / 1073741824.0, 2) }).ToList();
        double totalGB = Math.Round(totalBytes / 1073741824.0, 1);
        // Report anonymises the Site URL when the tenant's "Display concealed names in reports" setting is on.
        bool concealed = sites.Count > 0 && sites.All(s => string.IsNullOrEmpty(s.url));
        return Results.Json(new { totalSites = sites.Count, totalGB, totalTB = Math.Round(totalGB / 1024.0, 2), templates, templateStorage, top, all, concealed });
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
        int pct = assigned > 0 ? Math.Min(100, (int)Math.Round(active * 100.0 / assigned)) : 0;
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
        int Col(string name) => Array.FindIndex(header, h => h.Trim().Equals(name, StringComparison.OrdinalIgnoreCase));
        int Cell(string[] row, int i) => i >= 0 && i < row.Length && int.TryParse(row[i], out var v) ? v : 0;
        int teamsC = Col("Teams Active"), spC = Col("SharePoint Active"), exC = Col("Exchange Active"), odC = Col("OneDrive Active");
        int dateC = Col("Report Date"); if (dateC < 0) dateC = Col("Report Refresh Date");
        // This report is a per-day time series; use the most recent day that actually has
        // counts (the latest day or two can be blank because usage data lags ~1-2 days),
        // rather than trusting the CSV row order.
        var data = rows[1];
        DateTime best = DateTime.MinValue; bool picked = false;
        for (int r = 1; r < rows.Count; r++)
        {
            var row = rows[r];
            if (Cell(row, teamsC) + Cell(row, spC) + Cell(row, exC) + Cell(row, odC) == 0) continue;
            if (dateC >= 0 && dateC < row.Length && DateTime.TryParse(row[dateC], out var dt))
            { if (dt >= best) { best = dt; data = row; picked = true; } }
            else if (!picked) { data = row; picked = true; }
        }
        int Get(string name) { var i = Col(name); return i >= 0 && i < data.Length && int.TryParse(data[i], out var v) ? v : 0; }
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
        // Page every unified group for an accurate total.
        var unifiedAll = new List<Microsoft.Graph.Models.Group>();
        var groups = await G().Groups.GetAsync(rc =>
        {
            rc.QueryParameters.Select = new[] { "id", "displayName", "groupTypes" };
            rc.QueryParameters.Top = 999;
        });
        int gp = 0;
        while (groups is not null)
        {
            foreach (var g in groups.Value ?? new())
                if (g.GroupTypes != null && g.GroupTypes.Contains("Unified")) unifiedAll.Add(g);
            var gnext = groups.OdataNextLink;
            if (string.IsNullOrEmpty(gnext) || ++gp >= 200) break;
            groups = await G().Groups.WithUrl(gnext).GetAsync();
        }

        // Owner/member checks are 2 calls per group; bound concurrency so a large tenant can't
        // burst Graph into 429s, and cap the analysed set so the tile stays responsive.
        const int GovSampleCap = 500;
        var toCheck = unifiedAll.Take(GovSampleCap).ToList();
        using var gate = new SemaphoreSlim(8);
        var checks = await Task.WhenAll(toCheck.Select(async g =>
        {
            await gate.WaitAsync();
            try
            {
                var owners = await G().Groups[g.Id].Owners.GetAsync(rc => rc.QueryParameters.Top = 1);
                var members = await G().Groups[g.Id].Members.GetAsync(rc => rc.QueryParameters.Top = 1);
                return (ownerless: (owners?.Value?.Count ?? 0) == 0, empty: (members?.Value?.Count ?? 0) == 0);
            }
            finally { gate.Release(); }
        }));
        int ownerless = checks.Count(c => c.ownerless);
        int empty = checks.Count(c => c.empty);

        // Distinct admin role holders across all directory roles (members paged).
        var roles = await G().DirectoryRoles.GetAsync();
        var admins = new HashSet<string>();
        foreach (var role in roles?.Value ?? new())
        {
            var mem = await G().DirectoryRoles[role.Id].Members.GetAsync();
            int rp = 0;
            while (mem is not null)
            {
                foreach (var m in mem.Value ?? new()) if (m.Id != null) admins.Add(m.Id);
                var mnext = mem.OdataNextLink;
                if (string.IsNullOrEmpty(mnext) || ++rp >= 50) break;
                mem = await G().DirectoryRoles[role.Id].Members.WithUrl(mnext).GetAsync();
            }
        }

        // analyzed/sampled let the UI say the ownerless/empty figures come from a sample when a
        // tenant has more unified groups than the per-request cap, instead of implying full coverage.
        return Results.Json(new { groups = unifiedAll.Count, ownerless, empty, admins = admins.Count,
                                  analyzed = toCheck.Count, sampled = unifiedAll.Count > toCheck.Count });
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

// Compares two dotted version strings, ignoring a leading "v". >0 when a is newer than b.
static int CompareVersions(string a, string b)
{
    static int[] Parts(string s) => s.TrimStart('v', 'V').Split('.', '-')
        .Select(p => int.TryParse(p, out var n) ? n : 0).ToArray();
    var pa = Parts(a);
    var pb = Parts(b);
    for (int i = 0; i < Math.Max(pa.Length, pb.Length); i++)
    {
        int x = i < pa.Length ? pa[i] : 0, y = i < pb.Length ? pb[i] : 0;
        if (x != y) return x.CompareTo(y);
    }
    return 0;
}

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

record TenantDto(string? id, string label, string tenantId, string clientId, string clientSecret);
record TenantProfile(string Id, string Label, string TenantId, string ClientId, string ClientSecret);
record TenantStore(string? ActiveId, List<TenantProfile> Tenants);
record SnapshotDto(Dictionary<string, double?> metrics);
record Snapshot(string date, Dictionary<string, double?> metrics);
