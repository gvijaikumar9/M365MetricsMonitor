# M365 Metrics Monitor

A free, local, read-only Microsoft 365 dashboard. It runs on your own PC, hosts a
small web server on `localhost`, and opens a dashboard in your browser. One screen
with the numbers you check every morning, pulled from Microsoft Graph. Nothing
leaves the machine except the read-only Graph calls, and there is no account or
sign-up.

It is a **monitor**, not an admin tool. It shows you what needs attention across
service health, licences, storage, secrets, users, security and governance in one
place, so you do not have to click through five different admin portals. To *act* on
anything, you still use the Microsoft admin center.

## What it shows

- **Service health** (live) and the active advisories
- **Unused licences** and, with your cost rates entered in Settings, the monthly idle spend
- **Guest users**, **users by domain**
- **Storage used**, **sites by template**, **top sites by storage**
- **App secrets and certificates expiring** across all app registrations
- **Copilot adoption**, **workload activity**
- **Secure Score** and security posture
- **Governance**: ownerless groups, empty groups, admin role holders

Tiles are marked **LIVE** (current directory data) or **~2D** (from usage reports,
about one to two days behind).

## Run it

### Option A - the packaged app (no .NET needed)

1. Download and unzip the release folder.
2. Do the one-time app registration below and put your three values in `appsettings.json`.
3. Run `M365MetricsMonitor.exe`. The dashboard opens at `http://localhost:5137`.

### Option B - from source

```
dotnet run
```

Until `appsettings.json` is filled in, the dashboard runs on sample data and the
badge reads `SAMPLE`. Once configured it turns green and reads `LIVE DATA`.

## One-time app registration

Auth is **app-only (client credentials)**. No user sign-in, no device code, no
redirect URI. The app runs with its own identity using the Application permissions
you grant.

1. Entra admin center, **App registrations**, **New registration**. Name it
   `M365 Metrics Monitor`, single tenant.
2. **API permissions**, add these **Application** Microsoft Graph permissions, then
   **Grant admin consent**:
   - `Organization.Read.All`  (licences)
   - `User.Read.All`  (users, guests)
   - `Reports.Read.All`  (storage, workload, Copilot usage)
   - `Group.Read.All`  (governance groups)
   - `Directory.Read.All`  (app secrets, admin roles)
   - `ServiceHealth.Read.All`  (service health, message center)
   - `SecurityEvents.Read.All`  (Secure Score)
3. **Certificates & secrets** > **New client secret**. Copy the secret **Value**
   right away (it is shown only once).
4. Put three values into `appsettings.json`:
   - **Directory (tenant) ID** -> `Graph:TenantId`
   - **Application (client) ID** -> `Graph:ClientId`
   - the secret value -> `Graph:ClientSecret`

That is everything. There is nothing to sign in to.

### Optional (extended)

The **Stale accounts** and **guest inactivity** figures read `signInActivity`, which
needs `AuditLog.Read.All` plus an **Entra ID P1** licence. Without it those two stay
on sample values. Everything else works on the permissions above.

## Settings

Open **Settings** in the app to:
- enter a **monthly cost per licence** (Microsoft Graph does not expose prices), so the
  Unused Licences tile shows real idle spend,
- set alert thresholds.

Both are stored in your browser.

## Security

Your client secret lives in `appsettings.json`. It is **git-ignored** - never commit
it. Ship `appsettings.example.json` (placeholders) instead. If the secret ever leaks,
rotate it in the app registration.

## Build a single-file release

```
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
```

Output is in `bin/Release/net8.0/win-x64/publish`. Replace the published
`appsettings.json` with the placeholder version before sharing the folder.
