# Development Notes

Lessons learned and implementation notes from building `onenote-cli`.

## Architecture Decisions

### Bun + Yargs

- Bun is used as the runtime, allowing direct execution of TypeScript without a build step
- Yargs provides a natural command/subcommand CLI structure
- The `bin` field in package.json points directly to `src/index.ts` — no compilation needed

### MSAL Node for Authentication

- `@azure/msal-node` provides the `PublicClientApplication` class for device code flow
- Token caching is implemented via MSAL's `cachePlugin` interface, persisting to `~/.onenote-cli/msal-cache.json`
- Silent token acquisition is attempted first (using cached refresh tokens), falling back to device code flow

### Configuration Priority

Configuration is resolved in this order:
1. Environment variables (`ONENOTE_CLIENT_ID`, `ONENOTE_AUTHORITY`)
2. Config file (`~/.onenote-cli/config.json`)
3. Default values (which prompt the user to configure)

This allows both `.env.local` for development and the config file for installed usage.

## Implementation Details

### Device Code Flow

The device code flow is ideal for CLI applications because:
- No web server or redirect handler is needed
- Works in headless/SSH environments
- User authenticates in any browser, even on a different device

Flow:
1. CLI requests a device code from Azure AD
2. Azure AD returns a code and a URL
3. User opens the URL in a browser and enters the code
4. User authenticates and consents to permissions
5. CLI polls Azure AD and receives tokens once authentication completes

### Page Creation

OneNote pages are created by POSTing raw HTML (Content-Type: `text/html`), not JSON. The HTML must include a `<title>` in the `<head>` and content in the `<body>`. This is different from most Graph API endpoints which use JSON.

### App-Only Auth Not Supported

As of March 2025, Microsoft Graph OneNote API no longer supports app-only (client credentials) authentication. All access must use delegated permissions with a signed-in user. This means the CLI must always go through the device code flow for initial authentication.

## Azure Portal Navigation Tips

### Entra Admin Center vs Azure Portal

- App registrations are managed in the [Microsoft Entra admin center](https://entra.microsoft.com), not the classic Azure portal
- The path is: **Entra ID** > **App registrations**
- Direct URL: `https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade`

### Common Setup Mistakes

1. **Forgetting to enable public client flows** — Device code flow requires "Allow public client flows" to be set to Yes in Authentication > Settings. Without this, authentication fails with `AADSTS7000218`.

2. **Wrong account type selection** — For a CLI that should work with both personal and work accounts, select "Accounts in any organizational directory and personal Microsoft accounts".

3. **Missing redirect URI** — While device code flow doesn't strictly require a redirect URI, adding `https://login.microsoftonline.com/common/oauth2/nativeclient` ensures compatibility with the MSAL library.

4. **Not adding API permissions** — The default `User.Read` permission is not enough. You must explicitly add `Notes.Read` / `Notes.ReadWrite` delegated permissions under Microsoft Graph.

## Troubleshooting

### Token Refresh Failures

If silent token acquisition fails after a long period, the refresh token may have expired. Run `onenote login` again to re-authenticate.

### Graph API 403 Errors

Usually means the required permission was not consented. Check that:
- The permission is added in the app registration
- The user consented during login (or admin consent was granted)

### Graph API 404 Errors

OneNote resources use opaque IDs. If a notebook/section/page was deleted or moved, its ID becomes invalid.
