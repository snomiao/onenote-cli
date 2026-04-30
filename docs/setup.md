# onenote-cli Authentication Setup Guide

This guide walks you through setting up Azure AD authentication for `onenote-cli` to access Microsoft OneNote via the Graph API.

## Prerequisites

- A Microsoft account (personal or organizational)
- An Azure account with an active subscription — [create one for free](https://azure.microsoft.com/pricing/purchase-options/azure-account)
- [Bun](https://bun.sh) runtime installed

## Step 1: Register an Azure AD Application

1. Sign in to the [Microsoft Entra admin center](https://entra.microsoft.com)
2. Navigate to **Entra ID** > **App registrations** > **New registration**
3. Fill in the registration form:
   - **Name**: `onenote-cli`
   - **Supported account types**: Select **"Accounts in any organizational directory and personal Microsoft accounts"**
     (This allows both work/school and personal Microsoft accounts)
   - **Redirect URI**: Leave blank for now
4. Click **Register**
5. On the Overview page, copy the **Application (client) ID** — you will need this

## Step 2: Configure Platform for Device Code Flow

1. In your app registration, go to **Authentication** > **Add a platform**
2. Select **Mobile and desktop applications**
3. Check the box for `https://login.microsoftonline.com/common/oauth2/nativeclient`
4. Click **Configure**
5. Scroll down and set **Allow public client flows** to **Yes**
6. Click **Save**

> Device code flow requires "Allow public client flows" to be enabled. Without this, authentication will fail.

## Step 3: Configure API Permissions

1. In your app registration, go to **API permissions** > **Add a permission**
2. Select **Microsoft Graph** > **Delegated permissions**
3. Search and add the following permissions:
   - `Notes.Read` — Read user OneNote notebooks
   - `Notes.ReadWrite` — Read and write user OneNote notebooks
   - `Notes.Read.All` — Read all notebooks the user can access
   - `Notes.ReadWrite.All` — Read and write all notebooks the user can access
4. Click **Add permissions**

> For organizational accounts, an admin may need to **Grant admin consent** for these permissions.

## Step 4: Configure onenote-cli

You have two options to provide credentials:

### Option A: Environment variables (recommended)

Copy `.env.example` to `.env.local` and fill in your client ID:

```bash
cp .env.example .env.local
```

Edit `.env.local`:

```env
ONENOTE_CLIENT_ID=your-actual-client-id-here
ONENOTE_AUTHORITY=https://login.microsoftonline.com/common
```

### Option B: Config file

The CLI also reads from `~/.onenote-cli/config.json`. This file is auto-created on first run:

```json
{
  "clientId": "your-actual-client-id-here",
  "authority": "https://login.microsoftonline.com/common"
}
```

> Environment variables take priority over the config file.

## Step 5: Login

```bash
bun run src/index.ts login
```

This initiates the **device code flow**:

1. The CLI prints a URL and a code
2. Open the URL in your browser
3. Enter the code and sign in with your Microsoft account
4. Grant the requested permissions
5. Return to the terminal — you should see "Login successful!"

Tokens are cached at `~/.onenote-cli/msal-cache.json` so you don't need to login every time.

## Step 6: Verify

```bash
# List your notebooks
bun run src/index.ts notebooks list

# List all sections
bun run src/index.ts sections list

# List pages
bun run src/index.ts pages list
```

## Authority URL Reference

| Scenario | Authority URL |
|---|---|
| Multi-tenant + personal accounts | `https://login.microsoftonline.com/common` |
| Organizational accounts only (any tenant) | `https://login.microsoftonline.com/organizations` |
| Personal Microsoft accounts only | `https://login.microsoftonline.com/consumers` |
| Single tenant only | `https://login.microsoftonline.com/{tenant-id}` |

## Logout

To clear cached tokens:

```bash
bun run src/index.ts logout
```

## Troubleshooting

### "AADSTS7000218: The request body must contain the following parameter: 'client_assertion' or 'client_secret'"

→ Make sure **Allow public client flows** is set to **Yes** in your app's Authentication settings.

### "AADSTS65001: The user or administrator has not consented to use the application"

→ An admin needs to grant consent for the API permissions, or the user needs to consent during login.

### "AADSTS700016: Application with identifier '...' was not found"

→ Double-check that your `ONENOTE_CLIENT_ID` matches the Application (client) ID from Azure portal.

### Token expired

Tokens are automatically refreshed using the cached refresh token. If refresh fails, run `onenote login` again.
