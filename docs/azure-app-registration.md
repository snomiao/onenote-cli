# Azure AD App Registration for OneNote CLI

This document describes how to register an Azure AD application for use with `onenote-cli`, which uses the Microsoft Graph OneNote API with delegated (device code flow) authentication.

## Prerequisites

- A Microsoft account (personal or organizational)
- Access to [Microsoft Entra admin center](https://entra.microsoft.com)

## Step 1: Create the App Registration

1. Go to [Microsoft Entra admin center](https://entra.microsoft.com)
2. Navigate to **Entra ID** > **App registrations** > **New registration**
3. Fill in:
   - **Name**: `onenote-cli`
   - **Supported account types**: "Accounts in any organizational directory and personal Microsoft accounts"
   - **Redirect URI**: Leave blank (configured in the next step)
4. Click **Register**
5. Copy the **Application (client) ID** from the Overview page

## Step 2: Configure Authentication Platform

1. Go to **Authentication** tab in your app registration
2. Click **Add a platform** > **Mobile and desktop applications**
3. Check these redirect URIs:
   - `https://login.microsoftonline.com/common/oauth2/nativeclient`
   - `msal{client-id}://auth` (MSAL-specific URI)
4. Click **Configure**

## Step 3: Enable Public Client Flows

1. In the **Authentication** tab, go to the **Settings** sub-tab
2. Find **Allow public client flows** and set it to **Enabled**
3. Click **Save**

This is required for device code flow authentication. Without it, you will get error `AADSTS7000218`.

## Step 4: Add API Permissions

1. Go to **API permissions** > **Add a permission**
2. Select **Microsoft Graph** > **Delegated permissions**
3. Search for "Notes" and expand the **Notes** group
4. Select:
   - `Notes.Read` — Read user OneNote notebooks
   - `Notes.ReadWrite` — Read and write user OneNote notebooks
   - `Notes.ReadWrite.All` — Read and write all notebooks the user can access
5. Click **Add permissions**

The default `User.Read` permission is already included.

## Step 5: Configure the CLI

Set the client ID via environment variable or config file:

```bash
# Option A: .env.local
ONENOTE_CLIENT_ID=your-client-id-here
ONENOTE_AUTHORITY=https://login.microsoftonline.com/common

# Option B: ~/.onenote-cli/config.json
{
  "clientId": "your-client-id-here",
  "authority": "https://login.microsoftonline.com/common"
}
```

Environment variables take priority over the config file.

## Step 6: Login and Verify

```bash
bun run src/index.ts login
bun run src/index.ts notebooks list
```

## Notes

- The Microsoft Graph OneNote API **does not support app-only authentication** as of March 2025. Only delegated (user) authentication is supported.
- The authority URL `https://login.microsoftonline.com/common` supports both organizational and personal Microsoft accounts.
- Tokens are cached at `~/.onenote-cli/msal-cache.json` and automatically refreshed.
