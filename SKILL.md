---
name: onenote-cli
description: Search and operate Microsoft OneNote from the command line via Microsoft Graph API. Use when the user wants to search OneNote notes, list notebooks, read pages, or manage OneNote content.
---

# onenote-cli

CLI for Microsoft OneNote via Microsoft Graph. Built with Bun + yargs + MSAL.

## Quick Reference

```bash
onenote auth login                    # Device code flow login
onenote auth logout                   # Clear cached tokens
onenote auth whoami                   # Show current user
onenote auth setup                    # Show OAuth setup instructions

onenote notebooks list                # List notebooks
onenote notebooks get <id>            # Get notebook by ID
onenote notebooks create <name>       # Create notebook

onenote sections list -n <nb-id>      # List sections in a notebook
onenote sections create -n <nb> --name <name>

onenote pages list -s <sec-id>        # List pages in a section
onenote pages content <id>            # Get page HTML content
onenote pages create -s <sec> -t <title> -b <html>
onenote pages delete <id>

onenote sync                          # Build local cache (.one binary + page index)
onenote search <query>                # Full-text search across cached pages
onenote search <query> --online       # Online section-level search via Graph
```

## Setup

See [docs/setup.md](docs/setup.md) for full Azure AD app registration walkthrough.

Quick setup:
1. Register app at https://entra.microsoft.com → App registrations → New
2. Supported account types: "Accounts in any organizational directory and personal Microsoft accounts"
3. Add platform: Mobile and desktop applications, redirect URI `https://login.microsoftonline.com/common/oauth2/nativeclient`
4. Settings → Allow public client flows: **Yes**
5. API permissions → Microsoft Graph → Delegated: `Notes.Read`, `Notes.ReadWrite`, `Notes.ReadWrite.All`, `Files.Read`, `Files.Read.All`, `Sites.Read.All`
6. Copy Application (client) ID into `.env.local`:
   ```env
   ONENOTE_CLIENT_ID=your-client-id-here
   ONENOTE_AUTHORITY=https://login.microsoftonline.com/common
   ```

## How Search Works

`onenote search <query>` returns page-level results with **official OneNote URLs** that navigate directly to the matching page (bypassing OneNote Online's session caching).

The search:
1. Downloads `.one` files via OneDrive (cached locally in `<package>/.onenote/cache/`)
2. Extracts page GUIDs from the binary
3. Resolves official `oneNoteWebUrl` via `/me/onenote/sections/0-{guid}/pages` (works around the 5,000-item document library limit)
4. Searches the binary for the query (UTF-8 + UTF-16LE) and attributes matches to the correct page using context-based lookup

See [docs/local-search-architecture.md](docs/local-search-architecture.md) for details.

## File Locations

- `.env.local` — at the package root (auto-loaded via `import.meta.dir`)
- Token cache — `~/.onenote-cli/msal-cache.json`
- Config fallback — `~/.onenote-cli/config.json`
- Page cache — `<package>/.onenote/cache/` (override with `ONENOTE_CACHE_DIR`)

## Common Issues

| Error | Fix |
|---|---|
| `AADSTS7000218` | Enable "Allow public client flows" in app Authentication settings |
| `AADSTS65001` | Admin consent required for API permissions, or accept consent during login |
| `Graph API 403: error 10008` | Document library has > 5,000 OneNote items; use `onenote sync` and local search instead |
| Cache empty | Run `onenote sync` (auto-runs on first search) |
