# onenote-cli

A command-line tool for Microsoft OneNote built with Bun, yargs, and Microsoft Graph API.

**Full-text search across all your OneNote pages with page-level deep links** — even on accounts with 5,000+ items that break the Graph API.

## Features

- **Page-level search** — search inside all your OneNote pages, get results with URLs that open directly to the matching page in OneNote Online
- **Notebooks / Sections / Pages** — list, get, create, delete via Graph API
- **5,000-item workaround** — when Graph OneNote API is blocked by the SharePoint document library limit, automatically falls back to OneDrive file API + local binary parsing
- **Local cache** — downloads `.one` files, extracts text (UTF-8 + UTF-16LE), builds a searchable index
- **Official OneNote URLs** — resolves page GUIDs via `GET /me/onenote/sections/0-{guid}/pages` to get URLs that bypass OneNote Online's session caching
- **Device code flow auth** — works in SSH / headless / terminal environments
- **Cross-directory** — `.env.local` and cache are loaded from the package directory, so `onenote` works from any working directory

## Install

```bash
git clone https://github.com/snomiao/onenote-cli.git
cd onenote-cli
bun install
```

### As a Claude Code Skill

```bash
npx skills add snomiao/onenote-cli
```

## Setup

1. Register an Azure AD app at [entra.microsoft.com](https://entra.microsoft.com) (see [docs/setup.md](docs/setup.md) for full walkthrough)
2. Copy `.env.example` to `.env.local` and set your Client ID:
   ```bash
   cp .env.example .env.local
   # Edit .env.local with your Application (client) ID
   ```
3. Login:
   ```bash
   bun run src/index.ts auth login
   ```

## Usage

```bash
# Auth
onenote auth login          # Login via device code flow
onenote auth whoami         # Show current user
onenote auth setup          # Print setup instructions
onenote auth logout         # Clear tokens

# Notebooks
onenote notebooks list
onenote notebooks get <id>
onenote notebooks create <name>

# Sections (falls back to OneDrive if Graph API is blocked)
onenote sections list -n <notebook-id>
onenote sections create -n <notebook-id> --name <name>

# Pages
onenote pages list -s <section-id>
onenote pages get <id>
onenote pages content <id>
onenote pages create -s <section-id> -t <title> -b "<html>"
onenote pages delete <id>

# Search
onenote sync                # Download and cache all sections
onenote search <query>      # Full-text page-level search (local)
onenote search <query> -o   # Online section-level search (Graph API)
```

### Search Example

```
$ onenote search project plan

# (2024-11-03) Meeting notes
  Section: Work | Notebook: MyNotebook
  **project plan** ...
  https://contoso.sharepoint.com/.../Notebooks/MyNotebook?wd=target(...)

2 page-level results found.
```

Clicking the URL opens OneNote Online directly on the matching page.

## How Search Works

1. `onenote sync` downloads all `.one` section files from OneDrive
2. Extracts page GUIDs from the MS-ONESTORE binary format
3. Fetches official page URLs via OneNote Graph API (`/me/onenote/sections/0-{guid}/pages`)
4. On search, scans the binary for matches (UTF-8 and UTF-16LE), attributes each match to the correct page via context-based anchor lookup
5. Returns results with official OneNote URLs that navigate directly to the page

See [docs/local-search-architecture.md](docs/local-search-architecture.md) for the full technical design.

## API Permissions Required

| Permission | Type | Purpose |
|---|---|---|
| `Notes.Read` | Delegated | Read notebooks |
| `Notes.ReadWrite` | Delegated | Create/modify pages |
| `Notes.ReadWrite.All` | Delegated | Access all notebooks |
| `Files.Read` | Delegated | Download .one files from OneDrive |
| `Files.Read.All` | Delegated | Access all accessible files |
| `Sites.Read.All` | Delegated | Search via SharePoint listItem API |

## File Structure

```
src/
  index.ts      CLI entry point (yargs)
  auth.ts       MSAL device code flow + .env.local auto-loader
  graph.ts      Microsoft Graph API client (OneNote + OneDrive)
  cache.ts      Local cache, .one binary parser, page GUID extraction
docs/
  setup.md                     Azure AD registration walkthrough
  local-search-architecture.md Technical design of local search
  graph-api-endpoints.md       OneNote Graph API reference
  development-notes.md         Lessons learned
  onen0te-cli-analysis.md      Competitor UX analysis
```

## Configuration

| Source | Location | Priority |
|---|---|---|
| `.env.local` | Package root (auto-loaded) | Highest |
| `~/.onenote-cli/config.json` | Home directory | Fallback |
| `ONENOTE_CLIENT_ID` env var | Shell environment | Overrides all |

Cache location: `<package>/.onenote/cache/` (override with `ONENOTE_CACHE_DIR`)

## License

MIT
