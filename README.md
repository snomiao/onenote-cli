# onenote-cli

**Make your OneNote notebooks survive in the age of AI.**

A CLI that lets AI agents (and humans) search, read, and operate your OneNote — with full-text page-level search and deep links that open directly to the matching page.

## Features

- **Page-level search** — search inside all your OneNote pages, get results with URLs that open directly to the matching page in OneNote Online
- **Notebooks / Sections / Pages** — list, get, create, update, delete via Graph API
- **5,000-item workaround** — when Graph OneNote API is blocked by the SharePoint document library limit, automatically falls back to OneDrive file API + local binary parsing
- **Local cache** — downloads `.one` files, extracts text (UTF-8 + UTF-16LE), builds a searchable index
- **Official OneNote URLs** — resolves page GUIDs via `GET /me/onenote/sections/0-{guid}/pages` to get URLs that bypass OneNote Online's session caching
- **Device code flow auth** — works in SSH / headless / terminal environments
- **Cross-directory** — `.env.local` and cache are loaded from the package directory, so `onenote` works from any working directory

## Install

### As an AI Agent Skill

```bash
# Claude Code / OpenClaw / Codex / Cursor / any SKILL.md-compatible agent
npx skills add snomiao/onenote-cli
```

### Manual

```bash
git clone https://github.com/snomiao/onenote-cli.git
cd onenote-cli
bun install
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

# Notebooks  (<ref> = name, path, Graph ID, or URL)
onenote notebooks list
onenote notebooks get <ref>
onenote notebooks create <name>

# Sections (falls back to OneDrive if Graph API is blocked)
onenote sections list [<ref>]                 # e.g. "sno@ja" or notebook ID/URL
onenote sections create -n <ref> --name <name>

# Pages  (<ref> = "nb/sec/page", page ID, or OneNote URL)
onenote pages list [<ref>]                    # e.g. "sno@ja/visa@ja"
onenote pages get <ref>
onenote pages content <ref>                   # deprecated, use 'read --html'
onenote pages create -s <ref> -t <title> -b "<html>"
onenote pages create -s <ref> -t <title> --body "# Heading" --md
onenote pages append <ref> -c "- bullet" --md
onenote pages update <ref> --target "#element-id" --action replace -c "<p>new</p>"
onenote pages delete <ref>

# Top-level shortcuts
onenote ls [<path>]                           # auto: notebooks / sections / pages
onenote read <ref>                            # render page (or list section/notebook)
onenote open <ref>                            # open in browser
onenote mv <ref> <new-name>                   # rename (depth inferred)
onenote rm <ref>                              # delete page
onenote init                                  # first-run setup

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

## Roadmap

### ✅ v0.1 — Foundation (shipped)

- MSAL device code flow authentication with auto token refresh
- Notebooks / sections / section-groups / pages CRUD
- 5,000-item SharePoint limit workaround via OneDrive fallback
- Local `.one` binary cache with UTF-8 + UTF-16LE page extraction
- Page GUID extraction from binary
- Official OneNote page URL resolution via Graph API
- Full-text page-level search with context snippets
- `onenote read <url>` — page / section / notebook tree view
- Page editing via `rename` / `append` / `update` / `delete`
- URL-or-ID refs for `sections list`, `pages list`, `pages get`, `pages content`, and page write commands
- Markdown input for `pages create` / `append` / `update` via `--md`
- Incremental sync (compares `lastModifiedDateTime`)
- Size limit for sync (skips sections > 200MB)
- `.env.local` auto-load from package dir (cross-directory support)
- Markdown links in non-TTY, OSC 8 hyperlinks in TTY
- Published to npm, installable as AI agent skill via `npx skills add snomiao/onenote-cli`

### 🚀 v0.2 — AI Native

Make onenote-cli the standard way AI accesses OneNote.

- **MCP server mode** — `onenote mcp` starts an MCP server on stdio. Unlocks Claude Desktop, Cursor, Continue, and any MCP client.
- **`onenote ask <question>`** — RAG over your notebook. Searches locally, fetches top N pages, feeds to LLM for answer.
- **`onenote summarize <url>`** — LLM summary of any page.
- **Semantic search** — local embeddings via `@xenova/transformers`.

### 📦 v0.3 — Portability

Own your data. Never locked in.

- **`onenote export`** — convert pages to Markdown (preserves images, code blocks, tables, math).
- **`onenote import <dir>`** — ingest Markdown directory tree as notebook.
- **`onenote backup --incremental`** — daily diff-based snapshots.
- **Obsidian / Notion bridges** — `--format obsidian` preserves `[[wiki-links]]`.

### 🎨 v0.4 — Terminal UI

- **`onenote browse`** — interactive TUI with react-ink (tree view, arrow keys, `/` search, split-pane preview).
- **`onenote fzf`** — pipe all page titles to fuzzy finder.
- **Progress bars for sync** — real-time KB/s, ETA, spinner.
- **`onenote stats`** — total pages, size, word count, oldest note, etc.

### 🔧 v0.5 — Power User

Unix composition, scripting, automation.

- **Stdin piping** — `echo "note" | onenote append <id>`.
- **`--json` flag everywhere** — structured output for `jq` / scripts.
- **`onenote watch <file> --to-page <id>`** — live sync file → page.
- **`onenote new --daily`** — auto-create dated journal page with template.
- **Clipboard capture** — `onenote clip <id>` reads clipboard → page.
- **Screenshot + OCR** — `onenote snap <id>` captures screen → OCR → page.

### 🤝 v0.6 — Collaboration

- **Group notebooks** — `/groups/{id}/onenote/` support.
- **SharePoint site notebooks** — `/sites/{id}/onenote/`.
- **`onenote share <id> --user <email>`** — permissions management.
- **`onenote diff <url>`** — show changes since last sync.
- **`onenote subscribe <url>`** — webhooks on page changes.

### 🧠 v0.7 — Insights

- **`onenote graph --out graph.html`** — d3.js interactive page graph (nodes = pages, edges = cross-references).
- **`onenote duplicates`** — find and merge similar pages.
- **`onenote timeline`** — activity heatmap by date.
- **`onenote links <id>`** — incoming/outgoing references for a page.

### 🔐 v0.8 — Privacy & Local-First

- **Offline-only mode** — `--offline` flag uses only local cache.
- **SQLite FTS5 index** — sub-millisecond full-text search.
- **End-to-end encryption** — encrypt pages with `age` before sync.
- **Self-hosted relay** — optional proxy for air-gapped environments.

### 🛠 v1.0 — Platform

- **SDK** — `import { OneNote } from "onenote-cli"`.
- **HTTP API** — `onenote serve --port 3000` exposes OpenAPI 3.1.
- **Plugin system** — `onenote plugin install @company/jira-sync`.
- **Raycast / Alfred extensions** — native launcher integration.
- **Browser extension** — right-click → save to OneNote.
- **Proper MS-ONESTORE parser** — 100% accurate page attribution (replaces heuristic).

### Not planned

- GUI / Electron app (OneNote Online exists)
- Mobile apps (use Microsoft OneNote app)
- Rich text editor (keep it CLI-first)
- User management (stay delegated auth only)

## Contributing

Pick any unchecked roadmap item. Open an issue to claim it, or just ship a PR.

## License

MIT
