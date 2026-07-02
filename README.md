# onenote-cli

**Your OneNote, in a terminal. Fluent with AI.**

```bash
onenote ls                                  # browse your notebooks
onenote read NotebookA/SectionB/PageC       # read a page, rendered as markdown
onenote search "visa interview"             # find any word, anywhere
onenote open NotebookA/SectionB             # jump back into OneNote Online
onenote append NotebookA/SectionB/today -c "- met with alex" --md
```

Every page, every section, every word — addressable by path, openable from a prompt, pipe-able into anything.

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

# Sections
onenote sections list [<ref>]                 # e.g. "NotebookA"
onenote sections create -n <ref> --name <name>

# Pages  (<ref> = "nb/sec/page", page ID, or OneNote URL)
onenote pages list [<ref>]                    # e.g. "NotebookA/SectionB"
onenote pages get <ref>
onenote pages create -s <ref> -t <title> --body "# Heading" --md
onenote pages append <ref> -c "- bullet" --md   # ⚠ may reformat existing content (Graph API limitation)
onenote pages update <ref> --target "#element-id" --action replace -c "<p>new</p>" --sha <4-char>
onenote pages delete <ref> --sha <4-char>

# Top-level shortcuts
onenote ls [<path>]                           # auto: notebooks / sections / pages
onenote read <ref>                            # render page (or list section/notebook)
onenote export <page-ref> <out.md>            # export a page to Markdown + downloaded media
onenote export <page-ref> <out.html>          # export as HTML (format inferred from extension)
onenote export <section-ref> <dir>            # dump a section to a folder (one .md per page)
onenote export <notebook-ref> <dir>           # dump a whole notebook as a directory tree
onenote open <ref>                            # open in browser
onenote rename <ref> <new-name>               # rename (depth inferred)
onenote rm <ref>                              # dry-run: prints content + sha
onenote rm <ref> --sha=<4-char>               # confirm deletion with content sha
onenote init                                  # first-run setup

# Search
onenote sync                # Download and cache all sections
onenote search <query>      # Full-text page-level search (local)
onenote search <query> -o   # Online section-level search (Graph API)
```

### Append caveat

`pages append` uses the Microsoft Graph OneNote PATCH API (`action: "append"`). This is safe — it never deletes existing content — but the Graph API re-parses the entire page on the server side, which can normalize or strip custom fonts, indentation, and other formatting from **existing** content. New content is added correctly.

If preserving exact formatting matters, edit the page directly in OneNote Online or the desktop app.

### Export

```
onenote export <ref> <output>
```

Exports a **page**, **section**, **section group**, or **whole notebook** to local files plus their media. What `<ref>` resolves to decides the shape:

| `<ref>` resolves to | `<output>` | Result |
| --- | --- | --- |
| page | a file (`.md`/`.html`) | one file |
| page | a directory | `<dir>/<page-title>.md` |
| section | a directory | one file per page |
| section group / notebook | a directory | a directory tree mirroring the structure, one file per page |

Details:

- **Format** — for a single page it's inferred from the output extension (`.html`/`.htm` → HTML, else Markdown). For directory exports it defaults to Markdown; use `--format md|html`.
- **Filenames** — pages are named by their sanitized title (Unicode preserved; only filesystem-illegal characters are stripped). Collisions get ` (2)`, ` (3)`… suffixes.
- **Frontmatter** — every Markdown file gets YAML frontmatter with `title`, `source` (the original OneNote URL), `notebook`/`section` (when known), and `exported` date. HTML exports get the same as `<meta name="onenote-*">` tags.
- **Media** — downloaded next to each file, into `./assets` by default (safe to share across pages — assets are content-addressed, so no clashes). Override with `--assets-dir <dir>`. Links are rewritten to local relative paths, so folders are portable.
- **`[date]`** in the output path expands to today's date (`YYYY-MM-DD`).
- **Large libraries** — notebook/section-group exports normally enumerate via the OneNote API. When a OneDrive library exceeds the Graph API's 5,000-item limit, export automatically falls back to walking the OneDrive file tree (`.one` files → sections, subfolders → section groups), so even huge accounts export fully.

```
# One page → Markdown + ./assets next to it
onenote export "<page-url>" './[date]-my-page/page.md'

# One page → HTML
onenote export NotebookA/SectionB/PageC ./out/page.html

# A whole section → a folder of Markdown files (one per page)
onenote export "<section-url>" ./out/my-section

# A whole notebook → a directory tree of sections/section groups
onenote export NotebookA ./backup/NotebookA

# Shared media directory + HTML
onenote export NotebookA ./backup --format html --assets-dir ./backup/media
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

For architecture, permissions, and internals see [docs/](docs/).

## Roadmap

### ✅ v0.1 — Foundation (shipped)

- Browse, read, edit, create, delete — notebooks, sections, pages
- Path refs (`Notebook/Section/Page`) across every command
- Full-text page-level search with context snippets
- Incremental local cache, works on huge libraries
- Markdown-in / markdown-out for reads and writes
- Export a page to Markdown or HTML with all media downloaded locally
- Clickable links in the terminal
- Installable as an AI agent skill

### 🚀 v0.2 — AI Native

Make onenote-cli the standard way AI accesses OneNote.

- **MCP server mode** — `onenote mcp` starts an MCP server on stdio. Unlocks Claude Desktop, Cursor, Continue, and any MCP client.
- **`onenote ask <question>`** — RAG over your notebook. Searches locally, fetches top N pages, feeds to LLM for answer.
- **`onenote summarize <url>`** — LLM summary of any page.
- **Semantic search** — local embeddings via `@xenova/transformers`.

### 📦 v0.3 — Portability

Own your data. Never locked in.

- ✅ **`onenote export`** — convert a page to Markdown or HTML with media downloaded to a local assets folder. _(shipped)_
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

## Support

Love this tool? Help keep it moving:

- 💛 [Sponsor / donate](https://github.com/snomiao) — options on my homepage
- 🤖 Gift AI credits (Claude / OpenAI / etc.) — this project is built with AI and every token goes straight back into shipping features

## License

MIT
