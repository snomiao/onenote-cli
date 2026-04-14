# Roadmap

What onenote-cli is, where it's going, and how it gets there.

## ✅ v0.1 — Foundation (shipped)

- [x] MSAL device code flow authentication with auto token refresh
- [x] Notebooks / sections / section-groups / pages CRUD
- [x] 5,000-item SharePoint limit workaround via OneDrive fallback
- [x] Local `.one` binary cache with UTF-8 + UTF-16LE page extraction
- [x] Page GUID extraction from binary
- [x] Official OneNote page URL resolution via Graph API
- [x] Full-text page-level search with context snippets
- [x] `onenote read <url>` — page / section / notebook tree view
- [x] `onenote rename` / `append` / `delete` for pages, sections, notebooks
- [x] Incremental sync (compares `lastModifiedDateTime`)
- [x] Size limit for sync (skips sections > 200MB)
- [x] `.env.local` auto-load from package dir (cross-directory support)
- [x] Markdown links in non-TTY, OSC 8 hyperlinks in TTY
- [x] ANSI color output with keyword highlighting
- [x] Published to npm as `onenote-cli`
- [x] Installable as AI agent skill via `npx skills add snomiao/onenote-cli`
- [x] Modernized stack: Bun, oxlint, tsgo, husky, secretlint, semantic-release

---

## 🚀 v0.2 — AI Native (next)

Make onenote-cli the standard way AI accesses OneNote.

- [ ] **MCP server mode** — `onenote mcp` starts an MCP server on stdio.
      Unlocks Claude Desktop, Cursor, Continue, and any MCP client.
      *Why first: maximum leverage — one feature, infinite AI clients.*
- [ ] **`onenote ask <question>`** — RAG over your notebook.
      Searches locally, fetches top N pages, feeds to LLM for answer.
      Supports `--model claude-sonnet-4-6` / `--model gpt-5` / local via Ollama.
- [ ] **`onenote summarize <url>`** — LLM summary of any page.
- [ ] **Semantic search** — local embeddings via `@xenova/transformers`.
      `onenote search --semantic "things I forgot to finish"`

## 📦 v0.3 — Portability

Own your data. Never locked in.

- [ ] **`onenote export`** — convert pages to Markdown.
      `onenote export --all --out ./backup/` dumps everything.
      Preserves images, code blocks, tables, math.
- [ ] **`onenote import <dir>`** — ingest Markdown directory tree as notebook.
      Folders → sections, `.md` files → pages.
- [ ] **`onenote backup --incremental`** — daily diff-based snapshots.
      Cron-safe, resumable, git-friendly.
- [ ] **Obsidian / Notion bridges** — `--format obsidian` preserves `[[wiki-links]]`.

## 🎨 v0.4 — Terminal UI

Make it lovely to use by hand.

- [ ] **`onenote browse`** — interactive TUI with react-ink.
      Tree view: notebook → section → page. Arrow keys, `/` search,
      `Enter` to open, preview in split pane.
- [ ] **`onenote fzf`** — pipe all page titles to fuzzy finder.
- [ ] **Progress bars for sync** — real-time KB/s, ETA, spinner.
- [ ] **`onenote stats`** — total pages, size, word count, oldest note, etc.

## 🔧 v0.5 — Power User

Unix composition, scripting, automation.

- [ ] **Stdin piping** — `echo "note" | onenote append <id>`.
- [ ] **`--json` flag everywhere** — structured output for `jq` / scripts.
- [ ] **`onenote watch <file> --to-page <id>`** — live sync file → page.
- [ ] **`onenote new --daily`** — auto-create dated journal page with template.
- [ ] **Clipboard capture** — `onenote clip <id>` reads clipboard → page.
- [ ] **Screenshot + OCR** — `onenote snap <id>` captures screen → OCR → page.

## 🤝 v0.6 — Collaboration

Teams, shared notebooks, change tracking.

- [ ] **Group notebooks** — `/groups/{id}/onenote/` support.
- [ ] **SharePoint site notebooks** — `/sites/{id}/onenote/`.
- [ ] **`onenote share <id> --user <email>`** — permissions management.
- [ ] **`onenote diff <url>`** — show changes since last sync.
- [ ] **`onenote subscribe <url>`** — webhooks on page changes.

## 🧠 v0.7 — Insights

Understand your own knowledge graph.

- [ ] **`onenote graph --out graph.html`** — d3.js interactive page graph.
      Nodes = pages, edges = cross-references.
- [ ] **`onenote duplicates`** — find and merge similar pages.
- [ ] **`onenote timeline`** — activity heatmap by date.
- [ ] **`onenote links <id>`** — incoming/outgoing references for a page.

## 🔐 v0.8 — Privacy & Local-First

- [ ] **Offline-only mode** — `--offline` flag uses only local cache.
- [ ] **SQLite FTS5 index** — sub-millisecond full-text search.
- [ ] **End-to-end encryption** — encrypt pages with `age` before sync.
- [ ] **Self-hosted relay** — optional proxy for air-gapped environments.

## 🛠 v1.0 — Platform

Onenote-cli is no longer a CLI — it's a toolkit.

- [ ] **SDK** — `import { OneNote } from "onenote-cli"`.
- [ ] **HTTP API** — `onenote serve --port 3000` exposes OpenAPI 3.1.
- [ ] **Plugin system** — `onenote plugin install @company/jira-sync`.
- [ ] **Raycast / Alfred extensions** — native launcher integration.
- [ ] **Browser extension** — right-click → save to OneNote.
- [ ] **Proper MS-ONESTORE parser** — 100% accurate page attribution,
      replaces the current heuristic binary parser.

---

## Not planned

- GUI / Electron app (OneNote Online exists)
- Mobile apps (use Microsoft OneNote app)
- Rich text editor (keep it CLI-first)
- User management (stay delegated auth only)

---

## Contributing

Pick any unchecked item. Open an issue to claim it, or just ship a PR.

Foundation is solid — the Graph API works, the binary parser works,
auth works. Every feature above is additive.

**Current priority**: v0.2 MCP server mode. Start there for maximum impact.
