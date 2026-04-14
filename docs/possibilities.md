# Possibilities — What onenote-cli could become

A vision document: every imaginable feature, with use cases.

## 1. AI-Native OneNote

### `onenote ask`
Retrieval-augmented Q&A over your entire notebook history.

```bash
$ onenote ask "what did I decide about the auth architecture?"

[Searching 153 sections...]
Found 3 relevant pages:
  1. (20250312) Architecture decision log — 95% match
  2. (2024-09-02) Security review notes — 78% match
  3. (2024-09-03) Follow-up — 62% match

Based on your notes:
- You chose Device Code Flow over OAuth redirect (20250312) because
  the CLI needs to work in headless SSH environments.
- Security review flagged token caching in ~/.onenote-cli/ as acceptable
  given OS-level file permissions (20250318).
```

### `onenote summarize <url>`
```bash
$ onenote summarize "https://.../page.one?wd=target(...)"

Page: Q2 Planning Meeting (2025-03-15)
─────────────────────────────────────────
Decisions:
- Ship v2 API by April 30
- Deprecate v1 endpoints after 90-day notice
- Hire 2 backend engineers

Action items:
- [alice] Draft migration guide by 3/22
- [bob] Benchmark new auth flow
```

### `onenote search --semantic "tasks I forgot to finish"`
Uses local embeddings (Xenova/transformers.js) — no API cost, runs offline.
Finds notes by meaning, not just keywords. Great for surfacing forgotten work.

---

## 2. MCP Server Mode

```bash
$ onenote mcp
# Listening on stdio for MCP requests...
```

Add to `~/.config/claude/mcp.json`:
```json
{
  "mcpServers": {
    "onenote": {
      "command": "bunx",
      "args": ["onenote-cli", "mcp"]
    }
  }
}
```

Now in Claude Desktop / Cursor / Continue:
> "Find all my notes tagged #meeting from last month"
> → Claude directly queries your OneNote via MCP

---

## 3. Content Operations

### Import from Markdown
```bash
$ onenote import ./journal/2025/
```
Imports a directory structure:
```
journal/
  2025/
    work/         → becomes Section "work"
      q1.md       → becomes Page "q1"
      q2.md       → becomes Page "q2"
    personal/     → becomes Section "personal"
      goals.md    → becomes Page "goals"
```

### Export to Markdown
```bash
$ onenote export --all --format markdown --out ./onenote-backup/
Exported 1,547 pages across 22 notebooks in 4m 12s.
```

### Move/Copy
```bash
$ onenote pages move <page-id> --to-section <section-id>
$ onenote sections copy <section-id> --to-notebook <notebook-id>
```

---

## 4. Piping & Composition

### Unix philosophy: stdin/stdout
```bash
# Pipe content into a page
$ date | onenote append <page-id>

# Search, read, pipe to LLM
$ onenote search "meeting" --json \
  | jq '.results[0].webUrl' \
  | xargs onenote read \
  | claude "extract action items as JSON"

# Git-hook: log every commit to OneNote
$ git log -1 --pretty=format:"%h %s" | onenote append <dev-log-page>

# Screenshot → OCR → OneNote
$ screencapture -i /tmp/s.png && \
  ocr /tmp/s.png | onenote append <page-id>
```

### JSON output for scripting
```bash
$ onenote search "invoice" --json | jq '.results[] | select(.modified > "2025-01-01")'
```

---

## 5. Automation & Workflows

### Daily journal
```bash
$ onenote new --daily --section <journal-section>
# Creates "2025-04-14" page using template
# Pre-fills: weather, calendar events, task list
```

### Watch mode
```bash
$ onenote watch ./notes.md --to-page <id>
# Auto-syncs notes.md → OneNote page on every save
```

### Scheduled backup
```bash
$ onenote backup --cron "0 3 * * *" --out ~/onenote-backup/
# Cron-safe, resumable, incremental
```

---

## 6. Rich Terminal UI

### Interactive browse
```bash
$ onenote browse
┌─ OneNote ─────────────────────────────────────┐
│ ▸ Archive        (40 sections)            │
│ ▾ MyNotebook            (44 sections)             │
│    ▾ Work       (46 pages)                │
│       • (2024-09-01) Auth architecture [opened] │
│       • (2024-09-02) Security review            │
│       • (2024-09-03) Follow-up       │
│ ▸ Notebook6          (22 sections)              │
└───────────────────────────────────────────────┘
Navigate: ↑↓ Expand: → Open: Enter Search: /
```

Built with react-ink. Arrow-key navigation, live search, page preview in split pane.

### Fuzzy finder
```bash
$ onenote fzf
# Pipes all page titles to fzf for instant fuzzy search
```

---

## 7. AI Agent Integrations

### As a Claude skill
Already works: `npx skills add snomiao/onenote-cli`

### As a cron-triggered agent
```yaml
# .github/workflows/onenote-summary.yml
- run: |
    onenote search "this-week-project" --json \
    | llm "write weekly update in markdown" \
    | onenote pages create --section <weekly-notes>
```

### Voice-driven
```bash
$ say-to-onenote "create a page in work called bug triage with the following"
# Uses speech-to-text → onenote pages create
```

---

## 8. Multi-User / Team Features

### Shared notebooks
```bash
$ onenote groups notebooks list                      # M365 group notebooks
$ onenote sites notebooks list <sharepoint-url>      # SharePoint site notebooks
$ onenote share <notebook-id> --user alice@corp.com
```

### Change detection
```bash
$ onenote diff <url>
# Shows what changed in a page since last sync

$ onenote subscribe <url>
# Notifies on any change via webhook/email/desktop notification
```

---

## 9. Rich Content

### Markdown-first page creation
```bash
$ onenote pages create --section <id> --from-file draft.md
# Auto-converts: code blocks, tables, images, math (KaTeX)
```

### Embed images/files
```bash
$ onenote append <page-id> --image ./diagram.png --caption "System overview"
$ onenote append <page-id> --file ./report.pdf
```

### Clipboard capture
```bash
$ onenote clip <page-id>
# Reads clipboard (text + images) → appends to page
# Works with screenshot tools
```

---

## 10. Insights & Analytics

### Stats
```bash
$ onenote stats

Notebooks:     22
Sections:      153
Pages:         1,547
Total size:    3.2 GB
Oldest note:   2019-02-03
Newest note:   2025-04-14
Most active:   Today (58 pages)
Word count:    2.1M words
```

### Graph view
```bash
$ onenote graph --out graph.html
# Generates interactive HTML: pages as nodes, links between pages as edges
# Uses d3.js or cytoscape
```

### Duplicates detection
```bash
$ onenote duplicates
Found 14 likely duplicate pages:
  "Meeting notes (1)" vs "Meeting notes" (98% similar)
  "Todo" vs "TODO" in same section
  ...
Run `onenote duplicates --merge` to merge interactively.
```

---

## 11. Developer Experience

### SDK mode
```ts
import { OneNote } from "onenote-cli";

const on = new OneNote();
const results = await on.search("meeting");
const page = await on.read(results[0].url);
```

### Plugin system
```bash
$ onenote plugin install @company/onenote-jira-sync
# Plugins can add commands, transforms, webhooks
```

### OpenAPI spec
```bash
$ onenote spec --out openapi.yaml
# Generates an OpenAPI 3.1 spec of all commands for REST gateway mode
$ onenote serve --port 3000
# Runs as HTTP API
```

---

## 12. Platform Integrations

### Obsidian
```bash
$ onenote export --format obsidian --out ~/Obsidian/OneNote/
# Preserves [[wiki-links]] where pages reference each other
```

### Notion
```bash
$ onenote migrate notion --workspace <url>
# Two-way sync with Notion
```

### Git-backed
```bash
$ onenote init --git ~/onenote-repo
# Every sync commits changes — full version history
$ cd ~/onenote-repo && git log
```

### Raycast / Alfred
```bash
$ onenote raycast-extension
# Generates a Raycast command pack
```

---

## 13. Privacy & Local-First

### End-to-end encryption
```bash
$ onenote encrypt --page-id <id> --key ~/.onenote-cli/age.key
# Encrypts page content with age; only decrypts on your machine
```

### Offline mode
```bash
$ onenote search "..." --offline
# Works entirely from local cache, no network
```

### Full local SQLite FTS
```bash
$ onenote index rebuild
# Build SQLite FTS5 index for sub-millisecond search
```

---

## 14. Which of these would you ship next?

Current status:
- ✅ Search (local, page-level, official URLs)
- ✅ Read (page/section/notebook)
- ✅ Create/delete/rename pages, sections, notebooks
- ✅ Incremental sync with size limits
- ✅ `.env.local` auto-load, cross-directory
- ✅ Markdown/OSC 8 dual output
- ✅ Installable as AI agent skill

Next candidates (see `TODO.md` for details):
- 🎯 MCP server mode (highest leverage — unlocks Claude Desktop, Cursor, etc.)
- 🎯 Export to markdown (backup + portability)
- 🎯 `ask` command with RAG (flagship AI feature)
- 🎯 Browse mode TUI (best for humans)

Or invent something new — the binary works, the URLs work, the auth works.
What do you want OneNote to *feel* like?
