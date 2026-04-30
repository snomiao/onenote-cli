# onen0te-cli UX Analysis

Analysis of [fatihdumanli/onen0te-cli](https://github.com/fatihdumanli/onen0te-cli) — a Go-based OneNote CLI tool. Key takeaways for informing our `onenote-cli` design.

## Command Structure

```
nnote new [-i "text" | -f file] [-a alias] [-t title]   Create a note
nnote browse                                              Interactive notebook/section/page navigation
nnote search <phrase>                                     Search across all notebooks
nnote alias new [name]                                    Create a section alias
nnote alias list                                          List all aliases
nnote alias remove <name>                                 Delete an alias
```

## Auth Flow

- Uses **OAuth 2.0 Authorization Code Flow** (not device code)
- Starts a local HTTP server on `localhost:5992` for the redirect callback
- Opens system browser for login
- Tokens stored in Bitcask DB at `/tmp/nnote`
- Auto-refreshes expired tokens silently before each API call
- On first run, prompts: "You haven't setup a Onenote account yet, would you like to setup one now?"

Comparison with our approach:
- We use **device code flow** which works better in SSH/headless environments
- They use browser redirect which is more seamless on desktop
- Both auto-refresh tokens

## Key UX Patterns Worth Adopting

### 1. Interactive Selection Prompts

When creating a note without an alias, the user is prompted to select a notebook, then a section via interactive dropdown (uses `survey` library). This is much better than requiring users to copy-paste IDs.

### 2. Alias System

Maps a short name to a notebook+section pair. Avoids repeated interactive selection for frequently-used sections. Example:
```
nnote new -a work -i "Meeting notes"
```
After saving, if no alias exists for the section, suggests creating one.

### 3. Browse Mode (Interactive Navigation Loop)

`nnote browse` creates an interactive loop:
1. Select notebook
2. Select section
3. Select page
4. View content (HTML rendered to text)
5. Menu: back to sections / notebooks / open in browser / open in OneNote client / exit

Uses emoji indicators in menus for visual scanning.

### 4. Multiple Note Input Methods

- `-i "text"` — Inline text
- `-f /path/to/file` — Import from file
- No flags — Opens `$EDITOR` for composing

### 5. Spinner Animations

Shows spinners during API calls (GetNotebooks, GetSections, SaveNote, etc.) with success/fail status on completion.

### 6. Styled Output

- Color-coded messages: green=success, red=error, yellow=warning
- Table rendering for structured data (aliases, notebooks)
- Breadcrumb display with metadata when viewing pages

## Architecture Decisions

### Storage

- Uses **Bitcask** (embedded key-value store) at `/tmp/nnote`
- Stores both OAuth tokens and aliases in the same DB
- No config files — minimal setup burden
- Concern: `/tmp` is volatile on some systems; better to use `~/.config/` or `~/.local/`

### API Layer

- Custom REST client wrapper over `net/http`
- Endpoints used:
  - `GET /me/onenote/notebooks` — list notebooks
  - `GET /me/onenote/notebooks/{id}/sections` — list sections
  - `GET /me/onenote/sections/{id}/pages` — list pages
  - `GET /me/onenote/pages/{id}/content` — get page HTML
  - `POST /me/onenote/sections/{id}/pages` — create page
  - `GET /me/onenote/pages?search=...` — search (undocumented/deprecated?)
- Uses `html2text` library to convert page HTML to terminal-readable text

### Error Handling

- All errors wrapped with context via `pkg/errors`
- Exit codes for different failure modes (0-7)
- Styled error messages (not raw stack traces)

### Dependencies (Go)

| Library | Purpose |
|---------|---------|
| spf13/cobra | CLI framework |
| AlecAivazis/survey | Interactive prompts |
| pterm/pterm | Tables, spinners, styled output |
| k3a/html2text | HTML to text rendering |
| prologic/bitcask | Key-value storage |
| pkg/errors | Error wrapping |

## Notable Gaps

- No pagination for large result sets
- No retry logic for HTTP 503/504 (marked as TODO)
- No handling of the 5000-item SharePoint limit (would fail silently)
- Search uses an older/undocumented Graph API pattern
- Hardcoded OAuth client ID — not configurable
- Storage path `/tmp/nnote` is volatile
- No `--json` output option for scripting

## Ideas for onenote-cli

Based on this analysis, features worth considering for our CLI:

1. **Interactive prompts** — Use `@inquirer/prompts` or similar for notebook/section selection instead of requiring raw IDs
2. **Alias system** — Map short names to notebook+section pairs for quick note creation
3. **Browse mode** — Interactive navigation loop through notebooks → sections → pages
4. **HTML to text rendering** — Use `html-to-text` or `turndown` for terminal display of page content
5. **Spinner/progress indicators** — Show progress during API calls
6. **Open in browser/client** — Use the `links.oneNoteWebUrl` and `oneNoteClientUrl` from notebook metadata
7. **$EDITOR integration** — Launch editor for composing notes
8. **--json flag** — Output JSON for scripting/piping (improvement over onen0te-cli)
