# TODO

## High Priority

### npm publish
Publish to npm registry so users can run without cloning:
```bash
bunx onenote-cli search "keyword"
bunx onenote-cli notebooks list
```
Requires: add `files` field to package.json, set `"publishConfig"`, test with `npm pack`.

### Incremental sync
Current `onenote sync` re-downloads ALL `.one` files every time cache expires (1 hour).
Should compare `lastModifiedDateTime` from OneDrive API against cached timestamp,
and only download sections that actually changed. Expected 10x speedup for daily use.
Also store `eTag` from driveItem for more precise change detection.

### Sync progress display
Current output: `[downloading] section-name...` with no indication of progress.
Should show: download speed (KB/s), file size, progress bar, total progress (N/M sections).
Consider using `react-ink` for a dynamic TUI with real-time updates, or simpler
ANSI escape codes for inline progress. Also show indexing progress (pages extracted).

### Snippet noise cleanup
Search context still contains binary garbage like `픠३ᇓက❚떙` and `䴀匀 䜀漀琀栀椀挀`
(shifted UTF-16LE reads of "MS Gothic" font names). Improve `cleanSnippet()` to:
- Detect and strip shifted-ASCII CJK patterns (chars where high byte is ASCII)
- Remove OneNote internal metadata strings (font names, GUIDs, PageTitle markers)
- Keep only natural language text around the match

### Pagination for 100+ page sections
OneNote API enforces `$top=100` max per request. Sections with 100+ pages
only get the first 100 in `getOneNotePagesForSection()`. Should follow
`@odata.nextLink` to fetch all pages. Affects official URL resolution —
pages beyond 100 won't get official deep link URLs.

## Medium Priority

### `onenote open <query>`
Search and immediately open the first (or Nth) result in the default browser:
```bash
onenote open "meeting notes"        # opens first match
onenote open "meeting notes" -n 3   # opens 3rd match
```
Use `Bun.spawn(["open", url])` on macOS, `xdg-open` on Linux, `start` on Windows.

### `onenote export`
Export page content as markdown for use in other tools:
```bash
onenote export <page-id>              # stdout markdown
onenote export <page-id> -o file.md   # write to file
onenote export --section <id>         # export all pages in section
```
Use Graph API `GET /me/onenote/pages/{id}/content` for HTML, then convert
with `turndown` or `html-to-text`. Falls back to cached binary text if API blocked.

### Browse mode
Interactive terminal UI for navigating notebooks → sections → pages:
```bash
onenote browse
```
Using `@inquirer/prompts` (or react-ink) for selection menus.
Display page content as rendered text (HTML → terminal). Navigation options:
back, open in browser, open in OneNote client, copy URL.
Reference: onen0te-cli's browse implementation (see docs/onen0te-cli-analysis.md).

### Group/Site notebooks
Currently only supports `/me/onenote/`. Should also support:
- `/groups/{groupId}/onenote/` — Microsoft 365 group notebooks
- `/sites/{siteId}/onenote/` — SharePoint site notebooks
Add `--group <id>` and `--site <url>` flags to notebook/section/page commands.
Requires resolving group display names to IDs and SharePoint URLs to site IDs.

### `--json` output flag
Add `--json` or `-j` flag to all commands for machine-readable output:
```bash
onenote search "keyword" --json | jq '.results[].webUrl'
onenote notebooks list --json
```
Structured as `{ results: [...], total: N }` for search,
`{ items: [...] }` for list commands.

### Fuzzy search
Current search is exact substring match on binary content.
Add fuzzy/approximate matching for typo tolerance:
- Levenshtein distance for short queries
- Trigram similarity for longer queries
- Or integrate a lightweight full-text search library (e.g., `minisearch`, `fuse.js`)
Could also build a SQLite FTS5 index from cached text for fast queries.

## Low Priority

### GitHub Actions CI
Add `.github/workflows/ci.yml`:
- On push/PR: `bun install` → `bun run lint` → `bun run typecheck`
- On main push: semantic-release for automated npm publish + GitHub release
- Use `oven-sh/setup-bun@v2` action

### Tests
Add test coverage for critical paths:
- `src/auth.ts` — mock MSAL, test token refresh logic, env loading
- `src/cache.ts` — test `extractPages()`, `extractPageGuids()`, `findOwnerPage()`
  with known .one binary fixtures
- `src/graph.ts` — mock Graph API responses, test 5000-limit fallback
- `src/index.ts` — CLI integration tests with yargs

### MS-ONESTORE proper parser
Current binary parsing uses heuristics (UTF-8/UTF-16LE text extraction + GUID pattern matching).
A proper implementation would:
- Parse FileNodeListFragment structures per MS-ONE spec
- Resolve ObjectSpaceManifestList → RevisionManifestList → page content
- Extract page titles, bodies, and GUIDs with 100% accuracy
- Handle all page types including ink, images, embedded files
Reference: [MS-ONESTORE] specification from Microsoft Open Specifications.

### Multi-cloud support
Support sovereign clouds beyond `login.microsoftonline.com/common`:
- US Government: `login.microsoftonline.us`
- China (21Vianet): `login.chinacloudapi.cn`
- Graph endpoints also differ per cloud
Add `ONENOTE_CLOUD` env var or `--cloud` flag.

### react-ink TUI
Replace raw ANSI escape codes with react-ink for:
- Dynamic sync progress with multiple concurrent downloads
- Interactive search results (arrow keys to navigate, Enter to open)
- Spinner animations during API calls
- Responsive layout that adapts to terminal width
