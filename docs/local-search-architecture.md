# Local Page-Level Search Architecture

## Problem

The Microsoft Graph OneNote API has a hard limit: when a OneDrive for Business document library contains more than 5,000 OneNote items, the API returns error 10008 and blocks all section/page listing endpoints. This affects accounts with many notebooks.

The Graph Search API (`/search/query`) only returns section-level (.one file) results — it cannot identify individual pages within a section.

## Solution: Local Cache + Binary Text Extraction

### How it works

1. **Sync** (`onenote sync`): Downloads all `.one` section files from OneDrive to a local cache directory (`.onenote/cache/`)
2. **Extract**: Parses the MS-ONESTORE binary format to extract readable text blocks, then segments them into pages based on binary gaps
3. **Search** (`onenote search <query>`): Searches the local cache for matching text at the page level

### Cache Structure

```
.onenote/
  cache/
    {Notebook Name}/
      {Section Name}.json    # Extracted pages with text
```

Each `.json` file contains:
```json
{
  "section": "Section Name",
  "notebook": "Notebook Name",
  "webUrl": "https://...Doc.aspx?sourcedoc={GUID}&...",
  "pages": [
    { "title": "Page Title", "body": "Full text content..." },
    ...
  ],
  "cachedAt": "2025-01-01T00:00:00.000Z"
}
```

### Binary-Based Position Search

For accurate page attribution, the cache stores the original `.one` binary alongside extracted metadata. Search works by:

1. Searching the binary directly for the query string (both UTF-8 and UTF-16LE encodings)
2. For each match position, finding the nearest preceding page GUID anchor
3. Returning that page as the result

This bypasses the imperfect text-block-to-page heuristics and gives accurate page attribution because the binary positions are ground truth — the matched text is physically located near the page anchor it belongs to.

### .one Binary Text Extraction

The MS-ONESTORE format (`.one` files) stores text as UTF-8 and UTF-16LE encoded strings interspersed with binary data. The extraction algorithm:

1. Read the file twice — once as UTF-8 (ASCII content) and once as UTF-16LE (CJK and other Unicode)
2. Find contiguous runs of printable characters
3. Filter runs that don't contain enough "common" characters (ratio threshold)
4. Use the page GUID anchors (see below) to assign each text block to the correct page

### Page GUID Extraction

Page-level URLs require knowing each page's UUID. We extract these from the binary using a pattern observed in MS-ONESTORE files:

```
[UTF-16LE title text] 00 00 [10 00 00 00] [16-byte page GUID]
```

The `10 00 00 00` is a uint32-LE size marker meaning "16 bytes follow", and the next 16 bytes are a UUIDv4 (version=4, variant=8-B). The text immediately preceding (UTF-16LE encoded) is the page title.

For each page GUID found, all text blocks at later offsets up to the next anchor are assigned to that page. This anchor-based grouping correctly maps content (including text in non-Latin scripts) to the right page even when the file structure is fragmented.

### Page-Level URL Format

OneNote Online supports deep linking to specific pages via:

```
{sectionUrl}&wd=target({pageTitle}|{pageGuid}/)
```

Where:
- `{sectionUrl}` is the SharePoint Doc.aspx URL with the section's `sourcedoc` GUID
- `{pageTitle}` is the page title with `)` and `|` characters escaped as `\)` and `\|`
- `{pageGuid}` is the page's UUIDv4
- Trailing `/` is required

The full `wd=` parameter is then URL-encoded with strict encoding (parens encoded as `%28`/`%29`).

### Search Output

Each result shows:
- **Page title** — extracted from the page anchor
- **Section and notebook** — which section/notebook contains the match
- **Context snippet** — text surrounding the match with keyword highlighted in `**bold**`
- **URL** — page-level OneNote Online deep link

### Page-Level URL Resolution (Official URLs)

We use the OneNote Graph API endpoint `GET /me/onenote/sections/0-{guid}/pages` to fetch the official `links.oneNoteWebUrl.href` for each page. This endpoint works even when the 5,000-item document library limit blocks `/me/onenote/pages` and `/me/onenote/sections` listing — because the `0-{guid}` ID prefix targets the section directly via its sourcedoc GUID (extracted from the OneDrive driveItem webUrl).

Official OneNote URLs use the format:
```
{driveRootPath}/{notebook}?wd=target({sectionFile}|{sectionGroupGuid}/{pageTitle}|{pageGuid}/)
```

Unlike the simpler `Doc.aspx?sourcedoc=...&wd=target(...)` format, this URL **bypasses OneNote Online's session caching** and navigates directly to the specified page on first load. Verified working via browser automation testing.

The page navigation GUID (used in `wd=target`) is extracted from the `oneNoteWebUrl` URL itself (the last UUID in the URL), since the API's page `id` field uses a different identifier format (`1-{32hexchars}!{counter}-{sectionGuid}`) that doesn't match the navigation GUID.

### Limitations

- **Page coverage**: Our binary parser detects ~50% of pages in some sections (those whose GUID-title binary pattern matches our heuristic). The other pages exist in the cache but without a precise GUID, so search results for them fall back to section-level URLs.
- **Cache freshness**: Cache is valid for 1 hour by default. Run `onenote sync` to refresh.
- **Binary parsing heuristic**: The UTF-8 text extraction may miss some content or include some binary noise. Page boundary detection is approximate.
- **Large sections**: Sections with thousands of pages (e.g., 5000+ extracted "pages") may have over-segmented results due to the binary gap heuristic.

### Alternative Approaches Investigated

| Approach | Result |
|---|---|
| OneNote API `/me/onenote/pages` | Blocked by 5000 item limit (error 10008) |
| Graph Search API `driveItem` | Section-level only, no page granularity |
| Graph Search API `listItem` | Section-level only (pages not individually indexed) |
| OneDrive HTML conversion (`?format=html`) | Not supported for .one files (406) |
| SharePoint REST search | Requires separate OAuth scope, returns section-level |
| Site-specific OneNote API | Same 5000 limit applies |
| Beta Graph API endpoints | Same limitations |

### Future Improvements

- Implement proper MS-ONESTORE parser for accurate page extraction with page GUIDs
- Use page GUIDs to construct page-level deep links: `Doc.aspx?sourcedoc={guid}&wd=target(section|/pageTitle)`
- Incremental sync (only download changed sections based on `lastModifiedDateTime`)
- SQLite or full-text search index for faster queries on large caches
