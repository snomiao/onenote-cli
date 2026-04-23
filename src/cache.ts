import prettyBytes from "pretty-bytes";
import { getAccessToken, listAccounts, setCurrentAccount } from "./auth";
import { listNotebooks } from "./graph";
import { readFile, writeFile, mkdir, readdir, stat } from "node:fs/promises";
import { join } from "node:path";
import { Database } from "bun:sqlite";

import { homedir } from "node:os";
import { dirname } from "node:path";

// Cache directory: use package root (../.onenote/cache relative to src/) so that
// the cache lives alongside .env.local. Fall back to ~/.onenote-cli/cache if the
// package root is not writable (e.g. installed via npm).
const PKG_ROOT = dirname(import.meta.dir);
const CACHE_DIR = process.env.ONENOTE_CACHE_DIR
  || join(PKG_ROOT, ".onenote", "cache");

interface CachedPage {
  title: string;
  body: string;
  section: string;
  notebook: string;
  webUrl: string; // OneNote Online URL for this page (section + page GUID)
  pageGuid?: string;
  tagLines?: Array<{ tag: string; text: string }>; // text content of tagged elements
}

interface CacheIndex {
  updatedAt: string;
  notebooks: {
    id: string;
    displayName: string;
    sections: {
      driveItemId: string;
      displayName: string;
      webUrl: string;
      drivePath: string;
      cachedAt: string;
    }[];
  }[];
}

function getNotebookDrivePath(notebook: any): string | null {
  const webUrl = notebook.links?.oneNoteWebUrl?.href;
  if (!webUrl) return null;
  const match = decodeURIComponent(new URL(webUrl).pathname).match(/Documents\/(.+)/);
  return match?.[1] ?? null;
}

async function graphFetchRaw(path: string): Promise<Response> {
  const token = await getAccessToken();
  const url = path.startsWith("http")
    ? path
    : `https://graph.microsoft.com/v1.0${path}`;
  // Retry on 429 (rate limit) with exponential backoff, respecting Retry-After header
  for (let attempt = 0; attempt < 8; attempt++) {
    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (res.status !== 429) return res;
    const retryAfter = parseInt(res.headers.get("retry-after") ?? "0", 10);
    const delayMs = retryAfter > 0 ? retryAfter * 1000 : Math.min(5000 * 2 ** attempt, 120000);
    await new Promise((r) => setTimeout(r, delayMs));
  }
  return fetch(url, { headers: { Authorization: `Bearer ${token}` } });
}

function isReadableChar(code: number): boolean {
  return (
    (code >= 0x20 && code <= 0x7e) || // ASCII printable
    code === 0x0a || code === 0x0d || code === 0x09 || // whitespace
    (code >= 0xa0 && code <= 0x024f) || // Latin Extended
    (code >= 0x0370 && code <= 0x058f) || // Greek, Cyrillic, Armenian
    (code >= 0x0600 && code <= 0x06ff) || // Arabic
    (code >= 0x0900 && code <= 0x097f) || // Devanagari
    (code >= 0x0e00 && code <= 0x0e7f) || // Thai
    (code >= 0x1100 && code <= 0x11ff) || // Hangul Jamo
    (code >= 0x2000 && code <= 0x206f) || // General Punctuation
    (code >= 0x2100 && code <= 0x214f) || // Letterlike Symbols
    (code >= 0x2190 && code <= 0x21ff) || // Arrows
    (code >= 0x2200 && code <= 0x22ff) || // Mathematical Operators
    (code >= 0x2500 && code <= 0x257f) || // Box Drawing
    (code >= 0x3000 && code <= 0x303f) || // CJK Symbols and Punctuation
    (code >= 0x3040 && code <= 0x309f) || // Hiragana
    (code >= 0x30a0 && code <= 0x30ff) || // Katakana
    (code >= 0x3100 && code <= 0x312f) || // Bopomofo
    (code >= 0x3400 && code <= 0x4dbf) || // CJK Extension A
    (code >= 0x4e00 && code <= 0x9fff) || // CJK Unified Ideographs
    (code >= 0xac00 && code <= 0xd7af) || // Hangul Syllables
    (code >= 0xf900 && code <= 0xfaff) || // CJK Compatibility Ideographs
    (code >= 0xfe30 && code <= 0xfe4f) || // CJK Compatibility Forms
    (code >= 0xff00 && code <= 0xffef)    // Fullwidth Forms
  );
}

function extractTextBlocks(buf: Buffer): { offset: number; text: string }[] {
  const blocks: { offset: number; text: string }[] = [];

  // Extract UTF-8 text blocks
  const utf8 = buf.toString("utf-8");
  let start = -1;
  let chars = "";
  for (let i = 0; i < utf8.length; i++) {
    if (isReadableChar(utf8.charCodeAt(i))) {
      if (start < 0) start = i;
      chars += utf8[i];
    } else {
      if (chars.trim().length >= 6) blocks.push({ offset: start, text: chars.trim() });
      chars = "";
      start = -1;
    }
  }
  if (chars.trim().length >= 6) blocks.push({ offset: start, text: chars.trim() });

  // Extract UTF-16LE text blocks at both even and odd byte alignments
  // (OneNote may have UTF-16LE strings starting at either alignment)
  for (const startOffset of [0, 1]) {
    start = -1;
    chars = "";
    for (let i = startOffset; i < buf.length - 1; i += 2) {
      const code = buf[i] | (buf[i + 1] << 8);
      if (isReadableChar(code)) {
        if (start < 0) start = i;
        chars += String.fromCharCode(code);
      } else {
        if (chars.trim().length >= 4) blocks.push({ offset: start, text: chars.trim() });
        chars = "";
        start = -1;
      }
    }
    if (chars.trim().length >= 4) blocks.push({ offset: start, text: chars.trim() });
  }

  // Sort by offset
  blocks.sort((a, b) => a.offset - b.offset);

  // Filter noise: require blocks to have a reasonable ratio of common characters
  return blocks.filter((b) => {
    const common = b.text.replace(
      /[^a-zA-Z0-9\u3040-\u30ff\u4e00-\u9fff\u3000-\u303f\uff00-\uffef\s.,;:!?@#\-_()[\]{}'"/\\]/g,
      ""
    );
    if (common.length / b.text.length <= 0.6 || b.text.length < 4) return false;

    // Detect misaligned UTF-16LE reading of ASCII: characters where low byte = 0x00
    // and code is in the "shifted ASCII" range (0x2000-0x7E00 typically)
    let shiftedAsciiCount = 0;
    let cjkCount = 0;
    for (const ch of b.text) {
      const code = ch.charCodeAt(0);
      if ((code & 0xff) === 0 && code >= 0x2000 && code <= 0x7f00) {
        shiftedAsciiCount++;
      }
      if (code >= 0x4e00 && code <= 0x9fff) cjkCount++;
    }
    // If most characters look like shifted-ASCII garbage, reject
    if (shiftedAsciiCount > 3 && shiftedAsciiCount / b.text.length > 0.5) return false;
    return true;
  });
}

function groupIntoPages(blocks: { offset: number; text: string }[]): { title: string; body: string }[] {
  const pages: { title: string; body: string }[] = [];
  let group: typeof blocks = [];

  for (const block of blocks) {
    if (group.length > 0) {
      const prevEnd = group[group.length - 1].offset + group[group.length - 1].text.length * 2;
      const gap = block.offset - prevEnd;
      if (gap > 500) {
        const body = group.map((b) => b.text).join("\n");
        if (body.length > 10) {
          const lines = body.split("\n").filter((l) => l.trim().length > 0);
          pages.push({ title: lines[0]?.slice(0, 200) || "(untitled)", body });
        }
        group = [];
      }
    }
    group.push(block);
  }
  if (group.length > 0) {
    const body = group.map((b) => b.text).join("\n");
    if (body.length > 10) {
      const lines = body.split("\n").filter((l) => l.trim().length > 0);
      pages.push({ title: lines[0]?.slice(0, 200) || "(untitled)", body });
    }
  }

  return pages;
}

function bufToGuid(b: Buffer, off: number): string {
  return [
    b.readUInt32LE(off).toString(16).padStart(8, "0"),
    b.readUInt16LE(off + 4).toString(16).padStart(4, "0"),
    b.readUInt16LE(off + 6).toString(16).padStart(4, "0"),
    b.slice(off + 8, off + 10).toString("hex"),
    b.slice(off + 10, off + 16).toString("hex"),
  ].join("-");
}

function guidToBuf(guid: string): Buffer {
  const p = guid.replace(/-/g, "");
  const b = Buffer.alloc(16);
  b.writeUInt32LE(parseInt(p.slice(0, 8), 16), 0);
  b.writeUInt16LE(parseInt(p.slice(8, 12), 16), 4);
  b.writeUInt16LE(parseInt(p.slice(12, 16), 16), 6);
  Buffer.from(p.slice(16, 32), "hex").copy(b, 8);
  return b;
}

/**
 * Expand first-occurrence anchors to include ALL binary occurrences of each known GUID.
 * The .one format stores page objects non-contiguously; using only first occurrences as
 * boundaries causes NTP/AIS pairs from later chunks of page A to be attributed to page B.
 */
function expandAnchorsToAllOccurrences(
  buf: Buffer,
  anchors: { offset: number; guid: string }[]
): { offset: number; guid: string }[] {
  const guidHexToGuid = new Map<string, string>();
  for (const { guid } of anchors) {
    guidHexToGuid.set(guidToBuf(guid).toString("hex"), guid);
  }
  const SIZE_MARKER = Buffer.from([0x10, 0x00, 0x00, 0x00]);
  const result: { offset: number; guid: string }[] = [];
  let pos = 0;
  while (pos < buf.length - 20) {
    const markerPos = buf.indexOf(SIZE_MARKER, pos);
    if (markerPos < 0) break;
    const guidHex = buf.slice(markerPos + 4, markerPos + 20).toString("hex");
    const guid = guidHexToGuid.get(guidHex);
    if (guid) result.push({ offset: markerPos, guid });
    pos = markerPos + 1;
  }
  return result;
}

/**
 * Extract (pageGuid, title, offset) tuples from .one binary.
 * Pattern: [UTF-16LE title] 00 00 [10 00 00 00] [16-byte GUID]
 */
export function extractPageGuids(
  buf: Buffer
): { guid: string; title: string; offset: number }[] {
  const results: { guid: string; title: string; offset: number }[] = [];
  const seen = new Set<string>();

  for (let i = 0; i < buf.length - 20; i++) {
    // Size marker 10 00 00 00 (uint32 LE = 16)
    if (buf[i] !== 0x10 || buf[i + 1] !== 0 || buf[i + 2] !== 0 || buf[i + 3] !== 0) continue;

    // Check GUID validity (UUIDv4: version=4, variant=8-B)
    const v = (buf[i + 4 + 7] >> 4) & 0xf;
    const vr = (buf[i + 4 + 8] >> 4) & 0xf;
    if (v !== 4 || vr < 8 || vr > 0xb) continue;

    const guid = bufToGuid(buf, i + 4);
    if (seen.has(guid)) continue;

    // Walk backwards from i to find UTF-16LE title
    let j = i - 2;
    if (j >= 0 && buf[j] === 0 && buf[j + 1] === 0) j -= 2; // skip terminator
    let chars = "";
    while (j >= 0 && chars.length < 200) {
      const code = buf[j] | (buf[j + 1] << 8);
      if ((code >= 0x20 && code <= 0x7e) || (code >= 0xa0 && code <= 0xffef)) {
        chars = String.fromCharCode(code) + chars;
        j -= 2;
      } else {
        break;
      }
    }

    if (chars.length >= 3 && chars.length < 200) {
      // Filter out garbage titles: must contain a meaningful ratio of "real" characters
      // (ASCII letters/digits, common CJK ideographs, hiragana/katakana, punctuation)
      const meaningful = chars.replace(
        /[^a-zA-Z0-9\u3040-\u30ff\u4e00-\u9fff\u3000-\u303f\uff00-\uffef\s.,;:!?@#\-_()[\]{}'"/\\]/g,
        ""
      );
      if (meaningful.length / chars.length < 0.7) continue;

      // Reject titles where most CJK chars are likely shifted-ASCII garbage
      let shiftedCount = 0;
      for (const ch of chars) {
        const code = ch.charCodeAt(0);
        if ((code & 0xff) === 0 && code >= 0x2000) shiftedCount++;
      }
      if (shiftedCount > 2 && shiftedCount / chars.length > 0.3) continue;

      // Reject embedded object titles like ".jpg", ".png", ".pdf" that are
      // attachment GUIDs, not page GUIDs
      const trimmed = chars.trim();
      if (/^\.[a-z0-9]{2,5}$/i.test(trimmed)) continue;

      seen.add(guid);
      results.push({ guid, title: trimmed, offset: i });
    }
  }

  return results;
}

export function extractPages(
  buf: Buffer
): { title: string; body: string; pageGuid?: string }[] {
  const blocks = extractTextBlocks(buf);
  const guidEntries = extractPageGuids(buf);

  if (guidEntries.length === 0) {
    return groupIntoPages(blocks).map((p) => ({ ...p, pageGuid: undefined }));
  }

  // Sort guid entries by offset and dedupe to get unique pages with their FIRST offset
  guidEntries.sort((a, b) => a.offset - b.offset);
  const firstOffsetByGuid = new Map<string, { title: string; offset: number }>();
  const titleByGuid = new Map<string, string>();
  for (const e of guidEntries) {
    if (!firstOffsetByGuid.has(e.guid)) {
      firstOffsetByGuid.set(e.guid, { title: e.title.trim(), offset: e.offset });
      titleByGuid.set(e.guid, e.title.trim());
    }
  }

  // Build sorted list of (offset, guid) anchors using ALL occurrences
  const anchors = guidEntries
    .map((e) => ({ offset: e.offset, guid: e.guid }))
    .sort((a, b) => a.offset - b.offset);

  // Build a title -> guid map for boundary detection
  const titleToGuidMap = new Map<string, string>();
  for (const [guid, title] of titleByGuid) {
    if (title.length >= 4) titleToGuidMap.set(title, guid);
  }
  // Sort known titles by length desc for greedy match
  const knownTitlesSorted = [...titleToGuidMap.keys()].sort((a, b) => b.length - a.length);

  // For each text block:
  // - If it matches a known page title, switch to that page's GUID
  // - Otherwise, append to the current page's body
  const bodiesByGuid = new Map<string, string[]>();
  let anchorIdx = 0;
  let currentGuid: string | undefined;
  for (const block of blocks) {
    // Check if this block IS a known page title (boundary)
    let titleMatch: string | undefined;
    for (const t of knownTitlesSorted) {
      if (block.text === t || block.text.startsWith(t)) {
        titleMatch = t;
        break;
      }
    }
    if (titleMatch) {
      currentGuid = titleToGuidMap.get(titleMatch);
      // Skip pushing the title text into the body to avoid noise
      continue;
    }

    // Otherwise advance anchor by offset
    while (anchorIdx < anchors.length && anchors[anchorIdx].offset <= block.offset) {
      currentGuid = anchors[anchorIdx].guid;
      anchorIdx++;
    }
    if (!currentGuid) continue;
    const arr = bodiesByGuid.get(currentGuid) || [];
    arr.push(block.text);
    bodiesByGuid.set(currentGuid, arr);
  }

  // Build final pages
  const pages: { title: string; body: string; pageGuid?: string }[] = [];
  for (const [guid, info] of firstOffsetByGuid) {
    const body = (bodiesByGuid.get(guid) || []).join("\n");
    if (body.length < 5 && info.title.length < 3) continue;
    pages.push({ title: info.title || "(untitled)", body: body || info.title, pageGuid: guid });
  }
  return pages;
}

async function ensureDir(dir: string) {
  await mkdir(dir, { recursive: true });
}

async function downloadSection(
  drivePath: string
): Promise<Buffer | null> {
  const encoded = drivePath
    .split("/")
    .map((s) => encodeURIComponent(s))
    .join("/");
  const res = await graphFetchRaw(`/me/drive/root:/${encoded}:/content`);
  if (!res.ok) return null;
  return Buffer.from(await res.arrayBuffer());
}

async function getSectionWebUrl(drivePath: string): Promise<string> {
  const encoded = drivePath
    .split("/")
    .map((s) => encodeURIComponent(s))
    .join("/");
  try {
    const res = await graphFetchRaw(
      `/me/drive/root:/${encoded}?$select=webUrl`
    );
    if (res.ok) {
      const item = (await res.json()) as any;
      return item.webUrl?.split("&mobileredirect")[0] ?? "";
    }
  } catch {}
  return "";
}

/**
 * Get OneNote pages for a section via Graph API.
 * Uses the `0-{guid}` ID prefix which works even when the 5,000-item limit
 * blocks listing endpoints.
 */
async function getOneNotePagesForSection(
  sectionGuid: string
): Promise<{ id: string; title: string; webUrl: string }[]> {
  const results: { id: string; title: string; webUrl: string }[] = [];
  let url: string | null = `/me/onenote/sections/0-${sectionGuid}/pages?$select=id,title,links&$top=100`;
  try {
    while (url) {
      const res = await graphFetchRaw(url);
      if (!res.ok) break;
      const data = (await res.json()) as any;
      for (const p of data.value ?? []) {
        results.push({ id: p.id, title: p.title ?? "", webUrl: p.links?.oneNoteWebUrl?.href ?? "" });
      }
      url = data["@odata.nextLink"] ?? null;
    }
  } catch {}
  return results;
}

/**
 * Extract the page navigation GUID from the OneNote oneNoteWebUrl.
 * The webUrl contains `wd=target(...|{lastGuid}/)` where lastGuid is the page GUID
 * used for navigation (matches what we extract from the binary).
 */
function pageGuidFromWebUrl(webUrl: string): string | null {
  if (!webUrl) return null;
  // Find the LAST GUID in the URL (the page-level one, after the section group)
  // Pattern: ...{guid1}/{title}|{pageGuid}/)
  const decoded = decodeURIComponent(webUrl);
  const matches = decoded.match(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/gi);
  if (!matches || matches.length === 0) return null;
  return matches[matches.length - 1].toLowerCase();
}

async function listSectionFiles(
  notebook: any
): Promise<{ name: string; drivePath: string; lastModified: string; size: number; apiOnly?: boolean; sectionId?: string }[]> {
  const nbPath = getNotebookDrivePath(notebook);
  if (nbPath) {
    const encoded = nbPath.split("/").map((s) => encodeURIComponent(s)).join("/");
    try {
      const res = await graphFetchRaw(
        `/me/drive/root:/${encoded}:/children?$select=name,id,file,size,lastModifiedDateTime&$top=200`
      );
      if (res.ok) {
        const data = (await res.json()) as any;
        const files = (data.value ?? [])
          .filter((f: any) => f.name?.endsWith(".one"))
          .map((f: any) => ({
            name: f.name.replace(/\.one$/, ""),
            drivePath: `${nbPath}/${f.name}`,
            lastModified: f.lastModifiedDateTime ?? "",
            size: f.size ?? 0,
          }));
        if (files.length > 0) return files;
      }
    } catch {}
  }

  // Fallback: personal OneDrive / unresolvable path → use OneNote API to list sections.
  // These sections are synced HTML-only (no .one binary download).
  try {
    const res = await graphFetchRaw(
      `/me/onenote/notebooks/${notebook.id}/sections?$select=id,displayName,lastModifiedDateTime&$top=100`
    );
    if (!res.ok) return [];
    const data = (await res.json()) as any;
    return (data.value ?? []).map((s: any) => ({
      name: s.displayName ?? "unnamed",
      drivePath: "",
      lastModified: s.lastModifiedDateTime ?? "",
      size: 0,
      apiOnly: true,
      sectionId: s.id,
    }));
  } catch {
    return [];
  }
}

export const SEARCH_DB_PATH = join(PKG_ROOT, ".onenote", "search.db");

// Maps #alias → OneNote data-tag value(s). Multiple values = OR match.
export const TAG_ALIASES: Record<string, string[]> = {
  star:        ["star"],
  question:    ["question"],
  important:   ["important"],
  critical:    ["critical"],
  definition:  ["definition"],
  idea:        ["idea"],
  contact:     ["contact"],
  address:     ["address"],
  phone:       ["phone-number"],
  website:     ["web-site-to-visit"],
  password:    ["password"],
  remember:    ["remember-for-later"],
  book:        ["book-to-read"],
  music:       ["music-to-listen-to"],
  movie:       ["movie-to-see"],
  highlight:   ["highlight"],
  meeting:     ["schedule-meeting"],
  email:       ["send-in-email"],
  callback:    ["call-back"],
  discuss:     ["discuss-with-person-a", "discuss-with-person-b", "discuss-with-manager"],
  priority1:   ["to-do-priority-1"],
  priority2:   ["to-do-priority-2"],
  client:      ["client-request"],
};

function openSearchDb(): Database {
  const db = new Database(SEARCH_DB_PATH, { create: true });
  db.run("PRAGMA journal_mode=WAL");
  db.run(`
    CREATE TABLE IF NOT EXISTS pages (
      id INTEGER PRIMARY KEY,
      section TEXT NOT NULL,
      notebook TEXT NOT NULL,
      title TEXT,
      body TEXT,
      web_url TEXT,
      page_guid TEXT,
      has_todo INTEGER DEFAULT NULL,
      has_done INTEGER DEFAULT NULL,
      tags TEXT DEFAULT NULL,
      tag_lines TEXT DEFAULT NULL,
      account TEXT DEFAULT NULL
    )
  `);
  // Migrate existing DBs that lack these columns
  try { db.run("ALTER TABLE pages ADD COLUMN has_todo INTEGER DEFAULT NULL"); } catch {}
  try { db.run("ALTER TABLE pages ADD COLUMN has_done INTEGER DEFAULT NULL"); } catch {}
  try { db.run("ALTER TABLE pages ADD COLUMN tags TEXT DEFAULT NULL"); } catch {}
  try { db.run("ALTER TABLE pages ADD COLUMN tag_lines TEXT DEFAULT NULL"); } catch {}
  try { db.run("ALTER TABLE pages ADD COLUMN account TEXT DEFAULT NULL"); } catch {}
  db.run("CREATE INDEX IF NOT EXISTS pages_section ON pages(section, notebook)");
  db.run(`
    CREATE VIRTUAL TABLE IF NOT EXISTS pages_fts USING fts5(
      title, body,
      content="",
      tokenize="unicode61"
    )
  `);
  return db;
}

/** Parse tag filters out of a query string. Returns text-only FTS query and extracted tags.
 *  Recognizes: #todo  tag:todo  #done  tag:done
 */
/**
 * Build an FTS5 MATCH parameter from a user query string.
 * - Explicit boolean operators (AND/OR/NOT, case-insensitive) are normalized to uppercase.
 * - Otherwise each space-separated token is quoted for exact-term matching (implicit AND).
 * - User-supplied "quoted phrases" are preserved as-is.
 */
export function buildFtsParam(query: string): string {
  if (/\b(AND|OR|NOT)\b/i.test(query)) {
    // Normalize operators to uppercase; preserve user-quoted phrases
    return query.replace(/\b(AND|OR|NOT)\b/gi, (m) => m.toUpperCase());
  }
  // Split preserving quoted phrases, then quote individual bare terms
  const tokens = query.match(/(?:"[^"]*"|[^\s]+)/g) ?? [];
  return tokens
    .map((t) => (t.startsWith('"') ? t : `"${t.replace(/"/g, '""')}"`))
    .join(" ");
}

export function parseTagsFromQuery(query: string): {
  ftsQuery: string;
  hasTodo: boolean;
  hasDone: boolean;
  hasCheckbox: boolean;
  tagFilters: string[]; // OneNote data-tag values to filter by (OR within each alias group, AND across groups)
} {
  let hasTodo = false;
  let hasDone = false;
  let hasCheckbox = false;
  const tagFilters: string[] = [];

  let ftsQuery = query
    .replace(/#checkbox\b/gi, () => { hasCheckbox = true; return ""; })
    .replace(/\btag:checkbox\b/gi, () => { hasCheckbox = true; return ""; })
    .replace(/#todo\b/gi, () => { hasTodo = true; return ""; })
    .replace(/\btag:todo\b/gi, () => { hasTodo = true; return ""; })
    .replace(/#done\b/gi, () => { hasDone = true; return ""; })
    .replace(/\btag:done\b/gi, () => { hasDone = true; return ""; });

  // Replace all known #alias and tag:alias patterns
  for (const [alias, tagValues] of Object.entries(TAG_ALIASES)) {
    const escaped = alias.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    ftsQuery = ftsQuery
      .replace(new RegExp(`#${escaped}\\b`, "gi"), () => { tagFilters.push(...tagValues); return ""; })
      .replace(new RegExp(`\\btag:${escaped}\\b`, "gi"), () => { tagFilters.push(...tagValues); return ""; });
  }

  ftsQuery = ftsQuery.replace(/\s+/g, " ").trim();
  return { ftsQuery, hasTodo, hasDone, hasCheckbox, tagFilters };
}

const BINARY_BODY_LIMIT = 50_000; // chars extracted from .one binary per anchor range

function asciiRatio(s: string): number {
  if (!s.length) return 0;
  const printable = s.split("").filter((c) => c >= "\x20" && c <= "\x7E").length;
  return printable / s.length;
}

function cleanBodyForIndex(body: string): string {
  // Strip binary garbage lines before FTS indexing to keep the posting list small.
  // A line is kept if it has reasonable ASCII ratio OR contains meaningful CJK text.
  const lines = body
    .replace(/[\x00-\x09\x0B-\x1F\x7F]/g, " ")
    .split(/\n+/)
    .map((l) => l.replace(/\s+/g, " ").trim())
    .filter((l) => {
      if (l.length < 2) return false;
      const ratio = asciiRatio(l);
      // Keep lines with decent ASCII ratio or meaningful non-ASCII (CJK etc)
      if (ratio >= 0.3) return true;
      // Count CJK/Hiragana/Katakana characters (U+3000-U+9FFF, U+AC00-U+D7AF)
      const cjk = (l.match(/[\u3000-\u9FFF\uAC00-\uD7AF]/g) ?? []).length;
      return cjk / l.length >= 0.2;
    });
  return lines.join("\n");
}

// MS-ONESTORE property markers for note tags
const NOTE_TAGS_PROP     = Buffer.from([0x89, 0x34, 0x00, 0x40]); // 0x40003489: NoteTags property
const ACTION_ITEM_STATUS = Buffer.from([0x70, 0x34, 0x00, 0x10]); // 0x10003470: ActionItemStatus

/**
 * Extract page GUIDs that have at least one uncompleted action-item tag (To Do, Remember for later, etc.)
 * directly from the .one binary without any API calls.
 *
 * Detection: find NoteTags property (89 34 00 40) followed within 300 bytes by
 * ActionItemStatus (70 34 00 10) where the 2-byte status value has bit0=0 (not completed).
 * Map each match to the nearest preceding page anchor.
 *
 * Note: this catches ALL uncompleted action tags, not only pure To-Do checkboxes.
 * Shape-based discrimination (shape=1 = checkbox) requires full OID resolution from the
 * PropertySet graph and is not implemented here.
 */
export function extractTodoPageGuids(
  buf: Buffer,
  anchors: { offset: number; guid: string }[]
): { todo: Set<string>; done: Set<string> } {
  const sorted = expandAnchorsToAllOccurrences(buf, anchors).sort((a, b) => a.offset - b.offset);
  const todo = new Set<string>();
  const done = new Set<string>();

  let pos = 0;
  while (true) {
    const idx = buf.indexOf(NOTE_TAGS_PROP, pos);
    if (idx < 0) break;

    const window = buf.slice(idx, idx + 300);
    const aisRel = window.indexOf(ACTION_ITEM_STATUS);
    if (aisRel >= 0) {
      const aisAbs = idx + aisRel;
      // ActionItemStatus is a uint16 at AIS_prid_pos + 12.
      // The 8 bytes at +4..+11 are the FILETIME data for the preceding property (0x1400346f).
      // Empirically verified: status=0=unchecked, status=1=completed.
      if (aisAbs + 14 <= buf.length) {
        const status = buf.readUInt16LE(aisAbs + 12);
        if (status === 0 || status === 1) {
          // find nearest anchor (preceding, or following if very close)
          // MS-ONESTORE stores page content before the page GUID marker, so a NTP that falls
          // just before the next anchor likely belongs to that next page.
          let preceding: { offset: number; guid: string } | null = null;
          let following: { offset: number; guid: string } | null = null;
          for (const a of sorted) {
            if (a.offset <= idx) preceding = a;
            else { following = a; break; }
          }
          const LOOKAHEAD = 1000; // bytes: prefer following anchor if it's this close
          const nearest =
            following && following.offset - idx < LOOKAHEAD ? following.guid : preceding?.guid ?? null;
          if (nearest) {
            if (status === 0) todo.add(nearest);
            else done.add(nearest);
          }
        }
      }
    }
    pos = idx + 1;
  }
  return { todo, done };
}

export type PageTagInfo = {
  tags: string[]; // unique sorted tag values
  lines: Array<{ tag: string; text: string }>; // text of each tagged element
};

function stripHtml(s: string): string {
  return s
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(parseInt(n, 10)))
    .replace(/\s+/g, " ")
    .trim();
}

/** Extract text of each tagged element from OneNote HTML */
export function extractTagLines(html: string): Array<{ tag: string; text: string }> {
  const lines: Array<{ tag: string; text: string }> = [];
  // Matches elements with data-tag attribute and captures content up to closing tag.
  // Uses non-greedy match; OneNote generally doesn't nest tagged elements.
  const re = /<(\w+)([^>]*?)\s+data-tag="([^"]+)"([^>]*)>([\s\S]*?)<\/\1>/gi;
  for (const m of html.matchAll(re)) {
    const tagValues = m[3].split(/[\s,]+/).filter(Boolean);
    const text = stripHtml(m[5]).slice(0, 200);
    if (text.length > 0) {
      for (const t of tagValues) lines.push({ tag: t, text });
    }
  }
  return lines;
}

/**
 * Fetch page HTML content from OneNote API and collect all data-tag values + text per page.
 */
async function fetchTagsFromHtml(
  pages: Array<{ apiId: string; guid: string }>,
  onProgress?: (msg: string) => void
): Promise<Map<string, PageTagInfo>> {
  const result = new Map<string, PageTagInfo>();
  for (const page of pages) {
    try {
      const res = await graphFetchRaw(`/me/onenote/pages/${page.apiId}/content`);
      if (!res.ok) continue;
      const html = await res.text();
      const found = new Set<string>();
      for (const m of html.matchAll(/data-tag="([^"]+)"/g)) {
        for (const tag of m[1].split(/[\s,]+/)) {
          if (tag) found.add(tag);
        }
      }
      const lines = extractTagLines(html);
      if (found.size > 0) {
        const sorted = [...found].sort();
        result.set(page.guid, { tags: sorted, lines });
        onProgress?.(`    [tags:${sorted.join(",")}] ${page.guid}`);
      }
    } catch {}
  }
  return result;
}

function extractBinaryPageText(buf: Buffer, startOffset: number, endOffset: number): string {
  const segment = buf.slice(startOffset, Math.min(endOffset, startOffset + BINARY_BODY_LIMIT * 3));
  const clean = (s: string) =>
    s.replace(/[\x00-\x1F\x7F\uFFFD]/g, " ").replace(/\s+/g, " ").trim();
  const utf8 = clean(segment.toString("utf-8"));
  const utf16 = clean(segment.toString("utf16le"));
  const best = asciiRatio(utf8) >= asciiRatio(utf16) ? utf8 : utf16;
  return best.slice(0, BINARY_BODY_LIMIT);
}

function upsertSectionToIndex(
  db: Database,
  section: string,
  notebook: string,
  pages: Array<{ title: string; body?: string; pageGuid?: string; officialUrl?: string }>,
  webUrl: string,
  binBuf?: Buffer,
  anchors?: Array<{ offset: number; guid: string; title: string }>,
  officialPages?: Array<{ guid: string | null; title: string; webUrl?: string; apiId?: string }>,
  htmlTags?: Map<string, PageTagInfo>,
  accountEmail?: string
): void {
  // Build lookup maps
  const officialUrlByGuid = new Map<string, { url: string; title: string }>();
  for (const op of officialPages ?? []) {
    if (op.guid && op.webUrl) officialUrlByGuid.set(op.guid, { url: op.webUrl, title: op.title });
  }

  // Binary detection as fallback for pages not covered by HTML fetch
  const binaryGuids = binBuf && anchors?.length
    ? extractTodoPageGuids(binBuf, anchors)
    : { todo: new Set<string>(), done: new Set<string>() };

  // Pages checked via HTML: use HTML result (accurate). Others: use binary (fallback).
  const officialGuidSet = new Set(officialUrlByGuid.keys());
  const getHasTodo = (guid: string): 0 | 1 => {
    if (htmlTags && officialGuidSet.has(guid)) return (htmlTags.get(guid)?.tags ?? []).includes("to-do") ? 1 : 0;
    return binaryGuids.todo.has(guid) ? 1 : 0;
  };
  const getHasDone = (guid: string): 0 | 1 => {
    if (htmlTags && officialGuidSet.has(guid)) return (htmlTags.get(guid)?.tags ?? []).includes("to-do:completed") ? 1 : 0;
    return binaryGuids.done.has(guid) ? 1 : 0;
  };
  // tags column: pipe-delimited tag list, e.g. "|star|to-do|" for LIKE '%|star|%' matching
  const getTagsStr = (guid: string): string | null => {
    const t = htmlTags?.get(guid);
    if (!t || t.tags.length === 0) return null;
    return `|${t.tags.join("|")}|`;
  };
  // tag_lines column: JSON-encoded array of {tag, text} for display
  const getTagLinesStr = (guid: string): string | null => {
    const t = htmlTags?.get(guid);
    if (!t || t.lines.length === 0) return null;
    return JSON.stringify(t.lines);
  };

  db.transaction(() => {
    // Remove old FTS entries first (content table requires explicit delete)
    db.run(
      `INSERT INTO pages_fts(pages_fts, rowid, title, body)
       SELECT 'delete', id, title, body FROM pages WHERE section=? AND notebook=?`,
      [section, notebook]
    );
    db.run("DELETE FROM pages WHERE section=? AND notebook=?", [section, notebook]);

    const insert = db.prepare(
      "INSERT INTO pages(section, notebook, title, body, web_url, page_guid, has_todo, has_done, tags, tag_lines, account) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
    );

    if (binBuf && anchors && anchors.length > 0) {
      const sorted = [...anchors].sort((a, b) => a.offset - b.offset);
      const jsonBodyByGuid = new Map(pages.map((p) => [p.pageGuid, p.body]));
      for (let i = 0; i < sorted.length; i++) {
        const anchor = sorted[i]!;
        const nextOffset = sorted[i + 1]?.offset ?? binBuf.length;
        const binBody = extractBinaryPageText(binBuf, anchor.offset, nextOffset);
        const jsonBody = jsonBodyByGuid.get(anchor.guid) ?? "";
        const body = cleanBodyForIndex((jsonBody + " " + binBody).slice(0, BINARY_BODY_LIMIT));
        const official = officialUrlByGuid.get(anchor.guid);
        const url = official?.url ?? buildPageUrl(webUrl, anchor.title, anchor.guid);
        insert.run(section, notebook, official?.title ?? anchor.title, body, url, anchor.guid, getHasTodo(anchor.guid), getHasDone(anchor.guid), getTagsStr(anchor.guid), getTagLinesStr(anchor.guid), accountEmail ?? null);
      }
    } else {
      for (const p of pages) {
        const url = p.officialUrl ?? buildPageUrl(webUrl, p.title, p.pageGuid);
        insert.run(section, notebook, p.title ?? "", cleanBodyForIndex(p.body ?? ""), url, p.pageGuid ?? "", null, null, null, null, accountEmail ?? null);
      }
    }

    // Rebuild FTS for this section from the pages table
    db.run(
      `INSERT INTO pages_fts(rowid, title, body)
       SELECT id, title, body FROM pages WHERE section=? AND notebook=?`,
      [section, notebook]
    );
  })();
}

export type SyncEmit = (ev: import("./sync-ui").SyncEvent) => void;

/** Sanitize email for filesystem path usage (replace @ and any other unsafe chars). */
function accountDirName(email: string): string {
  return email.replace(/[^a-zA-Z0-9._-]/g, "_");
}

/** Move legacy .onenote/cache/<notebook>/ entries into .onenote/cache/<account>/<notebook>/ on first multi-account sync. */
async function migrateLegacyCacheToAccount(accountEmail: string, log: (msg: string) => void): Promise<void> {
  const { rename } = await import("node:fs/promises");
  const accountDir = join(CACHE_DIR, accountDirName(accountEmail));
  let entries: string[];
  try { entries = await readdir(CACHE_DIR); } catch { return; }
  const legacy = entries.filter((e) => !e.startsWith(".") && e !== accountDirName(accountEmail));
  if (legacy.length === 0) return;

  // Only migrate entries that look like notebook dirs (contain .json or .one files)
  const toMove: string[] = [];
  for (const name of legacy) {
    const p = join(CACHE_DIR, name);
    try {
      const s = await stat(p);
      if (!s.isDirectory()) continue;
      const files = await readdir(p);
      if (files.some((f) => f.endsWith(".json") || f.endsWith(".one"))) {
        // Skip if this dir itself looks like an account dir (contains subdirs that look like notebooks)
        const hasSubdirs = await Promise.all(files.map(async (f) => {
          try { return (await stat(join(p, f))).isDirectory(); } catch { return false; }
        }));
        if (hasSubdirs.some((x) => x)) continue; // already account-style
        toMove.push(name);
      }
    } catch {}
  }
  if (toMove.length === 0) return;

  log(`Migrating ${toMove.length} legacy notebook cache dir(s) to ${accountEmail}/...`);
  await ensureDir(accountDir);
  for (const name of toMove) {
    try { await rename(join(CACHE_DIR, name), join(accountDir, name)); } catch {}
  }
}

export async function syncCache(
  onProgress?: (msg: string) => void,
  emit?: SyncEmit
): Promise<void> {
  await ensureDir(CACHE_DIR);
  await ensureDir(join(PKG_ROOT, ".onenote"));
  const log = onProgress ?? console.log;

  const accounts = await listAccounts();
  if (accounts.length === 0) {
    log("No accounts logged in. Run `onenote auth login` first.");
    return;
  }

  // One-time migration: move legacy caches (.onenote/cache/<nb>/) into the first account's namespace
  await migrateLegacyCacheToAccount(accounts[0]!.username, log);

  log(`Syncing ${accounts.length} account${accounts.length > 1 ? "s" : ""}: ${accounts.map((a) => a.username).join(", ")}`);

  let idx = 0;
  for (const account of accounts) {
    idx++;
    setCurrentAccount(account);
    emit?.({ type: "account", email: account.username, index: idx, total: accounts.length });
    log(`\n=== ${account.username} ===`);
    try {
      await syncAccountCache(account.username, log, emit);
    } catch (err) {
      log(`  [error] ${account.username}: ${(err as Error).message}`);
    }
  }
  setCurrentAccount(undefined);
}

/** HTML-only sync for personal OneDrive accounts (no .one binary access).
 *  Lists pages via OneNote API, fetches each page's HTML, stores tags + text in DB. */
async function syncApiOnlySection(
  db: Database,
  notebookName: string,
  sectionName: string,
  sectionId: string,
  accountEmail: string,
  log: (msg: string) => void,
  emit?: SyncEmit
): Promise<number> {
  const pagesForHtml: { guid: string; apiId: string; title: string; webUrl: string }[] = [];
  let url: string | null = `/me/onenote/sections/${sectionId}/pages?$select=id,title,links&$top=100`;
  while (url) {
    const res = await graphFetchRaw(url);
    if (!res.ok) break;
    const data = (await res.json()) as any;
    for (const p of data.value ?? []) {
      const webUrl = p.links?.oneNoteWebUrl?.href ?? "";
      const guid = pageGuidFromWebUrl(webUrl) ?? p.id;
      pagesForHtml.push({ guid, apiId: p.id, title: p.title ?? "", webUrl });
    }
    url = data["@odata.nextLink"] ?? null;
  }
  if (pagesForHtml.length === 0) return 0;

  emit?.({ type: "tags", count: 0 });
  const htmlTags = await fetchTagsFromHtml(
    pagesForHtml.map((p) => ({ guid: p.guid, apiId: p.apiId })),
    log
  );

  // Also fetch HTML body text for search (separate pass to reuse fetchTagsFromHtml would double fetch; do combined)
  // Simpler: refetch HTML for the page body. But that doubles requests. Instead, extend fetchTagsFromHtml
  // to return bodies too. For now keep it simple — we already have tag_lines which contain the tagged text.
  db.transaction(() => {
    const insert = db.prepare(
      "INSERT INTO pages(section, notebook, title, body, web_url, page_guid, has_todo, has_done, tags, tag_lines, account) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) ON CONFLICT DO NOTHING"
    );
    // Delete existing entries for this section first
    db.run(
      `INSERT INTO pages_fts(pages_fts, rowid, title, body)
       SELECT 'delete', id, title, body FROM pages WHERE section=? AND notebook=? AND account=?`,
      [sectionName, notebookName, accountEmail]
    );
    db.run("DELETE FROM pages WHERE section=? AND notebook=? AND account=?", [sectionName, notebookName, accountEmail]);

    for (const p of pagesForHtml) {
      const info = htmlTags.get(p.guid);
      const tags = info?.tags ?? [];
      const lines = info?.lines ?? [];
      const tagsStr = tags.length > 0 ? `|${tags.join("|")}|` : null;
      const linesStr = lines.length > 0 ? JSON.stringify(lines) : null;
      const hasTodo = tags.includes("to-do") ? 1 : 0;
      const hasDone = tags.includes("to-do:completed") ? 1 : 0;
      // Body = concatenation of tag line texts (used for search when no binary body available)
      const body = lines.map((l) => l.text).join(" ").slice(0, 5000);
      insert.run(sectionName, notebookName, p.title, body, p.webUrl, p.guid, hasTodo, hasDone, tagsStr, linesStr, accountEmail);
    }
    // Rebuild FTS entries
    db.run(
      `INSERT INTO pages_fts(rowid, title, body)
       SELECT id, title, body FROM pages WHERE section=? AND notebook=? AND account=?`,
      [sectionName, notebookName, accountEmail]
    );
  })();

  return pagesForHtml.length;
}

async function syncAccountCache(
  accountEmail: string,
  log: (msg: string) => void,
  emit?: SyncEmit
): Promise<void> {
  const accountCacheDir = join(CACHE_DIR, accountDirName(accountEmail));
  await ensureDir(accountCacheDir);
  const db = openSearchDb();

  const notebooks = await listNotebooks();
  log(`Found ${notebooks.length} notebooks`);
  emit?.({ type: "total", total: 0, notebooks: notebooks.length });

  // Collect all sections across notebooks for progress tracking.
  // Sync smaller sections first so quick wins land early and huge sections
  // don't block progress.
  const sectionsByNotebook = await Promise.all(
    notebooks.map(async (nb) => {
      const nbDir = join(accountCacheDir, nb.displayName);
      await ensureDir(nbDir);
      const sections = await listSectionFiles(nb);
      return sections.map((sec) => ({
        nb,
        sec,
        nbDir,
        cachePath: join(nbDir, `${sec.name}.json`),
      }));
    })
  );
  const allSections = sectionsByNotebook
    .flat()
    .toSorted((a, b) => (a.sec.size ?? 0) - (b.sec.size ?? 0));

  let downloaded = 0;
  let skipped = 0;
  let retagged = 0;
  const total = allSections.length;
  log(`${total} sections across ${notebooks.length} notebooks`);
  emit?.({ type: "total", total, notebooks: notebooks.length });

  // Query for sections that already have tag_lines populated (to skip retagging them)
  const taggedSectionsStmt = db.prepare(
    "SELECT COUNT(*) as c FROM pages WHERE section=? AND notebook=? AND tag_lines IS NOT NULL"
  );
  const hasTagLines = (section: string, notebook: string): boolean =>
    ((taggedSectionsStmt.get(section, notebook) as { c: number } | null)?.c ?? 0) > 0;

  for (const { nb, sec, nbDir, cachePath } of allSections) {
      // Personal account / API-only sections: HTML-only sync (no binary download)
      if (sec.apiOnly && sec.sectionId) {
        const idx = downloaded + skipped + retagged + 1;
        log(`  [${idx}/${total}] ${nb.displayName}/${sec.name} (api-only)`);
        emit?.({ type: "section", index: idx, notebook: nb.displayName, section: sec.name });
        try {
          const pageCount = await syncApiOnlySection(
            db, nb.displayName, sec.name, sec.sectionId, accountEmail, log, emit
          );
          emit?.({ type: "done", notebook: nb.displayName, section: sec.name, pages: pageCount, status: "ok" });
          downloaded++;
        } catch (err) {
          log(`    [failed] ${sec.name}: ${(err as Error).message}`);
          emit?.({ type: "done", notebook: nb.displayName, section: sec.name, pages: 0, status: "failed" });
          skipped++;
        }
        continue;
      }

      // Incremental sync: skip if cache is fresh AND source hasn't changed
      let cacheIsFresh = false;
      try {
        const raw = await readFile(cachePath, "utf-8");
        const cached = JSON.parse(raw);
        if (cached.lastModified && sec.lastModified && cached.lastModified >= sec.lastModified) {
          cacheIsFresh = true;
        }
      } catch {}

      if (cacheIsFresh) {
        // Still run a light-weight tag refresh if DB doesn't have tag_lines for this section
        if (!hasTagLines(sec.name, nb.displayName)) {
          const webUrl = await getSectionWebUrl(sec.drivePath);
          const sourcedocMatch = webUrl.match(/sourcedoc=%7B([0-9a-f-]+)%7D/i);
          const sectionGuid = sourcedocMatch?.[1]?.toLowerCase();
          if (sectionGuid) {
            const officialPages = await getOneNotePagesForSection(sectionGuid);
            const pagesForHtml = officialPages
              .map((p) => { const g = pageGuidFromWebUrl(p.webUrl); return g ? { guid: g, apiId: p.id } : null; })
              .filter(Boolean) as { guid: string; apiId: string }[];
            if (pagesForHtml.length > 0) {
              log(`  [${retagged + downloaded + skipped + 1}/${total}] ${nb.displayName}/${sec.name} (retag only)`);
              emit?.({ type: "retag", index: retagged + downloaded + skipped + 1, notebook: nb.displayName, section: sec.name });
              const htmlTags = await fetchTagsFromHtml(pagesForHtml, log);
              if (htmlTags.size > 0) {
                emit?.({ type: "tags", count: htmlTags.size });
                const updateStmt = db.prepare(
                  "UPDATE pages SET tags=?, tag_lines=?, has_todo=?, has_done=?, account=COALESCE(account, ?) WHERE page_guid=? AND section=? AND notebook=?"
                );
                db.transaction(() => {
                  for (const [guid, info] of htmlTags) {
                    const tagsStr = `|${info.tags.join("|")}|`;
                    const linesStr = info.lines.length > 0 ? JSON.stringify(info.lines) : null;
                    const hasTodo = info.tags.includes("to-do") ? 1 : 0;
                    const hasDone = info.tags.includes("to-do:completed") ? 1 : 0;
                    updateStmt.run(tagsStr, linesStr, hasTodo, hasDone, accountEmail, guid, sec.name, nb.displayName);
                  }
                })();
              }
              retagged++;
              emit?.({ type: "done", notebook: nb.displayName, section: sec.name, pages: 0, status: "skipped" });
              continue;
            }
          }
        }
        skipped++;
        emit?.({ type: "done", notebook: nb.displayName, section: sec.name, pages: 0, status: "skipped" });
        continue;
      }

      const idx = downloaded + skipped + retagged + 1;
      const sizeStr = sec.size ? prettyBytes(sec.size, { binary: true }) : undefined;
      log(`  [${idx}/${total}] ${nb.displayName}/${sec.name}${sizeStr ? ` (${sizeStr})` : ""}`);
      emit?.({ type: "section", index: idx, notebook: nb.displayName, section: sec.name, size: sizeStr });
      const buf = await downloadSection(sec.drivePath);
      if (!buf) {
        log(`    [failed] ${sec.name}`);
        skipped++;
        emit?.({ type: "done", notebook: nb.displayName, section: sec.name, pages: 0, status: "failed" });
        continue;
      }
      downloaded++;

      const pages = extractPages(buf);
      const guidEntries = extractPageGuids(buf).sort((a, b) => a.offset - b.offset);
      const webUrl = await getSectionWebUrl(sec.drivePath);

      // Extract section's sourcedoc GUID from webUrl, then fetch official page list via OneNote API
      const sourcedocMatch = webUrl.match(/sourcedoc=%7B([0-9a-f-]+)%7D/i);
      const sectionGuid = sourcedocMatch?.[1]?.toLowerCase();
      const officialPages = sectionGuid
        ? await getOneNotePagesForSection(sectionGuid)
        : [];
      // Map: pageGuid -> official webUrl (extract GUID from webUrl, not from id)
      const officialUrlByGuid = new Map<string, { url: string; title: string }>();
      const officialApiIdByGuid = new Map<string, string>();
      for (const op of officialPages) {
        const guid = pageGuidFromWebUrl(op.webUrl);
        if (guid && op.webUrl) {
          officialUrlByGuid.set(guid, { url: op.webUrl, title: op.title });
          officialApiIdByGuid.set(guid, op.id);
        }
      }

      // Fetch HTML for each official page to collect all tag types
      const pagesForHtml = [...officialApiIdByGuid.entries()].map(([guid, apiId]) => ({ guid, apiId }));
      if (pagesForHtml.length > 0) log(`    fetching tags (${pagesForHtml.length} pages)...`);
      const htmlTags = await fetchTagsFromHtml(pagesForHtml, log);

      // Save the binary as base64 for accurate position-based search
      // Limit raw cache to <50MB sections to control disk usage
      const includeRaw = buf.length < 50 * 1024 * 1024;

      const cacheData = {
        section: sec.name,
        notebook: nb.displayName,
        webUrl,
        lastModified: sec.lastModified,
        pages: pages.map((p: any) => ({
          title: p.title,
          body: p.body,
          pageGuid: p.pageGuid,
          officialUrl: p.pageGuid ? officialUrlByGuid.get(p.pageGuid)?.url : undefined,
        })),
        // Use OneNote API page list as authoritative anchor list (with official URLs)
        officialPages: officialPages.map((p) => ({
          guid: pageGuidFromWebUrl(p.webUrl),
          title: p.title,
          webUrl: p.webUrl,
          apiId: p.id,
        })),
        anchors: guidEntries.map((e) => ({ offset: e.offset, guid: e.guid, title: e.title })),
        rawSize: buf.length,
        cachedAt: new Date().toISOString(),
      };

      // Save raw .one file alongside JSON for binary search
      if (includeRaw) {
        const binPath = cachePath.replace(/\.json$/, ".one");
        await writeFile(binPath, buf);
      }

      await writeFile(cachePath, JSON.stringify(cacheData));
      upsertSectionToIndex(
        db, sec.name, nb.displayName, cacheData.pages, webUrl,
        buf, // always pass buf to FTS (even >50MB sections); only disk save is gated on includeRaw
        cacheData.anchors,
        cacheData.officialPages,
        htmlTags,
        accountEmail
      );
      log(`    [ok] ${sec.name} (${pages.length} pages)`);
      emit?.({ type: "done", notebook: nb.displayName, section: sec.name, pages: pages.length, status: "ok" });
  }
  db.close();
  log(`Sync complete. ${downloaded} downloaded, ${retagged} retagged, ${skipped} up-to-date.`);
  emit?.({ type: "complete", downloaded, retagged, skipped });
}

/**
 * Build a page-level OneNote Online URL.
 * Format: {sectionUrl}&wd=target({escapedTitle}|{pageGuid}/)
 *
 * Note: OneNote Online caches the user's last-viewed page within a section.
 * When opened in a session that has previously viewed the section, OneNote may
 * redirect to the cached page instead of honoring the wd=target parameter.
 * The URL is still a correct page-level permalink.
 */
function buildPageUrl(
  sectionUrl: string,
  pageTitle: string,
  pageGuid?: string
): string {
  if (!pageGuid || !sectionUrl) return sectionUrl;
  // OneNote escapes only `)` and `|` in titles with `\`
  const escapedTitle = pageTitle
    .replace(/\\/g, "\\\\")
    .replace(/\)/g, "\\)")
    .replace(/\|/g, "\\|");
  const wd = `target(${escapedTitle}|${pageGuid}/)`;
  // Strict encoding (also encode parens) to match OneNote's own URL format
  const encoded = encodeURIComponent(wd).replace(
    /[!'()*]/g,
    (c) => "%" + c.charCodeAt(0).toString(16).toUpperCase()
  );
  const separator = sectionUrl.includes("?") ? "&" : "?";
  return `${sectionUrl}${separator}wd=${encoded}`;
}

export async function isCacheEmpty(): Promise<boolean> {
  try {
    const nbDirs = await readdir(CACHE_DIR);
    for (const nb of nbDirs) {
      const nbDir = join(CACHE_DIR, nb);
      const s = await stat(nbDir);
      if (!s.isDirectory()) continue;
      const files = await readdir(nbDir);
      if (files.some((f) => f.endsWith(".json"))) return false;
    }
  } catch {}
  return true;
}

/**
 * Find all binary positions where `needle` appears in `buf`.
 * Searches for both UTF-8 and UTF-16LE encodings.
 */
function findAllInBinary(buf: Buffer, needle: string): number[] {
  const results: number[] = [];
  const utf8 = Buffer.from(needle, "utf-8");
  let pos = 0;
  while ((pos = buf.indexOf(utf8, pos)) !== -1) { results.push(pos); pos++; }
  const utf16 = Buffer.from(needle, "utf16le");
  pos = 0;
  while ((pos = buf.indexOf(utf16, pos)) !== -1) { results.push(pos); pos += 2; }
  return results;
}

function getNearestPrecedingAnchor(
  anchors: { offset: number; guid: string; title: string }[],
  pos: number
): { guid: string; title: string } | null {
  let best: { guid: string; title: string } | null = null;
  for (const a of anchors) {
    if (a.offset <= pos) best = { guid: a.guid, title: a.title };
    else break;
  }
  return best;
}

/**
 * Find the page that owns this binary position by checking which page title
 * appears in the surrounding context. Falls back to nearest preceding anchor.
 */
function findOwnerPage(
  buf: Buffer,
  anchors: { offset: number; guid: string; title: string }[],
  pos: number
): { guid: string; title: string } | null {
  // Get a 20KB context around the match
  const ctxStart = Math.max(0, pos - 10000);
  const ctxEnd = Math.min(buf.length, pos + 10000);
  const ctx = buf.slice(ctxStart, ctxEnd);
  // Decode as both UTF-8 and UTF-16LE to catch all titles
  const ctxUtf8 = ctx.toString("utf-8");
  const ctxUtf16 = ctx.toString("utf16le");

  // For each known anchor, check if its title appears in context.
  // Prefer the longest matching title (more specific = more likely correct).
  const candidates: { anchor: typeof anchors[number]; matchLen: number }[] = [];
  for (const a of anchors) {
    const t = a.title.trim();
    if (t.length < 4) continue;
    // Use a substring of the title (first 30 chars) for matching
    const probe = t.slice(0, 30);
    if (ctxUtf8.includes(probe) || ctxUtf16.includes(probe)) {
      candidates.push({ anchor: a, matchLen: probe.length });
    }
  }

  if (candidates.length > 0) {
    // Return the candidate with the longest matched title
    candidates.sort((a, b) => b.matchLen - a.matchLen);
    return { guid: candidates[0].anchor.guid, title: candidates[0].anchor.title };
  }

  // Fallback: nearest preceding anchor
  return getNearestPrecedingAnchor(anchors, pos);
}

function extractContextFromBinary(
  buf: Buffer,
  pos: number,
  query: string
): string {
  // Try both UTF-8 and UTF-16LE; pick the one that contains the query
  // Use a wider context to be sure
  const queryByteLen = Math.max(
    Buffer.byteLength(query, "utf-8"),
    Buffer.byteLength(query, "utf16le")
  );
  const start = Math.max(0, pos - 200);
  const end = Math.min(buf.length, pos + queryByteLen + 200);
  const segment = buf.slice(start, end);

  const cleanText = (s: string): string =>
    s.replace(/[\x00-\x1F\x7F\uFFFD]/g, " ").replace(/\s+/g, " ").trim();

  // Try UTF-8
  const utf8Text = cleanText(segment.toString("utf-8"));
  if (utf8Text.includes(query)) return utf8Text;

  // Try UTF-16LE at this offset and offset+1
  const utf16Text = cleanText(segment.toString("utf16le"));
  if (utf16Text.includes(query)) return utf16Text;

  // Try UTF-16LE with shifted alignment
  if (segment.length > 1) {
    const shifted = cleanText(segment.slice(1).toString("utf16le"));
    if (shifted.includes(query)) return shifted;
  }

  // Fallback: return UTF-8 even if it doesn't contain the query
  return utf8Text || utf16Text;
}

function cleanSnippet(raw: string, query: string): string {
  const stripped = raw
    .replace(/<[^>]{0,300}>/g, " ")
    .replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&nbsp;/g, " ")
    .replace(/[\x00-\x09\x0B-\x1F\x7F]/g, " "); // keep \x0A (LF) so split(/\n+/) works
  const lq = query.toLowerCase();
  // Split into lines, find lines that contain the query
  const lines = stripped.split(/\n+/).map((l) => l.replace(/\s+/g, " ").trim()).filter(Boolean);
  const matchingLines = lines.filter((l) => l.toLowerCase().includes(lq));
  if (matchingLines.length === 0) return "";
  // Prefer lines with better ASCII ratio (less binary garbage)
  const best = matchingLines.sort((a, b) => asciiRatio(b) - asciiRatio(a))[0]!;
  // If the best line is still mostly garbage, say so
  if (asciiRatio(best) < 0.15) return "[binary content, no clean preview]";
  // Trim to a window around the query
  const idx = best.toLowerCase().indexOf(lq);
  const start = Math.max(0, idx - 30);
  const end = Math.min(best.length, idx + query.length + 30);
  const excerpt = best.slice(start, end).trim();
  return (start > 0 ? "..." : "") + excerpt + (end < best.length ? "..." : "");
}

async function searchSection(jsonPath: string, query: string): Promise<CachedPage[]> {
  const results: CachedPage[] = [];
  try {
    const raw = await readFile(jsonPath, "utf-8");
    const data = JSON.parse(raw);

    const binPath = jsonPath.replace(/\.json$/, ".one");
    let binBuf: Buffer | null = null;
    try {
      binBuf = await readFile(binPath);
    } catch {}

    if (binBuf && data.anchors) {
      const positions = findAllInBinary(binBuf, query);
      const byGuid = new Map<string, { title: string; positions: number[] }>();
      for (const pos of positions) {
        const anchor = findOwnerPage(binBuf, data.anchors, pos);
        if (!anchor) continue;
        const existing = byGuid.get(anchor.guid);
        if (existing) existing.positions.push(pos);
        else byGuid.set(anchor.guid, { title: anchor.title, positions: [pos] });
      }
      const officialByGuid = new Map<string, { url: string; title: string }>();
      for (const op of data.officialPages ?? []) {
        if (op.guid && op.webUrl) officialByGuid.set(op.guid, { url: op.webUrl, title: op.title });
      }
      const bodyByGuid = new Map<string, string>();
      for (const p of data.pages ?? []) {
        if (p.pageGuid && p.body) bodyByGuid.set(p.pageGuid, p.body);
      }
      for (const [guid, info] of byGuid) {
        const firstPos = info.positions[0]!;
        const cleanBody = bodyByGuid.get(guid);
        const rawContext = cleanBody ?? extractContextFromBinary(binBuf, firstPos, query);
        const official = officialByGuid.get(guid);
        const pageUrl = official?.url ?? buildPageUrl(data.webUrl, info.title, guid);
        results.push({
          title: official?.title ?? info.title,
          body: cleanSnippet(rawContext, query),
          section: data.section,
          notebook: data.notebook,
          webUrl: pageUrl,
          pageGuid: guid,
        });
      }
    } else {
      const lowerQuery = query.toLowerCase();
      for (const page of data.pages ?? []) {
        if (page.body?.toLowerCase().includes(lowerQuery)) {
          const pageUrl = page.officialUrl ?? buildPageUrl(data.webUrl, page.title, page.pageGuid);
          results.push({
            title: page.title,
            body: cleanSnippet(page.body, query),
            section: data.section,
            notebook: data.notebook,
            webUrl: pageUrl,
            pageGuid: page.pageGuid,
          });
        }
      }
    }
  } catch {}
  return results;
}

export async function searchLocal(
  query: string,
  { offset = 0, limit = 100, notebook, section }: { offset?: number; limit?: number; notebook?: string; section?: string } = {}
): Promise<CachedPage[]> {
  const { ftsQuery, hasTodo, hasDone, hasCheckbox, tagFilters } = parseTagsFromQuery(query);
  // Use FTS5 index if available (built during sync)
  try {
    const db = new Database(SEARCH_DB_PATH, { readonly: true });

    type Row = { section: string; notebook: string; title: string; body: string; web_url: string; page_guid: string; tag_lines: string | null };
    let rows: Row[];

    // Build tag filter SQL: each tagFilter value needs tags LIKE '%|value|%'
    const addTagConditions = (conditions: string[], params: (string | number | null)[], prefix = "") => {
      if (hasTodo) conditions.push(`${prefix}has_todo = 1`);
      if (hasDone) conditions.push(`${prefix}has_done = 1`);
      if (hasCheckbox) conditions.push(`(${prefix}has_todo = 1 OR ${prefix}has_done = 1)`);
      for (const tagVal of tagFilters) {
        conditions.push(`${prefix}tags LIKE ?`);
        params.push(`%|${tagVal}|%`);
      }
    };

    if (ftsQuery) {
      const ftsParam = buildFtsParam(ftsQuery);
      const conditions: string[] = ["pages_fts MATCH ?"];
      const params: (string | number | null)[] = [ftsParam];
      if (notebook) { conditions.push("p.notebook LIKE ?"); params.push(`%${notebook}%`); }
      if (section) { conditions.push("p.section LIKE ?"); params.push(`%${section}%`); }
      addTagConditions(conditions, params, "p.");
      params.push(limit, offset);
      rows = db.query<Row, (string | number | null)[]>(
        `SELECT p.section, p.notebook, p.title, p.body, p.web_url, p.page_guid, p.tag_lines
         FROM pages_fts f JOIN pages p ON f.rowid = p.id
         WHERE ${conditions.join(" AND ")} ORDER BY rank LIMIT ? OFFSET ?`
      ).all(...params);
    } else {
      const conditions: string[] = [];
      const params: (string | number | null)[] = [];
      if (notebook) { conditions.push("notebook LIKE ?"); params.push(`%${notebook}%`); }
      if (section) { conditions.push("section LIKE ?"); params.push(`%${section}%`); }
      addTagConditions(conditions, params);
      const where = conditions.length ? `WHERE ${conditions.join(" AND ")}` : "";
      params.push(limit, offset);
      // Prefer pages with actual tag_lines (fetched from HTML) over binary-only detections
      rows = db.query<Row, (string | number | null)[]>(
        `SELECT section, notebook, title, body, web_url, page_guid, tag_lines FROM pages ${where}
         ORDER BY (tag_lines IS NOT NULL) DESC, id LIMIT ? OFFSET ?`
      ).all(...params);
    }

    db.close();
    return rows.map((r) => {
      let tagLines: Array<{ tag: string; text: string }> | undefined;
      if (r.tag_lines) { try { tagLines = JSON.parse(r.tag_lines); } catch {} }
      return {
        title: r.title,
        body: ftsQuery ? cleanSnippet(r.body, ftsQuery) : r.body?.split("\n")[0] ?? "",
        section: r.section,
        notebook: r.notebook,
        webUrl: r.web_url,
        pageGuid: r.page_guid,
        tagLines,
      };
    });
  } catch {
    // DB not found or FTS error — fall back to file scan
  }

  // File-scan fallback (used before first sync or if DB is missing)
  let nbDirs: string[];
  try {
    nbDirs = await readdir(CACHE_DIR);
  } catch {
    return [];
  }

  const jsonPaths: string[] = [];
  for (const nbName of nbDirs) {
    const nbDir = join(CACHE_DIR, nbName);
    try {
      const s = await stat(nbDir);
      if (!s.isDirectory()) continue;
      const files = await readdir(nbDir);
      for (const file of files) {
        if (file.endsWith(".json")) jsonPaths.push(join(nbDir, file));
      }
    } catch {
      continue;
    }
  }

  const CONCURRENCY = 20;
  const allResults: CachedPage[] = [];
  for (let i = 0; i < jsonPaths.length; i += CONCURRENCY) {
    const chunk = jsonPaths.slice(i, i + CONCURRENCY);
    const chunkResults = await Promise.all(chunk.map((p) => searchSection(p, query)));
    for (const r of chunkResults) allResults.push(...r);
  }

  return allResults;
}

export async function rebuildSearchIndex(onProgress?: (msg: string) => void): Promise<void> {
  const log = onProgress ?? console.log;
  await ensureDir(join(PKG_ROOT, ".onenote"));
  // Drop and recreate: contentless FTS5 DELETE has no effect on the posting list,
  // and pages table truncation doesn't sync to FTS shadow tables.
  const { unlink } = await import("node:fs/promises");
  for (const ext of ["", "-shm", "-wal"]) {
    try { await unlink(SEARCH_DB_PATH + ext); } catch {}
  }
  const db = openSearchDb();

  let nbDirs: string[];
  try {
    nbDirs = await readdir(CACHE_DIR);
  } catch {
    db.close();
    return;
  }

  let count = 0;
  for (const nbName of nbDirs) {
    const nbDir = join(CACHE_DIR, nbName);
    try {
      const s = await stat(nbDir);
      if (!s.isDirectory()) continue;
      const files = await readdir(nbDir);
      for (const file of files) {
        if (!file.endsWith(".json")) continue;
        try {
          const jsonPath = join(nbDir, file);
          const raw = await readFile(jsonPath, "utf-8");
          const data = JSON.parse(raw);
          let binBuf: Buffer | undefined;
          try { binBuf = await readFile(jsonPath.replace(/\.json$/, ".one")); } catch {}
          upsertSectionToIndex(
            db, data.section, data.notebook, data.pages ?? [], data.webUrl ?? "",
            binBuf, data.anchors, data.officialPages
          );
          count++;
        } catch {}
      }
    } catch {}
  }

  db.close();
  log(`Search index rebuilt: ${count} sections indexed.`);
}
