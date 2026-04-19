import prettyBytes from "pretty-bytes";
import { getAccessToken } from "./auth";
import { listNotebooks } from "./graph";
import { readFile, writeFile, mkdir, readdir, stat } from "node:fs/promises";
import { join } from "node:path";

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
  return fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
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
  try {
    const res = await graphFetchRaw(
      `/me/onenote/sections/0-${sectionGuid}/pages?$select=id,title,links&$top=100`
    );
    if (!res.ok) return [];
    const data = (await res.json()) as any;
    return (data.value ?? []).map((p: any) => ({
      id: p.id,
      title: p.title ?? "",
      webUrl: p.links?.oneNoteWebUrl?.href ?? "",
    }));
  } catch {
    return [];
  }
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
): Promise<{ name: string; drivePath: string; lastModified: string; size: number }[]> {
  const nbPath = getNotebookDrivePath(notebook);
  if (!nbPath) return [];
  const encoded = nbPath
    .split("/")
    .map((s) => encodeURIComponent(s))
    .join("/");
  try {
    const res = await graphFetchRaw(
      `/me/drive/root:/${encoded}:/children?$select=name,id,file,size,lastModifiedDateTime&$top=200`
    );
    if (!res.ok) return [];
    const data = (await res.json()) as any;
    return (data.value ?? [])
      .filter((f: any) => f.name?.endsWith(".one"))
      .map((f: any) => ({
        name: f.name.replace(/\.one$/, ""),
        drivePath: `${nbPath}/${f.name}`,
        lastModified: f.lastModifiedDateTime ?? "",
        size: f.size ?? 0,
      }));
  } catch {
    return [];
  }
}

export async function syncCache(
  onProgress?: (msg: string) => void
): Promise<void> {
  await ensureDir(CACHE_DIR);
  const log = onProgress ?? console.log;

  const notebooks = await listNotebooks();
  log(`Found ${notebooks.length} notebooks`);

  // Collect all sections across notebooks for progress tracking.
  // Sync smaller sections first so quick wins land early and huge sections
  // don't block progress.
  const sectionsByNotebook = await Promise.all(
    notebooks.map(async (nb) => {
      const nbDir = join(CACHE_DIR, nb.displayName);
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
  const total = allSections.length;
  log(`${total} sections across ${notebooks.length} notebooks`);

  for (const { nb, sec, nbDir, cachePath } of allSections) {
      // Incremental sync: skip if cache is fresh AND source hasn't changed
      try {
        const raw = await readFile(cachePath, "utf-8");
        const cached = JSON.parse(raw);
        // If cache has a lastModified field, compare with OneDrive's lastModifiedDateTime
        if (cached.lastModified && sec.lastModified && cached.lastModified >= sec.lastModified) {
          skipped++;
          continue;
        }
      } catch {
        // No cache or unreadable — download
      }

      const progress = `[${downloaded + skipped + 1}/${total}]`;
      const sizeStr = sec.size ? ` (${prettyBytes(sec.size, { binary: true })})` : "";
      log(`  ${progress} ${nb.displayName}/${sec.name}${sizeStr}`);
      const buf = await downloadSection(sec.drivePath);
      if (!buf) {
        log(`    [failed] ${sec.name}`);
        skipped++;
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
      for (const op of officialPages) {
        const guid = pageGuidFromWebUrl(op.webUrl);
        if (guid && op.webUrl) officialUrlByGuid.set(guid, { url: op.webUrl, title: op.title });
      }

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
      log(`    [ok] ${sec.name} (${pages.length} pages)`);
  }
  log(`Sync complete. ${downloaded} downloaded, ${skipped} up-to-date.`);
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
  for (let i = 0; i < buf.length - utf8.length; i++) {
    let m = true;
    for (let j = 0; j < utf8.length; j++) if (buf[i + j] !== utf8[j]) { m = false; break; }
    if (m) results.push(i);
  }
  const utf16 = Buffer.from(needle, "utf16le");
  for (let i = 0; i < buf.length - utf16.length; i++) {
    let m = true;
    for (let j = 0; j < utf16.length; j++) if (buf[i + j] !== utf16[j]) { m = false; break; }
    if (m) results.push(i);
  }
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

function asciiRatio(s: string): number {
  if (!s.length) return 0;
  const printable = s.split("").filter((c) => c >= "\x20" && c <= "\x7E").length;
  return printable / s.length;
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

export async function searchLocal(query: string): Promise<CachedPage[]> {
  const results: CachedPage[] = [];

  let nbDirs: string[];
  try {
    nbDirs = await readdir(CACHE_DIR);
  } catch {
    return results;
  }

  for (const nbName of nbDirs) {
    const nbDir = join(CACHE_DIR, nbName);
    let files: string[];
    try {
      const s = await stat(nbDir);
      if (!s.isDirectory()) continue;
      files = await readdir(nbDir);
    } catch {
      continue;
    }

    for (const file of files) {
      if (!file.endsWith(".json")) continue;
      try {
        const jsonPath = join(nbDir, file);
        const raw = await readFile(jsonPath, "utf-8");
        const data = JSON.parse(raw);

        // Try binary-based search first if .one is cached
        const binPath = jsonPath.replace(/\.json$/, ".one");
        let binBuf: Buffer | null = null;
        try {
          binBuf = await readFile(binPath);
        } catch {}

        if (binBuf && data.anchors) {
          const positions = findAllInBinary(binBuf, query);
          // Group positions by owner page (using context-based lookup)
          const byGuid = new Map<string, { title: string; positions: number[] }>();
          for (const pos of positions) {
            const anchor = findOwnerPage(binBuf, data.anchors, pos);
            if (!anchor) continue;
            const existing = byGuid.get(anchor.guid);
            if (existing) existing.positions.push(pos);
            else byGuid.set(anchor.guid, { title: anchor.title, positions: [pos] });
          }
          // Build map of official URLs by GUID (from OneNote API)
          const officialByGuid = new Map<string, { url: string; title: string; body?: string }>();
          for (const op of data.officialPages ?? []) {
            if (op.guid && op.webUrl) officialByGuid.set(op.guid, { url: op.webUrl, title: op.title });
          }
          // Build body map from JSON pages for clean snippets (avoids binary garbage)
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
            const displayTitle = official?.title ?? info.title;
            results.push({
              title: displayTitle,
              body: cleanSnippet(rawContext, query),
              section: data.section,
              notebook: data.notebook,
              webUrl: pageUrl,
              pageGuid: guid,
            });
          }
        } else {
          // Fallback: search in extracted page bodies
          const lowerQuery = query.toLowerCase();
          for (const page of data.pages ?? []) {
            if (page.body?.toLowerCase().includes(lowerQuery)) {
              const pageUrl = buildPageUrl(data.webUrl, page.title, page.pageGuid);
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
      } catch {
        continue;
      }
    }
  }

  return results;
}
