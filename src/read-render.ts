import { createHash } from "node:crypto";
import { mkdir, stat, writeFile } from "node:fs/promises";
import { dirname, join, relative, sep } from "node:path";
import { getAccessToken } from "./auth";

const PKG_ROOT = dirname(import.meta.dir);
const READ_ASSET_DIR = process.env.ONENOTE_READ_ASSET_DIR
  || join(PKG_ROOT, ".onenote", "assets");

type ResourceReference = {
  alt: string;
  url: string;
  mediaType?: string;
};

export type RenderOptions = {
  /** When false, resources are not downloaded and links keep their remote URL. */
  downloadAssets?: boolean;
  /** Directory to write downloaded media into (default: package .onenote/assets). */
  assetDir?: string;
  /** Directory that emitted relative links are computed from (default: cwd). */
  linkBaseDir?: string;
};

function decodeHtmlEntities(text: string): string {
  return text
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
}

// Convert struck-through runs to markdown `~~…~~` before tags are stripped.
// OneNote emits both <s>/<strike>/<del> and styled spans (text-decoration:line-through).
function markStrikethrough(html: string): string {
  return html
    .replace(/<(?:s|strike|del)\b[^>]*>([\s\S]*?)<\/(?:s|strike|del)>/gi, "~~$1~~")
    .replace(
      /<span\b[^>]*text-decoration\s*:[^"';>]*line-through[^>]*>([\s\S]*?)<\/span>/gi,
      "~~$1~~"
    );
}

function extractAttr(tag: string, name: string): string | undefined {
  const match = tag.match(new RegExp(`${name}=(["'])(.*?)\\1`, "i"));
  return match?.[2];
}

function escapeMarkdownText(text: string): string {
  return text.replace(/[[\]\\]/g, "\\$&");
}

export type OneNoteLinkInfo = {
  basePath: string; // notebook web URL the link was authored against (may be stale)
  file: string; // section ".one" filename
  sectionId: string; // section object GUID
  pageId: string; // page object GUID
  title: string; // page title
};

/**
 * A resolver that, given a parsed OneNote link, returns the page's *current*
 * https URL (e.g. from the local sync cache or a fresh Graph lookup), or null to
 * fall back to reconstructing it from the link's embedded base-path. Injected by
 * the CLI so this module stays free of graph/cache dependencies.
 */
export type LinkResolver = (info: OneNoteLinkInfo) => Promise<string | null>;

let linkResolver: LinkResolver | null = null;
export function setLinkResolver(fn: LinkResolver | null): void {
  linkResolver = fn;
}

/** Parse a `onenote:` client deep-link into its parts, or null if not one. */
function parseOneNoteHref(href: string): OneNoteLinkInfo | null {
  if (!/^onenote:/i.test(href)) return null;
  const rest = href.replace(/^onenote:/i, "");
  const hashIdx = rest.indexOf("#");
  const file = decodeURIComponent(hashIdx >= 0 ? rest.slice(0, hashIdx) : rest);
  const frag = hashIdx >= 0 ? rest.slice(hashIdx + 1) : "";
  const params = frag.split("&");
  const title = decodeURIComponent(params[0] ?? "");
  const get = (k: string) => {
    const p = params.find((x) => x.toLowerCase().startsWith(`${k}=`));
    return p ? p.slice(k.length + 1) : "";
  };
  const sectionId = get("section-id").replace(/[{}]/g, "").toLowerCase();
  const pageId = get("page-id").replace(/[{}]/g, "").toLowerCase();
  const basePath = get("base-path");
  if (!basePath || !/^https?:\/\//i.test(basePath) || !file) return null;
  return { basePath, file, sectionId, pageId, title };
}

/**
 * Build the equivalent https OneNote Online URL
 * (`{base-path}?wd=target(file|section-guid/title|page-guid/)`) from parsed link
 * info. This is the same form Graph's oneNoteWebUrl uses; browsers and web
 * Markdown renderers can open it. Returns null without a section GUID.
 */
export function buildOneNoteWebUrl(info: OneNoteLinkInfo): string | null {
  // OneNote's own web URLs percent-encode the whole target(...) argument,
  // including parentheses (which encodeURIComponent leaves untouched).
  const pct = (s: string) =>
    encodeURIComponent(s).replace(/[()!*'~]/g, (c) => `%${c.charCodeAt(0).toString(16).toUpperCase()}`);
  let inner: string;
  if (info.sectionId && info.pageId) inner = `${info.file}|${info.sectionId}/${info.title}|${info.pageId}/`;
  else if (info.sectionId) inner = `${info.file}|${info.sectionId}/`;
  else return null;
  return `${info.basePath}?wd=target${pct(`(${inner})`)}`;
}

/** Resolve an anchor href to its best https/onenote URL for markdown output. */
async function resolveLinkHref(href: string): Promise<string> {
  if (/^https?:\/\//i.test(href)) return href;
  const info = parseOneNoteHref(href);
  if (!info) return href;
  if (linkResolver) {
    try {
      const resolved = await linkResolver(info);
      if (resolved) return resolved;
    } catch {}
  }
  return buildOneNoteWebUrl(info) ?? href;
}

function sanitizeStem(text: string): string {
  return text
    .replace(/[^a-zA-Z0-9._-]+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "")
    .slice(0, 120);
}

function toDisplayPath(path: string, baseDir: string = process.cwd()): string {
  const rel = relative(baseDir, path).split(sep).join("/");
  if (!rel || rel === "") return ".";
  if (rel.startsWith("../") || rel.startsWith("./")) return rel;
  if (rel.startsWith(".")) return `./${rel}`;
  return `./${rel}`;
}

function extensionFromMediaType(mediaType?: string): string {
  switch ((mediaType ?? "").toLowerCase()) {
    case "image/jpeg":
    case "image/jpg":
      return "jpg";
    case "image/png":
      return "png";
    case "image/gif":
      return "gif";
    case "image/webp":
      return "webp";
    case "image/bmp":
      return "bmp";
    case "image/tiff":
      return "tiff";
    case "image/svg+xml":
      return "svg";
    case "application/pdf":
      return "pdf";
    default: {
      const subtype = mediaType?.split("/")[1]?.split(";")[0];
      return subtype ? subtype.replace(/[^a-z0-9]/gi, "").toLowerCase() : "bin";
    }
  }
}

function getResourceId(resourceUrl: string): string {
  try {
    const { pathname } = new URL(resourceUrl);
    const match = pathname.match(/\/onenote\/resources\/([^/]+)\/(?:\$value|content)$/i);
    if (match?.[1]) return match[1];
  } catch {}
  return createHash("sha1").update(resourceUrl).digest("hex");
}

function sniffMediaType(buf: Buffer): string | undefined {
  if (buf.length >= 3 && buf[0] === 0xff && buf[1] === 0xd8 && buf[2] === 0xff) {
    return "image/jpeg";
  }
  if (
    buf.length >= 8
    && buf[0] === 0x89
    && buf[1] === 0x50
    && buf[2] === 0x4e
    && buf[3] === 0x47
  ) {
    return "image/png";
  }
  if (buf.length >= 6 && buf.subarray(0, 6).toString("ascii") === "GIF87a") return "image/gif";
  if (buf.length >= 6 && buf.subarray(0, 6).toString("ascii") === "GIF89a") return "image/gif";
  if (buf.length >= 12 && buf.subarray(0, 4).toString("ascii") === "RIFF" && buf.subarray(8, 12).toString("ascii") === "WEBP") {
    return "image/webp";
  }
  if (buf.length >= 4 && buf.subarray(0, 4).toString("ascii") === "%PDF") return "application/pdf";
  return undefined;
}

async function fetchAuthed(resourceUrl: string): Promise<Response> {
  const token = await getAccessToken();
  const url = normalizeOneNoteResourceUrl(resourceUrl);
  const doFetch = () => fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  let res = await doFetch();
  // Retry on throttling/transient errors with Retry-After / exponential backoff,
  // mirroring graphFetch so media downloads survive rate-limit windows.
  for (let attempt = 0; attempt < 6 && (res.status === 429 || res.status === 503 || res.status === 504); attempt++) {
    const retryAfter = parseInt(res.headers.get("retry-after") ?? "0", 10);
    const delayMs = retryAfter > 0 ? retryAfter * 1000 : Math.min(2000 * 2 ** attempt, 60000);
    await new Promise((r) => setTimeout(r, delayMs));
    res = await doFetch();
  }
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Graph API ${res.status}: ${body}`);
  }
  return res;
}

export function isOneNoteResourceUrl(url: string): boolean {
  try {
    const parsed = new URL(url);
    return /(^|\.)graph\.microsoft\.com$/i.test(parsed.hostname)
      && /\/onenote\/resources\/[^/]+\/(?:\$value|content)$/i.test(parsed.pathname);
  } catch {
    return false;
  }
}

function normalizeOneNoteResourceUrl(resourceUrl: string): string {
  try {
    const parsed = new URL(resourceUrl);
    if (!/(^|\.)graph\.microsoft\.com$/i.test(parsed.hostname)) return resourceUrl;

    parsed.pathname = parsed.pathname
      .replace(
        /^\/v1\.0\/siteCollections\/([^/]+)\/onenote\/resources\/([^/]+)\/\$value$/i,
        "/v1.0/sites/$1/onenote/resources/$2/content"
      )
      .replace(
        /(\/v1\.0\/(?:me|users\/[^/]+|groups\/[^/]+|sites\/[^/]+)\/onenote\/resources\/[^/]+)\/\$value$/i,
        "$1/content"
      );

    parsed.search = "";
    return parsed.toString();
  } catch {
    return resourceUrl;
  }
}

export async function cacheOneNoteResource(
  resourceUrl: string,
  mediaTypeHint?: string,
  opts?: { assetDir?: string; linkBaseDir?: string }
): Promise<{ absolutePath: string; displayPath: string; mediaType: string }> {
  const assetDir = opts?.assetDir ?? READ_ASSET_DIR;
  const baseDir = opts?.linkBaseDir;
  await mkdir(assetDir, { recursive: true });

  const canonicalUrl = normalizeOneNoteResourceUrl(resourceUrl);
  const resourceId = getResourceId(canonicalUrl);
  const hash = createHash("sha1").update(canonicalUrl).digest("hex").slice(0, 10);
  const baseName = sanitizeStem(`${resourceId}-${hash}`) || hash;
  let mediaType = mediaTypeHint ?? "";
  let ext = extensionFromMediaType(mediaTypeHint);
  let absolutePath = join(assetDir, `${baseName}.${ext}`);

  try {
    await stat(absolutePath);
    return { absolutePath, displayPath: toDisplayPath(absolutePath, baseDir), mediaType };
  } catch {}

  const res = await fetchAuthed(resourceUrl);
  const buf = Buffer.from(await res.arrayBuffer());
  const headerMediaType = res.headers.get("content-type") || "";
  const sniffedMediaType = sniffMediaType(buf);
  mediaType = mediaType
    || (headerMediaType && headerMediaType !== "application/octet-stream" ? headerMediaType : "")
    || sniffedMediaType
    || headerMediaType
    || "application/octet-stream";
  ext = extensionFromMediaType(mediaType);
  absolutePath = join(assetDir, `${baseName}.${ext}`);

  try {
    await stat(absolutePath);
    return { absolutePath, displayPath: toDisplayPath(absolutePath, baseDir), mediaType };
  } catch {}

  await writeFile(absolutePath, buf);
  return { absolutePath, displayPath: toDisplayPath(absolutePath, baseDir), mediaType };
}

function normalizeResourceReferences(html: string): ResourceReference[] {
  const tags = [...html.matchAll(/<(img|object)\b[^>]*>/gi)].map((match) => match[0]);
  return tags.flatMap((tag) => {
    if (/^<img\b/i.test(tag)) {
      const url = extractAttr(tag, "data-fullres-src") || extractAttr(tag, "src");
      if (!url) return [];
      return [{
        alt: decodeHtmlEntities(extractAttr(tag, "alt")?.trim() || "image"),
        url,
        mediaType: extractAttr(tag, "data-fullres-src-type") || extractAttr(tag, "data-src-type"),
      }];
    }

    const url = extractAttr(tag, "data");
    if (!url) return [];
    return [{
      alt: decodeHtmlEntities(extractAttr(tag, "data-attachment")?.trim() || "attachment"),
      url,
      mediaType: extractAttr(tag, "type"),
    }];
  });
}

async function resolveResourceTargets(
  html: string,
  opts?: { assetDir?: string; linkBaseDir?: string }
) {
  const refs = normalizeResourceReferences(html);
  const byUrl = new Map<string, { displayPath: string; mediaType?: string }>();

  await Promise.all(
    refs.map(async (ref) => {
      if (byUrl.has(ref.url)) return;
      if (!isOneNoteResourceUrl(ref.url)) {
        byUrl.set(ref.url, { displayPath: ref.url, mediaType: ref.mediaType });
        return;
      }
      try {
        const cached = await cacheOneNoteResource(ref.url, ref.mediaType, opts);
        byUrl.set(ref.url, { displayPath: cached.displayPath, mediaType: cached.mediaType });
      } catch {
        byUrl.set(ref.url, { displayPath: ref.url, mediaType: ref.mediaType });
      }
    })
  );

  return byUrl;
}

function stripTagsInline(html: string): string {
  return decodeHtmlEntities(
    markStrikethrough(html)
      .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, "")
      .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, "")
      .replace(/<br\s*\/?>/gi, " ")
      .replace(/<\/(p|div|li|h[1-6])>/gi, " ")
      .replace(/<[^>]+>/g, "")
      .replace(/\s+/g, " ")
      .trim()
  );
}

function maskNestedTables(html: string): { masked: string; slots: string[] } {
  const slots: string[] = [];
  let out = "";
  let i = 0;
  while (i < html.length) {
    const start = html.toLowerCase().indexOf("<table", i);
    if (start === -1) { out += html.slice(i); break; }
    out += html.slice(i, start);
    let depth = 0;
    let j = start;
    while (j < html.length) {
      const lower = html.slice(j).toLowerCase();
      if (lower.startsWith("<table")) {
        depth++;
        const close = html.indexOf(">", j);
        if (close === -1) { j = html.length; break; }
        j = close + 1;
      } else if (lower.startsWith("</table>")) {
        depth--;
        j += 8;
        if (depth === 0) break;
      } else {
        j++;
      }
    }
    const idx = slots.length;
    slots.push(html.slice(start, j));
    out += `\uE000TABLE${idx}\uE000`;
    i = j;
  }
  return { masked: out, slots };
}

function parseRowCells(rowHtml: string): string[] {
  return [...rowHtml.matchAll(/<t[hd]\b[^>]*>([\s\S]*?)<\/t[hd]>/gi)].map((m) => m[1] ?? "");
}

function renderCell(cellHtml: string): string {
  const { masked, slots } = maskNestedTables(cellHtml);
  const flattened = slots.map((t) => {
    const stripped = t.replace(/^<table\b[^>]*>/i, "").replace(/<\/table>\s*$/i, "");
    const inner = maskNestedTables(stripped);
    const rows = [...inner.masked.matchAll(/<tr\b[^>]*>([\s\S]*?)<\/tr>/gi)].map((m) => m[1] ?? "");
    return rows.map((r) => parseRowCells(r).map((c) => renderCell(c.replace(/\uE000TABLE(\d+)\uE000/g, (_, n) => inner.slots[Number(n)] ?? ""))).filter(Boolean).join(" / ")).filter(Boolean).join(" ; ");
  });
  let text = stripTagsInline(masked);
  text = text.replace(/\uE000TABLE(\d+)\uE000/g, (_, n) => ` ${flattened[Number(n)] ?? ""} `);
  return text.replace(/\|/g, "\\|").replace(/\s+/g, " ").trim();
}

function renderTable(tableHtml: string): string {
  const stripped = tableHtml.replace(/^<table\b[^>]*>/i, "").replace(/<\/table>\s*$/i, "");
  const { masked, slots } = maskNestedTables(stripped);
  const rows = [...masked.matchAll(/<tr\b[^>]*>([\s\S]*?)<\/tr>/gi)].map((m) => {
    const rowMasked = m[1] ?? "";
    return parseRowCells(rowMasked).map((c) => renderCell(c.replace(/\uE000TABLE(\d+)\uE000/g, (_, n) => slots[Number(n)] ?? "")));
  });
  if (rows.length === 0) return "";
  const width = Math.max(...rows.map((r) => r.length));
  const padded = rows.map((r) => [...r, ...Array(width - r.length).fill("")]);
  const header = padded[0]!.map((c) => c || " ");
  const body = padded.slice(1);
  const lines = [
    `| ${header.join(" | ")} |`,
    `| ${header.map(() => "---").join(" | ")} |`,
    ...body.map((r) => `| ${r.map((c) => c || " ").join(" | ")} |`),
  ];
  return `\n\n${lines.join("\n")}\n\n`;
}

function replaceTopLevelTables(html: string): string {
  let out = "";
  let i = 0;
  while (i < html.length) {
    const start = html.toLowerCase().indexOf("<table", i);
    if (start === -1) {
      out += html.slice(i);
      break;
    }
    out += html.slice(i, start);
    let depth = 0;
    let j = start;
    while (j < html.length) {
      const lower = html.slice(j).toLowerCase();
      if (lower.startsWith("<table")) {
        depth++;
        const close = html.indexOf(">", j);
        if (close === -1) { j = html.length; break; }
        j = close + 1;
      } else if (lower.startsWith("</table>")) {
        depth--;
        j += 8;
        if (depth === 0) break;
      } else {
        j++;
      }
    }
    out += renderTable(html.slice(start, j));
    i = j;
  }
  return out;
}

export async function renderHtmlForRead(
  html: string,
  options?: RenderOptions
): Promise<string> {
  const replacements = options?.downloadAssets === false
    ? new Map<string, { displayPath: string; mediaType?: string }>()
    : await resolveResourceTargets(html, options);

  // Convert <img>/<object> to markdown BEFORE table rendering so assets inside
  // table cells survive stripTagsInline() during cell flattening.
  let rendered = html.replace(/<img\b[^>]*>/gi, (tag) => {
    const url = extractAttr(tag, "data-fullres-src") || extractAttr(tag, "src");
    if (!url) return "";
    const alt = escapeMarkdownText(decodeHtmlEntities(extractAttr(tag, "alt")?.trim() || "image"));
    const target = replacements.get(url)?.displayPath || url;
    return `\n\n![${alt}](${target})\n\n`;
  });

  rendered = rendered.replace(/<object\b[^>]*>/gi, (tag) => {
    const url = extractAttr(tag, "data");
    if (!url) return "";
    const label = escapeMarkdownText(
      decodeHtmlEntities(extractAttr(tag, "data-attachment")?.trim() || "attachment")
    );
    const target = replacements.get(url)?.displayPath || url;
    return `\n\n[${label}](${target})\n\n`;
  });

  // Convert <a href> links to markdown. Mask each as a placeholder so its URL
  // (which may contain parens/spaces and thus need <angle-bracket> form) is not
  // eaten by the final tag-strip below; restored at the very end.
  const linkMark = String.fromCharCode(0xe001);
  const linkSlots: { text: string; href: string }[] = [];
  rendered = rendered.replace(/<a\b[^>]*>[\s\S]*?<\/a>/gi, (whole) => {
    const open = whole.match(/<a\b[^>]*>/i)?.[0] ?? "";
    const href = extractAttr(open, "href");
    const inner = whole.replace(/^<a\b[^>]*>/i, "").replace(/<\/a>\s*$/i, "");
    const text = decodeHtmlEntities(inner.replace(/<[^>]+>/g, "")).replace(/\s+/g, " ").trim();
    if (!href) return text;
    const idx = linkSlots.length;
    linkSlots.push({ text, href: decodeHtmlEntities(href).trim() });
    return `${linkMark}${idx}${linkMark}`;
  });

  rendered = replaceTopLevelTables(rendered);

  const text = decodeHtmlEntities(
    markStrikethrough(rendered)
      // Drop document <head> (title/meta) so the page title isn't re-emitted as body text.
      .replace(/<head\b[^>]*>[\s\S]*?<\/head>/gi, "")
      .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, "")
      .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, "")
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<\/p>/gi, "\n\n")
      .replace(/<\/div>/gi, "\n")
      .replace(/<\/h[1-6]>/gi, "\n\n")
      .replace(/<\/li>/gi, "\n")
      .replace(/<li[^>]*>/gi, "- ")
      .replace(/<[^>]+>/g, "")
  );

  // OneNote's HTML is pretty-printed with tab indentation; once tags are gone
  // that inter-tag whitespace survives as leading tabs and blank-but-not-empty
  // lines. Trim each line and collapse blank runs to keep the output readable.
  const collapsed = text
    .split("\n")
    .map((line) => line.replace(/^[ \t]+/, "").replace(/\s+$/, ""))
    .join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();

  // Resolve each link to its best current URL (may hit the sync cache / Graph),
  // then restore as markdown. URLs with parens/whitespace use the <angle-bracket>
  // form so they don't terminate the link early.
  const urls = await Promise.all(linkSlots.map((s) => resolveLinkHref(s.href)));
  const linkRe = new RegExp(`${linkMark}(\\d+)${linkMark}`, "g");
  return collapsed.replace(linkRe, (_, n: string) => {
    const slot = linkSlots[Number(n)];
    if (!slot) return "";
    const url = urls[Number(n)]!;
    const label = escapeMarkdownText(slot.text || url);
    const safe = /[()\s]/.test(url) ? `<${url}>` : url;
    return `[${label}](${safe})`;
  });
}

/**
 * Rewrite a page's HTML so that OneNote resource references (img/object) point
 * at locally-downloaded copies. Returns the (still complete) HTML document with
 * remote resource URLs replaced by relative paths, for `onenote export *.html`.
 */
export async function renderHtmlForExport(
  html: string,
  opts?: { assetDir?: string; linkBaseDir?: string }
): Promise<string> {
  const replacements = await resolveResourceTargets(html, opts);
  const localFor = (u?: string) => (u ? replacements.get(u)?.displayPath : undefined);
  const swapAttr = (tag: string, from: string | undefined, to: string): string => {
    if (!from) return tag;
    return tag.replace(`"${from}"`, `"${to}"`).replace(`'${from}'`, `'${to}'`);
  };

  let out = html.replace(/<img\b[^>]*>/gi, (tag) => {
    const full = extractAttr(tag, "data-fullres-src");
    const src = extractAttr(tag, "src");
    const target = localFor(full ?? src);
    if (!target) return tag;
    let t = tag;
    if (src) t = swapAttr(t, src, target);
    else t = t.replace(/<img\b/i, `<img src="${target}"`);
    if (full) t = swapAttr(t, full, target);
    return t;
  });

  out = out.replace(/<object\b[^>]*>/gi, (tag) => {
    const data = extractAttr(tag, "data");
    const target = localFor(data);
    return target ? swapAttr(tag, data, target) : tag;
  });

  // Rewrite <a href> links to their current https URL (same policy as markdown:
  // sync cache → notebook re-resolution → embedded base-path).
  const anchorHrefs = new Set<string>();
  for (const m of out.matchAll(/<a\b[^>]*>/gi)) {
    const raw = extractAttr(m[0], "href");
    if (raw) anchorHrefs.add(raw);
  }
  const resolvedByRaw = new Map<string, string>();
  await Promise.all(
    [...anchorHrefs].map(async (raw) => {
      resolvedByRaw.set(raw, await resolveLinkHref(decodeHtmlEntities(raw).trim()));
    })
  );
  const escapeHtmlAttr = (v: string) =>
    v.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  out = out.replace(/<a\b[^>]*>/gi, (tag) => {
    const raw = extractAttr(tag, "href");
    if (!raw) return tag;
    const resolved = resolvedByRaw.get(raw);
    return resolved ? swapAttr(tag, raw, escapeHtmlAttr(resolved)) : tag;
  });

  return out;
}

export async function renderResourceForRead(resourceUrl: string): Promise<{
  title: string;
  content: string;
  assetPath: string;
}> {
  const cached = await cacheOneNoteResource(resourceUrl);
  const title = cached.absolutePath.split(sep).at(-1) ?? "resource";
  return {
    title,
    content: `![${escapeMarkdownText(title)}](${cached.displayPath})`,
    assetPath: cached.displayPath,
  };
}
