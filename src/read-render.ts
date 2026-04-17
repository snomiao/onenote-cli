import { createHash } from "node:crypto";
import { mkdir, stat, writeFile } from "node:fs/promises";
import { join, relative, sep } from "node:path";
import { getAccessToken } from "./auth";

const READ_ASSET_DIR = process.env.ONENOTE_READ_ASSET_DIR
  || join(process.cwd(), ".onenote", "assets");

type ResourceReference = {
  alt: string;
  url: string;
  mediaType?: string;
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

function extractAttr(tag: string, name: string): string | undefined {
  const match = tag.match(new RegExp(`${name}=(["'])(.*?)\\1`, "i"));
  return match?.[2];
}

function escapeMarkdownText(text: string): string {
  return text.replace(/[[\]\\]/g, "\\$&");
}

function sanitizeStem(text: string): string {
  return text
    .replace(/[^a-zA-Z0-9._-]+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "")
    .slice(0, 120);
}

function toDisplayPath(path: string): string {
  const rel = relative(process.cwd(), path).split(sep).join("/");
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
  const res = await fetch(normalizeOneNoteResourceUrl(resourceUrl), {
    headers: { Authorization: `Bearer ${token}` },
  });
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
  mediaTypeHint?: string
): Promise<{ absolutePath: string; displayPath: string; mediaType: string }> {
  await mkdir(READ_ASSET_DIR, { recursive: true });

  const canonicalUrl = normalizeOneNoteResourceUrl(resourceUrl);
  const resourceId = getResourceId(canonicalUrl);
  const hash = createHash("sha1").update(canonicalUrl).digest("hex").slice(0, 10);
  const baseName = sanitizeStem(`${resourceId}-${hash}`) || hash;
  let mediaType = mediaTypeHint ?? "";
  let ext = extensionFromMediaType(mediaTypeHint);
  let absolutePath = join(READ_ASSET_DIR, `${baseName}.${ext}`);

  try {
    await stat(absolutePath);
    return { absolutePath, displayPath: toDisplayPath(absolutePath), mediaType };
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
  absolutePath = join(READ_ASSET_DIR, `${baseName}.${ext}`);

  try {
    await stat(absolutePath);
    return { absolutePath, displayPath: toDisplayPath(absolutePath), mediaType };
  } catch {}

  await writeFile(absolutePath, buf);
  return { absolutePath, displayPath: toDisplayPath(absolutePath), mediaType };
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

async function resolveResourceTargets(html: string) {
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
        const cached = await cacheOneNoteResource(ref.url, ref.mediaType);
        byUrl.set(ref.url, { displayPath: cached.displayPath, mediaType: cached.mediaType });
      } catch {
        byUrl.set(ref.url, { displayPath: ref.url, mediaType: ref.mediaType });
      }
    })
  );

  return byUrl;
}

export async function renderHtmlForRead(
  html: string,
  options?: { downloadAssets?: boolean }
): Promise<string> {
  let rendered = html;
  const replacements = options?.downloadAssets === false
    ? new Map<string, { displayPath: string; mediaType?: string }>()
    : await resolveResourceTargets(html);

  rendered = rendered.replace(/<img\b[^>]*>/gi, (tag) => {
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

  return decodeHtmlEntities(
    rendered
      .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, "")
      .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, "")
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<\/p>/gi, "\n\n")
      .replace(/<\/div>/gi, "\n")
      .replace(/<\/h[1-6]>/gi, "\n\n")
      .replace(/<\/li>/gi, "\n")
      .replace(/<li[^>]*>/gi, "- ")
      .replace(/<[^>]+>/g, "")
      .replace(/\n{3,}/g, "\n\n")
      .trim()
  );
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
