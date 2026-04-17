import { mkdir, readFile, writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { getAccessToken } from "./auth";
import {
  isOneNoteResourceUrl,
  renderHtmlForRead,
  renderResourceForRead,
} from "./read-render";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

const PKG_ROOT = dirname(import.meta.dir);
const PKG_CACHE_DIR = join(PKG_ROOT, ".onenote");
const NOTEBOOK_CACHE_PATH = join(PKG_CACHE_DIR, "notebooks.json");
const NOTEBOOK_CACHE_TTL_MS = 24 * 60 * 60 * 1000;

type CachedNotebook = {
  id: string;
  displayName: string;
  webUrl?: string;
  lastModifiedDateTime?: string;
};

async function loadNotebookCache(): Promise<CachedNotebook[] | null> {
  try {
    const raw = await readFile(NOTEBOOK_CACHE_PATH, "utf8");
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object") return null;
    const cachedAt = typeof parsed.cachedAt === "string" ? Date.parse(parsed.cachedAt) : NaN;
    if (!Number.isFinite(cachedAt) || Date.now() - cachedAt > NOTEBOOK_CACHE_TTL_MS) return null;
    if (!Array.isArray(parsed.notebooks)) return null;
    return parsed.notebooks as CachedNotebook[];
  } catch {
    return null;
  }
}

async function saveNotebookCache(notebooks: CachedNotebook[]): Promise<void> {
  await mkdir(PKG_CACHE_DIR, { recursive: true });
  await writeFile(
    NOTEBOOK_CACHE_PATH,
    JSON.stringify({ cachedAt: new Date().toISOString(), notebooks }, null, 2)
  );
}

async function lookupSectionFromSyncCache(
  notebookName: string,
  sectionName: string
): Promise<{ id: string; displayName: string; links?: any } | null> {
  try {
    const file = join(PKG_CACHE_DIR, "cache", notebookName, `${sectionName}.json`);
    const raw = await readFile(file, "utf8");
    const data = JSON.parse(raw);
    const url: string = data.webUrl ?? "";
    const m = decodeURIComponent(url).match(/sourcedoc=\{([0-9a-f-]+)\}/i);
    if (!m) return null;
    return { id: `0-${m[1]!.toLowerCase()}`, displayName: data.section ?? sectionName };
  } catch {
    return null;
  }
}

async function lookupPageFromSyncCache(
  notebookName: string,
  sectionName: string,
  pageTitle: string
): Promise<{ id: string; title: string; links?: any } | null> {
  try {
    const file = join(PKG_CACHE_DIR, "cache", notebookName, `${sectionName}.json`);
    const raw = await readFile(file, "utf8");
    const data = JSON.parse(raw);
    const pages = (data.officialPages ?? []) as Array<{ guid: string; title: string; webUrl?: string }>;
    const matches = pages.filter((p) => p.title === pageTitle);
    if (matches.length === 0) return null;
    if (matches.length > 1) {
      const lines = matches.map((p) => `  - ${p.guid} ${p.webUrl ?? ""}`).join("\n");
      throw new Error(
        `Multiple pages titled '${pageTitle}' in '${notebookName}/${sectionName}' (${matches.length} matches). Please rename for uniqueness or use URL/ID.\nCandidates:\n${lines}`
      );
    }
    return { id: `0-${matches[0]!.guid.toLowerCase()}`, title: matches[0]!.title };
  } catch (err: any) {
    if (err?.message?.startsWith("Multiple pages")) throw err;
    return null;
  }
}

async function getCachedNotebooks(force = false): Promise<CachedNotebook[]> {
  if (!force) {
    const cached = await loadNotebookCache();
    if (cached) return cached;
  }
  const fresh = await listNotebooks();
  const slim: CachedNotebook[] = fresh.map((nb: any) => ({
    id: nb.id,
    displayName: nb.displayName,
    webUrl: nb.links?.oneNoteWebUrl?.href ?? nb.links?.oneNoteClientUrl?.href,
    lastModifiedDateTime: nb.lastModifiedDateTime,
  }));
  await saveNotebookCache(slim);
  return slim;
}

interface GraphResponse<T> {
  value: T[];
  "@odata.nextLink"?: string;
}

async function graphFetch(path: string, init?: RequestInit): Promise<Response> {
  const token = await getAccessToken();
  const url = path.startsWith("http") ? path : `${GRAPH_BASE}${path}`;
  const res = await fetch(url, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...init?.headers,
    },
  });
  if (!res.ok) {
    const body = await res.text();
    let code = "";
    let message = "";
    try {
      const err = JSON.parse(body);
      code = err.error?.code ?? "";
      message = err.error?.message ?? body;
    } catch {
      message = body;
    }
    const error = new Error(`Graph API ${res.status}: ${message}`) as any;
    error.statusCode = res.status;
    error.graphCode = code;
    throw error;
  }
  return res;
}

function is5000LimitError(err: any): boolean {
  return err?.graphCode === "10008" || err?.graphCode === "20102";
}

// --- Notebook path helpers ---

function getNotebookDrivePath(notebook: any): string | null {
  const webUrl = notebook.links?.oneNoteWebUrl?.href;
  if (!webUrl) return null;
  const match = decodeURIComponent(new URL(webUrl).pathname).match(/Documents\/(.+)/);
  return match?.[1] ?? null;
}

async function listNotebookFilesViaDrive(notebook: any): Promise<any[]> {
  const drivePath = getNotebookDrivePath(notebook);
  if (!drivePath) return [];
  const encoded = drivePath.split("/").map(s => encodeURIComponent(s)).join("/");
  const res = await graphFetch(
    `/me/drive/root:/${encoded}:/children?$select=name,id,file,folder,size,lastModifiedDateTime,createdDateTime`
  );
  const data: GraphResponse<any> = await res.json();
  return data.value;
}

// --- Notebooks ---

export async function listNotebooks() {
  const res = await graphFetch("/me/onenote/notebooks");
  const data: GraphResponse<any> = await res.json();
  return data.value;
}

export async function getNotebook(ref: string) {
  const { apiBase, notebookId } = await resolveNotebookRef(ref);
  const res = await graphFetch(`${apiBase}/onenote/notebooks/${notebookId}`);
  return res.json();
}

async function getNotebookByRef(ref: string) {
  const { apiBase, notebookId } = await resolveNotebookRef(ref);
  const res = await graphFetch(`${apiBase}/onenote/notebooks/${notebookId}`);
  return res.json();
}

export async function createNotebook(displayName: string) {
  const res = await graphFetch("/me/onenote/notebooks", {
    method: "POST",
    body: JSON.stringify({ displayName }),
  });
  return res.json();
}

// --- Sections ---

export async function listSections(notebookId?: string) {
  try {
    let path = "/me/onenote/sections";
    if (notebookId) {
      const { apiBase, notebookId: resolvedNotebookId } = await resolveNotebookRef(notebookId);
      path = `${apiBase}/onenote/notebooks/${resolvedNotebookId}/sections`;
    }
    const res = await graphFetch(path);
    const data: GraphResponse<any> = await res.json();
    return data.value;
  } catch (err: any) {
    if (!is5000LimitError(err) || !notebookId) throw err;
    // Fallback: list .one files via OneDrive
    const notebook = await getNotebookByRef(notebookId);
    const files = await listNotebookFilesViaDrive(notebook);
    return files
      .filter((f: any) => f.name?.endsWith(".one"))
      .map((f: any) => ({
        id: f.id,
        displayName: f.name.replace(/\.one$/, ""),
        createdDateTime: f.createdDateTime,
        lastModifiedDateTime: f.lastModifiedDateTime,
        _source: "onedrive",
      }));
  }
}

export async function getSection(ref: string) {
  const { apiBase, sectionId } = await resolveSectionRef(ref);
  const res = await graphFetch(`${apiBase}/onenote/sections/${sectionId}`);
  return res.json();
}

export async function createSection(notebookId: string, displayName: string) {
  const { apiBase, notebookId: resolvedNotebookId } = await resolveNotebookRef(notebookId);
  const res = await graphFetch(`${apiBase}/onenote/notebooks/${resolvedNotebookId}/sections`, {
    method: "POST",
    body: JSON.stringify({ displayName }),
  });
  return res.json();
}

// --- Section Groups ---

export async function listSectionGroups(notebookId?: string) {
  try {
    let path = "/me/onenote/sectionGroups";
    if (notebookId) {
      const { apiBase, notebookId: resolvedNotebookId } = await resolveNotebookRef(notebookId);
      path = `${apiBase}/onenote/notebooks/${resolvedNotebookId}/sectionGroups`;
    }
    const res = await graphFetch(path);
    const data: GraphResponse<any> = await res.json();
    return data.value;
  } catch (err: any) {
    if (!is5000LimitError(err) || !notebookId) throw err;
    // Fallback: list folders via OneDrive
    const notebook = await getNotebookByRef(notebookId);
    const files = await listNotebookFilesViaDrive(notebook);
    return files
      .filter((f: any) => f.folder && !f.name?.startsWith("OneNote_") && f.name !== "deletePending")
      .map((f: any) => ({
        id: f.id,
        displayName: f.name,
        createdDateTime: f.createdDateTime,
        _source: "onedrive",
      }));
  }
}

// --- Pages ---

export async function listPages(sectionId?: string) {
  let path = "/me/onenote/pages";
  if (sectionId) {
    const { apiBase, sectionId: resolvedSectionId } = await resolveSectionRef(sectionId);
    path = `${apiBase}/onenote/sections/${resolvedSectionId}/pages`;
  }
  const res = await graphFetch(path);
  const data: GraphResponse<any> = await res.json();
  return data.value;
}

export async function getPage(id: string) {
  const { apiBase, pageId } = await resolvePageRef(id);
  const res = await graphFetch(`${apiBase}/onenote/pages/${pageId}`);
  return res.json();
}

export async function getPageContent(id: string): Promise<string> {
  const { apiBase, pageId } = await resolvePageRef(id);
  const res = await graphFetch(`${apiBase}/onenote/pages/${pageId}/content`);
  return res.text();
}

export type PageUpdateCommand = {
  target: string; // "body", "title", or "#element-id"
  action: "append" | "insert" | "prepend" | "replace";
  position?: "before" | "after";
  content: string;
};

/**
 * Update page content via PATCH. Commands follow MS Graph OneNote page update format.
 * Accepts either a page ID or a OneNote URL (SharePoint URLs use site-scoped
 * endpoints to bypass the 5000-item limit on /me/onenote/*).
 * @see https://learn.microsoft.com/graph/onenote-update-page
 */
export async function updatePage(urlOrId: string, commands: PageUpdateCommand[]) {
  const { apiBase, pageId } = await resolvePageRef(urlOrId);
  const res = await graphFetch(`${apiBase}/onenote/pages/${pageId}/content`, {
    method: "PATCH",
    body: JSON.stringify(commands),
  });
  return res.ok;
}

/**
 * Rename a page by replacing its title element.
 */
export async function renamePage(urlOrId: string, newTitle: string) {
  return updatePage(urlOrId, [
    { target: "title", action: "replace", content: newTitle },
  ]);
}

/**
 * Append HTML content to a page's body.
 */
export async function appendToPage(urlOrId: string, htmlContent: string) {
  return updatePage(urlOrId, [
    { target: "body", action: "append", content: htmlContent },
  ]);
}

/**
 * Rename a section.
 */
export async function renameSection(ref: string, newName: string) {
  const { apiBase, sectionId } = await resolveSectionRef(ref);
  const res = await graphFetch(`${apiBase}/onenote/sections/${sectionId}`, {
    method: "PATCH",
    body: JSON.stringify({ displayName: newName }),
  });
  return res.json();
}

/**
 * Rename a notebook.
 */
export async function renameNotebook(ref: string, newName: string) {
  const { apiBase, notebookId } = await resolveNotebookRef(ref);
  const res = await graphFetch(`${apiBase}/onenote/notebooks/${notebookId}`, {
    method: "PATCH",
    body: JSON.stringify({ displayName: newName }),
  });
  return res.json();
}

export async function createPage(sectionId: string, title: string, htmlBody: string) {
  const { apiBase, sectionId: resolvedSectionId } = await resolveSectionRef(sectionId);
  const html = `
<!DOCTYPE html>
<html>
  <head>
    <title>${title}</title>
  </head>
  <body>
    ${htmlBody}
  </body>
</html>`;

  const res = await graphFetch(`${apiBase}/onenote/sections/${resolvedSectionId}/pages`, {
    method: "POST",
    headers: { "Content-Type": "text/html" },
    body: html,
  });
  return res.json();
}

export async function deletePage(urlOrId: string) {
  const { apiBase, pageId } = await resolvePageRef(urlOrId);
  await graphFetch(`${apiBase}/onenote/pages/${pageId}`, { method: "DELETE" });
}

// --- Search ---

async function resolveOneNoteUrl(filePath: string): Promise<string> {
  // Get the driveItem to obtain the proper Doc.aspx URL with sourcedoc GUID
  try {
    const match = decodeURIComponent(new URL(filePath).pathname).match(/Documents\/(.+)/);
    if (!match) return filePath;
    const encoded = match[1].split("/").map((s) => encodeURIComponent(s)).join("/");
    const res = await graphFetch(`/me/drive/root:/${encoded}?$select=webUrl`);
    const item = (await res.json()) as any;
    // webUrl has format: ...Doc.aspx?sourcedoc={GUID}&file=name.one&action=edit...
    // Clean it up
    return item.webUrl?.split("&mobileredirect")[0] ?? filePath;
  } catch {
    return filePath;
  }
}

export async function searchPages(query: string) {
  // Use Microsoft Graph Search API with listItem to search OneNote content (full-text)
  const res = await graphFetch("/search/query", {
    method: "POST",
    body: JSON.stringify({
      requests: [
        {
          entityTypes: ["listItem"],
          query: { queryString: `${query} FileExtension:one` },
          from: 0,
          size: 25,
          fields: ["title", "path", "parentLink", "spWebUrl"],
        },
      ],
    }),
  });
  const data = (await res.json()) as any;
  const hits = data.value?.[0]?.hitsContainers?.[0]?.hits ?? [];
  const results = hits
    .map((hit: any) => {
      const r = hit.resource ?? {};
      const fields = r.fields ?? {};
      const rawSummary = (hit.summary ?? "")
        .replace(/<c0>/g, "**")
        .replace(/<\/c0>/g, "**")
        .replace(/<ddd\/>/g, "...")
        .replace(/<[^>]*>/g, "")
        .trim();
      const filePath = fields.path ?? r.webUrl ?? "";
      const pathMatch = filePath.match(/Documents\/(?:Notebooks\/)?([^/]+)\//);
      const notebook = pathMatch ? decodeURIComponent(pathMatch[1]) : "";
      return {
        id: r.id ?? hit.hitId,
        section: fields.title ?? (r.name ?? "").replace(/\.one$/, ""),
        notebook,
        summary: rawSummary,
        filePath,
        url: "",
        lastModifiedDateTime: r.lastModifiedDateTime ?? "",
      };
    })
    .filter((r: any) => r.summary.length > 0);

  // Resolve proper OneNote Online URLs in parallel
  await Promise.all(
    results.map(async (r: any) => {
      r.url = await resolveOneNoteUrl(r.filePath);
    })
  );
  return results;
}

// --- Read page by URL ---

/**
 * Parse a OneNote Online URL to determine type and extract GUIDs.
 */
interface ParsedOneNoteUrl {
  type: "page" | "section" | "notebook" | "unknown";
  sectionGuid?: string;
  pageGuid?: string;
  notebookId?: string;
  fileName?: string;       // from file= param or .one path basename
  // SharePoint / business notebook resolution hints
  siteRef?: string;        // e.g. "host:/personal/user" for /sites endpoint lookup
  notebookName?: string;   // from direct notebook path basename
  sectionName?: string;    // from wd=target first segment (with .one stripped)
  pageTitle?: string;      // from wd=target second segment (unescaped)
}

function unescapeOneNoteName(s: string): string {
  return s.replace(/\\(.)/g, "$1");
}

function parseOneNoteUrl(url: string): ParsedOneNoteUrl {
  const decoded = decodeURIComponent(url);
  const guids = decoded.match(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/gi) ?? [];
  let fileName: string | undefined;

  const sourcedocMatch = decoded.match(/sourcedoc=\{?([0-9a-f-]+)\}?/i);
  let sectionGuid = sourcedocMatch?.[1]?.toLowerCase();

  // SharePoint siteRef from the URL host + /personal/{user} or /sites/{name} segment
  let siteRef: string | undefined;
  let notebookName: string | undefined;
  try {
    const u = new URL(url);
    if (/sharepoint\.com$/i.test(u.hostname)) {
      const p = decodeURIComponent(u.pathname);
      const personal = p.match(/\/(personal\/[^/]+)/i);
      const site = p.match(/\/(sites|teams)\/([^/]+)/i);
      if (personal) siteRef = `${u.hostname}:/${personal[1]}:`;
      else if (site) siteRef = `${u.hostname}:/${site[1]}/${site[2]}:`;

      const lastSegment = p.split("/").filter(Boolean).at(-1);
      if (lastSegment && !/\.aspx$/i.test(lastSegment)) {
        if (/\.one$/i.test(lastSegment)) fileName = lastSegment.replace(/\.one$/i, "");
        else notebookName = lastSegment;
      }
    }
  } catch {}

  // Doc.aspx URLs expose the current .one file as the `file=` query parameter.
  const fileMatch = decoded.match(/[?&]file=([^&]+)/i);
  if (fileMatch) {
    fileName = decodeURIComponent(fileMatch[1]).replace(/\.one$/i, "");
  }

  // wd=target(sectionName|SECTION_GUID/pageTitle|PAGE_GUID/)
  // Page titles may contain escaped ')' as '\\)' — skip those when finding the terminator.
  const wdTargetMatch = decoded.match(/wd=target\(((?:[^)\\]|\\.)*)\)/i);
  let pageGuid: string | undefined;
  let sectionName: string | undefined;
  let pageTitle: string | undefined;
  if (wdTargetMatch) {
    const target = wdTargetMatch[1];
    // Capture "name|guid" pairs, where name may contain backslash escapes.
    const pairs = [...target.matchAll(/((?:[^|\\]|\\.)+)\|([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})/gi)];
    if (pairs.length >= 1) {
      sectionName = unescapeOneNoteName(pairs[0][1]).replace(/\.one$/i, "");
      sectionGuid = pairs[0][2].toLowerCase();
    }
    if (pairs.length >= 2) {
      // Second pair's name begins with "/" (segment separator) — strip it.
      pageTitle = unescapeOneNoteName(pairs[1][1].replace(/^\//, ""));
      pageGuid = pairs[1][2].toLowerCase();
    } else if (pairs.length === 0 && guids.length > 1) {
      pageGuid = guids[guids.length - 1].toLowerCase();
    }
  }

  if (!sectionName && fileName && sectionGuid) {
    sectionName = fileName;
  }

  const nbMatch = decoded.match(/notebooks\/(1-[0-9a-f-]+)/i);
  const notebookId = nbMatch?.[1];

  const base = { siteRef, fileName, notebookName, sectionName, pageTitle };
  if (pageGuid && sectionGuid) return { type: "page", sectionGuid, pageGuid, ...base };
  if (sectionGuid) return { type: "section", sectionGuid, ...base };
  if (notebookId || notebookName) return { type: "notebook", notebookId, ...base };
  return { type: "unknown", ...base };
}

// --- Site-scoped resolution (bypasses 5000-item limit on /me/onenote/*) ---

const siteIdCache = new Map<string, string>();

async function resolveSiteId(siteRef: string): Promise<string> {
  if (siteIdCache.has(siteRef)) return siteIdCache.get(siteRef)!;
  const res = await graphFetch(`/sites/${siteRef}?$select=id`);
  const site = (await res.json()) as any;
  siteIdCache.set(siteRef, site.id);
  return site.id;
}

function odataEscape(s: string): string {
  return s.replace(/'/g, "''");
}

function extractSectionIdFromClientUrl(href: string | undefined, wantGuid?: string): boolean {
  if (!href || !wantGuid) return false;
  const m = decodeURIComponent(href).match(/section-id=([0-9a-f-]+)/i);
  return (m?.[1]?.toLowerCase() ?? "") === wantGuid;
}

function extractPageIdFromClientUrl(href: string | undefined, wantGuid?: string): boolean {
  if (!href || !wantGuid) return false;
  const m = decodeURIComponent(href).match(/page-id=([0-9a-f-]+)/i);
  return (m?.[1]?.toLowerCase() ?? "") === wantGuid;
}

async function resolveSharePointNotebook(
  siteId: string,
  notebookName: string
): Promise<{ id: string; displayName: string } | null> {
  const res = await graphFetch(
    `/sites/${siteId}/onenote/notebooks?$filter=displayName eq '${odataEscape(notebookName)}'&$select=id,displayName&$top=50`
  );
  const data: GraphResponse<any> = await res.json();
  if (!data.value?.length) return null;
  if (data.value.length > 1) {
    const lines = data.value.map((n: any) => `  - ${n.id}`).join("\n");
    throw new Error(
      `Multiple notebooks named '${notebookName}' in site (${data.value.length} matches). Please rename for uniqueness or pass a Graph ID.\nCandidates:\n${lines}`
    );
  }
  return {
    id: data.value[0].id,
    displayName: data.value[0].displayName ?? notebookName,
  };
}

async function resolveSharePointSection(
  siteId: string,
  sectionName: string,
  sectionGuid?: string
): Promise<{ id: string; displayName: string } | null> {
  const res = await graphFetch(
    `/sites/${siteId}/onenote/sections?$filter=displayName eq '${odataEscape(sectionName)}'&$select=id,displayName,links&$top=50`
  );
  const data: GraphResponse<any> = await res.json();
  if (!data.value?.length) return null;
  if (data.value.length === 1 || !sectionGuid) return data.value[0];
  // Multiple matches — disambiguate by section-id in oneNoteClientUrl
  const match = data.value.find((s: any) =>
    extractSectionIdFromClientUrl(s.links?.oneNoteClientUrl?.href, sectionGuid)
  );
  return match ?? data.value[0];
}

async function resolveSharePointPage(
  siteId: string,
  sectionId: string,
  pageTitle: string,
  pageGuid?: string
): Promise<{ id: string; title: string; webUrl: string } | null> {
  const res = await graphFetch(
    `/sites/${siteId}/onenote/sections/${sectionId}/pages?$filter=title eq '${odataEscape(pageTitle)}'&$select=id,title,links&$top=50`
  );
  const data: GraphResponse<any> = await res.json();
  const mapOne = (p: any) => ({
    id: p.id,
    title: p.title ?? "(untitled)",
    webUrl: p.links?.oneNoteWebUrl?.href ?? "",
  });
  if (!data.value?.length) return null;
  if (data.value.length === 1 || !pageGuid) return mapOne(data.value[0]);
  const match = data.value.find((p: any) =>
    extractPageIdFromClientUrl(p.links?.oneNoteClientUrl?.href, pageGuid)
  );
  return mapOne(match ?? data.value[0]);
}

function looksLikeNotebookId(s: string): boolean {
  return /^[0-9]-[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(s);
}

function looksLikeGraphId(s: string): boolean {
  return /^[0-9]-[0-9a-f-]{10,}$/i.test(s);
}

export type ResolvedPath = {
  apiBase: string;
  notebookId?: string;
  sectionId?: string;
  pageId?: string;
  notebookName?: string;
  sectionName?: string;
  pageTitle?: string;
};

function uniqueOrThrow<T extends { id: string; displayName?: string; title?: string; links?: any }>(
  items: T[],
  kind: "notebook" | "section" | "page",
  name: string,
  parent?: string
): T {
  if (items.length === 0) {
    throw new Error(`${kind} '${name}'${parent ? ` in '${parent}'` : ""} not found.`);
  }
  if (items.length > 1) {
    const ctx = parent ? ` in '${parent}'` : "";
    const lines = items.map((i) => {
      const label = (i as any).displayName ?? (i as any).title ?? "(untitled)";
      const url = i.links?.oneNoteWebUrl?.href ?? i.links?.oneNoteClientUrl?.href ?? "";
      return `  - ${i.id}${url ? ` ${url}` : ""} [${label}]`;
    });
    throw new Error(
      `Multiple ${kind}s named '${name}'${ctx} (${items.length} matches). Please rename for uniqueness or use URL/ID.\nCandidates:\n${lines.join("\n")}`
    );
  }
  return items[0]!;
}

async function resolveByPath(path: string): Promise<ResolvedPath> {
  const segments = path.split("/").filter((s) => s.length > 0);
  if (segments.length === 0) throw new Error("Empty path.");
  if (segments.length > 3) throw new Error(`Path too deep (max 3 segments): ${path}`);

  // Segment 1: notebook
  const nbName = segments[0]!;
  const notebooks = (await getCachedNotebooks()).filter((n) => n.displayName === nbName);
  if (notebooks.length === 0) {
    const fresh = (await getCachedNotebooks(true)).filter((n) => n.displayName === nbName);
    if (fresh.length === 0) {
      throw new Error(`Notebook '${nbName}' not found. Run 'onenote notebooks list' to see available names.`);
    }
    notebooks.push(...fresh);
  }
  const nb = uniqueOrThrow(
    notebooks.map((n) => ({ id: n.id, displayName: n.displayName, links: { oneNoteWebUrl: { href: n.webUrl } } })),
    "notebook",
    nbName
  );
  const nbRef = nb.links.oneNoteWebUrl.href
    ? await resolveNotebookRef(nb.links.oneNoteWebUrl.href)
    : { apiBase: "/me", notebookId: nb.id, displayName: nb.displayName };

  if (segments.length === 1) {
    return { apiBase: nbRef.apiBase, notebookId: nbRef.notebookId, notebookName: nbRef.displayName };
  }

  // Segment 2: section — try site-scoped filter; fall back to sync cache if 5000-limit error.
  const sectionName = segments[1]!;
  let section: { id: string; displayName: string; links?: any };
  try {
    const sectRes = await graphFetch(
      `${nbRef.apiBase}/onenote/sections?$filter=displayName eq '${odataEscape(sectionName)}'&$expand=parentNotebook($select=id)&$select=id,displayName,links&$top=50`
    );
    const sectData: GraphResponse<any> = await sectRes.json();
    const candidates = (sectData.value ?? []).filter(
      (s: any) => !s.parentNotebook?.id || s.parentNotebook.id === nbRef.notebookId
    );
    section = uniqueOrThrow(candidates, "section", sectionName, nbName);
  } catch (err: any) {
    if (err?.message?.startsWith("Multiple sections") || err?.message?.startsWith("section ")) throw err;
    if (!is5000LimitError(err) && err?.statusCode !== 404) throw err;
    const fromCache = await lookupSectionFromSyncCache(nbName, sectionName);
    if (!fromCache) {
      throw new Error(
        `Section '${sectionName}' in '${nbName}' not resolvable: site exceeds Graph API 5000-item limit. Run 'onenote sync' first, or pass a full URL.`
      );
    }
    section = fromCache;
  }

  if (segments.length === 2) {
    return {
      apiBase: nbRef.apiBase,
      notebookId: nbRef.notebookId,
      sectionId: section.id,
      notebookName: nbRef.displayName,
      sectionName: section.displayName,
    };
  }

  // Segment 3: page
  const pageTitle = segments[2]!;
  let page: { id: string; title: string; links?: any };
  try {
    const pageRes = await graphFetch(
      `${nbRef.apiBase}/onenote/sections/${section.id}/pages?$filter=title eq '${odataEscape(pageTitle)}'&$select=id,title,links&$top=50`
    );
    const pageData: GraphResponse<any> = await pageRes.json();
    page = uniqueOrThrow(pageData.value ?? [], "page", pageTitle, `${nbName}/${sectionName}`);
  } catch (err: any) {
    if (err?.message?.startsWith("Multiple pages") || err?.message?.startsWith("page ")) throw err;
    if (!is5000LimitError(err) && err?.statusCode !== 404) throw err;
    const fromCache = await lookupPageFromSyncCache(nbName, sectionName, pageTitle);
    if (!fromCache) {
      throw new Error(
        `Page '${pageTitle}' in '${nbName}/${sectionName}' not resolvable: site exceeds Graph API 5000-item limit. Run 'onenote sync' first, or pass a full URL.`
      );
    }
    page = fromCache;
  }

  return {
    apiBase: nbRef.apiBase,
    notebookId: nbRef.notebookId,
    sectionId: section.id,
    pageId: page.id,
    notebookName: nbRef.displayName,
    sectionName: section.displayName,
    pageTitle: page.title,
  };
}

async function resolveNotebookRef(
  urlOrId: string
): Promise<{ apiBase: string; notebookId: string; displayName?: string }> {
  if (!/^https?:\/\//i.test(urlOrId)) {
    if (looksLikeNotebookId(urlOrId)) {
      return { apiBase: "/me", notebookId: urlOrId };
    }
    // Name or path — delegate to path resolver (handles uniqueness)
    const r = await resolveByPath(urlOrId);
    if (!r.notebookId) throw new Error(`Path '${urlOrId}' does not resolve to a notebook.`);
    return { apiBase: r.apiBase, notebookId: r.notebookId, displayName: r.notebookName };
  }

  const parsed = parseOneNoteUrl(urlOrId);
  if (parsed.notebookId) {
    return { apiBase: "/me", notebookId: parsed.notebookId, displayName: parsed.notebookName };
  }

  if (!parsed.siteRef || !parsed.notebookName) {
    throw new Error("URL does not point to a notebook with a resolvable name.");
  }

  const siteId = await resolveSiteId(parsed.siteRef);
  const notebook = await resolveSharePointNotebook(siteId, parsed.notebookName);
  if (!notebook) throw new Error(`Notebook '${parsed.notebookName}' not found.`);
  return { apiBase: `/sites/${siteId}`, notebookId: notebook.id, displayName: notebook.displayName };
}

async function resolveSectionRef(
  urlOrId: string
): Promise<{ apiBase: string; sectionId: string; displayName?: string }> {
  if (!/^https?:\/\//i.test(urlOrId)) {
    if (urlOrId.includes("/") || !looksLikeGraphId(urlOrId)) {
      const r = await resolveByPath(urlOrId);
      if (!r.sectionId) throw new Error(`Path '${urlOrId}' does not resolve to a section (need notebook/section).`);
      return { apiBase: r.apiBase, sectionId: r.sectionId, displayName: r.sectionName };
    }
    return { apiBase: "/me", sectionId: urlOrId };
  }

  const parsed = parseOneNoteUrl(urlOrId);
  if (parsed.siteRef) {
    const siteId = await resolveSiteId(parsed.siteRef);
    const sectionName = parsed.sectionName ?? parsed.fileName;
    if (sectionName) {
      const section = await resolveSharePointSection(siteId, sectionName, parsed.sectionGuid);
      if (section) {
        return { apiBase: `/sites/${siteId}`, sectionId: section.id, displayName: section.displayName };
      }
    }
  }

  if (parsed.sectionGuid) {
    return {
      apiBase: "/me",
      sectionId: `0-${parsed.sectionGuid}`,
      displayName: parsed.sectionName ?? parsed.fileName,
    };
  }

  throw new Error("URL does not point to a specific section.");
}

/**
 * Resolve a URL-or-ID reference to the correct Graph API base + page ID.
 * - If the input looks like an http(s) URL, parse it and resolve via the
 *   SharePoint site-scoped endpoint when applicable (bypasses 5000-item limit).
 * - Otherwise, treat the input as a Graph page ID under /me/onenote/.
 */
async function resolvePageRef(
  urlOrId: string
): Promise<{ apiBase: string; pageId: string; title?: string; webUrl?: string }> {
  if (!/^https?:\/\//i.test(urlOrId)) {
    if (urlOrId.includes("/") || !looksLikeGraphId(urlOrId)) {
      const r = await resolveByPath(urlOrId);
      if (!r.pageId) throw new Error(`Path '${urlOrId}' does not resolve to a page (need notebook/section/page).`);
      return { apiBase: r.apiBase, pageId: r.pageId, title: r.pageTitle };
    }
    return { apiBase: "/me", pageId: urlOrId };
  }
  const parsed = parseOneNoteUrl(urlOrId);
  if (!parsed.siteRef || !parsed.sectionName || !parsed.pageTitle) {
    throw new Error(
      "URL does not point to a specific page (need both section and page in wd=target)."
    );
  }
  const siteId = await resolveSiteId(parsed.siteRef);
  const section = await resolveSharePointSection(siteId, parsed.sectionName, parsed.sectionGuid);
  if (!section) throw new Error(`Section '${parsed.sectionName}' not found.`);
  const page = await resolveSharePointPage(siteId, section.id, parsed.pageTitle, parsed.pageGuid);
  if (!page) throw new Error(`Page '${parsed.pageTitle}' not found in section '${section.displayName}'.`);
  return { apiBase: `/sites/${siteId}`, pageId: page.id, title: page.title, webUrl: page.webUrl };
}

/**
 * List pages in a section by its sourcedoc GUID.
 * Returns array of { id, title, webUrl }.
 */
export async function listSectionPagesByGuid(
  sectionGuid: string
): Promise<{ id: string; title: string; webUrl: string }[]> {
  const res = await graphFetch(
    `/me/onenote/sections/0-${sectionGuid}/pages?$select=id,title,links&$top=100`
  );
  const data: GraphResponse<any> = await res.json();
  return (data.value ?? []).map((p: any) => ({
    id: p.id,
    title: p.title ?? "(untitled)",
    webUrl: p.links?.oneNoteWebUrl?.href ?? "",
  }));
}

/**
 * Read a OneNote URL. Behavior depends on URL type:
 * - Page URL → returns page content
 * - Section URL → returns page list (tree view)
 * - Notebook URL → returns section list
 */
export async function readOneNoteUrl(
  url: string,
  options?: { downloadAssets?: boolean }
): Promise<{
  type: "page" | "section" | "notebook" | "resource";
  title: string;
  content: string; // text content or tree view
  html?: string;
  pageUrl?: string;
  assetPath?: string;
  breadcrumb?: string;
}> {
  if (isOneNoteResourceUrl(url)) {
    const resource = await renderResourceForRead(url);
    return {
      type: "resource",
      title: resource.title,
      content: resource.content,
      assetPath: resource.assetPath,
    };
  }

  // Non-URL input: treat as path or Graph ID
  if (!/^https?:\/\//i.test(url)) {
    // Raw Graph IDs (no "/") — route directly via resolvePageRef to preserve ID support
    if (!url.includes("/") && looksLikeGraphId(url)) {
      try {
        const r = await resolvePageRef(url);
        const contentRes = await graphFetch(`${r.apiBase}/onenote/pages/${r.pageId}/content`);
        const html = await contentRes.text();
        return {
          type: "page",
          title: r.title ?? "(untitled)",
          content: await renderHtmlForRead(html, options),
          html,
        };
      } catch {
        // Fall through to path resolution
      }
    }
    const r = await resolveByPath(url);
    if (r.pageId) {
      const contentRes = await graphFetch(`${r.apiBase}/onenote/pages/${r.pageId}/content`);
      const html = await contentRes.text();
      const crumbs = [r.notebookName, r.sectionName, r.pageTitle].filter(Boolean);
      return {
        type: "page",
        title: r.pageTitle ?? "(untitled)",
        content: await renderHtmlForRead(html, options),
        html,
        breadcrumb: crumbs.length > 1 ? crumbs.join(" / ") : undefined,
      };
    }
    if (r.sectionId) {
      const pagesRes = await graphFetch(
        `${r.apiBase}/onenote/sections/${r.sectionId}/pages?$select=id,title&$top=100`
      );
      const pagesData: GraphResponse<any> = await pagesRes.json();
      const pageList = pagesData.value ?? [];
      const tree = pageList.map((p: any, i: number) => `${i + 1}. ${p.title ?? "(untitled)"}`).join("\n");
      return {
        type: "section",
        title: r.sectionName ?? "Section",
        content: `${r.sectionName} (${pageList.length} pages)\n\n${tree}`,
      };
    }
    if (r.notebookId) {
      const sections = await listSections(url);
      const tree = sections.map((s: any, i: number) => `${i + 1}. ${s.displayName ?? "(untitled)"}`).join("\n");
      return {
        type: "notebook",
        title: r.notebookName ?? "Notebook",
        content: `${r.notebookName} (${sections.length} sections)\n\n${tree}`,
      };
    }
  }

  const parsed = parseOneNoteUrl(url);

  // --- SharePoint (business) path: use site-scoped endpoints to bypass
  //     the 5000-OneNote-item limit on /me/onenote/*.
  if (parsed.siteRef && parsed.sectionName) {
    try {
      const siteId = await resolveSiteId(parsed.siteRef);
      const section = await resolveSharePointSection(siteId, parsed.sectionName, parsed.sectionGuid);
      if (section) {
        // Page?
        if (parsed.pageTitle) {
          const page = await resolveSharePointPage(siteId, section.id, parsed.pageTitle, parsed.pageGuid);
          if (page) {
            const contentRes = await graphFetch(`/sites/${siteId}/onenote/pages/${page.id}/content`);
            const html = await contentRes.text();
            return {
              type: "page",
              title: page.title,
              content: await renderHtmlForRead(html, options),
              html,
              pageUrl: page.webUrl,
            };
          }
        }
        // Section tree view
        const pagesRes = await graphFetch(
          `/sites/${siteId}/onenote/sections/${section.id}/pages?$select=id,title&$top=100`
        );
        const pagesData: GraphResponse<any> = await pagesRes.json();
        const pageList = pagesData.value ?? [];
        const tree = pageList.map((p: any, i: number) => `${i + 1}. ${p.title ?? "(untitled)"}`).join("\n");
        return {
          type: "section",
          title: section.displayName,
          content: `${section.displayName} (${pageList.length} pages)\n\n${tree}`,
        };
      }
    } catch (err: any) {
      // Fall through to legacy paths; include message in final error if none match.
      if (!is5000LimitError(err) && err?.statusCode !== 404) throw err;
    }
  }

  // --- Page ---
  if (parsed.type === "page" && parsed.sectionGuid) {
    const pages = await listSectionPagesByGuid(parsed.sectionGuid);
    // Find matching page by GUID
    let targetPage: any = null;
    for (const page of pages) {
      const urlDecoded = decodeURIComponent(page.webUrl);
      const urlGuids = urlDecoded.match(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/gi) ?? [];
      const lastGuid = urlGuids.length > 0 ? urlGuids[urlGuids.length - 1].toLowerCase() : "";
      if (parsed.pageGuid && lastGuid === parsed.pageGuid) {
        targetPage = page;
        break;
      }
    }
    if (!targetPage && pages.length > 0) targetPage = pages[0];
    if (targetPage) {
      const contentRes = await graphFetch(`/me/onenote/pages/${targetPage.id}/content`);
      const html = await contentRes.text();
      return {
        type: "page",
        title: targetPage.title,
        content: await renderHtmlForRead(html, options),
        html,
        pageUrl: targetPage.webUrl,
      };
    }
  }

  // --- Section ---
  if ((parsed.type === "section" || parsed.type === "page") && parsed.sectionGuid) {
    const pages = await listSectionPagesByGuid(parsed.sectionGuid);
    // Build tree view
    const tree = pages.map((p, i) => `${i + 1}. ${p.title}`).join("\n");
    // Get section name
    let sectionName = "Section";
    try {
      const secRes = await graphFetch(`/me/onenote/sections/0-${parsed.sectionGuid}?$select=displayName`);
      const sec = await secRes.json() as any;
      sectionName = sec.displayName ?? sectionName;
    } catch {}
    return {
      type: "section",
      title: sectionName,
      content: `${sectionName} (${pages.length} pages)\n\n${tree}`,
    };
  }

  // --- Notebook ---
  if (parsed.type === "notebook") {
    try {
      const { apiBase, notebookId } = await resolveNotebookRef(url);
      const nbRes = await graphFetch(`${apiBase}/onenote/notebooks/${notebookId}`);
      const nb = await nbRes.json() as any;
      const sections = await listSections(url);
      const tree = sections.map((s: any, i: number) => `${i + 1}. ${s.displayName ?? s.name ?? "(untitled)"}`).join("\n");
      return {
        type: "notebook",
        title: nb.displayName ?? "Notebook",
        content: `${nb.displayName} (${sections.length} sections)\n\n${tree}`,
      };
    } catch {
      // Fall through
    }
  }

  throw new Error("Could not parse URL or access content. Supported: page, section, or notebook URLs.");
}

// Keep backward compat
export async function readPageByUrl(url: string) {
  const result = await readOneNoteUrl(url);
  return { title: result.title, html: result.html ?? "", text: result.content, pageUrl: result.pageUrl ?? url };
}
