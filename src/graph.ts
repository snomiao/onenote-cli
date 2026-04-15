import { getAccessToken } from "./auth";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

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

export async function getNotebook(id: string) {
  const res = await graphFetch(`/me/onenote/notebooks/${id}`);
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
    const path = notebookId
      ? `/me/onenote/notebooks/${notebookId}/sections`
      : "/me/onenote/sections";
    const res = await graphFetch(path);
    const data: GraphResponse<any> = await res.json();
    return data.value;
  } catch (err: any) {
    if (!is5000LimitError(err) || !notebookId) throw err;
    // Fallback: list .one files via OneDrive
    const notebook = await getNotebook(notebookId);
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

export async function getSection(id: string) {
  const res = await graphFetch(`/me/onenote/sections/${id}`);
  return res.json();
}

export async function createSection(notebookId: string, displayName: string) {
  const res = await graphFetch(`/me/onenote/notebooks/${notebookId}/sections`, {
    method: "POST",
    body: JSON.stringify({ displayName }),
  });
  return res.json();
}

// --- Section Groups ---

export async function listSectionGroups(notebookId?: string) {
  try {
    const path = notebookId
      ? `/me/onenote/notebooks/${notebookId}/sectionGroups`
      : "/me/onenote/sectionGroups";
    const res = await graphFetch(path);
    const data: GraphResponse<any> = await res.json();
    return data.value;
  } catch (err: any) {
    if (!is5000LimitError(err) || !notebookId) throw err;
    // Fallback: list folders via OneDrive
    const notebook = await getNotebook(notebookId);
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
  const path = sectionId
    ? `/me/onenote/sections/${sectionId}/pages`
    : "/me/onenote/pages";
  const res = await graphFetch(path);
  const data: GraphResponse<any> = await res.json();
  return data.value;
}

export async function getPage(id: string) {
  const res = await graphFetch(`/me/onenote/pages/${id}`);
  return res.json();
}

export async function getPageContent(id: string): Promise<string> {
  const res = await graphFetch(`/me/onenote/pages/${id}/content`);
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
export async function renameSection(sectionId: string, newName: string) {
  const res = await graphFetch(`/me/onenote/sections/${sectionId}`, {
    method: "PATCH",
    body: JSON.stringify({ displayName: newName }),
  });
  return res.json();
}

/**
 * Rename a notebook.
 */
export async function renameNotebook(notebookId: string, newName: string) {
  const res = await graphFetch(`/me/onenote/notebooks/${notebookId}`, {
    method: "PATCH",
    body: JSON.stringify({ displayName: newName }),
  });
  return res.json();
}

export async function createPage(sectionId: string, title: string, htmlBody: string) {
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

  const res = await graphFetch(`/me/onenote/sections/${sectionId}/pages`, {
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
  // SharePoint / business notebook resolution hints
  siteRef?: string;        // e.g. "host:/personal/user" for /sites endpoint lookup
  notebookName?: string;   // from file= param (may have .one stripped elsewhere)
  sectionName?: string;    // from wd=target first segment (with .one stripped)
  pageTitle?: string;      // from wd=target second segment (unescaped)
}

function unescapeOneNoteName(s: string): string {
  return s.replace(/\\(.)/g, "$1");
}

function parseOneNoteUrl(url: string): ParsedOneNoteUrl {
  const decoded = decodeURIComponent(url);
  const guids = decoded.match(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/gi) ?? [];

  const sourcedocMatch = decoded.match(/sourcedoc=\{?([0-9a-f-]+)\}?/i);
  let sectionGuid = sourcedocMatch?.[1]?.toLowerCase();

  // SharePoint siteRef from the URL host + /personal/{user} or /sites/{name} segment
  let siteRef: string | undefined;
  try {
    const u = new URL(url);
    if (/sharepoint\.com$/i.test(u.hostname)) {
      const p = decodeURIComponent(u.pathname);
      const personal = p.match(/\/(personal\/[^/]+)/i);
      const site = p.match(/\/(sites|teams)\/([^/]+)/i);
      if (personal) siteRef = `${u.hostname}:/${personal[1]}:`;
      else if (site) siteRef = `${u.hostname}:/${site[1]}/${site[2]}:`;
    }
  } catch {}

  // file= gives the notebook name for Doc.aspx URLs
  const fileMatch = decoded.match(/[?&]file=([^&]+)/i);
  const notebookName = fileMatch ? decodeURIComponent(fileMatch[1]).replace(/\.one$/i, "") : undefined;

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

  const nbMatch = decoded.match(/notebooks\/(1-[0-9a-f-]+)/i);
  const notebookId = nbMatch?.[1];

  const base = { siteRef, notebookName, sectionName, pageTitle };
  if (pageGuid && sectionGuid) return { type: "page", sectionGuid, pageGuid, ...base };
  if (sectionGuid) return { type: "section", sectionGuid, ...base };
  if (notebookId) return { type: "notebook", notebookId, ...base };
  return { type: "unknown", ...base };
}

function htmlToText(html: string): string {
  return html
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, "")
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n\n")
    .replace(/<\/div>/gi, "\n")
    .replace(/<\/h[1-6]>/gi, "\n\n")
    .replace(/<\/li>/gi, "\n")
    .replace(/<li[^>]*>/gi, "- ")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
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
export async function readOneNoteUrl(url: string): Promise<{
  type: "page" | "section" | "notebook";
  title: string;
  content: string; // text content or tree view
  html?: string;
  pageUrl?: string;
}> {
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
              content: htmlToText(html),
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
        content: htmlToText(html),
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
  if (parsed.type === "notebook" && parsed.notebookId) {
    try {
      const nb = await getNotebook(parsed.notebookId);
      const sections = await listSections(parsed.notebookId);
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
