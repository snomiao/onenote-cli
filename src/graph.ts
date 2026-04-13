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

export async function deletePage(id: string) {
  await graphFetch(`/me/onenote/pages/${id}`, { method: "DELETE" });
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
function parseOneNoteUrl(url: string): {
  type: "page" | "section" | "notebook" | "unknown";
  sectionGuid?: string;
  pageGuid?: string;
  notebookId?: string;
} {
  const decoded = decodeURIComponent(url);
  const guids = decoded.match(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/gi) ?? [];

  // sourcedoc GUID (section)
  const sourcedocMatch = decoded.match(/sourcedoc=\{?([0-9a-f-]+)\}?/i);
  const sectionGuid = sourcedocMatch?.[1]?.toLowerCase();

  // Check for wd=target (means section or page level)
  const hasWdTarget = /wd=target/i.test(decoded);

  // Page GUID is the LAST GUID in wd=target (after section group GUID + title)
  const pageGuid = hasWdTarget && guids.length > 1
    ? guids[guids.length - 1].toLowerCase()
    : undefined;

  // Detect notebook ID from URL path like /notebooks/1-{guid}
  const nbMatch = decoded.match(/notebooks\/(1-[0-9a-f-]+)/i);
  const notebookId = nbMatch?.[1];

  if (pageGuid && sectionGuid) return { type: "page", sectionGuid, pageGuid };
  if (sectionGuid) return { type: "section", sectionGuid };
  if (notebookId) return { type: "notebook", notebookId };
  if (sectionGuid) return { type: "section", sectionGuid };
  return { type: "unknown" };
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
