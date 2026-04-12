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
