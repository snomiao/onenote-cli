import { mkdir, readdir, writeFile } from "node:fs/promises";
import { dirname, extname, join, resolve } from "node:path";
import * as graph from "./graph";
import { renderHtmlForExport, renderHtmlForRead } from "./read-render";

export type ExportFormat = "md" | "html";

/** YYYY-MM-DD in local time, for [date] token expansion and frontmatter. */
function todayStamp(): string {
  const d = new Date();
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

/** Expand path tokens: [date] -> today's date (YYYY-MM-DD). */
export function substituteTokens(p: string): string {
  return p.replace(/\[date\]/gi, todayStamp());
}

const FILE_EXTS = new Set([".md", ".markdown", ".html", ".htm"]);

export function inferFormat(pathOrExplicit: string, explicit?: string): ExportFormat {
  if (explicit) {
    const f = explicit.toLowerCase();
    if (f === "md" || f === "markdown") return "md";
    if (f === "html" || f === "htm") return "html";
    throw new Error(`Unknown format '${explicit}'. Use 'md' or 'html'.`);
  }
  const ext = extname(pathOrExplicit).toLowerCase();
  if (ext === ".html" || ext === ".htm") return "html";
  return "md";
}

/** Make a title safe as a single filesystem name. Unicode is preserved; only
 *  path-illegal characters, control chars, and trailing dots/spaces are removed. */
export function sanitizeFilename(title: string, fallback = "untitled"): string {
  let s = (title ?? "")
    // eslint-disable-next-line no-control-regex -- intentionally stripping control chars
    .replace(/[\u0000-\u001f\u007f]/g, " ") // newlines/tabs are invalid in filenames
    .replace(/[/\\:*?"<>|]/g, "-") // illegal on Windows/Unix
    .replace(/\s+/g, " ")
    .trim()
    .replace(/^\.+/, "") // no leading dots (hidden files / traversal)
    .replace(/[. ]+$/, ""); // no trailing dot/space (Windows)
  // Keep filenames comfortably under the 255-byte limit, leaving room for an
  // extension and a collision suffix. Truncate by codepoint on a byte budget.
  const MAX_BYTES = 180;
  if (Buffer.byteLength(s, "utf8") > MAX_BYTES) {
    let out = "";
    for (const ch of s) {
      if (Buffer.byteLength(out + ch, "utf8") > MAX_BYTES) break;
      out += ch;
    }
    s = out.trim().replace(/[. ]+$/, "");
  }
  // Windows reserved device names.
  if (/^(con|prn|aux|nul|com[1-9]|lpt[1-9])$/i.test(s)) s = `_${s}`;
  return s || fallback;
}

/** Return `name` (or `name.ext`) unique within `used`, appending " (2)", " (3)"… */
function uniqueName(base: string, used: Set<string>, ext = ""): string {
  const make = (n: number) => (n === 1 ? `${base}${ext}` : `${base} (${n})${ext}`);
  let n = 1;
  while (used.has(make(n).toLowerCase())) n++;
  const chosen = make(n);
  used.add(chosen.toLowerCase());
  return chosen;
}

function yamlQuote(v: string): string {
  return `"${v.replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/\n/g, " ")}"`;
}

function htmlAttr(v: string): string {
  return v.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

type FrontMatter = {
  title?: string;
  source?: string;
  notebook?: string;
  section?: string;
  exported?: string;
};

function mdFrontMatter(fm: FrontMatter): string {
  const lines = Object.entries(fm)
    .filter(([, v]) => v !== undefined && v !== "")
    .map(([k, v]) => `${k}: ${yamlQuote(String(v))}`);
  return lines.length ? `---\n${lines.join("\n")}\n---\n\n` : "";
}

/** Inject <meta name="onenote-*"> tags into an HTML document's <head>. */
function htmlWithMeta(html: string, fm: FrontMatter): string {
  const metas = Object.entries(fm)
    .filter(([, v]) => v !== undefined && v !== "")
    .map(([k, v]) => `<meta name="onenote-${k}" content="${htmlAttr(String(v))}" />`)
    .join("\n\t\t");
  if (!metas) return html;
  if (/<head\b[^>]*>/i.test(html)) {
    return html.replace(/<head\b[^>]*>/i, (m) => `${m}\n\t\t${metas}`);
  }
  return `<head>\n\t\t${metas}\n\t</head>\n${html}`;
}

async function countFiles(dir: string): Promise<number> {
  try {
    const entries = await readdir(dir, { withFileTypes: true });
    return entries.filter((e) => e.isFile()).length;
  } catch {
    return 0;
  }
}

/** Run `fn` over items with bounded concurrency, preserving order. */
async function mapPool<T, R>(items: T[], concurrency: number, fn: (item: T, i: number) => Promise<R>): Promise<R[]> {
  const results: R[] = Array.from({ length: items.length });
  let cursor = 0;
  const width = Math.max(1, Math.min(concurrency, items.length));
  await Promise.all(
    Array.from({ length: width }, async () => {
      while (true) {
        const i = cursor++;
        if (i >= items.length) break;
        results[i] = await fn(items[i]!, i);
      }
    })
  );
  return results;
}

export type ExportSummary = {
  rootPath: string;
  format: ExportFormat;
  kind: graph.ExportNode["kind"];
  title: string;
  pageCount: number;
  assetCount: number;
};

type Ctx = {
  format: ExportFormat;
  assetsDirOverride?: string;
  log?: (msg: string) => void;
  assetDirs: Set<string>;
  pageCount: { n: number };
};

/** Render one page's HTML into a Markdown/HTML file at `filePath`, with media
 *  downloaded next to it (or into the shared --assets-dir) and frontmatter. */
async function writePageFile(
  base: string,
  page: { id: string; title: string; webUrl?: string },
  filePath: string,
  ctx: Ctx,
  crumbs: { notebook?: string; section?: string }
): Promise<void> {
  const outDir = dirname(filePath);
  const assetDir = ctx.assetsDirOverride ?? join(outDir, "assets");
  const html = await graph.fetchPageHtml(base, page.id);

  const fm: FrontMatter = {
    title: page.title || undefined,
    source: page.webUrl,
    notebook: crumbs.notebook,
    section: crumbs.section,
    exported: todayStamp(),
  };

  let body: string;
  if (ctx.format === "html") {
    const rendered = await renderHtmlForExport(html, { assetDir, linkBaseDir: outDir });
    body = htmlWithMeta(rendered, fm);
  } else {
    const md = await renderHtmlForRead(html, { assetDir, linkBaseDir: outDir });
    const heading = page.title ? `# ${page.title}\n\n` : "";
    body = `${mdFrontMatter(fm)}${heading}${md}\n`;
  }

  await mkdir(outDir, { recursive: true });
  await writeFile(filePath, body);
  ctx.assetDirs.add(assetDir);
  ctx.pageCount.n++;
  ctx.log?.(`  ${filePath}`);
}

/** Write an already-listed set of pages into `dir` (one file per page, by title). */
async function writeSectionPages(
  base: string,
  pages: { id: string; title: string; webUrl?: string }[],
  dir: string,
  ctx: Ctx,
  crumbs: { notebook?: string; section?: string }
): Promise<void> {
  ctx.log?.(`${crumbs.section ?? "section"}: ${pages.length} page(s) -> ${dir}`);
  await mkdir(dir, { recursive: true });

  // Assign unique filenames sequentially before fetching pages in parallel.
  const used = new Set<string>();
  const ext = ctx.format === "html" ? ".html" : ".md";
  const planned = pages.map((p) => ({
    page: p,
    filePath: join(dir, uniqueName(sanitizeFilename(p.title, "untitled"), used, ext)),
  }));

  await mapPool(planned, 5, ({ page, filePath }) => writePageFile(base, page, filePath, ctx, crumbs));
}

/** Export a section (by Graph section id) into `dir`. */
async function exportSection(
  base: string,
  sectionId: string,
  dir: string,
  ctx: Ctx,
  crumbs: { notebook?: string; section?: string }
): Promise<void> {
  const pages = await graph.listSectionPages(base, sectionId);
  await writeSectionPages(base, pages, dir, ctx, crumbs);
}

/**
 * Recursively export a notebook/section-group folder via the OneDrive tree.
 * Used when the OneNote API refuses to enumerate a >5,000-item library.
 */
async function exportContainerViaDrive(
  drivePath: string,
  dir: string,
  ctx: Ctx,
  crumbs: { notebook?: string; section?: string }
): Promise<void> {
  const { sections, groups } = await graph.listDriveContainer(drivePath);
  await mkdir(dir, { recursive: true });
  const usedDirs = new Set<string>();
  for (const s of sections) {
    const subdir = join(dir, uniqueName(sanitizeFilename(s.name, "section"), usedDirs));
    const pages = await graph.listSectionPagesByGuidAll(s.sectionGuid);
    await writeSectionPages("/me", pages, subdir, ctx, { ...crumbs, section: s.name });
  }
  for (const g of groups) {
    const subdir = join(dir, uniqueName(sanitizeFilename(g.name, "group"), usedDirs));
    await exportContainerViaDrive(g.drivePath, subdir, ctx, crumbs);
  }
}

/** Recursively export a notebook or section group as a directory tree. */
async function exportContainer(
  node: graph.ExportNode,
  dir: string,
  ctx: Ctx,
  crumbs: { notebook?: string; section?: string }
): Promise<void> {
  const parentType = node.kind === "notebook" ? "notebooks" : "sectionGroups";
  const nextCrumbs = node.kind === "notebook" ? { ...crumbs, notebook: node.title } : crumbs;
  const drivePath = "drivePath" in node ? node.drivePath : undefined;
  await mkdir(dir, { recursive: true });

  let sections: { id: string; displayName: string }[];
  let groups: { id: string; displayName: string }[];
  try {
    [sections, groups] = await Promise.all([
      graph.listChildSections(node.base, parentType, node.id),
      graph.listChildSectionGroups(node.base, parentType, node.id),
    ]);
  } catch (err: any) {
    const over5000 =
      err?.statusCode === 403 && /5,?000 OneNote items|document librar/i.test(err?.message ?? "");
    if (over5000 && drivePath) {
      // Library too large for the OneNote API — enumerate via the OneDrive tree.
      ctx.log?.(
        `(library exceeds Graph 5,000-item limit — enumerating "${node.title}" via OneDrive)`
      );
      return exportContainerViaDrive(drivePath, dir, ctx, nextCrumbs);
    }
    if (over5000) {
      throw new Error(
        `Cannot enumerate the ${node.kind} "${node.title}": this OneDrive library exceeds the ` +
          `Graph API's 5,000-OneNote-item limit and its OneDrive path could not be resolved. ` +
          `Export sections individually instead — pass a section URL or path to 'onenote export'.`
      );
    }
    throw err;
  }

  const usedDirs = new Set<string>();
  for (const s of sections) {
    const subdir = join(dir, uniqueName(sanitizeFilename(s.displayName, "section"), usedDirs));
    await exportSection(node.base, s.id, subdir, ctx, { ...nextCrumbs, section: s.displayName });
  }
  for (const g of groups) {
    const subdir = join(dir, uniqueName(sanitizeFilename(g.displayName, "group"), usedDirs));
    await exportContainer(
      { kind: "sectionGroup", base: node.base, id: g.id, title: g.displayName },
      subdir,
      ctx,
      nextCrumbs
    );
  }
}

/**
 * Export a OneNote page, section, section group, or notebook.
 * - page  -> a single Markdown/HTML file (format from output extension).
 * - section/group/notebook -> a directory tree, one file per page named by its
 *   sanitized title, format chosen by --format (default Markdown).
 * Media is downloaded next to each file (`<dir>/assets`, or --assets-dir).
 */
export async function exportRef(
  ref: string,
  output: string,
  opts?: { assetsDir?: string; format?: string; log?: (msg: string) => void }
): Promise<ExportSummary> {
  const node = await graph.resolveExportNode(ref);
  const outPath = resolve(substituteTokens(output));
  const assetsDirOverride = opts?.assetsDir ? resolve(substituteTokens(opts.assetsDir)) : undefined;

  const ctx: Ctx = {
    format: "md",
    assetsDirOverride,
    log: opts?.log,
    assetDirs: new Set<string>(),
    pageCount: { n: 0 },
  };

  if (node.kind === "page") {
    // Output may be a concrete file, or a directory to place <title>.<ext> in.
    const asFile = FILE_EXTS.has(extname(outPath).toLowerCase());
    ctx.format = inferFormat(asFile ? outPath : "", opts?.format);
    const ext = ctx.format === "html" ? ".html" : ".md";
    const filePath = asFile
      ? outPath
      : join(outPath, `${sanitizeFilename(node.title, "untitled")}${ext}`);
    const pageForFm = { id: node.id, title: node.title, webUrl: node.webUrl ?? (/^https?:\/\//i.test(ref) ? ref : undefined) };
    await writePageFile(node.base, pageForFm, filePath, ctx, {});
    return {
      rootPath: filePath,
      format: ctx.format,
      kind: node.kind,
      title: node.title,
      pageCount: ctx.pageCount.n,
      assetCount: await sumAssets(ctx.assetDirs),
    };
  }

  // Container export -> directory tree.
  ctx.format = inferFormat("", opts?.format);
  if (node.kind === "section") {
    await exportSection(node.base, node.id, outPath, ctx, { section: node.title });
  } else {
    await exportContainer(node, outPath, ctx, {});
  }

  return {
    rootPath: outPath,
    format: ctx.format,
    kind: node.kind,
    title: node.title,
    pageCount: ctx.pageCount.n,
    assetCount: await sumAssets(ctx.assetDirs),
  };
}

async function sumAssets(dirs: Set<string>): Promise<number> {
  let total = 0;
  for (const d of dirs) total += await countFiles(d);
  return total;
}
