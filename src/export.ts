import { mkdir, readdir, writeFile } from "node:fs/promises";
import { dirname, extname, join, resolve } from "node:path";
import * as graph from "./graph";
import { renderHtmlForExport, renderHtmlForRead } from "./read-render";

export type ExportFormat = "md" | "html";

/** YYYY-MM-DD in local time, for [date] token expansion in output paths. */
function todayStamp(): string {
  const d = new Date();
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

/** Expand path tokens: [date] -> today's date (YYYY-MM-DD). */
export function substituteTokens(p: string): string {
  return p.replace(/\[date\]/gi, todayStamp());
}

/** Format is chosen by the output file extension; .html/.htm -> HTML, else Markdown. */
export function inferFormat(outputPath: string, explicit?: string): ExportFormat {
  if (explicit) {
    const f = explicit.toLowerCase();
    if (f === "md" || f === "markdown") return "md";
    if (f === "html" || f === "htm") return "html";
    throw new Error(`Unknown format '${explicit}'. Use 'md' or 'html'.`);
  }
  const ext = extname(outputPath).toLowerCase();
  if (ext === ".html" || ext === ".htm") return "html";
  return "md";
}

async function countFiles(dir: string): Promise<number> {
  try {
    const entries = await readdir(dir, { withFileTypes: true });
    return entries.filter((e) => e.isFile()).length;
  } catch {
    return 0;
  }
}

export type ExportResult = {
  outputPath: string;
  assetDir: string;
  format: ExportFormat;
  title: string;
  assetCount: number;
};

/**
 * Export a single OneNote page to a Markdown or HTML file, downloading all
 * referenced media into an assets directory (default: <output-dir>/assets) and
 * rewriting links to point at the local copies relative to the output file.
 */
export async function exportPage(
  url: string,
  output: string,
  opts?: { assetsDir?: string; format?: string }
): Promise<ExportResult> {
  const outputPath = resolve(substituteTokens(output));
  const format = inferFormat(outputPath, opts?.format);
  const outDir = dirname(outputPath);
  const assetDir = opts?.assetsDir
    ? resolve(substituteTokens(opts.assetsDir))
    : join(outDir, "assets");

  // Fetch raw HTML without downloading assets to the default cache; we render
  // (and download) into the export's own asset directory below.
  const result = await graph.readOneNoteUrl(url, { downloadAssets: false });
  if (result.type !== "page" || !result.html) {
    throw new Error(
      `URL does not resolve to a single page (got '${result.type}'). ` +
        `'export' operates on pages — open a specific page and copy its link.`
    );
  }

  await mkdir(outDir, { recursive: true });

  let body: string;
  if (format === "html") {
    body = await renderHtmlForExport(result.html, { assetDir, linkBaseDir: outDir });
  } else {
    const md = await renderHtmlForRead(result.html, { assetDir, linkBaseDir: outDir });
    const heading =
      result.title && result.title !== "(untitled)" ? `# ${result.title}\n\n` : "";
    body = `${heading}${md}\n`;
  }

  await writeFile(outputPath, body);

  return {
    outputPath,
    assetDir,
    format,
    title: result.title,
    assetCount: await countFiles(assetDir),
  };
}
