#!/usr/bin/env bun
import { createHash } from "node:crypto";
import yargs from "yargs";
import { hideBin } from "yargs/helpers";
import * as graph from "./graph";
import { login, logout, logoutAll, whoami, listAccounts } from "./auth";
import { syncCache, searchLocal, isCacheEmpty, rebuildSearchIndex, SEARCH_DB_PATH, parseTagsFromQuery, TAG_ALIASES } from "./cache";
import { stat } from "node:fs/promises";
import { markdownToHtml } from "./markdown";

const isTTY = process.stdout.isTTY ?? false;

// ANSI color helpers (only when TTY)
const bold = (s: string) => (isTTY ? `\x1b[1m${s}\x1b[0m` : s);
const dim = (s: string) => (isTTY ? `\x1b[2m${s}\x1b[0m` : s);
const yellow = (s: string) => (isTTY ? `\x1b[33m${s}\x1b[0m` : s);
const cyan = (s: string) => (isTTY ? `\x1b[36m${s}\x1b[0m` : s);
const green = (s: string) => (isTTY ? `\x1b[32m${s}\x1b[0m` : s);
const magenta = (s: string) => (isTTY ? `\x1b[35m${s}\x1b[0m` : s);
const pageUpdateActions = ["append", "insert", "prepend", "replace"] as const;
const pageInsertPositions = ["before", "after"] as const;

// OSC 8 hyperlink (clickable in supported terminals), markdown in non-TTY
function link(url: string, text: string): string {
  if (!isTTY) return `[${text}](${url})`;
  return `\x1b]8;;${url}\x1b\\${text}\x1b]8;;\x1b\\`;
}

function contentSha4(html: string): string {
  return createHash("sha256").update(html).digest("hex").slice(0, 4);
}

async function confirmPageSha(ref: string, providedSha: string | undefined, verb: string) {
  const result = await graph.readOneNoteUrl(ref, { downloadAssets: false });
  if (result.type !== "page" || !result.html) {
    throw new Error(`'${ref}' is not a page — ${verb} only operates on pages.`);
  }
  const sha = contentSha4(result.html);

  if (!providedSha) {
    console.log(bold(cyan(result.title)));
    if (result.breadcrumb) console.log(dim(result.breadcrumb));
    console.log(dim("─".repeat(Math.min(result.title.length + 10, 60))));
    console.log(result.content);
    console.error("");
    console.error(yellow(`Dry run. To ${verb}, re-run with --sha=${sha}`));
    return { confirmed: false as const, sha };
  }
  if (providedSha !== sha) {
    throw new Error(
      `sha mismatch: expected '${sha}', got '${providedSha}'. Re-read the page — content may have changed.`
    );
  }
  return { confirmed: true as const, sha };
}

type ListItem = { name: string; url?: string; subtitle?: string };

function printList(items: ListItem[]) {
  if (!items || items.length === 0) {
    console.log("No results found.");
    return;
  }
  for (const it of items) {
    const main = it.url ? link(it.url, it.name) : it.name;
    const sub = it.subtitle ? ` ${dim(it.subtitle)}` : "";
    console.log(`- ${main}${sub}`);
  }
}

function toListItem(raw: any, type?: "notebook" | "section" | "page"): ListItem {
  const url =
    raw?.links?.oneNoteWebUrl?.href
    ?? raw?.links?.oneNoteClientUrl?.href
    ?? raw?.webUrl
    ?? "";
  const name = raw?.displayName ?? raw?.title ?? "(untitled)";
  const date = raw?.lastModifiedDateTime ?? raw?.createdDateTime;
  const typeSuffix = type ? dim(` .${type}`) : "";
  let locationHint: string | undefined;
  if (type === "notebook" && url) {
    try {
      const pathname = decodeURIComponent(new URL(url).pathname.replace(/\/$/, ""));
      // Strip /personal/{user}/Documents/ prefix, then show parent folder only if non-standard
      const afterDocuments = pathname.replace(/^\/[^/]+\/[^/]+\/Documents\//, "");
      const segments = afterDocuments.split("/");
      // Standard path: Notebooks/{name} — hide, not informative
      if (segments[0] !== "Notebooks") {
        // Show parent folder(s) only (everything except the last segment)
        const parent = segments.slice(0, -1).join("/");
        locationHint = parent ? `${parent}/` : "Documents/";
      }
    } catch {}
  }
  const parts = [locationHint, date ? String(date).slice(0, 10) : undefined].filter(Boolean);
  return { name: `${name}${typeSuffix}`, url, subtitle: parts.length ? parts.join("  ") : undefined };
}

function outputList(items: any[], argv: { json?: boolean; limit?: number }, type?: "notebook" | "section" | "page") {
  const limited = typeof argv.limit === "number" ? items.slice(0, argv.limit) : items;
  if (argv.json) {
    console.log(JSON.stringify(limited, null, 2));
    return;
  }
  printList(limited.map((r) => toListItem(r, type)));
}

function formatTable(items: any[], columns: { key: string; label: string }[]) {
  if (!items || items.length === 0) {
    console.log("No results found.");
    return;
  }
  const widths = columns.map((col) =>
    Math.max(col.label.length, ...items.map((item) => String(item[col.key] ?? "").length))
  );
  const header = columns.map((col, i) => col.label.padEnd(widths[i])).join("  ");
  const separator = widths.map((w) => "-".repeat(w)).join("  ");
  console.log(header);
  console.log(separator);
  for (const item of items) {
    const row = columns.map((col, i) => String(item[col.key] ?? "").padEnd(widths[i])).join("  ");
    console.log(row);
  }
}

function normalizeRef(value: string | undefined): string | undefined {
  if (typeof value !== "string") return undefined;
  return value.replace(/^["']|["']$/g, "");
}

function renderHtmlContent(content: string, markdown?: boolean): string {
  return markdown ? markdownToHtml(content) : content;
}

function renderHtmlBody(content: string | undefined, markdown?: boolean): string {
  if (!content || content.trim() === "") return "<p></p>";
  return renderHtmlContent(content, markdown);
}

yargs(hideBin(process.argv))
  .scriptName("onenote")
  .usage("$0 <command> [options]")
  .demandCommand(1)

  // --- top-level ls (auto-detects depth by path segments) ---
  .command(
    ["ls [path]", "list [path]"],
    "List notebooks, sections, or pages based on path depth",
    (y) =>
      y.positional("path", {
        type: "string",
        describe: "Path: empty=notebooks, 'nb'=sections, 'nb/sec'=pages",
      }),
    async (argv) => {
      const path = normalizeRef(argv.path as string | undefined);
      const segments = path ? path.split("/").filter(Boolean) : [];
      if (segments.length === 0) {
        const notebooks = await graph.listNotebooks();
        printList((notebooks ?? []).map((r) => toListItem(r, "notebook")));
      } else if (segments.length === 1) {
        const sections = await graph.listSections(path);
        printList((sections ?? []).map((r) => toListItem(r, "section")));
      } else if (segments.length === 2) {
        const pages = await graph.listPages(path);
        printList((pages ?? []).map((r) => toListItem(r, "page")));
      } else {
        throw new Error(
          `Path '${path}' points to a page, not a listable container. Use 'onenote read ${path}' to view it.`
        );
      }
    }
  )

  // --- notebooks ---
  .command(
    "notebooks",
    "Manage notebooks",
    (yargs) =>
      yargs
        .command(
          ["list", "ls"],
          "List all notebooks",
          (y) =>
            y
              .option("json", { type: "boolean", describe: "Output JSON" })
              .option("limit", { type: "number", describe: "Max items (default: all)" }),
          async (argv) => {
            const notebooks = await graph.listNotebooks();
            outputList(notebooks ?? [], argv, "notebook");
          }
        )
        .command(
          "get <ref>",
          "Get a notebook by name, path, ID, or URL",
          (y) => y.positional("ref", { type: "string", demandOption: true }),
          async (argv) => {
            const nb = await graph.getNotebook(normalizeRef(argv.ref as string)!);
            console.log(JSON.stringify(nb, null, 2));
          }
        )
        .command(
          "create <name>",
          "Create a new notebook",
          (y) => y.positional("name", { type: "string", demandOption: true }),
          async (argv) => {
            const nb = await graph.createNotebook(argv.name as string);
            console.log("Created notebook:", nb.displayName);
            console.log("ID:", nb.id);
          }
        )
        .command(
          "rename <ref> <name>",
          "Rename a notebook",
          (y) =>
            y
              .positional("ref", { type: "string", demandOption: true })
              .positional("name", { type: "string", demandOption: true }),
          async (argv) => {
            const nb = await graph.renameNotebook(normalizeRef(argv.ref as string)!, argv.name as string);
            console.log("Notebook renamed to:", nb.displayName);
          }
        )
        .demandCommand(1),
    () => {}
  )

  // --- sections ---
  .command(
    "sections",
    "Manage sections",
    (yargs) =>
      yargs
        .command(
          ["list [ref]", "ls [ref]"],
          "List sections",
          (y) =>
            y
              .positional("ref", { type: "string", describe: "Notebook name, path, ID, or URL" })
              .option("notebook", { type: "string", alias: ["n", "notebook-id"], describe: "Filter by notebook name, path, ID, or URL" })
              .option("json", { type: "boolean", describe: "Output JSON" })
              .option("limit", { type: "number", describe: "Max items (default: all)" }),
          async (argv) => {
            const notebookRef = normalizeRef((argv.ref as string | undefined) ?? (argv.notebook as string | undefined));
            const sections = await graph.listSections(notebookRef);
            outputList(sections ?? [], argv, "section");
          }
        )
        .command(
          "get <ref>",
          "Get a section by path, ID, or URL",
          (y) => y.positional("ref", { type: "string", demandOption: true }),
          async (argv) => {
            const section = await graph.getSection(normalizeRef(argv.ref as string)!);
            console.log(JSON.stringify(section, null, 2));
          }
        )
        .command(
          "create",
          "Create a new section in a notebook",
          (y) =>
            y
              .option("notebook", { type: "string", alias: ["n", "notebook-id"], demandOption: true, describe: "Notebook name, path, ID, or URL" })
              .option("name", { type: "string", demandOption: true, describe: "Section name" }),
          async (argv) => {
            const section = await graph.createSection(
              normalizeRef(argv.notebook as string)!,
              argv.name as string
            );
            console.log("Created section:", section.displayName);
            console.log("ID:", section.id);
          }
        )
        .command(
          "rename <ref> <name>",
          "Rename a section",
          (y) =>
            y
              .positional("ref", { type: "string", demandOption: true })
              .positional("name", { type: "string", demandOption: true }),
          async (argv) => {
            const section = await graph.renameSection(normalizeRef(argv.ref as string)!, argv.name as string);
            console.log("Section renamed to:", section.displayName);
          }
        )
        .demandCommand(1),
    () => {}
  )

  // --- section-groups ---
  .command(
    "section-groups",
    "Manage section groups",
    (yargs) =>
      yargs
        .command(
          ["list [ref]", "ls [ref]"],
          "List section groups",
          (y) =>
            y
              .positional("ref", { type: "string", describe: "Notebook name, path, ID, or URL" })
              .option("notebook", { type: "string", alias: ["n", "notebook-id"], describe: "Filter by notebook name, path, ID, or URL" })
              .option("json", { type: "boolean", describe: "Output JSON" })
              .option("limit", { type: "number", describe: "Max items (default: all)" }),
          async (argv) => {
            const notebookRef = normalizeRef((argv.ref as string | undefined) ?? (argv.notebook as string | undefined));
            const groups = await graph.listSectionGroups(notebookRef);
            outputList(groups ?? [], argv);
          }
        )
        .demandCommand(1),
    () => {}
  )

  // --- pages ---
  .command(
    "pages",
    "Manage pages",
    (yargs) =>
      yargs
        .command(
          ["list [ref]", "ls [ref]"],
          "List pages",
          (y) =>
            y
              .positional("ref", { type: "string", describe: "Section name, path (notebook/section), ID, or URL" })
              .option("section", { type: "string", alias: ["s", "section-id"], describe: "Filter by section name, path, ID, or URL" })
              .option("json", { type: "boolean", describe: "Output JSON" })
              .option("limit", { type: "number", describe: "Max items (default: all)" }),
          async (argv) => {
            const sectionRef = normalizeRef((argv.ref as string | undefined) ?? (argv.section as string | undefined));
            const pages = await graph.listPages(sectionRef);
            outputList(pages ?? [], argv, "page");
          }
        )
        .command(
          "get <ref>",
          "Get page metadata (accepts path, page ID, or OneNote URL)",
          (y) => y.positional("ref", { type: "string", demandOption: true }),
          async (argv) => {
            const page = await graph.getPage(normalizeRef(argv.ref as string)!);
            console.log(JSON.stringify(page, null, 2));
          }
        )
        .command(
          "content <ref>",
          "Get page HTML content (deprecated: use 'onenote read <ref> --html')",
          (y) => y.positional("ref", { type: "string", demandOption: true }),
          async (argv) => {
            console.error(dim("[deprecated] Use 'onenote read <ref> --html' instead."));
            const html = await graph.getPageContent(normalizeRef(argv.ref as string)!);
            console.log(html);
          }
        )
        .command(
          "create",
          "Create a new page",
          (y) =>
            y
              .option("section", { type: "string", alias: ["s", "section-id"], demandOption: true, describe: "Section name, path, ID, or URL" })
              .option("title", { type: "string", alias: "t", demandOption: true, describe: "Page title" })
              .option("body", { type: "string", alias: "b", describe: "HTML body content. Use data-tag to add OneNote tags: data-tag=\"to-do\" (checkbox), data-tag=\"star\", data-tag=\"question\", data-tag=\"important\", data-tag=\"critical\", data-tag=\"idea\", data-tag=\"contact\", data-tag=\"definition\", data-tag=\"highlight\", data-tag=\"password\", data-tag=\"remember-for-later\", data-tag=\"to-do:completed\" (checked)" })
              .option("md", { type: "boolean", describe: "Treat --body as Markdown and convert it to HTML" }),
          async (argv) => {
            const page = await graph.createPage(
              normalizeRef(argv.section as string)!,
              argv.title as string,
              renderHtmlBody(argv.body as string | undefined, argv.md as boolean | undefined)
            );
            console.log("Created page:", page.title);
            console.log("ID:", page.id);
          }
        )
        .command(
          "delete <ref>",
          "Delete a page. Without --sha, dry-runs and prints content + sha.",
          (y) =>
            y
              .positional("ref", { type: "string", demandOption: true })
              .option("sha", {
                type: "string",
                describe: "4-char content sha from 'onenote read' to confirm deletion",
              }),
          async (argv) => {
            const ref = normalizeRef(argv.ref as string)!;
            const { confirmed } = await confirmPageSha(ref, argv.sha as string | undefined, "delete");
            if (!confirmed) return;
            await graph.deletePage(ref);
            console.log(green("Page deleted."));
          }
        )
        .command(
          "rename <ref> <title>",
          "Rename a page (accepts path, page ID, or OneNote URL)",
          (y) =>
            y
              .positional("ref", { type: "string", demandOption: true })
              .positional("title", { type: "string", demandOption: true }),
          async (argv) => {
            await graph.renamePage(
              normalizeRef(argv.ref as string)!,
              argv.title as string
            );
            console.log("Page renamed to:", argv.title);
          }
        )
        .command(
          "append <ref>",
          "Append HTML content to a page's body (accepts path, page ID, or OneNote URL). WARNING: the Microsoft Graph append API re-parses the whole page and may reformat or strip styles from existing content.",
          (y) =>
            y
              .positional("ref", { type: "string", demandOption: true })
              .option("content", { type: "string", alias: "c", demandOption: true, describe: "HTML content to append" })
              .option("md", { type: "boolean", describe: "Treat --content as Markdown and convert it to HTML" }),
          async (argv) => {
            await graph.appendToPage(
              normalizeRef(argv.ref as string)!,
              renderHtmlContent(argv.content as string, argv.md as boolean | undefined)
            );
            console.log("Appended to page.");
          }
        )
        .command(
          "update <ref>",
          "Apply a raw Graph page PATCH command (accepts path, page ID, or OneNote URL)",
          (y) =>
            y
              .positional("ref", { type: "string", demandOption: true })
              .option("target", {
                type: "string",
                alias: "t",
                demandOption: true,
                describe: 'Target selector: "body", "title", or "#element-id"',
              })
              .option("action", {
                type: "string",
                alias: "a",
                choices: pageUpdateActions,
                demandOption: true,
                describe: "Patch action",
              })
              .option("position", {
                type: "string",
                alias: "p",
                choices: pageInsertPositions,
                describe: "Required when --action=insert",
              })
              .option("content", {
                type: "string",
                alias: "c",
                demandOption: true,
                describe: "HTML content to apply",
              })
              .option("md", { type: "boolean", describe: "Treat --content as Markdown and convert it to HTML" })
              .option("sha", {
                type: "string",
                describe: "4-char content sha from 'onenote read' (required for --action=replace)",
              })
              .check((argv) => {
                if (argv.action === "insert" && !argv.position) {
                  throw new Error("--position is required when --action=insert.");
                }
                if (argv.action !== "insert" && argv.position) {
                  throw new Error("--position can only be used when --action=insert.");
                }
                if (argv.md && argv.target === "title") {
                  throw new Error("--md cannot be used with --target=title (titles are plain text).");
                }
                return true;
              }),
          async (argv) => {
            const ref = normalizeRef(argv.ref as string)!;
            if (argv.action === "replace") {
              const { confirmed } = await confirmPageSha(ref, argv.sha as string | undefined, "replace");
              if (!confirmed) return;
            }
            const command: graph.PageUpdateCommand = {
              target: argv.target as string,
              action: argv.action as graph.PageUpdateCommand["action"],
              content: renderHtmlContent(argv.content as string, argv.md as boolean | undefined),
            };
            if (argv.position) {
              command.position = argv.position as graph.PageUpdateCommand["position"];
            }
            await graph.updatePage(ref, [command]);
            console.log(green("Page updated."));
          }
        )
        .demandCommand(1),
    () => {}
  )

  // --- rm (delete page by ref; requires --sha confirmation) ---
  .command(
    ["rm <ref>", "delete <ref>"],
    "Delete a page. Without --sha, dry-runs and prints content + sha.",
    (y) =>
      y
        .positional("ref", { type: "string", demandOption: true })
        .option("sha", {
          type: "string",
          describe: "4-char content sha from 'onenote read' to confirm deletion",
        }),
    async (argv) => {
      const ref = normalizeRef(argv.ref as string)!;
      const { confirmed } = await confirmPageSha(ref, argv.sha as string | undefined, "delete");
      if (!confirmed) return;
      await graph.deletePage(ref);
      console.log(green("Page deleted."));
    }
  )

  // --- cp (copy section into another notebook; non-destructive) ---
  .command(
    "cp <src> <dst>",
    "Copy a section into another notebook (non-destructive; source remains)",
    (y) =>
      y
        .positional("src", { type: "string", demandOption: true, describe: "Source section (path/ID/URL)" })
        .positional("dst", { type: "string", demandOption: true, describe: "Destination notebook (path/ID/URL)" }),
    async (argv) => {
      const src = normalizeRef(argv.src as string)!;
      const dst = normalizeRef(argv.dst as string)!;
      const { operationUrl } = await graph.copySectionToNotebook(src, dst);
      console.log(dim(`Operation: ${operationUrl}`));
      const result = await graph.waitForOperation(operationUrl, {
        onProgress: (s) => console.log(dim(`  status: ${s}`)),
      });
      if (result.status === "failed") {
        throw new Error(`Copy failed: ${JSON.stringify(result.error)}`);
      }
      console.log(green("Copy completed."));
      if (result.resourceLocation) console.log(`New section: ${result.resourceLocation}`);
      console.log(
        dim("Source section was NOT deleted. Verify the copy, then 'onenote rm' the source manually if desired.")
      );
    }
  )

  // --- rename (by ref; depth-dispatched) ---
  .command(
    "rename <ref> <name>",
    "Rename a notebook, section, or page (depth inferred from path)",
    (y) =>
      y
        .positional("ref", { type: "string", demandOption: true })
        .positional("name", { type: "string", demandOption: true }),
    async (argv) => {
      const ref = normalizeRef(argv.ref as string)!;
      const name = argv.name as string;
      const isUrl = ref.includes("://");
      const segments = isUrl ? [] : ref.split("/").filter(Boolean);
      const isGraphId = !isUrl && /^[0-9]-[0-9a-f-]{10,}$/i.test(ref);
      if (isUrl || isGraphId) {
        throw new Error(
          "rename with a raw ID/URL is ambiguous. Use 'notebooks rename', 'sections rename', or 'pages rename' instead."
        );
      }
      if (segments.length === 1) {
        const nb = await graph.renameNotebook(ref, name);
        console.log("Notebook renamed to:", nb.displayName);
      } else if (segments.length === 2) {
        const sec = await graph.renameSection(ref, name);
        console.log("Section renamed to:", sec.displayName);
      } else {
        await graph.renamePage(ref, name);
        console.log("Page renamed to:", name);
      }
    }
  )

  // --- open (launch in browser) ---
  .command(
    "open <ref>",
    "Open a notebook, section, or page in the browser",
    (y) => y.positional("ref", { type: "string", demandOption: true }),
    async (argv) => {
      const ref = normalizeRef(argv.ref as string)!;
      const isUrl = ref.includes("://");
      const segments = isUrl ? [] : ref.split("/").filter(Boolean);
      const isGraphId = !isUrl && /^[0-9]-[0-9a-f-]{10,}$/i.test(ref);
      let url: string | undefined;
      if (isUrl) {
        url = ref;
      } else if (isGraphId) {
        // Page IDs start with "1-"; notebook/section IDs start with "0-" (ambiguous).
        if (ref.startsWith("1-")) {
          const page = await graph.getPage(ref);
          url = page?.links?.oneNoteWebUrl?.href ?? page?.links?.oneNoteClientUrl?.href;
        } else {
          throw new Error(
            "open with a raw 0-* ID is ambiguous (notebook or section). Use 'notebooks get' / 'sections get' to fetch the URL, or pass a path."
          );
        }
      } else if (segments.length === 1) {
        const nb = await graph.getNotebook(ref);
        url = nb?.links?.oneNoteWebUrl?.href ?? nb?.links?.oneNoteClientUrl?.href;
      } else if (segments.length === 2) {
        const sec = await graph.getSection(ref);
        url = sec?.links?.oneNoteWebUrl?.href ?? sec?.links?.oneNoteClientUrl?.href;
      } else {
        const page = await graph.getPage(ref);
        url = page?.links?.oneNoteWebUrl?.href ?? page?.links?.oneNoteClientUrl?.href;
      }
      if (!url) throw new Error(`Could not resolve URL for '${ref}'.`);
      const openerArgs =
        process.platform === "darwin"
          ? ["open", url]
          : process.platform === "win32"
            ? ["cmd", "/c", "start", "", url]
            : ["xdg-open", url];
      const proc = Bun.spawn(openerArgs, { stdout: "inherit", stderr: "inherit" });
      await proc.exited;
      console.log(url);
    }
  )

  // --- read ---
  .command(
    "read <url>",
    "Read a OneNote page, section, or notebook by path (nb/sec/page), ID, or URL",
    (y) =>
      y
        .positional("url", { type: "string", demandOption: true })
        .option("html", { type: "boolean", describe: "Output raw HTML instead of text" }),
    async (argv) => {
      const url = (argv.url as string).replace(/^["']|["']$/g, "");
      const result = await graph.readOneNoteUrl(url, {
        downloadAssets: !argv.html,
      });

      if (argv.html && result.html) {
        console.log(result.html);
        if (result.type === "page") console.error(dim(`sha: ${contentSha4(result.html)}`));
        return;
      }

      console.log(bold(cyan(result.title)));
      if (result.breadcrumb) console.log(dim(result.breadcrumb));
      console.log(dim("─".repeat(Math.min(result.title.length + 10, 60))));
      console.log(result.content);
      if (result.type === "page" && result.html) {
        console.error(dim(`sha: ${contentSha4(result.html)}`));
      }
    }
  )

  // --- sync ---
  .command(
    "sync",
    "Download and cache all OneNote sections for local search",
    () => {},
    async () => {
      const { runSyncUI } = await import("./sync-ui");
      await runSyncUI((emit) => syncCache(() => {}, emit));
    }
  )

  // --- reindex ---
  .command(
    "reindex",
    "Rebuild FTS search index from existing cache (no network needed)",
    () => {},
    async () => {
      await rebuildSearchIndex(console.log);
    }
  )

  // --- search ---
  .command(
    "search <query>",
    "Full-text search across OneNote pages (uses local cache)",
    (y) =>
      y
        .positional("query", { type: "string", demandOption: true })
        .option("online", { type: "boolean", alias: "o", describe: "Use online Graph Search API (section-level)" })
        .option("notebook", { type: "string", alias: "n", describe: "Limit to a notebook name" })
        .option("section", { type: "string", alias: "s", describe: "Limit to a section name" })
        .option("limit", { type: "number", alias: "l", default: 100, describe: "Max results per page" })
        .option("offset", { type: "number", alias: "p", default: 0, describe: "Skip first N results (for pagination)" })
        .option("account", { type: "string", alias: "a", describe: "Limit to a specific account email (e.g. snomiao@gmail.com)" })
        .option("group", { type: "boolean", alias: "g", describe: "Group results hierarchically by account > notebook > section" })
        .option("compact", { type: "boolean", alias: "c", describe: "One line per page: [count] title — section | notebook" })
        .option("max-lines", { type: "number", alias: "m", default: 5, describe: "Max tag lines to show per page (default 5)" })
        .option("since", { type: "string", describe: "Only pages modified after this date (YYYY-MM-DD or '30d' / '3m' / '1y')" })
        .epilog(
          [
            "Query syntax:",
            "  term1 term2       pages containing both terms (AND)",
            "  term1 OR term2    pages containing either term",
            "  term1 NOT term2   pages with term1 but not term2",
            "  \"phrase\"          exact phrase match",
            "  #todo             pages with unchecked checkboxes (requires: onenote sync)",
            "  #done             pages with completed (checked) checkboxes",
            "  #checkbox         pages with any checkbox (#todo OR #done)",
            "  #star             pages tagged with Star",
            "  #question         pages tagged with Question",
            "  #important        pages tagged Important",
            "  #critical         pages tagged Critical",
            "  #idea             pages tagged Idea",
            "  #contact          pages tagged Contact",
            "  #definition       pages tagged Definition",
            "  #highlight        pages with Highlight tag",
            "  #remember         pages tagged Remember for Later",
            "  #book #music #movie #website #phone #address #password",
            "  #meeting #email #callback #discuss #priority1 #priority2 #client",
            "  tag:<name>        alias for #<name> (e.g. tag:star = #star)",
            "",
            "Examples:",
            "  onenote search meeting",
            "  onenote search \"project plan\" --notebook Work",
            "  onenote search \"python OR javascript\"",
            "  onenote search \"#todo buy OR groceries\"",
            "  onenote search \"#todo\" --notebook Work",
            "  onenote search \"#done\"",
            "  onenote search meeting --online",
            "  onenote search meeting --limit 20 --offset 20",
          ].join("\n")
        ),
    async (argv) => {
      if (argv.online) {
        const results = await graph.searchPages(argv.query as string);
        if (!results || results.length === 0) {
          console.log("No results found.");
          return;
        }
        for (const r of results) {
          const heading = r.notebook ? `${r.section} (${r.notebook})` : r.section;
          console.log(`# ${heading}`);
          if (r.summary) console.log(`  ${r.summary}`);
          if (r.url) console.log(`  ${r.url}`);
          console.log();
        }
        console.log(`${results.length} results found.`);
        return;
      }

      // Auto-sync if cache is empty
      if (await isCacheEmpty()) {
        console.log("Cache is empty. Syncing notebooks...");
        await syncCache();
      }

      // Local page-level search
      // Strip surrounding quotes from query (shell may pass "word" or 'word')
      const query = (argv.query as string).replace(/^["']|["']$/g, "");
      const limit = argv.limit as number;
      const offset = argv.offset as number;
      const nbFilter = normalizeRef(argv.notebook as string | undefined);
      const secFilter = normalizeRef(argv.section as string | undefined);
      const accountFilter = normalizeRef(argv.account as string | undefined);
      const groupMode = argv.group as boolean | undefined;
      const compactMode = argv.compact as boolean | undefined;
      const maxLines = (argv["max-lines"] as number | undefined) ?? 5;

      // Parse --since: accepts "YYYY-MM-DD", "30d", "3m", "1y", or full ISO
      const parseSince = (s: string | undefined): string | undefined => {
        if (!s) return undefined;
        const rel = s.match(/^(\d+)([dmy])$/i);
        if (rel) {
          const n = parseInt(rel[1], 10);
          const unit = rel[2].toLowerCase();
          const d = new Date();
          if (unit === "d") d.setDate(d.getDate() - n);
          else if (unit === "m") d.setMonth(d.getMonth() - n);
          else if (unit === "y") d.setFullYear(d.getFullYear() - n);
          return d.toISOString();
        }
        // YYYY-MM-DD → to ISO start of day
        if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return new Date(s + "T00:00:00Z").toISOString();
        return s; // assume already ISO
      };
      const sinceIso = parseSince(argv.since as string | undefined);
      const results = await searchLocal(query, { offset, limit, notebook: nbFilter, section: secFilter, account: accountFilter, since: sinceIso });
      if (results.length === 0) {
        console.log("No results found.");
        return;
      }
      // Strip binary noise, then pick lines with sufficient real content for display.
      // Higher threshold than indexing: requires ≥50% quality chars (ASCII printable or CJK).
      const cleanSnippet = (s: string): string => {
        const lines = s
          .replace(/[\u0000-\u001F\u007F-\u009F]/g, " ")
          .split(/[\n\r]+/)
          .map((l) => l.replace(/\s+/g, " ").trim())
          .filter((l) => {
            if (l.length < 3) return false;
            const ascii = (l.match(/[\x20-\x7E]/g) ?? []).length;
            const cjk = (l.match(/[\u4E00-\u9FFF\u3000-\u30FF\uAC00-\uD7AF]/g) ?? []).length;
            const quality = ascii + cjk;
            if (quality / l.length < 0.5) return false;
            // Require either decent ASCII ratio or meaningful CJK (≥3 chars)
            return ascii / l.length >= 0.4 || (cjk >= 3 && cjk / l.length >= 0.4);
          });
        return lines.join(" ").replace(/\s+/g, " ").trim();
      };

      const { ftsQuery, hasTodo, hasDone, hasCheckbox, tagFilters } = parseTagsFromQuery(query);

      // Tag display prefix: emoji per filter active
      const TAG_EMOJI: Record<string, string> = {
        star: "★", question: "?", important: "❗", critical: "🔴", definition: "📖",
        idea: "💡", contact: "👤", address: "🏠", "phone-number": "📞",
        "web-site-to-visit": "🌐", password: "🔑", "remember-for-later": "🔔",
        "book-to-read": "📚", "music-to-listen-to": "🎵", "movie-to-see": "🎬",
        highlight: "🖍", "schedule-meeting": "📅", "send-in-email": "📧",
        "call-back": "📲", "to-do-priority-1": "P1", "to-do-priority-2": "P2",
        "client-request": "👔",
      };
      const tagPrefix = (() => {
        if (hasTodo) return yellow(bold("☐ "));
        if (hasDone) return green(bold("☑ "));
        if (hasCheckbox) return yellow(bold("☐")) + green(bold("☑ "));
        if (tagFilters.length > 0) {
          const icons = [...new Set(tagFilters.map((t) => TAG_EMOJI[t] ?? t))].join("");
          return bold(icons + " ");
        }
        return "";
      })();

      const lowerQuery = ftsQuery.toLowerCase();
      const cleanResults = results.filter((r) => {
        if (/^\.[a-z0-9]{2,5}$/i.test(r.title.trim())) return false;
        const printable = r.title.replace(/[^\x20-\x7E\u3000-\u30FF\u4E00-\u9FFF\uAC00-\uD7AF\uFF00-\uFFEF\u00A0-\u024F]/g, "");
        return printable.length >= 3 && printable.length / r.title.length >= 0.4;
      });

      type R = typeof results[0];
      const getMatchingLines = (r: R) => {
        const tagLines = r.tagLines ?? [];
        if (hasTodo) return tagLines.filter((l) => l.tag === "to-do");
        if (hasDone) return tagLines.filter((l) => l.tag === "to-do:completed");
        if (hasCheckbox) return tagLines.filter((l) => l.tag === "to-do" || l.tag === "to-do:completed");
        if (tagFilters.length > 0) return tagLines.filter((l) => tagFilters.includes(l.tag));
        return [];
      };

      const fmtDate = (iso: string | null | undefined): string => {
        if (!iso) return "";
        const d = new Date(iso);
        if (isNaN(d.getTime())) return "";
        return d.toISOString().slice(0, 10);
      };

      const renderPageBody = (r: R, indent: string = "  ") => {
        const matchingLines = getMatchingLines(r);
        const body = cleanSnippet(r.body);
        const idx = lowerQuery ? body.toLowerCase().indexOf(lowerQuery) : -1;
        if (idx >= 0) {
          const start = Math.max(0, idx - 40);
          const end = Math.min(body.length, idx + ftsQuery.length + 80);
          const before = body.slice(start, idx);
          const match = body.slice(idx, idx + ftsQuery.length);
          const after = body.slice(idx + ftsQuery.length, end);
          const snippet = (start > 0 ? "..." : "") + before + yellow(bold(match)) + after + (end < body.length ? "..." : "");
          console.log(`${indent}${tagPrefix}${snippet}`);
        } else if (matchingLines.length > 0) {
          for (const line of matchingLines.slice(0, maxLines)) {
            const lineIcon = line.tag === "to-do" ? yellow(bold("☐")) :
                             line.tag === "to-do:completed" ? green(bold("☑")) :
                             bold(TAG_EMOJI[line.tag] ?? line.tag);
            console.log(`${indent}${lineIcon} ${line.text}`);
          }
          if (matchingLines.length > maxLines) console.log(`${indent}${dim(`... and ${matchingLines.length - maxLines} more`)}`);
        } else {
          const snippet = body.length >= 20 ? body.slice(0, 120) : null;
          if (tagPrefix) {
            const label = hasTodo ? "(unchecked items)" : hasDone ? "(completed items)" : hasCheckbox ? "(checkbox items)" : `(${tagFilters.join(",")})`;
            console.log(`${indent}${tagPrefix}${snippet ?? label}`);
          } else if (snippet) {
            console.log(`${indent}${snippet}`);
          }
        }
      };

      // Show account info only when results span multiple accounts (null counts as distinct)
      const uniqueAccounts = new Set(cleanResults.map((r) => r.account ?? ""));
      const showAccounts = uniqueAccounts.size > 1;

      if (groupMode) {
        const byAccount = new Map<string, Map<string, Map<string, R[]>>>();
        for (const r of cleanResults) {
          const acct = r.account ?? "(local)";
          if (!byAccount.has(acct)) byAccount.set(acct, new Map());
          const nbMap = byAccount.get(acct)!;
          if (!nbMap.has(r.notebook)) nbMap.set(r.notebook, new Map());
          const secMap = nbMap.get(r.notebook)!;
          if (!secMap.has(r.section)) secMap.set(r.section, []);
          secMap.get(r.section)!.push(r);
        }

        const renderNotebook = (nbMap: Map<string, Map<string, R[]>>, nbIndent: string, pageIndent: string, bodyIndent: string) => {
          for (const [nb, secMap] of nbMap) {
            console.log(`${nbIndent}${bold(nb)}`);
            for (const [sec, pages] of secMap) {
              console.log(`${nbIndent}  ${dim(sec)}`);
              for (const r of pages) {
                const title = r.webUrl ? link(r.webUrl, bold(r.title)) : bold(r.title);
                console.log(`${pageIndent}${title}`);
                renderPageBody(r, bodyIndent);
              }
            }
          }
        };

        if (showAccounts) {
          for (const [acct, nbMap] of byAccount) {
            const acctTotal = [...nbMap.values()].reduce((s, m) => s + [...m.values()].reduce((a, b) => a + b.length, 0), 0);
            console.log(bold(cyan(`📘 ${acct}`)) + dim(`  (${acctTotal} page${acctTotal === 1 ? "" : "s"})`));
            renderNotebook(nbMap, "  ", "      ", "        ");
            console.log();
          }
        } else {
          // Single account (or all null) — skip account header, flatten one level
          const merged = new Map<string, Map<string, R[]>>();
          for (const nbMap of byAccount.values()) {
            for (const [nb, secMap] of nbMap) {
              if (!merged.has(nb)) merged.set(nb, new Map());
              const m = merged.get(nb)!;
              for (const [sec, pages] of secMap) {
                if (!m.has(sec)) m.set(sec, []);
                m.get(sec)!.push(...pages);
              }
            }
          }
          renderNotebook(merged, "", "    ", "      ");
        }
      } else if (compactMode) {
        // Compact: one line per page, [count] title — section | notebook [date]
        for (const r of cleanResults) {
          const matchingLines = getMatchingLines(r);
          const countStr = matchingLines.length > 0 ? `${yellow(bold(`[${matchingLines.length}]`))} ` : "";
          const titleStr = r.webUrl ? link(r.webUrl, bold(cyan(r.title))) : bold(cyan(r.title));
          const locStr = dim(`${r.section} | ${r.notebook}`);
          const acctStr = showAccounts && r.account ? dim(` | ${r.account}`) : "";
          const dateStr = r.lastModified ? dim(` [${fmtDate(r.lastModified)}]`) : "";
          console.log(`${countStr}${titleStr} ${dim("—")} ${locStr}${acctStr}${dateStr}`);
        }
      } else {
        // Flat: collapse consecutive same (account, notebook, section) header into one line
        let lastHeader = "";
        for (const r of cleanResults) {
          const header = `${r.section}|${r.notebook}|${r.account ?? ""}`;
          if (header !== lastHeader) {
            if (lastHeader !== "") console.log();
            const acctStr = showAccounts && r.account ? ` ${dim("|")} ${dim(r.account)}` : "";
            console.log(`${bold(r.notebook)} ${dim("›")} ${dim(r.section)}${acctStr}`);
            lastHeader = header;
          }
          const displayTitle = r.webUrl ? link(r.webUrl, bold(cyan(r.title))) : bold(cyan(r.title));
          const dateStr = r.lastModified ? dim(` [${fmtDate(r.lastModified)}]`) : "";
          console.log(`  ${displayTitle}${dateStr}`);
          renderPageBody(r, "    ");
        }
      }
      let syncedStr = "";
      try {
        const s = await stat(SEARCH_DB_PATH);
        syncedStr = ` • last synced: ${s.mtime.toLocaleString()}`;
      } catch {}
      const pageInfo = offset > 0 ? ` (offset ${offset})` : results.length === limit ? ` • use --offset ${offset + limit} for more` : "";
      console.log(green(`${results.length} page-level results found${pageInfo}${syncedStr}.`));
    }
  )

  // --- init ---
  .command(
    "init",
    "First-run setup: verify client ID and login",
    () => {},
    async () => {
      // getAccessToken reads ONENOTE_CLIENT_ID from env OR ~/.onenote-cli/config.json
      // and prints a setup guide if the clientId is missing/placeholder.
      const { getAccessToken } = await import("./auth");
      await getAccessToken();
      await whoami();
      console.log(green("\nReady. Try: onenote ls"));
    }
  )

  // --- auth ---
  .command(
    "auth",
    "Manage authentication",
    (yargs) =>
      yargs
        .command(
          "login",
          "Login to Microsoft account (device code flow). Run multiple times to add more accounts.",
          () => {},
          async () => {
            const existing = await listAccounts();
            if (existing.length > 0) {
              console.log(`Already logged in (${existing.length} account${existing.length > 1 ? "s" : ""}):`);
              for (const a of existing) console.log(`  - ${a.username}`);
              console.log("Adding another account...\n");
            }
            const username = await login();
            console.log(`\nLogged in as ${username}`);
          }
        )
        .command(
          "logout [email]",
          "Logout an account. With no args: logout if only one account, else show list. Use --all to logout all.",
          (y) => y
            .positional("email", { type: "string", describe: "Account email to logout" })
            .option("all", { type: "boolean", describe: "Logout all accounts" }),
          async (argv) => {
            if (argv.all) {
              await logoutAll();
            } else {
              await logout(argv.email as string | undefined);
            }
          }
        )
        .command(
          "list",
          "List all logged-in accounts",
          () => {},
          async () => {
            const accounts = await listAccounts();
            if (accounts.length === 0) {
              console.log("No accounts logged in. Run `onenote auth login` to authenticate.");
              return;
            }
            for (const a of accounts) {
              console.log(`${a.username}  (${a.name ?? ""})`);
            }
          }
        )
        .command(
          "whoami",
          "Show all authenticated accounts",
          () => {},
          async () => {
            await whoami();
          }
        )
        .command(
          "setup",
          "Guide for setting up OAuth client credentials",
          () => {},
          async () => {
            console.log(`
=== onenote-cli OAuth Setup ===

1. Register an Azure AD app:
   https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps

   - Click "New registration"
   - Name: onenote-cli
   - Supported account types: "Accounts in any organizational directory and personal Microsoft accounts"
   - Click "Register"

2. Configure authentication:
   - Go to "Authentication" > "Add a platform" > "Mobile and desktop applications"
   - Check: https://login.microsoftonline.com/common/oauth2/nativeclient
   - Enable "Allow public client flows" in Settings tab
   - Click "Save"

3. Add API permissions:
   - Go to "API permissions" > "Add a permission" > "Microsoft Graph" > "Delegated permissions"
   - Add: Notes.Read, Notes.ReadWrite, Notes.ReadWrite.All
   - Click "Add permissions"

4. Copy your Application (client) ID and set it:

   Option A — .env.local:
     ONENOTE_CLIENT_ID=<your-client-id>
     ONENOTE_AUTHORITY=https://login.microsoftonline.com/common

   Option B — ~/.onenote-cli/config.json:
     { "clientId": "<your-client-id>", "authority": "https://login.microsoftonline.com/common" }

5. Login:
   onenote auth login
`);
          }
        )
        .demandCommand(1),
    () => {}
  )

  .epilogue(
    [
      "<ref> accepts any of:",
      "  - path        e.g. 'NotebookA', 'NotebookA/SectionB', 'NotebookA/SectionB/PageC'",
      "  - Graph ID    e.g. '1-abc123...' (page) or '0-abc123...' (notebook/section)",
      "  - OneNote URL e.g. https://onedrive.live.com/redir?...",
      "",
      "Note: path segments must be unique — duplicates throw; rename or use ID/URL.",
    ].join("\n")
  )
  .strict()
  .help()
  .alias("h", "help")
  .version()
  .alias("v", "version")
  .showHelpOnFail(true)
  .fail((msg, err, yargs) => {
    if (err) {
      yargs.showHelp();
      console.error("\nError:", err.message);
      process.exit(1);
    }
    if (msg) {
      yargs.showHelp();
      console.error("\n" + msg);
      process.exit(1);
    }
  })
  .parse();
