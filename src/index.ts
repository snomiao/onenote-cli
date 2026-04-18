#!/usr/bin/env bun
import yargs from "yargs";
import { hideBin } from "yargs/helpers";
import * as graph from "./graph";
import { logout, whoami } from "./auth";
import { syncCache, searchLocal, isCacheEmpty } from "./cache";
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

function toListItem(raw: any): ListItem {
  const url =
    raw?.links?.oneNoteWebUrl?.href
    ?? raw?.links?.oneNoteClientUrl?.href
    ?? raw?.webUrl
    ?? "";
  const name = raw?.displayName ?? raw?.title ?? "(untitled)";
  const date = raw?.lastModifiedDateTime ?? raw?.createdDateTime;
  return { name, url, subtitle: date ? String(date).slice(0, 10) : undefined };
}

function outputList(items: any[], argv: { json?: boolean; limit?: number }) {
  const limited = typeof argv.limit === "number" ? items.slice(0, argv.limit) : items;
  if (argv.json) {
    console.log(JSON.stringify(limited, null, 2));
    return;
  }
  printList(limited.map(toListItem));
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
        printList((notebooks ?? []).map(toListItem));
      } else if (segments.length === 1) {
        const sections = await graph.listSections(path);
        printList((sections ?? []).map(toListItem));
      } else if (segments.length === 2) {
        const pages = await graph.listPages(path);
        printList((pages ?? []).map(toListItem));
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
            outputList(notebooks ?? [], argv);
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
            outputList(sections ?? [], argv);
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
            outputList(pages ?? [], argv);
          }
        )
        .command(
          "get <ref>",
          "Get page metadata (accepts a page ID or a OneNote URL)",
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
              .option("body", { type: "string", alias: "b", describe: "HTML body content" })
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
          "Delete a page (accepts a page ID or a OneNote URL)",
          (y) => y.positional("ref", { type: "string", demandOption: true }),
          async (argv) => {
            await graph.deletePage(normalizeRef(argv.ref as string)!);
            console.log("Page deleted.");
          }
        )
        .command(
          "rename <ref> <title>",
          "Rename a page (accepts a page ID or a OneNote URL)",
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
          "Append HTML content to a page's body (accepts a page ID or a OneNote URL)",
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
          "Apply a raw Graph page PATCH command (accepts a page ID or a OneNote URL)",
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
            const command: graph.PageUpdateCommand = {
              target: argv.target as string,
              action: argv.action as graph.PageUpdateCommand["action"],
              content: renderHtmlContent(argv.content as string, argv.md as boolean | undefined),
            };
            if (argv.position) {
              command.position = argv.position as graph.PageUpdateCommand["position"];
            }
            await graph.updatePage(normalizeRef(argv.ref as string)!, [command]);
            console.log("Page updated.");
          }
        )
        .demandCommand(1),
    () => {}
  )

  // --- rm (delete page by ref) ---
  .command(
    ["rm <ref>", "delete <ref>"],
    "Delete a page (by path, ID, or URL)",
    (y) => y.positional("ref", { type: "string", demandOption: true }),
    async (argv) => {
      await graph.deletePage(normalizeRef(argv.ref as string)!);
      console.log("Page deleted.");
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

  // --- mv (rename by ref; depth-dispatched) ---
  .command(
    ["mv <ref> <name>", "rename <ref> <name>"],
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
        // Raw IDs/URLs — ambiguous shape. Require user to disambiguate via subcommand.
        throw new Error(
          "mv with a raw ID/URL is ambiguous. Use 'notebooks rename', 'sections rename', or 'pages rename' instead."
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
        return;
      }

      console.log(bold(cyan(result.title)));
      if (result.breadcrumb) console.log(dim(result.breadcrumb));
      console.log(dim("─".repeat(Math.min(result.title.length + 10, 60))));
      console.log(result.content);
    }
  )

  // --- sync ---
  .command(
    "sync",
    "Download and cache all OneNote sections for local search",
    () => {},
    async () => {
      await syncCache();
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
        .option("section", { type: "string", alias: "s", describe: "Limit to a section name" }),
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
      const all = await searchLocal(query);
      const nbFilter = normalizeRef(argv.notebook as string | undefined)?.toLowerCase();
      const secFilter = normalizeRef(argv.section as string | undefined)?.toLowerCase();
      const results = all.filter((r) => {
        if (nbFilter && !String(r.notebook ?? "").toLowerCase().includes(nbFilter)) return false;
        if (secFilter && !String(r.section ?? "").toLowerCase().includes(secFilter)) return false;
        return true;
      });
      if (results.length === 0) {
        console.log("No results found.");
        return;
      }
      // Clean snippet text: remove binary noise characters
      const cleanSnippet = (s: string) =>
        s.replace(/[\u0000-\u001F\u007F-\u009F]/g, " ")
          .replace(/[^\x20-\x7E\u00A0-\u024F\u0370-\u058F\u0600-\u06FF\u3000-\u30FF\u3400-\u9FFF\uAC00-\uD7AF\uFF00-\uFFEF\s.,;:!?@#\-_()[\]{}'"/\\=+<>|~`^&*%$\u2000-\u206F]/g, "")
          .replace(/\s{2,}/g, " ")
          .trim();

      const lowerQuery = query.toLowerCase();
      for (const r of results) {
        // Skip results with garbage/attachment titles
        if (/^\.[a-z0-9]{2,5}$/i.test(r.title.trim())) continue;
        const printable = r.title.replace(/[^\x20-\x7E\u3000-\u30FF\u4E00-\u9FFF\uAC00-\uD7AF\uFF00-\uFFEF\u00A0-\u024F]/g, "");
        if (printable.length < 3 || printable.length / r.title.length < 0.4) continue;

        // Title as clickable hyperlink (OSC 8 in TTY, markdown in non-TTY)
        const displayTitle = r.webUrl ? link(r.webUrl, bold(cyan(r.title))) : bold(r.title);
        console.log(displayTitle);
        console.log(`  ${dim(r.section)} ${dim("|")} ${dim(r.notebook)}`);

        // Show context around the match with highlighted keyword
        const body = cleanSnippet(r.body);
        const idx = body.toLowerCase().indexOf(lowerQuery);
        if (idx >= 0) {
          const start = Math.max(0, idx - 40);
          const end = Math.min(body.length, idx + query.length + 80);
          const before = body.slice(start, idx);
          const match = body.slice(idx, idx + query.length);
          const after = body.slice(idx + query.length, end);
          const snippet = (start > 0 ? "..." : "") + before + yellow(bold(match)) + after + (end < body.length ? "..." : "");
          console.log(`  ${snippet}`);
        }
        console.log();
      }
      console.log(green(`${results.length} page-level results found.`));
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
          "Login to Microsoft account (device code flow)",
          () => {},
          async () => {
            const { getAccessToken } = await import("./auth");
            await getAccessToken();
            console.log("Login successful!");
          }
        )
        .command(
          "logout",
          "Clear cached authentication tokens",
          () => {},
          async () => {
            await logout();
          }
        )
        .command(
          "whoami",
          "Show current authenticated user",
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
