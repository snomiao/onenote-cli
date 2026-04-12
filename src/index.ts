#!/usr/bin/env bun
import yargs from "yargs";
import { hideBin } from "yargs/helpers";
import * as graph from "./graph";
import { logout, whoami } from "./auth";
import { syncCache, searchLocal, isCacheEmpty } from "./cache";

const isTTY = process.stdout.isTTY ?? false;

// ANSI color helpers (only when TTY)
const bold = (s: string) => (isTTY ? `\x1b[1m${s}\x1b[0m` : s);
const dim = (s: string) => (isTTY ? `\x1b[2m${s}\x1b[0m` : s);
const yellow = (s: string) => (isTTY ? `\x1b[33m${s}\x1b[0m` : s);
const cyan = (s: string) => (isTTY ? `\x1b[36m${s}\x1b[0m` : s);
const green = (s: string) => (isTTY ? `\x1b[32m${s}\x1b[0m` : s);
const magenta = (s: string) => (isTTY ? `\x1b[35m${s}\x1b[0m` : s);

// OSC 8 hyperlink (clickable in supported terminals), markdown in non-TTY
function link(url: string, text: string): string {
  if (!isTTY) return `[${text}](${url})`;
  return `\x1b]8;;${url}\x1b\\${text}\x1b]8;;\x1b\\`;
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

yargs(hideBin(process.argv))
  .scriptName("onenote")
  .usage("$0 <command> [options]")
  .demandCommand(1)

  // --- notebooks ---
  .command(
    "notebooks",
    "Manage notebooks",
    (yargs) =>
      yargs
        .command(
          "list",
          "List all notebooks",
          () => {},
          async () => {
            const notebooks = await graph.listNotebooks();
            formatTable(notebooks, [
              { key: "id", label: "ID" },
              { key: "displayName", label: "Name" },
              { key: "createdDateTime", label: "Created" },
              { key: "lastModifiedDateTime", label: "Modified" },
            ]);
          }
        )
        .command(
          "get <id>",
          "Get a notebook by ID",
          (y) => y.positional("id", { type: "string", demandOption: true }),
          async (argv) => {
            const nb = await graph.getNotebook(argv.id as string);
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
          "list",
          "List sections",
          (y) => y.option("notebook-id", { type: "string", alias: "n", describe: "Filter by notebook ID" }),
          async (argv) => {
            const sections = await graph.listSections(argv.notebookId as string | undefined);
            formatTable(sections, [
              { key: "id", label: "ID" },
              { key: "displayName", label: "Name" },
              { key: "createdDateTime", label: "Created" },
            ]);
          }
        )
        .command(
          "get <id>",
          "Get a section by ID",
          (y) => y.positional("id", { type: "string", demandOption: true }),
          async (argv) => {
            const section = await graph.getSection(argv.id as string);
            console.log(JSON.stringify(section, null, 2));
          }
        )
        .command(
          "create",
          "Create a new section in a notebook",
          (y) =>
            y
              .option("notebook-id", { type: "string", alias: "n", demandOption: true, describe: "Notebook ID" })
              .option("name", { type: "string", demandOption: true, describe: "Section name" }),
          async (argv) => {
            const section = await graph.createSection(argv.notebookId as string, argv.name as string);
            console.log("Created section:", section.displayName);
            console.log("ID:", section.id);
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
          "list",
          "List section groups",
          (y) => y.option("notebook-id", { type: "string", alias: "n", describe: "Filter by notebook ID" }),
          async (argv) => {
            const groups = await graph.listSectionGroups(argv.notebookId as string | undefined);
            formatTable(groups, [
              { key: "id", label: "ID" },
              { key: "displayName", label: "Name" },
              { key: "createdDateTime", label: "Created" },
            ]);
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
          "list",
          "List pages",
          (y) => y.option("section-id", { type: "string", alias: "s", describe: "Filter by section ID" }),
          async (argv) => {
            const pages = await graph.listPages(argv.sectionId as string | undefined);
            formatTable(pages, [
              { key: "id", label: "ID" },
              { key: "title", label: "Title" },
              { key: "createdDateTime", label: "Created" },
              { key: "lastModifiedDateTime", label: "Modified" },
            ]);
          }
        )
        .command(
          "get <id>",
          "Get page metadata",
          (y) => y.positional("id", { type: "string", demandOption: true }),
          async (argv) => {
            const page = await graph.getPage(argv.id as string);
            console.log(JSON.stringify(page, null, 2));
          }
        )
        .command(
          "content <id>",
          "Get page HTML content",
          (y) => y.positional("id", { type: "string", demandOption: true }),
          async (argv) => {
            const html = await graph.getPageContent(argv.id as string);
            console.log(html);
          }
        )
        .command(
          "create",
          "Create a new page",
          (y) =>
            y
              .option("section-id", { type: "string", alias: "s", demandOption: true, describe: "Section ID" })
              .option("title", { type: "string", alias: "t", demandOption: true, describe: "Page title" })
              .option("body", { type: "string", alias: "b", default: "<p></p>", describe: "HTML body content" }),
          async (argv) => {
            const page = await graph.createPage(
              argv.sectionId as string,
              argv.title as string,
              argv.body as string
            );
            console.log("Created page:", page.title);
            console.log("ID:", page.id);
          }
        )
        .command(
          "delete <id>",
          "Delete a page",
          (y) => y.positional("id", { type: "string", demandOption: true }),
          async (argv) => {
            await graph.deletePage(argv.id as string);
            console.log("Page deleted.");
          }
        )
        .demandCommand(1),
    () => {}
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
        .option("online", { type: "boolean", alias: "o", describe: "Use online Graph Search API (section-level)" }),
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
      const results = await searchLocal(query);
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

  .strict()
  .help()
  .alias("h", "help")
  .version()
  .alias("v", "version")
  .showHelpOnFail(true)
  .fail((msg, err, yargs) => {
    if (err) {
      console.error("Error:", err.message);
      process.exit(1);
    }
    if (msg) {
      yargs.showHelp();
      console.error("\n" + msg);
      process.exit(1);
    }
  })
  .parse();
