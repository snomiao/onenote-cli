#!/usr/bin/env node
// Writes rs/parity.json, the parity cases for onenote (Rust against
// TypeScript), and the index fixtures in rs/parity/db/.
//
//   node rs/parity.gen.mjs           # rewrite both
//   node rs/parity.gen.mjs --check   # fail if they are not what this file writes
//   node rs/parity.gen.mjs --list    # print the case names
//
// Then, from the repository root (after `bun install` and a release build of
// rs/):
//
//   node <rustify>/templates/parity/parity.mjs --cases rs/parity.json \
//     --ts "sh rs/parity/ts.sh" --rs rs/target/release/onenote-cli
//   node <rustify>/templates/parity/parity.mjs --cases rs/parity.json \
//     --rs rs/target/release/onenote-cli --handback ONENOTE_CLI_PACKAGE_ROOT=src/index.ts
//
// The cases are generated because nearly every one needs the same scenery.
// onenote keeps its search index in the directory it is installed in, so a
// case gets a package directory of its own (rs/parity/package.sh; the
// TypeScript is started from it by rs/parity/ts.sh), and most cases get an
// index built into it (rs/parity/make-db.ts, from a file in rs/parity/db/).
// The generated file has no path of this machine in it and is committed.
//
// Nothing here touches Microsoft. The cases run with an empty HOME, so there
// is no client id and no token: every command that needs Graph ends at the
// "Please configure your Azure AD app credentials" notice before it could
// open a connection. `auth login`, `auth logout` and `open` (which starts a
// browser) are not run at all.
//
// A case says who is expected to answer it:
//   add(...)   the Rust binary, alone.
//   back(...)  the JavaScript: the case carries "handback" with the reason.
// The runner's --handback check fails unless exactly the marked ones get
// there.

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const DB_DIR = path.join(HERE, "parity", "db");
const sq = (text) => `'${String(text).replace(/'/g, `'\\''`)}'`;

// ── index fixtures ───────────────────────────────────────────────────────────

const ALICE = "alice@example.com";
const BOB = "bob@contoso.com";
const SHAREPOINT =
  "https://contoso.sharepoint.com/sites/eng/_layouts/15/Doc.aspx?sourcedoc=%7B9f1c2b3a-0000-4000-8000-aaaaaaaaaaaa%7D&file=Meetings.one&action=edit&wd=target%28Meetings.one%7C9f1c2b3a-1111-4111-8111-bbbbbbbbbbbb%2F";
const guid = (n) => `${String(n).padStart(8, "0")}-0000-4000-8000-${String(n).padStart(12, "0")}`;
let nextGuid = 1;
/** A row of the `pages` table, with a page GUID of its own. */
const page = (row) => ({ page_guid: guid(nextGuid++), ...row });

/** An index as a sync of two accounts (and an older one, before accounts) leaves it. */
const notes = [
  page({
    notebook: "Work",
    section: "Meetings",
    account: ALICE,
    title: "Weekly sync",
    body: "Weekly sync\nDiscussed the project plan for Q3 and the migration timeline with the platform team.\nAction items below.\nAlice to draft the plan by Friday.",
    web_url: `${SHAREPOINT}Weekly%20sync%7C11111111-1111-4111-8111-111111111111%2F%29`,
    has_todo: 1,
    has_done: 0,
    tags: "|to-do|",
    tag_lines: [
      { tag: "to-do", text: "Draft the plan" },
      { tag: "to-do", text: "Book the room" },
    ],
    last_modified: "2026-03-01T10:00:00Z",
  }),
  page({
    notebook: "Work",
    section: "Meetings",
    account: ALICE,
    title: "Retro",
    body: "Retro\nWhat went well: shipping on time.\nWhat did not: the plan slipped twice.",
    web_url: "",
    has_todo: 0,
    has_done: 1,
    tags: "|star|to-do:completed|",
    tag_lines: [
      { tag: "to-do:completed", text: "Ship it" },
      { tag: "star", text: "Great work everyone" },
    ],
    last_modified: "2026-02-01T08:30:00.1234567Z",
  }),
  page({
    notebook: "Work",
    section: "Meetings",
    account: ALICE,
    title: "Offsite",
    body: "Offsite\nWe are meeting in the mountains this spring for three days of workshops and the plan is to split into four groups, each owning one theme, and to end every day with a short demo of what the group built so far.",
    web_url: `${SHAREPOINT}Offsite%7C33333333-3333-4333-8333-333333333333%2F%29`,
    has_todo: 0,
    has_done: 0,
    last_modified: "2026-01-15T23:30:00-05:00",
  }),
  page({
    notebook: "Work",
    section: "Projects",
    account: ALICE,
    title: "Roadmap 2026",
    body: '<p>Roadmap &amp; milestones</p>\n<span style="color:red">Plan B &lt;draft&gt;</span> for the launch&nbsp;window',
    web_url: `${SHAREPOINT}Roadmap%202026%7C44444444-4444-4444-8444-444444444444%2F%29`,
    has_todo: 0,
    has_done: 0,
    tags: "|important|question|",
    tag_lines: [
      { tag: "important", text: "Launch window" },
      { tag: "question", text: "Who owns plan B?" },
    ],
    last_modified: "2026-02-20T00:30:00+09:00",
  }),
  // An attachment and a title that is mostly unreadable: counted, not shown.
  page({ notebook: "Work", section: "Projects", account: ALICE, title: ".jpg", body: "plan diagram image", has_todo: 0, has_done: 0, last_modified: "2026-02-20T09:00:00Z" }),
  page({
    notebook: "Work",
    section: "Projects",
    account: ALICE,
    title: "؀؁؂؃؄؅ab",
    body: "plan garbage title",
    has_todo: 0,
    has_done: 0,
    last_modified: "2026-02-20T09:00:00Z",
  }),
  page({
    notebook: "Work",
    section: "Meetings",
    account: BOB,
    title: "1:1 with Alice",
    body: "1:1 with Alice\nPlan review and career chat.",
    // One of the three tags #discuss stands for.
    tags: "|discuss-with-manager|",
    tag_lines: [{ tag: "discuss-with-manager", text: "Promotion timeline" }],
    web_url: "https://contoso-my.sharepoint.com/personal/bob/_layouts/15/Doc.aspx?sourcedoc={55555555-5555-4555-8555-555555555555}&wd=target(Meetings.one|/1:1 with Alice|55555555-0000-4000-8000-000000000000/)",
    has_todo: 0,
    has_done: 0,
    last_modified: "2026-03-02T16:45:00Z",
  }),
  page({
    notebook: "Home",
    section: "Recipes",
    account: BOB,
    title: "Pancakes",
    body: "Pancakes\nFlour, eggs, milk.\nPlan ahead: rest the batter for 30 minutes.",
    web_url: "https://onedrive.live.com/redir?resid=ABC123%21456&page=Edit&wd=target%28Recipes.one%7C%2FPancakes%7C66666666-6666-4666-8666-666666666666%2F%29",
    has_todo: 0,
    has_done: 0,
    tags: "|book-to-read|remember-for-later|",
    tag_lines: [
      { tag: "remember-for-later", text: "Buy maple syrup" },
      { tag: "book-to-read", text: "The Food Lab" },
    ],
    last_modified: "2025-12-24T18:00:00Z",
  }),
  page({
    notebook: "Home",
    section: "Recipes",
    account: BOB,
    title: "Shopping list",
    body: "milk eggs",
    web_url: "https://onedrive.live.com/redir?resid=ABC123%21457",
    has_todo: 1,
    has_done: 1,
    tags: "|to-do|to-do:completed|",
    tag_lines: [
      ...["Milk", "Eggs", "Flour", "Butter", "Maple syrup", "Blueberries", "Yeast"].map((text) => ({ tag: "to-do", text })),
      { tag: "to-do:completed", text: "Sugar" },
    ],
    last_modified: "2026-03-05T07:00:00Z",
  }),
  page({
    notebook: "Home",
    section: "Travel",
    account: BOB,
    title: "Packing",
    body: "Packing\nPassport, charger, adapter and the travel plan printout.",
    has_todo: 0,
    has_done: 0,
    tags: "|discuss-with-manager|discuss-with-person-a|discuss-with-person-b|phone-number|",
    tag_lines: [
      { tag: "discuss-with-person-a", text: "Who waters the plants" },
      { tag: "discuss-with-person-b", text: "Airport pickup" },
      { tag: "discuss-with-manager", text: "Days off" },
      { tag: "phone-number", text: "Hotel: +81 75 000 0000" },
    ],
    last_modified: "2026-01-02T03:04:05Z",
  }),
  // Tagged, but with no tag lines stored (found in the .one file only).
  page({ notebook: "Home", section: "Travel", account: BOB, title: "Checklist", body: "short", has_todo: 1, has_done: 0, last_modified: "2026-01-03T00:00:00Z" }),
  page({
    notebook: "Home",
    section: "Travel",
    account: BOB,
    title: "Ideas",
    body: "Ideas for the next trip: islands, a night train, and a week without any plan at all, just a map and a pair of good shoes and a notebook.\nSecond line.",
    has_todo: 0,
    has_done: 0,
    tags: "|idea|",
    last_modified: "2026-01-04T00:00:00Z",
  }),
  page({ notebook: "Home", section: "Travel", account: BOB, title: "Spark", body: "tiny", has_todo: 0, has_done: 0, tags: "|idea|", last_modified: "2026-01-05T00:00:00Z" }),
  // Rows from before accounts and dates were stored.
  page({ notebook: "Archive", section: "Old", title: "Plan of record 2019", body: "The plan of record." }),
  page({ notebook: "Archive", section: "Old", title: "(untitled)", body: "Loose notes without a plan or a title.", web_url: "https://d.docs.live.net/abc/Archive/Old.one" }),
  page({ notebook: "Archive", section: "Old", title: "Oddments", body: "agent 007 reported 1e3 units, 0x10 in hex, a=b and c++; true or false; ls and search" }),
  // Both words of "project plan", apart.
  page({ notebook: "Archive", section: "Old", title: "Note to self", body: "This project still has no written plan." }),
];

/**
 * A hundred and thirty pages that all match one word, for paging. The word is
 * in the titles only, so a page is one line of output: with stdout a pipe or
 * a socket the TypeScript loses lines once the reader falls behind (bun
 * 1.4.0, see the ledger), and a case has to stay well below where that
 * starts (about 280 lines) to print the same thing every time.
 */
const bulk = Array.from({ length: 130 }, (_, i) =>
  page({
    notebook: "Bulk",
    section: i % 2 ? "Odd" : "Even",
    title: `Bulk keyword page ${String(i + 1).padStart(3, "0")}`,
    body: `filler ${i + 1}`,
  }),
);

/** Pages whose dates are far from any day the cases will run on. */
const ages = [
  page({ notebook: "Ages", section: "S", title: "From 2001", body: "era marker", last_modified: "2001-01-01T00:00:00Z" }),
  page({ notebook: "Ages", section: "S", title: "From 2010", body: "era marker", last_modified: "2010-06-15T12:00:00Z" }),
  page({ notebook: "Ages", section: "S", title: "From 2099", body: "era marker", last_modified: "2099-12-31T23:59:59Z" }),
  page({ notebook: "Ages", section: "S", title: "Undated", body: "era marker" }),
];

const EMOJI = "\u{1F600}";
/** Text that is cut, folded and cleaned in ways where UTF-16 and Unicode show. */
const text = [
  page({ notebook: "Text", section: "Emoji", title: "Cut at the end", body: `0123456789012345678901234567${EMOJI} needle1 ${EMOJI.repeat(40)} tail text after the emoji run` }),
  page({ notebook: "Text", section: "Emoji", title: "Cut at the start", body: `a${EMOJI.repeat(15)} needle2 and what follows it` }),
  page({ notebook: "Text", section: "Folding", title: "Greek", body: "ΟΔΟΣ ΣΟΦΙΑΣ 12, Athens office address" }),
  page({ notebook: "Text", section: "Folding", title: "Turkish", body: `${"İ".repeat(20)} İstanbul gezisi needle3 burada ve devamı da var, uzun bir satır olarak` }),
  page({ notebook: "Text", section: "Spaces", title: "Kinds of space", body: "alpha beta　gamma delta needle4\tend​zero﻿bom" }),
  page({ notebook: "Text", section: "Spaces", title: "Control characters", body: "ctrl\u0001char\u007fdel \u0085nel \u009fapc needle5 \u001fus" }),
  page({ notebook: "Text", section: "Spaces", title: "Carriage returns", body: "line one\r\nneedle11 line two\r\nline three" }),
  page({ notebook: "Text", section: "Scripts", title: "京都旅行", body: "京都 旅行 計画 メモ\n京都に行く計画を立てる\nKyoto trip 計画 notes for spring" }),
  page({ notebook: "Text", section: "Scripts", title: "План", body: "план проекта на третий квартал\nabc план проекта на квартал xyz\nNotes: план (Q3 draft) for the team" }),
  page({ notebook: "Text", section: "Scripts", title: "ＡＢＣ", body: "fullwidth needle13" }),
  // Lines with almost no ASCII in them.
  page({ notebook: "Text", section: "Scripts", title: "Russian notes", body: "заметки проекта на третий квартал" }),
  page({ notebook: "Text", section: "Scripts", title: "Kyoto notes", body: "京都に行く予定です 予定表 を立てるつもりです" }),
  page({
    notebook: "Text",
    section: "Markup",
    title: "Tags and entities",
    body: `needle6 <a href="x">link</a> &amp;amp; &lt;b&gt; <unclosed needle6b\nnext <${"x".repeat(301)}> needle7 after a long tag\n<${"y".repeat(300)}> needle8 after a tag of 300`,
  }),
  page({
    notebook: "Text",
    section: "Markup",
    title: "Plan (v2) [draft] | notes",
    body: "needle14 title with brackets",
    web_url: "https://example.com/a_(b)?x=[1]&y=a|b",
  }),
  page({ notebook: "Text", section: "Titles", title: "ab", body: "needle9 short title" }),
  page({ notebook: "Text", section: "Titles", title: " .png ", body: "needle9 attachment" }),
  page({ notebook: "Text", section: "Titles", title: "Titleonly needle12", body: "" }),
  page({ notebook: "Text", section: "Lines", title: "Best line", body: "план needle10 план план план план\nclean needle10 line here\nanother needle10 line" }),
  page({ notebook: "Text", section: "Lines", title: "Far in", body: `${"lorem ipsum dolor sit amet ".repeat(8)}needle15 ${"consectetur adipiscing elit ".repeat(8)}` }),
  page({ notebook: "École", section: "Français", title: "Leçon", body: "needle16 accents" }),
];

/** Rows no sync writes; each index holds one kind. */
const odd = {
  "null-title": [page({ notebook: "N", section: "S", title: null, body: "oddity here" })],
  "null-body": [page({ notebook: "N", section: "S", title: "Bodiless", body: null, tags: "|star|" })],
  // SQLite turns a number into text on its way into a TEXT column.
  "number-title": [page({ notebook: "N", section: "S", title: { integer: 4242 }, body: "oddity here" })],
  "blob-title": [page({ notebook: "N", section: "S", title: { blob: "426c6f62" }, body: "oddity here" })],
  "blob-body": [page({ notebook: "N", section: "S", title: "Blob", body: { blob: "6f6464697479" }, tags: "|star|" })],
  "tag-lines-object": [page({ notebook: "N", section: "S", title: "Object", body: "oddity here", has_todo: 1, tag_lines: '{"tag":"to-do","text":"x"}' })],
  "tag-lines-broken": [page({ notebook: "N", section: "S", title: "Broken", body: "oddity here", has_todo: 1, tag_lines: "not json" })],
  "tag-lines-no-text": [page({ notebook: "N", section: "S", title: "No text", body: "oddity here", has_todo: 1, tag_lines: '[{"tag":"to-do"}]' })],
  "tag-lines-null": [page({ notebook: "N", section: "S", title: "Null lines", body: "oddity here", has_todo: 1, tag_lines: "null" })],
  "date-in-words": [page({ notebook: "N", section: "S", title: "Worded date", body: "oddity here", last_modified: "March 1, 2026" })],
  "date-unreadable": [page({ notebook: "N", section: "S", title: "No date", body: "oddity here", last_modified: "soon" })],
  "empty-account": [
    page({ notebook: "N", section: "S", title: "Empty account", body: "oddity here", account: "" }),
    page({ notebook: "N", section: "S", title: "No account", body: "oddity here too" }),
    page({ notebook: "N", section: "T", title: "Named account", body: "oddity here as well", account: ALICE }),
  ],
  "bars-in-names": [
    page({ notebook: "c", section: "a|b", title: "First", body: "oddity here" }),
    page({ notebook: "b|c", section: "a", title: "Second", body: "oddity here" }),
  ],
  empty: [],
};

const DBS = { notes, bulk, ages, text, ...Object.fromEntries(Object.entries(odd).map(([name, rows]) => [`odd-${name}`, rows])) };

// ── cases ────────────────────────────────────────────────────────────────────

const cases = [];
const MTIME = "2026-03-08T09:05:07Z";
const ROOT = '"$ONENOTE_CLI_PACKAGE_ROOT"';
// What a case leaves behind: the directory it ran in, HOME, the index and
// what SQLite puts beside it (sizes), the index's modification time, and the
// files of ~/.onenote-cli.
const AFTER = [
  "ls -A <CWD> <HOME>",
  `find ${ROOT}/.onenote <HOME>/.onenote-cli -type f -exec stat -c '%s %n' {} + 2>&1 | sort; stat -c %Y ${ROOT}/.onenote/search.db 2>&1`,
];

/**
 * One case. `db` names the index to start from ("notes", or ["notes",
 * "--wal", "empty"]), `mtime` its modification time, `root` files to put in
 * the package directory, `tz` the TZ to run under (null: none, the system's
 * zone).
 */
function make(name, args, { db, mtime = MTIME, root = {}, tz = "UTC", env = {}, files, setup = [], stdin, after = AFTER, handback } = {}) {
  const lines = ['sh "<REPO>/rs/parity/package.sh"'];
  for (const [file, content] of Object.entries(root)) lines.push(`printf %s ${sq(content)} > ${ROOT}/${file}`);
  if (db) {
    const [which, ...flags] = [].concat(db);
    if (!DBS[which]) throw new Error(`${name}: no index called ${which}`);
    lines.push(`bun "<REPO>/rs/parity/make-db.ts" "<REPO>/rs/parity/db/${which}.json" --mtime ${mtime}${flags.length ? ` ${flags.join(" ")}` : ""}`);
  }
  lines.push(...setup);
  if (cases.some((c) => c.name === name)) throw new Error(`two cases are called ${name}`);
  cases.push({
    name,
    args,
    ...(stdin === undefined ? {} : { stdin }),
    env: { ...(tz === null ? {} : { TZ: tz }), ...env },
    ...(files ? { files } : {}),
    setup: lines,
    after,
    ...(handback ? { handback } : {}),
  });
}
/** A case the Rust binary answers alone. */
const add = (name, args, options) => make(name, args, options);
/** A case that is handed to the JavaScript, and why. */
const back = (name, why, args, options) => make(name, args, { ...options, handback: why });

const search = (...args) => ["search", ...args];
const NOTES = { db: "notes" };

// ── search: queries ──────────────────────────────────────────────────────────

add("search: one word", search("plan"), NOTES);
add("search: a word in capitals", search("PLAN"), NOTES);
add("search: two words that stand together", search("project plan"), NOTES);
add("search: two words that do not", search("plan project"), NOTES);
add("search: a word that is nowhere", search("nosuchword"), NOTES);
add("search: OR", search("pancakes OR retro"), NOTES);
add("search: or in small letters", search("pancakes or retro"), NOTES);
add("search: NOT", search("plan NOT project"), NOTES);
add("search: AND", search("plan AND batter"), NOTES);
add("search: a prefix in an expression", search("road* OR zzz"), NOTES);
add("search: a prefix as a bare word", search("road*"), NOTES);
add("search: a column filter in an expression", search("title:plan OR zzz"), NOTES);
add("search: NEAR in an expression", search("NEAR(project plan, 3) OR zzz"), NOTES);
add("search: brackets in an expression", search("(plan OR retro) NOT project"), NOTES);
add("search: punctuation in a bare word", search("(plan"), NOTES);
add("search: a hyphenated word", search("1:1"), NOTES);
add("search: an apostrophe inside", search("alice's"), NOTES);
add("search: a quote inside", search('pl"an'), NOTES);
add("search: the query in double quotes", search('"plan"'), NOTES);
add("search: the query in single quotes", search("'plan'"), NOTES);
add("search: a phrase in double quotes", search('"project plan"'), NOTES);
add("search: a phrase in two pairs of quotes", search('""project plan""'), NOTES);
add("search: a quote at the front only", search('"plan'), NOTES);
add("search: a quote at the end only", search("plan'"), NOTES);
add("search: a phrase after a word", search('discussed "project plan" with'), NOTES);
add("search: nothing but two quotes lists every page", search('""'), NOTES);
add("search: nothing but one quote lists every page", search('"'), NOTES);
add("search: nothing but blanks lists every page", search("   "), NOTES);
add("search: a number with leading zeros", search("007"), NOTES);
add("search: a number in e notation", search("1e3"), NOTES);
add("search: a hexadecimal number", search("0x10"), NOTES);
add("search: a year", search("2019"), NOTES);
add("search: the word true", search("true"), NOTES);
add("search: the name of a command", search("ls"), NOTES);
add("search: the word search", search("search"), NOTES);
add("search: an equals sign", search("a=b"), NOTES);
add("search: standard input is not read", search("plan"), { ...NOTES, stdin: "retro\n" });
back("search: an expression that ends in an operator", "FTS5 rejects it and the TypeScript reads the section files instead", search("plan AND"), NOTES);
back("search: an operator alone", "FTS5 rejects it and the TypeScript reads the section files instead", search("NOT"), NOTES);
back("search: an unknown column in an expression", "FTS5 rejects it and the TypeScript reads the section files instead", search("nope:plan OR zzz"), NOTES);
back("search: an open quote after a word", "FTS5 rejects it and the TypeScript reads the section files instead", search('retro "plan'), NOTES);

// ── search: options ──────────────────────────────────────────────────────────

add("search options: -n", search("plan", "-n", "Work"), NOTES);
add("search options: --notebook", search("plan", "--notebook", "Home"), NOTES);
add("search options: --notebook=", search("plan", "--notebook=Archive"), NOTES);
add("search options: the option before the query", search("-n", "Work", "plan"), NOTES);
add("search options: a notebook in small letters", search("plan", "-n", "work"), NOTES);
add("search options: part of a notebook's name", search("plan", "-n", "or"), NOTES);
add("search options: _ in a notebook stands for any character", search("plan", "-n", "W_rk"), NOTES);
add("search options: % as a notebook", search("plan", "-n", "%"), NOTES);
add("search options: a notebook in quotes", search("plan", "-n", '"Work"'), NOTES);
add("search options: a notebook in two pairs of quotes", search("plan", "-n", '""Work""'), NOTES);
add("search options: a notebook that is two quotes", search("plan", "-n", '""'), NOTES);
add("search options: a notebook that does not exist", search("plan", "-n", "Nope"), NOTES);
add("search options: a notebook called true", search("plan", "-n", "true"), NOTES);
add("search options: a notebook called 123", search("plan", "--notebook", "123"), NOTES);
add("search options: a notebook value with = in it", search("plan", "--notebook=a=b"), NOTES);
add("search options: a notebook value starting with a dash after =", search("plan", "--notebook=-x"), NOTES);
add("search options: an accented notebook", search("needle16", "-n", "École"), { db: "text" });
add("search options: an accented notebook in the other case", search("needle16", "-n", "éCOLE"), { db: "text" });
add("search options: -s", search("plan", "-s", "Meetings"), NOTES);
add("search options: --section=", search("plan", "--section=Trav"), NOTES);
add("search options: --section", search("plan", "--section", "recipes"), NOTES);
add("search options: notebook and section", search("plan", "-n", "work", "-s", "meet"), NOTES);
add("search options: -a", search("plan", "-a", ALICE), NOTES);
add("search options: --account", search("plan", "--account", BOB), NOTES);
add("search options: --account=", search("plan", `--account=${BOB}`), NOTES);
add("search options: part of an account is not enough", search("plan", "-a", "alice"), NOTES);
add("search options: an account in capitals is another", search("plan", "-a", ALICE.toUpperCase()), NOTES);
add("search options: an account in quotes", search("plan", "-a", `'${ALICE}'`), NOTES);
add("search options: -l", search("plan", "-l", "2"), NOTES);
add("search options: --limit", search("plan", "--limit", "3"), NOTES);
add("search options: --limit=", search("plan", "--limit=1"), NOTES);
add("search options: --limit 0", search("plan", "--limit", "0"), NOTES);
add("search options: a limit that is the number of results", search("plan", "-l", "10"), NOTES);
add("search options: -p", search("plan", "-p", "2"), NOTES);
add("search options: --offset", search("plan", "--offset", "9"), NOTES);
add("search options: --offset=", search("plan", "--offset=0"), NOTES);
add("search options: an offset past the end", search("plan", "-p", "50"), NOTES);
add("search options: limit and offset", search("plan", "-l", "2", "-p", "2"), NOTES);
add("search options: -g with two accounts", search("plan", "-g"), NOTES);
add("search options: --group with one account", search("plan", "--group", "-a", BOB), NOTES);
add("search options: -g with no account", search("plan", "-g", "-n", "Archive"), NOTES);
add("search options: -g before the query", search("-g", "plan"), NOTES);
add("search options: -c", search("plan", "-c"), NOTES);
add("search options: --compact with one account", search("plan", "--compact", "-a", ALICE), NOTES);
add("search options: -g wins over -c", search("plan", "-c", "-g"), NOTES);
add("search options: everything at once", search("-n", "Work", "plan", "--section=Meetings", "-a", ALICE, "-l", "5", "--offset", "0", "-c", "-m", "3", "--since", "2026-01-01"), NOTES);
add("search options: ONENOTE_CACHE_DIR does not matter to an index", search("plan"), { ...NOTES, env: { ONENOTE_CACHE_DIR: "<CWD>/nowhere" } });

// Paging: the last line says how to go on.
add("search paging: the first hundred", search("keyword"), { db: "bulk" });
add("search paging: the rest", search("keyword", "--offset", "100"), { db: "bulk" });
add("search paging: the second fifty", search("keyword", "-l", "50", "-p", "50"), { db: "bulk" });
add("search paging: all of them at once", search("keyword", "-l", "130"), { db: "bulk" });
add("search paging: more than there are", search("keyword", "-l", "999999999"), { db: "bulk" });
add("search paging: three from the third on, compact", search("keyword", "-l", "3", "-p", "2", "-c"), { db: "bulk" });
add("search paging: grouped", search("keyword", "-l", "7", "-g"), { db: "bulk" });
add("search paging: every page without a query", search('""', "-l", "4"), { db: "bulk" });
add("search paging: a number among the pages", search("17"), { db: "bulk" });

// ── search: --since ──────────────────────────────────────────────────────────

add("search since: a day", search("plan", "--since", "2026-02-15"), NOTES);
add("search since: a day, with =", search("plan", "--since=2026-01-01"), NOTES);
add("search since: a day after everything", search("plan", "--since", "2030-01-01"), NOTES);
add("search since: a full time is compared as text", search("plan", "--since", "2026-03-01T10:00:00Z"), NOTES);
add("search since: a word is compared as text", search("plan", "--since", "yesterday"), NOTES);
add("search since: a year alone is compared as text", search("plan", "--since", "2026"), NOTES);
add("search since: with tags", search("#todo", "--since", "2026-03-02"), NOTES);
add("search since: thirty days back", search("era", "--since", "30d"), { db: "ages" });
add("search since: three months back", search("era", "--since", "3m"), { db: "ages" });
add("search since: ten years back", search("era", "--since", "10y"), { db: "ages" });
add("search since: fifty years back", search("era", "--since", "50Y"), { db: "ages" });
add("search since: nine thousand days back", search("era", "--since", "9999D"), { db: "ages" });
add("search since: no days back", search("era", "--since", "0d"), { db: "ages" });
add("search since: months back in New York", search("era", "--since", "13M"), { db: "ages", tz: "America/New_York" });
add("search since: a span with a unit that is none", search("era", "--since", "3w"), { db: "ages" });
back("search since: a day that does not exist", "Date turns February 30th into March 2nd; Rust reads only days that exist", search("plan", "--since", "2026-02-30"), NOTES);
back("search since: a month that does not exist", "an invalid Date makes the TypeScript fail with yargs's help", search("plan", "--since", "2026-13-01"), NOTES);
back("search since: more days back than four digits", "Rust counts back at most 9999 units", search("era", "--since", "99999d"), { db: "ages" });

// ── search: tags ─────────────────────────────────────────────────────────────

add("search tags: #todo", search("#todo"), NOTES);
add("search tags: #done", search("#done"), NOTES);
add("search tags: #checkbox", search("#checkbox"), NOTES);
add("search tags: #star", search("#star"), NOTES);
add("search tags: tag:star", search("tag:star"), NOTES);
add("search tags: tag:todo", search("tag:todo"), NOTES);
add("search tags: #TODO in capitals", search("#TODO"), NOTES);
add("search tags: #todos is a word, not a tag", search("#todos"), NOTES);
add("search tags: an unknown tag is a word", search("#nope"), NOTES);
add("search tags: two tags", search("#question #important"), NOTES);
add("search tags: two tags no page has both of", search("#star #important"), NOTES);
add("search tags: the same tag twice", search("#star #star"), NOTES);
add("search tags: #discuss needs all three of its tags", search("#discuss"), NOTES);
add("search tags: #phone", search("#phone"), NOTES);
add("search tags: #remember and #book", search("#remember #book"), NOTES);
add("search tags: #idea, pages without tag lines", search("#idea"), NOTES);
add("search tags: a tag no page has", search("#priority1"), NOTES);
add("search tags: #todo and #done", search("#todo #done"), NOTES);
add("search tags: a tag and a word", search("#todo plan"), NOTES);
add("search tags: a tag and a word that is in the excerpt", search("#star retro"), NOTES);
add("search tags: a tag and a word with nothing in common", search("#todo pancakes"), NOTES);
add("search tags: a tag inside a word", search("re#todo"), NOTES);
add("search tags: #todo in a notebook", search("#todo", "-n", "Home"), NOTES);
add("search tags: #todo compact", search("#todo", "-c"), NOTES);
add("search tags: #todo grouped", search("#todo", "-g"), NOTES);
add("search tags: #checkbox compact, one account", search("#checkbox", "-c", "-a", BOB), NOTES);
add("search tags: -m 1", search("#todo", "-m", "1"), NOTES);
add("search tags: --max-lines 7", search("#todo", "--max-lines", "7"), NOTES);
add("search tags: --max-lines=0", search("#todo", "--max-lines=0"), NOTES);
add("search tags: #done, one line each", search("#done", "-m", "1"), NOTES);
add("search tags: limit and offset", search("#checkbox", "-l", "1", "-p", "1"), NOTES);

// ── search: text ─────────────────────────────────────────────────────────────

const TEXT = { db: "text" };
add("search text: an emoji cut in half at the end of the excerpt", search("needle1"), TEXT);
add("search text: an emoji cut in half at the start of the excerpt", search("needle2"), TEXT);
add("search text: Greek in capitals", search("ΟΔΟΣ"), TEXT);
add("search text: Greek in small letters", search("οδος"), TEXT);
add("search text: Greek with the other sigma", search("οδοσ"), TEXT);
add("search text: a line that grows when put in small letters", search("needle3"), TEXT);
add("search text: Turkish capital i", search("İstanbul"), TEXT);
add("search text: kinds of space", search("needle4"), TEXT);
add("search text: control characters", search("needle5"), TEXT);
add("search text: carriage returns", search("needle11"), TEXT);
add("search text: Japanese words set apart", search("京都"), TEXT);
add("search text: Japanese in a line with little ASCII", search("計画"), TEXT);
add("search text: Latin beside Japanese", search("kyoto"), TEXT);
add("search text: Cyrillic", search("план"), TEXT);
add("search text: Cyrillic, grouped", search("план", "-g"), TEXT);
add("search text: a Cyrillic line", search("квартал"), TEXT);
add("search text: a Japanese line", search("予定表"), TEXT);
add("search text: a Japanese word inside a longer run is not found", search("予定"), TEXT);
add("search text: a fullwidth title", search("needle13"), TEXT);
add("search text: tags and entities", search("needle6"), TEXT);
add("search text: an unclosed tag", search("needle6b"), TEXT);
add("search text: a tag too long to be one", search("needle7"), TEXT);
add("search text: a tag of exactly 300", search("needle8"), TEXT);
add("search text: brackets in title and address", search("needle14"), TEXT);
add("search text: brackets in title and address, compact", search("needle14", "-c"), TEXT);
add("search text: titles that are not shown", search("needle9"), TEXT);
add("search text: titles that are not shown, compact", search("needle9", "-c"), TEXT);
add("search text: titles that are not shown, grouped", search("needle9", "-g"), TEXT);
add("search text: a hit in the title only", search("needle12"), TEXT);
add("search text: the line with the most ASCII is shown", search("needle10"), TEXT);
add("search text: a hit far into a long line", search("needle15"), TEXT);
add("search text: every page, by first line", search('""'), TEXT);
add("search text: every page, compact", search('""', "-c"), TEXT);
add("search text: every page, grouped", search('""', "-g"), TEXT);

// ── search: the time in the last line ────────────────────────────────────────

const one = (zone, mtime) => ({ db: "notes", tz: zone, mtime });
add("search time: Tokyo", search("retro"), one("Asia/Tokyo", MTIME));
add("search time: New York, the second before summer time", search("retro"), one("America/New_York", "2026-03-08T06:59:59Z"));
add("search time: New York, the first second of summer time", search("retro"), one("America/New_York", "2026-03-08T07:00:00Z"));
add("search time: New York, the hour that comes twice", search("retro"), one("America/New_York", "2026-11-01T06:30:00Z"));
add("search time: Berlin in summer", search("retro"), one("Europe/Berlin", "2026-07-01T21:59:59Z"));
add("search time: Berlin in winter", search("retro"), one("Europe/Berlin", "2026-12-31T23:00:00Z"));
add("search time: Kolkata", search("retro"), one("Asia/Kolkata", "2026-03-08T18:30:00Z"));
add("search time: Pago Pago", search("retro"), one("Pacific/Pago_Pago", "2026-01-01T10:59:59Z"));
add("search time: Sydney", search("retro"), one("Australia/Sydney", "2026-01-01T12:59:59Z"));
add("search time: Kathmandu", search("retro"), one("Asia/Kathmandu", "2026-06-30T18:14:59Z"));
add("search time: US/Pacific, an older name", search("retro"), one("US/Pacific", MTIME));
add("search time: Etc/GMT+5", search("retro"), one("Etc/GMT+5", MTIME));
add("search time: GMT", search("retro"), one("GMT", MTIME));
add("search time: midnight", search("retro"), one("UTC", "2026-01-01T00:00:00Z"));
add("search time: noon", search("retro"), one("UTC", "2026-10-10T12:00:00Z"));
add("search time: the last second of a day", search("retro"), one("UTC", "2026-12-31T23:59:59Z"));
add("search time: the last century", search("retro"), one("UTC", "1999-12-31T12:34:56Z"));
add("search time: a far year", search("retro"), one("UTC", "2099-02-28T01:02:03Z"));
add("search time: the system's zone", search("retro"), one(null, MTIME));
add("search time: the system's zone, grouped", search("plan", "-g"), one(null, "2026-08-08T20:08:08Z"));
back("search time: TZ is a rule", "the runtime reads rules, paths and other spellings its own way; Rust takes zone names only", search("retro"), one("JST-9", MTIME));
back("search time: TZ with a colon", "the runtime reads rules, paths and other spellings its own way; Rust takes zone names only", search("retro"), one(":Asia/Tokyo", MTIME));
back("search time: TZ in small letters", "the runtime reads rules, paths and other spellings its own way; Rust takes zone names only", search("retro"), one("asia/tokyo", MTIME));
back("search time: TZ is a path", "the runtime reads rules, paths and other spellings its own way; Rust takes zone names only", search("retro"), one("/usr/share/zoneinfo/Asia/Tokyo", MTIME));
back("search time: TZ names no zone", "the runtime reads rules, paths and other spellings its own way; Rust takes zone names only", search("retro"), one("Nowhere/Land", MTIME));
back("search time: TZ is empty", "the runtime reads rules, paths and other spellings its own way; Rust takes zone names only", search("retro"), one("", MTIME));

// ── search: the state of the machine ─────────────────────────────────────────

add("search state: an index as a finished sync leaves it", search("plan"), { db: ["notes", "--wal", "empty"] });
add("search state: an index whose rows are still in the write-ahead log", search("plan"), { db: ["notes", "--wal", "pending"] });
add("search state: tags from the write-ahead log", search("#todo", "-c"), { db: ["notes", "--wal", "pending"] });
add("search state: a .env here that has nothing to say about it", search("retro"), {
  ...NOTES,
  files: { ".env": "# local settings\nFOO=bar\nexport API_URL='https://example.com'\nONENOTE_CACHE_DIR=/nonexistent\n" },
});
add("search state: .env.local and .env.development here", search("retro"), {
  ...NOTES,
  files: { ".env.local": "ONENOTE_CLIENT_ID=00000000-0000-0000-0000-000000000000\n", ".env.development": "DEBUG=1\n" },
});
add("search state: the package's own .env.local", search("retro"), {
  ...NOTES,
  root: { ".env.local": "ONENOTE_CLIENT_ID=00000000-0000-0000-0000-000000000000\nONENOTE_AUTHORITY=https://login.microsoftonline.com/common\n" },
});
add("search state: the package's .env moves the cache", search("retro"), { ...NOTES, root: { ".env": "ONENOTE_CACHE_DIR=/nonexistent\n" } });
back("search state: a .env here that sets TZ", "bun loads it, and TZ changes the time in the last line", search("retro"), { ...NOTES, files: { ".env": "TZ=Asia/Tokyo\n" } });
back("search state: a .env.local here that names NODE_ENV", "a variable that may change what the runtime does", search("retro"), {
  ...NOTES,
  files: { ".env.local": "NODE_ENV=production\n" },
});
back("search state: a .env here with tz inside another name", "the letters TZ anywhere in the file are enough to step aside", search("retro"), {
  ...NOTES,
  files: { ".env": "QUARTZ_HOST=localhost\n" },
});
back("search state: a .env.production here that sets TZ", "every .env file bun might load is looked at, whatever NODE_ENV is", search("retro"), {
  ...NOTES,
  files: { ".env.production": "TZ=Asia/Tokyo\n" },
});
back("search state: the package's .env.local sets TZ", "src/auth.ts loads it, and TZ changes the time in the last line", search("retro"), {
  ...NOTES,
  root: { ".env.local": "TZ=Asia/Tokyo\n" },
});
back("search state: the package's .env sets TZ", "src/auth.ts loads it, and TZ changes the time in the last line", search("retro"), { ...NOTES, root: { ".env": "TZ = 'Asia/Tokyo'\n" } });
back("search state: a bunfig.toml here", "bun reads it; it can preload scripts", search("retro"), { ...NOTES, files: { "bunfig.toml": "# nothing\n" } });
back("search state: a .bunfig.toml at home", "bun reads it; it can preload scripts", search("retro"), { ...NOTES, setup: ["printf '# nothing\\n' > <HOME>/.bunfig.toml"] });
back("search state: a .bunfig.toml in the configuration directory", "bun reads it; it can preload scripts", search("retro"), {
  ...NOTES,
  setup: ["mkdir -p <HOME>/.config && printf '# nothing\\n' > <HOME>/.config/.bunfig.toml"],
});
back("search state: NODE_ENV is set", "it changes which .env files bun loads", search("retro"), { ...NOTES, env: { NODE_ENV: "production" } });
back("search state: BUN_OPTIONS is set", "options for the runtime", search("retro"), { ...NOTES, env: { BUN_OPTIONS: "--smol" } });
back("search state: NODE_OPTIONS is set", "options for the runtime", search("retro"), { ...NOTES, env: { NODE_OPTIONS: "--no-deprecation" } });
back("search state: ONENOTE_CLI_NO_NATIVE=1", "the user's way out", search("retro"), { ...NOTES, env: { ONENOTE_CLI_NO_NATIVE: "1" } });
back("search state: nothing synced yet", "the TypeScript starts a sync, which asks for credentials", search("plan"));
back("search state: an index without pages", "the TypeScript looks through the cache directory and starts a sync", search("plan"), { db: "odd-empty" });
const SECTION_FILE = JSON.stringify({
  section: "Meetings",
  notebook: "Work",
  webUrl: "https://contoso.sharepoint.com/sites/eng/Meetings.one",
  pages: [
    { title: "Weekly sync", body: "Discussed the project plan for Q3.", pageGuid: "11111111-1111-4111-8111-111111111111" },
    { title: "Retro", body: "What went well." },
  ],
});
back("search state: section files but no index", "without an index the TypeScript reads the section files", search("plan"), {
  env: { ONENOTE_CACHE_DIR: "<CWD>/cache" },
  files: { "cache/Work/Meetings.json": SECTION_FILE },
});
back("search state: section files and an index without pages", "an index without pages is not searched by Rust", search("plan"), {
  db: "odd-empty",
  env: { ONENOTE_CACHE_DIR: "<CWD>/cache" },
  files: { "cache/Work/Meetings.json": SECTION_FILE },
});
back("search state: the index is not a database", "SQLite cannot read it and the TypeScript reads the section files instead", search("plan"), {
  setup: [`mkdir -p ${ROOT}/.onenote && printf 'not a database\\n' > ${ROOT}/.onenote/search.db && touch -d ${MTIME} ${ROOT}/.onenote/search.db`],
});
back("search state: an index from before the full-text table", "SQLite reports the missing table and the TypeScript reads the section files instead", search("plan"), {
  ...NOTES,
  setup: [`bun -e 'import { Database } from "bun:sqlite"; const db = new Database(process.env.ONENOTE_CLI_PACKAGE_ROOT + "/.onenote/search.db"); db.run("DROP TABLE pages_fts"); db.close(true);' && touch -d ${MTIME} ${ROOT}/.onenote/search.db && rm -f ${ROOT}/.onenote/search.db-wal ${ROOT}/.onenote/search.db-shm`],
});
add("search state: tags need no full-text table", search("#todo", "-c"), {
  ...NOTES,
  setup: [`bun -e 'import { Database } from "bun:sqlite"; const db = new Database(process.env.ONENOTE_CLI_PACKAGE_ROOT + "/.onenote/search.db"); db.run("DROP TABLE pages_fts"); db.close(true);' && touch -d ${MTIME} ${ROOT}/.onenote/search.db && rm -f ${ROOT}/.onenote/search.db-wal ${ROOT}/.onenote/search.db-shm`],
});

// Rows no sync writes.
back("search rows: a title that is NULL", "the TypeScript fails on it with yargs's help", search("oddity"), { db: "odd-null-title" });
add("search rows: a title stored from a number is text", search("oddity"), { db: "odd-number-title" });
back("search rows: a title that is a blob", "the TypeScript fails on it with yargs's help", search("oddity"), { db: "odd-blob-title" });
back("search rows: a body that is NULL", "Rust reads text only", search("#star"), { db: "odd-null-body" });
back("search rows: a body that is a blob", "Rust reads text only", search("#star"), { db: "odd-blob-body" });
back("search rows: tag lines that are an object", "the TypeScript fails on it with yargs's help", search("#todo"), { db: "odd-tag-lines-object" });
back("search rows: tag lines that are an object, unused", "Rust takes tag lines only as the list a sync stores", search("oddity"), { db: "odd-tag-lines-object" });
back("search rows: tag lines that are not JSON", "Rust takes tag lines only as the list a sync stores", search("#todo"), { db: "odd-tag-lines-broken" });
back("search rows: a tag line without text", "Rust takes tag lines only as the list a sync stores", search("#todo"), { db: "odd-tag-lines-no-text" });
add("search rows: tag lines that are null", search("#todo"), { db: "odd-tag-lines-null" });
back("search rows: a date in words", "Date reads many notations; Rust only the one Microsoft Graph writes", search("oddity"), { db: "odd-date-in-words" });
back("search rows: a date in words, compact", "Date reads many notations; Rust only the one Microsoft Graph writes", search("oddity", "-c"), { db: "odd-date-in-words" });
add("search rows: a date in words is not looked at when grouped", search("oddity", "-g"), { db: "odd-date-in-words" });
back("search rows: a date that is none", "Date reads many notations; Rust only the one Microsoft Graph writes", search("oddity"), { db: "odd-date-unreadable" });
add("search rows: an empty account beside none", search("oddity"), { db: "odd-empty-account" });
add("search rows: an empty account beside none, grouped", search("oddity", "-g"), { db: "odd-empty-account" });
add("search rows: an empty account beside none, compact", search("oddity", "-c"), { db: "odd-empty-account" });
add("search rows: bars in the names", search("oddity"), { db: "odd-bars-in-names" });
add("search rows: bars in the names, grouped", search("oddity", "-g"), { db: "odd-bars-in-names" });

// ── search: command lines that are yargs's to read ───────────────────────────

const YARGS = "a command line Rust does not read; yargs does";
for (const [name, args] of [
  ["--online", ["plan", "--online"]],
  ["-o", ["plan", "-o"]],
  ["--help", ["--help"]],
  ["-h after the query", ["plan", "-h"]],
  ["no query", []],
  ["no query but an option", ["-g"]],
  ["two words", ["plan", "retro"]],
  ["an unknown option", ["plan", "--fuzzy"]],
  ["an unknown short option", ["plan", "-x"]],
  ["short options in one", ["plan", "-gc"]],
  ["a short option with its value attached", ["plan", "-l5"]],
  ["a short option with =", ["plan", "-l=5"]],
  ["an option twice", ["plan", "-l", "2", "--limit", "3"]],
  ["a flag twice", ["plan", "-g", "--group"]],
  ["an option without its value", ["plan", "--limit"]],
  ["a limit below zero", ["plan", "--limit", "-1"]],
  ["a limit with a fraction", ["plan", "--limit", "1.5"]],
  ["a limit with a leading zero", ["plan", "--limit", "05"]],
  ["a limit in e notation", ["plan", "--limit", "1e1"]],
  ["a limit in words", ["plan", "--limit", "ten"]],
  ["a limit of ten digits", ["plan", "--limit", "1000000000"]],
  ["an empty limit", ["plan", "--limit="]],
  ["an offset in words", ["plan", "--offset=two"]],
  ["max-lines in words", ["#todo", "--max-lines", "all"]],
  ["an empty notebook", ["plan", "--notebook", ""]],
  ["a notebook starting with a dash", ["plan", "--notebook", "-x"]],
  ["--no-group", ["plan", "--no-group"]],
  ["--group=true", ["plan", "--group=true"]],
  ["--group true", ["plan", "--group", "true"]],
  ["-g false and no query", ["-g", "false"]],
  ["--maxLines", ["#todo", "--maxLines", "1"]],
  ["--notebook-id", ["plan", "--notebook-id", "Work"]],
  ["-- before the query", ["--", "plan"]],
  ["-- after the query", ["plan", "--"]],
  ["a query starting with a dash", ["-plan"]],
  ["a lone dash", ["-"]],
  ["an empty query", [""]],
  ["--version", ["plan", "--version"]],
]) {
  back(`search yargs: ${name}`, YARGS, search(...args), NOTES);
}
back("search yargs: an option before the command", YARGS, ["-g", "search", "plan"], NOTES);

// ── everything else is the JavaScript's ──────────────────────────────────────

const NOT_PORTED = "not ported";
back("cli: no command", "yargs's usage error", []);
back("cli: --help", "yargs's", ["--help"]);
back("cli: -h", "yargs's", ["-h"]);
back("cli: --version", "yargs's; it reads the version from package.json", ["--version"]);
back("cli: an unknown command", "yargs's usage error", ["frobnicate"]);
back("cli: ls", NOT_PORTED, ["ls"]);
back("cli: list with a path", NOT_PORTED, ["list", "Work/Meetings"]);
back("cli: notebooks list", NOT_PORTED, ["notebooks", "list", "--json"]);
back("cli: notebooks without a subcommand", "yargs's usage error", ["notebooks"]);
back("cli: sections list", NOT_PORTED, ["sections", "ls", "Work"]);
back("cli: section-groups list", NOT_PORTED, ["section-groups", "list"]);
back("cli: pages list", NOT_PORTED, ["pages", "list", "Work/Meetings", "--limit", "3"]);
back("cli: pages update with a bad action", "yargs's usage error", ["pages", "update", "x", "-t", "body", "-a", "erase", "-c", "y"]);
back("cli: read", NOT_PORTED, ["read", "Work/Meetings/Retro"]);
back("cli: rename with an ID", NOT_PORTED, ["rename", "0-1234567890abcdef", "New name"]);
back("cli: sync", NOT_PORTED, ["sync"]);
// reindex writes a new index, so its modification time is the time of the run.
const AFTER_REINDEX = [AFTER[0], AFTER[1].slice(0, AFTER[1].indexOf("; stat -c %Y"))];
back("cli: reindex with nothing cached", NOT_PORTED, ["reindex"], { after: AFTER_REINDEX });
back("cli: reindex of section files", NOT_PORTED, ["reindex"], {
  env: { ONENOTE_CACHE_DIR: "<CWD>/cache" },
  files: { "cache/Work/Meetings.json": SECTION_FILE },
  after: AFTER_REINDEX,
});
// Where a sync puts them today: <account>/<notebook>/<section>.json.
back("cli: reindex of section files under an account", NOT_PORTED, ["reindex"], {
  db: "notes",
  env: { ONENOTE_CACHE_DIR: "<CWD>/cache" },
  files: { "cache/alice_example.com/Work/Meetings.json": SECTION_FILE },
  after: AFTER_REINDEX,
});
back("search state: section files under an account but no index", "without an index the TypeScript reads the section files", search("plan"), {
  env: { ONENOTE_CACHE_DIR: "<CWD>/cache" },
  files: { "cache/alice_example.com/Work/Meetings.json": SECTION_FILE },
});
add("auth setup: the guide", ["auth", "setup"]);
add("auth setup: with an index, a cache directory and a .env around", ["auth", "setup"], {
  ...NOTES,
  tz: null,
  env: { ONENOTE_CACHE_DIR: "<CWD>/cache", ONENOTE_CLIENT_ID: "00000000-0000-0000-0000-000000000000" },
  files: { ".env": "FOO=bar\n" },
});
back("auth setup: --help", YARGS, ["auth", "setup", "--help"]);
back("auth setup: one word more", YARGS, ["auth", "setup", "now"]);
back("auth setup: an option", YARGS, ["auth", "setup", "--json"]);
back("auth setup: ONENOTE_CLI_NO_NATIVE=1", "the user's way out", ["auth", "setup"], { env: { ONENOTE_CLI_NO_NATIVE: "1" } });
back("auth setup: a bunfig.toml here", "bun reads it; it can preload scripts", ["auth", "setup"], { files: { "bunfig.toml": "# nothing\n" } });
back("cli: auth list", NOT_PORTED, ["auth", "list"]);
back("cli: auth whoami", NOT_PORTED, ["auth", "whoami"]);
back("cli: auth without a subcommand", "yargs's usage error", ["auth"]);
back("cli: SEARCH in capitals is no command", "yargs's usage error", ["SEARCH", "plan"], NOTES);

// ── output ───────────────────────────────────────────────────────────────────

// Every case runs in a package directory of its own, inside its HOME (which
// is where a global install puts a package, too).
const spec = { env: { ONENOTE_CLI_PACKAGE_ROOT: "<HOME>/pkg" }, cases };
const files = new Map([[path.join(HERE, "parity.json"), `${JSON.stringify(spec, null, 2)}\n`]]);
for (const [name, rows] of Object.entries(DBS)) files.set(path.join(DB_DIR, `${name}.json`), `${JSON.stringify(rows, null, 1)}\n`);

const mode = process.argv[2];
if (mode === "--list") {
  for (const c of cases) console.log(`${c.handback ? "back" : "rust"}  ${c.name}`);
} else if (mode === "--check") {
  const stale = [...files].filter(([file, content]) => !fs.existsSync(file) || fs.readFileSync(file, "utf8") !== content).map(([file]) => path.relative(process.cwd(), file));
  const extra = fs.existsSync(DB_DIR) ? fs.readdirSync(DB_DIR).filter((name) => !files.has(path.join(DB_DIR, name))) : [];
  if (stale.length || extra.length) {
    console.error(`not what rs/parity.gen.mjs writes: ${[...stale, ...extra.map((name) => `rs/parity/db/${name}`)].join(", ")}`);
    process.exit(1);
  }
} else {
  fs.rmSync(DB_DIR, { recursive: true, force: true });
  fs.mkdirSync(DB_DIR, { recursive: true });
  for (const [file, content] of files) fs.writeFileSync(file, content);
}
const native = cases.filter((c) => !c.handback).length;
console.error(`${cases.length} cases: ${native} for Rust, ${cases.length - native} handed back`);
