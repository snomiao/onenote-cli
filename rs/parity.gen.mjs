#!/usr/bin/env node
// Writes rs/parity.json, the parity cases for onenote (Rust against
// TypeScript), and the index fixtures in rs/parity/db/.
//
//   node rs/parity.gen.mjs           # rewrite both
//   node rs/parity.gen.mjs --check   # fail if they are not what this file writes
//   node rs/parity.gen.mjs --list    # print the case names, with who answers each where
//   node rs/parity.gen.mjs --groups  # print the groups of cases, for --filter
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
// index built into it (rs/parity/make-db.ts, from a file in rs/parity/db/,
// which this file writes, or from one of the two in rs/parity/rows/, which
// are kept as the independent verification wrote them).
// The generated file has no path of this machine in it and is committed.
//
// Nothing here touches Microsoft. The cases run with an empty HOME, so there
// is no client id and no token: every command that needs Graph ends at the
// "Please configure your Azure AD app credentials" notice before it could
// open a connection. `auth login`, `auth logout` and `open` (which starts a
// browser) are not run at all.
//
// A case says who is expected to answer it:
//   add(...)   the Rust binary, alone; a search only on Linux and macOS,
//              and on Windows the JavaScript (SEARCH_ELSEWHERE).
//   back(...)  the JavaScript, everywhere: the case carries "handback" with
//              the reason.
// The runner's --handback check fails unless exactly the marked ones get
// there, on the system it runs on. A case that cannot mean the same thing on
// a system carries "skip" for it, and is not run there.

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

/**
 * Rows this file does not write: rs/parity/rows/<name>.json, from the
 * independent verification. `odd-rows` is 37 pages that are unusual but
 * valid (titles with line ends, dates with offsets, tag lines with escapes,
 * letters that change length in small letters); `ties` is 300 pages, in a
 * shuffled order, that are equal in every column the search sorts by.
 */
const KEPT = ["odd-rows", "ties"];

// ── cases ────────────────────────────────────────────────────────────────────

const cases = [];
const MTIME = "2026-03-08T09:05:07Z";
const ROOT = '"$ONENOTE_CLI_PACKAGE_ROOT"';
// What a case leaves behind: the directory it ran in, HOME, the index and
// what SQLite puts beside it (sizes), the index's modification time, and the
// files of ~/.onenote-cli. (rs/parity/files.mjs, because macOS has no
// `stat -c`.)
const FILES = `node "<REPO>/rs/parity/files.mjs"`;
const AFTER = ["ls -A <CWD> <HOME>", `${FILES} --mtime ${ROOT}/.onenote/search.db ${ROOT}/.onenote <HOME>/.onenote-cli`];
/** Sets a file's time to the nanosecond, on every system (`touch -d` is GNU's or BSD's). */
const touch = (when, file) => `sh "<REPO>/rs/parity/touch.sh" ${when} ${file}`;

/**
 * Where the binary hands a search back that it answers elsewhere. The
 * search is ported for Linux and macOS, where its cases ran; on Windows it
 * is not compiled in (rs/src/main.rs), and every search goes to the
 * JavaScript.
 */
const SEARCH_ELSEWHERE = { win32: "the search is not ported for Windows: bun's SQLite, time zone data and console there were never compared" };
/**
 * An index without its -wal file: on macOS the system's SQLite, which
 * bun:sqlite is there, cannot open it read-only, and the TypeScript takes
 * the cache for empty and syncs; the binary leaves that to it
 * (rs/src/search.rs, openable_by_the_runtime).
 */
const CLOSED_ON_MACOS = "macOS's SQLite cannot open an index in WAL mode read-only without its -wal file, and the TypeScript then takes the cache for empty";

/**
 * One case. `db` names the index to start from ("notes", or ["notes",
 * "--wal", "pending"]): by default with an empty -wal file and the -shm file
 * beside it (make-db.ts --wal empty), which is what a finished `onenote sync`
 * leaves on macOS, and on Linux under bun 1.4.4 and older; with "--closed"
 * as one closed cleanly leaves it (the index alone), which is what a sync
 * leaves on Linux under bun 1.4.5 (make-db.ts says why). `closed` says that
 * the index has no -wal file when the command runs (set by "--closed"; for a
 * setup line that removes it). `mtime` is the index's
 * modification time, `root` files to put in the package directory, `tz` the
 * TZ to run under (null: none, the system's zone), `pkg` a package
 * directory other than <HOME>/pkg, `normalize` rules for what differs
 * between two runs of the same program, `skip` the systems on which the
 * case cannot mean the same thing, and why.
 *
 * `handback` is the reason the case goes to the JavaScript on every system,
 * or an object with a reason for each system where it does (`linux`,
 * `darwin`). A search the binary answers is still handed back on the
 * systems that SEARCH_ELSEWHERE names, and on macOS when the index is
 * `closed`.
 */
function make(name, args, { db, mtime = MTIME, root = {}, tz = "UTC", env = {}, files, setup = [], stdin, after = AFTER, handback, normalize, pkg, skip, closed = false } = {}) {
  const lines = ['sh "<REPO>/rs/parity/package.sh"'];
  for (const [file, content] of Object.entries(root)) lines.push(`printf %s ${sq(content)} > ${ROOT}/${file}`);
  if (db) {
    const [which, ...given] = [].concat(db);
    const rows = DBS[which] ? `db/${which}.json` : KEPT.includes(which) ? `rows/${which}.json` : null;
    if (!rows) throw new Error(`${name}: no index called ${which}`);
    closed ||= given.includes("--closed");
    const flags = given.includes("--closed") ? given.filter((flag) => flag !== "--closed") : given.includes("--wal") ? given : [...given, "--wal", "empty"];
    lines.push(`bun "<REPO>/rs/parity/make-db.ts" "<REPO>/rs/parity/${rows}" --mtime ${mtime}${flags.length ? ` ${flags.join(" ")}` : ""}`);
  }
  lines.push(...setup);
  if (cases.some((c) => c.name === name)) throw new Error(`two cases are called ${name}`);
  cases.push({
    name,
    args,
    ...(stdin === undefined ? {} : { stdin }),
    env: { ...(tz === null ? {} : { TZ: tz }), ...(pkg ? { ONENOTE_CLI_PACKAGE_ROOT: pkg } : {}), ...env },
    ...(files ? { files } : {}),
    setup: lines,
    after,
    ...(normalize ? { normalize } : {}),
    ...(typeof handback === "string" ? { handback } : args[0] === "search" ? { handback: { ...(closed ? { darwin: CLOSED_ON_MACOS } : {}), ...handback, ...SEARCH_ELSEWHERE } } : handback ? { handback } : {}),
    ...(skip ? { skip } : {}),
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

// What the bun that runs the cases leaves when a program ends as a sync
// does (make-db.ts --wal sync): on macOS, and under bun 1.4.4 and older, the
// index with an empty -wal file and the -shm file; on Linux under bun 1.4.5
// the index alone. Native everywhere the search is: should a bun on macOS
// ever leave no -wal file, the binary hands this back there and the
// --handback check says so.
add("search state: an index as a sync under this bun leaves it", search("plan"), { db: ["notes", "--wal", "sync"] });
add("search state: an index with an empty -wal file and the -shm file", search("plan"), { db: ["notes", "--wal", "empty"] });
add("search state: an index whose rows are still in the write-ahead log", search("plan"), { db: ["notes", "--wal", "pending"] });
add("search state: tags from the write-ahead log", search("#todo", "-c"), { db: ["notes", "--wal", "pending"] });
// Every other case starts from an index with an empty -wal file and the
// -shm file beside it. One that was closed cleanly has no -wal file, which
// macOS's SQLite needs to open it read-only; it makes a missing -shm file
// itself. A sync on Linux under bun 1.4.5 leaves it so.
add("search state: an index closed cleanly", search("plan"), { db: ["notes", "--closed"] });
add("search state: tags from an index closed cleanly", search("#todo", "-c"), { db: ["notes", "--closed"] });
add("search state: no results from an index closed cleanly", search("nosuchword"), { db: ["notes", "--closed"] });
add("search state: an index closed cleanly, grouped", search("plan", "-g"), { db: ["bulk", "--closed"] });
add("search state: an index with its -wal file and no -shm file", search("plan"), { ...NOTES, setup: [`rm ${ROOT}/.onenote/search.db-shm`] });
add("search state: an index with its -shm file and no -wal file", search("plan"), { ...NOTES, closed: true, setup: [`rm ${ROOT}/.onenote/search.db-wal`] });
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
  setup: [`mkdir -p ${ROOT}/.onenote && printf 'not a database\\n' > ${ROOT}/.onenote/search.db && ${touch(MTIME, `${ROOT}/.onenote/search.db`)}`],
});
back("search state: an index from before the full-text table", "SQLite reports the missing table and the TypeScript reads the section files instead", search("plan"), {
  ...NOTES,
  setup: [`bun -e 'import { Database } from "bun:sqlite"; const db = new Database(process.env.ONENOTE_CLI_PACKAGE_ROOT + "/.onenote/search.db"); db.run("DROP TABLE pages_fts"); db.close(true);' && ${touch(MTIME, `${ROOT}/.onenote/search.db`)} && rm -f ${ROOT}/.onenote/search.db-wal ${ROOT}/.onenote/search.db-shm`],
});
add("search state: tags need no full-text table", search("#todo", "-c"), {
  ...NOTES,
  closed: true,
  setup: [`bun -e 'import { Database } from "bun:sqlite"; const db = new Database(process.env.ONENOTE_CLI_PACKAGE_ROOT + "/.onenote/search.db"); db.run("DROP TABLE pages_fts"); db.close(true);' && ${touch(MTIME, `${ROOT}/.onenote/search.db`)} && rm -f ${ROOT}/.onenote/search.db-wal ${ROOT}/.onenote/search.db-shm`],
});
// The same with an empty -wal file beside it, which macOS's SQLite opens.
add("search state: tags need no full-text table, with an empty -wal file", search("#todo", "-c"), {
  ...NOTES,
  setup: [`bun -e 'import { Database } from "bun:sqlite"; const db = new Database(process.env.ONENOTE_CLI_PACKAGE_ROOT + "/.onenote/search.db"); db.run("DROP TABLE pages_fts"); db.close(true);' && ${touch(MTIME, `${ROOT}/.onenote/search.db`)} && rm -f ${ROOT}/.onenote/search.db-shm && : > ${ROOT}/.onenote/search.db-wal`],
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
const AFTER_REINDEX = [AFTER[0], `${FILES} ${ROOT}/.onenote <HOME>/.onenote-cli`];
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

// ── cases from the independent verification ──────────────────────────────────

// An independent verifier ran 196 cases of its own against both sides
// (queries and expressions, option values, unusual rows, three hundred pages
// that tie on every sort key, zones, variables, states of the machine). They
// are kept as written, but for two that are split in two (see the ties); `x`
// in a name became the group it belongs to. Twelve of them showed where the
// binary answered differently; those now go back.
const FTS_REJECTS = "FTS5 rejects it and the TypeScript reads the section files instead";
const NO_ZONE_NAME = "the runtime reads rules, paths and other spellings its own way; Rust takes zone names only";
const BEFORE_1970 = "a span that reaches back before 1970, where the two time zone databases were not compared";
const FOR_THE_RUNTIME = "a variable named for the runtime (BUN_*, NODE_*, MIMALLOC_*, ICU_*): what it does with it is not imitated";
const OTHER_ZONE_DATA = "the runtime's time zone database and the one compiled in differ for this zone at this time";
/** Where only bun's own data on Linux differs (rs/src/time.rs has a list for each system). */
const ON_LINUX = (why) => ({ linux: why });
const LAST_MILLISECOND = "the last millisecond of a second is left to the runtime, which rounds the file's time its own way";
// On macOS bun reads the system's ICU, which does not know these four names
// under TZ and shows UTC (rs/src/time.rs).
const UNKNOWN_ON_MACOS = "macOS's ICU does not know this zone name under TZ and shows UTC";
const GOES_BACK = {
  'search expression: "plan AND NOT project"': FTS_REJECTS,
  'search expression: "NOT plan"': FTS_REJECTS,
  'search expression: "- title : plan OR zzz"': YARGS,
  'search expression: "plan AND AND retro"': FTS_REJECTS,
  'search expression: "and"': FTS_REJECTS,
  'search expression: "plan or"': FTS_REJECTS,
  'search expression: "plan.or.retro"': FTS_REJECTS,
  'search expression: "rowid:1 OR plan"': FTS_REJECTS,
  'search expression: "plan OR alice\'s"': FTS_REJECTS,
  'search expression: "plan AND \\"of record\\""': FTS_REJECTS,
  'search expression: "plan* OR *plan"': FTS_REJECTS,
  'search expression: "tag:star AND retro"': FTS_REJECTS,
  'search expression: "#todo OR #done"': FTS_REJECTS,
  "search options: since the year 1": "Rust reads a day only in the years 1000 to 9999",
  "search options: since the year 999": "Rust reads a day only in the years 1000 to 9999",
  "search options: since 00030d": "Rust counts back at most 9999 units",
  "search options: since 9999y": BEFORE_1970,
  "search options: since 9999m": BEFORE_1970,
  "search options: since 1026y": BEFORE_1970,
  "search options: since empty": YARGS,
  "search options: --Max-Lines": YARGS,
  "search options: --LIMIT": YARGS,
  "search options: -G": YARGS,
  "search options: a limit with a blank after it": YARGS,
  "search options: a limit with a plus": YARGS,
  "search options: a limit in fullwidth digits": YARGS,
  "search options: an offset of 00": YARGS,
  "search options: -v after the query": YARGS,
  "search options: --online=false": YARGS,
  "search options: --no-online": YARGS,
  'search time: TZ="UTC0"': NO_ZONE_NAME,
  'search time: TZ="utc"': NO_ZONE_NAME,
  'search time: TZ="Z"': NO_ZONE_NAME,
  'search time: TZ="localtime"': NO_ZONE_NAME,
  'search time: TZ="posix/Asia/Tokyo"': NO_ZONE_NAME,
  'search time: TZ="Asia/Tokyo "': NO_ZONE_NAME,
  'search time: TZ="Etc/Unknown"': NO_ZONE_NAME,
  'search time: TZ="SystemV/EST5"': NO_ZONE_NAME,
  "search env: NODE_NO_WARNINGS=1": FOR_THE_RUNTIME,
  "search env: ICU_TIMEZONE_FILES_DIR=/nonexistent": FOR_THE_RUNTIME,
  "search state: .env here is a link to a file that sets TZ": "bun loads it, and TZ changes the time in the last line",
  "search state: .env here is a directory": "a .env that is not a file is not looked into",
  "search state: .env.test.local here sets TZ": "every .env file bun might load is looked at, whatever NODE_ENV is",
  "search state: the index's time has a fraction": LAST_MILLISECOND,
  "auth setup: an empty argument after it": YARGS,
  "search time: EST5EDT": { darwin: UNKNOWN_ON_MACOS },
};
const wentBack = new Set();
/** A case of the verification: for Rust, unless `GOES_BACK` has a reason. */
function also(name, args, options) {
  if (GOES_BACK[name]) wentBack.add(name);
  make(name, args, { ...options, handback: GOES_BACK[name] });
}

// queries
for (const [n, q] of [["blanks around a word", " plan "], ["a line end after a word", "plan\n"], ["a tab between words", "plan\tretro"], ["a line end between words", "plan\nretro"], ["null", "null"], ["undefined", "undefined"], ["1.0", "1.0"], [".5", ".5"], ["+5", "+5"], ["0", "0"], ["a number past 2^53", "9007199254740993"], ["0b11", "0b11"], ["__proto__", "__proto__"], ["constructor", "constructor"], ["$0", "$0"], ["a backslash", "\\"], ["%s", "%s"], ["an asterisk", "*"], ["eight words", "a b c d e f g h"], ["a hash alone", "#"], ["tag: alone", "tag:"], ["a=b at the front", "=plan"]]) also(`search query: ${n}`, search(q), NOTES);
// FTS5 expressions, two SQLite versions
for (const q of ["plan AND NOT project", "NOT plan", "plan NOT (project OR batter)", "^plan OR zzz", "title : plan OR zzz", "{title body} : plan OR zzz", "{title}: plan NOT retro", "- title : plan OR zzz", "plan + of + record OR zzz", "NEAR(plan project) OR zzz", 'NEAR("project plan" timeline, 10) OR zzz', "pl* AND ret*", "plan AND AND retro", "and", "plan or", "plan.or.retro", "body:plan AND title:retro", "rowid:1 OR plan", "plan OR alice's", "plan OR é", "((plan OR retro))", '"plan" OR "retro"', 'plan AND "of record"', "plan* OR *plan", "plan\nOR\tretro", "#todo plan OR milk", "tag:star AND retro", "#todo OR #done", 'title:"weekly sync" OR zzz', "plan NOT plan"]) also(`search expression: ${JSON.stringify(q)}`, search(q), NOTES);
// option values
for (const [n, a] of [["a notebook that is one blank", ["plan", "-n", " "]], ["a notebook that is one quote", ["plan", "-n", "'"]], ["a quote at the front of a notebook", ["plan", "-n", "'Work"]], ["a quote inside a notebook", ["plan", "-n", 'Wo"rk']], ["a quote at the end of a section", ["plan", "-s", "Meet'"]], ["a quote at the front of an account", ["plan", "-a", '"alice@example.com']], ["since in quotes is text", ["plan", "--since", '"2026-01-01"']], ["since with milliseconds is text", ["plan", "--since", "2026-01-01T00:00:00.000Z"]], ["since a leap day", ["plan", "--since", "2024-02-29"]], ["since the year 1", ["plan", "--since", "0001-01-01"]], ["since the year 999", ["plan", "--since", "0999-12-31"]], ["since the last day of 9999", ["plan", "--since", "9999-12-31"]], ["since 0m", ["plan", "--since", "0m"]], ["since 00030d", ["plan", "--since", "00030d"]], ["since 9999m", ["plan", "--since", "9999m"]], ["since 9999y", ["plan", "--since", "9999y"]], ["since 1026y", ["plan", "--since", "1026y"]], ["since 30d and a line end", ["plan", "--since", "30d\n"]], ["since in fullwidth digits", ["plan", "--since", "３０d"]], ["since with a blank in front", ["plan", "--since", " 30d"]], ["since 2026-1-1", ["plan", "--since", "2026-1-1"]], ["since true", ["plan", "--since", "true"]], ["since 0", ["plan", "--since", "0"]], ["since 1e3", ["plan", "--since", "1e3"]], ["since empty", ["plan", "--since", ""]], ["the limit before the query", ["-l", "3", "plan"]], ["limit, offset and max-lines with =", ["--limit=3", "--offset=1", "plan", "-m", "2"]], ["the largest offset", ["plan", "-p", "999999999"]], ["the largest max-lines", ["#todo", "-m", "999999999"]], ["-m 2 grouped", ["#todo", "-m", "2", "-g"]], ["-g and then 1 as the query", ["-g", "1"]], ["-c and then 0 as the query", ["-c", "0"]], ["--compact --group -l 4", ["plan", "--compact", "--group", "-l", "4"]], ["every filter, grouped", ["plan", "-n", "Work", "-s", "Meetings", "-a", "alice@example.com", "-g"]], ["a notebook with %", ["plan", "--notebook", "W%k"]], ["a section of seven _", ["plan", "--section", "_______"]], ["a notebook with a line end", ["plan", "--notebook", "Work\n"]], ["an account with a blank after it", ["plan", "--account", "alice@example.com "]], ["--Max-Lines", ["plan", "--Max-Lines", "3"]], ["--LIMIT", ["plan", "--LIMIT", "3"]], ["-G", ["plan", "-G"]], ["a limit with a blank after it", ["plan", "--limit", "3 "]], ["a limit with a plus", ["plan", "--limit", "+3"]], ["a limit in fullwidth digits", ["plan", "--limit", "３"]], ["an offset of 00", ["plan", "--offset", "00"]], ["-v after the query", ["plan", "-v"]], ["--online=false", ["plan", "--online=false"]], ["--no-online", ["plan", "--no-online"]]]) also(`search options: ${n}`, search(...a), NOTES);
// rows that are unusual but valid
for (const q of ["zeta", "ZETA", "#todo", "#done", "#checkbox", "#star", "#todo zeta", '""', "kelvin", "İstanbul", "one", "ſtar"]) for (const [l, f] of [["flat", []], ["compact", ["-c"]], ["grouped", ["-g"]]]) also(`search rows: ${JSON.stringify(q)} ${l}`, search(q, ...f), { db: "odd-rows" });
// three hundred pages that tie on every sort key. Two of these asked for all
// three hundred at once (`omega -l 300 -c` and `"" -l 300 -c`), which is 302
// lines: more than the TypeScript gets through a socket every time (see
// `bulk` above; under load it lost lines in 17 of 48 runs against itself).
// Each is two cases here, a half at a time.
for (const [n, a] of [["the first half, compact", ["omega", "-l", "150", "-c"]], ["the second half, compact", ["omega", "-l", "150", "-p", "150", "-c"]], ["a window, compact", ["omega", "-l", "40", "-p", "37", "-c"]], ["tags, all", ["#todo", "-l", "300", "-c"]], ["tags, a window", ["#todo", "-l", "25", "-p", "13"]], ["no query, the first half", ['""', "-l", "150", "-c"]], ["no query, the second half", ['""', "-l", "150", "-p", "150", "-c"]], ["a prefix", ["unique*", "-l", "300", "-c"]], ["grouped", ["omega", "-g", "-l", "120"]]]) also(`search ties: ${n}`, search(...a), { db: "ties" });
// zones: the same on both sides
for (const z of ["Asia/Calcutta", "Europe/Kiev", "Antarctica/Troll", "EST", "EST5EDT", "Factory", "Etc/GMT-14", "GMT0", "America/Ciudad_Juarez", "Pacific/Kanton", "Europe/Kyiv"]) also(`search time: ${z}`, search("retro"), one(z, MTIME));
for (const z of ["UTC0", "utc", "Z", "localtime", "posix/Asia/Tokyo", "Asia/Tokyo ", "Etc/Unknown", "SystemV/EST5"]) also(`search time: TZ=${JSON.stringify(z)}`, search("retro"), one(z, MTIME));
// locale and other variables that change nothing
for (const [k, v] of [["LANG", "de_DE.UTF-8"], ["LC_ALL", "ja_JP.UTF-8"], ["LC_ALL", "C"], ["LC_TIME", "fr_FR.UTF-8"], ["LANGUAGE", "fr"], ["FORCE_COLOR", "1"], ["NO_COLOR", "1"], ["DEBUG", "*"], ["NODE_NO_WARNINGS", "1"], ["TZDIR", "/nonexistent"], ["ICU_TIMEZONE_FILES_DIR", "/nonexistent"], ["COLUMNS", "20"]]) also(`search env: ${k}=${v}`, search("retro"), { ...NOTES, env: { [k]: v } });
also("search state: a package directory with blanks and letters outside ASCII", search("retro"), { ...NOTES, pkg: "<HOME>/my pkg ü 日本" });
also("search state: a package directory with # and %41", search("retro"), { ...NOTES, pkg: "<HOME>/c#%41&x=1" });
also("search state: .env here is a link to a file that sets TZ", search("retro"), { ...NOTES, files: { "real.env": "TZ=Asia/Tokyo\n" }, setup: ["ln -s real.env .env"] });
also("search state: .env here is a directory", search("retro"), { ...NOTES, setup: ["mkdir .env"] });
also("search state: .env.test.local here sets TZ", search("retro"), { ...NOTES, files: { ".env.test.local": "TZ=Asia/Tokyo\n" } });
also("search state: a .env here with CRLF and no letters that matter", search("retro"), { ...NOTES, files: { ".env": "FOO=bar\r\nONENOTE_CLIENT_ID=x\r\n" } });
also("search state: the index's time has a fraction", search("retro"), { ...NOTES, setup: [touch("2026-03-08T09:05:07.9996Z", `${ROOT}/.onenote/search.db`)] });
also("auth setup: LANG=de_DE.UTF-8", ["auth", "setup"], { env: { LANG: "de_DE.UTF-8" } });
also("auth setup: an empty argument after it", ["auth", "setup", ""]);
also("auth setup: a package directory with blanks", ["auth", "setup"], { pkg: "<HOME>/my pkg" });
for (const name of Object.keys(GOES_BACK)) if (!wentBack.has(name)) throw new Error(`GOES_BACK names a case that is not there: ${name}`);

// The twelve on which the two differed, and what was added when they were
// repaired.

// Time zones. The runtime has its own time zone database and the binary has
// one compiled in; rs/src/time.rs lists the zones and stretches of time on
// which the two differ (found by `cargo test`, which compares every zone),
// and the binary leaves those to the runtime, three days on either side
// included.
const WINTER_2026 = "2026-11-27T12:00:00Z";
// On macOS only Cuba's winter differs of these (and the four zones named by
// their rule, below); the rest is answered there.
for (const zone of ["America/Vancouver", "America/Edmonton", "America/Winnipeg", "Canada/Pacific", "America/Havana", "Africa/Casablanca", "America/Inuvik", "America/Yellowknife", "America/Rainy_River", "Canada/Mountain", "Canada/Central", "Cuba", "Africa/El_Aaiun"]) {
  const why = zone === "America/Havana" || zone === "Cuba" ? OTHER_ZONE_DATA : ON_LINUX(OTHER_ZONE_DATA);
  back(`search time: ${zone} on 2026-11-27, where the databases differ`, why, search("retro"), one(zone, WINTER_2026));
}
back("search time: Bogota on 1992-05-02, where the databases differ", ON_LINUX(OTHER_ZONE_DATA), search("retro"), one("America/Bogota", "1992-05-02T12:00:00Z"));
back("search time: Tehran on 1979-05-25, where the databases differ", ON_LINUX(OTHER_ZONE_DATA), search("retro"), one("Asia/Tehran", "1979-05-25T21:00:00Z"));
back("search time: Iran on 1979-05-25, where the databases differ", ON_LINUX(OTHER_ZONE_DATA), search("retro"), one("Iran", "1979-05-25T21:00:00Z"));
for (const zone of ["CST6CDT", "MST7MDT", "PST8PDT"]) back(`search time: ${zone}`, { darwin: UNKNOWN_ON_MACOS }, search("retro"), one(zone, MTIME));
// The same zones where the databases agree.
add("search time: Vancouver before the databases differ", search("retro"), one("America/Vancouver", MTIME));
add("search time: Vancouver, the last second before the margin", search("retro"), one("America/Vancouver", "2026-10-29T08:59:59Z"));
back("search time: Vancouver, the first second of the margin", ON_LINUX(OTHER_ZONE_DATA), search("retro"), one("America/Vancouver", "2026-10-29T09:00:00Z"));
add("search time: Casablanca before the databases differ", search("retro"), one("Africa/Casablanca", MTIME));
add("search time: Havana after the winter the databases differ on", search("retro"), one("America/Havana", "2027-06-01T12:00:00Z"));
add("search time: Bogota today", search("retro"), one("America/Bogota", MTIME));
add("search time: Bogota a week after the day the databases differ on", search("retro"), one("America/Bogota", "1992-05-09T12:00:00Z"));
add("search time: Tehran today", search("retro"), one("Asia/Tehran", MTIME));
// --since counts back on the local calendar, from now.
back("search since: a span in Vancouver, where the databases differ from November 2026 on", ON_LINUX(OTHER_ZONE_DATA), search("boundary", "--since", "100d"), {
  tz: "America/Vancouver",
  // One page, last modified 100 days and 30 minutes ago: inside the span by
  // the runtime's clock, outside it by the compiled-in one.
  setup: [
    `printf '[{"notebook":"N","section":"S","title":"Edge page","body":"boundary marker text","last_modified":"%s"}]' "$(node -e 'const d = new Date(); d.setDate(d.getDate() - 100); d.setMinutes(d.getMinutes() - 30); console.log(d.toISOString().slice(0, 19) + "Z")')" > edge.json && bun "<REPO>/rs/parity/make-db.ts" edge.json --mtime ${MTIME} --wal empty && rm edge.json`,
  ],
});
back("search since: thirty days back in Edmonton", ON_LINUX(OTHER_ZONE_DATA), search("era", "--since", "30d"), { db: "ages", tz: "America/Edmonton" });
add("search since: thirty days back in Bogota", search("era", "--since", "30d"), { db: "ages", tz: "America/Bogota" });
add("search since: a day in Vancouver needs no zone", search("plan", "--since", "2026-02-15"), { ...NOTES, tz: "America/Vancouver", mtime: MTIME });
back("search since: ninety-nine years back", BEFORE_1970, search("era", "--since", "99y"), { db: "ages" });
// The years that were compared: 1970 to 2099.
add("search time: the last second that was compared", search("retro"), one("UTC", "2099-12-31T23:59:59Z"));
back("search time: the year 2100", "the time zone databases were compared for 1970 to 2099", search("retro"), one("UTC", "2100-01-01T00:00:00Z"));
back("search time: an index from before 1970", "the time zone databases were compared for 1970 to 2099", search("retro"), one("UTC", "1969-12-31T23:59:59Z"));

// The file's time. The runtime turns it into milliseconds in floating point,
// and from 122 ns before a full second on (in 2026) shows the next second.
add("search state: the index's time is just short of the last millisecond", search("retro"), { ...NOTES, setup: [touch("2026-03-08T09:05:07.998999999Z", `${ROOT}/.onenote/search.db`)] });
back("search state: the index's time is in the last millisecond", LAST_MILLISECOND, search("retro"), { ...NOTES, setup: [touch("2026-03-08T09:05:07.999Z", `${ROOT}/.onenote/search.db`)] });
back("search state: the index's time is 120 ns short of the next second", LAST_MILLISECOND, search("retro"), {
  ...NOTES,
  setup: [touch("2026-03-08T09:05:07.999999900Z", `${ROOT}/.onenote/search.db`)],
});

// Variables the runtime reads for itself.
const PID = [["\\(node:\\d+\\)", "(node:<PID>)"]];
back("search env: BUN_JSC_dumpOptions=1 makes the engine list its options", FOR_THE_RUNTIME, search("retro"), { ...NOTES, env: { BUN_JSC_dumpOptions: "1" } });
back("auth setup: BUN_JSC_dumpOptions=1 makes the engine list its options", FOR_THE_RUNTIME, ["auth", "setup"], { env: { BUN_JSC_dumpOptions: "1" } });
// The warning carries the process id: the one normalize rule in this file.
back("search env: NODE_DEBUG=http makes the runtime warn", FOR_THE_RUNTIME, search("retro"), { ...NOTES, env: { NODE_DEBUG: "http" }, normalize: PID });
back("auth setup: NODE_DEBUG=http makes the runtime warn", FOR_THE_RUNTIME, ["auth", "setup"], { env: { NODE_DEBUG: "http" }, normalize: PID });
back("search env: a transpiler cache directory, which the runtime fills", FOR_THE_RUNTIME, search("retro"), { ...NOTES, env: { BUN_RUNTIME_TRANSPILER_CACHE_PATH: "<HOME>/transpiled" } });
back("search env: BUN_ENV", FOR_THE_RUNTIME, search("retro"), { ...NOTES, env: { BUN_ENV: "production" } });
back("search env: MIMALLOC_SHOW_ERRORS=1", FOR_THE_RUNTIME, search("retro"), { ...NOTES, env: { MIMALLOC_SHOW_ERRORS: "1" } });
back("search env: NODE_EXTRA_CA_CERTS", FOR_THE_RUNTIME, search("retro"), { ...NOTES, env: { NODE_EXTRA_CA_CERTS: "/nonexistent.pem" } });
back("search env: ICU_DATA", FOR_THE_RUNTIME, search("retro"), {
  ...NOTES,
  env: { ICU_DATA: "/nonexistent" },
  skip: { darwin: "bun 1.4.4 crashes under it on macOS, and its crash report differs from one run to the next" },
});
// On Windows a variable's name is the same in any case, so there this is
// NODE_DEBUG, and the warning with the process id comes too.
back("search env: node_debug in small letters", FOR_THE_RUNTIME, search("retro"), { ...NOTES, env: { node_debug: "http" }, normalize: PID });
// What bun's installer puts into the shell's profile.
add("search env: BUN_INSTALL, as bun's installer sets it", search("retro"), { ...NOTES, env: { BUN_INSTALL: "<HOME>/.bun" } });
add("auth setup: BUN_INSTALL, as bun's installer sets it", ["auth", "setup"], { env: { BUN_INSTALL: "<HOME>/.bun" } });
// Set by `bun run` and `npm run` for what they start.
add("search env: NODE and npm_config_user_agent, as a package script has them", search("retro"), { ...NOTES, env: { NODE: "/usr/bin/node", npm_config_user_agent: "bun/1.4.0 npm/? node/v26.3.0 linux x64" } });
// The allocator has read the environment long before bun reads a .env.
add("search state: a .env here that sets MIMALLOC_VERBOSE comes too late for the allocator", search("retro"), { ...NOTES, files: { ".env": "MIMALLOC_VERBOSE=1\n" } });
back("search state: a .env here that names ICU's data directory", "the letters ICU anywhere in the file are enough to step aside", search("retro"), {
  ...NOTES,
  files: { ".env": "ICU_TIMEZONE_FILES_DIR=/nonexistent\n" },
});

// A package directory bun cannot load a program from: the TypeScript ends
// with `Module not found`, and the binary hands back to get the same. (The
// --handback check gives every case its own stand-in package directory, so
// there these four are answered by the binary and cannot be marked.)
// Windows has neither: `?` cannot be in a file name there, and a backslash
// separates directories.
const NO_QUESTION_MARK = { win32: "a file name cannot hold ? on Windows" };
add("search state: a package directory with ? in its name cannot be loaded", search("retro"), { ...NOTES, pkg: "<HOME>/what?", skip: NO_QUESTION_MARK });
add("auth setup: a package directory with ? in its name cannot be loaded", ["auth", "setup"], { pkg: "<HOME>/what?", skip: NO_QUESTION_MARK });
add("search state: a package directory with a backslash in its name cannot be loaded", search("retro"), {
  ...NOTES,
  pkg: "<HOME>/back\\slash",
  skip: { win32: "a backslash separates directories on Windows" },
});
add("search state: a package directory that is a link to one with ? in its name", search("retro"), {
  ...NOTES,
  pkg: "<HOME>/link",
  setup: ['if [ -d "<HOME>/link" ]; then mv "<HOME>/link" "<HOME>/real?dir" && ln -s "real?dir" "<HOME>/link"; fi'],
  skip: NO_QUESTION_MARK,
});

// ── output ───────────────────────────────────────────────────────────────────

// Every case runs in a package directory of its own, inside its HOME (which
// is where a global install puts a package, too).
const spec = { env: { ONENOTE_CLI_PACKAGE_ROOT: "<HOME>/pkg" }, cases };
const files = new Map([[path.join(HERE, "parity.json"), `${JSON.stringify(spec, null, 2)}\n`]]);
for (const [name, rows] of Object.entries(DBS)) files.set(path.join(DB_DIR, `${name}.json`), `${JSON.stringify(rows, null, 1)}\n`);

/** Who answers on a system: "skip", "back" or "rust". */
const on = (c, system) => (c.skip?.[system] ? "skip" : typeof c.handback === "string" || c.handback?.[system] ? "back" : "rust");
const mode = process.argv[2];
/**
 * The groups the cases fall into, by the start of their names ("search
 * time: "): each name contains its own group's and no other, so that the
 * parity runner's --filter with each of them in turn runs every case once.
 */
function groups() {
  const starts = [...new Set(cases.map((c) => c.name.slice(0, c.name.indexOf(": ") + 2)))];
  for (const c of cases) {
    const holding = starts.filter((start) => c.name.includes(start));
    if (holding.length !== 1) throw new Error(`${c.name}: in ${holding.length} groups`);
  }
  return starts;
}
groups();
if (mode === "--groups") {
  for (const start of groups()) console.log(start);
} else if (mode === "--list") {
  console.log("linux darwin win32");
  for (const c of cases) console.log(`${["linux", "darwin", "win32"].map((system) => on(c, system).padEnd(4)).join(" ")}  ${c.name}`);
} else if (mode === "--check") {
  const stale = [...files].filter(([file, content]) => !fs.existsSync(file) || fs.readFileSync(file, "utf8") !== content).map(([file]) => path.relative(process.cwd(), file));
  const extra = fs.existsSync(DB_DIR) ? fs.readdirSync(DB_DIR).filter((name) => !files.has(path.join(DB_DIR, name))) : [];
  if (stale.length || extra.length) {
    console.error(`not what rs/parity.gen.mjs writes: ${[...stale, ...extra.map((name) => `rs/parity/db/${name}`)].join(", ")}`);
    process.exit(1);
  }
  const missing = KEPT.filter((name) => !fs.existsSync(path.join(HERE, "parity", "rows", `${name}.json`)));
  if (missing.length) {
    console.error(`missing: ${missing.map((name) => `rs/parity/rows/${name}.json`).join(", ")}`);
    process.exit(1);
  }
} else {
  fs.rmSync(DB_DIR, { recursive: true, force: true });
  fs.mkdirSync(DB_DIR, { recursive: true });
  for (const [file, content] of files) fs.writeFileSync(file, content);
}
const count = (system, who) => cases.filter((c) => on(c, system) === who).length;
console.error(
  `${cases.length} cases; ` +
    ["linux", "darwin", "win32"].map((system) => `${system}: ${count(system, "rust")} for Rust, ${count(system, "back")} handed back, ${count(system, "skip")} skipped`).join("; "),
);
