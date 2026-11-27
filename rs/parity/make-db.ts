#!/usr/bin/env bun
// Builds the search index a parity case starts from, the way `onenote sync`
// leaves it: <package root>/.onenote/search.db, an SQLite file with a `pages`
// table and a contentless FTS5 index over it.
//
//   bun rs/parity/make-db.ts <pages.json> [--mtime <ISO time>] [--wal empty|pending]
//
// The package root is $ONENOTE_CLI_PACKAGE_ROOT. The schema is the one in
// src/cache.ts (openSearchDb), and the rows go in the way
// upsertSectionToIndex puts them in. pages.json is an array of rows, each
// with the columns it wants to set; a column left out is NULL. A value can
// be given a storage class other than TEXT as { "integer": 42 } or
// { "blob": "<hex>" }, for rows no sync would write.
//
// --mtime sets the file's modification time, which `onenote search` prints
// as "last synced".
//
// --wal says what is left beside the file:
//   (not given)  nothing: one file with everything in it;
//   empty        search.db-wal (empty) and search.db-shm, which is what a
//                finished `onenote sync` leaves (it closes the database with
//                statements still open);
//   pending      a search.db-wal that still holds the rows, as after a sync
//                that was killed.
//
// The writing happens in a child process, so that this one can set the
// modification time after the last byte has been written.

import { Database } from "bun:sqlite";
import fs from "node:fs";
import path from "node:path";

const args = process.argv.slice(2);
const option = (name: string): string | undefined => {
  const at = args.indexOf(name);
  if (at < 0) return undefined;
  const [, value] = args.splice(at, 2);
  return value;
};
const flag = (name: string): boolean => {
  const at = args.indexOf(name);
  if (at >= 0) args.splice(at, 1);
  return at >= 0;
};
const isChild = flag("--child");
const mtime = option("--mtime");
const wal = option("--wal");
const root = process.env.ONENOTE_CLI_PACKAGE_ROOT;
const [source] = args;
if (!source || !root || (wal !== undefined && wal !== "empty" && wal !== "pending")) {
  console.error("usage: ONENOTE_CLI_PACKAGE_ROOT=<dir> make-db.ts <pages.json> [--mtime <ISO time>] [--wal empty|pending]");
  process.exit(2);
}
const dir = path.join(root, ".onenote");
const file = path.join(dir, "search.db");

type Cell = string | number | null | { integer: number } | { real: number } | { blob: string } | { tag: string; text: string }[];
type Row = Record<string, Cell | undefined>;
const COLUMNS = [
  "section",
  "notebook",
  "title",
  "body",
  "web_url",
  "page_guid",
  "has_todo",
  "has_done",
  "tags",
  "tag_lines",
  "account",
  "last_modified",
];

function cell(value: Cell | undefined): string | number | null | Uint8Array {
  if (value === undefined || value === null) return null;
  if (Array.isArray(value)) return JSON.stringify(value);
  if (typeof value === "object") {
    if ("integer" in value) return value.integer;
    if ("real" in value) return value.real;
    return Uint8Array.from(Buffer.from(value.blob, "hex"));
  }
  return value;
}

function write() {
  const rows = JSON.parse(fs.readFileSync(source, "utf8")) as Row[];
  const db = new Database(file, { create: true });
  db.run("PRAGMA journal_mode=WAL");
  db.run(`
    CREATE TABLE IF NOT EXISTS pages (
      id INTEGER PRIMARY KEY,
      section TEXT NOT NULL,
      notebook TEXT NOT NULL,
      title TEXT,
      body TEXT,
      web_url TEXT,
      page_guid TEXT,
      has_todo INTEGER DEFAULT NULL,
      has_done INTEGER DEFAULT NULL,
      tags TEXT DEFAULT NULL,
      tag_lines TEXT DEFAULT NULL,
      account TEXT DEFAULT NULL,
      last_modified TEXT DEFAULT NULL
    )
  `);
  db.run("CREATE INDEX IF NOT EXISTS pages_section ON pages(section, notebook)");
  db.run(`
    CREATE VIRTUAL TABLE IF NOT EXISTS pages_fts USING fts5(
      title, body,
      content="",
      tokenize="unicode61"
    )
  `);
  const insert = db.prepare(`INSERT INTO pages(${COLUMNS.join(", ")}) VALUES (${COLUMNS.map(() => "?").join(", ")})`);
  db.transaction(() => {
    for (const row of rows) insert.run(...COLUMNS.map((column) => cell(row[column])));
    db.run("INSERT INTO pages_fts(rowid, title, body) SELECT id, title, body FROM pages");
  })();
  if (wal === "pending") process.kill(process.pid, "SIGKILL");
  if (wal === "empty") {
    // As src/cache.ts does at the end of a sync: `insert` is still open, so
    // the database is only really closed when the process ends.
    db.close();
    return;
  }
  insert.finalize();
  db.close(true);
}

if (isChild) {
  write();
} else {
  fs.mkdirSync(dir, { recursive: true });
  for (const ext of ["", "-shm", "-wal"]) fs.rmSync(file + ext, { force: true });
  const child = Bun.spawnSync([process.execPath, import.meta.path, source, "--child", ...(wal ? ["--wal", wal] : [])], {
    stdio: ["ignore", "inherit", "inherit"],
    env: process.env,
  });
  const killed = wal === "pending" && child.signalCode === "SIGKILL";
  if (!killed && child.exitCode !== 0) process.exit(child.exitCode ?? 1);
  if (!wal) for (const ext of ["-shm", "-wal"]) fs.rmSync(file + ext, { force: true });
  if (mtime) {
    const when = new Date(mtime);
    fs.utimesSync(file, when, when);
  }
}
