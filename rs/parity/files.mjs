#!/usr/bin/env node
// What a parity case leaves behind, for an `after` line (rs/parity.gen.mjs):
// every regular file under the directories named, with its size in bytes,
// and with --mtime the modification time of one file in whole seconds.
//
//   node rs/parity/files.mjs [--mtime <file>] <directory>...
//
// It does what `find <directory>... -type f -exec stat -c '%s %n' {} + |
// sort; stat -c %Y <file>` did on Linux, in one way on Linux, macOS and
// Windows: macOS has no `stat -c`, and Git Bash's tools on Windows spell
// some things their own way. Paths are printed as given, with `/` before
// what lies below, so that the parity runner can put <HOME> and <ROOT> back
// into them; files come in the order of their paths, by code unit.

import fs from "node:fs";

const args = process.argv.slice(2);
let mtime = null;
if (args[0] === "--mtime") {
  mtime = args[1];
  args.splice(0, 2);
}

const found = [];
const missing = [];
function walk(dir, shown) {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const path = `${dir}/${entry.name}`;
    const name = `${shown}/${entry.name}`;
    // As find -type f: links are not followed, and only regular files count.
    if (entry.isDirectory()) walk(path, name);
    else if (entry.isFile()) found.push(`${fs.statSync(path).size} ${name}`);
  }
}
for (const dir of args) {
  if (fs.existsSync(dir)) walk(dir, dir);
  else missing.push(`${dir}: no such directory`);
}
const lines = [...missing, ...found.sort((a, b) => (a.slice(a.indexOf(" ")) < b.slice(b.indexOf(" ")) ? -1 : 1))];
if (mtime !== null) {
  try {
    lines.push(`modified: ${fs.statSync(mtime, { bigint: true }).mtimeNs / 1_000_000_000n}`);
  } catch {
    lines.push(`modified: ${mtime} is not there`);
  }
}
process.stdout.write(lines.map((line) => `${line}\n`).join(""));
