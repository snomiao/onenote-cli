#!/usr/bin/env bun
// What the JavaScript runtime's toLowerCase makes of every code point, for
// the test in rs/src/js.rs that holds the binary's lower case against it
// (`cargo test` in rs/; lower_case_is_the_runtimes).
//
//   bun rs/parity/runtime-lower.ts <file>
//
// writes one line per code point, from U+0000 to U+10FFFF in order (a
// surrogate as a lone code unit), each with five strings as JSON.stringify
// writes them, separated by tabs: the toLowerCase of
//
//   c    c + "Σ"    "A" + c + "Σ"    "AΣ" + c    "AΣ" + c + "A"
//
// The last four show what c is to a capital sigma before and after it,
// whose lower case is σ or ς by its neighbours. Written to a file, not to
// stdout: bun can drop what a pipe does not take at once.
import fs from "node:fs";

const file = process.argv[2];
if (!file) throw new Error("usage: bun rs/parity/runtime-lower.ts <file>");
const SIGMA = "Σ";
const lines: string[] = [];
for (let code = 0; code <= 0x10ffff; code++) {
  const c = String.fromCodePoint(code);
  const asked = [c, c + SIGMA, "A" + c + SIGMA, "A" + SIGMA + c, "A" + SIGMA + c + "A"];
  lines.push(asked.map((text) => JSON.stringify(text.toLowerCase())).join("\t"));
}
fs.writeFileSync(file, `${lines.join("\n")}\n`);
