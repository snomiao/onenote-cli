#!/usr/bin/env bun
// Writes rs/src/unicode.rs: what bun's `toLowerCase` does with every code
// point, recorded once, so that the binary does not take it from Rust's own
// tables. Those follow the toolchain that builds the binary: a Rust with a
// newer Unicode lower-cases characters that bun keeps (Rust 1.100, Unicode
// 18: U+AB6C, U+AB6D and 16 capitals from U+1DF40 to U+1DF7E, which bun 1.4
// leaves as they are), so the binary would answer according to whichever
// toolchain built it. rustdesk.skill's rs/unicode.gen.mjs is the model.
//
//   bun rs/unicode.gen.mjs            rewrite rs/src/unicode.rs
//   bun rs/unicode.gen.mjs --check    fail if this bun says otherwise
//
// For every code point it records:
// - String.fromCodePoint(c).toLowerCase() where that is not c itself;
// - whether bun's RegExp counts c as unassigned (\p{Cn});
// - for the code points of SIGMA_BLOCKS, what c is to a capital sigma
//   beside it, the one lower case that depends on its neighbours: Σ becomes
//   ς at the end of a word, where a cased letter comes before it and none
//   after it, the characters in between that are case-ignorable passed
//   over. Taken from toLowerCase itself: c ends the search as a cased letter
//   when (c + "Σ") ends in ς, and is passed over when ("A" + c + "Σ") does
//   and (c + "Σ") does not. That the same holds after the sigma is checked
//   here for each of them.
// Only bun gives these: node has another ICU. rs/src/js.rs answers from
// this table and hands back what it is not certain of (see lowercase_unsure
// there); rs/parity/proof.sh runs --check, and a test in rs/src/js.rs holds
// the binary's answers against the bun at hand (rs/parity/runtime-lower.ts),
// so a bun that lower-cases another way, or knows another Unicode, stops the
// proof.
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const TARGET = path.join(HERE, "src", "unicode.rs");
const check = process.argv.includes("--check");

if (!process.versions.bun) {
  console.error("rs/unicode.gen.mjs records bun's toLowerCase: run it with bun, not node");
  process.exit(2);
}
const unicode = process.versions.unicode;
if (!unicode) {
  console.error("this bun does not say which Unicode it has (process.versions.unicode)");
  process.exit(2);
}

// A Rust string literal that is ASCII: anything else as \u{…}.
const literal = (text) =>
  `"${[...text].map((ch) => {
    const code = ch.codePointAt(0);
    return code >= 0x20 && code < 0x7f && ch !== '"' && ch !== "\\" ? ch : `\\u{${code.toString(16)}}`;
  }).join("")}"`;
const hex = (code) => `0x${code.toString(16)}`;

/** Ranges [from, to] of the code points for which `test` holds, in order. */
const ranges = (test) => {
  const found = [];
  let from = null;
  for (let code = 0; code <= 0x10ffff; code++) {
    const yes = test(code);
    if (yes && from === null) from = code;
    if (!yes && from !== null) {
      found.push([from, code - 1]);
      from = null;
    }
  }
  if (from !== null) found.push([from, 0x10ffff]);
  return found;
};
const surrogate = (code) => code >= 0xd800 && code <= 0xdfff;

const lower = [];
for (let code = 0; code <= 0x10ffff; code++) {
  const ch = String.fromCodePoint(code);
  if (!surrogate(code) && ch.toLowerCase() !== ch) lower.push([code, ch.toLowerCase()]);
}
const unassigned = /^\p{Cn}$/u;
// A surrogate is assigned (Cs), and no Rust char.
const assigned = ranges((code) => surrogate(code) || !unassigned.test(String.fromCodePoint(code)));

// The capital sigma and its neighbours. A lone surrogate stops the search
// as a character that is not cased (the binary takes it so).
const SIGMA = "\u03a3";
const FINAL = "\u03c2";
const MEDIAL = "\u03c3";
// Where the binary answers for the neighbours of a capital sigma: Latin up
// to Latin Extended-B, the combining marks and Greek, Greek Extended, and
// General Punctuation. What a character there is to the sigma is the same
// in Unicode 13, 15.1, 16 and 17 (Python 3.9's and 3.13's tables, node 20's
// ICU 76.1, bun 1.4's ICU 78.3, and macOS's ICU 78.1, which bun takes
// there). Elsewhere it is not: U+0295 changed in 17, and macOS's ICU passes
// over private-use characters (U+F870 to U+F87F, among others) that bun's
// own ICU does not. Beside any other character the binary hands back.
// (Greek Extended and General Punctuation are one range: they adjoin.)
const SIGMA_BLOCKS = [[0x0, 0x24f], [0x300, 0x3ff], [0x1f00, 0x206f]];
const inBlocks = (code) => SIGMA_BLOCKS.some(([a, b]) => code >= a && code <= b);
const sigmaCased = new Uint8Array(0x110000);
const sigmaPassed = new Uint8Array(0x110000);
for (let code = 0; code <= 0x10ffff; code++) {
  if (!inBlocks(code) && !surrogate(code)) continue;
  const ch = String.fromCodePoint(code);
  const cased = (ch + SIGMA).toLowerCase().endsWith(FINAL);
  const passed = !cased && ("A" + ch + SIGMA).toLowerCase().endsWith(FINAL);
  // After the sigma: a cased letter makes it medial; one passed over leaves
  // it final at the end, and medial before a cased letter.
  const after = ("A" + SIGMA + ch).toLowerCase()[1];
  const between = ("A" + SIGMA + ch + "A").toLowerCase()[1];
  if (after !== (cased ? MEDIAL : FINAL) || between !== (cased || passed ? MEDIAL : FINAL)) {
    console.error(`U+${code.toString(16).toUpperCase()} is not the same to a capital sigma after it as before it in this bun: rs/src/js.rs cannot answer for its final sigma`);
    process.exit(1);
  }
  if (surrogate(code) && (cased || passed)) {
    console.error(`the lone surrogate U+${code.toString(16).toUpperCase()} does not stop the final-sigma search in this bun`);
    process.exit(1);
  }
  if (surrogate(code)) continue;
  sigmaCased[code] = cased ? 1 : 0;
  sigmaPassed[code] = passed ? 1 : 0;
}
const cased = ranges((code) => sigmaCased[code] === 1);
const passed = ranges((code) => sigmaPassed[code] === 1);

// What the table claims has to hold together: every character that has a
// lower case, every character in one, and every character the sigma looks
// at, is assigned.
const isAssigned = (code) => assigned.some(([a, b]) => code >= a && code <= b);
for (const [code, to] of lower) {
  for (const c of [code, ...[...to].map((ch) => ch.codePointAt(0))]) {
    if (!isAssigned(c)) {
      console.error(`U+${c.toString(16).toUpperCase()} is in bun's lower case of U+${code.toString(16).toUpperCase()}, and its RegExp says it is unassigned: this bun's ICU and its RegExp have different Unicodes`);
      process.exit(1);
    }
  }
}
for (const [a, b] of [...cased, ...passed]) {
  if (!isAssigned(a) || !isAssigned(b) || !assigned.some(([x, y]) => a >= x && b <= y)) {
    console.error(`U+${a.toString(16).toUpperCase()}..U+${b.toString(16).toUpperCase()} matter to a capital sigma, and bun's RegExp says some of them are unassigned: this bun's ICU and its RegExp have different Unicodes`);
    process.exit(1);
  }
}

const rows = (items, perLine) => {
  const lines = [];
  for (let i = 0; i < items.length; i += perLine) lines.push(`    ${items.slice(i, i + perLine).join(" ")}`);
  return lines.join("\n");
};
const table = (list) => rows(list.map(([a, b]) => `(${hex(a)}, ${hex(b)}),`), 6);
const text = `//! Written by rs/unicode.gen.mjs from bun (Unicode ${unicode}); do not edit.
//! \`bun rs/unicode.gen.mjs\` writes it again, \`--check\` fails when the bun at
//! hand says otherwise. rs/src/js.rs answers \`toLowerCase\` from it.

/// The Unicode of bun's ICU (\`process.versions.unicode\`), for the tests.
#[cfg(test)]
pub const VERSION: &str = ${literal(unicode)};

/// \`String.fromCodePoint(c).toLowerCase()\` for every code point whose lower
/// case is not itself, in order of code point.
#[rustfmt::skip]
pub static LOWER: [(u32, &str); ${lower.length}] = [
${rows(lower.map(([code, to]) => `(${hex(code)}, ${literal(to)}),`), 6)}
];

/// The code points bun's RegExp does not count as unassigned (\`\\P{Cn}\`), as
/// ranges in order: every character of this Unicode, private use, the
/// surrogates and nothing else.
#[rustfmt::skip]
pub static ASSIGNED: [(u32, u32); ${assigned.length}] = [
${table(assigned)}
];

/// The blocks whose characters the binary answers for beside a capital
/// sigma (see rs/unicode.gen.mjs); SIGMA_CASED and SIGMA_PASSED cover only
/// these.
#[rustfmt::skip]
pub static SIGMA_BLOCKS: [(u32, u32); ${SIGMA_BLOCKS.length}] = [
${table(SIGMA_BLOCKS)}
];

/// The characters that end the search for a cased letter beside a capital
/// sigma as one, as ranges in order: \`(c + "\\u{3a3}").toLowerCase()\` ends
/// in a final sigma.
#[rustfmt::skip]
pub static SIGMA_CASED: [(u32, u32); ${cased.length}] = [
${table(cased)}
];

/// The characters that search passes over (case-ignorable), as ranges in
/// order: \`("A" + c + "\\u{3a3}").toLowerCase()\` ends in a final sigma, and
/// c is not in SIGMA_CASED.
#[rustfmt::skip]
pub static SIGMA_PASSED: [(u32, u32); ${passed.length}] = [
${table(passed)}
];
`;

const counts = `${lower.length} lower cases, ${assigned.length} ranges assigned, ${cased.length} cased and ${passed.length} passed over beside a sigma`;
if (check) {
  const current = fs.existsSync(TARGET) ? fs.readFileSync(TARGET, "utf8") : "";
  if (current !== text) {
    const crlf = current.replace(/\r\n/g, "\n") === text;
    console.error(
      crlf
        ? "rs/src/unicode.rs differs from what rs/unicode.gen.mjs writes only in CR LF line ends: check out rs/ with LF"
        : `rs/src/unicode.rs is not what bun ${process.versions.bun} (Unicode ${unicode}) says: its toLowerCase or its Unicode has moved. ` +
            "Write the table again with `bun rs/unicode.gen.mjs`, and add the blocks where it moved to lowercase_unsure in rs/src/js.rs, " +
            "so that the binary does not answer for a bun from before the move either",
    );
    process.exit(1);
  }
  console.error(`rs/src/unicode.rs is what bun ${process.versions.bun} says (Unicode ${unicode}, ${counts})`);
} else {
  fs.writeFileSync(TARGET, text);
  console.error(`wrote rs/src/unicode.rs from bun ${process.versions.bun} (Unicode ${unicode}, ${counts})`);
}
