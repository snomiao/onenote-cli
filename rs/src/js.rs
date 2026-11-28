//! JavaScript's strings, as far as the ported commands use them.
//!
//! The TypeScript counts, cuts and searches text in UTF-16 code units
//! (`length`, `slice`, `indexOf`, regular expressions without the `u` flag),
//! so an emoji is two units and a cut can go through the middle of one. To
//! print the same bytes, text is kept here as code units too and turned into
//! UTF-8 only when a whole line is written, the way `console.log` does it:
//! half an emoji becomes U+FFFD.
//!
//! No answer depends on the Unicode tables of the Rust that built the binary:
//! white space is listed here, lower case comes from rs/src/unicode.rs,
//! recorded from bun (see `lowercase_unsure`).

use crate::unicode;

/// `text` as JavaScript holds it.
pub fn units(text: &str) -> Vec<u16> {
    text.encode_utf16().collect()
}

/// What `console.log` writes for a string.
pub fn text(units: &[u16]) -> String {
    String::from_utf16_lossy(units)
}

/// `\s` in a regular expression, and what `trim()` removes.
pub fn is_space(unit: u16) -> bool {
    matches!(
        unit,
        0x09..=0x0D | 0x20 | 0xA0 | 0x1680 | 0x2000..=0x200A | 0x2028 | 0x2029 | 0x202F | 0x205F | 0x3000 | 0xFEFF
    )
}

/// `\w` in a regular expression without the `u` flag.
pub fn is_word(unit: u16) -> bool {
    matches!(unit, 0x30..=0x39 | 0x41..=0x5A | 0x61..=0x7A | 0x5F)
}

/// `s.trim()`
pub fn trim(units: &[u16]) -> &[u16] {
    let start = units.iter().position(|&u| !is_space(u)).unwrap_or(units.len());
    let end = units.iter().rposition(|&u| !is_space(u)).map_or(start, |at| at + 1);
    &units[start..end]
}

/// `s.replace(/\s+/g, " ")`
pub fn collapse_spaces(units: &[u16]) -> Vec<u16> {
    let mut out = Vec::with_capacity(units.len());
    let mut in_space = false;
    for &unit in units {
        if is_space(unit) {
            if !in_space {
                out.push(0x20);
            }
            in_space = true;
        } else {
            out.push(unit);
            in_space = false;
        }
    }
    out
}

/// `s.replace(/\s+/g, " ").trim()`
pub fn squeeze(units: &[u16]) -> Vec<u16> {
    trim(&collapse_spaces(units)).to_vec()
}

/// `s.slice(start, end)` for indexes that are not negative.
pub fn slice(units: &[u16], start: usize, end: usize) -> &[u16] {
    let end = end.min(units.len());
    let start = start.min(end);
    &units[start..end]
}

/// `s.indexOf(needle)`
pub fn index_of(units: &[u16], needle: &[u16]) -> Option<usize> {
    if needle.is_empty() {
        return Some(0);
    }
    if needle.len() > units.len() {
        return None;
    }
    (0..=units.len() - needle.len()).find(|&at| units[at..at + needle.len()] == *needle)
}

/// `s.toLowerCase()`, or `None` where the runtime's answer is not certain
/// (see `lowercase_unsure`): the caller hands the invocation back.
///
/// The answer comes from rs/src/unicode.rs, recorded from bun by
/// rs/unicode.gen.mjs, never from Rust's own tables, which follow the
/// toolchain that built the binary (rustify's docs/porting.md, "Unicode
/// follows the toolchain, not bun"). Unicode's default lower-casing: one
/// letter can become two units, and a capital sigma depends on where in the
/// word it stands (see `final_sigma`). Half an emoji is neither a letter nor
/// part of a word and stays as it is.
pub fn to_lower_case(units: &[u16]) -> Option<Vec<u16>> {
    let points: Vec<Result<char, u16>> = char::decode_utf16(units.iter().copied()).map(|decoded| decoded.map_err(|lone| lone.unpaired_surrogate())).collect();
    let mut out = Vec::with_capacity(units.len());
    for (at, point) in points.iter().enumerate() {
        match *point {
            Err(lone) => out.push(lone),
            Ok(c) if c.is_ascii() => out.push(u16::from(c.to_ascii_lowercase() as u8)),
            Ok(SIGMA) => out.push(if final_sigma(&points, at)? { 0x3C2 } else { 0x3C3 }),
            Ok(c) if lowercase_unsure(c) => return None,
            Ok(c) => match table_lower(c) {
                Some(lower) => out.extend(lower.encode_utf16()),
                None => out.extend(c.encode_utf16(&mut [0; 2]).iter()),
            },
        }
    }
    Some(out)
}

/// The capital sigma, whose lower case is the only one that depends on its
/// neighbours.
const SIGMA: char = '\u{3a3}';

/// Characters whose `toLowerCase` the binary does not answer for, because a
/// runtime may give another lower case than the one recorded from bun 1.4 in
/// rs/src/unicode.rs (rustdesk.skill's `lowercase_unsure`, whose reasons
/// these are):
///
/// - the blocks whose case pairs came with Unicode 14 or later, which a bun
///   with an older ICU lower-cases otherwise or not at all (the Cyrillic of
///   U+1C89, Unicode 16, among them). Every code point whose lower case is
///   not the same in Unicode 13 (Python 3.9's tables), 15.1 (Python 3.13's)
///   or 16 (node 20's ICU 76.1) as in bun 1.4's 17 is in one of them;
/// - every code point the table's Unicode leaves unassigned, which a later
///   Unicode may give a lower case (Unicode 18 gave U+AB6C one);
/// - every code point whose lower case in Rust's own tables is not the
///   table's. Those tables follow the toolchain that built the binary; where
///   two Unicodes disagree, runtimes of either may.
///
/// The capital sigma is answered (see `final_sigma`). The answer itself
/// always comes from the table, never from Rust's tables, so whatever
/// toolchain built the binary, it answers the same where it answers at all,
/// and hands back the rest. A test below holds every answer against the bun
/// at hand.
pub fn lowercase_unsure(c: char) -> bool {
    matches!(c as u32,
        0x1c80..=0x1c8f // Cyrillic Extended-C
        | 0x2c00..=0x2c5f // Glagolitic
        | 0xa720..=0xa7ff // Latin Extended-D
        | 0x10570..=0x105bf // Vithkuqi
        | 0x10d40..=0x10d8f // Garay
        | 0x16ea0..=0x16edf // Beria Erfe
        | 0x1e900..=0x1e95f // Adlam
    ) || !assigned(c)
        || !rust_agrees(c)
}

/// The table's lower case of `c`, `None` where that is `c` itself.
fn table_lower(c: char) -> Option<&'static str> {
    let at = unicode::LOWER.binary_search_by_key(&(c as u32), |&(code, _)| code).ok()?;
    Some(unicode::LOWER[at].1)
}

/// Whether `c` is in one of `ranges` (in order, as the table has them).
fn in_ranges(ranges: &[(u32, u32)], c: char) -> bool {
    let code = c as u32;
    // The last range that starts at `code` or before it.
    let at = ranges.partition_point(|&(from, _)| from <= code);
    at > 0 && code <= ranges[at - 1].1
}

/// Whether the table's Unicode assigns `c`.
fn assigned(c: char) -> bool {
    in_ranges(&unicode::ASSIGNED, c)
}

/// Whether Rust's `to_lowercase`, from the toolchain's tables, gives what
/// the table says.
fn rust_agrees(c: char) -> bool {
    match table_lower(c) {
        Some(lower) => c.to_lowercase().eq(lower.chars()),
        None => c.to_lowercase().eq([c]),
    }
}

/// Whether the capital sigma at `at` becomes a final sigma (ς) rather than
/// σ: when a cased letter comes before it and none after it, the
/// case-ignorable characters in between passed over (Unicode's Final_Sigma,
/// as ICU and so bun read it). What each character is to the sigma comes
/// from the table, recorded from bun's toLowerCase (SIGMA_CASED,
/// SIGMA_PASSED). A lone surrogate ends the search, as a character that is
/// not cased. `None`, to hand back, when the search meets a character
/// outside SIGMA_BLOCKS, the blocks where that is the same in Unicode 13,
/// 15.1, 16 and 17 and in macOS's ICU (Latin up to Latin Extended-B, the
/// combining marks and Greek, Greek Extended and General Punctuation; see
/// rs/unicode.gen.mjs), or one the table's Unicode leaves unassigned.
fn final_sigma(points: &[Result<char, u16>], at: usize) -> Option<bool> {
    let cased_then = |around: &mut dyn Iterator<Item = &Result<char, u16>>| -> Option<bool> {
        for point in around {
            let Ok(c) = *point else { return Some(false) };
            if !in_ranges(&unicode::SIGMA_BLOCKS, c) || !assigned(c) {
                return None;
            }
            if !in_ranges(&unicode::SIGMA_PASSED, c) {
                return Some(in_ranges(&unicode::SIGMA_CASED, c));
            }
        }
        Some(false)
    };
    let before = cased_then(&mut points[..at].iter().rev())?;
    let after = cased_then(&mut points[at + 1..].iter())?;
    Some(before && !after)
}

/// `s.replace(/pattern/g, replacement)` for a pattern that is plain text.
pub fn replace_all(units: &[u16], pattern: &str, replacement: &str) -> Vec<u16> {
    let pattern = self::units(pattern);
    let replacement = self::units(replacement);
    let mut out = Vec::with_capacity(units.len());
    let mut at = 0;
    while at < units.len() {
        if units[at..].starts_with(&pattern) {
            out.extend_from_slice(&replacement);
            at += pattern.len();
        } else {
            out.push(units[at]);
            at += 1;
        }
    }
    out
}

/// `s.split(/\n+/)`, without the empty pieces (every caller drops them).
pub fn lines(units: &[u16]) -> impl Iterator<Item = &[u16]> {
    units.split(|&u| u == 0x0A).filter(|line| !line.is_empty())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn t(units: &[u16]) -> String {
        text(units)
    }

    #[test]
    fn spaces_are_javascripts() {
        let s = units("\u{feff}\u{3000} a\u{a0}\u{2003}\t\nb \u{2028}");
        assert_eq!(t(trim(&s)), "a\u{a0}\u{2003}\t\nb");
        assert_eq!(t(&squeeze(&s)), "a b");
        assert_eq!(t(&collapse_spaces(&units("  a  b "))), " a b ");
        // Not white space for JavaScript: U+0085, U+180E, U+200B.
        assert_eq!(t(trim(&units("\u{85}\u{180e}\u{200b}"))), "\u{85}\u{180e}\u{200b}");
    }

    #[test]
    fn cutting_counts_code_units() {
        let s = units("a😀b");
        assert_eq!(s.len(), 4);
        assert_eq!(t(slice(&s, 0, 2)), "a\u{fffd}");
        assert_eq!(t(slice(&s, 2, 9)), "\u{fffd}b");
        assert_eq!(t(slice(&s, 3, 1)), "");
        assert_eq!(t(slice(&s, 9, 12)), "");
        // Two halves put together again are the emoji again.
        let mut whole = slice(&s, 0, 2).to_vec();
        whole.extend_from_slice(slice(&s, 2, 4));
        assert_eq!(t(&whole), "a😀b");
    }

    #[test]
    fn finding_and_replacing() {
        let s = units("one two one");
        assert_eq!(index_of(&s, &units("one")), Some(0));
        assert_eq!(index_of(&s, &units("two")), Some(4));
        assert_eq!(index_of(&s, &units("three")), None);
        assert_eq!(index_of(&s, &[]), Some(0));
        assert_eq!(index_of(&units("ab"), &units("abc")), None);
        assert_eq!(t(&replace_all(&units("a&amp;amp;b&amp;"), "&amp;", "&")), "a&amp;b&");
        assert_eq!(lines(&units("\n\na\n\nb\n")).map(t).collect::<Vec<_>>(), ["a", "b"]);
    }

    #[test]
    fn lower_case_is_unicodes() {
        let lower = |text: &str| to_lower_case(&units(text)).map(|lower| t(&lower));
        assert_eq!(lower("ÀBC Ünï").as_deref(), Some("àbc ünï"));
        // One unit becomes two.
        assert_eq!(to_lower_case(&units("İ")), Some(units("i\u{307}")));
        // A capital sigma at the end of a word, also behind what is passed
        // over (an apostrophe, a combining mark), and not before a letter.
        assert_eq!(lower("ΟΔΟΣ ΣΟΣ").as_deref(), Some("οδος σος"));
        assert_eq!(lower("ΟΔΟΣ'").as_deref(), Some("οδος'"));
        assert_eq!(lower("Ο\u{301}Σ").as_deref(), Some("ο\u{301}ς"));
        assert_eq!(lower("ΟΣ'Α").as_deref(), Some("οσ'α"));
        assert_eq!(lower("Σ").as_deref(), Some("σ"));
        assert_eq!(lower("ΑΣ\u{2019}").as_deref(), Some("ας\u{2019}"));
        // Beside a character whose part in that may differ between
        // Unicodes (U+0295 changed in 17), or that is not assigned: none.
        assert_eq!(lower("ΑΣ\u{295}"), None);
        assert_eq!(lower("\u{295}Σ"), None);
        assert_eq!(lower("ΑΣ\u{378}"), None);
        // Far from the sigma, such a character is no matter to it.
        assert_eq!(lower("ΑΣ \u{295}").as_deref(), Some("ας \u{295}"));
        let mut half = units("AB");
        half.insert(1, 0xD83D);
        assert_eq!(to_lower_case(&half), Some(vec![0x61, 0xD83D, 0x62]));
        // A lone surrogate ends the search, as a character that is not cased.
        let mut cut = units("AΣ");
        cut.insert(1, 0xD83D);
        assert_eq!(to_lower_case(&cut), Some(vec![0x61, 0xD83D, 0x3C3]));
    }

    #[test]
    fn the_table_is_in_order() {
        assert!(unicode::LOWER.windows(2).all(|w| w[0].0 < w[1].0));
        for ranges in [&unicode::ASSIGNED[..], &unicode::SIGMA_BLOCKS[..], &unicode::SIGMA_CASED[..], &unicode::SIGMA_PASSED[..]] {
            assert!(ranges.iter().all(|&(a, b)| a <= b));
            assert!(ranges.windows(2).all(|w| w[0].1 + 1 < w[1].0));
        }
        assert!(unicode::LOWER.iter().all(|&(code, lower)| char::from_u32(code).is_some_and(assigned) && lower.chars().all(assigned)));
        assert!(assigned('\0') && assigned('\u{16ea0}') && assigned('\u{f0000}') && !assigned('\u{ab6c}') && !assigned('\u{1df40}') && !assigned('\u{10ffff}'));
        // Nothing is both, and nothing is either outside the blocks.
        for c in (0..=0x10ffffu32).filter_map(char::from_u32) {
            let (cased, passed) = (in_ranges(&unicode::SIGMA_CASED, c), in_ranges(&unicode::SIGMA_PASSED, c));
            assert!(!(cased && passed) && (in_ranges(&unicode::SIGMA_BLOCKS, c) || !(cased || passed)), "U+{:04X}", c as u32);
        }
    }

    #[test]
    fn lower_case_comes_from_the_table() {
        // U+AB6C and U+1DF40 are capitals in Unicode 18 (Rust 1.100) and
        // unassigned in bun's 17; U+1C89 came with 16.
        for unsure in ["\u{ab6c}", "\u{1df40}", "\u{1c89}", "\u{10ffff}", "\u{ffff}"] {
            assert_eq!(to_lower_case(&units(unsure)), None, "{unsure:?}");
        }
        // Wherever the binary answers, the answer is the table's, whatever
        // this toolchain's tables say; where they disagree, it answers nothing.
        let mut disagree = Vec::new();
        for c in (0..=0x10ffffu32).filter_map(char::from_u32) {
            let table: String = table_lower(c).map_or_else(|| c.to_string(), String::from);
            if let Some(ours) = to_lower_case(&units(&c.to_string())) {
                assert_eq!(t(&ours), table, "U+{:04X}", c as u32);
            }
            if c.to_lowercase().collect::<String>() != table {
                assert!(lowercase_unsure(c), "U+{:04X}", c as u32);
                disagree.push(format!("U+{:04X}", c as u32));
            }
        }
        let (major, minor, update) = char::UNICODE_VERSION;
        eprintln!("this toolchain's Unicode {major}.{minor}.{update}, the table's {}: they disagree on {} code point(s) {}", unicode::VERSION, disagree.len(), disagree.join(" "));
    }

    /// A string as JSON.stringify writes it, in code units (a lone surrogate
    /// is a \u escape there).
    fn json_units(json: &str) -> Vec<u16> {
        let inner = json.strip_prefix('"').and_then(|rest| rest.strip_suffix('"')).unwrap_or_else(|| panic!("not a JSON string: {json}"));
        let mut out = Vec::new();
        let mut chars = inner.chars();
        while let Some(c) = chars.next() {
            if c != '\\' {
                out.extend(c.encode_utf16(&mut [0; 2]).iter());
                continue;
            }
            match chars.next() {
                Some('u') => out.push(u16::from_str_radix(&chars.by_ref().take(4).collect::<String>(), 16).unwrap()),
                Some('n') => out.push(0x0A),
                Some('r') => out.push(0x0D),
                Some('t') => out.push(0x09),
                Some('b') => out.push(0x08),
                Some('f') => out.push(0x0C),
                Some(c @ ('"' | '\\' | '/')) => out.push(c as u16),
                other => panic!("{other:?} in {json}"),
            }
        }
        out
    }

    /// The alarm: every answer the binary gives is held against the bun at
    /// hand (rs/parity/runtime-lower.ts), every code point alone and beside
    /// a capital sigma, whatever toolchain built the test. When it fails, is
    /// rs/src/unicode.rs what this bun says? `bun rs/unicode.gen.mjs --check`.
    #[test]
    fn lower_case_is_the_runtimes() {
        let dir = std::env::temp_dir().join(format!("onenote-lower-{}", std::process::id()));
        std::fs::create_dir_all(&dir).unwrap();
        let answers = dir.join("answers");
        let script = concat!(env!("CARGO_MANIFEST_DIR"), "/parity/runtime-lower.ts");
        let status = std::process::Command::new("bun")
            .arg(script)
            .arg(&answers)
            .current_dir(&dir)
            .env_clear()
            .env("PATH", std::env::var_os("PATH").unwrap_or_default())
            .env("HOME", &dir)
            .status()
            .expect("bun, to compare with the runtime");
        assert!(status.success(), "{script}: {status}");
        let listed = std::fs::read_to_string(&answers).unwrap();
        std::fs::remove_dir_all(&dir).unwrap();
        let (a, sigma) = (units("A"), units("\u{3a3}"));
        let mut wrong = Vec::new();
        let mut answered = [0usize; 5];
        let mut lines = 0;
        for (code, line) in listed.lines().enumerate() {
            lines += 1;
            let c = match char::from_u32(code as u32) {
                Some(c) => units(&c.to_string()),
                None => vec![code as u16],
            };
            let asked = [c.clone(), [&c[..], &sigma].concat(), [&a[..], &c, &sigma].concat(), [&a[..], &sigma, &c].concat(), [&a[..], &sigma, &c, &a].concat()];
            let theirs: Vec<Vec<u16>> = line.split('\t').map(json_units).collect();
            assert_eq!(theirs.len(), asked.len(), "{line}");
            for (n, (question, answer)) in asked.iter().zip(&theirs).enumerate() {
                if let Some(ours) = to_lower_case(question) {
                    answered[n] += 1;
                    if ours != *answer {
                        wrong.push(format!("U+{code:04X} {:?}: bun {:?}, the binary {:?}", t(question), t(answer), t(&ours)));
                    }
                }
            }
        }
        assert_eq!(lines, 0x110000);
        assert!(wrong.is_empty(), "{} answers differ:\n{}", wrong.len(), wrong.iter().take(60).cloned().collect::<Vec<_>>().join("\n"));
        // Every character of bun's Unicode but the few unsure ones, and the
        // sigma beside those of the blocks it answers in.
        assert!(answered[0] > 290_000 && answered[1..].iter().all(|&n| n > 1_000), "{answered:?}");
        eprintln!("answers compared with bun's toLowerCase: {answered:?} (alone, c + \"\u{3a3}\", \"A\" + c + \"\u{3a3}\", \"A\u{3a3}\" + c, \"A\u{3a3}\" + c + \"A\")");
    }
}
