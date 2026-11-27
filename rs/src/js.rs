//! JavaScript's strings, as far as the ported commands use them.
//!
//! The TypeScript counts, cuts and searches text in UTF-16 code units
//! (`length`, `slice`, `indexOf`, regular expressions without the `u` flag),
//! so an emoji is two units and a cut can go through the middle of one. To
//! print the same bytes, text is kept here as code units too and turned into
//! UTF-8 only when a whole line is written, the way `console.log` does it:
//! half an emoji becomes U+FFFD.

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

/// `s.toLowerCase()`. Both sides follow Unicode's default lower-casing (one
/// letter can become two units, and a capital sigma depends on where in the
/// word it stands). Half an emoji is neither a letter nor part of a word and
/// stays as it is.
pub fn to_lower_case(units: &[u16]) -> Vec<u16> {
    let mut out = Vec::with_capacity(units.len());
    let mut run = String::new();
    for decoded in char::decode_utf16(units.iter().copied()) {
        match decoded {
            Ok(c) => run.push(c),
            Err(lone) => {
                out.extend(run.to_lowercase().encode_utf16());
                run.clear();
                out.push(lone.unpaired_surrogate());
            }
        }
    }
    out.extend(run.to_lowercase().encode_utf16());
    out
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
        assert_eq!(t(&to_lower_case(&units("ÀBC Ünï"))), "àbc ünï");
        // One unit becomes two.
        assert_eq!(to_lower_case(&units("İ")), units("i\u{307}"));
        // A capital sigma at the end of a word.
        assert_eq!(t(&to_lower_case(&units("ΟΔΟΣ ΣΟΣ"))), "οδος σος");
        let mut half = units("AB");
        half.insert(1, 0xD83D);
        assert_eq!(to_lower_case(&half), [0x61, 0xD83D, 0x62]);
    }
}
