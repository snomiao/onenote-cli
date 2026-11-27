//! What a search query means. Port of `parseTagsFromQuery`, `buildFtsParam`
//! and `TAG_ALIASES` in src/cache.ts.

use crate::js::{is_space, is_word, squeeze, units};

/// `#alias` → the OneNote `data-tag` values it stands for (TAG_ALIASES, in
/// the order of the TypeScript).
const TAG_ALIASES: [(&str, &[&str]); 23] = [
    ("star", &["star"]),
    ("question", &["question"]),
    ("important", &["important"]),
    ("critical", &["critical"]),
    ("definition", &["definition"]),
    ("idea", &["idea"]),
    ("contact", &["contact"]),
    ("address", &["address"]),
    ("phone", &["phone-number"]),
    ("website", &["web-site-to-visit"]),
    ("password", &["password"]),
    ("remember", &["remember-for-later"]),
    ("book", &["book-to-read"]),
    ("music", &["music-to-listen-to"]),
    ("movie", &["movie-to-see"]),
    ("highlight", &["highlight"]),
    ("meeting", &["schedule-meeting"]),
    ("email", &["send-in-email"]),
    ("callback", &["call-back"]),
    ("discuss", &["discuss-with-person-a", "discuss-with-person-b", "discuss-with-manager"]),
    ("priority1", &["to-do-priority-1"]),
    ("priority2", &["to-do-priority-2"]),
    ("client", &["client-request"]),
];

/// The result of `parseTagsFromQuery`.
pub struct Parsed {
    /// The words left for the full-text index.
    pub fts_query: Vec<u16>,
    pub has_todo: bool,
    pub has_done: bool,
    pub has_checkbox: bool,
    /// `data-tag` values a page must carry, one entry per time a tag was
    /// named (the TypeScript does not remove repeats).
    pub tag_filters: Vec<&'static str>,
}

/// One unit of the text against one letter of a pattern, the way the `i`
/// flag compares them without the `u` flag: ASCII letters in either case, and
/// nothing outside ASCII ever equals an ASCII letter.
fn same_letter(unit: u16, pattern: u8) -> bool {
    u8::try_from(unit).is_ok_and(|byte| byte.eq_ignore_ascii_case(&pattern))
}

/// `\b` at `at`: a word character on exactly one side.
fn boundary(text: &[u16], at: usize) -> bool {
    let before = at > 0 && is_word(text[at - 1]);
    let after = at < text.len() && is_word(text[at]);
    before != after
}

/// Where `/pattern\b/i` (with `\b` in front as well when `word_start`)
/// matches at `at`.
fn matches_at(text: &[u16], at: usize, pattern: &str, word_start: bool) -> bool {
    let pattern = pattern.as_bytes();
    at + pattern.len() <= text.len()
        && pattern.iter().enumerate().all(|(i, &letter)| same_letter(text[at + i], letter))
        && boundary(text, at + pattern.len())
        && (!word_start || boundary(text, at))
}

/// `text.replace(/pattern\b/gi, "")`, and how many times it matched.
fn remove(text: &[u16], pattern: &str, word_start: bool) -> (Vec<u16>, usize) {
    let mut out = Vec::with_capacity(text.len());
    let mut count = 0;
    let mut at = 0;
    while at < text.len() {
        if matches_at(text, at, pattern, word_start) {
            count += 1;
            at += pattern.len();
        } else {
            out.push(text[at]);
            at += 1;
        }
    }
    (out, count)
}

/// `parseTagsFromQuery(query)`
pub fn parse_tags(query: &[u16]) -> Parsed {
    let mut text = query.to_vec();
    let mut flag = |pattern: &str, word_start: bool| {
        let (rest, count) = remove(&text, pattern, word_start);
        text = rest;
        count
    };
    let has_checkbox = flag("#checkbox", false) + flag("tag:checkbox", true) > 0;
    let has_todo = flag("#todo", false) + flag("tag:todo", true) > 0;
    let has_done = flag("#done", false) + flag("tag:done", true) > 0;
    let mut tag_filters = Vec::new();
    for (alias, values) in TAG_ALIASES {
        let times = flag(&format!("#{alias}"), false) + flag(&format!("tag:{alias}"), true);
        for _ in 0..times {
            tag_filters.extend_from_slice(values);
        }
    }
    Parsed { fts_query: squeeze(&text), has_todo, has_done, has_checkbox, tag_filters }
}

/// The operator `/\b(AND|OR|NOT)\b/i` matches at `at`, if any.
fn operator_at(text: &[u16], at: usize) -> Option<&'static str> {
    ["AND", "OR", "NOT"].into_iter().find(|operator| matches_at(text, at, operator, true))
}

/// `buildFtsParam(query)`: what is handed to FTS5 as the MATCH expression.
pub fn build_fts_param(query: &[u16]) -> Vec<u16> {
    if (0..query.len()).any(|at| operator_at(query, at).is_some()) {
        // The query is the user's own expression; only the operators are
        // put in capitals.
        let mut out = Vec::with_capacity(query.len());
        let mut at = 0;
        while at < query.len() {
            match operator_at(query, at) {
                Some(operator) => {
                    out.extend(units(operator));
                    at += operator.len();
                }
                None => {
                    out.push(query[at]);
                    at += 1;
                }
            }
        }
        return out;
    }
    // query.match(/(?:"[^"]*"|[^\s]+)/g): a "quoted phrase", else a run of
    // anything that is not white space. Every bare term is quoted.
    const QUOTE: u16 = b'"' as u16;
    let mut out: Vec<u16> = Vec::new();
    let mut at = 0;
    while at < query.len() {
        if is_space(query[at]) {
            at += 1;
            continue;
        }
        let closing = (query[at] == QUOTE).then(|| query[at + 1..].iter().position(|&u| u == QUOTE)).flatten();
        let end = match closing {
            Some(offset) => at + 1 + offset + 1,
            None => at + query[at..].iter().position(|&u| is_space(u)).unwrap_or(query.len() - at),
        };
        let token = &query[at..end];
        if !out.is_empty() {
            out.push(0x20);
        }
        if token[0] == QUOTE {
            out.extend_from_slice(token);
        } else {
            out.push(QUOTE);
            for &unit in token {
                out.push(unit);
                if unit == QUOTE {
                    out.push(QUOTE);
                }
            }
            out.push(QUOTE);
        }
        at = end;
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::js::text;

    fn parsed(query: &str) -> (String, bool, bool, bool, Vec<&'static str>) {
        let p = parse_tags(&units(query));
        (text(&p.fts_query), p.has_todo, p.has_done, p.has_checkbox, p.tag_filters)
    }

    fn fts(query: &str) -> String {
        text(&build_fts_param(&units(query)))
    }

    // Expected values were printed by the TypeScript's own functions (bun
    // 1.4.0, src/cache.ts at 9af05e1).
    #[test]
    fn tags_are_taken_out_of_the_query() {
        assert_eq!(parsed("plan"), ("plan".into(), false, false, false, vec![]));
        assert_eq!(parsed("#todo buy  milk"), ("buy milk".into(), true, false, false, vec![]));
        assert_eq!(parsed("#TODO #Done tag:checkbox"), ("".into(), true, true, true, vec![]));
        assert_eq!(parsed("#todos tag:todoist xtag:todo"), ("#todos tag:todoist xtag:todo".into(), false, false, false, vec![]));
        assert_eq!(parsed("#star #star tag:phone"), ("".into(), false, false, false, vec!["star", "star", "phone-number"]));
        assert_eq!(
            parsed("#discuss x"),
            ("x".into(), false, false, false, vec!["discuss-with-person-a", "discuss-with-person-b", "discuss-with-manager"])
        );
        assert_eq!(parsed("#phone-number"), ("-number".into(), false, false, false, vec!["phone-number"]));
        assert_eq!(parsed("a#todo-b (tag:star)"), ("a-b ()".into(), true, false, false, vec!["star"]));
        assert_eq!(parsed("#priority1 #priority10"), ("#priority10".into(), false, false, false, vec!["to-do-priority-1"]));
        // Outside ASCII nothing is a word character and nothing folds.
        assert_eq!(parsed("#todoé #ſtar"), ("é #ſtar".into(), true, false, false, vec![]));
    }

    #[test]
    fn terms_are_quoted_for_fts5() {
        assert_eq!(fts("plan"), "\"plan\"");
        assert_eq!(fts("project  plan"), "\"project\" \"plan\"");
        assert_eq!(fts("say \"hello world\" now"), "\"say\" \"hello world\" \"now\"");
        assert_eq!(fts("a\"b c\" \"d"), "\"a\"\"b\" \"c\"\"\" \"d");
        assert_eq!(fts("\"a b\"c"), "\"a b\" \"c\"");
        assert_eq!(fts("c++ foo:bar"), "\"c++\" \"foo:bar\"");
        assert_eq!(fts(""), "");
        assert_eq!(fts("  "), "");
    }

    #[test]
    fn operators_keep_the_query_as_written() {
        assert_eq!(fts("python or javascript"), "python OR javascript");
        assert_eq!(fts("a And b nOt c"), "a AND b NOT c");
        assert_eq!(fts("android organ nothing"), "\"android\" \"organ\" \"nothing\"");
        assert_eq!(fts("x-or-y"), "x-OR-y");
        assert_eq!(fts("or_ _and and2"), "\"or_\" \"_and\" \"and2\"");
    }
}
