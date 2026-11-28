//! `onenote search <query>` against the local index.
//!
//! Port of the `search` command in src/index.ts and of `isCacheEmpty`,
//! `searchLocal` and `cleanSnippet` in src/cache.ts, for the case a synced
//! machine is in: `<package>/.onenote/search.db` exists and has pages.
//!
//! Everything is worked out before a byte is written, so whatever turns out
//! not to be implemented here can still be handed to the JavaScript:
//!
//! - `--online` (Microsoft Graph), and any spelling of the command line other
//!   than the plain ones (see `parse_args`);
//! - no index, or an empty one: the TypeScript then syncs, or reads the
//!   section files instead;
//! - an error from SQLite, a query FTS5 does not accept included: the
//!   TypeScript then falls back to reading the section files;
//! - a row that is not what `onenote sync` writes (a NULL title, tag lines
//!   that are not a list), or a date in a notation only `Date` reads;
//! - a `TZ` that is not the name of a zone, a zone at a time for which the
//!   runtime's time zone database says something else than the one compiled
//!   in, a time before 1970 or after 2099, and an index last written in the
//!   last millisecond of a second (see time.rs for each).

use std::path::Path;

use rusqlite::types::{Value, ValueRef};
use rusqlite::{Connection, OpenFlags};

use crate::js::{self, index_of, slice, squeeze, text, to_lower_case, trim, units};
use crate::query::{build_fts_param, parse_tags, Parsed};
use crate::{time, HandBack};

/// The command line of a search, in the form the handler sees it.
#[derive(Debug, Default, PartialEq)]
pub struct Args {
    pub query: String,
    pub notebook: Option<String>,
    pub section: Option<String>,
    pub account: Option<String>,
    pub since: Option<String>,
    pub limit: Option<u32>,
    pub offset: Option<u32>,
    pub max_lines: Option<u32>,
    pub group: bool,
    pub compact: bool,
}

/// A count as yargs turns it into a number without surprises: decimal digits,
/// no sign, no leading zero, small enough to be bound as an integer.
fn count(text: &str) -> Option<u32> {
    let plain = text == "0" || (!text.starts_with('0') && !text.is_empty());
    (plain && text.len() <= 9 && text.bytes().all(|b| b.is_ascii_digit())).then(|| text.parse().ok()).flatten()
}

/// The arguments after `search`, if they are written the plain way:
/// one query, and each option at most once as `--name value`, `--name=value`
/// or `-n value` (bare `--group`/`-g`, `--compact`/`-c`). Everything else is
/// yargs's to read: `--online`, `--help`, unknown or repeated options,
/// `-n5`, `-gc`, `--no-group`, `--group=true`, camel-case names, `--`, empty
/// arguments, values starting with a dash, a second word, no query at all.
pub fn parse_args(args: &[String]) -> Option<Args> {
    let mut parsed = Args::default();
    let mut query: Option<&str> = None;
    let mut seen: Vec<&str> = Vec::new();
    let mut at = 0;
    while at < args.len() {
        let arg = args[at].as_str();
        at += 1;
        if arg.is_empty() {
            return None;
        }
        if !arg.starts_with('-') {
            if query.replace(arg).is_some() {
                return None;
            }
            continue;
        }
        let (name, inline) = match arg.strip_prefix("--") {
            Some(long) => match long.split_once('=') {
                Some((name, value)) => (name, Some(value)),
                None => (long, None),
            },
            None => {
                let name = match arg {
                    "-n" => "notebook",
                    "-s" => "section",
                    "-a" => "account",
                    "-l" => "limit",
                    "-p" => "offset",
                    "-m" => "max-lines",
                    "-g" => "group",
                    "-c" => "compact",
                    _ => return None,
                };
                (name, None)
            }
        };
        if seen.contains(&name) {
            return None;
        }
        seen.push(name);
        if matches!(name, "group" | "compact") {
            // yargs takes a following `true` or `false` as the flag's value.
            if inline.is_some() || args.get(at).is_some_and(|next| next == "true" || next == "false") {
                return None;
            }
            match name {
                "group" => parsed.group = true,
                _ => parsed.compact = true,
            }
            continue;
        }
        if !matches!(name, "notebook" | "section" | "account" | "since" | "limit" | "offset" | "max-lines") {
            return None;
        }
        let value = match inline {
            Some(value) => value,
            None => {
                let value = args.get(at)?.as_str();
                at += 1;
                if value.starts_with('-') {
                    return None;
                }
                value
            }
        };
        if value.is_empty() {
            return None;
        }
        match name {
            "notebook" => parsed.notebook = Some(value.to_string()),
            "section" => parsed.section = Some(value.to_string()),
            "account" => parsed.account = Some(value.to_string()),
            "since" => parsed.since = Some(value.to_string()),
            "limit" => parsed.limit = Some(count(value)?),
            "offset" => parsed.offset = Some(count(value)?),
            _ => parsed.max_lines = Some(count(value)?),
        }
    }
    parsed.query = query?.to_string();
    Some(parsed)
}

/// `value.replace(/^["']|["']$/g, "")`: one quote off the front, one off the
/// end.
fn strip_quotes(value: &str) -> &str {
    let quote = |c: char| c == '"' || c == '\'';
    let value = value.strip_prefix(quote).unwrap_or(value);
    value.strip_suffix(quote).unwrap_or(value)
}

/// One `{ tag, text }` of a page's tag lines.
struct TagLine {
    tag: String,
    text: Vec<u16>,
}

/// One result, as `searchLocal` returns it.
struct Page {
    title: Vec<u16>,
    /// The excerpt `searchLocal` cut out (`cleanSnippet` in src/cache.ts),
    /// or the body's first line for a query that has only tags.
    body: Vec<u16>,
    section: Vec<u16>,
    notebook: Vec<u16>,
    web_url: Option<Vec<u16>>,
    account: Option<Vec<u16>>,
    last_modified: Option<String>,
    tag_lines: Vec<TagLine>,
}

fn ascii_ratio(line: &[u16]) -> f64 {
    if line.is_empty() {
        return 0.0;
    }
    line.iter().filter(|&&u| (0x20..=0x7E).contains(&u)).count() as f64 / line.len() as f64
}

/// `cleanSnippet(raw, query)` in src/cache.ts: the line of the body that
/// holds the query, cut to thirty characters either side. `HandBack` where
/// the lower case of the query or of a line is not certain.
fn excerpt(raw: &[u16], query: &[u16]) -> Result<Vec<u16>, HandBack> {
    // raw.replace(/<[^>]{0,300}>/g, " ")
    let mut stripped = Vec::with_capacity(raw.len());
    let mut at = 0;
    while at < raw.len() {
        if raw[at] == u16::from(b'<') {
            let close = raw[at + 1..].iter().position(|&u| u == u16::from(b'>'));
            if let Some(inside) = close.filter(|&inside| inside <= 300) {
                stripped.push(0x20);
                at += inside + 2;
                continue;
            }
        }
        stripped.push(raw[at]);
        at += 1;
    }
    for (entity, plain) in [("&amp;", "&"), ("&lt;", "<"), ("&gt;", ">"), ("&nbsp;", " ")] {
        stripped = js::replace_all(&stripped, entity, plain);
    }
    // .replace(/[\x00-\x09\x0B-\x1F\x7F]/g, " "): line feeds stay.
    for unit in &mut stripped {
        if matches!(*unit, 0x00..=0x09 | 0x0B..=0x1F | 0x7F) {
            *unit = 0x20;
        }
    }
    let lower_query = to_lower_case(query).ok_or(HandBack)?;
    let mut matching: Vec<Vec<u16>> = Vec::new();
    for line in js::lines(&stripped).map(squeeze) {
        if !line.is_empty() && index_of(&to_lower_case(&line).ok_or(HandBack)?, &lower_query).is_some() {
            matching.push(line);
        }
    }
    // matchingLines.sort((a, b) => asciiRatio(b) - asciiRatio(a))[0]: the
    // first of the lines with the highest share of ASCII.
    let Some(best) = matching.iter().reduce(|best, line| if ascii_ratio(line) > ascii_ratio(best) { line } else { best }) else {
        return Ok(Vec::new());
    };
    if ascii_ratio(best) < 0.15 {
        return Ok(units("[binary content, no clean preview]"));
    }
    // The position is found in the lower-cased line and used on the line
    // itself, as in the TypeScript, although the two can differ in length.
    let found = index_of(&to_lower_case(best).ok_or(HandBack)?, &lower_query).unwrap_or(0);
    let start = found.saturating_sub(30);
    let end = (found + query.len() + 30).min(best.len());
    let mut out = Vec::new();
    if start > 0 {
        out.extend(units("..."));
    }
    out.extend_from_slice(trim(slice(best, start, end)));
    if end < best.len() {
        out.extend(units("..."));
    }
    Ok(out)
}

/// A TEXT column. JavaScript would also be handed numbers and blobs, and
/// would decode bytes that are not UTF-8 in its own way.
fn column_text(value: ValueRef<'_>) -> Result<Option<&str>, HandBack> {
    match value {
        ValueRef::Null => Ok(None),
        ValueRef::Text(bytes) => match std::str::from_utf8(bytes) {
            Ok(text) if !text.contains('\0') => Ok(Some(text)),
            _ => Err(HandBack),
        },
        _ => Err(HandBack),
    }
}

/// `JSON.parse(r.tag_lines)`, for what `onenote sync` stores there: a list
/// of `{ tag, text }`.
fn tag_lines(stored: Option<&str>) -> Result<Vec<TagLine>, HandBack> {
    let Some(stored) = stored.filter(|stored| !stored.is_empty()) else {
        return Ok(Vec::new());
    };
    match serde_json::from_str::<serde_json::Value>(stored).map_err(|_| HandBack)? {
        serde_json::Value::Null => Ok(Vec::new()),
        serde_json::Value::Array(lines) => lines
            .iter()
            .map(|line| match (line.get("tag"), line.get("text")) {
                (Some(serde_json::Value::String(tag)), Some(serde_json::Value::String(text))) if line.is_object() => {
                    Ok(TagLine { tag: tag.clone(), text: units(text) })
                }
                _ => Err(HandBack),
            })
            .collect(),
        _ => Err(HandBack),
    }
}

/// What the handler needs from the command line, cleaned up as it does.
struct Request<'a> {
    parsed: &'a Parsed,
    notebook: Option<&'a str>,
    section: Option<&'a str>,
    account: Option<&'a str>,
    since: Option<&'a str>,
    limit: u32,
    offset: u32,
}

/// Whether the runtime's SQLite can open the index read-only, as far as this
/// binary can tell without asking it.
///
/// On macOS bun:sqlite is the system's SQLite, which cannot open a database
/// in WAL mode read-only while the `-wal` file beside it is missing
/// (SQLITE_CANTOPEN; a missing `-shm` file it makes itself), and the
/// TypeScript then takes the cache for empty and starts a sync. The SQLite
/// compiled in here, like the one inside bun on Linux, makes both files and
/// reads on. So on macOS the binary leaves the search to the JavaScript
/// whenever the `-wal` file is not there, whatever the journal mode. A sync
/// leaves it there (its last connection is never closed); an index that
/// something closed cleanly has none.
fn openable_by_the_runtime(file: &Path) -> Result<(), HandBack> {
    if cfg!(target_os = "macos") {
        let mut wal = file.as_os_str().to_owned();
        wal.push("-wal");
        if !std::fs::symlink_metadata(&wal).is_ok_and(|meta| meta.is_file()) {
            return Err(HandBack);
        }
    }
    Ok(())
}

/// `isCacheEmpty()` and `searchLocal()`, where the index answers.
fn search_index(file: &Path, request: &Request<'_>) -> Result<Vec<Page>, HandBack> {
    openable_by_the_runtime(file)?;
    let db = Connection::open_with_flags(file, OpenFlags::SQLITE_OPEN_READ_ONLY).map_err(|_| HandBack)?;
    let pages: i64 = db.query_row("SELECT COUNT(*) as c FROM pages", [], |row| row.get(0)).map_err(|_| HandBack)?;
    if pages <= 0 {
        // The TypeScript looks through the cache directory next, and syncs
        // when that is empty too.
        return Err(HandBack);
    }
    let parsed = request.parsed;
    let fts = !parsed.fts_query.is_empty();
    let prefix = if fts { "p." } else { "" };
    let mut conditions: Vec<String> = Vec::new();
    let mut params: Vec<Value> = Vec::new();
    if fts {
        conditions.push("pages_fts MATCH ?".to_string());
        params.push(Value::Text(String::from_utf16(&build_fts_param(&parsed.fts_query)).map_err(|_| HandBack)?));
    }
    if let Some(notebook) = request.notebook {
        conditions.push(format!("{prefix}notebook LIKE ?"));
        params.push(Value::Text(format!("%{notebook}%")));
    }
    if let Some(section) = request.section {
        conditions.push(format!("{prefix}section LIKE ?"));
        params.push(Value::Text(format!("%{section}%")));
    }
    if let Some(account) = request.account {
        conditions.push(format!("{prefix}account = ?"));
        params.push(Value::Text(account.to_string()));
    }
    if parsed.has_todo {
        conditions.push(format!("{prefix}has_todo = 1"));
    }
    if parsed.has_done {
        conditions.push(format!("{prefix}has_done = 1"));
    }
    if parsed.has_checkbox {
        conditions.push(format!("({prefix}has_todo = 1 OR {prefix}has_done = 1)"));
    }
    for tag in &parsed.tag_filters {
        conditions.push(format!("{prefix}tags LIKE ?"));
        params.push(Value::Text(format!("%|{tag}|%")));
    }
    if let Some(since) = request.since {
        conditions.push(format!("{prefix}last_modified >= ?"));
        params.push(Value::Text(since.to_string()));
    }
    let sql = if fts {
        format!(
            "SELECT p.section, p.notebook, p.title, p.body, p.web_url, p.page_guid, p.tag_lines, p.account, p.last_modified
         FROM pages_fts f JOIN pages p ON f.rowid = p.id
         WHERE {}
         ORDER BY rank, (p.account IS NULL), p.account, p.last_modified DESC, p.notebook, p.section, p.title
         LIMIT ? OFFSET ?",
            conditions.join(" AND ")
        )
    } else {
        let filter = if conditions.is_empty() { String::new() } else { format!("WHERE {}", conditions.join(" AND ")) };
        format!(
            "SELECT section, notebook, title, body, web_url, page_guid, tag_lines, account, last_modified FROM pages {filter}
         ORDER BY (tag_lines IS NOT NULL) DESC, (account IS NULL), account, last_modified DESC, notebook, section, title
         LIMIT ? OFFSET ?"
        )
    };
    params.push(Value::Integer(i64::from(request.limit)));
    params.push(Value::Integer(i64::from(request.offset)));

    let mut found = Vec::new();
    {
        let mut statement = db.prepare(&sql).map_err(|_| HandBack)?;
        let mut rows = statement.query(rusqlite::params_from_iter(params.iter())).map_err(|_| HandBack)?;
        while let Some(row) = rows.next().map_err(|_| HandBack)? {
            let column = |at: usize| row.get_ref(at).map_err(|_| HandBack).and_then(column_text);
            let section = column(0)?.ok_or(HandBack)?;
            let notebook = column(1)?.ok_or(HandBack)?;
            let title = column(2)?.ok_or(HandBack)?;
            let body = units(column(3)?.ok_or(HandBack)?);
            found.push(Page {
                title: units(title),
                body: if fts {
                    excerpt(&body, &parsed.fts_query)?
                } else {
                    // r.body.split("\n")[0]
                    body.split(|&u| u == 0x0A).next().unwrap_or_default().to_vec()
                },
                section: units(section),
                notebook: units(notebook),
                web_url: column(4)?.map(units),
                tag_lines: tag_lines(column(6)?)?,
                account: column(7)?.map(units),
                last_modified: column(8)?.map(str::to_string),
            });
        }
    }
    db.close().map_err(|_| HandBack)?;
    Ok(found)
}

/// The colour helpers of src/index.ts. Off a terminal they add nothing, and
/// a link is written the Markdown way.
struct Paint {
    tty: bool,
}

impl Paint {
    fn wrap(&self, code: &str, content: &[u16]) -> Vec<u16> {
        if !self.tty {
            return content.to_vec();
        }
        let mut out = units(&format!("\x1b[{code}m"));
        out.extend_from_slice(content);
        out.extend(units("\x1b[0m"));
        out
    }
    fn bold(&self, content: &[u16]) -> Vec<u16> {
        self.wrap("1", content)
    }
    fn dim(&self, content: &[u16]) -> Vec<u16> {
        self.wrap("2", content)
    }
    fn yellow(&self, content: &[u16]) -> Vec<u16> {
        self.wrap("33", content)
    }
    fn cyan(&self, content: &[u16]) -> Vec<u16> {
        self.wrap("36", content)
    }
    fn green(&self, content: &[u16]) -> Vec<u16> {
        self.wrap("32", content)
    }
    fn link(&self, url: &[u16], label: &[u16]) -> Vec<u16> {
        let (open, middle, close) = if self.tty { ("\x1b]8;;", "\x1b\\", "\x1b]8;;\x1b\\") } else { ("[", "](", ")") };
        let (first, second) = if self.tty { (url, label) } else { (label, url) };
        let mut out = units(open);
        out.extend_from_slice(first);
        out.extend(units(middle));
        out.extend_from_slice(second);
        out.extend(units(close));
        out
    }
}

/// TAG_EMOJI in src/index.ts.
fn tag_icon(tag: &str) -> &str {
    match tag {
        "star" => "\u{2605}",
        "question" => "?",
        "important" => "\u{2757}",
        "critical" => "\u{1F534}",
        "definition" => "\u{1F4D6}",
        "idea" => "\u{1F4A1}",
        "contact" => "\u{1F464}",
        "address" => "\u{1F3E0}",
        "phone-number" => "\u{1F4DE}",
        "web-site-to-visit" => "\u{1F310}",
        "password" => "\u{1F511}",
        "remember-for-later" => "\u{1F514}",
        "book-to-read" => "\u{1F4DA}",
        "music-to-listen-to" => "\u{1F3B5}",
        "movie-to-see" => "\u{1F3AC}",
        "highlight" => "\u{1F58D}",
        "schedule-meeting" => "\u{1F4C5}",
        "send-in-email" => "\u{1F4E7}",
        "call-back" => "\u{1F4F2}",
        "to-do-priority-1" => "P1",
        "to-do-priority-2" => "P2",
        "client-request" => "\u{1F454}",
        other => other,
    }
}

const TODO: &str = "\u{2610}";
const DONE: &str = "\u{2611}";

/// The lines the handler prints, as code units until a line is complete.
struct Output {
    text: String,
}

impl Output {
    fn line(&mut self, parts: &[&[u16]]) {
        let whole: Vec<u16> = parts.concat();
        self.text.push_str(&text(&whole));
        self.text.push('\n');
    }
}

/// Everything about the printing that is the same for every page.
struct Renderer<'a> {
    paint: Paint,
    parsed: &'a Parsed,
    lower_query: Vec<u16>,
    tag_prefix: Vec<u16>,
    max_lines: usize,
    out: Output,
}

impl Renderer<'_> {
    /// `getMatchingLines(r)`
    fn matching_lines<'p>(&self, page: &'p Page) -> Vec<&'p TagLine> {
        let parsed = self.parsed;
        let wanted = |line: &&TagLine| {
            if parsed.has_todo {
                line.tag == "to-do"
            } else if parsed.has_done {
                line.tag == "to-do:completed"
            } else if parsed.has_checkbox {
                line.tag == "to-do" || line.tag == "to-do:completed"
            } else {
                parsed.tag_filters.contains(&line.tag.as_str())
            }
        };
        page.tag_lines.iter().filter(wanted).collect()
    }

    /// `renderPageBody(r, indent)`. `HandBack` where the lower case of the
    /// body is not certain.
    fn page_body(&mut self, page: &Page, indent: &str) -> Result<(), HandBack> {
        let indent = units(indent);
        let matching = self.matching_lines(page);
        let body = display_body(&page.body);
        let query_length = self.parsed.fts_query.len();
        let found = if self.lower_query.is_empty() { None } else { index_of(&to_lower_case(&body).ok_or(HandBack)?, &self.lower_query) };
        if let Some(found) = found {
            let start = found.saturating_sub(40);
            let end = (found + query_length + 80).min(body.len());
            let hit = self.paint.yellow(&self.paint.bold(slice(&body, found, found + query_length)));
            let dots = units("...");
            let empty: &[u16] = &[];
            self.out.line(&[
                &indent,
                &self.tag_prefix,
                if start > 0 { &dots } else { empty },
                slice(&body, start, found),
                &hit,
                slice(&body, found + query_length, end),
                if end < body.len() { &dots } else { empty },
            ]);
        } else if !matching.is_empty() {
            for line in matching.iter().take(self.max_lines) {
                let icon = match line.tag.as_str() {
                    "to-do" => self.paint.yellow(&self.paint.bold(&units(TODO))),
                    "to-do:completed" => self.paint.green(&self.paint.bold(&units(DONE))),
                    tag => self.paint.bold(&units(tag_icon(tag))),
                };
                self.out.line(&[&indent, &icon, &units(" "), &line.text]);
            }
            if matching.len() > self.max_lines {
                let more = self.paint.dim(&units(&format!("... and {} more", matching.len() - self.max_lines)));
                self.out.line(&[&indent, &more]);
            }
        } else {
            let snippet = (body.len() >= 20).then(|| slice(&body, 0, 120));
            if !self.tag_prefix.is_empty() {
                let parsed = self.parsed;
                let label = if parsed.has_todo {
                    "(unchecked items)".to_string()
                } else if parsed.has_done {
                    "(completed items)".to_string()
                } else if parsed.has_checkbox {
                    "(checkbox items)".to_string()
                } else {
                    format!("({})", parsed.tag_filters.join(","))
                };
                let label = units(&label);
                self.out.line(&[&indent, &self.tag_prefix, snippet.unwrap_or(&label)]);
            } else if let Some(snippet) = snippet {
                self.out.line(&[&indent, snippet]);
            }
        }
        Ok(())
    }

    /// A page's title, as a link when the page has an address.
    fn title(&self, page: &Page, painted: &[u16]) -> Vec<u16> {
        match page.web_url.as_deref().filter(|url| !url.is_empty()) {
            Some(url) => self.paint.link(url, painted),
            None => painted.to_vec(),
        }
    }
}

/// `cleanSnippet` in the handler: the excerpt as one line, or nothing when
/// too little of it is readable.
fn display_body(body: &[u16]) -> Vec<u16> {
    let spaced: Vec<u16> = body.iter().map(|&u| if u <= 0x1F || (0x7F..=0x9F).contains(&u) { 0x20 } else { u }).collect();
    let line = squeeze(&spaced);
    if line.len() < 3 {
        return Vec::new();
    }
    let length = line.len() as f64;
    let ascii = line.iter().filter(|&&u| (0x20..=0x7E).contains(&u)).count();
    let cjk = line.iter().filter(|&&u| matches!(u, 0x4E00..=0x9FFF | 0x3000..=0x30FF | 0xAC00..=0xD7AF)).count();
    if ((ascii + cjk) as f64) / length < 0.5 {
        return Vec::new();
    }
    if ascii as f64 / length >= 0.4 || (cjk >= 3 && cjk as f64 / length >= 0.4) {
        line
    } else {
        Vec::new()
    }
}

/// The filter behind `cleanResults`: attachments (`.jpg`) and titles that
/// are mostly unreadable are counted but not shown.
fn shown(title: &[u16]) -> bool {
    let alnum = |u: &u16| u8::try_from(*u).is_ok_and(|b| b.is_ascii_alphanumeric());
    let trimmed = trim(title);
    if (3..=6).contains(&trimmed.len()) && trimmed[0] == u16::from(b'.') && trimmed[1..].iter().all(alnum) {
        return false;
    }
    let printable = title
        .iter()
        .filter(|&&u| matches!(u, 0x20..=0x7E | 0x3000..=0x30FF | 0x4E00..=0x9FFF | 0xAC00..=0xD7AF | 0xFF00..=0xFFEF | 0xA0..=0x24F))
        .count();
    printable >= 3 && printable as f64 / title.len() as f64 >= 0.4
}

/// Groups in the order they were first seen, as a JavaScript `Map` keeps
/// them.
struct Ordered<K, V> {
    entries: Vec<(K, V)>,
}

impl<K: PartialEq, V: Default> Ordered<K, V> {
    fn new() -> Self {
        Ordered { entries: Vec::new() }
    }
    fn entry(&mut self, key: K) -> &mut V {
        let at = match self.entries.iter().position(|(k, _)| *k == key) {
            Some(at) => at,
            None => {
                self.entries.push((key, V::default()));
                self.entries.len() - 1
            }
        };
        &mut self.entries[at].1
    }
}

impl<K: PartialEq, V: Default> Default for Ordered<K, V> {
    fn default() -> Self {
        Self::new()
    }
}

type Sections<'p> = Ordered<&'p [u16], Vec<&'p Page>>;
type Notebooks<'p> = Ordered<&'p [u16], Sections<'p>>;

/// The whole command: what it writes to standard output. `HandBack` for
/// everything this file's header lists.
pub fn run(args: &Args, package_root: &Path, tty: bool) -> Result<String, HandBack> {
    let query = units(strip_quotes(&args.query));
    let parsed = parse_tags(&query);
    let limit = args.limit.unwrap_or(100);
    let offset = args.offset.unwrap_or(0);
    let filter = |value: &'_ Option<String>| -> Option<String> {
        value.as_deref().map(strip_quotes).filter(|value| !value.is_empty()).map(str::to_string)
    };
    let (notebook, section, account) = (filter(&args.notebook), filter(&args.section), filter(&args.account));
    let since = match args.since.as_deref() {
        Some(value) => Some(time::since(value, jiff::Timestamp::now())?).filter(|since| !since.is_empty()),
        None => None,
    };
    // The time zone is needed for the last line; find out now whether it
    // is one this binary reads the way the JavaScript runtime does.
    let zone = time::local_zone()?;

    let index = package_root.join(".onenote").join("search.db");
    let request = Request {
        parsed: &parsed,
        notebook: notebook.as_deref(),
        section: section.as_deref(),
        account: account.as_deref(),
        since: since.as_deref(),
        limit,
        offset,
    };
    let results = search_index(&index, &request)?;
    if results.is_empty() {
        return Ok("No results found.\n".to_string());
    }

    let paint = Paint { tty };
    let tag_prefix = if parsed.has_todo {
        paint.yellow(&paint.bold(&units(&format!("{TODO} "))))
    } else if parsed.has_done {
        paint.green(&paint.bold(&units(&format!("{DONE} "))))
    } else if parsed.has_checkbox {
        [paint.yellow(&paint.bold(&units(TODO))), paint.green(&paint.bold(&units(&format!("{DONE} "))))].concat()
    } else if !parsed.tag_filters.is_empty() {
        // [...new Set(tagFilters.map((t) => TAG_EMOJI[t] ?? t))].join("")
        let mut icons: Vec<&str> = Vec::new();
        for tag in &parsed.tag_filters {
            if !icons.contains(&tag_icon(tag)) {
                icons.push(tag_icon(tag));
            }
        }
        paint.bold(&units(&format!("{} ", icons.concat())))
    } else {
        Vec::new()
    };
    let mut renderer = Renderer {
        paint,
        parsed: &parsed,
        lower_query: to_lower_case(&parsed.fts_query).ok_or(HandBack)?,
        tag_prefix,
        max_lines: args.max_lines.unwrap_or(5) as usize,
        out: Output { text: String::new() },
    };

    let clean: Vec<&Page> = results.iter().filter(|page| shown(&page.title)).collect();
    // Accounts are named only when the results come from more than one;
    // pages without an account count as one of them.
    let mut accounts: Vec<&[u16]> = Vec::new();
    for page in &clean {
        let account = page.account.as_deref().unwrap_or_default();
        if !accounts.contains(&account) {
            accounts.push(account);
        }
    }
    let show_accounts = accounts.len() > 1;
    let named_account = |page: &Page| page.account.clone().filter(|account| show_accounts && !account.is_empty());
    // ` [${fmtDate(r.lastModified)}]`
    let date = |page: &Page| -> Result<Option<String>, HandBack> {
        match page.last_modified.as_deref().filter(|modified| !modified.is_empty()) {
            Some(modified) => Ok(Some(format!(" [{}]", time::utc_day(modified)?))),
            None => Ok(None),
        }
    };

    if args.group {
        let local = units("(local)");
        let mut by_account: Ordered<&[u16], Notebooks> = Ordered::new();
        for page in &clean {
            let account = page.account.as_deref().unwrap_or(&local);
            by_account.entry(account).entry(&page.notebook).entry(&page.section).push(page);
        }
        let render_notebooks = |renderer: &mut Renderer, notebooks: &Notebooks, notebook_indent: &str, page_indent: &str, body_indent: &str| -> Result<(), HandBack> {
            for (notebook, sections) in &notebooks.entries {
                let name = renderer.paint.bold(notebook);
                renderer.out.line(&[&units(notebook_indent), &name]);
                for (section, pages) in &sections.entries {
                    let name = renderer.paint.dim(section);
                    renderer.out.line(&[&units(notebook_indent), &units("  "), &name]);
                    for page in pages {
                        let title = renderer.title(page, &renderer.paint.bold(&page.title));
                        renderer.out.line(&[&units(page_indent), &title]);
                        renderer.page_body(page, body_indent)?;
                    }
                }
            }
            Ok(())
        };
        if show_accounts {
            for (account, notebooks) in &by_account.entries {
                let total: usize = notebooks.entries.iter().map(|(_, sections)| sections.entries.iter().map(|(_, pages)| pages.len()).sum::<usize>()).sum();
                let mut heading = units("\u{1F4D8} ");
                heading.extend_from_slice(account);
                let heading = renderer.paint.bold(&renderer.paint.cyan(&heading));
                let pages = renderer.paint.dim(&units(&format!("  ({total} page{})", if total == 1 { "" } else { "s" })));
                renderer.out.line(&[&heading, &pages]);
                render_notebooks(&mut renderer, notebooks, "  ", "      ", "        ")?;
                renderer.out.line(&[]);
            }
        } else {
            // One account, or none: no account heading, one level less.
            let mut merged: Notebooks = Ordered::new();
            for (_, notebooks) in &by_account.entries {
                for (notebook, sections) in &notebooks.entries {
                    for (section, pages) in &sections.entries {
                        merged.entry(notebook).entry(section).extend(pages.iter().copied());
                    }
                }
            }
            render_notebooks(&mut renderer, &merged, "", "    ", "      ")?;
        }
    } else if args.compact {
        // One line per page: [count] title — section | notebook [date]
        for page in &clean {
            let matching = renderer.matching_lines(page).len();
            let paint = &renderer.paint;
            let count = if matching > 0 {
                let mut count = paint.yellow(&paint.bold(&units(&format!("[{matching}]"))));
                count.push(0x20);
                count
            } else {
                Vec::new()
            };
            let title = renderer.title(page, &paint.bold(&paint.cyan(&page.title)));
            let place = paint.dim(&[page.section.as_slice(), &units(" | "), &page.notebook].concat());
            let account = named_account(page).map(|account| paint.dim(&[units(" | "), account].concat())).unwrap_or_default();
            let date = date(page)?.map(|date| paint.dim(&units(&date))).unwrap_or_default();
            let dash = paint.dim(&units("\u{2014}"));
            let space = units(" ");
            renderer.out.line(&[&count, &title, &space, &dash, &space, &place, &account, &date]);
        }
    } else {
        // Flat: pages that follow each other in the same section of the same
        // notebook and account share one heading.
        let mut last_heading: Option<Vec<u16>> = None;
        for page in &clean {
            // `${r.section}|${r.notebook}|${r.account ?? ""}`
            let bar = units("|");
            let heading = [page.section.as_slice(), &bar, &page.notebook, &bar, page.account.as_deref().unwrap_or_default()].concat();
            if last_heading.as_ref() != Some(&heading) {
                if last_heading.is_some() {
                    renderer.out.line(&[]);
                }
                let paint = &renderer.paint;
                let account = match named_account(page) {
                    Some(account) => [units(" "), paint.dim(&units("|")), units(" "), paint.dim(&account)].concat(),
                    None => Vec::new(),
                };
                let line = [paint.bold(&page.notebook), units(" "), paint.dim(&units("\u{203A}")), units(" "), paint.dim(&page.section), account].concat();
                renderer.out.line(&[&line]);
                last_heading = Some(heading);
            }
            let paint = &renderer.paint;
            let title = renderer.title(page, &paint.bold(&paint.cyan(&page.title)));
            let date = date(page)?.map(|date| paint.dim(&units(&date))).unwrap_or_default();
            renderer.out.line(&[&units("  "), &title, &date]);
            renderer.page_body(page, "    ")?;
        }
    }

    // The index's modification time is what "last synced" shows.
    let synced = match std::fs::metadata(&index).and_then(|meta| meta.modified()) {
        Ok(modified) => format!(" \u{2022} last synced: {}", time::locale_string(time::modified_second(modified)?, &zone)?),
        Err(_) => return Err(HandBack),
    };
    let page_info = if offset > 0 {
        format!(" (offset {offset})")
    } else if results.len() == limit as usize {
        format!(" \u{2022} use --offset {} for more", offset + limit)
    } else {
        String::new()
    };
    let last = renderer.paint.green(&units(&format!("{} page-level results found{page_info}{synced}.", results.len())));
    renderer.out.line(&[&last]);
    Ok(renderer.out.text)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn args(list: &[&str]) -> Option<Args> {
        parse_args(&list.iter().map(|arg| arg.to_string()).collect::<Vec<_>>())
    }

    #[test]
    fn plain_command_lines_are_taken() {
        assert_eq!(args(&["plan"]), Some(Args { query: "plan".into(), ..Args::default() }));
        assert_eq!(
            args(&["-n", "Work", "plan", "--section=Meet=ings", "-l", "20", "--offset", "0", "-g", "--compact", "-m", "3", "--since", "30d", "-a", "a@b.c"]),
            Some(Args {
                query: "plan".into(),
                notebook: Some("Work".into()),
                section: Some("Meet=ings".into()),
                account: Some("a@b.c".into()),
                since: Some("30d".into()),
                limit: Some(20),
                offset: Some(0),
                max_lines: Some(3),
                group: true,
                compact: true,
            })
        );
        assert_eq!(args(&["-g", "plan"]).map(|a| (a.query, a.group)), Some(("plan".into(), true)));
        assert_eq!(args(&["--notebook=-x", "plan"]).and_then(|a| a.notebook), Some("-x".into()));
    }

    #[test]
    fn everything_else_is_left_to_yargs() {
        for line in [
            &[][..],
            &["--group"],
            &["plan", "extra"],
            &["plan", "--online"],
            &["plan", "-o"],
            &["plan", "--help"],
            &["plan", "-h"],
            &["plan", "--version"],
            &["plan", "--unknown"],
            &["plan", "-x"],
            &["plan", "-gc"],
            &["plan", "-l5"],
            &["plan", "-l", "5", "--limit", "6"],
            &["plan", "-g", "--group"],
            &["plan", "--limit"],
            &["plan", "--limit", "-1"],
            &["plan", "--limit", "1.5"],
            &["plan", "--limit", "05"],
            &["plan", "--limit", "1e2"],
            &["plan", "--limit", "abc"],
            &["plan", "--limit", "1000000000"],
            &["plan", "--limit="],
            &["plan", "--notebook", ""],
            &["plan", "--notebook", "-x"],
            &["plan", "--no-group"],
            &["plan", "--group=true"],
            &["-g", "true"],
            &["plan", "-c", "false"],
            &["plan", "--maxLines", "3"],
            &["plan", "--max_lines", "3"],
            &["plan", "--"],
            &["--", "plan"],
            &["plan", "-"],
            &[""],
            &["-n=Work", "plan"],
            &["plan", "--notebook-id", "x"],
        ] {
            assert_eq!(args(line), None, "{line:?}");
        }
    }

    #[test]
    fn one_quote_comes_off_each_end() {
        assert_eq!(strip_quotes("\"plan\""), "plan");
        assert_eq!(strip_quotes("'plan"), "plan");
        assert_eq!(strip_quotes("plan\""), "plan");
        assert_eq!(strip_quotes("\"\"plan\"\""), "\"plan\"");
        assert_eq!(strip_quotes("\""), "");
        assert_eq!(strip_quotes("\"\""), "");
        assert_eq!(strip_quotes("pl\"an"), "pl\"an");
    }

    fn cut(raw: &str, query: &str) -> String {
        text(&excerpt(&units(raw), &units(query)).unwrap())
    }

    // The parity cases compare whole searches; this pins the piece.
    #[test]
    fn the_excerpt_is_the_line_with_the_query() {
        assert_eq!(cut("one\nthe Plan is here\nthree", "plan"), "the Plan is here");
        assert_eq!(cut("nothing here", "plan"), "");
        assert_eq!(cut("<b>bold</b> plan &amp; <i\nx>more &lt;", "plan"), "bold plan & more <");
    }

    #[test]
    fn titles_that_are_not_pages_are_left_out() {
        assert!(shown(&units("Weekly sync")));
        assert!(!shown(&units(".jpg")));
        assert!(!shown(&units("  .PDF ")));
        assert!(shown(&units(".profile")));
        assert!(!shown(&units("ab")));
        assert!(!shown(&units("")));
        assert!(shown(&units("会議メモ")));
        assert!(!shown(&units("\u{0600}\u{0601}\u{0602}\u{0603}\u{0604}\u{0605}abc")));
    }
}
