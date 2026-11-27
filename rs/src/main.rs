//! onenote-cli — Rust entry point.
//!
//! Installed from npm, this binary is the command users run. A command runs
//! here once it has been ported and a parity test shows it behaves like the
//! TypeScript; until then it is handed back to the JavaScript the package
//! still ships, so nothing changes for users.

mod js_fallback;
mod runtime;

// The search, and what only it needs. Linux only: see `search` below.
#[cfg(target_os = "linux")]
mod js;
#[cfg(target_os = "linux")]
mod query;
#[cfg(target_os = "linux")]
mod search;
#[cfg(target_os = "linux")]
mod time;

use std::io::{IsTerminal, Write};

/// The package's JavaScript entry, relative to the package root: the file
/// `bin` names in the published package (`npm view onenote-cli bin`), which
/// bun runs as it is.
const JS_ENTRY: &str = "src/index.ts";

/// "Not mine": the invocation has to be handed to the JavaScript. Only ever
/// returned before the first output. Nothing has been changed by then
/// either, except that SQLite may have put its `-shm` and `-wal` files
/// beside an index it was asked to read, as it does for the TypeScript.
pub use js_fallback::HandBack;

/// `ONENOTE_CLI_NO_NATIVE=1` hands every command back to the JavaScript: the way
/// out for a user if a ported command misbehaves.
fn native() -> bool {
    std::env::var_os("ONENOTE_CLI_NO_NATIVE").is_none()
}

/// What a ported command needs to know about the machine.
#[cfg_attr(not(target_os = "linux"), allow(dead_code))]
struct Machine {
    /// The package directory: the JavaScript lives there, and so do the
    /// search index and the caches.
    package_root: std::path::PathBuf,
    /// `process.stdout.isTTY`: on a terminal the TypeScript colours its
    /// output and makes titles clickable.
    terminal: bool,
}

/// The machine, when it is in a state the ported commands were compared in:
/// installed as a package, and nothing set up that changes the JavaScript
/// runtime (see runtime.rs).
fn machine() -> Result<Machine, HandBack> {
    let terminal = std::io::stdout().is_terminal();
    // A Windows console was never compared, and what counts as a terminal
    // differs there: Rust says yes to a Git Bash window, the runtime no.
    if terminal && cfg!(windows) {
        return Err(HandBack);
    }
    let package_root = js_fallback::package_root().ok_or(HandBack)?;
    if !loadable(&package_root) {
        return Err(HandBack);
    }
    runtime::plain(&package_root)?;
    Ok(Machine { package_root, terminal })
}

/// Whether the runtime can load the program from that directory. bun 1.4.0
/// cannot when the path, or the path its links lead to, has a `?` or a `\`
/// in it (it takes the one for the start of a query and the other for a
/// `/`): the TypeScript then ends with `Module not found` before its first
/// line, and so must this.
#[cfg(unix)]
fn loadable(package_root: &std::path::Path) -> bool {
    let plain = |path: &std::path::Path| !path.as_os_str().as_encoded_bytes().iter().any(|byte| matches!(byte, b'?' | b'\\'));
    std::fs::canonicalize(package_root).is_ok_and(|real| plain(package_root) && plain(&real))
}

/// On Windows a `\` separates directories and a `?` cannot be in a name.
#[cfg(not(unix))]
fn loadable(_package_root: &std::path::Path) -> bool {
    true
}

/// `onenote search <query>`: what it prints, or `HandBack`.
///
/// On Linux only. The last line of a search carries a local time written by
/// `toLocaleString()`. On Linux bun's locale is en-US whatever the user's is
/// (compared under LANG and LC_ALL). On macOS and Windows a JavaScript
/// runtime usually takes the system's language and region, and what bun
/// prints there could not be tried; nor could bun's SQLite on those systems
/// (on macOS it is documented to be the system's, not its own). So there the
/// search stays the JavaScript's, and SQLite is not even compiled in.
#[cfg(target_os = "linux")]
fn search(args: &[String]) -> Result<String, HandBack> {
    let args = search::parse_args(args).ok_or(HandBack)?;
    let machine = machine()?;
    search::run(&args, &machine.package_root, machine.terminal)
}

#[cfg(not(target_os = "linux"))]
fn search(_args: &[String]) -> Result<String, HandBack> {
    Err(HandBack)
}

/// `onenote auth setup`: the guide, word for word (src/index.ts, the
/// `setup` command under `auth`).
fn auth_setup() -> Result<String, HandBack> {
    machine()?;
    Ok(include_str!("auth_setup.txt").to_string())
}

/// Runs a ported command. It works out everything it will print before
/// printing any of it, so `HandBack` (and a bug that makes it panic) leaves
/// nothing behind and the JavaScript can do the whole invocation
/// (`js_fallback::attempt`).
fn answer(command: impl FnOnce() -> Result<String, HandBack>) {
    if let Some(output) = js_fallback::attempt(command) {
        let mut stdout = std::io::stdout().lock();
        // A reader that has gone away (`| head`) is not an error.
        let _ = stdout.write_all(output.as_bytes()).and_then(|()| stdout.flush());
        std::process::exit(0);
    }
}

fn main() {
    // Arguments that are not valid UTF-8 are for the JavaScript to deal with.
    let args: Option<Vec<String>> = std::env::args_os().skip(1).map(|arg| arg.into_string().ok()).collect();
    if let (Some(args), true) = (&args, native()) {
        // Ported commands, one arm each, switched on once their parity cases
        // (rs/parity.json) pass. An arm takes only the invocations it
        // implements exactly; for the rest it returns and the JavaScript runs.
        match args.first().map(String::as_str) {
            Some("search") => answer(|| search(&args[1..])),
            // Exactly `auth setup`; with anything more it is yargs's.
            Some("auth") if args[1..] == ["setup"] => answer(auth_setup),
            _ => {}
        }
    }
    js_fallback::run(JS_ENTRY)
}
