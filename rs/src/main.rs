//! onenote-cli — Rust entry point.
//!
//! Installed from npm, this binary is the command users run. A command runs
//! here once it has been ported and a parity test shows it behaves like the
//! TypeScript; until then it is handed back to the JavaScript the package
//! still ships, so nothing changes for users.

mod js_fallback;

/// The package's JavaScript entry, relative to the package root: the file
/// `bin` names in the published package (`npm view onenote-cli bin`), which
/// bun runs as it is.
const JS_ENTRY: &str = "src/index.ts";

/// `ONENOTE_CLI_NO_NATIVE=1` hands every command back to the JavaScript: the way
/// out for a user if a ported command misbehaves.
#[allow(dead_code)]
fn native() -> bool {
    std::env::var_os("ONENOTE_CLI_NO_NATIVE").is_none()
}

fn main() {
    // Arguments that are not valid UTF-8 are for the JavaScript to deal with.
    let args: Option<Vec<String>> = std::env::args_os().skip(1).map(|arg| arg.into_string().ok()).collect();
    match args.as_deref().and_then(|args| args.first()).map(String::as_str) {
        // Ported commands go here, one arm each, once their parity cases pass.
        // An arm takes only the invocations it implements exactly and hands
        // back the rest, before it has printed or changed anything.
        _ => js_fallback::run(JS_ENTRY),
    }
}
