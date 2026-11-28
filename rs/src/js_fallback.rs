//! rustify JS fallback v6 — copy into the crate as `src/js_fallback.rs`.
//! (v2: waits for the JavaScript on Ctrl+C on Windows. v3: starts the
//! JavaScript with the runtime its own `#!` line names, and offers a
//! TypeScript entry to bun only. v4: says so when bun is what is missing.
//! v5: says on stderr that it handed back when `RUSTIFY_HANDBACK_TRACE` is
//! set; `attempt` runs a ported command so that it can decline, and so that a
//! bug in it costs the user nothing. v6: on Linux and macOS an entry with a
//! `#!` line is started as the system starts it, and only so; on Windows the
//! runtimes are looked up on PATH with PATHEXT, `bun.cmd` included, and
//! arguments a batch file cannot take are said to be so; `ROOT_ENV` is taken
//! out of the JavaScript's environment when it names this package. Also v6:
//! on Windows the runtimes are those the package managers' shims start: the
//! one the `#!` line names and no other, but bun after node, so an entry
//! whose line names bun never goes to node; a missing runtime is named, with
//! exit code 1. And on Windows a runtime found as an .exe is started by the
//! name the `#!` line gives, as the shims start it, where Windows finds the
//! same file by that name, so that it sees `bun` as its argv[0], not the
//! path.)
//!
//! For the time when the Rust binary is already the command users run but some
//! commands still exist only in the package's JavaScript. The binary handles
//! what it has ported and hands everything else to the old entry point:
//!
//! ```ignore
//! match command {
//!     "send" | "search" => run_natively(),
//!     _ => js_fallback::run("dist/cli.js"), // never returns
//! }
//! ```
//!
//! Ported commands start at native speed; the others behave exactly as before.
//!
//! A ported command seldom implements every invocation. `attempt` is for the
//! usual shape: the command works out its whole answer before it prints or
//! changes anything, and says `Err(HandBack)` for a command line or a state of
//! the machine it does not implement:
//!
//! ```ignore
//! "status" => {
//!     if let Some(text) = js_fallback::attempt(|| status::answer(&args[1..])) {
//!         // As console.log under bun: a reader that has gone away (`| head`)
//!         // is not an error. See "Printing" below for process.stdout.write.
//!         let _ = std::io::Write::write_all(&mut std::io::stdout(), text.as_bytes());
//!         return;
//!     }
//!     js_fallback::run("dist/cli.js")
//! }
//! ```
//!
//! Printing. bun treats a reader that has gone away (`| head`) in two ways
//! (measured with bun 1.4.0). `console.log` and `console.error` drop what they
//! cannot write and the program goes on, to exit 0 as usual: write with
//! `write_all` and ignore its error. `process.stdout.write` goes on too, to
//! the end of the code that runs at once, and then bun reports an uncaught
//! `EPIPE: broken pipe, write` with a stack trace on stderr and exits 1: for
//! such output, exit 1 when the write fails with `ErrorKind::BrokenPipe`. The
//! report itself names bun's internals and the installed file, so a port does
//! not copy it; that difference is for the owner to accept. Never `print!`,
//! which panics there.
//!
//! Who answered can be seen from outside: with `RUSTIFY_HANDBACK_TRACE=1` in
//! the environment, `run` first writes `rustify: handed back to the
//! JavaScript` to stderr, a panic that `attempt` swallowed is reported before
//! that line, and a native answer looks as it always does.
//!
//! Delete this file when nothing falls back any more, and take the `[native]`
//! and `[handback]` markers off the project's smoke lines with it: the install
//! smoke test finds out whether a binary can say who answered by looking for
//! that line's text in it, and fails a marked line when the text is missing.
use std::cell::Cell;
use std::ffi::OsString;
use std::io::Write;
#[cfg(unix)]
use std::os::unix::ffi::OsStringExt;
use std::path::PathBuf;
use std::process::Command;

/// Environment variable the npm launcher sets to the package directory when
/// the binary runs from the platform package. Use the same name as
/// `rustify.rootEnv` in package.json, and make it specific to the project
/// (`MYCLI_PACKAGE_ROOT`) so one tool starting another does not inherit it.
/// The JavaScript is started without it when it names this package, which is
/// what the launcher sets it to: the JavaScript then sees the environment it
/// had before this binary existed. A value of the user's own that names
/// another directory stays.
pub const ROOT_ENV: &str = "ONENOTE_CLI_PACKAGE_ROOT";

/// Set this variable to any value that is not empty and `run` writes
/// `TRACE_LINE` to stderr before it starts the JavaScript. That shows who
/// answered a command where no parity test runs: rustify's install smoke test
/// sets it for a line marked `[native]` or `[handback]`, and it finds out
/// whether a binary is able to say the line by looking for the text in the
/// file. So the name and the line are the same in every project: do not
/// change them. The JavaScript is started without the variable, so that a
/// tool it starts in turn cannot write the line in this binary's name.
pub const TRACE_ENV: &str = "RUSTIFY_HANDBACK_TRACE";
pub const TRACE_LINE: &str = "rustify: handed back to the JavaScript";

fn tracing() -> bool {
    std::env::var_os(TRACE_ENV).is_some_and(|value| !value.is_empty())
}

/// Directory of the npm package this binary was installed with, if any.
pub fn package_root() -> Option<PathBuf> {
    // Installed and linked: the binary is <package>/rustify/<command>, and the
    // packer puts a package.json marker next to it.
    if let Ok(exe) = std::env::current_exe() {
        // macOS reports the path the command was started through, and for an
        // npm install that is a symlink in the package manager's bin directory.
        #[cfg(unix)]
        let exe = std::fs::canonicalize(&exe).unwrap_or(exe);
        if let Some(bin) = exe.parent() {
            if bin.join("package.json").is_file() {
                if let Some(root) = bin.parent() {
                    if root.join("package.json").is_file() {
                        return Some(root.to_path_buf());
                    }
                }
            }
        }
    }
    // Running from the platform package: the launcher told us.
    let root = PathBuf::from(std::env::var_os(ROOT_ENV)?);
    root.join("package.json").is_file().then_some(root)
}

/// Runs the package's JavaScript entry with this process's arguments and ends
/// the process with its exit code. `entry` is relative to the package root.
///
/// The entry is started as it was before this binary existed. On Linux and
/// macOS, an entry with a `#!` line is started as the system starts it: with
/// the interpreter that line names and its argument, read as that system
/// reads them (see `shebang`), and with nothing else. `#!/usr/bin/env bun`
/// with no bun on PATH is then reported by env, in its words and with its
/// exit code (127), as when the command was the script itself. An entry
/// without such a line goes to bun, and to node when bun cannot be started
/// and the entry is not TypeScript. On Windows, where nothing but the package
/// managers' command shims reads the `#!` line, the entry is started as they
/// start it (see `runtimes`): with the runtime that line names, and with bun
/// where that is node and node is missing. Each is looked up in the
/// directories of PATH with the extensions of PATHEXT that a program can be
/// started from (see `find_program`; unlike a command prompt, not in the
/// working directory first): that finds `bun.exe`, and the `bun.cmd` that
/// `npm i -g bun` writes. When none can be started, the binary says which
/// and exits 1, as npm's shim does from a command prompt (see
/// `missing_runtime`).
pub fn run(entry: &str) -> ! {
    if tracing() {
        // A closed stderr must not stop the hand-back.
        let _ = writeln!(std::io::stderr(), "{TRACE_LINE}");
    }
    let Some(root) = package_root() else {
        eprintln!(
            "this command is not available in the standalone binary yet; \
             install the npm package to use it"
        );
        std::process::exit(1);
    };
    // In the separators of the system, as the package managers' shims name it.
    let script = root.join(if cfg!(windows) { entry.replace('/', "\\") } else { entry.to_string() });
    let args: Vec<OsString> = std::env::args_os().skip(1).collect();
    // The launcher's, not the user's: see ROOT_ENV.
    let hide_root = names_root(&root);
    let command = |program: &std::ffi::OsStr, words: &[OsString]| {
        let mut command = Command::new(program);
        command.args(words).arg(&script).args(&args).env_remove(TRACE_ENV);
        if hide_root {
            command.env_remove(ROOT_ENV);
        }
        command
    };
    #[cfg(unix)]
    {
        use std::os::unix::process::CommandExt;
        let line = shebang(&read_head(&script), cfg!(target_os = "macos"));
        // Never empty (see Shebang::Command); an empty one would go to bun below.
        if let Some((name, words)) = match &line { Shebang::Command(words) => words.split_first(), _ => None } {
            // As the system does it: argv[0] is the interpreter as the line
            // writes it, and one without a slash is found in the working
            // directory, not on PATH.
            let written = OsString::from_vec(name.clone());
            let program = if name.contains(&b'/') { written.clone() } else { OsString::from_vec([b"./".as_slice(), name].concat()) };
            let rest: Vec<OsString> = words.iter().cloned().map(OsString::from_vec).collect();
            // Replaces this process, so signals and the exit code need no relaying.
            let error = command(&program, &rest).arg0(&written).exec();
            let code = if error.kind() == std::io::ErrorKind::NotFound { 127 } else { 126 };
            eprintln!("cannot run {}: its #! line names {}, which could not be started ({error})", script.display(), written.to_string_lossy());
            std::process::exit(code);
        }
        let candidates = fallbacks(&script);
        let mut last_error = None;
        for name in candidates {
            last_error = Some(command(name.as_ref(), &[]).exec());
        }
        not_started(&script, candidates, last_error);
    }
    #[cfg(windows)]
    {
        stay_alive_on_ctrl_c();
        let path = std::env::var_os("PATH").unwrap_or_default();
        let pathext = std::env::var("PATHEXT").unwrap_or_else(|_| DEFAULT_PATHEXT.to_string());
        let line = interpreter(&first_line(&read_head(&script)));
        let named = line.is_some();
        let candidates = runtimes(&script, line);
        // For each runtime tried: None when it is not on PATH.
        let mut outcomes: Vec<(&str, Option<std::io::Error>)> = Vec::new();
        for runtime in &candidates {
            let Some(program) = find_program(&runtime[0], std::env::split_paths(&path), &pathext, |file| file.is_file()) else {
                outcomes.push((&runtime[0], None));
                continue;
            };
            let words: Vec<OsString> = runtime[1..].iter().map(OsString::from).collect();
            let start_as = start_name(&runtime[0], &program, &windows_first_dirs(), |file| file.is_file());
            match command(&start_as, &words).status() {
                Ok(status) => std::process::exit(status.code().unwrap_or(1)),
                // Rust refuses to start a batch file (bun.cmd) with an
                // argument that cmd.exe would not hand over unchanged.
                Err(error) if error.kind() == std::io::ErrorKind::InvalidInput && is_batch_file(&program) => {
                    eprintln!(
                        "cannot run {}: {} is a batch file, and Windows cannot hand it an argument with a line break \
                         unchanged ({error}); a runtime installed as an .exe, ahead of it on PATH, takes such arguments",
                        script.display(),
                        program.display()
                    );
                    std::process::exit(127);
                }
                Err(error) => outcomes.push((&runtime[0], Some(error))),
            }
        }
        eprintln!("{}", missing_runtime(&script, named, &outcomes));
        std::process::exit(1);
    }
    #[cfg(not(any(unix, windows)))]
    {
        let candidates = fallbacks(&script);
        let mut last_error = None;
        for name in candidates {
            match command(name.as_ref(), &[]).status() {
                Ok(status) => std::process::exit(status.code().unwrap_or(1)),
                Err(error) => last_error = Some(error),
            }
        }
        not_started(&script, candidates, last_error);
    }
}

/// Says that none of `tried` could be started, and ends the process (see
/// `missing_runtime` for Windows).
#[cfg_attr(windows, allow(dead_code))]
fn not_started(script: &std::path::Path, tried: &[&str], error: Option<std::io::Error>) -> ! {
    let reason = error.map(|e| e.to_string()).unwrap_or_default();
    if tried == ["bun"] {
        // A TypeScript entry: node is not offered, it cannot be relied on to run it.
        eprintln!("cannot run {}: it needs bun, which could not be started ({reason}). See https://bun.sh", script.display());
    } else {
        eprintln!("cannot run {}: none of {} could be started ({reason})", script.display(), tried.join(", "));
    }
    std::process::exit(127);
}

/// Whether `ROOT_ENV` holds the directory `root`, as the npm launcher sets it
/// (and as a test that points the binary at a package does). Only then is it
/// the launcher's, to be taken out of the JavaScript's environment.
pub fn names_root(root: &std::path::Path) -> bool {
    let Some(value) = std::env::var_os(ROOT_ENV) else {
        return false;
    };
    let named = PathBuf::from(value);
    if named.as_os_str().is_empty() {
        return false;
    }
    match (std::fs::canonicalize(&named), std::fs::canonicalize(root)) {
        (Ok(a), Ok(b)) => a == b,
        _ => named == root,
    }
}
/// What a ported command returns for an invocation, or a state of the
/// machine, that it does not implement exactly: the JavaScript has to answer.
#[allow(dead_code)] // unused until the first command is ported
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct HandBack;

thread_local! {
    /// Set while an attempt runs on this thread.
    static IN_ATTEMPT: Cell<bool> = const { Cell::new(false) };
    /// Set by `commit`, for the attempt running on this thread.
    static COMMITTED: Cell<bool> = const { Cell::new(false) };
}

// Both are asked by the panic hook, which can run while the thread is ending.
fn in_attempt() -> bool {
    IN_ATTEMPT.try_with(Cell::get).unwrap_or(false)
}

fn committed() -> bool {
    COMMITTED.try_with(Cell::get).unwrap_or(true)
}

/// Says that the command is about to print or change something, so the
/// invocation can no longer be handed back. Only a command that prints or
/// writes inside `attempt` needs this (one that reports progress while it
/// works, say); a command that returns its whole answer does not. It applies
/// to the attempt running on the calling thread, and the next attempt starts
/// afresh.
#[allow(dead_code)]
pub fn commit() {
    COMMITTED.with(|flag| flag.set(true));
}

/// Runs a ported command up to its answer. `Some(answer)` is for the caller
/// to print and carry out. `None` means the invocation goes to the
/// JavaScript (call `run` next): the command declined with `Err(HandBack)`,
/// or it panicked.
///
/// A panic is a bug in the port. As long as the command has printed and
/// changed nothing, the JavaScript can still answer in its place, so the
/// panic is swallowed without a word and the user gets what they asked for.
/// With `TRACE_ENV` set, the panic is reported as usual before the hand-back
/// line, which is how to find such a bug.
///
/// After `commit` there is no way back. A panic is then reported and ends the
/// process as any panic does (exit code 101), and returning `Err(HandBack)`
/// is itself a bug and panics.
///
/// The panic hook (what prints "thread 'main' panicked at …") belongs to the
/// whole process. While attempts run, a quieter one is in its place, and the
/// hook that was there before is put back when the last of them returns. In
/// the program, the quiet hook silences a panic on any thread, so a worker
/// thread the command started says nothing either. In the crate's own tests
/// (`cfg(test)`) it silences only the threads that run an attempt, so that a
/// test failing beside them keeps its message; there a worker thread started
/// inside an attempt is reported. Either way, a panic is reported on the
/// thread of an attempt that has committed, and everywhere when the trace is
/// on. Attempts may run on several threads at once and inside one another.
/// A test that puts a panic hook of its own in place should not run beside
/// tests that call `attempt`.
///
/// This relies on unwinding. A crate built with `panic = "abort"` cannot
/// catch anything: there a panic ends the process at once, with its usual
/// message, and the JavaScript is not started. The Cargo.toml that `rustify
/// scaffold` writes leaves the default, which unwinds.
#[allow(dead_code)]
pub fn attempt<T>(command: impl FnOnce() -> Result<T, HandBack>) -> Option<T> {
    use std::panic;
    // An attempt inside another starts afresh as well. When it returns, the
    // outer one is as it was, except that what the inner one committed stays
    // committed: it may have printed.
    let outer = (IN_ATTEMPT.replace(true), COMMITTED.replace(false));
    let outcome = if cfg!(panic = "unwind") {
        quiet_panics(true);
        let outcome = panic::catch_unwind(panic::AssertUnwindSafe(command));
        quiet_panics(false);
        outcome
    } else {
        Ok(command())
    };
    let committed = committed();
    IN_ATTEMPT.set(outer.0);
    COMMITTED.set(outer.1 || committed);
    match outcome {
        Ok(Ok(answer)) => Some(answer),
        Ok(Err(HandBack)) if committed => panic!("a ported command handed back after js_fallback::commit()"),
        Err(payload) if committed => panic::resume_unwind(payload),
        Ok(Err(HandBack)) | Err(_) => None,
    }
}

/// What `std::panic::take_hook` returns (`PanicHookInfo` needs Rust 1.81).
type PanicHook = Box<dyn Fn(&std::panic::PanicHookInfo<'_>) + Sync + Send + 'static>;

/// How many attempts are running, and the hook that was in place before the
/// first of them.
static QUIET: std::sync::Mutex<(usize, Option<std::sync::Arc<PanicHook>>)> = std::sync::Mutex::new((0, None));

/// Puts the quiet panic hook in place for an attempt (`true`), or takes it
/// away again when the last attempt is over (`false`).
fn quiet_panics(on: bool) {
    use std::panic;
    use std::sync::Arc;
    let mut state = QUIET.lock().unwrap_or_else(|poisoned| poisoned.into_inner());
    let (running, previous) = &mut *state;
    if on {
        if *running == 0 {
            let earlier: Arc<PanicHook> = Arc::new(panic::take_hook());
            let report = Arc::clone(&earlier);
            panic::set_hook(Box::new(move |info| {
                if committed() || tracing() || (cfg!(test) && !in_attempt()) {
                    report(info);
                }
            }));
            *previous = Some(earlier);
        }
        *running += 1;
        return;
    }
    *running -= 1;
    if *running == 0 {
        if let Some(earlier) = previous.take() {
            // Drops the quiet hook, and with it its share of the earlier one.
            drop(panic::take_hook());
            panic::set_hook(match Arc::try_unwrap(earlier) {
                Ok(hook) => hook,
                // The command kept the quiet hook somewhere, and with it a
                // share of the earlier one.
                Err(shared) => Box::new(move |info| shared(info)),
            });
        }
    }
}

/// The runtimes for an entry without a `#!` line: bun, and node too unless
/// the entry is TypeScript, which node cannot be relied on to run.
fn fallbacks(script: &std::path::Path) -> &'static [&'static str] {
    let is_typescript = matches!(script.extension().and_then(|e| e.to_str()), Some("ts" | "mts" | "cts" | "tsx"));
    if is_typescript {
        &["bun"]
    } else {
        &["bun", "node"]
    }
}

/// The most of a file that Linux (256 bytes, since Linux 5.1) or macOS (512)
/// reads for its `#!` line, and that the Windows shims are read for here.
const SHEBANG_MAX: usize = 512;

/// The first bytes of `script`: as many as a system reads for a `#!` line.
#[allow(dead_code)] // unused where neither Unix nor Windows
fn read_head(script: &std::path::Path) -> Vec<u8> {
    use std::io::Read;
    let mut head = Vec::new();
    if let Ok(file) = std::fs::File::open(script) {
        let _ = file.take(SHEBANG_MAX as u64).read_to_end(&mut head);
    }
    head
}

/// What a system makes of the first bytes of a script.
#[derive(Debug, PartialEq, Eq)]
#[allow(dead_code)] // unused on Windows, tested everywhere
enum Shebang {
    /// No `#!` at the very start of the file (behind a byte-order mark it is
    /// not one either).
    None,
    /// A `#!` line the system refuses to start the script with (ENOEXEC): one
    /// without an interpreter, on macOS one that does not end within 512
    /// bytes, on Linux one whose interpreter's name goes past 256. Taken as
    /// no `#!` line. So is one whose interpreter's name is empty because a
    /// NUL comes first (`#!\0…`, `#! \0…`), which the kernels try to start
    /// and cannot (EACCES on Linux, ENOENT on macOS).
    Refused,
    /// The interpreter, as the line writes it, and its arguments, which may
    /// be empty. Never empty itself: the interpreter's name is there.
    Command(Vec<Vec<u8>>),
}

/// The interpreter and its arguments that the system takes from a script's
/// first bytes, read as the kernel reads them: Linux (fs/binfmt_script.c)
/// passes everything after the interpreter as one argument, macOS (`split`;
/// XNU's exec_shell_imgact) splits it at blanks and also ends the line at a
/// `#`. Blanks are spaces and tabs only, so a carriage return stays part of
/// the line (`bun\r`), and a NUL ends what is read; a NUL right after a blank
/// leaves an empty argument, which both kernels pass on.
#[allow(dead_code)] // unused on Windows, tested everywhere
fn shebang(head: &[u8], split: bool) -> Shebang {
    let blank = |c: u8| c == b' ' || c == b'\t';
    if !head.starts_with(b"#!") {
        return Shebang::None;
    }
    if split {
        // XNU: the line ends at '\n' or '#', within the first 512 bytes.
        let mut vdata = [0u8; 512];
        let n = head.len().min(512);
        vdata[..n].copy_from_slice(&head[..n]);
        let eol = |c: u8| c == b'\n' || c == b'#';
        let Some(start) = (2..512).find(|&i| !blank(vdata[i])) else {
            return Shebang::Refused;
        };
        if eol(vdata[start]) {
            return Shebang::Refused;
        }
        let Some(mut end) = (start..512).find(|&i| eol(vdata[i])) else {
            return Shebang::Refused;
        };
        while eol(vdata[end]) || blank(vdata[end]) {
            end -= 1;
        }
        let line = &vdata[start..=end];
        // Read as a C string, and split as exec_extract_strings splits it: a
        // word ends at a blank, and the blanks after it are skipped, so a NUL
        // right after a blank leaves one empty word at the end.
        let mut rest = &line[..line.iter().position(|&c| c == 0).unwrap_or(line.len())];
        let mut words = Vec::new();
        loop {
            let at = rest.iter().position(|&c| blank(c)).unwrap_or(rest.len());
            words.push(rest[..at].to_vec());
            if at == rest.len() {
                break;
            }
            rest = &rest[at + 1..];
            rest = &rest[rest.iter().position(|&c| !blank(c)).unwrap_or(rest.len())..];
        }
        // A NUL first: the interpreter's name is empty (the kernel cannot
        // start that, ENOENT).
        if words[0].is_empty() {
            return Shebang::Refused;
        }
        return Shebang::Command(words);
    }
    // Linux: a buffer of 256 bytes, padded with NULs.
    let mut buf = [0u8; 256];
    let n = head.len().min(256);
    buf[..n].copy_from_slice(&head[..n]);
    let end = match buf.iter().position(|&c| c == b'\n') {
        Some(at) => at,
        None => {
            // No line end: the interpreter's name has to end within the
            // buffer (its last byte included), and the line is cut off
            // before that last byte.
            let Some(start) = (2..=255).find(|&i| !blank(buf[i])) else {
                return Shebang::Refused;
            };
            if !(start..=255).any(|i| blank(buf[i]) || buf[i] == 0) {
                return Shebang::Refused;
            }
            255
        }
    };
    let mut line = &buf[2..end];
    while let [rest @ .., last] = line {
        if !blank(*last) {
            break;
        }
        line = rest;
    }
    let line = &line[line.iter().position(|&c| !blank(c)).unwrap_or(line.len())..];
    let name_end = line.iter().position(|&c| blank(c) || c == 0).unwrap_or(line.len());
    // No name, or a NUL first (the kernel cannot start an empty name, EACCES).
    if name_end == 0 {
        return Shebang::Refused;
    }
    let mut words = vec![line[..name_end].to_vec()];
    if line.get(name_end).is_some_and(|&c| c != 0) {
        // The trailing blanks are gone, so something follows the blanks
        // here; up to a NUL, which can leave it empty, and the kernel passes
        // that empty argument too.
        let rest = &line[name_end..];
        let rest = &rest[rest.iter().position(|&c| !blank(c)).unwrap_or(rest.len())..];
        words.push(rest[..rest.iter().position(|&c| c == 0).unwrap_or(rest.len())].to_vec());
    }
    Shebang::Command(words)
}

/// The commands to try for running `script` on Windows, in order, as the
/// package managers' shims run it (one of them started the script before
/// this binary existed). `line` is what its `#!` line names (see
/// `interpreter`). Measured on Windows 11 with npm 11 and bun 1.4.2: both
/// start what the line names and nothing else, except that bun's shim starts
/// bun when the line names node and node is not on PATH. So a line naming
/// bun gets bun only, never node; one naming node gets node, then bun; any
/// other gets what it names. Without a `#!` line: bun, then node unless the
/// entry is TypeScript (see `fallbacks`).
#[allow(dead_code)] // used on Windows, tested everywhere
fn runtimes(script: &std::path::Path, line: Option<Vec<String>>) -> Vec<Vec<String>> {
    match line {
        Some(command) => {
            let node = command[0] == "node";
            let mut list = vec![command];
            if node {
                list.push(vec!["bun".to_string()]);
            }
            list
        }
        None => fallbacks(script).iter().map(|name| vec![name.to_string()]).collect(),
    }
}

/// The first line of a file's first bytes, as text.
#[allow(dead_code)] // used on Windows, tested everywhere
fn first_line(head: &[u8]) -> String {
    String::from_utf8_lossy(&head[..head.len().min(SHEBANG_MAX)]).lines().next().unwrap_or("").to_string()
}

/// What the binary says on Windows when no runtime for `script` could be
/// started: `named` when its `#!` line names the first of them, and for each
/// one tried, the error that kept it from starting, or None where it was not
/// on PATH. The shims say it in their own words: npm's `.cmd` lets cmd.exe
/// say `'"bun"' is not recognized as an internal or external command,` and
/// exits 1; bun's says `error: bun is not installed in %PATH%` and exits 255;
/// npm's `.ps1` reports a CommandNotFoundException and exits 0, its sh script
/// (Git Bash) says `exec: bun: not found` and exits 127. The binary names the
/// runtime and the entry, and exits 1, as the command prompt's shim does.
#[allow(dead_code)] // used on Windows, tested everywhere
fn missing_runtime(script: &std::path::Path, named: bool, outcomes: &[(&str, Option<std::io::Error>)]) -> String {
    let names: Vec<&str> = outcomes.iter().map(|(name, _)| *name).collect();
    let what = match (named, names.as_slice()) {
        (true, [first, ..]) => format!("its #! line names {first}; "),
        (false, ["bun"]) => "it needs bun; ".to_string(),
        _ => String::new(),
    };
    let reasons = if outcomes.iter().all(|(_, error)| error.is_none()) {
        match names.as_slice() {
            [one] => format!("{one} was not found in PATH"),
            [a, b] => format!("neither {a} nor {b} was found in PATH"),
            all => format!("none of {} was found in PATH", all.join(", ")),
        }
    } else {
        let each: Vec<String> = outcomes
            .iter()
            .map(|(name, error)| match error {
                None => format!("{name} was not found in PATH"),
                Some(error) => format!("{name} could not be started ({error})"),
            })
            .collect();
        each.join(", and ")
    };
    let hint = if names == ["bun"] { ". See https://bun.sh" } else { "" };
    format!("cannot run {}: {what}{reasons}{hint}", script.display())
}

/// The command a `#!` line asks for, without its directory, as the package
/// managers' shims read it: `#!/usr/bin/env node` is `["node"]`,
/// `#!/usr/bin/env -S npx tsx` is `["npx", "tsx"]`, and a line that is not a
/// `#!` line is `None`.
#[allow(dead_code)] // used on Windows, tested everywhere
fn interpreter(first_line: &str) -> Option<Vec<String>> {
    let mut words = first_line.strip_prefix("#!")?.split_whitespace();
    let name = |path: &str| path.rsplit(['/', '\\']).next().unwrap_or(path).to_string();
    let first = name(words.next()?);
    let command: Vec<String> = if first == "env" {
        // Skip env's own flags (-S) and VAR=value settings.
        words.skip_while(|w| w.starts_with('-') || w.contains('=')).map(String::from).collect()
    } else {
        std::iter::once(first).chain(words.map(String::from)).collect()
    };
    (!command.is_empty()).then_some(command)
}

/// What Windows takes for PATHEXT when it is not set.
#[allow(dead_code)] // used on Windows, tested everywhere
const DEFAULT_PATHEXT: &str = ".COM;.EXE;.BAT;.CMD;.VBS;.VBE;.JS;.JSE;.WSF;.WSH;.MSC";

/// The first `<dir>\<name><ext>` that is a file, for the directories of PATH
/// in order and, in each, the extensions of PATHEXT in order, as a command
/// prompt looks for a command, except that a command prompt looks in the
/// working directory first and this does not. Only the extensions a program
/// can be started from count: `.com`, `.exe`, and the `.bat` and `.cmd` files
/// that Rust starts through cmd.exe, with their arguments quoted for it.
/// They are tried in small letters, as the files are named (`bun.cmd`, not
/// PATHEXT's `bun.CMD`; Windows finds them either way). A name that already
/// ends in one of those is looked for as it is.
#[allow(dead_code)] // used on Windows, tested everywhere
fn find_program(name: &str, dirs: impl IntoIterator<Item = PathBuf>, pathext: &str, is_file: impl Fn(&std::path::Path) -> bool) -> Option<PathBuf> {
    const STARTABLE: [&str; 4] = [".com", ".exe", ".bat", ".cmd"];
    let ends_in = |ext: &str| name.len() > ext.len() && name.is_char_boundary(name.len() - ext.len()) && name[name.len() - ext.len()..].eq_ignore_ascii_case(ext);
    let extensions: Vec<String> = if STARTABLE.iter().any(|ext| ends_in(ext)) {
        vec![String::new()]
    } else {
        pathext.split(';').filter(|ext| STARTABLE.iter().any(|known| ext.eq_ignore_ascii_case(known))).map(str::to_ascii_lowercase).collect()
    };
    for dir in dirs {
        if dir.as_os_str().is_empty() {
            continue;
        }
        for ext in &extensions {
            let file = dir.join(format!("{name}{ext}"));
            if is_file(&file) {
                return Some(file);
            }
        }
    }
    None
}

/// What to start `program`, the file `find_program` found for `name`, as: the
/// runtime gets it as its argv[0], and bun and node print that (bun's hint
/// "Use `bun --trace-warnings ...` to show where the warning was created",
/// process.argv0). npm's command shim starts `bun` by that name, as the `#!`
/// line gives it, and so does this wherever Windows, looking the name up
/// itself, finds that same file. Rust looks a name without a path up in
/// `first` (the directory of this binary, then the system and Windows
/// directories: see `windows_first_dirs`), then in PATH, adding `.exe`;
/// `find_program` found no `name` in the directories of PATH before
/// `program`'s. Otherwise, and for a batch file, the path.
#[allow(dead_code)] // used on Windows, tested everywhere
fn start_name(name: &str, program: &std::path::Path, first: &[PathBuf], is_file: impl Fn(&std::path::Path) -> bool) -> OsString {
    let plain = !name.is_empty() && !name.contains(['/', '\\', '.', ':']);
    let exe = format!("{name}.exe");
    let full = program.to_string_lossy();
    // `<name>.exe`, in any letters, whichever separator the path uses.
    let is_name_exe = full.rsplit(['/', '\\']).next().is_some_and(|last| last.eq_ignore_ascii_case(&exe));
    let same = |file: &std::path::Path| file.to_string_lossy().replace('/', "\\").eq_ignore_ascii_case(&full.replace('/', "\\"));
    let elsewhere = first.iter().map(|dir| dir.join(&exe)).any(|file| is_file(&file) && !same(&file));
    if plain && is_name_exe && !elsewhere {
        OsString::from(name)
    } else {
        program.as_os_str().to_owned()
    }
}

/// The directories Rust looks a program up in before PATH on Windows: that of
/// this binary, the system directory and the Windows directory.
#[cfg(windows)]
fn windows_first_dirs() -> Vec<PathBuf> {
    use std::os::windows::ffi::OsStringExt;
    extern "system" {
        fn GetSystemDirectoryW(buffer: *mut u16, size: u32) -> u32;
        fn GetWindowsDirectoryW(buffer: *mut u16, size: u32) -> u32;
    }
    let mut dirs: Vec<PathBuf> = std::env::current_exe().ok().and_then(|exe| exe.parent().map(PathBuf::from)).into_iter().collect();
    for get in [GetSystemDirectoryW as unsafe extern "system" fn(*mut u16, u32) -> u32, GetWindowsDirectoryW] {
        let mut buffer = [0u16; 1024];
        let n = unsafe { get(buffer.as_mut_ptr(), buffer.len() as u32) } as usize;
        if n > 0 && n < buffer.len() {
            dirs.push(PathBuf::from(OsString::from_wide(&buffer[..n])));
        }
    }
    dirs
}

/// Whether Rust starts `program` through cmd.exe, as it does a `.bat` or a
/// `.cmd` file.
#[allow(dead_code)] // used on Windows, tested everywhere
fn is_batch_file(program: &std::path::Path) -> bool {
    program.extension().and_then(|e| e.to_str()).is_some_and(|e| e.eq_ignore_ascii_case("bat") || e.eq_ignore_ascii_case("cmd"))
}
/// Ctrl+C reaches every process attached to the console. Without a handler
/// this process would end at once and hand the prompt back while the
/// JavaScript is still shutting down. With one it waits, and exits with the
/// child's exit code. A handler, unlike the "ignore Ctrl+C" flag, is not
/// inherited, so the child still gets the key press.
#[cfg(windows)]
fn stay_alive_on_ctrl_c() {
    extern "system" {
        fn SetConsoleCtrlHandler(
            handler: Option<unsafe extern "system" fn(u32) -> i32>,
            add: i32,
        ) -> i32;
    }
    unsafe extern "system" fn handled(_event: u32) -> i32 {
        1
    }
    unsafe {
        SetConsoleCtrlHandler(Some(handled), 1);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn names(script: &std::path::Path) -> Vec<String> {
        runtimes(script, interpreter(&first_line(&read_head(script)))).into_iter().map(|c| c.join(" ")).collect()
    }

    /// The `#!` line as Linux (one argument) and macOS (split at blanks, `#`
    /// ends the line) read it; tests/e2e/toy/tests/handback.rs checks the
    /// same against the kernel of the system it runs on.
    #[test]
    fn the_shebang_is_read_as_each_kernel_reads_it() {
        let words = |head: &[u8], split: bool| match shebang(head, split) {
            Shebang::Command(words) => Some(words.into_iter().map(|w| String::from_utf8(w).unwrap()).collect::<Vec<_>>()),
            other => {
                assert_eq!(other, Shebang::Refused, "{head:?}");
                None
            }
        };
        let both = |head: &str, linux: Option<&[&str]>, macos: Option<&[&str]>| {
            assert_eq!(words(head.as_bytes(), false).as_deref(), linux.map(|w| w.iter().map(|s| s.to_string()).collect::<Vec<_>>()).as_deref(), "Linux, {head:?}");
            assert_eq!(words(head.as_bytes(), true).as_deref(), macos.map(|w| w.iter().map(|s| s.to_string()).collect::<Vec<_>>()).as_deref(), "macOS, {head:?}");
        };
        both("#!/usr/bin/env bun\nimport x", Some(&["/usr/bin/env", "bun"]), Some(&["/usr/bin/env", "bun"]));
        both("#! \t/usr/bin/env  bun \t\n", Some(&["/usr/bin/env", "bun"]), Some(&["/usr/bin/env", "bun"]));
        both("#!/usr/bin/env -S bun --flag\n", Some(&["/usr/bin/env", "-S bun --flag"]), Some(&["/usr/bin/env", "-S", "bun", "--flag"]));
        // A carriage return is part of the line for both.
        both("#!/usr/bin/env bun\r\n", Some(&["/usr/bin/env", "bun\r"]), Some(&["/usr/bin/env", "bun\r"]));
        // macOS ends the line at a `#` as well.
        both("#!/bin/sh one # two\n", Some(&["/bin/sh", "one # two"]), Some(&["/bin/sh", "one"]));
        // A relative interpreter is as written; the system finds it from the working directory.
        both("#!node --flag\n", Some(&["node", "--flag"]), Some(&["node", "--flag"]));
        // A NUL ends what is read.
        both("#!/bin/sh one\0two\n", Some(&["/bin/sh", "one"]), Some(&["/bin/sh", "one"]));
        both("#!/bin/sh\0 one\n", Some(&["/bin/sh"]), Some(&["/bin/sh"]));
        // Right after a blank, it leaves an empty argument, which both kernels pass.
        both("#!/bin/sh \0two\n", Some(&["/bin/sh", ""]), Some(&["/bin/sh", ""]));
        both("#!/bin/sh\t\0two\n", Some(&["/bin/sh", ""]), Some(&["/bin/sh", ""]));
        both("#!/bin/sh one \0two\n", Some(&["/bin/sh", "one "]), Some(&["/bin/sh", "one", ""]));
        // First, it leaves no interpreter: refused, as the kernels cannot start one.
        both("#!\0/usr/bin/env node\n", None, None);
        both("#!  \0/usr/bin/env node\n", None, None);
        both("#!\t\0", None, None);
        // Linux reads 256 bytes and cuts a longer line off; macOS reads 512 and refuses a longer one.
        let long = format!("#!/usr/bin/env {}\n", "a".repeat(300));
        both(&long, Some(&["/usr/bin/env", &"a".repeat(240)]), Some(&["/usr/bin/env", &"a".repeat(300)]));
        let longer = format!("#!/usr/bin/env {}\n", "a".repeat(600));
        both(&longer, Some(&["/usr/bin/env", &"a".repeat(240)]), None);
        both(&format!("#!/{}\n", "x".repeat(300)), None, Some(&[&format!("/{}", "x".repeat(300))]));
        // Linux: a name that a blank at the buffer's last byte ends is taken, and cut off before that byte.
        let edge = format!("#!/{} y{}\n", "x".repeat(252), "z".repeat(50));
        both(&edge, Some(&[&format!("/{}", "x".repeat(252))]), Some(&[&format!("/{}", "x".repeat(252)), &format!("y{}", "z".repeat(50))]));
        // No line end at all: Linux takes the file as the line; macOS finds no end.
        both("#!/bin/sh one", Some(&["/bin/sh", "one"]), None);
        both("#!\n", None, None);
        both("#!  \t\n", None, None);
        both("#!", None, None);
        assert_eq!(shebang(b"\xef\xbb\xbf#!/bin/sh\n", false), Shebang::None, "a byte-order mark first");
        assert_eq!(shebang(b"\xef\xbb\xbf#!/bin/sh\n", true), Shebang::None);
        assert_eq!(shebang(b"console.log(1)\n", false), Shebang::None);
        assert_eq!(shebang(b"", true), Shebang::None);
    }

    /// bun on Windows, as a command prompt finds it.
    #[test]
    fn a_runtime_is_found_as_a_command_prompt_finds_it() {
        let files = ["c:/npm/bun.cmd", "c:/npm/bun.ps1", "c:/npm/bun", "c:/bun/bun.exe", "c:/js/bun.js", "c:/npm/npx.cmd", "c:/bat/bun.bat"];
        // Windows finds a file whatever the letter case of its name.
        let spelt = |file: &std::path::Path| file.to_string_lossy().replace('\\', "/").to_lowercase();
        let is_file = |file: &std::path::Path| files.iter().any(|f| *f == spelt(file));
        let find = |name: &str, dirs: &[&str], pathext: &str| find_program(name, dirs.iter().map(PathBuf::from), pathext, is_file).map(|f| spelt(&f));
        // The first directory of PATH that has it, whatever the extension.
        assert_eq!(find("bun", &["C:/npm", "C:/bun"], DEFAULT_PATHEXT).as_deref(), Some("c:/npm/bun.cmd"));
        assert_eq!(find("bun", &["C:/bun", "C:/npm"], DEFAULT_PATHEXT).as_deref(), Some("c:/bun/bun.exe"));
        assert_eq!(find("bun", &["", "C:/none", "C:/bun"], ".exe;.CMD").as_deref(), Some("c:/bun/bun.exe"));
        // In one directory, in the order of PATHEXT.
        assert_eq!(find("bun", &["C:/bat"], ".CMD;.BAT").as_deref(), Some("c:/bat/bun.bat"));
        // Only what a program can be started from: not a .ps1, a sh script or a .js.
        assert_eq!(find("bun", &["C:/js", "C:/npm"], ".JS;.PS1;.CMD").as_deref(), Some("c:/npm/bun.cmd"));
        assert_eq!(find("bun", &["C:/npm"], ".EXE"), None);
        assert_eq!(find("bun", &[], DEFAULT_PATHEXT), None);
        // npx, as a `#!/usr/bin/env -S npx tsx` line names it.
        assert_eq!(find("npx", &["C:/bun", "C:/npm"], DEFAULT_PATHEXT).as_deref(), Some("c:/npm/npx.cmd"));
        // A name with its extension is looked for as it is.
        assert_eq!(find("bun.exe", &["C:/npm", "C:/bun"], ".CMD").as_deref(), Some("c:/bun/bun.exe"));
        // The file as it is named, not in PATHEXT's capitals.
        let found = find_program("bun", [PathBuf::from("C:/npm")], ".COM;.EXE;.BAT;.CMD", is_file).unwrap();
        assert_eq!(found.file_name().and_then(|f| f.to_str()), Some("bun.cmd"));
        assert!(is_batch_file(std::path::Path::new("C:/npm/bun.CMD")) && is_batch_file(std::path::Path::new("x.bat")));
        assert!(!is_batch_file(std::path::Path::new("C:/bun/bun.exe")));
    }

    /// On Windows a runtime found as an .exe is started by its name, as the
    /// shims start it, where Windows looks that name up to the same file.
    #[test]
    fn a_runtime_is_started_by_its_name_where_that_finds_the_same_file() {
        let files = ["c:/bun/bun.exe", "c:/npm/bun.cmd", "c:/app/node.exe", "c:/node/node.exe"];
        let spelt = |file: &std::path::Path| file.to_string_lossy().replace('\\', "/").to_lowercase();
        let is_file = |file: &std::path::Path| files.iter().any(|f| *f == spelt(file));
        let first = [PathBuf::from("C:/app"), PathBuf::from("C:/Windows/System32"), PathBuf::from("C:/Windows")];
        let start = |name: &str, program: &str| start_name(name, std::path::Path::new(program), &first, is_file).to_string_lossy().into_owned();
        assert_eq!(start("bun", "C:/bun/bun.exe"), "bun");
        assert_eq!(start("bun", "C:\\BUN\\Bun.EXE"), "bun", "in other letters and separators it is the same file");
        // A batch file is started as a path (Rust runs it through cmd.exe).
        assert_eq!(start("bun", "C:/npm/bun.cmd"), "C:/npm/bun.cmd");
        // Windows would find another node.exe first, in this binary's directory.
        assert_eq!(start("node", "C:/node/node.exe"), "C:/node/node.exe");
        assert_eq!(start("node", "C:/app/node.exe"), "node");
        // A name with an extension or a path is started as the path found.
        assert_eq!(start("bun.exe", "C:/bun/bun.exe"), "C:/bun/bun.exe");
        assert_eq!(start("tools/bun", "C:/bun/bun.exe"), "C:/bun/bun.exe");
    }

    /// As the package managers' shims read the `#!` line, on Windows.
    #[test]
    fn the_shebang_names_the_interpreter() {
        assert_eq!(interpreter("#!/usr/bin/env node"), Some(vec!["node".to_string()]));
        assert_eq!(interpreter("#!/usr/bin/env bun"), Some(vec!["bun".to_string()]));
        assert_eq!(interpreter("#!/usr/local/bin/node --no-warnings").unwrap(), ["node", "--no-warnings"]);
        assert_eq!(interpreter("#!/usr/bin/env -S npx tsx").unwrap(), ["npx", "tsx"]);
        assert_eq!(interpreter("#!/usr/bin/env npx tsx").unwrap(), ["npx", "tsx"]);
        assert_eq!(interpreter("#! /usr/bin/env  FOO=1 node").unwrap(), ["node"]);
        assert_eq!(interpreter("import x from \"y\";"), None);
        assert_eq!(interpreter("#!/usr/bin/env"), None);
        assert_eq!(interpreter(""), None);
    }

    /// Whether the test `test` of this module is to run in this process. The
    /// panic hook belongs to the whole process, and the crate's other tests
    /// may call `attempt`, put a hook in place or panic while a test watches
    /// it. So such a test runs again, alone, in a process of its own, once
    /// with the trace off and once with it on: `alone` returns true there.
    /// Anywhere else it starts those two runs, fails unless the test passed
    /// in both, and returns false.
    fn alone(test: &str) -> bool {
        const ALONE: &str = "RUSTIFY_FALLBACK_TEST_ALONE";
        if std::env::var_os(ALONE).is_some() {
            return true;
        }
        // The test binary names a test by its path without the crate's name.
        let module = module_path!().split_once("::").map_or(module_path!(), |(_, path)| path);
        let name = format!("{module}::{test}");
        for trace in [false, true] {
            let mut command = Command::new(std::env::current_exe().unwrap());
            command.args([name.as_str(), "--exact", "--test-threads=1"]).env(ALONE, "1");
            if trace {
                command.env(TRACE_ENV, "1");
            } else {
                command.env_remove(TRACE_ENV);
            }
            let run = command.output().unwrap();
            let output = format!("{}{}", String::from_utf8_lossy(&run.stdout), String::from_utf8_lossy(&run.stderr));
            assert!(
                run.status.success() && output.contains("1 passed"),
                "{name}, run alone with the trace {}:\n{output}",
                if trace { "on" } else { "off" }
            );
        }
        false
    }

    /// One test for all of `attempt`, in this order.
    #[test]
    fn attempt_answers_declines_and_hands_back_a_panic_until_commit() {
        use std::panic::catch_unwind;
        use std::sync::atomic::{AtomicUsize, Ordering::SeqCst};
        if !alone("attempt_answers_declines_and_hands_back_a_panic_until_commit") {
            return;
        }
        // Counts the panics that reach the hook that was there before, on any
        // thread, and passes them on, so that a failing assertion shows.
        static REPORTED: AtomicUsize = AtomicUsize::new(0);
        let reported = || REPORTED.load(SeqCst);
        let before = std::panic::take_hook();
        std::panic::set_hook(Box::new(move |info| {
            REPORTED.fetch_add(1, SeqCst);
            before(info);
        }));
        // With the trace on, a swallowed panic is reported too.
        let traced = usize::from(tracing());
        let mut expected = 0;
        let bug = || -> Result<i32, HandBack> { panic!("a bug in a ported command") };

        assert_eq!(attempt(|| Ok(7)), Some(7));
        assert_eq!(attempt(|| Err::<i32, _>(HandBack)), None);
        assert_eq!(attempt(bug), None);
        expected += traced;
        assert_eq!(reported(), expected, "a panic before commit is not reported");
        assert!(catch_unwind(|| panic!("outside attempt")).is_err());
        expected += 1;
        assert_eq!(reported(), expected, "the hook that was there before is back");

        // Attempts on several threads at once leave that hook in place too.
        let threads: Vec<_> = (0..4)
            .map(|_| std::thread::spawn(move || (0..50).filter(|_| attempt(bug).is_none()).count()))
            .collect();
        assert!(threads.into_iter().all(|thread| thread.join().unwrap() == 50));
        expected += 200 * traced;
        assert!(catch_unwind(|| panic!("after the threads")).is_err());
        expected += 1;
        assert_eq!(reported(), expected, "the hook is back after attempts on several threads");

        // In a test build, a panic on a thread that runs no attempt is
        // reported while another thread runs one.
        let (started, running) = std::sync::mpsc::channel();
        let (finish, finished) = std::sync::mpsc::channel::<()>();
        let beside = std::thread::spawn(move || {
            attempt(|| {
                started.send(()).unwrap();
                finished.recv().unwrap();
                Ok(())
            })
        });
        running.recv().unwrap();
        assert!(catch_unwind(|| panic!("a test failing beside an attempt")).is_err());
        expected += 1;
        assert_eq!(reported(), expected, "a panic beside an attempt is reported");
        finish.send(()).unwrap();
        assert_eq!(beside.join().unwrap(), Some(()));

        let late = || -> Result<i32, HandBack> {
            commit();
            panic!("a bug after the first output")
        };
        assert!(catch_unwind(|| attempt(late)).is_err(), "a panic after commit is not swallowed");
        expected += 1;
        assert_eq!(reported(), expected, "and it is reported");
        let late_hand_back = || -> Result<i32, HandBack> {
            commit();
            Err(HandBack)
        };
        assert!(catch_unwind(|| attempt(late_hand_back)).is_err(), "handing back after commit is a bug");
        expected += 1;
        assert_eq!(reported(), expected);
        assert_eq!(attempt(|| Err::<i32, _>(HandBack)), None, "the next attempt starts afresh");

        // Attempts inside one another.
        let outer = || -> Result<i32, HandBack> {
            assert_eq!(attempt(|| Err::<i32, _>(HandBack)), None);
            Ok(1)
        };
        assert_eq!(attempt(outer), Some(1), "an inner attempt that declines leaves the outer one going");
        let outer_bug = || -> Result<i32, HandBack> {
            assert_eq!(attempt(|| Ok(2)), Some(2));
            panic!("a bug after an inner attempt")
        };
        assert_eq!(attempt(outer_bug), None);
        expected += traced;
        assert_eq!(reported(), expected, "after an inner attempt, the outer one is still quiet");
        let inner_commit = || -> Result<i32, HandBack> {
            commit();
            Ok(3)
        };
        let outer_after_commit = || -> Result<i32, HandBack> {
            assert_eq!(attempt(inner_commit), Some(3));
            Err(HandBack)
        };
        assert!(catch_unwind(|| attempt(outer_after_commit)).is_err(), "what an inner attempt committed stays committed");
        expected += 1;
        assert_eq!(reported(), expected);
    }

    /// On Windows, as the shims start the entry: what the `#!` line names and
    /// nothing else, but bun after node; without a line, bun, then node (not
    /// for TypeScript).
    #[test]
    fn the_script_decides_which_runtime_is_tried_first() {
        let dir = std::env::temp_dir().join(format!("rustify-fallback-{}", std::process::id()));
        std::fs::create_dir_all(&dir).unwrap();
        let file = |name: &str, text: &str| {
            let path = dir.join(name);
            std::fs::write(&path, text).unwrap();
            path
        };
        assert_eq!(names(&file("a.js", "#!/usr/bin/env node\nx")), ["node", "bun"]);
        // A line naming bun never goes to node, whatever the file is.
        assert_eq!(names(&file("b.js", "#!/usr/bin/env bun\nx")), ["bun"]);
        assert_eq!(names(&file("b.mjs", "#!/usr/bin/env bun\r\nimport x from \"./src/server\";\r\n")), ["bun"]);
        assert_eq!(names(&file("b2.js", "#!/usr/bin/env -S bun --smol\nx")), ["bun --smol"]);
        assert_eq!(names(&file("c.js", "console.log(1)")), ["bun", "node"]);
        assert_eq!(names(&file("d.ts", "#!/usr/bin/env bun\nx")), ["bun"]);
        assert_eq!(names(&file("e.ts", "#!/usr/bin/env npx tsx\nx")), ["npx tsx"]);
        assert_eq!(names(&file("e.js", "#!/usr/bin/env tsx\nx")), ["tsx"]);
        assert_eq!(names(&file("f.ts", "#!/usr/bin/env node\nx")), ["node", "bun"]);
        assert_eq!(names(&dir.join("missing.js")), ["bun", "node"]);
        // Without a `#!` line, everywhere: bun, then node unless it is TypeScript.
        assert_eq!(fallbacks(&dir.join("g.js")), ["bun", "node"]);
        assert_eq!(fallbacks(&dir.join("h.mts")), ["bun"]);
        std::fs::remove_dir_all(&dir).unwrap();
    }

    /// What a Windows user reads when no runtime could be started.
    #[test]
    fn a_missing_runtime_is_named() {
        let script = std::path::Path::new("C:\\pkg\\wtx.mjs");
        let missing = |names: &[&'static str]| names.iter().map(|n| (*n, None)).collect::<Vec<_>>();
        assert_eq!(missing_runtime(script, true, &missing(&["bun"])), "cannot run C:\\pkg\\wtx.mjs: its #! line names bun; bun was not found in PATH. See https://bun.sh");
        assert_eq!(missing_runtime(script, true, &missing(&["node", "bun"])), "cannot run C:\\pkg\\wtx.mjs: its #! line names node; neither node nor bun was found in PATH");
        assert_eq!(missing_runtime(script, true, &missing(&["npx"])), "cannot run C:\\pkg\\wtx.mjs: its #! line names npx; npx was not found in PATH");
        assert_eq!(missing_runtime(script, false, &missing(&["bun", "node"])), "cannot run C:\\pkg\\wtx.mjs: neither bun nor node was found in PATH");
        assert_eq!(missing_runtime(script, false, &missing(&["bun"])), "cannot run C:\\pkg\\wtx.mjs: it needs bun; bun was not found in PATH. See https://bun.sh");
        let failed = [("node", Some(std::io::Error::from(std::io::ErrorKind::PermissionDenied))), ("bun", None)];
        let text = missing_runtime(script, true, &failed);
        assert!(text.starts_with("cannot run C:\\pkg\\wtx.mjs: its #! line names node; node could not be started (") && text.ends_with("), and bun was not found in PATH"), "{text}");
    }
}
