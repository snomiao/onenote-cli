//! What happens before the TypeScript's first line, and whether any of it
//! could change what a ported command prints.
//!
//! "The TypeScript" is `bun src/index.ts`, and bun does things first:
//!
//! - it loads `.env`, `.env.local`, `.env.development` and
//!   `.env.development.local` from the current directory into the
//!   environment (other names under another NODE_ENV), in a syntax of its
//!   own;
//! - it reads `bunfig.toml` from the directory and from the user's
//!   configuration, which can preload scripts;
//! - it takes options from BUN_OPTIONS and NODE_OPTIONS, and it and the
//!   libraries inside it read many more variables: NODE_DEBUG makes it print
//!   a warning, BUN_JSC_* and MIMALLOC_VERBOSE make the engine and the
//!   allocator report on standard error, BUN_RUNTIME_TRANSPILER_CACHE_PATH
//!   makes it leave files in a directory.
//!
//! Then src/auth.ts, which every command imports, reads `.env.local` and
//! `.env` from the package directory into `process.env` (`autoLoadEnv`).
//!
//! The ported commands read one variable that these files could set: `TZ`,
//! which decides the local time printed as "last synced". (ONENOTE_CACHE_DIR
//! and the credentials are not read on any path that stays in Rust.) So the
//! files are not interpreted here. A file is left alone when the letters
//! `TZ`, `BUN`, `NODE` and `ICU` appear nowhere in it, in any case: then it
//! defines no variable named so, however its lines are read. Anything else
//! is `HandBack`, and the JavaScript runs as before.
//!
//! The environment itself is not interpreted either: any variable named for
//! the runtime (see `for_the_runtime`) is `HandBack`.

use std::ffi::OsStr;
use std::fs;
use std::io::ErrorKind;
use std::path::{Path, PathBuf};

use crate::HandBack;

/// Every file bun may load from the current directory, under any NODE_ENV.
const BUN_DOTENV_FILES: [&str; 8] = [
    ".env",
    ".env.local",
    ".env.development",
    ".env.development.local",
    ".env.production",
    ".env.production.local",
    ".env.test",
    ".env.test.local",
];

/// The files `autoLoadEnv` reads from the package directory.
const PACKAGE_DOTENV_FILES: [&str; 2] = [".env.local", ".env"];

/// Letters that a variable changing the runtime or the local time has in its
/// name: TZ, BUN_*, NODE_*, ICU_* (where ICU, which tells the runtime the
/// time, looks for its data). MIMALLOC_* is not among them: the allocator
/// has read the environment long before bun reads a .env file (a case).
const NAMES_THAT_MATTER: [&[u8]; 4] = [b"TZ", b"BUN", b"NODE", b"ICU"];

fn mentions(content: &[u8], name: &[u8]) -> bool {
    content.windows(name.len()).any(|window| window.eq_ignore_ascii_case(name))
}

/// A file that cannot define any of the variables that matter: it is not
/// there, or none of their names is anywhere in it.
fn harmless(file: &Path) -> Result<(), HandBack> {
    match fs::metadata(file) {
        Err(err) if err.kind() == ErrorKind::NotFound => Ok(()),
        Ok(meta) if meta.is_file() => {
            let content = fs::read(file).map_err(|_| HandBack)?;
            if NAMES_THAT_MATTER.iter().any(|name| mentions(&content, name)) {
                Err(HandBack)
            } else {
                Ok(())
            }
        }
        // A directory of that name, a file that cannot be read.
        _ => Err(HandBack),
    }
}

fn absent(file: &Path) -> Result<(), HandBack> {
    match fs::symlink_metadata(file) {
        Err(err) if err.kind() == ErrorKind::NotFound => Ok(()),
        _ => Err(HandBack),
    }
}

/// How the variables begin that the runtime and the libraries inside it read
/// for themselves: bun and its JavaScript engine (BUN_ENV, BUN_OPTIONS,
/// BUN_JSC_*, BUN_RUNTIME_TRANSPILER_CACHE_PATH, …), what it takes over from
/// Node (NODE_ENV, NODE_OPTIONS, NODE_DEBUG, …), its allocator (MIMALLOC_*)
/// and ICU, which tells it the time (ICU_TIMEZONE_FILES_DIR, ICU_DATA).
const FOR_THE_RUNTIME: [&str; 4] = ["BUN_", "NODE_", "MIMALLOC_", "ICU_"];

/// A variable the runtime may act on before or beside the program: by name
/// only, whatever its value. Which .env files are loaded, extra options,
/// warnings and reports on standard error, files left in a cache: none of it
/// is imitated, and the names are too many to tell apart, so each of them
/// hands the invocation back.
///
/// BUN_INSTALL is the exception. bun's installer puts it into the shell's
/// profile, so nearly every user has it; it says where bun itself lives and
/// changes nothing about how a program runs (compared in rs/parity.json).
fn for_the_runtime(name: &OsStr) -> bool {
    let name = name.to_string_lossy().to_ascii_uppercase();
    name != "BUN_INSTALL" && FOR_THE_RUNTIME.iter().any(|start| name.starts_with(start))
}

/// `Ok` when the JavaScript would start with the environment this process
/// has, as far as the ported commands can tell.
pub fn plain(package_root: &Path) -> Result<(), HandBack> {
    if std::env::vars_os().any(|(name, _)| for_the_runtime(&name)) {
        return Err(HandBack);
    }
    let cwd = std::env::current_dir().map_err(|_| HandBack)?;
    // bunfig.toml: in the directory, in the home directory, in the
    // configuration directory.
    // Without a home directory in the environment bun asks the system for
    // one; not imitated.
    let homes: Vec<PathBuf> = ["HOME", "USERPROFILE"].iter().filter_map(std::env::var_os).filter(|home| !home.is_empty()).map(PathBuf::from).collect();
    if homes.is_empty() {
        return Err(HandBack);
    }
    absent(&cwd.join("bunfig.toml"))?;
    for home in homes {
        absent(&home.join(".bunfig.toml"))?;
    }
    if let Some(config) = std::env::var_os("XDG_CONFIG_HOME") {
        absent(&PathBuf::from(config).join(".bunfig.toml"))?;
    }
    for name in BUN_DOTENV_FILES {
        harmless(&cwd.join(name))?;
    }
    for name in PACKAGE_DOTENV_FILES {
        harmless(&package_root.join(name))?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_name_is_found_in_any_case_anywhere() {
        assert!(mentions(b"TZ=Asia/Tokyo\n", b"TZ"));
        assert!(mentions(b"export tz = x", b"TZ"));
        assert!(mentions(b"A=1\nQUARTZ=2\n", b"TZ"));
        assert!(mentions(b"node_env=production", b"NODE"));
        assert!(mentions(b"X=${BUN_INSTALL}", b"BUN"));
        assert!(mentions(b"ICU_TIMEZONE_FILES_DIR=/x", b"ICU"));
        assert!(!mentions(b"ONENOTE_CLIENT_ID=abc\nONENOTE_CACHE_DIR=/x\n", b"TZ"));
        assert!(!mentions(b"", b"TZ"));
        assert!(!mentions(b"T", b"TZ"));
    }

    #[test]
    fn a_variable_named_for_the_runtime_is_told_by_how_it_begins() {
        for name in ["NODE_ENV", "BUN_ENV", "BUN_OPTIONS", "NODE_OPTIONS", "NODE_DEBUG", "BUN_JSC_dumpOptions", "bun_jsc_logGC", "MIMALLOC_VERBOSE", "NODE_NO_WARNINGS", "BUN_INSTALL_CACHE_DIR", "ICU_DATA", "BUN_"] {
            assert!(for_the_runtime(OsStr::new(name)), "{name}");
        }
        for name in ["BUN_INSTALL", "PATH", "HOME", "TZ", "NODE", "BUN", "BUNDLE_PATH", "NODENV_ROOT", "ONENOTE_CLIENT_ID", "ONENOTE_CLI_NO_NATIVE", "npm_config_user_agent", "LANG", "DEBUG", ""] {
            assert!(!for_the_runtime(OsStr::new(name)), "{name}");
        }
    }
}
