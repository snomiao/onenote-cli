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
//! - it takes options from BUN_OPTIONS and NODE_OPTIONS.
//!
//! Then src/auth.ts, which every command imports, reads `.env.local` and
//! `.env` from the package directory into `process.env` (`autoLoadEnv`).
//!
//! The ported commands read one variable that these files could set: `TZ`,
//! which decides the local time printed as "last synced". (ONENOTE_CACHE_DIR
//! and the credentials are not read on any path that stays in Rust.) So the
//! files are not interpreted here. A file is left alone when the letters
//! `TZ`, `BUN` and `NODE` appear nowhere in it, in any case: then it defines
//! no variable named so, however its lines are read. Anything else is
//! `HandBack`, and the JavaScript runs as before.

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
/// name: TZ, BUN_*, NODE_*.
const NAMES_THAT_MATTER: [&[u8]; 3] = [b"TZ", b"BUN", b"NODE"];

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

/// `Ok` when the JavaScript would start with the environment this process
/// has, as far as the ported commands can tell.
pub fn plain(package_root: &Path) -> Result<(), HandBack> {
    // These change what bun itself does: which .env files it loads, extra
    // command-line options. Rare; not imitated.
    for name in ["NODE_ENV", "BUN_ENV", "BUN_OPTIONS", "NODE_OPTIONS"] {
        if std::env::var_os(name).is_some() {
            return Err(HandBack);
        }
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
        assert!(!mentions(b"ONENOTE_CLIENT_ID=abc\nONENOTE_CACHE_DIR=/x\n", b"TZ"));
        assert!(!mentions(b"", b"TZ"));
        assert!(!mentions(b"T", b"TZ"));
    }
}
