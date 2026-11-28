#!/bin/sh
# Gives a parity case the package directory it runs in. First setup line of
# every case in rs/parity.json.
#
# onenote keeps its search index, its caches and its .env.local in the
# directory it is installed in (dirname(import.meta.dir) in src/*.ts), not in
# HOME or the current directory. A case therefore needs a package directory of
# its own, or it would read and write the checkout's .onenote/. The cases name
# it in ONENOTE_CLI_PACKAGE_ROOT: the Rust binary reads the package location
# from there, and rs/parity/ts.sh starts the TypeScript from there.
#
#   parity run        the directory does not exist yet: it becomes a copy of
#                     src/ and package.json, with node_modules linked (on
#                     Windows by a junction: Git Bash's `ln -s` copies a
#                     directory unless it is allowed to make links, and a
#                     junction needs no permission; node and bun remove a
#                     junction without going into it).
#   --handback check  the runner points the variable at its stand-in package,
#                     one for all cases, whose src/index.ts only reports that
#                     it was reached: that is left alone, and what an earlier
#                     case put beside it is removed.
set -eu
# Without starting a program where that can be done: on Windows every
# program Git Bash starts costs a tenth of a second, and every case runs
# this once for each side. The cases start it as <REPO>/rs/parity/package.sh.
case "$0" in
  */rs/parity/package.sh) repo=${0%/rs/parity/package.sh} ;;
  *) repo=$(cd "$(dirname "$0")/../.." && pwd) ;;
esac
root=${ONENOTE_CLI_PACKAGE_ROOT:?}
# (On Windows the runner names its stand-in package with backslashes.)
case "$root" in
  */rustify-parity-* | *\\rustify-parity-*) ;;
  *) echo "package.sh: $root is not inside a parity run's directory" >&2; exit 1 ;;
esac
if [ -f "$root/package.json" ]; then
  rm -rf "$root/.onenote" "$root/.env" "$root/.env.local"
else
  mkdir -p "$root"
  cp -R "$repo/src" "$repo/package.json" "$root/"
  case "${OSTYPE:-}" in
    msys* | cygwin*)
      # Git Bash. node makes the junction: cmd's mklink would need the
      # directories quoted in cmd's way, which Git Bash cannot give it for a
      # name with & or % in it.
      node -e 'require("fs").symlinkSync(process.argv[1], process.argv[2], "junction")' "$repo/node_modules" "$root/node_modules"
      ;;
    *) ln -s "$repo/node_modules" "$root/node_modules" ;;
  esac
fi
