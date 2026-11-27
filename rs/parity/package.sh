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
#                     src/ and package.json, with node_modules linked.
#   --handback check  the runner points the variable at its stand-in package,
#                     one for all cases, whose src/index.ts only reports that
#                     it was reached: that is left alone, and what an earlier
#                     case put beside it is removed.
set -eu
here=$(cd "$(dirname "$0")" && pwd)
repo=$(cd "$here/../.." && pwd)
root=${ONENOTE_CLI_PACKAGE_ROOT:?}
case "$root" in
  */rustify-parity-*) ;;
  *) echo "package.sh: $root is not inside a parity run's directory" >&2; exit 1 ;;
esac
if [ -f "$root/package.json" ]; then
  rm -rf "$root/.onenote" "$root/.env" "$root/.env.local"
else
  mkdir -p "$root"
  cp -R "$repo/src" "$repo/package.json" "$root/"
  ln -s "$repo/node_modules" "$root/node_modules"
fi
