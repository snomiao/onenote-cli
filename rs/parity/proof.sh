#!/bin/sh
# Everything that shows the binary behaves like the TypeScript, in one go.
#
#   sh rs/parity/proof.sh <rustify's templates/parity/parity.mjs> <the onenote-cli binary>
#
# after `bun install`, on Linux, macOS or Windows (in Git Bash there). The
# CI workflow runs this on each of the three, against the binary it has just
# built for that system ($RUSTIFY_PLATFORM names the system; without it,
# `uname` does). First rs/parity.json is checked to be what rs/parity.gen.mjs
# writes; then these run side by side (the first alone takes six minutes on
# Linux, and the CI job has thirty for all of them):
#
#   plain     every case, TypeScript against the binary
#   terminal  the same with a pseudo-terminal as standard input, output and
#             error, where the search prints colours and clickable titles
#             (rs/parity/terminal.py, python3's pty)
#   handback  the binary alone against a stand-in package, which fails unless
#             exactly the cases marked "handback" for this system reach the
#             JavaScript; then the same on a pseudo-terminal
#   zones     `cargo test` in rs/: the unit tests, among them the one that
#             compares every time zone the binary reads with the runtime's
#             own time zone database, for 1970 to 2099 (rs/src/time.rs).
#             This one builds the crate again, from the same sources and
#             Cargo.lock; it does not run the binary named above.
#
# On Windows, where this runs in Git Bash:
#   - There is no pseudo-terminal (python's pty is for Linux and macOS, and
#     the parity runner has none there either), so `terminal` and the second
#     half of `handback` are left out; the binary hands a search on a Windows
#     console back anyway (rs/src/main.rs, machine).
#   - Nor is there a time zone to compare: the search, which alone reads
#     one, is not compiled for Windows, and with it rs/src/time.rs. So there
#     `cargo test` runs as `unit`, the unit tests of what is compiled.
#   - The TypeScript is started by rs/parity/ts.py instead of ts.sh: Git
#     Bash's sh would drop TZ on the way to bun.
#   - Every program a case starts takes a tenth of a second or more to start,
#     so a case takes seconds. `plain` and `handback` therefore run in lanes,
#     side by side, each lane taking every fourth (every second) group of
#     cases that rs/parity.gen.mjs --groups names, one run of the parity
#     runner per group.
#
# Each writes to a file. At the end one line is printed for each that
# passed (and how many cases were skipped on this system), and all of what a
# failed one wrote. The exit code is 0 only when all of them passed.
set -u
usage="usage: proof.sh <parity.mjs> <binary>"
absolute() {
  case "$1" in
    /* | [A-Za-z]:[\\/]*) printf '%s\n' "$1" ;;
    *) printf '%s/%s\n' "$here" "$1" ;;
  esac
}
# In Git Bash `pwd -W` gives the directory as Windows programs (node, the
# binary) read it.
here=$(pwd -W 2> /dev/null || pwd)
runner=$(absolute "${1:?$usage}")
binary=$(absolute "${2:?$usage}")
cd "$(dirname "$0")/../.." || exit 1

platform=${RUSTIFY_PLATFORM:-}
if [ -z "$platform" ]; then
  case "$(uname -s)" in
    Linux) platform=linux ;;
    Darwin) platform=darwin ;;
    MINGW* | MSYS* | CYGWIN*) platform=win32 ;;
    *) platform=$(uname -s) ;;
  esac
fi
case "$platform" in
  win32*) runs="plain0 plain1 plain2 plain3 handback0 handback1 unit" ts="python -I rs/parity/ts.py" ;;
  *) runs="plain terminal handback zones" ts="sh rs/parity/ts.sh" ;;
esac

node rs/parity.gen.mjs --check || exit 1

# In $TMPDIR, which macOS's mktemp passes over without a template.
logs=$(mktemp -d "${TMPDIR:-/tmp}/onenote-proof.XXXXXX")
start() {
  name=$1
  shift
  (
    "$@" > "$logs/$name.log" 2>&1
    echo $? > "$logs/$name.exit"
  ) &
}
handback() {
  node "$runner" --cases rs/parity.json --rs "$binary" --handback ONENOTE_CLI_PACKAGE_ROOT=src/index.ts || return
  node "$runner" --cases rs/parity.json --rs "python3 rs/parity/terminal.py $binary" --handback ONENOTE_CLI_PACKAGE_ROOT=src/index.ts
}
node rs/parity.gen.mjs --groups > "$logs/groups" 2> /dev/null || exit 1
# lane <number> <of how many> <runner options...>: the runner once for each
# of those groups, in turn.
lane() {
  number=$1 lanes=$2
  shift 2
  code=0 at=0
  while IFS= read -r group; do
    if [ $((at % lanes)) = "$number" ]; then
      node "$runner" --cases rs/parity.json --filter "$group" "$@" || code=1
    fi
    at=$((at + 1))
  done < "$logs/groups"
  return $code
}
for name in $runs; do
  case $name in
    plain) start plain node "$runner" --cases rs/parity.json --ts "$ts" --rs "$binary" ;;
    plain[0-9]) start "$name" lane "${name#plain}" 4 --ts "$ts" --rs "$binary" ;;
    terminal) start terminal node "$runner" --cases rs/parity.json --ts "python3 rs/parity/terminal.py $ts" --rs "python3 rs/parity/terminal.py $binary" ;;
    handback) start handback handback ;;
    handback[0-9]) start "$name" lane "${name#handback}" 2 --rs "$binary" --handback ONENOTE_CLI_PACKAGE_ROOT=src/index.ts ;;
    zones | unit) start "$name" cargo test --locked --manifest-path rs/Cargo.toml -- --nocapture ;;
  esac
done
wait

echo "proof on $platform: $runs"
failed=0
for name in $runs; do
  code=$(cat "$logs/$name.exit" 2> /dev/null || echo "no exit code")
  if [ "$code" = 0 ]; then
    grep -E 'case\(s\)|^test result|answers compared' "$logs/$name.log" | sed "s/^/$name: /"
  else
    echo "$name: FAILED ($code)"
    cat "$logs/$name.log"
    failed=1
  fi
done
rm -rf "$logs"
exit $failed
