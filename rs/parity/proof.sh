#!/bin/sh
# Everything that shows the binary behaves like the TypeScript, in one go.
#
#   sh rs/parity/proof.sh <rustify's templates/parity/parity.mjs> <the onenote-cli binary>
#
# after `bun install`. The CI workflow runs this against the Linux binary it
# has just built. First rs/parity.json is checked to be what
# rs/parity.gen.mjs writes; then four things run side by side (the first
# alone takes six minutes, and the CI job has thirty for all of them):
#
#   plain     every case, TypeScript against the binary
#   terminal  the same with a pseudo-terminal as standard output, where the
#             search prints colours and clickable titles
#   handback  the binary alone against a stand-in package, which fails unless
#             exactly the cases marked "handback" reach the JavaScript; then
#             the same on a pseudo-terminal
#   zones     `cargo test` in rs/: the unit tests, among them the one that
#             compares every time zone the binary reads with the runtime's
#             own time zone database, for 1970 to 2099 (rs/src/time.rs).
#             This one builds the crate again, from the same sources and
#             Cargo.lock; it does not run the binary named above.
#
# Each writes to a file. At the end one line is printed for each that
# passed, and all of what a failed one wrote. The exit code is 0 only when
# all four passed.
set -u
usage="usage: proof.sh <parity.mjs> <binary>"
absolute() {
  case "$1" in
    /*) printf '%s\n' "$1" ;;
    *) printf '%s/%s\n' "$PWD" "$1" ;;
  esac
}
runner=$(absolute "${1:?$usage}")
binary=$(absolute "${2:?$usage}")
cd "$(dirname "$0")/../.." || exit 1

node rs/parity.gen.mjs --check || exit 1

logs=$(mktemp -d)
start() {
  name=$1
  shift
  (
    "$@" > "$logs/$name.log" 2>&1
    echo $? > "$logs/$name.exit"
  ) &
}
handback() {
  node "$runner" --cases rs/parity.json --rs "$binary" --handback ONENOTE_CLI_PACKAGE_ROOT=src/index.ts &&
    node "$runner" --cases rs/parity.json --rs "python3 rs/parity/terminal.py $binary" --handback ONENOTE_CLI_PACKAGE_ROOT=src/index.ts
}
start plain node "$runner" --cases rs/parity.json --ts "sh rs/parity/ts.sh" --rs "$binary"
start terminal node "$runner" --cases rs/parity.json --ts "python3 rs/parity/terminal.py sh rs/parity/ts.sh" --rs "python3 rs/parity/terminal.py $binary"
start handback handback
start zones cargo test --locked --manifest-path rs/Cargo.toml -- --nocapture
wait

failed=0
for name in plain terminal handback zones; do
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
