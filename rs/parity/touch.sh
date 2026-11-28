#!/bin/sh
# Sets the modification and access time of a file, to the nanosecond, for a
# setup line of a parity case (rs/parity.gen.mjs):
#
#   sh rs/parity/touch.sh 2026-03-08T09:05:07.998999999Z <file>
#
# `touch -d` reads its date in one way with GNU's tools and in another with
# BSD's (macOS), and node sets a file's time only as precisely as a double
# holds seconds, which is a quarter of a microsecond in 2026. Python takes
# nanoseconds, and is on Linux, on macOS and on Windows, where the command
# can be `python` rather than `python3` (a `python3` that is only the
# Microsoft Store's stand-in fails to start and is passed over). Windows
# keeps a file's time in units of 100 ns, so there the time is cut to that.
set -eu
for python in python3 python; do
  if "$python" -I -c '' > /dev/null 2>&1; then
    exec "$python" -I -c '
import calendar, os, re, sys
when, file = sys.argv[1], sys.argv[2]
match = re.fullmatch(r"(\d{4})-(\d\d)-(\d\d)T(\d\d):(\d\d):(\d\d)(?:\.(\d{1,9}))?Z", when)
if not match:
    sys.exit("touch.sh: not a time written YYYY-MM-DDTHH:MM:SS[.fraction]Z: " + when)
parts = [int(part) for part in match.groups()[:6]]
seconds = calendar.timegm((*parts, 0, 0, 0))
nanoseconds = seconds * 1_000_000_000 + int((match.group(7) or "").ljust(9, "0"))
os.utime(file, ns=(nanoseconds, nanoseconds))
' "$@"
  fi
done
echo "touch.sh: no python3 or python that starts" >&2
exit 1
