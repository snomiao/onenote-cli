"""The TypeScript side of a parity case on Windows: what rs/parity/ts.sh is
elsewhere, `bun <the case's package>/src/index.ts <arguments>`.

    python rs/parity/ts.py <arguments>

On Windows ts.sh would run in Git Bash's sh, and every program of that
family drops TZ from the environment of a Windows program it starts (a TZ
that names a zone, such as America/New_York; Windows reads TZ otherwise).
The TypeScript would then show the system's local time where the binary,
which the parity runner starts itself, shows the zone's. Python passes the
environment on as it is, and so does the runner; it reads no variable of
the JavaScript runtime's. bun is started by its name, as the binary's
hand-back starts it, so that what bun says about itself (`bun
--trace-warnings`) is the same on both sides.
"""

import os
import subprocess
import sys

root = os.environ.get("ONENOTE_CLI_PACKAGE_ROOT")
if not root:
    sys.exit("ts.py: ONENOTE_CLI_PACKAGE_ROOT has to name the case's package")
# The path as ts.sh writes it.
sys.exit(subprocess.run(["bun", f"{root}/src/index.ts", *sys.argv[1:]]).returncode)
