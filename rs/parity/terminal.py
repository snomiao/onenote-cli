#!/usr/bin/env python3
"""Runs a command with a terminal as its standard input, output and error.

    python3 rs/parity/terminal.py <command> [args...]

On a terminal onenote colours its output and makes titles clickable, and the
parity runner has no terminal. Put in front of both implementations, this
gives each a pseudo-terminal of its own and passes on what was written to it,
byte for byte (line ends arrive as CR LF, as a terminal gets them), and the
command's exit code.
"""

import os
import pty
import sys


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__, file=sys.stderr)
        return 2
    status = pty.spawn(sys.argv[1:])
    return os.waitstatus_to_exitcode(status)


if __name__ == "__main__":
    sys.exit(main())
