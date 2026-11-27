#!/usr/bin/env python3
"""Runs a command with a terminal as its standard input, output and error.

    python3 rs/parity/terminal.py <command> [args...]

On a terminal onenote colours its output and makes titles clickable, and the
parity runner has no terminal. Put in front of both implementations, this
gives each a pseudo-terminal of its own and passes on what was written to it,
byte for byte (line ends arrive as CR LF, as a terminal gets them), and the
command's exit code. What comes in on its own standard input is typed into
the terminal.

This is what pty.spawn does, with the copying written out: in Python 3.9,
the python3 of macOS's command line tools, pty.spawn waits forever once the
command has ended there, because macOS's terminal reports the end as an
empty read and not as an error. Linux and macOS only; Windows has no pty.
"""

import os
import pty
import select
import sys


def write_all(fd: int, data: bytes) -> None:
    while data:
        data = data[os.write(fd, data):]


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__, file=sys.stderr)
        return 2
    pid, master = pty.fork()
    if pid == 0:
        try:
            os.execvp(sys.argv[1], sys.argv[1:])
        finally:
            os._exit(127)
    stdin = sys.stdin.fileno()
    reading = [master, stdin]
    while True:
        ready, _, _ = select.select(reading, [], [])
        if master in ready:
            try:
                data = os.read(master, 65536)
            except OSError:
                # Linux: EIO once the command and everything it started has
                # let go of the terminal.
                data = b""
            if not data:
                break
            write_all(sys.stdout.fileno(), data)
        if stdin in ready:
            data = os.read(stdin, 65536)
            if data:
                write_all(master, data)
            else:
                reading.remove(stdin)
    os.close(master)
    _, status = os.waitpid(pid, 0)
    return os.waitstatus_to_exitcode(status)


if __name__ == "__main__":
    sys.exit(main())
