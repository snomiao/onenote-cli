#!/usr/bin/env python3
"""Runs a command with a terminal as its standard input, output and error.

    python3 rs/parity/terminal.py <command> [args...]

On a terminal onenote colours its output and makes titles clickable, and the
parity runner has no terminal. Put in front of both implementations, this
gives each a pseudo-terminal of its own and passes on what was written to it,
byte for byte (line ends arrive as CR LF, as a terminal gets them), and the
command's exit code. What comes in on its own standard input is typed into
the terminal. The terminal is the command's controlling terminal, in a
session of its own, as pty.fork makes it.

This process holds the terminal open until the command has ended and what
it wrote has been read. On macOS the last close of a terminal waits at most
0.6 s for the other end to read what is still in it, and then throws it away
(60 ticks: ptsclose in XNU). The command's exit was that last close once
pty.fork had let go of the terminal here, so a command that printed and
ended at once lost all of its output whenever this process did not get to
run for 0.6 s, as on a busy CI runner: the binary did, and bun too when it
is started without sh. With the terminal held here, the command's exit waits
until its output has been read, however long that takes. (macOS's sh, which
ts.sh goes through, is bash, and bash gets the same by opening /dev/tty when
it starts; dash and zsh do not.) Linux keeps the output either way.

The copying is written out, as pty.spawn would do it: in Python 3.9, the
python3 of macOS's command line tools, pty.spawn waits forever once the
command has ended there, because macOS's terminal reports the end as an
empty read and not as an error. Linux and macOS only; Windows has no pty.
"""

import fcntl
import os
import select
import signal
import sys
import termios


def write_all(fd: int, data: bytes) -> None:
    while data:
        data = data[os.write(fd, data):]


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__, file=sys.stderr)
        return 2
    master, slave = os.openpty()
    # A byte comes down this pipe for every signal, so that the wait below
    # wakes up when the command has ended.
    wake_read, wake_write = os.pipe()
    os.set_blocking(wake_read, False)
    os.set_blocking(wake_write, False)
    signal.signal(signal.SIGCHLD, lambda number, frame: None)
    signal.set_wakeup_fd(wake_write)
    pid = os.fork()
    if pid == 0:
        try:
            # What pty.fork does in the child.
            os.setsid()
            if hasattr(termios, "TIOCSCTTY"):
                fcntl.ioctl(slave, termios.TIOCSCTTY, 0)
            else:
                os.close(os.open(os.ttyname(slave), os.O_RDWR))
            for fd in (0, 1, 2):
                os.dup2(slave, fd)
            for fd in (slave, master, wake_read, wake_write):
                if fd > 2:
                    os.close(fd)
            os.execvp(sys.argv[1], sys.argv[1:])
        finally:
            os._exit(127)

    def copy() -> bool:
        """Passes on what the terminal has; False once it has ended."""
        try:
            data = os.read(master, 65536)
        except OSError:
            # Linux: EIO once the command and everything it started has
            # let go of the terminal.
            data = b""
        write_all(sys.stdout.fileno(), data)
        return bool(data)

    stdin = sys.stdin.fileno()
    reading = [master, stdin, wake_read]
    status = None
    held = True
    while True:
        # A second at most while the command runs: the wake-up is the rule,
        # this is for a signal that got lost.
        ready, _, _ = select.select(reading, [], [], 1.0 if status is None else None)
        if master in ready and not copy():
            break
        if stdin in ready:
            data = os.read(stdin, 65536)
            if data:
                write_all(master, data)
            else:
                reading.remove(stdin)
        if status is None:
            try:
                os.read(wake_read, 4096)
            except BlockingIOError:
                pass
            done, waited = os.waitpid(pid, os.WNOHANG)
            if done:
                status = waited
                reading.remove(wake_read)
                # Nothing more comes from the command. Read what the terminal
                # still holds, let go of it, and go on until whatever else
                # the command started has let go of it too, as before.
                ended = False
                while not ended and select.select([master], [], [], 0)[0]:
                    ended = not copy()
                os.close(slave)
                held = False
                if ended:
                    break
    if held:
        os.close(slave)
    os.close(master)
    if status is None:
        _, status = os.waitpid(pid, 0)
    return os.waitstatus_to_exitcode(status)


if __name__ == "__main__":
    sys.exit(main())
