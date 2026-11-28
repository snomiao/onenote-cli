#!/bin/sh
# The TypeScript side of a parity case: `bun src/index.ts`, started from the
# case's own copy of the package (see package.sh), because the program keeps
# its search index and caches in the directory it is installed in.
exec bun "${ONENOTE_CLI_PACKAGE_ROOT:?}/src/index.ts" "$@"
