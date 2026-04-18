# CLAUDE.md

## Build & Run
- `bun install` — install deps
- `bun run start` — run CLI
- `bun run lint` — lint with oxlint
- `bun run lint:fix` — auto-fix lint issues
- `bun run typecheck` — type check with tsgo

## Commands
- `onenote auth login/logout/whoami/setup` — authentication
- `onenote notebooks list/get/create` — notebook management
- `onenote sections list/get/create` — section management
- `onenote pages list/get/content/create/delete` — page management
- `onenote sync` — download and cache all sections for local search
- `onenote search <query>` — full-text page-level search (local cache)
- `onenote search <query> --online` — online section-level search (Graph API)

## Conventions
- Use `import.meta.main` for script entry points
- Use `import.meta.dir` for package-relative paths (env, cache)
- Conventional commits for all commit messages
- `.env.local` at package root for credentials (auto-loaded)

## Safe by default
Destructive operations (anything that deletes or overwrites user content in OneNote) must be opt-in, not default. Without an explicit confirmation flag, they dry-run: show what would be affected + a short content sha, then exit.

- `rm` / `pages delete` / `pages update --action=replace` require `--sha=<4-char>` (sha256 of page HTML, first 4 hex chars). Mismatch → refuse.
- `read` prints the same sha to stderr so users can copy it straight from a prior read.
- Non-destructive ops (`cp`, `append`, additive `update` actions, `rename`, `create`) don't need confirmation.
- When in doubt between atomic-but-risky and two-step-but-safe (e.g. section move), ship the safe version first and leave the destructive step manual.

## Cache paths
All CLI caches live at **package root** (`<pkg>/.onenote/`), NOT cwd. The CLI is project-agnostic — caches must survive across `cd` and not pollute user directories. Resolve via `dirname(import.meta.dir)` in `src/*.ts`.

- `<pkg>/.onenote/cache/<notebook>/<section>.{json,one}` — full section sync (`onenote sync`)
- `<pkg>/.onenote/notebooks.json` — notebook metadata list (24h TTL), used for name→id resolution
- `<pkg>/.onenote/assets/` — downloaded resource files (images/PDFs) referenced by `onenote read`
- `~/.onenote-cli/` — user-level state (MSAL token cache, config)

Never use `process.cwd()` for cache paths. Override via env: `ONENOTE_CACHE_DIR`, `ONENOTE_READ_ASSET_DIR`.
