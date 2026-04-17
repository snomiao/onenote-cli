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

## Cache paths
All CLI caches live at **package root** (`<pkg>/.onenote/`), NOT cwd. The CLI is project-agnostic — caches must survive across `cd` and not pollute user directories. Resolve via `dirname(import.meta.dir)` in `src/*.ts`.

- `<pkg>/.onenote/cache/<notebook>/<section>.{json,one}` — full section sync (`onenote sync`)
- `<pkg>/.onenote/notebooks.json` — notebook metadata list (24h TTL), used for name→id resolution
- `<pkg>/.onenote/assets/` — downloaded resource files (images/PDFs) referenced by `onenote read`
- `~/.onenote-cli/` — user-level state (MSAL token cache, config)

Never use `process.cwd()` for cache paths. Override via env: `ONENOTE_CACHE_DIR`, `ONENOTE_READ_ASSET_DIR`.
