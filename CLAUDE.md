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
- Cache at `<package>/.onenote/cache/`
