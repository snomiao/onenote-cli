import {
  PublicClientApplication,
  type DeviceCodeRequest,
  type AuthenticationResult,
  type AccountInfo,
  CryptoProvider,
} from "@azure/msal-node";
import { readFile, writeFile } from "node:fs/promises";
import { readFileSync, existsSync } from "node:fs";
import { join, dirname } from "node:path";
import { homedir } from "node:os";

// Auto-load .env.local from the package root (one level up from src/) when
// running from a different working directory.
function autoLoadEnv() {
  const packageRoot = dirname(import.meta.dir);
  for (const name of [".env.local", ".env"]) {
    const path = join(packageRoot, name);
    if (!existsSync(path)) continue;
    try {
      const content = readFileSync(path, "utf-8");
      for (const line of content.split(/\r?\n/)) {
        const m = line.match(/^\s*([A-Z_][A-Z0-9_]*)\s*=\s*(.*)\s*$/);
        if (!m) continue;
        const key = m[1];
        let val = m[2];
        if ((val.startsWith('"') && val.endsWith('"')) || (val.startsWith("'") && val.endsWith("'"))) {
          val = val.slice(1, -1);
        }
        if (process.env[key] === undefined) process.env[key] = val;
      }
    } catch {}
  }
}
autoLoadEnv();

const CACHE_PATH = join(homedir(), ".onenote-cli", "msal-cache.json");
const CONFIG_PATH = join(homedir(), ".onenote-cli", "config.json");

const SCOPES = [
  "Notes.Read", "Notes.ReadWrite", "Notes.Read.All", "Notes.ReadWrite.All",
  "Files.Read", "Files.Read.All", "Files.ReadWrite", "Files.ReadWrite.All",
  "Sites.Read.All",
];

interface AppConfig {
  clientId: string;
  authority: string;
}

const DEFAULT_CONFIG: AppConfig = {
  clientId: "YOUR_CLIENT_ID",
  authority: "https://login.microsoftonline.com/common",
};

async function ensureDir(path: string) {
  const dir = path.substring(0, path.lastIndexOf("/"));
  await Bun.write(join(dir, ".keep"), "");
}

async function loadConfig(): Promise<AppConfig> {
  const envClientId = process.env.ONENOTE_CLIENT_ID;
  const envAuthority = process.env.ONENOTE_AUTHORITY;
  if (envClientId && envClientId !== "YOUR_CLIENT_ID") {
    return {
      clientId: envClientId,
      authority: envAuthority || DEFAULT_CONFIG.authority,
    };
  }

  try {
    const raw = await readFile(CONFIG_PATH, "utf-8");
    return JSON.parse(raw) as AppConfig;
  } catch {
    await ensureDir(CONFIG_PATH);
    await writeFile(CONFIG_PATH, JSON.stringify(DEFAULT_CONFIG, null, 2));
    return DEFAULT_CONFIG;
  }
}

async function createPca(): Promise<PublicClientApplication> {
  const config = await loadConfig();
  if (config.clientId === "YOUR_CLIENT_ID") {
    console.error(
      `Please configure your Azure AD app credentials in:\n  ${CONFIG_PATH}\n\n` +
        "To register an app:\n" +
        "  1. Go to https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps\n" +
        "  2. New registration -> Name: onenote-cli, Supported account types: Personal + Org\n" +
        "  3. Set platform to 'Mobile and desktop applications'\n" +
        "  4. Copy the Application (client) ID into config.json\n"
    );
    process.exit(1);
  }

  const pca = new PublicClientApplication({
    auth: {
      clientId: config.clientId,
      authority: config.authority,
    },
    cache: {
      cachePlugin: {
        beforeCacheAccess: async (ctx) => {
          try {
            const data = await readFile(CACHE_PATH, "utf-8");
            ctx.tokenCache.deserialize(data);
          } catch {
            // no cache yet
          }
        },
        afterCacheAccess: async (ctx) => {
          if (ctx.cacheHasChanged) {
            await ensureDir(CACHE_PATH);
            await writeFile(CACHE_PATH, ctx.tokenCache.serialize());
          }
        },
      },
    },
  });

  return pca;
}

export async function listAccounts(): Promise<AccountInfo[]> {
  const pca = await createPca();
  return pca.getTokenCache().getAllAccounts() as Promise<AccountInfo[]>;
}

// Module-level "current" account used by getAccessToken when no account is passed.
// syncCache sets this to loop over each cached account.
let currentAccount: AccountInfo | undefined;
export function setCurrentAccount(account: AccountInfo | undefined): void {
  currentAccount = account;
}

async function deviceCodeFlow(pca: PublicClientApplication): Promise<AuthenticationResult> {
  const request: DeviceCodeRequest = {
    scopes: SCOPES,
    deviceCodeCallback: async (response) => {
      console.error(response.message);
      // Auto-copy code to clipboard for convenience
      try {
        const clipboardy = await import("clipboardy");
        await clipboardy.default.write(response.userCode);
        console.error(`(Code "${response.userCode}" copied to clipboard)`);
      } catch {}
    },
  };
  const result = await pca.acquireTokenByDeviceCode(request);
  if (!result) throw new Error("Authentication failed");
  return result;
}

/** Extract the auth code (and state) from a pasted redirect URL or its query string. */
function parseRedirectInput(input: string): { code?: string; state?: string; error?: string } {
  const text = input.trim();
  if (!text) return {};
  if (!text.includes("code=") && !text.includes("error=")) return {};
  const query = text.includes("?") ? text.slice(text.indexOf("?") + 1) : text;
  const params = new URLSearchParams(query.split("#")[0]);
  const error = params.get("error");
  if (error) return { error: `${error}: ${params.get("error_description") ?? ""}`.trim() };
  return { code: params.get("code") ?? undefined, state: params.get("state") ?? undefined };
}

const page = (title: string, body: string) =>
  "<html><body style='font-family:system-ui;padding:40px;text-align:center'>" +
  `<h2>${title}</h2><p>${body}</p></body></html>`;

/**
 * Interactive auth via loopback redirect (PKCE). Starts a local server on a free port and opens
 * the browser to the Microsoft sign-in URL with redirect_uri=http://localhost:PORT.
 *
 * When the CLI runs on a remote host (VPS/SSH) but the browser is on your desktop, the redirect
 * to localhost fails to load — in that case paste the redirect URL from the address bar into the
 * terminal. Whichever arrives first (loopback or paste) completes the sign-in.
 */
async function interactiveFlow(pca: PublicClientApplication): Promise<AuthenticationResult> {
  const crypto = new CryptoProvider();
  const { verifier, challenge } = await crypto.generatePkceCodes();
  const expectedState = crypto.createNewGuid();

  let settle!: (r: { code?: string; state?: string; error?: string }) => void;
  const received = new Promise<{ code?: string; state?: string; error?: string }>((r) => (settle = r));

  const server = Bun.serve({
    hostname: "localhost",
    port: 0,
    fetch(req) {
      const url = new URL(req.url);
      if (url.pathname !== "/") return new Response("Not found", { status: 404 });
      const parsed = parseRedirectInput(url.search);
      if (!parsed.code && !parsed.error) return new Response("Waiting for sign-in...");
      // Defer so this response is flushed before the CLI proceeds (and possibly exits on error).
      setTimeout(() => settle(parsed), 50);
      const ok = !!parsed.code;
      return new Response(
        ok
          ? page("Signed in successfully", "You can close this tab and return to the terminal.")
          : page("Sign-in failed", "Please return to the terminal and try again."),
        { headers: { "Content-Type": "text/html; charset=utf-8" } },
      );
    },
  });
  const redirectUri = `http://localhost:${server.port}`;

  const { createInterface } = await import("node:readline");
  const rl = createInterface({ input: process.stdin, terminal: false });
  rl.on("line", (line) => {
    const parsed = parseRedirectInput(line);
    if (parsed.code || parsed.error) settle(parsed);
    else if (line.trim()) console.error("Couldn't find a code in that input — paste the full redirect URL.");
  });

  try {
    const authUrl = await pca.getAuthCodeUrl({
      scopes: SCOPES,
      redirectUri,
      codeChallenge: challenge,
      codeChallengeMethod: "S256",
      state: expectedState,
    });

    const cmd = process.platform === "darwin" ? "open" : process.platform === "win32" ? "start" : "xdg-open";
    try {
      const { spawn } = await import("node:child_process");
      spawn(cmd, [authUrl], { stdio: "ignore", detached: true }).on("error", () => {}).unref();
    } catch {}
    console.error(
      `\nIf your browser didn't open, visit:\n  ${authUrl}\n\n` +
        `After signing in you'll be redirected to ${redirectUri}/?code=...\n` +
        "If that page fails to load (e.g. the CLI runs on a remote server / VPS),\n" +
        "copy the full URL from the browser's address bar and paste it here, then press Enter:",
    );

    const { code, state, error } = await received;
    if (error) throw new Error(`Sign-in failed — ${error}`);
    if (state && state !== expectedState) throw new Error("Sign-in failed — state mismatch (stale or foreign redirect URL)");
    const result = await pca.acquireTokenByCode({ code: code!, scopes: SCOPES, redirectUri, codeVerifier: verifier });
    if (!result) throw new Error("Authentication failed");
    return result;
  } finally {
    rl.close();
    process.stdin.pause();
    server.stop();
  }
}

/** Add a new account via interactive browser auth. */
export async function login(): Promise<string> {
  const config = await loadConfig();
  if (config.clientId === "YOUR_CLIENT_ID") {
    console.error(`Please configure your Azure AD app credentials in:\n  ${CONFIG_PATH}`);
    process.exit(1);
  }
  // No cachePlugin: prevents silent reuse of an existing cached token
  const freshPca = new PublicClientApplication({
    auth: { clientId: config.clientId, authority: config.authority },
  });
  console.error("Opening browser for sign-in...");
  const result = await interactiveFlow(freshPca);

  // Merge new account's tokens into the shared on-disk cache
  const newTokens = JSON.parse(freshPca.getTokenCache().serialize()) as Record<string, Record<string, unknown>>;
  let existing: Record<string, Record<string, unknown>> = {};
  try {
    existing = JSON.parse(await readFile(CACHE_PATH, "utf-8"));
  } catch {}
  const merged = { ...existing };
  for (const [section, entries] of Object.entries(newTokens)) {
    merged[section] = { ...merged[section], ...entries };
  }
  await ensureDir(CACHE_PATH);
  await writeFile(CACHE_PATH, JSON.stringify(merged));
  return result.account?.username ?? "(unknown)";
}

/** Get an access token silently for the given account (or module-level currentAccount, or first cached account). */
export async function getAccessToken(account?: AccountInfo): Promise<string> {
  const pca = await createPca();
  const accounts = await pca.getTokenCache().getAllAccounts() as AccountInfo[];

  // Resolve: explicit > module-level current > match by homeAccountId > first
  let target = account ?? currentAccount;
  if (target) {
    // Re-fetch from the shared cache to ensure we have the latest state
    target = accounts.find((a) => a.homeAccountId === target!.homeAccountId) ?? target;
  } else {
    target = accounts[0];
  }
  if (target) {
    try {
      const result = await pca.acquireTokenSilent({ account: target, scopes: SCOPES });
      return result.accessToken;
    } catch {
      // silent failed — fall through to device code for this account
    }
  }

  // No cached account or silent refresh failed
  const result = await deviceCodeFlow(pca);
  return result.accessToken;
}

export async function logout(email?: string): Promise<void> {
  const pca = await createPca();
  const accounts = await pca.getTokenCache().getAllAccounts() as AccountInfo[];

  // Option C: 1 account → logout that one; multiple + no email → show list
  if (!email) {
    if (accounts.length === 0) {
      console.log("Already logged out (no cached tokens).");
      return;
    }
    if (accounts.length === 1) {
      email = accounts[0].username;
    } else {
      console.log("Multiple accounts found. Specify one or use --all:");
      for (const a of accounts) {
        console.log(`  ${a.username}`);
      }
      console.log("\nExample: onenote auth logout snomiao@snomiao.com");
      console.log("         onenote auth logout --all");
      return;
    }
  }

  const target = accounts.find((a) => a.username?.toLowerCase() === email!.toLowerCase());
  if (!target) {
    console.log(`Account '${email}' not found. Logged-in accounts:`);
    for (const a of accounts) console.log(`  ${a.username}`);
    return;
  }

  await pca.getTokenCache().removeAccount(target);
  console.log(`Logged out ${target.username}.`);
}

export async function logoutAll(): Promise<void> {
  try {
    const { unlink } = await import("node:fs/promises");
    await unlink(CACHE_PATH);
    console.log("Logged out all accounts. Token cache removed.");
  } catch {
    console.log("Already logged out (no cached tokens).");
  }
}

export async function whoami(): Promise<void> {
  const pca = await createPca();
  const accounts = await pca.getTokenCache().getAllAccounts() as AccountInfo[];
  if (accounts.length === 0) {
    console.log("Not logged in. Run `onenote auth login` to authenticate.");
    return;
  }
  for (const account of accounts) {
    console.log(`${account.username}  (${account.name ?? ""})`);
    console.log(`  tenant: ${account.tenantId ?? "unknown"}  env: ${account.environment ?? "unknown"}`);
  }
}
