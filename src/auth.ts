import {
  PublicClientApplication,
  type DeviceCodeRequest,
  type AuthenticationResult,
  type AccountInfo,
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

export async function getAccessToken(account?: AccountInfo): Promise<string> {
  const pca = await createPca();

  const accounts = await pca.getTokenCache().getAllAccounts();

  // Resolve which account to use
  let target: AccountInfo | undefined = account;
  if (!target && accounts.length > 0) {
    target = accounts[0] as AccountInfo;
  }

  if (target) {
    try {
      const result = await pca.acquireTokenSilent({
        account: target as AccountInfo,
        scopes: SCOPES,
      });
      return result.accessToken;
    } catch {
      // fall through to device code
    }
  }

  // Device code flow (adds new account to cache without removing existing ones)
  const request: DeviceCodeRequest = {
    scopes: SCOPES,
    deviceCodeCallback: (response) => {
      console.error(response.message);
    },
  };

  const result = await pca.acquireTokenByDeviceCode(request);
  if (!result) throw new Error("Authentication failed");
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
