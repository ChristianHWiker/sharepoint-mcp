import {
  PublicClientApplication,
  ConfidentialClientApplication,
  type Configuration,
  type TokenCacheContext,
} from "@azure/msal-node";
import fs from "node:fs";
import path from "node:path";
import os from "node:os";

const DELEGATED_SCOPES = [
  "Sites.ReadWrite.All",
  "Sites.Manage.All",
  "Files.ReadWrite.All",
  "User.Read",
];

const APP_ONLY_SCOPES = ["https://graph.microsoft.com/.default"];

const CACHE_DIR = path.join(
  process.env.APPDATA ?? path.join(os.homedir(), ".config"),
  "sharepoint-mcp",
);
const CACHE_FILE = path.join(CACHE_DIR, "token-cache.json");

function ensureCacheDir() {
  if (!fs.existsSync(CACHE_DIR)) fs.mkdirSync(CACHE_DIR, { recursive: true });
}

const cachePlugin = {
  beforeCacheAccess: async (ctx: TokenCacheContext) => {
    if (fs.existsSync(CACHE_FILE)) {
      ctx.tokenCache.deserialize(fs.readFileSync(CACHE_FILE, "utf8"));
    }
  },
  afterCacheAccess: async (ctx: TokenCacheContext) => {
    if (ctx.cacheHasChanged) {
      ensureCacheDir();
      fs.writeFileSync(CACHE_FILE, ctx.tokenCache.serialize());
    }
  },
};

type AuthMode =
  | { kind: "delegated"; client: PublicClientApplication }
  | { kind: "app-only"; client: ConfidentialClientApplication };

let cached: AuthMode | null = null;

function buildAuth(): AuthMode {
  const clientId = process.env.AZURE_CLIENT_ID;
  const tenantId = process.env.AZURE_TENANT_ID;
  const clientSecret = process.env.AZURE_CLIENT_SECRET;

  if (!clientId || !tenantId) {
    throw new Error(
      "AZURE_CLIENT_ID and AZURE_TENANT_ID must be set (see .env.example).",
    );
  }

  const authority = `https://login.microsoftonline.com/${tenantId}`;

  if (clientSecret) {
    const config: Configuration = {
      auth: { clientId, authority, clientSecret },
    };
    return { kind: "app-only", client: new ConfidentialClientApplication(config) };
  }

  const config: Configuration = {
    auth: { clientId, authority },
    cache: { cachePlugin },
  };
  return { kind: "delegated", client: new PublicClientApplication(config) };
}

function auth(): AuthMode {
  if (!cached) cached = buildAuth();
  return cached;
}

export function authMode(): "delegated" | "app-only" {
  return auth().kind;
}

async function acquireDelegated(pca: PublicClientApplication): Promise<string> {
  const accounts = await pca.getTokenCache().getAllAccounts();

  if (accounts.length > 0) {
    try {
      const result = await pca.acquireTokenSilent({
        account: accounts[0],
        scopes: DELEGATED_SCOPES,
      });
      if (result?.accessToken) return result.accessToken;
    } catch {
      // fall through to device code
    }
  }

  const result = await pca.acquireTokenByDeviceCode({
    scopes: DELEGATED_SCOPES,
    deviceCodeCallback: (response) => {
      // stdout is the MCP transport — must write to stderr
      process.stderr.write(`\n${response.message}\n`);
    },
  });

  if (!result?.accessToken) {
    throw new Error("Failed to acquire delegated access token.");
  }
  return result.accessToken;
}

async function acquireAppOnly(
  cca: ConfidentialClientApplication,
): Promise<string> {
  const result = await cca.acquireTokenByClientCredential({
    scopes: APP_ONLY_SCOPES,
  });
  if (!result?.accessToken) {
    throw new Error("Failed to acquire app-only access token.");
  }
  return result.accessToken;
}

export async function getAccessToken(): Promise<string> {
  const a = auth();
  if (a.kind === "app-only") return acquireAppOnly(a.client);
  return acquireDelegated(a.client);
}

export async function getSharePointToken(hostname: string): Promise<string> {
  const a = auth();
  const resourceScope = `https://${hostname}/.default`;

  if (a.kind === "app-only") {
    const result = await a.client.acquireTokenByClientCredential({
      scopes: [resourceScope],
    });
    if (!result?.accessToken) {
      throw new Error("Failed to acquire app-only SharePoint access token.");
    }
    return result.accessToken;
  }

  const accounts = await a.client.getTokenCache().getAllAccounts();
  if (accounts.length > 0) {
    try {
      const result = await a.client.acquireTokenSilent({
        account: accounts[0],
        scopes: [resourceScope],
      });
      if (result?.accessToken) return result.accessToken;
    } catch {
      // fall through
    }
  }

  const result = await a.client.acquireTokenByDeviceCode({
    scopes: [resourceScope],
    deviceCodeCallback: (response) => {
      process.stderr.write(`\n${response.message}\n`);
    },
  });

  if (!result?.accessToken) {
    throw new Error("Failed to acquire delegated SharePoint access token.");
  }
  return result.accessToken;
}
