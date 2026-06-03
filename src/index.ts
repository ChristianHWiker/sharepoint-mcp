#!/usr/bin/env node
import { config as loadDotenv } from "dotenv";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
for (const p of [
  path.resolve(scriptDir, "../.env"),
  path.resolve(process.cwd(), ".env"),
]) {
  if (fs.existsSync(p)) {
    loadDotenv({ path: p, quiet: true });
    break;
  }
}

const { McpServer } = await import("@modelcontextprotocol/sdk/server/mcp.js");
const { StdioServerTransport } = await import(
  "@modelcontextprotocol/sdk/server/stdio.js"
);
const { registerSiteTools } = await import("./tools/sites.js");
const { registerPageTools } = await import("./tools/pages.js");
const { registerFileTools } = await import("./tools/files.js");
const { registerListTools } = await import("./tools/lists.js");
const { registerSearchTools } = await import("./tools/search.js");
const { registerPermissionTools } = await import("./tools/permissions.js");
const { registerContentTypeTools } = await import("./tools/contentTypes.js");
const { registerTaxonomyTools } = await import("./tools/taxonomy.js");
const { registerSubscriptionTools } = await import("./tools/subscriptions.js");
const { authMode, getAccessToken } = await import("./auth.js");

async function runServer() {
  const server = new McpServer({
    name: "sharepoint-mcp",
    version: "0.1.0",
  });

  registerSiteTools(server);
  registerPageTools(server);
  registerFileTools(server);
  registerListTools(server);
  registerSearchTools(server);
  registerPermissionTools(server);
  registerContentTypeTools(server);
  registerTaxonomyTools(server);
  registerSubscriptionTools(server);

  const transport = new StdioServerTransport();
  await server.connect(transport);
  process.stderr.write(`[sharepoint-mcp] ready (auth: ${authMode()})\n`);
}

async function runAuth() {
  const mode = authMode();
  process.stderr.write(`[sharepoint-mcp] auth mode: ${mode}\n`);
  await getAccessToken();
  if (mode === "delegated") {
    process.stderr.write(
      `[sharepoint-mcp] sign-in complete; token cached.\n`,
    );
  } else {
    process.stderr.write(
      `[sharepoint-mcp] app-only credentials verified.\n`,
    );
  }
}

function printHelp() {
  process.stderr.write(
    `sharepoint-mcp\n\n` +
      `Usage:\n` +
      `  sharepoint-mcp           Run the MCP server on stdio (default).\n` +
      `  sharepoint-mcp serve     Same as default.\n` +
      `  sharepoint-mcp auth      Interactive sign-in; caches the token for the server.\n` +
      `  sharepoint-mcp --help    Show this message.\n`,
  );
}

const cmd = process.argv[2];

if (cmd === "--help" || cmd === "-h" || cmd === "help") {
  printHelp();
} else if (cmd === "auth") {
  runAuth().catch((err) => {
    process.stderr.write(
      `[sharepoint-mcp] auth failed: ${err?.stack ?? err}\n`,
    );
    process.exit(1);
  });
} else if (cmd === undefined || cmd === "serve") {
  runServer().catch((err) => {
    process.stderr.write(`[sharepoint-mcp] fatal: ${err?.stack ?? err}\n`);
    process.exit(1);
  });
} else {
  process.stderr.write(`[sharepoint-mcp] unknown command: ${cmd}\n`);
  printHelp();
  process.exit(2);
}
