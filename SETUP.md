# SharePoint MCP setup

A Model Context Protocol server that lets Claude work with a SharePoint Online tenant: browse sites, read/write files, manage lists and pages (with native web parts), search, and handle permissions.

## Prerequisites

- **Node.js 22 or newer.** Verify with `node --version`. Install from https://nodejs.org if missing.
- **Git.**
- **Claude Code CLI.** https://claude.com/claude-code
- A Microsoft Entra (Azure AD) **app registration** in your tenant (see below), and an account assigned to it.

## Create an app registration

In the Azure portal → **Microsoft Entra ID → App registrations → New registration**:

1. Give it a name (e.g. `sharepoint-mcp`).
2. Under **Authentication**, enable **Allow public client flows** (required for the device-code sign-in used by delegated auth).
3. Under **API permissions**, add Microsoft Graph **delegated** permissions: `Sites.ReadWrite.All`, `Sites.Manage.All`, `Files.ReadWrite.All`, `User.Read`. Grant admin consent.
4. Copy the **Application (client) ID** and **Directory (tenant) ID**.

> App-only mode is also supported: add a client secret and the equivalent **application** permissions, then set `AZURE_CLIENT_SECRET` in `.env`.

## Install

```powershell
git clone <repo-url> sharepointMCP
cd sharepointMCP
npm install
npm run build
```

## Configure

Copy `.env.example` to `.env` and fill in your values:

```
AZURE_TENANT_ID=<your tenant id>
AZURE_CLIENT_ID=<your app registration client id>
```

That is the entire config for delegated auth. Leave `AZURE_CLIENT_SECRET` unset to sign in as yourself via device code.

## Sign in (one time)

```powershell
npm run auth
```

You will see a message like:

```
To sign in, use a web browser to open the page https://microsoft.com/devicelogin
and enter the code XXXXXXXXX to authenticate.
```

Open the URL, paste the code, and complete sign-in. The token is cached at `%APPDATA%\sharepoint-mcp\token-cache.json` and refreshed automatically — you should not need to re-run `npm run auth` unless the cache is deleted or your credentials change.

## Register with Claude Code

From any directory:

```powershell
claude mcp add sharepoint -- node C:/full/path/to/sharepointMCP/dist/index.js
```

Use the full absolute path to `dist/index.js`. Confirm it registered:

```powershell
claude mcp list
```

You should see `sharepoint: ... - Connected`.

## Verify

Start Claude Code and ask:

> Use sharepoint to get the root site.

It should return a JSON blob describing your tenant root site.

## Troubleshooting

**"Need admin approval" during sign-in.** Your account isn't assigned to the app registration, or admin consent wasn't granted for the delegated permissions. Assign the user under *Entra ID → Enterprise applications → \<your app\> → Users and groups*, and grant admin consent on the API permissions.

**`invalid_client` during sign-in.** The app registration's **Allow public client flows** toggle is off. Turn it back on.

**Sign-in works but tool calls fail with auth errors.** The token cache is stale. Delete `%APPDATA%\sharepoint-mcp\token-cache.json` and run `npm run auth` again.

**`AZURE_CLIENT_ID and AZURE_TENANT_ID must be set`.** The `.env` file is missing or the values are blank.

**MCP shows "Connected" but tools never appear in Claude Code.** Restart Claude Code so it reattaches the server.

## Updating

```powershell
git pull
npm install
npm run build
```

The token cache survives upgrades. No need to re-auth.
