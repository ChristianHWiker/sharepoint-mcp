# SharePoint MCP

A [Model Context Protocol](https://modelcontextprotocol.io) server for **SharePoint Online**. It lets an MCP client (Claude Code, Claude Desktop, or any MCP host) work with a SharePoint tenant in natural language: browse sites, read and write files, manage lists and items, author modern pages with native web parts, run searches, and manage permissions.

Built on the Microsoft Graph API, with a SharePoint REST fallback for the parts Graph can't do (see below). Node + TypeScript, delegated or app-only auth.

> Not affiliated with or endorsed by Microsoft.

## Highlight: real modern-page web parts

Most tools that create SharePoint pages over Graph can only place text and simple parts — **Graph's Pages API cannot author web parts whose properties contain arrays of objects** (e.g. Quick Links). It rejects them at the OData layer regardless of payload shape.

This server sidesteps that by building the page's `CanvasContent1` and posting it through the SharePoint REST pages API (`_api/sitepages/pages` + `SavePageAsDraft`) — the same approach PnP uses. Each web part's default `properties` are seeded dynamically from the site's web-part manifest (`GetClientSideWebParts`), so you get correct, current schemas without hand-maintaining them. Supported page web parts: `text`, `image`, `divider`, `button`, `quickLinks`, and `standard` (any web part by GUID + data).

## Features

| Area | Tools (examples) |
| --- | --- |
| **Sites** | get root site, list/resolve sites, subsites, site drives & columns, follow/unfollow, create site, invite users |
| **Files & drives** | list/get/upload/download/copy/move/delete items, folders, check-in/out, file versions, sharing links |
| **Lists & items** | list/create/delete lists, query/create/update/delete items, columns, views, item versions, item permissions |
| **Pages** | list/get/create/update/publish/delete pages, set thumbnail — with native web parts |
| **Content types** | list/get/create content types, add columns, site content types & columns |
| **Taxonomy** | term store groups, term sets, terms |
| **Search** | tenant search across sites and content |
| **Permissions** | list/revoke permissions, create sharing links |
| **Subscriptions** | create/list/update/delete change webhooks |
| **Recycle bin** | list/restore/permanently delete items |

## Quick start

```bash
git clone <repo-url> sharepointMCP
cd sharepointMCP
npm install
npm run build
cp .env.example .env   # then fill in AZURE_TENANT_ID and AZURE_CLIENT_ID
npm run auth           # one-time device-code sign-in (delegated mode)
```

Register with Claude Code:

```bash
claude mcp add sharepoint -- node /full/path/to/sharepointMCP/dist/index.js
```

Full instructions, app-registration steps, and troubleshooting are in [SETUP.md](./SETUP.md).

## Authentication

- **Delegated (default):** set `AZURE_TENANT_ID` and `AZURE_CLIENT_ID`; sign in as yourself via device code. The token is cached under `%APPDATA%\sharepoint-mcp` and refreshed automatically. Requires *Allow public client flows* on the app registration.
- **App-only:** additionally set `AZURE_CLIENT_SECRET` to use client-credentials flow with application permissions.

Required Graph permissions: `Sites.ReadWrite.All`, `Sites.Manage.All`, `Files.ReadWrite.All`, `User.Read` (delegated) or their application equivalents.

## Development

```bash
npm run dev     # run from source with tsx
npm run build   # compile to dist/
```

`scripts/verify-e2e.mts` is a live smoke test for page authoring (create → read → update → delete across every web-part kind). Set `TEST_SITE_ID` to a `host,siteGuid,webGuid` site id and run `npx tsx scripts/verify-e2e.mts`.

## Notes & limitations

- The `image` web part is best-effort for externally addressable URLs. A **site-hosted** image needs `siteId`/`webId`/`listId`/`uniqueId` metadata — pass a captured blob via the `standard` web-part kind for those.
- `update_page` checks out the page before saving (SharePoint requires an active checkout to edit).

## License

ISC
