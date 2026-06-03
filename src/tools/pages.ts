import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { randomUUID } from "node:crypto";
import { z } from "zod";
import { graphClient } from "../graph.js";
import { getSharePointToken } from "../auth.js";
import { asJson, asText, wrapTool } from "../util.js";

// ---------------------------------------------------------------------------
// Web part model
// ---------------------------------------------------------------------------
type Webpart =
  | { kind: "text"; html: string }
  | { kind: "image"; imageUrl: string; altText?: string; caption?: string }
  | { kind: "divider" }
  | { kind: "button"; label: string; url: string; alignment?: "left" | "center" | "right" }
  | { kind: "quickLinks"; links: Array<{ title: string; url: string; description?: string }> }
  | { kind: "standard"; webPartType: string; data: Record<string, unknown> };

type Column = { width?: number; webparts: Webpart[] };
type Section = {
  layout?: "oneColumn" | "twoColumns" | "threeColumns" | "oneThirdLeftColumn" | "oneThirdRightColumn";
  columns: Column[];
};

// Canonical web part type GUIDs (from the Graph "supported web parts" reference).
const WP = {
  button: "0f087d7f-520e-42b7-89c0-496aaf979d58",
  divider: "2161a1c6-db61-4731-b97c-3cdb303f7cbb",
  image: "d1d91016-032f-456d-98a4-721247c305e8",
  quickLinks: "c70391ea-0b10-4ee9-b2b4-006d3fcad0cd",
} as const;

const capitalize = (s: string) => s.charAt(0).toUpperCase() + s.slice(1);

// ---------------------------------------------------------------------------
// Web part manifests (per-site cache)
//
// `_api/web/GetClientSideWebParts` returns every web part's manifest, including
// the canonical default `properties` and `version` (dataVersion). We fetch this
// once per web and use it to seed each web part so we never hand-maintain
// property schemas. The hand-coded values below remain as a fallback if the
// manifest call fails or a part is missing. (This is the same technique PnPjs
// uses under the hood.)
// ---------------------------------------------------------------------------
type WpManifest = {
  title: string;
  description: string;
  dataVersion: string;
  defaultProperties: Record<string, unknown>;
};
export type WpManifests = Map<string, WpManifest>;

function resolveLoc(v: unknown): string {
  if (!v) return "";
  if (typeof v === "string") return v;
  const o = v as Record<string, string>;
  return o.default ?? o["en-US"] ?? Object.values(o)[0] ?? "";
}

// ---------------------------------------------------------------------------
// CanvasContent1 builder
//
// SharePoint stores a modern page's layout as a JSON string ("CanvasContent1"):
// a flat array of "controls", each carrying a `position` (zone = horizontal
// section, sectionIndex = column, sectionFactor = column width, controlIndex =
// order within the column). We post this via the SharePoint REST pages API
// rather than Graph, because Graph's OData reader cannot accept web parts whose
// `properties` contain arrays of objects (e.g. Quick Links `items`) — it 400s
// with "resource set without entity set" / a server-side null reference.
// REST takes the canvas as an opaque JSON string, so every native web part works.
// ---------------------------------------------------------------------------
type Position = {
  zoneIndex: number;
  sectionIndex: number;
  sectionFactor: number;
  controlIndex: number;
  layoutIndex: number;
};

const emptySpc = () => ({ htmlStrings: {}, searchablePlainTexts: {}, imageSources: {}, links: {} });

type WebPartSpec = {
  title: string;
  description: string;
  dataVersion: string;
  properties: Record<string, unknown>;
  serverProcessedContent?: Record<string, unknown>;
};

/** Assemble a web part control, seeding meta + default properties from the manifest. */
function assembleWebPart(
  webPartId: string,
  position: Position,
  manifests: WpManifests | undefined,
  spec: WebPartSpec,
): Record<string, unknown> {
  const instanceId = randomUUID();
  const m = manifests?.get(webPartId.toLowerCase());
  return {
    controlType: 3,
    id: instanceId,
    position,
    emphasis: {},
    webPartId,
    webPartData: {
      id: webPartId,
      instanceId,
      title: m?.title || spec.title,
      description: m?.description || spec.description,
      dataVersion: m?.dataVersion || spec.dataVersion,
      // manifest defaults first, then our specific overrides
      properties: { ...(m?.defaultProperties ?? {}), ...spec.properties },
      serverProcessedContent: spec.serverProcessedContent ?? emptySpc(),
      dynamicDataPaths: {},
      dynamicDataValues: {},
    },
  };
}

function buildControl(
  wp: Webpart,
  position: Position,
  manifests?: WpManifests,
): Record<string, unknown> {
  switch (wp.kind) {
    case "text":
      return { controlType: 4, id: randomUUID(), position, emphasis: {}, innerHTML: wp.html };

    case "divider":
      return assembleWebPart(WP.divider, position, manifests, {
        title: "Divider",
        description: "Show a divider between content.",
        dataVersion: "2.0",
        properties: {},
      });

    case "button":
      // The Button web part reads its label and link from serverProcessedContent,
      // NOT from properties. Putting them in properties yields an empty button.
      return assembleWebPart(WP.button, position, manifests, {
        title: "Button",
        description: "Display a button with a custom label and link.",
        dataVersion: "1.3",
        properties: {
          alignment: capitalize(wp.alignment ?? "center"),
          minimumLayoutWidth: 1,
          isDynamicWidthEnabled: true,
        },
        serverProcessedContent: {
          htmlStrings: {},
          searchablePlainTexts: { label: wp.label },
          imageSources: {},
          links: { linkUrl: wp.url },
        },
      });

    case "quickLinks":
      return assembleWebPart(WP.quickLinks, position, manifests, {
        title: "Quick links",
        description: "Add links to important documents and pages.",
        dataVersion: "2.2",
        properties: {
          items: wp.links.map((l, i) => ({
            sourceItem: { itemType: 2, fileExtension: "", progId: "" },
            thumbnailType: 2,
            id: i + 1,
            description: l.description ?? "",
            altText: "",
          })),
          isMigrated: true,
          hideWebPartWhenEmpty: true,
        },
        serverProcessedContent: {
          htmlStrings: {},
          searchablePlainTexts: Object.fromEntries(
            wp.links.map((l, i) => [`items[${i}].title`, l.title]),
          ),
          imageSources: {},
          links: Object.fromEntries(
            wp.links.map((l, i) => [`items[${i}].sourceItem.url`, l.url]),
          ),
        },
      });

    case "image":
      // Best-effort for an externally addressable image URL. Site-hosted images
      // additionally require siteId/webId/listId/uniqueId metadata; pass those
      // via the `standard` kind if you need a library-hosted image.
      return assembleWebPart(WP.image, position, manifests, {
        title: "Image",
        description: "Show an image on your page.",
        dataVersion: "1.9",
        properties: {
          imageSourceType: 2,
          alignment: "Center",
          altText: wp.altText ?? "",
          overlayText: "",
          captionText: wp.caption ?? "",
          fixAspectRatio: false,
        },
        serverProcessedContent: {
          htmlStrings: {},
          searchablePlainTexts: {},
          imageSources: { imageSource: wp.imageUrl },
          links: {},
        },
      });

    case "standard": {
      const data = wp.data as Record<string, any>;
      return assembleWebPart(wp.webPartType, position, manifests, {
        title: data.title ?? "",
        description: data.description ?? "",
        dataVersion: data.dataVersion ?? "1.0",
        properties: (data.properties as Record<string, unknown>) ?? {},
        serverProcessedContent: data.serverProcessedContent as Record<string, unknown> | undefined,
      });
    }
  }
}

export function buildCanvasContent1(sections: Section[], manifests?: WpManifests): string {
  const controls: Record<string, unknown>[] = [];
  sections.forEach((section, si) => {
    const zoneIndex = si + 1;
    section.columns.forEach((column, ci) => {
      const sectionFactor = column.width ?? 12;
      column.webparts.forEach((wp, wi) => {
        controls.push(
          buildControl(
            wp,
            { zoneIndex, sectionIndex: ci + 1, sectionFactor, controlIndex: wi + 1, layoutIndex: 1 },
            manifests,
          ),
        );
      });
    });
  });
  return JSON.stringify(controls);
}

function textCanvasContent1(html: string): string {
  return buildCanvasContent1([
    { layout: "oneColumn", columns: [{ width: 12, webparts: [{ kind: "text", html }] }] },
  ]);
}

// ---------------------------------------------------------------------------
// SharePoint REST helpers
// ---------------------------------------------------------------------------
type WebContext = { webUrl: string; hostname: string };

async function resolveWeb(siteId: string): Promise<WebContext> {
  const res = await graphClient().api(`/sites/${siteId}`).select("webUrl").get();
  const webUrl: string = res.webUrl;
  return { webUrl, hostname: new URL(webUrl).hostname };
}

async function spFetch(
  ctx: WebContext,
  method: string,
  apiPath: string,
  body?: unknown,
  extraHeaders: Record<string, string> = {},
): Promise<any> {
  const token = await getSharePointToken(ctx.hostname);
  const res = await fetch(`${ctx.webUrl}/_api/${apiPath}`, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json;odata=nometadata",
      "Content-Type": "application/json;odata=nometadata",
      ...extraHeaders,
    },
    body: body === undefined ? undefined : JSON.stringify(body),
  });
  const text = await res.text();
  if (!res.ok) {
    throw new Error(`SharePoint REST ${method} ${apiPath} failed (${res.status}): ${text.slice(0, 600)}`);
  }
  return text ? JSON.parse(text) : null;
}

const manifestCache = new Map<string, WpManifests>();

/** Fetch (and cache) the site's client-side web part manifests by GUID. */
async function getManifests(ctx: WebContext): Promise<WpManifests | undefined> {
  const cached = manifestCache.get(ctx.webUrl);
  if (cached) return cached;
  try {
    const res = await spFetch(ctx, "GET", "web/GetClientSideWebParts");
    const map: WpManifests = new Map();
    for (const p of res?.value ?? []) {
      const id = String(p.Id ?? p.id ?? "").toLowerCase();
      if (!id) continue;
      let m: any = {};
      try { m = JSON.parse(p.Manifest ?? p.manifest ?? "{}"); } catch { /* ignore */ }
      const entry = m?.preconfiguredEntries?.[0] ?? {};
      map.set(id, {
        title: resolveLoc(entry.title),
        description: resolveLoc(entry.description),
        dataVersion: m?.version ?? "1.0",
        defaultProperties: (entry.properties as Record<string, unknown>) ?? {},
      });
    }
    manifestCache.set(ctx.webUrl, map);
    return map;
  } catch {
    // Fall back to hand-coded defaults baked into buildControl.
    return undefined;
  }
}

/** Resolve the integer list-item Id of a modern page from its Graph GUID id. */
async function resolvePageIntId(ctx: WebContext, pageGuid: string): Promise<number> {
  const res = await spFetch(
    ctx,
    "GET",
    `sitepages/pages?$select=Id,UniqueId&$filter=UniqueId eq guid'${pageGuid}'`,
  );
  const page = res?.value?.[0];
  if (!page) throw new Error(`Could not resolve page ${pageGuid} via SharePoint REST.`);
  return page.Id;
}

async function getPageGraph(siteId: string, pageId: string): Promise<unknown> {
  return graphClient()
    .api(`/sites/${siteId}/pages/${pageId}/microsoft.graph.sitePage`)
    .expand("canvasLayout")
    .get();
}

// ---------------------------------------------------------------------------
// Tool input schema
// ---------------------------------------------------------------------------
const sectionsSchema = z
  .array(
    z.object({
      layout: z
        .enum(["oneColumn", "twoColumns", "threeColumns", "oneThirdLeftColumn", "oneThirdRightColumn"])
        .optional(),
      columns: z.array(
        z.object({
          width: z.number().int().min(1).max(12).optional(),
          webparts: z.array(
            z.union([
              z.object({ kind: z.literal("text"), html: z.string() }),
              z.object({
                kind: z.literal("image"),
                imageUrl: z.string(),
                altText: z.string().optional(),
                caption: z.string().optional(),
              }),
              z.object({ kind: z.literal("divider") }),
              z.object({
                kind: z.literal("button"),
                label: z.string(),
                url: z.string(),
                alignment: z.enum(["left", "center", "right"]).optional(),
              }),
              z.object({
                kind: z.literal("quickLinks"),
                links: z.array(
                  z.object({
                    title: z.string(),
                    url: z.string(),
                    description: z.string().optional(),
                  }),
                ),
              }),
              z.object({
                kind: z.literal("standard"),
                webPartType: z.string(),
                data: z.record(z.string(), z.unknown()),
              }),
            ]),
          ),
        }),
      ),
    }),
  )
  .describe(
    "Page canvas: array of horizontal sections. Each section has columns; each column has webparts. " +
      "Webpart kinds: 'text' (innerHTML rich text), 'image', 'divider', 'button', 'quickLinks', and " +
      "'standard' (raw web part: pass webPartType GUID + data). Set each column's `width` (1-12) to lay " +
      "out multi-column sections (e.g. two width-6 columns). Native web parts are authored via the " +
      "SharePoint REST pages API (CanvasContent1), so button, divider, and quickLinks all work — prefer " +
      "them over emulating with styled HTML.",
  );

// ---------------------------------------------------------------------------
// Tools
// ---------------------------------------------------------------------------
export function registerPageTools(server: McpServer) {
  server.registerTool(
    "list_pages",
    {
      title: "List pages in a site",
      description: "List modern pages in a SharePoint site.",
      inputSchema: { siteId: z.string() },
    },
    wrapTool(async ({ siteId }: { siteId: string }) => {
      const res = await graphClient()
        .api(`/sites/${siteId}/pages`)
        .select("id,name,title,webUrl,pageLayout,publishingState")
        .get();
      return asJson(res.value);
    }),
  );

  server.registerTool(
    "get_page",
    {
      title: "Get a page with content",
      description: "Get a page including its canvasLayout (web parts and their content).",
      inputSchema: { siteId: z.string(), pageId: z.string() },
    },
    wrapTool(async ({ siteId, pageId }: { siteId: string; pageId: string }) => {
      return asJson(await getPageGraph(siteId, pageId));
    }),
  );

  server.registerTool(
    "create_page",
    {
      title: "Create a SharePoint page",
      description:
        "Create a new modern page. Pass either htmlContent (simple text body) OR sections (structured " +
        "canvas with native web parts). Page is created as a draft — call publish_page to make it visible.",
      inputSchema: {
        siteId: z.string(),
        name: z.string().describe("File name, must end with .aspx (e.g. 'my-page.aspx')"),
        title: z.string(),
        htmlContent: z
          .string()
          .optional()
          .describe("Optional HTML body, rendered as a single text web part."),
        sections: sectionsSchema.optional(),
      },
    },
    wrapTool(async (args: {
      siteId: string;
      name: string;
      title: string;
      htmlContent?: string;
      sections?: Section[];
    }) => {
      const { siteId, name, title, htmlContent, sections } = args;
      const ctx = await resolveWeb(siteId);

      const created = await spFetch(ctx, "POST", "sitepages/pages", {
        Title: title,
        Name: name,
        PageLayoutType: "Article",
      });
      const intId: number = created.Id;

      let canvas: string | null = null;
      if (sections?.length) {
        canvas = buildCanvasContent1(sections, await getManifests(ctx));
      } else if (htmlContent) {
        canvas = textCanvasContent1(htmlContent);
      }

      if (canvas) {
        await spFetch(ctx, "POST", `sitepages/pages(${intId})/SavePageAsDraft`, {
          CanvasContent1: canvas,
        });
      }

      // Return the Graph view so callers get the canonical GUID id + canvasLayout.
      return asJson(await getPageGraph(siteId, created.UniqueId));
    }),
  );

  server.registerTool(
    "update_page",
    {
      title: "Update a page",
      description:
        "Update a page's title and/or replace its body. Pass htmlContent for a simple text body or " +
        "sections for a structured canvas. Either replaces the page content entirely.",
      inputSchema: {
        siteId: z.string(),
        pageId: z.string(),
        title: z.string().optional(),
        htmlContent: z.string().optional(),
        sections: sectionsSchema.optional(),
      },
    },
    wrapTool(async (args: {
      siteId: string;
      pageId: string;
      title?: string;
      htmlContent?: string;
      sections?: Section[];
    }) => {
      const { siteId, pageId, title, htmlContent, sections } = args;
      const ctx = await resolveWeb(siteId);
      const intId = await resolvePageIntId(ctx, pageId);

      let canvas: string | undefined;
      if (sections?.length) {
        canvas = buildCanvasContent1(sections, await getManifests(ctx));
      } else if (htmlContent !== undefined) {
        canvas = textCanvasContent1(htmlContent);
      }

      const draftBody: Record<string, unknown> = {};
      if (title !== undefined) draftBody.Title = title;
      if (canvas !== undefined) draftBody.CanvasContent1 = canvas;

      if (Object.keys(draftBody).length > 0) {
        // Editing an existing page requires an active checkout, otherwise
        // SavePageAsDraft 409s with "a site member ended your editing session".
        await spFetch(ctx, "POST", `sitepages/pages(${intId})/checkoutpage`);
        await spFetch(ctx, "POST", `sitepages/pages(${intId})/SavePageAsDraft`, draftBody);
      }

      return asJson(await getPageGraph(siteId, pageId));
    }),
  );

  server.registerTool(
    "publish_page",
    {
      title: "Publish a page",
      description: "Publish a draft page so readers can see it.",
      inputSchema: { siteId: z.string(), pageId: z.string() },
    },
    wrapTool(async ({ siteId, pageId }: { siteId: string; pageId: string }) => {
      await graphClient()
        .api(`/sites/${siteId}/pages/${pageId}/microsoft.graph.sitePage/publish`)
        .post("");
      return asText("Published.");
    }),
  );

  server.registerTool(
    "delete_page",
    {
      title: "Delete a page",
      description: "Permanently delete a page from a SharePoint site.",
      inputSchema: { siteId: z.string(), pageId: z.string() },
    },
    wrapTool(async ({ siteId, pageId }: { siteId: string; pageId: string }) => {
      await graphClient().api(`/sites/${siteId}/pages/${pageId}`).delete();
      return asText("Deleted.");
    }),
  );

  server.registerTool(
    "set_page_thumbnail",
    {
      title: "Set a page thumbnail (banner image)",
      description:
        "Set the title-area banner image for a page using a publicly accessible URL or a drive item.",
      inputSchema: {
        siteId: z.string(),
        pageId: z.string(),
        imageUrl: z.string().describe("URL of the image (typically from a SharePoint drive)."),
      },
    },
    wrapTool(async ({ siteId, pageId, imageUrl }: { siteId: string; pageId: string; imageUrl: string }) => {
      const res = await graphClient()
        .api(`/sites/${siteId}/pages/${pageId}/microsoft.graph.sitePage`)
        .patch({
          "@odata.type": "#microsoft.graph.sitePage",
          titleArea: { imageWebUrl: imageUrl, layout: "imageAndTitle" },
        });
      return asJson(res);
    }),
  );
}
