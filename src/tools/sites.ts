import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { graphClient } from "../graph.js";
import { asJson, asText, wrapTool } from "../util.js";

export function registerSiteTools(server: McpServer) {
  server.registerTool(
    "list_sites",
    {
      title: "Search SharePoint sites",
      description:
        "Search SharePoint sites by keyword. Returns id, name, webUrl for each match.",
      inputSchema: {
        query: z
          .string()
          .describe("Search text, e.g. 'marketing'. Use '*' to list all."),
      },
    },
    wrapTool(async ({ query }: { query: string }) => {
      const client = graphClient();
      const res = await client
        .api("/sites")
        .query({ search: query })
        .select("id,name,displayName,webUrl")
        .get();
      return asJson(res.value);
    }),
  );

  server.registerTool(
    "get_site_by_path",
    {
      title: "Get site by URL path",
      description:
        "Resolve a SharePoint site by its hostname and server-relative path (e.g. hostname='contoso.sharepoint.com', sitePath='/sites/marketing').",
      inputSchema: {
        hostname: z
          .string()
          .describe("Tenant hostname, e.g. 'contoso.sharepoint.com'"),
        sitePath: z
          .string()
          .describe("Server-relative path, e.g. '/sites/marketing'"),
      },
    },
    wrapTool(async ({ hostname, sitePath }: { hostname: string; sitePath: string }) => {
      const client = graphClient();
      const normalized = sitePath.startsWith("/") ? sitePath : `/${sitePath}`;
      const res = await client
        .api(`/sites/${hostname}:${normalized}`)
        .select("id,name,displayName,webUrl")
        .get();
      return asJson(res);
    }),
  );

  server.registerTool(
    "get_root_site",
    {
      title: "Get tenant root site",
      description: "Return the tenant's root SharePoint site.",
      inputSchema: {},
    },
    wrapTool(async () => {
      const client = graphClient();
      const res = await client
        .api("/sites/root")
        .select("id,name,displayName,webUrl")
        .get();
      return asJson(res);
    }),
  );

  server.registerTool(
    "list_subsites",
    {
      title: "List subsites of a site",
      description: "List immediate subsites (child sites) of a SharePoint site.",
      inputSchema: { siteId: z.string() },
    },
    wrapTool(async ({ siteId }: { siteId: string }) => {
      const client = graphClient();
      const res = await client
        .api(`/sites/${siteId}/sites`)
        .select("id,name,displayName,webUrl")
        .get();
      return asJson(res.value);
    }),
  );

  server.registerTool(
    "get_site_drives",
    {
      title: "List document libraries on a site",
      description:
        "List all document libraries (drives) on a SharePoint site. Use the drive id with file tools.",
      inputSchema: { siteId: z.string() },
    },
    wrapTool(async ({ siteId }: { siteId: string }) => {
      const client = graphClient();
      const res = await client
        .api(`/sites/${siteId}/drives`)
        .select("id,name,description,driveType,webUrl")
        .get();
      return asJson(res.value);
    }),
  );

  server.registerTool(
    "get_site_columns",
    {
      title: "List site columns",
      description: "List site-level columns (fields) defined on a SharePoint site.",
      inputSchema: { siteId: z.string() },
    },
    wrapTool(async ({ siteId }: { siteId: string }) => {
      const client = graphClient();
      const res = await client.api(`/sites/${siteId}/columns`).get();
      return asJson(res.value);
    }),
  );

  server.registerTool(
    "follow_site",
    {
      title: "Follow a site",
      description:
        "Add the current user as a follower of this site (delegated auth only — has no effect with app-only).",
      inputSchema: { siteId: z.string() },
    },
    wrapTool(async ({ siteId }: { siteId: string }) => {
      const client = graphClient();
      await client.api("/me/followedSites/add").post({
        value: [{ id: siteId }],
      });
      return asText("Followed.");
    }),
  );

  server.registerTool(
    "unfollow_site",
    {
      title: "Unfollow a site",
      description:
        "Remove this site from the current user's followed sites (delegated auth only).",
      inputSchema: { siteId: z.string() },
    },
    wrapTool(async ({ siteId }: { siteId: string }) => {
      const client = graphClient();
      await client.api("/me/followedSites/remove").post({
        value: [{ id: siteId }],
      });
      return asText("Unfollowed.");
    }),
  );

  server.registerTool(
    "create_site_column",
    {
      title: "Create a site column",
      description: "Create a new column (field definition) at the site level, which can be reused across lists.",
      inputSchema: {
        siteId: z.string(),
        name: z.string().describe("Internal name of the column."),
        description: z.string().optional(),
        columnType: z
          .enum(["text", "number", "choice", "dateTime", "boolean", "currency", "lookup", "personOrGroup"])
          .describe("The type of field."),
        textDetails: z
          .object({
            allowMultipleLines: z.boolean().optional(),
            maxLength: z.number().int().optional(),
          })
          .optional(),
        numberDetails: z
          .object({
            minimum: z.number().optional(),
            maximum: z.number().optional(),
          })
          .optional(),
        choiceDetails: z
          .object({
            choices: z.array(z.string()),
            allowTextEntry: z.boolean().optional(),
            displayAs: z.enum(["dropDownMenu", "radioButtons", "checkboxes"]).optional(),
          })
          .optional(),
        dateTimeDetails: z
          .object({
            format: z.enum(["standard", "friendly"]).optional(),
          })
          .optional(),
        currencyDetails: z
          .object({
            locale: z.string().optional(),
          })
          .optional(),
        lookupDetails: z
          .object({
            listId: z.string().optional(),
            columnName: z.string().optional(),
            allowMultipleValues: z.boolean().optional(),
          })
          .optional(),
        personOrGroupDetails: z
          .object({
            allowMultipleSelection: z.boolean().optional(),
            chooseFromType: z.enum(["peopleOnly", "peopleAndGroups"]).optional(),
            displayAs: z.string().optional(),
          })
          .optional(),
      },
    },
    wrapTool(
      async (args: {
        siteId: string;
        name: string;
        description?: string;
        columnType: "text" | "number" | "choice" | "dateTime" | "boolean" | "currency" | "lookup" | "personOrGroup";
        textDetails?: { allowMultipleLines?: boolean; maxLength?: number };
        numberDetails?: { minimum?: number; maximum?: number };
        choiceDetails?: { choices: string[]; allowTextEntry?: boolean; displayAs?: "dropDownMenu" | "radioButtons" | "checkboxes" };
        dateTimeDetails?: { format?: "standard" | "friendly" };
        currencyDetails?: { locale?: string };
        lookupDetails?: { listId?: string; columnName?: string; allowMultipleValues?: boolean };
        personOrGroupDetails?: { allowMultipleSelection?: boolean; chooseFromType?: "peopleOnly" | "peopleAndGroups"; displayAs?: string };
      }) => {
        const {
          siteId,
          name,
          description,
          columnType,
          textDetails,
          numberDetails,
          choiceDetails,
          dateTimeDetails,
          currencyDetails,
          lookupDetails,
          personOrGroupDetails,
        } = args;
        const client = graphClient();
        const body: Record<string, unknown> = {
          name,
          description: description ?? "",
        };

        if (columnType === "text") body.text = textDetails ?? {};
        else if (columnType === "number") body.number = numberDetails ?? {};
        else if (columnType === "choice") body.choice = choiceDetails ?? {};
        else if (columnType === "dateTime") body.dateTime = dateTimeDetails ?? {};
        else if (columnType === "boolean") body.boolean = {};
        else if (columnType === "currency") body.currency = currencyDetails ?? {};
        else if (columnType === "lookup") body.lookup = lookupDetails ?? {};
        else if (columnType === "personOrGroup") body.personOrGroup = personOrGroupDetails ?? {};

        const res = await client.api(`/sites/${siteId}/columns`).post(body);
        return asJson(res);
      },
    ),
  );

  server.registerTool(
    "delete_site_column",
    {
      title: "Delete a site column",
      description: "Delete a column definition from a SharePoint site collection.",
      inputSchema: {
        siteId: z.string(),
        columnId: z.string(),
      },
    },
    wrapTool(async ({ siteId, columnId }: { siteId: string; columnId: string }) => {
      const client = graphClient();
      await client.api(`/sites/${siteId}/columns/${columnId}`).delete();
      return asText("Site column deleted.");
    }),
  );

  server.registerTool(
    "create_site",
    {
      title: "Create a site (Beta)",
      description:
        "Create a new SharePoint site collection (Team site or Communication site) via Microsoft Graph Beta.",
      inputSchema: {
        name: z.string().describe("The display name of the SharePoint site."),
        webUrl: z
          .string()
          .optional()
          .describe(
            "The absolute URL for the new site (e.g. 'https://tenant.sharepoint.com/sites/my-site'). Required if hostname is not specified.",
          ),
        hostname: z
          .string()
          .optional()
          .describe(
            "Tenant hostname (e.g. 'contoso.sharepoint.com') to automatically build the webUrl. Required if webUrl is not specified.",
          ),
        template: z
          .enum(["sts", "sitepagepublishing"])
          .optional()
          .describe(
            "The template type: 'sts' (Team Site) or 'sitepagepublishing' (Communication Site, default).",
          ),
        description: z.string().optional(),
      },
    },
    wrapTool(
      async (args: {
        name: string;
        webUrl?: string;
        hostname?: string;
        template?: "sts" | "sitepagepublishing";
        description?: string;
      }) => {
        const { name, webUrl, hostname, template, description } = args;
        let finalWebUrl = webUrl;
        if (!finalWebUrl) {
          if (!hostname) {
            throw new Error("Either webUrl or hostname must be specified.");
          }
          const slug = name
            .toLowerCase()
            .replace(/[^a-z0-9-]/g, "-")
            .replace(/-+/g, "-")
            .replace(/^-|-$/g, "");
          finalWebUrl = `https://${hostname}/sites/${slug}`;
        }

        const client = graphClient();
        const body = {
          name,
          webUrl: finalWebUrl,
          template: template ?? "sitepagepublishing",
          description: description ?? "",
        };

        const res = await client.api("/sites").version("beta").post(body);
        return asJson(res);
      },
    ),
  );

  server.registerTool(
    "list_recycle_bin_items",
    {
      title: "List recycle bin items (Beta)",
      description: "List deleted files, folders, lists, or items in a site's recycle bin using Graph Beta.",
      inputSchema: {
        siteId: z.string(),
      },
    },
    wrapTool(async ({ siteId }: { siteId: string }) => {
      const client = graphClient();
      const res = await client
        .api(`/sites/${siteId}/recycleBin/items`)
        .version("beta")
        .get();
      return asJson(res.value);
    }),
  );

  server.registerTool(
    "restore_recycle_bin_item",
    {
      title: "Restore recycle bin item (Beta)",
      description: "Restore a deleted item from the site's recycle bin using Graph Beta.",
      inputSchema: {
        siteId: z.string(),
        itemId: z.string(),
      },
    },
    wrapTool(async ({ siteId, itemId }: { siteId: string; itemId: string }) => {
      const client = graphClient();
      const res = await client
        .api(`/sites/${siteId}/recycleBin/items/${itemId}/restore`)
        .version("beta")
        .post({});
      return asJson(res);
    }),
  );

  server.registerTool(
    "delete_recycle_bin_item",
    {
      title: "Delete recycle bin item (Beta)",
      description: "Permanently delete (purge) an item from the site's recycle bin using Graph Beta.",
      inputSchema: {
        siteId: z.string(),
        itemId: z.string(),
      },
    },
    wrapTool(async ({ siteId, itemId }: { siteId: string; itemId: string }) => {
      const client = graphClient();
      await client
        .api(`/sites/${siteId}/recycleBin/items/${itemId}`)
        .version("beta")
        .delete();
      return asText("Permanently deleted from recycle bin.");
    }),
  );
}
