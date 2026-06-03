import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { graphClient } from "../graph.js";
import { asJson, asText, wrapTool } from "../util.js";

export function registerContentTypeTools(server: McpServer) {
  server.registerTool(
    "list_content_types",
    {
      title: "List content types",
      description: "List all content types defined on a SharePoint site.",
      inputSchema: {
        siteId: z.string(),
      },
    },
    wrapTool(async ({ siteId }: { siteId: string }) => {
      const client = graphClient();
      const res = await client.api(`/sites/${siteId}/contentTypes`).get();
      return asJson(res.value);
    }),
  );

  server.registerTool(
    "get_content_type",
    {
      title: "Get content type",
      description: "Retrieve a specific content type with its column schema (expanded fields).",
      inputSchema: {
        siteId: z.string(),
        contentTypeId: z.string(),
      },
    },
    wrapTool(async ({ siteId, contentTypeId }: { siteId: string; contentTypeId: string }) => {
      const client = graphClient();
      const res = await client
        .api(`/sites/${siteId}/contentTypes/${contentTypeId}`)
        .expand("columns")
        .get();
      return asJson(res);
    }),
  );

  server.registerTool(
    "create_site_content_type",
    {
      title: "Create site content type",
      description: "Create a new content type at the site level.",
      inputSchema: {
        siteId: z.string(),
        name: z.string().describe("The name of the custom content type."),
        description: z.string().optional(),
        parentId: z
          .string()
          .optional()
          .describe("The parent content type ID. Defaults to base Document '0x01' if omitted."),
      },
    },
    wrapTool(
      async (args: {
        siteId: string;
        name: string;
        description?: string;
        parentId?: string;
      }) => {
        const { siteId, name, description, parentId } = args;
        const client = graphClient();
        const body = {
          name,
          description: description ?? "",
          base: {
            id: parentId ?? "0x01",
          },
        };
        const res = await client.api(`/sites/${siteId}/contentTypes`).post(body);
        return asJson(res);
      },
    ),
  );

  server.registerTool(
    "add_content_type_column",
    {
      title: "Add column to content type",
      description: "Add a column definition reference to a content type.",
      inputSchema: {
        siteId: z.string(),
        contentTypeId: z.string(),
        columnId: z.string().describe("The ID of the site column to associate."),
      },
    },
    wrapTool(
      async (args: { siteId: string; contentTypeId: string; columnId: string }) => {
        const { siteId, contentTypeId, columnId } = args;
        const client = graphClient();
        const res = await client
          .api(`/sites/${siteId}/contentTypes/${contentTypeId}/columns`)
          .post({
            id: columnId,
          });
        return asJson(res);
      },
    ),
  );

  server.registerTool(
    "add_content_type_to_list",
    {
      title: "Add content type to list",
      description: "Associate an existing site content type with a specific list by copy.",
      inputSchema: {
        siteId: z.string(),
        listId: z.string(),
        contentTypeId: z.string().describe("The ID of the site content type to copy/add."),
      },
    },
    wrapTool(
      async (args: { siteId: string; listId: string; contentTypeId: string }) => {
        const { siteId, listId, contentTypeId } = args;
        const client = graphClient();
        const res = await client
          .api(`/sites/${siteId}/lists/${listId}/contentTypes/addCopy`)
          .post({
            contentType: `https://graph.microsoft.com/v1.0/sites/${siteId}/contentTypes/${contentTypeId}`,
          });
        return asJson(res);
      },
    ),
  );
}
