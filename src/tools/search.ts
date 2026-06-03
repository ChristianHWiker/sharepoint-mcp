import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { graphClient } from "../graph.js";
import { asJson, wrapTool } from "../util.js";

const EntityType = z.enum([
  "site",
  "list",
  "listItem",
  "drive",
  "driveItem",
]);

export function registerSearchTools(server: McpServer) {
  server.registerTool(
    "search",
    {
      title: "Microsoft Search across SharePoint",
      description:
        "Run a Microsoft Search query across the tenant. Choose what entity types to search (sites, listItems, driveItems, etc).",
      inputSchema: {
        query: z.string().describe("KQL/free-text query string."),
        entityTypes: z
          .array(EntityType)
          .optional()
          .describe("Default: ['driveItem', 'listItem', 'site']."),
        from: z.number().int().min(0).optional(),
        size: z.number().int().min(1).max(500).optional(),
      },
    },
    wrapTool(async (args: {
      query: string;
      entityTypes?: Array<z.infer<typeof EntityType>>;
      from?: number;
      size?: number;
    }) => {
      const { query, entityTypes, from, size } = args;
      const client = graphClient();
      const body = {
        requests: [
          {
            entityTypes: entityTypes ?? ["driveItem", "listItem", "site"],
            query: { queryString: query },
            from: from ?? 0,
            size: size ?? 25,
          },
        ],
      };
      const res = await client.api("/search/query").post(body);
      return asJson(res);
    }),
  );
}
