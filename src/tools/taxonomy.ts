import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { graphClient } from "../graph.js";
import { asJson, wrapTool } from "../util.js";

export function registerTaxonomyTools(server: McpServer) {
  server.registerTool(
    "list_term_store_groups",
    {
      title: "List term store groups",
      description: "List all term groups in the site term store.",
      inputSchema: {
        siteId: z.string(),
      },
    },
    wrapTool(async ({ siteId }: { siteId: string }) => {
      const client = graphClient();
      const res = await client.api(`/sites/${siteId}/termStore/groups`).get();
      return asJson(res.value);
    }),
  );

  server.registerTool(
    "list_term_sets",
    {
      title: "List term sets",
      description: "List all term sets under a specific term group.",
      inputSchema: {
        siteId: z.string(),
        groupId: z.string(),
      },
    },
    wrapTool(async ({ siteId, groupId }: { siteId: string; groupId: string }) => {
      const client = graphClient();
      const res = await client
        .api(`/sites/${siteId}/termStore/groups/${groupId}/sets`)
        .get();
      return asJson(res.value);
    }),
  );

  server.registerTool(
    "list_terms",
    {
      title: "List terms",
      description: "List all terms in a specific term set.",
      inputSchema: {
        siteId: z.string(),
        setId: z.string(),
      },
    },
    wrapTool(async ({ siteId, setId }: { siteId: string; setId: string }) => {
      const client = graphClient();
      const res = await client
        .api(`/sites/${siteId}/termStore/sets/${setId}/terms`)
        .get();
      return asJson(res.value);
    }),
  );

  server.registerTool(
    "create_term_group",
    {
      title: "Create term group",
      description: "Create a new term group in the site term store.",
      inputSchema: {
        siteId: z.string(),
        displayName: z.string().describe("The display name of the term group."),
      },
    },
    wrapTool(async ({ siteId, displayName }: { siteId: string; displayName: string }) => {
      const client = graphClient();
      const res = await client
        .api(`/sites/${siteId}/termStore/groups`)
        .post({ displayName });
      return asJson(res);
    }),
  );

  server.registerTool(
    "create_term_set",
    {
      title: "Create term set",
      description: "Create a new term set under a specific term group.",
      inputSchema: {
        siteId: z.string(),
        groupId: z.string(),
        localizedNames: z
          .array(
            z.object({
              locale: z.string().describe("E.g. 'en-US'"),
              name: z.string(),
            }),
          )
          .min(1)
          .describe("Localized names list."),
      },
    },
    wrapTool(
      async (args: {
        siteId: string;
        groupId: string;
        localizedNames: Array<{ locale: string; name: string }>;
      }) => {
        const { siteId, groupId, localizedNames } = args;
        const client = graphClient();
        const res = await client
          .api(`/sites/${siteId}/termStore/groups/${groupId}/sets`)
          .post({ localizedNames });
        return asJson(res);
      },
    ),
  );

  server.registerTool(
    "create_term",
    {
      title: "Create term",
      description: "Create a new term under a specific term set.",
      inputSchema: {
        siteId: z.string(),
        setId: z.string(),
        labels: z
          .array(
            z.object({
              locale: z.string().describe("E.g. 'en-US'"),
              name: z.string(),
              isDefault: z.boolean(),
            }),
          )
          .min(1)
          .describe("Labels for the term."),
      },
    },
    wrapTool(
      async (args: {
        siteId: string;
        setId: string;
        labels: Array<{ locale: string; name: string; isDefault: boolean }>;
      }) => {
        const { siteId, setId, labels } = args;
        const client = graphClient();
        const res = await client
          .api(`/sites/${siteId}/termStore/sets/${setId}/terms`)
          .post({ labels });
        return asJson(res);
      },
    ),
  );
}
