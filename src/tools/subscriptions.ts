import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { graphClient } from "../graph.js";
import { asJson, asText, wrapTool } from "../util.js";

export function registerSubscriptionTools(server: McpServer) {
  server.registerTool(
    "list_subscriptions",
    {
      title: "List webhook subscriptions",
      description: "List all active webhook change subscriptions created by this app registration.",
      inputSchema: {},
    },
    wrapTool(async () => {
      const client = graphClient();
      const res = await client.api("/subscriptions").get();
      return asJson(res.value);
    }),
  );

  server.registerTool(
    "create_subscription",
    {
      title: "Create webhook subscription",
      description:
        "Create a webhook change subscription. The notificationUrl must be publicly accessible and quickly respond to verification requests.",
      inputSchema: {
        changeType: z
          .enum(["updated", "deleted", "added"])
          .describe("SharePoint lists typically only support 'updated'."),
        notificationUrl: z
          .string()
          .describe("The HTTPS URL that will receive the notification payloads."),
        resource: z
          .string()
          .describe("The target resource path, e.g. 'sites/site-id/lists/list-id'."),
        expirationDateTime: z
          .string()
          .describe("ISO-8601 string when the subscription expires, e.g. '2026-06-01T11:00:00Z'."),
        clientState: z
          .string()
          .optional()
          .describe("A secret value sent back in notifications to verify authenticity."),
      },
    },
    wrapTool(
      async (args: {
        changeType: string;
        notificationUrl: string;
        resource: string;
        expirationDateTime: string;
        clientState?: string;
      }) => {
        const { changeType, notificationUrl, resource, expirationDateTime, clientState } = args;
        const client = graphClient();
        const body: Record<string, unknown> = {
          changeType,
          notificationUrl,
          resource,
          expirationDateTime,
        };
        if (clientState) body.clientState = clientState;

        const res = await client.api("/subscriptions").post(body);
        return asJson(res);
      },
    ),
  );

  server.registerTool(
    "update_subscription",
    {
      title: "Update subscription expiration",
      description: "Extend the lifetime of a webhook subscription.",
      inputSchema: {
        subscriptionId: z.string(),
        expirationDateTime: z
          .string()
          .describe("The new expiration timestamp (ISO-8601 string)."),
      },
    },
    wrapTool(
      async ({
        subscriptionId,
        expirationDateTime,
      }: {
        subscriptionId: string;
        expirationDateTime: string;
      }) => {
        const client = graphClient();
        const res = await client
          .api(`/subscriptions/${subscriptionId}`)
          .patch({ expirationDateTime });
        return asJson(res);
      },
    ),
  );

  server.registerTool(
    "delete_subscription",
    {
      title: "Delete webhook subscription",
      description: "Unsubscribe from webhook change notifications.",
      inputSchema: {
        subscriptionId: z.string(),
      },
    },
    wrapTool(async ({ subscriptionId }: { subscriptionId: string }) => {
      const client = graphClient();
      await client.api(`/subscriptions/${subscriptionId}`).delete();
      return asText("Subscription deleted.");
    }),
  );
}
