import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { graphClient } from "../graph.js";
import { asJson, asText, wrapTool } from "../util.js";

const LinkType = z.enum(["view", "edit", "embed"]);
const LinkScope = z.enum(["anonymous", "organization", "users"]);

export function registerPermissionTools(server: McpServer) {
  server.registerTool(
    "list_item_permissions",
    {
      title: "List permissions on a drive item",
      description:
        "List permissions (sharing links, direct grants) on a file or folder in the site's default library.",
      inputSchema: {
        siteId: z.string(),
        itemPath: z.string(),
      },
    },
    wrapTool(async ({
      siteId,
      itemPath,
    }: { siteId: string; itemPath: string }) => {
      const client = graphClient();
      const cleaned = itemPath.replace(/^\/+/, "");
      const res = await client
        .api(`/sites/${siteId}/drive/root:/${cleaned}:/permissions`)
        .get();
      return asJson(res.value);
    }),
  );

  server.registerTool(
    "create_sharing_link",
    {
      title: "Create a sharing link for a drive item",
      description:
        "Create a sharing link (view/edit/embed) at the chosen scope (anonymous/organization/users).",
      inputSchema: {
        siteId: z.string(),
        itemPath: z.string(),
        type: LinkType,
        scope: LinkScope.optional().describe("Default: organization."),
        password: z.string().optional(),
        expirationDateTime: z
          .string()
          .optional()
          .describe("ISO-8601 datetime, e.g. '2026-06-01T00:00:00Z'"),
      },
    },
    wrapTool(async (args: {
      siteId: string;
      itemPath: string;
      type: z.infer<typeof LinkType>;
      scope?: z.infer<typeof LinkScope>;
      password?: string;
      expirationDateTime?: string;
    }) => {
      const { siteId, itemPath, type, scope, password, expirationDateTime } = args;
      const client = graphClient();
      const cleaned = itemPath.replace(/^\/+/, "");
      const body: Record<string, unknown> = {
        type,
        scope: scope ?? "organization",
      };
      if (password) body.password = password;
      if (expirationDateTime) body.expirationDateTime = expirationDateTime;
      const res = await client
        .api(`/sites/${siteId}/drive/root:/${cleaned}:/createLink`)
        .post(body);
      return asJson(res);
    }),
  );

  server.registerTool(
    "invite_users",
    {
      title: "Invite users to a drive item",
      description:
        "Send a sharing invitation to one or more email addresses with a specified role.",
      inputSchema: {
        siteId: z.string(),
        itemPath: z.string(),
        emails: z.array(z.string()).min(1),
        roles: z
          .array(z.enum(["read", "write", "owner"]))
          .min(1)
          .describe("Roles to grant, e.g. ['read'] or ['write']."),
        message: z.string().optional(),
        requireSignIn: z.boolean().optional().describe("Default true."),
        sendInvitation: z.boolean().optional().describe("Default true."),
      },
    },
    wrapTool(async (args: {
      siteId: string;
      itemPath: string;
      emails: string[];
      roles: Array<"read" | "write" | "owner">;
      message?: string;
      requireSignIn?: boolean;
      sendInvitation?: boolean;
    }) => {
      const { siteId, itemPath, emails, roles, message, requireSignIn, sendInvitation } = args;
      const client = graphClient();
      const cleaned = itemPath.replace(/^\/+/, "");
      const res = await client
        .api(`/sites/${siteId}/drive/root:/${cleaned}:/invite`)
        .post({
          recipients: emails.map((email) => ({ email })),
          message: message ?? "",
          requireSignIn: requireSignIn ?? true,
          sendInvitation: sendInvitation ?? true,
          roles,
        });
      return asJson(res);
    }),
  );

  server.registerTool(
    "revoke_permission",
    {
      title: "Revoke a permission",
      description: "Remove a specific permission from a drive item by permission id.",
      inputSchema: {
        siteId: z.string(),
        itemPath: z.string(),
        permissionId: z.string(),
      },
    },
    wrapTool(async ({
      siteId,
      itemPath,
      permissionId,
    }: { siteId: string; itemPath: string; permissionId: string }) => {
      const client = graphClient();
      const cleaned = itemPath.replace(/^\/+/, "");
      await client
        .api(`/sites/${siteId}/drive/root:/${cleaned}:/permissions/${permissionId}`)
        .delete();
      return asText("Revoked.");
    }),
  );
}
