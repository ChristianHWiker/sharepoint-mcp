import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import fs from "node:fs";
import path from "node:path";
import { graphClient } from "../graph.js";
import { getAccessToken } from "../auth.js";
import { asJson, asText, wrapTool } from "../util.js";

const SMALL_UPLOAD_LIMIT = 4 * 1024 * 1024;
const UPLOAD_CHUNK_SIZE = 5 * 1024 * 1024;

function cleanPath(p: string): string {
  return p.replace(/^\/+/, "").replace(/\/+$/, "");
}

async function uploadSession(args: {
  siteId: string;
  targetPath: string;
  data: Buffer;
}): Promise<unknown> {
  const { siteId, targetPath, data } = args;
  const client = graphClient();
  const session = await client
    .api(`/sites/${siteId}/drive/root:/${cleanPath(targetPath)}:/createUploadSession`)
    .post({
      item: { "@microsoft.graph.conflictBehavior": "replace" },
    });
  const uploadUrl: string = session.uploadUrl;

  const token = await getAccessToken();
  const total = data.length;
  let offset = 0;
  let finalResponse: unknown = null;

  while (offset < total) {
    const end = Math.min(offset + UPLOAD_CHUNK_SIZE, total);
    const chunk = data.subarray(offset, end);
    const res = await fetch(uploadUrl, {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Length": String(chunk.length),
        "Content-Range": `bytes ${offset}-${end - 1}/${total}`,
      },
      body: new Uint8Array(chunk),
    });
    if (!res.ok && res.status !== 202) {
      const text = await res.text();
      throw new Error(`Upload chunk failed (${res.status}): ${text}`);
    }
    if (res.status === 200 || res.status === 201) {
      finalResponse = await res.json();
    }
    offset = end;
  }

  return finalResponse;
}

export function registerFileTools(server: McpServer) {
  server.registerTool(
    "upload_file",
    {
      title: "Upload a file to a site's default library",
      description:
        "Upload a local file to the site's default document library. Automatically uses an upload session for files >4MB.",
      inputSchema: {
        siteId: z.string(),
        targetPath: z
          .string()
          .describe(
            "Path within the library, e.g. 'images/banner.png'. Parent folders are created if missing.",
          ),
        localPath: z.string().describe("Absolute path to a local file."),
      },
    },
    wrapTool(async ({
      siteId,
      targetPath,
      localPath,
    }: { siteId: string; targetPath: string; localPath: string }) => {
      if (!fs.existsSync(localPath)) {
        throw new Error(`Local file not found: ${localPath}`);
      }
      const data = fs.readFileSync(localPath);
      if (data.length > SMALL_UPLOAD_LIMIT) {
        const result = await uploadSession({ siteId, targetPath, data });
        return asJson(result);
      }
      const client = graphClient();
      const res = await client
        .api(`/sites/${siteId}/drive/root:/${cleanPath(targetPath)}:/content`)
        .put(data);
      return asJson(res);
    }),
  );

  server.registerTool(
    "list_drive_items",
    {
      title: "List items in a site's default library folder",
      description:
        "List files and folders in a site's default document library, optionally under a sub-folder.",
      inputSchema: {
        siteId: z.string(),
        folderPath: z
          .string()
          .optional()
          .describe("Optional folder path, e.g. 'images'. Omit for root."),
      },
    },
    wrapTool(async ({
      siteId,
      folderPath,
    }: { siteId: string; folderPath?: string }) => {
      const client = graphClient();
      const endpoint = folderPath
        ? `/sites/${siteId}/drive/root:/${cleanPath(folderPath)}:/children`
        : `/sites/${siteId}/drive/root/children`;
      const res = await client
        .api(endpoint)
        .select("id,name,size,webUrl,folder,file")
        .get();
      return asJson(res.value);
    }),
  );

  server.registerTool(
    "get_drive_item",
    {
      title: "Get a drive item by path",
      description: "Get metadata for a file or folder in the site's default library.",
      inputSchema: {
        siteId: z.string(),
        itemPath: z.string().describe("Path within the library, e.g. 'docs/report.docx'"),
      },
    },
    wrapTool(async ({
      siteId,
      itemPath,
    }: { siteId: string; itemPath: string }) => {
      const client = graphClient();
      const res = await client
        .api(`/sites/${siteId}/drive/root:/${cleanPath(itemPath)}`)
        .get();
      return asJson(res);
    }),
  );

  server.registerTool(
    "create_folder",
    {
      title: "Create a folder",
      description:
        "Create a folder in the site's default library. Parent folders are created if missing.",
      inputSchema: {
        siteId: z.string(),
        folderPath: z.string().describe("Folder path, e.g. 'projects/2026/q1'"),
      },
    },
    wrapTool(async ({
      siteId,
      folderPath,
    }: { siteId: string; folderPath: string }) => {
      const cleaned = cleanPath(folderPath);
      const parent = path.posix.dirname(cleaned);
      const name = path.posix.basename(cleaned);
      const client = graphClient();
      const endpoint =
        parent === "." || parent === ""
          ? `/sites/${siteId}/drive/root/children`
          : `/sites/${siteId}/drive/root:/${parent}:/children`;
      const res = await client.api(endpoint).post({
        name,
        folder: {},
        "@microsoft.graph.conflictBehavior": "fail",
      });
      return asJson(res);
    }),
  );

  server.registerTool(
    "delete_drive_item",
    {
      title: "Delete a file or folder",
      description: "Permanently delete a file or folder from the site's default library.",
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
      await client
        .api(`/sites/${siteId}/drive/root:/${cleanPath(itemPath)}`)
        .delete();
      return asText("Deleted.");
    }),
  );

  server.registerTool(
    "move_drive_item",
    {
      title: "Move or rename a file or folder",
      description:
        "Move a drive item to a new parent folder and/or rename it. Source and destination must be in the same library.",
      inputSchema: {
        siteId: z.string(),
        itemPath: z.string().describe("Current path of the item."),
        newParentPath: z
          .string()
          .optional()
          .describe("New parent folder path. Omit to keep current parent."),
        newName: z.string().optional().describe("New name. Omit to keep current name."),
      },
    },
    wrapTool(async (args: {
      siteId: string;
      itemPath: string;
      newParentPath?: string;
      newName?: string;
    }) => {
      const { siteId, itemPath, newParentPath, newName } = args;
      const client = graphClient();
      const body: Record<string, unknown> = {};
      if (newName) body.name = newName;
      if (newParentPath !== undefined) {
        const parent = cleanPath(newParentPath);
        const parentRef = parent === ""
          ? await client.api(`/sites/${siteId}/drive/root`).select("id").get()
          : await client
              .api(`/sites/${siteId}/drive/root:/${parent}`)
              .select("id")
              .get();
        body.parentReference = { id: parentRef.id };
      }
      const res = await client
        .api(`/sites/${siteId}/drive/root:/${cleanPath(itemPath)}`)
        .patch(body);
      return asJson(res);
    }),
  );

  server.registerTool(
    "copy_drive_item",
    {
      title: "Copy a file or folder",
      description:
        "Copy a drive item. Graph returns a monitor URL; the operation may complete asynchronously.",
      inputSchema: {
        siteId: z.string(),
        itemPath: z.string(),
        newParentPath: z
          .string()
          .optional()
          .describe("Destination parent folder. Omit to copy into the same parent."),
        newName: z.string().optional().describe("Optional new name for the copy."),
      },
    },
    wrapTool(async (args: {
      siteId: string;
      itemPath: string;
      newParentPath?: string;
      newName?: string;
    }) => {
      const { siteId, itemPath, newParentPath, newName } = args;
      const client = graphClient();
      const body: Record<string, unknown> = {};
      if (newName) body.name = newName;
      if (newParentPath !== undefined) {
        const parent = cleanPath(newParentPath);
        const parentRef = parent === ""
          ? await client.api(`/sites/${siteId}/drive/root`).select("id,driveId").get()
          : await client
              .api(`/sites/${siteId}/drive/root:/${parent}`)
              .select("id,driveId")
              .get();
        body.parentReference = { id: parentRef.id, driveId: parentRef.parentReference?.driveId };
      }
      const res = await client
        .api(`/sites/${siteId}/drive/root:/${cleanPath(itemPath)}:/copy`)
        .post(body);
      return asJson(res ?? { status: "accepted" });
    }),
  );

  server.registerTool(
    "download_file",
    {
      title: "Download a file to disk",
      description:
        "Download a file from the site's default library and write it to a local path.",
      inputSchema: {
        siteId: z.string(),
        itemPath: z.string(),
        localPath: z.string().describe("Absolute local path to write the file to."),
      },
    },
    wrapTool(async ({
      siteId,
      itemPath,
      localPath,
    }: { siteId: string; itemPath: string; localPath: string }) => {
      const client = graphClient();
      const stream: NodeJS.ReadableStream = await client
        .api(`/sites/${siteId}/drive/root:/${cleanPath(itemPath)}:/content`)
        .getStream();
      await new Promise<void>((resolve, reject) => {
        const out = fs.createWriteStream(localPath);
        stream.pipe(out);
        out.on("finish", () => resolve());
        out.on("error", reject);
        stream.on("error", reject);
      });
      const stat = fs.statSync(localPath);
      return asJson({ localPath, bytes: stat.size });
    }),
  );

  server.registerTool(
    "checkout_file",
    {
      title: "Check out a file",
      description:
        "Check out a file in the site's default document library to prevent others from editing it.",
      inputSchema: {
        siteId: z.string(),
        itemPath: z.string().describe("Path of the file to check out, e.g. 'documents/report.docx'"),
      },
    },
    wrapTool(async ({ siteId, itemPath }: { siteId: string; itemPath: string }) => {
      const client = graphClient();
      await client
        .api(`/sites/${siteId}/drive/root:/${cleanPath(itemPath)}:/checkout`)
        .post({});
      return asText("File checked out.");
    }),
  );

  server.registerTool(
    "checkin_file",
    {
      title: "Check in a file",
      description:
        "Check in a checked-out file to the site's default document library, making the changes visible to others.",
      inputSchema: {
        siteId: z.string(),
        itemPath: z.string().describe("Path of the file to check in, e.g. 'documents/report.docx'"),
        comment: z.string().optional().describe("Check-in comment describing the changes."),
      },
    },
    wrapTool(
      async ({
        siteId,
        itemPath,
        comment,
      }: {
        siteId: string;
        itemPath: string;
        comment?: string;
      }) => {
        const client = graphClient();
        await client
          .api(`/sites/${siteId}/drive/root:/${cleanPath(itemPath)}:/checkin`)
          .post({ comment: comment ?? "" });
        return asText("File checked in.");
      },
    ),
  );

  server.registerTool(
    "discard_file_checkout",
    {
      title: "Discard file checkout",
      description:
        "Discard a file checkout on the site's default library, discarding any changes made since the checkout.",
      inputSchema: {
        siteId: z.string(),
        itemPath: z.string().describe("Path of the file, e.g. 'documents/report.docx'"),
      },
    },
    wrapTool(async ({ siteId, itemPath }: { siteId: string; itemPath: string }) => {
      const client = graphClient();
      await client
        .api(`/sites/${siteId}/drive/root:/${cleanPath(itemPath)}:/discardCheckout`)
        .post({});
      return asText("Checkout discarded.");
    }),
  );

  server.registerTool(
    "list_file_versions",
    {
      title: "List file versions",
      description:
        "List all versions of a file in the site's default document library.",
      inputSchema: {
        siteId: z.string(),
        itemPath: z.string().describe("Path of the file, e.g. 'documents/report.docx'"),
      },
    },
    wrapTool(async ({ siteId, itemPath }: { siteId: string; itemPath: string }) => {
      const client = graphClient();
      const res = await client
        .api(`/sites/${siteId}/drive/root:/${cleanPath(itemPath)}:/versions`)
        .get();
      return asJson(res.value);
    }),
  );

  server.registerTool(
    "restore_file_version",
    {
      title: "Restore a file version",
      description:
        "Restore a specific past version of a file to be the current version in the site's default document library.",
      inputSchema: {
        siteId: z.string(),
        itemPath: z.string().describe("Path of the file, e.g. 'documents/report.docx'"),
        versionId: z.string().describe("The ID of the version to restore."),
      },
    },
    wrapTool(
      async ({
        siteId,
        itemPath,
        versionId,
      }: {
        siteId: string;
        itemPath: string;
        versionId: string;
      }) => {
        const client = graphClient();
        await client
          .api(
            `/sites/${siteId}/drive/root:/${cleanPath(itemPath)}:/versions/${versionId}/restoreVersion`,
          )
          .post({});
        return asText(`Version ${versionId} restored.`);
      },
    ),
  );
}
