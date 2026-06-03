import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { graphClient } from "../graph.js";
import { asJson, asText, wrapTool } from "../util.js";
import { getSharePointToken } from "../auth.js";

const ListTemplate = z
  .enum(["genericList", "documentLibrary", "tasks", "survey", "links", "announcements", "contacts", "events"])
  .describe("SharePoint list template. Default: genericList.");

async function resolvePersonFields(
  client: any,
  siteId: string,
  listId: string,
  fields: Record<string, unknown>,
): Promise<Record<string, unknown>> {
  let columns: any[] = [];
  try {
    const colRes = await client.api(`/sites/${siteId}/lists/${listId}/columns`).get();
    columns = colRes.value || [];
  } catch (err) {
    return fields;
  }

  const personColumns = columns.filter((col) => col.personOrGroup);
  if (personColumns.length === 0) {
    return fields;
  }

  const updatedFields = { ...fields };

  const resolveEmail = async (email: string): Promise<number | null> => {
    try {
      const res = await client
        .api(`/sites/${siteId}/lists/users/items`)
        .expand("fields")
        .filter(`fields/EMail eq '${email}'`)
        .get();
      if (res.value?.length > 0) {
        return res.value[0].fields.id || res.value[0].id;
      }
    } catch {
      // ignore
    }

    try {
      const res = await client
        .api(`/sites/${siteId}/lists/users/items`)
        .expand("fields")
        .get();
      const match = res.value?.find(
        (item: any) => item.fields?.EMail?.toLowerCase() === email.toLowerCase(),
      );
      if (match) {
        return match.fields.id || match.id;
      }
    } catch {
      // ignore
    }
    return null;
  };

  for (const col of personColumns) {
    const colName = col.name;
    const value = fields[colName];

    if (value === undefined || value === null) {
      continue;
    }

    if (typeof value === "string" && value.includes("@")) {
      const lookupId = await resolveEmail(value);
      if (lookupId !== null) {
        delete updatedFields[colName];
        updatedFields[`${colName}LookupId`] = lookupId;
      } else {
        process.stderr.write(`[sharepoint-mcp] warning: could not resolve user email '${value}' to SharePoint Lookup ID. Leaving field as-is.\n`);
      }
    } else if (Array.isArray(value)) {
      const lookupIds: number[] = [];
      for (const item of value) {
        if (typeof item === "string" && item.includes("@")) {
          const lookupId = await resolveEmail(item);
          if (lookupId !== null) {
            lookupIds.push(lookupId);
          } else {
            process.stderr.write(`[sharepoint-mcp] warning: could not resolve user email '${item}' to SharePoint Lookup ID.\n`);
          }
        } else if (typeof item === "number") {
          lookupIds.push(item);
        }
      }
      if (lookupIds.length > 0) {
        delete updatedFields[colName];
        updatedFields[`${colName}LookupId@odata.type`] = "Collection(Edm.Int32)";
        updatedFields[`${colName}LookupId`] = lookupIds;
      }
    }
  }

  return updatedFields;
}

export function registerListTools(server: McpServer) {
  server.registerTool(
    "list_lists",
    {
      title: "List all lists on a site",
      description:
        "List all lists (including hidden ones) on a SharePoint site. Use list id with item tools.",
      inputSchema: {
        siteId: z.string(),
        includeHidden: z
          .boolean()
          .optional()
          .describe("Include system/hidden lists. Default false."),
      },
    },
    wrapTool(async ({
      siteId,
      includeHidden,
    }: { siteId: string; includeHidden?: boolean }) => {
      const client = graphClient();
      const res = await client
        .api(`/sites/${siteId}/lists`)
        .select("id,name,displayName,description,webUrl,list")
        .get();
      const items = res.value as Array<{ list?: { hidden?: boolean } }>;
      const filtered = includeHidden ? items : items.filter((l) => !l.list?.hidden);
      return asJson(filtered);
    }),
  );

  server.registerTool(
    "get_list",
    {
      title: "Get a list with its columns",
      description:
        "Get a list including its columns (field definitions). Useful before creating/updating items.",
      inputSchema: {
        siteId: z.string(),
        listId: z.string(),
      },
    },
    wrapTool(async ({
      siteId,
      listId,
    }: { siteId: string; listId: string }) => {
      const client = graphClient();
      const res = await client
        .api(`/sites/${siteId}/lists/${listId}`)
        .expand("columns")
        .get();
      return asJson(res);
    }),
  );

  server.registerTool(
    "create_list",
    {
      title: "Create a list",
      description:
        "Create a new list on a site. For document libraries, pages, etc. use the appropriate template.",
      inputSchema: {
        siteId: z.string(),
        displayName: z.string(),
        description: z.string().optional(),
        template: ListTemplate.optional(),
      },
    },
    wrapTool(async (args: {
      siteId: string;
      displayName: string;
      description?: string;
      template?: z.infer<typeof ListTemplate>;
    }) => {
      const { siteId, displayName, description, template } = args;
      const client = graphClient();
      const res = await client.api(`/sites/${siteId}/lists`).post({
        displayName,
        description: description ?? "",
        list: { template: template ?? "genericList" },
      });
      return asJson(res);
    }),
  );

  server.registerTool(
    "delete_list",
    {
      title: "Delete a list",
      description: "Permanently delete a list and all of its items.",
      inputSchema: { siteId: z.string(), listId: z.string() },
    },
    wrapTool(async ({
      siteId,
      listId,
    }: { siteId: string; listId: string }) => {
      const client = graphClient();
      await client.api(`/sites/${siteId}/lists/${listId}`).delete();
      return asText("Deleted.");
    }),
  );

  server.registerTool(
    "list_items",
    {
      title: "List items in a list",
      description:
        "List items in a SharePoint list. Supports OData $filter (e.g. \"fields/Status eq 'Open'\"), $orderby, $top. Fields are expanded by default.",
      inputSchema: {
        siteId: z.string(),
        listId: z.string(),
        filter: z.string().optional().describe("OData $filter on fields, e.g. fields/Status eq 'Open'"),
        orderby: z.string().optional(),
        top: z.number().int().min(1).max(5000).optional(),
        selectFields: z
          .array(z.string())
          .optional()
          .describe("Field names to include in the expanded fields projection."),
      },
    },
    wrapTool(async (args: {
      siteId: string;
      listId: string;
      filter?: string;
      orderby?: string;
      top?: number;
      selectFields?: string[];
    }) => {
      const { siteId, listId, filter, orderby, top, selectFields } = args;
      const client = graphClient();
      const expand = selectFields?.length
        ? `fields($select=${selectFields.join(",")})`
        : "fields";
      let req = client.api(`/sites/${siteId}/lists/${listId}/items`).expand(expand);
      if (filter) req = req.filter(filter);
      if (orderby) req = req.orderby(orderby);
      if (top) req = req.top(top);
      const res = await req.get();
      return asJson(res.value);
    }),
  );

  server.registerTool(
    "get_item",
    {
      title: "Get a list item",
      description: "Get a single list item with its fields expanded.",
      inputSchema: {
        siteId: z.string(),
        listId: z.string(),
        itemId: z.string(),
      },
    },
    wrapTool(async ({
      siteId,
      listId,
      itemId,
    }: { siteId: string; listId: string; itemId: string }) => {
      const client = graphClient();
      const res = await client
        .api(`/sites/${siteId}/lists/${listId}/items/${itemId}`)
        .expand("fields")
        .get();
      return asJson(res);
    }),
  );

  server.registerTool(
    "create_item",
    {
      title: "Create a list item",
      description:
        "Create a list item. 'fields' is a map of field internal name → value. Use get_list to discover field names.",
      inputSchema: {
        siteId: z.string(),
        listId: z.string(),
        fields: z
          .record(z.string(), z.unknown())
          .describe("Map of field internal names to values, e.g. { Title: 'Hello', Status: 'Open' }"),
      },
    },
    wrapTool(async ({
      siteId,
      listId,
      fields,
    }: { siteId: string; listId: string; fields: Record<string, unknown> }) => {
      const client = graphClient();
      const resolvedFields = await resolvePersonFields(client, siteId, listId, fields);
      const res = await client
        .api(`/sites/${siteId}/lists/${listId}/items`)
        .post({ fields: resolvedFields });
      return asJson(res);
    }),
  );

  server.registerTool(
    "update_item",
    {
      title: "Update a list item",
      description: "Update fields on a list item. Pass only the fields you want to change.",
      inputSchema: {
        siteId: z.string(),
        listId: z.string(),
        itemId: z.string(),
        fields: z.record(z.string(), z.unknown()),
      },
    },
    wrapTool(async ({
      siteId,
      listId,
      itemId,
      fields,
    }: {
      siteId: string;
      listId: string;
      itemId: string;
      fields: Record<string, unknown>;
    }) => {
      const client = graphClient();
      const resolvedFields = await resolvePersonFields(client, siteId, listId, fields);
      const res = await client
        .api(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`)
        .patch(resolvedFields);
      return asJson(res);
    }),
  );

  server.registerTool(
    "delete_item",
    {
      title: "Delete a list item",
      description: "Permanently delete a list item.",
      inputSchema: {
        siteId: z.string(),
        listId: z.string(),
        itemId: z.string(),
      },
    },
    wrapTool(async ({
      siteId,
      listId,
      itemId,
    }: { siteId: string; listId: string; itemId: string }) => {
      const client = graphClient();
      await client.api(`/sites/${siteId}/lists/${listId}/items/${itemId}`).delete();
      return asText("Deleted.");
    }),
  );

  server.registerTool(
    "list_item_versions",
    {
      title: "List list item versions",
      description: "List all version history for a specific SharePoint list item.",
      inputSchema: {
        siteId: z.string(),
        listId: z.string(),
        itemId: z.string(),
      },
    },
    wrapTool(
      async ({
        siteId,
        listId,
        itemId,
      }: {
        siteId: string;
        listId: string;
        itemId: string;
      }) => {
        const client = graphClient();
        const res = await client
          .api(`/sites/${siteId}/lists/${listId}/items/${itemId}/versions`)
          .get();
        return asJson(res.value);
      },
    ),
  );

  server.registerTool(
    "create_list_column",
    {
      title: "Create a list column",
      description: "Create a new column (field definition) in a SharePoint list.",
      inputSchema: {
        siteId: z.string(),
        listId: z.string(),
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
        listId: string;
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
          listId,
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

        const res = await client
          .api(`/sites/${siteId}/lists/${listId}/columns`)
          .post(body);
        return asJson(res);
      },
    ),
  );

  server.registerTool(
    "delete_list_column",
    {
      title: "Delete a list column",
      description: "Delete a column definition from a SharePoint list.",
      inputSchema: {
        siteId: z.string(),
        listId: z.string(),
        columnId: z.string(),
      },
    },
    wrapTool(
      async ({
        siteId,
        listId,
        columnId,
      }: {
        siteId: string;
        listId: string;
        columnId: string;
      }) => {
        const client = graphClient();
      await client
        .api(`/sites/${siteId}/lists/${listId}/columns/${columnId}`)
        .delete();
      return asText("Column deleted.");
    }),
  );

  server.registerTool(
    "list_views",
    {
      title: "List list views",
      description: "List all view configurations defined on a SharePoint list (SharePoint REST API).",
      inputSchema: {
        siteId: z.string(),
        listId: z.string(),
      },
    },
    wrapTool(async ({ siteId, listId }: { siteId: string; listId: string }) => {
      const res = await sharepointRestRequest({
        siteId,
        method: "GET",
        apiPath: `/_api/web/lists/getbyid('${listId}')/views`,
      });
      return asJson(res.results || res);
    }),
  );

  server.registerTool(
    "create_view",
    {
      title: "Create list view",
      description: "Create a new saved view on a SharePoint list, optionally specifying columns and sorting (SharePoint REST API).",
      inputSchema: {
        siteId: z.string(),
        listId: z.string(),
        title: z.string().describe("Display title of the view, e.g. 'Alle aktive'."),
        viewQuery: z
          .string()
          .optional()
          .describe("CAML query string for sorting/filtering, e.g. '<OrderBy><FieldRef Name=\"Created\" Ascending=\"FALSE\"/></OrderBy>'."),
        personalView: z.boolean().optional().describe("Create as a personal view. Default is false (public view)."),
        viewFields: z
          .array(z.string())
          .optional()
          .describe("List of column internal names to display in the view, e.g. ['LinkTitle', 'Ansvarlig']."),
      },
    },
    wrapTool(
      async (args: {
        siteId: string;
        listId: string;
        title: string;
        viewQuery?: string;
        personalView?: boolean;
        viewFields?: string[];
      }) => {
        const { siteId, listId, title, viewQuery, personalView, viewFields } = args;
        const createBody = {
          "__metadata": { "type": "SP.View" },
          "Title": title,
          "PersonalView": personalView ?? false,
          "ViewType": "HTML",
        } as any;
        if (viewQuery) createBody.ViewQuery = viewQuery;

        const viewRes = await sharepointRestRequest({
          siteId,
          method: "POST",
          apiPath: `/_api/web/lists/getbyid('${listId}')/views`,
          body: createBody,
        });

        const viewId = viewRes.Id;

        if (viewFields && viewFields.length > 0) {
          for (const field of viewFields) {
            await sharepointRestRequest({
              siteId,
              method: "POST",
              apiPath: `/_api/web/lists/getbyid('${listId}')/views/getbyid('${viewId}')/viewfields/addviewfield('${field}')`,
            });
          }
        }

        return asJson(viewRes);
      },
    ),
  );

  server.registerTool(
    "delete_view",
    {
      title: "Delete list view",
      description: "Permanently delete a saved view from a SharePoint list (SharePoint REST API).",
      inputSchema: {
        siteId: z.string(),
        listId: z.string(),
        viewId: z.string().describe("The ID of the view to delete."),
      },
    },
    wrapTool(
      async ({
        siteId,
        listId,
        viewId,
      }: {
        siteId: string;
        listId: string;
        viewId: string;
      }) => {
        await sharepointRestRequest({
          siteId,
          method: "DELETE",
          apiPath: `/_api/web/lists/getbyid('${listId}')/views/getbyid('${viewId}')`,
        });
        return asText("View deleted.");
      },
    ),
  );
}

async function getSiteHostnameAndUrl(siteId: string): Promise<{ hostname: string; webUrl: string }> {
  if (siteId.includes(",")) {
    const parts = siteId.split(",");
    const hostname = parts[0];
    const client = graphClient();
    const site = await client.api(`/sites/${siteId}`).select("webUrl").get();
    return { hostname, webUrl: site.webUrl };
  } else {
    const client = graphClient();
    const site = await client.api(`/sites/${siteId}`).select("webUrl").get();
    const hostname = new URL(site.webUrl).hostname;
    return { hostname, webUrl: site.webUrl };
  }
}

async function sharepointRestRequest(args: {
  siteId: string;
  method: "GET" | "POST" | "PATCH" | "DELETE";
  apiPath: string;
  body?: any;
}) {
  const { siteId, method, apiPath, body } = args;
  const { hostname, webUrl } = await getSiteHostnameAndUrl(siteId);
  const token = await getSharePointToken(hostname);

  const headers: Record<string, string> = {
    Authorization: `Bearer ${token}`,
    Accept: "application/json;odata=verbose",
  };

  if (method === "POST" || method === "PATCH" || method === "DELETE") {
    const digestRes = await fetch(`${webUrl}/_api/contextinfo`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json;odata=verbose",
      },
    });
    if (!digestRes.ok) {
      const errText = await digestRes.text();
      throw new Error(`Failed to get context info: ${errText}`);
    }
    const digestData: any = await digestRes.json();
    const digestValue = digestData?.d?.GetContextWebInformation?.FormDigestValue;
    if (!digestValue) throw new Error("Could not parse FormDigestValue");

    headers["X-RequestDigest"] = digestValue;
    headers["Content-Type"] = "application/json;odata=verbose";

    if (method === "PATCH") {
      headers["X-HTTP-Method"] = "MERGE";
      headers["IF-MATCH"] = "*";
    } else if (method === "DELETE") {
      headers["X-HTTP-Method"] = "DELETE";
      headers["IF-MATCH"] = "*";
    }
  }

  const fetchMethod = (method === "PATCH" || method === "DELETE") ? "POST" : method;
  const url = `${webUrl}${apiPath}`;

  const res = await fetch(url, {
    method: fetchMethod,
    headers,
    body: body ? JSON.stringify(body) : undefined,
  });

  if (!res.ok) {
    const errText = await res.text();
    throw new Error(`SharePoint REST request failed (${res.status}): ${errText}`);
  }

  if (res.status === 204) {
    return { success: true };
  }

  const resData: any = await res.json();
  return resData?.d || resData;
}
