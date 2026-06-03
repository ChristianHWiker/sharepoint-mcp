import { GraphError } from "@microsoft/microsoft-graph-client";

type TextContent = { type: "text"; text: string };
type ToolResult = {
  content: TextContent[];
  isError?: boolean;
};

export function asJson(data: unknown): ToolResult {
  return {
    content: [{ type: "text", text: JSON.stringify(data, null, 2) }],
  };
}

export function asText(text: string): ToolResult {
  return { content: [{ type: "text", text }] };
}

function formatGraphError(err: GraphError): string {
  const parts = [
    `Graph request failed (status ${err.statusCode ?? "?"}):`,
    err.code ? `code=${err.code}` : null,
    err.message ? `message=${err.message}` : null,
    err.requestId ? `requestId=${err.requestId}` : null,
  ].filter(Boolean);
  return parts.join(" | ");
}

export function wrapTool<TArgs, TResult extends ToolResult>(
  handler: (args: TArgs) => Promise<TResult>,
): (args: TArgs) => Promise<ToolResult> {
  return async (args: TArgs) => {
    try {
      return await handler(args);
    } catch (err) {
      if (err instanceof GraphError) {
        return {
          content: [{ type: "text", text: formatGraphError(err) }],
          isError: true,
        };
      }
      const message = err instanceof Error ? err.message : String(err);
      return {
        content: [{ type: "text", text: `Error: ${message}` }],
        isError: true,
      };
    }
  };
}
