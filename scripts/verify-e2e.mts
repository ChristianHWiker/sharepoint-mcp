// End-to-end test of the SHIPPED buildCanvasContent1 against live SharePoint:
// create (REST) -> get (Graph) -> update (REST) -> delete. Exercises every kind.
//
// Run: npx tsx scripts/verify-e2e.mts
import { config as loadDotenv } from "dotenv";
import path from "node:path";
import { fileURLToPath } from "node:url";

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
loadDotenv({ path: path.resolve(scriptDir, "../.env"), quiet: true });

const { buildCanvasContent1, parseCanvasContent1 } = await import("../src/tools/pages.js");
const { graphClient } = await import("../src/graph.js");
const { getSharePointToken } = await import("../src/auth.js");

// Set TEST_SITE_ID to a Graph composite site id (host,siteGuid,webGuid).
const SITE_ID = process.env.TEST_SITE_ID;
if (!SITE_ID) {
  console.error("Set TEST_SITE_ID env var to a 'host,siteGuid,webGuid' site id.");
  process.exit(1);
}

const web = await graphClient().api(`/sites/${SITE_ID}`).select("webUrl").get();
const webUrl: string = web.webUrl;
const hostname = new URL(webUrl).hostname;
const token = await getSharePointToken(hostname);

// Build the web part manifest map (the same one create_page fetches) so this
// test exercises the manifest-seeded path, not just the hand-coded fallback.
const manifestResp = await (await fetch(`${webUrl}/_api/web/GetClientSideWebParts`, {
  headers: { Authorization: `Bearer ${token}`, Accept: "application/json;odata=nometadata" },
})).json();
const manifests = new Map<string, any>();
for (const p of manifestResp.value ?? []) {
  const id = String(p.Id ?? p.id ?? "").toLowerCase();
  let m: any = {};
  try { m = JSON.parse(p.Manifest ?? "{}"); } catch {}
  const e = m?.preconfiguredEntries?.[0] ?? {};
  const loc = (v: any) => (!v ? "" : typeof v === "string" ? v : v.default ?? v["en-US"] ?? "");
  manifests.set(id, { title: loc(e.title), description: loc(e.description), dataVersion: m?.version ?? "1.0", defaultProperties: e.properties ?? {} });
}
console.log("manifests loaded:", manifests.size);
const H = {
  Authorization: `Bearer ${token}`,
  Accept: "application/json;odata=nometadata",
  "Content-Type": "application/json;odata=nometadata",
};
const api = async (m: string, p: string, b?: any, extra: any = {}) => {
  const r = await fetch(`${webUrl}/_api/${p}`, { method: m, headers: { ...H, ...extra }, body: b === undefined ? undefined : JSON.stringify(b) });
  const t = await r.text();
  if (!r.ok) throw new Error(`${m} ${p} -> ${r.status}: ${t.slice(0, 400)}`);
  return t ? JSON.parse(t) : null;
};

const sections = [
  { layout: "oneColumn", background: "neutral", columns: [{ width: 12, webparts: [{ kind: "text", html: "<h2>E2E test</h2><p>Built by the shipped builder.</p>" }, { kind: "divider" }, { kind: "spacer" }] }] },
  { layout: "twoColumns", background: "strong", columns: [
    { width: 6, webparts: [{ kind: "button", label: "Open ticket", url: "https://example.com/ticket", alignment: "left" }, { kind: "codeSnippet", code: "console.log('hi')", language: "javascript" }] },
    { width: 6, webparts: [{ kind: "quickLinks", links: [{ title: "Example", url: "https://example.com" }, { title: "Docs", url: "https://example.com/docs" }] }, { kind: "embed", embedCode: "<iframe src='https://www.youtube.com/embed/dQw4w9WgXcQ'></iframe>" }] },
  ] },
];

// CREATE
const created = await api("POST", "sitepages/pages", { Title: "zz e2e", Name: `zz-e2e-${Date.now()}.aspx`, PageLayoutType: "Article" });
await api("POST", `sitepages/pages(${created.Id})/SavePageAsDraft`, { CanvasContent1: buildCanvasContent1(sections as any, manifests) });
console.log("CREATE ok  intId:", created.Id, " guid:", created.UniqueId);

// GET via Graph (the canonical read path the MCP returns)
const page: any = await graphClient().api(`/sites/${SITE_ID}/pages/${created.UniqueId}/microsoft.graph.sitePage`).expand("canvasLayout").get();
const parts: string[] = [];
for (const s of page.canvasLayout.horizontalSections)
  for (const c of s.columns)
    for (const w of c.webparts)
      parts.push(w["@odata.type"] === "#microsoft.graph.textWebPart" ? "text" : (w.webPartType ?? "?"));
console.log("GET parts:", JSON.stringify(parts));
// confirm manifest-seeded defaults round-tripped (Quick Links should carry layoutId)
for (const s of page.canvasLayout.horizontalSections)
  for (const c of s.columns)
    for (const w of c.webparts)
      if (w.webPartType === "c70391ea-0b10-4ee9-b2b4-006d3fcad0cd")
        console.log("  QL properties keys:", Object.keys(w.data?.properties ?? {}).join(", "));

// READ-BACK: fetch raw CanvasContent1 and parse it (the get_page_canvas path).
// Confirms the round-trip the in-place-edit workflow relies on.
const found = await api("GET", `sitepages/pages?$select=Id&$filter=UniqueId eq guid'${created.UniqueId}'`);
const intId = found.value[0].Id;
const raw = await api("GET", `sitepages/pages(${intId})?$select=Title,CanvasContent1`);
const parsed = parseCanvasContent1(raw.CanvasContent1 ?? "[]");
const kinds = parsed.flatMap((s: any) => s.columns.flatMap((c: any) => c.webparts.map((w: any) => w.kind)));
console.log("CANVAS READ-BACK parsed sections:", parsed.length, " backgrounds:", JSON.stringify(parsed.map((s: any) => s.background)), " kinds:", JSON.stringify(kinds));
const expected = ["text", "divider", "spacer", "button", "codeSnippet", "quickLinks", "embed"];
const missing = expected.filter((k) => !kinds.includes(k));
console.log(missing.length ? `  ⚠ MISSING kinds after round-trip: ${missing.join(", ")}` : "  ✅ all new kinds round-tripped");

await api("POST", `sitepages/pages(${intId})/checkoutpage`);
await api("POST", `sitepages/pages(${intId})/SavePageAsDraft`, { CanvasContent1: buildCanvasContent1([{ layout: "oneColumn", columns: [{ width: 12, webparts: [{ kind: "text", html: "<p>updated</p>" }] }] }] as any) });
const page2: any = await graphClient().api(`/sites/${SITE_ID}/pages/${created.UniqueId}/microsoft.graph.sitePage`).expand("canvasLayout").get();
const n2 = page2.canvasLayout.horizontalSections.reduce((a: number, s: any) => a + s.columns.reduce((b: number, c: any) => b + c.webparts.length, 0), 0);
console.log("UPDATE ok  parts after update:", n2);

// DELETE
await fetch(`${webUrl}/_api/sitepages/pages(${intId})`, { method: "POST", headers: { ...H, "X-HTTP-Method": "DELETE", "IF-MATCH": "*" } });
console.log("DELETE ok");
