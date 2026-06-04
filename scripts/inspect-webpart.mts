// Capture the LIVE shape of a web part so we can author it correctly, instead
// of guessing its properties/serverProcessedContent. This is the methodology
// that got Button/QuickLinks right: read a real instance, then encode it.
//
// Two modes:
//   1) Dump the site's web-part manifests (alias -> GUID + default properties):
//        npx tsx scripts/inspect-webpart.mts manifests
//        npx tsx scripts/inspect-webpart.mts manifests hero      (filter by alias/title)
//   2) Dump the parsed CanvasContent1 of an existing page (add the web part you
//      want in the SharePoint UI first, then read it back):
//        npx tsx scripts/inspect-webpart.mts page <pageId-guid>
//
// Requires TEST_SITE_ID (Graph 'host,siteGuid,webGuid') in .env.
import { config as loadDotenv } from "dotenv";
import path from "node:path";
import { fileURLToPath } from "node:url";

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
loadDotenv({ path: path.resolve(scriptDir, "../.env"), quiet: true });

const { graphClient } = await import("../src/graph.js");
const { getSharePointToken } = await import("../src/auth.js");

const SITE_ID = process.env.TEST_SITE_ID;
if (!SITE_ID) {
  console.error("Set TEST_SITE_ID to a 'host,siteGuid,webGuid' site id.");
  process.exit(1);
}
const [mode, arg] = process.argv.slice(2);
if (!mode) {
  console.error("Usage: inspect-webpart.mts manifests [filter] | page <pageGuid>");
  process.exit(1);
}

const web = await graphClient().api(`/sites/${SITE_ID}`).select("webUrl").get();
const webUrl: string = web.webUrl;
const token = await getSharePointToken(new URL(webUrl).hostname);
const H = { Authorization: `Bearer ${token}`, Accept: "application/json;odata=nometadata" };
const api = async (p: string) => {
  const r = await fetch(`${webUrl}/_api/${p}`, { headers: H });
  const t = await r.text();
  if (!r.ok) throw new Error(`GET ${p} -> ${r.status}: ${t.slice(0, 400)}`);
  return t ? JSON.parse(t) : null;
};

if (mode === "manifests") {
  const res = await api("web/GetClientSideWebParts");
  const filter = (arg ?? "").toLowerCase();
  const loc = (v: any) => (!v ? "" : typeof v === "string" ? v : v.default ?? v["en-US"] ?? "");
  for (const p of res.value ?? []) {
    let m: any = {};
    try { m = JSON.parse(p.Manifest ?? "{}"); } catch {}
    const e = m?.preconfiguredEntries?.[0] ?? {};
    const alias = String(m?.alias ?? "");
    const title = loc(e.title);
    if (filter && !alias.toLowerCase().includes(filter) && !title.toLowerCase().includes(filter)) continue;
    console.log(`\n# ${title}  (alias: ${alias})`);
    console.log(`  id: ${p.Id}`);
    console.log(`  version: ${m?.version}`);
    console.log(`  defaultProperties: ${JSON.stringify(e.properties ?? {}, null, 2)}`);
  }
} else if (mode === "page") {
  if (!arg) { console.error("page mode needs a pageId guid"); process.exit(1); }
  const found = await api(`sitepages/pages?$select=Id&$filter=UniqueId eq guid'${arg}'`);
  const intId = found.value?.[0]?.Id;
  if (!intId) { console.error("page not found"); process.exit(1); }
  const page = await api(`sitepages/pages(${intId})?$select=Title,CanvasContent1`);
  const controls = JSON.parse(page.CanvasContent1 ?? "[]");
  console.log(`Title: ${page.Title}\nControls: ${controls.length}\n`);
  for (const c of controls) {
    if (c.controlType === 3 && c.webPartData) {
      console.log(`# ${c.webPartData.title}  (${c.webPartId})  v${c.webPartData.dataVersion}`);
      console.log(`  properties: ${JSON.stringify(c.webPartData.properties, null, 2)}`);
      console.log(`  serverProcessedContent: ${JSON.stringify(c.webPartData.serverProcessedContent, null, 2)}\n`);
    }
  }
}
