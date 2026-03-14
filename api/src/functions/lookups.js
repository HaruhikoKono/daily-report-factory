// ファイル名：C:\Users\senzo\Desktop\report-functions\src\functions\lookups.js
// 役割：SharePoint のマスタ（Lookup元）リストから候補を取って返す（フロントでプルダウンに使う）

const { app } = require("@azure/functions");

app.http("lookups", {
  methods: ["GET"],
  authLevel: "anonymous",
  handler: async (request, context) => {
    try {
      // ===== 1) 環境変数 =====
      const TENANT_ID = process.env.TENANT_ID;
      const CLIENT_ID = process.env.CLIENT_ID;
      const CLIENT_SECRET = process.env.CLIENT_SECRET;
      const SITE_ID = process.env.SITE_ID;

      // Lookup元リストID（あなたが貼ってくれた columns の listId をそのまま入れる想定）
      const LIST_WORKER = process.env.LIST_WORKER;   // 例: 5b5821ae-e671-4c3f-a250-2e5f7e06787d
      const LIST_VEHICLE = process.env.LIST_VEHICLE; // 例: d2f3d2f8-7613-48a3-a7dc-811cc707e7aa
      const LIST_CHAIN = process.env.LIST_CHAIN;     // 例: 20e8f681-1160-45c3-8f72-e52d54c39200
      const LIST_STORE = process.env.LIST_STORE;     // 例: 14b20563-c6dd-431e-9c10-590450da6a5c
      const LIST_TASK = process.env.LIST_TASK;       // 例: 431f5914-6549-4d92-a772-0da72e127227

      const missing = [];
      if (!TENANT_ID) missing.push("TENANT_ID");
      if (!CLIENT_ID) missing.push("CLIENT_ID");
      if (!CLIENT_SECRET) missing.push("CLIENT_SECRET");
      if (!SITE_ID) missing.push("SITE_ID");
      if (!LIST_WORKER) missing.push("LIST_WORKER");
      if (!LIST_VEHICLE) missing.push("LIST_VEHICLE");
      if (!LIST_CHAIN) missing.push("LIST_CHAIN");
      if (!LIST_STORE) missing.push("LIST_STORE");
      if (!LIST_TASK) missing.push("LIST_TASK");

      if (missing.length) {
        return {
          status: 500,
          headers: { "Content-Type": "application/json; charset=utf-8" },
          body: JSON.stringify(
            { ok: false, error: `local.settings.json の Values に ${missing.join(", ")} がありません` },
            null,
            2
          ),
        };
      }

      // ===== 2) アクセストークン取得（アプリ権限） =====
      const tokenRes = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id: CLIENT_ID,
          client_secret: CLIENT_SECRET,
          grant_type: "client_credentials",
          scope: "https://graph.microsoft.com/.default",
        }),
      });

      const tokenJson = await tokenRes.json();
      if (!tokenRes.ok) {
        return {
          status: 500,
          headers: { "Content-Type": "application/json; charset=utf-8" },
          body: JSON.stringify({ ok: false, where: "token", status: tokenRes.status, detail: tokenJson }, null, 2),
        };
      }
      const accessToken = tokenJson.access_token;

      // ===== 3) 共通：リストから items を取得する関数 =====
      // 返す形式：[{ id, title, fields }]
      const fetchListItems = async (listId) => {
        // $expand=fields で fields を取得。$top は大きめに。
        const url =
          `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${listId}/items` +
          `?$top=999&$expand=fields`;

        const res = await fetch(url, {
          method: "GET",
          headers: { Authorization: `Bearer ${accessToken}` },
        });
        const json = await res.json();
        if (!res.ok) {
          throw new Error(`list items fetch failed: ${listId} status=${res.status} detail=${JSON.stringify(json)}`);
        }

        const items = (json.value ?? []).map((it) => ({
          id: it.id, // SharePoint listItem id（数値文字列）
          title: it?.fields?.Title ?? "",
          fields: it?.fields ?? {},
        }));

        return items;
      };

      // ===== 4) どれを返すか（クエリで指定できる） =====
      // 例: /api/lookups?only=task  or  /api/lookups?only=all
      const only = (request.query.get("only") ?? "all").toLowerCase();

      const out = { ok: true };

      if (only === "all" || only === "worker") out.worker = await fetchListItems(LIST_WORKER);
      if (only === "all" || only === "vehicle") out.vehicle = await fetchListItems(LIST_VEHICLE);
      if (only === "all" || only === "chain") out.chain = await fetchListItems(LIST_CHAIN);
      if (only === "all" || only === "store") out.store = await fetchListItems(LIST_STORE);
      if (only === "all" || only === "task") out.task = await fetchListItems(LIST_TASK);

      return {
        status: 200,
        headers: { "Content-Type": "application/json; charset=utf-8" },
        body: JSON.stringify(out, null, 2),
      };
    } catch (e) {
      return {
        status: 500,
        headers: { "Content-Type": "application/json; charset=utf-8" },
        body: JSON.stringify({ ok: false, error: e?.message ?? String(e) }, null, 2),
      };
    }
  },
});
