const { app } = require('@azure/functions');

app.http('report', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request) => {
    try {
      // ==========================
      // 0. JSON 取得（不正なら 400）
      // ==========================
      let data;
      try {
        data = await request.json();
      } catch {
        return { status: 400, jsonBody: { ok: false, error: 'Invalid JSON' } };
      }

      const {
        TENANT_ID,
        CLIENT_ID,
        CLIENT_SECRET,
        SITE_ID,
        REPORT_LIST_ID
      } = process.env;

      // ==========================
      // 0.5 必須 env チェック（未設定なら 500）
      // ==========================
      const missing = [];
      if (!TENANT_ID) missing.push('TENANT_ID');
      if (!CLIENT_ID) missing.push('CLIENT_ID');
      if (!CLIENT_SECRET) missing.push('CLIENT_SECRET');
      if (!SITE_ID) missing.push('SITE_ID');
      if (!REPORT_LIST_ID) missing.push('REPORT_LIST_ID');
      if (missing.length) {
        return {
          status: 500,
          jsonBody: { ok: false, error: `local.settings.json の ${missing.join(', ')} が未設定です` }
        };
      }

      // ==========================
      // 1. Graph トークン取得
      // ==========================
      const tokenRes = await fetch(
        `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
          body: new URLSearchParams({
            client_id: CLIENT_ID,
            client_secret: CLIENT_SECRET,
            grant_type: 'client_credentials',
            scope: 'https://graph.microsoft.com/.default'
          })
        }
      );

      const tokenJson = await tokenRes.json();
      if (!tokenRes.ok) {
        return {
          status: 500,
          jsonBody: { ok: false, where: 'token', status: tokenRes.status, detail: tokenJson }
        };
      }

      const token = tokenJson.access_token;

      // ==========================
      // 2. ヘッダ（親レコード）作成
      // ==========================
      // ★確定：参照列の内部名（Graph 用）
      // Workername   → WorkernameLookupId
      // VehicleModel → VehicleModelLookupId
      // VehicleRegNo → VehicleRegNoLookupId
      //
      // Reportdate は "YYYY-MM-DD" を渡すと UTC 変換され、
      // 例: 2026-02-17 → 2026-02-16T15:00:00Z になることがある（JST→UTCのため）
      const headerFields = {
        Title: data.requestId || 'NO-TITLE',
        RequestId: data.requestId || '',
        Reportdate: data.reportDate || null,
        MeterStart: data.meterStart ?? null,
        MeterEnd: data.meterEnd ?? null,
        WorkernameLookupId: data.workerId || null,
        VehicleModelLookupId: data.vehicleId || null,
        VehicleRegNoLookupId: data.vehicleId || null
      };

      const headerRes = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${REPORT_LIST_ID}/items`,
        {
          method: 'POST',
          headers: {
            Authorization: `Bearer ${token}`,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({ fields: headerFields })
        }
      );

      const headerJson = await headerRes.json();
      if (!headerRes.ok) {
        return {
          status: 500,
          jsonBody: {
            ok: false,
            where: 'create-header',
            status: headerRes.status,
            detail: headerJson,
            sent: headerFields
          }
        };
      }

      // ※現状：親IDは使っていない（必要なら後で明細側に ParentLookupId 等を追加）
      const parentId = headerJson.id;

      // ==========================
      // 3. 明細を1件ずつ追加（成功/失敗を集計）
      // ==========================
      const details = Array.isArray(data.details) ? data.details : [];
      const detailResults = [];

      for (let i = 0; i < details.length; i++) {
        const d = details[i] || {};

        const detailFields = {
          Title: data.requestId || 'NO-TITLE',
          RequestId: data.requestId || '',
          Reportdate: data.reportDate || null,
          StartTime: combineDateTime(data.reportDate, d.startTime),
          EndTime: combineDateTime(data.reportDate, d.endTime),

          WorkernameLookupId: data.workerId || null,
          VehicleModelLookupId: data.vehicleId || null,
          VehicleRegNoLookupId: data.vehicleId || null,

          TasknameLookupId: d.taskId || null,
          ChainnameLookupId: d.chainId || null,
          StorenameLookupId: d.storeId || null,

          Note: d.note || ''
        };

        try {
          const detailRes = await fetch(
            `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${REPORT_LIST_ID}/items`,
            {
              method: 'POST',
              headers: {
                Authorization: `Bearer ${token}`,
                'Content-Type': 'application/json'
              },
              body: JSON.stringify({ fields: detailFields })
            }
          );

          const detailJson = await detailRes.json().catch(() => ({}));

          if (!detailRes.ok) {
            detailResults.push({
              index: i,
              ok: false,
              where: 'create-detail',
              status: detailRes.status,
              detail: detailJson,
              sent: detailFields
            });
            continue;
          }

          detailResults.push({
            index: i,
            ok: true,
            createdId: detailJson?.id ?? null
          });
        } catch (e) {
          detailResults.push({
            index: i,
            ok: false,
            where: 'create-detail-fetch',
            error: e?.message ?? String(e),
            sent: detailFields
          });
        }
      }

      const detailOk = detailResults.filter(r => r.ok).length;
      const detailNg = detailResults.length - detailOk;

      return {
        status: 200,
        jsonBody: {
          ok: true,
          headerCreated: headerJson,
          headerSent: headerFields,
          parentId, // 将来の紐付け用に返すだけ
          detailsTotal: detailResults.length,
          detailsOk: detailOk,
          detailsNg: detailNg,
          detailResults, // 失敗時の理由がここに残る
          received: data
        }
      };

    } catch (e) {
      return {
        status: 500,
        jsonBody: { ok: false, error: e?.message ?? String(e) }
      };
    }
  }
});

// ==========================
// 日付 + 時刻をISO形式に変換
// ==========================
function combineDateTime(dateStr, timeStr) {
  if (!dateStr || !timeStr) return null;
  return `${dateStr}T${timeStr}:00`;
}
