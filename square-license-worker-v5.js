// ============================================================
// square-license-worker.js  （v5: 原因究明モード）
//
// v5でやりたいこと:
//   「決済したのに解錠されない」の原因は、次のどれかしかない。
//     A. Webhookがそもそも届いていない
//     B. 届いているが署名検証で弾かれている（401）
//     C. 届いて保存もされたが、ブラウザが持ち帰るIDと保存したIDが違う
//   v5は、届いた事実と結果をKVに記録し、/last で見えるようにした。
//   これでA・B・Cのどれかが即座に分かる。
//
// 追加した入口:
//   GET /last         … 直近のWebhook受信の記録（届いたか / 署名OKか / 保存した鍵）
//   GET /peek?order=… … その番号が名簿にあるかを確認（解錠はしない）
//
// 原因が分かったら、handleLast / handlePeek / handleDebug は削除してよい。
// ============================================================

const PRODUCTS = {
  100: "tool-100",     // テスト用の100円リンク
  500: "yoshin-pro",   // 本番の500円リンク
};

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    try {
      if (request.method === "OPTIONS") return cors(new Response(null, { status: 204 }), env);
      if (url.pathname === "/debug" && request.method === "GET") return handleDebug(env);
      if (url.pathname === "/last" && request.method === "GET") return handleLast(env);
      if (url.pathname === "/peek" && request.method === "GET") return handlePeek(url, env);
      if (url.pathname === "/square/webhook" && request.method === "POST") return handleWebhook(request, env);
      if (url.pathname === "/verify" && request.method === "GET") return handleVerify(url, env);
      return new Response("Not found", { status: 404 });
    } catch (err) {
      return new Response("ERROR: " + (err && err.message ? err.message : String(err)), { status: 500 });
    }
  },
};

// --- 設定の健康診断 ---
function handleDebug(env) {
  return Response.json({
    has_LICENSES_KV: !!env.LICENSES,
    has_SQUARE_SIGNATURE_KEY: !!env.SQUARE_SIGNATURE_KEY,
    signature_key_length: env.SQUARE_SIGNATURE_KEY ? env.SQUARE_SIGNATURE_KEY.length : 0,
    WEBHOOK_URL: env.WEBHOOK_URL || null,
    ALLOW_ORIGIN: env.ALLOW_ORIGIN || null,
  });
}

// --- 直近のWebhook受信記録を見る（これが空なら「届いていない」） ---
async function handleLast(env) {
  const rec = await env.LICENSES.get("debug:last");
  return Response.json(rec ? JSON.parse(rec) : { note: "まだ一度もWebhookを受信していません" });
}

// --- その番号が名簿にあるか確認する（解錠はしない・調査用） ---
async function handlePeek(url, env) {
  const order = url.searchParams.get("order") || "";
  const rec = await env.LICENSES.get(`order:${order}`);
  return Response.json({ order, found: !!rec, record: rec ? JSON.parse(rec) : null });
}

// --- Squareからの支払い完了通知 ---
async function handleWebhook(request, env) {
  const raw = await request.text();          // 署名検証には"生の本文"が必要
  const signature = request.headers.get("x-square-hmacsha256-signature") || "";

  // ★届いた事実を、検証の成否にかかわらず必ず残す。
  //   これが無いと「届いていない」のか「弾かれた」のかが永遠に分からない。
  const log = {
    receivedAt: new Date().toISOString(),
    hasSignatureHeader: !!signature,
    bodyLength: raw.length,
    webhookUrlUsed: env.WEBHOOK_URL || null,
    signatureOk: null,
    eventType: null,
    paymentStatus: null,
    savedKeys: [],
    note: "",
  };

  try {
    const body = JSON.parse(raw);
    log.eventType = body?.type || null;
    const payment = body?.data?.object?.payment;
    log.paymentStatus = payment?.status || null;

    if (!env.SQUARE_SIGNATURE_KEY) { log.note = "SQUARE_SIGNATURE_KEY が未設定"; }
    else if (!env.WEBHOOK_URL) { log.note = "WEBHOOK_URL が未設定"; }
    else {
      log.signatureOk = await verifySquareSignature(
        env.WEBHOOK_URL + raw, signature, env.SQUARE_SIGNATURE_KEY);
      if (!log.signatureOk) {
        log.note = "署名が一致しない。SQUARE_SIGNATURE_KEY か WEBHOOK_URL が、Squareの購読設定と食い違っている";
      }
    }

    // 署名が通り、支払いが完了しているときだけ名簿に書く
    if (log.signatureOk && payment && payment.status === "COMPLETED") {
      const amount = payment.amount_money?.amount;
      const record = JSON.stringify({
        product: PRODUCTS[amount] || "unknown",
        amount,
        paidAt: new Date().toISOString(),
        active: true,
      });
      // ブラウザが持ち帰るIDは支払いID（?transactionId=）。注文IDとは別物なので両方を鍵にする
      if (payment.id) { await env.LICENSES.put(`order:${payment.id}`, record); log.savedKeys.push(payment.id); }
      if (payment.order_id) { await env.LICENSES.put(`order:${payment.order_id}`, record); log.savedKeys.push(payment.order_id); }
      if (!log.savedKeys.length) log.note = "payment に id も order_id も無かった";
    }
  } catch (e) {
    log.note = "本文の解析に失敗: " + (e && e.message ? e.message : String(e));
  }

  await env.LICENSES.put("debug:last", JSON.stringify(log));
  // Squareには常に200を返す（返さないと何度も再送されてくる）
  return new Response("ok", { status: 200 });
}

// --- ツールからの解錠確認 ---
async function handleVerify(url, env) {
  const order = url.searchParams.get("order") || "";
  const rec = order ? await env.LICENSES.get(`order:${order}`) : null;
  const data = rec ? JSON.parse(rec) : null;
  const res = Response.json({
    valid: !!data && data.active === true,
    product: data?.product ?? null,
  });
  res.headers.set("Cache-Control", "no-store");
  return cors(res, env);
}

// --- Square署名の検証（HMAC-SHA256）---
async function verifySquareSignature(message, signatureB64, secret) {
  const enc = new TextEncoder();
  const key = await crypto.subtle.importKey(
    "raw", enc.encode(secret), { name: "HMAC", hash: "SHA-256" }, false, ["sign"]);
  const sigBuf = await crypto.subtle.sign("HMAC", key, enc.encode(message));
  const expected = btoa(String.fromCharCode(...new Uint8Array(sigBuf)));
  return timingSafeEqual(expected, signatureB64);
}

function timingSafeEqual(a, b) {
  if (a.length !== b.length) return false;
  let out = 0;
  for (let i = 0; i < a.length; i++) out |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return out === 0;
}

function cors(res, env) {
  res.headers.set("Access-Control-Allow-Origin", env.ALLOW_ORIGIN || "*");
  res.headers.set("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
  res.headers.set("Access-Control-Allow-Headers", "content-type");
  return res;
}
