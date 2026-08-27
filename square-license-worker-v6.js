// ============================================================
// square-license-worker.js  （v6）
//
// v6の狙い:
//   「1回の決済 ＝ 1社分」を、サーバ側で守れるようにする。
//   ・最初に解錠した会社名（のハッシュ）を注文番号に結び付ける
//   ・以後、その会社でしか解錠できない
//   ・有効期間は最初の解錠から24時間。ダウンロードのやり直しは何度でもできる
//
//   画面のJSは誰でも読めるので、ブラウザ側だけの制限は意味を持たない。
//   だから判断はすべてこのWorkerが行い、画面はその答えに従うだけにする。
//
// 会社名そのものは受け取らない。ブラウザ側でハッシュ化した文字列だけを受け取る。
// 取引先の名前をこちらに残さないための措置。
//
// 入口:
//   POST /square/webhook  … 支払い完了通知を受け、注文番号を名簿(KV)に記録
//   GET  /verify?order=…&fp=… … 解錠してよいかを返す（初回に会社を結び付ける）
//   GET  /last            … 直近のWebhook受信の記録（調査用）
//   GET  /peek?order=…    … その番号の記録を見る（調査用・解錠はしない）
//   GET  /debug           … 設定の入り具合（調査用）
//
// 調査用の3つ（/last /peek /debug）は、本番公開前に削除してよい。
// ============================================================

const PRODUCTS = {
  100: "tool-100",     // テスト用の100円リンク
  500: "yoshin-pro",   // 本番の500円リンク
};

// 有効期間。最初に解錠した時点から数える。
const VALID_HOURS = 24;

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

function handleDebug(env) {
  return Response.json({
    has_LICENSES_KV: !!env.LICENSES,
    has_SQUARE_SIGNATURE_KEY: !!env.SQUARE_SIGNATURE_KEY,
    signature_key_length: env.SQUARE_SIGNATURE_KEY ? env.SQUARE_SIGNATURE_KEY.length : 0,
    WEBHOOK_URL: env.WEBHOOK_URL || null,
    ALLOW_ORIGIN: env.ALLOW_ORIGIN || null,
    valid_hours: VALID_HOURS,
  });
}

async function handleLast(env) {
  const rec = await env.LICENSES.get("debug:last");
  return Response.json(rec ? JSON.parse(rec) : { note: "まだ一度もWebhookを受信していません" });
}

async function handlePeek(url, env) {
  const order = url.searchParams.get("order") || "";
  const rec = await env.LICENSES.get(`order:${order}`);
  return Response.json({ order, found: !!rec, record: rec ? JSON.parse(rec) : null });
}

// --- Squareからの支払い完了通知 ---
async function handleWebhook(request, env) {
  const raw = await request.text();          // 署名検証には"生の本文"が必要
  const signature = request.headers.get("x-square-hmacsha256-signature") || "";

  // 届いた事実を、検証の成否にかかわらず必ず残す（調査用）
  const log = {
    receivedAt: new Date().toISOString(),
    hasSignatureHeader: !!signature,
    bodyLength: raw.length,
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

    if (!env.SQUARE_SIGNATURE_KEY) log.note = "SQUARE_SIGNATURE_KEY が未設定";
    else if (!env.WEBHOOK_URL) log.note = "WEBHOOK_URL が未設定";
    else {
      log.signatureOk = await verifySquareSignature(
        env.WEBHOOK_URL + raw, signature, env.SQUARE_SIGNATURE_KEY);
      if (!log.signatureOk) log.note = "署名が一致しない。SQUARE_SIGNATURE_KEY か WEBHOOK_URL が購読設定と食い違っている";
    }

    if (log.signatureOk && payment && payment.status === "COMPLETED") {
      const amount = payment.amount_money?.amount;
      const record = JSON.stringify({
        product: PRODUCTS[amount] || "unknown",
        amount,
        paidAt: new Date().toISOString(),
        active: true,
        boundFp: null,     // まだどの会社にも結び付いていない
        boundAt: null,     // 最初に解錠した時刻
        uses: 0,
      });
      // 支払いリンク経由では transactionId と order_id が同じ値だが、念のため両方を鍵にする
      if (payment.id) { await env.LICENSES.put(`order:${payment.id}`, record); log.savedKeys.push(payment.id); }
      if (payment.order_id && payment.order_id !== payment.id) {
        await env.LICENSES.put(`order:${payment.order_id}`, record); log.savedKeys.push(payment.order_id);
      }
      if (!log.savedKeys.length) log.note = "payment に id も order_id も無かった";
    }
  } catch (e) {
    log.note = "本文の解析に失敗: " + (e && e.message ? e.message : String(e));
  }

  await env.LICENSES.put("debug:last", JSON.stringify(log));
  return new Response("ok", { status: 200 });   // Squareには常に200を返す
}

// --- 解錠してよいかの判定（ここが本体） ---
async function handleVerify(url, env) {
  const order = url.searchParams.get("order") || "";
  const fp = url.searchParams.get("fp") || "";        // 会社名のハッシュ
  const rec = order ? await env.LICENSES.get(`order:${order}`) : null;

  if (!rec) return jsonNoStore({ valid: false, reason: "not_found" }, env);

  const data = JSON.parse(rec);
  if (data.active !== true) return jsonNoStore({ valid: false, reason: "inactive" }, env);

  const now = Date.now();

  // ① 初回：この注文番号を、いま判定している会社に結び付ける
  if (!data.boundFp) {
    if (!fp) return jsonNoStore({ valid: false, reason: "no_fingerprint" }, env);
    data.boundFp = fp;
    data.boundAt = now;
    data.uses = 1;
    await env.LICENSES.put(`order:${order}`, JSON.stringify(data));
    return jsonNoStore({
      valid: true, product: data.product,
      expiresAt: new Date(now + VALID_HOURS * 3600e3).toISOString(),
    }, env);
  }

  // ② 期限切れ
  const expires = data.boundAt + VALID_HOURS * 3600e3;
  if (now > expires) {
    return jsonNoStore({ valid: false, reason: "expired",
      expiredAt: new Date(expires).toISOString() }, env);
  }

  // ③ 別の会社に使い回そうとしている
  if (data.boundFp !== fp) {
    return jsonNoStore({ valid: false, reason: "other_company" }, env);
  }

  // ④ 同じ会社・期間内 → 何度でも解錠する
  data.uses = (data.uses || 0) + 1;
  await env.LICENSES.put(`order:${order}`, JSON.stringify(data));
  return jsonNoStore({
    valid: true, product: data.product, uses: data.uses,
    expiresAt: new Date(expires).toISOString(),
  }, env);
}

function jsonNoStore(obj, env) {
  const res = Response.json(obj);
  // 「未購入」の答えが残ると、購入後も解錠されなくなるため必ず毎回問い合わせさせる
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
