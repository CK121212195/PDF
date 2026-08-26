// ============================================================
// square-license-worker.js  （v4）
// 「決済確認 ＋ ライセンス認可」の最小バックエンド
//
// v4の変更点（重要）:
//   ・Squareが戻してくるIDと、名簿に保存するIDが食い違う問題を解消した。
//     Squareの支払いリンクは決済後 ?transactionId=（＝支払いID）を付けて戻す一方、
//     v3は order_id だけを保存していたため、照合できず解錠されなかった。
//     → order_id と payment_id の両方で引けるように保存し、/verify も両方を見る。
//   ・PRODUCTS に 500円を追加（100円はテスト用に残置）
//   ・/verify に no-store を付け、古い結果がキャッシュされないようにした
//   ・/debug は残置（動作確認が済んだら、この関数とルートを削除してよい）
//
// 入口:
//   1) POST /square/webhook … Squareの支払い完了通知を受け、名簿(KV)に記録
//   2) GET  /verify?order=…  … ツールが「このIDは解錠していい?」を確認
//   3) GET  /debug           … 設定の入り具合を確認
//   4) POST /ai              … 支払い済みのIDにだけAIを呼ぶ（今は未使用）
// ============================================================

// 「支払い金額(円)」→「商品ID」。Squareの支払いリンクの金額と一致させる。
const PRODUCTS = {
  100: "tool-100",     // テスト用の100円リンク
  500: "yoshin-pro",   // 本番の500円リンク
};

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    try {
      if (request.method === "OPTIONS") {
        return cors(new Response(null, { status: 204 }), env);
      }
      if (url.pathname === "/debug" && request.method === "GET") {
        return handleDebug(env);
      }
      if (url.pathname === "/square/webhook" && request.method === "POST") {
        return handleWebhook(request, env);
      }
      if (url.pathname === "/verify" && request.method === "GET") {
        return handleVerify(url, env);
      }
      if (url.pathname === "/ai" && request.method === "POST") {
        return handleAI(request, env);
      }
      return new Response("Not found", { status: 404 });
    } catch (err) {
      return new Response("ERROR: " + (err && err.message ? err.message : String(err)), {
        status: 500,
      });
    }
  },
};

// --- 0) 設定が入っているかの健康診断 ---
function handleDebug(env) {
  return Response.json({
    has_LICENSES_KV: !!env.LICENSES,
    has_SQUARE_SIGNATURE_KEY: !!env.SQUARE_SIGNATURE_KEY,
    signature_key_length: env.SQUARE_SIGNATURE_KEY ? env.SQUARE_SIGNATURE_KEY.length : 0,
    WEBHOOK_URL: env.WEBHOOK_URL || null,
    ALLOW_ORIGIN: env.ALLOW_ORIGIN || null,
    has_ANTHROPIC_API_KEY: !!env.ANTHROPIC_API_KEY,
  });
}

// --- 1) Squareからの「支払い完了」通知を受け取る ---
async function handleWebhook(request, env) {
  const raw = await request.text(); // 署名検証には"生の本文"が必要

  if (!env.SQUARE_SIGNATURE_KEY) return new Response("SQUARE_SIGNATURE_KEY is not set", { status: 500 });
  if (!env.WEBHOOK_URL) return new Response("WEBHOOK_URL is not set", { status: 500 });
  if (!env.LICENSES) return new Response("LICENSES KV binding is missing", { status: 500 });

  const signature = request.headers.get("x-square-hmacsha256-signature") || "";
  const ok = await verifySquareSignature(env.WEBHOOK_URL + raw, signature, env.SQUARE_SIGNATURE_KEY);
  if (!ok) return new Response("bad signature", { status: 401 });

  const body = JSON.parse(raw);
  const payment = body?.data?.object?.payment;

  if (payment && payment.status === "COMPLETED") {
    const amount = payment.amount_money?.amount;
    const record = JSON.stringify({
      product: PRODUCTS[amount] || "unknown",
      amount,
      paidAt: new Date().toISOString(),
      active: true,
    });
    // ★ここが肝。ブラウザが持ち帰るIDがどちらでも引けるように、両方を鍵にして保存する。
    //   支払いリンクの戻りは ?transactionId=（＝payment.id）で、order_id ではない。
    if (payment.order_id) await env.LICENSES.put(`order:${payment.order_id}`, record);
    if (payment.id) await env.LICENSES.put(`order:${payment.id}`, record);
  }
  // Squareには必ず200を返す（返さないと何度も再送されてくる）
  return new Response("ok", { status: 200 });
}

// --- 2) ツールが「このID、解錠していい?」と確認する窓口 ---
async function handleVerify(url, env) {
  const order = url.searchParams.get("order") || "";
  const rec = order ? await env.LICENSES.get(`order:${order}`) : null;
  const data = rec ? JSON.parse(rec) : null;
  const res = Response.json({
    valid: !!data && data.active === true,
    product: data?.product ?? null,
  });
  // 未購入の結果が残ると、購入後も解錠されないため必ず毎回問い合わせさせる
  res.headers.set("Cache-Control", "no-store");
  return cors(res, env);
}

// --- 3) 支払い済みのIDにだけAIを呼ぶ（今は未使用） ---
async function handleAI(request, env) {
  const { order, prompt } = await request.json();
  const rec = await env.LICENSES.get(`order:${order || ""}`);
  const data = rec ? JSON.parse(rec) : null;
  if (!data || data.active !== true) {
    return cors(Response.json({ error: "no_license" }, { status: 402 }), env);
  }
  if (!env.ANTHROPIC_API_KEY) {
    return cors(Response.json({ error: "api_key_not_set" }, { status: 503 }), env);
  }
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "x-api-key": env.ANTHROPIC_API_KEY,
      "anthropic-version": "2023-06-01",
    },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 1024,
      messages: [{ role: "user", content: prompt }],
    }),
  });
  return cors(Response.json(await res.json()), env);
}

// --- Square署名の検証（HMAC-SHA256）---
async function verifySquareSignature(message, signatureB64, secret) {
  const enc = new TextEncoder();
  const cryptoKey = await crypto.subtle.importKey(
    "raw",
    enc.encode(secret),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"]
  );
  const sigBuf = await crypto.subtle.sign("HMAC", cryptoKey, enc.encode(message));
  const expected = btoa(String.fromCharCode(...new Uint8Array(sigBuf)));
  return timingSafeEqual(expected, signatureB64);
}

function timingSafeEqual(a, b) {
  if (a.length !== b.length) return false;
  let out = 0;
  for (let i = 0; i < a.length; i++) out |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return out === 0;
}

// --- CORS ---
function cors(res, env) {
  res.headers.set("Access-Control-Allow-Origin", env.ALLOW_ORIGIN || "*");
  res.headers.set("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
  res.headers.set("Access-Control-Allow-Headers", "content-type");
  return res;
}
