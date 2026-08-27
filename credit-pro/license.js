/* ============================================================================
 * license.js — 決済確認とロック解除
 *
 * 仕組み:
 *   1. Squareの支払いリンクで決済すると ?transactionId=… を付けて戻ってくる
 *   2. その注文番号を localStorage に保存する
 *   3. Worker の /verify?order=… に問い合わせ、valid:true ならロックを外す
 *
 * 方針:
 *   ・画面上の判定は常に無料。課金の対象はExcelファイルの受け取りのみ
 *   ・購入前に判定結果を最後まで見せる。何を買うのか分からないまま払わせない
 *   ・Workerに障害があっても、画面上の判定は動き続ける（決済確認と計算を分離）
 * ========================================================================== */

const WORKER = "https://square-license.stats-okinawa.workers.dev";

/* ★★ ここだけ書き換えてください ★★
   【注意】このファイルを新しい版で上書きすると、下のPAY_URLも一緒に置き換わります。
   ZIPを差し替えたあとは、必ずこの行が自分のリンクになっているか確認してください。
   Squareの支払いリンクURL。Square管理画面で作成して貼り替えます。
   作成時の設定：
     金額 500円 ／ Frequency: One-time（Monthlyにすると定期課金になります）／
     Redirect to a website after checkout: ON
     リダイレクト先は「今そのページを公開しているURL」にします。
       テスト環境  : https://ck121212195.github.io/PDF/credit-pro/
       本番        : https://kazumono.com/credit-pro/
   未設定（XXXXXXXXのまま）だと購入ボタンは押せず、画面にその旨を表示します。 */
const PAY_URL = "https://square.link/u/E7Xmnxwy";

/** 支払いリンクが未設定かどうか。未設定のまま黙って遷移させないための判定 */
export function payUrlReady() {
  return typeof PAY_URL === "string" && /^https:\/\/square\.link\/u\/[A-Za-z0-9]+$/.test(PAY_URL)
    && !PAY_URL.includes("XXXXXXXX");
}

const KEY = "kazumono.credit-pro.order";

/** URLに注文番号が付いていれば保存し、URLからは消す（リロードで消えないように） */
function captureOrder() {
  const q = new URLSearchParams(location.search);
  const oid = q.get("orderId") || q.get("transactionId");
  if (oid) {
    try { localStorage.setItem(KEY, oid); } catch (e) { /* プライベートモード等 */ }
    q.delete("orderId"); q.delete("transactionId");
    const rest = q.toString();
    history.replaceState(null, "", location.pathname + (rest ? "?" + rest : ""));
    return oid;
  }
  try { return localStorage.getItem(KEY); } catch (e) { return null; }
}

/**
 * ライセンス状態を返す。
 * @returns {Promise<{state:"licensed"|"unlicensed"|"offline", order:string|null}>}
 *   offline は「確認できなかった」状態。決済済みの人を締め出さないため区別する。
 */
/**
 * 会社名を、そのままでは送らずにハッシュ化する。
 * 1回の決済を1社分に限るための目印として使う。
 * 取引先の名前をこちらのサーバに残さないため、必ずブラウザ内で変換してから送る。
 */
export async function companyFingerprint(name) {
  const norm = String(name || "")
    .normalize("NFKC")
    .replace(/[\s　]/g, "")
    .replace(/[（(].*?[）)]/g, "")
    .toLowerCase();
  const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode("co:" + norm));
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, "0")).join("").slice(0, 32);
}

/**
 * 解錠してよいかをWorkerに尋ねる。
 * 判定はすべてWorker側で行う。画面のJSは誰でも書き換えられるため、
 * ここで独自に判断しても意味がないので、返ってきた答えに従うだけにする。
 */
export async function checkLicense(fp) {
  const order = captureOrder();
  if (!order) return { state: "unlicensed", order: null, reason: null };
  try {
    const ctl = new AbortController();
    const timer = setTimeout(() => ctl.abort(), 8000);
    const q = `order=${encodeURIComponent(order)}` + (fp ? `&fp=${encodeURIComponent(fp)}` : "");
    const res = await fetch(`${WORKER}/verify?${q}`, { signal: ctl.signal, cache: "no-store" });
    clearTimeout(timer);
    if (!res.ok) return { state: "offline", order, reason: null };
    const data = await res.json();
    return {
      state: data.valid === true ? "licensed" : "unlicensed",
      order,
      reason: data.reason || null,
      expiresAt: data.expiresAt || null,
    };
  } catch (e) {
    return { state: "offline", order, reason: null };
  }
}

export function payUrl() { return PAY_URL; }

export function forgetOrder() {
  try { localStorage.removeItem(KEY); } catch (e) { /* noop */ }
}
