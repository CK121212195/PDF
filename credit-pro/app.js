/* ============================================================================
 * /credit-pro/app.js — 画面と判定エンジンの接続
 * 計算は engine.js、Excel生成は xlsx-export.js。ここはUIだけを担当する。
 * ========================================================================== */
import { evaluate, emptyInput, INDUSTRIES, CAPITAL_TIERS, LISTING_OPTIONS, POLICY }
  from "./engine.js?v=14";
import { downloadXlsx } from "./xlsx-export.js?v=14";
import { checkLicense, payUrl, payUrlReady, companyFingerprint, forgetOrder } from "./license.js?v=14";
import { scanPdf, buildPeriod, validatePeriod, toEngineFields } from "./pdf-extract.js?v=14";

const $ = (id) => document.getElementById(id);
const COLS = ["今期（直近）", "前期", "前々期"];

const PL_ROWS = [
  ["sales", "売上高", 0], ["cogs", "売上原価", 0],
  [null, "売上総利益（粗利益）", "grossProfit"],
  ["sga", "販売費及び一般管理費", 0],
  [null, "営業利益", "operatingProfit"],
  ["nonOpInc", "営業外収益", 0], ["nonOpExp", "営業外費用", 0],
  [null, "経常利益", "ordinaryProfit"],
  ["extraInc", "特別利益", 0], ["extraExp", "特別損失", 0],
  [null, "税引前当期純利益", "pretaxProfit"],
  ["tax", "法人税・住民税及び事業税等", 0],
  [null, "当期純利益", "netProfit"],
  ["depreciation", "減価償却費（販管費・製造原価の合計）", 0],
];
const BS_ROWS = [
  ["cash", "現金・預金", 0], ["receivables", "受取手形・売掛金", 0],
  ["inventory", "棚卸資産", 0], ["otherCurrentAssets", "その他流動資産", 0],
  [null, "流動資産合計", "currentAssets"],
  ["tangible", "有形固定資産", 0], ["otherFixedAssets", "無形固定資産・投資その他", 0],
  [null, "固定資産合計", "fixedAssets"],
  ["deferred", "繰延資産", 0],
  [null, "資産合計", "totalAssets"],
  ["payables", "支払手形・買掛金", 0],
  ["shortDebt", "短期借入金（1年内返済分を含む）", 0],
  ["otherCurrentLiab", "その他流動負債", 0],
  [null, "流動負債合計", "currentLiab"],
  ["longDebt", "長期借入金・社債", 0], ["otherFixedLiab", "その他固定負債", 0],
  [null, "固定負債合計", "fixedLiab"],
  [null, "負債合計", "totalLiab"],
  ["equity", "純資産合計（自己資本）", 0],
  [null, "負債・純資産合計（総資本）", "totalCapital"],
  [null, "【検算】資産合計 － 負債・純資産合計", "balanceCheck"],
];

let state = emptyInput();

/* ------------------------------------------------------------------ 表示補助 */
const yen = (n) => (!n ? "0" : (n < 0 ? "▲" : "") + Math.abs(Math.round(n)).toLocaleString());
const pct = (n) => (n * 100).toFixed(1) + "%";
const esc = (s) => String(s ?? "").replace(/[&<>"]/g,
  (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

function buildTable(el, rows, withTerms) {
  let h = "<thead><tr><th style='text-align:left'>項　目<span class=\"unit-tag\">単位：百万円</span></th>";
  COLS.forEach((c) => (h += `<th>${c}</th>`));
  h += "</tr></thead><tbody>";
  if (withTerms) {
    h += "<tr><td class='lb'>決算期</td>";
    for (let i = 0; i < 3; i++)
      h += `<td><input type="text" data-k="terms" data-i="${i}" style="text-align:center"></td>`;
    h += "</tr>";
  }
  for (const [key, label, calc] of rows) {
    const isCalc = key === null;
    h += `<tr class="${isCalc ? "total" : ""}"><td class="lb">${label}</td>`;
    for (let i = 0; i < 3; i++) {
      h += isCalc
        ? `<td class="sum" data-calc="${calc}" data-i="${i}">0</td>`
        : `<td><input type="number" step="1" data-k="${key}" data-i="${i}"></td>`;
    }
    h += "</tr>";
  }
  el.innerHTML = h + "</tbody>";
}

/* ---------------------------------------------------------------- 初期化 */
function init() {
  INDUSTRIES.forEach((r) => $("f_industry").add(new Option(r[0], r[0])));
  CAPITAL_TIERS.forEach((t) => $("f_capitalTier").add(new Option(t[0], t[0])));
  LISTING_OPTIONS.forEach((l) => $("f_listing").add(new Option(l, l)));
  buildTable($("tPL"), PL_ROWS, true);
  buildTable($("tBS"), BS_ROWS, false);
  $("tRepay").innerHTML =
    "<thead><tr><th style='text-align:left'>項　目</th><th>1年目</th><th>2年目</th><th>3年目</th></tr></thead>" +
    "<tbody><tr><td class='lb'>年間約定返済額（元金）</td>" +
    [0, 1, 2].map((i) => `<td><input type="number" step="1" data-k="repayment" data-i="${i}"></td>`).join("") +
    "</tr></tbody>";

  document.addEventListener("input", onInput);
  $("btnDemo").addEventListener("click", () => { state = demo(); paint(); render(); openManual(); });
  $("btnClear").addEventListener("click", () => {
    state = emptyInput();
    state.baseDate = new Date().toISOString().slice(0, 10);
    paint(); render();
  });
  $("btnXlsx").addEventListener("click", onDownload);
  const rc = $("btnRecheck");
  if (rc) rc.addEventListener("click", () => refreshLicense());
  const fg = $("btnForget");
  if (fg) fg.addEventListener("click", () => {
    if (!confirm("この端末に保存された購入情報を消します。まだ有効なお支払いがある場合は、消すと使えなくなります。よろしいですか？")) return;
    forgetOrder();
    refreshLicense();
  });
  $("btnRetry").addEventListener("click", refreshLicense);
  $("btnBuy").addEventListener("click", onBuy);
  initUploader();
  initShots();

  // 決済から戻ったときだけ、入力内容と判定を復元して段を開く。
  // 初回は空の状態で「決算書を置く」だけに集中してもらう。
  const restored = restoreDraft();
  if (!restored) {
    state = emptyInput();
    state.baseDate = new Date().toISOString().slice(0, 10);
  }
  paint(); render();
  if (restored) {
    showStep("step4", true);
    showStep("step5", true);
    refreshLicense();
  }
}

/* ------------------------------------------------------------ 決済ゲート */
let licensed = false;

function showGate(which) {
  ["gateWait", "gateBuy", "gateOk", "gateOffline"].forEach((id) => {
    $(id).hidden = (id !== which);
  });
}

/** 決済まわりで何が起きているかを画面に出し、切り分けられるようにする */
function showLicenseDiag(st, order, reason, expiresAt) {
  const el = $("licDiag");
  if (!el) return;
  const label = { licensed: "購入済み", unlicensed: "未購入", offline: "確認できず" }[st] || st;
  // 解錠されない理由は、利用者が次に何をすればよいか分かる言葉で出す
  const byReason = {
    other_company: "このお支払いは、別の会社の判定に使われています。1回のお支払いにつき1社分です。この会社の分をご入用の場合は、あらためてお求めください。",
    expired: "お支払いから24時間が過ぎたため、期限切れになりました。あらためてお求めください。",
    not_found: "決済の記録が見つかりません。決済直後の場合、記録が届くまで数十秒かかります。少しお待ちください。",
    no_fingerprint: "会社名が未入力です。会社の基本情報に会社名をご入力ください。",
  };
  const why =
    st === "licensed"
      ? "この会社の分を、" + (expiresAt ? expiresAt.slice(0, 16).replace("T", " ") + " まで" : "24時間") + "何度でもダウンロードできます。"
    : st === "offline" ? "決済の確認先に接続できませんでした。通信環境をご確認ください。"
    : (reason && byReason[reason]) ||
      (order ? "決済の記録が見つかりません。" : "まだ決済していない状態です。");
  // 状態が色で分かるようにする。緑＝購入済み、橙＝要対応、既定＝未購入
  el.classList.toggle("tip--warn", st === "unlicensed" && reason !== null);
  el.classList.toggle("tip--ok", st === "licensed");
  // すでに支払っている人に、もう一度払わせないための出し分け
  const paidButLocked = st === "unlicensed" && !!order && reason !== "other_company";
  const buy = $("btnBuy"), recheck = $("btnRecheck");
  if (buy) buy.style.opacity = paidButLocked ? ".45" : "";
  if (recheck) recheck.style.fontSize = paidButLocked ? "16px" : "";
  // 使えない番号（別会社に紐づき済み／期限切れ）が残っていると、
  // 何度開いても橙色の警告が出続ける。その場合だけ消す手段を出す。
  const stale = st === "unlicensed" && !!order &&
                (reason === "other_company" || reason === "expired");
  const fr = $("forgetRow");
  if (fr) fr.hidden = !stale;
  const dup = $("dupWarn");
  if (dup) {
    dup.hidden = !paidButLocked;
    dup.innerHTML = paidButLocked
      ? `<b>お支払いは受け付けられています。もう一度お支払いなさらないでください</b><span>` +
        `確認が済んでいないだけです。「購入状況をもう一度確認する」を押すか、時間をおいて開き直してください。</span>`
      : "";
  }
  el.innerHTML = `<b>決済の状態：${label}</b><span>注文番号：` +
    `<code>${order ? esc(String(order)) : "（なし）"}</code><br>${why}</span>`;
  el.hidden = false;
}

/**
 * 決済直後は、Squareからの通知がこちらの確認より遅れて届くことがある。
 * 1回で諦めると「払ったのに解錠されない」状態のまま終わってしまうため、
 * 注文番号を持っているのに未購入と出た場合だけ、数秒おきに数回だけ確認し直す。
 */
async function pollLicense(order) {
  // Squareからの通知は、実測で2分ほどかかることがある。
  // 「止まっている」と誤解されないよう、経過秒数を出しながら最大3分待つ。
  const TRIES = 60, WAIT = 3000;
  showGate("gateWait");
  for (let i = 0; i < TRIES; i++) {
    await new Promise((r) => setTimeout(r, WAIT));
    const sec = Math.round(((i + 1) * WAIT) / 1000);
    const el = $("licDiag");
    if (el) {
      el.classList.remove("tip--warn");
      el.innerHTML =
        `<b>お支払いを確認しています… 経過 ${sec} 秒</b><span>` +
        `Square からの入金通知が届くまで、<b>2分ほどかかることがあります</b>。` +
        `この画面を開いたままお待ちください。確認できしだい、ダウンロードボタンが出ます。<br>` +
        `ページを閉じても、あとで開き直せば続きから確認できます。<br>` +
        `注文番号：<code>${esc(String(order))}</code></span>`;
    }
    const wait = $("gateWait");
    if (wait) wait.innerHTML =
      `<p class="note-s"><b>お支払いを確認しています…（経過 ${sec} 秒／最大3分）</b><br>` +
      `Squareからの通知待ちです。この画面のままお待ちください。</p>`;
    const { state, expiresAt } = await checkLicense(await companyFingerprint(state_name()));
    if (state === "licensed") {
      licensed = true;
      showLicenseDiag("licensed", order, null, expiresAt);
      showGate("gateOk");
      if (window.gtag) gtag("event", "license_ok", { tool: "credit-pro" });
      return true;
    }
    if (state === "offline") break;
  }
  const el = $("licDiag");
  if (el) {
    el.classList.add("tip--warn");
    el.innerHTML = `<b>3分待っても確認できませんでした</b><span>` +
      `お支払いが完了していれば、記録は必ず後から届きます。<b>二重にお支払いなさらないでください。</b><br>` +
      `少し時間をおいて「購入状況をもう一度確認する」を押すか、このページを開き直してください。<br>` +
      `それでも解錠されない場合は、この注文番号を添えてお問い合わせください：<code>${esc(String(order))}</code></span>`;
    el.hidden = false;
  }
  showGate("gateBuy");
  return false;
}

/** 判定に使っている会社名（Workerには送らず、ハッシュ化して使う） */
function state_name() { return state && state.name ? state.name : ""; }

async function refreshLicense() {
  showGate("gateWait");
  const fp = await companyFingerprint(state_name());
  const { state: st, order, reason, expiresAt } = await checkLicense(fp);
  licensed = st === "licensed";
  showLicenseDiag(st, order, reason, expiresAt);
  // 決済直後は通知の到着が遅れることがあるので、記録が無いときだけ確認し直す
  if (st === "unlicensed" && order && reason === "not_found") { pollLicense(order); return; }
  showGate(st === "licensed" ? "gateOk" : st === "offline" ? "gateOffline" : "gateBuy");
  if (st === "licensed" && window.gtag) gtag("event", "license_ok", { tool: "credit-pro" });
}

/**
 * 購入へ進む。入力内容を保存してからSquareへ送る。
 * 決済後に戻ってきたとき、同じ内容のまま続けられるようにするため。
 */
function onBuy(e) {
  e.preventDefault();
  // 支払いリンクが未設定のまま押されたときは、遷移せずに理由を出す。
  // 黙ってSquareのエラーページへ飛ばすと、原因の切り分けができなくなる。
  // 会社名が空だと、決済しても「どの会社の分か」を確定できず解錠できない
  if (!state_name().trim()) {
    $("buyNote").textContent =
      "会社名をご入力ください。お支払いは1社分ごとのため、会社名が必要です（「会社の基本情報」欄）。";
    $("f_name")?.scrollIntoView({ behavior: "smooth", block: "center" });
    $("f_name")?.focus();
    return;
  }
  if (!payUrlReady()) {
    // 何が読み込まれているかを必ず表示する。
    // 「直したのに直らない」の大半は、ブラウザが古いlicense.jsを使っているだけなので、
    // 実際の値が見えれば一目で切り分けられる。
    $("buyNote").innerHTML =
      "支払いリンクが未設定です。credit-pro/license.js の PAY_URL に、Squareで作成したリンクを貼ってください。<br>" +
      "いま読み込まれている値：<code>" + esc(String(payUrl())) + "</code><br>" +
      "すでに貼り替えたのにこの表示が出る場合は、ブラウザが古いファイルを使っています。" +
      "Ctrl+Shift+R（Mac は Cmd+Shift+R）で読み込み直してください。";
    return;
  }
  try { sessionStorage.setItem("kazumono.credit-pro.draft", JSON.stringify(state)); } catch (err) { /* noop */ }
  if (window.gtag) gtag("event", "begin_checkout", { tool: "credit-pro", value: 500, currency: "JPY" });
  location.href = payUrl();
}

/** 決済から戻ったとき、入力内容を復元する */
function restoreDraft() {
  try {
    const raw = sessionStorage.getItem("kazumono.credit-pro.draft");
    if (!raw) return false;
    const d = JSON.parse(raw);
    if (d && typeof d === "object") { state = { ...emptyInput(), ...d }; return true; }
  } catch (e) { /* noop */ }
  return false;
}

function onInput(e) {
  const t = e.target, k = t.dataset.k, i = t.dataset.i;
  if (k !== undefined && i !== undefined) {
    state[k][+i] = t.type === "number" ? (t.value === "" ? 0 : parseFloat(t.value)) : t.value;
  } else if (t.id && t.id.startsWith("f_")) {
    const f = t.id.slice(2);
    state[f] = t.type === "number" ? (t.value === "" ? 0 : parseFloat(t.value)) : t.value;
  } else return;
  render();
}

function paint() {
  ["name", "industry", "capitalTier", "listing", "founded", "baseDate", "employees",
   "capital", "ceoName", "ceoAge", "industryYears", "ceoYears", "ownHome",
   "disclosure", "successor", "memo"].forEach((f) => {
    const el = $("f_" + f); if (el) el.value = state[f] ?? "";
  });
  document.querySelectorAll("[data-k][data-i]").forEach((el) => {
    const v = state[el.dataset.k]?.[+el.dataset.i];
    el.value = el.type === "number" ? (v === 0 ? "0" : v ?? "") : (v ?? "");
    // PDFから読み取れなかった項目は枠を赤くして、どこを埋めればよいか一目で分かるようにする
    const miss = missingCells[el.dataset.k];
    el.classList.toggle("is-missing", !!miss && miss.includes(+el.dataset.i));
  });
  syncFinAccordions();
}

/**
 * 貸借対照表・損益計算書の折りたたみを、読み取り結果に合わせて開閉する。
 * 全部読めていれば閉じて画面を短くし、抜けがあれば開いて赤い欄を見せる。
 */
function syncFinAccordions() {
  const BS = ["cash","receivables","inventory","otherCurrentAssets","tangible",
              "otherFixedAssets","deferred","payables","shortDebt","otherCurrentLiab",
              "longDebt","otherFixedLiab","equity"];
  const PL = ["sales","cogs","sga","nonOpInc","nonOpExp","extraInc","extraExp","tax","depreciation"];
  const apply = (id, keys, label) => {
    const acc = $(id), st = $(id + "-state");
    if (!acc) return;
    const n = keys.reduce((a, k) => a + (missingCells[k] ? missingCells[k].length : 0), 0);
    const touched = Object.keys(missingCells).length > 0 || anyValue(keys);
    if (!touched) { if (st) st.textContent = ""; return; }
    if (n > 0) {
      acc.open = true;
      if (st) { st.textContent = `${n}か所が未取得です`; st.className = "acc__state is-missing-tag"; }
    } else {
      acc.open = false;
      if (st) { st.textContent = "すべて読み取れました", st.className = "acc__state is-ok-tag"; }
    }
  };
  const anyValue = (keys) => keys.some((k) => Array.isArray(state[k]) && state[k].some((v) => v));
  apply("accBS", BS, "貸借対照表");
  apply("accPL", PL, "損益計算書");
}

/* ---------------------------------------------------------------- 再計算 */
function render() {
  const r = evaluate(state);
  document.querySelectorAll("[data-calc]").forEach((el) => {
    const v = r.fy[+el.dataset.i][el.dataset.calc];
    el.textContent = yen(v);
    if (el.dataset.calc === "balanceCheck")
      el.style.background = Math.abs(v) > 0.5 ? "#FBEDE6" : "";
  });
  const bad = r.fy.map((p, i) => (Math.abs(p.balanceCheck) > 0.5 ? COLS[i] : null)).filter(Boolean);
  const al = $("alertBalance");
  al.className = "alert" + (bad.length ? " on" : "");
  al.textContent = bad.length
    ? `${bad.join("・")}で、資産合計と負債・純資産合計が一致していません。指標がすべて狂うため、必ず0にしてください。`
    : "";
  $("resultCol").innerHTML = report(r);
}

function report(r) {
  const s = r.scores, cur = r.cur, bm = r.benchmark, red = r.redemption;
  const cls = s.rank === "C" ? " is-mid" : (s.rank === "D" || s.rank === "E") ? " is-low" : "";
  const axes = [
    ["① 業歴", s.gyoreki, 10], ["② 資本構成", s.shihon, 12], ["③ 規模", s.kibo, 18],
    ["④ 損益", s.soneki, 10], ["⑤ 経営者", s.keiei, 20], ["⑥ 償還余力", s.shokan, 30],
  ].map(([n, got, max]) => `
    <div class="axis"><div class="axis__top">
      <span class="axis__name">${n}</span>
      <span class="axis__val">${got} / ${max}点</span></div>
      <div class="bar"><i style="width:${Math.round((got / max) * 100)}%"></i></div></div>`).join("");

  const p = r.ratios.periods;
  const U = '<span class="unit-tag">百万円</span>', PC = '<span class="unit-tag">％</span>';
  const figs = [
    ["売上高", U, yen(cur.sales), yen(r.prev.sales), yen(r.prev2.sales)],
    ["経常利益", U, yen(cur.ordinaryProfit), yen(r.prev.ordinaryProfit), yen(r.prev2.ordinaryProfit)],
    ["当期純利益", U, yen(cur.netProfit), yen(r.prev.netProfit), yen(r.prev2.netProfit)],
    ["自己資本比率", PC, pct(p[0].equityRatio), pct(p[1].equityRatio), pct(p[2].equityRatio)],
    ["売上高経常利益率", PC, pct(p[0].ordinaryMargin), pct(p[1].ordinaryMargin), pct(p[2].ordinaryMargin)],
    ["有利子負債", U, yen(cur.interestBearingDebt), yen(r.prev.interestBearingDebt), yen(r.prev2.interestBearingDebt)],
    ["簡易キャッシュフロー", U, yen(cur.simpleCF), yen(r.prev.simpleCF), yen(r.prev2.simpleCF)],
  ].map(([n, u, a, b, c]) =>
    `<tr><th>${n}${u}</th><td>${a}</td><td>${b}</td><td>${c}</td></tr>`).join("");

  return `
  <div class="calc">
    <h2>判定結果</h2>
    <div class="result${cls}">
      <div class="result__label">信用程度</div>
      <div class="result__grade">${s.rank}</div>
      <div class="result__score"><b>${s.total}</b> / 100点</div>
      <p style="margin:10px 0 0;">${POLICY[s.rank]}</p>
    </div>
    <div class="axes">${axes}</div>
  </div>

  <div class="calc">
    <h2>与信限度額の目安</h2>
    <p class="calc__hint">自己資本を基準にした金額と月商を基準にした金額のうち、小さいほうです。一次スクリーニングの出発点としてお使いください。</p>
    <div class="result${cls}" style="text-align:center;">
      <div class="result__score"><b>${yen(r.creditLimit.value)}</b> 百万円<span class="unit-tag">単位：百万円</span></div>
    </div>
  </div>

  <div class="calc">
    <h2>財務ハイライト（直近3期）</h2>
    <p class="calc__hint">金額の単位は百万円です。比率は％で表示しています。</p>
    <table class="figs"><thead><tr><th>項　目<span class="unit-tag">百万円 ／ ％</span></th><th>今期</th><th>前期</th><th>前々期</th></tr></thead>
      <tbody>${figs}</tbody></table>
    <p class="calc__hint" style="margin-top:10px;">
      自己資本比率の業種基準は ${pct(bm.equityRatio)}、売上高経常利益率の業種基準は ${pct(bm.ordinaryMarginAvg3)} です
      （${esc(r.input.industry.trim())}／${esc(r.input.capitalTier)}）。
      債務償還年数は ${red.simpleCF <= 0 ? "算定不能（返済原資なし）" : red.years.toFixed(1) + "年"} です。</p>
  </div>

  <div class="calc">
    <h2>自動所見</h2>
    <p class="calc__hint">該当する項目だけが表示されます。稟議書の所見欄にそのままお使いいただけます。</p>
    <p style="font-weight:700;margin:14px 0 6px;">評価できる点</p>
    <ul class="remarks">${r.comments.strengths.map((t) => `<li>${esc(t)}</li>`).join("")}</ul>
    <p style="font-weight:700;margin:18px 0 6px;">留意すべき点</p>
    <ul class="remarks">${r.comments.concerns.map((t) => `<li>${esc(t)}</li>`).join("")}</ul>
  </div>`;
}

/* -------------------------------------------------------------- ダウンロード */
async function onDownload() {
  const btn0 = $("btnXlsx");
  $("dlNote").textContent = "確認しています…";
  // ボタンの表示状態だけに頼らず、実行の直前にもう一度確認する。
  // このとき会社名の指紋を必ず一緒に送る。送り忘れると
  // 「別の会社に使い回そうとしている」と判定され、解錠が取り消されてしまう。
  const { state: st, order, reason, expiresAt } =
    await checkLicense(await companyFingerprint(state_name()));
  if (st !== "licensed") {
    licensed = false;
    showLicenseDiag(st, order, reason, expiresAt);
    $("dlNote").textContent = "";
    showGate(st === "offline" ? "gateOffline" : "gateBuy");
    return;
  }
  $("dlNote").textContent = "";
  const r = evaluate(state);
  if (r.fy.some((x) => Math.abs(x.balanceCheck) > 0.5) &&
      !confirm("貸借対照表の検算が0になっていません。このまま出力しますか？")) return;
  const btn = $("btnXlsx");
  btn.disabled = true;
  const label = btn.textContent;
  btn.textContent = "Excelを作成中…";
  $("dlNote").textContent = "ファイルを組み立てています。数秒かかります。";
  try {
    await downloadXlsx(r);
    $("dlNote").textContent = "ダウンロードしました";
    if (window.gtag) gtag("event", "xlsx_download", { tool: "credit-pro" });
  } catch (e) {
    $("dlNote").textContent = "作成に失敗しました：" + e.message;
  } finally {
    btn.disabled = false;
    btn.textContent = label;
  }
}


/* ==========================================================================
 * 決算書PDFの読み取り
 * ・pdf.js はサイト内に置いてある（外部CDNを使わない）
 * ・読み取りはすべてブラウザ内。ファイルはどこにも送信しない
 * ・読み取った値は「そのまま反映」せず、必ず確認画面を挟む
 * ======================================================================== */
const PERIOD_LABEL = ["今期（直近）", "前期", "前々期"];
let pdfjsLib = null;
let pending = null;      // 確認待ちの読み取り結果
let accepted = null;     // 判定済みの読み取り結果（次のPDFを足せるよう保持する）
let metaCompany = null;  // PDFから読み取れた会社名
let metaCompanySeen = [];// 読み込んだPDFに出てきた会社名（取り違え検知用）
let missingCells = {};   // 読み取れなかった項目 { キー: [期のindex] }

async function getPdfjs() {
  if (pdfjsLib) return pdfjsLib;
  pdfjsLib = await import("./vendor/pdf.min.mjs");
  pdfjsLib.GlobalWorkerOptions.workerSrc = new URL("./vendor/pdf.worker.min.mjs", import.meta.url).href;
  return pdfjsLib;
}

function initUploader() {
  const zone = $("dropZone"), input = $("fileInput");
  zone.addEventListener("click", () => input.click());
  zone.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); input.click(); } });
  input.addEventListener("change", () => { if (input.files.length) readFiles([...input.files]); });
  ["dragenter", "dragover"].forEach((ev) =>
    zone.addEventListener(ev, (e) => { e.preventDefault(); zone.classList.add("is-over"); }));
  ["dragleave", "drop"].forEach((ev) =>
    zone.addEventListener(ev, (e) => { e.preventDefault(); zone.classList.remove("is-over"); }));
  zone.addEventListener("drop", (e) => {
    const fs = [...(e.dataTransfer?.files || [])].filter((f) => /\.pdf$/i.test(f.name));
    if (fs.length) readFiles(fs);
  });
  $("btnApply").addEventListener("click", applyRead);
  // 最初からやり直す：手入力した内容も含めて、まっさらな状態に戻す。
  // 前回の入力が残っていると、次の会社の判定に混ざって事故になる。
  $("btnReread").addEventListener("click", () => {
    if (!confirm("入力した内容をすべて消して、最初からやり直します。よろしいですか？")) return;
    try { sessionStorage.removeItem("kazumono.credit-pro.draft"); } catch (err) { /* noop */ }
    location.reload();
  });
  // PDFが無い場合の導線
  $("btnNoPdf").addEventListener("click", () => {
    showStep("step3", true);
    $("readBanner").innerHTML =
      '<div class="read-banner is-supp"><b>数値を直接ご入力ください</b>' +
      '<span>下の入力欄を上から順に埋めて、いちばん下の「この内容で判定する」を押してください。' +
      '決算書PDFをお持ちの場合は、上に置いていただくほうが早く済みます。</span></div>';
    $("readTable").innerHTML = "";
    $("readWarn").innerHTML = "";
    const fx = $("readFix"); if (fx) { fx.innerHTML = ""; fx.hidden = true; }
    $("btnApply").hidden = false;
    openManual();
    $("step3").scrollIntoView({ behavior: "smooth", block: "start" });
  });
  // 手入力欄の開閉
  $("btnManual").addEventListener("click", () => {
    const body = $("manualBody"), open = body.hidden;
    body.hidden = !open;
    $("btnManual").setAttribute("aria-expanded", String(open));
    $("btnManual").textContent = open ? "入力欄を閉じる" : "入力欄を開く";
  });
}

function setProgress(msg, ratio) {
  $("readState").hidden = false;
  $("readMsg").textContent = msg;
  $("progBar").style.width = Math.round((ratio || 0) * 100) + "%";
}

async function readFiles(files) {
  showStep("step3", false);
  try {
    const pdfjs = await getPdfjs();
    // 前回までに読み取れた期を引き継ぐ。
    // 「第10期、第9期…」と1件ずつ投げても積み上がるようにするため、毎回捨てない。
    const periods = (pending || accepted || []).slice();
    for (let fi = 0; fi < files.length; fi++) {
      const f = files[fi];
      setProgress(`${f.name} を読み取っています…`, 0);
      const buf = await f.arrayBuffer();
      const found = await scanPdf(pdfjs, buf, (pno, total) =>
        setProgress(`${f.name} を読み取っています…（${pno} / ${total} ページ）`, pno / total));
      // スキャン画像PDF（テキストレイヤーが無い）は、その旨を明示する
      if (found._noText) { periods.push({ file: f.name, failed: true, image: true }); continue; }
      if (found._outOfScope) { periods.push({ file: f.name, failed: true, scope: found._outOfScope }); continue; }
      if (!found.BS && !found.PL) { periods.push({ file: f.name, failed: true }); continue; }
      // 有報・短信は1本で2期分（当期・前期）取れる
      const isTwoYear = (found.BS || found.PL).kind === "years";
      const yis = isTwoYear ? [-1, 0] : [-1];
      const per = found._period || {};
      const unit = found._unit || null;
      if (found._company) {
        if (!metaCompany) metaCompany = found._company;
        if (!metaCompanySeen.includes(found._company)) metaCompanySeen.push(found._company);
      }
      // 同じファイル、または同じ決算期をすでに読み取っていれば置き換える（重複防止）
      const sameIdx = periods.findIndex((q) =>
        q.file === f.name ||
        (q.period && per.end && q.period.end === per.end) ||
        (q.period && typeof per.no === "number" && q.period.no === per.no && !per.end && !q.period.end));
      if (sameIdx >= 0) periods.splice(sameIdx, 1);
      for (const yi of yis) {
        const raw = buildPeriod(found, yi);
        const values = scaleToMillion(raw.values, unit);
        const source = raw.source, warnings = raw.warnings.slice();
        if (unit && unit.toMillion !== 1)
          warnings.push(`この決算書は${unit.label}単位で記載されていたため、百万円に換算しました。`);
        else if (!unit)
          warnings.push("金額の単位を読み取れませんでした。百万円として扱っています。千円単位の決算書の場合は、下の入力欄で桁をご確認ください。");
        const { diff, messages } = validatePeriod(values);
        const grade = gradePeriod(values, diff, messages);
        // 1本で2期取れる様式（有報・短信）は、yi=0 が前期。並べ替えの鍵をずらしておく
        const key = yi === 0 ? shiftBack(per) : per;
        periods.push({ file: f.name, yi, values, source, warnings, diff, messages, grade,
                       period: key, unit,
                       pages: Object.entries(found).map(([k, v]) => `${k} ${v.page}ページ`).join(" / ") });
      }
    }
    setProgress("読み取りが終わりました。", 1);
    sortPeriods(periods);
    // 3期を超えたぶんは、古いものから落とす（並べ替え済みなので後ろが古い）
    const live = periods.filter((p) => !p.failed);
    const kept = live.slice(0, 3);
    const dropped = live.length - kept.length;
    const failed = periods.filter((p) => p.failed);
    const merged = kept.concat(failed);
    if (dropped > 0) merged._dropped = dropped;
    showRead(merged);
  } catch (e) {
    setProgress("読み取りに失敗しました：" + e.message, 0);
  }
}






/**
 * 複数の会社名が混ざっていないかを知らせる。
 * 別会社の決算書を1つの判定に混ぜると、まったく意味のない数字が出るため、
 * 気づけるように必ず画面に出す。
 */
function warnCompany() {
  const box = $("companyWarn"), btn = $("btnApply");
  if (!box) return;
  const mixed = metaCompanySeen.length > 1;
  if (mixed) {
    // 別会社の決算書が混ざったままでは、出てくる数字に意味が無い。
    // 警告だけでは押されてしまうため、判定そのものを止める。
    box.hidden = false;
    box.classList.add("tip--warn");
    box.innerHTML = `<b>別の会社の決算書が混ざっています。このままでは判定できません</b><span>` +
      `読み取れた会社名：<b>${metaCompanySeen.map(esc).join("</b> と <b>")}</b>。` +
      `別々の会社の数字を1つにまとめても、結果は意味を持ちません。` +
      `下の「やり直す」で読み取りを消してから、1社分だけを置き直してください。</span>`;
    if (btn) { btn.disabled = true; btn.textContent = "会社が混ざっているため判定できません"; }
  } else {
    if (btn) { btn.disabled = false; btn.textContent = "この内容で判定する"; }
    if (metaCompany) {
      box.hidden = false;
      box.classList.remove("tip--warn");
      box.innerHTML = `<b>会社名を自動で入れました：${esc(metaCompany)}</b>` +
        `<span>PDFから読み取った名前です。誤っていれば、下の「会社の基本情報」で直してください。</span>`;
    } else {
      box.hidden = true;
    }
  }
}

/* --------------------------------------------------- Excel見本スライダー */
const SHOTS = [
  ["./assets/1-summary.jpg", "①判定サマリー",
   "総合評点・信用程度A〜E・6軸の評点内訳・与信限度額の目安・財務ハイライト・自動所見までを1枚に収めています。"],
  ["./assets/2-financial.jpg", "②財務分析",
   "損益計算書と貸借対照表の3期比較、主要財務指標14種、運転資金分析。業種基準との対比つきです。"],
  ["./assets/3-repayment.jpg", "③資金償還表",
   "簡易キャッシュフローから債務償還年数を算定し、今後3年の約定返済に返済原資が足りるかを見ます。"],
  ["./assets/4-scoring.jpg", "④配点内訳",
   "6軸それぞれの得点と、判定に用いた値。なぜその点数になったのかを稟議で説明できます。"],
  ["./assets/5-input.jpg", "⑤入力データ",
   "判定に使った数値をそのまま記録。あとから検証・引き継ぎができます。"],
];
let shotAt = 0;

function paintShot() {
  const [src, , cap] = SHOTS[shotAt];
  const img = $("shotImg");
  img.src = src;
  img.alt = SHOTS[shotAt][1] + "シートの見本";
  $("shotCap").textContent = cap;
  document.querySelectorAll(".shot__tabs button").forEach((b, i) => {
    b.classList.toggle("is-on", i === shotAt);
    b.setAttribute("aria-selected", String(i === shotAt));
  });
}

function initShots() {
  const tabs = document.querySelector(".shot__tabs");
  if (!tabs) return;
  tabs.addEventListener("click", (e) => {
    const b = e.target.closest("[data-shot]");
    if (b) { shotAt = +b.dataset.shot; paintShot(); }
  });
  $("shotPrev").addEventListener("click", () => { shotAt = (shotAt + SHOTS.length - 1) % SHOTS.length; paintShot(); });
  $("shotNext").addEventListener("click", () => { shotAt = (shotAt + 1) % SHOTS.length; paintShot(); });
  paintShot();
}

/* ------------------------------------------------------- 金額単位の換算 */
// 本シートは百万円で計算する。決算書は千円単位が多く、換算しないと
// 規模の配点と与信限度額が1000倍ずれる。桁が大きいだけで数字は自然に見えるため気づけない。
const MONEY_KEYS = ["sales","cogs","sga","nonOpInc","nonOpExp","extraInc","extraExp","tax",
  "depreciation","cash","receivables","inventory","otherCurrentAssets","tangible",
  "otherFixedAssets","deferred","payables","shortDebt","otherCurrentLiab","longDebt",
  "otherFixedLiab","equity","currentAssets","fixedAssets","currentLiab","fixedLiab",
  "totalCapital","ordinary","operating","net"];

/** 読み取った1期分を百万円に揃える。単位が読めなければ触らない */
function scaleToMillion(values, unit) {
  if (!unit || unit.toMillion === 1) return values;
  const out = { ...values };
  for (const k of MONEY_KEYS)
    if (typeof out[k] === "number") out[k] = Math.round(out[k] * unit.toMillion);

  // 各項目を個別に四捨五入すると、内訳の合計が小計と1単位ずれることがある。
  // 画面の表は内訳を足して小計を出すため、そのままだと決算書では合っていた貸借が
  // 合わなくなり、赤い警告が出てしまう。端数は「その他」の欄に寄せて辻褄を合わせる。
  reconcile(out);
  return out;
}

/**
 * 換算で生じた端数を「その他◯◯」に吸収させ、内訳の合計＝小計＝貸借一致に揃える。
 * 元の決算書で合っていたものを、換算のせいで狂わせないための処理。
 */
function reconcile(v) {
  const n = (x) => (typeof x === "number" ? x : 0);
  const fit = (parts, sub, slack) => {
    // 小計が取れていなければ、内訳の合計をそのまま小計とする
    if (typeof v[sub] !== "number") { v[sub] = parts.reduce((a, k) => a + n(v[k]), 0); return; }
    const gap = v[sub] - parts.reduce((a, k) => a + n(v[k]), 0);
    if (gap === 0 || Math.abs(gap) > 3) return;   // 大きなズレは読み取り誤りなので触らない
    v[slack] = n(v[slack]) + gap;                  // 端数は「その他」で調整する
  };
  fit(["cash", "receivables", "inventory", "otherCurrentAssets"], "currentAssets", "otherCurrentAssets");
  fit(["tangible", "otherFixedAssets"], "fixedAssets", "otherFixedAssets");
  fit(["payables", "shortDebt", "otherCurrentLiab"], "currentLiab", "otherCurrentLiab");
  fit(["longDebt", "otherFixedLiab"], "fixedLiab", "otherFixedLiab");

  // 最後に資産側と負債・純資産側を突き合わせ、残った端数もその他流動資産に寄せる
  const assets = n(v.currentAssets) + n(v.fixedAssets) + n(v.deferred);
  const liabEq = n(v.currentLiab) + n(v.fixedLiab) + n(v.equity);
  const gap = assets - liabEq;
  if (gap !== 0 && Math.abs(gap) <= 3) {
    v.otherCurrentAssets = n(v.otherCurrentAssets) - gap;
    v.currentAssets = n(v.currentAssets) - gap;
  }
}

/* ----------------------------------------------------------- 画面の段 */
function showStep(id, on) { const el = $(id); if (el) el.hidden = !on; }

/** 手入力欄を開き、判定と決済の段も見えるようにする */
function openManual() {
  const body = $("manualBody");
  if (body && body.hidden) {
    body.hidden = false;
    $("btnManual").setAttribute("aria-expanded", "true");
    $("btnManual").textContent = "入力欄を閉じる";
  }
}

/* -------------------------------------------------- 決算期の並べ替え */
/** 有報の「前期」列を、決算期の鍵として1年ぶん戻す */
function shiftBack(per) {
  const out = { end: null, no: null, label: "" };
  if (per.end) {
    const d = per.end.split("-");
    out.end = `${+d[0] - 1}-${d[1]}-${d[2]}`;
    out.label = `${+d[0] - 1}年${+d[1]}月期`;
  }
  if (typeof per.no === "number") {
    out.no = per.no - 1;
    out.label = out.end ? `第${out.no}期（${out.label}）` : `第${out.no}期`;
  }
  return out;
}

/**
 * 読み取れた期を新しい順（今期→前期→前々期）に並べ替える。
 * 決算日が取れたものを優先し、無ければ期数で比べる。
 * どちらも読めない期が混ざっている場合は、取り違えを避けるため並べ替えを行わず投入順のままにする。
 */
function sortPeriods(periods) {
  const live = periods.filter((p) => !p.failed);
  if (live.length < 2) return;
  const keyed = live.every((p) => p.period && (p.period.end || typeof p.period.no === "number"));
  if (!keyed) return;
  const val = (p) => (p.period.end ? p.period.end : "") ;
  const sorted = live.slice().sort((a, b) => {
    const ea = val(a), eb = val(b);
    if (ea && eb && ea !== eb) return eb.localeCompare(ea);
    const na = a.period.no ?? -Infinity, nb = b.period.no ?? -Infinity;
    if (na !== nb) return nb - na;
    return 0;
  });
  // 元配列の live 部分だけを並べ替え後の順に置き換える
  let i = 0;
  for (let k = 0; k < periods.length; k++)
    if (!periods[k].failed) periods[k] = sorted[i++];
}

/* ------------------------------------------------------------ 判定と手入力 */
// 抽出の成否を左右する中核項目。これが欠けたら「要確認」とする
const CORE_KEYS = ["sales", "cash", "currentAssets", "fixedAssets", "currentLiab", "fixedLiab", "equity"];

/**
 * 1期分の読み取り結果を3段階で評価する。
 *   "ok"    … 貸借が合い、中核項目も減価償却費も揃っている（そのまま使える）
 *   "supp"  … 抽出は成功。ただし決算書に単独で載っていない項目（減価償却費など）だけ手入力が要る
 *   "check" … 貸借不一致、経常利益の不整合、または中核項目が取れていない（数値の確認・入力が要る）
 */
function gradePeriod(values, diff, messages) {
  const coreMiss = CORE_KEYS.some((k) => values[k] === null || values[k] === undefined);
  const plNg = messages.some((m) => m.includes("経常利益"));
  if (diff !== 0 || plNg || coreMiss) return "check";
  if (values.depreciation === null || values.depreciation === undefined) return "supp";
  return "ok";
}

// 手入力欄に出す候補。
// 出す基準は「未取得だと判定結果が狂う項目」に限る。
// 決算書に載っていない/合計に吸収済みの内訳（売上原価の区分がない様式など）まで
// 欄にすると、「そのまま使えます」と言いながら大量の入力を求める矛盾した画面になる。
//   depreciation … ⑥償還余力（30点）の算定に必須
//   CORE_KEYS     … 貸借・規模・資本構成の土台。欠けると判定そのものが立たない
//   shortDebt / longDebt … 有利子負債。0で確定できないため、欠けたら必ず確認を求める
const MANUAL_LABELS = {
  depreciation: ["減価償却費", "販管費明細・製造原価報告書の合計。決算書に単独の行が無いことが多い項目です。"],
  sales: ["売上高", ""], cash: ["現金・預金", ""],
  currentAssets: ["流動資産合計", ""], fixedAssets: ["固定資産合計", ""],
  currentLiab: ["流動負債合計", ""], fixedLiab: ["固定負債合計", ""],
  equity: ["純資産合計", ""],
  shortDebt: ["短期借入金（1年内返済分を含む）", ""], longDebt: ["長期借入金・社債", ""],
};
// 表示順（重要な順）
const MANUAL_ORDER = ["depreciation", "sales", "cash", "currentAssets", "fixedAssets",
                      "currentLiab", "fixedLiab", "equity", "shortDebt", "longDebt"];

/** その期で実際に入力を求めるべきキーを返す */
function manualKeysFor(p) {
  const nil = (k) => p.values[k] === null || p.values[k] === undefined;
  return MANUAL_ORDER.filter((k) => {
    if (k === "shortDebt" || k === "longDebt") {
      // buildPeriod は見つからなければ0を入れ、警告を出す。
      // 「無借金なのか読み落としなのか」を利用者に確かめてもらう
      return p.warnings.some((w) => w.includes("有利子負債"));
    }
    return nil(k);
  });
}

const ROWS_SHOW = [
  ["売上高", "sales"], ["売上原価", "cogs"], ["販売費及び一般管理費", "sga"],
  ["営業外収益", "nonOpInc"], ["営業外費用", "nonOpExp"], ["法人税等", "tax"],
  ["減価償却費", "depreciation"], ["現金・預金", "cash"], ["受取手形・売掛金", "receivables"],
  ["棚卸資産", "inventory"], ["流動資産合計", "currentAssets"], ["固定資産合計", "fixedAssets"],
  ["支払手形・買掛金", "payables"], ["短期借入金", "shortDebt"], ["流動負債合計", "currentLiab"],
  ["長期借入金・社債", "longDebt"], ["固定負債合計", "fixedLiab"], ["純資産合計", "equity"],
];

function bannerFor(worst, hasFields) {
  if (worst === "ok")
    return hasFields
      ? ["good", "読み取れました", "貸借は一致しています。下の確認欄だけ目を通してから「反映する」を押してください。"]
      : ["good", "✓ そのまま使えます", "読み取った数値をご確認のうえ、下の「反映する」を押してください。"];
  if (worst === "supp")
    return ["supp", "あと少しで完成します", "決算書に単独で載っていない項目だけ、下の欄にご入力ください。入力すると判定とExcelが完成します。"];
  return ["check", "数値の確認・入力をお願いします", "読み取れなかった項目を下の欄に入力すると完成します。貸借がずれている場合は、各項目の値もあわせてご確認ください。"];
}

function showRead(periods) {
  const live = periods.filter((p) => !p.failed);
  const setFix = (html) => {
    const el = $("readFix");
    if (!el) return;
    el.innerHTML = html;
    el.hidden = !html;
  };

  showStep("step3", true);
  if (!live.length) {
    // 全滅：画像PDFかどうかで文言を変え、手入力へ誘導する
    const anyImage = periods.some((p) => p.image);
    const scope = periods.find((p) => p.scope);
    $("readTable").innerHTML = "";
    $("readBanner").innerHTML = scope
      ? `<div class="read-banner is-ng"><b>${scope.scope === "bank" ? "銀行" : "保険会社"}の決算書のようです。このツールの対象外です</b>` +
        `<span>${scope.scope === "bank" ? "銀行" : "保険会社"}の貸借対照表には流動・固定の区分が無く、損益計算書も売上高ではなく経常収益で構成されるため、` +
        `本ツールの判定モデルには載りません。判定に用いる業界基準の統計も金融業・保険業を対象外としています。` +
        `一般事業会社の決算書でお試しください。</span></div>`
      : `<div class="read-banner is-ng"><b>${anyImage ? "この決算書からは文字を取り出せませんでした" : "この様式は読み取れませんでした"}</b>` +
        `<span>下の入力欄に直接ご入力いただければ、判定もExcelの作成も問題なく行えます。</span></div>`;
    setFix("");
    $("readWarn").innerHTML = anyImage
      ? "<li>スキャンされた画像PDFのため、文字を読み取れませんでした（このツールは画像の文字起こし＝OCRは行いません）。お手数ですが、下の入力欄に直接ご入力ください。判定とExcelの作成は問題なく行えます。</li>"
      : "<li>この決算書は自動読み取りに対応していませんでした。お手数ですが、下の入力欄に直接ご入力ください。判定とExcelの作成は問題なく行えます。</li>";
    $("btnApply").hidden = false;   // 手入力だけで判定できるようにする
    openManual();
    return;
  }
  $("btnApply").hidden = false;
  pending = live;
  if (metaCompany) { state.name = metaCompany; paint(); }
  warnCompany();

  // ---- 判定バナー（最も注意の要る期に合わせる） ----
  const rank = { ok: 0, supp: 1, check: 2 };
  const worst = live.reduce((w, p) => (rank[p.grade] > rank[w] ? p.grade : w), "ok");
  // 入力欄が1つでも出るかを先に調べ、バナーの文言と矛盾しないようにする
  const anyFields = live.some((p) => manualKeysFor(p).length || p.diff !== 0);
  const [cls, title, lead] = bannerFor(worst, anyFields);
  $("readBanner").innerHTML =
    `<div class="read-banner is-${cls}"><b>${title}</b><span>${lead}</span></div>`;

  // ---- 読み取り結果の表 ----
  let h = '<thead><tr><th>科　目<span class="unit-tag">単位：百万円</span></th>';
  live.forEach((p, i) => {
    const lab = p.period && p.period.label ? `<br><span class="th-sub">${esc(p.period.label)}</span>` : "";
    // 元の決算書が何円単位だったかも出す。換算したことを隠さない
    const u = p.unit ? `<br><span class="th-sub">原本：${esc(p.unit.label)}</span>` : "";
    h += `<th>${PERIOD_LABEL[i]}${lab}${u}</th>`;
  });
  h += "</tr></thead><tbody>";
  h += "<tr><th>読み取り元</th>" + live.map((p) =>
    `<td style="font-size:12.5px">${esc(p.file)}<br>${esc(p.pages || "")}</td>`).join("") + "</tr>";
  for (const [label, key] of ROWS_SHOW) {
    h += `<tr><th>${label}</th>` + live.map((p) => {
      const v = p.values[key];
      return v === null || v === undefined
        ? '<td class="miss">未取得</td>'
        : `<td class="ok">${yen(v)}</td>`;
    }).join("") + "</tr>";
  }
  h += "<tr><th>貸借の検算</th>" + live.map((p) =>
    `<td class="${p.diff === 0 ? "ok" : "miss"}">${p.diff === 0 ? "一致" : yen(p.diff) + " のズレ"}</td>`).join("") + "</tr>";
  $("readTable").innerHTML = h + "</tbody>";

  // ---- ここだけ入力してください（未取得のエンジン項目だけを欄にする） ----
  let fix = "";
  let depHint = false;
  live.forEach((p, i) => {
    const miss = manualKeysFor(p);
    const unbalanced = p.diff !== 0;
    if (!miss.length && !unbalanced) return;
    if (miss.includes("depreciation")) depHint = true;
    fix += live.length > 1 ? `<p class="fix-period">${PERIOD_LABEL[i]}</p>` : "";
    if (unbalanced)
      fix += `<p class="fix-warn">資産合計と負債・純資産合計が <b>${yen(p.diff)}</b> ずれています。下の項目、または反映後の入力欄で各数値をご確認ください。</p>`;
    if (miss.length) {
      fix += '<div class="pro-2 fix-grid">';
      for (const k of miss) {
        const label = MANUAL_LABELS[k][0];
        fix += `<label class="pro-field"><span>${label}</span>` +
               `<input type="number" step="1" inputmode="numeric" placeholder="未入力" ` +
               `data-fixkey="${k}" data-fixidx="${i}"></label>`;
      }
      fix += "</div>";
    }
  });
  if (fix && depHint)
    fix += '<p class="note-s" style="margin-top:8px;">減価償却費は、損益計算書に単独の行が無い決算書が普通です。' +
           'キャッシュ・フロー計算書、製造原価報告書、販売費及び一般管理費の明細に載っています。無ければ0のままでも作成できますが、償還余力（30点）の判定には必要です。</p>';
  const fixHead = worst === "ok" ? "ここだけご確認ください" : "ここだけ入力してください";
  setFix(fix ? `<h3 class="fix-head">${fixHead}</h3>${fix}` : "");
  // 読み取った時点で入力欄へ流し込む。判定ボタンを押すまで0が並ぶのは分かりにくいため。
  // 不足項目の欄を作り直した後に呼ぶ（前回の入力値を拾わないようにする）。
  applyValues();

  // ---- 補足メッセージ ----
  const msgs = [];
  live.forEach((p, i) => {
    [...new Set([...p.warnings, ...p.messages])].forEach((w) =>
      msgs.push(`${PERIOD_LABEL[i]}：${w}`));
  });
  // 決算期が連続していない場合に知らせる。
  // 例：2026年3月期と2024年3月期だけが読めたとき、間の2025年3月期が抜けたまま
  // 「前期」の欄に前々期の数字が入る。黙って詰めると気づけないため必ず伝える。
  {
    const dated = live.filter((p) => p.period && p.period.end);
    if (dated.length >= 2) {
      const gaps = [];
      for (let i = 0; i < dated.length - 1; i++) {
        const y0 = +dated[i].period.end.slice(0, 4), y1 = +dated[i + 1].period.end.slice(0, 4);
        if (y0 - y1 > 1) gaps.push(`${dated[i + 1].period.label}と${dated[i].period.label}`);
      }
      if (gaps.length)
        msgs.push(`決算期が連続していません（${gaps.join("、")}のあいだが抜けています）。` +
          `間の期のPDFを追加で読み込むか、順番が意図どおりかご確認ください。`);
    }
  }
  if (periods._dropped)
    msgs.push(`4期以上を読み取ったため、新しい3期分だけを残しました（${periods._dropped}期分を除きました）。`);
  if (live.length < 3) {
    const dated = live.filter((p) => p.period && p.period.label).map((p) => p.period.label);
    const got = dated.length === live.length ? `（${dated.join("・")}）` : "";
    msgs.push(`${live.length}期分を読み取りました${got}。3期分そろうと、損益の推移と償還余力まで判定できます。` +
      `別の期の決算書PDFを、この画面にもう一度ドロップしてください。1件ずつでも、まとめてでもかまいません。` +
      `決算期は自動で読み取り、新しい順に並べ替えます。`);
  }
  $("readWarn").innerHTML = msgs.length
    ? msgs.map((m) => `<li>${esc(m)}</li>`).join("")
    : "<li>特に注意すべき点は検出されませんでした。念のため数値をご確認ください。</li>";
}

/**
 * 読み取った数値を入力欄へ反映する。
 * 判定ボタンを押す前に呼ぶので、赤い枠や3期分の数値がその場で見える。
 * 会社名・代表者情報・返済計画など、利用者が自分で入れた項目は消さない。
 */
function applyValues() {
  if (!pending) return;
  // 前回の読み取り値が残っていると、今回読めなかった項目に古い数字が居座る。
  // 3期分すべてを一度0に戻してから入れ直す。
  const blank = emptyInput();
  const KEEP = ["repayment"];   // 返済計画は決算書に載っておらず、利用者が入れたもの
  for (const k of Object.keys(blank))
    if (Array.isArray(blank[k]) && blank[k].length === 3 && !KEEP.includes(k)) state[k] = [0, 0, 0];
  state.terms = ["", "", ""];
  missingCells = {};   // 赤く出す対象を作り直す

  pending.forEach((p, i) => {
    const v = toEngineFields(p.values);
    // 読み取れなかった項目を覚えておき、入力欄を赤くする
    for (const [k, val] of Object.entries(v))
      if (val === null || val === undefined) (missingCells[k] ||= []).push(i);
    for (const [k, val] of Object.entries(v)) {
      if (!Array.isArray(state[k])) state[k] = [0, 0, 0];
      state[k][i] = val ?? 0;
    }
    // 決算期が読み取れていれば、決算期欄にも入れる（判定には使わないが、Excelの見出しになる）
    if (p.period && p.period.label) {
      if (!Array.isArray(state.terms)) state.terms = ["", "", ""];
      state.terms[i] = p.period.label;
    }
  });
  // 会社名。PDFから読めたものを優先して入れる。
  // 別の会社の決算書を続けて読ませたときに、前の会社名が残ったままにならないようにする。
  if (metaCompany) state.name = metaCompany;
  warnCompany();
  paint(); render();
}

/**
 * 「この内容で判定する」を押したとき。
 * 値の流し込みは読み取り直後に済ませてあるので、ここでは上書きしない。
 * 上書きすると、利用者が表で直した数値が読み取り値に戻ってしまう。
 * ここで取り込むのは「ここだけ入力してください」の欄だけ。
 */
function applyRead() {
  document.querySelectorAll("#readFix input[data-fixkey]").forEach((inp) => {
    if (inp.value === "") return;
    const num = parseFloat(inp.value);
    if (Number.isNaN(num)) return;
    const k = inp.dataset.fixkey, i = +inp.dataset.fixidx;
    if (!Array.isArray(state[k])) state[k] = [0, 0, 0];
    state[k][i] = num;
    // 埋まったので赤い表示を解除する
    if (missingCells[k]) missingCells[k] = missingCells[k].filter((x) => x !== i);
  });
  paint(); render();
  if (window.gtag && pending) gtag("event", "pdf_applied", { tool: "credit-pro", periods: pending.length });
  accepted = pending || accepted;   // 追加のPDFを置いたときに積み上げられるよう残す
  $("readState").hidden = true;
  showStep("step4", true);
  showStep("step5", true);
  refreshLicense();
  $("step4").scrollIntoView({ behavior: "smooth", block: "start" });
}

/* ------------------------------------------------------------------ 記入例 */
function demo() {
  const d = emptyInput();
  Object.assign(d, {
    name: "サンプル情報システム株式会社（記入例）", industry: "　情報通信業",
    capitalTier: "1億円以上10億円未満", listing: "未上場",
    founded: "1998-04-01", baseDate: new Date().toISOString().slice(0, 10),
    employees: 320, capital: 150,
    terms: ["2026年3月期", "2025年3月期", "2024年3月期"],
    sales: [10200, 9600, 9100], cogs: [7140, 6816, 6552], sga: [2150, 2080, 2030],
    nonOpInc: [20, 18, 16], nonOpExp: [30, 34, 38],
    extraInc: [0, 0, 0], extraExp: [0, 0, 30], tax: [290, 220, 150],
    depreciation: [260, 250, 240],
    cash: [2600, 2200, 1900], receivables: [1750, 1650, 1560],
    inventory: [320, 300, 290], otherCurrentAssets: [230, 210, 200],
    tangible: [1900, 1880, 1860], otherFixedAssets: [900, 850, 800], deferred: [0, 0, 0],
    payables: [780, 750, 720], shortDebt: [250, 280, 300], otherCurrentLiab: [850, 800, 780],
    longDebt: [700, 900, 1100], otherFixedLiab: [120, 120, 120], equity: [5000, 4240, 3590],
    ceoName: "見本　太郎", ceoAge: 54, industryYears: 26, ceoYears: 15,
    ownHome: "あり", disclosure: "あり", successor: "あり",
    repayment: [250, 230, 220],
    memo: "主力は金融機関向け業務システムの受託開発。上位5社で売上の約6割を占めるが、いずれも長期契約で取引関係は安定。3期連続の増収増益で、実質無借金。",
  });
  return d;
}

init();
