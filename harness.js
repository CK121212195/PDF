/* ============================================================================
 * harness.js — 決算書PDF 読み取り検証ハーネス
 *
 * 目的:
 *   多数の決算書PDFを一括で流し込み、抽出がどれだけ通るかを測定する。
 *
 * 自動での正誤判定:
 *   決算書の正解を人手で用意しなくても、次の2つで抽出の正しさを機械判定できる。
 *     ① 貸借の一致   … 資産側の小計合計 と 負債純資産側の小計合計 が一致するか
 *     ② 経常利益の整合 … 営業利益＋営業外収益−営業外費用 が記載の経常利益と一致するか
 *   この2つは互いに独立した経路の数値なので、両方通れば抽出はほぼ正しい。
 * ========================================================================== */
import { scanPdf, buildPeriod, validatePeriod, toEngineFields } from "./pdf-extract.js";
import { evaluate, emptyInput, INDUSTRIES, CAPITAL_TIERS, LISTING_OPTIONS } from "./engine.js";
import { downloadXlsx } from "./xlsx-export.js";

const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const yen = (n) => (n === null || n === undefined ? "—" : (n < 0 ? "▲" : "") + Math.abs(Math.round(n)).toLocaleString());

let pdfjsLib = null;
async function getPdfjs() {
  if (pdfjsLib) return pdfjsLib;
  pdfjsLib = await import("./vendor/pdf.min.mjs");
  pdfjsLib.GlobalWorkerOptions.workerSrc = new URL("./vendor/pdf.worker.min.mjs", import.meta.url).href;
  return pdfjsLib;
}

/* 抽出の成否を左右する中核項目。これが欠けたら抽出は失敗とみなす */
const CORE = [
  ["sales", "売上高"], ["cash", "現金・預金"], ["currentAssets", "流動資産合計"],
  ["fixedAssets", "固定資産合計"], ["currentLiab", "流動負債合計"],
  ["fixedLiab", "固定負債合計"], ["equity", "純資産合計"],
];
/* 決算書そのものに載っていないことがある項目。手入力で補える */
const SUPPLEMENT = [["depreciation", "減価償却費"]];

let results = [];
let stop = false;

/* -------------------------------------------------------------- 初期化 */
function init() {
  INDUSTRIES.forEach((r) => $("m_industry").add(new Option(r[0].trim(), r[0])));
  CAPITAL_TIERS.forEach((t) => $("m_tier").add(new Option(t[0], t[0])));
  LISTING_OPTIONS.forEach((l) => $("m_listing").add(new Option(l, l)));
  $("m_industry").value = "全産業（金融業、保険業を除く）";
  $("m_tier").value = "1,000万円以上1億円未満";
  $("m_listing").value = "未上場";

  const zone = $("drop"), input = $("file");
  zone.addEventListener("click", () => input.click());
  zone.addEventListener("keydown", (e) => { if (e.key === "Enter") input.click(); });
  input.addEventListener("change", () => input.files.length && run([...input.files]));
  ["dragenter", "dragover"].forEach((ev) => zone.addEventListener(ev, (e) => { e.preventDefault(); zone.classList.add("is-over"); }));
  ["dragleave", "drop"].forEach((ev) => zone.addEventListener(ev, (e) => { e.preventDefault(); zone.classList.remove("is-over"); }));
  zone.addEventListener("drop", (e) => {
    const fs = [...(e.dataTransfer?.files || [])].filter((f) => /\.pdf$/i.test(f.name));
    if (fs.length) run(fs);
  });
  $("btnStop").addEventListener("click", () => { stop = true; });
  $("btnCsv").addEventListener("click", exportCsv);
  $("btnClear").addEventListener("click", () => { results = []; draw(); });
  $("filter").addEventListener("change", draw);
  $("tbl").addEventListener("click", onTableClick);
}

/* ------------------------------------------------------------ 一括処理 */
async function run(files) {
  stop = false;
  $("progWrap").hidden = false;
  $("btnStop").hidden = false;
  const pdfjs = await getPdfjs();

  for (let i = 0; i < files.length; i++) {
    if (stop) break;
    const f = files[i];
    $("prog").style.width = ((i / files.length) * 100).toFixed(1) + "%";
    $("progMsg").textContent = `${i + 1} / ${files.length} 件目：${f.name}`;
    // 画面を更新させるため1フレーム譲る
    await new Promise((r) => setTimeout(r, 0));
    results.push(await one(pdfjs, f));
    if ((i + 1) % 10 === 0 || i === files.length - 1) draw();
  }
  $("prog").style.width = "100%";
  $("progMsg").textContent = stop
    ? `中断しました（${results.length} 件処理済み）`
    : `完了：${results.length} 件`;
  $("btnStop").hidden = true;
  draw();
}

async function one(pdfjs, file) {
  const t0 = performance.now();
  const base = { file: file.name, sizeKB: Math.round(file.size / 1024) };
  try {
    const buf = await file.arrayBuffer();
    const found = await scanPdf(pdfjs, buf);
    if (found._noText) {
      return { ...base, grade: "ng",
               reason: "スキャンされた画像PDFのため、文字を読み取れません（このツールはOCRを行いません）",
               ms: performance.now() - t0, pages: "", periods: [] };
    }
    if (!found.BS && !found.PL) {
      return { ...base, grade: "ng", reason: "財務諸表のページを検出できませんでした",
               ms: performance.now() - t0, pages: "", periods: [] };
    }
    const isTwoYear = (found.BS || found.PL).kind === "years";
    const yis = isTwoYear ? [-1, 0] : [-1];
    const periods = yis.map((yi) => {
      const { values, source, warnings } = buildPeriod(found, yi);
      const { diff, messages } = validatePeriod(values);
      const isNil = (k) => values[k] === null || values[k] === undefined;
      const missing = CORE.filter(([k]) => isNil(k));
      const lacking = SUPPLEMENT.filter(([k]) => isNil(k));
      return { yi, values, source, warnings, diff, messages, missing, lacking };
    });
    const p0 = periods[0];
    const balanced = p0.diff === 0;
    const plOk = !p0.messages.some((m) => m.includes("経常利益"));
    // 抽出そのものが成功したか（貸借が合い、中核項目が揃っているか）で判定する。
    // 決算書に載っていない補助項目の欠落は、抽出の失敗ではない。
    const extracted = balanced && plOk && p0.missing.length === 0;
    const grade = !extracted ? "wn" : (p0.lacking.length ? "sup" : "ok");
    return {
      ...base, grade,
      reason: grade === "ok" ? "" :
        grade === "sup" ? `${p0.lacking.map(([, n]) => n).join("・")}が決算書に見当たりません（手入力が必要）` :
        [!balanced ? "貸借不一致" : "", !plOk ? "経常利益の不整合" : "",
         p0.missing.length ? `未取得 ${p0.missing.map(([, n]) => n).join("・")}` : ""].filter(Boolean).join(" / "),
      pages: Object.entries(found)
        .map(([k, v]) => `${k}${v.page}${v.continued ? "+" + v.continued.join("+") : ""}${v.consolidated ? "(連)" : ""}`)
        .join(" "),
      kind: (found.BS || found.PL).kind, periods, ms: performance.now() - t0,
    };
  } catch (e) {
    return { ...base, grade: "ng", reason: "処理中にエラー：" + e.message,
             ms: performance.now() - t0, pages: "", periods: [] };
  }
}

/* -------------------------------------------------------------- 表示 */
function draw() {
  const n = results.length;
  $("sumCard").hidden = $("resCard").hidden = $("xlCard").hidden = n === 0;
  $("btnCsv").hidden = $("btnClear").hidden = n === 0;
  if (!n) { $("tbl").innerHTML = ""; return; }

  const c = { ok: 0, sup: 0, wn: 0, ng: 0 };
  results.forEach((r) => c[r.grade]++);
  const usable = c.ok + c.sup;
  const bal = results.filter((r) => r.periods[0]?.diff === 0).length;
  const dep = results.filter((r) => r.periods[0] && r.periods[0].values.depreciation != null).length;
  const two = results.filter((r) => r.periods.length > 1).length;
  const avg = (results.reduce((a, r) => a + r.ms, 0) / n / 1000).toFixed(2);

  $("kpi").innerHTML = `
    <div><b>${n}</b><span>投入したファイル</span></div>
    <div><b>${((usable / n) * 100).toFixed(1)}%</b><span>抽出成功（${usable}件）</span></div>
    <div><b>${c.ok}</b><span>そのまま使える</span></div>
    <div><b>${c.sup}</b><span>一部だけ手入力</span></div>
    <div><b>${((c.wn / n) * 100).toFixed(1)}%</b><span>要確認（${c.wn}件）</span></div>
    <div><b>${((c.ng / n) * 100).toFixed(1)}%</b><span>失敗（${c.ng}件）</span></div>
    <div><b>${((bal / n) * 100).toFixed(1)}%</b><span>貸借が一致</span></div>
    <div><b>${((dep / n) * 100).toFixed(1)}%</b><span>減価償却費を取得</span></div>
    <div><b>${two}</b><span>2期分取れたファイル</span></div>
    <div><b>${avg}秒</b><span>1件あたり平均</span></div>`;

  // 想定外の値でも表が空にならないようにする
  const raw = $("filter").value;
  const f = ["ok", "wn", "ng"].includes(raw) ? raw : "all";
  const rows = results.filter((r) => f === "all" || r.grade === f);
  $("cnt").textContent = `${rows.length} 件を表示`;
  let h = `<thead><tr><th style="width:26px"></th><th>ファイル名</th><th>判定</th><th>検出ページ</th><th>様式</th>
    <th class="n">売上高</th><th class="n">経常利益</th><th class="n">純資産</th><th class="n">貸借差</th>
    <th>指摘</th><th style="width:70px"></th></tr></thead><tbody>`;
  rows.forEach((r) => {
    const i = results.indexOf(r);
    const v = r.periods[0]?.values || {};
    const tag = { ok: "t-ok", sup: "t-sp", wn: "t-wn", ng: "t-ng" }[r.grade];
    const lbl = { ok: "成功", sup: "ほぼ成功", wn: "要確認", ng: "失敗" }[r.grade];
    h += `<tr class="${r.grade === "ng" ? "bad" : r.grade === "wn" ? "warn" : ""}">
      <td class="n">${i + 1}</td>
      <td>${esc(r.file)}<br><span class="note">${r.sizeKB}KB / ${(r.ms / 1000).toFixed(1)}秒</span></td>
      <td><span class="tag ${tag}">${lbl}</span></td>
      <td>${esc(r.pages)}</td>
      <td>${r.kind === "years" ? "2期並記" : r.kind === "accounts" ? "左右2段" : r.kind === "single" ? "単段" : "—"}</td>
      <td class="n">${yen(v.sales)}</td><td class="n">${yen(v.ordinary)}</td><td class="n">${yen(v.equity)}</td>
      <td class="n">${r.periods[0] ? yen(r.periods[0].diff) : "—"}</td>
      <td>${esc(r.reason)}${detail(r)}</td>
      <td>${r.periods.length ? `<button class="mini ghost" data-xl="${i}">Excel</button>` : ""}</td>
    </tr>`;
  });
  $("tbl").innerHTML = h + "</tbody>";
}

function detail(r) {
  if (!r.periods.length) return "";
  const p = r.periods[0];
  const lines = [];
  for (const [k, v] of Object.entries(p.values)) {
    if (v === null || v === undefined) continue;
    lines.push(`${k} = ${v.toLocaleString()}${p.source[k] ? `   ← ${p.source[k]}` : ""}`);
  }
  const warn = [...new Set([...p.warnings, ...p.messages])];
  return `<details><summary>詳細</summary><pre>${esc(
    (warn.length ? "【警告】\n" + warn.map((w) => "・" + w).join("\n") + "\n\n" : "") +
    "【抽出値】\n" + lines.join("\n"))}</pre></details>`;
}

/* ------------------------------------------------------------ Excel作成 */
async function onTableClick(e) {
  const btn = e.target.closest("[data-xl]");
  if (!btn) return;
  const r = results[+btn.dataset.xl];
  btn.disabled = true;
  const old = btn.textContent;
  btn.textContent = "作成中";
  try {
    const d = {
      ...emptyInput(),
      name: r.file.replace(/\.pdf$/i, ""),
      industry: $("m_industry").value, capitalTier: $("m_tier").value,
      listing: $("m_listing").value, founded: $("m_founded").value,
      baseDate: new Date().toISOString().slice(0, 10),
      employees: +$("m_emp").value || 0, disclosure: $("m_disc").value,
      terms: ["今期", "前期", "前々期"],
      memo: `検証ハーネスで ${r.file} から自動読み取りした数値です。会社情報は画面で指定した値を使用しています。`,
    };
    r.periods.forEach((p, i) => {
      const v = toEngineFields(p.values);
      for (const [k, val] of Object.entries(v)) {
        if (!Array.isArray(d[k])) d[k] = [0, 0, 0];
        d[k][i] = val ?? 0;
      }
    });
    await downloadXlsx(evaluate(d), `検証_${d.name}.xlsx`);
  } catch (err) {
    alert("Excelの作成に失敗しました：" + err.message);
  } finally {
    btn.disabled = false; btn.textContent = old;
  }
}

/* ---------------------------------------------------------------- CSV */
function exportCsv() {
  const head = ["No", "ファイル名", "判定", "指摘", "検出ページ", "様式", "期数",
    "貸借差", "売上高", "売上原価", "販管費", "営業外収益", "営業外費用", "経常利益",
    "法人税等", "減価償却費", "現金預金", "売上債権", "棚卸資産", "流動資産合計",
    "固定資産合計", "仕入債務", "短期借入金", "流動負債合計", "長期借入金", "固定負債合計",
    "純資産合計", "警告", "秒"];
  const K = ["sales", "cogs", "sga", "nonOpInc", "nonOpExp", "ordinary", "tax", "depreciation",
    "cash", "receivables", "inventory", "currentAssets", "fixedAssets", "payables",
    "shortDebt", "currentLiab", "longDebt", "fixedLiab", "equity"];
  const q = (s) => `"${String(s ?? "").replace(/"/g, '""')}"`;
  const rows = results.map((r, i) => {
    const p = r.periods[0];
    const v = p?.values || {};
    const w = p ? [...new Set([...p.warnings, ...p.messages])].join(" / ") : "";
    return [i + 1, r.file, { ok: "成功", sup: "ほぼ成功", wn: "要確認", ng: "失敗" }[r.grade], r.reason,
      r.pages, r.kind || "", r.periods.length, p ? p.diff : "",
      ...K.map((k) => (v[k] ?? "")), w, (r.ms / 1000).toFixed(2)].map(q).join(",");
  });
  const csv = "\uFEFF" + [head.map(q).join(","), ...rows].join("\r\n");
  const url = URL.createObjectURL(new Blob([csv], { type: "text/csv" }));
  const a = document.createElement("a");
  a.href = url;
  a.download = `pdf検証結果_${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

init();
