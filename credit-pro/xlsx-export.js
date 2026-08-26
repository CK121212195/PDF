/* ============================================================================
 * xlsx-export.js — ¥500版 Excel出力（ExcelJS）
 *
 * 方針:
 *   ・数式を一切書き込まない（値のみ）。Excel版Pro（数式入り）との差別化のため。
 *   ・DATA_業界指標 / DATA_判定基準 は同梱しない。しきい値・配点表は商品価値の核心。
 *   ・配点内訳の「該当区分」「算定根拠」列、資金償還表の解説列も出力しない。
 *   ・ブラウザ内で生成しダウンロードする。サーバーには何も送らない。
 * ========================================================================== */
// ExcelJS はモジュール先頭では読み込まない。
// 先頭で解決すると、読み込みに失敗したときにページ全体が動かなくなるため、
// 実際にExcelを作る瞬間まで遅延させる。
let _ExcelJS = null;
async function getExcelJS() {
  if (_ExcelJS) return _ExcelJS;
  if (typeof window !== "undefined" && window.ExcelJS) return (_ExcelJS = window.ExcelJS);
  try {
    _ExcelJS = (await import("exceljs")).default;   // Node（テスト時）
    return _ExcelJS;
  } catch (e) {
    throw new Error("Excel生成ライブラリを読み込めませんでした。ページを再読み込みしてお試しください。");
  }
}
import { POLICY } from "./engine.js";

/* ------------------------------------------------------------------ 配色 */
const C = {
  navy: "FF1B2A4A", navyD: "FF12203A", steel: "FF2E6E8E", steelL: "FFE8F0F5",
  line: "FFD6DCE5", muted: "FF5A6473", ink: "FF1A1A1A", calc: "FFF2F4F7",
  warn: "FFFDE7E7", ok: "FFE6F4EA", white: "FFFFFFFF",
};
const RANK_C = { A: "FF1B7F5C", B: "FF2E6E8E", C: "FFB8860B", D: "FFC4632B", E: "FFB03A3A" };
const F = "Meiryo UI";
const MONEY = '#,##0;"▲"#,##0;"-"';
const PCT = '0.0%;"▲"0.0%;"-"';
const PT = '0"点"';

const thin = { style: "thin", color: { argb: C.line } };
const BOX = { top: thin, left: thin, bottom: thin, right: thin };

const fill = (argb) => ({ type: "pattern", pattern: "solid", fgColor: { argb } });
const font = (size = 10, bold = false, color = C.ink) => ({ name: F, size, bold, color: { argb: color } });

/* --------------------------------------------------------------- 部品 */
function put(ws, addr, value, o = {}) {
  const c = ws.getCell(addr);
  c.value = value === undefined ? null : value;
  c.font = o.font || font();
  if (o.fill) c.fill = fill(o.fill);
  if (o.numFmt) c.numFmt = o.numFmt;
  c.alignment = o.align || { vertical: "middle", horizontal: "left", indent: 1 };
  if (o.border !== false) c.border = BOX;
  return c;
}
const AL = {
  l: { vertical: "middle", horizontal: "left", indent: 1 },
  c: { vertical: "middle", horizontal: "center" },
  r: { vertical: "middle", horizontal: "right", indent: 1 },
  w: { vertical: "top", horizontal: "left", indent: 1, wrapText: true },
  cw: { vertical: "middle", horizontal: "center", wrapText: true },
};

function band(ws, row, lastCol, text, o = {}) {
  ws.mergeCells(row, 1, row, lastCol);
  const c = ws.getCell(row, 1);
  c.value = "  " + text;
  c.font = font(o.size || 11, true, C.white);
  c.alignment = AL.l;
  for (let i = 1; i <= lastCol; i++) ws.getCell(row, i).fill = fill(o.bg || C.navy);
  ws.getRow(row).height = o.h || 24;
}

function header(ws, row, labels, startCol = 1, h = 22) {
  labels.forEach((t, i) => {
    const c = ws.getCell(row, startCol + i);
    c.value = t;
    c.font = font(9.5, true, C.navy);
    c.fill = fill(C.steelL);
    c.alignment = AL.cw;
    c.border = BOX;
  });
  ws.getRow(row).height = h;
}

const money = (v) => (typeof v === "number" ? Math.round(v) : 0);

/* ==========================================================================
 * 1. 判定サマリー
 * ======================================================================== */
function sheetSummary(wb, r) {
  const ws = wb.addWorksheet("①判定サマリー", { views: [{ showGridLines: false }] });
  ws.columns = [{ width: 3 }, { width: 22 }, { width: 13 }, { width: 12 }, { width: 11 },
                { width: 12 }, { width: 20 }, { width: 18 }, { width: 3 }];
  const s = r.scores, cur = r.cur, bm = r.benchmark, inp = r.input;
  const LC = 9;

  ws.mergeCells("A1:I1");
  put(ws, "A1", " 与信判断検討書", { font: font(16, true, C.white), align: AL.l, border: false });
  for (let i = 1; i <= LC; i++) ws.getCell(1, i).fill = fill(C.navyD);
  ws.getRow(1).height = 34;
  ws.mergeCells("A2:I2");
  put(ws, "A2", "  CREDIT ASSESSMENT SHEET", { font: font(9, false, "FFB9C4D4"), align: AL.l, border: false });
  for (let i = 1; i <= LC; i++) ws.getCell(2, i).fill = fill(C.navyD);

  // 会社情報
  const info = [
    ["会社名", inp.name || "", "評価基準日", inp.baseDate || ""],
    ["業　種", (inp.industry || "").trim(), "資本金階層", inp.capitalTier || ""],
    ["上場区分", `${inp.listing}　／　業歴 ${r.businessYears}年　／　従業員 ${money(inp.employees).toLocaleString()}名`,
      "代表者", inp.ceoName || ""],
  ];
  info.forEach(([k1, v1, k2, v2], i) => {
    const row = 4 + i;
    put(ws, `B${row}`, k1, { font: font(9, true, C.white), fill: C.steel, align: AL.c });
    ws.mergeCells(row, 3, row, 5);
    put(ws, `C${row}`, v1, { font: font(i === 0 ? 12 : 10, i === 0, C.navy), fill: C.white });
    for (let j = 3; j <= 5; j++) { ws.getCell(row, j).fill = fill(C.white); ws.getCell(row, j).border = BOX; }
    put(ws, `F${row}`, k2, { font: font(9, true, C.white), fill: C.steel, align: AL.c });
    ws.mergeCells(row, 7, row, 8);
    put(ws, `G${row}`, v2, { fill: C.white, align: AL.c });
    ws.getCell(row, 8).fill = fill(C.white); ws.getCell(row, 8).border = BOX;
    ws.getRow(row).height = i === 0 ? 26 : 20;
  });

  // 総合判定
  band(ws, 8, LC, "総 合 判 定");
  ws.mergeCells("B9:C11"); ws.mergeCells("D9:E11"); ws.mergeCells("F9:H11");
  put(ws, "B9", `${s.total}点`, { font: { name: F, size: 34, bold: true, color: { argb: C.navy } }, fill: C.steelL, align: AL.c });
  put(ws, "D9", s.rank, { font: { name: F, size: 34, bold: true, color: { argb: C.white } }, fill: RANK_C[s.rank], align: AL.c });
  put(ws, "F9", POLICY[s.rank], { font: font(11, true, C.navy), fill: C.steelL, align: AL.cw });
  for (let row = 9; row <= 11; row++) {
    ws.getRow(row).height = 22;
    for (let j = 2; j <= 8; j++) {
      const c = ws.getCell(row, j);
      if (!c.fill || c.fill.type !== "pattern") c.fill = fill(j >= 4 && j <= 5 ? RANK_C[s.rank] : C.steelL);
      c.border = BOX;
    }
  }
  ws.mergeCells("B12:C12"); put(ws, "B12", "総合評点（100点満点）", { font: font(9, false, C.muted), align: AL.c, border: false });
  ws.mergeCells("D12:E12"); put(ws, "D12", "信用程度（A〜E）", { font: font(9, false, C.muted), align: AL.c, border: false });
  ws.mergeCells("F12:H12"); put(ws, "F12", "取引方針の目安", { font: font(9, false, C.muted), align: AL.c, border: false });

  // 評点内訳
  band(ws, 14, LC, "評点内訳");
  header(ws, 15, ["評価項目", "得点", "満点", "達成率", "評価", "判定に用いた値"], 2);
  ws.getCell(15, 8).fill = fill(C.steelL);
  ws.getCell(15, 8).border = BOX;
  ws.mergeCells(15, 7, 15, 8);
  const items = [
    ["① 業歴", s.gyoreki, 10, `${r.businessYears}年`],
    ["② 資本構成", s.shihon, 12, `自己資本比率 ${(r.equityRatio * 100).toFixed(1)}％\n業種基準 ${(bm.equityRatio * 100).toFixed(1)}％`],
    ["③ 規模", s.kibo, 18, `売上高 ${money(cur.sales).toLocaleString()}百万円\n${inp.listing}`],
    ["④ 損益", s.soneki, 10, s.soneki_detail.pattern],
    ["⑤ 経営者", s.keiei, 20, `業界歴 ${money(inp.industryYears)}年 ／ 開示 ${inp.disclosure}`],
    ["⑥ 償還余力", s.shokan, 30, r.redemption.simpleCF <= 0 ? "返済原資なし" :
      `債務償還 ${r.redemption.years.toFixed(1)}年`],
  ];
  items.forEach(([n, got, max, memo], i) => {
    const row = 16 + i;
    put(ws, `B${row}`, n);
    put(ws, `C${row}`, got, { numFmt: PT, align: AL.c, font: font(10, true, C.navy), fill: C.calc });
    put(ws, `D${row}`, max, { numFmt: PT, align: AL.c, font: font(10, false, C.muted), fill: C.calc });
    put(ws, `E${row}`, got / max, { numFmt: "0%", align: AL.c, fill: C.calc });
    put(ws, `F${row}`, got / max >= 0.8 ? "◎" : got / max >= 0.6 ? "○" : got / max >= 0.4 ? "△" : "×",
        { align: AL.c, font: font(10, true), fill: C.calc });
    ws.mergeCells(row, 7, row, 8);
    put(ws, `G${row}`, memo, { font: font(9, false, C.muted), fill: C.calc, align: AL.w });
    ws.getCell(row, 8).fill = fill(C.calc); ws.getCell(row, 8).border = BOX;
    ws.getRow(row).height = 28;
  });
  const totRow = 22;
  put(ws, `B${totRow}`, "合　計", { font: font(11, true, C.white), fill: C.navy });
  put(ws, `C${totRow}`, s.total, { numFmt: PT, align: AL.c, font: font(12, true, C.white), fill: C.navy });
  put(ws, `D${totRow}`, 100, { numFmt: PT, align: AL.c, font: font(10, true, C.white), fill: C.navy });
  put(ws, `E${totRow}`, s.total / 100, { numFmt: "0%", align: AL.c, font: font(10, true, C.white), fill: C.navy });
  put(ws, `F${totRow}`, s.rank, { align: AL.c, font: font(11, true, C.white), fill: C.navy });
  ws.mergeCells(totRow, 7, totRow, 8);
  put(ws, `G${totRow}`, "", { fill: C.navy });
  ws.getCell(totRow, 8).fill = fill(C.navy); ws.getCell(totRow, 8).border = BOX;
  ws.getRow(totRow).height = 24;

  // 与信限度額（係数は伏せる）
  band(ws, 24, LC, "与信限度額の目安（一次スクリーニング）");
  put(ws, "B25", "★ 与信限度額の目安", { font: font(10, true), fill: C.steelL });
  ws.mergeCells("C25:D25");
  put(ws, "C25", money(r.creditLimit.value), { numFmt: '#,##0" 百万円"', align: AL.c, font: font(14, true, C.navy), fill: C.steelL });
  ws.getCell(25, 4).fill = fill(C.steelL); ws.getCell(25, 4).border = BOX;
  ws.mergeCells("E25:H25");
  put(ws, "E25", "自己資本基準と月商基準のいずれか小さい方（算定係数は非公開）", { font: font(9, false, C.muted), fill: C.steelL, align: AL.w });
  for (let j = 6; j <= 8; j++) { ws.getCell(25, j).fill = fill(C.steelL); ws.getCell(25, j).border = BOX; }
  ws.getRow(25).height = 32;

  // 財務ハイライト
  band(ws, 27, LC, "財務ハイライト（直近3期）");
  header(ws, 28, ["項　目", "今期（直近）", "前期", "前々期", "業種基準", "判　定"], 2);
  ws.getCell(28, 8).fill = fill(C.steelL);
  ws.getCell(28, 8).border = BOX;
  ws.mergeCells(28, 7, 28, 8);
  const p = r.ratios.periods;
  const hl = [
    ["売上高", cur.sales, r.prev.sales, r.prev2.sales, null],
    ["売上総利益", cur.grossProfit, r.prev.grossProfit, r.prev2.grossProfit, null],
    ["営業利益", cur.operatingProfit, r.prev.operatingProfit, r.prev2.operatingProfit, null],
    ["経常利益", cur.ordinaryProfit, r.prev.ordinaryProfit, r.prev2.ordinaryProfit, null],
    ["当期純利益", cur.netProfit, r.prev.netProfit, r.prev2.netProfit, null],
    ["売上高経常利益率", p[0].ordinaryMargin, p[1].ordinaryMargin, p[2].ordinaryMargin, bm.ordinaryMarginAvg3],
    ["自己資本比率", p[0].equityRatio, p[1].equityRatio, p[2].equityRatio, bm.equityRatio],
    ["有利子負債", cur.interestBearingDebt, r.prev.interestBearingDebt, r.prev2.interestBearingDebt, null],
    ["簡易キャッシュフロー", cur.simpleCF, r.prev.simpleCF, r.prev2.simpleCF, null],
  ];
  hl.forEach(([n, a, b, c2, bench], i) => {
    const row = 29 + i;
    const fmt = bench === null ? MONEY : PCT;
    put(ws, `B${row}`, n);
    [a, b, c2].forEach((v, j) => put(ws, `${"CDE"[j]}${row}`, bench === null ? money(v) : v,
      { numFmt: fmt, align: AL.r, fill: C.calc }));
    put(ws, `F${row}`, bench === null ? "－" : bench, { numFmt: bench === null ? undefined : PCT, align: AL.c, fill: C.calc });
    ws.mergeCells(row, 7, row, 8);
    put(ws, `G${row}`, bench === null ? (a > b ? "↗ 増加" : a < b ? "↘ 減少" : "→ 横ばい")
                                      : (a >= bench ? "○ 業種基準以上" : "△ 業種基準未満"),
        { align: AL.c, fill: C.calc, font: font(9.5) });
    ws.getCell(row, 8).fill = fill(C.calc); ws.getCell(row, 8).border = BOX;
  });

  // 自動所見
  let row = 39;
  band(ws, row++, LC, "自動所見");
  const blocks = [["評価できる点", r.comments.strengths, C.ok, "FF1B7F5C"],
                  ["留意すべき点", r.comments.concerns, C.warn, "FFB03A3A"]];
  for (const [title, list, bg, fg] of blocks) {
    ws.mergeCells(row, 2, row, 8);
    put(ws, `B${row}`, title, { font: font(10, true, fg), fill: bg });
    for (let j = 3; j <= 8; j++) { ws.getCell(row, j).fill = fill(bg); ws.getCell(row, j).border = BOX; }
    row++;
    for (const t of list) {
      ws.mergeCells(row, 2, row, 8);
      put(ws, `B${row}`, "・" + t, { fill: C.white, font: font(9.5) });
      for (let j = 3; j <= 8; j++) { ws.getCell(row, j).fill = fill(C.white); ws.getCell(row, j).border = BOX; }
      row++;
    }
  }

  // 定性メモ
  row++;
  band(ws, row++, LC, "定性情報メモ");
  ws.mergeCells(row, 2, row + 2, 8);
  for (let k = 0; k < 3; k++) ws.getRow(row + k).height = 16;
  put(ws, `B${row}`, inp.memo || "（記載なし）", { fill: C.white, align: AL.w, font: font(9.5) });
  for (let rr = row; rr <= row + 2; rr++)
    for (let j = 2; j <= 8; j++) { ws.getCell(rr, j).fill = fill(C.white); ws.getCell(rr, j).border = BOX; }
  row += 4;

  // 決裁欄
  band(ws, row++, LC, "決裁欄");
  ["承認", "作成者", "係　長", "課　長", "部　長", "決　裁"].forEach((t, i) =>
    put(ws, `${"BCDEFG"[i]}${row}`, t, { font: font(9, true, C.navy), fill: C.steelL, align: AL.c }));
  row++;
  for (let j = 2; j <= 7; j++) { ws.getCell(row, j).border = BOX; ws.getCell(row, j).fill = fill(C.white); }
  ws.getRow(row).height = 46;
  row += 2;

  ws.mergeCells(row, 2, row + 2, 8);
  for (let k = 0; k < 3; k++) ws.getRow(row + k).height = 16;
  put(ws, `B${row}`,
    "出典：財務省「年次別法人企業統計調査（令和6年度）」令和7年9月1日公表（統計法に基づく基幹統計調査）\n" +
    "本ファイルは1社分の判定結果です（計算式は含まれません）。配点・しきい値を自社基準に書き換えて何度でもお使いいただける" +
    "数式入りのExcel版「与信判断検討書類 Pro」もご用意しています。\n" +
    "本書は与信判断を支援する一次スクリーニング資料です。最終的な与信判断は貴社の決裁権者が総合的に行ってください。",
    { font: font(8.5, false, C.muted), align: AL.w, border: false });

  ws.pageSetup = { paperSize: 9, orientation: "portrait", fitToPage: true, fitToWidth: 1, fitToHeight: 0,
                   margins: { left: 0.3, right: 0.3, top: 0.4, bottom: 0.4, header: 0, footer: 0 } };
  ws.views = [{ showGridLines: false, state: "frozen", ySplit: 3 }];
  return ws;
}

/* ==========================================================================
 * 2. 財務分析
 * ======================================================================== */
function sheetFinance(wb, r) {
  const ws = wb.addWorksheet("②財務分析", { views: [{ showGridLines: false }] });
  ws.columns = [{ width: 3 }, { width: 32 }, { width: 15 }, { width: 15 }, { width: 15 },
                { width: 13 }, { width: 15 }, { width: 10 }];
  const LC = 8;
  ws.mergeCells("A1:H1");
  put(ws, "A1", " 財務分析（3期比較）", { font: font(14, true, C.white), align: AL.l, border: false });
  for (let i = 1; i <= LC; i++) ws.getCell(1, i).fill = fill(C.navyD);
  ws.getRow(1).height = 30;

  let row = 3;
  const table = (title, rows, cols) => {
    band(ws, row++, LC, title);
    header(ws, row++, cols, 2);
    for (const [name, vals, fmt, bold] of rows) {
      put(ws, `B${row}`, name, { font: font(10, !!bold), fill: bold ? C.steelL : undefined });
      vals.forEach((v, j) => put(ws, `${"CDEFG"[j]}${row}`, v, {
        numFmt: fmt, align: typeof v === "string" ? AL.c : AL.r,
        fill: C.calc, font: font(10, !!bold),
      }));
      row++;
    }
    row++;
  };

  const P = [r.cur, r.prev, r.prev2];
  const m = (k) => [money(P[0][k]), money(P[1][k]), money(P[2][k])];
  table("1. 損益計算書", [
    ["売上高", m("sales"), MONEY, true], ["売上原価", m("cogs"), MONEY],
    ["売上総利益", m("grossProfit"), MONEY, true], ["販売費及び一般管理費", m("sga"), MONEY],
    ["営業利益", m("operatingProfit"), MONEY, true], ["営業外収益", m("nonOpInc"), MONEY],
    ["営業外費用", m("nonOpExp"), MONEY], ["経常利益", m("ordinaryProfit"), MONEY, true],
    ["特別利益", m("extraInc"), MONEY], ["特別損失", m("extraExp"), MONEY],
    ["税引前当期純利益", m("pretaxProfit"), MONEY], ["法人税等", m("tax"), MONEY],
    ["当期純利益", m("netProfit"), MONEY, true], ["減価償却費", m("depreciation"), MONEY],
  ], ["項　目", "今期（直近）", "前期", "前々期"]);

  table("2. 貸借対照表", [
    ["現金・預金", m("cash"), MONEY], ["受取手形・売掛金", m("receivables"), MONEY],
    ["棚卸資産", m("inventory"), MONEY], ["その他流動資産", m("otherCurrentAssets"), MONEY],
    ["流動資産合計", m("currentAssets"), MONEY, true],
    ["有形固定資産", m("tangible"), MONEY], ["無形固定資産・投資その他", m("otherFixedAssets"), MONEY],
    ["固定資産合計", m("fixedAssets"), MONEY, true],
    ["資産合計", m("totalAssets"), MONEY, true],
    ["支払手形・買掛金", m("payables"), MONEY], ["短期借入金", m("shortDebt"), MONEY],
    ["その他流動負債", m("otherCurrentLiab"), MONEY],
    ["流動負債合計", m("currentLiab"), MONEY, true],
    ["長期借入金・社債", m("longDebt"), MONEY], ["その他固定負債", m("otherFixedLiab"), MONEY],
    ["固定負債合計", m("fixedLiab"), MONEY, true], ["負債合計", m("totalLiab"), MONEY, true],
    ["純資産合計", m("equity"), MONEY, true],
    ["負債・純資産合計（総資本）", m("totalCapital"), MONEY, true],
  ], ["項　目", "今期（直近）", "前期", "前々期"]);

  const q = r.ratios.periods, bm = r.benchmark;
  const rr = (k) => [q[0][k], q[1][k], q[2][k]];
  const ind = [
    ["売上高総利益率", rr("grossMargin"), PCT, "－", "業種により差"],
    ["売上高営業利益率", rr("operatingMargin"), PCT, bm.operatingMarginR6, "業種基準比較"],
    ["売上高経常利益率", rr("ordinaryMargin"), PCT, bm.ordinaryMarginAvg3, "業種基準比較"],
    ["総資本経常利益率（ROA）", rr("roa"), PCT, "－", "5％以上"],
    ["総資本回転率", rr("assetTurnover"), '0.00"回"', "－", "1.0回以上"],
    ["自己資本比率", rr("equityRatio"), PCT, bm.equityRatio, "業種基準比較"],
    ["流動比率", rr("currentRatio"), PCT, "－", "120％以上"],
    ["当座比率", rr("quickRatio"), PCT, "－", "90％以上"],
    ["固定比率", rr("fixedRatio"), PCT, "－", "100％以下"],
    ["有利子負債月商倍率", rr("gearingMonths"), '0.00"倍"', "－", "6倍以内"],
    ["債務償還年数", rr("debtYears").map((v) => (v >= 999 ? "算定不能" : v)), '0.0"年"', "－", "10年以内"],
    ["売上債権回転期間", rr("receivableMonths"), '0.0"ヶ月"', "－", "業種により差"],
    ["棚卸資産回転期間", rr("inventoryMonths"), '0.0"ヶ月"', "－", "業種により差"],
    ["買入債務回転期間", rr("payableMonths"), '0.0"ヶ月"', "－", "業種により差"],
  ];
  band(ws, row++, LC, "3. 主要財務指標（14種）");
  header(ws, row++, ["指　標", "今期（直近）", "前期", "前々期", "業種基準", "目　安"], 2);
  for (const [n, vals, fmt, bench, guide] of ind) {
    put(ws, `B${row}`, n);
    vals.forEach((v, j) => put(ws, `${"CDE"[j]}${row}`, v,
      { numFmt: typeof v === "string" ? undefined : fmt, align: AL.c, fill: C.calc }));
    put(ws, `F${row}`, bench, { numFmt: typeof bench === "number" ? PCT : undefined, align: AL.c, fill: C.calc, font: font(10, typeof bench === "number") });
    put(ws, `G${row}`, guide, { align: AL.c, font: font(9, false, C.muted), fill: C.calc });
    row++;
  }
  row++;

  table("4. 運転資金分析", [
    ["受取サイクル（棚卸資産＋売上債権）", rr("inventoryMonths").map((v, i) => v + rr("receivableMonths")[i]), '0.0"ヶ月"'],
    ["支払サイクル（買入債務）", rr("payableMonths"), '0.0"ヶ月"'],
    ["必要運転資金（正常運転資金）", m("workingCapital"), MONEY, true],
    ["簡易キャッシュフロー", m("simpleCF"), MONEY, true],
  ], ["項　目", "今期（直近）", "前期", "前々期"]);

  ws.pageSetup = { paperSize: 9, orientation: "portrait", fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
  return ws;
}

/* ==========================================================================
 * 3. 資金償還表（解説列は出力しない）
 * ======================================================================== */
function sheetRedemption(wb, r) {
  const ws = wb.addWorksheet("③資金償還表", { views: [{ showGridLines: false }] });
  ws.columns = [{ width: 3 }, { width: 46 }, { width: 20 }, { width: 3 }];
  const d = r.redemption;
  ws.mergeCells("A1:D1");
  put(ws, "A1", " 資金償還表（償還余力の算定）", { font: font(14, true, C.white), align: AL.l, border: false });
  for (let i = 1; i <= 4; i++) ws.getCell(1, i).fill = fill(C.navyD);
  ws.getRow(1).height = 30;

  let row = 3;
  const line = (label, value, fmt, bold) => {
    put(ws, `B${row}`, label, { font: font(10, !!bold), fill: bold ? C.steelL : undefined });
    put(ws, `C${row}`, value, { numFmt: typeof value === "string" ? undefined : fmt, align: AL.r, fill: C.calc, font: font(10, !!bold) });
    row++;
  };
  band(ws, row++, 4, "1. 簡易キャッシュフロー（返済原資）");
  line("税引後当期純利益", money(r.cur.netProfit), MONEY);
  line("（＋）減価償却費", money(r.cur.depreciation), MONEY);
  line("（＝）簡易キャッシュフロー", money(d.simpleCF), MONEY, true);
  row++;
  band(ws, row++, 4, "2. 要償還債務");
  line("有利子負債（短期借入金＋長期借入金・社債）", money(d.interestBearingDebt), MONEY);
  line("（－）現金・預金", money(d.cash), MONEY);
  line("（－）正常運転資金（売上債権＋棚卸資産－仕入債務）", money(d.workingCapital), MONEY);
  line("（＝）要償還債務", money(d.required), MONEY, true);
  row++;
  band(ws, row++, 4, "3. 償還余力の判定");
  line("(A) 債務償還年数", d.simpleCF <= 0 && d.required > 0 ? "算定不能" : Number(d.years.toFixed(1)), '0.0"年"');
  line("(A) スコア", d.scoreA, PT, true);
  line("(B) 3年返済充足率", d.ratio < 0 ? "算定不能" : d.ratio > 9 ? "返済予定なし" : Number(d.ratio.toFixed(2)), '0.00"倍"');
  line("(B) スコア", d.scoreB, PT, true);
  put(ws, `B${row}`, "償還余力 合計", { font: font(11, true, C.white), fill: C.navy });
  put(ws, `C${row}`, d.total, { numFmt: PT, align: AL.c, font: font(14, true, C.white), fill: C.navy });
  ws.getRow(row).height = 28;
  row += 2;
  band(ws, row++, 4, "4. 参考指標");
  line("有利子負債月商倍率", Number(d.gearingMonths.toFixed(2)), '0.00"倍"');
  line("手元流動性（月商倍率）", Number(d.liquidityMonths.toFixed(1)), '0.0"ヶ月"');
  line("DSCR（簡易CF ÷ 1年目約定返済額）", d.dscr > 90 ? "－" : Number(d.dscr.toFixed(2)), '0.00"倍"');
  return ws;
}

/* ==========================================================================
 * 4. 配点内訳（該当区分・算定根拠は出力しない）
 * ======================================================================== */
function sheetScore(wb, r) {
  const ws = wb.addWorksheet("④配点内訳", { views: [{ showGridLines: false }] });
  ws.columns = [{ width: 3 }, { width: 5 }, { width: 30 }, { width: 20 }, { width: 10 }, { width: 10 }, { width: 3 }];
  const s = r.scores, inp = r.input, cur = r.cur;
  ws.mergeCells("A1:G1");
  put(ws, "A1", " 配点内訳（100点満点）", { font: font(14, true, C.white), align: AL.l, border: false });
  for (let i = 1; i <= 7; i++) ws.getCell(1, i).fill = fill(C.navyD);
  ws.getRow(1).height = 30;

  header(ws, 3, ["", "評価項目", "判定に用いた値", "得点", "満点"], 2, 24);
  const rows = [
    ["①", "業歴", `${r.businessYears}年`, s.gyoreki, 10, true],
    ["②", "資本構成（自己資本比率）", `${(r.equityRatio * 100).toFixed(1)}％`, s.shihon, 12, true],
    ["③", "規模", "", s.kibo, 18, true],
    ["", "　③-1 業容（売上高）", `${money(cur.sales).toLocaleString()}百万円`, s.kibo_detail.gyoyo, 12],
    ["", "　③-2 年商", `${(cur.sales / 100).toFixed(0)}億円`, s.kibo_detail.nensho, 2],
    ["", "　③-3 上場区分", inp.listing, s.kibo_detail.listing, 2],
    ["", "　③-4 従業員数", `${money(inp.employees).toLocaleString()}人`, s.kibo_detail.employees, 2],
    ["④", "損益", s.soneki_detail.pattern, s.soneki, 10, true],
    ["⑤", "経営者", "", s.keiei, 20, true],
    ["", "　⑤-1 業界歴", `${money(inp.industryYears)}年`, s.keiei_detail.industry, 2],
    ["", "　⑤-2 経営者歴", `${money(inp.ceoYears)}年`, s.keiei_detail.ceo, 2],
    ["", "　⑤-3 持ち家", inp.ownHome, s.keiei_detail.home, 2],
    ["", "　⑤-4 決算書開示姿勢", inp.disclosure, s.keiei_detail.disclosure, 14],
    ["⑥", "償還余力", r.redemption.simpleCF <= 0 ? "返済原資なし" : `債務償還 ${r.redemption.years.toFixed(1)}年`, s.shokan, 30, true],
  ];
  let row = 4;
  for (const [no, name, val, got, max, bold] of rows) {
    put(ws, `B${row}`, no, { align: AL.c, font: font(10, true) });
    put(ws, `C${row}`, name, { font: font(10, !!bold), fill: bold ? C.steelL : undefined });
    put(ws, `D${row}`, val, { align: AL.c, fill: C.calc, font: font(9.5) });
    put(ws, `E${row}`, got, { numFmt: PT, align: AL.c, font: font(10, true, C.navy), fill: bold ? C.steelL : C.calc });
    put(ws, `F${row}`, max, { numFmt: PT, align: AL.c, font: font(10, false, C.muted), fill: C.calc });
    row++;
  }
  put(ws, `C${row}`, "合　計", { font: font(12, true, C.white), fill: C.navy });
  put(ws, `B${row}`, "", { fill: C.navy }); put(ws, `D${row}`, "", { fill: C.navy });
  put(ws, `E${row}`, s.total, { numFmt: PT, align: AL.c, font: font(13, true, C.white), fill: C.navy });
  put(ws, `F${row}`, 100, { numFmt: PT, align: AL.c, font: font(11, true, C.white), fill: C.navy });
  ws.getRow(row).height = 26;
  row++;
  put(ws, `C${row}`, "信用程度（ランク）", { font: font(10, true), fill: C.steelL });
  put(ws, `B${row}`, ""); put(ws, `D${row}`, "");
  put(ws, `E${row}`, s.rank, { align: AL.c, font: font(13, true, C.navy), fill: C.calc });
  put(ws, `F${row}`, "");
  row += 2;
  ws.mergeCells(row, 2, row + 1, 6);
  put(ws, `B${row}`,
    "※ 各評価項目のしきい値・配点表は本ファイルには含まれません。配点を自社の与信方針に合わせて変更したい場合は、" +
    "数式入りのExcel版「与信判断検討書類 Pro」をご利用ください。",
    { font: font(9, false, C.muted), align: AL.w, border: false });
  return ws;
}

/* ==========================================================================
 * 5. 入力データ（記録用）
 * ======================================================================== */
function sheetInput(wb, r) {
  const ws = wb.addWorksheet("⑤入力データ", { views: [{ showGridLines: false }] });
  ws.columns = [{ width: 3 }, { width: 36 }, { width: 16 }, { width: 16 }, { width: 16 }, { width: 3 }];
  const inp = r.input, P = [r.cur, r.prev, r.prev2];
  ws.mergeCells("A1:F1");
  put(ws, "A1", " 入力データ（記録用）", { font: font(14, true, C.white), align: AL.l, border: false });
  for (let i = 1; i <= 6; i++) ws.getCell(1, i).fill = fill(C.navyD);
  ws.getRow(1).height = 30;

  let row = 3;
  band(ws, row++, 6, "会社基本情報");
  for (const [k, v] of [["会社名", inp.name], ["業種", (inp.industry || "").trim()],
      ["資本金階層", inp.capitalTier], ["上場区分", inp.listing], ["設立年月日", inp.founded],
      ["評価基準日", inp.baseDate], ["業歴（年）", r.businessYears], ["期末従業員数（人）", money(inp.employees)],
      ["資本金（百万円）", money(inp.capital)], ["代表者名", inp.ceoName],
      ["業界経験年数", money(inp.industryYears)], ["代表取締役在任年数", money(inp.ceoYears)],
      ["持ち家", inp.ownHome], ["決算書開示姿勢", inp.disclosure], ["後継者の有無", inp.successor]]) {
    put(ws, `B${row}`, k, { fill: C.steelL });
    ws.mergeCells(row, 3, row, 5);
    put(ws, `C${row}`, v ?? "", { fill: C.calc, align: typeof v === "number" ? AL.r : AL.l });
    for (let j = 4; j <= 5; j++) { ws.getCell(row, j).fill = fill(C.calc); ws.getCell(row, j).border = BOX; }
    row++;
  }
  row++;
  band(ws, row++, 6, "決算数値（単位：百万円）");
  header(ws, row++, ["項　目", "今期（直近）", "前期", "前々期"], 2);
  put(ws, `B${row}`, "決算期", { fill: C.steelL });
  (inp.terms || ["", "", ""]).forEach((t, j) => put(ws, `${"CDE"[j]}${row}`, t || "", { align: AL.c, fill: C.calc }));
  row++;
  const keys = [["売上高", "sales"], ["売上原価", "cogs"], ["売上総利益", "grossProfit"],
    ["販売費及び一般管理費", "sga"], ["営業利益", "operatingProfit"], ["営業外収益", "nonOpInc"],
    ["営業外費用", "nonOpExp"], ["経常利益", "ordinaryProfit"], ["特別利益", "extraInc"],
    ["特別損失", "extraExp"], ["税引前当期純利益", "pretaxProfit"], ["法人税等", "tax"],
    ["当期純利益", "netProfit"], ["減価償却費", "depreciation"], ["現金・預金", "cash"],
    ["受取手形・売掛金", "receivables"], ["棚卸資産", "inventory"], ["その他流動資産", "otherCurrentAssets"],
    ["有形固定資産", "tangible"], ["無形固定資産・投資その他", "otherFixedAssets"], ["繰延資産", "deferred"],
    ["資産合計", "totalAssets"], ["支払手形・買掛金", "payables"], ["短期借入金", "shortDebt"],
    ["その他流動負債", "otherCurrentLiab"], ["長期借入金・社債", "longDebt"],
    ["その他固定負債", "otherFixedLiab"], ["純資産合計", "equity"],
    ["負債・純資産合計", "totalCapital"], ["【検算】資産－負債純資産", "balanceCheck"]];
  for (const [name, key] of keys) {
    put(ws, `B${row}`, name);
    [0, 1, 2].forEach((i) => put(ws, `${"CDE"[i]}${row}`, money(P[i][key]),
      { numFmt: MONEY, align: AL.r, fill: key === "balanceCheck" && Math.abs(P[i][key]) > 0.5 ? C.warn : C.calc }));
    row++;
  }
  row++;
  band(ws, row++, 6, "借入金返済計画（今後3年）");
  put(ws, `B${row}`, "年間約定返済額");
  (inp.repayment || [0, 0, 0]).forEach((v, j) => put(ws, `${"CDE"[j]}${row}`, money(v), { numFmt: MONEY, align: AL.r, fill: C.calc }));
  return ws;
}

/* ------------------------------------------------------------------ 公開API */
export async function buildWorkbook(result) {
  const ExcelJS = await getExcelJS();
  const wb = new ExcelJS.Workbook();
  wb.creator = "数字のものさし｜与信判断検討書";
  wb.created = new Date();
  sheetSummary(wb, result);
  sheetFinance(wb, result);
  sheetRedemption(wb, result);
  sheetScore(wb, result);
  sheetInput(wb, result);
  return wb;
}

export async function downloadXlsx(result, filename) {
  const wb = await buildWorkbook(result);
  const buf = await wb.xlsx.writeBuffer();
  const name = filename || `与信判断検討書_${(result.input.name || "無題").replace(/[\\/:*?"<>|]/g, "")}.xlsx`;
  const url = URL.createObjectURL(new Blob([buf],
    { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  const a = document.createElement("a");
  a.href = url; a.download = name; a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
