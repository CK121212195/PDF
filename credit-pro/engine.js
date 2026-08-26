/* ============================================================================
 * engine.js — 与信判断検討書 Pro / 判定エンジン
 *
 * Excel版「与信判断検討書類_Pro.xlsx」の DATA_業界指標 / DATA_判定基準 /
 * ⑤配点内訳 / ④資金償還表 / ②判定サマリー の数式を 1:1 で移植したもの。
 * DOM に依存しないため Node からも読める（パリティテスト用）。
 *
 * 出典（業界指標）:
 *   財務省「年次別法人企業統計調査（令和6年度）」令和7年9月1日公表
 *   https://www.mof.go.jp/pri/reference/ssc/results/r6.pdf
 *   第4表 売上高利益率の推移 / 第12表 自己資本比率の推移
 * ========================================================================== */

/* ---------------------------------------------------------------- 業界指標 */
// [業種名, 大区分, 経常R2, R3, R4, R5, R6, 営業利益率R6]
export const INDUSTRIES = [
  ["全産業（金融業、保険業を除く）", "全産業", 0.046, 0.058, 0.060, 0.065, 0.068, 0.050],
  ["製造業", "製造業", 0.060, 0.083, 0.079, 0.086, 0.086, 0.054],
  ["　食料品", "製造業", 0.036, 0.046, 0.037, 0.047, 0.055, 0.036],
  ["　化学", "製造業", 0.104, 0.128, 0.116, 0.121, 0.123, 0.085],
  ["　石油・石炭", "製造業", 0.036, 0.063, 0.009, 0.027, 0.013, 0.006],
  ["　鉄鋼", "製造業", 0.006, 0.067, 0.065, 0.070, 0.050, 0.037],
  ["　金属製品", "製造業", 0.044, 0.063, 0.049, 0.062, 0.055, 0.040],
  ["　はん用機械", "製造業", 0.068, 0.092, 0.090, 0.118, 0.111, 0.067],
  ["　生産用機械", "製造業", 0.078, 0.094, 0.110, 0.105, 0.112, 0.078],
  ["　業務用機械", "製造業", 0.083, 0.118, 0.134, 0.135, 0.153, 0.091],
  ["　電気機械", "製造業", 0.075, 0.105, 0.107, 0.088, 0.114, 0.071],
  ["　情報通信機械", "製造業", 0.060, 0.102, 0.081, 0.066, 0.086, 0.063],
  ["　輸送用機械", "製造業", 0.056, 0.072, 0.092, 0.113, 0.102, 0.051],
  ["非製造業", "非製造業", 0.041, 0.048, 0.053, 0.058, 0.061, 0.049],
  ["　建設業", "非製造業", 0.052, 0.051, 0.052, 0.050, 0.054, 0.045],
  ["　卸売業、小売業", "非製造業", 0.023, 0.030, 0.034, 0.035, 0.037, 0.024],
  ["　不動産業", "非製造業", 0.121, 0.125, 0.128, 0.130, 0.136, 0.122],
  ["　物品賃貸業", "非製造業", 0.035, 0.026, 0.011, 0.070, 0.074, 0.068],
  ["　情報通信業", "非製造業", 0.094, 0.102, 0.113, 0.112, 0.099, 0.082],
  ["　運輸業、郵便業", "非製造業", -0.015, 0.019, 0.055, 0.059, 0.065, 0.053],
  ["　電気業", "非製造業", 0.036, 0.011, -0.014, 0.070, 0.059, 0.052],
  ["　サービス業", "非製造業", 0.061, 0.074, 0.081, 0.077, 0.091, 0.081],
];

// 第12表 自己資本比率（令和6年度）
export const EQUITY_SEGMENT = { 全産業: 0.421, 製造業: 0.503, 非製造業: 0.392 };
export const CAPITAL_TIERS = [
  ["1,000万円未満", 0.213],
  ["1,000万円以上1億円未満", 0.455],
  ["1億円以上10億円未満", 0.416],
  ["10億円以上", 0.435],
];
const ALL_INDUSTRY_EQUITY = 0.421; // 補正係数の分母

export const LISTING_OPTIONS = ["プライム", "スタンダード・グロース", "未上場"];

/* ------------------------------------------------------------ しきい値表 */
export const T = {
  gyoreki:  { keys: [0, 1, 3, 5, 10, 20],                              vals: [2, 2, 4, 6, 8, 10] },
  shihon:   { keys: [-9999, 0, 0.3, 0.5, 0.7, 0.9, 1.1, 1.3, 1.6],     vals: [0, 1, 3, 5, 7, 9, 10, 11, 12] },
  gyoyo:    { keys: [0, 100, 300, 1000, 3000, 10000, 30000, 100000, 300000, 1000000, 3000000],
              vals: [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12] },
  nensho:   { keys: [0, 1000, 3000],                                   vals: [0, 1, 2] },
  jugyoin:  { keys: [0, 1000, 3000],                                   vals: [0, 1, 2] },
  keireki:  { keys: [0, 3, 10],                                        vals: [0, 1, 2] },
  shokanA:  { keys: [0, 3.0001, 5.0001, 7.0001, 10.0001, 15.0001, 20.0001], vals: [15, 13, 11, 8, 5, 2, 0] },
  shokanB:  { keys: [-9999, 0.5, 0.8, 1, 1.2, 1.3, 1.5],               vals: [0, 2, 5, 8, 11, 13, 15] },
  rank:     { keys: [0, 36, 51, 66, 86],                               vals: ["E", "D", "C", "B", "A"] },
};

export const LISTING_POINTS = { "プライム": 2, "スタンダード・グロース": 1, "未上場": 0 };
export const SONEKI = { threeYear: 7, twoYear: 6, currentOnly: 4, pastOnly: 2, none: 0 };
export const SONEKI_BONUS = { threshold: 50, perPeriod: 1, cap: 10 };
export const KEIEI_POINTS = { home: 2, disclosure: 14 };
export const LIMIT_COEF = {
  A: { equity: 0.30, months: 3.0 },
  B: { equity: 0.20, months: 2.0 },
  C: { equity: 0.10, months: 1.0 },
  D: { equity: 0.05, months: 0.5 },
  E: { equity: 0.00, months: 0.0 },
};

/* -------------------------------------------------------------- ユーティリティ */
// Excel の LOOKUP（昇順ベクトル・以下最大値を拾う）と同じ挙動
export function lookup(value, table) {
  const v = num(value);
  let idx = -1;
  for (let i = 0; i < table.keys.length; i++) if (table.keys[i] <= v) idx = i;
  return table.vals[idx < 0 ? 0 : idx];
}
const num = (x) => (typeof x === "number" && isFinite(x) ? x : parseFloat(x) || 0);
const safeDiv = (a, b, fallback = 0) => (num(b) === 0 ? fallback : num(a) / num(b));

/** DATEDIF(start, end, "y") 相当 */
export function yearsBetween(startISO, endISO) {
  if (!startISO || !endISO) return 0;
  const s = new Date(startISO), e = new Date(endISO);
  if (isNaN(s) || isNaN(e)) return 0;
  let y = e.getFullYear() - s.getFullYear();
  const m = e.getMonth() - s.getMonth();
  if (m < 0 || (m === 0 && e.getDate() < s.getDate())) y--;
  return Math.max(0, y);
}

/* ------------------------------------------------------------- 入力の器 */
export function emptyInput() {
  const p3 = () => [0, 0, 0]; // [今期, 前期, 前々期]
  return {
    name: "", industry: "　金属製品", capitalTier: "1,000万円以上1億円未満",
    listing: "未上場", founded: "", baseDate: new Date().toISOString().slice(0, 10),
    employees: 0, capital: 0,
    terms: ["", "", ""],
    // PL
    sales: p3(), cogs: p3(), sga: p3(), nonOpInc: p3(), nonOpExp: p3(),
    extraInc: p3(), extraExp: p3(), tax: p3(), depreciation: p3(),
    // BS
    cash: p3(), receivables: p3(), inventory: p3(), otherCurrentAssets: p3(),
    tangible: p3(), otherFixedAssets: p3(), deferred: p3(),
    payables: p3(), shortDebt: p3(), otherCurrentLiab: p3(),
    longDebt: p3(), otherFixedLiab: p3(), equity: p3(),
    // 代表者
    ceoName: "", ceoAge: 0, industryYears: 0, ceoYears: 0,
    ownHome: "なし", disclosure: "あり", successor: "なし",
    // 返済計画
    repayment: [0, 0, 0],
    memo: "",
  };
}

/* --------------------------------------------------------------- 決算集計 */
function derive(inp) {
  const at = (arr, i) => num(arr && arr[i]);
  const per = (i) => {
    const sales = at(inp.sales, i), cogs = at(inp.cogs, i), sga = at(inp.sga, i);
    const grossProfit = sales - cogs;
    const operatingProfit = grossProfit - sga;
    const ordinaryProfit = operatingProfit + at(inp.nonOpInc, i) - at(inp.nonOpExp, i);
    const pretaxProfit = ordinaryProfit + at(inp.extraInc, i) - at(inp.extraExp, i);
    const netProfit = pretaxProfit - at(inp.tax, i);

    const currentAssets = at(inp.cash, i) + at(inp.receivables, i)
                        + at(inp.inventory, i) + at(inp.otherCurrentAssets, i);
    const fixedAssets = at(inp.tangible, i) + at(inp.otherFixedAssets, i);
    const totalAssets = currentAssets + fixedAssets + at(inp.deferred, i);
    const currentLiab = at(inp.payables, i) + at(inp.shortDebt, i) + at(inp.otherCurrentLiab, i);
    const fixedLiab = at(inp.longDebt, i) + at(inp.otherFixedLiab, i);
    const totalLiab = currentLiab + fixedLiab;
    const totalCapital = totalLiab + at(inp.equity, i);

    return {
      sales, cogs, grossProfit, sga, operatingProfit,
      nonOpInc: at(inp.nonOpInc, i), nonOpExp: at(inp.nonOpExp, i), ordinaryProfit,
      extraInc: at(inp.extraInc, i), extraExp: at(inp.extraExp, i),
      pretaxProfit, tax: at(inp.tax, i), netProfit, depreciation: at(inp.depreciation, i),
      cash: at(inp.cash, i), receivables: at(inp.receivables, i),
      inventory: at(inp.inventory, i), otherCurrentAssets: at(inp.otherCurrentAssets, i),
      currentAssets, tangible: at(inp.tangible, i),
      otherFixedAssets: at(inp.otherFixedAssets, i), fixedAssets,
      deferred: at(inp.deferred, i), totalAssets,
      payables: at(inp.payables, i), shortDebt: at(inp.shortDebt, i),
      otherCurrentLiab: at(inp.otherCurrentLiab, i), currentLiab,
      longDebt: at(inp.longDebt, i), otherFixedLiab: at(inp.otherFixedLiab, i),
      fixedLiab, totalLiab, equity: at(inp.equity, i), totalCapital,
      balanceCheck: totalAssets - totalCapital,
      interestBearingDebt: at(inp.shortDebt, i) + at(inp.longDebt, i),
      workingCapital: at(inp.receivables, i) + at(inp.inventory, i) - at(inp.payables, i),
      simpleCF: netProfit + at(inp.depreciation, i),
    };
  };
  return [per(0), per(1), per(2)];
}

/* --------------------------------------------------------------- 業界基準 */
export function benchmark(industryName, capitalTier) {
  const row = INDUSTRIES.find((r) => r[0] === industryName) || INDUSTRIES[0];
  const segment = row[1];
  const ordinaryAvg3 = (row[4] + row[5] + row[6]) / 3;   // 直近3期平均（令4〜令6）
  const tier = CAPITAL_TIERS.find((t) => t[0] === capitalTier) || CAPITAL_TIERS[0];
  const coefficient = tier[1] / ALL_INDUSTRY_EQUITY;
  const segmentEquity = EQUITY_SEGMENT[segment] ?? EQUITY_SEGMENT["全産業"];
  return {
    segment,
    ordinaryMarginAvg3: ordinaryAvg3,
    ordinaryMarginR6: row[6],
    operatingMarginR6: row[7],
    segmentEquityRatio: segmentEquity,
    tierCoefficient: coefficient,
    equityRatio: segmentEquity * coefficient,   // ★基準自己資本比率
  };
}

/* ------------------------------------------------------------ 償還余力 */
function redemption(cur, repayment) {
  const simpleCF = cur.simpleCF;
  const required = Math.max(0, cur.interestBearingDebt - cur.cash - cur.workingCapital);
  const years = required <= 0 ? 0 : (simpleCF <= 0 ? 999 : required / simpleCF);
  const scoreA = lookup(years, T.shokanA);

  const repay3 = num(repayment[0]) + num(repayment[1]) + num(repayment[2]);
  const cf3 = simpleCF * 3;
  let ratio;
  if (repay3 <= 0) ratio = simpleCF > 0 ? 9.99 : 0;
  else if (simpleCF <= 0) ratio = -9999;
  else ratio = cf3 / repay3;
  const scoreB = lookup(ratio, T.shokanB);

  const monthlySales = cur.sales / 12;
  return {
    simpleCF, interestBearingDebt: cur.interestBearingDebt, cash: cur.cash,
    workingCapital: cur.workingCapital, required, years, scoreA,
    repay3, cf3, ratio, scoreB, total: scoreA + scoreB,
    gearingMonths: monthlySales === 0 ? 0 : cur.interestBearingDebt / monthlySales,
    liquidityMonths: monthlySales === 0 ? 0 : cur.cash / monthlySales,
    dscr: num(repayment[0]) <= 0 ? 99 : simpleCF / num(repayment[0]),
    grossYears: simpleCF <= 0 ? 999 : cur.interestBearingDebt / simpleCF,
  };
}

/* ------------------------------------------------------------ メイン計算 */
export function evaluate(input) {
  const inp = { ...emptyInput(), ...input };
  const fy = derive(inp);
  const cur = fy[0], prev = fy[1], prev2 = fy[2];
  const bm = benchmark(inp.industry, inp.capitalTier);

  // ① 業歴
  const businessYears = yearsBetween(inp.founded, inp.baseDate);
  const s1 = lookup(businessYears, T.gyoreki);

  // ② 資本構成
  const equityRatio = safeDiv(cur.equity, cur.totalCapital, 0);
  const equityMultiple = bm.equityRatio === 0 ? -9999 : equityRatio / bm.equityRatio;
  const s2 = lookup(equityMultiple, T.shihon);

  // ③ 規模
  const s3_1 = lookup(cur.sales, T.gyoyo);
  const s3_2 = lookup(cur.sales / 100, T.nensho);
  const s3_3 = LISTING_POINTS[inp.listing] ?? 0;
  const s3_4 = lookup(num(inp.employees), T.jugyoin);
  const s3 = s3_1 + s3_2 + s3_3 + s3_4;

  // ④ 損益
  const o = [cur.ordinaryProfit, prev.ordinaryProfit, prev2.ordinaryProfit];
  let base, pattern;
  if (o[0] > 0 && o[1] > 0 && o[2] > 0)      { base = SONEKI.threeYear;  pattern = "3期連続黒字"; }
  else if (o[0] > 0 && o[1] > 0)             { base = SONEKI.twoYear;    pattern = "直近2期黒字"; }
  else if (o[0] > 0)                          { base = SONEKI.currentOnly; pattern = "直近期のみ黒字"; }
  else if (o[1] > 0 || o[2] > 0)             { base = SONEKI.pastOnly;   pattern = "直近期赤字"; }
  else                                        { base = SONEKI.none;       pattern = "3期連続赤字"; }
  const bonus = o.filter((x) => x >= SONEKI_BONUS.threshold).length * SONEKI_BONUS.perPeriod;
  const s4 = Math.min(base + bonus, SONEKI_BONUS.cap);

  // ⑤ 経営者
  const s5_1 = lookup(num(inp.industryYears), T.keireki);
  const s5_2 = lookup(num(inp.ceoYears), T.keireki);
  const s5_3 = inp.ownHome === "あり" ? KEIEI_POINTS.home : 0;
  const s5_4 = inp.disclosure === "あり" ? KEIEI_POINTS.disclosure : 0;
  const s5 = s5_1 + s5_2 + s5_3 + s5_4;

  // ⑥ 償還余力
  const red = redemption(cur, inp.repayment);
  const s6 = red.total;

  const total = s1 + s2 + s3 + s4 + s5 + s6;
  const rank = lookup(total, T.rank);

  // 与信限度額
  const coef = LIMIT_COEF[rank];
  const equityBasis = Math.max(0, cur.equity) * coef.equity;
  const salesBasis = Math.max(0, cur.sales / 12) * coef.months;
  const creditLimit = cur.equity < 0 ? 0 : Math.min(equityBasis, salesBasis);

  return {
    input: inp, fy, cur, prev, prev2, benchmark: bm,
    businessYears, equityRatio, equityMultiple, redemption: red,
    scores: {
      gyoreki: s1, shihon: s2,
      kibo: s3, kibo_detail: { gyoyo: s3_1, nensho: s3_2, listing: s3_3, employees: s3_4 },
      soneki: s4, soneki_detail: { pattern, base, bonus },
      keiei: s5, keiei_detail: { industry: s5_1, ceo: s5_2, home: s5_3, disclosure: s5_4 },
      shokan: s6, total, rank,
    },
    creditLimit: { equityBasis, salesBasis, value: creditLimit, coef },
    ratios: financialRatios(fy, bm),
    comments: buildComments(cur, prev, prev2, inp, equityMultiple, red),
  };
}

/* ---------------------------------------------------------- 財務指標14種 */
function financialRatios(fy, bm) {
  const calc = (p) => {
    const ms = p.sales / 12;
    const required = Math.max(0, p.interestBearingDebt - p.cash - p.workingCapital);
    const debtYears = required <= 0 ? 0 : (p.simpleCF <= 0 ? 999 : required / p.simpleCF);
    return {
      grossMargin: safeDiv(p.grossProfit, p.sales),
      operatingMargin: safeDiv(p.operatingProfit, p.sales),
      ordinaryMargin: safeDiv(p.ordinaryProfit, p.sales),
      roa: safeDiv(p.ordinaryProfit, p.totalCapital),
      assetTurnover: safeDiv(p.sales, p.totalCapital),
      equityRatio: safeDiv(p.equity, p.totalCapital),
      currentRatio: safeDiv(p.currentAssets, p.currentLiab),
      quickRatio: safeDiv(p.cash + p.receivables, p.currentLiab),
      fixedRatio: safeDiv(p.fixedAssets, p.equity),
      gearingMonths: ms === 0 ? 0 : p.interestBearingDebt / ms,
      debtYears,
      receivableMonths: ms === 0 ? 0 : p.receivables / ms,
      inventoryMonths: ms === 0 ? 0 : p.inventory / ms,
      payableMonths: ms === 0 ? 0 : p.payables / ms,
      workingCapital: p.workingCapital,
      simpleCF: p.simpleCF,
    };
  };
  return { periods: fy.map(calc), benchmark: bm };
}

/* ------------------------------------------------------------- 自動所見 */
function buildComments(cur, prev, prev2, inp, equityMultiple, red) {
  const strengths = [], concerns = [];
  const push = (arr, cond, text) => { if (cond) arr.push(text); };

  push(strengths, equityMultiple >= 1,
    "自己資本比率が業種・規模別の基準値を上回っており、財務の安全性は相対的に良好です。");
  push(strengths, cur.ordinaryProfit > 0 && prev.ordinaryProfit > 0 && prev2.ordinaryProfit > 0,
    "直近3期連続で経常黒字を計上しており、収益の安定性が認められます。");
  push(strengths, red.simpleCF > 0 && red.years <= 10,
    "債務償還年数が10年以内であり、有利子負債の返済能力に問題は見られません。");
  push(strengths, red.ratio >= 1.2,
    "今後3年間の約定返済額に対し、キャッシュフローで十分な返済原資を確保できる見込みです。");
  push(strengths, inp.disclosure === "あり",
    "決算書の開示を受けられており、財務実態の把握が可能です。");
  push(strengths, cur.sales > prev.sales && prev.sales > prev2.sales,
    "売上高は3期連続で増収基調にあります。");

  push(concerns, cur.equity < 0,
    "債務超過の状態にあります。原則として新規与信は不可とし、既存与信は保全・回収を優先してください。");
  push(concerns, cur.ordinaryProfit <= 0,
    "直近期の経常損益が赤字です。損益構造の問題点と資金繰りへの影響を確認してください。");
  push(concerns, red.simpleCF > 0 && red.years > 15,
    "債務償還年数が15年を超えており、有利子負債の水準が過大です。");
  push(concerns, red.simpleCF <= 0,
    "簡易キャッシュフローがマイナスです。返済原資が確保できておらず、資金繰りに注意が必要です。");
  push(concerns, red.ratio < 1,
    "今後3年間の約定返済額に対しキャッシュフローが不足する見込みです。返済計画の見直し状況を確認してください。");
  push(concerns, inp.disclosure === "なし",
    "決算書が非開示です。財務実態が確認できないため、与信は特に慎重に判断してください。");
  push(concerns, Math.abs(cur.balanceCheck) > 0.5,
    "【入力エラー】貸借対照表の資産合計と負債・純資産合計が一致していません。入力内容を確認してください。");

  if (!strengths.length) strengths.push("本シートで自動判定できる範囲では、特筆すべき評価点は検出されませんでした。");
  if (!concerns.length) concerns.push("本シートで自動検出できる範囲では、特段の懸念事項はありません。ただし業界動向・取引条件・保全状況は別途確認してください。");
  return { strengths, concerns };
}

/* -------------------------------------------------------------- 取引方針 */
export const POLICY = {
  A: "積極的に取り組んで差し支えない水準です。",
  B: "通常の与信取引に支障はない水準です。",
  C: "概ね可としつつ、定期的なモニタリングが必要です。",
  D: "慎重な対応が必要です。保全策の検討を推奨します。",
  E: "原則として新規与信は見合わせ、既存与信は回収・保全を優先してください。",
};
