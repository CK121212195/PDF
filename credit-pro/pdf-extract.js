/* ============================================================================
 * pdf-extract.js — 決算書PDFから数値を取り出す
 *
 * pdf.js のテキストレイヤーだけを使う。OCRもサーバー送信も行わない。
 * Python版（検証済み）と同じアルゴリズムを移植したもの。
 *
 * 対応している様式:
 *   ・計算書類（会社計算規則）  … 資産の部／負債の部 の左右2段組み・単年
 *   ・有価証券報告書 / 決算短信 … 前期／当期 の2年並記・単段
 *
 * 気をつけている点:
 *   ・分散配置（「現 金 及 び 預 金」）→ 空白を全除去してから照合する
 *   ・注記番号（※1,※2）→ 科目名から除去する
 *   ・2段組みには「科目が2列」と「年度が2列」の2種類がある。
 *     取り違えると前期の数字を当期として静かに取り込むため、必ず判別する
 *   ・有利子負債は「社債＋長期借入金」のように合算が要る。選択ではなく加算
 * ========================================================================== */

/* ------------------------------------------------------------------ 正規化 */
/** 分散配置・全角・注記番号を吸収して比較可能な形にする */
export function norm(s) {
  return String(s)
    .normalize("NFKC")
    .replace(/[\s\u3000]+/g, "")
    .replace(/[※＊*]\s*[\d,、，]*/g, "")
    .replace(/[（(]注\d*[）)]/g, "")
    // 小計ラベルを囲む隅付き括弧・かぎ括弧を除去する。
    // 例:「【流動資産】」「【売上高】」を辞書の「流動資産」「売上高」と一致させるため。
    // 単位表記（単位：千円）や（連結）は丸括弧なので影響しない。
    .replace(/[【】「」『』〔〕［］]/g, "");
}

const NUM_RE = /^[△▲\-−(（]?[\d,]+[)）]?$/;

export function toNum(tok) {
  const t0 = String(tok).normalize("NFKC").replace(/,/g, "");
  const neg = /^[△▲\-−(（]/.test(t0);
  const t = t0.replace(/^[△▲\-−(（]/, "").replace(/[)）]$/, "");
  if (!/^\d+$/.test(t)) return null;
  const v = parseInt(t, 10);
  return neg ? -v : v;
}

const isNum = (tok) => NUM_RE.test(String(tok).normalize("NFKC")) && toNum(tok) !== null;

/* -------------------------------------------------- pdf.js の item → 語 */
/**
 * pdf.js は文字単位でitemを返すことがある（分散配置の決算書は特に）。
 * y座標で行にまとめ、x方向に近いものを1語として連結する。
 */
function itemsToWords(items, viewportHeight) {
  const raw = [];
  for (const it of items) {
    const s = it.str;
    if (!s || !s.trim()) continue;
    const x0 = it.transform[4];
    const y = it.transform[5];
    raw.push({ text: s, x0, x1: x0 + (it.width || 0), top: viewportHeight - y });
  }
  raw.sort((a, b) => a.top - b.top || a.x0 - b.x0);

  // 行にまとめる
  const lines = [];
  for (const r of raw) {
    const ln = lines.find((l) => Math.abs(l.top - r.top) <= 3);
    if (ln) { ln.items.push(r); ln.top = (ln.top + r.top) / 2; }
    else lines.push({ top: r.top, items: [r] });
  }

  // 行内でx方向に近いものを連結して語にする
  const words = [];
  for (const ln of lines) {
    ln.items.sort((a, b) => a.x0 - b.x0);
    let cur = null;
    for (const it of ln.items) {
      const h = 10; // 想定文字高
      if (cur && it.x0 - cur.x1 <= h * 0.9 && isNum(it.text) === isNum(cur.text)) {
        cur.text += it.text; cur.x1 = it.x1;
      } else {
        if (cur) words.push(cur);
        cur = { text: it.text, x0: it.x0, x1: it.x1, top: ln.top };
      }
    }
    if (cur) words.push(cur);
  }
  return words;
}

/* ------------------------------------------------------------ レイアウト判定 */
/** 右揃えで3回以上現れるx1＝金額列。見出しの年号やページ番号は頻度で除外する */
function amountBands(words, pageWidth) {
  const nums = words.filter((w) => isNum(w.text));
  if (nums.length < 4) return [];
  // 注記番号の列（「10」「15」など）を金額列と誤認しないようにする。
  // 注記番号は桁が短く符号も付かないため、列ごとの「最大桁数」で見分ける。
  const cnt = new Map(), width = new Map();
  for (const w of nums) {
    const k = Math.round(w.x1);
    cnt.set(k, (cnt.get(k) || 0) + 1);
    const digits = w.text.replace(/[^\d]/g, "").length;
    const signed = /^[△▲\-−(（]/.test(w.text.normalize("NFKC"));
    width.set(k, Math.max(width.get(k) || 0, digits + (signed ? 10 : 0)));
  }
  let xs = [...cnt.entries()].filter(([, c]) => c >= 3).map(([x]) => x).sort((a, b) => a - b);
  if (xs.length > 1) {
    // 最も「金額らしい」列の桁数を基準に、明らかに桁の小さい列（注記番号）を落とす
    const maxW = Math.max(...xs.map((x) => width.get(x)));
    const kept = xs.filter((x) => width.get(x) >= Math.min(4, maxW) || width.get(x) >= maxW - 2);
    if (kept.length) xs = kept;
  }
  if (!xs.length) return [];
  const bands = [];
  let cur = [xs[0]];
  for (const x of xs.slice(1)) {
    if (x - cur[cur.length - 1] < pageWidth * 0.10) cur.push(x);
    else { bands.push(cur); cur = [x]; }
  }
  bands.push(cur);
  return bands.map((b) => Math.max(...b));
}

/**
 * 段組みを判定する。
 *  "single"   … 単段
 *  "accounts" … 科目が左右2列（計算書類）→ splitX で分割して読む
 *  "years"    … 年度が2列（有報・短信）→ 分割せず、行内の数値配列で持つ
 */
function classify(words, pageWidth, pageText) {
  const bands = amountBands(words, pageWidth);
  if (bands.length < 2) return { kind: "single", splitX: null };
  if (/(前連結会計年度|前事業年度|前期).{0,40}(当連結会計年度|当事業年度|当期)/.test(pageText))
    return { kind: "years", splitX: null };
  // 左段の金額と右段の金額の“あいだ”に科目名があれば「科目が2列」、なければ「年度が2列」。
  // 判定を誤ると当期と前期を取り違えるため、次の2点を必ず確認する。
  //   ・あいだの文字が、左段の金額のすぐ右から始まっているか（表の途中の注釈を拾わない）
  //   ・その文字が複数行にわたって存在するか（1行だけの脚注を科目列と誤認しない）
  const lo = bands[0], hi = bands[1];
  const mids = words.filter((w) => !isNum(w.text) && w.x0 > lo + 2 && w.x0 < hi - 20);
  const midRows = new Set(mids.map((w) => Math.round(w.top / 3)));
  const numRows = new Set(words.filter((w) => isNum(w.text) && Math.abs(w.x1 - hi) <= 4)
                               .map((w) => Math.round(w.top / 3)));
  // 右段の金額がある行のうち、あいだに科目名も存在する行の割合
  let share = 0;
  if (numRows.size) {
    let n = 0;
    for (const r of numRows) if (midRows.has(r)) n++;
    share = n / numRows.size;
  }
  if (mids.length >= 5 && midRows.size >= 3 && share >= 0.6)
    return { kind: "accounts", splitX: lo + 6 };
  return { kind: "years", splitX: null };
}

/* ---------------------------------------------------------------- 行の抽出 */
function rowsOf(words, pageWidth, pageText) {
  const { kind, splitX } = classify(words, pageWidth, pageText);
  const buckets = new Map();
  for (const w of words) {
    const col = splitX === null || w.x0 < splitX ? 0 : 1;
    const key = col + ":" + Math.round(w.top / 3);
    if (!buckets.has(key)) buckets.set(key, []);
    buckets.get(key).push(w);
  }
  // 金額列のx1位置。ここに右端が揃っている数値だけを金額として採用し、
  // 注記番号（左寄り・少桁）を金額として拾ってしまうのを防ぐ。
  const bands = amountBands(words, pageWidth);
  const onBand = (w) => bands.some((b) => Math.abs(w.x1 - b) <= 4);

  const rows = [];
  for (const ws of buckets.values()) {
    ws.sort((a, b) => a.x0 - b.x0);
    const nums = ws.filter((w) => isNum(w.text));
    const labels = ws.filter((w) => !isNum(w.text)).map((w) => w.text);
    let amountWords = bands.length ? nums.filter(onBand) : nums;
    if (!amountWords.length) amountWords = nums;        // バンドを検出できない表は従来どおり
    const amounts = amountWords.map((w) => toNum(w.text));
    if (!labels.length || !amounts.length) continue;
    rows.push({ label: norm(labels.join("")), amounts });
  }
  return { rows, kind };
}

/* ------------------------------------------------------------------ 科目辞書 */
// agg   … 集計済みの科目。1つでも見つかればそれを採用
// parts … 内訳科目。aggが無いときに全部足す
const BS_DICT = {
  cash: { agg: ["現金及び預金", "現金預金", "現金・預金", "現金及び現金同等物"], parts: ["現金", "預金"] },
  receivables: { agg: ["受取手形及び売掛金", "売上債権", "受取手形、売掛金及び契約資産", "売掛金及びその他の短期債権", "受取手形及び売掛金(純額)"], parts: ["受取手形", "売掛金", "契約資産", "電気事業未収金"] },
  inventory: { agg: ["棚卸資産"], parts: ["商品及び製品", "商品", "製品", "仕掛品", "原材料及び貯蔵品", "原材料", "貯蔵品"] },
  tangible: { agg: ["有形固定資産合計", "有形固定資産"], parts: [] },
  currentAssets: { agg: ["流動資産合計", "流動資産"], parts: [] },
  // IFRSは「非流動資産合計」。日本基準の「固定資産合計」と同じ位置づけ
  fixedAssets: { agg: ["固定資産合計", "非流動資産合計", "固定資産"], parts: [] },
  deferred: { agg: ["繰延資産合計", "繰延資産"], parts: [] },
  payables: { agg: ["支払手形及び買掛金", "仕入債務", "買掛金及びその他の短期債務"], parts: ["支払手形", "買掛金", "電気事業未払金"] },
  shortDebt: { agg: [], parts: ["短期借入金", "1年内返済予定の長期借入金", "1年以内に期限到来の固定負債",
                                "1年内償還予定の社債", "コマーシャル・ペーパー", "リース債務",
                                "その他の短期金融負債", "リース負債", "短期借入債務"] },
  currentLiab: { agg: ["流動負債合計", "流動負債"], parts: [] },
  longDebt: { agg: [], parts: ["長期借入金", "社債", "リース債務", "長期金融負債", "リース負債", "長期借入債務"] },
  fixedLiab: { agg: ["固定負債合計", "非流動負債合計", "固定負債"], parts: [] },
  equity: { agg: ["純資産合計", "純資産の部合計", "資本合計"], parts: [] },
  totalCapital: { agg: ["負債純資産合計", "負債及び純資産合計", "負債及び資本合計"], parts: [] },
};
const PL_DICT = {
  sales: { agg: ["売上高", "営業収益", "純営業収益", "売上収益", "営業収益合計"], parts: [] },
  cogs: { agg: ["売上原価"], parts: [] },
  sga: { agg: ["販売費及び一般管理費"], parts: [] },
  operating: { agg: ["営業利益", "営業損失", "営業利益又は営業損失", "営業損失(△)"], parts: [] },
  // IFRSには営業外収益/費用が無い。金融収益・その他収益・持分法投資利益を合算して同じ位置に置く
  nonOpInc: { agg: ["営業外収益", "営業外収益合計"], parts: ["金融収益", "その他収益", "持分法による投資利益"] },
  nonOpExp: { agg: ["営業外費用", "営業外費用合計"], parts: ["金融費用", "その他費用", "持分法による投資損失"] },
  ordinary: { agg: ["経常利益", "経常損失", "当期経常利益", "経常利益又は経常損失"], parts: [] },
  extraInc: { agg: ["特別利益", "特別利益合計"], parts: [] },
  extraExp: { agg: ["特別損失", "特別損失合計"], parts: [] },
  tax: { agg: ["法人税等", "法人税等合計", "法人所得税費用"], parts: [] },
  net: { agg: ["当期純利益", "当期純損失", "当期利益", "中間利益", "中間(当期)利益"], parts: [] },
};
const DEP = { agg: ["減価償却費", "減価償却費及びその他の償却費", "減価償却費及びその他の償却費", "減価償却費及び償却費", "減価償却費及びのれん償却額"], parts: [] };
const OPEX = ["営業費用", "営業費用合計"];

/** yi: 0=左（有報なら前期）／-1=右（当期） */
/**
 * yi: -1=当期（右端）／0=前期（右から2番目）
 * 有報・短信には「注記」列があり、注記番号が数値として拾われることがある。
 * 先頭から数えると注記番号を金額として取り込むため、必ず右端から数える。
 */
function pull(rows, spec, yi) {
  const val = (r) => {
    const a = r.amounts;
    if (yi === -1) return a[a.length - 1];
    return a.length >= 2 ? a[a.length - 2] : a[a.length - 1];
  };
  for (const name of spec.agg) {
    const r = rows.find((x) => x.label === name);
    if (r) return { value: val(r), source: name };
  }
  const hits = rows.filter((x) => spec.parts.includes(x.label));
  if (hits.length) {
    return { value: hits.reduce((a, r) => a + val(r), 0), source: hits.map((h) => h.label).join("＋") };
  }
  return { value: null, source: null };
}

/* ------------------------------------------------------------ 財務諸表の特定 */
/**
 * 見出しの検出。
 * pdf.js はテキストを content stream の順で返すため、見出しがページ末尾に来ることがある。
 * したがって「先頭◯文字」ではなくページ全文から探す。
 */
// 「要約中間連結損益計算書」「連結財政状態計算書」「貸借対照表」など語順・修飾語の差を吸収する
const HEAD_RE = /(?:【([^】]{0,12}?)】?|([^\n]{0,12}?))(財政状態計算書|貸借対照表|損益計算書|キャッシュ・?フロー計算書)/g;

function headingsOf(pageText) {
  const out = [];
  HEAD_RE.lastIndex = 0;
  let m;
  while ((m = HEAD_RE.exec(pageText)) !== null) {
    // 直前12文字のうち、見出しに直接つながる修飾語だけを見る。
    // 「…包括利益計算書】【要約中間連結損益計算書」のように直前に別の見出しが
    // 来ることがあるため、最後の「】」より後ろだけを修飾語として扱う。
    let mod = (m[1] || m[2] || "");
    const cut = mod.lastIndexOf("】");
    if (cut >= 0) mod = mod.slice(cut + 1);
    mod = mod.replace(/^[^ぁ-んァ-ヶ一-龥]+/, "");     // 先頭の数字・記号を落とす
    if (/包括利益$/.test(mod)) continue;               // 包括利益計算書はPLではない
    if (/注記|計上額|について|に関する|における/.test(mod)) continue;  // 本文中の言及を除く
    const kind = /財政状態|貸借/.test(m[3]) ? "BS" : /損益/.test(m[3]) ? "PL" : "CF";
    out.push({ kind, text: mod + m[3], consolidated: mod.includes("連結") });
  }
  return out;
}

/** 見出しではなく中身で財務諸表と判断できるか（見出しが別ページにある場合の保険） */
function bodyLooksLike(kind, labels) {
  const has = (...ks) => ks.some((k) => labels.has(k));
  if (kind === "BS") return has("流動資産合計", "非流動資産合計", "固定資産合計", "資産合計",
                                "流動負債合計", "非流動負債合計", "負債合計", "資本合計", "純資産合計");
  // 利益科目が無いものは損益計算書とみなさない。
  // 「売上高」だけを条件にすると、経営指標の推移表などを誤って拾う。
  if (kind === "PL") return has("営業利益", "営業損失", "営業利益又は営業損失", "経常利益", "当期経常利益",
                                "税引前利益", "税引前中間利益", "税引前当期純利益", "税金等調整前当期純利益",
                                "当期純利益", "当期純損失", "中間利益", "売上総利益");
  if (kind === "CF") return has("減価償却費", "減価償却費及びその他の償却費");
  return false;
}

/**
 * PDF全体を走査して BS/PL/CF を特定する。
 *
 * 重要な設計:
 *   ・連結を単体より優先する。混ざると貸借が合わなくなり、しかも数字は自然に見えるため気づけない
 *   ・BSは資産の部と資本の部で2ページに分かれることがある（IFRS・有報で頻出）。
 *     見出しページに続く「同じ表の続き」を結合する
 */
/* ------------------------------------------------- 会社名・金額単位の検出 */
const CO_KINDS = ["株式会社", "有限会社", "合同会社", "合資会社", "合名会社"];
// 会計用語や住所の断片が混ざっていたら、それは会社名ではなく本文の切れ端
const CO_NG = /資産|負債|利益|損失|費用|収益|合計|計算書|報告書|剰余金|余金|原価|税引|除却|配当|事項|営業外|注記|明細|科目|金額|株主資本|変動|附属|監査|決算|貸借|損益|現在|まで|から|支社|支店|本社|営業所|御中|様|[0-9]\s*[年月日]/;
// 提出会社ではないのに有価証券報告書の表紙に必ず載る名前。これを社名として拾わない。
const CO_EXCLUDE = /証券取引所|取引所|信託銀行|証券代行|監査法人|会計事務所|印刷|EDINET|縦覧/;
const CO_OK = /^[ぁ-んァ-ヶ一-龥ａ-ｚＡ-Ｚa-zA-Z0-9・ー]{2,20}$/;

/**
 * 決算書の金額単位を返す。決算書は必ず「（単位：千円）」等を明記している。
 * 本シートは百万円で計算するため、換算に使う。取り違えると規模と与信限度額が桁違いになる。
 */
export function detectUnit(pageTexts) {
  const t = pageTexts.join("");
  // ①「（単位：千円）」が最も確実。決算書・計算書類はほぼこの形
  let m = t.match(/単位[：:]?\s*([百千]?万?円)/);
  // ②有価証券報告書は「売上高(百万円)」「営業収益(百万円)」のように項目名に単位を付ける
  if (!m) m = t.match(/(?:売上高|営業収益|経常収益|純資産額|総資産額)[（(]([百千]?万?円)[）)]/);
  if (!m) return null;
  const u = m[1];
  if (u === "百万円") return { label: "百万円", toMillion: 1 };
  if (u === "千円") return { label: "千円", toMillion: 1 / 1000 };
  if (u === "万円") return { label: "万円", toMillion: 1 / 100 };
  if (u === "円") return { label: "円", toMillion: 1 / 1000000 };
  return null;
}

/**
 * 会社名を返す。誤った名前を自動入力すると手入力より危険なので、
 * 確度が低い候補は採用せず null を返す（空欄のままにする）。
 * 採用条件は「2回以上出てくる」または「1ページ目の冒頭＝表題部にある」こと。
 */
export function detectCompany(pageTexts) {
  const t = pageTexts.join("");

  // ① 有価証券報告書の表紙は「【会社名】株式会社◯◯」と明記している。
  //    ここが読めれば推測は要らないので、最優先で採用する。
  // 本文の正規化で【】は既に外れているため、括弧なしの形で照合する。
  // 例：「…事業年度第13期(…)会社名株式会社マネーフォワード英訳名MoneyForward…」
  const KIND_RE = "株式会社|有限会社|合同会社|合資会社|合名会社";
  const labeled = (pageTexts[0] || "").match(new RegExp(
    "(?:会社名|商号|名称)\\s*((?:" + KIND_RE + ")[ぁ-んァ-ヶ一-龥ａ-ｚＡ-Ｚa-zA-Z0-9・ー]{1,20}" +
    "|[ぁ-んァ-ヶ一-龥ａ-ｚＡ-Ｚa-zA-Z0-9・ー]{1,20}(?:" + KIND_RE + "))"));
  if (labeled) {
    // 直後の見出し語（英訳名 など）を巻き込んでいたら落とす
    const n = labeled[1].replace(/(英訳名|代表者|本店|電話番号|事務連絡者|最寄り|縦覧).*$/, "");
    if (n.length >= 4) return n;
  }

  const cands = new Map();
  const add = (n) => { if (n) cands.set(n, (cands.get(n) || 0) + 1); };
  for (const kind of CO_KINDS) {
    let i = -1;
    while ((i = t.indexOf(kind, i + 1)) !== -1) {
      // 後置形（株式会社○○）：直後を名前として読む。次の会社種別・記号・数字で打ち切る
      let after = t.slice(i + kind.length, i + kind.length + 14)
                   .split(/[（(【］\]、。：:0-9０-９]/)[0];
      for (const k2 of CO_KINDS) { const p2 = after.indexOf(k2); if (p2 >= 0) after = after.slice(0, p2); }
      if (CO_OK.test(after) && !CO_NG.test(after) && !CO_EXCLUDE.test(kind + after)) add(kind + after);
      // 前置形（○○株式会社）：直前を名前として読む。長い候補から順に見る
      const before = t.slice(Math.max(0, i - 20), i);
      for (let s = 0; s < before.length; s++) {
        let cand = before.slice(s);
        // 「…でOTNet株式会社」のように直前の助詞を巻き込むのを防ぐ
        cand = cand.replace(/^[ぁ-ん]{1,2}(?=[ァ-ヶ一-龥A-Za-zＡ-Ｚ0-9])/, "");
        // 「…3月31日沖縄電力株式会社」のように日付の末尾文字を巻き込むのを防ぐ
        cand = cand.replace(/^[年月日期至自現在]+/, "");
        // 「有価証券報告書GMO…」のように直前の見出し語を巻き込むのを防ぐ
        cand = cand.replace(/^(?:有価証券報告書|報告書|書類|表紙|提出会社|会社名|商号|名称|当社|同社)+/, "");
        if (CO_OK.test(cand) && !CO_NG.test(cand) && !CO_EXCLUDE.test(cand + kind)) { add(cand + kind); break; }
      }
      i += kind.length;
    }
  }
  // 1ページ目の表題部にある社名が最も信頼できる。有価証券報告書は必ずここに提出会社名を書く。
  // 本文には「株式会社を設立」のような文章の断片が何度も出るため、頻度だけで選ぶと負ける。
  const head = (pageTexts[0] || "").slice(0, 400);
  const inHead = [...cands.keys()].filter((k) => head.includes(k));
  if (inHead.length) {
    // 表題部に複数あるときは、
    //  ① 表題部での出現位置が最も後ろ（提出会社名は見出しの最後に書かれる）
    //  ② 同じ位置なら短いほう（余計な語を巻き込んでいない）
    // の順で選ぶ。長いほうを選ぶと「役員…縦覧に供する場所株式会社」のような塊を掴む。
    return inHead.sort((a, b) => {
      const d = head.indexOf(a) - head.indexOf(b);
      return d !== 0 ? d : a.length - b.length;
    })[0];
  }
  // 表題部で見つからないときだけ、2回以上出てくるものを採る
  let best = 0, company = null;
  for (const [k, v] of cands) if (v >= 2 && v > best) { best = v; company = k; }
  return company;
}

/**
 * 本ツールの判定モデルに載らない業態かどうかを返す。
 * 銀行・保険の決算書は、貸借対照表に流動／固定の区分が無く、
 * 損益計算書も売上高ではなく経常収益で構成されるため、そもそも様式が違う。
 * また業界基準に用いる統計も金融業・保険業を対象外としている。
 * 「読み取れませんでした」ではなく、対象外であることを伝えるために使う。
 */
export function detectOutOfScope(pageTexts) {
  const t = pageTexts.join("");
  const bank = /経常収益/.test(t) && /(?:預金|貸出金|コールローン|資金運用収益)/.test(t);
  if (bank) return "bank";
  const ins = /(?:保険料等収入|責任準備金|支払備金)/.test(t);
  if (ins) return "insurance";
  return null;
}

/* ------------------------------------------------------------------ 決算期 */
/**
 * 決算期を特定する。決算書の様式によって書き方が違うため、次の順で探す。
 *   ① 会計期間「自2025年4月1日 至2026年3月31日」… 決算日が確実に分かる
 *   ② 貸借対照表日「2026年3月31日現在」
 *   ③ 期数「第47期」… 決算日は分からないが、期の前後関係は分かる
 * 戻り値の end は並べ替えに使う（新しいものが後ろ）。label は画面表示用。
 */
export function detectPeriod(pageTexts) {
  const t = pageTexts.join("");
  const out = { end: null, no: null, label: "" };

  let m = t.match(/自(\d{4})年(\d{1,2})月(\d{1,2})日至(\d{4})年(\d{1,2})月(\d{1,2})日/);
  if (m) out.end = `${m[4]}-${String(m[5]).padStart(2, "0")}-${String(m[6]).padStart(2, "0")}`;
  if (!out.end) {
    m = t.match(/(\d{4})年(\d{1,2})月(\d{1,2})日現在/);
    if (m) out.end = `${m[1]}-${String(m[2]).padStart(2, "0")}-${String(m[3]).padStart(2, "0")}`;
  }
  m = t.match(/第(\d{1,3})期/);
  if (m) out.no = parseInt(m[1], 10);

  if (out.end) {
    const [y, mo] = out.end.split("-");
    out.label = `${y}年${parseInt(mo, 10)}月期`;
    if (out.no) out.label = `第${out.no}期（${out.label}）`;
  } else if (out.no) {
    out.label = `第${out.no}期`;
  }
  return out;
}

export async function scanPdf(pdfjs, buf, onProgress) {
  // cMap（文字コード変換表）を必ず渡す。
  // 有価証券報告書は MS-Gothic 等のCIDフォントを埋め込みつつ ToUnicode を持たないものが多く、
  // これが無いと文字が1文字も取り出せず「画像PDF」と誤判定してしまう。
  const doc = await pdfjs.getDocument({
    data: new Uint8Array(buf),
    useWorkerFetch: false,
    isEvalSupported: false,
    cMapUrl: new URL("../vendor/cmaps/", import.meta.url).href,
    cMapPacked: true,
  }).promise;
  const pages = [];
  const pageTexts = [];
  let totalChars = 0;

  for (let i = 1; i <= doc.numPages; i++) {
    if (onProgress) onProgress(i, doc.numPages);
    const page = await doc.getPage(i);
    const vp = page.getViewport({ scale: 1 });
    const tc = await page.getTextContent();
    const words = itemsToWords(tc.items, vp.height);
    const pageText = norm(words.map((w) => w.text).join(""));
    totalChars += pageText.length;
    if (pageTexts.length < 12) pageTexts.push(pageText);
    const { rows, kind } = words.length ? rowsOf(words, vp.width, pageText) : { rows: [], kind: "single" };
    pages.push({ no: i, rows, kind, labels: new Set(rows.map((r) => r.label)),
                 headings: headingsOf(pageText) });
  }
  await doc.destroy();

  // テキストレイヤーが無いPDF（スキャン画像）は、その旨を明示して返す
  if (totalChars < 200) return { _noText: true };

  // 決算期を拾う。複数のPDFを投入されたとき、新しい順に並べ替えるために使う。
  // 読めなくても処理は続ける（並べ替えを諦めて投入順にするだけ）。
  const period = detectPeriod(pageTexts);
  const outOfScope = detectOutOfScope(pageTexts);
  const unit = detectUnit(pageTexts);
  const company = detectCompany(pageTexts);

  const found = {};
  // 減価償却費を全ページから探すために保持する。列挙対象に混ざらないよう非列挙にする。
  Object.defineProperty(found, "_pages", { value: pages, enumerable: false });
  Object.defineProperty(found, "_period", { value: period, enumerable: false });
  Object.defineProperty(found, "_unit", { value: unit, enumerable: false });
  Object.defineProperty(found, "_company", { value: company, enumerable: false });
  Object.defineProperty(found, "_outOfScope", { value: outOfScope, enumerable: false });
  for (const wantConsolidated of [true, false]) {          // 連結を先に探す
    for (const kind of ["BS", "PL", "CF"]) {
      if (found[kind]) continue;
      for (const pg of pages) {
        const h = pg.headings.find((x) => x.kind === kind && x.consolidated === wantConsolidated);
        if (!h || !bodyLooksLike(kind, pg.labels)) continue;
        found[kind] = { page: pg.no, rows: pg.rows.slice(), kind: pg.kind,
                        consolidated: h.consolidated, heading: h.text };
        // 続きのページを結合する。
        // 貸借対照表だけが「資産の部」と「資本の部」で2ページに割れるため、BSに限定する。
        // 損益計算書に適用すると、次ページの包括利益計算書まで飲み込んでしまう。
        if (kind === "BS") {
          const need = ["資本合計", "純資産合計", "負債及び資本合計", "負債純資産合計", "負債及び純資産合計"];
          const hasEquity = (labels) => need.some((k) => labels.has(k));
          for (let n = pg.no; n < pages.length && !hasEquity(new Set(found.BS.rows.map((r) => r.label))); n++) {
            const nx = pages[n];
            if (!nx || nx.headings.length || !nx.rows.length) break;
            if (!bodyLooksLike("BS", nx.labels)) break;
            found.BS.rows.push(...nx.rows);
            found.BS.continued = (found.BS.continued || []).concat(nx.no);
          }
        }
        break;
      }
    }
    if (found.BS && found.PL) break;   // 連結で揃ったら単体は見ない
  }
  return found;
}

/** 負債の部を流動セクション／固定（非流動）セクションに切り分ける */
function sliceSection(rows, which) {
  const endCur = rows.findIndex((r) => /^(流動負債合計)$/.test(r.label));
  const endFix = rows.findIndex((r) => /^(固定負債合計|非流動負債合計)$/.test(r.label));
  if (endCur < 0 || endFix < 0 || endFix <= endCur) return rows;   // 判別できなければ全体を対象
  // 流動負債の節は「流動負債」の見出し行から合計行まで。見出しが無ければ合計行の手前20行を見る
  const startCur = Math.max(0, (() => {
    const i = rows.findIndex((r) => /^(流動負債)$/.test(r.label));
    return i >= 0 && i < endCur ? i : endCur - 20;
  })());
  return which === "current"
    ? rows.slice(startCur, endCur + 1)
    : rows.slice(endCur + 1, endFix + 1);
}

/* -------------------------------------------------------- 1期分の組み立て */
export function buildPeriod(found, yi) {
  const out = {}, source = {}, warnings = [];
  const bs = found.BS?.rows || [], pl = found.PL?.rows || [], cf = found.CF?.rows || [];

  for (const [k, spec] of Object.entries(BS_DICT)) {
    // 有利子負債は「リース負債」のように流動・非流動で同名の科目があるため、
    // 区間を限定して拾わないと二重計上になる
    const scope = k === "shortDebt" ? sliceSection(bs, "current")
                : k === "longDebt" ? sliceSection(bs, "fixed") : bs;
    const { value, source: s } = pull(scope, spec, yi);
    out[k] = value; if (s) source[k] = s;
  }
  for (const [k, spec] of Object.entries(PL_DICT)) {
    const { value, source: s } = pull(pl, spec, yi);
    out[k] = value; if (s) source[k] = s;
  }

  // 減価償却費：CF計算書 →（無ければ）損益計算書 →（無ければ）全ページ の順に探す。
  // 中小企業の計算書類にはCF計算書が無く、販売費及び一般管理費明細書や
  // 製造原価報告書に載っていることが多いため、最後は全ページを走査する。
  let dep = pull(cf, DEP, yi);
  if (dep.value === null) {
    dep = pull(pl, DEP, yi);
    if (dep.value !== null) dep.source += "（損益計算書より）";
  }
  if (dep.value === null && Array.isArray(found._pages)) {
    // 販管費明細と製造原価報告書の両方にある場合は合算する（両方が費用計上分）
    const hits = [];
    for (const pg of found._pages) {
      if (found.CF && pg.no === found.CF.page) continue;
      if (found.PL && pg.no === found.PL.page) continue;
      const r = pg.rows.find((x) => DEP.agg.includes(x.label));
      if (r) {
        const a = r.amounts;
        const v = yi === -1 ? a[a.length - 1] : (a.length >= 2 ? a[a.length - 2] : a[a.length - 1]);
        if (typeof v === "number") hits.push({ page: pg.no, value: v });
      }
    }
    if (hits.length) {
      dep = { value: hits.reduce((a, h) => a + h.value, 0),
              source: hits.map((h) => `${h.page}ページ`).join("＋") + "（明細より）" };
      if (hits.length > 1) {
        warnings.push("減価償却費を複数ページから合算しました（" +
          hits.map((h) => `${h.page}ページ ${h.value.toLocaleString()}`).join(" ＋ ") +
          "）。二重計上でないかご確認ください。");
      }
    }
  }
  out.depreciation = dep.value;
  if (dep.source) source.depreciation = dep.source;
  if (dep.value === null)
    warnings.push("減価償却費が見つかりませんでした。償還余力の算定に必要です。製造原価明細・販管費明細からご入力ください。");

  // 売上原価／販管費の区分が無い様式（電気事業・通信事業など）
  if (out.cogs === null && out.sga === null) {
    for (const name of OPEX) {
      const r = pl.find((x) => x.label === name);
      if (r) {
        out.sga = yi === -1 ? r.amounts[r.amounts.length - 1] : r.amounts[yi];
        out.cogs = 0;
        source.sga = name + "（売上原価との区分なし）";
        warnings.push("この様式には売上原価と販管費の区分がありません。営業費用の全額を販管費として扱っています。売上総利益率は実態を表しません。");
        break;
      }
    }
  }

  // IFRSの損益計算書は費用を「△835,371」のようにマイナス表記する。
  // 本シートは費用を正の数で扱うため、符号を揃える。
  for (const k of ["cogs", "sga", "nonOpExp", "extraExp", "tax"]) {
    if (typeof out[k] === "number" && out[k] < 0) out[k] = -out[k];
  }
  // 営業外収益がマイナスで返るケース（その他費用を含めて相殺された場合）も正に寄せる
  if (typeof out.nonOpInc === "number" && out.nonOpInc < 0) {
    out.nonOpExp = (out.nonOpExp || 0) + Math.abs(out.nonOpInc);
    out.nonOpInc = 0;
  }

  // 「その他」は小計からの差額で埋める。こうすると小計が必ず一致する
  const d = (total, ...items) =>
    out[total] === null || out[total] === undefined
      ? null
      : out[total] - items.reduce((a, k) => a + (out[k] || 0), 0);
  out.otherCurrentAssets = d("currentAssets", "cash", "receivables", "inventory");
  out.otherFixedAssets = d("fixedAssets", "tangible");
  out.otherCurrentLiab = d("currentLiab", "payables", "shortDebt");
  out.otherFixedLiab = d("fixedLiab", "longDebt");

  // 決算書自体の端数処理で、資産側と負債側が1〜数単位ずれることがある。
  // 重要性のない差は「その他流動資産」で吸収して検算を0にする。
  // 吸収しないと、正しく読み取れているのに赤い警告が出て利用者を不安にさせる。
  {
    const a = (out.currentAssets || 0) + (out.fixedAssets || 0) + (out.deferred || 0);
    const l = (out.currentLiab || 0) + (out.fixedLiab || 0) + (out.equity || 0);
    const gap = a - l;
    const tol = Math.max(2, Math.abs(a) * 0.0001);   // 総資産の0.01%まで
    if (gap !== 0 && Math.abs(gap) <= tol && out.otherCurrentAssets !== null) {
      out.otherCurrentAssets -= gap;
      out.currentAssets -= gap;
      source.roundingAdjust = `端数調整 ${gap > 0 ? "−" : "＋"}${Math.abs(gap)}`;
    }
  }

  if (out.tangible === null && out.fixedAssets !== null)
    warnings.push("有形固定資産の内訳が特定できませんでした。固定資産合計は正しいですが、有形と無形の内訳はご確認ください。");
  for (const k of ["shortDebt", "longDebt"]) {
    if (out[k] === null) {
      out[k] = 0;
      warnings.push("有利子負債に該当する科目が見つかりませんでした。無借金でない場合は必ずご入力ください。");
      break;
    }
  }
  return { values: out, source, warnings };
}

/** 独立に抽出した資産側小計と負債側小計を突き合わせる */
export function validatePeriod(v) {
  const a = (v.currentAssets || 0) + (v.fixedAssets || 0) + (v.deferred || 0);
  const l = (v.currentLiab || 0) + (v.fixedLiab || 0) + (v.equity || 0);
  const msgs = [];
  if (Math.abs(a - l) > Math.max(2, Math.abs(a) * 0.0001))
    msgs.push(`貸借が一致しません（資産計 ${a.toLocaleString()} と 負債純資産計 ${l.toLocaleString()} の差 ${(a - l).toLocaleString()}）`);
  if ([v.ordinary, v.operating, v.nonOpInc, v.nonOpExp].every((x) => x !== null && x !== undefined)) {
    const calc = v.operating + v.nonOpInc - v.nonOpExp;
    if (Math.abs(calc - v.ordinary) > 2)
      msgs.push(`経常利益が積み上がりません（計算 ${calc.toLocaleString()} と 記載 ${v.ordinary.toLocaleString()}）`);
  }
  return { diff: a - l, messages: msgs };
}

/** エンジン入力のキー名に合わせて1期分を返す */
export function toEngineFields(v) {
  return {
    sales: v.sales, cogs: v.cogs, sga: v.sga,
    nonOpInc: v.nonOpInc, nonOpExp: v.nonOpExp,
    extraInc: v.extraInc, extraExp: v.extraExp, tax: v.tax,
    depreciation: v.depreciation,
    cash: v.cash, receivables: v.receivables, inventory: v.inventory,
    otherCurrentAssets: v.otherCurrentAssets, tangible: v.tangible,
    otherFixedAssets: v.otherFixedAssets, deferred: v.deferred,
    payables: v.payables, shortDebt: v.shortDebt,
    otherCurrentLiab: v.otherCurrentLiab, longDebt: v.longDebt,
    otherFixedLiab: v.otherFixedLiab, equity: v.equity,
  };
}
