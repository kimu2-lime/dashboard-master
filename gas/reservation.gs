// ============================================================
//  予約CSV分析（HPB/サロンボード予約エクスポート）
//  ・SS_RESERVE … 予約CSVを貼り付けたスプシ（月ごとタブ ＋「店名エイリアス」タブ）
//  ・店名エイリアス … A列=お店名(掲載名)  B列=略称  C列=メモ
//  会計CSV(SS_SALON)を「正の数字」とし、予約CSVは予約固有情報専用。
//  突合キー: お客様名(フリガナ) ＋ 来店日≈会計日 ＋ 店舗(略称)
//
//  ▼まずはこの Phase1a:「読込＋エイリアス変換が正しく効くか」をログ検証する。
//    testReservationParse() を実行 → 実行ログで件数・未マップ店名を確認。
// ============================================================

// ── 予約CSVの列位置（0始まり・エクスポート列順は固定） ──
const RESV_COL = {
  STORE:        0,   // お店名（掲載名）
  STATUS:       1,   // ステータス
  NOMINATE:     5,   // 指名予約有無
  STAFF:        4,   // スタッフ名
  ROUTE:        15,  // 予約経路
  VISIT_DATE:   7,   // 来店日(YYYYMMDD)
  HOPE1_DATE:   11,  // 第一来店希望日（≠来店日でリスケ判定）
  RSV_MENU:     18,  // 予約時メニュー
  RSV_COUPON:   19,  // 予約時HotPepperBeautyクーポン
  KANA:         24,  // フリガナ
  GENDER:       27,  // 性別
  NAME_KANA:    28,  // 氏名(カナ)
  NAME_KANJI:   29,  // 氏名(漢字)
  CUST_NO:      31,  // お客様番号（予約CSVではほぼ空）
  RSV_AMT:      32,  // 予約時合計金額
  VISIT_AMT:    38,  // 来店処理金額
  FIRST_VISIT:  47,  // このサロンに行くのは初めてですか？
  ACT_CATEGORY: 49,  // 会計時メニューカテゴリ
  ACT_MENU:     50,  // 会計時メニュー（契約/消化/利用 ラベルを含む）
  ACT_AMT:      56   // 会計時合計金額
};

// ステータス分類
const RESV_STATUS_VISITED  = ['会計済み', '済み'];           // 来店完了
const RESV_STATUS_CXL_CUST = ['お客様キャンセル'];           // 客都合
const RESV_STATUS_CXL_SALON= ['サロンキャンセル'];           // サロン都合
const RESV_STATUS_CXL_NOSHOW=['無断キャンセル'];             // 無断
const ALIAS_SHEET_NAME = '店名エイリアス';
const ALIAS_EXCLUDE    = '（対象外）';                       // 集計から除外する略称

// ── SS_RESERVE を開く（URL丸ごと貼られていてもIDを抽出する） ──
function getReserveSS_() {
  const props = PropertiesService.getScriptProperties();
  let raw = props.getProperty('SS_RESERVE');
  if (!raw) throw new Error('スクリプトプロパティ SS_RESERVE が未設定です');
  raw = raw.toString().trim();
  // URLが貼られていたら /d/●●●/ の●●●を取り出す
  const m = raw.match(/\/d\/([a-zA-Z0-9_-]+)/);
  const id = m ? m[1] : raw;
  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error(
      'SS_RESERVE のスプシを開けません（ID="' + id + '"）。\n' +
      '原因の可能性: ①IDが間違い ②このGASを動かしているGoogleアカウントに共有されていない。\n' +
      'スプシをGASと同じアカウントで開けるか / 共有設定を確認してください。\n元エラー: ' + e.message
    );
  }
}

// ── 店名エイリアス表を読む → { お店名: 略称 }（対象外は除外して別管理） ──
function loadReserveAlias_(ss) {
  const sheet = ss.getSheetByName(ALIAS_SHEET_NAME);
  if (!sheet) throw new Error('「' + ALIAS_SHEET_NAME + '」タブが見つかりません');
  const rows = sheet.getDataRange().getValues();
  const map = {};       // お店名 → 略称
  const excluded = {};  // お店名 → true（集計対象外）
  for (let i = 1; i < rows.length; i++) {            // 1行目はヘッダー
    const name  = (rows[i][0] || '').toString().trim();
    const short = (rows[i][1] || '').toString().trim();
    if (!name) continue;
    if (short === ALIAS_EXCLUDE || short === '') { excluded[name] = true; continue; }
    map[name] = short;
  }
  return { map: map, excluded: excluded };
}

// ── 月タブの名前一覧 ──
//   「店名エイリアス」「設定」を除き、A1に「お店名」を持つ＝予約データタブを月タブとみなす。
//   タブ名の形式（2026-06 / 2026年6月 / 202606 / シート1 等）は問わない。
function listReserveMonthSheets_(ss) {
  const skip = [ALIAS_SHEET_NAME, '設定'];
  return ss.getSheets()
    .filter(function(s){
      if (skip.indexOf(s.getName()) >= 0) return false;
      if (s.getLastRow() < 2 || s.getLastColumn() < 1) return false;
      const a1 = (s.getRange(1, 1).getValue() || '').toString().trim();
      return a1.indexOf('お店名') >= 0;   // ヘッダー先頭=お店名 なら予約データ
    })
    .map(function(s){ return s.getName(); })
    .sort();
}

// ── 1つの月タブを読み、行オブジェクトの配列を返す ──
//   alias で略称へ変換し、対象外・未マップは meta に記録（捨てない＝検証用）
function readReservationRows_(ss, monthTab, alias) {
  const sheet = ss.getSheetByName(monthTab);
  if (!sheet) throw new Error('タブが見つかりません: ' + monthTab);
  const values = sheet.getDataRange().getValues();
  const rows = [];
  const unmapped = {};   // 未マップのお店名 → 件数
  let excludedCount = 0;
  const g = function(row, idx) { return idx < row.length ? (row[idx] == null ? '' : row[idx].toString().trim()) : ''; };

  for (let i = 1; i < values.length; i++) {           // 1行目ヘッダー
    const row = values[i];
    const storeName = g(row, RESV_COL.STORE);
    if (!storeName) continue;
    if (alias.excluded[storeName]) { excludedCount++; continue; }
    const short = alias.map[storeName];
    if (!short) { unmapped[storeName] = (unmapped[storeName] || 0) + 1; continue; }

    rows.push({
      short:       short,
      storeName:   storeName,
      status:      g(row, RESV_COL.STATUS),
      nominate:    g(row, RESV_COL.NOMINATE),
      staff:       g(row, RESV_COL.STAFF),
      route:       g(row, RESV_COL.ROUTE),
      visitDate:   g(row, RESV_COL.VISIT_DATE).replace(/[-\/]/g, ''),
      hope1Date:   g(row, RESV_COL.HOPE1_DATE).replace(/[-\/]/g, ''),
      rsvMenu:     g(row, RESV_COL.RSV_MENU),
      rsvCoupon:   g(row, RESV_COL.RSV_COUPON),
      gender:      g(row, RESV_COL.GENDER),
      kana:        g(row, RESV_COL.KANA) || g(row, RESV_COL.NAME_KANA),
      firstVisit:  g(row, RESV_COL.FIRST_VISIT),
      actCategory: g(row, RESV_COL.ACT_CATEGORY),
      actMenu:     g(row, RESV_COL.ACT_MENU),
      rsvAmt:      resvNum_(g(row, RESV_COL.RSV_AMT)),
      visitAmt:    resvNum_(g(row, RESV_COL.VISIT_AMT)),
      actAmt:      resvNum_(g(row, RESV_COL.ACT_AMT))
    });
  }
  return { rows: rows, unmapped: unmapped, excludedCount: excludedCount, totalRows: values.length - 1 };
}

function resvNum_(s) {
  if (!s) return 0;
  const n = parseFloat(s.toString().replace(/[^\d.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

// 来店完了か
function resvIsVisited_(status) { return RESV_STATUS_VISITED.indexOf(status) >= 0; }
// 新規契約か（会計時メニューに「契約」を含む）※会計CSV突合前の予約ベース簡易判定
function resvIsContract_(actMenu) { return /契約/.test(actMenu || ''); }

// ============================================================
//  ★検証用：実行してログを確認する（Phase1a）
//  期待：・total と mapped がほぼ一致（未マップが0に近い）
//        ・ステータス内訳・来店完了の契約率がそれっぽい数字
// ============================================================
function testReservationParse() {
  const ss = getReserveSS_();
  const alias = loadReserveAlias_(ss);
  const months = listReserveMonthSheets_(ss);
  Logger.log('SS_RESERVE 月タブ: ' + JSON.stringify(months));
  Logger.log('エイリアス登録数: ' + Object.keys(alias.map).length + ' / 対象外: ' + Object.keys(alias.excluded).length);
  if (!months.length) { Logger.log('⚠ A1セルが「お店名」のデータタブが見つかりません（予約データが貼られているか確認）'); return; }

  const target = months[months.length - 1];           // 最新月で検証
  const res = readReservationRows_(ss, target, alias);
  Logger.log('── 検証対象タブ: ' + target + ' ──');
  Logger.log('総行数: ' + res.totalRows + ' / マップ成功: ' + res.rows.length +
             ' / 対象外: ' + res.excludedCount + ' / 未マップ: ' + Object.keys(res.unmapped).length + '店名');
  if (Object.keys(res.unmapped).length) {
    Logger.log('▼未マップ店名（エイリアス表に追加が必要）:');
    Object.keys(res.unmapped).sort(function(a,b){return res.unmapped[b]-res.unmapped[a];})
      .slice(0, 30).forEach(function(n){ Logger.log('   ' + res.unmapped[n] + '件  ' + n); });
  }

  // ステータス内訳
  const st = {};
  res.rows.forEach(function(r){ st[r.status] = (st[r.status]||0) + 1; });
  Logger.log('▼ステータス内訳: ' + JSON.stringify(st));

  // 来店完了の予約ベース契約率（会計CSV突合は次フェーズ）
  const visited = res.rows.filter(function(r){ return resvIsVisited_(r.status); });
  const contract = visited.filter(function(r){ return resvIsContract_(r.actMenu); });
  Logger.log('▼来店完了: ' + visited.length + ' / うち契約: ' + contract.length +
             ' (' + (visited.length ? Math.round(contract.length/visited.length*1000)/10 : 0) + '%)');

  // 略称ごとの件数 上位10（マッピングが店舗単位で効いているか確認）
  const byStore = {};
  res.rows.forEach(function(r){ byStore[r.short] = (byStore[r.short]||0) + 1; });
  const top = Object.keys(byStore).sort(function(a,b){return byStore[b]-byStore[a];}).slice(0, 10);
  Logger.log('▼略称別 件数 上位10: ' + top.map(function(s){ return s + ':' + byStore[s]; }).join(' / '));
}

// ============================================================
//  Phase1b: 会計CSV(SS_SALON)突合
//  ・会計データは YYYYMM_実績 シートに会計ID単位で入っている
//  ・突合キー: フリガナ(正規化) ＋ 会計日(YYYYMMDD) ＋ 店舗(略称)
//  ・会計CSVを「正」とし、契約/消化/利用・金額・新規再来を取得
// ============================================================

// フリガナ正規化（空白除去＋ひらがな→カタカナ）
function normKana_(s) {
  if (!s) return '';
  return s.toString()
    .replace(/[\s　]/g, '')
    .replace(/[ぁ-ん]/g, function(c){ return String.fromCharCode(c.charCodeAt(0) + 0x60); });
}

// 日付を YYYYMMDD 文字列に
function toYmd_(v) {
  if (v == null || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, 'JST', 'yyyyMMdd');
  }
  const s = v.toString().trim().replace(/[-\/]/g, '');
  return /^\d{8}/.test(s) ? s.slice(0, 8) : '';
}

// 会計CSVを読み、突合インデックスを作る
//   key = フリガナ正規化 | 会計日YYYYMMDD | 略称
//   value = { hasContract, hasShoka, hasRiyou, amount, shinki, staff, accId }
function readAccountingIndex_(alias, yyyymm) {
  const ss = getSS('SS_SALON');
  const sheetName = yyyymm + '_実績';
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log('⚠ 会計シートが見つかりません: ' + sheetName + ' （SS_SALONのタブ名を確認）');
    return { index: {}, stats: { found: false, rows: 0, accIds: 0, unmatchedStores: {} } };
  }
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0].map(function(h){ return h.toString().trim(); });
  const col = {};
  headers.forEach(function(h, i){ col[h] = i; });
  // メニュー列はヘッダーに空白ゆらぎがあるため部分一致で特定
  let menuIdx = -1;
  headers.forEach(function(h, i){ if (h.indexOf('メニュー') >= 0 && h.indexOf('店販') >= 0) menuIdx = i; });
  const ci = function(name){ return col[name] !== undefined ? col[name] : -1; };
  const iStore = ci('お店名'), iDate = ci('会計日'), iKubun = ci('会計区分'),
        iCat = ci('カテゴリ'), iAmt = ci('金額'), iStaff = ci('スタッフ'),
        iFuri = ci('お客様名（フリガナ）'), iShinki = ci('新規再来'), iAccId = ci('会計ID');

  const gs = function(row, idx){ return idx >= 0 && idx < row.length && row[idx] != null ? row[idx].toString().trim() : ''; };
  const amt = function(row){ const v = iAmt>=0 ? row[iAmt] : 0; if (typeof v === 'number') return v; return parseFloat((v||'').toString().replace(/[^\d.\-]/g,'')) || 0; };

  // 会計IDごとにまとめる（会計区分=会計 のみ。取り消しは除外）
  const byAcc = {};
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (gs(row, iKubun) !== '会計') continue;
    const accId = gs(row, iAccId);
    if (!accId) continue;
    if (!byAcc[accId]) byAcc[accId] = [];
    byAcc[accId].push(row);
  }

  const index = {};
  let accIdCount = 0;
  const unmatchedStores = {};
  Object.keys(byAcc).forEach(function(accId){
    const lines = byAcc[accId];
    const first = lines[0];
    const storeName = gs(first, iStore);
    const short = alias.map[storeName];
    if (!short) {
      if (!alias.excluded[storeName]) unmatchedStores[storeName] = (unmatchedStores[storeName]||0) + 1;
      return;
    }
    const ymd = toYmd_(iDate>=0 ? first[iDate] : '');
    const furi = normKana_(gs(first, iFuri));
    if (!furi || !ymd) return;
    let hasContract = false, hasShoka = false, hasRiyou = false, total = 0, shinki = '', staff = '';
    lines.forEach(function(ln){
      const menu = menuIdx>=0 ? gs(ln, menuIdx) : '';
      const cat  = gs(ln, iCat);
      if (/契約/.test(menu) && /フェイシャル|ボディ|その他/.test(cat)) hasContract = true;
      if (/消化/.test(menu)) hasShoka = true;
      if (/利用|サブスク/.test(menu)) hasRiyou = true;
      total += amt(ln);
      const sk = gs(ln, iShinki); if (sk === '新規' || sk === '再来') shinki = sk;
      const st = gs(ln, iStaff); if (st && !staff) staff = st;
    });
    accIdCount++;
    const key = furi + '|' + ymd + '|' + short;
    // 同キー複数会計IDは契約優先でマージ
    const prev = index[key];
    if (!prev) {
      index[key] = { hasContract: hasContract, hasShoka: hasShoka, hasRiyou: hasRiyou, amount: total, shinki: shinki, staff: staff, accId: accId };
    } else {
      prev.hasContract = prev.hasContract || hasContract;
      prev.hasShoka    = prev.hasShoka || hasShoka;
      prev.hasRiyou    = prev.hasRiyou || hasRiyou;
      prev.amount     += total;
    }
  });
  return { index: index, stats: { found: true, rows: rows.length - 1, accIds: accIdCount, unmatchedStores: unmatchedStores } };
}

// ★検証用：突合の歩留まりと「予約CSV契約判定 vs 会計CSV契約判定」の一致を確認
function testAccountingJoin() {
  const ss = getReserveSS_();
  const alias = loadReserveAlias_(ss);
  const months = listReserveMonthSheets_(ss);
  if (!months.length) { Logger.log('⚠ 予約データタブがありません'); return; }
  const res = readReservationRows_(ss, months[months.length - 1], alias);
  const visited = res.rows.filter(function(r){ return resvIsVisited_(r.status); });

  // 来店日の最頻 YYYYMM を会計月とする
  const ymCount = {};
  visited.forEach(function(r){ const ym = (r.visitDate||'').slice(0,6); if (ym) ymCount[ym] = (ymCount[ym]||0)+1; });
  const yyyymm = Object.keys(ymCount).sort(function(a,b){return ymCount[b]-ymCount[a];})[0] || '';
  Logger.log('会計対象月: ' + yyyymm + '_実績 を読みます');

  const acc = readAccountingIndex_(alias, yyyymm);
  if (!acc.stats.found) return;
  Logger.log('会計CSV: ' + acc.stats.rows + '行 / 会計ID(会計のみ): ' + acc.stats.accIds + ' / 突合キー数: ' + Object.keys(acc.index).length);
  if (Object.keys(acc.stats.unmatchedStores).length) {
    Logger.log('▼会計CSVで略称マップできなかった店名（エイリアス要追加）:');
    Object.keys(acc.stats.unmatchedStores).slice(0, 15).forEach(function(n){ Logger.log('   ' + acc.stats.unmatchedStores[n] + '件  ' + n); });
  }

  // 突合
  let matched = 0, both = 0, resvOnly = 0, accOnly = 0, agree = 0;
  visited.forEach(function(r){
    const key = normKana_(r.kana) + '|' + r.visitDate + '|' + r.short;
    const a = acc.index[key];
    if (!a) return;
    matched++;
    const resvC = resvIsContract_(r.actMenu);   // 予約CSVベース
    const accC  = a.hasContract;                 // 会計CSVベース（正）
    if (resvC && accC) both++;
    else if (resvC && !accC) resvOnly++;
    else if (!resvC && accC) accOnly++;
    if (resvC === accC) agree++;
  });
  const rate = visited.length ? Math.round(matched/visited.length*1000)/10 : 0;
  Logger.log('▼突合結果: 来店完了' + visited.length + ' 件中 ' + matched + ' 件が会計CSVと一致 (' + rate + '%)');
  Logger.log('▼契約判定の一致(突合できた' + matched + '件): 予約=会計 ' + agree + '件 (' +
             (matched?Math.round(agree/matched*1000)/10:0) + '%)');
  Logger.log('   両方契約:' + both + ' / 予約だけ契約:' + resvOnly + ' / 会計だけ契約:' + accOnly);
}

// ============================================================
//  Phase1c: 店舗別インサイト集計 → reservation_insights.json
//  予約CSV単体で集計（契約判定はC50＝会計と99.8%一致で実証済み）
// ============================================================

// キーワード辞書（knowledge/hpb_keywords.md ＋ データで判明した高転換ワード）
// ※ナレッジ自動成長フェーズで converting_keywords.json とマージ予定
const RESV_KEYWORDS = [
  'ニキビ','毛穴','いちご鼻','黒ずみ','引き締め','ざらつき','肌質改善','肌管理','ポテンツァ',
  'ハーブピーリング','剥離なし','剥離あり','ダウンタイム','当日メイク','水光肌','陶器肌','白玉肌',
  '満足度','口コミ','カウンセリング','小顔','むくみ','毛穴洗浄','撃退','撃ち放題','お試し','専門',
  '回数券','フェイシャル','痩身','脱毛','メンズ','学割'
];

function rate1_(num, den) { return den > 0 ? Math.round(num / den * 1000) / 10 : 0; }

// ステータス優先度（重複排除で代表行を選ぶ用：完了＞キャンセル＞受付待ち＞サロンC）
function statusPriority_(s) {
  if (RESV_STATUS_VISITED.indexOf(s) >= 0) return 5;
  if (s === '施術中') return 4;
  if (RESV_STATUS_CXL_CUST.indexOf(s) >= 0 || RESV_STATUS_CXL_NOSHOW.indexOf(s) >= 0) return 3;
  if (s === '受付待ち') return 2;
  if (RESV_STATUS_CXL_SALON.indexOf(s) >= 0) return 1;  // サロンCは別ブランド自動キャンセルの産物が多い→最下位
  return 0;
}

// 重複排除：電話予約のブランド複製(LIME/Lift/Belle)を1予約に畳む
//   キー = フリガナ(空なら漢字名) | 来店日 | 開始時間 | 略称
//   グループ内は最優先ステータスの行を代表に採用
function dedupeBookings_(rows) {
  const groups = {};
  rows.forEach(function(r){
    const who = r.kana || r.nameKanji || '?';
    const id = who + '|' + r.visitDate + '|' + r.startTime + '|' + r.short;
    (groups[id] = groups[id] || []).push(r);
  });
  const out = [];
  Object.keys(groups).forEach(function(id){
    const g = groups[id];
    g.sort(function(a,b){ return statusPriority_(b.status) - statusPriority_(a.status); });
    out.push(g[0]);
  });
  return out;
}

// 店舗別＋全店インサイトを構築して返す
function buildReservationInsights_(yyyymm) {
  const ss = getReserveSS_();
  const alias = loadReserveAlias_(ss);
  const months = listReserveMonthSheets_(ss);
  if (!months.length) throw new Error('予約データタブがありません');
  const res = readReservationRows_(ss, months[months.length - 1], alias);
  const rawRows = res.rows;
  const rows = dedupeBookings_(rawRows);   // ★ブランド複製を畳んだ「予約」単位

  // 来店日の最頻YYYYMM
  if (!yyyymm) {
    const ymCount = {};
    rows.forEach(function(r){ const ym=(r.visitDate||'').slice(0,6); if(ym) ymCount[ym]=(ymCount[ym]||0)+1; });
    yyyymm = Object.keys(ymCount).sort(function(a,b){return ymCount[b]-ymCount[a];})[0] || '';
  }

  // 店舗ごとに予約を束ねる
  const byStore = {};
  rows.forEach(function(r){ (byStore[r.short] = byStore[r.short] || []).push(r); });

  const isHpb = function(r){ return r.route === 'HOT PEPPER Beauty'; };
  const stores = {};

  Object.keys(byStore).forEach(function(short){
    const list = byStore[short];
    const total = list.length;
    const visited = list.filter(function(r){ return resvIsVisited_(r.status); });
    const cust  = list.filter(function(r){ return RESV_STATUS_CXL_CUST.indexOf(r.status)>=0; }).length;
    const salon = list.filter(function(r){ return RESV_STATUS_CXL_SALON.indexOf(r.status)>=0; }).length;
    const nosh  = list.filter(function(r){ return RESV_STATUS_CXL_NOSHOW.indexOf(r.status)>=0; }).length;
    const hpbList = list.filter(isHpb);
    const hpbCxl  = hpbList.filter(function(r){ return RESV_STATUS_CXL_CUST.indexOf(r.status)>=0 || RESV_STATUS_CXL_SALON.indexOf(r.status)>=0 || RESV_STATUS_CXL_NOSHOW.indexOf(r.status)>=0; }).length;
    const contract = visited.filter(function(r){ return resvIsContract_(r.actMenu); }).length;
    // 男女比（HPB予約のみ＝性別が取れる）
    const genF = list.filter(function(r){ return r.gender==='女性'; }).length;
    const genM = list.filter(function(r){ return r.gender==='男性'; }).length;
    // 指名
    const shimei = list.filter(function(r){ return r.nominate==='指名予約'; }).length;
    // リスケ（第一希望日があり来店日と異なる）
    const resched = list.filter(function(r){ return r.hope1Date && r.visitDate && r.hope1Date!==r.visitDate; }).length;
    // 初回
    const firstYes = visited.filter(function(r){ return /はい/.test(r.firstVisit); });
    const firstContract = firstYes.filter(function(r){ return resvIsContract_(r.actMenu); }).length;

    stores[short] = {
      total_reservations: total,
      visited:            visited.length,
      cancel_rate:        rate1_(cust+salon+nosh, total),
      cancel_customer:    cust,
      cancel_salon:       salon,
      cancel_noshow:      nosh,
      hpb_cancel_rate:    rate1_(hpbCxl, hpbList.length),
      hpb_rate:           rate1_(hpbList.length, total),
      shimei_rate:        rate1_(shimei, total),
      gender_female:      genF,
      gender_male:        genM,
      gender_female_pct:  rate1_(genF, genF+genM),
      reschedule:         resched,
      contract_count:     contract,
      contract_rate:      rate1_(contract, visited.length),
      first_visit_count:  firstYes.length,
      first_contract_rate:rate1_(firstContract, firstYes.length)
    };
  });

  // 全店キーワード別 契約率（来店完了が母数）
  const visitedAll = rows.filter(function(r){ return resvIsVisited_(r.status); });
  const kw = [];
  RESV_KEYWORDS.forEach(function(k){
    const m = visitedAll.filter(function(r){ return (r.rsvCoupon.indexOf(k)>=0) || (r.rsvMenu.indexOf(k)>=0); });
    if (m.length >= 15) {
      const c = m.filter(function(r){ return resvIsContract_(r.actMenu); }).length;
      kw.push({ keyword: k, n: m.length, contract: c, rate: rate1_(c, m.length) });
    }
  });
  kw.sort(function(a,b){ return b.rate - a.rate; });

  // 全店クーポン別 契約率（上位）
  const couponMap = {};
  visitedAll.forEach(function(r){
    if (!r.rsvCoupon) return;
    const c = couponMap[r.rsvCoupon] = couponMap[r.rsvCoupon] || { n:0, contract:0 };
    c.n++; if (resvIsContract_(r.actMenu)) c.contract++;
  });
  const coupons = Object.keys(couponMap).filter(function(k){ return couponMap[k].n>=20; })
    .map(function(k){ return { coupon:k, n:couponMap[k].n, contract:couponMap[k].contract, rate:rate1_(couponMap[k].contract, couponMap[k].n) }; })
    .sort(function(a,b){ return b.rate - a.rate; });

  return {
    updated_at: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'),
    month: yyyymm,
    source_rows: res.totalRows,
    mapped_rows: rawRows.length,
    bookings: rows.length,        // 重複排除後の予約数
    stores: stores,
    overall: {
      visited: visitedAll.length,
      contract: visitedAll.filter(function(r){ return resvIsContract_(r.actMenu); }).length,
      keywords: kw,
      coupons: coupons.slice(0, 30)
    }
  };
}

// 汎用：JSONをGitHubにcommit（exportToGitHub と同方式）
function commitJsonToGitHub_(path, obj, message) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('GITHUB_TOKEN');
  const owner = props.getProperty('GITHUB_OWNER') || 'kimu2-lime';
  const repo  = props.getProperty('GITHUB_REPO')  || 'dashboard-master';
  if (!token) { Logger.log('❌ GITHUB_TOKEN が未設定'); return false; }
  const apiBase = 'https://api.github.com/repos/' + owner + '/' + repo + '/contents/' + path;
  const headers = { 'Authorization':'token '+token, 'Accept':'application/vnd.github.v3+json', 'User-Agent':'GAS-LIME-FIT' };
  let sha = null;
  try {
    const r = UrlFetchApp.fetch(apiBase, { headers: headers, muteHttpExceptions: true });
    if (r.getResponseCode() === 200) sha = JSON.parse(r.getContentText()).sha;
  } catch(e) { Logger.log('SHA取得エラー: ' + e); }
  const body = { message: message, content: Utilities.base64Encode(JSON.stringify(obj), Utilities.Charset.UTF_8) };
  if (sha) body.sha = sha;
  const put = UrlFetchApp.fetch(apiBase, { method:'PUT', headers:headers, payload:JSON.stringify(body), muteHttpExceptions:true });
  const code = put.getResponseCode();
  if (code===200 || code===201) { Logger.log('✅ push成功: ' + path); return true; }
  Logger.log('❌ push失敗 HTTP ' + code + ': ' + put.getContentText().slice(0,200)); return false;
}

// 本番：集計してGitHubに出力
function exportReservationInsights() {
  const data = buildReservationInsights_(null);
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  commitJsonToGitHub_('data/reservation_insights.json', data, '📈 予約インサイト更新 ' + now);
}

// ★検証用：集計結果をログで確認（commitしない）
function testBuildInsights() {
  const data = buildReservationInsights_(null);
  Logger.log('対象月: ' + data.month + ' / 元行: ' + data.source_rows + ' / マップ: ' + data.mapped_rows + ' / 重複排除後の予約: ' + data.bookings);
  Logger.log('全店 来店完了: ' + data.overall.visited + ' / 契約: ' + data.overall.contract +
             ' (' + rate1_(data.overall.contract, data.overall.visited) + '%)');
  Logger.log('▼キーワード別 契約率（上位15）:');
  data.overall.keywords.slice(0, 15).forEach(function(k){
    Logger.log('   ' + k.rate + '%  (n=' + k.n + ' 契=' + k.contract + ')  ' + k.keyword);
  });
  // サンプル店舗（件数最大）のインサイト
  const top = Object.keys(data.stores).sort(function(a,b){ return data.stores[b].total_reservations - data.stores[a].total_reservations; })[0];
  Logger.log('▼サンプル店舗「' + top + '」: ' + JSON.stringify(data.stores[top]));
  Logger.log('（店舗数: ' + Object.keys(data.stores).length + '）');
}

