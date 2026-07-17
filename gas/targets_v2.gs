// ============================================================
//  加盟店目標 V2 ── 縦持ち・1タブ集約システム
// ------------------------------------------------------------
//  ねらい：
//   ・担当者ごとのタブ分割をやめ、1タブで全担当者・全月を縦持ち管理
//   ・1行＝「年月 × 店舗」。月を追加するときは下に行を追記していく
//   ・KPIの定義は KV2_DEFS（レジストリ）1か所に集約 → 列追加は1行足すだけ
//   ・列はヘッダー名で読む（位置固定をやめる）ので、列の追加・並べ替えで壊れない
//   ・突合キー（略称）＝ ブランド ＋ 店舗名（スペース無し連結）。実績データと一致する
//
//  【スクリプトプロパティ】
//   SS_KAMEL_V2 : この目標タブ／本部収益実績タブを置くスプシID（共有スプシ）
//                 ※ 未設定なら SS_KAMEL にフォールバック。
//                    移行中は SS_KAMEL_V2 だけ新スプシに設定し、本番(SS_KAMEL=旧)を
//                    動かしたままテストできる。Stage 4 で正式に切替える。
//   SS_MASTER   : 店舗一覧マスター（フル名→略称の対応・加盟店リスト取得元）
//
//  ※ Stage 1〜3（このファイル）は生成系のみ。既存 main.gs の読み取りには影響しない。
//    パーサ差し替え（Stage 4）は別途、シート生成・入力確認後に行う。
// ============================================================

const KV2 = {
  SHEET_TARGET:    '目標',           // 縦持ち目標タブ名
  SHEET_HQ_ACTUAL: '本部収益実績',    // 本部収益の実績入力タブ名
  HEADER_ROW:      1,
  DATA_START_ROW:  2,
  BRANDS:          ['Belle', 'LIME', 'Lift', 'LABO'],
  // 雛形セットアップ時に初期生成する月（必要に応じて編集／後から addKamelV2Month で追記可）
  INIT_MONTHS:     ['2026年4月', '2026年5月', '2026年6月']
};

// 目標タブの隠し列ヘッダー：各店舗行の「月」を保持し、サマリの月別SUMIFSの突合キーにする。
// 表示は月バナーのみ・各行の年月は空、という見た目を保ったまま、サマリ側は数式で自動反映できる。
const KV2_MONTHKEY_LABEL = '_月キー';

// V2 用スプシ取得（SS_KAMEL_V2 優先・無ければ SS_KAMEL にフォールバック）
function getSSKamelV2_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SS_KAMEL_V2');
  if (id) return SpreadsheetApp.openById(id);
  return getSS('SS_KAMEL');
}

// ヘッダー名から列番号(1始まり)を返す。見つからなければ0。※列の並べ替えに強い
function kv2FindCol_(sheet, label) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return 0;
  const header = sheet.getRange(KV2.HEADER_ROW, 1, 1, lastCol).getValues()[0];
  for (let i = 0; i < header.length; i++) {
    if (String(header[i]).trim() === label) return i + 1;
  }
  return 0;
}

// "202604" → "2026年4月"
function kv2YYYYMMToLabel_(ym) {
  const s = String(ym);
  if (s.length < 6) return s;
  return s.slice(0, 4) + '年' + parseInt(s.slice(4), 10) + '月';
}

// 「_月キー」ヘッダーを持つ列を全部返す（並べ替えで重複が出ても検知できる）
function kv2AllMonthKeyCols_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return [];
  const header = sheet.getRange(KV2.HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const cols = [];
  for (let i = 0; i < header.length; i++) {
    if (String(header[i]).trim() === KV2_MONTHKEY_LABEL) cols.push(i + 1);
  }
  return cols;
}

// 実データ（月ラベル）が最も多く入っている「_月キー」列を返す。重複時の取り違え防止。
function kv2MonthKeyCol_(sheet) {
  const cols = kv2AllMonthKeyCols_(sheet);
  if (cols.length <= 1) return cols[0] || 0;
  const lastRow = sheet.getLastRow();
  let best = cols[0], bestCount = -1;
  if (lastRow >= KV2.DATA_START_ROW) {
    cols.forEach(function(c){
      const vals = sheet.getRange(KV2.DATA_START_ROW, c, lastRow - KV2.HEADER_ROW, 1).getValues();
      let cnt = 0;
      for (let k = 0; k < vals.length; k++) if (String(vals[k][0] || '').trim()) cnt++;
      if (cnt > bestCount) { bestCount = cnt; best = c; }
    });
  }
  return best;
}

// 診断：_月キー列の状況をログ出力（取り違え・重複の確認用）
function kv2DebugMonthKey() {
  const ss = getSSKamelV2_();
  const sheet = ss.getSheetByName(KV2.SHEET_TARGET);
  if (!sheet) { Logger.log('目標タブがありません'); return; }
  const lastCol = sheet.getLastColumn(), lastRow = sheet.getLastRow();
  Logger.log('lastRow=' + lastRow + ' / lastCol=' + lastCol);
  const cols = kv2AllMonthKeyCols_(sheet);
  Logger.log('「' + KV2_MONTHKEY_LABEL + '」ヘッダーのある列: ' + (cols.length ? cols.join(', ') : 'なし'));
  cols.forEach(function(c){
    const vals = sheet.getRange(KV2.DATA_START_ROW, c, Math.max(lastRow - 1, 1), 1).getValues();
    let cnt = 0; const sample = [];
    for (let k = 0; k < vals.length; k++) {
      const v = String(vals[k][0] || '').trim();
      if (v) { cnt++; if (sample.length < 3) sample.push(v); }
    }
    Logger.log('  列' + c + ': 非空 ' + cnt + ' 件  例: ' + sample.join(' / '));
  });
  Logger.log('→ 採用される列(kv2MonthKeyCol_): ' + kv2MonthKeyCol_(sheet));
}

// ============================================================
//  KPI定義レジストリ（このシステムの単一の真実）
//   type : 'id'（識別列） | 'manual'（手入力） | 'auto'（数式自動計算）
//   tkey : data.json の targets/actual_data 側のキー（突合用・nullは取込対象外）
//   fmt  : 'int' | 'pct' | 'yen' | 'text'
//   formula(col): 自動列のA1数式を返す関数。col(label)で他列の列文字を取得
//  ※ 列の並びはこの配列順。追加・並べ替えはここを編集するだけ。
// ============================================================
const KV2_DEFS = [
  // ── 識別列 ──
  { label: '年月',               type: 'id', tkey: null,                        fmt: 'text', group: '識別' },
  { label: '担当者',             type: 'id', tkey: null,                        fmt: 'text', group: '識別' },
  { label: 'ブランド',           type: 'id', tkey: null,                        fmt: 'text', group: '識別' },
  { label: '店舗名',             type: 'id', tkey: null,                        fmt: 'text', group: '識別' },
  { label: 'ロイヤリティパターン', type: 'id', tkey: null,                       fmt: 'text', group: '識別' },

  // ── 左寄せ集計（一番見たい結果を識別列の直後に配置）──
  { label: '合計売上',           type: 'auto',   tkey: 'total_sales',           fmt: 'yen', group: '集計',
    formula: function(c){ return '=' + c('新規契約売上') + '+' + c('再アポ契約売上') + '+' + c('継続契約売上')
                                 + '+' + c('クーポン売上') + '+' + c('オプション売上') + '+' + c('物販売上'); } },
  { label: 'ロイヤリティ_固定含', type: 'auto',  tkey: 'royalty',               fmt: 'yen', group: '集計',
    formula: function(c){ return buildRoyaltyFormula(c('ロイヤリティパターン'), c('合計売上')); } },
  { label: '商材売上',           type: 'manual', tkey: 'shozai_sales',          fmt: 'yen', group: '本部収益' },

  // ── 新規系 ──
  { label: '新規数',             type: 'manual', tkey: 'new_count',             fmt: 'int', group: '新規' },
  { label: '新規契約率',         type: 'manual', tkey: 'new_contract_rate',     fmt: 'pct', group: '新規' },
  { label: '新規契約単価',       type: 'manual', tkey: 'new_contract_unit_price', fmt: 'yen', group: '新規' },
  { label: '新規契約数',         type: 'auto',   tkey: 'new_contract_count',    fmt: 'int', group: '新規',
    formula: function(c){ return '=ROUND(' + c('新規数') + '*' + c('新規契約率') + ',0)'; } },
  { label: '新規契約売上',       type: 'auto',   tkey: 'new_contract_sales',    fmt: 'yen', group: '新規',
    formula: function(c){ return '=' + c('新規契約数') + '*' + c('新規契約単価'); } },

  // ── 再アポ系 ──
  { label: '再アポ数',           type: 'manual', tkey: 'reappo_count',          fmt: 'int', group: '再アポ' },
  { label: '再アポ契約率',       type: 'manual', tkey: 'reappo_contract_rate',  fmt: 'pct', group: '再アポ' },
  { label: '再アポ契約単価',     type: 'manual', tkey: 'reappo_contract_unit_price', fmt: 'yen', group: '再アポ' },
  { label: '再アポ契約数',       type: 'auto',   tkey: 'reappo_contract_count', fmt: 'int', group: '再アポ',
    formula: function(c){ return '=ROUND(' + c('再アポ数') + '*' + c('再アポ契約率') + ',0)'; } },
  { label: '再アポ契約売上',     type: 'auto',   tkey: 'reappo_contract_sales', fmt: 'yen', group: '再アポ',
    formula: function(c){ return '=' + c('再アポ契約数') + '*' + c('再アポ契約単価'); } },

  // ── 継続系 ──
  { label: '継続対象数',         type: 'manual', tkey: 'keizoku_end_count',     fmt: 'int', group: '継続' },
  { label: '継続契約率',         type: 'manual', tkey: 'keizoku_rate',          fmt: 'pct', group: '継続' },
  { label: '継続契約単価',       type: 'manual', tkey: 'keizoku_continue_unit_price', fmt: 'yen', group: '継続' },
  { label: '継続契約数',         type: 'auto',   tkey: 'keizoku_continue_count', fmt: 'int', group: '継続',
    formula: function(c){ return '=ROUND(' + c('継続対象数') + '*' + c('継続契約率') + ',0)'; } },
  { label: '継続契約売上',       type: 'auto',   tkey: 'keizoku_continue_sales', fmt: 'yen', group: '継続',
    formula: function(c){ return '=' + c('継続契約数') + '*' + c('継続契約単価'); } },

  // ── 店舗売上 構成要素 ──
  { label: 'クーポン単価',       type: 'manual', tkey: null,                    fmt: 'yen', group: '構成' },
  { label: 'クーポン売上',       type: 'auto',   tkey: 'coupon_sales',          fmt: 'yen', group: '構成',
    formula: function(c){ return '=' + c('新規数') + '*' + c('クーポン単価'); } },
  { label: 'オプション売上',     type: 'manual', tkey: 'option_sales',          fmt: 'yen', group: '構成' },
  { label: '物販売上',           type: 'manual', tkey: 'buppan_sales',          fmt: 'yen', group: '構成' },

  // ── 集計（自動・参考列）──
  { label: 'ロイヤリティ_純%',   type: 'auto',   tkey: null,                    fmt: 'yen', group: '集計',
    formula: function(c){ return buildRoyaltyPureFormula(c('ロイヤリティパターン'), c('合計売上')); } },
  { label: '新規+再アポ契約数',  type: 'auto',   tkey: 'total_contract_count',  fmt: 'int', group: '集計',
    formula: function(c){ return '=' + c('新規契約数') + '+' + c('再アポ契約数'); } },
  { label: '新規+再アポ契約率',  type: 'auto',   tkey: 'total_contract_rate',   fmt: 'pct', group: '集計',
    formula: function(c){ return '=IFERROR((' + c('新規契約数') + '+' + c('再アポ契約数') + ')/('
                                 + c('新規数') + '+' + c('再アポ数') + '),0)'; } },

  // ── 本部収益目標（手入力・発生月のみ入力でOK／商材売上は左の集計ブロックへ移動済み）──
  { label: '追加加盟売上',       type: 'manual', tkey: 'add_franchise_sales',   fmt: 'yen', group: '本部収益' },
  { label: '追加マシン購入売上', type: 'manual', tkey: 'add_machine_sales',     fmt: 'yen', group: '本部収益' },
  { label: 'マシン等レンタル売上', type: 'manual', tkey: 'machine_rental_sales', fmt: 'yen', group: '本部収益' },
  { label: '研修売上',           type: 'manual', tkey: 'training_sales',        fmt: 'yen', group: '本部収益' },
  { label: '採用売上',           type: 'manual', tkey: 'recruit_sales',         fmt: 'yen', group: '本部収益' },
  { label: 'その他',             type: 'manual', tkey: 'other_hq_sales',        fmt: 'yen', group: '本部収益' }
];

// 本部収益 実績入力タブの列（手入力）。tkeyは目標側と揃える
const KV2_HQ_DEFS = [
  { label: '年月',               type: 'id', tkey: null,                       fmt: 'text' },
  { label: '担当者',             type: 'id', tkey: null,                       fmt: 'text' },
  { label: 'ブランド',           type: 'id', tkey: null,                       fmt: 'text' },
  { label: '店舗名',             type: 'id', tkey: null,                       fmt: 'text' },
  { label: '商材売上',           type: 'manual', tkey: 'shozai_sales',         fmt: 'yen' },
  { label: '追加加盟売上',       type: 'manual', tkey: 'add_franchise_sales',  fmt: 'yen' },
  { label: '追加マシン購入売上', type: 'manual', tkey: 'add_machine_sales',    fmt: 'yen' },
  { label: 'マシン等レンタル売上', type: 'manual', tkey: 'machine_rental_sales', fmt: 'yen' },
  { label: '研修売上',           type: 'manual', tkey: 'training_sales',       fmt: 'yen' },
  { label: '採用売上',           type: 'manual', tkey: 'recruit_sales',        fmt: 'yen' },
  { label: 'その他',             type: 'manual', tkey: 'other_hq_sales',       fmt: 'yen' }
];

// ============================================================
//  ロイヤリティ「純%（固定加算・下限を除いた純粋な%）」数式
//   ・固定パターンは純%対象外 → 空欄
//   ・"10%+4.4万" は +4.4万を落とした 10%
//   ・"15%（最低13.2万）" は下限を落とした 15%
//   ※ シート上の参考表示用。ダッシュボードには出さない
// ============================================================
function buildRoyaltyPureFormula(p, s) {
  return '=IF(' + s + '=0,0,' +
    'IF(REGEXMATCH(IFERROR(' + p + ',""),"固定"),"",' +
    'IF(REGEXMATCH(IFERROR(' + p + ',""),"段階|70万"),' +
      'IF(' + s + '<700000,ROUND(' + s + '*0.08,0),' +
      'IF(' + s + '<850000,ROUND(' + s + '*0.09,0),' +
      'IF(' + s + '<950000,ROUND(' + s + '*0.10,0),' +
      'IF(' + s + '<1000000,ROUND(' + s + '*0.12,0),' +
      'ROUND(' + s + '*0.15,0))))),' +
    'IF(REGEXMATCH(IFERROR(' + p + ',""),"15[%％]"),ROUND(' + s + '*0.15,0),' +
    'IF(REGEXMATCH(IFERROR(' + p + ',""),"90万"),' +
      'IF(' + s + '<900000,ROUND(' + s + '*0.08,0),ROUND(' + s + '*0.10,0)),' +
    'IF(REGEXMATCH(IFERROR(' + p + ',""),"8[%％]"),ROUND(' + s + '*0.08,0),' +
    'IF(REGEXMATCH(IFERROR(' + p + ',""),"5[%％]"),ROUND(' + s + '*0.05,0),' +
    'ROUND(' + s + '*0.10,0)' +
    ')))))))';
}

// ============================================================
//  ヘルパー：レジストリから「ラベル→列番号(1始まり)」マップを作る
// ============================================================
function kv2ColMap_(defs) {
  const map = {};
  defs.forEach(function(d, i) { map[d.label] = i + 1; });
  return map;
}

// ラベル→「同一行の列文字」を返す関数を生成（数式組み立て用）
function kv2RowColRef_(colMap, row) {
  return function(label) {
    const c = colMap[label];
    if (!c) throw new Error('KV2_DEFS に列が無い: ' + label);
    return columnLetter(c) + row;
  };
}

// ============================================================
//  加盟店リスト取得（店舗一覧マスターから・直営を除く）
//   返り値: [{short, brand, location, person, roy}, ...]
//   略称→ブランド/店舗名 はブランド接頭辞で分解
// ============================================================
function kv2GetKamelStores_() {
  const ssMaster  = getSS('SS_MASTER');
  const listSheet = ssMaster ? ssMaster.getSheetByName('店舗一覧') : null;
  const out = [];
  if (!listSheet) return out;

  const data = listSheet.getDataRange().getValues();
  const seen = {};
  for (let i = 1; i < data.length; i++) {
    const short  = data[i][1] ? data[i][1].toString().trim() : '';
    const type   = data[i][2] ? data[i][2].toString().trim() : '';
    const person = data[i][3] ? data[i][3].toString().trim() : '';
    const roy    = data[i][4] ? data[i][4].toString().trim() : '';
    if (!short || type === '直営' || seen[short]) continue;
    seen[short] = true;

    // 略称をブランド接頭辞で分解（例: "Belle溝の口" → brand="Belle", location="溝の口"）
    let brand = '', location = short;
    for (let b = 0; b < KV2.BRANDS.length; b++) {
      if (short.indexOf(KV2.BRANDS[b]) === 0) {
        brand = KV2.BRANDS[b];
        location = short.slice(brand.length).replace(/^[\s　]+/, '');
        break;
      }
    }
    out.push({ short: short, brand: brand, location: location, person: person, roy: roy });
  }
  return out;
}

// 担当者フルネーム候補（プルダウン用）。マスターの担当者＋既知の担当者
function kv2PersonList_(stores) {
  const set = {};
  stores.forEach(function(s){ if (s.person && s.person !== '担当なし') set[s.person] = true; });
  return Object.keys(set);
}

// ============================================================
//  Stage 2：目標タブのセットアップ（ヘッダー作成）
// ============================================================
function setupKamelV2() {
  const ss = getSSKamelV2_();
  let sheet = ss.getSheetByName(KV2.SHEET_TARGET);
  if (sheet) {
    throw new Error('「' + KV2.SHEET_TARGET + '」タブは既に存在します。作り直す場合は手動で削除してから実行してください（入力データ保護のため自動削除しません）。');
  }
  sheet = ss.insertSheet(KV2.SHEET_TARGET);

  const colMap   = kv2ColMap_(KV2_DEFS);
  const totalCols = KV2_DEFS.length;

  // ヘッダー行
  const headers = KV2_DEFS.map(function(d){ return d.label; });
  sheet.getRange(KV2.HEADER_ROW, 1, 1, totalCols).setValues([headers])
    .setBackground('#2d3561').setFontColor('#ffffff').setFontWeight('bold').setFontSize(10)
    .setVerticalAlignment('middle').setWrap(true);

  // グループ別にヘッダー背景を薄く色分け（視認性）
  const groupColor = { '識別':'#1a1a2e', '新規':'#243b6b', '再アポ':'#2b5a4a', '継続':'#5a3b2b', '構成':'#4a4a6b', '集計':'#3a2b5a', '本部収益':'#6b2b4a' };
  KV2_DEFS.forEach(function(d, i){
    if (groupColor[d.group]) sheet.getRange(KV2.HEADER_ROW, i + 1).setBackground(groupColor[d.group]);
  });

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(colMap['商材売上']); // 識別列＋左寄せ集計（合計売上・ロイヤリティ・商材売上）まで固定
  sheet.setColumnWidth(colMap['年月'], 90);
  sheet.setColumnWidth(colMap['担当者'], 90);
  sheet.setColumnWidth(colMap['ブランド'], 70);
  sheet.setColumnWidth(colMap['店舗名'], 110);
  sheet.setColumnWidth(colMap['ロイヤリティパターン'], 150);
  sheet.setColumnWidth(colMap['合計売上'], 100);
  sheet.setColumnWidth(colMap['ロイヤリティ_固定含'], 100);
  sheet.setColumnWidth(colMap['商材売上'], 90);

  // 隠し列「_月キー」：サマリの月別SUMIFS用。totalCols+1 に置いて非表示にする
  const keyCol = totalCols + 1;
  sheet.getRange(KV2.HEADER_ROW, keyCol).setValue(KV2_MONTHKEY_LABEL);
  sheet.hideColumns(keyCol);

  Logger.log('✅ 目標タブのヘッダーを作成しました。続けて月を追加するには addKamelV2Month("2026年6月") を実行してください。');
  return sheet;
}

// セットアップ＋初期月（KV2.INIT_MONTHS）をまとめて生成する便利関数
function setupKamelV2WithMonths() {
  setupKamelV2();
  KV2.INIT_MONTHS.forEach(function(m){ addKamelV2Month(m); });
  Logger.log('✅ 目標タブ作成＋初期月（' + KV2.INIT_MONTHS.join(', ') + '）を生成しました。');
}

// ============================================================
//  Stage 2：指定月の店舗行をタブ下部に追記（縦持ち）
//   ・加盟店マスターから全店舗を展開し、識別列をプリフィル
//   ・手入力列は空欄（黄）、自動列は数式
//   ・追記したブロックを行グループ化（過去月は折りたたみ可能）
// ============================================================
function addKamelV2Month(monthLabel) {
  if (!monthLabel) throw new Error('月ラベルを指定してください（例: "2026年6月"）');
  const ss    = getSSKamelV2_();
  const sheet = ss.getSheetByName(KV2.SHEET_TARGET);
  if (!sheet) throw new Error('先に setupKamelV2() を実行してください。');

  const colMap    = kv2ColMap_(KV2_DEFS);
  const totalCols = KV2_DEFS.length;
  const stores    = kv2GetKamelStores_();
  if (stores.length === 0) throw new Error('店舗一覧マスターから加盟店を取得できませんでした（SS_MASTERの店舗一覧タブを確認）。');

  // 既に同月が存在しないかチェック（バナー行の年月を YYYYMM 比較で走査）
  const lastRow  = sheet.getLastRow();
  const targetYM = kv2MonthToYYYYMM_(monthLabel);
  if (lastRow >= KV2.DATA_START_ROW) {
    const monthCol = sheet.getRange(KV2.DATA_START_ROW, colMap['年月'], lastRow - KV2.HEADER_ROW, 1).getValues();
    for (let i = 0; i < monthCol.length; i++) {
      if (monthCol[i][0] && kv2MonthToYYYYMM_(monthCol[i][0]) === targetYM) {
        throw new Error('「' + monthLabel + '」は既に存在します。');
      }
    }
  }

  const persons = kv2PersonList_(stores);

  // ── 月バナー行（年月は1ブロック1回だけここに書く／空行を1つ空けてから）──
  const bannerRow = lastRow >= KV2.DATA_START_ROW ? lastRow + 2 : KV2.DATA_START_ROW; // 既存があれば空行を1つはさむ
  sheet.getRange(bannerRow, colMap['年月']).setValue('▼ ' + monthLabel);
  sheet.getRange(bannerRow, 1, 1, totalCols)
    .setBackground('#2d3561').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11)
    .setVerticalAlignment('middle');
  sheet.setRowHeight(bannerRow, 26);

  const startRow = bannerRow + 1;

  // プルダウン
  const brandRule  = SpreadsheetApp.newDataValidation().requireValueInList(KV2.BRANDS, true).setAllowInvalid(true).build();
  const royRule    = SpreadsheetApp.newDataValidation().requireValueInList(ROYALTY_PATTERNS, true).setAllowInvalid(true).build();
  const personRule = persons.length > 0
    ? SpreadsheetApp.newDataValidation().requireValueInList(persons, true).setAllowInvalid(true).build()
    : null;

  // 行データ（識別列のみ値・年月は空＝バナーからfill-downで継承）。自動列は後で setFormula
  const values = [];
  for (let i = 0; i < stores.length; i++) {
    const row = new Array(totalCols).fill('');
    // 年月セルは意図的に空（バナー行の月を fill-down で適用）
    row[colMap['担当者'] - 1]             = stores[i].person || '';
    row[colMap['ブランド'] - 1]           = stores[i].brand || '';
    row[colMap['店舗名'] - 1]             = stores[i].location || '';
    row[colMap['ロイヤリティパターン'] - 1] = stores[i].roy || '';
    values.push(row);
  }
  sheet.getRange(startRow, 1, values.length, totalCols).setValues(values);

  // 隠し列「_月キー」に月ラベルを各行へ（非表示・サマリの月別SUMIFS突合キー）
  const keyCol = totalCols + 1;
  if (sheet.getRange(KV2.HEADER_ROW, keyCol).getValue() !== KV2_MONTHKEY_LABEL) {
    sheet.getRange(KV2.HEADER_ROW, keyCol).setValue(KV2_MONTHKEY_LABEL);
    sheet.hideColumns(keyCol);
  }
  const keyVals = [];
  for (let i = 0; i < stores.length; i++) keyVals.push([monthLabel]);
  sheet.getRange(startRow, keyCol, stores.length, 1).setValues(keyVals);

  // 自動列：1行ずつ数式（行番号が変わるため）
  for (let i = 0; i < stores.length; i++) {
    const r = startRow + i;
    const ref = kv2RowColRef_(colMap, r);
    KV2_DEFS.forEach(function(d){
      if (d.type === 'auto' && typeof d.formula === 'function') {
        sheet.getRange(r, colMap[d.label]).setFormula(d.formula(ref));
      }
    });
  }

  const endRow = startRow + stores.length - 1;

  // プルダウン適用
  sheet.getRange(startRow, colMap['ブランド'], stores.length, 1).setDataValidation(brandRule);
  sheet.getRange(startRow, colMap['ロイヤリティパターン'], stores.length, 1).setDataValidation(royRule);
  if (personRule) sheet.getRange(startRow, colMap['担当者'], stores.length, 1).setDataValidation(personRule);

  // 書式：手入力＝黄背景・青字／自動＝灰背景・斜体／%・¥フォーマット
  KV2_DEFS.forEach(function(d){
    const rng = sheet.getRange(startRow, colMap[d.label], stores.length, 1);
    if (d.type === 'manual') rng.setBackground('#fffde7').setFontColor('#1565c0');
    if (d.type === 'auto')   rng.setBackground('#f5f5f5').setFontStyle('italic');
    if (d.fmt === 'yen')     rng.setNumberFormat('¥#,##0');
    if (d.fmt === 'pct')     rng.setNumberFormat('0.0%');
    if (d.fmt === 'int')     rng.setNumberFormat('#,##0');
  });

  // 月ブロック（店舗行）を行グループ化（バナーは残し、明細を折りたためる）
  sheet.getRange(startRow, 1, stores.length, 1).shiftRowGroupDepth(1);

  // 担当者サマリを最新化
  buildKamelSummaryV2();

  Logger.log('✅ 「' + monthLabel + '」を ' + stores.length + ' 店舗ぶん追記しました（バナー行 ' + bannerRow + ' / 明細 ' + startRow + '〜' + endRow + '）。');
}

// ============================================================
//  前月ブロックを丸ごと複製して新しい月を作る（前月の入力値ごとコピー）
//   ・fromLabel 省略時は「最新月」を自動採用
//   ・年月バナー・_月キー(AK)を新月に書き換え、サマリも自動再生成
//   ・自動列の数式は copyTo の相対参照調整で正しく追従する
//   用途：前月をベースに微調整して当月目標を作る運用
// ============================================================
function copyKamelV2Month(toLabel, fromLabel) {
  if (!toLabel) throw new Error('追加する月ラベルを指定してください（例: "2026年7月"）');
  const ss    = getSSKamelV2_();
  const sheet = ss.getSheetByName(KV2.SHEET_TARGET);
  if (!sheet) throw new Error('先に setupKamelV2() を実行してください。');

  // 列はヘッダー名で解決（位置固定をやめる＝列の並べ替えに強い／_月キーはデータのある列を選ぶ）
  const keyCol  = kv2MonthKeyCol_(sheet);
  const ymCol   = kv2FindCol_(sheet, '年月');
  if (!keyCol) throw new Error('目標タブに「' + KV2_MONTHKEY_LABEL + '」列がありません。先に repairKamelV2MonthKey を実行してください。');
  if (!ymCol)  throw new Error('目標タブに「年月」列がありません。');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < KV2.DATA_START_ROW) {
    throw new Error('まだ月がありません。先に setupKamelV2WithMonths を実行してください。');
  }

  // _月キー列を読み、月ラベル→行範囲（1始まり）を把握
  const keyVals    = sheet.getRange(KV2.DATA_START_ROW, keyCol, lastRow - KV2.HEADER_ROW, 1).getValues();
  const monthRows  = {};
  const monthOrder = [];
  for (let i = 0; i < keyVals.length; i++) {
    const lab = String(keyVals[i][0] || '').trim();
    if (!lab) continue;
    const row = KV2.DATA_START_ROW + i;
    if (!monthRows[lab]) { monthRows[lab] = { first: row, last: row }; monthOrder.push(lab); }
    else monthRows[lab].last = row;
  }
  if (monthOrder.length === 0) {
    throw new Error('「' + KV2_MONTHKEY_LABEL + '」列が空です。repairKamelV2MonthKey を実行して月キーを再構築してください。');
  }

  // 重複チェック（YYYYMM比較）
  const toYM = kv2MonthToYYYYMM_(toLabel);
  for (let i = 0; i < monthOrder.length; i++) {
    if (kv2MonthToYYYYMM_(monthOrder[i]) === toYM) {
      throw new Error('「' + toLabel + '」は既に存在します。');
    }
  }

  // コピー元（既定＝最新月＝ym最大）
  let src = fromLabel;
  if (!src) {
    src = monthOrder.slice().sort(function(a, b){
      const ya = kv2MonthToYYYYMM_(a) || '', yb = kv2MonthToYYYYMM_(b) || '';
      return ya < yb ? -1 : ya > yb ? 1 : 0;
    }).pop();
  }
  if (!monthRows[src]) throw new Error('コピー元の月「' + src + '」が見つかりません。');

  const srcFirst   = monthRows[src].first;
  const srcLast    = monthRows[src].last;
  const bannerRow  = srcFirst - 1;                 // 店舗行の1つ上＝バナー
  const blockRows  = srcLast - bannerRow + 1;      // バナー＋店舗行
  const storeCount = srcLast - srcFirst + 1;

  // 末尾の下に空行を1つはさんで複製先を決定
  const destBanner = lastRow + 2;

  // バナー〜店舗行を全列まるごとコピー（値・数式・書式・入力規則すべて／相対参照は自動調整）
  sheet.getRange(bannerRow, 1, blockRows, lastCol)
       .copyTo(sheet.getRange(destBanner, 1, blockRows, lastCol));

  // 新バナーの年月・新ブロックの_月キーを当月に書き換え（列はヘッダー解決済み）
  sheet.getRange(destBanner, ymCol).setValue('▼ ' + toLabel);
  const destStoreFirst = destBanner + 1;
  const newKeys = [];
  for (let i = 0; i < storeCount; i++) newKeys.push([toLabel]);
  sheet.getRange(destStoreFirst, keyCol, storeCount, 1).setValues(newKeys);

  // 明細を行グループ化（折りたたみ用。copyToはグループを引き継がないので付け直す）
  sheet.getRange(destStoreFirst, 1, storeCount, 1).shiftRowGroupDepth(1);

  // サマリ再生成（新月ブロック追加・過去月は自動折りたたみ）
  buildKamelSummaryV2();

  Logger.log('✅ 「' + src + '」を複製して「' + toLabel + '」を作成（バナー ' + destBanner +
             ' / 明細 ' + destStoreFirst + '〜' + (destStoreFirst + storeCount - 1) +
             '）。値は前月のコピーです。必要箇所を編集してください。');
}

// 引数なしで実行できる版：最新月の「次の月」を自動算出して前月複製で追加。
//   GASエディタで関数を選んで Run するだけ（メニューが使えない環境向け）。
function addNextKamelV2Month() {
  const ss    = getSSKamelV2_();
  const sheet = ss.getSheetByName(KV2.SHEET_TARGET);
  if (!sheet) throw new Error('先に setupKamelV2() を実行してください。');
  const keyCol  = kv2MonthKeyCol_(sheet);   // データのある_月キー列を選ぶ（重複/並べ替えに強い）
  if (!keyCol) throw new Error('目標タブに「' + KV2_MONTHKEY_LABEL + '」列がありません。先に repairKamelV2MonthKey を実行してください。');
  const lastRow = sheet.getLastRow();
  if (lastRow < KV2.DATA_START_ROW) throw new Error('まだ月がありません。先に setupKamelV2WithMonths を実行してください。');

  const keyVals = sheet.getRange(KV2.DATA_START_ROW, keyCol, lastRow - KV2.HEADER_ROW, 1).getValues();
  let maxYM = '';
  for (let i = 0; i < keyVals.length; i++) {
    const lab = String(keyVals[i][0] || '').trim();
    if (!lab) continue;
    const ym = kv2MonthToYYYYMM_(lab);
    if (ym && ym > maxYM) maxYM = ym;
  }
  if (!maxYM) throw new Error('「' + KV2_MONTHKEY_LABEL + '」列が空です。repairKamelV2MonthKey を実行して月キーを再構築してください。');

  const y = parseInt(maxYM.slice(0, 4), 10), m = parseInt(maxYM.slice(4), 10);
  const ny = m === 12 ? y + 1 : y;
  const nm = m === 12 ? 1 : m + 1;
  const nextLabel = ny + '年' + nm + '月';
  copyKamelV2Month(nextLabel);
  Logger.log('✅ 次月「' + nextLabel + '」を最新月の複製で作成しました。');
}

// ============================================================
//  修復：目標タブの「_月キー」列を再構築する
//   ・列が無ければ右端に作成して非表示
//   ・各店舗行に、上のバナー年月から fill-down したクリーン月ラベルを書き込む
//   ・列を並べ替えた／旧版で作った／キーが空、等の復旧に使う。実行後はサマリも更新
// ============================================================
function repairKamelV2MonthKey() {
  const ss    = getSSKamelV2_();
  const sheet = ss.getSheetByName(KV2.SHEET_TARGET);
  if (!sheet) throw new Error('目標タブがありません。');
  const lastRow = sheet.getLastRow();
  if (lastRow < KV2.DATA_START_ROW) throw new Error('データ行がありません。');

  // 重複した「_月キー」列を掃除：データの多い1本だけ残し、他（空の重複列）は削除
  const allKeyCols = kv2AllMonthKeyCols_(sheet);
  if (allKeyCols.length > 1) {
    const keep = kv2MonthKeyCol_(sheet);
    const drop = allKeyCols.filter(function(c){ return c !== keep; }).sort(function(a, b){ return b - a; }); // 右から削除
    drop.forEach(function(c){ sheet.deleteColumn(c); });
    Logger.log('重複した「' + KV2_MONTHKEY_LABEL + '」列を削除: ' + drop.join(', '));
  }

  // 列インデックスは削除で変わるのでここで取り直す
  const ymCol    = kv2FindCol_(sheet, '年月');
  const brandCol = kv2FindCol_(sheet, 'ブランド');
  const locCol   = kv2FindCol_(sheet, '店舗名');
  if (!ymCol) throw new Error('「年月」列が見つかりません。ヘッダーを確認してください。');

  let keyCol = kv2FindCol_(sheet, KV2_MONTHKEY_LABEL);
  if (!keyCol) {
    keyCol = sheet.getLastColumn() + 1;
    sheet.getRange(KV2.HEADER_ROW, keyCol).setValue(KV2_MONTHKEY_LABEL);
  }

  const numRows = lastRow - KV2.HEADER_ROW;
  const wide    = Math.max(ymCol, brandCol, locCol);
  const data    = sheet.getRange(KV2.DATA_START_ROW, 1, numRows, wide).getValues();

  const out = [];
  let curLabel = '';
  for (let i = 0; i < numRows; i++) {
    const ymCell = String(data[i][ymCol - 1] || '').trim();
    if (ymCell) {
      const ym = kv2MonthToYYYYMM_(ymCell);
      if (ym) curLabel = kv2YYYYMMToLabel_(ym);     // クリーンな "2026年4月"
    }
    const brand = brandCol ? String(data[i][brandCol - 1] || '').trim() : '';
    const loc   = locCol   ? String(data[i][locCol   - 1] || '').trim() : '';
    const isStore = (brand || loc);                  // 店舗行だけにキーを入れる
    out.push([ isStore && curLabel ? curLabel : '' ]);
  }
  sheet.getRange(KV2.DATA_START_ROW, keyCol, out.length, 1).setValues(out);
  sheet.hideColumns(keyCol);

  buildKamelSummaryV2();
  const filled = out.filter(function(r){ return r[0]; }).length;
  Logger.log('✅ 「' + KV2_MONTHKEY_LABEL + '」を再構築（列 ' + keyCol + ' / 店舗行 ' + filled +
             ' 行に月を設定）。サマリも更新しました。');
}

// ============================================================
//  Stage 3：本部収益 実績入力タブのセットアップ
// ============================================================
function setupHQActualV2() {
  const ss = getSSKamelV2_();
  let sheet = ss.getSheetByName(KV2.SHEET_HQ_ACTUAL);
  if (sheet) {
    throw new Error('「' + KV2.SHEET_HQ_ACTUAL + '」タブは既に存在します。');
  }
  sheet = ss.insertSheet(KV2.SHEET_HQ_ACTUAL);

  const colMap    = kv2ColMap_(KV2_HQ_DEFS);
  const totalCols = KV2_HQ_DEFS.length;
  const headers   = KV2_HQ_DEFS.map(function(d){ return d.label; });

  sheet.getRange(1, 1, 1, totalCols).setValues([headers])
    .setBackground('#6b2b4a').setFontColor('#ffffff').setFontWeight('bold').setFontSize(10).setWrap(true);

  // 識別列のプルダウン
  const stores  = kv2GetKamelStores_();
  const persons = kv2PersonList_(stores);
  if (persons.length > 0) {
    const pr = SpreadsheetApp.newDataValidation().requireValueInList(persons, true).setAllowInvalid(true).build();
    sheet.getRange(2, colMap['担当者'], 500, 1).setDataValidation(pr);
  }
  const br = SpreadsheetApp.newDataValidation().requireValueInList(KV2.BRANDS, true).setAllowInvalid(true).build();
  sheet.getRange(2, colMap['ブランド'], 500, 1).setDataValidation(br);

  // 金額列フォーマット
  KV2_HQ_DEFS.forEach(function(d){
    if (d.fmt === 'yen') sheet.getRange(2, colMap[d.label], 500, 1).setNumberFormat('¥#,##0');
    if (d.type === 'manual') sheet.getRange(2, colMap[d.label], 500, 1).setBackground('#fffde7').setFontColor('#1565c0');
  });

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(colMap['年月'], 90);
  sheet.setColumnWidth(colMap['店舗名'], 110);
  Logger.log('✅ 本部収益 実績入力タブ「' + KV2.SHEET_HQ_ACTUAL + '」を作成しました（発生した月・店舗の行だけ手入力してください）。');
  return sheet;
}

// ============================================================
//  Stage 4（パーサ）：V2「目標」タブを読み取る
//   ・ヘッダー名で列を解決（列の並べ替え・追加に強い）
//   ・年月 fill-down（月セルは1ブロックに1回でOK・空欄は直近月を継承）
//   ・ブランド/店舗名が空の行はスキップ（空行・区切り行を許容）
//   ・突合キー = ブランド + 店舗名（スペース無し連結＝実績データの略称）
//   ・返り値: { 略称: { "YYYYMM": { tkey: 値, ... } } }  ← buildData の targets と同形
//   ※ まだ buildData には接続していない。テスト関数で検証してから切替える。
// ============================================================

// "2026年4月" / "2026/4" / "202604" / Date型 → "202604"
function kv2MonthToYYYYMM_(label) {
  if (label === null || label === undefined || label === '') return null;
  // 日付(Date)型セル対応：年・月をそのまま取り出す（時刻の誤読を防ぐ）
  if (typeof label === 'object' && typeof label.getMonth === 'function') {
    return label.getFullYear() + ('0' + (label.getMonth() + 1)).slice(-2);
  }
  const s = String(label).trim();
  // "2026年4月" / "2026/4" / "2026-4" 形式を優先（区切り文字で年と月を分離）
  let m = s.match(/(\d{4})\s*[年\/\-\.]\s*(\d{1,2})/);
  if (m) return m[1] + ('0' + m[2]).slice(-2);
  // "202604"（6桁）形式
  m = s.match(/(\d{4})(\d{2})/);
  if (m) return m[1] + m[2];
  return null;
}

function getKamelTargetsV2_() {
  const ss    = getSSKamelV2_();
  const sheet = ss.getSheetByName(KV2.SHEET_TARGET);
  const out = {};
  if (!sheet) return out;

  const data = sheet.getDataRange().getValues();
  if (data.length < KV2.DATA_START_ROW) return out;

  // 1行目ヘッダー → ラベル→列index(0始まり)
  const header = data[KV2.HEADER_ROW - 1];
  const idx = {};
  header.forEach(function(h, i){ const k = String(h).trim(); if (k) idx[k] = i; });

  // 識別列の存在チェック（ヘッダー読みの前提）
  ['年月', 'ブランド', '店舗名'].forEach(function(req){
    if (idx[req] === undefined) {
      throw new Error('目標タブのヘッダーに「' + req + '」が見つかりません。ヘッダー文字列を確認してください。');
    }
  });

  // tkeyを持つ列だけ収集（ラベル→tkey/fmt/列）。シートに無い列は除外
  const valueCols = KV2_DEFS
    .filter(function(d){ return d.tkey; })
    .map(function(d){ return { tkey: d.tkey, fmt: d.fmt, col: idx[d.label] }; })
    .filter(function(c){ return c.col !== undefined; });

  let curMonth = '';
  for (let r = KV2.DATA_START_ROW - 1; r < data.length; r++) {
    const row = data[r];

    // 年月 fill-down：値があれば更新、無ければ直近を継承（Date型もそのまま保持）
    const mCell = row[idx['年月']];
    if (mCell !== '' && mCell !== null && mCell !== undefined) curMonth = mCell;

    const brand = String(row[idx['ブランド']] || '').trim();
    const loc   = String(row[idx['店舗名']]  || '').trim();
    if (!brand && !loc) continue;          // 空行・区切り行はスキップ
    if (!curMonth) continue;

    const yyyymm = kv2MonthToYYYYMM_(curMonth);
    if (!yyyymm) continue;
    const short = brand + loc;             // 突合キー（略称）

    const rec = {};
    let hasNonZero = false;
    valueCols.forEach(function(c){
      let v = row[c.col];
      if (v === '' || v === null || v === undefined) return;
      v = parseFloat(v);
      if (isNaN(v)) return;
      if (c.fmt === 'pct') v = Math.round(v * 1000) / 10;  // 0-1 → 0-100（旧targets互換）
      else                 v = Math.round(v);
      rec[c.tkey] = v;
      if (v !== 0) hasNonZero = true;
    });

    // 合計売上（total_sales）が未入力の店舗対策：シートの「合計売上」数式が欠けていても、
    // 構成要素（各契約売上＋クーポン＋オプション＋物販）から算出する。※シートの合計売上式と同じ
    if (!rec.total_sales) {
      const sumTotal = (rec.new_contract_sales || 0) + (rec.reappo_contract_sales || 0)
                     + (rec.keizoku_continue_sales || 0) + (rec.coupon_sales || 0)
                     + (rec.option_sales || 0) + (rec.buppan_sales || 0);
      if (sumTotal > 0) { rec.total_sales = sumTotal; hasNonZero = true; }
    }

    if (!hasNonZero) continue;   // 全項目0＝未入力店舗はスキップ（旧パーサ互換）

    if (!out[short]) out[short] = {};
    out[short][yyyymm] = rec;
  }
  return out;
}

// 担当者ごとの月次合計（④用）。返り値: { 担当者: { "YYYYMM": { tkey: 合計, ... } } }
//   ・合計対象は KV2_DEFS のうち集計して意味のある金額/件数系（pctは合計しない）
function getKamelPersonTotalsV2_() {
  const ss    = getSSKamelV2_();
  const sheet = ss.getSheetByName(KV2.SHEET_TARGET);
  const out = {};
  if (!sheet) return out;

  const data = sheet.getDataRange().getValues();
  if (data.length < KV2.DATA_START_ROW) return out;

  const header = data[KV2.HEADER_ROW - 1];
  const idx = {};
  header.forEach(function(h, i){ const k = String(h).trim(); if (k) idx[k] = i; });
  if (idx['年月'] === undefined || idx['担当者'] === undefined) return out;

  // 合計する列（pct以外でtkeyあり）
  const sumCols = KV2_DEFS
    .filter(function(d){ return d.tkey && d.fmt !== 'pct'; })
    .map(function(d){ return { tkey: d.tkey, col: idx[d.label] }; })
    .filter(function(c){ return c.col !== undefined; });

  let curMonth = '';
  for (let r = KV2.DATA_START_ROW - 1; r < data.length; r++) {
    const row = data[r];
    const mCell = row[idx['年月']];
    if (mCell !== '' && mCell !== null && mCell !== undefined) curMonth = mCell;
    const person = String(row[idx['担当者']] || '').trim();
    if (!person || person === '担当なし') continue;
    if (!curMonth) continue;
    const yyyymm = kv2MonthToYYYYMM_(curMonth);
    if (!yyyymm) continue;

    if (!out[person]) out[person] = {};
    if (!out[person][yyyymm]) out[person][yyyymm] = {};
    const bucket = out[person][yyyymm];
    sumCols.forEach(function(c){
      const v = parseFloat(row[c.col]);
      if (!isNaN(v)) bucket[c.tkey] = (bucket[c.tkey] || 0) + v;
    });
  }
  // 端数整理
  Object.keys(out).forEach(function(p){
    Object.keys(out[p]).forEach(function(m){
      Object.keys(out[p][m]).forEach(function(k){ out[p][m][k] = Math.round(out[p][m][k]); });
    });
  });
  return out;
}

// ============================================================
//  テスト：ヘッダー読みが効くか検証（buildDataには影響なし）
//   ・ヘッダー→列マップ、店舗数・月、サンプル店舗の値をログ出力
//   ・列を並べ替えた後にもう一度実行 → 同じ結果になればヘッダー読みが効いている証拠
// ============================================================
function testKamelTargetsV2() {
  const targets = getKamelTargetsV2_();
  const stores  = Object.keys(targets);
  Logger.log('── V2目標パーサ テスト ──');
  Logger.log('店舗数: ' + stores.length);

  if (stores.length === 0) {
    Logger.log('⚠️ 0件。「目標」タブに入力行があるか、ヘッダー文字列が KV2_DEFS と一致するか確認。');
    return;
  }

  // 月の一覧
  const monthSet = {};
  stores.forEach(function(s){ Object.keys(targets[s]).forEach(function(m){ monthSet[m] = true; }); });
  Logger.log('検出した月: ' + Object.keys(monthSet).sort().join(', '));

  // サンプル：最初の3店舗を表示
  stores.slice(0, 3).forEach(function(s){
    Object.keys(targets[s]).sort().forEach(function(m){
      Logger.log('  [' + s + '] ' + m + ' = ' + JSON.stringify(targets[s][m]));
    });
  });

  // 担当者合計サンプル
  const pt = getKamelPersonTotalsV2_();
  const persons = Object.keys(pt);
  Logger.log('── 担当者合計 (' + persons.length + '名) ──');
  persons.slice(0, 3).forEach(function(p){
    Object.keys(pt[p]).sort().forEach(function(m){
      const b = pt[p][m];
      Logger.log('  [' + p + '] ' + m + ' 合計売上=' + (b.total_sales||0) +
                 ' ロイヤリティ=' + (b.royalty||0) + ' 商材=' + (b.shozai_sales||0));
    });
  });
  Logger.log('✅ ここまで出ればヘッダー読みOK。列を並べ替えて再実行し、同じ値が出るか確認してください。');
}

// ============================================================
//  担当者サマリ タブ（④）：月ごとに「担当者×目標」を集計
//   ・各セルは目標タブ参照の SUMIFS 数式 → 目標タブに入力すると即反映（トリガー不要）
//   ・月別ブロックを行グループ化し、過去月（当月より前）は自動で折りたたむ
//   ・突合キー：目標タブの隠し列「_月キー」＋「担当者」
//   ・月を追加（addKamelV2Month）すると構造ごと自動再生成される
// ============================================================
const KV2_SUMMARY_SHEET = '担当者サマリ';

// 目標タブから「月リスト(ym昇順)」「担当者リスト(出現順)」「必要列の列文字」を取得
function getKamelV2MonthsAndPersons_() {
  const ss = getSSKamelV2_();
  const sheet = ss.getSheetByName(KV2.SHEET_TARGET);
  const out = { months: [], persons: [], headerLetters: {}, ok: false };
  if (!sheet) return out;
  const data = sheet.getDataRange().getValues();
  if (data.length < KV2.DATA_START_ROW) return out;

  const header = data[KV2.HEADER_ROW - 1];
  const idx = {};
  header.forEach(function(h, i){ const k = String(h).trim(); if (k) idx[k] = i; });

  // 必要ラベル → 列文字（1つでも欠ければ ok:false）
  const need = ['担当者', KV2_MONTHKEY_LABEL, '合計売上', 'ロイヤリティ_固定含', 'ロイヤリティ_純%',
                '商材売上', '追加加盟売上', '追加マシン購入売上', 'マシン等レンタル売上',
                '研修売上', '採用売上', 'その他'];
  for (let i = 0; i < need.length; i++) {
    if (idx[need[i]] === undefined) return out;
    out.headerLetters[need[i]] = columnLetter(idx[need[i]] + 1);
  }

  const monthSeen = {}, personSeen = {};
  const keyCol = idx[KV2_MONTHKEY_LABEL];
  for (let r = KV2.DATA_START_ROW - 1; r < data.length; r++) {
    const row = data[r];
    const mk = String(row[keyCol] || '').trim();
    if (mk && !monthSeen[mk]) { monthSeen[mk] = true; out.months.push(mk); }
    const person = String(row[idx['担当者']] || '').trim();
    if (person && person !== '担当なし' && !personSeen[person]) {
      const brand = idx['ブランド'] !== undefined ? String(row[idx['ブランド']] || '').trim() : '';
      const loc   = idx['店舗名']   !== undefined ? String(row[idx['店舗名']]   || '').trim() : '';
      if (brand || loc) { personSeen[person] = true; out.persons.push(person); }
    }
  }
  out.months.sort(function(a, b){
    const ya = kv2MonthToYYYYMM_(a) || '', yb = kv2MonthToYYYYMM_(b) || '';
    return ya < yb ? -1 : ya > yb ? 1 : 0;
  });
  out.ok = true;
  return out;
}

// 担当者サマリ：月ごとブロック＋全社行。各値は目標タブ参照のSUMIFS数式（入力で即反映）。
function buildKamelSummaryV2() {
  const ss = getSSKamelV2_();
  // 既存のグループ残りを避けるため作り直す
  const old = ss.getSheetByName(KV2_SUMMARY_SHEET);
  if (old) ss.deleteSheet(old);
  const sheet = ss.insertSheet(KV2_SUMMARY_SHEET);

  const info = getKamelV2MonthsAndPersons_();
  if (!info.ok) {
    sheet.getRange(1, 1).setValue('目標タブが無いか「' + KV2_MONTHKEY_LABEL +
      '」列がありません。setupKamelV2WithMonths で目標タブを作り直してください。');
    Logger.log('⚠️ サマリ生成中止：目標タブ or _月キー列なし。');
    return sheet;
  }

  const L = info.headerLetters;
  const T = "'" + KV2.SHEET_TARGET + "'";
  function colRng(label){ return T + '!' + L[label] + ':' + L[label]; }

  const header = ['個人', '合計', '末端売上目標', 'ROY目標(固定含)', 'ROY目標(固定省)',
                  '商材売上', '追加加盟売上', '追加機械売上', '機械レンタル売上',
                  '研修売上', '採用売上', 'その他'];
  // サマリ列C〜L → 目標ラベル（合計Bは下の sumFormula で別計算）
  const colDefs = [
    '合計売上',            // C 末端売上目標
    'ロイヤリティ_固定含',  // D ROY固定含
    'ロイヤリティ_純%',     // E ROY固定省
    '商材売上',            // F
    '追加加盟売上',         // G
    '追加マシン購入売上',   // H 追加機械売上
    'マシン等レンタル売上', // I 機械レンタル売上
    '研修売上',            // J
    '採用売上',            // K
    'その他'               // L
  ];
  // 合計(B)=D+F+G+H+I+J+K+L（末端C・ROY固定省Eは除外）
  function sumFormula(r){ return '=D'+r+'+F'+r+'+G'+r+'+H'+r+'+I'+r+'+J'+r+'+K'+r+'+L'+r; }
  // SUMIFS（月＝リテラル／担当者＝同行のA。全社行は担当者条件なし）
  function sumifs(label, monthLabel, withPerson, r){
    let f = 'SUMIFS(' + colRng(label) + ',' + colRng(KV2_MONTHKEY_LABEL) + ',"' + monthLabel + '"';
    if (withPerson) f += ',' + colRng('担当者') + ',$A' + r;
    return '=' + f + ')';
  }

  const rows = [];
  const blocks = [];
  info.months.forEach(function(monthLabel){
    const ym = kv2MonthToYYYYMM_(monthLabel) || monthLabel;
    const titleRow0 = rows.length;
    rows.push(['■ ' + monthLabel].concat(new Array(header.length - 1).fill('')));
    const hdrRow0 = rows.length;
    rows.push(header.slice());
    info.persons.forEach(function(p){
      const r = rows.length + 1;
      const row = [p, sumFormula(r)];
      colDefs.forEach(function(lab){ row.push(sumifs(lab, monthLabel, true, r)); });
      rows.push(row);
    });
    const rz = rows.length + 1;
    const zrow = ['【全社】', sumFormula(rz)];
    colDefs.forEach(function(lab){ zrow.push(sumifs(lab, monthLabel, false, rz)); });
    rows.push(zrow);
    const lastRow0 = rows.length - 1;
    rows.push(new Array(header.length).fill(''));  // ブロック間の空行
    blocks.push({ titleRow0: titleRow0, hdrRow0: hdrRow0, lastRow0: lastRow0, ym: ym });
  });

  if (rows.length === 0) rows.push(['（月がありません。目標タブに月を追加してください）']);

  const maxCols = header.length;
  const norm = rows.map(function(r){ const a = r.slice(); while (a.length < maxCols) a.push(''); return a; });
  sheet.getRange(1, 1, norm.length, maxCols).setValues(norm);  // 先頭"="は数式として解釈される

  const now = new Date();
  const curYM = now.getFullYear() + ('0' + (now.getMonth() + 1)).slice(-2);

  blocks.forEach(function(b){
    const titleR = b.titleRow0 + 1;
    const hdrR   = b.hdrRow0 + 1;
    const firstP = hdrR + 1;
    const lastR  = b.lastRow0 + 1;
    // タイトルはmergeしない（列固定との競合回避）。A列にラベル＋行全体を濃色で帯状に
    sheet.getRange(titleR, 1, 1, maxCols)
      .setBackground('#1a1a2e').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11)
      .setVerticalAlignment('middle');
    sheet.setRowHeight(titleR, 24);
    sheet.getRange(hdrR, 1, 1, maxCols)
      .setBackground('#2d3561').setFontColor('#ffffff').setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
    sheet.getRange(firstP, 2, lastR - firstP + 1, maxCols - 1).setNumberFormat('¥#,##0');
    sheet.getRange(firstP, 2, lastR - firstP + 1, 1).setFontWeight('bold');  // 合計列
    sheet.getRange(lastR, 1, 1, maxCols).setBackground('#fce4ec').setFontWeight('bold');
    sheet.getRange(hdrR, 1, lastR - hdrR + 1, maxCols)
      .setBorder(true, true, true, true, true, true, '#cfd2e0', SpreadsheetApp.BorderStyle.SOLID);
    // 明細（ヘッダー〜全社）をグループ化 → 過去月は折りたたむ（タイトルは残る）
    sheet.getRange(hdrR, 1, lastR - hdrR + 1, 1).shiftRowGroupDepth(1);
    if (b.ym < curYM) {
      const g = sheet.getRowGroup(hdrR, 1);
      if (g) g.collapse();
    }
  });

  sheet.setFrozenColumns(1);
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 115);
  for (let i = 3; i <= maxCols; i++) sheet.setColumnWidth(i, 108);

  Logger.log('✅ 「' + KV2_SUMMARY_SHEET + '」を月別SUMIFS数式で更新（月 ' + info.months.length +
             ' / 担当者 ' + info.persons.length + '名）。目標タブ入力で自動反映します。');
  return sheet;
}

// ============================================================
//  スプレッドシート メニュー（手動操作用）
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('加盟店目標V2')
    .addItem('① 目標タブ作成＋初期月', 'setupKamelV2WithMonths')
    .addItem('② 月を追加（空欄）…', 'promptAddKamelV2Month')
    .addItem('③ 前月を複製して月追加…', 'promptCopyKamelV2Month')
    .addItem('④ 次の月を複製作成（ワンクリック）', 'addNextKamelV2Month')
    .addSeparator()
    .addItem('担当者サマリを更新', 'buildKamelSummaryV2')
    .addItem('_月キー列を修復', 'repairKamelV2MonthKey')
    .addItem('_月キー診断（ログ）', 'kv2DebugMonthKey')
    .addItem('パーサ動作テスト（ログ）', 'testKamelTargetsV2')
    .addToUI();
}

// メニューから月ラベルを入力して追加（空欄の新月）
function promptAddKamelV2Month() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('月を追加（空欄）', '追加する月を入力（例: 2026年7月）', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const label = res.getResponseText().trim();
  if (!label) { ui.alert('月ラベルが空です。'); return; }
  try {
    addKamelV2Month(label);
    ui.alert('「' + label + '」を追加しました。');
  } catch (e) {
    ui.alert('追加できませんでした: ' + e.message);
  }
}

// メニューから月ラベルを入力して、前月（最新月）を複製して追加
function promptCopyKamelV2Month() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('前月を複製して月追加', '追加する月を入力（例: 2026年7月）\n※最新月の入力値ごとコピーします', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const label = res.getResponseText().trim();
  if (!label) { ui.alert('月ラベルが空です。'); return; }
  try {
    copyKamelV2Month(label);
    ui.alert('最新月を複製して「' + label + '」を作成しました。値を必要に応じて編集してください。');
  } catch (e) {
    ui.alert('追加できませんでした: ' + e.message);
  }
}

// ============================================================
//  診断：新「目標」スプシの実構造をダンプする（パーサ作成用）
//   ・全シート名＋gid＋サイズを一覧
//   ・「目標」タブ（または gid 631051371 のタブ）の先頭ブロックを
//     行×列のグリッドそのままログ出力
//   ※ Apps Script エディタでこの関数を選んで実行し、ログ全文を貼ってください。
// ============================================================
const NEW_TARGET_SS_ID  = '1Sw3akWjt_qi85vs5NT3y7XRIhMJjm2MFLRmSCUGE19g';
const NEW_TARGET_GID     = 631051371;       // URL で指定された「目標」タブの gid
const NEW_TARGET_DUMP_ROWS = 60;            // ダンプする最大行数
const NEW_TARGET_DUMP_COLS = 24;            // ダンプする最大列数

function dumpNewTargetSheet() {
  // どのアカウントで実行されているか（共有先を判断するため最初に出す）
  try { Logger.log('実行アカウント(effective): ' + Session.getEffectiveUser().getEmail()); } catch (e) {}
  try { Logger.log('実行アカウント(active): '    + Session.getActiveUser().getEmail()); } catch (e) {}
  Logger.log('開こうとしているスプシID: ' + NEW_TARGET_SS_ID);

  let ss;
  try {
    ss = SpreadsheetApp.openById(NEW_TARGET_SS_ID);
  } catch (e) {
    Logger.log('❌ 新スプシを開けませんでした: ' + e.message);
    Logger.log('→ 上の「実行アカウント」のメールアドレスに、新スプシを「閲覧者」以上で共有してから再実行してください。');
    return;
  }

  // 1) 全シート一覧（名前 / gid / 行数 / 列数）
  Logger.log('=== シート一覧（' + ss.getName() + '） ===');
  const sheets = ss.getSheets();
  let gidSheet = null, namedSheet = null;
  sheets.forEach(function(sh){
    const gid = sh.getSheetId();
    Logger.log('  「' + sh.getName() + '」 gid=' + gid +
               ' rows=' + sh.getLastRow() + ' cols=' + sh.getLastColumn());
    if (gid === NEW_TARGET_GID) gidSheet = sh;
    if (sh.getName() === '目標') namedSheet = sh;
  });

  // 2) ダンプ対象：gid 一致を最優先、無ければ名前「目標」
  const target = gidSheet || namedSheet;
  if (!target) {
    Logger.log('⚠️ gid=' + NEW_TARGET_GID + ' も 名前「目標」も見つかりません。上の一覧から対象タブ名を教えてください。');
    return;
  }
  Logger.log('=== ダンプ対象タブ: 「' + target.getName() + '」 gid=' + target.getSheetId() + ' ===');

  const lastRow = Math.min(target.getLastRow(), NEW_TARGET_DUMP_ROWS);
  const lastCol = Math.min(target.getLastColumn(), NEW_TARGET_DUMP_COLS);
  if (lastRow < 1 || lastCol < 1) { Logger.log('（空のタブです）'); return; }

  // 表示値（フォーマット後）でダンプ＝¥や%付きで人が読める形に
  const disp = target.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  for (let r = 0; r < disp.length; r++) {
    // 行末の空セルを落として見やすく
    let row = disp[r].slice();
    while (row.length && String(row[row.length - 1]).trim() === '') row.pop();
    Logger.log('R' + (r + 1) + ': ' + row.map(function(c, i){ return columnLetter(i + 1) + '=' + c; }).join(' | '));
  }
  Logger.log('=== ダンプ終了（' + lastRow + '行 × ' + lastCol + '列まで） ===');
}
