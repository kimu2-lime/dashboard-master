// ============================================================
//  Limeからの仕入れ（BASE発注管理　担当別売上スプシ）
//
//  依頼35（2026-10-30）：お店がLimeから買った商材（仕入れ）を画面に出す。
//
//  ★このスプシは加藤さんのものです。読むだけ。絶対に書き込まないこと。
//    （キム兄さんの指示。openById → getValues しかしていません）
//
//  読むタブ：03_明細
//    使う列：年月・店舗名・正式商品名・合計金額・状態
//    状態が「OK」以外の行は数えない（除外件数はログに出す）。
//    ※配送先の宛名・住所・郵便番号は個人情報なので読みません。
//
//  店名の突き合わせ：
//    商材スプシの店名は「Belle梅田」のようにブランド＋地名で、
//    店舗一覧の略称と一致することが多いが、会計単位でまとめている店
//    （梅田・浦和・直営 駒沢大学駅前）は一致しない。
//    店舗一覧スプシの「商材店名対応」タブ（A:商材スプシの店名 B:略称 C:メモ）
//    を先に見て、無ければ略称そのままで突き合わせる。
//    それでも合わない店名は shozai_name_issues に入れて画面上部に出す。
//
//  業務用／物販用：
//    商品名では判定できない（例：LMハーブsoft は施術で使う材料）。
//    店舗一覧スプシの「商材区分」タブ（A:正式商品名 B:業務用／物販用 C:メモ）
//    から読む。空欄の商品は「未分類」として出す。
//    「値引き」はクーポン行のことで、仕入れ合計には入れるが区分には数えない。
//    （2026-10-30 キム兄さん確認：物販用は6商品だけ、残りは全部業務用）
// ============================================================

const SS_SHOZAI_FALLBACK_ID = '1lfTiDvr6mCSjxQKviSPzMaSrwrn2ZrQvZt6_06TGeB4';

function getShozaiSS_() {
  const id = PropertiesService.getScriptProperties().getProperty('SS_SHOZAI');
  if (!id) {
    Logger.log('⚠️ スクリプトプロパティ SS_SHOZAI が未設定です。'
             + '既定のID（' + SS_SHOZAI_FALLBACK_ID + '）で開きます。');
    return SpreadsheetApp.openById(SS_SHOZAI_FALLBACK_ID);
  }
  return SpreadsheetApp.openById(id);
}

// 商材スプシの店名 → 店舗一覧の略称
function readShozaiNameMap_() {
  const out = {};
  try {
    const sh = getSS('SS_MASTER').getSheetByName('商材店名対応');
    if (!sh) { Logger.log('商材店名対応タブなし（略称そのままで突き合わせます）'); return out; }
    const v = sh.getDataRange().getValues();
    for (let i = 1; i < v.length; i++) {
      const a = v[i][0] ? v[i][0].toString().trim() : '';
      const b = v[i][1] ? v[i][1].toString().trim() : '';
      if (a && b) out[a] = b;
    }
    addNormAliases_(out);
    Logger.log('商材店名対応: ' + Object.keys(out).length + '件');
  } catch (e) {
    Logger.log('⚠️ 商材店名対応タブを読めません: ' + e);
  }
  return out;
}

// 正式商品名 → 区分（業務用／物販用／値引き）。空欄は入れない（＝未分類）
function readShozaiKubun_() {
  const out = {};
  try {
    const sh = getSS('SS_MASTER').getSheetByName('商材区分');
    if (!sh) { Logger.log('商材区分タブなし（すべて未分類で出します）'); return out; }
    const v = sh.getDataRange().getValues();
    for (let i = 1; i < v.length; i++) {
      const a = v[i][0] ? v[i][0].toString().trim() : '';
      const b = v[i][1] ? v[i][1].toString().trim() : '';
      if (a && b) out[a] = b;
    }
    addNormAliases_(out);
    Logger.log('商材区分: ' + Object.keys(out).length + '件');
  } catch (e) {
    Logger.log('⚠️ 商材区分タブを読めません: ' + e);
  }
  return out;
}

// 仕入れを集計する。
//   knownShorts … 店舗一覧の略称の集合（Set）
//   戻り値 { purchase, name_issues, kubun }
//     purchase[YYYYMM][略称] = { total, by_kind:{...}, items:{商品名: 金額} }
function buildShozai_(since, knownShorts) {
  const purchase = {};
  const name_issues = [];
  const kubun = readShozaiKubun_();
  const nameMap = readShozaiNameMap_();

  let sheet = null;
  try {
    const ss = getShozaiSS_();
    sheet = ss ? ss.getSheetByName('03_明細') : null;
  } catch (e) {
    Logger.log('⚠️ 商材スプシを開けません: ' + e);
    return { purchase: purchase, name_issues: name_issues, kubun: kubun };
  }
  if (!sheet) {
    Logger.log('⚠️ 商材スプシに「03_明細」タブがありません');
    return { purchase: purchase, name_issues: name_issues, kubun: kubun };
  }

  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return { purchase: purchase, name_issues: name_issues, kubun: kubun };

  const col = {};
  rows[0].forEach(function (h, i) { col[h ? h.toString().trim() : ''] = i; });
  const need = ['年月', '店舗名', '正式商品名', '合計金額'];
  for (let i = 0; i < need.length; i++) {
    if (col[need[i]] === undefined) {
      Logger.log('⚠️ 商材スプシ 03_明細 に「' + need[i] + '」列がありません → 仕入れをスキップ');
      return { purchase: purchase, name_issues: name_issues, kubun: kubun };
    }
  }

  const gs_ = function (row, name) {
    const i = col[name];
    return (i !== undefined && row[i] != null) ? row[i].toString().trim() : '';
  };
  const amt_ = function (row, name) {
    const i = col[name];
    if (i === undefined) return 0;
    const v = row[i];
    if (v === '' || v === null || v === undefined) return 0;
    if (typeof v === 'number') return v;
    return parseFloat(v.toString().replace(/[¥円, ]/g, '')) || 0;
  };

  const unmatched = {};     // 店名 → {rows, amount}
  let skipped = 0;          // 状態がOK以外
  let used = 0;

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const raw = gs_(row, '店舗名');
    if (!raw) continue;

    // 状態が「OK」以外は数えない（列が無ければ全部OK扱い）
    if (col['状態'] !== undefined) {
      const st = gs_(row, '状態');
      if (st && st !== 'OK') { skipped++; continue; }
    }

    // 年月：「2026/10」「202610」どちらでも拾う
    const ym = gs_(row, '年月').replace(/[^\d]/g, '').slice(0, 6);
    if (ym.length !== 6) continue;
    if (since && ym < since) continue;

    // 店名 → 略称。対応表が先、無ければ略称そのまま
    let short = pickByName_(nameMap, raw);
    if (!short) {
      const n = normStoreName_(raw);
      if (knownShorts && (knownShorts.has(raw) || knownShorts.has(n))) short = raw;
    }
    if (!short) {
      if (!unmatched[raw]) unmatched[raw] = { rows: 0, amount: 0 };
      unmatched[raw].rows++;
      unmatched[raw].amount += amt_(row, '合計金額');
      continue;
    }

    const amount = amt_(row, '合計金額');
    const item = gs_(row, '正式商品名') || '（商品名なし）';
    const kind = pickByName_(kubun, item) || '未分類';

    if (!purchase[ym]) purchase[ym] = {};
    if (!purchase[ym][short]) purchase[ym][short] = { total: 0, by_kind: {}, items: {} };
    const p = purchase[ym][short];
    p.total += amount;
    p.by_kind[kind] = (p.by_kind[kind] || 0) + amount;
    p.items[item] = (p.items[item] || 0) + amount;
    used++;
  }

  // 丸め
  Object.keys(purchase).forEach(function (ym) {
    Object.keys(purchase[ym]).forEach(function (st) {
      const p = purchase[ym][st];
      p.total = Math.round(p.total);
      Object.keys(p.by_kind).forEach(function (k) { p.by_kind[k] = Math.round(p.by_kind[k]); });
      Object.keys(p.items).forEach(function (k) { p.items[k] = Math.round(p.items[k]); });
    });
  });

  Object.keys(unmatched).forEach(function (nm) {
    name_issues.push({ name: nm, rows: unmatched[nm].rows, amount: Math.round(unmatched[nm].amount) });
  });
  name_issues.sort(function (a, b) { return b.amount - a.amount; });

  Logger.log('仕入れ: ' + used + '行を集計 / 状態OK以外 ' + skipped + '行を除外 / '
           + '店名未一致 ' + name_issues.length + '件'
           + (name_issues.length ? '（' + name_issues.map(function (x) { return x.name; }).join(', ') + '）' : ''));
  return { purchase: purchase, name_issues: name_issues, kubun: kubun };
}

// 仕入れのアラート。
//   ① 今月まだ仕入れていない加盟店（施術があるのに仕入れ0）… 施術の多い順
//   ② 施術はあるのに直近の仕入れが少ない加盟店
//      … 1施術あたりの仕入れ額が、加盟店の中位値の40%未満
//   直営はBASEで買っていないので対象外（キム兄さんの指示）。
//   撤退した店も対象外。
const SHOZAI_LOW_RATIO = 0.4;
const SHOZAI_LOW_MIN_TREATMENTS = 20;

function buildShozaiAlerts_(purchase, actual, store_type, store_closed_month) {
  const out = { month: '', no_purchase: [], low: [], median_per_treatment: 0 };
  const months = Object.keys(actual || {}).sort();
  if (!months.length) return out;
  const month = months[months.length - 1];
  out.month = month;

  const am = actual[month] || {};
  const pm = (purchase || {})[month] || {};

  const isKamel = function (st) {
    const t = store_type ? store_type[st] : '';
    if (!t || t.toString().indexOf('加盟') < 0) return false;
    if (store_closed_month && store_closed_month[st] && store_closed_month[st] <= month) return false;
    return true;
  };

  const rates = [];
  Object.keys(am).forEach(function (st) {
    if (!isKamel(st)) return;
    const tr = am[st].total_treatments || 0;
    const p = pm[st] ? pm[st].total : 0;
    if (tr > 0 && !p) {
      out.no_purchase.push({ store: st, treatments: tr });
    }
    if (tr >= SHOZAI_LOW_MIN_TREATMENTS && p > 0) {
      rates.push({ store: st, treatments: tr, purchase: p, per_treatment: Math.round(p / tr) });
    }
  });

  out.no_purchase.sort(function (a, b) { return b.treatments - a.treatments; });

  if (rates.length) {
    const sorted = rates.map(function (r) { return r.per_treatment; }).sort(function (a, b) { return a - b; });
    const mid = Math.floor(sorted.length / 2);
    const median = sorted.length % 2 ? sorted[mid] : Math.round((sorted[mid - 1] + sorted[mid]) / 2);
    out.median_per_treatment = median;
    const line = median * SHOZAI_LOW_RATIO;
    out.low = rates.filter(function (r) { return r.per_treatment < line; })
                   .sort(function (a, b) { return a.per_treatment - b.per_treatment; });
  }

  Logger.log('仕入れアラート(' + month + '): 未仕入れ ' + out.no_purchase.length + '店 / '
           + '少ない ' + out.low.length + '店（1施術あたり中位値 ' + out.median_per_treatment + '円）');
  return out;
}
