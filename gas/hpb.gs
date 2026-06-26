// ============================================================
//  HPB集客レポート取り込み（PV / CVR / ACR ＋ 同プラン同エリア平均）
//  ・SS_HPB スプシ、タブ「基本情報」に HPBレポート-基本情報CSV を貼る
//    └ CSVに店舗識別が無いので、先頭に「略称」列を足して各行に略称を入れる
//  ・各行 = 1店舗 × 1月。列はヘッダー名で特定（並び順が変わってもOK）
//  使い方：①SS_HPB設定 → ②testHpb()でログ確認 → ③exportHpbReport()でjson出力
// ============================================================

function getHpbSS_() {
  const id = PropertiesService.getScriptProperties().getProperty('SS_HPB');
  if (!id) throw new Error('スクリプトプロパティ SS_HPB が未設定です');
  const m = id.toString().match(/\/d\/([a-zA-Z0-9_-]+)/);
  return SpreadsheetApp.openById(m ? m[1] : id.toString().trim());
}

function hpbNum_(v) {
  if (v == null || v === '') return null;
  const n = parseFloat(v.toString().replace(/[^\d.\-]/g, ''));
  return isNaN(n) ? null : n;
}

// 基本情報タブを読む → [{short, ym, cvr, cvrAvg, acr, acrAvg, yoyaku, ...}]
function readHpbBasic_(ss) {
  const sheet = ss.getSheetByName('基本情報');
  if (!sheet) throw new Error('「基本情報」タブが見つかりません');
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return { rows: [], unlabeled: [] };
  const H = {}; rows[0].forEach(function(h, i){ H[h.toString().trim()] = i; });
  const ci = function(name){ return H[name] !== undefined ? H[name] : -1; };
  const i = {
    short: ci('略称'), ym: ci('対象年月'),
    cvr: ci('CVR(自店)'), cvrA: ci('CVR(同P同A平均)'),
    acr: ci('ACR(自店)'), acrA: ci('ACR(同P同A平均)'),
    y: ci('予約数(自店)'), yA: ci('予約数(同P同A平均)'),
    u: ci('予約売上高(自店)'), uA: ci('予約売上高(同P同A平均)'),
    tp: ci('総PV数(自店)'), tpA: ci('総PV数(同P同A平均)'),
    sp: ci('サロン情報PV数(自店)'), spA: ci('サロン情報PV数(同P同A平均)'),
    cp: ci('クーポンメニューPV数(自店)'), cpA: ci('クーポンメニューPV数(同P同A平均)')
  };
  const g = function(row, idx){ return idx >= 0 ? hpbNum_(row[idx]) : null; };
  const out = []; const unlabeled = [];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const short = i.short >= 0 ? (row[i.short] || '').toString().trim() : '';
    const ym = i.ym >= 0 ? (row[i.ym] || '').toString().trim().replace(/[^\d]/g, '').slice(0, 6) : '';
    if (!ym) continue;
    if (!short) { unlabeled.push(ym); continue; }
    out.push({
      short: short, ym: ym,
      cvr: g(row, i.cvr), cvrAvg: g(row, i.cvrA),
      acr: g(row, i.acr), acrAvg: g(row, i.acrA),
      yoyaku: g(row, i.y), yoyakuAvg: g(row, i.yA),
      uriage: g(row, i.u), uriageAvg: g(row, i.uA),
      totalPv: g(row, i.tp), totalPvAvg: g(row, i.tpA),
      salonPv: g(row, i.sp), salonPvAvg: g(row, i.spA),
      couponPv: g(row, i.cp), couponPvAvg: g(row, i.cpA)
    });
  }
  return { rows: out, unlabeled: unlabeled };
}

// 略称ごと最新月のレポートを構築
function buildHpbReport_() {
  const ss = getHpbSS_();
  const d = readHpbBasic_(ss);
  const byStore = {};
  d.rows.forEach(function(r){ const cur = byStore[r.short]; if (!cur || r.ym > cur.ym) byStore[r.short] = r; });
  const ymc = {}; d.rows.forEach(function(r){ ymc[r.ym] = (ymc[r.ym] || 0) + 1; });
  const month = Object.keys(ymc).sort(function(a, b){ return ymc[b] - ymc[a]; })[0] || '';
  return {
    updated_at: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'),
    month: month, stores: byStore
  };
}

// ★検証：ログで確認（commitしない）
function testHpb() {
  const ss = getHpbSS_();
  const d = readHpbBasic_(ss);
  Logger.log('基本情報 行数: ' + d.rows.length + ' / 略称なし行: ' + d.unlabeled.length);
  d.rows.slice(0, 12).forEach(function(r){
    Logger.log(r.short + ' [' + r.ym + '] CVR' + r.cvr + '%(平均' + r.cvrAvg + ') ACR' + r.acr
      + '%(平均' + r.acrAvg + ') 総PV' + r.totalPv + '(平均' + r.totalPvAvg + ') 予約' + r.yoyaku + '(平均' + r.yoyakuAvg + ')');
  });
}

// 本番：json出力（commitJsonToGitHub_ は reservation.gs にある共通関数）
function exportHpbReport() {
  const data = buildHpbReport_();
  commitJsonToGitHub_('data/hpb_report.json', data,
    '📊 HPB集客レポート更新 ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'));
}
