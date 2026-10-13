// ============================================================
//  HPB集客レポート取り込み（PV / CVR / ACR ＋ 同プラン同エリア平均）
//  ・自動DLスクリプトが SS_HPB スプシの「貼り付け」タブに
//    [ファイル名, サロンID, 対象年月, ...] 形式で追記していく
//  ・サロンID(H000...)→略称 は店舗一覧マスターのHPB_URL(F列)から自動解決
//  ・ヘッダー行が途中に何度混ざってもOK（対象年月が数値の行だけ拾う）
//  使い方：①SS_HPB設定 → ②testHpb()でログ確認 → ③exportHpbReport()でjson出力
// ============================================================

const HPB_TAB_CANDIDATES = ['貼り付け', '基本情報'];   // この順で最初に見つかったタブを使う

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

// サロンID(H000...) → 略称 のマップを店舗一覧マスター(B列=略称, F列=HPB_URL)から作る
function buildSalonIdToShort_() {
  const map = {};
  try {
    const sheet = getSS('SS_MASTER').getSheetByName('店舗一覧');
    if (!sheet) return map;
    const rows = sheet.getDataRange().getValues();
    for (let r = 1; r < rows.length; r++) {
      const short = (rows[r][1] || '').toString().trim();      // B列 略称
      const url   = (rows[r][5] || '').toString();             // F列 HPB_URL
      const m = url.match(/sln(H\d+)/i);
      if (short && m) map[m[1].toUpperCase()] = short;
    }
  } catch (e) {}
  return map;
}

// 貼り付けタブを読む → [{short, ym, cvr, cvrAvg, acr, ...}]
function readHpbBasic_(ss) {
  let sheet = null;
  for (let k = 0; k < HPB_TAB_CANDIDATES.length; k++) {
    const s = ss.getSheetByName(HPB_TAB_CANDIDATES[k]);
    if (s) { sheet = s; break; }
  }
  if (!sheet) throw new Error('「貼り付け」または「基本情報」タブが見つかりません');
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return { rows: [], unlabeled: [] };

  // ヘッダー行を探す（「対象年月」列を持つ最初の行）。途中に混ざっていてもOK
  let H = null;
  for (let r = 0; r < rows.length; r++) {
    const map = {}; let hit = false;
    rows[r].forEach(function(h, i){ const k = (h == null ? '' : h.toString().trim()); map[k] = i; if (k === '対象年月') hit = true; });
    if (hit) { H = map; break; }
  }
  if (!H) throw new Error('ヘッダー（「対象年月」列）が見つかりません。貼り付け内容を確認してください');

  const ci = function(name){ return H[name] !== undefined ? H[name] : -1; };
  const iSalon = ci('サロンID');
  const idMap = iSalon >= 0 ? buildSalonIdToShort_() : {};
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
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    const ym = i.ym >= 0 ? (row[i.ym] || '').toString().trim().replace(/[^\d]/g, '').slice(0, 6) : '';
    if (ym.length < 6) continue;   // 対象年月が数値YYYYMMの行＝データ行のみ拾う（ヘッダー等はskip）

    let short = i.short >= 0 ? (row[i.short] || '').toString().trim() : '';
    if (!short && iSalon >= 0) {
      const sid = (row[iSalon] || '').toString().trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
      short = idMap[sid] || '';
    }
    if (!short) { unlabeled.push(ym + (iSalon >= 0 ? ' /ID:' + (row[iSalon] || '') : '')); continue; }

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
  Logger.log('HPB 読み取りデータ行数: ' + d.rows.length + ' / 略称マップできなかった行: ' + d.unlabeled.length);
  if (d.unlabeled.length) {
    Logger.log('▼略称マップ不可（マスターのHPB_URLに無いサロンID）: ' + d.unlabeled.slice(0, 15).join(' , '));
  }
  // 略称ごと最新月を表示
  const rep = buildHpbReport_();
  Logger.log('対象月(最頻): ' + rep.month + ' / 店舗数: ' + Object.keys(rep.stores).length);
  Object.keys(rep.stores).slice(0, 15).forEach(function(s){
    const x = rep.stores[s];
    Logger.log(s + ' [' + x.ym + '] CVR' + x.cvr + '%(平均' + x.cvrAvg + ') ACR' + x.acr
      + '%(平均' + x.acrAvg + ') 総PV' + x.totalPv + '(平均' + x.totalPvAvg + ') 予約' + x.yoyaku + '(平均' + x.yoyakuAvg + ')');
  });
}

// 本番：json出力（commitJsonToGitHub_ は reservation.gs にある共通関数）
function exportHpbReport() {
  const data = buildHpbReport_();
  commitJsonToGitHub_('data/hpb_report.json', data,
    '📊 HPB集客レポート更新 ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'));
}
