// ============================================================
//  LIME FIT 目標・実績データ取得スクリプト（メインGAS）
//  【スクリプトプロパティに設定するもの】
//  SS_MASTER : 店舗一覧スプシID
//  SS_KAMEL  : 加盟店目標スプシID
//  SS_DIRECT : 直営店目標スプシID
//  SS_SALON  : サロンボードCSVスプシID
//  SS_BM     : BMCSVスプシID
//  GITHUB_TOKEN / GITHUB_OWNER / GITHUB_REPO
// ============================================================

const EXCLUDE_SHEETS = ['店舗リスト', '雛形', 'HPB_URL', 'HPB_Cache'];

// HPBスクレイピング設定
const HPB_FETCH_DELAY_MS  = 1500;   // 店舗間スリープ（HPB負荷対策）
const HPB_FETCH_TIMEOUT_S = 270;    // 約4.5分でカットオフ（GAS制限対策）
const HPB_CACHE_SHEET     = 'HPB_Cache';
const HPB_CACHE_HEADERS   = [
  '略称','取得日時','キャッチコピー','評価','口コミ件数',
  '写真枚数','最新口コミ日','クーポン枚数','クーポンJSON','エラー'
];

const TARGET_MONTHS = [
  '202501','202502','202503','202504','202505','202506',
  '202507','202508','202509','202510','202511','202512',
  '202601','202602','202603','202604','202605','202606',
  '202607','202608','202609','202610','202611','202612'
];

const MONTH_COL = {
  '202501': 4,  '202502': 5,  '202503': 6,  '202504': 7,
  '202505': 8,  '202506': 9,  '202507': 10, '202508': 11,
  '202509': 12, '202510': 13, '202511': 14, '202512': 15,
  '202601': 4,  '202602': 5,  '202603': 6,  '202604': 7,
  '202605': 8,  '202606': 9,  '202607': 10, '202608': 11,
  '202609': 12, '202610': 13, '202611': 14, '202612': 15
};

const BLOCK      = 12;
const N_BLOCKS   = 40;
const DATA_START = 4; // 0始まりで行4（Excel行5）

// 加盟店新フォーマット定数
const KAMEL_BLOCK      = 16;
const KAMEL_MAX_STORES = 40;
const KAMEL_DATA_START = 5; // 1始まり

// ============================================================
//  スプレッドシート取得ヘルパー
// ============================================================
function getSS(propKey) {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(propKey);
  if (id) return SpreadsheetApp.openById(id);
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ============================================================
//  ロイヤリティパターン別計算
// ============================================================
function calcRoyalty(sales, pattern) {
  const s = Math.round(sales || 0);
  if (!pattern || pattern === '') return Math.round(s * 0.10 + 44000);
  const p = pattern.toString().trim();
  if (/固定.*8[.．]?8万/.test(p)) return 88000;
  if (/固定.*6[.．]?6万/.test(p)) return 66000;
  if (/固定.*5[.．]?5万/.test(p)) return 55000;
  if (/段階/.test(p) || (/70万/.test(p) && /85万/.test(p))) {
    if (s <  700000) return Math.round(s * 0.08);
    if (s <  850000) return Math.round(s * 0.09);
    if (s <  950000) return Math.round(s * 0.10);
    if (s < 1000000) return Math.round(s * 0.12);
    return Math.round(s * 0.15);
  }
  if (/90万/.test(p)) return s < 900000 ? Math.round(s * 0.08) : Math.round(s * 0.10);
  if (/55万/.test(p)) return s < 550000 ? 44000 : Math.round(s * 0.10 + 44000);
  if (/40万/.test(p) && /8[.．]?8万/.test(p)) return s < 400000 ? 88000 : Math.round(s * 0.10 + 44000);
  if (/40万/.test(p) && /4[.．]?4万/.test(p)) return s < 400000 ? 44000 : Math.round(s * 0.10 + 44000);
  if (/15[%％]/.test(p)) return Math.max(Math.round(s * 0.15), 132000);
  if (/10[%％]/.test(p) && /4[.．]?4万/.test(p)) return Math.round(s * 0.10 + 44000);
  if (/10[%％]/.test(p)) return Math.round(s * 0.10);
  if (/8[%％]/.test(p))  return Math.round(s * 0.08);
  if (/5[%％]/.test(p))  return Math.round(s * 0.05);
  return Math.round(s * 0.10 + 44000);
}

// ============================================================
//  担当者タブを動的に取得
// ============================================================
function getPersons() {
  const ss = getSS('SS_KAMEL');
  return ss.getSheets()
    .map(function(s) { return s.getName(); })
    .filter(function(name) {
      return !EXCLUDE_SHEETS.includes(name) && !name.includes('_実績');
    });
}

// ============================================================
//  略称を自動生成
// ============================================================
function generateAbbr(fullName) {
  const m = fullName.match(/(LIME|Belle|Lift)\s*(plus\s*)?([一-龯぀-ゟ゠-ヿー\w]{1,10})/i);
  if (m) {
    const brand = m[1].charAt(0).toUpperCase() + m[1].slice(1).toLowerCase();
    const loc   = m[3].slice(0, 6);
    return brand + loc;
  }
  return fullName.replace(/【[^】]*】/g, '').replace(/\[[^\]]*\]/g, '').trim().slice(0, 10);
}

// ============================================================
//  メインデータ生成
// ============================================================
function buildData(since, targetsOnly) {
  const shortToFull     = {};
  const store_short     = {};
  const store_type      = {};
  const royalty_pattern = {};
  const store_person    = {};
  const targets         = {};
  const store_hpb_url   = {};
  const store_closed_month = {};   // 略称→撤退月(YYYYMM)。この月までは表示、翌月以降は非表示
  const personOrder     = [];   // 担当者の出現順（店舗一覧D列ベース・ダッシュボードのpersons用）
  const personSeen      = {};

  // 店舗マスターをSS_MASTERの店舗一覧タブから読む
  // A=フル名 B=略称 C=直営加盟 D=担当者 E=ロイヤリティパターン F=HPB_URL
  const ssMaster   = getSS('SS_MASTER');
  const listSheet  = ssMaster ? ssMaster.getSheetByName('店舗一覧') : null;
  if (listSheet) {
    const listData = listSheet.getDataRange().getValues();
    // ヘッダーから「撤退月」列を探す（位置が変わってもよいように名前で特定）
    const headerRow = listData[0] || [];
    let tcol = -1;
    headerRow.forEach(function(h, idx){ if (h && h.toString().indexOf('撤退月') >= 0) tcol = idx; });
    for (let i = 1; i < listData.length; i++) {
      const full   = listData[i][0] ? listData[i][0].toString().trim() : '';
      const short  = listData[i][1] ? listData[i][1].toString().trim() : '';
      const type   = listData[i][2] ? listData[i][2].toString().trim() : '';
      const person = listData[i][3] ? listData[i][3].toString().trim() : '';
      const royPat = listData[i][4] ? listData[i][4].toString().trim() : '';
      const hpbUrl = listData[i][5] ? listData[i][5].toString().trim() : '';
      // 撤退月（YYYYMMに正規化。日付/「2026/5」/「202605」いずれもOK）
      if (tcol >= 0 && short) {
        const cv = listData[i][tcol];
        let closed = '';
        if (cv !== '' && cv != null) {
          if (Object.prototype.toString.call(cv) === '[object Date]') {
            closed = Utilities.formatDate(cv, 'Asia/Tokyo', 'yyyyMM');
          } else {
            const p = cv.toString().trim().split(/[^\d]+/).filter(function(x){ return x; });
            if (p.length >= 2) closed = p[0] + ('0' + p[1]).slice(-2);
            else if (p.length === 1) closed = p[0].slice(0, 6);
          }
        }
        if (closed) store_closed_month[short] = closed;
      }
      if (full && short)  shortToFull[short]      = full;
      if (full && short)  store_short[full]        = short;
      if (full && type)   store_type[full]         = type;
      if (full && royPat) royalty_pattern[full]    = royPat;
      if (full && hpbUrl) store_hpb_url[full]      = hpbUrl;
      // 担当者情報もここから読む（personsの並び順も店舗一覧D列の出現順で確定）
      if (full && person && person !== '担当なし') {
        store_person[full] = person;
        if (!personSeen[person]) { personSeen[person] = true; personOrder.push(person); }
      }
    }
  }

  // ── 加盟店目標：新「目標」スプシ（V2縦持ち1タブ）から読む ──
  //   旧 SS_KAMEL の担当者タブ読み取りは廃止。新スプシIDは スクリプトプロパティ
  //   SS_KAMEL_V2 に設定する（未設定だと SS_KAMEL にフォールバックするので必ず設定）。
  //   getKamelTargetsV2_() の返り値 { 略称: { "YYYYMM": { tkey: 値 } } } を targets にマージ。
  //   キー(略称=ブランド+店舗名)・KPI名(total_sales/royalty/new_count/new_contract_count/
  //   new_contract_rate/new_contract_unit_price …)は従来の targets 形と互換。
  const kamelTargetsV2 = getKamelTargetsV2_();
  const unmatchedKamel = [];
  Object.keys(kamelTargetsV2).forEach(function(short) {
    // 略称(ブランド+店舗名)が店舗一覧マスターに存在するものだけ取り込む（旧挙動と同等）
    if (!shortToFull[short]) { unmatchedKamel.push(short); return; }
    if (!targets[short]) targets[short] = {};
    const byMonth = kamelTargetsV2[short];
    Object.keys(byMonth).forEach(function(ym) {
      targets[short][ym] = byMonth[ym];
    });
  });
  Logger.log('加盟店目標(V2): 取込 ' + (Object.keys(kamelTargetsV2).length - unmatchedKamel.length) +
             ' 店舗 / マスター未一致 ' + unmatchedKamel.length +
             (unmatchedKamel.length ? ' → ' + unmatchedKamel.join(', ') : ''));

  // 実績データ
  const ssSalon = getSS('SS_SALON');
  const actual_data = targetsOnly ? {} : getActualData(ssSalon, new Set(Object.keys(shortToFull)), listSheet, royalty_pattern, since);

  if (!targetsOnly) {
    const ssBM    = getSS('SS_BM');
    const bm_data = getActualDataBM(ssBM, royalty_pattern, since);
    Object.keys(bm_data).forEach(function(month) {
      if (!actual_data[month]) actual_data[month] = {};
      Object.keys(bm_data[month]).forEach(function(store) {
        if (!actual_data[month][store]) {
          actual_data[month][store] = bm_data[month][store];
        } else {
          const sb = actual_data[month][store];
          const bm = bm_data[month][store];
          const total_sales = (sb.total_sales||0) + (bm.total_sales||0);
          const royPat = royalty_pattern[store] || '';
          sb.total_sales             = total_sales;
          sb.royalty                 = calcRoyalty(total_sales, royPat);
          sb.new_count               = (sb.new_count||0)               + (bm.new_count||0);
          sb.new_contract_count      = (sb.new_contract_count||0)      + (bm.new_contract_count||0);
          sb.total_treatments        = (sb.total_treatments||0)        + (bm.total_treatments||0);
          sb.contract_sales          = (sb.contract_sales||0)          + (bm.contract_sales||0);
          sb.new_contract_sales      = (sb.new_contract_sales||0)      + (bm.new_contract_sales||0);
          // ▼ 再来系KPI 合算（BMは0だが将来対応のため）
          sb.reappo_count            = (sb.reappo_count||0)            + (bm.reappo_count||0);
          sb.reappo_contract_count   = (sb.reappo_contract_count||0)   + (bm.reappo_contract_count||0);
          sb.reappo_contract_sales   = (sb.reappo_contract_sales||0)   + (bm.reappo_contract_sales||0);
          sb.keizoku_end_count       = (sb.keizoku_end_count||0)       + (bm.keizoku_end_count||0);
          sb.keizoku_continue_count  = (sb.keizoku_continue_count||0)  + (bm.keizoku_continue_count||0);
          sb.keizoku_continue_sales  = (sb.keizoku_continue_sales||0)  + (bm.keizoku_continue_sales||0);
          // ▼ 売上内訳（BMは0だが将来対応のため）
          sb.coupon_sales            = (sb.coupon_sales||0)            + (bm.coupon_sales||0);
          sb.option_sales            = (sb.option_sales||0)            + (bm.option_sales||0);
          sb.buppan_sales            = (sb.buppan_sales||0)            + (bm.buppan_sales||0);
          sb.other_sales             = (sb.other_sales||0)             + (bm.other_sales||0);
          // ▼ 日次売上：日付キーごとに加算
          const dmerged = {};
          if (sb.daily_sales) Object.keys(sb.daily_sales).forEach(function(k){ dmerged[k] = (dmerged[k]||0) + sb.daily_sales[k]; });
          if (bm.daily_sales) Object.keys(bm.daily_sales).forEach(function(k){ dmerged[k] = (dmerged[k]||0) + bm.daily_sales[k]; });
          sb.daily_sales             = dmerged;
          const totalContracts = sb.new_contract_count;
          const totalNew       = sb.new_count;
          sb.new_contract_rate       = totalNew > 0 ? Math.round((totalContracts / totalNew) * 1000) / 10 : 0;
          sb.new_contract_unit_price = totalContracts > 0 ? Math.round(sb.new_contract_sales / totalContracts) : 0;
          // ▼ 再来系 派生KPI 再計算
          const _rc  = sb.reappo_count;
          const _rcc = sb.reappo_contract_count;
          const _ke  = sb.keizoku_end_count;
          const _kc  = sb.keizoku_continue_count;
          const _newLost = Math.max(0, totalNew - totalContracts);
          sb.reappo_contract_rate        = _rc > 0  ? Math.round(_rcc/_rc*1000)/10 : 0;
          sb.reappo_contract_unit_price  = _rcc > 0 ? Math.round(sb.reappo_contract_sales/_rcc) : 0;
          sb.keizoku_continue_unit_price = _kc > 0  ? Math.round(sb.keizoku_continue_sales/_kc) : 0;
          sb.total_contract_count        = totalContracts + _rcc;
          sb.total_contract_rate        = (totalNew + _rc) > 0 ? Math.round(sb.total_contract_count/(totalNew + _rc)*1000)/10 : 0;
          sb.keizoku_rate               = _ke > 0 ? Math.round(_kc/_ke*1000)/10 : 0;
          sb.dropout_count              = Math.max(0, _ke - _kc);
          sb.dropout_rate               = _ke > 0 ? Math.round(sb.dropout_count/_ke*1000)/10 : 0;
          sb.reappo_recovery_rate       = _newLost > 0 ? Math.round(_rcc/_newLost*1000)/10 : 0;
          if (bm.staff_data) {
            if (!sb.staff_data) sb.staff_data = {};
            Object.keys(bm.staff_data).forEach(function(staff) {
              if (!sb.staff_data[staff]) {
                sb.staff_data[staff] = bm.staff_data[staff];
              } else {
                const ss2 = sb.staff_data[staff];
                const bms = bm.staff_data[staff];
                ss2.total_sales        = (ss2.total_sales||0)        + (bms.total_sales||0);
                ss2.new_count          = (ss2.new_count||0)          + (bms.new_count||0);
                ss2.new_contract_count = (ss2.new_contract_count||0) + (bms.new_contract_count||0);
                ss2.new_contract_rate  = ss2.new_count > 0 ? Math.round((ss2.new_contract_count / ss2.new_count) * 1000) / 10 : 0;
              }
            });
          }
        }
      });
    });
  }

  // 直営店目標をマージ（店名のゆらぎ＝末尾「店」を略称に吸収）
  const knownShorts = new Set(Object.values(store_short));
  const direct_targets = getDirectTargets(since);
  Object.keys(direct_targets).forEach(function(storeName) {
    let key = storeName;
    if (!knownShorts.has(key)) {
      const stripped = key.replace(/店$/, '');     // 「直営 大宮店」→「直営 大宮」等
      if (knownShorts.has(stripped)) key = stripped;
    }
    if (!targets[key]) targets[key] = {};
    Object.keys(direct_targets[storeName]).forEach(function(month) {
      targets[key][month] = direct_targets[storeName][month];
    });
  });


  // ── 実績データのキーをフル名→略称に統一（同じ略称は合算）──
  function mergeStoreData(a, b) {
    const numKeys = ['total_sales','new_count','new_contract_count',
      'new_contract_sales','contract_sales','total_treatments','royalty',
      // ▼ 再来系KPI（実数）
      'reappo_count','reappo_contract_count','reappo_contract_sales',
      'keizoku_end_count','keizoku_continue_count','keizoku_continue_sales',
      // ▼ 売上内訳（円グラフ用）
      'coupon_sales','option_sales','buppan_sales','other_sales'];
    const merged = JSON.parse(JSON.stringify(a));
    numKeys.forEach(function(k) { merged[k] = (a[k]||0) + (b[k]||0); });
    // ▼ 日次売上：日付キーごとに加算
    const md = {};
    if (a.daily_sales) Object.keys(a.daily_sales).forEach(function(k){ md[k] = (md[k]||0) + a.daily_sales[k]; });
    if (b.daily_sales) Object.keys(b.daily_sales).forEach(function(k){ md[k] = (md[k]||0) + b.daily_sales[k]; });
    merged.daily_sales = md;
    const nc = merged.new_contract_count;
    const nn = merged.new_count;
    merged.new_contract_rate       = nn > 0 ? Math.round(nc/nn*1000)/10 : 0;
    merged.new_contract_unit_price = nc > 0 ? Math.round(merged.new_contract_sales/nc) : 0;
    // ▼ 再来系 派生KPI 再計算
    const rc   = merged.reappo_count;
    const rcc  = merged.reappo_contract_count;
    const ke   = merged.keizoku_end_count;
    const kc   = merged.keizoku_continue_count;
    const newLost = Math.max(0, nn - nc);
    merged.reappo_contract_rate        = rc > 0  ? Math.round(rcc/rc*1000)/10 : 0;
    merged.reappo_contract_unit_price  = rcc > 0 ? Math.round(merged.reappo_contract_sales/rcc) : 0;
    merged.keizoku_continue_unit_price = kc > 0  ? Math.round(merged.keizoku_continue_sales/kc) : 0;
    merged.total_contract_count        = nc + rcc;
    merged.total_contract_rate        = (nn + rc) > 0 ? Math.round(merged.total_contract_count/(nn + rc)*1000)/10 : 0;
    merged.keizoku_rate               = ke > 0 ? Math.round(kc/ke*1000)/10 : 0;
    merged.dropout_count              = Math.max(0, ke - kc);
    merged.dropout_rate               = ke > 0 ? Math.round(merged.dropout_count/ke*1000)/10 : 0;
    merged.reappo_recovery_rate       = newLost > 0 ? Math.round(rcc/newLost*1000)/10 : 0;
    // onda/peelingも合算
    if (a.onda && b.onda) {
      merged.onda = mergeStoreData(a.onda, b.onda);
    }
    if (a.peeling && b.peeling) {
      merged.peeling = mergeStoreData(a.peeling, b.peeling);
    }
    // last_dateは新しい方を使う
    if (b.last_date && (!a.last_date || b.last_date > a.last_date)) {
      merged.last_date = b.last_date;
    }
    return merged;
  }

  const normalized_actual = {};
  Object.keys(actual_data).forEach(function(month) {
    normalized_actual[month] = {};
    Object.keys(actual_data[month]).forEach(function(fullOrShort) {
      // store_shortで略称に変換、なければそのまま
      const short = store_short[fullOrShort] || fullOrShort;
      if (!normalized_actual[month][short]) {
        normalized_actual[month][short] = actual_data[month][fullOrShort];
      } else {
        // 同じ略称のデータは合算
        normalized_actual[month][short] = mergeStoreData(
          normalized_actual[month][short],
          actual_data[month][fullOrShort]
        );
      }
    });
  });

  // store_type・store_person・royalty_pattern・store_hpb_urlを略称キーのみで整理
  // フル名キーは含めず、略称に変換したものだけを格納
  const st_short = {};
  const sp_short = {};
  const rp_short = {};
  const hu_short = {};
  Object.keys(store_short).forEach(function(full) {
    const short = store_short[full];
    if (!short) return;
    if (store_type[full])      st_short[short] = store_type[full];
    if (store_person[full])    sp_short[short] = store_person[full];
    if (royalty_pattern[full]) rp_short[short] = royalty_pattern[full];
    if (store_hpb_url[full])   hu_short[short] = store_hpb_url[full];
  });
  // 直営店目標スプシのtargetsキー（略称）もstore_typeに登録
  Object.keys(direct_targets).forEach(function(short) {
    st_short[short] = '直営';
  });

  return {
    updated_at:            new Date().toISOString(),
    persons:               personOrder,   // 店舗一覧D列の担当者（旧SS_KAMELのタブ名は参照しない）
    store_person:          sp_short,
    store_type:            st_short,
    store_short:           store_short,
    store_closed_month:    store_closed_month,
    store_royalty_pattern: rp_short,
    store_hpb_url:         hu_short,
    store_hpb_data:        readHpbCache(),
    targets:               targets,
    actual_data:           normalized_actual
  };
}

// ============================================================
//  BM実績データ読み取り
// ============================================================
function getActualDataBM(ss, royalty_pattern, since) {
  const actual_data = {};

  function isValidStaffName(name) {
    if (!name || name.toString().trim() === '') return false;
    const n = name.toString().trim();
    if (n === 'フリー' || n === '担当者不明') return false;
    if (/^[ァ-ヶー]{1,4}$/.test(n)) return false;
    if (/^\d+$/.test(n)) return false;
    return true;
  }

  function getShinkiFromRow(itemName, shinkiRaiCol) {
    if (itemName) {
      const s = itemName.toString();
      if (s.includes('【カテゴリー】新規')) return '新規';
      if (s.includes('【カテゴリー】再来')) return '再来';
    }
    if (shinkiRaiCol) {
      const v = shinkiRaiCol.toString().trim();
      if (v === '新規' || v === '再来') return v;
    }
    return '';
  }

  function isContractBM(itemName) {
    if (!itemName) return false;
    const s = itemName.toString();
    return s.includes('●契約') || s.includes('契約');
  }

  function isOnda(accRows, itemIdx) {
    return accRows.some(function(row) {
      const item = row[itemIdx] ? row[itemIdx].toString() : '';
      return item.includes('オンダ') || item.includes('マイクロ');
    });
  }

  function parseAmtBM(v) {
    if (v === null || v === undefined || v === '') return 0;
    if (typeof v === 'number') return v;
    return parseFloat(v.toString().replace(/[¥円, ]/g, '')) || 0;
  }

  TARGET_MONTHS.forEach(function(month) {
    if (since && month < since) return;
    const sheet = ss.getSheetByName(month + '_実績_BM');
    if (!sheet) return;
    const rows = sheet.getDataRange().getValues();
    if (rows.length < 2) return;

    const headers = rows[0].map(function(h) { return h.toString().trim(); });
    const col = {};
    headers.forEach(function(h, i) { col[h] = i; });

    const iStore  = col['店舗名']    !== undefined ? col['店舗名']    : 0;
    const iAccId  = col['会計ID']    !== undefined ? col['会計ID']    : 4;
    const iItem   = col['項目名']    !== undefined ? col['項目名']    : 8;
    const iAmt    = col['金額']      !== undefined ? col['金額']      : 12;
    const iStaff  = col['施術担当者'] !== undefined ? col['施術担当者'] : 15;
    const iShinki = col['新規再来']  !== undefined ? col['新規再来']  : 21;
    const iDateBM = col['来店日']    !== undefined ? col['来店日']    : col['会計日'] !== undefined ? col['会計日'] : 1;

    const byAccId = {};
    const storeLastDateBM = {};

    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      const accId = row[iAccId] ? row[iAccId].toString().trim() : '';
      if (!accId) continue;
      if (!byAccId[accId]) byAccId[accId] = [];
      byAccId[accId].push(row);

      const storeName = row[iStore] ? row[iStore].toString().trim() : '';
      const dateVal = row[iDateBM];
      if (storeName && dateVal) {
        let d;
        if (dateVal instanceof Date) {
          d = dateVal;
        } else {
          const s = dateVal.toString().trim();
          const m2 = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
          d = m2 ? new Date(parseInt(m2[1]), parseInt(m2[2])-1, parseInt(m2[3])) : new Date(s);
        }
        if (d && !isNaN(d.getTime())) {
          if (!storeLastDateBM[storeName] || d > storeLastDateBM[storeName]) {
            storeLastDateBM[storeName] = d;
          }
        }
      }
    }

    const storeMap = {};

    Object.keys(byAccId).forEach(function(accId) {
      const accRows = byAccId[accId];
      const firstRow = accRows[0];
      const storeName = firstRow[iStore] ? firstRow[iStore].toString().trim() : '';
      if (!storeName) return;

      let shinki = '';
      for (let i = 0; i < accRows.length && !shinki; i++) {
        shinki = getShinkiFromRow(accRows[i][iItem], accRows[i][iShinki]);
      }

      let staffName = '';
      for (let i = 0; i < accRows.length; i++) {
        const amt = parseAmtBM(accRows[i][iAmt]);
        const s = accRows[i][iStaff] ? accRows[i][iStaff].toString().trim() : '';
        if (amt > 0 && isValidStaffName(s)) { staffName = s; break; }
      }
      if (!staffName) {
        for (let i = 0; i < accRows.length; i++) {
          const s = accRows[i][iStaff] ? accRows[i][iStaff].toString().trim() : '';
          if (isValidStaffName(s)) { staffName = s; break; }
        }
      }
      if (!staffName) staffName = '担当者不明';

      let totalAmt = 0;
      let isContract = false;
      let contractAmt = 0;
      accRows.forEach(function(row) {
        const amt = parseAmtBM(row[iAmt]);
        totalAmt += amt;
        if (isContractBM(row[iItem])) { isContract = true; contractAmt += amt; }
      });

      const ondaFlag = isOnda(accRows, iItem);

      if (!storeMap[storeName]) {
        storeMap[storeName] = {
          total_sales: 0, new_count: 0, new_contract_count: 0,
          total_treatments: 0, contract_sales: 0, new_contract_sales: 0,
          onda_total_sales: 0, onda_new_count: 0, onda_new_contract_count: 0,
          onda_new_contract_sales: 0,
          staff_data: {}
        };
      }
      const sm = storeMap[storeName];
      sm.total_sales      += totalAmt;
      sm.total_treatments += 1;
      sm.contract_sales   += contractAmt;
      if (shinki === '新規') {
        sm.new_count += 1;
        if (isContract) { sm.new_contract_count += 1; sm.new_contract_sales += contractAmt; }
      }
      if (ondaFlag) {
        sm.onda_total_sales += totalAmt;
        if (shinki === '新規') {
          sm.onda_new_count += 1;
          if (isContract) { sm.onda_new_contract_count += 1; sm.onda_new_contract_sales += contractAmt; }
        }
      }

      if (!sm.staff_data[staffName]) {
        sm.staff_data[staffName] = { total_sales: 0, new_count: 0, new_contract_count: 0, new_contract_rate: 0 };
      }
      const sd = sm.staff_data[staffName];
      sd.total_sales += totalAmt;
      if (shinki === '新規') {
        sd.new_count += 1;
        if (isContract) sd.new_contract_count += 1;
      }
    });

    if (!actual_data[month]) actual_data[month] = {};
    Object.keys(storeMap).forEach(function(storeName) {
      const sm = storeMap[storeName];
      Object.keys(sm.staff_data).forEach(function(staff) {
        const sd = sm.staff_data[staff];
        sd.new_contract_rate       = sd.new_count > 0 ? Math.round((sd.new_contract_count / sd.new_count) * 1000) / 10 : 0;
        sd.new_contract_unit_price = sd.new_contract_count > 0 ? Math.round(sd.new_contract_sales / sd.new_contract_count) : 0;
      });
      const royPat = royalty_pattern[storeName] || '';
      const lastDateBMObj = storeLastDateBM[storeName];
      const lastDateBMStr = lastDateBMObj ? Utilities.formatDate(lastDateBMObj, 'Asia/Tokyo', 'yyyy-MM-dd') : null;

      const peel_total_sales       = Math.round(sm.total_sales)        - Math.round(sm.onda_total_sales);
      const peel_new_count         = sm.new_count           - sm.onda_new_count;
      const peel_new_contract_count= sm.new_contract_count  - sm.onda_new_contract_count;
      const peel_new_contract_sales= Math.round(sm.new_contract_sales) - Math.round(sm.onda_new_contract_sales);

      actual_data[month][storeName] = {
        total_sales:             Math.round(sm.total_sales),
        royalty:                 calcRoyalty(Math.round(sm.total_sales), royPat),
        new_count:               sm.new_count,
        new_contract_count:      sm.new_contract_count,
        new_contract_unit_price: sm.new_contract_count > 0 ? Math.round(sm.new_contract_sales / sm.new_contract_count) : 0,
        new_contract_rate:       sm.new_count > 0 ? Math.round((sm.new_contract_count / sm.new_count) * 1000) / 10 : 0,
        total_treatments:        sm.total_treatments,
        contract_sales:          Math.round(sm.contract_sales),
        new_contract_sales:      Math.round(sm.new_contract_sales),
        // ▼ 売上内訳（BMは内訳取得未対応、その他に寄せる）
        coupon_sales:            0,
        option_sales:            0,
        buppan_sales:            0,
        other_sales:             Math.round(sm.total_sales) - Math.round(sm.contract_sales),
        // ▼ 日次売上（BMは未取得 — 空オブジェクト）
        daily_sales:             {},
        // ▼ 再来系KPI（BMは項目構造が異なるため一旦0初期化。後でBM側判定追加予定）
        reappo_count:                0,
        reappo_contract_count:       0,
        reappo_contract_sales:       0,
        reappo_contract_rate:        0,
        reappo_contract_unit_price:  0,
        keizoku_end_count:           0,
        keizoku_continue_count:      0,
        keizoku_continue_sales:      0,
        keizoku_continue_unit_price: 0,
        total_contract_count:        sm.new_contract_count,
        total_contract_rate:         sm.new_count > 0 ? Math.round((sm.new_contract_count / sm.new_count) * 1000) / 10 : 0,
        keizoku_rate:                0,
        dropout_count:               0,
        dropout_rate:                0,
        reappo_recovery_rate:        0,
        last_date:               lastDateBMStr,
        onda: {
          total_sales:             Math.round(sm.onda_total_sales),
          new_count:               sm.onda_new_count,
          new_contract_count:      sm.onda_new_contract_count,
          new_contract_sales:      Math.round(sm.onda_new_contract_sales),
          new_contract_unit_price: sm.onda_new_contract_count > 0 ? Math.round(sm.onda_new_contract_sales / sm.onda_new_contract_count) : 0,
          new_contract_rate:       sm.onda_new_count > 0 ? Math.round((sm.onda_new_contract_count / sm.onda_new_count) * 1000) / 10 : 0
        },
        peeling: {
          total_sales:             peel_total_sales,
          new_count:               peel_new_count,
          new_contract_count:      peel_new_contract_count,
          new_contract_sales:      peel_new_contract_sales,
          new_contract_unit_price: peel_new_contract_count > 0 ? Math.round(peel_new_contract_sales / peel_new_contract_count) : 0,
          new_contract_rate:       peel_new_count > 0 ? Math.round((peel_new_contract_count / peel_new_count) * 1000) / 10 : 0
        },
        staff_data: sm.staff_data
      };
    });
  });

  return actual_data;
}

// ============================================================
//  サロンボード実績データ読み取り（会計IDベース・取り消し差し引き）
//  各数値の計算ルール：
//  ・売上         = 会計区分=会計 の金額合計 - 取り消し会計 の金額合計
//  ・総施術数     = 会計区分=会計 のユニーク会計ID数 - 取り消し会計 のユニーク会計ID数
//  ・新規数       = 新規+会計 のユニーク会計ID数 - 新規+取り消し のユニーク会計ID数
//  ・新規契約数   = 新規+会計+カテゴリ=フェイシャル+メニューに「契約」含む ユニーク会計ID数 - 取り消し同条件
//  ・契約売上     = カテゴリ=フェイシャル+メニューに「契約or消化or利用」含む行の金額合計 - 取り消し同条件
//  ・新規契約売上 = 新規+会計+カテゴリ=フェイシャル+メニューに「契約」含む行の金額合計 - 取り消し同条件
// ============================================================
function getActualData(ss, fullNamesSet, listSheet, royalty_pattern, since) {
  const actual_data = {};

  TARGET_MONTHS.forEach(function(month) {
    if (since && month < since) return;
    const sheet = ss.getSheetByName(month + '_実績');
    if (!sheet) return;
    const rows = sheet.getDataRange().getValues();
    if (rows.length < 2) return;

    const headers = rows[0].map(function(h) { return h.toString().trim(); });
    const col = {};
    headers.forEach(function(h, i) { col[h] = i; });

    const g = function(row, colName) {
      return col[colName] !== undefined ? row[col[colName]] : '';
    };
    const gs = function(row, colName) {
      const v = g(row, colName);
      return v ? v.toString().trim() : '';
    };
    const parseAmt = function(v) {
      if (v === null || v === undefined || v === '') return 0;
      if (typeof v === 'number') return v;
      return parseFloat(v.toString().replace(/[¥円, ]/g, '')) || 0;
    };

    // 会計IDごとに行をグループ化
    const byAccId = {};
    const storeLastDate = {};

    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      const accId = gs(row, '会計ID');
      if (!accId) continue;
      if (!byAccId[accId]) byAccId[accId] = [];
      byAccId[accId].push(row);

      // 最終来店日（会計区分=会計のみ）
      const kaikeiKubun = gs(row, '会計区分');
      if (kaikeiKubun === '会計') {
        const storeName = gs(row, 'お店名');
        const dateVal = g(row, '会計日');
        if (storeName && dateVal) {
          let d = null;
          if (dateVal instanceof Date) {
            d = dateVal;
          } else {
            const s = dateVal.toString().trim().replace(/[-\/]/g, '');
            if (/^\d{8}$/.test(s)) {
              d = new Date(parseInt(s.slice(0,4)), parseInt(s.slice(4,6))-1, parseInt(s.slice(6,8)));
            }
          }
          if (d && !isNaN(d.getTime())) {
            if (!storeLastDate[storeName] || d > storeLastDate[storeName]) {
              storeLastDate[storeName] = d;
            }
          }
        }
      }
    }

    // 店舗ごとに集計
    const storeMap = {};

    Object.keys(byAccId).forEach(function(accId) {
      const accRows = byAccId[accId];
      const firstRow = accRows[0];
      const storeName = gs(firstRow, 'お店名');
      if (!storeName) return;

      const kaikeiKubun = gs(firstRow, '会計区分');
      const isCancelled = kaikeiKubun === '取り消し会計';
      const sign = isCancelled ? -1 : 1; // 取り消しは引き算

      // 会計区分=会計 or 取り消し会計 以外はスキップ
      if (kaikeiKubun !== '会計' && kaikeiKubun !== '取り消し会計') return;

      // 新規再来判定（会計ID単位）
      let shinki = '';
      for (let i = 0; i < accRows.length && !shinki; i++) {
        const v = gs(accRows[i], '新規再来');
        if (v === '新規' || v === '再来') shinki = v;
      }

      // スタッフ名（金額>0の行を優先、会計のみ）
      let staffName = '';
      if (!isCancelled) {
        for (let i = 0; i < accRows.length; i++) {
          const amt = parseAmt(g(accRows[i], '金額'));
          const s = gs(accRows[i], 'スタッフ');
          if (amt > 0 && s) { staffName = s; break; }
        }
        if (!staffName) {
          for (let i = 0; i < accRows.length; i++) {
            const s = gs(accRows[i], 'スタッフ');
            if (s) { staffName = s; break; }
          }
        }
      }

      // ▼ 会計ID単位フラグ（正式仕様）
      //   hasFirst : 回数券（初回）  ← 育成フェーズで初契約
      //   hasUse   : 回数券（消化）  ← 既に契約状態
      //   hasEnd   : 回数券（終了）  ← 契約満了
      //   hasAdd   : 回数券（追加）  ← 継続意思あり（追加購入）
      //   hasCont  : 回数券（継続）  ← 継続意思あり（再購入）
      let hasFirst = false;
      let hasUse   = false;
      let hasEnd   = false;
      let hasAdd   = false;
      let hasCont  = false;

      // ▼ 行ごとの集計
      let totalAmt         = 0;     // 会計ID全行の金額合計
      let contractSalesAmt = 0;     // 契約売上（広義: 契約|消化|利用）
      let allContractAmt   = 0;     // 厳密「契約」行の金額（カテゴリ=F/B/O & I「契約」）
      let newContractAmt   = 0;     // 新規契約売上
      let isNewContract    = false; // 新規契約フラグ

      accRows.forEach(function(row) {
        const amt      = parseAmt(g(row, '金額'));
        const category = gs(row, 'カテゴリ');
        const menu     = gs(row, 'メニュー・店販・ 割引・サービス・オプション');
        totalAmt += amt;

        // ▼ 契約売上はカテゴリ=フェイシャル/ボディ/その他 から取得（H=サービスは¥0ラベル行のため除外）
        const isContractCategory = category === 'フェイシャル' || category === 'ボディ' || category === 'その他';

        // 契約売上（広義: 契約・消化・利用・分割・都度）
        //   ・契約時の支払（新規/再アポ/継続契約売上）
        //   ・回数券の消化中の支払（分割払い・都度払い等を含む）
        if (isContractCategory && /契約|消化|利用|分割|都度/.test(menu)) {
          contractSalesAmt += amt;
        }
        // 厳密「契約」行 → 再アポ・継続契約売上、新規契約売上の元金
        if (isContractCategory && /契約/.test(menu)) {
          allContractAmt += amt;
          if (shinki === '新規') {
            newContractAmt += amt;
            isNewContract   = true;
          }
        }

        // ▼ 会計ID単位フラグ（H=サービス行のラベルから抽出）
        if (category === 'サービス') {
          if (/初回|初めて契約|初めて購入/.test(menu)) hasFirst = true;
          if (/消化/.test(menu))                       hasUse   = true;
          if (/終了/.test(menu))                       hasEnd   = true;
          if (/追加/.test(menu))                       hasAdd   = true;
          if (/継続|再購入/.test(menu))                hasCont  = true;
        }
      });

      // ▼ 既存契約状態フラグ：既に回数券契約状態にある顧客
      const hasExistingContract = hasUse || hasEnd || hasAdd || hasCont;

      // ▼ 再アポ判定（正式仕様）
      //   再アポ数    = U=再来 AND NOT hasExistingContract（未契約状態の再来＝育成対象）
      //   再アポ契約数 = 再アポ AND hasFirst（再来時にクーポン施術後 or その場で契約）
      const isReappo         = (shinki === '再来') && !hasExistingContract;
      const isReappoContract = isReappo && hasFirst;
      const reappoContractAmt = isReappoContract ? allContractAmt : 0;

      // ▼ 継続フェーズ判定（正式仕様・会計IDユニーク）
      //   継続対象数 = hasEnd OR hasCont OR hasAdd（再来時の新規/再来を問わない）
      //   継続成功数 = hasCont OR hasAdd（顧客が継続意思を示したか重視 → 継続+追加を成功扱い）
      //   継続契約売上 = 上記成功会計ID × 契約行金額（F/B/O & I=契約）
      const isKeizokuTarget  = hasEnd || hasCont || hasAdd;
      const isKeizokuSuccess = hasCont || hasAdd;
      const keizokuContinueAmt = isKeizokuSuccess ? allContractAmt : 0;

      if (!storeMap[storeName]) {
        storeMap[storeName] = {
          total_sales: 0, new_count: 0, new_contract_count: 0,
          total_treatments: 0, contract_sales: 0, new_contract_sales: 0,
          // ▼ 再来系KPI
          reappo_count: 0,
          reappo_contract_count: 0,
          reappo_contract_sales: 0,
          keizoku_end_count: 0,
          keizoku_continue_count: 0,
          keizoku_continue_sales: 0,
          // ▼ 売上内訳（円グラフ用）
          coupon_sales: 0,
          option_sales: 0,
          buppan_sales: 0,
          // ▼ 日次売上（折れ線グラフ用）
          daily_sales: {},
          staff_data: {}
        };
      }
      const sm = storeMap[storeName];
      // 取り消し会計のM列はCSV上ですでにマイナス値なのでそのまま加算
      sm.total_sales      += totalAmt;
      sm.total_treatments += isCancelled ? -1 : 1; // 件数だけsignを使う
      sm.contract_sales   += contractSalesAmt;

      // ▼ 売上内訳集計（行単位で再走査）
      //   - クーポン売上: H列カテゴリ=「メニュー付クーポン」AND I列メニューに「契約/消化/回数券/利用」を含まない
      //   - オプション売上: H列カテゴリ=「オプション」
      //   - 物販売上: F列区分=「店販」
      //   ※ 取り消し会計は CSV上で金額がマイナスのためそのまま加算でOK
      accRows.forEach(function(row) {
        const amt      = parseAmt(g(row, '金額'));
        const kubun    = gs(row, '区分');
        const category = gs(row, 'カテゴリ');
        const menu     = gs(row, 'メニュー・店販・ 割引・サービス・オプション');
        if (kubun === '店販') {
          sm.buppan_sales += amt;
        } else if (category === 'オプション') {
          sm.option_sales += amt;
        } else if (category === 'メニュー付クーポン' && !/契約|消化|回数券|利用/.test(menu)) {
          sm.coupon_sales += amt;
        }
      });

      // ▼ 日次売上：会計日キーに会計ID単位の金額合計を加算
      const dateVal0 = g(firstRow, '会計日');
      let dayKey = '';
      if (dateVal0) {
        if (dateVal0 instanceof Date) {
          dayKey = Utilities.formatDate(dateVal0, 'Asia/Tokyo', 'yyyy-MM-dd');
        } else {
          const s0 = dateVal0.toString().trim().replace(/[\/]/g, '-');
          const m0 = s0.match(/^(\d{4})-?(\d{1,2})-?(\d{1,2})/);
          if (m0) dayKey = m0[1] + '-' + ('0' + m0[2]).slice(-2) + '-' + ('0' + m0[3]).slice(-2);
        }
      }
      if (dayKey) {
        sm.daily_sales[dayKey] = (sm.daily_sales[dayKey] || 0) + totalAmt;
      }

      if (shinki === '新規') {
        sm.new_count += isCancelled ? -1 : 1;
        if (isNewContract) {
          sm.new_contract_count += isCancelled ? -1 : 1;
          sm.new_contract_sales += newContractAmt;
        }
      }

      // ▼ 再アポフェーズ集計（再来 & 未契約状態を対象）
      if (isReappo) {
        sm.reappo_count += isCancelled ? -1 : 1;
        if (isReappoContract) {
          sm.reappo_contract_count += isCancelled ? -1 : 1;
          sm.reappo_contract_sales += reappoContractAmt;
        }
      }

      // ▼ 継続フェーズ集計（会計IDユニーク・shinki不問）
      //   同一会計ID内で hasEnd / hasCont / hasAdd が複数立っても二重計上しない
      if (isKeizokuTarget)  sm.keizoku_end_count      += isCancelled ? -1 : 1;
      if (isKeizokuSuccess) {
        sm.keizoku_continue_count += isCancelled ? -1 : 1;
        sm.keizoku_continue_sales += keizokuContinueAmt;
      }

      // スタッフ別集計（会計のみ）
      if (staffName && !isCancelled) {
        if (!sm.staff_data[staffName]) {
          sm.staff_data[staffName] = {
            total_sales: 0, new_count: 0, new_contract_count: 0,
            new_contract_rate: 0, new_contract_sales: 0, new_contract_unit_price: 0,
            // ▼ 再来系
            reappo_count: 0, reappo_contract_count: 0, reappo_contract_sales: 0,
            keizoku_end_count: 0, keizoku_continue_count: 0, keizoku_continue_sales: 0
          };
        }
        const sd = sm.staff_data[staffName];
        sd.total_sales += totalAmt;
        if (shinki === '新規') {
          sd.new_count += 1;
          if (isNewContract) {
            sd.new_contract_count += 1;
            sd.new_contract_sales += newContractAmt;
          }
        }
        if (isReappo) {
          sd.reappo_count += 1;
          if (isReappoContract) {
            sd.reappo_contract_count += 1;
            sd.reappo_contract_sales += reappoContractAmt;
          }
        }
        if (isKeizokuTarget)  sd.keizoku_end_count      += 1;
        if (isKeizokuSuccess) {
          sd.keizoku_continue_count += 1;
          sd.keizoku_continue_sales += keizokuContinueAmt;
        }
      }
    });

    if (!actual_data[month]) actual_data[month] = {};
    Object.keys(storeMap).forEach(function(storeName) {
      const sm = storeMap[storeName];
      Object.keys(sm.staff_data).forEach(function(staff) {
        const sd = sm.staff_data[staff];
        sd.new_contract_rate       = sd.new_count > 0 ? Math.round((sd.new_contract_count / sd.new_count) * 1000) / 10 : 0;
        sd.new_contract_unit_price = sd.new_contract_count > 0 ? Math.round(sd.new_contract_sales / sd.new_contract_count) : 0;
        // ▼ 再来系 派生KPI（スタッフ別）
        sd.reappo_contract_rate        = sd.reappo_count > 0 ? Math.round((sd.reappo_contract_count / sd.reappo_count) * 1000) / 10 : 0;
        sd.reappo_contract_unit_price  = sd.reappo_contract_count > 0 ? Math.round(sd.reappo_contract_sales / sd.reappo_contract_count) : 0;
        sd.keizoku_continue_unit_price = sd.keizoku_continue_count > 0 ? Math.round(sd.keizoku_continue_sales / sd.keizoku_continue_count) : 0;
        sd.total_contract_count        = sd.new_contract_count + sd.reappo_contract_count;
        const totalVisits              = sd.new_count + sd.reappo_count;
        sd.total_contract_rate         = totalVisits > 0 ? Math.round((sd.total_contract_count / totalVisits) * 1000) / 10 : 0;
        sd.keizoku_rate                = sd.keizoku_end_count > 0 ? Math.round((sd.keizoku_continue_count / sd.keizoku_end_count) * 1000) / 10 : 0;
        sd.dropout_count               = Math.max(0, sd.keizoku_end_count - sd.keizoku_continue_count);
        sd.dropout_rate                = sd.keizoku_end_count > 0 ? Math.round((sd.dropout_count / sd.keizoku_end_count) * 1000) / 10 : 0;
      });
      const royPat = royalty_pattern[storeName] || '';
      const lastDateObj = storeLastDate[storeName];
      const lastDateStr = lastDateObj ? Utilities.formatDate(lastDateObj, 'Asia/Tokyo', 'yyyy-MM-dd') : null;

      // ▼ 派生KPI算出
      const newCount      = Math.max(0, sm.new_count);
      const newContCount  = Math.max(0, sm.new_contract_count);
      const reappoCount   = Math.max(0, sm.reappo_count);
      const reappoContCount = Math.max(0, sm.reappo_contract_count);
      const keizokuEnd    = Math.max(0, sm.keizoku_end_count);
      const keizokuCont   = Math.max(0, sm.keizoku_continue_count);
      const dropoutCount  = Math.max(0, keizokuEnd - keizokuCont);
      const totalContractCount = newContCount + reappoContCount;
      const totalVisits        = newCount + reappoCount;
      const newLost            = Math.max(0, newCount - newContCount); // 初回逃した数

      // ▼ 売上内訳（円グラフ用） — その他は残差
      const couponSalesR = Math.round(sm.coupon_sales);
      const optionSalesR = Math.round(sm.option_sales);
      const buppanSalesR = Math.round(sm.buppan_sales);
      const contractSalesR = Math.round(sm.contract_sales);
      const totalSalesR  = Math.round(sm.total_sales);
      const otherSalesR  = totalSalesR - contractSalesR - couponSalesR - optionSalesR - buppanSalesR;

      // ▼ 日次売上を丸めて出力
      const daily_sales = {};
      Object.keys(sm.daily_sales).forEach(function(k) {
        daily_sales[k] = Math.round(sm.daily_sales[k]);
      });

      actual_data[month][storeName] = {
        total_sales:             totalSalesR,
        royalty:                 calcRoyalty(totalSalesR, royPat),
        new_count:               newCount,
        new_contract_count:      newContCount,
        new_contract_unit_price: sm.new_contract_count > 0 ? Math.round(sm.new_contract_sales / sm.new_contract_count) : 0,
        new_contract_rate:       sm.new_count > 0 ? Math.round((sm.new_contract_count / sm.new_count) * 1000) / 10 : 0,
        total_treatments:        Math.max(0, sm.total_treatments),
        contract_sales:          contractSalesR,
        new_contract_sales:      Math.round(sm.new_contract_sales),
        // ▼ 売上内訳（円グラフ用）
        coupon_sales:            couponSalesR,
        option_sales:            optionSalesR,
        buppan_sales:            buppanSalesR,
        other_sales:             otherSalesR,
        // ▼ 日次売上（折れ線グラフ用）
        daily_sales:             daily_sales,
        // ▼ 再来系KPI（実数）
        reappo_count:                reappoCount,
        reappo_contract_count:       reappoContCount,
        reappo_contract_sales:       Math.round(sm.reappo_contract_sales),
        reappo_contract_rate:        reappoCount > 0 ? Math.round((reappoContCount / reappoCount) * 1000) / 10 : 0,
        reappo_contract_unit_price:  reappoContCount > 0 ? Math.round(sm.reappo_contract_sales / reappoContCount) : 0,
        keizoku_end_count:           keizokuEnd,
        keizoku_continue_count:      keizokuCont,
        keizoku_continue_sales:      Math.round(sm.keizoku_continue_sales),
        keizoku_continue_unit_price: keizokuCont > 0 ? Math.round(sm.keizoku_continue_sales / keizokuCont) : 0,
        // ▼ 派生KPI
        total_contract_count:        totalContractCount,                                                       // ⑥
        total_contract_rate:         totalVisits > 0 ? Math.round((totalContractCount / totalVisits) * 1000) / 10 : 0, // ⑦
        keizoku_rate:                keizokuEnd > 0 ? Math.round((keizokuCont / keizokuEnd) * 1000) / 10 : 0,  // ⑩
        dropout_count:               dropoutCount,                                                              // ⑪
        dropout_rate:                keizokuEnd > 0 ? Math.round((dropoutCount / keizokuEnd) * 1000) / 10 : 0,  // ⑫
        reappo_recovery_rate:        newLost > 0 ? Math.round((reappoContCount / newLost) * 1000) / 10 : 0,    // 再アポ回収率
        last_date:               lastDateStr,
        staff_data:              sm.staff_data
      };
    });
  });

  return actual_data;
}

// ============================================================
//  直営店目標読み取り（SS_DIRECTの月タブから）
// ============================================================

const DIRECT_ROWS = {
  header:              7,
  total_sales:         8,
  new_contract_sales:  9,
  coupon_sales:        10,
  total_treatments:    11,
  return_count:        12,
  new_count:           13,
  peel_new_count:      14,
  onda_new_count:      15,
  new_contract_count:  16,
  peel_contract_count: 17,
  onda_contract_count: 18,
  contract_rate:       19,
  peel_rate:           20,
  onda_rate:           21,
  contract_unit_price: 22,
  peel_unit_price:     23,
  onda_unit_price:     24,
  keizoku_sales:       25,
  keizoku_unit_price:  26,
  keizoku_count:       27,
  option_sales:        28,
  reappo_sales:        29,
};

function getDirectTargets(since) {
  const ssDirect = getSS('SS_DIRECT');
  const targets = {};
  const r = DIRECT_ROWS;

  TARGET_MONTHS.forEach(function(month) {
    if (since && month < since) return;
    const sheet = ssDirect.getSheetByName(month);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    if (data.length < r.reappo_sales) return;

    const headerRow = data[r.header - 1];
    if (!headerRow) return;

    for (let ci = 2; ci < headerRow.length; ci++) {
      const storeName = headerRow[ci] ? headerRow[ci].toString().trim() : '';
      if (!storeName) continue;

      const v = function(row) {
        const rowData = data[row - 1];
        if (!rowData) return 0;
        const val = rowData[ci];
        if (val === null || val === undefined || val === '') return 0;
        if (typeof val === 'number') return val;
        const s = val.toString().replace(/[%¥,\s]/g, '');
        return parseFloat(s) || 0;
      };

      const peelNewCount       = v(r.peel_new_count);
      const ondaNewCount       = v(r.onda_new_count);
      // 契約率：スプシで30%入力→内部値0.3で保存される場合と30で保存される場合の両対応
      const peelRateRaw        = v(r.peel_rate);
      const ondaRateRaw        = v(r.onda_rate);
      const peelRate           = peelRateRaw <= 1 ? peelRateRaw * 100 : peelRateRaw;
      const ondaRate           = ondaRateRaw <= 1 ? ondaRateRaw * 100 : ondaRateRaw;
      const peelUnitPrice      = v(r.peel_unit_price);
      const ondaUnitPrice      = v(r.onda_unit_price);
      const peelContractCount  = Math.round(peelNewCount * peelRate / 100);
      const ondaContractCount  = Math.round(ondaNewCount * ondaRate / 100);
      const totalNewCount      = peelNewCount + ondaNewCount;
      const totalContractCount = peelContractCount + ondaContractCount;
      const newContractSales   = peelContractCount * peelUnitPrice + ondaContractCount * ondaUnitPrice;
      const couponSales        = peelNewCount * 5500 + ondaNewCount * 15000;
      const keizokuUnitPrice   = v(r.keizoku_unit_price);
      const keizokuCount       = v(r.keizoku_count);
      const keizokuSales       = keizokuCount * keizokuUnitPrice;
      const optionSales        = v(r.option_sales);
      const reappoSales        = v(r.reappo_sales);
      const totalSales         = newContractSales + couponSales + keizokuSales + optionSales + reappoSales;

      const peelNewContractSales = peelContractCount * peelUnitPrice;
      const ondaNewContractSales = ondaContractCount * ondaUnitPrice;
      const peelCouponSales      = peelNewCount * 5500;
      const ondaCouponSales      = ondaNewCount * 15000;

      if (!targets[storeName]) targets[storeName] = {};
      targets[storeName][month] = {
        total_sales:             Math.round(totalSales),
        new_count:               totalNewCount,
        new_contract_count:      totalContractCount,
        new_contract_rate:       totalNewCount > 0 ? Math.round(totalContractCount / totalNewCount * 1000) / 10 : 0,
        new_contract_unit_price: totalContractCount > 0 ? Math.round(newContractSales / totalContractCount) : 0,
        royalty:                 0,
        // ピーリング目標
        peeling: {
          total_sales:             Math.round(peelNewContractSales + peelCouponSales),
          new_count:               peelNewCount,
          new_contract_count:      peelContractCount,
          new_contract_rate:       peelNewCount > 0 ? Math.round(peelContractCount / peelNewCount * 1000) / 10 : 0,
          new_contract_unit_price: peelUnitPrice,
          new_contract_sales:      Math.round(peelNewContractSales)
        },
        // オンダリフト目標
        onda: {
          total_sales:             Math.round(ondaNewContractSales + ondaCouponSales),
          new_count:               ondaNewCount,
          new_contract_count:      ondaContractCount,
          new_contract_rate:       ondaNewCount > 0 ? Math.round(ondaContractCount / ondaNewCount * 1000) / 10 : 0,
          new_contract_unit_price: ondaUnitPrice,
          new_contract_sales:      Math.round(ondaNewContractSales)
        },
        direct: {
          peel_new_count:      peelNewCount,
          onda_new_count:      ondaNewCount,
          peel_contract_count: peelContractCount,
          onda_contract_count: ondaContractCount,
          peel_rate:           Math.round(peelRate * 10) / 10,
          onda_rate:           Math.round(ondaRate * 10) / 10,
          peel_unit_price:     peelUnitPrice,
          onda_unit_price:     ondaUnitPrice,
          new_contract_sales:  Math.round(newContractSales),
          coupon_sales:        Math.round(couponSales),
          keizoku_unit_price:  keizokuUnitPrice,
          keizoku_count:       keizokuCount,
          keizoku_sales:       Math.round(keizokuSales),
          option_sales:        Math.round(optionSales),
          reappo_sales:        Math.round(reappoSales)
        }
      };
    }
  });

  return targets;
}

// ============================================================
//  GitHub Pages用 data.json エクスポート
// ============================================================
function exportToGitHub() {
  const props  = PropertiesService.getScriptProperties();
  const token  = props.getProperty('GITHUB_TOKEN');
  const owner  = props.getProperty('GITHUB_OWNER') || 'kimu2-lime';
  const repo   = props.getProperty('GITHUB_REPO')  || 'dashboard-master';
  const path   = props.getProperty('GITHUB_PATH')  || 'data/data.json';

  if (!token) { Logger.log('❌ GITHUB_TOKEN が未設定'); return; }

  Logger.log('📊 データ生成中...');
  Logger.log('SS_KAMEL_V2(加盟店目標): ' + (props.getProperty('SS_KAMEL_V2') || '⚠️未設定→SS_KAMELにフォールバック'));
  Logger.log('SS_SALON: ' + (props.getProperty('SS_SALON') || 'アクティブシート'));
  Logger.log('SS_BM: '    + (props.getProperty('SS_BM')    || 'アクティブシート'));

  // CSVから新店舗を検出して店舗一覧スプシに追記
  try { syncNewStoresToMaster(); } catch(e) { Logger.log('⚠️ syncNewStoresToMaster: ' + e); }

  // HPB取得はここでは行わない（遅いため別関数 refreshHpbCache / refreshHpbThenExport に分離）。
  // ダッシュボードには HPB_Cache シートの内容（前回スクレイプ済み）が readHpbCache 経由で反映される。

  const result  = buildData(null, false);
  const jsonStr = JSON.stringify(result);
  const encoded = Utilities.base64Encode(jsonStr, Utilities.Charset.UTF_8);

  const apiBase = 'https://api.github.com/repos/' + owner + '/' + repo + '/contents/' + path;
  const headers = {
    'Authorization': 'token ' + token,
    'Accept': 'application/vnd.github.v3+json',
    'User-Agent': 'GAS-LIME-FIT'
  };

  // SHAをリトライ付きで取得（Bandwidth quota対策）
  let sha = null;
  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      const getResp = UrlFetchApp.fetch(apiBase, { headers: headers, muteHttpExceptions: true });
      const code = getResp.getResponseCode();
      if (code === 200) {
        sha = JSON.parse(getResp.getContentText()).sha;
        break;
      } else if (code === 403 || code === 429) {
        Logger.log('SHA取得待機中... (' + (attempt+1) + '回目)');
        Utilities.sleep(5000);
      } else {
        Logger.log('SHA取得失敗 HTTP ' + code);
        break;
      }
    } catch(e) {
      Logger.log('SHA取得エラー: ' + e);
      Utilities.sleep(3000);
    }
  }
  if (!sha) Logger.log('⚠️ SHAなしでpushを試みます（新規ファイル作成）');

  const now  = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  const body = { message: '📊 data更新 ' + now, content: encoded };
  if (sha) body.sha = sha;

  const putResp = UrlFetchApp.fetch(apiBase, {
    method: 'PUT', headers: headers,
    payload: JSON.stringify(body), muteHttpExceptions: true
  });

  const code = putResp.getResponseCode();
  if (code === 200 || code === 201) {
    Logger.log('✅ GitHubへのpush成功！ (' + now + ')');
  } else {
    Logger.log('❌ push失敗 HTTP ' + code + ': ' + putResp.getContentText().slice(0, 200));
  }
}

// HPB（ホットペッパー）をスクレイピングして HPB_Cache を更新し、その後 export する。
// 「ホットペッパーまで含めて最新化したい」とき用。通常の定期更新は exportToGitHub のみでよい。
function refreshHpbThenExport() {
  try { refreshHpbCache(); } catch(e) { Logger.log('⚠️ refreshHpbCache: ' + e); }
  exportToGitHub();
}

// ============================================================
//  店舗一覧スプシへの新店舗自動検出・追記
// ============================================================
function syncNewStoresToMaster() {
  const ssMaster = getSS('SS_MASTER');
  const masterSheet = ssMaster.getSheetByName('店舗一覧');
  if (!masterSheet) { Logger.log('⚠️ 店舗一覧タブが見つかりません'); return; }

  const masterData = masterSheet.getDataRange().getValues();
  const existingNames = new Set();
  for (let i = 1; i < masterData.length; i++) {
    const name = masterData[i][0] ? masterData[i][0].toString().trim() : '';
    if (name) existingNames.add(name);
  }

  const newStores = [];
  const ssSalon = getSS('SS_SALON');
  const ssBM    = getSS('SS_BM');

  [ssSalon, ssBM].forEach(function(csvSS) {
    try {
      csvSS.getSheets().forEach(function(sheet) {
        const name = sheet.getName();
        if (!name.includes('_実績')) return;
        const rows = sheet.getDataRange().getValues();
        if (rows.length < 2) return;
        const headers = rows[0].map(function(h) { return h.toString().trim(); });
        const storeCol = headers.indexOf('お店名') !== -1 ? headers.indexOf('お店名') : headers.indexOf('店舗名');
        if (storeCol === -1) return;
        for (let i = 1; i < rows.length; i++) {
          const storeName = rows[i][storeCol] ? rows[i][storeCol].toString().trim() : '';
          if (storeName && !existingNames.has(storeName) && !newStores.find(function(s) { return s.name === storeName; })) {
            newStores.push({ name: storeName });
            existingNames.add(storeName);
          }
        }
      });
    } catch(e) { Logger.log('CSV読み取りエラー: ' + e); }
  });

  if (newStores.length > 0) {
    const lastRow = masterSheet.getLastRow();
    newStores.forEach(function(store, idx) {
      const row = lastRow + 1 + idx;
      masterSheet.getRange(row, 1).setValue(store.name);
      masterSheet.getRange(row, 2).setValue(generateAbbr(store.name));
    });
    Logger.log('✅ 新店舗追記: ' + newStores.map(function(s) { return s.name; }).join(', '));
  } else {
    Logger.log('✅ 新規店舗なし');
  }
}

// ============================================================
//  直営店目標タブ生成定数・関数
// ============================================================
const DIRECT_LABELS = {
  8:  '合計売上', 9:  '新規契約売上', 10: '新規クーポン売上',
  11: '総施術数', 12: '再来数', 13: '新規数',
  14: '　ピーリング新規数', 15: '　オンダ新規数',
  16: '新規契約数', 17: '　ピーリング新規契約数', 18: '　オンダ新規契約数',
  19: '契約率', 20: '　ピーリング契約率', 21: '　オンダ契約率',
  22: '新規契約単価', 23: '　ピーリング契約単価', 24: '　オンダ契約単価',
  25: '継続/契約売上', 26: '継続契約単価', 27: '継続早期見込み件数',
  28: 'オプション', 29: '再アポ・都度払い',
};
const DIRECT_INPUT_ROWS  = [11, 12, 14, 15, 20, 21, 23, 24, 26, 27, 28, 29];
const DIRECT_CALC_ROWS   = [8, 9, 10, 13, 16, 17, 18, 19, 22, 25];
const DIRECT_PARENT_ROWS = [8, 13, 16, 19, 22];

function columnLetter(col) {
  let letter = '';
  while (col > 0) {
    const mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

function createDirectTargetSheet(yyyymm) {
  // yyyymmを省略した場合は「yyyymm」という名前の雛形タブを作成
  const tabName = yyyymm || 'yyyymm';
  const ss = getSS('SS_DIRECT');
  let sheet = ss.getSheetByName(tabName);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(tabName);

  // 直営店リストをSS_MASTERの店舗一覧タブ（C列=直営）から取得
  const ssMasterD   = getSS('SS_MASTER');
  const listSheetD  = ssMasterD ? ssMasterD.getSheetByName('店舗一覧') : null;
  const directStores = [];
  const seenShorts = new Set();
  if (listSheetD) {
    const data = listSheetD.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const short = data[i][1] ? data[i][1].toString().trim() : '';
      const type  = data[i][2] ? data[i][2].toString().trim() : '';
      if (short && type === '直営' && !seenShorts.has(short)) {
        directStores.push(short);
        seenShorts.add(short);
      }
    }
  }

  const storeCount = directStores.length;
  const totalRows  = 29;
  const totalCols  = 2 + storeCount;

  const y = tabName !== 'yyyymm' ? tabName.slice(0, 4) : '';
  const m = tabName !== 'yyyymm' ? parseInt(tabName.slice(4, 6)) : '';
  const titleText = tabName !== 'yyyymm' ? (m + '月 直営店目標') : '直営店目標（雛形）';
  sheet.getRange(1, 1).setValue(titleText);
  sheet.getRange(2, 1).setValue('（1〜6行目は自由に記入できます：店長名・メモ等）');

  const headerValues = ['項目', '合計'];
  directStores.forEach(function(s) { headerValues.push(s); });
  sheet.getRange(DIRECT_ROWS.header, 1, 1, totalCols).setValues([headerValues]);

  for (let row = 8; row <= totalRows; row++) {
    sheet.getRange(row, 1).setValue(DIRECT_LABELS[row] || '');
  }

  for (let ci = 0; ci < storeCount; ci++) {
    const col = ci + 3;
    const c   = columnLetter(col);
    const r   = DIRECT_ROWS;
    sheet.getRange(r.peel_contract_count, col).setFormula('=ROUND(' + c + r.peel_new_count + '*' + c + r.peel_rate + '/100,0)');
    sheet.getRange(r.onda_contract_count, col).setFormula('=ROUND(' + c + r.onda_new_count + '*' + c + r.onda_rate + '/100,0)');
    sheet.getRange(r.new_count, col).setFormula('=' + c + r.peel_new_count + '+' + c + r.onda_new_count);
    sheet.getRange(r.new_contract_count, col).setFormula('=' + c + r.peel_contract_count + '+' + c + r.onda_contract_count);
    sheet.getRange(r.contract_rate, col).setFormula('=IF(' + c + r.new_count + '=0,"",ROUND(' + c + r.new_contract_count + '/' + c + r.new_count + '*100,1)&"%")');
    sheet.getRange(r.contract_unit_price, col).setFormula('=IF(' + c + r.new_contract_count + '=0,"",ROUND((' + c + r.peel_contract_count + '*' + c + r.peel_unit_price + '+' + c + r.onda_contract_count + '*' + c + r.onda_unit_price + ')/' + c + r.new_contract_count + ',0))');
    sheet.getRange(r.new_contract_sales, col).setFormula('=' + c + r.peel_contract_count + '*' + c + r.peel_unit_price + '+' + c + r.onda_contract_count + '*' + c + r.onda_unit_price);
    sheet.getRange(r.coupon_sales, col).setFormula('=' + c + r.peel_new_count + '*5500+' + c + r.onda_new_count + '*15000');
    sheet.getRange(r.keizoku_sales, col).setFormula('=' + c + r.keizoku_count + '*' + c + r.keizoku_unit_price);
    sheet.getRange(r.total_sales, col).setFormula('=' + c + r.new_contract_sales + '+' + c + r.coupon_sales + '+' + c + r.keizoku_sales + '+' + c + r.option_sales + '+' + c + r.reappo_sales);
  }

  if (storeCount > 0) {
    const firstC = columnLetter(3);
    const lastC  = columnLetter(2 + storeCount);
    const sumRows = [9,10,11,12,13,14,15,16,17,18,25,26,27,28,29];
    sumRows.forEach(function(row) { sheet.getRange(row, 2).setFormula('=SUM(' + firstC + row + ':' + lastC + row + ')'); });
    sheet.getRange(DIRECT_ROWS.total_sales, 2).setFormula('=B' + DIRECT_ROWS.new_contract_sales + '+B' + DIRECT_ROWS.coupon_sales + '+B' + DIRECT_ROWS.keizoku_sales + '+B' + DIRECT_ROWS.option_sales + '+B' + DIRECT_ROWS.reappo_sales);
    sheet.getRange(DIRECT_ROWS.contract_rate, 2).setFormula('=IF(B' + DIRECT_ROWS.new_count + '=0,"",ROUND(B' + DIRECT_ROWS.new_contract_count + '/B' + DIRECT_ROWS.new_count + '*100,1)&"%")');
    sheet.getRange(DIRECT_ROWS.contract_unit_price, 2).setFormula('=IF(B' + DIRECT_ROWS.new_contract_count + '=0,"",ROUND(B' + DIRECT_ROWS.new_contract_sales + '/B' + DIRECT_ROWS.new_contract_count + ',0))');
    sheet.getRange(DIRECT_ROWS.peel_rate, 2).setFormula('=IF(B' + DIRECT_ROWS.peel_new_count + '=0,"",ROUND(B' + DIRECT_ROWS.peel_contract_count + '/B' + DIRECT_ROWS.peel_new_count + '*100,1)&"%")');
    sheet.getRange(DIRECT_ROWS.onda_rate, 2).setFormula('=IF(B' + DIRECT_ROWS.onda_new_count + '=0,"",ROUND(B' + DIRECT_ROWS.onda_contract_count + '/B' + DIRECT_ROWS.onda_new_count + '*100,1)&"%")');
    sheet.getRange(DIRECT_ROWS.peel_unit_price, 2).setFormula('=IF(B' + DIRECT_ROWS.peel_contract_count + '=0,"",ROUND((SUMPRODUCT(' + firstC + DIRECT_ROWS.peel_contract_count + ':' + lastC + DIRECT_ROWS.peel_contract_count + '*' + firstC + DIRECT_ROWS.peel_unit_price + ':' + lastC + DIRECT_ROWS.peel_unit_price + '))/B' + DIRECT_ROWS.peel_contract_count + ',0))');
    sheet.getRange(DIRECT_ROWS.onda_unit_price, 2).setFormula('=IF(B' + DIRECT_ROWS.onda_contract_count + '=0,"",ROUND((SUMPRODUCT(' + firstC + DIRECT_ROWS.onda_contract_count + ':' + lastC + DIRECT_ROWS.onda_contract_count + '*' + firstC + DIRECT_ROWS.onda_unit_price + ':' + lastC + DIRECT_ROWS.onda_unit_price + '))/B' + DIRECT_ROWS.onda_contract_count + ',0))');
    sheet.getRange(DIRECT_ROWS.keizoku_sales, 2).setFormula('=B' + DIRECT_ROWS.keizoku_count + '*B' + DIRECT_ROWS.keizoku_unit_price);
  }

  sheet.getRange(1, 1, totalRows, totalCols).setFontFamily('Arial').setFontSize(10);
  sheet.getRange(1, 1, 6, totalCols).setBackground('#e8f4f8');
  sheet.getRange(1, 1).setFontSize(12).setFontWeight('bold');
  sheet.getRange(DIRECT_ROWS.header, 1, 1, totalCols).setBackground('#1a1a2e').setFontColor('#ffffff').setFontWeight('bold');
  DIRECT_INPUT_ROWS.forEach(function(row) { if (storeCount > 0) sheet.getRange(row, 3, 1, storeCount).setBackground('#fffde7'); });
  DIRECT_CALC_ROWS.forEach(function(row)  { sheet.getRange(row, 2, 1, 1 + (storeCount > 0 ? storeCount : 0)).setBackground('#f5f5f5').setFontStyle('italic'); });
  DIRECT_PARENT_ROWS.forEach(function(row) { sheet.getRange(row, 1, 1, totalCols).setFontWeight('bold'); });
  const yenRows = [8, 9, 10, 22, 23, 24, 25, 26, 28, 29];
  yenRows.forEach(function(row) { if (storeCount > 0) sheet.getRange(row, 2, 1, Math.max(1, storeCount)).setNumberFormat('¥#,##0'); });
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 110);
  for (let ci = 0; ci < storeCount; ci++) sheet.setColumnWidth(ci + 3, 100);
  sheet.setFrozenRows(7);
  sheet.setFrozenColumns(1);
  Logger.log('✅ 直営目標タブ作成完了: ' + tabName + '（' + storeCount + '店舗）');
  return sheet;
}


// ============================================================
//  直営店目標 雛形タブ作成（手動実行）
//  実行すると直営店目標スプシに「yyyymm」という雛形タブが作られる
//  → タブを右クリック→コピー→タブ名を202605などに変更して使う
// ============================================================
function createDirectTarget() { createDirectTargetSheet(); }


// ============================================================
//  加盟店担当者タブ 雛形生成
// ============================================================
const KAMEL_BLOCK_DEF    = 15;  // ロイヤリティパターン行削除で15行
const KAMEL_MONTH_START_COL = 4;
const KAMEL_MONTHS = [
  '2026年4月','2026年5月','2026年6月','2026年7月',
  '2026年8月','2026年9月','2026年10月','2026年11月','2026年12月'
];
const ROYALTY_PATTERNS = [
  '固定8.8万円','固定6.6万円','固定5.5万円',
  '売上5%','売上8%','売上10%','売上10%+4.4万円',
  '売上15%（最低13.2万円）',
  '40万円基準:4.4万/10%+4.4万','40万円基準:8.8万/10%+4.4万',
  '55万円基準:4.4万/10%+4.4万','90万円基準:8%/10%',
  '段階制（70/85/95/100万）'
];
// ロイヤリティパターン行を削除（店舗一覧E列で一元管理）→ 15行ブロック
const KB = {
  store_name:    0,  // 店舗名（A列）＋C列にクーポン単価
  total_sales:   1,  // 合計売上（自動）
  royalty:       2,  // ロイヤリティ（自動）
  new_cont_sales:3,  // 新規契約売上（自動）
  coupon_sales:  4,  // 新規クーポン売上（自動）
  new_count:     5,  // 新規数（手入力）
  new_cont_count:6,  // 新規契約数（自動）
  new_cont_rate: 7,  // 新規契約率（手入力）
  new_cont_price:8,  // 新規契約単価（手入力）
  keizoku_sales: 9,  // 継続契約売上（自動）
  keizoku_price: 10, // 継続契約単価（手入力）
  keizoku_count: 11, // 継続見込み件数（手入力）
  option_sales:  12, // オプション（手入力）
  reappo_sales:  13, // 再アポ/分割/都度払い（手入力）
  blank:         14, // 空白区切り
};
const KAMEL_LABELS_DEF = {
  0:'',
  1:'合計売上', 2:'ロイヤリティ', 3:'新規契約売上', 4:'新規クーポン売上',
  5:'新規数', 6:'新規契約数', 7:'新規契約率', 8:'新規契約単価',
  9:'継続契約売上', 10:'継続契約単価', 11:'継続見込み件数',
  12:'オプション', 13:'再アポ/分割/都度払い', 14:'',
};
const KAMEL_INPUT_OFFSETS = [5, 7, 8, 10, 11, 12, 13];
const KAMEL_CALC_OFFSETS  = [1, 2, 3, 4, 6, 9];

function buildRoyaltyFormula(royPatCell, salesCell) {
  const s = salesCell;
  const p = royPatCell;
  return '=IF(' + s + '=0,0,' +
    'IF(REGEXMATCH(IFERROR(' + p + ',""),"固定.*8.?8万"),88000,' +
    'IF(REGEXMATCH(IFERROR(' + p + ',""),"固定.*6.?6万"),66000,' +
    'IF(REGEXMATCH(IFERROR(' + p + ',""),"固定.*5.?5万"),55000,' +
    'IF(REGEXMATCH(IFERROR(' + p + ',""),"段階|70万"),' +
      'IF(' + s + '<700000,ROUND(' + s + '*0.08,0),' +
      'IF(' + s + '<850000,ROUND(' + s + '*0.09,0),' +
      'IF(' + s + '<950000,ROUND(' + s + '*0.10,0),' +
      'IF(' + s + '<1000000,ROUND(' + s + '*0.12,0),' +
      'ROUND(' + s + '*0.15,0))))),' +
    'IF(REGEXMATCH(IFERROR(' + p + ',""),"90万"),' +
      'IF(' + s + '<900000,ROUND(' + s + '*0.08,0),ROUND(' + s + '*0.10,0)),' +
    'IF(REGEXMATCH(IFERROR(' + p + ',""),"55万"),' +
      'IF(' + s + '<550000,44000,ROUND(' + s + '*0.10+44000,0)),' +
    'IF(REGEXMATCH(IFERROR(' + p + ',""),"40万.*8.?8万"),' +
      'IF(' + s + '<400000,88000,ROUND(' + s + '*0.10+44000,0)),' +
    'IF(REGEXMATCH(IFERROR(' + p + ',""),"40万"),' +
      'IF(' + s + '<400000,44000,ROUND(' + s + '*0.10+44000,0)),' +
    'IF(REGEXMATCH(IFERROR(' + p + ',""),"15%"),' +
      'MAX(ROUND(' + s + '*0.15,0),132000),' +
    'IF(REGEXMATCH(IFERROR(' + p + ',""),"10%.*4.?4万"),' +
      'ROUND(' + s + '*0.10+44000,0),' +
    'IF(REGEXMATCH(IFERROR(' + p + ',""),"10%"),' +
      'ROUND(' + s + '*0.10,0),' +
    'IF(REGEXMATCH(IFERROR(' + p + ',""),"8%"),' +
      'ROUND(' + s + '*0.08,0),' +
    'IF(REGEXMATCH(IFERROR(' + p + ',""),"5%"),' +
      'ROUND(' + s + '*0.05,0),' +
    'ROUND(' + s + '*0.10+44000,0)' +
    '))))))))))))))))';
}

function createKamelTemplate() {
  const ss = getSS('SS_KAMEL');
  let sheet = ss.getSheetByName('雛形');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('雛形');

  // 加盟店略称リストをSS_MASTERの店舗一覧タブから取得
  const ssMasterK  = getSS('SS_MASTER');
  const listSheetK = ssMasterK ? ssMasterK.getSheetByName('店舗一覧') : null;
  const shortNames = [];
  if (listSheetK) {
    const data = listSheetK.getDataRange().getValues();
    const seen = new Set();
    for (let i = 1; i < data.length; i++) {
      const short = data[i][1] ? data[i][1].toString().trim() : '';
      const type  = data[i][2] ? data[i][2].toString().trim() : '';
      if (short && type !== '直営' && !seen.has(short)) { shortNames.push(short); seen.add(short); }
    }
  }

  const totalCols = KAMEL_MONTH_START_COL - 1 + KAMEL_MONTHS.length;
  const totalRows = 4 + KAMEL_BLOCK_DEF * KAMEL_MAX_STORES;

  sheet.getRange(1, 1, 1, totalCols).merge().setValue('担当者名 ／ 月次KPI目標値シート')
    .setBackground('#1a1a2e').setFontColor('#ffffff').setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(2, 1).setValue('前提');
  sheet.getRange(2, 2).setValue('クーポン単価は各店舗ブロックの3行目に入力してください');
  sheet.getRange(2, 1, 1, totalCols).setBackground('#f0f4ff');
  sheet.getRange(3, 1).setValue('【青文字】手入力  ／  【黒文字】自動計算  ／  ロイヤリティはパターン選択で自動計算');
  sheet.getRange(3, 1, 1, totalCols).setBackground('#fff8e1').setFontSize(9).setFontColor('#888');

  sheet.getRange(4, 1).setValue('店舗名');
  sheet.getRange(4, 2).setValue('項目');
  sheet.getRange(4, 3).setValue('ロイヤリティパターン');
  KAMEL_MONTHS.forEach(function(m, i) { sheet.getRange(4, KAMEL_MONTH_START_COL + i).setValue(m); });
  sheet.getRange(4, 1, 1, totalCols).setBackground('#2d3561').setFontColor('#ffffff').setFontWeight('bold').setFontSize(10);

  for (let bi = 0; bi < KAMEL_MAX_STORES; bi++) {
    const baseRow = 5 + bi * KAMEL_BLOCK_DEF;

    const storeCell = sheet.getRange(baseRow, 1);
    if (shortNames.length > 0) {
      const rule = SpreadsheetApp.newDataValidation().requireValueInList(shortNames, true).setAllowInvalid(true).build();
      storeCell.setDataValidation(rule);
    }
    sheet.getRange(baseRow, 1, 1, 3).setBackground('#e8eaf6').setFontWeight('bold');
    storeCell.setFontSize(12).setFontColor('#1a237e');
    sheet.getRange(baseRow, 2).setValue('← 店舗を選択').setFontColor('#9e9e9e').setFontSize(9).setFontStyle('italic');
    sheet.getRange(baseRow, 3).setBackground('#fffde7').setFontColor('#1565c0').setNumberFormat('#,##0').setNote('クーポン単価を入力（例: 6000）');

    const royRow = baseRow + KB.royalty_pat;
    const royRule = SpreadsheetApp.newDataValidation().requireValueInList(ROYALTY_PATTERNS, true).setAllowInvalid(true).build();
    sheet.getRange(royRow, 2).setValue('ロイヤリティパターン').setFontColor('#555').setFontSize(10);
    sheet.getRange(royRow, 3).setDataValidation(royRule).setBackground('#fff3e0').setFontColor('#e65100').setFontSize(10);

    for (let offset = 2; offset < KAMEL_BLOCK_DEF; offset++) {
      const row   = baseRow + offset;
      const label = KAMEL_LABELS_DEF[offset] || '';
      if (label) sheet.getRange(row, 2).setValue(label).setFontColor('#444').setFontSize(10);
    }

    for (let mi = 0; mi < KAMEL_MONTHS.length; mi++) {
      const mc = KAMEL_MONTH_START_COL + mi;
      const cl = columnLetter(mc);
      const couponPrice  = 'C' + (baseRow + KB.store_name);
      const newCount     = cl + (baseRow + KB.new_count);
      const newRate      = cl + (baseRow + KB.new_cont_rate);
      const newPrice     = cl + (baseRow + KB.new_cont_price);
      const newContCount = cl + (baseRow + KB.new_cont_count);
      const keizokuPrice = cl + (baseRow + KB.keizoku_price);
      const keizokuCount = cl + (baseRow + KB.keizoku_count);
      const optionSales  = cl + (baseRow + KB.option_sales);
      const reappoSales  = cl + (baseRow + KB.reappo_sales);
      const newContSales = cl + (baseRow + KB.new_cont_sales);
      const couponSales  = cl + (baseRow + KB.coupon_sales);
      const keizokuSales = cl + (baseRow + KB.keizoku_sales);
      const totalSales   = cl + (baseRow + KB.total_sales);
      const royPat       = 'C' + (baseRow + KB.royalty_pat);

      sheet.getRange(baseRow + KB.new_cont_count, mc).setFormula('=ROUND(' + newCount + '*' + newRate + ',0)');
      sheet.getRange(baseRow + KB.new_cont_sales, mc).setFormula('=' + newContCount + '*' + newPrice);
      sheet.getRange(baseRow + KB.coupon_sales,   mc).setFormula('=' + newCount + '*' + couponPrice);
      sheet.getRange(baseRow + KB.keizoku_sales,  mc).setFormula('=' + keizokuPrice + '*' + keizokuCount);
      sheet.getRange(baseRow + KB.total_sales,    mc).setFormula('=' + newContSales + '+' + couponSales + '+' + keizokuSales + '+' + optionSales + '+' + reappoSales);
      sheet.getRange(baseRow + KB.royalty,        mc).setFormula(buildRoyaltyFormula(royPat, totalSales));
    }

    KAMEL_INPUT_OFFSETS.forEach(function(offset) {
      sheet.getRange(baseRow + offset, KAMEL_MONTH_START_COL, 1, KAMEL_MONTHS.length).setBackground('#fffde7').setFontColor('#1565c0');
    });
    KAMEL_CALC_OFFSETS.forEach(function(offset) {
      sheet.getRange(baseRow + offset, KAMEL_MONTH_START_COL, 1, KAMEL_MONTHS.length).setBackground('#f5f5f5').setFontStyle('italic');
    });
    [KB.total_sales, KB.royalty, KB.new_cont_sales, KB.coupon_sales, KB.keizoku_sales, KB.keizoku_price, KB.new_cont_price, KB.option_sales, KB.reappo_sales].forEach(function(offset) {
      sheet.getRange(baseRow + offset, KAMEL_MONTH_START_COL, 1, KAMEL_MONTHS.length).setNumberFormat('¥#,##0');
    });
    sheet.getRange(baseRow + KB.new_cont_rate, KAMEL_MONTH_START_COL, 1, KAMEL_MONTHS.length).setNumberFormat('0.0%');
    sheet.getRange(baseRow + KB.total_sales, 1, 1, totalCols).setBackground('#e8f5e9').setFontWeight('bold');
    sheet.getRange(baseRow + KB.royalty, 1, 1, totalCols).setBackground('#fff3e0');
    sheet.getRange(baseRow + KB.blank, 1, 1, totalCols).setBackground('#eeeeee');
  }

  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 170);
  for (let i = 0; i < KAMEL_MONTHS.length; i++) sheet.setColumnWidth(KAMEL_MONTH_START_COL + i, 95);
  sheet.setFrozenRows(4);
  sheet.setFrozenColumns(2);
  Logger.log('✅ 加盟店担当者タブ雛形作成完了');
  return sheet;
}

function createKamelTemplateRun() { createKamelTemplate(); Logger.log('完了'); }


// デバッグ用：直営店リストの読み取り結果を確認
function debugDirectStores() {
  const ssMasterD  = getSS('SS_MASTER');
  const listSheetD = ssMasterD ? ssMasterD.getSheetByName('店舗一覧') : null;
  if (!listSheetD) { Logger.log('❌ 店舗一覧タブが見つかりません'); return; }

  const data = listSheetD.getDataRange().getValues();
  const directRows = [];
  const seenShorts = new Set();
  const duplicates = [];

  for (let i = 1; i < data.length; i++) {
    const full  = data[i][0] ? data[i][0].toString().trim() : '';
    const short = data[i][1] ? data[i][1].toString().trim() : '';
    const type  = data[i][2] ? data[i][2].toString().trim() : '';
    if (type !== '直営' || !short) continue;

    if (seenShorts.has(short)) {
      duplicates.push('重複: ' + short + ' ← ' + full);
    } else {
      seenShorts.add(short);
      directRows.push(short);
    }
  }

  Logger.log('直営店（重複除外後）: ' + directRows.length + '店舗');
  Logger.log(directRows.join(', '));
  if (duplicates.length > 0) {
    Logger.log('重複略称（スキップ）: ' + duplicates.length + '件');
    duplicates.forEach(function(d) { Logger.log(d); });
  } else {
    Logger.log('✅ 重複なし');
  }
}

// サロンボードCSVのヘッダー確認用デバッグ関数
function debugSalonHeaders() {
  const ss = getSS('SS_SALON');
  const sheets = ss.getSheets().filter(function(s) { return s.getName().includes('_実績') && !s.getName().includes('_BM'); });
  if (sheets.length === 0) { Logger.log('❌ 実績タブが見つかりません'); return; }

  const sheet = sheets[0];
  Logger.log('タブ名: ' + sheet.getName());
  const headers = sheet.getRange(1, 1, 1, 30).getValues()[0];
  headers.forEach(function(h, i) {
    if (h) Logger.log((i+1) + '列目: [' + h + ']');
  });

  // 2行目のデータも確認
  const row2 = sheet.getRange(2, 1, 1, 30).getValues()[0];
  Logger.log('2行目のデータ:');
  row2.forEach(function(v, i) {
    if (v !== '') Logger.log('  ' + (i+1) + '列目: ' + v);
  });
}

// ============================================================
//  店舗一覧D列の担当者割り当てを加盟店目標スプシに一括反映
// ============================================================
function syncAllPersonAssignments() {
  const ssMaster  = getSS('SS_MASTER');
  const ssKamel   = getSS('SS_KAMEL');
  const masterSheet = ssMaster.getSheetByName('店舗一覧');
  if (!masterSheet) { Logger.log('❌ 店舗一覧タブが見つかりません'); return; }

  const data = masterSheet.getDataRange().getValues();
  const personToShorts = {};
  for (let i = 1; i < data.length; i++) {
    const short  = data[i][1] ? data[i][1].toString().trim() : '';
    const type   = data[i][2] ? data[i][2].toString().trim() : '';
    const person = data[i][3] ? data[i][3].toString().trim() : '';
    if (!short || !person || person === '担当なし' || type === '直営') continue;
    if (!personToShorts[person]) personToShorts[person] = [];
    if (!personToShorts[person].includes(short)) personToShorts[person].push(short);
  }

  const persons = Object.keys(personToShorts);
  Logger.log('対象担当者: ' + persons.length + '名');

  persons.forEach(function(person) {
    const sheet = ssKamel.getSheetByName(person);
    if (!sheet) { Logger.log('⚠️ タブなし: ' + person); return; }
    const sheetData = sheet.getDataRange().getValues();
    const isNewFormat = sheetData[3] && sheetData[3][1] && sheetData[3][1].toString().includes('項目');
    const blockSize   = isNewFormat ? KAMEL_BLOCK_DEF : KAMEL_BLOCK;
    const dataStart   = isNewFormat ? KAMEL_DATA_START - 1 : DATA_START;
    const existingShorts = new Set();
    for (let i = 0; i < KAMEL_MAX_STORES; i++) {
      const baseRow = dataStart + i * blockSize;
      if (baseRow >= sheetData.length) break;
      const val = sheetData[baseRow][0] ? sheetData[baseRow][0].toString().trim() : '';
      if (val) existingShorts.add(val);
    }
    let addCount = 0;
    personToShorts[person].forEach(function(short) {
      if (existingShorts.has(short)) return;
      for (let i = 0; i < KAMEL_MAX_STORES; i++) {
        const baseRowIdx = dataStart + i * blockSize;
        const baseRow1   = baseRowIdx + 1;
        if (baseRowIdx >= sheetData.length) { sheet.getRange(baseRow1, 1).setValue(short); addCount++; existingShorts.add(short); return; }
        const val = sheetData[baseRowIdx][0] ? sheetData[baseRowIdx][0].toString().trim() : '';
        if (!val) { sheet.getRange(baseRow1, 1).setValue(short); addCount++; existingShorts.add(short); sheetData[baseRowIdx][0] = short; return; }
      }
      Logger.log('⚠️ 空きブロックなし: ' + person + ' / ' + short);
    });
    Logger.log('✅ ' + person + ': ' + addCount + '店舗を追加');
  });
  Logger.log('✅ 一括同期完了');
}

// ============================================================
//  全担当者タブのロイヤリティ数式を店舗一覧E列パターンで一括更新
// ============================================================
function updateRoyaltyFormulasInPersonTabs() {
  const ssMaster  = getSS('SS_MASTER');
  const ssKamel   = getSS('SS_KAMEL');
  const masterSheet = ssMaster.getSheetByName('店舗一覧');
  if (!masterSheet) { Logger.log('❌ 店舗一覧タブが見つかりません'); return; }
  const masterData = masterSheet.getDataRange().getValues();
  const royaltyMap = {};
  for (let i = 1; i < masterData.length; i++) {
    const short  = masterData[i][1] ? masterData[i][1].toString().trim() : '';
    const royPat = masterData[i][4] ? masterData[i][4].toString().trim() : '';
    if (short && royPat) royaltyMap[short] = royPat;
  }
  Logger.log('ロイヤリティパターン登録数: ' + Object.keys(royaltyMap).length);

  const persons = getPersons();
  persons.forEach(function(person) {
    const sheet = ssKamel.getSheetByName(person);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    const isNewFormat = data[3] && data[3][1] && data[3][1].toString().includes('項目');
    const blockSize   = isNewFormat ? KAMEL_BLOCK_DEF : KAMEL_BLOCK;
    const dataStart   = isNewFormat ? KAMEL_DATA_START - 1 : DATA_START;
    const royaltyOffset = isNewFormat ? KB.royalty : 3;
    let updateCount = 0;
    for (let i = 0; i < KAMEL_MAX_STORES; i++) {
      const baseRowIdx = dataStart + i * blockSize;
      if (baseRowIdx >= data.length) break;
      const short = data[baseRowIdx][0] ? data[baseRowIdx][0].toString().trim() : '';
      if (!short) continue;
      const royPat = royaltyMap[short];
      if (!royPat) { Logger.log('  パターン未設定: ' + short); continue; }
      for (let mi = 0; mi < KAMEL_MONTHS.length; mi++) {
        const mc = KAMEL_MONTH_START_COL + mi;
        const cl = columnLetter(mc);
        const totalSalesCell = cl + (baseRowIdx + 1 + KB.total_sales);
        const royaltyRow     = baseRowIdx + 1 + royaltyOffset;
        sheet.getRange(royaltyRow, mc).setFormula(buildRoyaltyFormulaFixed(royPat, totalSalesCell));
      }
      updateCount++;
    }
    Logger.log('✅ ' + person + ': ' + updateCount + '店舗のロイヤリティ数式を更新');
  });
  Logger.log('✅ ロイヤリティ数式一括更新完了');
}

// ロイヤリティ計算式：パターン文字列を直接埋め込む版
function buildRoyaltyFormulaFixed(royPat, salesCell) {
  const s = salesCell;
  if (!royPat) return '=0';
  if (/固定.*8[.・]?8万/.test(royPat)) return '=IF(' + s + '=0,0,88000)';
  if (/固定.*6[.・]?6万/.test(royPat)) return '=IF(' + s + '=0,0,66000)';
  if (/固定.*5[.・]?5万/.test(royPat)) return '=IF(' + s + '=0,0,55000)';
  if (/段階/.test(royPat)) return '=IF(' + s + '=0,0,IF(' + s + '<700000,ROUND(' + s + '*0.08,0),IF(' + s + '<850000,ROUND(' + s + '*0.09,0),IF(' + s + '<950000,ROUND(' + s + '*0.10,0),IF(' + s + '<1000000,ROUND(' + s + '*0.12,0),ROUND(' + s + '*0.15,0))))))';
  if (/90万/.test(royPat)) return '=IF(' + s + '=0,0,IF(' + s + '<900000,ROUND(' + s + '*0.08,0),ROUND(' + s + '*0.10,0)))';
  if (/55万/.test(royPat)) return '=IF(' + s + '=0,0,IF(' + s + '<550000,44000,ROUND(' + s + '*0.10+44000,0)))';
  if (/40万.*8[.・]?8万/.test(royPat)) return '=IF(' + s + '=0,0,IF(' + s + '<400000,88000,ROUND(' + s + '*0.10+44000,0)))';
  if (/40万/.test(royPat)) return '=IF(' + s + '=0,0,IF(' + s + '<400000,44000,ROUND(' + s + '*0.10+44000,0)))';
  if (/15[%％]/.test(royPat)) return '=IF(' + s + '=0,0,MAX(ROUND(' + s + '*0.15,0),132000))';
  if (/10[%％].*4[.・]?4万/.test(royPat)) return '=IF(' + s + '=0,0,ROUND(' + s + '*0.10+44000,0))';
  if (/10[%％]/.test(royPat)) return '=IF(' + s + '=0,0,ROUND(' + s + '*0.10,0))';
  if (/8[%％]/.test(royPat)) return '=IF(' + s + '=0,0,ROUND(' + s + '*0.08,0))';
  if (/5[%％]/.test(royPat)) return '=IF(' + s + '=0,0,ROUND(' + s + '*0.05,0))';
  return '=IF(' + s + '=0,0,ROUND(' + s + '*0.10+44000,0))';
}

// ============================================================
//  加盟店目標の読み取り状況をデバッグ
// ============================================================
function debugKamelTargets() {
  const ss = getSS('SS_KAMEL');
  const persons = getPersons();
  Logger.log('担当者タブ数: ' + persons.length + ' → ' + persons.join(', '));
  persons.forEach(function(person) {
    const sheet = ss.getSheetByName(person);
    if (!sheet) { Logger.log('タブなし: ' + person); return; }
    const data = sheet.getDataRange().getValues();
    if (data.length < 5) { Logger.log(person + ': データが少なすぎ(' + data.length + '行)'); return; }
    const isNewFormat = data[3] && data[3][1] && data[3][1].toString().includes('項目');
    const blockSize   = isNewFormat ? KAMEL_BLOCK_DEF : KAMEL_BLOCK;
    const dataStart   = isNewFormat ? KAMEL_DATA_START - 1 : DATA_START;
    Logger.log(person + ': ' + (isNewFormat ? '新' : '旧') + 'フォーマット blockSize=' + blockSize + ' dataStart=' + dataStart);
    for (let i = 0; i < 3; i++) {
      const baseRow = dataStart + i * blockSize;
      if (baseRow >= data.length) break;
      const storeName = data[baseRow][0] ? data[baseRow][0].toString().trim() : '';
      if (!storeName) continue;
      const col = isNewFormat ? KAMEL_MONTH_START_COL - 1 : 6; // 4月の列（0始まり）
      const totalSales = data[baseRow + KB.total_sales] ? data[baseRow + KB.total_sales][col] : '?';
      Logger.log('  ' + storeName + ': total_sales=' + totalSales);
    }
  });
}

// ============================================================
//  HPB（ホットペッパービューティー）スクレイピング
//  店舗一覧F列 HPB_URL を起点にサロンTOP + クーポンページを取得し、
//  SS_MASTER の HPB_Cache タブにキャッシュする
// ============================================================
function ensureHpbCacheSheet() {
  const ss = getSS('SS_MASTER');
  let sheet = ss.getSheetByName(HPB_CACHE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(HPB_CACHE_SHEET);
    sheet.getRange(1, 1, 1, HPB_CACHE_HEADERS.length).setValues([HPB_CACHE_HEADERS])
      .setBackground('#1a1a2e').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, HPB_CACHE_HEADERS.length, 120);
    sheet.setColumnWidth(3, 280);  // キャッチコピー
    sheet.setColumnWidth(9, 400);  // クーポンJSON
    Logger.log('✅ HPB_Cache シートを作成');
  }
  return sheet;
}

function readHpbCache() {
  const ss = getSS('SS_MASTER');
  const sheet = ss.getSheetByName(HPB_CACHE_SHEET);
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const result = {};
  for (let i = 1; i < data.length; i++) {
    const short = data[i][0] ? data[i][0].toString().trim() : '';
    if (!short) continue;
    let coupons = [];
    try { coupons = JSON.parse(data[i][8] || '[]'); } catch(e) {}
    result[short] = {
      fetched_at:       data[i][1] ? new Date(data[i][1]).toISOString() : null,
      catch_copy:       data[i][2] || '',
      rating:           parseFloat(data[i][3]) || null,
      review_count:     parseInt(data[i][4])   || 0,
      photo_count:      parseInt(data[i][5])   || 0,
      last_review_date: data[i][6] || null,
      coupon_count:     parseInt(data[i][7])   || 0,
      coupons:          coupons,
      error:            data[i][9] || null
    };
  }
  return result;
}

function refreshHpbCache() {
  const ssMaster = getSS('SS_MASTER');
  const listSheet = ssMaster.getSheetByName('店舗一覧');
  if (!listSheet) { Logger.log('⚠️ 店舗一覧タブなし → HPB取得スキップ'); return; }

  const cacheSheet = ensureHpbCacheSheet();
  const listData   = listSheet.getDataRange().getValues();
  const cacheData  = cacheSheet.getDataRange().getValues();
  const rowByShort = {};
  for (let i = 1; i < cacheData.length; i++) {
    const short = cacheData[i][0] ? cacheData[i][0].toString().trim() : '';
    if (short) rowByShort[short] = i + 1;
  }

  const start = Date.now();
  let updated = 0, skipped = 0, errors = 0;

  for (let i = 1; i < listData.length; i++) {
    if ((Date.now() - start) / 1000 > HPB_FETCH_TIMEOUT_S) {
      Logger.log('⏱ タイムアウト前にカットオフ（次回継続）');
      break;
    }
    const short  = listData[i][1] ? listData[i][1].toString().trim() : '';
    const hpbUrl = listData[i][5] ? listData[i][5].toString().trim() : '';
    if (!short || !hpbUrl) { skipped++; continue; }

    const parsed = scrapeHpbStore(hpbUrl);
    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
    const rowValues = [
      short, now,
      parsed.catch_copy || '',
      parsed.rating || '',
      parsed.review_count || 0,
      parsed.photo_count || 0,
      parsed.last_review_date || '',
      parsed.coupon_count || 0,
      JSON.stringify(parsed.coupons || []),
      parsed.error || ''
    ];
    const targetRow = rowByShort[short] || (cacheSheet.getLastRow() + 1);
    cacheSheet.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);
    rowByShort[short] = targetRow;

    if (parsed.error) errors++; else updated++;
    Utilities.sleep(HPB_FETCH_DELAY_MS);
  }
  Logger.log('✅ HPB取得完了: 更新=' + updated + ' / スキップ=' + skipped + ' / エラー=' + errors);
}

function scrapeHpbStore(salonUrl) {
  const result = { catch_copy:'', rating:null, review_count:0, photo_count:0, last_review_date:'', coupon_count:0, coupons:[], error:null };
  try {
    const topResp = UrlFetchApp.fetch(salonUrl, { muteHttpExceptions: true, followRedirects: true });
    if (topResp.getResponseCode() !== 200) {
      result.error = 'TOP HTTP ' + topResp.getResponseCode();
      return result;
    }
    parseHpbSalonTop(topResp.getContentText(), result);

    const couponUrl = salonUrl.replace(/\/$/, '') + '/coupon/';
    const couponResp = UrlFetchApp.fetch(couponUrl, { muteHttpExceptions: true, followRedirects: true });
    if (couponResp.getResponseCode() === 200) {
      parseHpbCouponPage(couponResp.getContentText(), result);
    }
  } catch(e) {
    result.error = String(e).slice(0, 200);
  }
  return result;
}

function parseHpbSalonTop(html, result) {
  let m;
  m = html.match(/評価[\s\S]{0,80}?([0-9]\.[0-9])/);
  if (m) result.rating = parseFloat(m[1]);

  m = html.match(/口コミ[\s\S]{0,100}?(\d{1,5})\s*件/);
  if (m) result.review_count = parseInt(m[1]);

  m = html.match(/<meta[^>]+name=["']description["'][^>]+content=["']([^"']+)["']/i);
  if (m) result.catch_copy = decodeHtmlEntities(m[1]).trim().slice(0, 200);

  m = html.match(/フォトギャラリー[\s\S]{0,80}?[(（]\s*(\d{1,4})\s*[)）]/);
  if (m) result.photo_count = parseInt(m[1]);

  m = html.match(/口コミ[\s\S]{0,300}?(20\d{2})[\/年-](\d{1,2})[\/月-](\d{1,2})/);
  if (m) {
    const y = m[1], mo = ('0'+m[2]).slice(-2), d = ('0'+m[3]).slice(-2);
    result.last_review_date = y + '-' + mo + '-' + d;
  }
}

function parseHpbCouponPage(html, result) {
  const titles = [];
  const titleRegex = /class=["'][^"']*couponNm[^"']*["'][^>]*>([\s\S]*?)</gi;
  let tm;
  while ((tm = titleRegex.exec(html)) !== null) {
    titles.push(decodeHtmlEntities(stripTags(tm[1])).trim());
  }
  const prices = [];
  const priceRegex = /¥\s*([\d,]+)/g;
  let pm;
  while ((pm = priceRegex.exec(html)) !== null) {
    prices.push(parseInt(pm[1].replace(/,/g,'')));
  }
  const coupons = [];
  const n = Math.min(titles.length, prices.length);
  for (let i = 0; i < n; i++) {
    coupons.push({ title: titles[i].slice(0, 120), price: prices[i] });
  }
  result.coupons = coupons.slice(0, 30);
  result.coupon_count = coupons.length;
}

function stripTags(s) { return String(s||'').replace(/<[^>]*>/g, ''); }
function decodeHtmlEntities(s) {
  return String(s||'')
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&nbsp;/g, ' ');
}

// 単体テスト用：1店舗だけスクレイプして結果をログ出力
// 'https://beauty.hotpepper.jp/slnH000XXXXXX/' を実URLに書き換えて実行
function debugHpbScrape() {
  const url = 'https://beauty.hotpepper.jp/slnH000XXXXXX/';
  const result = scrapeHpbStore(url);
  Logger.log(JSON.stringify(result, null, 2));
}
