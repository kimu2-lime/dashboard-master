// ============================================================
//  ダッシュボードの画面から入れたデータの保存先（依頼45の段階4）
//  ------------------------------------------------------------
//  ■ どこに保存するか
//    新しいスプレッドシート「ダッシュボード入力データ」を1つ作って、タブに分ける。
//      対策メモ       … 店ごとの状態・ロープレ予定・サロボ修正・メモ（③で使う）
//      週メモ         … 週ごとのメモ（③で使う）
//      店舗収支設定   … 家賃・初期投資など（④で使う）★お金の情報
//      月次手入力     … 月ごとの広告費・スタッフ数の上書き（④で使う）
//      アクション状態 … 未着手／実行中／完了（③④で使う）
//      書き込みログ   … 誰が・いつ・何を書いたか
//
//  ■ 共有は最小限
//    Webアプリは「自分（デプロイした人）として実行」なので、
//    画面を見る人はこのスプシに権限が無くても書き込める。
//    だから **このスプシは作った人だけに共有したままでよい**。
//    家賃・初期投資はオーナーのお金の情報なので、共有を広げないこと。
//
//  ■ 安全のための決まり
//    ・書き込みは下の決まった関数だけ。「好きなセルに書く」関数は作らない
//    ・値はGAS側で必ず検証する（画面から来た値をそのまま書かない）
//    ・店名は店舗一覧にある略称だけ受け付ける
//    ・誰が・いつ書いたかを毎回残す
//
//  ■ 最初に1回だけ
//    GASエディタで udSetup() を実行すると、スプシを作ってURLをログに出す。
//    作らなくても、最初の保存のときに自動で作られる。
// ============================================================

const UD_PROP_KEY   = 'SS_USERDATA';
const UD_FILE_NAME  = 'ダッシュボード入力データ';

// タブの名前 → 列の並び（1行目の見出し）。キー列は先頭から数えた「キーの数」で決める
const UD_TABS = {
  memo: {
    sheet: '対策メモ',
    keys:  ['月', '店舗'],
    cols:  ['月', '店舗', '状態', 'ロープレ予定日', 'ロープレ済', 'サロボ修正予定日',
            'サロボ修正済', 'サロボ修正内容', 'メモ', '更新者', '更新日時']
  },
  weekmemo: {
    sheet: '週メモ',
    keys:  ['月', '週', '店舗'],
    cols:  ['月', '週', '店舗', 'メモ', '更新者', '更新日時']
  },
  finance: {
    sheet: '店舗収支設定',
    keys:  ['店舗'],
    cols:  ['店舗', 'オープン年月', '家賃', '初期投資', '上位回数券比率', '広告費の決め方',
            '広告費', '2人目を採用する', '採用開始月', '回収期限', '更新者', '更新日時']
  },
  monthly: {
    sheet: '月次手入力',
    keys:  ['店舗', '月'],
    cols:  ['店舗', '月', '広告費', 'スタッフ数', '更新者', '更新日時']
  },
  action: {
    sheet: 'アクション状態',
    keys:  ['店舗', 'アクション'],
    cols:  ['店舗', 'アクション', '状態', '更新者', '更新日時']
  }
};
const UD_LOG_SHEET = '書き込みログ';
const UD_LOG_COLS  = ['日時', '更新者', '種類', 'キー', '中身'];

// 値の決まり。ここに無い項目は書き込めない
const UD_STATUS   = ['未着手', '予定あり', '実施中', '完了'];
const UD_ACTION   = ['未着手', '実行中', '完了'];
const UD_AD_MODE  = ['固定額', '売上比'];
const UD_MAX_TEXT = 2000;     // メモの最大文字数
const UD_MAX_ROWS = 20000;    // 1タブの上限（暴走したときの歯止め）
// 店に紐づかない「その週全体のメモ」を入れるときの、店舗欄の決まった値
const UD_TEAM_STORE = '（全体）';

// ------------------------------------------------------------
//  スプシを開く／作る
// ------------------------------------------------------------
function udGetSS_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(UD_PROP_KEY);
  if (id) {
    try { return SpreadsheetApp.openById(id); }
    catch (e) { Logger.log('⚠️ ' + UD_PROP_KEY + ' のスプシを開けません: ' + e); }
  }
  // 無ければ作る。作った人だけに共有された状態になる（共有を広げないこと）
  const ss = SpreadsheetApp.create(UD_FILE_NAME);
  props.setProperty(UD_PROP_KEY, ss.getId());
  Logger.log('✅ ' + UD_FILE_NAME + ' を作りました: ' + ss.getUrl());
  return ss;
}

function udSheet_(ss, name, cols) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, cols.length).setValues([cols]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, cols.length).setFontWeight('bold');
  } else if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, cols.length).setValues([cols]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/** 最初に1回だけ実行する。スプシを作ってURLを出す */
function udSetup() {
  const ss = udGetSS_();
  Object.keys(UD_TABS).forEach(function (k) {
    udSheet_(ss, UD_TABS[k].sheet, UD_TABS[k].cols);
  });
  udSheet_(ss, UD_LOG_SHEET, UD_LOG_COLS);
  // 最初から付いてくる「シート1」は邪魔なので消す
  const first = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (first && ss.getSheets().length > 1) ss.deleteSheet(first);
  Logger.log('URL: ' + ss.getUrl());
  Logger.log('タブ: ' + ss.getSheets().map(function (s) { return s.getName(); }).join(', '));
  Logger.log('★このスプシは共有を広げないでください（家賃・初期投資が入ります）');
  return ss.getUrl();
}

// ------------------------------------------------------------
//  検証（画面から来た値はここを通さないと書かない）
// ------------------------------------------------------------
function udFail_(msg) { throw new Error('入力を保存できません: ' + msg); }

function udStoreSet_() {
  // 店舗一覧スプシの略称（B列）だけを通す。実在しない店名は書かせない
  const out = {};
  try {
    const sh = getSS('SS_MASTER').getSheetByName('店舗一覧');
    const rows = sh.getDataRange().getValues();
    for (let r = 1; r < rows.length; r++) {
      const short = (rows[r][1] || '').toString().trim();
      if (short) out[short] = true;
    }
  } catch (e) {
    Logger.log('⚠️ 店舗一覧が読めません: ' + e);
  }
  return out;
}

function udText_(v, label, max) {
  const t = (v == null ? '' : v.toString()).trim();
  if (t.length > (max || UD_MAX_TEXT)) udFail_(label + 'が長すぎます（' + (max || UD_MAX_TEXT) + '文字まで）');
  return t;
}
function udNum_(v, label, min, max) {
  if (v === '' || v == null) return '';
  const n = Number(v.toString().replace(/[,，\s円]/g, ''));
  if (!isFinite(n)) udFail_(label + 'は数で入れてください');
  if (n < min || n > max) udFail_(label + 'は' + min + '〜' + max + 'の範囲で入れてください');
  return Math.round(n * 100) / 100;
}
function udYm_(v, label) {
  // 「2026-01」「2026/1」「202601」「2026年1月」を受ける。
  //   ★区切りを消してから6桁で見ると「2026/1」が 20261 になって弾かれるので、別々に見る
  //   （BMの来店日で同じ踏み方をした。2026-11-13）
  const t = (v == null ? '' : v.toString()).trim();
  if (!t) return '';
  const m = t.match(/^(\d{4})[\/\-.年]?\s*(\d{1,2})\s*月?$/);
  if (!m) udFail_(label + 'は「2026-01」の形で入れてください');
  const mo = parseInt(m[2], 10);
  if (mo < 1 || mo > 12) udFail_(label + 'の月がおかしいです');
  return m[1] + ('0' + mo).slice(-2);
}
function udDate_(v, label) {
  if (!v) return '';
  const a = accYmd_(v);
  if (!a) udFail_(label + 'は「2026-11-20」の形で入れてください');
  const pad = function (n) { return ('0' + n).slice(-2); };
  return a[0] + '-' + pad(a[1]) + '-' + pad(a[2]);
}
function udBool_(v) { return (v === true || v === 'true' || v === 1 || v === '1' || v === '済') ? '済' : ''; }
function udPick_(v, list, label) {
  const t = (v == null ? '' : v.toString()).trim();
  if (!t) return list[0];
  if (list.indexOf(t) < 0) udFail_(label + 'は ' + list.join('／') + ' から選んでください');
  return t;
}

/** 1件ぶんを、タブの列の並びに合わせた配列にする */
function udRowOf_(kind, rec, stores, email, now) {
  const t = UD_TABS[kind];
  if (!t) udFail_('知らない種類です: ' + kind);
  if (!rec) udFail_('中身がありません');
  const store = udText_(rec.store, '店舗', 120);
  if (t.keys.indexOf('店舗') >= 0) {
    if (!store) udFail_('店舗がありません');
    // 週メモだけは「（全体）」を使える（その週全体のメモ。店に紐づかない）
    const isTeam = (kind === 'weekmemo' && store === UD_TEAM_STORE);
    if (!isTeam && Object.keys(stores).length && !stores[store]) {
      udFail_('店舗一覧にない店です: ' + store);
    }
  }
  if (kind === 'memo') {
    return [udYm_(rec.month, '月'), store,
            udPick_(rec.status, UD_STATUS, '状態'),
            udDate_(rec.roleplay_date, 'ロープレ予定日'), udBool_(rec.roleplay_done),
            udDate_(rec.sb_date, 'サロボ修正予定日'),   udBool_(rec.sb_done),
            udText_(rec.sb_detail, 'サロボ修正内容'), udText_(rec.memo, 'メモ'),
            email, now];
  }
  if (kind === 'weekmemo') {
    return [udYm_(rec.month, '月'), udNum_(rec.week, '週', 1, 6), store,
            udText_(rec.memo, 'メモ'), email, now];
  }
  if (kind === 'finance') {
    return [store, udYm_(rec.open_ym, 'オープン年月'),
            udNum_(rec.rent, '家賃（万円）', 0, 999),
            udNum_(rec.capex, '初期投資（万円）', 0, 99999),
            udNum_(rec.high_ratio, '上位回数券の比率（%）', 0, 100),
            udPick_(rec.ad_mode, UD_AD_MODE, '広告費の決め方'),
            udNum_(rec.ad_cost, '広告費（万円）', 0, 999),
            udBool_(rec.hire_second), udYm_(rec.hire_ym, '採用活動の開始月'),
            udNum_(rec.payback_months, '回収期限（か月）', 1, 240),
            email, now];
  }
  if (kind === 'monthly') {
    return [store, udYm_(rec.month, '月'),
            udNum_(rec.ad_cost, '広告費（万円）', 0, 999),
            udNum_(rec.staff_count, 'スタッフ数', 0, 99),
            email, now];
  }
  if (kind === 'action') {
    return [store, udText_(rec.action, 'アクション', 200),
            udPick_(rec.status, UD_ACTION, '状態'), email, now];
  }
  udFail_('知らない種類です: ' + kind);
}

// ------------------------------------------------------------
//  読む（画面から google.script.run で呼ぶ）
// ------------------------------------------------------------
/**
 * 入力データを読む。
 *   kind  … 'memo' / 'weekmemo' / 'finance' / 'monthly' / 'action'
 *   store … 指定すればその店だけ（将来の店舗別リンク用。本部は省略）
 * 戻り値: { ok:true, cols:[...], rows:[[...]] }
 */
function udRead(kind, store) {
  if (!webappIsAllowed_()) return { ok: false, error: '閲覧権限がありません' };
  const t = UD_TABS[kind];
  if (!t) return { ok: false, error: '知らない種類です: ' + kind };
  try {
    const ss = udGetSS_();
    const sh = udSheet_(ss, t.sheet, t.cols);
    const last = sh.getLastRow();
    if (last < 2) return { ok: true, cols: t.cols, rows: [] };
    const vals = sh.getRange(2, 1, last - 1, t.cols.length).getValues();
    const iStore = t.cols.indexOf('店舗');
    const want = (store || '').toString().trim();
    const rows = vals.filter(function (r) {
      if (!r.join('')) return false;
      if (want && iStore >= 0 && (r[iStore] || '').toString().trim() !== want) return false;
      return true;
    }).map(function (r) {
      // 日付セルが Date で返ることがあるので文字にそろえる
      return r.map(function (v) {
        if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
        return v;
      });
    });
    return { ok: true, cols: t.cols, rows: rows };
  } catch (e) {
    return { ok: false, error: String(e && e.message ? e.message : e) };
  }
}

/** まとめて読む。画面の初期表示用 */
function udReadAll(store) {
  const out = {};
  Object.keys(UD_TABS).forEach(function (k) { out[k] = udRead(k, store); });
  return out;
}

// ------------------------------------------------------------
//  書く（画面から google.script.run で呼ぶ）
// ------------------------------------------------------------
/**
 * 入力データを保存する。キーが同じ行があれば上書き、無ければ足す。
 *   kind … UD_TABS のキー
 *   recs … 1件のオブジェクト、またはその配列
 * 戻り値: { ok:true, saved:n, by:'メール', at:'日時' }
 */
function udWrite(kind, recs) {
  if (!webappIsAllowed_()) return { ok: false, error: '書き込み権限がありません' };
  const t = UD_TABS[kind];
  if (!t) return { ok: false, error: '知らない種類です: ' + kind };
  const list = (recs && recs.length !== undefined) ? recs : [recs];
  if (!list.length) return { ok: true, saved: 0 };
  if (list.length > 500) return { ok: false, error: '一度に保存できるのは500件までです' };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (e) {
    return { ok: false, error: 'ほかの保存と重なりました。少し待ってもう一度お願いします' };
  }
  try {
    const email = webappUserEmail_() || '不明';
    const now   = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    const stores = udStoreSet_();
    const ss = udGetSS_();
    const sh = udSheet_(ss, t.sheet, t.cols);

    // いまある行のキー → 行番号
    const last = sh.getLastRow();
    const have = {};
    const keyIdx = t.keys.map(function (k) { return t.cols.indexOf(k); });
    if (last >= 2) {
      const vals = sh.getRange(2, 1, last - 1, t.cols.length).getValues();
      vals.forEach(function (r, i) {
        const k = keyIdx.map(function (ci) { return (r[ci] == null ? '' : r[ci]).toString().trim(); }).join('\u0001');
        if (k.replace(/\u0001/g, '')) have[k] = i + 2;
      });
    }
    if (last - 1 + list.length > UD_MAX_ROWS) {
      return { ok: false, error: '行が多すぎます（' + UD_MAX_ROWS + '行まで）' };
    }

    // 先に全部検証する（途中で失敗して半分だけ書くのを防ぐ）
    const rows = list.map(function (rec) { return udRowOf_(kind, rec, stores, email, now); });

    let updated = 0, added = 0;
    const appends = [];
    rows.forEach(function (row) {
      const k = keyIdx.map(function (ci) { return (row[ci] == null ? '' : row[ci]).toString().trim(); }).join('\u0001');
      if (have[k]) {
        sh.getRange(have[k], 1, 1, t.cols.length).setValues([row]);
        updated++;
      } else {
        appends.push(row);
        have[k] = -1;   // 同じ呼び出しの中で重複しても1行にする
        added++;
      }
    });
    if (appends.length) {
      sh.getRange(sh.getLastRow() + 1, 1, appends.length, t.cols.length).setValues(appends);
    }

    udLog_(ss, now, email, t.sheet, rows, updated, added);
    return { ok: true, saved: rows.length, updated: updated, added: added, by: email, at: now };
  } catch (e) {
    return { ok: false, error: String(e && e.message ? e.message : e) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function udLog_(ss, now, email, sheetName, rows, updated, added) {
  try {
    const sh = udSheet_(ss, UD_LOG_SHEET, UD_LOG_COLS);
    const keys = rows.map(function (r) { return r[0] + '/' + (r[1] || ''); }).slice(0, 20).join(', ');
    sh.appendRow([now, email, sheetName,
                  keys + (rows.length > 20 ? ' ほか' + (rows.length - 20) + '件' : ''),
                  '上書き' + updated + '件・追加' + added + '件']);
  } catch (e) {
    Logger.log('⚠️ 書き込みログを残せません: ' + e);
  }
}

// ------------------------------------------------------------
//  動作確認（画面の「接続を確認」から呼ぶ）
// ------------------------------------------------------------
/** 保存先の状態を返す。書き込みはしない */
function udStatus() {
  if (!webappIsAllowed_()) return { ok: false, error: '閲覧権限がありません' };
  try {
    const props = PropertiesService.getScriptProperties();
    const id = props.getProperty(UD_PROP_KEY);
    if (!id) {
      return { ok: true, ready: false, by: webappUserEmail_() || '不明',
               message: 'まだ保存先がありません。最初の保存のときに作られます' };
    }
    const ss = SpreadsheetApp.openById(id);
    const tabs = {};
    Object.keys(UD_TABS).forEach(function (k) {
      const sh = ss.getSheetByName(UD_TABS[k].sheet);
      tabs[UD_TABS[k].sheet] = sh ? Math.max(0, sh.getLastRow() - 1) : null;
    });
    return { ok: true, ready: true, name: ss.getName(), by: webappUserEmail_() || '不明', tabs: tabs };
  } catch (e) {
    return { ok: false, error: String(e && e.message ? e.message : e) };
  }
}

/** GASエディタから動作確認する（テスト用の行を書いて、読んで、消す） */
function udSelfTest() {
  const probe = '__動作確認__';
  const r1 = udWrite('action', { store: probe, action: 'せつぞく確認', status: '完了' });
  Logger.log('書き込み: ' + JSON.stringify(r1));
  if (!r1.ok && String(r1.error).indexOf('店舗一覧にない店') >= 0) {
    Logger.log('→ 店名の検証が効いています（想定どおり）');
  }
  const st = udStatus();
  Logger.log('状態: ' + JSON.stringify(st));
  const r2 = udRead('action');
  Logger.log('読み取り: ' + (r2.ok ? r2.rows.length + '行' : r2.error));
  return st;
}
