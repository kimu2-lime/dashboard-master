// ============================================================
//  ダッシュボードを「社内のGoogleアカウントだけ」見られる形で配信する
//
//  いまの GitHub Pages 版はそのまま動かしたまま、こちらを並行で足している。
//  切り替えの判断がつくまで、どちらでも同じ画面が見られる。
//
//  ・doGet()        … index.html を配信する（Googleログインが要る）
//  ・getDashboardJson() … データをGoogleドライブの非公開フォルダから返す
//  ・chatAsk()      … チャットを google.script.run から直接呼ぶ（公開の doPost を通さない）
//
//  画面側（index.html）は1ファイルで両対応している。
//    google.script.run がある → こちら（ドライブから読む・パスワード画面なし）
//    なければ                 → これまでどおり data/data.json を取りに行く
//
//  【スクリプトプロパティ】
//    DRIVE_DATA_FOLDER_ID : データを置くドライブのフォルダID（無ければ自動で作る）
//    ALLOWED_DOMAIN       : 許可するメールのドメイン（既定 lime-fit.com）
// ============================================================

const WEBAPP_FOLDER_NAME = 'LIMEFIT ダッシュボード データ（非公開）';
const WEBAPP_DEFAULT_DOMAIN = 'lime-fit.com';

// 画面に渡すファイル名 → ドライブ上のファイル名
const WEBAPP_FILES = {
  'data': 'data.json',
  'reservation_insights': 'reservation_insights.json',
  'ai-analysis': 'ai-analysis.json'
};


// ── 見てよい人か ──
function webappAllowedDomain_() {
  return (PropertiesService.getScriptProperties().getProperty('ALLOWED_DOMAIN')
          || WEBAPP_DEFAULT_DOMAIN).toLowerCase();
}

function webappUserEmail_() {
  try {
    return (Session.getActiveUser().getEmail() || '').toLowerCase();
  } catch (e) {
    return '';
  }
}

/**
 * アクセスしてよい人かどうか。
 *
 * デプロイ側で「ドメイン内の全員」に絞るのが本線。
 * ここは二重のふたで、メールが取れたときだけドメインを見る。
 * （取れないことがあるので、取れない＝拒否にはしない。閉め出しを避ける）
 */
function webappIsAllowed_() {
  const mail = webappUserEmail_();
  if (!mail) return true;
  return mail.indexOf('@' + webappAllowedDomain_()) >= 0;
}


// ── ドライブの置き場所 ──
function webappFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('DRIVE_DATA_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) {
      Logger.log('⚠️ DRIVE_DATA_FOLDER_ID のフォルダを開けません: ' + e);
    }
  }
  // 無ければ作る。作ったIDはプロパティに覚える（共有はしない＝作った人だけ）
  const it = DriveApp.getFoldersByName(WEBAPP_FOLDER_NAME);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(WEBAPP_FOLDER_NAME);
  props.setProperty('DRIVE_DATA_FOLDER_ID', folder.getId());
  Logger.log('📁 データ置き場: ' + folder.getName() + ' (' + folder.getId() + ')');
  return folder;
}

/** ドライブにファイルを置く（同名があれば中身を差し替える）。 */
function webappSaveFile_(name, text) {
  try {
    const folder = webappFolder_();
    const it = folder.getFilesByName(name);
    if (it.hasNext()) {
      const f = it.next();
      f.setContent(text);
      while (it.hasNext()) it.next().setTrashed(true);   // 重複があれば片付ける
      return f.getId();
    }
    return folder.createFile(name, text, MimeType.PLAIN_TEXT).getId();
  } catch (e) {
    Logger.log('⚠️ ドライブに保存できません（' + name + '）: ' + e);
    return null;
  }
}

/** ドライブからファイルの中身を読む。無ければ null。 */
function webappReadFile_(name) {
  try {
    const it = webappFolder_().getFilesByName(name);
    if (!it.hasNext()) return null;
    return it.next().getBlob().getDataAsString('UTF-8');
  } catch (e) {
    Logger.log('⚠️ ドライブから読めません（' + name + '）: ' + e);
    return null;
  }
}


// ── 画面から呼ばれるもの ──

/**
 * ダッシュボードのデータを返す（文字列のまま。画面側で JSON.parse する）。
 * name は 'data' / 'reservation_insights' / 'ai-analysis'。
 */
function getDashboardJson(name) {
  if (!webappIsAllowed_()) throw new Error('閲覧権限がありません');
  const file = WEBAPP_FILES[name];
  if (!file) throw new Error('知らないデータです: ' + name);
  return webappReadFile_(file);
}

/** チャット。公開の doPost を通さず、ログイン済みの人からだけ呼べる。 */
function chatAsk(question, history) {
  if (!webappIsAllowed_()) return { error: '閲覧権限がありません' };
  try {
    const budget = chatBudgetJpy_();
    const spent = chatSpent_();
    if (spent >= budget) {
      return { error: '今月の利用上限（' + budget + '円）に達しました。来月またご利用ください。',
               spent: Math.round(spent), budget: budget };
    }
    const r = chatCall_(question, history || []);
    if (r.error) return { error: r.error };
    chatAddSpend_(r.costJpy);
    return { answer: r.text, spent: Math.round(chatSpent_()), budget: budget };
  } catch (err) {
    return { error: 'サーバエラー: ' + err };
  }
}


// ── 配信 ──
function doGet(e) {
  if (!webappIsAllowed_()) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;padding:40px;text-align:center">'
      + '<h2>閲覧権限がありません</h2>'
      + '<p>' + webappAllowedDomain_() + ' のアカウントでログインしてください。</p>'
      + '<p style="color:#888;font-size:13px">いまログイン中：'
      + (webappUserEmail_() || '(不明)') + '</p></div>');
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('LIME FIT 売上ダッシュボード')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ── 手で動かすもの ──

/**
 * いまGitHubに出ている3つのJSONを、ドライブにも置く。
 * 初回の移行用。これ以降は exportToGitHub / reservation 側が自動で置く。
 */
function mirrorAllToDrive() {
  const owner = PropertiesService.getScriptProperties().getProperty('GITHUB_OWNER') || 'kimu2-lime';
  const repo = PropertiesService.getScriptProperties().getProperty('GITHUB_REPO') || 'dashboard-master';
  const base = 'https://raw.githubusercontent.com/' + owner + '/' + repo + '/main/data/';
  ['data.json', 'reservation_insights.json', 'ai-analysis.json'].forEach(function (f) {
    try {
      const r = UrlFetchApp.fetch(base + f, { muteHttpExceptions: true });
      if (r.getResponseCode() !== 200) {
        Logger.log('⚠️ ' + f + ' を取れません（HTTP ' + r.getResponseCode() + '）'); return;
      }
      const t = r.getContentText();
      webappSaveFile_(f, t);
      Logger.log('✅ ' + f + ' → ドライブ（' + Math.round(t.length / 1024) + ' KB）');
    } catch (err) {
      Logger.log('⚠️ ' + f + ': ' + err);
    }
  });
  const kb = webappReadFile_('chat_knowledge.md');
  if (!kb) {
    try {
      const r = UrlFetchApp.fetch(
        'https://raw.githubusercontent.com/' + owner + '/' + repo + '/main/knowledge/chat_knowledge.md',
        { muteHttpExceptions: true });
      if (r.getResponseCode() === 200) {
        webappSaveFile_('chat_knowledge.md', r.getContentText());
        Logger.log('✅ chat_knowledge.md → ドライブ');
      }
    } catch (err) { Logger.log('⚠️ chat_knowledge.md: ' + err); }
  }
  Logger.log('📁 置き場所: ' + webappFolder_().getUrl());
}

/** 設定とドライブの中身を確かめる。 */
function checkWebapp() {
  const folder = webappFolder_();
  Logger.log('ログイン中: ' + (webappUserEmail_() || '(取れません)'));
  Logger.log('許可ドメイン: ' + webappAllowedDomain_());
  Logger.log('フォルダ: ' + folder.getName() + '  ' + folder.getUrl());
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    Logger.log('  - %s  %s KB  %s', f.getName(),
               Math.round(f.getSize() / 1024), f.getLastUpdated());
  }
}
