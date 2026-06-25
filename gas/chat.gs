// ============================================================
//  ダッシュボード埋め込みチャット（Claude API プロキシ）
//  ・ダッシュボード(GitHub Pages) → このGAS Webアプリ(doPost) → Claude API
//  ・APIキーはGAS側に隠す（クライアントに出さない）
//  【スクリプトプロパティ】
//    ANTHROPIC_API_KEY : Claude APIキー（console.anthropic.comで発行）
//    CHAT_BUDGET_JPY   : 月の利用上限(円)。未設定なら5000
//  使い方：①プロパティ設定 → ②testChat()でログ確認 → ③Webアプリとしてデプロイ
// ============================================================

const CHAT_MODEL    = 'claude-sonnet-4-6';
const CHAT_PRICE_IN = 3;     // 入力 $/1Mトークン
const CHAT_PRICE_OUT= 15;    // 出力 $/1Mトークン
const CHAT_USD_JPY  = 155;   // 概算レート（円換算用。多少のズレは安全側に倒すなら大きめに）
const CHAT_INSIGHTS_URL = 'https://raw.githubusercontent.com/kimu2-lime/dashboard-master/main/data/reservation_insights.json';

// ── 予算管理（月ごと・円） ──
function chatBudgetJpy_() {
  const v = PropertiesService.getScriptProperties().getProperty('CHAT_BUDGET_JPY');
  return v ? parseFloat(v) : 5000;
}
function chatSpendKey_() {
  return 'CHAT_SPEND_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMM');
}
function chatSpent_() {
  const v = PropertiesService.getScriptProperties().getProperty(chatSpendKey_());
  return v ? parseFloat(v) : 0;
}
function chatAddSpend_(jpy) {
  PropertiesService.getScriptProperties().setProperty(chatSpendKey_(), (chatSpent_() + jpy).toFixed(2));
}

// ── 予約インサイトを文脈として読む（GitHub rawをCacheServiceで10分キャッシュ） ──
function chatInsights_() {
  const cache = CacheService.getScriptCache();
  let s = cache.get('chat_insights');
  if (!s) {
    try {
      const r = UrlFetchApp.fetch(CHAT_INSIGHTS_URL + '?t=' + Date.now(), { muteHttpExceptions: true });
      if (r.getResponseCode() === 200) { s = r.getContentText(); cache.put('chat_insights', s, 600); }
    } catch (e) {}
  }
  try { return s ? JSON.parse(s) : null; } catch (e) { return null; }
}

// ── システムプロンプト（実データを文脈に） ──
function chatSystemPrompt_() {
  const RI = chatInsights_();
  let c = 'あなたはLIME FIT 売上ダッシュボードのアシスタントです。店長・SV・本部の質問に、以下の実データを根拠に日本語で簡潔・具体的に答えてください。'
        + '数字は提供データに基づき、無い情報は「データにありません」と答える。憶測で数字を作らない。\n\n';
  if (RI) {
    c += '【対象月】' + RI.month + '（予約CSV由来。契約=会計時メニューに「契約」を含む新規契約、母数=来店完了）\n\n';
    if (RI.overall && RI.overall.keywords) {
      c += '【全店：予約クーポンのキーワード別 契約率】\n';
      RI.overall.keywords.forEach(function(k){ c += '・' + k.keyword + ' ' + k.rate + '%(n=' + k.n + ')\n'; });
      c += '\n';
    }
    if (RI.overall && RI.overall.coupons) {
      c += '【全店：クーポン別 契約率 上位】\n';
      RI.overall.coupons.slice(0, 15).forEach(function(x){ c += '・' + x.rate + '% ' + x.coupon + '(n=' + x.n + ')\n'; });
      c += '\n';
    }
    c += '【店舗別 予約インサイト】略称: 予約/来店/CXL率/HPB率/指名率/女性%/初回契約率/新規契約\n';
    Object.keys(RI.stores).forEach(function(name){
      const s = RI.stores[name];
      c += name + ': 予約' + s.total_reservations + ' 来店' + s.visited + ' CXL' + s.cancel_rate + '% HPB' + s.hpb_rate
         + '% 指名' + s.shimei_rate + '% 女' + s.gender_female_pct + '% 初回契約' + s.first_contract_rate
         + '% 新規契約' + s.contract_count + '\n';
    });
    c += '\n';
  } else {
    c += '（予約インサイトデータを取得できませんでした）\n\n';
  }
  c += '【ナレッジ】悩み訴求ワード(ニキビ/毛穴/いちご鼻/黒ずみ/水光肌)は契約率が高い傾向。'
     + 'クーポン名を提案する時は「共感(悩み)＋手法(韓国式ハーブピーリング)＋信頼(口コミ★4.9等)」の型で、不自然でない自然な日本語で作る。\n\n'
     + '【回答フォーマット】チャットUIで表示されるので、マークダウンの表は使わず、短い文と箇条書き(・)で簡潔に。長くなりすぎない。';
  return c;
}

// ── Claude API 呼び出し ──
function chatCall_(question, history) {
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) return { error: 'ANTHROPIC_API_KEY が未設定です（スクリプトプロパティに設定してください）' };
  const messages = (history && history.length) ? history.slice() : [];
  messages.push({ role: 'user', content: String(question || '').slice(0, 4000) });
  const payload = {
    model: CHAT_MODEL,
    max_tokens: 1024,
    system: [{ type: 'text', text: chatSystemPrompt_(), cache_control: { type: 'ephemeral' } }],
    messages: messages
  };
  const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const code = resp.getResponseCode();
  let body; try { body = JSON.parse(resp.getContentText()); } catch (e) { return { error: 'API応答の解析に失敗' }; }
  if (code !== 200) return { error: 'API ' + code + ': ' + (body.error ? body.error.message : resp.getContentText().slice(0, 200)) };
  const text = (body.content || []).filter(function(b){ return b.type === 'text'; }).map(function(b){ return b.text; }).join('');
  const u = body.usage || {};
  const costUsd = ((u.input_tokens || 0) * CHAT_PRICE_IN
                 + (u.cache_creation_input_tokens || 0) * CHAT_PRICE_IN * 1.25
                 + (u.cache_read_input_tokens || 0) * CHAT_PRICE_IN * 0.1
                 + (u.output_tokens || 0) * CHAT_PRICE_OUT) / 1e6;
  return { text: text, costJpy: costUsd * CHAT_USD_JPY, usage: u };
}

// ── Webアプリ受け口（ダッシュボードから fetch される） ──
function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const budget = chatBudgetJpy_();
    const spent  = chatSpent_();
    if (spent >= budget) {
      out = { error: '今月の利用上限（' + budget + '円）に達しました。来月またご利用ください。', spent: Math.round(spent), budget: budget };
    } else {
      const r = chatCall_(req.question, req.history);
      if (r.error) {
        out = { error: r.error };
      } else {
        chatAddSpend_(r.costJpy);
        out = { answer: r.text, spent: Math.round(chatSpent_()), budget: budget };
      }
    }
  } catch (err) {
    out = { error: 'サーバエラー: ' + err };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// ── 動作確認（Webアプリ化の前に実行してログを見る） ──
function testChat() {
  const r = chatCall_('契約に一番効いてるクーポンのキーワードを上位3つ、数字付きで教えて', []);
  Logger.log(JSON.stringify(r, null, 2));
  if (r.text) {
    Logger.log('--- 回答 ---\n' + r.text);
    Logger.log('この1回のコスト概算: ' + (r.costJpy ? r.costJpy.toFixed(2) : '?') + '円');
  }
  Logger.log('今月の累計使用額: ' + chatSpent_().toFixed(2) + '円 / 上限: ' + chatBudgetJpy_() + '円');
}
