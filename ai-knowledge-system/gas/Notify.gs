/**
 * ============================================================
 *  Notify.gs : Google Chat への通知
 * ------------------------------------------------------------
 *  「実行結果」を Google Chat のスペースに流し、全員がリンクから
 *  すぐ結果を見られるようにする。人はDriveにファイルを置くだけ。
 *
 *  準備(1回だけ・とても簡単):
 *   1. Google Chat で通知用スペースを作る（例:「AI営業ボット」）
 *   2. スペース名 →「アプリと統合」→「Webhook」→ 追加 → URLをコピー
 *   3. スクリプトプロパティ CHAT_WEBHOOK_URL にそのURLを貼る
 *  （未設定でも動作します。その場合は通知だけスキップされます）
 * ============================================================
 */

/**
 * 結果メールの送信先を返す（カンマ区切り文字列）。
 *  「配信先」シートの有効な行 ∪ スクリプトプロパティ REPORT_RECIPIENTS を統合・重複除去。
 *  → 社員が増えたら「配信先」シートにメールを1行足すだけで届くようになる（管理画面化）。
 */
function recipients_() {
  var set = {};
  // (1) 配信先シート
  try {
    var sh = openBook_().getSheetByName(CONFIG.SHEET_RECIPIENTS);
    if (sh && sh.getLastRow() > 1) {
      var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
      vals.forEach(function (r) {
        var email = String(r[0] || '').trim();
        var status = String(r[2] || '').trim().toUpperCase();
        if (email && email.indexOf('@') > 0 && status !== 'OFF') set[email] = true;
      });
    }
  } catch (e) { Logger.log('配信先シート読取り失敗: ' + e); }
  // (2) 従来のプロパティも常に含める（管理者アドレスの保険）
  var fallback = prop_('REPORT_RECIPIENTS', false) || '';
  fallback.split(',').forEach(function (e) {
    var em = e.trim(); if (em && em.indexOf('@') > 0) set[em] = true;
  });
  return Object.keys(set).join(',');
}

/** Chatにテキスト通知（未設定なら黙ってスキップ） */
function notifyChat_(text) {
  var url = prop_('CHAT_WEBHOOK_URL', false);
  if (!url) { Logger.log('[Chat未設定のため通知スキップ] ' + text); return; }
  try {
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json; charset=UTF-8',
      payload: JSON.stringify({ text: text }),
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log('Chat通知失敗: ' + e);
  }
}

/** Chat用のリンク書式 <URL|ラベル> */
function chatLink_(url, label) {
  if (!url) return label || '';
  return '<' + url + '|' + (label || url) + '>';
}

/** 通知が届くかのテスト（手動実行用） */
function testChatNotify() {
  notifyChat_('✅ 通知テスト：AI営業スイートと Google Chat がつながりました。');
  Logger.log('testChatNotify 実行。CHAT_WEBHOOK_URL 未設定ならログのみ。');
}
