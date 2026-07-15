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
