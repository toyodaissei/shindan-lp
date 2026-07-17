/**
 * ============================================================
 *  AutoRun.gs : ぜんぶ自動の司令塔
 * ------------------------------------------------------------
 *  これ1本を時間トリガー(15分毎)で回すだけで、
 *   ・ナレッジ要約   ・商談→提案書
 *   ・DM録画/スクショ解析   ・DM提案生成
 *  を順番に実行し、新しく出来たものだけを Google Chat に通知する。
 *
 *  現場の人がやること = Driveの決まったフォルダにファイルを置くだけ。
 *  Apps Script は二度と開かなくてよい。
 * ============================================================
 */

/** メニュー「▶ 今すぐ全部まとめて実行」用：手動実行なので毎回メールを送る */
function runNow() {
  autoRunAll(true);
}

/**
 * @param {boolean} manual  手動実行(メニュー)なら true。true のときは結果が無くても
 *                          「実行しました」メールを必ず送る。自動トリガーは新規がある時だけ送る。
 */
function autoRunAll(manual) {
  var lines = [];

  // 🅱 ナレッジ要約（NotebookLM母艦Docにも自動追記）
  safe_(function () {
    var n = processPending();
    if (n) lines.push('🧠 社内ナレッジを ' + n + ' 件、要約・蓄積しました');
  }, 'ナレッジ要約');

  // 🅰 商談 → 提案書スライド
  safe_(function () {
    var deals = processDeals() || [];
    if (deals.length) {
      lines.push('📊 商談の提案書を ' + deals.length + ' 件つくりました');
      deals.forEach(function (d) { lines.push('　・' + d.customer + ' 様の提案書: ' + d.slideUrl); });
    }
  }, '商談提案書');

  // 🆕 DM録画/スクショ 解析
  safe_(function () {
    var got = ingestDmRecordings();
    if (got) lines.push('📹 DMの録画/スクショを ' + got + ' 件、解析しました');
  }, 'DM解析');

  if (CONFIG.APPROVAL_MODE) {
    // 承認モード：提案は自動生成せず、Chatに承認カードを送る（人がボタンで承認）
    safe_(function () {
      var cards = postPendingApprovalCards_();
      if (cards) lines.push('🔔 承認待ちの案件 ' + cards + ' 件をChatに送りました（「✅提案を作成」を押してください）');
    }, 'DM承認カード');
  } else {
    // 全自動：そのまま提案まで生成
    safe_(function () {
      var props = proposeForPending() || [];
      if (props.length) {
        lines.push('📝 エージェント開拓の提案書を ' + props.length + ' 件つくりました');
        props.forEach(function (p) { lines.push('　・' + p.customer + ' の提案書: ' + p.url); });
      }
    }, 'DM提案');
  }

  var hasNew = lines.length > 0;
  if (!hasNew && !manual) { Logger.log('autoRunAll: 新規なし（通知なし）'); return; }

  // 参照リンク（結果の置き場所）
  var refs = [];
  try { refs.push('📂 データ一覧(スプレッドシート): ' + openBook_().getUrl()); } catch (e) {}
  try {
    var pf = prop_('PROPOSAL_FOLDER', false);
    if (pf) refs.push('🗂 提案書フォルダ: ' + DriveApp.getFolderById(pf).getUrl());
  } catch (e2) {}

  var head = hasNew
    ? '✅ 処理が完了しました。生成物は下記のとおりです。'
    : 'ℹ 今回は新しく処理する対象がありませんでした（フォルダに未処理の新しいファイルが無かった可能性があります）。';
  var body = head + '\n\n' + (hasNew ? lines.join('\n') + '\n\n' : '') +
    '── 保存場所 ──\n' + refs.join('\n');

  // 📧 メール通知（毎回：手動実行 or 新規があった時）→ REPORT_RECIPIENTS 宛
  var to = prop_('REPORT_RECIPIENTS', false);
  if (to) {
    try {
      MailApp.sendEmail(to,
        '【AI営業スイート】実行結果 ' + Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'M/d HH:mm'),
        body);
    } catch (e3) { Logger.log('メール送信失敗: ' + e3); }
  } else {
    Logger.log('REPORT_RECIPIENTS未設定のためメール送信スキップ');
  }
  // Chat通知（設定している場合のみ）
  notifyChat_(body);
}

/** 週次レポートもChatへ流す（トリガーから呼ぶ） */
function autoWeeklyReports() {
  safe_(function () { generateWeeklyReport(); }, '週次経営レポート');
  safe_(function () { generateDmReport(); }, 'DM営業レポート');
  notifyChat_('🗓 今週のレポートを更新しました（経営レポート / DM営業レポート）。メールもご確認ください。');
}

/** エラーが出ても全体を止めない小さなラッパ */
function safe_(fn, label) {
  try { fn(); }
  catch (e) {
    Logger.log('[' + label + '] 失敗: ' + e);
    notifyChat_('⚠ ' + label + ' の処理でエラーが出ました: ' + String(e).slice(0, 200));
  }
}
