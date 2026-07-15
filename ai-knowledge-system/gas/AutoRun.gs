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

function autoRunAll() {
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
      deals.forEach(function (d) { lines.push('　・' + chatLink_(d.slideUrl, d.customer + ' 様の提案書')); });
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
        props.forEach(function (p) { lines.push('　・' + chatLink_(p.url, p.customer + ' の提案書')); });
      }
    }, 'DM提案');
  }

  // 承認モードでは承認カードを個別に送るので、まとめ通知は他に新規がある時だけ
  if (lines.length) {
    var ss = '';
    try { ss = '\n\n📂 ' + chatLink_(openBook_().getUrl(), 'データ一覧(スプレッドシート)'); } catch (e) {}
    notifyChat_('✅ 自動処理が完了しました\n\n' + lines.join('\n') + ss);
  } else {
    Logger.log('autoRunAll: 新規なし（通知なし）');
  }
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
