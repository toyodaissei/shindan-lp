/**
 * ============================================================
 *  Chatbot.gs : Google Chat 承認ボタン フロー（Cloud設定なし版）
 * ------------------------------------------------------------
 *  仕組み:
 *   ・Chatには Incoming Webhook で「カード＋ボタン」を送る
 *   ・ボタンは Apps Script の Webアプリ(doGet) を開くリンク
 *   ・開くと承認アクション(提案生成/送付)が実行され、確認ページを表示
 *  → フルのChatボット(Google Cloud設定)なしで「承認して実行」を実現。
 *
 *  必要な準備:
 *   1. CHAT_WEBHOOK_URL … 通知用スペースのWebhook（Notify.gs参照）
 *   2. Webアプリをデプロイ（デプロイ→新しいデプロイ→ウェブアプリ／アクセス:全員）
 *      → setup()が自動でURLを掴みますが、掴めない場合は WEBAPP_URL に貼る
 *   3. APPROVE_TOKEN … setup()が自動生成（ボタンURLの改ざん防止トークン）
 * ============================================================
 */

/** DM案件1件の「承認カード」をChatに送る */
function postDmApprovalCard_(caseId, customer, summary, outcome, fileUrl) {
  var proposeUrl = actionUrl_('propose', caseId);
  var skipUrl = actionUrl_('skip', caseId);
  var buttons = [];
  if (proposeUrl) buttons.push(cardButton_('✅ 提案を作成', proposeUrl));
  if (fileUrl)    buttons.push(cardButton_('🎬 録画/スクショを見る', fileUrl));
  if (skipUrl)    buttons.push(cardButton_('⏭ 見送る', skipUrl));

  postCard_({
    cardId: 'dm-approve-' + caseId,
    card: {
      header: { title: '📹 新しいDM解析', subtitle: customer || '（相手不明）' },
      sections: [{
        widgets: [
          { textParagraph: { text: '<b>結果:</b> ' + (outcome || '-') + '<br><b>要約:</b> ' + trim_(summary, 300) } },
          { buttonList: { buttons: buttons } }
        ]
      }]
    }
  });
}

/** 提案が出来たことを知らせる「送信カード」をChatに送る */
function postProposalReadyCard_(caseId, customer, docUrl) {
  var buttons = [];
  if (docUrl) buttons.push(cardButton_('📄 提案書を開く', docUrl));
  var sendUrl = actionUrl_('send', caseId);
  if (sendUrl) buttons.push(cardButton_('📧 担当に送信', sendUrl));
  var regenUrl = actionUrl_('regen', caseId);
  if (regenUrl) buttons.push(cardButton_('🔁 作り直す', regenUrl));

  postCard_({
    cardId: 'dm-ready-' + caseId,
    card: {
      header: { title: '📝 提案書ができました', subtitle: customer || '' },
      sections: [{
        widgets: [
          { textParagraph: { text: 'マニュアル型＋提案型の提案書を作成しました。内容を確認して「送信」してください。' } },
          { buttonList: { buttons: buttons } }
        ]
      }]
    }
  });
}

/**
 * 承認待ちの案件にカードを送る（autoRunAllが承認モードで呼ぶ）。
 * カード送信済みは status を「承認待ち」にして二重送信を防ぐ。
 * @return {number} 送ったカード数
 */
function postPendingApprovalCards_() {
  var sh = sheet_(CONFIG.SHEET_DM, DM_HEADERS);
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var values = sh.getRange(2, 1, last - 1, DM_HEADERS.length).getValues();
  var sent = 0;
  for (var i = 0; i < values.length; i++) {
    if (values[i][DM_COL.STATUS - 1] !== '解析済(提案待ち)') continue;
    var r = values[i];
    var fileUrl = '';
    try { if (r[DM_COL.FILE_ID - 1]) fileUrl = DriveApp.getFileById(r[DM_COL.FILE_ID - 1]).getUrl(); } catch (e) {}
    postDmApprovalCard_(r[DM_COL.ID - 1], r[DM_COL.CUSTOMER - 1], r[DM_COL.SUMMARY - 1], r[DM_COL.OUTCOME - 1], fileUrl);
    sh.getRange(i + 2, DM_COL.STATUS).setValue('承認待ち');
    sent++;
  }
  return sent;
}

/**
 * Webアプリのボタン押下を処理（Code.gsのdoGetから呼ばれる）。
 * ?action=propose|send & case=<id> & token=<APPROVE_TOKEN>
 */
function handleApproveAction_(e) {
  var p = (e && e.parameter) || {};
  var token = prop_('APPROVE_TOKEN', false);
  if (!token || p.token !== token) {
    return htmlPage_('リンクが無効です', 'このリンクは無効か期限切れです。Chatの最新カードからお試しください。');
  }
  var found = findDmRowById_(p['case']);
  if (!found) return htmlPage_('見つかりません', '対象の案件が見つかりませんでした。');
  var sh = found.sheet, rowNum = found.rowNum, row = found.values;
  var customer = row[DM_COL.CUSTOMER - 1];

  if (p.action === 'propose') {
    if (row[DM_COL.STATUS - 1] === '提案済' || row[DM_COL.STATUS - 1] === '送付済') {
      return htmlPage_('作成済みです', customer + ' の提案書はすでに作成済みです。Chatのカードからご確認ください。');
    }
    var url = proposeDmStrategy_(row);
    sh.getRange(rowNum, DM_COL.PROPOSAL_URL).setValue(url);
    sh.getRange(rowNum, DM_COL.STATUS).setValue('提案済');
    postProposalReadyCard_(row[DM_COL.ID - 1], customer, url);
    return htmlPage_('✅ 提案を作成しました',
      customer + ' の提案書を作成し、Chatにリンクを送りました。<br><br>' +
      '<a href="' + url + '" target="_blank">📄 提案書を今すぐ開く</a>');
  }

  if (p.action === 'send') {
    var docUrl = row[DM_COL.PROPOSAL_URL - 1];
    if (!docUrl) return htmlPage_('まだ提案がありません', '先に「✅ 提案を作成」を押してください。');
    var to = recipients_();
    if (to) {
      MailApp.sendEmail({
        to: to,
        subject: '【エージェント開拓 提案書】' + customer,
        body: customer + ' 向けのエージェント開拓提案書です。\n\n提案書: ' + docUrl +
          '\n\n内容を確認のうえ、相手へのアプローチにご活用ください。'
      });
    }
    sh.getRange(rowNum, DM_COL.STATUS).setValue('送付済');
    notifyChat_('📧 ' + customer + ' の提案書を担当（' + (to || '未設定') + '）へ送付しました。');
    return htmlPage_('📧 送信しました', customer + ' の提案書を担当者へメールしました。');
  }

  if (p.action === 'skip') {
    sh.getRange(rowNum, DM_COL.STATUS).setValue('見送り');
    return htmlPage_('⏭ 見送りにしました', customer + ' の案件を見送りにしました。今後この案件のカードは届きません。');
  }

  if (p.action === 'regen') {
    var url2 = proposeDmStrategy_(row);
    sh.getRange(rowNum, DM_COL.PROPOSAL_URL).setValue(url2);
    sh.getRange(rowNum, DM_COL.STATUS).setValue('提案済');
    postProposalReadyCard_(row[DM_COL.ID - 1], customer, url2);
    return htmlPage_('🔁 作り直しました',
      customer + ' の提案書を作り直し、Chatに新しいリンクを送りました。<br><br>' +
      '<a href="' + url2 + '" target="_blank">📄 提案書を開く</a>');
  }

  return htmlPage_('不明な操作', '対応していない操作です。');
}

/* ---------- 小物ヘルパー ---------- */

function postCard_(cardV2) {
  var url = prop_('CHAT_WEBHOOK_URL', false);
  if (!url) { Logger.log('[Chat未設定のためカード送信スキップ]'); return; }
  try {
    UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json; charset=UTF-8',
      payload: JSON.stringify({ cardsV2: [cardV2] }), muteHttpExceptions: true
    });
  } catch (e) { Logger.log('カード送信失敗: ' + e); }
}

function cardButton_(text, url) {
  return { text: text, onClick: { openLink: { url: url } } };
}

/** WebアプリのベースURL（自動取得→無ければWEBAPP_URLプロパティ） */
function webappUrl_() {
  try {
    var u = ScriptApp.getService().getUrl();
    if (u) return u;
  } catch (e) {}
  return prop_('WEBAPP_URL', false) || '';
}

function actionUrl_(action, caseId) {
  var base = webappUrl_();
  var token = prop_('APPROVE_TOKEN', false);
  if (!base || !token) return '';
  return base + '?action=' + encodeURIComponent(action) +
    '&case=' + encodeURIComponent(caseId) + '&token=' + encodeURIComponent(token);
}

function findDmRowById_(caseId) {
  if (!caseId) return null;
  var sh = openBook_().getSheetByName(CONFIG.SHEET_DM);
  if (!sh || sh.getLastRow() < 2) return null;
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, DM_HEADERS.length).getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[i][DM_COL.ID - 1] === caseId) {
      return { sheet: sh, rowNum: i + 2, values: values[i] };
    }
  }
  return null;
}

function trim_(s, n) {
  s = String(s || '');
  return s.length > n ? s.slice(0, n) + '…' : s;
}

/** ボタン押下後に表示する簡易ページ */
function htmlPage_(title, bodyHtml) {
  var html =
    '<!doctype html><html lang="ja"><head><meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<style>body{font-family:system-ui,-apple-system,"Noto Sans JP",sans-serif;background:#0f1620;color:#e7edf5;' +
    'display:flex;min-height:100vh;margin:0;align-items:center;justify-content:center;padding:24px}' +
    '.c{background:#182230;border:1px solid #28323f;border-radius:16px;padding:28px 26px;max-width:420px;text-align:center}' +
    'h1{font-size:20px;margin:0 0 10px}p,a{font-size:15px;line-height:1.7}a{color:#5c9dff}</style></head>' +
    '<body><div class="c"><h1>' + title + '</h1><p>' + bodyHtml + '</p>' +
    '<p style="opacity:.6;margin-top:16px">このタブは閉じて大丈夫です。</p></div></body></html>';
  return HtmlService.createHtmlOutput(html).setTitle(title);
}
