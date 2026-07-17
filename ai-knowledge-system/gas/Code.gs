/**
 * ============================================================
 *  🅱 社内ナレッジ蓄積 → 経営レポート → NotebookLM 社内AI
 *     （動画①「入江が開発 / AIシステムに込めた経営哲学」の再現）
 * ------------------------------------------------------------
 *  流れ:
 *   1. 議事録/商談メモを「ナレッジ」シートに蓄積（doPost / addKnowledge）
 *   2. processPending() が新規行をAI要約 → 構造化（課題/示唆/次アクション等）
 *   3. NotebookLM用の母艦Docに追記（=会社を全部知ったAIの元データ）
 *   4. generateWeeklyReport() が週次で経営レポートを生成 → メール & Doc
 * ============================================================
 */

// ナレッジシートの列（1始まり）
var KN_COL = {
  ID: 1, TS: 2, TYPE: 3, WHO: 4, RAW: 5,
  SUMMARY: 6, ISSUES: 7, NEXT: 8, INSIGHT: 9, CATEGORY: 10, IMPORTANCE: 11,
  PROCESSED: 12, PROCESSED_AT: 13
};
var KN_HEADERS = ['id', '日時', '種別', '参加者/顧客', '生ログ(文字起こし)',
  '要約', '顧客の課題', 'ネクストアクション', '経営への示唆', 'カテゴリ', '重要度(1-5)',
  '処理済', '処理日時'];

/**
 * Webエンドポイント（n8n / Notta / フォーム等からPOSTで投入）
 * body 例:
 *   { "type":"商談", "who":"A社 田中様", "transcript":"...", "ts":"2026-07-15T10:00:00Z" }
 */
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    // 商談(提案書生成)ルートなら振り分け
    var routed = routePost_(e, body);
    if (routed) return jsonOut_(routed);
    var id = addKnowledge({
      type: body.type || body.種別 || '会議',
      who: body.who || body.参加者 || '',
      transcript: body.transcript || body.raw || body.text || '',
      ts: body.ts
    });
    return jsonOut_({ ok: true, id: id });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

/** ナレッジを1件追加（processed=false で積む） */
function addKnowledge(o) {
  var sh = sheet_(CONFIG.SHEET_KNOWLEDGE, KN_HEADERS);
  var id = newId_('kn');
  var ts = o.ts ? new Date(o.ts) : new Date();
  var row = [];
  row[KN_COL.ID - 1] = id;
  row[KN_COL.TS - 1] = Utilities.formatDate(ts, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm');
  row[KN_COL.TYPE - 1] = o.type || '会議';
  row[KN_COL.WHO - 1] = o.who || '';
  row[KN_COL.RAW - 1] = o.transcript || '';
  row[KN_COL.PROCESSED - 1] = false;
  sh.appendRow(row);
  return id;
}

/**
 * 未処理のナレッジをAIで構造化する（時間トリガー：15分毎など）
 */
function processPending() {
  var sh = sheet_(CONFIG.SHEET_KNOWLEDGE, KN_HEADERS);
  var last = sh.getLastRow();
  if (last < 2) return;
  var values = sh.getRange(2, 1, last - 1, KN_HEADERS.length).getValues();
  var processed = 0;

  for (var i = 0; i < values.length && processed < CONFIG.MAX_PROCESS_PER_RUN; i++) {
    var r = values[i];
    if (r[KN_COL.PROCESSED - 1] === true) continue;
    var raw = r[KN_COL.RAW - 1];
    if (!raw) continue;

    var prompt =
      'あなたは経営コンサル兼ナレッジマネージャーです。以下は社内の会議/商談の記録です。\n' +
      '内容を分析し、必ず次のJSONだけを返してください。\n' +
      '{"summary":"3〜5行の要約","issues":"顧客/現場の課題（箇条書きを改行で）",' +
      '"next":"ネクストアクション（改行区切り）","insight":"経営者が知るべき示唆・気づき",' +
      '"category":"カテゴリ(例: 営業/採用/顧客対応/プロダクト/資金/組織 のいずれか)",' +
      '"importance":1〜5の整数}\n\n' +
      '種別: ' + r[KN_COL.TYPE - 1] + ' / 相手: ' + r[KN_COL.WHO - 1] + '\n' +
      '記録:\n' + String(raw).slice(0, 12000);

    try {
      var o = callGeminiJson_(prompt, { temperature: 0.3 });
      var rowNum = i + 2;
      sh.getRange(rowNum, KN_COL.SUMMARY).setValue(o.summary || '');
      sh.getRange(rowNum, KN_COL.ISSUES).setValue(o.issues || '');
      sh.getRange(rowNum, KN_COL.NEXT).setValue(o.next || '');
      sh.getRange(rowNum, KN_COL.INSIGHT).setValue(o.insight || '');
      sh.getRange(rowNum, KN_COL.CATEGORY).setValue(o.category || '');
      sh.getRange(rowNum, KN_COL.IMPORTANCE).setValue(o.importance || '');
      sh.getRange(rowNum, KN_COL.PROCESSED).setValue(true);
      sh.getRange(rowNum, KN_COL.PROCESSED_AT).setValue(
        Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm'));
      appendToNotebookDoc_(r, o);
      processed++;
    } catch (err) {
      Logger.log('processPending 行' + (i + 2) + ' 失敗: ' + err);
    }
  }
  Logger.log('processPending: ' + processed + '件を処理');
  return processed;
}

/** NotebookLM用の母艦Docに1件追記（=会社を全部知ったAIの元データ） */
function appendToNotebookDoc_(r, o) {
  var docId = prop_('NOTEBOOKLM_DOC_ID', true);
  var doc = DocumentApp.openById(docId);
  var b = doc.getBody();
  b.appendParagraph('■ ' + r[KN_COL.TS - 1] + ' / ' + r[KN_COL.TYPE - 1] + ' / ' + r[KN_COL.WHO - 1])
    .setHeading(DocumentApp.ParagraphHeading.HEADING3);
  b.appendParagraph('要約: ' + (o.summary || ''));
  if (o.issues)  b.appendParagraph('課題: ' + o.issues);
  if (o.next)    b.appendParagraph('次アクション: ' + o.next);
  if (o.insight) b.appendParagraph('経営示唆: ' + o.insight);
  b.appendParagraph('カテゴリ: ' + (o.category || '') + ' / 重要度: ' + (o.importance || ''));
  b.appendHorizontalRule();
  doc.saveAndClose();
}

/**
 * 週次経営レポートを生成 → メール送付 + Docに追記（時間トリガー：週1）
 *  「デイリー/週報が経営者に上がってくる」を再現
 */
function generateWeeklyReport() { return generateReport_(CONFIG.REPORT_DAYS, '週次'); }
function generateDailyReport()  { return generateReport_(1, 'デイリー'); }

function generateReport_(days, label) {
  var sh = sheet_(CONFIG.SHEET_KNOWLEDGE, KN_HEADERS);
  var last = sh.getLastRow();
  if (last < 2) { Logger.log('ナレッジがありません'); return; }
  var values = sh.getRange(2, 1, last - 1, KN_HEADERS.length).getValues();

  var since = new Date();
  since.setDate(since.getDate() - days);
  var picked = values.filter(function (r) {
    if (r[KN_COL.PROCESSED - 1] !== true) return false;
    var t = new Date(r[KN_COL.TS - 1]);
    return t >= since;
  });
  if (!picked.length) { Logger.log('対象期間のナレッジがありません'); return; }

  var digest = picked.map(function (r) {
    return '・[' + r[KN_COL.CATEGORY - 1] + '/重要度' + r[KN_COL.IMPORTANCE - 1] + '] ' +
      r[KN_COL.WHO - 1] + '：' + r[KN_COL.SUMMARY - 1] +
      (r[KN_COL.INSIGHT - 1] ? '（示唆:' + r[KN_COL.INSIGHT - 1] + '）' : '');
  }).join('\n');

  var prompt =
    'あなたは経営者の右腕（参謀）です。以下は直近' + days + '日間の社内ナレッジ（' + picked.length + '件）の要約リストです。\n' +
    'これを経営者向けの' + label + 'レポートにまとめてください。冗長にせず、意思決定に効く形で。\n' +
    '構成:\n' +
    '1. 今週のハイライト（3点）\n' +
    '2. 見えてきた顧客/市場の課題\n' +
    '3. 社内の弱み・改善提案（率直に）\n' +
    '4. 経営者が今すぐ判断すべきこと（優先順に）\n' +
    '5. 数字で見るべきKPI提案\n\n' +
    'ナレッジ:\n' + digest.slice(0, 14000);

  var report = callGemini_(prompt, { temperature: 0.5, maxTokens: 4096 });

  // Docに追記
  var docId = prop_('KEIEI_REPORT_DOC_ID', true);
  var doc = DocumentApp.openById(docId);
  var b = doc.getBody();
  var title = '【' + label + '経営レポート】' +
    Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
  b.insertParagraph(0, '').appendHorizontalRule();
  var h = b.insertParagraph(0, title);
  h.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  b.insertParagraph(1, report);
  doc.saveAndClose();

  // メール送付（配信先シート＋REPORT_RECIPIENTS）
  var to = recipients_();
  if (to) {
    MailApp.sendEmail({
      to: to,
      subject: title,
      body: report + '\n\n----\nレポートDoc: ' + doc.getUrl() +
        '\nNotebookLM母艦Doc: ' + DocumentApp.openById(prop_('NOTEBOOKLM_DOC_ID', true)).getUrl()
    });
  }

  logReport_(label + '経営レポート', title, to);
  Logger.log(label + 'レポートを送信しました → ' + to);
}

/** doGet: 承認ボタンの処理 or 動作確認 */
function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    return handleApproveAction_(e);   // Chat承認ボタンからのリンク
  }
  return jsonOut_({ ok: true, service: 'AI営業インテリジェンス スイート', time: new Date().toISOString() });
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
