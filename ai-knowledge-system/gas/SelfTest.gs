/**
 * ============================================================
 *  SelfTest.gs : 初回セットアップ後の動作確認セット
 * ------------------------------------------------------------
 *  実行順のおすすめ:
 *   1. checkConfig()        … プロパティの過不足 & Geminiキー疎通を確認
 *   2. smokeTestKnowledge() … 🅱 ナレッジ→要約→NotebookLM Doc→経営レポート を一気通貫
 *   3. demoDealProposal()   … 🅰 サンプル商談→提案書スライドを生成
 *   4. demoDmProposal()     … 🆕 録画なし(テキスト)でDMの「2つの型」提案Docを生成
 *
 *  ※ いずれも実データではなく「サンプル」を入れて動きを確認するものです。
 *    テスト行はシート上で「(TEST)」と分かるようにしています。
 * ============================================================
 */

/** ① 設定チェック + Gemini疎通 */
function checkConfig() {
  var need = ['GEMINI_API_KEY', 'REPORT_RECIPIENTS', 'SPREADSHEET_ID',
    'NOTEBOOKLM_DOC_ID', 'KEIEI_REPORT_DOC_ID', 'DM_REPORT_DOC_ID',
    'DM_RECORDINGS_FOLDER', 'PROPOSAL_FOLDER'];
  var lines = ['=== 設定チェック ==='];
  var missing = [];
  need.forEach(function (k) {
    var v = PropertiesService.getScriptProperties().getProperty(k);
    lines.push((v ? '✅ ' : '❌ ') + k + (v ? '' : '（未設定）'));
    if (!v) missing.push(k);
  });

  if (missing.indexOf('SPREADSHEET_ID') >= 0) {
    lines.push('\n→ SPREADSHEET_ID 等が無い場合は先に setup() を実行してください。');
  }

  // Gemini 疎通
  if (PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY')) {
    try {
      var ping = callGemini_('「疎通OK」とだけ返して。', { maxTokens: 20, temperature: 0 });
      lines.push('\n✅ Gemini疎通: ' + ping);
    } catch (e) {
      lines.push('\n❌ Gemini疎通失敗: ' + e);
    }
  } else {
    lines.push('\n❌ GEMINI_API_KEY 未設定のため疎通テスト不可');
  }
  var msg = lines.join('\n');
  Logger.log(msg);
  return msg;
}

/** ② 🅱 ナレッジの一気通貫テスト */
function smokeTestKnowledge() {
  var sample =
    '(TEST) A社との定例MTG。担当の田中様より「現行ツールの操作が複雑で現場が使いこなせていない」' +
    'との課題共有。競合B社の乗り換え提案も来ている様子。当社は運用サポート強化を提案し、' +
    '来週トライアル導入のデモを実施予定。予算感は月20万まで。決裁は部長の承認が必要。';
  var id = addKnowledge({ type: '商談', who: '(TEST) A社 田中様', transcript: sample });
  Logger.log('サンプルナレッジ投入: ' + id);

  processPending();               // AI要約 + NotebookLM Docへ追記
  generateReport_(30, '(TEST)週次'); // 直近30日で1件でもレポートを作る

  var out = 'smokeTestKnowledge 完了。\n' +
    'ナレッジシート / NotebookLM母艦Doc / 経営レポートDoc & 受信メールを確認してください。\n' +
    'レポートDoc: ' + DocumentApp.openById(prop_('KEIEI_REPORT_DOC_ID', true)).getUrl();
  Logger.log(out);
  return out;
}

/** ③ 🅰 商談→提案書スライドのテスト */
function demoDealProposal() {
  var sample =
    '(TEST) C社の新規商談。ECの売上が伸び悩み、特にカゴ落ちが多いのが悩み。' +
    'メール施策は手作業で属人的。CVRを今の1.2%から2%に上げたい。' +
    '予算は初期50万・月10万程度。導入は3ヶ月以内希望。担当は鈴木様(マーケ責任者)。';
  var id = addDeal({ customer: '(TEST) C社 鈴木様', transcript: sample });
  Logger.log('サンプル商談投入: ' + id);
  processDeals();

  var sh = sheet_(CONFIG.SHEET_DEALS, DEAL_HEADERS);
  var last = sh.getLastRow();
  var url = sh.getRange(last, DEAL_COL.SLIDE_URL).getValue();
  var out = 'demoDealProposal 完了。\n提案書スライド: ' + url;
  Logger.log(out);
  return out;
}

/** ④ 🆕 録画なし(テキスト)でDM「2つの型」提案を生成（エージェント開拓シーン） */
function demoDmProposal() {
  // 本来は Gemini が録画から作る解析結果を、ここでは手打ちで用意
  var fakeAnalysis = {
    customer: '(TEST) 就活系インフルエンサー E様',
    product: '27卒面談送客案件',
    platform: 'InstagramDM',
    summary: '(TEST) 就活系の発信をしているインフルエンサーに、27卒学生の面談送客案件を' +
      'エージェントとして一緒にやりませんかとDMで打診。案件概要は伝えたが、報酬条件が曖昧で返信が停滞。',
    technique: '冒頭で発信への共感→案件概要の提示→「一緒にやりませんか」の打診、という型。ただし報酬・稼働イメージが不明瞭。',
    reaction: '既読はつき1返信あるが、条件面が見えず様子見。',
    outcome: '返信',
    sent: 15, replies: 5,
    good: '相手の発信内容に触れて入った点は好反応。',
    bad: '報酬(送客単価)や稼働負荷、実績が曖昧で“自分がやる理由”が弱い。面談への導線が無い。',
    messages: [
      { step: 1, message: 'はじめまして!27卒向けの発信いつも拝見してます。実は学生の面談送客の案件をご一緒できないかと思いご連絡しました。', intent: '共感+用件提示', reaction: '既読→返信', effect: '良' },
      { step: 2, message: '面談1件あたり〇〇円で、審査もこちらで巻き取ります。', intent: '条件提示', reaction: '様子見', effect: '普' },
      { step: 3, message: 'よければ詳細お送りしますね!', intent: 'CTA', reaction: '停滞', effect: '悪' }
    ]
  };

  var sh = sheet_(CONFIG.SHEET_DM, DM_HEADERS);
  var fakeFile = { getId: function () { return 'TEST-' + newId_('f'); }, getName: function () { return '(TEST) DMサンプル.mp4'; } };
  writeDmCase_(sh, fakeFile, fakeAnalysis);
  appendDmMessages_(fakeAnalysis);
  Logger.log('サンプルDM案件を投入。提案を生成します…');

  proposeForPending();

  var last = sh.getLastRow();
  var url = sh.getRange(last, DM_COL.PROPOSAL_URL).getValue();
  var out = 'demoDmProposal 完了。\n' +
    '「マニュアル型 + 提案型」の提案書Doc: ' + url + '\n' +
    '（録画を使う本番は、Driveの「DM営業_画面録画」フォルダに動画を入れて ingestDmRecordings() を実行）';
  Logger.log(out);
  return out;
}

/** おまけ: テスト行をまとめて掃除したい時（(TEST)を含む行を削除） */
function cleanupTestRows() {
  [CONFIG.SHEET_KNOWLEDGE, CONFIG.SHEET_DEALS, CONFIG.SHEET_DM, CONFIG.SHEET_DM_MSG].forEach(function (name) {
    var sh = openBook_().getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return;
    var vals = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    for (var i = vals.length - 1; i >= 0; i--) {
      if (vals[i].join('|').indexOf('(TEST)') >= 0) sh.deleteRow(i + 2);
    }
  });
  Logger.log('(TEST)行を削除しました');
}
