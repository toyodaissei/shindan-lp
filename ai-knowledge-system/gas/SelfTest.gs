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
    objective: '代理店/協業パートナーの開拓（相手にハブになってもらい学生を送客してもらう）',
    customer: '(TEST) 学生送客をしている個人事業者 F様',
    product: '学生送客の代理店開拓',
    platform: 'ThreadsDM',
    summary: '(TEST) Threadsで学生送客をしている個人に、定型文で「ぜひ協業させてください」とDMを送り、' +
      'アポを取ってミートに繋げる代理店開拓。相手にハブになってもらい学生を送客してもらうのが目的。',
    technique: '定型文で協業を打診 → 返信が来たらアポ打診 → 日程調整してミート設定、という開拓の型。',
    reaction: '返信あり。協業自体には前向きだが、条件と実績を知りたい様子。',
    outcome: '返信',
    sent: 20, replies: 7,
    good: '定型文でも「協業」という明確な用件で返信率が出ている。',
    bad: '協業のメリット・条件・実績が薄く、面談化までの導線が弱い。',
    messages: [
      { step: 1, message: 'はじめまして!学生送客をされていると拝見しました。ぜひ一度協業させていただけないかと思いご連絡しました。', intent: '協業の打診(定型文)', reaction: '返信あり', effect: '良' },
      { step: 2, message: 'よければ15分ほどお話しできませんか?こちらの案件と条件をご説明します。', intent: 'アポ打診', reaction: '日程調整へ', effect: '良' },
      { step: 3, message: '直近だと〇/〇か〇/〇はいかがでしょう?', intent: '日程提示', reaction: '様子見', effect: '普' }
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
