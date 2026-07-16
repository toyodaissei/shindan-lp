/**
 * ============================================================
 *  Setup.gs : 初期セットアップ・トリガー登録・メニュー・共通ログ
 * ------------------------------------------------------------
 *  使い方（初回）:
 *   1. GEMINI_API_KEY と REPORT_RECIPIENTS をスクリプトプロパティに設定
 *      （プロジェクトの設定 → スクリプトプロパティ）
 *   2. setup() を実行 … スプレッドシート/各Doc/Driveフォルダを自動作成しIDを保存
 *   3. installTriggers() を実行 … 定期実行を登録
 *   4. （Web連携するなら）「デプロイ → 新しいデプロイ → ウェブアプリ」でURL発行
 * ============================================================
 */

/** 初期セットアップ：必要なシート・Doc・フォルダを作りIDを保存 */
function setup() {
  var props = PropertiesService.getScriptProperties();

  // スプレッドシート
  var ssId = props.getProperty('SPREADSHEET_ID');
  var ss;
  if (ssId) { ss = SpreadsheetApp.openById(ssId); }
  else {
    ss = SpreadsheetApp.create('AI営業インテリジェンス スイート – データ');
    props.setProperty('SPREADSHEET_ID', ss.getId());
  }
  // 各シートをヘッダ付きで用意
  sheet_(CONFIG.SHEET_KNOWLEDGE, KN_HEADERS);
  sheet_(CONFIG.SHEET_DEALS, DEAL_HEADERS);
  sheet_(CONFIG.SHEET_DM, DM_HEADERS);
  sheet_(CONFIG.SHEET_DM_MSG, DM_MSG_HEADERS);
  sheet_(CONFIG.SHEET_REPORTLOG, ['日時', '種別', 'タイトル', '送信先']);
  // デフォルトの空シートを掃除
  var def = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);

  // 各種Doc
  ensureDoc_('NOTEBOOKLM_DOC_ID', '社内ナレッジ母艦（NotebookLM用）');
  ensureDoc_('KEIEI_REPORT_DOC_ID', '経営レポート');
  ensureDoc_('DM_REPORT_DOC_ID', 'DM営業レポート');

  // Driveフォルダ（DM録画の投入先 / 提案書の保存先）
  ensureFolder_('DM_RECORDINGS_FOLDER', 'DM営業_画面録画');
  ensureFolder_('PROPOSAL_FOLDER', 'AI生成_提案書');

  // 承認ボタンURLの改ざん防止トークン（無ければ生成）
  if (!props.getProperty('APPROVE_TOKEN')) {
    props.setProperty('APPROVE_TOKEN', Utilities.getUuid().replace(/-/g, ''));
  }
  // Webアプリ公開済みならURLを保存（承認ボタンのリンク先）
  try {
    var wu = ScriptApp.getService().getUrl();
    if (wu) props.setProperty('WEBAPP_URL', wu);
  } catch (e) {}

  // 未設定チェック
  var miss = [];
  if (!props.getProperty('GEMINI_API_KEY')) miss.push('GEMINI_API_KEY');
  if (!props.getProperty('REPORT_RECIPIENTS')) miss.push('REPORT_RECIPIENTS');

  var msg = 'setup 完了。\n\n' +
    'スプレッドシート: ' + ss.getUrl() + '\n' +
    'NotebookLM母艦Doc: ' + DocumentApp.openById(props.getProperty('NOTEBOOKLM_DOC_ID')).getUrl() + '\n' +
    '経営レポートDoc: ' + DocumentApp.openById(props.getProperty('KEIEI_REPORT_DOC_ID')).getUrl() + '\n' +
    'DMレポートDoc: ' + DocumentApp.openById(props.getProperty('DM_REPORT_DOC_ID')).getUrl() + '\n' +
    'DM録画フォルダ: ' + DriveApp.getFolderById(props.getProperty('DM_RECORDINGS_FOLDER')).getUrl() + '\n' +
    '提案書フォルダ: ' + DriveApp.getFolderById(props.getProperty('PROPOSAL_FOLDER')).getUrl() + '\n\n' +
    (miss.length ? '⚠ 未設定のプロパティ: ' + miss.join(', ') + '（手動で設定してください）'
                 : '✅ 必要なプロパティは揃っています。installTriggers() を実行してください。');
  Logger.log(msg);
  return msg;
}

function ensureDoc_(propKey, name) {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty(propKey)) return;
  var doc = DocumentApp.create(name);
  props.setProperty(propKey, doc.getId());
}

function ensureFolder_(propKey, name) {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty(propKey)) return;
  var folder = DriveApp.createFolder(name);
  props.setProperty(propKey, folder.getId());
}

/** 定期トリガーを登録（重複登録は自動で防止） */
function installTriggers() {
  removeTriggers_();
  // ぜんぶ自動の司令塔：15分毎に一括処理（誰もボタンを押さなくてよい）
  ScriptApp.newTrigger('autoRunAll').timeBased().everyMinutes(15).create();
  // レポート：毎週月曜 8:00（経営レポート & DM営業レポートをまとめて）
  ScriptApp.newTrigger('autoWeeklyReports').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  // スプレッドシートを開いた時に「AI営業スイート」メニューを出す（誰でも手動実行できる）
  try {
    ScriptApp.newTrigger('onOpen').forSpreadsheet(prop_('SPREADSHEET_ID', true)).onOpen().create();
  } catch (e) {
    Logger.log('メニュー用トリガー登録スキップ: ' + e);
  }
  Logger.log('トリガーを登録しました（autoRunAll:15分毎 / 週次レポート / シートのメニュー）');
}

function removeTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
}

/** スプレッドシートを開いた時のメニュー（手動実行しやすく） */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AI営業スイート')
    .addItem('① 初期セットアップ(setup)', 'setup')
    .addItem('② 自動運用をON(トリガー登録)', 'installTriggers')
    .addItem('③ Chat通知テスト', 'testChatNotify')
    .addSeparator()
    .addItem('▶ 今すぐ全部まとめて実行', 'autoRunAll')
    .addSeparator()
    .addItem('🅱 ナレッジ要約を今すぐ', 'processPending')
    .addItem('🅱 週次経営レポートを今すぐ', 'generateWeeklyReport')
    .addItem('🅰 商談→提案書を今すぐ', 'processDeals')
    .addItem('🆕 DM録画/スクショを取り込む', 'ingestDmRecordings')
    .addItem('🆕 DM提案を生成', 'proposeForPending')
    .addItem('🆕 DM営業レポートを今すぐ', 'generateDmReport')
    .addToUi();
}

/** 共通：レポート履歴に記録 */
function logReport_(kind, title, to) {
  var sh = sheet_(CONFIG.SHEET_REPORTLOG, ['日時', '種別', 'タイトル', '送信先']);
  sh.appendRow([
    Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm'),
    kind, title, to || ''
  ]);
}

/** APIキーだけ手早く設定したい時のヘルパー（値を直書きして1回実行） */
function setApiKey() {
  // setProp_('GEMINI_API_KEY', 'ここにキー');
  // setProp_('REPORT_RECIPIENTS', 'you@example.com, boss@example.com');
  throw new Error('この関数内のコメントを外し、キーとメールを入れてから実行してください。');
}
