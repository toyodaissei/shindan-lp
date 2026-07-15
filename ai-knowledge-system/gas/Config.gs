/**
 * ============================================================
 *  AI営業インテリジェンス スイート  (Google Apps Script)
 * ------------------------------------------------------------
 *  3つのシステムを1プロジェクトで動かします。
 *
 *   🅱 社内ナレッジ蓄積 → 経営レポート → NotebookLM 社内AI   … Code.gs
 *       （動画①「入江が開発 / AIシステムに込めた経営哲学」の再現）
 *   🅰 商談 → 提案書(スライド) 自動生成 → 送付               … Proposal.gs
 *       （動画②「Nottaで録音→GASで5分後にパワポ提案書」の再現）
 *   🆕 営業DM 最適解提案システム（画面録画をAIが視聴）        … DmCoach.gs
 *       DM営業の画面録画をDriveに溜め → 各案件を解析 →
 *       「マニュアル型」＋「提案型」の2つで最適解を提案 → レポート
 * ------------------------------------------------------------
 *  Config.gs : 設定・定数まとめ
 *
 *  ● 秘密情報（APIキー・各種ID）は「スクリプト プロパティ」に入れます。
 *    Apps Script エディタ →「プロジェクトの設定」→「スクリプト プロパティ」
 *    もしくは Setup.gs の setup() を1回実行すると、
 *    スプレッドシート/Doc/Driveフォルダを自動作成し、IDを埋めてくれます。
 *
 *  必要なプロパティ（setup() 実行後に手動で埋めるのは ★ のみ）:
 *    ★ GEMINI_API_KEY        … Google AI Studio で無料発行
 *                              https://aistudio.google.com/apikey
 *    ★ REPORT_RECIPIENTS     … レポート送付先メール（カンマ区切り）
 *      SPREADSHEET_ID        … ナレッジ等を貯めるスプレッドシート
 *      NOTEBOOKLM_DOC_ID     … NotebookLM に食わせる母艦 Google Doc
 *      KEIEI_REPORT_DOC_ID   … 経営レポート出力先 Google Doc
 *      DM_RECORDINGS_FOLDER  … DM営業の画面録画を入れる Drive フォルダ
 *      DM_REPORT_DOC_ID      … DM分析レポート出力先 Google Doc
 *      PROPOSAL_FOLDER       … 生成した提案書(スライド)の保存先フォルダ
 * ============================================================
 */

var CONFIG = {
  // ---- シート名（1つのスプレッドシート内のタブ） ----
  SHEET_KNOWLEDGE: 'ナレッジ',       // 🅱 議事録・商談メモを1行ずつ蓄積
  SHEET_DEALS:     '商談',           // 🅰 商談ごとの提案書生成キュー
  SHEET_DM:        'DM案件',         // 🆕 DM営業の案件と録画解析結果
  SHEET_DM_MSG:    'DMメッセージ',    // 🆕 録画から抽出した個々のDM文面（学習用）
  SHEET_REPORTLOG: 'レポート履歴',    // 全システム共通のレポート履歴

  // ---- 使用するLLM (Gemini) ----
  //  Workspace と相性が良く、無料枠があり、動画も直接解析できるため既定は Gemini。
  //  ※ モデル名は時期により提供状況が変わります（2.0-flashが無料枠0/2.5-flashが404 等）。
  //    実際に使えるモデルは AI Studio のモデル一覧で確認し、ここを差し替えてください。
  //    2026/07時点で動作確認できたモデル: gemini-3-flash-preview
  GEMINI_MODEL:       'gemini-3-flash-preview',   // 通常の要約・生成
  GEMINI_MODEL_VIDEO: 'gemini-3-flash-preview',   // 画面録画の視聴解析（マルチモーダル）
  GEMINI_API_BASE:    'https://generativelanguage.googleapis.com',

  //  思考(thinking)型モデル対策：
  //   思考トークンが出力上限を食い尽くし回答が空になるのを防ぐため思考をスキップ。
  //   0=思考オフ。思考型でない旧モデルに戻す場合は null にすると本設定を送りません。
  GEMINI_THINKING_BUDGET: 0,

  // ---- 挙動 ----
  MAX_PROCESS_PER_RUN: 20,   // 1回の自動処理で捌く最大件数（API/時間の節約）
  REPORT_DAYS:         7,    // 週次レポートで集計する日数
  TIMEZONE:            'Asia/Tokyo'
};

/** スクリプトプロパティを安全に取得 */
function prop_(key, required) {
  var v = PropertiesService.getScriptProperties().getProperty(key);
  if (required && !v) {
    throw new Error('スクリプトプロパティ「' + key + '」が未設定です。setup() を実行するか、手動で設定してください。');
  }
  return v;
}

function setProp_(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, value);
}

/** 共通：設定済みスプレッドシートを開く */
function openBook_() {
  var id = prop_('SPREADSHEET_ID', true);
  return SpreadsheetApp.openById(id);
}

/** 共通：シートを名前で取得（無ければ作る） */
function sheet_(name, headers) {
  var ss = openBook_();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers && headers.length) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

/** 共通：連番ID（タイムスタンプ + ランダム） */
function newId_(prefix) {
  return (prefix || 'id') + '-' +
    Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd-HHmmss') + '-' +
    Math.floor(Math.random() * 1000);
}
