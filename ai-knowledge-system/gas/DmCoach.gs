/**
 * ============================================================
 *  🆕 営業DM 最適解提案システム（オリジナル）
 * ------------------------------------------------------------
 *  コンセプト:
 *   DM営業をしている「画面録画」を Google ドライブに溜め込むと、
 *   AI(Gemini)がその録画を“視聴”して手法を抽出・学習し、
 *   各案件に対して最適な営業DM手法を「2つの型」で提案する。
 *   さらに、返信率・転換率などの数値レポートも出る。
 *
 *  提案の2つの型:
 *   ① マニュアル型 … 新人が1→10まで見てそのまま真似できる手順書
 *                    （送る順番・そのままコピペできる文面・タイミング）
 *   ② 提案型      … そこからさらにアイデアが湧くきっかけ
 *                    （切り口の提案＋想定 返信率／転換率などの数値）
 *
 *  流れ:
 *   1. Driveの「DM営業録画」フォルダに画面録画(.mp4/.mov等)を入れる
 *   2. ingestDmRecordings() が新規録画をAIに視聴させ「DM案件」シートに構造化
 *      （送った文面の流れ・使った技術・顧客反応・結果・良かった点/改善点）
 *   3. proposeForPending() が各案件に「マニュアル型＋提案型」の提案書(Doc)を生成
 *      過去の高成績パターンも学習データとして反映
 *   4. generateDmReport() が返信率・転換率などの数値レポートを生成
 * ============================================================
 */

var DM_COL = {
  ID: 1, TS: 2, FILE_ID: 3, FILE_NAME: 4, CUSTOMER: 5, PRODUCT: 6, PLATFORM: 7,
  SUMMARY: 8, TECHNIQUE: 9, REACTION: 10, OUTCOME: 11,
  SENT: 12, REPLIES: 13, APPTS: 14, DEALS: 15, REPLY_RATE: 16, CONV_RATE: 17,
  GOOD: 18, BAD: 19, PROPOSAL_URL: 20, STATUS: 21, PROCESSED_AT: 22
};
var DM_HEADERS = ['id', '日時', 'ファイルID', 'ファイル名', '相手(エージェント)/案件', '案件/商材', 'チャネル',
  '録画の要約', '使った営業手法', '相手の反応', '結果(未返信/返信/面談/提携)',
  '送信数', '返信数', '面談数', '提携数', '返信率', '提携率',
  '良かった点', '改善点', '提案書URL', 'ステータス', '処理日時'];

var DM_MSG_COL = { ID: 1, CASE_ID: 2, CUSTOMER: 3, PLATFORM: 4, STEP: 5, MESSAGE: 6, INTENT: 7, REACTION: 8, EFFECT: 9 };
var DM_MSG_HEADERS = ['id', '案件ID', '顧客', 'プラットフォーム', 'ステップ', '送信メッセージ', '狙い/技術', '顧客反応', '効果(良/普/悪)'];

/**
 * Driveフォルダの新規録画をAIに視聴させて「DM案件」に取り込む
 *  時間トリガー: 30分〜1時間毎 など
 */
function ingestDmRecordings() {
  var folderId = prop_('DM_RECORDINGS_FOLDER', true);
  var folder = DriveApp.getFolderById(folderId);
  var sh = sheet_(CONFIG.SHEET_DM, DM_HEADERS);
  var known = getKnownFileIds_(sh);

  var files = folder.getFiles();
  var processed = 0;
  while (files.hasNext() && processed < CONFIG.MAX_PROCESS_PER_RUN) {
    var file = files.next();
    var mime = file.getMimeType();
    // 動画(画面録画)と画像(スクショ)の両方に対応
    if (mime.indexOf('video') !== 0 && mime.indexOf('image') !== 0) continue;
    if (known[file.getId()]) continue;               // 取込済みはスキップ

    try {
      var analysis = analyzeDmRecording_(file);
      writeDmCase_(sh, file, analysis);
      appendDmMessages_(analysis);
      processed++;
    } catch (err) {
      Logger.log('録画解析 失敗 ' + file.getName() + ': ' + err);
      // 失敗もエラー行として記録（無限リトライ防止）
      var row = [];
      row[DM_COL.ID - 1] = newId_('dm');
      row[DM_COL.TS - 1] = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm');
      row[DM_COL.FILE_ID - 1] = file.getId();
      row[DM_COL.FILE_NAME - 1] = file.getName();
      row[DM_COL.STATUS - 1] = '解析失敗: ' + String(err).slice(0, 80);
      sh.appendRow(row);
    }
  }
  Logger.log('ingestDmRecordings: ' + processed + '件の録画/スクショを解析');
  return processed;
}

/** Geminiに画面録画(動画)またはスクショ(画像)を解析させDM営業(エージェント開拓)を構造化 */
function analyzeDmRecording_(file) {
  var mime = file.getMimeType();
  var prompt =
    'これは「代理店/エージェント開拓」の営業担当者が、SNSのDM等で相手(集客・送客してくれるエージェント候補)に' +
    '案件を持ちかけている「画面録画またはスクリーンショット」です。画面に映るチャット/DMのやり取りを読み取り、営業手法を分析してください。' +
    '必ず次のJSONのみ返す:\n' +
    '{' +
    '"customer":"相手(エージェント)名/案件名（分かれば。無ければ推定や空文字）",' +
    '"product":"持ちかけている案件/商材(例:27卒面談送客案件 など)",' +
    '"platform":"チャネル(Instagram/InstagramDM/Threads/ThreadsDM/X/LINE/公式LINE/Facebook/その他)",' +
    '"summary":"録画で何が行われたかの要約(4〜6行)",' +
    '"technique":"使われた営業手法・トークの型(改行区切りで具体的に)",' +
    '"reaction":"相手の反応・温度感",' +
    '"outcome":"未返信 / 返信 / 面談 / 提携 のいずれか(判断できなければ最も近いもの。提携=成約)",' +
    '"sent":送ったメッセージ数(整数),"replies":相手の返信数(整数),' +
    '"good":"良かった点(なぜ効いたか)",' +
    '"bad":"改善点(もっとこうすれば返信/面談/提携が増えるか)",' +
    '"messages":[{"step":1,"message":"実際に送った文面(できるだけ原文)","intent":"その一手の狙い/技術","reaction":"相手の反応","effect":"良/普/悪"}]' +
    '}\n' +
    '文面は新人が真似できるよう、できるだけ原文に近い形で書き出すこと。';
  if (mime.indexOf('image') === 0) {
    return analyzeImage_(file.getBlob(), prompt, { json: true, maxTokens: 8192 });
  }
  return analyzeVideo_(file.getBlob(), prompt, { json: true, maxTokens: 8192 });
}

function writeDmCase_(sh, file, a) {
  var id = newId_('dm');
  var sent = Number(a.sent) || 0;
  var replies = Number(a.replies) || 0;
  var deals = (a.outcome === '提携' || a.outcome === '成約') ? 1 : 0;
  var row = [];
  row[DM_COL.ID - 1] = id;
  row[DM_COL.TS - 1] = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm');
  row[DM_COL.FILE_ID - 1] = file.getId();
  row[DM_COL.FILE_NAME - 1] = file.getName();
  row[DM_COL.CUSTOMER - 1] = a.customer || file.getName();
  row[DM_COL.PRODUCT - 1] = a.product || '';
  row[DM_COL.PLATFORM - 1] = a.platform || '';
  row[DM_COL.SUMMARY - 1] = a.summary || '';
  row[DM_COL.TECHNIQUE - 1] = a.technique || '';
  row[DM_COL.REACTION - 1] = a.reaction || '';
  row[DM_COL.OUTCOME - 1] = a.outcome || '';
  row[DM_COL.SENT - 1] = sent;
  row[DM_COL.REPLIES - 1] = replies;
  row[DM_COL.APPTS - 1] = (a.outcome === '面談' || a.outcome === 'アポ' || a.outcome === '提携' || a.outcome === '成約') ? 1 : 0;
  row[DM_COL.DEALS - 1] = deals;
  row[DM_COL.REPLY_RATE - 1] = sent ? Math.round(replies / sent * 100) + '%' : '';
  row[DM_COL.CONV_RATE - 1] = sent ? Math.round(deals / sent * 100) + '%' : '';
  row[DM_COL.GOOD - 1] = a.good || '';
  row[DM_COL.BAD - 1] = a.bad || '';
  row[DM_COL.STATUS - 1] = '解析済(提案待ち)';
  row[DM_COL.PROCESSED_AT - 1] = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm');
  sh.appendRow(row);
  a._caseId = id;
  a._caseCustomer = row[DM_COL.CUSTOMER - 1];
}

/** 抽出した個々のDM文面を学習コーパスとして蓄積 */
function appendDmMessages_(a) {
  if (!a.messages || !a.messages.length) return;
  var sh = sheet_(CONFIG.SHEET_DM_MSG, DM_MSG_HEADERS);
  var rows = a.messages.map(function (m) {
    var r = [];
    r[DM_MSG_COL.ID - 1] = newId_('msg');
    r[DM_MSG_COL.CASE_ID - 1] = a._caseId || '';
    r[DM_MSG_COL.CUSTOMER - 1] = a._caseCustomer || '';
    r[DM_MSG_COL.PLATFORM - 1] = a.platform || '';
    r[DM_MSG_COL.STEP - 1] = m.step || '';
    r[DM_MSG_COL.MESSAGE - 1] = m.message || '';
    r[DM_MSG_COL.INTENT - 1] = m.intent || '';
    r[DM_MSG_COL.REACTION - 1] = m.reaction || '';
    r[DM_MSG_COL.EFFECT - 1] = m.effect || '';
    return r;
  });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, DM_MSG_HEADERS.length).setValues(rows);
}

/** 提案待ちの案件に「マニュアル型＋提案型」の提案書を生成（時間トリガー可） */
function proposeForPending() {
  var sh = sheet_(CONFIG.SHEET_DM, DM_HEADERS);
  var last = sh.getLastRow();
  if (last < 2) return;
  var values = sh.getRange(2, 1, last - 1, DM_HEADERS.length).getValues();
  var made = [];
  for (var i = 0; i < values.length; i++) {
    if (values[i][DM_COL.STATUS - 1] !== '解析済(提案待ち)') continue;
    try {
      var url = proposeDmStrategy_(values[i]);
      sh.getRange(i + 2, DM_COL.PROPOSAL_URL).setValue(url);
      sh.getRange(i + 2, DM_COL.STATUS).setValue('提案済');
      made.push({ customer: values[i][DM_COL.CUSTOMER - 1], url: url });
    } catch (err) {
      Logger.log('提案生成 失敗 行' + (i + 2) + ': ' + err);
    }
  }
  return made;
}

/**
 * 提案JSONの必須構造をスキーマで強制（Geminiが変数一覧やtalkScriptを落とさないように）
 */
var DM_PROPOSAL_SCHEMA = {
  type: 'object',
  properties: {
    variables: {
      type: 'array',
      items: {
        type: 'object',
        properties: { key: { type: 'string' }, desc: { type: 'string' }, example: { type: 'string' } },
        required: ['key', 'desc']
      }
    },
    manual: {
      type: 'object',
      properties: {
        goal: { type: 'string' },
        steps: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              no: { type: 'integer' }, action: { type: 'string' }, channel: { type: 'string' },
              script: { type: 'string' }, timing: { type: 'string' }, point: { type: 'string' }
            },
            required: ['no', 'action', 'channel', 'script', 'timing', 'point']
          }
        },
        talkScript: {
          type: 'object',
          properties: {
            opening: { type: 'string' },
            hearing: { type: 'array', items: { type: 'string' } },
            proposal: { type: 'string' },
            objection: {
              type: 'array',
              items: {
                type: 'object',
                properties: { 'if': { type: 'string' }, say: { type: 'string' } },
                required: ['if', 'say']
              }
            },
            closing: { type: 'string' }
          },
          required: ['opening', 'hearing', 'proposal', 'objection', 'closing']
        },
        ng: { type: 'array', items: { type: 'string' } }
      },
      required: ['goal', 'steps', 'talkScript', 'ng']
    },
    ideas: {
      type: 'object',
      properties: {
        angles: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              title: { type: 'string' }, why: { type: 'string' }, 'try': { type: 'string' },
              expReplyRate: { type: 'string' }, expMeetingRate: { type: 'string' },
              expPartnerRate: { type: 'string' }, basis: { type: 'string' }
            },
            required: ['title', 'why', 'try', 'expReplyRate', 'expMeetingRate', 'expPartnerRate', 'basis']
          }
        },
        experiments: { type: 'array', items: { type: 'string' } }
      },
      required: ['angles', 'experiments']
    }
  },
  required: ['variables', 'manual', 'ideas']
};

/** 1案件について、過去の高成績パターンを踏まえ「2つの型」で提案 → Docに書き出す */
function proposeDmStrategy_(caseRow) {
  var bestPatterns = collectBestPatterns_();

  var prompt =
    'あなたは「代理店/エージェント開拓」に強いトップセールス兼営業コーチです。\n' +
    '私たちは広告主と代理店の間に立つ代理店で、広告主から受けた案件(例:27卒面談送客案件)を、\n' +
    'SNSのDM等でエージェント(集客・送客してくれる個人/法人パートナー)に持ちかけ、提携してもらうのが営業です。\n' +
    '目的は「エージェント開拓の営業ナレッジを再現性ある形にし、新人でも回せるトークスクリプト/DMを提示する」こと。\n' +
    '以下の対象案件に対する最適解を、過去の高成績パターンも学習材料に作ってください。\n\n' +
    '=== 対象案件 ===\n' +
    '相手(エージェント)/案件: ' + caseRow[DM_COL.CUSTOMER - 1] + '\n' +
    '案件/商材: ' + caseRow[DM_COL.PRODUCT - 1] + '\n' +
    '主なチャネル: ' + caseRow[DM_COL.PLATFORM - 1] +
      '（想定チャネル: Instagram/InstagramDM/Threads/ThreadsDM/X/LINE/公式LINE/Facebook）\n' +
    '現状の要約: ' + caseRow[DM_COL.SUMMARY - 1] + '\n' +
    '使った手法: ' + caseRow[DM_COL.TECHNIQUE - 1] + '\n' +
    '相手の反応: ' + caseRow[DM_COL.REACTION - 1] + ' / 結果: ' + caseRow[DM_COL.OUTCOME - 1] + '\n' +
    '改善点メモ: ' + caseRow[DM_COL.BAD - 1] + '\n\n' +
    '=== 過去の高成績パターン(学習データ) ===\n' + bestPatterns + '\n\n' +
    '必ず次のJSONのみ返す:\n' +
    '{' +
    '"variables":[{"key":"{{相手の名前}}","desc":"何を入れるか","example":"例"}],' +
    '"manual":{' +
      '"goal":"このマニュアルで到達するゴール(例:面談アポ獲得)",' +
      '"steps":[{"no":1,"action":"やること(新人が迷わない粒度)","channel":"どのチャネルで送るか","script":"そのままコピペで送れる文面。差し込み変数は必ず {{変数名}} 形式で埋め込む","timing":"送るタイミング/条件","point":"つまづき注意"}],' +
      '"talkScript":{"opening":"面談/通話の掴み(変数可)","hearing":["ヒアリングで聞く質問"],"proposal":"案件の提案トーク","objection":[{"if":"よくある断り/懸念","say":"切り返しトーク"}],"closing":"クロージング(次アクション確定)"},' +
      '"ng":["やってはいけないこと"]' +
    '},' +
    '"ideas":{' +
      '"angles":[{"title":"切り口/アイデア名","why":"なぜ効くかの仮説","try":"具体的な試し方","expReplyRate":"想定返信率(例:25〜35%)","expMeetingRate":"想定 面談化率(例:8〜12%)","expPartnerRate":"想定 提携(成約)率(例:3〜5%)","basis":"その数値の根拠(過去データや一般値)"}],' +
      '"experiments":["A/Bで試すべき比較案(数値で検証できる形)"]' +
    '}' +
    '}\n' +
    '重要:\n' +
    '・文面テンプレは必ず {{変数名}} の差し込み形式にする。数値や固有名詞の空欄に「◯◯」等の伏字を絶対に使わず、' +
    '必ず {{想定報酬額}} {{想定月収}} {{フォロワー数}} {{実績数値}} のような変数にする。使った変数は漏れなく "variables" に列挙する。\n' +
    '・steps は各チャネル特性(Instagram/ThreadsのDMは短く軽快、公式LINEは段階的、Xは簡潔、Facebookはやや丁寧)を踏まえ、6〜10個。\n' +
    '・"talkScript" は面談/通話でそのまま読めるレベルで具体的に。\n' +
    '・"talkScript.objection" には、エージェント開拓で頻出する次の懸念への切り返しを必ず含める:' +
    '「報酬条件が曖昧で自分がやる理由が分からない」「稼働・作業負荷が見えず不安」。' +
    'これらに加え、対象案件の状況に応じた懸念も足す。\n' +
    '・"ideas" は必ず 想定返信率・面談化率・提携率 の3数値を添え、さらにアイデアが広がる“きっかけ”にする。\n' +
    '・"basis"(根拠)は、上記「過去の高成績パターン」に実データがあればそれを引用し、無い場合は「一般値/仮説」と明記する。';

  var o = callGeminiJson_(prompt, { temperature: 0.6, maxTokens: 8192, schema: DM_PROPOSAL_SCHEMA });

  // 保険：必須ブロックが欠けたら一度だけ強めに再生成
  if (!o || !o.variables || !o.manual || !o.manual.talkScript || !o.ideas) {
    o = callGeminiJson_(prompt +
      '\n\n【厳守】前回 variables / manual.talkScript / ideas のいずれかが欠落しました。' +
      '3ブロックすべてを必ず埋めて返してください。',
      { temperature: 0.5, maxTokens: 8192, schema: DM_PROPOSAL_SCHEMA });
  }

  return writeDmProposalDoc_(caseRow[DM_COL.CUSTOMER - 1], caseRow[DM_COL.PRODUCT - 1], caseRow[DM_COL.PLATFORM - 1], o);
}

/** 提案JSON → 見やすいGoogle Docに整形して書き出し、URLを返す */
function writeDmProposalDoc_(customer, product, platform, o) {
  var folderId = prop_('PROPOSAL_FOLDER', false);
  var name = 'エージェント開拓提案_' + (customer || '案件') + '_' +
    Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd-HHmm');
  var doc = DocumentApp.create(name);
  var b = doc.getBody();

  b.appendParagraph('エージェント開拓 最適解 提案書').setHeading(DocumentApp.ParagraphHeading.TITLE);
  b.appendParagraph('相手/案件: ' + (customer || '') + '　/　案件・商材: ' + (product || '') +
    '　/　主なチャネル: ' + (platform || ''));
  b.appendHorizontalRule();

  // 差し込み変数一覧（新人がまず埋めるもの）
  if (o.variables && o.variables.length) {
    b.appendParagraph('◇ 差し込み変数一覧（送信前にここを埋める）')
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);
    o.variables.forEach(function (v) {
      b.appendListItem((v.key || '') + ' … ' + (v.desc || '') +
        (v.example ? '（例: ' + v.example + '）' : ''));
    });
  }

  // ① マニュアル型
  b.appendParagraph('① マニュアル型（新人が1→10でそのまま真似できる手順書）')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  var m = o.manual || {};
  if (m.goal) b.appendParagraph('ゴール：' + m.goal).editAsText().setBold(true);

  b.appendParagraph('― DM手順（チャネル別・コピペ文面）―').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  (m.steps || []).forEach(function (s) {
    var head = 'STEP ' + (s.no || '') + '：' + (s.action || '');
    if (s.channel) head += '　【' + s.channel + '】';
    b.appendParagraph(head).setHeading(DocumentApp.ParagraphHeading.HEADING3);
    if (s.script) {
      b.appendParagraph('▼ そのまま送れる文面（{{ }}を差し替え）');
      var q = b.appendParagraph(s.script);
      q.setIndentStart(18);
      q.editAsText().setItalic(true);
    }
    if (s.timing) b.appendParagraph('タイミング：' + s.timing);
    if (s.point)  b.appendParagraph('注意：' + s.point);
  });

  // 会話トークスクリプト（面談/通話用）
  var t = m.talkScript;
  if (t) {
    b.appendParagraph('― 面談/通話トークスクリプト ―').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    if (t.opening) b.appendParagraph('■ 掴み：' + t.opening);
    if (t.hearing && t.hearing.length) {
      b.appendParagraph('■ ヒアリング');
      t.hearing.forEach(function (x) { b.appendListItem(x); });
    }
    if (t.proposal) b.appendParagraph('■ 提案：' + t.proposal);
    if (t.objection && t.objection.length) {
      b.appendParagraph('■ 反論処理');
      t.objection.forEach(function (x) { b.appendListItem('「' + (x.if || '') + '」→ ' + (x.say || '')); });
    }
    if (t.closing) b.appendParagraph('■ クロージング：' + t.closing);
  }

  if (m.ng && m.ng.length) {
    b.appendParagraph('やってはいけないこと').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    m.ng.forEach(function (x) { b.appendListItem('× ' + x); });
  }

  // ② 提案型
  b.appendParagraph('② 提案型（さらにアイデアが広がるきっかけ・数値つき）')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  var idea = o.ideas || {};
  (idea.angles || []).forEach(function (a) {
    b.appendParagraph('◆ ' + (a.title || '')).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    if (a.why)  b.appendParagraph('なぜ効く：' + a.why);
    if (a.try)  b.appendParagraph('試し方：' + a.try);
    b.appendParagraph('想定 返信率：' + (a.expReplyRate || '-') +
      '　/　面談化率：' + (a.expMeetingRate || '-') +
      '　/　提携(成約)率：' + (a.expPartnerRate || '-'))
      .editAsText().setBold(true);
    if (a.basis) b.appendParagraph('根拠：' + a.basis);
  });
  if (idea.experiments && idea.experiments.length) {
    b.appendParagraph('数値で検証すべき実験(A/B)').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    idea.experiments.forEach(function (x) { b.appendListItem(x); });
  }

  doc.saveAndClose();

  // フォルダへ移動
  if (folderId) {
    var file = DriveApp.getFileById(doc.getId());
    DriveApp.getFolderById(folderId).addFile(file);
    try { DriveApp.getRootFolder().removeFile(file); } catch (e) {}
  }
  return doc.getUrl();
}

/** 過去の効果が良かった文面/手法を学習データとして要約 */
function collectBestPatterns_() {
  var sh = openBook_().getSheetByName(CONFIG.SHEET_DM_MSG);
  if (!sh || sh.getLastRow() < 2) return '（まだ学習データが少ないため一般的なベストプラクティスで提案します）';
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, DM_MSG_HEADERS.length).getValues();
  var good = vals.filter(function (r) { return String(r[DM_MSG_COL.EFFECT - 1]).indexOf('良') === 0; });
  if (!good.length) return '（効果「良」の文面がまだありません。一般的なベストプラクティスで提案します）';
  return good.slice(0, 40).map(function (r) {
    return '・[' + r[DM_MSG_COL.PLATFORM - 1] + '/step' + r[DM_MSG_COL.STEP - 1] + '] ' +
      '狙い:' + r[DM_MSG_COL.INTENT - 1] + ' 文面:「' + String(r[DM_MSG_COL.MESSAGE - 1]).slice(0, 120) + '」';
  }).join('\n');
}

/** DM営業の数値レポート（返信率・転換率）を生成 → Doc & メール */
function generateDmReport() {
  var sh = openBook_().getSheetByName(CONFIG.SHEET_DM);
  if (!sh || sh.getLastRow() < 2) { Logger.log('DM案件がありません'); return; }
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, DM_HEADERS.length).getValues();

  var since = new Date(); since.setDate(since.getDate() - CONFIG.REPORT_DAYS);
  var rows = vals.filter(function (r) {
    var t = new Date(r[DM_COL.TS - 1]); return t >= since && r[DM_COL.SENT - 1];
  });
  if (!rows.length) { Logger.log('対象期間のDM案件がありません'); return; }

  // 集計
  var sent = 0, replies = 0, appts = 0, deals = 0;
  var byPlatform = {};
  rows.forEach(function (r) {
    sent += Number(r[DM_COL.SENT - 1]) || 0;
    replies += Number(r[DM_COL.REPLIES - 1]) || 0;
    appts += Number(r[DM_COL.APPTS - 1]) || 0;
    deals += Number(r[DM_COL.DEALS - 1]) || 0;
    var p = r[DM_COL.PLATFORM - 1] || '不明';
    byPlatform[p] = byPlatform[p] || { sent: 0, replies: 0, deals: 0 };
    byPlatform[p].sent += Number(r[DM_COL.SENT - 1]) || 0;
    byPlatform[p].replies += Number(r[DM_COL.REPLIES - 1]) || 0;
    byPlatform[p].deals += Number(r[DM_COL.DEALS - 1]) || 0;
  });
  var pct = function (a, b) { return b ? Math.round(a / b * 100) + '%' : '-'; };

  var stats = '全体：送信 ' + sent + ' / 返信 ' + replies + '（返信率 ' + pct(replies, sent) + '）' +
    ' / アポ ' + appts + ' / 成約 ' + deals + '（転換率 ' + pct(deals, sent) + '）\n';
  Object.keys(byPlatform).forEach(function (p) {
    var d = byPlatform[p];
    stats += '・' + p + '：返信率 ' + pct(d.replies, d.sent) + ' / 転換率 ' + pct(d.deals, d.sent) +
      '（送信' + d.sent + '）\n';
  });

  var techniques = rows.map(function (r) {
    return '・' + r[DM_COL.CUSTOMER - 1] + '(' + r[DM_COL.OUTCOME - 1] + ')：良かった点=' +
      r[DM_COL.GOOD - 1] + ' / 改善点=' + r[DM_COL.BAD - 1];
  }).join('\n').slice(0, 8000);

  var prompt =
    'あなたはDM営業の分析官です。直近' + CONFIG.REPORT_DAYS + '日のDM営業データから、返信率・転換率を上げる示唆をレポートしてください。\n' +
    '構成: 1.サマリー 2.数字から見える傾向 3.勝ちパターン 4.負けパターン 5.来週の改善アクション(数値目標つき)。\n\n' +
    '【数値】\n' + stats + '\n【各案件の良し悪し】\n' + techniques;
  var body = callGemini_(prompt, { temperature: 0.5, maxTokens: 4096 });

  var docId = prop_('DM_REPORT_DOC_ID', true);
  var doc = DocumentApp.openById(docId);
  var b = doc.getBody();
  var title = '【DM営業レポート】' + Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
  b.insertParagraph(0, '').appendHorizontalRule();
  b.insertParagraph(0, stats);
  b.insertParagraph(0, body);
  b.insertParagraph(0, title).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  doc.saveAndClose();

  var to = prop_('REPORT_RECIPIENTS', false);
  if (to) {
    MailApp.sendEmail({ to: to, subject: title, body: body + '\n\n【数値】\n' + stats + '\nレポートDoc: ' + doc.getUrl() });
  }
  logReport_('DM営業レポート', title, to || '');
  Logger.log('DM営業レポートを生成しました');
}

/** 既に取り込み済みのファイルIDセットを取得 */
function getKnownFileIds_(sh) {
  var map = {};
  var last = sh.getLastRow();
  if (last < 2) return map;
  var ids = sh.getRange(2, DM_COL.FILE_ID, last - 1, 1).getValues();
  ids.forEach(function (r) { if (r[0]) map[r[0]] = true; });
  return map;
}
