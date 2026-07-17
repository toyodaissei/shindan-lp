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
  var root = DriveApp.getFolderById(folderId);
  var sh = sheet_(CONFIG.SHEET_DM, DM_HEADERS);
  // 失敗/対象なしの行は毎回クリア → 後から素材を入れたフォルダも再挑戦できる
  cleanDmErrorRows_(sh);
  var known = getKnownDoneIds_(sh);   // 成功状態のIDのみ「処理済み」として扱う
  var processed = 0;
  var errors = [];

  // (1) サブフォルダ = 1案件（複数スクショ/録画をまとめて時系列で解析）
  var subs = root.getFolders();
  while (subs.hasNext() && processed < CONFIG.MAX_PROCESS_PER_RUN) {
    var sub = subs.next();
    if (known[sub.getId()]) continue;
    try {
      var a = analyzeDmCaseFromFolder_(sub);
      if (!a) {
        markDmError_(sh, sub.getId(), sub.getName(), '画像/動画が無い（アップロード完了後に再実行してください）');
        errors.push(sub.getName() + '：フォルダ内に画像/動画が見つかりません');
        continue;
      }
      writeDmCase_(sh, idFile_(sub.getId(), sub.getName()), a);
      appendDmMessages_(a);
      processed++;
    } catch (err) {
      Logger.log('案件フォルダ解析 失敗 ' + sub.getName() + ': ' + err);
      markDmError_(sh, sub.getId(), sub.getName(), '解析失敗: ' + String(err).slice(0, 80));
      errors.push(sub.getName() + '：' + String(err).slice(0, 120));
    }
  }

  // (2) 直下の単体ファイル = 1案件（従来どおり）
  var files = root.getFiles();
  while (files.hasNext() && processed < CONFIG.MAX_PROCESS_PER_RUN) {
    var file = files.next();
    var mime = file.getMimeType();
    if (mime.indexOf('video') !== 0 && mime.indexOf('image') !== 0) continue;
    if (known[file.getId()]) continue;
    try {
      var analysis = analyzeDmRecording_(file);
      writeDmCase_(sh, file, analysis);
      appendDmMessages_(analysis);
      processed++;
    } catch (err) {
      Logger.log('録画解析 失敗 ' + file.getName() + ': ' + err);
      markDmError_(sh, file.getId(), file.getName(), '解析失敗: ' + String(err).slice(0, 80));
      errors.push(file.getName() + '：' + String(err).slice(0, 120));
    }
  }

  Logger.log('ingestDmRecordings: ' + processed + '件を解析 / エラー' + errors.length + '件');
  return { processed: processed, errors: errors };
}

/** 成功状態(処理済み)のIDだけを返す。失敗/対象なしは含めない＝再挑戦できる */
function getKnownDoneIds_(sh) {
  var map = {};
  var last = sh.getLastRow();
  if (last < 2) return map;
  var done = { '解析済(提案待ち)': 1, '承認待ち': 1, '提案済': 1, '送付済': 1, '見送り': 1 };
  var vals = sh.getRange(2, 1, last - 1, DM_HEADERS.length).getValues();
  vals.forEach(function (r) {
    var id = r[DM_COL.FILE_ID - 1];
    if (id && done[r[DM_COL.STATUS - 1]]) map[id] = true;
  });
  return map;
}

/** 失敗/対象なしの行を削除（毎回クリアして再挑戦できるように・下から削除） */
function cleanDmErrorRows_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return;
  var done = { '解析済(提案待ち)': 1, '承認待ち': 1, '提案済': 1, '送付済': 1, '見送り': 1 };
  var vals = sh.getRange(2, 1, last - 1, DM_HEADERS.length).getValues();
  for (var i = vals.length - 1; i >= 0; i--) {
    if (!done[vals[i][DM_COL.STATUS - 1]]) sh.deleteRow(i + 2);
  }
}

/** DM解析の共通プロンプト（単体でもフォルダまとめでも同じ観点） */
function dmAnalysisPrompt_() {
  return 'これは当社の営業担当者が、SNS(Threads/Instagram/X/LINE等)のDM・スレッドで行った「実際の営業・開拓のやり取り」の' +
    '画面録画またはスクリーンショットです。画面に映る事実だけを正確に読み取り、営業手法を分析してください。\n' +
    '【厳守】録画に映っていない事柄を推測で足さない。特に「相手のフォロワーに広告を出す/リーチを買う」等の' +
    'インフルエンサーマーケティングを、実際にそうでない限り絶対に当てはめない。\n' +
    '必ず次のJSONのみ返す:\n' +
    '{' +
    '"objective":"この営業の目的を一言で。録画の実態に基づく。例:『代理店/協業パートナーの開拓(相手にハブになってもらい顧客/学生を送客してもらう)』『自社送客の直接営業』『特定商材の販売』など",' +
    '"customer":"相手(アカウント名/相手の属性)。分からなければ推定や空文字",' +
    '"product":"案件/トピック名(例:学生送客の代理店開拓 / 〇〇商材の販売 など、今回のスコープ)",' +
    '"platform":"チャネル(Instagram/InstagramDM/Threads/ThreadsDM/X/LINE/公式LINE/Facebook/その他)",' +
    '"summary":"実際に何が行われたかの要約(4〜6行)",' +
    '"technique":"使われた営業手法・トークの型(改行区切りで具体的に。例:定型文で打診→アポ取り→ミート設定 など実際の流れ)",' +
    '"reaction":"相手の反応・温度感",' +
    '"outcome":"未返信 / 返信 / 面談 / 提携 のいずれか(判断できなければ最も近いもの。提携=成約)",' +
    '"sent":送ったメッセージ数(整数),"replies":相手の返信数(整数),' +
    '"good":"良かった点(なぜ効いたか)",' +
    '"bad":"改善点(もっとこうすれば返信/面談/提携が増えるか)",' +
    '"messages":[{"step":1,"message":"実際に送った文面(できるだけ原文)","intent":"その一手の狙い/技術","reaction":"相手の反応","effect":"良/普/悪"}]' +
    '}\n' +
    '文面は新人が真似できるよう、できるだけ原文に近い形で書き出すこと。';
}

/** 単体ファイル(動画/画像)を解析 */
function analyzeDmRecording_(file) {
  var mime = file.getMimeType();
  var prompt = dmAnalysisPrompt_();
  if (mime.indexOf('image') === 0) {
    return analyzeImage_(file.getBlob(), prompt, { json: true, maxTokens: 8192 });
  }
  return analyzeVideo_(file.getBlob(), prompt, { json: true, maxTokens: 8192 });
}

/** 1案件フォルダ内の複数スクショ/録画をまとめて解析 */
function analyzeDmCaseFromFolder_(folder) {
  var imgs = [], vids = [];
  var it = folder.getFiles();
  while (it.hasNext()) {
    var f = it.next();
    var m = f.getMimeType() || '';
    if (m.indexOf('image') === 0) imgs.push(f);
    else if (m.indexOf('video') === 0) vids.push(f);
  }
  // 作成/アップロード日時の順に並べる（番号付け不要。撮った順＝会話の時系列）
  sortByCreated_(imgs);
  var prompt = dmAnalysisPrompt_();
  if (imgs.length) {
    var blobs = imgs.slice(0, 12).map(function (f) { return f.getBlob(); }); // API負荷対策で最大12枚
    return analyzeImages_(blobs,
      prompt + '\n※複数枚のスクショは同じ会話の断片で、おおむね時系列順に並べてあります。' +
      '画面内のタイムスタンプや文脈から必要なら順序を補正し、全体を1件の営業として分析してください。',
      { json: true, maxTokens: 8192 });
  }
  if (vids.length) {
    return analyzeVideo_(vids[0].getBlob(), prompt, { json: true, maxTokens: 8192 });
  }
  return null;
}

/** 作成/アップロード日時の昇順に並べる（同時刻は名前で安定化）。番号付け不要にするための自動整列 */
function sortByCreated_(files) {
  files.sort(function (a, b) {
    var d = a.getDateCreated().getTime() - b.getDateCreated().getTime();
    return d !== 0 ? d : (a.getName() < b.getName() ? -1 : 1);
  });
  return files;
}

/** writeDmCase_ に渡す簡易ファイル風オブジェクト（IDと名前だけ） */
function idFile_(id, name) {
  return { getId: function () { return id; }, getName: function () { return name; } };
}

/** 解析できなかった対象をエラー行として記録（無限リトライ防止） */
function markDmError_(sh, id, name, msg) {
  var row = [];
  row[DM_COL.ID - 1] = newId_('dm');
  row[DM_COL.TS - 1] = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm');
  row[DM_COL.FILE_ID - 1] = id;
  row[DM_COL.FILE_NAME - 1] = name;
  row[DM_COL.STATUS - 1] = msg;
  sh.appendRow(row);
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
  row[DM_COL.SUMMARY - 1] = (a.objective ? '【営業の目的】' + a.objective + '\n' : '') + (a.summary || '');
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
    'あなたは営業の「型化(再現性化)」のプロです。以下は、当社の営業担当者が実際に行った営業・開拓のやり取り' +
    '(画面録画/スクショから解析済み)です。この“実際のやり方”を、他の社員がそっくり真似して再現できるマニュアルと、' +
    'そこから発想を広げるための提案に落とします。\n' +
    '【最重要・絶対厳守】\n' +
    '1) 提案は必ず「録画から読み取った実際の営業目的・相手・手法」に忠実に沿わせる。' +
    '勝手に別のビジネスモデル(例:インフルエンサーマーケティング／相手のフォロワーにリーチや広告を売る 等)を持ち込まない。\n' +
    '2) 下記「営業の目的」が代理店/協業パートナーの開拓(相手にハブになってもらい顧客/学生を送客してもらう)なら、' +
    'その"協業関係を結ぶ"提案にする。相手のフォロワーに広告を出す話には絶対にしない。\n' +
    '3) 案件のトピック(例:学生送客/特定商材/自社送客 など)に厳密にスコープを合わせ、他トピックを混ぜない。\n' +
    '4) マニュアル型=録画の実際のやり方（定型文→アポ→ミート等）を新人がそのまま再現できる手順に。' +
    '提案型=そのやり方を土台に、新しいアイデアが湧くきっかけ(切り口＋数値)に。\n' +
    '過去の高成績パターンも学習材料に使ってください。\n\n' +
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
    '・"talkScript.objection" には、この案件の目的・相手に即して相手が実際に抱く懸念への切り返しを2〜4個入れる' +
    '(例:協業のメリットが不明/手間がかかりそう/実績・信頼性は?/既に他社とやっている 等、案件に合うものを選ぶ)。' +
    '案件と無関係な一般論の懸念は入れない。\n' +
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

  return writeDmProposalDoc_(caseRow[DM_COL.CUSTOMER - 1], caseRow[DM_COL.PRODUCT - 1],
    caseRow[DM_COL.PLATFORM - 1], o, caseRow[DM_COL.FILE_ID - 1]);
}

/** 提案書のカラーパレット（LPと統一） */
var DOC_STYLE = {
  navy: '#1c2b4a', gray: '#5a6577',
  varc: '#1558b0',   // 変数セクション（青）
  manual: '#0e9f7e', // マニュアル型（ティール）
  idea: '#c77d17',   // 提案型（アンバー）
  box: '#eef3f8'     // コピペ文面の背景
};

/** セクション見出し（色バンド） */
function docBand_(b, text, hex) {
  var p = b.appendParagraph('  ' + text + '  ');
  p.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  p.setSpacingBefore(20).setSpacingAfter(10);
  p.editAsText().setForegroundColor('#ffffff').setBackgroundColor(hex).setBold(true).setFontSize(14);
  return p;
}
/** 小見出し（色文字） */
function docSub_(b, text, hex) {
  var p = b.appendParagraph(text);
  p.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  p.setSpacingBefore(12).setSpacingAfter(4);
  p.editAsText().setForegroundColor(hex).setBold(true).setFontSize(12.5);
  return p;
}
/** 本文行 */
function docText_(b, text, opt) {
  opt = opt || {};
  var p = b.appendParagraph(text);
  p.setSpacingAfter(opt.after != null ? opt.after : 3);
  var t = p.editAsText();
  t.setFontSize(opt.size || 11).setForegroundColor(opt.color || DOC_STYLE.navy).setBold(!!opt.bold);
  if (opt.indent) p.setIndentStart(opt.indent);
  return p;
}

/**
 * 「📸 実際のやりとり画面（参考）」セクションを作り、案件の元ファイルを配置。
 *  画像(スクショ)は埋め込み、動画はリンク。①の構成に合わせて変数一覧の直後に置く。
 */
function embedSourceMedia_(b, id) {
  if (!id || String(id).indexOf('TEST-') === 0) return; // テスト用ダミーは除外
  var S = DOC_STYLE;

  // フォルダ(複数スクショ) か 単体ファイル かを判定
  var folder = null, file = null;
  try { folder = DriveApp.getFolderById(id); } catch (e) {}
  if (!folder) { try { file = DriveApp.getFileById(id); } catch (e2) { return; } }

  docBand_(b, '📸 実際のやりとり画面（参考）', S.navy);
  try {
    if (folder) {
      var imgs = [], vids = [];
      var it = folder.getFiles();
      while (it.hasNext()) {
        var f = it.next(); var m = f.getMimeType() || '';
        if (m.indexOf('image') === 0) imgs.push(f);
        else if (m.indexOf('video') === 0) vids.push(f);
      }
      sortByCreated_(imgs);
      imgs.forEach(function (f, i) { embedOneImage_(b, f, (i + 1) + '/' + imgs.length + '　' + f.getName()); });
      vids.forEach(function (f) { linkOneVideo_(b, f); });
      if (!imgs.length && !vids.length) docText_(b, '（このフォルダに画像/動画がありません）', { color: S.gray, size: 9, after: 8 });
    } else {
      var mime = file.getMimeType() || '';
      if (mime.indexOf('image') === 0) embedOneImage_(b, file, file.getName());
      else linkOneVideo_(b, file);
    }
  } catch (e) {
    docText_(b, '（元ファイルの表示に失敗：' + e + '）', { color: S.gray, size: 9, after: 8 });
  }
}

/** 画像1枚をページ幅に収めて埋め込む */
function embedOneImage_(b, file, caption) {
  var S = DOC_STYLE;
  try {
    var img = b.appendImage(file.getBlob());
    var maxW = 460;
    if (img.getWidth() && img.getWidth() > maxW) {
      var h = Math.round(img.getHeight() * (maxW / img.getWidth()));
      img.setWidth(maxW).setHeight(h);
    }
    docText_(b, caption, { color: S.gray, size: 9, after: 8 });
  } catch (e) {
    docText_(b, '（画像の埋め込み失敗：' + file.getName() + '）', { color: S.gray, size: 9, after: 6 });
  }
}

/** 動画はリンクで置く */
function linkOneVideo_(b, file) {
  var link = b.appendParagraph('🎬 録画を見る（' + file.getName() + '）');
  link.setSpacingAfter(8);
  link.editAsText().setLinkUrl(file.getUrl()).setFontSize(11);
}

/** 提案JSON → 色分け・余白・文字サイズを整えたGoogle Docに書き出し、URLを返す */
function writeDmProposalDoc_(customer, product, platform, o, fileId) {
  var S = DOC_STYLE;
  var folderId = prop_('PROPOSAL_FOLDER', false);
  var topic = product || customer || '営業案件';
  var name = '営業ナレッジ提案_' + topic + '_' +
    Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd-HHmm');
  var doc = DocumentApp.create(name);
  var b = doc.getBody();
  b.setMarginTop(48).setMarginBottom(48).setMarginLeft(56).setMarginRight(56);

  // タイトル
  var title = b.appendParagraph('営業ナレッジ 提案書');
  title.setHeading(DocumentApp.ParagraphHeading.TITLE).setSpacingAfter(2);
  title.editAsText().setForegroundColor(S.navy).setFontSize(24).setBold(true);
  // スコープ（何の案件か）
  docText_(b, '対象トピック：' + topic, { color: S.manual, bold: true, size: 13, after: 1 });
  docText_(b, '相手：' + (customer || '-') + '　｜　主なチャネル：' + (platform || '-'),
    { color: S.gray, size: 10, after: 2 });
  b.appendHorizontalRule();

  // ◇ 差し込み変数一覧
  if (o.variables && o.variables.length) {
    docBand_(b, '◇ 差し込み変数一覧（送信前にここを埋める）', S.varc);
    o.variables.forEach(function (v) {
      var li = b.appendListItem((v.key || '') + ' … ' + (v.desc || '') +
        (v.example ? '（例: ' + v.example + '）' : ''));
      li.setSpacingAfter(2);
      li.editAsText().setFontSize(11).setForegroundColor(S.navy);
      li.editAsText().setForegroundColor(0, Math.max(0, (v.key || '').length - 1), S.varc);
    });
  }

  // 📸 実際のやりとり画面（参考）… 変数一覧の直後に元スクショ/録画を集約
  embedSourceMedia_(b, fileId);

  // ① マニュアル型
  var m = o.manual || {};
  docBand_(b, '① マニュアル型（新人が1→10でそのまま真似できる手順書）', S.manual);
  if (m.goal) docText_(b, '🎯 ゴール：' + m.goal, { bold: true, size: 11.5, after: 6 });

  docSub_(b, '― DM手順（チャネル別・コピペ文面）―', S.manual);
  (m.steps || []).forEach(function (s) {
    var head = b.appendParagraph('STEP ' + (s.no || '') + '　' + (s.action || ''));
    head.setHeading(DocumentApp.ParagraphHeading.HEADING3).setSpacingBefore(8).setSpacingAfter(2);
    head.editAsText().setForegroundColor(S.navy).setBold(true).setFontSize(11.5);
    if (s.channel) docText_(b, '📱 チャネル：' + s.channel, { color: S.manual, size: 10, after: 2 });
    if (s.script) {
      docText_(b, '▼ そのまま送れる文面（{{ }}を差し替え）', { color: S.gray, size: 9, bold: true, after: 1 });
      var q = b.appendParagraph(s.script);
      q.setIndentStart(14).setIndentEnd(6).setSpacingAfter(6).setSpacingBefore(2);
      q.editAsText().setBackgroundColor(S.box).setForegroundColor(S.navy).setFontSize(11);
    }
    if (s.timing) docText_(b, '⏰ タイミング：' + s.timing, { color: S.gray, size: 10, after: 1 });
    if (s.point)  docText_(b, '⚠️ 注意：' + s.point, { color: S.gray, size: 10, after: 4 });
  });

  // 会話トークスクリプト
  var t = m.talkScript;
  if (t) {
    docSub_(b, '― 面談/通話トークスクリプト ―', S.manual);
    if (t.opening) docText_(b, '■ 掴み：' + t.opening, { after: 4 });
    if (t.hearing && t.hearing.length) {
      docText_(b, '■ ヒアリング', { bold: true, after: 1 });
      t.hearing.forEach(function (x) { b.appendListItem(x).editAsText().setFontSize(11).setForegroundColor(S.navy); });
    }
    if (t.proposal) docText_(b, '■ 提案：' + t.proposal, { after: 4 });
    if (t.objection && t.objection.length) {
      docText_(b, '■ 反論処理', { bold: true, after: 1 });
      t.objection.forEach(function (x) {
        b.appendListItem('「' + (x['if'] || '') + '」→ ' + (x.say || ''))
          .editAsText().setFontSize(11).setForegroundColor(S.navy);
      });
    }
    if (t.closing) docText_(b, '■ クロージング：' + t.closing, { after: 4 });
  }

  if (m.ng && m.ng.length) {
    docSub_(b, '🚫 やってはいけないこと', '#c0392b');
    m.ng.forEach(function (x) {
      b.appendListItem('× ' + x).editAsText().setFontSize(11).setForegroundColor('#c0392b');
    });
  }

  // ② 提案型
  docBand_(b, '② 提案型（さらにアイデアが広がるきっかけ・数値つき）', S.idea);
  var idea = o.ideas || {};
  (idea.angles || []).forEach(function (a) {
    docSub_(b, '◆ ' + (a.title || ''), S.idea);
    if (a.why)  docText_(b, 'なぜ効く：' + a.why, { after: 2 });
    if (a['try']) docText_(b, '試し方：' + a['try'], { after: 3 });
    var kpi = b.appendParagraph('📊 想定 返信率 ' + (a.expReplyRate || '-') +
      '　｜　面談化率 ' + (a.expMeetingRate || '-') +
      '　｜　提携(成約)率 ' + (a.expPartnerRate || '-'));
    kpi.setSpacingAfter(2).setIndentStart(8);
    kpi.editAsText().setBackgroundColor('#fbeecf').setForegroundColor(S.navy).setBold(true).setFontSize(11);
    if (a.basis) docText_(b, '根拠：' + a.basis, { color: S.gray, size: 10, after: 6 });
  });
  if (idea.experiments && idea.experiments.length) {
    docSub_(b, '🧪 数値で検証すべき実験(A/B)', S.idea);
    idea.experiments.forEach(function (x) { b.appendListItem(x).editAsText().setFontSize(11).setForegroundColor(S.navy); });
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
