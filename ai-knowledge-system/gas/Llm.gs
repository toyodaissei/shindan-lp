/**
 * ============================================================
 *  Llm.gs : Gemini 呼び出しの共通ヘルパー
 *   - callGemini_()      … テキスト生成（JSON強制も可）
 *   - callGeminiJson_()  … JSONで返させて parse まで面倒みる
 *   - analyzeVideo_()    … 画面録画(動画ファイル)を Gemini に「視聴」させて解析
 *                          → Drive のDM営業録画を理解するための中核
 * ============================================================
 */

/**
 * Gemini呼び出し用のフェッチ（一時的な混雑 429/500/503 は自動リトライ）。
 * 2s→4s→8s→16s と待って最大5回試す。GAS実行時間の都合で上限は控えめ。
 */
function geminiFetch_(url, options) {
  var delay = 2000, res;
  for (var i = 0; i < 5; i++) {
    res = UrlFetchApp.fetch(url, options);
    var code = res.getResponseCode();
    if ((code === 429 || code === 500 || code === 503) && i < 4) {
      Logger.log('Gemini混雑(' + code + ') リトライ ' + (i + 1) + '回目 … ' + (delay / 1000) + '秒待機');
      Utilities.sleep(delay);
      delay = Math.min(delay * 2, 16000);
      continue;
    }
    return res;
  }
  return res;
}

/** テキストプロンプト → 文字列を返す */
function callGemini_(prompt, opts) {
  opts = opts || {};
  var key = prop_('GEMINI_API_KEY', true);
  var model = opts.model || CONFIG.GEMINI_MODEL;
  var url = CONFIG.GEMINI_API_BASE + '/v1beta/models/' + model + ':generateContent?key=' + key;

  var payload = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: opts.temperature != null ? opts.temperature : 0.4,
      maxOutputTokens: opts.maxTokens || 4096
    }
  };
  if (opts.json) {
    payload.generationConfig.responseMimeType = 'application/json';
  }
  // 構造化出力：必須フィールドをスキーマで強制（Geminiが要素を落とすのを防ぐ）
  if (opts.schema) {
    payload.generationConfig.responseMimeType = 'application/json';
    payload.generationConfig.responseSchema = opts.schema;
  }
  applyThinkingConfig_(payload.generationConfig);
  if (opts.system) {
    payload.systemInstruction = { parts: [{ text: opts.system }] };
  }

  var res = geminiFetch_(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  return parseGeminiResponse_(res);
}

/** JSONで返させて object を返す（失敗時は例外） */
function callGeminiJson_(prompt, opts) {
  opts = opts || {};
  opts.json = true;
  var text = callGemini_(prompt, opts);
  try {
    return JSON.parse(text);
  } catch (e) {
    // まれに```json ... ```で包まれる事があるので救済
    var m = text.match(/\{[\s\S]*\}/);
    if (m) return JSON.parse(m[0]);
    throw new Error('GeminiのJSON解析に失敗: ' + text.slice(0, 300));
  }
}

/**
 * 画面録画(動画)を Gemini に視聴させて解析する。
 *  1) Gemini File API に動画をアップロード（resumable）
 *  2) ACTIVE になるまで待機
 *  3) fileUri を付けて generateContent
 *
 * @param {Blob}   blob   動画Blob（Driveファイルなど）
 * @param {string} prompt 指示文
 * @param {Object} opts   {json:true, model:...}
 * @return {string|Object}
 */
function analyzeVideo_(blob, prompt, opts) {
  opts = opts || {};
  var key = prop_('GEMINI_API_KEY', true);
  var model = opts.model || CONFIG.GEMINI_MODEL_VIDEO;
  var base = CONFIG.GEMINI_API_BASE;
  var bytes = blob.getBytes();
  var mime = blob.getContentType() || 'video/mp4';

  // --- 1) アップロード開始（resumable, 1リクエスト完結） ---
  var startRes = UrlFetchApp.fetch(base + '/upload/v1beta/files?key=' + key, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'X-Goog-Upload-Protocol': 'resumable',
      'X-Goog-Upload-Command': 'start',
      'X-Goog-Upload-Header-Content-Length': String(bytes.length),
      'X-Goog-Upload-Header-Content-Type': mime
    },
    payload: JSON.stringify({ file: { display_name: blob.getName() || 'recording' } }),
    muteHttpExceptions: true
  });
  var uploadUrl = startRes.getHeaders()['X-Goog-Upload-URL'] ||
                  startRes.getHeaders()['x-goog-upload-url'];
  if (!uploadUrl) {
    throw new Error('Gemini File APIのアップロードURL取得に失敗: ' + startRes.getContentText().slice(0, 300));
  }

  // --- 2) バイト送信 + finalize ---
  var upRes = UrlFetchApp.fetch(uploadUrl, {
    method: 'post',
    headers: {
      'X-Goog-Upload-Offset': '0',
      'X-Goog-Upload-Command': 'upload, finalize'
    },
    contentType: mime,
    payload: bytes,
    muteHttpExceptions: true
  });
  var fileInfo = JSON.parse(upRes.getContentText());
  var fileUri = fileInfo.file && fileInfo.file.uri;
  var fileName = fileInfo.file && fileInfo.file.name;
  if (!fileUri) {
    throw new Error('Gemini File APIのアップロード失敗: ' + upRes.getContentText().slice(0, 300));
  }

  // --- 3) ACTIVE になるまで待機（動画は処理に数秒〜数十秒） ---
  var state = fileInfo.file.state;
  var tries = 0;
  while (state === 'PROCESSING' && tries < 30) {
    Utilities.sleep(4000);
    var chk = UrlFetchApp.fetch(base + '/v1beta/' + fileName + '?key=' + key, { muteHttpExceptions: true });
    state = JSON.parse(chk.getContentText()).state;
    tries++;
  }
  if (state !== 'ACTIVE') {
    throw new Error('動画の前処理が完了しませんでした（state=' + state + '）');
  }

  // --- 4) 生成 ---
  var payload = {
    contents: [{
      role: 'user',
      parts: [
        { fileData: { mimeType: mime, fileUri: fileUri } },
        { text: prompt }
      ]
    }],
    generationConfig: {
      temperature: opts.temperature != null ? opts.temperature : 0.3,
      maxOutputTokens: opts.maxTokens || 8192
    }
  };
  if (opts.json) payload.generationConfig.responseMimeType = 'application/json';
  applyThinkingConfig_(payload.generationConfig);

  var genRes = geminiFetch_(base + '/v1beta/models/' + model + ':generateContent?key=' + key, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  var text = parseGeminiResponse_(genRes);

  // 後始末（無料枠の容量節約。失敗しても無視）
  try { UrlFetchApp.fetch(base + '/v1beta/' + fileName + '?key=' + key, { method: 'delete', muteHttpExceptions: true }); } catch (e) {}

  if (opts.json) {
    try { return JSON.parse(text); }
    catch (e) { var m = text.match(/\{[\s\S]*\}/); if (m) return JSON.parse(m[0]); throw e; }
  }
  return text;
}

/**
 * 思考(thinking)型モデル対策：generationConfigにthinkingConfigを付与。
 * 思考トークンが出力上限を食い尽くして回答が空になるのを防ぐ。
 * CONFIG.GEMINI_THINKING_BUDGET が null の場合は何もしない（旧モデル互換）。
 */
function applyThinkingConfig_(generationConfig) {
  var budget = CONFIG.GEMINI_THINKING_BUDGET;
  if (budget == null) return;
  generationConfig.thinkingConfig = { thinkingBudget: budget };
}

/**
 * スクリーンショット(画像)を Gemini に解析させる（inlineData方式・軽量）。
 * DM営業のスクショからチャット内容を読み取る用途。
 */
function analyzeImage_(blob, prompt, opts) {
  opts = opts || {};
  var key = prop_('GEMINI_API_KEY', true);
  var model = opts.model || CONFIG.GEMINI_MODEL_VIDEO;
  var payload = {
    contents: [{
      role: 'user',
      parts: [
        { inlineData: { mimeType: blob.getContentType() || 'image/png', data: Utilities.base64Encode(blob.getBytes()) } },
        { text: prompt }
      ]
    }],
    generationConfig: {
      temperature: opts.temperature != null ? opts.temperature : 0.3,
      maxOutputTokens: opts.maxTokens || 8192
    }
  };
  if (opts.json) payload.generationConfig.responseMimeType = 'application/json';
  applyThinkingConfig_(payload.generationConfig);

  var res = geminiFetch_(CONFIG.GEMINI_API_BASE + '/v1beta/models/' + model + ':generateContent?key=' + key, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(payload), muteHttpExceptions: true
  });
  var text = parseGeminiResponse_(res);
  if (opts.json) {
    try { return JSON.parse(text); }
    catch (e) { var m = text.match(/\{[\s\S]*\}/); if (m) return JSON.parse(m[0]); throw e; }
  }
  return text;
}

/**
 * 複数のスクショ(画像)をまとめて Gemini に解析させる（会話の時系列を1件として読む）。
 */
function analyzeImages_(blobs, prompt, opts) {
  opts = opts || {};
  if (!blobs || !blobs.length) throw new Error('画像がありません');
  var key = prop_('GEMINI_API_KEY', true);
  var model = opts.model || CONFIG.GEMINI_MODEL_VIDEO;
  var parts = blobs.map(function (bl) {
    return { inlineData: { mimeType: bl.getContentType() || 'image/png', data: Utilities.base64Encode(bl.getBytes()) } };
  });
  parts.push({ text: prompt });
  var payload = {
    contents: [{ role: 'user', parts: parts }],
    generationConfig: {
      temperature: opts.temperature != null ? opts.temperature : 0.3,
      maxOutputTokens: opts.maxTokens || 8192
    }
  };
  if (opts.json) payload.generationConfig.responseMimeType = 'application/json';
  applyThinkingConfig_(payload.generationConfig);

  var res = geminiFetch_(CONFIG.GEMINI_API_BASE + '/v1beta/models/' + model + ':generateContent?key=' + key, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(payload), muteHttpExceptions: true
  });
  var text = parseGeminiResponse_(res);
  if (opts.json) {
    try { return JSON.parse(text); }
    catch (e) { var m = text.match(/\{[\s\S]*\}/); if (m) return JSON.parse(m[0]); throw e; }
  }
  return text;
}

/** Gemini レスポンス共通パーサ */
function parseGeminiResponse_(res) {
  var code = res.getResponseCode();
  var body = res.getContentText();
  if (code !== 200) {
    throw new Error('Gemini APIエラー(' + code + '): ' + body.slice(0, 400));
  }
  var data = JSON.parse(body);
  if (data.promptFeedback && data.promptFeedback.blockReason) {
    throw new Error('Geminiにブロックされました: ' + data.promptFeedback.blockReason);
  }
  var cand = data.candidates && data.candidates[0];
  if (!cand || !cand.content || !cand.content.parts) {
    throw new Error('Geminiの応答が空です: ' + body.slice(0, 300));
  }
  return cand.content.parts.map(function (p) { return p.text || ''; }).join('').trim();
}
