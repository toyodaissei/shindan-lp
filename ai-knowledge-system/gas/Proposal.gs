/**
 * ============================================================
 *  🅰 商談 → 提案書(Googleスライド) 自動生成 → 送付
 *     （動画②「Nottaで録音 → GASで組んで → 5分後に提案書」の再現）
 * ------------------------------------------------------------
 *  流れ:
 *   1. 商談の文字起こしを「商談」シートに投入（doPostDeal / n8n / Notta連携）
 *   2. processDeals() が未処理の商談をAIで提案書構成に整形
 *   3. Googleスライドの提案書を自動生成（PDF化も）
 *   4. 依頼者へメール（顧客への送付はレビュー後 or 自動、設定で切替）
 * ============================================================
 */

var DEAL_COL = {
  ID: 1, TS: 2, CUSTOMER: 3, OWNER_EMAIL: 4, RAW: 5,
  STATUS: 6, SLIDE_URL: 7, PDF_URL: 8, PROCESSED_AT: 9
};
var DEAL_HEADERS = ['id', '日時', '顧客', '担当メール', '商談文字起こし',
  'ステータス', '提案書(スライド)', '提案書(PDF)', '生成日時'];

/**
 * Webエンドポイント（Notta/n8n から商談の文字起こしをPOST）
 *  ※ doPost はCode.gs（ナレッジ用）が使うので、こちらはパラメータで振り分け:
 *     POST ?target=deal  もしくは body.target="deal"
 */
function routePost_(e, body) {
  if ((e.parameter && e.parameter.target === 'deal') || body.target === 'deal') {
    var id = addDeal({
      customer: body.customer || body.顧客 || '',
      ownerEmail: body.ownerEmail || body.担当 || prop_('REPORT_RECIPIENTS', false),
      transcript: body.transcript || body.text || '',
      ts: body.ts
    });
    return { ok: true, id: id, kind: 'deal' };
  }
  return null; // ナレッジとして扱う
}

/** 商談を1件追加 */
function addDeal(o) {
  var sh = sheet_(CONFIG.SHEET_DEALS, DEAL_HEADERS);
  var id = newId_('deal');
  var ts = o.ts ? new Date(o.ts) : new Date();
  var row = [];
  row[DEAL_COL.ID - 1] = id;
  row[DEAL_COL.TS - 1] = Utilities.formatDate(ts, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm');
  row[DEAL_COL.CUSTOMER - 1] = o.customer || '';
  row[DEAL_COL.OWNER_EMAIL - 1] = o.ownerEmail || '';
  row[DEAL_COL.RAW - 1] = o.transcript || '';
  row[DEAL_COL.STATUS - 1] = '未処理';
  sh.appendRow(row);
  return id;
}

/** 未処理の商談を提案書化（時間トリガー：5〜10分毎） */
function processDeals() {
  var sh = sheet_(CONFIG.SHEET_DEALS, DEAL_HEADERS);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, DEAL_HEADERS.length).getValues();
  var created = [];

  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    if (r[DEAL_COL.STATUS - 1] !== '未処理') continue;
    var rowNum = i + 2;
    try {
      var content = buildProposalContent_(r[DEAL_COL.CUSTOMER - 1], r[DEAL_COL.RAW - 1]);
      var made = buildSlides_(r[DEAL_COL.CUSTOMER - 1], content);
      created.push({ customer: r[DEAL_COL.CUSTOMER - 1], slideUrl: made.slideUrl });
      sh.getRange(rowNum, DEAL_COL.SLIDE_URL).setValue(made.slideUrl);
      sh.getRange(rowNum, DEAL_COL.PDF_URL).setValue(made.pdfUrl);
      sh.getRange(rowNum, DEAL_COL.STATUS).setValue('生成済');
      sh.getRange(rowNum, DEAL_COL.PROCESSED_AT).setValue(
        Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm'));

      var to = r[DEAL_COL.OWNER_EMAIL - 1] || prop_('REPORT_RECIPIENTS', false);
      if (to) {
        MailApp.sendEmail({
          to: to,
          subject: '【提案書できました】' + r[DEAL_COL.CUSTOMER - 1] + ' 様',
          body: '商談の文字起こしから提案書を自動生成しました。\n\n' +
            'スライド: ' + made.slideUrl + '\nPDF: ' + made.pdfUrl +
            '\n\n内容を確認のうえ、お客様へご送付ください。'
        });
      }
    } catch (err) {
      sh.getRange(rowNum, DEAL_COL.STATUS).setValue('失敗: ' + String(err).slice(0, 80));
      Logger.log('processDeals 行' + rowNum + ' 失敗: ' + err);
    }
  }
  return created;
}

/** 商談文字起こし → 提案書の構成をJSONで得る */
function buildProposalContent_(customer, transcript) {
  var prompt =
    'あなたは一流の法人営業です。以下は「' + customer + '」様との商談の文字起こしです。\n' +
    'これを提案書スライドの構成に落としてください。必ず次のJSONのみ返す:\n' +
    '{"title":"提案タイトル","subtitle":"サブタイトル",' +
    '"slides":[{"heading":"見出し","bullets":["要点1","要点2","要点3"]}]}\n' +
    '推奨スライド構成: ①現状の課題整理 ②目指す姿(ゴール) ③ご提案内容 ④期待できる効果(できれば数値) ' +
    '⑤導入ステップ ⑥お見積り/プラン ⑦次のアクション。\n' +
    '各スライドの bullets は3〜5個、簡潔に。\n\n' +
    '商談文字起こし:\n' + String(transcript).slice(0, 12000);
  return callGeminiJson_(prompt, { temperature: 0.4, maxTokens: 4096 });
}

/** 構成JSON → Googleスライドを生成し、URLとPDFを返す */
function buildSlides_(customer, content) {
  var folderId = prop_('PROPOSAL_FOLDER', false);
  var name = '提案書_' + customer + '_' +
    Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd-HHmm');

  var pres = SlidesApp.create(name);
  var slides = pres.getSlides();

  // 表紙
  var cover = slides[0];
  cover.getPlaceholders().forEach(function (ph) {
    var t = ph.asShape().getPlaceholderType();
    if (t === SlidesApp.PlaceholderType.CENTERED_TITLE || t === SlidesApp.PlaceholderType.TITLE) {
      ph.asShape().getText().setText(content.title || (customer + ' 様 ご提案'));
    } else if (t === SlidesApp.PlaceholderType.SUBTITLE) {
      ph.asShape().getText().setText(content.subtitle || '');
    }
  });

  // 本文スライド
  (content.slides || []).forEach(function (s) {
    var slide = pres.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
    var phs = slide.getPlaceholders();
    phs.forEach(function (ph) {
      var t = ph.asShape().getPlaceholderType();
      if (t === SlidesApp.PlaceholderType.TITLE) {
        ph.asShape().getText().setText(s.heading || '');
      } else if (t === SlidesApp.PlaceholderType.BODY) {
        ph.asShape().getText().setText((s.bullets || []).map(function (x) { return '• ' + x; }).join('\n'));
      }
    });
  });
  pres.saveAndClose();

  var file = DriveApp.getFileById(pres.getId());
  if (folderId) {
    var folder = DriveApp.getFolderById(folderId);
    folder.addFile(file);
    try { DriveApp.getRootFolder().removeFile(file); } catch (e) {}
  }
  // PDF書き出し
  var pdfBlob = file.getAs('application/pdf').setName(name + '.pdf');
  var pdfFile = (folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder()).createFile(pdfBlob);

  return { slideUrl: file.getUrl(), pdfUrl: pdfFile.getUrl() };
}
