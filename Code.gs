/**
 * 製品図鑑・マニュアル管理 — GAS バックエンド
 * スプレッドシートをDBとして使用
 * 全データを1回のAPIコールで取得・保存するシンプル設計
 * （旧「BOM Pro」部品表管理システムを土台に、部品BOM関連機能を廃止して作り替え）
 */

// ── シート名定数 ──
var S = {
  MODELS:  '機種図鑑',
  BOARDS:  '基板図鑑',
  MANUALS: 'マニュアル',
  FLOWS:   'フロー'
};

// ── プロパティ ──
function getSetting(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

// ── doGet ──
function doGet() {
  try {
    var userEmail = Session.getActiveUser().getEmail();
    var tmpl = HtmlService.createTemplateFromFile('Index');
    tmpl.userEmail = userEmail;
    return tmpl.evaluate()
      .setTitle('製品図鑑・マニュアル管理')
      .addMetaTag('viewport','width=device-width,initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch(e) {
    return HtmlService.createHtmlOutput(
      '<body style="font-family:sans-serif;padding:40px;background:#f7f7f8;">' +
      '<h2 style="color:#dc2626;">⚠ 起動エラー</h2>' +
      '<p style="margin-top:12px;color:#52525b;">' + e.message + '</p>' +
      '<hr style="margin:20px 0;border:none;border-top:1px solid #e2e2e5;"/>' +
      '<p style="font-size:13px;color:#a1a1aa;line-height:1.8;">確認事項：<br>' +
      '① GAS「プロジェクトの設定」→「スクリプトプロパティ」に <b>BOARD_SS_ID</b>（連携するスプレッドシートのID）を設定する<br>' +
      '② 再デプロイ（新バージョン）する</p></body>'
    ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── スプレッドシート取得 ──
// プロパティ名は旧システム（BOM Pro）から継続利用（既存デプロイ環境の設定を活かすため）
function _ss() {
  var id = getSetting('BOARD_SS_ID');
  if (!id) throw new Error('スクリプトプロパティ BOARD_SS_ID が未設定です');
  return SpreadsheetApp.openById(id);
}

function _sheet(name, headers) {
  var ss = _ss();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers && headers.length) sh.appendRow(headers);
  }
  return sh;
}

// シート → オブジェクト配列
function _read(name, headers) {
  var sh = _sheet(name, headers);
  var vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  var h = vals[0].map(String);
  return vals.slice(1).map(function(row) {
    var obj = {};
    h.forEach(function(k,i){ if(k) obj[k] = (row[i] !== null && row[i] !== undefined) ? row[i] : ''; });
    return obj;
  });
}

// オブジェクト配列 → シートを丸ごと上書き
function _write(name, headers, rows) {
  var sh = _sheet(name, headers);
  sh.clearContents();
  sh.appendRow(headers);
  if (rows.length > 0) {
    var data = rows.map(function(obj) {
      return headers.map(function(h){ var v = obj[h]; return (v === null || v === undefined) ? '' : v; });
    });
    sh.getRange(2, 1, data.length, headers.length).setValues(data);
  }
}

// ── ヘッダー定義 ──
var H = {
  MODELS:  ['機種コード','機種名','種類','ブランド','発売日','M基板','D基板','DE基板','E基板','C基板','S基板','写真URL','概要','特徴・ポイント','学習メモ','関連マニュアルID','備考','更新日時'],
  BOARDS:  ['基板ID','基板名','分類','バージョン','ステータス','写真URL','機能概要','主要部品','学習ポイント','関連機種コード','備考','更新日時'],
  MANUALS: ['マニュアルID','タイトル','カテゴリ','対象システム','ファイルURL','概要','タグ','更新者','更新日時'],
  FLOWS:   ['フローID','タイトル','カテゴリ','概要','ステップ内容','図解URL','関連マニュアルID','備考','更新日時']
  // 「ステップ内容」列には構造化ステップ配列のJSON文字列を保存する（[{no,type,title,detail,owner,branchYes,branchNo,nextNo,manualIds,refLink,done,note}, ...]）
  // JSONとして読めない場合はフロント側で1ステップの自由テキストとしてフォールバック表示する
};

// ══════════════════════════════════════════════════
// メインAPI: 全データを1回で取得
// ══════════════════════════════════════════════════
function apiLoadAll() {
  return _wrap(function() {
    var models = {}, boards = {}, manuals = {}, flows = {};

    _read(S.MODELS, H.MODELS).forEach(function(r) {
      if (r['機種コード']) models[r['機種コード']] = {
        code: String(r['機種コード']),
        name: String(r['機種名'] || ''),
        type: String(r['種類'] || ''),
        brand: String(r['ブランド'] || ''),
        date: String(r['発売日'] || ''),
        m:  String(r['M基板'] || ''),
        d:  String(r['D基板'] || ''),
        de: String(r['DE基板'] || ''),
        e:  String(r['E基板'] || ''),
        c:  String(r['C基板'] || ''),
        s:  String(r['S基板'] || ''),
        photoUrl: String(r['写真URL'] || ''),
        summary: String(r['概要'] || ''),
        features: String(r['特徴・ポイント'] || ''),
        studyNote: String(r['学習メモ'] || ''),
        manualIds: String(r['関連マニュアルID'] || ''),
        note: String(r['備考'] || ''),
        updatedAt: String(r['更新日時'] || '')
      };
    });

    _read(S.BOARDS, H.BOARDS).forEach(function(r) {
      if (r['基板ID']) boards[r['基板ID']] = {
        id: String(r['基板ID']),
        name: String(r['基板名'] || ''),
        category: String(r['分類'] || ''),
        version: String(r['バージョン'] || ''),
        status: String(r['ステータス'] || ''),
        photoUrl: String(r['写真URL'] || ''),
        summary: String(r['機能概要'] || ''),
        mainParts: String(r['主要部品'] || ''),
        studyPoint: String(r['学習ポイント'] || ''),
        relatedModels: String(r['関連機種コード'] || ''),
        note: String(r['備考'] || ''),
        updatedAt: String(r['更新日時'] || '')
      };
    });

    _read(S.MANUALS, H.MANUALS).forEach(function(r) {
      if (r['マニュアルID']) manuals[r['マニュアルID']] = {
        id: String(r['マニュアルID']),
        title: String(r['タイトル'] || ''),
        category: String(r['カテゴリ'] || ''),
        targetSystem: String(r['対象システム'] || ''),
        fileUrl: String(r['ファイルURL'] || ''),
        summary: String(r['概要'] || ''),
        tags: String(r['タグ'] || ''),
        updatedBy: String(r['更新者'] || ''),
        updatedAt: String(r['更新日時'] || '')
      };
    });

    _read(S.FLOWS, H.FLOWS).forEach(function(r) {
      if (r['フローID']) flows[r['フローID']] = {
        id: String(r['フローID']),
        title: String(r['タイトル'] || ''),
        category: String(r['カテゴリ'] || ''),
        summary: String(r['概要'] || ''),
        steps: String(r['ステップ内容'] || ''),
        diagramUrl: String(r['図解URL'] || ''),
        manualIds: String(r['関連マニュアルID'] || ''),
        note: String(r['備考'] || ''),
        updatedAt: String(r['更新日時'] || '')
      };
    });

    return { models: models, boards: boards, manuals: manuals, flows: flows };
  });
}

// ── 機種図鑑 保存 ──
function apiSaveModels(modelsObj) {
  return _wrap(function() {
    var rows = Object.values(modelsObj).map(function(m) {
      return {
        '機種コード': m.code, '機種名': m.name || '', '種類': m.type || '', 'ブランド': m.brand || '',
        '発売日': m.date || '', 'M基板': m.m || '', 'D基板': m.d || '', 'DE基板': m.de || '',
        'E基板': m.e || '', 'C基板': m.c || '', 'S基板': m.s || '',
        '写真URL': m.photoUrl || '', '概要': m.summary || '', '特徴・ポイント': m.features || '',
        '学習メモ': m.studyNote || '', '関連マニュアルID': m.manualIds || '', '備考': m.note || '',
        '更新日時': m.updatedAt || ''
      };
    });
    _write(S.MODELS, H.MODELS, rows);
    return { saved: rows.length };
  });
}

// ── 基板図鑑 保存 ──
function apiSaveBoards(boardsObj) {
  return _wrap(function() {
    var rows = Object.values(boardsObj).map(function(b) {
      return {
        '基板ID': b.id, '基板名': b.name || '', '分類': b.category || '', 'バージョン': b.version || '',
        'ステータス': b.status || '', '写真URL': b.photoUrl || '', '機能概要': b.summary || '',
        '主要部品': b.mainParts || '', '学習ポイント': b.studyPoint || '', '関連機種コード': b.relatedModels || '',
        '備考': b.note || '', '更新日時': b.updatedAt || ''
      };
    });
    _write(S.BOARDS, H.BOARDS, rows);
    return { saved: rows.length };
  });
}

// ── マニュアル 保存 ──
function apiSaveManuals(manualsObj) {
  return _wrap(function() {
    var rows = Object.values(manualsObj).map(function(mn) {
      return {
        'マニュアルID': mn.id, 'タイトル': mn.title || '', 'カテゴリ': mn.category || '',
        '対象システム': mn.targetSystem || '', 'ファイルURL': mn.fileUrl || '', '概要': mn.summary || '',
        'タグ': mn.tags || '', '更新者': mn.updatedBy || '', '更新日時': mn.updatedAt || ''
      };
    });
    _write(S.MANUALS, H.MANUALS, rows);
    return { saved: rows.length };
  });
}

// ── フロー 保存 ──
function apiSaveFlows(flowsObj) {
  return _wrap(function() {
    var rows = Object.values(flowsObj).map(function(f) {
      return {
        'フローID': f.id, 'タイトル': f.title || '', 'カテゴリ': f.category || '',
        '概要': f.summary || '', 'ステップ内容': f.steps || '', '図解URL': f.diagramUrl || '',
        '関連マニュアルID': f.manualIds || '', '備考': f.note || '', '更新日時': f.updatedAt || ''
      };
    });
    _write(S.FLOWS, H.FLOWS, rows);
    return { saved: rows.length };
  });
}

// ── アップロード先フォルダ（未設定なら自動作成して記憶） ──
function _uploadFolder() {
  var id = getSetting('UPLOAD_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch(e) { /* フォールバックへ */ }
  }
  var it = DriveApp.getFoldersByName('図鑑・マニュアル_アップロード');
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder('図鑑・マニュアル_アップロード');
  PropertiesService.getScriptProperties().setProperty('UPLOAD_FOLDER_ID', folder.getId());
  return folder;
}

// ── ファイルアップロード（写真・マニュアルPDF等） ──
function apiUploadFile(p) {
  return _wrap(function() {
    if (!p.base64Data || !p.fileName) throw new Error('ファイルデータ不足');
    var mimeType = p.mimeType || 'application/octet-stream';
    var blob = Utilities.newBlob(Utilities.base64Decode(p.base64Data), mimeType, p.fileName);
    var file = _uploadFolder().createFile(blob);
    return { url: file.getUrl(), fileId: file.getId(), fileName: p.fileName };
  });
}

// ══════════════════════════════════════════════════
// テンプレート（雛形）
// フロー: Googleスプレッドシート雛形 → ネイティブexport URLで.xlsxとしてダウンロード可能
// マニュアル: Googleドキュメント/スライド雛形 → 同様に.docx/.pptxでダウンロード可能
// ══════════════════════════════════════════════════
var FLOW_TEMPLATE_HEADERS = ['No','種別','タイトル','詳細内容','担当','分岐:Yes条件','分岐:No条件','次のステップNo(分岐時)','関連マニュアルID','参考リンク','備考'];

function _shareViewable(file) {
  try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); }
  catch(e) { try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e2) {} }
}

// 新規作成したファイルをマイドライブ直下からアップロード共有フォルダへ移動
function _relocateToUploadFolder(file) {
  var parents = file.getParents();
  while (parents.hasNext()) { parents.next().removeFile(file); }
  _uploadFolder().addFile(file);
}

function _ensureFlowTemplate() {
  var id = getSetting('TPL_FLOW_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch(e) { /* 削除済みなら作り直す */ } }
  var ss = SpreadsheetApp.create('【雛形】業務フロー');
  var sh = ss.getSheets()[0];
  sh.setName('フロー');
  sh.appendRow(FLOW_TEMPLATE_HEADERS);
  sh.appendRow(['1','ステップ','客先から見積依頼を受付','見積依頼メールの件名・添付を確認する','営業担当','','','','','','記入例：1行＝1ステップ。分岐する場合は「種別」を「分岐」にし、Yes/No条件と分岐先Noを入力']);
  sh.appendRow(['2','分岐','対象基板は新規設計か？','過去に見積提出の実績があるか確認する','営業担当','新規設計の場合は3へ','既存流用の場合は5へ','3 / 5','','','']);
  sh.getRange(1,1,1,FLOW_TEMPLATE_HEADERS.length).setFontWeight('bold').setBackground('#eef2ff');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, FLOW_TEMPLATE_HEADERS.length);
  var noteSheet = ss.insertSheet('記入方法');
  noteSheet.getRange('A1').setValue(
    '■ 業務フロー テンプレートの使い方\n\n' +
    '・1行 = 1ステップ（または1分岐）として記入してください。\n' +
    '・種別は「ステップ」または「分岐」のどちらかを入力してください。\n' +
    '・「分岐」の場合は Yes/No それぞれの条件と、進む先のステップNoを入力してください。\n' +
    '・関連マニュアルIDは、マニュアル管理画面で確認できるIDをカンマ区切りで入力してください（空欄でも可）。\n' +
    '・記入が終わったら、このシートのURLをコピーし、「フロー管理」→「スプレッドシートから取り込む」に貼り付けてください。'
  ).setWrap(true);
  noteSheet.setColumnWidth(1, 600);
  var file = DriveApp.getFileById(ss.getId());
  _relocateToUploadFolder(file);
  _shareViewable(file);
  PropertiesService.getScriptProperties().setProperty('TPL_FLOW_ID', ss.getId());
  return ss;
}

function _ensureManualDocTemplate() {
  var id = getSetting('TPL_MANUAL_DOC_ID');
  if (id) { try { return DocumentApp.openById(id); } catch(e) { /* 削除済みなら作り直す */ } }
  var doc = DocumentApp.create('【雛形】マニュアル（Word）');
  var body = doc.getBody();
  body.appendParagraph('マニュアルタイトル').setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph('（ここにマニュアルのタイトルを記入してください）');
  body.appendParagraph('概要').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('このマニュアルの目的・対象範囲を記入してください。');
  body.appendParagraph('対象者').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('例：営業担当、新入社員 など');
  body.appendParagraph('手順').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('① 手順1のタイトル').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('手順の詳細を記入してください。スクリーンショットも貼り付けられます。');
  body.appendParagraph('② 手順2のタイトル').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('手順の詳細を記入してください。');
  body.appendParagraph('注意事項').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('間違えやすいポイント、注意点を記入してください。');
  body.appendParagraph('更新履歴').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendTable([['日付','更新者','内容'],['2026/01/01','（記入例）','新規作成']]);
  doc.saveAndClose();
  var file = DriveApp.getFileById(doc.getId());
  _relocateToUploadFolder(file);
  _shareViewable(file);
  PropertiesService.getScriptProperties().setProperty('TPL_MANUAL_DOC_ID', doc.getId());
  return doc;
}

function _ensureManualSlidesTemplate() {
  var id = getSetting('TPL_MANUAL_SLIDES_ID');
  if (id) { try { return SlidesApp.openById(id); } catch(e) { /* 削除済みなら作り直す */ } }
  var pres = SlidesApp.create('【雛形】マニュアル（スライド）');
  var first = pres.getSlides()[0];
  var t1 = first.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
  if (t1) t1.asShape().getText().setText('マニュアルタイトル');
  var b1 = first.getPlaceholder(SlidesApp.PlaceholderType.BODY) || first.getPlaceholder(SlidesApp.PlaceholderType.SUBTITLE);
  if (b1) b1.asShape().getText().setText('対象システム／対象者をここに記入');

  [['① 手順1のタイトル','手順の詳細・スクリーンショットをここに追加してください。'],
   ['② 手順2のタイトル','手順の詳細をここに記入してください。'],
   ['まとめ・注意事項','間違えやすいポイントや注意点をここに記入してください。']
  ].forEach(function(pair) {
    var s = pres.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
    var t = s.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
    if (t) t.asShape().getText().setText(pair[0]);
    var b = s.getPlaceholder(SlidesApp.PlaceholderType.BODY);
    if (b) b.asShape().getText().setText(pair[1]);
  });

  var file = DriveApp.getFileById(pres.getId());
  _relocateToUploadFolder(file);
  _shareViewable(file);
  PropertiesService.getScriptProperties().setProperty('TPL_MANUAL_SLIDES_ID', pres.getId());
  return pres;
}

function _ensureTemplates() {
  _ensureFlowTemplate();
  _ensureManualDocTemplate();
  _ensureManualSlidesTemplate();
}

// ── テンプレート情報取得（ダウンロードURL・編集URL） ──
function apiGetTemplateInfo() {
  return _wrap(function() {
    _ensureTemplates();
    var flowId = getSetting('TPL_FLOW_ID');
    var docId = getSetting('TPL_MANUAL_DOC_ID');
    var slidesId = getSetting('TPL_MANUAL_SLIDES_ID');
    return {
      flow:         { name:'【雛形】業務フロー',            editUrl:'https://docs.google.com/spreadsheets/d/'+flowId+'/edit',    downloadUrl:'https://docs.google.com/spreadsheets/d/'+flowId+'/export?format=xlsx' },
      manualDoc:    { name:'【雛形】マニュアル（Word）',     editUrl:'https://docs.google.com/document/d/'+docId+'/edit',        downloadUrl:'https://docs.google.com/document/d/'+docId+'/export?format=docx' },
      manualSlides: { name:'【雛形】マニュアル（スライド）', editUrl:'https://docs.google.com/presentation/d/'+slidesId+'/edit', downloadUrl:'https://docs.google.com/presentation/d/'+slidesId+'/export/pptx' }
    };
  });
}

// ── テンプレートをコピーして自分用の編集可能ファイルを作成 ──
function apiCopyTemplate(p) {
  return _wrap(function() {
    _ensureTemplates();
    var idKey = p.kind === 'flow' ? 'TPL_FLOW_ID' : p.kind === 'manualDoc' ? 'TPL_MANUAL_DOC_ID' : p.kind === 'manualSlides' ? 'TPL_MANUAL_SLIDES_ID' : null;
    if (!idKey) throw new Error('不明なテンプレート種別: ' + p.kind);
    var srcFile = DriveApp.getFileById(getSetting(idKey));
    var baseName = (p.name || srcFile.getName().replace('【雛形】', '')).trim();
    var copyName = baseName + '_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmm');
    var copy = srcFile.makeCopy(copyName, _uploadFolder());
    return { url: copy.getUrl(), name: copy.getName(), id: copy.getId() };
  });
}

// URL または ID からスプレッドシートIDを取り出す
function _extractSheetId(url) {
  var m = String(url || '').match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  var trimmed = String(url || '').trim();
  if (/^[a-zA-Z0-9_-]{20,}$/.test(trimmed)) return trimmed;
  throw new Error('スプレッドシートのURLが正しくありません');
}

// ── 記入済みのフロー用スプレッドシートから構造化ステップを取り込む（保存はフロント側で実施） ──
function apiImportFlowFromSheet(p) {
  return _wrap(function() {
    var id = _extractSheetId(p.sheetUrl);
    var ss = SpreadsheetApp.openById(id);
    var sh = ss.getSheets()[0];
    var vals = sh.getDataRange().getValues();
    if (vals.length < 2) throw new Error('シートにデータがありません');
    var headers = vals[0].map(String);
    var idx = {};
    FLOW_TEMPLATE_HEADERS.forEach(function(h){ idx[h] = headers.indexOf(h); });
    if (idx['No'] < 0 || idx['タイトル'] < 0) throw new Error('テンプレートの列見出しと一致しません（No・タイトル列が見つかりません）。テンプレートをコピーして使用してください。');
    var steps = [];
    vals.slice(1).forEach(function(row) {
      var no = row[idx['No']], title = row[idx['タイトル']];
      if (!no && !title) return; // 空行はスキップ
      var g = function(key){ return idx[key] >= 0 ? String(row[idx[key]] || '') : ''; };
      steps.push({
        no: String(no || ''),
        type: g('種別') === '分岐' ? 'branch' : 'step',
        title: String(title || ''),
        detail: g('詳細内容'),
        owner: g('担当'),
        branchYes: g('分岐:Yes条件'),
        branchNo: g('分岐:No条件'),
        nextNo: g('次のステップNo(分岐時)'),
        manualIds: g('関連マニュアルID'),
        refLink: g('参考リンク'),
        note: g('備考'),
        done: false
      });
    });
    if (steps.length === 0) throw new Error('取り込めるステップが見つかりませんでした');
    return { steps: steps, sourceName: ss.getName() };
  });
}

// ── エラーラッパー ──
function _wrap(fn) {
  try {
    var r = fn();
    r.success = true;
    return r;
  } catch(e) {
    Logger.log('ERROR: ' + e.message + '\n' + e.stack);
    return { success: false, error: e.message };
  }
}
