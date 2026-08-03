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
  MODELS:  ['機種コード','機種名','種類','ブランド','遊技機タイプ','発売日','M基板','D基板','DE基板','E基板','C基板','S基板','写真URL','概要','特徴・ポイント','学習メモ','関連マニュアルID','備考','更新日時'],
  BOARDS:  ['基板ID','基板名','分類','ブランド','バージョン','ステータス','写真URL','機能概要','主要部品','学習ポイント','関連機種コード','備考','更新日時'],
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
        gameType: String(r['遊技機タイプ'] || ''),
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
        brand: String(r['ブランド'] || ''),
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
        '遊技機タイプ': m.gameType || '',
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
        '基板ID': b.id, '基板名': b.name || '', '分類': b.category || '', 'ブランド': b.brand || '', 'バージョン': b.version || '',
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

var MODEL_TEMPLATE_HEADERS = ['機種コード','機種名','ブランド','遊技機タイプ','主基板(M)','図柄制御基板(D)','液晶IF基板(DE)','演出基板(E)','払出制御基板(C)','S基板','備考'];
var BOARD_TEMPLATE_HEADERS = ['基板ID','分類','ブランド','状態','採用機種','見積書タイトル','見積書ファイルURL','部品構成URL','仕掛基板URL','PCB基板URL','価格','備考'];

function _ensureModelTemplate() {
  var id = getSetting('TPL_MODEL_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch(e) { /* 削除済みなら作り直す */ } }
  var ss = SpreadsheetApp.create('【雛形】機種・基板構成');
  var sh = ss.getSheets()[0];
  sh.setName('機種構成');
  sh.appendRow(MODEL_TEMPLATE_HEADERS);
  sh.appendRow(['E33','世界で一番','SF20','回胴','M1602B','D1401C','DE1607A','E1501C','','','記入例：機種別ハードウェア構成一覧表から書き写してください']);
  sh.appendRow(['A72','とある禁書目録','PF40','パチンコ','M1904B(共通①)','D1401C','DE1802A','E1403C','','','']);
  sh.getRange(1,1,1,MODEL_TEMPLATE_HEADERS.length).setFontWeight('bold').setBackground('#eef2ff');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, MODEL_TEMPLATE_HEADERS.length);
  var noteSheet = ss.insertSheet('記入方法');
  noteSheet.getRange('A1').setValue(
    '■ 機種・基板構成 テンプレートの使い方\n\n' +
    '・1行 = 1機種として記入してください。\n' +
    '・「遊技機タイプ」には「パチンコ」または「回胴」を入力してください。\n' +
    '・主基板(M)／図柄制御基板(D)／液晶IF基板(DE)／演出基板(E)／払出制御基板(C)／S基板には、\n' +
    '　機種別ハードウェア構成一覧表に記載の基板コードを記入してください（不明な列は空欄でOK）。\n' +
    '・記入が終わったら、このシートのURLをコピーし、「機種図鑑」→「スプレッドシートから取り込む」に貼り付けてください。'
  ).setWrap(true);
  noteSheet.setColumnWidth(1, 600);
  var file = DriveApp.getFileById(ss.getId());
  _relocateToUploadFolder(file);
  _shareViewable(file);
  PropertiesService.getScriptProperties().setProperty('TPL_MODEL_ID', ss.getId());
  return ss;
}

function _ensureBoardTemplate() {
  var id = getSetting('TPL_BOARD_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch(e) { /* 削除済みなら作り直す */ } }
  var ss = SpreadsheetApp.create('【雛形】基板マスタ');
  var sh = ss.getSheets()[0];
  sh.setName('基板マスタ');
  sh.appendRow(BOARD_TEMPLATE_HEADERS);
  sh.appendRow(['M2003A2','製品','FF+','新品','D56','','https://drive.google.com/...','https://docs.google.com/...','','','10000','記入例：部品マスタ表から書き写してください']);
  sh.getRange(1,1,1,BOARD_TEMPLATE_HEADERS.length).setFontWeight('bold').setBackground('#eef2ff');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, BOARD_TEMPLATE_HEADERS.length);
  var noteSheet = ss.insertSheet('記入方法');
  noteSheet.getRange('A1').setValue(
    '■ 基板マスタ テンプレートの使い方\n\n' +
    '・1行 = 1基板として記入してください（基板IDが重複する行は後の行で上書きされます）。\n' +
    '・「採用機種」には、その基板が使われている機種コードをカンマ区切りで入力してください。\n' +
    '・見積書タイトル／各URL／価格は、参考情報として基板図鑑の備考欄にまとめて取り込まれます。\n' +
    '・記入が終わったら、このシートのURLをコピーし、「基板図鑑」→「スプレッドシートから取り込む」に貼り付けてください。'
  ).setWrap(true);
  noteSheet.setColumnWidth(1, 600);
  var file = DriveApp.getFileById(ss.getId());
  _relocateToUploadFolder(file);
  _shareViewable(file);
  PropertiesService.getScriptProperties().setProperty('TPL_BOARD_ID', ss.getId());
  return ss;
}

function _ensureTemplates() {
  _ensureFlowTemplate();
  _ensureManualDocTemplate();
  _ensureManualSlidesTemplate();
  _ensureModelTemplate();
  _ensureBoardTemplate();
}

// ── テンプレート情報取得（ダウンロードURL・編集URL） ──
function apiGetTemplateInfo() {
  return _wrap(function() {
    _ensureTemplates();
    var flowId = getSetting('TPL_FLOW_ID');
    var docId = getSetting('TPL_MANUAL_DOC_ID');
    var slidesId = getSetting('TPL_MANUAL_SLIDES_ID');
    var modelId = getSetting('TPL_MODEL_ID');
    var boardId = getSetting('TPL_BOARD_ID');
    return {
      flow:         { name:'【雛形】業務フロー',            editUrl:'https://docs.google.com/spreadsheets/d/'+flowId+'/edit',    downloadUrl:'https://docs.google.com/spreadsheets/d/'+flowId+'/export?format=xlsx' },
      manualDoc:    { name:'【雛形】マニュアル（Word）',     editUrl:'https://docs.google.com/document/d/'+docId+'/edit',        downloadUrl:'https://docs.google.com/document/d/'+docId+'/export?format=docx' },
      manualSlides: { name:'【雛形】マニュアル（スライド）', editUrl:'https://docs.google.com/presentation/d/'+slidesId+'/edit', downloadUrl:'https://docs.google.com/presentation/d/'+slidesId+'/export/pptx' },
      model:        { name:'【雛形】機種・基板構成',        editUrl:'https://docs.google.com/spreadsheets/d/'+modelId+'/edit',   downloadUrl:'https://docs.google.com/spreadsheets/d/'+modelId+'/export?format=xlsx' },
      board:        { name:'【雛形】基板マスタ',            editUrl:'https://docs.google.com/spreadsheets/d/'+boardId+'/edit',   downloadUrl:'https://docs.google.com/spreadsheets/d/'+boardId+'/export?format=xlsx' }
    };
  });
}

// ── テンプレートをコピーして自分用の編集可能ファイルを作成 ──
function apiCopyTemplate(p) {
  return _wrap(function() {
    _ensureTemplates();
    var idKeyMap = { flow:'TPL_FLOW_ID', manualDoc:'TPL_MANUAL_DOC_ID', manualSlides:'TPL_MANUAL_SLIDES_ID', model:'TPL_MODEL_ID', board:'TPL_BOARD_ID' };
    var idKey = idKeyMap[p.kind];
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

// ── 記入済みの機種構成スプレッドシートから機種一覧を取り込む ──
function apiImportModelsFromSheet(p) {
  return _wrap(function() {
    var id = _extractSheetId(p.sheetUrl);
    var ss = SpreadsheetApp.openById(id);
    var sh = ss.getSheets()[0];
    var vals = sh.getDataRange().getValues();
    if (vals.length < 2) throw new Error('シートにデータがありません');
    var headers = vals[0].map(String);
    var idx = {};
    MODEL_TEMPLATE_HEADERS.forEach(function(h){ idx[h] = headers.indexOf(h); });
    if (idx['機種コード'] < 0) throw new Error('テンプレートの列見出しと一致しません（機種コード列が見つかりません）。テンプレートをコピーして使用してください。');
    var g = function(row, key){ return idx[key] >= 0 ? String(row[idx[key]] || '') : ''; };
    var models = [];
    vals.slice(1).forEach(function(row) {
      var code = g(row, '機種コード');
      if (!code) return;
      models.push({
        code: code, name: g(row,'機種名'), brand: g(row,'ブランド'), gameType: g(row,'遊技機タイプ'),
        m: g(row,'主基板(M)'), d: g(row,'図柄制御基板(D)'), de: g(row,'液晶IF基板(DE)'), e: g(row,'演出基板(E)'),
        c: g(row,'払出制御基板(C)'), s: g(row,'S基板'), note: g(row,'備考')
      });
    });
    if (models.length === 0) throw new Error('取り込める機種が見つかりませんでした');
    return { models: models, sourceName: ss.getName() };
  });
}

// ── 記入済みの基板マスタスプレッドシートから基板一覧を取り込む ──
function apiImportBoardsFromSheet(p) {
  return _wrap(function() {
    var id = _extractSheetId(p.sheetUrl);
    var ss = SpreadsheetApp.openById(id);
    var sh = ss.getSheets()[0];
    var vals = sh.getDataRange().getValues();
    if (vals.length < 2) throw new Error('シートにデータがありません');
    var headers = vals[0].map(String);
    var idx = {};
    BOARD_TEMPLATE_HEADERS.forEach(function(h){ idx[h] = headers.indexOf(h); });
    if (idx['基板ID'] < 0) throw new Error('テンプレートの列見出しと一致しません（基板ID列が見つかりません）。テンプレートをコピーして使用してください。');
    var g = function(row, key){ return idx[key] >= 0 ? String(row[idx[key]] || '') : ''; };
    var boards = [];
    vals.slice(1).forEach(function(row) {
      var id2 = g(row, '基板ID');
      if (!id2) return;
      var noteParts = [];
      var state = g(row,'状態'); if (state) noteParts.push('状態: ' + state);
      var qTitle = g(row,'見積書タイトル'); if (qTitle) noteParts.push('見積書: ' + qTitle);
      var price = g(row,'価格'); if (price) noteParts.push('見積価格: ¥' + price);
      ['見積書ファイルURL','部品構成URL','仕掛基板URL','PCB基板URL'].forEach(function(k){
        var v = g(row,k); if (v) noteParts.push(k + ': ' + v);
      });
      var remark = g(row,'備考'); if (remark) noteParts.push(remark);
      boards.push({
        id: id2, name: id2, category: g(row,'分類'), brand: g(row,'ブランド'),
        relatedModels: g(row,'採用機種'), note: noteParts.join('\n')
      });
    });
    if (boards.length === 0) throw new Error('取り込める基板が見つかりませんでした');
    return { boards: boards, sourceName: ss.getName() };
  });
}

// ══════════════════════════════════════════════════
// 実データ投入（部品マスタxlsx・機種別ハードウェア構成表PDFから抽出したデータ）
// 既存データは残したまま、同じキーのみ追加・上書きするマージ方式
// ══════════════════════════════════════════════════
function apiSeedRealData() {
  return _wrap(function() {
    var models = _read(S.MODELS, H.MODELS);
    var boards = _read(S.BOARDS, H.BOARDS);
    var modelMap = {}, boardMap = {};
    models.forEach(function(r){ if (r['機種コード']) modelMap[r['機種コード']] = _rowToModel(r); });
    boards.forEach(function(r){ if (r['基板ID']) boardMap[r['基板ID']] = _rowToBoard(r); });

    SEED_MODELS.forEach(function(m){ modelMap[m.code] = m; });
    SEED_BOARDS.forEach(function(b){ boardMap[b.id] = b; });

    apiSaveModels(modelMap);
    apiSaveBoards(boardMap);
    return { models: SEED_MODELS.length, boards: SEED_BOARDS.length };
  });
}
function _rowToModel(r) {
  return { code:String(r['機種コード']||''), name:String(r['機種名']||''), type:String(r['種類']||''), brand:String(r['ブランド']||''),
    gameType:String(r['遊技機タイプ']||''), date:String(r['発売日']||''), m:String(r['M基板']||''), d:String(r['D基板']||''),
    de:String(r['DE基板']||''), e:String(r['E基板']||''), c:String(r['C基板']||''), s:String(r['S基板']||''),
    photoUrl:String(r['写真URL']||''), summary:String(r['概要']||''), features:String(r['特徴・ポイント']||''),
    studyNote:String(r['学習メモ']||''), manualIds:String(r['関連マニュアルID']||''), note:String(r['備考']||''), updatedAt:String(r['更新日時']||'') };
}
function _rowToBoard(r) {
  return { id:String(r['基板ID']||''), name:String(r['基板名']||''), category:String(r['分類']||''), brand:String(r['ブランド']||''),
    version:String(r['バージョン']||''), status:String(r['ステータス']||''), photoUrl:String(r['写真URL']||''), summary:String(r['機能概要']||''),
    mainParts:String(r['主要部品']||''), studyPoint:String(r['学習ポイント']||''), relatedModels:String(r['関連機種コード']||''),
    note:String(r['備考']||''), updatedAt:String(r['更新日時']||'') };
}

// 機種別ハードウェア構成一覧表(回胴)_260701.pdf / (ぱちんこ)_260701.pdf より抽出（判読できた範囲）
var SEED_MODELS = [
  {code:'E33',name:'世界で一番',brand:'SF20',gameType:'回胴',m:'M1602B',d:'D1401C',de:'DE1607A',e:'E1501C'},
  {code:'E32',name:'リング2',brand:'SF20',gameType:'回胴',m:'M1602B',d:'D1401C',de:'DE1607A',e:'E1501C'},
  {code:'E34',name:'地獄少女2',brand:'SF20',gameType:'回胴',m:'M1602B',d:'D1401C',de:'DE1607A',e:'E1501C'},
  {code:'E35',name:'FT',brand:'SF20',gameType:'回胴',m:'M1602B',d:'D1401C',de:'DE1606A',e:'E1501C'},
  {code:'E36',name:'呪怨2',brand:'SF20',gameType:'回胴',m:'M1602B',d:'D1401C',de:'DE1607A',e:'E1501C'},
  {code:'E36B',name:'呪怨2',brand:'SF20',gameType:'回胴',m:'M1805D',d:'D1401C',de:'DE1607A',e:'E1501C',note:'E36のリユース版'},
  {code:'E37',name:'喰霊',brand:'SF20',gameType:'回胴',m:'M1805D',d:'D1401C',de:'DE1802A',e:'E1501C'},
  {code:'E39',name:'リング3',brand:'SF20',gameType:'回胴',m:'M1805D',d:'D1401C',de:'DE1607A',e:'E1501C'},
  {code:'E40',name:'地獄少女3',brand:'SF20',gameType:'回胴',m:'M1805D',d:'D1401C',de:'DE1802A',e:'E1501C'},
  {code:'E40B',name:'地獄3高純増',brand:'SF20',gameType:'回胴',m:'M1805D',d:'D1401C',de:'DE1802A',e:'E1501C'},
  {code:'E41',name:'リング4',brand:'SF20',gameType:'回胴',m:'M1805D',d:'D1401C',de:'DE1802A',e:'E1501C'},
  {code:'E42',name:'フェアリーテイル2',brand:'SF20',gameType:'回胴',m:'M1805D(代)VerB',d:'D2101A',de:'DE1606A',e:'E1902B'},
  {code:'E43',name:'レールガン',brand:'SF20',gameType:'回胴',m:'M1805D(代)VerC',d:'D2101A',de:'DE1606A',e:'E1902B'},
  {code:'E44',name:'アリア2',brand:'SF20',gameType:'回胴',m:'M1805D',d:'D1401C',de:'DE1802A',e:'E1902B'},
  {code:'E45',name:'ゴブリンスレイヤー',brand:'SFK10',gameType:'回胴',m:'M2104A',d:'D1401C',de:'DE1802A',e:'E2102B'},
  {code:'E46B',name:'とある魔術の禁書目録',brand:'SFK15',gameType:'回胴',m:'M2104A1(リセットIC対応)',d:'D1401C',de:'DE1802A',e:'E1902B'},

  {code:'A72',name:'とある禁書目録',brand:'PF40',gameType:'パチンコ',m:'M1904B(共通①)',d:'D1401C',de:'DE1802A',e:'E1403C'},
  {code:'A72W',name:'とある禁書目録',brand:'PF40',gameType:'パチンコ',m:'M1904B(共通①)',d:'D2101A(旭化成対応)',de:'DE1802A',e:'E1901B'},
  {code:'A73',name:'暴れん坊8',brand:'PF40',gameType:'パチンコ',m:'M1904B(共通①)',d:'D1401C',de:'DE1607A',e:'E1403C'},
  {code:'A73W',name:'暴れん坊9',brand:'PF40',gameType:'パチンコ',m:'M1904B(共通①)',d:'D1401C',de:'DE1607A',e:'E1403C'},
  {code:'A74',name:'アリア4',brand:'PF40',gameType:'パチンコ',m:'M2002B',d:'D1901B(代)',de:'DE1802A',e:'E1403C'},
  {code:'A74(ミドル)',name:'アリア4',brand:'PF40',gameType:'パチンコ',m:'M2002B(代)VerB',d:'D1901B',de:'DE1802A',e:'E1403C'},
  {code:'A74W',name:'アリア4',brand:'PF40',gameType:'パチンコ',m:'M2002B',d:'D1901B',de:'DE1802A',e:'E1403C'},
  {code:'A70B',name:'地獄4.5',brand:'PF40',gameType:'パチンコ',m:'M1904B(共通①)',d:'D1401C',de:'DE1802A',e:'E1403C'},
  {code:'C41',name:'どないやねんDX',brand:'PF40',gameType:'パチンコ',m:'M1904B(共通①)',d:'D1201B',de:'DE1901A',e:'L1302A'},
  {code:'C42',name:'アレジン',brand:'PF40',gameType:'パチンコ',m:'M2003A',e:'L1302A',note:'液晶制御基板/液晶IF基板は原本の表記が不明瞭のため未取り込み'},
  {code:'D53',name:'地獄少女5',brand:'PF40',gameType:'パチンコ',m:'M2003A',d:'D2101A',de:'DE2101A',e:'E1901B'},
  {code:'D53H',name:'地獄少女5',brand:'PF40',gameType:'パチンコ',m:'M2003A',d:'D2101A',de:'DE2101A',e:'E1901B'},
  {code:'D53W',name:'地獄少女5',brand:'PF40',gameType:'パチンコ',m:'M2003A',d:'D2101A',de:'DE2101A',e:'E1901B'},
  {code:'D52',name:'とある超電磁砲',brand:'PF40',gameType:'パチンコ',m:'M2003A',d:'D2101A',de:'DE2101A(代)VerB',e:'E1901B'},
  {code:'D52W',name:'とある超電磁砲',brand:'PF40',gameType:'パチンコ',m:'M2003A(代)VerB',d:'D2101A',de:'DE2101A',e:'E1901B'},
  {code:'A76',name:'ストリートファイター',brand:'PF40',gameType:'パチンコ',m:'M2003A',d:'D2101A',de:'DE1607A',e:'E1901B'},
  {code:'A76B',name:'ストリートファイター',brand:'PF40',gameType:'パチンコ',m:'M2003A',d:'D2101A',de:'DE1607A',e:'E1901B'},
  {code:'A76W',name:'ストリートファイター',brand:'PF40',gameType:'パチンコ',m:'M2003A',d:'D2101A',de:'DE1607A',e:'E2002B'},
  {code:'D55',name:'サラリーマン金太郎',brand:'PF40',gameType:'パチンコ',m:'M2003A(代)VerB',d:'D2101A',de:'DE1802A',e:'E2002B'},
  {code:'D55L(D55サブ)',name:'サラリーマン金太郎',brand:'PF40',gameType:'パチンコ',m:'M2003A(代)VerB',d:'D2101A',de:'DE1802A',e:'E2002B',note:'D55のサブ機'}
].map(function(m){ return Object.assign({type:'',date:'',photoUrl:'',summary:'',features:'',studyNote:'',manualIds:'',c:'',s:'',note:'',updatedAt:_now_()}, m); });

// マスタ表（アミューズメント事業部、遊技機一括管理用アプリ）.xlsx「⑦製品マスタ」より抽出・重複統合
var SEED_BOARDS = [
  {id:'C1601E',name:'C1601E',category:''},
  {id:'C2101B',name:'C2101B',category:''},
  {id:'C2202B',name:'C2202B',category:''},
  {id:'C2401A',name:'C2401A',category:''},
  {id:'C2501B',name:'C2501B',category:''},
  {id:'D1401C',name:'D1401C',category:''},
  {id:'D2101A',name:'D2101A',category:''},
  {id:'D2101A1',name:'D2101A1',category:''},
  {id:'DE1802A',name:'DE1802A',category:''},
  {id:'DE2101A',name:'DE2101A',category:''},
  {id:'DE2101A1',name:'DE2101A1',category:''},
  {id:'DE2103AverB',name:'DE2103AverB',category:''},
  {id:'DE2502A',name:'DE2502A',category:'回路設計',brand:'FJ+',note:'状態: 新品\n見積価格: ¥2566000\n見積書ファイルURL: https://drive.google.com/file/d/1ZMS6L3YMBinHx4fUdypXJvunO22Shkig/view?usp=sharing'},
  {id:'E2002B',name:'E2002B',category:''},
  {id:'E2101B',name:'E2101B',category:''},
  {id:'E2102B',name:'E2102B',category:''},
  {id:'E2301B',name:'E2301B',category:''},
  {id:'E2501B',name:'E2501B',category:''},
  {id:'E2503A',name:'E2503A',category:'回路設計',brand:'FJ+',note:'状態: 新品\n見積価格: ¥3880000\n見積書ファイルURL: https://drive.google.com/file/d/1qwn7zt18L6CSX_gdfMHL3I2H0bCPORgt/view?usp=sharing'},
  {id:'M2003A',name:'M2003A',category:'仕掛→製品',relatedModels:'A84',note:'状態: 中古(A84向け)→新品(A84向け見積 ¥1,000)\n※同名で複数の案件段階の記録あり'},
  {id:'M2003A1',name:'M2003A1',category:'仕掛',brand:'FF+',note:'状態: 中古'},
  {id:'M2003A2',name:'M2003A2',category:'製品',brand:'FF+',relatedModels:'D56',note:'状態: 新品\n見積価格: ¥10000\n見積書ファイルURL: https://drive.google.com/drive/folders/1vhFIJDGrKab83qewgWQGEAm_krCwY9lD?usp=drive_link\n部品構成URL: https://docs.google.com/spreadsheets/d/1rDTsxnPLEoTBK-_viElyLH-RnybJHQOW/edit?usp=drive_link\n仕掛基板URL: https://drive.google.com/file/d/1ELwPiT6rGDPTZ5muWSRIbfVGj7-zbWNp/view?usp=drive_link'},
  {id:'M2003A5',name:'M2003A5',category:''},
  {id:'M2104A',name:'M2104A',category:''},
  {id:'M2202A',name:'M2202A',category:''},
  {id:'M2401A',name:'M2401A',category:''},
  {id:'M2402A',name:'M2402A',category:''},
  {id:'M2503A',name:'M2503A',category:''},
  {id:'メダル数制御基板(C2101B)',name:'メダル数制御基板(C2101B)',category:'製品',relatedModels:'SJK15',note:'状態: 新品\n見積書: SJK15メダル数制御基板(C2101B)見積り 2023年5月価\n見積書ファイルURL: https://drive.google.com/file/d/1gNthp90V1dsoxCL4r5Cltno2Cb0O8y1b/view?usp=drive_link'}
].map(function(b){ return Object.assign({version:'',status:'',photoUrl:'',summary:'',mainParts:'',studyPoint:'',relatedModels:'',note:'',updatedAt:_now_()}, b); });

function _now_() { return new Date().toISOString(); }

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
