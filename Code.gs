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
  if (id) return SpreadsheetApp.openById(id);
  // 未設定の場合、このスクリプトが紐づくコンテナ（スプレッドシート）を自動検出して記憶する
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    PropertiesService.getScriptProperties().setProperty('BOARD_SS_ID', active.getId());
    return active;
  }
  throw new Error('スクリプトプロパティ BOARD_SS_ID が未設定です');
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
  MODELS:  ['機種コード','機種名','種類','ブランド','遊技機タイプ','発売日','M基板','D基板','DE基板','E基板','C基板','S基板','R基板','関連資料','売上台数','売上目標','リユース・エコ投入台数','在庫数','写真URL','概要','特徴・ポイント','学習メモ','関連マニュアルID','備考','更新日時'],
  BOARDS:  ['基板ID','基板名','種類','分類','ブランド','対応区分','バージョン','ステータス','仕様','改定履歴','関連資料','写真URL','機能概要','主要部品','学習ポイント','関連機種コード','備考','更新日時'],
  MANUALS: ['マニュアルID','タイトル','カテゴリ','対象システム','ファイルURL','概要','タグ','更新者','更新日時'],
  FLOWS:   ['フローID','タイトル','カテゴリ','概要','ステップ内容','図解URL','関連マニュアルID','備考','更新日時']
  // 「ステップ内容」列には構造化ステップ配列のJSON文字列を保存する（[{no,type,title,detail,owner,branchYes,branchNo,nextNo,manualIds,refLink,done,note}, ...]）
  // 基板図鑑の「仕様」列には[{label,value}, ...]、「改定履歴」列には[{version,date,changes}, ...]のJSON文字列を保存する
  // 機種図鑑の「関連資料」列には[{category,title,url,note}, ...]のJSON文字列を保存する（構成表/見積書/組立基準書/生産計画書/納品計画書/部品表など）
  // JSONとして読めない場合はフロント側で1件の自由テキストとしてフォールバック表示する
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
        r:  String(r['R基板'] || ''),
        relatedDocs: String(r['関連資料'] || ''),
        salesUnits: String(r['売上台数'] || ''),
        salesTarget: String(r['売上目標'] || ''),
        reuseEcoUnits: String(r['リユース・エコ投入台数'] || ''),
        stockQty: String(r['在庫数'] || ''),
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
        type: String(r['種類'] || '') || _extractSpecValue(String(r['仕様'] || ''), '基板種別'),
        category: String(r['分類'] || ''),
        brand: String(r['ブランド'] || ''),
        gameType: String(r['対応区分'] || ''),
        version: String(r['バージョン'] || ''),
        status: String(r['ステータス'] || ''),
        specs: String(r['仕様'] || ''),
        revisions: String(r['改定履歴'] || ''),
        relatedDocs: String(r['関連資料'] || ''),
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
        'E基板': m.e || '', 'C基板': m.c || '', 'S基板': m.s || '', 'R基板': m.r || '',
        '関連資料': m.relatedDocs || '', '売上台数': m.salesUnits || '', '売上目標': m.salesTarget || '',
        'リユース・エコ投入台数': m.reuseEcoUnits || '', '在庫数': m.stockQty || '',
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
        '基板ID': b.id, '基板名': b.name || '', '種類': b.type || '', '分類': b.category || '', 'ブランド': b.brand || '',
        '対応区分': b.gameType || '', 'バージョン': b.version || '',
        'ステータス': b.status || '', '仕様': b.specs || '', '改定履歴': b.revisions || '',
        '関連資料': b.relatedDocs || '', '写真URL': b.photoUrl || '', '機能概要': b.summary || '',
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
    _shareViewable(file); // 画像を<img>で直接表示するには閲覧権限が必要
    return {
      url: file.getUrl(),
      directUrl: 'https://drive.google.com/uc?export=view&id=' + file.getId(), // <img src>で直接表示できる形式
      fileId: file.getId(),
      fileName: p.fileName
    };
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
    de:String(r['DE基板']||''), e:String(r['E基板']||''), c:String(r['C基板']||''), s:String(r['S基板']||''), r:String(r['R基板']||''),
    relatedDocs:String(r['関連資料']||''), salesUnits:String(r['売上台数']||''), salesTarget:String(r['売上目標']||''),
    reuseEcoUnits:String(r['リユース・エコ投入台数']||''), stockQty:String(r['在庫数']||''),
    photoUrl:String(r['写真URL']||''), summary:String(r['概要']||''), features:String(r['特徴・ポイント']||''),
    studyNote:String(r['学習メモ']||''), manualIds:String(r['関連マニュアルID']||''), note:String(r['備考']||''), updatedAt:String(r['更新日時']||'') };
}
function _rowToBoard(r) {
  return { id:String(r['基板ID']||''), name:String(r['基板名']||''),
    type:String(r['種類']||'') || _extractSpecValue(String(r['仕様']||''), '基板種別'),
    category:String(r['分類']||''), brand:String(r['ブランド']||''),
    gameType:String(r['対応区分']||''), version:String(r['バージョン']||''), status:String(r['ステータス']||''),
    specs:String(r['仕様']||''), revisions:String(r['改定履歴']||''), relatedDocs:String(r['関連資料']||''),
    photoUrl:String(r['写真URL']||''), summary:String(r['機能概要']||''),
    mainParts:String(r['主要部品']||''), studyPoint:String(r['学習ポイント']||''), relatedModels:String(r['関連機種コード']||''),
    note:String(r['備考']||''), updatedAt:String(r['更新日時']||'') };
}
// 「仕様」JSON配列から指定labelの値を取り出す（旧データの「基板種別」を「種類」列へフォールバック取得するため）
function _extractSpecValue(specsJson, label) {
  try {
    var arr = JSON.parse(specsJson || '[]');
    if (!Array.isArray(arr)) return '';
    var found = arr.filter(function(s){ return s && s.label === label; })[0];
    return found ? String(found.value || '') : '';
  } catch(e) { return ''; }
}
function _parseJsonArraySafe(str) {
  try { var a = JSON.parse(str||'[]'); return Array.isArray(a) ? a : []; } catch(e) { return []; }
}
// 配列フィールドの安全な合成（同じキー[field]の値が既にあれば重複追加しない）
function _unionByField(a, b, field) {
  var seen = {}, out = [];
  a.concat(b).forEach(function(item) {
    var k = (item && item[field]) || JSON.stringify(item);
    if (!seen[k]) { seen[k] = true; out.push(item); }
  });
  return out;
}
// SEED配列を既存データへ安全にマージ（配列フィールドは追記合成、スカラーは上書き）
function _mergeSeedInto(map, key, seedItem, arrayFields) {
  var existing = map[key] || {};
  var merged = Object.assign({}, existing, seedItem);
  (arrayFields||[]).forEach(function(f) {
    var joined = _unionByField(_parseJsonArraySafe(existing[f]), _parseJsonArraySafe(seedItem[f]), f === 'specs' ? 'label' : 'url');
    if (joined.length) merged[f] = JSON.stringify(joined);
  });
  map[key] = merged;
}

// ══════════════════════════════════════════════════
// 実データ投入 第2弾（基板PCB価格・機種一覧スプレッドシートから抽出したデータ）
// 原価・公表単価・見積書リンク・機種の資料リンクなどを追加する。既存データは配列項目は合成、その他は上書き
// ══════════════════════════════════════════════════
function apiSeedRealData2() {
  return _wrap(function() {
    var models = _read(S.MODELS, H.MODELS);
    var boards = _read(S.BOARDS, H.BOARDS);
    var modelMap = {}, boardMap = {};
    models.forEach(function(r){ if (r['機種コード']) modelMap[r['機種コード']] = _rowToModel(r); });
    boards.forEach(function(r){ if (r['基板ID']) boardMap[r['基板ID']] = _rowToBoard(r); });

    SEED_MODELS_2.forEach(function(m){ _mergeSeedInto(modelMap, m.code, m, ['relatedDocs']); });
    SEED_BOARDS_2.forEach(function(b){ _mergeSeedInto(boardMap, b.id, b, ['relatedDocs','specs']); });

    apiSaveModels(modelMap);
    apiSaveBoards(boardMap);
    return { models: SEED_MODELS_2.length, boards: SEED_BOARDS_2.length };
  });
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

// 基板PCB価格・機種一覧スプレッドシート（ユーザー提供）より抽出。原価/公表単価/見積書リンク等を含む
var SEED_MODELS_2 = [
  {code:"A63BW", name:"A63BW(リング４.5)", gameType:"パチンコ", m:"M1703A", c:"C1601E", e:"E1403C", d:"D1401C", de:"DE1802A", r:"MSPB620KR", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「A63BW」商品説明会用資料.xlsm\", \"url\": \"https://drive.google.com/open?id=1RkJ3c4eniWdNNtmw_RW0iUpCHXSwARi5&usp=drive_copy\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/file/d/15d-s4LU6-TaoEYSi3vi5IJxtiTCtibal/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"払出制御基板(C) 資料1\", \"url\": \"https://drive.google.com/file/d/19wjx1ciwRw3Ja21TaX381gybqf_65OiJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"演出基板(E) 資料1\", \"url\": \"https://drive.google.com/file/d/1oLrw5ZKXrIDGR2D6xBfyzEi2mJz2L1WA/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"図柄制御基板(D) 資料1\", \"url\": \"https://drive.google.com/file/d/1b15c4cfEMVNDFMslUnnCIHWP-DLcYVaJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"液晶IF基板(DE) 資料1\", \"url\": \"https://drive.google.com/file/d/1okwv9-lskK-tD5e2EYpDPJp6iyM2h2B5/view?usp=drivesdk\"}]"},
  {code:"A64B", name:"Ａ６４Ｂ(地獄少女3.5 小当たりラッシュ機", gameType:"パチンコ", m:"M1802A", c:"C1601E", e:"E1403C", d:"D1401C", de:"DE1801A", r:"MSPB620KR", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「A64B」商品説明会用資料.xlsx\", \"url\": \"https://docs.google.com/spreadsheets/d/1qNUFkaLNWRX4tfgMwDWKB-OEt6BwnTJk/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/file/d/1FmHxrsssFQ8-qpM-rKNBA0omhBPa1k0_/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"払出制御基板(C) 資料1\", \"url\": \"https://drive.google.com/file/d/19wjx1ciwRw3Ja21TaX381gybqf_65OiJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"演出基板(E) 資料1\", \"url\": \"https://drive.google.com/file/d/1oLrw5ZKXrIDGR2D6xBfyzEi2mJz2L1WA/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"図柄制御基板(D) 資料1\", \"url\": \"https://drive.google.com/file/d/1b15c4cfEMVNDFMslUnnCIHWP-DLcYVaJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"液晶IF基板(DE) 資料1\", \"url\": \"https://drive.google.com/file/d/1id-3A4C-Qt540gKv7ekUv3w_ESqfD_1H/view?usp=drivesdk\"}]"},
  {code:"A64W", name:"Ａ６４Ｗ(地獄少女3.5 甘", gameType:"パチンコ", m:"M1801A", c:"C1601E", e:"E1403C", d:"D1401C", de:"DE1801A", r:"MSPB620KR", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「A64Ｗ」商品説明会用資料.xlsm\", \"url\": \"https://docs.google.com/spreadsheets/d/1a7DkuRucwcSlZqhJVJErqn74EvokLJhi/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/file/d/14oKDbdl8wCFrqrWgU6fN4PDsESBHy0PO/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"払出制御基板(C) 資料1\", \"url\": \"https://drive.google.com/file/d/19wjx1ciwRw3Ja21TaX381gybqf_65OiJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"演出基板(E) 資料1\", \"url\": \"https://drive.google.com/file/d/1oLrw5ZKXrIDGR2D6xBfyzEi2mJz2L1WA/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"図柄制御基板(D) 資料1\", \"url\": \"https://drive.google.com/file/d/1b15c4cfEMVNDFMslUnnCIHWP-DLcYVaJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"液晶IF基板(DE) 資料1\", \"url\": \"https://drive.google.com/file/d/1id-3A4C-Qt540gKv7ekUv3w_ESqfD_1H/view?usp=drivesdk\"}]"},
  {code:"A68", name:"Ａ６８（リング５）", gameType:"パチンコ", m:"M1801A", c:"C1601E", e:"E1403C", d:"D1401C", de:"DE1802A", r:"MSPB620KR", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「A68」商品説明会用資料.xlsm\", \"url\": \"https://docs.google.com/spreadsheets/d/1jouobewEWRyEQRzFcC0-D9_hAIbqvskb/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/file/d/14oKDbdl8wCFrqrWgU6fN4PDsESBHy0PO/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"払出制御基板(C) 資料1\", \"url\": \"https://drive.google.com/file/d/19wjx1ciwRw3Ja21TaX381gybqf_65OiJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"演出基板(E) 資料1\", \"url\": \"https://drive.google.com/file/d/1oLrw5ZKXrIDGR2D6xBfyzEi2mJz2L1WA/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"図柄制御基板(D) 資料1\", \"url\": \"https://drive.google.com/file/d/1b15c4cfEMVNDFMslUnnCIHWP-DLcYVaJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"液晶IF基板(DE) 資料1\", \"url\": \"https://drive.google.com/file/d/1okwv9-lskK-tD5e2EYpDPJp6iyM2h2B5/view?usp=drivesdk\"}]"},
  {code:"A70", name:"Ａ７０（地獄少女４）", gameType:"パチンコ", m:"M1904B", c:"C1601E", e:"E1403C", d:"D1401C", de:"DE1802A", r:"MSPB620KR", salesUnits:"15000", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「A70」商品説明会用資料.xlsm\", \"url\": \"https://drive.google.com/open?id=1nByUvB0XPF1jaG4hDnCCbc9qIt20Fmgf&usp=drive_copy\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/file/d/1B2OzQ29MJaWzZoS-wPT3-65s5ZoNjbs5/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"払出制御基板(C) 資料1\", \"url\": \"https://drive.google.com/file/d/19wjx1ciwRw3Ja21TaX381gybqf_65OiJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"演出基板(E) 資料1\", \"url\": \"https://drive.google.com/file/d/1oLrw5ZKXrIDGR2D6xBfyzEi2mJz2L1WA/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"図柄制御基板(D) 資料1\", \"url\": \"https://drive.google.com/file/d/1b15c4cfEMVNDFMslUnnCIHWP-DLcYVaJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"液晶IF基板(DE) 資料1\", \"url\": \"https://drive.google.com/file/d/1okwv9-lskK-tD5e2EYpDPJp6iyM2h2B5/view?usp=drivesdk\"}]"},
  {code:"A71", name:"リング６", gameType:"パチンコ", m:"M1904A", c:"C1601E", e:"E1403C", de:"DE1802A", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「A71」商品説明会用資料.xlsm\", \"url\": \"https://docs.google.com/spreadsheets/d/1GOob1CD8eW9A-pOsmyzR6ECsHhfosVLT/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/file/d/1emkx14a0UgEzNAcsqT4D5pYljwnasz4N/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"払出制御基板(C) 資料1\", \"url\": \"https://drive.google.com/file/d/19wjx1ciwRw3Ja21TaX381gybqf_65OiJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"演出基板(E) 資料1\", \"url\": \"https://drive.google.com/file/d/1oLrw5ZKXrIDGR2D6xBfyzEi2mJz2L1WA/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"液晶IF基板(DE) 資料1\", \"url\": \"https://drive.google.com/file/d/1okwv9-lskK-tD5e2EYpDPJp6iyM2h2B5/view?usp=drivesdk\"}]"},
  {code:"A72W"},
  {code:"A73", name:"Ａ７３（ 暴れん坊将軍８）", gameType:"パチンコ", m:"M1904B", c:"C1601E", e:"E1403C", d:"D1401C", de:"DE1607A", r:"AX75611", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「A73」商品説明会用資料.xlsm\", \"url\": \"https://drive.google.com/open?id=1Wq4olc-LPhfz3GF5IqgYch6YphUyRgcK&usp=drive_copy\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/file/d/1B2OzQ29MJaWzZoS-wPT3-65s5ZoNjbs5/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"払出制御基板(C) 資料1\", \"url\": \"https://drive.google.com/file/d/19wjx1ciwRw3Ja21TaX381gybqf_65OiJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"演出基板(E) 資料1\", \"url\": \"https://drive.google.com/file/d/1oLrw5ZKXrIDGR2D6xBfyzEi2mJz2L1WA/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"図柄制御基板(D) 資料1\", \"url\": \"https://drive.google.com/file/d/1b15c4cfEMVNDFMslUnnCIHWP-DLcYVaJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"液晶IF基板(DE) 資料1\", \"url\": \"https://drive.google.com/file/d/1wUJA20ljZy6_-F_SGrhgyIAXoQdTIfqC/view?usp=drivesdk\"}]"},
  {code:"A74", name:"アリア４", gameType:"パチンコ", m:"M2002B", c:"C1601E", e:"E1403C", de:"DE1802A", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「A74」商品説明会用資料.xlsm\", \"url\": \"https://docs.google.com/spreadsheets/d/1nHg0wmDG02NUnjHdBBcIsJEc1zkL-cZl/edit?gid=51764943\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/file/d/1Dtcgglign0jQj1t77aHV4ToeiSE36neX/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"払出制御基板(C) 資料1\", \"url\": \"https://drive.google.com/file/d/19wjx1ciwRw3Ja21TaX381gybqf_65OiJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"演出基板(E) 資料1\", \"url\": \"https://drive.google.com/file/d/1oLrw5ZKXrIDGR2D6xBfyzEi2mJz2L1WA/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"液晶IF基板(DE) 資料1\", \"url\": \"https://drive.google.com/file/d/1okwv9-lskK-tD5e2EYpDPJp6iyM2h2B5/view?usp=drivesdk\"}]"},
  {code:"A74W", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「A74W」商品説明会用資料.xlsm\", \"url\": \"https://docs.google.com/spreadsheets/d/1sTpJgZI9_f2ifZ5fVVQt-Zz-Gj7ISEfD/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\"}]"},
  {code:"A76", name:"A76　ストリートファイター", gameType:"パチンコ", m:"M2003A", c:"C1601E", e:"E1901B", d:"D2101A", de:"DE1607A", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「A76」商品説明会用資料.xlsm\", \"url\": \"https://docs.google.com/spreadsheets/d/1wZP4z7v-ueF9eZxbT8PFiyvNpQl8bK2B/edit?gid=397039102\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/file/d/1Dtcgglign0jQj1t77aHV4ToeiSE36neX/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料2\", \"url\": \"https://drive.google.com/open?id=1fcMjqT6f-qBPcxWQ12_XYZnY6N7nRcGg&usp=drive_copy\"}, {\"category\": \"構成表-見本機\", \"title\": \"払出制御基板(C) 資料1\", \"url\": \"https://drive.google.com/file/d/19wjx1ciwRw3Ja21TaX381gybqf_65OiJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"演出基板(E) 資料1\", \"url\": \"https://drive.google.com/file/d/1Q3WrRljasxX_tc2eCzGJC1EsOgDbSG0G/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"図柄制御基板(D) 資料1\", \"url\": \"https://drive.google.com/file/d/1xIlg3O1VjuC3wipKqtGUWNBiArKc7sYL/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"液晶IF基板(DE) 資料1\", \"url\": \"https://drive.google.com/file/d/1wUJA20ljZy6_-F_SGrhgyIAXoQdTIfqC/view?usp=drivesdk\"}]"},
  {code:"A77", name:"eリング7", gameType:"パチンコ", m:"M2105C", e:"E2101B", d:"D2101A", de:"DE2103A", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「eA77」商品説明会用資料.xlsm\", \"url\": \"https://docs.google.com/spreadsheets/d/1wG3qwiaCqoGKXMh4AILbmGZV8XwTkvNb/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/file/d/1zI_sVb0ZjkHewRXXrgjiYdtPaV5uZ6Db/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料2\", \"url\": \"https://drive.google.com/open?id=1ueKULiAD9q0MzihIM_hR5lmgfCZhPsiu&usp=drive_copy\"}, {\"category\": \"構成表-見本機\", \"title\": \"演出基板(E) 資料1\", \"url\": \"https://drive.google.com/file/d/1xIlg3O1VjuC3wipKqtGUWNBiArKc7sYL/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"液晶IF基板(DE) 資料1\", \"url\": \"https://drive.google.com/file/d/180jCa_EwyOZH1qNLS1HtSVMMpJHA_RXd/view?usp=drivesdk\"}]"},
  {code:"A78", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「A78」商品説明会用資料.xlsm\", \"url\": \"https://docs.google.com/spreadsheets/d/1dBns94YaPas3tnQu1ziYLkSTIH2RHAxD/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\"}]"},
  {code:"A79", relatedDocs:"[{\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/open?id=1y7puUkL1tu2lQN_fd-5gZVFN62cxZRXg&usp=drive_copy\"}]"},
  {code:"A79e", name:"Ａ７９ｅ", m:"M2105C1", e:"E2101B", d:"D1401C", de:"DE2103A", relatedDocs:"[{\"category\": \"その他\", \"title\": \"機種説明（A79e）_チェックボックス自動化_190617.xlsm\", \"url\": \"https://drive.google.com/open?id=1blATeHx6D741SPv_-WL6grgfGArI5ZGw&usp=drive_copy\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/file/d/1zI_sVb0ZjkHewRXXrgjiYdtPaV5uZ6Db/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"演出基板(E) 資料1\", \"url\": \"https://drive.google.com/file/d/1xIlg3O1VjuC3wipKqtGUWNBiArKc7sYL/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"図柄制御基板(D) 資料1\", \"url\": \"https://drive.google.com/file/d/1b15c4cfEMVNDFMslUnnCIHWP-DLcYVaJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"液晶IF基板(DE) 資料1\", \"url\": \"https://drive.google.com/file/d/180jCa_EwyOZH1qNLS1HtSVMMpJHA_RXd/view?usp=drivesdk\"}]"},
  {code:"A80", relatedDocs:"[{\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/open?id=1ourrqM5XSnTz7bFIrQA0VF9ExMRRLbTp&usp=drive_copy\"}]"},
  {code:"A81", relatedDocs:"[{\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/open?id=1xtEaJACI3yDzIEJog93yi1S07sRtjp-P&usp=drive_copy\"}]"},
  {code:"A81W", relatedDocs:"[{\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/open?id=1AfrB9_jlLpsMCADHwPbMyzfzMKJEVZxO&usp=drive_copy\"}]"},
  {code:"A82", name:"世界最高の暗殺者", gameType:"パチンコ", relatedDocs:"[{\"category\": \"その他\", \"title\": \"A82商品説明会用資料.xlsx\", \"url\": \"https://drive.google.com/open?id=1nHIWyjEpcILW6IWaNg5ZYTqTSdydO5Ff&usp=drive_copy\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/open?id=1cK7OmdovW0yJo3yNqdLkI7B9cXe2gK7R&usp=drive_copy\"}]"},
  {code:"A82サブ"},
  {code:"A83", name:"防振り", gameType:"パチンコ", relatedDocs:"[{\"category\": \"その他\", \"title\": \"A83商品説明会用資料.xlsx\", \"url\": \"https://docs.google.com/spreadsheets/d/1z9RHySDERZwyGzZN2EYaEiY31pMLwtTO/edit?gid=47541224\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/open?id=1PLFdgrXGFBqgNx511EZm46-Ysx1_e-q6&usp=drive_copy\"}]"},
  {code:"A84", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「A84」商品説明会用資料.xlsm\", \"url\": \"https://docs.google.com/spreadsheets/d/105_-nzGWgl3cB96zVteb8b2253JzKRVSewBTurl70SU/edit?gid=750903847\"}]"},
  {code:"A84B"},
  {code:"A84C"},
  {code:"A85", relatedDocs:"[{\"category\": \"その他\", \"title\": \"機種説明A85基板構成.xlsm.xlsx\", \"url\": \"https://drive.google.com/open?id=1iGi3BBx6NLn6_Ds8AUMtZ2I0qM8UAr8x&usp=drive_copy\"}]"},
  {code:"A85甘"},
  {code:"A86", relatedDocs:"[{\"category\": \"その他\", \"title\": \"機種説明A86基板構成.xlsm\", \"url\": \"https://drive.google.com/open?id=1JeDYf3NQgdfh8sC6LQyPpCmZOtQZNU5y0MdE4FRdqzc&usp=drive_copy\"}]"},
  {code:"A86甘"},
  {code:"A87", relatedDocs:"[{\"category\": \"その他\", \"title\": \"機種説明A87基板構成.xlsm\", \"url\": \"https://drive.google.com/open?id=1SQj90QvcRXQaoFy89scvyZLLuM-2LqiauCrZS-ykeHs&usp=drive_copy\"}]"},
  {code:"A87サブ"},
  {code:"A88", relatedDocs:"[{\"category\": \"その他\", \"title\": \"機種説明A88基板構成(修正20250916).xlsm.xlsx\", \"url\": \"https://drive.google.com/open?id=1_OiDvQcatuvpdEvpDwJVC7qOVfJD0ItW&usp=drive_copy\"}]"},
  {code:"A89", name:"A89(片田舎のおっさん)", relatedDocs:"[{\"category\": \"その他\", \"title\": \"機種説明A89基板構成(修正20250916).xlsm.xlsx\", \"url\": \"https://drive.google.com/open?id=1iSOSA0ZArHZ8wye8s5gI-NUnJ5zssH1m&usp=drive_copy\"}]"},
  {code:"C40", name:"戦国恋姫（改）の焼き直し", gameType:"パチンコ", m:"M1904B", c:"C1601E", e:"E1403C", d:"D1401C", de:"DE1607A", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「C40」商品説明会用資料.xlsm\", \"url\": \"https://docs.google.com/spreadsheets/d/1-5id4x2p77TK75UJTOIYJn_Tsi5kgVh1/edit?gid=835724895\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/file/d/1B2OzQ29MJaWzZoS-wPT3-65s5ZoNjbs5/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"払出制御基板(C) 資料1\", \"url\": \"https://drive.google.com/file/d/19wjx1ciwRw3Ja21TaX381gybqf_65OiJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"演出基板(E) 資料1\", \"url\": \"https://drive.google.com/file/d/1oLrw5ZKXrIDGR2D6xBfyzEi2mJz2L1WA/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"図柄制御基板(D) 資料1\", \"url\": \"https://drive.google.com/file/d/1b15c4cfEMVNDFMslUnnCIHWP-DLcYVaJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"液晶IF基板(DE) 資料1\", \"url\": \"https://drive.google.com/file/d/1wUJA20ljZy6_-F_SGrhgyIAXoQdTIfqC/view?usp=drivesdk\"}]"},
  {code:"C41", name:"どないやねんDX", gameType:"パチンコ", m:"FJ/JJ+M1904B", c:"C1601E", e:"E1403C", d:"D1201B", de:"DE1901A", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「C41」商品説明会用資料.xlsm\", \"url\": \"https://drive.google.com/open?id=1MOv8x63rIBjDXcm4NBUTNI9DGJUeYAmx&usp=drive_copy\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/file/d/1Sxbg22blhFCCAbIIMBx3HDYTZQc9_Rpn/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"払出制御基板(C) 資料1\", \"url\": \"https://drive.google.com/file/d/19wjx1ciwRw3Ja21TaX381gybqf_65OiJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"演出基板(E) 資料1\", \"url\": \"https://drive.google.com/file/d/1oLrw5ZKXrIDGR2D6xBfyzEi2mJz2L1WA/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"図柄制御基板(D) 資料1\", \"url\": \"https://drive.google.com/file/d/1D3cRE7D7OY3ZXBeYlMrL6Rp-XOYBo0QY/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"液晶IF基板(DE) 資料1\", \"url\": \"https://drive.google.com/file/d/1AqfbzkUUSYtDbpc631QkaO7s2qWOOUJQ/view?usp=drivesdk\"}]"},
  {code:"C42", name:"アレンジ", gameType:"パチンコ", m:"FJ/JJ+M2003A", c:"C1601E", e:"E1403C", d:"D1201B", de:"DE1902A", r:"FJ+R2001A", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「C42」商品説明会用資料.xlsm\", \"url\": \"https://drive.google.com/open?id=1ZZCh4-stJYb3jZDHqvOfOJPLKCNoMl2p&usp=drive_copy\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/file/d/1Dtcgglign0jQj1t77aHV4ToeiSE36neX/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"払出制御基板(C) 資料1\", \"url\": \"https://drive.google.com/file/d/19wjx1ciwRw3Ja21TaX381gybqf_65OiJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"演出基板(E) 資料1\", \"url\": \"https://drive.google.com/file/d/1oLrw5ZKXrIDGR2D6xBfyzEi2mJz2L1WA/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"図柄制御基板(D) 資料1\", \"url\": \"https://drive.google.com/file/d/1D3cRE7D7OY3ZXBeYlMrL6Rp-XOYBo0QY/view?usp=drivesdk\"}]"},
  {code:"C43", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「C43」商品説明会用資料.xlsm\", \"url\": \"https://docs.google.com/spreadsheets/d/1FqK8XzSJgGKX7H25VucoY11X3K_RGiqVJqemEWSKzXw/edit?gid=1676251525\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/open?id=1WXT54DhEZoIS1DFM53-5MWEmVhO6t7le&usp=drive_copy\"}]"},
  {code:"C44", name:"RAVE３", gameType:"パチンコ", m:"M2105B", e:"E2101B", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「C44」商品説明会用資料.xlsm\", \"url\": \"https://docs.google.com/spreadsheets/d/1F3DmKVzxBhs44s7Mvod3rbpvBjg7QOPj/edit?gid=1844705279\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/file/d/1LPyxw8R-oPPag38pC_hf8PyfESeKO4Gh/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"演出基板(E) 資料1\", \"url\": \"https://drive.google.com/file/d/1xIlg3O1VjuC3wipKqtGUWNBiArKc7sYL/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"図柄制御基板(D) 資料1\", \"url\": \"https://drive.google.com/file/d/1aL9OYerVnLsG3EIH-XHnLcfY1HMtUjFW/view?usp=drivesdk\"}]"},
  {code:"C45", name:"貞子", gameType:"パチンコ", relatedDocs:"[{\"category\": \"その他\", \"title\": \"C45商品説明会用資料.xlsx\", \"url\": \"https://docs.google.com/spreadsheets/d/1spAl03xLdZuAgjk_lzr-aZ-s-TFKLRnf/edit?gid=47541224\"}]"},
  {code:"C45", relatedDocs:"[{\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/open?id=1FHRLMimB12dH97seEhRPL1P4k6Gxd-DT&usp=drive_copy\"}]"},
  {code:"C46", relatedDocs:"[{\"category\": \"その他\", \"title\": \"機種説明C46基板構成.xlsm\", \"url\": \"https://drive.google.com/open?id=1xj2061NnzmcYyEf-2mIo1puwxQxUqoR5uJ5X0hPy8xY&usp=drive_copy\"}]"},
  {code:"C46B"},
  {code:"C47"},
  {code:"D45CW", name:"Ｄ４５ＣＷ(喰霊零Ｃ 甘)", gameType:"パチンコ", m:"M1801A", c:"C1601E", e:"E1403C", d:"D1401C", de:"DE1802A", r:"MSPB620KR", relatedDocs:"[{\"category\": \"その他\", \"title\": \"D45CW_商品説明会用資料_20180202.xlsm\", \"url\": \"https://docs.google.com/spreadsheets/d/1-6SBANaG1Wnlv0Umkss59ct3821l2zCQ/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/file/d/14oKDbdl8wCFrqrWgU6fN4PDsESBHy0PO/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"払出制御基板(C) 資料1\", \"url\": \"https://drive.google.com/file/d/19wjx1ciwRw3Ja21TaX381gybqf_65OiJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"演出基板(E) 資料1\", \"url\": \"https://drive.google.com/file/d/1oLrw5ZKXrIDGR2D6xBfyzEi2mJz2L1WA/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"図柄制御基板(D) 資料1\", \"url\": \"https://drive.google.com/file/d/1b15c4cfEMVNDFMslUnnCIHWP-DLcYVaJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"液晶IF基板(DE) 資料1\", \"url\": \"https://drive.google.com/file/d/1okwv9-lskK-tD5e2EYpDPJp6iyM2h2B5/view?usp=drivesdk\"}]"},
  {code:"D46W", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「D46W」商品説明会用資料.xlsm\", \"url\": \"https://drive.google.com/open?id=1FthRgGXaB3v7-9AQGAy_srmsVU6c4w6h&usp=drive_copy\"}]"},
  {code:"D48", name:"Ｄ４８", gameType:"パチンコ", m:"M1802A", c:"C1601E", e:"E1403C", d:"D1401C", de:"DE1802A", r:"MSPB620KR", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「D48」商品説明会用資料.xlsm\", \"url\": \"https://docs.google.com/spreadsheets/d/1aO8LJvaCXNvSxn52x4REoYmlFOAJf0F_/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/file/d/1FmHxrsssFQ8-qpM-rKNBA0omhBPa1k0_/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"払出制御基板(C) 資料1\", \"url\": \"https://drive.google.com/file/d/19wjx1ciwRw3Ja21TaX381gybqf_65OiJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"演出基板(E) 資料1\", \"url\": \"https://drive.google.com/file/d/1oLrw5ZKXrIDGR2D6xBfyzEi2mJz2L1WA/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"図柄制御基板(D) 資料1\", \"url\": \"https://drive.google.com/file/d/1b15c4cfEMVNDFMslUnnCIHWP-DLcYVaJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"液晶IF基板(DE) 資料1\", \"url\": \"https://drive.google.com/file/d/1okwv9-lskK-tD5e2EYpDPJp6iyM2h2B5/view?usp=drivesdk\"}]"},
  {code:"D49", name:"Ｄ４９", gameType:"パチンコ", m:"M1801A", c:"C1601E", e:"E1403C", d:"D1401C", de:"DE1802A", r:"AHSF2-640R-N", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「D49」商品説明会用資料 (1).xlsm\", \"url\": \"https://docs.google.com/spreadsheets/d/1YmtHEotK9QK9dzCimaay3P1WYle7Tc5A/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/file/d/14oKDbdl8wCFrqrWgU6fN4PDsESBHy0PO/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"払出制御基板(C) 資料1\", \"url\": \"https://drive.google.com/file/d/19wjx1ciwRw3Ja21TaX381gybqf_65OiJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"演出基板(E) 資料1\", \"url\": \"https://drive.google.com/file/d/1oLrw5ZKXrIDGR2D6xBfyzEi2mJz2L1WA/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"図柄制御基板(D) 資料1\", \"url\": \"https://drive.google.com/file/d/1b15c4cfEMVNDFMslUnnCIHWP-DLcYVaJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"液晶IF基板(DE) 資料1\", \"url\": \"https://drive.google.com/file/d/1okwv9-lskK-tD5e2EYpDPJp6iyM2h2B5/view?usp=drivesdk\"}]"},
  {code:"D50", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「D50」商品説明会用資料.xlsm\", \"url\": \"https://drive.google.com/open?id=1k5OhJpCfN7ZQ3tUGgp7Zo4UlZcfehjAu&usp=drive_copy\"}]"},
  {code:"D51", relatedDocs:"[{\"category\": \"その他\", \"title\": \"機種説明（D51）_チェックボックス自動化_190617.xlsm\", \"url\": \"https://drive.google.com/open?id=1cFtAYrrKp-__IYJJJoStqEAN3HClHLSg&usp=drive_copy\"}]"},
  {code:"D53", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「D53」商品説明会用資料.xlsm\", \"url\": \"https://docs.google.com/spreadsheets/d/1dy5fnSlSn8ehxD9n9E1HHY8kdG074CDP/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\"}]"},
  {code:"D54", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「D54」商品説明会用資料.xlsm\", \"url\": \"https://docs.google.com/spreadsheets/d/1ISST9Mhjvbm7_2PGKpiaiQr1QjjmnkSC/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/open?id=1CNly6j38YuFA-xeG__LE49fw1_VF-fCu&usp=drive_copy\"}, {\"category\": \"構成表-見本機\", \"title\": \"図柄制御基板(D) 資料1\", \"url\": \"https://drive.google.com/open?id=1iQnotfx7yh2Klpnlf3F_RaK294v6q_38&usp=drive_copy\"}]"},
  {code:"D55", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「D55」商品説明会用資料.xlsm\", \"url\": \"https://docs.google.com/spreadsheets/d/1JzMwJcyzfUn5L5t8F6kCVvtNp8a3C33F/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/open?id=1VGwd3-p9cmU-Ud2XDgA5Z5izrviZSqHQ&usp=drive_copy\"}]"},
  {code:"D56", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「D56」商品説明会用資料.xlsx\", \"url\": \"https://docs.google.com/spreadsheets/d/1gBUOtaqvx-8SwjCgrGO6D5SiTFBLVPza/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/open?id=13yD-jGk4MriSmIsjYGU_26gAs6F9gpEs&usp=drive_copy\"}]"},
  {code:"D57", relatedDocs:"[{\"category\": \"その他\", \"title\": \"機種説明D57基板構成.xlsm\", \"url\": \"https://drive.google.com/open?id=1Xk0MaoQOg67xNvndvhNOchU5TZnYC_6Z&usp=drive_copy\"}]"},
  {code:"D57B", relatedDocs:"[{\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/open?id=1O2hSiS1tXdTHWXmIS1qP7PHd_qfxlooW&usp=drive_copy\"}]"},
  {code:"D57C", relatedDocs:"[{\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/open?id=1o1M4M3wCY8U_bQzrAh__XtU8XNLpVq1g&usp=drive_copy\"}]"},
  {code:"D58", relatedDocs:"[{\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/open?id=1zOimgYSZzicCPFPqO0Z9Ku1ZX8EumC7b&usp=drive_copy\"}]"},
  {code:"D58B", relatedDocs:"[{\"category\": \"その他\", \"title\": \"機種説明D58B基板構成.xlsm\", \"url\": \"https://drive.google.com/open?id=1LJX4LSnGXBinorEqGUyATO6rL0oRT203&usp=drive_copy\"}]"},
  {code:"D59"},
  {code:"D59", relatedDocs:"[{\"category\": \"その他\", \"title\": \"機種説明D59基板構成.xlsm\", \"url\": \"https://drive.google.com/open?id=1eio8gL4d0ccBLM9dgzfltzxRaWZsHf1S&usp=drive_copy\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/open?id=1yQXXYfnw5PRe02waP1fFruCMasKk0WJf&usp=drive_copy\"}]"},
  {code:"D59サブ"},
  {code:"D60", name:"いせれべ", gameType:"パチンコ", salesTarget:"いせれべ", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「D60」商品説明会用資料.xlsx\", \"url\": \"https://drive.google.com/open?id=1FSoJOZ6Giy_5-LEDYF3yTFBChiJCKW0f&usp=drive_copy\"}]"},
  {code:"D60甘"},
  {code:"D61", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「D61」商品説明会用資料.xlsx\", \"url\": \"https://drive.google.com/open?id=1Hy2V-ISCET0cZkk4o8h3wifwqV4BLnkJ&usp=drive_copy\"}]"},
  {code:"D61"},
  {code:"D61甘"},
  {code:"D62", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「D62」商品説明会用資料.xlsx\", \"url\": \"https://drive.google.com/open?id=1ptBDwo0cl_bZ36Y6AXeS4Yvza8zp5YWl&usp=drive_copy\"}]"},
  {code:"D62"},
  {code:"D63"},
  {code:"E36B", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「E36B」商品説明会用資料20180827.xlsx\", \"url\": \"https://docs.google.com/spreadsheets/d/1Q4XhncVeVaLuYgI2S6JTwadcXGks9pEH/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\"}]"},
  {code:"E37", name:"Ｅ３７(喰霊)", m:"M1805D", e:"E1501C", d:"D1401C", de:"DE1802A", r:"AHSF2-640\nAXCEL(64G)", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「E37」商品説明会用資料20181031.xlsx\", \"url\": \"https://drive.google.com/open?id=11BPsWqSu_6igZJMDX0YOWu3UMX7p_8Uf&usp=drive_copy\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/file/d/1Rsf5VuN72a1cPBSQsH9wxmzSS-Jd7lJl/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"演出基板(E) 資料1\", \"url\": \"https://drive.google.com/file/d/1v5v3UycLBMJ_FA36WtpLYLN_rwk_lsX5/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"図柄制御基板(D) 資料1\", \"url\": \"https://drive.google.com/file/d/1b15c4cfEMVNDFMslUnnCIHWP-DLcYVaJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"液晶IF基板(DE) 資料1\", \"url\": \"https://drive.google.com/file/d/1okwv9-lskK-tD5e2EYpDPJp6iyM2h2B5/view?usp=drivesdk\"}]"},
  {code:"E39", name:"Ｅ３９(リング３)", m:"M1805D", e:"E1501C", d:"D1401C", de:"DE1607A", r:"AHSF2-640R-N", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「E39」商品説明会用資料20190111.xlsx\", \"url\": \"https://drive.google.com/open?id=1Fl4V_qgpey5ai0rQxDspIkWL0KcvU1Gk&usp=drive_copy\"}, {\"category\": \"構成表-見本機\", \"title\": \"演出基板(E) 資料1\", \"url\": \"https://drive.google.com/file/d/1v5v3UycLBMJ_FA36WtpLYLN_rwk_lsX5/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"図柄制御基板(D) 資料1\", \"url\": \"https://drive.google.com/file/d/1b15c4cfEMVNDFMslUnnCIHWP-DLcYVaJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"液晶IF基板(DE) 資料1\", \"url\": \"https://drive.google.com/file/d/1wUJA20ljZy6_-F_SGrhgyIAXoQdTIfqC/view?usp=drivesdk\"}]"},
  {code:"E40", name:"E40「地獄少女3」", gameType:"回胴", m:"M1805D", e:"E1501C", de:"DE1802A", relatedDocs:"[{\"category\": \"その他\", \"title\": \"機種説明（E40）.xlsm\", \"url\": \"https://drive.google.com/open?id=1WH96OPiNv4ZkxF2nh3D-I3FfGWezSSWR&usp=drive_copy\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/file/d/1zmXisJyTj0DIUGox0EemFofBDL_M5WLA/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"演出基板(E) 資料1\", \"url\": \"https://drive.google.com/file/d/1v5v3UycLBMJ_FA36WtpLYLN_rwk_lsX5/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"液晶IF基板(DE) 資料1\", \"url\": \"https://drive.google.com/file/d/1okwv9-lskK-tD5e2EYpDPJp6iyM2h2B5/view?usp=drivesdk\"}]"},
  {code:"E44", name:"アリア２", gameType:"パチンコ", m:"M1805D", c:"C1601E", e:"E1902B", d:"D2101A", de:"DE1802A", relatedDocs:"[{\"category\": \"その他\", \"title\": \"「E44」商品説明会用資料.xlsm\", \"url\": \"https://drive.google.com/open?id=1xW95Gb30ooKosktAAZgIG5_nm2TTXeTE&usp=drive_copy\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/file/d/1Rsf5VuN72a1cPBSQsH9wxmzSS-Jd7lJl/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"払出制御基板(C) 資料1\", \"url\": \"https://drive.google.com/file/d/19wjx1ciwRw3Ja21TaX381gybqf_65OiJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"演出基板(E) 資料1\", \"url\": \"https://drive.google.com/file/d/1qjJRiFWxTUGhubYcweZ7gsHm31PNBWgX/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"図柄制御基板(D) 資料1\", \"url\": \"https://drive.google.com/file/d/1xIlg3O1VjuC3wipKqtGUWNBiArKc7sYL/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"液晶IF基板(DE) 資料1\", \"url\": \"https://drive.google.com/file/d/1okwv9-lskK-tD5e2EYpDPJp6iyM2h2B5/view?usp=drivesdk\"}]"},
  {code:"E45", name:"Ｅ４５", gameType:"回胴", m:"M2104A", e:"E2102B", d:"D1401C", de:"DE1802A", relatedDocs:"[{\"category\": \"その他\", \"title\": \"機種説明（E45).xlsx\", \"url\": \"https://drive.google.com/open?id=1PjlTKjaPBgagpRf9z3-Z-OlKo11Uf9v5&usp=drive_copy\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/file/d/12VDUtgbYRQ8EBOjleZx4msZIAjIFbhSV/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"演出基板(E) 資料1\", \"url\": \"https://drive.google.com/file/d/1ONlFLdkMbZL4gKDA8825jXAWsNnXq1dy/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"図柄制御基板(D) 資料1\", \"url\": \"https://drive.google.com/file/d/1b15c4cfEMVNDFMslUnnCIHWP-DLcYVaJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"液晶IF基板(DE) 資料1\", \"url\": \"https://drive.google.com/file/d/1okwv9-lskK-tD5e2EYpDPJp6iyM2h2B5/view?usp=drivesdk\"}]"},
  {code:"E46", relatedDocs:"[{\"category\": \"その他\", \"title\": \"機種説明(E46).xlsx\", \"url\": \"https://drive.google.com/open?id=1AcfD_IKnyVBPzv6uCS7Uq8F2q3l5XSh0&usp=drive_copy\"}]"},
  {code:"E47", name:"E47", gameType:"回胴", m:"M2104A", e:"E2102B", d:"D1401C", de:"DE1802A", relatedDocs:"[{\"category\": \"その他\", \"title\": \"機種説明(E47).xlsx\", \"url\": \"https://drive.google.com/open?id=1hhRsXFgvapZ62WXz2D0XZF94aG5SqBIB&usp=drive_copy\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/file/d/12VDUtgbYRQ8EBOjleZx4msZIAjIFbhSV/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"演出基板(E) 資料1\", \"url\": \"https://drive.google.com/file/d/1ONlFLdkMbZL4gKDA8825jXAWsNnXq1dy/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"図柄制御基板(D) 資料1\", \"url\": \"https://drive.google.com/file/d/1b15c4cfEMVNDFMslUnnCIHWP-DLcYVaJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"液晶IF基板(DE) 資料1\", \"url\": \"https://drive.google.com/file/d/1okwv9-lskK-tD5e2EYpDPJp6iyM2h2B5/view?usp=drivesdk\"}]"},
  {code:"E48", name:"E48 アクセラレータ（とある科学の一方通行 ）", gameType:"回胴", m:"M2104A", e:"E2102B", d:"D1401C", de:"DE1802A", relatedDocs:"[{\"category\": \"その他\", \"title\": \"機種説明(E48).xlsx\", \"url\": \"https://drive.google.com/open?id=1TvmK5rnIAr4y6Cd1MdzjU_6Nglq3AFp-&usp=drive_copy\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料1\", \"url\": \"https://drive.google.com/file/d/12VDUtgbYRQ8EBOjleZx4msZIAjIFbhSV/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"主基板(M) 資料2\", \"url\": \"https://drive.google.com/open?id=189vIuRSpZbgw6Pj3NUN7rYXlcs4Avvh5&usp=drive_copy\"}, {\"category\": \"構成表-見本機\", \"title\": \"演出基板(E) 資料1\", \"url\": \"https://drive.google.com/file/d/1ONlFLdkMbZL4gKDA8825jXAWsNnXq1dy/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"図柄制御基板(D) 資料1\", \"url\": \"https://drive.google.com/file/d/1b15c4cfEMVNDFMslUnnCIHWP-DLcYVaJ/view?usp=drivesdk\"}, {\"category\": \"構成表-見本機\", \"title\": \"液晶IF基板(DE) 資料1\", \"url\": \"https://drive.google.com/file/d/1okwv9-lskK-tD5e2EYpDPJp6iyM2h2B5/view?usp=drivesdk\"}]"},
  {code:"E49", relatedDocs:"[{\"category\": \"その他\", \"title\": \"機種説明（E49).xlsx\", \"url\": \"https://drive.google.com/open?id=1EuHGWjfFmF12dIBk5IPp-cy9s92EyqMp&usp=drive_copy\"}]"},
  {code:"E60", relatedDocs:"[{\"category\": \"その他\", \"title\": \"機種説明E60基板構成.xlsm\", \"url\": \"https://drive.google.com/open?id=1QsXUTbC9vvkmcjjEsmRrYuvEUyjIUd8DxNuLALI8s8I&usp=drive_copy\"}]"},
  {code:"E61", relatedDocs:"[{\"category\": \"その他\", \"title\": \"機種説明E61基板構成.xlsm.xlsx\", \"url\": \"https://drive.google.com/open?id=14F3XleJ3PtDx-xvLQ066daS-uU9jMeo6&usp=drive_copy\"}]"},
  {code:"E62", relatedDocs:"[{\"category\": \"その他\", \"title\": \"機種説明E62基板構成.xlsx\", \"url\": \"https://drive.google.com/open?id=1d2vUje1OeiMR87dOQ0Q2qdaGPiOaHqrR&usp=drive_copy\"}]"},
  {code:"E63"},
  {code:"E64"},
  {code:"E65"},
  {code:"E66"},
  {code:"E67"},
  {code:"E68"},
  {code:"NA07", relatedDocs:"[{\"category\": \"その他\", \"title\": \"機種説明（NA07).xlsx\", \"url\": \"https://drive.google.com/open?id=1F1SXjdiOYUrMbLoyCmxunE7m9hlGfP3k&usp=drive_copy\"}]"},
  {code:"NA09"},
  {code:"NA10"},
  {code:"PD56"}
].map(function(m){ return Object.assign({type:'',brand:'',date:'',m:'',d:'',de:'',e:'',c:'',s:'',r:'',relatedDocs:'',salesUnits:'',salesTarget:'',reuseEcoUnits:'',stockQty:'',photoUrl:'',summary:'',features:'',studyNote:'',manualIds:'',note:'',updatedAt:_now_()}, m); });

var SEED_BOARDS_2 = [
  {id:"M1602B", name:"M1602B", gameType:"回胴", specs:"[{\"label\": \"基板種別\", \"value\": \"主制御基板\"}, {\"label\": \"基板メーカー\", \"value\": \"板橋精機\"}, {\"label\": \"PCB原価\", \"value\": \"¥389\"}, {\"label\": \"仕掛原価\", \"value\": \"¥372\"}, {\"label\": \"仕掛原価(PCB抜き)\", \"value\": \"¥3127\"}, {\"label\": \"仕掛公表単価\", \"value\": \"¥5068\"}, {\"label\": \"公表単価合計\", \"value\": \"¥5068\"}]", relatedDocs:"[{\"category\": \"見積書-PCB原価\", \"title\": \"PCB原価見積\", \"url\": \"https://drive.google.com/file/d/11b3-p496S630sllQWL1rsgJjF5C3MQBC/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛原価\", \"title\": \"仕掛原価見積\", \"url\": \"https://drive.google.com/file/d/1Mz1nSz1flqkQxJxcCmoZJErmME2q9Qil/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛公表単価\", \"title\": \"仕掛公表単価見積\", \"url\": \"https://drive.google.com/file/d/1uvZXgS6Bey-yNhERG5K8N9e3p8w7ECoo/view?usp=drive_link\"}, {\"category\": \"部品表\", \"title\": \"FJ+M1602B.xls\", \"url\": \"https://docs.google.com/spreadsheets/d/1TVhi9zkS3nwL1oulh1HSKsGHRQ96uPM7/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\"}]"},
  {id:"M1805D", name:"M1805D", gameType:"回胴", specs:"[{\"label\": \"基板種別\", \"value\": \"主制御基板\"}, {\"label\": \"基板メーカー\", \"value\": \"板橋精機\"}, {\"label\": \"PCB原価\", \"value\": \"¥475\"}, {\"label\": \"仕掛原価\", \"value\": \"¥397\"}, {\"label\": \"仕掛原価(PCB抜き)\", \"value\": \"¥2399\"}, {\"label\": \"仕掛公表単価\", \"value\": \"¥4375\"}, {\"label\": \"公表単価合計\", \"value\": \"¥4375\"}]", relatedDocs:"[{\"category\": \"見積書-PCB原価\", \"title\": \"PCB原価見積\", \"url\": \"https://drive.google.com/file/d/1rCOuH40rdWhFAvhvPcDQJghQsRt2KyM3/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛原価\", \"title\": \"仕掛原価見積\", \"url\": \"https://drive.google.com/file/d/1y85rPpE9hpbpR_EC19svIrNrVaLnl_kL/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛公表単価\", \"title\": \"仕掛公表単価見積\", \"url\": \"https://drive.google.com/file/d/1Jb2lMsFpgYsuuh7WR6tIBg-33f45-w8s/view?usp=drive_link\"}, {\"category\": \"部品表\", \"title\": \"JJ+M1805D.xls\", \"url\": \"https://docs.google.com/spreadsheets/d/1ctIbjZMET6Do88Td8OGQpWZ6kVomwGI0/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\"}]"},
  {id:"M2104A", name:"M2104A", gameType:"回胴", specs:"[{\"label\": \"基板種別\", \"value\": \"主制御基板\"}, {\"label\": \"基板メーカー\", \"value\": \"板橋精機\"}, {\"label\": \"PCB原価\", \"value\": \"¥491\"}, {\"label\": \"仕掛原価\", \"value\": \"¥581\"}, {\"label\": \"仕掛原価(PCB抜き)\", \"value\": \"¥2592\"}, {\"label\": \"仕掛公表単価\", \"value\": \"¥4729\"}, {\"label\": \"公表単価合計\", \"value\": \"¥4729\"}]", relatedDocs:"[{\"category\": \"見積書-PCB原価\", \"title\": \"PCB原価見積\", \"url\": \"https://drive.google.com/file/d/1nPt597KNmKgMQqYlzbKBo-07nmhNkcdW/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛原価\", \"title\": \"仕掛原価見積\", \"url\": \"https://drive.google.com/file/d/15am5Y3mZjpWBb-W1S5r9BgOE9cbDZoh9/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛公表単価\", \"title\": \"仕掛公表単価見積\", \"url\": \"https://drive.google.com/file/d/1Jb2lMsFpgYsuuh7WR6tIBg-33f45-w8s/view?usp=drive_link\"}, {\"category\": \"基板仕様書\", \"title\": \"14-04_基板仕様書_M2104A_220207.pdf\", \"url\": \"https://drive.google.com/open?id=1KDAp2vY4ME7ySLQJtDRl1P0NzeJ263dz&usp=drive_copy\"}, {\"category\": \"関連基板情報\", \"title\": \"HB237(FJ+M2104A2)\", \"url\": \"https://docs.google.com/spreadsheets/d/1Es9fC3NA0soY4yiGiiGzAYHmmy6quLIGVRa_eiaJD-k/edit?gid=1607512910\"}]"},
  {id:"M2104A1", name:"M2104A1", gameType:"回胴", specs:"[{\"label\": \"基板種別\", \"value\": \"主制御基板\"}, {\"label\": \"公表単価合計\", \"value\": \"¥0\"}]"},
  {id:"M2402A", name:"M2402A", gameType:"回胴", specs:"[{\"label\": \"基板種別\", \"value\": \"主制御基板\"}, {\"label\": \"基板メーカー\", \"value\": \"板橋精機\"}, {\"label\": \"PCB原価\", \"value\": \"¥551\"}, {\"label\": \"仕掛原価\", \"value\": \"¥470\"}, {\"label\": \"仕掛原価(PCB抜き)\", \"value\": \"¥2692\"}, {\"label\": \"PCB公表単価\", \"value\": \"¥762\"}, {\"label\": \"仕掛公表単価\", \"value\": \"¥3716\"}, {\"label\": \"公表単価合計\", \"value\": \"¥4478\"}]", relatedDocs:"[{\"category\": \"見積書-PCB原価\", \"title\": \"PCB原価見積\", \"url\": \"https://drive.google.com/file/d/1aktZfgBRBGGOdkBAPJVqRoFsj0mH001l/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛原価\", \"title\": \"仕掛原価見積\", \"url\": \"https://drive.google.com/file/d/1mFC9uMNEf_qKZf3FX9kTBsd0tzxOPhKD/view?usp=drive_link\"}, {\"category\": \"見積書-PCB公表単価\", \"title\": \"PCB公表単価見積\", \"url\": \"https://drive.google.com/file/d/1G6pOLsXqAUdlz4pV4RxTZuEi2KQ3d-A8/view?usp=sharing\"}, {\"category\": \"部品表\", \"title\": \"FJ+M2402A.xls\", \"url\": \"https://docs.google.com/spreadsheets/d/1oDPm1mQY_ED3ViKtwOpowTxO_CZeE55P/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\"}, {\"category\": \"関連基板情報\", \"title\": \"HB239(FJ+M2402A)\", \"url\": \"https://docs.google.com/spreadsheets/d/1v1oQkGwJVSjHrM3tD6UpS9gJDxw5UDPv5ruHSFyy_EE/edit?usp=drive_link\"}]"},
  {id:"M2401A", name:"M2401A", gameType:"パチンコ", specs:"[{\"label\": \"基板種別\", \"value\": \"主制御基板\"}, {\"label\": \"基板メーカー\", \"value\": \"板橋精機\"}, {\"label\": \"PCB原価\", \"value\": \"¥558\"}, {\"label\": \"仕掛原価\", \"value\": \"¥441\"}, {\"label\": \"仕掛原価(PCB抜き)\", \"value\": \"¥1333\"}, {\"label\": \"PCB公表単価\", \"value\": \"¥731\"}, {\"label\": \"仕掛公表単価\", \"value\": \"¥1753\"}, {\"label\": \"公表単価合計\", \"value\": \"¥2484\"}]", relatedDocs:"[{\"category\": \"見積書-PCB原価\", \"title\": \"PCB原価見積\", \"url\": \"https://drive.google.com/file/d/1kfht4V95itkOsrZv-ZGT3kYHtQ7GuAim/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛原価\", \"title\": \"仕掛原価見積\", \"url\": \"https://drive.google.com/file/d/1P5xQSNWfG02xcXeFj-MxyCMKFK7i5wEV/view?usp=drive_link\"}, {\"category\": \"見積書-PCB公表単価\", \"title\": \"PCB公表単価見積\", \"url\": \"https://drive.google.com/file/d/1G6pOLsXqAUdlz4pV4RxTZuEi2KQ3d-A8/view?usp=sharing\"}, {\"category\": \"部品表\", \"title\": \"FJ+M2104A.xls\", \"url\": \"https://docs.google.com/spreadsheets/d/1zhxCIRYEkUP0bC_CFw-NKd7QgV3iZ2Fe/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\"}, {\"category\": \"基板仕様書\", \"title\": \"14-04_基板仕様書(FJ+M2401)_240828.pdf\", \"url\": \"https://drive.google.com/open?id=1d76-tQ9xY_X1LJ0o4ziv3sRP3nsLxwBz&usp=drive_copy\"}, {\"category\": \"関連基板情報\", \"title\": \"HB238(FJ+M2401A)\", \"url\": \"https://docs.google.com/spreadsheets/d/1TVZeGR0ZH7prggd2IxwQHyrUjXy9vOtpH0P93Is_EdM/edit?usp=drive_link\"}]"},
  {id:"M2503A", name:"M2503A", gameType:"パチンコ", specs:"[{\"label\": \"基板種別\", \"value\": \"主制御基板\"}, {\"label\": \"基板メーカー\", \"value\": \"板橋精機\"}, {\"label\": \"PCB原価\", \"value\": \"¥321\"}, {\"label\": \"仕掛原価\", \"value\": \"¥424.4\"}, {\"label\": \"仕掛原価(PCB抜き)\", \"value\": \"¥1101\"}, {\"label\": \"PCB公表単価\", \"value\": \"¥407\"}, {\"label\": \"仕掛公表単価\", \"value\": \"¥1923\"}, {\"label\": \"公表単価合計\", \"value\": \"¥2330\"}]", relatedDocs:"[{\"category\": \"見積書-仕掛原価\", \"title\": \"仕掛原価見積\", \"url\": \"https://drive.google.com/file/d/1H2TeXo3MNL_h4Naws0iuS9LBCnxGgasA/view?usp=drive_link\"}, {\"category\": \"見積書-PCB公表単価\", \"title\": \"PCB公表単価見積\", \"url\": \"https://drive.google.com/file/d/1G6pOLsXqAUdlz4pV4RxTZuEi2KQ3d-A8/view?usp=sharing\"}, {\"category\": \"関連基板情報\", \"title\": \"HB242(FJ+M2503A)：PFK30\", \"url\": \"https://docs.google.com/spreadsheets/d/1BldU893DlXRwzUbg-759I0fqt7V_KouFb5boAAhQFtc/edit?usp=drive_link\"}]"},
  {id:"D1401C", name:"D1401C", gameType:"共通", specs:"[{\"label\": \"基板種別\", \"value\": \"液晶制御基板（D）\"}, {\"label\": \"基板メーカー\", \"value\": \"リンクステックEPC\"}, {\"label\": \"PCB原価\", \"value\": \"¥490\"}, {\"label\": \"仕掛原価\", \"value\": \"¥434\"}, {\"label\": \"仕掛原価(PCB抜き)\", \"value\": \"¥8571\"}, {\"label\": \"仕掛公表単価\", \"value\": \"¥12449\"}, {\"label\": \"公表単価合計\", \"value\": \"¥12449\"}]", relatedDocs:"[{\"category\": \"見積書-PCB原価\", \"title\": \"PCB原価見積\", \"url\": \"https://drive.google.com/file/d/15fYj2DNt3gfHgxLk9DSIwgN2BJXYlA92/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛原価\", \"title\": \"仕掛原価見積\", \"url\": \"https://drive.google.com/file/d/1fO_XQN1YsD7Om8zLw4XgrauU-OlkIAAr/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛公表単価\", \"title\": \"仕掛公表単価見積\", \"url\": \"https://drive.google.com/file/d/1WImxw_o2uh7B-MINjoTm7h3juu5rHiTu/view?usp=drive_link\"}]"},
  {id:"D2101A", name:"D2101A", gameType:"共通", specs:"[{\"label\": \"基板種別\", \"value\": \"液晶制御基板（D）\"}, {\"label\": \"基板メーカー\", \"value\": \"キョウデン\"}, {\"label\": \"PCB原価\", \"value\": \"¥416\"}, {\"label\": \"仕掛原価\", \"value\": \"¥544\"}, {\"label\": \"仕掛原価(PCB抜き)\", \"value\": \"¥8486\"}, {\"label\": \"PCB公表単価\", \"value\": \"¥537\"}, {\"label\": \"仕掛公表単価\", \"value\": \"¥12605\"}, {\"label\": \"公表単価合計\", \"value\": \"¥13142\"}]", relatedDocs:"[{\"category\": \"見積書-PCB原価\", \"title\": \"PCB原価見積\", \"url\": \"https://drive.google.com/file/d/1QIOrV807nls9kA7e-7YIujyx2wgBN4m4/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛原価\", \"title\": \"仕掛原価見積\", \"url\": \"https://drive.google.com/file/d/1-AxtxwWhMeTF44UnLRASYyBHhxzUZto8/view?usp=drive_link\"}, {\"category\": \"見積書-PCB公表単価\", \"title\": \"PCB公表単価見積\", \"url\": \"https://drive.google.com/file/d/1G6pOLsXqAUdlz4pV4RxTZuEi2KQ3d-A8/view?usp=sharing\"}]"},
  {id:"D2101A1", name:"D2101A1", gameType:"共通", specs:"[{\"label\": \"基板種別\", \"value\": \"液晶制御基板（D）\"}, {\"label\": \"公表単価合計\", \"value\": \"¥0\"}]"},
  {id:"D1901B", name:"D1901B", specs:"[{\"label\": \"基板種別\", \"value\": \"液晶制御基板（D）\"}, {\"label\": \"VDP\", \"value\": \"AG5\"}, {\"label\": \"公表単価合計\", \"value\": \"¥0\"}]", relatedDocs:"[{\"category\": \"基板仕様書\", \"title\": \"14-04_基板設計仕様書(D1901B)_200114.pdf\", \"url\": \"https://drive.google.com/open?id=1ak3syEToHliY4S9nFbY_kofpgYUIiP01&usp=drive_copy\"}, {\"category\": \"関連基板情報\", \"title\": \"HD185(FJ+D1901B)\", \"url\": \"https://docs.google.com/spreadsheets/d/1jHIs4rKjcokUTqp6B2TK8XQpgHKj9lTl96Ae4P29mTA/edit?usp=sharing\"}]"},
  {id:"DE2502A", name:"DE2502A", specs:"[{\"label\": \"基板種別\", \"value\": \"液晶IF基板（DE）\"}, {\"label\": \"基板メーカー\", \"value\": \"リンクステックEPC\"}, {\"label\": \"PCB原価\", \"value\": \"¥354\"}, {\"label\": \"仕掛原価\", \"value\": \"-\"}, {\"label\": \"PCB公表単価\", \"value\": \"¥574\"}, {\"label\": \"仕掛公表単価\", \"value\": \"¥2057\"}, {\"label\": \"公表単価合計\", \"value\": \"¥2631\"}]", relatedDocs:"[{\"category\": \"見積書-PCB原価\", \"title\": \"PCB原価見積\", \"url\": \"https://drive.google.com/file/d/1wQdjwgWnvKLn6sSfc6mTkYOu3EiSakDv/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛原価\", \"title\": \"仕掛原価見積\", \"url\": \"https://drive.google.com/file/d/1gGgQsF05y1ognloVtoIomqfrPU35fk10/view?usp=drive_link\"}, {\"category\": \"見積書-PCB公表単価\", \"title\": \"PCB公表単価見積\", \"url\": \"https://drive.google.com/file/d/1KXqqYnMn4Cj6SUazCjNZpO_BBjcNmXHQ/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛公表単価\", \"title\": \"仕掛公表単価見積\", \"url\": \"https://drive.google.com/file/d/10yT03Rc4qm9h5duXglcOiybpQmoBbh7d/view?usp=drive_link\"}]"},
  {id:"DE1802A", name:"DE1802A", gameType:"パチンコ", specs:"[{\"label\": \"基板種別\", \"value\": \"液晶IF基板（DE）\"}, {\"label\": \"基板メーカー\", \"value\": \"リンクステックEPC\"}, {\"label\": \"PCB原価\", \"value\": \"¥343\"}, {\"label\": \"仕掛原価\", \"value\": \"¥188\"}, {\"label\": \"仕掛原価(PCB抜き)\", \"value\": \"¥1089\"}, {\"label\": \"PCB公表単価\", \"value\": \"¥460\"}, {\"label\": \"仕掛公表単価\", \"value\": \"¥1858\"}, {\"label\": \"公表単価合計\", \"value\": \"¥2318\"}]", relatedDocs:"[{\"category\": \"見積書-PCB原価\", \"title\": \"PCB原価見積\", \"url\": \"https://drive.google.com/file/d/1dsceNf9zVm8wenrV-knVUfzhaMff7CM7/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛原価\", \"title\": \"仕掛原価見積\", \"url\": \"https://drive.google.com/file/d/1T0_jzGe2ZhBdeMZ8DoN2peX5Pj2Ud849/view?usp=drive_link\"}, {\"category\": \"見積書-PCB公表単価\", \"title\": \"PCB公表単価見積\", \"url\": \"https://drive.google.com/file/d/1G32dimsGMItFCM5PLpUZ-PZURGDlmHW1/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛公表単価\", \"title\": \"仕掛公表単価見積\", \"url\": \"https://drive.google.com/file/d/1YR1Fk1AzjX7w3KKMDxBpWPjDzhwO9ulY/view?usp=drive_link\"}, {\"category\": \"部品表\", \"title\": \"FJ+DE1802A.xls\", \"url\": \"https://docs.google.com/spreadsheets/d/1YJP1AQVo1zRs6lJlah4wRPPEJas0KcuU/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\"}]"},
  {id:"DE2101A", name:"DE2101A", gameType:"パチンコ", specs:"[{\"label\": \"基板種別\", \"value\": \"液晶IF基板（DE）\"}, {\"label\": \"基板メーカー\", \"value\": \"板橋精機\"}, {\"label\": \"PCB原価\", \"value\": \"¥352\"}, {\"label\": \"仕掛原価\", \"value\": \"¥247\"}, {\"label\": \"仕掛原価(PCB抜き)\", \"value\": \"¥1348\"}, {\"label\": \"仕掛公表単価\", \"value\": \"¥2729\"}, {\"label\": \"公表単価合計\", \"value\": \"¥2729\"}]", relatedDocs:"[{\"category\": \"見積書-PCB原価\", \"title\": \"PCB原価見積\", \"url\": \"https://drive.google.com/file/d/1FaovLmkbXEvg4TW9Su4sZkQb52Xmm9Q1/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛原価\", \"title\": \"仕掛原価見積\", \"url\": \"https://drive.google.com/file/d/13sgg_G27oijvddUSqbkbqsiPHT8TQMsd/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛公表単価\", \"title\": \"仕掛公表単価見積\", \"url\": \"https://drive.google.com/file/d/1Eexd1mQ4GWbXjTbnwueKeHE48_Xht4U5/view?usp=drive_link\"}, {\"category\": \"部品表\", \"title\": \"FJ+DE2101A.xls\", \"url\": \"https://docs.google.com/spreadsheets/d/15t2pNezczKJ41GXqCMgA8oV1z0w60iuV/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\"}]"},
  {id:"DE2103A1", name:"DE2103A1", gameType:"パチンコ", specs:"[{\"label\": \"基板種別\", \"value\": \"液晶IF基板（DE）\"}, {\"label\": \"公表単価合計\", \"value\": \"¥0\"}]"},
  {id:"DE2101A1◎", name:"DE2101A1◎", gameType:"パチンコ", specs:"[{\"label\": \"基板種別\", \"value\": \"液晶IF基板（DE）\"}, {\"label\": \"基板メーカー\", \"value\": \"リンクステックEPC\"}, {\"label\": \"PCB原価\", \"value\": \"¥399\"}, {\"label\": \"仕掛原価\", \"value\": \"¥217\"}, {\"label\": \"仕掛原価(PCB抜き)\", \"value\": \"¥1307\"}, {\"label\": \"仕掛公表単価\", \"value\": \"¥2362\"}, {\"label\": \"公表単価合計\", \"value\": \"¥2362\"}]", relatedDocs:"[{\"category\": \"見積書-PCB原価\", \"title\": \"PCB原価見積\", \"url\": \"https://drive.google.com/file/d/13XPtL8TfkkmyjffmtvVo_11AUgj3RCz_/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛原価\", \"title\": \"仕掛原価見積\", \"url\": \"https://drive.google.com/file/d/1YDug0D3ygbZqju45_FCVVfgi0G4dzTKS/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛公表単価\", \"title\": \"仕掛公表単価見積\", \"url\": \"https://drive.google.com/file/d/1zrAirdDKR79gMGPhDkmB1Zge_e_dMhRZ/view?usp=sharing\"}, {\"category\": \"部品表\", \"title\": \"FJ+DE2101A1.xls\", \"url\": \"https://docs.google.com/spreadsheets/d/1vV5n5l67g5XTLzlDkzkj-EQUYGXeNUDd/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\"}, {\"category\": \"見積書-過去見積\", \"title\": \"過去見積\", \"url\": \"https://drive.google.com/file/d/1RKlA-tDzJFdMvjdw5FLmVM3JpXpqEekk/view?usp=drive_link\"}, {\"category\": \"関連基板情報\", \"title\": \"HD200(FJ+DE2101A_VerB)\", \"url\": \"https://docs.google.com/spreadsheets/d/1J1up4vF6RgiNwHBii6JZBTCPQcP1RRoMNoqnetNjAF0/edit\"}]"},
  {id:"DE2101A", name:"DE2101A", gameType:"パチンコ", specs:"[{\"label\": \"基板種別\", \"value\": \"液晶IF基板（DE）\"}, {\"label\": \"基板メーカー\", \"value\": \"リンクステックEPC\"}, {\"label\": \"公表単価合計\", \"value\": \"¥0\"}]", relatedDocs:"[{\"category\": \"関連基板情報\", \"title\": \"HD239(FJ+DE2502A)\", \"url\": \"https://docs.google.com/spreadsheets/d/1nTIsjMEFaoV9ymukW73wtOAzO_gLybb1zU3RccawkLM/edit?usp=drive_link\"}]"},
  {id:"DE1607A", name:"DE1607A", gameType:"回胴", specs:"[{\"label\": \"基板種別\", \"value\": \"液晶IF基板（DE）\"}, {\"label\": \"基板メーカー\", \"value\": \"リンクステックEPC\"}, {\"label\": \"PCB原価\", \"value\": \"¥343\"}, {\"label\": \"仕掛原価\", \"value\": \"¥175\"}, {\"label\": \"仕掛原価(PCB抜き)\", \"value\": \"¥984\"}, {\"label\": \"仕掛公表単価\", \"value\": \"¥2054\"}, {\"label\": \"公表単価合計\", \"value\": \"¥2054\"}]", relatedDocs:"[{\"category\": \"見積書-PCB原価\", \"title\": \"PCB原価見積\", \"url\": \"https://drive.google.com/file/d/1LEhylDbH8zO45QR5aB-x64A5BuK8FJNM/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛原価\", \"title\": \"仕掛原価見積\", \"url\": \"https://drive.google.com/file/d/1vPNFG2wKHJgpddjNKucDGfl6a3TGI-vI/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛公表単価\", \"title\": \"仕掛公表単価見積\", \"url\": \"https://drive.google.com/file/d/1Jb2lMsFpgYsuuh7WR6tIBg-33f45-w8s/view?usp=drive_link\"}, {\"category\": \"部品表\", \"title\": \"FJ+DE1607A.xls\", \"url\": \"https://docs.google.com/spreadsheets/d/1zPF5_cCLnYwfW1sH2OHt99a-ASHLQirC/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\"}, {\"category\": \"見積書-過去見積\", \"title\": \"過去見積\", \"url\": \"https://drive.google.com/file/d/1xxTQsOsj1KSF6Obld0Z_GzeqiNcX-Fk8/view?usp=drive_link\"}]"},
  {id:"DE1606A", name:"DE1606A", gameType:"回胴", specs:"[{\"label\": \"基板種別\", \"value\": \"液晶IF基板（DE）\"}, {\"label\": \"基板メーカー\", \"value\": \"リンクステックEPC\"}, {\"label\": \"PCB原価\", \"value\": \"¥343\"}, {\"label\": \"仕掛原価\", \"value\": \"¥201\"}, {\"label\": \"仕掛原価(PCB抜き)\", \"value\": \"¥1241\"}, {\"label\": \"仕掛公表単価\", \"value\": \"¥2486\"}, {\"label\": \"公表単価合計\", \"value\": \"¥2486\"}]", relatedDocs:"[{\"category\": \"見積書-PCB原価\", \"title\": \"PCB原価見積\", \"url\": \"https://drive.google.com/file/d/1zWoSYYNQq5XI9eMomTBRfbeIhx72ZSMQ/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛原価\", \"title\": \"仕掛原価見積\", \"url\": \"https://drive.google.com/file/d/1eg-ns8KB37mS5fEDgZtKDtIiGgxwkcqh/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛公表単価\", \"title\": \"仕掛公表単価見積\", \"url\": \"https://drive.google.com/file/d/1Jb2lMsFpgYsuuh7WR6tIBg-33f45-w8s/view?usp=drive_link\"}, {\"category\": \"部品表\", \"title\": \"FJ+DE1606A.xls\", \"url\": \"https://docs.google.com/spreadsheets/d/1xpWH3TvPO1BjauLdr_t76lNa6FSQ-ctu/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\"}, {\"category\": \"見積書-過去見積\", \"title\": \"過去見積\", \"url\": \"https://drive.google.com/file/d/1xxTQsOsj1KSF6Obld0Z_GzeqiNcX-Fk8/view?usp=drive_link\"}]"},
  {id:"DE1802A", name:"DE1802A", gameType:"回胴", specs:"[{\"label\": \"基板種別\", \"value\": \"液晶IF基板（DE）\"}, {\"label\": \"基板メーカー\", \"value\": \"リンクステックEPC\"}, {\"label\": \"公表単価合計\", \"value\": \"¥0\"}]"},
  {id:"E1501C", name:"E1501C", gameType:"回胴", specs:"[{\"label\": \"基板種別\", \"value\": \"演出基板（E）\"}, {\"label\": \"基板メーカー\", \"value\": \"リンクステックEPC\"}, {\"label\": \"PCB原価\", \"value\": \"¥699\"}, {\"label\": \"仕掛原価\", \"value\": \"¥493\"}, {\"label\": \"仕掛原価(PCB抜き)\", \"value\": \"¥5472\"}, {\"label\": \"仕掛公表単価\", \"value\": \"¥7133\"}, {\"label\": \"公表単価合計\", \"value\": \"¥7133\"}]", relatedDocs:"[{\"category\": \"見積書-PCB原価\", \"title\": \"PCB原価見積\", \"url\": \"https://drive.google.com/file/d/14zIB4jPqCdgbScHIkzqOcmW-qRRqnk_d/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛原価\", \"title\": \"仕掛原価見積\", \"url\": \"https://drive.google.com/file/d/1Dhx2tPtQzaPvtU3LAscgPj3H2V4f2wXI/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛公表単価\", \"title\": \"仕掛公表単価見積\", \"url\": \"https://drive.google.com/file/d/17_QvdIKi3uG0uA6Pi0zgI61NVl-N53F2/view?usp=drive_link\"}, {\"category\": \"部品表\", \"title\": \"FJ+E1501C.xls\", \"url\": \"https://docs.google.com/spreadsheets/d/1vw1TBQpzDQUQPmTkwA_ffJUrY6DWANpH/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\"}]"},
  {id:"E1901B", name:"E1901B", specs:"[{\"label\": \"基板種別\", \"value\": \"演出基板（E）\"}, {\"label\": \"基板メーカー\", \"value\": \"板橋精機\"}, {\"label\": \"PCB原価\", \"value\": \"¥556\"}, {\"label\": \"仕掛原価\", \"value\": \"¥633\"}, {\"label\": \"仕掛原価(PCB抜き)\", \"value\": \"¥5019\"}, {\"label\": \"仕掛公表単価\", \"value\": \"¥7100\"}, {\"label\": \"公表単価合計\", \"value\": \"¥7100\"}]", relatedDocs:"[{\"category\": \"見積書-PCB原価\", \"title\": \"PCB原価見積\", \"url\": \"https://drive.google.com/file/d/1ULqYmRBiab7Vftmz_ILVPGcRwgNmfNtr/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛原価\", \"title\": \"仕掛原価見積\", \"url\": \"https://drive.google.com/file/d/1kvkHSsC8meZsPf2gEf_FfpvCGoPN8VQB/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛公表単価\", \"title\": \"仕掛公表単価見積\", \"url\": \"https://drive.google.com/file/d/1Jb2lMsFpgYsuuh7WR6tIBg-33f45-w8s/view?usp=drive_link\"}, {\"category\": \"部品表\", \"title\": \"FJ+E1901B.xls\", \"url\": \"https://docs.google.com/spreadsheets/d/1jHbzfcKOoTqJAWfZ6LLQh9IMQgP-l8__/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\"}]"},
  {id:"E1902B", name:"E1902B", gameType:"回胴", specs:"[{\"label\": \"基板種別\", \"value\": \"演出基板（E）\"}, {\"label\": \"基板メーカー\", \"value\": \"板橋精機\"}, {\"label\": \"PCB原価\", \"value\": \"¥742\"}, {\"label\": \"仕掛原価\", \"value\": \"¥654\"}, {\"label\": \"仕掛原価(PCB抜き)\", \"value\": \"¥5244\"}, {\"label\": \"仕掛公表単価\", \"value\": \"¥7571\"}, {\"label\": \"公表単価合計\", \"value\": \"¥7571\"}]", relatedDocs:"[{\"category\": \"見積書-PCB原価\", \"title\": \"PCB原価見積\", \"url\": \"https://drive.google.com/file/d/1a6r5rmnEPwctg8cw3VGXcqn4uWozTYJx/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛原価\", \"title\": \"仕掛原価見積\", \"url\": \"https://drive.google.com/file/d/1p0UwrV2Gjg2cI5GTT366Lm0R9iD-iKX2/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛公表単価\", \"title\": \"仕掛公表単価見積\", \"url\": \"https://drive.google.com/file/d/1Jb2lMsFpgYsuuh7WR6tIBg-33f45-w8s/view?usp=drive_link\"}, {\"category\": \"部品表\", \"title\": \"FJ+E1902B.xls\", \"url\": \"https://docs.google.com/spreadsheets/d/19r8lLSl5RqMRjlhhEsnbiQN7IaVKUgvq/edit?usp=drive_link&ouid=102382963600657219027&rtpof=true&sd=true\"}]"},
  {id:"E2102B", name:"E2102B", gameType:"回胴", specs:"[{\"label\": \"基板種別\", \"value\": \"演出基板（E）\"}, {\"label\": \"基板メーカー\", \"value\": \"板橋精機\"}, {\"label\": \"PCB原価\", \"value\": \"¥887\"}, {\"label\": \"仕掛原価\", \"value\": \"¥666\"}, {\"label\": \"PCB公表単価\", \"value\": \"¥1097\"}, {\"label\": \"仕掛公表単価\", \"value\": \"¥8089\"}, {\"label\": \"公表単価合計\", \"value\": \"¥9186\"}]", relatedDocs:"[{\"category\": \"見積書-PCB原価\", \"title\": \"PCB原価見積\", \"url\": \"https://drive.google.com/file/d/1oaVjQZaRU9hCF48BD2Hyji29mbUBv9OY/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛原価\", \"title\": \"仕掛原価見積\", \"url\": \"https://drive.google.com/file/d/1nnFHKl_48wQ9onga4K-A0-LJdflJ8jGz/view?usp=drive_link\"}, {\"category\": \"見積書-PCB公表単価\", \"title\": \"PCB公表単価見積\", \"url\": \"https://drive.google.com/file/d/1r9Ftnt4WyC7JgAkqz8pUju0c0-sP_My-/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛公表単価\", \"title\": \"仕掛公表単価見積\", \"url\": \"https://drive.google.com/file/d/1qlTwXtkQONeJXfKZEmL-a6YpVdyRKlVz/view?usp=drive_link\"}]"},
  {id:"E2002B", name:"E2002B", gameType:"パチンコ", specs:"[{\"label\": \"基板種別\", \"value\": \"演出基板（E）\"}, {\"label\": \"基板メーカー\", \"value\": \"板橋精機\"}, {\"label\": \"PCB原価\", \"value\": \"¥556\"}, {\"label\": \"仕掛原価\", \"value\": \"¥648\"}, {\"label\": \"仕掛公表単価\", \"value\": \"¥7244\"}, {\"label\": \"公表単価合計\", \"value\": \"¥7244\"}]", relatedDocs:"[{\"category\": \"見積書-PCB原価\", \"title\": \"PCB原価見積\", \"url\": \"https://drive.google.com/file/d/117exYeAoGwIwIie97S54_eoQTglxYaVd/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛原価\", \"title\": \"仕掛原価見積\", \"url\": \"https://drive.google.com/file/d/1Qqb9R3n123qadCfArEmhDutW3671UWbY/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛公表単価\", \"title\": \"仕掛公表単価見積\", \"url\": \"https://drive.google.com/file/d/1Jb2lMsFpgYsuuh7WR6tIBg-33f45-w8s/view?usp=drive_link\"}, {\"category\": \"基板仕様書\", \"title\": \"14-04_基板設計仕様書(E2002B)_210708.pdf\", \"url\": \"https://drive.google.com/open?id=1SLVpGbI1Wl8xZEiHbzyTlhksIblMKLT-&usp=drive_copy\"}, {\"category\": \"関連基板情報\", \"title\": \"HG181(FJ+E2002B)\", \"url\": \"https://docs.google.com/spreadsheets/d/1_GiYUmBAH7jMIydMcDC4-jUbk4Kf5omNHwezFRM-rtM/edit?usp=sharing\"}]"},
  {id:"E2101B", name:"E2101B", gameType:"パチンコ", specs:"[{\"label\": \"基板種別\", \"value\": \"演出基板（E）\"}, {\"label\": \"基板メーカー\", \"value\": \"板橋精機\"}, {\"label\": \"PCB原価\", \"value\": \"¥887\"}, {\"label\": \"仕掛原価\", \"value\": \"¥539\"}, {\"label\": \"PCB公表単価\", \"value\": \"¥1124\"}, {\"label\": \"仕掛公表単価\", \"value\": \"¥8049\"}, {\"label\": \"公表単価合計\", \"value\": \"¥9173\"}]", relatedDocs:"[{\"category\": \"見積書-PCB原価\", \"title\": \"PCB原価見積\", \"url\": \"https://drive.google.com/file/d/1V4v8im0vQoXJvPr5cXVm3W-20eK8uQJK/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛原価\", \"title\": \"仕掛原価見積\", \"url\": \"https://drive.google.com/file/d/1iNcm2W4ioiuO3Y7CZvw72Q6yG_HHL2yh/view?usp=drive_link\"}, {\"category\": \"見積書-PCB公表単価\", \"title\": \"PCB公表単価見積\", \"url\": \"https://drive.google.com/file/d/1G6pOLsXqAUdlz4pV4RxTZuEi2KQ3d-A8/view?usp=sharing\"}, {\"category\": \"基板仕様書\", \"title\": \"14-04_基板設計仕様書(E2101B)_211125.pdf\", \"url\": \"https://drive.google.com/open?id=1C_eUNoQdl2T_-Ph9E-29jmJqFGgzzzvS&usp=drive_copy\"}, {\"category\": \"関連基板情報\", \"title\": \"HG183(FJ+E2101B)\", \"url\": \"https://docs.google.com/spreadsheets/d/1V3FFrs-FvskrQlBPTINaaRvU6SBJf310V0gRD6zHdMM/edit?usp=sharing\"}]"},
  {id:"E2102B", name:"E2102B", gameType:"パチンコ", specs:"[{\"label\": \"基板種別\", \"value\": \"演出基板（E）\"}, {\"label\": \"基板メーカー\", \"value\": \"板橋精機\"}, {\"label\": \"公表単価合計\", \"value\": \"¥0\"}]", relatedDocs:"[{\"category\": \"基板仕様書\", \"title\": \"14-04_基板設計仕様書_FJ+E2102B.pdf\", \"url\": \"https://drive.google.com/open?id=1fATcvVULFz09TDOLX6JAfC3cA12wbHAf&usp=drive_copy\"}, {\"category\": \"関連基板情報\", \"title\": \"HG182(FJ+E2102B)\", \"url\": \"https://docs.google.com/spreadsheets/d/1ahX4yavzLwEOsvaTHujLklz2XDynGnZ9_EV0Iy9EElA/edit\"}]"},
  {id:"E2301B", name:"E2301B", gameType:"パチンコ", specs:"[{\"label\": \"基板種別\", \"value\": \"演出基板（E）\"}, {\"label\": \"基板メーカー\", \"value\": \"板橋精機\"}, {\"label\": \"PCB原価\", \"value\": \"¥734\"}, {\"label\": \"仕掛原価\", \"value\": \"¥509\"}, {\"label\": \"PCB公表単価\", \"value\": \"¥956\"}, {\"label\": \"仕掛公表単価\", \"value\": \"¥4942\"}, {\"label\": \"公表単価合計\", \"value\": \"¥5898\"}]", relatedDocs:"[{\"category\": \"見積書-PCB原価\", \"title\": \"PCB原価見積\", \"url\": \"https://drive.google.com/file/d/1U7ox1RjVWuHwsXOSdzFjd6Z9oVb6g6ra/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛原価\", \"title\": \"仕掛原価見積\", \"url\": \"https://drive.google.com/file/d/1zfTrUyPQiSCd6kkvsTTu3qyOgRB32Ejk/view?usp=drive_link\"}, {\"category\": \"見積書-PCB公表単価\", \"title\": \"PCB公表単価見積\", \"url\": \"https://drive.google.com/file/d/1G6pOLsXqAUdlz4pV4RxTZuEi2KQ3d-A8/view?usp=sharing\"}, {\"category\": \"基板仕様書\", \"title\": \"14-04_基板設計仕様書(E2301B)_240228.pdf\", \"url\": \"https://drive.google.com/open?id=1Nsk0yDrZ0Y6v6qck7R0WGvPBp27CBpGr&usp=drive_copy\"}, {\"category\": \"関連基板情報\", \"title\": \"HG184(FJ+E2301A)\", \"url\": \"https://docs.google.com/spreadsheets/d/1E-5iyLQ0ciHN1KWj24MEcrmStrjaIcvLgylKw6MkoaU/edit\"}]"},
  {id:"E2501B", name:"E2501B", gameType:"パチンコ", specs:"[{\"label\": \"基板種別\", \"value\": \"演出基板（E）\"}, {\"label\": \"基板メーカー\", \"value\": \"板橋精機\"}, {\"label\": \"公表単価合計\", \"value\": \"¥0\"}]", relatedDocs:"[{\"category\": \"基板仕様書\", \"title\": \"14-04_基板設計仕様書(TEST2501)_250619.pdf\", \"url\": \"https://drive.google.com/open?id=1Zpnz_dA23Ytx5hHLTLY4-l_nzB_gAYCK&usp=drive_copy\"}, {\"category\": \"関連基板情報\", \"title\": \"FJ+E2501B\", \"url\": \"https://docs.google.com/spreadsheets/d/1oc9OhjGk8Y_Eory_Rjpbv5NqgZPsno0oT-8mZOzrJ0M/edit?usp=drive_link\"}]"},
  {id:"E2503A", name:"E2503A", gameType:"パチンコ", specs:"[{\"label\": \"基板種別\", \"value\": \"演出基板（E）\"}, {\"label\": \"基板メーカー\", \"value\": \"板橋精機\"}, {\"label\": \"公表単価合計\", \"value\": \"¥0\"}]", relatedDocs:"[{\"category\": \"関連基板情報\", \"title\": \"FJ+E2503A\", \"url\": \"https://docs.google.com/spreadsheets/d/1d50PCBb1tqiLFNT2zS_QMBktaXbF9RgwNHpd_gRn0J8/edit?gid=1607512910\"}]"},
  {id:"E2501B", name:"E2501B", gameType:"パチンコ", specs:"[{\"label\": \"基板種別\", \"value\": \"演出基板（E）\"}, {\"label\": \"VDP\", \"value\": \"AG6R\"}, {\"label\": \"基板メーカー\", \"value\": \"板橋精機\"}, {\"label\": \"PCB原価\", \"value\": \"¥699\"}, {\"label\": \"仕掛原価\", \"value\": \"¥476\"}, {\"label\": \"PCB公表単価\", \"value\": \"¥974\"}, {\"label\": \"仕掛公表単価\", \"value\": \"¥4700\"}, {\"label\": \"公表単価合計\", \"value\": \"¥5674\"}]", relatedDocs:"[{\"category\": \"見積書-PCB原価\", \"title\": \"PCB原価見積\", \"url\": \"https://drive.google.com/file/d/1gxyOQ7r3sA-fFqoY0kX9MGGRU-51ds63/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛原価\", \"title\": \"仕掛原価見積\", \"url\": \"https://drive.google.com/file/d/1QyJwCusb4JpZnA1axycSPm_EoRbI--m9/view?usp=drive_link\"}, {\"category\": \"見積書-PCB公表単価\", \"title\": \"PCB公表単価見積\", \"url\": \"https://drive.google.com/file/d/1G6pOLsXqAUdlz4pV4RxTZuEi2KQ3d-A8/view?usp=sharing\"}, {\"category\": \"基板仕様書\", \"title\": \"14-04_基板設計仕様書(TEST2501)_250619.pdf\", \"url\": \"https://drive.google.com/open?id=1Zpnz_dA23Ytx5hHLTLY4-l_nzB_gAYCK&usp=drive_copy\"}]"},
  {id:"E2503B", name:"E2503B", specs:"[{\"label\": \"基板種別\", \"value\": \"演出基板（E）\"}, {\"label\": \"基板メーカー\", \"value\": \"リンクステックEPC\"}, {\"label\": \"PCB原価\", \"value\": \"¥814\"}, {\"label\": \"仕掛原価\", \"value\": \"-\"}, {\"label\": \"PCB公表単価\", \"value\": \"¥1118\"}, {\"label\": \"仕掛公表単価\", \"value\": \"¥8449\"}, {\"label\": \"公表単価合計\", \"value\": \"¥9567\"}]", relatedDocs:"[{\"category\": \"見積書-PCB原価\", \"title\": \"PCB原価見積\", \"url\": \"https://drive.google.com/file/d/1wLvTHiAlv8S49vywYjB38V7m-tzssGNg/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛原価\", \"title\": \"仕掛原価見積\", \"url\": \"https://drive.google.com/file/d/1XGPLeAc_-etEnczNNsw0HjyS2ZBtpLYD/view?usp=drive_link\"}, {\"category\": \"見積書-PCB公表単価\", \"title\": \"PCB公表単価見積\", \"url\": \"https://drive.google.com/file/d/1Yci0PDwaKkU-KGKS4KMBHWmJzzcx-cSP/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛公表単価\", \"title\": \"仕掛公表単価見積\", \"url\": \"https://drive.google.com/file/d/1F8wCDjvMiRUX6wEXbCDG1vOkYb3T6NlH/view?usp=drive_link\"}]"},
  {id:"C1601E", name:"C1601E", specs:"[{\"label\": \"基板種別\", \"value\": \"サブ制御基板（C）\"}, {\"label\": \"公表単価合計\", \"value\": \"¥0\"}]"},
  {id:"C1901A", name:"C1901A", specs:"[{\"label\": \"基板種別\", \"value\": \"サブ制御基板（C）\"}, {\"label\": \"公表単価合計\", \"value\": \"¥0\"}]", relatedDocs:"[{\"category\": \"基板仕様書\", \"title\": \"14-04_(仮)基板仕様書_FJ C1901A_191024.pdf\", \"url\": \"https://drive.google.com/open?id=1ZgpxtAE9XBD1HIPBbjrQriLeH4tdHXsa&usp=drive_copy\"}]"},
  {id:"C2101B", name:"C2101B", specs:"[{\"label\": \"基板種別\", \"value\": \"サブ制御基板（C）\"}, {\"label\": \"基板メーカー\", \"value\": \"板橋精機\"}, {\"label\": \"PCB原価\", \"value\": \"¥378\"}, {\"label\": \"仕掛原価\", \"value\": \"¥343\"}, {\"label\": \"PCB公表単価\", \"value\": \"¥445\"}, {\"label\": \"仕掛公表単価\", \"value\": \"¥1902\"}, {\"label\": \"公表単価合計\", \"value\": \"¥2347\"}]", relatedDocs:"[{\"category\": \"見積書-PCB原価\", \"title\": \"PCB原価見積\", \"url\": \"https://drive.google.com/file/d/1IJCf5xVXLY8rlxbwDe9Vnw5-Ijz2cojR/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛原価\", \"title\": \"仕掛原価見積\", \"url\": \"https://drive.google.com/file/d/1BCO1u9mJvlvDEj_tdsM8aZld19JXvXFn/view?usp=drive_link\"}, {\"category\": \"見積書-PCB公表単価\", \"title\": \"PCB公表単価見積\", \"url\": \"https://drive.google.com/file/d/1r9Ftnt4WyC7JgAkqz8pUju0c0-sP_My-/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛公表単価\", \"title\": \"仕掛公表単価見積\", \"url\": \"https://drive.google.com/file/d/1qlTwXtkQONeJXfKZEmL-a6YpVdyRKlVz/view?usp=drive_link\"}]"},
  {id:"C2401A", name:"C2401A", specs:"[{\"label\": \"基板種別\", \"value\": \"サブ制御基板（C）\"}, {\"label\": \"基板メーカー\", \"value\": \"板橋精機\"}, {\"label\": \"PCB原価\", \"value\": \"¥749\"}, {\"label\": \"仕掛原価\", \"value\": \"¥630.4\"}, {\"label\": \"PCB公表単価\", \"value\": \"¥941\"}, {\"label\": \"仕掛公表単価\", \"value\": \"¥3624\"}, {\"label\": \"公表単価合計\", \"value\": \"¥4565\"}]", relatedDocs:"[{\"category\": \"見積書-PCB原価\", \"title\": \"PCB原価見積\", \"url\": \"https://drive.google.com/file/d/1LCfXUQddHT1rB5ijDctqOu1NIelG7-sH/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛原価\", \"title\": \"仕掛原価見積\", \"url\": \"https://drive.google.com/file/d/1P4lBSHPR51GhSx9N6-4-WE4bYLkYAP47/view?usp=drive_link\"}, {\"category\": \"見積書-PCB公表単価\", \"title\": \"PCB公表単価見積\", \"url\": \"https://drive.google.com/file/d/1G6pOLsXqAUdlz4pV4RxTZuEi2KQ3d-A8/view?usp=sharing\"}]"},
  {id:"C2202B", name:"C2202B", specs:"[{\"label\": \"基板種別\", \"value\": \"サブ制御基板（C）\"}, {\"label\": \"基板メーカー\", \"value\": \"板橋精機\"}, {\"label\": \"公表単価合計\", \"value\": \"¥0\"}]", relatedDocs:"[{\"category\": \"基板仕様書\", \"title\": \"14-04_FJ+C2202A_基板仕様書_230117.pdf\", \"url\": \"https://drive.google.com/open?id=1ciWl8sRL5TcUsXkA0mx-r2VlzKoUu7cN&usp=drive_copy\"}, {\"category\": \"関連基板情報\", \"title\": \"HC030(FJ+C2202B)：PF45\", \"url\": \"https://docs.google.com/spreadsheets/d/1lhz181exIl7Tynhn7c3QK0vTP5vdu9rof7hUWzTq6vc/edit\"}]"},
  {id:"C2501C", name:"C2501C", specs:"[{\"label\": \"基板種別\", \"value\": \"サブ制御基板（C）\"}, {\"label\": \"基板メーカー\", \"value\": \"板橋精機\"}, {\"label\": \"PCB原価\", \"value\": \"¥728\"}, {\"label\": \"仕掛原価\", \"value\": \"¥502\"}, {\"label\": \"PCB公表単価\", \"value\": \"¥854\"}, {\"label\": \"仕掛公表単価\", \"value\": \"¥3169\"}, {\"label\": \"公表単価合計\", \"value\": \"¥4023\"}]", relatedDocs:"[{\"category\": \"見積書-PCB原価\", \"title\": \"PCB原価見積\", \"url\": \"https://drive.google.com/file/d/1mJLfOXS3QDYMRa5XTFZYT8tYM4r99mZ6/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛原価\", \"title\": \"仕掛原価見積\", \"url\": \"https://drive.google.com/file/d/1g763qKIlEAfOG0b8YUTgcGfN__gq4EwG/view?usp=drive_link\"}, {\"category\": \"見積書-PCB公表単価\", \"title\": \"PCB公表単価見積\", \"url\": \"https://drive.google.com/file/d/1G6pOLsXqAUdlz4pV4RxTZuEi2KQ3d-A8/view?usp=sharing\"}, {\"category\": \"関連基板情報\", \"title\": \"HC036(FJ+C2501C)：PFK30\", \"url\": \"https://docs.google.com/spreadsheets/d/1k6SkKeJy6TcYcYm0kkD2MqN0TnK749JwRDlFPGs-BCw/edit?gid=1607512910\"}]"},
  {id:"C2502A", name:"C2502A", specs:"[{\"label\": \"基板種別\", \"value\": \"サブ制御基板（C）\"}, {\"label\": \"基板メーカー\", \"value\": \"板橋精機\"}, {\"label\": \"PCB原価\", \"value\": \"¥352\"}, {\"label\": \"仕掛原価\", \"value\": \"¥350\"}, {\"label\": \"PCB公表単価\", \"value\": \"¥491\"}, {\"label\": \"仕掛公表単価\", \"value\": \"¥1926\"}, {\"label\": \"公表単価合計\", \"value\": \"¥2417\"}]", relatedDocs:"[{\"category\": \"見積書-PCB公表単価\", \"title\": \"PCB公表単価見積\", \"url\": \"https://drive.google.com/file/d/1jkcA4Yf3YRFaomKLaGnrJ7iKNfh4E9ds/view?usp=drive_link\"}, {\"category\": \"見積書-仕掛公表単価\", \"title\": \"仕掛公表単価見積\", \"url\": \"https://drive.google.com/file/d/1DcaQt2h94U58VzbUohZe14js9Ho5sTr-/view?usp=drive_link\"}, {\"category\": \"関連基板情報\", \"title\": \"HC037(FJ+C2502A)：SFK18\", \"url\": \"https://docs.google.com/spreadsheets/d/1rRZjy2-3OFre2n6xHOABftc6ZN7mAq9KLv507igX1F4/edit?gid=1607512910\"}]"},
  {id:"C2501B", name:"C2501B", specs:"[{\"label\": \"基板種別\", \"value\": \"サブ制御基板（C）\"}, {\"label\": \"公表単価合計\", \"value\": \"¥0\"}]", relatedDocs:"[{\"category\": \"関連基板情報\", \"title\": \"HC035(FJ+C2501B)：PFK30\", \"url\": \"https://docs.google.com/spreadsheets/d/1Iu0MTSOaFSRPRnHjVTFCKJjV2EmNuTs3qP86j_OASOU/edit?usp=drive_link\"}]"},
  {id:"SNB5163A-00", name:"SNB5163A-00", specs:"[{\"label\": \"基板種別\", \"value\": \"SNB基板\"}, {\"label\": \"仕掛原価\", \"value\": \"https://drive.google.com/file/d/1lxty11pjp_7TkTUsZEFc4P2DxQVkjbEw/view?usp=drive_link\"}, {\"label\": \"公表単価合計\", \"value\": \"¥8146\"}]", relatedDocs:"[{\"category\": \"見積書-仕掛原価\", \"title\": \"仕掛原価見積\", \"url\": \"https://drive.google.com/file/d/1lxty11pjp_7TkTUsZEFc4P2DxQVkjbEw/view?usp=drive_link\"}]"}
].map(function(b){ return Object.assign({name:'',category:'',brand:'',gameType:'',version:'',status:'',specs:'',revisions:'',relatedDocs:'',photoUrl:'',summary:'',mainParts:'',studyPoint:'',relatedModels:'',note:'',updatedAt:_now_()}, b); });

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
