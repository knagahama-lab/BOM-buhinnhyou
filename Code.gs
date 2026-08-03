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
