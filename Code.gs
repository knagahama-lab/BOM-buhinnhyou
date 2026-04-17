/**
 * BOM Pro — GAS バックエンド
 * スプレッドシートをDBとして使用
 * 全データを1回のAPIコールで取得・保存するシンプル設計
 */

// ── シート名定数 ──
var S = {
  PARTS:    '部品マスタ',
  BOM:      'BOM',
  MACHINES: '機種マスタ',
  BOARDS:   '基板マスタ',
  STOCK:    '在庫',
  CHANGELOG:'変更履歴',
  SSLINKS:  'SS連携設定',
  PRICE_H:  '価格履歴',
  META:     'メタ'
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
      .setTitle('BOM Pro — パチンコ基板部品表管理')
      .addMetaTag('viewport','width=device-width,initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch(e) {
    return HtmlService.createHtmlOutput(
      '<body style="font-family:sans-serif;padding:40px;background:#f7f7f8;">' +
      '<h2 style="color:#dc2626;">⚠ 起動エラー</h2>' +
      '<p style="margin-top:12px;color:#52525b;">' + e.message + '</p>' +
      '<hr style="margin:20px 0;border:none;border-top:1px solid #e2e2e5;"/>' +
      '<p style="font-size:13px;color:#a1a1aa;line-height:1.8;">確認事項：<br>' +
      '① GAS「プロジェクトの設定」→「スクリプトプロパティ」に <b>BOARD_SS_ID</b> を設定する<br>' +
      '② 再デプロイ（新バージョン）する</p></body>'
    ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── スプレッドシート取得 ──
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
  PARTS:    ['部品コード','部品名','メーカー名','公表単価','仕入先','廃番フラグ','RoHS対応','認定部品','備考','更新日時'],
  BOM:      ['基板ID','部品コード','使用数量','単位','レベル','親ID','種別','備考'],
  MACHINES: ['機種コード','機種名','種類','ブランド','発売日','M基板','D基板','DE基板','E基板','C基板','S基板','備考','更新日時'],
  BOARDS:   ['基板ID','基板名','基板分類','バージョン','ステータス','作成日','備考','更新日時'],
  STOCK:    ['部品コード','在庫数','安全在庫','備考','更新日時'],
  CHANGELOG:['日時','担当者','ECN番号','対象','変更種別','変更前','変更後','備考'],
  SSLINKS:  ['名前','SSID','シート名','種別','備考'],
  PRICE_H:  ['日時','部品コード','旧単価','新単価','変化率']
};

// ══════════════════════════════════════════════════
// メインAPI: 全データを1回で取得
// ══════════════════════════════════════════════════
function apiLoadAll() {
  return _wrap(function() {
    var parts = {}, bom = {}, machines = {}, boards = {}, stock = {};

    _read(S.PARTS, H.PARTS).forEach(function(r) {
      if (r['部品コード']) parts[r['部品コード']] = {
        code: String(r['部品コード']),
        name: String(r['部品名'] || ''),
        maker: String(r['メーカー名'] || ''),
        price: String(r['公表単価'] || ''),
        supplier: String(r['仕入先'] || ''),
        eol: String(r['廃番フラグ'] || '0'),
        rohs: String(r['RoHS対応'] || '0'),
        certified: String(r['認定部品'] || '0'),
        note: String(r['備考'] || ''),
        updatedAt: String(r['更新日時'] || '')
      };
    });

    _read(S.BOM, H.BOM).forEach(function(r) {
      var bid = r['基板ID'], pc = r['部品コード'];
      if (bid && pc) {
        var key = bid + '|' + pc;
        bom[key] = {
          boardId: String(bid), partCode: String(pc),
          qty: String(r['使用数量'] || '1'), unit: String(r['単位'] || '個'),
          level: String(r['レベル'] || '1'), parentId: String(r['親ID'] || ''),
          type: String(r['種別'] || '部品'), note: String(r['備考'] || '')
        };
      }
    });

    _read(S.MACHINES, H.MACHINES).forEach(function(r) {
      if (r['機種コード']) machines[r['機種コード']] = {
        code:String(r['機種コード']), name:String(r['機種名']||''),
        type:String(r['種類']||''), brand:String(r['ブランド']||''),
        date:String(r['発売日']||''), m:String(r['M基板']||''),
        d:String(r['D基板']||''), de:String(r['DE基板']||''),
        e:String(r['E基板']||''), c:String(r['C基板']||''),
        s:String(r['S基板']||''), updatedAt:String(r['更新日時']||'')
      };
    });

    _read(S.BOARDS, H.BOARDS).forEach(function(r) {
      if (r['基板ID']) boards[r['基板ID']] = {
        id:String(r['基板ID']), name:String(r['基板名']||''),
        category:String(r['基板分類']||''), version:String(r['バージョン']||''),
        status:String(r['ステータス']||'設計中'),
        date:String(r['作成日']||''), note:String(r['備考']||''),
        updatedAt:String(r['更新日時']||'')
      };
    });

    _read(S.STOCK, H.STOCK).forEach(function(r) {
      if (r['部品コード']) stock[r['部品コード']] = {
        code:String(r['部品コード']), qty:String(r['在庫数']||''),
        safe:String(r['安全在庫']||''), note:String(r['備考']||''),
        updatedAt:String(r['更新日時']||'')
      };
    });

    var changelog = _read(S.CHANGELOG, H.CHANGELOG).map(function(r) {
      return {
        date:   r['日時']      ? new Date(r['日時']).toISOString()     : '',
        user:   String(r['担当者']  || ''),
        ecn:    String(r['ECN番号'] || ''),
        target: String(r['対象']    || ''),
        type:   String(r['変更種別'] || ''),
        before: String(r['変更前']  || ''),
        after:  String(r['変更後']  || ''),
        note:   String(r['備考']   || '')
      };
    });

    var sslinks = _read(S.SSLINKS, H.SSLINKS).map(function(r) {
      return { name:String(r['名前']||''), ssId:String(r['SSID']||''), sheet:String(r['シート名']||''), type:String(r['種別']||'inventory'), note:String(r['備考']||'') };
    });

    return { parts:parts, bom:bom, machines:machines, boards:boards, stock:stock, changelog:changelog, sslinks:sslinks, priceHistory:[] };
  });
}

// ── 部品マスタ保存 ──
function apiSaveParts(partsObj) {
  return _wrap(function() {
    var rows = Object.values(partsObj).map(function(p) {
      return { '部品コード':p.code, '部品名':p.name||'', 'メーカー名':p.maker||'', '公表単価':p.price||'', '仕入先':p.supplier||'', '廃番フラグ':p.eol||'0', 'RoHS対応':p.rohs||'0', '認定部品':p.certified||'0', '備考':p.note||'', '更新日時':p.updatedAt||'' };
    });
    _write(S.PARTS, H.PARTS, rows);
    return { saved: rows.length };
  });
}

// ── BOM保存 ──
function apiSaveBom(bomObj) {
  return _wrap(function() {
    var rows = Object.values(bomObj).map(function(r) {
      return { '基板ID':r.boardId, '部品コード':r.partCode, '使用数量':r.qty||'1', '単位':r.unit||'個', 'レベル':r.level||'1', '親ID':r.parentId||'', '種別':r.type||'部品', '備考':r.note||'' };
    });
    _write(S.BOM, H.BOM, rows);
    return { saved: rows.length };
  });
}

// ── 機種マスタ保存 ──
function apiSaveMachines(machinesObj) {
  return _wrap(function() {
    var rows = Object.values(machinesObj).map(function(m) {
      return { '機種コード':m.code, '機種名':m.name||'', '種類':m.type||'', 'ブランド':m.brand||'', '発売日':m.date||'', 'M基板':m.m||'', 'D基板':m.d||'', 'DE基板':m.de||'', 'E基板':m.e||'', 'C基板':m.c||'', 'S基板':m.s||'', '備考':'', '更新日時':m.updatedAt||'' };
    });
    _write(S.MACHINES, H.MACHINES, rows);
    return { saved: rows.length };
  });
}

// ── 基板マスタ保存 ──
function apiSaveBoards(boardsObj) {
  return _wrap(function() {
    var rows = Object.values(boardsObj).map(function(b) {
      return { '基板ID':b.id, '基板名':b.name||'', '基板分類':b.category||'', 'バージョン':b.version||'', 'ステータス':b.status||'設計中', '作成日':b.date||'', '備考':b.note||'', '更新日時':b.updatedAt||'' };
    });
    _write(S.BOARDS, H.BOARDS, rows);
    return { saved: rows.length };
  });
}

// ── 在庫保存 ──
function apiSaveStock(stockObj) {
  return _wrap(function() {
    var rows = Object.values(stockObj).map(function(s) {
      return { '部品コード':s.code, '在庫数':s.qty||'', '安全在庫':s.safe||'', '備考':s.note||'', '更新日時':s.updatedAt||'' };
    });
    _write(S.STOCK, H.STOCK, rows);
    return { saved: rows.length };
  });
}

// ── 変更履歴を1件追加 ──
function apiAppendChangelog(entry) {
  return _wrap(function() {
    var sh = _sheet(S.CHANGELOG, H.CHANGELOG);
    if (sh.getLastRow() === 0) sh.appendRow(H.CHANGELOG);
    sh.appendRow([
      entry.date || new Date(),
      entry.user  || '—',
      entry.ecn   || '',
      entry.target|| '',
      entry.type  || '',
      entry.before|| '',
      entry.after || '',
      entry.note  || ''
    ]);
    // 価格変更の場合は価格履歴にも記録
    if (entry.type === '単価変更' && entry.target) {
      var ph = _sheet(S.PRICE_H, H.PRICE_H);
      if (ph.getLastRow() === 0) ph.appendRow(H.PRICE_H);
      var oldP = parseFloat(entry.before || 0);
      var newP = parseFloat(entry.after  || 0);
      var pct  = oldP > 0 ? (((newP - oldP) / oldP) * 100).toFixed(1) + '%' : '—';
      ph.appendRow([new Date(), entry.target, entry.before, entry.after, pct]);
    }
    return {};
  });
}

// ── SS連携設定保存 ──
function apiSaveSSLinks(links) {
  return _wrap(function() {
    var rows = (links || []).map(function(l) {
      return { '名前':l.name||'', 'SSID':l.ssId||'', 'シート名':l.sheet||'', '種別':l.type||'inventory', '備考':l.note||'' };
    });
    _write(S.SSLINKS, H.SSLINKS, rows);
    return { saved: rows.length };
  });
}

// ── 在庫を外部SSから同期 ──
function apiSyncInventoryFromSS() {
  return _wrap(function() {
    var links = _read(S.SSLINKS, H.SSLINKS);
    var synced = 0;
    var errors = [];
    var stockMap = {};

    // 既存在庫を読み込み
    _read(S.STOCK, H.STOCK).forEach(function(r) {
      if (r['部品コード']) stockMap[r['部品コード']] = r;
    });

    links.forEach(function(link) {
      if (!link['SSID']) return;
      try {
        var extSS    = SpreadsheetApp.openById(String(link['SSID']));
        var shName   = link['シート名'] || extSS.getSheets()[0].getName();
        var extSheet = extSS.getSheetByName(shName);
        if (!extSheet) { errors.push(link['名前'] + ': シート「' + shName + '」が見つかりません'); return; }
        var vals = extSheet.getDataRange().getValues();
        if (vals.length < 2) return;
        var headers = vals[0].map(String);
        // 列を自動検出
        var codeIdx  = _findColIdx(headers, ['部品コード','品番','PartCode','part_code','コード']);
        var stockIdx = _findColIdx(headers, ['在庫数','在庫','Stock','stock_qty','現在庫']);
        var safeIdx  = _findColIdx(headers, ['安全在庫','安全在庫数','SafeStock','safe_stock']);
        if (codeIdx < 0) { errors.push(link['名前'] + ': 部品コード列が見つかりません'); return; }
        if (stockIdx < 0) { errors.push(link['名前'] + ': 在庫数列が見つかりません'); return; }
        vals.slice(1).forEach(function(row) {
          var code = String(row[codeIdx] || '').trim();
          if (!code) return;
          stockMap[code] = {
            '部品コード': code,
            '在庫数':     row[stockIdx] !== '' ? row[stockIdx] : (stockMap[code] ? stockMap[code]['在庫数'] : ''),
            '安全在庫':   safeIdx >= 0 ? row[safeIdx] : (stockMap[code] ? stockMap[code]['安全在庫'] : ''),
            '備考':       '外部SS: ' + link['名前'],
            '更新日時':   new Date().toISOString()
          };
          synced++;
        });
      } catch(e) {
        errors.push(link['名前'] + ': ' + e.message);
        Logger.log('SS Sync Error [' + link['名前'] + ']: ' + e.message);
      }
    });

    _write(S.STOCK, H.STOCK, Object.values(stockMap));
    return { synced: synced, errors: errors };
  });
}

function _findColIdx(headers, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    var idx = headers.indexOf(candidates[i]);
    if (idx >= 0) return idx;
  }
  return -1;
}

// ── Gemini AI 部品調査 ──
function apiAskAI(partName) {
  return _wrap(function() {
    var key = getSetting('GEMINI_API_KEY');
    if (!key) throw new Error('GEMINI_API_KEY が未設定です（スクリプトプロパティに追加してください）');
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + key;
    var res = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ contents: [{ parts: [{ text:
        'パチンコ・パチスロ基板用電子部品「' + partName + '」について以下を簡潔に回答してください：\n' +
        '1. 主な仕様・スペック\n2. 基板上での用途・役割\n3. 代替品・互換品の候補\n4. 廃番・入手性の注意点\nプレーンテキストで300字程度で回答してください。'
      }]}]})
    });
    var json = JSON.parse(res.getContentText());
    return { answer: json.candidates[0].content.parts[0].text };
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
