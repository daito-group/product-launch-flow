/**
 * 新商品導入フロー 進捗管理アプリ ― 共有用バックエンド（Google Apps Script）
 *
 * ▼設置手順は「共有版セットアップ手順.md」を参照。
 * ▼このスクリプトは、アプリのデータ（JSON）をこのスプレッドシートに保存/読み出しします。
 *   データは "DATA" シートのA列に分割保存されます。
 *
 * ★★ 必ず TOKEN を、あなただけが知る「合言葉」に書き換えてください ★★
 *    （英数字で20文字程度を推奨。例をそのまま使わないこと）
 */
var TOKEN = 'donyu-2026-CHANGE-THIS';   // ← ここを必ず変更

var SHEET_NAME = 'DATA';
var CHUNK = 40000;   // 1セルあたりの最大文字数（Sheetsの上限5万に対する余裕）

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); }
  return sh;
}

function readAll_() {
  var sh = sheet_();
  var version = Number(sh.getRange('B1').getValue()) || 0;
  var last = sh.getLastRow();
  if (last < 1) { return { version: version, data: '' }; }
  var vals = sh.getRange(1, 1, last, 1).getValues();
  var s = '';
  for (var i = 0; i < vals.length; i++) { s += (vals[i][0] == null ? '' : vals[i][0]); }
  return { version: version, data: s };
}

function writeAll_(dataStr, version) {
  var sh = sheet_();
  sh.clearContents();
  dataStr = String(dataStr || '');
  var rows = [];
  for (var i = 0; i < dataStr.length; i += CHUNK) { rows.push([dataStr.substr(i, CHUNK)]); }
  if (rows.length) { sh.getRange(1, 1, rows.length, 1).setValues(rows); }
  sh.getRange('B1').setValue(version);
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 読み出し（GETでも簡易確認できる）
function doGet(e) {
  var token = (e && e.parameter && e.parameter.token) || '';
  if (token !== TOKEN) { return out_({ ok: false, error: 'token' }); }
  var r = readAll_();
  return out_({ ok: true, version: r.version, data: r.data });
}

// 読み書き（アプリはこちらを使用）
function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'badjson' }); }
  if ((body.token || '') !== TOKEN) { return out_({ ok: false, error: 'token' }); }

  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (err) { return out_({ ok: false, error: 'busy' }); }
  try {
    if (body.action === 'load') {
      var r = readAll_();
      return out_({ ok: true, version: r.version, data: r.data });
    }
    if (body.action === 'save') {
      var cur = readAll_();
      // 楽観ロック：読み込み以降に他の人が保存していたら上書きしない
      if (typeof body.baseVersion === 'number' && body.baseVersion !== cur.version) {
        return out_({ ok: false, error: 'conflict', version: cur.version });
      }
      var nv = cur.version + 1;
      writeAll_(body.data, nv);
      return out_({ ok: true, version: nv });
    }
    return out_({ ok: false, error: 'action' });
  } finally {
    lock.releaseLock();
  }
}
