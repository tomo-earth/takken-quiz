/**
 * アース建設 現場日報アプリ — サーバー側（Google Apps Script）
 *
 * 台帳スプレッドシートに日報行を追記し、写真を共有ドライブの現場フォルダへ保存する。
 * 台帳のルール:
 *   - 自社は作業員1人＝1行（人工は原則1.0）、業者は1社＝1行（0.5刻み）
 *   - B列(曜日)・E列(工程)は台帳側の数式のため書き込まない
 *   - 書き込む列: A日付 C現場 D区分 F工種 G作業員/業者名 H作業内容 I人工 J機械 K備考
 *                 M開始時刻 N終了時刻 O入力者 P送信日時 Q写真
 */

// 台帳スプレッドシートの ID は「セットアップ」実行時に自動で保存されます（手入力不要）
var FOLDER_ID = '1dRBCpHN1mKeTslkhLNfoMzyIpUm4-cem';            // 07.日報フォルダ
var LEDGER_NAME = '現場日報_台帳（アプリ用）';

var SHEET_LEDGER = '台帳';
var SHEET_MASTER = 'マスター';
var SHEET_CONFIG = '設定';
var LEDGER_MAX_ROW = 2000;   // 数式・書式を用意してある最終行

function doGet() {
  return HtmlService.createHtmlOutput(APP_HTML)
    .setTitle('アース建設 現場日報')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

function ss_() { return SpreadsheetApp.openById(ledgerId_()); }

function ledgerId_() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) return id;
  var it = DriveApp.getFolderById(FOLDER_ID).getFilesByName(LEDGER_NAME);
  if (it.hasNext()) {
    id = it.next().getId();
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', id);
    return id;
  }
  throw new Error('台帳がまだ作られていません。Apps Script の画面で「セットアップ」を1回実行してください。');
}

function colValues_(sheet, col, fromRow) {
  var last = sheet.getLastRow();
  if (last < fromRow) return [];
  return sheet.getRange(fromRow, col, last - fromRow + 1, 1).getValues()
    .map(function (r) { return String(r[0]).trim(); })
    .filter(function (v) { return v !== ''; });
}

/** 画面の初期データ（マスター・今日の入力状況） */
function api_init() {
  var ss = ss_();
  var master = ss.getSheetByName(SHEET_MASTER);
  var kouShuNames = colValues_(master, 5, 5);          // E列 工種
  var kouShuParents = master.getRange(5, 6, kouShuNames.length || 1, 1)
    .getValues().map(function (r) { return String(r[0]).trim(); });
  var kouShu = kouShuNames.map(function (n, i) {
    return { name: n, parent: kouShuParents[i] || '' };
  });
  var email = Session.getActiveUser().getEmail();
  return {
    user: email,
    userName: guessUserName_(email),
    today: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd'),
    sites: colValues_(master, 1, 5),                   // A列 現場
    workers: colValues_(master, 8, 5),                 // H列 作業員(自社)
    gyousha: colValues_(master, 10, 5),                // J列 業者
    machines: colValues_(master, 12, 5),               // L列 機械
    contents: colValues_(master, 14, 5),               // N列 作業内容
    kouTei: colValues_(master, 3, 5),                  // C列 工程
    kouShu: kouShu,
    todayEntries: readEntries_(new Date(), 'all'),
    ledgerUrl: ss.getUrl()
  };
}

/** メールアドレス→作業員名の推定（履歴の初期表示用） */
function guessUserName_(email) {
  var local = String(email).split('@')[0].toLowerCase();
  var map = { tomoyose: '友寄', miyagi: '宮城', uehara: '上原',
              yamazato: '山里', taira: '平良', toyama: '當山', nikadori: '荷川取' };
  for (var key in map) { if (local.indexOf(key) >= 0) return map[key]; }
  return '';
}

/**
 * 日報の送信。
 * payload = { date:'yyyy-MM-dd', site, biko,
 *   works: [{kubun:'自社'|'業者', koushu, koushuParent(新規工種時のみ),
 *            workers:[..](自社) | gyousha, ninku(業者),
 *            content, machine, start, end, addToMaster:{koushu,gyousha,content,machine}}],
 *   photos: [{name, mime, dataB64}] }
 */
function api_submit(payload) {
  if (!payload || !payload.date || !payload.site) throw new Error('日付と現場は必須です。');
  if (!payload.works || !payload.works.length) throw new Error('作業が1件もありません。');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = ss_();
    var ledger = ss.getSheetByName(SHEET_LEDGER);
    var email = Session.getActiveUser().getEmail();
    var now = new Date();

    // 追記位置: A列(2行目〜)の最初の空きセル
    var aVals = ledger.getRange(2, 1, LEDGER_MAX_ROW - 1, 1).getValues();
    var row = 2;
    for (var i = aVals.length - 1; i >= 0; i--) {
      if (String(aVals[i][0]) !== '') { row = i + 3; break; }
    }
    var rows = expandRows_(payload);
    if (row + rows.length - 1 > LEDGER_MAX_ROW) {
      throw new Error('台帳が' + LEDGER_MAX_ROW + '行に達しました。友寄さんに「行の追加」を依頼してください。');
    }

    // 写真の保存（先に保存してリンクを行に書ける状態にする）
    var photoLinks = [];
    if (payload.photos && payload.photos.length) {
      photoLinks = savePhotos_(payload.site, payload.date, email, payload.photos);
    }

    // 行の書き込み（B・E列の数式は残す）
    for (var r = 0; r < rows.length; r++) {
      var w = rows[r];
      var tr = row + r;
      ledger.getRange(tr, 1).setValue(parseDate_(payload.date));                 // A 日付
      ledger.getRange(tr, 3, 1, 2).setValues([[payload.site, w.kubun]]);         // C,D
      ledger.getRange(tr, 6, 1, 6).setValues([[w.koushu, w.name, w.content,
        w.ninku === '' ? '' : w.ninku, w.machine, w.biko]]);                     // F..K
      ledger.getRange(tr, 13, 1, 5).setValues([[w.start, w.end, email,
        Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm'),
        r === 0 ? photoLinks.join('\n') : '']]);                                 // M..Q
    }

    addToMaster_(ss, payload.works);
    return { added: rows.length, photos: photoLinks.length,
             entries: readEntries_(parseDate_(payload.date), 'all') };
  } finally {
    lock.releaseLock();
  }
}

/** 1件の日報を台帳ルールで行に展開する（クライアントの確認画面と同じ規則） */
function expandRows_(payload) {
  var out = [];
  payload.works.forEach(function (w) {
    var base = { kubun: w.kubun, koushu: String(w.koushu || '').trim(),
                 content: String(w.content || '').trim(),
                 machine: String(w.machine || '').trim(),
                 start: String(w.start || ''), end: String(w.end || ''),
                 biko: String(payload.biko || '').trim() };
    if (!base.koushu) throw new Error('工種が未入力の作業があります。');
    if (w.kubun === '自社') {
      if (!w.workers || !w.workers.length) throw new Error('自社の作業に作業員が選ばれていません。');
      w.workers.forEach(function (name) {
        out.push(Object.assign({}, base, { name: name, ninku: (w.ninkuEach || 1) }));
      });
    } else {
      if (!w.gyousha) throw new Error('業者名が未入力の作業があります。');
      var n = Number(w.ninku);
      out.push(Object.assign({}, base, { name: String(w.gyousha).trim(),
        ninku: isNaN(n) || n <= 0 ? '' : n }));
    }
  });
  return out;
}

function parseDate_(s) {
  var p = String(s).split('-');
  return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
}

/** リストに無い値のマスター追記（アプリ側で「追加する」を選んだもののみ） */
function addToMaster_(ss, works) {
  var master = ss.getSheetByName(SHEET_MASTER);
  function appendCol(col, value) {
    if (!value) return;
    var vals = colValues_(master, col, 5);
    if (vals.indexOf(value) >= 0) return;
    master.getRange(5 + vals.length, col).setValue(value);
  }
  works.forEach(function (w) {
    var a = w.addToMaster || {};
    if (a.koushu && w.koushuParent) {           // 工種は親工程とセットで追記
      var names = colValues_(master, 5, 5);
      if (names.indexOf(String(w.koushu).trim()) < 0) {
        master.getRange(5 + names.length, 5, 1, 2)
          .setValues([[String(w.koushu).trim(), String(w.koushuParent).trim()]]);
      }
    }
    if (a.gyousha) appendCol(10, String(w.gyousha || '').trim());
    if (a.content) appendCol(14, String(w.content || '').trim());
    if (a.machine) appendCol(12, String(w.machine || '').trim());
  });
}

/** 写真を 現場フォルダ/06_写真/日報/yyyy-MM/ に保存し、URL の配列を返す */
function savePhotos_(site, dateStr, email, photos) {
  var folder = resolvePhotoFolder_(site, dateStr);
  var userName = guessUserName_(email) || String(email).split('@')[0];
  var links = [];
  photos.forEach(function (p, i) {
    var ext = (p.name && p.name.indexOf('.') >= 0)
      ? p.name.slice(p.name.lastIndexOf('.')) : '.jpg';
    var name = dateStr + '_' + userName + '_' +
      Utilities.formatString('%02d', i + 1) + ext;
    var blob = Utilities.newBlob(Utilities.base64Decode(p.dataB64),
      p.mime || 'image/jpeg', name);
    links.push(folder.createFile(blob).getUrl());
  });
  return links;
}

/** 設定シートの割り当て→名前検索→写真置き場 の順で保存先フォルダを決める */
function resolvePhotoFolder_(site, dateStr) {
  var config = ss_().getSheetByName(SHEET_CONFIG);
  var vals = config.getRange(1, 1, config.getLastRow(), 2).getValues();
  var conf = {};
  var siteFolderId = '';
  vals.forEach(function (r, i) {
    var k = String(r[0]).trim();
    if (k === '写真ルートフォルダID' || k === '写真置き場フォルダID' ||
        k === '写真サブフォルダ名' || k === '日報写真フォルダ名') conf[k] = String(r[1]).trim();
    if (i >= 9 && k && k === String(site).trim()) siteFolderId = String(r[1]).trim();
  });

  var target = null;
  try {
    if (siteFolderId) target = DriveApp.getFolderById(siteFolderId);
  } catch (e) { target = null; }
  if (!target && conf['写真ルートフォルダID']) {
    // 00.工事部 直下から現場名を含むフォルダを探す
    var root = DriveApp.getFolderById(conf['写真ルートフォルダID']);
    var it = root.getFolders();
    var key = String(site).replace(/\s/g, '');
    while (it.hasNext()) {
      var f = it.next();
      if (f.getName().replace(/\s/g, '').indexOf(key) >= 0) { target = f; break; }
    }
  }
  if (target) {
    target = subFolder_(target, conf['写真サブフォルダ名'] || '06_写真');
    target = subFolder_(target, conf['日報写真フォルダ名'] || '日報');
  } else {
    // 現場が特定できない → 写真置き場
    target = DriveApp.getFolderById(conf['写真置き場フォルダID']);
  }
  return subFolder_(target, String(dateStr).slice(0, 7));   // yyyy-MM
}

function subFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/** 履歴: 指定日の行を返す。scope='mine' は入力者=自分 or 作業員名=自分 */
function api_history(dateStr, scope) {
  return readEntries_(parseDate_(dateStr), scope || 'all');
}

function readEntries_(date, scope) {
  var ledger = ss_().getSheetByName(SHEET_LEDGER);
  var last = ledger.getLastRow();
  if (last < 2) return [];
  var vals = ledger.getRange(2, 1, last - 1, 17).getValues();
  var email = Session.getActiveUser().getEmail();
  var myName = guessUserName_(email);
  var key = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');
  var out = [];
  vals.forEach(function (r) {
    if (!(r[0] instanceof Date)) return;
    if (Utilities.formatDate(r[0], 'Asia/Tokyo', 'yyyy-MM-dd') !== key) return;
    var mine = String(r[14]) === email || (myName && String(r[6]).trim() === myName);
    if (scope === 'mine' && !mine) return;
    out.push({ site: String(r[2]), kubun: String(r[3]), kouTei: String(r[4]),
      koushu: String(r[5]), name: String(r[6]), content: String(r[7]),
      ninku: r[8] === '' ? '' : Number(r[8]), machine: String(r[9]),
      biko: String(r[10]), start: String(r[12]), end: String(r[13]),
      mine: mine, photos: String(r[16] || '') });
  });
  return out;
}

/* =====================================================================
 *  セットアップ（初回に1回だけ、Apps Script の画面から ▶ 実行する）
 *   1. 07.日報フォルダ内の台帳 Excel（.xlsm/.xlsx）を Google スプレッドシートに変換
 *   2. 台帳シートに M〜Q 列（開始時刻・終了時刻・入力者・送信日時・写真）の見出しを追加
 *   3. 「設定」シート（写真フォルダの割り当て）を作成
 *   4. 台帳の ID をこのスクリプトに記憶（以後の手入力は不要）
 * ===================================================================== */
function セットアップ() {
  var folder = DriveApp.getFolderById(FOLDER_ID);
  var ssId = null;
  var it = folder.getFilesByName(LEDGER_NAME);
  if (it.hasNext()) {
    ssId = it.next().getId();
    Logger.log('既存の台帳を使います: ' + LEDGER_NAME);
  } else {
    var src = pickSourceFile_(folder);
    if (!src) throw new Error('07.日報フォルダ に台帳の Excel ファイル（現場日報台帳）が見つかりません。');
    Logger.log('変換元: ' + src.getName());
    var res = UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' + src.getId() + '/copy?supportsAllDrives=true', {
        method: 'post', contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        payload: JSON.stringify({ name: LEDGER_NAME,
          mimeType: 'application/vnd.google-apps.spreadsheet', parents: [FOLDER_ID] }),
        muteHttpExceptions: true });
    if (res.getResponseCode() >= 300) throw new Error('スプレッドシートへの変換に失敗しました: ' + res.getContentText());
    ssId = JSON.parse(res.getContentText()).id;
  }
  var ss = SpreadsheetApp.openById(ssId);
  ensureLedgerColumns_(ss);
  ensureConfigSheet_(ss);
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ssId);
  var msg = '✅ セットアップ完了。台帳: ' + ss.getUrl() +
    '\n次は右上「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」で公開してください。';
  Logger.log(msg);
  return msg;
}

function pickSourceFile_(folder) {
  var best = null, bestScore = 0;
  var files = folder.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    var name = f.getName(), mt = f.getMimeType();
    var isExcel = /spreadsheetml|ms-excel/.test(mt) || /\.xls[xm]$/i.test(name);
    if (!isExcel) continue;
    var score = 1;
    if (name.indexOf('クラウド版') >= 0) score += 2;
    if (name.indexOf('日報') >= 0) score += 1;
    if (score > bestScore) { best = f; bestScore = score; }
  }
  return best;
}

function ensureLedgerColumns_(ss) {
  var sh = ss.getSheetByName(SHEET_LEDGER);
  if (!sh) throw new Error('「台帳」シートが見つかりません。');
  if (sh.getMaxColumns() < 17) sh.insertColumnsAfter(sh.getMaxColumns(), 17 - sh.getMaxColumns());
  if (String(sh.getRange('M1').getValue()).trim() === '') {
    sh.getRange('M1:Q1').setValues([['開始時刻', '終了時刻', '入力者', '送信日時', '写真']]);
    sh.getRange('L1').copyTo(sh.getRange('M1:Q1'), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  }
}

function ensureConfigSheet_(ss) {
  if (ss.getSheetByName(SHEET_CONFIG)) return;
  var sh = ss.insertSheet(SHEET_CONFIG);
  var rows = [
    ['現場日報アプリ 設定シート（この表の場所・見出しは変えないでください）', '', ''],
    ['', '', ''],
    ['項目', '値', '説明'],
    ['写真ルートフォルダID', '1kFKgVW8pRDY2O-SEPVaAfC-oZYbiP_fi', '共有ドライブ「アース建設03案件」内の 00.工事部 フォルダ'],
    ['写真置き場フォルダID', '18tJl4_plrS1A53U5VsdbODN5--_-ODs0', '現場が特定できない写真の保存先（99_写真置き場）'],
    ['写真サブフォルダ名', '06_写真', '各現場フォルダ内の写真フォルダの名前'],
    ['日報写真フォルダ名', '日報', '06_写真 の下に自動で作る日報写真用フォルダ（月別フォルダも自動作成）'],
    ['', '', ''],
    ['現場フォルダの割り当て（現場名 → フォルダID。行を足せば現場を増やせます）', '', ''],
    ['現場名', '現場フォルダID', '備考'],
    ['谷直美様住宅新築工事', '14NO0tzYYFYRNyNAAuvV8bpyaTBQ1rLKI', '01_工事No　谷直美様住宅新築工事']
  ];
  sh.getRange(1, 1, rows.length, 3).setValues(rows);
  sh.getRange('A3:C3').setFontWeight('bold'); sh.getRange('A10:C10').setFontWeight('bold');
  sh.setColumnWidth(1, 260); sh.setColumnWidth(2, 320); sh.setColumnWidth(3, 460);
}

/* ===== 画面（index.html をそのまま埋め込み） ===== */
var APP_HTML = `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<link rel="apple-touch-icon" href="data:image/png;base64,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">
<link rel="icon" type="image/png" href="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAMAAAADACAYAAABS3GwHAABIXUlEQVR4nO29eaAdRZn3/6nqPvtdc7NBFkISAoGERUBAQBGQZVABR3HDbdBx19eZ8R3HcZxXnVeH1/k5MzrjNo7rgCKyyioICMi+SiBsgRBICFnvfpburvr9UVXdfc69N9yb5Z6T5D7Juf093X26q6uf56nneeqpKlHsadNM0RTtpeQjml2EKZqi5pE0GycFaWmYwlN4T8cCHyEQgG56YfZGjK17g6bw5GJB3AIYEgLcC5rCk4ETTZTQFJ5MLIUAnd4/RZNKru7T72EKTw7WAkRxZnsd/7umYQrveoz9PoWbgwXOBBL2oJOMKTwpGJHCTOFJxfYjSrM6rGA4mWgV+dzbsHtDU3gysSjNatfWK5uiyaa696GJ38MU3vXY1r0ozXYtwBRN0d5HsQ9Qt53CU3hvwAJEaZ+pFmCK9l6SAMLaR247hafw3oCFEIjSvp1TLUDTKO0FT9Hkkqn7KR+gqVi0QBn2VmzqXrTNqW8BWidCu+dj7PdWKc/eiEe0AKMqpyk8hfcwHO8rze3UAoHGJEQ3aqYpvOuwsN+0PSKm8KRh9zaksG/EbM2hKTw5mFTdiyk8qRj3HtrmddX5AFM0RXsTTUWBpvDeiwWItvlTLcAU7b0k6782WyT3TCyEQAiJEAIpJVJ6SCnxPInneXaf/Qhhz2+d8u/J2I/TRJlqCHaUBMI4VkIghdEtWiuiSBGpCKUilNIordBK2cibZXYhYuaX0iPj+3ieh3OYtdYorUHrqTe100jgg0YIgdbCpElrc2AKjwebrZRGu0dKEYYhtVqNqFYFBPgeHW3ttLW1kc/nKeTyFIslOto7yGR8wiBkuFKmWq0QBAFBGDI4NMiWrVspDw6AVqA0+D75XI5sJovneRhZUChlX6UtzxQePxYCRNuC7imFMkESQuBJidYQhAHlShldC8gUCkyf1sP8ufM56IADOXjJUg5ffhhzZu9LW6mNUqmNtlKJfC4/6nWDMKBWq9HX38f6Da+wbv06nn9hNc+vWc3KZ57kiaeeYP2GV6iVh0FKcvkCuVwOKQRKKdNCTNH4yHbEiPYGAWh213TLYsf0QK1WpTxsmHBmzwyWHbyM1x97AiccczzLlh7CjJ4ZSNngXjWQTjFrvb0/Nm3p3cqKlSt4ZMWj3Pvgfdzz4L08v2Y1OgzJFQrkc3mEkMbU0ro16q1FsSPRtn+3FqmDk9EDujthz5o3tSBgeHgYtGbe3Hmc8vpTOOONp3Hc0ccyf848GimKonobH0ZsGyktFFrr+LtG40lvxO+29G7ljnvu4Lqbb+DWO29j1ernUGFIsVQim82hlEJrZa/RGvXZSlgAon3/bl2/q1WK10yc2PVDw8ME1QrTZ8zihGOO521nncNpJ53KrBmz4nO11ihrjDvNP16tPhFy93GC5XlefKyvv4+7H7iHK2/4LVffcDUvr32JbLFIMV+0voLajnrYc3GsnNoWTovnBdL20N78VwqJkJKh4SGCapXDDz2CD77r/bzlTWexcMFCHBntqhEyifhMNsWCJ8CTiTC88NIafnHpRVz0m4t58umV+JkMbW3tKKXRKqI+8rd3Y9G+cFqSBdcawtkU7Gz8weEhglqNw5cfzqc//Enefe47KeQLADZ8qePWoZUo3Qq5lmFgcIDLrrmCH/z8R9zzwN3k8nmKhSJhFEHaYW6B+m8KFiDaF01L1cReSBo83yMMQwb7+zn80CP41AWf4D3nvotCwTB+FEVN1fQTJSeoThDCMORnv/4F3/j2haxa9QxtHZ34vk8URU0uaRNpSgCs1vc8evt66Wjv4G8/9Td89iOfoZhi/FbU9uMl1yo4Qdi4eRPf+v6/8f2f/ZDe/l66OrtiU25vJdG+eG8TANP+eZ5HEAQMDQxw2smnc+GXvs7hyw4DIIzCUaMuuytprYmUwreC8KeVj/GVb/4Tl193BW2lEn4mSxSFtmbSfw3tqRhAml3mRTd7lP7kYMj4GfoHBigVS/zHhd/hxl9dy+HLDiOMQrTW+J6/xzA/mOf3Pc8IQhRx6NLlXPbjS/jehf9BJpulf6CfjJ8xbG8f22CxR2OBQLQfME0nnvEeTBqENI7u1s2bOP641/Pjb/2QJYsOiGPur9Z5tadQOmT72MoVfPKLn+WOu26js6sH0CilQbROpGbXYZCCZHRSWgz2JAwazzc9pFt7t/KJCz7BTZdcx5JFBxCGYZylubeQyzwNo5DlS5dxy6U38sXPfYmh8hBhFOH5HqT4Yk/FAhAdS3o0pORCYxYO2FMw4Pse5XKZXDbHN7/8z3zkvRcARhM2g/H1CEAssZPdFqfr4FdX/Zq//PzHqQUBxUKBKIxMEW19ar0HYawAtC/p2UNTIUxvn7H3+5mzz7786rv/w7FHHjMp0R2NqfD0QGwBjOeWSrvncKkUu1YwnG/g+z533Hsn53/qg6xdv47Ojk6CMIjv3QpGy87Gov3AHp2Mkk9mh9i9sQCt8TMZevt6WbJwMVf++DIOXLSEIArJeD67gpRleIGwsw2MJK2hGmmqkaIcaiKlyXqCjCfISkExM3qL5IRCMj4h2h4KoxDf83l29Sre+dH38NDjj9Ld2U0YBlZZaMslpn5NS7t74jgVouPA6frVKmb3I43vZ9jau5XDDjmUK3/8G/abM58oiuryZ3bKnTQoNF4DVw4FijX9NdYOhLzQX+P53hobyxGV0DB+OVQMBppQafK+IOdJ8r6gMyeZUfDZty3Dfh0ZlkzLsaAzS8EXI+4pRdq23Tnk6mjz1i2882Pv4fd33sq07mkEYbhLW6HJJyvKHQftYQKgwc/49Pb1ctyRx3L5j37NzJ4ZhFEUx8F3Bilba2lN/+zWKo9sqPDg+jKPb6qwsRxRDhQK8IRpFaQAKQ32bAhCa1B2pFeoNEqBHedCKSOYUfRZ3JXlqNkFXjenyLz2THzPyMbwxmpxtoecEPQP9HPW+8/mjw/cTXdXN2EQTr6TsqvI+QN7mgD4vmH+1x7xWq77+VV0d3YRKYW3k5zdRsZfPxRy8+oBbnx+gFW9NQYDjRSQ8wRZT1omh8TLTXss9ZRuRNyvIq0JIkU10igN3TmPQ2fkecO8IifOLdFT8EYt146Sq7ONmzdy5vveysOPP0JXRxdhGO6cG7QIGQHYQ7xgz/MYGh5iv7nz+f0lNzJvn7k7zezRzga3DLZiY4VLnuzljy8Ns7EcWhNG4nvCnq8tR5sCinTnizDHRT3HmzPtz3TqfbizhIBAacqBIlKa2aUMb1pQ4s2L2lnclQWMIOwsp9lFiF5av5Yzzn8zT616mvZSe30OUQu89+3GAkTH0um6/ujuSVJ61GpVCoUCN//yBg4/+NCdxvyRTmz8xzZW+J/Ht3LrmiEqoaYtK8l6Eu3cKgGQityIxuiPFQBSjB3/Td5B49tI0nW0vZ6gGiqGA0V7VnL8nCLvP6SLJd2JIOyM1sDV4XNrnuOUd5/J+g3ryecKKOWEoNlcvGNYdCydrs2geDc43h7YHbBjNoxGrQU1Lv/hrznzjafHEY0dIcd0QsCG4ZDvP7yZ3z7bT01Be9bDl9ZWj5lcxIwuRhOG+Nx6ZhciiUokgqHjVsA9qimPDaymyhYpzUBN0ZYRvHlRO+87uJOZRb+u/DtCri7vuO+PnP6+N5PxM3GdN50HdgALIRAdB8/YrVW/wJg+W3u38oMLv8tH3v0XhGGI7+8Y86c16JXP9PH9h7ewdjCgM+fhSZEEXoVlfJwA1O93ZXTCgD2PlAmk0clxG6ojnrLYYIRO9SvY83TSXyAxDvRATTGz6PGBQ7o478AOe3THTSJXp//5s+/zqS99hmndPYRRGAvt7kqi8+AZ2tYvrkfM1n/LYzBO75atm/nAO97PT/+//9opzO9MnpcHA75xz0Zue3GQgi8p+NJEXYSZwwfhMMSa3oUmnVA4HLcM9u8E7HSn+d3fuCXQ7rgTFBMarYaKoVBxyvwSnzq8m33bMjvsF2hsZ5nn84G/uoCf/+YXTOueThSFSQvVQrzxatjxu+g4ZMZu2xPsSUm5Umb+nPncfcUf6OrsjjXx9pBjLingrrVDfPWPr7BuMKKr4PyIhMllrOlFqgVgRAvgSiKcwGhimyQ+KkZ7SuodYrtDx2dYAbCtgsPKtgaegOf7arxxXol/feMssl7SGm0vGUHTDAwOcNJ5b2Lls09RLBSJVBRft1V4YzxYAKLzkBmpKFCqtlsd2wcoV8pc+9MrOeX4k3cotyddOT9+bAvffWgzUgqKviQCJG7KwhTTCxAj9ifXEc7eFykTSLie6rSgatsiCGv3J82xbnACErOH2KZ1AuCEN4gUw6HmzxaU+OihXXGodEeY31GkIjzpcdeD93DKe84gn8vZe+vYtDM3a21sygtSCwEI3Mh4LdgtsO959Pb38jcf/RynHH9ynN+zPRT70xq++sdX+Jf7NpL3TVhTYacslAIpk63pzEph910IPGmySz0pU8cM9oTEEwLPE3gSPAm+FGafEPgSfLv1pN165pjL4vTsNePr22tmPcFwoCllPb50TA9/f0wP0wteYortBPKkRxRFvO7IY/n4+z5Kb99WfN+Lb9AKvDEeHAcoOpfN0HVxORd6aGEspaRcGebAhQdy9xV/IJfN1TmgEyH32EGk+Yc71nPNc/30FDK4SUSkkHVaX6ZbAEDIUVoCW8PxObjfpJLbUo9l3sgoBbObxO4ntvdN77E5Q2D8lsGa5uhZWT59WBfz2v2d2ieQJmXnGhoYHOCYs09kzboXyefyZn+r2DfjwQJkg3dWX2MtioUUVIOAL3/2ixTyhZGdSuMkZzdXQsVf3bKOa1cNGOa35onLGBWNGt8xd6z1Zaz9TWshY41uNDxkUpo9I8H3ICMFvhRkPMgIYb4LQUaY/X78W2F+Y1sLX9rr2t7mmgIhJB87tIMLT5jOvHafyJpDO5v5wSgFrTSd7Z18/W+/RjWoIVzILN3ctDoGROfymbtVT7DJUenjtNe/iWv++wqU1tud5uA05Bdue5lrVg0wo2gZRzptn2j1NE6OWc1vX75MnSsBVyxpz3MMWadz7B8R9xw7SkV5SFqAdKq01tBX0yzp9Pno8naWdmdtNKjhUruInD/wzk+fz2+uu5yuji4iFe02vIQA0XnoTLe79UkAWhBGAX/41c0cueyI+CVMlFyo81v3beTHj21lumP+lGavY/RtYCcYnjuGTXpLmUTGdDKC4EyfWBgcw7hHTDGQM4FMhMeYQFIIKqEi1PBn+xV4z5I2cp4g0qYMk0Wu7h97agXHv/2NxgdLO5u7Abl5/KjbtiQGX/r0D/Tx/redv4PMb7Izf7Wyl5+s2Mq0Qkrzp5zZWKM7Bzjl9Apr6kgh8IR1ZJ0JkzZTZGLqpE2gjISMZ/d7wpo5ziSy53oi9XuzzfmCSqSZVfT4wpGdfGhpOzlPoCaZ+cE6xCpi+YHLeNsZ59A/2I/v3kdL8MyrYCEQnYfNSvcrUaeOWggLAUppfM/j3svvYNF+C9megeyuh3fl5gofvO6lOAIjUmaPtGnKhvnlmCaQFDbNWWK1vmF6M2hFxFo/dpxJWgNIm0KNnKuTv7YVEEIQKE0l1Bw7K8d7lpTozMpd5uiOl1zo+U9PruDE896IkF7cqqWGz7QkFoj6JZJEYpS2FBbCaJuBoQHOOvlMFi9YtF3M71rmcqj42l0bCBX4ttlOmJ8U84vYlEk+JgwZhy8lNnwp4lYgDl867S0EvnCa3x232l7K+lYh5ei6ViDrGeZv8wUfXtrGx5e1x8y/qxzd8ZKUEqUUhx60jLedeS79g/140rMinLLrWhT7QqS1P/W12UJYaU02m+Uv32UGtG/PbGYawzDfe3gzj22s0FM0EZ/YpBEgSDR+vR/gTCBjargBLp4U1vY3++NBL7FfkPQSp51gs7WCXucAJGUVGLt/OFQsm5blvEVFZhW9WOvvzEEwO0LGT9H8rw9+istuuNL0DIsGX1i0HtaANM1rSwjjmFh6HoPDA5x49PEcf+RxaK0nnObstOWD64f5n8d76c77dl9K+8c2v6yz+YX9OBvfT2txUa/R/TrtngpleqOdR921MqljOU/YkWRw7v4lPrW8PWb+Zmv9RvKkCYsuP3AZrzvyWAaGB5H2/Tj+akUshOnhNwdcjYrWwk44lVJ8+LwPIYSIJ3YaL7m2ItLw3Ye3JIxtNXfykUnHVkrrC+voerZ3Nm3ieLZF8EV9z65fhxMTyPdGColxiI2gZD3IeoKq0sxv8/jEIe28aW4+fo5W0fqOlFKEkRklJqXkY+/+SFzhLv/JfKH1MOC75tdloifJrS2ChaBcq7Jw/kLOfMNpANvl+HoCrn62nwfWl+ku+LFjmTZx0r28sdmDY36bqiCInV+3X1rhkELgQeJLCPBIBE0Kl1Nkk6ljx9gwihQQRCah7Q375DltXp6sFDttcMvOIjfprusslJhJtm67+3Yuu/EKCsViknvfwF+thv1kCrwG0WgRLD1JeWCYNx73BjraOiY8ystpzYGa4kePbiHvy1jUx7Tz05rfObwi0fYu8uN6emN/wEWERLIv7gSzs7C670KkmEOYyFE50kzLS06bW+CgLpPSrVuI+d1M0p7nxe/gmdXPcvnvruLyG6/kkScfI4pCOtraUwOW0vzVahh8rPap91haBFvoeT5vPeUstoec9r92VT+r+wJ6SjZHRhI30YnJQ5354xjYaP2EyWNTyLUGsWCknWDTE+yR0v62qtM5Q1IIlNYECg7ryXLqnBztmVSEp8nMn55i3bW8/YP9/O7O33PJdZdy2713sGnrJvLZPG3FEkIIM2Y4fo+idbEAf4Tib77Sj7EQgmq1wsJ5C3j9UScAEzd/pIBapLnimX7yGWlz7FNx/pTzm+7hjRnbmTZyJPOPaAFkWvOnpkGpaw3cfUyPbi2Cgi84fXaOw3p27nje7SWzKLdJB/SkjLX9A489xKU3XM5vb7mWp1c/C0BbscT07ul2mVZV34XTQrw0AlvaNVOk7STypEe5Wubk151Ee1v7hHt+XWrArWsGeWpLla68j4v5131GMYdihhfpWH+K+ccUAtdSuHl/6oXAhUvRgpqC/Ts8Ttk3z7ScjBu+ZjF/emUZT5h6fnnDy1xz2/Vcev3l3Pvo/fQPDVDMF+lq74x/45zg3YpsHfumRUgPyk6cl+ZiuxWCU4872ZQ2ZRqNhxwjXflMv2FGa0+4XB2jjc0Ok9pAyvRJGNmlO8SOb/wxQiGlqDOB0s5xfQsg8IQm0CbSc9zsLEdNzyJontZ32t48t7HZgjDg9vvv5NfXX8aNd9zEiy+/hO9nKBWKTO/uQSlFpM2sEM6Bd7h1+OfVsUAYH8Dtdlug6VgICKOIaV3TOPTAZWbvBAxil1y5pj/g8U1Vir41f5waTmt/klhF2oE1TCxiG9/Y9fWaXzoHOT5mfYN0VMi2EEJATcG+RY83zM4xq5DEyieb+escWqvtn3lhFVfedDWX33QVj678E9WgRqlQpLtrGqCJIkUYRXFdubLvnthsEx+gxUhISa1aZsn+i9l/7gJgYva/QuMhuHPtEH01RU/BRzGKM+p6gWUqv0emtL1MtQKNZo+N76f3S1nfUjiBcT0Xr52R47XTM/jSdHRJRjVNdzm5uhwYGuCmu27hkusu47b7bmfjlo3ksjmKhSKlUokoUibFmaSgE2yIW5TMw5gpVwWx8xJHr5qMpZBUgyqHHXSoWdFwgva/FKZFu+PFITKeSGlZgdH0DR1h9riJ2yfmS70JlI70JH0CddEh1zIIJyCCmtL05CUnzsoxr5TS+tvz3naQtNaEUcjjz6zk0hsu4+pbruWp559BaUVbsURP9zS7nrAiSvu0LcQbOw2LUaJAaSujqdjOlHD08iMBJpT74x5u3WDAyi1VCr5jtYTpEcYWjP+JJBlOkkR/Yvt9RAuRMHnC/EnkyLdSFaE5uMvn+Fk58p6Iy9YMre/6UH55za+54EsfR2vIZbN0tpv5gyKlkmkPW40fdjJ25BsJT+bYd2zWTAwm3aFQKPCaQw4HTIswXlKY+PtTW6oM1jTtOWNoNEZ+0i1AOlTp1TnDKc2f6ghzHWQjQqNSkBEQamj3JcfNzLKk03ZqMfpLmCxyPtTBi5dSyBfI53KG6ZVqGcd0srCwCtZ3L0SkTqPZWBgHuKujizkz9zV7J8Q5Rqyf2FwlVGa2Zsd8cYNnBcE5wyZNwQ1ndFo/FcUh1cMr6wXDOc0Zm/5Q07Cw3efEWTk6s+knay5JaUKtrznkCA5fehj3PXY/bcU2tEvNsLQ3YLdHplgitW0yFoIwDJnVMyOON0+E3KRVj2+qknHDpBLOT83kYD9i5FaCFYTEBPJIBMX5Bq5VyEiItDGHTpyV483z8jHzp5+y2RRFEVIITjjqeGpBkFIsLfDeJx23aBRICkkQhczsmUkhX0BpNW4TyDHcYE3x0kBgZ25Oks7qmFEk+7bVe5tObXDhTTdYxo0oqynN3JLHSftkmZn36srSinTiYUfyTc/WaasWchKoNXuCBSgVMW/2HIBUYtWrk3MyN5dDBgOFJxOJr3OAHfOnGFyMsU2bQOlhj77N1BRojp6R5biZOXxBUx3dbZHS4Fum31zcn2KpC6V2/wlud4TsoHj7rRXcc+ugamD+vvPNvu14O5srEQM1hQsAJbdItQOxyWP9AIvT05vELYQUda2FL81id105wdn7FThxlmX+hsdpBdK40DJsrih+sHKIu8sz6J6+L2EYQLp1bREemBQswE++6HhncmLzsEbT0dbBRMnJyivDodXOJKaOSNn+jdsGLN3HtQYkjO/yeA7vyXDSPnmKfuvZ+o7iFAsBD22qce2aMr2ViK5SiULHTLase3qkxLYID+x6LPDBLXJgmMOE25uL3Qsp5gsAcfhqItRfNX2vQti5/EdTACJxchtDocmEWIkp5EuoKujKCk6Zk+fQaRlbvtZjfEi0/lCo+d1LFR7aVEUiKPqgpEexoweNit8/YN/F3oGFawEs79mtOampGMP0xUKR7aVKOFJojCJ0f0XD/sbzkhbB5fGUQ83B3VnOnG+zNy3ntxrzx62RgOcGQm5aW2HDcETRFwSRmRTM9zMUO2aidVRXAY4X9gasSadCkBwUTcag8TyffDaHKd5EWMxcaThMxg3HDmkDlwsS80CM9sF0bAXKTL945vw8J8w2ZWp2zv5Y5OoxVHD3hioPbKoRKShmBLXQnGGeS5LNF60Qm4d1gwM1ez52FeVrUfe9rhKbhTVmAqx8Ns/2UtW2AO7ZcN9ESuGJeO+Ij4v4DIWKRR0+5+5fYl6bF1+r1Zg/7YNsrCjueKXKS0MROSkI0dSU6f1MjzLTWmH6RSTSjF0DIElIx343/lDaFI3XMYhvbn4h0HHlCtzK7PY9pFrLZmO3tS2AiCdnFalFBJqFjVmmUUxs9oc0eTK5lrlsqg9gBOPbv8JgTwpCrdEKTpmT58z5BTKy9QanO0oz64regIc2B1RCTdEX1KKEmescfowAqKhGUB0iCipoOy7AiYHDNWVSO9yYATdrtpQSM318EjnDNqfut/HoslguWoTHLL/72j2w9QXcQzQLg9G8YRQyODxkX3CieV6dzDXas15Kg40e5xap5sAxh5QwFCjmtPmct6jIIdPMMMVWGpyeJsfcg6Hmgc011gyaSamynplhAlJ6wLxzpPQIqoMc8rp3MXvxsQRBjWplmGq1TK0yRLU6RFAdJqgOUS4PMSdXpUtWGa6UqdaqlKsVKrUKlWrF4GqFIKgRKY1SkRkeqUwadcbPUMwX0fGKC27TXOxaQX/EcZ2YCc3AriAa6O3vTXZMkDqyydIHxu4Vdc8Zz8pmNZMUZghlGGhO3DfPOxYVac+2zuD0RkqbPGuGIv60NWAw1OQ9o/WNBndrCifngnl2rUK6ZuxH+6wDCBWEWhAqQaAUQaQIlckO3VKO+Mxh7Zyxb3LvWlCjGtSoBTVqNbOt1CoMDg8xVB5mqDxE/9AAlWqFm++5lctuvpJSoWTmcxKQTPrURGxrY2QqhGgBjHlJWwd6Rzs0LmrLytTvLEo5Q2jsUEiTyzMcaHoKkvcsaeP43cTRrSl4ojfghaEIrSFrHXan7cFpfrOjXogFYVgjjCoEShMoTRhBqDWhxZHWRFWFF0giVQBrHmYzWbKZ7LjKWq1Vuei6S2gvtccry7QKj0GL5gJpYRLWeq0ATKSM7tRpBS9t6tUddArAE4bJBwLFa2fm+dDSNmYVvTi82WrMn9b6m6qKJ/tC+mqajDCObpSy88RoP1T1x81IOA+BRmqN9EBE2rQcEoTWeJ5gdlsGT0q7JBMjxmZodCxw2v41a7Z5rN34MtKT9U1QC5HJBUpzSqOj0AyMawH6tvvBZhV9OrMegdb4aU620BMwHGraMoL3L2nnnIWmz6HVtb7SsHowYvVQRKRNFmqQWprrVS+Qah5GPKa7iD0QKWjLSLpz9SPxGlPT0+Zlmjzp8dIra43JlfzY3Ei3AEaYjjDTSiZbVwfNwggz9fmGLRuSCh4nuXczq+jTXfBYNxiSSb0/OzSG3qpi2fQsn1zewZKuTCx/rcj8YOplMNQ8NxCxtarxrM6on5BEj4JSFzBLzMTHRz3HnicwJlBXVjItb1PGJlA3Lnv3pQ1r8TwTRI15zQpMc7HZ1plA6QppJtYaMpkML21YR7lappAb/0J47uGynmCfks+a/jB+RE9ANQIhNO88oMT7l7aT93aP8Ob6suLFoYiaMikZRusn6wabejNPH/s42vG7jpWfO0SyO95qndonTGfaPkW/LtdpXOW272q4Msym3s1k/Gx8zdYhN07cSb2wTVsLYI1ZC+CljevYsGXThB/N2ar7d2TjEWFSQF9VMavo8bXjpvGXyzpamvnBVElVaZ4diFgzpNAIG61KlkqNNbn9E+8Tuo7ht8X8YzFmoDSLu2y+06vaWCOpd6CfDVs34md8dJxsKZL7NRkLkfIB4nVrnag2EYOxHweHB3l+3Wr222eeWQ1ygrHIQ6ZnydjVVSqR5owFbXz68C46c8nSQq3G/Gl+3FLVrC1HVCNTzlBps0qkTlaL1Npo3LgViPc1fCel7Rtag7Tmj/drox0Xd058yIhrAV7evJ7+oQEyGSNE8Vyszg5vInZhd+lsb4Rr+kTzMWb8aiWo8vQLz9pKHX+vsMtuXzY9j8SYQ39/TA9fOmZazPwuxbmVyDF/pGHNkGL1YBR3ZoXaMH+kTZeSxgiBwnxXWifC4Jhck4rcjGR8ZZFpMcxJThhCrenMCQ7sdsw7/udQtrl4fNVKyrVKMpt3i2j+9PP46QdLP2OzsTOFnnlxVfx9vOROnVn0Of/gTs5Y0MaCzkzLa30B9AWal4YU5cjE7ZU2iwMa5rdMjbbMX6/djSCkfAOtbSpCIhhKW2GxN1a6/hpgClINNct7sswseBOy/93vAf707AqUTbduOW1jyZpAqT3pp20iVlqTy+Z49JnHMGvjTnwaKSngY4d1A7tHeHPtsGJ9RdmmWhMpM3W60/AKiACliJnaMX38IbUlZQYRu8fW7En7CCbvKX1OoODwGcls1RNZgtUtXP7wU4+SzWTjFqEV+KpREFMjwlIntADWaAq5PH9a9TjPr3uBhXMWTGhwvCM3KqzVmD/9yIOh5vmBiIFA48apq1SUx2l/BUTKdHhF2H1WWaS32mn5lIZPC0l6f9r+d/siBaWM4OiZpkd8IjXu3tHql9fw5OqnyefyuEmOY2q2eeGwxq4RJlI7WwRroclkMmzu38Jdj90LgHYG7QSolfN4AF4ainhkS0hvzUhqpLXJw9HG6Q21ju1/t19httEozG/MJd3A9DrxGZyQ0GhK6TjSMxwqDurOML/dr+9JH8+z2Xd03+MPsKlvM9lMxgoArfXBbCWOQVLblsBJGbntoTsMbjVO3g5yzD8cah7aEvBkf0SkNBptc3JM6nHoNL22QqG1yc2xH2MS6VQrkJhGUR3jJ4JRZw6lbX8nBPYTaXjT/EJc3u2hOx65C+V+3WxeGgObMKhwJdSpLS2BldYU8gXue/xBhitlivnxd4i1GqVNnpeGIp7sjwiUJmPHHiRLZJqzXRZ/rM2hjtkjjKniBCTe4oRAJ0IBiR8RtwIjzSYwwz4XtPscM8uaPxOoajfderVW5e4V95PP5VHoOKBhHsm+31bACGRq+IPdtg7WGEd49foXePipR602216d1DxyzF+NNPdvDnhgc0AlMqnKxsxJtLz7RMp0RIXaLKhhWgTD6GFDSxDvUynmVvUmUPJdp/Yn/oKxXEx/yZvmFeLVKSf0nMro/LtX3MeTLzxFMZ8nWS0yacFbATuS2qqlZCtaCktPUqlVuOwPV+EySXYXSrdp68oRN71c5bmB0IY3LWNb5g2cNrfbIBaK9DbR7FHsJ2AHophruv2uFYh0cqyR8eMQqS1nJdTsW/J449x8oign8rzCPOsVf7iGahiYNQhaiJfS2PUJxKkQ6W0rYaUVxWKJa+/6HVv6t+JJb0Q6biuSiz6FGu7fVOOWl6smyiNEovWdk6tSefja+AKxM6xc5Cd1rrItgjsWn2OZXiVmkfMJTPg0MXtiJ5qkV7kcas5bXKLNLSY4gefVWuNJj4HhAW5+4FaKhSKRVrEAtAIvNWJEPDOcfdS0yLcI1kDemkE33HOziV5McKX4ySQnmlLAhkrE1WvKPLIlMOanZW6n1QOrwY35YvYHukEo7HmNghApXSdIkd1Xx+BxB1rSkZa2/bU9BjBUUyzryXLSnMJ29Zm4wS7X3X0Tq9Y+TyGXS9rqFuGlEVgI5Ajmj8WkRbAlKSW/vvUKBBNbKmkySaW05oObaly+uswr5QhfmHykwGn9lH3vNHZAKvITM39i4jgzKXAtQKzt7Xn2E7ljaeZ3ZhLU+QaR7QBzx96zpDShDq80SSFRSvHf1/wCz/NGMn8r8FIjpkVHhDWS0or2Yht3PHoXf1r1OMsXHYJSqmUEwfqPSAFbq4qb11ZYPRhR8M0qkrVIm1mlI40S6ekWTdTGNcemv0jEJp4LUyaJbzpOh4ijPHoU82cUcyh9XtpRRkBfRXHOwhKHTMtul/Z3y1fd/sgfuevxe2krtcXjfwX1vlArYWiYHLfZLdJYWAuzWvxAeYjvXPaDlpJZp/WFgEc31/jJU0M80xfa9QK0mYlN6dhWd/a7axHc9zCyZk+k7XG7P97WfxyjO78h8QUSU6cuRDoiZGrKP1hTLOnK8O4lpQl3ejly0bsf/vanhGFoHEv37swJrYdtY2A7wkyBtSt4q2EEkY7obG/nitt/y4rnVyKlbLov4LTlYKC59LlhLn1umOFQ2QErrkc3MWUC5+g6Iahj9kQwYqzNdRKzKWUe2d/FguXMJd3gQGvj/MadZynzqBaZ4aIfX95BzpvYoJe4DmxL/MTqJ7nxvlvMguY6lQCX4q+Wwhj+qvMBhN22HjaS60ufgfIg3770e+YBmhQSdfFxKWDl1oDvrOjnvleqZKQpZ01poijR0IHT1FHyPYioOxY0fOrPT1qBZPaGxJ8IXQuT6kNw/kEUpXqUVcL8WsNQqPjIIe3s3+Fvd7Kgewf/eul3GagMjLT/W8HWHwMLQMx468LWjynGJMB2hv3h29dxyIKDiJSKsw8ngxyjBEpzzeoyt62r4EtBwTOCatYQ03Zp1WSGabeOGKR9APcy3Bpd5vkANAK0WbrQZW4aDW72R2Ac2Aa/wNn9yrYIKk6jMA5waKV3c0XxzsVF3r2kbbuZ39n+9zx+P6d//m3ks7ndIkQNxOaeNDaRmyFaxPkgrYlNV/tgZYgLL/439yg7vXJGI1cOKeC5/pALH+7j2jXD8ZSANRfT187W14lmTmnvdBQo0fQ02PzGFwhSZpEzkyKlrSllTCqn3Z15Fdv6sfmV+AahNoK2qRxx0pz8DjE/GAHWWvO1X3yTIAoRQrYIn7w6do5bskpkvE20UCviSCm62jr5ze1Xcd4bz+XNx50+4UW0J0ppJrlm9TBXPj9MpKEtIwiVQmu7fjDCjMW12l4LM3tDJASenZg2QiCdxnczt7kbpRlRJ0KHNpmbrucWbUKYdb27KskDUqN1ill3aUtFccysHJ9Y1m5s/u1kflfnv7n9am556A90tXURqWinBD0mB9s2d8Y5i3aPNitFUkoq1QrzZ83jjm9fT3uxDRDx6pA7izTEc4K+PBTx3ysHeGRjjfastEuiCrt+cLKSfLKipG1ibbaVMYfMFI1pUwiEEQTHjS4Eas0hZwJBYgLFGZ66Ppkt7v2Nw51J6BMM85+wT47PHNZBRm6f0wtYk0xTqVU54dOns2rdagq5fNODEttDyXiAtCpqcay0olQo8OSLT/ONi7+FFHJCY4bHSwLDpDe/WObzd23h/leqFDPChDedyVLXi2ujMVGCneNZF91RSf5PowmUToeud35TnWIuwpQymeKok0ruG6Y6vTZXIk6Zm+dzh3ca5tfbx/yQDHr554v/lcdXP0mpUEjm/XQV1+rYfsSMty2yVWH1QVwzrY3d89SCGjdceBnHLD1qpzrEGthSifjpygFufKFCISPJ2yn0feHWBxax1ndObryMKgZ7CBD1jm/KBE3eSaL849fhWiDs1phBWIe4ftSXsseTjFDzu+HQCOjbF5d4+6JS/Gzbzfw27HnLI7dzzpfeSyGXj932VuGNcWFbXpmuivjYboA1IKQk0orP/OcX6BvqJ92Lur3k7O5Aaf75/q388qkB2rMCiTbhzYaOqLSD6UZxxekJsWYexfGN+wHs/ijVQjgHOB02bbhOfSeZdYLtNZxAbK0o2jKCv31NJ29fVIqfbXuZ343F2DKwlc/+xxeMeSftrG+O0hdvZWy/pkaEuVRRsZtg0xS3Fdt4ZNVjfO67X0QKQaTVDgmBq6OMFPzVa7p56/4lNpdDapEZrJ44lynG0wmjughQY7Jao9kSxEKi4k4uF+1x10ibQ42mTpIjZIXP9g2AphxqequKI2dm+dox3RwxI0dkld/2Mj9gnVzB//7hl3lm7SqKhZJJgrPvRaTe0e6AtWjIBXKadXfCoQrp6ZzGRbdcyhEHHMqnz/lLwijC97Y/KuSqZJ+Sx5ePmcby6YNc9NQgWyoR7VmJtMsOedb51NYc0sJEfqRIliNyC9W4yI+InWFwU8YKMVJgdfzHtUo6mdoEUvP+6HjS50hBf00xoyD5i6XtnDzXLDGl9MRmdRiNwijE93x+dP3Puej3lzKto5tImbENu10UxZIAxMy3L9bmVZgIhevZ2x2wcI9hX261VuWqr17MGw49foeFAKz9bS+/YTjil08P8Ls1ZarKzJqcseEcZ/sLEp9AiGTJVbcAd2z7k7S6wtW+ML8HNxhS2NFvAjc7BDQIgC1nqMyqNr4UvGHfPO88oER3TtYpjB0hV5c33H8z7/r6BeT8LEJIlFYj3sXuhQVixjsO0DtaQc0mjZmLphpU6eno4fp/+jWL5yzcaf0D6X6AlVtqXPLMIA9uqDEcakoZSd5z62MlYdBE8xsWidcdxjK6DYnG3RyNL0EDQqO1E3wrEClfsxZphkJNXsJRs3K8ZUGRA+xcnjvSwZUmV4ePPPsYZ/3DO6kEVTJ+Jg55ulvo3RADiJnnHbC7tmAjyJMeg+VBFu6zP1d/9WLmz5i70yJDLiLjmOrZ3oDfrRnm9nVVNpUjfE9Q9AUZKZI0h5QJlER+RNwCxC/CyEPdm6lzY+xxhbH/K6HxO7pzHkdMz3LGfgUWdyaMX3ftHSDH/Gs3reP0L76dFzeupZQvEqloJ1y9NcgIQFoLNduo3xEM+J5P31AfyxYczNVf+SUzu6bv1LEDjQy2pRJxzytV7nm5wpNbQ3prEZ4Q5DxB1hP4soHhU2YQ8ff6FRASbWUiS4HSVC3PtfmC/Tt9jp+d5+hZObpz5rmcwOysvkBXZ72DfZz9lffywNMP01XqJIzC1njXOwMLEDPfeYCO96SD0bsx9j2frYN9HHvQkVz5jxfRUWzf6ekS6YxQR+uHIx7aUOXhjTWeHwjZXIkYDjUSs+C2J8C3WyGs42yd2EbnNrIhx4InmJaXLOnKsLwnw0HdWfYtedssx46Sq6tNfZt55zcu4J4n70+Yv86YgOZz8Y5ggZj5riV6T2kA0jjj+Wwe2MrJh53IRf/7h3S1de6SnCFNYm+nebAaaV4cDFnVF/LSYMiGcsSWimIwUNQikzxXUxpfGNOp4AvyniDvC7pzHvPaPPYtecwqesxp88mnwjjOHNtZpk6a0mbPO77+IR5ZtYLu9k7CKDF7WuH97qQGADHz3Uv2GB+gkXzPp3ewl8MXLucXn/8+C2cv2CnRobHIaXHB2Bq5HCbZnYEyyXJ5ay5lvbEXg3Ka3plTu4JcqPO59at5x9c/xFMvPkNXWydBFL76j3c3ik2gdy/RcQ9qs0VyJ2JnGfmeT/9wP3N69uGnf/1djj3wKMIoxJMeu3KGOaelnaaB8Zkp7ndp2hWavu6e2qzm7kmPh1f9ifO/+THWbHyRzmInQRTYMiS97HsMFoJkMV0BdSun7+bY7BCEKqSj1MErfRs592vv46p7rsP3fBC7dnoV1wrEGaK2nnX6o+u/p3+X/uxK5lfKDF/0pMcvbvk1Z33l3azbup6OYgeBClP1yR6GbXBi1nsPrNM3aY21J2EpJUEYEEYhf/22T/F37/gcnpS7fCxBq5IGImvyVIMqf/ezr/H9639KKVck4/tESo0QvFZ4jzsbi5nvPVC7zhgtwPXK70kY+3zSpitsHezljctP5Fsf+SeWzl0Sp/JuzyIcuyOlw8JPr13FJ7/3eW5//G6md0xDaZVMQ2/rbY/FAsSs8w/UscHcQqHMXYl9z6N3qJ+uUidfec/fcsGbzgdMBEQKuUt9g2ZS2tYPo4gf3PBTLrzs22wd3EpnqYMwDFvmHU0KFgIx630HOZ1J80Vy8rAnfYIwYKg6xNnH/hl/9/bPsXy/pYAVBCmZyALdrUxaayKl4ujXvU8/yJcv+ga3P3437YU2Ml6GSO1pMf7xYIGY9X4nAHsfmYQ1Sd9QHx3FDi449b18+s0fYWbXDGD3bxE0GqVU7ONsGdzKt678Lj+48WdUgyodxQ4iFe3wGIrdmcSsDxy0R3aEjQc78qRHFIX0DfezcNYCPnf2J3jfSe8glzGLROxugpA2dQAGy4NcdPtlfO/6H/PU2mfpLHXYZ46MHUzz30WTXAAjAPW7WqV4k4tdKLBSqzBcLXPU4sP5i1Pfy9uOezMdhXYAmwQmbL5/awmD0fYmrcs5uOVahYtv/w3fv+GnrHhhJcV8kUI2b7W++ZWh5td/00yg2R9c6tLeW6ZYzcBgmEgKiRSSwcoQtTBgyb6LeNeJ5/Ke1/8586fPxVErCIPR9BrQdaHcFzet49oHf8fPbv0Vjz6/gnw2TzFbQGmVyuE31Cr13wztrwEx60NL91oTaCwMWEEQlGsVhqvDzOnZhzcfeTpvPvo0jllyFO35UnxupBSuBneVQLg1ttw8/M5/cVSuVbj9ibu49K6rue2xO1m3dT25TI5SrmgHyauWqNtWwoJGAbBgb8bpGtIYc0IKQS2oMVgdJuNlWDR7AScdcjxnvOYUjltyFO2FNtJkpjFPOpLMwBcRX3us6FKyeHVqBBh2lolR+ig29m9ixZqnuHPlPVz30E08/uJTRFFIKV8i62fMypBWYNwLn8IpARAgZv/FUj2yYWAKN1Sb0bgeGkWlVqVcq5DxMyyevT9HLz6CoxYdxrL5S1k6ZwmdpQ7GQ8nKiWMLhaMgDNg8uJUn1z7DA6se4d6nHmTFiytZv3UD1bBKLpOnmMvjhlIaxk+/9l1dV7sjFojZFxzs9k7ROMmZH1prKkGFSlBFKUVbvsTMzukcNOcAluy7iP1mzGfBjHnM7p7JrM4ZdBbbyfpZk4vUQFpr+suDDJQH6C8P0jfcz5qNL/LES0/z3Csv8NwrL/Dy1lfoHe6jXC3jez75bI6sl8HMyal3y5nZmk1i9oenBGBHyNj8JqcwUoogCqgEVcIostMiCoq5At1tXSbPxvPJeBlymSyFbIGM51MNa/QO9TNUHWK4WmaoOsxwZZhQRfEsbFk/g+/5+J6P5xhe6706hr8zyKiiKQtou7GJwkRx/WV8w9xmcWh3jmLrYC+b+jfjZnjQlnmdLSqFxJOesfelpFQoEWe3Un9+6Mbkpq2mFqqT3QYLOy+QwMxq7LbmnCm8PRgwy4Nqu9/WdMbP4JMxPJvKS6lbRR3qBETReB13L8Zdnik8NhYIfJymItmaFzGFdyZW9ptpFRJhGTH6xf6kFcq8p2NNeoZu+4EpPJmYKdwU7LbSzUhmmlUBgik8WRhT92IKTzp2fJ+aG1S4/+gpPEnYkE59m8KTgwEQ4MdLxaT2TuHJxWLqbxP+mu1usVL8XkGmbZ7Ck4lxK8U36KIp3AQspvCkY2wYNHUktZ3CU3jPx7YjLGkZdOrwFJ7CezIWWBPI9SymFwqYwlN4T8da4JbWNTFSYZuDKTx5mCncFOy2vhYmIGQyI8zOxFGewlN4z8SO36Xdj9s6OIUnAQvz0an9U3hysPv4cW6EO8AUnmzcMgK5N2HbFPjJW9CpLS2D04lMyROMH2s90edquOd2X2cy6mT7r2MeZ+R+IWTd/mQR1m2UJx6uLACNSwFvhbp6tXfd8v0AkVbJqhOp8o8Lk57wdvz3DVRUdx2JwJNu2dHxXUcIgTeByXbd7bRWdpDN2NffoTrR5uvI2SuEnWal4dnrnmFkedx8ROnD21Pnk45NNlzrjwjL+3kzqzMTI2G1djWs2Qd+9XuZVSAF00pdqQNQCwOGaxWkFOO6jkBQC2sM1YZjhhP2+oyC0eDGGef8LIVsHoWun6XZPYOAvJ+zA/QnWiuGpBDUwiBZ7dGWwfM8sn4WrU3amEJTC2uj84Ytjyc9ctksSps5lpWyv2kR/tmWcjQCYF9YK40IQ2JfguDij36X+dPmEqpw3NOXa62RUrKhfxPv+sFHGSgP4nlefM2R9zUafutQL58/45P8r9M+QqRUPN539aYXeeu330ctDJBSbuM6RvsN1YZ57f5HcMEJ70Gh7CokY7+Nalilt9zPypef4f7nH+Gp9avIZ7Jk/RxKRbiwKUCoI/77Q//KgbMXT6hOHIUqIp/J8vVr/52f/vHXdBc7AegrD3DWoafyzfO+TC2skfEyrOt9hfP/6xMMVAbxUgInMMM2+8sDvOXw07jw7f9ANaza36znvO/9JbWw9qp11Wwcp0O73Yk+aS52LbMQggU989inaxbbQ1k/a6YJNKtTj3FfEw6LtKK92MY7jn4LHfn2uuscOncpxy46iutX3EJnoZ1IqTHLL6UkiAIWTJ/PO45+y4TLPFgd4ppHf8c/XfPvvNK/gXymEE9qhQQi2LdrNnO6Z0/42mBMFoGgPd+G1gphWzWFopQrMqcruW7Wy5jjkLwUJwbC/KYtX2Lf1PvxPd/8JmUFJaajbiEcm4K2E0xQ10nQXOwqT1ANa3Z67yhlG4+PSrli6oJj30tKyXBtmKMXHMFBsxfHs6+BWThOa825rznTOIRi2+UHEFJQi2qEKiKIgnh2h22R0oowCmnLlXjXa8/lik/+mJkdMwiimjG9Us12MVfYxnV0asrEkeR6ezJepq53VNgyaK2pRUHKfEzdG6OU4o4lRHy/WhjYbS25ovUz0gqtZbDd+tYXiM08QewfNA3Hf4UrsFEjUgh+cfdv+OOq+ynZuS5HI/cs5aBCJaiYVkDo0Z9RgJCSUEWcffgZsWPYXx5g0+AW9p8xH4Hg1KWvZ960OWwe3ErG860mHVl+19JIIfGlR6SMnVyulbnt6XsIozB+GRqjZWe093Do3IPxPR+tNaEKOWDWQv7PW/+GD//8r8lmslZraTwp+cerv0lXoSN+Hmd2nbD4GM4/9s+JdIQnPDYObObCG/6TMApMdAcjGFIIHlrzmJk2ETPtSiLYyfSOQiQPF5c5VW/JsdF/47R/q/BVPY8ZsjM0JSKgU6GiZmG3SeswN//NrU/9kZ/ceREdpS7LTCKeRjDG1r4TCDoKbQhri452LyGgFtWY3TWTM5a90e3lT2tXctmD1/Cv7/wqQRTS09bNyQedwM/u/jXd2U6iSI1e/tSLN6UyaNPgVj7y87+irzyAJ71YgECQ9TMsnX0A33n3/+XQuQfHk26ddegpLJ65gJe2vmymOtRGWC976Fqzhpcw9eJ7Pv2DvYQq4vxj/zxe8bO/MsCP/3gxlaAazxht6tK0IvmMcV6ldGVOs0b6bZghhHVmXyzso1Hjgdbgq7r3bvl9xIgwlxjRTAzYyAgjqC1foru9h+5iZxLFGIM0xOcIMfq9pJQMVwb5s+WnsG/XLGpRQNbLcPsz93DNit/zjT//e3JeFoBzjjiDi+67nLhLYLTyu/psKLsUgo5iBwjwpV8XwREI7lv9MF+44utc86mfx45tMVtk3rQ5rNr0ArlMDmFj8t2lTpwpA6aF8aSkLTVhL4AnJD1t06iGVaTw6uo30sr4AE5jj1Jmt1MIGiJxaY0/2m9ItRp1bXrL/HXbJEBR13S1AJaM/kasJtLj+BfpKHXtMe6FEY5zDz8TMEyjtOKOZ+7lpd6Xuee5BxFCoLTidYuO5sDZiyiHFZNFONY1R1nbVAtjY0daEWG3qe/dbd28MrCRweoQQggT6weyfsaUUSb3ibQi1NGIT6NJqAVE9ljUcK52wurKPdp6rMLUsxrxUSZMW9dG1/8ucSxbiKcazDNE4qu3FIkRbyKh4VqFvkoffZWBV/kMvqrTLIWkEtZYOHMBb1hyHFqbefYfW/skK15+kozn8ds/3QSY8GEhk+eMQ05O+RUTeSajqX3Px5de8rFTJfZXBugpddOWK5lyWEHsLfebEO5YzFZ3h+059ipl9myZPS/18eOtGCMMu313nHxKUiFid2X0WPVkYi0Y4a04U+0TJ32As5afQtbLjMkUAkEtCvg/1/wLr/RvIuNn4mnC088qpWQ4KHPa0tfTnm+jFtbI+lmuf+IW+ioDtOfbuO3puxioDtKWNebF2Yedxvfv+BkhdqEJ0VBvaQ2TIqUVW4Z76bMMTconEcD8aXP44pmfQQpJEIX40mPVxhd4csOz5LN5q7HHqLdR7heTtGaKtHU6Wp1bZe2u4Uywme3TufLjP4kX0673ygSRiugsmJmw0z6GEPYdSveT1uCrkdimQpivyTY5pTkYUhEVu8NpmtfMW85r5i1nPPT/bvouERFZkRn1XhGKfDbPOYefAZgYttKK3z95J4Vsnnw2zwtbX+Le5x/i1INeT6Qils85mMPnLePe1Y/QnivG6wDopKB1lptrzbqKnfzfs79ALazFwgymVZjR1sPxi45mets0QrdCpRD8+63/RV95gK5CpzXnxnhH6aZ9NEqKNmo9JOZg/QUyXoYlMxeOq65Jvz8nUK4+RitzC2CBjQI5xZBWEM3EAmze0lhvdHxk3mv9Ndy9XOz/sDkHc+T8w+JZmO967gHufv4Bcn6W/sogveV+fvXgVbEAZP0sbzn0Tdy56j6ELGGn76xrANLc6Ji9PdfGX7zuXdssr9Ya3y519K3f/5BfPnAVnYUOIhJnf7R6M3cbQwJiRSKom4O0vgEY/bfbQWlH2YapWoavGrEWLh3alhfhGobm4qQtTlWsXdj43tUPsWrTC2S9bGIC6eQnbgHkMAoZrA4bc8MpuvSlpaQa1jhr2alkPN9GfyTPb36R1y06ms58O0orgsisn1sOKuTtqpFnLTuVb978fWpRzZoLur7829LGDRTa9Xk9IYl0xK1P3cVP77mEa1f8PumtJVEMjIK3ec9YJRulMup1Gn7vmHi4VubGlX8giIL6KJA2gl2LauzfM5/j9j8yrgEB5r3J+LajlrlVsO0IM9pBIBAieZRmYffi0u/T2PCSH9z5C35y18V0FLtMLDz1wuqwgI58u+lcQtvm315fmB7e7lIXb1n+JgB8G5t/91Hn8N6jz6WR3OJySiv2mzaX1y06iutW/D5OjXDlH4uZqmGNR156nEhFcVSpkMnzmnnL4z4KreE7t/031z9xG/Om7UsQBrj3I1Is1ojHFgBRtzHmyMjruN/GnV22H2Hz0BY+cckX6C33m/qxTZ1L3Osf7uP8Y88zAqAVJi+6viDjKX8zsOOGEenQ6c6QZmG7o66r3lF7vp1p7dPpsv0AYwoAJnLjHjjtvkkh6a8McuqBJ3DAzIUmec76GKPGtUkcQ8cc5xx6BtesuBknqfX2dOq57PkbBjbxjh9/hL7hAXzfR2sIo4DLPvwj3nTQ6wmikIzn8y9v+0dWrH/KJJOlIizbrDdrw48suo6PJQI5ynWEe4aRz9xT6ibj+Xi2l9rVrZQevpehPVe/PlqdMKaa3FbgqzQWtglIZodObZuN3We01RbjOLRd8jPSY2PXRT/iXlKgdMTZhxrn18Xca1FAX7mf/sog/ZWB+NNXGWCwOgQkgnDKgScyf9ocqlHVdAi567t7NBRdCpOA1l5opz3XRlehg3w2z1du+BbDtTKeTcdYMnMhnz3pw/RVBvBcZOVV6s2cM5oEuLwX4g6vUa9jf99YZk2qH0Eln/i7Ds0aBiNva68nms5Lr8ZjrTkiLLVppFCFVMMa1aj2qj3B2Li+5/k4/SaEyYWf070Ppy99AxBPjcEFF/81d666j7ZcyQwMSRUkVCG//MB3ec285QRRyLRiJ6cceAI/uedXdGdy9bn1oxUF4uVKFSaEWMoWeXTtE3zvzp/x1yd/DGU7sz5+wvu55vGbefilFbRlS0mH3lj1lj42gtJsPkadj1FmW2HUSxqp79v6YbpILcJXo48Ie/VnmHQS1Du2JLin1M1+PXPpzLfHmnsskkIyUBmkrzIYM7mUHsOVQc5efDoz2noIVYgvfZ7Z+By3PPNHIqWoRrXEwUbgS49NQ1u4asXvrM1u7nv28tP4xf2XWTs6Edx0TH3Ec8nkWEREZ7GD79zxU966/HQOmLE/oQrJ+lm+/pYv8JYffhAl0hXRaKSkrrute7r7jvHz5PejXKChzOO5Z2wFbuueLUJ+LM06tYWm4kShJbVrclngS6d/li+c+slRzaM0mZBlhh/edTFfuvb/0VPsMpoUM1jmnOWnx2nWvvS57olbGawO0lOalqzB5e4tBZ2Fdm5++na+cOonyfs5lNYcu+BIDpq1iGc2raZg94FAjykB8YMBxun1pU9fuY+v3fhv/Pz8f0cISaQijp5/OB953Xv59z/8iJ5SN2EUUfeu6upNjM288SGr9Uarc8To9dmo9evg2K2AFg110CJ8NQILmwuU2InUV1rTsAWC2KZ3n5yfoy1XopQtvuon7+fJ+hnTWytACEE5rLBk1kLesPg4hBDkfBPavP7JW8j4WSIUJnXa5roIY7bk/RxPbXiO+9Y8UhfFOWvZqSY3KDXwxkScdJ3drFBxGdLa07QCnfz2iZu55vGb8YS04whC/uqkj3DgrEUM1oaRnnzVejMrR9bXV6zZt/Fb92n8vXPgG8+t/z76PVuHl8bGwgmAC4WlQ2KtgAHacqU4t17a5UjHQ57tUMp6JpkMIRDS5P6cuuQEpDTmUTWscc8LD/PYy09SzBaS1dpT5dGAkIJQhVzx2PVUwirloEw5qHDqkhPpzLcTahPedIyY8TNmvKyfxZMenfl24pUj0wyEOT/rZfjaTf/GcFAm52fJeD7dxS6+8+f/ZITYOvWj1VV8Ty9jllT1skghTfqGq7Bt1LP7ve/59vfmOqVs0d6zXmjT/JKR/oh7xu+ohXhpNKwBMe8rR9dZaa4/oJk4/f2IucvIWy09MTLmyNq+9Ty3+QV8aYY+KK2Y1T6DtqxJY/Ckx9bhPjYPbYmFZqyyaW0Ye27n7DglGg0v9a0jjJwACCpBlQNnLuKspSfHqQ0DlUEueeRqqja82dDrEQ9qOeOgk1i+z1KCKABhTKRLHr6Kl/s3xvlPo5UtVBGz22dwwIz9433loMKj656I+zDGei6BIFABM9umc+DMRWYcrzCj8R5Z+ziRNuOSG99RqCJmtU9nyYyFdb95eO0KM0i+7m00n6/qsZWHuV89Ou4ldgVuNk7vG6qV4zGxY8X7t4WzXsbY7PaIBGpRGHdImfQDn4znGxtevErZtKYaBXX3yvkZXF+DsHcPooDhoBz/WEpJm9WojHZ9DZ4UDFbLBFHgsgjQaDrybVZoxi6bxCQAVsJqXDajxQvjrH/Te14JK/FzSSEoZgqjltmJb6gCykE1rnMhBMVMMe54hMnnn/Fgt/VpIN0COC6ggHY7yCPVbzkh7BaYTldAxs+QJZtiYntOSmU1ls0dEkJQyOTq7mXuY88RRkiyfiZOnXAUxR1z9dfEPqvSmrZ8sU7Lghu8Uv8MjWVzZlcuk60rm7HlGfO3rsK11vi+T2emI/4tgLITAAhGe0fG7Mr6WbvP3FcplTpnG/dtInb14bsIgdNfSZ9p8zAkeXvJII/tbQMw3NWw3xkSY52zrbIlQ0FS54h6bAbkqPr7xq3LyGs6PNbzJsK57bKptG02WtnGekahYYJldncZkWou6s9pFb4aWee7wcxwU3gK70pcP5yndcq19+Apah6J0ZZIEvGxKbwrse0ldXUf75/Ck4YFLTokcq/BYgd+O4V3ArbTouh4Z3o7hScDu7oXU3jSscb6AI1DQqfw5GCgLvQ6hScXCwFS2zfitkIwhScJx9KQ2j+FJw+DSMYEJ13ESTfMFN7VuNlGwN6MDclEGqiDU3gycNMLsBdjs/3/Acpvnw1+YfUjAAAAAElFTkSuQmCC">
<title>アース建設 現場日報</title>
<style>
:root{--bg:#FAFAF6;--panel:#fff;--ink:#22271F;--sub:#5D6357;--line:#DAD7CB;
 --green:#1E6B4F;--green-soft:#E4EFE4;--green-ink:#175640;
 --orange:#C96F2E;--orange-soft:#F9E8D8;--orange-ink:#9E551F;--warn:#B3261E}
*{box-sizing:border-box;-webkit-tap-highlight-color:transparent}
html,body{margin:0;background:var(--bg);color:var(--ink);
 font-family:-apple-system,"Hiragino Kaku Gothic ProN","Noto Sans JP",sans-serif;font-size:16px}
.app{max-width:480px;margin:0 auto;min-height:100vh;display:flex;flex-direction:column;background:var(--bg)}
.appbar{background:var(--green);color:#fff;padding:13px 16px;display:flex;align-items:center;gap:10px;
 position:sticky;top:0;z-index:5}
.appbar .t{font-weight:700;font-size:17px}
.logomark{height:30px;width:auto;flex:none}
.appbar .u{margin-left:auto;font-size:12px;background:rgba(255,255,255,.22);padding:3px 10px;border-radius:99px}
.appbar .back{background:none;border:none;color:#fff;font-size:20px;padding:0 4px;cursor:pointer}
main{flex:1;padding:16px 15px 40px}
.screen{display:none}.screen.on{display:block}
label{display:block;font-size:12px;font-weight:700;color:var(--sub);margin:14px 0 4px;letter-spacing:.04em}
input,select,textarea{width:100%;font-size:16px;padding:11px 12px;border:1px solid var(--line);
 border-radius:10px;background:var(--panel);color:var(--ink);font-family:inherit}
textarea{min-height:64px}
.row2{display:flex;gap:10px;align-items:center}
.seg{display:flex;border:1.5px solid var(--line);border-radius:11px;overflow:hidden;margin-top:4px}
.seg button{flex:1;padding:11px 0;font-size:15px;font-weight:700;border:none;background:var(--panel);
 color:var(--sub);cursor:pointer}
.seg button.on{background:var(--green);color:#fff}
.seg button.on.orange{background:var(--orange)}
.chips{display:flex;flex-wrap:wrap;gap:8px;margin-top:4px}
.chip{border:1.5px solid var(--line);border-radius:99px;padding:8px 16px;font-size:15px;
 background:var(--panel);cursor:pointer;user-select:none}
.chip.on{background:var(--green-soft);border-color:var(--green);color:var(--green-ink);font-weight:700}
.chip.small{padding:6px 12px;font-size:13px}
.bigbtn{display:block;width:100%;background:var(--green);color:#fff;border:none;border-radius:12px;
 padding:15px 0;font-size:17px;font-weight:700;margin:16px 0 0;cursor:pointer;font-family:inherit}
.bigbtn.sub{background:var(--panel);color:var(--green-ink);border:1.5px solid var(--green)}
.bigbtn.orange{background:var(--orange)}
.bigbtn:disabled{opacity:.5}
.card{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:11px 13px;margin:8px 0}
.rowline{display:flex;align-items:center;gap:8px}
.rowline .grow{flex:1;font-weight:700}
.meta{font-size:13px;color:var(--sub);margin-top:3px}
.tag{display:inline-block;font-size:12px;font-weight:700;padding:2px 10px;border-radius:99px;white-space:nowrap}
.tag.g{background:var(--green-soft);color:var(--green-ink)}
.tag.o{background:var(--orange-soft);color:var(--orange-ink)}
.addbtn{width:100%;border:1.5px dashed var(--green);color:var(--green-ink);border-radius:12px;
 background:none;padding:13px 0;font-size:15px;font-weight:700;margin-top:10px;cursor:pointer;font-family:inherit}
.hint{font-size:12.5px;color:var(--sub);margin-top:6px;line-height:1.6}
.today{font-size:13px;color:var(--sub);font-weight:700;margin:16px 0 6px}
.toast{position:fixed;left:50%;bottom:30px;transform:translateX(-50%);background:var(--green-ink);
 color:#fff;padding:11px 22px;border-radius:99px;font-size:14px;font-weight:700;z-index:99;
 opacity:0;transition:opacity .25s;pointer-events:none;max-width:90%;text-align:center}
.toast.on{opacity:1}
.toast.err{background:var(--warn)}
.banner{background:var(--green-soft);color:var(--green-ink);border-radius:11px;padding:10px 14px;
 font-size:14.5px;font-weight:700;text-align:center;margin-bottom:10px}
.prev{border-bottom:1px dashed var(--line);padding:8px 0;display:flex;gap:8px;align-items:center;font-size:14px}
.prev:last-child{border-bottom:none}
.prev .nm{flex:1;font-weight:700}
.prev .nk{font-variant-numeric:tabular-nums;font-weight:700}
.navrow{display:flex;justify-content:space-between;margin:12px 2px;font-weight:700;color:var(--green-ink)}
.navrow button{background:none;border:none;color:var(--green-ink);font-size:15px;font-weight:700;cursor:pointer}
.thumbs{display:flex;flex-wrap:wrap;gap:8px;margin-top:6px}
.thumbs img{width:72px;height:72px;object-fit:cover;border-radius:10px;border:1px solid var(--line)}
.thumbs .del{position:absolute;top:-6px;right:-6px;background:var(--warn);color:#fff;border:none;
 border-radius:50%;width:22px;height:22px;font-size:12px;cursor:pointer}
.thumbwrap{position:relative}
.photobtn{border:1.5px dashed var(--line);border-radius:10px;width:72px;height:72px;background:none;
 font-size:24px;color:var(--sub);cursor:pointer}
.spin{display:inline-block;width:16px;height:16px;border:2.5px solid rgba(255,255,255,.4);
 border-top-color:#fff;border-radius:50%;animation:sp 1s linear infinite;vertical-align:-3px;margin-right:8px}
@keyframes sp{to{transform:rotate(360deg)}}
.loading{padding:60px 0;text-align:center;color:var(--sub)}
</style>
</head>
<body>
<div class="app">
  <div class="appbar">
    <button class="back" id="backBtn" hidden onclick="goBack()">‹</button>
    <svg xmlns="http://www.w3.org/2000/svg" class="logomark" viewBox="95 115 570 570">
  <defs>
    <linearGradient id="em-bg" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0" stop-color="#0A130C"/>
      <stop offset="0.45" stop-color="#123D1C"/>
      <stop offset="1" stop-color="#1E8A38"/>
    </linearGradient>
    <radialGradient id="em-globe" cx="0.5" cy="0.58" r="0.72">
      <stop offset="0" stop-color="#EAF7FF"/>
      <stop offset="0.28" stop-color="#A8DFF6"/>
      <stop offset="0.62" stop-color="#4FB4E8"/>
      <stop offset="1" stop-color="#1E8ECC"/>
    </radialGradient>
    <clipPath id="em-rim"><circle cx="380" cy="400" r="268"/></clipPath>
    <mask id="em-mouth" maskUnits="userSpaceOnUse" x="0" y="0" width="760" height="1000">
      <rect width="760" height="1000" fill="#fff"/>
      <g transform="rotate(-35 380 400)">
        <path d="M 680.8 475.0 A 310 310 0 0 1 624.3 590.9 L 498.2 492.4 A 150 150 0 0 0 525.6 436.3 Z" fill="#000"/>
      </g>
    </mask>
  </defs>
  <g mask="url(#em-mouth)">
    <g transform="rotate(-35 380 400)">
      <circle cx="380" cy="400" r="270" fill="#FFFFFF"/>
      <circle cx="380" cy="400" r="198" fill="url(#em-globe)"/>
      <!-- 横棒（円の内側に収める） -->
      <rect x="110" y="340" width="540" height="120" fill="#FFFFFF" clip-path="url(#em-rim)"/>
    </g>
  </g>
</svg>
    <span class="t" id="barTitle">アース建設 現場日報</span>
    <span class="u" id="barUser">…</span>
  </div>
  <main>

    <!-- ホーム -->
    <section class="screen" id="scr-home">
      <div class="today" id="homeDate"></div>
      <button class="bigbtn" onclick="openForm()">＋　今日の日報を入力</button>
      <div class="today">今日の入力状況</div>
      <div class="card" id="homeToday"><div class="hint">読み込み中…</div></div>
      <button class="bigbtn sub" onclick="openHistory()">じぶんの履歴を見る</button>
      <button class="bigbtn sub" id="ledgerBtn">台帳・集計を見る</button>
      <div class="hint">※「台帳・集計」はスプレッドシートが開きます。編集は台帳のルールに沿ってください。</div>
    </section>

    <!-- 今日の日報（作業一覧） -->
    <section class="screen" id="scr-form">
      <label>日付</label><input type="date" id="fDate">
      <label>現場</label><select id="fSite"></select>
      <label>今日の作業（自社も業者も、何件でも）</label>
      <div id="workList"></div>
      <button class="addbtn" onclick="openAdd(-1)">＋　作業を追加（自社／業者）</button>
      <label>今日の写真（任意・自動でドライブの現場フォルダへ）</label>
      <div class="thumbs" id="thumbs">
        <button class="photobtn" onclick="pickPhoto()">📷</button>
      </div>
      <input type="file" id="photoInput" accept="image/*" multiple hidden>
      <label>備考（あれば）</label><textarea id="fBiko" placeholder="例：午後から雨のため2時上がり"></textarea>
      <button class="bigbtn" id="toConfirm" onclick="openConfirm()">確認へ進む</button>
    </section>

    <!-- 作業を追加 -->
    <section class="screen" id="scr-add">
      <label>区分</label>
      <div class="seg">
        <button id="segJi" onclick="setKubun('自社')">自社</button>
        <button id="segGy" onclick="setKubun('業者')">業者</button>
      </div>
      <label>工種（選ぶ／手入力 どちらもOK）</label>
      <input id="aKoushu" list="dlKoushu" placeholder="タップして選択か入力">
      <datalist id="dlKoushu"></datalist>
      <div class="hint" id="koushuParent"></div>
      <div id="newKoushuBox" hidden>
        <label style="color:var(--orange-ink)">この工種はマスターに無いため、工程を選んでください</label>
        <select id="aParent"></select>
        <label class="chips" style="margin-top:8px"><span class="chip small on" id="addKoushuChip" onclick="toggleChip(this)">マスターに追加する（推奨）</span></label>
      </div>
      <div id="jiBox">
        <label>作業員（複数選べます・1人＝1行）</label>
        <div class="chips" id="workerChips"></div>
        <label>1人あたりの人工</label>
        <select id="aNinkuEach"><option value="1" selected>1.0（通常。半日でも1）</option><option value="0.5">0.5（特別な場合のみ）</option></select>
      </div>
      <div id="gyBox" hidden>
        <label>業者名（選ぶ／手入力 どちらもOK）</label>
        <input id="aGyousha" list="dlGyousha" placeholder="タップして選択か入力">
        <datalist id="dlGyousha"></datalist>
        <label>人工（0.5刻み：半日＝0.5、2人1日＝2.0）</label>
        <select id="aNinku"></select>
      </div>
      <label>作業内容（選ぶ／手入力 どちらもOK）</label>
      <input id="aContent" list="dlContent" placeholder="タップして選択か入力">
      <datalist id="dlContent"></datalist>
      <label>機械（あれば・選ぶ／手入力OK）</label>
      <input id="aMachine" list="dlMachine" placeholder="タップして選択か入力">
      <datalist id="dlMachine"></datalist>
      <label>時間帯（任意）</label>
      <div class="row2">
        <select id="aStart"></select><span>〜</span><select id="aEnd"></select>
      </div>
      <button class="bigbtn" onclick="saveWork()">この作業を追加</button>
      <button class="bigbtn sub" id="delWork" hidden onclick="deleteWork()">この作業を削除</button>
    </section>

    <!-- 確認 -->
    <section class="screen" id="scr-confirm">
      <div class="banner" id="confirmBanner"></div>
      <div class="card" id="confirmRows"></div>
      <div class="hint" id="confirmMeta"></div>
      <button class="bigbtn" id="sendBtn" onclick="send()">この内容で送信</button>
      <button class="bigbtn sub" onclick="show('form')">戻って直す</button>
    </section>

    <!-- 履歴 -->
    <section class="screen" id="scr-history">
      <div class="chips">
        <span class="chip small" onclick="histJump(1)">昨日</span>
        <span class="chip small" onclick="histJump(3)">3日前</span>
        <span class="chip small" onclick="histJump(7)">1週間前</span>
        <span class="chip small" onclick="histJump(30)">1か月前</span>
      </div>
      <label>日付を指定</label><input type="date" id="hDate" onchange="loadHistory()">
      <label>表示する範囲</label>
      <div class="seg">
        <button id="hMine" class="on" onclick="setScope('mine')">じぶんの分</button>
        <button id="hAll" onclick="setScope('all')">現場全体</button>
      </div>
      <div class="navrow">
        <button onclick="histShift(-1)">◀ 前の日</button>
        <span class="today" id="hLabel" style="margin:0"></span>
        <button onclick="histShift(1)">次の日 ▶</button>
      </div>
      <div class="card" id="histRows"><div class="hint">読み込み中…</div></div>
      <div class="hint">「じぶんの分」はアプリから自分が送信した行と、作業員名が自分の行です。アプリ導入前（9/2以前）の日報は「現場全体」でご覧ください。</div>
    </section>

    <div class="loading" id="loading">読み込み中です…<br>初回はログイン確認が出ます。</div>
  </main>
</div>
<div class="toast" id="toast"></div>

<script>
var INIT = null;            // サーバーからの初期データ
var works = [];             // 追加された作業
var photos = [];            // {name, mime, dataB64, url}
var editIndex = -1;         // 作業の編集中インデックス
var curKubun = '自社';
var histDate = null;
var histScope = 'mine';
var screenStack = [];

function $(id){ return document.getElementById(id); }
function toast(msg, err){
  var t = $('toast'); t.textContent = msg;
  t.className = 'toast on' + (err ? ' err' : '');
  setTimeout(function(){ t.className = 'toast'; }, err ? 4200 : 2500);
}
function show(name){
  document.querySelectorAll('.screen').forEach(function(s){ s.classList.remove('on'); });
  $('scr-' + name).classList.add('on');
  var titles = {home:'アース建設 現場日報', form:'今日の日報', add:'作業を追加',
                confirm:'内容の確認', history:'じぶんの履歴'};
  $('barTitle').textContent = titles[name];
  $('backBtn').hidden = (name === 'home');
  if (screenStack[screenStack.length-1] !== name) screenStack.push(name);
  window.scrollTo(0,0);
}
function goBack(){
  screenStack.pop();
  show(screenStack.pop() || 'home');
}

/* ---------- 初期化 ---------- */
google.script.run.withSuccessHandler(function(d){
  INIT = d;
  $('loading').style.display = 'none';
  $('barUser').textContent = (d.userName || d.user.split('@')[0]) + 'さん';
  $('homeDate').textContent = fmtJa(d.today);
  $('ledgerBtn').onclick = function(){ window.open(d.ledgerUrl, '_blank'); };
  fillDatalist('dlKoushu', d.kouShu.map(function(k){ return k.name; }));
  fillDatalist('dlGyousha', d.gyousha);
  fillDatalist('dlContent', d.contents);
  fillDatalist('dlMachine', d.machines);
  fillSelect('fSite', d.sites);
  fillSelect('aParent', d.kouTei);
  var chips = $('workerChips');
  d.workers.forEach(function(w){
    var c = document.createElement('span');
    c.className = 'chip'; c.textContent = w;
    c.onclick = function(){ toggleChip(c); };
    chips.appendChild(c);
  });
  var nk = $('aNinku');
  for (var v = 0.5; v <= 10; v += 0.5){
    var o = document.createElement('option'); o.value = v; o.textContent = v.toFixed(1);
    if (v === 1) o.selected = true; nk.appendChild(o);
  }
  fillTimes($('aStart'), '');
  fillTimes($('aEnd'), '');
  $('fDate').value = d.today;
  renderTodayCard(d.todayEntries);
  restoreDraft();
  show('home');
}).withFailureHandler(function(e){
  $('loading').textContent = '読み込みに失敗しました：' + e.message;
}).api_init();

function fillDatalist(id, arr){
  $(id).innerHTML = arr.map(function(v){ return '<option value="' + esc(v) + '">'; }).join('');
}
function fillSelect(id, arr){
  $(id).innerHTML = arr.map(function(v){ return '<option>' + esc(v) + '</option>'; }).join('');
}
function fillTimes(sel, def){
  var html = '<option value="">--:--</option>';
  for (var h = 6; h <= 20; h++) ['00','30'].forEach(function(m){
    var t = h + ':' + m; html += '<option' + (t === def ? ' selected' : '') + '>' + t + '</option>';
  });
  sel.innerHTML = html;
}
function esc(s){ return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/"/g,'&quot;'); }
function fmtJa(iso){
  var d = new Date(iso + 'T00:00:00');
  return (d.getMonth()+1) + '月' + d.getDate() + '日（' + '日月火水木金土'[d.getDay()] + '）';
}
function toggleChip(c){ c.classList.toggle('on'); }

/* ---------- ホーム ---------- */
function renderTodayCard(entries){
  var el = $('homeToday');
  if (!entries || !entries.length){
    el.innerHTML = '<div class="hint">まだ今日の日報はありません。</div>'; return;
  }
  el.innerHTML = entries.map(function(e){
    return '<div class="prev"><span class="tag ' + (e.kubun === '自社' ? 'g' : 'o') + '">' + e.kubun +
      '</span><span class="nm">' + esc(e.name) + '　<span class="meta">' + esc(e.koushu) + '</span></span>' +
      '<span class="nk">' + (e.ninku === '' ? '—' : e.ninku.toFixed(1)) + '</span></div>';
  }).join('');
}
function openForm(){ show('form'); }
function openHistory(){
  histDate = INIT.today; $('hDate').value = histDate;
  show('history'); loadHistory();
}

/* ---------- 作業一覧 ---------- */
function renderWorks(){
  var el = $('workList');
  el.innerHTML = works.map(function(w, i){
    var who = w.kubun === '自社'
      ? w.workers.join('・') + '　＝ ' + w.workers.length + '行'
      : esc(w.gyousha) + '　' + Number(w.ninku).toFixed(1) + '人工';
    var time = (w.start || w.end) ? '　' + (w.start||'') + '〜' + (w.end||'') : '';
    return '<div class="card" onclick="openAdd(' + i + ')"><div class="rowline">' +
      '<span class="tag ' + (w.kubun==='自社'?'g':'o') + '">' + w.kubun + '</span>' +
      '<span class="grow">' + esc(w.koushu) + '</span><span class="meta">編集 ›</span></div>' +
      '<div class="meta">' + who + time + (w.content ? '　／ ' + esc(w.content) : '') + '</div></div>';
  }).join('');
  var n = countRows();
  $('toConfirm').textContent = n ? '確認へ進む（' + n + '行）' : '確認へ進む';
  saveDraft();
}
function countRows(){
  return works.reduce(function(n, w){
    return n + (w.kubun === '自社' ? w.workers.length : 1);
  }, 0);
}

/* ---------- 作業を追加 ---------- */
function openAdd(i){
  editIndex = i;
  var w = i >= 0 ? works[i] : null;
  setKubun(w ? w.kubun : '自社');
  $('aKoushu').value = w ? w.koushu : '';
  $('aContent').value = w ? w.content : '';
  $('aMachine').value = w ? w.machine : '';
  $('aGyousha').value = w && w.gyousha ? w.gyousha : '';
  $('aNinku').value = w && w.ninku ? w.ninku : 1;
  $('aNinkuEach').value = w && w.ninkuEach ? w.ninkuEach : 1;
  $('aStart').value = w ? w.start : '';
  $('aEnd').value = w ? w.end : '';
  document.querySelectorAll('#workerChips .chip').forEach(function(c){
    c.classList.toggle('on', !!(w && w.workers && w.workers.indexOf(c.textContent) >= 0));
  });
  $('delWork').hidden = (i < 0);
  onKoushuInput();
  show('add');
}
$('aKoushu').addEventListener('input', onKoushuInput);
function onKoushuInput(){
  var v = $('aKoushu').value.trim();
  var hit = INIT ? INIT.kouShu.filter(function(k){ return k.name === v; })[0] : null;
  $('koushuParent').textContent = hit ? '→ 工程：' + hit.parent + '（マスターから自動）' : '';
  $('newKoushuBox').hidden = !v || !!hit;
}
function setKubun(k){
  curKubun = k;
  $('segJi').className = k === '自社' ? 'on' : '';
  $('segGy').className = k === '業者' ? 'on orange' : '';
  $('jiBox').hidden = k !== '自社';
  $('gyBox').hidden = k !== '業者';
}
function saveWork(){
  var koushu = $('aKoushu').value.trim();
  if (!koushu) return toast('工種を選ぶか入力してください。', true);
  var isNew = !INIT.kouShu.some(function(k){ return k.name === koushu; });
  var w = { kubun: curKubun, koushu: koushu,
    koushuParent: isNew ? $('aParent').value : '',
    content: $('aContent').value.trim(), machine: $('aMachine').value.trim(),
    start: $('aStart').value, end: $('aEnd').value,
    addToMaster: {
      koushu: isNew && $('addKoushuChip').classList.contains('on'),
      gyousha: false, content: false, machine: false } };
  if (curKubun === '自社'){
    w.workers = [].slice.call(document.querySelectorAll('#workerChips .chip.on'))
      .map(function(c){ return c.textContent; });
    if (!w.workers.length) return toast('作業員を選んでください。', true);
    w.ninkuEach = Number($('aNinkuEach').value);
  } else {
    w.gyousha = $('aGyousha').value.trim();
    if (!w.gyousha) return toast('業者名を選ぶか入力してください。', true);
    w.ninku = Number($('aNinku').value);
    w.addToMaster.gyousha = INIT.gyousha.indexOf(w.gyousha) < 0;
  }
  w.addToMaster.content = !!w.content && INIT.contents.indexOf(w.content) < 0;
  w.addToMaster.machine = !!w.machine && INIT.machines.indexOf(w.machine) < 0;
  if (editIndex >= 0) works[editIndex] = w; else works.push(w);
  renderWorks(); show('form');
}
function deleteWork(){
  if (editIndex >= 0) works.splice(editIndex, 1);
  renderWorks(); show('form');
}

/* ---------- 写真 ---------- */
function pickPhoto(){ $('photoInput').click(); }
$('photoInput').addEventListener('change', function(){
  [].slice.call(this.files).forEach(addPhoto);
  this.value = '';
});
function addPhoto(file){
  var img = new Image();
  var reader = new FileReader();
  reader.onload = function(){ img.src = reader.result; };
  img.onload = function(){
    var MAX = 1600, s = Math.min(1, MAX / Math.max(img.width, img.height));
    var cv = document.createElement('canvas');
    cv.width = Math.round(img.width * s); cv.height = Math.round(img.height * s);
    cv.getContext('2d').drawImage(img, 0, 0, cv.width, cv.height);
    var dataUrl = cv.toDataURL('image/jpeg', 0.82);
    photos.push({ name: file.name || 'photo.jpg', mime: 'image/jpeg',
      dataB64: dataUrl.split(',')[1], url: dataUrl });
    renderThumbs();
  };
  reader.readAsDataURL(file);
}
function renderThumbs(){
  var el = $('thumbs');
  el.innerHTML = photos.map(function(p, i){
    return '<span class="thumbwrap"><img src="' + p.url + '">' +
      '<button class="del" onclick="delPhoto(' + i + ')">✕</button></span>';
  }).join('') + '<button class="photobtn" onclick="pickPhoto()">📷</button>';
}
function delPhoto(i){ photos.splice(i, 1); renderThumbs(); }

/* ---------- 確認・送信 ---------- */
function buildRows(){
  var rows = [];
  works.forEach(function(w){
    if (w.kubun === '自社') w.workers.forEach(function(nm){
      rows.push({ kubun:'自社', name:nm, koushu:w.koushu, ninku:w.ninkuEach || 1 });
    });
    else rows.push({ kubun:'業者', name:w.gyousha, koushu:w.koushu, ninku:w.ninku });
  });
  return rows;
}
function openConfirm(){
  if (!works.length) return toast('作業を1件以上追加してください。', true);
  var rows = buildRows();
  $('confirmBanner').textContent = '台帳に ' + rows.length + ' 行追加します';
  $('confirmRows').innerHTML = rows.map(function(r){
    return '<div class="prev"><span class="tag ' + (r.kubun==='自社'?'g':'o') + '">' + r.kubun +
      '</span><span class="nm">' + esc(r.name) + '</span><span class="meta">' + esc(r.koushu) +
      '</span><span class="nk">' + Number(r.ninku).toFixed(1) + '</span></div>';
  }).join('');
  $('confirmMeta').innerHTML = fmtJa($('fDate').value) + '・' + esc($('fSite').value) +
    (photos.length ? '<br>📷 写真' + photos.length + '枚 → ドライブの現場フォルダへ保存' : '');
  show('confirm');
}
function send(){
  var btn = $('sendBtn');
  btn.disabled = true;
  btn.innerHTML = '<span class="spin"></span>送信中…（写真がある時は少し時間がかかります）';
  var payload = { date: $('fDate').value, site: $('fSite').value,
    biko: $('fBiko').value.trim(), works: works,
    photos: photos.map(function(p){ return { name:p.name, mime:p.mime, dataB64:p.dataB64 }; }) };
  google.script.run.withSuccessHandler(function(res){
    toast('台帳に ' + res.added + ' 行追加しました' + (res.photos ? '・写真' + res.photos + '枚保存' : ''));
    works = []; photos = []; $('fBiko').value = '';
    renderWorks(); renderThumbs(); clearDraft();
    renderTodayCard(res.entries);
    btn.disabled = false; btn.textContent = 'この内容で送信';
    show('home');
  }).withFailureHandler(function(e){
    btn.disabled = false; btn.textContent = 'この内容で送信';
    toast('送信できませんでした：' + e.message, true);
  }).api_submit(payload);
}

/* ---------- 履歴 ---------- */
function setScope(s){
  histScope = s;
  $('hMine').className = s === 'mine' ? 'on' : '';
  $('hAll').className = s === 'all' ? 'on' : '';
  loadHistory();
}
function histJump(daysAgo){
  var d = new Date(INIT.today + 'T00:00:00');
  d.setDate(d.getDate() - daysAgo);
  histDate = isoDate(d); $('hDate').value = histDate; loadHistory();
}
function histShift(dir){
  var d = new Date(($('hDate').value || INIT.today) + 'T00:00:00');
  d.setDate(d.getDate() + dir);
  histDate = isoDate(d); $('hDate').value = histDate; loadHistory();
}
function isoDate(d){
  return d.getFullYear() + '-' + ('0'+(d.getMonth()+1)).slice(-2) + '-' + ('0'+d.getDate()).slice(-2);
}
function loadHistory(){
  histDate = $('hDate').value || INIT.today;
  $('hLabel').textContent = fmtJa(histDate);
  $('histRows').innerHTML = '<div class="hint">読み込み中…</div>';
  google.script.run.withSuccessHandler(function(rows){
    if (!rows.length){
      $('histRows').innerHTML = '<div class="hint">この日の記録はありません。</div>'; return;
    }
    $('histRows').innerHTML = rows.map(function(e){
      var t = (e.start || e.end) ? '　' + e.start + '〜' + e.end : '';
      return '<div class="prev"><span class="tag ' + (e.kubun==='自社'?'g':'o') + '">' + e.kubun +
        '</span><span class="nm">' + esc(e.name) + '<div class="meta">' + esc(e.site) + '／' +
        esc(e.koushu) + (e.content ? '／' + esc(e.content) : '') + t + '</div></span>' +
        '<span class="nk">' + (e.ninku === '' ? '—' : Number(e.ninku).toFixed(1)) + '</span></div>';
    }).join('');
  }).withFailureHandler(function(e){
    $('histRows').innerHTML = '<div class="hint">読み込みに失敗しました：' + e.message + '</div>';
  }).api_history(histDate, histScope);
}

/* ---------- 下書き（送信前にアプリを閉じても消えない） ---------- */
function saveDraft(){
  try{ localStorage.setItem('nippoDraft', JSON.stringify({
    date: $('fDate').value, site: $('fSite').value, biko: $('fBiko').value, works: works })); }catch(e){}
}
function restoreDraft(){
  try{
    var d = JSON.parse(localStorage.getItem('nippoDraft') || 'null');
    if (d && d.works && d.works.length && d.date === INIT.today){
      works = d.works; $('fBiko').value = d.biko || ''; renderWorks();
      toast('前回の下書きを復元しました');
    } else { renderWorks(); }
  }catch(e){ renderWorks(); }
}
function clearDraft(){ try{ localStorage.removeItem('nippoDraft'); }catch(e){} }
</script>
</body>
</html>
`;
