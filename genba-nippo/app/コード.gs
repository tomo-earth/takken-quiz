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

// ★★ 初期設定: 台帳スプレッドシートの ID を貼り付けてください ★★
// （スプレッドシートの URL の /d/ と /edit の間の文字列）
var SPREADSHEET_ID = 'ここに台帳スプレッドシートのIDを貼り付け';

var SHEET_LEDGER = '台帳';
var SHEET_MASTER = 'マスター';
var SHEET_CONFIG = '設定';
var LEDGER_MAX_ROW = 2000;   // 数式・書式を用意してある最終行

function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('アース建設 現場日報')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

function ss_() { return SpreadsheetApp.openById(SPREADSHEET_ID); }

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
    try { ensureDateRows_(ss); } catch (e) { /* 集計の延長に失敗しても日報の登録は成功扱い */ }
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
 *  集計シートの日付範囲を台帳に合わせて自動で延ばす
 *  （日別集計・工種別日別・印刷01・印刷02 の4シートに行を足し、
 *    数式・書式をコピーして合計行の範囲も広げる）
 * ===================================================================== */
function ensureDateRows_(ss) {
  var daily = ss.getSheetByName('日別集計');
  if (!daily) return 0;

  // 集計側の最終日付行を探す（合計行の1つ上）
  var colA = daily.getRange(1, 1, daily.getLastRow(), 1).getValues();
  var lastDateRow = 0, lastDate = null;
  for (var i = colA.length - 1; i >= 0; i--) {
    if (colA[i][0] instanceof Date) { lastDateRow = i + 1; lastDate = colA[i][0]; break; }
  }
  if (!lastDateRow) return 0;

  // 台帳の最大日付
  var vals = ss.getSheetByName(SHEET_LEDGER)
    .getRange(2, 1, LEDGER_MAX_ROW - 1, 1).getValues();
  var maxDate = null;
  for (var j = 0; j < vals.length; j++) {
    var v = vals[j][0];
    if (v instanceof Date && (!maxDate || v > maxDate)) maxDate = v;
  }
  if (!maxDate) return 0;

  var need = Math.round((dayOnly_(maxDate) - dayOnly_(lastDate)) / 86400000);
  if (need <= 0) return 0;
  if (need > 400) need = 400;          // 日付の打ち間違い対策

  extendDateSheet_(daily, lastDateRow, need, true);
  extendDateSheet_(ss.getSheetByName('工種別日別'), lastDateRow, need, true);
  extendDateSheet_(ss.getSheetByName('印刷01_日別人工'), lastDateRow, need, false);
  extendDateSheet_(ss.getSheetByName('印刷02_累計推移'), lastDateRow, need, false);
  return need;
}

function dayOnly_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime(); }

function extendDateSheet_(sh, lastDateRow, need, fillDate) {
  if (!sh) return;
  var lastCol = sh.getLastColumn();
  sh.insertRowsAfter(lastDateRow, need);
  sh.getRange(lastDateRow, 1, 1, lastCol)
    .copyTo(sh.getRange(lastDateRow + 1, 1, need, lastCol));
  if (!fillDate) return;

  // 日付を1日ずつ入れる
  var base = sh.getRange(lastDateRow, 1).getValue();
  var out = [];
  for (var i = 1; i <= need; i++) {
    var d = new Date(base); d.setDate(d.getDate() + i); out.push([d]);
  }
  sh.getRange(lastDateRow + 1, 1, need, 1).setValues(out);

  // 合計行の SUM を新しい範囲に広げる（SUM の式が入っている列だけ）
  var totalRow = lastDateRow + need + 1;
  if (String(sh.getRange(totalRow, 1).getValue()).indexOf('合計') < 0) return;
  var formulas = sh.getRange(totalRow, 1, 1, lastCol).getFormulas()[0];
  for (var c = 0; c < formulas.length; c++) {
    if (/^=SUM\(/i.test(formulas[c])) {
      var letter = colLetter_(c + 1);
      sh.getRange(totalRow, c + 1)
        .setFormula('=SUM(' + letter + '5:' + letter + (lastDateRow + need) + ')');
    }
  }
}

function colLetter_(n) {
  var s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m) / 26; }
  return s;
}

/* =====================================================================
 *  マスター修正（1回だけ実行）
 *   ・工種「外構工事」の工程が「準備・仮設工事」になっていたのを直す
 *   ・工種「内部大工」をマスターに追加（工程＝木工事）
 *   ・集計シートの日付範囲を台帳の最新日まで延ばす
 * ===================================================================== */
function マスター修正() {
  var ss = ss_();
  var master = ss.getSheetByName(SHEET_MASTER);
  var log = [];

  var fixes = [['外構工事', '外構工事'], ['内部大工', '木工事']];
  fixes.forEach(function (pair) {
    var koushu = pair[0], koutei = pair[1];
    var names = colValues_(master, 5, 5);
    var idx = names.indexOf(koushu);
    if (idx >= 0) {
      var row = 5 + idx;
      var cur = String(master.getRange(row, 6).getValue()).trim();
      if (cur !== koutei) {
        master.getRange(row, 6).setValue(koutei);
        log.push('工種「' + koushu + '」の工程を ' + (cur || '空欄') + ' → ' + koutei + ' に修正');
      }
    } else {
      master.getRange(5 + names.length, 5, 1, 2).setValues([[koushu, koutei]]);
      log.push('工種「' + koushu + '」を追加（工程＝' + koutei + '）');
    }
  });

  var n = ensureDateRows_(ss);
  if (n) log.push('集計シートに ' + n + ' 日分の行を追加しました');

  var msg = log.length ? '✅ ' + log.join('\n✅ ') : '修正の必要はありませんでした。';
  Logger.log(msg);
  return msg;
}
