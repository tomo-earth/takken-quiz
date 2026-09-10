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

/* ---- 日付はすべて「スプレッドシートの時計（タイムゾーン）」基準で扱う ----
 *  スクリプトとシートの時計がずれていても、台帳の日付が 00:00 ちょうどで書かれ、
 *  日別集計の日付と一致するようにする */
var TZ_CACHE_ = null;
function tz_() {
  if (!TZ_CACHE_) TZ_CACHE_ = ss_().getSpreadsheetTimeZone() || 'Asia/Tokyo';
  return TZ_CACHE_;
}
function dateOnly_(d) { return Utilities.formatDate(d, tz_(), 'yyyy-MM-dd'); }          // Date → 'yyyy-MM-dd'
function toSheetDate_(ymd) { return Utilities.parseDate(ymd, tz_(), 'yyyy-MM-dd'); }   // 'yyyy-MM-dd' → 00:00 の Date
function addDays_(ymd, n) {
  var p = ymd.split('-');
  return Utilities.formatDate(new Date(Date.UTC(+p[0], +p[1] - 1, +p[2] + n)), 'UTC', 'yyyy-MM-dd');
}
function daysBetween_(ymdA, ymdB) {
  var a = ymdA.split('-'), b = ymdB.split('-');
  return Math.round((Date.UTC(+b[0], +b[1] - 1, +b[2]) - Date.UTC(+a[0], +a[1] - 1, +a[2])) / 86400000);
}
function todayYmd_() { return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd'); }

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
    today: todayYmd_(),
    sites: colValues_(master, 1, 5),                   // A列 現場
    workers: colValues_(master, 8, 5),                 // H列 作業員(自社)
    gyousha: colValues_(master, 10, 5),                // J列 業者
    machines: colValues_(master, 12, 5),               // L列 機械
    contents: colValues_(master, 14, 5),               // N列 作業内容
    kouTei: colValues_(master, 3, 5),                  // C列 工程
    kouShu: kouShu,
    todayEntries: readEntries_(todayYmd_(), 'all'),
    ledgerUrl: ss.getUrl()
  };
}

/** メールアドレス→作業員名（設定シートの「メールアドレス／作業員名」表を優先） */
var USER_MAP_CACHE_ = null;
function guessUserName_(email) {
  var e = String(email || '').trim().toLowerCase();
  if (!e) return '';
  if (USER_MAP_CACHE_ === null) {
    USER_MAP_CACHE_ = {};
    try {
      var cfg = ss_().getSheetByName(SHEET_CONFIG);
      if (cfg) cfg.getRange(1, 1, cfg.getLastRow(), 2).getValues().forEach(function (r) {
        var a = String(r[0]).trim().toLowerCase(), b = String(r[1]).trim();
        if (a.indexOf('@') > 0 && b) USER_MAP_CACHE_[a] = b;
      });
    } catch (err) { /* 設定シートが無くても動く */ }
  }
  if (USER_MAP_CACHE_[e]) return USER_MAP_CACHE_[e];
  var local = e.split('@')[0];
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
      ledger.getRange(tr, 1).setValue(toSheetDate_(payload.date));               // A 日付（シートの時計で 00:00）
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
             entries: readEntries_(payload.date, 'all') };
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
  return readEntries_(String(dateStr), scope || 'all');
}

function readEntries_(ymd, scope) {
  var ledger = ss_().getSheetByName(SHEET_LEDGER);
  var last = ledger.getLastRow();
  if (last < 2) return [];
  var vals = ledger.getRange(2, 1, last - 1, 17).getValues();
  var email = Session.getActiveUser().getEmail();
  var myName = guessUserName_(email);
  var key = String(ymd);
  var out = [];
  vals.forEach(function (r) {
    if (!(r[0] instanceof Date)) return;
    if (dateOnly_(r[0]) !== key) return;
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
  // 台帳の最大日付（'yyyy-MM-dd'）
  var vals = ss.getSheetByName(SHEET_LEDGER)
    .getRange(2, 1, LEDGER_MAX_ROW - 1, 1).getValues();
  var maxYmd = '';
  for (var j = 0; j < vals.length; j++) {
    var v = vals[j][0];
    if (v instanceof Date) { var y = dateOnly_(v); if (y > maxYmd) maxYmd = y; }
  }
  if (!maxYmd) return 0;

  // 4シートをそれぞれ独立に延ばす（途中で止まっても再実行で残りが直る）
  var targets = [['日別集計', true], ['工種別日別', true],
                 ['印刷01_日別人工', false], ['印刷02_累計推移', false]];
  var added = 0;
  targets.forEach(function (t) {
    var sh = ss.getSheetByName(t[0]);
    if (!sh) return;
    var last = lastDateRow_(sh);
    if (!last.row) return;
    var need = daysBetween_(last.ymd, maxYmd);
    if (need > 400) need = 400;          // 日付の打ち間違い対策
    if (need > 0) { extendDateSheet_(sh, last.row, need, t[1]); added = Math.max(added, need); }
    if (t[1]) fixTotalRow_(sh, last.row + Math.max(need, 0));
  });
  return added;
}

/** A列の一番下の日付行を返す { row, ymd } */
function lastDateRow_(sh) {
  var colA = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (var i = colA.length - 1; i >= 0; i--) {
    if (colA[i][0] instanceof Date) return { row: i + 1, ymd: dateOnly_(colA[i][0]) };
  }
  return { row: 0, ymd: '' };
}

/** アプリが書いた行の日付に時刻が混ざっていたら 00:00 に直す（表示されている日付は変えない） */
function normalizeLedgerDates_(ss) {
  var ledger = ss.getSheetByName(SHEET_LEDGER);
  var last = ledger.getLastRow();
  if (last < 2) return 0;
  var a = ledger.getRange(2, 1, last - 1, 1).getValues();
  var o = ledger.getRange(2, 15, last - 1, 1).getValues();     // O列 入力者
  var fixed = 0;
  for (var i = 0; i < a.length; i++) {
    var v = a[i][0];
    if (!(v instanceof Date) || String(o[i][0]).trim() === '') continue;
    var want = toSheetDate_(dateOnly_(v));
    if (want.getTime() !== v.getTime()) { ledger.getRange(i + 2, 1).setValue(want); fixed++; }
  }
  return fixed;
}

/** 合計行（最終日付行の直下）の SUM の範囲を 5行目〜最終日付行 に合わせ直す */
function fixTotalRow_(sh, lastDataRow) {
  var totalRow = 0;
  for (var r = lastDataRow + 1; r <= lastDataRow + 3 && r <= sh.getLastRow(); r++) {
    if (String(sh.getRange(r, 1).getValue()).indexOf('合計') >= 0) { totalRow = r; break; }
  }
  if (!totalRow) return;
  var lastCol = sh.getLastColumn();
  var formulas = sh.getRange(totalRow, 1, 1, lastCol).getFormulas()[0];
  for (var c = 0; c < formulas.length; c++) {
    if (/^=SUM\(/i.test(formulas[c])) {
      var letter = colLetter_(c + 1);
      var want = '=SUM(' + letter + '5:' + letter + lastDataRow + ')';
      if (formulas[c] !== want) sh.getRange(totalRow, c + 1).setFormula(want);
    }
  }
}

function dayOnly_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime(); }

function extendDateSheet_(sh, lastDateRow, need, fillDate) {
  if (!sh) return;
  var lastCol = sh.getLastColumn();
  sh.insertRowsAfter(lastDateRow, need);
  sh.getRange(lastDateRow, 1, 1, lastCol)
    .copyTo(sh.getRange(lastDateRow + 1, 1, need, lastCol));
  if (!fillDate) return;

  // 日付を1日ずつ入れる（シートの時計で 00:00）
  var baseYmd = dateOnly_(sh.getRange(lastDateRow, 1).getValue());
  var out = [];
  for (var i = 1; i <= need; i++) out.push([toSheetDate_(addDays_(baseYmd, i))]);
  sh.getRange(lastDateRow + 1, 1, need, 1).setValues(out);
}

function colLetter_(n) {
  var s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
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

  var fixedDates = normalizeLedgerDates_(ss);
  if (fixedDates) log.push('台帳の日付 ' + fixedDates + ' 行から余分な時刻を取り除きました（日別集計に数えられるようになります）');

  var n = ensureDateRows_(ss);
  if (n) log.push('集計シートに ' + n + ' 日分の行を追加しました');

  if (ensureUserTable_(ss)) log.push('設定シートに「社員の割り当て（メールアドレス → 作業員名）」の表を追加しました。各自のメールアドレスを入れてください');

  var msg = log.length ? '✅ ' + log.join('\n✅ ') : '修正の必要はありませんでした。';
  Logger.log(msg);
  return msg;
}

/** 設定シートに社員（メール→作業員名）の表が無ければ末尾に追加する。追加したら true */
function ensureUserTable_(ss) {
  var cfg = ss.getSheetByName(SHEET_CONFIG);
  if (!cfg) return false;
  var vals = cfg.getRange(1, 1, cfg.getLastRow(), 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).indexOf('社員の割り当て') >= 0) return false;
  }
  var start = cfg.getLastRow() + 2;
  var rows = [
    ['社員の割り当て（メールアドレス → 作業員名。アプリの名前表示と「じぶんの履歴」に使います）', '', ''],
    ['メールアドレス', '作業員名', '備考'],
    ['tomoyose@earth-kensetsu.jp', '友寄', ''],
    ['', '荷川取', '← 会社のメールアドレスを入れてください'],
    ['', '當山', '← 同上'],
    ['', '宮城', '← 同上'],
    ['', '上原', '← 同上'],
    ['', '山里', '← 同上'],
    ['', '平良', '← 同上']
  ];
  cfg.getRange(start, 1, rows.length, 3).setValues(rows);
  cfg.getRange(start + 1, 1, 1, 3).setFontWeight('bold');
  return true;
}

/* =====================================================================
 *  本人の名前登録（初回にアプリが聞く → 設定シートの社員表に自動保存）
 * ===================================================================== */
function api_setMyName(name) {
  name = String(name || '').trim();
  if (!name) throw new Error('名前が空です。');
  var email = String(Session.getActiveUser().getEmail() || '').trim();
  if (!email) throw new Error('ログイン情報が取得できませんでした。');
  var ss = ss_();
  ensureUserTable_(ss);
  var cfg = ss.getSheetByName(SHEET_CONFIG);
  var vals = cfg.getRange(1, 1, cfg.getLastRow(), 2).getValues();
  var done = false;
  for (var i = 0; i < vals.length && !done; i++) {                 // 既にメールがある行
    if (String(vals[i][0]).trim().toLowerCase() === email.toLowerCase()) {
      cfg.getRange(i + 1, 2).setValue(name); done = true;
    }
  }
  for (var j = 0; j < vals.length && !done; j++) {                 // 名前だけの空き行
    if (String(vals[j][0]).trim() === '' && String(vals[j][1]).trim() === name) {
      cfg.getRange(j + 1, 1, 1, 3).setValues([[email, name, 'アプリで登録']]); done = true;
    }
  }
  if (!done) cfg.getRange(cfg.getLastRow() + 1, 1, 1, 3).setValues([[email, name, 'アプリで登録']]);
  USER_MAP_CACHE_ = null;
  return name;
}

/* =====================================================================
 *  更新（自己更新）: GitHub に置いた最新版を取り込み、新バージョンとして公開し、
 *  マスター修正まで自動で行う。以後、コードの貼り付けは不要。
 *   事前に1回だけ: https://script.google.com/home/usersettings で
 *   「Google Apps Script API」を ON にする
 * ===================================================================== */
var RELEASE_BASE = 'https://raw.githubusercontent.com/tomo-earth/takken-quiz/claude/genba-macro-phase1-2hv0lc/genba-nippo/app/release/';

function 更新() {
  var log = [];
  var version = fetchText_(RELEASE_BASE + 'version.txt').trim();
  var current = PropertiesService.getScriptProperties().getProperty('APP_VERSION') || '(初回)';
  if (version === current) {
    log.push('すでに最新版（' + version + '）です。');
  } else {
    var source = fetchText_(RELEASE_BASE + 'bundle.gs');
    var manifest = fetchText_(RELEASE_BASE + 'appsscript.json');
    var scriptId = ScriptApp.getScriptId();
    var base = 'https://script.googleapis.com/v1/projects/' + scriptId;

    apiCall_('PUT', base + '/content', {
      files: [{ name: 'コード', type: 'SERVER_JS', source: source },
              { name: 'appsscript', type: 'JSON', source: manifest }]
    });
    log.push('コードを ' + version + ' に更新しました');

    var ver = apiCall_('POST', base + '/versions', { description: '現場日報 ' + version });
    var list = apiCall_('GET', base + '/deployments', null);
    var target = null;
    (list.deployments || []).forEach(function (d) {
      var isWeb = (d.entryPoints || []).some(function (e) { return e.entryPointType === 'WEB_APP'; });
      if (isWeb && d.deploymentConfig && d.deploymentConfig.versionNumber) target = d;
    });
    if (target) {
      apiCall_('PUT', base + '/deployments/' + target.deploymentId, {
        deploymentConfig: { scriptId: scriptId, versionNumber: ver.versionNumber,
                            manifestFileName: 'appsscript', description: '現場日報 ' + version }
      });
      log.push('アプリを新バージョン（' + ver.versionNumber + '）で公開しました。URL は変わりません');
    } else {
      log.push('公開中のウェブアプリが見つからないため、公開の更新は手動で行ってください（デプロイ → デプロイを管理）');
    }
    PropertiesService.getScriptProperties().setProperty('APP_VERSION', version);
  }
  try { log.push(マスター修正()); } catch (e) { log.push('マスター修正でエラー: ' + e.message); }
  var msg = log.join('\n');
  Logger.log(msg);
  return msg;
}

/** 毎日1回、自動で「更新」を実行する（止めるときは 自動更新をOFF を実行） */
function 自動更新をON() {
  自動更新をOFF();
  ScriptApp.newTrigger('更新').timeBased().everyDays(1).atHour(5).create();
  Logger.log('✅ 毎日5時ごろに自動更新します。');
}
function 自動更新をOFF() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === '更新') ScriptApp.deleteTrigger(t);
  });
  Logger.log('自動更新を止めました。');
}

function fetchText_(url) {
  var res = UrlFetchApp.fetch(url + '?t=' + Date.now(), { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('取得に失敗: ' + url + ' (' + res.getResponseCode() + ')');
  return res.getContentText('UTF-8');
}

function apiCall_(method, url, body) {
  var res = UrlFetchApp.fetch(url, {
    method: method, contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: body ? JSON.stringify(body) : null, muteHttpExceptions: true
  });
  var codeNum = res.getResponseCode();
  if (codeNum >= 300) {
    var body = res.getContentText();
    var hint = '';
    if (body.indexOf('insufficient authentication scopes') >= 0) {
      hint = '\n→ 新しい権限の許可がまだです。appsscript.json を保存したうえで、もう一度「更新」を▶実行し、'
           + '「承認が必要です」→ 許可 と進んでください。許可画面が出ない場合は '
           + 'https://myaccount.google.com/permissions で「無題のプロジェクト（または現場日報アプリ）」のアクセスを削除してから再実行してください。';
    } else if (codeNum === 403 || codeNum === 404) {
      hint = '\n→ https://script.google.com/home/usersettings で「Google Apps Script API」を ON にしてから、もう一度「更新」を実行してください。';
    }
    throw new Error('Apps Script API エラー ' + codeNum + ': ' + body.slice(0, 300) + hint);
  }
  return res.getContentText() ? JSON.parse(res.getContentText()) : {};
}
