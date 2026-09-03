# -*- coding: utf-8 -*-
"""コード.gs と index.html を1ファイルの Apps Script（全部入り）にまとめる。
- HTML は文字列として埋め込み（HtmlService.createHtmlOutput）
- 台帳の ID は「セットアップ」実行時に自動取得してスクリプトプロパティへ保存（手書き不要）
- セットアップ: 07.日報フォルダ内の Excel(台帳) を Google スプレッドシートへ変換し、
  台帳 M〜Q 列の見出しと「設定」シートを自動で用意する
出力: 現場日報アプリ_全部入り.gs
"""
import re

FOLDER_ID = "1dRBCpHN1mKeTslkhLNfoMzyIpUm4-cem"        # 07.日報フォルダ
PHOTO_ROOT_ID = "1kFKgVW8pRDY2O-SEPVaAfC-oZYbiP_fi"    # 00.工事部
PHOTO_UNKNOWN_ID = "18tJl4_plrS1A53U5VsdbODN5--_-ODs0" # 99_写真置き場
TANI_FOLDER_ID = "14NO0tzYYFYRNyNAAuvV8bpyaTBQ1rLKI"   # 01_谷直美様邸

code = open("コード.gs", encoding="utf-8").read()
html = open("index.html", encoding="utf-8").read()
assert "`" not in html and "${" not in html, "HTML にバッククォート/${ が含まれています"

# 1) SPREADSHEET_ID 定数 → プロパティ/自動検出に置換
code = code.replace(
    "// ★★ 初期設定: 台帳スプレッドシートの ID を貼り付けてください ★★\n"
    "// （スプレッドシートの URL の /d/ と /edit の間の文字列）\n"
    "var SPREADSHEET_ID = 'ここに台帳スプレッドシートのIDを貼り付け';\n",
    "// 台帳スプレッドシートの ID は「セットアップ」実行時に自動で保存されます（手入力不要）\n"
    f"var FOLDER_ID = '{FOLDER_ID}';            // 07.日報フォルダ\n"
    "var LEDGER_NAME = '現場日報_台帳（アプリ用）';\n")
code = code.replace(
    "function ss_() { return SpreadsheetApp.openById(SPREADSHEET_ID); }",
    "function ss_() { return SpreadsheetApp.openById(ledgerId_()); }\n\n"
    "function ledgerId_() {\n"
    "  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');\n"
    "  if (id) return id;\n"
    "  var it = DriveApp.getFolderById(FOLDER_ID).getFilesByName(LEDGER_NAME);\n"
    "  if (it.hasNext()) {\n"
    "    id = it.next().getId();\n"
    "    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', id);\n"
    "    return id;\n"
    "  }\n"
    "  throw new Error('台帳がまだ作られていません。Apps Script の画面で「セットアップ」を1回実行してください。');\n"
    "}")
# 2) doGet: ファイル参照 → 埋め込みHTML
code = code.replace(
    "return HtmlService.createTemplateFromFile('index').evaluate()",
    "return HtmlService.createHtmlOutput(APP_HTML)")

setup = f'''
/* =====================================================================
 *  セットアップ（初回に1回だけ、Apps Script の画面から ▶ 実行する）
 *   1. 07.日報フォルダ内の台帳 Excel（.xlsm/.xlsx）を Google スプレッドシートに変換
 *   2. 台帳シートに M〜Q 列（開始時刻・終了時刻・入力者・送信日時・写真）の見出しを追加
 *   3. 「設定」シート（写真フォルダの割り当て）を作成
 *   4. 台帳の ID をこのスクリプトに記憶（以後の手入力は不要）
 * ===================================================================== */
function セットアップ() {{
  var folder = DriveApp.getFolderById(FOLDER_ID);
  var ssId = null;
  var it = folder.getFilesByName(LEDGER_NAME);
  if (it.hasNext()) {{
    ssId = it.next().getId();
    Logger.log('既存の台帳を使います: ' + LEDGER_NAME);
  }} else {{
    var src = pickSourceFile_(folder);
    if (!src) throw new Error('07.日報フォルダ に台帳の Excel ファイル（現場日報台帳）が見つかりません。');
    Logger.log('変換元: ' + src.getName());
    var res = UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' + src.getId() + '/copy?supportsAllDrives=true', {{
        method: 'post', contentType: 'application/json',
        headers: {{ Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }},
        payload: JSON.stringify({{ name: LEDGER_NAME,
          mimeType: 'application/vnd.google-apps.spreadsheet', parents: [FOLDER_ID] }}),
        muteHttpExceptions: true }});
    if (res.getResponseCode() >= 300) throw new Error('スプレッドシートへの変換に失敗しました: ' + res.getContentText());
    ssId = JSON.parse(res.getContentText()).id;
  }}
  var ss = SpreadsheetApp.openById(ssId);
  ensureLedgerColumns_(ss);
  ensureConfigSheet_(ss);
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ssId);
  var msg = '✅ セットアップ完了。台帳: ' + ss.getUrl() +
    '\\n次は右上「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」で公開してください。';
  Logger.log(msg);
  return msg;
}}

function pickSourceFile_(folder) {{
  var best = null, bestScore = 0;
  var files = folder.getFiles();
  while (files.hasNext()) {{
    var f = files.next();
    var name = f.getName(), mt = f.getMimeType();
    var isExcel = /spreadsheetml|ms-excel/.test(mt) || /\\.xls[xm]$/i.test(name);
    if (!isExcel) continue;
    var score = 1;
    if (name.indexOf('クラウド版') >= 0) score += 2;
    if (name.indexOf('日報') >= 0) score += 1;
    if (score > bestScore) {{ best = f; bestScore = score; }}
  }}
  return best;
}}

function ensureLedgerColumns_(ss) {{
  var sh = ss.getSheetByName(SHEET_LEDGER);
  if (!sh) throw new Error('「台帳」シートが見つかりません。');
  if (sh.getMaxColumns() < 17) sh.insertColumnsAfter(sh.getMaxColumns(), 17 - sh.getMaxColumns());
  if (String(sh.getRange('M1').getValue()).trim() === '') {{
    sh.getRange('M1:Q1').setValues([['開始時刻', '終了時刻', '入力者', '送信日時', '写真']]);
    sh.getRange('L1').copyTo(sh.getRange('M1:Q1'), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  }}
}}

function ensureConfigSheet_(ss) {{
  if (ss.getSheetByName(SHEET_CONFIG)) return;
  var sh = ss.insertSheet(SHEET_CONFIG);
  var rows = [
    ['現場日報アプリ 設定シート（この表の場所・見出しは変えないでください）', '', ''],
    ['', '', ''],
    ['項目', '値', '説明'],
    ['写真ルートフォルダID', '{PHOTO_ROOT_ID}', '共有ドライブ「アース建設03案件」内の 00.工事部 フォルダ'],
    ['写真置き場フォルダID', '{PHOTO_UNKNOWN_ID}', '現場が特定できない写真の保存先（99_写真置き場）'],
    ['写真サブフォルダ名', '06_写真', '各現場フォルダ内の写真フォルダの名前'],
    ['日報写真フォルダ名', '日報', '06_写真 の下に自動で作る日報写真用フォルダ（月別フォルダも自動作成）'],
    ['', '', ''],
    ['現場フォルダの割り当て（現場名 → フォルダID。行を足せば現場を増やせます）', '', ''],
    ['現場名', '現場フォルダID', '備考'],
    ['谷直美様住宅新築工事', '{TANI_FOLDER_ID}', '01_工事No　谷直美様住宅新築工事']
  ];
  sh.getRange(1, 1, rows.length, 3).setValues(rows);
  sh.getRange('A3:C3').setFontWeight('bold'); sh.getRange('A10:C10').setFontWeight('bold');
  sh.setColumnWidth(1, 260); sh.setColumnWidth(2, 320); sh.setColumnWidth(3, 460);
}}
'''

bundle = (code.rstrip() + "\n" + setup +
          "\n/* ===== 画面（index.html をそのまま埋め込み） ===== */\n"
          "var APP_HTML = `" + html + "`;\n")
open("現場日報アプリ_全部入り.gs", "w", encoding="utf-8").write(bundle)
print("wrote 現場日報アプリ_全部入り.gs", len(bundle), "bytes")
