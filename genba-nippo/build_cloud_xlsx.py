# -*- coding: utf-8 -*-
"""クラウド（Googleスプレッドシート）版の台帳ブックを生成する。

元の xlsx から zip 手術で:
 1. 台帳シートに M〜Q 列の見出し（開始時刻/終了時刻/入力者/送信日時/写真）を追加
 2. アプリ用「設定」シート（写真フォルダの ID・現場フォルダの割り当て）を追加
既存の数式・入力規則・書式はバイト単位で保持する（openpyxl は使わない）。

使い方: python3 build_cloud_xlsx.py <入力xlsx> <出力xlsx>
"""
import re
import sys
import zipfile

PHOTO_ROOT_ID = "1kFKgVW8pRDY2O-SEPVaAfC-oZYbiP_fi"        # 共有ドライブ 00.工事部
PHOTO_UNKNOWN_ID = "18tJl4_plrS1A53U5VsdbODN5--_-ODs0"     # 99_写真置き場（現場不明）
TANI_FOLDER_ID = "14NO0tzYYFYRNyNAAuvV8bpyaTBQ1rLKI"       # 01_工事No 谷直美様住宅新築工事

NEW_HEADERS = ["開始時刻", "終了時刻", "入力者", "送信日時", "写真"]  # M1..Q1

SETTINGS_ROWS = [
    ["現場日報アプリ 設定シート（この表の場所・見出しは変えないでください）", "", ""],
    ["", "", ""],
    ["項目", "値", "説明"],
    ["写真ルートフォルダID", PHOTO_ROOT_ID, "共有ドライブ「アース建設03案件」内の 00.工事部 フォルダ"],
    ["写真置き場フォルダID", PHOTO_UNKNOWN_ID, "現場が特定できない写真の保存先（99_写真置き場）"],
    ["写真サブフォルダ名", "06_写真", "各現場フォルダ内の写真フォルダの名前"],
    ["日報写真フォルダ名", "日報", "06_写真 の下に自動で作る日報写真用フォルダ（さらに 2026-09 のような月別フォルダを自動作成）"],
    ["", "", ""],
    ["現場フォルダの割り当て（現場名 → フォルダID。行を足せば現場を増やせます）", "", ""],
    ["現場名", "現場フォルダID", "備考"],
    ["谷直美様住宅新築工事", TANI_FOLDER_ID, "01_工事No　谷直美様住宅新築工事"],
]


def esc(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
             .replace('"', "&quot;"))


def build_settings_sheet_xml():
    rows = []
    for r, row in enumerate(SETTINGS_ROWS, 1):
        cells = []
        for c, val in enumerate(row):
            if val == "":
                continue
            ref = chr(ord("A") + c) + str(r)
            cells.append(f'<c r="{ref}" t="inlineStr"><is><t>{esc(val)}</t></is></c>')
        rows.append(f'<row r="{r}">{"".join(cells)}</row>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<dimension ref="A1:C11"/>'
        '<cols><col min="1" max="1" width="34" customWidth="1"/>'
        '<col min="2" max="2" width="42" customWidth="1"/>'
        '<col min="3" max="3" width="60" customWidth="1"/></cols>'
        f'<sheetData>{"".join(rows)}</sheetData>'
        "</worksheet>"
    )


def build(src, dst):
    zin = zipfile.ZipFile(src)
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        name = item.filename
        if name == "xl/worksheets/sheet3.xml":            # 台帳
            text = data.decode("utf-8")
            text = text.replace('<dimension ref="A1:L2000"/>',
                                '<dimension ref="A1:Q2000"/>')
            extra = "".join(
                f'<c r="{col}1" s="5" t="inlineStr"><is><t>{esc(h)}</t></is></c>'
                for col, h in zip("MNOPQ", NEW_HEADERS))
            text = text.replace(
                '<row r="1" spans="1:12" ht="28.5" x14ac:dyDescent="0.2">',
                '<row r="1" spans="1:17" ht="28.5" x14ac:dyDescent="0.2">')
            # 行1の閉じタグ直前に新セルを挿入
            m = re.search(r'(<row r="1"[^>]*>)(.*?)(</row>)', text)
            text = text[:m.start()] + m.group(1) + m.group(2) + extra + m.group(3) + text[m.end():]
            data = text.encode("utf-8")
        elif name == "xl/workbook.xml":
            text = data.decode("utf-8")
            text = text.replace(
                "</sheets>",
                '<sheet name="設定" sheetId="14" r:id="rId18"/></sheets>')
            data = text.encode("utf-8")
        elif name == "xl/_rels/workbook.xml.rels":
            text = data.decode("utf-8")
            text = text.replace(
                "</Relationships>",
                '<Relationship Id="rId18" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                'Target="worksheets/sheet14.xml"/></Relationships>')
            data = text.encode("utf-8")
        elif name == "[Content_Types].xml":
            text = data.decode("utf-8")
            text = text.replace(
                "</Types>",
                '<Override PartName="/xl/worksheets/sheet14.xml" '
                'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
            data = text.encode("utf-8")
        zi = zipfile.ZipInfo(name, date_time=item.date_time)
        zi.compress_type = zipfile.ZIP_DEFLATED
        zout.writestr(zi, data)
    zout.writestr("xl/worksheets/sheet14.xml", build_settings_sheet_xml())
    zout.close()
    zin.close()
    print(f"wrote {dst}")


if __name__ == "__main__":
    build(sys.argv[1] if len(sys.argv) > 1 else "谷直美様邸_現場日報台帳_集計.xlsx",
          sys.argv[2] if len(sys.argv) > 2 else "現場日報_台帳_クラウド版.xlsx")
