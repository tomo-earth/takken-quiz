# -*- coding: utf-8 -*-
"""印刷メニュー用 VBA ソース（付録A準拠・単一ソース）。

ここから 印刷メニュー_マクロ.bas（Shift-JIS/CRLF）と
vbaProject.bin 内のモジュールストリームの両方を生成する。
注意: cp932 に無い U+301C(〜) は使わず、U+FF5E(～) を使うこと。
"""

MODULE_NAME = "印刷メニュー"

VBA_CODE = '''Attribute VB_Name = "印刷メニュー"
Option Explicit
' 「印刷メニュー」シートのB列に ○ を付けたシートをまとめて印刷する

Private Const MENU_SHEET As String = "印刷メニュー"
Private Const FIRST_ROW As Long = 6      ' シート一覧の先頭行

Private Function SelectedSheets() As Variant
    Dim ws As Worksheet, r As Long, n As Long, names() As String
    Set ws = ThisWorkbook.Worksheets(MENU_SHEET)
    r = FIRST_ROW
    Do While Trim(ws.Cells(r, 1).Value) <> "" And Left(ws.Cells(r, 1).Value, 1) <> "○"
        If Trim(ws.Cells(r, 2).Value) = "○" Then
            If SheetExists(ws.Cells(r, 1).Value) Then
                n = n + 1
                ReDim Preserve names(1 To n)
                names(n) = ws.Cells(r, 1).Value
            End If
        End If
        r = r + 1
    Loop
    If n = 0 Then SelectedSheets = Empty Else SelectedSheets = names
End Function

Private Function SheetExists(ByVal name As String) As Boolean
    Dim s As Worksheet
    On Error Resume Next
    Set s = ThisWorkbook.Worksheets(name)
    SheetExists = Not s Is Nothing
    On Error GoTo 0
End Function

Private Sub PrintSheets(ByVal names As Variant, ByVal preview As Boolean)
    Dim cur As Worksheet
    Set cur = ActiveSheet
    If IsEmpty(names) Then
        MsgBox "印刷するシートのB列に ○ を付けてください。", vbExclamation, "印刷メニュー"
        Exit Sub
    End If
    ThisWorkbook.Worksheets(names).Select
    If preview Then
        ActiveWindow.SelectedSheets.PrintPreview
    Else
        ActiveWindow.SelectedSheets.PrintOut
    End If
    cur.Select   ' グループ解除して元のシートへ戻る
End Sub

Public Sub 印刷_○のシートをプレビュー()
    PrintSheets SelectedSheets(), True
End Sub

Public Sub 印刷_○のシートを印刷()
    Dim names As Variant
    names = SelectedSheets()
    If IsEmpty(names) Then
        MsgBox "印刷するシートのB列に ○ を付けてください。", vbExclamation, "印刷メニュー"
        Exit Sub
    End If
    If MsgBox(UBound(names) & " シートを印刷します。よろしいですか？", vbOKCancel + vbQuestion, "印刷メニュー") = vbOK Then
        PrintSheets names, False
    End If
End Sub

Public Sub 印刷_印刷用シートを全部印刷()
    Dim s As Worksheet, n As Long, names() As String
    For Each s In ThisWorkbook.Worksheets
        If Left(s.Name, 2) = "印刷" And s.Name <> MENU_SHEET Then
            n = n + 1
            ReDim Preserve names(1 To n)
            names(n) = s.Name
        End If
    Next
    If n = 0 Then Exit Sub
    If MsgBox("印刷用シート " & n & " 枚組を印刷します。よろしいですか？", vbOKCancel + vbQuestion, "印刷メニュー") = vbOK Then
        PrintSheets names, False
    End If
End Sub

Public Sub 印刷_全シートを印刷()
    If MsgBox("ブック内のすべてのシートを印刷します。よろしいですか？", vbOKCancel + vbQuestion, "印刷メニュー") = vbOK Then
        ThisWorkbook.PrintOut
    End If
End Sub

Public Sub 印刷_○をすべて付ける()
    SetAllMarks "○"
End Sub

Public Sub 印刷_○をすべて外す()
    SetAllMarks ""
End Sub

Private Sub SetAllMarks(ByVal v As String)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets(MENU_SHEET)
    r = FIRST_ROW
    Do While Trim(ws.Cells(r, 1).Value) <> "" And Left(ws.Cells(r, 1).Value, 1) <> "○"
        ws.Cells(r, 2).Value = v
        r = r + 1
    Loop
End Sub

' 印刷メニューシートの G 列にボタンを配置（初回のみ実行）
Public Sub 印刷_ボタンを配置()
    Dim ws As Worksheet, btn As Object, i As Long
    Dim labels, macros
    labels = Array("○のシートをプレビュー", "○のシートを印刷", "印刷用シート(01～06)を全部印刷", "全シートを印刷", "○をすべて付ける", "○をすべて外す")
    macros = Array("印刷_○のシートをプレビュー", "印刷_○のシートを印刷", "印刷_印刷用シートを全部印刷", "印刷_全シートを印刷", "印刷_○をすべて付ける", "印刷_○をすべて外す")
    Set ws = ThisWorkbook.Worksheets(MENU_SHEET)
    For Each btn In ws.Buttons
        btn.Delete
    Next
    For i = 0 To UBound(labels)
        Set btn = ws.Buttons.Add(ws.Range("G6").Left, ws.Range("G6").Top + i * 32, 200, 26)
        btn.Caption = labels(i)
        btn.OnAction = macros(i)
        btn.Font.Size = 10
    Next
    MsgBox "印刷メニューシートのG列にボタンを配置しました。", vbInformation
End Sub
'''

# 文書モジュール（ThisWorkbook / SheetN）の空スタブ。
# VB_Base の GUID: Workbook=00020819, Worksheet=00020820
DOC_MODULE_TEMPLATE = '''Attribute VB_Name = "{name}"
Attribute VB_Base = "0{{{guid}}}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
'''

GUID_WORKBOOK = "00020819-0000-0000-C000-000000000046"
GUID_WORKSHEET = "00020820-0000-0000-C000-000000000046"


def to_vba_bytes(text: str) -> bytes:
    """cp932 + CRLF に変換"""
    return text.replace("\r\n", "\n").replace("\n", "\r\n").encode("cp932")


if __name__ == "__main__":
    import sys
    out = sys.argv[1] if len(sys.argv) > 1 else "印刷メニュー_マクロ.bas"
    with open(out, "wb") as f:
        f.write(to_vba_bytes(VBA_CODE))
    print(f"wrote {out} ({len(to_vba_bytes(VBA_CODE))} bytes, cp932/CRLF)")
