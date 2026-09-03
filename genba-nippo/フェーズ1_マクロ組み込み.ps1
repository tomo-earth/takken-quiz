# フェーズ1（予備手段）: xlsx に VBA を組み込み .xlsm を作る PowerShell スクリプト
# ※ 通常は同梱の「谷直美様邸_現場日報台帳_集計.xlsm」をそのまま使えます。
#    もし xlsm が Excel で開けない・修復扱いになる場合のみ、このスクリプトで作り直してください。
#
# 事前準備（1回だけ・手動）:
#   Excel → ファイル → オプション → トラストセンター → トラストセンターの設定
#   → マクロの設定 → 「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を ON
#
# 実行方法:
#   このスクリプトと 谷直美様邸_現場日報台帳_集計.xlsx と 印刷メニュー_マクロ.bas を
#   同じフォルダに置き、右クリック →「PowerShell で実行」
#   （または PowerShell で:  .\フェーズ1_マクロ組み込み.ps1 ）

$ErrorActionPreference = "Stop"
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$src = Join-Path $dir "谷直美様邸_現場日報台帳_集計.xlsx"
$bas = Join-Path $dir "印刷メニュー_マクロ.bas"
$dst = Join-Path $dir "谷直美様邸_現場日報台帳_集計.xlsm"

if (-not (Test-Path $src)) { Write-Error "見つかりません: $src" }
if (-not (Test-Path $bas)) { Write-Error "見つかりません: $bas" }
if (Test-Path $dst) {
    $bak = Join-Path $dir ("谷直美様邸_現場日報台帳_集計_旧" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".xlsm")
    Move-Item $dst $bak
    Write-Host "既存の xlsm を退避しました: $bak"
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($src)
    try {
        $wb.VBProject.VBComponents.Import($bas) | Out-Null
    } catch {
        throw ("VBA の取り込みに失敗しました。トラストセンターで " +
               "「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を ON にしてから再実行してください。`n" + $_)
    }
    $wb.SaveAs($dst, 52)             # 52 = xlOpenXMLWorkbookMacroEnabled
    $xl.Run("印刷_ボタンを配置")      # 印刷メニューシート G 列にボタン 6 個を配置
    $wb.Save()
    $wb.Close($false)
    Write-Host "完成: $dst"
    Write-Host "元の xlsx はそのまま残っています: $src"
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
Read-Host "Enter キーで閉じます"
