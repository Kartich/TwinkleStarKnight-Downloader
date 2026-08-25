param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$RestArgs
)

$OutputEncoding = [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$env:PYTHONUTF8 = "1"
$env:PYTHONUNBUFFERED = "1"

try {
    $Host.UI.RawUI.WindowTitle = "Unity Addressables Bundle Downloader V5"
} catch {}

Clear-Host
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " Unity Addressables Bundle Downloader / Export V5" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "修复 Windows Export 状态文件 WinError 5 / PPtr raw 导出失败。" -ForegroundColor DarkGray
Write-Host "下载与 Export 状态固定行实时刷新；完成后自动清理临时工作文件。" -ForegroundColor DarkGray
Write-Host ""

$script = Join-Path $PSScriptRoot "smart_bundle_downloader_export_v5.py"

$py = Get-Command py -ErrorAction SilentlyContinue
if ($py) {
    & py -3 -u $script @RestArgs
} else {
    & python -u $script @RestArgs
}
$code = $LASTEXITCODE

Write-Host ""
if ($code -eq 0) {
    Write-Host "完成。" -ForegroundColor Green
} else {
    Write-Host "程序退出码: $code" -ForegroundColor Yellow
}

Read-Host "按 Enter 关闭窗口" | Out-Null
exit $code
