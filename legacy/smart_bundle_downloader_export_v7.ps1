param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$RestArgs
)

$OutputEncoding = [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$env:PYTHONUTF8 = "1"
$env:PYTHONUNBUFFERED = "1"

# 适配 V7：标题/摘要 + 20 条线程进度 + 菜单，目标约 132x44。
try {
    $raw = $Host.UI.RawUI
    $targetWidth = 132
    $targetHeight = 44

    try {
        $maxSize = $raw.MaxPhysicalWindowSize
        if ($maxSize.Width -gt 0)  { $targetWidth  = [Math]::Min($targetWidth,  $maxSize.Width) }
        if ($maxSize.Height -gt 0) { $targetHeight = [Math]::Min($targetHeight, $maxSize.Height) }
    } catch {}

    if ($targetWidth -lt 96) { $targetWidth = 96 }
    if ($targetHeight -lt 32) { $targetHeight = 32 }

    $bufferWidth = [Math]::Max($raw.BufferSize.Width, $targetWidth)
    $bufferHeight = [Math]::Max($raw.BufferSize.Height, 1000)
    $raw.BufferSize = New-Object System.Management.Automation.Host.Size($bufferWidth, $bufferHeight)
    $raw.WindowSize = New-Object System.Management.Automation.Host.Size($targetWidth, $targetHeight)
    $raw.WindowTitle = "Unity Addressables Bundle Downloader / Export V7"
} catch {
    # Windows Terminal 等宿主可能禁止调整 WindowSize，忽略即可。
}

Clear-Host
$script = Join-Path $PSScriptRoot "smart_bundle_downloader_export_v7.py"
$py = Get-Command py -ErrorAction SilentlyContinue
if ($py) {
    & py -3 -u $script @RestArgs
} else {
    $python = Get-Command py -ErrorAction SilentlyContinue
if ($python) {
    & $python.Source -3 -u $script @RestArgs
} else {
    $python = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($python) {
        & $python.Source -u $script @RestArgs
    } else {
        $candidates = @(
            "$env:LOCALAPPDATA\Programs\Python\Python313\python.exe",
            "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe",
            "$env:LOCALAPPDATA\Programs\Python\Python311\python.exe",
            "$env:LOCALAPPDATA\Programs\Python\Python310\python.exe",
            "$env:ProgramFiles\Python313\python.exe",
            "$env:ProgramFiles\Python312\python.exe",
            "$env:ProgramFiles\Python311\python.exe",
            "$env:ProgramFiles\Python310\python.exe"
        )
        $pythonPath = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
        if (-not $pythonPath) {
            Write-Host "ERROR: Python was not found. Please install Python 3.10+ or add python.exe to PATH."
            exit 1
        }
        & $pythonPath -u $script @RestArgs
    }
}
}
$code = $LASTEXITCODE

if ($code -ne 0) {
    Write-Host ""
    Write-Host "程序退出码: $code" -ForegroundColor Yellow
    Read-Host "按 Enter 关闭窗口" | Out-Null
}
exit $code
