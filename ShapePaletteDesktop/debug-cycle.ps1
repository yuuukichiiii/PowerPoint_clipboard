# ビルド→（必要なら登録）→PowerPoint起動→ログ確認 を自動化するデバッグ用スクリプト。
param(
    [switch]$Register,
    [switch]$Restore,
    [switch]$SelfTest,
    [int]$WaitSeconds = 12,
    [switch]$NoLaunch
)
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$msbuild = "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe"
$sln = Join-Path $root "ShapePalette.sln"
$log = Join-Path $env:TEMP "ShapePalette.log"

Write-Host "[1] PowerPoint を穏やかに終了..." -ForegroundColor Cyan
# graceful close（未保存の空プレゼンならプロンプトなしで閉じる）。これで「異常終了」フラグを残さない。
& taskkill /IM POWERPNT.EXE 2>$null | Out-Null
Start-Sleep -Seconds 3
# まだ残っていれば強制終了（最後の手段）
$still = Get-Process POWERPNT -ErrorAction SilentlyContinue
if ($still) { $still | Stop-Process -Force -ErrorAction SilentlyContinue; Start-Sleep -Seconds 2 }
# 異常終了に伴うセーフモード/復元/無効化フラグを丸ごと掃除（クラッシュ汚染の連鎖を断つ）
Remove-Item "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Resiliency" -Recurse -Force -ErrorAction SilentlyContinue
# LoadBehavior を必ず 3 に戻す（前回クラッシュで 2 に落とされる場合がある）
$ak = "HKCU:\Software\Microsoft\Office\PowerPoint\AddIns\ShapePalette.Connect"
if (Test-Path $ak) { Set-ItemProperty $ak -Name LoadBehavior -Value 3 -Type DWord }

if ($Restore) {
    Write-Host "[2a] 復元..." -ForegroundColor Cyan
    & $msbuild $sln /t:Restore /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
}
Write-Host "[2] ビルド..." -ForegroundColor Cyan
& $msbuild $sln /t:Build /p:Configuration=Debug /p:Platform="Any CPU" /v:minimal /nologo
if ($LASTEXITCODE -ne 0) { Write-Host "ビルド失敗" -ForegroundColor Red; exit 1 }

if ($Register) {
    Write-Host "[3] 登録..." -ForegroundColor Cyan
    & (Join-Path $root "register.ps1") -Configuration Debug
}

Remove-Item $log -ErrorAction SilentlyContinue
Write-Host "[4] ログをクリア" -ForegroundColor Cyan

if (-not $NoLaunch) {
    Write-Host "[5] PowerPoint 起動..." -ForegroundColor Cyan
    if ($SelfTest) { $env:SHAPEPALETTE_SELFTEST = "1" } else { Remove-Item Env:\SHAPEPALETTE_SELFTEST -ErrorAction SilentlyContinue }
    Start-Process "powerpnt.exe" -ArgumentList "/B"
    Start-Sleep -Seconds 6
    # 起動時に出る可能性のあるダイアログ（「アドインを無効にしますか？」等）を「いいえ」で閉じる
    try {
        $ws = New-Object -ComObject WScript.Shell
        if ($ws.AppActivate("PowerPoint")) { Start-Sleep -Milliseconds 500; $ws.SendKeys("n"); Start-Sleep -Milliseconds 200; $ws.SendKeys("{ESC}") }
    } catch {}
    Write-Host "    残り待機..."
    Start-Sleep -Seconds ([Math]::Max(2, $WaitSeconds - 6))
    Write-Host "[6] ログ:" -ForegroundColor Green
    if (Test-Path $log) { Get-Content $log } else { Write-Host "(ログなし)" -ForegroundColor Yellow }
}
