# Release ビルド → DLL署名 → インストーラー作成 → インストーラー署名 を一括で行う。
# 使い方:  powershell -ExecutionPolicy Bypass -File build-installer.ps1   （-NoSign で署名なし）
param([switch]$NoSign)
$ErrorActionPreference = "Stop"

$here    = Split-Path -Parent $MyInvocation.MyCommand.Path     # installer フォルダ
$desktop = Split-Path -Parent $here                            # ShapePaletteDesktop
$msbuild = "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe"
$iscc    = "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
$sln     = Join-Path $desktop "ShapePalette.sln"
$relDir  = Join-Path $desktop "ShapePalette\bin\Release"
$ts      = "http://timestamp.digicert.com"
$subject = "CN=Shape Palette (Yuki)"

Write-Host "[1] Release ビルド..." -ForegroundColor Cyan
& $msbuild $sln /t:Restore  /p:Configuration=Release /p:Platform="Any CPU" /v:quiet   /nologo
& $msbuild $sln /t:Rebuild  /p:Configuration=Release /p:Platform="Any CPU" /v:minimal /nologo
if ($LASTEXITCODE -ne 0) { throw "ビルド失敗" }

$cert = $null
if (-not $NoSign) {
    $cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert -ErrorAction SilentlyContinue |
            Where-Object { $_.Subject -eq $subject } | Select-Object -First 1
    if (-not $cert) {
        Write-Host "[*] 自己署名証明書を新規作成" -ForegroundColor Yellow
        $cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject $subject `
            -CertStoreLocation Cert:\CurrentUser\My -KeyUsage DigitalSignature `
            -KeyExportPolicy Exportable -NotAfter (Get-Date).AddYears(5)
        # このPCで信頼させる（他PCでは ShapePalette-PublicCert.cer を信頼ストアに入れる）
        foreach ($s in @("Root","TrustedPublisher")) {
            $st = New-Object System.Security.Cryptography.X509Certificates.X509Store($s,"CurrentUser")
            $st.Open("ReadWrite"); $st.Add($cert); $st.Close()
        }
        Export-Certificate -Cert $cert -FilePath (Join-Path $here "ShapePalette-PublicCert.cer") -Type CERT | Out-Null
    }
    Write-Host "[2] DLL 署名..." -ForegroundColor Cyan
    Set-AuthenticodeSignature -FilePath (Join-Path $relDir "ShapePalette.dll") -Certificate $cert -TimeStampServer $ts -HashAlgorithm SHA256 | Out-Null
}

Write-Host "[3] インストーラー作成..." -ForegroundColor Cyan
& $iscc (Join-Path $here "ShapePalette.iss")
if ($LASTEXITCODE -ne 0) { throw "ISCC 失敗" }

$setup = Join-Path $here "Output\ShapePaletteSetup.exe"
if ($cert) {
    Write-Host "[4] インストーラー署名..." -ForegroundColor Cyan
    Set-AuthenticodeSignature -FilePath $setup -Certificate $cert -TimeStampServer $ts -HashAlgorithm SHA256 | Out-Null
    Write-Host ("    署名状態: " + (Get-AuthenticodeSignature $setup).Status)
}
Write-Host "完成: $setup" -ForegroundColor Green
