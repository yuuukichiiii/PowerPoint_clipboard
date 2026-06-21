# Shape Palette アドインを「現在のユーザーのみ」で COM 登録する（管理者不要）。
# HKCU\Software\Classes に per-user COM 登録し、PowerPoint の AddIns キーを書く。
param(
    [string]$Configuration = "Debug"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$dll = Join-Path $root "ShapePalette\bin\$Configuration\ShapePalette.dll"
if (-not (Test-Path $dll)) { throw "DLL が見つかりません: $dll （先にビルドしてください）" }

$asmName    = "ShapePalette"
$asmVersion = "1.0.0.0"
$runtime    = "v4.0.30319"
$codeBase   = "file:///" + ($dll -replace '\\','/')
$asmString  = "$asmName, Version=$asmVersion, Culture=neutral, PublicKeyToken=null"

function Register-ComClass {
    param([string]$Clsid, [string]$ProgId, [string]$ClassFullName)

    # ProgId -> CLSID
    $pidKey = "HKCU:\Software\Classes\$ProgId"
    New-Item -Path $pidKey -Force | Out-Null
    New-Item -Path "$pidKey\CLSID" -Force | Out-Null
    Set-ItemProperty -Path "$pidKey\CLSID" -Name "(default)" -Value "{$Clsid}"

    # CLSID 本体
    $clsidKey = "HKCU:\Software\Classes\CLSID\{$Clsid}"
    New-Item -Path $clsidKey -Force | Out-Null
    Set-ItemProperty -Path $clsidKey -Name "(default)" -Value $ClassFullName

    foreach ($sub in @("InprocServer32", "InprocServer32\$asmVersion")) {
        $k = "$clsidKey\$sub"
        New-Item -Path $k -Force | Out-Null
        Set-ItemProperty -Path $k -Name "(default)"      -Value "mscoree.dll"
        Set-ItemProperty -Path $k -Name "ThreadingModel" -Value "Both"
        Set-ItemProperty -Path $k -Name "Class"          -Value $ClassFullName
        Set-ItemProperty -Path $k -Name "Assembly"       -Value $asmString
        Set-ItemProperty -Path $k -Name "RuntimeVersion" -Value $runtime
        Set-ItemProperty -Path $k -Name "CodeBase"       -Value $codeBase
    }

    New-Item -Path "$clsidKey\ProgId" -Force | Out-Null
    Set-ItemProperty -Path "$clsidKey\ProgId" -Name "(default)" -Value $ProgId
}

Register-ComClass -Clsid "352F4B34-C24C-42FB-AF36-FE303D48BB57" -ProgId "ShapePalette.Connect"     -ClassFullName "ShapePalette.Connect"
Register-ComClass -Clsid "7F4C3A87-FB96-43D8-8B52-B9D92EF2A81B" -ProgId "ShapePalette.PaletteHost" -ClassFullName "ShapePalette.UI.PaletteHost"

# PowerPoint アドイン登録
$addinKey = "HKCU:\Software\Microsoft\Office\PowerPoint\AddIns\ShapePalette.Connect"
New-Item -Path $addinKey -Force | Out-Null
Set-ItemProperty -Path $addinKey -Name "FriendlyName" -Value "Shape Palette"
Set-ItemProperty -Path $addinKey -Name "Description"  -Value "図形をタブ別に並べ、クリックでスライドに挿入するパレット"
Set-ItemProperty -Path $addinKey -Name "LoadBehavior" -Value 3 -Type DWord

Write-Host "登録完了 (per-user)。" -ForegroundColor Green
Write-Host "  DLL: $dll"
Write-Host "PowerPoint を再起動してください。"
