# Shape Palette アドインの per-user COM 登録を解除する（管理者不要）。
$ErrorActionPreference = "SilentlyContinue"

# PowerPoint アドイン登録の削除
Remove-Item -Path "HKCU:\Software\Microsoft\Office\PowerPoint\AddIns\ShapePalette.Connect" -Recurse -Force

# ProgId / CLSID の削除
foreach ($pid_ in @("ShapePalette.Connect", "ShapePalette.PaletteHost")) {
    Remove-Item -Path "HKCU:\Software\Classes\$pid_" -Recurse -Force
}
foreach ($clsid in @("352F4B34-C24C-42FB-AF36-FE303D48BB57", "7F4C3A87-FB96-43D8-8B52-B9D92EF2A81B")) {
    Remove-Item -Path "HKCU:\Software\Classes\CLSID\{$clsid}" -Recurse -Force
}

Write-Host "登録解除完了。PowerPoint を再起動してください。" -ForegroundColor Yellow
