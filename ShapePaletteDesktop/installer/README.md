# Shape Palette — 配布用インストーラー

PowerPoint 図形パレット アドインの **per-user（管理者不要）インストーラー**。

## ビルド方法

```powershell
powershell -ExecutionPolicy Bypass -File build-installer.ps1
```

これで以下を一括実行します:
1. Release ビルド（`ShapePalette.dll` ＋ `Newtonsoft.Json.dll`）
2. 自己署名証明書で DLL に署名（初回は証明書を自動作成し、このPCの信頼ストアに登録）
3. Inno Setup でインストーラー作成
4. インストーラーに署名

出力: `Output\ShapePaletteSetup.exe`（配布する単一ファイル）

> 署名不要なら `-NoSign` を付ける。Inno Setup と Visual Studio(MSBuild) が必要。

## インストール（各ユーザー）

1. **`ShapePaletteSetup.exe` をダブルクリック**（管理者不要）
2. SmartScreen の警告（「WindowsによってPCが保護されました」）が出たら **「詳細情報」→「実行」**
   - 署名証明書を信頼させれば警告は出なくなる（下記）
3. PowerPoint を起動 → ホームタブに「**パレットを開く**」ボタン

インストール先: `%LOCALAPPDATA%\ShapePalette\`（DLL）。データ（登録図形）は `%APPDATA%\ShapePalette\`。

## 警告を消す（任意・研究室配布向け）

自己署名のため、他PCでは「不明な発行元」警告が出ます。消すには公開証明書を信頼させます:

1. `ShapePalette-PublicCert.cer` を対象PCにコピー
2. 右クリック →「証明書のインストール」→ **現在のユーザー** →「証明書をすべて次のストアに配置する」
   → **信頼された発行元** に入れる（＋**信頼されたルート証明機関**にも入れると確実）
3. 以降、署名済みインストーラーは警告なしで実行可能

ドメイン管理なら GPO/Intune で証明書を一括配布可能。少人数なら「詳細情報→実行」で1回通すだけでもOK。

## アンインストール

設定 → アプリ → 「Shape Palette」→ アンインストール（または `%LOCALAPPDATA%\ShapePalette\unins000.exe`）。
登録図形データ（`%APPDATA%\ShapePalette\`）は残るので、再インストールで復元されます。
