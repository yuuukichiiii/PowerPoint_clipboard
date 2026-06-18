# Shape Palette — Claude Code 引き継ぎ

PowerPoint 用「図形クリップボード／パレット」タスクペインアドイン。図形をタブ別に並べ、
クリックまたはキーボードで現在のスライドに挿入する。Claude.ai 上で設計〜初期実装まで完了し、
ここから先の実装・デプロイを Claude Code に引き継ぐ。

---

## 1. 成果物（現在のファイル）

| ファイル | 役割 |
|---|---|
| `taskpane.html` | パネル本体。UI + 全ロジックを単一ファイルに内包（CSS は `<style>`、JS は `<script>` 内）。office.js は CDN 読み込み。 |
| `commands.html` / `commands.js` | ショートカットから呼ぶ関数（`showAsTaskpane` / `hide`）。`Office.actions.associate` で登録。 |
| `shortcuts.json` | キー定義。`Ctrl+Shift+1`=表示 / `Ctrl+Shift+2`=非表示（Mac は Cmd）。 |
| `manifest.xml` | add-in only manifest（XML）。タスクペインボタン + `ExtendedOverrides` でショートカット参照。 |
| `README.md` | セットアップ・使い方・制約。 |

すべて素の HTML/JS でビルド不要（webpack 等は使っていない）。

---

## 2. 設計判断の経緯（なぜこうなっているか）

- **VSTO(C#) ではなく Office.js を採用**。理由: クロスプラットフォーム・配布が容易・モダンな UI 開発。
  保存対象がシンプルな図形・矢印・テキスト入りパーツ中心と確認できたため、VSTO の唯一の優位点
  （クリップボード経由の完全コピペ）は重要度が低いと判断。
- **挿入方法は「クリック挿入 → 手動で移動」**。Office.js からスライドへの真のドラッグ&ドロップは
  制約が強いため。ユーザー合意済み。
- **挿入位置は「選択中の図形の右隣」**（なければスライド中央＋カスケード）。理由は §3 参照。
- **取り込みは "レシピ方式"**。図形の見た目（塗り・枠・テキスト・フォント・サイズ）をプロパティとして
  読み取り保存し、挿入時に同じ見た目を再生成する。OOXML 丸ごとコピーではない。

---

## 3. 確定した技術的制約（Office.js 由来・調査済み）

これらは API の仕様上の壁。回避策込みで実装済み。蒸し返さないこと。

1. **表示領域（ビューポート）の中心・ズーム・スクロール位置は取得できない。**
   → 「画面表示の中央」への挿入は不可能。代替として「選択図形の隣」に挿入している。
2. **既存図形のジオメトリ種別（角丸四角／五角形など）は読み取れない。**
   → 取り込み時は `roundRectangle` 固定。編集モードのドロップダウンで後から変更可能。
3. **テーマカラー（Theme Colors）で塗った図形は塗り色が `#000000` として読まれるバグがある。**
   → 標準の色 / カスタム RGB で塗れば正しく取れる。取り込み時に黒を検出したら警告トーストを表示。
4. **グラデーション塗り・グループ図形・アニメーションの完全再現は不可。**
   → 単色塗り中心の設計。複数図形は「複合パーツ」（相対座標保存）で近似的に再現。

使用している主な API: `addGeometricShape`, `getSelectedShapes`, `getSelectedSlides`,
`fill.setSolidColor` / `fill.foregroundColor`, `lineFormat.color/weight/dashStyle/visible`,
`textFrame.textRange.text` / `.font.color/size/bold`。要件セットは PowerPointApi 1.x 相当。
新しい API を使う場合は対象 Office バージョンの要件セット対応を確認すること。

---

## 4. データスキーマ（localStorage キー: `shapePalette.v3`）

```jsonc
{
  "activeTab": "t_shapes",
  "tabs": [
    { "id": "t_fav",    "name": "お気に入り", "items": [ /* recipe */ ] },  // 特殊タブ
    { "id": "t_recent", "name": "最近",       "items": [ /* recipe */ ] },  // 特殊タブ（自動）
    { "id": "t_shapes", "name": "図形",       "items": [ /* recipe */ ] },
    { "id": "t_arrows", "name": "矢印",       "items": [ /* recipe */ ] }
  ]
}
```

単一図形 recipe:
```jsonc
{
  "id": "i...", "kind": "single", "geometricType": "roundRectangle", "name": "…",
  "width": 120, "height": 90,
  "fill": "#RRGGBB" | null,
  "line": { "color": "#RRGGBB", "weight": 1.5, "dash": "Solid" } | null,
  "text": "…" | null,
  "font": { "color": "#RRGGBB", "size": 18, "bold": false } | null,
  "favOf": "<元id>",     // お気に入りタブ内のクローンのみ
  "recentOf": "<元id>"   // 最近タブ内のクローンのみ（重複排除キー）
}
```

複合パーツ recipe:
```jsonc
{
  "id": "i...", "kind": "composite", "name": "複合パーツ N",
  "width": <bbox幅>, "height": <bbox高>,
  "parts": [
    { "geometricType": "roundRectangle", "fill": …, "line": …, "text": …, "font": …,
      "dx": <左上からの相対X>, "dy": <相対Y>, "w": <幅>, "h": <高> }
  ]
}
```

- 特殊タブ ID は `t_fav` / `t_recent`（`SPECIAL` 配列・`isSpecial()`）。リネーム・削除不可。
- 最近タブは上限 `RECENT_MAX = 24`、`pushRecent()` が先頭追加＋重複排除。
- ストレージは `OfficeRuntime.storage` → `localStorage` → メモリ の順にフォールバック（`Store`）。
- スキーマ変更時は新キー（`shapePalette.v4` …）に上げ、`load()` にマイグレーションを追加すること
  （`t_recent` 補完の前例あり）。

---

## 5. 実装済み機能（taskpane.html 内の主な関数）

- タブ: `renderTabs` / `addTab` / `deleteTab` / `setActiveTab`、編集モード `toggleEdit`
- 取り込み: `captureSelection` → `shapeLoad` / `extractProps` → `commitIndividual` / `commitComposite`
  （複数選択時はバナーで「まとめて1つ／個別」を選択）
- 挿入: `insertRecipe`（選択図形の隣 → 中央フォールバック、単一/複合両対応）
- キーボード: `onKey`（`1`–`9` 挿入 / `↑↓←→` 移動 / `Enter`・`Space` / `[` `]` タブ / `Esc`）、`moveFocus` / `applyFocus`
- 履歴: `pushRecent` / `clearRecent`
- お気に入り: `toggleFav`
- プレビュー SVG 生成: `svgPreview` / `partSvg` / `svgComposite` / `SHAPES`(ジオメトリ→SVG)
- 入出力: `exportJSON` / `importJSON` / `resetAll`
- 非 PowerPoint 環境（ブラウザ直開き）では demo モードで動作（挿入はトースト、取り込みは見本生成）。

---

## 6. 次にやること（優先度順）

### A. 導入の簡易化（最優先）
今は localhost + 自己署名証明書 + 手動サイドロードで手間が大きい。次の形にしたい。
1. `taskpane.html` / `commands.html` / `commands.js` / `shortcuts.json` を **GitHub Pages**
   （または Netlify / Azure Static Web Apps）にデプロイ。
2. `manifest.xml` 内の全 URL（`SourceLocation`、`Taskpane.Url`、`Commands.Url`、
   `ExtendedOverrides Url`、各 IconUrl）を `https://localhost:3000/...` から公開 URL に置換。
3. これで証明書・ローカルサーバー不要に。導入は manifest の Web アップロード or 共有フォルダ登録のみ。
4. 開発用に `package.json` を作り、`start`/`stop` スクリプトを用意（下記参照）。
   - `office-addin-dev-certs`, `office-addin-debugging`, `http-server` を devDependencies に。
   - `"start": "office-addin-debugging start manifest.xml desktop"` 等。
5. 余裕があれば `assets/` に 16/32/80px のアイコン PNG を用意（現在マニフェスト参照のみ）。
6. 研究室配布を見据えるなら、大学 M365 の集中展開（一元展開）手順も README に追記。

### B. 機能追加（ユーザー選択済みの次フェーズ候補）
- **検索 / 絞り込み**: 上部に検索ボックス、全タブ横断で name フィルタ。タイル増加時に必須。
- **ドラッグ並べ替え**: タイル順・タブ順の入れ替え（編集モード内、pointer events 推奨）。
- **挿入時リカラー**: 同じ図形を色だけ変えて挿入（テーマ色問題の回避にも）。
- **挿入位置の設定**: 「選択図形の隣 / スライド中央 / 直前と同じ位置」を切替。
- **画像・アイコン登録**: ローカル画像アップロード→ base64 保存→挿入。
  ※ 既存スライド上の画像取り込みは Office.js の制約があり得るので、先に挙動確認すること。

### C. クリーンアップ
- `tile()` 内、星表示の条件式 `if(tab.id!=="t_fav"&&rec.kind!=="composite"||tab.id!=="t_fav")`
  が冗長（実質 `tab.id!=="t_fav"`）。整理する。

---

## 7. 開発セットアップ（現状の手順・要点）

```powershell
# フォルダ内で
npx office-addin-dev-certs install                         # 初回のみ。localhost証明書を信頼
npx http-server . -S -C <localhost.crt> -K <localhost.key> -p 3000   # 配信（ポートは manifest と一致）
# 別ターミナル
npx office-addin-debugging start manifest.xml desktop      # PowerPoint 起動＋サイドロード
npx office-addin-debugging stop  manifest.xml              # 終了時
```

ハマりどころ: 配信ポートは manifest の `https://localhost:3000` と一致必須 / start 前に
PowerPoint は閉じておく / パネル真っ白は証明書未信頼 / Microsoft 365 アカウントでサインイン必須 /
`npx office-addin-manifest validate manifest.xml` で manifest 検証可。

> §6-A を実施すれば、この localhost 手順は開発時のみで済み、配布は公開 URL 経由になる。