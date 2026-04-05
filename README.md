# Excel Hanko Add-in - 電子印鑑（データ印）

日本式の電子印鑑（データ印）を Excel に挿入する Office Web アドインです。
Excel Online（Web版）およびデスクトップ版 Excel の両方で動作します。

![Office Add-in](https://img.shields.io/badge/Office-Add--in-blue)
![Platform](https://img.shields.io/badge/Platform-Excel%20Online%20%7C%20Desktop-green)
![License](https://img.shields.io/badge/License-MIT-yellow)

## 機能

- 丸型の日本式データ印（電子印鑑）を Canvas で動的生成
- 上段（苗字）・中段（日付）・下段（部署名等）の3段構成
- リアルタイムプレビュー付きの作業ウィンドウ UI
- ワンクリックでアクティブセルの位置に PNG 画像として挿入
- 日付は今日の日付を自動入力（YYYY/MM/DD 形式）
- 押印ごとにユニークなトラッキングID を自動付与（簡易証跡機能）

### 印鑑デザイン

```
    ╭──────────╮
   │   山 田    │  ← 上段: 苗字など
   │──────────│
   │ 2026/04/05 │  ← 中段: 日付
   │──────────│
   │   営業部   │  ← 下段: 部署名（空欄可）
    ╰──────────╯
                a3f1b20e  ← トラッキングID（極小グレー文字）
```

- 外枠: 赤い円（丸型）
- 区切り: 円の弦による横線で3分割
- 文字色: すべて赤色
- フォントサイズ: 領域に合わせて自動調整
- トラッキングID: 円の右下に8桁の16進数を極小グレー文字で描画

## トラッキング機能

「押印」ボタンを押すたびに、`crypto.getRandomValues` で8桁のランダムな16進数ID（例: `a3f1b20e`）を生成します。

- ID は印鑑画像の右下にグレーの極小文字（8px）で埋め込まれる
- 押印完了後、ステータスバーにも ID が表示される（例: `押印しました (ID: a3f1b20e)`）
- プレビューには ID は表示されず、実際の押印時のみ付与される
- 見た目上のシリアルナンバーとして、どの印影がいつ生成されたかの簡易的な証跡になる

## テクノロジースタック

| 技術 | 用途 |
|---|---|
| HTML / CSS | 作業ウィンドウ UI |
| JavaScript | 印鑑描画ロジック・Office.js 連携 |
| HTML5 Canvas | 印鑑画像の動的生成 |
| Office.js (Excel.js API) | Excel へのシェイプ（画像）挿入 |
| Web Crypto API | トラッキングID の乱数生成 |
| Node.js / http-server | ローカル HTTPS 開発サーバー |

## プロジェクト構成

```
excel-hanko-addin/
├── manifest.xml          # Office アドインマニフェスト
├── package.json          # 依存関係・起動スクリプト
├── README.md
└── src/
    ├── taskpane.html     # 作業ウィンドウ UI
    ├── taskpane.js       # 印鑑生成 & Excel 挿入ロジック
    └── assets/
        └── icon-32.png   # リボンアイコン
```

## セットアップ

### 前提条件

- [Node.js](https://nodejs.org/) v18 以上
- Microsoft 365 アカウント（Excel Online を使用する場合）

### インストール

```bash
git clone https://github.com/<your-username>/excel-hanko-addin.git
cd excel-hanko-addin
npm install
```

### 開発サーバーの起動

```bash
npm start
```

初回起動時に `office-addin-dev-certs` が自己署名 SSL 証明書を自動生成・インストールします。
Windows の場合、証明書の信頼ダイアログが表示されたら「はい」を選択してください。

サーバーが起動すると `https://localhost:3000` でアクセス可能になります。

> **Note:** `npm start` の証明書パスは環境に合わせて `package.json` 内の `scripts.start` を編集してください。

## Excel Online へのサイドロード

1. [office.com](https://office.com) で Excel Online を開く
2. **挿入** タブ → **アドイン** → **マイ アドイン**
3. **カスタム アドインのアップロード** をクリック
4. プロジェクト内の `manifest.xml` を選択してアップロード
5. ホームタブに「押印」ボタンが表示される

## デスクトップ版 Excel へのサイドロード (Windows)

1. 任意の場所にマニフェスト共有フォルダを作成（例: `C:\AddinManifests`）
2. `manifest.xml` をそのフォルダにコピー
3. Excel → **ファイル** → **オプション** → **トラスト センター** → **トラスト センターの設定**
4. **信頼できるアドイン カタログ** にフォルダパスを追加
5. Excel を再起動 → **挿入** → **アドイン** → **共有フォルダ** から「電子印鑑」を追加

## 使い方

1. アドインの作業ウィンドウが開いたら、各欄に入力:
   - **上段**: 苗字など（例: 山田）
   - **中段**: 日付（今日の日付が自動入力済み）
   - **下段**: 部署名など（空欄可）
2. プレビューで印鑑の見た目を確認
3. Excel 上で挿入先のセルを選択
4. **「押印」** ボタンをクリック
5. 選択セルの位置に印鑑画像が挿入される（トラッキングID 付き）
6. ステータスバーに挿入結果と ID が表示される

## カスタマイズ

`src/taskpane.js` 内の定数を変更することで、印鑑のサイズや色を調整できます:

```js
const STAMP_SIZE = 200;            // Canvas 描画解像度 (px)
const INSERT_SIZE_PX = 60;         // Excel 挿入サイズ (px)
const STAMP_COLOR = "#c62828";     // 印鑑の色（赤）
const TRACKING_COLOR = "#aaaaaa";  // トラッキングID の文字色（グレー）
const TRACKING_FONT_SIZE = 8;      // トラッキングID の文字サイズ (px)
```

## API 要件

このアドインは以下の Excel JavaScript API を使用します:

- `Excel.run()` — バッチ処理コンテキスト
- `Workbook.getSelectedRange()` — 選択セルの取得
- `Worksheet.shapes.addImage()` — 画像シェイプの挿入（**要件セット: ExcelApi 1.9 以上**）

## ライセンス

MIT
