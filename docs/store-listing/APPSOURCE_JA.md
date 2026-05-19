# Microsoft AppSource 提出用テンプレート（日本語）

Excel Hanko Add-in を Microsoft AppSource に提出する際に使用する説明文・スクリーンショット要件・メタデータの雛形です。

---

## アプリ名

```
電子印鑑 for Excel（データ印）
```

英語版：`Excel Hanko - Japanese Digital Stamp`

---

## 短い説明（最大 100 文字）

```
Excel に日本式の電子印鑑（データ印）をワンクリックで挿入。脱ハンコ・在宅勤務対応の社内承認に。
```

---

## 詳細説明（最大 4,000 文字）

```
■ 概要
紙の書類に印鑑を押すために出社・郵送する非効率を、Excel ファイル内の
ワンクリック押印に置き換えます。日付・部署・氏名を含む日本式データ印を
Canvas で動的に生成し、選択セルの位置に PNG 画像として挿入します。

■ こんな方におすすめ
・脱ハンコを進めたい中小企業の総務・経理担当
・在宅勤務で書類押印が手間なリモートワーカー
・Excel Online / Microsoft 365 を全社展開している組織

■ 主な機能
・丸型データ印（電子印鑑）のリアルタイムプレビュー
・上段（苗字）・中段（日付）・下段（部署名）の3段構成
・ワンクリックでアクティブセルに PNG として挿入
・押印ごとにユニークなトラッキングID を自動付与（簡易証跡）
・日付は今日の日付を自動入力（YYYY/MM/DD 形式）

■ 動作環境
・Excel Online（Web版）
・Excel デスクトップ版（Microsoft 365）
・ExcelApi 1.9 以上が利用可能な環境

■ プライバシー・セキュリティ
・氏名・部署名等の入力データは端末内のメモリでのみ処理され、
　外部サーバーへ送信されることはありません。
・ネットワーク通信は印影 PNG の生成・挿入のみ（ローカル処理）。

■ ライセンス
個人利用は無料（MIT ライセンス）。
法人向けのカスタマイズ・証跡 CSV エクスポート・SSO 連携は別途ご相談ください。

■ サポート
GitHub: https://github.com/highdefinitionaudiodriver/excel-hanko-addin
お問い合わせ: highdefinitionaudiodriver@gmail.com
```

---

## カテゴリ

- 主：**生産性 (Productivity)**
- 副：**業務プロセスとワークフロー (Business Process and Workflow)**

---

## キーワード（最大 7 個）

```
電子印鑑, デジタル印鑑, ハンコ, 脱ハンコ, 在宅勤務, 承認ワークフロー, 押印
```

---

## スクリーンショット要件

AppSource 提出には **5 枚以上** のスクリーンショットが必要です。以下の場面を撮影してください：

1. **作業ウィンドウ全体**（プレビューとフォーム入力欄が見える状態）
2. **押印実行後の Excel シート**（実際に印影が挿入された様子）
3. **3 段構成の入力フォーム**（苗字・日付・部署名を入力中）
4. **トラッキング ID の拡大**（印影右下のグレー極小文字 a3f1b20e 等）
5. **Excel Online + デスクトップ版の動作例**（両環境で動くことを示す）

推奨サイズ：**1280×720 以上**、PNG または JPG、ファイルサイズ < 1MB。

---

## アイコン

- **48×48 PNG**（リボン用、既存 `src/assets/icon-32.png` から拡大）
- **300×300 PNG**（AppSource 一覧用、透過背景推奨）

---

## サポート URL / プライバシーポリシー URL

GitHub Pages を有効化して以下を準備：

- サポート URL：`https://highdefinitionaudiodriver.github.io/excel-hanko-addin/support`
- プライバシーポリシー URL：`https://highdefinitionaudiodriver.github.io/excel-hanko-addin/privacy`
- 利用規約 URL：`https://highdefinitionaudiodriver.github.io/excel-hanko-addin/terms`

（テンプレート HTML は `docs/store-listing/pages/` に配置予定）

---

## 提出時のチェックリスト

- [ ] `manifest.xml` の `<ProviderName>` を本名または組織名に設定
- [ ] `<SupportUrl>` を有効な URL に変更
- [ ] `<AppDomains>` にホスティング先ドメインを追加
- [ ] 商用提出には Microsoft Partner Center アカウントが必要
- [ ] 動画デモ（30 秒以内）があると審査が早い傾向
- [ ] Microsoft 365 認定（オプション）は SOC 2 / GDPR 対応で加点
