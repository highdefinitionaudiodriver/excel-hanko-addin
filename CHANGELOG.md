# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- README に「これは何？（30秒で）」「想定ユースケース・価格帯」セクションを追加
- SECURITY.md を追加（脆弱性報告フロー）
- 商用利用・カスタマイズ依頼の連絡先を README 末尾に明記
- **証跡（押印履歴）機能**：押印ごとに `timestamp / trackingId / host_app / 上中下段テキスト` を localStorage に永続化
- **CSV エクスポート機能**：UTF-8 BOM 付き RFC4180 互換 CSV を `hanko_audit_YYYYMMDD_HHMMSS.csv` でダウンロード
- 履歴クリアボタン（誤操作防止の確認ダイアログ付き）
- docs/store-listing/APPSOURCE_JA.md — Microsoft AppSource 提出用テンプレート
- **印影バリエーション 3 形状**（丸 / 角 / スタンプ）
  - サイドバーに 3 形状のラジオセレクタを追加
  - drawStamp() を形状ディスパッチに刷新（drawCircleFrame_ / drawSquareFrame_ / drawStampFrame_）
  - 角印は正方形外枠 + 水平 3 分割
  - スタンプ風は角丸長方形（横長）で部署印・会社印を想定
  - 証跡 CSV に `shape` 列を追加（どの形状で押印したか監査可能）

## [0.1.0]

### Added
- 初版リリース
