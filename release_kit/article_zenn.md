---
title: "Excel はんこアドイン を作った — ローカル完結で動かす実用ツール"
emoji: "🛠️"
type: "tech"
topics: ["python", "個人開発", "oss"]
published: false
---

> 本記事は Zenn 用の下書きです。Qiita に出す場合は先頭の frontmatter を削除してください。

## TL;DR

日本式の電子印鑑（データ印）を Excel に挿入する Office Web アドインです。 Excel Online（Web版）およびデスクトップ版 Excel の両方で動作します。

- リポジトリ: https://github.com/highdefinitionaudiodriver/excel-hanko-addin
- ライセンス: MIT / バージョン: v1.1.0

## 作った背景・課題

（なぜ作ったか。既存ツールの不満、手作業の手間などを 2〜3 段落で。）

## できること

- 丸型の日本式データ印（電子印鑑）を Canvas で動的生成
- 上段（苗字）・中段（日付）・下段（部署名等）の3段構成
- リアルタイムプレビュー付きの作業ウィンドウ UI
- ワンクリックでアクティブセルの位置に PNG 画像として挿入
- 日付は今日の日付を自動入力（YYYY/MM/DD 形式）
- 押印ごとにユニークなトラッキングID を自動付与（簡易証跡機能）
- ID は印鑑画像の右下にグレーの極小文字（8px）で埋め込まれる
- 押印完了後、ステータスバーにも ID が表示される（例: 押印しました (ID: a3f1b20e)）

## 仕組み / 工夫した点

（設計上のポイント。ローカル完結・プライバシー配慮・依存の少なさ など。）

## 使い方

```bash
# インストール・起動例（README から転記）
```

## ハマったところ

（開発中の課題と解決。）

## おわりに

フィードバックは Issues / Star をいただけると励みになります。

リポジトリ: https://github.com/highdefinitionaudiodriver/excel-hanko-addin
