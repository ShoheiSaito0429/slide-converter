# Slide Converter セットアップ指示書

## 概要

「画像→編集可能PowerPoint変換ツール」のプロトタイプ。
Kirigami.ai (https://kirigami.app) のような、AI生成スライド画像をpptxに変換するWebアプリ。

Claude Vision APIで画像を解析 → JSON形式でレイアウト取得 → python-pptxでpptx生成、という流れ。

## セットアップ手順

### 1. 依存パッケージのインストール

```bash
pip install flask python-pptx Pillow
```

### 2. 動作確認（デモモード）

```bash
cd ~/slide-converter
python3 app.py
```

ブラウザで http://localhost:8081 にアクセスし、「デモモード」ボタンで動作確認。
APIキー不要でサンプルpptxが生成される。

### 3. ポート

- 8081を使用
- Tailscale経由: http://100.65.45.31:8081

## 機能

- Claude Vision APIで画像内の要素を解析
- テキスト、図形、画像領域をJSON形式で取得
- python-pptxで編集可能なpptxを生成
- デモモード対応（APIキー不要で動作確認）

詳細は README.md を参照してください。
