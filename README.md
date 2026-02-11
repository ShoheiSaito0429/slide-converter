# Slide Converter

画像→編集可能PowerPoint変換ツール（プロトタイプ）

Claude Vision APIで画像を解析し、python-pptxで編集可能なPowerPointファイルを生成します。

## セットアップ

```bash
pip install -r requirements.txt
python3 app.py
```

ブラウザで http://localhost:8081 にアクセス

## 機能

- スライド画像（PNG/JPG/WEBP）をアップロード
- Claude Vision APIで自動解析
- テキスト、図形、画像領域を抽出
- 編集可能なPowerPoint (.pptx) を生成
- デモモード対応（APIキー不要で動作確認）

## 技術スタック

- Flask (Webサーバー)
- Claude API (画像解析)
- python-pptx (PowerPoint生成)
- Pillow (画像処理)

詳細は `INSTRUCTIONS.md` を参照してください。
