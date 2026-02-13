# 📸 写真貼り付けシステム Webアプリ

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://share.streamlit.io)

表形式で写真をMicrosoft Wordファイルに貼り付けるWebアプリケーションです。

## 🚀 クイックスタート

### オンラインで使用（推奨）
デプロイ後のURLにアクセスするだけで使えます！

### ローカルで起動
```bash
pip install -r requirements.txt
streamlit run app.py
```

## ✨ 主な機能

- 📊 表形式での写真配置
- 📷 複数画像の一括処理（JPG, PNG, HEIC対応）
- 📝 ファイル名から部品名を自動抽出
- 🎨 文字数に応じた自動フォントサイズ調整
- 📄 複数ページの自動生成
- 💾 Wordファイルのダウンロード

## 📦 技術スタック

- Python 3.7+
- Streamlit
- python-docx
- Pillow / pillow-heif

## 📖 使い方

1. 画像ファイルをアップロード
2. 表の設定を調整（行数・列数・セルサイズなど）
3. 「Wordファイルを生成」ボタンをクリック
4. 生成されたWordファイルをダウンロード

## 🎯 ファイル名規則

偶数行に部品名を挿入する場合：
```
番号_部品名_重量_単位_素材ID_加工ID_実施者ID_写真区分_特記事項.拡張子
```

例：`001_フロントパネル_500_g_ST001_CUT001_USER001_P_特記なし.jpg`

## 📄 ライセンス

MIT License
