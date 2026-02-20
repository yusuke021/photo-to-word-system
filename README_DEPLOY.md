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

- 📊 表形式での写真配置（行数・列数・セルサイズ・罫線・配置をカスタマイズ可能）
- 📷 複数画像の一括処理（JPG, PNG, GIF, BMP, HEIC対応）
- 📝 ファイル名から部品名を自動抽出（サイドバー一番上で設定）
- 🎨 文字数に応じた自動フォントサイズ調整（7pt〜12pt）
- 🖼️ 画像品質（PPI）選択：印刷用220ppi（デフォルト）、高性能300ppi、標準150ppi
- 📄 複数ページの自動生成
- 📤 既存Wordファイルへの追記対応（アップロード可能）
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

## � セキュリティとプライバシー

- ✅ すべてのデータ処理はクライアント側で実行
- ✅ アップロードされたファイルはサーバーに保存されません
- ✅ セッション終了後、データは自動的に削除されます
- ✅ HTTPS接続により通信は暗号化されます

詳細は [SECURITY.md](SECURITY.md) をご参照ください。

## �📄 ライセンス

MIT License
