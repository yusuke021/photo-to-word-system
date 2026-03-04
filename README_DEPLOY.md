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

- � **テンプレート機能（NEW!）：** `template.docx` を配置すると新規ファイルに自動適用
- �📊 表形式での写真配置（行数・列数・セルサイズ・罫線・配置をカスタマイズ可能）
- 📷 複数画像の一括処理（JPG, PNG, GIF, BMP, HEIC対応）
- 📝 ファイル名から部品名を自動抽出（サイドバー一番上で設定）
- 🎨 文字数に応じた自動フォントサイズ調整（7pt〜12pt）
- 🖼️ 画像品質（PPI）選択：印刷用220ppi（デフォルト）、高性能300ppi、標準150ppi
- 📄 複数ページの自動生成
- 📤 既存Wordファイルへの追記対応（アップロード可能）
- 💾 Wordファイルのダウンロード

## 📦 技術スタック

- **Python 3.11** (推奨 - 互換性保証)
- **Streamlit** (>=1.28.0, <2.0.0) - Webフレームワーク
- **python-docx** (>=0.8.11, <2.0.0) - Word文書操作
- **Pillow** (>=9.0.0, <11.0.0) - 画像処理
- **pillow-heif** (>=0.10.0, <1.0.0) - HEIC形式サポート

### プロジェクト構成
- `.python-version`: Python 3.11を指定 (互換性のため)
- `.streamlit/config.toml`: Streamlit設定 (テーマ、サーバー設定)
- `requirements.txt`: 依存パッケージのバージョン固定
- `template.docx`: オプションのテンプレートファイル

## 📖 使い方

1. 画像ファイルをアップロード
2. サイドバーで設定を調整（名称挿入、表の設定、セルサイズ、画像品質など）
3. （オプション）既存のWordファイルをアップロード（追記したい場合）
4. 「Wordファイルを生成」ボタンをクリック
5. 生成されたWordファイルをダウンロード

### 📋 テンプレート機能

新規Wordファイル作成時にテンプレートを自動適用できます：

- プロジェクトルートに `template.docx` を配置
- 既存Wordファイルをアップロードしない場合、自動的にテンプレートが使用されます
- ヘッダー・フッター、ページ設定、スタイルが継承されます

詳細は [TEMPLATE_GUIDE.md](TEMPLATE_GUIDE.md) を参照してください。

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
## 🔧 Streamlit Cloudデプロイ設定

### 必要なファイル
1. **app.py** - メインアプリケーション
2. **requirements.txt** - Python依存パッケージ
3. **.python-version** - Python 3.11を指定（重要）
4. **.streamlit/config.toml** - Streamlit設定
5. **template.docx** - オプションのテンプレート

### 主要設定
```toml
[server]
maxUploadSize = 200  # 最大200MBまでアップロード可能
enableXsrfProtection = true
enableCORS = false

[browser]
gatherUsageStats = false
```

### デプロイ手順
1. GitHubリポジトリを作成
2. 上記のファイルをプッシュ
3. https://share.streamlit.io でアプリを作成
4. リポジトリと `app.py` を指定
5. 自動でPython 3.11が検出されます

## ⚠️ トラブルシューティング

### エラー: "MediaFileStorageError"
**原因**: セッション管理の問題、古いダウンロードリンクの参照  
**解決策**: 
- ブラウザのキャッシュをクリア (Ctrl+Shift+R)
- アプリを再起動 (Manage app → Reboot)
- 最新コードにセッション状態管理が実装済み

### エラー: "Error running app"
**原因**: Pythonバージョンの非互換性、依存パッケージの問題  
**解決策**:
- `.python-version` がリポジトリにあることを確認
- `requirements.txt` のバージョン指定を確認
- Logsを確認 (Manage app → Logs)

### ログ確認方法
1. アプリページ右下の「Manage app」をクリック
2. 「Logs」タブを開く
3. 詳細なエラー情報が表示されます

### パフォーマンスの最適化
- 大きな画像ファイルは事前に圧縮することを推奨
- 一度にアップロードする画像は50枚以下を推奨
- 印刷用(220ppi)がファイルサイズと品質のバランスが良い
## �📄 ライセンス

MIT License
