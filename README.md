# 写真貼り付けシステム v2.0

ローカルファイルから写真を選択してMicrosoft Wordファイルに表形式で貼り付けるシステムです。

**デスクトップ版**と**Webアプリ版**の両方を提供しています！

---

## 🌐 Webアプリ版（推奨）

### ローカルで起動

```bash
# 依存関係をインストール
pip install -r requirements.txt

# Webアプリを起動
streamlit run app.py
```

ブラウザで `http://localhost:8501` が自動的に開きます。

### 無料デプロイ方法（Streamlit Cloud）

1. **GitHubにプッシュ**
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git remote add origin <your-repo-url>
   git push -u origin main
   ```

2. **Streamlit Cloudにデプロイ**
   - https://share.streamlit.io/ にアクセス
   - GitHubでサインイン
   - "New app" をクリック
   - リポジトリを選択
   - Main file: `app.py`
   - Deploy!

**完全無料**で全世界からアクセス可能なWebアプリになります！

---

## 🖥️ デスクトップ版

### 起動方法

```bash
python image_to_word.py
```

---

## 機能

### 📊 表形式での配置
- ✅ 行数・列数を自由に設定（デフォルト: 8行×2列）
- ✅ 奇数行に写真、偶数行に部品名
- ✅ 罫線設定（なし/すべて/外枠のみ）
- ✅ 表の配置（左/中央/右）

### 📷 画像処理
- ✅ 複数画像の一括処理
- ✅ アスペクト比を維持した自動リサイズ
- ✅ 対応形式: JPG, PNG, GIF, BMP, HEIC
- ✅ セルの高さに完全フィット
- ✅ 画像品質（PPI）設定: 印刷用220ppi（デフォルト）、高性能300ppi、標準150ppi

### 📝 部品名の自動挿入（サイドバー一番上で設定）
- ✅ ファイル名から部品名を自動抽出
- ✅ 文字数に応じた自動フォントサイズ調整（7pt〜12pt）
- ✅ 日本語: MS明朝 / 英数字: Times New Roman
- ✅ 写真区分でフィルタリング（P/M）

### 📄 Wordファイル操作
- ✅ 新規Wordファイルの作成（**template.docx**があれば自動的にテンプレート適用）
- ✅ 既存Wordファイルへの追記（アップロード対応）
- ✅ 1ページに収まらない場合は自動改ページ
- ✅ すべてのページで同じ表形式を維持

### 📋 テンプレート機能（NEW!）
- ✅ プロジェクトルートに `template.docx` を配置すると自動的に適用
- ✅ ヘッダー・フッター、ページ設定、スタイルを継承
- ✅ テンプレートはGitリポジトリに含めて管理可能
- ✅ テンプレートなしの場合は空白のWordファイルを作成
- 📖 詳細は [TEMPLATE_GUIDE.md](TEMPLATE_GUIDE.md) を参照

---

## 📝 使い方

### Webアプリ版の基本的な使い方

1. **Webアプリを起動**
   ```bash
   streamlit run app.py
   ```

2. **サイドバーで設定**
   - 📝 名称挿入設定：部品名を自動挿入するか選択
   - ⚙️ 表の設定：行数、列数、罫線、配置を設定
   - 📐 セルサイズ：奇数行（写真用）と偶数行（説明用）のサイズを調整
   - 🎨 画像品質設定：印刷用220ppi（推奨）、高性能300ppi、標準150ppi

3. **メイン画面で操作**
   - 📄 既存のWordファイルをアップロード（オプション）
     - アップロードしない場合：`template.docx` があれば自動適用、なければ新規作成
     - アップロードした場合：既存ファイルに表を追記
   - 📷 写真ファイルをアップロード（複数選択可）

4. **Wordファイルを生成**
   - 「Wordファイルを生成」ボタンをクリック
   - 自動的にダウンロードが開始されます

### テンプレート機能の使い方

**テンプレートファイルの準備：**
1. フォーマット済みのWordファイルを `template.docx` という名前で保存
2. プロジェクトのルートディレクトリ（`app.py` と同じ場所）に配置
3. Gitリポジトリに含める場合：
   ```bash
   git add template.docx
   git commit -m "Add template file"
   git push origin main
   ```

**テンプレートに含められる内容：**
- ヘッダー・フッター（会社ロゴ、ページ番号など）
- ページ設定（余白、用紙サイズ、向きなど）
- デフォルトのフォント・スタイル設定
- 既存のコンテンツ（タイトル、説明文など）

**動作：**
- テンプレートの内容はそのまま維持され、その後に写真の表が追加されます
- 既存のWordファイルをアップロードした場合は、テンプレートは使用されません

詳細は [TEMPLATE_GUIDE.md](TEMPLATE_GUIDE.md) を参照してください。

---

## トラブルシューティング

### HEICファイルが開けない場合
```bash
pip install pillow-heif
```

### Streamlitが起動しない場合
```bash
pip install --upgrade streamlit
streamlit hello  # 動作確認
```

### tkinterが見つからない場合（デスクトップ版）

**macOS:**
```bash
brew install python-tk
```

**Ubuntu/Debian:**
```bash
sudo apt-get install python3-tk
```

---

## 🔒 セキュリティ

このアプリケーションはすべてのデータをローカルで処理します：
- ✅ アップロードされた画像はサーバーに保存されません
- ✅ 生成されたWordファイルは即座にダウンロードされます
- ✅ ユーザーデータの収集や追跡は一切行いません

詳細は [SECURITY.md](SECURITY.md) をご覧ください。

脆弱性を発見した場合は、[Security Advisory](https://github.com/yusuke021/photo-to-word-system/security/advisories/new) から報告してください。

---

## ライセンス

MIT License
