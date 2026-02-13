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

### 📝 部品名の自動挿入
- ✅ ファイル名から部品名を自動抽出
- ✅ 文字数に応じた自動フォントサイズ調整（7pt〜12pt）
- ✅ 日本語: MS明朝 / 英数字: Times New Roman
- ✅ 写真区分でフィルタリング（P/M）

### 📄 複数ページ対応
- ✅ 1ページに収まらない場合は自動改ページ
- ✅ すべてのページで同じ表形式を維持

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

## ライセンス

MIT License
