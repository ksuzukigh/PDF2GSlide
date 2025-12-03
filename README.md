# PDF to Google Slides Converter マニュアル

## 概要
NotebookLMなどで作成したPDFファイルを、ページごとに画像化してGoogleスライドに変換するWebアプリです。

---

## ファイル構成

| ファイル名 | 役割 |
|-----------|------|
| `app.py` | メインのWebアプリ本体 |
| `make_slides.py` | コマンドライン版（単体実行用） |
| `start.command` | Mac用の起動スクリプト |
| `token.json` | Google認証トークン（自動生成） |
| `credentials.json` | Google Cloud認証情報（要準備） |

---

## 事前準備

1. **Python環境の構築**
   - Anacondaで `pdf-slide` という名前の仮想環境を作成
   - 必要なライブラリをインストール：
     - `streamlit`（Web画面）
     - `PyMuPDF`（PDF処理）
     - `google-auth`、`google-auth-oauthlib`、`google-api-python-client`（Google連携）

2. **Google Cloud設定**
   - Google Cloud Consoleで「Google Slides API」と「Google Drive API」を有効化
   - OAuth 2.0クライアントIDを作成し、`credentials.json`をダウンロード
   - `credentials.json`を`app.py`と同じフォルダに配置

---

## 起動方法

**方法1: start.commandをダブルクリック**（Mac推奨）

**方法2: ターミナルから起動**
```bash
cd /Users/keiichi/Desktop/Work/python/PDF2GSlide
conda activate pdf-slide
python -m streamlit run app.py
```

起動後、ブラウザで `http://localhost:8501` が自動的に開きます。

---

## 使い方

1. **PDFをアップロード**
   - 画面の「PDFファイルをここにドロップしてください」にPDFをドラッグ＆ドロップ

2. **オプション選択**
   - 「スライドごとの画像(PNG)もダウンロードする」にチェックを入れると、変換した画像のZIPファイルも取得可能

3. **変換開始**
   - 「🚀 スライド作成を開始」ボタンをクリック
   - 初回のみGoogleアカウント認証画面が表示される

4. **完了後**
   - 「Googleスライドを開く」リンクから完成したスライドを確認
   - オプションを選択していた場合は「画像(ZIP)をダウンロード」ボタンが表示される

---

## 処理の流れ

```
PDFアップロード
    ↓
各ページを高解像度PNG画像に変換（2倍解像度）
    ↓
Google Driveに画像をアップロード（公開設定）
    ↓
Googleスライドを新規作成
    ↓
各スライドに画像を全画面配置（720pt x 405pt、16:9比率）
    ↓
完成したスライドのURLを表示
```

---

## 注意事項

- 変換中はGoogle Driveに画像ファイルが一時的にアップロードされます（公開設定）
- 大きなPDFは処理に時間がかかります
- 認証トークン（`token.json`）は有効期限があり、期限切れ時は再認証が必要です

