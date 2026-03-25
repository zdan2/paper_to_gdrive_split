# Dropbox Paper → ローカル確認 → Google Drive 3段構成

## フォルダ構成

```text
review_workspace/
  exported/   # Part 1 が Dropbox Paper を書き出す場所
  approved/   # 人の目で確認して「これを上げる」と決めたファイルだけ置く場所
```

## 事前準備

### 1. Python パッケージ

```bash
python -m pip install -r requirements.txt
```

### 2. Dropbox

- Dropbox App Console で API アプリを作成
- アクセストークンを発行
- 実行前に環境変数 `DROPBOX_TOKEN` に設定

Windows PowerShell:

```powershell
$env:DROPBOX_TOKEN="your_dropbox_token"
```

Windows コマンドプロンプト:

```cmd
set DROPBOX_TOKEN=your_dropbox_token
```

### 3. Google Drive

- Google Cloud で Drive API を有効化
- OAuth client (Desktop app) を作成
- JSON を `credentials.json` としてこのフォルダに置く
- 初回実行時にブラウザ認証され、`token.json` が作られる

## Part 1: Dropbox Paper をローカルへ書き出す

```bash
python 01_export_paper_to_local.py --dropbox-root "/議事録" --preferred-format docx
```

ポイント:
- `.paper` だけを対象にします
- `.papert` は対象外です
- `--preferred-format docx` は「docx が取れれば優先」という意味で、見つからなければ Dropbox の既定形式へ自動で戻します

## 人の確認

- `review_workspace/exported/` を開いて内容確認
- 問題ないものだけ `review_workspace/approved/` へ移動またはコピー
- サブフォルダ構造はそのまま保つと、Google Drive 側にも再現されます

## Part 2: approved フォルダのアップロード計画を作る

```bash
python 02_plan_upload_from_local.py
```

出力:
- `review_workspace/upload_plan.csv`

## Part 3: approved フォルダから Google Drive へ転送

```bash
python 03_upload_approved_to_gdrive.py --drive-folder-id "保存先フォルダID" --skip-existing
```

動作:
- Google Drive の importFormats に照らして、変換可能なものは Google ドキュメント化
- 変換できないものは通常ファイルとして Drive にアップロード
- approved 配下のサブフォルダを Drive 側にも作成

## 補足

- 変換対象の判定は Google Drive API の `about.importFormats` を参照します
- 保存先フォルダ ID を省略するとマイドライブ直下に保存します
- 同名ファイルがあると困る場合は `--skip-existing` を使ってください
