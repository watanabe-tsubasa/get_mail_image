# README

## setup

### install uv

https://docs.astral.sh/uv/ こちらを参照

### git clone

当repositoryを `clone` してください。

### uv sync

該当のルートに移動して

```sh
uv sync
```

## how to use

### 画像のみ、件名に「請求書」を含むメールの添付を保存

uv run main.py -o "C:\temp\mail_attachments" -s "請求書"

### 完全一致で検索

uv run main.py -o "C:\temp\mail_attachments" -s "【重要】請求書の送付" --exact

### 受信トレイ内のサブフォルダーを対象

uv run main.py -o "C:\temp\mail_attachments" -s "請求書" -f "Inbox/2025-08"

### 共有メールボックス（例: "経理部"）の「受信トレイ/請求書」配下を対象

uv run main.py -o "C:\temp\mail_attachments" -s "請求書" --store "経理部" -f "受信トレイ/請求書"

### 画像以外の添付も含めて全保存

uv run main.py -o "C:\temp\mail_attachments" -s "見積" --all
