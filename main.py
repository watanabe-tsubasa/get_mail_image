# save_mail_attachments.py
# -*- coding: utf-8 -*-
import argparse
import os
import re
import sys
from datetime import datetime
try:
    import win32com.client  # pywin32
except ImportError:
    print("pywin32 が見つかりません。`pip install pywin32` を実行してください。")
    sys.exit(1)

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp", ".heic"}

def sanitize_filename(name: str) -> str:
    # Windows 禁止文字を置換
    return re.sub(r'[\\/:*?"<>|]', "_", name)

def ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)

def connect_outlook():
    # Outlook（MAPI）へ接続
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    return namespace

def get_folder(namespace, store_name: str | None, folder_path: str):
    """
    folder_path 例:
      - "Inbox"（受信トレイ）
      - "Inbox/サブフォルダ"
      - "受信トレイ/請求書"
    store_name を指定すると別メールボックス（共有/追加）も選択可能。
    """
    if store_name:
        root = None
        for f in namespace.Folders:
            if f.Name == store_name:
                root = f
                break
        if root is None:
            raise RuntimeError(f"メールボックス '{store_name}' が見つかりません。")
    else:
        root = namespace.GetDefaultFolder(6).Parent  # 6 = olFolderInbox のストアルート

    parts = folder_path.replace("\\", "/").split("/")
    current = root
    for p in parts:
        if not p:
            continue
        if p.lower() in ("inbox", "受信トレイ"):  # 言語差吸収
            current = namespace.GetDefaultFolder(6) if current == root else current.Folders[p]
        else:
            current = current.Folders[p]
    return current

def iter_target_mails(items, subject: str, exact: bool):
    # 新しい順で安定取得（Outlook Items は並びに敏感）
    items.Sort("[ReceivedTime]", True)
    for item in items:
        # メール以外(会議通知など)をスキップ
        if item.Class != 43:  # 43 = olMail
            continue
        subj = (item.Subject or "")
        if exact:
            if subj == subject:
                yield item
        else:
            if subject.lower() in subj.lower():
                yield item

def save_attachments_from_mail(mail, outdir: str, images_only: bool):
    saved = []
    # メールごとにサブフォルダを切ると管理しやすい（任意）
    mail_dir = os.path.join(
        outdir,
        sanitize_filename(f"{mail.ReceivedTime:%Y%m%d_%H%M%S}_{mail.EntryID[-8:]}")
    )
    ensure_dir(mail_dir)

    for att in mail.Attachments:
        # ファイル名
        name = sanitize_filename(att.FileName or "attachment")
        root, ext = os.path.splitext(name)
        ext_lower = ext.lower()

        if images_only and ext_lower not in IMAGE_EXTS:
            # 画像のみ保存オプション
            continue

        # 同名回避
        target = os.path.join(mail_dir, name)
        i = 1
        while os.path.exists(target):
            target = os.path.join(mail_dir, f"{root}({i}){ext}")
            i += 1

        att.SaveAsFile(target)
        saved.append(target)

    return saved

def main():
    parser = argparse.ArgumentParser(
        description="Outlook(Win32 MAPI)から件名でメールを指定し、添付画像を一括保存します。"
    )
    parser.add_argument(
        "-s", "--subject", required=False,
        help="対象メールの件名（部分一致）。未指定の場合は対話的に入力します。"
    )
    parser.add_argument(
        "-o", "--outdir", required=True,
        help="保存先フォルダのパス（例: C:\\temp\\mail_attachments）"
    )
    parser.add_argument(
        "-f", "--folder", default="Inbox",
        help="検索対象フォルダ（例: 'Inbox', 'Inbox/請求書' など）"
    )
    parser.add_argument(
        "--store", default=None,
        help="メールボックス名（共有/別アカウント等を使う場合に指定。未指定で既定のメールボックス）"
    )
    parser.add_argument(
        "--exact", action="store_true",
        help="件名を完全一致にする（デフォルトは部分一致）"
    )
    parser.add_argument(
        "--all", action="store_true",
        help="画像以外の添付も含めて全て保存（デフォルトは画像のみ）"
    )
    args = parser.parse_args()

    subject = args.subject
    if not subject:
        try:
            subject = input("対象メールの件名（部分一致可）を入力してください: ").strip()
        except KeyboardInterrupt:
            print("\nキャンセルしました。")
            return
    if not subject:
        print("件名が空です。終了します。")
        return

    ensure_dir(args.outdir)

    print("Outlook に接続中...")
    ns = connect_outlook()
    print(f"フォルダ取得中: store={args.store or '(既定)'} / {args.folder}")
    folder = get_folder(ns, args.store, args.folder)

    print(f"メール検索中: 件名{'(完全一致)' if args.exact else '(部分一致)'}='{subject}'")
    matched = list(iter_target_mails(folder.Items, subject, args.exact))
    if not matched:
        print("該当するメールが見つかりませんでした。")
        return

    total_saved = 0
    for mail in matched:
        saved_paths = save_attachments_from_mail(mail, args.outdir, images_only=not args.all)
        if saved_paths:
            print(f"保存: {len(saved_paths)} 件 / 受信日時: {mail.ReceivedTime} / 件名: {mail.Subject}")
            for p in saved_paths:
                print(f"  - {p}")
            total_saved += len(saved_paths)

    print(f"完了: 添付保存合計 {total_saved} 件")

if __name__ == "__main__":
    main()
