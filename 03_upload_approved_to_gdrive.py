from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from googleapiclient.http import MediaFileUpload

from pipeline_common import (
    GOOGLE_DOC_MIME,
    build_drive_service,
    can_convert_to_google_doc,
    ensure_drive_path_for_local_parent,
    find_existing_drive_file,
    guess_mime_type,
    iter_local_files,
    load_drive_import_formats,
    write_csv_rows,
)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="approved フォルダのファイルを Google Drive へアップロードします。必要なら Google ドキュメントへ変換します。"
    )
    parser.add_argument(
        "--approved-root",
        default="./review_workspace/approved",
        help="アップロード対象フォルダ",
    )
    parser.add_argument(
        "--drive-folder-id",
        default="",
        help="保存先 Google Drive フォルダ ID。未指定ならマイドライブ直下。",
    )
    parser.add_argument(
        "--credentials",
        default="./credentials.json",
        help="Google OAuth client JSON",
    )
    parser.add_argument(
        "--token-cache",
        default="./token.json",
        help="Google OAuth トークン保存先",
    )
    parser.add_argument(
        "--skip-existing",
        action="store_true",
        help="保存先に同名ファイルがあればスキップ",
    )
    parser.add_argument(
        "--sleep",
        type=float,
        default=0.1,
        help="各ファイル間の待機秒",
    )
    parser.add_argument(
        "--result-csv",
        default="./review_workspace/upload_result.csv",
        help="アップロード結果 CSV",
    )
    args = parser.parse_args()

    approved_root = Path(args.approved_root).resolve()
    result_csv = Path(args.result_csv).resolve()
    if not approved_root.exists():
        print(f"approved フォルダが見つかりません: {approved_root}", file=sys.stderr)
        return 2

    try:
        service = build_drive_service(args.credentials, args.token_cache)
        import_formats = load_drive_import_formats(service)
    except Exception as e:
        print(f"Google 認証に失敗しました: {e}", file=sys.stderr)
        return 2

    folder_cache: Dict[Tuple[str, str], str] = {}
    rows: List[dict] = []
    files = list(iter_local_files(approved_root))
    ok = 0
    ng = 0

    for idx, local_file in enumerate(files, start=1):
        try:
            parent_id = ensure_drive_path_for_local_parent(
                service=service,
                local_file=local_file,
                local_root=approved_root,
                drive_root_folder_id=args.drive_folder_id or None,
                folder_cache=folder_cache,
            )

            source_mime = guess_mime_type(local_file)
            convertible = can_convert_to_google_doc(source_mime, import_formats)
            target_name = local_file.stem if convertible else local_file.name
            target_mime = GOOGLE_DOC_MIME if convertible else source_mime

            if args.skip_existing:
                existing = find_existing_drive_file(
                    service=service,
                    name=target_name,
                    parent_id=parent_id,
                    mime_type=target_mime if convertible else None,
                )
                if existing:
                    print(f"[{idx}/{len(files)}] SKIP {local_file}")
                    rows.append(
                        {
                            "local_path": str(local_file),
                            "status": "skipped_existing",
                            "source_mime": source_mime,
                            "uploaded_mime": existing.get("mimeType", ""),
                            "drive_id": existing.get("id", ""),
                            "web_view_link": existing.get("webViewLink", ""),
                        }
                    )
                    continue

            media = MediaFileUpload(
                str(local_file),
                mimetype=source_mime,
                resumable=local_file.stat().st_size > 5 * 1024 * 1024,
            )

            body = {"name": target_name}
            if convertible:
                body["mimeType"] = GOOGLE_DOC_MIME
            if parent_id:
                body["parents"] = [parent_id]

            created = service.files().create(
                body=body,
                media_body=media,
                fields="id, name, mimeType, webViewLink",
                supportsAllDrives=True,
            ).execute()

            print(
                f"[{idx}/{len(files)}] OK   {local_file} -> {created.get('webViewLink', created['id'])}"
            )
            rows.append(
                {
                    "local_path": str(local_file),
                    "status": "uploaded",
                    "source_mime": source_mime,
                    "uploaded_mime": created.get("mimeType", ""),
                    "drive_id": created.get("id", ""),
                    "web_view_link": created.get("webViewLink", ""),
                }
            )
            ok += 1
            time.sleep(max(args.sleep, 0.0))
        except Exception as e:
            print(f"[{idx}/{len(files)}] ERR  {local_file}: {e}", file=sys.stderr)
            rows.append(
                {
                    "local_path": str(local_file),
                    "status": f"error: {e}",
                    "source_mime": "",
                    "uploaded_mime": "",
                    "drive_id": "",
                    "web_view_link": "",
                }
            )
            ng += 1

    write_csv_rows(result_csv, rows)
    print()
    print(f"完了: 成功 {ok} / 失敗 {ng}")
    print(f"結果 CSV: {result_csv}")
    return 0 if ng == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
