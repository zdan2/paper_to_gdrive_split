from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import List

from pipeline_common import (
    GOOGLE_DOC_MIME,
    build_drive_service,
    can_convert_to_google_doc,
    guess_mime_type,
    iter_local_files,
    load_drive_import_formats,
    write_csv_rows,
)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="approved フォルダを走査し、Google Drive へどうアップロードされるかの計画表を作ります。"
    )
    parser.add_argument(
        "--approved-root",
        default="./review_workspace/approved",
        help="人の確認後にアップロード対象を置くローカルフォルダ",
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
        "--plan-csv",
        default="./review_workspace/upload_plan.csv",
        help="出力する計画 CSV",
    )
    args = parser.parse_args()

    approved_root = Path(args.approved_root).resolve()
    plan_csv = Path(args.plan_csv).resolve()

    if not approved_root.exists():
        print(f"approved フォルダが見つかりません: {approved_root}", file=sys.stderr)
        return 2

    try:
        service = build_drive_service(args.credentials, args.token_cache)
        import_formats = load_drive_import_formats(service)
    except Exception as e:
        print(f"Google 認証または importFormats の取得に失敗しました: {e}", file=sys.stderr)
        return 2

    rows: List[dict] = []
    count = 0
    for local_file in iter_local_files(approved_root):
        count += 1
        mime_type = guess_mime_type(local_file)
        convertible = can_convert_to_google_doc(mime_type, import_formats)
        rows.append(
            {
                "local_path": str(local_file),
                "mime_type": mime_type,
                "will_convert_to_google_doc": "yes" if convertible else "no",
                "target_mime_type": GOOGLE_DOC_MIME if convertible else mime_type,
            }
        )

    write_csv_rows(plan_csv, rows)
    print(f"対象ファイル: {count} 件")
    print(f"計画表を書き出しました: {plan_csv}")
    print("yes のものは Google ドキュメントへ変換、no のものは通常ファイルとして Drive に置かれます。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
