from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path
from typing import List

from pipeline_common import (
    ensure_dir,
    export_paper_file,
    iter_paper_paths,
    load_dropbox_client_from_env,
    make_local_export_path,
    write_csv_rows,
)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Dropbox Paper(.paper) をローカル確認用フォルダへ書き出します。"
    )
    parser.add_argument("--dropbox-root", default="", help="Dropbox 側の走査開始パス。例: /議事録")
    parser.add_argument(
        "--export-root",
        default="./review_workspace/exported",
        help="書き出し先ローカルフォルダ。ここを人の目で確認します。",
    )
    parser.add_argument(
        "--preferred-format",
        default="",
        help="希望エクスポート形式。例: docx / html。見つからなければ Dropbox の既定形式を使います。",
    )
    parser.add_argument(
        "--skip-existing",
        action="store_true",
        help="同名のローカルファイルが既にあればスキップ",
    )
    parser.add_argument("--sleep", type=float, default=0.1, help="各ファイル間の待機秒")
    parser.add_argument(
        "--manifest",
        default="./review_workspace/export_manifest.csv",
        help="書き出し結果の CSV 保存先",
    )
    args = parser.parse_args()

    export_root = Path(args.export_root).resolve()
    manifest_path = Path(args.manifest).resolve()
    ensure_dir(export_root)

    try:
        dbx = load_dropbox_client_from_env()
    except Exception as e:
        print(f"Dropbox 認証に失敗しました: {e}", file=sys.stderr)
        return 2

    paper_paths = list(iter_paper_paths(dbx, args.dropbox_root))
    print(f"検出件数: {len(paper_paths)} 件")

    rows: List[dict] = []
    ok = 0
    ng = 0

    for idx, dropbox_path in enumerate(paper_paths, start=1):
        try:
            tentative_local_path = make_local_export_path(
                dropbox_path=dropbox_path,
                exported_filename=Path(dropbox_path).stem,
                dropbox_root=args.dropbox_root,
                export_root=export_root,
            )

            # 既存スキップは "同じ stem の何か" があるかでざっくり判定
            if args.skip_existing and tentative_local_path.parent.exists():
                stem = tentative_local_path.name
                existing = list(tentative_local_path.parent.glob(stem + ".*"))
                if existing:
                    print(f"[{idx}/{len(paper_paths)}] SKIP {dropbox_path}")
                    rows.append(
                        {
                            "dropbox_path": dropbox_path,
                            "local_path": str(existing[0]),
                            "status": "skipped_existing",
                            "selected_export_format": "",
                            "available_export_options": "",
                        }
                    )
                    continue

            local_path, selected_format, options = export_paper_file(
                dbx=dbx,
                dropbox_path=dropbox_path,
                local_path=tentative_local_path,
                export_format=args.preferred_format or None,
            )

            print(f"[{idx}/{len(paper_paths)}] OK   {dropbox_path} -> {local_path}")
            rows.append(
                {
                    "dropbox_path": dropbox_path,
                    "local_path": local_path,
                    "status": "exported",
                    "selected_export_format": selected_format or "",
                    "available_export_options": " | ".join(options),
                }
            )
            ok += 1
            time.sleep(max(args.sleep, 0.0))
        except Exception as e:
            print(f"[{idx}/{len(paper_paths)}] ERR  {dropbox_path}: {e}", file=sys.stderr)
            rows.append(
                {
                    "dropbox_path": dropbox_path,
                    "local_path": "",
                    "status": f"error: {e}",
                    "selected_export_format": "",
                    "available_export_options": "",
                }
            )
            ng += 1

    write_csv_rows(manifest_path, rows)
    print()
    print(f"完了: 成功 {ok} / 失敗 {ng}")
    print(f"確認用フォルダ: {export_root}")
    print(f"マニフェスト: {manifest_path}")
    print("次の手順: 確認して問題ないファイルだけを approved フォルダへ移動またはコピーしてください。")
    return 0 if ng == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
