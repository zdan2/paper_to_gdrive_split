from __future__ import annotations

import csv
import mimetypes
import os
from pathlib import Path, PurePosixPath
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import dropbox
from google.auth.transport.requests import Request as GoogleRequest
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

GOOGLE_SCOPES = ["https://www.googleapis.com/auth/drive"]
GOOGLE_DOC_MIME = "application/vnd.google-apps.document"
GOOGLE_FOLDER_MIME = "application/vnd.google-apps.folder"

# Useful MIME overrides because mimetypes can vary by platform.
COMMON_MIME_OVERRIDES = {
    ".doc": "application/msword",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".html": "text/html",
    ".htm": "text/html",
    ".rtf": "application/rtf",
    ".txt": "text/plain",
    ".md": "text/markdown",
    ".odt": "application/vnd.oasis.opendocument.text",
    ".pdf": "application/pdf",
}


def normalize_dropbox_path(path: str) -> str:
    path = (path or "").strip()
    if not path:
        return ""
    return path if path.startswith("/") else f"/{path}"


def ensure_parent_dir(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def relative_parts_under_root(dropbox_path: str, root: str) -> Tuple[str, ...]:
    full_path = normalize_dropbox_path(dropbox_path)
    root = normalize_dropbox_path(root)

    if root:
        base = root.rstrip("/")
        if full_path.startswith(base + "/"):
            rel = full_path[len(base) + 1 :]
        else:
            rel = full_path.lstrip("/")
    else:
        rel = full_path.lstrip("/")

    rel_posix = PurePosixPath(rel)
    return rel_posix.parts


def make_local_export_path(
    dropbox_path: str,
    exported_filename: str,
    dropbox_root: str,
    export_root: Path,
) -> Path:
    rel_parts = relative_parts_under_root(dropbox_path, dropbox_root)
    rel_parent = rel_parts[:-1]
    return export_root.joinpath(*rel_parent, exported_filename)

def load_dropbox_client_from_env() -> dropbox.Dropbox:
    token = os.environ.get("DROPBOX_TOKEN")
    if not token:
        raise RuntimeError("環境変数 DROPBOX_TOKEN が必要です。")

    dbx = dropbox.Dropbox(oauth2_access_token=token, timeout=300)

    # team space の root に切り替える
    try:
        from dropbox.common import PathRoot

        acct = dbx.users_get_current_account()
        root_info = getattr(acct, "root_info", None)
        root_namespace_id = getattr(root_info, "root_namespace_id", None)
        if root_namespace_id:
            dbx = dbx.with_path_root(PathRoot.root(root_namespace_id))
    except Exception:
        pass

    return dbx
def iter_paper_paths(dbx: dropbox.Dropbox, root: str) -> Iterable[str]:
    result = dbx.files_list_folder(normalize_dropbox_path(root), recursive=True)
    while True:
        for entry in result.entries:
            if isinstance(entry, dropbox.files.FileMetadata) and entry.name.lower().endswith(".paper"):
                yield entry.path_display or entry.path_lower or entry.name
        if not result.has_more:
            break
        result = dbx.files_list_folder_continue(result.cursor)


def try_get_export_info(dbx: dropbox.Dropbox, path: str) -> Tuple[Optional[str], List[str]]:
    try:
        meta = dbx.files_get_metadata(path)
    except Exception:
        return None, []

    export_info = getattr(meta, "export_info", None)
    if export_info is None:
        return None, []

    default_format = getattr(export_info, "export_as", None)
    options = list(getattr(export_info, "export_options", None) or [])
    return default_format, options


def choose_export_format(default_format: Optional[str], options: Sequence[str], preferred: Optional[str]) -> Optional[str]:
    """
    Dropbox の export_format は実際の値がアカウントやファイル種別で異なることがあるため、
    ここでは完全一致→部分一致の順で緩く選びます。
    選べなければ None を返して Dropbox 側のデフォルト形式に委ねます。
    """
    if not preferred:
        return None

    preferred_lower = preferred.strip().lower()
    candidates = [c for c in options if c]
    if default_format:
        candidates = [default_format, *candidates]

    # 1) exact match
    for value in candidates:
        if value.lower() == preferred_lower:
            return value

    # 2) suffix / token match
    aliases = {
        "docx": ["docx", "wordprocessingml"],
        "html": ["html"],
        "markdown": ["markdown", "md"],
        "md": ["markdown", "md"],
        "pdf": ["pdf"],
        "txt": ["plain", "text/plain", "txt"],
        "rtf": ["rtf"],
    }
    tokens = aliases.get(preferred_lower, [preferred_lower])
    for value in candidates:
        lv = value.lower()
        if any(token in lv for token in tokens):
            return value

    return None


def export_paper_file(
    dbx: dropbox.Dropbox,
    dropbox_path: str,
    local_path: Path,
    export_format: Optional[str] = None,
) -> Tuple[str, Optional[str], List[str]]:
    default_format, options = try_get_export_info(dbx, dropbox_path)
    selected = choose_export_format(default_format, options, export_format)

    if selected:
        export_result, response = dbx.files_export(dropbox_path, export_format=selected)
    else:
        export_result, response = dbx.files_export(dropbox_path)

    # Dropbox returns the exported filename with the new extension.
    exported_name = export_result.export_metadata.name
    local_path = local_path.with_name(exported_name)
    ensure_parent_dir(local_path)
    with local_path.open("wb") as f:
        f.write(response.content)
    response.close()
    return str(local_path), selected or default_format, list(options)


def build_drive_service(credentials_path: str, token_cache_path: str):
    creds: Optional[Credentials] = None

    if os.path.exists(token_cache_path):
        creds = Credentials.from_authorized_user_file(token_cache_path, GOOGLE_SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(GoogleRequest())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, GOOGLE_SCOPES)
            creds = flow.run_local_server(port=0)

        with open(token_cache_path, "w", encoding="utf-8") as f:
            f.write(creds.to_json())

    return build("drive", "v3", credentials=creds, cache_discovery=False)


def guess_mime_type(path: Path) -> str:
    ext = path.suffix.lower()
    if ext in COMMON_MIME_OVERRIDES:
        return COMMON_MIME_OVERRIDES[ext]
    mime, _ = mimetypes.guess_type(str(path))
    return mime or "application/octet-stream"


def load_drive_import_formats(service) -> Dict[str, List[str]]:
    about = service.about().get(fields="importFormats").execute()
    raw = about.get("importFormats", {})
    converted: Dict[str, List[str]] = {}
    for source_mime, targets in raw.items():
        if isinstance(targets, list):
            converted[source_mime] = [str(x) for x in targets]
        elif isinstance(targets, dict) and "items" in targets and isinstance(targets["items"], list):
            converted[source_mime] = [str(x) for x in targets["items"]]
        else:
            converted[source_mime] = [str(targets)]
    return converted


def can_convert_to_google_doc(source_mime: str, import_formats: Dict[str, List[str]]) -> bool:
    targets = import_formats.get(source_mime, [])
    return GOOGLE_DOC_MIME in targets


def iter_local_files(root: Path) -> Iterable[Path]:
    for path in root.rglob("*"):
        if path.is_file():
            yield path


def relative_parts_under_local_root(path: Path, root: Path) -> Tuple[str, ...]:
    rel = path.relative_to(root)
    return rel.parts


def drive_query_escape(value: str) -> str:
    return value.replace("\\", "\\\\").replace("'", "\\'")


def ensure_drive_folder(service, folder_name: str, parent_id: Optional[str], cache: Dict[Tuple[str, str], str]) -> str:
    cache_key = (parent_id or "ROOT", folder_name)
    if cache_key in cache:
        return cache[cache_key]

    q = (
        "trashed = false and "
        f"mimeType = '{GOOGLE_FOLDER_MIME}' and "
        f"name = '{drive_query_escape(folder_name)}'"
    )
    if parent_id:
        q += f" and '{parent_id}' in parents"

    resp = service.files().list(
        q=q,
        spaces="drive",
        fields="files(id, name)",
        pageSize=10,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    files = resp.get("files", [])

    if files:
        folder_id = files[0]["id"]
    else:
        body = {"name": folder_name, "mimeType": GOOGLE_FOLDER_MIME}
        if parent_id:
            body["parents"] = [parent_id]
        folder_id = service.files().create(
            body=body,
            fields="id",
            supportsAllDrives=True,
        ).execute()["id"]

    cache[cache_key] = folder_id
    return folder_id


def ensure_drive_path_for_local_parent(
    service,
    local_file: Path,
    local_root: Path,
    drive_root_folder_id: Optional[str],
    folder_cache: Dict[Tuple[str, str], str],
) -> Optional[str]:
    parts = relative_parts_under_local_root(local_file, local_root)
    parent_id = drive_root_folder_id
    for folder_name in parts[:-1]:
        parent_id = ensure_drive_folder(service, folder_name, parent_id, folder_cache)
    return parent_id


def find_existing_drive_file(service, name: str, parent_id: Optional[str], mime_type: Optional[str] = None) -> Optional[dict]:
    q = f"trashed = false and name = '{drive_query_escape(name)}'"
    if parent_id:
        q += f" and '{parent_id}' in parents"
    if mime_type:
        q += f" and mimeType = '{mime_type}'"

    resp = service.files().list(
        q=q,
        spaces="drive",
        fields="files(id, name, mimeType, webViewLink)",
        pageSize=10,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    files = resp.get("files", [])
    return files[0] if files else None


def write_csv_rows(csv_path: Path, rows: List[dict]) -> None:
    ensure_parent_dir(csv_path)
    if not rows:
        csv_path.write_text("", encoding="utf-8")
        return
    fieldnames = list(rows[0].keys())
    with csv_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
