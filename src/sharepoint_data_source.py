from __future__ import annotations

import hashlib
import json
import os
import re
import time
from dataclasses import dataclass
from pathlib import Path

import requests


FILE_PATTERN = re.compile(
    r"^Solutions_Revenue_(?P<month>[A-Za-z]+)_(?P<year>\d{4}).*\.xlsx$",
    re.IGNORECASE,
)


class DataSourceError(RuntimeError):
    pass


@dataclass(frozen=True)
class SharePointRuntimeConfig:
    tenant_id: str
    client_id: str
    client_secret: str
    site_id: str
    drive_id: str
    folder_path: str
    template_file: str

    @classmethod
    def from_env(cls) -> "SharePointRuntimeConfig":
        cfg = cls(
            tenant_id=os.getenv("RAS_SP_TENANT_ID", "").strip(),
            client_id=os.getenv("RAS_SP_CLIENT_ID", "").strip(),
            client_secret=os.getenv("RAS_SP_CLIENT_SECRET", "").strip(),
            site_id=os.getenv("RAS_SP_SITE_ID", "").strip(),
            drive_id=os.getenv("RAS_SP_DRIVE_ID", "").strip(),
            folder_path=os.getenv("RAS_SP_FOLDER_PATH", "").strip().strip("/"),
            template_file=os.getenv("RAS_SP_TEMPLATE_FILE", "Revenue Output Dashboard Sample.xlsx").strip(),
        )
        missing = [
            name
            for name, value in {
                "RAS_SP_TENANT_ID": cfg.tenant_id,
                "RAS_SP_CLIENT_ID": cfg.client_id,
                "RAS_SP_CLIENT_SECRET": cfg.client_secret,
                "RAS_SP_SITE_ID": cfg.site_id,
                "RAS_SP_DRIVE_ID": cfg.drive_id,
                "RAS_SP_FOLDER_PATH": cfg.folder_path,
                "RAS_SP_TEMPLATE_FILE": cfg.template_file,
            }.items()
            if not value
        ]
        if missing:
            raise DataSourceError(f"Missing SharePoint configuration: {', '.join(missing)}")
        return cfg


@dataclass(frozen=True)
class MirrorResult:
    data_folder: Path
    template_file: Path
    downloaded_revenue_files: int


RETRYABLE_STATUS_CODES = {429, 500, 502, 503, 504}
MAX_ATTEMPTS = 4


def _mirror_root() -> Path:
    if os.name == "nt":
        base = Path(os.getenv("LOCALAPPDATA", str(Path.home() / "AppData" / "Local")))
    else:
        base = Path(os.getenv("XDG_CACHE_HOME", str(Path.home() / ".cache")))
    root = base / "revenue-analysis-dashboard" / "sharepoint-mirror"
    root.mkdir(parents=True, exist_ok=True)
    return root


def _folder_signature(config: SharePointRuntimeConfig) -> str:
    payload = "|".join(
        [
            config.tenant_id,
            config.site_id,
            config.drive_id,
            config.folder_path,
            config.template_file,
        ]
    )
    return hashlib.sha1(payload.encode("utf-8")).hexdigest()[:12]


def _token(config: SharePointRuntimeConfig) -> str:
    token_url = f"https://login.microsoftonline.com/{config.tenant_id}/oauth2/v2.0/token"
    payload = {
        "client_id": config.client_id,
        "scope": "https://graph.microsoft.com/.default",
        "client_secret": config.client_secret,
        "grant_type": "client_credentials",
    }
    response = requests.post(token_url, data=payload, timeout=30)
    if not response.ok:
        raise DataSourceError(f"Unable to get Graph token ({response.status_code}): {response.text[:400]}")
    body = response.json()
    token = body.get("access_token", "")
    if not token:
        raise DataSourceError("Graph token response did not include access_token.")
    return token


def _graph_get(url: str, token: str) -> dict:
    headers = {"Authorization": f"Bearer {token}"}
    response = _request_with_retry("get", url, headers=headers, timeout=60)
    return response.json()


def _graph_download(url: str, token: str, output_file: Path) -> None:
    headers = {"Authorization": f"Bearer {token}"}
    response = _request_with_retry("get", url, headers=headers, stream=True, timeout=120)
    with output_file.open("wb") as handle:
        for chunk in response.iter_content(chunk_size=1024 * 1024):
            if chunk:
                handle.write(chunk)


def _request_with_retry(method: str, url: str, **kwargs) -> requests.Response:
    last_error = ""
    for attempt in range(1, MAX_ATTEMPTS + 1):
        response = requests.request(method, url, **kwargs)
        if response.ok:
            return response

        status = response.status_code
        last_error = f"HTTP {status}: {response.text[:400]}"
        if status not in RETRYABLE_STATUS_CODES or attempt == MAX_ATTEMPTS:
            raise DataSourceError(f"Graph call failed ({status}): {response.text[:400]}")

        retry_after = response.headers.get("Retry-After")
        wait_seconds = float(retry_after) if retry_after and retry_after.isdigit() else 2 ** (attempt - 1)
        time.sleep(wait_seconds)

    raise DataSourceError(f"Graph call failed after retries: {last_error}")


def _manifest_path(folder: Path) -> Path:
    return folder / "mirror_manifest.json"


def _load_manifest(folder: Path) -> dict[str, dict[str, str]]:
    path = _manifest_path(folder)
    if not path.exists():
        return {}
    try:
        with path.open("r", encoding="utf-8") as handle:
            data = json.load(handle)
            if isinstance(data, dict):
                return {str(k): dict(v) for k, v in data.items() if isinstance(v, dict)}
    except Exception:
        return {}
    return {}


def _save_manifest(folder: Path, data: dict[str, dict[str, str]]) -> None:
    with _manifest_path(folder).open("w", encoding="utf-8") as handle:
        json.dump(data, handle, indent=2, sort_keys=True)


def _item_signature(item: dict) -> str:
    etag = str(item.get("eTag", ""))
    modified = str(item.get("lastModifiedDateTime", ""))
    size = str(item.get("size", ""))
    return "|".join([etag, modified, size])


def _list_drive_children(config: SharePointRuntimeConfig, token: str) -> list[dict]:
    url = (
        f"https://graph.microsoft.com/v1.0/sites/{config.site_id}/drives/{config.drive_id}"
        f"/root:/{config.folder_path}:/children?$top=999"
    )
    rows: list[dict] = []
    while url:
        body = _graph_get(url, token)
        rows.extend(body.get("value", []))
        url = body.get("@odata.nextLink")
    return rows


def mirror_sharepoint_files(config: SharePointRuntimeConfig) -> MirrorResult:
    token = _token(config)
    children = _list_drive_children(config, token)

    folder = _mirror_root() / _folder_signature(config)
    monthly_dir = folder / "monthly"
    monthly_dir.mkdir(parents=True, exist_ok=True)
    existing_manifest = _load_manifest(folder)
    updated_manifest: dict[str, dict[str, str]] = {}

    monthly_count = 0
    template_found = False

    for item in children:
        if "file" not in item:
            continue

        file_name = str(item.get("name", ""))
        item_id = str(item.get("id", ""))
        if not file_name or not item_id:
            continue

        target = None
        if FILE_PATTERN.match(file_name):
            target = monthly_dir / file_name
            monthly_count += 1
        elif file_name.lower() == config.template_file.lower():
            target = folder / config.template_file
            template_found = True

        if target is None:
            continue

        signature = _item_signature(item)
        manifest_key = str(target)
        previous = existing_manifest.get(manifest_key, {})
        updated_manifest[manifest_key] = {"signature": signature}

        if previous.get("signature") == signature and target.exists():
            continue

        download_url = (
            f"https://graph.microsoft.com/v1.0/sites/{config.site_id}/drives/{config.drive_id}"
            f"/items/{item_id}/content"
        )
        _graph_download(download_url, token, target)

    if monthly_count == 0:
        raise DataSourceError(
            "No monthly files matched pattern Solutions_Revenue_<Month>_<Year>*.xlsx in configured SharePoint folder."
        )
    if not template_found:
        raise DataSourceError(
            f"Template file '{config.template_file}' not found in configured SharePoint folder."
        )

    _save_manifest(folder, updated_manifest)

    return MirrorResult(
        data_folder=monthly_dir,
        template_file=folder / config.template_file,
        downloaded_revenue_files=monthly_count,
    )
