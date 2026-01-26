import requests

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def _headers(token: str, extra: dict | None = None):
    h = {"Authorization": f"Bearer {token}"}
    if extra:
        h.update(extra)
    return h


def resolve_site_id(token: str, site_url: str) -> str:
    """
    site_url examples:
      https://philsagov.sharepoint.com/sites/SMCOD
      https://philsagov.sharepoint.com/teams/SMCOD
    """
    # Convert full URL to Graph "hostname:/path" format
    # e.g. hostname=philsagov.sharepoint.com, path=/sites/SMCOD
    site_url = site_url.rstrip("/")
    if "://" in site_url:
        site_url = site_url.split("://", 1)[1]
    host, path = site_url.split("/", 1)
    path = "/" + path

    url = f"{GRAPH_BASE}/sites/{host}:{path}"
    r = requests.get(url, headers=_headers(token), timeout=60)
    r.raise_for_status()
    return r.json()["id"]


def get_default_drive_id(token: str, site_id: str) -> str:
    url = f"{GRAPH_BASE}/sites/{site_id}/drive"
    r = requests.get(url, headers=_headers(token), timeout=60)
    r.raise_for_status()
    return r.json()["id"]


def _item_by_path(token: str, drive_id: str, path: str):
    """
    Get item metadata for /root:/path:
    """
    path = path.strip("/")
    url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{path}"
    r = requests.get(url, headers=_headers(token), timeout=60)
    if r.status_code == 404:
        return None
    r.raise_for_status()
    return r.json()


def _children(token: str, drive_id: str, folder_item_id: str):
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{folder_item_id}/children"
    r = requests.get(url, headers=_headers(token), timeout=60)
    r.raise_for_status()
    return r.json().get("value", [])


def ensure_folder(token: str, drive_id: str, parent_item_id: str, folder_name: str) -> dict:
    """
    Ensure folder exists under parent item by name. Returns folder item.
    """
    kids = _children(token, drive_id, parent_item_id)
    for k in kids:
        if k.get("name") == folder_name and k.get("folder") is not None:
            return k

    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{parent_item_id}/children"
    payload = {"name": folder_name, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"}
    r = requests.post(url, headers=_headers(token, {"Content-Type": "application/json"}), json=payload, timeout=60)
    if r.status_code == 409:
        # created by someone else between check and create
        kids = _children(token, drive_id, parent_item_id)
        for k in kids:
            if k.get("name") == folder_name and k.get("folder") is not None:
                return k
    r.raise_for_status()
    return r.json()


def ensure_path(token: str, drive_id: str, root_path: str, parts: list[str]) -> dict:
    """
    Ensure /root_path/part1/part2/... exists.
    Returns the final folder item (dict).
    """
    root_item = _item_by_path(token, drive_id, root_path)
    if not root_item:
        raise RuntimeError(f"Root path not found in drive: {root_path}")

    current = root_item
    for name in parts:
        current = ensure_folder(token, drive_id, current["id"], name)
    return current


def upload_file_to_folder(token: str, drive_id: str, folder_item_id: str, filename: str, content_bytes: bytes, content_type: str):
    """
    Upload (overwrite) a file into a folder.
    """
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{folder_item_id}:/{filename}:/content"
    r = requests.put(url, headers=_headers(token, {"Content-Type": content_type}), data=content_bytes, timeout=120)
    r.raise_for_status()
    return r.json()


def check_duplicate_ir(token: str, drive_id: str, root_path: str, year: str, city: str, incident_folder_name: str) -> bool:
    """
    True if folder exists at /root_path/year/city/incident_folder_name
    """
    path = f"{root_path}/{year}/{city}/{incident_folder_name}"
    item = _item_by_path(token, drive_id, path)
    return bool(item and item.get("folder") is not None)


def list_text_files(token: str, drive_id: str, folder_path: str) -> list[dict]:
    """
    List text-ish files in the folder path (non-recursive).
    """
    folder = _item_by_path(token, drive_id, folder_path)
    if not folder or folder.get("folder") is None:
        raise RuntimeError(f"Folder not found: {folder_path}")

    kids = _children(token, drive_id, folder["id"])
    allowed = (".txt", ".md", ".log", ".csv")
    out = []
    for k in kids:
        name = k.get("name", "")
        if k.get("file") and name.lower().endswith(allowed):
            out.append({"id": k["id"], "name": name})
    return sorted(out, key=lambda x: x["name"].lower())


def download_file_text(token: str, drive_id: str, file_item_id: str) -> str:
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{file_item_id}/content"
    r = requests.get(url, headers=_headers(token), timeout=120)
    r.raise_for_status()
    # Assume utf-8 text
    return r.content.decode("utf-8", errors="replace")


def update_file_text(token: str, drive_id: str, file_item_id: str, new_text: str):
    """
    Overwrite an existing file's contents (text).
    """
    # Need file name and parent folder to PUT to the correct path
    meta_url = f"{GRAPH_BASE}/drives/{drive_id}/items/{file_item_id}"
    meta = requests.get(meta_url, headers=_headers(token), timeout=60)
    meta.raise_for_status()
    meta = meta.json()

    parent_id = meta["parentReference"]["id"]
    filename = meta["name"]

    content_bytes = (new_text or "").encode("utf-8")
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{parent_id}:/{filename}:/content"
    r = requests.put(url, headers=_headers(token, {"Content-Type": "text/plain; charset=utf-8"}), data=content_bytes, timeout=120)
    r.raise_for_status()
    return r.json()
