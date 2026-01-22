import re
import requests

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"

def _headers(access_token: str):
    return {"Authorization": f"Bearer {access_token}"}

def graph_get(access_token: str, url: str, params=None):
    r = requests.get(url, headers=_headers(access_token), params=params, timeout=60)
    if r.status_code >= 400:
        raise RuntimeError(f"GET {url} failed: {r.status_code} {r.text}")
    return r.json()

def graph_post(access_token: str, url: str, json_body=None):
    r = requests.post(
        url,
        headers={**_headers(access_token), "Content-Type": "application/json"},
        json=json_body,
        timeout=60,
    )
    if r.status_code >= 400:
        raise RuntimeError(f"POST {url} failed: {r.status_code} {r.text}")
    return r.json()

def graph_put_bytes(access_token: str, url: str, content_bytes: bytes, content_type="application/octet-stream"):
    r = requests.put(
        url,
        headers={**_headers(access_token), "Content-Type": content_type},
        data=content_bytes,
        timeout=120,
    )
    if r.status_code >= 400:
        raise RuntimeError(f"PUT {url} failed: {r.status_code} {r.text}")
    return r.json()

# ---------------- SharePoint Resolution ----------------

def resolve_site_id(access_token: str, sharepoint_site_url: str) -> str:
    """
    sharepoint_site_url example:
      https://philsagov.sharepoint.com/sites/SMCOD
      https://philsagov.sharepoint.com/teams/SMCOD
    """
    m = re.match(r"^https://([^/]+)(/.*)$", sharepoint_site_url.strip())
    if not m:
        raise ValueError("Invalid SharePoint site URL.")
    host = m.group(1)
    path = m.group(2).rstrip("/")
    url = f"{GRAPH_ROOT}/sites/{host}:{path}"
    data = graph_get(access_token, url)
    return data["id"]

def get_default_drive_id(access_token: str, site_id: str) -> str:
    data = graph_get(access_token, f"{GRAPH_ROOT}/sites/{site_id}/drive")
    return data["id"]

def get_item_by_path(access_token: str, drive_id: str, path: str) -> dict:
    """
    Gets a DriveItem by path relative to drive root.
    """
    path = path.strip("/")
    url = f"{GRAPH_ROOT}/drives/{drive_id}/root:/{path}"
    return graph_get(access_token, url)

def list_children(access_token: str, drive_id: str, item_id: str) -> list[dict]:
    url = f"{GRAPH_ROOT}/drives/{drive_id}/items/{item_id}/children"
    items = []
    next_url = url
    while next_url:
        data = graph_get(access_token, next_url)
        items.extend(data.get("value", []))
        next_url = data.get("@odata.nextLink")
    return items

# ---------------- IR folder logic ----------------

def list_city_folders(access_token: str, drive_id: str, incident_reports_root_path: str, year: str) -> list[str]:
    year_item = get_item_by_path(access_token, drive_id, f"{incident_reports_root_path}/{year}")
    children = list_children(access_token, drive_id, year_item["id"])
    return sorted([c["name"] for c in children if "folder" in c])

def list_ir_folders_in_city(access_token: str, drive_id: str, incident_reports_root_path: str, year: str, city_folder: str) -> list[str]:
    city_item = get_item_by_path(access_token, drive_id, f"{incident_reports_root_path}/{year}/{city_folder}")
    children = list_children(access_token, drive_id, city_item["id"])
    return sorted([c["name"] for c in children if "folder" in c])

def parse_last_serial(ir_folder_names: list[str], site_code: str, year: str) -> int:
    pat = re.compile(rf"^SMCOD-IR-GS-{re.escape(site_code)}-{re.escape(str(year))}-(\d{{4}})$")
    max_n = 0
    for name in ir_folder_names:
        m = pat.match(name.strip())
        if m:
            max_n = max(max_n, int(m.group(1)))
    return max_n

def suggest_next_serial(access_token: str, drive_id: str, incident_reports_root_path: str, year: str, city_folder: str, site_code: str) -> str:
    folders = list_ir_folders_in_city(access_token, drive_id, incident_reports_root_path, year, city_folder)
    last_n = parse_last_serial(folders, site_code, year)
    return f"{last_n + 1:04d}"

def check_duplicate_ir(access_token: str, drive_id: str, incident_reports_root_path: str, year: str, city_folder: str, full_ir_no: str) -> bool:
    folders = list_ir_folders_in_city(access_token, drive_id, incident_reports_root_path, year, city_folder)
    return full_ir_no in set(folders)

# ---------------- Write ops (requires Sites.ReadWrite.All) ----------------

def ensure_folder(access_token: str, drive_id: str, parent_item_id: str, folder_name: str) -> dict:
    children = list_children(access_token, drive_id, parent_item_id)
    for c in children:
        if c.get("name") == folder_name and "folder" in c:
            return c

    url = f"{GRAPH_ROOT}/drives/{drive_id}/items/{parent_item_id}/children"
    body = {"name": folder_name, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"}
    return graph_post(access_token, url, body)

def ensure_path(access_token: str, drive_id: str, root_path: str, parts: list[str]) -> dict:
    root_item = get_item_by_path(access_token, drive_id, root_path)
    cur = root_item
    for p in parts:
        cur = ensure_folder(access_token, drive_id, cur["id"], p)
    return cur

def upload_file_to_folder(access_token: str, drive_id: str, folder_item_id: str, filename: str, content_bytes: bytes, content_type: str):
    url = f"{GRAPH_ROOT}/drives/{drive_id}/items/{folder_item_id}:/{filename}:/content"
    return graph_put_bytes(access_token, url, content_bytes, content_type=content_type)
