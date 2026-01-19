import base64
import re
import requests

GRAPH = "https://graph.microsoft.com/v1.0"

def share_link_to_share_id(url: str) -> str:
    b = base64.urlsafe_b64encode(url.encode("utf-8")).decode("utf-8").rstrip("=")
    return "u!" + b

def _headers(token: str):
    return {"Authorization": f"Bearer {token}"}

def graph_get(url: str, token: str, params=None):
    r = requests.get(url, headers=_headers(token), params=params, timeout=30)
    r.raise_for_status()
    return r.json()

def graph_post(url: str, token: str, body: dict):
    r = requests.post(
        url,
        headers={**_headers(token), "Content-Type": "application/json"},
        json=body,
        timeout=30,
    )
    r.raise_for_status()
    return r.json()

def graph_put(url: str, token: str, data: bytes, content_type: str):
    r = requests.put(
        url,
        headers={**_headers(token), "Content-Type": content_type},
        data=data,
        timeout=120,
    )
    r.raise_for_status()
    return r.json()

def resolve_root_folder_from_share_link(share_link: str, token: str) -> tuple[str, str]:
    """
    Returns (drive_id, root_item_id) for the shared folder link.
    """
    share_id = share_link_to_share_id(share_link)
    item = graph_get(f"{GRAPH}/shares/{share_id}/driveItem", token)
    root_item_id = item["id"]
    drive_id = item["parentReference"]["driveId"]
    return drive_id, root_item_id

def list_children(drive_id: str, item_id: str, token: str) -> list[dict]:
    data = graph_get(f"{GRAPH}/drives/{drive_id}/items/{item_id}/children", token)
    return data.get("value", [])

def find_child_folder(children: list[dict], folder_name: str) -> dict | None:
    for it in children:
        if it.get("name") == folder_name and "folder" in it:
            return it
    return None

def create_folder(drive_id: str, parent_item_id: str, folder_name: str, token: str) -> dict:
    url = f"{GRAPH}/drives/{drive_id}/items/{parent_item_id}/children"
    body = {
        "name": folder_name,
        "folder": {},
        "@microsoft.graph.conflictBehavior": "fail",
    }
    return graph_post(url, token, body)

def ensure_folder(drive_id: str, parent_item_id: str, folder_name: str, token: str, create_if_missing=True) -> dict:
    children = list_children(drive_id, parent_item_id, token)
    found = find_child_folder(children, folder_name)
    if found:
        return found
    if not create_if_missing:
        raise FileNotFoundError(f"Folder not found: {folder_name}")
    return create_folder(drive_id, parent_item_id, folder_name, token)

def compute_next_incident_id(children: list[dict], year: int, site_code: str) -> str:
    """
    children here are items under the CITY folder. Incidents are folders named:
    SMCOD-IR-GS-DVO-2025-0001
    """
    pat = re.compile(rf"^SMCOD-IR-GS-{re.escape(site_code)}-{year}-(\d{{4}})$", re.IGNORECASE)
    nums = []
    for it in children:
        if "folder" not in it:
            continue
        name = it.get("name", "")
        m = pat.match(name)
        if m:
            nums.append(int(m.group(1)))
    nxt = (max(nums) + 1) if nums else 1
    return f"SMCOD-IR-GS-{site_code}-{year}-{nxt:04d}"

def ensure_incident_folder(drive_id: str, city_folder_id: str, incident_id: str, token: str) -> dict:
    """
    Creates incident folder under the CITY folder. If exists, raises to prevent overwrite.
    """
    return create_folder(drive_id, city_folder_id, incident_id, token)

def upload_file_to_folder(drive_id: str, folder_item_id: str, filename: str, file_bytes: bytes, token: str) -> dict:
    """
    Upload into folder using simple PUT. Good for small/medium files.
    If you expect huge PDFs/images, ask me for upload-session chunking.
    """
    if filename.lower().endswith(".docx"):
        ctype = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    elif filename.lower().endswith(".pdf"):
        ctype = "application/pdf"
    elif filename.lower().endswith((".jpg", ".jpeg")):
        ctype = "image/jpeg"
    elif filename.lower().endswith(".png"):
        ctype = "image/png"
    else:
        ctype = "application/octet-stream"

    url = f"{GRAPH}/drives/{drive_id}/items/{folder_item_id}:/{filename}:/content"
    return graph_put(url, token, file_bytes, ctype)
