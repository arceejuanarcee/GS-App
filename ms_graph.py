import os
import time
import json
import base64
import streamlit as st
import msal
import streamlit.components.v1 as components

DEFAULT_SCOPES_READONLY = ["User.Read", "Sites.Read.All"]
DEFAULT_SCOPES_WRITE = ["User.Read", "Sites.ReadWrite.All"]

_FLOW_QP_KEY = "ms_flow_b64"


def _cfg():
    # NEVER use st.secrets["..."] to avoid KeyError on import or missing keys.
    s = st.secrets.get("ms_graph", {})
    tenant_id = s.get("tenant_id") or os.getenv("MS_TENANT_ID", "")
    client_id = s.get("client_id") or os.getenv("MS_CLIENT_ID", "")
    client_secret = s.get("client_secret") or os.getenv("MS_CLIENT_SECRET", "")
    redirect_uri = s.get("redirect_uri") or os.getenv("MS_REDIRECT_URI", "")

    authority = s.get("authority")
    if not authority and tenant_id:
        authority = f"https://login.microsoftonline.com/{tenant_id}"

    return {
        "client_id": client_id,
        "client_secret": client_secret,
        "tenant_id": tenant_id,
        "redirect_uri": redirect_uri,
        "authority": authority or "",
    }


def _require_cfg():
    cfg = _cfg()
    missing = [k for k in ["client_id", "client_secret", "tenant_id", "redirect_uri", "authority"] if not cfg.get(k)]
    if missing:
        st.error("Microsoft Graph config is missing in Streamlit secrets: " + ", ".join(missing))
        st.stop()
    return cfg


def _msal_app():
    cfg = _require_cfg()
    return msal.ConfidentialClientApplication(
        client_id=cfg["client_id"],
        client_credential=cfg["client_secret"],
        authority=cfg["authority"],
    )


def _b64e(obj: dict) -> str:
    raw = json.dumps(obj).encode("utf-8")
    return base64.urlsafe_b64encode(raw).decode("utf-8")


def _b64d(s: str) -> dict:
    # Padding-safe decode
    s = (s or "").strip()
    pad = "=" * (-len(s) % 4)
    raw = base64.urlsafe_b64decode((s + pad).encode("utf-8"))
    return json.loads(raw.decode("utf-8"))


def _reset_login_state(clear_url=True):
    for k in ["ms_flow", "ms_token", "ms_scopes"]:
        st.session_state.pop(k, None)

    if clear_url:
        try:
            st.query_params.clear()
        except Exception:
            pass


def logout():
    _reset_login_state(clear_url=True)
    st.rerun()


def _ensure_flow(app, scopes):
    """
    Create or restore MSAL auth flow.
    Persist flow to query params so it survives Streamlit Cloud session loss.
    """
    cfg = _require_cfg()

    # Prefer session flow
    if "ms_flow" in st.session_state:
        return st.session_state["ms_flow"]

    # Try restoring from query params
    qp = st.query_params
    flow_b64 = qp.get(_FLOW_QP_KEY, None)
    if flow_b64:
        try:
            flow = _b64d(flow_b64)
            st.session_state["ms_flow"] = flow
            return flow
        except Exception:
            # If corrupted, just create a new flow
            pass

    # Create new flow
    flow = app.initiate_auth_code_flow(scopes=scopes, redirect_uri=cfg["redirect_uri"])
    st.session_state["ms_flow"] = flow

    # Persist into URL
    try:
        st.query_params[_FLOW_QP_KEY] = _b64e(flow)
    except Exception:
        pass

    return flow


ddef login_ui(scopes=None):
    """
    ONE button UI (no JS):
      - If callback comes back with ?code=, redeem token
      - Otherwise show a single Sign In link button
    """
    app = _msal_app()

    if scopes is None:
        scopes = DEFAULT_SCOPES_READONLY
    st.session_state["ms_scopes"] = scopes

    # Already logged in
    if st.session_state.get("ms_token"):
        return

    qp = st.query_params

    # CALLBACK
    if qp.get("code"):
        flow = _ensure_flow(app, scopes)
        auth_response = {k: qp.get(k) for k in qp.keys()}

        try:
            result = app.acquire_token_by_auth_code_flow(flow, auth_response)
        except ValueError:
            _reset_login_state(clear_url=True)
            st.error("Login failed. Click Sign In again.")
            return

        if "access_token" in result:
            st.session_state["ms_token"] = result
            try:
                st.query_params.clear()
            except Exception:
                pass
            st.rerun()
        else:
            _reset_login_state(clear_url=True)
            st.error("Login failed. Click Sign In again.")
            return

    # START LOGIN (single button, no JS)
    flow = _ensure_flow(app, scopes)
    auth_url = flow["auth_uri"]
    st.link_button("Sign In", auth_url)



def get_access_token():
    token = st.session_state.get("ms_token")
    if not token:
        return None

    expires_at = token.get("expires_at")
    if not expires_at:
        expires_in = token.get("expires_in", 3599)
        token["expires_at"] = int(time.time()) + int(expires_in)
        expires_at = token["expires_at"]

    if int(expires_at) - int(time.time()) < 120:
        _reset_login_state(clear_url=True)
        st.rerun()

    return token.get("access_token")
