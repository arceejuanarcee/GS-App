import os
import time
import json
import base64
import streamlit as st
import msal

DEFAULT_SCOPES_READONLY = ["User.Read", "Sites.Read.All"]
DEFAULT_SCOPES_WRITE = ["User.Read", "Sites.ReadWrite.All"]

_FLOW_QP_KEY = "ms_flow_b64"


def _cfg() -> dict:
    """
    Read config safely (no KeyError).
    Expected secrets:
      [ms_graph]
      client_id = "..."
      client_secret = "..."
      tenant_id = "..."
      redirect_uri = "https://<app>.streamlit.app/"
    """
    s = st.secrets.get("ms_graph", {})
    tenant_id = s.get("tenant_id") or os.getenv("MS_TENANT_ID", "")
    client_id = s.get("client_id") or os.getenv("MS_CLIENT_ID", "")
    client_secret = s.get("client_secret") or os.getenv("MS_CLIENT_SECRET", "")
    redirect_uri = s.get("redirect_uri") or os.getenv("MS_REDIRECT_URI", "")

    authority = s.get("authority")
    if not authority and tenant_id:
        authority = f"https://login.microsoftonline.com/{tenant_id}"

    return {
        "tenant_id": tenant_id,
        "client_id": client_id,
        "client_secret": client_secret,
        "redirect_uri": redirect_uri,
        "authority": authority or "",
    }


def _require_cfg() -> dict:
    cfg = _cfg()
    missing = [k for k in ["tenant_id", "client_id", "client_secret", "redirect_uri", "authority"] if not cfg.get(k)]
    if missing:
        st.error("Missing MS Graph config in secrets: " + ", ".join(missing))
        st.stop()
    return cfg


def _msal_app() -> msal.ConfidentialClientApplication:
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
    s = (s or "").strip()
    pad = "=" * (-len(s) % 4)
    raw = base64.urlsafe_b64decode((s + pad).encode("utf-8"))
    return json.loads(raw.decode("utf-8"))


def _reset_login_state(clear_url: bool = True) -> None:
    for k in ["ms_flow", "ms_token", "ms_scopes"]:
        st.session_state.pop(k, None)

    if clear_url:
        try:
            st.query_params.clear()
        except Exception:
            pass


def logout() -> None:
    _reset_login_state(clear_url=True)
    st.rerun()


def _ensure_flow(app: msal.ConfidentialClientApplication, scopes: list[str]) -> dict:
    """
    Ensure we have an MSAL flow. Store it in URL query param so it survives
    Streamlit Cloud new-session behavior.
    """
    cfg = _require_cfg()

    # Prefer existing flow in session
    if "ms_flow" in st.session_state:
        return st.session_state["ms_flow"]

    # Restore from URL if present
    qp = st.query_params
    flow_b64 = qp.get(_FLOW_QP_KEY, None)
    if flow_b64:
        try:
            flow = _b64d(flow_b64)
            st.session_state["ms_flow"] = flow
            return flow
        except Exception:
            # If corrupted, ignore and create new
            pass

    # Create new flow
    flow = app.initiate_auth_code_flow(scopes=scopes, redirect_uri=cfg["redirect_uri"])
    st.session_state["ms_flow"] = flow

    # Persist to URL
    try:
        st.query_params[_FLOW_QP_KEY] = _b64e(flow)
    except Exception:
        pass

    return flow


def login_ui(scopes: list[str] | None = None) -> None:
    """
    Minimal UI:
      - If not logged in, show ONE "Sign In" link button
      - If callback has ?code=, redeem and store token, then rerun
    """
    app = _msal_app()

    if scopes is None:
        scopes = DEFAULT_SCOPES_READONLY
    st.session_state["ms_scopes"] = scopes

    # Already logged in
    if st.session_state.get("ms_token"):
        return

    qp = st.query_params

    # Callback handler
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

    # Start login (ONE button)
    flow = _ensure_flow(app, scopes)
    auth_url = flow["auth_uri"]
    st.link_button("Sign In", auth_url)


def get_access_token() -> str | None:
    token = st.session_state.get("ms_token")
    if not token:
        return None

    expires_at = token.get("expires_at")
    if not expires_at:
        expires_in = token.get("expires_in", 3599)
        token["expires_at"] = int(time.time()) + int(expires_in)
        expires_at = token["expires_at"]

    # Refresh by forcing re-login if about to expire
    if int(expires_at) - int(time.time()) < 120:
        _reset_login_state(clear_url=True)
        st.rerun()

    return token.get("access_token")
