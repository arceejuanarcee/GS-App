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
    s = st.secrets.get("ms_graph", {})
    tenant_id = s.get("tenant_id") or os.getenv("MS_TENANT_ID", "")
    return {
        "client_id": s.get("client_id") or os.getenv("MS_CLIENT_ID", ""),
        "client_secret": s.get("client_secret") or os.getenv("MS_CLIENT_SECRET", ""),
        "tenant_id": tenant_id,
        "redirect_uri": s.get("redirect_uri") or os.getenv("MS_REDIRECT_URI", ""),
        "authority": s.get("authority") or f"https://login.microsoftonline.com/{tenant_id}",
    }


def _msal_app():
    cfg = _cfg()
    missing = [k for k in ["client_id", "client_secret", "tenant_id", "redirect_uri"] if not cfg.get(k)]
    if missing:
        st.error(f"Missing MS Graph config in secrets: {', '.join(missing)}")
        st.stop()

    return msal.ConfidentialClientApplication(
        client_id=cfg["client_id"],
        client_credential=cfg["client_secret"],
        authority=cfg["authority"],
    )


def _b64e(obj: dict) -> str:
    raw = json.dumps(obj).encode("utf-8")
    return base64.urlsafe_b64encode(raw).decode("utf-8")


def _b64d(s: str) -> dict:
    raw = base64.urlsafe_b64decode(s.encode("utf-8"))
    return json.loads(raw.decode("utf-8"))


def _reset_login_state(clear_url=True):
    for k in ["ms_flow", "ms_token", "ms_scopes"]:
        st.session_state.pop(k, None)
    if clear_url:
        try:
            st.query_params.clear()
        except Exception:
            pass


def logout_button():
    if st.button("Log out"):
        _reset_login_state()
        st.rerun()


def _ensure_flow(app, scopes):
    """
    Create a flow and persist it BOTH in session_state and in query params.
    This survives Streamlit Cloud session loss after MS redirects back.
    """
    cfg = _cfg()

    # If flow already exists in session, use it
    if "ms_flow" in st.session_state:
        return st.session_state["ms_flow"]

    # If flow is in query params, restore it
    qp = st.query_params
    if _FLOW_QP_KEY in qp:
        try:
            flow = _b64d(qp.get(_FLOW_QP_KEY))
            st.session_state["ms_flow"] = flow
            return flow
        except Exception:
            # If it's corrupted, ignore and create new flow
            pass

    # Otherwise create new flow
    flow = app.initiate_auth_code_flow(scopes=scopes, redirect_uri=cfg["redirect_uri"])
    st.session_state["ms_flow"] = flow

    # Persist flow into URL so it survives a new session
    try:
        st.query_params[_FLOW_QP_KEY] = _b64e(flow)
    except Exception:
        pass

    return flow


def login_ui(scopes=None):
    """
    ONE button only: Sign In
    - If already logged in: returns
    - If callback (?code=...): redeems token and reruns
    - Otherwise: shows Sign In button and redirects same tab
    """
    app = _msal_app()

    if scopes is None:
        scopes = DEFAULT_SCOPES_READONLY
    st.session_state["ms_scopes"] = scopes

    # Already logged in
    if st.session_state.get("ms_token"):
        # (Optional) keep it minimal
        return

    qp = st.query_params

    # CALLBACK: if we have a code, redeem it
    if "code" in qp:
        flow = _ensure_flow(app, scopes)  # restore flow even if session was lost
        auth_response = {k: qp.get(k) for k in qp.keys()}

        try:
            result = app.acquire_token_by_auth_code_flow(flow, auth_response)
        except ValueError:
            # State mismatch, stale flow, etc.
            _reset_login_state(clear_url=True)
            st.error("Login failed. Please click Sign In again.")
            return

        if "access_token" in result:
            st.session_state["ms_token"] = result
            # Clear URL params after success
            try:
                st.query_params.clear()
            except Exception:
                pass
            st.rerun()
        else:
            _reset_login_state(clear_url=True)
            st.error("Login failed. Please click Sign In again.")
            return

    # START LOGIN (no callback yet)
    flow = _ensure_flow(app, scopes)
    auth_url = flow["auth_uri"]

    # ONE button only
    if st.button("Sign In"):
        # Same-tab redirect
        components.html(
            f"""
            <script>
              window.top.location.href = "{auth_url}";
            </script>
            """,
            height=0,
        )


def get_access_token():
    token = st.session_state.get("ms_token")
    if not token:
        return None

    expires_at = token.get("expires_at")
    if not expires_at:
        expires_in = token.get("expires_in", 3599)
        token["expires_at"] = int(time.time()) + int(expires_in)
        expires_at = token["expires_at"]

    # expire soon => force relogin
    if int(expires_at) - int(time.time()) < 120:
        _reset_login_state()
        st.rerun()

    return token.get("access_token")
