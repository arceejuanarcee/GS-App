import os
import time
import streamlit as st
import msal
import streamlit.components.v1 as components

DEFAULT_SCOPES_READONLY = ["User.Read", "Sites.Read.All"]
DEFAULT_SCOPES_WRITE = ["User.Read", "Sites.ReadWrite.All"]

def _cfg():
    s = st.secrets.get("ms_graph", {})
    tenant_id = s.get("tenant_id") or os.getenv("MS_TENANT_ID", "")
    return {
        "client_id": s.get("client_id") or os.getenv("MS_CLIENT_ID", ""),
        "client_secret": s.get("client_secret") or os.getenv("MS_CLIENT_SECRET", ""),
        "tenant_id": tenant_id,
        "redirect_uri": s.get("redirect_uri") or os.getenv("MS_REDIRECT_URI", ""),
        "authority": s.get("authority") or f"https://login.microsoftonline.com/{tenant_id}",
        "debug": bool(s.get("debug", False)),
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

def _reset_login_state(clear_url=True):
    for k in ["ms_flow", "ms_token", "ms_scopes", "ms_auth_clicked"]:
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
    cfg = _cfg()
    if "ms_flow" not in st.session_state:
        st.session_state["ms_flow"] = app.initiate_auth_code_flow(
            scopes=scopes,
            redirect_uri=cfg["redirect_uri"],
        )
    return st.session_state["ms_flow"]

def login_ui(scopes=None):
    """
    Minimal UI:
      - Shows only "Sign In"
      - First tries same-tab redirect via JS
      - If JS is blocked, shows a single link_button also labeled "Sign In"
    """
    cfg = _cfg()
    app = _msal_app()

    if scopes is None:
        scopes = DEFAULT_SCOPES_READONLY
    st.session_state["ms_scopes"] = scopes

    # Already logged in
    if st.session_state.get("ms_token"):
        st.success("Logged in.")
        logout_button()
        return

    qp = st.query_params

    # Optional debug (goes to app log, not UI)
    if cfg["debug"]:
        # Safe debug: show only qp keys
        st.write({"debug_query_param_keys": list(qp.keys())})

    # CALLBACK: redeem code
    if "code" in qp:
        flow = st.session_state.get("ms_flow")
        if not flow:
            # Can't redeem without original flow; restart cleanly
            _reset_login_state(clear_url=True)
            st.rerun()

        auth_response = {k: qp.get(k) for k in qp.keys()}
        try:
            result = app.acquire_token_by_auth_code_flow(flow, auth_response)
        except ValueError:
            _reset_login_state(clear_url=True)
            st.rerun()

        if "access_token" in result:
            st.session_state["ms_token"] = result
            try:
                st.query_params.clear()
            except Exception:
                pass
            st.rerun()
        else:
            _reset_login_state(clear_url=True)
            st.rerun()

    # START LOGIN
    flow = _ensure_flow(app, scopes)
    auth_url = flow["auth_uri"]

    clicked = st.button("Sign In")
    if clicked:
        st.session_state["ms_auth_clicked"] = True
        # same-tab redirect attempt
        components.html(
            f"""
            <script>
              try {{
                window.top.location.href = "{auth_url}";
              }} catch (e) {{}}
            </script>
            """,
            height=0,
        )

    # Show link fallback ONLY after click (so it doesn't appear immediately)
    if st.session_state.get("ms_auth_clicked"):
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
        _reset_login_state()
        st.rerun()

    return token.get("access_token")
