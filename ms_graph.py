import os
import time
import streamlit as st
import msal

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
    Ensure we have a flow stored in session_state.
    """
    cfg = _cfg()
    if "ms_flow" not in st.session_state:
        st.session_state["ms_flow"] = app.initiate_auth_code_flow(
            scopes=scopes,
            redirect_uri=cfg["redirect_uri"],
        )
    return st.session_state["ms_flow"]

def login_ui(scopes=None):
    """
    Streamlit Cloud friendly auth flow.
    Key behavior:
      - Always shows a single auth link (user should use same tab)
      - If callback arrives but flow missing, it restarts flow automatically
    """
    app = _msal_app()

    if scopes is None:
        scopes = DEFAULT_SCOPES_READONLY
    st.session_state["ms_scopes"] = scopes

    # Already logged in
    if st.session_state.get("ms_token"):
        st.success("Logged in to Microsoft.")
        logout_button()
        return

    qp = st.query_params

    # --- CALLBACK ---
    if "code" in qp:
        # If flow missing, recreate flow and ask user to sign in again.
        # We cannot safely redeem an auth code without the original flow state.
        flow = st.session_state.get("ms_flow")
        if not flow:
            st.warning(
                "Login callback received but session state was lost (Streamlit Cloud behavior). "
                "Please click the sign-in link again (same tab)."
            )
            # Clear the old callback params, restart flow
            _reset_login_state(clear_url=True)
            flow = _ensure_flow(app, scopes)
            st.markdown(f"[Sign in with Microsoft]({flow['auth_uri']})")
            st.stop()

        try:
            auth_response = {k: qp.get(k) for k in qp.keys()}
            result = app.acquire_token_by_auth_code_flow(flow, auth_response)
        except ValueError as e:
            # state mismatch or other flow errors
            st.warning(
                f"Login interrupted ({str(e)}). Please sign in again and avoid multiple tabs/refresh."
            )
            _reset_login_state(clear_url=True)
            flow = _ensure_flow(app, scopes)
            st.markdown(f"[Sign in with Microsoft]({flow['auth_uri']})")
            st.stop()

        if "access_token" in result:
            st.session_state["ms_token"] = result
            try:
                st.query_params.clear()
            except Exception:
                pass
            st.success("Login successful.")
            st.rerun()
        else:
            st.error(f"Login failed: {result.get('error')} - {result.get('error_description')}")
            _reset_login_state(clear_url=True)
            flow = _ensure_flow(app, scopes)
            st.markdown(f"[Sign in with Microsoft]({flow['auth_uri']})")
            st.stop()

    # --- START LOGIN ---
    flow = _ensure_flow(app, scopes)

    st.markdown("### Microsoft Sign-in required")
    st.caption("Important: Use the sign-in link in the SAME tab. Avoid refresh while signing in.")
    st.markdown(f"[Sign in with Microsoft]({flow['auth_uri']})")

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
        st.warning("Session expired. Please log in again.")
        _reset_login_state()
        st.rerun()

    return token.get("access_token")
