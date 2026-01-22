import os
import time
import streamlit as st
import msal

# IMPORTANT: Do NOT include "openid", "profile", "offline_access" in scopes for MSAL Python.
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

def _reset_login_state():
    for k in ["ms_flow", "ms_token", "ms_scopes"]:
        st.session_state.pop(k, None)
    try:
        st.query_params.clear()
    except Exception:
        pass

def logout_button():
    if st.button("Log out"):
        _reset_login_state()
        st.rerun()

def login_ui(scopes=None):
    """
    Streamlit/MSAL auth code flow:
    - Avoids regenerating auth flow each rerun
    - Handles callback
    - Recovers from state mismatch cleanly
    """
    cfg = _cfg()
    app = _msal_app()

    if scopes is None:
        scopes = DEFAULT_SCOPES_READONLY

    st.session_state["ms_scopes"] = scopes

    if st.session_state.get("ms_token"):
        st.success("Logged in to Microsoft.")
        logout_button()
        return

    qp = st.query_params

    # Callback
    if "code" in qp:
        flow = st.session_state.get("ms_flow")
        if not flow:
            st.warning("Login session expired. Please sign in again.")
            _reset_login_state()
            st.stop()

        try:
            auth_response = {k: qp.get(k) for k in qp.keys()}
            result = app.acquire_token_by_auth_code_flow(flow, auth_response)
        except ValueError as e:
            if "state mismatch" in str(e).lower():
                st.warning(
                    "Login was interrupted (state mismatch). "
                    "Please sign in again and avoid multiple tabs/refresh during login."
                )
                _reset_login_state()
                st.stop()
            raise

        if "access_token" in result:
            st.session_state["ms_token"] = result
            st.query_params.clear()
            st.success("Login successful.")
            st.rerun()
        else:
            st.error(f"Login failed: {result.get('error')} - {result.get('error_description')}")
            _reset_login_state()
            st.stop()

    # Start login (create flow only once)
    if "ms_flow" not in st.session_state:
        st.session_state["ms_flow"] = app.initiate_auth_code_flow(
            scopes=scopes,
            redirect_uri=cfg["redirect_uri"],
        )

    auth_url = st.session_state["ms_flow"].get("auth_uri")

    st.markdown("### Microsoft Sign-in required")
    st.markdown("Tip: click sign-in **once** and avoid refreshing during login.")

    if st.button("Sign in with Microsoft"):
        st.markdown(f"[Continue to Microsoft sign-in]({auth_url})")

    st.caption("Fallback link:")
    st.markdown(f"[Microsoft Sign-in Link]({auth_url})")

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
