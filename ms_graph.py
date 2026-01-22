import os
import time
import streamlit as st
import msal

# IMPORTANT: Do NOT include "openid", "profile", "offline_access" in SCOPES for MSAL python
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
    # also clear URL params so we don't re-trigger the same bad callback
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
    Safe Streamlit/MSAL auth-code flow:
    - Handles callback once
    - Avoids re-creating flow each rerun
    - If state mismatch happens, resets and asks user to try again
    """
    cfg = _cfg()
    app = _msal_app()

    if scopes is None:
        scopes = DEFAULT_SCOPES_READONLY
    st.session_state["ms_scopes"] = scopes

    # Already logged in
    if st.session_state.get("ms_token"):
        st.success("Logged in to Microsoft.")
        logout_button()
        return

    # --- CALLBACK HANDLER ---
    qp = st.query_params  # streamlit >=1.30
    if "code" in qp:
        flow = st.session_state.get("ms_flow")
        if not flow:
            # Flow missing usually means rerun/new session; restart cleanly
            st.warning("Login session expired. Please sign in again.")
            _reset_login_state()
            st.stop()

        try:
            # Convert query params to normal dict[str,str]
            auth_response = {k: qp.get(k) for k in qp.keys()}
            result = app.acquire_token_by_auth_code_flow(flow, auth_response)

        except ValueError as e:
            # This catches "state mismatch" and similar flow issues
            if "state mismatch" in str(e).lower():
                st.warning(
                    "Login was interrupted (state mismatch). "
                    "Please click Sign in again and avoid opening multiple tabs."
                )
                _reset_login_state()
                st.stop()
            raise

        if "access_token" in result:
            st.session_state["ms_token"] = result
            # clear query params after successful login
            st.query_params.clear()
            st.success("Login successful.")
            st.rerun()
        else:
            st.error(f"Login failed: {result.get('error')} - {result.get('error_description')}")
            _reset_login_state()
            st.stop()

    # --- START LOGIN (ONLY CREATE FLOW ONCE) ---
    if "ms_flow" not in st.session_state:
        st.session_state["ms_flow"] = app.initiate_auth_code_flow(
            scopes=scopes,
            redirect_uri=cfg["redirect_uri"],
        )

    auth_url = st.session_state["ms_flow"].get("auth_uri")

    st.markdown("### Microsoft Sign-in required")
    st.markdown(
        "To avoid errors: **click the sign-in link once**, and do not open it in multiple tabs."
    )

    # Use a button to prevent accidental multi-click reruns
    if st.button("Sign in with Microsoft"):
        st.markdown(f"[Continue to Microsoft sign-in]({auth_url})")

    # Show link as fallback
    st.caption("If the button does not open, use the link below:")
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
