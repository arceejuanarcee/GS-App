import os
import time
import streamlit as st
import msal

DEFAULT_SCOPES_READONLY = ["User.Read", "Sites.Read.All"]
DEFAULT_SCOPES_WRITE = ["User.Read", "Sites.ReadWrite.All"]

_STATE_QP_KEY = "ms_state"


def _cfg() -> dict:
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


@st.cache_resource
def _flow_store():
    # dict[state] = flow
    return {}


def _reset_login_state(clear_url: bool = True) -> None:
    for k in ["ms_token", "ms_scopes"]:
        st.session_state.pop(k, None)

    if clear_url:
        try:
            st.query_params.clear()
        except Exception:
            pass


def logout() -> None:
    _reset_login_state(clear_url=True)
    st.rerun()


def _start_flow(app: msal.ConfidentialClientApplication, scopes: list[str]) -> tuple[str, str]:
    """
    Create a new auth code flow, store it in cache by state,
    and return (auth_url, state).
    """
    cfg = _require_cfg()
    flow = app.initiate_auth_code_flow(scopes=scopes, redirect_uri=cfg["redirect_uri"])
    state = flow.get("state")
    if not state:
        raise RuntimeError("MSAL did not return a state value.")

    store = _flow_store()
    store[state] = flow

    # Put only the state in the URL (small, safe)
    try:
        st.query_params[_STATE_QP_KEY] = state
    except Exception:
        pass

    return flow["auth_uri"], state


def login_ui(scopes: list[str] | None = None) -> None:
    """
    ONE button UI:
      - If callback has ?code=, redeem using cached flow by state
      - Otherwise show single Sign In link_button
    """
    app = _msal_app()

    if scopes is None:
        scopes = DEFAULT_SCOPES_READONLY
    st.session_state["ms_scopes"] = scopes

    if st.session_state.get("ms_token"):
        return

    qp = st.query_params

    # Callback handling
    if qp.get("code"):
        state = qp.get("state") or qp.get(_STATE_QP_KEY)
        if not state:
            st.error("Login failed: missing state in callback.")
            return

        store = _flow_store()
        flow = store.get(state)
        if not flow:
            # flow not found (expired cache / new server instance)
            st.error("Login failed: session expired. Click Sign In again.")
            return

        auth_response = {k: qp.get(k) for k in qp.keys()}

        try:
            result = app.acquire_token_by_auth_code_flow(flow, auth_response)
        except ValueError as e:
            st.error(f"Login failed: {e}")
            return

        if "access_token" in result:
            st.session_state["ms_token"] = result
            try:
                st.query_params.clear()
            except Exception:
                pass
            st.rerun()
        else:
            err = result.get("error")
            desc = result.get("error_description")
            st.error(f"Login failed: {err} - {desc}")
            return

    # Start flow / show sign in
    auth_url, _ = _start_flow(app, scopes)
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

    if int(expires_at) - int(time.time()) < 120:
        _reset_login_state(clear_url=True)
        st.rerun()

    return token.get("access_token")
