import streamlit as st
import msal

# Do NOT include: "openid", "profile", "offline_access"
SCOPES = [
    "User.Read",
    "Files.ReadWrite.All",
    "Sites.ReadWrite.All",
]

def _cfg():
    return st.secrets["ms_graph"]

def build_msal_app():
    authority = f"https://login.microsoftonline.com/{_cfg()['tenant_id']}"
    return msal.ConfidentialClientApplication(
        client_id=_cfg()["client_id"],
        client_credential=_cfg()["client_secret"],
        authority=authority,
    )

def login_ui():
    """
    Renders a login link, and exchanges ?code=... for tokens.
    Stores tokens in st.session_state['ms_token'].
    """
    st.session_state.setdefault("ms_flow", None)

    app = build_msal_app()

    if st.session_state["ms_flow"] is None:
        st.session_state["ms_flow"] = app.initiate_auth_code_flow(
            scopes=SCOPES,
            redirect_uri=_cfg()["redirect_uri"],
        )

    flow = st.session_state["ms_flow"]
    st.markdown(f"[Login to Microsoft]({flow['auth_uri']})")

    params = dict(st.query_params)

    # When redirected back, the URL will contain ?code=...&state=...
    if "code" in params and "ms_token" not in st.session_state:
        result = app.acquire_token_by_auth_code_flow(flow, params)
        if "access_token" not in result:
            raise RuntimeError(f"Token error: {result.get('error')} - {result.get('error_description')}")
        st.session_state["ms_token"] = result
        st.success("Logged in to Microsoft.")

def access_token():
    tok = st.session_state.get("ms_token", {})
    return tok.get("access_token")
