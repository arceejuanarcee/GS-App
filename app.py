# app.py
from pathlib import Path
import streamlit as st
import ms_graph

# -----------------------------
# Config
# -----------------------------
st.set_page_config(
    page_title="SMCOD Tools Portal",
    layout="wide",
)

ROOT = Path(__file__).parent
LOGO_PATH = ROOT / "graphics" / "PhilSA_v4-01.png"

# -----------------------------
# CSS – clean, no teal
# -----------------------------
st.markdown(
    """
<style>
header[data-testid="stHeader"] {
    visibility: hidden;
    height: 0px;
}
.block-container {
    padding-top: 1.5rem;
}
.logo-center {
    display: flex;
    justify-content: center;
    margin-bottom: 20px;
}
.portal-title {
    text-align: center;
    font-size: 26px;
    font-weight: 700;
}
.portal-sub {
    text-align: center;
    font-size: 14px;
    color: #666;
    margin-bottom: 30px;
}
div[data-testid="stButton"] > button {
    width: 100%;
    height: 90px;
    border-radius: 0;
    border: 2px solid #333;
    background: #0f5e73;
    color: white;
    font-size: 18px;
    font-weight: 600;
}
div[data-testid="stButton"] > button:hover {
    background: #0c4c5d;
}
</style>
""",
    unsafe_allow_html=True,
)

# -----------------------------
# Logo
# -----------------------------
if LOGO_PATH.exists():
    st.markdown("<div class='logo-center'>", unsafe_allow_html=True)
    st.image(str(LOGO_PATH), width=160)
    st.markdown("</div>", unsafe_allow_html=True)

st.markdown("<div class='portal-title'>PhilSA Tools Portal</div>", unsafe_allow_html=True)
st.markdown(
    "<div class='portal-sub'>Sign in with your PhilSA email to access internal tools.</div>",
    unsafe_allow_html=True,
)

# -----------------------------
# LOGIN (Microsoft Graph)
# -----------------------------
ms_graph.login_ui(scopes=ms_graph.DEFAULT_SCOPES_WRITE)
token = ms_graph.get_access_token()

# Stop here if not logged in
if not token:
    st.stop()

# -----------------------------
# Logout
# -----------------------------
col1, col2 = st.columns([6, 1])
with col2:
    if st.button("Logout"):
        ms_graph.logout()

st.write("")

# -----------------------------
# Dashboard
# -----------------------------
r1c1, r1c2, r1c3 = st.columns(3, gap="large")
r2c1, r2c2, r2c3 = st.columns(3, gap="large")

with r1c1:
    if st.button("Incident Report Generator"):
        st.switch_page("pages/1_Incident_Report_Generator.py")

with r1c2:
    if st.button("DGS Fault & Track Mapper"):
        st.switch_page("pages/2_DGS_Fault_Track_Mapper.py")

with r1c3:
    if st.button("SC Converter / Submit to TU"):
        st.switch_page("pages/3_SC_Converter_Submit.py")

with r2c1:
    st.button("Tool 4 (Coming soon)")

with r2c2:
    st.button("Tool 5 (Coming soon)")

with r2c3:
    st.button("Tool 6 (Coming soon)")
