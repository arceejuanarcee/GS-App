
import streamlit as st

st.set_page_config(page_title="DGS Fault & Track Mapper", layout="wide")
st.title("DGS Fault & Track Mapper")
st.info("Placeholder. Upload/point me to your Fault & Track Mapper Streamlit script, then I’ll wire it here.")
st.write("Back to portal:")
if st.button("⬅ Home"):
    st.switch_page("app.py")
