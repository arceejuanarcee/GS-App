
import streamlit as st

st.set_page_config(page_title="SC Converter / Submit to TU", layout="wide")
st.title("SC Converter / Submit to TU")
st.info("Placeholder. Upload/point me to your SC Converter / Submit Commands script, then I’ll wire it here.")
st.write("Back to portal:")
if st.button("⬅ Home"):
    st.switch_page("app.py")
