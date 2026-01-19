import io
import os
import re
import tempfile
import subprocess
from datetime import date, datetime

import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Inches

from ms_graph import login_ui, access_token
from sp_folder_graph import (
    resolve_root_folder_from_share_link,
    ensure_folder,
    list_children,
    compute_next_incident_id,
    ensure_incident_folder,
    upload_file_to_folder,
)

TEMPLATE_PATH = "Incident Report Template_blank (1).docx"

# City -> site code mapping (adjust if needed)
CITY_CODES = {
    "Davao City": "DVO",
    "Quezon City": "QZN",
}

# -------- PDF conversion (optional) --------
def convert_docx_to_pdf_bytes(docx_bytes: bytes) -> bytes:
    """
    Requires LibreOffice 'soffice' on the server.
    Streamlit Cloud usually does NOT have libreoffice by default.
    If you need PDF there, we can switch to:
      - docx-only export, or
      - external conversion service, or
      - host on your own Linux VM with libreoffice installed.
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = os.path.join(tmpdir, "report.docx")
        with open(docx_path, "wb") as f:
            f.write(docx_bytes)

        cmd = [
            "soffice", "--headless", "--nologo", "--nofirststartwizard",
            "--convert-to", "pdf", "--outdir", tmpdir, docx_path,
        ]
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        pdf_path = os.path.join(tmpdir, "report.pdf")
        with open(pdf_path, "rb") as f:
            return f.read()

# -------- docx utilities --------
def _clear_table_rows_except_header(table, header_rows=1):
    while len(table.rows) > header_rows:
        tbl = table._tbl
        tr = table.rows[-1]._tr
        tbl.remove(tr)

def _set_2col_table_value(table, label, value):
    for row in table.rows:
        if row.cells[0].text.strip() == label.strip():
            row.cells[1].text = "" if value is None else str(value)
            return True
    return False

def _set_paragraph_after_heading(doc, heading_text, new_text):
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip() == heading_text.strip():
            if i + 1 < len(doc.paragraphs):
                doc.paragraphs[i + 1].text = "" if new_text is None else str(new_text)
                return True
    return False

def _append_images_after_heading(doc, heading_text, uploaded_files, width_in=5.5):
    if not uploaded_files:
        return False
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip() == heading_text.strip():
            # Insert after the body paragraph that follows heading
            insert_after = i + 1
            if insert_after >= len(doc.paragraphs):
                insert_after = len(doc.paragraphs) - 1
            # Insert images as new paragraphs after that paragraph
            anchor = doc.paragraphs[insert_after]
            for f in uploaded_files:
                img_bytes = f.getvalue() if hasattr(f, "getvalue") else f.read()
                new_p = anchor.insert_paragraph_after()
                run = new_p.add_run()
                run.add_picture(io.BytesIO(img_bytes), width=Inches(width_in))
                anchor = new_p
            return True
    return False

def _fill_sequence_table_from_df(table, seq_df: pd.DataFrame):
    _clear_table_rows_except_header(table, header_rows=0)  # template's table has no header row
    for _, r in seq_df.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(r.get("Date", ""))
        row_cells[1].text = str(r.get("Time", ""))
        row_cells[2].text = str(r.get("Category", ""))
        row_cells[3].text = str(r.get("Message", ""))

def _fill_actions_table(table, actions_df: pd.DataFrame):
    _clear_table_rows_except_header(table, header_rows=1)
    for _, r in actions_df.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(r.get("Date", ""))
        row_cells[1].text = str(r.get("Time", ""))
        row_cells[2].text = str(r.get("Performed by", ""))
        row_cells[3].text = str(r.get("Action", ""))
        row_cells[4].text = str(r.get("Result", ""))

def generate_docx_bytes(template_path, data):
    doc = Document(template_path)

    # Template tables (based on your uploaded template)
    t0, t1, t2, t3 = doc.tables[0], doc.tables[1], doc.tables[2], doc.tables[3]

    # Table 0
    _set_2col_table_value(t0, "Reported by", data["reported_by"])
    _set_2col_table_value(t0, "Position", data["position"])
    _set_2col_table_value(t0, "Date of Report", data["date_of_report"])
    _set_2col_table_value(t0, "Incident No.", data["incident_no"])

    # Table 1
    _set_2col_table_value(t1, "Date (YYYY-MM-DD)", data["incident_date"])
    _set_2col_table_value(t1, "Time", data["incident_time"])
    _set_2col_table_value(t1, "Location", data["location"])
    _set_2col_table_value(t1, "Current Status", data["current_status"])

    # Paragraph sections
    _set_paragraph_after_heading(doc, "Nature of Incident", data["nature_of_incident"])
    _set_paragraph_after_heading(doc, "Damages Incurred (if any)", data["damages_incurred"])
    _set_paragraph_after_heading(doc, "Investigation and Analysis", data["investigation"])
    _set_paragraph_after_heading(doc, "Conclusion and Recommendations", data["conclusion"])

    # Tables
    _fill_sequence_table_from_df(t2, data["seq_df"])
    _fill_actions_table(t3, data["actions_df"])

    # Append images under headings
    _append_images_after_heading(doc, "Sequence of Events", data["seq_images"])
    _append_images_after_heading(doc, "Damages Incurred (if any)", data["damages_images"])
    _append_images_after_heading(doc, "Investigation and Analysis", data["investigation_images"])
    _append_images_after_heading(doc, "Conclusion and Recommendations", data["conclusion_images"])

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out.read()

# ---------------- Streamlit UI ----------------
st.set_page_config(page_title="IR Generator", layout="wide")
st.title("Incident Report Generator")

login_ui()
token = access_token()

if not token:
    st.warning("Please login to Microsoft first (link above).")
    st.stop()

cfg = st.secrets["ms_graph"]
share_link = cfg["share_folder_link"]

# Resolve root shared folder
drive_id, root_id = resolve_root_folder_from_share_link(share_link, token)

# Select year + city
year = st.selectbox("Year", [datetime.now().year, datetime.now().year - 1, datetime.now().year + 1], index=0)
city = st.selectbox("City folder", list(CITY_CODES.keys()), index=0)
site_code = CITY_CODES[city]

# Ensure folders: Year -> City
year_folder = ensure_folder(drive_id, root_id, str(year), token, create_if_missing=True)
city_folder = ensure_folder(drive_id, year_folder["id"], city, token, create_if_missing=True)

# Compute next incident ID based on EXISTING INCIDENT FOLDERS in that city folder
city_children = list_children(drive_id, city_folder["id"], token)
incident_no = compute_next_incident_id(city_children, year, site_code)

st.text_input("Incident No. (auto)", value=incident_no, disabled=True)

st.divider()

with st.form("ir_form"):
    c1, c2 = st.columns(2)

    with c1:
        reported_by = st.text_input("Reported by")
        position = st.text_input("Position")
        date_of_report = st.date_input("Date of Report", value=date.today()).strftime("%Y-%m-%d")

    with c2:
        incident_date = st.date_input("Incident Date", value=date.today()).strftime("%Y-%m-%d")
        incident_time = st.text_input("Incident Time (HH:MM:SS)", value=datetime.now().strftime("%H:%M:%S"))
        location = st.text_input("Location", value=city)
        current_status = st.selectbox("Current Status", ["Resolved", "Ongoing", "Monitoring", "Open"], index=0)

    st.subheader("Nature of Incident")
    nature_of_incident = st.text_area("Nature of Incident", height=120)

    st.subheader("Sequence of Events (table)")
    seq_df_default = pd.DataFrame([{"Date": incident_date, "Time": "", "Category": "", "Message": ""}])
    seq_df = st.data_editor(seq_df_default, num_rows="dynamic", use_container_width=True)

    seq_images = st.file_uploader(
        "Sequence of Events Photos (optional, appended under Sequence of Events)",
        type=["png", "jpg", "jpeg"],
        accept_multiple_files=True,
    )

    st.subheader("Damages / Investigation / Conclusion")
    damages_incurred = st.text_area("Damages Incurred (if any)", height=80, value="None")
    damages_images = st.file_uploader("Damages Photos (optional)", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

    investigation = st.text_area("Investigation and Analysis", height=120)
    investigation_images = st.file_uploader("Investigation Photos (optional)", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

    conclusion = st.text_area("Conclusion and Recommendations", height=120)
    conclusion_images = st.file_uploader("Conclusion/Recommendations Photos (optional)", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

    st.subheader("Response and Actions Taken (table)")
    actions_default = pd.DataFrame([{"Date": incident_date, "Time": "", "Performed by": "", "Action": "", "Result": ""}])
    actions_df = st.data_editor(actions_default, num_rows="dynamic", use_container_width=True)

    export_pdf = st.checkbox("Also export PDF (requires LibreOffice on host)", value=False)

    submitted = st.form_submit_button("Generate and Upload")

if submitted:
    if not os.path.exists(TEMPLATE_PATH):
        st.error(f"Template not found in repo root: {TEMPLATE_PATH}")
        st.stop()

    # Create incident folder first (prevents duplicates by failing if it exists)
    try:
        incident_folder = ensure_incident_folder(drive_id, city_folder["id"], incident_no, token)
    except Exception as e:
        st.error(f"Could not create incident folder (maybe already exists): {e}")
        st.stop()

    incident_folder_id = incident_folder["id"]

    data = {
        "reported_by": reported_by,
        "position": position,
        "date_of_report": date_of_report,
        "incident_no": incident_no,
        "incident_date": incident_date,
        "incident_time": incident_time,
        "location": location,
        "current_status": current_status,
        "nature_of_incident": nature_of_incident,
        "seq_df": seq_df,
        "actions_df": actions_df,
        "damages_incurred": damages_incurred,
        "investigation": investigation,
        "conclusion": conclusion,
        "seq_images": seq_images or [],
        "damages_images": damages_images or [],
        "investigation_images": investigation_images or [],
        "conclusion_images": conclusion_images or [],
    }

    docx_bytes = generate_docx_bytes(TEMPLATE_PATH, data)

    # Upload DOCX
    upload_file_to_folder(drive_id, incident_folder_id, f"{incident_no}.docx", docx_bytes, token)

    # Upload images as separate files too (optional but useful for records)
    def upload_images(prefix, files):
        for idx, f in enumerate(files, start=1):
            ext = "jpg"
            name_lower = f.name.lower()
            if name_lower.endswith(".png"):
                ext = "png"
            elif name_lower.endswith(".jpeg"):
                ext = "jpeg"
            elif name_lower.endswith(".jpg"):
                ext = "jpg"
            else:
                # default if unknown
                ext = name_lower.split(".")[-1] if "." in name_lower else "jpg"

            fname = f"{prefix}_{idx:02d}.{ext}"
            upload_file_to_folder(drive_id, incident_folder_id, fname, f.getvalue(), token)

    upload_images("sequence", seq_images or [])
    upload_images("damages", damages_images or [])
    upload_images("investigation", investigation_images or [])
    upload_images("conclusion", conclusion_images or [])

    # Optional PDF (may fail on Streamlit Cloud)
    pdf_bytes = None
    if export_pdf:
        try:
            pdf_bytes = convert_docx_to_pdf_bytes(docx_bytes)
            upload_file_to_folder(drive_id, incident_folder_id, f"{incident_no}.pdf", pdf_bytes, token)
        except Exception as e:
            st.warning(f"PDF conversion failed on this host: {e}")

    st.success(f"Done. Uploaded to {year}/{city}/{incident_no}/")

    # Provide downloads too
    st.download_button(
        "Download DOCX",
        data=docx_bytes,
        file_name=f"{incident_no}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

    if pdf_bytes:
        st.download_button(
            "Download PDF",
            data=pdf_bytes,
            file_name=f"{incident_no}.pdf",
            mime="application/pdf",
        )
