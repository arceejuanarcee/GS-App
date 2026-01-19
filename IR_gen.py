import io
from datetime import date, datetime

import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Inches
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

# ---------------- CONFIG ----------------
TEMPLATE_PATH = "Incident Report Template_blank (1).docx"

CITY_CODES = {
    "Davao City": "DVO",
    "Quezon City": "QZN",
}

# ---------------- DOCX HELPERS ----------------
def _clear_table_rows_except_header(table, header_rows=1):
    while len(table.rows) > header_rows:
        tbl = table._tbl
        tr = table.rows[-1]._tr
        tbl.remove(tr)

def _set_2col_table_value(table, label, value):
    for row in table.rows:
        if row.cells[0].text.strip() == label.strip():
            row.cells[1].text = "" if value is None else str(value)
            return

def _set_paragraph_after_heading(doc, heading_text, new_text):
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip() == heading_text.strip():
            if i + 1 < len(doc.paragraphs):
                doc.paragraphs[i + 1].text = new_text or ""
            return

def _insert_paragraph_after(paragraph):
    """
    Insert a new empty paragraph after the given paragraph and return it.
    python-docx does NOT provide insert_paragraph_after() publicly.
    """
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    return Paragraph(new_p, paragraph._parent)

def _append_images_after_heading(doc, heading_text, files, width=5.5):
    """
    Appends uploaded images under a given heading by inserting new paragraphs
    after the body paragraph that follows the heading (or after heading if none).
    """
    if not files:
        return

    for i, p in enumerate(doc.paragraphs):
        if p.text.strip() == heading_text.strip():
            # place images after the paragraph right after the heading if it exists
            anchor = doc.paragraphs[i + 1] if i + 1 < len(doc.paragraphs) else p

            for f in files:
                img_bytes = f.getvalue()
                new_p = _insert_paragraph_after(anchor)
                run = new_p.add_run()
                run.add_picture(io.BytesIO(img_bytes), width=Inches(width))
                anchor = new_p

            return

def _fill_sequence_table(table, df):
    # Your template's Sequence of Events table has no header row (based on your earlier script),
    # so we clear all rows first.
    _clear_table_rows_except_header(table, header_rows=0)

    for _, r in df.iterrows():
        cells = table.add_row().cells
        cells[0].text = str(r.get("Date", ""))
        cells[1].text = str(r.get("Time", ""))
        cells[2].text = str(r.get("Category", ""))
        cells[3].text = str(r.get("Message", ""))

def _fill_actions_table(table, df):
    # Actions table has a header row in the template
    _clear_table_rows_except_header(table, header_rows=1)

    for _, r in df.iterrows():
        cells = table.add_row().cells
        cells[0].text = str(r.get("Date", ""))
        cells[1].text = str(r.get("Time", ""))
        cells[2].text = str(r.get("Performed by", ""))
        cells[3].text = str(r.get("Action", ""))
        cells[4].text = str(r.get("Result", ""))

def generate_docx(data):
    doc = Document(TEMPLATE_PATH)

    # Assumes template table order:
    # 0: Report header table
    # 1: Incident details table
    # 2: Sequence of events table
    # 3: Response/actions table
    t0, t1, t2, t3 = doc.tables[0], doc.tables[1], doc.tables[2], doc.tables[3]

    # Header tables
    _set_2col_table_value(t0, "Reported by", data["reported_by"])
    _set_2col_table_value(t0, "Position", data["position"])
    _set_2col_table_value(t0, "Date of Report", data["date_of_report"])
    _set_2col_table_value(t0, "Incident No.", data["incident_no"])

    _set_2col_table_value(t1, "Date (YYYY-MM-DD)", data["incident_date"])
    _set_2col_table_value(t1, "Time", data["incident_time"])
    _set_2col_table_value(t1, "Location", data["location"])
    _set_2col_table_value(t1, "Current Status", data["current_status"])

    # Text sections
    _set_paragraph_after_heading(doc, "Nature of Incident", data["nature"])
    _set_paragraph_after_heading(doc, "Damages Incurred (if any)", data["damages"])
    _set_paragraph_after_heading(doc, "Investigation and Analysis", data["investigation"])
    _set_paragraph_after_heading(doc, "Conclusion and Recommendations", data["conclusion"])

    # Tables
    _fill_sequence_table(t2, data["sequence_df"])
    _fill_actions_table(t3, data["actions_df"])

    # Images (inserted under headings)
    _append_images_after_heading(doc, "Sequence of Events", data["sequence_images"])
    _append_images_after_heading(doc, "Damages Incurred (if any)", data["damages_images"])
    _append_images_after_heading(doc, "Investigation and Analysis", data["investigation_images"])
    _append_images_after_heading(doc, "Conclusion and Recommendations", data["conclusion_images"])

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out.read()

# ---------------- STREAMLIT UI ----------------
st.set_page_config(page_title="IR Generator", layout="wide")
st.title("Incident Report Generator")

year = st.selectbox("Year", [datetime.now().year, datetime.now().year - 1])
city = st.selectbox("City", list(CITY_CODES.keys()))
site_code = CITY_CODES[city]

# Local auto-increment (session only)
st.session_state.setdefault("ir_counter", 1)
incident_no = f"SMCOD-IR-GS-{site_code}-{year}-{st.session_state['ir_counter']:04d}"

st.text_input("Incident No. (auto)", value=incident_no, disabled=True)

with st.form("ir_form"):
    c1, c2 = st.columns(2)

    with c1:
        reported_by = st.text_input("Reported by")
        position = st.text_input("Position")
        date_of_report = st.date_input("Date of Report", value=date.today()).strftime("%Y-%m-%d")

    with c2:
        incident_date = st.date_input("Incident Date", value=date.today()).strftime("%Y-%m-%d")
        incident_time = st.text_input("Incident Time", value=datetime.now().strftime("%H:%M:%S"))
        location = st.text_input("Location", value=city)
        current_status = st.selectbox("Current Status", ["Resolved", "Ongoing", "Monitoring", "Open"])

    nature = st.text_area("Nature of Incident", height=120)

    st.subheader("Sequence of Events")
    seq_df = st.data_editor(
        pd.DataFrame([{"Date": incident_date, "Time": "", "Category": "", "Message": ""}]),
        num_rows="dynamic",
        use_container_width=True,
    )
    seq_imgs = st.file_uploader("Sequence Photos", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

    damages = st.text_area("Damages Incurred", value="None")
    dmg_imgs = st.file_uploader("Damage Photos", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

    investigation = st.text_area("Investigation and Analysis", height=120)
    inv_imgs = st.file_uploader("Investigation Photos", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

    conclusion = st.text_area("Conclusion and Recommendations", height=120)
    con_imgs = st.file_uploader("Conclusion Photos", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

    st.subheader("Response and Actions Taken")
    actions_df = st.data_editor(
        pd.DataFrame([{"Date": incident_date, "Time": "", "Performed by": "", "Action": "", "Result": ""}]),
        num_rows="dynamic",
        use_container_width=True,
    )

    submit = st.form_submit_button("Generate Report")

if submit:
    data = {
        "reported_by": reported_by,
        "position": position,
        "date_of_report": date_of_report,
        "incident_no": incident_no,
        "incident_date": incident_date,
        "incident_time": incident_time,
        "location": location,
        "current_status": current_status,
        "nature": nature,
        "damages": damages,
        "investigation": investigation,
        "conclusion": conclusion,
        "sequence_df": seq_df,
        "actions_df": actions_df,
        "sequence_images": seq_imgs or [],
        "damages_images": dmg_imgs or [],
        "investigation_images": inv_imgs or [],
        "conclusion_images": con_imgs or [],
    }

    docx_bytes = generate_docx(data)
    st.session_state["ir_counter"] += 1

    st.success("Incident Report generated.")
    st.download_button(
        "Download Incident Report (DOCX)",
        data=docx_bytes,
        file_name=f"{incident_no}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
