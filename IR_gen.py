import io
import os
import re
import tempfile
import subprocess
from datetime import date, datetime

import pandas as pd
import streamlit as st
from docx import Document

TEMPLATE_PATH = "Incident Report Template_blank (1).docx"  # put this beside app.py


# -------------------------
# DOCX helpers
# -------------------------
def _clear_table_rows_except_header(table, header_rows=1):
    """Remove all rows except the first header_rows rows."""
    # python-docx doesn't have a direct delete-row, but this works:
    while len(table.rows) > header_rows:
        tbl = table._tbl
        tr = table.rows[-1]._tr
        tbl.remove(tr)

def _set_2col_table_value(table, label, value):
    """
    For tables like:
    [Label | Value]
    Find row where first cell == label, set second cell to value.
    """
    for row in table.rows:
        if row.cells[0].text.strip() == label.strip():
            row.cells[1].text = "" if value is None else str(value)
            return True
    return False

def _set_paragraph_after_heading(doc, heading_text, new_text):
    """
    Find a paragraph whose text == heading_text, then replace the next
    non-empty paragraph with new_text. If next is empty, it will still be replaced.
    """
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip() == heading_text.strip():
            # find next paragraph index
            if i + 1 < len(doc.paragraphs):
                doc.paragraphs[i + 1].text = "" if new_text is None else str(new_text)
                return True
    return False

def _set_sequence_intro(doc, intro_text):
    """
    In the template, there's a paragraph under "Sequence of Events" that contains
    the log source/path. We'll overwrite that paragraph.
    """
    # In your template it is right after "Sequence of Events" heading.
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip() == "Sequence of Events":
            if i + 1 < len(doc.paragraphs):
                doc.paragraphs[i + 1].text = "" if intro_text is None else str(intro_text)
                return True
    return False

def _parse_sequence_text_to_rows(raw_text):
    """
    Accepts pasted logs in either:
      date<TAB or many spaces>time<TAB>category<TAB>message
    or CSV-like lines: date,time,category,message
    Returns list of dicts with keys: date, time, category, message
    """
    rows = []
    if not raw_text.strip():
        return rows

    for line in raw_text.splitlines():
        line = line.strip()
        if not line:
            continue

        # Try comma-separated first (but allow commas inside message by maxsplit=3)
        if "," in line:
            parts = [p.strip() for p in line.split(",", 3)]
            if len(parts) == 4:
                rows.append(
                    {"date": parts[0], "time": parts[1], "category": parts[2], "message": parts[3]}
                )
                continue

        # Fall back to whitespace/tab split into 4 columns
        parts = re.split(r"\s{2,}|\t+", line, maxsplit=3)
        if len(parts) == 4:
            rows.append(
                {"date": parts[0], "time": parts[1], "category": parts[2], "message": parts[3]}
            )
        else:
            # If format is messy, put whole line in message
            rows.append({"date": "", "time": "", "category": "", "message": line})

    return rows

def _fill_sequence_table(table, sequence_rows):
    """
    Sequence table in your template has 4 columns:
    Date | Time | (Type) | (Message)
    We'll wipe existing rows and re-add from sequence_rows.
    """
    _clear_table_rows_except_header(table, header_rows=0)  # your sequence table has no header row
    for r in sequence_rows:
        row_cells = table.add_row().cells
        row_cells[0].text = str(r.get("date", ""))
        row_cells[1].text = str(r.get("time", ""))
        row_cells[2].text = str(r.get("category", ""))
        row_cells[3].text = str(r.get("message", ""))

def _fill_actions_table(table, actions_df: pd.DataFrame):
    """
    Actions table has header row:
    Date | Time | Performed by | Action | Result
    We'll keep header row and rebuild the rest.
    """
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

    # Tables based on your template:
    # 0: Reported by / Position / Date of Report / Incident No.
    # 1: Incident Details (Date/Time/Location/Current Status)
    # 2: Sequence of Events (4 cols, many rows)
    # 3: Response and Actions Taken (header + rows)
    t0, t1, t2, t3 = doc.tables[0], doc.tables[1], doc.tables[2], doc.tables[3]

    # Fill table 0
    _set_2col_table_value(t0, "Reported by", data["reported_by"])
    _set_2col_table_value(t0, "Position", data["position"])
    _set_2col_table_value(t0, "Date of Report", data["date_of_report"])
    _set_2col_table_value(t0, "Incident No.", data["incident_no"])

    # Fill table 1
    _set_2col_table_value(t1, "Date (YYYY-MM-DD)", data["incident_date"])
    _set_2col_table_value(t1, "Time", data["incident_time"])
    _set_2col_table_value(t1, "Location", data["location"])
    _set_2col_table_value(t1, "Current Status", data["current_status"])

    # Paragraph sections (replace the paragraph right after headings)
    _set_paragraph_after_heading(doc, "Nature of Incident", data["nature_of_incident"])
    _set_sequence_intro(doc, data["sequence_intro"])
    _set_paragraph_after_heading(doc, "Damages Incurred (if any)", data["damages_incurred"])
    _set_paragraph_after_heading(doc, "Investigation and Analysis", data["investigation"])
    _set_paragraph_after_heading(doc, "Conclusion and Recommendations", data["conclusion"])

    # Sequence of events
    seq_rows = data["sequence_rows"]
    _fill_sequence_table(t2, seq_rows)

    # Response/actions
    _fill_actions_table(t3, data["actions_df"])

    # Save to bytes
    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out.read()

def convert_docx_to_pdf_bytes(docx_bytes):
    """
    Uses LibreOffice headless conversion. Works best on Linux servers.
    If LibreOffice isn't installed, this will raise.
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = os.path.join(tmpdir, "report.docx")
        with open(docx_path, "wb") as f:
            f.write(docx_bytes)

        cmd = [
            "soffice",
            "--headless",
            "--nologo",
            "--nofirststartwizard",
            "--convert-to",
            "pdf",
            "--outdir",
            tmpdir,
            docx_path,
        ]
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        pdf_path = os.path.join(tmpdir, "report.pdf")
        with open(pdf_path, "rb") as f:
            return f.read()


# -------------------------
# Streamlit UI
# -------------------------
st.set_page_config(page_title="Incident Report Generator", layout="wide")
st.title("Incident Report Generator (DOCX / PDF)")

st.caption("Fills your existing incident report template and exports Word/PDF.")

with st.form("incident_form"):
    c1, c2 = st.columns(2)

    with c1:
        reported_by = st.text_input("Reported by", value="")
        position = st.text_input("Position", value="")
        date_of_report = st.date_input("Date of Report", value=date.today()).strftime("%Y-%m-%d")
        incident_no = st.text_input("Incident No.", value="SMCOD-IR-GS-____-____-____")

    with c2:
        incident_date = st.date_input("Incident Date", value=date.today()).strftime("%Y-%m-%d")
        incident_time = st.text_input("Incident Time (HH:MM:SS)", value=datetime.now().strftime("%H:%M:%S"))
        location = st.text_input("Location", value="")
        current_status = st.selectbox("Current Status", ["Resolved", "Ongoing", "Monitoring", "Open"], index=0)

    st.subheader("Incident Description")
    nature_of_incident = st.text_area("Nature of Incident", height=120)

    st.subheader("Sequence of Events")
    sequence_intro = st.text_input(
        "Sequence Source/Intro (shows under 'Sequence of Events')",
        value="These are taken from event logs found at: /var/log/track/acu/events/<logfile>.log.txt",
    )

    st.write("Paste logs here (one per line). Format supported:")
    st.code("date  time  category  message   (tab or 2+ spaces)\nOR\ndate,time,category,message")
    seq_text = st.text_area("Sequence raw logs", height=200)

    st.subheader("Damages / Investigation / Conclusion")
    damages_incurred = st.text_area("Damages Incurred (if any)", height=80, value="None")
    investigation = st.text_area("Investigation and Analysis", height=120)
    conclusion = st.text_area("Conclusion and Recommendations", height=120)

    st.subheader("Response and Actions Taken (table)")
    default_actions = pd.DataFrame(
        [
            {"Date": incident_date, "Time": "09:00", "Performed by": "", "Action": "", "Result": ""},
        ]
    )
    actions_df = st.data_editor(default_actions, num_rows="dynamic", use_container_width=True)

    export_pdf = st.checkbox("Also export PDF (requires LibreOffice on server)", value=False)
    submitted = st.form_submit_button("Generate Report")

if submitted:
    if not os.path.exists(TEMPLATE_PATH):
        st.error(f"Template not found: {TEMPLATE_PATH}\nPut the .docx beside app.py.")
        st.stop()

    sequence_rows = _parse_sequence_text_to_rows(seq_text)

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
        "sequence_intro": sequence_intro,
        "sequence_rows": sequence_rows,
        "damages_incurred": damages_incurred,
        "investigation": investigation,
        "conclusion": conclusion,
        "actions_df": actions_df,
    }

    docx_bytes = generate_docx_bytes(TEMPLATE_PATH, data)

    safe_no = re.sub(r"[^A-Za-z0-9_-]+", "_", incident_no.strip()) or "incident_report"
    docx_name = f"{safe_no}.docx"

    st.success("Report generated.")
    st.download_button(
        "Download Word (.docx)",
        data=docx_bytes,
        file_name=docx_name,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

    if export_pdf:
        try:
            pdf_bytes = convert_docx_to_pdf_bytes(docx_bytes)
            st.download_button(
                "Download PDF (.pdf)",
                data=pdf_bytes,
                file_name=f"{safe_no}.pdf",
                mime="application/pdf",
            )
        except Exception as e:
            st.warning(
                "PDF conversion failed. If you're on Linux, install LibreOffice (`soffice`). "
                f"Error: {e}"
            )
