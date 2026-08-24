"""Tender Pipeline - a plain website version of the tender extractor.

Runs on Streamlit Community Cloud (free), deployed straight from GitHub.
Paste a Drive link, optionally a Gemini key, press Extract, the dashboard
renders on the page. No notebook, no code visible, no menus - a page with a
form and a button.

All extraction logic lives in tender_extractor.py, in this same repository.
This file is only the web page around it.
"""
import contextlib
import io
import time
from pathlib import Path

import streamlit as st

import tender_extractor as te

st.set_page_config(page_title="Tender Pipeline", page_icon="🧾", layout="centered")

if "result" not in st.session_state:
    st.session_state.result = None  # (dash_html, excel_bytes, log_text) | None

st.title("Tender Pipeline")
st.markdown(
    "Paste your tender folder's Google Drive link below, press **Extract "
    "tenders**, and wait. Your dashboard appears on this page — deadlines, "
    "amounts, eligibility, scope, all 13 fields, each traceable to its "
    "source file and page.\n\n"
    "**`WORKING FOLDER` is skipped on purpose** in every tender — it holds "
    "your own draft submissions, not the department's documents. "
    "**Amber needs your eye, red means not found** — this removes the "
    "typing, not the review; always confirm EMD, fees and the deadline "
    "before acting on them."
)

drive_link = st.text_input(
    "Google Drive folder link",
    placeholder="https://drive.google.com/drive/folders/...")
gemini_key = st.text_input(
    "Gemini API key (optional — free at aistudio.google.com/apikey)",
    type="password", placeholder="Leave blank to run without AI")

go = st.button("Extract tenders", type="primary")

if go:
    link = drive_link.strip()
    if not link:
        st.error("Paste a Google Drive folder link first — the box above is empty.")
    elif "drive.google.com" not in link:
        st.error("That doesn't look like a Google Drive link. It should look "
                 "like https://drive.google.com/drive/folders/....")
    else:
        te.CONFIG.update({
            "source_mode": "drive_link",
            "drive_folder_url": link,
            "output_xlsx": "Tender_Summary.xlsx",
            "output_html": "Tender_Dashboard.html",
            "gemini_api_key": gemini_key.strip(),
            "use_gemini": bool(gemini_key.strip()),
            "work_dir": "tender_work",
            "use_cache": True,
        })
        log = io.StringIO()
        with st.spinner("Reading your tender documents — this can take a "
                        "few minutes for a large folder or scanned PDFs..."):
            try:
                with contextlib.redirect_stdout(log):
                    te.run()
            except Exception as e:
                st.session_state.result = None
                st.error(f"Something went wrong: {e}")
                with st.expander("Technical log"):
                    st.code(log.getvalue() or "(no output captured)")
                st.stop()

        dash_html = te.build_dashboard_html(
            te.LAST_ROWS, standalone=True,
            stamp="Run " + time.strftime("%d %b %Y, %H:%M"))
        excel_path = Path("Tender_Summary.xlsx")
        excel_bytes = excel_path.read_bytes() if excel_path.exists() else None
        st.session_state.result = (dash_html, excel_bytes, log.getvalue())

if st.session_state.result:
    dash_html, excel_bytes, log_text = st.session_state.result
    if excel_bytes:
        st.download_button(
            "Download Tender_Summary.xlsx", data=excel_bytes,
            file_name="Tender_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.iframe(dash_html, height="content")
    with st.expander("Technical log"):
        st.code(log_text or "(no output captured)")
