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


class LiveLog(io.TextIOBase):
    """Mirrors the engine's print() output into the page while it runs.

    Without this the whole run is invisible behind a spinner: the engine can
    spend several minutes downloading, OCR-ing and calling Gemini, and a
    silent spinner is indistinguishable from a hang. Renders are throttled so
    a chatty engine does not flood the websocket.
    """

    MAX_LINES = 16
    MIN_INTERVAL = 0.35  # seconds between re-renders

    def __init__(self, placeholder):
        self.placeholder = placeholder
        self.lines: list[str] = []
        self._partial = ""
        self._last_render = 0.0

    def write(self, s):
        self._partial += s
        while "\n" in self._partial:
            line, self._partial = self._partial.split("\n", 1)
            self.lines.append(line)
        now = time.monotonic()
        if now - self._last_render >= self.MIN_INTERVAL:
            self._last_render = now
            self._render()
        return len(s)

    def flush(self):
        self._render()

    def _render(self):
        tail = [ln for ln in self.lines[-self.MAX_LINES:]]
        if self._partial:
            tail.append(self._partial)
        self.placeholder.code("\n".join(tail) or "Starting...", language=None)

    def getvalue(self):
        return "\n".join(self.lines)


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
st.caption(
    "With a key the freeform documents read much better, but the run takes "
    "noticeably longer. Without one it still reads GeM bids well. Either way "
    "the live progress below tells you exactly what it is doing."
)

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
            # --- tuned down for a free shared cloud container ---------------
            # 300 dpi pixmaps are memory-hungry and this tier has little RAM;
            # 200 dpi still reads printed tender scans well.
            "ocr_dpi": 200,
            "ocr_max_pages_per_file": 25,
            # Fail fast rather than appear frozen. The engine's own backoff
            # can otherwise wait minutes per busy model, and on a web page a
            # long silence reads as a crash.
            "gemini_retries": 3,
        })

        st.info("Working. This normally takes a few minutes — longer with a "
                "Gemini key or scanned PDFs. You can leave this tab open.")
        progress_box = st.empty()
        live = LiveLog(progress_box)

        try:
            with st.spinner("Reading your tender documents..."):
                with contextlib.redirect_stdout(live):
                    te.run()
            live.flush()
        except Exception as e:
            live.flush()
            st.session_state.result = None
            st.error(f"Something went wrong: {e}")
            with st.expander("Technical log", expanded=True):
                st.code(live.getvalue() or "(no output captured)")
            st.stop()

        dash_html = te.build_dashboard_html(
            te.LAST_ROWS, standalone=True,
            stamp="Run " + time.strftime("%d %b %Y, %H:%M"))
        excel_path = Path("Tender_Summary.xlsx")
        excel_bytes = excel_path.read_bytes() if excel_path.exists() else None
        st.session_state.result = (dash_html, excel_bytes, live.getvalue())
        progress_box.empty()

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
