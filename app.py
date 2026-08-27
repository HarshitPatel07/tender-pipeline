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
import re

import streamlit as st

import tender_extractor as te

st.set_page_config(page_title="Tender Pipeline", page_icon="🧾", layout="centered")

if "result" not in st.session_state:
    st.session_state.result = None  # (dash_html, excel_bytes, log_text) | None


class ProgressLog(io.TextIOBase):
    """Streams engine output with structured progress tracking.

    Parses key milestones (downloading, reading PDFs, OCR, AI extraction)
    and displays progress bars with timing and current file info.
    """

    def __init__(self, progress_container, log_container):
        self.progress_container = progress_container
        self.log_container = log_container
        self.lines: list[str] = []
        self._partial = ""
        self._last_render = 0.0
        self.start_time = time.monotonic()
        self.current_file = None
        self.current_step = "Starting..."
        self.files_done = 0
        self.total_files = None

    def write(self, s):
        self._partial += s
        while "\n" in self._partial:
            line, self._partial = self._partial.split("\n", 1)
            self.lines.append(line)
            self._parse_line(line)

        now = time.monotonic()
        if now - self._last_render >= 0.5:
            self._last_render = now
            self._render()
        return len(s)

    def flush(self):
        self._render()

    def _parse_line(self, line):
        """Extract progress info from engine output lines."""
        # Detect which file is being processed
        if "Reading" in line and ".pdf" in line.lower():
            match = re.search(r"Reading\s+([^.]+\.pdf)", line, re.IGNORECASE)
            if match:
                self.current_file = match.group(1)
                self.current_step = "📄 Reading PDF"
        elif "reading" in line.lower() and "docx" in line.lower():
            match = re.search(r"([^\s]+\.docx)", line, re.IGNORECASE)
            if match:
                self.current_file = match.group(1)
                self.current_step = "📄 Reading DOCX"
        elif "reading" in line.lower() and ("xls" in line.lower() or "excel" in line.lower()):
            match = re.search(r"([^\s]+\.xls[x]?)", line, re.IGNORECASE)
            if match:
                self.current_file = match.group(1)
                self.current_step = "📊 Reading Excel"
        elif "ocr" in line.lower() or "tesseract" in line.lower():
            self.current_step = "🔍 OCR scanning"
        elif "gemini" in line.lower() or "ai" in line.lower():
            self.current_step = "🤖 AI extraction"
        elif "merge" in line.lower() or "comparing" in line.lower():
            self.current_step = "✓ Verifying data"
        elif "cached" in line.lower():
            self.current_step = "💾 Using cache"

        # Track tender count
        if "tender" in line.lower() and "total" in line.lower():
            match = re.search(r"(\d+)\s+tender", line, re.IGNORECASE)
            if match:
                self.total_files = int(match.group(1))

        # Track completion
        if "writing" in line.lower() and ("excel" in line.lower() or "xlsx" in line.lower()):
            self.files_done = self.total_files or self.files_done + 1

    def _render(self):
        """Display progress with bars, timing, and current file."""
        elapsed = int(time.monotonic() - self.start_time)
        elapsed_str = f"{elapsed//60}m {elapsed%60}s" if elapsed >= 60 else f"{elapsed}s"

        with self.progress_container.container():
            col1, col2 = st.columns([3, 1])
            with col1:
                st.markdown(f"**{self.current_step}**")
                if self.current_file:
                    st.caption(f"📁 {self.current_file}")
            with col2:
                st.caption(f"⏱️ {elapsed_str}")

            # Progress bar if we know totals
            if self.total_files and self.total_files > 0:
                progress = min(self.files_done / self.total_files, 1.0)
                st.progress(progress, text=f"{self.files_done}/{self.total_files} tenders")

        # Recent log lines
        with self.log_container.container():
            tail = self.lines[-8:]
            log_text = "\n".join(tail)
            st.code(log_text or "Starting...", language=None)

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

fresh = st.checkbox(
    "Force a fresh read (ignore anything remembered from a previous run)")
st.caption(
    "Results are remembered between runs so a repeat is fast, and a tender is "
    "re-read automatically whenever its documents change in Drive. Tick this "
    "only if you want every document read again from scratch regardless."
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
            "use_cache": not fresh,
            "ocr_dpi": 200,
            "ocr_max_pages_per_file": 25,
            "gemini_retries": 3,
        })

        st.info("Working. This normally takes a few minutes — longer with a "
                "Gemini key or scanned PDFs. You can leave this tab open.")
        progress_box = st.empty()
        log_box = st.empty()
        progress_log = ProgressLog(progress_box, log_box)

        try:
            with st.spinner("Reading your tender documents..."):
                with contextlib.redirect_stdout(progress_log):
                    te.run()
            progress_log.flush()
        except Exception as e:
            progress_log.flush()
            st.session_state.result = None
            st.error(f"Something went wrong: {e}")
            with st.expander("Technical log", expanded=True):
                st.code(progress_log.getvalue() or "(no output captured)")
            st.stop()

        dash_html = te.build_dashboard_html(
            te.LAST_ROWS, standalone=True,
            stamp="Run " + time.strftime("%d %b %Y, %H:%M"))
        excel_path = Path("Tender_Summary.xlsx")
        excel_bytes = excel_path.read_bytes() if excel_path.exists() else None
        st.session_state.result = (dash_html, excel_bytes, progress_log.getvalue())
        progress_box.empty()
        log_box.empty()

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
