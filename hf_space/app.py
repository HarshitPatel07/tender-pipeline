"""Tender Pipeline - AI-Powered Tender Document Extractor & Dashboard.

Runs locally or on Streamlit Community Cloud.
Extracts 13 critical tender fields using Groq 120B & Gemini Flash AI engines
with page-level citations, dual-reader verification, and Excel/Dashboard export.
"""
from __future__ import annotations

import contextlib
import io
import os
import re
import time
from pathlib import Path

import streamlit as st
import tender_extractor as te

# Page configuration
st.set_page_config(
    page_title="Tender Pipeline AI",
    page_icon="📑",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom CSS for rich aesthetics
st.markdown(
    """
    <style>
    /* Modern typography and container spacing */
    .main-header {
        font-size: 2.2rem;
        font-weight: 800;
        color: #1E293B;
        margin-bottom: 0.2rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }
    .sub-header {
        font-size: 1.05rem;
        color: #64748B;
        margin-bottom: 1.5rem;
    }
    .stat-card {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 1rem 1.25rem;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .stat-val {
        font-size: 1.8rem;
        font-weight: 700;
        color: #0F172A;
    }
    .stat-lbl {
        font-size: 0.85rem;
        font-weight: 600;
        color: #64748B;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    .badge-success {
        background-color: #DEF7EC;
        color: #03543F;
        padding: 0.2rem 0.5rem;
        border-radius: 6px;
        font-size: 0.8rem;
        font-weight: 600;
    }
    .badge-amber {
        background-color: #FEF08A;
        color: #713F12;
        padding: 0.2rem 0.5rem;
        border-radius: 6px;
        font-size: 0.8rem;
        font-weight: 600;
    }
    .badge-danger {
        background-color: #FEE2E2;
        color: #991B1B;
        padding: 0.2rem 0.5rem;
        border-radius: 6px;
        font-size: 0.8rem;
        font-weight: 600;
    }
    .css-1y4p8pa { padding: 1.5rem 1rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

if "result" not in st.session_state:
    st.session_state.result = None


class ProgressLog(io.TextIOBase):
    """Streams engine output with real-time progress parsing."""

    def __init__(self, progress_container, log_container):
        self.progress_container = progress_container
        self.log_container = log_container
        self.lines: list[str] = []
        self._partial = ""
        self._last_render = 0.0
        self.start_time = time.monotonic()
        self.current_file = None
        self.current_step = "Initializing..."
        self.files_done = 0
        self.total_files = None

    def write(self, s):
        self._partial += s
        while "\n" in self._partial:
            line, self._partial = self._partial.split("\n", 1)
            self.lines.append(line)
            self._parse_line(line)

        now = time.monotonic()
        if now - self._last_render >= 0.4:
            self._last_render = now
            self._render()
        return len(s)

    def flush(self):
        self._render()

    def _parse_line(self, line):
        low = line.lower()
        if "reading" in low and ".pdf" in low:
            match = re.search(r"reading\s+([^.]+\.pdf)", line, re.I)
            if match:
                self.current_file = match.group(1)
                self.current_step = "📄 Reading PDF Document"
        elif "reading" in low and ".docx" in low:
            self.current_step = "📄 Reading Word Document"
        elif "ocr" in low or "scanned" in low:
            self.current_step = "🔍 Processing Document Pages (OCR / Layout)"
        elif "groq" in low:
            self.current_step = "🤖 Groq 120B AI Extraction"
        elif "gemini" in low:
            self.current_step = "✨ Gemini Flash AI Extraction"
        elif "ai pass" in low:
            self.current_step = "🤖 Running Deep AI Extraction"
        elif "writing excel" in low:
            self.current_step = "📊 Generating Workbook & Evidence"

        if "tender(s):" in low:
            m = re.search(r"(\d+)\s+tender", line, re.I)
            if m:
                self.total_files = int(m.group(1))

        if "done in" in low:
            self.files_done = self.total_files or 1

    def _render(self):
        elapsed = int(time.monotonic() - self.start_time)
        elapsed_str = f"{elapsed//60}m {elapsed%60}s" if elapsed >= 60 else f"{elapsed}s"

        with self.progress_container.container():
            col1, col2 = st.columns([3, 1])
            with col1:
                st.markdown(f"**Status:** `{self.current_step}`")
                if self.current_file:
                    st.caption(f"Active file: {self.current_file}")
            with col2:
                st.markdown(f"**Elapsed:** `{elapsed_str}`")

        with self.log_container.container():
            tail = self.lines[-6:]
            st.code("\n".join(tail) or "Processing...", language="text")

    def getvalue(self):
        return "\n".join(self.lines)


# Sidebar Configuration
with st.sidebar:
    st.markdown("### ⚙️ AI Engine Settings")

    # Load default keys from environment
    default_groq = os.environ.get("GROQ_API_KEY", "")
    default_gemini = os.environ.get("GEMINI_API_KEY", "")

    ai_provider = st.selectbox(
        "Primary AI Engine",
        options=["auto", "groq", "gemini", "both"],
        format_func=lambda x: {
            "auto": "🚀 Auto (Groq 120B + Gemini Flash)",
            "groq": "⚡ Groq (120B Model - Fast & Free)",
            "gemini": "✨ Google Gemini Flash",
            "both": "🛡️ Dual Mode (Cross-Check Both)",
        }[x],
        index=0,
        help="Select which AI models analyze the tender documents.",
    )

    groq_key_input = st.text_input(
        "Groq API Key",
        value=default_groq,
        type="password",
        help="Free API key from console.groq.com",
    )

    gemini_key_input = st.text_input(
        "Gemini API Key",
        value=default_gemini,
        type="password",
        help="Free key from aistudio.google.com/apikey",
    )

    st.markdown("---")
    st.markdown("### 🛠️ Execution Options")
    require_ai = st.checkbox(
        "Compulsory AI Extraction",
        value=True,
        help="Ensures deep AI extraction is always performed rather than basic regex.",
    )
    fresh_run = st.checkbox(
        "Force Fresh Extraction",
        value=False,
        help="Ignore cached tender runs and re-parse all documents from scratch.",
    )

    st.markdown("---")
    st.caption("💡 **Tip:** Groq 120B provides instant extraction for large tender PDFs without hitting quota limits.")


# Main App Header
st.markdown('<div class="main-header">📑 Tender Pipeline AI</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-header">'
    "Automated end-to-end tender document extraction with multi-model AI, "
    "page-level evidence citations, and Excel/Dashboard export."
    "</div>",
    unsafe_allow_html=True,
)

# Input Section
col_input, col_mode = st.columns([3, 1])

with col_mode:
    source_type = st.radio("Source Location", ["Google Drive Link", "Local Folder"], index=0)

with col_input:
    if source_type == "Google Drive Link":
        folder_input = st.text_input(
            "Google Drive Folder Link",
            placeholder="https://drive.google.com/drive/folders/...",
            help="Paste the link to your tender folder. Make sure link sharing is set to 'Anyone with the link can view'.",
        )
    else:
        folder_input = st.text_input(
            "Local Folder Path",
            value="tender_work/docs",
            help="Relative or absolute path to the local folder containing tender documents.",
        )

btn_extract = st.button("🚀 Extract Tender Details", type="primary", use_container_width=True)

if btn_extract:
    raw_path = folder_input.strip()
    if not raw_path:
        st.error("Please provide a valid Google Drive folder link or local directory path.")
    elif source_type == "Google Drive Link" and "drive.google.com" not in raw_path:
        st.error("Invalid link format. Please provide a standard Google Drive folder URL.")
    else:
        # Configure tender extractor
        groq_k = groq_key_input.strip()
        gemini_k = gemini_key_input.strip()

        te.CONFIG.update({
            "source_mode": "drive_link" if source_type == "Google Drive Link" else "local_folder",
            "drive_folder_url": raw_path if source_type == "Google Drive Link" else "",
            "local_folder": raw_path if source_type != "Google Drive Link" else "",
            "output_xlsx": "Tender_Summary.xlsx",
            "output_html": "Tender_Dashboard.html",
            "work_dir": "tender_work",
            "groq_api_key": groq_k,
            "use_groq": bool(groq_k),
            "gemini_api_key": gemini_k,
            "use_gemini": bool(gemini_k),
            "ai_provider": ai_provider,
            "require_ai": require_ai,
            "use_cache": not fresh_run,
            "ocr_dpi": 200,
            "ocr_max_pages_per_file": 40,
        })

        progress_container = st.empty()
        log_container = st.empty()
        progress_log = ProgressLog(progress_container, log_container)

        try:
            with st.spinner("Extracting tender fields with AI engine..."):
                with contextlib.redirect_stdout(progress_log):
                    out_path = te.run()
            progress_log.flush()
        except Exception as e:
            progress_log.flush()
            st.error(f"Extraction failed: {e}")
            with st.expander("Show Technical Execution Log", expanded=True):
                st.code(progress_log.getvalue() or "No output captured.")
            st.stop()

        # Build output assets
        dash_html = te.build_dashboard_html(
            te.LAST_ROWS,
            standalone=True,
            stamp="Extracted on " + time.strftime("%d %b %Y, %H:%M"),
        )
        xlsx_file = Path("Tender_Summary.xlsx")
        xlsx_data = xlsx_file.read_bytes() if xlsx_file.exists() else None

        st.session_state.result = {
            "rows": te.LAST_ROWS,
            "dash_html": dash_html,
            "excel_bytes": xlsx_data,
            "log": progress_log.getvalue(),
        }
        progress_container.empty()
        log_container.empty()
        st.success("✅ Extraction completed successfully!")


# Results View
if st.session_state.result:
    res = st.session_state.result
    rows = res.get("rows", [])
    dash_html = res.get("dash_html", "")
    excel_bytes = res.get("excel_bytes")
    log_text = res.get("log", "")

    st.markdown("---")

    # Metrics Overview
    total_tenders = len(rows)
    total_fields = total_tenders * len(te.FIELDS)
    found_fields = sum(
        1 for r in rows for k, _ in te.FIELDS if r["results"][k].value != te.NOT_FOUND
    )
    flagged_fields = sum(
        1 for r in rows for k, _ in te.FIELDS if r["results"][k].flag
    )
    accuracy_pct = int((found_fields / max(1, total_fields)) * 100)

    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.markdown(
            f'<div class="stat-card"><div class="stat-val">{total_tenders}</div>'
            f'<div class="stat-lbl">Tenders Processed</div></div>',
            unsafe_allow_html=True,
        )
    with m2:
        st.markdown(
            f'<div class="stat-card"><div class="stat-val">{found_fields} / {total_fields}</div>'
            f'<div class="stat-lbl">Fields Extracted ({accuracy_pct}%)</div></div>',
            unsafe_allow_html=True,
        )
    with m3:
        st.markdown(
            f'<div class="stat-card"><div class="stat-val">{flagged_fields}</div>'
            f'<div class="stat-lbl">Needs Human Eye (Amber)</div></div>',
            unsafe_allow_html=True,
        )
    with m4:
        st.markdown(
            f'<div class="stat-card"><div class="stat-val">13 / 13</div>'
            f'<div class="stat-lbl">Field Schema Depth</div></div>',
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)

    # Action Toolbar
    col_dl1, col_dl2, _ = st.columns([1.5, 1.5, 3])
    with col_dl1:
        if excel_bytes:
            st.download_button(
                label="📥 Download Excel (.xlsx)",
                data=excel_bytes,
                file_name="Tender_Summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    with col_dl2:
        if dash_html:
            st.download_button(
                label="🌐 Download Dashboard (.html)",
                data=dash_html,
                file_name="Tender_Dashboard.html",
                mime="text/html",
                use_container_width=True,
            )

    # Tabs for Data View
    tab_table, tab_dashboard, tab_logs = st.tabs([
        "📋 Extracted Tender Summary",
        "🖥️ Interactive Visual Dashboard",
        "📜 Technical Audit Log",
    ])

    with tab_table:
        for idx, r in enumerate(rows, start=1):
            t_name = r["results"].get("tender_name", te.Result()).value
            if not t_name or t_name == te.NOT_FOUND:
                t_name = r.get("tender", f"Tender {idx}")

            with st.expander(f"📌 {t_name}", expanded=True):
                field_data = []
                for k, label in te.FIELDS:
                    res_obj = r["results"].get(k, te.Result())
                    val = res_obj.value
                    flag = res_obj.flag
                    ref = res_obj.ref or "N/A"
                    status = "✅ Verified"
                    if flag:
                        status = f"⚠️ {flag}"
                    elif val == te.NOT_FOUND:
                        status = "❌ Not Found"

                    field_data.append({
                        "Field": label,
                        "Extracted Value": val,
                        "Source Page Citation": ref,
                        "Status / Confidence": status,
                    })

                st.table(field_data)

    with tab_dashboard:
        if dash_html:
            st.components.v1.html(dash_html, height=800, scrolling=True)

    with tab_logs:
        st.code(log_text or "No log available.", language="text")
