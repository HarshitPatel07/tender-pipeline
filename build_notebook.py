"""Packs tender_extractor.py into a single-button Colab notebook.

Design goal: the person running this has never seen code and never will.
Everything collapses into ONE cell behind a form (two boxes: Drive link,
Gemini key) and ONE play button. The dashboard renders right there in the
page afterwards - no file to find in Downloads, no second notebook cell to
remember to run in order.
"""
import base64
import json
from pathlib import Path

HERE = Path(__file__).parent
ENGINE = (HERE / "tender_extractor.py").read_text(encoding="utf-8")
ENGINE_B64 = base64.b64encode(ENGINE.encode("utf-8")).decode("ascii")


def md(text):
    return {"cell_type": "markdown", "metadata": {},
            "source": text.strip("\n").splitlines(keepends=True)}


def code(text):
    return {"cell_type": "code", "metadata": {}, "execution_count": None,
            "outputs": [], "source": text.strip("\n").splitlines(keepends=True)}


CELL_INTRO = """
# Tender Pipeline

**What this does:** paste your tender folder's Google Drive link below, press
the one Play button, and wait. Your dashboard appears right on this page -
no downloading, no other apps, nothing to install.

### The three steps
1. Click the little arrow/Play button on the grey box below to expand it if
   it's collapsed.
2. Paste your Drive folder link into the first box. The second box (Gemini
   key) is optional - leave it blank the first time if you don't have one yet.
3. Press the **&#9654; Play button** on the left of that box, then wait.
   You'll see technical-looking text scroll by while it reads your documents
   - that's normal, it's just showing its work. When it says **DONE**, your
   dashboard is right below it on this same page.

### Two things worth knowing
- **`WORKING FOLDER` is skipped on purpose** in every tender - it holds your
  own draft submissions (covering letter, financial bid), not the
  department's tender documents.
- **Amber needs your eye, red means not found.** Every figure on the
  dashboard shows which file and page it came from - open *Where each value
  came from* on any card to check it. This removes the typing, not the
  review: always confirm EMD, fees and the deadline before acting on them.

**Cost: zero.** This page (Google Colab) is free, the document reading is
free, and the free Gemini key is free.

**One click to come back later:** after this first run, use **File -&gt; Save a
copy in Drive**, then bookmark that copy. Next time, open the bookmark and
you're straight back to these same two boxes.
"""

CELL_RUN = '''
#@title \u25b6 Paste your tender folder link, then press the Play button on the left { display-mode: "form" }
#@markdown &nbsp;
DRIVE_LINK = "https://drive.google.com/drive/folders/1qyGp7N8HVGynwn0HNFTiYHhF6KN5rNcc"  #@param {type:"string"}
#@markdown Optional - a free key from **aistudio.google.com/apikey** gives much better results on freeform documents (not just GeM). Leave blank to run without it.
GEMINI_KEY = ""  #@param {type:"string"}

import base64, sys, time
from pathlib import Path
from IPython.display import HTML, display
import html as _html

print("Setting up (about a minute the first time this session) ...")
get_ipython().system('pip install -q pymupdf gdown python-docx openpyxl pandas xlrd requests pytesseract pillow google-genai')
get_ipython().system('apt-get -qq install -y tesseract-ocr > /dev/null 2>&1')

Path("tender_extractor.py").write_text(
    base64.b64decode("__ENGINE_B64__").decode("utf-8"), encoding="utf-8")

sys.path.insert(0, ".")
import importlib
import tender_extractor as te
importlib.reload(te)

te.CONFIG.update({
    "source_mode": "drive_link",
    "drive_folder_url": DRIVE_LINK,
    "output_xlsx": "Tender_Summary.xlsx",
    "output_html": "Tender_Dashboard.html",
    "gemini_api_key": GEMINI_KEY.strip(),
    "use_gemini": bool(GEMINI_KEY.strip()),
    "work_dir": "tender_work",
    "use_cache": True,
})

print("-" * 60)
print("Reading your tender documents now. This can take a few minutes")
print("for a large folder or scanned PDFs - the text below is normal.")
print("-" * 60)
te.run()

print("\\nBuilding your dashboard ...")
dash_html = te.build_dashboard_html(
    te.LAST_ROWS, standalone=True,
    stamp="Run " + time.strftime("%d %b %Y, %H:%M"))
escaped = _html.escape(dash_html, quote=True)

display(HTML(f"""
<script>
  window.addEventListener('message', function(e) {{
    if (e.data && e.data.tenderDashHeight) {{
      var f = document.getElementById('tenderFrame');
      if (f) f.style.height = Math.min(e.data.tenderDashHeight + 24, 4000) + 'px';
    }}
  }});
</script>
<div style="border:2px solid #14655A;border-radius:6px;overflow:hidden;margin-top:14px;">
  <iframe id="tenderFrame" srcdoc="{escaped}"
    style="width:100%;height:640px;border:none;display:block;"></iframe>
</div>
"""))

from google.colab import files
files.download("Tender_Dashboard.html")
files.download("Tender_Summary.xlsx")
print("\\nDONE. Your dashboard is above. Both files were also saved to your")
print("computer's Downloads folder as a backup.")
'''.replace("__ENGINE_B64__", ENGINE_B64)

CELL_LOCAL = """
## If you'd rather run this on your own PC (no browser needed)

Needs a one-time Python install. After that:

```bash
pip install pymupdf gdown python-docx openpyxl pandas xlrd requests pytesseract pillow google-genai
python tender_extractor.py --drive "<your drive link>" --key "<gemini key>"
```

Or point it at a folder already on disk, works fully offline:

```bash
python tender_extractor.py --folder "D:\\Tenders\\Master Tender Uploads" --no-ai
```

Re-running is cheap either way: a tender whose files have not changed is
served from cache, so only newly added tender folders get reprocessed.
"""

import ast
try:
    ast.parse(CELL_RUN)
except SyntaxError as e:
    raise SystemExit(f"CELL_RUN is not valid Python - would crash in Colab: {e}")

# A doubled brace {{ or }} is only ever correct inside the one nested
# f-string (the JS block below) - everywhere else in CELL_RUN it means a
# dict literal got wrongly wrapped in a set, e.g. te.CONFIG.update({{...}})
# calling .update() with a single (unhashable) dict inside a set. That is
# valid Python syntax (ast.parse above will not catch it) but fails at
# runtime - exactly what shipped and crashed on 27-Aug.
_fstring_start = CELL_RUN.index('display(HTML(f"""')
_fstring_end = CELL_RUN.index('"""))', _fstring_start) + len('"""))')
_outside = CELL_RUN[:_fstring_start] + CELL_RUN[_fstring_end:]
if "{{" in _outside or "}}" in _outside:
    raise SystemExit(
        "CELL_RUN has a doubled brace outside the nested f-string - this "
        "will crash in Colab with 'unhashable type: dict'. Check for a "
        "leftover {{ or }} left over from string-template editing.")

nb = {
    "nbformat": 4, "nbformat_minor": 0,
    "metadata": {
        "colab": {"name": "Tender_Pipeline.ipynb",
                  "provenance": [], "toc_visible": True},
        "kernelspec": {"name": "python3", "display_name": "Python 3"},
        "language_info": {"name": "python"},
    },
    "cells": [
        md(CELL_INTRO),
        code(CELL_RUN),
        md(CELL_LOCAL),
    ],
}

out = HERE / "Tender_Pipeline.ipynb"
out.write_text(json.dumps(nb, indent=1, ensure_ascii=False), encoding="utf-8")
print(f"wrote {out}  ({out.stat().st_size:,} bytes, {len(nb['cells'])} cells, "
      f"engine embedded as {len(ENGINE_B64):,} base64 chars)")
