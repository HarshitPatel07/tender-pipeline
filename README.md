# Tender Pipeline

Paste a Google Drive tender folder link into one box, press one Play button,
and a dashboard of all its tenders appears on the same page — deadlines,
amounts, eligibility, scope, all 13 fields, each traceable to its source file
and page. No coding, no installing anything, free.

| Field | |
|---|---|
| 1 Tender Name | 8 Scope of Work |
| 2 Location / Address | 9 Penalty |
| 3 Purpose / Audit type | 10 Tender EMD |
| 4 Period | 11 Tender SD |
| 5 Tender Estimated cost | 12 Tender Fees |
| 6 Assignment fees | 13 Tender submission date |
| 7 Eligibility Criteria | |

## For non-technical use — start here

1. Open [Google Colab](https://colab.research.google.com) → **File → Upload
   notebook** → pick **`Tender_Pipeline.ipynb`**.
2. Paste your Drive folder link into the box in the one grey cell. The Gemini
   key box is optional (free key: <https://aistudio.google.com/apikey>) —
   leave it blank the first time.
3. Press the **▶ Play button** on that cell. Wait — you'll see technical text
   scroll by, that's normal. When it says **DONE**, your dashboard is right
   there on the page: search, filters, deadlines sorted to the top, and the
   source file/page behind every figure.
4. **File → Save a copy in Drive**, bookmark that copy. Next time it's just:
   open the bookmark, paste the new folder link, press Play.

The Drive folder must be shared as **Anyone with the link – Viewer**.

## Files

| File | What it is |
|---|---|
| `Tender_Pipeline.ipynb` | **Start here.** One button. Paste a link, press Play, the dashboard renders on the page. |
| `tender_extractor.py` | The engine underneath, for running locally from a terminal. |
| `test_rules.py` | Offline self-test on synthetic tender text. No internet, no API key needed. |
| `test_regressions.py` | Guards against the specific false positives seen in real runs. Run after any change to the rules layer. |
| `dashboard_from_excel.py` | Rebuilds the HTML dashboard from an existing `Tender_Summary.xlsx`, without re-reading the PDFs. |
| `SAMPLE_Tender_Summary.xlsx` | Example output, so you can see the format before running. |
| `build_notebook.py` | Rebuilds `Tender_Pipeline.ipynb` after you edit the engine. |

## What you get

**The dashboard, inline on the page** — deadline-sorted cards, colour by
urgency, search, and two filters (*needs review*, *closing in 14 days*). Every
card expands to show the file and page each value came from. It also saves
`Tender_Dashboard.html` and `Tender_Summary.xlsx` to your computer as a
backup — the HTML opens the same dashboard standalone in any browser, and the
Excel is the same data as a table.

## Honest limit of the one-button version

Google Colab has no way to run a cell the instant the page opens (that's a
deliberate security choice on Google's side, not something this notebook can
turn off) — pressing Play once is the minimum. If that one click is still one
too many, the only ways around it are installing Python locally (a `.bat` you
could then just double-click) or hosting this on a server permanently online,
which means your tender documents would sit on a third-party host rather than
processing entirely inside your own free Google account. Say if you want that
traded off differently.

## Run it locally

```bash
pip install pymupdf gdown python-docx openpyxl pandas xlrd requests pytesseract pillow google-genai
python tender_extractor.py --drive "<drive folder link>" --key "<gemini key>"
```

OCR for scanned PDFs needs Tesseract installed separately
(Windows: the UB-Mannheim installer; macOS: `brew install tesseract`;
Linux: `apt install tesseract-ocr`).

Point it at a folder already on disk instead — works with no internet at all:

```bash
python tender_extractor.py --folder "D:\Tenders\Master Tender Uploads" --no-ai
```

Flags: `--out` Excel name · `--html` dashboard name · `--no-ai` rules only ·
`--no-cache` force reprocess.

## How it works

1. **Collect** — walks the Drive folder recursively, at any nesting depth.
   Skips `WORKING FOLDER` (your own draft submissions) and temp files (`~$…`).
2. **Read** — PyMuPDF for PDF text; any page with almost no text layer is
   rasterised and OCR-ed with Tesseract. `python-docx` for DOCX, pandas for
   XLS/BOQ. Documents are read in priority order, so the GeM bid document and
   the NIB/tender-summary are consulted before annexures.
3. **Extract, twice, independently**
   - *Rules layer*: label-anchored regex plus section capture. A label is only
     trusted at full confidence when it heads a line or carries a colon, so a
     stray mention inside prose can never outrank a real field. Handles the
     fixed GeM template natively (bid number, EMD, ePBG %, estimated value,
     bid end date, contract period, buyer address).
     A number is only accepted as money when it carries a currency marker
     (`Rs.`, `INR`, `₹`), a `/-` suffix, or a multiplier word (`lakh`, `crore`).
     Bare digits pass only under a structured template label such as GeM's
     "EMD Amount", and even then years, leading-zero identifiers, numbers
     embedded in references like `T-168`, and file-name digits are rejected.
     This is what stops a bid number or an account number being reported as an
     EMD.
   - *AI layer*: Gemini free tier, instructed to cite the file and page for
     every value and to answer `NOT FOUND` rather than guess. Long tenders are
     chunked and the chunk results merged.
4. **Merge and flag** — where the two readers disagree on a money or date
   field, both readings appear in the cell and it is flagged. Nothing is ever
   silently guessed.
   If the AI layer is unreachable (Gemini 503s under load), the engine rotates
   through six models before giving up, falls back to the rules layer, and
   **does not cache that tender** — so a later re-run retries only the tenders
   that missed the AI pass.
5. **Cache** — a tender whose files have not changed is served from cache, so
   re-runs only process newly added folders.

## The output workbook

- **Tender Summary** — one row per tender, the 13 fields, plus a *Needs Review* column.
- **Evidence** — per field: source file and page, confidence, and what each of
  the two readers read. This is how you verify a value in seconds.
- **Documents Read** — which files were opened, page counts, how many needed OCR.
- **Read Me** — the colour legend.

Amber = needs your eye (readers disagreed / low confidence / came off a scan).
Red = not stated in the documents that were read.

## Honest limits

- This removes the typing, not the review. **Always confirm EMD, fees and the
  submission deadline against the source page before acting on them.**
- Scanned PDFs depend on scan quality. OCR can mangle digits, so values from
  OCR-ed pages are flagged.
- Legacy `.doc` files need LibreOffice installed; otherwise they are reported
  as skipped in the *Documents Read* sheet rather than silently ignored.
- Gemini's free tier is rate-limited. The engine backs off and retries; a very
  large batch may simply take longer.

## Tuning it

Almost all accuracy work happens in two dictionaries in `tender_extractor.py`:

- `LABELS` — add a label phrasing you see in a tender that was missed, e.g. a
  new wording for tender fees. Order matters: most specific first.
- `SECTION_HEADINGS` — add heading wordings for Scope / Eligibility / Penalty.

After editing, run both test suites, then refresh the notebook:

```bash
python test_rules.py && python test_regressions.py && python build_notebook.py
```
