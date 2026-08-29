"""
patch_accuracy.py — applies the accuracy fixes found by auditing tender
2026-9533 (NTPC Solapur ash-disposal audit) against its source PDFs.

Findings this patch addresses:
  1. SD read "Rs. 1,00,000" from a clause that says Security Deposit *up to*
     Rs 1,00,000 must be paid by EFT — a payment-mode threshold, not the
     deposit. Truth: 5.0% of total Contract Price (SCC cl.46).
  2. EMD reported NOT FOUND although three documents state "Not Applicable".
  3. Penalty captured the right clause header then ran on into two unrelated
     clauses about site visits.
  4. Location returned "Ntpc Limited" (the Organisation Name) instead of the
     site address that sits under "Location of Work" / "Site location".
  5. Purpose returned the generic GeM category "Custom Bid for Services"
     instead of "Statutory audit of ash disposal for FY25-26".
  6. Eligibility stopped at the end of its first page, dropping 85% of the
     criteria including the CPCB-authorised-auditor gate.
  7. Money fields were marked "high" confidence on a single reader when the
     AI cross-check had not run at all.

Idempotent: re-running detects the marker and does nothing.

    python patch_accuracy.py
"""

import io
import os
import sys

SRC = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   "tender_extractor.py")
MARKER = "# --- accuracy patch 2026-08-26 ---"


def main():
    s = io.open(SRC, encoding="utf-8").read()
    if MARKER in s:
        print("already patched; nothing to do")
        return 0
    before = len(s)
    edits = 0

    def sub(old, new, label):
        nonlocal s, edits
        if old not in s:
            print("!! ANCHOR MISSING: %s" % label)
            sys.exit(1)
        if s.count(old) != 1:
            print("!! ANCHOR NOT UNIQUE (%d): %s" % (s.count(old), label))
            sys.exit(1)
        s = s.replace(old, new, 1)
        edits += 1
        print("   ok  %s" % label)

    # ---------------------------------------------------------------- E1 ---
    # New parsing helpers: explicit-nil, percentage-of-basis, threshold
    # qualifiers, and unreadable-glyph detection.
    sub(
        'def parse_date(window: str) -> str:\n'
        '    m = DATE_RE.search(window)\n'
        '    return m.group(0).strip() if m else ""\n',

        'def parse_date(window: str) -> str:\n'
        '    m = DATE_RE.search(window)\n'
        '    return m.group(0).strip() if m else ""\n'
        '\n'
        '\n' + MARKER + '\n'
        '\n'
        '# "Not Applicable" against EMD is an ANSWER, not a miss. Reporting it as\n'
        '# NOT FOUND sends the reader hunting for something the tender already\n'
        '# settled. Seen in three separate documents of tender 2026-9533.\n'
        'EXPLICIT_NIL_RE = re.compile(\n'
        r'    r"\b(?:not\s*applicable|not\s*required|nil|none|n\.?a\.?|exempt(?:ed)?|'\
        '"\n'
        r'    r"no\s+emd|without\s+emd|waived)\b", re.I)'
        '\n'
        '\n'
        '# Words that turn a nearby rupee figure into a THRESHOLD or a CEILING\n'
        '# rather than the amount payable. Real culprit in 2026-9533:\n'
        '#   "Security Deposit amount up to Rs. 1,00,000/- must be submitted\n'
        '#    through Electronic Fund Transfer (EFT) only"\n'
        '# — a payment-mode rule, read as the deposit itself.\n'
        'MONEY_QUALIFIER_RE = re.compile(\n'
        r'    r"\b(?:up\s*to|upto|not\s*exceeding|exceeding|in\s*excess\s*of|"'
        '\n'
        r'    r"more\s+than|less\s+than|below|above|at\s*least|minimum\s+of|"'
        '\n'
        r'    r"maximum\s+of|whichever|threshold|per\s*day|per\s*week|per\s*month|"'
        '\n'
        r'    r"per\s*visit|per\s*branch|per\s*sitting|slab)\b", re.I)'
        '\n'
        '\n'
        '# Security deposit / performance security in these tenders is stated as a\n'
        '# percentage of contract value, never as a rupee figure. Capturing the\n'
        '# percentage is both correct and more useful than a derived amount.\n'
        'PCT_OF_RE = re.compile(\n'
        r'    r"(\d{1,2}(?:\.\d+)?)\s*(?:%|per\s*cent|percent)\s*(?:of|on)\s+"'
        '\n'
        r'    r"((?:the\s+)?(?:total\s+)?(?:awarded\s+)?"'
        '\n'
        r'    r"(?:contract|order|bid|tender|work|purchase\s*order)\s*"'
        '\n'
        r'    r"(?:price|value|amount|cost))", re.I)'
        '\n'
        '\n'
        '# PyMuPDF leaves an unmapped glyph as U+FFFD. The NTPC GCC states the\n'
        '# liquidated-damages rate with a vulgar fraction, so the rate extracts as\n'
        '# "damages@ \\ufffd percent per week". Emitting that silently would be worse\n'
        '# than saying the character could not be read.\n'
        'UNREADABLE_GLYPH = "\\ufffd"\n'
        '\n'
        '\n'
        'def parse_percent_of(window: str) -> str:\n'
        '    """Return e.g. "5.0% of the total Contract Price", or ""."""\n'
        '    m = PCT_OF_RE.search(window)\n'
        '    if not m:\n'
        '        return ""\n'
        '    basis = re.sub(r"\\s+", " ", m.group(2)).strip()\n'
        '    return "%s%% of %s" % (m.group(1), basis)\n'
        '\n'
        '\n'
        'def parse_amount_at(window: str, allow_bare: bool = False):\n'
        '    """parse_amount, but also returns where in the window the figure sat.\n'
        '\n'
        '    The offset is what lets the caller check whether a qualifier such as\n'
        '    "up to" appears between the label and the number.\n'
        '    """\n'
        '    for m in AMOUNT_RE.finditer(window):\n'
        '        val, disp = parse_amount(m.group(0), allow_bare=allow_bare)\n'
        '        if val is not None:\n'
        '            return val, disp, m.start()\n'
        '    return None, "", -1\n',
        "E1 parsing helpers")

    # ---------------------------------------------------------------- E2 ---
    # Location: the real address labels must outrank the organisation name.
    sub(
        '    "location": [\n'
        '        r"Buyer\\s*Name\\s*/?\\s*Address",\n'
        '        # Full form only. Matching bare "Consignee" hits the plural heading\n'
        '        # "Consignees/Reporting Officer and Quantity" and captures its leftovers.\n'
        '        r"Consignee\\s*/\\s*Reporting\\s*Officer\\s*/\\s*Address",\n'
        '        r"Organisation\\s*Name",\n'
        '        r"Office\\s*Name",\n'
        '        r"Address\\s*of\\s*(?:the\\s*)?(?:Office|Department|Organisation)",\n'
        '        r"Place\\s*of\\s*(?:Work|Audit|Posting)",\n'
        '    ],',

        '    "location": [\n'
        '        # Site/work address first. "Organisation Name" used to win and\n'
        '        # returned "Ntpc Limited" for a tender whose site address was\n'
        '        # spelled out three times under these labels.\n'
        '        r"Location\\s*of\\s*Work",\n'
        '        r"Site\\s*location",\n'
        '        r"Delivery\\s*Address",\n'
        '        r"Place\\s*of\\s*(?:Work|Audit|Posting)",\n'
        '        r"Address\\s*of\\s*(?:the\\s*)?(?:Office|Department|Organisation)",\n'
        '        r"Buyer\\s*Name\\s*/?\\s*Address",\n'
        '        # Full form only. Matching bare "Consignee" hits the plural heading\n'
        '        # "Consignees/Reporting Officer and Quantity" and captures its leftovers.\n'
        '        r"Consignee\\s*/\\s*Reporting\\s*Officer\\s*/\\s*Address",\n'
        '        r"\\bLocation\\b",\n'
        '        r"Delivery\\s*district",\n'
        '        # Names of the buying entity, not a place. Last resort only.\n'
        '        r"Office\\s*Name",\n'
        '        r"Organisation\\s*Name",\n'
        '    ],',
        "E2 location label order")

    # Purpose: the descriptive work name beats the generic GeM category.
    sub(
        '    "purpose": [\n'
        '        r"Item\\s*Category",\n'
        '        r"Name\\s*of\\s*(?:the\\s*)?Work",\n'
        '        r"Nature\\s*of\\s*(?:Work|Service|Assignment)",\n'
        '        r"Type\\s*of\\s*Audit",\n'
        '        r"Subject",\n'
        '    ],',

        '    "purpose": [\n'
        '        # "Item Category" on a GeM custom bid is always the useless\n'
        '        # "Custom Bid for Services". The work name carries the real answer.\n'
        '        r"Name\\s*of\\s*(?:the\\s*)?Work",\n'
        '        r"Type\\s*of\\s*Audit",\n'
        '        r"Nature\\s*of\\s*(?:Work|Service|Assignment)",\n'
        '        r"Nature\\s*of\\s*Requirement",\n'
        '        r"Products?",\n'
        '        r"Subject",\n'
        '        r"Item\\s*Category",\n'
        '    ],',
        "E2 purpose label order")

    # ---------------------------------------------------------------- E3 ---
    # Reject generic placeholder values for location / purpose so the search
    # carries on to a real one.
    sub(
        'def _line_anchored(text: str, start: int, end: int) -> bool:',

        '# Values that are technically present but say nothing. Accepting one of\n'
        '# these stops the search and hides the real answer further down the page.\n'
        'GENERIC_VALUES = {\n'
        '    "purpose": re.compile(\n'
        '        r"^(?:custom\\s*bid(?:\\s*for\\s*services?)?|services?|goods?|works?|"\n'
        '        r"n\\.?a\\.?|not\\s*applicable|others?|miscellaneous)\\s*$", re.I),\n'
        '    "location": re.compile(\n'
        '        r"^(?:n\\.?a\\.?|not\\s*applicable|india|as\\s*per\\s*.*|"\n'
        '        r"[a-z0-9 .&-]*(?:limited|ltd\\.?|corporation|corp\\.?|"\n'
        '        r"nigam|board|authority|company)\\s*)$", re.I),\n'
        '}\n'
        '\n'
        '# An address is expected to carry a PIN code or several comma-separated\n'
        '# parts. A single proper noun is a name, not a place.\n'
        'ADDRESSISH_RE = re.compile(r"\\b\\d{6}\\b|,.*,|\\bP\\.?O\\.?\\b|\\bDist\\b|\\bPIN\\b", re.I)\n'
        '\n'
        '\n'
        'def _value_is_weak(field: str, value: str) -> bool:\n'
        '    """True when a value parsed cleanly but carries no information."""\n'
        '    rx = GENERIC_VALUES.get(field)\n'
        '    if rx and rx.match(value.strip()):\n'
        '        return True\n'
        '    if field == "location" and len(value) < 60 and not ADDRESSISH_RE.search(value):\n'
        '        return True\n'
        '    return False\n'
        '\n'
        '\n'
        'def _line_anchored(text: str, start: int, end: int) -> bool:',
        "E3 weak-value guard")

    # ---------------------------------------------------------------- E4 ---
    # _label_lookup: field-aware; percentage and explicit-nil handling for
    # money; qualifier rejection; weak-value fallback for text.
    sub(
        'def _label_lookup(pages: list[Page], patterns: list[str], kind: str) -> Cand:',
        'def _label_lookup(pages: list[Page], patterns: list[str], kind: str,\n'
        '                  field: str = "") -> Cand:',
        "E4 signature")

    sub(
        '                    if kind == "money":\n'
        '                        # Bare digits only from a structured label, matched\n'
        '                        # line-anchored. In prose, currency must be present.\n'
        '                        val, disp = parse_amount(\n'
        '                            window, allow_bare=strict and pat in BARE_OK_LABELS)\n'
        '                        if val is not None:\n'
        '                            return Cand(disp, pg.ref, conf, val, m.group(0))\n'
        '                    elif kind == "date":\n'
        '                        d = parse_date(window)\n'
        '                        if d:\n'
        '                            return Cand(d, pg.ref, conf, None, m.group(0))\n'
        '                    else:\n'
        '                        v = _text_value(window)\n'
        '                        if len(v) >= 3:\n'
        '                            return Cand(v[:400], pg.ref,\n'
        '                                        "medium" if strict else "low",\n'
        '                                        None, m.group(0))\n'
        '    return Cand()',

        '                    if kind == "money":\n'
        '                        # A percentage of contract value is how security\n'
        '                        # deposits and performance guarantees are actually\n'
        '                        # specified. Prefer it over any rupee figure nearby.\n'
        '                        pct = parse_percent_of(window)\n'
        '                        if pct:\n'
        '                            return Cand(pct, pg.ref, conf, None, m.group(0))\n'
        '\n'
        '                        # Bare digits only from a structured label, matched\n'
        '                        # line-anchored. In prose, currency must be present.\n'
        '                        val, disp, at = parse_amount_at(\n'
        '                            window, allow_bare=strict and pat in BARE_OK_LABELS)\n'
        '                        if val is not None:\n'
        '                            # "...deposit up to Rs 1,00,000 must be paid by\n'
        '                            # EFT" states a threshold, not the deposit.\n'
        '                            if MONEY_QUALIFIER_RE.search(window[:at]):\n'
        '                                weak_money.append(\n'
        '                                    Cand(disp, pg.ref, "low", val, m.group(0)))\n'
        '                            else:\n'
        '                                return Cand(disp, pg.ref, conf, val, m.group(0))\n'
        '                        elif EXPLICIT_NIL_RE.search(window[:120]):\n'
        '                            # Stated as nil/not applicable: a real answer.\n'
        '                            return Cand("Not Applicable / Nil (as stated)",\n'
        '                                        pg.ref, conf, 0.0, m.group(0))\n'
        '                    elif kind == "date":\n'
        '                        d = parse_date(window)\n'
        '                        if d:\n'
        '                            return Cand(d, pg.ref, conf, None, m.group(0))\n'
        '                    else:\n'
        '                        v = _text_value(window)\n'
        '                        if len(v) >= 3:\n'
        '                            c = Cand(v[:400], pg.ref,\n'
        '                                     "medium" if strict else "low",\n'
        '                                     None, m.group(0))\n'
        '                            # Keep a placeholder value in reserve, but go on\n'
        '                            # looking for one that actually says something.\n'
        '                            if _value_is_weak(field, v):\n'
        '                                weak_text.append(c)\n'
        '                            else:\n'
        '                                return c\n'
        '    if weak_money:\n'
        '        c = weak_money[0]\n'
        '        return Cand(c.value + "   [context suggests a threshold, not the "\n'
        '                    "amount payable - verify]", c.ref, "low", c.numeric, c.raw)\n'
        '    if weak_text:\n'
        '        return weak_text[0]\n'
        '    return Cand()',
        "E4 money/text branches")

    sub(
        '    # Pass 1 trusts only labels that head a line or carry a colon - that is a\n'
        '    # real field. Pass 2 relaxes it, so a label buried in prose is a fallback,\n'
        '    # never a first choice.\n'
        '    for strict in (True, False):',

        '    # Pass 1 trusts only labels that head a line or carry a colon - that is a\n'
        '    # real field. Pass 2 relaxes it, so a label buried in prose is a fallback,\n'
        '    # never a first choice.\n'
        '    weak_money: list[Cand] = []\n'
        '    weak_text: list[Cand] = []\n'
        '    for strict in (True, False):',
        "E4 reserve lists")

    # ---------------------------------------------------------------- E5 ---
    # _capture_section: caller-tunable break threshold, and the ability to run
    # across consecutive pages of one file.
    sub(
        'def _capture_section(pages: list[Page], patterns: list[str],\n'
        '                     max_chars: int = 4000) -> Cand:',
        'def _capture_section(pages: list[Page], patterns: list[str],\n'
        '                     max_chars: int = 4000,\n'
        '                     min_before_break: int = 200,\n'
        '                     span_pages: bool = False) -> Cand:',
        "E5 capture signature")

    sub(
        '    best = Cand()\n'
        '    for pat in patterns:\n'
        '        rx = re.compile(pat, re.I)\n'
        '        for pg in _ordered_pages(pages):',
        '    best = Cand()\n'
        '    units = _file_blocks(pages) if span_pages else _ordered_pages(pages)\n'
        '    for pat in patterns:\n'
        '        rx = re.compile(pat, re.I)\n'
        '        for pg in units:',
        "E5 capture units")

    sub(
        '                joined_len = len("\\n".join(lines))\n'
        '                if (joined_len > 200 and NEXT_HEADING_RE.match(ln)\n'
        '                        and not SUBITEM_RE.match(ln)):\n'
        '                    break',
        '                joined_len = len("\\n".join(lines))\n'
        '                if (joined_len > min_before_break and NEXT_HEADING_RE.match(ln)\n'
        '                        and not SUBITEM_RE.match(ln)):\n'
        '                    break',
        "E5 break threshold")

    # A page-spanning pseudo-page, so a section that runs over a page break is
    # captured whole instead of being cut at the footer.
    sub(
        'def _ordered_pages(pages: list[Page]) -> list[Page]:\n'
        '    """Read the documents most likely to carry the facts first."""\n'
        '    return sorted(pages, key=lambda p: (-file_priority(p.file), p.file, p.page))\n',

        'def _ordered_pages(pages: list[Page]) -> list[Page]:\n'
        '    """Read the documents most likely to carry the facts first."""\n'
        '    return sorted(pages, key=lambda p: (-file_priority(p.file), p.file, p.page))\n'
        '\n'
        '\n'
        'class _Block:\n'
        '    """One file\'s pages joined into a single searchable text.\n'
        '\n'
        '    Eligibility criteria routinely run over several pages; capturing per\n'
        '    page cut tender 2026-9533\'s criteria off after the first of four,\n'
        '    losing the CPCB-authorised-auditor requirement that decides whether\n'
        '    the firm may bid at all. `ref` reports the page the match started on.\n'
        '    """\n'
        '\n'
        '    def __init__(self, pages: list[Page]):\n'
        '        pages = sorted(pages, key=lambda p: p.page)\n'
        '        self.file = pages[0].file\n'
        '        self._starts: list[tuple[int, Page]] = []\n'
        '        chunks, pos = [], 0\n'
        '        for p in pages:\n'
        '            self._starts.append((pos, p))\n'
        '            chunks.append(p.text)\n'
        '            pos += len(p.text) + 1\n'
        '        self.text = "\\n".join(chunks)\n'
        '        self._match_at = 0\n'
        '\n'
        '    @property\n'
        '    def ref(self) -> str:\n'
        '        page = self._starts[0][1]\n'
        '        for start, p in self._starts:\n'
        '            if start <= self._match_at:\n'
        '                page = p\n'
        '            else:\n'
        '                break\n'
        '        return page.ref\n'
        '\n'
        '\n'
        'def _file_blocks(pages: list[Page]) -> list[_Block]:\n'
        '    by_file: dict[str, list[Page]] = {}\n'
        '    for p in pages:\n'
        '        by_file.setdefault(p.file, []).append(p)\n'
        '    blocks = [_Block(v) for v in by_file.values()]\n'
        '    return sorted(blocks, key=lambda b: (-file_priority(b.file), b.file))\n',
        "E5 file blocks")

    # Record where the heading matched so _Block.ref names the right page.
    sub(
        '            m = next((mm for mm in rx.finditer(pg.text)\n'
        '                      if _line_anchored(pg.text, mm.start(), mm.end())), None)\n'
        '            if not m:\n'
        '                continue',
        '            m = next((mm for mm in rx.finditer(pg.text)\n'
        '                      if _line_anchored(pg.text, mm.start(), mm.end())), None)\n'
        '            if not m:\n'
        '                continue\n'
        '            if isinstance(pg, _Block):\n'
        '                pg._match_at = m.start()',
        "E5 block ref tracking")

    # ---------------------------------------------------------------- E6 ---
    # Penalty: pull the actual rate sentence rather than a clause header plus
    # whatever followed it.
    sub(
        'GEM_BID_NO_RE = re.compile(r"(GEM/\\d{4}/[BR]/\\d+)", re.I)',

        '# The sentence that actually states the liquidated-damages rate, e.g.\n'
        '#   "liable for payment of liquidated damages@ 1/2 percent per week, not as\n'
        '#    penalty, on the Contract Value ... subject to a maximum of 5% of the\n'
        '#    Contract Value"\n'
        '# Capturing this beats capturing the clause heading, which in 2026-9533\n'
        '# said only "Liquidated Damages / As per GCC" and then bled into the next\n'
        '# two clauses about site inspection.\n'
        'LD_RATE_RE = re.compile(\n'
        r'    r"(?:liquidated\s+damages?|penalt(?:y|ies))\b[^.]{0,400}?"'
        '\n'
        r'    r"(?:\d{1,2}(?:\.\d+)?\s*(?:%|per\s*cent|percent)|' + "\\ufffd" + r'|\u00bd|\u00bc|\u00be)"'
        '\n'
        r'    r"[^.]{0,400}?\.", re.I | re.S)'
        '\n'
        '\n'
        'LD_CAP_RE = re.compile(\n'
        r'    r"maximum\s+of\s+(\d{1,2}(?:\.\d+)?)\s*(?:%|per\s*cent|percent)\s*"'
        '\n'
        r'    r"of\s+the\s+([A-Za-z ]{0,40}?(?:contract|order)\s*(?:value|price))", re.I)'
        '\n'
        '\n'
        '\n'
        'def find_penalty_rate(pages: list[Page]) -> Cand:\n'
        '    """The liquidated-damages rate sentence, if any document states one."""\n'
        '    for pg in _ordered_pages(pages):\n'
        '        m = LD_RATE_RE.search(pg.text)\n'
        '        if not m:\n'
        '            continue\n'
        '        body = _clean(m.group(0)).strip()\n'
        '        if len(body) < 40:\n'
        '            continue\n'
        '        cap = LD_CAP_RE.search(pg.text[m.start(): m.start() + 1200])\n'
        '        if cap and "maximum" not in body.lower():\n'
        '            body += "  (capped at %s%% of the %s)" % (cap.group(1), cap.group(2))\n'
        '        if UNREADABLE_GLYPH in body:\n'
        '            # The rate is a vulgar fraction the PDF font does not map.\n'
        '            # Say so rather than shipping a mystery character.\n'
        '            body = body.replace(UNREADABLE_GLYPH, "[?]")\n'
        '            return Cand(body + "   [the rate character did not extract - "\n'
        '                        "read it off the source page]", pg.ref, "low")\n'
        '        return Cand(body, pg.ref, "high")\n'
        '    return Cand()\n'
        '\n'
        '\n'
        'GEM_BID_NO_RE = re.compile(r"(GEM/\\d{4}/[BR]/\\d+)", re.I)',
        "E6 penalty rate extractor")

    # ---------------------------------------------------------------- E7 ---
    # Wire the new behaviour into rules_extract.
    sub(
        '    for f in MONEY_FIELDS:\n'
        '        out[f] = _label_lookup(pages, LABELS.get(f, []), "money")\n'
        '    for f in DATE_FIELDS:\n'
        '        out[f] = _label_lookup(pages, LABELS.get(f, []), "date")\n'
        '    for f in ("period", "location", "purpose", "tender_name"):\n'
        '        out[f] = _label_lookup(pages, LABELS.get(f, []), "text")\n'
        '    for f, pats in SECTION_HEADINGS.items():\n'
        '        out[f] = _capture_section(pages, pats)',

        '    for f in MONEY_FIELDS:\n'
        '        out[f] = _label_lookup(pages, LABELS.get(f, []), "money", f)\n'
        '    for f in DATE_FIELDS:\n'
        '        out[f] = _label_lookup(pages, LABELS.get(f, []), "date", f)\n'
        '    for f in ("period", "location", "purpose", "tender_name"):\n'
        '        out[f] = _label_lookup(pages, LABELS.get(f, []), "text", f)\n'
        '    for f, pats in SECTION_HEADINGS.items():\n'
        '        if f == "eligibility":\n'
        '            # Criteria span pages; a per-page capture loses most of them.\n'
        '            out[f] = _capture_section(pages, pats, max_chars=14000,\n'
        '                                      span_pages=True)\n'
        '        elif f == "penalty":\n'
        '            # Break at the very next numbered clause: the SCC lists\n'
        '            # penalties as one-line table rows, so a 200-char run-on\n'
        '            # swallows the clauses that follow.\n'
        '            out[f] = _capture_section(pages, pats, min_before_break=40)\n'
        '        else:\n'
        '            out[f] = _capture_section(pages, pats)\n'
        '\n'
        '    # A stated rate beats a clause heading. "Liquidated Damages / As per\n'
        '    # GCC" is a pointer; the reader needs the number it points at.\n'
        '    rate = find_penalty_rate(pages)\n'
        '    if rate.value:\n'
        '        head = out.get("penalty", Cand()).value\n'
        '        if head and head[:60] not in rate.value:\n'
        '            rate = Cand(rate.value + "\\n\\n[clause reference: "\n'
        '                        + _clean(head)[:200] + "]", rate.ref, rate.conf)\n'
        '        out["penalty"] = rate',
        "E7 rules_extract wiring")

    # ---------------------------------------------------------------- E8 ---
    # Confidence must reflect how many readers actually looked.
    sub(
        'def merge(rules: dict[str, Cand], ai: dict[str, Cand]) -> dict[str, Result]:',
        'def merge(rules: dict[str, Cand], ai: dict[str, Cand],\n'
        '          ai_ran: bool = True) -> dict[str, Result]:',
        "E8 merge signature")

    sub(
        '        if not res.flag and res.conf == "low":\n'
        '            res.flag = "low confidence"\n'
        '        if not res.flag and any(w in (res.ref or "").upper() for w in ("OCR",)):\n'
        '            res.flag = "from OCR - verify digits"\n'
        '        out[key] = res\n'
        '    return out',

        '        # One reader is not a cross-check. Reporting "high" confidence on a\n'
        '        # money field that only the rules layer saw is what let a wrong\n'
        '        # security deposit through unflagged in tender 2026-9533.\n'
        '        if not ai_ran and key in (MONEY_FIELDS | DATE_FIELDS):\n'
        '            if res.conf == "high":\n'
        '                res.conf = "medium"\n'
        '            if not res.flag:\n'
        '                res.flag = "single reader - AI cross-check did not run"\n'
        '\n'
        '        if not res.flag and res.conf == "low":\n'
        '            res.flag = "low confidence"\n'
        '        if not res.flag and any(w in (res.ref or "").upper() for w in ("OCR",)):\n'
        '            res.flag = "from OCR - verify digits"\n'
        '        out[key] = res\n'
        '    return out',
        "E8 single-reader calibration")

    sub(
        '    rules = rules_extract(pages, name)\n'
        '    ai, ai_ok = ai_extract(pages, client, models)\n'
        '    results = merge(rules, ai)',
        '    rules = rules_extract(pages, name)\n'
        '    ai, ai_ok = ai_extract(pages, client, models)\n'
        '    results = merge(rules, ai, ai_ran=bool(client) and ai_ok)',
        "E8 merge call")

    # ---------------------------------------------------------------- E9 ---
    sub(
        '    ("Important", "This tool eliminates the typing, not the review. Always "\n'
        '                  "confirm EMD, fees, and the submission deadline against the "\n'
        '                  "source page before acting on them."),',
        '    ("Important", "This tool eliminates the typing, not the review. Always "\n'
        '                  "confirm EMD, fees, and the submission deadline against the "\n'
        '                  "source page before acting on them."),\n'
        '    ("Single reader", "Without a Gemini API key only the rules layer runs. "\n'
        '                      "Money and date fields are then capped at medium "\n'
        '                      "confidence and flagged, because nothing "\n'
        '                      "cross-checked them."),\n'
        '    ("Percentages", "Security deposit and performance guarantees are "\n'
        '                    "normally a percentage of contract value, so that is "\n'
        '                    "what is reported. A rupee figure sitting next to a "\n'
        '                    "phrase like \\"up to\\" is a threshold, not the amount "\n'
        '                    "payable, and is flagged as such."),',
        "E9 legend")

    # Cache must not serve results from the old, weaker logic.
    sub('CACHE_VERSION = "v2"', 'CACHE_VERSION = "v3"', "E10 cache bump")

    io.open(SRC, "w", encoding="utf-8", newline="\n").write(s)
    print("\n%d edits applied; %d -> %d bytes" % (edits, before, len(s)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
