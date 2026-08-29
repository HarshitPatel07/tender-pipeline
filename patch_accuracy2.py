"""
patch_accuracy2.py — second round, fixing what the first round exposed.

  1. Purpose came back as "Date". Putting "Name of Work" first made it match a
     blank proforma line in TENDER.pdf p.389 ("Name of Work: ......  Date"),
     and the wrap-follower picked up the next label. A value needs substance.
  2. Penalty found "penalty upto 10% of the contract value" on TENDER.pdf p.42
     — a banning/non-performance clause — instead of the operative liquidated
     damages rate in GCC 25.5.1 on p.90. First match won; best match should.
  3. Eligibility still stopped at page 3. The GeM criteria table numbers its
     sub-questions ("4. Documents required to prove turnover:"), and those look
     exactly like the next clause heading, so the capture broke on its own
     content.
  4. Penalty's clause reference still carried the run-on text, because the
     40-char floor was never reached before the stray clause arrived.

    python patch_accuracy2.py
"""

import io
import os
import sys

SRC = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   "tender_extractor.py")
MARKER = "# --- accuracy patch round 2 ---"


def main():
    s = io.open(SRC, encoding="utf-8").read()
    if MARKER in s:
        print("already patched; nothing to do")
        return 0
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

    # ---------------------------------------------------------------- F1 ---
    # A value must say something. Blank proforma lines and bare form labels
    # are not answers.
    sub(
        '# An address is expected to carry a PIN code or several comma-separated\n'
        '# parts. A single proper noun is a name, not a place.\n'
        'ADDRESSISH_RE = re.compile(r"\\b\\d{6}\\b|,.*,|\\bP\\.?O\\.?\\b|\\bDist\\b|\\bPIN\\b", re.I)',

        MARKER + '\n'
        '# Words that are the NEXT label on a blank proforma, not a value. Tender\n'
        '# documents are full of "Name of Work: ..........  Date: .........."; the\n'
        '# dotted blank is skipped as junk and the following label gets picked up.\n'
        'FORM_LABEL_RE = re.compile(\n'
        '    r"^(?:date|dated|signature|sign|name|address|place|seal|stamp|"\n'
        '    r"designation|witness|to|from|ref|reference|subject|sub|page|"\n'
        '    r"sr\\.?\\s*no\\.?|s\\.?\\s*no\\.?|annexure|appendix|amount|total|nil)"\n'
        '    r"\\s*:?\\s*$", re.I)\n'
        '\n'
        '# An address is expected to carry a PIN code or several comma-separated\n'
        '# parts. A single proper noun is a name, not a place.\n'
        'ADDRESSISH_RE = re.compile(r"\\b\\d{6}\\b|,.*,|\\bP\\.?O\\.?\\b|\\bDist\\b|\\bPIN\\b", re.I)',
        "F1 form-label regex")

    sub(
        'def _value_is_weak(field: str, value: str) -> bool:\n'
        '    """True when a value parsed cleanly but carries no information."""\n'
        '    rx = GENERIC_VALUES.get(field)\n'
        '    if rx and rx.match(value.strip()):\n'
        '        return True\n'
        '    if field == "location" and len(value) < 60 and not ADDRESSISH_RE.search(value):\n'
        '        return True\n'
        '    return False',

        'def _value_is_weak(field: str, value: str) -> bool:\n'
        '    """True when a value parsed cleanly but carries no information."""\n'
        '    v = value.strip()\n'
        '    rx = GENERIC_VALUES.get(field)\n'
        '    if rx and rx.match(v):\n'
        '        return True\n'
        '    if field in ("purpose", "location") and FORM_LABEL_RE.match(v):\n'
        '        return True\n'
        '    if field == "purpose" and len(re.findall(r"[A-Za-z]{2,}", v)) < 3:\n'
        '        # "Date", "Audit", "Services" - a label or a bare category.\n'
        '        return True\n'
        '    if field == "location" and len(v) < 60 and not ADDRESSISH_RE.search(v):\n'
        '        return True\n'
        '    return False',
        "F1 substance guard")

    # ---------------------------------------------------------------- F2 ---
    # Penalty: score every candidate rate sentence and take the best, rather
    # than the first one encountered in page order.
    sub(
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
        '    return Cand()',

        'RATE_PER_RE = re.compile(r"per\\s*(?:week|day|month|fortnight)", re.I)\n'
        '\n'
        '\n'
        'def _score_penalty(body: str, has_cap: bool) -> int:\n'
        '    """Rank rate sentences. The operative clause states a rate AND a cap.\n'
        '\n'
        '    Needed because a tender mentions penalties in several places -\n'
        '    banning policy, non-performance, fraud - and only one of them is the\n'
        '    delay penalty the bidder is exposed to. In 2026-9533 the first match\n'
        '    in page order was a 10% banning-policy clause on p.42, while the\n'
        '    operative liquidated damages sat in GCC 25.5.1 on p.90.\n'
        '    """\n'
        '    low = body.lower()\n'
        '    score = 0\n'
        '    if "liquidated damage" in low:\n'
        '        score += 4\n'
        '    if RATE_PER_RE.search(low):\n'
        '        score += 3\n'
        '    if has_cap or "maximum of" in low:\n'
        '        score += 2\n'
        '    if "contract value" in low or "contract price" in low:\n'
        '        score += 1\n'
        '    # A clause about forfeiting or banning is not the delay penalty.\n'
        '    if any(w in low for w in ("banning", "blacklist", "debar", "forfeit")):\n'
        '        score -= 3\n'
        '    return score\n'
        '\n'
        '\n'
        'def find_penalty_rate(pages: list[Page]) -> Cand:\n'
        '    """The best liquidated-damages / penalty rate sentence in the tender."""\n'
        '    best, best_score = Cand(), 0\n'
        '    for pg in _ordered_pages(pages):\n'
        '        for m in LD_RATE_RE.finditer(pg.text):\n'
        '            body = _clean(m.group(0)).strip()\n'
        '            if len(body) < 40:\n'
        '                continue\n'
        '            cap = LD_CAP_RE.search(pg.text[m.start(): m.start() + 1200])\n'
        '            score = _score_penalty(body, bool(cap))\n'
        '            if score <= 0 or score <= best_score:\n'
        '                continue\n'
        '            if cap and "maximum" not in body.lower():\n'
        '                body += ("  (capped at %s%% of the %s)"\n'
        '                         % (cap.group(1), cap.group(2)))\n'
        '            if UNREADABLE_GLYPH in body:\n'
        '                # The rate is a vulgar fraction the PDF font does not map.\n'
        '                # Say so rather than shipping a mystery character.\n'
        '                best = Cand(body.replace(UNREADABLE_GLYPH, "[?]")\n'
        '                            + "   [the rate character did not extract - "\n'
        '                            "read it off the source page]", pg.ref, "low")\n'
        '            else:\n'
        '                best = Cand(body, pg.ref, "high")\n'
        '            best_score = score\n'
        '    return best',
        "F2 penalty scoring")

    # ---------------------------------------------------------------- F3 ---
    # Let the caller supply the "next heading" test, so a numbered list that IS
    # the content does not terminate the capture.
    sub(
        'def _capture_section(pages: list[Page], patterns: list[str],\n'
        '                     max_chars: int = 4000,\n'
        '                     min_before_break: int = 200,\n'
        '                     span_pages: bool = False) -> Cand:',
        'def _capture_section(pages: list[Page], patterns: list[str],\n'
        '                     max_chars: int = 4000,\n'
        '                     min_before_break: int = 200,\n'
        '                     span_pages: bool = False,\n'
        '                     break_re: "re.Pattern | None" = None) -> Cand:',
        "F3 capture signature")

    sub(
        '                joined_len = len("\\n".join(lines))\n'
        '                if (joined_len > min_before_break and NEXT_HEADING_RE.match(ln)\n'
        '                        and not SUBITEM_RE.match(ln)):\n'
        '                    break',
        '                joined_len = len("\\n".join(lines))\n'
        '                stop = break_re or NEXT_HEADING_RE\n'
        '                if (joined_len > min_before_break and stop.match(ln)\n'
        '                        and not SUBITEM_RE.match(ln)):\n'
        '                    break',
        "F3 pluggable break")

    sub(
        'SUBITEM_RE = re.compile(r"^\\s*\\(?[a-z0-9ivx]{1,3}[).]")',
        'SUBITEM_RE = re.compile(r"^\\s*\\(?[a-z0-9ivx]{1,3}[).]")\n'
        '\n'
        '# For eligibility only: the GeM criteria table numbers its own\n'
        '# sub-questions, so a numbered line is content, not the next heading.\n'
        '# Stop only at a genuine all-caps heading or a document division.\n'
        'ELIG_BREAK_RE = re.compile(\n'
        '    r"^\\s*(?:[A-Z][A-Z0-9 ,\\-&/()\'.]{14,}\\s*:?\\s*$"\n'
        '    r"|(?:ANNEXURE|APPENDIX|SECTION|CHAPTER|PART|CHECKLIST|'
        'DECLARATION)\\b)")',
        "F3 eligibility break regex")

    # ---------------------------------------------------------------- F4 ---
    sub(
        '        if f == "eligibility":\n'
        '            # Criteria span pages; a per-page capture loses most of them.\n'
        '            out[f] = _capture_section(pages, pats, max_chars=14000,\n'
        '                                      span_pages=True)\n'
        '        elif f == "penalty":\n'
        '            # Break at the very next numbered clause: the SCC lists\n'
        '            # penalties as one-line table rows, so a 200-char run-on\n'
        '            # swallows the clauses that follow.\n'
        '            out[f] = _capture_section(pages, pats, min_before_break=40)',

        '        if f == "eligibility":\n'
        '            # Criteria span pages; a per-page capture loses most of them.\n'
        '            out[f] = _capture_section(pages, pats, max_chars=14000,\n'
        '                                      span_pages=True,\n'
        '                                      break_re=ELIG_BREAK_RE)\n'
        '        elif f == "penalty":\n'
        '            # Break at the very next numbered clause: the SCC lists\n'
        '            # penalties as one-line table rows, so any run-on swallows\n'
        '            # the unrelated clauses that follow.\n'
        '            out[f] = _capture_section(pages, pats, min_before_break=0)',
        "F4 eligibility + penalty wiring")

    io.open(SRC, "w", encoding="utf-8", newline="\n").write(s)
    print("\n%d edits applied" % edits)
    return 0


if __name__ == "__main__":
    sys.exit(main())
