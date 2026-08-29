"""Regression tests built from the real false positives in the 26-Aug run.

Every case here is something the tool actually got wrong on Tushar's tenders.
Run this after any change to the rules layer.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import tender_extractor as te

FAILS = []


def check(name, got, want, note=""):
    ok = got == want
    FAILS.append(name) if not ok else None
    print(f"{'PASS' if ok else 'FAIL'}  {name}")
    if not ok:
        print(f"      got    : {got!r}")
        print(f"      wanted : {want!r}   {note}")


# ---------------------------------------------------------------- amounts ----
print("\n--- parse_amount: bare numbers must be rejected in prose ---")
check("bid number as EMD (was Rs. 9,722,687)",
      te.parse_amount("9722687")[0], None,
      "a bare 7-digit number in prose is not money")
check("year as fee (was Rs. 1,976)",
      te.parse_amount("1976, the Act provides", allow_bare=True)[0], None,
      "1900-2100 bare = a year")
check("year 2028 as SD (was Rs. 2,028)",
      te.parse_amount("2028 onwards", allow_bare=True)[0], None)
check("account number (was Rs. 404,214,000,030)",
      te.parse_amount("0404214000030,", allow_bare=True)[0], None,
      "leading zero = identifier")
check("tender number T-168 (was Rs. 168)",
      te.parse_amount("T-168 Estt. for replacement", allow_bare=True)[0], None,
      "embedded in an identifier")
check("file name digits",
      te.parse_amount("GeM-Bidding-9722687.pdf", allow_bare=True)[0], None)
check("percentage not money",
      te.parse_amount("5.00 % of contract value")[0], None)

print("\n--- parse_amount: real amounts must still be read ---")
check("Rs. with /-", te.parse_amount("Rs. 5,500/-")[0], 5500.0)
check("Indian grouping", te.parse_amount("Rs. 2,50,000/-")[0], 250000.0)
check("lakh words", te.parse_amount("Rs. 4.50 Lakh")[0], 450000.0)
check("crore words", te.parse_amount("Rs. 6 Crores")[0], 6e7)
check("GeM bare estimated value", te.parse_amount("1500000", allow_bare=True)[0],
      1500000.0)
check("GeM bare EMD", te.parse_amount("30000", allow_bare=True)[0], 30000.0)
check("bare with decimals", te.parse_amount("145600.0", allow_bare=True)[0],
      145600.0)

print("\n--- _num: multiplier must only apply next to the digits ---")
check("words-in-brackets must not scale (was the false 'readers disagree')",
      te._num("Rs. 2,50,000/- (Rupees Two Lakhs fifty thousand only)"),
      250000.0)
check("real lakh still scales", te._num("Rs. 4.50 Lakh"), 450000.0)

# --------------------------------------------------------------- sections ----
MIDSENTENCE = """The bidder shall comply with the terms related to suspension of
seller / service provider, where such suspension period has already expired and
the scope of work defined in clause III of Section B shall apply throughout.
"""
GENUINE = """2. SCOPE OF WORK
(a) Audit of annual accounts including Balance Sheet and Income & Expenditure.
(b) Verification of vouchers, cash book, bank book and ledgers.
(c) Reconciliation of bank accounts and the fixed asset register.
(d) Verification of TDS, GST and EPF deductions and their timely deposit.
"""

print("\n--- section capture ---")
mid = te._capture_section([te.Page("x.pdf", 1, MIDSENTENCE)],
                          te.SECTION_HEADINGS["scope_of_work"])
check("mid-sentence 'scope of work' reference is not a section",
      mid.value, "", "heading must start its own line")
real = te._capture_section([te.Page("x.pdf", 1, GENUINE)],
                           te.SECTION_HEADINGS["scope_of_work"])
check("genuine heading is still captured",
      real.value.startswith("(a) Audit of annual accounts"), True)

# ---------------------------------------------------------------- labels -----
CONSIGNEE = """Consignees/Reporting Officer and Quantity
1 Deputy Commissioner (Finance) 1
"""
print("\n--- label fragments ---")
loc = te._label_lookup([te.Page("g.pdf", 4, CONSIGNEE)],
                       te.LABELS["location"], "text")
check("plural heading must not yield 's/Reporting Officer and Quantity'",
      loc.value, "")

# ---------------------------------------------- real-document regressions ----
# Found on Tushar's actual tenders on 24-Aug, not the synthetic samples above.
print("\n--- real-document false positives (26-Aug and 24-Aug runs) ---")

BILINGUAL = "N/a काया%लय का नाम/Office Name"
check("bilingual Hindi echo of an unanswered field is not a location",
      te._is_junk_line(BILINGUAL), True)

PLACEHOLDER = "….............................., dated: …................ for Providing"
check("unfilled dot-leader template line is not a tender name",
      te._is_junk_line(PLACEHOLDER), True)

check("a genuine short value is not junk",
      te._is_junk_line("Commissioner, Nagar Nigam Jaipur Greater"), False)

# NTPC's GeM summary PDF dumps one table cell per line with no real
# structure. "EMD BG ... to be uploaded on GeM Portal" describes a submission
# format, not an amount - the real risk is the old 320-char window drifting
# past it into "51 Lakh" from a completely unrelated turnover requirement
# further down the same page.
FRAGMENTED_TABLE = (
    "EMD BG (Bank Guarantee) in a separate envelope or any other acceptable "
    "form of EMD BG,\ncopy of EMD BG to be uploaded on GeM Portal\nMandatory\n"
    "Sample\nSubmission\nNo\nOrganization\nTender ID\n9900330303\nScope\n"
    "Classification\nServices\nEligibility\nCriteria\nEligibility Criteria\n"
    "Sub-Question\nAnswer\n1. Minimum\naverage annual\nturnover\nrequired:\n"
    "51 Lakh"
)
emd = te._label_lookup([te.Page("Tender-Summary.pdf", 3, FRAGMENTED_TABLE)],
                       te.LABELS["emd"], "money")
check("EMD must not reach past its own paragraph into an unrelated turnover figure",
      emd.value, "")

# ------------------------------------------------------------- GeM tender fee
GEM = """Bid Number: GEM/2026/B/7888499
EMD Amount 30000
Estimated Bid Value 1500000
"""
print("\n--- GeM defaults ---")
r = te.rules_extract([te.Page("GeM-Bidding-1.pdf", 1, GEM)], "001")
check("GeM tender fee reported as Nil, not blank",
      r["tender_fees"].value.startswith("Nil"), True)
check("GeM EMD still read", r["emd"].numeric, 30000.0)
check("GeM estimated value still read", r["estimated_cost"].numeric, 1500000.0)

print("\n" + "=" * 60)
if FAILS:
    print(f"{len(FAILS)} FAILURE(S): " + ", ".join(FAILS))
    sys.exit(1)
print("ALL REGRESSION TESTS PASS")
