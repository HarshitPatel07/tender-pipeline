"""Offline self-test of the rules layer + Excel writer (no PDFs, no API key)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import tender_extractor as te

# ---- Case 1: a GeM bid document, as the text layer actually comes out -------
GEM = """
Bid Document
Bid Details/
Bid Number/  : GEM/2026/B/9722687
Dated/  : 07-08-2026
Bid Description : Internal Audit Services
Bid End Date/Time 21-08-2026 15:00:00
Bid Opening Date/Time 21-08-2026 15:30:00
Bid Offer Validity (From End Date) 120 (Days)
Ministry/State Name Government Of Rajasthan
Department Name Local Self Government Department
Organisation Name Nagar Nigam Jaipur Greater
Office Name Zone Office Jaipur
Item Category Internal Audit Services - Concurrent Audit
Contract Period 12 Month(s)
Minimum Average Annual Turnover of the bidder (For 3 Years) 25 Lakh (s)
Years of Past Experience Required 3 Year (s)
MSE Exemption for Years of Experience and Turnover No
Estimated Bid Value 1500000
EMD Amount 30000
ePBG Percentage(%) 3.00
Buyer Name / Address Commissioner, Nagar Nigam Jaipur Greater,
Lal Kothi, Jaipur, Rajasthan - 302015
Consignee / Reporting Officer / Address Deputy Commissioner (Finance), Jaipur
Tender Document Fee Rs. 1,000/-
"""

# ---- Case 2: a state e-procurement NIB, freeform prose ---------------------
NIB = """
NOTICE INVITING BID
NIB No. 2026-9520/Finance/CA-Audit                      Dated 05.08.2026

Name of Work: Appointment of Chartered Accountant Firm for Statutory Audit of
Accounts of the Municipal Council, Bhilwara for the Financial Year 2025-26.

Estimated Cost of Work: Rs. 4.50 Lakh
Earnest Money Deposit: Rs. 9,000/-
Cost of Tender Document: Rs. 500/-
Last Date and Time for Online Submission of Bid: 26/08/2026 upto 18:00 Hrs
Place of Work: Municipal Council, Bhilwara, Rajasthan

1. ELIGIBILITY CRITERIA
(a) The firm must be registered with the Institute of Chartered Accountants of
India and must hold a valid firm registration number.
(b) The firm should have at least 3 partners, out of which at least 1 shall be
a Fellow Chartered Accountant (FCA).
(c) The firm must have minimum 5 years of experience in statutory audit of
urban local bodies or government departments.
(d) Average annual turnover of the firm shall not be less than Rs. 10.00 Lakh
during the last three financial years.
(e) The firm should not have been blacklisted by any Government Department.

2. SCOPE OF WORK
(a) Audit of the annual accounts including Balance Sheet, Income and
Expenditure Statement and Receipts and Payments Account.
(b) Verification of all vouchers, cash book, bank book and ledgers.
(c) Reconciliation of bank accounts and verification of fixed asset register.
(d) Verification of statutory deductions including TDS, GST and EPF and their
timely deposit.
(e) Submission of audit report along with management letter within 45 days of
award of work.

3. SECURITY DEPOSIT
The successful bidder shall deposit a Security Deposit of Rs. 22,500/- being 5%
of the contract value, valid for the entire contract period.

4. PENALTY
In case of delay in submission of the audit report beyond the stipulated period
of 45 days, a penalty of Rs. 1,000 per day of delay shall be levied, subject to
a maximum of 10% of the assignment fees. Repeated deficiency in service shall
render the firm liable for termination and forfeiture of Security Deposit.

5. ASSIGNMENT FEES
The audit fee payable for the assignment shall be Rs. 3,75,000 (Rupees Three
Lakh Seventy Five Thousand only) inclusive of all out-of-pocket expenses and
exclusive of GST. The period of contract shall be one year from the date of
award, extendable by one more year on satisfactory performance.

ANNEXURE - I
Format of undertaking to be submitted by the bidder.
"""


def show(title, pages, folder):
    print("=" * 72)
    print(title)
    print("=" * 72)
    rules = te.rules_extract(pages, folder)
    merged = te.merge(rules, {})
    for key, label in te.FIELDS:
        r = merged[key]
        val = r.value.replace("\n", " / ")
        val = (val[:150] + " ...") if len(val) > 150 else val
        mark = "  <-- " + r.flag if r.flag else ""
        print(f"{label:34s} {val}{mark}")
        if r.ref:
            print(f"{'':34s} src: {r.ref}")
    print()
    return {"tender": folder, "results": merged,
            "files": [{"name": p.file, "pages": 1, "ocr": False, "note": ""}
                      for p in pages]}


rows = []
rows.append(show("CASE 1 - GeM bid document",
                 [te.Page("GeM-Bidding-9722687.pdf", 1, GEM)], "001_2026-9445"))
rows.append(show("CASE 2 - State NIB (freeform)",
                 [te.Page("NIB-Document 1786942307.pdf", 1, NIB)], "004_2026-9520"))

out = te.write_excel(rows, Path(__file__).parent / "_test_output.xlsx")
print(f"Excel written: {out}  ({out.stat().st_size:,} bytes)")
