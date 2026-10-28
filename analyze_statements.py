"""
Fast Eddies Financial Spending and Income Analysis - analyzer.

Reads redacted statements (PDF) and optional spend lists (Excel/CSV) and builds one Excel workbook
(income vs expenses by month, spend by category and card, reconciliation, review list) plus an
advisor-ready two-tab summary.

Input : PDFs with 'redacted' in the name (credit card statements - Citi, Apple Card and Synchrony store card layouts - and credit-union
        statements, detected by content), and any .xlsx/.csv spend list, anywhere under <BASE>\\Redacted
        (subfolders are included). A spend list needs columns for date, payee/provider, amount and
        (optionally) category; total rows are used to check it. A changed spend list replaces its earlier version.
Output: <BASE>\\Analysis\\Income_vs_Expense_<YEAR>_<FirstMonth>-<LastMonth>.xlsx
        (Credit_Card_Analysis_<YEAR>_... when the year has no bank statements) and
        <Advisor>_Summary_<YEAR>_... .xlsx (values only, safe to send).
Ledger: <BASE>\\Analysis\\processed_ledger.json
        Remembers every statement already processed (by file hash AND by card +
        statement end date) so nothing is ever counted twice. Delete the ledger to start over.
BASE is C:\\CCardStatements unless the environment variable FEFA_BASE says otherwise.

Personal rules (your own vendors, payees, income sources, names) belong in personal_config.py next to this
script. That file is ignored by git. See README_FIRST.md and personal_config.example.py.

Bank statements: every row of every sub-account is classified as Income (by source), Expense (by category) or
Transfer (card payments, money sent to the brokerage, moves between the credit union's own accounts). A payment
to a card becomes a transfer only when that card's own statement shows the same payment. Each month must roll
from the statement's opening balance to its closing balance and match the credit union's own totals.

How the numbers are built
  * Every statement is read in full and checked against its own Account Summary
    (previous balance, payments, credits, purchases, fees, interest, new balance).
    The Reconciliation tab shows the result for each statement.
  * A statement counts in the month it closes, so each month ties to the bank's own figures.
  * If anything on a statement cannot be read, the gap is shown as its own line
    ("UNRECONCILED - REVIEW") instead of being silently dropped.
"""
import csv
import re
import json
import hashlib
import os
import subprocess
import sys
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path


def _need(module, pip_name):
    try:
        __import__(module)
    except ImportError:
        print(f"Installing {pip_name} (one time only)...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pip_name])


_need("openpyxl", "openpyxl")
try:
    import pymupdf as fitz
except ImportError:
    _need("pymupdf", "pymupdf")
    import pymupdf as fitz
import openpyxl
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# Personal rules live in personal_config.py (next to this script, never committed). Every name in it is optional.
sys.path.insert(0, str(Path(__file__).resolve().parent))
try:
    import personal_config as PC
except ImportError:
    PC = None


def _pc(name, default):
    return getattr(PC, name, default) if PC is not None else default


# ----------------------------------------------------------------------------
# Settings
# ----------------------------------------------------------------------------
BASE = Path(os.environ.get("FEFA_BASE") or r"C:\CCardStatements")
REDACTED_DIR = BASE / "Redacted"
ANALYSIS_DIR = BASE / "Analysis"
LEDGER_PATH = ANALYSIS_DIR / "processed_ledger.json"
LEDGER_VERSION = 5

TOP_N = 25  # rows shown in each Pareto table; the rest roll into "All other"

# Vendors involved in security breaches or fraud. Spend with them is flagged in the
# Transactions tab and listed on every summary tab. Set BREACH_VENDORS in personal_config.py.
# Format: (regex, note), for example (r"\\bACME CORP\\b", "Breach disclosed 2025-03").
BREACH_VENDORS = list(_pc("BREACH_VENDORS", []))

# Merge cards whose account number changed. Format: {"Citi Premier 1111": "Citi Premier 2222"}
CARD_ALIASES = dict(_pc("CARD_ALIASES", {}))

# A statement that closes in one month but should count in another (for example a store-card cycle that runs
# 12/10 to 01/09 and is wanted in December). Format: {"Amazon Store Card 1234|2026-01-09": "2025-12"}
STATEMENT_MONTH_OVERRIDES = dict(_pc("STATEMENT_MONTH_OVERRIDES", {}))

# Account name for Excel/CSV spend lists, chosen by a pattern in the file name.
# Anything not matched is named after the file. Format: [(regex on file name, account name)]
ACCOUNT_BY_FILE = list(_pc("ACCOUNT_BY_FILE", []))
# Spend lists that are only a cross-check (not counted). The bank statements are the source.
REFERENCE_ACCOUNTS = set(_pc("REFERENCE_ACCOUNTS", set()))
# Money the brokerage reports as withdrawn each year, compared with what the bank shows as received.
# These figures are NOT kept in this script: put them in <BASE>\\Analysis\\brokerage_yearend.json
# (ignored by git), for example {"2025": [["Transfer withdrawals", 66600], ["IRA rollover withdrawals", 43500]]}
BROKER_YEAREND = {}
CHECK_LABELS = {}   # check number -> (payee, category), from check_payees.json
VENDOR_OVERRIDES = []   # (regex on the statement description, vendor, category), from vendor_overrides.json


def load_vendor_overrides():
    """Card charges that sit under the wrong merchant live in vendor_overrides.json next to the ledger:
    {"EXAMPLE BILLED NAME": ["What it really was", "Home Improvement"], ...}. The key is a regex on the description."""
    VENDOR_OVERRIDES.clear()
    path = ANALYSIS_DIR / "vendor_overrides.json"
    if not path.exists():
        return
    try:
        for pat, item in json.loads(path.read_text(encoding="utf-8")).items():
            VENDOR_OVERRIDES.append((re.compile(pat, re.I), str(item[0]).strip(), str(item[1]).strip()))
    except (ValueError, OSError, TypeError, IndexError, AttributeError, re.error) as exc:
        print(f"Could not read {path.name} ({exc}); no vendor overrides applied.")


def load_check_labels():
    """Check payees live in check_payees.json next to the ledger: {"3312": ["Payee", "Category"], ...}."""
    path = ANALYSIS_DIR / "check_payees.json"
    if not path.exists():
        return
    try:
        for num, item in json.loads(path.read_text(encoding="utf-8")).items():
            CHECK_LABELS[str(num).strip()] = (str(item[0]).strip(), str(item[1]).strip())
    except (ValueError, OSError, TypeError, IndexError, AttributeError) as exc:
        print(f"Could not read {path.name} ({exc}); checks keep the generic label.")


def load_settings_file():
    """Personal figures live in brokerage_yearend.json next to the ledger, never in the script."""
    load_check_labels()
    load_vendor_overrides()
    path = ANALYSIS_DIR / "brokerage_yearend.json"
    if not path.exists():
        return
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        for year, items in data.items():
            BROKER_YEAREND[int(year)] = [(str(label), float(amount)) for label, amount in items]
    except (ValueError, OSError, TypeError) as exc:
        print(f"Could not read {path.name} ({exc}); Brokerage cross-check skipped.")
MEDICAL = "Medical & Dental"
# Sub-types inside Medical & Dental (the labels used in a medical spend list). Used for card
# charges; a spend list's own category label always wins. First match wins; anything else
# in Medical & Dental is "Medical Provider". Add your own in personal_config.py (MEDICAL_TYPES_EXTRA).
MEDICAL_TYPES = [
    (r"\bCVS\b|WALGREENS|RITE AID|PHARMACY", "Prescription"),
    (r"DENTAL|DENTIST|ENDODONT|ORTHODONT|\bDMD\b|\bDDS\b|PERIODON", "Dental"),
]
MEDICAL_TYPES = list(_pc("MEDICAL_TYPES_EXTRA", [])) + MEDICAL_TYPES
# Category labels in a spend list mapped to a main category (label itself is kept as sub-type).
FILE_CATEGORY_MAP = {"medical provider": MEDICAL, "prescription": MEDICAL, "dental": MEDICAL,
                     "therapy": MEDICAL, "vision": MEDICAL}

# (regex on the statement description, clean vendor name or None, category).
# First match wins, so specific rules come before general ones. A vendor of None
# keeps the auto-cleaned name. Anything matching no rule is "Uncategorized" and the
# Pareto then ranks it by vendor instead of by category.
FRAUD = "Disputed / Suspected Fraud"
CREDITS = "Credits & Rewards"
CATEGORY_RULES = [
    # --- credits and rewards
    (r"CONDITIONAL CREDIT FOR DISPUTE", "CITIBANK DISPUTE CREDIT", CREDITS),
    (r"THANKYOU POINTS REDEEMED", "THANKYOU POINTS REDEEMED", CREDITS),
    (r"CATEGORY SPEND BONUS", "CATEGORY SPEND BONUS", CREDITS),
    # --- merchants seen on the Apple Card statements
    (r"APPLE\.COM/BILL", "APPLE.COM/BILL", "Electronics & Software"),
    (r"^GIANT\b", "GIANT", "Groceries"),
    (r"DOLLAR-GENERAL", "DOLLAR GENERAL", "General Retail"),
    (r"BATTERIES PLUS|HUDSON NEWS", None, "General Retail"),
    (r"7-ELEVEN", "7-ELEVEN", "Fuel & Convenience"),
    (r"HARBOR FREIGHT", "HARBOR FREIGHT TOOLS", "Home Improvement"),
    (r"OLD NAVY|FAMOUS ?FOOTWEAR", None, "Clothing"),
    (r"UPS STORE", "THE UPS STORE", "Government & Postal"),
    (r"^TEMU|TIKTOK ?SHOP", None, "Online Shopping"),
    (r"PARKING", None, "Auto & Transportation"),
    (r"OLIVE GARDEN", "OLIVE GARDEN", "Dining"),
    (r"ARBYS|\bKFC\b|POPEYES|DAIRY QUEEN|5GUYS|^DD \*", None, "Dining"),
    # --- warehouse club, fuel, groceries, general retail
    (r"COSTCO GAS", "COSTCO GAS", "Fuel & Convenience"),
    (r"WWW COSTCO COM", "COSTCO.COM", "Warehouse Club"),
    (r"COSTCO", "COSTCO WAREHOUSE", "Warehouse Club"),
    (r"SHELL OIL", "SHELL", "Fuel & Convenience"),
    (r"^BP#|^BP ", "BP", "Fuel & Convenience"),
    (r"SAFEWAY FUEL", "SAFEWAY FUEL", "Fuel & Convenience"),
    (r"\bWAWA\b", "WAWA", "Fuel & Convenience"),
    (r"ROYAL FARMS", "ROYAL FARMS", "Fuel & Convenience"),
    (r"\b(SUNOCO|EXXON|MOBIL|SPEEDWAY|CHEVRON|SHEETZ|GULF OIL|SUNOCO)\b", None, "Fuel & Convenience"),
    (r"SAFEWAY", "SAFEWAY", "Groceries"),
    (r"\b(WEGMANS|ACME|GIANT FOOD|SHOPRITE|ALDI|TRADER JOE|WHOLE FOODS|KROGER|WEIS|LIDL|FOOD LION|SPROUTS)\b",
     None, "Groceries"),
    (r"WM SUPERCENTER|WAL-MART|WALMART", "WALMART", "General Retail"),
    (r"OCEAN STATE JOB LOT", "OCEAN STATE JOB LOT", "General Retail"),
    (r"\b(TARGET|KOHLS|TJ MAXX|MARSHALLS|ROSS STORES|DOLLAR TREE|DOLLAR GENERAL|FIVE BELOW)\b", None, "General Retail"),
    (r"AMAZON|AMZN", "AMAZON", "Online Shopping"),
    (r"\b(EBAY|ETSY)\b", None, "Online Shopping"),
    # --- home
    (r"HOME DEPOT", "HOME DEPOT", "Home Improvement"),
    (r"\bLOWES\b|LOWE'S", "LOWES", "Home Improvement"),
    (r"\b(MENARDS|ACE HARDWARE)\b", None, "Home Improvement"),
    # --- medical and dental
    (r"\bCVS\b", "CVS PHARMACY", "Medical & Dental"),
    (r"\b(WALGREENS|RITE AID)\b|PHARMACY", None, "Medical & Dental"),
    (r"\b(DENTAL|DENTIST|ORTHODON|PEDIATRIC|HOSPITAL|URGENT CARE|LABCORP|QUEST DIAG)\b", None, "Medical & Dental"),
    # --- insurance, auto, transportation
    (r"\b(GEICO|STATE FARM|PROGRESSIVE|ALLSTATE)\b", None, "Insurance"),
    (r"ENTERPRISE RENT", "ENTERPRISE RENT-A-CAR", "Auto & Transportation"),
    (r"AAA MIDATL|\bAAA\b", "AAA", "Auto & Transportation"),
    (r"\bUBER\b|\bLYFT\b", "UBER", "Auto & Transportation"),
    # --- travel and lodging
    (r"TRAVELOCITY", "TRAVELOCITY", "Travel & Lodging"),
    (r"^UNITED\b", "UNITED AIRLINES", "Travel & Lodging"),
    (r"^AMERICAN ?AI|^AMERICAN\d", "AMERICAN AIRLINES", "Travel & Lodging"),
    (r"^DELTA AIR", "DELTA AIR LINES", "Travel & Lodging"),
    (r"\b(SOUTHWEST|JETBLUE|FRONTIER AIR|SPIRIT AIR|AMTRAK|MARRIOTT|HILTON|HYATT|AIRBNB|VRBO|EXPEDIA)\b", None, "Travel & Lodging"),
    (r"HOTEL|RESORT|INN & SUITES|LODGE", None, "Travel & Lodging"),
    # --- dining
    (r"BURGER KING", "BURGER KING", "Dining"),
    (r"DUNKIN", "DUNKIN", "Dining"),
    (r"WAFFLE HOUSE", "WAFFLE HOUSE", "Dining"),
    (r"APPLEBEE", "APPLEBEES", "Dining"),
    (r"GOLDEN CORRAL", "GOLDEN CORRAL", "Dining"),
    (r"MCDONALD|STARBUCKS|CHIPOTLE|PANERA|DOMINO|PIZZA|RESTAURANT|CAFE\b|GRILL|DINER|DOORDASH|GRUBHUB|UBER ?EATS|BURGER|TAVERN|BREWERY|BREWING|BAKERY|DELI\b|STEAKHOUSE|PANCAKE", None, "Dining"),
    # --- recreation, clothing, subscriptions, other
    (r"\b(AMC THEAT|REGAL CINEMA|FANDANGO|TICKETMASTER|STUBHUB)\b", None, "Recreation & Entertainment"),
    (r"NETFLIX", "NETFLIX", "Subscriptions & Media"),
    (r"\b(SPOTIFY|HULU|DISNEY|YOUTUBE|PRIME VIDEO|HBO|PARAMOUNT|PEACOCK|AUDIBLE|KINDLE)\b", None, "Subscriptions & Media"),
    (r"BEST BUY|MYBESTBUY", "BEST BUY", "Electronics & Software"),
    (r"APPLE STORE", "APPLE STORE", "Electronics & Software"),
    (r"MICROSOFT", "MICROSOFT STORE", "Electronics & Software"),
    (r"\b(APPLE\.COM|GOOGLE|ADOBE|OPENAI|ANTHROPIC|CLAUDE\.AI|DROPBOX)\b", None, "Electronics & Software"),
    (r"T-?MOBILE", "T-MOBILE", "Utilities & Telecom"),
    (r"\b(VERIZON|COMCAST|XFINITY|AT&T|WATER AUTH)\b", None, "Utilities & Telecom"),
    (r"USPS", "USPS", "Government & Postal"),
]
# Your own vendors (local shops, providers, fraud lookalikes) go in personal_config.py as CATEGORY_RULES_FIRST,
# same format as above; they are tried before the rules in this file.
CATEGORY_RULES = list(_pc("CATEGORY_RULES_FIRST", [])) + CATEGORY_RULES

# ----------------------------------------------------------------------------
# Constants
# ----------------------------------------------------------------------------
PURCHASE, REFUND, PAYMENT, INTEREST, FEE, CBR = (
    "Purchase", "Refund/Credit", "Payment", "Interest", "Fee", "Credit Balance Refund")
UNRECON = "UNRECONCILED - REVIEW"

AMT = re.compile(r"^(minus|\+|-)?\s*\$?\s*(\d{1,3}(?:,\d{3})*\.\d{2}|\d+\.\d{2})$", re.I)
DATE_RE = re.compile(r"^\d{2}/\d{2}$")
PAYMENT_RE = re.compile(r"\b(payment|autopay|auto-pmt|thank you)\b", re.I)
US_STATES = set("AL AK AZ AR CA CO CT DE FL GA HI ID IL IN IA KS KY LA ME MD MA MI MN MS MO MT NE NV NH NJ NM NY NC ND OH OK OR PA RI SC SD TN TX UT VT VA WA WV WI WY DC".split())
SUMMARY_LABELS = [
    ("prev", r"previous balance"), ("payments", r"payments"), ("credits", r"credits"),
    ("purchases", r"purchases"), ("cash", r"cash advances"), ("fees", r"fees"),
    ("interest", r"interest"), ("new", r"new balance"),
]


# ----------------------------------------------------------------------------
# Ledger
# ----------------------------------------------------------------------------
def empty_ledger():
    return {"version": LEDGER_VERSION, "files": {}, "statements": {}, "transactions": [], "review": [],
            "recon": [], "bank_txns": [], "bank_recon": [], "bank_stmts": {}, "reference": {}}


def load_ledger():
    if not LEDGER_PATH.exists():
        return empty_ledger()
    try:
        ledger = json.loads(LEDGER_PATH.read_text(encoding="utf-8"))
    except (ValueError, OSError):
        ledger = None
    if not ledger or ledger.get("version") != LEDGER_VERSION:
        bak = LEDGER_PATH.with_name("processed_ledger_old_" + datetime.now().strftime("%H%M%S") + ".json")
        LEDGER_PATH.replace(bak)
        print(f"Old ledger moved to {bak.name}; statements will be re-read with the corrected parser.")
        return empty_ledger()
    for key, default in empty_ledger().items():
        ledger.setdefault(key, default)
    return ledger


def save_ledger(ledger):
    ANALYSIS_DIR.mkdir(parents=True, exist_ok=True)
    tmp = LEDGER_PATH.with_suffix(".tmp")
    tmp.write_text(json.dumps(ledger, indent=1), encoding="utf-8")
    tmp.replace(LEDGER_PATH)


def file_hash(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


# ----------------------------------------------------------------------------
# PDF parsing
# ----------------------------------------------------------------------------
def money(tok):
    m = AMT.match(tok.strip())
    if not m:
        return None
    v = float(m.group(2).replace(",", ""))
    return -v if (m.group(1) or "").lower() in ("minus", "-") else v


def read_lines(path):
    doc = fitz.open(path)
    out = []
    for page in doc:
        for ln in page.get_text().splitlines():
            ln = ln.strip()
            if ln:
                out.append(ln)
    doc.close()
    return out


def infer_date(tok, closing):
    if not tok:
        return None
    mth, day = int(tok[:2]), int(tok[3:5])
    yr = closing.year - 1 if mth > closing.month else closing.year
    try:
        return date(yr, mth, day)
    except ValueError:
        return None


def parse_statement(path):
    lines = read_lines(path)
    text = "\n".join(lines)
    res = {"rows": [], "summary": {}, "notes": []}

    m = re.search(r"Billing Period:\s*(\d{2}/\d{2}/\d{2})\s*-\s*(\d{2}/\d{2}/\d{2})", text)
    res["start"] = datetime.strptime(m.group(1), "%m/%d/%y").date() if m else None
    res["end"] = datetime.strptime(m.group(2), "%m/%d/%y").date() if m else None
    if not res["end"]:
        m = re.search(r"New balance as of (\d{2}/\d{2}/\d{2})", text)
        res["end"] = datetime.strptime(m.group(1), "%m/%d/%y").date() if m else None
    m = re.search(r"Account number ending in:?\s*(\d{4})", text)
    res["last4"] = m.group(1) if m else "UNK"
    head = " ".join(lines[:40]).lower()
    if "costco anywhere" in head or "costco" in head:
        res["card_name"] = "Citi Costco"
    elif "premier" in head:
        res["card_name"] = "Citi Premier"
    else:
        res["card_name"] = "Unknown Card"
    if not res["end"]:
        return res

    # ---- Account Summary (page 1)
    summ = {}
    try:
        i0 = next(i for i, l in enumerate(lines) if l.lower() == "account summary")
    except StopIteration:
        i0 = -1
    i = i0 + 1
    while 0 <= i0 and i < min(len(lines) - 1, i0 + 60) and len(summ) < 8:
        low = lines[i].lower()
        for key, pat in SUMMARY_LABELS:
            if key not in summ and re.fullmatch(pat, low):
                v = money(lines[i + 1])
                if v is not None:
                    summ[key] = v
                    i += 1
                break
        i += 1
    res["summary"] = summ

    # ---- Transaction sections
    section, cur = None, None
    rows = []
    for idx, l in enumerate(lines):
        low = l.lower()
        if re.match(r"payments, credits and adjustments", low):
            section, cur = "pay", None
            continue
        if low == "standard purchases" or low.startswith("promo purchase"):
            section, cur = "purch", None
            continue
        if re.fullmatch(r"cash advances?", low):
            section, cur = "cash", None
            continue
        if low == "fees charged":
            section, cur = "fees", None
            continue
        if low == "interest charged":
            section, cur = "int", None
            continue
        if re.search(r"totals year-to-date|interest charge calculation", low):
            section, cur = None, None
            continue
        if section is None:
            continue
        if re.match(r"total (fees|interest) (for this period|charged in this billing period)", low):
            cur = None
            nxt = money(lines[idx + 1]) if idx + 1 < len(lines) else None
            if nxt is not None and idx + 1 < len(lines):
                lines[idx + 1] = "~total~"   # consume the amount that follows the label
            continue
        if l == "~total~":
            continue
        # amount printed BEFORE its "Total ... " label (older Premier layout)
        if (cur is None and money(l) is not None and idx + 1 < len(lines)
                and re.match(r"total (fees|interest) (for this period|charged in this billing period)",
                             lines[idx + 1].lower())):
            continue
        if DATE_RE.match(l):
            if cur is None or cur["desc"] or cur["amt"] is not None:
                cur = {"sec": section, "d1": l, "d2": None, "desc": [], "amt": None}
                rows.append(cur)
            elif cur["d2"] is None:
                cur["d2"] = l
            continue
        v = money(l)
        if v is not None and cur is not None and cur["amt"] is None:
            cur["amt"] = v
            continue
        if cur is not None and cur["amt"] is None:
            cur["desc"].append(l)
    res["rows"] = rows
    return res


def build_rows(p):
    """Turn parsed rows into typed transactions and reconcile to the statement's own totals."""
    end, s = p["end"], p["summary"]
    out, incomplete = [], 0
    for r in p["rows"]:
        if r["amt"] is None:
            incomplete += 1
            continue
        desc = r["desc"][0] if r["desc"] else ""
        amt = r["amt"]
        sec = r["sec"]
        if sec == "pay":
            typ = (PAYMENT if PAYMENT_RE.search(desc) else REFUND) if amt < 0 else CBR
            amount = abs(amt)
        elif sec in ("purch", "cash"):
            typ, amount = (PURCHASE, amt) if amt >= 0 else (REFUND, abs(amt))
        elif sec == "fees":
            typ, amount = FEE, amt
        else:
            typ, amount = INTEREST, amt
        out.append({"date": infer_date(r["d1"], end), "post": infer_date(r["d2"], end), "type": typ,
                    "desc": desc, "amount": round(amount, 2), "cash": sec == "cash", "sec": sec})

    def tot(types, sec=None):
        return round(sum(t["amount"] * (-1 if t["type"] == REFUND and t["sec"] != "pay" else 1)
                         for t in out if t["type"] in types and (sec is None or t["sec"] == sec)), 2)

    g = lambda k: s.get(k)
    have = len(s) >= 7
    # fees / interest: summary is authoritative; add the gap as its own line
    for key, typ, label in (("fees", FEE, "Fees charged (statement total)"),
                            ("interest", INTEREST, "Interest charged (statement total)")):
        if g(key) is not None:
            diff = round(g(key) - tot([typ]), 2)
            if abs(diff) > 0.005:
                out.append({"date": end, "post": None, "type": typ, "desc": label, "amount": diff,
                            "cash": False, "sec": "adj"})

    # purchases: standard/promo purchases (net of in-section credits) + credit-balance refunds
    pur_rows = round(sum((t["amount"] if t["type"] == PURCHASE else -t["amount"])
                         for t in out if t["sec"] == "purch") + sum(t["amount"] for t in out if t["type"] == CBR), 2)
    cash_rows = round(sum(t["amount"] for t in out if t["sec"] == "cash"), 2)
    pur_diff = round((g("purchases") or 0) - pur_rows, 2) if g("purchases") is not None else 0.0
    if abs(pur_diff) > 0.005:
        out.append({"date": end, "post": None, "type": PURCHASE if pur_diff > 0 else REFUND,
                    "desc": "Amount on statement that could not be itemized",
                    "amount": abs(pur_diff), "cash": False, "sec": "unrecon"})

    # payments + credits
    pay_rows = round(sum(t["amount"] for t in out if t["type"] in (PAYMENT, REFUND) and t["sec"] == "pay"), 2)
    pay_target = None
    if have or ("new" in s and "prev" in s):
        credits = g("credits") or 0
        pay = g("payments")
        if pay is None:  # label printed without a value; derive from the balance identity
            pay = round(g("new") - g("prev") - credits - (g("purchases") or 0) - (g("cash") or 0)
                        - (g("fees") or 0) - (g("interest") or 0), 2)
        pay_target = round(-(pay + credits), 2)
    pay_diff = round(pay_rows - pay_target, 2) if pay_target is not None else None

    ident = None
    if "new" in s and "prev" in s:
        pay = g("payments")
        if pay is not None:
            ident = round(g("prev") + pay + (g("credits") or 0) + (g("purchases") or 0) + (g("cash") or 0)
                          + (g("fees") or 0) + (g("interest") or 0) - g("new"), 2)

    problems = []
    if not have:
        problems.append("Account Summary not fully read")
    if incomplete:
        problems.append(f"{incomplete} row(s) without an amount")
    if abs(pur_diff) > 0.005:
        problems.append(f"purchases off by {pur_diff:,.2f}")
    if pay_diff is not None and abs(pay_diff) > 0.005:
        problems.append(f"payments/credits off by {pay_diff:,.2f}")
    if ident is not None and abs(ident) > 0.005:
        problems.append(f"balance check off by {ident:,.2f}")
    recon = {
        "prev": g("prev"), "new": g("new"),
        "s_purchases": g("purchases"), "p_purchases": pur_rows,
        "s_paycred": pay_target, "p_paycred": pay_rows,
        "s_fees": g("fees"), "p_fees": tot([FEE]),
        "s_interest": g("interest"), "p_interest": tot([INTEREST]),
        "cash": g("cash"), "balance_check": ident,
        "status": "OK" if not problems else "CHECK: " + "; ".join(problems),
    }
    return out, recon


# ----------------------------------------------------------------------------
# Apple Card statements (Goldman Sachs): one PDF per month. Charges carry a Daily Cash column; Daily Cash is
# paid into Apple Cash, not credited to the card balance, so it is shown as a memo and never netted.
# ----------------------------------------------------------------------------
APPLE_CARD = "Apple Card"
APPLE_AMT = re.compile(r"^-?\$[\d,]+\.\d\d$")
APPLE_DATE = re.compile(r"^\d\d/\d\d/\d{4}$")
APPLE_PCT = re.compile(r"^-?\d+(?:\.\d+)?%$")
APPLE_PERIOD = re.compile(r"([A-Z][a-z]{2}) (\d{1,2}) — ([A-Z][a-z]{2}) (\d{1,2}), (\d{4})")
APPLE_TAG = re.compile(r"\((RETURN|DISPUTE CREDIT|DISPUTE REVERSAL)\)\s*$")
APPLE_LABELS = {
    "prev_total": "Previous Total Balance", "new": "Total Balance",
    "payments": "Total payments for this period", "charges": "Total charges, credits and returns",
    "dcash": "Total Daily Cash this month", "interest": "Total interest for this month",
}


def _apple_money(tok):
    t = tok.strip()
    if not APPLE_AMT.match(t):
        return None
    v = float(t.replace("-", "").replace("$", "").replace(",", ""))
    return -v if t.startswith("-") else v


def is_apple_pdf(path):
    try:
        doc = fitz.open(path)
        text = "\n".join(doc[i].get_text() for i in range(min(len(doc), 3)))
        doc.close()
    except Exception:
        return False
    return "Apple Card Customer" in text and "Goldman Sachs" in text


def apple_merchant(desc):
    """Merchant name from an Apple Card description ('NAME STORE# ADDRESS CITY ZIP ST COUNTRY')."""
    toks = APPLE_TAG.sub("", desc).split()
    keep = []
    for i, tok in enumerate(toks):
        if i > 0:
            if tok.endswith(",") or tok in ("-", "PO", "UNKNOWN") or tok.startswith("#"):
                break
            if re.search(r"\d", tok):
                m = re.fullmatch(r"([A-Za-z.&'\-]{3,})\d+", tok)      # e.g. TOOLS31962 -> TOOLS
                if m:
                    keep.append(m.group(1))
                break
        keep.append(tok)
    return " ".join(keep) or desc


def parse_apple_statement(path):
    doc = fitz.open(path)
    res = {"rows": [], "summary": {}, "start": None, "end": None, "notes": []}
    text_lines = []
    sec, cur = None, None
    rows = []

    def close():
        nonlocal cur
        cur = None

    for page in doc:
        text_lines += [l.strip() for l in page.get_text().splitlines() if l.strip()]
        spans = []
        for b in page.get_text("dict")["blocks"]:
            for ln in b.get("lines", []):
                for sp in ln["spans"]:
                    t = sp["text"].strip()
                    if t:
                        spans.append((sp["bbox"][1], sp["bbox"][0], t))
        spans.sort()
        lines, grp = [], []
        for y, x, t in spans:                      # group spans that sit on one printed line
            if grp and y - grp[0][0] > 2.5:
                lines.append(sorted(grp, key=lambda s: s[1]))
                grp = []
            grp.append((y, x, t))
        if grp:
            lines.append(sorted(grp, key=lambda s: s[1]))
        for ln in lines:
            y, x0, first = ln[0]
            if y < 115 or y > 742:                 # running header / footer of every page
                continue
            if x0 < 60:
                if first == "Payments":
                    sec, cur = "pay", None
                    continue
                if first == "Transactions":
                    sec, cur = "tx", None
                    continue
                if first in ("Date", "Interest Charged", "Daily Cash", "Legal", "Installments") or first.startswith("Total"):
                    if first != "Date":
                        sec = None if first in ("Interest Charged", "Daily Cash", "Legal", "Installments") else sec
                    close()
                    continue
                if APPLE_DATE.match(first) and sec:
                    cur = {"kind": sec, "date": first, "desc": [], "amt": None, "dcash": 0.0, "adj": 0.0, "sign": 1}
                    rows.append(cur)
                    items = ln[1:]
                else:
                    items = ln if cur is not None else []
            else:
                items = ln if cur is not None else []
            if cur is None:
                continue
            adj_line = any("Daily Cash" in t for _, _, t in ln)     # 'Daily Cash Adjustment' line of a return
            for _, x, t in items:
                if x >= 500 and APPLE_AMT.match(t):
                    if adj_line:      # a return takes back the Daily Cash earned on it: printed in the Amount column,
                        v = _apple_money(t)                       # charged to the card, and counted in the statement's charges
                        cur["adj"] += v
                        cur["dcash"] += cur["sign"] * abs(v)
                    elif cur["amt"] is None:
                        cur["amt"] = _apple_money(t)
                elif 410 <= x < 500:
                    if APPLE_PCT.match(t):
                        cur["sign"] = -1 if t.startswith("-") else 1
                    elif APPLE_AMT.match(t):
                        cur["dcash"] += cur["sign"] * abs(_apple_money(t))
                elif x >= 95 and not APPLE_PCT.match(t) and "Daily Cash" not in t:
                    cur["desc"].append(t)
    doc.close()

    text = "\n".join(text_lines)
    m = APPLE_PERIOD.search(text)
    if m:
        try:
            ey = int(m.group(5))
            e = datetime.strptime(f"{m.group(3)} {m.group(4)} {ey}", "%b %d %Y").date()
            sm = datetime.strptime(m.group(1), "%b").month
            s = date(ey - 1 if sm > e.month else ey, sm, int(m.group(2)))
            res["start"], res["end"] = s, e
        except ValueError:
            pass
    for key, label in APPLE_LABELS.items():
        for i, l in enumerate(text_lines):
            if l == label:
                for nxt in text_lines[i + 1:i + 4]:
                    v = _apple_money(nxt)
                    if v is not None:
                        res["summary"][key] = v
                        break
                break
    for r in rows:
        r["desc"] = re.sub(r"\s+", " ", " ".join(r["desc"])).strip()
        r["dcash"] = round(r["dcash"], 2)
        r["adj"] = round(r["adj"], 2)
    res["rows"] = rows
    return res


def build_apple_rows(p):
    """Typed transactions for one Apple Card statement, reconciled to the statement's own totals."""
    end, s = p["end"], p["summary"]
    out, incomplete = [], 0
    for r in p["rows"]:
        if r["amt"] is None:
            incomplete += 1
            continue
        d = datetime.strptime(r["date"], "%m/%d/%Y").date()
        if r["kind"] == "pay":
            desc = re.sub(r"account ending in \d{4}", "bank account", r["desc"], flags=re.I)
            out.append({"date": d, "post": None, "type": PAYMENT, "desc": desc, "amount": round(-r["amt"], 2),
                        "cash": False, "sec": "pay", "dcash": 0.0, "name": ""})
        else:
            net_amt = round(r["amt"] + r["adj"], 2)      # a return nets the Daily Cash taken back on it
            typ, amount = (PURCHASE, net_amt) if net_amt >= 0 else (REFUND, -net_amt)
            desc = r["desc"] + (f" | incl. Daily Cash taken back {r['adj']:+,.2f}" if r["adj"] else "")
            out.append({"date": d, "post": None, "type": typ, "desc": desc, "amount": round(amount, 2),
                        "cash": False, "sec": "purch", "dcash": r["dcash"], "name": apple_merchant(r["desc"])})
    g = s.get
    pur_rows = round(sum((t["amount"] if t["type"] == PURCHASE else -t["amount"])
                         for t in out if t["sec"] == "purch"), 2)
    pur_diff = round(g("charges", pur_rows) - pur_rows, 2)
    if abs(pur_diff) > 0.005:
        out.append({"date": end, "post": None, "type": PURCHASE if pur_diff > 0 else REFUND,
                    "desc": "Amount on statement that could not be itemized", "amount": abs(pur_diff),
                    "cash": False, "sec": "unrecon", "dcash": 0.0, "name": ""})
    interest = g("interest", 0.0)
    if abs(interest) > 0.005:
        out.append({"date": end, "post": None, "type": INTEREST, "desc": "Interest charged (statement total)",
                    "amount": round(interest, 2), "cash": False, "sec": "adj", "dcash": 0.0, "name": ""})
    pay_rows = round(sum(t["amount"] for t in out if t["type"] == PAYMENT), 2)
    pay_target = round(-g("payments", -pay_rows), 2)
    pay_diff = round(pay_rows - pay_target, 2)
    dc_rows = round(sum(t["dcash"] for t in out), 2)                          # net of Daily Cash taken back on returns
    dc_earned = round(sum(t["dcash"] for t in out if t["dcash"] > 0), 2)      # what the statement totals
    dc_stmt = g("dcash")
    ident = None
    if "new" in s and "prev_total" in s:
        ident = round(s["prev_total"] + g("charges", pur_rows) + interest - pay_target - s["new"], 2)
    problems = []
    required = ("prev_total", "new", "payments", "charges")
    if any(k not in s for k in required):
        problems.append("Account figures not fully read")
    if incomplete:
        problems.append(f"{incomplete} row(s) without an amount")
    if abs(pur_diff) > 0.005:
        problems.append(f"charges off by {pur_diff:,.2f}")
    if abs(pay_diff) > 0.005:
        problems.append(f"payments off by {pay_diff:,.2f}")
    if ident is not None and abs(ident) > 0.005:
        problems.append(f"balance check off by {ident:,.2f}")
    if dc_stmt is not None and abs(dc_earned - dc_stmt) > 0.015:
        problems.append(f"Daily Cash rows {dc_earned:,.2f} vs statement {dc_stmt:,.2f}")
    recon = {
        "prev": g("prev_total"), "new": g("new"),
        "s_purchases": g("charges"), "p_purchases": pur_rows,
        "s_paycred": pay_target, "p_paycred": pay_rows,
        "s_fees": 0.0, "p_fees": 0.0, "s_interest": interest, "p_interest": interest,
        "cash": None, "balance_check": ident, "dcash": dc_rows,
        "status": "OK" if not problems else "CHECK: " + "; ".join(problems),
    }
    return out, recon


def process_apple_pdf(ledger, pdf, h):
    """Read one Apple Card statement. Returns True when it was recorded."""
    try:
        parsed = parse_apple_statement(pdf)
    except Exception as exc:
        print(f"ERROR reading {pdf.name}: {exc}. Skipped, not recorded; fix or remove the file.")
        return False
    if not parsed["end"]:
        print(f"Could not read the statement period, NOT recorded (will retry next run): {pdf.name}")
        return False
    card_tab = CARD_ALIASES.get(APPLE_CARD, APPLE_CARD)
    stmt_key = f"{card_tab}|{parsed['end'].isoformat()}"
    entry = {"name": pdf.name, "card": card_tab, "period_end": parsed["end"].strftime("%m/%d/%Y"),
             "rows": 0, "processed_at": datetime.now().isoformat(timespec="seconds")}
    if stmt_key in ledger["statements"]:
        entry["status"] = "Duplicate statement (same card and end date); not counted"
        print(f"Duplicate statement period, not counted: {pdf.name} ({card_tab} {entry['period_end']})")
    else:
        rows, recon = build_apple_rows(parsed)
        txns = [make_txn(t, card_tab, parsed["end"], pdf.name) for t in rows]
        ledger["transactions"].extend(txns)
        ledger["statements"][stmt_key] = h
        recon.update({"card": card_tab, "stmt_end": parsed["end"].isoformat(), "source": pdf.name})
        ledger["recon"].append(recon)
        entry["rows"] = len(txns)
        if recon["status"] != "OK":
            ledger["review"].append({"source": pdf.name, "line": recon["status"]})
        entry["status"] = "Processed - " + ("ties to statement" if recon["status"] == "OK" else recon["status"])
        print(f"Processed: {pdf.name} -> {card_tab}, ends {entry['period_end']}, {len(txns)} rows, "
              f"{'ties to statement' if recon['status'] == 'OK' else recon['status']}")
    ledger["files"][h] = entry
    return True


# ----------------------------------------------------------------------------
# Synchrony store cards (Amazon Prime Store Card and other Synchrony layouts): one PDF per billing cycle.
# Every purchase line carries the item name, which feeds the Store Card Items tab (what was bought).
# A card payment becomes a transfer only when the statement it pays is loaded (see link_card_payments).
# ----------------------------------------------------------------------------
SYN_CYCLE = re.compile(r"(\d+) Day Billing Cycle from (\d\d/\d\d/\d{4}) to (\d\d/\d\d/\d{4})")
SYN_MONEY = re.compile(r"^([+-])?\s*\$?\s*([\d,]*\d\.\d\d)$")
SYN_REF = re.compile(r"^(?=.*\d)[A-Z0-9]{8,}$")
SYN_TOKEN = re.compile(r"^[A-Za-z0-9]{12}$")
SYN_NOISE = re.compile(r"^(\(Continued on next page\)|PAGE \d+ of \d+|Visit us at .*|Transaction Detail \(Continued\)|"
                       r"Date|Reference #|Description|Amount|Account Number ending in \d{4})$")
SYN_SECTIONS = {"Payments": "pay", "Other Credits": "credit", "Purchases and Other Debits": "purch",
                "Fees Charged": "fees", "Fees": "fees", "Interest Charged": "interest"}


def _syn_money(tok):
    m = SYN_MONEY.match(tok.strip().replace("−", "-"))
    if not m:
        return None
    v = float(m.group(2).replace(",", ""))
    return -v if m.group(1) == "-" else v


def is_synchrony_pdf(path):
    try:
        doc = fitz.open(path)
        text = "\n".join(doc[i].get_text() for i in range(min(len(doc), 3)))
        doc.close()
    except Exception:
        return False
    return bool(re.search(r"\.syf\.com", text, re.I)) and "Account Number ending in" in text


def parse_synchrony_statement(path):
    lines = read_lines(path)
    text = "\n".join(lines)
    res = {"rows": [], "summary": {}, "table": None, "detail_fees": None, "detail_interest": None,
           "start": None, "end": None, "last4": "UNK", "brand": ""}
    m = re.search(r"Account Number ending in (\d{4})", text)
    if m:
        res["last4"] = m.group(1)
    m = re.search(r"https?://(?:www\.)?([\w-]+)\.syf\.com", text, re.I)
    if m:
        res["brand"] = m.group(1).lower()
    m = SYN_CYCLE.search(text)
    if m:
        res["start"] = datetime.strptime(m.group(2), "%m/%d/%Y").date()
        res["end"] = datetime.strptime(m.group(3), "%m/%d/%Y").date()
    # page-1 Account Summary: label line, then the value on the next line
    try:
        a0 = lines.index("Account Summary")
    except ValueError:
        a0 = None
    if a0 is not None:
        for i in range(a0 + 1, min(a0 + 40, len(lines) - 1)):
            ln = lines[i]
            if SYN_CYCLE.search(ln):
                break
            v = _syn_money(lines[i + 1])
            if v is None:
                continue
            if ln.startswith("Previous Balance as of"):
                res["summary"]["prev"] = v
            elif ln.startswith("New Balance as of"):
                res["summary"]["new"] = v
            elif ln == "Payments":
                res["summary"]["payments"] = v
            elif ln == "Other Credits":
                res["summary"]["credits"] = v
            elif ln.startswith("Purchases"):
                res["summary"]["purchases"] = v
            elif "Fee" in ln:
                res["summary"]["fees"] = v
            elif "Interest" in ln:
                res["summary"]["interest"] = v
    # Account Balance Summary table: Total row = previous, payments & credits, purchases+fees, interest, expired, new
    try:
        b0 = lines.index("Account Balance Summary")
        t0 = next(i for i in range(b0, len(lines)) if lines[i] == "Total")
        vals = []
        for ln in lines[t0 + 1:t0 + 7]:
            vals.append(0.0 if ln in ("-", "–") else _syn_money(ln))
        if len(vals) == 6 and None not in vals:
            res["table"] = vals
    except (ValueError, StopIteration):
        pass
    # Transaction Detail
    try:
        d0 = lines.index("Transaction Detail")
    except ValueError:
        return res
    rows, sec, i, n = [], None, d0 + 1, len(lines)
    after_totals = False
    while i < n:
        ln = lines[i]
        if re.search(r"Year-to-Date Fees and Interest|Interest Charge Calculation", ln):
            break
        if SYN_NOISE.match(ln):
            i += 1
            continue
        if ln.startswith("Total Fees Charged This Period") and i + 1 < n:
            res["detail_fees"] = _syn_money(lines[i + 1])
            after_totals = True
            i += 2
            continue
        if ln.startswith("Total Interest Charged This Period") and i + 1 < n:
            res["detail_interest"] = _syn_money(lines[i + 1])
            after_totals = True
            i += 2
            continue
        if ln in SYN_SECTIONS and i + 1 < n and _syn_money(lines[i + 1]) is not None:
            sec = SYN_SECTIONS[ln]
            i += 2
            continue
        if DATE_RE.match(ln) and (sec or after_totals):
            row = {"kind": "zone" if after_totals else sec, "d": ln, "ref": "", "desc": "", "amt": None, "item": ""}
            i += 1
            if i < n and SYN_REF.match(lines[i]) and _syn_money(lines[i]) is None:
                row["ref"] = lines[i]
                i += 1
            if i < n and _syn_money(lines[i]) is None and not DATE_RE.match(lines[i]):
                row["desc"] = lines[i]
                i += 1
            if i < n and _syn_money(lines[i]) is not None:
                row["amt"] = _syn_money(lines[i])
                i += 1
            parts, first = [], True
            while i < n:
                nx = lines[i]
                if SYN_NOISE.match(nx):
                    i += 1
                    continue
                if (DATE_RE.match(nx) or nx.startswith("Total Fees Charged") or nx.startswith("Total Interest Charged")
                        or re.search(r"Year-to-Date Fees and Interest|Interest Charge Calculation", nx)
                        or (nx in SYN_SECTIONS and i + 1 < n and _syn_money(lines[i + 1]) is not None)):
                    break
                if first and SYN_TOKEN.match(nx) and not re.search(r"\s", nx):
                    first = False
                    i += 1
                    continue
                first = False
                parts.append(nx)
                i += 1
            row["item"] = re.sub(r"\s+", " ", " ".join(parts)).strip()
            rows.append(row)
            continue
        i += 1
    res["rows"] = rows
    return res


def synchrony_card_name(p):
    brand = p.get("brand") or ""
    name = "Amazon Store Card" if brand == "amazon" else (f"{brand.title()} Store Card" if brand else "Synchrony Card")
    return f"{name} {p['last4']}"


def item_category(item, channel, base_category):
    """Category for a store-card purchase from the item name printed on the statement (keyword rules)."""
    if base_category not in ("Online Shopping", "Uncategorized", "General Retail"):
        return base_category
    up = (item or "").upper()
    if up:
        for pat, cat in ITEM_RULES:
            if re.search(pat, up):
                return cat
    if "DIGITAL" in (channel or "").upper():
        return "Subscriptions & Media"
    return base_category


def build_synchrony_rows(p):
    """Typed transactions for one Synchrony statement, reconciled to the statement's own totals."""
    end, tb = p["end"], p["table"]
    out, incomplete = [], 0
    for r in p["rows"]:
        if r["amt"] is None:
            incomplete += 1
            continue
        d = infer_date(r["d"], end)
        amt, kind, desc = r["amt"], r["kind"], r["desc"]
        channel = re.sub(r"\s+SEATTLE(\s+WA)?\s*$", "", desc).strip()
        item = r["item"]
        base = {"date": d, "post": None, "cash": False, "dcash": 0.0, "name": "", "item": "", "channel": ""}
        if kind == "pay":
            out.append({**base, "type": PAYMENT, "desc": desc, "amount": round(-amt, 2), "sec": "pay"})
        elif kind in ("credit", "purch"):
            if kind == "credit":
                typ, amount = (REFUND, abs(amt)) if amt <= 0 else (PURCHASE, amt)
            else:
                typ, amount = (PURCHASE, amt) if amt >= 0 else (REFUND, -amt)
            out.append({**base, "type": typ, "desc": channel + (f" | {item}" if item else ""),
                        "amount": round(amount, 2), "sec": "credit" if kind == "credit" else "purch",
                        "name": channel, "item": item, "channel": channel})
        else:   # fees / interest section or the rows printed after the period totals
            up = desc.upper()
            if kind == "interest" or up.startswith("INTEREST CHARGE"):
                if abs(amt) > 0.004:
                    out.append({**base, "type": INTEREST, "desc": desc, "amount": round(amt, 2), "sec": "interest"})
            elif kind == "fees" or "FEE" in up:
                if abs(amt) > 0.004:
                    out.append({**base, "type": FEE, "desc": desc, "amount": round(amt, 2), "sec": "fees"})
            elif abs(amt) > 0.004:
                incomplete += 1
    s_fees = p["detail_fees"] if p["detail_fees"] is not None else 0.0
    s_int = p["detail_interest"] if p["detail_interest"] is not None else (tb[3] if tb else 0.0)
    sm = p["summary"]
    # Purchases on the statement = the table's "Purchases, Fees & Other Debits" less the fees printed separately
    s_pur = round(tb[2] - s_fees, 2) if tb else sm.get("purchases")
    pur_rows = round(sum((t["amount"] if t["type"] == PURCHASE else -t["amount"])
                         for t in out if t["sec"] == "purch"), 2)
    pur_diff = round((s_pur if s_pur is not None else pur_rows) - pur_rows, 2)
    if abs(pur_diff) > 0.005:
        out.append({"date": end, "post": None, "type": PURCHASE if pur_diff > 0 else REFUND, "desc":
                    "Amount on statement that could not be itemized", "amount": abs(pur_diff), "cash": False,
                    "sec": "unrecon", "dcash": 0.0, "name": "", "item": "", "channel": ""})
    for typ, tot_stmt, label in ((FEE, s_fees, "Fees charged (statement total)"),
                                 (INTEREST, s_int, "Interest charged (statement total)")):
        got = round(sum(t["amount"] for t in out if t["type"] == typ), 2)
        diff = round(tot_stmt - got, 2)
        if abs(diff) > 0.005:
            out.append({"date": end, "post": None, "type": typ, "desc": label, "amount": diff, "cash": False,
                        "sec": "adj", "dcash": 0.0, "name": "", "item": "", "channel": ""})
    pay_rows = round(sum(t["amount"] for t in out if t["type"] == PAYMENT), 2)
    cred_rows = round(sum(t["amount"] for t in out if t["type"] == REFUND and t["sec"] == "credit"), 2)
    paycred_rows = round(pay_rows + cred_rows, 2)
    s_paycred = tb[1] if tb else (round(-(sm.get("payments", 0) + sm.get("credits", 0)), 2) if "payments" in sm else None)
    ident = None
    if tb:
        ident = round(tb[0] - tb[1] + tb[2] + tb[3] + tb[4] - tb[5], 2)
    problems = []
    if not tb:
        problems.append("Account Balance Summary not read")
    if incomplete:
        problems.append(f"{incomplete} row(s) not read")
    if abs(pur_diff) > 0.005:
        problems.append(f"purchases off by {pur_diff:,.2f}")
    if s_paycred is not None and abs(paycred_rows - s_paycred) > 0.005:
        problems.append(f"payments/credits off by {round(paycred_rows - s_paycred, 2):,.2f}")
    if ident is not None and abs(ident) > 0.005:
        problems.append(f"balance check off by {ident:,.2f}")
    if tb and abs(tb[4]) > 0.005:
        problems.append(f"promotional balance adjustment {tb[4]:,.2f} not itemized")
    if tb and sm:   # two independent reads of the same figures (page-1 summary vs the balance table)
        for key, want, label in (("prev", tb[0], "previous balance"), ("new", tb[5], "new balance")):
            if key in sm and abs(sm[key] - want) > 0.005:
                problems.append(f"{label} differs between the summary and the balance table")
        if "payments" in sm and "credits" in sm and abs(-(sm["payments"] + sm["credits"]) - tb[1]) > 0.005:
            problems.append("payments/credits differ between the summary and the balance table")
        if "purchases" in sm and abs(sm["purchases"] - tb[2]) > 0.005:
            problems.append("purchases differ between the summary and the balance table")
    recon = {
        "prev": tb[0] if tb else sm.get("prev"), "new": tb[5] if tb else sm.get("new"),
        "s_purchases": s_pur, "p_purchases": pur_rows,
        "s_paycred": s_paycred, "p_paycred": paycred_rows,
        "s_fees": s_fees, "p_fees": round(sum(t["amount"] for t in out if t["type"] == FEE), 2),
        "s_interest": s_int, "p_interest": round(sum(t["amount"] for t in out if t["type"] == INTEREST), 2),
        "cash": None, "balance_check": ident,
        "status": "OK" if not problems else "CHECK: " + "; ".join(problems),
    }
    return out, recon


def process_synchrony_pdf(ledger, pdf, h):
    """Read one Synchrony statement. Returns True when it was recorded."""
    try:
        parsed = parse_synchrony_statement(pdf)
    except Exception as exc:
        print(f"ERROR reading {pdf.name}: {exc}. Skipped, not recorded; fix or remove the file.")
        return False
    if not parsed["end"]:
        print(f"Could not read the billing cycle, NOT recorded (will retry next run): {pdf.name}")
        return False
    card_tab = CARD_ALIASES.get(synchrony_card_name(parsed), synchrony_card_name(parsed))
    stmt_key = f"{card_tab}|{parsed['end'].isoformat()}"
    entry = {"name": pdf.name, "card": card_tab, "period_end": parsed["end"].strftime("%m/%d/%Y"),
             "rows": 0, "processed_at": datetime.now().isoformat(timespec="seconds")}
    if stmt_key in ledger["statements"]:
        entry["status"] = "Duplicate statement (same card and end date); not counted"
        print(f"Duplicate statement period, not counted: {pdf.name} ({card_tab} {entry['period_end']})")
    else:
        rows, recon = build_synchrony_rows(parsed)
        month_of = None
        ov = STATEMENT_MONTH_OVERRIDES.get(stmt_key)
        if ov:
            try:
                y, mth = (int(x) for x in str(ov).split("-"))
                month_of = date(y, mth, 1)
            except ValueError:
                print(f"STATEMENT_MONTH_OVERRIDES entry for {stmt_key} is not YYYY-MM; ignored.")
        txns = [make_txn(t, card_tab, parsed["end"], pdf.name, month_of) for t in rows]
        ledger["transactions"].extend(txns)
        ledger["statements"][stmt_key] = h
        recon.update({"card": card_tab, "stmt_end": parsed["end"].isoformat(), "source": pdf.name})
        ledger["recon"].append(recon)
        entry["rows"] = len(txns)
        if recon["status"] != "OK":
            ledger["review"].append({"source": pdf.name, "line": recon["status"]})
        note = ""
        if month_of:
            note = f" Counted in {month_of:%Y-%m} (STATEMENT_MONTH_OVERRIDES), not the closing month."
            ledger["review"].append({"source": pdf.name, "line": (
                f"{card_tab} statement closing {parsed['end']:%m/%d/%Y} is counted in {month_of:%B %Y} as configured "
                f"(STATEMENT_MONTH_OVERRIDES), so purchases on it dated after that month are included there too.")})
        entry["status"] = "Processed - " + ("ties to statement" if recon["status"] == "OK" else recon["status"]) + note
        print(f"Processed: {pdf.name} -> {card_tab}, ends {entry['period_end']}, {len(txns)} rows, "
              f"{'ties to statement' if recon['status'] == 'OK' else recon['status']}"
              + (f", counted in {month_of:%Y-%m}" if month_of else ""))
    ledger["files"][h] = entry
    return True


ITEM_RULES = list(_pc("ITEM_RULES_FIRST", [])) + [
    # (regex on the item name printed on a store-card statement, category). First match wins; a store-card line that
    # matches nothing stays 'Online Shopping'. Add your own as ITEM_RULES_FIRST in personal_config.py.
    (r"GIFT CARD", "Gifts"),
    (r"ROUTER|WI-?FI|\bUSB|HDMI|CABLE|CHARGER|ADAPTER|\bMOUSE\b|KEYBOARD|MONITOR|LAPTOP|MACBOOK|\bSSD\b|HARD DRIVE|"
     r"BLUETOOTH|EARBUD|HEADPHONE|SPEAKER|WEBCAM|CAMERA|TABLET|SMART PLUG|BATTER(Y|IES)|ETHERNET|AIRTAG|GL\.?INET|"
     r"EXTENSION CORD|POWER STRIP|SURGE PROTECT", "Electronics & Software"),
    (r"VITAMIN|SUPPLEMENT|MAGNESIUM|PROBIOTIC|OMEGA|CHEWABLE|SALINE|FIRST AID|BANDAGE|THERMOMETER|BLOOD PRESSURE|"
     r"PAIN RELIEF|ANTACID|ALLERGY|MELATONIN|COLLAGEN|BIOTIN|CALCIUM|ZINC|REMEDY|NASAL|\bENSURE\b", "Health & Fitness"),
    (r"NOZZLE|FLAP DISC|SANDING|SANDPAPER|DRILL|WRENCH|SCREW|\bBOLTS?\b|HARDWARE|\bSAW\b|\bHOSE\b|SPRAYER|CAULK|PLUMB|"
     r"FAUCET|VALVE|PAINT|ADHESIVE|EPOXY|WELD|GRINDER|\bTOOLS?\b|WORK ?LIGHT|LADDER|CONNECTOR|INSULATION|WEATHERSTRIP|"
     r"SEALANT", "Home Improvement"),
    (r"\bRACK\b|ORGANIZER|SHELF|CABINET|FURNITURE|\bCHAIR|\bTABLE\b|\bDESK\b|STORAGE|\bHOOKS?\b|CURTAIN|DRYER BALL|LAUNDRY|"
     r"TOWEL|BATHTUB|SHOWER|TRASH|VACUUM|\bMOP\b|CLEANER|CLEANING|SPONGE|DISPENSER|KITCHEN|COOKWARE|SERVING|POTPOURRI|"
     r"CANDLE|\bVASE\b|DECOR|CHRISTMAS|HALLOWEEN|WREATH|ORNAMENT|FLOWERS?|FAUX|FLORAL|PILLOW|BLANKET|\bSHEETS?\b|"
     r"\bLAMP\b|WATERING|IRRIGATION|GARDEN|\bPLANT|SEED|\bSOIL\b|FILTER|\bMATS?\b|WOOD|WINDOW|PURIFIER|HUMIDIFIER|"
     r"WASHING MACHINE|DETERGENT|OXICLEAN", "Household"),
    (r"\bSHIRTS?\b|\bPANTS\b|SOCKS|JACKET|\bCOAT\b|\bSHOES?\b|\bBOOTS?\b|GLOVES?|\bHAT\b|\bDRESS\b|UNDERWEAR|JEANS|"
     r"SWEATER|SLIPPERS?|HOODIE|\bROBES?\b|PAJAMA", "Clothing"),
    (r"SHAMPOO|CONDITIONER|\bSOAP\b|LOTION|RAZOR|TOOTHBRUSH|TOOTHPASTE|DEODORANT|SKIN ?CARE|MOISTURIZ|\bHAIR\b|\bNAIL|"
     r"MAKEUP|COSMETIC|EYE MASK", "Personal Care"),
    (r"\bTEA\b|COFFEE|SNACK|CANDY|COOKIE|SAUCE|SPICE|\bNUTS?\b|FLOUR|SUGAR|HONEY|JERKY|CEREAL|OLIVE OIL|PROTEIN|"
     r"CHOCOLATE", "Groceries"),
    (r"\bBOOKS?\b|NOVEL|PAPERBACK|\bDVD\b|BLU-?RAY", "Recreation & Entertainment"),
]


def link_card_payments(ledger):
    """A payment to a card that has statements loaded is a transfer, but only when the card's own statement shows
    the same payment (same amount, within a few days). Unmatched payments stay expenses so nothing is lost.
    Runs every time, so loading more statements later fixes earlier runs."""
    ledger["review"] = [r for r in ledger["review"] if not r.get("line", "").startswith("Card payments:")]
    for t in ledger["bank_txns"]:
        if "link_orig" in t:
            t["flow"], t["cat"] = t.pop("link_orig")
            t["note"] = ""
    for vendor, card, kind, need_loaded in CARD_LINKS:
        card_pays = [t for t in ledger["transactions"] if t["card"].startswith(card) and t["type"] == PAYMENT]
        if not card_pays:
            continue
        if need_loaded:
            recon = [rc for rc in ledger["recon"] if rc["card"].startswith(card)]
            for t in card_pays:     # does the statement this payment settles (previous balance) exist in the ledger?
                cur = next((rc for rc in recon if rc["stmt_end"] == t["stmt_end"]), None)
                t["settles_loaded"] = bool(cur and cur.get("prev") and cur["prev"] > 0.005 and any(
                    rc["stmt_end"] < t["stmt_end"] and rc.get("new") is not None and abs(rc["new"] - cur["prev"]) <= 0.005
                    for rc in recon))
            card_pays = [t for t in card_pays if t["settles_loaded"]]
        pays = sorted((datetime.strptime(t["date"], "%Y-%m-%d").date(), t["amount"], t["source"]) for t in card_pays)
        cands = [t for t in ledger["bank_txns"]
                 if t["vendor"] == vendor and t["flow"] == "Expense" and t["cat"] == OTHER_CARDS and t["amount"] < 0]
        used, unmatched = set(), []
        for pd, amt, src in pays:
            best = None
            for i, t in enumerate(cands):
                if i in used or abs(-t["amount"] - amt) > 0.005:
                    continue
                gap = (datetime.strptime(t["date"], "%Y-%m-%d").date() - pd).days
                if -3 <= gap <= 10 and (best is None or abs(gap) < best[0]):
                    best = (abs(gap), i)
            if best:
                t = cands[best[1]]
                used.add(best[1])
                t["link_orig"] = [t["flow"], t["cat"]]
                t["flow"], t["cat"] = "Transfer", kind
                t["note"] = f"Matches {card} statement payment of {pd:%m/%d/%Y}"
            else:
                unmatched.append((pd, amt, src))
        for pd, amt, src in unmatched:
            ledger["review"].append({"source": src, "line": (
                f"Card payments: {card} statement shows a payment of ${amt:,.2f} on {pd:%m/%d/%Y} with no matching "
                f"{BANK_NAME} withdrawal (same amount, within a few days). It was made from another account or not yet "
                f"loaded; the card spend is still counted.")})
        left = [t for i, t in enumerate(cands) if i not in used]
        if left and need_loaded:
            ledger["review"].append({"source": left[0]["source"], "line": (
                f"Card payments: {len(left)} {BANK_NAME} payment(s) to {card} totalling ${-sum(t['amount'] for t in left):,.2f} "
                f"do not settle a loaded {card} statement (the statement they pay is not loaded), so they stay counted as "
                f"expenses ('{OTHER_CARDS}'). Load the earlier statements to clear them; nothing is dropped or counted twice.")})
        elif left:
            ledger["review"].append({"source": left[0]["source"], "line": (
                f"Card payments: {len(left)} {BANK_NAME} payment(s) to {card} totalling ${-sum(t['amount'] for t in left):,.2f} "
                f"match no payment on the loaded {card} statements, so they stay counted as expenses "
                f"('{OTHER_CARDS}'). Load the statement for that month to clear them.")})


# ----------------------------------------------------------------------------
# Excel / CSV spend lists
# ----------------------------------------------------------------------------
DATE_FORMATS = ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%m-%d-%Y", "%b %d, %Y", "%B %d, %Y")
VENDOR_HEADS = {"provider", "vendor", "payee", "merchant", "paid to", "description", "name", "who"}
AMOUNT_HEADS = {"amount", "charge", "cost", "debit", "withdrawal", "paid amount"}


def to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        for f in DATE_FORMATS:
            try:
                return datetime.strptime(v.strip(), f).date()
            except ValueError:
                pass
    return None


def to_amount(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        t = v.strip().replace("$", "").replace(",", "")
        neg = t.startswith("(") and t.endswith(")")
        t = t.strip("()")
        try:
            x = float(t)
        except ValueError:
            return None
        return -x if neg else x
    return None


def read_tables(path):
    """Yield (sheet name, rows) for every sheet of an Excel file, or the one table of a CSV."""
    if path.suffix.lower() == ".csv":
        with open(path, newline="", encoding="utf-8-sig") as f:
            yield path.stem, [list(r) for r in csv.reader(f)]
        return
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    for ws in wb.worksheets:
        yield ws.title, [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()


def parse_ledger_sheet(path):
    """Read a manual spend list. Returns None if no table with date + payee + amount columns is found."""
    for _, rows in read_tables(path):
        head = None
        for i, r in enumerate(rows[:15]):
            cells = [str(c).strip().lower() if c is not None else "" for c in r]
            di = next((j for j, c in enumerate(cells) if "date" in c), None)
            vi = next((j for j, c in enumerate(cells) if c in VENDOR_HEADS), None)
            ai = next((j for j, c in enumerate(cells) if c in AMOUNT_HEADS), None)
            if di is not None and vi is not None and ai is not None:
                ci = next((j for j, c in enumerate(cells) if c == "category"), None)
                mi = next((j for j, c in enumerate(cells) if c == "month"), None)
                head = (i, di, vi, ai, ci, mi)
                break
        if not head:
            continue
        i0, di, vi, ai, ci, mi = head
        out, problems, stated, cat_totals = [], [], None, {}
        in_totals = False
        for r in rows[i0 + 1:]:
            r = list(r) + [None] * 8
            text = " ".join(str(c) for c in r if isinstance(c, str)).lower()
            if "total" in text and out and to_date(r[di]) is None:
                in_totals = True
            if in_totals:
                amt = to_amount(r[ai])
                if amt is not None:
                    label = str(r[0]).strip() if r[0] is not None else ""
                    if "grand total" in text or (stated is None and re.search(r"\btotal\b", text)
                                                  and not text.startswith("category")):
                        stated = amt if "grand total" in text or stated is None else stated
                    elif label:
                        cat_totals[label.lower()] = amt
                continue
            d, amt = to_date(r[di]), to_amount(r[ai])
            vendor = str(r[vi]).strip() if r[vi] is not None else ""
            if d is None and amt is None and not vendor:
                continue
            if d is None or amt is None or not vendor:
                problems.append(f"row with missing date, payee or amount skipped: {vendor or r[:4]}")
                continue
            cat = str(r[ci]).strip() if ci is not None and r[ci] is not None else ""
            if mi is not None and isinstance(r[mi], str) and r[mi].strip() \
                    and r[mi].strip().lower()[:3] != d.strftime("%b").lower():
                problems.append(f"month label '{r[mi]}' does not match date {d:%m/%d/%Y} ({vendor})")
            out.append({"date": d, "vendor": vendor, "category": cat, "amount": amt})
        if out:
            return {"rows": out, "stated": stated, "cat_totals": cat_totals, "problems": problems}
    return None


def account_for(path):
    for pat, name in ACCOUNT_BY_FILE:
        if re.search(pat, path.stem, re.I):
            return name
    return re.sub(r"[_\-]+", " ", path.stem).strip().title()


def make_sheet_txn(row, account, source):
    desc = row["vendor"]
    vendor, category = classify(desc, PURCHASE, "sheet", strip_state=False)
    label = row["category"]
    if category == "Uncategorized" and label:
        category = FILE_CATEGORY_MAP.get(label.lower(), label.title())
    subcat = ""
    if category == MEDICAL:
        subcat = label.title() if label and FILE_CATEGORY_MAP.get(label.lower()) == MEDICAL \
            else med_type(vendor, desc)
    amt = row["amount"]
    typ = REFUND if amt < 0 else PURCHASE
    flag, note = breach_info(vendor, desc)
    d = row["date"]
    return {
        "year": d.year, "month": d.month, "card": account, "type": typ, "vendor": vendor, "category": category,
        "subcat": subcat, "driver": f"{vendor} (vendor)" if category == "Uncategorized" else category,
        "amount": abs(amt), "breach": flag, "breach_note": note, "date": d.isoformat(), "post": "",
        "stmt_end": d.isoformat(), "source": source, "desc": desc,
    }


def sheet_recon(parsed, txns, account, source):
    read = round(sum(t["amount"] if t["type"] == PURCHASE else -t["amount"] for t in txns), 2)
    problems = list(parsed["problems"])
    if parsed["stated"] is None:
        problems.append("no total row in the file to check against")
    elif abs(read - parsed["stated"]) > 0.005:
        problems.append(f"rows add to {read:,.2f} but the file's total says {parsed['stated']:,.2f}")
    by_label = defaultdict(float)
    for r in parsed["rows"]:
        by_label[r["category"].lower()] += r["amount"]
    for label, tot in parsed["cat_totals"].items():
        if label in by_label and abs(round(by_label[label], 2) - tot) > 0.005:
            problems.append(f"category '{label}' rows add to {by_label[label]:,.2f}, file says {tot:,.2f}")
    return {
        "card": account, "stmt_end": max(t["date"] for t in txns), "source": source, "prev": None, "new": None,
        "s_purchases": parsed["stated"], "p_purchases": read, "s_paycred": None, "p_paycred": None,
        "s_fees": None, "s_interest": None, "cash": None, "balance_check": None,
        "status": "OK" if not problems else "CHECK: " + "; ".join(problems),
    }


def flag_cross_duplicates(ledger):
    """Same payee, same amount, within 5 days, on two different accounts: list for review (both stay counted)."""
    ledger["review"] = [r for r in ledger["review"] if not r.get("line", "").startswith("Possible duplicate")]
    is_sheet = lambda t: Path(t["source"]).suffix.lower() in (".xlsx", ".xlsm", ".csv")
    buys = [t for t in ledger["transactions"] if t["type"] == PURCHASE and t["date"]]
    for a in (t for t in buys if is_sheet(t)):
        da = datetime.strptime(a["date"], "%Y-%m-%d").date()
        for b in buys:
            if b["card"] == a["card"] or abs(b["amount"] - a["amount"]) > 0.005 or b["vendor"] != a["vendor"]:
                continue
            db = datetime.strptime(b["date"], "%Y-%m-%d").date()
            if abs((db - da).days) <= 5:
                ledger["review"].append({"source": a["source"], "line": (
                    f"Possible duplicate: {a['vendor']} ${a['amount']:,.2f} on {da:%m/%d/%Y} ({a['card']}) and "
                    f"${b['amount']:,.2f} on {db:%m/%d/%Y} ({b['card']}). Both are counted; remove one if it is the same bill.")})


# ----------------------------------------------------------------------------
# Vendor, category, breach
# ----------------------------------------------------------------------------
def clean_vendor(desc, strip_state=True):
    parts = re.split(r"\s{2,}", desc.strip())
    name = re.sub(r"\s+", " ", parts[0].upper())
    name = re.sub(r"\s*[#*]\s*\d+.*$", "", name)
    toks = [t for t in name.split() if not re.fullmatch(r"[\d\-\(\)\.]{4,}", t)]
    while strip_state and len(toks) > 1 and toks[-1] in US_STATES:
        toks.pop()
    name = re.sub(r"^(SQ|TST|PY|YPS|PHR|CBI|PP|GGL|SP)\s*\*\s*", "", " ".join(toks))
    name = re.sub(r"[*?~]", " ", name)
    name = re.sub(r"\s+", " ", name).strip(" -.,")
    return name or desc.upper()[:30]


def classify(desc, rtype, sec, strip_state=True):
    if rtype == PAYMENT:
        return "PAYMENT", "Payments"
    if rtype == CBR:
        return "CREDIT BALANCE REFUND", "Credit Balance Refund"
    if rtype == INTEREST:
        return "INTEREST CHARGED", "Interest"
    if rtype == FEE:
        return clean_vendor(desc) if desc and sec != "adj" else "FEES CHARGED", "Fees"
    if sec == "unrecon":
        return UNRECON, UNRECON
    upper = re.sub(r"\s+", " ", desc.upper())
    for pat, vendor, cat in CATEGORY_RULES:
        if re.search(pat, upper):
            return (vendor or clean_vendor(desc, strip_state)), cat
    return clean_vendor(desc, strip_state), "Uncategorized"


def breach_info(vendor, desc):
    hay = f"{vendor} {desc}".upper()
    for pat, note in BREACH_VENDORS:
        if re.search(pat, hay, re.I):
            return "Y", note
    return "", ""


def med_type(vendor, desc):
    hay = f"{vendor} {desc}".upper()
    for pat, label in MEDICAL_TYPES:
        if re.search(pat, hay):
            return label
    return "Medical Provider"


def make_txn(t, card_tab, end, source, month_of=None):
    vendor, category = classify(t.get("name") or t["desc"], t["type"], t["sec"])
    if "channel" in t and t["type"] in (PURCHASE, REFUND):      # store-card line: the item name sets the category
        category = item_category(t["item"], t["channel"], category)
    flag, note = ("", "")
    if t["type"] in (PURCHASE, REFUND):
        flag, note = breach_info(vendor, t["desc"])
    driver = f"{vendor} (vendor)" if category == "Uncategorized" else category
    iso = lambda d: d.isoformat() if d else ""
    mo = month_of or end
    rec = {
        "subcat": med_type(vendor, t["desc"]) if category == MEDICAL else "",
        "year": mo.year, "month": mo.month, "card": card_tab, "type": t["type"], "vendor": vendor,
        "category": category, "driver": driver, "amount": t["amount"], "breach": flag, "breach_note": note,
        "date": iso(t["date"]), "post": iso(t["post"]), "stmt_end": end.isoformat(), "source": source,
        "desc": t["desc"], "dcash": t.get("dcash", 0.0),
    }
    if "channel" in t:
        rec["channel"], rec["item"] = t["channel"], t["item"]
    return rec


# ----------------------------------------------------------------------------
# Bank statements (credit union): cash flow = income vs expenses, plus transfers
# ----------------------------------------------------------------------------
BANK_NAME = _pc("BANK_NAME", "Credit union")
BROKER = _pc("BROKERAGE_NAME", "Brokerage")             # where IRA / brokerage money moves from and to
ADVISOR = _pc("ADVISOR_NAME", "Advisor")                  # prefix of the advisor summary tabs and file
ADVISOR_FIRM = _pc("ADVISOR_FIRM", "your financial advisor")
SSA_OWNERS = _pc("SSA_OWNERS", None)                    # e.g. ("Alex", "Sam"): two monthly deposits, larger = first name
WHITE = 16777215
NUMTOK = re.compile(r"^-?\$?[\d,]*\d\.\d{2}$")
DATETOK = re.compile(r"^\d{2}/\d{2}$")

# Income sources, in the order shown on the Cash Flow tab.
INC_SSA1 = f"Social Security - {SSA_OWNERS[0]}" if SSA_OWNERS else "Social Security"
INC_SSA2 = f"Social Security - {SSA_OWNERS[1]}" if SSA_OWNERS else None
INC_BROKER = f"{BROKER} transfers in" + ("" if BROKER == "Brokerage" else " (IRA/brokerage)")
INC_CHECKS = "Check & branch deposits"
INC_INT = "Interest & dividends"
INC_REFUND = "Tax refunds"
INC_OTHER = "Other income"
# Deposit rules after Social Security: (regex on the description, income source label, vendor). First match wins.
# Replace the whole list with INCOME_RULES in personal_config.py (one entry per pension or employer).
INCOME_RULES = list(_pc("INCOME_RULES", [
    (r"^(?=.*REFUND)(?=.*(IRS|TREAS|DEPT OF REV|TAX))", INC_REFUND, "TAX REFUND"),
    (r"PENSION|ANNUIT", "Pension & annuity", "PENSION / ANNUITY"),
]))
_labels = []
for _l in [INC_SSA1, INC_SSA2] + [r[1] for r in INCOME_RULES]:
    if _l and _l != INC_REFUND and _l not in _labels:
        _labels.append(_l)
INCOME_SOURCES = _labels + [INC_BROKER, INC_CHECKS, INC_INT, INC_REFUND, INC_OTHER]

# Transfers (money moved, not earned or spent). Internal = between the credit union's own accounts.
TR_CITI = "Citi card payments (spend counted from the Citi statements)"
TR_APPLE = "Apple Card payments (spend counted from the Apple Card statements)"
TR_AMAZON = "Amazon Store Card payments (spend counted from the Amazon Store Card statements)"
TR_BROKER_OUT = f"Sent back to {BROKER}"
TR_INTERNAL = f"Between {BANK_NAME} accounts (nets to zero)"
TRANSFER_KINDS = [TR_CITI, TR_APPLE, TR_AMAZON, TR_BROKER_OUT, TR_INTERNAL]
# Card payments made from the bank that become transfers once the card's own statement shows the same payment.
# (bank vendor text, card tab name starts with, transfer kind, only when the statement it pays is loaded).
# A store card's payment on a statement settles the PREVIOUS statement, so it is a transfer only when that earlier
# statement (the purchases it pays for) is loaded; otherwise it stays an expense and the spend is not lost.
# Add Barclays / Chase here when their statements are read.
CARD_LINKS = [("APPLE CARD", "Apple Card", TR_APPLE, False),
              ("AMAZON STORE CARD", "Amazon Store Card", TR_AMAZON, True)]

OTHER_CARDS = "Credit card payments - other cards (no statements loaded)"
PLAN_529 = _pc("PLAN_529_LABEL", "529 plan contributions")
PLAN_529_VENDOR = _pc("PLAN_529_VENDOR", "529 PLAN CONTRIBUTION")
PLAN_529_PAT = _pc("PLAN_529_PAT", None)            # regex on ACH text for 529 contributions (None = no special rule)
MORTGAGE_PAT = _pc("MORTGAGE_PAT", r"MTG PYMT|MORTGAGE")
MORTGAGE_VENDOR = _pc("MORTGAGE_VENDOR", "MORTGAGE PAYMENT")
MORTGAGE_EXTRA_MIN = _pc("MORTGAGE_EXTRA_MIN", 0)   # a mortgage payment this large or larger is extra principal (0 = off)
CHECKS_CAT = "Checks written (payee not on statement)"
BANK_RULES = [
    # (regex on the full description, category, vendor or None)  -- first match wins
    (r"OVERDRAWN", "Bank & Card Fees", "OVERDRAFT FEE"),
    (r"RTN CK FEE|RETURN(ED)? ?(CK|CHECK) FEE", "Bank & Card Fees", "RETURNED CHECK FEE"),
    (r"INTERNATIONAL SERVICE ASSESSMENT|DEBIT CARD FEE|ATM FEE|SERVICE CHARGE", "Bank & Card Fees", "CARD / SERVICE FEE"),
    (r"TO LOAN", "Loan payments", "LOAN PAYMENT"),
    (MORTGAGE_PAT, "Housing - Mortgage", MORTGAGE_VENDOR),
    (r"USATAXPYMT|\bIRS\b", "Taxes", "IRS"),
    (r"VERIZ", "Utilities & Telecom", "VERIZON"),
    (r"T-MOBILE|T-M\b|PCS SVC", "Utilities & Telecom", "T-MOBILE"),
    (r"\bWM\b.*(AUTO|PMT)|WM AU|CO: WM|WASTE MGMT|WASTE MANAGEMENT", "Utilities & Telecom", "WASTE MANAGEMENT"),
    (r"EZ ?PASS", "Auto & Transportation", "EZ PASS"),
    (r"URGENT CARE|\bDR\.? |DERMATOLOGY|PHR\*ENT", "Medical & Dental", None),
    (r"APPLE CASH|CASH APP", "Person-to-person (Apple Cash / Cash App)", None),
    (r"FOREIGN CURRENCY", "Travel & Lodging", "FOREIGN CURRENCY"),
    (r"APPLE\.COM|APPLE COM ONE", "Electronics & Software", "APPLE"),
    (r"ATM #|AT ATM|^WITHDRAWAL$|^WITHDRAWAL A( |$)", "Cash withdrawals (ATM / branch)", "CASH WITHDRAWAL"),
    (r"^DRAFT \d+", CHECKS_CAT, None),
    (r"GIANT FUEL", "Fuel & Convenience", "GIANT FUEL"),
    (r"\bGIANT\b", "Groceries", "GIANT"),
    (r"HOME ?DEPOT", "Home Improvement", "HOME DEPOT"),
    (r"NISSAN|CAR WASH|SUBARU|HONDA|FORD\b", "Auto & Transportation", None),
    (r"NORDSTROM|MACY'?S|\bLOFT\b|J\.? ?CREW|AMERICAN EAGLE|AERIE|OLD NAVY|WACOAL|JOCKEY|DECKERS|UGG\b|LANDS ?END|BELK|FOOTWEAR|"
     r"TJ ?MAXX|KOHL", "Clothing", None),
    (r"NAIL|HAIR CUTTERY|\bULTA\b|SALON|BARBER", "Personal Care", None),
    (r"BUNDT|CHICKEN SALAD|CHICK-FIL-A|BBQ|SHAKE SHACK|TACO BELL|HOAGIE|SUSHI|WAWA|SONIC|WENDY|SUBWAY|\bKFC\b|PIZZA|DINER", "Dining", None),
    (r"DYSON|DOLLAR.GENERAL|OLD TIME POTTERY|THE COMPANY STORE|DICKS ?SPORTING|WAL-?MART|WALMART|NOTHING BUNDT", "General Retail", None),
]
# Your own payees (loan, insurer, utility, doctor, dealer) go in personal_config.py as BANK_RULES_FIRST, same format.
BANK_RULES = list(_pc("BANK_RULES_FIRST", [])) + BANK_RULES
# Card payments by ACH (text has TYPE:) to cards for which no statements are loaded.
OTHER_CARD_PAT = (r"BARCLAY|SYNCHRONY|SYN +NK|SYN +B|CAPITAL ?ONE|CAPITA\b|APPLECARD|APP +BANK|APP +ANK|"
                  r"AMZ.?_?PMT|AMZ_STORECR|AMZ +\b|BEST ?B(UY)?\b.*(PYMT|PMT)|BEST BUY CREDIT|CHASE CR|DIT C|"
                  r"EPAY|COMENITY|KOHL'?S|CHG PYMT|CO: 'S")
CITI_PAT = r"CITI AUTOPAY|CITI CARD|CITIBANK|CITICTP|CITIB\w*|CITICARDS|CITI +"
BROKER_PAT = _pc("BROKER_PAT", r"BROKERAGE")
SUMMARY_COLS = (187.2, 266.4, 352.8, 417.6, 496.8, 594.0)   # right edges of the SUMMARY FOR PERIOD columns
SSA_PAT = r"XXSOC|SSA +TRE|SSA TREAS|SOC SEC"


def _num(tok):
    return round(float(tok.replace("$", "").replace(",", "")), 2)


def _page_rows(page):
    items = []
    for b in page.get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            t = "".join(s["text"] for s in l["spans"]).strip()
            if t:
                items.append({"y": l["bbox"][1], "x0": l["bbox"][0], "x1": l["bbox"][2], "t": t,
                              "white": l["spans"][0]["color"] == WHITE})
    items.sort(key=lambda i: (round(i["y"], 0), i["x0"]))
    rows, cur = [], None
    for it in items:
        if cur is not None and abs(it["y"] - cur["y"]) <= 1.6:
            cur["items"].append(it)
        else:
            cur = {"y": it["y"], "items": [it]}
            rows.append(cur)
    for r in rows:
        r["items"].sort(key=lambda i: i["x0"])
    return rows


def is_bank_pdf(path):
    try:
        doc = fitz.open(path)
        text = "\n".join(doc[i].get_text() for i in range(min(len(doc), 6)))
        doc.close()
    except Exception:
        return False
    low = text.lower()
    return "summary for period" in low and ("member number" in low or BANK_NAME.lower() in low) and "statement of account" in low


def parse_bank_statement(path):
    doc = fitz.open(path)
    full = "\n".join(p.get_text() for p in doc)
    m = re.search(r"(\d\d/\d\d/\d\d)\s*-\s*(\d\d/\d\d/\d\d)", full)
    if not m:
        return None
    start = datetime.strptime(m.group(1), "%m/%d/%y").date()
    end = datetime.strptime(m.group(2), "%m/%d/%y").date()
    accts, order = {}, []
    acct = cur = None
    last_y = None
    mode = "txn"
    summary = []
    in_summary = False
    for pi, page in enumerate(doc):
        amt_x1 = bal_x1 = None
        cur, last_y = None, None            # a transaction never continues onto the next page
        for row in _page_rows(page):
            its = row["items"]
            texts = [i["t"] for i in its]
            line = " ".join(texts)
            low = line.lower()
            if any(i["white"] for i in its):
                label = re.sub(r"\s+", " ", re.sub(r"\(Cont\.\)", "", line)).strip()
                mm = re.match(r"^(.*?\D)\s*(\d{4})$", label)
                if not mm:
                    continue                  # promotional banner, not an account header
                name, tail = mm.group(1).strip(), mm.group(2)
                if name not in accts:
                    kind = "deposit" if re.search(r"SHARES|CLUB|CHECKING|SAVINGS|MONEY MARKET|CERT|IRA", name, re.I) else "loan"
                    accts[name] = {"name": name, "last4": tail, "kind": kind, "txns": [], "open": None, "end": None}
                    order.append(name)
                acct = accts[name]
                mode = "txn"
                cur, last_y = None, row["y"]
                continue
            if "summary for period" in low:
                in_summary = True
                continue
            if in_summary:
                nums = [(i["x1"], _num(i["t"])) for i in its if NUMTOK.match(i["t"])]
                if len(nums) >= 3:     # one value per column; a blanked cell stays None
                    cells = [None] * 6
                    for x, v in nums:
                        cells[min(range(6), key=lambda k: abs(SUMMARY_COLS[k] - x))] = v
                    summary.append(cells)
                continue
            if acct is None:
                continue
            if "AMOUNT" in texts and "BALANCE" in texts:
                for i in its:
                    if i["t"] == "AMOUNT":
                        amt_x1 = i["x1"]
                    if i["t"] == "BALANCE":
                        bal_x1 = i["x1"]
                last_y = row["y"]
                mode = "txn"
                continue
            if any(k in low for k in ("summary of cleared checks", "total overdraft", "nsf fees")):
                mode, cur = "skip", None
                continue
            if mode == "skip":
                continue
            first = its[0]
            if DATETOK.match(first["t"]) and first["x0"] < 70:
                amount = balance = None
                desc = []
                for i in its[1:]:
                    if NUMTOK.match(i["t"]) and i["x1"] > 470:
                        if bal_x1 and abs(i["x1"] - bal_x1) <= 4:
                            balance = _num(i["t"])
                        elif amt_x1 and abs(i["x1"] - amt_x1) <= 8:
                            amount = _num(i["t"])
                    else:
                        desc.append(i["t"])
                d = " ".join(desc).strip()
                cur = {"md": first["t"], "amount": amount, "balance": balance, "desc": [d] if d else []}
                acct["txns"].append(cur)
                last_y = row["y"]
                continue
            if cur is not None and last_y is not None and 0 <= row["y"] - last_y <= 12.5 and first["x0"] >= 70:
                cur["desc"].append(line)
                last_y = row["y"]
    doc.close()
    out = []
    for n in order:
        a = accts[n]
        if a["kind"] != "deposit":
            continue
        txns = []
        for t in a["txns"]:
            head = (t["desc"][0] if t["desc"] else "").lower()
            if head.startswith("balance forward"):
                a["open"] = t["balance"]
            elif head.startswith("ending balance"):
                a["end"] = t["balance"]
            elif t["amount"] is not None:
                txns.append(t)
        a["txns"] = txns
        out.append(a)
    # summary table rows follow the order of the deposit accounts
    for a, srow in zip(out, summary):
        a["summary"] = srow
    return {"start": start, "end": end, "accts": out,
            "loans": [accts[n] for n in order if accts[n]["kind"] != "deposit"]}


def _merchant_text(desc_lines):
    """The payee text of a transaction: the ACH 'CO:' name, or the lines after the transaction type."""
    full = " ".join(desc_lines)
    co = re.search(r"CO:\s*(.*?)(?:\s+NAME:|\s+ID:|$)", full)
    if co and co.group(1).strip():
        return co.group(1).strip()
    rest = " ".join(desc_lines[1:]) if len(desc_lines) > 1 else ""
    rest = re.sub(r"\b[A-Z]{2} - \d\d/\d\d\b.*$", "", rest)
    rest = re.sub(r"\b\d{3}[- ]\d{3}[- ]\d{4}\b|\b1 \d{3} \d{3}-\d{3}\b", " ", rest)
    return re.sub(r"\s+", " ", rest).strip()


def _card_name(U):
    for pat, name in (("BARCLAY", "BARCLAYCARD"), ("APPLECARD|APP +BANK|APP +ANK|APP BANK", "APPLE CARD"),
                      ("AMZ", "AMAZON STORE CARD"), ("SYN", "SYNCHRONY"), ("COMENITY", "COMENITY"),
                      ("BEST ?B", "BEST BUY CARD"), ("CAPITA", "CAPITAL ONE"), ("CHASE|DIT C|EPAY", "CHASE"),
                      ("KOHL|'S", "KOHL'S CARD")):
        if re.search(pat, U):
            return name
    return "OTHER CARD"


def classify_bank(t):
    """Returns (flow, category, vendor). flow: Income | Expense | Transfer."""
    lines = t["desc"] or [""]
    full = re.sub(r"\s+", " ", " ".join(lines)).strip()
    U = full.upper()
    amt = t["amount"]
    first = U.split(" ")[0] if U else ""
    reversal = first in ("WITHDRAWAL", "ADJUSTMENT", "DEBIT", "CREDIT") and amt > 0   # bank counts these as debits
    merchant = _merchant_text(lines) or full

    if re.search(r"TRANSFER (TO|FROM) SHARE|TRANSFER (TO|FROM) (CHECKING|CLUB|SAVINGS)", U) and "LOAN" not in U:
        return "Transfer", TR_INTERNAL, f"{BANK_NAME.upper()} ACCOUNT TRANSFER"
    if re.search(r"FED RTN CK|RETURNED? (CHECK|ITEM)", U):
        return "Income", INC_CHECKS, "RETURNED CHECK (REVERSAL)"
    if re.search(OVERDRAWN_PAT, U):
        return "Expense", "Bank & Card Fees", "OVERDRAFT FEE"
    if amt > 0 and not reversal:
        if re.search(SSA_PAT, U):
            return "Income", INC_SSA1, "SOCIAL SECURITY (SSA TREAS 310)"      # owner split later
        for pat, label, vendor in INCOME_RULES:
            if re.search(pat, U):
                return "Income", label, vendor
        if re.search(BROKER_PAT, U):
            return "Income", INC_BROKER, _pc("BROKER_VENDOR_IN", f"{BROKER.upper()} TRANSFER IN")
        if re.search(r"CITIADJUST|CITICARDS|CITI", U):
            return "Transfer", TR_CITI, "CITICARDS RETURN/ADJUSTMENT"
        if re.search(r"FOREIGN CURRENCY", U):
            return "Expense", "Travel & Lodging", "FOREIGN CURRENCY (RETURNED)"
        if re.search(r"DEBIT CARD FEE", U):
            return "Expense", "Bank & Card Fees", "CARD FEE REFUND"
        if re.search(r"DISPUTE", U):
            for pat, cat, vendor in BANK_RULES:
                if re.search(pat, U) and cat == "Disputed / Suspected Fraud":
                    return "Expense", cat, ((vendor or "DISPUTED ITEM").replace("(DISPUTED)", "").strip() + " (DISPUTE CREDIT)")
            vendor, cat = classify(merchant, PURCHASE, "")
            return "Expense", cat, vendor
        if re.search(r"DIVIDEND", U):
            return "Income", INC_INT, "DIVIDEND"
        if re.search(r"SHARED BRANCH|MOBILE|^DEPOSIT$|^DEPOSIT (BRANCH|ATM)|CHECK", U) or first == "DEPOSIT" and len(U) <= 12:
            return "Income", INC_CHECKS, "CHECK / BRANCH DEPOSIT"
        return "Income", INC_OTHER, merchant[:40] or "UNKNOWN DEPOSIT"
    # money out, or a merchant refund (positive "Withdrawal Adjustment ... Credit Voucher")
    if amt < 0:
        if PLAN_529_PAT and re.search(PLAN_529_PAT, U):
            return "Expense", PLAN_529, PLAN_529_VENDOR
        if re.search(BROKER_PAT, U) and not re.search(r"DEBIT CARD|POS #", U):
            return "Transfer", TR_BROKER_OUT, BROKER.upper()
        if re.search(r"CITI", U) and re.search(r"ACH|BILL PAYMENT|TYPE:", U):
            return "Transfer", TR_CITI, "CITI CARD PAYMENT"
        if re.search(r"ACH|TYPE:", U) and not re.search(r"POS #|DEBIT CARD|CHECK CARD", U) and re.search(OTHER_CARD_PAT, U):
            return "Expense", OTHER_CARDS, _card_name(U)
        if MORTGAGE_EXTRA_MIN and re.search(MORTGAGE_PAT, U) and amt <= -MORTGAGE_EXTRA_MIN:
            return "Expense", "Housing - Mortgage extra principal", f"{MORTGAGE_VENDOR} (EXTRA PRINCIPAL)"
    for pat, cat, vendor in BANK_RULES:
        if re.search(pat, U):
            return "Expense", cat, (vendor or clean_vendor(merchant))
    vendor, cat = classify(merchant, PURCHASE, "")
    if cat == "Uncategorized" and (vendor.startswith("WITHDRAWAL") or
                                   all(w in LOCATION_WORDS or w.isdigit() for w in vendor.split())):
        vendor = "MERCHANT NAME NOT LEGIBLE (" + vendor.title() + ")"
    return "Expense", cat, vendor


LOCATION_WORDS = set("MAIN STRT BLDG ROAD".split()) | set(_pc("LOCATION_WORDS_EXTRA", "").split())   # town names on statements


OVERDRAWN_PAT = r"OVERDRAWN"


def _is_reversal(t):
    head = (t["desc"][0] if t["desc"] else "").lower()
    return t["amount"] > 0 and bool(re.match(r"(withdrawal|adjustment)", head))


def bank_account_name(a):
    return f"{BANK_NAME} {a['name'].title()} {a['last4']}"


def make_bank_txns(parsed, source):
    end = parsed["end"]
    out = []
    for a in parsed["accts"]:
        for t in a["txns"]:
            mth, day = int(t["md"][:2]), int(t["md"][3:5])
            yr = end.year - 1 if mth > end.month else end.year
            flow, cat, vendor = classify_bank(t)
            # cash month = statement month (a row dated 01/31 that posts on the February statement counts in
            # February), so every month ties to the credit union's own opening and closing balances
            out.append({"acct": bank_account_name(a), "date": date(yr, mth, day).isoformat(), "year": end.year,
                        "month": end.month,
                        "amount": t["amount"], "balance": t["balance"], "flow": flow, "cat": cat, "vendor": vendor,
                        "desc": " | ".join(t["desc"]), "source": source, "note": ""})
    return out


def bank_recon_row(a, parsed, source):
    sm = a.get("summary") or [None] * 6
    op = a["open"] if a["open"] is not None else (sm[0] if sm[0] is not None and not a["txns"] else 0.0)
    en = a["end"] if a["end"] is not None else (sm[4] if sm[4] is not None and not a["txns"] else op)
    total = round(sum(t["amount"] for t in a["txns"]), 2)
    dep = round(sum(t["amount"] for t in a["txns"] if t["amount"] > 0 and not _is_reversal(t)), 2)
    wd = round(total - dep, 2)
    bal, pend, chain = op, 0.0, 0
    for t in a["txns"]:
        pend += t["amount"]
        if t["balance"] is not None:
            if abs(bal + pend - t["balance"]) > 0.015:
                chain += 1
            bal, pend = t["balance"], 0.0
    issues = []
    if abs(op + total - en) > 0.015:
        issues.append(f"opening {op:,.2f} + rows {total:,.2f} does not equal closing {en:,.2f}")
    if chain:
        issues.append(f"{chain} running-balance break(s)")
    if sm[1] is not None and abs(sm[1] - dep) > 0.015:
        issues.append(f"deposits read {dep:,.2f} vs statement {sm[1]:,.2f}")
    if sm[2] is not None and abs(sm[2] - wd) > 0.015:
        issues.append(f"withdrawals read {wd:,.2f} vs statement {sm[2]:,.2f}")
    if sm[4] is not None and abs(sm[4] - en) > 0.015:
        issues.append(f"closing {en:,.2f} vs summary {sm[4]:,.2f}")
    return {"acct": bank_account_name(a), "stmt_end": parsed["end"].isoformat(), "year": parsed["end"].year,
            "month": parsed["end"].month, "source": source, "open": op, "close": en, "rows_sum": total,
            "dep_read": dep, "dep_stmt": sm[1], "wd_read": wd, "wd_stmt": sm[2], "chain": chain, "n": len(a["txns"]),
            "status": "OK" if not issues else "CHECK: " + "; ".join(issues)}


SSA_VENDOR = "SOCIAL SECURITY (SSA TREAS 310)"


def apply_check_labels(ledger):
    """Give written checks the payee and category from check_payees.json. Runs every time, so editing the file and
    re-running is enough (no need to delete the ledger)."""
    for t in ledger["bank_txns"]:
        m = re.search(r"\bDRAFT\s+(\d{3,6})\b", t["desc"], re.I)
        if not m or t["flow"] != "Expense":
            continue
        label = CHECK_LABELS.get(m.group(1))
        if label:
            t.setdefault("cat_orig", t["cat"])
            t.setdefault("vendor_orig", t["vendor"])
            t["vendor"], t["cat"], t["note"] = label[0], label[1], f"Check {m.group(1)}"
        elif "cat_orig" in t:      # label removed from check_payees.json: back to the generic check line
            t["cat"], t["vendor"], t["note"] = t.pop("cat_orig"), t.pop("vendor_orig"), ""
        elif t["note"].startswith("Check ") and t["cat"] != CHECKS_CAT:
            t["cat"], t["vendor"], t["note"] = CHECKS_CAT, "DRAFT", ""


def apply_vendor_overrides(ledger):
    """Re-label card charges listed in vendor_overrides.json. Runs every time, so editing the file and re-running is
    enough; removing an entry restores the original vendor and category."""
    for t in ledger["transactions"]:
        if t["type"] not in (PURCHASE, REFUND):
            continue
        hit = next(((v, c) for rx, v, c in VENDOR_OVERRIDES if rx.search(t["desc"])), None)
        if hit:
            t.setdefault("ov_orig", [t["vendor"], t["category"], t["driver"], t.get("subcat", "")])
            t["vendor"], t["category"] = hit
            t["driver"] = f"{hit[0]} (vendor)" if hit[1] == "Uncategorized" else hit[1]
            t["subcat"] = med_type(hit[0], t["desc"]) if hit[1] == MEDICAL else ""
        elif "ov_orig" in t:
            t["vendor"], t["category"], t["driver"], t["subcat"] = t.pop("ov_orig")


def assign_ssa_owners(ledger):
    """With SSA_OWNERS set, two Social Security deposits arrive each month: the larger belongs to the first name."""
    ledger["review"] = [r for r in ledger["review"] if not r.get("line", "").startswith("Social Security:")]
    if not SSA_OWNERS:
        return
    by = defaultdict(list)
    for t in ledger["bank_txns"]:
        if t["vendor"] == SSA_VENDOR:
            by[(t["year"], t["month"])].append(t)
    for (yr, mth), items in sorted(by.items()):
        items.sort(key=lambda t: -t["amount"])
        for i, t in enumerate(items):
            t["cat"] = INC_SSA1 if i == 0 else INC_SSA2
        if len(items) != 2:
            ledger["review"].append({"source": items[0]["source"], "line": (
                f"Social Security: {len(items)} deposit(s) in {MONTHS[mth - 1]} {yr} (expected 2: {SSA_OWNERS[0]} larger, "
                f"{SSA_OWNERS[1]} smaller). The larger one is counted as {SSA_OWNERS[0]}'s.")})


def bank_review_items(ledger):
    ledger["review"] = [r for r in ledger["review"] if not r.get("line", "").startswith("Bank:")]
    tx = ledger["bank_txns"]
    if not tx:
        return
    src = tx[0]["source"]
    chk = [t for t in tx if t["cat"] == CHECKS_CAT]
    if chk:
        ledger["review"].append({"source": src, "line": (
            f"Bank: {len(chk)} written check(s) totalling ${-sum(t['amount'] for t in chk):,.2f} show only a check number "
            f"(the statement has no payee) and have no entry in check_payees.json. They are counted as expenses under "
            f"'{CHECKS_CAT}'.")})
    unc = [t for t in tx if t["flow"] == "Expense" and t["cat"] == "Uncategorized"]
    if unc:
        ledger["review"].append({"source": src, "line": (
            f"Bank: {len(unc)} payment(s) totalling ${-sum(t['amount'] for t in unc):,.2f} could not be categorized "
            f"(merchant name missing or unfamiliar). Filter Bank Transactions on Category = Uncategorized.")})
    oth = [t for t in tx if t["flow"] == "Income" and t["cat"] == INC_OTHER]
    for t in oth:
        ledger["review"].append({"source": t["source"], "line": (
            f"Bank: unclassified deposit ${t['amount']:,.2f} on {t['date']} counted as 'Other income': {t['desc'][:80]}")})
    big = [t for t in tx if t["cat"] == INC_CHECKS and t["amount"] >= 5000]
    for t in big:
        ledger["review"].append({"source": t["source"], "line": (
            f"Bank: large deposit ${t['amount']:,.2f} on {t['date']} counted as 'Check & branch deposits' "
            f"(the statement gives no source).")})
    for yr, items in BROKER_YEAREND.items():
        got = round(sum(t["amount"] for t in tx if t["year"] == yr and t["cat"] == INC_BROKER), 2)
        said = round(sum(a for _, a in items), 2)
        if got and abs(got - said) > 1:
            ledger["review"].append({"source": src, "line": (
                f"Bank: {BANK_NAME} shows ${got:,.2f} received from {BROKER} in {yr}, but {BROKER}'s year-end figures total "
                f"${said:,.2f} withdrawn ({'; '.join(f'{l} ${a:,.0f}' for l, a in items)}). "
                f"Difference ${got - said:,.2f}; see the Transfers tab. "
                + (f"Possible year-end timing: a {BROKER} withdrawal dated in the last days of the year may post at {BANK_NAME} "
                   f"in January. Check {BROKER} history for withdrawals dated 12/26-12/31 and the next January statement."
                   if got < said else
                   "Possible prior-year timing: early-January deposits may be withdrawals made the previous December."))})
    internal = round(sum(t["amount"] for t in tx if t["cat"] == TR_INTERNAL), 2)
    if abs(internal) > 0.005:
        ledger["review"].append({"source": src, "line": (
            f"Bank: transfers between {BANK_NAME} accounts do not net to zero (${internal:,.2f}); a matching leg may be "
            f"missing from a statement.")})


def process_bank_pdf(ledger, pdf, h):
    """Read one credit-union statement. Returns True when it was recorded."""
    try:
        parsed = parse_bank_statement(pdf)
    except Exception as exc:
        print(f"ERROR reading {pdf.name}: {exc}. Skipped, not recorded; fix or remove the file.")
        return False
    if not parsed or not parsed["accts"]:
        print(f"Could not read the statement period or accounts, NOT recorded (will retry next run): {pdf.name}")
        return False
    key = f"{BANK_NAME}|{parsed['end'].isoformat()}"
    entry = {"name": pdf.name, "card": f"{BANK_NAME} (bank)", "period_end": parsed["end"].strftime("%m/%d/%Y"),
             "rows": 0, "processed_at": datetime.now().isoformat(timespec="seconds")}
    if key in ledger["bank_stmts"]:
        entry["status"] = "Duplicate statement (same period); not counted"
        print(f"Duplicate statement period, not counted: {pdf.name}")
    else:
        txns = make_bank_txns(parsed, pdf.name)
        recons = [bank_recon_row(a, parsed, pdf.name) for a in parsed["accts"]]
        ledger["bank_txns"].extend(txns)
        ledger["bank_recon"].extend(recons)
        ledger["bank_stmts"][key] = h
        bad = [r for r in recons if r["status"] != "OK"]
        for r in bad:
            ledger["review"].append({"source": pdf.name, "line": f"{r['acct']}: {r['status']}"})
        entry["rows"] = len(txns)
        entry["status"] = "Processed - ties to statement" if not bad else "Processed - " + "; ".join(r["status"] for r in bad)
        print(f"Processed: {pdf.name} -> {BANK_NAME}, period ends {entry['period_end']}, {len(txns)} rows, "
              f"{'ties to statement' if not bad else 'DOES NOT TIE: ' + bad[0]['status']}")
    ledger["files"][h] = entry
    return True


# ----------------------------------------------------------------------------
# Workbook
# ----------------------------------------------------------------------------
FONT = "Arial"
HDR_FILL = PatternFill("solid", fgColor="1F3864")
SUB_FILL = PatternFill("solid", fgColor="D9E1F2")
WARN_FILL = PatternFill("solid", fgColor="FFF2CC")
MONEY = '$#,##0.00;($#,##0.00);"-"'
PCT = "0.0%"
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
TX_COLS = ["Stmt Month", "Stmt Year", "Card", "Type", "Vendor", "Category", "Pareto Driver", "Amount",
           "Spend Amount", "Breach", "Breach Note", "Trans Date", "Post Date", "Statement End",
           "Source File", "Description", "Sub-Category", "Daily Cash (memo)"]
TX = {"month": "A", "year": "B", "card": "C", "type": "D", "vendor": "E", "category": "F", "driver": "G",
      "amount": "H", "spend": "I", "breach": "J", "subcat": "Q", "dcash": "R"}


def net(t):
    return t["amount"] if t["type"] == PURCHASE else (-t["amount"] if t["type"] == REFUND else 0.0)


def style_header(ws, row, c1, c2):
    for c in range(c1, c2 + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(name=FONT, bold=True, color="FFFFFF", size=10)
        cell.fill = HDR_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def put(ws, ref, value, bold=False, fmt=None, fill=None, italic=False, size=10):
    c = ws[ref]
    c.value = value
    c.font = Font(name=FONT, size=size, bold=bold, italic=italic)
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    return c


class Builder:
    def __init__(self, last_row):
        self.last = max(last_row, 2)

    def rng(self, key):
        col = TX[key]
        return f"Transactions!${col}$2:${col}${self.last}"

    def sumifs(self, sum_key, crits):
        parts = [self.rng(sum_key)]
        for key, val in crits:
            parts += [self.rng(key), val]
        return "SUMIFS(" + ",".join(parts) + ")"


def write_summary(ws, bld, year, card, scope):
    ws.sheet_view.showGridLines = False
    put(ws, "A1", f"{card or 'Grand Totals - All Cards'} | {ylabel(year)}", bold=True, size=14)
    put(ws, "A2", "Year", bold=True); put(ws, "B2", year)
    put(ws, "A3", "Card", bold=True); put(ws, "B3", card or "All cards")
    last = max((t["stmt_end"] for t in scope), default="")
    put(ws, "A4", "Data thru", bold=True)
    put(ws, "B4", datetime.strptime(last, "%Y-%m-%d").strftime("%m/%d/%Y") if last else "n/a")
    put(ws, "A5", "Each card statement counts in the month it closes, so every month ties to the bank's own "
                  "figures (see Reconciliation tab); an Excel/CSV list counts in the month of each payment date. "
                  "Spend = purchases minus refunds/credits.",
        italic=True, size=9)

    def crit(extra):
        c = [("year", "$B$2")] + extra
        if card:
            c.append(("card", "$B$3"))
        return c

    KF = 7           # key figures block
    hdr = 18         # monthly table header row
    # ---- monthly table
    heads = ["Mo", "Month", "Purchases", "Refunds/Credits", "Net Spend", "Payments", "Interest", "Fees",
             "Cumulative Net Spend"]
    for i, h in enumerate(heads, start=1):
        ws.cell(row=hdr, column=i, value=h)
    style_header(ws, hdr, 1, len(heads))
    for m in range(1, 13):
        r = hdr + m
        put(ws, f"A{r}", m)
        put(ws, f"B{r}", MONTHS[m - 1])
        for col, typ in (("C", PURCHASE), ("D", REFUND), ("F", PAYMENT), ("G", INTEREST), ("H", FEE)):
            put(ws, f"{col}{r}", "=" + bld.sumifs("amount", crit([("month", f"$A{r}"), ("type", f'"{typ}"')])),
                fmt=MONEY)
        put(ws, f"E{r}", f"=C{r}-D{r}", fmt=MONEY)
        put(ws, f"I{r}", f"=SUM(E${hdr + 1}:E{r})", fmt=MONEY)
    tot = hdr + 13
    put(ws, f"B{tot}", "Total", bold=True, fill=SUB_FILL)
    for col in "CDEFGH":
        put(ws, f"{col}{tot}", f"=SUM({col}{hdr + 1}:{col}{hdr + 12})", bold=True, fmt=MONEY, fill=SUB_FILL)
    total_ref = f"$E${tot}"
    row = tot + 3

    # ---- by-card table (grand totals only)
    if card is None:
        cards = sorted({t["card"] for t in scope})
        put(ws, f"A{row}", "By card", bold=True)
        row += 1
        for i, h in enumerate(["", "Card", "Purchases", "Refunds/Credits", "Net Spend", "Payments",
                               "Interest", "Fees"], start=1):
            ws.cell(row=row, column=i, value=h)
        style_header(ws, row, 1, 8)
        first = row + 1
        for c in cards:
            row += 1
            put(ws, f"B{row}", c)
            for col, typ in (("C", PURCHASE), ("D", REFUND), ("F", PAYMENT), ("G", INTEREST), ("H", FEE)):
                put(ws, f"{col}{row}", "=" + bld.sumifs(
                    "amount", [("year", "$B$2"), ("card", f"$B{row}"), ("type", f'"{typ}"')]), fmt=MONEY)
            put(ws, f"E{row}", f"=C{row}-D{row}", fmt=MONEY)
        row += 1
        put(ws, f"B{row}", "Total", bold=True, fill=SUB_FILL)
        for col in "CDEFGH":
            put(ws, f"{col}{row}", f"=SUM({col}{first}:{col}{row - 1})", bold=True, fmt=MONEY, fill=SUB_FILL)
        row += 3

    # ---- Pareto helpers
    def totals_by(key, positive_only=False):
        d = defaultdict(float)
        for t in scope:
            d[t[key]] += net(t)
        keep = (lambda v: v > 0.005) if positive_only else (lambda v: abs(v) > 0.005)
        return sorted(((k, v) for k, v in d.items() if keep(v)), key=lambda kv: -kv[1])

    def pareto(title, key, basis_fn, extra_head=None, extra_fn=None):
        nonlocal row
        put(ws, f"A{row}", title, bold=True)
        row += 1
        heads = ["#", "Item", "Basis" if key == "driver" else "Category", "Net Spend", "% of Spend",
                 "Cumulative %", "In top 80%"]
        if extra_head:
            heads.append(extra_head)
        for i, h in enumerate(heads, start=1):
            ws.cell(row=row, column=i, value=h)
        style_header(ws, row, 1, len(heads))
        head_row = row
        positives = totals_by(key, positive_only=True)
        items = positives[:TOP_N]
        n = len(items)
        base = round(sum(v for _, v in positives), 2)   # spend before credits: the 100% line
        base_row = head_row + n + 2
        for rank, (label, _) in enumerate(items, start=1):
            row += 1
            put(ws, f"A{row}", rank)
            put(ws, f"B{row}", label)
            put(ws, f"C{row}", basis_fn(label))
            put(ws, f"D{row}", "=" + bld.sumifs("spend", crit([(key, f"$B{row}")])), fmt=MONEY)
            put(ws, f"E{row}", f"=IF($D${base_row}=0,0,D{row}/$D${base_row})", fmt=PCT)
            put(ws, f"F{row}", f"=E{row}" if rank == 1 else f"=F{row - 1}+E{row}", fmt=PCT)
            put(ws, f"G{row}", f'=IF(F{row}-E{row}<0.8,"Yes","")')
            if extra_fn:
                put(ws, f"H{row}", extra_fn(label))
        if n:
            row += 1
            put(ws, f"B{row}", "All other items")
            put(ws, f"D{row}", f"=D{base_row}-SUM(D{head_row + 1}:D{head_row + n})", fmt=MONEY)
            put(ws, f"E{row}", f"=IF($D${base_row}=0,0,D{row}/$D${base_row})", fmt=PCT)
            put(ws, f"F{row}", f"=F{row - 1}+E{row}", fmt=PCT)
            row += 1
            put(ws, f"B{row}", "Spend before credits (100% base)", bold=True, fill=SUB_FILL)
            put(ws, f"D{row}", base, bold=True, fmt=MONEY, fill=SUB_FILL)
            row += 1
            put(ws, f"B{row}", "Credits, refunds and rewards netted")
            put(ws, f"D{row}", f"={total_ref}-D{base_row}", fmt=MONEY)
            row += 1
            put(ws, f"B{row}", "Total net spend", bold=True, fill=SUB_FILL)
            put(ws, f"D{row}", f"=D{base_row}+D{row - 1}", bold=True, fmt=MONEY, fill=SUB_FILL)
        row += 3
        return head_row, n

    drv_basis = {t["driver"]: ("Vendor (no category)" if t["driver"].endswith("(vendor)") else "Category")
                 for t in scope}
    p_head, p_n = pareto("Pareto - where the money goes (by category; by vendor where no category could be derived)",
                         "driver", lambda k: drv_basis.get(k, ""))
    vend_cat = {}
    for t in scope:
        vend_cat.setdefault(t["vendor"], t["category"])
    breach_vendors = {t["vendor"]: t["breach_note"] for t in scope if t["breach"] == "Y"}
    v_head, v_n = pareto("Pareto - top vendors", "vendor", lambda k: vend_cat.get(k, ""),
                         extra_head="Breach/Fraud", extra_fn=lambda k: "Y" if k in breach_vendors else "")

    # ---- breach exposure
    put(ws, f"A{row}", "Spend with breach/fraud-flagged vendors", bold=True)
    row += 1
    for i, h in enumerate(["", "Vendor", "Net Spend", "Note"], start=1):
        ws.cell(row=row, column=i, value=h)
    style_header(ws, row, 1, 4)
    if breach_vendors:
        first = row + 1
        for v, note in sorted(breach_vendors.items()):
            row += 1
            put(ws, f"B{row}", v)
            put(ws, f"D{row}", note)
            put(ws, f"C{row}", "=" + bld.sumifs("spend", crit([("vendor", f"$B{row}"), ("breach", '"Y"')])), fmt=MONEY)
        row += 1
        put(ws, f"B{row}", "Total", bold=True, fill=SUB_FILL)
        put(ws, f"C{row}", f"=SUM(C{first}:C{row - 1})", bold=True, fmt=MONEY, fill=SUB_FILL)
    else:
        row += 1
        put(ws, f"B{row}", "None flagged. Add patterns to BREACH_VENDORS in the script.")
    row += 3

    # ---- spend by category by month
    cats = [k for k, _ in totals_by("category")]
    if cats:
        put(ws, f"A{row}", "Net spend by category and month", bold=True)
        row += 1
        mo_row = row
        for m in range(1, 13):
            c = put(ws, f"{get_column_letter(m + 2)}{row}", m, size=8)
            c.font = Font(name=FONT, size=8, color="808080")
        row += 1
        mh = row
        put(ws, f"B{row}", "Category")
        for m in range(1, 13):
            ws.cell(row=row, column=m + 2, value=MONTHS[m - 1])
        ws.cell(row=row, column=15, value="Total")
        style_header(ws, row, 2, 15)
        first = row + 1
        for cat in cats:
            row += 1
            put(ws, f"B{row}", cat)
            for m in range(1, 13):
                col = get_column_letter(m + 2)
                put(ws, f"{col}{row}", "=" + bld.sumifs("spend", crit([("category", f"$B{row}"),
                                                                       ("month", f"{col}${mo_row}")])), fmt=MONEY)
            put(ws, f"O{row}", f"=SUM(C{row}:N{row})", fmt=MONEY, bold=True)
        row += 1
        put(ws, f"B{row}", "Total", bold=True, fill=SUB_FILL)
        for m in range(1, 14):
            col = get_column_letter(m + 2)
            put(ws, f"{col}{row}", f"=SUM({col}{first}:{col}{row - 1})", bold=True, fmt=MONEY, fill=SUB_FILL)
        row += 2

    # ---- medical and dental by type
    sub_tot = defaultdict(float)
    for t in scope:
        if t["category"] == MEDICAL and t.get("subcat"):
            sub_tot[t["subcat"]] += net(t)
    if sub_tot:
        put(ws, f"A{row}", "Medical & Dental by type", bold=True)
        row += 1
        mo2 = row
        for m in range(1, 13):
            c = put(ws, f"{get_column_letter(m + 2)}{row}", m, size=8)
            c.font = Font(name=FONT, size=8, color="808080")
        row += 1
        put(ws, f"B{row}", "Type")
        for m in range(1, 13):
            ws.cell(row=row, column=m + 2, value=MONTHS[m - 1])
        ws.cell(row=row, column=15, value="Total")
        style_header(ws, row, 2, 15)
        first = row + 1
        for sub, _ in sorted(sub_tot.items(), key=lambda kv: -kv[1]):
            row += 1
            put(ws, f"B{row}", sub)
            for m in range(1, 13):
                col = get_column_letter(m + 2)
                put(ws, f"{col}{row}", "=" + bld.sumifs("spend", crit([
                    ("category", f'"{MEDICAL}"'), ("subcat", f"$B{row}"), ("month", f"{col}${mo2}")])), fmt=MONEY)
            put(ws, f"O{row}", f"=SUM(C{row}:N{row})", fmt=MONEY, bold=True)
        row += 1
        put(ws, f"B{row}", "Total", bold=True, fill=SUB_FILL)
        for m in range(1, 14):
            col = get_column_letter(m + 2)
            put(ws, f"{col}{row}", f"=SUM({col}{first}:{col}{row - 1})", bold=True, fmt=MONEY, fill=SUB_FILL)
        row += 2

    # ---- key figures (filled last, now that table positions are known)
    put(ws, f"A{KF}", "Key figures", bold=True)
    kf = [
        ("Total net spend", f"={total_ref}", MONEY),
        ("Average per month with spend", f'=IFERROR({total_ref}/COUNTIF($C${hdr + 1}:$C${hdr + 12},">0"),0)', MONEY),
        ("Payments made", f"=$F${tot}", MONEY),
        ("Interest paid", f"=$G${tot}", MONEY),
        ("Fees paid", f"=$H${tot}", MONEY),
        ("Largest single purchase", "=_xlfn.MAXIFS(" + bld.rng("spend") + "," + bld.rng("year") + ",$B$2"
         + ("," + bld.rng("card") + ",$B$3" if card else "") + ")", MONEY),
        ("Biggest spend driver", f'=IF({p_n}=0,"n/a",B{p_head + 1})', None),
        ("Share of spend in biggest driver", f"=IF({p_n}=0,0,E{p_head + 1})", PCT),
        ("Items that make up 80% of spend", f'=COUNTIF($G${p_head + 1}:$G${p_head + max(p_n, 1)},"Yes")', "0"),
    ]
    for i, (lab, f, fmt) in enumerate(kf, start=1):
        put(ws, f"B{KF + i}", lab)
        put(ws, f"C{KF + i}", f, fmt=fmt, bold=True)
        ws[f"C{KF + i}"].alignment = Alignment(horizontal="left")

    if any(t.get("dcash") for t in scope):      # Daily Cash (Apple Card): a memo, never netted from spend
        put(ws, f"E{KF}", "Apple Card Daily Cash (memo)", bold=True)
        put(ws, f"E{KF + 1}", "Earned in the year, net of reversals")
        put(ws, f"H{KF + 1}", "=" + bld.sumifs("dcash", crit([])), fmt=MONEY, bold=True)
        put(ws, f"E{KF + 2}", "Paid into Apple Cash, not credited to the card;", italic=True, size=9)
        put(ws, f"E{KF + 3}", "spend above is gross and does not net it.", italic=True, size=9)
        ws.cell(row=hdr, column=10, value="Daily Cash (memo)")
        style_header(ws, hdr, 10, 10)
        for m in range(1, 13):
            put(ws, f"J{hdr + m}", "=" + bld.sumifs("dcash", crit([("month", f"$A{hdr + m}")])), fmt=MONEY)
        put(ws, f"J{tot}", f"=SUM(J{hdr + 1}:J{hdr + 12})", bold=True, fmt=MONEY, fill=SUB_FILL)

    for col, w in zip("ABCDEFGHI", (11, 34, 24, 16, 16, 16, 14, 14, 20)):
        ws.column_dimensions[col].width = w
    for col in "JKLMNO":
        ws.column_dimensions[col].width = 13
    ws.freeze_panes = "A6"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

    bar = BarChart()
    bar.title = "Net Spend by Month"
    bar.add_data(Reference(ws, min_col=5, min_row=hdr, max_row=hdr + 12), titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=2, min_row=hdr + 1, max_row=hdr + 12))
    bar.legend = None
    bar.x_axis.delete = False
    bar.y_axis.delete = False
    bar.height, bar.width = 8.5, 16
    polish_chart(bar)
    ws.add_chart(bar, "K7")

    if p_n:
        k = min(p_n, 10)
        pb = BarChart()
        pb.title = "Pareto - Top Spend Drivers"
        pb.add_data(Reference(ws, min_col=4, min_row=p_head, max_row=p_head + k), titles_from_data=True)
        pb.set_categories(Reference(ws, min_col=2, min_row=p_head + 1, max_row=p_head + k))
        pb.x_axis.delete = False
        pb.y_axis.delete = False
        pb.legend = None
        ln = LineChart()
        ln.add_data(Reference(ws, min_col=6, min_row=p_head, max_row=p_head + k), titles_from_data=True)
        ln.y_axis.axId = 200
        ln.y_axis.txPr = _txt(900)
        ln.y_axis.number_format = "0%"
        ln.y_axis.scaling.max = 1
        ln.y_axis.scaling.min = 0
        ln.y_axis.crosses = "max"
        ln.y_axis.delete = False
        pb += ln
        pb.height, pb.width = 10, 16
        polish_chart(pb, right=0.11, rotate=True)
        ws.add_chart(pb, f"K{max(p_head - 1, 24)}")
    return {"hdr": hdr, "p_head": p_head, "p_n": p_n, "v_head": v_head, "v_n": v_n}


def _txt(size, bold=False):
    from openpyxl.chart.text import RichText
    from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties
    cp = CharacterProperties(sz=size, b=bold)
    return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)])


def polish_chart(ch, horizontal=False, right=0.04, rotate=False):
    """Explicit layout so Excel never draws the title, legend and axis labels on top of one another."""
    from openpyxl.chart.layout import Layout, ManualLayout
    from openpyxl.drawing.text import CharacterProperties, ParagraphProperties
    if ch.title is not None:
        ch.title.overlay = False
        cp = CharacterProperties(sz=1200, b=True)
        for para in ch.title.tx.rich.p:
            para.pPr = ParagraphProperties(defRPr=cp)
            for run in para.r or []:
                run.rPr = cp
    ch.x_axis.txPr = _txt(900)
    ch.y_axis.txPr = _txt(900)
    ch.y_axis.number_format = KFMT
    if rotate:
        ch.x_axis.txPr.bodyPr.rot = -2700000
        ch.x_axis.txPr.bodyPr.vert = "horz"
    has_legend = ch.legend is not None
    if has_legend:
        ch.legend.overlay = False
        ch.legend.position = "b"
        ch.legend.txPr = _txt(900)
    left = 0.30 if horizontal else 0.11
    ch.layout = Layout(manualLayout=ManualLayout(layoutTarget="inner", xMode="edge", yMode="edge",
                                                 x=left, y=0.17, w=1 - right - left,
                                                 h=(0.50 if rotate else 0.56) if has_legend else (0.48 if rotate else 0.68)))


def write_charts(ws, bld, grand, pos, year, cards):
    """Dedicated Charts tab. Reads the Grand Totals tab and the Transactions tab."""
    ws.sheet_view.showGridLines = False
    put(ws, "A1", f"Charts | {ylabel(year)}", bold=True, size=14)
    put(ws, "A2", "All charts update from the Grand Totals and Transactions tabs.", italic=True, size=9)
    hdr = pos["hdr"]
    # helper table: net spend by month for each card
    put(ws, "A4", "Mo", bold=True)
    heads = ["Mo", "Month"] + list(cards)
    for i, h in enumerate(heads, start=1):
        ws.cell(row=4, column=i, value=h)
    style_header(ws, 4, 1, len(heads))
    for m in range(1, 13):
        r = 4 + m
        put(ws, f"A{r}", m)
        put(ws, f"B{r}", MONTHS[m - 1])
        for ci, _ in enumerate(cards):
            col = get_column_letter(3 + ci)
            put(ws, f"{col}{r}", "=" + bld.sumifs("spend", [("year", "'Grand Totals'!$B$2"), ("month", f"$A{r}"),
                                                           ("card", f"{col}$4")]), fmt=MONEY)
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 12
    for ci in range(len(cards)):
        ws.column_dimensions[get_column_letter(3 + ci)].width = 20

    def style(ch, title, w=17, h=8.5):
        ch.title = title
        ch.x_axis.delete = False
        ch.y_axis.delete = False
        ch.width, ch.height = w, h
        return ch

    cats_g = Reference(grand, min_col=2, min_row=hdr + 1, max_row=hdr + 12)

    c1 = style(BarChart(), "Net Spend by Month and Card")
    c1.type = "col"; c1.grouping = "stacked"; c1.overlap = 100
    c1.add_data(Reference(ws, min_col=3, max_col=2 + len(cards), min_row=4, max_row=16), titles_from_data=True)
    c1.set_categories(Reference(ws, min_col=2, min_row=5, max_row=16))
    polish_chart(c1)
    ws.add_chart(c1, "A19")

    c2 = style(BarChart(), "Net Spend vs Payments by Month")
    c2.add_data(Reference(grand, min_col=5, max_col=6, min_row=hdr, max_row=hdr + 12), titles_from_data=True)
    c2.set_categories(cats_g)
    polish_chart(c2)
    ws.add_chart(c2, "J19")

    c3 = style(LineChart(), "Cumulative Net Spend")
    c3.add_data(Reference(grand, min_col=9, min_row=hdr, max_row=hdr + 12), titles_from_data=True)
    c3.set_categories(cats_g)
    c3.legend = None
    for sr in c3.series:
        sr.smooth = False
    polish_chart(c3)
    ws.add_chart(c3, "A37")

    if pos["p_n"]:
        ph, k = pos["p_head"], min(pos["p_n"], 10)
        pb = style(BarChart(), "Pareto - Top Spend Drivers")
        pb.add_data(Reference(grand, min_col=4, min_row=ph, max_row=ph + k), titles_from_data=True)
        pb.set_categories(Reference(grand, min_col=2, min_row=ph + 1, max_row=ph + k))
        pb.legend = None
        ln = LineChart()
        ln.add_data(Reference(grand, min_col=6, min_row=ph, max_row=ph + k), titles_from_data=True)
        ln.y_axis.axId = 200
        ln.y_axis.txPr = _txt(900)
        ln.y_axis.number_format = "0%"
        ln.y_axis.scaling.max = 1
        ln.y_axis.scaling.min = 0
        ln.y_axis.crosses = "max"
        ln.y_axis.delete = False
        pb += ln
        pb.height = 10
        polish_chart(pb, right=0.11, rotate=True)
        ws.add_chart(pb, "J37")

    if pos["v_n"]:
        vh, k = pos["v_head"], min(pos["v_n"], 10)
        vb = style(BarChart(), "Top 10 Vendors by Net Spend", h=10)
        vb.type = "bar"
        vb.add_data(Reference(grand, min_col=4, min_row=vh, max_row=vh + k), titles_from_data=True)
        vb.set_categories(Reference(grand, min_col=2, min_row=vh + 1, max_row=vh + k))
        vb.x_axis.scaling.orientation = "maxMin"
        vb.y_axis.crosses = "max"
        vb.legend = None
        polish_chart(vb, horizontal=True)
        ws.add_chart(vb, "A55")
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)


# ----------------------------------------------------------------------------
# Workbook: cash-flow tabs (income vs expenses), transfers, bank detail and bank reconciliation
# ----------------------------------------------------------------------------
BT_COLS = ["Month", "Year", "Account", "Date", "Flow", "Category / Source", "Payee / Vendor", "Amount (+in / -out)",
           "Balance After", "Income Amt", "Expense Amt", "Transfer Amt", "Description", "Source File", "Note"]
BR_COLS = ["Account", "Month", "Year", "Statement End", "Source File", "Opening Balance", "Closing Balance",
           "Net of Rows Read", "Difference", "Deposits (read)", "Deposits (statement)", "Withdrawals (read)",
           "Withdrawals (statement)", "Balance-Chain Breaks", "Rows", "Status"]
MCOLS = [get_column_letter(c) for c in range(2, 14)]      # B..M = Jan..Dec


def _bt(col, last):
    return f"'Bank Transactions'!${col}$2:${col}${last}"


def _br(col, last):
    return f"'Bank Reconciliation'!${col}$2:${col}${last}"


def _bank_sumifs(sum_col, last, extra):
    parts = [_bt(sum_col, last), _bt("B", last), "$B$3"]
    for col, crit in extra:
        parts += [_bt(col, last), crit]
    return "SUMIFS(" + ",".join(parts) + ")"


def write_bank_transactions(ws, txns):
    for i, h in enumerate(BT_COLS, start=1):
        ws.cell(row=1, column=i, value=h)
    style_header(ws, 1, 1, len(BT_COLS))
    ws.row_dimensions[1].height = 30
    for r, t in enumerate(txns, start=2):
        vals = [t["month"], t["year"], t["acct"], datetime.strptime(t["date"], "%Y-%m-%d").date(), t["flow"], t["cat"],
                t["vendor"], t["amount"], t["balance"], f'=IF(E{r}="Income",H{r},0)', f'=IF(E{r}="Expense",-H{r},0)',
                f'=IF(E{r}="Transfer",H{r},0)', t["desc"], t["source"], t.get("note", "")]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = Font(name=FONT, size=10)
            if c == 4:
                cell.number_format = "mm/dd/yyyy"
            elif c in (8, 9, 10, 11, 12):
                cell.number_format = MONEY
    for i, w in enumerate((8, 7, 30, 12, 10, 40, 34, 16, 14, 13, 13, 13, 70, 30, 30), start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(BT_COLS))}{max(len(txns) + 1, 2)}"


def write_bank_recon(ws, ledger, year, bt_last):
    for i, h in enumerate(BR_COLS, start=1):
        ws.cell(row=1, column=i, value=h)
    style_header(ws, 1, 1, len(BR_COLS))
    ws.row_dimensions[1].height = 42
    rows = sorted(ledger["bank_recon"], key=lambda x: (x["stmt_end"], x["acct"]))
    d = lambda s: datetime.strptime(s, "%Y-%m-%d").date()
    for r, rc in enumerate(rows, start=2):
        vals = [rc["acct"], rc["month"], rc["year"], d(rc["stmt_end"]), rc["source"], rc["open"], rc["close"],
                rc["rows_sum"], f"=ROUND(F{r}+H{r}-G{r},2)", rc["dep_read"], rc["dep_stmt"], rc["wd_read"],
                rc["wd_stmt"], rc["chain"], rc["n"], rc["status"]]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = Font(name=FONT, size=10)
            if c == 4:
                cell.number_format = "mm/dd/yyyy"
            elif 6 <= c <= 13:
                cell.number_format = MONEY
        if rc["status"] != "OK":
            for c in range(1, len(BR_COLS) + 1):
                ws.cell(row=r, column=c).fill = WARN_FILL
    n = len(rows) + 1
    put(ws, f"A{n + 2}", "Statements that tie (opening + rows = closing, running balance, bank's own totals):", bold=True)
    put(ws, f"F{n + 2}", f'=COUNTIF(P2:P{n},"OK")&" of "&COUNTA(P2:P{n})', bold=True)
    put(ws, f"A{n + 3}", "Deposits and withdrawals are classified the way the credit union does it (a positive "
                         "'Withdrawal Adjustment' / debit-card credit voucher counts against withdrawals). A blank "
                         "'(statement)' cell means that figure was not legible on the redacted PDF; the balance "
                         "checks still apply.", italic=True, size=9)
    # cross-checks
    r0 = n + 5
    put(ws, f"A{r0}", f"Cross-checks, {year}", bold=True)
    ref = next((v for k, v in ledger.get("reference", {}).items() if k in REFERENCE_ACCOUNTS), None)
    if ref:
        put(ws, f"A{r0 + 1}", "Your manual medical list (reference only, not counted)")
        put(ws, f"F{r0 + 1}", ref["total"], fmt=MONEY)
        put(ws, f"A{r0 + 2}", f"Medical & Dental found on the {BANK_NAME} statements")
        put(ws, f"F{r0 + 2}", "=" + _bank_sumifs("K", bt_last, [("E", '"Expense"'), ("F", '"Medical & Dental"')]).replace("$B$3", str(year)),
            fmt=MONEY)
        put(ws, f"A{r0 + 3}", "Difference (list minus statements)")
        put(ws, f"F{r0 + 3}", f"=ROUND(F{r0 + 1}-F{r0 + 2},2)", fmt=MONEY)
        put(ws, f"G{r0 + 1}", "The statements also include medical items paid by check or on other accounts, and the "
                              "list may include items paid by card; small differences are normal.", italic=True, size=9)
    def card_xcheck(r, name, prefix, kind, show_diff, settled_only=False):
        if settled_only:    # store card: only payments that settle a loaded statement are transfers
            got = round(sum(t["amount"] for t in ledger["transactions"] if t["card"].startswith(prefix)
                            and t["type"] == PAYMENT and t.get("settles_loaded") and t["date"][:4] == str(year)), 2)
        else:
            got = round(sum(rc["s_paycred"] for rc in ledger["recon"] if rc.get("s_paycred") is not None
                            and rc["card"].startswith(prefix) and str(rc["stmt_end"])[:4] == str(year)), 2)
        if not got and not any(t["cat"] == kind for t in ledger["bank_txns"]):
            return False
        put(ws, f"A{r}", f"Payments shown on the {name} statements that settle a loaded statement (dated in year)"
            if settled_only else f"Payments shown on the {name} statements (statement months in year)")
        put(ws, f"F{r}", got, fmt=MONEY)
        put(ws, f"A{r + 1}", f"Net paid to {name} from {BANK_NAME} in the year (payments less any return)")
        put(ws, f"F{r + 1}", "=-" + _bank_sumifs("L", bt_last, [("F", f'"{kind}"')]).replace("$B$3", str(year)), fmt=MONEY)
        if show_diff:
            put(ws, f"A{r + 2}", f"Difference (statements minus {BANK_NAME})")
            put(ws, f"F{r + 2}", f"=ROUND(F{r}-F{r + 1},2)", fmt=MONEY)
        return True

    r1 = r0 + 5
    if card_xcheck(r1, "Citi", "Citi", TR_CITI, False):
        put(ws, f"G{r1}", "These differ by timing (a December payment is on the January statement) and by "
                          "credits/refunds that are not payments.", italic=True, size=9)
        r1 += 3
    if card_xcheck(r1, "Apple Card", APPLE_CARD, TR_APPLE, True):
        put(ws, f"G{r1}", f"Every Apple Card statement payment is matched to a {BANK_NAME} withdrawal (same amount, within "
                          "a few days), so these should agree; the Review tab lists any that do not.",
            italic=True, size=9)
        r1 += 3
    if card_xcheck(r1, "Amazon Store Card", "Amazon Store Card", TR_AMAZON, True, settled_only=True):
        put(ws, f"G{r1}", f"A payment on a store-card statement settles the previous statement, so it is a transfer only "
                          f"when that previous statement is loaded. Other payments stay expenses ('{OTHER_CARDS}') until "
                          "the earlier statements are loaded; the Review tab lists them.", italic=True, size=9)
    for i, w in enumerate((34, 8, 7, 14, 32, 16, 16, 16, 12, 16, 16, 16, 16, 12, 8, 70), start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    return n


def card_group_label(ledger, year):
    fams = sorted({re.sub(r"\s+Card$", "", re.sub(r"\s+\d{4}$", "", t["card"]))
                   for t in ledger["transactions"] if t["year"] == year})
    label = "Cards: " + ", ".join(fams) + " - net spend by month"
    return label if len(label) <= 64 else f"Credit cards ({len(fams)}) - net spend by statement month"


def write_cashflow(ws, ledger, year, bt_last, card_bld):
    """Income vs expenses by month. Returns row positions for the charts."""
    ws.sheet_view.showGridLines = False
    tx = [t for t in ledger["bank_txns"] if t["year"] == year]
    put(ws, "A1", f"Cash Flow - Income vs Expenses | {ylabel(year)}", bold=True, size=14)
    put(ws, "A2", f"Income and expenses paid from {BANK_NAME} (all accounts) plus credit card spend (every card with statements loaded). Transfers "
                  "are shown on the Transfers tab and are not income or expense; payments to cards whose statements are "
                  "loaded are left out of expenses because the card purchases themselves are counted instead.",
        italic=True, size=9)
    put(ws, "A3", "Year", bold=True)
    put(ws, "B3", year)
    put(ws, "A4", "Month #", size=8, italic=True)
    for m, col in enumerate(MCOLS, start=1):
        put(ws, f"{col}4", m, size=8, italic=True)
    heads = ["Item"] + MONTHS + ["Total", "% of section"]
    for i, h in enumerate(heads, start=1):
        ws.cell(row=5, column=i, value=h)
    style_header(ws, 5, 1, 15)

    def section(r, text):
        for c in range(1, 16):
            ws.cell(row=r, column=c).fill = SUB_FILL
        put(ws, f"A{r}", text, bold=True, fill=SUB_FILL)

    def line(r, label, formula_fn, bold=False, pct_of=None):
        put(ws, f"A{r}", label, bold=bold)
        for col in MCOLS:
            put(ws, f"{col}{r}", "=" + formula_fn(col), fmt=MONEY, bold=bold)
        put(ws, f"N{r}", f"=SUM(B{r}:M{r})", fmt=MONEY, bold=True)
        if pct_of:
            put(ws, f"O{r}", f"=IF({pct_of}=0,0,N{r}/{pct_of})", fmt=PCT)

    pos = {"hdr": 5, "inc_rows": {}, "exp_rows": {}}
    section(6, "INCOME")
    r = 7
    pos["inc_first"] = r
    inc_total_row = 7 + len(INCOME_SOURCES)
    for src in INCOME_SOURCES:
        line(r, src, lambda col, r=r: _bank_sumifs("J", bt_last, [("A", f"{col}$4"), ("E", '"Income"'), ("F", f"$A{r}")]),
             pct_of=f"$N${inc_total_row}")
        pos["inc_rows"][src] = r
        if src == INC_BROKER:
            pos["broker_row"] = r
        if src == INC_CHECKS:
            pos["checks_row"] = r
        r += 1
    pos["inc_last"] = r - 1
    pos["inc_total"] = r
    line(r, "TOTAL INCOME", lambda col: f"SUM({col}{pos['inc_first']}:{col}{pos['inc_last']})", bold=True)
    for c in range(1, 16):
        ws.cell(row=r, column=c).fill = SUB_FILL

    r += 2
    section(r, f"EXPENSES - paid from {BANK_NAME} (by category)")
    r += 1
    totals = defaultdict(float)
    for t in tx:
        if t["flow"] == "Expense":
            totals[t["cat"]] -= t["amount"]
    cats = sorted(totals, key=lambda c: (-totals[c], c))
    pos["exp_first"] = r
    exp_total_row = r + len(cats) + 2
    for cat in cats:
        line(r, cat, lambda col, r=r: _bank_sumifs("K", bt_last, [("A", f"{col}$4"), ("E", '"Expense"'), ("F", f"$A{r}")]),
             pct_of=f"$N${exp_total_row}")
        pos["exp_rows"][cat] = r
        r += 1
    pos["exp_last_bank"] = r - 1
    line(r, card_group_label(ledger, year),
         lambda col: card_bld.sumifs("spend", [("year", "$B$3"), ("month", f"{col}$4")]), pct_of=f"$N${exp_total_row}")
    pos["citi_row"] = r
    r += 1
    line(r, "Card interest & fees (charged on the card statements)",
         lambda col: "+".join(card_bld.sumifs("amount", [("year", "$B$3"), ("month", f"{col}$4"), ("type", f'"{t}"')])
                              for t in (INTEREST, FEE)), pct_of=f"$N${exp_total_row}")
    pos["card_fee_row"] = r
    r += 1
    pos["exp_total"] = r
    line(r, "TOTAL EXPENSES", lambda col: f"SUM({col}{pos['exp_first']}:{col}{r - 1})", bold=True)
    for c in range(1, 16):
        ws.cell(row=r, column=c).fill = SUB_FILL

    r += 2
    pos["net"] = r
    line(r, "NET CASH FLOW (income - expenses)", lambda col: f"{col}{pos['inc_total']}-{col}{pos['exp_total']}", bold=True)
    r += 1
    pos["net_ex"] = r
    line(r, f"Net cash flow before {BROKER} transfers in", lambda col: f"{col}{pos['net']}-{col}{pos['broker_row']}")
    r += 1
    put(ws, f"A{r}", f"{BROKER} transfers in, % of total income", italic=True)
    put(ws, f"N{r}", f"=IF(N{pos['inc_total']}=0,0,N{pos['broker_row']}/N{pos['inc_total']})", fmt=PCT)
    r += 2
    owners = (f"{SSA_OWNERS[0]}'s Social Security is the larger of the two monthly deposits, {SSA_OWNERS[1]}'s the "
              f"smaller. ") if SSA_OWNERS else ""
    put(ws, f"A{r}", owners + "Check & branch deposits are net of returned checks. Mortgage extra principal is shown on "
                     "its own line. Checks show only a number on the statement; payees come from check_payees.json.",
        italic=True, size=9)
    ws.column_dimensions["A"].width = 62
    for col in MCOLS:
        ws.column_dimensions[col].width = 12
    ws.column_dimensions["N"].width = 14
    ws.column_dimensions["O"].width = 12
    ws.freeze_panes = "B6"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    return pos


def write_transfers(ws, ledger, year, bt_last, br_last, pos):
    ws.sheet_view.showGridLines = False
    put(ws, "A1", f"Transfers and Balance Tie-out | {ylabel(year)}", bold=True, size=14)
    put(ws, "A2", f"Money moved to or from {BANK_NAME} that is neither income nor expense. Negative = left {BANK_NAME}.",
        italic=True, size=9)
    put(ws, "A3", "Year", bold=True)
    put(ws, "B3", year)
    put(ws, "A4", "Month #", size=8, italic=True)
    for m, col in enumerate(MCOLS, start=1):
        put(ws, f"{col}4", m, size=8, italic=True)
    for i, h in enumerate(["Item"] + MONTHS + ["Total"], start=1):
        ws.cell(row=5, column=i, value=h)
    style_header(ws, 5, 1, 14)

    def row(r, label, fn, bold=False):
        put(ws, f"A{r}", label, bold=bold)
        for col in MCOLS:
            put(ws, f"{col}{r}", "=" + fn(col), fmt=MONEY, bold=bold)
        put(ws, f"N{r}", f"=SUM(B{r}:M{r})", fmt=MONEY, bold=True)

    r = 6
    for kind in TRANSFER_KINDS:
        row(r, kind, lambda col, r=r: _bank_sumifs("L", bt_last, [("A", f"{col}$4"), ("E", '"Transfer"'), ("F", f"$A{r}")]))
        r += 1
    t_first, t_last = 6, r - 1
    row(r, "Net transfers", lambda col: f"SUM({col}{t_first}:{col}{t_last})", bold=True)
    net_tr = r
    r += 2
    put(ws, f"A{r}", f"{BROKER} (IRA / brokerage)", bold=True, fill=SUB_FILL)
    for c in range(2, 15):
        ws.cell(row=r, column=c).fill = SUB_FILL
    r += 1
    sw_in = r
    row(r, f"Received from {BROKER} (counted as income)",
        lambda col: _bank_sumifs("J", bt_last, [("A", f"{col}$4"), ("E", '"Income"'), ("F", f'"{INC_BROKER}"')]))
    r += 1
    sw_out = r
    row(r, f"Sent back to {BROKER}",
        lambda col: "-" + _bank_sumifs("L", bt_last, [("A", f"{col}$4"), ("E", '"Transfer"'), ("F", f'"{TR_BROKER_OUT}"')]))
    r += 1
    row(r, f"Net from {BROKER}", lambda col: f"{col}{sw_in}-{col}{sw_out}", bold=True)
    r += 2
    put(ws, f"A{r}", f"Balance tie-out (all {BANK_NAME} deposit accounts)", bold=True, fill=SUB_FILL)
    for c in range(2, 15):
        ws.cell(row=r, column=c).fill = SUB_FILL
    r += 1
    op = r
    for col in MCOLS:
        put(ws, f"{col}{r}", "=SUMIFS(" + ",".join([_br("F", br_last), _br("B", br_last), f"{col}$4",
                                                  _br("C", br_last), "$B$3"]) + ")", fmt=MONEY)
    put(ws, f"A{r}", "Opening balance (statements)")
    put(ws, f"N{r}", f"=B{r}", fmt=MONEY, bold=True)
    r += 1
    inc = r
    row(r, "+ Income", lambda col: f"CashFlowRef_inc_{col}")
    r += 1
    exp = r
    row(r, f"- Expenses paid from {BANK_NAME}", lambda col: f"CashFlowRef_exp_{col}")
    r += 1
    trn = r
    row(r, "+ Net transfers", lambda col: f"{col}{net_tr}")
    r += 1
    calc = r
    put(ws, f"A{r}", "Computed closing balance", bold=True)
    for col in MCOLS:
        put(ws, f"{col}{r}", f"={col}{op}+{col}{inc}-{col}{exp}+{col}{trn}", fmt=MONEY, bold=True)
    put(ws, f"N{r}", f"=M{r}", fmt=MONEY, bold=True)
    r += 1
    stm = r
    put(ws, f"A{r}", "Closing balance (statements)")
    for col in MCOLS:
        put(ws, f"{col}{r}", "=SUMIFS(" + ",".join([_br("G", br_last), _br("B", br_last), f"{col}$4",
                                                  _br("C", br_last), "$B$3"]) + ")", fmt=MONEY)
    put(ws, f"N{r}", f"=M{r}", fmt=MONEY, bold=True)
    r += 1
    put(ws, f"A{r}", "Difference (must be zero)", bold=True)
    for col in MCOLS:
        put(ws, f"{col}{r}", f"=IF(AND({col}{op}=0,{col}{stm}=0),0,ROUND({col}{calc}-{col}{stm},2))", fmt=MONEY, bold=True)
    put(ws, f"N{r}", f"=SUM(B{r}:M{r})", fmt=MONEY, bold=True)
    pos["tie_diff"] = r
    pos["sw_in"], pos["sw_out"] = sw_in, sw_out
    # fill the references to the Cash Flow tab
    for col in MCOLS:
        ws[f"{col}{inc}"].value = f"='Cash Flow'!{col}{pos['inc_total']}"
        bank_exp = f"SUM('Cash Flow'!{col}{pos['exp_first']}:{col}{pos['exp_last_bank']})"
        ws[f"{col}{exp}"].value = f"={bank_exp}"
    put(ws, f"A{r + 2}", f"Expenses here are only those paid from {BANK_NAME} (card spend is excluded because payments to "
                         "cards with statements loaded are transfers). Opening and closing balances come from the statements.",
        italic=True, size=9)
    ye = BROKER_YEAREND.get(year)
    if ye:
        q = r + 5
        put(ws, f"A{q}", f"{BROKER} cross-check ({BROKER}'s year-end withdrawals vs {BANK_NAME} deposits)", bold=True, fill=SUB_FILL)
        for c in range(2, 15):
            ws.cell(row=q, column=c).fill = SUB_FILL
        for k, (label, amt) in enumerate(ye, start=1):
            put(ws, f"A{q + k}", f"{BROKER} year-end: {label}")
            put(ws, f"N{q + k}", amt, fmt=MONEY)
        e = q + len(ye)
        put(ws, f"A{e + 1}", f"Total withdrawn per {BROKER}", bold=True)
        put(ws, f"N{e + 1}", f"=SUM(N{q + 1}:N{e})", fmt=MONEY, bold=True)
        put(ws, f"A{e + 2}", f"Received from {BROKER} per {BANK_NAME} statements", bold=True)
        put(ws, f"N{e + 2}", f"=N{sw_in}", fmt=MONEY, bold=True)
        put(ws, f"A{e + 3}", f"Difference ({BANK_NAME} received - {BROKER} withdrawn; negative = {BROKER} shows more)", bold=True)
        put(ws, f"N{e + 3}", f"=ROUND(N{e + 2}-N{e + 1},2)", fmt=MONEY, bold=True)
        put(ws, f"A{e + 4}", f"Net received after money sent back to {BROKER}")
        put(ws, f"N{e + 4}", f"=N{sw_in}-N{sw_out}", fmt=MONEY)
        put(ws, f"A{e + 5}", f"A withdrawal dated 12/30-12/31 can post at {BANK_NAME} in January (and a late-December 2024 one "
                             "in early January 2025), so a small year-end difference is normal. Otherwise compare "
                             f"{BROKER}'s transfer history for the year, by account, with the {BROKER} deposits on Bank Transactions.",
            italic=True, size=9)
    ws.column_dimensions["A"].width = 62
    for col in MCOLS:
        ws.column_dimensions[col].width = 12
    ws.column_dimensions["N"].width = 14
    ws.freeze_panes = "B6"
    return pos


CHART_SHORT = {
    INC_SSA1: f"SSA - {SSA_OWNERS[0]}" if SSA_OWNERS else "Social Security", INC_BROKER: f"{BROKER} transfers",
    INC_CHECKS: "Check deposits", INC_INT: "Interest", INC_REFUND: "Tax refunds", INC_OTHER: "Other income",
    OTHER_CARDS: "Other card payments", "Housing - Mortgage extra principal": "Mortgage extra principal",
    CHECKS_CAT: "Checks written", PLAN_529: "529 plans", "Cash withdrawals (ATM / branch)": "Cash / ATM",
    "Person-to-person (Apple Cash / Cash App)": "Apple Cash / Cash App", "Disputed / Suspected Fraud": "Disputed / fraud",
    "Loan payments": "Loan", "Housing - Mortgage": "Mortgage",
}
if INC_SSA2:
    CHART_SHORT[INC_SSA2] = f"SSA - {SSA_OWNERS[1]}"
CHART_SHORT.update(_pc("CHART_SHORT_EXTRA", {}))        # short labels for your own income sources / categories
KFMT = '[>=1000]$#,##0,"K";[<=-1000]-$#,##0,"K";$0'      # 40000 -> $40K


def _short(label):
    return CHART_SHORT.get(label) or re.sub(r"\s*\(.*?\)", "", label).strip()


def write_cashflow_charts(ws, cf, tr, pos, ledger, year):
    """Charts read a small helper table (columns W onward) with short labels, so nothing overlaps."""
    from openpyxl.chart import Series
    ws.sheet_view.showGridLines = False
    put(ws, "A1", f"Income vs Expense Charts | {ylabel(year)}", bold=True, size=14)
    put(ws, "A2", "All charts read the Cash Flow and Transfers tabs (via the chart data table in columns W onward).",
        italic=True, size=9)
    for c in range(1, 20):
        ws.column_dimensions[get_column_letter(c)].width = 11
    ws.column_dimensions["W"].width = 26
    W, X = 23, 24          # label column, first value column

    inc_tot = defaultdict(float)
    exp_tot = defaultdict(float)
    for t in ledger["bank_txns"]:
        if t["year"] == year and t["flow"] == "Income":
            inc_tot[t["cat"]] += t["amount"]
        elif t["year"] == year and t["flow"] == "Expense":
            exp_tot[t["cat"]] -= t["amount"]
    citi_total = sum(net(t) for t in ledger["transactions"] if t["year"] == year)

    # ---- helper block 1: monthly series
    put(ws, "W3", "Chart data (feeds the charts)", bold=True)
    ws.cell(row=4, column=W, value="Series")
    for m in range(12):
        ws.cell(row=4, column=X + m, value=MONTHS[m])
    style_header(ws, 4, W, X + 11)
    rows = {}
    r = 5

    def monthly(label, sheet, srow):
        nonlocal r
        put(ws, f"{get_column_letter(W)}{r}", label)
        for m, col in enumerate(MCOLS):
            put(ws, f"{get_column_letter(X + m)}{r}", f"='{sheet}'!{col}{srow}", fmt=MONEY)
        rows[label] = r
        r += 1

    inc_shown = [src for src in INCOME_SOURCES if abs(inc_tot.get(src, 0)) > 0.004]
    for src in inc_shown:
        monthly(_short(src), "Cash Flow", pos["inc_rows"][src])
    first_src, last_src = rows[_short(inc_shown[0])], rows[_short(inc_shown[-1])]
    for label, sheet, srow in (("Total income", "Cash Flow", pos["inc_total"]),
                               ("Total expenses", "Cash Flow", pos["exp_total"]),
                               ("Net cash flow", "Cash Flow", pos["net"]),
                               (f"Received from {BROKER}", "Transfers", pos["sw_in"]),
                               (f"Sent to {BROKER}", "Transfers", pos["sw_out"])):
        monthly(label, sheet, srow)

    # ---- helper block 2: income by source, year total (largest first)
    r += 1
    put(ws, f"W{r}", "Income by source - year", bold=True)
    r += 1
    inc_first = r
    for src in sorted(inc_shown, key=lambda k: -inc_tot[k]):
        put(ws, f"W{r}", _short(src))
        put(ws, f"X{r}", f"='Cash Flow'!N{pos['inc_rows'][src]}", fmt=MONEY)
        r += 1
    inc_last = r - 1

    # ---- helper block 3: expenses by category, year total (largest first, incl. Citi cards)
    r += 1
    put(ws, f"W{r}", "Expenses by category - year", bold=True)
    r += 1
    exp_items = [(k, v, pos["exp_rows"][k]) for k, v in exp_tot.items() if v > 0.004]
    exp_items.append(("Credit cards (net spend)", citi_total, pos["citi_row"]))
    exp_items.sort(key=lambda x: -x[1])
    exp_first = r
    for label, _, srow in exp_items[:12]:
        put(ws, f"W{r}", _short(label) if label in exp_tot else label)
        put(ws, f"X{r}", f"='Cash Flow'!N{srow}", fmt=MONEY)
        r += 1
    exp_last = r - 1

    def style(ch, title, w=18, h=9.5):
        ch.title = title
        ch.x_axis.delete = False
        ch.y_axis.delete = False
        ch.y_axis.number_format = KFMT
        ch.width, ch.height = w, h
        return ch

    def hbar(ch):
        ch.type = "bar"
        ch.x_axis.scaling.orientation = "maxMin"     # largest on top
        ch.y_axis.crosses = "max"                    # value axis labels at the bottom
        ch.legend = None

    cats_m = Reference(ws, min_col=X, max_col=X + 11, min_row=4)

    def mrow(label):
        return Reference(ws, min_col=W, max_col=X + 11, min_row=rows[label], max_row=rows[label])

    c1 = style(BarChart(), "Income by Source and Month", h=10)
    c1.type = "col"; c1.grouping = "stacked"; c1.overlap = 100
    for src in inc_shown:
        c1.add_data(mrow(_short(src)), titles_from_data=True, from_rows=True)
    c1.set_categories(cats_m)
    polish_chart(c1)
    ws.add_chart(c1, "A4")

    c2 = style(BarChart(), "Income by Source - Year Total", h=10)
    hbar(c2)
    c2.series.append(Series(Reference(ws, min_col=X, min_row=inc_first, max_row=inc_last), title="Year total"))
    c2.set_categories(Reference(ws, min_col=W, min_row=inc_first, max_row=inc_last))
    polish_chart(c2, horizontal=True)
    ws.add_chart(c2, "K4")

    c3 = style(BarChart(), "Income vs Expenses by Month (line = net)", h=10)
    c3.type = "col"
    c3.add_data(mrow("Total income"), titles_from_data=True, from_rows=True)
    c3.add_data(mrow("Total expenses"), titles_from_data=True, from_rows=True)
    c3.set_categories(cats_m)
    c3.x_axis.tickLblPos = "low"                     # month names stay at the bottom when net goes negative
    c3.legend.position = "b"
    ln = LineChart()
    ln.add_data(mrow("Net cash flow"), titles_from_data=True, from_rows=True)
    for sr in ln.series:
        sr.smooth = False
    c3 += ln
    polish_chart(c3)
    ws.add_chart(c3, "A25")

    c4 = style(BarChart(), "Expenses by Category - Year Total", h=10)
    hbar(c4)
    c4.series.append(Series(Reference(ws, min_col=X, min_row=exp_first, max_row=exp_last), title="Year total"))
    c4.set_categories(Reference(ws, min_col=W, min_row=exp_first, max_row=exp_last))
    polish_chart(c4, horizontal=True)
    ws.add_chart(c4, "K25")

    c5 = style(BarChart(), f"{BROKER}: Received vs Sent Back", h=10)
    c5.type = "col"
    c5.add_data(mrow(f"Received from {BROKER}"), titles_from_data=True, from_rows=True)
    c5.add_data(mrow(f"Sent to {BROKER}"), titles_from_data=True, from_rows=True)
    c5.set_categories(cats_m)
    polish_chart(c5)
    ws.add_chart(c5, "A46")

    if "Check deposits" in rows:
        c6 = style(BarChart(), "Check & Branch Deposits by Month", h=10)
        c6.type = "col"
        c6.add_data(mrow("Check deposits"), titles_from_data=True, from_rows=True)
        c6.set_categories(cats_m)
        c6.legend = None
        polish_chart(c6)
        ws.add_chart(c6, "K46")
    ws.print_area = "A1:T67"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)


PERIOD_LAST = {}   # year -> last month with data, set per workbook


def ylabel(year):
    """'2025' for a full year, '2026 year to date through Sep' when later months have no statements yet."""
    last = PERIOD_LAST.get(year, 12)
    return f"{year}" if last >= 12 else f"{year} year to date through {MONTHS[last - 1]}"


# ----------------------------------------------------------------------------
# Advisor tabs: a high-level, advisor-ready view (income, transfers, spending by month and by category)
# ----------------------------------------------------------------------------
MONEY0 = '$#,##0;($#,##0);"-"'
ADVISOR_TAB = "548235"
CARD_FEES_CAT = "Card interest & fees"
OTHER_GROUP = "Fees, disputes & other"
ADVISOR_GROUPS = [
    ("Housing & home", ["Housing - Mortgage", "Home Improvement", "Home Services", "Household", "Utilities & Telecom"]),
    ("Extra mortgage principal", ["Housing - Mortgage extra principal"]),
    ("Food", ["Groceries", "Dining", "Warehouse Club"]),
    ("Transportation", ["Auto & Transportation", "Loan payments", "Fuel & Convenience"]),
    ("Healthcare", [MEDICAL, "Health & Fitness"]),
    ("Insurance", ["Insurance"]),
    ("Taxes & professional fees", ["Taxes", "Professional Services"]),
    ("Travel & leisure", ["Travel & Lodging", "Recreation & Entertainment", "Subscriptions & Media"]),
    ("Shopping & personal", ["Clothing", "General Retail", "Online Shopping", "Personal Care", "Electronics & Software",
                             "Gifts", "Government & Postal", "Childcare"]),
    ("Giving & family", ["Donations", PLAN_529]),
    ("Cash & checks", ["Cash withdrawals (ATM / branch)", CHECKS_CAT, "Person-to-person (Apple Cash / Cash App)"]),
    ("Card payments not yet itemized", [OTHER_CARDS]),
    (OTHER_GROUP, ["Bank & Card Fees", CARD_FEES_CAT, FRAUD, CREDITS, "Uncategorized"]),
]
for _cat, _grp in _pc("ADVISOR_EXTRA_CATEGORIES", {}).items():      # {"My loan category": "Transportation"}
    for _g, _cats in ADVISOR_GROUPS:
        if _g == _grp:
            _cats.append(_cat)
ADVISOR_LABELS = {OTHER_CARDS: "Card payments - other cards (purchases not itemized yet)"}


def advisor_group(cat):
    for g, cats in ADVISOR_GROUPS:
        if cat in cats:
            return g
    return OTHER_GROUP


def advisor_data(ledger, year):
    """Monthly figures (Jan..Dec) for the advisor tabs, computed straight from the ledger."""
    vec = lambda: [0.0] * 12
    inc, trn, bexp, cexp, fee = defaultdict(vec), defaultdict(vec), defaultdict(vec), defaultdict(vec), vec()
    for t in ledger["bank_txns"]:
        if t["year"] != year:
            continue
        m = t["month"] - 1
        if t["flow"] == "Income":
            inc[t["cat"]][m] += t["amount"]
        elif t["flow"] == "Expense":
            bexp[t["cat"]][m] -= t["amount"]
        else:
            trn[t["cat"]][m] += t["amount"]
    for t in ledger["transactions"]:
        if t["year"] != year:
            continue
        m = t["month"] - 1
        if t["type"] in (PURCHASE, REFUND):
            cexp[t["category"]][m] += net(t)
        elif t["type"] in (INTEREST, FEE):
            fee[m] += t["amount"]
    cats = {}
    for c in set(bexp) | set(cexp):
        cats[c] = [round(a + b, 2) for a, b in zip(bexp.get(c, vec()), cexp.get(c, vec()))]
    if any(abs(v) > 0.004 for v in fee):
        cats[CARD_FEES_CAT] = [round(v, 2) for v in fee]
    order = {g: i for i, (g, _) in enumerate(ADVISOR_GROUPS)}
    rows = [(c, advisor_group(c), v) for c, v in cats.items() if any(abs(x) > 0.004 for x in v)]
    rows.sort(key=lambda r: (order[r[1]], -sum(r[2]), r[0]))
    return {"inc": inc, "trn": trn, "bexp": bexp, "cexp": cexp, "cats": rows,
            "groups": [g for g, _ in ADVISOR_GROUPS if any(r[1] == g for r in rows)]}


def write_advisor(wb, ledger, year, bt_last, bld, cf_total_ref=None, live=True):
    """Adds the advisor Summary and Categories tabs. live=True: base cells are SUMIFS on the workbook's own tabs.
    live=False: base cells are values, so the two tabs can be sent on their own."""
    d = advisor_data(ledger, year)
    nmo = PERIOD_LAST.get(year, 12)
    cs = wb.create_sheet(f"{ADVISOR} Categories")
    sm = wb.create_sheet(f"{ADVISOR} Summary")
    for ws in (cs, sm):
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.tabColor = ADVISOR_TAB
    bt = lambda col: f"'Bank Transactions'!${col}$2:${col}${bt_last}"
    bank_f = lambda col, m, flow, cat: (f"SUMIFS({bt(col)},{bt('B')},{year},{bt('A')},{m},{bt('E')},\"{flow}\","
                                        f"{bt('F')},\"{cat}\")")
    mcol = lambda i, first: get_column_letter(first + i)      # i = 0..11

    def months_header(ws, row, c0, labels_first):
        heads = labels_first + MONTHS + ["Total"]
        for i, h in enumerate(heads, start=c0):
            ws.cell(row=row, column=i, value=h)
        return row

    # ------------------------------------------------------------ Advisor Categories
    put(cs, "A1", f"Spending by Category and Month | {ylabel(year)}", bold=True, size=14)
    put(cs, "A2", f"Everything paid from {BANK_NAME} plus credit card purchases (every card with statements loaded), grouped. Payments to cards "
                  "whose statements are loaded are not spending; the purchases are.", italic=True, size=9)
    hdr = 4
    for i, h in enumerate(["Category", "Group"] + MONTHS + ["Total", "% of spending", "Avg / month"], start=1):
        cs.cell(row=hdr, column=i, value=h)
    style_header(cs, hdr, 1, 17)
    first = hdr + 1
    r = hdr
    for cat, grp, vals in d["cats"]:
        r += 1
        put(cs, f"A{r}", ADVISOR_LABELS.get(cat, cat))
        put(cs, f"B{r}", grp)
        for i in range(12):
            col = mcol(i, 3)
            m = i + 1
            if live:
                parts = []
                if cat == CARD_FEES_CAT:
                    parts = [bld.sumifs("amount", [("year", str(year)), ("month", str(m)), ("type", f'"{t}"')])
                             for t in (INTEREST, FEE)]
                else:
                    if cat in d["bexp"]:
                        parts.append(bank_f("K", m, "Expense", cat))
                    if cat in d["cexp"]:
                        parts.append(bld.sumifs("spend", [("year", str(year)), ("month", str(m)),
                                                          ("category", f'"{cat}"')]))
                v = "=" + "+".join(parts)
            else:
                v = vals[i]
            put(cs, f"{col}{r}", v, fmt=MONEY0)
        put(cs, f"O{r}", f"=SUM(C{r}:N{r})", fmt=MONEY0, bold=True)
    last = r
    tot = last + 1
    put(cs, f"A{tot}", "TOTAL SPENDING", bold=True, fill=SUB_FILL)
    put(cs, f"B{tot}", "", fill=SUB_FILL)
    for i in range(13):
        col = get_column_letter(3 + i)
        put(cs, f"{col}{tot}", f"=SUM({col}{first}:{col}{last})", fmt=MONEY0, bold=True, fill=SUB_FILL)
    for rr in range(first, tot + 1):
        put(cs, f"P{rr}", f"=IF($O${tot}=0,0,O{rr}/$O${tot})", fmt=PCT, bold=(rr == tot),
            fill=SUB_FILL if rr == tot else None)
        put(cs, f"Q{rr}", f"=O{rr}/{nmo}", fmt=MONEY0, bold=(rr == tot), fill=SUB_FILL if rr == tot else None)
    n = tot + 2
    if live and cf_total_ref:
        put(cs, f"A{n}", "Cash Flow tab total expenses (must match)")
        put(cs, f"O{n}", f"={cf_total_ref}", fmt=MONEY0)
        put(cs, f"A{n + 1}", "Difference")
        put(cs, f"O{n + 1}", f"=ROUND(O{tot}-O{n},2)", fmt=MONEY0, bold=True)
        n += 3
    put(cs, f"A{n}", f"Card payments - other cards: payments from {BANK_NAME} to cards whose statements are not loaded yet "
                     "(Barclays, Amazon, Synchrony, Capital One, Chase, Best Buy, Kohl's, Comenity). They are counted when "
                     "paid and will be replaced by itemized purchases as those statements arrive.", italic=True, size=9)
    put(cs, f"A{n + 1}", "Card interest & fees are charged on the Citi statements and paid through the card payments.",
        italic=True, size=9)
    cs.column_dimensions["A"].width = 50
    cs.column_dimensions["B"].width = 26
    for i in range(12):
        cs.column_dimensions[mcol(i, 3)].width = 10
    for col, w in (("O", 12), ("P", 12), ("Q", 12)):
        cs.column_dimensions[col].width = w
    cs.freeze_panes = f"C{hdr + 1}"
    cs.page_setup.orientation = "landscape"
    cs.page_setup.fitToWidth = 1
    cs.page_setup.fitToHeight = 1
    cs.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

    # ------------------------------------------------------------ Advisor Summary
    put(sm, "A1", f"Household Cash Flow Summary | {ylabel(year)}", bold=True, size=14)
    put(sm, "A2", f"Prepared for {ADVISOR_FIRM}. Cash basis from the credit union statements (all accounts) plus credit card "
                  "purchases counted when made. Payments to those cards are transfers, not spending.", italic=True, size=9)
    H = 12                                        # table header row
    heads = ["Item"] + MONTHS + ["Total", "Avg / month"]
    for i, h in enumerate(heads, start=1):
        sm.cell(row=H, column=i, value=h)
    style_header(sm, H, 1, 15)

    def section(r, text):
        for c in range(1, 16):
            sm.cell(row=r, column=c).fill = SUB_FILL
        put(sm, f"A{r}", text, bold=True, fill=SUB_FILL)

    def line(r, label, per_month, bold=False, fill=None, avg=True):
        put(sm, f"A{r}", label, bold=bold, fill=fill)
        for i in range(12):
            col = mcol(i, 2)
            put(sm, f"{col}{r}", per_month(i, col), fmt=MONEY0, bold=bold, fill=fill)
        put(sm, f"N{r}", f"=SUM(B{r}:M{r})", fmt=MONEY0, bold=True, fill=fill)
        if avg:
            put(sm, f"O{r}", f"=N{r}/{nmo}", fmt=MONEY0, bold=bold, fill=fill)
        elif fill:
            put(sm, f"O{r}", "", fill=fill)

    def base(live_f, vals):
        return lambda i, col: (("=" + live_f(i + 1)) if live else vals[i])

    r = H + 1
    section(r, "INCOME RECEIVED")
    inc_rows = []
    for src in INCOME_SOURCES:
        if src == INC_BROKER or not any(abs(x) > 0.004 for x in d["inc"].get(src, [0])):
            continue
        r += 1
        label = "Deposits - checks and branch (source not on statement)" if src == INC_CHECKS else src
        line(r, label, base(lambda m, s=src: bank_f("J", m, "Income", s), d["inc"][src]))
        inc_rows.append(r)
    r += 1
    inc_tot = r
    line(r, "TOTAL INCOME", lambda i, col: f"=SUM({col}{inc_rows[0]}:{col}{inc_rows[-1]})", bold=True, fill=SUB_FILL)

    r += 2
    section(r, "TRANSFERS (money moved, not income and not spending)")
    r += 1
    sw_in = r
    line(r, f"Withdrawn from {BROKER} (IRA / brokerage)",
         base(lambda m: bank_f("J", m, "Income", INC_BROKER), d["inc"].get(INC_BROKER, [0.0] * 12)))
    r += 1
    sw_out = r
    line(r, f"Less: sent back to {BROKER}",
         base(lambda m: bank_f("L", m, "Transfer", TR_BROKER_OUT), d["trn"].get(TR_BROKER_OUT, [0.0] * 12)))
    r += 1
    sw_net = r
    line(r, "Net portfolio withdrawals", lambda i, col: f"={col}{sw_in}+{col}{sw_out}", bold=True)
    r += 1
    cardpay = [-(a + b + c) for a, b, c in zip(d["trn"].get(TR_CITI, [0.0] * 12), d["trn"].get(TR_APPLE, [0.0] * 12),
                                               d["trn"].get(TR_AMAZON, [0.0] * 12))]
    line(r, "Memo: paid to cards whose statements are loaded (purchases counted as spending)",
         base(lambda m: f"-({bank_f('L', m, 'Transfer', TR_CITI)}+{bank_f('L', m, 'Transfer', TR_APPLE)}"
                        f"+{bank_f('L', m, 'Transfer', TR_AMAZON)})",
              [round(x, 2) for x in cardpay]))
    r += 1
    line(r, f"Memo: moves between {BANK_NAME} accounts (net zero)",
         base(lambda m: bank_f("L", m, "Transfer", TR_INTERNAL), d["trn"].get(TR_INTERNAL, [0.0] * 12)))
    r += 2
    cash_in = r
    line(r, "TOTAL CASH IN (income + net portfolio withdrawals)",
         lambda i, col: f"={col}{inc_tot}+{col}{sw_net}", bold=True, fill=SUB_FILL)

    r += 2
    section(r, "SPENDING (detail by category on the next tab)")
    g_first = r + 1
    xp = None
    nrows = len(d["cats"])
    cat_last = hdr + nrows
    for g in d["groups"]:
        r += 1
        gv = [round(sum(v[i] for c, gg, v in d["cats"] if gg == g), 2) for i in range(12)]
        line(r, g, lambda i, col, rr=r: (f"=SUMIFS('{ADVISOR} Categories'!{mcol(i, 3)}${first}:{mcol(i, 3)}${cat_last},"
                                         f"'{ADVISOR} Categories'!$B${first}:$B${cat_last},$A{rr})"))
        if g == "Extra mortgage principal":
            xp = r
    g_last = r
    r += 1
    sp_tot = r
    line(r, "TOTAL SPENDING", lambda i, col: f"=SUM({col}{g_first}:{col}{g_last})", bold=True, fill=SUB_FILL)

    r += 2
    section(r, "RESULT")
    r += 1
    line(r, "Income less spending (before portfolio withdrawals)",
         lambda i, col: f"={col}{inc_tot}-{col}{sp_tot}", bold=True)
    res1 = r
    r += 1
    line(r, "Plus: net portfolio withdrawals", lambda i, col: f"={col}{sw_net}")
    r += 1
    line(r, "Cash in less spending", lambda i, col: f"={col}{res1}+{col}{r - 1}", bold=True, fill=SUB_FILL)

    # at-a-glance block (top)
    put(sm, "A4", "At a glance", bold=True)
    glance = [
        ("Total income (excluding portfolio withdrawals)", f"=N{inc_tot}"),
        (f"Net portfolio withdrawals from {BROKER}", f"=N{sw_net}"),
        ("Total spending", f"=N{sp_tot}"),
        ("Average spending per month", f"=O{sp_tot}"),
        ("Average per month excluding extra mortgage principal", f"=(N{sp_tot}-N{xp})/{nmo}" if xp else f"=O{sp_tot}"),
    ]
    for k, (lab, f) in enumerate(glance, start=5):
        put(sm, f"A{k}", lab)
        put(sm, f"B{k}", f, fmt=MONEY0, bold=True)
        sm[f"B{k}"].alignment = Alignment(horizontal="right")

    n = r + 2
    notes = [
        f"Social Security and pension deposits are recurring income. {BROKER} deposits are withdrawals from IRA / brokerage "
        "accounts and are shown as transfers, not income.",
        "Credit card purchases are counted in the month of the statement. "
        "Payments to those cards are excluded from spending so nothing is counted twice.",
        "Card payments to other cards (Barclays, Amazon, Synchrony, Capital One, Chase, others) are spending counted when "
        "paid; their purchases will be itemized when those statements are loaded.",
        "Extra mortgage principal is shown separately because it is a payoff, not a living expense.",
        "Behind these tabs, every month is tied to the credit union's own opening and closing balances.",
    ]
    for k, text in enumerate(notes):
        put(sm, f"A{n + k}", text, italic=True, size=9)
    sm.column_dimensions["A"].width = 62
    for i in range(12):
        sm.column_dimensions[mcol(i, 2)].width = 10
    sm.column_dimensions["N"].width = 12
    sm.column_dimensions["O"].width = 12
    sm.freeze_panes = f"B{H + 1}"
    sm.page_setup.orientation = "landscape"
    sm.page_setup.fitToWidth = 1
    sm.page_setup.fitToHeight = 1
    sm.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    for ws in (sm, cs):                      # advisor tabs go first
        wb.move_sheet(ws, offset=-wb.index(ws))
    wb.move_sheet(sm, offset=-wb.index(sm))
    return d


def build_advisor_file(year, ledger, path):
    """Stand-alone copy of the two advisor tabs with values only: safe to send without the underlying transactions."""
    PERIOD_LAST.setdefault(year, 12)
    wb = Workbook()
    wb.remove(wb.active)
    write_advisor(wb, ledger, year, 2, Builder(2), live=False)
    try:
        wb.save(path)
        return path
    except PermissionError:
        alt = path.with_name(path.stem + "_" + datetime.now().strftime("%H%M%S") + path.suffix)
        wb.save(alt)
        print(f"  {path.name} is open in Excel; saved {alt.name} instead.")
        return alt


def write_store_items(ws, ledger, year, bld):
    """What was bought on the store cards (Synchrony): one row per purchase or credit, with the item name printed
    on the statement, a category from the item name, and a tie-out to the card tabs."""
    ws.sheet_view.showGridLines = False
    rows = sorted((t for t in ledger["transactions"]
                   if "channel" in t and t["year"] == year and t["type"] in (PURCHASE, REFUND)),
                  key=lambda t: (t["date"], t["card"], t["desc"]))
    put(ws, "A1", f"Store Card Items - what was bought | {ylabel(year)}", bold=True, size=14)
    put(ws, "A2", "Each purchase and credit on the store card statements with the item name as the statement prints it "
                  "(Synchrony shows the first 30 characters or so). Category comes from keywords in the item name; add "
                  "your own rules as ITEM_RULES_FIRST in personal_config.py. The statement month is the month the "
                  "statement is counted in. Payments are not listed: they are transfers or already counted.",
        italic=True, size=9)
    put(ws, "A3", "Year", bold=True)
    put(ws, "B3", year)
    heads = ["Date", "Month", "Year", "Card", "Channel", "Item (as printed)", "Category", "Type", "Net spend",
             "Statement end", "Source file"]
    hdr = 6
    for i, h in enumerate(heads, start=1):
        ws.cell(row=hdr, column=i, value=h)
    style_header(ws, hdr, 1, len(heads))
    d = lambda s: datetime.strptime(s, "%Y-%m-%d").date() if s else None
    for r, t in enumerate(rows, start=hdr + 1):
        vals = [d(t["date"]), t["month"], t["year"], t["card"], t["channel"], t["item"] or (t["desc"] if not t["channel"] else "(no item shown)"),
                t["category"], t["type"], round(net(t), 2), d(t["stmt_end"]), t["source"]]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = Font(name=FONT, size=10)
            if c in (1, 10):
                cell.number_format = "mm/dd/yyyy"
            elif c == 9:
                cell.number_format = MONEY
    last = max(hdr + len(rows), hdr + 1)
    put(ws, "A4", "Net spend shown (follows the filter)", bold=True)
    put(ws, "I4", f"=SUBTOTAL(109,I{hdr + 1}:I{last})", bold=True, fmt=MONEY)
    cats = defaultdict(float)
    for t in rows:
        cats[t["category"]] += net(t)
    order = sorted(cats, key=lambda c: -cats[c])
    for i, h in enumerate(["Category", "Lines", "Net spend", "% of total"], start=13):
        ws.cell(row=hdr, column=i, value=h)
    style_header(ws, hdr, 13, 16)
    rng = lambda col: f"${col}${hdr + 1}:${col}${last}"
    r = hdr
    for cat in order:
        r += 1
        put(ws, f"M{r}", cat)
        put(ws, f"N{r}", f'=COUNTIFS({rng("G")},M{r},{rng("C")},$B$3)')
        put(ws, f"O{r}", f'=SUMIFS({rng("I")},{rng("G")},M{r},{rng("C")},$B$3)', fmt=MONEY)
    cat_first, cat_last = hdr + 1, r
    r += 1
    put(ws, f"M{r}", "Total", bold=True, fill=SUB_FILL)
    put(ws, f"N{r}", f"=SUM(N{cat_first}:N{cat_last})", bold=True, fill=SUB_FILL)
    put(ws, f"O{r}", f"=SUM(O{cat_first}:O{cat_last})", bold=True, fmt=MONEY, fill=SUB_FILL)
    for rr in range(cat_first, cat_last + 1):
        put(ws, f"P{rr}", f'=IF($O${r}=0,0,O{rr}/$O${r})', fmt="0.0%")
    tot_row = r
    r += 2
    put(ws, f"M{r}", "Tie-out to the card tabs", bold=True)
    r += 1
    for i, h in enumerate(["Card", "This tab", "Card tab", "Difference"], start=13):
        ws.cell(row=r, column=i, value=h)
    style_header(ws, r, 13, 16)
    for card in sorted({t["card"] for t in rows}):
        r += 1
        put(ws, f"M{r}", card)
        put(ws, f"N{r}", f'=SUMIFS({rng("I")},{rng("D")},M{r},{rng("C")},$B$3)', fmt=MONEY)
        put(ws, f"O{r}", "=" + bld.sumifs("spend", [("year", "$B$3"), ("card", f"$M{r}")]), fmt=MONEY)
        put(ws, f"P{r}", f"=ROUND(N{r}-O{r},2)", fmt=MONEY)
    for i, w in enumerate((11, 7, 6, 22, 20, 36, 24, 14, 12, 13, 34, 3, 24, 8, 13, 11), start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = f"A{hdr + 1}"
    ws.auto_filter.ref = f"A{hdr}:K{last}"
    return tot_row


def build_workbook(year, ledger, path):
    months = ({t["month"] for t in ledger["transactions"] if t["year"] == year}
              | {t["month"] for t in ledger["bank_txns"] if t["year"] == year})
    PERIOD_LAST[year] = max(months) if months else 12
    all_txns = sorted(ledger["transactions"], key=lambda t: (t["stmt_end"], t["card"], t["date"]))
    year_txns = [t for t in all_txns if t["year"] == year]
    bld = Builder(len(all_txns) + 1)
    bank_all = sorted(ledger["bank_txns"], key=lambda t: (t["date"], t["acct"]))
    has_bank = any(t["year"] == year for t in bank_all)
    bt_last = max(len(bank_all) + 1, 2)
    br_last = max(len(ledger["bank_recon"]) + 1, 2)

    wb = Workbook()
    if has_bank:
        cf = wb.active
        cf.title = "Cash Flow"
        pos_cf = write_cashflow(cf, ledger, year, bt_last, bld)
        ch = wb.create_sheet("Income vs Expense Charts")
        tr = wb.create_sheet("Transfers")
        pos_cf.update(write_transfers(tr, ledger, year, bt_last, br_last, pos_cf))
        write_cashflow_charts(ch, cf, tr, pos_cf, ledger, year)
    if year_txns:
        grand = wb.create_sheet("Grand Totals") if has_bank else wb.active
        grand.title = "Grand Totals"
        pos = write_summary(grand, bld, year, None, year_txns)
        cards = sorted({t["card"] for t in year_txns})
        write_charts(wb.create_sheet("Charts"), bld, grand, pos, year, cards)
        for card in cards:
            ws = wb.create_sheet(re.sub(r"[\\/*?:\[\]]", "-", card)[:31])
            write_summary(ws, bld, year, card, [t for t in year_txns if t["card"] == card])
        if any("channel" in t for t in year_txns):
            write_store_items(wb.create_sheet("Store Card Items"), ledger, year, bld)

    ws = wb.create_sheet("Transactions")
    for i, h in enumerate(TX_COLS, start=1):
        ws.cell(row=1, column=i, value=h)
    style_header(ws, 1, 1, len(TX_COLS))
    d = lambda s: datetime.strptime(s, "%Y-%m-%d").date() if s else None
    for r, t in enumerate(all_txns, start=2):
        vals = [t["month"], t["year"], t["card"], t["type"], t["vendor"], t["category"], t["driver"],
                t["amount"], f'=IF(D{r}="{PURCHASE}",H{r},IF(D{r}="{REFUND}",-H{r},0))',
                t["breach"], t["breach_note"], d(t["date"]), d(t["post"]), d(t["stmt_end"]),
                t["source"], t["desc"], t.get("subcat", ""), t.get("dcash", 0.0)]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = Font(name=FONT, size=10)
            if c in (12, 13, 14):
                cell.number_format = "mm/dd/yyyy"
            elif c in (8, 9, 18):
                cell.number_format = MONEY
    for i, w in enumerate((10, 9, 20, 22, 26, 24, 30, 14, 14, 8, 24, 12, 12, 14, 38, 46, 18, 14), start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(TX_COLS))}{max(len(all_txns) + 1, 2)}"

    if has_bank:
        write_bank_transactions(wb.create_sheet("Bank Transactions"), bank_all)

    ws = wb.create_sheet("Reconciliation")
    heads = ["Card", "Statement End", "Source File", "Previous Balance", "New Balance", "Purchases (statement)",
             "Purchases (read)", "Difference", "Payments+Credits (statement)", "Payments+Credits (read)",
             "Difference", "Fees (statement)", "Interest (statement)", "Status", "Daily Cash (memo)"]
    for i, h in enumerate(heads, start=1):
        ws.cell(row=1, column=i, value=h)
    style_header(ws, 1, 1, len(heads))
    ws.row_dimensions[1].height = 42
    for r, rc in enumerate(sorted(ledger["recon"], key=lambda x: (x["card"], x["stmt_end"])), start=2):
        vals = [rc["card"], d(rc["stmt_end"]), rc["source"], rc["prev"], rc["new"], rc["s_purchases"],
                rc["p_purchases"], f"=IF(OR(F{r}=\"\",G{r}=\"\"),\"\",ROUND(G{r}-F{r},2))",
                rc["s_paycred"], rc["p_paycred"], f"=IF(OR(I{r}=\"\",J{r}=\"\"),\"\",ROUND(J{r}-I{r},2))",
                rc["s_fees"], rc["s_interest"], rc["status"], rc.get("dcash")]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = Font(name=FONT, size=10)
            if c == 2:
                cell.number_format = "mm/dd/yyyy"
            elif 4 <= c <= 13 or c == 15:
                cell.number_format = MONEY
        if rc["status"] != "OK":
            for c in range(1, len(heads) + 1):
                ws.cell(row=r, column=c).fill = WARN_FILL
    n = len(ledger["recon"]) + 1
    put(ws, f"A{n + 2}", "Statements that tie to the bank's own totals:", bold=True)
    put(ws, f"D{n + 2}", f'=COUNTIF(N2:N{n},"OK")&" of "&COUNTA(N2:N{n})', bold=True)
    put(ws, f"A{n + 3}", "Purchases (read) = standard/promo purchases net of in-section credits, plus any "
                         "credit-balance refund Citi counts as a purchase. Payments+Credits are shown as "
                         "positive amounts. For an Excel/CSV spend list, 'Purchases' is the file's own total row.",
        italic=True, size=9)
    put(ws, f"A{n + 4}", "Apple Card: Purchases = charges net of returns and dispute credits ('Total charges, credits and "
                         "returns'); Previous Balance is the Previous Total Balance. Daily Cash is paid into Apple Cash, "
                         "not credited to the card balance, so it is a memo and is not netted from spend.",
        italic=True, size=9)
    put(ws, f"A{n + 5}", "Store cards (Synchrony): Purchases = the balance table's 'Purchases, Fees & Others Debits' less fees; "
                         "Payments+Credits = payments plus other credits; the page-1 summary and the balance table are "
                         "both read and must agree.", italic=True, size=9)
    for i, w in enumerate((20, 14, 36, 15, 15, 17, 15, 12, 19, 17, 12, 13, 13, 60, 14), start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    if has_bank:
        write_bank_recon(wb.create_sheet("Bank Reconciliation"), ledger, year, bt_last)

    if has_bank:
        write_advisor(wb, ledger, year, bt_last, bld, cf_total_ref=f"'Cash Flow'!N{pos_cf['exp_total']}")

    ws = wb.create_sheet("Processed Files")
    heads = ["File", "Hash (first 12)", "Card", "Statement End", "Transactions", "Status", "Processed At"]
    for i, h in enumerate(heads, start=1):
        ws.cell(row=1, column=i, value=h)
    style_header(ws, 1, 1, len(heads))
    for r, (h, f) in enumerate(sorted(ledger["files"].items(), key=lambda kv: kv[1].get("name", "")), start=2):
        for c, v in enumerate([f.get("name", ""), h[:12], f.get("card", ""), f.get("period_end", ""),
                               f.get("rows", 0), f.get("status", ""), f.get("processed_at", "")], start=1):
            ws.cell(row=r, column=c, value=v).font = Font(name=FONT, size=10)
    for i, w in enumerate((44, 16, 22, 14, 13, 60, 20), start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws = wb.create_sheet("Review")
    for i, h in enumerate(["Source File", "What needs a look"], start=1):
        ws.cell(row=1, column=i, value=h)
    style_header(ws, 1, 1, 2)
    for r, item in enumerate(ledger["review"], start=2):
        ws.cell(row=r, column=1, value=item.get("source", "")).font = Font(name=FONT, size=10)
        ws.cell(row=r, column=2, value=item.get("line", "")).font = Font(name=FONT, size=10)
    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 100

    try:
        wb.save(path)
        return path
    except PermissionError:
        alt = path.with_name(path.stem + "_" + datetime.now().strftime("%H%M%S") + path.suffix)
        wb.save(alt)
        print(f"  {path.name} is open in Excel; saved {alt.name} instead.")
        return alt


# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------
def main():
    ANALYSIS_DIR.mkdir(parents=True, exist_ok=True)
    load_settings_file()
    found = sorted(p for p in REDACTED_DIR.rglob("*") if p.is_file() and not p.name.startswith("~$"))
    all_pdfs = [p for p in found if p.suffix.lower() == ".pdf"]
    pdfs = [p for p in all_pdfs if "redacted" in p.name.lower()]
    sheets = [p for p in found if p.suffix.lower() in (".xlsx", ".xlsm", ".csv")]
    print(f"Found {len(all_pdfs)} PDF(s) ({len(pdfs)} with 'redacted' in the name) and "
          f"{len(sheets)} Excel/CSV file(s) in {REDACTED_DIR}.")
    for p in all_pdfs:
        if p not in pdfs:
            print(f"Ignored (no 'redacted' in name): {p.name}")
    if not pdfs and not sheets:
        raise SystemExit(f"No redacted PDF or Excel/CSV file found in {REDACTED_DIR}")

    ledger = load_ledger()
    new_files = 0
    new_statements = 0
    for pdf in pdfs:
        h = file_hash(pdf)
        if h in ledger["files"]:
            print(f"Already processed, skipped: {pdf.name}")
            continue
        if is_bank_pdf(pdf):
            if process_bank_pdf(ledger, pdf, h):
                new_files += 1
                new_statements += 1
            continue
        if is_apple_pdf(pdf):
            if process_apple_pdf(ledger, pdf, h):
                new_files += 1
                new_statements += 1
            continue
        if is_synchrony_pdf(pdf):
            if process_synchrony_pdf(ledger, pdf, h):
                new_files += 1
                new_statements += 1
            continue
        try:
            parsed = parse_statement(pdf)
        except Exception as exc:
            print(f"ERROR reading {pdf.name}: {exc}. Skipped, not recorded; fix or remove the file.")
            continue
        if not parsed["end"]:
            print(f"Could not read the billing period, NOT recorded (will retry next run): {pdf.name}")
            continue

        card_tab = f"{parsed['card_name']} {parsed['last4']}"
        card_tab = CARD_ALIASES.get(card_tab, card_tab)
        stmt_key = f"{card_tab}|{parsed['end'].isoformat()}"
        entry = {"name": pdf.name, "card": card_tab, "period_end": parsed["end"].strftime("%m/%d/%Y"),
                 "rows": 0, "processed_at": datetime.now().isoformat(timespec="seconds")}

        if stmt_key in ledger["statements"]:
            entry["status"] = "Duplicate statement (same card and end date); not counted"
            print(f"Duplicate statement period, not counted: {pdf.name} ({card_tab} {entry['period_end']})")
        else:
            rows, recon = build_rows(parsed)
            txns = [make_txn(t, card_tab, parsed["end"], pdf.name) for t in rows]
            ledger["transactions"].extend(txns)
            ledger["statements"][stmt_key] = h
            new_statements += 1
            recon.update({"card": card_tab, "stmt_end": parsed["end"].isoformat(), "source": pdf.name})
            ledger["recon"].append(recon)
            entry["rows"] = len(txns)
            notes = []
            if parsed["card_name"] == "Unknown Card":
                notes.append("card name not found")
            if parsed["last4"] == "UNK":
                notes.append("last four not found")
            if recon["status"] != "OK":
                notes.append(recon["status"])
                ledger["review"].append({"source": pdf.name, "line": recon["status"]})
            entry["status"] = "Processed" + (" - " + "; ".join(notes) if notes else " - ties to statement")
            print(f"Processed: {pdf.name} -> {card_tab}, ends {entry['period_end']}, {len(txns)} rows, "
                  f"{'ties to statement' if recon['status'] == 'OK' else recon['status']}")
        ledger["files"][h] = entry
        new_files += 1

    for path in sheets:
        h = file_hash(path)
        if h in ledger["files"]:
            print(f"Already processed, skipped: {path.name}")
            continue
        try:
            parsed = parse_ledger_sheet(path)
        except Exception as exc:
            print(f"ERROR reading {path.name}: {exc}. Skipped, not recorded; fix or remove the file.")
            continue
        if not parsed:
            print(f"Ignored (no date, payee and amount columns found): {path.name}")
            continue
        account = account_for(path)
        if account in REFERENCE_ACCOUNTS:      # a cross-check list only; the bank statements are the source
            total = round(sum(r["amount"] for r in parsed["rows"]), 2)
            ledger["reference"][account] = {"file": path.name, "total": total, "rows": len(parsed["rows"])}
            ledger["files"][h] = {"name": path.name, "card": account, "period_end": "", "rows": len(parsed["rows"]),
                                  "processed_at": datetime.now().isoformat(timespec="seconds"),
                                  "status": "Reference only (not counted); compared with the bank statements"}
            print(f"Reference only (not counted): {path.name} ({len(parsed['rows'])} rows, ${total:,.2f})")
            new_files += 1
            continue
        key = f"{account}|{path.name}"
        replaced = key in ledger["statements"]
        if replaced:   # same list, edited since last run: the new version replaces the old one
            ledger["transactions"] = [t for t in ledger["transactions"]
                                      if not (t["card"] == account and t["source"] == path.name)]
            ledger["recon"] = [r for r in ledger["recon"] if not (r["card"] == account and r["source"] == path.name)]
            ledger["review"] = [r for r in ledger["review"] if r.get("source") != path.name]
            ledger["files"] = {k: v for k, v in ledger["files"].items()
                               if not (v.get("name") == path.name and v.get("card") == account)}
        txns = [make_sheet_txn(r, account, path.name) for r in parsed["rows"]]
        recon = sheet_recon(parsed, txns, account, path.name)
        ledger["transactions"].extend(txns)
        ledger["recon"].append(recon)
        ledger["statements"][key] = h
        if recon["status"] != "OK":
            ledger["review"].append({"source": path.name, "line": recon["status"]})
        ledger["files"][h] = {
            "name": path.name, "card": account, "period_end": recon["stmt_end"][5:7] + "/" + recon["stmt_end"][8:]
            + "/" + recon["stmt_end"][:4], "rows": len(txns), "processed_at": datetime.now().isoformat(timespec="seconds"),
            "status": ("Updated list" if replaced else "Processed") + (" - rows add up to the file's totals"
                                                                       if recon["status"] == "OK" else " - " + recon["status"])}
        print(f"{'Updated' if replaced else 'Processed'}: {path.name} -> {account}, {len(txns)} rows, "
              f"{'adds up to the file totals' if recon['status'] == 'OK' else recon['status']}")
        new_files += 1
        new_statements += 1

    flag_cross_duplicates(ledger)
    apply_check_labels(ledger)
    apply_vendor_overrides(ledger)
    link_card_payments(ledger)
    assign_ssa_owners(ledger)
    bank_review_items(ledger)
    save_ledger(ledger)

    years = sorted({t["year"] for t in ledger["transactions"]} | {t["year"] for t in ledger["bank_txns"]})
    if not years:
        print("No transactions were read from any statement, so no workbook was written.")
        return
    for y in years:
        months = sorted({t["month"] for t in ledger["transactions"] if t["year"] == y}
                        | {t["month"] for t in ledger["bank_txns"] if t["year"] == y})
        stem = "Income_vs_Expense" if any(t["year"] == y for t in ledger["bank_txns"]) else "Credit_Card_Analysis"
        name = f"{stem}_{y}_{MONTHS[months[0] - 1]}-{MONTHS[months[-1] - 1]}.xlsx"
        out = build_workbook(y, ledger, ANALYSIS_DIR / name)
        print(f"Workbook: {out}")
        if any(t["year"] == y for t in ledger["bank_txns"]):
            mname = f"{ADVISOR}_Summary_{y}_{MONTHS[months[0] - 1]}-{MONTHS[months[-1] - 1]}.xlsx"
            print(f"For {ADVISOR} (values only, two tabs): {build_advisor_file(y, ledger, ANALYSIS_DIR / mname)}")
    bad = [r for r in ledger["recon"] if r["status"] != "OK"] + [r for r in ledger["bank_recon"] if r["status"] != "OK"]
    print(f"New files this run: {new_files} ({new_statements} statements counted). "
          f"Statements in total: {len(ledger['recon']) + len(ledger['bank_recon'])}. Statements that do not tie: {len(bad)}.")
    print("Check the Reconciliation tab, and the Review tab for anything that needs a look.")


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        import traceback
        err = traceback.format_exc()
        print(err)
        try:
            ANALYSIS_DIR.mkdir(parents=True, exist_ok=True)
            (ANALYSIS_DIR / "last_error.txt").write_text(err, encoding="utf-8")
            print(f"Error details also saved to {ANALYSIS_DIR / 'last_error.txt'}")
        except OSError:
            pass
        raise SystemExit(1)
