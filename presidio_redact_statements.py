"""Fast Eddies Financial Spending and Income Analysis - redactor.
Reads statement PDFs from <BASE>\\Inputs and writes redacted copies to <BASE>\\Redacted (names, addresses, e-mail,
account numbers removed; last four digits kept). BASE is C:\\CCardStatements unless FEFA_BASE is set.
Run the redactor first, then analyze_statements.py (which reads only the redacted copies)."""
import os
import re
import sys
import json
from collections import Counter
from datetime import datetime
from pathlib import Path

import pymupdf as fitz
from presidio_analyzer import AnalyzerEngine

BASE = Path(os.environ.get("FEFA_BASE") or r"C:\CCardStatements")
INPUT_DIR = BASE / "Inputs"
OUTPUT_DIR = BASE / "Redacted"
REPORT_PATH = OUTPUT_DIR / "redaction_review.json"

# Usage:
#   python presidio_redact_statements.py              all PDFs directly in Inputs
#   python presidio_redact_statements.py BankName    all PDFs in Inputs\BankName (and its subfolders)
#   add --test to either to redact only the first PDF (safe check before the full run)
ARGS = [a for a in sys.argv[1:] if not a.startswith("--")]
TEST_ONE_FILE = "--test" in sys.argv[1:]

# Leave transaction rows (lines that start with a date) completely untouched,
# including the vendor/creditor description text.
SKIP_TRANSACTION_ROWS = True

# Number of trailing card digits that must stay visible.
KEEP_LAST = 4

# Entities redacted on non-transaction lines (header, summary, address blocks).
REDACT_ENTITIES = {"PERSON", "US_BANK_NUMBER", "US_DRIVER_LICENSE", "US_PASSPORT", "US_SSN", "EMAIL_ADDRESS"}

ACCOUNT_LABEL = re.compile(
    r"\b(account|acct|member|customer)\s*(number|no\.?|#|id)\b|\baccount\s*:\s*$",
    re.IGNORECASE,
)
# Credit-union style sub-account header, e.g. "PREMIUM CHECKING1234567890123": label + long account id.
# Also matches loan headers such as "2025 AUTO LOAN2000012345678".
SUBACCOUNT_HEADER = re.compile(r"^\s*([A-Za-z0-9][A-Za-z0-9 &()/'.-]*?)(\d{10,})\s*$")
SUBACCOUNT_HEADER13 = re.compile(r"^\s*([A-Za-z0-9][A-Za-z0-9 &()/'.-]*?)(\d{13})\s*$")
MONEY_TOKEN = re.compile(r"^-?\$?[\d,]*\d\.\d{2}$")
CARD_LABEL = re.compile(r"\b(card|credit card)\s*(number|no\.?|#)\b", re.IGNORECASE)
DIGIT_RUN = re.compile(r"(?<!\d)(?:\d[ -]?){8,19}(?!\d)")

# A line is a transaction row if it starts with MM/DD, MM/DD/YY(YY) or "Mon DD".
ROW_DATE = re.compile(
    r"^\s*(?:\d{1,2}/\d{1,2}(?:/\d{2,4})?(?![\d/])"
    r"|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2}\b)",
    re.IGNORECASE,
)

# E-mail addresses are redacted on every line, including lines that share a row with a date.
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+'-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")

# A statement-period header such as "Apr 1 - Apr 30, 2025" starts with a date but is not a transaction row.
PERIOD_HEADER = re.compile(
    r"^\s*[A-Za-z]{3,9}\.?\s+\d{1,2}(?:,\s*\d{4})?\s*[\u2014\u2013-]\s*[A-Za-z]{3,9}\.?\s+\d{1,2},?\s*\d{4}", re.IGNORECASE)

# A line that names the account holder; the line right after it (name, e-mail) is redacted whole.
HOLDER_LABEL = re.compile(r"^\s*(Apple Card Customer|Cardholder|Card Holder|Account Holder|Primary Account Holder)\s*:?\s*$",
                          re.IGNORECASE)

# Text immediately before a span that means it is a masked card / "ending in" number.
MASK_BEFORE = re.compile(
    r"(?:[xX*\u2022\u25cf]{2,}[\s-]*|ending(?:\s+in)?\s*|last\s*(?:4|four)\s*(?:digits)?[:\s]*)$",
    re.IGNORECASE,
)


def clean_digits(text):
    return re.sub(r"\D", "", text)


def parse_us_date(text):
    for fmt in ("%m/%d/%y", "%m/%d/%Y"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    return None


def detect_statement_month(text, filename):
    """Return 'YYYY-MM' for the statement closing date, or None."""
    m = re.search(r"Statement\s+[A-Za-z]{3,9}\.?\s+\d{1,2}\s*[\u2014\u2013-]\s*([A-Za-z]{3,9})\.?\s+\d{1,2},?\s*(\d{4})", text, re.I)
    if m:
        for fmt in ("%b %Y", "%B %Y"):
            try:
                return datetime.strptime(f"{m.group(1)} {m.group(2)}", fmt).strftime("%Y-%m")
            except ValueError:
                pass
    m = re.search(r"billing period:?\s*(\d{1,2}/\d{1,2}/\d{2,4})\s*(?:-|\u2013|to)\s*(\d{1,2}/\d{1,2}/\d{2,4})", text, re.I)
    if m and parse_us_date(m.group(2)):
        return parse_us_date(m.group(2)).strftime("%Y-%m")
    m = re.search(r"(?:statement|closing) date:?\s*(\d{1,2}/\d{1,2}/\d{2,4})", text, re.I)
    if m and parse_us_date(m.group(1)):
        return parse_us_date(m.group(1)).strftime("%Y-%m")
    m = re.search(r"(?:statement|closing) date:?\s*([A-Za-z]{3,9})\.? \d{1,2},? (\d{4})", text, re.I)
    if m:
        for fmt in ("%b %Y", "%B %Y"):
            try:
                return datetime.strptime(f"{m.group(1)} {m.group(2)}", fmt).strftime("%Y-%m")
            except ValueError:
                pass
    m = re.search(r"(?:statement\s+period:?\s*)?(\d{1,2}/\d{1,2}/\d{2,4})\s*(?:-|\u2013|to)\s*(\d{1,2}/\d{1,2}/\d{2,4})", text, re.I)
    if m and parse_us_date(m.group(2)):
        return parse_us_date(m.group(2)).strftime("%Y-%m")
    m = re.search(r"(20\d\d)[-_ ]?(0[1-9]|1[0-2])", filename)
    return f"{m.group(1)}-{m.group(2)}" if m else None


def find_last4_candidates(text):
    """Last four digits of the card/account number, read from the ORIGINAL text."""
    cands = re.findall(r"ending\s+in\s*:?\s*(\d{4})\b", text, re.I)
    for m in re.finditer(r"(?:card|account)\s*(?:number|no\.?|#)\s*:?\s*", text, re.I):
        digits = ""
        for tok in re.split(r"[\s\-]+", text[m.end():m.end() + 40].strip()):
            if re.fullmatch(r"[xX*\u2022\d]{4,19}", tok):
                digits += re.sub(r"\D", "", tok)
            else:
                break
        if len(digits) >= 4:
            cands.append(digits[-4:])
    cands += re.findall(r"(?:[xX*\u2022]{4}[ -]?){1,3}(\d{4})\b", text)
    return cands


def add_redaction(rects, rect, reason):
    key = tuple(round(v, 2) for v in rect)
    rects.setdefault(key, {"rect": rect, "reasons": set()})
    rects[key]["reasons"].add(reason)


def get_lines(page):
    """Return [(text, [char_rect, ...])] using exact per-character boxes."""
    lines = []
    for block in page.get_text("rawdict")["blocks"]:
        if block.get("type") != 0:
            continue
        for line in block["lines"]:
            chars = [c for span in line["spans"] for c in span["chars"]]
            if not chars:
                continue
            text = "".join(c["c"] for c in chars)
            rects = [fitz.Rect(c["bbox"]) for c in chars]
            lines.append((text, rects))
    return lines


def transaction_row_flags(lines):
    """Flag every line sharing a vertical band with a date-leading line."""
    bands = []
    for text, rects in lines:
        if ROW_DATE.match(text) and not PERIOD_HEADER.match(text):
            y0 = min(r.y0 for r in rects)
            y1 = max(r.y1 for r in rects)
            bands.append((y0, y1))
    flags = []
    for text, rects in lines:
        mid = (min(r.y0 for r in rects) + max(r.y1 for r in rects)) / 2
        flags.append(any(y0 - 1 <= mid <= y1 + 1 for y0, y1 in bands))
    return flags


def span_rect(text, rects, start, end, keep_last=0):
    """
    Union rect covering text[start:end]. With keep_last, the final keep_last
    digits are excluded so they stay visible. Returns None if nothing to cover.
    """
    while start < end and text[start].isspace():
        start += 1
    while end > start and text[end - 1].isspace():
        end -= 1
    if start >= end:
        return None

    boundary = None
    if keep_last:
        digit_idx = [i for i in range(start, end) if text[i].isdigit()]
        if len(digit_idx) <= keep_last:
            return None
        boundary = digit_idx[-keep_last]  # first digit that stays visible
        end = boundary

    rect = fitz.Rect(rects[start])
    for i in range(start + 1, end):
        rect.include_rect(rects[i])

    if boundary is not None:
        # MuPDF removes any glyph whose box touches the redaction rect, so pull
        # the edge back slightly to protect the first kept digit.
        rect.x1 = min(rect.x1, rects[boundary].x0 - 0.5)
        if rect.x1 <= rect.x0:
            return None
    return rect


def is_protected(text, start, end):
    """Never redact short digit runs or masked / 'ending in' card references."""
    span = text[start:end]
    digits = clean_digits(span)
    if digits and len(digits) <= KEEP_LAST:
        return True
    if MASK_BEFORE.search(text[:start]):
        return True
    return False


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    src = INPUT_DIR / ARGS[0] if ARGS else INPUT_DIR
    pdfs = sorted(src.rglob("*.pdf")) if ARGS else sorted(src.glob("*.pdf"))
    if not pdfs:
        raise SystemExit(f"No PDF files found in {src}")
    print(f"Redacting {len(pdfs) if not TEST_ONE_FILE else 1} PDF(s) from {src}")

    files_to_process = pdfs[:1] if TEST_ONE_FILE else pdfs
    analyzer = AnalyzerEngine()
    analyze_entities = sorted(REDACT_ENTITIES | {"CREDIT_CARD"})
    report = {"input_folder": str(src), "test_one_file": TEST_ONE_FILE, "files": []}

    used_names = set()

    for pdf_path in files_to_process:
        subaccounts = []   # (label, last four) of sub-account headers, used for the output name
        holder_text = set()   # name / e-mail text removed from holder lines, re-checked after saving
        doc = fitz.open(pdf_path)
        # Read the original text first; the output name uses statement month + last four.
        pre_text = "\n".join(p.get_text() for p in doc)
        last4_votes = Counter(find_last4_candidates(pre_text))
        file_info = {"input": str(pdf_path), "output": None, "pages": []}

        for page_no, page in enumerate(doc, start=1):
            lines = get_lines(page)
            row_flags = transaction_row_flags(lines) if SKIP_TRANSACTION_ROWS else [False] * len(lines)

            redactions = {}
            reinserts = []   # last four digits of sub-account headers, redrawn after the full number is removed
            page_hits = []
            skipped_rows = 0

            prev_was_holder_label = False
            for (text, rects), is_row in zip(lines, row_flags):
                was_label, prev_was_holder_label = prev_was_holder_label, bool(HOLDER_LABEL.match(text))
                if text.strip():
                    for em in EMAIL_RE.finditer(text):
                        rect = span_rect(text, rects, em.start(), em.end())
                        if rect is not None:
                            add_redaction(redactions, rect, "EMAIL_ADDRESS")
                            holder_text.add(em.group(0))
                            page_hits.append({"type": "EMAIL_ADDRESS", "value": "[redacted]"})
                if was_label and text.strip() and not ACCOUNT_LABEL.search(text):
                    rect = span_rect(text, rects, 0, len(text))
                    if rect is not None:
                        add_redaction(redactions, rect, "ACCOUNT HOLDER line")
                        holder_text.update(t.strip() for t in re.split(r"[,;]", text) if len(t.strip()) >= 4)
                        page_hits.append({"type": "ACCOUNT_HOLDER_LINE", "value": "[redacted]"})
                    continue
                if is_row:
                    skipped_rows += 1
                    continue
                if not text.strip():
                    continue

                def redact(start, end, reason, keep_last=0):
                    rect = span_rect(text, rects, start, end, keep_last)
                    if rect is not None:
                        add_redaction(redactions, rect, reason)

                results = analyzer.analyze(text=text, language="en", entities=analyze_entities)

                # Sub-account header ("PREMIUM CHECKING1234567890123"): keep the last four digits only.
                sub = SUBACCOUNT_HEADER13.match(text) or SUBACCOUNT_HEADER.match(text)
                if sub:
                    # Remove the whole number, then redraw the last four digits (glyph removal inside a
                    # text run can shift the kept digits under the box on some PDFs).
                    redact(sub.start(2), sub.end(2), "SUB-ACCOUNT ID (keep last four)")
                    k0 = sub.end(2) - KEEP_LAST
                    box = fitz.Rect(rects[sub.start(2)])
                    h = rects[k0].y1 - rects[k0].y0
                    box.include_rect(rects[sub.end(2) - 1])
                    pix = page.get_pixmap(clip=box, dpi=36)
                    px = pix.samples
                    avg = sum(px) / max(len(px), 1)   # dark bar -> white digits, light page -> black digits
                    reinserts.append(((rects[k0].x0, rects[k0].y1 - 0.212 * h / 1.15), sub.group(2)[-KEEP_LAST:],
                                      h / 1.15, (1, 1, 1) if avg < 128 else (0, 0, 0)))
                    subaccounts.append((sub.group(1).strip(), sub.group(2)[-4:]))
                    page_hits.append({"type": "SUBACCOUNT_ID",
                                      "value": f"[leading digits redacted; last four: {sub.group(2)[-4:]}]"})
                    continue

                # Credit card numbers: redact everything except the last four digits.
                card_spans = []
                for result in results:
                    if result.entity_type != "CREDIT_CARD":
                        continue
                    digits = clean_digits(text[result.start:result.end])
                    if 12 <= len(digits) <= 19:
                        card_spans.append((result.start, result.end))
                        last4_votes[digits[-4:]] += 1
                        redact(result.start, result.end, "CREDIT_CARD (keep last four)", KEEP_LAST)
                        page_hits.append({
                            "type": "CREDIT_CARD",
                            "value": f"[leading digits redacted; last four: {digits[-4:]}]"
                        })

                # Other PII (name, SSN, etc.), skipping card overlaps and protected spans.
                for result in results:
                    if result.entity_type not in REDACT_ENTITIES:
                        continue
                    if any(result.start < ce and result.end > cs for cs, ce in card_spans):
                        continue
                    if is_protected(text, result.start, result.end):
                        continue
                    if MONEY_TOKEN.match(text[result.start:result.end].strip()):
                        continue   # an amount such as 20,778.45 is not an identifier
                    redact(result.start, result.end, result.entity_type)
                    page_hits.append({"type": result.entity_type, "value": "[redacted]"})

                # Labeled account numbers (e.g. "Account Number: 12345678").
                if ACCOUNT_LABEL.search(text):
                    for match in DIGIT_RUN.finditer(text):
                        digits = clean_digits(match.group(0))
                        if not 6 <= len(digits) <= 20:
                            continue
                        if any(match.start() < ce and match.end() > cs for cs, ce in card_spans):
                            continue
                        if is_protected(text, match.start(), match.end()):
                            continue
                        redact(match.start(), match.end(), "ACCOUNT NUMBER label")
                        page_hits.append({"type": "ACCOUNT_NUMBER_CONTEXT", "value": "[redacted]"})

                # Labeled card numbers Presidio did not catch (e.g. failed Luhn check).
                if CARD_LABEL.search(text):
                    for match in DIGIT_RUN.finditer(text):
                        digits = clean_digits(match.group(0))
                        if not 12 <= len(digits) <= 19:
                            continue
                        if any(match.start() < ce and match.end() > cs for cs, ce in card_spans):
                            continue
                        redact(match.start(), match.end(), "CARD NUMBER label (keep last four)", KEEP_LAST)
                        last4_votes[digits[-4:]] += 1
                        page_hits.append({
                            "type": "CARD_NUMBER_CONTEXT",
                            "value": f"[leading digits redacted; last four: {digits[-4:]}]"
                        })

            for entry in redactions.values():
                page.add_redact_annot(entry["rect"], fill=(0, 0, 0))
            if redactions:
                page.apply_redactions(images=0, graphics=0, text=0)
            for pt, digits, size, color in reinserts:
                page.insert_text(pt, digits, fontsize=size, fontname="helv", color=color)

            file_info["pages"].append({
                "page": page_no,
                "redaction_count": len(redactions),
                "transaction_rows_skipped": skipped_rows,
                "detections": page_hits,
            })

        month = detect_statement_month(pre_text, pdf_path.name) or "UnknownMonth"
        last4 = last4_votes.most_common(1)[0][0] if last4_votes else "UNK"
        if subaccounts:   # name the file after the checking account when there is one, else the first account
            last4 = next((n4 for lab, n4 in subaccounts if "CHECKING" in lab.upper()), subaccounts[0][1])
        if re.search(r"Apple Card Customer", pre_text):   # the only digits on an Apple Card statement are the bank's
            last4 = "AppleCard"
        out_path = OUTPUT_DIR / f"{month}_REDACTED_{last4}.pdf"
        n = 2
        while out_path in used_names:
            out_path = OUTPUT_DIR / f"{month}_REDACTED_{last4}_{n}.pdf"
            n += 1
        used_names.add(out_path)
        file_info["output"] = str(out_path)
        if month == "UnknownMonth" or last4 == "UNK":
            print(f"CHECK NAME for {pdf_path.name}: month={month}, last4={last4}")

        doc.save(out_path, garbage=4, deflate=True, clean=True)
        doc.close()
        out_text = "\n".join(pg.get_text() for pg in fitz.open(out_path))
        leaks = sorted({e for e in EMAIL_RE.findall(out_text)} | {t for t in holder_text if t in out_text})
        if leaks:
            out_path.unlink()
            print(f"LEAK CHECK FAILED for {pdf_path.name}: {len(leaks)} identifying item(s) are still readable, so the "
                  f"output was deleted. Do not use this file; tell Claude.")
            file_info["leak_check"] = "FAILED, output deleted"
            report["files"].append(file_info)
            continue
        file_info["leak_check"] = "passed"
        report["files"].append(file_info)
        print(f"Created: {out_path.name}  (from {pdf_path.name})")
        print(f"Pages: {len(file_info['pages'])}; redaction boxes: {sum(p['redaction_count'] for p in file_info['pages'])}")

    REPORT_PATH.write_text(json.dumps(report, indent=2), encoding="utf-8")
    print(f"Review report: {REPORT_PATH}")
    print("IMPORTANT: Review the redacted PDF visually before processing all files.")
    if TEST_ONE_FILE:
        print("This run processed only the first PDF in alphabetical order.")


if __name__ == "__main__":
    main()
