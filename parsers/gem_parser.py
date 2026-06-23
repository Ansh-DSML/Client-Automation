"""
GeM Bid PDF Parser
==================
Parses GeM (Government e Marketplace) bid PDFs and extracts structured
data for the RFQ pipeline Excel output.

Works in "offline" mode — reads the PDF text + tables directly.
The parse() function is called by router.py with:
    parse(text, pdf_bytes=pdf_bytes)
"""

import re
import os
import io
import sys
import time
import glob
import shutil
import tempfile
import logging
from datetime import datetime, timedelta
from pathlib import Path

import pdfplumber
import pandas as pd

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────
#  CONSTANTS
# ─────────────────────────────────────────────────────────────
GEM_BASE_URL  = "https://bidplus.gem.gov.in"
GEM_LOGIN_URL = "https://gem.gov.in/login"
DEFAULT_STS   = "WORKING"
DEFAULT_MFG   = "N/A"


# ─────────────────────────────────────────────────────────────
#  SECTION 1 – TEXT CLEANING
# ─────────────────────────────────────────────────────────────

def clean_text(t: str) -> str:
    """Remove CID garbage characters from pdfplumber output."""
    t = re.sub(r'\(cid:\d+\)', '', t)
    t = re.sub(r'\s+', ' ', t).strip()
    return t


# ─────────────────────────────────────────────────────────────
#  SECTION 2 – FIELD PARSERS
# ─────────────────────────────────────────────────────────────

def parse_bid_number(text: str) -> str:
    """Extract  GEM/2026/B/7587781  from the PDF text."""
    m = re.search(r'Bid Number[:\s]*(GEM/\d+/[A-Z]/\d+)', text, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    # fallback: look for GEM/ pattern anywhere
    m = re.search(r'(GEM/\d{4}/[A-Z]/\d+)', text)
    return m.group(1).strip() if m else "UNKNOWN"


def parse_organisation(text: str) -> str:
    """Return e.g. Hindustan Aeronautics Limited (hal)."""
    m = re.search(r'Organisation Name\s+(.+?)(?:\n|Department)', text, re.IGNORECASE)
    return m.group(1).strip() if m else ""


def parse_city_from_address(text: str) -> str:
    """
    Address block looks like: '***********Rangareddi 15 60'
    Rangareddi/Rangareddy is in Hyderabad district → map to HYDERABAD.
    Extend this dict as needed.
    """
    DISTRICT_TO_CITY = {
        "rangareddi":    "HYDERABAD",
        "rangareddy":    "HYDERABAD",
        "hyderabad":     "HYDERABAD",
        "secunderabad":  "HYDERABAD",
        "bangalore":     "BANGALORE",
        "bengaluru":     "BANGALORE",
        "chennaicity":   "CHENNAI",
        "chennai":       "CHENNAI",
        "delhi":         "DELHI",
        "newdelhi":      "DELHI",
        "mumbai":        "MUMBAI",
        "pune":          "PUNE",
        "kolkata":       "KOLKATA",
        "lucknow":       "LUCKNOW",
        "nasik":         "NASIK",
        "nashik":        "NASIK",
        "koraput":       "KORAPUT",
        "kanpur":        "KANPUR",
        "korwa":         "KORWA",
        "ozar":          "NASIK",
        "sulur":         "COIMBATORE",
        "tambaram":      "CHENNAI",
        "barrackpore":   "KOLKATA",
    }
    # Pattern: asterisks followed by a city/district name, then numbers
    m = re.search(r'\*+([A-Za-z]+)\s+\d+', text)
    if m:
        raw = m.group(1).strip().lower()
        for key, city in DISTRICT_TO_CITY.items():
            if raw.startswith(key) or key in raw:
                return city
        return m.group(1).strip().upper()
    return "UNKNOWN"


def parse_customer(text: str) -> str:
    """
    Build the CUST string: '<ORG SHORT>-<CITY>'
    e.g. 'HAL-HYDERABAD'
    """
    org = parse_organisation(text).upper()
    city = parse_city_from_address(text)

    # Short name rules
    ORG_SHORT = {
        "HINDUSTAN AERONAUTICS LIMITED (HAL)": "HAL",
        "HINDUSTAN AERONAUTICS LIMITED":       "HAL",
        "HAL":                                 "HAL",
        "DRDO":                                "DRDO",
        "BEL":                                 "BEL",
        "ORDNANCE FACTORY":                    "OF",
        "INDIAN ARMY":                         "IA",
        "INDIAN NAVY":                         "IN",
        "INDIAN AIR FORCE":                    "IAF",
    }
    short = org
    for k, v in ORG_SHORT.items():
        if k in org:
            short = v
            break

    return f"{short}-{city}"


def parse_bid_end_date_time(text: str):
    """
    Returns (due_date_str, time_str).
    due_date is end_date MINUS 3 days, formatted DD-MM-YYYY.
    """
    m = re.search(
        r'Bid End Date/Time\s+(\d{2}-\d{2}-\d{4})\s+(\d{2}:\d{2}:\d{2})',
        text, re.IGNORECASE
    )
    if m:
        raw_date = m.group(1)    # e.g. 16-06-2026
        raw_time = m.group(2)    # e.g. 14:00:00
        dt = datetime.strptime(raw_date, "%d-%m-%Y")
        due_dt = dt - timedelta(days=3)
        return due_dt.strftime("%d-%m-%Y"), raw_time
    return "", ""


def parse_delivery_days(text: str) -> str:
    """
    Extract the first Delivery Days value from the consignee tables.
    Format in PDF: '***Rangareddi 15 60' where 60 is the delivery days.
    """
    # Pattern: asterisks + city + qty + delivery_days
    m = re.search(r'\*+[A-Za-z]+\s+(\d+)\s+(\d+)', text)
    if m:
        return m.group(2)   # second number = delivery days
    # fallback: explicit Delivery Days label
    m = re.search(r'Delivery Days?\s+(\d+)', text, re.IGNORECASE)
    return m.group(1) if m else ""


def parse_evaluation_method(text: str) -> str:
    """Map Evaluation Method to IWE or TVWE."""
    m = re.search(r'Evaluation Method\s+([^\n]+)', text, re.IGNORECASE)
    if m:
        val = m.group(1).strip().lower()
        if "item wise" in val or "item-wise" in val:
            return "IWE"
        if "total value" in val:
            return "TVWE"
    return "IWE"   # safe default


# ─────────────────────────────────────────────────────────────
#  SECTION 2b – PART NUMBER EXTRACTION HELPER
# ─────────────────────────────────────────────────────────────

def _extract_part_no_from_category(category_text: str) -> str:
    """
    Extract the HAL/GeM part number from an Item Category string.

    Patterns handled (tested across 6 PDFs):
      '68756716 - Self Solder Sleeve...'   → 68756716      (digits then ' - ')
      '51126961-circular Connector...'     → 51126961      (digits then '-word')
      '51107159 Crystal Oscillator...'     → 51107159      (digits then space+word)
      '51106932 - CONNECTOR 530721-3...'   → 51106932      (digits then ' - ')
      'Connector, 33505964011, 80 Pin...'  → 33505964011   (word first, digits later)
      '901316P51S, SMA STR PLUG'           → 901316P51S    (alphanumeric mixed)
    """
    # Normalise whitespace — pdfplumber puts \n in multi-line cells
    text = re.sub(r'\s+', ' ', category_text).strip()

    # Case A: starts with digits+letters ending at comma or ' - '
    # Handles: "901316P51S, SMA STR PLUG", "68756716 - Self Solder...", "51106932 - CONNECTOR..."
    m = re.match(r'^(\d[\w]*?)(?:\s*,|\s+-\s)', text)
    if m:
        return m.group(1).strip()

    # Case B: starts with pure digits then hyphen (e.g. "51126961-circular")
    m = re.match(r'^(\d+)-', text)
    if m:
        return m.group(1).strip()

    # Case C: starts with pure digits then space (e.g. "51107159 Crystal...")
    m = re.match(r'^(\d+)\s', text)
    if m:
        return m.group(1).strip()

    # Case D: digits appear after a word (e.g. "Connector, 33505964011, ...")
    m = re.search(r'(?<![A-Za-z\-])(\d{6,})', text)
    if m:
        return m.group(1).strip()

    # Fallback — should not reach here for GeM PDFs
    return text.upper()


# ─────────────────────────────────────────────────────────────
#  SECTION 3 – SCHEDULE & CONSIGNEE TABLE PARSERS
# ─────────────────────────────────────────────────────────────

def parse_schedules_from_text(text: str) -> list[dict]:
    """
    Text-based fallback schedule extraction.
    Looks for lines like: 'Schedule N  <description text>  <qty>'
    """
    items = []
    pattern = re.compile(
        r'Schedule\s+(\d+)\s+'       # Schedule N
        r'(.+?)\s+'                   # description (non-greedy)
        r'(\d+)\s*$',                 # trailing quantity
        re.IGNORECASE | re.MULTILINE
    )
    for m in pattern.finditer(text):
        sched_no  = int(m.group(1))
        desc_raw  = m.group(2).strip()
        qty       = m.group(3).strip()
        part_no   = _extract_part_no_from_category(desc_raw)
        items.append({
            "schedule_no": sched_no,
            "part_no":     part_no,
            "description": desc_raw,
            "qty":         qty,
        })
    return items


def parse_schedules_from_tables(pdf_bytes: bytes) -> list[dict]:
    """
    Parse evaluation schedules from the PDF's Evaluation Schedules table.

    Table structure (detected by the data rows, NOT the garbled Hindi header):
      3-column: ['Schedule N', 'Item/Category text', 'Quantity']
      4-column: ['Schedule N', 'Estimated Value', 'Item/Category text', 'Quantity']

    Returns list of dicts: {schedule_no, part_no, description, qty}
    """
    items = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if not table or len(table) < 2:
                        continue

                    # Detect this is the Evaluation Schedules table by checking
                    # whether any data row starts with "Schedule N"
                    schedule_rows = []
                    for row in table:
                        if not row:
                            continue
                        cell0 = str(row[0] or "").strip()
                        m = re.match(r'Schedule\s+(\d+)', cell0, re.IGNORECASE)
                        if m:
                            schedule_rows.append((int(m.group(1)), row))

                    if not schedule_rows:
                        continue

                    # Determine layout: 3-col or 4-col
                    # Detect by checking if column 1 looks like a large number (Estimated Value)
                    # vs a descriptive string (Item/Category)
                    for sched_no, row in schedule_rows:
                        ncols = len(row)
                        if ncols >= 4:
                            # Check if col[1] is a numeric estimated value
                            col1 = str(row[1] or "").strip().replace(",", "")
                            if re.match(r'^\d+(\.\d+)?$', col1):
                                # 4-col layout: col[2]=description, col[3]=qty
                                description = re.sub(r'\s+', ' ', str(row[2] or "")).strip()
                                qty = str(row[3] or "").strip()
                            else:
                                # 4-col but col[1] is actually description
                                description = re.sub(r'\s+', ' ', str(row[1] or "")).strip()
                                qty = str(row[-1] or "").strip()
                        elif ncols == 3:
                            # 3-col layout: col[1]=description, col[2]=qty
                            description = re.sub(r'\s+', ' ', str(row[1] or "")).strip()
                            qty = str(row[2] or "").strip()
                        else:
                            continue

                        # Extract part number from the leading numeric token
                        part_no = _extract_part_no_from_category(description)

                        if description and qty:
                            items.append({
                                "schedule_no": sched_no,
                                "part_no":     part_no,
                                "description": description,
                                "qty":         qty,
                            })
    except Exception as e:
        log.warning(f"Table-based schedule extraction failed: {e}")
    return items


def parse_consignee_delivery_days(pdf_bytes: bytes) -> list[str]:
    """
    Extract delivery days from each consignee table, in schedule order.

    Consignee tables are 5-column tables:
      [S.No | Consignee | Address | Quantity | Delivery Days]
    They are identified by '***' in the Address column (col index 2).

    Returns a list of delivery_days strings, one per schedule, in order.
    e.g. ['179', '180', '180'] for a 3-schedule bid.
    """
    delivery_days = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if not table or len(table) < 2:
                        continue
                    for row in table:
                        if not row or len(row) < 5:
                            continue
                        # Address column (index 2) contains '***'
                        addr = str(row[2] or "")
                        if "***" not in addr:
                            continue
                        # Delivery days is the last column
                        days = str(row[-1] or "").strip()
                        if days.isdigit():
                            delivery_days.append(days)
    except Exception as e:
        log.warning(f"Consignee delivery days extraction failed: {e}")
    return delivery_days


def parse_item_category_parts(text: str) -> list[str]:
    """
    Extract part numbers from the Item Category field.
    E.g.: 'TWCB476K050CCYZ0000 , LM1086ISX-ADJ , VSSC8L45-M3 9AT'
    Or from GeMARPTS searched strings.
    """
    parts = []
    # Look for Item Category section
    m = re.search(r'Item Category\s+(.+?)(?:\n|GeMARPTS)', text, re.IGNORECASE | re.DOTALL)
    if m:
        raw = m.group(1).strip()
        # Split by commas
        for token in re.split(r'\s*,\s*', raw):
            token = token.strip()
            if token and len(token) >= 3 and not token.startswith('(cid:'):
                parts.append(token.upper())
    return parts


def parse_single_item_from_text(text: str, pdf_bytes: bytes = None) -> list[dict]:
    """
    Fallback for single-item bids that have NO Evaluation Schedules table.

    In these bids the item lives in the 'Item Category' field on page 1,
    and the quantity lives in the consignee table (col[3]).

    Returns a list with one dict, or [] if nothing can be extracted.
    """
    # Extract Item Category from text
    m = re.search(r'/Item Category\s+(.+?)(?:\n|GeMARPTS)', text, re.IGNORECASE | re.DOTALL)
    if not m:
        return []

    item_cat = re.sub(r'\s+', ' ', m.group(1)).strip()
    if not item_cat:
        return []

    part_no = _extract_part_no_from_category(item_cat)

    # Extract quantity from the consignee table (col[3], the row with '***' in col[2])
    qty = "1"
    if pdf_bytes:
        try:
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                for page in pdf.pages:
                    for table in (page.extract_tables() or []):
                        if not table:
                            continue
                        for row in table:
                            if row and len(row) >= 4 and "***" in str(row[2] or ""):
                                q = str(row[3] or "").strip()
                                if q.isdigit():
                                    qty = q
                                    break
        except Exception as e:
            log.warning(f"Single item qty extraction failed: {e}")

    return [{
        "schedule_no": 1,
        "part_no":     part_no,
        "description": item_cat,
        "qty":         qty,
    }]


def determine_customer_remarks(customer: str) -> str:
    """Fixed remarks if the customer is HAL (any city)."""
    if customer.upper().startswith("HAL"):
        return "OEM / AD COC | DATE CODE IS MANDATORY WITHIN 5 YEARS"
    return ""


# ─────────────────────────────────────────────────────────────
#  SECTION 4 – HELPER UTILITIES
# ─────────────────────────────────────────────────────────────

def parse_part_number_from_description(desc: str) -> str:
    """
    From 'CAPACITOR, P/N:TWCB476K050CCYZ0000 ,HIGH TEMPERATURE...'
    extract 'TWCB476K050CCYZ0000'
    """
    m = re.search(r'P/N[:\s]*([A-Za-z0-9\-\.]+)', desc, re.IGNORECASE)
    return m.group(1).strip() if m else ""


def _extract_part_from_title(title: str) -> str:
    """
    Title examples:
      'RF COAXIAL PANEL MOUNT  R12551200'          → R12551200
      'STRAIGHT JACK CON,Pno  R222.426.020'        → R222.426.020
    Strategy: take the last token that looks like a part number.
    """
    m = re.search(r'(?:Pno|P/N)[:\s]+([^\s,]+)', title, re.IGNORECASE)
    if m:
        return m.group(1).strip()

    tokens = title.split()
    for tok in reversed(tokens):
        tok = tok.rstrip(",.:;")
        if re.search(r'\d', tok) and len(tok) >= 4:
            return tok

    return title.strip()


# ─────────────────────────────────────────────────────────────
#  SECTION 5 – ROW BUILDER
# ─────────────────────────────────────────────────────────────

def _build_output_rows(
    bid_number: str,
    customer: str,
    due_date: str,
    bid_time: str,
    lead_time: str,
    eval_method: str,
    remarks: str,
    line_items: list[dict],
    today_str: str,
) -> list[dict]:
    """
    Construct one Excel row per line item.
    """
    rows = []
    for sn, li in enumerate(line_items, start=1):
        # Use per-item lead time if available, else fall back to bid-level
        item_lt = li.get("lead_time", lead_time)
        rows.append({
            "Enquiry":       bid_number,
            "Cust":          customer,
            "RFQ D":         today_str,
            "Due D":         due_date,
            "Time":          bid_time,
            "L/T":           item_lt,          # ← per-schedule delivery days
            "STS":           DEFAULT_STS,
            "SN":            sn,
            "CUST PART NO.": li.get("part_no", ""),
            "Description":   li.get("description", ""),
            "Enq. Part No*": li.get("enq_part_no", li.get("part_no", "")),
            "Enq. Mfg":      DEFAULT_MFG,
            "Qty":           li.get("qty", ""),
            "UOM":           li.get("uom", "Numbers"),
            "TP":            eval_method,
            "Remarks":       remarks,
        })
    return rows


# ─────────────────────────────────────────────────────────────
#  SECTION 6 – MAIN PARSE FUNCTION
#  Called by router.py: parse(text, pdf_bytes=pdf_bytes)
# ─────────────────────────────────────────────────────────────

def parse(text: str, pdf_bytes: bytes = None) -> list[dict]:
    """
    Main entry point for the GeM parser.
    Accepts the full extracted text and optional raw PDF bytes.
    Returns a list of dicts ready for the GeM Excel writer.
    """
    # ── STEP 1: Parse header fields from text ────────────────
    bid_number  = parse_bid_number(text)
    customer    = parse_customer(text)
    due_date, bid_time = parse_bid_end_date_time(text)
    eval_method = parse_evaluation_method(text)
    remarks     = determine_customer_remarks(customer)
    today_str   = datetime.today().strftime("%d-%m-%Y")

    log.info(f"  GeM Bid No   : {bid_number}")
    log.info(f"  Customer     : {customer}")
    log.info(f"  Due Date     : {due_date}  Time: {bid_time}")
    log.info(f"  Eval Method  : {eval_method}")

    # ── STEP 2: Extract schedule items from tables ────────────────
    schedule_items = []
    if pdf_bytes:
        schedule_items = parse_schedules_from_tables(pdf_bytes)
        log.info(f"  Schedules (from tables): {len(schedule_items)} items")

    # Fallback 1: text-based schedule regex
    if not schedule_items:
        schedule_items = parse_schedules_from_text(text)
        log.info(f"  Schedules (from text): {len(schedule_items)} items")

    # Fallback 2: single-item bid — extract from Item Category field
    if not schedule_items:
        schedule_items = parse_single_item_from_text(text, pdf_bytes)
        log.info(f"  Schedules (single item fallback): {len(schedule_items)} items")

    # ── STEP 3: Extract per-schedule delivery days ────────────────
    delivery_days_list = []
    if pdf_bytes:
        delivery_days_list = parse_consignee_delivery_days(pdf_bytes)
        log.info(f"  Delivery days per schedule: {delivery_days_list}")

    # If single delivery days available, use it as default for all
    default_lt = delivery_days_list[0] if delivery_days_list else parse_delivery_days(text)

    # ── STEP 4: Build line_items directly from schedule data ──────
    if schedule_items:
        line_items = []
        for i, sched in enumerate(schedule_items):
            lt = delivery_days_list[i] if i < len(delivery_days_list) else default_lt
            line_items.append({
                "part_no":     sched["part_no"],
                "description": sched["description"],
                "enq_part_no": sched["part_no"],
                "qty":         sched["qty"],
                "uom":         "Numbers",
                "lead_time":   lt,     # per-schedule delivery days
            })
        log.info(f"  → Built {len(line_items)} items from schedule tables")

    else:
        # Absolute fallback if no schedule table found at all
        line_items = [{
            "part_no":     bid_number,
            "description": "GeM Bid - manual review required",
            "enq_part_no": bid_number,
            "qty":         "1",
            "uom":         "Numbers",
            "lead_time":   default_lt,
        }]
        log.warning("  → No schedule data found. Single placeholder row.")

    log.info(f"  Total line items: {len(line_items)}")

    # ── STEP 5: Build output rows ────────────────────────────
    rows = _build_output_rows(
        bid_number, customer, due_date, bid_time,
        default_lt, eval_method, remarks,
        line_items, today_str
    )

    return rows
