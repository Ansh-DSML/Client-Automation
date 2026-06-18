import re
from datetime import datetime, timedelta

COLUMNS = [
    "Enquiry", "Cust", "RFQ D", "Due D", "Due Time",
    "Delivery Time", "Status", "SN", "CPN", "Description",
    "Enq. Part No*", "Enq. Mfg", "Qty", "Unit",
    "ITEM/VALUE.(EVALUATION)", "TP", "Remark", "ACT.Due DT"
]

def _empty_row() -> dict:
    return {col: "NA" for col in COLUMNS}


def _get_item_value(text: str, cust_code: str) -> str:
    code = cust_code.upper()
    if "N6" in code:
        m = re.search(
            r"(Year of manufacturing[^.]*?within\s+\d+\s+years?[^.]*\.)",
            text, re.IGNORECASE
        )
        if m:
            return m.group(1).strip()
    if "CH1" in code:
        m = re.search(
            r"(MANUFACTURING DATE CODE[^.\n]*?WITHIN\s+\d+\s+YEARS?[^.\n]*(?:IS MANDATORY|ORDER)[^.\n]*)",
            text, re.IGNORECASE
        )
        if m:
            return m.group(1).strip()
    return "N/A"


def parse(text: str) -> list[dict]:
    """
    Accepts full extracted text from a BEL RFQ PDF.
    Returns a list of dicts, one per line item in Bid Details.
    Each dict has exactly the 18 keys defined in COLUMNS.
    """

    # 1. Enquiry number
    enq_m = re.search(r"RFx number\s+(\d+)", text)
    enquiry = enq_m.group(1) if enq_m else "NA"

    # 2. Customer code — derived from Description field
    desc_m = re.search(r"Description:\s*(\S+)", text)
    cust = "BEL"
    if desc_m:
        parts = desc_m.group(1).split("/")
        cust = "BEL-" + parts[1] if len(parts) > 1 else "BEL-" + parts[0]

    item_value = _get_item_value(text, cust)
    rfq_d = datetime.today()

    # 3. Submission period — always use the last date (end date)
    sub_line = re.search(r"Submission period:\s*(.+)", text, re.IGNORECASE)
    due_d = due_time = act_due = None
    if sub_line:
        all_dates = re.findall(r"(\d{2}\.\d{2}\.\d{4})\s+(\d{2}:\d{2})", sub_line.group(1))
        if all_dates:
            last_date, last_time = all_dates[-1]
            sub_date = datetime.strptime(last_date, "%d.%m.%Y")
            due_d    = sub_date - timedelta(days=3)
            t        = datetime.strptime(last_time, "%H:%M")
            due_time = t.strftime("%I:%M %p").lstrip("0")
            act_due  = sub_date

    # 4. Delivery time
    lt_m = re.search(
        r"(?:OUR\s+REQUIRED\s+)?DELIVERY\s+SCHDULE?\s*:?\s*([\d\-–]+\s*(?:WEEKS?|DAYS?))",
        text, re.IGNORECASE
    )
    delivery_time = "NA"
    if lt_m:
        lt_raw  = lt_m.group(1).strip()
        range_m = re.match(r"(\d+)[\-–](\d+)\s*DAYS?", lt_raw, re.IGNORECASE)
        days_m  = re.match(r"(\d+)\s*DAYS?",            lt_raw, re.IGNORECASE)
        weeks_m = re.match(r"([\d\-–]+)\s*WEEKS?",      lt_raw, re.IGNORECASE)
        if range_m:
            avg = (int(range_m.group(1)) + int(range_m.group(2))) / 2
            delivery_time = f"{round(avg / 7)} WEEKS"
        elif days_m:
            delivery_time = f"{round(int(days_m.group(1)) / 7)} WEEKS"
        elif weeks_m:
            delivery_time = weeks_m.group(1) + " WEEKS"

    # 5. Bid Details table
    NOISE = re.compile(
        r"^(?:Item\s+Material|Qty/Unit|Quantity\s*$|Page\s+\d|Date\s*:|"
        r"Bid Invitation|Product no\.|info@|nklamin@|orantselectro@|"
        r"sales\.|support\.|\d{7,10}$)",
        re.IGNORECASE
    )
    in_bid    = False
    bid_lines = []
    for line in text.splitlines():
        s = line.strip()
        if not s:
            continue
        if re.search(r"^Bid Details$", s):
            in_bid = True
            continue
        if in_bid:
            if NOISE.match(s):
                continue
            bid_lines.append(s)

    rows = []
    i = 0
    while i < len(bid_lines):
        item_m = re.match(
            r"^(\d+)\s+(\d{8,})\s+(.+?)\s+([\d,\.]+)\s+([A-Z]+)$",
            bid_lines[i]
        )
        if item_m:
            sn      = int(item_m.group(1)) // 10
            cpn     = item_m.group(2)
            desc    = item_m.group(3).strip()
            qty_raw = item_m.group(4).replace(",", "")
            try:
                qty = float(qty_raw)
                qty = int(qty) if qty == int(qty) else qty
            except Exception:
                qty = qty_raw
            unit = item_m.group(5)

            mfg_list, mpn_list = [], []
            while i + 1 < len(bid_lines):
                nxt = bid_lines[i + 1]
                if re.match(r"^\d+\s+\d{8,}", nxt):
                    break
                if "-" not in nxt and "/" not in nxt:
                    break
                nxt_n   = re.sub(r",\s*-", "-", nxt)
                split_m = re.match(r"^(.*?[A-Z\)\/\.])\\s*-\\s*(.+)$", nxt_n)
                if split_m:
                    mfg_list.append(split_m.group(1).strip().rstrip(",/").strip())
                    mpn_list.append(split_m.group(2).strip())
                else:
                    break
                i += 1

            rows.append({
                "Enquiry":                 enquiry,
                "Cust":                    cust,
                "RFQ D":                   rfq_d,
                "Due D":                   due_d,
                "Due Time":                due_time or "NA",
                "Delivery Time":           delivery_time,
                "Status":                  "Working",
                "SN":                      sn,
                "CPN":                     cpn,
                "Description":             desc,
                "Enq. Part No*":           " // ".join(mpn_list) if mpn_list else "NA",
                "Enq. Mfg":                " // ".join(mfg_list) if mfg_list else "NA",
                "Qty":                     qty,
                "Unit":                    unit,
                "ITEM/VALUE.(EVALUATION)": item_value,
                "TP":                      "",
                "Remark":                  "N/A",
                "ACT.Due DT":              act_due,
            })
        i += 1

    if not rows:
        fallback = _empty_row()
        fallback.update({
            "Enquiry": enquiry, "Cust": cust, "RFQ D": rfq_d,
            "Due D": due_d, "Due Time": due_time or "NA",
            "Delivery Time": delivery_time, "Status": "Working",
            "ITEM/VALUE.(EVALUATION)": item_value,
            "Remark": "N/A", "ACT.Due DT": act_due, "TP": "",
        })
        rows = [fallback]

    return rows
