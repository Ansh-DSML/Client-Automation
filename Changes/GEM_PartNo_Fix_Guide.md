# GEM Parser — Part Number & Description Fix Guide
## Fixes 2 Bugs Confirmed by Live Diagnosis on All 3 PDFs

---

## Root Cause Analysis

Running pdfplumber on **GeM-Bidding-9105321.pdf** (PDF 1) reveals that the
Evaluation Schedules table returns these raw cell values:

```
col[1] = 'Connector, 33505964011, 80 Pin Module Flange\nPlug'
col[1] = 'Connector, 63500004000, 80 Pin Connector Kit'
```

This causes **two separate bugs** in `parsers/gem_parser.py`:

---

## Bug 1 — Wrong CUST PART NO. and Enq. Part No* for PDF 1

**File:** `parsers/gem_parser.py`
**Function:** `_extract_part_no_from_category()` (line ~190)

**What it does now:**
```python
m = re.match(r'^(\d+)', text)
if m:
    return m.group(1).strip()
return text.upper()    # ← FALLS HERE for PDF 1 rows
```

When `text = "Connector, 33505964011, 80 Pin Module Flange\nPlug"`, the regex
`r'^(\d+)'` fails (string starts with a letter). So the function falls through
to `return text.upper()`, returning the **entire description** as the part
number. This is exactly what the Excel shows:

```
CUST PART NO. = "CONNECTOR, 33505964011, 80 PIN MODULE FLANGE\nPLUG"   ← wrong
Enq. Part No* = "CONNECTOR, 33505964011, 80 PIN MODULE FLANGE\nPLUG"   ← wrong
```

PDF 2 and PDF 3 are unaffected because their descriptions start with digits
(`51107159 Crystal...`, `51126961-Circular...`), so the leading-digit match
succeeds.

**The fix:**  After the leading-digit attempt fails, search for the first
long numeric token anywhere in the string. The HAL part number in this format
is always an 8–11 digit number sitting after `"Connector, "`.

---

## Bug 2 — Embedded `\n` in Description and CUST PART NO.

**What it does now:**
`parse_schedules_from_tables()` (line ~245) reads the raw cell text and passes
it straight into `description` and then into `_extract_part_no_from_category()`
without stripping embedded newlines. pdfplumber splits multi-line cells with
`\n`, so the raw value is:

```
'Connector, 33505964011, 80 Pin Module Flange\nPlug'
```

The `\n` survives into the Excel output for both `CUST PART NO.` (via the
fallback return) and `Description`.

PDFs 2 and 3 also have this problem in their descriptions:
```
'51107159 Crystal Oscillator 36.864 Mhz Qt41hc9m\n36.864 Mhz Or Cxo200lmno36m864...'
'51126961-circular Connector, C Size, Wall Mount Receptacle, Square\nFlanze...'
```

**The fix:** Normalise whitespace (collapse all `\s+` including `\n` to a
single space) at the top of `_extract_part_no_from_category()`. Since this
function is called before building the description string in
`parse_schedules_from_tables()`, applying the clean there handles both columns.
Also apply the same cleanup explicitly inside `parse_schedules_from_tables()`
before the description is stored.

---

## No Excel-Inside-PDF Issue

`pdfdetach` returns 0 files for all 3 PDFs. The "Specification Document
Download" link is an external URL, not an embedded file. The parser does not
need to open anything inside the PDF. The existing code paths for BOQ/Excel
attachment already return empty and are harmless.

---

## The 2 Changes to Make

### Change 1 — Fix `_extract_part_no_from_category()`

**Location:** `parsers/gem_parser.py`, the function starting at line ~190.

Replace the **entire function body** with:

```python
def _extract_part_no_from_category(category_text: str) -> str:
    """
    Extract the HAL/GeM part number from an Item Category string.

    Patterns handled:
      'Connector, 33505964011, 80 PIN MODULE FLANGE PLUG'  → 33505964011  (PDF 1)
      '51107159 Crystal Oscillator 36.864 Mhz...'          → 51107159     (PDF 2)
      '51126961-Circular Connector, C Size...'              → 51126961     (PDF 2/3)
      '68838104 Copper Braid HAL-6176...'                   → 68838104     (PDF 2)
    """
    # Normalise whitespace first — pdfplumber puts \n in multi-line cells
    text = re.sub(r'\s+', ' ', category_text).strip()

    # Case A: string starts with digits (PDFs 2 & 3 — most common)
    m = re.match(r'^(\d+)', text)
    if m:
        return m.group(1).strip()

    # Case B: digits are preceded by a non-digit separator (PDF 1 pattern)
    # e.g. "Connector, 33505964011, ..." — find first run of 6+ digits
    m = re.search(r'(?<![A-Za-z\-])(\d{6,})', text)
    if m:
        return m.group(1).strip()

    # Fallback: return cleaned text (should never reach here for GeM PDFs)
    return text.upper()
```

---

### Change 2 — Strip `\n` from description in `parse_schedules_from_tables()`

**Location:** `parsers/gem_parser.py`, inside `parse_schedules_from_tables()`,
in the block where `description` is assigned (lines ~288–297).

There are **three branches** that assign `description`. In all three, wrap the
assignment with `re.sub(r'\s+', ' ', ...).strip()`.

Find this block (the if/elif/else that sets `description` and `qty`):

```python
                        if ncols >= 4:
                            col1 = str(row[1] or "").strip().replace(",", "")
                            if re.match(r'^\d+(\.\d+)?$', col1):
                                description = str(row[2] or "").strip()
                                qty = str(row[3] or "").strip()
                            else:
                                description = str(row[1] or "").strip()
                                qty = str(row[-1] or "").strip()
                        elif ncols == 3:
                            description = str(row[1] or "").strip()
                            qty = str(row[2] or "").strip()
                        else:
                            continue
```

Replace it with:

```python
                        if ncols >= 4:
                            col1 = str(row[1] or "").strip().replace(",", "")
                            if re.match(r'^\d+(\.\d+)?$', col1):
                                description = re.sub(r'\s+', ' ', str(row[2] or "")).strip()
                                qty = str(row[3] or "").strip()
                            else:
                                description = re.sub(r'\s+', ' ', str(row[1] or "")).strip()
                                qty = str(row[-1] or "").strip()
                        elif ncols == 3:
                            description = re.sub(r'\s+', ' ', str(row[1] or "")).strip()
                            qty = str(row[2] or "").strip()
                        else:
                            continue
```

---

## Expected Output After Both Changes

### PDF 1 — GEM/2026/B/7342473

| SN | CUST PART NO. | Description | Enq. Part No* | Qty | L/T |
|---|---|---|---|---|---|
| 1 | `33505964011` | `Connector, 33505964011, 80 Pin Module Flange Plug` | `33505964011` | 20 | 60 |
| 2 | `63500004000` | `Connector, 63500004000, 80 Pin Connector Kit` | `63500004000` | 20 | 60 |

### PDF 2 — GEM/2026/B/7347198

| SN | CUST PART NO. | Description | Qty | L/T |
|---|---|---|---|---|
| 1 | `51107159` | `51107159 Crystal Oscillator 36.864 Mhz Qt41hc9m 36.864 Mhz Or Cxo200lmno36m864 Or Equivalent, Qte` | 86 | 179 |
| 2 | `68838104` | `68838104 Copper Braid Hal-6176 16x24 Sukriti Vidyut Udyog` | 20 | 180 |
| 3 | `51125756` | `51125756 Back Shell P/no: 2m620ms065-nf14 (or) 620ms065-nf14 Make: M/s: Amphenol/souriau/fci/deutsc` | 60 | 180 |

### PDF 3 — GEM/2026/B/7355310

| SN | CUST PART NO. | Description | Qty | L/T |
|---|---|---|---|---|
| 1 | `51126961` | `51126961-circular Connector, C Size, Wall Mount Receptacle, Square Flanze, 4 No's Of 16 Awg Pc Tail` | 405 | 180 |
| 2 | `51106879` | `51106879-back Shell(10 Pin Rect Straight Cable Clamp Type) M85049/38-13w M/s Amphenol` | 322 | 180 |
| 3 | `51106151` | `51106151-pcb Connector 127-33jm1yc Of M/s Amphenol Socapex (a Subsidary Of Amphenol, Formed Afte` | 15 | 180 |
| 4 | `51106152` | `51106152-pcb Connector 127 33 Pf1z Of M/s M/s Amphenol Socapex (a Subsidary Of Amphenol, Formed A` | 13 | 180 |

---

## What Stays Unchanged

Everything else in `gem_parser.py` is correct for these 3 PDFs. No other
function needs to be touched for this fix.

---

## Summary

| Bug | Symptom in Excel | Root Cause | Fix Location |
|---|---|---|---|
| Wrong CUST PART NO. / Enq. Part No* for PDF 1 | Entire description string e.g. `CONNECTOR, 33505964011, 80 PIN MODULE FLANGE\nPLUG` | `_extract_part_no_from_category()` only tried leading-digit match; fell back to full string when description starts with `"Connector, "` | Add fallback `re.search(r'(?<![A-Za-z\-])(\d{6,})', text)` in `_extract_part_no_from_category()` |
| `\n` embedded in Description (all 3 PDFs) | Newline visible in cell e.g. `80 Pin Module Flange\nPlug` | pdfplumber returns `\n` for multi-line table cells; no cleanup applied before storing description | Wrap description assignments in `re.sub(r'\s+', ' ', ...).strip()` in `parse_schedules_from_tables()` |
