import json
from groq import Groq
from config import GROQ_API_KEY, GROQ_MODEL, KNOWN_COMPANIES

_groq_client: Groq = None

def get_groq_client() -> Groq:
    global _groq_client
    if _groq_client is None:
        _groq_client = Groq(api_key=GROQ_API_KEY)
    return _groq_client


def detect_company(full_text: str) -> dict:
    """
    Sends the first 800 characters of the PDF text to Groq.
    Returns a dict with this exact structure:
    {
        "company_code": "BEL",
        "company_full_name": "Bharat Electronics Limited",
        "confidence": "high"
    }

    company_code will be one of: BEL, HAL, DRDO, BEML, NAL, UNKNOWN
    confidence will be one of: high, medium, low
    """
    client = get_groq_client()

    # Only send the first 800 chars — the company name is always in the header
    header_text = full_text[:800]

    system_prompt = f"""You are a document classifier for Indian defence and aerospace procurement PDFs.

Your job is to identify which company issued this RFQ (Request for Quotation) document.

Known companies and their codes:
- Bharat Electronics Limited → BEL
- Hindustan Aeronautics Limited → HAL
- Defence Research and Development Organisation → DRDO
- Bharat Earth Movers Limited → BEML
- National Aerospace Laboratories → NAL
- Government e Marketplace (GeM Portal) → GEM

IMPORTANT: GeM portal PDFs often contain garbled text like "(cid:1)" but will have a
bid number pattern starting with "GEM/" (e.g. "GEM/2025/B/6998572"). If you see "GEM/"
anywhere in the text, classify it as GEM with high confidence.

If the company does not match any of the above, use code: UNKNOWN
(Note: GEM is specifically for GeM portal bids identified by GEM/ bid numbers.)

You must respond with ONLY a valid JSON object. No explanation. No markdown. No extra text.

Required format:
{{
  "company_code": "BEL",
  "company_full_name": "Bharat Electronics Limited",
  "confidence": "high"
}}

confidence must be one of: high, medium, low"""

    user_prompt = f"""Identify the company from this RFQ document header:

{header_text}"""

    response = client.chat.completions.create(
        model=GROQ_MODEL,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user",   "content": user_prompt},
        ],
        temperature=0,           # deterministic — we want the same answer every time
        max_tokens=150,          # response is tiny
        response_format={"type": "json_object"},  # forces JSON output
    )

    raw = response.choices[0].message.content.strip()

    try:
        result = json.loads(raw)
    except json.JSONDecodeError:
        # Defensive fallback: if somehow JSON is malformed
        result = {
            "company_code": "UNKNOWN",
            "company_full_name": "Unknown",
            "confidence": "low",
        }

    # Sanitise — make sure company_code is always in our known list
    if result.get("company_code") not in KNOWN_COMPANIES + ["UNKNOWN"]:
        result["company_code"] = "UNKNOWN"

    return result
