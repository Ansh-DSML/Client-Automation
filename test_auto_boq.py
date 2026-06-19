"""
End-to-end test: verify auto-download of BOQ CSV from PDF annotations.
"""
import sys, io
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, r'd:\rfq_pipeline')

import pdfplumber
from parsers import gem_parser

for fname in ['GeM-Bidding-8710858.pdf', 'GeM-Bidding-9382650.pdf']:
    print(f'\n===== {fname} =====')
    with open(fr'd:\rfq_pipeline\GeM-PDF\{fname}', 'rb') as f:
        pdf_bytes = f.read()
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        text = '\n'.join(p.extract_text() or '' for p in pdf.pages)

    rows = gem_parser.parse(text, pdf_bytes=pdf_bytes)
    print(f'Total rows: {len(rows)}')
    for r in rows:
        print(f"  SN={r['SN']}  PART={str(r['CUST PART NO.'])[:22]:<22}  UOM={r['UOM']}")
