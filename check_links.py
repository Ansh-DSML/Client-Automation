"""
Inspect GeM PDFs for hyperlinks, annotations, embedded files - to understand
exactly how 'View File' links inside the PDF work.
"""
import sys, io, re
sys.stdout.reconfigure(encoding='utf-8')

# Try with pdfplumber's underlying pdfminer
import pdfplumber

for fname in ['GeM-Bidding-8710858.pdf', 'GeM-Bidding-9382650.pdf']:
    print(f'\n===== {fname} =====')
    with open(fr'd:\rfq_pipeline\GeM-PDF\{fname}', 'rb') as f:
        raw = f.read()

    # 1. Look for URI patterns in raw bytes
    print('--- URIs in raw bytes ---')
    for m in re.finditer(rb'/URI\s*\(([^)]+)\)', raw):
        print(' ', m.group(1).decode('latin-1', errors='replace'))

    # 2. Look for /F (file) specifications
    for m in re.finditer(rb'/F\s*\(([^)]+)\)', raw):
        val = m.group(1).decode('latin-1', errors='replace')
        if 'gem' in val.lower() or 'http' in val.lower() or '.xls' in val.lower() or '.pdf' in val.lower():
            print('  /F:', val)

    # 3. pdfplumber hyperlinks
    with pdfplumber.open(io.BytesIO(raw)) as pdf:
        for i, page in enumerate(pdf.pages):
            # Check page annotations
            annots = page.annots
            if annots:
                print(f'  Page {i} annots:')
                for a in annots:
                    print(f'    {a}')

    # 4. Look for any http URLs in the raw bytes
    for m in re.finditer(rb'https?://[^\s\x00-\x1f\x80-\xff)>]{10,}', raw):
        url = m.group(0).decode('latin-1', errors='replace')
        if 'gem' in url.lower():
            print(f'  URL: {url}')
