import pdfplumber, sys, io, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# Get full schedule table from PDF 2
path2 = r'd:\rfq_pipeline\GeM-PDF\GeM-Bidding-9382650.pdf'
pdf2 = pdfplumber.open(path2)

# ALL tables from page 5 (where schedule table is)
print("=== P5 ALL tables PDF 2 ===")
page5 = pdf2.pages[4]  # 0-indexed
for ti, table in enumerate(page5.extract_tables()):
    print(f"\n  T{ti} ({len(table)} rows):")
    for ri, row in enumerate(table):
        sr = [str(c).encode('ascii', errors='replace').decode('ascii') if c else '' for c in row]
        print(f"    R{ri}: {sr}")

# ALL consignee tables from PDF2 (page 6-10)
print("\n=== PDF 2 Consignee tables pages 6-10 ===")
for i in range(5, min(10, len(pdf2.pages))):
    page = pdf2.pages[i]
    for ti, table in enumerate(page.extract_tables()):
        for ri, row in enumerate(table):
            sr = [str(c).encode('ascii', errors='replace').decode('ascii') if c else '' for c in row]
            if any('***' in str(c) for c in row if c) or any('onsignee' in str(c) for c in row if c) or any('Schedule' in str(c) for c in row if c):
                print(f"  P{i+1} T{ti} R{ri}: {sr}")

# Check the Item Category line from PDF 1 - more context
path1 = r'd:\rfq_pipeline\GeM-PDF\GeM-Bidding-8710858.pdf'
pdf1 = pdfplumber.open(path1)
text1 = '\n'.join([p.extract_text() or '' for p in pdf1.pages])
safe1 = text1.encode('ascii', errors='replace').decode('ascii')

print("\n=== PDF 1 Lines around Item Category ===")
lines = safe1.split('\n')
for idx, line in enumerate(lines):
    if 'item categ' in line.lower():
        for j in range(max(0, idx-3), min(len(lines), idx+10)):
            print(f"  L{j}: {lines[j].strip()[:300]}")
        break

# Check the schedule table for PDF 2 - just to see if there are 8 items
print("\n=== PDF 2 Schedule 6-8 check ===")
text2 = '\n'.join([p.extract_text() or '' for p in pdf2.pages])
safe2 = text2.encode('ascii', errors='replace').decode('ascii')
for line in safe2.split('\n'):
    if 'Schedule' in line and any(c.isdigit() for c in line):
        print(f"  {line.strip()[:200]}")

pdf1.close()
pdf2.close()
