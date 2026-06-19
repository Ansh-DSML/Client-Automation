import requests, sys, io
import pandas as pd
sys.stdout.reconfigure(encoding='utf-8')

csv_url = 'https://mkp.gem.gov.in/uploaded_documents/51/16/877/OrderItem/BoqLineItemsDocument/2025/12/13/boq_item_593-27-2_2025-12-13-12-06-30_fcb6511fca5e2cac5c908711935553c6.csv'
headers = {'User-Agent': 'Mozilla/5.0'}

r = requests.get(csv_url, timeout=15, headers=headers)
print(f'Status: {r.status_code}')
df = pd.read_csv(io.StringIO(r.text))
print(f'Columns: {list(df.columns)}')
print(df.to_string())
