"""Setup Supabase: check table existence."""
import json
import urllib.request

SUPABASE_URL = "https://ffbnrypmxzlixefxcdxx.supabase.co"
SERVICE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZmYm5yeXBteHpsaXhlZnhjZHh4Iiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc4MTUyMDc2NCwiZXhwIjoyMDk3MDk2NzY0fQ.W7xYoygO-VBy56B_MkD495RS08U8CC9hmpDruyMi390"

def make_request(url, data=None, method="POST"):
    body = json.dumps(data).encode("utf-8") if data else None
    req = urllib.request.Request(url, data=body, method=method)
    req.add_header("Authorization", f"Bearer {SERVICE_KEY}")
    req.add_header("apikey", SERVICE_KEY)
    req.add_header("Content-Type", "application/json")
    req.add_header("Prefer", "return=representation")
    try:
        with urllib.request.urlopen(req) as resp:
            return resp.status, resp.read().decode()
    except urllib.error.HTTPError as e:
        return e.code, e.read().decode()

# Check if bucket exists
print("Checking storage bucket...")
status, body = make_request(
    f"{SUPABASE_URL}/storage/v1/bucket/rfq-pdfs",
    method="GET"
)
print(f"  Bucket check status: {status}")
if status == 200:
    print("  [OK] Bucket 'rfq-pdfs' exists.")
else:
    print(f"  [WARN] Bucket response: {body}")

# Check if table exists
print("\nChecking rfq_uploads table...")
status, body = make_request(
    f"{SUPABASE_URL}/rest/v1/rfq_uploads?select=id&limit=1",
    method="GET"
)
print(f"  Table check status: {status}")
if status == 200:
    print("  [OK] Table 'rfq_uploads' exists!")
else:
    print(f"  [INFO] Table response: {body}")
    print("  Table may not exist yet. Please create it in Supabase SQL Editor.")
    print("  SQL:")
    print("  CREATE TABLE rfq_uploads (")
    print("      id UUID DEFAULT gen_random_uuid() PRIMARY KEY,")
    print("      filename TEXT NOT NULL,")
    print("      storage_path TEXT NOT NULL,")
    print("      company_detected TEXT DEFAULT NULL,")
    print("      uploaded_at TIMESTAMPTZ DEFAULT NOW(),")
    print("      processed BOOLEAN DEFAULT FALSE,")
    print("      row_count INTEGER DEFAULT 0")
    print("  );")

print("\nDone.")
