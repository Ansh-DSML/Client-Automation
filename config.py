import os
from dotenv import load_dotenv

load_dotenv()

SUPABASE_URL        = os.getenv("SUPABASE_URL")
SUPABASE_SERVICE_KEY = os.getenv("SUPABASE_SERVICE_KEY")
SUPABASE_BUCKET     = os.getenv("SUPABASE_BUCKET", "rfq-pdfs")

GROQ_API_KEY        = os.getenv("GROQ_API_KEY")
GROQ_MODEL          = os.getenv("GROQ_MODEL", "llama-3.3-70b-versatile")

OUTPUT_EXCEL_PATH   = os.getenv("OUTPUT_EXCEL_PATH", r"D:\rfq_pipeline\output\RFQ_Log.xlsx")
GEM_OUTPUT_EXCEL_PATH = os.getenv("GEM_OUTPUT_EXCEL_PATH", r"D:\rfq_pipeline\output\gem_excel_output.xlsx")

# All companies your client deals with.
# Add new ones here as you onboard more parsers.
KNOWN_COMPANIES = ["BEL", "HAL", "DRDO", "BEML", "NAL", "GEM"]
