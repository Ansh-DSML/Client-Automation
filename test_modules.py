"""Quick test of all modules."""
from parsers import bel_parser, unknown_parser
print(f"BEL columns: {len(bel_parser.COLUMNS)}")

r = unknown_parser.parse("test")
print(f"Unknown parser remark: {r[0]['Remark']}")

from router import route
print("Router imported OK")

from excel_writer import append_rows
print("Excel writer imported OK")

from extractor import extract_text
print("Extractor imported OK")

from groq_detector import detect_company
print("Groq detector imported OK")

print("\nAll modules validated successfully!")
