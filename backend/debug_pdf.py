import pdfplumber
import re

path = r"uploads\CWI_Final Report_Kotak Mahindra Bank Ltd_Dombivli (W).pdf"

print("=== PLAIN TEXT (first 3 pages) ===")
with pdfplumber.open(path) as pdf:
    for i, page in enumerate(pdf.pages[:3]):
        print(f"\n--- PAGE {i+1} PLAIN TEXT ---")
        t = page.extract_text()
        print(repr(t[:2000]) if t else "NONE")

        print(f"\n--- PAGE {i+1} TABLES ---")
        tables = page.extract_tables()
        for ti, table in enumerate(tables):
            print(f"  Table {ti}:")
            for row in table:
                print(f"    {row}")

        print(f"\n--- PAGE {i+1} WORDS (first 50) ---")
        words = page.extract_words()
        print(" | ".join(w['text'] for w in words[:50]))