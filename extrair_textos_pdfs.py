"""
Extracts text from all PDFs listed in rows_to_process.json using pdfplumber.
Saves each unique PDF's text to a .txt file in the 'textos_pdfs' subfolder.
"""
import json, os
from pathlib import Path
from collections import defaultdict

try:
    import pdfplumber
except ImportError:
    print("ERROR: pdfplumber not installed. Run: pip install pdfplumber")
    raise

try:
    from docx import Document as DocxDocument
    DOCX_OK = True
except ImportError:
    DOCX_OK = False
    print("NOTE: python-docx not installed — DOCX files will be skipped.")

BASE_DIR   = Path(__file__).parent
OUTPUT_DIR = BASE_DIR / "textos_pdfs"
OUTPUT_DIR.mkdir(exist_ok=True)

COUNTRY_FOLDER = {
    "Argentina":  "Argentina",
    "Brazil":     "Brasil",
    "Chile":      "Chile",
    "Colombia":   "Colombia",
    "Ecuador":    "Equador",
    "Guatemala":  "Guatemala",
    "Costa Rica": "Costa Rica",
    "Uruguay":    "Uruguai",
    "Paraguay":   "Paraguay",
    "Peru":       "Peru",
}

def find_pdf(pdf_name, country):
    folder_name = COUNTRY_FOLDER.get(country, country)
    country_dir = BASE_DIR / folder_name
    if not country_dir.exists():
        for d in BASE_DIR.iterdir():
            if d.is_dir() and d.name.lower() == folder_name.lower():
                country_dir = d
                break
        else:
            return None
    pdf_name_clean = pdf_name.strip()
    for root, _, files in os.walk(country_dir):
        for fname in files:
            if fname.strip() == pdf_name_clean:
                return Path(root) / fname
    direct = BASE_DIR / pdf_name_clean
    if direct.exists():
        return direct
    return None

def extract_pdf_text(pdf_path: Path) -> str:
    texts = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                t = page.extract_text()
                if t and t.strip():
                    texts.append(f"--- Page {i+1} ---\n{t.strip()}")
        return "\n\n".join(texts) if texts else "[No extractable text found in PDF]"
    except Exception as e:
        return f"[PDF extraction error: {e}]"

def extract_docx_text(docx_path: Path) -> str:
    if not DOCX_OK:
        return "[DOCX extraction skipped — python-docx not installed]"
    try:
        doc = DocxDocument(str(docx_path))
        paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
        return "\n\n".join(paragraphs) if paragraphs else "[No text found in DOCX]"
    except Exception as e:
        return f"[DOCX extraction error: {e}]"

# Load rows
rows = json.loads((BASE_DIR / "rows_to_process.json").read_text(encoding="utf-8"))

# Deduplicate by (country, pdf_name)
seen = set()
tasks = []
for r in rows:
    key = (r["country"], r["pdf"])
    if key not in seen:
        seen.add(key)
        tasks.append(r)

print(f"Unique PDFs to extract: {len(tasks)}")
ok, failed = 0, 0

# Safe filename for output
def safe_name(country, pdf_name):
    s = f"{country}__{pdf_name}"
    for c in r'\/:*?"<>|':
        s = s.replace(c, "_")
    return s + ".txt"

for r in tasks:
    pdf_path = find_pdf(r["pdf"], r["country"])
    if not pdf_path:
        print(f"  NOT FOUND: {r['country']} / {r['pdf']}")
        failed += 1
        continue

    out_name = safe_name(r["country"], r["pdf"])
    out_path = OUTPUT_DIR / out_name

    suffix = pdf_path.suffix.lower()
    if suffix == ".pdf":
        text = extract_pdf_text(pdf_path)
    elif suffix == ".docx":
        text = extract_docx_text(pdf_path)
    else:
        text = f"[Unsupported file type: {suffix}]"

    header = (
        f"DOCUMENT: {r['pdf']}\n"
        f"COUNTRY: {r['country']}\n"
        f"ORIGINAL PATH: {pdf_path}\n"
        f"{'=' * 60}\n\n"
    )
    out_path.write_text(header + text, encoding="utf-8")

    preview = text[:80].replace("\n", " ")
    status = "OK" if not text.startswith("[") else "WARN"
    print(f"  [{status}] {r['country']:<12} | {r['pdf']:<45} | {preview}...")
    ok += 1

print(f"\nExtracted: {ok} files  |  Failed: {failed}")
print(f"Output folder: {OUTPUT_DIR}")
