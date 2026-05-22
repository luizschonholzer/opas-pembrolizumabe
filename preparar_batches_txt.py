"""
Regenerates batch files pointing to extracted .txt files instead of raw PDFs.
Run AFTER extrair_textos_pdfs.py completes.
"""
import json, os
from pathlib import Path
from collections import defaultdict

BASE_DIR    = Path(__file__).parent
TXT_DIR     = BASE_DIR / "textos_pdfs"
COUNTRY_FOLDER = {
    "Argentina":  "Argentina", "Brazil": "Brasil",
    "Chile":      "Chile",     "Colombia": "Colombia",
    "Ecuador":    "Equador",   "Guatemala": "Guatemala",
    "Costa Rica": "Costa Rica","Uruguay": "Uruguai",
}

def safe_name(country, pdf_name):
    s = f"{country}__{pdf_name}"
    for c in r'\/:*?"<>|':
        s = s.replace(c, "_")
    return s + ".txt"

rows = json.loads((BASE_DIR / "rows_to_process.json").read_text(encoding="utf-8"))

# Group rows by (country, pdf) → include txt_path
pdf_groups = defaultdict(list)
missing_txt = []

for r in rows:
    txt_name = safe_name(r["country"], r["pdf"])
    txt_path = TXT_DIR / txt_name
    if txt_path.exists():
        key = (r["country"], r["pdf"])
        pdf_groups[key].append({
            "row":      r["row"],
            "case":     r["case"],
            "N":        r["N"],
            "O":        r["O"],
            "txt_path": str(txt_path),
        })
    else:
        missing_txt.append(r)

print(f"Unique PDFs with text file: {len(pdf_groups)}")
print(f"Missing text files:         {len(missing_txt)}")

tasks = []
for (country, pdf_name), row_list in pdf_groups.items():
    tasks.append({
        "country":  country,
        "pdf_name": pdf_name,
        "txt_path": row_list[0]["txt_path"],
        "rows":     row_list,
    })

BATCH_SIZE = 12
batches = [tasks[i:i+BATCH_SIZE] for i in range(0, len(tasks), BATCH_SIZE)]

print(f"Total tasks: {len(tasks)}")
print(f"Batches:     {len(batches)} (up to {BATCH_SIZE} each)")

for idx, batch in enumerate(batches):
    out = BASE_DIR / f"batch_txt_{idx+1:02d}.json"
    out.write_text(json.dumps(batch, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"  {out.name}  ({len(batch)} tasks)")

if missing_txt:
    (BASE_DIR / "txt_missing.json").write_text(
        json.dumps(missing_txt, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"\nMissing files saved to txt_missing.json")

print("\nBatch files ready.")
