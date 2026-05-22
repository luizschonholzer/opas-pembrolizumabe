"""
Creates batch task files for parallel agent PDF analysis.
Groups rows by unique PDF (per country), then splits into balanced batches.
"""
import json, os
from pathlib import Path
from collections import defaultdict

BASE_DIR = Path(r"C:\Users\Claudio\Downloads\Processos - OPAS")
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
                return str(Path(root) / fname)
    direct = BASE_DIR / pdf_name_clean
    if direct.exists():
        return str(direct)
    return None

rows = json.loads((BASE_DIR / "rows_to_process.json").read_text(encoding="utf-8"))

# Group rows by (country, pdf) -> list of row metadata
pdf_groups = defaultdict(list)
not_found = []

for r in rows:
    pdf_path = find_pdf(r["pdf"], r["country"])
    if pdf_path:
        key = (r["country"], r["pdf"])
        pdf_groups[key].append({
            "row":      r["row"],
            "case":     r["case"],
            "N":        r["N"],
            "O":        r["O"],
            "pdf_path": pdf_path,
        })
    else:
        not_found.append(r)

print(f"Unique PDFs found on disk:   {len(pdf_groups)}")
print(f"Rows with PDF not on disk:   {len(not_found)}")

# Build flat list of unique PDF tasks
tasks = []
for (country, pdf_name), row_list in pdf_groups.items():
    tasks.append({
        "country":  country,
        "pdf_name": pdf_name,
        "pdf_path": row_list[0]["pdf_path"],   # same path for all rows
        "rows":     row_list,
    })

# Split into ~8 balanced batches
BATCH_SIZE = 12
batches = [tasks[i:i+BATCH_SIZE] for i in range(0, len(tasks), BATCH_SIZE)]

print(f"Total tasks:  {len(tasks)}")
print(f"Batches:      {len(batches)} (up to {BATCH_SIZE} PDFs each)")

# Save batch files
for idx, batch in enumerate(batches):
    batch_file = BASE_DIR / f"batch_{idx+1:02d}.json"
    with open(batch_file, "w", encoding="utf-8") as f:
        json.dump(batch, f, ensure_ascii=False, indent=2)
    print(f"  {batch_file.name}  ({len(batch)} PDFs)")

# Save not-found list
if not_found:
    nf_file = BASE_DIR / "pdfs_not_found.json"
    with open(nf_file, "w", encoding="utf-8") as f:
        json.dump(not_found, f, ensure_ascii=False, indent=2)
    print(f"\nPDFs not found saved to: {nf_file.name}")
    for r in not_found:
        print(f"  Row {r['row']} | {r['country']} | {r['pdf']}")

print("\nDone. Batch files ready.")
