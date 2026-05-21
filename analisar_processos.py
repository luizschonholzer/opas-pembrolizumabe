import openpyxl
import pdfplumber
import os
import sys
from datetime import datetime

BASE_DIR = r"C:\Users\Claudio\Downloads\Processos - OPAS"
XLSX_PATH = os.path.join(BASE_DIR, "OPAS_com_links_PDF_BACKUP.xlsx")

COUNTRY_FOLDER = {
    "Argentina": "Argentina",
    "Brazil": "Brasil",
    "Chile": "Chile",
    "Colombia": "Colombia",
    "Ecuador": "Equador",
    "Guatemala": "Guatemala",
    "Costa Rica": "Costa Rica",
    "Uruguay": "Uruguai",
    "Paraguay": "Paraguay",
    "Peru": "Peru",
    "Bolivia": "Bolivia",
    "Venezuela": "Venezuela",
    "Guyana": "Guyana",
    "Canada": "Canada",
}

def find_pdf(pdf_name, country):
    if not pdf_name or str(pdf_name).strip() in ("", "None"):
        return None
    pdf_name = str(pdf_name).strip()
    folder = COUNTRY_FOLDER.get(country, country)
    country_dir = os.path.join(BASE_DIR, folder)
    if os.path.exists(country_dir):
        for root, dirs, files in os.walk(country_dir):
            for f in files:
                if f.strip() == pdf_name:
                    return os.path.join(root, f)
    direct = os.path.join(BASE_DIR, pdf_name)
    if os.path.exists(direct):
        return direct
    return None

def extract_pdf_text(pdf_path, max_chars=12000):
    try:
        text = ""
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[:15]:
                t = page.extract_text()
                if t:
                    text += t + "\n"
                if len(text) > max_chars:
                    break
        return text[:max_chars].strip()
    except Exception as e:
        return f"[ERRO AO LER PDF: {e}]"

def val(v):
    if v is None:
        return ""
    return str(v).strip()

def load_rows():
    wb = openpyxl.load_workbook(XLSX_PATH, data_only=True)
    ws = wb["Sheet1"]
    headers = []
    rows = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            headers = [str(c) if c is not None else "" for c in row]
            continue
        if all(c is None or str(c).strip() == "" for c in row):
            continue
        row_dict = {}
        for j, h in enumerate(headers):
            row_dict[h] = row[j] if j < len(row) else None
        rows.append(row_dict)
    return rows

def is_skippable(row):
    skip_markers = ["procedural issue", "not about pembrolizumab", "out of scope",
                    "procedural matter", "not found"]
    disease = val(row.get("Disease", "")).lower()
    for m in skip_markers:
        if m in disease:
            return True, disease
    return False, disease

def format_date(d):
    if d is None:
        return "não localizado nos autos"
    s = str(d)
    if s.isdigit():
        # Excel serial date
        try:
            from datetime import date
            delta = date(1899, 12, 30)
            from datetime import timedelta
            return (delta + timedelta(days=int(s))).strftime("%d/%m/%Y")
        except:
            return s
    return s[:10] if len(s) > 10 else s

def analyze_row(idx, row, pdf_text):
    country = val(row.get("Country"))
    case_num = val(row.get("Case Number"))
    court = val(row.get("Court"))
    date = format_date(row.get("Date"))
    defendant = val(row.get("Defendant"))
    other_meds = val(row.get("Other Medications Besides Pembrolizumab"))
    incorporated = val(row.get("Incorporated"))
    within_protocol = val(row.get("Within Protocol"))
    registration = val(row.get("Registration for Patient's Disease"))
    disease_n = val(row.get("Disease"))
    clinical_o = val(row.get("Further clinical detalis"))
    legal_claim = val(row.get("Legal Basis of the Claim"))
    legal_denial = val(row.get("Legal Basis of the Denial"))
    legal_decision = val(row.get("Legal Basis of the Decision"))
    evidence = val(row.get("Scientific Evidence Supporting Decision"))
    granted = val(row.get("Pembrolizumab Granted"))
    notes = val(row.get("Notes"))
    esmo = val(row.get("ESMO Score"))
    who_eml = val(row.get("WHO EML "))
    pdf_link = val(row.get("PDF Link"))

    # Determine decision result
    if val(granted) == "1" or val(granted).lower() == "yes":
        resultado = "DEFERIDO"
    elif val(granted) == "0" or val(granted).lower() == "no":
        resultado = "INDEFERIDO"
    elif val(granted).lower() in ["", "none"]:
        resultado = "não localizado nos autos"
    else:
        resultado = val(granted)

    # Check column N adequacy
    n_status = "adequada"
    n_obs = ""
    if not disease_n or disease_n.lower() in ["ni", "none", ""]:
        n_status = "incompleta"
        n_obs = "Sem diagnóstico registrado"
    elif len(disease_n) < 10:
        n_status = "incompleta"
        n_obs = "Diagnóstico muito sucinto"

    # Check column O adequacy
    o_status = "adequada"
    o_obs = ""
    if not clinical_o or clinical_o.lower() in ["ni", "none", ""]:
        o_status = "incompleta"
        o_obs = "Sem detalhes clínicos registrados"
    elif len(clinical_o) < 30:
        o_status = "incompleta"
        o_obs = "Detalhes clínicos muito sucintos"

    # Build corrections from PDF text
    corrections = []
    pdf_excerpt = ""
    if pdf_text and not pdf_text.startswith("[ERRO"):
        pdf_excerpt = pdf_text[:3000]
        if n_status == "incompleta":
            corrections.append(f"Coluna N (Disease): complementar com diagnóstico extraído do PDF.")
        if o_status == "incompleta":
            corrections.append(f"Coluna O (Further clinical details): complementar com detalhes clínicos do PDF.")

    lines = []
    lines.append(f"\n{'='*70}")
    lines.append(f"PROCESSO #{idx}: {country} — {case_num}")
    lines.append(f"{'='*70}")
    lines.append(f"Tribunal: {court}")
    lines.append(f"Data: {date}")
    lines.append(f"Réu/Demandado: {defendant}")
    lines.append(f"PDF: {pdf_link if pdf_link else 'não disponível'}")

    lines.append(f"\nRESUMO CLÍNICO:")
    lines.append(f"  Diagnóstico (col. N):       {disease_n if disease_n else 'não localizado nos autos'}")
    lines.append(f"  Detalhes clínicos (col. O): {clinical_o if clinical_o else 'não localizado nos autos'}")
    lines.append(f"  Outros medicamentos:         {other_meds if other_meds else 'não localizado nos autos'}")
    lines.append(f"  Incorporado:                 {incorporated if incorporated else 'não localizado nos autos'}")
    lines.append(f"  Dentro do protocolo:         {within_protocol if within_protocol else 'não localizado nos autos'}")
    lines.append(f"  Registro para a doença:      {registration if registration else 'não localizado nos autos'}")
    lines.append(f"  ESMO Score:                  {esmo if esmo else 'não localizado nos autos'}")
    lines.append(f"  WHO EML:                     {who_eml if who_eml else 'não localizado nos autos'}")

    lines.append(f"\nANÁLISE DA DECISÃO:")
    lines.append(f"  Resultado:        {resultado}")
    lines.append(f"  Base jurídica (demandante): {legal_claim if legal_claim else 'não localizado nos autos'}")
    lines.append(f"  Base jurídica (negativa):   {legal_denial if legal_denial else 'não localizado nos autos'}")
    lines.append(f"  Fundamentação decisão:      {legal_decision if legal_decision else 'não localizado nos autos'}")
    lines.append(f"  Evidência científica:       {evidence if evidence else 'não localizado nos autos'}")
    lines.append(f"  Observações:                {notes if notes else 'não localizado nos autos'}")

    lines.append(f"\nCONFERÊNCIA PLANILHA:")
    lines.append(f"  Coluna N: {n_status}{' — ' + n_obs if n_obs else ''}")
    lines.append(f"  Coluna O: {o_status}{' — ' + o_obs if o_obs else ''}")
    overall = "OK" if n_status == "adequada" and o_status == "adequada" else "REVISAR"
    lines.append(f"  Status geral: {overall}")

    if pdf_text and not pdf_text.startswith("[ERRO"):
        lines.append(f"\nEXCERTO DO PDF (primeiros 3000 caracteres):")
        lines.append("-" * 50)
        lines.append(pdf_excerpt)
        lines.append("-" * 50)

    lines.append(f"\nCORREÇÕES SUGERIDAS:")
    if corrections:
        for c in corrections:
            lines.append(f"  - {c}")
    else:
        lines.append("  Nenhuma correção necessária com base nos dados disponíveis.")

    return "\n".join(lines)

def main():
    start = int(sys.argv[1]) if len(sys.argv) > 1 else 1
    end = int(sys.argv[2]) if len(sys.argv) > 2 else None

    rows = load_rows()
    total = len(rows)
    print(f"Total de processos carregados: {total}")

    if end is None:
        end = total

    processed = 0
    for i, row in enumerate(rows):
        idx = i + 1
        if idx < start or idx > end:
            continue

        country = val(row.get("Country"))
        case_num = val(row.get("Case Number"))
        pdf_link = val(row.get("PDF Link"))

        skippable, disease = is_skippable(row)
        if skippable:
            print(f"\n{'='*70}")
            print(f"PROCESSO #{idx}: {country} — {case_num}")
            print(f"  [IGNORADO] Processo fora do escopo: {disease}")
            continue

        # Find and read PDF
        pdf_path = find_pdf(pdf_link, country)
        pdf_text = ""
        if pdf_path:
            print(f"\n[Lendo PDF {idx}/{total}: {os.path.basename(pdf_path)}]", file=sys.stderr)
            pdf_text = extract_pdf_text(pdf_path)
        else:
            if pdf_link:
                print(f"\n[PDF não encontrado localmente: {pdf_link}]", file=sys.stderr)

        report = analyze_row(idx, row, pdf_text)
        print(report)
        processed += 1

    print(f"\n\nTotal analisados: {processed} de {end - start + 1} no intervalo #{start}-#{end}")

if __name__ == "__main__":
    main()
