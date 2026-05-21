# -*- coding: utf-8 -*-
import openpyxl
import pdfplumber
import os
import sys
import re
from datetime import date, timedelta

BASE_DIR = r"C:\Users\Claudio\Downloads\Processos - OPAS"
XLSX_PATH = os.path.join(BASE_DIR, "OPAS_com_links_PDF_BACKUP.xlsx")
OUTPUT_FILE = os.path.join(BASE_DIR, "RELATORIO_ANALISE_PROCESSOS.txt")

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

SKIP_MARKERS = [
    "procedural issue", "not about pembrolizumab", "out of scope",
    "procedural matter", "not found", "não é sobre pembro",
    "not assessable", "membership continuation matter"
]

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

def extract_pdf_text(pdf_path, max_pages=20):
    try:
        text = ""
        with pdfplumber.open(pdf_path) as pdf:
            n = min(max_pages, len(pdf.pages))
            for page in pdf.pages[:n]:
                t = page.extract_text()
                if t:
                    text += t + "\n"
        return text.strip()
    except Exception as e:
        return f"[ERRO AO LER PDF: {e}]"

def val(v):
    if v is None:
        return ""
    s = str(v).strip()
    if s.lower() in ("none", "nan"):
        return ""
    return s

def is_skip(row):
    disease = val(row.get("Disease", "")).lower()
    notes_all = " ".join([
        val(row.get("Disease", "")),
        val(row.get("Further clinical detalis", "")),
        val(row.get("Pembrolizumab Granted", "")),
        val(row.get("Defendant", "")),
    ]).lower()
    for m in SKIP_MARKERS:
        if m in notes_all:
            return True, disease
    return False, disease

def format_date(d):
    if d is None:
        return "não localizado nos autos"
    s = str(d)
    if s.isdigit() and int(s) > 10000:
        try:
            return (date(1899, 12, 30) + timedelta(days=int(s))).strftime("%d/%m/%Y")
        except:
            return s
    if "00:00:00" in s:
        return s[:10]
    return s

# -------- PDF clinical extraction helpers --------

def find_pattern(text, patterns):
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE | re.DOTALL)
        if m:
            snippet = m.group(0)[:300].strip()
            return snippet
    return "não localizado nos autos"

def extract_age(text):
    patterns = [
        r"(\d{1,3})\s*(años?|años de edad|year[s]? old|age[d]?\s*\d|anos? de idade|anos?)\b",
        r"\b(age|edad|idade)\s*[:\-]?\s*(\d{1,3})",
        r"paciente\s+de\s+(\d{1,3})\s+años?",
        r"(\d{1,3})\s*-year-old",
        r"\((\d{2,3})\s*años?\)",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            groups = [g for g in m.groups() if g and g.isdigit()]
            if groups:
                age = int(groups[0])
                if 1 <= age <= 120:
                    return str(age) + " anos"
    return "não localizado nos autos"

def extract_cid(text):
    patterns = [
        r"(?:CID|ICD|C\.I\.D\.?|CIE)[\s\-:]*(C\d{2}(?:\.\d{1,2})?)",
        r"\b(C\d{2}\.\d{1,2})\b",
        r"(?:código|code|código CID)[:\s]*(C\d{2}(?:\.\d{1,2})?)",
    ]
    found = []
    for pat in patterns:
        for m in re.finditer(pat, text, re.IGNORECASE):
            cid = m.group(1) if m.lastindex >= 1 else m.group(0)
            if re.match(r"C\d{2}", cid):
                found.append(cid.upper())
    if found:
        return ", ".join(sorted(set(found)))
    return "não localizado nos autos"

def extract_staging(text):
    patterns = [
        r"(?:stage|staging|estadio|etapa|estadiamento)[:\s]*([IVX]{1,4}[A-C]?(?:\s*[A-D])?)",
        r"(?:TNM|pT\d[a-d]?\s*(?:pN|N)\d[a-d]?\s*(?:M|pM)\d)",
        r"(?:Etapa|Stage|Estadio)\s+([IVX]{1,4}[A-C]?)",
        r"\b(pT\d[a-d]?\s*(?:pN|N)\d[a-d]?\s*(?:M|pM)\d[a-d]?)\b",
    ]
    return find_pattern(text, patterns)

def extract_ecog(text):
    m = re.search(r"ECOG\s*[:\-]?\s*(\d)", text, re.IGNORECASE)
    if m:
        return "ECOG " + m.group(1)
    return "não localizado nos autos"

def extract_prior_treatments(text):
    keywords = [
        "chemotherapy", "quimioterapia", "radiotherapy", "radioterapia",
        "surgery", "cirugía", "cirurgia", "nephrectomy", "nefrectomia",
        "ABVD", "FOLFOX", "XELOX", "carboplatino", "paclitaxel", "cisplatino",
        "gemcitabine", "bevacizumab", "carboplatin", "docetaxel",
        "treatment line", "línea de tratamiento", "primeira linha", "segunda linha",
        "first.?line", "second.?line"
    ]
    found = []
    for kw in keywords:
        m = re.search(r"[^.]*" + kw + r"[^.]*\.", text, re.IGNORECASE)
        if m:
            snippet = m.group(0)[:200].strip()
            if snippet not in found:
                found.append(snippet)
        if len(found) >= 3:
            break
    if found:
        return " | ".join(found[:3])
    return "não localizado nos autos"

def extract_metastasis(text):
    patterns = [
        r"metástas[ei][s]?[^.]*\.",
        r"metásta[^.]*\.",
        r"metastati[^.]*\.",
        r"metástases[^.]*\.",
        r"disseminação[^.]*\.",
        r"spreads?\s+to[^.]*\.",
    ]
    return find_pattern(text, patterns)

def extract_urgency_risk(text):
    patterns = [
        r"(?:risco\s+de\s+morte|risk\s+of\s+death|riesgo\s+de\s+(?:muerte|fallecimiento))[^.]*\.",
        r"(?:vida\s+em\s+risco|vida\s+en\s+riesgo|life.?threatening)[^.]*\.",
        r"(?:urgente|urgency|urgencia)[^.]*\.",
        r"(?:grave|gravidade|gravedad|serious condition)[^.]*\.",
        r"(?:progressão|progresión|progression|disease\s+progression)[^.]*\.",
        r"(?:paliativ[ao]|palliative)[^.]*\.",
        r"(?:irreversível|irreversible)[^.]*\.",
    ]
    return find_pattern(text, patterns)

def extract_medical_justification(text):
    patterns = [
        r"(?:indicad[ao]|prescribed|prescrit[ao]|indicación)[^.]*\.",
        r"(?:tratamento\s+de\s+escolha|treatment\s+of\s+choice|tratamiento\s+de\s+elección)[^.]*\.",
        r"(?:única\s+alternativa|only\s+(?:option|alternative)|única\s+opción)[^.]*\.",
        r"(?:approve[d]?|aprovad[ao]|aprobad[ao])[^.]*FDA[^.]*\.",
        r"(?:FDA|EMA|ANMAT|ANVISA)[^.]*(?:approv|aprovad|aprobad)[^.]*\.",
        r"(?:médico\s+tratante|treating\s+physician|médico\s+tratante)[^.]*prescri[^.]*\.",
    ]
    return find_pattern(text, patterns)

def determine_decision(row, pdf_text):
    granted = val(row.get("Pembrolizumab Granted", ""))
    if granted == "1" or granted.lower() == "yes":
        return "DEFERIDO"
    elif granted == "0" or granted.lower() == "no":
        return "INDEFERIDO"
    elif "partially" in granted.lower() or "parcial" in granted.lower():
        return "DEFERIDO PARCIALMENTE"
    elif granted == "" and pdf_text:
        # Try to detect from PDF
        defer_kw = ["hace lugar", "se ordena", "granted", "deferido", "ordenar a",
                    "se otorga", "lugar a la acción", "amparo procedente"]
        indefer_kw = ["rechaza", "denied", "indeferido", "no hacer lugar", "improcedente",
                      "no lugar", "dismissed"]
        text_lower = pdf_text.lower()
        defer_hits = sum(1 for k in defer_kw if k in text_lower)
        indefer_hits = sum(1 for k in indefer_kw if k in text_lower)
        if defer_hits > indefer_hits:
            return "DEFERIDO (inferido do PDF)"
        elif indefer_hits > defer_hits:
            return "INDEFERIDO (inferido do PDF)"
    return "não localizado nos autos"

def assess_column(colname, value, pdf_text):
    if not value or value.lower() in ["ni", "none", "", "nan"]:
        return "incompleta", f"Sem {colname} registrado"
    if len(value) < 15:
        return "incompleta", f"{colname} muito sucinto"
    return "adequada", ""

def generate_corrections(row, pdf_text, n_status, o_status, n_obs, o_obs,
                          age, cid, staging, ecog, prior_tx, metastasis,
                          urgency, med_just):
    corrections = []
    disease_n = val(row.get("Disease", ""))
    clinical_o = val(row.get("Further clinical detalis", ""))

    if n_status != "adequada":
        if age != "não localizado nos autos":
            corrections.append(f"Coluna N: adicionar idade do paciente ({age})")
        if cid != "não localizado nos autos":
            corrections.append(f"Coluna N: adicionar CID ({cid})")
        if staging != "não localizado nos autos":
            corrections.append(f"Coluna N: adicionar estadiamento ({staging[:80]})")

    if o_status != "adequada":
        details = []
        if prior_tx != "não localizado nos autos":
            details.append(f"Tratamentos prévios: {prior_tx[:150]}")
        if metastasis != "não localizado nos autos":
            details.append(f"Metástases: {metastasis[:150]}")
        if urgency != "não localizado nos autos":
            details.append(f"Gravidade/urgência: {urgency[:150]}")
        if med_just != "não localizado nos autos":
            details.append(f"Justificativa médica: {med_just[:150]}")
        if details:
            corrections.append("Coluna O: complementar com — " + "; ".join(details))
        else:
            corrections.append("Coluna O: sem detalhes clínicos nos autos — verificar manualmente")

    if not corrections:
        if pdf_text and not pdf_text.startswith("[ERRO"):
            corrections.append("Dados da planilha parecem completos. Verificar coerência com PDF se necessário.")
        else:
            corrections.append("Sem PDF disponível para verificação adicional.")

    return corrections

def analyze(idx, row, pdf_text):
    country = val(row.get("Country"))
    case_num = val(row.get("Case Number"))
    court = val(row.get("Court"))
    date_str = format_date(row.get("Date"))
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
    notes = val(row.get("Notes"))
    esmo = val(row.get("ESMO Score"))
    who_eml = val(row.get("WHO EML "))
    pdf_link = val(row.get("PDF Link"))

    # Extract from PDF
    age = cid = staging = ecog = prior_tx = metastasis = urgency = med_just = "não localizado nos autos"
    if pdf_text and not pdf_text.startswith("[ERRO"):
        age = extract_age(pdf_text)
        cid = extract_cid(pdf_text)
        staging = extract_staging(pdf_text)
        ecog = extract_ecog(pdf_text)
        prior_tx = extract_prior_treatments(pdf_text)
        metastasis = extract_metastasis(pdf_text)
        urgency = extract_urgency_risk(pdf_text)
        med_just = extract_medical_justification(pdf_text)

    resultado = determine_decision(row, pdf_text)

    n_status, n_obs = assess_column("diagnóstico", disease_n, pdf_text)
    o_status, o_obs = assess_column("detalhes clínicos", clinical_o, pdf_text)

    corrections = generate_corrections(row, pdf_text, n_status, o_status, n_obs, o_obs,
                                        age, cid, staging, ecog, prior_tx, metastasis,
                                        urgency, med_just)

    SEP = "=" * 70
    sep2 = "-" * 70
    NL = "não localizado nos autos"

    lines = [
        f"\n{SEP}",
        f"PROCESSO #{idx}: {country} — {case_num}",
        SEP,
        f"Tribunal: {court}",
        f"Data: {date_str}",
        f"Réu/Demandado: {defendant if defendant else NL}",
        f"PDF: {pdf_link if pdf_link else 'não disponível'}",
        f"",
        f"RESUMO CLÍNICO:",
        f"  Diagnóstico:              {disease_n if disease_n else NL}",
        f"  CID:                      {cid}",
        f"  Idade:                    {age}",
        f"  Estadiamento:             {staging[:120] if staging != NL else NL}",
        f"  ECOG:                     {ecog}",
        f"  Histórico/Tratamentos:    {prior_tx[:200] if prior_tx != NL else NL}",
        f"  Metástases:               {metastasis[:200] if metastasis != NL else NL}",
        f"  Gravidade/Urgência:       {urgency[:200] if urgency != NL else NL}",
        f"  Justificativa médica:     {med_just[:200] if med_just != NL else NL}",
        f"  Outros medicamentos:      {other_meds if other_meds else NL}",
        f"  ESMO Score:               {esmo if esmo else NL}",
        f"  WHO EML:                  {who_eml if who_eml else NL}",
        f"  Detalhes clínicos (col.O):{clinical_o if clinical_o else NL}",
        f"",
        f"ANÁLISE DA DECISÃO:",
        f"  Resultado:                {resultado}",
        f"  Base jurídica (demandante): {legal_claim if legal_claim else NL}",
        f"  Base jurídica (negativa):   {legal_denial if legal_denial else NL}",
        f"  Fundamentação decisão:      {legal_decision[:300] if legal_decision else NL}",
        f"  Evidência científica:       {evidence if evidence else NL}",
        f"  Observações:                {notes[:300] if notes else NL}",
        f"  Incorporado ao protocolo:   {incorporated if incorporated else NL}",
        f"  Dentro do protocolo:        {within_protocol if within_protocol else NL}",
        f"  Registro para a doença:     {registration if registration else NL}",
        f"",
        f"CONFERÊNCIA PLANILHA:",
        f"  Coluna N: {n_status}{' — ' + n_obs if n_obs else ''}",
        f"  Coluna O: {o_status}{' — ' + o_obs if o_obs else ''}",
        f"  Status geral: {'OK' if n_status == 'adequada' and o_status == 'adequada' else 'REVISAR'}",
        f"",
        f"CORREÇÕES SUGERIDAS:",
    ]
    for c in corrections:
        lines.append(f"  - {c}")

    return "\n".join(lines)

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
        rows.append({headers[j]: row[j] for j in range(min(len(headers), len(row)))})
    return rows

def main():
    start = int(sys.argv[1]) if len(sys.argv) > 1 else 1
    end_arg = int(sys.argv[2]) if len(sys.argv) > 2 else None
    save_file = sys.argv[3] == "save" if len(sys.argv) > 3 else False

    rows = load_rows()
    total = len(rows)
    end = end_arg if end_arg else total

    out_lines = []
    header = (
        f"RELATÓRIO DE ANÁLISE DE PROCESSOS — OPAS Pembrolizumab\n"
        f"Gerado em: {date.today().strftime('%d/%m/%Y')}\n"
        f"Total de registros na planilha: {total}\n"
        f"Intervalo analisado: #{start} a #{end}\n"
        f"{'=' * 70}"
    )
    out_lines.append(header)
    print(header)

    analyzed = skipped = 0
    for i, row in enumerate(rows):
        idx = i + 1
        if idx < start or idx > end:
            continue

        country = val(row.get("Country"))
        case_num = val(row.get("Case Number"))
        pdf_link = val(row.get("PDF Link"))

        skippable, disease = is_skip(row)
        if skippable:
            skip_line = f"\nPROCESSO #{idx}: {country} — {case_num}\n  [FORA DO ESCOPO: {disease or 'questão processual/não relacionado a pembrolizumab'}]"
            out_lines.append(skip_line)
            print(skip_line)
            skipped += 1
            continue

        pdf_path = find_pdf(pdf_link, country)
        pdf_text = ""
        pdf_status = "PDF não encontrado"
        if pdf_path:
            print(f"  [Lendo PDF {idx}: {os.path.basename(pdf_path)}]", file=sys.stderr)
            pdf_text = extract_pdf_text(pdf_path)
            pdf_status = f"PDF lido: {os.path.basename(pdf_path)}"
        elif pdf_link:
            pdf_status = f"PDF não localizado localmente: {pdf_link}"

        report = analyze(idx, row, pdf_text)
        out_lines.append(report)
        print(report)
        analyzed += 1

    summary = (
        f"\n{'=' * 70}\n"
        f"RESUMO FINAL\n"
        f"  Total de registros: {total}\n"
        f"  Analisados: {analyzed}\n"
        f"  Fora do escopo (ignorados): {skipped}\n"
        f"  Intervalo: #{start}-#{end}\n"
        f"{'=' * 70}"
    )
    out_lines.append(summary)
    print(summary)

    if save_file:
        with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
            f.write("\n".join(out_lines))
        print(f"\nRelatório salvo em: {OUTPUT_FILE}", file=sys.stderr)

if __name__ == "__main__":
    main()
