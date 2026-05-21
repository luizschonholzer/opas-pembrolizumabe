# OPAS — Análise de Processos Judiciais sobre Pembrolizumabe

Repositório de scripts Python desenvolvidos para análise automatizada de processos judiciais envolvendo o fornecimento de **Pembrolizumabe** (imunoterapia oncológica) em países da América Latina.

## Contexto

Este projeto foi desenvolvido como parte de pesquisa orientada pela OPAS (Organização Pan-Americana da Saúde). O corpus consiste em decisões judiciais de 8 países — Argentina, Brasil, Chile, Colômbia, Equador, Guatemala, Costa Rica e Uruguai — onde pacientes com câncer solicitaram judicialmente o fornecimento do medicamento pelo sistema público de saúde.

## Metodologia de extração de dados

### Como os PDFs foram lidos

Os scripts utilizam a biblioteca [`pdfplumber`](https://github.com/jsvine/pdfplumber) para converter o texto dos PDFs em formato legível por máquina. A extração clínica é feita em duas etapas:

#### 1. Vinculação PDF ↔ Planilha (`add_pdf_links.py`)

Cada processo na planilha é associado ao seu PDF correspondente por meio de um **algoritmo de pontuação** (`score_match`) que compara:
- Número do processo vs. nome do arquivo PDF
- Data da decisão vs. data no nome do arquivo (útil especialmente para o Uruguai)
- Correspondência numérica e textual normalizada

#### 2. Extração de dados clínicos (via expressões regulares)

Os scripts de análise (`analisar_v4.py`, `processar_pdfs.py`, `analyze_pdfs_full.py`) utilizam **expressões regulares (regex)** para identificar campos clínicos no texto extraído dos PDFs. Cada campo possui padrões específicos:

| Campo clínico | Exemplo de padrão usado |
|---|---|
| Idade do paciente | `(\d{1,3})\s*anos?\s+de\s+edad` |
| Estadiamento | `[Ee]stadio\s+(IV\|III\|II\|I)` |
| PD-L1 | `PD[L-]?1\s*[:\-]?\s*([\d<>]+\s*%?)` |
| MSI / dMMR | `MSI[- ]?H\|dMMR\|deficient mismatch` |
| ECOG | `ECOG\s*[:\-]?\s*([0-4])` |
| Metástases | `metástasis\s+.*\s+(pulmón\|hígado\|hueso...)` |
| Tratamentos anteriores | Lista de ~40 drogas e procedimentos |
| Urgência clínica | `riesgo de muerte\|risk of death\|urgente` |

> **Importante:** Esta abordagem identifica padrões de texto, mas **não interpreta semanticamente** o conteúdo. Por exemplo, o script encontra "PD-L1: 50%" no texto, mas não avalia se isso é clinicamente relevante para aquele caso específico.

#### 3. Análise de lacunas (`find_gaps` em `analisar_v4.py`)

Após a extração, o script compara os dados encontrados no PDF com o conteúdo já presente nas colunas N (diagnóstico) e O (detalhes clínicos) da planilha, identificando informações ausentes que deveriam ser registradas.

#### 4. Atualização da planilha (`aplicar_resultados.py`, `update_n_o_columns.py`)

Os dados extraídos são aplicados de volta à planilha Excel, preenchendo campos vazios e sinalizando inconsistências com formatação visual (células destacadas em amarelo/vermelho).

## Estrutura dos scripts

```
├── add_pdf_links.py          # Vincula PDFs às linhas da planilha (scoring)
├── extrair_dados.py          # Extração inicial de texto e dados
├── analisar_processos.py     # Análise v1: relatório por processo
├── analisar_processos_v2.py  # Análise v2: com leitura de PDF
├── analisar_v3.py            # Análise v3: extração clínica estruturada
├── analisar_v4.py            # Análise v4: comparação N/O vs PDF + coluna nova
├── analyze_pdfs_full.py      # Análise completa com relatório por seções
├── analyze_pdfs_v2.py        # Versão alternativa da análise completa
├── processar_pdfs.py         # Processamento e complemento da coluna AB
├── aplicar_resultados.py     # Aplica resultados JSON à planilha
└── update_n_o_columns.py     # Correções manuais nas colunas N e O
```

## Países cobertos

| País | Pasta | Período |
|---|---|---|
| Argentina | `Argentina/` | 2020–2025 |
| Brasil | `Brasil/` | — |
| Chile | `Chile/` | 2016 |
| Colômbia | `Colombia/` | — |
| Equador | `Equador/` | 2021 |
| Guatemala | `Guatemala/` | 2023–2025 |
| Costa Rica | `Costa Rica/` | 2019–2024 |
| Uruguai | `Uruguai/` | 2019–2026 |

## Dependências

```bash
pip install pdfplumber openpyxl python-docx
```

## Limitações metodológicas

- A extração por regex é sensível à qualidade do PDF (PDFs escaneados sem OCR retornam texto vazio)
- Padrões em espanhol, português e inglês foram incluídos, mas variações ortográficas regionais podem não ser capturadas
- O algoritmo de vinculação PDF ↔ processo pode gerar falsos positivos em casos com numeração ambígua
- A análise não substitui revisão humana especializada
