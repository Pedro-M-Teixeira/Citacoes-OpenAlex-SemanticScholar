#!/usr/bin/env python3
# =============================================================================
#  ANÁLISE BIBLIOMÉTRICA — RSL Literacia em Dados / Administração Pública
#  OpenAlex + Semantic Scholar
#  Versão 2.0 — Junho 2026
#
#  USO RÁPIDO:
#    pip install -r requirements.txt
#    python bibliometria_rsl.py --ris 39_Artigos_Finais_RSL.ris
#
#  OUTPUT: bibliometria_resultados.xlsx  (4 folhas)
# =============================================================================

import argparse, hashlib, json, os, sys, time
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests
import rispy
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from tqdm import tqdm

# ── Configuração ──────────────────────────────────────────────────────────────
EMAIL             = "COLOQUEAQUIOSEUEMAIL@EMAIL"   # polite pool OpenAlex
PAUSA             = 0.4     # segundos entre requests
MAX_RETRY         = 3
OUTPUT_DEFAULT    = "bibliometria_resultados.xlsx"
OA_BASE           = "https://api.openalex.org"
SS_BASE           = "https://api.semanticscholar.org/graph/v1"

# Paleta visual Excel
C_AZUL_ESC  = "1A3A5C"
C_AZUL_MED  = "2E75B6"
C_AZUL_CLR  = "D6E8F7"
C_VERDE_ESC = "1A4D2E"
C_VERDE_CLR = "D6F0E0"
C_AMARELO   = "FFF2CC"
C_BRANCO    = "FFFFFF"
C_CINZA     = "F5F5F5"

# ── Parser RIS ────────────────────────────────────────────────────────────────

def _normalizar_doi(raw: str) -> str:
    doi = raw.strip().lower()
    for pref in ("https://doi.org/", "http://doi.org/", "doi:", "https://dx.doi.org/"):
        if doi.startswith(pref):
            doi = doi[len(pref):]
    return doi.strip()

def parse_ris(caminho: str) -> list[dict]:
    print(f"\n📂  A ler: {caminho}")
    with open(caminho, "r", encoding="utf-8") as f:
        entries = rispy.load(f)

    artigos, sem_doi = [], 0
    for idx, e in enumerate(entries, 1):
        doi_raw = (e.get("doi") or e.get("DO") or
                   e.get("electronic_resource_num") or "")
        doi = _normalizar_doi(doi_raw)

        autores_raw = e.get("authors") or e.get("first_authors") or []
        autores = "; ".join(autores_raw) if isinstance(autores_raw, list) else str(autores_raw)

        ano = str(e.get("year") or e.get("publication_year") or "")
        if isinstance(ano, list): ano = ano[0] if ano else ""

        artigos.append({
            "n":            idx,
            "ris_doi":      doi,
            "ris_titulo":   (e.get("title") or e.get("primary_title") or "").strip(),
            "ris_autores":  autores,
            "ris_ano":      ano,
            "ris_revista":  (e.get("journal_name") or e.get("secondary_title") or
                             e.get("JO") or "").strip(),
            "ris_volume":   str(e.get("volume") or ""),
            "ris_numero":   str(e.get("number") or ""),
            "ris_paginas":  f"{e.get('start_page','')}-{e.get('end_page','')}".strip("-"),
            "ris_keywords": "; ".join(e.get("keywords") or [])[:300],
            "ris_abstract": (e.get("abstract") or "")[:500],
            "ris_lang":     e.get("language") or "",
        })
        if not doi: sem_doi += 1

    print(f"   ✅  {len(artigos)} artigos  |  {len(artigos)-sem_doi} com DOI  |  {sem_doi} sem DOI")
    return artigos

def parse_dois(caminho: str) -> list[dict]:
    with open(caminho, "r", encoding="utf-8") as f:
        linhas = [l.strip() for l in f if l.strip()]
    return [{"n": i+1, "ris_doi": _normalizar_doi(d), "ris_titulo": "",
             "ris_autores": "", "ris_ano": "", "ris_revista": "",
             "ris_volume": "", "ris_numero": "", "ris_paginas": "",
             "ris_keywords": "", "ris_abstract": "", "ris_lang": ""}
            for i, d in enumerate(linhas)]

# ── HTTP robusto ──────────────────────────────────────────────────────────────

def _get(url, params=None, headers=None):
    for t in range(MAX_RETRY):
        try:
            r = requests.get(url, params=params, headers=headers, timeout=15)
            if r.status_code == 200:   return r.json()
            if r.status_code == 404:   return None
            if r.status_code == 429:
                espera = 15 * (t+1)
                print(f"\n   ⏳  Rate-limit. Aguardar {espera}s...")
                time.sleep(espera)
            else:
                time.sleep(3 * (t+1))
        except Exception as ex:
            print(f"\n   ⚠️  Rede ({t+1}/{MAX_RETRY}): {ex}")
            time.sleep(4 * (t+1))
    return None

# ── OpenAlex ──────────────────────────────────────────────────────────────────

def oa_artigo(doi: str, titulo: str = "") -> dict | None:
    d = None
    if doi:
        d = _get(f"{OA_BASE}/works/https://doi.org/{doi}",
                 params={"mailto": EMAIL})
        time.sleep(PAUSA)
    if d is None and titulo and len(titulo) > 12:
        res = _get(f"{OA_BASE}/works",
                   params={"filter": f"title.search:{titulo[:120]}",
                           "per-page": 1, "mailto": EMAIL})
        time.sleep(PAUSA)
        if res and res.get("results"):
            d = res["results"][0]
    if not d or "id" not in d: return None

    oa_info  = d.get("open_access") or {}
    fonte    = ((d.get("primary_location") or {}).get("source") or {})
    autores  = []
    for a in (d.get("authorships") or []):
        ai  = a.get("author") or {}
        ins = [i.get("display_name","") for i in (a.get("institutions") or [])]
        autores.append({"nome": ai.get("display_name",""),
                        "oa_id": ai.get("id",""),
                        "afil": "; ".join(ins)})
    conc = [c.get("display_name","") for c in (d.get("concepts") or [])[:6]]
    return {
        "oa_id":       d.get("id",""),
        "oa_doi":      _normalizar_doi(d.get("doi","") or ""),
        "oa_titulo":   d.get("display_name") or d.get("title",""),
        "oa_ano":      d.get("publication_year",""),
        "oa_cits":     d.get("cited_by_count") or 0,
        "oa_revista":  fonte.get("display_name",""),
        "oa_issn":     ", ".join(fonte.get("issn") or []),
        "oa_is_oa":    "Sim" if oa_info.get("is_oa") else "Não",
        "oa_url_oa":   oa_info.get("oa_url",""),
        "oa_tipo":     d.get("type",""),
        "oa_conceitos": "; ".join(conc),
        "oa_autores":  autores,
    }

def oa_autor(oa_id: str) -> dict:
    if not oa_id: return {}
    d = _get(oa_id, params={"mailto": EMAIL})
    time.sleep(PAUSA)
    if not d: return {}
    s  = d.get("summary_stats") or {}
    li = d.get("last_known_institution") or {}
    return {
        "oa_h":    s.get("h_index") or d.get("h_index") or 0,
        "oa_cits": d.get("cited_by_count") or 0,
        "oa_wks":  d.get("works_count") or 0,
        "oa_i10":  s.get("i10_index") or 0,
        "oa_afil": li.get("display_name",""),
        "orcid":   (d.get("ids") or {}).get("orcid",""),
    }

# ── Semantic Scholar ──────────────────────────────────────────────────────────

_SS_FIELDS = ("title,year,citationCount,influentialCitationCount,"
              "isOpenAccess,openAccessPdf,venue,journal,authors,"
              "fieldsOfStudy,referenceCount,externalIds")
_SS_A_FLD  = "name,hIndex,citationCount,paperCount,affiliations"

def ss_artigo(doi: str, titulo: str = "") -> dict | None:
    d = None
    if doi:
        d = _get(f"{SS_BASE}/paper/DOI:{doi}", params={"fields": _SS_FIELDS})
        time.sleep(PAUSA)
    if d is None and titulo and len(titulo) > 12:
        res = _get(f"{SS_BASE}/paper/search",
                   params={"query": titulo[:200], "fields": _SS_FIELDS, "limit": 1})
        time.sleep(PAUSA)
        if res and res.get("data"): d = res["data"][0]
    if not d or "paperId" not in d: return None

    rev = ""
    if d.get("journal"):  rev = d["journal"].get("name","")
    elif d.get("venue"):  rev = d["venue"]

    autores = [{"nome": a.get("name",""), "ss_id": a.get("authorId","")}
               for a in (d.get("authors") or [])]
    return {
        "ss_id":       d.get("paperId",""),
        "ss_cits":     d.get("citationCount") or 0,
        "ss_infl":     d.get("influentialCitationCount") or 0,
        "ss_refs":     d.get("referenceCount") or 0,
        "ss_is_oa":    "Sim" if d.get("isOpenAccess") else "Não",
        "ss_pdf":      (d.get("openAccessPdf") or {}).get("url",""),
        "ss_revista":  rev,
        "ss_areas":    "; ".join(d.get("fieldsOfStudy") or []),
        "ss_autores":  autores,
    }

def ss_autor(ss_id: str) -> dict:
    if not ss_id: return {}
    d = _get(f"{SS_BASE}/author/{ss_id}", params={"fields": _SS_A_FLD})
    time.sleep(PAUSA)
    if not d: return {}
    afils_raw = d.get("affiliations") or []
    afil = ", ".join(
        a.get("name", "") if isinstance(a, dict) else str(a)
        for a in afils_raw
    )
    return {
        "ss_h":    d.get("hIndex") or 0,
        "ss_cits": d.get("citationCount") or 0,
        "ss_wks":  d.get("paperCount") or 0,
        "ss_afil": afil,
    }

# ── Pipeline ──────────────────────────────────────────────────────────────────

def processar(corpus: list[dict]):
    artigos_out, autores_db = [], {}

    print(f"\n{'─'*62}")
    print(f"  Corpus: {len(corpus)} artigos  |  APIs: OpenAlex + Semantic Scholar")
    print(f"{'─'*62}\n")

    for art in tqdm(corpus, desc="Artigos", unit="art"):
        doi, titulo = art["ris_doi"], art["ris_titulo"]
        row = {**art}

        # ── OpenAlex ────────────────────────────────────────────────────
        oa = oa_artigo(doi, titulo)
        if oa:
            row.update({
                "oa_ok":        "✓",
                "oa_cits":      oa["oa_cits"],
                "oa_revista":   oa["oa_revista"],
                "oa_issn":      oa["oa_issn"],
                "oa_is_oa":     oa["oa_is_oa"],
                "oa_url_oa":    oa["oa_url_oa"],
                "oa_tipo":      oa["oa_tipo"],
                "oa_conceitos": oa["oa_conceitos"],
                "oa_id":        oa["oa_id"],
            })
            nomes_oa = []
            for a in oa["oa_autores"]:
                n_norm = a["nome"].lower().strip()
                nomes_oa.append(a["nome"])
                if n_norm and n_norm not in autores_db:
                    p = oa_autor(a["oa_id"])
                    autores_db[n_norm] = {
                        "nome": a["nome"],
                        "afil_oa": a["afil"] or p.get("oa_afil",""),
                        "oa_h": p.get("oa_h",""), "oa_cits": p.get("oa_cits",""),
                        "oa_wks": p.get("oa_wks",""), "oa_i10": p.get("oa_i10",""),
                        "orcid": p.get("orcid",""),
                        "ss_h": "", "ss_cits": "", "ss_wks": "", "ss_afil": "",
                        "artigos": [],
                    }
                if n_norm: autores_db[n_norm]["artigos"].append(titulo or doi)
            row["oa_autores_str"] = "; ".join(nomes_oa)
        else:
            for k in ("oa_ok","oa_cits","oa_revista","oa_issn","oa_is_oa",
                      "oa_url_oa","oa_tipo","oa_conceitos","oa_id","oa_autores_str"):
                row[k] = "" if k != "oa_ok" else "✗"

        # ── Semantic Scholar ─────────────────────────────────────────────
        ss = ss_artigo(doi, titulo)
        if ss:
            row.update({
                "ss_ok":    "✓",
                "ss_cits":  ss["ss_cits"],
                "ss_infl":  ss["ss_infl"],
                "ss_refs":  ss["ss_refs"],
                "ss_is_oa": ss["ss_is_oa"],
                "ss_pdf":   ss["ss_pdf"],
                "ss_areas": ss["ss_areas"],
                "ss_id":    ss["ss_id"],
            })
            nomes_ss = []
            for a in ss["ss_autores"]:
                n_norm = a["nome"].lower().strip()
                nomes_ss.append(a["nome"])
                if n_norm and n_norm in autores_db and not autores_db[n_norm]["ss_h"]:
                    p2 = ss_autor(a["ss_id"])
                    autores_db[n_norm].update({
                        "ss_h": p2.get("ss_h",""), "ss_cits": p2.get("ss_cits",""),
                        "ss_wks": p2.get("ss_wks",""), "ss_afil": p2.get("ss_afil",""),
                    })
                elif n_norm and n_norm not in autores_db and a["nome"]:
                    p2 = ss_autor(a["ss_id"])
                    autores_db[n_norm] = {
                        "nome": a["nome"], "afil_oa": "",
                        "oa_h": "", "oa_cits": "", "oa_wks": "", "oa_i10": "", "orcid": "",
                        "ss_h": p2.get("ss_h",""), "ss_cits": p2.get("ss_cits",""),
                        "ss_wks": p2.get("ss_wks",""), "ss_afil": p2.get("ss_afil",""),
                        "artigos": [titulo or doi],
                    }
            row["ss_autores_str"] = "; ".join(nomes_ss)
        else:
            for k in ("ss_ok","ss_cits","ss_infl","ss_refs","ss_is_oa",
                      "ss_pdf","ss_areas","ss_id","ss_autores_str"):
                row[k] = "" if k != "ss_ok" else "✗"

        # Média citações
        vals = [v for v in [row.get("oa_cits"), row.get("ss_cits")]
                if isinstance(v, (int, float)) and v != ""]
        row["media_cits"] = round(sum(vals)/len(vals), 1) if vals else ""

        artigos_out.append(row)

    lista_aut = sorted([
        {
            "Nome":                  v["nome"],
            "Afiliação (OA)":        v["afil_oa"] or v["ss_afil"],
            "OA: h-index":           v["oa_h"],
            "OA: Citações Totais":   v["oa_cits"],
            "OA: N.º Publicações":   v["oa_wks"],
            "OA: i10-index":         v["oa_i10"],
            "ORCID":                 v["orcid"],
            "SS: h-index":           v["ss_h"],
            "SS: Citações Totais":   v["ss_cits"],
            "SS: N.º Publicações":   v["ss_wks"],
            "N.º Artigos no Corpus": len(set(v["artigos"])),
        }
        for v in autores_db.values()
    ], key=lambda x: (x["OA: h-index"] or 0), reverse=True)

    return artigos_out, lista_aut

# ── Excel ─────────────────────────────────────────────────────────────────────

COLS_ARTIGOS = [
    ("n",            "#"),
    ("ris_titulo",   "Título"),
    ("ris_autores",  "Autores (RIS)"),
    ("ris_ano",      "Ano"),
    ("ris_revista",  "Revista (RIS)"),
    ("ris_doi",      "DOI"),
    ("oa_ok",        "OA ✓/✗"),
    ("oa_cits",      "OA: Citações"),
    ("oa_revista",   "OA: Revista"),
    ("oa_issn",      "OA: ISSN"),
    ("oa_is_oa",     "OA: Acesso Aberto"),
    ("oa_conceitos", "OA: Conceitos Temáticos"),
    ("oa_tipo",      "OA: Tipo"),
    ("oa_url_oa",    "OA: URL Acesso Aberto"),
    ("oa_autores_str","OA: Autores"),
    ("ss_ok",        "SS ✓/✗"),
    ("ss_cits",      "SS: Citações"),
    ("ss_infl",      "SS: Citações Influentes"),
    ("ss_refs",      "SS: N.º Referências"),
    ("ss_is_oa",     "SS: Acesso Aberto"),
    ("ss_areas",     "SS: Áreas Científicas"),
    ("ss_pdf",       "SS: URL PDF"),
    ("ss_autores_str","SS: Autores"),
    ("media_cits",   "Média Citações (OA+SS)"),
    ("ris_keywords", "Palavras-Chave"),
    ("ris_abstract", "Resumo (300 car.)"),
]


def _cab(ws, row_n, fill_hex):
    fill = PatternFill("solid", fgColor=fill_hex)
    ft   = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
    for cell in ws[row_n]:
        cell.fill = fill; cell.font = ft
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row_n].height = 36

def _zebra(ws, even_hex, odd_hex="FFFFFF"):
    f_par  = PatternFill("solid", fgColor=even_hex)
    f_imp  = PatternFill("solid", fgColor=odd_hex)
    ft     = Font(name="Calibri", size=10)
    for i, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=2):
        for cell in row:
            cell.fill = f_par if i % 2 == 0 else f_imp
            cell.font = ft
            cell.alignment = Alignment(vertical="center", wrap_text=True)

def _largura(ws, mn=10, mx=55):
    for col in ws.columns:
        ml = max((len(str(c.value or "")) for c in col), default=mn)
        ws.column_dimensions[col[0].column_letter].width = min(max(ml+2, mn), mx)

def exportar(artigos_out, lista_aut, caminho):
    print(f"\n📊  A gerar Excel: {caminho}")
    n = len(artigos_out)

    # ── Artigos ─────────────────────────────────────────────────────────
    rows_art = []
    for a in artigos_out:
        rows_art.append({label: a.get(key,"") for key, label in COLS_ARTIGOS})
    df_art = pd.DataFrame(rows_art)

    # ── Autores ─────────────────────────────────────────────────────────
    df_aut = pd.DataFrame(lista_aut) if lista_aut else pd.DataFrame(
        columns=["Nome","Afiliação (OA)","OA: h-index","OA: Citações Totais",
                 "OA: N.º Publicações","OA: i10-index","ORCID",
                 "SS: h-index","SS: Citações Totais","SS: N.º Publicações",
                 "N.º Artigos no Corpus"])

    # ── Cobertura ────────────────────────────────────────────────────────
    n_oa    = sum(1 for a in artigos_out if a.get("oa_ok")=="✓")
    n_ss    = sum(1 for a in artigos_out if a.get("ss_ok")=="✓")
    n_amb   = sum(1 for a in artigos_out if a.get("oa_ok")=="✓" and a.get("ss_ok")=="✓")
    n_none  = sum(1 for a in artigos_out if a.get("oa_ok")!="✓" and a.get("ss_ok")!="✓")
    cits_oa = [a["oa_cits"] for a in artigos_out if isinstance(a.get("oa_cits"),(int,float)) and a.get("oa_cits")!=""]
    cits_ss = [a["ss_cits"] for a in artigos_out if isinstance(a.get("ss_cits"),(int,float)) and a.get("ss_cits")!=""]

    df_cob = pd.DataFrame([
        {"Indicador": "Total de artigos no corpus",                    "N": n,       "%": "100%",    "Notas": ""},
        {"Indicador": "Artigos encontrados no OpenAlex",               "N": n_oa,    "%": f"{n_oa/n*100:.1f}%", "Notas": ""},
        {"Indicador": "Artigos encontrados no Semantic Scholar",       "N": n_ss,    "%": f"{n_ss/n*100:.1f}%", "Notas": ""},
        {"Indicador": "Artigos encontrados em ambas as APIs",          "N": n_amb,   "%": f"{n_amb/n*100:.1f}%","Notas": ""},
        {"Indicador": "Artigos não encontrados em nenhuma API",        "N": n_none,  "%": f"{n_none/n*100:.1f}%","Notas": "Pesquisar manualmente no Google Scholar"},
        {"Indicador": "Total de autores únicos identificados",         "N": len(lista_aut), "%": "—","Notas": ""},
        {"Indicador": "─── Citações (OpenAlex) ───",                  "N": "",      "%": "",        "Notas": ""},
        {"Indicador": "Total de citações (soma corpus, OA)",           "N": sum(cits_oa) if cits_oa else "N/D", "%": "", "Notas": ""},
        {"Indicador": "Média de citações por artigo (OA)",             "N": round(sum(cits_oa)/len(cits_oa),1) if cits_oa else "N/D", "%": "", "Notas": ""},
        {"Indicador": "Máximo de citações (artigo mais citado, OA)",   "N": max(cits_oa) if cits_oa else "N/D", "%": "", "Notas": ""},
        {"Indicador": "Mediana de citações (OA)",                      "N": sorted(cits_oa)[len(cits_oa)//2] if cits_oa else "N/D", "%": "", "Notas": ""},
        {"Indicador": "─── Citações (Semantic Scholar) ───",          "N": "",      "%": "",        "Notas": ""},
        {"Indicador": "Total de citações (soma corpus, SS)",           "N": sum(cits_ss) if cits_ss else "N/D", "%": "", "Notas": ""},
        {"Indicador": "Média de citações por artigo (SS)",             "N": round(sum(cits_ss)/len(cits_ss),1) if cits_ss else "N/D", "%": "", "Notas": ""},
        {"Indicador": "Artigos com citações influentes > 0 (SS)",      "N": sum(1 for a in artigos_out if isinstance(a.get("ss_infl"),(int,float)) and (a.get("ss_infl") or 0) > 0), "%": "", "Notas": ""},
    ])

    # Artigos não encontrados
    nao_enc = [a for a in artigos_out if a.get("oa_ok")!="✓" and a.get("ss_ok")!="✓"]
    if nao_enc:
        extras = [{"Indicador": "", "N": "", "%": "", "Notas": ""}]
        extras.append({"Indicador": "ARTIGOS NÃO ENCONTRADOS NAS APIs — pesquisar manualmente:", "N": "", "%": "", "Notas": ""})
        for x in nao_enc:
            extras.append({"Indicador": f"  [{x['ris_ano']}] {x['ris_titulo'][:80]}", "N": x["ris_doi"], "%": "", "Notas": ""})
        df_cob = pd.concat([df_cob, pd.DataFrame(extras)], ignore_index=True)

    # ── Metadados ────────────────────────────────────────────────────────
    sha = hashlib.sha256(Path(__file__).read_bytes()).hexdigest()
    df_meta = pd.DataFrame([
        {"Parâmetro": "Script",             "Valor": Path(__file__).name},
        {"Parâmetro": "Versão",             "Valor": "2.0"},
        {"Parâmetro": "Data de execução",   "Valor": datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
        {"Parâmetro": "SHA-256 do script",  "Valor": sha},
        {"Parâmetro": "Python",             "Valor": sys.version.split()[0]},
        {"Parâmetro": "API OpenAlex",       "Valor": OA_BASE},
        {"Parâmetro": "API Sem. Scholar",   "Valor": SS_BASE},
        {"Parâmetro": "Email polite pool",  "Valor": EMAIL},
        {"Parâmetro": "Total artigos",      "Valor": str(n)},
        {"Parâmetro": "Nota metodológica",  "Valor": "Indicadores extraídos de APIs abertas (OpenAlex e Semantic Scholar). Recomenda-se triangulação com Scopus/WoS para fins formais. Depositar este script no Zenodo para citação na tese."},
        {"Parâmetro": "Como citar (APA 7)", "Valor": f"[Pedro Apelido] ({datetime.now().year}). Script Python para análise bibliométrica via OpenAlex e Semantic Scholar (v2.0) [Software]. Zenodo. https://doi.org/10.5281/zenodo.XXXXXXX"},
        {"Parâmetro": "Incluir na tese",    "Valor": f"SHA-256: {sha}"},
    ])

    # ── Escrever e formatar ──────────────────────────────────────────────
    with pd.ExcelWriter(caminho, engine="openpyxl") as w:
        df_art.to_excel(w,  sheet_name="1_Artigos",   index=False)
        df_aut.to_excel(w,  sheet_name="2_Autores",   index=False)
        df_cob.to_excel(w,  sheet_name="3_Cobertura", index=False)
        df_meta.to_excel(w, sheet_name="4_Metadados", index=False)

    wb = load_workbook(caminho)
    configs = [("1_Artigos", C_AZUL_ESC, C_AZUL_CLR),
               ("2_Autores", C_VERDE_ESC, C_VERDE_CLR),
               ("3_Cobertura", "444444", C_CINZA),
               ("4_Metadados", "444444", C_CINZA)]

    for nome, cor_cab, cor_par in configs:
        ws = wb[nome]
        ws.freeze_panes = "B2"
        _cab(ws, 1, cor_cab)
        _zebra(ws, cor_par)
        _largura(ws)
        ws.auto_filter.ref = ws.dimensions

    wb.save(caminho)

    # ── Relatório terminal ───────────────────────────────────────────────
    print(f"\n{'═'*62}")
    print(f"  ✅  CONCLUÍDO")
    print(f"{'─'*62}")
    print(f"  Artigos processados      : {n}")
    print(f"  OpenAlex encontrados     : {n_oa} ({n_oa/n*100:.1f}%)")
    print(f"  Semantic Scholar encon.  : {n_ss} ({n_ss/n*100:.1f}%)")
    print(f"  Não encontrados          : {n_none}")
    print(f"  Autores únicos           : {len(lista_aut)}")
    if cits_oa:
        print(f"  Total citações (OA)      : {sum(cits_oa)}")
        print(f"  Média citações (OA)      : {sum(cits_oa)/len(cits_oa):.1f}")
        print(f"  Artigo mais citado (OA)  : {max(cits_oa)}")
    print(f"{'─'*62}")
    print(f"  Ficheiro : {caminho}")
    print(f"  SHA-256  : {sha[:40]}...")
    print(f"{'═'*62}\n")
    print("  📌  Próximos passos:")
    print("  1. Deposita o script no Zenodo → obtém DOI citable")
    print("  2. Copia o SHA-256 (folha 4_Metadados) para o Apêndice da tese")
    print("  3. Corre: pip freeze > requirements_bibliometria.txt")
    print("  4. Inclui requirements_bibliometria.txt no depósito Zenodo\n")

# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    global EMAIL
    ap = argparse.ArgumentParser(
        description="Análise bibliométrica RSL — OpenAlex + Semantic Scholar",
        epilog="Exemplo: python bibliometria_rsl.py --ris 39_Artigos_Finais_RSL.ris")
    ap.add_argument("--ris",    help="Ficheiro .ris (Zotero/Scopus/WoS)")
    ap.add_argument("--dois",   help="Ficheiro .txt com um DOI por linha")
    ap.add_argument("--output", default=OUTPUT_DEFAULT)
    ap.add_argument("--email",  default=EMAIL, help="Email para polite pool OpenAlex")
    args = ap.parse_args()
    EMAIL = args.email

    if not args.ris and not args.dois:
        ap.print_help(); sys.exit(1)

    corpus = parse_ris(args.ris) if args.ris else parse_dois(args.dois)
    if not corpus:
        print("❌  Nenhum artigo carregado."); sys.exit(1)

    artigos_out, lista_aut = processar(corpus)
    exportar(artigos_out, lista_aut, args.output)

if __name__ == "__main__":
    main()
