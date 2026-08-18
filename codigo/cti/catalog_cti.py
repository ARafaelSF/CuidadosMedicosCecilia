# -*- coding: utf-8 -*-
"""Catalog CTI lab PDFs and compare with Todos.pdf."""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

from pypdf import PdfReader

sys.stdout.reconfigure(encoding="utf-8")

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from caminhos import CTI, TEMP

OUT = TEMP / "cti_catalog.json"

HEADER = re.compile(
    r"Médico:\s*Dr\(a\)\.\s*(?P<medico>.+?)\s+Local:\s*(?P<local>.+?)\s*"
    r"Data coleta:\s*(?P<data>\d{2}/\d{2}/\d{4}).+?"
    r"Prescrição:\s*(?P<presc>\d+)",
    re.S,
)
SECTION = re.compile(
    r"^(HEMOGRAMA|PROTEÍNA C REATIVA|GLICOSE|SÓDIO|POTÁSSIO|UREIA|CREATININA|"
    r"GASOMETRIA|COAGULOGRAMA|CÁLCIO|CLORO|ÁCIDO LÁTICO|MAGNÉSIO|FÓSFORO|"
    r"TGO|TGP|BILIRRUBINA|ALBUMINA|PROTEÍNAS TOTAIS|AMÔNIA|PCR|"
    r"ELEMENTOS ANORMAIS|UROCULTURA|HEMOCULTURA|GRAM|"
    r"TSH|T4|CORTISOL|FERRITINA|FERRO|LACTATO|"
    r"TROPONINA|CK-MB|CPK|"
    r"HEMOCULTURA -|ANTIBIOGRAMA|"
    r"PROTEÍNA C REATIVA - PCR|"
    r"TEMPO DE PROTOMBINA|RNI|PTT|"
    r"CÁLCIO IÔNICO|CÁLCIO TOTAL|"
    r"GASOMETRIA ARTERIAL|GASOMETRIA VENOSA|"
    r"CO2 Total|pH\.\.\.\.\.\.\.\.\.\.\.|"
    r"CULTURA DE|Pesquisa de|"
    r"AMILASE|LIPASE|"
    r"LEUCÓCITOS|PLAQUETAS|"
    r"HEMATÓCRITO|"
    r"PROTEÍNA C REATIVA)",
    re.M | re.I,
)
TITLE_LINE = re.compile(
    r"^(?P<title>[A-ZÁÉÍÓÚÂÊÔÃÕÇ0-9 /.-]{6,80}):?\s*$",
    re.M,
)


def page_text(page) -> str:
    t = page.extract_text() or ""
    t = t.replace("\xa0", " ")
    return t


def parse_header(text: str) -> dict:
    m = HEADER.search(text)
    if not m:
        # fallback piece by piece
        medico = re.search(r"Médico:\s*Dr\(a\)\.\s*(.+?)\s+Local:", text)
        local = re.search(r"Local:\s*(.+?)\s*\n", text)
        data = re.search(r"Data coleta:\s*(\d{2}/\d{2}/\d{4})", text)
        presc = re.search(r"Prescrição:\s*(\d+)", text)
        atend = re.search(r"Nº Atend\.:\s*([0-9.]+)", text)
        idade = re.search(r"Idade\s+(.+?)\s+Sexo", text)
        return {
            "medico": (medico.group(1).strip() if medico else None),
            "local": (local.group(1).strip() if local else None),
            "data": (data.group(1) if data else None),
            "presc": (presc.group(1) if presc else None),
            "atend": (atend.group(1) if atend else None),
            "idade": (idade.group(1).strip() if idade else None),
        }
    atend = re.search(r"Nº Atend\.:\s*([0-9.]+)", text)
    idade = re.search(r"Idade\s+(.+?)\s+Sexo", text)
    return {
        "medico": m.group("medico").strip(),
        "local": m.group("local").strip(),
        "data": m.group("data"),
        "presc": m.group("presc"),
        "atend": (atend.group(1) if atend else None),
        "idade": (idade.group(1).strip() if idade else None),
    }


def find_exams(text: str) -> list[str]:
    found = []
    patterns = [
        ("Hemograma", r"\bHEMOGRAMA\b"),
        ("PCR", r"PROTEÍNA C REATIVA"),
        ("Glicose", r"\bGLICOSE[.:]"),
        ("Sódio", r"\bSÓDIO"),
        ("Potássio", r"\bPOTÁSSIO"),
        ("Cloro", r"\bCLORO[.:]"),
        ("Cálcio iônico", r"CÁLCIO IÔNICO"),
        ("Cálcio", r"CÁLCIO TOTAL|\bCÁLCIO:"),
        ("Ureia", r"\bUREIA"),
        ("Creatinina", r"CREATININA"),
        ("Gasometria arterial", r"GASOMETRIA ARTERIAL"),
        ("Gasometria venosa", r"GASOMETRIA VENOSA"),
        ("Gasometria", r"\bGASOMETRIA\b"),
        ("Ácido lático", r"ÁCIDO LÁTICO|LACTATO"),
        ("Coagulograma", r"COAGULOGRAMA"),
        ("EAS", r"ELEMENTOS ANORMAIS E SEDIMENTOSCOPIA"),
        ("Gram urina", r"GRAM BACTERIOSCOPIA"),
        ("Urocultura", r"\bUROCULTURA\b"),
        ("Hemocultura", r"\bHEMOCULTURA"),
        ("TGO", r"\bTGO\b|ASPARTATO"),
        ("TGP", r"\bTGP\b|ALANINA"),
        ("Bilirrubinas", r"BILIRRUBINA"),
        ("Magnésio", r"MAGNÉSIO"),
        ("Fósforo", r"FÓSFORO"),
        ("Albumina", r"ALBUMINA"),
        ("Proteínas totais", r"PROTEÍNAS TOTAIS"),
        ("Amônia", r"AM[ÔO]NIA"),
        ("TSH", r"\bTSH\b"),
        ("T4 livre", r"T4 LIVRE"),
        ("Cortisol", r"CORTISOL"),
        ("Ferritina", r"FERRITINA"),
        ("Cultura", r"\bCULTURA DE"),
        ("Antibiograma", r"ANTIBIOGRAMA"),
        ("Troponina", r"TROPONINA"),
        ("CK", r"\bCK-MB\b|\bCPK\b"),
        ("Amilase", r"AMILASE"),
        ("Lipase", r"LIPASE"),
        ("Proteína", r"PROTEÍNA[.:]"),
        ("pH", r"^pH\.\.\."),
        ("HCO3", r"HCO3"),
    ]
    for name, pat in patterns:
        if re.search(pat, text, re.I | re.M):
            found.append(name)
    # unique preserve order
    seen = set()
    out = []
    for x in found:
        if x not in seen:
            seen.add(x)
            out.append(x)
    return out


def catalog_file(path: Path) -> dict:
    reader = PdfReader(str(path))
    pages_meta = []
    full = []
    for i, page in enumerate(reader.pages):
        t = page_text(page)
        full.append(t)
        h = parse_header(t)
        h["page"] = i + 1
        h["exams"] = find_exams(t)
        h["chars"] = len(t)
        pages_meta.append(h)
    joined = "\n".join(full)
    header0 = pages_meta[0] if pages_meta else {}
    exams = []
    for p in pages_meta:
        exams.extend(p.get("exams") or [])
    seen = set()
    exams_u = []
    for x in exams:
        if x not in seen:
            seen.add(x)
            exams_u.append(x)
    prescs = sorted({p["presc"] for p in pages_meta if p.get("presc")})
    datas = sorted({p["data"] for p in pages_meta if p.get("data")})
    medicos = sorted({p["medico"] for p in pages_meta if p.get("medico")})
    locais = sorted({p["local"] for p in pages_meta if p.get("local")})
    return {
        "file": path.name,
        "pages": len(reader.pages),
        "size": path.stat().st_size,
        "prescs": prescs,
        "datas": datas,
        "medicos": medicos,
        "locais": locais,
        "exams": exams_u,
        "pages_meta": pages_meta,
        "text": joined,
    }


def main():
    files = sorted(
        [p for p in CTI.iterdir() if p.suffix.lower() == ".pdf" and p.name.lower() != "todos.pdf"],
        key=lambda p: p.name,
    )
    records = []
    for p in files:
        rec = catalog_file(p)
        # drop full text from per-file dump later; keep for comparison
        records.append(rec)
        print(
            f"{rec['file']:12} p={rec['pages']:2} data={','.join(rec['datas']) or '-':12} "
            f"presc={','.join(rec['prescs']) or '-':8} exams={len(rec['exams'])}"
        )

    todos = catalog_file(CTI / "Todos.pdf")
    print("TODOS pages", todos["pages"], "prescs", len(todos["prescs"]), "datas", todos["datas"][:3], "...", todos["datas"][-3:])

    indiv_prescs = set()
    for r in records:
        indiv_prescs.update(r["prescs"])
    todos_prescs = set(todos["prescs"])
    missing_in_todos = sorted(indiv_prescs - todos_prescs)
    extra_in_todos = sorted(todos_prescs - indiv_prescs)
    print("indiv files", len(records), "indiv prescs", len(indiv_prescs))
    print("todos prescs", len(todos_prescs))
    print("missing_in_todos", missing_in_todos)
    print("extra_in_todos", extra_in_todos)
    print("page sum indiv", sum(r["pages"] for r in records), "todos pages", todos["pages"])

    dump = {
        "indiv": [
            {k: v for k, v in r.items() if k != "text"}
            for r in records
        ],
        "todos": {k: v for k, v in todos.items() if k != "text"},
        "missing_in_todos": missing_in_todos,
        "extra_in_todos": extra_in_todos,
    }
    OUT.write_text(json.dumps(dump, ensure_ascii=False, indent=2), encoding="utf-8")
    # keep texts separately for value extraction
    texts = {r["file"]: r["text"] for r in records}
    texts["Todos.pdf"] = todos["text"]
    (TEMP / "cti_texts.json").write_text(
        json.dumps(texts, ensure_ascii=False), encoding="utf-8"
    )
    print("wrote", OUT)


if __name__ == "__main__":
    main()
