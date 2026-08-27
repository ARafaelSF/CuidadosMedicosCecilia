# -*- coding: utf-8 -*-
"""Gera pastas mensais de renovação CEAF (LME + Monitoramento GH).

Uso:
  py -3 gerar_mes.py 2026-08
  py -3 gerar_mes.py 2026-09
"""
from __future__ import annotations

import shutil
import sys
from datetime import date
from pathlib import Path

import pymupdf

ROOT = Path(__file__).resolve().parents[2]  # _Organizado
BASE = ROOT / "Solicitação Medicamento SUS"
MODELOS = BASE / "Modelos"
LME_MODELO = MODELOS / "LME-EDITAVEL.pdf"
MON_MODELO = MODELOS / "Formulario-GH-MONITORAMENTO-EDITAVEL.pdf"

# --- Dados fixos da Cecília (complete o que faltar) ---
PACIENTE = {
    "nome": "Cecília Maria Albergaria Silva",
    "nome_social": "",
    "mae": "",  # preencher quando souber
    "dn": date(2020, 6, 30),
    "sexo": "F",
    "cns": "898006204638698",
    "altura_cm": "98",  # última medida (5a8m); atualizar no mês
    "peso_kg": "",  # atualizar no mês
}

MEDICO = {
    "nome": "Sayonara Figueiredo de Faria",
    "cns": "",  # CNS do médico, se tiver
    "crm": "CRM-MG",  # completar
}

MEDICAMENTO = {
    "nome": "Somatropina",
    "cid": "E23.0",
    "diagnostico": "Deficiência do hormônio de crescimento em criança",
    "anamnese": (
        "Paciente em acompanhamento por deficiência de hormônio de crescimento. "
        "Início do tratamento com somatropina em 18/04/2025. "
        "Solicita-se continuidade do tratamento conforme PCDT/CEAF."
    ),
}

# Três colunas do monitoramento (mais recentes conhecidas). Atualize a cada ciclo.
CONSULTAS = [
    {
        "data": "28/07/2025",
        "idade_cron": "5a1m",
        "peso": "",
        "altura": "92",
        "idade_altura": "",
        "idade_peso": "",
        "idade_ossea": "2a6m (fev/25)",
        "glicose": "77",
        "igf1": "122",
        "tsh": "1,92",
        "t4l": "1,20",
    },
    {
        "data": "03/02/2026",
        "idade_cron": "5a7m",
        "peso": "",
        "altura": "95",
        "idade_altura": "",
        "idade_peso": "",
        "idade_ossea": "3a6m (fev/26)",
        "glicose": "82",
        "igf1": "68",
        "tsh": "1,42",
        "t4l": "1,29",
    },
    {
        "data": "12/08/2026",
        "idade_cron": "6a1m",
        "peso": "",
        "altura": "98",
        "idade_altura": "",
        "idade_peso": "",
        "idade_ossea": "4a2m (jul/26)",
        "glicose": "80",
        "igf1": "173",
        "tsh": "1,50",
        "t4l": "1,23",
    },
]

VELOCIDADE_CM_ANO = "≈9"  # estimativa a partir da curva; médico confirma


def idade_em(ref: date) -> str:
    anos = ref.year - PACIENTE["dn"].year
    meses = ref.month - PACIENTE["dn"].month
    if ref.day < PACIENTE["dn"].day:
        meses -= 1
    if meses < 0:
        anos -= 1
        meses += 12
    return f"{anos}a{meses}m"


def set_text(doc: pymupdf.Document, name: str, value: str) -> None:
    if value is None or value == "":
        return
    for page in doc:
        for w in page.widgets() or []:
            if w.field_name == name and w.field_type_string == "Text":
                w.field_value = str(value)
                w.update()


def set_radio(doc: pymupdf.Document, name: str, which: int) -> None:
    """which=1 primeiro botão, which=2 segundo (ordem visual esquerda→direita)."""
    radios = []
    for page in doc:
        for w in page.widgets() or []:
            if w.field_name == name and w.field_type_string == "RadioButton":
                radios.append(w)
    radios.sort(key=lambda w: w.rect.x0)
    if not radios:
        return
    target = radios[0] if which == 1 else radios[min(1, len(radios) - 1)]
    on = target.on_state()
    for w in radios:
        w.field_value = on if w is target else False
        w.update()


def set_check(doc: pymupdf.Document, name: str, checked: bool = True) -> None:
    for page in doc:
        for w in page.widgets() or []:
            if w.field_name == name and w.field_type_string == "CheckBox":
                on = w.on_state()
                w.field_value = on if checked else False
                w.update()


def preencher_lme(dest: Path, ref: date) -> None:
    doc = pymupdf.open(LME_MODELO)
    set_text(doc, "Text3", PACIENTE["nome"])
    set_text(doc, "Text4", PACIENTE["nome_social"])
    set_text(doc, "Text5", PACIENTE["mae"])
    set_text(doc, "Text6", PACIENTE["peso_kg"])
    set_text(doc, "Text7", PACIENTE["altura_cm"])
    set_text(doc, "Text8", MEDICAMENTO["nome"])
    set_text(doc, "Text14", MEDICAMENTO["cid"])
    set_text(doc, "Text15", MEDICAMENTO["diagnostico"])
    set_text(doc, "Text31", MEDICAMENTO["anamnese"])
    # 12. tratamento prévio = SIM
    set_radio(doc, "Button29", 2)
    set_text(doc, "Text30", "Em uso de somatropina desde 18/04/2025 (continuidade).")
    # 13. incapaz = SIM (menor)
    set_radio(doc, "Button38", 2)
    set_text(doc, "Text39", PACIENTE["mae"] or "Responsável legal (mãe)")
    set_text(doc, "Text46", MEDICO["nome"])
    set_text(doc, "Text47", MEDICO["cns"])
    set_text(doc, "Text48", ref.strftime("%d/%m/%Y"))
    # 18. preenchido por Responsável
    set_check(doc, "Button70", True)
    # 21. documento = CNS
    set_radio(doc, "Button82", 2)
    set_text(doc, "Text51", PACIENTE["cns"])
    doc.need_appearances(True)
    doc.save(dest, incremental=False, encryption=pymupdf.PDF_ENCRYPT_KEEP)
    doc.close()


def preencher_monitoramento(dest: Path, ref: date) -> None:
    doc = pymupdf.open(MON_MODELO)
    set_text(doc, "Text52", PACIENTE["nome"])
    set_text(doc, "Text53", PACIENTE["nome_social"])
    set_radio(doc, "Button56", 1 if PACIENTE["sexo"] == "F" else 2)
    dn = PACIENTE["dn"]
    set_text(doc, "Text57", f"{dn.day:02d}")
    set_text(doc, "Text59", f"{dn.month:02d}")
    set_text(doc, "Text58", f"{dn.year}")
    set_text(doc, "Text54", idade_em(ref))
    set_text(doc, "Text55", VELOCIDADE_CM_ANO)

    # Campos da tabela (PDF não tem campo editável na linha "DATA:" — só nas métricas)
    cols = [
        ["Text60", "Text61", "Text62", "Text63", "Text64", "Text65", "Text66", "Text70", "Text69", "Text68"],
        ["Text80", "Text71", "Text72", "Text73", "Text74", "Text75", "Text76", "Text77", "Text78", "Text79"],
        ["Text82", "Text83", "Text84", "Text85", "Text86", "Text87", "Text88", "Text89", "Text90", "Text91"],
    ]
    keys = [
        "idade_cron", "peso", "altura", "idade_altura", "idade_peso",
        "idade_ossea", "glicose", "igf1", "tsh", "t4l",
    ]
    for col_fields, consult in zip(cols, CONSULTAS):
        for field, key in zip(col_fields, keys):
            set_text(doc, field, consult.get(key, ""))

    datas = " | ".join(c["data"] for c in CONSULTAS)
    set_text(
        doc,
        "Text92",
        f"Datas das colunas (esq→dir): {datas}. "
        "Valores de laboratório e idade óssea conforme laudos na pasta Exames. "
        "Altura conforme curva/pontos_altura. Peso e dose a completar pelo médico.",
    )
    set_text(doc, "Text94", f"{ref.day:02d}")
    set_text(doc, "Text95", f"{ref.month:02d}")
    set_text(doc, "Text96", f"{ref.year}")

    doc.need_appearances(True)
    doc.save(dest, incremental=False, encryption=pymupdf.PDF_ENCRYPT_KEEP)
    doc.close()


def gerar(ym: str) -> Path:
    year, month = map(int, ym.split("-"))
    ref = date(year, month, 1)
    pasta = BASE / ym
    pasta.mkdir(parents=True, exist_ok=True)
    lme = pasta / f"LME-{ym}.pdf"
    mon = pasta / f"Monitoramento-GH-{ym}.pdf"
    preencher_lme(lme, ref)
    preencher_monitoramento(mon, ref)
    readme = pasta / "LEIA-ME.txt"
    readme.write_text(
        f"Renovação CEAF — {ym}\n"
        f"Paciente: {PACIENTE['nome']} (DN {PACIENTE['dn'].strftime('%d/%m/%Y')})\n"
        f"Idade em {ym}: {idade_em(ref)}\n\n"
        "Arquivos:\n"
        f"- {lme.name} (pré-preenchido; médico completa CNES, quantidades, assinatura)\n"
        f"- {mon.name} (monitoramento/continuidade; médico confere e assina)\n\n"
        "Ainda falta preencher à mão/digital: nome da mãe, peso, CNES, CNS do médico,\n"
        "quantidades mensais da somatropina, dose na orientação de aplicação, assinaturas.\n"
        "Checklist SES-MG: renovação a cada 6 meses (LME + receita + este formulário).\n",
        encoding="utf-8",
    )
    return pasta


def garantir_modelos() -> None:
    MODELOS.mkdir(parents=True, exist_ok=True)
    src_lme = BASE / "LME-EDITAVEL.pdf"
    src_mon = BASE / "Formulario Especifico - GH - MONITORAMENTO CONTINUIDADE - SES-MG.pdf"
    if src_lme.exists() and not LME_MODELO.exists():
        shutil.copy2(src_lme, LME_MODELO)
    if src_mon.exists() and not MON_MODELO.exists():
        shutil.copy2(src_mon, MON_MODELO)
    if not LME_MODELO.exists() or not MON_MODELO.exists():
        raise SystemExit(f"Faltam modelos em {MODELOS}")


def main() -> None:
    ym = sys.argv[1] if len(sys.argv) > 1 else "2026-08"
    garantir_modelos()
    pasta = gerar(ym)
    print("OK", pasta)


if __name__ == "__main__":
    main()
