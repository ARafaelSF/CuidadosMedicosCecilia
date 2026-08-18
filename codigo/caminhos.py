# -*- coding: utf-8 -*-
"""Caminhos da pasta _Organizado. Importe daqui; não hardcode OneDrive."""
from pathlib import Path

CODIGO = Path(__file__).resolve().parent
ROOT = CODIGO.parent  # .../Cecília/_Organizado
EXAMES = ROOT / "Exames"
RELATORIOS = ROOT / "Relatórios"
CTI = EXAMES / "Exames Laboratorias - CTI"
DOCUMENTOS = ROOT / "Documentos"

TEMP = Path.home() / "AppData" / "Local" / "Temp"
DEST_EVOL_XLSX = TEMP / "EvolucaoExames.xlsx"
DEST_EVOL_XLSM_TMP = TEMP / "EvolucaoExames.xlsm"
DEST_EVOL_XLSM = EXAMES / "Evolução Exames.xlsm"

RESUMO_EXAMES = EXAMES / "Resumo Exames.xlsx"
RESUMO_RELATORIOS = RELATORIOS / "Resumo Relatórios.xlsx"
RESUMO_CTI = CTI / "Resumo Exames CTI.xlsx"
EVOL_CTI = CTI / "Evolução Exames CTI.xlsx"

TIPOS_EXAME = [
    "Sangue", "Urina", "Imagem", "Pezinho", "Audiologia", "EEG",
    "Suor", "Alta", "Documento",
]
TIPOS_RELATORIO = [
    "Pediatria", "Endocrinologia", "Neurologia", "Genética", "Ortopedia",
    "Oftalmologia", "Fonoaudiologia", "Fisioterapia", "TO", "Psicologia",
    "Pedagogia", "Nutrição", "Funcional", "Natação", "Equoterapia",
]
