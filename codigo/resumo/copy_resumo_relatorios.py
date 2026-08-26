# -*- coding: utf-8 -*-
"""Copia ResumoRelatorios.xlsm do TEMP para Relatórios/ (evita mojibake do PowerShell)."""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from caminhos import RELATORIOS, TEMP  # noqa: E402

src = TEMP / "ResumoRelatorios.xlsm"
dst = RELATORIOS / "Resumo Relatórios.xlsm"
xlsx = RELATORIOS / "Resumo Relatórios.xlsx"
if not src.exists():
    raise SystemExit(f"Fonte inexistente: {src}")
shutil.copy2(src, dst)
print("SAVED_XLSM", dst, dst.stat().st_size)
if xlsx.exists():
    xlsx.unlink()
    print("REMOVED_XLSX")
