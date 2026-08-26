# -*- coding: utf-8 -*-
"""Evolução de exames: uma tabela por item, só datas reais, escolha Sim/Não."""
from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path
import re
import sys

import xlsxwriter

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from caminhos import DEST_EVOL_XLSX, EXAMES

DEST = DEST_EVOL_XLSX

ABA_ESCOLHER = "Escolher"
ABA_DADOS_COMP = "Dados Completo"
ABA_DADOS_SEL = "Dados Selecionados"
ABA_GRAF_COMP = "Graficos Completo"
ABA_GRAF_SEL = "Graficos Selecionados"
ABA_UNICOS = "Unicos"

# Cada tabela: só as coletas que existem (sem coluna vazia).
TABLES = [
    {
        "grupo": "Crescimento",
        "titulo": "IGF-1 - Somatomedina C",
        "mostra": "Hormônio de crescimento ao longo do tempo",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2021, 4, 20), 15.0, "ng/mL", "18 a 172 (nessa idade, na época)", "Abaixo da faixa"],
            [date(2021, 12, 18), 18.0, "ng/mL", "18 a 172", "No piso da faixa"],
            [date(2024, 10, 15), 39, "ng/mL", "4–6 anos: cerca de 22 a 208", ""],
            [date(2025, 7, 28), 122, "ng/mL", "4–6 anos: cerca de 22 a 208", "Maior valor da série"],
            [date(2026, 2, 3), 68, "ng/mL", "4–6 anos: cerca de 22 a 208", "Recuou em relação a jul/25"],
            [date(2026, 8, 17), 173, "ng/mL", "♀ 4–6 a: 35 a 232 (faixa nova neste laudo)", "Maior valor da série; coleta 17/08"],
        ],
    },
    {
        "grupo": "Crescimento",
        "titulo": "IGFBP-3 - proteína ligadora-3 do IGF",
        "mostra": "Proteína que transporta o IGF-1",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2021, 4, 20), 1.00, "mcg/mL", "muda com a idade (ver laudo)", ""],
            [date(2021, 12, 18), 1.50, "mcg/mL", "muda com a idade (ver laudo)", ""],
            [date(2024, 10, 15), 2.0, "mcg/mL", "muda com a idade (ver laudo)", ""],
            [date(2025, 7, 28), 3.6, "mcg/mL", "muda com a idade (ver laudo)", "Acompanhou a subida do IGF-1"],
            [date(2026, 8, 12), 3.3, "mcg/mL", "6 a: 1,3 a 5,6", ""],
        ],
    },
    {
        "grupo": "Crescimento",
        "titulo": "GH - hormônio de crescimento (basal)",
        "mostra": "Dosagem isolada (não substitui o teste de estímulo)",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2021, 12, 7), 1.08, "ng/mL", "basal isolado não define deficiência", ""],
            [date(2025, 1, 13), 2.17, "ng/mL", "basal do teste com glucagon", "Ver tabela do teste"],
        ],
    },
    {
        "grupo": "Crescimento",
        "titulo": "Teste do GH com glucagon",
        "mostra": "13/01/2025 · laudo pede GH ≥ 5 ng/mL em qualquer tempo",
        "arquivo": "Sangue - 2025-01-13 - Cristiano",
        "cols": ["Tempo", "GH (ng/mL)", "Glicose (mg/dL)", "Cortisol (mcg/dL)", "Nota"],
        "linhas": [
            ["Basal", 2.17, None, 10.3, ""],
            ["60 min", 0.43, 199, 12.1, ""],
            ["90 min", 1.06, 112, 15.6, ""],
            ["120 min", 3.07, 73, 29.0, "Pico do GH"],
            ["180 min", 0.97, 70, 22.2, "Pico 3,07 — abaixo de 5"],
        ],
    },
    {
        "grupo": "Idade óssea",
        "titulo": "RX mão/punho (Greulich-Pyle, feminino)",
        "mostra": "Idade cronológica × idade óssea",
        "arq_nota": "Punho",
        "cols": ["Data", "Cronológica", "Óssea", "Atraso", "Nota"],
        "linhas": [
            [date(2025, 2, 10), "4a7m", "2a6m", "25 meses", ""],
            [date(2026, 2, 18), "5a7m", "3a6m", "25 meses", "Atraso estável"],
            [date(2026, 7, 29), "6a1m", "4a2m", "23 meses", "Ainda cerca de 2 anos"],
        ],
    },
    {
        "grupo": "Tireoide",
        "titulo": "TSH ultra sensível",
        "mostra": "Hormônio que regula a tireoide",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2021, 4, 20), 3.43, "µUI/mL", "2–12 anos: 0,70 a 6,55", "Dentro da faixa"],
            [date(2021, 12, 7), 2.56, "µUI/mL", "2–12 anos: 0,70 a 6,55", ""],
            [date(2024, 10, 15), 2.36, "µUI/mL", "2–12 anos: 0,70 a 6,55", ""],
            [date(2025, 7, 28), 1.92, "µUI/mL", "2–12 anos: 0,70 a 6,55", ""],
            [date(2026, 2, 3), 1.42, "µUI/mL", "2–12 anos: 0,70 a 6,55", ""],
            [date(2026, 8, 12), 1.50, "µUI/mL", "2–12 anos: 0,70 a 6,55", ""],
        ],
    },
    {
        "grupo": "Tireoide",
        "titulo": "T4 livre",
        "mostra": "Hormônio da tireoide",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2021, 4, 20), 1.22, "ng/dL", "cerca de 0,85 a 1,67", "Dentro da faixa"],
            [date(2021, 12, 7), 1.07, "ng/dL", "cerca de 0,85 a 1,67", ""],
            [date(2024, 10, 15), 1.15, "ng/dL", "cerca de 0,85 a 1,67", ""],
            [date(2025, 7, 28), 1.20, "ng/dL", "cerca de 0,85 a 1,67", ""],
            [date(2026, 2, 3), 1.29, "ng/dL", "cerca de 0,85 a 1,67", ""],
            [date(2026, 8, 12), 1.23, "ng/dL", "2–12 a: 1,05 a 1,67 (faixa nova a partir de 06/07/26)", ""],
        ],
    },
    {
        "grupo": "Hemograma",
        "titulo": "Hemoglobina",
        "mostra": "Série vermelha",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2020, 10, 17), 12.3, "g/dL", "10,0 a 14,0 (nessa idade)", ""],
            [date(2021, 4, 20), 14.7, "g/dL", "10,5 a 13,5 na época", "Acima da faixa da época"],
            [date(2021, 12, 7), 13.9, "g/dL", "muda com a idade", ""],
            [date(2022, 2, 9), 12.7, "g/dL", "10,5 a 13,5 na época", ""],
            [date(2022, 3, 1), 12.5, "g/dL", "10,5 a 13,5 na época", ""],
            [date(2022, 3, 3), 12.3, "g/dL", "10,5 a 13,5 na época", ""],
            [date(2024, 10, 15), 12.2, "g/dL", "11,5 a 13,5", ""],
            [date(2025, 5, 28), 12.7, "g/dL", "11,5 a 13,5", ""],
            [date(2025, 7, 28), 13.2, "g/dL", "11,5 a 13,5", ""],
            [date(2026, 2, 3), 12.7, "g/dL", "11,5 a 13,5", ""],
            [date(2026, 7, 29), 13.0, "g/dL", "11,5 a 13,5", ""],
            [date(2026, 8, 12), 12.9, "g/dL", "11,5 a 13,5", ""],
        ],
    },
    {
        "grupo": "Hemograma",
        "titulo": "Hemácias",
        "mostra": "Contagem de hemácias (eritrograma)",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2020, 10, 17), 4.73, "10⁶/mm³", "3,20 a 4,50", "Acima da faixa"],
            [date(2021, 4, 20), 5.35, "10⁶/mm³", "3,70 a 5,70", ""],
            [date(2021, 12, 7), 5.07, "10⁶/mm³", "3,70 a 5,20", ""],
            [date(2022, 2, 9), 4.52, "10⁶/mm³", "3,80 a 5,00", ""],
            [date(2022, 3, 1), 4.50, "10⁶/mm³", "3,80 a 5,00", ""],
            [date(2022, 3, 3), 4.48, "10⁶/mm³", "3,80 a 5,00", ""],
            [date(2024, 10, 15), 4.95, "10⁶/mm³", "3,90 a 5,30", ""],
            [date(2025, 5, 28), 4.90, "10⁶/mm³", "3,90 a 5,30", ""],
            [date(2025, 7, 28), 4.95, "10⁶/mm³", "3,90 a 5,30", ""],
            [date(2026, 2, 3), 4.70, "10⁶/mm³", "3,90 a 5,30", ""],
            [date(2026, 7, 29), 4.78, "10⁶/mm³", "3,90 a 5,30", ""],
            [date(2026, 8, 12), 4.86, "10⁶/mm³", "3,90 a 5,30", ""],
        ],
    },
    {
        "grupo": "Hemograma",
        "titulo": "Hematócrito",
        "mostra": "Volume de hemácias no sangue",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2020, 10, 17), 37.5, "%", "28 a 42", ""],
            [date(2021, 4, 20), 44.2, "%", "32 a 40", "Acima da faixa"],
            [date(2021, 12, 7), 41.0, "%", "33 a 42", ""],
            [date(2022, 2, 9), 36.7, "%", "24 a 44", ""],
            [date(2022, 3, 1), 35.9, "%", "24 a 44", ""],
            [date(2022, 3, 3), 35.9, "%", "24 a 44", ""],
            [date(2024, 10, 15), 40.7, "%", "34 a 40", "Acima da faixa"],
            [date(2025, 5, 28), 38.8, "%", "34 a 40", ""],
            [date(2025, 7, 28), 40.5, "%", "34 a 40", "Acima da faixa"],
            [date(2026, 2, 3), 38.0, "%", "34 a 40", ""],
            [date(2026, 7, 29), 38.4, "%", "34 a 40", ""],
            [date(2026, 8, 12), 38.5, "%", "34 a 40", ""],
        ],
    },
    {
        "grupo": "Hemograma",
        "titulo": "VCM",
        "mostra": "Tamanho médio das hemácias",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2020, 10, 17), 79.3, "fL", "78 a 100", ""],
            [date(2021, 4, 20), 82.6, "fL", "74 a 88", ""],
            [date(2021, 12, 7), 80.9, "fL", "75 a 88", ""],
            [date(2022, 2, 9), 81, "fL", "78 a 92", ""],
            [date(2022, 3, 1), 80, "fL", "78 a 92", ""],
            [date(2022, 3, 3), 80, "fL", "78 a 92", ""],
            [date(2024, 10, 15), 82.2, "fL", "75 a 87", ""],
            [date(2025, 5, 28), 79.2, "fL", "75 a 87", ""],
            [date(2025, 7, 28), 81.8, "fL", "75 a 87", ""],
            [date(2026, 2, 3), 80.9, "fL", "75 a 87", ""],
            [date(2026, 7, 29), 80.3, "fL", "75 a 87", ""],
            [date(2026, 8, 12), 79.2, "fL", "75 a 87", ""],
        ],
    },
    {
        "grupo": "Hemograma",
        "titulo": "HCM e CHCM",
        "mostra": "Índices de cor das hemácias",
        "cols": ["Data", "HCM (pg)", "CHCM", "Faixa", "Nota"],
        "linhas": [
            [date(2020, 10, 17), 26.0, 32.8, "HCM 26–34 · CHCM 30–36", ""],
            [date(2021, 4, 20), 27.5, 33.3, "HCM 25–34 · CHCM 30–36", ""],
            [date(2021, 12, 7), 27.4, 33.9, "HCM 26–33 · CHCM 31–36", ""],
            [date(2022, 2, 9), 28, 35, "HCM 26–31 · CHCM 31–34", "CHCM no teto / um pouco acima"],
            [date(2022, 3, 1), 28, 35, "HCM 26–31 · CHCM 31–34", "CHCM no teto / um pouco acima"],
            [date(2022, 3, 3), 27, 34, "HCM 26–31 · CHCM 31–34", ""],
            [date(2024, 10, 15), 24.6, 30.0, "HCM 24–30 · CHCM 32–36 g/dL", "CHCM abaixo da faixa"],
            [date(2025, 5, 28), 25.9, 32.7, "HCM 24–30 · CHCM 32–36 g/dL", ""],
            [date(2025, 7, 28), 26.7, 32.6, "HCM 24–30 · CHCM 32–36 g/dL", ""],
            [date(2026, 2, 3), 27.0, 33.4, "HCM 24–30 · CHCM 32–36 g/dL", ""],
            [date(2026, 7, 29), 27.2, 33.9, "HCM 24–30 · CHCM 32–36 g/dL", ""],
            [date(2026, 8, 12), 26.5, 33.5, "HCM 24–30 · CHCM 32–36 g/dL", ""],
        ],
    },
    {
        "grupo": "Hemograma",
        "titulo": "RDW",
        "mostra": "Variação de tamanho das hemácias",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2020, 10, 17), 13.7, "%", "típico ~11,5 a 15", ""],
            [date(2021, 4, 20), 13.5, "%", "típico ~11,5 a 15", ""],
            [date(2021, 12, 7), 13.5, "%", "típico ~11,5 a 15", ""],
            [date(2022, 2, 9), 12.4, "%", "11,3 a 15,0", ""],
            [date(2022, 3, 1), 12.1, "%", "11,3 a 15,0", ""],
            [date(2022, 3, 3), 12.1, "%", "11,3 a 15,0", ""],
            [date(2024, 10, 15), 14.6, "%", "11,5 a 14,8", "No teto da faixa"],
            [date(2025, 5, 28), 14.1, "%", "11,5 a 14,8", ""],
            [date(2025, 7, 28), 13.1, "%", "11,5 a 14,8", ""],
            [date(2026, 2, 3), 13.1, "%", "11,5 a 14,8", ""],
            [date(2026, 7, 29), 13.2, "%", "11,5 a 14,8", ""],
            [date(2026, 8, 12), 12.4, "%", "11,5 a 14,8", ""],
        ],
    },
    {
        "grupo": "Hemograma",
        "titulo": "Leucócitos",
        "mostra": "Glóbulos brancos",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2020, 10, 17), 6350, "/mm³", "cerca de 5.000 a 14.500", ""],
            [date(2021, 4, 20), 9590, "/mm³", "cerca de 6.000 a 11.000 na época", ""],
            [date(2021, 12, 7), 7420, "/mm³", "cerca de 6.000 a 11.000 na época", ""],
            [date(2022, 2, 9), 9400, "/mm³", "faixa infantil do laudo", ""],
            [date(2022, 3, 1), 7100, "/mm³", "faixa infantil do laudo", ""],
            [date(2022, 3, 3), 8200, "/mm³", "faixa infantil do laudo", ""],
            [date(2024, 10, 15), 9930, "/mm³", "5.000 a 14.500", ""],
            [date(2025, 5, 28), 8950, "/mm³", "5.000 a 14.500", ""],
            [date(2025, 7, 28), 13580, "/mm³", "5.000 a 14.500", "No teto da faixa"],
            [date(2026, 2, 3), 8810, "/mm³", "5.000 a 14.500", ""],
            [date(2026, 7, 29), 4310, "/mm³", "5.000 a 14.500", "Abaixo da faixa"],
            [date(2026, 8, 12), 7560, "/mm³", "5.000 a 14.500", "Voltou para a faixa"],
        ],
    },
    {
        "grupo": "Hemograma",
        "titulo": "Neutrófilos",
        "mostra": "Segmentados (absolutos). Parte do leucograma",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2020, 10, 17), 1270, "/mm³", "1.000 a 8.500", ""],
            [date(2021, 4, 20), 1055, "/mm³", "1.500 a 8.500", "Abaixo da faixa"],
            [date(2021, 12, 7), 1484, "/mm³", "1.500 a 8.500", "Abaixo da faixa"],
            [date(2022, 2, 9), 4774, "/mm³", "2.000 a 5.000", ""],
            [date(2022, 3, 1), 2785, "/mm³", "2.000 a 5.000", ""],
            [date(2022, 3, 3), 2545, "/mm³", "2.000 a 5.000", ""],
            [date(2024, 10, 15), 4600, "/mm³", "1.500 a 8.000", ""],
            [date(2025, 5, 28), 3230, "/mm³", "1.500 a 8.000", ""],
            [date(2025, 7, 28), 6720, "/mm³", "1.500 a 8.000", ""],
            [date(2026, 2, 3), 3990, "/mm³", "1.500 a 8.000", ""],
            [date(2026, 7, 29), 2240, "/mm³", "1.500 a 8.000", ""],
            [date(2026, 8, 12), 2440, "/mm³", "1.500 a 8.000", ""],
        ],
    },
    {
        "grupo": "Hemograma",
        "titulo": "Linfócitos",
        "mostra": "Absolutos. Parte do leucograma",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2020, 10, 17), 4572, "/mm³", "4.000 a 13.500", ""],
            [date(2021, 4, 20), 7672, "/mm³", "4.000 a 10.500", ""],
            [date(2021, 12, 7), 5194, "/mm³", "1.500 a 7.000", ""],
            [date(2022, 2, 9), 3557, "/mm³", "5.500 a 9.000 (laudo PA)", "Abaixo da faixa da época"],
            [date(2022, 3, 1), 3856, "/mm³", "5.500 a 9.000 (laudo PA)", "Abaixo da faixa da época"],
            [date(2022, 3, 3), 4680, "/mm³", "5.500 a 9.000 (laudo PA)", "Abaixo da faixa da época"],
            [date(2024, 10, 15), 4260, "/mm³", "1.500 a 7.000", ""],
            [date(2025, 5, 28), 4700, "/mm³", "1.500 a 7.000", ""],
            [date(2025, 7, 28), 5270, "/mm³", "1.500 a 7.000", ""],
            [date(2026, 2, 3), 3690, "/mm³", "1.500 a 7.000", ""],
            [date(2026, 7, 29), 1030, "/mm³", "1.500 a 7.000", "Abaixo da faixa"],
            [date(2026, 8, 12), 4180, "/mm³", "1.500 a 7.000", "Voltou para a faixa"],
        ],
    },
    {
        "grupo": "Hemograma",
        "titulo": "Monócitos",
        "mostra": "Absolutos. Parte do leucograma",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2020, 10, 17), 444, "/mm³", "80 a 1.100", ""],
            [date(2021, 4, 20), 671, "/mm³", "80 a 1.000", ""],
            [date(2021, 12, 7), 668, "/mm³", "80 a 800", ""],
            [date(2022, 2, 9), 936, "/mm³", "200 a 1.300", ""],
            [date(2022, 3, 1), 428, "/mm³", "200 a 1.300", ""],
            [date(2022, 3, 3), 903, "/mm³", "200 a 1.300", ""],
            [date(2024, 10, 15), 940, "/mm³", "200 a 1.000", ""],
            [date(2025, 5, 28), 820, "/mm³", "200 a 1.000", ""],
            [date(2025, 7, 28), 1420, "/mm³", "200 a 1.000", "Acima da faixa"],
            [date(2026, 2, 3), 910, "/mm³", "200 a 1.000", ""],
            [date(2026, 7, 29), 1030, "/mm³", "200 a 1.000", "Acima da faixa"],
            [date(2026, 8, 12), 800, "/mm³", "200 a 1.000", "Voltou para a faixa"],
        ],
    },
    {
        "grupo": "Hemograma",
        "titulo": "Eosinófilos",
        "mostra": "Absolutos. Parte do leucograma",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2020, 10, 17), 64, "/mm³", "40 a 850", ""],
            [date(2021, 4, 20), 192, "/mm³", "40 a 700", ""],
            [date(2021, 12, 7), 74, "/mm³", "40 a 650", ""],
            [date(2022, 2, 9), 94, "/mm³", "70 a 650", ""],
            [date(2022, 3, 1), 71, "/mm³", "70 a 650", ""],
            [date(2024, 10, 15), 100, "/mm³", "100 a 1.000", "No piso"],
            [date(2025, 5, 28), 140, "/mm³", "100 a 1.000", ""],
            [date(2025, 7, 28), 100, "/mm³", "100 a 1.000", "No piso"],
            [date(2026, 2, 3), 190, "/mm³", "100 a 1.000", ""],
            [date(2026, 7, 29), 0, "/mm³", "100 a 1.000", "Abaixo da faixa"],
            [date(2026, 8, 12), 100, "/mm³", "100 a 1.000", "No piso"],
        ],
    },
    {
        "grupo": "Hemograma",
        "titulo": "Basófilos",
        "mostra": "Absolutos. Parte do leucograma (só nos laudos que trazem o valor)",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2024, 10, 15), 30, "/mm³", "até 200", ""],
            [date(2025, 5, 28), 60, "/mm³", "até 200", ""],
            [date(2025, 7, 28), 70, "/mm³", "até 200", ""],
            [date(2026, 2, 3), 30, "/mm³", "até 200", ""],
            [date(2026, 7, 29), 10, "/mm³", "até 200", ""],
            [date(2026, 8, 12), 40, "/mm³", "até 200", ""],
        ],
    },
    {
        "grupo": "Hemograma",
        "titulo": "Plaquetas",
        "mostra": "Coagulação",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2020, 10, 17), 180000, "/mm³", "150.000 a 450.000", ""],
            [date(2021, 4, 20), 185000, "/mm³", "150.000 a 450.000", ""],
            [date(2021, 12, 7), 231000, "/mm³", "150.000 a 450.000", ""],
            [date(2022, 2, 9), 303000, "/mm³", "200.000 a 500.000 no laudo", ""],
            [date(2022, 3, 1), 350000, "/mm³", "200.000 a 500.000 no laudo", ""],
            [date(2022, 3, 3), 367000, "/mm³", "200.000 a 500.000 no laudo", ""],
            [date(2024, 10, 15), 379000, "/mm³", "150.000 a 450.000", ""],
            [date(2025, 5, 28), 365000, "/mm³", "150.000 a 450.000", ""],
            [date(2025, 7, 28), 401000, "/mm³", "150.000 a 450.000", ""],
            [date(2026, 2, 3), 309000, "/mm³", "150.000 a 450.000", ""],
            [date(2026, 7, 29), 202000, "/mm³", "150.000 a 450.000", ""],
            [date(2026, 8, 12), 458000, "/mm³", "150.000 a 450.000", "Acima da faixa"],
        ],
    },
    {
        "grupo": "Hemograma",
        "titulo": "Fator XIII",
        "mostra": "Único laudo disponível",
        "arquivo": "Sangue - 2020-10-17 - Fator XIII",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2020, 10, 17), 28.0, "%", "70 a 140%", "Abaixo da faixa"],
        ],
    },
    {
        "grupo": "Hemograma",
        "titulo": "PCR - proteína C reativa quantitativa",
        "mostra": "Marcador de inflamação",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2022, 2, 9), 1.9, "mg/L", "0 a 10", "PA"],
            [date(2022, 3, 1), 1.0, "mg/L", "0 a 10", "PA"],
            [date(2026, 7, 29), 4.0, "mg/L", "< 10", "Limite inferior do ensaio"],
        ],
    },
    {
        "grupo": "Urina",
        "titulo": "EAS, Gram e culturas",
        "mostra": "PA e internação 2022 + urina 2023",
        "arquivo": "Sangue 2022-02-09 · Urina 2022-03-01 · Sangue 2022-03-02 (culturas) · Urina 2023-02-14",
        "cols": ["Data", "Exame", "Resumo", "", "Nota"],
        "largo": True,
        "linhas": [
            [date(2022, 2, 9), "EAS + Gram", "Cetônicos negativos · piócitos 5–10 · Gram sem bactérias", "", "No mesmo PDF do hemograma"],
            [date(2022, 3, 1), "EAS + Gram", "Cetônicos +++ · hemácias 5–10 · raros bastonetes GN", "", "PA"],
            [date(2022, 3, 2), "Urocultura e hemocultura", "Ambas negativas", "", "Internação"],
            [date(2023, 2, 14), "EAS", "Cetônicos 10 mg/dL · densidade 1,004 · resto sem alteração", "", "Ambulatorial"],
        ],
    },
    {
        "grupo": "Ferro e vitaminas",
        "titulo": "Ferritina sérica",
        "mostra": "Reserva de ferro",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2024, 10, 15), 11.6, "ng/mL", "6 meses–15 anos ♀: cerca de 7 a 140", "Baixa"],
            [date(2025, 7, 28), 6.9, "ng/mL", "6 meses–15 anos ♀: cerca de 7 a 140", "No piso da faixa"],
            [date(2026, 2, 3), 19.5, "ng/mL", "6 meses–15 anos ♀: cerca de 7 a 140", "Subiu"],
            [date(2026, 8, 12), 44.1, "ng/mL", "6 meses–15 anos ♀: cerca de 7 a 140", "Maior valor da série"],
        ],
    },
    {
        "grupo": "Ferro e vitaminas",
        "titulo": "Ferro sérico",
        "mostra": "Ferro circulante",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2025, 7, 28), 44, "mcg/dL", "1–12 anos: cerca de 50 a 120", "Abaixo da faixa"],
            [date(2026, 2, 3), 83, "mcg/dL", "1–12 anos: cerca de 50 a 120", "Normalizou"],
            [date(2026, 8, 12), 123, "mcg/dL", "1–12 anos: cerca de 50 a 120", "No teto / um pouco acima"],
        ],
    },
    {
        "grupo": "Ferro e vitaminas",
        "titulo": "Índice de saturação da transferrina",
        "mostra": "Quanto do transporte de ferro está ocupado",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2025, 7, 28), 11, "%", "mulheres: 15 a 50%", "Abaixo da faixa"],
            [date(2026, 8, 12), 37, "%", "mulheres: 15 a 50%", "Normalizou"],
        ],
    },
    {
        "grupo": "Ferro e vitaminas",
        "titulo": "Capacidade total de combinação do ferro",
        "mostra": "CTFF / TIBC",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2026, 2, 3), 332, "mcg/dL", "250 a 425", ""],
            [date(2026, 8, 12), 333, "mcg/dL", "250 a 425", ""],
        ],
    },
    {
        "grupo": "Ferro e vitaminas",
        "titulo": "25-Hidroxivitamina D",
        "mostra": "Dosagens disponíveis",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2021, 12, 7), 29.9, "ng/mL", "adequado 20–60 · ideal grupos de risco 30–60", "Adequado"],
            [date(2024, 10, 15), 28.9, "ng/mL", "adequado 20–60 · ideal grupos de risco 30–60", "Adequado"],
            [date(2026, 8, 12), 28.2, "ng/mL", "adequado 20–60 · ideal grupos de risco 30–60", "Estável"],
        ],
    },
    {
        "grupo": "Ferro e vitaminas",
        "titulo": "Vitamina B12",
        "mostra": "Única dosagem na pasta",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2025, 5, 28), 672, "pg/mL", "172 a 890", "Dentro da faixa"],
        ],
    },
    {
        "grupo": "Química",
        "titulo": "Glicose - jejum",
        "mostra": "Açúcar no sangue",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2021, 4, 20), 94, "mg/dL", "60 a 99", ""],
            [date(2021, 12, 7), 60, "mg/dL", "60 a 99", "No piso da faixa"],
            [date(2022, 3, 3), 97, "mg/dL", "jejum 60 a 99", "Arterial, sem jejum (internação)"],
            [date(2024, 10, 15), 87, "mg/dL", "60 a 99", ""],
            [date(2025, 7, 28), 77, "mg/dL", "60 a 99", ""],
            [date(2026, 2, 3), 82, "mg/dL", "60 a 99", ""],
            [date(2026, 8, 12), 80, "mg/dL", "60 a 99", ""],
        ],
    },
    {
        "grupo": "Química",
        "titulo": "Sódio e potássio",
        "mostra": "Eletrólitos",
        "cols": ["Data", "Sódio (mEq/L)", "Potássio (mEq/L)", "Faixa", "Nota"],
        "linhas": [
            [date(2020, 10, 17), 135, 5.1, "Na 136–145 · K 3,5–5,1", "Na no limite inferior"],
            [date(2021, 4, 20), 136, 4.9, "Na 136–145 · K 3,5–5,1", ""],
            [date(2021, 12, 7), 138, 5.1, "Na 136–145 · K 3,5–5,1", "K no teto"],
            [date(2022, 3, 3), 134, 3.8, "Na 136–145 · K 3,5–5,1", "Na abaixo da faixa (arterial)"],
        ],
    },
    {
        "grupo": "Química",
        "titulo": "Cálcio iônico",
        "mostra": "Gráfico em mmol/L (o de 2022 veio em mg/dL; convertido ×4, como no laudo de 2024)",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2022, 3, 3), 1.325, "mmol/L", "1,10 a 1,48 (convertido de 4,40–5,92 mg/dL)", "Arterial; 5,30 mg/dL no laudo"],
            [date(2024, 10, 15), 1.23, "mmol/L", "até 18 a: 1,20 a 1,38", ""],
        ],
    },
    {
        "grupo": "Química",
        "titulo": "Gasometria arterial",
        "mostra": "Internação 03/03/2022",
        "cols": ["Data", "pH", "pCO2 / HCO3", "Na / K / glicose", "Nota"],
        "linhas": [
            [date(2022, 3, 3), 7.493, "30,9 / 25,1", "134 / 3,8 / 97", "Alcalemia respiratória leve; Na baixo"],
        ],
    },
    {
        "grupo": "Química",
        "titulo": "Cortisol (matutino / série)",
        "mostra": "Dosagens soltas; a série do teste GH está na tabela do GH",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2021, 4, 20), 22.8, "mcg/dL", "7–9h: 5,3 a 22,5", "No teto / um pouco acima"],
            [date(2021, 12, 7), 26.3, "mcg/dL", "7–9h: 5,3 a 22,5", "Acima da faixa matutina"],
            [date(2024, 10, 15), 9.5, "mcg/dL", "7–9h: 5,3 a 22,5", "Dentro da faixa"],
        ],
    },
    {
        "grupo": "Química",
        "titulo": "Creatinina",
        "mostra": "Rim",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2021, 4, 20), 0.16, "mg/dL", "0–1 a: 0,17 a 0,52", "No limite inferior"],
            [date(2021, 12, 7), 0.18, "mg/dL", "2–4 a: 0,18 a 0,49", "No piso"],
            [date(2024, 10, 15), 0.22, "mg/dL", "2–4 a: 0,18 a 0,49", ""],
            [date(2026, 8, 12), 0.28, "mg/dL", "♀ 5–8 a: 0,30 a 0,61", "Abaixo da faixa"],
        ],
    },
    {
        "grupo": "Química",
        "titulo": "Ureia",
        "mostra": "Rim",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2021, 4, 20), 14, "mg/dL", "crianças: 10,8 a 38,4", ""],
            [date(2026, 8, 12), 27.8, "mg/dL", "crianças: 10,8 a 38,4", ""],
        ],
    },
    {
        "grupo": "Química",
        "titulo": "Fósforo",
        "mostra": "Osso / rim",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2021, 12, 7), 5.1, "mg/dL", "1–4 a ♀: 4,4 a 6,2", ""],
            [date(2024, 10, 15), 5.2, "mg/dL", "1–4 a ♀: 4,4 a 6,2", ""],
        ],
    },
    {
        "grupo": "Química",
        "titulo": "IgE total",
        "mostra": "Alergia / atopia",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2024, 10, 15), 104.7, "UI/mL", "1–4 a: até 313,5", ""],
            [date(2026, 2, 3), 120.3, "UI/mL", "5–10 a: até 555,1", ""],
        ],
    },
    {
        "grupo": "Química",
        "titulo": "Fosfatase alcalina",
        "mostra": "Osso / fígado",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2021, 12, 18), 514, "U/L", "♀ 1–3 a: 129 a 376", "Acima da faixa"],
            [date(2024, 10, 15), 228, "U/L", "♀ 4–6 a: 114 a 353", ""],
        ],
    },
    {
        "grupo": "Química",
        "titulo": "Gama GT",
        "mostra": "Gama-GT · fígado / via biliar",
        "cols": ["Data", "Resultado", "Unidade", "Faixa do laudo", "Nota"],
        "linhas": [
            [date(2021, 12, 18), 13.4, "U/L", "♀ < 38", ""],
            [date(2024, 10, 15), 15, "U/L", "♀ < 38", ""],
        ],
    },
    {
        "grupo": "Química",
        "titulo": "Transaminase oxalacética e pirúvica",
        "mostra": "TGO/AST e TGP/ALT — nomes do laudo Hermes Pardini",
        "cols": ["Data", "Oxalacética (U/L)", "Pirúvica (U/L)", "Faixa", "Nota"],
        "linhas": [
            [date(2021, 4, 20), 40.0, 18.0, "TGO ≤ 40 · TGP ≤ 41 (♀)", "TGO no teto"],
            [date(2021, 12, 18), 32.0, 17.0, "TGO ≤ 40 · TGP ≤ 41 (♀)", ""],
            [date(2024, 10, 15), 22, 15, "TGO ≤ 40 · TGP ≤ 41 (♀)", ""],
            [date(2025, 7, 28), 29, 20, "TGO ≤ 40 · TGP ≤ 41 (♀)", ""],
            [date(2026, 8, 12), 20, 15, "TGO ≤ 33 · TGP ≤ 41 (♀)", "Faixa de TGO deste laudo: mulheres ≤ 33"],
        ],
    },
    {
        "grupo": "Química",
        "titulo": "Dosagens pontuais",
        "mostra": "Exames feitos uma vez (ou sem série)",
        "cols": ["Data", "Exame", "Resultado", "Faixa do laudo", "Nota"],
        "largo": True,
        "linhas": [
            [date(2020, 10, 17), "Reticulócitos", "1,0% · 47.300/mm³", "0,5–1,5% ou 24.000–84.000", ""],
            [date(2021, 4, 20), "ACTH", "6,0 pg/mL", "< 46,0", ""],
            [date(2021, 4, 20), "T4 total", "9,7 mcg/dL", "4,5 a 12,3", ""],
            [date(2021, 4, 20), "T3 total", "1,39 ng/mL", "0,60 a 1,71", ""],
            [date(2021, 4, 20), "T3 livre", "4,1 pg/mL", "2,30 a 4,20", "No teto"],
            [date(2022, 3, 3), "Coagulograma", "ver laudo", "internação", ""],
            [date(2024, 10, 15), "Colesterol total e frações", "131 / HDL 44 / LDL 74 mg/dL", "cri. CT <170 · HDL >45 · LDL <110", "HDL um pouco abaixo"],
            [date(2024, 10, 15), "Triglicérides", "50 mg/dL", "0–9 a jejum: < 75", ""],
            [date(2024, 10, 15), "Imunoglobulinas IgA", "154 mg/dL", "4–6 a: 27 a 195", ""],
            [date(2024, 10, 15), "Imunoglobulinas IgG", "980 mg/dL", "ver faixa etária no laudo", ""],
            [date(2024, 10, 15), "Imunoglobulinas IgM", "74 mg/dL", "ver faixa etária no laudo", ""],
            [date(2024, 10, 15), "TTG, anticorpos anti-transglutaminase tecidual-IgA", "0,8 U/mL", "não reagente < 7,0", ""],
            [date(2024, 10, 15), "Paratormônio PTH intacto", "43,9 pg/mL", "18,5 a 88,0", ""],
            [date(2024, 10, 15), "Albumina", "4,5 g/dL", "adultos no laudo: 3,7 a 5,2", ""],
            [date(2024, 10, 15), "Proteínas totais e fracionadas", "6,9 g/dL", "1–18 a: 5,7 a 8,0", "globulinas 2,5 · A/G 1,8"],
            [date(2025, 5, 28), "Gasometria venosa", "pH 7,41 · HCO3 24 · BE −0,9", "pH 7,32–7,43 · HCO3 22–29", ""],
            [date(2025, 5, 28), "Homocisteína", "9,64 µmol/L", "♀ 5,75 a 18,89", ""],
            [date(2025, 5, 28), "Amônia", "23 µmol/L", "11 a 32", ""],
            [date(2025, 5, 28), "Ácido lático - lactato", "10,8 mg/dL", "4,5 a 19,8", ""],
            [date(2025, 5, 28), "Perfil de aminoácidos - quantitativo", "12 aa + ASA", "ver faixas no PDF", ""],
            [date(2025, 7, 28), "Hormônio folículo estimulante - FSH", "0,91 mUI/mL", "♀ 4–9 a: até 4,78", ""],
            [date(2025, 7, 28), "Hormônio luteinizante - LH", "0,07 mUI/mL", "pré-púbere ≤ 0,30", "Limite inferior do ensaio"],
            [date(2025, 7, 28), "Estradiol, 17 beta", "19,0 pg/mL", "limite inferior do ensaio", "Pré-púbere"],
            [date(2026, 8, 12), "Transferrina", "271 mg/dL", "200 a 360", ""],
            [date(2026, 8, 12), "IgE específico para clara de ovo (f1)", "0,19 kU/L", "classe 0: < 0,35", "Muito baixo / classe 0"],
        ],
    },
    {
        "grupo": "Genética",
        "titulo": "Cariótipo, array-CGH e exoma",
        "mostra": "Um exame por linha (não se repetem)",
        "arquivo": "Sangue 2022-05-25 · 2022-10-04 · 2022-12-28 · 2025-04-21 (reanálise)",
        "cols": ["Data", "Exame", "Resultado", "", "Nota"],
        "linhas": [
            [date(2022, 5, 25), "Cariótipo com banda G", "46,XX", "", "Sem alteração na resolução analisada"],
            [date(2022, 10, 4), "Análise cromossômica por array-CGH 400K (CGH+SNP)", "arr(1-22,X)x2", "", "Sem CNV clínica"],
            [date(2022, 12, 28), "Exoma com análise de CNV e DNA mitocondrial", "Sem variante reportada", "", "Reanálise 2025: mesma conclusão"],
            [date(2025, 4, 21), "Reanálise de dados brutos de exoma", "Mesma conclusão", "", ""],
        ],
    },
    {
        "grupo": "Pezinho",
        "titulo": "Teste do pezinho",
        "mostra": "Dois cartões no período neonatal",
        "cols": ["Data", "Laboratório", "Resumo", "", "Nota"],
        "largo": True,
        "linhas": [
            [date(2020, 7, 3), "NeoCenter (ampliado)", "TSH 1,2 (normal) · T4 18,0 (acima 2,3–11) · demais sem alteração", "", "T4 alto neste cartão"],
            [date(2020, 7, 21), "Programa público", "Hemoglobinas normais · IRT, Phe, 17-OHP, TSH e biotinidase no alvo", "", "Repetição dentro da referência"],
        ],
    },
    {
        "grupo": "Imagens e outros",
        "titulo": "Imagens da cabeça (período neonatal)",
        "mostra": "Ago/2020 · internada (convulsão / hemorragia)",
        "cols": ["Data", "Exame", "Resumo do laudo", "", "Nota"],
        "largo": True,
        "linhas": [
            [date(2020, 8, 4), "US transfontanela", "HIV grau II/III bilateral; ventrículos assimétricos com conteúdo ecogênico (coágulo, + à dir.)", "", "Dr. Antonio Carlos"],
            [date(2020, 8, 4), "TC crânio", "Área hiperdensa no caudado direito + sangue no sistema ventricular", "", "Indicação: convulsão"],
            [date(2020, 8, 8), "US transfontanela", "HIV grau III bilateral; ventriculomegalia; sem regressão em relação a 04/08", "", "Dra. Marlice"],
            [date(2020, 8, 11), "TC + angio", "Hemoventrículo acentuado, dilatação, cânula de derivação", "", ""],
            [date(2020, 8, 16), "RM encéfalo", "Hemorragia tálamo/caudado dir. com extensão ventricular e dilatação", "", ""],
        ],
    },
    {
        "grupo": "Imagens e outros",
        "titulo": "TC crânio (acompanhamento)",
        "mostra": "Controle da DVP · jan/2025",
        "arquivo": "Imagem - 2025-01-19 - TC Crânio",
        "cols": ["Data", "Exame", "Resumo do laudo", "", "Nota"],
        "largo": True,
        "linhas": [
            [date(2025, 1, 19), "TC crânio sem contraste", "Trepano parietal dir. (cateter de DVP; extremidade distal aparentemente fora do ventrículo); material no seio esfenoidal dir.; tecido em rinofaringe reduzindo via aérea. Sem hemorragia ou isquemia aguda.", "", "Solicitante: Dra. Silvania"],
        ],
    },
    {
        "grupo": "Imagens e outros",
        "titulo": "RM da hipófise",
        "mostra": "Mesmo exame do portal: RM da hipófise",
        "cols": ["Data", "Exame", "Resumo do laudo", "", "Nota"],
        "largo": True,
        "linhas": [
            [date(2025, 2, 4), "RM da hipófise", "Discreta hipoplasia da adeno-hipófise (altura até 3,5 mm). DVP tópica, sem dilatação.", "", "Resto sem alteração no método"],
        ],
    },
    {
        "grupo": "Imagens e outros",
        "titulo": "Grupo sanguíneo",
        "mostra": "Grupo sanguíneo + fator Rh/Du",
        "arquivo": "Sangue - 2024-10-15 - Grupo RH",
        "cols": ["Data", "Resultado", "", "", "Nota"],
        "linhas": [
            [date(2024, 10, 15), "O · Rh (D) positivo", "", "", ""],
        ],
    },
    {
        "grupo": "Imagens e outros",
        "titulo": "Suor e imagens em foto",
        "mostra": "Número não extraído (PDF escaneado)",
        "arquivo": "Suor - 2024-10-18 - Guilherme Rache · Imagem - 2025-06-04",
        "cols": ["Data", "Exame", "Resumo", "", "Nota"],
        "largo": True,
        "linhas": [
            [date(2024, 10, 18), "Teste do suor", "Laudo em foto (número não extraído)", "", "Dr. Guilherme Rache"],
            [date(2025, 6, 4), "Radiografia digital do crânio (PA e perfil)", "Laudo em foto", "", "Dra. Milena · scan"],
        ],
    },
    {
        "grupo": "Imagens e outros",
        "titulo": "EEG, córnea, US e audiologia",
        "mostra": "Exames únicos (sem série numérica)",
        "cols": ["Data", "Exame", "Resumo do laudo", "", "Nota"],
        "largo": True,
        "linhas": [
            [date(2020, 7, 1), "Orelhinha (EOA)", "Triagem neonatal", "", "Aline Cornelio"],
            [date(2020, 10, 23), "EEG sono espontâneo", "Baixa voltagem, sem ritmos regionais; sono mal definido; sem assimetria nem irritativa. Discreta disfunção córtico-subcortical inespecífica.", "", "Dra. Luciana"],
            [date(2023, 10, 31), "Topografia de córnea", "Astigmatismo discretamente irregular, boa simetria; ceratometria acima do usual. Controle em 6 meses.", "", "Dra. Cristiana · Orbscan"],
            [date(2024, 11, 14), "Audiometria, imitação, EOApd e EOAT", "Quatro laudos (Amanda Dias)", "", "Ver PDFs"],
            [date(2024, 11, 25), "PEATE", "Bruna Pereira", "", "Ver PDF"],
            [date(2025, 7, 28), "US abdome inferior", "Útero e ovários pré-púberes (útero 0,6 cm³). Folículos anecoicos. Líquido livre anexial (~1,6 e 2,7 cm³).", "", "Dra. Júlia"],
        ],
    },
]


# PDF na pasta Exames, por linha (abas Dados Completo / Dados Selecionados).
# Nome da pasta = Tipo - AAAA-MM-DD; a data da coleta às vezes cai noutro PDF.
_PDF_NOME = re.compile(r"^(.+) - (\d{4}-\d{2}-\d{2}) - (.+)$")
_PDF_SUFIXO = (
    "Fator XIII", "Grupo RH", "Culturas", "Punho", "TC Crânio",
    "Audio e imitancia", "EOApd", "EOAT",
)
_TIPO_GRUPO = {
    "Crescimento": ("Sangue",),
    "Idade óssea": ("Imagem",),
    "Tireoide": ("Sangue",),
    "Hemograma": ("Sangue",),
    "Urina": ("Sangue", "Urina"),
    "Ferro e vitaminas": ("Sangue",),
    "Química": ("Sangue",),
    "Genética": ("Sangue",),
    "Pezinho": ("Pezinho",),
    "Imagens e outros": ("Imagem", "Sangue", "Suor", "EEG", "Audiologia"),
}
# Coleta numa data, laudo impresso noutro PDF.
_ARQ_FORCAR = {
    ("IGF-1 - Somatomedina C", date(2021, 12, 18)): "Sangue - 2021-12-07",
    ("IGFBP-3 - proteína ligadora-3 do IGF", date(2021, 12, 18)): "Sangue - 2021-12-07",
    ("Fosfatase alcalina", date(2021, 12, 18)): "Sangue - 2021-12-07",
    ("IGF-1 - Somatomedina C", date(2026, 8, 17)): "Sangue - 2026-08-12 - Fernanda Silva",
}


def _indexar_pdfs():
    idx = defaultdict(list)
    if not EXAMES.is_dir():
        return idx
    for p in EXAMES.glob("*.pdf"):
        stem = p.stem
        m = _PDF_NOME.match(stem)
        if not m:
            continue
        tipo, ds, _rest = m.group(1), m.group(2), m.group(3)
        idx[(tipo, date.fromisoformat(ds))].append(stem)
    return idx


_PDFS = _indexar_pdfs()


def _nome_curto(stem):
    m = _PDF_NOME.match(stem)
    if not m:
        return stem
    tipo, ds, rest = m.group(1), m.group(2), m.group(3)
    head = f"{tipo} - {ds}"
    for suf in _PDF_SUFIXO:
        if rest.endswith(suf):
            return f"{head} - {suf}"
    irmaos = _PDFS.get((tipo, date.fromisoformat(ds)), [])
    if len(irmaos) > 1:
        nome = rest.split(" - ")[0]
        bits = nome.split()
        if len(bits) >= 2:
            return f"{head} - {bits[0]} {bits[-1]}"
        if bits:
            return f"{head} - {bits[0]}"
    return head


def _juntar(stems):
    vistos = []
    for s in stems:
        c = _nome_curto(s)
        if c not in vistos:
            vistos.append(c)
    return " · ".join(vistos)


def _cands(tipos, d):
    out = []
    for tipo in tipos:
        out.extend(_PDFS.get((tipo, d), []))
    return out


def _ctx(t, linha):
    bits = [t.get("titulo", ""), t.get("arquivo", ""), t.get("arq_nota", "")]
    if len(linha) > 1 and isinstance(linha[1], str):
        bits.append(linha[1])
    return " ".join(bits).lower()


def arquivo_linha(t, linha):
    """Nome curto do PDF daquela linha (sem .pdf), para achar na pasta Exames."""
    titulo = t["titulo"]
    d = linha[0] if linha else None
    if isinstance(d, date):
        forcar = _ARQ_FORCAR.get((titulo, d))
        if forcar:
            return forcar
    elif t.get("arquivo"):
        return t["arquivo"]

    exam = linha[1] if len(linha) > 1 else ""
    exam_l = exam.lower() if isinstance(exam, str) else ""
    ctx = _ctx(t, linha)

    tipos = list(_TIPO_GRUPO.get(t["grupo"], ("Sangue",)))
    if titulo.startswith("EAS"):
        if "cultura" in exam_l:
            tipos = ["Sangue"]
        elif "eas" in exam_l:
            tipos = ["Urina", "Sangue"]
    elif "orelhinha" in exam_l or "audiometr" in exam_l or "peate" in exam_l:
        tipos = ["Audiologia"]
    elif exam_l.startswith("eeg"):
        tipos = ["EEG"]
    elif "suor" in exam_l:
        tipos = ["Suor"]
    elif any(k in exam_l for k in ("us ", "tc ", "rm ", "topografia", "rx ")):
        tipos = ["Imagem"]
    elif titulo == "Grupo sanguíneo":
        tipos = ["Sangue"]
    elif titulo == "Fator XIII":
        tipos = ["Sangue"]

    if not isinstance(d, date):
        return t.get("arquivo") or ""

    cands = _cands(tipos, d)
    if "Urina" in tipos:
        so_urina = _PDFS.get(("Urina", d), [])
        if so_urina and "eas" in exam_l:
            cands = so_urina

    if not cands:
        return t.get("arquivo") or ""

    if "fator xiii" in ctx:
        hit = [s for s in cands if "fator xiii" in s.lower()]
        if hit:
            return _juntar(hit)
    if "grupo" in ctx and ("rh" in ctx or "sanguíneo" in ctx or "sanguineo" in ctx):
        hit = [s for s in cands if "grupo rh" in s.lower()]
        if hit:
            return _juntar(hit)
    if "cultura" in ctx:
        hit = [s for s in cands if "culturas" in s.lower()]
        if hit:
            return _juntar(hit)
    if "punho" in ctx or t.get("arq_nota") == "Punho":
        hit = [s for s in cands if "punho" in s.lower()]
        if hit:
            return _juntar(hit)
    if "tc crânio" in titulo.lower() or "tc cranio" in titulo.lower() or exam_l.startswith("tc"):
        hit = [s for s in cands if "tc crânio" in s.lower() or "tc cranio" in s.lower()
               or "sem solicitante" in s.lower() or "silvania" in s.lower()]
        if hit and not exam_l.startswith("us"):
            if "angio" in exam_l:
                hit = [s for s in cands if "silvania" in s.lower()] or hit
            elif "2025" not in str(d):
                hit = [s for s in cands if "sem solicitante" in s.lower()] or hit
            return _juntar(hit)
    if "transfont" in exam_l:
        hit = [s for s in cands if "antonio carlos" in s.lower() or "marlice" in s.lower()]
        if hit:
            return _juntar(hit)
    if exam_l.startswith("rm"):
        hit = [s for s in cands if "punho" not in s.lower()]
        if hit:
            return _juntar(hit)

    extras = [s for s in cands if any(suf.lower() in s.lower() for suf in _PDF_SUFIXO)]
    base = [s for s in cands if s not in extras]
    if "audiometr" in exam_l or "eoap" in exam_l:
        return _juntar(cands)
    return _juntar(base or cands)


# Gráficos de faixa: (data, valor, piso, teto) — piso/teto do laudo naquela idade/data.
# Vários gráficos por tabela (ex.: Na e K) ficam no mesmo item da caixinha.
def n_validos(spec):
    """Quantos pontos numéricos dá para plotar. Gráfico só com 2 ou mais."""
    tipo = spec.get("tipo", "valor")
    pts = spec.get("pontos") or []
    if tipo == "idade":
        return len(pts)
    if tipo == "tempo":
        return sum(1 for p in pts if p.get("v") is not None)
    return sum(1 for p in pts if p.get("v") is not None)


def _fmt_num_br(v):
    if v is None:
        return ""
    if isinstance(v, float):
        s = f"{v:.4f}".rstrip("0").rstrip(".")
        return s.replace(".", ",")
    return str(v).replace(".", ",")


def arquivo_do_exame(t, d):
    if not isinstance(d, date):
        return t.get("arquivo") or ""
    for linha in t.get("linhas") or []:
        if linha and linha[0] == d:
            return arquivo_linha(t, linha)
    return arquivo_linha(t, [d])


def linha_unico_dosagem(t, linha):
    """Uma linha da tabela 'Dosagens pontuais' (Resultados) para o final dos gráficos."""
    return {
        "exame": linha[1] if len(linha) > 1 else t["titulo"],
        "data": linha[0],
        "valor": linha[2] if len(linha) > 2 else "",
        "unidade": "",
        "faixa": linha[3] if len(linha) > 3 else "",
        "arquivo": arquivo_linha(t, linha),
    }


def linha_unico(t, spec):
    """Uma linha para a tabela de exames com um único resultado."""
    tipo = spec.get("tipo", "valor")
    pts = spec.get("pontos") or []
    titulo_c = spec.get("titulo") or t["titulo"]
    unidade = spec.get("unidade") or ""
    if tipo == "tempo":
        ok = [p for p in pts if p.get("v") is not None]
        if len(ok) != 1:
            return None
        p = ok[0]
        return {
            "exame": titulo_c,
            "data": p.get("cat") or "",
            "valor": p["v"],
            "unidade": unidade,
            "faixa": "",
            "arquivo": t.get("arquivo") or "",
        }
    if tipo == "idade":
        if len(pts) != 1:
            return None
        p = pts[0]
        return {
            "exame": titulo_c,
            "data": p["data"],
            "valor": f"{p['crono']} / {p['ossea']}",
            "unidade": unidade,
            "faixa": "",
            "arquivo": arquivo_do_exame(t, p["data"]),
        }
    ok = [p for p in pts if p.get("v") is not None]
    if len(ok) != 1:
        return None
    p = ok[0]
    lo, hi = p.get("lo"), p.get("hi")
    faixa = ""
    if lo is not None and hi is not None:
        faixa = f"{_fmt_num_br(lo)} a {_fmt_num_br(hi)}"
    return {
        "exame": titulo_c,
        "data": p["data"],
        "valor": p["v"],
        "unidade": unidade,
        "faixa": faixa,
        "arquivo": arquivo_do_exame(t, p["data"]),
    }


def _pt(d, v, lo=None, hi=None):
    return {"data": d, "v": v, "lo": lo, "hi": hi}


GRAF = {
    "IGF-1 - Somatomedina C": [
        {"unidade": "ng/mL", "pontos": [
            _pt(date(2021, 4, 20), 15.0, 18, 172),
            _pt(date(2021, 12, 18), 18.0, 18, 172),
            _pt(date(2024, 10, 15), 39, 22, 208),
            _pt(date(2025, 7, 28), 122, 22, 208),
            _pt(date(2026, 2, 3), 68, 22, 208),
            _pt(date(2026, 8, 17), 173, 35, 232),
        ]},
    ],
    "IGFBP-3 - proteína ligadora-3 do IGF": [
        {"unidade": "mcg/mL", "pontos": [
            _pt(date(2021, 4, 20), 1.00),
            _pt(date(2021, 12, 18), 1.50),
            _pt(date(2024, 10, 15), 2.0),
            _pt(date(2025, 7, 28), 3.6),
            _pt(date(2026, 8, 12), 3.3),
        ]},
    ],
    "GH - hormônio de crescimento (basal)": [
        {"unidade": "ng/mL", "pontos": [
            _pt(date(2021, 12, 7), 1.08),
            _pt(date(2025, 1, 13), 2.17),
        ]},
    ],
    "Teste do GH com glucagon": [
        {"titulo": "GH no teste com glucagon", "unidade": "ng/mL", "tipo": "tempo",
         "nota": "Sem barra de faixa. O laudo pede GH ≥ 5 ng/mL em qualquer tempo (pico 3,07).",
         "pontos": [
            {"cat": "Basal", "v": 2.17},
            {"cat": "60 min", "v": 0.43},
            {"cat": "90 min", "v": 1.06},
            {"cat": "120 min", "v": 3.07},
            {"cat": "180 min", "v": 0.97},
        ]},
        {"titulo": "Glicose no teste GH", "unidade": "mg/dL", "tipo": "tempo",
         "nota": "Sem faixa de referência: glicose durante o teste de estímulo.",
         "pontos": [
            {"cat": "Basal", "v": None},
            {"cat": "60 min", "v": 199},
            {"cat": "90 min", "v": 112},
            {"cat": "120 min", "v": 73},
            {"cat": "180 min", "v": 70},
        ]},
        {"titulo": "Cortisol no teste GH", "unidade": "mcg/dL", "tipo": "tempo",
         "nota": "Sem faixa de referência: cortisol durante o teste de estímulo.",
         "pontos": [
            {"cat": "Basal", "v": 10.3},
            {"cat": "60 min", "v": 12.1},
            {"cat": "90 min", "v": 15.6},
            {"cat": "120 min", "v": 29.0},
            {"cat": "180 min", "v": 22.2},
        ]},
    ],
    "RX mão/punho (Greulich-Pyle, feminino)": [
        {"tipo": "idade", "unidade": "meses", "pontos": [
            {"data": date(2025, 2, 10), "crono": 4 * 12 + 7, "ossea": 2 * 12 + 6},
            {"data": date(2026, 2, 18), "crono": 5 * 12 + 7, "ossea": 3 * 12 + 6},
            {"data": date(2026, 7, 29), "crono": 6 * 12 + 1, "ossea": 4 * 12 + 2},
        ]},
    ],
    "TSH ultra sensível": [
        {"unidade": "µUI/mL", "pontos": [
            _pt(date(2021, 4, 20), 3.43, 0.70, 6.55),
            _pt(date(2021, 12, 7), 2.56, 0.70, 6.55),
            _pt(date(2024, 10, 15), 2.36, 0.70, 6.55),
            _pt(date(2025, 7, 28), 1.92, 0.70, 6.55),
            _pt(date(2026, 2, 3), 1.42, 0.70, 6.55),
            _pt(date(2026, 8, 12), 1.50, 0.70, 6.55),
        ]},
    ],
    "T4 livre": [
        {"unidade": "ng/dL", "pontos": [
            _pt(date(2021, 4, 20), 1.22, 0.85, 1.67),
            _pt(date(2021, 12, 7), 1.07, 0.85, 1.67),
            _pt(date(2024, 10, 15), 1.15, 0.85, 1.67),
            _pt(date(2025, 7, 28), 1.20, 0.85, 1.67),
            _pt(date(2026, 2, 3), 1.29, 0.85, 1.67),
            _pt(date(2026, 8, 12), 1.23, 1.05, 1.67),
        ]},
    ],
    "Hemoglobina": [
        {"unidade": "g/dL", "pontos": [
            _pt(date(2020, 10, 17), 12.3, 10.0, 14.0),
            _pt(date(2021, 4, 20), 14.7, 10.5, 13.5),
            _pt(date(2021, 12, 7), 13.9, 10.5, 14.0),
            _pt(date(2022, 2, 9), 12.7, 10.5, 13.5),
            _pt(date(2022, 3, 1), 12.5, 10.5, 13.5),
            _pt(date(2022, 3, 3), 12.3, 10.5, 13.5),
            _pt(date(2024, 10, 15), 12.2, 11.5, 13.5),
            _pt(date(2025, 5, 28), 12.7, 11.5, 13.5),
            _pt(date(2025, 7, 28), 13.2, 11.5, 13.5),
            _pt(date(2026, 2, 3), 12.7, 11.5, 13.5),
            _pt(date(2026, 7, 29), 13.0, 11.5, 13.5),
            _pt(date(2026, 8, 12), 12.9, 11.5, 13.5),
        ]},
    ],
    "Hemácias": [
        {"unidade": "10⁶/mm³", "pontos": [
            _pt(date(2020, 10, 17), 4.73, 3.20, 4.50),
            _pt(date(2021, 4, 20), 5.35, 3.70, 5.70),
            _pt(date(2021, 12, 7), 5.07, 3.70, 5.20),
            _pt(date(2022, 2, 9), 4.52, 3.80, 5.00),
            _pt(date(2022, 3, 1), 4.50, 3.80, 5.00),
            _pt(date(2022, 3, 3), 4.48, 3.80, 5.00),
            _pt(date(2024, 10, 15), 4.95, 3.90, 5.30),
            _pt(date(2025, 5, 28), 4.90, 3.90, 5.30),
            _pt(date(2025, 7, 28), 4.95, 3.90, 5.30),
            _pt(date(2026, 2, 3), 4.70, 3.90, 5.30),
            _pt(date(2026, 7, 29), 4.78, 3.90, 5.30),
            _pt(date(2026, 8, 12), 4.86, 3.90, 5.30),
        ]},
    ],
    "Hematócrito": [
        {"unidade": "%", "pontos": [
            _pt(date(2020, 10, 17), 37.5, 28, 42),
            _pt(date(2021, 4, 20), 44.2, 32, 40),
            _pt(date(2021, 12, 7), 41.0, 33, 42),
            _pt(date(2022, 2, 9), 36.7, 24, 44),
            _pt(date(2022, 3, 1), 35.9, 24, 44),
            _pt(date(2022, 3, 3), 35.9, 24, 44),
            _pt(date(2024, 10, 15), 40.7, 34, 40),
            _pt(date(2025, 5, 28), 38.8, 34, 40),
            _pt(date(2025, 7, 28), 40.5, 34, 40),
            _pt(date(2026, 2, 3), 38.0, 34, 40),
            _pt(date(2026, 7, 29), 38.4, 34, 40),
            _pt(date(2026, 8, 12), 38.5, 34, 40),
        ]},
    ],
    "VCM": [
        {"unidade": "fL", "pontos": [
            _pt(date(2020, 10, 17), 79.3, 78, 100),
            _pt(date(2021, 4, 20), 82.6, 74, 88),
            _pt(date(2021, 12, 7), 80.9, 75, 88),
            _pt(date(2022, 2, 9), 81, 78, 92),
            _pt(date(2022, 3, 1), 80, 78, 92),
            _pt(date(2022, 3, 3), 80, 78, 92),
            _pt(date(2024, 10, 15), 82.2, 75, 87),
            _pt(date(2025, 5, 28), 79.2, 75, 87),
            _pt(date(2025, 7, 28), 81.8, 75, 87),
            _pt(date(2026, 2, 3), 80.9, 75, 87),
            _pt(date(2026, 7, 29), 80.3, 75, 87),
            _pt(date(2026, 8, 12), 79.2, 75, 87),
        ]},
    ],
    "HCM e CHCM": [
        {"titulo": "HCM", "unidade": "pg", "pontos": [
            _pt(date(2020, 10, 17), 26.0, 26, 34),
            _pt(date(2021, 4, 20), 27.5, 25, 34),
            _pt(date(2021, 12, 7), 27.4, 26, 33),
            _pt(date(2022, 2, 9), 28, 26, 31),
            _pt(date(2022, 3, 1), 28, 26, 31),
            _pt(date(2022, 3, 3), 27, 26, 31),
            _pt(date(2024, 10, 15), 24.6, 24, 30),
            _pt(date(2025, 5, 28), 25.9, 24, 30),
            _pt(date(2025, 7, 28), 26.7, 24, 30),
            _pt(date(2026, 2, 3), 27.0, 24, 30),
            _pt(date(2026, 7, 29), 27.2, 24, 30),
            _pt(date(2026, 8, 12), 26.5, 24, 30),
        ]},
        {"titulo": "CHCM", "unidade": "g/dL", "pontos": [
            _pt(date(2020, 10, 17), 32.8, 30, 36),
            _pt(date(2021, 4, 20), 33.3, 30, 36),
            _pt(date(2021, 12, 7), 33.9, 31, 36),
            _pt(date(2022, 2, 9), 35, 31, 34),
            _pt(date(2022, 3, 1), 35, 31, 34),
            _pt(date(2022, 3, 3), 34, 31, 34),
            _pt(date(2024, 10, 15), 30.0, 32, 36),
            _pt(date(2025, 5, 28), 32.7, 32, 36),
            _pt(date(2025, 7, 28), 32.6, 32, 36),
            _pt(date(2026, 2, 3), 33.4, 32, 36),
            _pt(date(2026, 7, 29), 33.9, 32, 36),
            _pt(date(2026, 8, 12), 33.5, 32, 36),
        ]},
    ],
    "RDW": [
        {"unidade": "%", "pontos": [
            _pt(date(2020, 10, 17), 13.7, 11.3, 15.3),
            _pt(date(2021, 4, 20), 13.5, 11.3, 15.3),
            _pt(date(2021, 12, 7), 13.5, 11.3, 15.3),
            _pt(date(2022, 2, 9), 12.4, 11.3, 15.0),
            _pt(date(2022, 3, 1), 12.1, 11.3, 15.0),
            _pt(date(2022, 3, 3), 12.1, 11.3, 15.0),
            _pt(date(2024, 10, 15), 14.6, 11.5, 14.8),
            _pt(date(2025, 5, 28), 14.1, 11.5, 14.8),
            _pt(date(2025, 7, 28), 13.1, 11.5, 14.8),
            _pt(date(2026, 2, 3), 13.1, 11.5, 14.8),
            _pt(date(2026, 7, 29), 13.2, 11.5, 14.8),
            _pt(date(2026, 8, 12), 12.4, 11.5, 14.8),
        ]},
    ],
    "Leucócitos": [
        {"unidade": "/mm³", "pontos": [
            _pt(date(2020, 10, 17), 6350, 5000, 14500),
            _pt(date(2021, 4, 20), 9590, 6000, 11000),
            _pt(date(2021, 12, 7), 7420, 6000, 11000),
            _pt(date(2022, 2, 9), 9400, 6000, 17500),
            _pt(date(2022, 3, 1), 7100, 6000, 17500),
            _pt(date(2022, 3, 3), 8200, 6000, 17500),
            _pt(date(2024, 10, 15), 9930, 5000, 14500),
            _pt(date(2025, 5, 28), 8950, 5000, 14500),
            _pt(date(2025, 7, 28), 13580, 5000, 14500),
            _pt(date(2026, 2, 3), 8810, 5000, 14500),
            _pt(date(2026, 7, 29), 4310, 5000, 14500),
            _pt(date(2026, 8, 12), 7560, 5000, 14500),
        ]},
    ],
    "Neutrófilos": [
        {"unidade": "/mm³", "pontos": [
            _pt(date(2020, 10, 17), 1270, 1000, 8500),
            _pt(date(2021, 4, 20), 1055, 1500, 8500),
            _pt(date(2021, 12, 7), 1484, 1500, 8500),
            _pt(date(2022, 2, 9), 4774, 2000, 5000),
            _pt(date(2022, 3, 1), 2785, 2000, 5000),
            _pt(date(2022, 3, 3), 2545, 2000, 5000),
            _pt(date(2024, 10, 15), 4600, 1500, 8000),
            _pt(date(2025, 5, 28), 3230, 1500, 8000),
            _pt(date(2025, 7, 28), 6720, 1500, 8000),
            _pt(date(2026, 2, 3), 3990, 1500, 8000),
            _pt(date(2026, 7, 29), 2240, 1500, 8000),
            _pt(date(2026, 8, 12), 2440, 1500, 8000),
        ]},
    ],
    "Linfócitos": [
        {"unidade": "/mm³", "pontos": [
            _pt(date(2020, 10, 17), 4572, 4000, 13500),
            _pt(date(2021, 4, 20), 7672, 4000, 10500),
            _pt(date(2021, 12, 7), 5194, 1500, 7000),
            _pt(date(2022, 2, 9), 3557, 5500, 9000),
            _pt(date(2022, 3, 1), 3856, 5500, 9000),
            _pt(date(2022, 3, 3), 4680, 5500, 9000),
            _pt(date(2024, 10, 15), 4260, 1500, 7000),
            _pt(date(2025, 5, 28), 4700, 1500, 7000),
            _pt(date(2025, 7, 28), 5270, 1500, 7000),
            _pt(date(2026, 2, 3), 3690, 1500, 7000),
            _pt(date(2026, 7, 29), 1030, 1500, 7000),
            _pt(date(2026, 8, 12), 4180, 1500, 7000),
        ]},
    ],
    "Monócitos": [
        {"unidade": "/mm³", "pontos": [
            _pt(date(2020, 10, 17), 444, 80, 1100),
            _pt(date(2021, 4, 20), 671, 80, 1000),
            _pt(date(2021, 12, 7), 668, 80, 800),
            _pt(date(2022, 2, 9), 936, 200, 1300),
            _pt(date(2022, 3, 1), 428, 200, 1300),
            _pt(date(2022, 3, 3), 903, 200, 1300),
            _pt(date(2024, 10, 15), 940, 200, 1000),
            _pt(date(2025, 5, 28), 820, 200, 1000),
            _pt(date(2025, 7, 28), 1420, 200, 1000),
            _pt(date(2026, 2, 3), 910, 200, 1000),
            _pt(date(2026, 7, 29), 1030, 200, 1000),
            _pt(date(2026, 8, 12), 800, 200, 1000),
        ]},
    ],
    "Eosinófilos": [
        {"unidade": "/mm³", "pontos": [
            _pt(date(2020, 10, 17), 64, 40, 850),
            _pt(date(2021, 4, 20), 192, 40, 700),
            _pt(date(2021, 12, 7), 74, 40, 650),
            _pt(date(2022, 2, 9), 94, 70, 650),
            _pt(date(2022, 3, 1), 71, 70, 650),
            _pt(date(2024, 10, 15), 100, 100, 1000),
            _pt(date(2025, 5, 28), 140, 100, 1000),
            _pt(date(2025, 7, 28), 100, 100, 1000),
            _pt(date(2026, 2, 3), 190, 100, 1000),
            _pt(date(2026, 7, 29), 0, 100, 1000),
            _pt(date(2026, 8, 12), 100, 100, 1000),
        ]},
    ],
    "Basófilos": [
        {"unidade": "/mm³", "pontos": [
            _pt(date(2024, 10, 15), 30, 0, 200),
            _pt(date(2025, 5, 28), 60, 0, 200),
            _pt(date(2025, 7, 28), 70, 0, 200),
            _pt(date(2026, 2, 3), 30, 0, 200),
            _pt(date(2026, 7, 29), 10, 0, 200),
            _pt(date(2026, 8, 12), 40, 0, 200),
        ]},
    ],
    "Plaquetas": [
        {"unidade": "/mm³", "pontos": [
            _pt(date(2020, 10, 17), 180000, 150000, 450000),
            _pt(date(2021, 4, 20), 185000, 150000, 450000),
            _pt(date(2021, 12, 7), 231000, 150000, 450000),
            _pt(date(2022, 2, 9), 303000, 200000, 500000),
            _pt(date(2022, 3, 1), 350000, 200000, 500000),
            _pt(date(2022, 3, 3), 367000, 200000, 500000),
            _pt(date(2024, 10, 15), 379000, 150000, 450000),
            _pt(date(2025, 5, 28), 365000, 150000, 450000),
            _pt(date(2025, 7, 28), 401000, 150000, 450000),
            _pt(date(2026, 2, 3), 309000, 150000, 450000),
            _pt(date(2026, 7, 29), 202000, 150000, 450000),
            _pt(date(2026, 8, 12), 458000, 150000, 450000),
        ]},
    ],
    "Fator XIII": [
        {"unidade": "%", "pontos": [
            _pt(date(2020, 10, 17), 28.0, 70, 140),
        ]},
    ],
    "PCR - proteína C reativa quantitativa": [
        {"unidade": "mg/L", "pontos": [
            _pt(date(2022, 2, 9), 1.9, 0, 10),
            _pt(date(2022, 3, 1), 1.0, 0, 10),
            _pt(date(2026, 7, 29), 4.0, 0, 10),
        ]},
    ],
    "Ferritina sérica": [
        {"unidade": "ng/mL", "pontos": [
            _pt(date(2024, 10, 15), 11.6, 7, 140),
            _pt(date(2025, 7, 28), 6.9, 7, 140),
            _pt(date(2026, 2, 3), 19.5, 7, 140),
            _pt(date(2026, 8, 12), 44.1, 7, 140),
        ]},
    ],
    "Ferro sérico": [
        {"unidade": "mcg/dL", "pontos": [
            _pt(date(2025, 7, 28), 44, 50, 120),
            _pt(date(2026, 2, 3), 83, 50, 120),
            _pt(date(2026, 8, 12), 123, 50, 120),
        ]},
    ],
    "Índice de saturação da transferrina": [
        {"unidade": "%", "pontos": [
            _pt(date(2025, 7, 28), 11, 15, 50),
            _pt(date(2026, 8, 12), 37, 15, 50),
        ]},
    ],
    "Capacidade total de combinação do ferro": [
        {"unidade": "mcg/dL", "pontos": [
            _pt(date(2026, 2, 3), 332, 250, 425),
            _pt(date(2026, 8, 12), 333, 250, 425),
        ]},
    ],
    "25-Hidroxivitamina D": [
        {"unidade": "ng/mL", "pontos": [
            _pt(date(2021, 12, 7), 29.9, 20, 60),
            _pt(date(2024, 10, 15), 28.9, 20, 60),
            _pt(date(2026, 8, 12), 28.2, 20, 60),
        ]},
    ],
    "Vitamina B12": [
        {"unidade": "pg/mL", "pontos": [
            _pt(date(2025, 5, 28), 672, 172, 890),
        ]},
    ],
    "Glicose - jejum": [
        {"unidade": "mg/dL", "pontos": [
            _pt(date(2021, 4, 20), 94, 60, 99),
            _pt(date(2021, 12, 7), 60, 60, 99),
            _pt(date(2022, 3, 3), 97, 60, 99),
            _pt(date(2024, 10, 15), 87, 60, 99),
            _pt(date(2025, 7, 28), 77, 60, 99),
            _pt(date(2026, 2, 3), 82, 60, 99),
            _pt(date(2026, 8, 12), 80, 60, 99),
        ]},
    ],
    "Sódio e potássio": [
        {"titulo": "Sódio", "unidade": "mEq/L", "pontos": [
            _pt(date(2020, 10, 17), 135, 136, 145),
            _pt(date(2021, 4, 20), 136, 136, 145),
            _pt(date(2021, 12, 7), 138, 136, 145),
            _pt(date(2022, 3, 3), 134, 136, 145),
        ]},
        {"titulo": "Potássio", "unidade": "mEq/L", "pontos": [
            _pt(date(2020, 10, 17), 5.1, 3.5, 5.1),
            _pt(date(2021, 4, 20), 4.9, 3.5, 5.1),
            _pt(date(2021, 12, 7), 5.1, 3.5, 5.1),
            _pt(date(2022, 3, 3), 3.8, 3.5, 5.1),
        ]},
    ],
    "Cálcio iônico": [
        {"unidade": "mmol/L", "pontos": [
            _pt(date(2022, 3, 3), 1.325, 1.10, 1.48),
            _pt(date(2024, 10, 15), 1.23, 1.20, 1.38),
        ]},
    ],
    "Gasometria arterial": [
        {"titulo": "pH arterial", "unidade": "pH", "pontos": [
            _pt(date(2022, 3, 3), 7.493, 7.35, 7.45),
        ]},
    ],
    "Cortisol (matutino / série)": [
        {"unidade": "mcg/dL", "pontos": [
            _pt(date(2021, 4, 20), 22.8, 5.3, 22.5),
            _pt(date(2021, 12, 7), 26.3, 5.3, 22.5),
            _pt(date(2024, 10, 15), 9.5, 5.3, 22.5),
        ]},
    ],
    "Creatinina": [
        {"unidade": "mg/dL", "pontos": [
            _pt(date(2021, 4, 20), 0.16, 0.17, 0.52),
            _pt(date(2021, 12, 7), 0.18, 0.18, 0.49),
            _pt(date(2024, 10, 15), 0.22, 0.18, 0.49),
            _pt(date(2026, 8, 12), 0.28, 0.30, 0.61),
        ]},
    ],
    "Ureia": [
        {"unidade": "mg/dL", "pontos": [
            _pt(date(2021, 4, 20), 14, 10.8, 38.4),
            _pt(date(2026, 8, 12), 27.8, 10.8, 38.4),
        ]},
    ],
    "Fósforo": [
        {"unidade": "mg/dL", "pontos": [
            _pt(date(2021, 12, 7), 5.1, 4.4, 6.2),
            _pt(date(2024, 10, 15), 5.2, 4.4, 6.2),
        ]},
    ],
    "IgE total": [
        {"unidade": "UI/mL", "pontos": [
            _pt(date(2024, 10, 15), 104.7, 0, 313.5),
            _pt(date(2026, 2, 3), 120.3, 0, 555.1),
        ]},
    ],
    "Fosfatase alcalina": [
        {"unidade": "U/L", "pontos": [
            _pt(date(2021, 12, 18), 514, 129, 376),
            _pt(date(2024, 10, 15), 228, 114, 353),
        ]},
    ],
    "Gama GT": [
        {"unidade": "U/L", "pontos": [
            _pt(date(2021, 12, 18), 13.4, 0, 38),
            _pt(date(2024, 10, 15), 15, 0, 38),
        ]},
    ],
    "Transaminase oxalacética e pirúvica": [
        {"titulo": "Transaminase oxalacética", "unidade": "U/L", "pontos": [
            _pt(date(2021, 4, 20), 40.0, 0, 40),
            _pt(date(2021, 12, 18), 32.0, 0, 40),
            _pt(date(2024, 10, 15), 22, 0, 40),
            _pt(date(2025, 7, 28), 29, 0, 40),
            _pt(date(2026, 8, 12), 20, 0, 33),
        ]},
        {"titulo": "Transaminase pirúvica", "unidade": "U/L", "pontos": [
            _pt(date(2021, 4, 20), 18.0, 0, 41),
            _pt(date(2021, 12, 18), 17.0, 0, 41),
            _pt(date(2024, 10, 15), 15, 0, 41),
            _pt(date(2025, 7, 28), 20, 0, 41),
            _pt(date(2026, 8, 12), 15, 0, 41),
        ]},
    ],
}


def fmt_value(v):
    if v is None or v == "":
        return None
    if isinstance(v, date):
        return v
    return v


def build(path: Path):
    faltou = [
        (t["titulo"], linha[0])
        for t in TABLES
        for linha in t["linhas"]
        if not arquivo_linha(t, linha)
    ]
    if faltou:
        print("ARQ_FALTOU", faltou)

    wb = xlsxwriter.Workbook(str(path), {"default_date_format": "DD/MM/YYYY"})

    navy = "#1B4F72"
    ink = "#1C2833"
    muted = "#5D6D7E"
    cream = "#FBF6EE"
    green = "#E8F6EF"
    green_dk = "#196F3D"
    white = "#FFFFFF"
    band = "#F4F6F7"
    line = "#D5D8DC"

    f_title = wb.add_format({
        "bold": True, "font_size": 16, "font_color": navy, "font_name": "Calibri",
        "valign": "vcenter",
    })
    f_sub = wb.add_format({
        "font_size": 10, "font_color": muted, "font_name": "Calibri", "text_wrap": True,
    })
    f_step = wb.add_format({
        "bold": True, "font_size": 11, "font_color": navy, "bg_color": cream,
        "font_name": "Calibri", "valign": "vcenter",
    })
    f_box = wb.add_format({
        "font_size": 16, "align": "center", "valign": "vcenter", "font_name": "Calibri",
        "bold": True,
    })
    f_print = wb.add_format({
        "font_size": 12, "align": "center", "valign": "vcenter", "font_name": "Calibri",
        "bold": True, "bg_color": green, "font_color": green_dk, "border": 1,
        "border_color": "#A9DFBF",
    })
    f_grp = wb.add_format({
        "font_size": 10, "font_name": "Calibri", "valign": "vcenter", "font_color": muted,
    })
    f_tab = wb.add_format({
        "font_size": 10, "font_name": "Calibri", "valign": "vcenter", "bold": True,
        "font_color": ink, "text_wrap": True,
    })
    f_desc = wb.add_format({
        "font_size": 8, "font_name": "Calibri", "valign": "vcenter", "font_color": muted,
        "text_wrap": True,
    })
    f_arq = wb.add_format({
        "font_size": 8, "font_name": "Calibri", "valign": "vcenter", "font_color": navy,
        "text_wrap": True, "border": 1, "border_color": line,
    })
    f_arq_b = wb.add_format({
        "font_size": 8, "font_name": "Calibri", "valign": "vcenter", "font_color": navy,
        "text_wrap": True, "border": 1, "border_color": line, "bg_color": band,
    })
    f_th = wb.add_format({
        "bold": True, "font_size": 9, "font_color": white, "bg_color": navy,
        "font_name": "Calibri", "align": "left", "valign": "vcenter",
    })
    f_h2 = wb.add_format({
        "bold": True, "font_size": 12, "font_color": navy, "font_name": "Calibri",
        "bg_color": "#D4E6F1", "valign": "vcenter",
    })
    f_h3 = wb.add_format({
        "italic": True, "font_size": 8, "font_color": muted, "font_name": "Calibri",
    })
    f_cab_m = wb.add_format({
        "bold": True, "font_size": 8, "font_color": white, "bg_color": navy,
        "font_name": "Calibri", "valign": "vcenter", "text_wrap": True, "align": "left",
    })
    f_cab = f_cab_m
    f_long = wb.add_format({
        "font_size": 9, "font_name": "Calibri", "valign": "top", "text_wrap": True,
        "border": 1, "border_color": line, "font_color": ink, "align": "left",
    })
    f_long_b = wb.add_format({
        "font_size": 9, "font_name": "Calibri", "valign": "top", "text_wrap": True,
        "border": 1, "border_color": line, "bg_color": band, "font_color": ink, "align": "left",
    })
    f_cell = wb.add_format({
        "font_size": 9, "font_name": "Calibri", "valign": "vcenter", "text_wrap": True,
        "border": 1, "border_color": line, "font_color": ink,
    })
    f_cell_b = wb.add_format({
        "font_size": 9, "font_name": "Calibri", "valign": "vcenter", "text_wrap": True,
        "border": 1, "border_color": line, "bg_color": band, "font_color": ink,
    })
    f_num = wb.add_format({
        "font_size": 9, "font_name": "Calibri", "valign": "vcenter", "align": "center",
        "border": 1, "border_color": line, "num_format": "#,##0.##",
    })
    f_num_b = wb.add_format({
        "font_size": 9, "font_name": "Calibri", "valign": "vcenter", "align": "center",
        "border": 1, "border_color": line, "bg_color": band, "num_format": "#,##0.##",
    })
    f_date = wb.add_format({
        "font_size": 9, "font_name": "Calibri", "valign": "vcenter", "align": "center",
        "border": 1, "border_color": line, "num_format": "DD/MM/YYYY",
    })
    f_date_b = wb.add_format({
        "font_size": 9, "font_name": "Calibri", "valign": "vcenter", "align": "center",
        "border": 1, "border_color": line, "bg_color": band, "num_format": "DD/MM/YYYY",
    })
    f_empty = wb.add_format({"font_size": 2})
    f_foot = wb.add_format({
        "italic": True, "font_size": 8, "font_color": muted, "font_name": "Calibri",
        "text_wrap": True,
    })
    f_help = wb.add_format({
        "font_size": 11, "font_name": "Calibri", "text_wrap": True, "valign": "top",
    })

    # --------- Escolher ----------
    ws = wb.add_worksheet(ABA_ESCOLHER)
    ws.set_tab_color(navy)
    ws.hide_gridlines(2)
    ws.set_column("A:A", 4)
    ws.set_column("B:B", 3, None, {"hidden": True})
    ws.set_column("C:C", 16)
    ws.set_column("D:D", 32)
    ws.set_column("E:E", 42)
    ws.set_column("F:F", 26)
    ws.set_column("G:G", 20)
    ws.set_column("H:H", 4, None, {"hidden": True})
    ws.set_row(0, 22)
    ws.set_row(1, 32)
    ws.set_row(2, 48)
    ws.set_row(3, 20)

    n = len(TABLES)
    last_row = 4 + n  # header row 4 (1-based 5), data 5..4+n
    ws.merge_range(0, 0, 0, 4, "Cecília Maria Albergaria Silva  ·  Evolução dos exames", f_title)
    ws.merge_range(
        1, 0, 1, 4,
        "Nascimento 30/06/2020.  Passo 1: clique na caixinha à esquerda.  "
        "Passo 2: abra Dados Selecionados (tabelas) ou Graficos Selecionados.  "
        "Para PDF em retrato: 3 ou 4 por folha, Imprimir selecionados ou Imprimir todos.  Dados Completo e Graficos Completo sempre mostram tudo.",
        f_sub,
    )
    ws.merge_range(3, 0, 3, 4, "Clique na caixinha. PDF: 3 ou 4 por folha e os botões à direita.", f_step)
    ws.write(2, 7, 1)  # H3: 1 = 3 por folha, 2 = 4 por folha

    headers = [" ", "_ligado", "Grupo", "Tabela", "O que aparece"]
    for c, h in enumerate(headers):
        ws.write(4, c, h, f_th)

    for i, t in enumerate(TABLES):
        r = 5 + i
        ws.set_row(r, 22)
        ws.write(r, 0, None)
        ws.write_boolean(r, 1, False)
        ws.write(r, 2, t["grupo"], f_grp)
        ws.write(r, 3, t["titulo"], f_tab)
        ws.write(r, 4, t["mostra"], f_desc)

    ws.freeze_panes(5, 0)
    ws.write(
        5 + n + 1,
        2,
        "Dica: clique no quadradinho, não no texto. Habilite macros se o Excel pedir. "
        "Dados Completo / Graficos Completo = tudo. Dados Selecionados / Graficos Selecionados = só o marcado. "
        "PDF: escolha 3 ou 4 por folha e clique Imprimir selecionados ou Imprimir todos. "
        "O PDF de cada valor está na coluna Arquivo das abas Dados Completo e Dados Selecionados.",
        f_sub,
    )
    ws.set_landscape()
    ws.set_paper(9)
    ws.set_margins(0.4, 0.4, 0.4, 0.4)
    ws.fit_to_pages(1, 0)
    ws.print_area(0, 0, last_row + 2, 4)
    ws.repeat_rows(0, 4)


    f_title_exam = wb.add_format({
        "bold": True, "font_size": 11, "font_color": navy, "font_name": "Calibri",
        "valign": "vcenter", "bg_color": "#D6EAF8", "border": 0,
    })

    def setup_print_sheet(sh):
        sh.set_landscape()
        sh.set_paper(9)
        sh.set_margins(0.4, 0.4, 0.45, 0.45)
        # Zoom fixo (não fit-to-page): assim as quebras verticais ficam
        # previsíveis e dá para evitar título órfão no VBA.
        sh.set_print_scale(82)
        sh.hide_gridlines(2)
        sh.center_horizontally()
        sh.set_column("A:A", 14)
        sh.set_column("B:B", 26)
        sh.set_column("C:C", 14)
        sh.set_column("D:D", 28)
        sh.set_column("E:E", 22)
        sh.set_column("F:F", 32)
        sh.set_header("&CCecília — evolução dos exames")
        sh.set_footer("&LConferido nos laudos da pasta Exames  ·  coluna Arquivo = PDF  ·  não substitui o laudo&R&P / &N")

    def write_blocks(sh, titulo, subtitulo):
        sh.set_row(0, 22)
        sh.merge_range(0, 0, 0, 5, titulo, f_title)
        sh.merge_range(1, 0, 1, 5, subtitulo, f_sub)
        r = 3
        current_grupo = None
        mapa_local = []
        for i, t in enumerate(TABLES):
            start0 = r
            if t["grupo"] != current_grupo:
                current_grupo = t["grupo"]
                sh.set_row(r, 20)
                sh.merge_range(r, 0, r, 5, current_grupo, f_h2)
                r += 1
            sh.set_row(r, 16)
            sh.merge_range(r, 0, r, 5, t["titulo"], f_title_exam)
            r += 1
            sh.write(r, 0, t["mostra"], f_h3)
            r += 1
            largo = bool(t.get("largo"))
            if largo:
                sh.write(r, 0, t["cols"][0], f_cab)
                sh.write(r, 1, t["cols"][1], f_cab)
                sh.merge_range(r, 2, r, 3, t["cols"][2], f_cab)
                sh.write(r, 4, t["cols"][4] or "Nota", f_cab)
                sh.write(r, 5, "Arquivo", f_cab)
            else:
                for c, name in enumerate(t["cols"]):
                    sh.write(r, c, name or None, f_cab)
                sh.write(r, 5, "Arquivo", f_cab)
            r += 1
            for j, linha in enumerate(t["linhas"]):
                odd = j % 2
                cf = f_cell_b if odd else f_cell
                nf = f_num_b if odd else f_num
                df = f_date_b if odd else f_date
                lf = f_long_b if odd else f_long
                af = f_arq_b if odd else f_arq
                arq = arquivo_linha(t, linha)
                if largo:
                    resumo = linha[2] if len(linha) > 2 else ""
                    extra = len(str(resumo or "")) + len(arq)
                    sh.set_row(r, 40 if extra > 70 else 30)
                    val0 = fmt_value(linha[0])
                    if isinstance(val0, date):
                        sh.write_datetime(r, 0, val0, df)
                    else:
                        sh.write(r, 0, val0, cf)
                    sh.write(r, 1, linha[1] or None, cf)
                    sh.merge_range(r, 2, r, 3, resumo or "", lf)
                    sh.write(r, 4, linha[4] if len(linha) > 4 else None, lf)
                    sh.write(r, 5, arq or None, af)
                else:
                    sh.set_row(r, 15 if len(arq) < 34 else 24)
                    for c, val in enumerate(linha):
                        val = fmt_value(val)
                        if isinstance(val, date):
                            sh.write_datetime(r, c, val, df)
                        elif isinstance(val, (int, float)):
                            sh.write_number(r, c, val, nf)
                        elif val:
                            sh.write(r, c, val, cf)
                        else:
                            sh.write(r, c, None, cf)
                    sh.write(r, 5, arq or None, af)
                r += 1
            r += 1
            sh.set_row(r - 1, 6, f_empty)
            mapa_local.append((6 + i, start0 + 1, r))
        sh.merge_range(
            r, 0, r + 1, 5,
            "Legenda: isto resume os laudos da pasta Exames. A coluna Arquivo é o PDF (nome na pasta, sem .pdf). "
            "Não substitui o laudo. Interpretação na coluna Nota é só leitura do próprio laudo (faixa, subiu, caiu). Conduta é com o médico.",
            f_foot,
        )
        sh.print_area(0, 0, r + 1, 5)
        return mapa_local

    comp = wb.add_worksheet(ABA_DADOS_COMP)
    comp.set_tab_color("#7D3C98")
    setup_print_sheet(comp)
    mapa = write_blocks(
        comp,
        "Cecília Maria Albergaria Silva  ·  Evolução dos exames (todas as tabelas)",
        "Cada bloco é um exame, só com as datas em que ele foi feito. A coluna Arquivo é o PDF na pasta Exames (sem .pdf).",
    )

    sel = wb.add_worksheet(ABA_DADOS_SEL)
    sel.set_tab_color("#148F77")
    setup_print_sheet(sel)
    write_blocks(
        sel,
        "Cecília Maria Albergaria Silva  ·  Exames selecionados",
        "Igual à aba Dados Completo, só o marcado em Escolher. Coluna Arquivo = PDF na pasta Exames.",
    )

    # --------- Gráficos (faixa do laudo + ponto) ----------
    ds = wb.add_worksheet("DadosG")
    ds.hide()
    g_full = wb.add_worksheet(ABA_GRAF_COMP)
    g_sel = wb.add_worksheet(ABA_GRAF_SEL)
    g_full.set_tab_color("#1ABC9C")
    g_sel.set_tab_color("#117A65")
    graf_last_col = 11
    chart_px_w, chart_px_h, chart_rows = 720, 300, 16
    for sh in (g_full, g_sel):
        sh.set_landscape()
        sh.set_paper(9)
        sh.set_margins(0.4, 0.4, 0.45, 0.45)
        sh.fit_to_pages(1, 0)
        sh.hide_gridlines(2)
        sh.set_column("A:L", 10)
        sh.set_header("&CCecília — gráficos × faixa do laudo")
        sh.set_footer("&LBarra = faixa naquela data  ·  verde dentro  ·  vermelho fora&R&P / &N")

    def write_graf_header(sh, titulo, sub):
        sh.set_row(0, 22)
        sh.merge_range(0, 0, 0, graf_last_col, titulo, f_title)
        sh.merge_range(1, 0, 1, graf_last_col, sub, f_sub)
        sh.set_row(1, 32)

    write_graf_header(
        g_full,
        "Cecília — gráficos (todos os exames com faixa numérica)",
        "Em cada data, se o laudo tem faixa, a barra clara é essa faixa (muda com a idade). "
        "Ponto verde = dentro. Ponto vermelho = fora. Sem faixa de referência, só a linha do resultado. "
        "Gráfico só com dois ou mais pontos; um resultado só vai na tabela no final.",
    )
    write_graf_header(
        g_sel,
        "Cecília — gráficos selecionados",
        "Iguais aos da aba Graficos Completo, só com o que estiver marcado em Escolher. "
        "Habilite macros. Esta aba se atualiza ao você abrir ela. "
        "Barra clara = faixa do laudo; verde = dentro; vermelho = fora.",
    )

    def y_pad(values):
        nums = [x for x in values if x is not None]
        raw_min, raw_max = min(nums), max(nums)
        span = raw_max - raw_min
        if span <= 0:
            span = abs(raw_max) if raw_max else 1
        ymin = raw_min - span * 0.18
        ymax = raw_max + span * 0.18
        if raw_min >= 0 and ymin < 0:
            ymin = 0
        return ymin, ymax

    def date_pad(pontos):
        ds = [p["data"] for p in pontos if isinstance(p.get("data"), date)]
        if not ds:
            return None, None
        dmin, dmax = min(ds), max(ds)
        span = max((dmax - dmin).days, 30)
        pad_l = max(16, int(span * 0.04))
        pad_r = max(20, int(span * 0.05))
        return dmin - timedelta(days=pad_l), dmax + timedelta(days=pad_r)

    def style_chart(ch, title, y_name, ymin, ymax, date_axis=True, legend=None, x_name=None, xmin=None, xmax=None):
        ch.set_title({"none": True})
        xa = {
            "num_font": {"size": 11, "name": "Calibri", "rotation": 0},
            "label_position": "low",
        }
        if date_axis:
            xa["num_format"] = "MM/YYYY"
            xa["date_axis"] = True
            if xmin is not None:
                xa["min"] = xmin
            if xmax is not None:
                xa["max"] = xmax
                dias = (xmax - xmin).days if xmin is not None else 0
                if dias > 1400:
                    xa["major_unit"] = 1
                    xa["major_unit_type"] = "years"
                elif dias > 400:
                    xa["major_unit"] = 6
                    xa["major_unit_type"] = "months"
                else:
                    xa["major_unit"] = 3
                    xa["major_unit_type"] = "months"
        ch.set_x_axis(xa)
        ya = {
            "name": y_name,
            "name_font": {"size": 11, "name": "Calibri"},
            "num_font": {"size": 11, "name": "Calibri"},
        }
        if ymin is not None:
            ya["min"] = ymin
        if ymax is not None:
            ya["max"] = ymax
        ch.set_y_axis(ya)
        if legend is None:
            ch.set_legend({"none": True})
            ch.set_plotarea({"layout": {"x": 0.09, "y": 0.05, "width": 0.80, "height": 0.74}})
        else:
            legend["position"] = "bottom"
            ch.set_legend(legend)
            ch.set_plotarea({"layout": {"x": 0.09, "y": 0.04, "width": 0.80, "height": 0.62}})
        ch.set_size({"width": chart_px_w, "height": chart_px_h})
        ch.set_style(10)
        return ch

    def make_range_chart(col, first, last, title, y_name, ymin, ymax, xmin=None, xmax=None):
        ch = wb.add_chart({"type": "column", "subtype": "stacked"})
        cats = ["DadosG", first, col, last, col]
        ch.add_series({
            "name": "_base",
            "categories": cats,
            "values": ["DadosG", first, col + 4, last, col + 4],
            "fill": {"none": True},
            "border": {"none": True},
            "gap": 90,
        })
        ch.add_series({
            "name": "Faixa do laudo",
            "categories": cats,
            "values": ["DadosG", first, col + 5, last, col + 5],
            "fill": {"color": "#AED6F1"},
            "border": {"color": "#5DADE2", "width": 0.75},
        })
        ln = wb.add_chart({"type": "line"})
        ln.add_series({
            "name": "Evolução",
            "categories": cats,
            "values": ["DadosG", first, col + 1, last, col + 1],
            "line": {"color": "#85929E", "width": 1.25},
            "marker": {"type": "none"},
        })
        ln.add_series({
            "name": "Dentro da faixa",
            "categories": cats,
            "values": ["DadosG", first, col + 6, last, col + 6],
            "line": {"none": True},
            "marker": {
                "type": "circle", "size": 9,
                "border": {"color": "#145A32"}, "fill": {"color": "#196F3D"},
            },
        })
        ln.add_series({
            "name": "Fora da faixa",
            "categories": cats,
            "values": ["DadosG", first, col + 7, last, col + 7],
            "line": {"none": True},
            "marker": {
                "type": "diamond", "size": 10,
                "border": {"color": "#7B241C"}, "fill": {"color": "#C0392B"},
            },
        })
        ch.combine(ln)
        return style_chart(
            ch, title, y_name, ymin, ymax, date_axis=True,
            legend={"position": "bottom", "font": {"size": 11, "name": "Calibri"}, "delete_series": [0]},
            xmin=xmin, xmax=xmax,
        )

    def make_line_chart(col, first, last, title, y_name, ymin, ymax, date_axis=True, xmin=None, xmax=None):
        ln = wb.add_chart({"type": "line"})
        cats = ["DadosG", first, col, last, col]
        ln.add_series({
            "name": title,
            "categories": cats,
            "values": ["DadosG", first, col + 1, last, col + 1],
            "line": {"color": navy, "width": 1.5},
            "marker": {
                "type": "circle", "size": 8,
                "border": {"color": navy}, "fill": {"color": navy},
            },
        })
        return style_chart(ln, title, y_name, ymin, ymax, date_axis=date_axis, xmin=xmin, xmax=xmax)

    def make_idade_chart(col, first, last, title, xmin=None, xmax=None):
        ln = wb.add_chart({"type": "line"})
        cats = ["DadosG", first, col, last, col]
        ln.add_series({
            "name": "Idade cronológica",
            "categories": cats,
            "values": ["DadosG", first, col + 1, last, col + 1],
            "line": {"color": navy, "width": 1.75},
            "marker": {"type": "circle", "size": 8, "border": {"color": navy}, "fill": {"color": navy}},
        })
        ln.add_series({
            "name": "Idade óssea",
            "categories": cats,
            "values": ["DadosG", first, col + 2, last, col + 2],
            "line": {"color": "#C0392B", "width": 1.75},
            "marker": {"type": "diamond", "size": 8, "border": {"color": "#7B241C"}, "fill": {"color": "#C0392B"}},
        })
        return style_chart(
            ln, title, "meses", None, None, date_axis=True,
            legend={"position": "bottom", "font": {"size": 11, "name": "Calibri"}},
            xmin=xmin, xmax=xmax,
        )

    def place_chart(sh_row, caption, nota, builder):
        nonlocal row_g
        for sh in (g_full, g_sel):
            sh.set_row(row_g, 18)
            sh.merge_range(row_g, 0, row_g, graf_last_col, caption, f_title_exam)
        row_g += 1
        for sh in (g_full, g_sel):
            sh.write(row_g, 0, nota, f_h3)
        row_g += 1
        g_full.insert_chart(row_g, 0, builder(), {"object_position": 3})
        g_sel.insert_chart(row_g, 0, builder(), {"object_position": 3})
        row_g += chart_rows

    chart_idx = 0
    row_g = 3
    current_grupo = None
    mapa_graf = []
    unicos = []

    for i, t in enumerate(TABLES):
        series_list = GRAF.get(t["titulo"])
        if t["titulo"] == "Dosagens pontuais":
            for linha in t.get("linhas") or []:
                u = linha_unico_dosagem(t, linha)
                u["escolher_row"] = 6 + i
                unicos.append(u)
            mapa_graf.append((0, 0))
            continue
        if not series_list:
            mapa_graf.append((0, 0))
            continue
        specs_ok = [s for s in series_list if n_validos(s) >= 2]
        for spec in series_list:
            if n_validos(spec) < 2:
                u = linha_unico(t, spec)
                if u:
                    u["escolher_row"] = 6 + i
                    unicos.append(u)
        if not specs_ok:
            mapa_graf.append((0, 0))
            continue
        start_g = row_g
        if t["grupo"] != current_grupo:
            current_grupo = t["grupo"]
            for sh in (g_full, g_sel):
                sh.set_row(row_g, 20)
                sh.merge_range(row_g, 0, row_g, graf_last_col, current_grupo, f_h2)
            row_g += 1
        for spec in specs_ok:
            titulo_c = spec.get("titulo") or t["titulo"]
            unidade = spec.get("unidade") or ""
            tipo = spec.get("tipo", "valor")
            pontos = spec["pontos"]
            col = chart_idx * 9
            first, last = 2, 1 + len(pontos)
            xmin, xmax = date_pad(pontos)

            if tipo == "idade":
                ds.write(0, col, titulo_c)
                ds.write_row(1, col, ["Data", "Cronológica", "Óssea"])
                for j, p in enumerate(pontos):
                    r = 2 + j
                    ds.write_datetime(r, col, p["data"])
                    ds.write_number(r, col + 1, p["crono"])
                    ds.write_number(r, col + 2, p["ossea"])
                place_chart(
                    row_g,
                    f"{titulo_c}  (meses)",
                    "Sem faixa de laboratório: compara idade cronológica × idade óssea.",
                    lambda c=col, xa=xmin, xb=xmax: make_idade_chart(c, first, last, titulo_c, xa, xb),
                )
            elif tipo == "tempo":
                ds.write(0, col, titulo_c)
                ds.write_row(1, col, ["Tempo", "Valor"])
                vs = []
                for j, p in enumerate(pontos):
                    r = 2 + j
                    ds.write(r, col, p["cat"])
                    if p["v"] is None:
                        ds.write_formula(r, col + 1, "=NA()")
                    else:
                        ds.write_number(r, col + 1, p["v"])
                        vs.append(p["v"])
                ymin, ymax = y_pad(vs) if vs else (None, None)
                place_chart(
                    row_g,
                    f"{titulo_c}  ({unidade})",
                    spec.get("nota") or "Sem faixa de referência neste laudo: só a evolução do resultado.",
                    lambda c=col, yn=ymin, yx=ymax, un=unidade, tt=titulo_c: make_line_chart(
                        c, first, last, tt, un, yn, yx, date_axis=False
                    ),
                )
            else:
                tem_faixa = all(p.get("lo") is not None and p.get("hi") is not None for p in pontos)
                ds.write(0, col, titulo_c)
                ds.write_row(1, col, ["Data", "Valor", "Piso", "Teto", "Pad", "Faixa", "Dentro", "Fora"])
                vs, los, his = [], [], []
                for j, p in enumerate(pontos):
                    r = 2 + j
                    v, lo, hi = p["v"], p.get("lo"), p.get("hi")
                    ds.write_datetime(r, col, p["data"])
                    ds.write_number(r, col + 1, v)
                    vs.append(v)
                    if tem_faixa:
                        ds.write_number(r, col + 2, lo)
                        ds.write_number(r, col + 3, hi)
                        ds.write_number(r, col + 4, lo)
                        ds.write_number(r, col + 5, max(hi - lo, 0))
                        if lo <= v <= hi:
                            ds.write_number(r, col + 6, v)
                            ds.write_formula(r, col + 7, "=NA()")
                        else:
                            ds.write_formula(r, col + 6, "=NA()")
                            ds.write_number(r, col + 7, v)
                        los.append(lo)
                        his.append(hi)
                if tem_faixa:
                    ymin, ymax = y_pad(vs + los + his)
                    place_chart(
                        row_g,
                        f"{titulo_c}  ({unidade})",
                        "Barra azul-clara = faixa do laudo nesta data. Verde = dentro. Vermelho = fora.",
                        lambda c=col, yn=ymin, yx=ymax, un=unidade, tt=titulo_c, xa=xmin, xb=xmax: make_range_chart(
                            c, first, last, tt, un, yn, yx, xa, xb
                        ),
                    )
                else:
                    ymin, ymax = y_pad(vs)
                    place_chart(
                        row_g,
                        f"{titulo_c}  ({unidade})",
                        "Sem faixa de referência neste laudo: só a evolução do resultado.",
                        lambda c=col, yn=ymin, yx=ymax, un=unidade, tt=titulo_c, xa=xmin, xb=xmax: make_line_chart(
                            c, first, last, tt, un, yn, yx, True, xa, xb
                        ),
                    )
            chart_idx += 1
        mapa_graf.append((start_g + 1, row_g))

    if unicos:
        row_g += 1
        for sh in (g_full, g_sel):
            sh.set_row(row_g, 22)
            sh.merge_range(row_g, 0, row_g, graf_last_col, "Exames com um único resultado", f_h2)
        row_g += 1
        for sh in (g_full, g_sel):
            sh.write(row_g, 0, "Sem gráfico: só há um ponto no laudo (inclui as dosagens pontuais). A evolução aparece quando houver uma nova coleta.", f_h3)
        row_g += 1
        cab_u = ["Exame", "Data", "Resultado", "Unidade", "Faixa do laudo", "Arquivo"]
        for sh in (g_full, g_sel):
            for c, name in enumerate(cab_u):
                sh.write(row_g, c, name, f_cab)
        row_g += 1
        for j, u in enumerate(unicos):
            odd = j % 2
            cf = f_cell_b if odd else f_cell
            nf = f_num_b if odd else f_num
            df = f_date_b if odd else f_date
            af = f_arq_b if odd else f_arq
            for sh in (g_full, g_sel):
                sh.set_row(row_g, 20)
                sh.write(row_g, 0, u["exame"], cf)
                if isinstance(u["data"], date):
                    sh.write_datetime(row_g, 1, u["data"], df)
                else:
                    sh.write(row_g, 1, u["data"] or None, cf)
                if isinstance(u["valor"], (int, float)):
                    sh.write_number(row_g, 2, u["valor"], nf)
                else:
                    sh.write(row_g, 2, u["valor"], cf)
                sh.write(row_g, 3, u["unidade"] or None, cf)
                sh.write(row_g, 4, u["faixa"] or None, cf)
                sh.write(row_g, 5, u["arquivo"] or None, af)
            row_g += 1

    for sh in (g_full, g_sel):
        sh.print_area(0, 0, max(row_g, 4), graf_last_col)

    print("charts", chart_idx, "unicos", len(unicos))

    un = wb.add_worksheet(ABA_UNICOS)
    un.write_row(0, 0, ["exame", "data", "valor", "unidade", "faixa", "arquivo", "escolher_row"])
    for j, u in enumerate(unicos, 1):
        un.write(j, 0, u["exame"])
        if isinstance(u["data"], date):
            un.write_datetime(j, 1, u["data"])
        else:
            un.write(j, 1, u["data"] or None)
        if isinstance(u["valor"], (int, float)):
            un.write_number(j, 2, u["valor"])
        else:
            un.write(j, 2, u["valor"])
        un.write(j, 3, u["unidade"] or None)
        un.write(j, 4, u["faixa"] or None)
        un.write(j, 5, u["arquivo"] or None)
        un.write_number(j, 6, u["escolher_row"])
    un.hide()

    mp = wb.add_worksheet("Mapa")
    mp.write_row(0, 0, ["escolher_row", "sel_ini", "sel_fim", "graf_ini", "graf_fim"])
    for j, row in enumerate(mapa, 1):
        gi, gf = mapa_graf[j - 1]
        mp.write_row(j, 0, [row[0], row[1], row[2], gi, gf])
    mp.hide()

    # --------- Como usar ----------
    how = wb.add_worksheet("Como usar")
    how.hide_gridlines(2)
    how.set_column("A:A", 110)
    how.write(0, 0, "Como usar", f_title)
    lines = [
        "",
        "Aba Escolher — clique no quadradinho à esquerda. Marca e desmarca a cada clique.",
        "Aba Dados Completo — sempre todas as tabelas, independente das caixinhas. Coluna Arquivo = PDF daquela linha.",
        "Aba Dados Selecionados — as tabelas que você marcou. Também tem a coluna Arquivo. Imprimir em paisagem. Precisa habilitar macros.",
        "Para imprimir gráficos em retrato: escolha 3 ou 4 por folha e clique Imprimir selecionados ou Imprimir todos.",
        "Gráfico só com dois ou mais resultados. Exame com um dado só (ex.: pH, B12, Fator XIII) vai numa tabela no final, sem gráfico.",
        "Aba Graficos Completo — exames numéricos. Se o laudo tem faixa, a barra é essa faixa naquela data (verde dentro, vermelho fora). Sem faixa, aparece só a linha do resultado.",
        "Aba Graficos Selecionados — os mesmos gráficos, só dos itens marcados. Habilite macros; ela atualiza ao você abrir ela.",
        "A faixa muda com a idade (ex.: IGF-1, hemoglobina, leucócitos). Cada ponto usa a faixa daquele laudo.",
        "Sem gráfico: o que é qualitativo (urina descritiva, genética, imagens, pezinho, grupo sanguíneo, suor em foto).",
        "Valores conferidos nos PDFs da pasta Exames. Não substitui o laudo. Conduta é com o médico.",
    ]
    for i, line in enumerate(lines):
        how.set_row(i + 1, 32 if line else 10)
        how.write(i + 1, 0, line, f_help)

    ws.activate()
    wb.close()


def main():
    dest = DEST
    try:
        build(dest)
        print("SAVED", dest)
    except PermissionError:
        alt = dest.with_name("Evolução Exames - novo.xlsx")
        build(alt)
        print("SAVED_ALT", alt)


if __name__ == "__main__":
    main()
