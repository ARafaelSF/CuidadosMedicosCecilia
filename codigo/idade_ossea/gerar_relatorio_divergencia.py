# -*- coding: utf-8 -*-
"""PDF: mini relatório da divergência idade óssea jul/2026, com imagens de referência."""
from __future__ import annotations

from datetime import date
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
from matplotlib.backends.backend_pdf import PdfPages
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parents[2]
EXAMES = ROOT / "Exames"
OUT = EXAMES / "Idade Ossea - Relatorio divergencia jul-2026.pdf"
ASSETS = Path(__file__).resolve().parent / "_assets_relatorio"
DOB = date(2020, 6, 30)
TRATAMENTO = date(2025, 4, 18)

RX_FEV = EXAMES / "Imagem - 2026-02-18 - Júlia Martins Azevedo Eyer Thomaz - Punho.pdf"
RX_JUL = EXAMES / "Imagem - 2026-07-29 - Fernanda de Souza Silva - Punho.pdf"


def _font(size: int = 16, bold: bool = False):
    names = ["arialbd.ttf", "Arial Bold.ttf"] if bold else ["arial.ttf", "Arial.ttf", "segoeui.ttf"]
    for name in names:
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            continue
    return ImageFont.load_default()


def pdf_page(pdf_path: Path, page: int, scale: float = 2.0) -> Image.Image:
    import pymupdf

    doc = pymupdf.open(pdf_path)
    pix = doc[page].get_pixmap(matrix=pymupdf.Matrix(scale, scale), alpha=False)
    return Image.frombytes("RGB", [pix.width, pix.height], pix.samples)


def crop_wrist(im: Image.Image) -> Image.Image:
    w, h = im.size
    return im.crop((int(w * 0.10), int(h * 0.28), int(w * 0.74), int(h * 0.76)))


def crop_ulnar(im: Image.Image) -> Image.Image:
    w, h = im.size
    return im.crop((int(w * 0.02), int(h * 0.12), int(w * 0.38), int(h * 0.42)))


def annotate(im: Image.Image, labels: list[tuple[tuple[int, int], str, str]]) -> Image.Image:
    out = im.copy()
    dr = ImageDraw.Draw(out)
    f = _font(14)
    for (x, y), text, color in labels:
        dr.ellipse((x - 9, y - 9, x + 9, y + 9), outline=color, width=3)
        dr.text((x + 12, y - 7), text, fill=color, font=f)
    return out


def side_by_side(
    left: Image.Image,
    right: Image.Image,
    title_left: str,
    title_right: str,
    header: str,
    footer: str = "",
) -> Image.Image:
    lh = max(left.height, right.height)
    left = left.resize((int(left.width * lh / left.height), lh), Image.Resampling.LANCZOS)
    right = right.resize((int(right.width * lh / right.height), lh), Image.Resampling.LANCZOS)
    pad, bar = 24, 56
    w = left.width + right.width + pad * 3
    h = lh + bar * 2 + (40 if footer else 0)
    out = Image.new("RGB", (w, h), "white")
    dr = ImageDraw.Draw(out)
    dr.text((pad, 12), header, fill="#1F4E79", font=_font(20, bold=True))
    dr.text((pad, bar - 8), title_left, fill="#C0392B", font=_font(15, bold=True))
    dr.text((left.width + pad * 2, bar - 8), title_right, fill="#E67E22", font=_font(15, bold=True))
    out.paste(left, (pad, bar))
    out.paste(right, (left.width + pad * 2, bar))
    if footer:
        dr.text((pad, h - 32), footer, fill="#555", font=_font(13))
    return out


def gp_reference_table() -> Image.Image:
    w, h = 1500, 820
    im = Image.new("RGB", (w, h), "white")
    dr = ImageDraw.Draw(im)
    dr.text((36, 20), "Marcos Greulich-Pyle (feminino) — referência para leitura", fill="#1F4E79", font=_font(22, bold=True))
    dr.text((36, 58), "Descrição publicada dos estágios (atlas GP, 1959). Não substitui lâmina oficial.", fill="#666", font=_font(13))

    blocks = [
        ("3a6m — laudo fev/26 ✓ concordante", "#C0392B", [
            "Capitato + hamato maduros",
            "Semilunar + piriforme visíveis",
            "Trapezo incipiente",
            "Escafóide ausente/incipiente",
            "Epífise ulnar AUSENTE",
            "≈ 4 ossos do carpo",
        ], "Corrobora: RX fev/2026 (foto ao lado)"),
        ("4a2m — laudo jul/26", "#E67E22", [
            "Escafóide + trapezoide visíveis",
            "Trapezo presente",
            "Epífise radial bem formada",
            "Epífise ulnar ainda mínima/ausente no GP clássico",
            "≈ 5–6 ossos do carpo",
            "Pisiforme ausente",
        ], "Corrobora: progressão +8m ósseos em 5m (GH)"),
        ("~5a2m — leitura IA (jul/26)", "#27AE60", [
            "7 ossos do carpo (sem pisiforme)",
            "Escafóide/trapezoide definidos",
            "Epífise ulnar incipiente visível",
            "No GP clássico ulnar ≈ 6a",
            "Pode adiantar com somatropina",
        ], "Corrobora: zoom ulnar jul/26"),
    ]
    cw = 460
    for i, (title, color, lines, note) in enumerate(blocks):
        x = 36 + i * (cw + 24)
        dr.rectangle((x, 100, x + cw, 760), outline=color, width=3)
        dr.text((x + 14, 118), title, fill=color, font=_font(16, bold=True))
        y = 160
        for ln in lines:
            dr.text((x + 18, y), "• " + ln, fill="#222", font=_font(14))
            y += 32
        dr.text((x + 14, 680), note, fill=color, font=_font(13, bold=True))
    return im


def pil_to_ax(ax, im: Image.Image) -> None:
    ax.imshow(np.asarray(im))
    ax.axis("off")


def add_footer(fig, text: str, y: float = 0.03) -> None:
    fig.text(0.06, y, text, fontsize=8.5, color="#555", va="bottom", wrap=True)


def main() -> None:
    ASSETS.mkdir(exist_ok=True)

    rx_fev_full = pdf_page(RX_FEV, 1, 2.0)
    rx_jul_full = pdf_page(RX_JUL, 1, 2.0)
    laudo_fev = pdf_page(RX_FEV, 0, 2.0)
    laudo_jul = pdf_page(RX_JUL, 0, 2.0)

    wrist_fev = crop_wrist(rx_fev_full)
    wrist_jul = crop_wrist(rx_jul_full)
    ulnar_fev = crop_ulnar(wrist_fev)
    ulnar_jul = crop_ulnar(wrist_jul)

    wf, hf = wrist_jul.size
    jul_ann = annotate(
        wrist_jul,
        [
            ((int(wf * 0.52), int(hf * 0.55)), "Capitato", "#C0392B"),
            ((int(wf * 0.45), int(hf * 0.62)), "Semilunar", "#C0392B"),
            ((int(wf * 0.38), int(hf * 0.52)), "Escafóide", "#E67E22"),
            ((int(wf * 0.58), int(hf * 0.44)), "Trapezoide", "#E67E22"),
            ((int(wf * 0.72), int(hf * 0.18)), "Ep. radial", "#3498DB"),
            ((int(wf * 0.10), int(hf * 0.22)), "Ep. ulnar?", "#27AE60"),
        ],
    )
    wf2, hf2 = wrist_fev.size
    fev_ann = annotate(
        wrist_fev,
        [
            ((int(wf2 * 0.52), int(hf2 * 0.55)), "Capitato", "#C0392B"),
            ((int(wf2 * 0.45), int(hf2 * 0.62)), "Semilunar", "#C0392B"),
            ((int(wf2 * 0.72), int(hf2 * 0.18)), "Ep. radial", "#3498DB"),
            ((int(wf2 * 0.10), int(hf2 * 0.22)), "Sem ep. ulnar", "#C0392B"),
        ],
    )

    panel_compare = side_by_side(
        fev_ann, jul_ann,
        "18/02/2026 — laudo 3a6m (IA concordante)",
        "29/07/2026 — laudo 4a2m (IA ~5a2m)",
        "Referência interna: mesmo paciente, 5 meses depois",
        "Vermelho = marcos do laudo fev · Laranja/verde = novos marcos que puxaram a IA em jul",
    )
    panel_ulnar = side_by_side(
        ulnar_fev, ulnar_jul,
        "Fev/26 — sem epífise ulnar",
        "Jul/26 — epífise ulnar incipiente (seta verde na página anterior)",
        "Zoom: extremidade ulnar distal (marco crítico da divergência)",
        "No GP clássico, epífise ulnar feminina ≈ 6a — mas pode adiantar com GH.",
    )
    gp_table = gp_reference_table()

    for name, img in [
        ("panel_compare", panel_compare),
        ("panel_ulnar", panel_ulnar),
        ("jul_ann", jul_ann),
        ("gp_table", gp_table),
    ]:
        img.save(ASSETS / f"{name}.png")

    with PdfPages(OUT) as pdf:
        meta = pdf.infodict()
        meta["Title"] = "Idade óssea — relatório divergência jul/2026"
        meta["Author"] = "Análise visual assistida (não substitui radiologista)"

        # Pág 1 — capa
        fig = plt.figure(figsize=(8.27, 11.69))
        fig.patch.set_facecolor("white")
        fig.text(0.5, 0.88, "Mini relatório — idade óssea", ha="center", fontsize=18, fontweight="bold", color="#1F4E79")
        fig.text(0.5, 0.83, "Divergência laudo × IA · RX punho 29/07/2026", ha="center", fontsize=12, color="#444")
        lines = [
            f"Cecília Maria Albergaria Silva · nasc. {DOB.strftime('%d/%m/%Y')}",
            "",
            "Resumo",
            "• Cronológica: 6a1m",
            "• Laudo radiologista: 4a2m (atraso ~23 meses)",
            "• Estimativa IA: ~5a2m (atraso ~11 meses)",
            f"• Somatropina desde {TRATAMENTO.strftime('%d/%m/%Y')}",
            "",
            "Este PDF reúne:",
            "1. Radiografias da Cecília (fev e jul/2026) anotadas",
            "2. Zoom na epífise ulnar (marco da divergência)",
            "3. Tabela de referência Greulich-Pyle (feminino)",
            "4. Laudos escaneados (o que o radiologista registrou)",
            "5. Conclusão: por que a IA divergiu e por que o laudo 4a2m é plausível",
            "",
            "⚠ Análise de apoio — não substitui laudo oficial.",
        ]
        y = 0.74
        for ln in lines:
            fw = "bold" if ln in ("Resumo", "Este PDF reúne:") else "normal"
            fig.text(0.1, y, ln, fontsize=10 if ln else 6, fontweight=fw, color="#222")
            y -= 0.028 if ln else 0.012
        pdf.savefig(fig)
        plt.close(fig)

        # Pág 2 — laudos (referência do que foi laudado)
        fig = plt.figure(figsize=(8.27, 11.69))
        fig.text(0.5, 0.97, "Laudos — referência do radiologista", ha="center", fontsize=13, fontweight="bold", color="#1F4E79")
        ax1 = fig.add_axes([0.06, 0.52, 0.88, 0.40])
        ax2 = fig.add_axes([0.06, 0.08, 0.88, 0.40])
        pil_to_ax(ax1, laudo_fev)
        pil_to_ax(ax2, laudo_jul)
        ax1.set_title("10/02/2025 — laudo 2a6m (sem RX neste PDF)", fontsize=9)
        ax2.set_title("29/07/2026 — laudo 4a2m  ←  caso em análise", fontsize=9, color="#C0392B")
        add_footer(fig, "Texto oficial assinado pelo Dr. Alexandre Paci Galvão (CRM-MG 19336). Método: Greulich-Pyle feminino.")
        pdf.savefig(fig)
        plt.close(fig)

        # Pág 3 — RX completas
        fig = plt.figure(figsize=(8.27, 11.69))
        fig.text(0.5, 0.97, "Radiografias completas (mão/punho E)", ha="center", fontsize=13, fontweight="bold", color="#1F4E79")
        ax1 = fig.add_axes([0.06, 0.52, 0.88, 0.40])
        ax2 = fig.add_axes([0.06, 0.08, 0.88, 0.40])
        pil_to_ax(ax1, rx_fev_full)
        pil_to_ax(ax2, rx_jul_full)
        ax1.set_title("18/02/2026", fontsize=9)
        ax2.set_title("29/07/2026", fontsize=9)
        add_footer(fig, "Imagens originais dos exames. Análise de punho nas páginas seguintes.")
        pdf.savefig(fig)
        plt.close(fig)

        # Pág 4 — comparativo anotado
        fig = plt.figure(figsize=(8.27, 11.69))
        ax = fig.add_axes([0.04, 0.10, 0.92, 0.84])
        pil_to_ax(ax, panel_compare)
        pdf.savefig(fig)
        plt.close(fig)

        # Pág 5 — zoom ulnar
        fig = plt.figure(figsize=(8.27, 11.69))
        ax = fig.add_axes([0.04, 0.12, 0.92, 0.80])
        pil_to_ax(ax, panel_ulnar)
        add_footer(fig, "Corroboração da divergência: epífise ulnar ausente em fev (3a6m) e incipiente em jul.")
        pdf.savefig(fig)
        plt.close(fig)

        # Pág 6 — GP reference
        fig = plt.figure(figsize=(11.69, 8.27))
        ax = fig.add_axes([0.03, 0.08, 0.94, 0.86])
        pil_to_ax(ax, gp_table)
        pdf.savefig(fig, orientation="landscape")
        plt.close(fig)

        # Pág 7 — jul anotado grande
        fig = plt.figure(figsize=(8.27, 11.69))
        fig.text(0.5, 0.97, "Jul/2026 — punho anotado (achados IA)", ha="center", fontsize=13, fontweight="bold", color="#1F4E79")
        ax = fig.add_axes([0.08, 0.18, 0.84, 0.72])
        pil_to_ax(ax, jul_ann)
        legend = (
            "Vermelho: carpais presentes em fev e jul · Laranja: escafóide/trapezoide (GP ~4a+)\n"
            "Azul: epífise radial · Verde: epífise ulnar incipiente (marco que elevou a IA para ~5a)"
        )
        add_footer(fig, legend, y=0.06)
        pdf.savefig(fig)
        plt.close(fig)

        # Pág 8 — conclusão
        fig = plt.figure(figsize=(8.27, 11.69))
        fig.text(0.08, 0.94, "Conclusão", fontsize=14, fontweight="bold", color="#1F4E79")
        conclusion = [
            "Por que a IA leu ~5a2m em jul/2026:",
            "• Contagem de ~7 ossos do carpo + escafóide/trapezoide visíveis",
            "• Epífise ulnar distal incipiente (zoom pág. 5) — no GP clássico ≈ 6a em meninas",
            "",
            "Por que o laudo 4a2m permanece plausível:",
            "• Fev/2026: laudo 3a6m concordou com a imagem (~4 carpais, sem ep. ulnar)",
            "• Jul/2026: progressão de +8 meses ósseos em ~5 meses reais — coerente com GH",
            "• Método GP compara a mão inteira ao padrão, não um único marco",
            "• Epífise ulnar pode adiantar-se com somatropina sem equivaler a 5–6a cronológicos",
            "",
            "Recomendação:",
            "Levar este relatório e o comparativo (PDF irmão) ao endocrino.",
            "Se houver dúvida clínica, solicitar releitura radiológica.",
            "",
            "Regenerar: py -3 codigo/idade_ossea/gerar_relatorio_divergencia.py",
        ]
        y = 0.88
        for ln in conclusion:
            bold = ln.startswith("Por ") or ln.startswith("Recomendação")
            fig.text(0.08, y, ln, fontsize=10, fontweight="bold" if bold else "normal", color="#222")
            y -= 0.032 if ln else 0.015
        pdf.savefig(fig)
        plt.close(fig)

    print(f"OK {OUT} ({8} páginas)")


if __name__ == "__main__":
    main()
