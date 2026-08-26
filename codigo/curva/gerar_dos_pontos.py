# -*- coding: utf-8 -*-
"""Gera curvas OMS com pontos do pontos_altura.csv + seta início tratamento.

Calibração 0–5 derivada das linhas vetoriais do PDF oficial OMS.
"""
from __future__ import annotations

import csv
from datetime import date
from pathlib import Path

import cv2
import pymupdf
from PIL import Image, ImageDraw, ImageFont

TMP = Path(r"C:\Users\arafa\AppData\Local\Temp\curva")
OUT = Path(r"C:\Users\arafa\OneDrive\Documentos\Cecília\_Organizado") / "Exames" / "Curva de Crescimento"
CSV_PATH = OUT / "pontos_altura.csv"
DOB = date(2020, 6, 30)
TRATAMENTO = date(2025, 4, 18)

# Coordenadas PDF (pt) das grades do cht-lhfa-girls-z-0-5
PDF_0_5 = {
    "age0_x": 110.86,   # Birth
    "age1_x": 715.37,   # 5 years
    "h_top_y": 119.92,  # 125 cm
    "h_bot_y": 483.92,  # 45 cm
    "age0": 0.0,
    "age1": 5.0,
    "h0": 45.0,
    "h1": 125.0,
}


def idade_anos(anos: int, meses: int) -> float:
    return anos + meses / 12.0


def idade_em(d: date) -> float:
    return (d - DOB).days / 365.25


def load_pontos(path: Path) -> list[dict]:
    text = path.read_text(encoding="utf-8-sig")
    lines = [ln for ln in text.splitlines() if ln.strip()]
    if lines and lines[0].lower().startswith("idade"):
        lines = lines[1:]
    rows = []
    for r in csv.DictReader(lines, delimiter=";"):
        anos = int(str(r["anos"]).strip())
        meses = int(str(r["meses"]).strip())
        alt = float(str(r["altura_cm"]).strip().replace(",", "."))
        rows.append({
            "anos": anos,
            "meses": meses,
            "idade": idade_anos(anos, meses),
            "altura": alt,
            "fonte": (r.get("fonte") or "").strip(),
            "label": f"{anos}a{meses}m",
        })
    return rows


def data_to_xy(age, height, bbox, age0, age1, h0, h1):
    x0, y0, x1, y1 = bbox
    px = x0 + (age - age0) / (age1 - age0) * (x1 - x0)
    py = y1 - (height - h0) / (h1 - h0) * (y1 - y0)
    return px, py


def bbox_from_pdf(calib: dict, scale: float) -> tuple[int, int, int, int]:
    return (
        int(round(calib["age0_x"] * scale)),
        int(round(calib["h_top_y"] * scale)),
        int(round(calib["age1_x"] * scale)),
        int(round(calib["h_bot_y"] * scale)),
    )


def plot_bbox_5_19(w: int, h: int) -> tuple[int, int, int, int]:
    return int(w * 0.125), int(h * 0.180), int(w * 0.880), int(h * 0.860)


def try_font(size: int):
    for name in ("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/calibri.ttf"):
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            continue
    return ImageFont.load_default()


def mark_chart(
    src_pdf: Path,
    dest_pdf: Path,
    pontos: list[dict],
    *,
    age0: float,
    age1: float,
    h0: float,
    h1: float,
    bbox: tuple[int, int, int, int] | None,
    bbox_fn,
    titulo: str,
    tratamento_age: float | None,
    tratamento_altura: float | None,
    scale: float = 3.0,
) -> None:
    doc = pymupdf.open(src_pdf)
    page = doc[0]
    pdf_w, pdf_h = page.rect.width, page.rect.height
    pix = page.get_pixmap(matrix=pymupdf.Matrix(scale, scale), alpha=False)
    png = TMP / (src_pdf.stem + "_hires.png")
    pix.save(str(png))
    doc.close()

    img = cv2.imread(str(png))
    h, w = img.shape[:2]
    if bbox is None:
        bbox = bbox_fn(w, h)
    print(f"  bbox {dest_pdf.name}: {bbox} size={w}x{h}")

    pil = Image.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB)).convert("RGBA")
    overlay = Image.new("RGBA", pil.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    font = try_font(22)
    font_sm = try_font(16)

    pts_px = []
    for p in pontos:
        if not (age0 <= p["idade"] <= age1 and h0 <= p["altura"] <= h1):
            continue
        px, py = data_to_xy(p["idade"], p["altura"], bbox, age0, age1, h0, h1)
        pts_px.append((px, py, p))
        print(f"    {p['label']:6} age={p['idade']:.3f} h={p['altura']} -> x={px:.0f} y={py:.0f}")

    if len(pts_px) >= 2:
        draw.line([(x, y) for x, y, _ in pts_px], fill=(21, 101, 192, 210), width=4)

    for px, py, p in pts_px:
        r = 8
        draw.ellipse(
            (px - r, py - r, px + r, py + r),
            fill=(21, 101, 192, 235),
            outline=(13, 71, 161, 255),
            width=2,
        )

    if tratamento_age is not None and tratamento_altura is not None:
        if age0 <= tratamento_age <= age1 and h0 <= tratamento_altura <= h1:
            tx, ty = data_to_xy(tratamento_age, tratamento_altura, bbox, age0, age1, h0, h1)
            print(f"    TRATAMENTO age={tratamento_age:.3f} -> x={tx:.0f} y={ty:.0f} (limite age1 x={bbox[2]})")
            r = 11
            draw.ellipse((tx - r, ty - r, tx + r, ty + r), outline=(198, 40, 40, 255), width=4)
            tip_y = ty - 16
            base_y = ty - 110
            draw.line((tx, base_y + 36, tx, tip_y), fill=(198, 40, 40, 255), width=4)
            draw.polygon(
                [(tx, tip_y), (tx - 11, tip_y - 20), (tx + 11, tip_y - 20)],
                fill=(198, 40, 40, 255),
            )
            label = "Início do tratamento"
            sub = "18/04/2025"
            tw, th = 260, 54
            # Preferir à esquerda do ponto para não cruzar a borda dos 5 anos
            bx0 = tx - tw - 12
            by0 = base_y - 20
            if bx0 < 8:
                bx0 = tx + 16
            if by0 < 8:
                by0 = 8
            if bx0 + tw > w - 8:
                bx0 = w - 8 - tw
            box = (bx0, by0, bx0 + tw, by0 + th)
            draw.rounded_rectangle(box, radius=8, fill=(255, 255, 255, 240), outline=(198, 40, 40, 255), width=2)
            draw.text((box[0] + 12, box[1] + 8), label, fill=(183, 28, 28, 255), font=font_sm)
            draw.text((box[0] + 12, box[1] + 28), sub, fill=(80, 80, 80, 255), font=font_sm)

    draw.rounded_rectangle((16, 12, 640, 72), radius=8, fill=(255, 255, 255, 230), outline=(21, 101, 192, 255), width=2)
    draw.text((28, 20), titulo, fill=(13, 71, 161, 255), font=font)
    draw.text((28, 46), "Fonte das medidas: pontos_altura.csv", fill=(90, 90, 90, 255), font=font_sm)

    out_img = Image.alpha_composite(pil, overlay).convert("RGB")
    out_png = TMP / (src_pdf.stem + "_marcado.png")
    out_img.save(out_png)

    marked = pymupdf.open()
    pg = marked.new_page(width=pdf_w, height=pdf_h)
    pg.insert_image(pg.rect, filename=str(out_png))
    dest_pdf.parent.mkdir(parents=True, exist_ok=True)
    marked.save(dest_pdf)
    marked.close()
    print("SAVED", dest_pdf.name, "pontos", len(pts_px))


def main() -> None:
    TMP.mkdir(exist_ok=True)
    pontos = load_pontos(CSV_PATH)
    print("PONTOS", len(pontos))
    for p in pontos:
        print(f"  {p['label']:6} idade={p['idade']:.3f}a  {p['altura']} cm")

    trat_age = idade_em(TRATAMENTO)
    sorted_p = sorted(pontos, key=lambda x: x["idade"])
    trat_h = None
    for i, p in enumerate(sorted_p):
        if abs(p["idade"] - trat_age) < 0.05:
            trat_h = p["altura"]
            break
        if p["idade"] > trat_age and i > 0:
            a, b = sorted_p[i - 1], p
            t = (trat_age - a["idade"]) / (b["idade"] - a["idade"])
            trat_h = a["altura"] + t * (b["altura"] - a["altura"])
            break
    if trat_h is None and sorted_p:
        trat_h = sorted_p[-1]["altura"]
    print(f"TRATAMENTO age={trat_age:.3f}a (~{trat_age * 12:.1f}m) altura≈{trat_h:.1f} cm")

    src_0_5 = TMP / "cht-lhfa-girls-z-0-5.pdf"
    src_5_19 = TMP / "cht-hfa-girls-z-5-19years.pdf"
    if not src_0_5.exists():
        src_0_5 = OUT / "01 - OMS oficial limpo - Altura meninas 0-5 anos (z-score).pdf"
    if not src_5_19.exists():
        src_5_19 = OUT / "03 - OMS oficial limpo - Altura meninas 5-19 anos (z-score).pdf"

    scale = 3.0
    bbox_0_5 = bbox_from_pdf(PDF_0_5, scale)

    pts_0_5 = [p for p in pontos if p["idade"] <= 5.0]
    mark_chart(
        src_0_5,
        OUT / "02 - OMS Altura meninas 0-5 anos COM PONTOS.pdf",
        pts_0_5,
        age0=0.0,
        age1=5.0,
        h0=45.0,
        h1=125.0,
        bbox=bbox_0_5,
        bbox_fn=None,
        titulo="Cecília · DN 30/06/2020 · Altura 0–5 anos",
        tratamento_age=trat_age if trat_age <= 5.0 else None,
        tratamento_altura=trat_h if trat_age <= 5.0 else None,
        scale=scale,
    )

    pts_5_19 = [p for p in pontos if p["idade"] >= 5.0]
    mark_chart(
        src_5_19,
        OUT / "04 - OMS Altura meninas 5-19 anos COM PONTOS.pdf",
        pts_5_19,
        age0=5.0,
        age1=19.0,
        h0=90.0,
        h1=180.0,
        bbox=None,
        bbox_fn=plot_bbox_5_19,
        titulo="Cecília · DN 30/06/2020 · Altura 5–19 anos",
        tratamento_age=trat_age if trat_age >= 5.0 else None,
        tratamento_altura=trat_h if trat_age >= 5.0 else None,
        scale=scale,
    )
    print("OK")


if __name__ == "__main__":
    main()
