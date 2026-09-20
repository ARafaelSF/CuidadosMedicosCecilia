# -*- coding: utf-8 -*-
"""Junta PDFs de Relatórios por tipo + profissional → pasta Imprimir/.

Originais intactos. Saída: Imprimir - {Tipo} - {Profissional}.pdf
Exceção: Alta → um único Imprimir - Alta.pdf (todos os profissionais, por data).

Natação: remove capa + folha de frequência; data no topo (só Imprimir/).
Fisioterapia com --so-novos: só PDFs ainda não listados no manifest (já impressos).
"""
from __future__ import annotations

import re
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import pymupdf
from pypdf import PdfReader, PdfWriter

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from caminhos import RELATORIOS  # noqa: E402

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:  # noqa: BLE001
        pass

PAT = re.compile(r"^(.+?) - (\d{4}-\d{2}-\d{2}) - (.+)\.pdf$", re.IGNORECASE)
OUT_DIR = RELATORIOS / "Imprimir"
COR_DATA = (0.12, 0.31, 0.47)


def profissional_de(rest: str) -> str:
    """Remove sufixo (- CASU, - Bayley, - Pauta …); fica só o profissional."""
    return rest.split(" - ", 1)[0].strip()


def data_br(iso: str) -> str:
    d = datetime.strptime(iso, "%Y-%m-%d")
    return d.strftime("%d/%m/%Y")


def manifest_path(nome_pdf: str) -> Path:
    return OUT_DIR / f"{Path(nome_pdf).stem}.manifest.txt"


def ler_manifest(nome_pdf: str) -> set[str]:
    mf = manifest_path(nome_pdf)
    if not mf.exists():
        return set()
    return {ln.strip() for ln in mf.read_text(encoding="utf-8").splitlines() if ln.strip()}


def gravar_manifest(nome_pdf: str, nomes: list[str]) -> None:
    mf = manifest_path(nome_pdf)
    mf.parent.mkdir(parents=True, exist_ok=True)
    mf.write_text("\n".join(sorted(nomes)) + "\n", encoding="utf-8")


def coletar() -> dict[tuple[str, str], list[tuple[str, Path]]]:
    groups: dict[tuple[str, str], list[tuple[str, Path]]] = defaultdict(list)
    for p in sorted(RELATORIOS.glob("*.pdf"), key=lambda x: x.name.casefold()):
        if p.name.casefold().startswith("imprimir"):
            continue
        m = PAT.match(p.name)
        if not m:
            print("UNPARSED", p.name)
            continue
        tipo, data, rest = m.group(1), m.group(2), m.group(3)
        prof = profissional_de(rest)
        groups[(tipo, prof)].append((data, p))
    for key in groups:
        groups[key].sort(key=lambda x: x[0])
    return groups


def pagina_em_branco_pypdf(page) -> bool:
    res = page.get("/Resources")
    if res is not None:
        res = res.get_object() if hasattr(res, "get_object") else res
        xobj = res.get("/XObject") if res else None
        if xobj is not None:
            xobj = xobj.get_object() if hasattr(xobj, "get_object") else xobj
            for _name, obj in xobj.items():
                o = obj.get_object() if hasattr(obj, "get_object") else obj
                if str(o.get("/Subtype", "")) == "/Image":
                    return False
    text = (page.extract_text() or "").strip()
    if len(text) >= 20:
        return False
    contents = page.get_contents()
    raw = b""
    if contents is not None:
        if isinstance(contents, list):
            for c in contents:
                raw += c.get_data() if hasattr(c, "get_data") else b""
        elif hasattr(contents, "get_data"):
            raw = contents.get_data()
    return len(raw) < 800


def pagina_em_branco_pymupdf(page: pymupdf.Page) -> bool:
    text = (page.get_text() or "").strip()
    if len(text) >= 20:
        return False
    if page.get_images():
        return False
    return len(page.read_contents() or b"") < 800


def paginas_natacao(doc: pymupdf.Document) -> list[int]:
    """Pula capa (p1) e frequência (p2) quando o PDF segue o layout Acqua Kids."""
    n = doc.page_count
    if n <= 2:
        return list(range(n))
    t1 = (doc[0].get_text() or "").upper()
    t2 = (doc[1].get_text() or "").upper()
    capa = any(
        k in t1
        for k in (
            "CONNECT ACQUA",
            "EQUIPE MULTIDISCIPLINAR",
            "GESTÃO DE DESENVOLVIMENTO /",
        )
    ) or len(t1.strip()) < 5
    freq = "FREQUÊNCIA" in t2 or "FREQUENCIA" in t2
    if capa and freq:
        return list(range(2, n))
    return list(range(n))


def adicionar_data_topo(page: pymupdf.Page, rotulo: str) -> None:
    rect = page.rect
    faixa = pymupdf.Rect(0, 0, rect.width, 22)
    page.draw_rect(faixa, color=(1, 1, 1), fill=(1, 1, 1), overlay=True)
    page.insert_text(
        (36, 16),
        rotulo,
        fontsize=10,
        fontname="helv",
        color=COR_DATA,
    )


def juntar_natacao(paths: list[Path], dest: Path) -> tuple[int, int]:
    """Compacta natação: sem capa/frequência; data no topo de cada folha."""
    out = pymupdf.open()
    pages = 0
    skipped = 0
    for path in paths:
        m = PAT.match(path.name)
        iso = m.group(2) if m else ""
        rotulo = f"Natação — {data_br(iso)}" if iso else "Natação"
        doc = pymupdf.open(path)
        for i in paginas_natacao(doc):
            src = doc[i]
            if pagina_em_branco_pymupdf(src):
                skipped += 1
                print(f"  SKIP_BLANK {path.name} p{i + 1}")
                continue
            pg = out.new_page(width=src.rect.width, height=src.rect.height)
            pg.show_pdf_page(pg.rect, doc, i)
            adicionar_data_topo(pg, rotulo)
            pages += 1
        doc.close()
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp = dest.with_suffix(".tmp.pdf")
    out.save(tmp)
    out.close()
    try:
        if dest.exists():
            dest.unlink()
        tmp.replace(dest)
    except OSError:
        from caminhos import TEMP  # noqa: E402

        fallback = TEMP / dest.name
        if fallback.exists():
            fallback.unlink()
        tmp.replace(fallback)
        print(f"AVISO: feche {dest.name} se estiver aberto.")
        print(f"SALVO em {fallback}")
        return pages, skipped


def juntar(paths: list[Path], dest: Path) -> tuple[int, int]:
    """Junta PDFs sem transformação (padrão)."""
    writer = PdfWriter()
    pages = 0
    skipped = 0
    for path in paths:
        reader = PdfReader(str(path))
        if reader.is_encrypted:
            try:
                reader.decrypt("")
            except Exception as exc:  # noqa: BLE001
                print("SKIP_ENCRYPTED", path.name, exc)
                continue
        for i, page in enumerate(reader.pages):
            if pagina_em_branco_pypdf(page):
                skipped += 1
                print(f"  SKIP_BLANK {path.name} p{i + 1}")
                continue
            writer.add_page(page)
            pages += 1
    dest.parent.mkdir(parents=True, exist_ok=True)
    with dest.open("wb") as f:
        writer.write(f)
    return pages, skipped


def nome_saida(tipo: str, prof: str) -> str:
    if tipo.casefold() == "alta":
        return "Imprimir - Alta.pdf"
    return f"Imprimir - {tipo} - {prof}.pdf"


def resolver_nome_e_paths(
    tipo: str,
    profissional: str | None,
    groups: dict[tuple[str, str], list[tuple[str, Path]]],
    so_novos: bool = False,
) -> tuple[str, list[Path]]:
    items: list[tuple[str, Path]] = []
    for (t, prof), files in groups.items():
        if t.casefold() != tipo.casefold():
            continue
        if profissional is not None and profissional.casefold() not in prof.casefold():
            continue
        items.extend(files)
    if not items:
        raise SystemExit(f"Nenhum PDF para tipo={tipo!r} prof={profissional!r}")

    items.sort(key=lambda x: x[0])
    paths = [p for _, p in items]

    if tipo.casefold() == "alta":
        return "Imprimir - Alta.pdf", paths

    profs = {profissional_de(PAT.match(p.name).group(3)) for p in paths}
    if len(profs) > 1 and profissional is None:
        raise SystemExit(f"Vários profissionais em {tipo}: {sorted(profs)}")
    prof_nome = next(iter(profs)) if len(profs) == 1 else profissional_de(
        PAT.match(paths[0].name).group(3)
    )
    nome = f"Imprimir - {tipo} - {prof_nome}.pdf"

    if so_novos:
        manifest = ler_manifest(nome)
        if not manifest:
            if tipo.casefold() == "fisioterapia":
                # Pacote anterior já impresso; estes scans entraram depois
                ja_impressos = {
                    "Fisioterapia - 2021-11-29 - Silvia Figueiredo.pdf",
                    "Fisioterapia - 2022-02-10 - Silvia Figueiredo.pdf",
                }
                manifest = {p.name for p in paths if p.name not in ja_impressos}
                gravar_manifest(nome, sorted(manifest))
                print(f"MANIFEST criado ({len(manifest)} já impressos)")
            else:
                print("MANIFEST vazio — tratando todos como novos")
        paths = [p for p in paths if p.name not in manifest]
        if not paths:
            raise SystemExit(f"Nenhum PDF novo para {nome} (todos já no manifest)")

    return nome, paths


def atualizar_pacote(
    tipo: str,
    profissional: str | None = None,
    *,
    so_novos: bool = False,
) -> None:
    groups = coletar()
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    nome, paths = resolver_nome_e_paths(tipo, profissional, groups, so_novos=so_novos)
    dest = OUT_DIR / nome

    if tipo.casefold() == "natação" or tipo.casefold() == "natacao":
        pages, skipped = juntar_natacao(paths, dest)
    else:
        pages, skipped = juntar(paths, dest)

    extra = f" (pulou {skipped} em branco)" if skipped else ""
    modo = " só novos" if so_novos else ""
    print(f"OK {len(paths):2} arq / {pages:3} pag{modo}  ->  {nome}{extra}")
    print(f"ATUALIZADO {dest}")


def main() -> None:
    import argparse

    ap = argparse.ArgumentParser(description="Junta PDFs de Relatórios para Imprimir/")
    ap.add_argument(
        "--apenas",
        nargs=2,
        metavar=("TIPO", "PROFISSIONAL"),
        help="Atualiza só este pacote (ex.: Fisioterapia Silvia). Não apaga os demais.",
    )
    ap.add_argument(
        "--so-novos",
        action="store_true",
        help="Só PDFs ainda não listados no manifest (já impressos).",
    )
    ap.add_argument(
        "--marcar-impresso",
        action="store_true",
        help="Após gerar, adiciona os PDFs incluídos ao manifest (registrar impressão).",
    )
    ap.add_argument(
        "--tudo",
        action="store_true",
        help="Recria todos os pacotes (apaga Imprimir - *.pdf existentes).",
    )
    args = ap.parse_args()

    if args.apenas:
        atualizar_pacote(args.apenas[0], args.apenas[1], so_novos=args.so_novos)
        if args.marcar_impresso:
            groups = coletar()
            nome, paths = resolver_nome_e_paths(
                args.apenas[0], args.apenas[1], groups, so_novos=False
            )
            _, paths_novos = resolver_nome_e_paths(
                args.apenas[0], args.apenas[1], groups, so_novos=True
            )
            manifest = ler_manifest(nome)
            manifest.update(p.name for p in paths_novos)
            gravar_manifest(nome, sorted(manifest))
            print(f"MANIFEST atualizado ({len(manifest)} impressos)")
        return

    if not args.tudo:
        print("Use --apenas TIPO PROFISSIONAL  ou  --tudo")
        print("Ex.: py -3 juntar_relatorios.py --apenas Fisioterapia Silvia --so-novos")
        return

    groups = coletar()
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    for old in OUT_DIR.glob("Imprimir - *.pdf"):
        old.unlink()

    packs: dict[str, list[Path]] = defaultdict(list)
    for (tipo, prof), items in groups.items():
        nome = nome_saida(tipo, prof)
        packs[nome].extend(p for _, p in items)

    def data_key(path: Path) -> str:
        m = PAT.match(path.name)
        return m.group(2) if m else path.name

    total_src = 0
    total_out = 0
    for nome in sorted(packs.keys(), key=lambda s: s.casefold()):
        paths = sorted(packs[nome], key=data_key)
        seen: set[Path] = set()
        uniq: list[Path] = []
        for p in paths:
            if p in seen:
                continue
            seen.add(p)
            uniq.append(p)
        paths = uniq
        total_src += len(paths)
        dest = OUT_DIR / nome
        tipo = PAT.match(paths[0].name).group(1) if paths else ""
        if tipo.casefold() in ("natação", "natacao"):
            pages, skipped = juntar_natacao(paths, dest)
        else:
            pages, skipped = juntar(paths, dest)
        total_out += 1
        extra = f" (pulou {skipped} em branco)" if skipped else ""
        print(f"OK {len(paths):2} arq / {pages:3} pag  ->  {nome}{extra}")

    print(f"ORIGINAIS {total_src}  ->  ARQUIVOS_IMPRIMIR {total_out}  em  {OUT_DIR}")


if __name__ == "__main__":
    main()
