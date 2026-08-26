# -*- coding: utf-8 -*-
"""Junta PDFs de Relatórios por tipo + profissional → pasta Imprimir/.

Originais intactos. Saída: Imprimir - {Tipo} - {Profissional}.pdf
Exceção: Alta → um único Imprimir - Alta.pdf (todos os profissionais, por data).
"""
from __future__ import annotations

import re
import sys
from collections import defaultdict
from pathlib import Path

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


def profissional_de(rest: str) -> str:
    """Remove sufixo (- CASU, - Bayley, - Pauta …); fica só o profissional."""
    return rest.split(" - ", 1)[0].strip()


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
        groups[key].sort(key=lambda x: x[0])  # por data
    return groups


def pagina_em_branco(page) -> bool:
    """Página sem imagem e sem texto útil (ex.: folha em branco no fim do PDF)."""
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


def juntar(paths: list[Path], dest: Path) -> tuple[int, int]:
    """Retorna (páginas escritas, páginas em branco puladas)."""
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
            if pagina_em_branco(page):
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
    """Alta: um único pacote. Demais: tipo + profissional."""
    if tipo.casefold() == "alta":
        return "Imprimir - Alta.pdf"
    return f"Imprimir - {tipo} - {prof}.pdf"


def atualizar_pacote(tipo: str, profissional: str | None = None) -> None:
    """Atualiza só um pacote Imprimir, sem recriar/apagar os demais."""
    groups = coletar()
    OUT_DIR.mkdir(parents=True, exist_ok=True)

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
        nome = "Imprimir - Alta.pdf"
    else:
        # usa o profissional do primeiro arquivo do grupo
        first_prof = profissional_de(PAT.match(paths[0].name).group(3))
        # se filtro parcial (ex. Silvia), preferir nome completo do arquivo
        nome = f"Imprimir - {tipo} - {first_prof}.pdf"
        # se vários profissionais no mesmo tipo (não Alta), exigir profissional
        profs = {profissional_de(PAT.match(p.name).group(3)) for p in paths}
        if len(profs) > 1 and profissional is None:
            raise SystemExit(f"Vários profissionais em {tipo}: {sorted(profs)}")
        if len(profs) == 1:
            nome = f"Imprimir - {tipo} - {next(iter(profs))}.pdf"

    dest = OUT_DIR / nome
    pages, skipped = juntar(paths, dest)
    extra = f" (pulou {skipped} em branco)" if skipped else ""
    print(f"OK {len(paths):2} arq / {pages:3} pag  ->  {nome}{extra}")
    print(f"ATUALIZADO {dest} (outros Imprimir preservados)")


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
        "--tudo",
        action="store_true",
        help="Recria todos os pacotes (apaga Imprimir - *.pdf existentes).",
    )
    args = ap.parse_args()

    if args.apenas:
        atualizar_pacote(args.apenas[0], args.apenas[1])
        return

    if not args.tudo:
        # padrão seguro: não apagar pacotes já impressos/removidos
        print("Use --apenas TIPO PROFISSIONAL  ou  --tudo")
        print("Ex.: py -3 juntar_relatorios.py --apenas Fisioterapia Silvia")
        return

    groups = coletar()
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    # Remove versões antigas desta pasta (só Imprimir - *.pdf)
    for old in OUT_DIR.glob("Imprimir - *.pdf"):
        old.unlink()

    # Alta: juntar todos os profissionais num único arquivo (por data)
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
        pages, skipped = juntar(paths, dest)
        total_out += 1
        extra = f" (pulou {skipped} em branco)" if skipped else ""
        print(f"OK {len(paths):2} arq / {pages:3} pag  ->  {nome}{extra}")

    print(f"ORIGINAIS {total_src}  ->  ARQUIVOS_IMPRIMIR {total_out}  em  {OUT_DIR}")


if __name__ == "__main__":
    main()
