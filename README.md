# Cuidados médicos — Cecília

Arquivo organizado dos exames, relatórios e do código que monta os Excel da **Cecília Maria Albergaria Silva** (nasc. 30/06/2020).

Repositório **privado**. São dados de saúde: não tornar público.

## Pastas neste repositório

| Pasta | O que tem |
|---|---|
| `codigo/` | Scripts que geram os resumos e a evolução (Python + macros do Excel). |
| `Exames/` | Laudos (sangue, imagem, audiologia, pezinho, EEG, suor). O CTI laboratorial (ago–set/2020, Hospital BH) fica em `Exames/Exames Laboratorias - CTI/` — **não misturar** com o resto. |
| `Relatórios/` | Consultas, terapias, escola e **sumários de alta** hospitalar. |
| `Solicitação Medicamento SUS/` | Pedido e documentos da solicitação de medicamento. |

Não entram no Git (ficam só no computador): Batizado, Documentos, Livros e Apostilas, CTI - Fotos e Recados.

## Arquivos principais

- `Exames/Resumo Exames.xlsm` — resumo com tabela nativa e macro de altura das linhas
- `Relatórios/Resumo Relatórios.xlsx`
- `Exames/Exames Laboratorias - CTI/Resumo Exames CTI.xlsx`
- `Exames/Evolução Exames.xlsm` — evolução com caixinhas (**arquivo ativo**; habilitar macros)
- `Exames/Exames Laboratorias - CTI/Evolução Exames CTI.xlsx` — evolução do CTI, sem macros

## Como nomear

Padrão (Exames e Relatórios):

```
Tipo - AAAA-MM-DD - Profissional.pdf
```

Exemplos:

- `Sangue - 2026-08-12 - Maria Leticia Gambogi Teixeira.pdf`
- `Imagem - 2026-07-29 - Fernanda de Souza Silva - Punho.pdf`

Regras:

- Data **ISO** (`2026-08-12`), da **coleta** no laudo (não da impressão).
- Profissional como no PDF. Se não houver: `Sem solicitante`.
- Dois arquivos no mesmo tipo/data/profissional: sufixo depois do hífen (`- Fator XIII`, `- Punho`).
- **Não renomear** os PDFs do CTI (`3063463.PDF` …).
- PDF novo ainda sem nome: colocar na raiz da pasta `_Organizado` no computador (essa raiz não vai para o Git).

Tipos de exame: Sangue, Urina, Imagem, Pezinho, Audiologia, EEG, Suor.

Tipos de relatório (exemplos): Alta, Endocrino, Neuro, Fonoaudiologia, Fisioterapia, Escola, Funcional, Natação, Pedagogia, Psicologia, TO.

## PDF novo

1. Ler o laudo.
2. CTI ago–set/2020 no Hospital BH → pasta CTI, **sem** renomear; atualizar só os Excel do CTI.
3. Laboratório / imagem / triagem → `Exames/`.
4. Consulta, terapia, escola, parecer ou **sumário de alta** → `Relatórios/`.
5. Renomear no padrão, uma linha no resumo, e se for numérico incluir na evolução.

## Evolução (fora do CTI)

Detalhes em `codigo/README.md`. Resumo:

```powershell
cd codigo\evolucao
powershell -ExecutionPolicy Bypass -File .\gerar.ps1
```

Fecha o Excel antes. O `.xlsm` vai para `Exames\`.

## Evolução do CTI

```powershell
cd codigo\cti
py -3 catalog_cti.py
py -3 build_cti.py
```

## Dependências

Python 3: `pip install -r codigo/requirements.txt` (`openpyxl`, `xlsxwriter`, `pypdf`).

Excel instalado (macros).
