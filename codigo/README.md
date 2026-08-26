# Código de organização dos exames da Cecília

Esta pasta (`codigo/`) é o que vai para o Git. Os PDFs e Excel ficam em `_Organizado/` (Exames, Relatórios, Documentos). Não versionar laudos.

Paciente: **Cecília Maria Albergaria Silva**, nasc. 30/06/2020.

## Pastas

| Pasta | O que entra |
|---|---|
| `Exames/` | Laudos laboratoriais, imagem, audiologia, pezinho, EEG. **Não** misturar o CTI. |
| `Exames/Exames Laboratorias - CTI/` | Internação ago–set/2020 (Hospital BH). Nomes originais das prescrições. `Todos.pdf` é a juntada. |
| `Relatórios/` | Consultas, terapias, escola e sumários de alta hospitalar. |
| `Relatórios/Imprimir/` | PDFs juntados por tipo+profissional (`Imprimir - …`). Originais intactos. Regenerar: `py -3 codigo/imprimir/juntar_relatorios.py`. |
| `Documentos/` | Cartão SUS, certidão, declaração de nascido vivo, etc. |
| `_Organizado/` (raiz) | Só PDFs **novos**, ainda sem nome. Depois de processar, a raiz fica vazia. |

Resumos: `Exames/Resumo Exames.xlsm`, `Relatórios/Resumo Relatórios.xlsm` (ambos com macro de altura), `Exames/Exames Laboratorias - CTI/Resumo Exames CTI.xlsx`.

Evolução (com caixinhas): `Exames/Evolução Exames.xlsm` — **é o arquivo ativo**. O `.xlsx` intermediário não deve ficar na pasta (OneDrive/Excel travam). Macros: caixinhas, impressão de gráficos (títulos destacados + tabela zebrada dos únicos) e **ajuste de altura** nas abas Dados Completo / Dados Selecionados.

Evolução do CTI: `Exames/Exames Laboratorias - CTI/Evolução Exames CTI.xlsx` (sem macros).

## Como nomear

Padrão (Exames e Relatórios):

```
Tipo - AAAA-MM-DD - Profissional.pdf
```

Exemplos:

- `Sangue - 2026-08-12 - Maria Leticia Gambogi Teixeira.pdf`
- `Imagem - 2026-07-29 - Fernanda de Souza Silva - Punho.pdf`
- `Funcional - 2025-06-27 - ....pdf`

Regras:

- Data **ISO** (`2026-08-12`), extraída da **coleta** no laudo (não da impressão).
- Profissional como no PDF. Se não houver: `Sem solicitante`.
- Se dois arquivos no mesmo tipo/data/profissional, acrescente um sufixo depois de um hífen: `- Fator XIII`, `- Culturas`, `- Punho`.
- **Não renomear** os PDFs do CTI (`3063463.PDF` …). O Excel do CTI aponta o nome original.
- Não misturar sangue de consulta com a pasta do CTI.

Tipos de exame usados: Sangue, Urina, Imagem, Pezinho, Audiologia, EEG, Suor.

Tipos de relatório (exemplos): Alta, Pediatria, Endocrinologia, Neurologia, Genética, Fisioterapia, Funcional, Natação, Pedagogia, Psicologia, TO, Fonoaudiologia, Escola.

## Como classificar um PDF novo

1. Ler o laudo (texto ou, se for foto, o carimbo).
2. É período **CTI ago–set/2020** no Hospital BH? → pasta CTI, **sem** renomear; atualizar só os Excel da pasta CTI.
3. É laudo de laboratório/imagem/triagem? → `Exames/`.
4. É consulta, terapia, escola, parecer ou **sumário de alta**? → `Relatórios/`.
5. Renomear no padrão, mover, **uma linha** no resumo da pasta, e se for numérico incluir na evolução.

O resumo de **Exames** (`TabelaExames`): Data | Tipo | Profissional | Descrição | Arquivo. Arquivo ativo: **`.xlsm`** (macro `AjustarAlturasResumo`). Regenerar: `codigo/resumo/gerar.ps1`.

O resumo de **Relatórios** (`TabelaRelatorios`): Data | Tipo | Profissional | Descrição | Arquivo | Conferido | Impresso. Regenerar: `codigo/resumo/gerar_relatorios.ps1`. No CTI também há coluna **Arquivo**.

Impressão dos resumos: paisagem A4, caber na largura, repetir título nas páginas.

## Evolução (fora do CTI)

Arquivo: `codigo/evolucao/build_evolucao.py`

- Lista `TABLES`: uma tabela por exame (só datas reais).
- Dicionário `GRAF`: gráficos. Se o laudo tem piso e teto, use `_pt(data, valor, piso, teto)` — a barra é a faixa daquela data (verde dentro, vermelho fora). Sem referência: `_pt(data, valor)` — só a linha.
- Qualitativo (urina descritiva, genética, imagens, pezinho, suor em foto): tabela sim, gráfico não.

Abas do `.xlsm`:

- **Escolher** — caixinhas. Precisa **habilitar macros**.
- **Dados Completo** / **Graficos Completo** — sempre tudo.
- **Dados Selecionados** / **Graficos Selecionados** — só o marcado (habilite macros).
- Excel em português: o módulo ThisWorkbook se chama `EstaPastaDeTrabalho`.

Gerar de novo (fechar o Excel antes):

```powershell
cd codigo\evolucao
powershell -ExecutionPolicy Bypass -File .\gerar.ps1
```

Isso chama `build_evolucao.py` e `rebuild_xlsm.ps1` (caixinhas + VBA). O `.xlsm` é copiado para `Exames\`.

## Evolução do CTI

```powershell
cd codigo\cti
py -3 catalog_cti.py
py -3 build_cti.py
```

`catalog_cti.py` confere se `Todos.pdf` tem as mesmas prescrições/páginas que os PDFs individuais.

## Como pedir para a IA incluir exames novos

Coloque os PDFs novos na **raiz** de `_Organizado` (ou na pasta CTI, se for o caso) e peça, por exemplo:

> Processa os PDFs novos na raiz: renomeia no padrão, move para a pasta certa, atualiza o resumo e a evolução. Usa `codigo/README.md`.

A IA deve:

1. Extrair data, médico e exames de cada PDF.
2. Renomear e mover (exceto CTI).
3. Atualizar o Excel de resumo da pasta.
4. Incluir valores em `TABLES` e `GRAF` em `build_evolucao.py` (ou nos scripts do CTI) e rodar `gerar.ps1`.
5. Não misturar CTI com a evolução ambulatorial.
6. Não apagar PDFs sem conferir duplicata.
7. Responder em português.

## Dependências

Python 3: `pip install -r codigo/requirements.txt` (`openpyxl`, `xlsxwriter`, `pypdf`).

Excel instalado (macros). O `rebuild_xlsm.ps1` liga temporariamente `AccessVBOM` no registro do usuário para injetar VBA.

## Git (quando for criar o repositório)

Versionar **`codigo/`**. Os PDFs e Excel com dados de saúde podem ficar de fora (`.gitignore` na raiz do repo) ou num remoto privado. Não commitar `~$*` nem `__pycache__/`.
