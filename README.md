# Automação de Painel — ADD & LIMPEZA (v7)

Este repositório reúne dois scripts em Python que automatizam o fluxo de **adição** e **limpeza** de TCLs em planilhas Excel,
utilizando `openpyxl` e regras de ciclos de marketing para organização automática dos arquivos.

- **ADDPAINELv7.py** — Prepara a planilha para **adição** de TCLs a partir de uma base, valida e separa BRICKs, cria a aba `ADICAO`
  com registros filtrados e salva o arquivo em pastas por **ciclo de marketing**.
- **limpezaPainelv7.py** — Realiza **limpeza**/remoção de TCLs não pertencentes à lista de BRICKs informada, gerando a aba `Limpeza`
  resumida e salvando em pastas por **ciclo de marketing**.

> **Tecnologias:** Python 3.x • openpyxl • `re`, `datetime`, `os`, `shutil`

---

##  Visão Geral

### ADDPAINELv7.py
- **Entrada interativa:** Setor do Representante + lista de BRICKs (espaço, vírgula ou quebra de linha).
- **Normalização:** Formata para `BR_XXXXXXX` (prefixo `BR_` + 7 dígitos), ignorando vazios.
- **Base Excel:** Carrega `BASE_ADD.xlsx` (aba ativa), **separa BRICKs** em colunas quando houver múltiplos na mesma célula.
- **VLOOKUP:** Insere duas colunas auxiliares (lista de BRICKs e fórmula `IFERROR(VLOOKUP(...))`).
- **Aba `ADICAO`:** Cria/limpa, localiza cabeçalhos, **filtra linhas** onde qualquer BRICK aparece na lista informada e **remove duplicidades por `Account ID_18`**.
- **Salvar/organizar:** Gera nome único (`ADD TCLs- <Setor>.xlsx`, com versões `_v2`, `_v3`...), detecta o ciclo pela data de criação e **move** para a subpasta do ciclo.

### limpezaPainelv7.py
- **Entrada interativa:** Setor do Representante + lista de BRICKs.
- **Separação na base:** Descobre a quantidade máxima de BRICKs por célula (divide por espaço) e **insere colunas** para espalhar os valores.
- **VLOOKUP auxiliar:** Adiciona duas colunas (códigos do usuário + célula para fórmula `IFERROR(VLOOKUP(...))`).
- **Aba `Limpeza`:** Cria/limpa e **coleta linhas que NÃO contêm** BRICKs informados, gerando um resumo com primeiros nomes e ciclo.
- **Salvar/organizar:** Salva como `DELETE TCLs- <Setor>.xlsx` e **move** para a subpasta do ciclo correspondente.

---

##  Fluxo de Trabalho Recomendido

1. **Adicionar (ADDPAINELv7):** Gere a aba `ADICAO` com contas/contatos que DEVEM ser adicionados conforme BRICKs informados.
2. **Limpar (limpezaPainelv7):** Gere a aba `Limpeza` com contas/contatos que NÃO pertencem aos BRICKs informados (candidatos a remoção/ajuste).
3. **Publicar/arquivar:** Cada saída é salva/movida para a pasta do ciclo, mantendo o histórico por período.

---

## Estrutura de Pastas e Arquivos

- **Saída (`pasta_base`)**: caminho configurável nos scripts
  - `ADD TCLs- <Setor>.xlsx` / `ADD TCLs- <Setor>_vN.xlsx`
  - `DELETE TCLs- <Setor>.xlsx`
  - `CICLO XX/` (subpastas de ciclos; os arquivos são movidos para cá após a criação)

> No `ADDPAINELv7.py` os caminhos padrão usam Windows + OneDrive; ajuste para o seu ambiente.

---

## Requisitos

- **Python 3.x**
- **openpyxl** (leitura/escrita `.xlsx`)

Instalação:
```bash
pip install openpyxl
```

---

## Como Usar

### 1) ADDPAINELv7.py
```bash
python ADDPAINELv7.py
```
Responda:
- **Setor do Representante** (ex.: `Sul`)
- **Lista de BRICKs** (ex.: `123, 456 789` ou cada um em uma linha)

Saída:
- Arquivo `ADD TCLs- <Setor>.xlsx` com aba `ADICAO` filtrada e sem duplicidades de `Account ID_18`.
- Movido para a subpasta do ciclo detectado.

### 2) limpezaPainelv7.py
Antes, configure `pasta_base` e `arquivo_origem` no topo do arquivo.
```bash
python limpezaPainelv7.py
```
Responda:
- **Setor do Representante**
- **Lista de BRICKs**

Saída:
- Arquivo `DELETE TCLs- <Setor>.xlsx` com aba `Limpeza` contendo itens fora dos BRICKs.
- Movido para a subpasta do ciclo detectado.

---

## Cabeçalhos esperados (na base Excel)

- `Ciclo de Marketing`
- `Alvo: Território` / `Alvo: Alvos`
- `Account ID_18`
- `Nome da conta`
- `Specialty 1`
- `Contact ID_18`
- `Licença Médica Legal`
- `Lista de clientes-alvo: Nome`

> Se a nomenclatura variar, ajuste os dicionários de cabeçalhos nas funções que criam as abas `ADICAO` e `Limpeza`.

---

## Limitações e Observações

- As **fórmulas de VLOOKUP** são escritas como texto; a avaliação ocorre no Excel ao abrir o arquivo.
- A **coluna de BRICK** na base é tratada como **índice 7 (coluna G)**; ajuste `colunaBrick` se necessário.
- Em `limpezaPainelv7.py`, **pasta_base** e **arquivo_origem** estão vazios por padrão — **configure antes de rodar**.
- A detecção de ciclo usa a **data de criação** do arquivo salvo; adapte para outra referência se preciso.
- Separador de BRICK na base: **espaço**; para outros separadores, ajuste as funções de split.

---

## Roadmap

- Parametrização via `.ini`/`.yaml` (caminhos, cabeçalhos, coluna de BRICK).
- Exportar `ADICAO`/`Limpeza` como arquivos separados adicionais.
- Logs detalhados e métricas (linhas filtradas, tempo de execução).
- Testes unitários (`pytest`) para funções de formatação e separação.

---

## Licença

Defina uma licença (ex.: MIT) conforme sua necessidade.

---

## Autor

Murilo Paz Lima — Automação de suporte administrativo (São Paulo, SP)
