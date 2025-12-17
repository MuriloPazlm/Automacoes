# ADDPAINELv7 — Automação para Adição de TCLs e Filtragem por BRICK

Este script em Python automatiza a preparação de uma planilha de **Adição de TCLs** a partir de uma base Excel, **separando e validando códigos BRICK**, aplicando um **VLOOKUP automatizado**, gerando uma aba filtrada (**ADICAO**) apenas com registros relevantes, e **salvando o arquivo em uma estrutura de pastas por ciclo de marketing**, com nome de arquivo único e incremental.

> **Tecnologias:** Python 3.x • openpyxl • `re`, `datetime`, `os`, `shutil`

---

## ✨ Principais Funcionalidades

- **Entrada interativa:**
  - Solicita o **Setor do Representante** (texto livre).
  - Solicita a lista de **códigos BRICK** (aceita separados por **espaço, vírgula ou quebra de linha**).

- **Normalização de BRICKs:**
  - Formata automaticamente cada código para o padrão `BR_XXXXXXX` (prefixo `BR_` + **7 dígitos** com `zfill`).
  - Ignora entradas vazias e espaços extras.

- **Preparação da base (Excel / openpyxl):**
  - Carrega a planilha base (`BASE_ADD.xlsx`) e usa a **aba ativa**.
  - **Separa BRICKs** quando há múltiplos códigos na mesma célula (divide por espaço e espalha em colunas adicionais).
  - Calcula **quantas colunas** são necessárias para acomodar todos os BRICKs separados por linha.

- **VLOOKUP automatizado:**
  - Insere **duas colunas auxiliares** na direita do bloco de BRICKs.
  - Preenche a primeira coluna auxiliar com a lista de BRICKs formatados.
  - Na segunda, escreve uma **fórmula de VLOOKUP** (com `IFERROR`) que faz a validação/consulta dos BRICKs por linha.
  - Obs.: A fórmula é escrita como texto na célula, pronta para cálculo no Excel.

- **Criação da aba `ADICAO`:**
  - Gera/limpa a aba `ADICAO` e escreve um cabeçalho padronizado:
    - `Ciclo de Marketing`, `Alvo: Território`, `Account ID_18`, `Nome da conta`, `Specialty 1`, `Contact ID_18`, `Licença Médica Legal`.
  - **Filtra linhas** da base original onde **qualquer** coluna de BRICK (após separação) aparece nos BRICKs informados.
  - **Elimina duplicidades por `Account ID_18`**, mantendo apenas a primeira ocorrência.

- **Salvar com nome único e organizar por ciclo:**
  - Salva o arquivo como `ADD TCLs- <Setor>.xlsx` (ex.: `ADD TCLs- Sul.xlsx`).
  - Se já existir, cria versão incremental: `ADD TCLs- <Setor>_v2.xlsx`, `..._v3.xlsx`, etc.
  - Detecta o **ciclo de marketing** pelo timestamp de criação do arquivo e **move** para a pasta do ciclo correspondente:
    - `CICLO 07 (2025-07-18 a 2025-08-15)`
    - `CICLO 08 (2025-08-18 a 2025-09-15)`
    - `CICLO 09 (2025-09-16 a 2025-10-14)`
    - `CICLO 10 (2025-10-15 a 2025-11-12)`
    - `CICLO 11 (2025-11-13 a 2025-12-17)`

---

## 📂 Estrutura de Pastas e Arquivos

- **`arquivo_origem`**: `C:\\Users\\pazlimx1\\OneDrive - Abbott\\Documents\\AUTOMACAO\\ADD TCL\\BASE\\BASE_ADD.xlsx`
- **`pasta_base` (saída)**: `C:\\Users\\pazlimx1\\OneDrive - Abbott\\Documents\\AUTOMACAO\\ADICAO PAINEL`
  - `ADD TCLs- <Setor>.xlsx` ou `ADD TCLs- <Setor>_vN.xlsx`
  - `CICLO XX\\ADD TCLs- <Setor>.xlsx` (arquivo movido para a subpasta do ciclo)

> Ajuste esses caminhos nas constantes do script se necessário.

---

## 🔧 Requisitos

- **Python 3.x**
- **openpyxl** (leitura/escrita de arquivos Excel `.xlsx`)
- Acesso de escrita/leitura aos caminhos configurados.

Instalação (se necessário):
```bash
pip install openpyxl
```

---

## ▶️ Como Usar

1. Garanta que o arquivo **`BASE_ADD.xlsx`** está no caminho configurado e que a aba ativa contém:
   - Cabeçalhos com os nomes esperados em português (p.ex. `Account ID_18`, `Ciclo de Marketing`, etc.).
   - Coluna **G** (índice 7) contendo os BRICKs (podem estar múltiplos por célula).

2. Execute o script:
```bash
python ADDPAINELv7.py
```

3. Informe:
   - **Setor do Representante** (ex.: `Sul`)
   - **Lista de BRICKs** (ex.: `123, 456 789` ou em linhas diferentes)

4. Ao finalizar:
   - O script salvará o arquivo nomeado em `pasta_base`, criará versão se já existir, e **moverá** para a subpasta do **ciclo** correspondente conforme a data de criação do arquivo.

---

## 🧠 Como o script funciona (fluxo)

1. **Configura ciclos** (datas início/fim) e converte para `datetime`.
2. **Coleta entradas** do usuário e normaliza BRICKs (`BR_` + 7 dígitos).
3. **Carrega a base** via `openpyxl` e identifica a coluna de BRICK (fixa: **7**).
4. **Separa BRICKs** por espaço em colunas novas (quantidade dinâmica).
5. **Insere colunas auxiliares** e escreve fórmula de VLOOKUP com `IFERROR`.
6. **Cria/limpa a aba `ADICAO`**, mapeia índices das colunas de interesse pelo cabeçalho, filtra linhas por presença de BRICK e remove duplicidades de `Account ID_18`.
7. **Salva com nome único**, determina o ciclo pela data de criação e **move** o arquivo para a pasta do ciclo.
8. **Mensagens de erro amigáveis** para casos de arquivo aberto ou permissões.

---

## 📎 Cabeçalhos esperados na base

O script busca estes nomes de coluna (sensíveis a grafia):
- `Ciclo de Marketing`
- `Alvo: Território`
- `Account ID_18`
- `Nome da conta`
- `Specialty 1`
- `Contact ID_18`
- `Licença Médica Legal`

> Se a base usar nomes diferentes, atualize o dicionário `cabecalhos` na função `criar_aba_adicao`.

---

## ⚠️ Limitações e Observações

- A **fórmula de VLOOKUP** escrita nas células assume que o Excel calculará após abrir o arquivo (o script não avalia fórmulas).
- O **separador de BRICK** é **espaço** na célula; se houver vírgulas/pontos e vírgulas dentro da planilha base, ajuste a função `separar_bricks`.
- O índice da coluna de BRICK está **fixo em 7** (`colunaBrick = 7`); altere se a estrutura da base mudar.
- O **ciclo** é determinado pela **data de criação** do arquivo salvo; se precisar usar outra referência (p.ex. data de sistema), adapte `salvar_e_mover_arquivo`.
- Os caminhos são **Windows + OneDrive**; em outros ambientes, atualize `pasta_base` e `arquivo_origem`.

---

## 🗺️ Roadmap (idéias de evolução)

- Parametrizar `colunaBrick` e nomes de cabeçalhos via arquivo `.ini` ou `.yaml`.
- Suportar separadores múltiplos na base (`;`, `,`) além de espaço.
- Validar BRICKs por **regex** (apenas dígitos) antes de formatar.
- Exportar a aba `ADICAO` como arquivo separado (ex.: `ADICAO_<Setor>.xlsx`).
- Log estruturado (arquivo `.log`) com contagem de linhas filtradas e tempo de execução.
- Testes unitários com `pytest` para `formatar_bricks`, `separar_bricks` e `gerar_nome_unico`.

---

## 📄 Licença

Defina uma licença (ex.: MIT) conforme sua necessidade.

---

## 👤 Autor

Murilo Paz Lima — Automação de suporte administrativo (São Paulo, SP)
