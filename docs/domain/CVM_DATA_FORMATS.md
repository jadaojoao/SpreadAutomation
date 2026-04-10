# Formatos de Dados CVM

> Referencia dos formatos de dados da Comissao de Valores Mobiliarios usados como fonte pelo sistema.

---

## 1. DadosDocumento.xlsx (Arquivo Individual)

Arquivo Excel gerado automaticamente pela CVM. Estrutura padronizada para todas as empresas.

### Abas disponiveis

| Aba | Usada | Funcao |
|-----|-------|--------|
| Composicao Capital | Nao | Composicao acionaria |
| DF Ind Ativo | Sim (individual) | BP Individual — Ativo |
| DF Ind Passivo | Sim (individual) | BP Individual — Passivo |
| DF Ind Resultado Periodo | Sim (individual) | DRE Individual |
| DF Ind Resultado Abrangente | Nao | Resultado abrangente |
| DF Ind Fluxo de Caixa | Sim (individual) | DFC Individual |
| DF Ind Valor Adicionado | Nao | DVA Individual |
| DF Ind DMPL Ultimo | Sim (individual, anual) | DMPL ultimo exercicio |
| DF Ind DMPL Penultimo | Nao | DMPL penultimo exercicio |
| DF Ind DMPL Antepenultimo | Nao | DMPL antepenultimo exercicio |
| DF Cons Ativo | Sim (consolidado) | BP Consolidado — Ativo |
| DF Cons Passivo | Sim (consolidado) | BP Consolidado — Passivo |
| DF Cons Resultado Periodo | Sim (consolidado) | DRE Consolidado |
| DF Cons Resultado Abrangente | Nao | Resultado abrangente consolidado |
| DF Cons Fluxo de Caixa | Sim (consolidado) | DFC Consolidado |
| DF Cons Valor Adicionado | Nao | DVA Consolidado |
| DF Cons DMPL Ultimo | Sim (consolidado, anual) | DMPL Consolidado ultimo exercicio |
| DF Cons DMPL Penultimo | Nao | DMPL Consolidado penultimo |
| DF Cons DMPL Antepenultimo | Nao | DMPL Consolidado antepenultimo |

### Colunas padrao — BP, DRE, DFC (DFP anual)

| Coluna | Conteudo |
|--------|----------|
| Codigo Conta | Codigo contabil CVM hierarquico (ex.: `1.01`, `3.01`) |
| Descricao Conta | Nome da conta (ex.: "Receita de Venda de Bens e/ou Servicos") |
| Valor Ultimo Exercicio | Valor do periodo mais recente (ex.: 2024) |
| Valor Penultimo Exercicio | Valor do periodo anterior (ex.: 2023) |
| Valor Antepenultimo Exercicio | Valor de dois periodos atras (ex.: 2022) |

### Colunas padrao — BP, DRE, DFC (ITR trimestral)

| Coluna | Conteudo |
|--------|----------|
| Codigo Conta | Igual ao DFP |
| Descricao Conta | Igual ao DFP |
| Valor Trimestre Atual | Valor do trimestre (BP) |
| Valor Exercicio Anterior | Valor do exercicio anterior (BP) |
| Valor Acumulado Atual Exercicio | Valor acumulado YTD (DRE/DFC) |
| Valor Acumulado Exercicio Anterior | Valor acumulado do ano anterior (DRE/DFC) |

**Nota:** O pipeline usa as colunas "Acumulado" para DRE/DFC trimestral, nao as colunas "Trimestre".

### Colunas padrao — DMPL

| Coluna | Conteudo |
|--------|----------|
| CodigoConta | Codigo contabil (SEM espaco — diferente das outras abas) |
| DescricaoConta | Nome da conta (SEM espaco) |
| Capital Social Integralizado | Coluna de capital |
| Patrimonio Liquido | Coluna usada para Individual |
| Patrimonio liquido Consolidado | Coluna usada para Consolidado |

### Renomeacao de abas pelo sistema

O `ingestion/cvm_excel.py` (atual `processing/origin.py`) renomeia as abas:

```
DF Cons Ativo              → "cons ativos"
DF Cons Passivo            → "cons passivos"
DF Cons Resultado Periodo  → "cons DRE"
DF Cons Fluxo de Caixa     → "cons DFC"
DF Cons DMPL Ultimo        → "cons DMPL"    (anual)
DF Cons DMPL Atual         → "cons DMPL"    (trimestral — aba diferente)
```

E renomeia as colunas de valor para o nome do periodo:
```
"Valor Ultimo Exercicio"         → "2024"
"Valor Penultimo Exercicio"      → "2023"
"Valor Antepenultimo Exercicio"  → "2022"
```

---

## 2. CSVs Dados Abertos CVM (Portal Bulk)

Disponiveis em `dados.cvm.gov.br`. Contem dados de TODAS as empresas em formato tabular.

### Campos relevantes

| Campo CSV | Equivalente no xlsx | Tipo | Observacao |
|-----------|---------------------|------|-----------|
| CD_CONTA | Codigo Conta | str | Hierarquico (1.01, 3.04.02) |
| DS_CONTA | Descricao Conta | str | Textual |
| VL_CONTA | Valor Ultimo Exercicio etc. | decimal(29,10) | Valor numerico |
| ORDEM_EXERC | (coluna de periodo) | str | "ULTIMO" / "PENULTIMO" / "ANTEPENULTIMO" |
| ST_CONTA_FIXA | — | str | "S" = conta fixa (sempre presente); "N" = discricionaria |
| ESCALA_MOEDA | — | str | "Mil" ou "Unidade" |
| GRUPO_DFP | (nome da aba) | str | Grupo/agregacao da demonstracao |
| VERSAO | — | int | Incrementa em reapresentacoes |
| COLUNA_DF | (colunas DMPL) | str | Apenas DMPL: identifica coluna de patrimonio |

### Campos criticos

#### ST_CONTA_FIXA
- `S` = conta fixa: sempre presente em todas as empresas. Segura para matching Layer 1.
- `N` = conta discricionaria: empresa decide se reporta. Pode nao existir em outra empresa.

O Plano de Contas Fixas DFP (`docs/reference/plano-contas-fixas-DFP/`) lista todas as contas com `S`.

#### ESCALA_MOEDA
O pipeline assume escala consistente entre periodos. Se uma empresa mudar de "Unidade" para "Mil" entre periodos, o matching por valor falhara silenciosamente.

**Mitigacao planejada:** validar na ingestion e alertar o usuario se ESCALA_MOEDA divergir.

#### VERSAO
Incrementa quando a empresa reapresenta demonstracoes. O pipeline deve sempre usar a versao mais recente.

**Mitigacao planejada:** verificar e filtrar por max(VERSAO) na ingestion.

---

## 3. Estrutura da DFC (atencao especial)

A DFC tem apenas **10 contas fixas** CVM, todas de alto nivel:

| CD_CONTA | Descricao |
|----------|-----------|
| 6.01 | Caixa Liquido Atividades Operacionais |
| 6.01.01 | Caixa Gerado nas Operacoes |
| 6.01.02 | Variacoes nos Ativos e Passivos |
| 6.02 | Caixa Liquido Atividades de Investimento |
| 6.03 | Caixa Liquido Atividades de Financiamento |
| 6.04 | Variacao Cambial |
| 6.05 | Aumento (Reducao) de Caixa |
| 6.05.01 | Saldo Inicial de Caixa |
| 6.05.02 | Saldo Final de Caixa |

Sub-itens como "Depreciacao e Amortizacao" NAO sao contas fixas — cada empresa cria seus proprios sub-niveis. Por isso a DFC usa busca textual (regex) em vez de Layer 1.

**Risco critico documentado (ADR-H03):** `6.03.02 "Amortizacoes"` (pagamento de emprestimo) pode ser confundida com depreciacao se o filtro por secao 6.01 nao estiver ativo.

---

## 4. Estrutura da DMPL (atencao especial)

A DMPL tem estrutura diferente das outras demonstracoes:

- Colunas sao `CodigoConta` e `DescricaoConta` (sem espaco)
- Campo `COLUNA_DF` identifica qual coluna de patrimonio o valor pertence
- No xlsx, essas viram colunas separadas: "Capital Social Integralizado", "Patrimonio Liquido", etc.

O sistema usa regex para encontrar a coluna correta:
1. Busca por `patrim.*consolidado` (para consolidado)
2. Fallback: busca por `patrim.*liquido` (para individual)

---

## 5. Logica de Periodos

| Input | Atual | Anterior | Anterior 2 | Trimestral? |
|-------|-------|----------|------------|-------------|
| "2024" | "2024" | "2023" | "2022" | Nao |
| "2023" | "2023" | "2022" | "2021" | Nao |
| "4T24" | "4T24" | "4T23" | "4T22" | Sim |
| "1T25" | "1T25" | "1T24" | "1T23" | Sim |

A flag `is_trim` afeta:
- Quais colunas sao lidas na Origem (nomes de cabecalho diferentes)
- Se a DRE manual e aplicada (apenas trimestral)
- Qual aba de DMPL e selecionada ("Ultimo" para anual, "Atual" para trimestral)

---

## 6. Material de Referencia no Repositorio

| Arquivo | Conteudo | Uso |
|---------|----------|-----|
| `docs/reference/plano-contas-fixas-DFP/` | Plano de Contas Fixas DFP — ENET 10.0 | Fonte primaria para Layer 1 (quais CD_CONTA sao fixos) |
| `docs/reference/meta_dfp_cia_aberta_txt/` | Schemas dos campos dos CSVs abertos CVM | Referencia dos campos disponiveis nos CSVs |
| `docs/reference/cvm_account_dictionary.csv` | 17.068 DS_CONTA unicos por statement_type | Base para account_synonyms.json (Layer 3 fuzzy) |
| `docs/reference/Manual XML ITR v2.xls` | Leiaute do Formulario DFP/ITR — ENET | Referencia da estrutura XML dos formularios |
| `docs/reference/manual-de-envio-*.pdf` | Manual oficial CVM | Regras de envio, prazos, reapresentacoes |

---

## Historico

| Data | Mudanca |
|------|---------|
| 2026-04-09 | Criacao com DadosDocumento.xlsx, CSVs CVM, DFC, DMPL, periodos, e material de referencia |
