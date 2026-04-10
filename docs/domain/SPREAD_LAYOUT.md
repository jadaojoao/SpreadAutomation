# Layout do Spread Proxy

> Referencia da estrutura imutavel do Spread Proxy. Esta estrutura NUNCA muda. Qualquer codigo que interaja com o Spread deve respeitar este layout.

---

## 1. Regra Fundamental

- A estrutura final do Spread Proxy e **fixa e imutavel**
- As linhas sempre estao nas mesmas posicoes
- Independentemente da origem dos dados, a saida final respeita exatamente este layout
- O sistema preenche esta estrutura com os valores corretos

---

## 2. Grid de Colunas

A aba "Entrada de Dados" usa colunas fixas com separadores ocultos entre elas.

| Col | Letra | Indice (0-based) | Oculta | Funcao |
|-----|-------|-------------------|--------|--------|
| 1 | A | 0 | Sim | Separador oculto |
| 2 | B | 1 | Nao | **ROTULOS** (nomes das linhas) |
| 3 | C | 2 | Sim | Separador oculto |
| 4 | D | 3 | Nao | **DADOS ANUAIS** — periodo mais antigo |
| 5 | E | 4 | Sim | Metadados ocultos (data, moeda, cambio) |
| 6 | F | 5 | Nao | **DADOS ANUAIS** |
| 7 | G | 6 | Sim | Separador oculto |
| 8 | H | 7 | Nao | **DADOS ANUAIS** |
| 9 | I | 8 | Sim | Separador oculto |
| 10 | J | 9 | Nao | **DADOS ANUAIS** — periodo mais recente |
| 11 | K | 10 | Sim | Separador oculto |
| 12 | L | 11 | Nao | **BALANCETE / TRIMESTRAL** |

### Colunas de dados
```
Anuais:      D, F, H, J    (indices 3, 5, 7, 9)
Trimestral:  L              (indice 11)
Rotulos:     B              (indice 1)
```

### Colunas ocultas (NUNCA usar como destino)
```
A, C, E, G, I, K    (indices 0, 2, 4, 6, 8, 10)
```

### Progressao de preenchimento
O Spread e preenchido da esquerda para a direita ao longo dos anos:
- Ano 1 (mais antigo) → coluna D
- Ano 2 → coluna F
- Ano 3 → coluna H
- Ano 4 (mais recente) → coluna J
- Balancete trimestral → coluna L (sempre)

---

## 3. Zonas do Spread

### Zona de metadados (linhas 1-26)

| Linha | Coluna B | Conteudo |
|-------|----------|----------|
| 3 | Anos | Data do periodo em cada coluna de dados (ex.: `31/12/2023`) |
| 4 | Numero de meses do Balanco | Meses cobertos (12 = anual, 3/6/9 = trimestral) |
| 5 | Nome do Arquivo: | Nome do arquivo de origem |
| 6 | Auditado por | Nome da auditoria |
| 14 | Analista de Credito: | Nome do analista |
| 22-24 | Cambio | Cambio final e medio |
| 26 | — | Repeticao da data (formula `=IF(E3<>"",E3,"")`) |

### Zona de dados (linhas 27-326)

A coluna B contem os rotulos das linhas. As colunas D/F/H/J/L contem os valores financeiros.

| Linhas (aprox.) | Secao | Origem CVM |
|-----------------|-------|------------|
| 27-130 | **Balanco Patrimonial** (Ativo + Passivo) | `DF Cons/Ind Ativo` + `DF Cons/Ind Passivo` |
| 149-180 | **DRE** (Demonstracao de Resultado) | `DF Cons/Ind Resultado Periodo` |
| 180-200 | **DFC** (Fluxo de Caixa) | `DF Cons/Ind Fluxo de Caixa` |
| 200-220 | **Mutacoes PL / Dividendos** | `DF Cons/Ind DMPL` |

---

## 4. Linhas Especiais

Linhas que recebem tratamento especial pelo sistema (NAO passam pela varredura padrao):

| Linha | Rotulo no Spread | Tratamento | Modulo responsavel |
|-------|------------------|------------|-------------------|
| **199** | Amortizacao Total | Soma negativa de depreciacao + amortizacao da DFC (secao 6.01 apenas) | `processing/dfc.py` → `engine/workflow_*.py` |
| **209** | Dividendos Pagos (Res. + P.L.) (-) | Soma positiva de dividendos/JCP da DMPL | `processing/dmpl.py` → `engine/workflow_*.py` |
| **210** | I.R. no P.L. (-) | Soma negativa de dividendos/JCP da DMPL | `processing/dmpl.py` → `engine/workflow_*.py` |
| **213** | Reavaliacao / (Reversao) no Imobilizado | Aumentos de capital da DMPL | `processing/dmpl.py` → `engine/workflow_*.py` |

Estas linhas estao no `skip_rows` do `spread_schema.json` e sao excluidas da varredura padrao.

---

## 5. spread_schema.json

Representacao formal deste layout como arquivo de configuracao:

```json
{
  "sheet_name": "Entrada de Dados",
  "data_start_row": 27,
  "columns": {
    "labels": "B",
    "annual": ["D", "F", "H", "J"],
    "quarterly": "L",
    "hidden": ["A", "C", "E", "G", "I", "K"]
  },
  "special_rows": {
    "amortizacao_total": 199,
    "dividendos_pagos_positivo": 209,
    "dividendos_pagos_negativo": 210,
    "reavaliacao_imobilizado": 213
  },
  "metadata_rows": {
    "data_periodo": 3,
    "meses_balanco": 4,
    "nome_arquivo": 5
  },
  "skip_rows": [199, 209, 210, 213]
}
```

Todos os modulos devem obter posicoes deste arquivo via `core/schema.py`. Nenhuma constante numerica de posicao deve existir no codigo.

---

## 6. Template Vazio

O repositorio contera um template vazio em `templates/Spread Proxy Template.xlsx`:
- Estrutura completa (326 linhas, 12 colunas)
- Rotulos na coluna B
- Formulas nas linhas calculadas (totais, subtotais)
- Formatacao visual (bordas, cores de secao)
- Nenhum valor em nenhuma coluna de dados (D/F/H/J/L)

O hash SHA-256 do template e registrado para detectar mudancas acidentais.

---

## Historico

| Data | Mudanca |
|------|---------|
| 2026-04-09 | Criacao com layout completo, linhas especiais, e referencia ao spread_schema.json |
