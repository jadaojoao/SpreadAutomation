# Estrategia de Mapeamento

> Descreve as 3 camadas de correspondencia entre fontes de dados e a estrutura fixa do Spread, incluindo scoring de confianca, validacoes e limitacoes conhecidas.

---

## 1. Principio: Confianca Decrescente

O sistema aplica ate 3 camadas de correspondencia em cascata. Cada camada tem uma confianca associada. A primeira camada que encontra match vence.

```
Layer 1 (codigo CVM)     confidence = 1.00    deterministico
    |
    v (se nao encontrou)
Layer 2 (valor numerico)  confidence ~ 0.85    heuristico
    |
    v (se nao encontrou)
Layer 3 (fuzzy textual)   confidence = variavel  semantico
```

**Nota historica (ADR-003):** no codigo atual, Layer 2 executa antes de Layer 1. A ordem foi mantida por compatibilidade com outputs validados. A inversao (Layer 1 primeiro) sera avaliada apos a Phase 1.

---

## 2. Layer 1 — Matching por CD_CONTA Exato

### Mecanismo
1. Le o rotulo da coluna B do Spread (ex.: "CMV Total")
2. Consulta `mapping_tables/conta_spread_map.json`
3. Obtem lista de CD_CONTA CVM (ex.: `["3.02"]`)
4. Busca esses codigos na coluna `Codigo Conta` do `FinancialDataSet`
5. Soma os valores encontrados

### Caracteristicas
- **Confidence:** 1.0 (match exato, sem ambiguidade)
- **Quando funciona:** sempre que a fonte tem CD_CONTA (CVM Excel, CVM CSV)
- **Quando NAO funciona:** PDFs (nao tem CD_CONTA)
- **Cobertura atual:** ~25 entradas para BP e DRE
- **Fonte dos mapeamentos:** Plano de Contas Fixas DFP — ENET 10.0

### Mapeamentos atuais validados

#### Balanco Patrimonial — Ativo
| Rotulo no Spread | CD_CONTA CVM | Status |
|------------------|-------------|--------|
| Disponibilidades | 1.01.01, 1.01.02 | Validado (Minerva, Sabesp) |
| Dupls Receber CP MN | 1.01.03.01 | Validado (Minerva) |
| Estoques Operacionais CP MN | 1.01.04 | Validado (Minerva) |
| Ativos Biologicos CP | 1.01.05 | Validado (Minerva — zero p/ Sabesp, esperado) |
| Impostos a Recuperar CP | 1.01.06 | Validado |
| Desp. Antecipadas | 1.01.07 | Validado |
| Disponibilidades LP | 1.02.01.01, 1.02.01.02 | Validado (zero p/ Minerva 4T24) |
| Dupls Receber LP MN | 1.02.01.03.01 | Validado (nao encontrado p/ Sabesp — sub-niveis) |
| Dupls Receber LP ME | 1.02.01.03.02 | Validado |
| Credito Tributario I.R. LP | 1.02.01.06 | Validado (zero p/ Minerva 4T24) |
| C/C Coligadas RLP MN | 1.02.01.08.01 | Validado |

#### Balanco Patrimonial — Passivo
| Rotulo no Spread | CD_CONTA CVM | Status |
|------------------|-------------|--------|
| Fornecedores Operac. CP MN | 2.01.02.01 | Validado (Minerva; zero p/ Sabesp) |
| Despesas Provisionadas | 2.01.01 | Validado |
| Provisao IR/CS | 2.01.03.01.01 | Validado |
| Dividendos a Pagar | 2.01.05.02.01, 2.01.05.02.02 | Validado |
| Provisao IR/CS LP | 2.02.03.01 | Validado |
| Participacao dos Minoritarios | 2.03.09 | Validado |

#### DRE
| Rotulo no Spread | CD_CONTA CVM | Status |
|------------------|-------------|--------|
| Vendas Mercado Externo | 3.01 | Validado (convencao: receita total CVM, ADR-H01) |
| CMV Total | 3.02 | Validado |
| Lucro Bruto | 3.03 | Validado |
| Despesas de Vendas | 3.04.01 | Validado |
| Despesas Administrativas | 3.04.02 | Validado |
| Outras Despesas Operacionais | 3.04.05, 3.04.03 | Validado (inclui Perdas NRA) |
| Outras Receitas Operacionais | 3.04.04 | Validado |
| Equivalencia Patrimonial | 3.04.06 | Validado |
| Receitas Financeiras Caixa | 3.06.01 | Validado |
| Despesas Financeiras Caixa | 3.06.02 | Validado |
| Imposto de Renda | 3.08 | Validado |

#### DMPL (usados por processing/dmpl.py)
| Conta | CD_CONTA | Status |
|-------|----------|--------|
| Dividendos | 5.04.06 | Validado |
| JCP | 5.04.07 | Validado |
| Aumentos de Capital | 5.04.01 | Validado |

### Limitacoes deliberadas

Linhas NAO incluidas no CONTA_SPREAD_MAP, com justificativa:

| Linha no Spread | Razao da omissao |
|-----------------|-----------------|
| Bancos CP/LP MN/ME | CVM classifica por instrumento (Emprestimos/Debentures), Spread por moeda (MN/ME) — taxonomias ortogonais |
| Imobilizado bruto/depreciacao acum. | Spread decompoe em sub-linhas bancarias (terrenos, edificios) sem equivalente CVM direto |
| Reserva de Capital | Nomenclatura divergente entre Spread e CVM |
| Investimentos | CVM 1.02.02 = apenas investimentos financeiros; Spread inclui mais itens |
| Diferido Normal / Amortizacao | CVM mostra 1.02.04 ja liquido; Spread decompoe em bruto - amortizacao |

Essas linhas caem para Layer 2 (matching por valor).

---

## 3. Layer 2 — Matching por Valor Numerico

### Mecanismo
1. Le o valor na coluna-fonte do Spread (periodo anterior)
2. Procura valor numerico identico na coluna `ant` do FinancialDataSet
3. Se encontra: retorna o valor da coluna `atual` da mesma linha

### Tratamento de formulas
- **Formula literal** (`=123456+789012`): substitui cada numero via Layer 2
- **Formula complexa** (`=SUM(H28:H31)`): desloca referencias de coluna (H→J) E substitui literais
- **Valor puro**: substituicao direta

### Caracteristicas
- **Confidence:** ~0.85 (heuristico, pode ter ambiguidade)
- **Quando funciona:** quando existe periodo anterior no Spread com valores
- **Quando NAO funciona:** primeiro periodo do Mode 1B; valores reapresentados; ESCALA_MOEDA inconsistente

### Riscos conhecidos

| Risco | Descricao | Frequencia | Impacto |
|-------|-----------|-----------|---------|
| Valores duplicados | Duas contas com mesmo valor no mesmo periodo → match errado | Medio | Alto — valor vai para a linha errada |
| Reapresentacao | Empresa republicou demonstracao; valor anterior mudou | Baixo | Alto — Layer 2 nao encontra match |
| ESCALA_MOEDA | Empresa mudou de "Unidade" para "Mil" → valores ~1000x menores | Raro | Alto — nenhum match |
| Zero como valor | Muitas contas zeradas → match no primeiro zero encontrado | Medio | Baixo — escreveria 0 de qualquer forma |

### Criterio de parada
A varredura para apos 30 linhas vazias consecutivas no Spread.

---

## 4. Layer 3 — Matching Fuzzy Textual

### Mecanismo
1. Para cada `FinancialAccount` sem CD_CONTA (vindo de PDF):
2. Compara `account.description` com cada rotulo do Spread usando `rapidfuzz.fuzz.token_sort_ratio()`
3. Compara tambem contra `account_synonyms.json` (dicionario de sinonimos CVM)
4. Score final = max(match direto, match via sinonimo)

### Thresholds

| Score | Classificacao | Acao |
|-------|--------------|------|
| >= 0.95 | AUTO | Escreve automaticamente |
| 0.60 – 0.95 | REVIEW | Candidato para revisao — GUI mostra |
| < 0.60 | REJECTED | Descartado com log |

### account_synonyms.json
Construido a partir de `docs/reference/cvm_account_dictionary.csv` (17.068 DS_CONTA unicos).

Formato:
```json
{
  "BPA": [
    {
      "descriptions": ["Caixa e Equivalentes de Caixa", "Disponibilidades"],
      "canonical_label": "Disponibilidades",
      "codes": ["1.01.01"]
    }
  ],
  "DRE": [
    {
      "descriptions": ["Receita Operacional Liquida", "Receita de Venda de Bens e/ou Servicos"],
      "canonical_label": "Vendas Mercado Externo",
      "codes": ["3.01"]
    }
  ]
}
```

### Normalizacao antes da comparacao
- Remover acentos (unidecode)
- Lowercase
- Remover artigos e preposicoes comuns ("de", "do", "da", "e", "ou")
- Normalizar espacos

### Caracteristicas
- **Confidence:** variavel (0.0 – 1.0)
- **Quando funciona:** PDFs com descricoes textuais legiveis
- **Quando NAO funciona:** PDFs escaneados; tabelas com headers genericos; dados sem descricao textual

---

## 5. DRE Trimestral — Caso Especial

A DRE trimestral tem um mecanismo de matching proprio: label-based matching via `dre_spread_map.json`.

### Mecanismo
1. Escaneia a coluna B do Spread de `start_row` ate o final
2. Para cada descricao CVM no `dre_spread_map.json`:
   - Encontra a linha do Spread cujo rotulo corresponde ao rotulo-alvo
   - Soma valores se multiplas contas CVM apontam para o mesmo rotulo
3. Escreve os totais acumulados

### Mapeamento DRE Trimestral

| Descricao CVM | Rotulo no Spread |
|---------------|-----------------|
| Receita de Venda de Bens e/ou Servicos | Vendas Mercado Externo |
| Custo dos Bens e/ou Servicos Vendidos | CMV Total |
| Resultado Bruto | Lucro Bruto |
| Despesas Gerais e Administrativas | Despesas Administrativas |
| Despesas com Vendas | Despesas de Vendas |
| Outras Despesas Operacionais | Outras Despesas Operacionais |
| Perdas pela Nao Recuperabilidade de Ativos | Outras Despesas Operacionais (somado) |
| Outras Receitas Operacionais | Outras Receitas Operacionais |
| Despesas Financeiras | Despesas Financeiras Caixa |
| Receitas Financeiras | Receitas Financeiras Caixa |
| Resultado de Equivalencia Patrimonial | Equivalencia Patrimonial |
| Imposto de Renda e CS sobre o Lucro | Imposto de Renda |

### Quando ativado
Apenas quando `is_trim == True` (periodo no formato "nTYY").

---

## 6. DFC e DMPL — Tratamentos Especiais

### DFC — Depreciacao/Amortizacao
- Filtra apenas secao 6.01 (atividades operacionais) — ADR-H03
- Busca textual por regex: `deprecia|amortiza|exaustao`
- Soma e grava como negativo na linha `special_rows.amortizacao_total`
- NAO usa Layer 1 (contas DFC nao sao fixas no CVM)

### DMPL — Dividendos e Capital
- Layer 1: busca por CD_CONTA 5.04.06 (dividendos) + 5.04.07 (JCP)
- Layer 2 fallback: busca textual por `dividendo|juros sobre capital`
- Negativos → `special_rows.dividendos_pagos_negativo`
- Positivos → `special_rows.dividendos_pagos_positivo`
- Aumentos de capital: CD_CONTA 5.04.01 → `special_rows.reavaliacao_imobilizado`

---

## Historico

| Data | Mudanca |
|------|---------|
| 2026-04-09 | Criacao com 3 layers, DRE/DFC/DMPL especiais, mapeamentos validados |
