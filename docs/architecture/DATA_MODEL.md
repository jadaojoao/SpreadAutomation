# Modelo de Dados

> Define os objetos de domínio que formam o contrato entre camadas do sistema.

---

## 1. FinancialAccount

Representa uma única conta contábil extraída de qualquer fonte.

```
FinancialAccount
├── code: str | None           CD_CONTA CVM (ex.: "3.02")
│                              None quando vem de PDF (sem código)
├── description: str           DS_CONTA ou rótulo extraído
│                              Ex.: "Custo dos Bens e/ou Serviços Vendidos"
├── value: Decimal             Valor numérico da conta
│                              Sempre em unidade (não em milhares)
├── statement_type: str        "BPA", "BPP", "DRE", "DFC", "DMPL"
├── confidence: float          0.0 – 1.0
│                              1.0 = fonte CVM com CD_CONTA
│                              <1.0 = extração de PDF
└── source: str                "cvm_xlsx", "cvm_csv", "cvm_analysis", "pdf"
```

### Invariantes
- `description` nunca é vazio
- `value` nunca é None (usar 0 para contas zeradas)
- `confidence` está em [0.0, 1.0]
- Se `source` == "cvm_xlsx" ou "cvm_csv", `code` nunca é None
- Se `source` == "pdf", `code` é sempre None

---

## 2. FinancialDataSet

Conjunto de contas financeiras extraídas de uma fonte, representando uma empresa em um período.

```
FinancialDataSet
├── company: str               Nome da empresa (ex.: "Minerva")
├── period: str                Período (ex.: "2024", "4T24")
├── entity_type: str           "consolidated" ou "individual"
├── source_type: str           "cvm_xlsx", "cvm_csv", "cvm_analysis", "pdf"
├── accounts: list[FinancialAccount]
│                              Todas as contas de todos os demonstrativos
├── metadata: dict             Informações adicionais (opcional)
│   ├── scale: str             "Unidade", "Mil" (se detectado)
│   ├── version: int           VERSAO do documento CVM (se disponível)
│   └── file_path: str         Caminho do arquivo de origem
└── warnings: list[str]        Avisos gerados durante ingestão
```

### Invariantes
- `accounts` nunca é vazio (se não há contas, a ingestão falhou)
- `period` segue o formato `"YYYY"` ou `"nTYY"`
- `entity_type` é estritamente "consolidated" ou "individual"

### Agrupamento por demonstrativo

O `FinancialDataSet` contém contas de todos os demonstrativos misturadas. Para acessar por demonstrativo:

```python
bp_ativo = [a for a in dataset.accounts if a.statement_type == "BPA"]
dre      = [a for a in dataset.accounts if a.statement_type == "DRE"]
```

---

## 3. SpreadSchema

Configuração do layout do Spread Proxy, carregada de `spread_schema.json`.

```
SpreadSchema (frozen dataclass)
├── sheet_name: str            "Entrada de Dados"
├── data_start_row: int        27
├── columns: SpreadColumns
│   ├── labels: str            "B"
│   ├── annual: list[str]      ["D", "F", "H", "J"]
│   ├── quarterly: str         "L"
│   └── hidden: list[str]      ["A", "C", "E", "G", "I", "K"]
├── special_rows: dict[str, int]
│   ├── amortizacao_total: 199
│   ├── dividendos_pagos_positivo: 209
│   ├── dividendos_pagos_negativo: 210
│   └── reavaliacao_imobilizado: 213
├── metadata_rows: dict[str, int]
│   ├── data_periodo: 3
│   ├── meses_balanco: 4
│   └── nome_arquivo: 5
└── skip_rows: list[int]       [199, 209, 210, 213]
```

### Invariantes
- `data_start_row` > 0
- `skip_rows` == set dos valores de `special_rows`
- `annual` tem exatamente 4 colunas
- Todas as colunas de `annual` + `quarterly` não estão em `hidden`

### Carregamento
- `core/schema.py` carrega o JSON e valida invariantes no startup
- Se o JSON é inválido, `SchemaError` é levantado antes de qualquer operação
- Schema é um singleton — carregado uma vez, reutilizado

---

## 4. MappingResult

Resultado de um mapeamento de uma conta para uma posição no Spread.

```
MappingResult
├── spread_row: int            Linha no Spread (ex.: 42)
├── spread_label: str          Rótulo da col B (ex.: "Disponibilidades")
├── source_account: FinancialAccount | None
│                              A conta que originou o match
│                              None se não houve match
├── mapped_value: int | None   Valor a ser escrito
│                              None se não houve match
├── confidence: float          Score do match (0.0 – 1.0)
├── layer: str                 "code", "value", "fuzzy", "none"
│                              "none" = nenhuma camada encontrou match
├── candidates: list[tuple[str, float]] | None
│                              Top 3 candidatos fuzzy (Layer 3)
│                              Usado na tela de revisão de PDFs
└── formula: str | None        Fórmula deslocada (se a célula original era fórmula)
```

### Invariantes
- Se `layer` == "none", `mapped_value` é None e `confidence` == 0.0
- Se `layer` == "code", `confidence` == 1.0
- Se `layer` == "fuzzy", `candidates` não é None
- `spread_row` >= `schema.data_start_row`

---

## 5. WorkflowResult

Output de qualquer workflow executado pelo engine.

```
WorkflowResult
├── output_path: Path          Caminho do Spread gerado
├── report: CompletenessReport
│   ├── matched: int           Linhas preenchidas com sucesso
│   ├── unmatched: int         Linhas do Spread sem correspondência
│   ├── unused: int            Accounts da fonte não utilizados
│   └── details: list[str]     Detalhes por linha não correspondida
├── warnings: list[str]        Avisos (ESCALA_MOEDA, reapresentação, etc.)
├── pending_review: list[MappingResult] | None
│                              Itens com confidence < 0.95 (fluxo PDF)
│                              GUI deve mostrar para confirmação
└── consistency: ConsistencyReport | None
    ├── asset_liability_match: bool
    ├── delta: int             Diferença Ativo - (Passivo+PL)
    └── warnings: list[str]
```

---

## 6. Relação entre Objetos

```
                     ┌──────────────┐
                     │ SpreadSchema │ (singleton, carregado do JSON)
                     └──────┬───────┘
                            │ usado por
         ┌──────────────────┼──────────────────┐
         │                  │                  │
┌────────▼────────┐ ┌──────▼──────┐ ┌─────────▼─────────┐
│ SpreadReader    │ │ SpreadWriter│ │ AccountMapper     │
│ lê labels+vals  │ │ escreve     │ │ mapeia             │
└────────┬────────┘ └──────▲──────┘ └────┬──────────────┘
         │                 │             │
         │    ┌────────────┘             │
         │    │ List[MappingResult]      │
         │    │                          │
         │    │    ┌─────────────────────┘
         │    │    │
         │    │    │ FinancialDataSet
         │    │    │
    ┌────▼────┴────▼────┐
    │ WorkflowEngine    │
    │ (engine/)         │
    └───────────────────┘
```

---

## Histórico

| Data | Mudança |
|------|---------|
| 2026-04-09 | Criação com 5 objetos de domínio e suas relações |
