# Design do Sistema

> Descreve as camadas, módulos, responsabilidades e contratos que formam a arquitetura v2 do Spread Automation.

---

## 1. Princípio Fundamental

O sistema é um pipeline ETL com schema de destino fixo:

```
[Fontes variáveis] → [Modelo canônico] → [Mapper] → [Spread Proxy fixo]
```

- **Extract:** adaptadores por fonte convertem dados brutos em `FinancialDataSet`
- **Transform:** mapper aplica 3 camadas de correspondência, produzindo `MappingResult` com scores de confiança
- **Load:** writer escreve os resultados no Spread Proxy respeitando o layout fixo

---

## 2. Diagrama de Camadas

```
┌─────────────────────────────────────────────────────────────┐
│  PRESENTATION                                                │
│  app/                                                        │
│  CustomTkinter — drag-drop, progress, log, screens           │
│  Fala SOMENTE com engine/                                    │
└──────────────────────────┬──────────────────────────────────┘
                           │ ProgressCallback
┌──────────────────────────▼──────────────────────────────────┐
│  ORCHESTRATION                                               │
│  engine/                                                     │
│  WorkflowEngine — detecta modo, valida, coordena pipeline    │
│  Expõe: run(**kwargs) → WorkflowResult                       │
└──────┬──────────────────────────────────────────────┬───────┘
       │                                              │
┌──────▼──────────────┐                    ┌──────────▼───────┐
│  INGESTION           │                    │  SPREAD I/O       │
│  ingestion/          │                    │  spread/          │
│                      │                    │                   │
│  BaseIngestionAdapter│                    │  SpreadReader     │
│    ├─ CvmExcelAdapter│                    │  SpreadWriter     │
│    ├─ CvmCsvAdapter  │                    │  SpreadTemplate   │
│    ├─ CvmAnalysis(*) │                    │  Highlights       │
│    └─ PdfAdapter     │                    │                   │
│                      │                    │  Motor primário:  │
│  Output:             │                    │    openpyxl       │
│  FinancialDataSet    │                    │  Motor bônus:     │
│                      │                    │    xlwings        │
└──────┬──────────────┘                    └──────────▲───────┘
       │                                              │
       │              ┌──────────────────────┐        │
       └──────────────▶  MAPPING             ├────────┘
                      │  mapping/            │
                      │                      │
                      │  AccountMapper       │
                      │    ├─ Layer 1: Code  │
                      │    ├─ Layer 2: Value │
                      │    └─ Layer 3: Fuzzy │
                      │                      │
                      │  Input:              │
                      │    FinancialDataSet   │
                      │    + SpreadLabels     │
                      │  Output:             │
                      │    List[MappingResult]│
                      └──────────┬───────────┘
                                 │
                      ┌──────────▼───────────┐
                      │  VALIDATION          │
                      │  validation/         │
                      │                      │
                      │  CompletenessChecker │
                      │  ConsistencyChecker  │
                      │  PeriodCoverage      │
                      │  Reporter            │
                      └──────────┬───────────┘
                                 │
                      ┌──────────▼───────────┐
                      │  CORE                │
                      │  core/               │
                      │                      │
                      │  models.py           │
                      │  schema.py           │
                      │  periods.py          │
                      │  exceptions.py       │
                      │  resources.py        │
                      │                      │
                      │  ZERO deps externas  │
                      └──────────────────────┘

(*) = placeholder até estrutura fornecida
```

---

## 3. Responsabilidades por Camada

### Core (`core/`)
**Regra de ouro:** zero importações de camadas superiores, zero dependências externas.

| Módulo | Responsabilidade |
|--------|-----------------|
| `models.py` | `FinancialAccount`, `FinancialDataSet`, `MappingResult`, `WorkflowResult` |
| `schema.py` | Carrega e valida `spread_schema.json`; expõe `SpreadSchema` (dataclass frozen) |
| `periods.py` | `periodos()`, `normaliza_num()`, `col_txt_to_idx()`, `shift_formula()`, `adjust_complex_formula()` |
| `exceptions.py` | `SpreadAutomationError`, `IngestionError`, `MappingError`, `ValidationError`, `SchemaError` |
| `resources.py` | `get_resource_path()` — localiza arquivos em dev e no bundle .exe |

### Ingestion (`ingestion/`)
Adaptadores que convertem fontes brutas em `FinancialDataSet`.

| Módulo | Fonte | Output |
|--------|-------|--------|
| `base.py` | — | `BaseIngestionAdapter` ABC com `ingest() → FinancialDataSet` |
| `cvm_excel.py` | DadosDocumento.xlsx | `FinancialDataSet` com `confidence=1.0`, `source_type="cvm_xlsx"` |
| `cvm_csv.py` | CSVs dados abertos CVM | `FinancialDataSet` com `confidence=1.0`, `source_type="cvm_csv"` |
| `cvm_analysis.py` | Estrutura CVM Analysis | Placeholder |
| `pdf/extractor.py` | PDF (qualquer) | Markdown string via markitdown |
| `pdf/parser.py` | Markdown string | `List[RawFinancialTable]` agrupadas por seção |
| `pdf/normalizer.py` | Números brutos | Inteiros normalizados (formato BR → int) |
| `pdf/adapter.py` | PDF completo | `FinancialDataSet` com `confidence<1.0`, `code=None`, `source_type="pdf"` |

### Mapping (`mapping/`)
Três camadas de correspondência com confiança decrescente.

| Módulo | Layer | Confiança | Quando usado |
|--------|-------|-----------|-------------|
| `layer1_code.py` | CD_CONTA exato | 1.0 | CVM Excel/CSV (sempre) |
| `layer2_value.py` | Valor numérico | ~0.85 | Quando Layer 1 não resolve + período anterior existe |
| `layer3_fuzzy.py` | Similaridade textual | 0.0–1.0 | PDFs (sem CD_CONTA) |
| `mapper.py` | Orquestrador | — | Coordena layers; expõe `map_dataset()` |
| `registry.py` | — | — | Carrega e cacheia `mapping_tables/*.json` |

### Spread I/O (`spread/`)

| Módulo | Responsabilidade |
|--------|-----------------|
| `reader.py` | Lê rótulos (col B), valores existentes, metadados do Spread |
| `writer.py` | Escreve `List[MappingResult]` no Spread (openpyxl/xlwings) |
| `template.py` | Cria cópia do template vazio para Mode 1B/2B |
| `highlights.py` | Destaques visuais (verde=match, azul=novo) no DadosDocumento_tratado |

### Validation (`validation/`)

| Módulo | O que valida |
|--------|-------------|
| `completeness.py` | Linhas do Spread sem correspondência; accounts não utilizados |
| `consistency.py` | Ativo == Passivo + PL; Receita + Custos == Lucro Bruto |
| `period_coverage.py` | Gaps de período para Mode 1B (quais anos/trimestres faltam) |
| `reporter.py` | Formata relatórios de validação para a GUI |

### Engine (`engine/`)
Orquestra fluxos de trabalho completos.

| Módulo | Modo | Descrição |
|--------|------|-----------|
| `base_workflow.py` | — | `BaseWorkflow` ABC com `run() → WorkflowResult` |
| `workflow_1a.py` | 1A | CVM Excel + Spread existente |
| `workflow_1b.py` | 1B | Múltiplos CVM Excel + template vazio |
| `workflow_1c.py` | 1C | CVM Analysis (placeholder) |
| `workflow_2a.py` | 2A | PDF + Spread existente |
| `workflow_2b.py` | 2B | PDF + template vazio |
| `progress.py` | — | Protocolo de callback de progresso |

### App (`app/`)

| Módulo | Responsabilidade |
|--------|-----------------|
| `gui.py` | Janela principal, navegação multi-tela |
| `widgets/file_drop.py` | Widget drag-and-drop |
| `widgets/progress_bar.py` | Barra de progresso com texto descritivo |
| `widgets/log_panel.py` | Painel de log colapsável |
| `screens/mode_selector.py` | Tela inicial com seleção de modo |
| `screens/screen_1a.py` | Inputs para Mode 1A |
| `screens/screen_1b.py` | Inputs para Mode 1B |
| `screens/screen_2.py` | Inputs + revisão para Modes 2A/2B |

---

## 4. Contratos entre Camadas

### Ingestion → Mapping
```
FinancialDataSet {
    company, period, entity_type, source_type,
    accounts: List[FinancialAccount]
}
```
O mapper não sabe e não precisa saber de onde vieram os dados.

### Mapping → Spread Writer
```
List[MappingResult] {
    spread_row, spread_label, source_account,
    mapped_value, confidence, layer
}
```
O writer não sabe como o valor foi obtido. Escreve se `confidence >= threshold`.

### Engine → App
```
ProgressCallback = Callable[[str, float], None]
    str: mensagem ("Lendo DadosDocumento...")
    float: progresso 0.0 – 1.0

WorkflowResult {
    output_path, report, warnings, pending_review
}
```
A GUI nunca importa diretamente de ingestion, mapping ou spread.

---

## 5. Regras de Importação

```
core/           → importa: stdlib, typing
ingestion/      → importa: core/, libs externas (pandas, openpyxl, markitdown)
mapping/        → importa: core/
spread/         → importa: core/, openpyxl, xlwings (opcional)
validation/     → importa: core/
engine/         → importa: core/, ingestion/, mapping/, spread/, validation/
app/            → importa: engine/ (SOMENTE)
```

**Proibido:**
- `app/` importar de `processing/`, `ingestion/`, `mapping/`, `spread/`
- `core/` importar de qualquer outra camada
- Importações circulares entre camadas

---

## 6. Motor Duplo (xlwings vs openpyxl)

| Motor | Quando usado | Vantagem | Limitação |
|-------|-------------|----------|-----------|
| openpyxl | Sempre (primary) | Funciona sem Excel, portátil, PyInstaller-safe | Não atualiza Excel aberto ao vivo |
| xlwings | Desenvolvimento apenas (bonus) | Atualiza Excel aberto, fórmulas recalculam | Requer COM, Excel instalado, Windows only |

**Regra:** o writer tenta xlwings se disponível e se Excel estiver aberto. Se falha, usa openpyxl. Nunca executa os dois para a mesma operação.

**No bundle .exe:** xlwings não é incluído. openpyxl é o único motor.

---

## 7. Dependências Externas

| Pacote | Tipo | Uso | Incluído no .exe |
|--------|------|-----|-------------------|
| `pandas` | Core | DataFrames de leitura | Sim (avaliar remoção na Phase 5) |
| `openpyxl` | Core | Leitura/escrita Excel | Sim |
| `customtkinter` | Core | GUI | Sim |
| `markitdown[pdf]` | Phase 3 | PDF → Markdown | Sim |
| `rapidfuzz` | Phase 3 | Fuzzy matching | Sim |
| `tkinterdnd2` | Phase 4 | Drag-and-drop | Sim |
| `xlwings` | Opcional | Excel ao vivo | Não |

---

## Histórico

| Data | Mudança |
|------|---------|
| 2026-04-09 | Criação com arquitetura v2 completa (6 camadas, contratos, regras de importação) |
