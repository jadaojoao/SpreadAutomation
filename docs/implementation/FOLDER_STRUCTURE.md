# Estrutura de Pastas Alvo

> Organizacao do codigo-fonte apos todas as phases serem concluidas. Cada modulo listado inclui a phase em que sera criado.

---

## 1. Visao Geral

```
spread_automation/
├── core/                          # Camada base — sem dependencias internas
│   ├── __init__.py
│   ├── models.py                  # Phase 0 — FinancialAccount, FinancialDataSet, MappingResult
│   ├── schema.py                  # Phase 0 — Loader + validator de spread_schema.json
│   ├── exceptions.py              # Phase 0 — Hierarquia de excecoes
│   ├── resources.py               # Phase 0 — get_resource_path() (PyInstaller-compatible)
│   ├── periods.py                 # Phase 0 — Funcoes puras de periodo (migrado de utils.py)
│   └── utils.py                   # Existente → mantido com re-export ate Phase 1
│
├── ingestion/                     # Adaptadores de fonte → FinancialDataSet
│   ├── __init__.py
│   ├── base.py                    # Phase 1 — BaseIngestionAdapter ABC
│   ├── cvm_excel.py               # Phase 1 — DadosDocumento.xlsx (absorve origin.py)
│   ├── cvm_analysis.py            # Phase 1C — CVM Analysis (bloqueado)
│   ├── cvm_csv.py                 # Phase 1C — CSVs portal CVM (se aplicavel)
│   └── pdf/                       # Phase 3 — Pipeline PDF
│       ├── __init__.py
│       ├── extractor.py           # Phase 3 — markitdown → Markdown
│       ├── parser.py              # Phase 3 — Identifica tabelas financeiras
│       ├── normalizer.py          # Phase 3 — Normaliza numeros BR
│       └── adapter.py             # Phase 3 — Orquestra pipeline → FinancialDataSet
│
├── mapping/                       # Camadas de correspondencia
│   ├── __init__.py
│   ├── registry.py                # Phase 0 — Carrega JSONs via resources.py
│   ├── layer1_code.py             # Phase 1 — CD_CONTA exact match
│   ├── layer2_value.py            # Phase 1 — Valor numerico match
│   ├── layer3_fuzzy.py            # Phase 3 — rapidfuzz textual match
│   └── mapper.py                  # Phase 1 — Orquestra layers + sub-mappers
│
├── spread/                        # Leitura e escrita do Spread Proxy
│   ├── __init__.py
│   ├── reader.py                  # Phase 1 — Le rotulos, valores, metadados
│   ├── writer.py                  # Phase 1 — Escreve MappingResults (openpyxl + xlwings)
│   ├── template.py                # Phase 2 — Copia e valida template vazio
│   └── highlights.py              # Phase 1 — Destaques visuais (migrado de processing/)
│
├── validation/                    # Verificacoes pos-mapeamento
│   ├── __init__.py
│   ├── completeness.py            # Phase 1 — Matched/unmatched/unused
│   ├── consistency.py             # Phase 2 — Ativo == Passivo + PL
│   ├── period_coverage.py         # Phase 2 — Gaps de periodo
│   └── reporter.py                # Phase 1 — Formata reports para GUI
│
├── engine/                        # Orquestracao de workflows
│   ├── __init__.py
│   ├── progress.py                # Phase 1 — ProgressCallback type
│   ├── base_workflow.py           # Phase 1 — BaseWorkflow ABC
│   ├── workflow_1a.py             # Phase 1 — Mode 1A (preenchimento existente)
│   ├── workflow_1b.py             # Phase 2 — Mode 1B (construcao do zero)
│   ├── workflow_1c.py             # Phase 1C — Mode 1C (CVM Analysis)
│   ├── workflow_2a.py             # Phase 3 — Mode 2A (PDF + Spread existente)
│   └── workflow_2b.py             # Phase 3 — Mode 2B (PDF + template)
│
├── app/                           # Interface grafica
│   ├── __init__.py
│   ├── gui.py                     # Existente → refatorado Phase 1, redesenhado Phase 4
│   ├── screens/                   # Phase 4 — Telas separadas
│   │   ├── __init__.py
│   │   ├── mode_selector.py       # Phase 4 — Tela inicial: escolha de modo
│   │   ├── screen_1a.py           # Phase 4 — Tela Mode 1A
│   │   ├── screen_1b.py           # Phase 4 — Tela Mode 1B
│   │   └── screen_2.py            # Phase 4 — Tela Modes 2A/2B (PDF)
│   └── widgets/                   # Phase 4 — Componentes reutilizaveis
│       ├── __init__.py
│       ├── file_drop.py           # Phase 4 — Drag-and-drop (tkinterdnd2)
│       ├── progress_bar.py        # Phase 4 — Barra animada com texto
│       └── log_panel.py           # Phase 4 — Painel colapsavel de mensagens
│
├── mapping_tables/                # Tabelas de mapeamento (JSON)
│   ├── conta_spread_map.json      # Phase 0 — Migrado de CONTA_SPREAD_MAP
│   ├── dre_spread_map.json        # Phase 0 — Migrado de DRE_SPREAD_MAP
│   └── account_synonyms.json      # Phase 3 — Sinonimos para fuzzy matching
│
├── templates/                     # Templates de Spread
│   └── Spread Proxy Template.xlsx # Phase 2 — Template vazio versionado
│
├── themes/                        # Design System
│   └── (a definir)                # Phase 4 — Cores, fontes, assets
│
├── processing/                    # DEPRECATED apos Phase 1
│   ├── __init__.py                #   Mantido para compatibilidade
│   ├── origin.py                  #   → absorvido por ingestion/cvm_excel.py
│   ├── spread.py                  #   → absorvido por mapping/ + spread/
│   ├── dre.py                     #   → absorvido por mapping/mapper.py
│   ├── dfc.py                     #   → absorvido por mapping/mapper.py
│   ├── dmpl.py                    #   → absorvido por mapping/mapper.py
│   ├── highlights.py              #   → movido para spread/highlights.py
│   └── pipeline.py                #   → substituido por engine/workflow_*.py
│
├── data/                          # Dados de trabalho (gitignored)
├── docs/                          # Documentacao do projeto
│   ├── PROJECT_MASTER.md          # Documento mestre
│   ├── product/
│   │   └── VISION.md
│   ├── architecture/
│   │   ├── DECISIONS.md
│   │   ├── SYSTEM_DESIGN.md
│   │   ├── DATA_MODEL.md
│   │   ├── FLOWS.md
│   │   └── MAPPING_STRATEGY.md
│   ├── domain/
│   │   ├── SPREAD_LAYOUT.md
│   │   └── CVM_DATA_FORMATS.md
│   ├── implementation/
│   │   ├── PLAN.md
│   │   ├── FOLDER_STRUCTURE.md
│   │   └── PACKAGING.md
│   ├── governance/
│   │   ├── parallel-lanes.md
│   │   ├── operators-runbook.md
│   │   └── rollback-recovery.md
│   └── reference/
│       ├── plano-contas-fixas-DFP/
│       ├── meta_dfp_cia_aberta_txt/
│       ├── cvm_account_dictionary.csv
│       ├── Manual XML ITR v2.xls
│       └── manual-de-envio-*.pdf
│
├── main.py                        # Ponto de entrada
├── build.spec                     # Phase 0 (esboco) → Phase 5 (finalizado)
├── requirements.txt
├── CONTEXT.md                     # Referencia historica (superseded por docs/)
├── MEMORIADASIA.md                # Memoria historica (superseded por docs/)
├── AGENTS.md                      # Contrato operacional
├── CONTRIBUTING.md
├── CHANGELOG.md
├── LICENSE
└── README.md
```

---

## 2. Regras de Import

As camadas seguem uma hierarquia estrita de dependencias. Imports sao permitidos apenas de cima para baixo.

```
app/        → engine/
engine/     → ingestion/, mapping/, spread/, validation/, core/
ingestion/  → core/
mapping/    → core/
spread/     → core/
validation/ → core/
core/       → (nenhum import interno)
```

### Proibicoes explicitas

| De | Para | Razao |
|----|------|-------|
| `core/` | qualquer outro pacote | Core e a base; nao pode ter dependencias circulares |
| `app/` | `processing/`, `ingestion/`, `mapping/`, `spread/` | GUI fala apenas com engine |
| `mapping/` | `spread/` | Mapper recebe dados via parametro, nao le o Spread diretamente |
| `ingestion/` | `mapping/` | Ingestao produz dados; nao sabe como serao mapeados |

---

## 3. Progressao por Phase

### Phase 0 — Fundacao

Arquivos criados:
```
core/models.py
core/schema.py
core/exceptions.py
core/resources.py
core/periods.py
mapping/registry.py
mapping_tables/conta_spread_map.json
mapping_tables/dre_spread_map.json
build.spec (esboco)
```

Arquivos modificados:
```
core/utils.py          → funcoes migradas para periods.py; re-export mantido
processing/*.py        → usa schema + registry em vez de constantes
app/gui.py             → usa schema em vez de constantes hardcoded
```

### Phase 1 — Mode 1A Refatorado

Arquivos criados:
```
ingestion/base.py
ingestion/cvm_excel.py
mapping/layer1_code.py
mapping/layer2_value.py
mapping/mapper.py
spread/reader.py
spread/writer.py
spread/highlights.py
validation/completeness.py
validation/reporter.py
engine/progress.py
engine/base_workflow.py
engine/workflow_1a.py
```

Arquivos modificados:
```
app/gui.py             → fala com engine, nao com processing/
processing/pipeline.py → docstring deprecated
```

### Phase 2 — Mode 1B

Arquivos criados:
```
templates/Spread Proxy Template.xlsx
spread/template.py
validation/period_coverage.py
validation/consistency.py
engine/workflow_1b.py
```

Arquivos modificados:
```
app/gui.py             → mode selector + tela 1B
```

### Phase 3 — PDF Pipeline

Arquivos criados:
```
ingestion/pdf/extractor.py
ingestion/pdf/parser.py
ingestion/pdf/normalizer.py
ingestion/pdf/adapter.py
mapping/layer3_fuzzy.py
mapping_tables/account_synonyms.json
engine/workflow_2a.py
engine/workflow_2b.py
```

Arquivos modificados:
```
mapping/mapper.py      → integra Layer 3
requirements.txt       → markitdown[pdf], rapidfuzz
app/gui.py             → tela de revisao PDF
```

### Phase 4 — UX + Design System

Arquivos criados:
```
app/screens/mode_selector.py
app/screens/screen_1a.py
app/screens/screen_1b.py
app/screens/screen_2.py
app/widgets/file_drop.py
app/widgets/progress_bar.py
app/widgets/log_panel.py
themes/                → assets do Design System
```

Arquivos modificados:
```
app/gui.py             → refatorado para multi-tela
requirements.txt       → tkinterdnd2
```

### Phase 5 — Empacotamento

Arquivos modificados:
```
build.spec             → finalizado com todos os recursos
core/resources.py      → validado em contexto .exe
```

---

## 4. Diretorio `processing/` — Plano de Deprecacao

O diretorio `processing/` nao sera deletado. Seus modulos serao marcados como deprecated com docstrings e mantidos ate que todos os consumidores migrem.

| Modulo antigo | Substituto | Phase de migracao |
|---------------|-----------|-------------------|
| `processing/origin.py` | `ingestion/cvm_excel.py` | Phase 1 |
| `processing/spread.py` | `mapping/layer1_code.py` + `mapping/layer2_value.py` + `spread/reader.py` | Phase 1 |
| `processing/dre.py` | `mapping/mapper.py` (sub-mapper DRE) | Phase 1 |
| `processing/dfc.py` | `mapping/mapper.py` (sub-mapper DFC) | Phase 1 |
| `processing/dmpl.py` | `mapping/mapper.py` (sub-mapper DMPL) | Phase 1 |
| `processing/highlights.py` | `spread/highlights.py` | Phase 1 |
| `processing/pipeline.py` | `engine/workflow_1a.py` | Phase 1 |

Criterio para remocao: zero imports de `processing/` em qualquer modulo. Verificavel com `grep -r "from processing" --include="*.py"`.

---

## 5. Convencoes de Nomeacao

| Tipo | Convencao | Exemplo |
|------|-----------|---------|
| Modulos | `snake_case.py` | `cvm_excel.py` |
| Classes | `PascalCase` | `CvmExcelAdapter` |
| Funcoes | `snake_case` | `ingest()` |
| Constantes | `UPPER_SNAKE_CASE` | `SKIP_ROWS` |
| Arquivos JSON | `snake_case.json` | `conta_spread_map.json` |
| Workflows | `workflow_{mode}.py` | `workflow_1a.py` |
| Screens | `screen_{mode}.py` | `screen_1a.py` |

---

## Historico

| Data | Mudanca |
|------|---------|
| 2026-04-09 | Criacao com estrutura alvo completa, regras de import, progressao por phase |
