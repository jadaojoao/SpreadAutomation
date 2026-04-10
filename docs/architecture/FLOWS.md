# Fluxos de Trabalho

> Diagramas detalhados de cada modo de operação do sistema.

---

## 1. Visão Geral

| Modo | Frente | Entrada | Saída | Layers usados | Revisão? | Status |
|------|--------|---------|-------|---------------|----------|--------|
| **1A** | ITR/DFP | DadosDocumento.xlsx + Spread existente | Spread + novo período | L1 → L2 | Não | Funcional (v1) |
| **1B** | ITR/DFP | N arquivos DFP/ITR | Spread completo do zero | L1 (+ L2 a partir do 2o período) | Não | Planejado |
| **1C** | CVM Analysis | Estrutura CVM Analysis | Spread preenchido | L1 | Não | Bloqueado |
| **2A** | PDF | PDFs + Spread existente | Spread preenchido | L3 → L2 | Sim | Planejado |
| **2B** | PDF | PDFs | Spread do zero | L3 | Sim | Planejado |

---

## 2. Mode 1A — Preenchimento de Spread Existente

O fluxo que existe hoje, refatorado na arquitetura v2.

```
ENTRADA
  DadosDocumento.xlsx (CVM)
  Spread Proxy.xlsx (com período(s) anterior(es))
  Tipo: consolidado | individual
  Período: "2024" ou "4T24"
       │
       ▼
┌──────────────────────────────────────┐
│  1. INGESTÃO                          │
│  CvmExcelAdapter.ingest()            │
│  • Seleciona abas (Cons/Ind)         │
│  • Renomeia colunas de período       │
│  • Detecta DFP vs ITR               │
│  → FinancialDataSet (confidence=1.0) │
└──────────────┬───────────────────────┘
               │
               ▼
┌──────────────────────────────────────┐
│  2. LEITURA DO SPREAD                 │
│  SpreadReader.read()                 │
│  • Rótulos col B (row→label)         │
│  • Valores período anterior (row→val)│
│  • Detecta próxima coluna livre      │
└──────────────┬───────────────────────┘
               │
               ▼
┌──────────────────────────────────────┐
│  3. MAPEAMENTO                        │
│  AccountMapper.map_dataset()         │
│                                      │
│  Para cada linha do Spread:          │
│    Layer 2 (valor): busca valor      │
│    idêntico no período anterior      │
│    ↓ se não encontrou                │
│    Layer 1 (código): busca CD_CONTA  │
│    via conta_spread_map.json         │
│                                      │
│  Tratamento especial:                │
│  • DRE trimestral (label-based)      │
│  • DFC depreciação (regex 6.01.*)    │
│  • DMPL dividendos/capital           │
│  • Fórmulas (shift + adjust)         │
│                                      │
│  → List[MappingResult]              │
└──────────────┬───────────────────────┘
               │
               ▼
┌──────────────────────────────────────┐
│  4. VALIDAÇÃO                         │
│  CompletenessChecker.check()         │
│  • Linhas matched vs unmatched       │
│  • Accounts unused                   │
│  → CompletenessReport               │
└──────────────┬───────────────────────┘
               │
               ▼
┌──────────────────────────────────────┐
│  5. ESCRITA                           │
│  SpreadWriter.write()                │
│  • openpyxl (primary)                │
│  • xlwings (bonus, se disponível)    │
│  → "Spread Proxy 2024.xlsx"          │
└──────────────┬───────────────────────┘
               │
               ▼
┌──────────────────────────────────────┐
│  6. DESTAQUES                         │
│  Highlights                          │
│  • Verde: valores correspondidos     │
│  • Azul: valores novos (era 0)       │
│  → "DadosDocumento_tratado.xlsx"     │
└──────────────────────────────────────┘

SAÍDA
  Spread Proxy {período}.xlsx
  DadosDocumento_tratado.xlsx
  CompletenessReport (no log)
```

### Suporte multi-período (1A)
O Mode 1A também suporta preenchimento de 2 períodos consecutivos (ant → atual). O workflow executa o pipeline acima duas vezes em sequência, usando o output do primeiro como input do segundo.

---

## 3. Mode 1B — Construção do Zero

```
ENTRADA
  N arquivos DadosDocumento.xlsx (ex.: Minerva 2021, 2022, 2023, 2024)
  Anos desejados: [2021, 2022, 2023, 2024]
  Visão: anual | trimestral | ambas
  Tipo: consolidado | individual
       │
       ▼
┌──────────────────────────────────────┐
│  1. INGESTÃO (para cada arquivo)      │
│  CvmExcelAdapter.ingest()            │
│  → List[FinancialDataSet]            │
│  (um por período)                    │
└──────────────┬───────────────────────┘
               │
               ▼
┌──────────────────────────────────────┐
│  2. VALIDAÇÃO DE COBERTURA            │
│  PeriodCoverageChecker.check()       │
│  • "Encontrados: 2021, 2022, 2024"   │
│  • "Faltando: 2023"                  │
│  → CoverageReport                   │
│                                      │
│  Se gaps: retorna para GUI           │
│  Usuário confirma ou cancela         │
└──────────────┬───────────────────────┘
               │ (se completo ou confirmado)
               ▼
┌──────────────────────────────────────┐
│  3. CRIAÇÃO DO TEMPLATE               │
│  SpreadTemplate.create_from_template()│
│  → Spread vazio com rótulos e        │
│    fórmulas, sem valores             │
└──────────────┬───────────────────────┘
               │
               ▼
┌──────────────────────────────────────┐
│  4. LOOP DE PREENCHIMENTO             │
│  Para cada período (cronológico):    │
│                                      │
│  Período 1 (mais antigo):            │
│    Apenas Layer 1 (sem anterior)     │
│    Escreve na coluna D               │
│                                      │
│  Período 2:                          │
│    Layer 2 (usa col D como ref)      │
│    + Layer 1 (fallback)              │
│    Escreve na coluna F               │
│                                      │
│  Período 3:                          │
│    Layer 2 (usa col F) + Layer 1     │
│    Escreve na coluna H               │
│                                      │
│  Período 4:                          │
│    Layer 2 (usa col H) + Layer 1     │
│    Escreve na coluna J               │
└──────────────┬───────────────────────┘
               │
               ▼
┌──────────────────────────────────────┐
│  5. VALIDAÇÃO DE CONSISTÊNCIA         │
│  ConsistencyChecker.check()          │
│  • Ativo == Passivo + PL (por col)   │
│  → warnings se divergência           │
└──────────────┬───────────────────────┘
               │
               ▼
SAÍDA
  Spread Proxy {empresa}.xlsx (completo)
  CompletenessReport por período
  ConsistencyReport
```

### Limitação do primeiro período
Sem Spread anterior, Layer 2 (matching por valor) não funciona no primeiro período. Apenas Layer 1 (CD_CONTA) está disponível. Linhas não cobertas pela Layer 1 ficarão vazias. Essa é a principal motivação para expandir a cobertura do `conta_spread_map.json`.

---

## 4. Mode 1C — CVM Analysis (placeholder)

```
ENTRADA
  Estrutura CVM Analysis (formato a definir)
       │
       ▼
┌──────────────────────────────────────┐
│  CvmAnalysisAdapter.ingest()         │
│  → FinancialDataSet                  │
│                                      │
│  Se tiver CD_CONTA: confidence=1.0   │
│  Se não: necessita Layer 3           │
└──────────────┬───────────────────────┘
               │
               ▼
  (restante do pipeline idêntico ao 1A ou 1B)
```

**Potencial:** se a estrutura CVM Analysis tiver CD_CONTA e cobrir todos os períodos, ela pode substituir DadosDocumento.xlsx como fonte padrão, tornando Mode 1B trivial (todos os anos em um download).

---

## 5. Mode 2A — Preenchimento a partir de PDFs

```
ENTRADA
  PDFs de demonstrações financeiras
  Spread Proxy.xlsx existente
       │
       ▼
┌──────────────────────────────────────┐
│  1. EXTRAÇÃO                          │
│  MarkitdownExtractor.extract()       │
│  • markitdown.convert(pdf) → MD      │
│  • Detecta PDF escaneado (output     │
│    vazio) → RECUSA com mensagem      │
└──────────────┬───────────────────────┘
               │
               ▼
┌──────────────────────────────────────┐
│  2. PARSING                           │
│  MarkdownFinancialParser.parse()     │
│  • Identifica tabelas no Markdown    │
│  • Detecta seção (BP/DRE/DFC/DMPL)  │
│  • Filtra tabelas não-financeiras    │
└──────────────┬───────────────────────┘
               │
               ▼
┌──────────────────────────────────────┐
│  3. NORMALIZAÇÃO                      │
│  BrazilianNumberNormalizer.normalize()│
│  • "1.234.567" → 1234567            │
│  • "(1.234)" → -1234                │
│  • "R$ 1.234 mil" → 1234000         │
└──────────────┬───────────────────────┘
               │
               ▼
┌──────────────────────────────────────┐
│  4. ADAPTAÇÃO                         │
│  PdfIngestionAdapter                 │
│  → FinancialDataSet                  │
│    code = None (PDFs não têm código) │
│    confidence = baseado na extração  │
│    source = "pdf"                    │
└──────────────┬───────────────────────┘
               │
               ▼
┌──────────────────────────────────────┐
│  5. MAPEAMENTO                        │
│  AccountMapper.map_dataset()         │
│  • Layer 3 (fuzzy): descrição do PDF │
│    vs rótulos do Spread              │
│    vs account_synonyms.json          │
│  • Layer 2 (valor): se Spread tem    │
│    período anterior                  │
│                                      │
│  Cada match recebe confidence score  │
└──────────────┬───────────────────────┘
               │
               ▼
┌──────────────────────────────────────┐
│  6. CLASSIFICAÇÃO                     │
│                                      │
│  confidence >= 0.95 → AUTO           │
│  0.60 <= confidence < 0.95 → REVIEW  │
│  confidence < 0.60 → REJECTED        │
└──────┬────────────┬──────────────────┘
       │            │
  AUTO │       REVIEW
       │            │
       │     ┌──────▼──────────────────┐
       │     │  7. TELA DE REVISÃO     │
       │     │  GUI mostra candidatos  │
       │     │  Usuário confirma/rej.  │
       │     │  Botão "Confirmar todos │
       │     │  acima de X%"           │
       │     └──────┬──────────────────┘
       │            │ confirmados
       ▼            ▼
┌──────────────────────────────────────┐
│  8. ESCRITA                           │
│  SpreadWriter.write()                │
│  Escreve apenas items AUTO +         │
│  items confirmados na revisão        │
└──────────────────────────────────────┘

SAÍDA
  Spread Proxy preenchido
  Log de extração (o que foi extraído, o que foi descartado)
  Relatório de confiança
```

---

## 6. Mode 2B — Construção do Zero a partir de PDFs

Idêntico ao Mode 2A, com uma diferença:
- Em vez de receber Spread existente, usa `SpreadTemplate.create_from_template()`
- Layer 2 (valor) não disponível no primeiro período (sem referência anterior)
- Depende ainda mais da qualidade da Layer 3 (fuzzy)

---

## 7. Componentes Compartilhados

| Componente | 1A | 1B | 1C | 2A | 2B |
|------------|----|----|----|----|-----|
| `core/models.py` | x | x | x | x | x |
| `core/schema.py` | x | x | x | x | x |
| `mapping/layer1_code.py` | x | x | x | - | - |
| `mapping/layer2_value.py` | x | x | x | x | x |
| `mapping/layer3_fuzzy.py` | - | - | - | x | x |
| `mapping/mapper.py` | x | x | x | x | x |
| `spread/reader.py` | x | x | x | x | x |
| `spread/writer.py` | x | x | x | x | x |
| `spread/template.py` | - | x | - | - | x |
| `validation/completeness.py` | x | x | x | x | x |
| `validation/consistency.py` | - | x | - | - | x |
| `validation/period_coverage.py` | - | x | - | - | - |
| `spread/highlights.py` | x | x | x | - | - |
| `ingestion/cvm_excel.py` | x | x | - | - | - |
| `ingestion/pdf/adapter.py` | - | - | - | x | x |

---

## Histórico

| Data | Mudança |
|------|---------|
| 2026-04-09 | Criação com fluxos detalhados para os 5 modos + componentes compartilhados |
