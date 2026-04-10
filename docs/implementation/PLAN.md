# Plano de Implementacao

> Plano completo de implementacao em 5+1 phases. Cada phase tem objetivo, entregaveis, dependencias, riscos e definicao de pronto.
>
> Aprovado em 2026-04-09. Executar em ordem estrita (0→1→2→3→4→5).

---

## Mapa de Dependencias

```
Phase 0 (Fundacao)
  │
  ├──→ Phase 1 (Mode 1A refatorado)
  │      │
  │      ├──→ Phase 2 (Mode 1B: build from scratch)
  │      │
  │      ├──→ Phase 3 (PDF pipeline: markitdown)
  │      │
  │      └──→ Phase 1C (CVM Analysis) [bloqueada ate estrutura fornecida]
  │
  └──→ (nenhuma phase comeca sem Phase 0 completa)

Phase 2 + Phase 3
  │
  └──→ Phase 4 (UX + Design System) [bloqueada ate DS fornecido]
         │
         └──→ Phase 5 (Empacotamento .exe)
```

---

## Phase 0 — Fundacao

**Objetivo:** Desacoplar todas as constantes hardcoded (linhas, colunas, mapas de contas) do codigo e move-las para arquivos de configuracao versionados. Criar os modelos de dominio. Estabelecer suporte a empacotamento desde o dia 1.

Nenhum comportamento muda. Nenhuma funcionalidade nova.

### Entregaveis

| # | Entregavel | Descricao |
|---|-----------|-----------|
| E0.1 | `spread_schema.json` | Layout fixo do Spread: linhas especiais, colunas, start_row, sheet_name |
| E0.2 | `core/schema.py` | Loader + validator do JSON; expoe `SpreadSchema` dataclass frozen |
| E0.3 | `core/models.py` | `FinancialAccount`, `FinancialDataSet`, `MappingResult` dataclasses |
| E0.4 | `core/exceptions.py` | `SpreadAutomationError`, `IngestionError`, `MappingError`, `ValidationError`, `SchemaError` |
| E0.5 | `core/resources.py` | `get_resource_path()` compativel com PyInstaller (`sys._MEIPASS`) |
| E0.6 | `core/periods.py` | Migrar funcoes puras de `core/utils.py`; manter re-export temporario |
| E0.7 | `mapping_tables/conta_spread_map.json` | Migracao exata de `CONTA_SPREAD_MAP` de Python para JSON |
| E0.8 | `mapping_tables/dre_spread_map.json` | Migracao exata de `DRE_SPREAD_MAP` de Python para JSON |
| E0.9 | `mapping/registry.py` | Carrega JSONs via `resources.py`; cache em memoria |
| E0.10 | Atualizacao dos modulos existentes | `processing/*.py` e `app/gui.py` usam schema + registry |
| E0.11 | `build.spec` | PyInstaller spec esboco com `--add-data` para JSONs e themes |
| E0.12 | Teste de regressao | Minerva 4T24 produz output identico ao baseline |

### Dependencias
Nenhuma externa. Trabalha apenas com codigo existente.

### Riscos
- Re-export de `core/utils.py` quebra imports → Mitigacao: grep completo antes de alterar
- `spread_schema.json` com valor errado → Mitigacao: teste de regressao e o gate
- `build.spec` falha com customtkinter → Nao bloqueia Phase 0; refinado na Phase 5

### Definicao de Pronto
1. Zero constantes numericas de posicao fora de `spread_schema.json`
2. Zero imports de mapas de Python — tudo via JSONs e registry
3. Minerva 4T24 processada com output identico ao baseline
4. `core/models.py` com dataclasses importaveis
5. `build.spec` gera executavel que abre a GUI

---

## Phase 1 — Mode 1A Refatorado

**Objetivo:** Reimplementar a funcionalidade atual usando a arquitetura de camadas (ingestion → mapping → validation → spread writer). `processing/pipeline.py` deve poder ser deprecated.

### Entregaveis

| # | Entregavel | Descricao |
|---|-----------|-----------|
| E1.1 | `ingestion/base.py` | `BaseIngestionAdapter` ABC |
| E1.2 | `ingestion/cvm_excel.py` | DadosDocumento.xlsx → `FinancialDataSet` (absorve `origin.py`) |
| E1.3 | `mapping/layer1_code.py` | CD_CONTA exact match via registry; confidence=1.0 |
| E1.4 | `mapping/layer2_value.py` | Value matching numerico; preserva logica de formulas |
| E1.5 | `mapping/mapper.py` | Orquestra layers; inclui sub-mappers DFC/DMPL/DRE |
| E1.6 | `spread/reader.py` | Le rotulos, valores existentes, metadados via SpreadSchema |
| E1.7 | `spread/writer.py` | Escreve MappingResults; openpyxl primary, xlwings bonus |
| E1.8 | `spread/highlights.py` | Migrado de `processing/highlights.py` |
| E1.9 | `validation/completeness.py` | Matched/unmatched/unused report |
| E1.10 | `validation/reporter.py` | Formata report para GUI |
| E1.11 | `engine/progress.py` | `ProgressCallback = Callable[[str, float], None]` |
| E1.12 | `engine/base_workflow.py` | `BaseWorkflow` ABC |
| E1.13 | `engine/workflow_1a.py` | Orquestra pipeline completo Mode 1A |
| E1.14 | Atualizacao `app/gui.py` | GUI fala com WorkflowEngine; thread separada |
| E1.15 | Deprecar `processing/pipeline.py` | Docstring deprecated; nao deletar |

### Dependencias
Phase 0 completa.

### Riscos
- Prioridade Layer1 vs Layer2 invertida → Manter Layer2 first (ADR-003)
- Logica de formulas dificil de encapsular → Manter como utilities em core/periods.py
- Thread GUI + xlwings COM → xlwings na main thread, openpyxl na worker
- DRE trimestral nao encaixa no mapper → Sub-mapper `DreQuarterlyMapper`

### Definicao de Pronto
1. `Workflow1A.run()` output identico a `processar()` para Minerva 4T24
2. Multi-periodo reproduzido por `Workflow1A.run(multi=True)`
3. `app/gui.py` nao importa de `processing/`
4. GUI nao congela durante processamento
5. Relatorio de completude no log

---

## Phase 2 — Mode 1B (Build from Scratch)

**Objetivo:** Construir Spread completo do zero a partir de multiplos DFPs/ITRs.

### Entregaveis

| # | Entregavel | Descricao |
|---|-----------|-----------|
| E2.1 | `templates/Spread Proxy Template.xlsx` | Template vazio versionado no repo |
| E2.2 | `spread/template.py` | Copia do template; validacao vs schema |
| E2.3 | `validation/period_coverage.py` | Gaps de periodo; "faltam 2022, 2023" |
| E2.4 | `validation/consistency.py` | Ativo == Passivo + PL |
| E2.5 | `engine/workflow_1b.py` | Loop multi-periodo: valida → por periodo → map → write |
| E2.6 | Tela Mode 1B na GUI | Multi-file selector, coverage report |
| E2.7 | Mode selector na GUI | Tela inicial: 1A / 1B |

### Dependencias
- Phase 1 completa
- Template vazio do Spread Proxy disponivel

### Riscos
- Primeiro periodo sem Layer 2 → Expandir CONTA_SPREAD_MAP
- Template incompativel com schema → Validacao automatica no startup
- Deteccao de periodo falha → Fallback: perguntar ao usuario
- Visao trimestral requer muitos arquivos → Coverage report mostra gaps

### Definicao de Pronto
1. 4 DFPs Minerva (2021-2024) → Spread com D/F/H/J preenchidos
2. Coverage report identifica gaps corretamente
3. Consistency check funciona
4. Mode selector funcional na GUI

---

## Phase 3 — PDF Pipeline (markitdown)

**Objetivo:** Suportar extracao de dados financeiros de PDFs nativos usando markitdown + parsing proprio.

### Entregaveis

| # | Entregavel | Descricao |
|---|-----------|-----------|
| E3.1 | `requirements.txt` atualizado | markitdown[pdf], rapidfuzz >= 3.0 |
| E3.2 | `ingestion/pdf/extractor.py` | markitdown → Markdown; detecta PDF escaneado |
| E3.3 | `ingestion/pdf/parser.py` | Identifica tabelas financeiras no Markdown |
| E3.4 | `ingestion/pdf/normalizer.py` | Normaliza numeros BR (1.234.567 → int) |
| E3.5 | `ingestion/pdf/adapter.py` | Orquestra extractor → parser → normalizer → FinancialDataSet |
| E3.6 | `mapping_tables/account_synonyms.json` | Sinonimos de contas CVM para fuzzy |
| E3.7 | `mapping/layer3_fuzzy.py` | rapidfuzz matching; thresholds auto/review/reject |
| E3.8 | `engine/workflow_2a.py` | PDF + Spread existente → revisao → escrita |
| E3.9 | `engine/workflow_2b.py` | PDF + template vazio |
| E3.10 | Tela de revisao PDF na GUI | Lista candidatos < 0.95 para confirmacao |

### Dependencias
- Phase 1 completa
- markitdown e rapidfuzz instalaveis
- PDF de referencia para testes

### Riscos
- markitdown produz Markdown com tabelas mal formatadas → Fallback manual
- Layout multi-coluna confunde parser → Detectar e avisar
- Fuzzy false positives → Threshold conservador (0.95) + revisao obrigatoria
- ESCALA_MOEDA diferente → Heuristica de deteccao
- Acentos afetam score → Normalizar com unidecode

### Definicao de Pronto
1. PDF nativo real → >= 70% linhas com confidence >= 0.95
2. Tela de revisao funcional
3. PDF escaneado detectado e recusado
4. Log detalhado de extracao

---

## Phase 4 — UX + Design System

**Objetivo:** Transformar a GUI funcional em experiencia profissional usando o Design System fornecido.

**BLOQUEADO** ate Design System ser fornecido.

### Entregaveis

| # | Entregavel | Descricao |
|---|-----------|-----------|
| E4.1 | Design System integrado | Cores, tipografia, spacing aplicados |
| E4.2 | `app/widgets/file_drop.py` | Drag-and-drop via tkinterdnd2 |
| E4.3 | `app/widgets/progress_bar.py` | Barra animada com texto descritivo |
| E4.4 | `app/widgets/log_panel.py` | Painel colapsavel; mensagens amigaveis |
| E4.5 | Todas as telas em `app/screens/` | mode_selector, screen_1a, screen_1b, screen_2 |
| E4.6 | `app/gui.py` refatorado | Multi-tela, worker thread, handler de excecao |
| E4.7 | Logging profissional | Rotation, formato padrao, stack trace no arquivo |

### Dependencias
- Phases 0-3 completas
- Design System fornecido
- tkinterdnd2 instalavel

### Riscos
- DS incompativel com customtkinter → Avaliar ao receber
- tkinterdnd2 DLL issues → Fallback: botao de selecao
- Worker thread deadlock → Pattern queue + after()

### Definicao de Pronto
1. GUI segue Design System
2. Drag-and-drop funciona
3. Progress bar com texto em todas as workflows
4. Log panel amigavel (sem traceback)
5. Nenhum congelamento de UI

---

## Phase 5 — Empacotamento .exe

**Objetivo:** Gerar executavel Windows distribuivel que funciona sem Python.

### Entregaveis

| # | Entregavel | Descricao |
|---|-----------|-----------|
| E5.1 | `build.spec` finalizado | Inclui JSONs, templates, DLLs, hidden imports; exclui xlwings |
| E5.2 | `core/resources.py` validado em .exe | Todos os recursos acessiveis no bundle |
| E5.3 | Teste em maquina limpa | VM sem Python; todos os modos funcionam |
| E5.4 | Teste em ambiente corporativo | Proxy, antivirus, diretorios restritos |
| E5.5 | README de distribuicao | Instrucoes para usuario final (1 pagina) |

### Dependencias
- Phase 4 completa
- PyInstaller no ambiente de build
- Maquina de teste limpa

### Riscos
- Bundle grande (pandas ~50MB) → Avaliar remoção de pandas
- Antivirus bloqueia .exe → Code signing ou whitelist
- tkinterdnd2 DLLs → Receita PyInstaller conhecida
- Paths com acentos → Usar Path objects; testar com "C:\Users\Joao\"

### Definicao de Pronto
1. `pyinstaller build.spec` sem erros
2. .exe funciona sem Python
3. Windows Defender nao bloqueia
4. Todos os modos funcionam
5. Bundle < 150MB
6. Paths com acentos funcionam

---

## Phase 1C — CVM Analysis (Paralela, Bloqueada)

**Objetivo:** Integrar CVM Analysis como fonte alternativa (e potencialmente primaria).

**BLOQUEADO** ate estrutura CVM Analysis ser fornecida.

### Entregaveis
| # | Entregavel | Descricao |
|---|-----------|-----------|
| E-1C.1 | Analise da estrutura | Documentar formato, campos, CD_CONTA |
| E-1C.2 | `ingestion/cvm_analysis.py` | Adapter → FinancialDataSet |
| E-1C.3 | `ingestion/cvm_csv.py` (se aplicavel) | CSVs portal CVM → FinancialDataSet |
| E-1C.4 | `engine/workflow_1c.py` | Reutiliza mapper + writer |
| E-1C.5 | Opcao na GUI | Terceira fonte no mode selector |

### Dependencias
- Phase 1 completa
- Estrutura CVM Analysis fornecida

---

## Decisoes Pendentes

| Decisao | Bloqueia | Input necessario |
|---------|----------|-----------------|
| Estrutura CVM Analysis | Phase 1C | Joao fornecer a estrutura |
| Design System | Phase 4 | Joao fornecer o DS |
| Template vazio do Spread | Phase 2 | Confirmar que existe copia limpa |
| PDF de referencia | Phase 3 | Joao fornecer PDF real para teste |
| Prioridade Layer 1 vs Layer 2 | Phase 1 | Decidir: inverter ou manter? |

---

## Historico

| Data | Mudanca |
|------|---------|
| 2026-04-09 | Criacao com 5+1 phases, 56 entregaveis, gates por phase |
