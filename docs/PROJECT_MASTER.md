# Spread Automation — Documento Mestre

> Fonte central de verdade do projeto. Leia este documento primeiro.
> Para profundidade em qualquer tópico, siga os links para os documentos satélite.

---

## 1. Resumo Executivo

**Spread Automation** automatiza a geração de planilhas de Spread Proxy usadas em análise de crédito corporativa. O sistema recebe dados financeiros de empresas brasileiras de capital aberto (via CVM ou PDFs) e os organiza em uma estrutura padronizada fixa.

**Estado atual:** sistema funcional para um modo de operação (Mode 1A — preenchimento de Spread existente a partir de DadosDocumento.xlsx CVM). Código monolítico, sem testes automatizados, sem empacotamento .exe.

**Estado futuro:** sistema multi-modo (5 fluxos de trabalho), suporte a PDFs, construção do zero, empacotamento em .exe para usuários não técnicos, arquitetura em camadas com modelo canônico de dados.

**Decisão arquitetural principal:** rearquitetar em torno de um modelo canônico (`FinancialDataSet`) que desacopla as fontes de entrada da estrutura fixa do Spread. Todas as fontes produzem o mesmo objeto; um único mapper escreve no Spread.

---

## 2. Índice de Documentação

### Ordem de leitura recomendada

| # | Documento | Propósito |
|---|-----------|-----------|
| 1 | **Este documento** | Visão completa do projeto |
| 2 | [Visão do Produto](product/VISION.md) | Objetivos, restrições, stakeholders, frentes de entrada |
| 3 | [Decisões de Arquitetura](architecture/DECISIONS.md) | Registro formal de todas as decisões técnicas (ADR) |
| 4 | [Design do Sistema](architecture/SYSTEM_DESIGN.md) | Camadas, módulos, dependências, contratos |
| 5 | [Modelo de Dados](architecture/DATA_MODEL.md) | FinancialDataSet, SpreadSchema, MappingResult |
| 6 | [Fluxos de Trabalho](architecture/FLOWS.md) | Diagramas detalhados dos modos 1A, 1B, 1C, 2A, 2B |
| 7 | [Estratégia de Mapeamento](architecture/MAPPING_STRATEGY.md) | 3 camadas de matching, scoring de confiança |
| 8 | [Layout do Spread](domain/SPREAD_LAYOUT.md) | Estrutura imutável do Spread Proxy |
| 9 | [Formatos CVM](domain/CVM_DATA_FORMATS.md) | Estrutura do DadosDocumento.xlsx e CSVs CVM |
| 10 | [Plano de Implementação](implementation/PLAN.md) | Phases 0–5 com tasks, gates, riscos |
| 11 | [Estrutura de Pastas](implementation/FOLDER_STRUCTURE.md) | Organização alvo do código-fonte |
| 12 | [Empacotamento .exe](implementation/PACKAGING.md) | Estratégia PyInstaller e distribuição |

### Documentação operacional

| Documento | Propósito |
|-----------|-----------|
| [AGENTS.md](../AGENTS.md) | Contrato operacional para contribuidores e agentes |
| [CONTRIBUTING.md](../CONTRIBUTING.md) | Guia de contribuição |
| [CHANGELOG.md](../CHANGELOG.md) | Histórico de mudanças |
| [Parallel Lanes](governance/parallel-lanes.md) | Governança de trabalho paralelo |
| [Operators Runbook](governance/operators-runbook.md) | Runbook operacional |
| [Path Policy](../.github/guardrails/path-policy.json) | Política de ownership por path |

### Referência de domínio (materiais CVM)

| Arquivo | Conteúdo |
|---------|----------|
| `docs/reference/cvm_account_dictionary.csv` | 17.068 DS_CONTA únicos por statement_type |
| `docs/reference/plano-contas-fixas-DFP/` | Plano de Contas Fixas DFP — ENET 10.0 |
| `docs/reference/meta_dfp_cia_aberta_txt/` | Schemas dos campos CSV dados abertos CVM |
| `docs/reference/Manual XML ITR v2.xls` | Leiaute do Formulário DFP/ITR — ENET |

---

## 3. Visão do Produto (resumo)

**Objetivo final:** gerar automaticamente um Spread Proxy Excel no formato padronizado usado em análise de crédito. A estrutura do Spread é fixa e imutável — as linhas sempre estão nas mesmas posições, independentemente da origem dos dados.

**Usuário final:** analista de crédito em ambiente corporativo restrito. Não técnico. Sem Python instalado. Fluxo ideal: abrir .exe, arrastar arquivos, clicar processar, receber Spread pronto.

**Restrições inegociáveis:**
- Sem WebApp, sem API (por enquanto)
- Sem dependência de instalação na máquina do usuário
- Executável .exe standalone
- Robustez e previsibilidade são requisitos, não features

**Duas frentes de entrada:**
- **Frente 1 — ITRs e DFPs:** dados estruturados da CVM (modos 1A, 1B, 1C)
- **Frente 2 — PDFs:** dados não estruturados de demonstrações financeiras (modos 2A, 2B)

> Detalhes completos: [Visão do Produto](product/VISION.md)

---

## 4. Arquitetura (resumo)

O sistema é um pipeline ETL com schema de destino fixo:

```
[Fontes variáveis] → [Modelo canônico] → [Mapper] → [Spread Proxy fixo]
```

### Camadas do sistema

```
┌─────────────────────────────────────────────────────────┐
│  PRESENTATION          app/                              │
│  CustomTkinter — drag-drop, progress, log, screens      │
└────────────────────────────┬────────────────────────────┘
                             │
┌────────────────────────────▼────────────────────────────┐
│  ORCHESTRATION         engine/                           │
│  WorkflowEngine — detecta modo, valida, drive pipeline  │
└──────┬─────────────────────────────────────┬────────────┘
       │                                     │
┌──────▼──────────┐               ┌──────────▼────────────┐
│  INGESTION       │               │  SPREAD I/O           │
│  ingestion/      │               │  spread/              │
│  Adaptadores     │               │  Reader + Writer      │
│  → FinancialDS   │               │  Template             │
└──────┬──────────┘               └──────────▲────────────┘
       │                                     │
┌──────▼──────────────────────────────────────┤
│  MAPPING           mapping/                  │
│  Layer 1: CD_CONTA exato                     │
│  Layer 2: valor numérico                     │
│  Layer 3: fuzzy textual (PDFs)               │
│  → MappingResult com confidence score        │
└──────┬──────────────────────────────────────┘
       │
┌──────▼──────────────────────────────────────┐
│  VALIDATION        validation/               │
│  Completude, consistência, cobertura         │
└──────┬──────────────────────────────────────┘
       │
┌──────▼──────────────────────────────────────┐
│  CORE              core/                     │
│  Modelos, schema, períodos, exceções         │
│  Zero dependências externas                  │
└─────────────────────────────────────────────┘
```

> Detalhes completos: [Design do Sistema](architecture/SYSTEM_DESIGN.md)

---

## 5. Fluxos de Trabalho (resumo)

| Modo | Entrada | Saída | Mapping | Status |
|------|---------|-------|---------|--------|
| **1A** | DadosDocumento.xlsx + Spread existente | Spread preenchido | Layer 1 → Layer 2 | Funcional (v1) |
| **1B** | Múltiplos DFPs/ITRs (sem Spread) | Spread do zero | Layer 1 (+ Layer 2 a partir do 2o período) | Planejado |
| **1C** | CVM Analysis structure | Spread preenchido | Layer 1 | Bloqueado (estrutura pendente) |
| **2A** | PDFs + Spread existente | Spread preenchido | Layer 3 → Layer 2 | Planejado |
| **2B** | PDFs (sem Spread) | Spread do zero | Layer 3 | Planejado |

> Detalhes completos: [Fluxos de Trabalho](architecture/FLOWS.md)

---

## 6. Plano de Implementação (resumo)

| Phase | Objetivo | Gate de Conclusão | Status |
|-------|----------|-------------------|--------|
| **0** | Fundação — desacoplar hardcoded → config | Minerva 4T24 idêntico ao atual | Pendente |
| **1** | Mode 1A refatorado na nova arquitetura | pipeline.py deprecated | Pendente |
| **2** | Mode 1B — build from scratch | 4 anos Minerva → Spread do zero | Pendente |
| **3** | PDF pipeline via markitdown | ≥70% linhas PDF com confidence ≥0.95 | Pendente |
| **4** | UX + Design System | UI profissional, drag-drop | Bloqueado (DS pendente) |
| **5** | Empacotamento .exe | Funciona sem Python | Pendente |
| **1C** | CVM Analysis como fonte | Output igual ao Mode 1A | Bloqueado (estrutura pendente) |

**Dependências:**

```
Phase 0 → Phase 1 → Phase 2
                   → Phase 3
                   → Phase 1C (quando disponível)
Phase 2 + 3 → Phase 4 (quando DS disponível) → Phase 5
```

> Detalhes completos: [Plano de Implementação](implementation/PLAN.md)

---

## 7. Riscos Ativos

| # | Risco | Severidade | Mitigação |
|---|-------|-----------|-----------|
| 1 | Linhas hardcoded (199, 209, 210, 213) — se template mudar, código quebra | Alta | Phase 0: migrar para `spread_schema.json` |
| 2 | PDF com tabelas mal formatadas — markitdown/pdfplumber não extrai | Alta | Threshold conservador + revisão obrigatória + mensagem clara de limitação |
| 3 | Sem testes automatizados — qualquer mudança pode quebrar silenciosamente | Alta | Fixture de regressão com Minerva 4T24 (gate de cada phase) |
| 4 | Antivírus corporativo bloqueia .exe não assinado | Alta | Investigar code signing; plano B: whitelist via TI |
| 5 | ESCALA_MOEDA inconsistente entre períodos (Mil vs Unidade) | Média | Validação na ingestion com aviso ao usuário |

---

## 8. Pendências e Bloqueios

### Aguardando input

| Item | Bloqueia | Impacto |
|------|----------|---------|
| Estrutura do CVM Analysis | Phase 1C | Pode ser a melhor fonte primária para todo o projeto |
| Design System | Phase 4 | Toda UI final depende disso |
| Template vazio do Spread Proxy | Phase 2 | Necessário para "build from scratch" |
| PDF de referência para teste | Phase 3 | Necessário para calibrar fuzzy matching |
| Decisão: prioridade Layer 1 vs Layer 2 | Phase 1 | Hoje Layer 2 é prioritário; inverter ou manter? |

### Decisões em aberto

| Decisão | Contexto | Recomendação |
|---------|----------|--------------|
| Substituir pandas por openpyxl puro | Reduzir tamanho do .exe (~50MB com pandas) | Avaliar na Phase 5 |
| Suportar PDFs escaneados (OCR) | markitdown-ocr requer API (LLM Vision) | Fora do escopo atual; registrar para futuro |
| Reconhecer reapresentações CVM | VERSAO > 1 muda valores retroativamente | Implementar validação na ingestion (Phase 1) |

---

## 9. Glossário

| Termo | Significado |
|-------|-------------|
| **CVM** | Comissão de Valores Mobiliários — regulador do mercado de capitais brasileiro |
| **DFP** | Demonstrações Financeiras Padronizadas — relatório anual obrigatório |
| **ITR** | Informações Trimestrais — relatório trimestral obrigatório |
| **Spread Proxy** | Planilha Excel de análise de crédito em formato bancário padronizado |
| **DadosDocumento.xlsx** | Arquivo Excel gerado pela CVM com demonstrações financeiras |
| **BP** | Balanço Patrimonial (Ativo + Passivo) |
| **DRE** | Demonstração do Resultado do Exercício |
| **DFC** | Demonstração dos Fluxos de Caixa |
| **DMPL** | Demonstração das Mutações do Patrimônio Líquido |
| **CD_CONTA** | Código contábil CVM hierárquico (ex.: `3.02`) |
| **DS_CONTA** | Descrição textual da conta CVM |
| **ST_CONTA_FIXA** | Flag CVM: `S` = conta sempre presente; `N` = discricionária |
| **Layer 1** | Matching determinístico por CD_CONTA exato (confidence = 1.0) |
| **Layer 2** | Matching heurístico por valor numérico idêntico (confidence ~ 0.85) |
| **Layer 3** | Matching fuzzy por similaridade textual via rapidfuzz (confidence variável) |
| **FinancialDataSet** | Modelo canônico de dados financeiros — contrato entre ingestion e mapping |
| **MappingResult** | Resultado de um mapeamento: valor + confidence + layer + destino no Spread |
| **SpreadSchema** | Configuração do layout do Spread carregada de `spread_schema.json` |
| **markitdown** | Biblioteca Microsoft para conversão de documentos → Markdown |
| **Consolidado** | Dados incluindo controladas (grupo econômico completo) |
| **Individual** | Dados apenas da empresa-mãe |

---

## Histórico deste documento

| Data | Mudança |
|------|---------|
| 2026-04-09 | Criação inicial com arquitetura v2 completa |
