# Decisões de Arquitetura (ADR)

> Registro formal de todas as decisões técnicas do projeto.
> Formato: contexto, opções avaliadas, decisão, consequências.
> Decisões históricas (migradas de MEMORIADASIA.md) prefixadas com `H`.

---

## ADR-001: Modelo canônico FinancialDataSet como contrato entre camadas

**Data:** 2026-04-09
**Status:** Aprovado

**Contexto:**
O sistema precisa suportar múltiplas fontes de entrada (CVM Excel, CVM CSV, PDFs, API futura) que produzem dados em formatos diferentes, mas que precisam alimentar o mesmo Spread fixo.

**Opções avaliadas:**
1. Cada adaptador escreve diretamente no Spread (acoplamento fonte-destino)
2. Modelo intermediário canônico que todas as fontes produzem (desacoplamento)
3. Interface genérica com dict/DataFrame (sem tipagem)

**Decisão:** Opção 2 — criar `FinancialDataSet` como contrato tipado entre ingestion e mapping.

**Consequências:**
- Novas fontes requerem apenas um novo adaptador, sem tocar no mapper ou writer
- O mapper é escrito uma vez e serve para todas as fontes
- A validação pode ser aplicada sobre o modelo canônico, não sobre formatos específicos
- Custo: refatoração da lógica atual que mistura leitura e escrita

---

## ADR-002: SpreadSchema como arquivo de configuração

**Data:** 2026-04-09
**Status:** Aprovado

**Contexto:**
Linhas especiais (199, 209, 210, 213), start_row (27), colunas (D/F/H/J/L), e skip_rows estão hardcoded em 6+ módulos diferentes. Se o template do Spread mudar, é necessário alterar código em múltiplos arquivos.

**Opções avaliadas:**
1. Manter hardcoded (status quo)
2. Constantes em um único módulo Python (centralizar, não desacoplar)
3. Arquivo JSON externo (`spread_schema.json`) carregado em runtime

**Decisão:** Opção 3 — arquivo JSON versionado, carregado por `core/schema.py`.

**Consequências:**
- Mudança de template requer edição de um único arquivo JSON
- JSON é incluído no bundle .exe via `--add-data`
- Necessário `core/resources.py` para localização PyInstaller-safe
- Validação de schema obrigatória no startup

---

## ADR-003: Manter prioridade Layer 2 > Layer 1 (compatibilidade)

**Data:** 2026-04-09
**Status:** Aprovado (temporário — reavaliar após Phase 1)

**Contexto:**
No código atual (`processing/spread.py`), o matching por valor numérico (Layer 2) executa primeiro. O matching por CD_CONTA (Layer 1) só entra como fallback quando Layer 2 não encontra correspondência. Isso é contraintuitivo (Layer 1 é mais confiável), mas é o comportamento validado com dados reais.

**Opções avaliadas:**
1. Inverter: Layer 1 primeiro, Layer 2 fallback (semanticamente correto)
2. Manter: Layer 2 primeiro, Layer 1 fallback (compatibilidade com outputs validados)

**Decisão:** Opção 2 — manter para não quebrar outputs durante refatoração.

**Consequências:**
- Outputs da Phase 1 refatorada serão idênticos ao sistema atual (gate de regressão)
- Após Phase 1 estabilizar, reavaliar inversão com testes A/B contra dados reais
- Documentar claramente que a ordem é por compatibilidade, não por design

---

## ADR-004: markitdown como camada de ingestão PDF

**Data:** 2026-04-09
**Status:** Aprovado

**Contexto:**
O usuário já testou pdfplumber diretamente e o resultado não foi satisfatório para PDFs financeiros brasileiros. Uma solução OCR pura não é viável em ambiente corporativo sem API.

**Opções avaliadas:**
1. `pdfplumber` direto (já testado, insatisfatório)
2. `camelot-py` (depende de Ghostscript e OpenCV — pesado para .exe)
3. `tabula-py` (depende de Java — inaceitável para .exe corporativo)
4. `markitdown` da Microsoft (usa pdfplumber internamente, normaliza output em Markdown)
5. `docling` da IBM (conversão de documentos, deps pesadas)

**Decisão:** Opção 4 — markitdown.

**Raciocínio:**
- Embora use pdfplumber por baixo, o output Markdown é significativamente mais fácil de parsear do que output bruto de pdfplumber
- Suporte unificado a Excel/Word/PPTX/PDF → Markdown (uma camada para tudo)
- Dependência leve, PyInstaller-compatible
- Plugin `markitdown-ocr` disponível para futuro com API (LLM Vision)
- Funciona 100% offline para conversão de formatos

**Consequências:**
- Mesmas limitações de pdfplumber para PDFs complexos (multi-coluna, tabelas aninhadas)
- PDFs escaneados continuam fora do escopo (sem OCR nativo)
- Pipeline: markitdown → Markdown → parser próprio → FinancialDataSet

---

## ADR-005: xlwings excluído do bundle .exe

**Data:** 2026-04-09
**Status:** Aprovado

**Contexto:**
xlwings requer COM automation e Excel instalado. Em ambiente corporativo, Excel pode não estar instalado ou COM pode estar restrito.

**Opções avaliadas:**
1. Incluir xlwings no bundle (risco de falha COM)
2. Excluir xlwings; openpyxl como único writer no .exe
3. Detectar em runtime e usar se disponível

**Decisão:** Opção 2 para o bundle .exe. Opção 3 para execução em desenvolvimento.

**Consequências:**
- openpyxl é o motor primário em todas as circunstâncias
- xlwings continua disponível ao rodar do código-fonte com `pip install xlwings`
- Código mantém `try/except ImportError` para xlwings (já existente)
- Sem escrita "ao vivo" no Excel aberto quando rodando como .exe

---

## ADR-006: Mapping tables em JSON, não em código Python

**Data:** 2026-04-09
**Status:** Aprovado

**Contexto:**
`CONTA_SPREAD_MAP` está em `core/conta_map.py` e `DRE_SPREAD_MAP` está em `core/utils.py` como dicts Python. Funciona, mas mistura dados com lógica e dificulta edição por não-programadores.

**Decisão:** Migrar para `mapping_tables/conta_spread_map.json` e `mapping_tables/dre_spread_map.json`.

**Consequências:**
- Editar mapeamentos não requer conhecimento de Python
- JSONs incluídos no bundle .exe via `--add-data`
- Necessário `mapping/registry.py` como loader e cache
- Validação de schema dos JSONs no startup

---

## ADR-007: CustomTkinter como framework de GUI

**Data:** 2026-04-09
**Status:** Aprovado

**Contexto:**
O projeto já usa CustomTkinter. Alternativas: PyQt/PySide (pesado, licença), Tkinter puro (feio), DearPyGui (menos maduro), web-based (proibido).

**Decisão:** Manter CustomTkinter.

**Raciocínio:**
- Já funciona no projeto
- Boa aparência visual sem configuração
- Compatível com PyInstaller
- Leve comparado a PyQt
- Suporta temas customizáveis (futuro Design System)

---

## ADR-008: PyInstaller --onedir para distribuição

**Data:** 2026-04-09
**Status:** Aprovado

**Contexto:**
PyInstaller oferece `--onefile` (um único .exe) e `--onedir` (diretório com .exe + dependências).

**Opções avaliadas:**
1. `--onefile`: um .exe, mais simples para o usuário, mas extrai para temp a cada execução (lento, antivírus suspeita)
2. `--onedir`: diretório distribuído como .zip, startup mais rápido, debugging mais fácil

**Decisão:** Opção 2 — `--onedir`.

**Raciocínio:**
- Dependências complexas (pdfminer, markitdown, customtkinter themes) são mais confiáveis em --onedir
- Antivírus corporativos suspeiam menos de diretórios do que de binários que extraem para temp
- Startup significativamente mais rápido
- Distribuição: compactar como .zip; usuário extrai e executa

---

## ADR-009: PDFs escaneados fora do escopo inicial

**Data:** 2026-04-09
**Status:** Aprovado

**Contexto:**
PDFs escaneados (imagens) não têm texto extraível sem OCR. Tesseract requer instalação separada. markitdown-ocr requer API.

**Decisão:** PDFs escaneados são detectados e recusados com mensagem clara.

**Mensagem ao usuário:** "Este PDF parece ser uma imagem digitalizada. Converta para PDF pesquisável antes de usar, ou forneça os dados em formato Excel."

**Reavaliar quando:** integração com API estiver disponível (markitdown-ocr).

---

## ADR-010: Revisão obrigatória para confidence < 0.95 no fluxo PDF

**Data:** 2026-04-09
**Status:** Aprovado

**Contexto:**
Fuzzy matching textual pode gerar false positives, especialmente em contas com nomes similares ("Receita Operacional" vs "Receita Operacional Líquida").

**Decisão:** O sistema nunca escreve no Spread com confidence < 0.95 sem confirmação do usuário.

**Thresholds:**
- `>= 0.95`: escrita automática
- `0.60 – 0.95`: candidato para revisão (usuário confirma ou rejeita)
- `< 0.60`: descartado com log

**Consequências:**
- GUI do fluxo PDF inclui tela de revisão interativa
- O usuário vê: descrição do PDF, label sugerido no Spread, score, ação (confirmar/rejeitar/trocar)
- Botão de atalho: "Confirmar todos acima de X%"

---

## Decisões Históricas

Migradas de `MEMORIADASIA.md`. Representam decisões tomadas antes da rearquitetura v2.

### ADR-H01: Receita CVM 3.01 sempre vai para "Vendas Mercado Externo"

**Data:** 2026-03-12 (Sessão 2)
**Contexto:** O Spread tem linhas separadas para "Vendas Mercado Interno" e "Vendas Mercado Externo". O CVM 3.01 é a receita total (sem decomposição por mercado).
**Decisão:** CVM 3.01 é sempre gravado na linha "Vendas Mercado Externo", conforme convenção do usuário.
**Validação:** Confirmado manualmente pelo João.

### ADR-H02: Lucro Líquido NÃO é escrito pelo pipeline

**Data:** 2026-03-12 (Sessão 2)
**Contexto:** A linha L195 do Spread calcula Lucro Líquido por fórmula a partir das linhas acima.
**Decisão:** O pipeline nunca escreve nessa linha. Se o LL calculado divergir do CVM, o usuário investiga manualmente.
**Implementação:** Entrada removida de `DRE_SPREAD_MAP`.

### ADR-H03: DFC filtrada por seção 6.01 (atividades operacionais)

**Data:** 2026-03-25 (Sessão 5)
**Contexto:** Bug crítico — regex `amortiza` capturava `6.03.02 "Amortizações"` (pagamento de empréstimo) junto com depreciação real (`6.01.01.06`). Para Sabesp, erro de -85%.
**Decisão:** Filtrar `df_dfc` por `Codigo Conta` começando com `"6.01"` antes do regex.
**Validação:** Sabesp 4T24 e Minerva 4T24 produzem valores corretos.

### ADR-H04: PDFs reativados como fonte de dados

**Data:** 2026-04-09
**Contexto:** Em 2026-03-12 (Sessão 1), extração de PDFs foi descartada pelo João. Em 2026-04-09, a Frente 2 (PDFs) foi reativada como requisito do produto, usando markitdown em vez de abordagem anterior.
**Decisão:** PDFs são uma frente de entrada suportada, com pipeline próprio e revisão obrigatória.

### ADR-H05: Matching híbrido Camada 1 + 2 aprovado

**Data:** 2026-03-25 (Sessão 4)
**Contexto:** O matching original era apenas por valor numérico (frágil). A Camada 1 (CD_CONTA) foi proposta como matching determinístico prioritário.
**Decisão:** Matching híbrido: Layer 2 (valor) prioritário, Layer 1 (código) como fallback. Validado contra Minerva e Sabesp.
**Cobertura Layer 1:** 25+ entradas para BP e DRE. DFC sem Camada 1 (contas não fixas). DMPL com Camada 1 parcial (5.04.01, 5.04.06, 5.04.07).

### ADR-H06: UI deve seguir Design System (a ser fornecido)

**Data:** 2026-04-09
**Contexto:** O produto final é para usuários não técnicos em ambiente corporativo. Visual coerente importa.
**Decisão:** Não investir em UI definitiva antes de receber o Design System. GUI funcional mínima nas phases 0-3. Phase 4 aplica o DS sobre essa base.

---

## Histórico

| Data | Mudança |
|------|---------|
| 2026-04-09 | Criação com 10 ADRs + 6 decisões históricas migradas de MEMORIADASIA.md |
