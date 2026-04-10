# Estrategia de Ingestao de PDFs

> Analise tecnica de MarkItDown vs Docling como tecnologias de extracao para o pipeline PDF (Phase 3 / Modes 2A-2B). Inclui ADR-011, trade-offs, estrategia de POC e criterios de decisao.
>
> Aprovado em: 2026-04-10

---

## ADR-011 — Extrator PDF: MarkItDown como primario, Docling como fallback futuro

**Status:** Aprovado

**Contexto:**
O projeto precisa de uma camada de extracao de PDFs de demonstracoes financeiras para os modos 2A e 2B. O extrator deve ser empacotavel em `.exe` standalone, funcionar offline em ambiente corporativo e produzir output suficientemente estruturado para o parser de dominio financeiro.

**Decisao:**
Usar `markitdown[pdf]` como extrator primario para PDFs nativos. Docling nao entra como dependencia do `.exe` — fica reservado para expansao futura em versao instalada com Python, se o escopo expandir para documentos nao-CVM.

**Consequencias:**
- Bundle `.exe` nao cresce mais que ~70MB por esta dependencia.
- PDFs escaneados continuam fora de escopo (ADR-009).
- O investimento principal vai para o parser de dominio (`ingestion/pdf/parser.py`), nao para o extrator.
- Se MarkItDown produzir <40% de reconhecimento em PDFs reais de DFP, a decisao sera revisada antes da Phase 3.

**Alternativas rejeitadas:**
- Docling como primario: inviavel para `.exe` (~500MB de modelos ONNX, download em runtime).
- pdfplumber direto: interface mais complexa para texto corrido; ja foi testado com resultado insatisfatorio.
- PyMuPDF/fitz: boa alternativa tecnica, mas cobre menos casos de uso que MarkItDown (sem Excel, DOCX).

---

## 1. Premissa Critica

O problema central do pipeline PDF nao e a conversao do arquivo — e a **extracao estruturada de tabelas numericas**. Converter um PDF para Markdown e a parte facil. Reconstituir linhas de balanco com os valores corretos, na escala certa, sem perder hierarquia de contas, e onde 80% do trabalho esta.

Nenhuma ferramenta generica resolve isso automaticamente. O extrator (MarkItDown ou Docling) determina a qualidade do material de entrada. O **parser de dominio** (`ingestion/pdf/parser.py`) e o componente que realmente resolve o problema.

---

## 2. O que o MarkItDown faz e nao faz

### Faz
- Converte PDFs nativos (texto selecionavel) para Markdown via pdfplumber interno
- Representa tabelas em pipe format (`| col | col |`)
- Trata multiplos formatos com interface unificada (PDF, Excel, DOCX, PPTX)
- Funciona offline apos instalacao
- API simples: `MarkItDown().convert(path).text_content`

### Nao faz
- Nao entende estrutura de documentos financeiros
- Nao garante alinhamento de colunas em tabelas complexas do PDF
- Nao lida com merge de celulas, headers multi-linha, subtotais indentados
- Nao detecta hierarquia de contas (1.01, 1.01.01)
- Nao interpreta negativos em parenteses `(1.234)` vs sinal `-1.234`
- Nao sabe nada de escala monetaria (Mil vs Unidade)

### Relacao com pdfplumber
MarkItDown usa pdfplumber internamente para PDFs. Se o problema anterior com pdfplumber era a qualidade da extracao de tabelas via `extract_table()`, o MarkItDown provavelmente produz resultado similar ou pior, pois nao usa `extract_table()` — usa extracao de texto com tentativa de reconstrucao de tabelas.

**O MarkItDown nao e um substituto de pdfplumber — e uma camada acima dele.** Para tabelas numericas precisas, `pdfplumber.extract_table()` com configuracao fina pode superar o Markdown gerado pelo MarkItDown.

---

## 3. O que o Docling faz e nao faz

### Faz
- Usa modelos de ML (layout analysis via DocLayNet, table structure via TableFormer/ONNX)
- Detecta celulas com coordenadas, headers multi-linha, hierarquia visual
- Exporta para JSON estruturado com metadados de posicao
- Qualidade de extracao de tabelas significativamente superior em PDFs complexos
- Suporta exportacao para Markdown, HTML, JSON, DoclingDocument

### Nao faz
- Nao e rapido: 10-60s por pagina em CPU para PDFs complexos
- Nao e leve: ~500MB de modelos ONNX (DocLayNet + TableFormer)
- Nao funciona offline sem pre-download dos modelos
- Nao conhece convencoes contabeis brasileiras
- Nao resolve PDFs escaneados sem OCR adicional

### Por que Docling e inviavel como primario agora

| Restricao | Impacto |
|-----------|---------|
| Modelos ONNX no bundle | `.exe` passaria de 600MB (meta: <150MB) |
| Download em runtime | Falha em ambiente corporativo sem internet |
| Antivirus corporativo | Modelos ONNX sao vetor de falso positivo |
| Velocidade | 10-60s/pagina e inaceitavel para uso operacional |
| Complexidade de empacotamento | Requer engenharia significativa de PyInstaller |

---

## 4. Comparacao Objetiva

| Dimensao | MarkItDown | Docling |
|----------|-----------|---------|
| Qualidade extracao tabela PDF | Baixa-media | Media-alta |
| Velocidade | Rapida (<5s/PDF tipico) | Lenta (10-60s/pagina em CPU) |
| Tamanho de dependencias | ~50-70MB | ~500MB+ |
| Empacotamento .exe | Viavel | Inviavel (curto prazo) |
| Saida estruturada | Markdown (texto) | JSON com coordenadas de celula |
| Funcionamento offline | Sim | Requer pre-download de modelos |
| PDFs escaneados | Nao | Nao (sem OCR adicional) |
| Convencoes BR (1.234.567) | Nao sabe | Nao sabe |
| Hierarquia contabil CVM | Nao sabe | Nao sabe |
| Maturidade producao | Alta | Media (v2+, em crescimento) |

---

## 5. Markdown como Formato Intermediario

Markdown e o formato de transporte adequado para **texto corrido** (notas explicativas, comentarios de gestao, secoes narrativas). Para **tabelas numericas financeiras**, Markdown e um formato intermediario ruim:

- Pipe tables nao preservam merge de celulas
- Valores como `1.234.567` podem ser interpretados como separadores de coluna
- Indentacao hierarquica de contas se perde
- Sem metadados de alinhamento numerico

**O formato intermediario correto para tabelas e uma lista de dicts:**

```python
# Resultado esperado do parser, independente do extrator usado
[
    {
        "code": None,           # PDFs nao tem CD_CONTA
        "description": "Receita de Venda de Bens e/ou Servicos",
        "value": 4823456789,    # ja normalizado para inteiro
        "period": "2024",
        "section": "DRE",
        "source": "pdf",
        "confidence": 0.92
    },
    ...
]
```

**Regra pratica:** MarkItDown (ou Docling) entrega Markdown/JSON. O `parser.py` converte imediatamente para lista de dicts. Markdown nao deve persistir alem da camada de extracao.

---

## 6. Arquitetura do Pipeline PDF

```
PDF Input (nativo, texto selecionavel)
    │
    ▼
[ingestion/pdf/extractor.py]
    │
    ├─ MarkItDown().convert(path).text_content
    │  → Raw Markdown
    │
    └─ Detecta PDF escaneado (output vazio ou <100 chars)
       → Recusa com mensagem clara (ADR-009)
    │
    ▼
[ingestion/pdf/parser.py]   ← investimento principal
    │
    ├─ Identifica secoes (## Ativo, ## DRE, etc.)
    ├─ Extrai linhas de tabela do Markdown
    ├─ Detecta colunas de periodo (header da tabela)
    ├─ Filtra linhas de subtotal/cabecalho/rodape
    └─ → List[Dict{description, raw_value, period, section}]
    │
    ▼
[ingestion/pdf/normalizer.py]
    │
    ├─ "1.234.567" → 1234567       (inteiro BR)
    ├─ "(1.234)" → -1234           (negativo entre parenteses)
    ├─ "R$ 1.234 mil" → 1234000    (escala explicita)
    ├─ Detecta ESCALA_MOEDA (heuristica)
    └─ Normaliza texto para fuzzy (unidecode, lowercase)
    │
    ▼
[ingestion/pdf/adapter.py]
    │
    └─ → FinancialDataSet(
            source_type="pdf",
            accounts=[FinancialAccount(code=None, confidence=X)]
         )
    │
    ▼
[mapping/mapper.py]
    │
    ├─ Layer 3 (fuzzy): description vs rotulos do Spread + account_synonyms.json
    └─ Layer 2 (valor): se Spread tem periodo anterior
    │
    ▼
[SpreadWriter] → Spread Proxy preenchido
```

### Fallback quando extracao falha

```
MarkItDown falha ou producao < threshold
    │
    ├─ Tentativa 1: pdfplumber.extract_table() direto
    │   (melhor para tabelas bem delimitadas)
    │
    ├─ Tentativa 2: pdfplumber.extract_text() por pagina
    │   (para layouts de texto corrido)
    │
    └─ Fallback final: rejeitar com relatorio de falha
    │   "PDF nao processavel automaticamente.
    │    Insira os dados manualmente ou use DadosDocumento.xlsx."
```

---

## 7. Prova de Conceito (POC)

### Objetivo
Descobrir em uma tarde se MarkItDown produz Markdown "parseable" para DFPs brasileiras reais.

### Script de validacao

```python
from markitdown import MarkItDown
from pathlib import Path
import re

def poc_pdf(pdf_path: str):
    md = MarkItDown()
    result = md.convert(pdf_path)
    content = result.text_content

    print(f"[1] Chars extraidos: {len(content)}")

    # Tabelas em pipe format
    table_lines = [l for l in content.split('\n') if l.strip().startswith('|')]
    print(f"[2] Linhas de tabela: {len(table_lines)}")

    # Numeros BR reconheciveis
    br_numbers = re.findall(r'\d{1,3}(?:\.\d{3})+(?:,\d+)?', content)
    print(f"[3] Numeros BR: {len(br_numbers)} | Exemplos: {br_numbers[:5]}")

    # Termos CVM presentes
    termos = ["Receita", "Custo", "Ativo", "Passivo", "Patrimonio", "Caixa",
              "Resultado", "Despesa", "Financeiro", "Imposto"]
    encontrados = [t for t in termos if t in content]
    print(f"[4] Termos CVM: {len(encontrados)}/10 — {encontrados}")

    # Salva para inspecao visual
    Path("poc_output.md").write_text(content, encoding="utf-8")
    print("[5] Salvo em poc_output.md — inspecione as tabelas visualmente")

poc_pdf("caminho/para/seu_dfp.pdf")
```

### O que observar no `poc_output.md`

1. **Tabelas aparecem como pipe tables ou texto corrido?**
   Se texto corrido: o parser precisara de regex mais agressivo, nao extrai_table.

2. **Valores numericos estao na mesma linha da descricao?**
   Se colunas misturadas: alinhamento perdido, problema serio.

3. **Hierarquia (1.01, 1.01.01) aparece?**
   CVM PDFs geralmente omitem o codigo — so descricao. Esperado.

4. **Garbage text?**
   Numeros de pagina, headers repetidos, rodapes. Filtragem necessaria no parser.

5. **Texto acentuado intacto?**
   Abrir output.md com UTF-8. "Receita de Venda de Bens e/ou Servicos" deve aparecer completo.

### Criterio de decisao rapido

| Resultado do POC | Acao |
|-----------------|------|
| >60% das contas CVM reconheciveis + valores na mesma linha | Prosseguir com MarkItDown |
| Tabelas presentes mas colunas desalinhadas | Investigar parser + pdfplumber.extract_table() como complemento |
| Texto extraido mas sem estrutura de tabela | Usar pdfplumber.extract_table() direto em vez de MarkItDown |
| <30% de reconhecimento | Testar Docling como comparativo (sem .exe ainda) |
| Output vazio ou <100 chars | PDF escaneado — fora de escopo (ADR-009) |

---

## 8. Criterios Objetivos de Decisao Final

### Seguir com MarkItDown como primario

- POC mostra ≥60% das contas CVM reconheciveis em PDFs nativos de DFP
- Bundle `.exe` com `markitdown[pdf]` fica abaixo de 120MB
- Velocidade de conversao <10s para um DFP tipico (50-100 paginas)
- Normalizer consegue parsear ≥80% dos valores numericos corretamente

### Adicionar Docling como modo alternativo (nao padrao)

- PDFs de laminas/prospectos tem layout mais complexo que DFPs (multiplas colunas, tabelas aninhadas)
- Projeto for distribuido como instalacao Python gerenciada (nao .exe)
- Resultado do MarkItDown ficar <40% apos melhoria no parser

### Usar pdfplumber.extract_table() diretamente

- MarkItDown produz texto sem estrutura de tabela (<20% pipe table lines)
- Tabelas do PDF tem bordas bem definidas (caso comum em DFPs gerados por sistemas)
- Resultado superior ao Markdown na inspecao visual do POC

### Nao usar nenhum dos dois

- PDFs do cliente forem majoritariamente escaneados
- Cobertura esperada do modo PDF for <30% das linhas do Spread
- Custo de manutencao do parser de dominio superar o ganho de cobertura

---

## 9. O que Nao Fazer

- **Nao usar `markitdown[all]` no `.exe`** — instala dependencias de audio, imagem, LLM irrelevantes. Use `markitdown[pdf]`.
- **Nao usar Docling antes de validar MarkItDown** — o custo de integracao e alto; valide a premissa mais barata primeiro.
- **Nao tratar Markdown como formato final de tabelas** — converta para lista de dicts imediatamente no parser.
- **Nao construir pipeline completo antes do POC** — o POC leva uma tarde; o pipeline completo leva dias.
- **Nao confiar que o extrator resolve a hierarquia contabil** — PDFs nao tem CD_CONTA, nunca terao. Layer 3 fuzzy e mandatoria.

---

## Historico

| Data | Mudanca |
|------|---------|
| 2026-04-10 | Criacao com analise completa, ADR-011, POC, criterios de decisao |
