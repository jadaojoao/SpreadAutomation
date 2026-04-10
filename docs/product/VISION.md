# Visão do Produto

> Define o que o Spread Automation deve se tornar, para quem, sob quais restrições, e quais possibilidades futuras devem ser consideradas sem serem implementadas agora.

---

## 1. Objetivo Final

Gerar automaticamente um **Spread Proxy** Excel no formato padronizado usado em análise de crédito.

### A regra mais importante do projeto

- A estrutura final do Spread Proxy é **fixa e imutável**.
- As linhas sempre estarão nas mesmas posições.
- Independentemente da origem dos dados, a saída final deve respeitar exatamente esse layout.
- O sistema preenche essa estrutura com os valores corretos.

### Experiência ideal do usuário final

1. Abrir o .exe
2. Arrastar arquivos ou informar diretórios
3. Escolher algumas opções simples
4. Receber o Spread Proxy pronto

Sem instalar Python, bibliotecas ou dependências.

---

## 2. Perfil do Usuário Final

| Atributo | Descrição |
|----------|-----------|
| Perfil | Analista de crédito em instituição financeira |
| Nível técnico | Não técnico — opera Excel, não programa |
| Ambiente | Windows corporativo com restrições de TI |
| Expectativa | Ferramenta robusta, previsível, que não falha silenciosamente |
| Tolerância a erro | Baixa — prefere que o sistema recuse a processar do que gere resultado errado |

---

## 3. Restrições do Ambiente

Estas restrições são inegociáveis e definem o espaço de solução:

| Restrição | Implicação técnica |
|-----------|-------------------|
| Ambiente corporativo controlado | Sem acesso livre a internet, sem instalação de software arbitrário |
| Sem WebApp | Interface desktop obrigatória |
| Sem API neste momento | Tudo roda localmente; dados vêm de arquivos |
| Sem liberdade de instalação | Solução deve ser .exe standalone |
| Python aceito como base de desenvolvimento | Mas o usuário final nunca vê Python |
| Robustez e previsibilidade são requisitos | Tratamento de erro, validação, logs claros, mensagens úteis |

---

## 4. Frentes de Entrada de Dados

O sistema precisa suportar duas grandes frentes de entrada. Cada frente tem múltiplos modos de operação.

### Frente 1 — ITRs e DFPs (dados estruturados CVM)

Dados padronizados da Comissão de Valores Mobiliários. Formato conhecido, campos definidos, códigos de conta fixos.

| Modo | Descrição | Entrada | Saída |
|------|-----------|---------|-------|
| **1A** | Preenchimento de Spread existente | DadosDocumento.xlsx + Spread Proxy preenchido parcialmente | Spread com novo período adicionado |
| **1B** | Construção do zero | Múltiplos DFPs/ITRs (vários anos) | Spread completo gerado a partir de template vazio |
| **1C** | Origem via CVM Analysis | Estrutura do CVM Analysis (formato a definir) | Spread preenchido |

**Modo 1A** é o fluxo que existe hoje e funciona. O usuário tem um Spread com períodos anteriores e quer adicionar o próximo.

**Modo 1B** é para quando não existe Spread anterior. O usuário quer gerar tudo do zero, escolhendo anos, visão (anual/trimestral/ambas). O sistema valida cobertura temporal, aponta gaps, e gera o Spread final completo.

**Modo 1C** usa como origem a estrutura do CVM Analysis. Essa estrutura será fornecida posteriormente. Ao pensar a arquitetura, considerar que:
- Pode ser a melhor origem-base para o projeto
- Se fizer sentido tecnicamente, deve ser priorizada
- Deve ser tratada como potencial padrão futuro da camada de origem

### Frente 2 — PDFs (dados não estruturados)

PDFs contendo demonstrações financeiras. O cenário mais desafiador: ruído, informações excedentes, baixa padronização.

| Modo | Descrição | Entrada | Saída |
|------|-----------|---------|-------|
| **2A** | Preenchimento a partir de PDFs | PDFs + Spread Proxy existente | Spread preenchido (com revisão) |
| **2B** | Construção do zero a partir de PDFs | PDFs (sem Spread) | Spread gerado a partir de template (com revisão) |

**Vantagem dos PDFs:** praticamente toda companhia possui PDFs disponíveis. Isso torna a solução aplicável a qualquer empresa, não apenas às que têm DadosDocumento.xlsx.

**Restrição:** o fluxo de PDFs exige revisão interativa pelo usuário antes da escrita. Nunca escrever com baixa confiança sem confirmação.

---

## 5. Premissas Técnicas

Estes são princípios de arquitetura, não sugestões:

1. A estrutura final do Spread nunca muda
2. As fontes de entrada podem mudar bastante
3. O sistema deve separar claramente: ingestão, parsing, normalização, mapeamento, validação, preenchimento, exportação
4. O sistema deve ser expansível para novas fontes no futuro
5. O sistema deve ser resiliente a arquivos incompletos, formatos inesperados e inconsistências
6. O sistema deve ter logs claros, rastreabilidade e mensagens úteis para o usuário
7. O sistema deve ser compatível com execução local em ambiente corporativo
8. O design deve facilitar o empacotamento futuro em .exe
9. Não é aceitável uma solução frágil, acoplada ou excessivamente hardcoded sem estratégia
10. Quando houver ambiguidades de mapeamento, o sistema deve ter uma estratégia clara de resolução

---

## 6. Possibilidades Futuras

Registradas na arquitetura, mas **fora do escopo atual de implementação**.

### API de dados CVM
No futuro pode existir integração com uma API que fornece os mesmos dados do DadosDocumento.xlsx de forma programática. Possivelmente a mesma API que alimentará o Excel do CVM Analysis.

**Impacto arquitetural:** a camada de ingestion já é desenhada com adaptadores plugáveis. Um `ingestion/cvm_api.py` seria apenas mais um adaptador que produz `FinancialDataSet`.

### markitdown-ocr para PDFs escaneados
O plugin `markitdown-ocr` (Microsoft) usa LLM Vision para extrair texto de imagens em PDFs. Requer chamada de API.

**Impacto arquitetural:** quando API estiver disponível, o `MarkitdownExtractor` pode ser configurado com o plugin OCR. O restante do pipeline de PDF não muda.

### Processamento em lote
Processar múltiplas empresas de uma vez, gerando vários Spreads em batch.

**Impacto arquitetural:** o `WorkflowEngine` já recebe parâmetros por execução. Um loop externo com progress reporting é suficiente.

---

## Histórico

| Data | Mudança |
|------|---------|
| 2026-04-09 | Criação com visão completa do produto, frentes de entrada, e premissas |
