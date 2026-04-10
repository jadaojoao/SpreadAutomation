# spread_automation

Automação de Spreads financeiros a partir de demonstrações CVM (DFP/ITR).
Compara o período anterior, mapeia valores e atualiza a planilha de Spread com destaques visuais.

## Como Funciona

1. Baixe o arquivo ZIP da CVM (DFP ou ITR)
2. Renomeie para `[Empresa] [Período].zip` (ex.: `Minerva 4T24.zip`)
3. Extraia na pasta `data/` → ficará `data/Minerva 4T24/`
4. Dentro da pasta, garanta que existam:
   - `DadosDocumento.xlsx` — arquivo Origem
   - `Spread Proxy.xlsx` — arquivo Spread (aba "Entrada de Dados")
5. Execute `python main.py`
6. Selecione os arquivos, configure período e colunas, clique **Processar**

## Estrutura do Projeto

```
spread_automation/
├── core/                   # Funções utilitárias e constantes
│   ├── __init__.py
│   └── utils.py
├── processing/             # Pipeline principal de Spread
│   ├── __init__.py
│   ├── origin.py           # Normalização do DadosDocumento.xlsx
│   ├── spread.py           # Varredura e mapeamento de valores
│   ├── dre.py              # DRE trimestral
│   ├── dfc.py              # Depreciação/amortização
│   ├── dmpl.py             # Dividendos, JCP, aumentos de capital
│   ├── highlights.py       # Destaques visuais
│   └── pipeline.py         # Orquestrador
├── app/                    # Interface gráfica
│   ├── __init__.py
│   └── gui.py
├── data/                   # Dados de trabalho
│   └── Minerva 4T24/       # Exemplo: empresa + período
│       ├── DadosDocumento.xlsx
│       └── Spread Proxy.xlsx
├── main.py                 # Ponto de entrada
├── requirements.txt
├── CONTEXT.md
└── README.md
```

## Instalação

```bash
cd spread_automation
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Uso

```bash
python main.py
```

Na GUI:
- **Arquivo Origem** → selecione `DadosDocumento.xlsx`
- **Arquivo Spread** → selecione `Spread Proxy.xlsx`
- **Tipo** → Consolidado ou Individual
- **Período** → ano (`2024`) ou trimestre (`1T25`, `4T24`)
- **Coluna Origem / Destino** → colunas no Spread para ler e escrever

## Desenvolvimento

### Mapeamento DRE Trimestral (ITR)

Para ITR, a DRE usa **label-based matching**: o sistema escaneia a coluna B do Spread
e insere cada conta CVM na linha cujo rótulo corresponde ao definido em `DRE_SPREAD_MAP`
(`core/utils.py`). Múltiplas contas CVM podem mapear para o mesmo rótulo (valores somados).

Para adicionar ou ajustar o mapeamento DRE (ex: nova conta ou rótulo diferente no Spread),
edite apenas `DRE_SPREAD_MAP` em `core/utils.py`.

### Adicionando um Novo Demonstrativo

1. Crie módulo em `processing/` (ex.: `processing/dva.py`)
2. Importe e integre em `processing/pipeline.py`

### Adicionando uma Nova Empresa

1. Baixe o ZIP da CVM
2. Renomeie para `[Empresa] [Período].zip`
3. Extraia em `data/`
4. Garanta que `DadosDocumento.xlsx` e `Spread Proxy.xlsx` existam

### Boas Práticas

- Imports via pacote: `from core.utils import normaliza_num`
- Um módulo por demonstrativo em `processing/`
- Constantes centralizadas em `core/utils.py`

## Colaboração

- Guia de contribuição: `CONTRIBUTING.md`
- Histórico de mudanças: `CHANGELOG.md`
- Licença: `LICENSE`

## Referências de Projeto

- Contexto técnico e arquitetura: `CONTEXT.md`
- Memória de decisões e validações: `MEMORIADASIA.md`

## Governança de Trabalho

- Contrato operacional: `AGENTS.md`
- Configuração de governança: `.github/governance.config.yaml`
- Política de paths e riscos: `.github/guardrails/path-policy.json`
- Runbook operacional: `docs/governance/operators-runbook.md`
- Recuperação e rollback: `docs/governance/rollback-recovery.md`
- Gate de conclusão: tarefa só pode ser dada como concluída após merge do PR e fechamento da issue vinculada.

Scripts úteis:
- Bootstrap dry-run/apply: `scripts/bootstrap-governance.ps1` e `scripts/bootstrap-governance.sh`
- Worktree create/status/remove: `scripts/worktree_create.ps1`, `scripts/worktree_status.ps1`, `scripts/worktree_remove.ps1`
- Merge operacional de PR: `scripts/pr_complete.ps1`
