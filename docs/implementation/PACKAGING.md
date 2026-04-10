# Estrategia de Empacotamento .exe

> Diretrizes para gerar um executavel Windows distribuivel com PyInstaller, sem dependencia de Python instalado.

---

## 1. Decisao Arquitetural

**PyInstaller --onedir** (ADR-008). Escolhido por:
- Startup mais rapido que --onefile (nao descompacta em temp a cada execucao)
- Facilita debug (DLLs e recursos visiveis na pasta)
- Atualizacoes parciais possiveis (substituir apenas o que mudou)
- Compativel com antivirus corporativo (menos falsos positivos)

---

## 2. Recursos Incluidos no Bundle

Todos os recursos devem ser acessiveis via `core/resources.py` → `get_resource_path()`, que detecta `sys._MEIPASS` (frozen) vs `__file__` (desenvolvimento).

### Arquivos de dados

| Recurso | Path relativo | `--add-data` |
|---------|--------------|--------------|
| `spread_schema.json` | `./spread_schema.json` | `spread_schema.json;.` |
| `conta_spread_map.json` | `./mapping_tables/conta_spread_map.json` | `mapping_tables;mapping_tables` |
| `dre_spread_map.json` | `./mapping_tables/dre_spread_map.json` | (incluso acima) |
| `account_synonyms.json` | `./mapping_tables/account_synonyms.json` | (incluso acima) |
| Template Spread | `./templates/Spread Proxy Template.xlsx` | `templates;templates` |
| Themes/assets | `./themes/*` | `themes;themes` |

### Hidden imports

Modulos que PyInstaller nao detecta automaticamente:

| Pacote | Hidden imports | Razao |
|--------|---------------|-------|
| `customtkinter` | `customtkinter` | Importa recursos via `__file__` |
| `openpyxl` | `openpyxl.cell._writer` | Import condicional |
| `markitdown` | `markitdown`, `pdfplumber` | Plugin lazy-load |
| `rapidfuzz` | `rapidfuzz.fuzz`, `rapidfuzz.process` | C extensions |
| `tkinterdnd2` | `tkinterdnd2` | DLL nativa |

### Exclusoes explicitas

| Pacote | Razao | Economia estimada |
|--------|-------|-------------------|
| `xlwings` | Requer Excel instalado; nao funciona em .exe standalone | ~15 MB |
| `pytest` | Dev-only | ~5 MB |
| `pip`, `setuptools` | Dev-only | ~10 MB |
| `numpy` (se possivel) | Avaliar se pandas precisa | ~30 MB |

---

## 3. Esboco do `build.spec`

```python
# build.spec — PyInstaller spec para Spread Automation
# Phase 0: esboco inicial
# Phase 5: finalizado com todos os recursos

import sys
from pathlib import Path
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# Coleta recursos do customtkinter
ctk_data = collect_data_files('customtkinter')

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('spread_schema.json', '.'),
        ('mapping_tables', 'mapping_tables'),
        ('templates', 'templates'),
        ('themes', 'themes'),
    ] + ctk_data,
    hiddenimports=[
        'customtkinter',
        'openpyxl.cell._writer',
        'rapidfuzz.fuzz',
        'rapidfuzz.process',
        # markitdown e tkinterdnd2 adicionados na Phase 3/4
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'xlwings',
        'pytest',
        'pip',
        'setuptools',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='SpreadAutomation',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,          # GUI, sem console
    icon='themes/icon.ico', # Icone do app (Phase 4)
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='SpreadAutomation',
)
```

---

## 4. `core/resources.py` — Padrao de Acesso

```python
import sys
from pathlib import Path

def get_resource_path(relative_path: str) -> Path:
    """Retorna path absoluto para recurso, compativel com PyInstaller."""
    if getattr(sys, 'frozen', False):
        # Executando como .exe — recursos estao em sys._MEIPASS
        base = Path(sys._MEIPASS)
    else:
        # Executando como script — recursos relativos ao projeto
        base = Path(__file__).resolve().parent.parent
    return base / relative_path
```

Todos os modulos que acessam JSONs, templates ou assets devem usar `get_resource_path()`. Nunca `Path(__file__).parent / "..."` diretamente.

---

## 5. Dependencias e Tamanho do Bundle

### Dependencias obrigatorias

| Pacote | Versao minima | Tamanho estimado | Fase |
|--------|--------------|------------------|------|
| `customtkinter` | 5.2 | ~5 MB | Phase 0 |
| `openpyxl` | 3.1 | ~8 MB | Phase 0 |
| `pandas` | 2.0 | ~50 MB | Phase 0 |
| `markitdown[pdf]` | latest | ~15 MB | Phase 3 |
| `rapidfuzz` | 3.0 | ~5 MB | Phase 3 |
| `tkinterdnd2` | 0.3 | ~2 MB | Phase 4 |
| Python runtime | 3.11+ | ~35 MB | — |

### Meta de tamanho

**Bundle total < 150 MB** (definicao de pronto Phase 5).

### Estrategia de reducao

1. **UPX** — comprimir binarios (ativado no spec)
2. **Excludes** — remover xlwings, pytest, pip, setuptools
3. **Avaliar pandas** — se unico uso for `read_excel`, substituir por `openpyxl` direto
4. **`--strip`** — remover simbolos de debug (apenas em release)

---

## 6. Riscos e Mitigacoes

| Risco | Probabilidade | Impacto | Mitigacao |
|-------|--------------|---------|-----------|
| Windows Defender bloqueia .exe | Media | Alto | UPX off em release; code signing se disponivel |
| customtkinter nao encontra assets | Alta | Alto | `collect_data_files('customtkinter')` no spec |
| tkinterdnd2 DLL nao carrega | Media | Medio | Receita PyInstaller conhecida; fallback para botao |
| Paths com acentos falham | Media | Alto | Usar `Path` objects; testar com `C:\Users\Joao\` |
| Bundle > 150 MB | Media | Baixo | Avaliar remocao de pandas; UPX |
| markitdown com dependencias pesadas | Baixa | Medio | Medir apos Phase 3; considerar lazy import |

---

## 7. Checklist de Teste

### Ambiente de build
- [ ] `pyinstaller build.spec` completa sem erros
- [ ] Pasta `dist/SpreadAutomation/` contem todos os recursos
- [ ] `spread_schema.json` acessivel via `get_resource_path()`
- [ ] `mapping_tables/*.json` acessiveis
- [ ] Template Spread acessivel
- [ ] Themes/assets carregam

### Maquina limpa (VM sem Python)
- [ ] .exe abre GUI sem erros
- [ ] Mode 1A: processa Minerva 4T24 com output correto
- [ ] Mode 1B: constroi Spread do zero (se Phase 2 concluida)
- [ ] Mode 2A: processa PDF (se Phase 3 concluida)
- [ ] Drag-and-drop funciona (se Phase 4 concluida)
- [ ] Nenhum traceback visivel ao usuario

### Ambiente corporativo
- [ ] Funciona atras de proxy corporativo
- [ ] Antivirus nao bloqueia
- [ ] Funciona em diretorios restritos (ex.: `C:\Users\...`)
- [ ] Funciona com nomes de usuario acentuados
- [ ] Funciona sem permissao de administrador

---

## 8. Distribuicao

### Estrutura do pacote distribuivel

```
SpreadAutomation/
├── SpreadAutomation.exe       # Executavel principal
├── _internal/                 # Recursos e dependencias (PyInstaller)
│   ├── spread_schema.json
│   ├── mapping_tables/
│   ├── templates/
│   ├── themes/
│   └── ...                    # DLLs, .pyd, etc.
└── README.txt                 # Instrucoes para o usuario (1 pagina)
```

### Instrucoes de uso (README.txt)

1. Extraia a pasta `SpreadAutomation/` para qualquer local
2. Execute `SpreadAutomation.exe`
3. Selecione os arquivos e configure o processamento
4. Nao mova ou renomeie arquivos dentro de `_internal/`

### Atualizacoes

Sem auto-update. Distribuicao manual:
1. Gere novo build
2. Comprima em .zip
3. Envie ao usuario
4. Usuario substitui a pasta antiga pela nova

---

## Historico

| Data | Mudanca |
|------|---------|
| 2026-04-09 | Criacao com estrategia PyInstaller, spec esboco, riscos e checklist |
