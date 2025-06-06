# app_spread.py · v1 (Com o git e code agente ChatGPT)
# --------------------------------------------------------------------
# • Coluna-origem / destino por LETRA ou índice
# • Linha inicial global  +  Linha inicial DRE (trimestre)
# • Cabeçalhos corretos para ano × trimestre
# • DRE trimestral: linhas mapeadas manualmente
# • Atualiza planilha ABERTA via xlwings; fallback openpyxl
# • Destaca valores usados na origem tratada
# pip install openpyxl xlwings customtkinter pandas
# pip install -U customtkinter  # se necessário
# test com Python 3.10+ e xlwings >= 0.30.0
# --------------------------------------------------------------------
from __future__ import annotations

import re, sys, tempfile
from pathlib import Path
from typing import Dict, List, Tuple, Callable

import customtkinter as ctk
import pandas as pd
from tkinter import filedialog
from openpyxl import load_workbook
from openpyxl.utils import (
    column_index_from_string as col2idx,
    get_column_letter as idx2col,
)

try:
    import xlwings as xw
    XLWINGS = True
except ImportError:
    XLWINGS = False

# ═════════════ helpers num / período ════════════════════════════════
def normaliza_num(v) -> int | None:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)):
        return int(v)
    if isinstance(v, str):
        s = v.strip().replace(".", "").replace(",", "")
        if re.fullmatch(r"-?\d+", s):
            return int(s)
    return None


def periodos(p: str) -> Tuple[str, str, str, bool]:
    """Devolve (atual, anterior, ante-anterior, is_trimester)."""
    p = p.upper().strip()
    if re.fullmatch(r"\d{4}", p):
        a = int(p)
        return str(a), str(a - 1), str(a - 2), False
    m = re.fullmatch(r"([1-4])T(\d{2})", p)
    if not m:
        raise ValueError("Período deve ser AAAA ou nTAA (ex.: 2024 ou 1T25).")
    tri, aa = int(m.group(1)), int(m.group(2))
    f = lambda y: f"{tri}T{y:02d}"
    return f(aa), f(aa - 1), f(aa - 2), True


def col_txt_to_idx(txt: str) -> int:
    """'A' ➜ 0 · 'AB' ➜ 27 · '0' ➜ 0 …"""
    txt = txt.strip().upper()
    if txt.isdigit():
        return int(txt)
    return col2idx(txt) - 1


# ═══════════ renomeia cabeçalhos origem ═════════════════════════════
# ═══════════ renomeia cabeçalhos + grava origem tratada ════════════
def prepara_origem(
    path: Path,
    tipo: str,
    atual: str,
    ant: str,
    ant2: str,
    is_trim: bool,
    out_dir: Path,              # ← novo parâmetro
) -> Path:
    """
    Cria o arquivo *origem_tratada* (xlsx/xlsm) em `out_dir` com:
      • apenas as abas relevantes;
      • colunas de período renomeadas (ano × trimestre).

    Devolve o Path desse arquivo.
    """
    mapa = {
        "consolidado": {
            "DF Cons Ativo": "cons ativos",
            "DF Cons Passivo": "cons passivos",
            "DF Cons Resultado Periodo": "cons DRE",
            "DF Cons Fluxo de Caixa": "cons DFC",
        },
        "individual": {
            "DF Ind Ativo": "ind ativos",
            "DF Ind Passivo": "ind passivos",
            "DF Ind Resultado Periodo": "ind DRE",
            "DF Ind Fluxo de Caixa": "ind DFC",
        },
    }[tipo]

    # cabeçalhos em minúsculas
    H_ANO      = ("valor ultimo exercicio",
                  "valor penultimo exercicio",
                  "valor antepenultimo exercicio")
    H_TRI_AP   = ("valor trimestre atual",
                  "valor exercicio anterior")
    H_TRI_RES  = ("valor acumulado atual exercicio",
                  "valor acumulado exercicio anterior")

    def make_ren(sheet_orig: str) -> Callable[[str], str]:
        low = sheet_orig.lower()
        is_ap  = any(k in low for k in ("ativo", "passivo"))
        is_res = "resultado" in low

        def ren(col: str) -> str:
            c = col.lower().strip()
            # ---- cabeçalhos trimestre --------------------------------
            if is_trim and is_ap:
                if c.startswith(H_TRI_AP[0]):  return atual
                if c.startswith(H_TRI_AP[1]):  return ant
            elif is_trim and is_res:
                if c.startswith(H_TRI_RES[0]): return atual
                if c.startswith(H_TRI_RES[1]): return ant
            # ---- cabeçalhos ano ---------------------------------------
            if c.startswith(H_ANO[0]):  return atual
            if c.startswith(H_ANO[1]):  return ant
            if c.startswith(H_ANO[2]):  return ant2
            return col
        return ren

    # --------- cria pasta destino, nomeia arquivo --------------------
    out_dir.mkdir(parents=True, exist_ok=True)
    out_name = f"{path.stem}_tratado{path.suffix}"
    out_path = out_dir / out_name

    engine = "openpyxl" if path.suffix.lower() in (".xlsx", ".xlsm") else None
    xls = pd.ExcelFile(path, engine=engine)

    with pd.ExcelWriter(out_path, engine="openpyxl") as wr:
        for aba_orig, aba_nova in mapa.items():
            if aba_orig in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=aba_orig, engine=engine)
                df = df.rename(columns=make_ren(aba_orig))
                df.to_excel(wr, sheet_name=aba_nova, index=False)

    return out_path



# ═══════════ util fórmulas complexas ════════════════════════════════
def shift_formula(f: str, delta: int) -> str:
    pat = re.compile(
        r"(?<![A-Za-z0-9_])(?:'[^']+'|[A-Za-z0-9_]+)?!"
        r"|(?<![A-Za-z0-9_])(\$?)([A-Za-z]{1,3})(?=\$?\d|:)", flags=re.I)
    def repl(m: re.Match) -> str:
        if m.group(1) is None:
            return m.group(0)
        abs_, col = m.group(1), m.group(2)
        try:
            new = idx2col(col2idx(col.upper()) + delta)
        except ValueError:
            new = col
        return f"{abs_}{new}"
    return pat.sub(repl, f)


def adjust_complex_formula(formula: str, delta: int, map_number,
                           used_vals: set[int] | None = None) -> str:
    num_pat = re.compile(r"(?<![A-Za-z])[-+]?\d[\d\.,]*")
    f2 = shift_formula(formula, delta)

    def repl(m: re.Match) -> str:
        n = normaliza_num(m.group(0))
        novo = map_number(n)
        if novo is not None:
            if used_vals is not None:
                used_vals.add(novo)
            return str(novo)
        return m.group(0)

    return num_pat.sub(repl, f2)


# ═══════════ atualizar worksheet (normal) ═══════════════════════════
def valor_corresp(abas: Dict[str, pd.DataFrame], n: int,
                  prev: str, curr: str) -> int | None:
    for df in abas.values():
        if prev not in df.columns or curr not in df.columns:
            continue
        hit = df[df[prev].apply(normaliza_num) == n]
        if not hit.empty:
            return normaliza_num(hit[curr].iloc[0])
    return None


def atualizar_ws(ws,
                 get_val, set_val,
                 abas: Dict[str, pd.DataFrame],
                 src_idx: int, dst_idx: int,
                 atual: str, ant: str,
                 start_row: int) -> tuple[list[int], set[int], set[int]]:
    """
    Copia / ajusta dados da coluna-origem para a coluna-destino
    (ver documentação v17) e devolve:

        • skipped_rows → linhas (1-based) cujo valor numérico ≠ 0 não
          pôde ser gravado na coluna-destino;
        • skipped_vals → conjunto dos números correspondentes.
        • used_vals    → conjunto dos números gravados com sucesso.
    """
    c_src, c_dst = src_idx + 1, dst_idx + 1
    delta = c_dst - c_src
    num_pat = re.compile(r"[-+]?\d[\d\.,]*")

    skipped_rows: list[int] = []
    skipped_vals: set[int] = set()
    used_vals:    set[int] = set()

    empty_streak = 0
    r = start_row
    while empty_streak < 30 and r <= 1_048_576:
        v = get_val(r, c_src)
        if v in (None, ""):
            empty_streak += 1
            r += 1
            continue
        empty_streak = 0

        wrote = False                      # se algo foi realmente escrito
        destino: object

        # ── 1. fórmula apenas com números ---------------------------
        if isinstance(v, str) and v.startswith("=") and not re.search(r"[A-Za-z]", v[1:]):
            def repl(m: re.Match) -> str:
                tok = m.group(0)
                sign = "+" if tok[0] == "+" else "-" if tok[0] == "-" else ""
                n_int = normaliza_num(tok.lstrip("+-"))
                novo  = valor_corresp(abas, n_int, ant, atual)
                if novo is not None:
                    used_vals.add(novo)
                    return sign + str(abs(novo))
                return tok
            destino = "=" + num_pat.sub(repl, v[1:])
            wrote = destino != v

        # ── 2. fórmula complexa -------------------------------------
        elif isinstance(v, str) and v.startswith("="):
            mp = lambda n: valor_corresp(abas, n, ant, atual)
            destino = adjust_complex_formula(v, delta, mp, used_vals)
            wrote = destino != v

        # ── 3. número isolado ---------------------------------------
        elif (n := normaliza_num(v)) is not None:
            novo = valor_corresp(abas, n, ant, atual)
            destino = novo if novo is not None else v
            wrote = novo is not None
            if novo is not None:
                used_vals.add(novo)

        # ── 4. texto / outro tipo -----------------------------------
        else:
            destino = v
            wrote = True                       # sempre copia

        # grava destino na coluna-destino
        try:
            set_val(r, c_dst, destino)
        except Exception:
            try:
                set_val(r, c_dst, get_val(r, c_src))
            except Exception:
                wrote = False

        # se não escreveu e valor numérico ≠ 0, registra
        n_val = normaliza_num(v)
        if (not wrote) and n_val not in (None, 0):
            skipped_rows.append(r)
            skipped_vals.add(n_val)

        r += 1

    return skipped_rows, skipped_vals, used_vals


# ═══════════ mapa de linhas DRE (trimestre) ═════════════════════════
DRE_MAP = {
     0: "Receita de Venda de Bens e/ou Serviços",
    13: "Custo dos Bens e/ou Serviços Vendidos",
    23: "Despesas Gerais e Administrativas",
    25: "Despesas com Vendas",
    26: "Outras Receitas Operacionais",
    27: "Outras Despesas Operacionais",
    29: "Despesas Financeiras",
    30: "Receitas Financeiras",
    41: "Resultado de Equivalência Patrimonial",
    43: "Imposto de Renda e Contribuição Social sobre o Lucro",
}

def aplicar_dre_manual(
    df_dre: pd.DataFrame,
    sheet,                # xlwings.Sheet  OU  openpyxl.Worksheet
    col_dst_1based: int,  # coluna-destino (1-based)
    dre_start: int,       # linha inicial da DRE (1-based)
    col_valor: str,       # nome da coluna “atual” no DataFrame
    is_xlwings: bool,
):
    """
    Copia valores específicos da DRE trimestral, removendo pontos/virgulas
    de milhar e retornos de string.  Usa a tabela DRE_MAP.
    """
    for offset, desc in DRE_MAP.items():
        linha = dre_start + offset
        try:
            raw_val = df_dre.loc[
                df_dre["Descricao Conta"].str.strip() == desc,
                col_valor,
            ].iloc[0]
        except Exception:
            continue                              # descrição não achada

        # ── limpa separadores de milhar ────────────────────────────
        num_val = normaliza_num(raw_val)
        valor   = num_val if num_val is not None else raw_val

        # ── grava na planilha de destino ───────────────────────────
        if is_xlwings:
            sheet.cells(linha, col_dst_1based).value = valor
        else:
            sheet.cell(linha, col_dst_1based, value=valor)

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font


def destacar_inseridos(orig_tratada: Path,
                       used_vals: set[int],
                       atual: str) -> None:
    """
    Realça (fundo verde claro + negrito) todas as células da(s)
    coluna(s) cujo cabeçalho == ``atual`` **e** cujo valor numérico
    está em ``used_vals``. Salva o arquivo no mesmo caminho.
    """
    if not used_vals:
        return  # nada a destacar

    wb = load_workbook(orig_tratada)
    fill = PatternFill("solid", fgColor="CCFFCC")   # verde claro
    bold = Font(bold=True)

    for ws in wb.worksheets:
        atual_cols = [
            cell.column
            for cell in ws[1]
            if str(cell.value).strip() == atual
        ]
        if not atual_cols:
            continue

        for row in ws.iter_rows(min_row=2, values_only=False):
            for c in atual_cols:
                cell = row[c - 1]
                if normaliza_num(cell.value) in used_vals:
                    cell.fill = fill
                    cell.font = bold

    wb.save(orig_tratada)


# ═══════════ pipeline principal (processar) ═════════════════════════
from openpyxl.styles import PatternFill, Font
from openpyxl import load_workbook

def processar(ori: Path, spr: Path, tipo: str,
              periodo: str,
              src_txt: str, dst_txt: str,
              start_row: int,
              dre_start: int,
              out_dir: Path,
              log=lambda _msg: None) -> Path:
    """Processa spread, salva origem tratada e realça valores usados e
    pendentes."""
    src_idx, dst_idx = col_txt_to_idx(src_txt), col_txt_to_idx(dst_txt)
    atual, ant, ant2, is_trim = periodos(periodo)

    # --- cria / grava a origem tratada ----------------------------------
    orig_tratada = prepara_origem(
        ori, tipo, atual, ant, ant2, is_trim, out_dir)

    abas = pd.read_excel(orig_tratada, sheet_name=None, engine="openpyxl")
    dre_sheet = "cons DRE" if tipo == "consolidado" else "ind DRE"
    df_dre = abas.get(dre_sheet)

    # ---------- xlwings --------------------------------------------------
    if XLWINGS and spr.suffix.lower() in {".xlsx", ".xlsm"}:
        try:
            for bk in xw.books:
                if Path(bk.fullname).resolve() == spr.resolve():
                    wb = bk; break
            else:
                wb = xw.Book(str(spr))

            sht = wb.sheets[0]

            get_val = lambda r, c: (
                sht.cells(r, c).formula or sht.cells(r, c).value)
            set_val = lambda r, c, v: (
                sht.cells(r, c).__setattr__("formula" if isinstance(v, str) and v.startswith("=") else "value", v))

            skipped, skipped_vals, used_vals = atualizar_ws(
                sht, get_val, set_val, abas,
                src_idx, dst_idx, atual, ant, start_row)

            if is_trim and df_dre is not None:
                aplicar_dre_manual(df_dre, sht, dst_idx + 1,
                                   dre_start, atual, True)

            wb.app.calculate(); wb.save()
        except Exception as exc:
            log(f"xlwings ⟶ fallback: {exc}")
            skipped, skipped_vals, used_vals = [], set(), set()

    # ---------- fallback openpyxl ---------------------------------------
    else:
        is_xlsm = spr.suffix.lower() == ".xlsm"
        wb = load_workbook(spr, keep_vba=is_xlsm)
        ws = wb.active

        skipped, skipped_vals, used_vals = atualizar_ws(
            ws,
            lambda r, c: ws.cell(r, c).value,
            lambda r, c, v: ws.cell(r, c).__setattr__("value", v),
            abas, src_idx, dst_idx, atual, ant, start_row)

        if is_trim and df_dre is not None:
            aplicar_dre_manual(df_dre, ws, dst_idx + 1,
                               dre_start, atual, False)

        out_name = f"{spr.stem} {atual}{'.xlsm' if is_xlsm else '.xlsx'}"
        spr = spr.with_name(out_name)
        wb.save(spr)

    # ---------- destaca valores pendentes na origem tratada ----------
    destacar_inseridos(orig_tratada, used_vals, atual)

    log(f"Origem tratada em: {orig_tratada}")
    return spr



# sem modificações, pois permanece idêntica.

# ═══════════════════   GUI (CustomTkinter)   ════════════════════════
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Atualizador Spread")
        self.grid_columnconfigure((0, 1), weight=1)

        # ── arquivos --------------------------------------------------------
        self.var_ori  = ctk.StringVar()
        self._campo_arquivo("Arquivo origem", 0, self.var_ori)

        self.var_spr  = ctk.StringVar()
        self._campo_arquivo("Arquivo Spread", 1, self.var_spr)

        # ── pasta para salvar origem tratada --------------------------------
        self.var_outdir = ctk.StringVar(value=str(Path.cwd()))
        self._campo_dir("Pasta origem tratada", 2, self.var_outdir)

        # ── opções ----------------------------------------------------------
        self.var_tipo = ctk.StringVar(value="consolidado")
        ctk.CTkLabel(self, text="Tipo").grid(row=3, column=0, sticky="w", padx=4)
        ctk.CTkOptionMenu(self, variable=self.var_tipo,
                          values=["consolidado", "individual"]
                          ).grid(row=3, column=1, sticky="ew", padx=4)

        self.var_per   = ctk.StringVar()
        self._campo_txt("Período (2024 / 1T25)", 4, self.var_per)

        self.var_src   = ctk.StringVar(value="A")
        self._campo_txt("Coluna origem (A ou 0…)", 5, self.var_src, width=80)

        self.var_dst   = ctk.StringVar(value="B")
        self._campo_txt("Coluna destino", 6, self.var_dst, width=80)

        self.var_start = ctk.StringVar(value="1")
        self._campo_txt("Linha inicial geral", 7, self.var_start, width=80)

        # campo DRE aparece apenas se período for trimestre
        self.var_dre   = ctk.StringVar(value="1")
        self.lbl_dre   = ctk.CTkLabel(self, text="Linha inicial DRE (Receita)")
        self.ent_dre   = ctk.CTkEntry(self, textvariable=self.var_dre, width=80)

        # ── botões ----------------------------------------------------------
        ctk.CTkButton(self, text="Processar", command=self._run
                      ).grid(row=10, column=0, pady=6, padx=4, sticky="ew")
        ctk.CTkButton(self, text="Sair", fg_color="gray",
                      command=self.destroy
                      ).grid(row=10, column=1, pady=6, padx=4, sticky="ew")

        # ── log -------------------------------------------------------------
        self.log = ctk.CTkTextbox(self, width=600, height=150, state="disabled")
        self.log.grid(row=11, column=0, columnspan=3, pady=6, padx=4)

        # monitora período para exibir/esconder campo DRE
        self.var_per.trace_add("write", self._toggle_dre)
        self._toggle_dre()

    # ───────────────────────── helpers UI ───────────────────────────
    def _campo_arquivo(self, rotulo: str, linha: int, var: ctk.StringVar):
        ctk.CTkLabel(self, text=rotulo).grid(row=linha, column=0, sticky="w", padx=4)
        ctk.CTkEntry(self, textvariable=var, width=420
                     ).grid(row=linha, column=1, sticky="ew", padx=4)
        def escolher():
            f = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm *.xls")])
            if f: var.set(f)
        ctk.CTkButton(self, text="…", width=30, command=escolher
                      ).grid(row=linha, column=2, padx=2)

    def _campo_dir(self, rotulo: str, linha: int, var: ctk.StringVar):
        ctk.CTkLabel(self, text=rotulo).grid(row=linha, column=0, sticky="w", padx=4)
        ctk.CTkEntry(self, textvariable=var, width=420
                     ).grid(row=linha, column=1, sticky="ew", padx=4)
        def escolher():
            d = filedialog.askdirectory()
            if d: var.set(d)
        ctk.CTkButton(self, text="…", width=30, command=escolher
                      ).grid(row=linha, column=2, padx=2)

    def _campo_txt(self, rotulo: str, linha: int, var: ctk.StringVar, width=420):
        ctk.CTkLabel(self, text=rotulo).grid(row=linha, column=0, sticky="w", padx=4)
        ctk.CTkEntry(self, textvariable=var, width=width
                     ).grid(row=linha, column=1, sticky="w", padx=4)

    def _toggle_dre(self, *_):
        per = self.var_per.get().strip().upper()
        if re.fullmatch(r"[1-4]T\d{2}", per):
            self.lbl_dre.grid(row=8, column=0, sticky="w", padx=4)
            self.ent_dre.grid(row=8, column=1, sticky="w", padx=4)
        else:
            self.lbl_dre.grid_remove()
            self.ent_dre.grid_remove()

    def _log(self, msg: str):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.configure(state="disabled")
        self.log.see("end")

    # ───────────────────────── run ---------------------------------
    def _run(self):
        try:
            ori  = Path(self.var_ori.get())
            spr  = Path(self.var_spr.get())
            outd = Path(self.var_outdir.get() or Path.cwd())

            if not ori.exists() or not spr.exists():
                self._log("Selecione arquivos válidos."); return

            out_spread = processar(
                ori, spr, self.var_tipo.get(), self.var_per.get(),
                self.var_src.get(), self.var_dst.get(),
                int(self.var_start.get()), int(self.var_dre.get()),
                out_dir=outd,
                log=self._log)

            self._log(f"✔️  Terminado: {out_spread}")
        except Exception as exc:
            self._log(f"Erro: {exc}")



if __name__ == "__main__":
    App().mainloop()