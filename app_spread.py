# app_spread.py · v2 (Com o git e code agente ChatGPT)
# --------------------------------------------------------------------
# • Coluna-origem / destino por LETRA ou índice
# • Linha inicial global  +  Linha inicial DRE (trimestre)
# • Cabeçalhos corretos para ano × trimestre
# • DRE trimestral: linhas mapeadas manualmente
# • Atualiza planilha ABERTA via xlwings; fallback openpyxl
# • Destaca valores usados na origem tratada
# • Depreciação/Amortização na DFC trimestral negativa
# pip install openpyxl xlwings customtkinter pandas
# pip install -U customtkinter  # se necessário
# test com Python 3.10+ e xlwings >= 0.30.0
# --------------------------------------------------------------------
import logging
import re
import sys
import tempfile
from pathlib import Path
from tkinter import filedialog
from typing import Callable, Dict, List, Tuple

import customtkinter as ctk
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import (
    column_index_from_string as col2idx,
    get_column_letter as idx2col,
)

try:
    import xlwings as xw

    XLWINGS = True
except ImportError:
    XLWINGS = False


def normaliza_num(v) -> int | None:
    """Normalize a string or number to an integer or None."""
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
    """Return (current, previous, pre-previous, is_trimester)."""
    p = p.upper().strip()
    if re.fullmatch(r"\d{4}", p):
        a = int(p)
        return str(a), str(a - 1), str(a - 2), False
    m = re.fullmatch(r"([1-4])T(\d{2})", p)
    if not m:
        raise ValueError("Period must be YYYY or QTYY (e.g., 2024 or 1T25).")
    tri, aa = int(m.group(1)), int(m.group(2))
    f = lambda y: f"{tri}T{y:02d}"
    return f(aa), f(aa - 1), f(aa - 2), True


def col_txt_to_idx(txt: str) -> int:
    """'A' -> 0 | 'AB' -> 27 | '0' -> 0 ..."""
    txt = txt.strip().upper()
    if txt.isdigit():
        return int(txt)
    return col2idx(txt) - 1


def prepara_origem(
    path: Path,
    tipo: str,
    atual: str,
    ant: str,
    ant2: str,
    is_trim: bool,
    out_dir: Path | None,
) -> Path:
    """
    Create <file>_tratado.xlsx (or .xlsm) without overwriting the
    original file and return the generated Path.
    """
    dst_dir = out_dir or path.parent
    dst_dir.mkdir(parents=True, exist_ok=True)
    dst_path = dst_dir / f"{path.stem}_tratado{path.suffix}"

    sheet_map = {
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

    H_ANO = (
        "valor ultimo exercicio",
        "valor penultimo exercicio",
        "valor antepenultimo exercicio",
    )
    H_TRI_AP = ("valor trimestre atual", "valor exercicio anterior")
    H_TRI_RES = (
        "valor acumulado atual exercicio",
        "valor acumulado exercicio anterior",
    )

    def ren_factory(sheet_orig: str) -> Callable[[str], str]:
        low = sheet_orig.lower()
        is_ap = any(k in low for k in ("ativo", "passivo"))
        is_res = "resultado" in low

        def ren(c: str) -> str:
            c_low = c.lower().strip()
            if is_trim and is_ap:
                if c_low.startswith(H_TRI_AP[0]):
                    return atual
                if c_low.startswith(H_TRI_AP[1]):
                    return ant
            elif is_trim and is_res:
                if c_low.startswith(H_TRI_RES[0]):
                    return atual
                if c_low.startswith(H_TRI_RES[1]):
                    return ant
            if c_low.startswith(H_ANO[0]):
                return atual
            if c_low.startswith(H_ANO[1]):
                return ant
            if c_low.startswith(H_ANO[2]):
                return ant2
            return c

        return ren

    engine = "openpyxl" if path.suffix.lower() in (".xlsx", ".xlsm") else None
    xls = pd.ExcelFile(path, engine=engine)

    with pd.ExcelWriter(dst_path, engine="openpyxl") as wr:
        for orig, novo in sheet_map.items():
            if orig not in xls.sheet_names:
                continue
            df = pd.read_excel(xls, sheet_name=orig, engine=engine)
            df = df.rename(columns=ren_factory(orig))
            df.to_excel(wr, sheet_name=novo, index=False)

    wb = load_workbook(dst_path)
    if wb.sheetnames:
        wb[wb.sheetnames[0]].sheet_state = "visible"
    wb.save(dst_path)

    return dst_path


def shift_formula(f: str, delta: int) -> str:
    """Shift column references in a formula by a given delta."""
    pat = re.compile(
        r"(?<![A-Za-z0-9_])(?:'[^']+'|[A-Za-z0-9_]+)?!"
        r"|(?<![A-Za-z0-9_])(\$?)([A-Za-z]{1,3})(?=\$?\d|:)",
        flags=re.I,
    )

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


def adjust_complex_formula(
    formula: str, delta: int, map_number, used_vals: set[int] | None = None
) -> str:
    """Adjust a complex formula by shifting columns and mapping numbers."""
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


def valor_corresp(
    abas: Dict[str, pd.DataFrame], n: int, prev: str, curr: str
) -> int | None:
    """Find the corresponding value for a number in the source sheets."""
    for df in abas.values():
        if prev not in df.columns or curr not in df.columns:
            continue
        hit = df[df[prev].apply(normaliza_num) == n]
        if not hit.empty:
            return normaliza_num(hit[curr].iloc[0])
    return None


def atualizar_ws(
    ws,
    get_val,
    set_val,
    abas: Dict[str, pd.DataFrame],
    src_idx: int,
    dst_idx: int,
    atual: str,
    ant: str,
    start_row: int,
) -> tuple[list[int], set[int], set[int]]:
    """
    Copy/adjust data from source to destination column and return skipped rows/values
    and used values.
    """
    c_src, c_dst = src_idx + 1, dst_idx + 1
    delta = c_dst - c_src
    skipped_rows, skipped_vals, used_vals = [], set(), set()
    empty_streak, r = 0, start_row

    while empty_streak < 30 and r <= 1_048_576:
        v = get_val(r, c_src)
        if v in (None, ""):
            empty_streak += 1
            r += 1
            continue
        empty_streak = 0
        wrote, destino = False, v

        if isinstance(v, str) and v.startswith("="):
            if not re.search(r"[A-Za-z]", v[1:]):
                def repl(m):
                    n = valor_corresp(
                        abas, normaliza_num(m.group(0).lstrip("+-")), ant, atual
                    )
                    if n is not None:
                        sign = ""
                        if m.group(0)[0] in "+-":
                            sign = m.group(0)[0]
                        return f"{sign}{abs(n)}"
                    return m.group(0)

                destino = "=" + re.sub(r"[-+]?\d[\d\.,]*", repl, v[1:])
                wrote = destino != v
            else:
                mp = lambda n: valor_corresp(abas, n, ant, atual)
                destino = adjust_complex_formula(v, delta, mp, used_vals)
                wrote = destino != v
        elif (n := normaliza_num(v)) is not None:
            novo = valor_corresp(abas, n, ant, atual)
            if novo is not None:
                destino, wrote = novo, True
                used_vals.add(novo)
        else:
            wrote = True

        try:
            set_val(r, c_dst, destino)
        except Exception:
            wrote = False

        if not wrote and normaliza_num(v) not in (None, 0):
            skipped_rows.append(r)
            skipped_vals.add(normaliza_num(v))
        r += 1
    return skipped_rows, skipped_vals, used_vals


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
    sheet,
    col_dst_1based: int,
    dre_start: int,
    col_valor: str,
    is_xlwings: bool,
):
    """Copy specific values from the quarterly DRE using DRE_MAP."""
    for offset, desc in DRE_MAP.items():
        linha = dre_start + offset
        try:
            raw_val = df_dre.loc[
                df_dre["Descricao Conta"].str.strip() == desc, col_valor
            ].iloc[0]
            valor = normaliza_num(raw_val) or raw_val
            if is_xlwings:
                sheet.cells(linha, col_dst_1based).value = valor
            else:
                sheet.cell(linha, col_dst_1based, value=valor)
        except (IndexError, KeyError):
            continue


def inserir_depreciacao_dfc(
    df_dfc: pd.DataFrame,
    sheet,
    col_dst_1based: int,
    linha: int,
    col_valor: str,
    is_xlwings: bool,
) -> int | None:
    """Read Depreciation/Amortization from DFC and write it as a negative value."""
    if df_dfc is None or col_valor not in df_dfc.columns:
        return None
    desc = df_dfc["Descricao Conta"].astype(str)
    mask = desc.str.contains("deprecia|amortiza", case=False, na=False)
    try:
        raw_val = df_dfc.loc[mask, col_valor].iloc[0]
    except (IndexError, KeyError):
        return None

    num_val = normaliza_num(raw_val)
    valor = -abs(num_val) if num_val is not None else f"-{str(raw_val).lstrip('+-')}"
    num_val = normaliza_num(valor)

    if is_xlwings:
        sheet.cells(linha, col_dst_1based).value = valor
    else:
        sheet.cell(linha, col_dst_1based, value=valor)
    return num_val


def destacar_inseridos(
    orig_tratada: Path, used_vals: set[int], atual: str, prefer_xlwings: bool = True
):
    """Highlight cells in the source sheet that were successfully used."""
    if not used_vals:
        return

    if prefer_xlwings and XLWINGS:
        try:
            wb = xw.Book(str(orig_tratada))
            for sht in wb.sheets:
                headers = sht.range("A1").expand("right").value
                if not headers:
                    continue
                headers = [headers] if not isinstance(headers, list) else headers
                for i, val in enumerate(headers):
                    if str(val).strip() == atual:
                        last_row = sht.cells.last_cell.row
                        rng = sht.range((2, i + 1), (last_row, i + 1)).value
                        rng = [rng] if last_row == 1 else rng
                        for r_idx, cell_val in enumerate(rng or [], start=2):
                            if normaliza_num(cell_val) in used_vals:
                                cell = sht.cells(r_idx, i + 1)
                                cell.color = (204, 255, 204)
                                cell.api.Font.Bold = True
            wb.save()
            return
        except Exception:
            pass  # Fallback to openpyxl

    wb = load_workbook(orig_tratada, keep_vba=orig_tratada.suffix.lower() == ".xlsm")
    fill, bold = PatternFill("solid", fgColor="CCFFCC"), Font(bold=True)
    for ws in wb.worksheets:
        atual_cols = [c.column for c in ws[1] if str(c.value).strip() == atual]
        for row in ws.iter_rows(min_row=2):
            for c_idx in atual_cols:
                cell = row[c_idx - 1]
                if normaliza_num(cell.value) in used_vals:
                    cell.fill, cell.font = fill, bold
    wb.save(orig_tratada)


def processar(
    ori: Path,
    spr: Path,
    tipo: str,
    periodo: str,
    src_txt: str,
    dst_txt: str,
    start_row: int,
    dre_start: int,
    out_dir: Path | None = None,
    log: Callable[[str], None] = print,
) -> Path:
    """Process the spreadsheet and highlight used/pending values in the source."""
    src_idx, dst_idx = col_txt_to_idx(src_txt), col_txt_to_idx(dst_txt)
    atual, ant, ant2, is_trim = periodos(periodo)
    orig_path = prepara_origem(ori, tipo, atual, ant, ant2, is_trim, out_dir)
    abas = pd.read_excel(orig_path, sheet_name=None, engine="openpyxl")
    dre_sheet = f"{'cons' if tipo == 'consolidado' else 'ind'} DRE"
    df_dre = abas.get(dre_sheet)
    dfc_sheet = f"{'cons' if tipo == 'consolidado' else 'ind'} DFC"
    df_dfc = abas.get(dfc_sheet)
    used_vals = set()

    if XLWINGS and spr.suffix.lower() in {".xlsx", ".xlsm"}:
        try:
            wb = xw.Book(str(spr))
            sht = wb.sheets[0]
            get_val = lambda r, c: sht.cells(r, c).formula or sht.cells(r, c).value
            
            def set_val(r, c, v):
                prop = "formula" if isinstance(v, str) and v.startswith("=") else "value"
                setattr(sht.cells(r, c), prop, v)

            _, _, used_vals = atualizar_ws(
                sht, get_val, set_val, abas, src_idx, dst_idx, atual, ant, start_row
            )
            if is_trim and df_dre is not None:
                aplicar_dre_manual(df_dre, sht, dst_idx + 1, dre_start, atual, True)
            if df_dfc is not None and (
                v199 := inserir_depreciacao_dfc(
                    df_dfc, sht, dst_idx + 1, 199, atual, True
                )
            ):
                used_vals.add(v199)
            wb.app.calculate()
            wb.save()
        except Exception as exc:
            log(f"xlwings failed, falling back: {exc}")
    else:
        is_xlsm = spr.suffix.lower() == ".xlsm"
        wb = load_workbook(spr, keep_vba=is_xlsm)
        ws = wb.active
        _, _, used_vals = atualizar_ws(
            ws,
            lambda r, c: ws.cell(r, c).value,
            lambda r, c, v: setattr(ws.cell(r, c), "value", v),
            abas, src_idx, dst_idx, atual, ant, start_row,
        )
        if is_trim and df_dre is not None:
            aplicar_dre_manual(df_dre, ws, dst_idx + 1, dre_start, atual, False)
        if df_dfc is not None and (
            v199 := inserir_depreciacao_dfc(
                df_dfc, ws, dst_idx + 1, 199, atual, False
            )
        ):
            used_vals.add(v199)
        out_name = f"{spr.stem} {atual}{'.xlsm' if is_xlsm else '.xlsx'}"
        spr = spr.with_name(out_name)
        wb.save(spr)

    destacar_inseridos(orig_path, used_vals, atual, prefer_xlwings=XLWINGS)
    log(f"Processed source saved to: {orig_path}")
    return spr


class App(ctk.CTk):
    """Main application GUI."""

    def __init__(self):
        super().__init__()
        self.title("Spread Updater")
        self.grid_columnconfigure((0, 1), weight=1)

        self.var_ori = ctk.StringVar()
        self._campo_arquivo("Source File", 0, self.var_ori)
        self.var_spr = ctk.StringVar()
        self._campo_arquivo("Spreadsheet File", 1, self.var_spr)
        self.var_outdir = ctk.StringVar(value=str(Path.cwd()))
        self._campo_dir("Processed Source Folder", 2, self.var_outdir)

        self.var_tipo = ctk.StringVar(value="consolidado")
        ctk.CTkLabel(self, text="Type").grid(row=3, column=0, sticky="w", padx=4)
        ctk.CTkOptionMenu(
            self, variable=self.var_tipo, values=["consolidado", "individual"]
        ).grid(row=3, column=1, sticky="ew", padx=4)

        self.var_per = ctk.StringVar()
        self._campo_txt("Period (2024 / 1T25)", 4, self.var_per)
        self.var_src = ctk.StringVar(value="A")
        self._campo_txt("Source Column (A or 0…)", 5, self.var_src, width=80)
        self.var_dst = ctk.StringVar(value="B")
        self._campo_txt("Destination Column", 6, self.var_dst, width=80)
        self.var_start = ctk.StringVar(value="27")
        self._campo_txt("General Start Row", 7, self.var_start, width=80)

        self.var_dre = ctk.StringVar(value="150")
        self.lbl_dre = ctk.CTkLabel(self, text="DRE Start Row (Revenue)")
        self.ent_dre = ctk.CTkEntry(self, textvariable=self.var_dre, width=80)
        self.var_per.trace_add("write", self._toggle_dre)

        ctk.CTkButton(self, text="Process", command=self._run).grid(
            row=10, column=0, pady=6, padx=4, sticky="ew"
        )
        ctk.CTkButton(
            self, text="Exit", fg_color="gray", command=self.destroy
        ).grid(row=10, column=1, pady=6, padx=4, sticky="ew")

        self.log = ctk.CTkTextbox(self, width=600, height=150, state="disabled")
        self.log.grid(row=11, column=0, columnspan=3, pady=6, padx=4)
        
        self._toggle_dre()


    def _campo_arquivo(self, rotulo: str, linha: int, var: ctk.StringVar):
        ctk.CTkLabel(self, text=rotulo).grid(row=linha, column=0, sticky="w", padx=4)
        ctk.CTkEntry(self, textvariable=var, width=420).grid(
            row=linha, column=1, sticky="ew", padx=4
        )

        def choose_file():
            f = filedialog.askopenfilename(
                filetypes=[("Excel", "*.xlsx *.xlsm *.xls")]
            )
            if f:
                var.set(f)

        button = ctk.CTkButton(self, text="…", width=30, command=choose_file)
        button.grid(row=linha, column=2, padx=2)

    def _campo_dir(self, rotulo: str, linha: int, var: ctk.StringVar):
        ctk.CTkLabel(self, text=rotulo).grid(row=linha, column=0, sticky="w", padx=4)
        ctk.CTkEntry(self, textvariable=var, width=420).grid(
            row=linha, column=1, sticky="ew", padx=4
        )

        def choose_dir():
            d = filedialog.askdirectory()
            if d:
                var.set(d)

        button = ctk.CTkButton(self, text="…", width=30, command=choose_dir)
        button.grid(row=linha, column=2, padx=2)

    def _campo_txt(self, rotulo: str, linha: int, var: ctk.StringVar, width=420):
        ctk.CTkLabel(self, text=rotulo).grid(row=linha, column=0, sticky="w", padx=4)
        ctk.CTkEntry(self, textvariable=var, width=width).grid(
            row=linha, column=1, sticky="w", padx=4
        )

    def _toggle_dre(self, *_):
        is_trim = re.fullmatch(r"[1-4]T\d{2}", self.var_per.get().strip().upper())
        if is_trim:
            self.lbl_dre.grid(row=8, column=0, sticky="w", padx=4)
            self.ent_dre.grid(row=8, column=1, sticky="w", padx=4)
        else:
            self.lbl_dre.grid_remove()
            self.ent_dre.grid_remove()

    def _log(self, msg: str):
        self.log.configure(state="normal")
        self.log.insert("end", f"{msg}\n")
        self.log.configure(state="disabled")
        self.log.see("end")

    def _run(self):
        try:
            ori, spr = Path(self.var_ori.get()), Path(self.var_spr.get())
            if not (ori.exists() and spr.exists()):
                self._log("Please select valid files.")
                return
            out_spread = processar(
                ori,
                spr,
                self.var_tipo.get(),
                self.var_per.get(),
                self.var_src.get(),
                self.var_dst.get(),
                int(self.var_start.get()),
                int(self.var_dre.get()),
                out_dir=None,
                log=self._log,
            )
            self._log(f"✔️ Finished: {out_spread}")
        except Exception as exc:
            logging.exception("Processing failed")
            self._log(f"Error: {exc}")


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    app = App()
    app.mainloop()