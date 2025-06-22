# app_spread.py · v4 (com mapa DMPL e destaque completo)
# --------------------------------------------------------------------
# • Mapeamento correto das abas (DF Cons/Ind + DMPL)
# • Cabeçalhos distintos para ano × trimestre
# • DRE trimestral mapeado manualmente
# • Depreciação/Amortização negativa na DFC
# • Atualização ao vivo com xlwings; fallback openpyxl
# • Destaque em verde dos valores usados (used_vals)
# • Destaque em amarelo de todos os valores que apareceram no Spread
# • Relatório de linhas pendentes na GUI
# --------------------------------------------------------------------
from __future__ import annotations
import logging
import re
import sys
import tempfile
from pathlib import Path
from typing import Callable, Dict, List, Set, Tuple

import customtkinter as ctk
import pandas as pd
from tkinter import filedialog
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
    """Normaliza um valor numérico ou texto para inteiro, ou retorna None."""
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
    """
    Retorna (atual, anterior, ante-anterior, is_trimestre).
    Aceita '2024' ou '1T25'.
    """
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
    """
    Converte letra de coluna Excel (ex.: 'A', 'BC') ou dígito '0' em índice 0-based.
    """
    t = txt.strip().upper()
    if t.isdigit():
        return int(t)
    return col2idx(t) - 1


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
    Gera <stem>_tratado.xlsx/.xlsm sem jamais sobrescrever o original.
    Renomeia abas e cabeçalhos por ano ou trimestre.
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
            "DF Cons DMPL Ultimo": "cons DMPL",
        },
        "individual": {
            "DF Ind Ativo": "ind ativos",
            "DF Ind Passivo": "ind passivos",
            "DF Ind Resultado Periodo": "ind DRE",
            "DF Ind Fluxo de Caixa": "ind DFC",
            "DF Ind DMPL Ultimo": "ind DMPL",
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
            cl = c.lower().strip()
            if is_trim and is_ap:
                if cl.startswith(H_TRI_AP[0]):
                    return atual
                if cl.startswith(H_TRI_AP[1]):
                    return ant
            elif is_trim and is_res:
                if cl.startswith(H_TRI_RES[0]):
                    return atual
                if cl.startswith(H_TRI_RES[1]):
                    return ant
            if cl.startswith(H_ANO[0]):
                return atual
            if cl.startswith(H_ANO[1]):
                return ant
            if cl.startswith(H_ANO[2]):
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

    # garante que ao menos a primeira aba fique visível
    wb = load_workbook(dst_path)
    if wb.sheetnames:
        wb[wb.sheetnames[0]].sheet_state = "visible"
    wb.save(dst_path)

    return dst_path


def shift_formula(f: str, delta: int) -> str:
    """Desloca referências de coluna em fórmulas complexas por um delta."""
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
    formula: str,
    delta: int,
    map_number: Callable[[int], int | None],
    used_vals: Set[int] | None = None,
) -> str:
    """Ajusta fórmula complexa deslocando colunas e mapeando literais."""
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
    """Retorna valor correspondente de n em abas, de prev→curr."""
    for df in abas.values():
        if prev not in df.columns or curr not in df.columns:
            continue
        hit = df[df[prev].apply(normaliza_num) == n]
        if not hit.empty:
            return normaliza_num(hit[curr].iloc[0])
    return None


def atualizar_ws(
    ws,
    get_val: Callable[[int, int], object],
    set_val: Callable[[int, int, object], None],
    abas: Dict[str, pd.DataFrame],
    src_idx: int,
    dst_idx: int,
    atual: str,
    ant: str,
    start_row: int,
) -> tuple[List[int], Set[int], Set[int]]:
    """
    Copia e ajusta valores/fórmulas da coluna origem→destino.
    Retorna (skipped_rows, skipped_vals, used_vals).
    """
    c_src, c_dst = src_idx + 1, dst_idx + 1
    delta = c_dst - c_src
    skipped_rows: List[int] = []
    skipped_vals: Set[int] = set()
    used_vals: Set[int] = set()

    num_pat = re.compile(r"[-+]?\d[\d\.,]*")
    # determina corretamente o número de linhas, seja openpyxl ou xlwings
    try:
        max_row = ws.max_row
    except AttributeError:
        # xlwings Sheet não tem max_row
        max_row = ws.cells.last_cell.row

    empty_streak = 0
    r = start_row
    while empty_streak < 30 and r <= max_row:  # tipo openpyxl
        v = get_val(r, c_src)
        if v in (None, ""):
            empty_streak += 1
            r += 1
            continue
        empty_streak = 0

        wrote = False
        destino = v

        if isinstance(v, str) and v.startswith("="):
            if not re.search(r"[A-Za-z]", v[1:]):
                # soma/subtração de literais
                def lit_repl(m: re.Match) -> str:
                    tok = m.group(0)
                    n0 = normaliza_num(tok.lstrip("+-"))
                    n1 = valor_corresp(abas, n0, ant, atual)
                    if n1 is not None:
                        used_vals.add(n1)
                        sign = tok[0] if tok[0] in "+-" else ""
                        return f"{sign}{abs(n1)}"
                    return tok

                destino = "=" + num_pat.sub(lit_repl, v[1:])
                wrote = destino != v
            else:
                # fórmula complexa
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
            skipped_vals.add(normaliza_num(v) or 0)

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
) -> None:
    """Insere manualmente linhas da DRE trimestral a partir de DRE_MAP."""
    for offset, desc in DRE_MAP.items():
        linha = dre_start + offset
        try:
            raw = df_dre.loc[
                df_dre["Descricao Conta"].str.strip() == desc, col_valor
            ].iloc[0]
        except Exception:
            continue
        val = normaliza_num(raw) if normaliza_num(raw) is not None else raw
        if is_xlwings:
            sheet.cells(linha, col_dst_1based).value = val
        else:
            sheet.cell(linha, col_dst_1based, value=val)


def inserir_depreciacao_dfc(
    df_dfc: pd.DataFrame,
    sheet,
    col_dst_1based: int,
    linha: int,
    col_valor: str,
    is_xlwings: bool,
) -> int | None:
    """Lê Depreciação/Amortização da DFC e grava sempre como negativo."""
    if df_dfc is None or col_valor not in df_dfc.columns:
        return None
    desc = df_dfc["Descricao Conta"].astype(str)
    mask = desc.str.contains("deprecia|amortiza", case=False, na=False)
    try:
        raw = df_dfc.loc[mask, col_valor].iloc[0]
    except Exception:
        return None
    nv = normaliza_num(raw)
    if nv is not None:
        nv = -abs(nv)
        val = nv
    else:
        s = str(raw).lstrip("+-")
        val = f"-{s}"
        nv = normaliza_num(val)
    if is_xlwings:
        sheet.cells(linha, col_dst_1based).value = val
    else:
        sheet.cell(linha, col_dst_1based, value=val)
    return nv

def destacar_inseridos(
    orig_tratada: Path, used_vals: Set[int], atual: str, prefer_xlwings: bool = True
) -> None:
    """
    Destaca (verde+negrito) todas as células da coluna do período `atual`
    cujo valor numérico esteja em `used_vals`.
    """
    if not used_vals:
        return

    # xlwings
    if prefer_xlwings and XLWINGS:
        try:
            wb = xw.Book(str(orig_tratada))
            for sht in wb.sheets:
                hdrs = sht.range("A1").expand("right").value
                hdrs = hdrs if isinstance(hdrs, list) else [hdrs]
                cols = [i + 1 for i, h in enumerate(hdrs) if str(h).strip() == atual]
                last = sht.cells.last_cell.row
                for c in cols:
                    vals = sht.range((2, c), (last, c)).value
                    vals = vals if isinstance(vals, list) else [vals]
                    for idx, v in enumerate(vals, start=2):
                        if normaliza_num(v) in used_vals:
                            cell = sht.cells(idx, c)
                            cell.color = (204, 255, 204)
                            cell.api.Font.Bold = True
            wb.save()
            return
        except Exception:
            pass

    # openpyxl
    wb = load_workbook(orig_tratada, keep_vba=orig_tratada.suffix.lower() == ".xlsm")
    fill, bold = PatternFill("solid", fgColor="CCFFCC"), Font(bold=True)
    for ws in wb.worksheets:
        cols = [c.column for c in ws[1] if str(c.value).strip() == atual]
        for row in ws.iter_rows(min_row=2):
            for c in cols:
                cell = row[c - 1]
                if normaliza_num(cell.value) in used_vals:
                    cell.fill, cell.font = fill, bold
    wb.save(orig_tratada)

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

def destacar_novos(orig_tratada: Path,
                   prev: str,
                   atual: str,
                   prefer_xlwings: bool = True) -> None:
    """
    Destaca em azul-claro (+negrito) todas as células onde:
      • em 'prev' o valor era 0
      • em 'atual' o valor é != 0
    Usa xlwings em bloco para ser rápido mesmo com o arquivo aberto.
    """
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill, Font

    # cores
    rgb = (153, 204, 255)
    fill_op = PatternFill("solid", fgColor="99CCFF")
    bold_op = Font(bold=True)

    if prefer_xlwings and XLWINGS:
        try:
            # 1) anexa ao workbook já aberto (ou abre se fechar)
            for bk in xw.books:
                if Path(bk.fullname).resolve() == orig_tratada.resolve():
                    wb = bk
                    break
            else:
                wb = xw.Book(str(orig_tratada))

            for sht in wb.sheets:
                # 2) cabeçalhos na linha 1
                hdrs = sht.range("A1").expand("right").value
                hdrs = hdrs if isinstance(hdrs, list) else [hdrs]

                if prev not in hdrs or atual not in hdrs:
                    continue
                c_prev  = hdrs.index(prev)  + 1
                c_atual = hdrs.index(atual) + 1

                # 3) pega vetor de valores de prev e atual de uma vez
                last = sht.cells.last_cell.row
                vals_prev  = sht.range((2, c_prev),  (last, c_prev)).value
                vals_atual = sht.range((2, c_atual), (last, c_atual)).value

                # normaliza para lista
                if not isinstance(vals_prev, list):  vals_prev  = [vals_prev]
                if not isinstance(vals_atual, list): vals_atual = [vals_atual]

                # 4) decide quais linhas destacar
                to_highlight: List[int] = []
                for i, (pv, av) in enumerate(zip(vals_prev, vals_atual), start=2):
                    if normaliza_num(pv) == 0 and normaliza_num(av) not in (None, 0):
                        to_highlight.append(i)

                # 5) faz só as chamadas COM de pintura
                for row in to_highlight:
                    cell = sht.cells(row, c_atual)
                    cell.color = rgb
                    cell.api.Font.Bold = True

            wb.save()
            return
        except Exception:
            # cai no fallback se der qualquer erro
            pass

    # --- fallback openpyxl (arquivo fechado) ---
    wb = load_workbook(orig_tratada,
                       keep_vba=orig_tratada.suffix.lower() == ".xlsm")
    fill = PatternFill("solid", fgColor="99CCFF")
    bold = Font(bold=True)

    for ws in wb.worksheets:
        headers = {cell.value: cell.column for cell in ws[1]}
        col_prev  = headers.get(prev)
        col_atual = headers.get(atual)
        if not col_prev or not col_atual:
            continue

        for row in ws.iter_rows(min_row=2):
            vprev = row[col_prev - 1].value
            vat   = row[col_atual - 1]
            if normaliza_num(vprev) == 0 and normaliza_num(vat.value) not in (None, 0):
                vat.fill = fill
                vat.font = bold

    wb.save(orig_tratada)



def coletar_vals_do_spread(spread_path: Path, dst_idx: int, start_row: int) -> Set[int]:
    """
    Abre o Spread processado e coleta todos os ints da coluna-destino,
    a partir de start_row, até 30 vazios seguidos.
    """
    wb = load_workbook(spread_path, data_only=True)
    ws = wb.active
    vals: Set[int] = set()
    empty = 0
    r = start_row
    max_r = ws.max_row
    while empty < 30 and r <= max_r:
        raw = ws.cell(r, dst_idx + 1).value
        if raw in (None, ""):
            empty += 1
        else:
            empty = 0
            n = normaliza_num(raw)
            if n is not None:
                vals.add(n)
        r += 1
    return vals


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
    """Pipeline principal: gera Spread, origem tratada e destaques."""
    src_idx = col_txt_to_idx(src_txt)
    dst_idx = col_txt_to_idx(dst_txt)
    atual, ant, ant2, is_trim = periodos(periodo)

    orig_path = prepara_origem(ori, tipo, atual, ant, ant2, is_trim, out_dir)
    abas = pd.read_excel(orig_path, sheet_name=None, engine="openpyxl")
    dre_sheet = f"{'cons' if tipo=='consolidado' else 'ind'} DRE"
    df_dre = abas.get(dre_sheet)
    df_dfc = abas.get(f"{'cons' if tipo=='consolidado' else 'ind'} DFC")

    used_vals: Set[int] = set()

    # tenta xlwings
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
                v199 := inserir_depreciacao_dfc(df_dfc, sht, dst_idx + 1, 199, atual, True)
            ):
                used_vals.add(v199)
            wb.app.calculate()
            wb.save()
        except Exception as exc:
            log(f"xlwings falhou, usando fallback: {exc}")

    # fallback openpyxl
    is_xlsm = spr.suffix.lower() == ".xlsm"
    wb2 = load_workbook(spr, keep_vba=is_xlsm)
    ws2 = wb2.active
    _, _, used_vals = atualizar_ws(
        ws2,
        lambda r, c: ws2.cell(r, c).value,
        lambda r, c, v: setattr(ws2.cell(r, c), "value", v),
        abas, src_idx, dst_idx, atual, ant, start_row
    )
    if is_trim and df_dre is not None:
        aplicar_dre_manual(df_dre, ws2, dst_idx + 1, dre_start, atual, False)
    if df_dfc is not None and (
        v199 := inserir_depreciacao_dfc(df_dfc, ws2, dst_idx + 1, 199, atual, False)
    ):
        used_vals.add(v199)
    out_name = f"{spr.stem} {atual}{'.xlsm' if is_xlsm else '.xlsx'}"
    spr = spr.with_name(out_name)
    wb2.save(spr)

    # destaca valores usados + todos que apareceram no Spread
    spread_vals = coletar_vals_do_spread(spr, dst_idx, start_row)
    highlight = used_vals.union(spread_vals)
    destacar_inseridos(orig_path, highlight, atual, prefer_xlwings=XLWINGS)
    # pinta de azul‐claro onde prev era 0 e atual ≠ 0
    destacar_novos(orig_path, ant, atual, prefer_xlwings=XLWINGS)



    log(f"Origem tratada salva em: {orig_path}")
    missing = spread_vals - used_vals
    if missing:
        log(f"⚠️  {len(missing)} valores entraram no Spread mas não estavam em used_vals:")
        log(f"    {sorted(missing)}")
    return spr


class App(ctk.CTk):
    """Interface GUI em CustomTkinter para o atualizador de Spread."""
    def __init__(self):
        super().__init__()
        self.title("Atualizador de Spread")
        self.grid_columnconfigure((0, 1), weight=1)
        # Arquivos
        self.var_ori = ctk.StringVar()
        self._campo_arquivo("Arquivo Origem", 0, self.var_ori)
        self.var_spr = ctk.StringVar()
        self._campo_arquivo("Arquivo Spread", 1, self.var_spr)
        # Tipo
        self.var_tipo = ctk.StringVar(value="consolidado")
        ctk.CTkLabel(self, text="Tipo").grid(row=2, column=0, sticky="w", padx=4)
        ctk.CTkOptionMenu(self, variable=self.var_tipo,
                          values=["consolidado", "individual"]
                          ).grid(row=2, column=1, sticky="ew", padx=4)
        # Período e colunas
        self.var_per = ctk.StringVar()
        self._campo_txt("Período (Ex: 2024 ou 1T25)", 3, self.var_per)
        self.var_src = ctk.StringVar(value="A")
        self._campo_txt("Coluna Origem", 4, self.var_src, width=80)
        self.var_dst = ctk.StringVar(value="B")
        self._campo_txt("Coluna Destino", 5, self.var_dst, width=80)
        # Botões e log
        ctk.CTkButton(self, text="Processar", command=self._run
                      ).grid(row=10, column=0, pady=10, padx=4, sticky="ew")
        ctk.CTkButton(self, text="Sair", fg_color="gray",
                      command=self.destroy
                      ).grid(row=10, column=1, pady=10, padx=4, sticky="ew")
        self.log = ctk.CTkTextbox(self, width=600, height=150, state="disabled")
        self.log.grid(row=11, column=0, columnspan=2, pady=(5,10), padx=4)
        logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

    def _campo_arquivo(self, rotulo: str, linha: int, var: ctk.StringVar):
        ctk.CTkLabel(self, text=rotulo).grid(row=linha, column=0,
                                            sticky="w", padx=4)
        ctk.CTkEntry(self, textvariable=var, width=420
                     ).grid(row=linha, column=1, sticky="ew", padx=4)
        def escolhe():
            f = filedialog.askopenfilename(
                filetypes=[("Excel", "*.xlsx *.xlsm *.xls")])
            if f:
                var.set(f)
        ctk.CTkButton(self, text="…", width=30, command=escolhe
                      ).grid(row=linha, column=2, padx=2)

    def _campo_txt(self, rotulo: str, linha: int,
                   var: ctk.StringVar, width: int = 420):
        ctk.CTkLabel(self, text=rotulo).grid(row=linha, column=0,
                                            sticky="w", padx=4)
        ctk.CTkEntry(self, textvariable=var, width=width
                     ).grid(row=linha, column=1, sticky="w", padx=4)

    def _log(self, msg: str):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.configure(state="disabled")
        self.log.see("end")

    def _run(self):
        try:
            ori = Path(self.var_ori.get())
            spr = Path(self.var_spr.get())
            if not (ori.exists() and spr.exists()):
                self._log("Selecione arquivos válidos.")
                return
            out = processar(
                ori=ori, spr=spr, tipo=self.var_tipo.get(),
                periodo=self.var_per.get(),
                src_txt=self.var_src.get(), dst_txt=self.var_dst.get(),
                start_row=27, dre_start=150, out_dir=None,
                log=self._log
            )
            self._log(f"✔️  Finalizado: {out}")
        except Exception as e:
            logging.exception("Erro no processamento")
            self._log(f"Erro: {e}")


if __name__ == "__main__":
    App().mainloop()
