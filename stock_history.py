# -*- coding: utf-8 -*-
"""
ibov_compare_auto.py — v2.0
• Lista de tickers fixa no topo (dar RUN e pronto)
• Janela padrão: 3 anos (IBOV + N ativos)
• Não encurta a janela: cada série é rebasada em 100 na sua 1ª data válida
• Exporta Excel (Base100, Excesso_vs_IBOV) e PNG
• (Opcional) injeta/atualiza planilha via xlwings
"""

import os
import logging
import datetime as dt
from typing import List

import pandas as pd
import yfinance as yf
import matplotlib.pyplot as plt

# ===================== CONFIG =====================
YEARS_BACK = 2

# Edite aqui sua lista (sem .SA; o script acrescenta automaticamente):
TICKERS = [
    "PETR4", "VALE3", "BBAS3",  # exemplos — ajuste à sua lista real
]

# Pasta de saída (se vazio, cria ./output/AAAAMMDD)
OUT_DIR = ""

# Salvar Excel/PNG? (True/False)
SAVE_EXCEL = True
SAVE_PNG = True

# ---- Integração opcional com xlwings ----
USE_XLWINGS = False                     # mude para True se quiser injetar no Excel
XLWINGS_TARGET_PATH = r""               # ex.: r"C:\Relatorios\Modelo_IBOV.xlsx"
XLWINGS_SHEET_BASE100 = "Base100"       # aba a receber a tabela base 100
XLWINGS_SHEET_EXCESSO = "Excesso"       # aba a receber a tabela de excesso
XLWINGS_CREATE_LINE_CHART = True        # cria/atualiza um gráfico simples com Base100
XLWINGS_CHART_SHEET = "Comparacao"      # aba onde o gráfico será criado/atualizado
XLWINGS_CHART_NAME = "Chart_Base100"
# ========================================


def setup_logging(out_dir: str) -> None:
    os.makedirs(out_dir, exist_ok=True)
    log_path = os.path.join(out_dir, "run.log")
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[logging.FileHandler(log_path, encoding="utf-8"), logging.StreamHandler()]
    )
    logging.info("Logging iniciado.")


def today_dates() -> tuple[pd.Timestamp, pd.Timestamp]:
    end = pd.Timestamp.today().normalize()
    start = end - pd.DateOffset(years=YEARS_BACK)
    return start, end


def ensure_sa_suffix(t: str) -> str:
    t = t.strip().upper()
    if t.startswith("^"):
        return t
    return t if t.endswith(".SA") else f"{t}.SA"


def download_adj_close(tickers: List[str], start: pd.Timestamp, end: pd.Timestamp) -> pd.DataFrame:
    all_ticks = ["^BVSP"] + [ensure_sa_suffix(t) for t in tickers]
    # auto_adjust=False garante coluna "Adj Close" nas versões recentes do yfinance
    data = yf.download(
        tickers=all_ticks,
        start=start.date().isoformat(),
        end=end.date().isoformat(),
        auto_adjust=False,
        progress=False,
        group_by="column",
        threads=True
    )

    # Seleciona o painel de Adj Close (DataFrame de 2D)
    if isinstance(data, pd.DataFrame) and ("Adj Close" in data.columns):
        adj = data["Adj Close"].copy()
    else:
        # Caso o formato venha "achatado"
        adj = data.copy()
        if isinstance(adj.columns, pd.MultiIndex):
            adj = adj.xs("Adj Close", level=0, axis=1)

    # Ordena por data e mantém union de datas (podem existir NaNs antes de começar cada série)
    adj = adj.sort_index()
    logging.info(f"Datas: {adj.index.min().date()} → {adj.index.max().date()}")
    return adj


def rebase_per_column(df: pd.DataFrame) -> pd.DataFrame:
    """
    Rebase em 100 por coluna, usando o 1º valor NÃO-NaN de cada série.
    Mantém NaNs antes do início da série (não encurta a janela).
    """
    out = pd.DataFrame(index=df.index)
    for col in df.columns:
        s = df[col]
        first_valid_idx = s.first_valid_index()
        if first_valid_idx is None:
            out[col] = pd.NA
            continue
        base_val = s.loc[first_valid_idx]
        out[col] = (s / base_val) * 100.0
    return out


def build_outputs(base100: pd.DataFrame, out_dir: str) -> tuple[str | None, str | None]:
    # Excesso vs IBOV
    if "^BVSP" not in base100.columns:
        raise ValueError("Coluna ^BVSP não encontrada em Base100.")
    excesso = base100.drop(columns=["^BVSP"]).subtract(base100["^BVSP"], axis=0)

    xlsx_path = None
    png_path = None

    if SAVE_EXCEL:
        from openpyxl.styles import numbers  # ← (novo) para number_format

        xlsx_path = os.path.join(out_dir, "series_ibov_base100.xlsx")

        # (novo) nome do índice para ficar claro na planilha
        base100.index.name = "Data"
        excesso.index.name = "Data"

        # --- Aqui entra a Opção B (openpyxl) ---
        with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
            base100.to_excel(writer, sheet_name="Base100")
            excesso.to_excel(writer, sheet_name="Excesso_vs_IBOV")

            # Após escrever, pegue as abas pelo writer.sheets
            ws_b = writer.sheets["Base100"]
            ws_e = writer.sheets["Excesso_vs_IBOV"]

            # Formatar coluna A (datas) como dd/mm/yyyy
            for ws in (ws_b, ws_e):
                # linhas começam em 2 (linha 1 = cabeçalho)
                for r in range(2, ws.max_row + 1):
                    ws.cell(row=r, column=1).number_format = "dd/mm/yyyy"
        # --- fim Opção B ---

        logging.info(f"Excel salvo: {xlsx_path}")

    if SAVE_PNG:
        plt.figure(figsize=(11, 6))
        base100.plot(ax=plt.gca(), linewidth=1.2)
        plt.title("Ações vs IBOV — Base 100 (cada série no seu D0 válido)")
        plt.xlabel("Data")
        plt.ylabel("Base 100")
        plt.grid(True, linewidth=0.3)
        png_path = os.path.join(out_dir, "comparacao_base100.png")
        plt.tight_layout()
        plt.savefig(png_path, dpi=150)
        plt.close()
        logging.info(f"Gráfico salvo: {png_path}")

    return xlsx_path, png_path



def inject_with_xlwings(base100: pd.DataFrame, out_dir: str) -> None:
    if not USE_XLWINGS or not XLWINGS_TARGET_PATH:
        return
    try:
        import xlwings as xw
    except ImportError:
        logging.warning("xlwings não instalado — pulando injeção no Excel.")
        return

    # Excesso (recalcula aqui para evitar passar adiante)
    excesso = base100.drop(columns=["^BVSP"]).subtract(base100["^BVSP"], axis=0)

    logging.info(f"Abrindo planilha: {XLWINGS_TARGET_PATH}")
    app = xw.App(visible=False)  # rodar “headless” no servidor/VM
    try:
        wb = xw.Book(XLWINGS_TARGET_PATH)

        # Escreve/atualiza Base100
        sht_b = wb.sheets[XLWINGS_SHEET_BASE100] if XLWINGS_SHEET_BASE100 in [s.name for s in wb.sheets] else wb.sheets.add(XLWINGS_SHEET_BASE100)
        sht_b.clear()
        sht_b["A1"].value = base100  # xlwings entende DataFrame (índice + colunas)

        # Escreve/atualiza Excesso
        sht_e = wb.sheets[XLWINGS_SHEET_EXCESSO] if XLWINGS_SHEET_EXCESSO in [s.name for s in wb.sheets] else wb.sheets.add(XLWINGS_SHEET_EXCESSO)
        sht_e.clear()
        sht_e["A1"].value = excesso

        # Gráfico simples de linhas com Base100
        if XLWINGS_CREATE_LINE_CHART:
            sht_c = wb.sheets[XLWINGS_CHART_SHEET] if XLWINGS_CHART_SHEET in [s.name for s in wb.sheets] else wb.sheets.add(XLWINGS_CHART_SHEET)
            # Detecta região usada da Base100 para o gráfico
            used = sht_b["A1"].current_region
            # Cria/atualiza
            chart = None
            for ch in sht_c.charts:
                if ch.name == XLWINGS_CHART_NAME:
                    chart = ch
                    break
            if chart is None:
                chart = sht_c.charts.add(left=10, top=20, width=700, height=380, name=XLWINGS_CHART_NAME)
            chart.set_source_data(used)
            chart.chart_type = "line"
        wb.save()
        logging.info("Planilha atualizada com xlwings.")
    finally:
        app.quit()


def main():
    # Pasta de saída
    stamp = dt.date.today().strftime("%Y%m%d")
    out_dir = OUT_DIR or os.path.join(os.getcwd(), "output", stamp)
    setup_logging(out_dir)

    if not TICKERS:
        raise SystemExit("Defina ao menos 1 ticker em TICKERS.")

    logging.info(f"Tickers informados: {TICKERS} (IBOV incluído automaticamente)")
    start, end = today_dates()
    logging.info(f"Janela-alvo: {start.date()} → {end.date()} (~{YEARS_BACK} anos)")

    # Download
    adj = download_adj_close(TICKERS, start, end)

    # Rebase por coluna (sem encurtar a janela)
    base100 = rebase_per_column(adj)

    # Outputs (Excel/PNG)
    build_outputs(base100, out_dir)

    # Injetar no Excel (opcional)
    inject_with_xlwings(base100, out_dir)

    logging.info("Concluído com sucesso.")


if __name__ == "__main__":
    main()
