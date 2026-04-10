# processing/dre.py — Aplicação manual de DRE trimestral
# v2: label-based matching — escaneia col B do Spread para encontrar linhas pelos
# rótulos, sem depender de offsets fixos. Suporta múltiplas contas CVM somadas
# na mesma linha do Spread (ex: Perdas NRA + Outras Despesas Operacionais).

from __future__ import annotations

from collections import defaultdict

import pandas as pd

from core.exceptions import SchemaError
from core.utils import normaliza_num
from processing.runtime_bridge import mapping_registry, spread_schema


def aplicar_dre_manual(
    df_dre: pd.DataFrame, sheet,
    col_dst_1based: int, start_row: int,
    col_valor: str, is_xlwings: bool
) -> None:
    """
    Preenche as linhas da DRE no Spread usando o MappingRegistry Layer 2.

      1. Resolve a descrição CVM para um row key canônico
      2. Acumula múltiplas contas CVM no mesmo row key
      3. Resolve a linha final pelo SpreadSchema
      4. Grava o total acumulado na célula encontrada

    Args:
        df_dre:        DataFrame da DRE da Origem (com col "Descricao Conta")
        sheet:         Worksheet do Spread (openpyxl ou xlwings)
        col_dst_1based: Índice 1-based da coluna destino no Spread
        start_row:     Mantido por compatibilidade com os callers legados
        col_valor:     Nome da coluna no df_dre com os valores (ex: "3T25")
        is_xlwings:    True se sheet for xlwings, False se for openpyxl
    """
    del start_row
    registry = mapping_registry()
    schema = spread_schema()

    sums: dict[str, int | None] = defaultdict(lambda: None)
    for _, record in df_dre.iterrows():
        desc = record.get("Descricao Conta")
        if pd.isna(desc):
            continue
        row_key = registry.layer2(str(desc))
        if row_key is None:
            continue
        raw = record.get(col_valor)
        val = normaliza_num(raw)
        if val is None:
            continue
        prev = sums[row_key]
        sums[row_key] = val if prev is None else prev + val

    for row_key, total in sums.items():
        if total is None:
            continue
        try:
            row_number = schema.row_for(row_key)
        except SchemaError:
            continue
        if is_xlwings:
            sheet.cells(row_number, col_dst_1based).value = total
        else:
            sheet.cell(row_number, col_dst_1based, value=total)
