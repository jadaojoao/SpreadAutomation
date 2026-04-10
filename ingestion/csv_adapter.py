import pandas as pd
from pathlib import Path
from decimal import Decimal

from .base import IngestionAdapter
from core.models import FinancialDataSet, FinancialAccount, EntityType, SourceType


class CVMCSVAdapter(IngestionAdapter):
    """
    Adapter to parse Data provided by CVM in open data standard (CSV format).
    """

    def load(
        self,
        path: str | Path,
        company: str,
        period: str,
        entity_type: EntityType,
        cnpj: str | None = None,
        section: str | None = None,
    ) -> FinancialDataSet:
        path = Path(path)
        
        import pandas as pd
        df = pd.read_csv(path, sep=";", encoding="utf-8")
        
        # Apply filters
        if cnpj:
            df = df[df["CNPJ_CIA"] == cnpj]
        else:
            df = df[df["DENOM_CIA"].str.contains(company, case=False, na=False)]
            
        accounts = []
        for _, row in df.iterrows():
            codigo = str(row.get("CD_CONTA", "")).strip()
            descricao = str(row.get("DS_CONTA", "")).strip()
            valor = row.get("VL_CONTA")
            
            if not codigo or not descricao or pd.isna(valor):
                continue
                
            accounts.append(
                FinancialAccount(
                    code=codigo,
                    description=descricao,
                    value=Decimal(str(valor)),
                    period=period,  # Needs proper mapping dynamically based on CSV DT_REFER/ORDEM
                    section=section or "UNKNOWN", 
                    source=SourceType.CVM_CSV
                )
            )

        return FinancialDataSet(
            company=company,
            cnpj=cnpj,
            period=period,
            entity_type=entity_type,
            source_type=SourceType.CVM_CSV,
            accounts=accounts
        )
