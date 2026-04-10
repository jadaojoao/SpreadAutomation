from pathlib import Path
from decimal import Decimal
import pandas as pd

from .base import IngestionAdapter
from core.models import FinancialDataSet, FinancialAccount, EntityType, SourceType


class CVMExcelAdapter(IngestionAdapter):
    """
    Adapter to parse CVM 'DadosDocumento.xlsx' individual or consolidated
    files into a canonical FinancialDataSet.
    """

    def load(
        self,
        path: str | Path,
        company: str,
        period: str,
        entity_type: EntityType,
        cnpj: str | None = None,
        previous_period: str | None = None,
        previous_previous_period: str | None = None,
    ) -> FinancialDataSet:
        path = Path(path)
        is_trim = "T" in period.upper()
        
        chapa = "Cons" if entity_type == EntityType.CONSOLIDATED else "Ind"
        aba_dm = f"DF {chapa} DMPL {'Atual' if is_trim else 'Ultimo'}"
        
        sheet_map = {
            f"DF {chapa} Ativo": "ATIVO",
            f"DF {chapa} Passivo": "PASSIVO",
            f"DF {chapa} Resultado Periodo": "DRE",
            f"DF {chapa} Fluxo de Caixa": "DFC",
            aba_dm: "DMPL",
        }

        engine = "openpyxl" if path.suffix.lower() in (".xlsx", ".xlsm") else None
        xls = pd.ExcelFile(path, engine=engine)
        
        accounts = []
        
        for sheet_orig, section in sheet_map.items():
            if sheet_orig not in xls.sheet_names:
                continue
                
            df = pd.read_excel(xls, sheet_name=sheet_orig, engine=engine)
            
            for _, row in df.iterrows():
                # Extract code and description
                codigo = str(row.get("Codigo Conta", row.get("CodigoConta", ""))).strip()
                descricao = str(row.get("Descricao Conta", row.get("DescricaoConta", ""))).strip()
                
                if not codigo or not descricao:
                    continue
                    
                # Identify value columns
                if section == "DMPL":
                    # For DMPL, we look for 'Patrimonio liquido Consolidado' or 'Patrimonio Liquido'
                    val_col = "Patrimonio liquido Consolidado" if entity_type == EntityType.CONSOLIDATED else "Patrimonio Liquido"
                    if val_col in df.columns:
                        val = row.get(val_col)
                        if pd.notna(val):
                            accounts.append(FinancialAccount(
                                code=codigo,
                                description=descricao,
                                value=Decimal(str(val)),
                                period=period,
                                section=section,
                                source=SourceType.CVM_EXCEL
                            ))
                else:
                    # Period matching logic based on core/origin
                    val_atual = None
                    val_ant = None
                    val_ant2 = None
                    
                    if is_trim:
                        if section in ("ATIVO", "PASSIVO"):
                            val_atual = row.get("Valor Trimestre Atual")
                            val_ant = row.get("Valor Exercicio Anterior")
                        else:
                            val_atual = row.get("Valor Acumulado Atual Exercicio")
                            val_ant = row.get("Valor Acumulado Exercicio Anterior")
                    else:
                        val_atual = row.get("Valor Ultimo Exercicio")
                        val_ant = row.get("Valor Penultimo Exercicio")
                        val_ant2 = row.get("Valor Antepenultimo Exercicio")
                    
                    # Add current period account
                    if val_atual is not None and pd.notna(val_atual):
                        accounts.append(FinancialAccount(
                            code=codigo,
                            description=descricao,
                            value=Decimal(str(val_atual)),
                            period=period,
                            section=section,
                            source=SourceType.CVM_EXCEL
                        ))
                    
                    # Add previous period account
                    if val_ant is not None and pd.notna(val_ant) and previous_period:
                        accounts.append(FinancialAccount(
                            code=codigo,
                            description=descricao,
                            value=Decimal(str(val_ant)),
                            period=previous_period,
                            section=section,
                            source=SourceType.CVM_EXCEL
                        ))
                        
                    # Add previous-previous period account if DFP
                    if not is_trim and val_ant2 is not None and pd.notna(val_ant2) and previous_previous_period:
                        accounts.append(FinancialAccount(
                            code=codigo,
                            description=descricao,
                            value=Decimal(str(val_ant2)),
                            period=previous_previous_period,
                            section=section,
                            source=SourceType.CVM_EXCEL
                        ))

        return FinancialDataSet(
            company=company,
            cnpj=cnpj,
            period=period,
            entity_type=entity_type,
            source_type=SourceType.CVM_EXCEL,
            accounts=accounts
        )
