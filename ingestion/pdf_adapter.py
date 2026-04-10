"""PDF adapter for ingestion."""

from pathlib import Path

from core.models import EntityType, SourceType, FinancialDataSet
from ingestion.base import IngestionAdapter
from .pdf.extractor import MarkdownExtractor
from .pdf.parser import MarkdownParser


class PDFAdapter(IngestionAdapter):
    """
    Adapter to parse PDF Document representations into FinancialDataSet.
    """

    def load(
        self,
        path: str | Path,
        company: str,
        period: str,
        entity_type: EntityType = EntityType.CONSOLIDATED,
        cnpj: str | None = None,
    ) -> FinancialDataSet:
        extractor = MarkdownExtractor()
        raw_md = extractor.extract(path)

        parser = MarkdownParser()
        accounts = list(parser.parse(raw_md, period=period))

        return FinancialDataSet(
            company=company,
            cnpj=cnpj,
            period=period,
            entity_type=entity_type,
            source_type=SourceType.PDF,
            accounts=accounts
        )
