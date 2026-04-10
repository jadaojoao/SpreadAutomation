"""Workflow for Mode 1B (Multi-period initialization from scratch)."""

from __future__ import annotations

from pathlib import Path
from decimal import Decimal
from typing import Sequence, Any

from core.models import EntityType, SourceType, FinancialDataSet
from ingestion import CVMExcelAdapter, CVMCSVAdapter
from mapping import MappingRegistry, Mapper
from spread import SpreadReader, SpreadWriter, Highlights, TemplateManager
from validation import ValidationReporter
from validation.coverage import CoverageValidator


class Mode1BWorkflow:
    """
    Orchestrates building a new multi-period Proxy Spread entirely from scratch 
    based on a blank template and an array of datasets.
    """

    def __init__(
        self,
        template_manager: TemplateManager,
        registry_dir: Path | None = None
    ):
        self.template_manager = template_manager
        self.registry = MappingRegistry.load(registry_dir)
        self.mapper = Mapper(self.registry)
        self.reporter = ValidationReporter()
        self.coverage_validator = CoverageValidator()

    def execute(
        self,
        datasets: Sequence[dict[str, Any]],  # source configs
        dest_spread: str | Path,
    ) -> dict[str, dict]:
        """
        Executes Mode 1B:
        1. Validates coverage (no multi-year gaps).
        2. Copies the empty Spread Proxy template into dest_spread.
        3. Sorts datasets chronologically.
        4. For the oldest dataset -> Layer 1 only (no prior_value).
        5. For subsequent datasets -> Inject Layer 2 parameters matching prior_dataset.
        6. Apply Highlights.
        """
        dest_path = Path(dest_spread)
        
        # Extract periods
        periods = [cfg.get("period") for cfg in datasets if cfg.get("period")]
        if not periods:
            return {"error": {"message": "No valid periods found in datasets"}}
            
        gaps = self.coverage_validator.validate_gaps(periods)
        if gaps:
            return {"error": {"message": f"Gaps detected in coverage: {gaps}"}}
            
        # Create empty template
        self.template_manager.create_from_template(dest_path=dest_path, overwrite=True)
        
        # Simple string-sort for chronology of annual dates
        sorted_configs = sorted(datasets, key=lambda x: str(x.get("period", "")))
        
        reports = {}
        prior_financial_dataset: FinancialDataSet | None = None
        
        # Read the skeleton schema ONCE from the newly copied clean template
        reader = SpreadReader(dest_path)
        base_schema = reader.extract_schema(label_col="B", start_row=10, end_row=1000)
        
        for idx, config in enumerate(sorted_configs):
            source_path = config.get("source_path")
            company = config.get("company", "Company")
            period = config.get("period")
            dest_col = config.get("dest_col", "E") 
            entity_type = config.get("entity_type", EntityType.CONSOLIDATED)
            source_type = config.get("source_type", SourceType.CVM_EXCEL)
            
            # 1. Ingest
            if source_type == SourceType.CVM_EXCEL:
                adapter = CVMExcelAdapter()
            else:
                adapter = CVMCSVAdapter()
                
            current_dataset = adapter.load(
                path=source_path,
                company=company,
                period=period,
                entity_type=entity_type
            )
            
            # For iteration 0, we can't use Layer 2 because we have no previous Spread column 
            # to parse target_prior_values from.
            # Base schema extracted from an empty template handles prior_value=None elegantly.
            
            # Map
            mapped_results = self.mapper.map_dataset(
                target_schema=base_schema, 
                current_dataset=current_dataset,
                prior_dataset=prior_financial_dataset 
            )
            
            # Since Mapper sets prior_value heuristics via target_schema AND 
            # checks prior_dataset accounts, if we don't have a prior target value, 
            # it falls back to Layer 1 safely.
            
            writer = SpreadWriter(dest_path)
            writer.write_results(dest_col=dest_col, results=mapped_results, overwrite=True)
            
            highlights = Highlights(dest_path)
            highlights.apply_styles(col=dest_col, results=mapped_results)
            
            # Validation
            rep = self.reporter.report(schema_targets=base_schema, mapped_results=mapped_results)
            
            reports[period] = {
                "status": "success",
                "is_valid": rep.is_valid,
                "discrepancy": str(rep.discrepancy),
                "mapped_count": len(mapped_results)
            }
            
            # Keep previous dataset state for Layer 2 heuristical comparison via Mapper
            # (Note that MAPPING_STRATEGY describes L2 by searching Spread's previous column, 
            # but Mapper logic built internally in engine uses prior_dataset accounts)
            prior_financial_dataset = current_dataset

        return {"reports": reports}
