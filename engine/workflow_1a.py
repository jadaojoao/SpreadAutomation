"""Workflow for Mode 1A (Filling an existing Spread with a single period)."""

from __future__ import annotations

from pathlib import Path
from decimal import Decimal

from core.models import EntityType, SourceType
from ingestion import CVMExcelAdapter
from mapping import MappingRegistry, Mapper
from spread import SpreadReader, SpreadWriter, Highlights
from validation import ValidationReporter


class Mode1AWorkflow:
    """
    Orchestrates the entire logic flow for filling out an existing Spread Proxy
    for a given target CVM period.
    """

    def __init__(self, registry_dir: Path | None = None):
        self.registry = MappingRegistry.load(registry_dir)
        self.mapper = Mapper(self.registry)
        self.reporter = ValidationReporter()

    def execute(
        self,
        source_path: str | Path,
        spread_path: str | Path,
        company: str,
        period: str,
        dest_col: str,
        prior_col: str | None = None,
        entity_type: EntityType = EntityType.CONSOLIDATED
    ) -> dict:
        """
        Executes Mode 1A:
        1. Ingest CVM data (assume Excel CVM template for 1A right now).
        2. Read Spread Proxy schema (with fallback prior values).
        3. Match accounts using core mapper.
        4. Overwrite targets in Spread Proxy.
        5. Apply conditional formatting highlights.
        6. Run validation checks.
        """
        source_path = Path(source_path)
        spread_path = Path(spread_path)

        # 1. Ingest Data
        adapter = CVMExcelAdapter()
        dataset = adapter.load(
            path=source_path,
            company=company,
            period=period,
            entity_type=entity_type,
            # We don't strictly need prior periods here because target spread already has the heuristic columns
        )

        # 2. Extract Spread Schema
        reader = SpreadReader(spread_path)
        target_schema = reader.extract_schema(
            label_col="B",
            prior_val_col=prior_col,
            start_row=10,  # Assumption based on MAPPING_STRATEGY or proxy templates
            end_row=1000
        )

        # 3. Apply Mapping Engine
        mapped_results = self.mapper.map_dataset(
            target_schema=target_schema,
            current_dataset=dataset,
            prior_dataset=None  # Mode 1A relies on the target_schema prior_values
        )

        # 4. Write output
        writer = SpreadWriter(spread_path)
        writer.write_results(dest_col=dest_col, results=mapped_results, overwrite=True)

        # 5. Apply graphical highlights
        highlights = Highlights(spread_path)
        highlights.apply_styles(col=dest_col, results=mapped_results)

        # 6. Validate
        report = self.reporter.report(schema_targets=target_schema, mapped_results=mapped_results)

        return {
            "status": "success",
            "mapped_count": len(mapped_results),
            "validation": {
                "is_valid": report.is_valid,
                "missing": report.missing_labels,
                "discrepancy": str(report.discrepancy)
            }
        }
