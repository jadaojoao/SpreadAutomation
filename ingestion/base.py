from abc import ABC, abstractmethod
from typing import Any

from core.models import FinancialDataSet


class IngestionAdapter(ABC):
    """
    Base class for all ingestion adapters.
    Adapters are responsible for taking a raw source document (Excel, CSV, PDF, etc.)
    and transforming it into a canonical FinancialDataSet.
    """

    @abstractmethod
    def load(self, *args: Any, **kwargs: Any) -> FinancialDataSet:
        """
        Parses the source data and returns a standardized FinancialDataSet.
        """
        pass
