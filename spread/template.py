"""Manages the creation and initialization of Spread Proxies from a blank template."""

from __future__ import annotations

import shutil
from pathlib import Path


class TemplateManager:
    """Manages copying a blank template Spread into a target file."""

    def __init__(self, template_path: str | Path):
        self.template_path = Path(template_path)

    def create_from_template(self, dest_path: str | Path, overwrite: bool = False) -> Path:
        """
        Creates a new ready-to-use Spread Proxy from the system template.
        """
        dest_path = Path(dest_path)
        
        if dest_path.exists() and not overwrite:
            raise FileExistsError(f"Destination {dest_path} already exists. Use overwrite=True if desired.")
        
        dest_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(self.template_path, dest_path)
        
        return dest_path
