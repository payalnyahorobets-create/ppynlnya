from __future__ import annotations

import importlib.util
import subprocess
import sys
from pathlib import Path


def ensure_dependencies(requirements_path: Path) -> None:
    missing = []
    for package in ("flask", "pandas", "openpyxl"):
        if importlib.util.find_spec(package) is None:
            missing.append(package)

    if missing:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "-r", str(requirements_path)]
        )
