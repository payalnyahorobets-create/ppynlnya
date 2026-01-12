from __future__ import annotations

from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

MONTH_NAMES = [
    "Січень",
    "Лютий",
    "Березень",
    "Квітень",
    "Травень",
    "Червень",
    "Липень",
    "Серпень",
    "Вересень",
    "Жовтень",
    "Листопад",
    "Грудень",
]


def default_excel_path() -> Path:
    return Path(__file__).resolve().parent / "Анализ NEW NOT FINAL (6).xlsx"


@dataclass(frozen=True)
class ExcelStore:
    path: Path

    @property
    def sheet_names(self) -> list[str]:
        excel = pd.ExcelFile(self.path)
        return excel.sheet_names

    def month_sheets(self) -> list[str]:
        available = set(self.sheet_names)
        sheets: list[str] = []
        for year in range(2024, 2027):
            for month in MONTH_NAMES:
                name = f"{month} {year}"
                if name in available:
                    sheets.append(name)
        return sheets

    @lru_cache(maxsize=32)
    def load_sheet(self, name: str, limit: int | None = None) -> pd.DataFrame:
        return pd.read_excel(self.path, sheet_name=name, nrows=limit)

    def head_as_records(self, name: str, limit: int = 200) -> list[dict]:
        data = self.load_sheet(name, limit)
        return data.fillna("").to_dict(orient="records")

    def columns(self, name: str) -> list[str]:
        data = self.load_sheet(name, 1)
        return [str(col) for col in data.columns]

    @lru_cache(maxsize=64)
    def row_count(self, name: str) -> int:
        workbook = load_workbook(self.path, read_only=True, data_only=True)
        try:
            sheet = workbook[name]
            max_row = sheet.max_row or 0
            return max(0, max_row - 1)
        finally:
            workbook.close()


@dataclass(frozen=True)
class ExcelSummary:
    products_count: int
    analysis_count: int
    month_count: int
    month_sheets: list[str]


def build_summary(store: ExcelStore) -> ExcelSummary:
    month_sheets = store.month_sheets()
    return ExcelSummary(
        products_count=store.row_count("Номенклатура"),
        analysis_count=store.row_count("Аналіз ABC-XYZ"),
        month_count=len(month_sheets),
        month_sheets=month_sheets,
    )
