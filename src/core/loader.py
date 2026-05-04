"""InfoLoader — reads the two sheets from the E&O Excel workbook.

Expected sheets
---------------
Summary  : one row per GTK Supplier + Platform combination,
           includes Actual Payment and other per-deal columns.
Info     : supplier contract meta-data (Master Agreement, Signer, etc.)
           keyed on GTK Supplier + Platform.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

from src.config.paths import EXCEL_FILE


class InfoLoader:
    """Loads *Summary* and *Info* sheets from the E&O Excel workbook."""

    SHEET_SUMMARY = "Summary"
    SHEET_INFO = "Info"

    def __init__(self, excel_path: Path = EXCEL_FILE) -> None:
        self.excel_path = excel_path
        self._summary: pd.DataFrame | None = None
        self._info: pd.DataFrame | None = None

    # ------------------------------------------------------------------
    # Last legitimate column in the Info sheet.  Everything to the right is
    # considered junk entered by users and is silently discarded.
    INFO_LAST_COLUMN = "Address"

    def load(self) -> "InfoLoader":
        """Read both sheets; return *self* for chaining."""
        xl = pd.ExcelFile(self.excel_path, engine="openpyxl")
        self._summary = pd.read_excel(xl, sheet_name=self.SHEET_SUMMARY)

        raw_info = pd.read_excel(xl, sheet_name=self.SHEET_INFO)
        # Truncate all columns after the last legitimate one.
        if self.INFO_LAST_COLUMN in raw_info.columns:
            cutoff = raw_info.columns.get_loc(self.INFO_LAST_COLUMN)
            raw_info = raw_info.iloc[:, : cutoff + 1]
        self._info = raw_info
        return self

    # ------------------------------------------------------------------
    @property
    def summary(self) -> pd.DataFrame:
        if self._summary is None:
            raise RuntimeError("Call load() before accessing summary.")
        return self._summary

    @property
    def info(self) -> pd.DataFrame:
        if self._info is None:
            raise RuntimeError("Call load() before accessing info.")
        return self._info

    @property
    def summary_count(self) -> int:
        return 0 if self._summary is None else len(self._summary)

    @property
    def info_count(self) -> int:
        return 0 if self._info is None else len(self._info)
