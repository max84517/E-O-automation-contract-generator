"""DataProcessor — merges Summary + Info DataFrames and provides
formatting helpers used by both the writer and the GUI.
"""
from __future__ import annotations

from datetime import date, datetime

import pandas as pd
from num2words import num2words


class DataProcessor:
    """Merges the two source DataFrames and exposes static formatters."""

    ELAN_SYNA_KEYWORDS: tuple[str, ...] = ("elan", "synaptics", "syna")

    # Columns to pull from Info; avoids column-name conflicts with Summary
    # (GBU, Sub-Category, ODM overlap → keep only these from Info).
    _INFO_KEEP: list[str] = [
        "Supplier",
        "Supplier name",
        "Master Agreement",
        "Effective Date",
        "Signer",
        "Signer title",
    ]

    def __init__(self, summary: pd.DataFrame, info: pd.DataFrame) -> None:
        self.summary = summary
        self.info = info
        self._merged: pd.DataFrame | None = None

    # ------------------------------------------------------------------
    def process(self) -> "DataProcessor":
        """Left-join Summary → Info on GTK Supplier = Supplier.

        Info is keyed by Supplier (one row per supplier regardless of
        platform).  Only the contract-metadata columns are kept from
        Info to avoid column-name collisions (GBU, Sub-Category, ODM
        appear in both sheets).
        """
        available = [c for c in self._INFO_KEEP if c in self.info.columns]
        info_slim = (
            self.info[available]
            .copy()
            .drop_duplicates(subset=["Supplier"], keep="first")
        )

        self._merged = pd.merge(
            self.summary,
            info_slim,
            left_on=["GTK Supplier"],
            right_on=["Supplier"],
            how="left",
        )
        return self

    @property
    def records(self) -> list[dict]:
        if self._merged is None:
            raise RuntimeError("Call process() before accessing records.")
        return self._merged.to_dict(orient="records")

    # ------------------------------------------------------------------
    # Static formatting helpers (used by WordWriter and the GUI)
    # ------------------------------------------------------------------

    @staticmethod
    def format_amount(amount) -> str:
        """Return a USD-formatted string, e.g. '$5,270.50'. Returns 'N/A' for NaN."""
        try:
            v = float(amount)
            if v != v:  # NaN
                return "N/A"
            return f"${v:,.2f}"
        except (TypeError, ValueError):
            return "N/A"

    @staticmethod
    def format_date(value) -> str:
        """Return 'Month DD, YYYY', e.g. 'January 01, 2025'."""
        if isinstance(value, (datetime, date)):
            return value.strftime("%B %d, %Y")
        if isinstance(value, str):
            for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"):
                try:
                    return datetime.strptime(value, fmt).strftime("%B %d, %Y")
                except ValueError:
                    continue
        return str(value)

    @staticmethod
    def amount_to_words(amount) -> str:
        """Convert a numeric amount to uppercase legal English words.

        Example: 5270.50 → 'FIVE THOUSAND TWO HUNDRED AND SEVENTY DOLLARS AND FIFTY CENTS'
        Returns 'N/A' for NaN or non-numeric values.
        """
        try:
            v = float(amount)
            if v != v:  # NaN
                return "N/A"
        except (TypeError, ValueError):
            return "N/A"
        total_cents = round(v * 100)
        dollars = total_cents // 100
        cents = total_cents % 100

        def _cardinal(n: int) -> str:
            text = num2words(n, lang="en")
            return text.replace(",", "").replace("-", " ").upper()

        dollar_label = "DOLLAR" if dollars == 1 else "DOLLARS"
        result = f"{_cardinal(dollars)} {dollar_label}"

        if cents > 0:
            cent_label = "CENT" if cents == 1 else "CENTS"
            result += f" AND {_cardinal(cents)} {cent_label}"

        return result

    # ------------------------------------------------------------------

    @classmethod
    def is_elan_syna(cls, supplier: str) -> bool:
        """Return True when the supplier name contains Elan or Synaptics."""
        low = supplier.lower()
        return any(kw in low for kw in cls.ELAN_SYNA_KEYWORDS)
