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
        "Address",
    ]

    def __init__(self, summary: pd.DataFrame, info: pd.DataFrame) -> None:
        self.summary = summary
        self.info = info
        self._merged: pd.DataFrame | None = None

    # ------------------------------------------------------------------
    def process(self) -> "DataProcessor":
        """Merge Summary → Info with GBU-aware Signer selection.

        Most suppliers have one Info row  → join on GTK Supplier only.
        Suppliers with multiple Info rows differentiated by GBU (e.g. Chicony NB / DT)
        → join on (GTK Supplier, GBU suffix).

        Summary.GBU looks like 'cNB', 'bDT' etc.
        Info.GBU looks like 'NB', 'DT' (or NaN for single-Signer suppliers).
        We normalize Summary.GBU to the last 2 chars for comparison.
        """
        keep_cols = [c for c in self._INFO_KEEP if c in self.info.columns]

        # Include Info.GBU temporarily for the split; remove it from final output.
        extra = ["GBU"] if "GBU" in self.info.columns else []
        info_work = self.info[keep_cols + extra].copy()

        # ── Normalise all string data columns in Info (strip whitespace) ──
        for col in keep_cols:
            if info_work[col].dtype == object:
                info_work[col] = info_work[col].str.strip()

        # Normalised join keys (case-insensitive, no surrounding whitespace)
        info_work["_supp_norm"] = info_work["Supplier"].str.strip().str.lower()
        if extra:
            info_work["_gbu_norm_info"] = (
                info_work["GBU"].astype(str).str.strip().str.upper()
            )

        # ── Suppliers that have GBU-differentiated rows in Info ─────────
        if extra:
            info_has_gbu = info_work[info_work["GBU"].notna()].copy()
            specific_suppliers_norm: set[str] = set(info_has_gbu["_supp_norm"].unique())
        else:
            info_has_gbu = pd.DataFrame()
            specific_suppliers_norm = set()

        # ── All other suppliers: one Info row per Supplier ───────────────
        info_nogbu = (
            info_work[info_work["GBU"].isna()][keep_cols + ["_supp_norm"]]
            if extra
            else info_work[keep_cols + ["_supp_norm"]].copy()
        )
        info_nogbu = info_nogbu.drop_duplicates(subset=["_supp_norm"], keep="first")

        # Normalize Summary GBU → last 2 chars upper (cNB→NB, bDT→DT)
        # Also normalize GTK Supplier for case-insensitive join
        summary_work = self.summary.copy()
        # Strip whitespace from all Summary string columns
        for col in summary_work.columns:
            if summary_work[col].dtype == object:
                summary_work[col] = summary_work[col].str.strip()
        summary_work["_gbu_norm"] = (
            summary_work["GBU"].astype(str).str.strip().str[-2:].str.upper()
        )
        summary_work["_supp_norm"] = (
            summary_work["GTK Supplier"].str.strip().str.lower()
        )

        results = []

        # ── Group 1: GBU-specific suppliers (e.g. Chicony) ──────────────
        if specific_suppliers_norm:
            s_specific = summary_work[
                summary_work["_supp_norm"].isin(specific_suppliers_norm)
            ].copy()
            m1 = pd.merge(
                s_specific,
                info_has_gbu.rename(columns={"_gbu_norm_info": "_info_gbu"}),
                left_on=["_supp_norm", "_gbu_norm"],
                right_on=["_supp_norm", "_info_gbu"],
                how="left",
            ).drop(columns=["_info_gbu", "GBU"], errors="ignore")
            results.append(m1)

        # ── Group 2: all other suppliers ────────────────────────────────
        s_others = summary_work[
            ~summary_work["_supp_norm"].isin(specific_suppliers_norm)
        ]
        m2 = pd.merge(
            s_others,
            info_nogbu,
            on=["_supp_norm"],
            how="left",
        )
        results.append(m2)

        self._merged = pd.concat(results, ignore_index=True).drop(
            columns=["_gbu_norm", "_supp_norm", "_gbu_norm_info", "GBU_y"], errors="ignore"
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
        """Return 'Month DD, YYYY', e.g. 'January 01, 2025'.
        Returns empty string for NaT, None, or unparseable values.
        """
        # pandas NaT passes isinstance(datetime) but raises on strftime
        import pandas as pd
        if value is None or value is pd.NaT:
            return ""
        try:
            if isinstance(value, (datetime, date)):
                return value.strftime("%B %d, %Y")
        except (ValueError, TypeError, AttributeError):
            pass
        if isinstance(value, str):
            for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"):
                try:
                    return datetime.strptime(value, fmt).strftime("%B %d, %Y")
                except ValueError:
                    continue
        return ""

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
