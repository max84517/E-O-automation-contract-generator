"""GUI main window for the E&O Settlement Contract Generator.

Layout (top → bottom)
----------------------
Row 0  Status bar    – Excel filename + row counts
Row 1  Table header  – column labels
Row 2  Scroll area   – lightweight tk rows grouped by status
Row 3  Select bar    – Select All / Deselect All
Row 4  Action bar    – Refresh | progress bar | Generate Selected
Row 5  Log           – real-time generation log

Performance note: data rows use native tk.Frame / tk.Label / tk.Checkbutton
(~1 canvas object each) instead of CTk equivalents (~5 canvas objects each),
which keeps scrolling smooth for 200+ rows.

All UI updates from the background worker go through self.after(0, ...).
"""
from __future__ import annotations

import threading
import tkinter as tk
from tkinter import ttk

import customtkinter as ctk

from src.config.paths import EXCEL_FILE, OUTPUT_DIR
from src.core.loader import InfoLoader
from src.core.processor import DataProcessor
from src.core.writer import WordWriter

# Must be called at module level before any CTk widget is instantiated.
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

_COL_SUPPLIER = "GTK Supplier"
_COL_PLATFORM = "Platform"
_COL_PAYMENT  = "Actual Payment"
_COL_EFF_DATE = "Effective Date"

# (header_label, pixel_width) – kept in sync with column layout below
_COLUMNS: list[tuple[str, int]] = [
    ("",               44),
    ("GTK Supplier",  200),
    ("Platform",      150),
    ("Actual Payment", 120),
    ("Effective Date", 130),
    ("Status",        120),
]

# Dark-mode palette (approx. CTkScrollableFrame dark bg)
_SCROLL_BG   = "#2B2B2B"
_EVEN_A      = "#2B2B2B"
_ODD_A       = "#323232"
_MISSING_BG  = "#3A1E1E"   # reddish – missing data
_DONE_BG_A   = "#1E2A1E"   # greenish – already generated
_DONE_BG_B   = "#243024"
_FG_NORMAL   = "#DCE4EE"
_FG_DIM      = "#888888"
_FG_SECTION  = "#AAAAAA"
_CHK_SELECT  = "#1F538D"   # blue check mark


# ---------------------------------------------------------------------------
# Module-level helpers
# ---------------------------------------------------------------------------

def _output_filename(record: dict) -> str:
    s = str(record.get(_COL_SUPPLIER, "")).strip()
    p = str(record.get(_COL_PLATFORM, "")).strip()
    return f"Settlement Form_{s} {p}.docx"


# All 9 Word placeholder sources that must be present to generate a contract.
_REQUIRED_FIELDS: list[tuple[str, str]] = [
    (_COL_SUPPLIER,  "GTK Supplier"),
    (_COL_PLATFORM,  "Platform"),
    (_COL_PAYMENT,   "Actual Payment"),
    (_COL_EFF_DATE,  "Effective Date"),
    ("Master Agreement", "Master Agreement"),
    ("Supplier name",    "Supplier name"),
    ("Sub-Category",     "Sub-Category"),
    ("Signer",           "Signer"),
    ("Signer title",     "Signer title"),
]


def _get_missing_fields(record: dict) -> list[str]:
    """Return display names of required fields that are empty / NaN."""
    missing: list[str] = []
    for col, label in _REQUIRED_FIELDS:
        raw = record.get(col)
        # Numeric check (Actual Payment)
        if col == _COL_PAYMENT:
            try:
                f = float(raw)   # type: ignore[arg-type]
                if f != f:       # NaN
                    missing.append(label)
            except (TypeError, ValueError):
                missing.append(label)
            continue
        # Date check (Effective Date)
        if col == _COL_EFF_DATE:
            if not DataProcessor.format_date(raw):
                missing.append(label)
            continue
        # String fields
        v = str(raw).strip() if raw is not None else ""
        if not v or v.lower() in ("nan", "nat", "none"):
            missing.append(label)
    return missing


def _record_is_valid(record: dict) -> bool:
    return len(_get_missing_fields(record)) == 0


def _existing_docx_names() -> set[str]:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    return {p.name for p in OUTPUT_DIR.iterdir() if p.suffix == ".docx"}


# ---------------------------------------------------------------------------
# Data container
# ---------------------------------------------------------------------------

class _RowEntry:
    """Holds state for one data row in the table."""

    __slots__ = ("record", "frame", "var", "is_valid")

    def __init__(
        self,
        record: dict,
        frame: tk.Frame,
        var: tk.BooleanVar,
        is_valid: bool,
    ) -> None:
        self.record   = record
        self.frame    = frame
        self.var      = var
        self.is_valid = is_valid


# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------

class App(ctk.CTk):
    """Main application window."""

    def __init__(self) -> None:
        super().__init__()
        self.title("E&O Settlement Contract Generator")
        self.geometry("1000x680")
        self.minsize(800, 500)

        self._processor: DataProcessor | None = None
        self._writer = WordWriter()
        self._rows: list[_RowEntry] = []
        # All tk widgets created inside the scroll area (rows + section headers).
        # Destroyed together on each rebuild to avoid accumulation.
        self._scroll_widgets: list[tk.Widget] = []

        self._build_ui()
        self.after(200, self._load_data)
        # Keep window-level mousewheel bound so scrolling works everywhere.
        self.bind_all("<MouseWheel>", self._on_mousewheel)

    # ------------------------------------------------------------------
    # UI construction (CTk widgets only – built once)
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # Row 0 – status bar
        self._lbl_status = ctk.CTkLabel(
            self, text="Loading…", anchor="w", font=("Arial", 12)
        )
        self._lbl_status.grid(row=0, column=0, sticky="ew", padx=16, pady=(12, 2))

        # Row 1 – fixed column header
        header = ctk.CTkFrame(self, fg_color=("gray60", "gray30"), corner_radius=6)
        header.grid(row=1, column=0, sticky="ew", padx=16, pady=(2, 0))
        for label, width in _COLUMNS:
            ctk.CTkLabel(
                header, text=label, width=width, anchor="w",
                font=("Arial", 11, "bold"),
            ).pack(side="left", padx=(8, 0), pady=5)

        # Row 2 – scrollable area
        # Built from a plain tk.Canvas + tk.Scrollbar so we own the inner
        # frame directly — no private CTk attributes needed.
        scroll_outer = ctk.CTkFrame(self, corner_radius=6)
        scroll_outer.grid(row=2, column=0, sticky="nsew", padx=16, pady=0)
        scroll_outer.grid_rowconfigure(0, weight=1)
        scroll_outer.grid_columnconfigure(0, weight=1)

        self._canvas = tk.Canvas(
            scroll_outer, bg=_SCROLL_BG, highlightthickness=0, bd=0
        )
        self._scrollbar = ttk.Scrollbar(
            scroll_outer, orient="vertical", command=self._canvas.yview
        )
        self._canvas.configure(yscrollcommand=self._scrollbar.set)
        self._canvas.grid(row=0, column=0, sticky="nsew")
        self._scrollbar.grid(row=0, column=1, sticky="ns")

        self._inner_frame = tk.Frame(self._canvas, bg=_SCROLL_BG)
        self._canvas_win = self._canvas.create_window(
            (0, 0), window=self._inner_frame, anchor="nw"
        )
        self._inner_frame.bind("<Configure>", self._on_inner_configure)
        self._canvas.bind("<Configure>", self._on_canvas_resize)

        # Row 3 – select helpers
        sel_bar = ctk.CTkFrame(self, fg_color="transparent")
        sel_bar.grid(row=3, column=0, sticky="w", padx=16, pady=(4, 0))
        ctk.CTkButton(
            sel_bar, text="Select All", width=90, height=28,
            command=self._select_all,
        ).pack(side="left", padx=(0, 6))
        ctk.CTkButton(
            sel_bar, text="Deselect All", width=90, height=28,
            command=self._deselect_all,
        ).pack(side="left")

        # Row 4 – action bar
        action = ctk.CTkFrame(self, fg_color="transparent")
        action.grid(row=4, column=0, sticky="ew", padx=16, pady=6)
        action.grid_columnconfigure(1, weight=1)

        ctk.CTkButton(
            action, text="Refresh", width=80, height=32, command=self._refresh
        ).grid(row=0, column=0, padx=(0, 10))

        self._progress = ctk.CTkProgressBar(action, height=18)
        self._progress.set(0)
        self._progress.grid(row=0, column=1, sticky="ew")

        self._btn_gen = ctk.CTkButton(
            action, text="Generate Selected", width=140, height=32,
            command=self._on_generate,
        )
        self._btn_gen.grid(row=0, column=2, padx=(10, 0))

        # Row 5 – log
        self._log = ctk.CTkTextbox(self, height=130, font=("Consolas", 11))
        self._log.grid(row=5, column=0, sticky="ew", padx=16, pady=(0, 12))
        self._log.configure(state="disabled")

    # ------------------------------------------------------------------
    # Canvas / scroll helpers
    # ------------------------------------------------------------------

    def _on_inner_configure(self, _event=None) -> None:
        """Update the canvas scroll region when the inner frame changes size."""
        self._canvas.configure(scrollregion=self._canvas.bbox("all"))

    def _on_canvas_resize(self, event) -> None:
        """Keep the inner frame as wide as the canvas."""
        self._canvas.itemconfigure(self._canvas_win, width=event.width)

    def _on_mousewheel(self, event) -> None:
        """Scroll the canvas with the mouse wheel (Windows delta units)."""
        self._canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    # ------------------------------------------------------------------
    # Data loading
    # ------------------------------------------------------------------

    def _load_data(self) -> None:
        try:
            loader = InfoLoader(EXCEL_FILE).load()
            self._processor = DataProcessor(loader.summary, loader.info).process()
            self._lbl_status.configure(
                text=(
                    f"Excel: {EXCEL_FILE.name}  |  "
                    f"Summary: {loader.summary_count} rows  |  "
                    f"Info: {loader.info_count} rows"
                )
            )
            self._rebuild_table()
        except FileNotFoundError:
            self._lbl_status.configure(text=f"⚠ Excel file not found: {EXCEL_FILE}")
        except Exception as exc:  # noqa: BLE001
            self._lbl_status.configure(text=f"⚠ Load failed: {exc}")

    # ------------------------------------------------------------------
    # Table rebuild
    # ------------------------------------------------------------------

    def _rebuild_table(self) -> None:
        """Destroy all scroll-area widgets and rebuild sorted by status."""
        for w in self._scroll_widgets:
            w.destroy()
        self._scroll_widgets.clear()
        self._rows.clear()

        if self._processor is None:
            return

        existing = _existing_docx_names()

        # ── Categorise ──────────────────────────────────────────────────
        pending: list[dict] = []
        missing: list[dict] = []
        done:    list[dict] = []

        for record in self._processor.records:
            fname = _output_filename(record)
            if fname in existing:
                done.append(record)
            elif _record_is_valid(record):
                pending.append(record)
            else:
                missing.append(record)

        # ── Render groups ───────────────────────────────────────────────
        groups = [
            (f"⬜  Pending  ({len(pending)})",      pending,  _EVEN_A,     _ODD_A,     "Pending"),
            (f"⚠  Missing Data  ({len(missing)})",  missing,  _MISSING_BG, _MISSING_BG, "Missing data"),
            (f"✅  Generated  ({len(done)})",        done,     _DONE_BG_A,  _DONE_BG_B, "Generated"),
        ]

        data_idx = 0   # counter for alternating row colour
        inner = self._inner_frame
        for sec_label, records, bg_a, bg_b, status_txt in groups:
            if not records:
                continue

            # Section header label
            sec_frm = tk.Frame(inner, bg=_SCROLL_BG)
            sec_frm.pack(fill="x", padx=2, pady=(8, 1))
            ctk.CTkLabel(
                sec_frm, text=sec_label,
                fg_color=_SCROLL_BG, text_color=_FG_SECTION,
                font=("Arial", 11, "bold"), anchor="w",
            ).pack(side="left", padx=8, pady=2)
            self._scroll_widgets.append(sec_frm)

            # Thin separator
            sep = tk.Frame(inner, bg="#444444", height=1)
            sep.pack(fill="x", padx=4)
            self._scroll_widgets.append(sep)

            for record in records:
                is_valid = _record_is_valid(record)
                is_done  = (status_txt == "Generated")

                row_bg = bg_a if data_idx % 2 == 0 else bg_b
                data_idx += 1

                # Default check state: valid+pending → checked, everything else → unchecked
                var = tk.BooleanVar(value=(is_valid and not is_done))

                row = tk.Frame(inner, bg=row_bg)
                row.pack(fill="x", padx=2, pady=1)
                self._scroll_widgets.append(row)

                # Checkbox — CTkCheckBox so CTk scales it identically to the header
                ctk.CTkCheckBox(
                    row,
                    text="",
                    variable=var,
                    width=_COLUMNS[0][1],
                    checkbox_width=20,
                    checkbox_height=20,
                    state="normal" if is_valid else "disabled",
                    fg_color=_CHK_SELECT,
                ).pack(side="left", padx=(8, 0), pady=4)

                sup      = str(record.get(_COL_SUPPLIER, "")).strip()
                plat     = str(record.get(_COL_PLATFORM, "")).strip()
                pay_str  = DataProcessor.format_amount(record.get(_COL_PAYMENT))
                eff_date = DataProcessor.format_date(record.get(_COL_EFF_DATE, ""))
                text_color = _FG_NORMAL if is_valid else _FG_DIM

                # CTkLabel handles DPI scaling automatically — identical to the header.
                for text, col_width in (
                    (sup,      _COLUMNS[1][1]),
                    (plat,     _COLUMNS[2][1]),
                    (pay_str,  _COLUMNS[3][1]),
                    (eff_date, _COLUMNS[4][1]),
                ):
                    ctk.CTkLabel(
                        row, text=text,
                        font=("Arial", 11),
                        fg_color=row_bg,
                        text_color=text_color,
                        anchor="w",
                        width=col_width,
                    ).pack(side="left", padx=(8, 0), pady=2)

                icon = "✅" if is_done else ("⚠" if not is_valid else "⬜")
                ctk.CTkLabel(
                    row, text=f"{icon} {status_txt}",
                    font=("Arial", 11),
                    fg_color=row_bg,
                    text_color=text_color,
                    anchor="w",
                    width=_COLUMNS[5][1],
                ).pack(side="left", padx=(8, 0), pady=2)

                self._rows.append(_RowEntry(record, row, var, is_valid))

    # ------------------------------------------------------------------
    # Button callbacks
    # ------------------------------------------------------------------

    def _select_all(self) -> None:
        """Select all rows that are eligible for generation."""
        for e in self._rows:
            if e.is_valid:
                e.var.set(True)

    def _deselect_all(self) -> None:
        for e in self._rows:
            e.var.set(False)

    def _refresh(self) -> None:
        self._load_data()
        self._log_write("🔄 Refreshed\n")

    # ------------------------------------------------------------------
    # Contract generation
    # ------------------------------------------------------------------

    def _on_generate(self) -> None:
        selected = [e for e in self._rows if e.is_valid and e.var.get()]
        if not selected:
            self._log_write("⚠ Please select at least one item\n")
            return
        self._btn_gen.configure(state="disabled")
        self._progress.set(0)
        threading.Thread(target=self._worker, args=(selected,), daemon=True).start()

    def _worker(self, selected: list[_RowEntry]) -> None:
        """Background worker — all UI updates go through self.after()."""
        total = len(selected)
        ok = fail = 0

        for idx, entry in enumerate(selected, start=1):
            record   = entry.record
            supplier = str(record.get(_COL_SUPPLIER, "")).strip()
            platform = str(record.get(_COL_PLATFORM, "")).strip()

            # Double-check required fields (guard against race with refresh)
            missing = _get_missing_fields(record)
            if missing:
                self.after(0, self._log_write,
                           f"❌ {supplier} {platform}: Missing fields — {', '.join(missing)}\n")
                fail += 1
                self.after(0, self._progress.set, idx / total)
                continue

            try:
                out = self._writer.write(
                    record,
                    is_elan_syna=DataProcessor.is_elan_syna(supplier),
                    format_amount=DataProcessor.format_amount,
                    amount_to_words=DataProcessor.amount_to_words,
                    format_date=DataProcessor.format_date,
                )
                self.after(0, self._log_write, f"✅ {out.name}\n")
                ok += 1
            except Exception as exc:  # noqa: BLE001
                self.after(0, self._log_write, f"❌ {supplier} {platform}: {exc}\n")
                fail += 1

            self.after(0, self._progress.set, idx / total)

        self.after(0, self._log_write, f"\nDone: {ok} succeeded, {fail} failed.\n")
        self.after(0, self._rebuild_table)
        self.after(0, lambda: self._btn_gen.configure(state="normal"))

    # ------------------------------------------------------------------
    # Log helper
    # ------------------------------------------------------------------

    def _log_write(self, text: str) -> None:
        self._log.configure(state="normal")
        self._log.insert("end", text)
        self._log.see("end")
        self._log.configure(state="disabled")
