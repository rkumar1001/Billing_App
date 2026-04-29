from __future__ import annotations

import os
import platform
import subprocess
import traceback
from datetime import date, datetime
from pathlib import Path
from tkinter import messagebox
from typing import TYPE_CHECKING

import customtkinter as ctk

from ..services.auto_detect import ExcelDetection, WordDetection, WordLoc
from ..services.calendar_util import MONTH_NAMES, days_in_month
from ..services.folder_analyzer import FolderAnalysis
from ..services.invoice_generator import (
    GenerateRequest,
    GeneratorError,
    generate,
    preview,
)
from .dialogs import ask_directory
from .cell_picker import ExcelCellPicker, WordLocationPicker

if TYPE_CHECKING:
    from .app import BillingApp


def open_in_os(path: str | Path) -> None:
    p = str(path)
    if not p or not os.path.exists(p):
        return
    try:
        if platform.system() == "Windows":
            os.startfile(p)  # type: ignore[attr-defined]
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", p])
        else:
            subprocess.Popen(["xdg-open", p])
    except OSError:
        pass


class GeneratorScreen(ctk.CTkFrame):
    def __init__(self, parent, app: "BillingApp") -> None:
        super().__init__(parent, fg_color="transparent")
        self.app = app

        self._analysis: FolderAnalysis | None = None
        self._excel_det: ExcelDetection | None = None
        self._word_det: WordDetection | None = None

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=24, pady=(20, 8))
        ctk.CTkLabel(
            header, text="Generate next month's invoice",
            font=ctk.CTkFont(size=24, weight="bold"),
        ).grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(
            header,
            text="Point at the previous month's folder — the app copies it, bumps dates, renames files, and saves the new folder next to the source.",
            font=ctk.CTkFont(size=12),
            text_color=("gray40", "gray70"),
            wraplength=900, justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.grid(row=1, column=0, sticky="nsew", padx=16, pady=(0, 16))
        body.grid_columnconfigure(0, weight=1)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)

        self._form = ctk.CTkScrollableFrame(body, label_text="Inputs")
        self._form.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self._form.grid_columnconfigure(1, weight=1)

        self._preview = ctk.CTkFrame(body, corner_radius=10)
        self._preview.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        self._preview.grid_columnconfigure(0, weight=1)
        self._preview.grid_rowconfigure(99, weight=1)

        self._build_form()
        self._build_preview()

        last = getattr(app.cfg, "last_source_folder", "") or app.cfg.default_output_root
        if last and Path(last).is_dir():
            self._folder_var.set(last)
            self._refresh_analysis()

    # ---- build -------------------------------------------------------

    def _build_form(self) -> None:
        f = self._form
        r = [0]

        def grow() -> int:
            r[0] += 1
            return r[0]

        # Source folder.
        ctk.CTkLabel(
            f, text="Source folder (previous month)",
            font=ctk.CTkFont(size=12, weight="bold"),
        ).grid(row=r[0], column=0, columnspan=2, sticky="w", padx=8, pady=(6, 2))
        row_frame = ctk.CTkFrame(f, fg_color="transparent")
        row_frame.grid(row=grow(), column=0, columnspan=2, sticky="ew", padx=8, pady=2)
        row_frame.grid_columnconfigure(0, weight=1)
        self._folder_var = ctk.StringVar()
        self._folder_var.trace_add("write", lambda *_: self._schedule_refresh())
        ctk.CTkEntry(row_frame, textvariable=self._folder_var).grid(row=0, column=0, sticky="ew")
        ctk.CTkButton(
            row_frame, text="Browse…", width=90,
            command=self._pick_source_folder,
        ).grid(row=0, column=1, padx=(6, 0))

        self._analysis_label = ctk.CTkLabel(
            f, text="", wraplength=380, justify="left", anchor="w",
            font=ctk.CTkFont(size=11),
            text_color=("gray20", "gray80"),
        )
        self._analysis_label.grid(row=grow(), column=0, columnspan=2, sticky="w", padx=8, pady=(2, 10))

        # Target month / year.
        ctk.CTkLabel(
            f, text="Target month",
            font=ctk.CTkFont(size=12, weight="bold"),
        ).grid(row=grow(), column=0, columnspan=2, sticky="w", padx=8, pady=(8, 2))

        today = date.today()
        self._month_var = ctk.StringVar(value=MONTH_NAMES[today.month - 1])
        self._year_var = ctk.StringVar(value=str(today.year))

        month_row = ctk.CTkFrame(f, fg_color="transparent")
        month_row.grid(row=grow(), column=0, columnspan=2, sticky="ew", padx=8, pady=2)
        ctk.CTkOptionMenu(
            month_row, variable=self._month_var, values=MONTH_NAMES,
            command=lambda *_: self._recompute(), width=120,
        ).pack(side="left")
        ctk.CTkEntry(month_row, textvariable=self._year_var, width=70).pack(side="left", padx=(6, 0))
        ctk.CTkButton(
            month_row, text="Next", width=60,
            command=self._set_next_month,
        ).pack(side="left", padx=(6, 0))
        ctk.CTkButton(
            month_row, text="Today", width=60,
            command=self._set_current_month,
        ).pack(side="left", padx=(6, 0))
        self._year_var.trace_add("write", lambda *_: self._recompute())

        # Invoice date.
        ctk.CTkLabel(
            f, text="Invoice date (YYYY-MM-DD)",
            font=ctk.CTkFont(size=12, weight="bold"),
        ).grid(row=grow(), column=0, columnspan=2, sticky="w", padx=8, pady=(8, 2))
        self._invoice_date_var = ctk.StringVar(value=today.strftime("%Y-%m-%d"))
        ctk.CTkEntry(f, textvariable=self._invoice_date_var).grid(
            row=grow(), column=0, columnspan=2, sticky="ew", padx=8, pady=2
        )

        # Optional overrides.
        ctk.CTkLabel(
            f, text="Overrides (optional — blank uses Excel values)",
            font=ctk.CTkFont(size=12, weight="bold"),
            wraplength=380, justify="left",
        ).grid(row=grow(), column=0, columnspan=2, sticky="w", padx=8, pady=(12, 2))
        self._rate_var = ctk.StringVar()
        self._hpd_var = ctk.StringVar()
        self._guards_var = ctk.StringVar()
        self._invoice_num_var = ctk.StringVar()
        for lbl, var in (
            ("Hourly rate",       self._rate_var),
            ("Hours per day",     self._hpd_var),
            ("Guards",            self._guards_var),
            ("Invoice number (override)", self._invoice_num_var),
        ):
            ctk.CTkLabel(f, text=lbl).grid(row=grow(), column=0, sticky="w", padx=8, pady=2)
            entry = ctk.CTkEntry(f, textvariable=var)
            entry.grid(row=r[0], column=1, sticky="ew", padx=8, pady=2)
            var.trace_add("write", lambda *_: self._recompute())

        # Overwrite checkbox.
        self._overwrite_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            f, text="Overwrite target folder if it exists",
            variable=self._overwrite_var,
        ).grid(row=grow(), column=0, columnspan=2, sticky="w", padx=8, pady=(12, 4))

        # Buttons.
        btn_row = ctk.CTkFrame(f, fg_color="transparent")
        btn_row.grid(row=grow(), column=0, columnspan=2, sticky="ew", padx=8, pady=(16, 2))
        self._generate_btn = ctk.CTkButton(
            btn_row, text="Generate", command=self._generate, height=38,
        )
        self._generate_btn.pack(side="left", fill="x", expand=True)

        btn_row2 = ctk.CTkFrame(f, fg_color="transparent")
        btn_row2.grid(row=grow(), column=0, columnspan=2, sticky="ew", padx=8, pady=(4, 8))
        ctk.CTkButton(
            btn_row2, text="Re-detect", command=self._refresh_analysis,
            fg_color="transparent", border_width=1, height=32,
            text_color=("gray10", "gray90"),
            border_color=("gray60", "gray40"),
            hover_color=("gray85", "gray25"),
        ).pack(side="left", fill="x", expand=True)

    def _build_preview(self) -> None:
        p = self._preview
        ctk.CTkLabel(
            p, text="Preview",
            font=ctk.CTkFont(size=16, weight="bold"),
        ).grid(row=0, column=0, sticky="w", padx=14, pady=(14, 6))

        self._preview_lines: dict[str, ctk.CTkLabel] = {}
        for i, key in enumerate([
            "Folder detected", "Source month", "Invoice number (source)",
            "Target folder", "Target period",
            "Days in month", "Hours per day × guards",
            "Hourly rate", "Total hours", "Grand total",
            "Next invoice number",
        ]):
            lbl = ctk.CTkLabel(p, text=f"{key}:  —", anchor="w", justify="left")
            lbl.grid(row=i + 1, column=0, sticky="ew", padx=14, pady=2)
            self._preview_lines[key] = lbl

        self._warning_box = ctk.CTkLabel(
            p, text="", text_color=("gray35", "gray75"), anchor="w", justify="left",
            wraplength=440, font=ctk.CTkFont(size=11),
        )
        self._warning_box.grid(row=50, column=0, sticky="ew", padx=14, pady=(12, 4))

        self._result_box = ctk.CTkLabel(
            p, text="", text_color=("gray20", "gray85"), anchor="w", justify="left",
            wraplength=440, font=ctk.CTkFont(size=12),
        )
        self._result_box.grid(row=60, column=0, sticky="ew", padx=14, pady=(8, 14))

    # ---- events ------------------------------------------------------

    def _pick_source_folder(self) -> None:
        start = (self._folder_var.get() or self.app.cfg.default_output_root or "").strip()
        initialdir = start if Path(start).is_dir() else None
        parent = self.winfo_toplevel()
        folder = ask_directory(parent=parent, initialdir=initialdir, title="Choose source folder")
        if folder:
            self._folder_var.set(folder)

    def _set_next_month(self) -> None:
        if self._analysis and self._analysis.source_month:
            m = self._analysis.source_month
            y = self._analysis.source_year or int(self._year_var.get() or date.today().year)
        else:
            today = date.today()
            m, y = today.month, today.year
        m += 1
        if m > 12:
            m = 1
            y += 1
        self._month_var.set(MONTH_NAMES[m - 1])
        self._year_var.set(str(y))
        self._recompute()

    def _set_current_month(self) -> None:
        today = date.today()
        self._month_var.set(MONTH_NAMES[today.month - 1])
        self._year_var.set(str(today.year))
        self._recompute()

    def _schedule_refresh(self) -> None:
        # Debounce: only fire after the path looks like a real directory.
        path = self._folder_var.get().strip()
        if path and Path(path).is_dir():
            self.after(250, self._refresh_analysis)

    def _refresh_analysis(self) -> None:
        path = self._folder_var.get().strip()
        if not path:
            self._analysis_label.configure(text="")
            return
        if not Path(path).is_dir():
            self._analysis_label.configure(
                text=f"Not a folder: {path}", text_color=("red", "#ff8080"),
            )
            return
        try:
            analysis, excel_det, word_det = preview(Path(path))
        except Exception as e:  # noqa: BLE001
            self._analysis_label.configure(
                text=f"Analysis failed: {e}", text_color=("red", "#ff8080"),
            )
            return

        self._analysis = analysis
        self._excel_det = excel_det
        self._word_det = word_det
        self.app.cfg.last_source_folder = path  # type: ignore[attr-defined]
        self.app.save_config()

        parts: list[str] = []
        if analysis.excel_path:
            if analysis.excel_is_legacy_xls:
                parts.append(
                    f"Excel: {analysis.excel_path.name}  (legacy .xls — "
                    "will be auto-converted to .xlsx in the new folder)"
                )
            else:
                parts.append(f"Excel: {analysis.excel_path.name}")
        else:
            parts.append("No Excel (.xlsx or .xls) found")
        if analysis.word_path:
            parts.append(f"Word: {analysis.word_path.name}")
        else:
            parts.append("No Word (.docx) found")
        if analysis.invoice_number is not None:
            parts.append(
                f"Invoice: {analysis.invoice_prefix}{analysis.invoice_number} → {analysis.invoice_prefix}{analysis.next_invoice_number}"
            )
        if analysis.source_month:
            parts.append(
                f"Source month: {MONTH_NAMES[analysis.source_month - 1]}"
                + (f" {analysis.source_year}" if analysis.source_year else "")
            )
        if analysis.warnings:
            parts.append("⚠ " + "; ".join(analysis.warnings))
        self._analysis_label.configure(
            text="\n".join(parts), text_color=("gray30", "gray80"),
        )

        if self._excel_det and self._excel_det.detected_year and not self._year_var.get():
            self._year_var.set(str(self._excel_det.detected_year))

        self._update_warning_box()
        self._recompute()

    def _update_warning_box(self) -> None:
        warnings: list[str] = []
        if self._excel_det:
            excel_missing = self._excel_det.missing_fields()
            if excel_missing:
                warnings.append(
                    "Some Excel fields need manual selection: "
                    + ", ".join(excel_missing)
                    + ". Use the Picker buttons below, or fill in the "
                    "Overrides fields manually."
                )
        if self._word_det:
            word_missing = self._word_det.missing_fields()
            if word_missing:
                warnings.append(
                    "Some Word fields need manual selection: "
                    + ", ".join(word_missing)
                    + ". Use the Picker buttons below."
                )
        if not (self._excel_det and self._word_det):
            warnings.append("Select a source folder to begin.")
        self._warning_box.configure(text="\n".join(warnings))
        self._ensure_picker_buttons()

    def _ensure_picker_buttons(self) -> None:
        # Lazily (re)build the picker button row.
        if hasattr(self, "_picker_bar") and self._picker_bar.winfo_exists():
            self._picker_bar.destroy()
        self._picker_bar = ctk.CTkFrame(self._preview, fg_color="transparent")
        self._picker_bar.grid(row=51, column=0, sticky="ew", padx=14, pady=(2, 4))

        if not (self._analysis and self._analysis.excel_path):
            return

        excel_missing = self._excel_det.missing_fields() if self._excel_det else []
        word_missing = self._word_det.missing_fields() if self._word_det else []

        def add_btn(label: str, cmd) -> None:
            ctk.CTkButton(
                self._picker_bar, text=label, command=cmd,
                fg_color="transparent", border_width=1, height=28,
                text_color=("gray10", "gray90"),
                border_color=("gray60", "gray40"),
                hover_color=("gray85", "gray25"),
            ).pack(side="top", anchor="w", pady=2)

        for field in excel_missing:
            add_btn(
                f"Pick Excel cell for '{field}'",
                lambda f=field: self._pick_excel_field(f),
            )
        for field in word_missing:
            add_btn(
                f"Pick Word cell for '{field}'",
                lambda f=field: self._pick_word_field(f),
            )

    def _pick_excel_field(self, field_name: str) -> None:
        if not (self._analysis and self._analysis.excel_path):
            return
        picker = ExcelCellPicker(
            self, self._analysis.excel_path,
            title=f"Pick cell for {field_name}",
        )
        self.wait_window(picker)
        chosen = picker.selected_cell
        if chosen and self._excel_det:
            sheet, cell = chosen
            self._excel_det.sheet = sheet
            setattr(self._excel_det, field_name, cell)
            self._update_warning_box()

    def _pick_word_field(self, field_name: str) -> None:
        if not (self._analysis and self._analysis.word_path):
            return
        picker = WordLocationPicker(
            self, self._analysis.word_path,
            title=f"Pick Word location for {field_name}",
        )
        self.wait_window(picker)
        chosen = picker.selected_location
        if chosen and self._word_det:
            setattr(self._word_det, field_name, chosen)
            self._update_warning_box()

    def _recompute(self) -> None:
        if not self._analysis:
            return
        try:
            year = int(self._year_var.get())
            month = MONTH_NAMES.index(self._month_var.get()) + 1
        except (ValueError, IndexError):
            return

        det = self._excel_det
        rate = _parse_float(self._rate_var.get(), det.hourly_rate if det else 0) or 0.0
        hpd = _parse_float(self._hpd_var.get(), det.hours_per_day if det else 0) or 0.0
        guards = int(_parse_float(
            self._guards_var.get(), det.guards if det else 1,
        ) or 1)

        n_days = days_in_month(year, month)
        total_hours = round(n_days * hpd * guards, 2)
        grand_total = round(total_hours * rate, 2)

        invoice_int = None
        override = self._invoice_num_var.get().strip()
        if override.isdigit():
            invoice_int = int(override)
        elif self._analysis.invoice_number is not None:
            invoice_int = self._analysis.next_invoice_number

        target_name = _swap_month(
            self._analysis.folder.name,
            self._analysis.folder_month_token,
            month,
        )

        from pathlib import Path as _P
        lines = self._preview_lines
        lines["Folder detected"].configure(text=f"Folder detected:  {self._analysis.folder}")
        if self._analysis.source_month:
            lines["Source month"].configure(
                text=f"Source month:  {MONTH_NAMES[self._analysis.source_month - 1]} "
                     f"{self._analysis.source_year or ''}".rstrip()
            )
        else:
            lines["Source month"].configure(text="Source month:  (unknown)")
        if self._analysis.invoice_number is not None:
            lines["Invoice number (source)"].configure(
                text=f"Invoice number (source):  {self._analysis.invoice_prefix}{self._analysis.invoice_number}"
            )
        lines["Target folder"].configure(
            text=f"Target folder:  {(_P(self._analysis.folder.parent) / target_name)}"
        )
        lines["Target period"].configure(
            text=f"Target period:  {MONTH_NAMES[month - 1]} {year}"
        )
        lines["Days in month"].configure(text=f"Days in month:  {n_days}")
        lines["Hours per day × guards"].configure(text=f"Hours per day × guards:  {hpd:g} × {guards}")
        lines["Hourly rate"].configure(text=f"Hourly rate:  ${rate:,.2f}")
        lines["Total hours"].configure(text=f"Total hours:  {total_hours:g}")
        lines["Grand total"].configure(text=f"Grand total:  ${grand_total:,.2f}")
        if invoice_int is not None:
            lines["Next invoice number"].configure(
                text=f"Next invoice number:  {self._analysis.invoice_prefix}{invoice_int}"
            )

    def _generate(self) -> None:
        if not self._analysis:
            messagebox.showerror("No source", "Pick a source folder first.", parent=self)
            return
        missing = self._analysis.missing()
        if missing:
            messagebox.showerror(
                "Incomplete source folder",
                "Source folder is missing: " + ", ".join(missing),
                parent=self,
            )
            return
        try:
            year = int(self._year_var.get())
            month = MONTH_NAMES.index(self._month_var.get()) + 1
            issue_date = datetime.strptime(self._invoice_date_var.get(), "%Y-%m-%d").date()
        except ValueError as e:
            messagebox.showerror("Invalid input", str(e), parent=self)
            return

        override = self._invoice_num_var.get().strip()
        explicit_num = int(override) if override.isdigit() else None

        req = GenerateRequest(
            source_folder=self._analysis.folder,
            target_month=month,
            target_year=year,
            invoice_date=issue_date,
            hourly_rate=_parse_float(self._rate_var.get(), None),
            hours_per_day=_parse_float(self._hpd_var.get(), None),
            guards=(int(float(self._guards_var.get())) if self._guards_var.get().strip() else None),
            explicit_invoice_number=explicit_num,
            excel_overrides=self._excel_det,
            word_overrides=self._word_det,
            overwrite=self._overwrite_var.get(),
        )

        try:
            result = generate(req)
        except GeneratorError as e:
            messagebox.showerror("Generation failed", str(e), parent=self)
            return
        except Exception as e:  # noqa: BLE001
            messagebox.showerror(
                "Unexpected error",
                f"{type(e).__name__}: {e}\n\n{traceback.format_exc(limit=5)}",
                parent=self,
            )
            return

        lines = [
            f"Created folder: {result.copied_folder}",
            f"Excel:   {result.excel_path or '—'}",
            f"Word:    {result.word_path or '—'}",
            f"Invoice: {result.invoice_number}",
            f"Total hours: {result.total_hours:g}   Grand total: ${result.grand_total:,.2f}",
        ]
        if result.unresolved_excel or result.unresolved_word:
            lines.append("")
            lines.append(
                "Some fields couldn't be updated because we didn't know where they live: "
                + ", ".join(result.unresolved_excel + result.unresolved_word)
                + ". Use the picker buttons, then re-generate."
            )

        self._result_box.configure(text="\n".join(lines))

        if messagebox.askyesno(
            "Done",
            f"Invoice {result.invoice_number} generated. Open the folder?",
            parent=self,
        ):
            open_in_os(result.copied_folder)


def _parse_float(raw: str, fallback):
    if not raw or not raw.strip():
        return fallback
    try:
        return float(raw)
    except ValueError:
        return fallback


def _swap_month(name: str, token: str, target_month: int) -> str:
    import re as _re
    target_name = MONTH_NAMES[target_month - 1]
    if token:
        return _re.sub(_re.escape(token), target_name, name, flags=_re.IGNORECASE)
    return f"{name} → {target_name}"
