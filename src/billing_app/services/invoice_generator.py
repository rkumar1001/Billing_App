"""Folder-in, folder-out invoice duplicator.

Given a source folder (e.g. `.../Bella Vista/March/`) and a target month,
this module:
  1. Copies the whole folder next to itself with the target month name.
  2. Renames the Excel and Word files to reflect the new month and the next
     invoice number (Bella16 → Bella17).
  3. Auto-detects field locations on first contact (cached in
     `.billingapp.json` inside the source folder) and applies updates in the
     copy: month number, invoice date, dates column in Excel; invoice number,
     invoice date, billing period, total hours, grand total in Word.

User-provided values (target month, invoice date, and optional overrides for
rate / hours / guards) drive the updates; everything else is inferred.
"""
from __future__ import annotations

import platform
import shutil
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Any

from docx import Document
from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_from_string

from .auto_detect import (
    DetectionCache,
    ExcelDetection,
    WordDetection,
    WordLoc,
    detect_excel,
    detect_word,
    load_cache,
    save_cache,
)
from .calendar_util import MONTH_NAMES, days_in_month, month_name
from .folder_analyzer import FolderAnalysis, analyze, rename_for_target


@dataclass
class GenerateRequest:
    source_folder: Path
    target_month: int
    target_year: int
    invoice_date: date
    target_folder_name: str | None = None
    hourly_rate: float | None = None
    hours_per_day: float | None = None
    guards: int | None = None
    explicit_invoice_number: int | None = None
    excel_overrides: ExcelDetection | None = None
    word_overrides: WordDetection | None = None
    overwrite: bool = False


@dataclass
class GenerateResult:
    copied_folder: Path
    excel_path: Path | None
    word_path: Path | None
    invoice_number: str
    invoice_number_int: int
    total_hours: float
    grand_total: float
    analysis: FolderAnalysis
    excel_detection: ExcelDetection
    word_detection: WordDetection
    unresolved_excel: list[str] = field(default_factory=list)
    unresolved_word: list[str] = field(default_factory=list)


class GeneratorError(Exception):
    pass


def preview(source: Path) -> tuple[FolderAnalysis, ExcelDetection | None, WordDetection | None]:
    """Run analysis + detection without making any changes. Used for UI."""
    analysis = analyze(Path(source))

    excel_det: ExcelDetection | None = None
    word_det: WordDetection | None = None

    cache = load_cache(analysis.folder)
    if cache:
        excel_det = cache.excel
        word_det = cache.word

    if analysis.excel_path:
        try:
            detected = detect_excel(analysis.excel_path)
        except Exception as e:  # noqa: BLE001
            detected = ExcelDetection()
            detected.warnings.append(f"Excel detection failed: {e}")
        if excel_det is None:
            excel_det = detected
        else:
            for attr in (
                "sheet", "month_number_cell", "invoice_date_cell",
                "dates_column_start", "dates_row_end",
            ):
                if getattr(excel_det, attr) is None:
                    setattr(excel_det, attr, getattr(detected, attr))
            # Always refresh derived fields from the live detection.
            excel_det.detected_month = detected.detected_month
            excel_det.detected_year = detected.detected_year
            excel_det.hourly_rate = detected.hourly_rate
            excel_det.hours_per_day = detected.hours_per_day
            excel_det.guards = detected.guards

    if analysis.word_path and analysis.invoice_number is not None:
        invoice_str = f"{analysis.invoice_prefix}{analysis.invoice_number}"
        try:
            detected_word = detect_word(analysis.word_path, invoice_str)
        except Exception as e:  # noqa: BLE001
            detected_word = WordDetection()
            detected_word.warnings.append(f"Word detection failed: {e}")
        if word_det is None:
            word_det = detected_word
        else:
            for attr in ("invoice_number", "invoice_date", "billing_period",
                          "total_hours", "grand_total"):
                if getattr(word_det, attr) is None:
                    setattr(word_det, attr, getattr(detected_word, attr))

    return analysis, excel_det, word_det


def generate(req: GenerateRequest) -> GenerateResult:
    source = Path(req.source_folder)
    if not source.is_dir():
        raise GeneratorError(f"Source folder does not exist: {source}")

    analysis = analyze(source)
    missing = analysis.missing()
    if missing:
        raise GeneratorError(
            "Source folder is missing required items: " + ", ".join(missing)
        )

    # Decide target folder name.
    target_name = req.target_folder_name
    if not target_name:
        target_name = _swap_month_in_name(
            source.name, analysis.folder_month_token, req.target_month
        )
    target_folder = source.parent / target_name
    if target_folder.exists():
        if not req.overwrite:
            raise GeneratorError(
                f"Target folder already exists: {target_folder}. "
                f"Enable overwrite to replace."
            )
        shutil.rmtree(target_folder)

    shutil.copytree(source, target_folder)

    # Determine next invoice number.
    invoice_int = (
        req.explicit_invoice_number
        if req.explicit_invoice_number is not None
        else analysis.next_invoice_number
    )

    new_excel, new_word = rename_for_target(
        analysis, target_folder, req.target_month, req.target_year, invoice_int,
    )

    # Resolve detection: overrides > cache > fresh detection.
    cache = load_cache(source)
    excel_det = req.excel_overrides or (cache.excel if cache else None)
    if excel_det is None and analysis.excel_path:
        excel_det = detect_excel(analysis.excel_path)
    if excel_det is None:
        excel_det = ExcelDetection()

    word_det = req.word_overrides or (cache.word if cache else None)
    if word_det is None and analysis.word_path and analysis.invoice_number is not None:
        word_det = detect_word(
            analysis.word_path,
            f"{analysis.invoice_prefix}{analysis.invoice_number}",
        )
    if word_det is None:
        word_det = WordDetection()

    # Compute totals.
    n_days = days_in_month(req.target_year, req.target_month)
    rate = _first_not_none(
        req.hourly_rate, excel_det.hourly_rate, 0.0,
    )
    hpd = _first_not_none(
        req.hours_per_day, excel_det.hours_per_day, 0.0,
    )
    guards = int(_first_not_none(req.guards, excel_det.guards, 1))

    total_hours = round(n_days * hpd * guards, 2)
    grand_total = round(total_hours * rate, 2)

    invoice_number_str = f"{analysis.invoice_prefix}{invoice_int}"

    # Apply to Excel copy.
    unresolved_excel: list[str] = []
    if new_excel is not None:
        unresolved_excel = _update_excel(
            new_excel,
            excel_det,
            source_month=analysis.source_month or excel_det.detected_month,
            target_year=req.target_year,
            target_month=req.target_month,
            invoice_date=req.invoice_date,
        )

    # Apply to Word copy.
    unresolved_word: list[str] = []
    if new_word is not None:
        period_string = _format_period(req.target_year, req.target_month)
        invoice_date_string = _format_invoice_date(req.invoice_date)
        unresolved_word = _update_word(
            new_word,
            word_det,
            values={
                "invoice_number": invoice_number_str,
                "invoice_date": invoice_date_string,
                "billing_period": period_string,
                "total_hours": _format_number(total_hours),
                "grand_total": _format_currency(grand_total),
            },
        )

    # Write cache back into the SOURCE folder so next run is fast.
    try:
        if analysis.excel_path and analysis.word_path:
            save_cache(source, DetectionCache(excel=excel_det, word=word_det))
    except Exception:  # noqa: BLE001
        pass

    return GenerateResult(
        copied_folder=target_folder,
        excel_path=new_excel,
        word_path=new_word,
        invoice_number=invoice_number_str,
        invoice_number_int=invoice_int,
        total_hours=total_hours,
        grand_total=grand_total,
        analysis=analysis,
        excel_detection=excel_det,
        word_detection=word_det,
        unresolved_excel=unresolved_excel,
        unresolved_word=unresolved_word,
    )


# ---------------------------------------------------------- helpers --------

def _swap_month_in_name(name: str, token: str, target_month: int) -> str:
    import re as _re
    target_name = MONTH_NAMES[target_month - 1]
    if token:
        return _re.sub(_re.escape(token), target_name, name, flags=_re.IGNORECASE)
    return target_name  # folder didn't contain a month name; use raw target name


def _first_not_none(*values: Any) -> Any:
    for v in values:
        if v is not None and v != 0:
            return v
    for v in values:
        if v is not None:
            return v
    return None


def _update_excel(
    path: Path,
    det: ExcelDetection,
    *,
    source_month: int | None,
    target_year: int,
    target_month: int,
    invoice_date: date,
) -> list[str]:
    unresolved: list[str] = []
    wb = load_workbook(filename=str(path))
    try:
        ws = wb[det.sheet] if det.sheet and det.sheet in wb.sheetnames else wb.active

        if det.month_number_cell:
            ws[det.month_number_cell] = target_month
        else:
            unresolved.append("month_number_cell")

        if det.invoice_date_cell:
            ws[det.invoice_date_cell] = invoice_date
            ws[det.invoice_date_cell].number_format = "dddd, mmmm d, yyyy"
        else:
            unresolved.append("invoice_date_cell")

        if det.dates_column_start:
            col_letter, start_row = coordinate_from_string(det.dates_column_start)
            col_idx = _col_to_index(col_letter)
            target_days = days_in_month(target_year, target_month)
            # Write new month dates.
            for i in range(target_days):
                d = date(target_year, target_month, i + 1)
                cell = ws.cell(row=start_row + i, column=col_idx, value=d)
                cell.number_format = "dddd, mmmm d, yyyy"
            # Clear any trailing rows from the source month if it was longer.
            source_days_guess = det.dates_row_end and (det.dates_row_end - start_row + 1)
            source_days = source_days_guess or (
                days_in_month(target_year, source_month) if source_month else 31
            )
            if source_days > target_days:
                for i in range(target_days, source_days):
                    ws.cell(row=start_row + i, column=col_idx).value = None
        else:
            unresolved.append("dates_column_start")

        wb.save(str(path))
    finally:
        wb.close()
    return unresolved


def _update_word(
    path: Path,
    det: WordDetection,
    *,
    values: dict[str, str],
) -> list[str]:
    unresolved: list[str] = []
    doc = Document(str(path))

    for field_name in ("invoice_number", "invoice_date", "billing_period",
                       "total_hours", "grand_total"):
        loc: WordLoc | None = getattr(det, field_name)
        value = values.get(field_name, "")
        if loc is None:
            unresolved.append(field_name)
            continue
        try:
            cell = doc.tables[loc.table_index].rows[loc.row].cells[loc.col]
        except IndexError:
            unresolved.append(field_name)
            continue
        _set_cell_text(cell, value, loc.paragraph_index)

    doc.save(str(path))
    return unresolved


def _set_cell_text(cell, text: str, paragraph_index: int = 0) -> None:
    if paragraph_index >= len(cell.paragraphs):
        paragraph_index = 0
    paragraph = cell.paragraphs[paragraph_index]
    if paragraph.runs:
        paragraph.runs[0].text = text
        for extra in paragraph.runs[1:]:
            extra.text = ""
    else:
        paragraph.add_run(text)


def _format_period(year: int, month: int) -> str:
    fmt = "%#m/%#d/%Y" if platform.system() == "Windows" else "%-m/%-d/%Y"
    start = date(year, month, 1).strftime(fmt)
    end = date(year, month, days_in_month(year, month)).strftime(fmt)
    return f"{start} – {end}"


def _format_invoice_date(d: date) -> str:
    return d.strftime("%B %d, %Y").replace(" 0", " ")


def _format_number(value: float) -> str:
    if float(value).is_integer():
        return str(int(value))
    return f"{value:.2f}"


def _format_currency(value: float) -> str:
    if float(value).is_integer():
        return f"${int(value):,}"
    return f"${value:,.2f}"


def _col_to_index(letters: str) -> int:
    acc = 0
    for c in letters.upper():
        acc = acc * 26 + (ord(c) - ord("A") + 1)
    return acc
