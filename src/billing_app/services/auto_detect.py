"""Auto-detect field locations in an Excel breakdown + Word invoice pair.

We never want the end user to manually map cells. Instead we infer positions by
looking for visual cues that match the Bella-Vista-style template (but written
to generalize to any similar-layout invoice):

Excel:
  - month_number:  the first integer 1..12 found in the top-right area of the
    top 5 rows.
  - invoice_date:  a cell with a fill colour that isn't white/transparent in
    the top 5 rows — falling back to any date-typed cell in that area.
  - dates_column_start:  the column whose cells contain a run of consecutive
    dates (day 1..N of the same month) — the top of that run is the start row.
  - hourly_rate / hours_per_day / guards:  read from the first date-row by
    locating the headers in the row just above.

Word:
  - invoice_number: the table cell whose text exactly matches the invoice
    number parsed from the filename (e.g. "Bella16").
  - invoice_date:   a nearby sibling cell that parses as a date.
  - billing_period: cell whose text matches `M/D/YYYY – M/D/YYYY`.
  - total_hours:    in the same row as billing_period, the rightmost numeric
    cell.
  - grand_total:    last cell whose text matches a currency string.
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any

from docx import Document
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


@dataclass
class ExcelDetection:
    sheet: str = ""
    month_number_cell: str | None = None
    invoice_date_cell: str | None = None
    dates_column_start: str | None = None
    dates_row_end: int | None = None  # last row that held a source-month date
    detected_month: int | None = None
    detected_year: int | None = None
    hourly_rate: float | None = None
    hours_per_day: float | None = None
    guards: int | None = None
    warnings: list[str] = field(default_factory=list)

    def missing_fields(self) -> list[str]:
        out = []
        if self.month_number_cell is None:
            out.append("month_number_cell")
        if self.invoice_date_cell is None:
            out.append("invoice_date_cell")
        if self.dates_column_start is None:
            out.append("dates_column_start")
        return out

    def to_dict(self) -> dict[str, Any]:
        return {
            "sheet": self.sheet,
            "month_number_cell": self.month_number_cell,
            "invoice_date_cell": self.invoice_date_cell,
            "dates_column_start": self.dates_column_start,
            "dates_row_end": self.dates_row_end,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "ExcelDetection":
        return cls(
            sheet=data.get("sheet", ""),
            month_number_cell=data.get("month_number_cell"),
            invoice_date_cell=data.get("invoice_date_cell"),
            dates_column_start=data.get("dates_column_start"),
            dates_row_end=data.get("dates_row_end"),
        )


@dataclass
class WordLoc:
    table_index: int
    row: int
    col: int
    paragraph_index: int = 0

    def to_dict(self) -> dict[str, Any]:
        return {
            "table": self.table_index, "row": self.row,
            "col": self.col, "paragraph": self.paragraph_index,
        }

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> "WordLoc":
        return cls(
            table_index=int(d["table"]),
            row=int(d["row"]),
            col=int(d["col"]),
            paragraph_index=int(d.get("paragraph", 0)),
        )


@dataclass
class WordDetection:
    invoice_number: WordLoc | None = None
    invoice_date: WordLoc | None = None
    billing_period: WordLoc | None = None
    total_hours: WordLoc | None = None
    grand_total: WordLoc | None = None
    warnings: list[str] = field(default_factory=list)

    def missing_fields(self) -> list[str]:
        out = []
        for name in ("invoice_number", "invoice_date", "billing_period",
                     "total_hours", "grand_total"):
            if getattr(self, name) is None:
                out.append(name)
        return out

    def to_dict(self) -> dict[str, Any]:
        return {
            name: (getattr(self, name).to_dict() if getattr(self, name) else None)
            for name in ("invoice_number", "invoice_date", "billing_period",
                         "total_hours", "grand_total")
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "WordDetection":
        kwargs: dict[str, WordLoc | None] = {}
        for name in ("invoice_number", "invoice_date", "billing_period",
                     "total_hours", "grand_total"):
            v = data.get(name)
            kwargs[name] = WordLoc.from_dict(v) if v else None
        return cls(**kwargs)


# ---------------------------------------------------------------- Excel ----

def detect_excel(xlsx_path: Path) -> ExcelDetection:
    wb = load_workbook(filename=str(xlsx_path), data_only=True)
    try:
        ws = wb.active
        result = ExcelDetection(sheet=ws.title)

        # Month number: integer 1..12 in top 5 rows, rightmost wins.
        month_cell: str | None = None
        best_col = -1
        for row in range(1, 6):
            for col in range(1, 20):
                val = ws.cell(row=row, column=col).value
                if isinstance(val, int) and 1 <= val <= 12 and col > best_col:
                    # Skip the "BELLA VISTA COMMUNITY"-style text-number mix.
                    left = ws.cell(row=row, column=col - 1).value if col > 1 else None
                    if left in (None, ""):
                        best_col = col
                        month_cell = f"{get_column_letter(col)}{row}"
        result.month_number_cell = month_cell

        # Invoice date: a filled (non-default) cell containing a date, in top 5.
        date_cell: str | None = None
        fallback_date_cell: str | None = None
        for row in range(1, 6):
            for col in range(1, 20):
                cell = ws.cell(row=row, column=col)
                val = cell.value
                is_date = isinstance(val, (date, datetime))
                if not is_date:
                    continue
                fill = cell.fill
                coloured = (
                    fill
                    and fill.fgColor
                    and fill.fgColor.type == "rgb"
                    and fill.fgColor.rgb
                    and fill.fgColor.rgb not in ("00000000", "FFFFFFFF", "FFFFFF", "00FFFFFF")
                )
                addr = f"{get_column_letter(col)}{row}"
                if coloured and date_cell is None:
                    date_cell = addr
                if fallback_date_cell is None:
                    fallback_date_cell = addr
        result.invoice_date_cell = date_cell or fallback_date_cell

        # Dates column: look for a column where consecutive rows hold dates
        # of the same month, starting with day 1.
        best: tuple[str, int, int, int] | None = None  # (col_letter, start_row, month, year)
        for col in range(1, 20):
            run_start = None
            run_month = None
            run_year = None
            run_len = 0
            for row in range(1, 60):
                val = ws.cell(row=row, column=col).value
                d = _as_date(val)
                if d is None:
                    if run_len >= 20:  # good enough
                        break
                    run_start = None
                    run_month = None
                    run_len = 0
                    continue
                if run_start is None:
                    if d.day != 1:
                        continue
                    run_start = row
                    run_month = d.month
                    run_year = d.year
                    run_len = 1
                else:
                    expected_day = run_len + 1
                    if d.month == run_month and d.day == expected_day:
                        run_len += 1
                    else:
                        break
            if run_len >= 20:
                col_letter = get_column_letter(col)
                if best is None or run_len > best[2]:
                    best = (col_letter, run_start, run_len, run_year or 0)
                    result.dates_column_start = f"{col_letter}{run_start}"
                    result.dates_row_end = run_start + run_len - 1
                    result.detected_month = run_month
                    result.detected_year = run_year

        # Read rate, hours per day, guards from the first date row.
        if result.dates_column_start:
            start_cell = result.dates_column_start
            col_letter = re.match(r"([A-Z]+)", start_cell).group(1)
            col_idx = _col_to_index(col_letter)
            start_row = int(start_cell[len(col_letter):])
            # Scan a handful of neighbouring columns on the start row and pick
            # the first small integer (hours) and first currency-ish value (rate).
            for col in range(col_idx + 1, col_idx + 8):
                val = ws.cell(row=start_row, column=col).value
                if isinstance(val, (int, float)) and val > 0:
                    if result.hours_per_day is None and 1 <= val <= 24 and isinstance(val, int):
                        result.hours_per_day = float(val)
                    elif result.guards is None and 1 <= val <= 50 and isinstance(val, int):
                        result.guards = int(val)
                    elif result.hourly_rate is None and val >= 5:
                        result.hourly_rate = float(val)
            if result.hours_per_day is None:
                result.hours_per_day = 0.0
            if result.guards is None:
                result.guards = 1
            if result.hourly_rate is None:
                result.hourly_rate = 0.0

        return result
    finally:
        wb.close()


def _as_date(val: Any) -> date | None:
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    return None


def _col_to_index(letters: str) -> int:
    acc = 0
    for c in letters.upper():
        acc = acc * 26 + (ord(c) - ord("A") + 1)
    return acc


# ---------------------------------------------------------------- Word -----

_DATE_RANGE_RE = re.compile(r"\b\d{1,2}/\d{1,2}/\d{4}\b\s*[–\-—]\s*\b\d{1,2}/\d{1,2}/\d{4}\b")
_CURRENCY_RE = re.compile(r"\$\s*\d[\d,]*(?:\.\d+)?")
_DATE_ONLY_RE = re.compile(
    r"(?:\b\d{1,2}/\d{1,2}/\d{2,4}\b|"
    r"(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|"
    r"Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|"
    r"Dec(?:ember)?)\s+\d{1,2},?\s*\d{4})",
    re.IGNORECASE,
)
_NUMERIC_RE = re.compile(r"^\s*\d+(?:\.\d+)?\s*$")


def detect_word(docx_path: Path, invoice_number_str: str) -> WordDetection:
    doc = Document(str(docx_path))
    result = WordDetection()

    for t_idx, table in enumerate(doc.tables):
        for r_idx, row in enumerate(table.rows):
            for c_idx, cell in enumerate(row.cells):
                text = cell.text.strip()
                if not text:
                    continue
                if (result.invoice_number is None
                        and invoice_number_str
                        and text == invoice_number_str):
                    result.invoice_number = WordLoc(t_idx, r_idx, c_idx)
                    # invoice_date is commonly the cell in the next row, same col.
                    if r_idx + 1 < len(table.rows):
                        below = table.rows[r_idx + 1].cells[c_idx].text.strip()
                        if _DATE_ONLY_RE.search(below):
                            result.invoice_date = WordLoc(t_idx, r_idx + 1, c_idx)
                if result.billing_period is None and _DATE_RANGE_RE.search(text):
                    result.billing_period = WordLoc(t_idx, r_idx, c_idx)
                    # total_hours = rightmost numeric cell on the same row.
                    for scan_col in range(len(row.cells) - 1, c_idx, -1):
                        scan = row.cells[scan_col].text.strip()
                        if _NUMERIC_RE.match(scan):
                            result.total_hours = WordLoc(t_idx, r_idx, scan_col)
                            break
                if _CURRENCY_RE.fullmatch(text) or _CURRENCY_RE.match(text):
                    # Heuristic: currency strings in tables with a "Grand" label
                    # win over other currency values.
                    prev_cell_text = (
                        row.cells[c_idx - 1].text.strip().lower() if c_idx > 0 else ""
                    )
                    label_row_above = ""
                    if r_idx > 0:
                        label_row_above = table.rows[r_idx - 1].cells[c_idx].text.strip().lower()
                    if (
                        "grand" in prev_cell_text
                        or "total" in prev_cell_text
                        or "grand" in label_row_above
                    ):
                        result.grand_total = WordLoc(t_idx, r_idx, c_idx)
                    elif result.grand_total is None:
                        # Fall back to whichever currency we see first.
                        result.grand_total = WordLoc(t_idx, r_idx, c_idx)

    # If invoice_date still unknown, search for any date-like cell in the same
    # table as invoice_number.
    if result.invoice_date is None and result.invoice_number is not None:
        t = doc.tables[result.invoice_number.table_index]
        for r_idx, row in enumerate(t.rows):
            for c_idx, cell in enumerate(row.cells):
                text = cell.text.strip()
                if _DATE_ONLY_RE.search(text) and (r_idx, c_idx) != (
                    result.invoice_number.row, result.invoice_number.col
                ):
                    result.invoice_date = WordLoc(
                        result.invoice_number.table_index, r_idx, c_idx
                    )
                    break
            if result.invoice_date:
                break

    return result


# ---------------------------------------------------------------- cache ----

CACHE_FILENAME = ".billingapp.json"


@dataclass
class DetectionCache:
    excel: ExcelDetection
    word: WordDetection

    def to_dict(self) -> dict[str, Any]:
        return {"excel": self.excel.to_dict(), "word": self.word.to_dict()}

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "DetectionCache":
        return cls(
            excel=ExcelDetection.from_dict(data.get("excel", {})),
            word=WordDetection.from_dict(data.get("word", {})),
        )


def load_cache(folder: Path) -> DetectionCache | None:
    path = folder / CACHE_FILENAME
    if not path.exists():
        return None
    try:
        return DetectionCache.from_dict(json.loads(path.read_text(encoding="utf-8")))
    except (json.JSONDecodeError, OSError, KeyError, ValueError):
        return None


def save_cache(folder: Path, cache: DetectionCache) -> None:
    path = folder / CACHE_FILENAME
    try:
        path.write_text(json.dumps(cache.to_dict(), indent=2), encoding="utf-8")
    except OSError:
        pass
