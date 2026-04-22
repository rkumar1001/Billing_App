"""Analyze an invoice folder to discover its layout without user input.

The PDF workflow hands the app a single folder (e.g. `.../Bella Vista/March/`)
that contains exactly one Excel breakdown and one Word invoice. From file names
and folder name alone we can infer the source month/year and the invoice number
we're about to bump.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

from .calendar_util import MONTH_NAMES

# Matches "...Bella16.docx" → ("Bella", "16"). Tolerates spaces and separators
# so "Invoice Bella 16.docx" also works.
_INVOICE_NUM_RE = re.compile(
    r"(?P<prefix>[A-Za-z][A-Za-z _-]*?)\s*(?P<number>\d+)\s*$"
)

_MONTH_LOOKUP = {name.lower(): i + 1 for i, name in enumerate(MONTH_NAMES)}
_MONTH_LOOKUP.update({name[:3].lower(): i + 1 for i, name in enumerate(MONTH_NAMES)})


@dataclass
class FolderAnalysis:
    folder: Path
    excel_path: Path | None = None
    word_path: Path | None = None
    extra_files: list[Path] = field(default_factory=list)
    source_month: int | None = None
    source_year: int | None = None
    invoice_prefix: str = ""
    invoice_number: int | None = None
    folder_month_token: str = ""  # e.g. "March" — token we'll swap in filenames
    warnings: list[str] = field(default_factory=list)

    @property
    def next_invoice_number(self) -> int:
        return (self.invoice_number or 0) + 1

    def missing(self) -> list[str]:
        out: list[str] = []
        if not self.excel_path:
            out.append("Excel file (.xlsx)")
        if not self.word_path:
            out.append("Word file (.docx)")
        if self.invoice_number is None:
            out.append("invoice number (parsed from Word filename)")
        return out


def analyze(folder: Path) -> FolderAnalysis:
    folder = Path(folder)
    result = FolderAnalysis(folder=folder)
    if not folder.is_dir():
        result.warnings.append(f"{folder} is not a folder")
        return result

    # Only look at top-level files. Subfolders (Attachments/, Documents/) ride
    # along during copy but never get edited.
    for item in folder.iterdir():
        if not item.is_file():
            continue
        lower = item.name.lower()
        if lower.endswith(".xlsx") or lower.endswith(".xlsm"):
            if result.excel_path is None:
                result.excel_path = item
            else:
                result.extra_files.append(item)
                result.warnings.append(
                    f"multiple Excel files found; using {result.excel_path.name}"
                )
        elif lower.endswith(".docx"):
            if result.word_path is None:
                result.word_path = item
            else:
                result.extra_files.append(item)
                result.warnings.append(
                    f"multiple Word files found; using {result.word_path.name}"
                )
        else:
            result.extra_files.append(item)

    # Detect invoice number / prefix from the Word filename.
    if result.word_path is not None:
        stem = result.word_path.stem  # strip .docx
        # Drop a leading "Invoice " if present so the regex latches onto the
        # client-specific prefix+number.
        stripped = re.sub(r"^(?:invoice[\s_-]*)", "", stem, flags=re.IGNORECASE)
        m = _INVOICE_NUM_RE.search(stripped)
        if m:
            prefix_raw = m.group("prefix").strip(" _-")
            result.invoice_prefix = prefix_raw
            result.invoice_number = int(m.group("number"))
        else:
            result.warnings.append(
                f"could not parse invoice number from '{result.word_path.name}'"
            )

    # Detect source month from folder name, then (as a fallback) from xlsx name.
    month_token, month_num = _find_month_token(folder.name)
    if month_num is None and result.excel_path is not None:
        month_token, month_num = _find_month_token(result.excel_path.stem)
    if month_num is not None:
        result.source_month = month_num
        result.folder_month_token = month_token

    # Year: best guess is whatever year appears in the folder name or xlsx
    # name; we'll refine it later from the dates column in the Excel itself.
    for candidate in (folder.name, result.excel_path.stem if result.excel_path else ""):
        year_match = re.search(r"\b(20\d{2})\b", candidate)
        if year_match:
            result.source_year = int(year_match.group(1))
            break

    return result


def _find_month_token(text: str) -> tuple[str, int | None]:
    """Return (original_casing_token, month_number) found in text, or ('', None)."""
    # Prefer full names over 3-letter abbreviations when both overlap.
    for full in MONTH_NAMES:
        m = re.search(rf"\b({re.escape(full)})\b", text, re.IGNORECASE)
        if m:
            return m.group(1), _MONTH_LOOKUP[full.lower()]
    for full in MONTH_NAMES:
        short = full[:3]
        m = re.search(rf"\b({re.escape(short)})\b", text, re.IGNORECASE)
        if m:
            return m.group(1), _MONTH_LOOKUP[short.lower()]
    return "", None


def rename_for_target(
    analysis: FolderAnalysis,
    copied_folder: Path,
    target_month: int,
    target_year: int,
    new_invoice_number: int,
) -> tuple[Path | None, Path | None]:
    """After copying, rename the Excel + Word inside `copied_folder`.

    Returns the new (excel_path, word_path) paths. Files that couldn't be
    located in the copy are returned as None.
    """
    new_excel: Path | None = None
    new_word: Path | None = None

    if analysis.excel_path is not None:
        src_copy = copied_folder / analysis.excel_path.name
        if src_copy.exists():
            target_month_name = MONTH_NAMES[target_month - 1]
            new_name = src_copy.name
            token = analysis.folder_month_token
            if token:
                new_name = re.sub(
                    re.escape(token), target_month_name, new_name, flags=re.IGNORECASE
                )
            new_excel = src_copy.with_name(new_name)
            if new_excel != src_copy:
                src_copy.rename(new_excel)
            else:
                new_excel = src_copy

    if analysis.word_path is not None:
        src_copy = copied_folder / analysis.word_path.name
        if src_copy.exists():
            old_stem = analysis.word_path.stem
            if analysis.invoice_number is not None:
                old_suffix = f"{analysis.invoice_prefix}{analysis.invoice_number}"
                new_suffix = f"{analysis.invoice_prefix}{new_invoice_number}"
                new_stem = old_stem.replace(old_suffix, new_suffix)
                if new_stem == old_stem:
                    new_stem = re.sub(
                        rf"{re.escape(analysis.invoice_prefix)}\s*\d+",
                        new_suffix,
                        old_stem,
                    )
                new_word = src_copy.with_name(new_stem + src_copy.suffix)
            else:
                new_word = src_copy
            if new_word != src_copy:
                src_copy.rename(new_word)

    return new_excel, new_word
