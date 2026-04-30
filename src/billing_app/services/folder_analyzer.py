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
    excel_is_legacy_xls: bool = False
    # Auxiliary spreadsheets in the same folder that should be carried
    # along and bumped each month, keyed by detected role:
    # "payment_request_form", "g703", etc. Each entry is the source path
    # plus a flag for whether it's a legacy .xls.
    auxiliary_excels: list[tuple[Path, str, bool]] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def next_invoice_number(self) -> int:
        return (self.invoice_number or 0) + 1

    def missing(self) -> list[str]:
        out: list[str] = []
        if not self.excel_path:
            out.append("Excel file (.xlsx or .xls)")
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

    # Collect every spreadsheet first; we'll classify by content afterwards
    # so we always pick the daily-breakdown file as the primary, even when
    # auxiliary forms (G703, payment-request-form) sit in the same folder.
    spreadsheets: list[tuple[Path, bool]] = []  # (path, is_legacy_xls)
    for item in folder.iterdir():
        if not item.is_file():
            continue
        lower = item.name.lower()
        if lower.endswith(".xlsx") or lower.endswith(".xlsm"):
            spreadsheets.append((item, False))
        elif lower.endswith(".xls"):
            spreadsheets.append((item, True))
        elif lower.endswith(".docx"):
            if result.word_path is None:
                result.word_path = item
            else:
                result.extra_files.append(item)
                result.warnings.append(
                    f"multiple Word files found; using {result.word_path.name}"
                )
        elif lower.endswith(".pdf"):
            continue  # ignore exports
        else:
            result.extra_files.append(item)

    # Classify spreadsheets. Legacy .xls files require a real-fidelity
    # conversion before we can read them — to keep `analyze` cheap we read
    # the .xlsx ones directly and treat any single .xls as a likely
    # breakdown unless an .xlsx breakdown is already present.
    from .file_role import classify_workbook
    breakdown: tuple[Path, bool] | None = None
    breakdown_role_known = False
    auxiliaries: list[tuple[Path, str, bool]] = []

    classified: list[tuple[Path, bool, str]] = []
    for path, is_legacy in spreadsheets:
        if is_legacy:
            # Defer classification — converting just to peek at content is
            # expensive. We still try below if we end up needing it.
            classified.append((path, True, "unknown"))
            continue
        role = classify_workbook(path)
        classified.append((path, False, role))

    # First pick: a confidently-classified breakdown (.xlsx).
    for path, is_legacy, role in classified:
        if role == "breakdown" and breakdown is None:
            breakdown = (path, is_legacy)
            breakdown_role_known = True

    # Fallback: if no .xlsx breakdown, promote a single legacy .xls (it's
    # almost always the breakdown — auxiliary forms are .xlsx in practice).
    if breakdown is None:
        legacy_candidates = [c for c in classified if c[1]]
        if len(legacy_candidates) == 1:
            path, is_legacy, _ = legacy_candidates[0]
            breakdown = (path, is_legacy)
        elif legacy_candidates:
            # Multiple .xls — pick the one whose name mentions "breakdown".
            for path, is_legacy, _ in legacy_candidates:
                if "breakdown" in path.name.lower():
                    breakdown = (path, is_legacy)
                    break
            if breakdown is None:
                breakdown = (legacy_candidates[0][0], True)

    if breakdown is not None:
        result.excel_path = breakdown[0]
        result.excel_is_legacy_xls = breakdown[1]

    # Everything else that classified as a known auxiliary role gets
    # carried along.
    for path, is_legacy, role in classified:
        if breakdown is not None and path == breakdown[0]:
            continue
        if role in ("payment_request_form", "g703"):
            auxiliaries.append((path, role, is_legacy))
        else:
            result.extra_files.append(path)
    result.auxiliary_excels = auxiliaries

    # Detect invoice number / prefix from the Word filename first, then
    # fall back to the doc body (templates like Hartford's name files
    # "Invoice HART April.docx" with no number in the filename).
    if result.word_path is not None:
        stem = result.word_path.stem  # strip .docx
        stripped = re.sub(r"^(?:invoice[\s_-]*)", "", stem, flags=re.IGNORECASE)
        m = _INVOICE_NUM_RE.search(stripped)
        if m:
            result.invoice_prefix = m.group("prefix").strip(" _-")
            result.invoice_number = int(m.group("number"))
        else:
            body_match = _read_invoice_number_from_docx(result.word_path)
            if body_match is not None:
                prefix, number = body_match
                result.invoice_prefix = prefix
                result.invoice_number = number
            else:
                result.warnings.append(
                    f"could not parse invoice number from "
                    f"'{result.word_path.name}' or its contents"
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


def _read_invoice_number_from_docx(path: Path) -> tuple[str, int] | None:
    """Best-effort scan of a .docx for an invoice number near a label.

    Looks for a paragraph whose text matches ``Invoice #`` / ``Invoice
    Number`` / ``Invoice``, then takes the next non-empty paragraph as
    the value (e.g. ``HART 23``). Returns ``(prefix, number)`` on success.
    """
    try:
        from docx import Document
        from docx.oxml.ns import qn
    except ImportError:
        return None
    try:
        doc = Document(str(path))
    except Exception:  # noqa: BLE001
        return None

    p_tag = qn("w:p")
    t_tag = qn("w:t")
    paragraphs: list[str] = []
    for p in doc.element.body.iter(p_tag):
        # Only leaf paragraphs (no nested w:p) — same rule as auto_detect.
        if any(c is not p for c in p.iter(p_tag)):
            continue
        text = "".join((t.text or "") for t in p.iter(t_tag)).strip()
        if text:
            paragraphs.append(text)

    label_re = re.compile(r"^\s*invoice(?:\s*#|\s*number|\s*no\.?)?\s*$", re.IGNORECASE)
    for i, text in enumerate(paragraphs):
        if label_re.match(text):
            for j in range(i + 1, min(i + 5, len(paragraphs))):
                cand = paragraphs[j]
                # Skip "Date" / "Invoice Date" labels that often follow.
                if re.match(r"^\s*(?:invoice\s*)?date\s*$", cand, re.IGNORECASE):
                    continue
                m = _INVOICE_NUM_RE.search(cand)
                if m:
                    return m.group("prefix").strip(" _-"), int(m.group("number"))
                break
    return None


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
            new_stem = old_stem
            if analysis.invoice_number is not None:
                old_suffix = f"{analysis.invoice_prefix}{analysis.invoice_number}"
                new_suffix = f"{analysis.invoice_prefix}{new_invoice_number}"
                replaced = old_stem.replace(old_suffix, new_suffix)
                if replaced == old_stem:
                    replaced = re.sub(
                        rf"{re.escape(analysis.invoice_prefix)}\s*\d+",
                        new_suffix,
                        old_stem,
                    )
                new_stem = replaced
            # Always also swap any month token (e.g. "April") in the
            # filename — this is the only rename signal for templates that
            # don't put the invoice number in the filename.
            new_stem = swap_any_month_token(new_stem, target_month)
            new_word = src_copy.with_name(new_stem + src_copy.suffix)
            if new_word != src_copy:
                src_copy.rename(new_word)
            else:
                new_word = src_copy

    return new_excel, new_word


def swap_any_month_token(text: str, target_month: int) -> str:
    """Replace any month name (full or 3-letter abbrev) found in `text`
    with the target month's full name. Used to rename files whose
    filenames carry a stale month (e.g. ``FCC Application for Payment
    February.xlsx`` -> ``FCC Application for Payment May.xlsx``).

    Handles underscore-bounded abbreviations too (``SOV_FEB_..`` ->
    ``SOV_May_..``) since underscores aren't word boundaries to ``\\b``.
    """
    target_full = MONTH_NAMES[target_month - 1]

    def _try(name: str) -> str | None:
        # Match `name` only when bounded by something that ISN'T an
        # alphanumeric (so "May" matches in "Payment May.xlsx" and
        # "FEB" matches in "SOV_FEB_..."). Use a negative-look-behind /
        # negative-look-ahead so we don't consume the surrounding chars.
        m = re.search(
            rf"(?<![A-Za-z0-9]){re.escape(name)}(?![A-Za-z0-9])",
            text,
            re.IGNORECASE,
        )
        if m:
            return text[: m.start()] + target_full + text[m.end():]
        return None

    for full in MONTH_NAMES:
        out = _try(full)
        if out is not None:
            return out
    for full in MONTH_NAMES:
        out = _try(full[:3])
        if out is not None:
            return out
    return text
