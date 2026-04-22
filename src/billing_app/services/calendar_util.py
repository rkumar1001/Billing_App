from __future__ import annotations

import calendar
from datetime import date
from typing import Iterator

MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]


def days_in_month(year: int, month: int) -> int:
    return calendar.monthrange(year, month)[1]


def days_of_month(year: int, month: int) -> list[date]:
    n = days_in_month(year, month)
    return [date(year, month, d) for d in range(1, n + 1)]


def iter_months(start_year: int, count: int = 24) -> Iterator[tuple[int, int]]:
    y, m = start_year, 1
    for _ in range(count):
        yield y, m
        m += 1
        if m > 12:
            m = 1
            y += 1


def period_strings(year: int, month: int, fmt: str = "%-m/%-d/%Y") -> tuple[str, str]:
    # Use "#" on Windows, "-" elsewhere for non-padded day/month.
    import platform
    if platform.system() == "Windows":
        fmt = fmt.replace("%-", "%#")
    start = date(year, month, 1)
    end = date(year, month, days_in_month(year, month))
    return start.strftime(fmt), end.strftime(fmt)


def month_name(month: int) -> str:
    return MONTH_NAMES[month - 1]
