"""Fiscal periods (Apr-Mar) and their comparison windows (LY / previous period)."""
from __future__ import annotations
from dataclasses import dataclass
import calendar
import pandas as pd

MONTHS = ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"]
KINDS = ["Month", "Quarter", "Half-year", "YTD (Apr to month)", "Full year", "Custom range"]


def fy_label(fy: int, short=False) -> str:
    return f"{(fy - 1) % 100:02d}-{fy % 100:02d}" if short else f"FY {fy - 1}-{fy % 100:02d}"


def mi_of(fy: int, fm: int) -> int:
    """absolute month index of fiscal month fm (Apr=1) in fiscal year fy."""
    return (fy - 1) * 12 + 3 + (fm - 1)


def mi_name(mi: int, with_year=True) -> str:
    y, m = divmod(mi, 12)
    n = calendar.month_abbr[m + 1]
    return f"{n}'{y % 100:02d}" if with_year else n


def month_start(mi: int) -> pd.Timestamp:
    y, m = divmod(mi, 12)
    return pd.Timestamp(year=y, month=m + 1, day=1)


def month_end(mi: int) -> pd.Timestamp:
    return month_start(mi) + pd.offsets.MonthEnd(0)


@dataclass(frozen=True)
class Period:
    kind: str
    fy: int
    fm_start: int
    fm_end: int

    @property
    def mi_start(self): return mi_of(self.fy, self.fm_start)
    @property
    def mi_end(self): return mi_of(self.fy, self.fm_end)
    @property
    def n_months(self): return self.fm_end - self.fm_start + 1
    @property
    def fms(self): return list(range(self.fm_start, self.fm_end + 1))

    def label(self, shift=0) -> str:
        a, b = self.mi_start + shift, self.mi_end + shift
        if self.kind == "Full year" and shift % 12 == 0:
            return fy_label(self.fy + shift // 12)
        if a == b: return mi_name(a)
        return f"{mi_name(a, False)}-{mi_name(b)}"

    def title(self) -> str:
        if self.kind == "Quarter": return f"Q{(self.fm_start - 1) // 3 + 1} {fy_label(self.fy)} ({self.label()})"
        if self.kind == "Half-year": return f"H{(self.fm_start - 1) // 6 + 1} {fy_label(self.fy)} ({self.label()})"
        return f"{self.label()}  ·  {fy_label(self.fy)}"


def make_period(kind: str, fy: int, fm: int = 1, q: int = 1, h: int = 1, fm_from: int = 1, fm_to: int = 12) -> Period:
    if kind == "Month": return Period(kind, fy, fm, fm)
    if kind == "Quarter": return Period(kind, fy, 3 * (q - 1) + 1, 3 * q)
    if kind == "Half-year": return Period(kind, fy, 6 * (h - 1) + 1, 6 * h)
    if kind.startswith("YTD"): return Period(kind, fy, 1, fm)
    if kind == "Full year": return Period(kind, fy, 1, 12)
    return Period(kind, fy, min(fm_from, fm_to), max(fm_from, fm_to))


@dataclass(frozen=True)
class Window:
    start: pd.Timestamp
    end: pd.Timestamp
    label: str
    partial: bool = False   # end was cut because data stops mid-month (like-for-like cut applied)


def windows(p: Period, last_date: pd.Timestamp | None) -> dict[str, Window]:
    """cy / ly / prev windows. If data ends mid-month inside the period, LY & prev are cut at the same
    day so growth is like-for-like."""
    start, end = month_start(p.mi_start), month_end(p.mi_end)
    partial = False
    if last_date is not None and pd.notna(last_date) and last_date < end:
        end, partial = max(last_date, start), True
    out = {"cy": Window(start, end, p.label(), partial)}
    for key, shift in (("ly", -12), ("prev", -p.n_months)):
        s = month_start(p.mi_start + shift)
        e = end - pd.DateOffset(months=-shift) if partial else month_end(p.mi_end + shift)
        out[key] = Window(s, e, p.label(shift) if key == "ly" else p.label(shift), partial)
    return out
