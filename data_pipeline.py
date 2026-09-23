"""
Step 3 — validation, standardization, and merging of the Excel (OneDrive)
transaction data and the SharePoint regional-performance data, following
the activity's own sample-code sequence: profile missingness, standardize
dates/types, verify + align merge keys, merge with explicit parameters,
then summarize.

Design choice, matching this activity's own Question 2 answer ("Validate
and decide according to your analysis needs"): unparseable dates or numbers
are never silently dropped or zero-filled here. Every standardization
function returns a validation report alongside the cleaned DataFrame, so
the caller sees exactly what couldn't be parsed and decides what to do
about it — dropping, flagging for manual review, etc. — rather than the
pipeline making that call invisibly.
"""

from __future__ import annotations

import logging
import time
from dataclasses import dataclass, field

import pandas as pd

logger = logging.getLogger(__name__)

ISO_DATE_FORMAT = "%Y-%m-%d"


@dataclass
class ValidationReport:
    """What was found while standardizing one DataFrame — never raised as
    an error by itself, so the caller can inspect it and decide."""

    missing_value_counts: dict[str, int] = field(default_factory=dict)
    unparseable_dates: dict[str, list[int]] = field(default_factory=dict)  # column -> row indices
    unparseable_numbers: dict[str, list[int]] = field(default_factory=dict)

    def has_issues(self) -> bool:
        return bool(
            any(count > 0 for count in self.missing_value_counts.values())
            or self.unparseable_dates
            or self.unparseable_numbers
        )


def profile_missing_values(df: pd.DataFrame) -> dict[str, int]:
    """`df.isnull().sum()` as a plain dict — the activity's own suggested
    first step before any transformation."""
    return df.isnull().sum().to_dict()


def standardize_dates(df: pd.DataFrame, columns: list[str]) -> tuple[pd.DataFrame, dict[str, list[int]]]:
    """Converts each column to ISO format (YYYY-MM-DD) strings, accepting
    mixed input formats (pandas' date parser handles "01/15/2026",
    "2026-01-15", etc. within the same column). Rows that fail to parse
    become None in the output — never a fabricated date — and their
    original row indices are returned in the report so nothing is silently
    lost.
    """
    df = df.copy()
    unparseable: dict[str, list[int]] = {}

    for column in columns:
        if column not in df.columns:
            continue
        # format="mixed": confirmed directly that pandas.to_datetime otherwise infers
        # ONE format from the column's first value and silently fails every row that
        # doesn't match it — e.g. a column starting with "01/15/2026" would turn a
        # perfectly valid "2026-02-01" into NaT too. That's the exact "inconsistent
        # date formats" failure mode this activity is about, so format="mixed" (parse
        # each value's format independently) is not optional here.
        parsed = pd.to_datetime(df[column], errors="coerce", format="mixed")
        failed_mask = parsed.isna() & df[column].notna()  # genuinely unparseable, not originally missing
        if failed_mask.any():
            unparseable[column] = df.index[failed_mask].tolist()

        df[column] = parsed.dt.strftime(ISO_DATE_FORMAT)
        df.loc[failed_mask, column] = None

    return df, unparseable


def standardize_numeric(df: pd.DataFrame, columns: list[str]) -> tuple[pd.DataFrame, dict[str, list[int]]]:
    """Converts each column to float, stripping common non-numeric
    formatting ("$", ",", surrounding whitespace) before conversion so
    "$1,200.50" and 1200.50 both land as the same float. Genuinely
    unparseable values become NaN and are reported, not guessed at."""
    df = df.copy()
    unparseable: dict[str, list[int]] = {}

    for column in columns:
        if column not in df.columns:
            continue

        original = df[column]
        cleaned = (
            original.astype(str)
            .str.replace(r"[\$,]", "", regex=True)
            .str.strip()
        )
        converted = pd.to_numeric(cleaned, errors="coerce")
        failed_mask = converted.isna() & original.notna()
        if failed_mask.any():
            unparseable[column] = df.index[failed_mask].tolist()

        df[column] = converted

    return df, unparseable


def validate_and_standardize(
    df: pd.DataFrame,
    date_columns: list[str] | None = None,
    numeric_columns: list[str] | None = None,
) -> tuple[pd.DataFrame, ValidationReport]:
    report = ValidationReport(missing_value_counts=profile_missing_values(df))

    result = df
    if date_columns:
        result, unparseable_dates = standardize_dates(result, date_columns)
        report.unparseable_dates = unparseable_dates
    if numeric_columns:
        result, unparseable_numbers = standardize_numeric(result, numeric_columns)
        report.unparseable_numbers = unparseable_numbers

    return result, report


def derive_month_key(df: pd.DataFrame, date_column: str, key_column: str = "Month") -> pd.DataFrame:
    """Derives a YYYY-MM month key from a standardized (ISO YYYY-MM-DD)
    date column — needed to merge per-transaction data (day-level, from
    Excel) against per-region-per-month summary data (month-level, from
    SharePoint). Without this, merging on the raw transaction date would
    never match anything: a day-level value and a month-level value are two
    different granularities of the same concept, and pandas.merge compares
    values, not concepts. Rows with an unparseable/missing date keep a
    missing month key, so they correctly fail to match rather than being
    grouped under a fabricated month.
    """
    df = df.copy()
    df[key_column] = df[date_column].astype(str).str.slice(0, 7)
    df.loc[df[date_column].isna(), key_column] = None
    return df


def align_and_merge(
    left: pd.DataFrame,
    right: pd.DataFrame,
    on: list[str],
    rename_right: dict[str, str] | None = None,
    how: str = "inner",
) -> pd.DataFrame:
    """Verifies + aligns the merge keys (existence, matching names, matching
    dtypes) before calling pandas.merge, exactly following the activity's
    own sample-code sequence, rather than merging blind and hoping the keys
    line up. Logs the merge's timing so slow merges surface in normal
    operation, not just when someone goes looking for them (Question 4's
    performance-monitoring theme).
    """
    if rename_right:
        right = right.rename(columns=rename_right)

    missing_left = [key for key in on if key not in left.columns]
    missing_right = [key for key in on if key not in right.columns]
    if missing_left or missing_right:
        raise ValueError(
            f"Merge key(s) missing — left missing {missing_left}, right missing {missing_right}. "
            "Rename columns via `rename_right` or fix the source data before merging."
        )

    for key in on:
        if left[key].dtype != right[key].dtype:
            logger.info(
                "Casting merge key %r to align dtypes (%s -> %s) before merging",
                key, right[key].dtype, left[key].dtype,
            )
            right[key] = right[key].astype(left[key].dtype)

    start = time.perf_counter()
    merged = pd.merge(left, right, on=on, how=how, indicator=True)
    elapsed = time.perf_counter() - start
    logger.info("Merge on %s completed in %.4f seconds (%d rows)", on, elapsed, len(merged))

    unmatched_left = int((merged["_merge"] == "left_only").sum())
    unmatched_right = int((merged["_merge"] == "right_only").sum())
    if unmatched_left or unmatched_right:
        logger.warning(
            "Merge produced %d left-only and %d right-only rows — check for missing counterparts.",
            unmatched_left, unmatched_right,
        )

    return merged.drop(columns=["_merge"])


def summarize(merged: pd.DataFrame, region_column: str = "Region", month_column: str = "Month",
              sales_column: str = "Sales", target_column: str = "TargetSales") -> dict:
    """Summary analytics the activity asks for: totals by region,
    month-over-month totals, and actual-vs-target variance where a target
    column is present."""
    summary: dict = {}

    if sales_column in merged.columns and region_column in merged.columns:
        summary["total_sales_by_region"] = merged.groupby(region_column)[sales_column].sum().to_dict()

    if sales_column in merged.columns and month_column in merged.columns:
        summary["total_sales_by_month"] = merged.groupby(month_column)[sales_column].sum().to_dict()

    if sales_column in merged.columns and target_column in merged.columns:
        merged = merged.copy()
        merged["sales_vs_target"] = merged[sales_column] - merged[target_column]
        summary["sales_vs_target_by_region"] = (
            merged.groupby(region_column)["sales_vs_target"].sum().to_dict()
            if region_column in merged.columns
            else None
        )

    return summary
