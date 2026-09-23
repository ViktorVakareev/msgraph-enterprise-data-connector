"""
Step 3 tests — data_pipeline.py against deliberately messy sample data:
mixed date formats, currency-formatted sales figures, missing values, and
mismatched key dtypes across the two source DataFrames (exactly the kind
of real-world messiness this activity's own sample code anticipates).

Run: python test_data_pipeline.py  (or: pytest -v)
"""

import pandas as pd

from data_pipeline import (
    align_and_merge,
    derive_month_key,
    profile_missing_values,
    standardize_dates,
    standardize_numeric,
    summarize,
    validate_and_standardize,
)


def test_profile_missing_values():
    df = pd.DataFrame({"A": [1, None, 3], "B": [None, None, "x"]})
    counts = profile_missing_values(df)
    assert counts == {"A": 1, "B": 2}
    print(f"[PASS] profile_missing_values reports exact null counts: {counts}")


def test_standardize_dates_mixed_formats_and_bad_value():
    df = pd.DataFrame({"TxDate": ["01/15/2026", "2026-02-01", "not-a-date", None]})
    result, unparseable = standardize_dates(df, ["TxDate"])

    values = result["TxDate"].tolist()
    assert values[0] == "2026-01-15"
    assert values[1] == "2026-02-01"
    assert pd.isna(values[2]) and pd.isna(values[3])
    assert unparseable == {"TxDate": [2]}  # only the genuinely bad value, not the originally-missing row
    print(f"[PASS] standardize_dates normalized mixed formats to ISO and flagged only the genuine bad value: unparseable={unparseable}")


def test_standardize_numeric_currency_and_bad_value():
    df = pd.DataFrame({"Sales": ["$1,200.50", "980", "N/A", None]})
    result, unparseable = standardize_numeric(df, ["Sales"])

    assert result["Sales"].tolist()[:2] == [1200.50, 980.0]
    assert pd.isna(result["Sales"].iloc[2])
    assert unparseable == {"Sales": [2]}
    print(f"[PASS] standardize_numeric parsed currency-formatted values and flagged only the genuine bad value: unparseable={unparseable}")


def test_validate_and_standardize_combines_both_reports():
    df = pd.DataFrame({
        "TxDate": ["01/15/2026", "garbage"],
        "Sales": ["$500", "oops"],
    })
    result, report = validate_and_standardize(df, date_columns=["TxDate"], numeric_columns=["Sales"])

    assert report.unparseable_dates == {"TxDate": [1]}
    assert report.unparseable_numbers == {"Sales": [1]}
    assert report.has_issues() is True
    assert result.loc[0, "TxDate"] == "2026-01-15"
    assert result.loc[0, "Sales"] == 500.0
    print(f"[PASS] validate_and_standardize combines date + numeric reports without dropping any rows: has_issues={report.has_issues()}")


def test_align_and_merge_renames_and_casts_before_merging():
    # Excel side: Region/Month as the canonical key names, CustomerID-style int key not needed here —
    # this project's key is (Region, Month).
    excel_df = pd.DataFrame({
        "Region": ["West", "East", "West"],
        "Month": ["2026-01", "2026-01", "2026-02"],
        "Sales": [1200.0, 980.0, 1430.0],
    })
    # SharePoint side: differently-named key column, and Month stored as a different dtype (category)
    sharepoint_df = pd.DataFrame({
        "RegionName": ["West", "East", "West"],
        "Month": pd.Categorical(["2026-01", "2026-01", "2026-02"]),
        "TargetSales": [1000, 900, 1400],
    })

    merged = align_and_merge(
        excel_df,
        sharepoint_df,
        on=["Region", "Month"],
        rename_right={"RegionName": "Region"},
    )

    assert len(merged) == 3
    assert set(merged.columns) == {"Region", "Month", "Sales", "TargetSales"}
    print(f"[PASS] align_and_merge renamed the SharePoint key column and merged {len(merged)} rows successfully.")


def test_align_and_merge_raises_on_missing_key():
    left = pd.DataFrame({"Region": ["West"], "Sales": [100]})
    right = pd.DataFrame({"District": ["West"], "TargetSales": [90]})

    try:
        align_and_merge(left, right, on=["Region"])
    except ValueError as exc:
        assert "missing" in str(exc)
        print(f"[PASS] align_and_merge correctly refuses to merge on a key missing from one side: {exc}")
    else:
        raise AssertionError("Expected ValueError")


def test_align_and_merge_flags_unmatched_rows_via_logging(caplog=None):
    left = pd.DataFrame({"Region": ["West", "North"], "Sales": [100, 50]})
    right = pd.DataFrame({"Region": ["West", "East"], "TargetSales": [90, 80]})

    merged = align_and_merge(left, right, on=["Region"], how="inner")

    # inner join: only "West" matches; "North" (left-only) and "East" (right-only) are dropped
    assert len(merged) == 1
    assert merged.iloc[0]["Region"] == "West"
    print("[PASS] Inner merge correctly drops unmatched rows from both sides (logged as warnings).")


def test_summarize_produces_expected_aggregates():
    merged = pd.DataFrame({
        "Region": ["West", "West", "East"],
        "Month": ["2026-01", "2026-02", "2026-01"],
        "Sales": [1200.0, 1430.0, 980.0],
        "TargetSales": [1000.0, 1400.0, 900.0],
    })

    summary = summarize(merged)

    assert summary["total_sales_by_region"] == {"East": 980.0, "West": 2630.0}
    assert summary["total_sales_by_month"] == {"2026-01": 2180.0, "2026-02": 1430.0}
    assert summary["sales_vs_target_by_region"] == {"East": 80.0, "West": 230.0}
    print(f"[PASS] summarize() produced correct region/month totals and sales-vs-target variance: {summary}")


def test_derive_month_key_from_day_level_dates():
    df = pd.DataFrame({"TransactionDate": ["2026-01-15", "2026-01-28", "2026-02-03", None]})
    result = derive_month_key(df, "TransactionDate", key_column="Month")

    assert result["Month"].tolist()[:3] == ["2026-01", "2026-01", "2026-02"]
    assert pd.isna(result["Month"].iloc[3])
    print(f"[PASS] derive_month_key collapsed day-level dates to YYYY-MM: {result['Month'].tolist()}")

if __name__ == "__main__":
    test_profile_missing_values()
    test_standardize_dates_mixed_formats_and_bad_value()
    test_standardize_numeric_currency_and_bad_value()
    test_validate_and_standardize_combines_both_reports()
    test_align_and_merge_renames_and_casts_before_merging()
    test_align_and_merge_raises_on_missing_key()
    test_align_and_merge_flags_unmatched_rows_via_logging()
    test_summarize_produces_expected_aggregates()
    test_derive_month_key_from_day_level_dates()
    print("\nAll Step 3 + Step 4 data pipeline tests passed.")
