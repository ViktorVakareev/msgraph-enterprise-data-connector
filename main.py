"""
Step 4 — end-to-end pipeline: fetch (or fall back to bundled sample data)
-> validate/standardize -> derive a merge-ready month key -> merge ->
summarize -> visualize, with logging and timing throughout.

Run: python main.py

If AZURE_CLIENT_ID / AZURE_TENANT_ID / AZURE_CLIENT_SECRET, plus
ONEDRIVE_USER and SHAREPOINT_SITE_ID, are all set in .env, this fetches
real data via Microsoft Graph. Otherwise (and this environment has no real
Azure tenant to test against) it falls back to bundled sample data — the
same "use a sample dataset if the real source isn't available" allowance
used in the earlier Research Plugin Development activity's integration
test, so the full pipeline is always demonstrable end to end.
"""

from __future__ import annotations

import logging
import time

import pandas as pd

from config import load_graph_config
from data_pipeline import align_and_merge, derive_month_key, summarize, validate_and_standardize
from graph_auth import GraphAuthenticator, GraphAuthError
from graph_client import GraphClient, GraphRequestError
from onedrive_fetcher import OneDriveFetcher
from sharepoint_fetcher import SharePointFetcher
from visualize import plot_monthly_sales_by_region

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

# Deliberately messy, mirroring the real-world problems this activity is
# about: mixed date formats, a currency-formatted value, one genuinely bad
# row in each column.
SAMPLE_TRANSACTIONS = pd.DataFrame({
    "TransactionID": [1, 2, 3, 4, 5, 6],
    "Region": ["West", "East", "West", "East", "West", "East"],
    "TransactionDate": ["01/15/2026", "2026-01-20", "02/03/2026", "2026-02-10", "03/01/2026", "not-a-date"],
    "Sales": ["$1,200.50", "980", "1,430.25", "1,050", "1,600", "N/A"],
})

SAMPLE_REGIONAL_PERFORMANCE = pd.DataFrame({
    "RegionName": ["West", "East", "West", "East", "West", "East"],
    "Month": ["2026-01", "2026-01", "2026-02", "2026-02", "2026-03", "2026-03"],
    "TargetSales": [1000, 900, 1400, 1000, 1500, 1100],
    "Status": ["On Track", "Behind", "On Track", "Behind", "On Track", "On Track"],
})


def load_sources() -> tuple[pd.DataFrame, pd.DataFrame, bool]:
    """Tries real Graph credentials + resource ids first; falls back to
    bundled sample data on any missing config or fetch failure. Returns
    (excel_df, sharepoint_df, used_live_data)."""
    try:
        config = load_graph_config()
    except RuntimeError as exc:
        logger.warning("Graph credentials not configured (%s) — using bundled sample data.", exc)
        return SAMPLE_TRANSACTIONS.copy(), SAMPLE_REGIONAL_PERFORMANCE.copy(), False

    if not config.onedrive_user or not config.sharepoint_site_id:
        logger.warning(
            "ONEDRIVE_USER and/or SHAREPOINT_SITE_ID not set in .env — using bundled sample data."
        )
        return SAMPLE_TRANSACTIONS.copy(), SAMPLE_REGIONAL_PERFORMANCE.copy(), False

    try:
        authenticator = GraphAuthenticator(config)
        client = GraphClient(authenticator)

        onedrive = OneDriveFetcher(client, user=config.onedrive_user)
        excel_df = onedrive.fetch_excel_as_dataframe("transactions.xlsx", use_cache=True)

        sharepoint = SharePointFetcher(client, site_id=config.sharepoint_site_id)
        sharepoint_df = sharepoint.fetch_list_as_dataframe("RegionalPerformance", use_cache=True)
    except (GraphAuthError, GraphRequestError, FileNotFoundError, ValueError) as exc:
        logger.warning("Live Graph fetch failed (%s) — falling back to bundled sample data.", exc)
        return SAMPLE_TRANSACTIONS.copy(), SAMPLE_REGIONAL_PERFORMANCE.copy(), False

    return excel_df, sharepoint_df, True


def run_pipeline(output_chart_path: str = "monthly_sales_by_region.png") -> dict:
    overall_start = time.perf_counter()

    fetch_start = time.perf_counter()
    excel_df, sharepoint_df, used_live_data = load_sources()
    fetch_elapsed = time.perf_counter() - fetch_start
    logger.info(
        "Data source: %s (fetch took %.3fs)",
        "live Microsoft Graph" if used_live_data else "bundled sample data",
        fetch_elapsed,
    )

    validate_start = time.perf_counter()
    excel_clean, excel_report = validate_and_standardize(
        excel_df, date_columns=["TransactionDate"], numeric_columns=["Sales"]
    )
    excel_clean = derive_month_key(excel_clean, "TransactionDate", key_column="Month")
    validate_elapsed = time.perf_counter() - validate_start

    if excel_report.has_issues():
        logger.warning(
            "Transaction data had validation issues — unparseable_dates=%s unparseable_numbers=%s "
            "(rows kept, flagged rather than dropped or zero-filled)",
            excel_report.unparseable_dates, excel_report.unparseable_numbers,
        )
    logger.info("Validation + standardization took %.3fs", validate_elapsed)

    rename_right = {"RegionName": "Region"} if "RegionName" in sharepoint_df.columns else None
    merged = align_and_merge(excel_clean, sharepoint_df, on=["Region", "Month"], rename_right=rename_right)

    summary = summarize(merged)

    chart_path = plot_monthly_sales_by_region(merged, output_chart_path)
    logger.info("Chart written to %s", chart_path)

    overall_elapsed = time.perf_counter() - overall_start
    logger.info("Full pipeline completed in %.3fs", overall_elapsed)

    # Documented bottlenecks, per the activity's own request ("Document at
    # least two bottlenecks in comments or a short note"):
    #
    # 1. Graph API fetch calls are the largest fixed cost per pipeline run
    #    — each is a full network round trip, and OneDriveFetcher's file
    #    lookup plus SharePointFetcher's list-then-items sequence are each
    #    at least two serial calls. get_all_pages_cached (used above via
    #    use_cache=True) removes the cost on any *repeated* run within the
    #    same process, but the first run's network cost is unavoidable.
    # 2. get_all_pages() fetches pages serially, one full round trip at a
    #    time — for a very large SharePoint list this dominates runtime.
    #    Not fixed here (out of this activity's scope); a production
    #    version would request a larger page size via $top or fetch pages
    #    concurrently.

    return {
        "merged": merged,
        "summary": summary,
        "chart_path": chart_path,
        "used_live_data": used_live_data,
        "validation_report": excel_report,
        "timings": {
            "fetch_seconds": fetch_elapsed,
            "validate_seconds": validate_elapsed,
            "total_seconds": overall_elapsed,
        },
    }


if __name__ == "__main__":
    result = run_pipeline()

    print(f"\nData source: {'live Microsoft Graph' if result['used_live_data'] else 'bundled sample data'}")
    print(f"Merged rows: {len(result['merged'])}")
    print("\nSummary:")
    for key, value in result["summary"].items():
        print(f"  {key}: {value}")
    print(f"\nChart: {result['chart_path']}")
    print(f"\nTimings: {result['timings']}")
