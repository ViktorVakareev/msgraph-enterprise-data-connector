"""
Step 4 — the summary visualization the activity's own demo calls for:
"Sample bar graph with monthly sales totals."

Uses matplotlib (Agg backend — no display needed, correct for a script
running in a pipeline/container). Static PNG, so this follows the data-viz
skill's form and color guidance (fixed-order categorical palette, thin
recessive gridlines/axes, direct value labels when there aren't too many
bars to stay legible, a legend whenever more than one region is plotted)
but not its interactive-layer requirements (hover/tooltip, dark-mode
toggle, live filters) — those apply to an interactive web chart, and this
is a static image artifact produced by a Python script, the same medium
the activity's own demo video uses.

Palette: the first slots of the validated default categorical order from
the dataviz skill's reference palette (blue, orange, aqua, yellow, magenta,
green, violet, red) — chosen in that fixed order, never cycled arbitrarily,
and never assigned by rank (a given region always gets the same color
across runs, based on first-seen order).
"""

from __future__ import annotations

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
import pandas as pd

CATEGORICAL_PALETTE = [
    "#2a78d6",  # blue
    "#eb6834",  # orange
    "#1baf7a",  # aqua
    "#eda100",  # yellow
    "#e87ba4",  # magenta
    "#008300",  # green
    "#4a3aa7",  # violet
    "#e34948",  # red
]
MAX_DIRECT_SERIES = len(CATEGORICAL_PALETTE)

CHART_SURFACE = "#fcfcfb"
GRIDLINE_COLOR = "#e1e0d9"
AXIS_COLOR = "#c3c2b7"
MUTED_TEXT = "#898781"
PRIMARY_TEXT = "#0b0b0b"
SECONDARY_TEXT = "#52514e"


def plot_monthly_sales_by_region(
    merged: pd.DataFrame,
    output_path: str,
    region_column: str = "Region",
    month_column: str = "Month",
    sales_column: str = "Sales",
) -> str:
    """Grouped bar chart of total sales per region, per month. Regions
    beyond the validated palette's 8 slots fold into "Other" rather than
    cycling colors past what's been checked for CVD-safety."""
    pivot = (
        merged.groupby([month_column, region_column])[sales_column]
        .sum()
        .unstack(region_column)
        .sort_index()
    )
    months = pivot.index.tolist()
    regions = pivot.columns.tolist()

    if len(regions) > MAX_DIRECT_SERIES:
        top_regions = pivot.sum().sort_values(ascending=False).index[: MAX_DIRECT_SERIES - 1]
        overflow = pivot.drop(columns=list(top_regions)).sum(axis=1)
        pivot = pivot[list(top_regions)].copy()
        pivot["Other"] = overflow
        regions = pivot.columns.tolist()

    fig, ax = plt.subplots(figsize=(8, 5), facecolor=CHART_SURFACE)
    ax.set_facecolor(CHART_SURFACE)

    n_regions = max(len(regions), 1)
    n_months = len(months)
    bar_width = 0.8 / n_regions
    x_positions = range(n_months)

    show_direct_labels = n_regions * n_months <= 12  # selective, not on every bar past legibility

    for i, region in enumerate(regions):
        offsets = [xi - 0.4 + bar_width * (i + 0.5) for xi in x_positions]
        values = pivot[region].fillna(0).tolist()
        color = CATEGORICAL_PALETTE[i % len(CATEGORICAL_PALETTE)]
        bars = ax.bar(offsets, values, width=bar_width * 0.9, color=color, label=str(region))

        if show_direct_labels:
            for bar, value in zip(bars, values):
                if value:
                    ax.text(
                        bar.get_x() + bar.get_width() / 2,
                        bar.get_height(),
                        f"{value:,.0f}",
                        ha="center",
                        va="bottom",
                        fontsize=8,
                        color=SECONDARY_TEXT,
                    )

    ax.set_xticks(list(x_positions))
    ax.set_xticklabels([str(m) for m in months], color=PRIMARY_TEXT)
    ax.set_ylabel("Sales", color=SECONDARY_TEXT)
    ax.set_title("Monthly Sales Totals by Region", color=PRIMARY_TEXT, fontsize=13, loc="left", fontweight="bold")

    ax.yaxis.grid(True, color=GRIDLINE_COLOR, linewidth=0.8)
    ax.set_axisbelow(True)
    for side in ("top", "right", "left"):
        ax.spines[side].set_visible(False)
    ax.spines["bottom"].set_color(AXIS_COLOR)
    ax.tick_params(colors=MUTED_TEXT)

    if n_regions >= 2:
        legend = ax.legend(frameon=False, loc="upper left", bbox_to_anchor=(1.0, 1.0))
        for text in legend.get_texts():
            text.set_color(SECONDARY_TEXT)

    fig.tight_layout()
    fig.savefig(output_path, dpi=150, facecolor=CHART_SURFACE)
    plt.close(fig)
    return output_path


if __name__ == "__main__":
    sample = pd.DataFrame({
        "Region": ["West", "East", "West", "East", "North"],
        "Month": ["2026-01", "2026-01", "2026-02", "2026-02", "2026-01"],
        "Sales": [1200.0, 980.0, 1430.0, 1050.0, 640.0],
    })
    path = plot_monthly_sales_by_region(sample, "sample_monthly_sales.png")
    print(f"Sample chart written to {path}")
