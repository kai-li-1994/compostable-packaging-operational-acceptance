#!/usr/bin/env python3
"""
Generate Figure 2 for the compostable-packaging results section.

Figure logic:
    Panel A: Source authority distribution.
    Panel B: Stated organic-waste treatment / collection route.
    Panel C: Source-level acceptance outcome.
    Panel D: Stated rejection/restriction rationale.

Design:
    - Four coordinated horizontal bar panels.
    - No external icons or web downloads are used.
    - Bars are sorted descending by share within every panel.
    - Each panel uses one muted RWTH Aachen-inspired colour.
    - No shared legend is drawn because categories are directly labelled.

Input structure:
    The script assumes the final display CSV files are always located at:

        results/display/

    relative to the folder where this script is run.

Required input files:
    results/display/source_level_category_shares_display.csv
    results/display/rejection_rationale_reporting_shares_display.csv

Outputs:
    figures/figure_2_source_level_operational_rules.png
    figures/figure_2_source_level_operational_rules.pdf
    figures/figure_2_source_level_operational_rules.csv

"""

from __future__ import annotations

from pathlib import Path
import matplotlib
matplotlib.use("Agg")

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


# ============================================================
# 1. Paths and basic settings
# ============================================================

SCRIPT_DIR = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()

DISPLAY_DIR = SCRIPT_DIR / "results" / "display"
OUTPUT_DIR = SCRIPT_DIR / "figures"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

SOURCE_SHARE_PATH = DISPLAY_DIR / "source_level_category_shares_display.csv"
RATIONALE_SHARE_PATH = DISPLAY_DIR / "rejection_rationale_reporting_shares_display.csv"

OUTPUT_PNG = OUTPUT_DIR / "figure_2_source_level_operational_rules.png"
OUTPUT_PDF = OUTPUT_DIR / "figure_2_source_level_operational_rules.pdf"
OUTPUT_DATA = OUTPUT_DIR / "figure_2_source_level_operational_rules.csv"

# Figure settings
DPI = 300
FIG_WIDTH = 16.5
FIG_HEIGHT = 11.8

# Font settings
FONT_FALLBACKS = ["Arial", "DejaVu Sans"]

FONT_SIZE_PANEL_TITLE = 18
FONT_SIZE_AXIS_LABEL = 16
FONT_SIZE_TICK_LABEL = 16
FONT_SIZE_BAR_LABEL = 16
FONT_SIZE_LEGEND = 16
# Layout controls
PANEL_WSPACE = 0.7
PANEL_HSPACE = 0.38
LEFT_MARGIN = 0.105
RIGHT_MARGIN = 0.985
TOP_MARGIN = 0.95
BOTTOM_MARGIN = 0.075

# Bar and grid controls
BAR_HEIGHT = 0.62
GRID_ALPHA = 0.22
GRID_WIDTH = 0.8
SPINE_WIDTH = 0.8

plt.rcParams.update({
    "font.family": "sans-serif",
    "font.sans-serif": FONT_FALLBACKS,
    "axes.titlesize": FONT_SIZE_PANEL_TITLE,
    "axes.labelsize": FONT_SIZE_AXIS_LABEL,
    "xtick.labelsize": FONT_SIZE_TICK_LABEL,
    "ytick.labelsize": FONT_SIZE_TICK_LABEL,
    "legend.fontsize": FONT_SIZE_LEGEND,
    "pdf.fonttype": 42,   # keep text editable in Illustrator/Inkscape
    "ps.fonttype": 42,
})


# ============================================================
# 2. Modular colour parameters
# ============================================================

# Muted RWTH Aachen-inspired colour system, aligned with Figure 3.
# Figure 2 uses one colour per panel. The four panel colours come from the
# same final muted family used in Figure 3:
#   - blue: source authority
#   - green: treatment route
#   - warm beige: acceptance
#   - muted lavender-grey: rejection/restriction rationale
# Neutral greys are retained for text, grid lines, and panel borders.

RWTH_BLUE_MUTED = "#9DBBD5"
RWTH_GREEN_MUTED = "#DCE8C4"
RWTH_WARM_MUTED = "#DECBAE"
RWTH_CONFLICT_MUTED = "#C7BFD3"
RWTH_GREY_MUTED = "#D7D7D7"
RWTH_GREY_LIGHT = "#F1F1F1"

COLOR_TEXT = "#111827"
COLOR_MUTED_TEXT = "#6B7280"
COLOR_GRID = "#CBD5E1"
COLOR_PANEL_BORDER = "#334155"

# One colour per panel.
COLOR_SOURCE_AUTHORITY_PANEL = RWTH_BLUE_MUTED
COLOR_TREATMENT_ROUTE_PANEL = RWTH_GREEN_MUTED
COLOR_ACCEPTANCE_PANEL = RWTH_WARM_MUTED
COLOR_REJECTION_RATIONALE_PANEL = RWTH_CONFLICT_MUTED

PANEL_COLORS = {
    "source_authority": COLOR_SOURCE_AUTHORITY_PANEL,
    "treatment_route": COLOR_TREATMENT_ROUTE_PANEL,
    "acceptance": COLOR_ACCEPTANCE_PANEL,
    "rejection_rationale": COLOR_REJECTION_RATIONALE_PANEL,
}


# Retained only for compatibility with existing plotting functions that expect
# category-colour dictionaries. Because each panel now uses a single colour,
# every category within a panel maps to the same panel colour.
SOURCE_AUTHORITY_COLORS = {
    "Municipal sorting rule": COLOR_SOURCE_AUTHORITY_PANEL,
    "Waste operator or treatment-facility rule": COLOR_SOURCE_AUTHORITY_PANEL,
    "National rule or policy": COLOR_SOURCE_AUTHORITY_PANEL,
    "State, provincial, regional, or canton rule": COLOR_SOURCE_AUTHORITY_PANEL,
    "Intermunicipal or regional waste authority": COLOR_SOURCE_AUTHORITY_PANEL,
    "Supranational framework": COLOR_SOURCE_AUTHORITY_PANEL,
}

TREATMENT_ROUTE_COLORS = {
    "Composting": COLOR_TREATMENT_ROUTE_PANEL,
    "Treatment route not stated": COLOR_TREATMENT_ROUTE_PANEL,
    "AD/biogas + composting": COLOR_TREATMENT_ROUTE_PANEL,
    "AD/biogas only": COLOR_TREATMENT_ROUTE_PANEL,
    "AD / biogas": COLOR_TREATMENT_ROUTE_PANEL,
    "Other organic recovery": COLOR_TREATMENT_ROUTE_PANEL,
    "Other valorisation": COLOR_TREATMENT_ROUTE_PANEL,
}

ACCEPTANCE_COLORS = {
    "Rejected": COLOR_ACCEPTANCE_PANEL,
    "Collection liners only": COLOR_ACCEPTANCE_PANEL,
    "Dedicated/controlled route only": COLOR_ACCEPTANCE_PANEL,
    "Listed/dedicated route only": COLOR_ACCEPTANCE_PANEL,
    "Accepted broadly": COLOR_ACCEPTANCE_PANEL,
    "Accepted": COLOR_ACCEPTANCE_PANEL,
    "Local decision required": COLOR_ACCEPTANCE_PANEL,
    "No explicit compostable-packaging rule": COLOR_ACCEPTANCE_PANEL,
    "No clear organics route": COLOR_ACCEPTANCE_PANEL,
    "No clear rule/route": COLOR_ACCEPTANCE_PANEL,
}

RATIONALE_COLORS = {
    "Compost/digestate quality or contamination concern": COLOR_REJECTION_RATIONALE_PANEL,
    "No compatible organic-treatment route": COLOR_REJECTION_RATIONALE_PANEL,
    "Slow degradation / residence-time mismatch": COLOR_REJECTION_RATIONALE_PANEL,
    "AD/biogas incompatibility": COLOR_REJECTION_RATIONALE_PANEL,
    "Explicit no-packaging / positive-list restriction": COLOR_REJECTION_RATIONALE_PANEL,
    "Pre-treatment/screening/equipment constraint": COLOR_REJECTION_RATIONALE_PANEL,
    "Equipment/operational disruption": COLOR_REJECTION_RATIONALE_PANEL,
    "Chemical/additive contamination concern": COLOR_REJECTION_RATIONALE_PANEL,
    "Mechanical pre-treatment/screening removal": COLOR_REJECTION_RATIONALE_PANEL,
    "Legal/positive-list/no-packaging exclusion": COLOR_REJECTION_RATIONALE_PANEL,
    "Residence-time/degradation mismatch": COLOR_REJECTION_RATIONALE_PANEL,
}

# Optional sequential palette for later heatmap-style figures.
RWTH_COLD_TO_WARM = [
    RWTH_BLUE_MUTED,
    "#B7CDD0",
    RWTH_GREEN_MUTED,
    "#E5D9BF",
    RWTH_WARM_MUTED,
]

# ============================================================
# 3. Helper functions
# ============================================================

def require_input_file(path: Path) -> None:
    if not path.exists():
        raise FileNotFoundError(
            f"Missing required input file:\n{path}\n\n"
            "Expected structure:\n"
            "compostable_packaging/\n"
            "  results/\n"
            "    display/\n"
            "      source_level_category_shares_display.csv\n"
            "      rejection_rationale_reporting_shares_display.csv"
        )


def wrap_label(text: str, width: int = 30, max_lines: int = 2) -> str:
    """Wrap category labels into at most two rows.

    Long labels are shortened with an ellipsis after the second row so that
    y-axis text does not consume excessive horizontal space.
    """
    words = str(text).split()
    lines: list[str] = []
    current = ""
    for word in words:
        candidate = word if not current else current + " " + word
        if len(candidate) <= width:
            current = candidate
            continue
        if current:
            lines.append(current)
        current = word
        if len(lines) == max_lines:
            break
    if len(lines) < max_lines and current:
        lines.append(current)

    label = "\n".join(lines[:max_lines])
    original = str(text)
    if len(" ".join(lines)) < len(original):
        split_lines = label.split("\n")
        split_lines[-1] = split_lines[-1].rstrip(" .,") + "…"
        label = "\n".join(split_lines)
    return label


def pct_label(value: float, n: int) -> str:
    return f"{value:.1f}%\n(n={int(n)})"


def get_indicator(df: pd.DataFrame, indicator: str) -> pd.DataFrame:
    out = df[df["indicator"].eq(indicator)].copy()
    if out.empty:
        raise ValueError(f"No rows found for indicator: {indicator}")
    return out


def order_by_share_desc(df: pd.DataFrame) -> pd.DataFrame:
    """Sort categories from largest to smallest share for display."""
    return (
        df.copy()
        .sort_values(["share_percent", "n", "category"], ascending=[False, False, True])
        .reset_index(drop=True)
    )


def style_axis(ax) -> None:
    """Use an open panel style while preserving the current font-size settings."""
    ax.tick_params(axis="both", colors=COLOR_TEXT)

    # Remove only the top/right box frame for a cleaner manuscript style.
    for spine in ["top", "right"]:
        ax.spines[spine].set_visible(False)

    # Keep left/bottom axes to anchor the category labels and x-axis scale.
    for spine in ["left", "bottom"]:
        ax.spines[spine].set_color("#94A3B8")
        ax.spines[spine].set_linewidth(0.7)


def add_horizontal_bar_labels(ax, values, counts, x_offset=1.0) -> None:
    for i, (value, count) in enumerate(zip(values, counts)):
        ax.text(
            value + x_offset,
            i,
            pct_label(value, count),
            va="center",
            ha="left",
            fontsize=FONT_SIZE_BAR_LABEL,
            color=COLOR_TEXT,
        )


# ============================================================
# 4. Plot helper
# ============================================================

def plot_ordered_bar_panel(
    ax,
    df: pd.DataFrame,
    title: str,
    xlabel: str,
    panel_color: str,
    label_width: int = 30,
    x_expand: float = 1.32,
) -> None:
    """Plot a coordinated horizontal bar panel sorted from high to low."""
    df = order_by_share_desc(df)
    y = np.arange(len(df))
    values = df["share_percent"].to_numpy(dtype=float)
    counts = df["n"].to_numpy(dtype=int)
    categories = df["category"].astype(str).tolist()

    ax.barh(y, values, color=panel_color, height=BAR_HEIGHT)
    ax.set_yticks(y)
    ax.set_yticklabels([wrap_label(c, label_width, max_lines=2) for c in categories], fontsize=FONT_SIZE_TICK_LABEL)
    ax.invert_yaxis()
    ax.set_xlim(0, max(values) * x_expand)
    ax.set_xlabel(xlabel, fontsize=FONT_SIZE_AXIS_LABEL)
    ax.set_title(title, loc="left", fontweight="bold", fontsize=FONT_SIZE_PANEL_TITLE, color=COLOR_TEXT)
    ax.grid(axis="x", alpha=GRID_ALPHA, color=COLOR_GRID, linewidth=GRID_WIDTH)
    add_horizontal_bar_labels(ax, values, counts, x_offset=max(values) * 0.018)
    style_axis(ax)


# ============================================================
# 5. Load data
# ============================================================

require_input_file(SOURCE_SHARE_PATH)
require_input_file(RATIONALE_SHARE_PATH)

source_shares = pd.read_csv(SOURCE_SHARE_PATH)
rationale_shares = pd.read_csv(RATIONALE_SHARE_PATH)

required_cols = {"indicator", "category", "n", "denominator", "share_percent"}
for name, df in [
    ("source_level_category_shares_display", source_shares),
    ("rejection_rationale_reporting_shares_display", rationale_shares),
]:
    missing = required_cols - set(df.columns)
    if missing:
        raise ValueError(f"{name} is missing required columns: {sorted(missing)}")


# ============================================================
# 6. Prepare panel data
# ============================================================

df_authority = order_by_share_desc(get_indicator(source_shares, "Source Authority"))
df_treatment = order_by_share_desc(get_indicator(source_shares, "Stated Organic-waste Treatment Route"))
df_acceptance = order_by_share_desc(get_indicator(source_shares, "Acceptance"))
df_rationale = order_by_share_desc(rationale_shares.copy())

# Export the exact data used for plotting.
plot_data = []
for panel, df in [
    ("a_source_authority", df_authority),
    ("b_treatment_route", df_treatment),
    ("c_acceptance", df_acceptance),
    ("d_rejection_rationale", df_rationale),
]:
    temp = df.copy()
    temp.insert(0, "panel", panel)
    plot_data.append(temp)
pd.concat(plot_data, ignore_index=True).to_csv(OUTPUT_DATA, index=False, encoding="utf-8-sig")



def add_panel_denominator(ax, text: str) -> None:
    """Add a small denominator note inside the panel."""
    ax.text(
        0.98,
        0.03,
        text,
        transform=ax.transAxes,
        ha="right",
        va="bottom",
        fontsize=FONT_SIZE_TICK_LABEL - 1,
        color=COLOR_MUTED_TEXT,
    )


# ============================================================
# 7. Plot
# ============================================================

fig, axes = plt.subplots(2, 2, figsize=(FIG_WIDTH, FIG_HEIGHT))
ax_a, ax_b, ax_c, ax_d = axes.flatten()

plot_ordered_bar_panel(
    ax=ax_a,
    df=df_authority,
    title="a. Rule authority",
    xlabel="Share of included sources (%)",
    panel_color=PANEL_COLORS["source_authority"],
    label_width=31,
)

plot_ordered_bar_panel(
    ax=ax_b,
    df=df_treatment,
    title="b. Stated treatment route",
    xlabel="Share of included sources (%)",
    panel_color=PANEL_COLORS["treatment_route"],
    label_width=31,
)

plot_ordered_bar_panel(
    ax=ax_c,
    df=df_acceptance,
    title="c. Acceptance outcome",
    xlabel="Share of included sources (%)",
    panel_color=PANEL_COLORS["acceptance"],
    label_width=31,
)

plot_ordered_bar_panel(
    ax=ax_d,
    df=df_rationale,
    title="d. Rejection rationale",
    xlabel="Share of sources with specified rationale (%)",
    panel_color=PANEL_COLORS["rejection_rationale"],
    label_width=31,
)

add_panel_denominator(ax_a, "n = 184 included sources")
add_panel_denominator(ax_b, "n = 184 included sources")
add_panel_denominator(ax_c, "n = 184 included sources")
add_panel_denominator(ax_d, "n = 97 sources with explicit\nrejection/restriction rationale")

# No legend is drawn: categories are directly labelled on each bar panel.

fig.subplots_adjust(
    left=LEFT_MARGIN,
    right=RIGHT_MARGIN,
    top=TOP_MARGIN,
    bottom=BOTTOM_MARGIN,
    wspace=PANEL_WSPACE,
    hspace=PANEL_HSPACE,
)

fig.savefig(OUTPUT_PNG, dpi=DPI, bbox_inches="tight")
fig.savefig(OUTPUT_PDF, bbox_inches="tight")
plt.close(fig)

print(f"Input display folder: {DISPLAY_DIR}")
print(f"Saved: {OUTPUT_PNG}")
print(f"Saved: {OUTPUT_PDF}")
print(f"Saved: {OUTPUT_DATA}")
print("Panel icons and legend are disabled; no external icon downloads are attempted.")
