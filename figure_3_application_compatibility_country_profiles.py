#!/usr/bin/env python3
"""
Generate a two-panel application-by-country figure with country-only clustering.

Panel A:
    Application-level compatibility shares under lower, central, and upper
    scenarios. Bars show the central scenario; whiskers show the lower–upper
    scenario range.

Panel B:
    Country × Application Group evidence-status matrix. Countries are ordered
    by hierarchical clustering based on central-scenario compatibility shares,
    while Application Group order is kept fixed. The dendrogram is not shown;
    clustered country groups are shown using separator lines and right-side
    bracket labels.

Input structure:
    The script assumes the final display CSV files are always located at:

        results/display/

    relative to the folder where this script is run.

Required input files:
    results/display/application_decision_long_display.csv
    results/display/application_compatibility_scenarios_display.csv
    results/display/source_level_coding_display.csv

Outputs:
    figures/figure_3_application_compatibility_country_profiles.png
    figures/figure_3_application_compatibility_country_profiles.pdf

PDF export note:
    The heatmap is drawn as editable vector rectangles rather than imshow/raster.
    This avoids Illustrator reinterpreting rasterized colormap colours during PDF import.
    figures/figure_3_application_compatibility_country_profiles_status.csv
    figures/figure_application_country_clustered_order.csv
"""

from __future__ import annotations

from pathlib import Path

import matplotlib
matplotlib.use("Agg")  # file-only rendering; avoids interactive display issues

import matplotlib.pyplot as plt
from matplotlib.colors import BoundaryNorm, ListedColormap
from matplotlib.lines import Line2D
from matplotlib.patches import Rectangle
import numpy as np
import pandas as pd
from scipy.cluster.hierarchy import fcluster, leaves_list, linkage


# ============================================================
# 1. Paths and basic settings
# ============================================================

SCRIPT_DIR = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()

DEFAULT_DISPLAY_DIR = SCRIPT_DIR / "results" / "display"
# Primary local workflow: display CSVs inside results/display/.
# Fallback: allow running this script from a folder where the display CSVs sit
# next to the script, which is useful when testing or exchanging figure scripts.
_REQUIRED_DISPLAY_FILENAMES = [
    "application_decision_long_display.csv",
    "application_compatibility_scenarios_display.csv",
    "source_level_coding_display.csv",
]
if DEFAULT_DISPLAY_DIR.exists():
    DISPLAY_DIR = DEFAULT_DISPLAY_DIR
elif all((SCRIPT_DIR / name).exists() for name in _REQUIRED_DISPLAY_FILENAMES):
    DISPLAY_DIR = SCRIPT_DIR
else:
    DISPLAY_DIR = DEFAULT_DISPLAY_DIR
OUTPUT_DIR = SCRIPT_DIR / "figures"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

OUTPUT_PNG = OUTPUT_DIR / "figure_3_application_compatibility_country_profiles.png"
OUTPUT_PDF = OUTPUT_DIR / "figure_3_application_compatibility_country_profiles.pdf"
OUTPUT_STATUS_CSV = OUTPUT_DIR / "figure_3_application_compatibility_country_profiles_status.csv"
OUTPUT_ORDER_CSV = OUTPUT_DIR / "figure_application_country_clustered_order.csv"

# Clustering settings
N_CLUSTERS = 4
CLUSTER_METHOD = "ward"
CLUSTER_METRIC = "euclidean"

# Figure settings
DPI = 300
FIG_WIDTH = 15.8

# Increase this if country rows look crowded after changing font sizes.
# Current value is set higher than the previous 0.34 to accommodate larger text.
ROW_HEIGHT = 0.55
MIN_FIG_HEIGHT = 13.5

# Font settings
# Arial will be used when available; Matplotlib will fall back to DejaVu Sans otherwise.
FONT_FAMILY = "Arial"
FONT_FALLBACKS = ["Arial", "DejaVu Sans"]

# Main font-size controls. These are intentionally centralized so the whole
# figure can be scaled later without searching through the plotting code.
FONT_SIZE_PANEL_TITLE = 18
FONT_SIZE_AXIS_LABEL = 16
FONT_SIZE_TICK_LABEL = 16
FONT_SIZE_BAR_LABEL = 16
FONT_SIZE_COUNTRY_LABEL = 15
FONT_SIZE_APPLICATION_LABEL = 16
FONT_SIZE_CELL_TEXT = 16
FONT_SIZE_GROUP_LABEL = 16
FONT_SIZE_LEGEND = 14

# Legend symbol controls
LEGEND_MARKER_SIZE = 13
LEGEND_PANEL_A_BAR_WIDTH = 7
LEGEND_PANEL_A_WHISKER_WIDTH = 1.4

# Legend layout control.
# With the current 12 legend entries, 3 columns gives 4 rows.
LEGEND_NCOL = 3

# Layout controls
GRID_WIDTH = 1.3
SEPARATOR_WIDTH = 1
BRACKET_WIDTH = 0.8

# Extend only the horizontal cluster-separator lines slightly to the left of
# the heatmap. This does not draw a vertical bracket and does not create an
# empty heatmap column.
# More negative values extend the separator farther left.
LEFT_SEPARATOR_X = -2

PANEL_VERTICAL_GAP = 0.08
BOTTOM_MARGIN = 0.15

# Country evidence label abbreviations
SOURCE_ABBR = "SR"
APP_RECORD_ABBR = "SA"

# ============================================================
# 2. Modular colour parameters
# ============================================================
# Keep all figure colours here so they can be harmonized across figures later.

COLOR_PANEL_A_BAR = "#B8D0E4"
COLOR_PANEL_A_ERROR = "#1F2937"
COLOR_PANEL_A_GRID = "#E5E7EB"
COLOR_TEXT = "#111827"
COLOR_MUTED_TEXT = "#6B7280"
COLOR_SEPARATOR = "#334155"

# Cell-status colours. These are solid pre-lightened RGB hex colours, not alpha-blended colours.
# That keeps the PDF colours stable when opened as editable vector artwork in Illustrator.
COLOR_NO_RECORD = "#F1F1F1"
COLOR_UNCLEAR_ONLY = "#D7D7D7"
COLOR_REJECTED_DOMINANT = "#E3D2B3"
COLOR_COMPATIBLE_DOMINANT = "#DDE8C4"
COLOR_BALANCED_CONFLICT = "#CFC6DA"

STATUS_COLORS = {
    0: COLOR_NO_RECORD,
    1: COLOR_UNCLEAR_ONLY,
    2: COLOR_REJECTED_DOMINANT,
    3: COLOR_COMPATIBLE_DOMINANT,
    4: COLOR_BALANCED_CONFLICT,
}

# Optional group-label colours. By default, neutral text is used.
COLOR_GROUP_LABEL = COLOR_MUTED_TEXT

# Apply font settings globally. Arial will be used where installed.
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
    "pdf.compression": 0,  # larger file, but easier for Illustrator to import/edit reliably
    "image.composite_image": False,  # avoid unexpected colour compositing in vector backends
})


# ============================================================
# 3. Helper functions
# ============================================================

def find_col(df: pd.DataFrame, candidates: list[str]) -> str:
    """Return the first matching column name from candidates, case-insensitive."""
    lower_map = {str(c).strip().lower(): c for c in df.columns}
    for candidate in candidates:
        key = candidate.strip().lower()
        if key in lower_map:
            return lower_map[key]
    raise KeyError(
        f"None of these columns found: {candidates}\n"
        f"Available columns: {list(df.columns)}"
    )


def safe_str_series(s: pd.Series) -> pd.Series:
    return s.fillna("").astype(str).str.strip()


def wrap_application_label(label: str) -> str:
    """
    Use full application names, wrapped into compact multi-line labels.
    No rotation is used in the final figure.
    """
    label_map = {
        "Food-waste liners / collection bags": "Food-waste\nliners /\ncollection bags",
        "Food-service ware / takeaway packaging": "Food-service\nware / takeaway\npackaging",
        "Tea/coffee preparation items": "Tea/coffee\npreparation\nitems",
        "Food-soiled paper / fibre packaging": "Food-soiled\npaper / fibre\npackaging",
        "Shopping/produce bags": "Shopping /\nproduce\nbags",
        "Flexible films/wraps/pouches": "Flexible films /\nwraps /\npouches",
        "Generic compostable packaging / plastics": "Generic\ncompostable\npackaging / plastics",
    }
    return label_map.get(label, label)





def require_input_file(path: Path) -> None:
    if not path.exists():
        raise FileNotFoundError(
            f"Missing required input file:\n{path}\n\n"
            "Expected structure:\n"
            "compostable_packaging/\n"
            "  results/\n"
            "    display/\n"
            "      application_decision_long_display.csv\n"
            "      application_compatibility_scenarios_display.csv\n"
            "      source_level_coding_display.csv"
        )


def compact_country_label(country: str, source_counts: pd.Series, app_record_counts: pd.Series) -> str:
    """Compact country label.

    Format: Country (SR;SA), where SR = source count and SA =
    source × application-group record count. The abbreviations are explained
    in the figure legend/caption rather than repeated in every country label.
    """
    return (
        f"{country} "
        f"({int(source_counts.get(country, 0))};"
        f"{int(app_record_counts.get(country, 0))})"
    )


def describe_cluster_profile(cluster_df: pd.DataFrame) -> str:
    """
    Assign a short descriptive label to a cluster from its displayed country × application rows.

    This is intentionally simple and transparent. It uses the visible C/R-derived
    status profile, not hidden model parameters.
    """
    if cluster_df.empty:
        return "Sparse evidence"

    decisive = cluster_df[cluster_df["n_compatible"].add(cluster_df["n_rejected"]).gt(0)]
    if decisive.empty:
        return "Sparse evidence"

    total_c = decisive["n_compatible"].sum()
    total_r = decisive["n_rejected"].sum()
    total_decisive = total_c + total_r

    liner = decisive[decisive["application_group"].eq("Food-waste liners / collection bags")]
    liner_c = liner["n_compatible"].sum()
    liner_r = liner["n_rejected"].sum()

    non_liner = decisive[~decisive["application_group"].eq("Food-waste liners / collection bags")]
    non_liner_c = non_liner["n_compatible"].sum()
    non_liner_r = non_liner["n_rejected"].sum()

    # Conflict means at least one visible balanced cell or both compatible and rejected
    # records present within the cluster for a major application.
    has_balanced = (decisive["status"].eq("Balanced conflict")).any()

    if total_decisive <= 2:
        return "Sparse evidence"

    if total_c > total_r and non_liner_c >= max(2, non_liner_r):
        return "Broader compatibility"

    if liner_c > liner_r and non_liner_r > non_liner_c:
        return "Liner-focused"

    if has_balanced or (liner_c > 0 and liner_r > 0):
        return "Fragmented"

    if total_r >= total_c:
        return "Restrictive"

    return "Mixed profile"


def wrap_cluster_label(label: str) -> str:
    """Wrap right-side cluster labels into short horizontal multi-line text."""
    label_map = {
        "Broader compatibility": "Broader\ncompatibility",
        "Liner-focused": "Liner-\nfocused",
        "Fragmented": "Fragmented",
        "Restrictive": "Restrictive",
        "Mixed profile": "Mixed\nprofile",
        "Sparse evidence": "Sparse\nevidence",
    }
    return label_map.get(label, label)


def draw_vector_status_matrix(ax: plt.Axes, status_values: np.ndarray) -> None:
    """Draw heatmap cells as solid vector rectangles for Illustrator-safe PDFs.

    Matplotlib's imshow can be embedded in PDFs as a raster image. Acrobat often
    previews that correctly, but Illustrator may reinterpret the raster colour
    profile or compositing when the PDF is opened for editing. Drawing each cell
    as a solid RGB rectangle keeps the colour mapping editable and stable.
    """
    n_rows, n_cols = status_values.shape
    for i in range(n_rows):
        for j in range(n_cols):
            try:
                status_value = int(status_values[i, j])
            except Exception:
                status_value = 0
            ax.add_patch(
                Rectangle(
                    (j - 0.5, i - 0.5),
                    1.0,
                    1.0,
                    facecolor=STATUS_COLORS.get(status_value, COLOR_NO_RECORD),
                    edgecolor="white",
                    linewidth=GRID_WIDTH,
                    alpha=1.0,
                    antialiased=False,
                    joinstyle="miter",
                    zorder=0,
                )
            )
    ax.set_xlim(-0.5, n_cols - 0.5)
    ax.set_ylim(n_rows - 0.5, -0.5)
    ax.set_aspect("auto")


# ============================================================
# 4. Load input data
# ============================================================

app_path = DISPLAY_DIR / "application_decision_long_display.csv"
comp_path = DISPLAY_DIR / "application_compatibility_scenarios_display.csv"
src_path = DISPLAY_DIR / "source_level_coding_display.csv"

for p in [app_path, comp_path, src_path]:
    require_input_file(p)

app = pd.read_csv(app_path)
comp = pd.read_csv(comp_path)
src = pd.read_csv(src_path)

country_col = find_col(app, ["country_or_region", "country", "Country / region"])
app_group_col = find_col(app, ["application_type", "application_group", "Application Group"])
decision_col = find_col(app, ["application_decision", "Application Decision"])

src_country_col = find_col(src, ["country_or_region", "country", "Country / region"])
src_id_col = find_col(src, ["citation_id", "Citation ID", "source_id", "reference_id"])

comp_app_col = find_col(comp, ["application_type", "application_group", "Application Group"])

app[country_col] = safe_str_series(app[country_col])
app[app_group_col] = safe_str_series(app[app_group_col])
app[decision_col] = safe_str_series(app[decision_col])

src[src_country_col] = safe_str_series(src[src_country_col])
src[src_id_col] = safe_str_series(src[src_id_col])

comp[comp_app_col] = safe_str_series(comp[comp_app_col])


# ============================================================
# 5. Application order and decision categories
# ============================================================

preferred_application_order = [
    "Food-waste liners / collection bags",
    "Food-service ware / takeaway packaging",
    "Tea/coffee preparation items",
    "Food-soiled paper / fibre packaging",
    "Shopping/produce bags",
    "Flexible films/wraps/pouches",
    "Generic compostable packaging / plastics",
]

summary_order = comp[comp_app_col].dropna().astype(str).tolist()
application_order = [a for a in preferred_application_order if a in set(summary_order)]
if len(application_order) == 0:
    application_order = summary_order

wrapped_application_labels = [wrap_application_label(a) for a in application_order]


# ============================================================
# 6. Build source and application-record counts
# ============================================================

source_counts = (
    src.groupby(src_country_col)[src_id_col]
    .nunique()
    .rename("n_sources")
)

app_record_counts = (
    app.groupby(country_col)
    .size()
    .rename("n_application_records")
)

all_countries = sorted(source_counts.index.astype(str).tolist())


# ============================================================
# 7. Build country × application count/status table
# ============================================================

group_counts = (
    app.groupby([country_col, app_group_col, decision_col])
    .size()
    .unstack(fill_value=0)
)

for col in ["Accepted", "Accepted with conditions", "Rejected", "Unclear or not stated"]:
    if col not in group_counts.columns:
        group_counts[col] = 0

records = []

for country in all_countries:
    for application in application_order:
        try:
            row = group_counts.loc[(country, application)]
            n_accepted = int(row.get("Accepted", 0))
            n_conditional = int(row.get("Accepted with conditions", 0))
            n_rejected = int(row.get("Rejected", 0))
            n_unclear = int(row.get("Unclear or not stated", 0))
        except KeyError:
            n_accepted = 0
            n_conditional = 0
            n_rejected = 0
            n_unclear = 0

        n_compatible = n_accepted + n_conditional
        n_decisive = n_compatible + n_rejected
        n_total = n_decisive + n_unclear

        if n_decisive > 0:
            cluster_value = n_compatible / n_decisive
            cell_text = f"{n_compatible}/{n_rejected}"

            if n_compatible > n_rejected:
                status = "Compatible-dominant"
                status_value = 3
            elif n_rejected > n_compatible:
                status = "Rejected-dominant"
                status_value = 2
            else:
                status = "Balanced conflict"
                status_value = 4
        elif n_unclear > 0:
            cluster_value = np.nan
            cell_text = "u"
            status = "Unclear only"
            status_value = 1
        else:
            cluster_value = np.nan
            cell_text = "—"
            status = "No application-level record"
            status_value = 0

        records.append({
            "country": country,
            "application_group": application,
            "n_accepted": n_accepted,
            "n_accepted_with_conditions": n_conditional,
            "n_compatible": n_compatible,
            "n_rejected": n_rejected,
            "n_unclear": n_unclear,
            "n_total_records": n_total,
            "cell_text": cell_text,
            "status": status,
            "status_value": status_value,
            "cluster_value": cluster_value,
        })

matrix_long = pd.DataFrame(records)

status_matrix = (
    matrix_long
    .pivot(index="country", columns="application_group", values="status_value")
    .reindex(index=all_countries, columns=application_order)
)

cell_text_matrix = (
    matrix_long
    .pivot(index="country", columns="application_group", values="cell_text")
    .reindex(index=all_countries, columns=application_order)
)

cluster_matrix = (
    matrix_long
    .pivot(index="country", columns="application_group", values="cluster_value")
    .reindex(index=all_countries, columns=application_order)
)


# ============================================================
# 8. Country-only hierarchical clustering
# ============================================================

countries_with_decisive = cluster_matrix.index[cluster_matrix.notna().any(axis=1)].tolist()
countries_without_decisive = cluster_matrix.index[~cluster_matrix.notna().any(axis=1)].tolist()

if len(countries_with_decisive) > 1:
    cluster_input = cluster_matrix.loc[countries_with_decisive].copy()

    # Column-mean imputation for clustering only.
    for col in cluster_input.columns:
        col_mean = cluster_input[col].mean(skipna=True)
        if pd.isna(col_mean):
            col_mean = 0.0
        cluster_input[col] = cluster_input[col].fillna(col_mean)

    Z = linkage(cluster_input.values, method=CLUSTER_METHOD, metric=CLUSTER_METRIC)
    leaf_order = leaves_list(Z)
    clustered_countries = [countries_with_decisive[i] for i in leaf_order]

    cluster_labels_raw = fcluster(Z, t=N_CLUSTERS, criterion="maxclust")
    cluster_label_by_country = {
        country: int(label)
        for country, label in zip(countries_with_decisive, cluster_labels_raw)
    }
else:
    clustered_countries = countries_with_decisive
    cluster_label_by_country = {c: 1 for c in countries_with_decisive}

# Countries with no decisive record go at the bottom.
country_order = clustered_countries + countries_without_decisive

status_plot = status_matrix.reindex(index=country_order, columns=application_order)
cell_text_plot = cell_text_matrix.reindex(index=country_order, columns=application_order)


# ============================================================
# 9. Build contiguous displayed cluster blocks and descriptive labels
# ============================================================

displayed_clusters = [cluster_label_by_country.get(c, np.nan) for c in country_order]

blocks = []
start = 0
while start < len(country_order):
    current = displayed_clusters[start]
    end = start + 1
    while end < len(country_order):
        next_value = displayed_clusters[end]
        same_cluster = (pd.isna(current) and pd.isna(next_value)) or (next_value == current)
        if not same_cluster:
            break
        end += 1

    countries_in_block = country_order[start:end]
    if pd.isna(current):
        block_label = "Sparse evidence"
    else:
        block_df = matrix_long[matrix_long["country"].isin(countries_in_block)]
        block_label = describe_cluster_profile(block_df)

    blocks.append({
        "start": start,
        "end": end,
        "cluster_raw": current,
        "label": block_label,
        "countries": countries_in_block,
    })
    start = end

# Merge adjacent displayed blocks when they have the same descriptive label.
merged_blocks = []
for block in blocks:
    if merged_blocks and merged_blocks[-1]["label"] == block["label"]:
        merged_blocks[-1]["end"] = block["end"]
        merged_blocks[-1]["countries"].extend(block["countries"])
    else:
        merged_blocks.append({
            "start": block["start"],
            "end": block["end"],
            "cluster_raw": block["cluster_raw"],
            "label": block["label"],
            "countries": list(block["countries"]),
        })
blocks = merged_blocks


# ============================================================
# 10. Panel A values: lower, central, upper scenario shares
# ============================================================

lower_col = find_col(comp, ["lower_bound_compatible_share"])
central_col = find_col(comp, ["central_compatible_share"])
upper_col = find_col(comp, ["upper_bound_compatible_share"])

panel_a = comp[comp[comp_app_col].isin(application_order)].copy()
panel_a[comp_app_col] = pd.Categorical(panel_a[comp_app_col], categories=application_order, ordered=True)
panel_a = panel_a.sort_values(comp_app_col)

lower_values = panel_a[lower_col].astype(float).to_numpy()
central_values = panel_a[central_col].astype(float).to_numpy()
upper_values = panel_a[upper_col].astype(float).to_numpy()

yerr_lower = central_values - lower_values
yerr_upper = upper_values - central_values
yerr = np.vstack([yerr_lower, yerr_upper])


# ============================================================
# 11. Export matrix and order diagnostics
# ============================================================

matrix_long_out = matrix_long.copy()
matrix_long_out["country_order"] = matrix_long_out["country"].map(
    {country: i + 1 for i, country in enumerate(country_order)}
)
matrix_long_out["cluster_label_raw"] = matrix_long_out["country"].map(cluster_label_by_country)
matrix_long_out.to_csv(OUTPUT_STATUS_CSV, index=False, encoding="utf-8-sig")

order_out = pd.DataFrame({
    "country_order": range(1, len(country_order) + 1),
    "country": country_order,
    "cluster_label_raw": [cluster_label_by_country.get(c, np.nan) for c in country_order],
    "cluster_profile_label": [
        next((b["label"] for b in blocks if c in b["countries"]), "Sparse evidence")
        for c in country_order
    ],
    "n_sources": [int(source_counts.get(c, 0)) for c in country_order],
    "n_application_records": [int(app_record_counts.get(c, 0)) for c in country_order],
})
order_out.to_csv(OUTPUT_ORDER_CSV, index=False, encoding="utf-8-sig")


# ============================================================
# 12. Plot
# ============================================================

cmap = ListedColormap([STATUS_COLORS[i] for i in range(5)])
norm = BoundaryNorm([-0.5, 0.5, 1.5, 2.5, 3.5, 4.5], cmap.N)

n_countries = len(country_order)
n_apps = len(application_order)

fig_height = max(MIN_FIG_HEIGHT, n_countries * ROW_HEIGHT)

fig = plt.figure(figsize=(FIG_WIDTH, fig_height))

# Layout:
# top-left: Panel A
# top-right: empty group-label margin
# bottom-left: Panel B matrix
# bottom-right: vertical group labels
gs = fig.add_gridspec(
    nrows=2,
    ncols=2,
    width_ratios=[7.0, 1.25],
    height_ratios=[1.08, max(5.0, n_countries * 0.18)],
    hspace=PANEL_VERTICAL_GAP,
    wspace=0.03,
)

ax_a = fig.add_subplot(gs[0, 0])
ax_empty = fig.add_subplot(gs[0, 1])
ax_empty.axis("off")

ax_b = fig.add_subplot(gs[1, 0], sharex=ax_a)
ax_group = fig.add_subplot(gs[1, 1], sharey=ax_b)


# ------------------------------------------------------------
# Panel A: application compatibility scenario bars
# ------------------------------------------------------------

x = np.arange(n_apps)

ax_a.bar(x, central_values, width=0.72, color=COLOR_PANEL_A_BAR)
ax_a.errorbar(
    x,
    central_values,
    yerr=yerr,
    fmt="none",
    ecolor=COLOR_PANEL_A_ERROR,
    elinewidth=1.0,
    capsize=3,
)

ax_a.set_ylim(0, 1)
ax_a.set_ylabel("Compatibility share", color=COLOR_TEXT, fontsize=FONT_SIZE_AXIS_LABEL)
ax_a.set_title(
    "a. Application-level compatibility scenario",
    loc="left",
    fontweight="bold",
    color=COLOR_TEXT,
    fontsize=FONT_SIZE_PANEL_TITLE,
)

ax_a.set_xticks(x)
ax_a.tick_params(axis="x", bottom=False, labelbottom=False)
ax_a.tick_params(axis="y", labelsize=FONT_SIZE_TICK_LABEL, colors=COLOR_TEXT)
# No background grid lines in Panel A.
ax_a.grid(False)

for i, value in enumerate(central_values):
    # Place percentage labels above the upper whisker to avoid overlap.
    label_y = min(upper_values[i] + 0.055, 0.98)
    ax_a.text(
        i,
        label_y,
        f"{value * 100:.1f}%",
        ha="center",
        va="bottom",
        fontsize=FONT_SIZE_BAR_LABEL,
        color=COLOR_TEXT,
    )


# ------------------------------------------------------------
# Right-side cluster brackets and labels
# ------------------------------------------------------------

ax_group.set_xlim(0, 1)
ax_group.set_ylim(n_countries - 0.5, -0.5)
ax_group.axis("off")

# Bracket geometry in the right-side label axis.
# The y-axis uses heatmap row coordinates; x-axis runs from 0 to 1.
BRACKET_X = 0.06
BRACKET_TICK_X = 0.18
LABEL_X = 0.25

for block in blocks:
    y_top = block["start"] - 0.5
    y_bottom = block["end"] - 0.5
    y_mid = (block["start"] + block["end"] - 1) / 2

    # Draw a bracket to show the vertical range covered by this cluster profile.
    ax_group.plot(
        [BRACKET_X, BRACKET_X],
        [y_top, y_bottom],
        color=COLOR_SEPARATOR,
        linewidth=BRACKET_WIDTH,
        solid_capstyle="butt",
        clip_on=False,
    )
    ax_group.plot(
        [BRACKET_X, BRACKET_TICK_X],
        [y_top, y_top],
        color=COLOR_SEPARATOR,
        linewidth=BRACKET_WIDTH,
        solid_capstyle="butt",
        clip_on=False,
    )
    ax_group.plot(
        [BRACKET_X, BRACKET_TICK_X],
        [y_bottom, y_bottom],
        color=COLOR_SEPARATOR,
        linewidth=BRACKET_WIDTH,
        solid_capstyle="butt",
        clip_on=False,
    )

    ax_group.text(
        LABEL_X,
        y_mid,
        wrap_cluster_label(block["label"]),
        ha="left",
        va="center",
        rotation=0,
        fontsize=FONT_SIZE_GROUP_LABEL,
        fontweight="bold",
        color=COLOR_GROUP_LABEL,
        linespacing=1.0,
        clip_on=False,
    )


# ------------------------------------------------------------
# Panel B: country × application evidence-status matrix
# ------------------------------------------------------------

# Draw heatmap as vector rectangles, not imshow/raster, to keep colours stable
# when opening the PDF in Adobe Illustrator.
draw_vector_status_matrix(ax_b, status_plot.values.astype(float))

ax_b.set_title(
    "b. Country-level compatibility profiles",
    loc="left",
    fontweight="bold",
    color=COLOR_TEXT,
    fontsize=FONT_SIZE_PANEL_TITLE,
)

country_labels = [
    compact_country_label(country, source_counts, app_record_counts)
    for country in country_order
]

ax_b.set_yticks(np.arange(n_countries))
ax_b.set_yticklabels(country_labels, fontsize=FONT_SIZE_COUNTRY_LABEL, color=COLOR_TEXT)

ax_b.set_xticks(x)
x_tick_texts = ax_b.set_xticklabels(
    wrapped_application_labels,
    rotation=0,
    ha="center",
    va="top",
    fontsize=FONT_SIZE_APPLICATION_LABEL,
)
for t in x_tick_texts:
    t.set_linespacing(0.95)
ax_b.tick_params(axis="x", labeltop=False, top=False, bottom=True, labelbottom=True, pad=10)

# Add cell text.
for i in range(n_countries):
    for j in range(n_apps):
        txt = cell_text_plot.iloc[i, j]
        if pd.isna(txt):
            txt = "—"
        ax_b.text(j, i, str(txt), ha="center", va="center", fontsize=FONT_SIZE_CELL_TEXT, color=COLOR_TEXT)

# Cell borders are drawn directly as vector rectangle edges. Keep minor ticks off
# to avoid duplicate grid objects in Illustrator.
ax_b.set_xticks(np.arange(-0.5, n_apps, 1), minor=True)
ax_b.set_yticks(np.arange(-0.5, n_countries, 1), minor=True)
ax_b.tick_params(which="minor", bottom=False, left=False)

# Cluster separator lines across the heatmap, extended slightly to the left.
# The extension is drawn outside the heatmap with clip_on=False, but the x-axis
# limits are reset afterwards so it does not create an empty column.
for block in blocks[1:]:
    y = block["start"] - 0.5
    ax_b.plot(
        [LEFT_SEPARATOR_X, n_apps - 0.5],
        [y, y],
        color=COLOR_SEPARATOR,
        linewidth=SEPARATOR_WIDTH,
        solid_capstyle="butt",
        clip_on=False,
    )

# Keep the heatmap x-limits fixed after drawing outside-axis separator extensions.
ax_b.set_xlim(-0.5, n_apps - 0.5)

# Legend: no title, left-aligned.
legend_elements = [
    Line2D([0], [0], color=COLOR_PANEL_A_BAR, lw=LEGEND_PANEL_A_BAR_WIDTH,
           label="Panel A bar = central scenario"),
    Line2D([0], [0], color=COLOR_PANEL_A_ERROR, lw=LEGEND_PANEL_A_WHISKER_WIDTH, marker="_",
           markersize=LEGEND_MARKER_SIZE, label="Panel A whisker = lower–upper scenario range"),
    Line2D([0], [0], marker="s", color="w", label="Compatible-dominant",
           markerfacecolor=STATUS_COLORS[3], markersize=LEGEND_MARKER_SIZE),
    Line2D([0], [0], marker="s", color="w", label="Rejected-dominant",
           markerfacecolor=STATUS_COLORS[2], markersize=LEGEND_MARKER_SIZE),
    Line2D([0], [0], marker="s", color="w", label="Balanced conflict",
           markerfacecolor=STATUS_COLORS[4], markersize=LEGEND_MARKER_SIZE),
    Line2D([0], [0], marker="s", color="w", label="Unclear only",
           markerfacecolor=STATUS_COLORS[1], markersize=LEGEND_MARKER_SIZE),
    Line2D([0], [0], marker="s", color="w", label="No application-level record",
           markerfacecolor=STATUS_COLORS[0], markersize=LEGEND_MARKER_SIZE),
    Line2D([], [], linestyle="None", label="Cell text C/R = compatible/rejected records"),
    Line2D([], [], linestyle="None", label="Example: 5/2 = 5 compatible, 2 rejected"),
    Line2D([], [], linestyle="None", label="u = unclear-only record"),
    Line2D([], [], linestyle="None", label="Country label example: (3;8)"),
    Line2D([], [], linestyle="None", label="3 = sources; 8 = source × application-group records"),

]

legend = fig.legend(
    handles=legend_elements,
    loc="lower center",
    bbox_to_anchor=(0.50, 0.018),
    frameon=False,
    ncol=LEGEND_NCOL,
    handlelength=1.4,
    handletextpad=0.6,
    columnspacing=1.2,
    alignment="left",
    fontsize=FONT_SIZE_LEGEND,
)
try:
    legend._legend_box.align = "left"
except Exception:
    pass

# Manual layout is more stable than tight_layout for outside legends.
fig.subplots_adjust(
    left=0.12,
    right=0.97,
    top=0.96,
    bottom=BOTTOM_MARGIN,
)

fig.savefig(OUTPUT_PNG, dpi=DPI, bbox_inches="tight", facecolor="white", edgecolor="white", transparent=False)
fig.savefig(OUTPUT_PDF, bbox_inches="tight", facecolor="white", edgecolor="white", transparent=False)
plt.close(fig)

print(f"Input display folder: {DISPLAY_DIR}")
print(f"Saved: {OUTPUT_PNG}")
print(f"Saved: {OUTPUT_PDF}")
print(f"Saved: {OUTPUT_STATUS_CSV}")
print(f"Saved: {OUTPUT_ORDER_CSV}")
print(f"Countries shown ({len(country_order)}): {', '.join(country_order)}")
print(f"Clustering: method={CLUSTER_METHOD}, metric={CLUSTER_METRIC}, clusters={N_CLUSTERS}")
print("Cluster blocks:")
for block in blocks:
    print(f"  rows {block['start'] + 1}-{block['end']}: {block['label']} ({len(block['countries'])} countries)")
