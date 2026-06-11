#!/usr/bin/env python3
"""
Generate Figure 4: Certification Sufficiency.

Updated logic
-------------
Panel A:
    Source-level Certification and Approval Basis for all included sources.
    The seven categories are shown as badge cards. The five certification-
    relevant basis categories used for Panel B are marked as "Included in
    Panel B"; generic compostable/biodegradable wording and no named standard/
    approval are marked as "Not in Panel B denominator".

Panel B:
    Certification Sufficiency × Application Group bubble matrix using the
    revised four-category Certification Sufficiency indicator:
        - Certification accepted
        - Certification conditionally accepted
        - Certification rejected
        - Certification unclear or not stated

    Bubble area shows certification-relevant source × application-group records.
    The script reads certification_sufficiency_bubble_matrix_display.csv when
    available, so the plotted matrix is the same base data used in the Excel
    workbook.

Required input files:
    source_level_category_shares_display.csv
    certification_sufficiency_bubble_matrix_display.csv
        OR certification_sufficiency_by_application_display.csv

Outputs:
    figures/figure_4_certification_sufficiency.png
    figures/figure_4_certification_sufficiency.pdf
    figures/figure_4_certification_sufficiency.csv
"""

from __future__ import annotations

from pathlib import Path
import textwrap

import matplotlib
matplotlib.use("Agg")

import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Rectangle
from matplotlib.lines import Line2D
import numpy as np
import pandas as pd


# ============================================================
# 1. Paths
# ============================================================

SCRIPT_DIR = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()

DISPLAY_DIR_CANDIDATES = [
    SCRIPT_DIR / "results" / "display",
    SCRIPT_DIR,
]

DISPLAY_DIR = None
for candidate in DISPLAY_DIR_CANDIDATES:
    if (candidate / "source_level_category_shares_display.csv").exists() and (
        (candidate / "certification_sufficiency_bubble_matrix_display.csv").exists()
        or (candidate / "certification_sufficiency_by_application_display.csv").exists()
    ):
        DISPLAY_DIR = candidate
        break

if DISPLAY_DIR is None:
    raise FileNotFoundError(
        "Could not find required display CSVs. Expected source_level_category_shares_display.csv "
        "and certification_sufficiency_bubble_matrix_display.csv or certification_sufficiency_by_application_display.csv "
        "in one of the configured display folders."
    )

PRIMARY_OUTPUT_DIR = SCRIPT_DIR / "figures"
FALLBACK_OUTPUT_DIR = SCRIPT_DIR / "figures_fallback"


def choose_output_dir(primary: Path, fallback: Path) -> Path:
    for candidate in [primary, fallback]:
        try:
            candidate.mkdir(parents=True, exist_ok=True)
            test_file = candidate / ".write_test"
            test_file.write_text("ok", encoding="utf-8")
            test_file.unlink(missing_ok=True)
            return candidate
        except OSError:
            continue
    raise PermissionError(f"Neither {primary} nor {fallback} is writable.")


OUTPUT_DIR = choose_output_dir(PRIMARY_OUTPUT_DIR, FALLBACK_OUTPUT_DIR)
OUTPUT_PNG = OUTPUT_DIR / "figure_4_certification_sufficiency.png"
OUTPUT_PDF = OUTPUT_DIR / "figure_4_certification_sufficiency.pdf"
OUTPUT_MATRIX_CSV = OUTPUT_DIR / "figure_4_certification_sufficiency.csv"


# ============================================================
# 2. Layout, fonts, and colours
# ============================================================

DPI = 300

# Figure size is content-driven rather than manually fixed.
PANEL_A_HEIGHT_IN = 4.35
PANEL_B_ROW_HEIGHT_IN = 0.92
PANEL_B_COLUMN_WIDTH_IN = 0.62
FIGURE_LEFT_LABEL_WIDTH_IN = 2.05
FIGURE_RIGHT_LEGEND_WIDTH_IN = 1.05
FIGURE_VERTICAL_PADDING_IN = 1.85
FIGURE_MIN_WIDTH_IN = 8.2
FIGURE_MIN_HEIGHT_IN = 9.8

# Panel B physical spacing.
PANEL_B_X_SPACING = 0.44
PANEL_B_X_HALF_RANGE = 2.10
PANEL_B_GRID_LINE_PAD = 0.06

TOP_MARGIN = 0.925
BOTTOM_MARGIN = 0.185
LEFT_MARGIN = 0.138
RIGHT_MARGIN = 0.965
PANEL_VERTICAL_GAP = 0.36

FONT_FALLBACKS = ["Arial", "DejaVu Sans"]
FONT_SIZE_PANEL_TITLE = 13.8
FONT_SIZE_PANEL_SUBTITLE = 8.8
FONT_SIZE_CARD_LABEL = 7.7
FONT_SIZE_CARD_VALUE = 10.8
FONT_SIZE_CARD_PERCENT = 8.0
FONT_SIZE_CARD_TAG = 6.7
FONT_SIZE_TICK_LABEL_X = 8.0
FONT_SIZE_TICK_LABEL_Y = 10.3
FONT_SIZE_CELL_COUNT = 9.3
FONT_SIZE_LEGEND = 9.2

BUBBLE_MIN_AREA = 100
BUBBLE_MAX_AREA = 1500
BUBBLE_EDGE_WIDTH = 0.80
BUBBLE_COUNT_THRESHOLD = 1
# Exponent > 1 exaggerates differences in count-based bubble area.
BUBBLE_AREA_EXPONENT = 0.72

CARD_ROUNDING = 0.030
CARD_EDGE_WIDTH = 1.0
CARD_PROGRESS_HEIGHT = 0.032
CARD_LABEL_XPAD = 0.072
CARD_VALUE_XPAD = 0.072

RWTH_BLUE_MUTED = "#A9C4DA"
RWTH_BLUE_LIGHT = "#E7F0F7"
RWTH_GREEN_MUTED = "#DCE8C4"
RWTH_WARM_MUTED = "#DECBAE"
RWTH_CONFLICT_MUTED = "#C7BFD3"
RWTH_GREY_MUTED = "#D7D7D7"
RWTH_GREY_LIGHT = "#F1F1F1"

COLOR_TEXT = "#111827"
COLOR_MUTED_TEXT = "#6B7280"
COLOR_GRID = "#E5E7EB"
COLOR_WHITE = "#FFFFFF"

PANEL_A_ACCENT_COLOR = RWTH_BLUE_MUTED
PANEL_B_BUBBLE_FILL = "#E8F0F6"
PANEL_B_BUBBLE_EDGE = "#88AFC8"

CARD_STYLE_INCLUDED = {"bg": PANEL_B_BUBBLE_FILL, "edge": PANEL_B_BUBBLE_EDGE, "accent": PANEL_B_BUBBLE_EDGE}
CARD_STYLE_EXCLUDED = {"bg": "#F3F4F6", "edge": "#C7CDD4", "accent": "#B9C0C7"}

plt.rcParams.update({
    "font.family": "sans-serif",
    "font.sans-serif": FONT_FALLBACKS,
    "axes.titlesize": FONT_SIZE_PANEL_TITLE,
    "xtick.labelsize": FONT_SIZE_TICK_LABEL_X,
    "ytick.labelsize": FONT_SIZE_TICK_LABEL_Y,
    "legend.fontsize": FONT_SIZE_LEGEND,
    "pdf.fonttype": 42,
    "ps.fonttype": 42,
})


def compute_auto_figsize(n_applications: int, n_status_rows: int) -> tuple[float, float]:
    grid_width = max(n_applications - 1, 1) * PANEL_B_COLUMN_WIDTH_IN
    width = FIGURE_LEFT_LABEL_WIDTH_IN + grid_width + FIGURE_RIGHT_LEGEND_WIDTH_IN
    width = max(width, FIGURE_MIN_WIDTH_IN)

    panel_b_height = max(n_status_rows, 1) * PANEL_B_ROW_HEIGHT_IN
    height = PANEL_A_HEIGHT_IN + panel_b_height + FIGURE_VERTICAL_PADDING_IN
    height = max(height, FIGURE_MIN_HEIGHT_IN)

    return width, height


# ============================================================
# 3. Labels and category order
# ============================================================

CERTIFICATION_BASIS_ORDER = [
    "Generic compostable/biodegradable wording only",
    "No named standard/approval stated",
    "EN 13432 / OK compost / Seedling",
    "Government/programme approval",
    "BPI / ASTM / CMA",
    "OK compost HOME / NF T 51-800 / AS 5810",
    "Country-specific standards/certifications (BNQ / AS 4736 / DINplus)",
]

CERTIFICATION_RELEVANT_BASIS = {
    "EN 13432 / OK compost / Seedling",
    "Government/programme approval",
    "BPI / ASTM / CMA",
    "OK compost HOME / NF T 51-800 / AS 5810",
    "Country-specific standards/certifications (BNQ / AS 4736 / DINplus)",
}

CERTIFICATION_BASIS_RENAME = {
    "No standard stated": "No named standard/approval stated",
    "Generic compostable/biodegradable claim only": "Generic compostable/biodegradable wording only",
    "Government/programme-approved": "Government/programme approval",
    "National standards": "Country-specific standards/certifications (BNQ / AS 4736 / DINplus)",
    "National standards (BNQ / AS 4736 / DINplus etc.)": "Country-specific standards/certifications (BNQ / AS 4736 / DINplus)",
    "BNQ / AS 4736 / DINplus / national standards": "Country-specific standards/certifications (BNQ / AS 4736 / DINplus)",
    "Country-specific standards/certifications (BNQ / AS 4736 / DINplus etc.)": "Country-specific standards/certifications (BNQ / AS 4736 / DINplus)",
}

CERTIFICATION_BASIS_SHORT = {
    "Generic compostable/biodegradable wording only": "Generic compostable /\nbiodegradable\nwording",
    "No named standard/approval stated": "No named standard /\napproval stated",
    "EN 13432 / OK compost / Seedling": "EN 13432 /\nOK compost /\nSeedling",
    "Government/programme approval": "Government /\nprogramme\napproval",
    "BPI / ASTM / CMA": "BPI / ASTM /\nCMA",
    "OK compost HOME / NF T 51-800 / AS 5810": "OK compost HOME /\nNF T 51-800 /\nAS 5810",
    "Country-specific standards/certifications (BNQ / AS 4736 / DINplus)": "Country-specific\nstandards /\ncertifications",
}

APPLICATION_ORDER = [
    "Food-waste liners / collection bags",
    "Food-service ware / takeaway packaging",
    "Tea/coffee preparation items",
    "Food-soiled paper / fibre packaging",
    "Shopping/produce bags",
    "Flexible films/wraps/pouches",
    "Generic compostable packaging / plastics",
]

APPLICATION_LABELS = {
    "Food-waste liners / collection bags": "Food-waste\nliners /\nbags",
    "Food-service ware / takeaway packaging": "Food-service\nware /\ntakeaway",
    "Tea/coffee preparation items": "Tea/coffee\npreparation\nitems",
    "Food-soiled paper / fibre packaging": "Food-soiled\npaper / fibre\npackaging",
    "Shopping/produce bags": "Shopping /\nproduce\nbags",
    "Flexible films/wraps/pouches": "Flexible films /\nwraps /\npouches",
    "Generic compostable packaging / plastics": "Generic\npackaging /\nplastics",
}

STATUS_ORDER = [
    "Certification accepted",
    "Certification conditionally accepted",
    "Certification rejected",
    "Certification unclear or not stated",
]

STATUS_LABELS = {
    "Certification accepted": "Certification\naccepted",
    "Certification conditionally accepted": "Certification\nconditionally\naccepted",
    "Certification rejected": "Certification\nrejected",
    "Certification unclear or not stated": "Certification\nunclear or\nnot stated",
}

# Backward compatibility with older Certification Sufficiency outputs.
STATUS_RENAME = {
    "Named certification/approval supports acceptance": "Certification accepted",
    "Certification/approval required but not sufficient": "Certification conditionally accepted",
    "Rejected despite named certification/approval": "Certification rejected",
    "Acceptance unclear or not stated": "Certification unclear or not stated",
    "Certification sufficient for ordinary acceptance": "Certification accepted",
    "Certification mentioned but operationally rejected": "Certification rejected",
    "Compatibility unclear or not stated": "Certification unclear or not stated",
}


# ============================================================
# 4. Helper functions
# ============================================================

def require_file(path: Path) -> None:
    if not path.exists():
        raise FileNotFoundError(f"Missing required input file: {path}")


def pct(x: float) -> str:
    return f"{x:.1f}%"


def bubble_area(values: np.ndarray, max_value: float | None = None) -> np.ndarray:
    """Scale positive counts to bubble areas.

    Zero-count cells return zero area and are not plotted. Positive counts use
    an exponent > 1 so large-count cells stand out more clearly from small-count
    cells in the revised four-row matrix.
    """
    values = np.asarray(values, dtype=float)
    if values.size == 0:
        return values
    if max_value is None:
        positive = values[values > 0]
        max_value = np.nanmax(positive) if positive.size else 1.0
    max_value = max(float(max_value), 1.0)

    out = np.zeros_like(values, dtype=float)
    positive_mask = values > 0
    scaled = (np.clip(values[positive_mask], 0, None) / max_value) ** BUBBLE_AREA_EXPONENT
    out[positive_mask] = BUBBLE_MIN_AREA + scaled * (BUBBLE_MAX_AREA - BUBBLE_MIN_AREA)
    return out


def draw_badge_card(
    ax,
    x: float,
    y: float,
    w: float,
    h: float,
    category: str,
    label: str,
    n: int,
    denominator: int,
    share_percent: float,
) -> None:
    included = category in CERTIFICATION_RELEVANT_BASIS
    style = CARD_STYLE_INCLUDED if included else CARD_STYLE_EXCLUDED
    accent = style["accent"]

    card = FancyBboxPatch(
        (x, y), w, h,
        boxstyle=f"round,pad=0.012,rounding_size={CARD_ROUNDING}",
        linewidth=CARD_EDGE_WIDTH,
        edgecolor=style["edge"],
        facecolor=style["bg"],
        zorder=2,
    )
    ax.add_patch(card)

    ax.text(
        x + w * CARD_LABEL_XPAD,
        y + h * 0.68,
        label,
        ha="left",
        va="center",
        fontsize=FONT_SIZE_CARD_LABEL,
        fontweight="bold",
        color=COLOR_TEXT,
        linespacing=1.0,
        zorder=6,
    )
    ax.text(
        x + w * CARD_VALUE_XPAD,
        y + h * 0.33,
        f"{n} / {denominator}",
        ha="left",
        va="center",
        fontsize=FONT_SIZE_CARD_VALUE,
        fontweight="bold",
        color=COLOR_TEXT,
        zorder=6,
    )
    ax.text(
        x + w * CARD_VALUE_XPAD,
        y + h * 0.19,
        pct(share_percent),
        ha="left",
        va="center",
        fontsize=FONT_SIZE_CARD_PERCENT,
        color=COLOR_MUTED_TEXT,
        zorder=6,
    )

    track_x = x + w * CARD_VALUE_XPAD
    track_y = y + h * 0.070
    track_w = w * 0.42
    track_h = h * CARD_PROGRESS_HEIGHT
    ax.add_patch(Rectangle((track_x, track_y), track_w, track_h, facecolor="#E5E7EB", edgecolor="none", zorder=3))
    ax.add_patch(Rectangle(
        (track_x, track_y),
        track_w * max(0, min(share_percent / 100.0, 1.0)),
        track_h,
        facecolor=accent,
        edgecolor="none",
        zorder=4,
    ))


def build_bubble_df(display_dir: Path) -> pd.DataFrame:
    """Read the revised long bubble matrix if available; otherwise build it from the wide table."""
    long_path = display_dir / "certification_sufficiency_bubble_matrix_display.csv"
    wide_path = display_dir / "certification_sufficiency_by_application_display.csv"

    if long_path.exists():
        df = pd.read_csv(long_path)
        required = {"application_type", "certification_sufficiency_status", "n"}
        missing = required - set(df.columns)
        if missing:
            raise ValueError(f"{long_path} is missing columns: {sorted(missing)}")
        out = df[list(required)].copy()
        out["certification_sufficiency_status"] = out["certification_sufficiency_status"].replace(STATUS_RENAME)
        out = out[out["certification_sufficiency_status"].isin(STATUS_ORDER)].copy()
        return out.rename(columns={"application_type": "application_group"})

    require_file(wide_path)
    wide = pd.read_csv(wide_path).rename(columns=STATUS_RENAME)
    if "application_type" not in wide.columns:
        raise KeyError(f"Expected column 'application_type' in {wide_path}")

    records = []
    for _, row in wide.iterrows():
        app = str(row["application_type"])
        for status in STATUS_ORDER:
            records.append({
                "application_group": app,
                "certification_sufficiency_status": status,
                "n": int(row.get(status, 0)) if status in wide.columns else 0,
            })
    return pd.DataFrame(records)


# ============================================================
# 5. Load and prepare data
# ============================================================

source_share_path = DISPLAY_DIR / "source_level_category_shares_display.csv"
require_file(source_share_path)

source_shares = pd.read_csv(source_share_path)
source_shares["category"] = source_shares["category"].replace(CERTIFICATION_BASIS_RENAME)

cert_basis = source_shares[source_shares["indicator"].eq("Certification and Approval Basis")].copy()
cert_basis["category"] = pd.Categorical(cert_basis["category"], categories=CERTIFICATION_BASIS_ORDER, ordered=True)
cert_basis = cert_basis.sort_values("category")

present = set(cert_basis["category"].astype(str))
missing = [c for c in CERTIFICATION_BASIS_ORDER if c not in present]
if missing:
    raise ValueError(
        "Missing Certification and Approval Basis categories from source-level share file:\n"
        + "\n".join(missing)
    )

bubble_df = build_bubble_df(DISPLAY_DIR)

# Ensure all application × status combinations exist.
available_apps = [a for a in APPLICATION_ORDER if a in set(bubble_df["application_group"].astype(str))]
if not available_apps:
    available_apps = [str(a) for a in bubble_df["application_group"].dropna().unique().tolist()]

full_index = pd.MultiIndex.from_product(
    [available_apps, STATUS_ORDER],
    names=["application_group", "certification_sufficiency_status"],
)
bubble_df = (
    bubble_df.groupby(["application_group", "certification_sufficiency_status"], as_index=True)["n"]
    .sum()
    .reindex(full_index, fill_value=0)
    .reset_index()
)

bubble_df.to_csv(OUTPUT_MATRIX_CSV, index=False, encoding="utf-8-sig")

max_count = max(int(bubble_df["n"].max()), 1)
cert_relevant_n = int(bubble_df["n"].sum())


# ============================================================
# 6. Plot
# ============================================================

fig_width, fig_height = compute_auto_figsize(len(available_apps), len(STATUS_ORDER))
fig = plt.figure(figsize=(fig_width, fig_height))
gs = fig.add_gridspec(nrows=2, ncols=1, height_ratios=[1.15, 1.30], hspace=PANEL_VERTICAL_GAP)

ax_a = fig.add_subplot(gs[0, 0])
ax_b = fig.add_subplot(gs[1, 0])

# ------------------------------------------------------------
# Panel A
# ------------------------------------------------------------
ax_a.set_xlim(0, 1)
ax_a.set_ylim(0, 1)
ax_a.axis("off")

ax_a.text(
    0.0, 1.135,
    "a. Composting certification and approval basis",
    ha="left", va="bottom",
    fontsize=FONT_SIZE_PANEL_TITLE,
    fontweight="bold",
    color=COLOR_TEXT,
    transform=ax_a.transAxes,
)
ax_a.text(
    0.0, 1.005,
    "Blue tiles feed Panel B; grey tiles are excluded from the Panel B denominator.\n"
    "All 184 sources are shown.",
    ha="left", va="bottom",
    fontsize=FONT_SIZE_PANEL_SUBTITLE,
    color=COLOR_MUTED_TEXT,
    linespacing=1.08,
    transform=ax_a.transAxes,
)

card_w = 0.178
card_h = 0.37
x_positions_top = [0.006, 0.230, 0.454, 0.678]
x_positions_bottom = [0.118, 0.342, 0.566]
y_top = 0.515
y_bottom = 0.065

for i, (_, row) in enumerate(cert_basis.iterrows()):
    category = str(row["category"])
    label = CERTIFICATION_BASIS_SHORT.get(category, category)
    n = int(row["n"])
    denominator = int(row["denominator"])
    share_percent = float(row["share_percent"])

    x = x_positions_top[i] if i < 4 else x_positions_bottom[i - 4]
    y = y_top if i < 4 else y_bottom

    draw_badge_card(
        ax=ax_a,
        x=x, y=y, w=card_w, h=card_h,
        category=category, label=label,
        n=n, denominator=denominator, share_percent=share_percent,
    )

# ------------------------------------------------------------
# Panel B
# ------------------------------------------------------------
ax_b.text(
    0.0, 1.175,
    "b. Composting certification sufficiency",
    ha="left", va="bottom",
    fontsize=FONT_SIZE_PANEL_TITLE,
    fontweight="bold",
    color=COLOR_TEXT,
    transform=ax_b.transAxes,
)
ax_b.text(
    0.0, 1.025,
    f"By application group; bubble area shows certification-relevant source × application-group records (n = {cert_relevant_n}).\n"
    "Zero-count cells are not plotted.",
    ha="left", va="bottom",
    fontsize=FONT_SIZE_PANEL_SUBTITLE,
    color=COLOR_MUTED_TEXT,
    linespacing=1.10,
    transform=ax_b.transAxes,
)

x_positions_raw = np.arange(len(available_apps), dtype=float) * PANEL_B_X_SPACING
x_positions_centered = x_positions_raw - x_positions_raw.mean()
x_index = {app: x for app, x in zip(available_apps, x_positions_centered)}
y_index = {status: i for i, status in enumerate(STATUS_ORDER)}

x_positions_b = [x_index[a] for a in available_apps]
for x in x_positions_b:
    ax_b.axvline(x, color=COLOR_GRID, linewidth=0.8, zorder=0)
for y in range(len(STATUS_ORDER)):
    ax_b.hlines(
        y,
        x_positions_b[0] - PANEL_B_GRID_LINE_PAD,
        x_positions_b[-1] + PANEL_B_GRID_LINE_PAD,
        color=COLOR_GRID,
        linewidth=0.55,
        zorder=0,
    )

for status in STATUS_ORDER:
    sub = bubble_df[
        bubble_df["certification_sufficiency_status"].eq(status)
        & bubble_df["n"].gt(0)
    ].copy()
    if sub.empty:
        continue

    xs = [x_index[a] for a in sub["application_group"]]
    ys = [y_index[status]] * len(sub)
    counts = sub["n"].to_numpy()
    sizes = bubble_area(counts, max_value=max_count)

    ax_b.scatter(
        xs, ys,
        s=sizes,
        facecolors=PANEL_B_BUBBLE_FILL,
        edgecolors=PANEL_B_BUBBLE_EDGE,
        linewidths=BUBBLE_EDGE_WIDTH,
        alpha=1.0,
        zorder=3,
    )

    for x, y, count, size in zip(xs, ys, counts, sizes):
        if count >= BUBBLE_COUNT_THRESHOLD:
            ax_b.text(
                x, y, str(int(count)),
                ha="center", va="center",
                fontsize=FONT_SIZE_CELL_COUNT,
                color=COLOR_TEXT,
                zorder=4,
            )

ax_b.set_xlim(-PANEL_B_X_HALF_RANGE, PANEL_B_X_HALF_RANGE)
ax_b.set_ylim(len(STATUS_ORDER) - 0.6, -0.6)

ax_b.set_xticks(x_positions_b)
x_tick_texts = ax_b.set_xticklabels(
    [APPLICATION_LABELS.get(a, "\n".join(textwrap.wrap(a, 14))) for a in available_apps],
    ha="center",
    va="top",
    fontsize=FONT_SIZE_TICK_LABEL_X,
)
for tick in x_tick_texts:
    tick.set_linespacing(0.92)

ax_b.set_yticks(range(len(STATUS_ORDER)))
y_tick_texts = ax_b.set_yticklabels(
    [STATUS_LABELS.get(s, s) for s in STATUS_ORDER],
    fontsize=FONT_SIZE_TICK_LABEL_Y,
    color=COLOR_TEXT,
)
for tick in y_tick_texts:
    tick.set_linespacing(0.92)

ax_b.tick_params(axis="x", pad=8, length=0)
ax_b.tick_params(axis="y", pad=-10, length=0)

for spine in ax_b.spines.values():
    spine.set_visible(False)

legend_counts = [5, 20, max_count]
legend_counts = sorted(set([v for v in legend_counts if 0 < v <= max_count]))
size_handles = [
    Line2D(
        [0], [0],
        marker="o",
        linestyle="None",
        color="none",
        markerfacecolor=PANEL_B_BUBBLE_FILL,
        markeredgecolor=PANEL_B_BUBBLE_EDGE,
        markersize=np.sqrt(bubble_area(np.array([v]), max_value=max_count)[0]) / 1.32,
        label=str(v),
    )
    for v in legend_counts
]

ax_b.legend(
    handles=size_handles,
    title="Bubble area\n(records)",
    loc="upper right",
    bbox_to_anchor=(1.09, 1.235),
    frameon=False,
    fancybox=False,
    framealpha=0.0,
    ncol=len(size_handles),
    handletextpad=0.55,
    columnspacing=0.80,
    fontsize=FONT_SIZE_LEGEND,
    title_fontsize=FONT_SIZE_LEGEND,
    borderpad=0.25,
)

fig.subplots_adjust(
    left=LEFT_MARGIN,
    right=RIGHT_MARGIN,
    top=TOP_MARGIN,
    bottom=BOTTOM_MARGIN,
)

fig.savefig(OUTPUT_PNG, dpi=DPI, bbox_inches="tight", facecolor="white", edgecolor="white", transparent=False)
fig.savefig(OUTPUT_PDF, bbox_inches="tight", facecolor="white", edgecolor="white", transparent=False)
plt.close(fig)

print(f"Input display folder: {DISPLAY_DIR}")
print(f"Certification-relevant records in Panel B: {cert_relevant_n}")
print(f"Saved: {OUTPUT_PNG}")
print(f"Saved: {OUTPUT_PDF}")
print(f"Saved: {OUTPUT_MATRIX_CSV}")
