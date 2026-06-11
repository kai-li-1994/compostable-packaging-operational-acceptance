#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Figure 1 - Analytical workflow schematic.

Design logic (mirrors the organization of the study):
  EVIDENCE BASE  ->  THREE ANALYTICAL ROUTES  ->  ANALYTICAL OUTPUTS
  - Left   : corpus construction + the two coding levels
  - Middle : Route 1 / Route 2 / Route 3 as three parallel indicator lanes
  - Right  : one output box per route, mapped to the research questions

Coding-level encoding:
  SOLID  border  = source level             (n = 184)
  DASHED border  = source x application grp  (n = 295)
  FILLED chip    = derived analytical output

Palette is taken directly from the paper's other figures
(figure_source_level.py / figure_application_country.py /
 figure_certification_sufficiency.py): muted, low-saturation RWTH tones.

Output: editable vector PDF + SVG (text kept as text, Type-42 fonts) for
Adobe Illustrator, plus a 300-dpi PNG preview. Numbers live in COUNTS.

Output base name matches this script name: figure_1_schematic.
"""

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from pathlib import Path
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

Path("./figures").mkdir(parents=True, exist_ok=True)

# ----------------------------------------------------------------------
# Editable-text settings, matched to the paper's other figures
# ----------------------------------------------------------------------
plt.rcParams["pdf.fonttype"]    = 42
plt.rcParams["ps.fonttype"]     = 42
plt.rcParams["svg.fonttype"]    = "none"
plt.rcParams["font.family"]     = "sans-serif"
plt.rcParams["font.sans-serif"] = ["Arial", "DejaVu Sans"]

# ----------------------------------------------------------------------
# Headline counts (from the coded corpus)
# ----------------------------------------------------------------------
COUNTS = dict(
    countries=36, sources=184, records=295,
    cert_sources=75, cert_records=158,
    national=25, below_national=159,
)

# ----------------------------------------------------------------------
# Palette - copied from the paper's figures (muted RWTH tones)
# ----------------------------------------------------------------------
INK    = "#111827"   # COLOR_TEXT
SUBINK = "#6B7280"   # COLOR_MUTED_TEXT
SEP    = "#334155"   # COLOR_PANEL_BORDER / separator
ARROW  = "#94A3B8"   # spine grey
GRID   = "#CBD5E1"
PANEL_BG = "#F1F1F1"  # RWTH_GREY_LIGHT
GREY_MUTED = "#D7D7D7"

# route families: line = muted edge, fill = paper pastel, soft = lane tint
R1 = dict(line="#88AFC8", fill="#9DBBD5", soft="#EAF1F7")   # blue  (source authority)
R2 = dict(line="#A7C07F", fill="#DCE8C4", soft="#F1F6E9")   # green (compatibility)
R3 = dict(line="#A99FC0", fill="#C7BFD3", soft="#F0ECF4")   # plum  (certification/conflict)

# ----------------------------------------------------------------------
# Canvas
# ----------------------------------------------------------------------
W, H = 200.0, 106.0
fig = plt.figure(figsize=(10.0, 10.0 * H / W))
ax = fig.add_axes([0, 0, 1, 1])
ax.set_xlim(0, W); ax.set_ylim(0, H); ax.axis("off")

# ----------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------
def rbox(x, y, w, h, fc="white", ec=INK, lw=1.0, ls="solid", r=1.6, z=2):
    p = FancyBboxPatch((x, y), w, h,
                       boxstyle=f"round,pad=0,rounding_size={r}",
                       fc=fc, ec=ec, lw=lw, linestyle=ls, zorder=z,
                       mutation_aspect=1.0)
    ax.add_patch(p)

def txt(x, y, s, size=7.0, color=INK, weight="normal", ha="center", va="center",
        style="normal", z=5, lh=1.12):
    ax.text(x, y, s, fontsize=size, color=color, fontweight=weight, ha=ha, va=va,
            fontstyle=style, zorder=z, linespacing=lh)

def arrow(p1, p2, color=ARROW, lw=1.2, z=1):
    a = FancyArrowPatch(p1, p2, arrowstyle="-|>", mutation_scale=9,
                        lw=lw, color=color, zorder=z, shrinkA=1, shrinkB=1)
    ax.add_patch(a)

# ======================================================================
# LEGEND (top strip)
# ======================================================================
ly = H - 5.0
lx = 52
def legend_item(x, kind, label):
    if kind == "solid":
        rbox(x, ly - 2.2, 5.0, 4.4, fc="white", ec=SEP, lw=1.1, r=0.8)
    elif kind == "dashed":
        rbox(x, ly - 2.2, 5.0, 4.4, fc="white", ec=SEP, lw=1.1, ls=(0, (3, 2)), r=0.8)
    else:
        rbox(x, ly - 2.2, 5.0, 4.4, fc=GREY_MUTED, ec=SUBINK, lw=1.0, r=0.8)
    txt(x + 6.4, ly, label, size=6.6, color=SUBINK, ha="left")

txt(lx - 1, ly, "Coding level:", size=6.8, color=INK, weight="bold", ha="right")
legend_item(lx + 1,   "solid",  f"source level  (n = {COUNTS['sources']})")
legend_item(lx + 58,  "dashed", f"source × application group  (n = {COUNTS['records']})")
legend_item(lx + 132, "output", "derived analytical output")

# ======================================================================
# LANE GEOMETRY
# ======================================================================
LANE_X0, LANE_X1 = 51.0, 150.0
LANE_H = 16.0
cy = {"R1": 79.0, "R2": 51.0, "R3": 23.0}
hdr_dy = LANE_H / 2 + 3.4
RES_X0, RES_X1 = 154.0, 198.0

# ======================================================================
# 1. EVIDENCE BASE (left column)
# ======================================================================
EB_X, EB_W = 2.0, 44.0
EB_Y0, EB_Y1 = cy["R3"] - LANE_H / 2 - 1, cy["R1"] + LANE_H / 2 + 1
rbox(EB_X, EB_Y0, EB_W, EB_Y1 - EB_Y0, fc=PANEL_BG, ec=GRID, lw=1.0, r=2.0, z=1)
txt(EB_X + EB_W / 2, EB_Y1 - 3.6, "EVIDENCE BASE", size=8.2, weight="bold", color=INK)

# Card A: corpus construction
ca_y, ca_h = EB_Y1 - 39.0, 31.0
rbox(EB_X + 2.5, ca_y, EB_W - 5.0, ca_h, fc="white", ec=GRID, lw=1.0, r=1.4)
txt(EB_X + EB_W / 2, ca_y + ca_h - 3.4, "Corpus construction", size=7.2, weight="bold", color=INK)
txt(EB_X + EB_W / 2, ca_y + ca_h - 9.6,
    f"{COUNTS['countries']} OECD countries\n{COUNTS['sources']} grey-literature sources",
    size=7.0, color=INK, weight="bold")
txt(EB_X + 3.6, ca_y + ca_h - 16.6,
    "Discovery:  multilingual web ·\ntargeted official sites · LLM-assisted",
    size=6.2, color=SUBINK, ha="left")
txt(EB_X + 3.6, ca_y + 5.0,
    "Inclusion (3 criteria):  public ·\nauthority type · operational rule",
    size=6.2, color=SUBINK, ha="left")

# Card B: coding levels
cb_y, cb_h = EB_Y0 + 3.0, 24.0
rbox(EB_X + 2.5, cb_y, EB_W - 5.0, cb_h, fc="white", ec=GRID, lw=1.0, r=1.4)
txt(EB_X + EB_W / 2, cb_y + cb_h - 3.4, "Coding framework", size=7.2, weight="bold", color=INK)
txt(EB_X + EB_W / 2, cb_y + cb_h - 7.6, "regex-based Python coding",
    size=6.0, color=SUBINK, style="italic")
rbox(EB_X + 4.0, cb_y + 7.0, EB_W - 8.0, 5.0, fc="white", ec=SEP, lw=1.1, r=0.8)
txt(EB_X + EB_W / 2, cb_y + 9.5, f"Source level   n = {COUNTS['sources']}", size=6.6, color=INK)
rbox(EB_X + 4.0, cb_y + 1.2, EB_W - 8.0, 5.0, fc="white", ec=SEP, lw=1.1, ls=(0, (3, 2)), r=0.8)
txt(EB_X + EB_W / 2, cb_y + 3.7, f"Source × application   n = {COUNTS['records']}", size=6.6, color=INK)

for k in cy:
    arrow((EB_X + EB_W + 0.5, cy[k]), (LANE_X0 - 0.5, cy[k]), lw=1.4)

# ======================================================================
# 2. THREE ANALYTICAL ROUTES (middle)
# ======================================================================
def lane_bg(key, col):
    y0 = cy[key] - LANE_H / 2
    rbox(LANE_X0 - 1.5, y0, (LANE_X1 - LANE_X0) + 3.0, LANE_H,
         fc=col["soft"], ec=col["line"], lw=1.0, r=2.0, z=1)

def chips(key, col, items, gap=4.0):
    n = len(items)
    cw = ((LANE_X1 - LANE_X0) - gap * (n - 1)) / n
    centers = []
    for i, it in enumerate(items):
        x = LANE_X0 + i * (cw + gap)
        if it.get("output"):
            fc, ec, tc = col["fill"], col["line"], INK
        else:
            fc, ec, tc = "white", col["line"], INK
        ls = (0, (3, 2)) if it["level"] == "SxA" else "solid"
        ch = 11.5
        rbox(x, cy[key] - ch / 2, cw, ch, fc=fc, ec=ec, lw=1.1, ls=ls, r=1.3, z=3)
        txt(x + cw / 2, cy[key] + (1.0 if it.get("note") else 0.0),
            it["label"], size=7.0, color=tc, weight="bold")
        if it.get("note"):
            txt(x + cw / 2, cy[key] - 3.2, it["note"], size=5.7, color=SUBINK)
        centers.append((x, x + cw))
    for i in range(n - 1):
        arrow((centers[i][1] + 0.3, cy[key]), (centers[i + 1][0] - 0.3, cy[key]),
              color=col["line"], lw=1.2)
    return centers

def lane_header(key, col, text):
    sx = LANE_X0 - 1.5
    rbox(sx, cy[key] + hdr_dy - 1.6, 3.0, 3.0, fc=col["fill"], ec=col["line"], lw=0.8, r=0.5)
    txt(sx + 4.4, cy[key] + hdr_dy, text, size=7.3, weight="bold", color=INK, ha="left")

lane_bg("R1", R1)
lane_header("R1", R1, "ROUTE 1 · Source-level operational rules")
chips("R1", R1, [
    dict(label="Source\nAuthority", level="S"),
    dict(label="Stated Treatment\nRoute", level="S"),
    dict(label="Acceptance", level="S"),
    dict(label="Rejection\nRationale", level="S"),
])

lane_bg("R2", R2)
lane_header("R2", R2, "ROUTE 2 · Application-level compatibility")
chips("R2", R2, [
    dict(label="Application\nGroup", level="SxA", note="7 groups"),
    dict(label="Application\nDecision", level="SxA"),
    dict(label="Compatibility\nscenario", level="SxA", output=True, note="lower · central · upper"),
])

lane_bg("R3", R3)
lane_header("R3", R3, "ROUTE 3 · Certification sufficiency")
chips("R3", R3, [
    dict(label="Certification &\nApproval Basis", level="S",
         note=f"{COUNTS['cert_sources']} certification-relevant sources"),
    dict(label="Certification\nSufficiency", level="SxA", output=True,
         note=f"{COUNTS['cert_records']} records"),
])

# ======================================================================
# 3. ANALYTICAL OUTPUTS (right column) - one box per route
# ======================================================================
res = {
    "R1": ("RQ1 · Who governs acceptance?", "Authority level of\nacceptance rules", R1),
    "R2": ("RQ2 · Which applications are compatible?", "Function-specific\ncompatibility profile", R2),
    "R3": ("RQ3 · Is certification sufficient?", "Certification–acceptance\ngap", R3),
}
RES_W = RES_X1 - RES_X0
for key, (rq, finding, col) in res.items():
    y0 = cy[key] - LANE_H / 2
    arrow((LANE_X1 + 1.8, cy[key]), (RES_X0 - 0.5, cy[key]), color=col["line"], lw=1.5)
    rbox(RES_X0, y0, RES_W, LANE_H, fc=col["soft"], ec=col["line"], lw=1.1, r=1.8, z=3)
    txt(RES_X0 + RES_W / 2, cy[key] + 3.4, rq, size=6.9, weight="bold", color=INK)
    txt(RES_X0 + RES_W / 2, cy[key] - 2.4, finding, size=6.6, color=SUBINK)

txt((RES_X0 + RES_X1) / 2, cy["R1"] + hdr_dy, "ANALYTICAL OUTPUTS",
    size=8.2, weight="bold", color=INK)

# ======================================================================
# 4. SENSITIVITY strip (bottom)
# ======================================================================
sy = cy["R3"] - LANE_H / 2 - 8.5
rbox(LANE_X0 - 1.5, sy, (RES_X1 - LANE_X0) + 1.5, 5.6,
     fc="white", ec=SUBINK, lw=1.0, ls=(0, (2, 2)), r=1.2, z=2)
txt((LANE_X0 + RES_X1) / 2 - 0.7, sy + 2.8,
    f"Sensitivity check:  national sources (n = {COUNTS['national']})   "
    f"vs   sources below national level (n = {COUNTS['below_national']})",
    size=6.6, color=SUBINK, weight="bold")
arrow(((LANE_X0 + LANE_X1) / 2, cy["R3"] - LANE_H / 2 - 0.5),
      ((LANE_X0 + LANE_X1) / 2, sy + 5.9), color=SUBINK, lw=1.0)

# ======================================================================
# Export
# ======================================================================
for ext in ("pdf", "svg", "png"):
    fig.savefig(f"./figures/figure_1_schematic.{ext}",
                dpi=300, bbox_inches="tight", pad_inches=0.04)
print("done")
