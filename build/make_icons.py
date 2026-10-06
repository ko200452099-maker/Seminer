#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Line icons for the asphaltene deck — one stroke weight, palette colours, transparent PNG."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Circle, Polygon, Rectangle, FancyBboxPatch, Wedge
import numpy as np
import os

NAVY = "#1B365D"; TEAL = "#0D7377"; ORANGE = "#E07A3D"
OUT = "/home/user/Seminar_Asphaltenes/_build/figs"
os.makedirs(OUT, exist_ok=True)
LW = 7.5          # single stroke weight for every icon
S = 512           # canvas

def new_ax():
    fig, ax = plt.subplots(figsize=(S/100, S/100), dpi=100)
    ax.set_xlim(0, 100); ax.set_ylim(0, 100); ax.axis("off")
    ax.set_aspect("equal")
    return fig, ax

def save(fig, name):
    fig.tight_layout(pad=0)
    fig.savefig(os.path.join(OUT, name), transparent=True, dpi=150)
    plt.close(fig)
    print("icon:", name)

# ---- 1. precipitation flask (Erlenmeyer + droplets) ----
fig, ax = new_ax()
ax.plot([42, 42, 46], [90, 56, 44], color=TEAL, lw=LW, solid_capstyle="round")
ax.plot([58, 58, 54], [90, 56, 44], color=TEAL, lw=LW, solid_capstyle="round")
ax.plot([46, 54], [44, 44], color=TEAL, lw=LW, solid_capstyle="round")
ax.plot([36, 64], [90, 90], color=TEAL, lw=LW, solid_capstyle="round")
for (x, y, r) in ((48, 58, 2.4), (53, 62, 1.8), (51, 52, 3.0)):
    ax.add_patch(Circle((x, y), r, facecolor=ORANGE, edgecolor="none", zorder=4))
ax.add_patch(Polygon([[44, 50], [56, 50], [54, 44.5], [46, 44.5]], closed=True,
                     facecolor="none", edgecolor=TEAL, lw=LW*0.7, zorder=3))
save(fig, "icon_flask.png")

# ---- 2. flocculation clusters (three cores clumping) ----
fig, ax = new_ax()
def core(cx, cy, r):
    a = np.linspace(0, 2*np.pi, 9)[:-1]
    rr = r * (0.82 + 0.3*np.random.default_rng(int(cx*7+cy)).random(len(a)))
    ax.add_patch(Polygon(np.column_stack([cx + rr*np.cos(a), cy + rr*np.sin(a)]),
                         closed=True, facecolor="none", edgecolor=NAVY, lw=LW, zorder=3))
core(44, 56, 13); core(60, 62, 10); core(58, 44, 9)
for (x, y) in ((30, 62), (30, 44), (74, 52), (66, 74), (38, 34)):
    ax.add_patch(Circle((x, y), 2.6, facecolor=TEAL, edgecolor="none", zorder=4))
save(fig, "icon_clusters.png")

# ---- 3. well schematic (well + tubing + reservoir layer) ----
fig, ax = new_ax()
ax.add_patch(Rectangle((22, 22), 56, 10, facecolor="none", edgecolor=ORANGE, lw=LW*0.75, zorder=2))
ax.plot([50, 50], [92, 30], color=NAVY, lw=LW, solid_capstyle="round")
ax.plot([44, 44], [88, 32], color=TEAL, lw=LW*0.8, solid_capstyle="round")
ax.plot([22, 78], [92, 92], color=NAVY, lw=LW*0.75, solid_capstyle="round")
ax.add_patch(Circle((50, 30), 3.4, facecolor=ORANGE, edgecolor="none", zorder=5))
save(fig, "icon_well.png")

# ---- 4. injection mandrel (tubing + side-pocket + capillary line) ----
fig, ax = new_ax()
ax.plot([38, 38], [92, 14], color=NAVY, lw=LW, solid_capstyle="round")
ax.plot([58, 58], [92, 14], color=NAVY, lw=LW, solid_capstyle="round")
ax.add_patch(FancyBboxPatch((43, 34), 12, 20, boxstyle="round,pad=0.6,rounding_size=3",
                            facecolor="none", edgecolor=TEAL, lw=LW, zorder=4))
ax.plot([34, 34], [92, 44], color=ORANGE, lw=LW*0.85, solid_capstyle="round")   # capillary
ax.plot([34, 43], [44, 44], color=ORANGE, lw=LW*0.85, solid_capstyle="round")
for y in (30, 24, 18):
    ax.add_patch(Circle((48, y), 2.2, facecolor=ORANGE, edgecolor="none", zorder=5))
save(fig, "icon_mandrel.png")

# ---- 5. money / breakeven ----
fig, ax = new_ax()
ax.add_patch(Rectangle((18, 22), 64, 46, facecolor="none", edgecolor=NAVY, lw=LW, zorder=2))
ax.plot([28, 28, 72], [30, 30, 30], color=NAVY, lw=LW*0.7, zorder=3)
xs = np.array([30, 42, 54, 66]); ys = np.array([36, 44, 57, 62])
ax.plot(xs, ys, color=TEAL, lw=LW*0.85, zorder=4, solid_capstyle="round")
ax.plot([30, 66], [62, 62], color=ORANGE, lw=LW*0.7, ls=(0, (4, 3)), zorder=4)
ax.add_patch(Circle((66, 62), 3.2, facecolor=ORANGE, edgecolor="none", zorder=5))
ax.text(50, 80, "$", color=NAVY, fontsize=30, ha="center", va="center", fontweight="bold")
save(fig, "icon_money.png")
print("all icons built")
