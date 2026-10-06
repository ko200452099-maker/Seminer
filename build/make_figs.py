#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Publication-quality figures for the asphaltene deck — v4
   Exports: 300 dpi PNG (for the PPTX) + SVG (vector, for PowerPoint insert / thesis).

   Figures
     1. envelope_pt.png        slide 10 — P-T stability envelope
     2. depth_prediction.png   slide 13 — flowing pressure vs AOP(d), onset depth
     3. pipe_views.png         slide 4  — 3 panels: clean bore | deposit | gauge-ring profile
     4. core_shell.png         slide 3  — colloidal stability schematic (3 states)
     5. well_schematic.png     slide 16 — capillary string, mandrel, predicted onset depth
"""
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Circle, Polygon, Rectangle, FancyBboxPatch, Wedge
import os

# ---------------- tokens (locked design system) ----------------
NAVY  = "#1B365D"; TEAL = "#0D7377"; ORANGE = "#E07A3D"; ORANGE_D = "#B35A18"
RED   = "#B43B2E"; BLUE = "#3E6FA3"; GRAY = "#4A5568"; AMBER = "#E07A3D"

plt.rcParams.update({
    "font.family": "DejaVu Sans", "font.size": 11,
    "axes.edgecolor": GRAY, "axes.labelcolor": NAVY, "text.color": NAVY,
    "xtick.color": GRAY, "ytick.color": GRAY, "axes.linewidth": 0.9,
    "figure.facecolor": "white", "svg.fonttype": "path",
})
FIGS = "/home/user/Seminar_Asphaltenes/_build/figs"
os.makedirs(FIGS, exist_ok=True)

def export(fig, name):
    """PNG at 300 dpi for the deck + SVG (vector) for PowerPoint/thesis."""
    fig.savefig(os.path.join(FIGS, name + ".png"), dpi=300)
    fig.savefig(os.path.join(FIGS, name + ".svg"))
    plt.close(fig)
    print("exported:", name, "(png 300 dpi + svg)")

# ---------------- shared P-T model (consistent deck-wide) ----------------
AOP = lambda T: 1080 - 0.9 * (T - 60)          # asphaltene onset pressure, bar
PB  = lambda T: 640 + 2.2 * (T - 60)           # bubble point (illustrative)
X = 930 / (1000/85.0 + 0.9)
T_X, P_X = 60 + X, AOP(60 + X)                 # onset: 133 °C / 1,014 bar
D_X = 930 / ((1000/5500.0) + (0.9*85/5500.0))  # onset depth: 4,752 m

SRC_ENV = ("Illustrative envelope reconstructed from the published deepwater Gulf of Mexico\n"
           "case data (Wylde & Punase, JPT, 30 Apr 2020); P-T path illustrative.")
SRC_DEP = ("Illustrative P-T model. Geometry — perforations 17,700-18,000 ft, mandrel 14,600 ft — from the\n"
           "documented deepwater Gulf of Mexico case (Wylde & Punase, JPT, 30 Apr 2020).")
SRC_PIPE = ("Cross-sections to scale from the measured tubing-ID reduction 2.30 to 1.35 in (Wylde & Punase, JPT, 2020); "
            "gauge-ring profile schematic, after the Hassi Messaoud field practice (Haskett & Tartera, SPE-994-PA, 1965).")
SRC_WELL = ("Hardware schematic, not to scale, after the documented case: mandrel 14,600 ft (4,450 m); perforations "
            "17,700-18,000 ft. Onset depth from the illustrative model on slide 13 (Wylde & Punase, JPT, 2020).")

# =========================================================
# 1 — STABILITY ENVELOPE  (slide 10)
# =========================================================
fig, ax = plt.subplots(figsize=(8.4, 4.6), dpi=300)
T = np.linspace(60, 145, 200)
ax.fill_between(T, 0, AOP(T), color=RED, alpha=0.07, zorder=1)
ax.plot(T, AOP(T), color=ORANGE, lw=3.2, ls="--", zorder=4, label="AOP  asphaltene onset pressure")
ax.plot(T, PB(T), color=BLUE, lw=2.0, zorder=4, label="Bubble point  Pb")
patht = np.linspace(60, 145, 200)
ax.plot(patht, 150 + 1000*(patht-60)/85.0, color=NAVY, lw=3.4, zorder=5,
        label="Flowing path  reservoir \u2192 wellhead")
ax.annotate("", xy=(60.8, 159), xytext=(72, 296),
            arrowprops=dict(arrowstyle="-|>", color=NAVY, lw=3.4, mutation_scale=22), zorder=6)
ax.scatter([145], [1150], s=80, color=TEAL, zorder=7)
ax.annotate("RESERVOIR  145 \u00b0C / 1,150 bar\nstable:  P > AOP",
            xy=(144, 1160), xytext=(142, 1330), fontsize=10.5, color=TEAL,
            ha="right", va="center", arrowprops=dict(arrowstyle="-", color=TEAL, lw=0.9))
ax.scatter([60], [150], s=80, color=NAVY, zorder=7)
ax.annotate("WELLHEAD  60 \u00b0C / 150 bar", xy=(61, 155), xytext=(66, 88),
            fontsize=10.5, color=NAVY, ha="left", va="center",
            arrowprops=dict(arrowstyle="-", color=NAVY, lw=0.9))
ax.scatter([T_X], [P_X], s=230, marker="*", color=ORANGE, edgecolor=NAVY, linewidth=1.4, zorder=8)
ax.annotate(f"ONSET\npath crosses AOP\n\u2248 {P_X:,.0f} bar @ {T_X:.0f} \u00b0C",
            xy=(T_X, P_X), xytext=(134, 620), fontsize=11, color=RED, fontweight="bold",
            ha="center", va="center", arrowprops=dict(arrowstyle="->", color=RED, lw=1.3))
ax.text(63, 950, "THERMODYNAMIC INSTABILITY ZONE\nP < AOP: asphaltenes can precipitate",
        fontsize=10.5, color=ORANGE_D, va="center", ha="left")
ax.text(120, 430, "P and T both fall\nas oil rises", fontsize=10, color=NAVY,
        alpha=0.75, ha="center", va="center")
ax.set_xlim(40, 160); ax.set_ylim(0, 1400)
ax.set_xlabel("Temperature (\u00b0C)"); ax.set_ylabel("Pressure (bar)")
ax.set_title("Asphaltene stability envelope — the oil is safe only above the orange line",
             fontsize=13, fontweight="bold", color=NAVY, pad=8)
ax.legend(loc="lower right", fontsize=10, frameon=True, framealpha=0.95)
ax.grid(False)
fig.text(0.012, 0.055, SRC_ENV, fontsize=10.5, color=GRAY, linespacing=1.35)
fig.tight_layout(rect=[0, 0.105, 1, 1])
export(fig, "envelope_pt")

# =========================================================
# 2 — DEPTH / PRESSURE, ONSET DEPTH  (slide 13)
# =========================================================
D = np.linspace(0, 5500, 400)
Td = 60 + 85 * D / 5500.0
Pd = 150 + 1000 * D / 5500.0
AOPd = AOP(Td)
P_STAR = 150 + 1000 * D_X / 5500.0          # flowing pressure at the onset depth
P_MAND, D_MAND = 150 + 1000 * 4450 / 5500.0, 4450.0
P_TGT = 150 + 1000 * 5000 / 5500.0

fig, ax = plt.subplots(figsize=(8.4, 4.6), dpi=300)
ax.axhspan(4450, 5490, color=AMBER, alpha=0.11, zorder=1)
ax.axhspan(0, D_X, color=RED, alpha=0.055, zorder=0)
ax.plot(Pd, D, color=NAVY, lw=3.4, zorder=5)
ax.plot(AOPd, D, color=ORANGE, lw=3.0, ls="--", zorder=5)
# --- direct curve labels (no legend box to collide with the curves) ---
ax.text(640, 780, "Flowing pressure  P(d)", fontsize=10.5, color=NAVY,
        fontweight="bold", ha="center", va="center")
ax.text(1005, 1300, "AOP(d) — onset pressure", fontsize=10.5, color=ORANGE_D,
        fontweight="bold", ha="right", va="center")
# --- onset ---
ax.scatter([P_STAR], [D_X], s=230, marker="*", color=ORANGE,
           edgecolor=NAVY, linewidth=1.4, zorder=8)
ax.axhline(D_X, color=ORANGE, lw=1.1, ls=":", alpha=0.8)
ax.annotate(f"PREDICTED\nONSET DEPTH\n\u2248 {D_X:,.0f} m",
            xy=(P_STAR - 8, D_X - 120), xytext=(830, 2250), fontsize=11,
            color=RED, fontweight="bold", ha="center", va="center", zorder=9,
            arrowprops=dict(arrowstyle="->", color=RED, lw=1.2))
# --- mandrel line ---
ax.axhline(D_MAND, color=TEAL, lw=1.3, ls="--", alpha=0.85, zorder=3)
ax.text(880, 4250, "existing mandrel  4,450 m (14,600 ft)", fontsize=10, color=TEAL,
        ha="right", va="center", fontweight="bold")
# --- unprotected interval label ---
ax.text(30, 5090, "3,400 ft UNPROTECTED interval", fontsize=10.5,
        color=ORANGE_D, va="top", ha="left")
# --- target injection ---
ax.scatter([P_TGT], [5000], s=110, marker="v", color=TEAL, zorder=6)
ax.text(985, 4950, "target injection \u2248 5,000 m", fontsize=10.5, color=TEAL,
        ha="right", va="center", fontweight="bold")
ax2 = ax.twiny()
ax2.plot(Td, D, color=AMBER, lw=1.2, alpha=0.8)
ax2.set_xlim(50, 160); ax2.set_xlabel("Temperature (\u00b0C)", color=ORANGE_D, fontsize=10)
ax2.tick_params(axis="x", colors=ORANGE_D, labelsize=9.5)
ax.set_ylim(5500, 0); ax.set_xlim(0, 1300)
ax.set_xlabel("Pressure (bar)"); ax.set_ylabel("Depth (m)")
ax.set_title("Where it deposits — the number that sets the injection point",
             fontsize=13, fontweight="bold", color=NAVY, pad=26)
ax.grid(axis="y", alpha=0.12, lw=0.6)
ax.grid(axis="x", visible=False)
fig.text(0.012, 0.055, SRC_DEP, fontsize=10.5, color=GRAY, linespacing=1.35)
fig.tight_layout(rect=[0, 0.105, 1, 1])
export(fig, "depth_prediction")

# =========================================================
# 3 — TUBING: 3 PANELS (slide 4)  — cross-sections + gauge-ring profile
# =========================================================
fig = plt.figure(figsize=(9.6, 4.3), dpi=300)
gs = fig.add_gridspec(1, 3, width_ratios=[1, 1, 1.5], wspace=0.30,
                      left=0.035, right=0.985, top=0.83, bottom=0.185)
axA, axB, axC = (fig.add_subplot(gs[0, i]) for i in range(3))

R_OUT, R_IN, R_DEP = 1.15, 0.944, 0.554   # OD 2.80, ID 2.30, eff. ID 1.35 in (to scale)

def draw_pipe(ax, deposit=None):
    ax.add_patch(Wedge((0, 0), R_OUT, 0, 360, width=R_OUT - R_IN,
                       facecolor="#E4ECF2", edgecolor=GRAY, lw=1.2, hatch="///", zorder=3))
    if deposit is not None:
        ax.add_patch(Wedge((0, 0), R_IN, 0, 360, width=R_IN - deposit,
                           facecolor="#2B2118", edgecolor="#120C07", lw=1.0, zorder=4))
        for rr in np.linspace(deposit + 0.06, R_IN - 0.06, 4):
            ax.add_patch(Circle((0, 0), rr, fill=False, edgecolor="#4A3A2A", lw=0.6, zorder=5))
        ax.add_patch(Circle((0, 0), deposit, facecolor="white", edgecolor="none", zorder=6))

for ax in (axA, axB):
    ax.set_aspect("equal"); ax.axis("off")
    ax.set_xlim(-1.55, 1.55); ax.set_ylim(-1.55, 1.75)

draw_pipe(axA)
axA.annotate("", xy=(0, R_IN), xytext=(0, -R_IN),
             arrowprops=dict(arrowstyle="<->", color=NAVY, lw=1.5), zorder=8)
axA.text(0, 0, "ID\n2.30 in", fontsize=10, fontweight="bold", color=NAVY,
         va="center", ha="center", zorder=9,
         bbox=dict(boxstyle="round,pad=0.14", fc="white", ec="none", alpha=0.95))
axA.set_title("BEFORE \u2014 clean bore", fontsize=12, fontweight="bold", color=NAVY, pad=4)

draw_pipe(axB, deposit=R_DEP)
axB.annotate("", xy=(0, R_DEP), xytext=(0, -R_DEP),
             arrowprops=dict(arrowstyle="<->", color=NAVY, lw=1.5), zorder=8)
axB.text(0, 0, "ID\n1.35 in", fontsize=10, fontweight="bold", color=NAVY,
         va="center", ha="center", zorder=9,
         bbox=dict(boxstyle="round,pad=0.10", fc="white", ec="none", alpha=0.95))
axB.annotate("deposit\n\u2248 0.47 in/side", xy=(0.54, 0.70), xytext=(0.10, 1.52),
             fontsize=9.5, color=RED, ha="center", va="center", zorder=9,
             arrowprops=dict(arrowstyle="-|>", color=RED, lw=1.1))
axB.set_title("AFTER \u2014 deposit on the wall", fontsize=12, fontweight="bold", color=RED, pad=4)

# --- panel C: gauge-ring deposit profile (smooth, realistic shape) ---
Dc = np.linspace(800, 5600, 400)
# smooth beta-like profile: zero until ~1,600 m, peak ~4,300 m, tapering by TD
t = (Dc - 1600) / (5600 - 1600)
prof = 66 * np.clip(t, 0, 1)**1.5 * (1 - np.clip(t, 0, 1))**0.35
prof = 66 * prof / prof.max()
axC.fill_betweenx(Dc, 0, prof, color=ORANGE, alpha=0.12, zorder=2)
axC.plot(prof, Dc, color=ORANGE, lw=3.0, zorder=4, solid_capstyle="round")
for d_run, val in ((2600, None), (3400, None), (4200, None), (4900, None)):
    idx = np.argmin(np.abs(Dc - d_run))
    axC.plot([0, 71.5], [d_run, d_run], color=GRAY, lw=0.75, ls=(0, (2, 2)), alpha=0.5, zorder=3)
    axC.plot([prof[idx]], [d_run], marker="o", ms=5, color=NAVY, zorder=5)
axC.annotate("successive gauge-ring runs\nmap the profile",
             xy=(prof[np.argmin(np.abs(Dc-3400))], 3400), xytext=(30, 1750),
             fontsize=9.5, color=NAVY, ha="left", va="center",
             arrowprops=dict(arrowstyle="->", color=NAVY, lw=1.0))
axC.axvline(66.7, color=RED, lw=1.4, ls=":")
axC.annotate("measured maximum\n\u2248 \u2154 of radius",
             xy=(65.5, 4850), xytext=(26, 5320), fontsize=9.5, color=RED,
             ha="center", va="center",
             arrowprops=dict(arrowstyle="->", color=RED, lw=1.1))
axC.set_ylim(5600, 800); axC.set_xlim(0, 74)
axC.set_xlabel("deposit thickness  (% of tubing radius)", fontsize=10)
axC.set_ylabel("Depth (m)", fontsize=10)
axC.set_title("Gauge-ring deposit profile", fontsize=12, fontweight="bold", color=NAVY, pad=4)
axC.grid(alpha=0.12, lw=0.6)
axC.tick_params(labelsize=9.5)

fig.suptitle("Tubing damage, three views \u2014 two-thirds of the bore can disappear",
             fontsize=13.5, fontweight="bold", color=NAVY, y=0.97)
fig.text(0.012, 0.058,
         "Cross-sections to scale from the measured ID reduction 2.30 \u2192 1.35 in (Wylde & Punase, JPT, 2020).",
         fontsize=9, color=GRAY)
fig.text(0.012, 0.012,
         "Gauge-ring profile: schematic after the Hassi Messaoud field practice (Haskett & Tartera, SPE-994-PA, 1965).",
         fontsize=9, color=GRAY)
export(fig, "pipe_views")

# =========================================================
# 4 — CORE-SHELL STABILITY SCHEMATIC (slide 3) — unchanged content, vector export
# =========================================================
rng = np.random.default_rng(7)
fig, axes = plt.subplots(3, 1, figsize=(6.4, 7.2), dpi=300)
CY = 1.8

def jagged(ax, cx, cy, r, seed=1, z=3):
    a = np.linspace(0, 2*np.pi, 15)[:-1]
    rr = r * (0.78 + 0.34 * np.random.default_rng(seed).random(len(a)))
    ax.add_patch(Polygon(np.column_stack([cx + rr*np.cos(a), cy + rr*np.sin(a)]),
                         closed=True, facecolor="#2B2118", edgecolor="#120C07", lw=1.0, zorder=z))
    ax.add_patch(Circle((cx, cy), r*0.55, facecolor="#3E3020", edgecolor="none", zorder=z+1))

def molecule(ax, x, y, r=0.11, color=TEAL, z=4):
    ax.add_patch(Circle((x, y), r, facecolor=color, edgecolor="white", lw=0.7, zorder=z))

for ax in axes:
    ax.set_xlim(0, 10); ax.set_ylim(0, 3.6); ax.axis("off"); ax.set_aspect("equal")

ax = axes[0]
ax.add_patch(Circle((5, CY), 1.18, facecolor="#F8FAFC", edgecolor="#E8EFF5", lw=1.2, zorder=1))
jagged(ax, 5, CY, 0.50, seed=3)
for a in np.linspace(0, 2*np.pi, 16, endpoint=False):
    molecule(ax, 5 + 0.92*np.cos(a), CY + 0.92*np.sin(a))
ax.text(5, 3.42, "1 \u00b7 STABLE — the resin / aromatic shell is intact", fontsize=13,
        fontweight="bold", color=NAVY, va="center", ha="center")
ax.text(5, 0.22, "The shell keeps the asphaltene cores dispersed in the oil.",
        fontsize=11.5, color=GRAY, va="center", ha="center")

ax = axes[1]
jagged(ax, 5, CY, 0.50, seed=3)
for a in [0.5, 2.2, 2.9, 5.2]:
    molecule(ax, 5 + 0.92*np.cos(a), CY + 0.92*np.sin(a))
for a in [0.9, 3.4]:
    molecule(ax, 5 + 0.92*np.cos(a), CY + 0.92*np.sin(a), color="#8FBFc4")
for _ in range(7):
    a = rng.uniform(0, 2*np.pi); R = rng.uniform(1.15, 1.42)
    molecule(ax, 5 + R*np.cos(a), CY + R*np.sin(a), color="#8FBFc4")
ax.text(5, 3.42, "2 \u00b7 DESTABILISED — the shell is stripped away", fontsize=13,
        fontweight="bold", color=ORANGE_D, va="center", ha="center")
ax.text(5, 0.22, "Pressure drop or injected gas removes the protective shell.",
        fontsize=11.5, color=GRAY, va="center", ha="center")

ax = axes[2]
for (cx, cy, s) in ((4.55, 1.75, 11), (5.60, 2.00, 3), (5.50, 1.30, 5)):
    jagged(ax, cx, cy, 0.46, seed=s)
for _ in range(5):
    a = rng.uniform(0, 2*np.pi); R = rng.uniform(1.15, 1.45)
    molecule(ax, 5 + R*np.cos(a), CY + R*np.sin(a), color="#8FBFc4")
ax.text(5, 3.42, "3 \u00b7 FLOCCULATED — cores clump and can stick", fontsize=13,
        fontweight="bold", color=RED, va="center", ha="center")
ax.text(5, 0.22, "Precipitation is not deposition — but flocculation enables it.",
        fontsize=11.5, color=GRAY, va="center", ha="center")
fig.tight_layout(h_pad=0.3)
export(fig, "core_shell")

# =========================================================
# 5 — WELL SCHEMATIC: capillary string, mandrel, onset  (slide 16)
# =========================================================
fig, ax = plt.subplots(figsize=(4.95, 3.32), dpi=300)
ax.set_xlim(0, 10); ax.set_ylim(6000, -350); ax.axis("off")
ax.set_aspect("auto")

CX = 2.2                     # well centreline (data x)
# reservoir layer
ax.add_patch(Rectangle((0, 5300), 10, 700, facecolor="#E8EFF5", edgecolor="none", zorder=0))
# casing + tubing + capillary
ax.add_patch(Rectangle((CX-0.55, 0), 1.10, 5300, facecolor="#E4ECF2", edgecolor=GRAY, lw=1.1, zorder=1))
ax.add_patch(Rectangle((CX-0.30, 0), 0.60, 5300, facecolor="white", edgecolor=NAVY, lw=1.4, zorder=2))
ax.plot([CX-0.44, CX-0.44], [60, 4450], color=ORANGE, lw=2.4, zorder=4, solid_capstyle="round")
ax.plot([CX-0.44, CX-0.30], [4450, 4450], color=ORANGE, lw=2.4, zorder=4)
ax.add_patch(FancyBboxPatch((CX-0.52, 4380), 0.46, 120, boxstyle="round,pad=3,rounding_size=10",
                            facecolor="none", edgecolor=ORANGE, lw=2.0, zorder=5))
# wellhead
ax.add_patch(Rectangle((CX-1.0, -60), 2.0, 110, facecolor=NAVY, edgecolor=NAVY, zorder=6))
# perforations
for yy in np.arange(5400, 5470, 20):
    ax.plot([CX-0.55, CX-1.05], [yy, yy], color=TEAL, lw=1.8, zorder=3)
    ax.plot([CX+0.55, CX+1.05], [yy, yy], color=TEAL, lw=1.8, zorder=3)
# intervals
ax.add_patch(Rectangle((0, 4450), 10, 1040, facecolor=ORANGE, alpha=0.10, zorder=0))
ax.plot([0, 10], [4752, 4752], color=ORANGE, lw=2.0, ls="--", zorder=6)
ax.scatter([CX], [4752], marker="*", s=200, color=ORANGE, edgecolor=NAVY, linewidth=1.2, zorder=8)
# depth bracket mandrel -> onset
ax.annotate("", xy=(CX-1.02, 4450), xytext=(CX-1.02, 4752),
            arrowprops=dict(arrowstyle="<->", color=RED, lw=1.5, mutation_scale=9), zorder=9)
ax.text(CX-1.32, 4600, "\u2248300 m", fontsize=9, color=RED, rotation=90,
        ha="center", va="center", zorder=9)

# ---- right-hand label column (fixed screen positions: no collisions) ----
def note(frac, txt, color, depth_point, weight="normal"):
    ax.annotate(txt, xy=(CX + 0.60, depth_point), xycoords="data",
                xytext=(0.40, frac), textcoords="axes fraction",
                fontsize=10.5, color=color, ha="left", va="center", fontweight=weight,
                zorder=10,
                arrowprops=dict(arrowstyle="-", color=color, lw=0.9, alpha=0.75))

note(0.98, "wellhead", NAVY, -5)
note(0.87, "production tubing", NAVY, 1900)
note(0.76, "capillary string\n(chemical dosing)", ORANGE, 3400)
note(0.59, "EXISTING MANDREL\n4,450 m (14,600 ft)", ORANGE, 4450, "bold")
note(0.47, "PREDICTED ONSET\n4,752 m", RED, 4752, "bold")
note(0.30, "3,400 ft UNPROTECTED\n(below the mandrel)", ORANGE_D, 5300)
note(0.15, "perforations\n17,700\u201318,000 ft", TEAL, 5435)

ax.set_title("Where the chemical must arrive", fontsize=12, fontweight="bold",
             color=NAVY, pad=5)
fig.text(0.012, 0.064,
         "Schematic, not to scale. Case geometry after Wylde & Punase, JPT, 2020.",
         fontsize=9.5, color=GRAY)
fig.text(0.012, 0.015,
         "Onset depth from the illustrative model on slide 13.", fontsize=9.5, color=GRAY)
fig.tight_layout(rect=[0, 0.075, 1, 1])
export(fig, "well_schematic")
print("ALL FIGURES DONE")
