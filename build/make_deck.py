#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Asphaltenes seminar deck — 23 core + 2 backup slides, speaker notes, timing clocks."""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import os

# ------------- palette (same visual language as the previous package) -------------
# ---------------- DESIGN TOKENS (single source of truth) ----------------
NAVY   = RGBColor(0x1B, 0x36, 0x5D)   # titles, table headers        (contrast vs white 12.1:1)
NAVY2  = RGBColor(0x2A, 0x4A, 0x73)   # secondary navy, box borders  (8.4:1)
TEAL   = RGBColor(0x0D, 0x73, 0x77)   # section labels, benign elements (5.3:1)
AMBER  = RGBColor(0xE0, 0x7A, 0x3D)   # call-out borders, icons      (graphic use)
ORANGE = RGBColor(0xE0, 0x7A, 0x3D)   # danger / onset / deposit
AMB_D  = RGBColor(0xB3, 0x5A, 0x18)   # dark orange for small text   (5.4:1)
GRAY   = RGBColor(0x4A, 0x55, 0x68)   # captions, footers            (7.6:1)
TEXT   = RGBColor(0x1A, 0x1A, 0x1A)   # body text                    (17.4:1)
LIGHT  = RGBColor(0xF8, 0xFA, 0xFC)   # box fill
LIGHT2 = RGBColor(0xE8, 0xEF, 0xF5)   # alternate row / muted fill
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
RED    = RGBColor(0xB4, 0x3B, 0x2E)   # reframes "this is wrong/danger" (6.1:1)
REDBG  = RGBColor(0xFB, 0xEA, 0xE5)   # tinted fill for the key distinction box
# LAYOUT ARCHETYPES (every slide declares one; geometry lives in the builder fns)
#  A-title | B-section | C-content-split | D-full-width | E-table | F-closing | G-backup
# ------------------------------------------------------------------------

FONT = "Calibri"
SW, SH = Inches(13.333), Inches(7.5)
FIGS = "/home/user/Seminar_Asphaltenes/_build/figs"

prs = Presentation()
prs.slide_width, prs.slide_height = SW, SH

def blank_slide():
    return prs.slides.add_slide(prs.slide_layouts[6])

def box(slide, x, y, w, h):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    return tf

def body_up(size):
    """Type scale: body text below 19 pt is lifted by 1 pt (content-slide floor ~13 pt effective)."""
    return size + 1 if 12 <= size <= 18 else size

def par(tf, text, size=16, bold=False, color=TEXT, first=False, align=PP_ALIGN.LEFT,
        space_after=8, italic=False, line=1.02):
    size = body_up(size)
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align; p.space_after = Pt(space_after); p.line_spacing = line
    r = p.add_run(); r.text = text
    r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic
    r.font.color.rgb = color; r.font.name = FONT
    return p

def rect(slide, x, y, w, h, fill=None, line=None, lw=1.0, shape=MSO_SHAPE.RECTANGLE, radius=None):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None: s.fill.background()
    else: s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None: s.line.fill.background()
    else: s.line.color.rgb = line; s.line.width = Pt(lw)
    s.shadow.inherit = False
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        try: s.adjustments[0] = radius
        except Exception: pass
    tf = s.text_frame; tf.word_wrap = True
    tf.margin_left = Inches(0.12); tf.margin_right = Inches(0.12)
    tf.margin_top = Inches(0.06); tf.margin_bottom = Inches(0.06)
    return s, tf

TARGETS = {1:15,2:45,3:35,4:35,5:35,6:20,7:35,8:40,9:35,10:45,11:40,12:45,13:55,
           14:35,15:45,16:40,17:40,18:40,19:45,20:50,21:40,22:40,23:10}
def clocks():
    out = {}; c = 0
    for n in range(1, 24):
        c += TARGETS[n]; out[n] = f"{c//60}:{c%60:02d}"
    return out
CLK = clocks()
print("total speaking plan:", sum(TARGETS.values()), "s =", f"{sum(TARGETS.values())//60}:{sum(TARGETS.values())%60:02d}")

def header(slide, kicker, headline, n, target=None, clock=None, backup=False):
    tf = box(slide, 0.62, 0.30, 12.1, 0.35)
    par(tf, kicker, size=11.5, bold=True, color=TEAL, first=True, space_after=0)
    tf2 = box(slide, 0.62, 0.56, 12.24, 0.95)
    par(tf2, headline, size=24 if len(headline) < 95 else 22, bold=True, color=NAVY,
        first=True, space_after=0, line=1.0)
    rect(slide, 0.66, 1.545, 1.15, 0.055, fill=AMBER)
    ftf = box(slide, 0.62, 7.06, 8.4, 0.3)
    par(ftf, "Asphaltenes in production wellbores — flow assurance  ·  Seminar  ·  [Your name]",
        size=9, color=GRAY, first=True, space_after=0)
    rtf = box(slide, 8.9, 7.06, 3.8, 0.3)
    label = ("BACKUP slide %d/25 — show only if asked" % n) if backup else \
            ("Slide %d/25 · target %ds · clock %s" % (n, target, clock))
    par(rtf, label, size=9, color=GRAY, first=True, space_after=0, align=PP_ALIGN.RIGHT)

def parruns(tf, runs, size=13, first=False, align=PP_ALIGN.LEFT, space_after=0, line=1.02):
    """runs = list of (text, bold, color)"""
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align; p.space_after = Pt(space_after); p.line_spacing = line
    for text, bold, color in runs:
        r = p.add_run(); r.text = text
        r.font.size = Pt(body_up(size)); r.font.bold = bold
        r.font.color.rgb = color; r.font.name = FONT
    return p

def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text

def table(slide, data, x, y, w, col_w=None, fs=12.5, hfs=12.5, row_h=0.36):
    fs = min(fs + 0.5, 13); hfs = min(hfs + 0.5, 13.5); row_h = row_h + 0.02
    rows, cols = len(data), len(data[0])
    gf = slide.shapes.add_table(rows, cols, Inches(x), Inches(y), Inches(w), Inches(row_h*rows))
    tbl = gf.table; tbl.first_row = False; tbl.horz_banding = False
    if col_w:
        tot = sum(col_w)
        for i, cw in enumerate(col_w):
            tbl.columns[i].width = Emu(int(Inches(w) * cw / tot))
    for r in range(rows):
        for c in range(cols):
            cell = tbl.cell(r, c); cell.text_frame.word_wrap = True
            cell.margin_left = Inches(0.08); cell.margin_right = Inches(0.06)
            cell.margin_top = Inches(0.02); cell.margin_bottom = Inches(0.02)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            run = p.add_run(); run.text = str(data[r][c])
            run.font.name = FONT; run.font.size = Pt(hfs if r == 0 else fs)
            run.font.bold = (r == 0)
            run.font.color.rgb = WHITE if r == 0 else TEXT
            if r == 0: cell.fill.solid(); cell.fill.fore_color.rgb = NAVY
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = WHITE if r % 2 == 1 else LIGHT
    return tbl

# =========================================================
# 1 — TITLE
# =========================================================
s = blank_slide()
rect(s, 0, 0, 13.333, 7.5, fill=WHITE)
rect(s, 0, 0, 0.30, 7.5, fill=NAVY)
rect(s, 0.30, 0, 0.10, 7.5, fill=AMBER)
tf = box(s, 1.0, 0.72, 11.4, 0.4)
par(tf, "SEMINAR  ·  PETROLEUM AND GAS ENGINEERING  ·  [COURSE CODE]", size=12.5, bold=True,
    color=TEAL, first=True, space_after=0)
tf = box(s, 1.0, 1.22, 11.3, 2.3)
par(tf, "Asphaltenes: Predicting and Preventing Wellbore Plugging", size=38, bold=True,
    color=NAVY, first=True, space_after=4, line=1.0)
par(tf, "Flow assurance in production wellbores — and the seed of my graduation project",
    size=19, color=GRAY, space_after=0, line=1.05)
tf = box(s, 1.0, 4.00, 6.6, 2.5)
par(tf, "Presenter:  [Your full name]   ·   ID [xxxxxx]", size=15, color=TEXT, first=True, space_after=6)
par(tf, "Supervisor:  [Prof. name]", size=15, color=TEXT, space_after=6)
par(tf, "Programme:  BSc Petroleum and Gas Engineering — [University]", size=15, color=TEXT, space_after=6)
par(tf, "[Month] 2026   ·   15-minute presentation   ·   23 slides + 2 backup", size=15, color=TEAL, space_after=0)
r, t = rect(s, 8.1, 4.00, 4.2, 2.3, fill=LIGHT, line=LIGHT2, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
par(t, "ONE-LINE SUMMARY", size=11, bold=True, color=TEAL, first=True, space_after=4)
par(t, "Field engineers call asphaltenes the cholesterol of crude oil. This seminar shows when they block wells, how we predict it — and the project that will quantify it for a real case.",
    size=13, color=TEXT, space_after=0, line=1.1)
notes(s, """TARGET 15 s (clock 0:15).
[EDIT] Replace all bracketed placeholders.
Opening lines, memorised verbatim: "Good morning. My name is [—]. In fifteen minutes I will show you how the heaviest fraction of crude oil can choke a well to death — and how we can predict it before it happens." """)

# =========================================================
# 2 — WHY IT MATTERS (numbers)
# =========================================================
s = blank_slide()
header(s, "PART 1 · THE PROBLEM", "Asphaltenes are a billion-dollar flow-assurance problem", 2, 45, CLK[2])
cards = [
    ("billions $/yr", "industry-wide cost of deposition — Farooq et al., Energy & Fuels (2021) [1]"),
    ("≈ $70 M", "one deepwater GoM well: shut-in + cleanup — Farooq et al. (2021) [1]"),
    ("up to $1.2 M/day", "deferred production during a plugging event — Stratiev et al., Processes (2025) [2]"),
    ("up to ⅔", "of tubing radius: measured deposits, Hassi Messaoud — SPE-994-PA (1965) [3]"),
]
x = 0.62
for big, cap in cards:
    r, t = rect(s, x, 1.80, 2.94, 1.95, fill=LIGHT, line=TEAL, lw=1.2,
                shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.07)
    par(t, big, size=23, bold=True, color=NAVY, first=True, space_after=4, align=PP_ALIGN.CENTER)
    par(t, cap, size=10.5, color=GRAY, space_after=0, align=PP_ALIGN.CENTER, line=1.05)
    x += 3.10
tf = box(s, 0.62, 4.00, 12.1, 2.7)
par(tf, "When production pressure drops, the heaviest molecules in the crude crash out of solution and build a hard, tar-like deposit that chokes the tubing.",
    size=14.5, first=True, space_after=10)
par(tf, "Not an exotic failure mode: asphaltenes are managed somewhere in the world every day — North Sea, Gulf of Mexico, Algerian Sahara.",
    size=14.5, space_after=10)
par(tf, "The sting is not only the money — the deposit arrives unplanned, and the industry's default answers are reactive.",
    size=14.5, space_after=10)
par(tf, "So the question of this seminar: can we predict WHERE it will deposit — before the well chokes?",
    size=15, bold=True, color=NAVY, space_after=0)
notes(s, """TARGET 45 s (clock 1:00).
VERIFIED (exact quotes in 04_Source_Log.docx):
[1] Farooq, U.; Laedre, S.; Gawel, K. Review of Asphaltenes in an Electric Field. Energy & Fuels 2021, 35(9), 7285-7304 — "Asphaltene deposition costs the oil industry billions of dollars every year"; "the cost of a well shut-in for a cleanup operation due to asphaltene deposition was approximately $70 million/well."
[2] Stratiev, D. et al. Mitigation of Asphaltene Deposit Formation via Chemical Additives: A Review. Processes 2025, 13(1), 141 — "the financial loss due to lost production can be as high as USD 1.2 M/day."
[3] Haskett, C.E.; Tartera, M. SPE-994-PA / JPT 1965 — Hassi Messaoud; deposits ≈ 2/3 of tubing radius in five wells (restated in later field-case literature).
Delivery: land each number slowly; the $70 M and $1.2 M/day figures do the heavy lifting.""")

# =========================================================
# 3 — WHAT ARE ASPHALTENES
# =========================================================
s = blank_slide()
header(s, "PART 1 · THE PROBLEM", "What are asphaltenes? The heaviest, most polar fraction of crude — held in suspension", 3, 35, CLK[3])
tf = box(s, 0.62, 1.80, 7.4, 4.9)
items = [
    ("Operational definition (not one molecule, but a solubility class):", "insoluble in n-heptane, soluble in toluene — the basis of the SARA analysis (Saturates · Aromatics · Resins · Asphaltenes): ASTM D6560 / IP 143 [12]."),
    ("Why they stay in the oil:", "aromatic-rich resins form the stabilising shell around nano-aggregates of ~2 nm; the crude is stable while that dispersion is maintained (Yen–Mullins picture) [11]."),
    ("Why the nickname fits:", "field engineers call them the cholesterol of crude oil — and like cholesterol, the danger is not their presence, it is precipitation from circulation."),
    ("Scale:", "typically a few per cent of the oil by weight — enough to plug a well completely."),
]
for i, (a, b) in enumerate(items):
    par(tf, "▪  " + a, size=14.5, bold=True, color=NAVY, first=(i == 0), space_after=2)
    par(tf, b, size=13.5, color=TEXT, space_after=10, line=1.05)
s.shapes.add_picture(os.path.join(FIGS, "core_shell.png"), Inches(8.25), Inches(1.78),
                     width=Inches(4.45), height=Inches(5.0))
notes(s, """TARGET 35 s (clock 1:35).
ANIMATION: the right-hand schematic has three panels — cover panels 2 and 3 with white rectangles and reveal them (Appear, on click) as you say "destabilised" and "flocculated". This is the single highest-value build in the deck.
Key correction to make explicitly (committees test this): asphaltenes do not "keep the oil stable" — they are kept dispersed BY resins and aromatics.
VERIFIED: ASTM D6560 = determination of asphaltenes (n-heptane insolubles) in crude petroleum; equivalent to IP 143. Cite the edition year of the copy you consult.""")

# =========================================================
# 4 — VISUAL EVIDENCE + IMPACT
# =========================================================
s = blank_slide()
header(s, "PART 1 · THE PROBLEM", "Inside the tubing: two-thirds of the flow area, gone", 4, 35, CLK[4])
s.shapes.add_picture(os.path.join(FIGS, "pipe_views.png"), Inches(0.62), Inches(1.80),
                     width=Inches(7.55), height=Inches(3.38))
tf = box(s, 0.62, 5.30, 7.55, 1.5)
par(tf, "Left and centre: cross-sections drawn TO SCALE from the measured ID reduction 2.30 \u2192 1.35 in [4].",
    size=10, italic=True, color=GRAY, first=True, space_after=3, line=1.05)
par(tf, "Right: gauge-ring deposit profile — the field method that maps where the deposit sits, panel by panel, so treatment volumes can be calculated [3].",
    size=10, italic=True, color=GRAY, space_after=3, line=1.05)
par(tf, "If you can obtain a real cut-tubing photograph with permission, it replaces the cross-sections — keep the profile panel either way.",
    size=9.5, italic=True, color=AMB_D, space_after=0, line=1.05)
tf = box(s, 8.42, 1.85, 4.30, 5.0)
items = [
    ("Flow area shrinks.", "ID 2.30 → 1.35 in: about two-thirds of the flow area lost [4]."),
    ("Pressure is consumed downhole.", "Severe fields lost 20–25 % of wellhead pressure within 15–20 days [3]."),
    ("Production declines, then stops.", "A choking well can cease flowing within days."),
    ("Invisible until expensive.", "No downhole window — the pressure signature arrives late."),
]
for i, (a, b) in enumerate(items):
    p = par(tf, "▪  " + a, size=13, bold=True, color=NAVY, first=(i == 0), space_after=2)
    par(tf, b, size=12, color=TEXT, space_after=10, line=1.04)
notes(s, """TARGET 35 s (clock 2:10).
VERIFIED — Ref [4] = Wylde & Punase, JPT (2020): the 2.3 → 1.35 in ID reduction comes from their transient multiphase flow model. Ref [3] = Hassi Messaoud (SPE-994-PA, 1965): "Wells often lost 20 to 25 per cent of the wellhead pressure in 15 to 20 days."
Delivery: point at the flow restriction in the image: "this is a pipe that still has a hole in it — for now." """)

# =========================================================
# 5 — OBJECTIVE & SCOPE
# =========================================================
s = blank_slide()
header(s, "PART 1 · THE PROBLEM", "Objective: explain the mechanism, review the defenses, propose the project", 5, 35, CLK[5])
labels = [
    ("1 · EXPLAIN", "what destabilises asphaltenes, and the chain from precipitation to deposition", TEAL),
    ("2 · REVIEW", "today's defenses — and why they are mostly reactive", AMBER),
    ("3 · PROPOSE", "the graduation project: predicting onset depth and injection point for a real well", NAVY2),
]
x = 0.62
for lab, txt, col in labels:
    r, t = rect(s, x, 1.85, 3.95, 2.5, fill=WHITE, line=col, lw=1.6,
                shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
    par(t, lab, size=16, bold=True, color=col, first=True, space_after=8)
    par(t, txt, size=14, color=TEXT, space_after=0, line=1.08)
    x += 4.13
r, t = rect(s, 0.62, 4.70, 12.1, 1.15, fill=LIGHT, line=None, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.10)
par(t, "Scope statement:  this is a flow-assurance review and a feasibility study for the thesis — not a completed field intervention study.",
    size=14.5, bold=True, color=NAVY, first=True, space_after=4)
par(t, "Saying this out loud protects you in the Q&A: the committee knows exactly what you claim and what you do not.",
    size=12.5, italic=True, color=GRAY, space_after=0)
notes(s, """TARGET 35 s (clock 2:45).
Key message: three verbs — explain, review, propose. The scope statement is your shield.
Delivery: same discipline as a good paper abstract: promise less, deliver exactly.""")

# =========================================================
# 6 — ROADMAP
# =========================================================
s = blank_slide()
header(s, "PART 1 · THE PROBLEM", "How the next 14 minutes are organised", 6, 20, CLK[6])
rows = [
    ("PART 1", "The problem: what asphaltenes are, and what they cost", "0:00 – 2:45"),
    ("PART 2", "The science: stability, triggers, and the P–T envelope", "2:45 – 6:20"),
    ("PART 3", "Field evidence: a documented deepwater case study", "6:20 – 9:20"),
    ("PART 4", "Defenses then and now — and the economics of prevention", "9:20 – 11:20"),
    ("PART 5", "The graduation project: prediction workflow and plan", "11:20 – 14:25"),
]
y = 1.85
for p, t, c in rows:
    r, tfr = rect(s, 0.62, y, 12.1, 0.80, fill=LIGHT if p != "PART 5" else LIGHT2,
                  line=None, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.12)
    par(tfr, p + "    " + t, size=15.5, bold=(p != "PART 5"), color=NAVY if p != "PART 5" else NAVY2,
        first=True, space_after=0)
    tbx = box(s, 10.4, y + 0.22, 2.2, 0.4)
    par(tbx, c, size=13, bold=True, color=TEAL, first=True, space_after=0, align=PP_ALIGN.RIGHT)
    y += 0.90
r, t = rect(s, 0.62, 6.35, 12.1, 0.50, fill=LIGHT2, line=TEAL, lw=1.2,
            shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.14)
par(t, "SCOPE — a flow-assurance review plus a thesis feasibility study; not a completed field intervention.",
    size=12.5, bold=True, color=NAVY, first=True, space_after=0, align=PP_ALIGN.CENTER)
notes(s, """TARGET 20 s (clock 3:05).
Scope line (bottom of the slide): say it once, do not apologise for it — it defines what is and is not being claimed.
Five parts, not four — the field case study gets its own part because the audience needs to see real evidence before hearing about defenses.
Delivery: fast and confident; do not linger.""")

# =========================================================
# 7 — WHAT KEEPS OIL STABLE
# =========================================================
s = blank_slide()
header(s, "PART 2 · THE SCIENCE", "Stability is a fragile balance — pressure, temperature and composition all pull on it", 7, 35, CLK[7])
tf = box(s, 0.62, 1.80, 6.3, 4.9)
items = [
    ("Two engineering views, one working rule:", "colloidal (aggregates protected by resins) and solubility (asphaltenes dissolve best in aromatic, dense fluids)."),
    ("Field screening uses solubility parameters:", "condensed into the Colloidal Instability Index [5] (Yen et al.; threshold scale after Asomaning) — higher CII, less stable crude."),
    ("Three levers move the balance:", "pressure, temperature, and composition — gas or condensate changes the oil's solvent quality."),
    ("Engineering consequence:", "anything that lightens the oil or drops its pressure pushes it toward instability."),
]
for i, (a, b) in enumerate(items):
    par(tf, "▪  " + a, size=14, bold=True, color=NAVY, first=(i == 0), space_after=2)
    par(tf, b, size=13, color=TEXT, space_after=9, line=1.05)
r, t = rect(s, 7.15, 1.90, 5.55, 3.6, fill=LIGHT, line=NAVY2, lw=1.3, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
par(t, "THE THREE LEVERS", size=13, bold=True, color=NAVY2, first=True, space_after=10, align=PP_ALIGN.CENTER)
par(t, "PRESSURE ↓", size=17, bold=True, color=RED, space_after=2, align=PP_ALIGN.CENTER)
par(t, "main trigger — loss of light-end solvency as P falls", size=12, color=GRAY, space_after=10, align=PP_ALIGN.CENTER)
par(t, "COMPOSITION changed", size=17, bold=True, color=RED, space_after=2, align=PP_ALIGN.CENTER)
par(t, "gas lift, gas breakthrough, CO₂ / rich-gas injection, commingling", size=12, color=GRAY, space_after=10, align=PP_ALIGN.CENTER)
par(t, "TEMPERATURE", size=17, bold=True, color=AMB_D, space_after=2, align=PP_ALIGN.CENTER)
par(t, "secondary modifier — shifts the envelope, direction is oil-specific", size=12, color=GRAY, space_after=0, align=PP_ALIGN.CENTER)
r, t = rect(s, 0.62, 5.60, 12.1, 1.05, fill=LIGHT, line=None,
            shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.10)
parruns(t, [("CII = (", False, NAVY), ("Saturates + Asphaltenes", True, RGBColor(0xE8, 0x63, 0x0A)),
            (") / (", False, NAVY), ("Resins + Aromatics", True, TEAL), (")", False, NAVY)],
        size=16, first=True, space_after=3)
parruns(t, [("orange numerator = the destabilising fraction", False, RGBColor(0xE8, 0x63, 0x0A)),
            ("   ·   ", False, GRAY),
            ("teal denominator = the crude's own dispersant", False, TEAL),
            ("   ·   ", False, GRAY),
            ("higher CII = less stable crude", True, NAVY)],
        size=11.5, space_after=0)
notes(s, """TARGET 35 s (clock 3:40).
VERIFIED — Ref [5] = Yen, A.; Yin, Y.R.; Asomaning, S. Evaluating Asphaltene Inhibitors: Laboratory Tests and Field Studies, SPE 65376 (2001), for the CII; thresholds after Asomaning (2003): CII ≥ 0.9 unstable, 0.7–0.9 uncertain, < 0.7 stable (confirmed in three independent sources; thresholds remain laboratory-specific).
Delivery: keep the "two views, one working rule" framing — it shows you read beyond a single textbook.""")

# =========================================================
# 8 — ASPHALTENES vs WAX
# =========================================================
s = blank_slide()
header(s, "PART 2 · THE SCIENCE", "Asphaltenes are not wax — the trigger, the physics and the fix all differ", 8, 40, CLK[8])
data = [
    ["", "Paraffin wax", "Asphaltenes"],
    ["Main trigger", "temperature falls below the wax appearance temperature (WAT)", "pressure falls below the onset pressure; composition change"],
    ["Nature of the solid", "crystalline — it can melt", "amorphous — it does not remelt; thermal methods largely fail"],
    ["Where solids appear", "in flowlines and surface lines, cold spots", "wherever the flowing P–T path crosses the envelope — often deep in the tubing"],
    ["Deposition mechanism", "diffusion + shear; hard wax layer grows inward", "adhesion of flocculated aggregates; can form throughout the flow path"],
    ["Removal", "hot oil, thermal insulation, scrapers", "aromatic solvents (toluene/xylene), milling, dispersants"],
    ["Inhibition", "crystal modifiers / pour-point depressants", "polymeric dispersants (inhibitors) injected continuously below the onset"],
]
table(s, data, 0.62, 1.72, 12.1, col_w=[2.5, 4.6, 5.0], fs=11.5, hfs=12, row_h=0.60)
tf = box(s, 0.62, 5.98, 12.1, 0.34)
par(tf, "NEXT  →  how that instability becomes a deposit in the tubing.", size=12,
    italic=True, color=TEAL, first=True, space_after=0)
tf = box(s, 0.62, 6.35, 12.1, 0.5)
par(tf, "Temperature matters — but as a modifier of the envelope, not the primary trigger. Juries test this distinction.",
    size=12.5, italic=True, color=GRAY, first=True, space_after=0)
notes(s, """TARGET 40 s (clock 4:20).
Key message: wax = cooling story; asphaltenes = pressure/composition story. This single slide is the credibility marker of the seminar.
Delivery: do not read rows; say only three contrasts — trigger, nature of the solid, removal.""")

# =========================================================
# 9 — MECHANISM CHAIN
# =========================================================
s = blank_slide()
header(s, "PART 2 · THE SCIENCE", "From stable oil to plugged tubing: four steps — and the critical distinction", 9, 35, CLK[9])
steps = [
    ("1 · PRECIPITATION", "Thermodynamic. The oil crosses the envelope and asphaltenes separate from solution as nano-aggregates.", RED),
    ("2 · FLOCCULATION", "Particles stick together into larger, irregular clusters.", AMB_D),
    ("3 · AGGREGATION", "Clusters grow; some can approach millimetre scale.", TEAL),
    ("4 · DEPOSITION", "Particles adhere to the steel and form a deposit — balanced by shear removal.", NAVY2),
]
ICONS_S9 = ["icon_flask.png", "icon_clusters.png", "icon_clusters.png", "icon_well.png"]
x = 0.62
for i, (lab, txt, col) in enumerate(steps):
    r, t = rect(s, x, 1.85, 2.94, 2.85, fill=WHITE, line=col, lw=1.6,
                shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
    s.shapes.add_picture(os.path.join(FIGS, ICONS_S9[i]), Inches(x + 0.16), Inches(1.98),
                         width=Inches(0.68), height=Inches(0.68))
    par(t, "", size=11, first=True, space_after=29)
    par(t, lab, size=13.5, bold=True, color=col, space_after=6)
    par(t, txt, size=12.5, color=TEXT, space_after=0, line=1.06)
    x += 3.10
r, t = rect(s, 0.62, 4.88, 12.1, 1.72, fill=REDBG, line=RED, lw=2.6,
            shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
par(t, "PRECIPITATION ≠ DEPOSITION", size=17, bold=True, color=RED, first=True,
    space_after=4, align=PP_ALIGN.CENTER)
par(t, "Many wells carry precipitated asphaltenes to surface with no deposit at all. Deposition is a transport phenomenon on top of the thermodynamic one — governed by adhesion, particle size and shear. This is exactly where prediction is weakest today.",
    size=13, color=TEXT, space_after=0, line=1.10, align=PP_ALIGN.CENTER)
notes(s, """TARGET 35 s (clock 4:55).
Key message: four steps, but the exam-grade point is step 4 — deposition is not automatic after precipitation.
Delivery: slow down on the red box — it is now the visual anchor of the slide; this idea returns in the gap analysis (slide 18) and the Q&A.""")

# =========================================================
# 10 — STABILITY ENVELOPE (chart A)
# =========================================================
s = blank_slide()
header(s, "PART 2 · THE SCIENCE", "The envelope: precipitation begins the moment the flowing path crosses the AOP", 10, 45, CLK[10])
s.shapes.add_picture(os.path.join(FIGS, "envelope_pt.png"), Inches(4.62), Inches(1.80),
                     width=Inches(8.15), height=Inches(4.46))
tf = box(s, 0.62, 1.85, 3.75, 4.9)
items = [
    ("AOP:", "asphaltene onset pressure — above the line the oil is stable."),
    ("Above the bubble point:", "the upper onset sits above Pb and precipitation peaks near Pb — instability windows reach thousands of psi [6, 15]."),
    ("Below Pb:", "asphaltenes partly re-dissolve; some oils then show a second, lower onset (backup slide)."),
    ("Unstable region:", "below the AOP line, precipitation is thermodynamically possible."),
    ("The flowing path:", "as oil rises, P and T fall together — the crossing is the onset."),
]
for i, (a, b) in enumerate(items):
    p = par(tf, "▪  " + a, size=12.5, bold=True, color=NAVY, first=(i == 0), space_after=2)
    par(tf, b, size=12, color=TEXT, space_after=7, line=1.03)
notes(s, """TARGET 45 s (clock 5:40).
[6] = de Boer, R.B. et al. (1995) SPE Prod. Facil. 10(1) 55-61, doi:10.2118/24987-PA — CAUTION: the de Boer PLOT is in-situ density vs gas under-saturation (P_res - Pb), three risk zones; risk greatest for light, undersaturated crudes. It is NOT an AOP-vs-Pb cross-plot; that is a separate envelope rule used on this slide and stated correctly on backup slide 24. [15] = Mohammed et al., J. Pet. Sci. Eng. 2021, 197, 107956 — instability windows of thousands of psi (2,600-3,800 psi reported) and precipitation peaking near Pb. Exact quotes in 04_Source_Log.docx and 06_Technical_Accuracy_and_Risk_Register.docx (item A3).
Delivery: trace the path with the pointer: reservoir → star (onset) → wellhead. This chart is your central technical exhibit.""")

# =========================================================
# 11 — SCREENING & MEASUREMENT
# =========================================================
s = blank_slide()
header(s, "PART 2 · THE SCIENCE", "How risk is measured: screening first, onset measurement second", 11, 40, CLK[11])
data = [
    ["Method", "What it gives you", "Practical role"],
    ["SARA analysis (ASTM D2007-class)", "the four fractions; feeds CII and solubility modelling", "first-pass screen on any new crude [12]"],
    ["CII / stability indices", "a single number: stable vs unstable tendency", "flags the crude before any lab work [5]"],
    ["De Boer plot", "AOP vs bubble point vs reservoir pressure", "field-level screening; if AOP > Pb, expect problems [6]"],
    ["Automated heptane titration (ASTM D6703)", "the asphaltene precipitation onset in titrations", "quantitative ranking of crudes and blends [12]"],
    ["High-pressure onset measurement (NIR / light scattering in a PVT cell)", "the AOP as a function of temperature — the input a simulator must match", "the calibration anchor for every model [4]"],
]
table(s, data, 0.62, 1.78, 12.1, col_w=[3.6, 5.5, 3.8], fs=11.5, hfs=12, row_h=0.62)
tf = box(s, 0.62, 6.10, 12.1, 0.8)
par(tf, "The best field practice uses several: a screening number for early decisions, and a measured AOP curve to calibrate the prediction. My project uses both.",
    size=13, italic=True, color=GRAY, first=True, space_after=0, line=1.08)
notes(s, """TARGET 40 s (clock 6:20).
VERIFIED: automated Heithaus titrimetry is ASTM D6703 (asphaltene precipitation onset). SARA quantification: ASTM D2007-class / IP 143. Confirm the edition years of the copies you consult.
Delivery: three methods maximum if pressed for time — SARA, CII, AOP measurement.""")

# =========================================================
# 12 — TRIGGERS  (CHECKPOINT 1)
# =========================================================
s = blank_slide()
header(s, "PART 2 · THE SCIENCE", "What triggers it in the wellbore — and where deposits actually form", 12, 45, CLK[12])
tf = box(s, 0.62, 1.80, 6.15, 4.9)
items = [
    ("Pressure drop — the main trigger.", "As oil rises, pressure falls; when the path crosses the AOP, precipitation begins."),
    ("Gas lift changes the solvent.", "Injected gas strips resins off the aggregates; the oil loses solvency."),
    ("Gas breakthrough / CO₂ / rich gas", "shifts the envelope — CO₂ is a documented destabiliser."),
    ("Commingling and blending", "individually stable crudes can form an unstable mixture."),
    ("Acid jobs and workovers", "spent acids destabilise the near-well region."),
]
for i, (a, b) in enumerate(items):
    p = par(tf, "▪  " + a, size=13.5, bold=True, color=NAVY, first=(i == 0), space_after=2)
    par(tf, b, size=12.5, color=TEXT, space_after=6, line=1.03)
r, t = rect(s, 7.0, 1.85, 5.7, 4.35, fill=LIGHT, line=ORANGE, lw=1.5, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
par(t, "CORRECTING A CLASSIC MISCONCEPTION", size=12.5, bold=True, color=RED, first=True, space_after=8)
par(t, "“Deposits form at the top of the well, where the pressure drops fastest.”  Not true — the pressure gradient along the tubing is roughly steady, so the deposit forms wherever the flowing path first crosses the envelope. In deep, hot wells that is often TENS of metres — sometimes a kilometre — below the wellhead.",
    size=13, color=TEXT, space_after=8, line=1.08)
par(t, "Why it matters: that depth decides where you inject inhibitor — and the field case study (next slide) shows what happens when the injection point is too shallow.",
    size=13, bold=True, color=NAVY, space_after=8, line=1.08)
par(t, "NEXT  →  the field case that shows the cost of getting that depth wrong.", size=12,
    italic=True, color=TEAL, space_after=0)
notes(s, """TARGET 45 s (clock 7:05) — CHECKPOINT 1: you should be on this slide at 7:05.
Key message: five triggers, one dominant (pressure), and a misconception corrected that sets up the case study.
Delivery: the misconception box is the pivot of the whole talk — deliver it as a small revelation.""")

# =========================================================
# 13 — CASE STUDY (chart B)  (CHECKPOINT 2)
# =========================================================
s = blank_slide()
header(s, "PART 3 · FIELD EVIDENCE", "GoM case study: injection placement above the onset left 3,400 ft unprotected", 13, 55, CLK[13])
s.shapes.add_picture(os.path.join(FIGS, "depth_prediction.png"), Inches(4.62), Inches(1.80),
                     width=Inches(8.15), height=Inches(4.46))
tf = box(s, 0.62, 1.85, 3.75, 4.9)
items = [
    ("The well  (all depths labelled on the figure):", "deepwater GoM; perfs 17,700–18,000 ft; mandrel 14,600 ft [4]."),
    ("The problem:", "3,400 ft of tubing BELOW the mandrel had no chemical protection."),
    ("The model:", "effective tubing ID fell from 2.3 in to 1.35 in while producing."),
    ("The fix:", "solvent soak, then continuous inhibitor — 750 ppm, stabilised near 500 ppm."),
    ("The result:", "+16 % production over 4 months, including shut-in time [4]."),
    ("NEXT  →  what you can actually do about it.", "", TEAL),
]
for i, (a, b, *rest) in enumerate(items):
    col = rest[0] if rest else NAVY
    if b == "":
        par(tf, a, size=12, italic=True, color=col, space_after=0)
    else:
        par(tf, "▪  " + a, size=12.5, bold=True, color=col, first=(i == 0), space_after=2)
        par(tf, b, size=11.5, color=TEXT, space_after=6, line=1.03)
tf = box(s, 0.62, 6.55, 12.1, 0.45)
par(tf, "Reconstructed from published data: geometry (perforations 17,700–18,000 ft; mandrel 14,600 ft) from Wylde & Punase, JPT (2020) [4]. The P–T model itself is illustrative — the thesis rebuilds it with the field's own PVT data.",
    size=10, italic=True, color=GRAY, first=True, space_after=0, line=1.05)
notes(s, """TARGET 55 s (clock 8:00) — CHECKPOINT 2: on this slide at 8:00.
VERIFIED — Ref [4] = Wylde, J.; Punase, A. "Asphaltenes: A Complex and Challenging Flow Assurance Issue To Measure and Quantify Risk," JPT, 30 April 2020 (Clariant Oil Services). Verbatim: "perforated between 17,700 and 18,000 ft, while the downhole chemical injection mandrel was situated at 14,600 ft, resulting in 3,400 ft of unprotected tubing"; "solvent soak for 2 days"; "continuous AI-1 injection at a high dose rate of 750 ppm"; "a decrease in the readings to around 21 at 500-ppm dosage"; "16% overall increase in production (including overall shut-in time) over a 4-month period." The 2.3 to 1.35 in ID reduction comes from their transient multiphase flow model.
POINTING CUE: while you say the third bullet, point to the ≈300 m bracket on the figure — the gap between the mandrel (4450 m) and the predicted onset (4752 m). That bracket IS the whole argument of the slide.
Delivery: the punchline is geometric — the well HAD inhibition, but 3,400 ft of tubing sat below it. Prediction of onset depth is not academic; it is the difference between protected and unprotected steel.""")

# =========================================================
# 14 — DETECTION & SURVEILLANCE
# =========================================================
s = blank_slide()
header(s, "PART 4 · DEFENSES", "Detecting the problem before the well chokes", 14, 35, CLK[14])
tf = box(s, 0.62, 1.80, 12.1, 4.9)
items = [
    ("Rate and pressure:", "a rising ΔP for the same rate (ΔP/Q) is the earliest fingerprint of a growing deposit."),
    ("Wellhead signature:", "severe fields lose 20–25 % of wellhead pressure in 15–20 days — a clean, quantitative alarm [3]."),
    ("Wireline gauge rings / calipers:", "successive runs map the deposit profile and volume, so treatment volumes can be sized [3]."),
    ("Sampling:", "optical analysis of produced fluids tuned the inhibitor dose in near-real time in the case well [4]."),
    ("Monitoring + periodic CT cleanouts:", "deliberate surveillance where prevention alone is not trusted [4]."),
]
for i, (a, b) in enumerate(items):
    p = par(tf, "▪  " + a, size=14.5, bold=True, color=NAVY, first=(i == 0), space_after=2)
    par(tf, b, size=13.5, color=TEXT, space_after=11, line=1.05)
notes(s, """TARGET 35 s (clock 8:35).
Key message: surveillance is cheap compared with a surprise plugging — and it produces the data your prediction model needs to be validated.
Delivery: name the ΔP/Q fingerprint and the gauge-ring method; both are real field practice.""")

# =========================================================
# 15 — REACTIVE DEFENSES & WHY THEY FAIL
# =========================================================
s = blank_slide()
header(s, "PART 4 · DEFENSES", "Reactive cleanouts are temporary, expensive — and guarantee repeat NPT", 15, 45, CLK[15])
data = [
    ["Method", "How it works", "Why it is not the answer"],
    ["Wireline scrapers / cutting", "mechanical removal of the deposit from tubing", "slow, impractical in bad wells; routine scraping was 'expensive and cumbersome' at Hassi Messaoud [3]"],
    ["Coiled-tubing milling / jetting", "mechanical cleanout of the full string", "rig-up + NPT; the well is shut in while it happens"],
    ["Aromatic solvent washes (toluene / xylene)", "dissolves the deposit; volume calculated from gauge-ring profile", "HSE exposure (BTEX), flammability, cost; only a temporary fix — deposits return"],
    ["Acid / stimulation jobs", "dissolve scale around the near-wellbore", "can destabilise the crude and worsen asphaltene damage — used with great care"],
]
table(s, data, 0.62, 1.74, 12.1, col_w=[3.3, 4.2, 5.4], fs=12, hfs=12.5, row_h=0.74)
tf = box(s, 0.62, 5.66, 12.1, 1.15)
par(tf, "Thermal methods are not a melting story — asphaltene deposits have no melting point, and heat can even weaken the resins that keep them dispersed. That is the wax toolkit, not this one [15].",
    size=12.5, italic=True, color=GRAY, first=True, space_after=7, line=1.06)
par(tf, "Every reactive method is a cleanout — and every cleanout is NPT, deferred production and a repeat visit scheduled by the physics itself.",
    size=13.5, bold=True, color=NAVY, space_after=0)

notes(s, """TARGET 45 s (clock 9:20).
VERIFIED — Ref [3] = Haskett & Tartera, SPE-994-PA (1965): verbatim, 'This method, while satisfactory for cleaning the tubings, was expensive and cumbersome.' Deposit composition: 83.4% asphaltenes, 3.3% carbenes, 13.3% resins. Still the canonical field description.
Delivery: do not read all five rows — wireline, solvents at Hassi Messaoud, and the acid warning are enough.""")

# =========================================================
# 16 — PROACTIVE: CONTINUOUS INHIBITION
# =========================================================
s = blank_slide()
header(s, "PART 4 · DEFENSES", "The proactive answer: keep the asphaltenes dispersed — don't let them land", 16, 40, CLK[16])
tf = box(s, 0.62, 1.80, 7.25, 4.9)
items = [
    ("Chemical inhibitors:", "polymer/surfactant molecules adsorb onto the aggregates — polar head on the particle, alkyl tail into the oil — and keep them suspended. Inhibitors shift the flocculation onset; dispersants shrink the aggregates. Both prevent formation; neither dissolves an existing deposit [15, 16]."),
    ("Delivery — capillary injection string:", "thin tubing clamped to the production string, dosing continuously at depth with a surface pump."),
    ("Dosage is field-specific:", "hundreds to a few thousand ppm; here 750 ppm initial, stabilised ≈ 500 ppm [4]."),
    ("Placement rule (the lesson of Part 3):", "the injection point must sit BELOW the deepest credible onset depth — onset plus a margin sized by the sensitivity study (backup slide 24). Rules of thumb leave tubing unprotected."),
    ("Chemical-free alternative:", "re-injecting dead oil restores solvent quality — promising, still developing [7]."),
]
for i, (a, b) in enumerate(items):
    p = par(tf, "▪  " + a, size=13.5, bold=True, color=NAVY, first=(i == 0), space_after=2)
    par(tf, b, size=12.5, color=TEXT, space_after=7, line=1.03)
s.shapes.add_picture(os.path.join(FIGS, "well_schematic.png"), Inches(8.10), Inches(1.80),
                     width=Inches(4.60), height=Inches(3.09))
r, t = rect(s, 8.10, 4.98, 4.62, 1.65, fill=LIGHT, line=TEAL, lw=1.4,
            shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
par(t, "WHY IT PAYS", size=12, bold=True, color=TEAL, first=True, space_after=5)
par(t, "No rig-up, no deferred production, no repeat workover. The well keeps producing while the problem is managed continuously — and the schematic shows exactly where the chemical has to arrive.",
    size=11.5, color=TEXT, space_after=0, line=1.06)
notes(s, """TARGET 40 s (clock 10:00).
VERIFIED — Ref [7] = Khaleel, A.T.; Abutaqiya, M.I.L.; Sisco, C.J.; Vargas, F.M. Mitigation of Asphaltene Deposition by Re-injection of Dead Oil. Fluid Phase Equilibria 2020, 514, 112552 (correct journal: Fluid Phase Equilibria, not Fuel). Ref [4] = dosage numbers from the deepwater case (Wylde & Punase, JPT 2020).
Delivery: the placement rule is the bridge from the case study to your project — say it as a design principle, not an anecdote.""")

# =========================================================
# 17 — ECONOMICS  (CHECKPOINT 3)
# =========================================================
s = blank_slide()
header(s, "PART 4 · DEFENSES", "The business case: reactive cleanouts vs continuous inhibition", 17, 40, CLK[17])
data = [
    ["Cost element", "Reactive approach", "Proactive approach"],
    ["Routine cost", "treatment documented up to ≈ $3 M/well — Stratiev et al. (2025) [2]", "inhibitor: $31–46 k/well/yr (Middle East) to $330–390 k/well/yr (GoM) — Cenegy (2001) survey, via Khaleel et al. (2020) [8, 9]"],
    ["Deferred production", "up to ≈ $1.2 M/day during a plugging/shut-in event — Stratiev et al. (2025) [2]", "near zero — the well keeps producing"],
    ["Catastrophic tail", "≈ $70 M/well shut-in + cleanup (GoM); up to $100 M if the deposit forms at the SCSSV — Farooq et al. (2021); Elsevier monograph (2021) [1, 10]", "residual risk only if dosage/placement is wrong"],
    ["Frequency", "scheduled by the physics — deposits return", "managed continuously; occasional CT verification runs"],
]
table(s, data, 0.62, 1.78, 12.1, col_w=[2.6, 4.9, 4.6], fs=11.5, hfs=12, row_h=0.72)
r, t = rect(s, 0.62, 5.52, 12.1, 1.40, fill=LIGHT, line=AMBER, lw=1.3, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.10)
t.margin_left = Inches(0.95)
s.shapes.add_picture(os.path.join(FIGS, "icon_money.png"), Inches(0.80), Inches(5.86),
                     width=Inches(0.58), height=Inches(0.58))
par(t, "The thesis supplies the missing breakeven frequency for this specific well — the cleanout rate at which continuous inhibition pays for itself. Today that number is argued by rule of thumb.",
    size=13, bold=True, color=NAVY, first=True, space_after=3)
par(t, "Attribution: the inhibitor ranges trace to Cenegy's 2001 survey; Fluid Phase Equilibria (2020) and the Elsevier monograph (2021) assign the high range to the GoM, Energy & Fuels (2021) to the Middle East — I quote the majority and flag the discrepancy.",
    size=11, italic=True, color=GRAY, space_after=0)
notes(s, """TARGET 40 s (clock 10:40) — CHECKPOINT 3: on this slide at 10:40.
VERIFIED: [1] Farooq et al. 2021 ($70M/well GoM); [2] Stratiev et al. 2025 ($1.2M/day; treatment up to $3M); [8] Khaleel et al., Fluid Phase Equilibria 2020 ($31-46k ME; $330-390k GoM, citing Cenegy 2001); [10] Elsevier monograph 2021 (up to $100M if at SCSSV; removal $0.3-3.5M/well; sidetrack ~$50M; downtime ~$700k/day). If asked about the attribution conflict: see Q26 of the Q&A pack.
Delivery: the honest caveat at the bottom builds more credibility than any single number.""")

# =========================================================
# 18 — GAP ANALYSIS
# =========================================================
s = blank_slide()
header(s, "PART 5 · TOWARD THE THESIS", "Gap analysis: prediction is still the weak link", 18, 40, CLK[18])
tf = box(s, 0.62, 1.80, 8.1, 4.9)
items = [
    "Onset prediction is sensitive to fluid characterisation — small input errors move the AOP and therefore the predicted depth.",
    "Precipitation ≠ deposition: transport/adhesion models are far less mature than the thermodynamics.",
    "Inhibitor dosage is still chosen largely by trial and error — published criticism: lab validation methods are inadequate, and field results can disappoint [7].",
    "Injection-point placement is rarely optimised against a predicted onset depth — the case study showed 3,400 ft left unprotected [4].",
    "No consolidated, public workflow connects screening → onset curve → deposition depth → injection point → dosage → economics.",
]
for i, txt in enumerate(items):
    par(tf, "▪  " + txt, size=13.5, first=(i == 0), space_after=10, line=1.05)
r, t = rect(s, 8.95, 2.10, 3.77, 3.4, fill=LIGHT, line=TEAL, lw=1.5, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
par(t, "THE OPPORTUNITY", size=12.5, bold=True, color=TEAL, first=True, space_after=8)
par(t, "A transparent, auditable workflow — every assumption visible, every number reproducible — applied end-to-end on a documented well. At Bachelor's level, documented rigour is the contribution.",
    size=13, color=TEXT, space_after=0, line=1.12)
notes(s, """TARGET 40 s (clock 11:20).
VERIFIED — Ref [7] = Khaleel et al., Fluid Phase Equilibria 2020, verbatim: 'the laboratory methods used to validate the effectiveness of these chemicals are inadequate and the chemicals are often ineffective or actually worsen the problem when applied in the field.'
Delivery: this is the pivot from review to research. Say the last box slowly.""")

# =========================================================
# 19 — PROPOSED PROJECT
# =========================================================
s = blank_slide()
header(s, "PART 5 · TOWARD THE THESIS", "Graduation project: onset and deposition-depth prediction for a real well", 19, 45, CLK[19])
r, t = rect(s, 0.62, 1.80, 12.1, 1.05, fill=LIGHT, line=NAVY2, lw=1.2, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.09)
par(t, "AIM:  build and document a transparent screening → PVT → wellbore workflow that predicts asphaltene onset depth, selects the inhibitor injection point, and quantifies the economics — validated against a documented field case.",
    size=13.5, bold=True, color=NAVY, first=True, space_after=0, line=1.1)
r, t = rect(s, 0.62, 3.00, 5.95, 3.1, fill=WHITE, line=TEAL, lw=1.4, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
par(t, "RESEARCH QUESTIONS", size=13, bold=True, color=TEAL, first=True, space_after=6)
for line in [
    "RQ1  How sensitive is predicted onset depth to fluid characterisation (SARA/CII inputs, EoS choice, AOP data)?",
    "RQ2  What injection point and dosage window keep the whole tubing above the AOP — with what uncertainty margin?",
    "RQ3  At what cleanout frequency does continuous inhibition break even for the case well?",
]:
    par(t, line, size=13, color=TEXT, space_after=9, line=1.05)
r, t = rect(s, 6.80, 3.00, 5.92, 3.1, fill=WHITE, line=AMBER, lw=1.4, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
par(t, "WORK PACKAGES", size=13, bold=True, color=AMB_D, first=True, space_after=6)
for line in [
    "WP1  Fluid screening: SARA, CII, de Boer assessment of the case fluid",
    "WP2  EoS modelling: asphaltene onset curve (CPA / PC-SAFT class model), tuned to measured AOP",
    "WP3  Wellbore model: P–T profile → onset depth → injection-point design",
    "WP4  Economics: breakeven frequency and reactive-vs-proactive comparison",
]:
    par(t, line, size=12.5, color=TEXT, space_after=7, line=1.05)
notes(s, """TARGET 45 s (clock 12:05).
Delivery: read the AIM; summarise RQs as three one-liners; WP1–WP4 map to the methodology slide.
Confirm with your supervisor before the seminar that a modelling-based thesis will be accepted — his agreement is your best Q&A shield.""")

# =========================================================
# 20 — METHODOLOGY
# =========================================================
s = blank_slide()
header(s, "PART 5 · TOWARD THE THESIS", "Methodology: screening → modelling → validation → economics", 20, 50, CLK[20])
steps = [
    ("1 · SCREEN", ["SARA + CII + de Boer on published fluid data", "Output: stability verdict and risk flag"], TEAL),
    ("2 · MODEL", ["Asphaltene onset curve vs temperature (CPA/PC-SAFT-class EoS)", "Tune to any published/measured AOP", "Output: the envelope"], AMB_D),
    ("3 · PREDICT", ["Wellbore P–T profile → crossing depth", "Sensitivity: ±ΔP, ±AOP, ±GOR", "Output: onset depth + margin"], NAVY2),
    ("4 · DECIDE", ["Injection point + dosage window", "Breakeven cleanout-frequency model", "Output: decision matrix for the field"], RED),
]
x = 0.62
for lab, lines, col in steps:
    r, t = rect(s, x, 1.85, 2.94, 4.15, fill=WHITE, line=col, lw=1.6, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
    par(t, lab, size=14.5, bold=True, color=col, first=True, space_after=10)
    for ln in lines:
        par(t, "▪  " + ln, size=12, color=TEXT, space_after=9, line=1.06)
    x += 3.10
chips = [
    ("1 · LICENSED, if available", "PVTsim · Multiflash · OLGA / PIPESIM", NAVY2),
    ("2 · OPEN-SOURCE FALLBACK", "NeqSim (CPA-EoS) + Python — verified", TEAL),
    ("3 · FLOOR", "correlation-based screening", AMB_D),
]
cx = 0.62
for lab, sub, col in chips:
    r, t = rect(s, cx, 6.08, 3.84, 0.56, fill=WHITE, line=col, lw=1.4,
                shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.12)
    parruns(t, [(lab + "   ", True, col), (sub, False, GRAY)], size=11, first=True, space_after=0)
    if cx < 8.0:
        ab = box(s, cx + 3.87, 6.14, 0.24, 0.4)
        par(ab, "→", size=16, bold=True, color=GRAY, first=True, space_after=0, align=PP_ALIGN.CENTER)
    cx += 4.11
tf = box(s, 0.62, 6.66, 12.1, 0.28)
par(tf, "TOOL HIERARCHY — the project cannot stall on software: licence check is a Week-1 action, and route 2 is already verified.",
    size=10.5, italic=True, color=GRAY, first=True, space_after=0)
notes(s, """TARGET 50 s (clock 12:55).
Key message: four steps with outputs — and a declared tool hierarchy so no licence situation can kill the project.
Validation (say it if time allows): reproduce the documented case's observed behaviour — onset deeper than the existing mandrel, ID reduction, dosage window — within a stated tolerance; cross-check onset with NeqSim and correlations.
ACTION (Week 1): confirm which licensed tools the department has; NeqSim fallback verified as available (open source, Apache-2.0).""")

# =========================================================
# 21 — TIMELINE, DELIVERABLES, LIMITS
# =========================================================
s = blank_slide()
header(s, "PART 5 · TOWARD THE THESIS", "Deliverables, honest limitations, and a realistic 16-week plan", 21, 40, CLK[21])
r, t = rect(s, 0.62, 1.85, 3.75, 4.3, fill=WHITE, line=TEAL, lw=1.4, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
par(t, "DELIVERABLES", size=13, bold=True, color=TEAL, first=True, space_after=8)
for ln in ["Thesis with full calculation chain", "Documented prediction workflow (spreadsheet / Python)", "Case-study report: onset depth and injection point", "Breakeven economics model", "Decision matrix: monitor / inhibit / clean out"]:
    par(t, "▪  " + ln, size=12, color=TEXT, space_after=8, line=1.05)
r, t = rect(s, 4.55, 1.85, 3.75, 4.3, fill=WHITE, line=RED, lw=1.4, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
par(t, "LIMITATIONS & ETHICS (stated up-front)", size=13, bold=True, color=RED, first=True, space_after=8)
for ln in ["Literature-based fluids — no live sample unless a partner provides one", "Single documented case: conditional conclusions", "Deposition-transport modelling kept first-order", "Steady-state wellbore profiles", "Cost figures indicative, from published ranges", "Solvent chemistries compared on paper only — no uncontrolled BTEX handling, no vendor recommended"]:
    par(t, "▪  " + ln, size=12, color=TEXT, space_after=8, line=1.05)
data = [
    ["Weeks", "Milestone"],
    ["1 – 3", "Literature, standards, case selection"],
    ["4 – 6", "Screen fluid; build the envelope model"],
    ["7 – 9", "Wellbore model; onset-depth prediction"],
    ["10 – 11", "Validation against the case; sensitivity"],
    ["12 – 14", "Economics + thesis writing"],
    ["15 – 16", "Review, defence preparation"],
]
table(s, data, 8.50, 1.85, 4.2, col_w=[1.3, 2.9], fs=11, hfs=11, row_h=0.40)
tf = box(s, 0.62, 6.25, 12.1, 0.34)
par(tf, "NEXT  →  three messages to take away.", size=12, italic=True, color=TEAL,
    first=True, space_after=0)
notes(s, """TARGET 40 s (clock 13:35).
Delivery: the limitations list is a strength — deliver it confidently, as a scope contract rather than an apology.""")

# =========================================================
# 22 — THREE MESSAGES + THANK YOU
# =========================================================
s = blank_slide()
header(s, "CLOSING", "Three messages — then your questions", 22, 40, CLK[22])
msgs = [
    ("1", "Asphaltene plugging is a pressure-and-composition story, not a cooling story — and precipitation is not the same as deposition.", RED),
    ("2", "The decisive number is the depth where the flowing path crosses the onset envelope: it sets the injection point, the mitigation strategy and the money.", AMBER),
    ("3", "My graduation project builds a transparent screening-to-economics workflow for that depth — validated on a documented field case.", TEAL),
]
for i, (num, txt, col) in enumerate(msgs):
    yy = 1.90 + i * 1.55
    rect(s, 0.62, yy, 12.1, 1.35, fill=LIGHT, line=None, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.10)
for i, (num, txt, col) in enumerate(msgs):
    yy = 1.90 + i * 1.55
    nb = box(s, 0.95, yy + 0.30, 0.9, 0.9)
    par(nb, num, size=38, bold=True, color=col, first=True, space_after=0, align=PP_ALIGN.CENTER)
    tb = box(s, 1.95, yy + 0.18, 10.5, 1.05)
    par(tb, txt, size=15, color=NAVY, first=True, space_after=0, line=1.1)
r, t = rect(s, 0.62, 6.35, 12.1, 0.55, fill=LIGHT2, line=None, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.18)
par(t, "Thank you — I am happy to take your questions.", size=14, bold=True, color=NAVY,
    first=True, space_after=0, align=PP_ALIGN.CENTER)
notes(s, """TARGET 40 s (clock 14:15) — CONCLUSIONS START HERE, no exceptions.
SOUNDBITES: rehearse each numbered message as a standalone 15-second answer — one idea, one example, one number — and practise them OUT OF ORDER, because an examiner may ask for any single one.
Delivery: one message per breath, pause between them; finish with the thank-you line, then STOP TALKING.
Never cut this slide, even if everything before it had to be compressed.""")

# =========================================================
# 23 — REFERENCES
# =========================================================
s = blank_slide()
header(s, "CLOSING", "Key references (full list per course citation style)", 23, 10, CLK[23])
left = [
    "[1]  Farooq, U.; Lædre, S.; Gawel, K. Review of Asphaltenes in an Electric Field. Energy & Fuels 2021, 35 (9), 7285–7304. doi:10.1021/acs.energyfuels.0c03962",
    "[2]  Stratiev, D.; Nikolova, R.; Veli, A.; Shishkova, I.; Toteva, V.; Georgiev, G. Mitigation of Asphaltene Deposit Formation via Chemical Additives: A Review. Processes 2025, 13 (1), 141. doi:10.3390/pr13010141",
    "[3]  Haskett, C. E.; Tartera, M. A Practical Solution to the Problem of Asphaltene Deposits — Hassi Messaoud Field, Algeria. J. Pet. Technol. 1965, 17 (4), 387–391. SPE-994-PA. doi:10.2118/994-PA",
    "[4]  Wylde, J.; Punase, A. Asphaltenes: A Complex and Challenging Flow Assurance Issue To Measure and Quantify Risk. J. Pet. Technol., 30 Apr 2020 (deepwater GoM case study).",
    "[5]  Yen, A.; Yin, Y. R.; Asomaning, S. Evaluating Asphaltene Inhibitors: Laboratory Tests and Field Studies. SPE 65376, 2001. Thresholds after Asomaning (2003): CII ≥ 0.9 unstable; < 0.7 stable.",
    "[6]  de Boer, R. B.; Leerlooyer, K.; Eigner, M. R. P.; van Bergen, A. R. D. Screening of Crude Oils for Asphalt Precipitation: Theory, Practice and the Selection of Inhibitors. SPE Prod. Facil. 1995, 10 (1), 55–61. doi:10.2118/24987-PA",
    "[7]  Khaleel, A. T.; Abutaqiya, M. I. L.; Sisco, C. J.; Vargas, F. M. Mitigation of Asphaltene Deposition by Re-injection of Dead Oil. Fluid Phase Equilib. 2020, 514, 112552. doi:10.1016/j.fluid.2020.112552",
]
right = [
    "[8]  Cenegy, L. M. Survey of Successful World-Wide Asphaltene Inhibitor Treatments in Oil Production Fields. SPE ATCE, New Orleans, 2001 (original source of the inhibitor cost data).",
    "[9]  Asphaltene Deposition Control by Chemical Inhibitors, 1st ed.; Gulf Professional/Elsevier: 2021; Ch. 1 (cost table: removal $0.3–3.5 M/well; sidetrack ≈ $50 M; downtime ≈ $700 k/day).",
    "[10] Critical Analysis of Different Techniques Used To Screen Asphaltene Stability in Crude Oils. Fuel 2021, 299, 120839. [Author list to confirm at library]",
    "[11] Mullins, O. C. The Modified Yen Model. Energy & Fuels 2010, 24 (4), 2179–2207; reviewed in Mullins et al., Energy & Fuels 2012. doi:10.1021/ef300185p",
    "[12] ASTM D6560-22 / IP 143:21 (asphaltenes, n-heptane insolubles); D2007 (clay-gel SARA); D4124 (four fractions); D3279 (n-heptane insolubles); D6703 (automated Heithaus titrimetry).",
    "[13] NeqSim — open-source thermodynamic toolkit (CPA-EoS; asphaltene screening). Equinor, Apache-2.0 licence, v3.22. github.com/equinor/neqsim",
    "[14] Evaluation and Modeling of Asphaltene Deposition in Oil Wells. SPE-206366-MS, SPE ATCE, Dubai, 2021. doi:10.2118/206366-MS",
    "[15] Mohammed, I.; Mahmoud, M.; Al Shehri, D.; El-Husseini, A.; Alade, O. Asphaltene Precipitation and Deposition: A Critical Review. J. Pet. Sci. Eng. 2021, 197, 107956 (thermal/steam limits; inhibitor mechanisms). doi:10.1016/j.petrol.2020.107956",
    "[16] Kelland, M. A. Production Chemicals for the Oil and Gas Industry. CRC Press, 2009 (inhibitor-versus-dispersant distinction; accessed via AADE-24-FTCE-073).",
]
for lines, x in ((left, 0.62), (right, 6.85)):
    tfx = box(s, x, 1.80, 6.0, 5.0)
    for i, ln in enumerate(lines):
        par(tfx, ln, size=9.5, first=(i == 0), space_after=6, line=1.02)
tf = box(s, 0.62, 6.35, 12.1, 0.6)
par(tf, "Full citation details, exact quotes and page references: 04_Source_Log.docx. Technical cross-check, risk register and sensitivity study: 06_Technical_Accuracy_and_Risk_Register.docx. Keyed to every number in the deck.",
    size=11.5, italic=True, color=GRAY, first=True, space_after=0)
notes(s, """TARGET 10 s (clock 14:25).
Delivery: do not read. One sentence: "My references are on this slide and in the handout; exact quotes and page references are in my source log." [Hand 04_Source_Log.docx to the examiner if any number is challenged.] If behind schedule, skip entirely.""")

# =========================================================
# 24 — BACKUP 1
# =========================================================
s = blank_slide()
header(s, "BACKUP", "Derivation and assumptions: envelope, indices and the onset-depth model", 24, backup=True)
tf = box(s, 0.62, 1.80, 12.1, 5.0)
for i, (a, b) in enumerate([
    ("CII definition:", "CII = (Saturates + Asphaltenes) / (Resins + Aromatics); boundary ≈ 0.9 unstable, < 0.7 stable — thresholds are laboratory-specific."),
    ("De Boer plot (correctly stated):", "in-situ density (x) vs gas under-saturation, P_res − P_b (y); three risk zones; risk greatest for LIGHT, undersaturated crudes. The onset-above-Pb rule is a SEPARATE envelope screen."),
    ("Onset curve and depth model:", "AOP(T) is linear here for illustration — the thesis tunes a CPA / PC-SAFT-class EoS to measured AOP. Onset depth = where P(d) falls below AOP(T(d)); steady-state profiles here, multiphase in the thesis."),
    ("Sensitivity study (quantified):", "base 4,752 m. ±50 bar → ±255 m · ±100 bar → ±511 m · ±10 % gradient → −404 / +487 m · gas-lift (gradient −10 %, AOP +50 bar) → 5,520 m: the whole tubing unstable, so the DOSE does the work. Design rule: inject below the deepest credible onset."),
    ("Deliberately excluded:", "deposition growth (first-order only), lower-AOP phenomena, near-wellbore damage."),
]):
    p = par(tf, "▪  " + a, size=13, bold=True, color=NAVY, first=(i == 0), space_after=2)
    par(tf, b, size=12, color=TEXT, space_after=9, line=1.04)
notes(s, "Show only if asked for the maths. The sensitivity line is the most valuable one: it justifies the safety margin on the injection point.")

# =========================================================
# 25 — BACKUP 2
# =========================================================
s = blank_slide()
header(s, "BACKUP", "Data sources, tools and validation plan", 25, backup=True)
tf = box(s, 0.62, 1.80, 12.1, 5.0)
for i, (a, b) in enumerate([
    ("Fluid data:", "published PVT/SARA reports from case-study literature; field case in Wylde & Punase, JPT (2020); Hassi Messaoud classic data (SPE-994-PA); other published well databases via the university library."),
    ("Licensed tools (if available):", "PVT simulator with asphaltene onset (PVTsim / Multiflash / WinProp-class) and a wellbore flow model (OLGA with asphaltene module / PIPESIM). Confirm with the department in Week 1."),
    ("Open-source fallback:", "NeqSim (Equinor, Apache-2.0) — CPA-EoS asphaltene screening and PVT; Python for the workflow, plots and the economics model."),
    ("Validation targets:", "reproduce the documented case behaviour — predicted onset below the existing mandrel; ID-reduction trend; dosage within the field's 750→500 ppm window — within a stated tolerance; cross-check onset against correlations."),
    ("Risk register (top 3):", "(1) no live fluid sample → use published data and declare it; (2) licensed software unavailable → NeqSim + correlations (project already designed for this); (3) case-data conflicts across sources → document the variance and choose conservative values."),
    ("Ethics & HSE framing:", "aromatic solvents (BTEX) are a workplace-safety topic; the project compares alternatives explicitly and never recommends handling without controls."),
]):
    p = par(tf, "▪  " + a, size=13.5, bold=True, color=NAVY, first=(i == 0), space_after=2)
    par(tf, b, size=12.5, color=TEXT, space_after=10, line=1.04)
notes(s, "Show only if asked about data, tools or risk. Naming NeqSim and the exact licence fallback is what pre-empts the 'software access' objection.")

LAYOUT_OF = {1:"A-title", 2:"B-cards", 3:"C-split", 4:"C-split", 5:"B-cards",
             6:"C-split", 7:"C-split", 8:"D-table", 9:"B-cards", 10:"C-split",
             11:"D-table", 12:"C-split", 13:"C-split", 14:"C-split", 15:"D-table",
             16:"C-split", 17:"D-table", 18:"C-split", 19:"C-split", 20:"B-cards",
             21:"D-table", 22:"F-closing", 23:"E-fullwidth", 24:"G-backup", 25:"G-backup"}
MARGIN_L, MARGIN_R = 0.62, 12.88   # content band; card grids + titles end at 12.86
_pics = sum(1 for sl in prs.slides for sh in sl.shapes if sh.shape_type == 13)
print("layout map:", len(LAYOUT_OF), "slides | archetypes used:",
      sorted(set(LAYOUT_OF.values())))

prs.core_properties.title = "Asphaltenes: Predicting and Preventing Wellbore Plugging — Seminar"
prs.core_properties.author = "Prepared for [Your name] — Seminar, BSc Petroleum and Gas Engineering"
out = "/home/user/Seminar_Asphaltenes/Asphaltene_Slides_23plus2.pptx"
prs.save(out)
print("saved", out, "| slides:", len(prs.slides._sldIdLst))
