#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Builds 06_Technical_Accuracy_and_Risk_Register.docx
   Part A — verification of the seminar's technical claims against original sources
   Part B — corrections applied to the deck / script / Q&A as a result
   Part C — project & presentation risk register
   Part D — quantified sensitivity study (backup slide 24)
   Part E — HSE / ethics position
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

NAVY = RGBColor(0x10, 0x2A, 0x43); TEAL = RGBColor(0x1F, 0x7A, 0x84)
GRAY = RGBColor(0x5A, 0x66, 0x72); RED = RGBColor(0xB4, 0x3B, 0x2E); AMBER = RGBColor(0xA8, 0x6B, 0x00)

doc = Document()
st = doc.styles["Normal"]; st.font.name = "Calibri"; st.font.size = Pt(9)
st.paragraph_format.space_after = Pt(4); st.paragraph_format.line_spacing = 1.0
for name, size, color in (("Heading 1", 13, NAVY), ("Heading 2", 11, TEAL)):
    sty = doc.styles[name]; sty.font.name = "Calibri"; sty.font.size = Pt(size); sty.font.bold = True
    sty.font.color.rgb = color; sty.paragraph_format.space_before = Pt(8); sty.paragraph_format.space_after = Pt(3)
sec = doc.sections[0]
sec.top_margin = Cm(1.4); sec.bottom_margin = Cm(1.4); sec.left_margin = Cm(1.4); sec.right_margin = Cm(1.4)

def para(text, size=9, bold=False, italic=False, color=None, space_after=4):
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(text); r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic
    if color: r.font.color.rgb = color
    return p

def rich(runs, size=9, space_after=4):
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(space_after)
    for text, bold, color in runs:
        r = p.add_run(text); r.font.size = Pt(size); r.font.bold = bold
        if color: r.font.color.rgb = color
    return p

def shade(cell, hexc):
    tcPr = cell._tc.get_or_add_tcPr(); shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hexc); tcPr.append(shd)

def tbl(rows, widths, fs=7.6, highlight_col=None, highlight_map=None):
    t = doc.add_table(rows=len(rows), cols=len(rows[0])); t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = t.cell(ri, ci); cell.text = ""
            p = cell.paragraphs[0]; p.paragraph_format.space_after = Pt(1); p.paragraph_format.line_spacing = 1.0
            r = p.add_run(str(val)); r.font.size = Pt(fs); r.font.name = "Calibri"
            if ri == 0:
                r.font.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); shade(cell, "102A43")
            else:
                if ri % 2 == 0:
                    shade(cell, "F4F7F9")
                if highlight_col is not None and ci == highlight_col and highlight_map:
                    key = str(val).strip().upper()
                    if "CONFIRMED" in key: r.font.color.rgb = TEAL; r.font.bold = True
                    elif "CORRECT" in key: r.font.color.rgb = RED; r.font.bold = True
                    elif "REFINE" in key: r.font.color.rgb = AMBER; r.font.bold = True
    total = sum(widths)
    for ci, w in enumerate(widths):
        for ri in range(len(rows)):
            t.cell(ri, ci).width = Cm(18.4 * w / total)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)
    return t

# ---------------- header ----------------
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("TECHNICAL ACCURACY CROSS-CHECK & RISK REGISTER"); r.font.size = Pt(11); r.font.bold = True; r.font.color.rgb = TEAL
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Asphaltenes: Predicting and Preventing Wellbore Plugging"); r.font.size = Pt(15); r.font.bold = True; r.font.color.rgb = NAVY
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Companion to 04_Source_Log.docx. Part A re-checks every load-bearing technical claim against the original paper; "
              "Part B lists what was changed as a result; Parts C–E are the risk, sensitivity and HSE/ethics records.")
r.font.size = Pt(9); r.font.italic = True; r.font.color.rgb = GRAY
para("Verdict key:   CONFIRMED = the claim as stated in the pack is supported by the original source.   "
     "REFINED = supported but imprecise as first written; wording tightened.   "
     "CORRECTED = the pack said something the original source does not support; wording replaced.",
     size=8, italic=True, color=AMBER, space_after=6)

# ================= PART A =================
doc.add_heading("Part A — Verification register (claim → verdict → evidence → action)", level=1)
rows = [["#", "Claim in the pack", "Verdict", "Evidence from the original source", "Action taken"]]
data = [
 ("A1", "CII = (Saturates + Asphaltenes) / (Resins + Aromatics); ≥ 0.9 unstable, 0.7–0.9 uncertain, < 0.7 stable [slide 7, Q5]",
  "CONFIRMED",
  "Formula and threshold scale reproduced identically across the indexing literature, e.g. “the sum of asphaltene and saturates (flocculants) divided by the sum of aromatics and resins (peptizers)”; " 
  "“if CII ≥ 0.9 … unstable; 0.7 ≤ CII ≤ 0.9 … uncertain; CII ≤ 0.7 … stable.” Origin credited to Yen et al.; thresholds from Asomaning (2003).",
  "Slide 7 now names the origin: “CII [5] (Yen et al.; thresholds after Asomaning)”. Caveat retained."),
 ("A2", "CII thresholds are laboratory-specific and CII alone is not a verdict [slide 7 notes, Q5, Q22]",
  "CONFIRMED",
  "Asomaning’s evaluation was based on oils characterised by two different SARA routes (liquid chromatography ASTM D2007 and IP 143), and later work (ACS Omega 2020) found CII thresholds “greatly differ from real field results” for some crudes. "
  "A widely-cited textbook variant even uses CII > 1 as the flag — the scale is not universal.",
  "Kept as a stated caveat on slide 7. CII is presented as ONE of three screens, never as a pass/fail."),
 ("A3", "“De Boer screening: plot AOP against bubble point …” [backup slide 24] and “if the onset is above the bubble point, expect problems — that is the de Boer criterion” [slide 10 script]",
  "CORRECTED",
  "The de Boer plot is a cross-plot of IN-SITU OIL DENSITY (x) against the DEGREE OF GAS UNDER-SATURATION, i.e. (reservoir pressure − bubble-point pressure) (y), with three risk zones (severe / moderate / minimal). "
  "de Boer et al. (1995) conclude that risk is greatest for LIGHT, gas-undersaturated crudes with low asphaltene content. It is not an AOP-versus-Pb cross-plot.",
  "Slide 10, the script, backup slide 24 and Q6 rewritten. The AOP-above-Pb observation is retained as a separate envelope fact — not attributed to the de Boer plot."),
 ("A4", "Precipitation increases as pressure falls and peaks near the bubble point; below Pb some asphaltenes re-dissolve [slide 10 bullet, Q9]",
  "CONFIRMED",
  "“As pressure decreases below the onset precipitation pressure, the amount of asphaltene precipitation increases and reaches its maximum value when pressure approaches the bubble point pressure.” "
  "“As pressure declines below the saturation pressure … already precipitated asphaltenes dissolve back … because of a shift in asphaltene solubility in the denser oil (de Boer et al. 1995).”",
  "Kept; this is the correct home for the AOP/Pb discussion now that A3 is fixed."),
 ("A5", "Onset sits “100–200 bar above the bubble point” [slide 10 script]",
  "REFINED",
  "Published instability windows range from a few hundred bar to several thousand psi depending on the fluid (“some oils have an instability window of several thousand psi”; “region of instability 2,600–3,800 psi”). "
  "A single 100–200 bar figure is not defensible as a general statement.",
  "Script now says the upper onset lies above Pb “by hundreds of bar or more, several thousand psi in severe oils — fluid-specific, and the thesis measures it for the case fluid.”"),
 ("A6", "Asphaltenes are quantified by ASTM D6560 / IP 143 (n-heptane insolubles) [slide 3, Q1, ref [12]]",
  "CONFIRMED",
  "ASTM D6560 / IP 143/21, “Determination of Asphaltenes (Heptane Insolubles) in Crude Petroleum and Petroleum Products”; current edition D6560-22, precision applicable 0.50–30.0 % m/m. "
  "Definition in the standard: “soluble in a specified aromatic solvent but separates upon addition of an excess of a specified paraffinic solvent” — hot toluene and heptane.",
  "Reference line updated to “ASTM D6560-22 / IP 143:21”. Note added: the standard itself records that benzene was dropped for health reasons — an HSE point that supports the ethics line."),
 ("A7", "“ASTM D6703 (standard test for asphaltene precipitation)” — previously listed as unverified [ref [12], source log]",
  "CORRECTED",
  "D6703 is titled “Standard Test Method for Automated Heithaus Titrimetry” (D6703-19). It is an asphaltene/resin stability titration, not a general asphaltene-precipitation test. "
  "The relevant SARA/insolubles standards are D6560 (IP 143), D2007 (clay-gel), D4124 (separation into four fractions) and D3279 (n-heptane insolubles).",
  "Reference line rewritten with correct titles and editions; the open library item from the source log is closed."),
 ("A8", "Asphaltenes exist as nano-aggregates stabilised by resins/aromatics [slide 3, ref [11]]",
  "CONFIRMED",
  "Yen–Mullins (modified Yen) model: most probable molecular weight ≈ 750 g/mol, island architecture dominant; molecules form nanoaggregates (aggregation number < 10, ≈ 6, size ≈ 2 nm); nanoaggregates form clusters (≈ 8, ≈ 5 nm). "
  "Mullins, Energy & Fuels 2010, 24(4), 2179–2207; reviewed in Mullins et al., Energy & Fuels 2012, doi:10.1021/ef300185p.",
  "Refs [11] extended; slide 3 wording aligned (“nano-aggregates ≈ 2 nm”). Debate note added to the risk register (island vs archipelago is still open)."),
 ("A9", "Thermal methods are largely ineffective — asphaltene deposits are amorphous and do not remelt [slide 15, Q3]",
  "CONFIRMED",
  "“Asphaltenes have no melting point and typically cannot be remediated via heating”; of a 497 °C steam treatment: “begs the question of the state of asphaltene as it does not melt but decompose into solid-like material which can pose problems within the formation”; "
  "“among [mechanical, chemical, thermal] removal methods … the thermal methods are less efficient.”",
  "Kept, with a nuance added (heat weakens the resin peptisation that keeps asphaltenes dispersed — so heating can aggravate, not cure). One counter-example paper is anticipated in Q&A (see Q30)."),
 ("A10", "Inhibitors are dispersants — they prevent formation and do not dissolve deposits [slide 16, Q13]",
  "REFINED",
  "Mechanism confirmed: polar head adsorbs on the aggregate, alkyl tail provides steric hindrance (Rogel & León 2001; Fuel 2021 review: dispersants “work by reducing the asphaltene aggregate size thus keeping them in suspension”). "
  "But the literature separates two classes: INHIBITORS delay/shift the flocculation onset, DISPERSANTS shrink aggregates — “inhibitors can also act as dispersants, but in general the inverse is not true” (Kelland 2009).",
  "Slide 16 and Q13 now state the distinction explicitly. Both classes still “prevent, not dissolve” — that part was right."),
 ("A11", "Dosage is field-specific; more is not simply better [slide 16, Q14]",
  "REFINED",
  "Same source family reports the opposite of a linear dose-response: “at high heptane vol%, inhibitor over-treatment (above recommended dose) increases deposition.” Inhibitor self-aggregation reduces adsorption efficiency at excessive concentration.",
  "Added to slide 16 notes and to the risk register (R7). New Q31 covers it."),
 ("A12", "Placement rule: inject deeper than the shallowest predicted onset depth [slide 16, Q15]",
  "CORRECTED",
  "The requirement is to dose the oil BEFORE it enters the unstable window, i.e. the injection point must sit BELOW the depth at which the flowing path enters the envelope. "
  "In the documented case the mandrel (4,450 m) sat ABOVE the predicted onset (4,752 m), leaving ~300 m of unstable, unprotected tubing — the exact failure the rule exists to prevent. "
  "Under uncertainty the conservative design uses the DEEPEST credible onset depth, not the shallowest.",
  "Slide 16, the script, Q15 and backup slide 24 rewritten as: “below the deepest credible onset depth — onset + a sensitivity-sized margin.”"),
 ("A13", "Aromatic solvents (toluene/xylene) carry real HSE exposure — flagged as a workplace-safety topic [slide 15, Q12]",
  "CONFIRMED",
  "“BTX chemicals (benzene, toluene, and xylene) … are acutely toxic and harmful to the environment”; “BTX solvents have low flash point, high acute toxicity and low biodegradability”; "
  "a field review of xylene cleanouts concludes the solvent “has limited effectiveness in addition to undesirable HSE effects.”",
  "Flag retained on slide 15; the limitations box on slide 21 is now titled “LIMITATIONS & ETHICS” and carries the no-uncontrolled-handling statement."),
 ("A14", "Sensitivity levers: ±50–100 bar on AOP, ±10 % on gradient, GOR scenarios [backup slide 24]",
  "CONFIRMED & QUANTIFIED",
  "The three levers are the correct ones (AOP, pressure gradient, composition/GOR). They were, however, only described — not quantified. See Part D: the arithmetic is now on the slide.",
  "Backup slide 24 now carries the numbers: ±255 m (±50 bar), ±511 m (±100 bar), −404 / +487 m (±10 % gradient), +768 m in the gas-lift scenario."),
]
rows += [list(r) for r in data]
tbl(rows, [1.0, 4.6, 1.9, 7.6, 4.3], fs=7.4, highlight_col=2)

# ================= PART B =================
doc.add_heading("Part B — What changed in the pack (before → after)", level=1)
rows = [["Where", "Before this pass", "After this pass"]]
data = [
 ("Slide 10 (bullets + script)", "“De Boer: if AOP > Pb → asphaltene-prone.”", "Two separate facts, correctly attributed: (i) the de Boer PLOT screens on in-situ density vs gas under-saturation; (ii) the upper onset lies above Pb, with precipitation peaking near Pb."),
 ("Slide 10 (script)", "“Onset can sit 100–200 bar above the bubble point.”", "“Above Pb by hundreds of bar or more — several thousand psi in severe oils; fluid-specific.”"),
 ("Slide 16 (bullets + Q13)", "“Chemical inhibitors (dispersants) … they do not dissolve deposits.”", "Kept, plus the correct distinction: inhibitors shift the onset; dispersants shrink aggregates; neither dissolves an existing deposit."),
 ("Slide 16 (bullet + Q15) and backup 24", "“Injection point must sit DEEPER than the shallowest predicted onset depth.”", "“Injection point must sit BELOW the deepest credible onset depth — onset plus a margin sized by the sensitivity study.”"),
 ("Slide 15 (caption)", "“Thermal methods fail here by design.”", "“Thermal methods are not a melting story — asphaltene deposits have no melting point; heat can even weaken the resins that keep them dispersed.”"),
 ("Slide 21 (box title + bullets)", "“LIMITATIONS (stated up-front)” — five technical limits.", "Retitled “LIMITATIONS & ETHICS”, with a sixth line: no uncontrolled solvent handling; chemistries compared on paper, no vendor recommended."),
 ("Backup slide 24", "“Sensitivity plan: ±50–100 bar … the onset depth moves by hundreds of metres.”", "Quantified table (Part D) + the design conclusion: the gas-lift scenario pushes the onset to the perforations, where placement alone cannot protect."),
 ("Refs [11], [12] + new [15]", "Mullins 2010 alone; “ASTM D6703 (asphaltene precipitation)”; no thermal reference.", "Mullins 2010 + Mullins et al. 2012 review; ASTM line corrected (D6560-22/IP 143:21, D2007, D4124, D3279, D6703 = Heithaus titrimetry); [15] added for the thermal/steam finding."),
 ("Q&A pack", "Q6, Q7, Q13, Q15 as above.", "Corrected in place; Q29–Q31 added (de Boer definition, the steam counter-example, over-dosing)."),
]
rows += [list(r) for r in data]
tbl(rows, [3.4, 6.2, 8.8], fs=7.8)

# ================= PART C =================
doc.add_heading("Part C — Risk register (project + presentation)", level=1)
para("L = likelihood (1 low → 3 high), I = impact (1 low → 3 high). Exposure = L × I. Reviewed weekly; any trigger firing moves the mitigation from “planned” to “active”.",
     size=8, italic=True, color=GRAY, space_after=4)
rows = [["ID", "Risk", "Area", "L", "I", "Exposure", "Mitigation / control in place", "Trigger to watch"]]
data = [
 ("R1", "AOP input uncertainty moves the onset depth by hundreds of metres, so the injection margin is mis-sized.", "Technical — model", "3", "3", "9",
  "Sensitivity study quantified (Part D) and shown on backup 24; margin sized from the worst credible case, not from a rule of thumb; tolerance stated with every prediction.", "Any case fluid without a measured AOP curve."),
 ("R2", "CII gives a misleading screening verdict (thresholds laboratory-specific; poor performance reported for some crudes).", "Technical — screening", "2", "2", "4",
  "CII is never used alone: always paired with the de Boer density/under-saturation screen and a measured onset curve; the caveat is stated on the slide.", "CII lands in the 0.7–0.9 band."),
 ("R3", "A classic misconception trap in Q&A (de Boer plot definition; “deposits form where pressure drops fastest”).", "Presentation — credibility", "2", "3", "6",
  "Both corrected in the deck, the script and Q6/Q29; the misconception is now turned into a deliberate teaching point on slide 12.", "An examiner opens with “isn’t the de Boer plot…”.	"),
 ("R4", "Over-claiming: presenting a thermodynamic onset model as a deposition predictor.", "Technical — scope", "2", "3", "6",
  "Precipitation ≠ deposition is the red box of slide 9; limitations on slide 21; deposition-transport modelling declared first-order.", "Any question about deposit thickness or growth rate."),
 ("R5", "Licensed PVT / flow-assurance software unavailable.", "Project — tools", "2", "2", "4",
  "Tool hierarchy declared on slide 20 and in Q20: licensed → NeqSim (CPA-EoS, verified available, Apache-2.0) → correlation screening as the floor.", "Week-1 licence enquiry returns negative."),
 ("R6", "The chosen case well lacks the data needed to tune the model (no measured AOP, no reliable production history).", "Project — data", "2", "3", "6",
  "Week 1: case selection against a data-availability checklist; fallback is a published AOP for a comparable fluid, declared explicitly.", "Case data sheet incomplete at Week 2 review."),
 ("R7", "Over-dosing: above the recommended dose at high destabilisation, inhibitor can INCREASE deposition.", "Technical — chemistry", "2", "2", "4",
  "Dosage presented as a window, never “more is better”; thesis models the optimum; new Q31 arms the answer.", "Any recommendation phrased as a single number."),
 ("R8", "HSE exposure if lab or field work with aromatic solvents is added later.", "HSE / ethics", "1", "3", "3",
  "No uncontrolled handling is recommended anywhere in the pack; any future lab work requires COSHH/MSDS, supervisor sign-off and a waste route before the first sample is opened.", "A supervisor offers bench work with BTEX solvents."),
 ("R9", "Scope creep — the project grows into deposition-transport physics and misses the 16-week plan.", "Project — schedule", "2", "3", "6",
  "Four work packages each with a named output; scope statement frozen from Week 1; the economics package is the declared cut line.", "Any week without a deliverable."),
 ("R10", "Seminar overruns 15 minutes.", "Presentation — delivery", "2", "2", "4",
  "Checkpoints at slides 12 / 13 / 17 / 22 with a rehearsed cut list; script word-budgeted to 14:10.", "Checkpoint missed by > 20 s in rehearsal."),
 ("R11", "Cost-attribution conflict challenged (Middle East vs GoM inhibitor ranges).", "Presentation — sources", "2", "2", "4",
  "Caveat printed on slide 17; full explanation in Q26 and 04_Source_Log.docx; the majority attribution is quoted and the discrepancy flagged.", "Examiner cites the Energy & Fuels (2021) attribution."),
 ("R12", "Thermal-method claim challenged with a counter-example (steam treatments reporting high permeability recovery).", "Technical — technical challenge", "2", "2", "4",
  "Slide 15 caption refined; Q30 gives the honest answer: thermal is not a melting/dissolution mechanism, reported steam successes work by other means, and 497 °C steam questions the deposit’s state rather than restoring solubility.", "A field case with steam success is quoted."),
 ("R13", "Island-vs-archipelago debate over asphaltene molecular structure put to you as if settled.", "Technical — science", "1", "2", "2",
  "Slide 3 speaks of nano-aggregates and the Yen–Mullins model without claiming molecular architecture is closed; the open debate is recorded in Part A (A8).", "“Archipelago model disproves Mullins” challenge."),
 ("R14", "Perceived endorsement of a specific chemical vendor or product class.", "Ethics", "1", "2", "2",
  "No product, brand or vendor is named anywhere; the pack compares chemistry CLASSES and cites published ranges only.", "Any question asking “which product should we buy?”."),
]
rows += [list(r) for r in data]
tbl(rows, [0.9, 5.0, 1.8, 0.7, 0.7, 1.2, 6.2, 3.2], fs=7.2)

# ================= PART D =================
doc.add_heading("Part D — Sensitivity study of the onset depth (the numbers behind backup slide 24)", level=1)
para("Method. The illustrative model used throughout the deck: P(d) = 150 + 1000·(d/5500) bar; T(d) = 60 + 85·(d/5500) °C; "
     "AOP(T) = 1080 − 0.9·(T − 60) bar. The onset depth is where P(d) = AOP(T(d)) — the base case gives 4,752 m, consistent with the deck and the figures.",
     size=8.5)
rows = [["Scenario", "Perturbation applied", "Onset depth", "Shift vs base", "What it means for the injection point"]]
data = [
 ("Base case", "as modelled", "4,752 m", "—", "Target injection ≈ 5,000 m = onset + 250 m margin."),
 ("AOP ±50 bar", "fluid less / more stable than modelled", "5,007 / 4,496 m", "± 255 m", "A 250 m margin covers realistic AOP scatter — but only just."),
 ("AOP ±100 bar", "worst credible characterisation error", "5,262 / 4,241 m", "± 511 m", "At the upper bound a 250 m margin is insufficient: the injector must move deeper (≈ 5,300 m)."),
 ("Gradient ±10 %", "lighter / heavier flowing column", "5,238 / 4,348 m", "−404 / +487 m", "A lighter column (gas lift) pushes the onset DEEPER — the error direction is the dangerous one."),
 ("Gas-lift scenario", "gradient −10 % AND AOP +50 bar (gas strips resins)", "5,520 m", "+ 768 m", "The onset falls below the perforations (5,395–5,486 m): the ENTIRE tubing is unstable. Placement alone cannot protect the well — dosage effectiveness does the work, with the injector set as deep as practical."),
 ("Better-solvent case", "gradient +10 % AND AOP −50 bar", "4,114 m", "− 638 m", "Favourable case: a shallower injector would suffice — which is exactly why a conservative design must not assume it."),
]
rows += [list(r) for r in data]
tbl(rows, [2.4, 4.4, 2.3, 2.0, 7.3], fs=7.6)
rich([("Design conclusion.  ", True, RED),
      ("The injection point is placed below the DEEPEST credible onset depth from this table — not below the base case. "
       "In the gas-lift scenario that depth approaches the perforations, so the deliverable is a placement PLUS a dosage-window recommendation, "
       "and the uncertainty margin is quoted with the number. This is the sensitivity analysis promised as WP3 output — backup slide 24 now carries it as arithmetic, not as a plan.", False, None)],
     size=9, space_after=6)

# ================= PART E =================
doc.add_heading("Part E — HSE and ethics position (stated, not implied)", level=1)
for t in [
 "Aromatic solvents (toluene, xylene — BTEX class) are correctly presented as a workplace-safety topic, not as a casual fix: acute toxicity, low flash point, low biodegradability, and limited long-term effectiveness.",
 "The seminar compares CLASSES of mitigation on published evidence. It never recommends uncontrolled handling of any solvent, and names no vendor or product.",
 "The D6560 standard itself records that benzene was dropped as the aromatic solvent for health reasons — a useful, citable illustration that the industry has already had to make this trade-off.",
 "If laboratory work with these solvents is added to the thesis, COSHH/MSDS assessment, supervisor sign-off and a waste-route plan are prerequisites — stated here so the boundary is on the record before any work starts.",
 "Data and attribution ethics: where published cost or performance figures conflict (R11), the pack quotes the majority attribution, flags the discrepancy and cites both — rather than quietly selecting the more impressive number.",
]:
    p = doc.add_paragraph(t, style="List Bullet"); p.paragraph_format.space_after = Pt(2)
    for r in p.runs: r.font.size = Pt(8.5)

para("Prepared as part of the seminar package. The verification in Part A was performed against the sources recorded in 04_Source_Log.docx; "
     "the sensitivity arithmetic in Part D is reproducible from _build/make_accuracy.py.",
     size=8, italic=True, color=GRAY, space_after=0)

doc.save("/home/user/Seminar_Asphaltenes/06_Technical_Accuracy_and_Risk_Register.docx")
print("saved 06_Technical_Accuracy_and_Risk_Register.docx")
