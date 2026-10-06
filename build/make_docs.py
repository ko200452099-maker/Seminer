#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Builds the three companion documents for the asphaltene seminar package."""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

NAVY = RGBColor(0x10, 0x2A, 0x43); TEAL = RGBColor(0x1F, 0x7A, 0x84)
GRAY = RGBColor(0x5A, 0x66, 0x72); RED = RGBColor(0xB4, 0x3B, 0x2E)
AMBER = RGBColor(0xA8, 0x6B, 0x00)

def new_doc(footer_text):
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = "Calibri"; st.font.size = Pt(10.5)
    st.paragraph_format.space_after = Pt(6); st.paragraph_format.line_spacing = 1.08
    for name, size, color in (("Heading 1", 16, NAVY), ("Heading 2", 13, TEAL), ("Heading 3", 11.5, NAVY)):
        s = doc.styles[name]; s.font.name = "Calibri"; s.font.size = Pt(size)
        s.font.bold = True; s.font.color.rgb = color
        s.paragraph_format.space_before = Pt(12); s.paragraph_format.space_after = Pt(4)
    sec = doc.sections[0]
    sec.top_margin = Cm(1.9); sec.bottom_margin = Cm(1.9); sec.left_margin = Cm(2.0); sec.right_margin = Cm(2.0)
    p = sec.footer.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(footer_text + "  —  page "); r.font.size = Pt(8); r.font.color.rgb = GRAY
    r2 = p.add_run()
    f1 = OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = "PAGE"
    f2 = OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"), "end")
    r2._r.append(f1); r2._r.append(it); r2._r.append(f2); r2.font.size = Pt(8); r2.font.color.rgb = GRAY
    return doc

def para(doc, text, size=10.5, bold=False, italic=False, color=None, align=None, space_after=6, indent=None):
    p = doc.add_paragraph()
    if align: p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    if indent is not None: p.paragraph_format.left_indent = Cm(indent)
    r = p.add_run(text); r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic
    if color: r.font.color.rgb = color
    return p

def rich(doc, parts, size=10.5, space_after=5):
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(space_after)
    for t, b, i, c in parts:
        r = p.add_run(t); r.font.size = Pt(size); r.font.bold = b; r.font.italic = i
        if c: r.font.color.rgb = c
    return p

def bullet(doc, text, size=10.5):
    p = doc.add_paragraph(style="List Bullet"); p.paragraph_format.space_after = Pt(3)
    r = p.add_run(text); r.font.size = Pt(size)
    return p

def shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hexcolor)
    tcPr.append(shd)

def tbl(doc, rows, widths=None, fs=9.5):
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = t.cell(ri, ci); cell.text = ""
            p = cell.paragraphs[0]; p.paragraph_format.space_after = Pt(2)
            r = p.add_run(str(val)); r.font.size = Pt(fs); r.font.name = "Calibri"
            if ri == 0:
                r.font.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); shade(cell, "102A43")
            elif ri % 2 == 0:
                shade(cell, "F1F5F8")
    if widths:
        total = sum(widths)
        for ci, w in enumerate(widths):
            for ri in range(len(rows)):
                t.cell(ri, ci).width = Cm(17.0 * w / total)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t

# =====================================================================
# DOC 1 — MASTER PLAN
# =====================================================================
plan = new_doc("Seminar Master Plan — Asphaltenes: Predicting and Preventing Wellbore Plugging")
p = plan.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("SEMINAR — MASTER PLAN AND DELIVERY PACKAGE"); r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = TEAL
p = plan.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Asphaltenes: Predicting and Preventing Wellbore Plugging"); r.font.size = Pt(23); r.font.bold = True; r.font.color.rgb = NAVY
p = plan.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Flow assurance in production wellbores — and the seed of the graduation project"); r.font.size = Pt(13); r.font.color.rgb = GRAY
plan.add_paragraph()
tbl(plan, [
    ["Parameter", "Value"],
    ["Course", "Seminar — BSc Petroleum and Gas Engineering, [University]"],
    ["Student", "[Your full name] — ID [xxxxxx]"],
    ["Supervisor", "[Prof. name]"],
    ["Format", "15-minute oral presentation · 23 core slides + 2 backup slides"],
    ["Topic character", "Flow-assurance review of a studied subject + graduation-project preview"],
    ["Companion files", "Asphaltene_Slides_23plus2.pptx (+300-dpi PDF proof) · 01_Asphaltene_Speaking_Script.docx · 02_Asphaltene_QandA_Pack.docx · 03_Animation_and_Design_Guide.docx · 04_Source_Log.docx · 05_Projector_Preflight.docx · 06_Technical_Accuracy_and_Risk_Register.docx"],
    ["Delivery month", "[Month] 2026"],
], widths=[3.2, 13.8], fs=9.5)
para(plan, "Note: this package replaces the earlier casing-design package at your request. The casing package remains in the workspace as an alternative.", size=10, italic=True, color=GRAY)

plan.add_heading("0.  The arithmetic that drives everything", level=1)
plan.add_heading("0.1  What changed from the original blueprint (and why)", level=2)
tbl(plan, [
    ["Blueprint element", "This package"],
    ["Slide-by-slide list only", "Full timing architecture, checkpoints, cut list, and a built deck with speaker notes"],
    ["No buffer (parts summed to exactly 15:00)", "Speaking ends at ≈14:15–14:25 with a deliberate buffer; references slide is skippable"],
    ["Slides 22–23 mixed thanks with Q&A", "Conclusions and thanks are separate (slide 22), references (23), backup (24–25)"],
    ["'Asphaltenes keep the fluid stable'", "Corrected: asphaltenes are KEPT DISPERSED by resins and aromatics — committees test this"],
    ["'Deposits form near the top where P drops fastest'", "Corrected: deposits form where the flowing path crosses the envelope — often deep in the tubing (this became the case study)"],
    ["Generic graduation options", "One committed route (modelling) with a declared tool hierarchy so no licence problem can stall the project"],
    ["Two Q&A questions", "31 questions with model answers, including refined versions of the original two"],
], widths=[6.0, 11.0], fs=9)
plan.add_heading("0.2  The two hard numbers", level=2)
bullet(plan, "23 core slides + 2 labelled backup = 25 total: compliant whether your professor counts presented or total slides.", 10)
bullet(plan, "Speaking budget ≈ 14:25 with ≈ 1,830 measured words at a rehearsed 128–132 wpm — the speaking script is written to that budget and word-counted.", 10)
rich(plan, [("The four time checkpoints:  ", True, False, NAVY),
            ("slide 12 at 7:05  ·  slide 13 at 8:00  ·  slide 17 at 10:40  ·  conclusions (slide 22) start by 14:15.", False, False, None)])
para(plan, "Customisation rule: replace every [bracketed] placeholder and resolve every [VERIFY] flag BEFORE the first rehearsal.", size=10, italic=True, color=AMBER)

plan.add_heading("1.  The topic — and its two anchors", level=1)
para(plan, "The deck is built on two verified, citable anchors. Learn them cold; they convert a 'review' into a research narrative:", size=10)
tbl(plan, [
    ["Anchor", "What it gives the seminar"],
    ["The documented deepwater Gulf of Mexico well (JPT/SPE, 2021)", "A real case where the chemical injection mandrel sat 3,400 ft ABOVE the point where the flowing pressure crossed the onset — tubing ID choked from 2.3 to 1.35 in. This is the proof that injection-point placement is a design decision, and it is the case your thesis will reproduce."],
    ["Hassi Messaoud (Algeria, SPE-994-PA and successors)", "The classic field: deposits 83.4 % asphaltenes; wells losing 20–25 % wellhead pressure in 15–20 days; gauge-ring profiling practice; and the counterintuitive discovery that producing at LOW wellhead pressure drastically reduced cleanouts."],
], widths=[4.6, 12.4], fs=9)

plan.add_heading("1.3  How the seminar becomes the thesis (chapter mapping)", level=2)
para(plan, "The project-linkage guarantee: every slide already has a home in the thesis. Nothing presented is wasted work.", size=10, italic=True, color=NAVY)
tbl(plan, [
    ["Seminar element", "Thesis destination"],
    ["Slides 2–4 — problem, definition, impact, cross-section", "Chapter 1 — Introduction, problem statement, objectives"],
    ["Slides 7–12 — stability, wax vs asphaltene, mechanism, envelope, screening, triggers", "Chapter 2 — Literature review and fluid-characterisation methods"],
    ["Slides 13–14 — documented GoM case + detection practice", "Chapter 3 — Case-study definition, data set, validation targets"],
    ["Slides 15–17 — reactive limits, proactive inhibition, economics", "Chapter 4 — Mitigation review and the breakeven model"],
    ["Slides 18–21 — gap, aim, RQs, methodology, plan, limitations", "Chapter 1.4 + Chapter 5 — Methodology and results"],
    ["Slide 22 + Q&A pack + 04_Source_Log.docx + 06_Technical_Accuracy_...docx", "Chapter 6 — Conclusions; defence preparation"],
], widths=[7.5, 9.5], fs=9)

plan.add_heading("2.  Timing architecture", level=1)
tbl(plan, [
    ["Part", "Slides", "Time slot", "Job of this part"],
    ["1 · The problem", "1–6", "0:00 – 3:05", "Hook with four numbers; define asphaltenes correctly; scope; roadmap"],
    ["2 · The science", "7–12", "3:05 – 7:05", "Stability balance, wax-vs-asphaltene, mechanism chain, envelope, screening, triggers"],
    ["3 · Field evidence", "13", "7:05 – 8:00", "The documented case: onset depth vs injection depth"],
    ["4 · Defenses", "14–17", "8:00 – 10:40", "Detection, reactive limits, proactive inhibition, economics"],
    ["5 · The project", "18–21", "10:40 – 13:35", "Gap, aim, RQs, methodology, deliverables, plan"],
    ["Close", "22–23", "14:15 – 14:25", "Three messages; references; questions"],
], widths=[2.6, 1.8, 2.6, 10.0], fs=9.5)
para(plan, "Slide map with per-slide targets and messages: identical structure to the deck footer (target time + cumulative clock on every slide). The cue-card version is Appendix B.", size=10, italic=True)

doc_rows = [["#", "Slide", "Target", "Clock", "The one message"]]
data = [
    (1,"Title",15,"0:15","Who, what, and the thesis promise"),
    (2,"A billion-dollar problem",45,"1:00","Four numbers that establish the stakes"),
    (3,"What are asphaltenes",35,"1:35","Solubility class; kept dispersed by resins — not 'stabilising' the oil"),
    (4,"Inside the tubing (image)",35,"2:10","ID 2.3→1.35 in; WHP losses; invisible until expensive"),
    (5,"Objective & scope",35,"2:45","Explain / review / propose; not a field intervention study"),
    (6,"Roadmap",20,"3:05","You own the clock"),
    (7,"Stability is a balance",35,"3:40","Three levers; CII as screening"),
    (8,"Asphaltenes vs wax",40,"4:20","Not a cooling story; amorphous, pressure/composition-driven"),
    (9,"Mechanism chain",35,"4:55","Precipitation ≠ deposition — the critical distinction"),
    (10,"The envelope (chart)",45,"5:40","AOP, de Boer screening, path crossing = onset"),
    (11,"Screening & measurement",40,"6:20","SARA → CII → de Boer → measured AOP curve"),
    (12,"Triggers + misconception",45,"7:05","CHECKPOINT 1 — five triggers; deposits form where the path crosses, often deep"),
    (13,"Field case (chart)",55,"8:00","CHECKPOINT 2 — mandrel above onset: 3,400 ft unprotected"),
    (14,"Detection",35,"8:35","ΔP/Q fingerprint; gauge rings; sampling"),
    (15,"Reactive toolkit fails",45,"9:20","Every cleanout is NPT and a repeat visit"),
    (16,"Proactive inhibition",40,"10:00","Dispersants keep particles flowing; placement rule"),
    (17,"Economics",40,"10:40","CHECKPOINT 3 — breakeven model is the missing piece"),
    (18,"Gap analysis",40,"11:20","Prediction is the weak link; documented rigour is the contribution"),
    (19,"Proposed project",45,"12:05","Aim, RQs, work packages — one case well, end to end"),
    (20,"Methodology",50,"12:55","Screening → modelling → validation → economics; tool hierarchy declared"),
    (21,"Deliverables & limits",40,"13:35","Bounded scope; honest limitations"),
    (22,"Three messages + thanks",40,"14:15","CONCLUSIONS START HERE — memorised verbatim"),
    (23,"References",10,"14:25","One sentence; skippable"),
    (24,"BACKUP — derivations",0,"—","CII, de Boer, onset model, sensitivity, exclusions"),
    (25,"BACKUP — data & tools",0,"—","Sources, NeqSim fallback, validation targets, risk register"),
]
for n, t, tg, cl, msg in data:
    doc_rows.append([n, t, f"{tg} s" if tg else "backup", cl, msg])
tbl(plan, doc_rows, widths=[0.7, 4.8, 1.2, 1.2, 10.1], fs=8.5)

plan.add_heading("2.3  Cut list (decide before the presentation, not during it)", level=2)
tbl(plan, [
    ["Situation", "Action"],
    ["On slides 12 / 13 / 17 at 7:05 / 8:00 / 10:40", "You are on time — continue as rehearsed"],
    ["Up to 20 s behind", "Compress slide 11 (three methods only) and slide 15 (three rows only)"],
    ["30–60 s behind", "Speak slide 23 in one sentence or skip; compress slides 14 and 16"],
    ["More than 60 s behind", "Drop slide 23 entirely; give Part 2 as headlines. NEVER cut slide 22"],
    ["Ahead of schedule", "Expand slide 13 (case study detail) — do not invent new claims"],
], widths=[6.5, 10.5], fs=9.5)

plan.add_heading("3.  Work plan (six weeks + presentation week)", level=1)
tbl(plan, [
    ["Week", "Focus", "Deliverable by end of week"],
    ["1", "Lock the topic with your supervisor; confirm tool licences (PVTsim/Multiflash/OLGA vs NeqSim); resolve [VERIFY] list; start bibliography", "Approved title in writing + data/tool decision logged"],
    ["2", "Write the story: 25 headlines and key messages; build the onset-depth calculation sheet with the case geometry", "Storyboard + working spreadsheet"],
    ["3", "Build slides 1–13; adapt speaker notes; write half the script", "Slides 1–13 + script Part 1–3"],
    ["4", "Build slides 14–25; finish script; first read-aloud with stopwatch", "Complete deck + complete script"],
    ["5", "Rehearsals 2–4; fix timings; prepare Q&A pack; memorise every field number", "Frozen deck; corrected script; Q&A pack"],
    ["6", "Dress rehearsal with listeners + question drill; polish; export PDF; submit", "Submitted deck + handout"],
    ["Deadline week", "Light review; 30 minutes early; test projector", "15 calm, controlled minutes"],
], widths=[1.6, 8.4, 7.0], fs=9)
bullet(plan, "Content freezes after Week 5. No new ideas in presentation week.", 10)
bullet(plan, "Rule: any number on a slide must trace to a source on slide 23 — this topic's numbers span fields and years, and the 17's caveat line ('never as a single industry cost') is deliberate.", 10)

plan.add_heading("4.  Sources and tools", level=1)
tbl(plan, [
    ["Category", "Sources", "Used on"],
    ["Reviews (costs & chemistry)", "ACS Energy & Fuels (2021); MDPI Processes (2025); Khaleel et al., Fluid Phase Equilibria (2020)", "Slides 2, 15–18, 23"],
    ["Field cases", "JPT/SPE (2021) deepwater GoM well; Hassi Messaoud SPE-994-PA and successors; west Kuwait Marrat monitoring study", "Slides 4, 13–16, 23"],
    ["Screening & standards", "de Boer et al. (1995); CII screening literature; ASTM D6560/IP 143; ASTM D6703; Mullins — Modified Yen Model", "Slides 7, 8, 10, 11, 24"],
    ["Tools", "Licensed: PVTsim / Multiflash / WinProp-class PVT; OLGA (asphaltene module) or PIPESIM. Open-source: NeqSim (CPA-EoS asphaltene screening). Python for the workflow.", "Slides 20, 25"],
], widths=[3.4, 10.0, 3.6], fs=9)

plan.add_heading("5.  Rehearsal protocol (same five-rehearsal ladder)", level=1)
tbl(plan, [
    ["#", "Format", "Pass criterion"],
    ["1", "Script read aloud, seated, no slides", "≤ 2 stumbles; no rewriting during the read"],
    ["2", "Timed, standing, cue card only", "Within ±30 s of 14:25"],
    ["3", "With slides, clicker, full timing", "Every checkpoint within ±15 s"],
    ["4", "Record on your phone and watch it", "Fillers halved; pace 128–132 wpm"],
    ["5", "Dress: 2–3 listeners + question drill", "14:25 ± 15 s and survive 5 questions"],
], widths=[0.8, 6.2, 10.0], fs=9)
para(plan, "Memorise verbatim: the opening two sentences, the slide-12 misconception line, the slide-13 punchline ('the well had inhibition — but 3,400 ft of tubing sat below it'), and the three closing messages.", size=10, italic=True, color=NAVY)

plan.add_heading("6.  Q&A overview", level=1)
para(plan, "The full pack is in 02_Asphaltene_QandA_Pack.docx (31 questions). The six most probable, with soundbites:", size=10)
tbl(plan, [
    ["Question", "Two-sentence answer"],
    ["Why do asphaltenes precipitate when we gas-lift?", "The injected light gas lowers the liquid's solvency and can strip the stabilising resin/aromatic fraction, so the colloidal dispersion collapses. It is a composition change, not a temperature one."],
    ["Why not simply choke the well to keep pressure high?", "Choking keeps the flowing pressure above the AOP but sacrifices rate and revenue — and Hassi Messaoud showed the opposite lever works too: producing at LOW wellhead pressure reduced cleanouts. The right answer is the P–T path relative to the envelope, which my project quantifies per well."],
    ["Are precipitated asphaltenes automatically a deposit?", "No — that is the central distinction. Deposition requires adhesion and is balanced by shear removal; many wells produce precipitated asphaltenes with no deposit. Transport modelling is the least mature part of the science."],
    ["Given CO₂ injection and sequestration growth, does this matter more?", "Yes — CO₂ is a documented destabiliser of asphaltenic crudes, so the screening and prediction workflow applies directly to CO₂ EOR and storage projects."],
    ["Isn't the deposit caused by cooling?" , "Cooling is the primary trigger for wax, not asphaltenes. Temperature modifies the asphaltene envelope, but pressure and composition dominate — and amorphous deposits do not remelt."],
    ["What data will you use?", "Published fluid and field data: the documented GoM case, Hassi Messaoud data, and PVT/SARA reports from the literature, with NeqSim as the open-source modelling fallback if licensed tools are unavailable."],
], widths=[5.6, 11.4], fs=9)

plan.add_heading("7.  Risk register", level=1)
tbl(plan, [
    ["Risk", "L", "I", "Mitigation"],
    ["No live crude sample available", "H", "M", "Use published fluid data and declare it; the method is the contribution, not the sample"],
    ["Licensed software unavailable", "M", "H", "Declared hierarchy: licensed → NeqSim (open source) → correlations. Project designed to survive any outcome"],
    ["Case-data conflicts across sources", "M", "M", "Document variance; choose conservative values; state the range on the slide"],
    ["Running over time", "M", "H", "Checkpoints + cut list; script word-budgeted to 14:25"],
    ["Committee pushes into deposition-transport modelling", "H", "M", "Declared first-order; offered as thesis extension, with sensitivity range"],
    ["Topic redirected by supervisor", "M", "H", "The casing/geothermal package in the workspace remains a ready alternative"],
], widths=[4.8, 0.9, 0.9, 10.4], fs=9)

plan.add_heading("8.  Pre-submission checklist", level=1)
for t in [
    "All bracketed placeholders replaced.",
    "Every [VERIFY] resolved: exact citations for CII, ASTM standards, AOP screening, cost figures.",
    "Every slide number traceable to slide 23; cost figures carry their field/year context.",
    "Slide count compliant (23 core / 25 with backup); backup labelled.",
    "Deck exported to PDF; script + Q&A printed; USB + cloud copies prepared.",
    "Checkpoints re-measured after final edits (±15 s).",
    "The five memorised lines rehearsed to perfection.",
    "Q&A pack skimmed the evening before and the morning of.",
]:
    bullet(plan, "[  ]   " + t, 10)

plan.add_page_break()
plan.add_heading("Appendix A — Numbers and formulas cheat sheet (keep at the lectern)", level=1)
tbl(plan, [
    ["Quantity", "Relation / value", "Notes to say aloud"],
    ["CII (Colloidal Instability Index)", "CII = (Saturates + Asphaltenes) / (Resins + Aromatics)", "Boundary commonly ≈ 0.9; thresholds are lab-specific"],
    ["De Boer screening", "AOP vs bubble point vs reservoir pressure", "If AOP > Pb → expect asphaltene problems"],
    ["AOP position for unstable oils", "typically 100–200 bar above the bubble point", "Cite the screening literature you use"],
    ["Onset depth", "depth where flowing P(d) falls below AOP(T(d))", "The number that sets the injection point"],
    ["Asphaltene definition", "insoluble in n-heptane; soluble in toluene (SARA)", "ASTM D6560 / IP 143 [verify edition]"],
    ["Case well numbers (GoM, JPT/SPE 2021)", "perfs 17,700–18,000 ft · mandrel 14,600 ft · 3,400 ft unprotected · ID 2.3→1.35 in · 750→500 ppm · +16 % / 4 months", "Memorise all six"],
    ["Hassi Messaoud numbers", "deposit 83.4 % asphaltenes · 20–25 % WHP loss in 15–20 days · gauge-ring profiling", "The classic field; cite SPE-994-PA"],
    ["Cost anchors", "industry-wide billions/yr · ≈ $70 M single-well GoM shut-in · up to $1.2 M/day deferred · inhibitor $31–46 k (ME) vs $330–390 k (GoM) per well/yr", "State field and year with each figure"],
    ["Tools", "PVTsim / Multiflash / WinProp · OLGA / PIPESIM · NeqSim (open source, CPA-EoS)", "The hierarchy is a project feature, not a backup plan"],
], widths=[4.4, 7.2, 5.4], fs=9)

plan.add_heading("Appendix B — Cue card (25 slides, folded, at the lectern)", level=1)
tbl(plan, doc_rows, widths=[0.7, 5.4, 1.3, 1.3, 9.3], fs=8)
plan.save("/home/user/Seminar_Asphaltenes/00_Asphaltene_Master_Plan.docx")
print("saved plan")

# =====================================================================
# DOC 2 — SPEAKING SCRIPT
# =====================================================================
script = [
 (1, "Title", 15, "0:15",
  "Good morning. My name is [—]. In fifteen minutes I will show you how the heaviest fraction of crude oil can choke a well to death — and how we can predict it before it happens. The title: asphaltenes — predicting and preventing wellbore plugging. Field engineers call them the cholesterol of crude oil. Let me show you why.",
  "Introduce yourself; state the title; promise the thesis link.", None),
 (2, "A billion-dollar problem", 45, "1:00",
  "First, the stakes. Four numbers. Asphaltene deposition costs the industry billions of dollars every year. In a documented deepwater Gulf of Mexico well, a single shut-in and cleanup cost about seventy million dollars. Deferred production during a plugging event reaches one point two million per day. And in severe fields, the deposit has been measured at up to two thirds of the tubing radius. This is not exotic — it is managed somewhere every day.",
  "Land each figure with a pause. Point at each card.", None),
 (3, "What are asphaltenes", 35, "1:35",
  "So what are they? Asphaltenes are a solubility class — defined operationally: insoluble in n-heptane, soluble in toluene, the heaviest and most polar fraction of the crude. They exist in the oil as nano-aggregates, kept dispersed by a solvating shell of resins and aromatics. That shell is the stability — not the asphaltene itself. Disturb the shell, and the aggregates flocculate and come out of solution.",
  "Correct the common inversion: resins stabilise the dispersion.", None),
 (4, "Inside the tubing", 35, "2:10",
  "Here is what that looks like. In the documented Gulf of Mexico case, the effective tubing diameter fell from two point three inches to one point three five — the pipe lost roughly two thirds of its flow area. In severe fields, wells lost twenty to twenty-five percent of wellhead pressure within fifteen to twenty days. First the rate declines, then the well stops flowing altogether — and until the pressure signature is obvious, the deposit is invisible.",
  "Point into the image: 'this pipe still has a hole in it — for now.'",
  "If late: keep the ID number and the 20–25 % figure only."),
 (5, "Objective and scope", 35, "2:45",
  "So, three verbs for this seminar. Explain what destabilises asphaltenes. Review today's defenses — and why they are mostly reactive. And propose the project: predicting onset depth and the injection point for a real well. Let me be explicit: this is a flow-assurance review and a feasibility study — not a completed field intervention study. I claim the roadmap, not the destination.",
  "The scope sentence is your shield. Say it slowly.", None),
 (6, "Roadmap", 20, "3:05",
  "The structure: five parts. The problem — to minute three. The science of stability — to minute seven. The field evidence — the case study that carries the argument. The defenses, then and now — to about minute eleven. And the graduation project. Then three messages, and your questions. One reminder in advance, at the bottom: this is a flow-assurance review plus a thesis feasibility study — not a completed field intervention.",
  "Fast and confident; do not linger. Gesture once at the scope line, then move on.", None),
 (7, "Stability is a balance", 35, "3:40",
  "Why does stable oil become unstable? Two engineering views — colloidal and solubility — and one working rule: asphaltenes are most soluble in aromatic, dense fluids. Screening condenses this into the Colloidal Instability Index: saturates plus asphaltenes, over resins plus aromatics. Higher means less stable. Three levers move the balance: pressure, composition, temperature — in that order.",
  "Say the order of importance explicitly: pressure first.", None),
 (8, "Asphaltenes are not wax", 40, "4:20",
  "Now the distinction that matters most in this talk. Wax is a cooling story: temperature falls below the wax appearance temperature, crystals form, and they can be melted away. Asphaltenes are a pressure-and-composition story: pressure falls below the onset, or the composition changes, and amorphous particles separate. Amorphous means they do not remelt — which is exactly why thermal methods that work for wax fail here. Different trigger, different physics, different fix.",
  "Three contrasts only; do not read the table.", None),
 (9, "Mechanism chain", 35, "4:55",
  "The chain has four steps. Precipitation — thermodynamic: the oil crosses the envelope. Flocculation — particles stick into clusters. Aggregation — clusters grow. And deposition — particles adhere to steel, balanced by shear removal. And here is the distinction that separates understanding from theory: precipitation is not deposition. Many wells carry precipitated asphaltenes all the way to surface with no deposit at all. Deposition is a transport phenomenon on top of the thermodynamic one — and it is where prediction is weakest today.",
  "The red box is the exam-grade point. Slow down.", None),
 (10, "The envelope", 45, "5:40",
  "This chart is the heart of the science. The dashed line is the onset pressure versus temperature — above it the oil is stable; below it, precipitation is possible. Two details. First, the upper onset can sit above the bubble point by hundreds of bar — thousands of psi in severe oils — and precipitation peaks near the bubble point. Second, keep the screens straight: the de Boer plot is a different screen, in-situ density against gas under-saturation; the onset-versus-bubble-point rule comes from the envelope. And the dark line is the flowing path: as oil rises, pressure and temperature fall together until the crossing — the moment the clock starts.",
  "Trace the path with the pointer: reservoir, crossing star, wellhead.", None),
 (11, "Screening and measurement", 40, "6:20",
  "How do we get these numbers? In two layers. Screening first: SARA analysis — saturates, aromatics, resins, asphaltenes; the Colloidal Instability Index; and the de Boer plot. Then measurement: heptane titration for onset ranking, and — the anchor for any model — a high-pressure onset measurement in a PVT cell, typically by near-infrared or light scattering, giving onset pressure versus temperature. Screening decides if we worry; measurement calibrates the prediction.",
  "Three methods maximum if short: SARA, CII, measured onset curve.", None),
 (12, "Triggers and the misconception", 45, "7:05",
  "What triggers it downhole? The pressure drop is the main trigger. Gas lift changes the solvent — the light gas strips the stabilising shell. Gas breakthrough, CO2 and rich-gas injection shift the envelope. Commingling two stable crudes can create an unstable blend. And a misconception I want to correct: that deposits form at the top of the well, where pressure drops fastest. Not true — the gradient along the tubing is roughly steady, so the deposit forms wherever the flowing path first crosses the envelope. In deep wells that can be hundreds of metres, even kilometres, below the wellhead.",
  "CHECKPOINT 1 at seven oh five. The misconception is the pivot of the talk.", None),
 (13, "The field case", 55, "8:00",
  "And here is why that matters — a documented deepwater Gulf of Mexico well. The perforations sit at seventeen thousand seven hundred to eighteen thousand feet; the chemical injection mandrel at fourteen thousand six hundred. The onset pressure of this fluid sits within reservoir conditions — so the bottom three thousand four hundred feet of tubing sat below the mandrel, with no chemical protection at all. While producing, the effective tubing diameter fell from two point three to one point three five inches. The fix was a solvent soak, then continuous inhibitor — seven hundred fifty parts per million initially, stabilised near five hundred. Result: sixteen percent more production over four months. The lesson is geometric: the well had inhibition — the placement was wrong.",
  "CHECKPOINT 2 at eight minutes. POINT AT THE ≈300 m BRACKET on the figure while you say the third sentence — the gap between the mandrel and the predicted onset is the argument. The punchline is the last sentence.", None),
 (14, "Detection", 35, "8:35",
  "How do we detect it before the well chokes? The earliest fingerprint is a rising pressure drop for the same rate — delta P over Q. Then the wellhead signature: in severe fields, twenty to twenty-five percent pressure loss within weeks. For direct measurement: wireline gauge rings, run successively, map the deposit profile — the classic Hassi Messaoud practice. And in the deepwater case, produced-fluid sampling with optical analysis tuned the inhibitor dose in near-real time.",
  "Name the delta-P/Q fingerprint first — it is the cheapest alarm.", None),
 (15, "Reactive defenses fail", 45, "9:20",
  "The reactive toolkit. Wireline scrapers cut the deposit out; coiled tubing mills the full string; aromatic solvents dissolve it; acid jobs are sometimes used; thermal methods are not a melting story — asphaltene deposits have no melting point, and heat can even weaken the resins that keep them dispersed. Each method works. But look at the pattern: every one is a cleanout. Every cleanout is NPT, deferred production, and a repeat visit scheduled by the physics itself. At Hassi Messaoud, the classic assessment of routine scraping and solvent washing was short and honest: expensive and cumbersome.",
  "Three rows maximum: scrapers, solvents, thermal.", None),
 (16, "Proactive inhibition", 40, "10:00",
  "So the proactive answer: don't remove the deposit — prevent it. Inhibitors and dispersants share one principle — a polar head adsorbs on the aggregate, an alkyl tail keeps it suspended — but they differ: inhibitors shift the flocculation onset, dispersants shrink aggregates, and neither dissolves a deposit. They are injected continuously through a capillary string. Dosage is field-specific — hundreds to a few thousand parts per million — and more is not better: over-dosing can increase deposition. The placement rule from the case study: inject below the deepest credible onset depth, with a margin from the sensitivity study.",
  "Placement rule is the bridge from case study to project.", None),
 (17, "Economics", 40, "10:40",
  "Now, the money. Reactive costs: cleanouts, treatment up to three million dollars, deferred production up to one point two million per day, and a catastrophic tail — seventy million for one deepwater shut-in. Proactive: inhibitor at roughly thirty to forty-six thousand per well per year in the Middle East, against three hundred thirty to three hundred ninety thousand in the Gulf of Mexico. Honest caveat: these ranges span different fields and years — never quote them as one number. What is missing is the breakeven frequency for this specific well. The thesis supplies it.",
  "CHECKPOINT 3 at ten forty. The caveat line builds credibility.", None),
 (18, "Gap analysis", 40, "11:20",
  "So where is the gap? Onset prediction is sensitive to fluid characterisation — small input errors move the depth. Precipitation and deposition are different physics, and transport models are far less mature. Dosage is still chosen by trial and error — the literature calls lab validation inadequate. Injection-point placement is rarely optimised against a predicted onset depth — our case had three thousand four hundred feet unprotected. And no consolidated public workflow runs from screening to economics. That last item is my opportunity.",
  "The last line is the pivot from review to research.", None),
 (19, "Proposed project", 45, "12:05",
  "So this is the project. The aim: build and document a transparent workflow — screening, then PVT modelling, then wellbore analysis — that predicts asphaltene onset depth, selects the inhibitor injection point, and quantifies the economics, validated against a documented field case. Three research questions: how sensitive is the predicted depth to fluid characterisation; what injection point and dosage window keep the whole tubing above the onset; and at what cleanout frequency does inhibition break even. Four work packages, from SARA screening to the economics model.",
  "Read the aim slowly; one line per research question.", None),
 (20, "Methodology", 50, "12:55",
  "The method is four steps, each with an output. Screen: SARA, CII and de Boer on published fluid data — a stability verdict. Model: the onset curve versus temperature, using a CPA or PC-SAFT-class equation of state tuned to measured onset data — the envelope. Predict: the wellbore pressure-temperature profile, the crossing depth, and sensitivity to plus or minus fifty to a hundred bar — onset depth with a margin. Decide: injection point, dosage window and the breakeven model — a decision matrix for the field. Tools: licensed software if the department has it; NeqSim, the open-source fallback, otherwise.",
  "Four steps, four outputs. Walk the three tool-hierarchy chips along the bottom — they prevent the software question. Sensitivity is already quantified: fifty bar on the onset shifts the depth by two hundred and fifty metres, and in the gas-lift scenario the whole tubing goes unstable — the dose, not the placement, does the work.", None),
 (21, "Deliverables and limits", 40, "13:35",
  "The deliverables: the thesis, the documented workflow, the case-study report, the breakeven model, and a practical decision matrix. The limitations and ethics, stated up front: literature-based fluids — no live sample unless a partner provides one; a single case, so conditional conclusions; deposition transport kept first-order; steady-state wellbore profiles; indicative costs; and solvent chemistries compared on paper only — no uncontrolled handling of aromatic solvents and no vendor recommendation. And a sixteen-week plan. Two semesters is enough — precisely because the scope is this narrow.",
  "Limitations delivered confidently, not as an apology.", None),
 (22, "Three messages", 40, "14:15",
  "So — three messages. One: asphaltene plugging is a pressure-and-composition story, not a cooling story — and precipitation is not deposition. Two: the decisive number is the depth where the flowing path crosses the onset envelope — it sets the injection point, the mitigation strategy, and the money. Three: my graduation project builds a transparent screening-to-economics workflow for that depth, validated on a documented field case. Thank you for your attention — I am happy to take your questions.",
  "CONCLUSIONS START HERE. Memorised verbatim. Practise each of the three as a standalone 15-second soundbite, out of order — an examiner may ask for any one of them. Then stop talking.", None),
 (23, "References", 10, "14:25",
  "My references are on this slide, and the detailed list is in the handout.",
  "One sentence. Skip entirely if behind schedule.", "If behind: skip."),
]

sc = new_doc("Speaking Script — Asphaltenes seminar")
p = sc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("WORD-FOR-WORD SPEAKING SCRIPT — 15 MINUTES"); r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = TEAL
p = sc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Asphaltenes: Predicting and Preventing Wellbore Plugging"); r.font.size = Pt(20); r.font.bold = True; r.font.color.rgb = NAVY
p = sc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Companion to Asphaltene_Slides_23plus2.pptx · designed for a 14:25 speaking budget"); r.font.size = Pt(11); r.font.color.rgb = GRAY
para(sc, "How to use: read it, cut it to your voice, then stop reading it — by rehearsal 3 you work from the cue card, with this script as the safety net. Numbers are written phonetically on purpose: you say 'two point three to one point three five', never mumble a decimal. Memorise the five scripted lines flagged in the Master Plan.", size=10, italic=True, color=GRAY)

total_words = 0
for num, title, tg, cl, say, do, behind in script:
    sc.add_heading(f"Slide {num} — {title}   (target {tg} s · clock {cl})", level=3)
    rich(sc, [("SAY:  ", True, False, NAVY), (say, False, False, None)], size=10.5, space_after=4)
    rich(sc, [("DO:  ", True, False, TEAL), (do, False, False, None)], size=10, space_after=2)
    if behind:
        rich(sc, [("IF BEHIND:  ", True, False, RED), (behind, False, False, None)], size=10, space_after=6)
    total_words += len(say.split())

sc.add_heading("Final counts", level=2)
para(sc, f"Spoken words in this script: {total_words}. At a rehearsed pace of 128–132 words/minute this is ≈ {total_words/130:.1f} minutes of speech; with pauses for numbers the delivery lands at ≈ 14:20–14:45 — inside the hard stop, with the references slide available to sacrifice if needed.", size=10.5)
para(sc, "Checkpoints to verify with a stopwatch in rehearsal: slide 12 at 7:05 · slide 13 at 8:00 · slide 17 at 10:40 · slide 22 at 14:15. If any checkpoint is more than 20 seconds late, apply the cut list — never improvise a shortened version live.", size=10.5, bold=True, color=NAVY)
sc.save("/home/user/Seminar_Asphaltenes/01_Asphaltene_Speaking_Script.docx")
print("saved script; words =", total_words, "->", round(total_words/130, 2), "min")

# =====================================================================
# DOC 3 — Q&A PACK
# =====================================================================
qa = new_doc("Q&A Defence Pack — Asphaltenes")
p = qa.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Q&A DEFENCE PACK — 31 QUESTIONS WITH MODEL ANSWERS"); r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = TEAL
p = qa.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Asphaltenes: Predicting and Preventing Wellbore Plugging"); r.font.size = Pt(20); r.font.bold = True; r.font.color.rgb = NAVY
qa.add_heading("How to run the question session", level=2)
for t in [
    "Answer in TWO sentences, then stop. The next question is often easier than the answer you would have improvised.",
    "Never bluff a number. Correct sentence: “I don't have that figure in front of me — I will verify it and send it to you.” Then do it, the same day.",
    "If a question leaves your scope, bridge: “That is outside my declared scope — what I can tell you is…” then one solid in-scope fact.",
    "If hostile: agree with the strongest part of the point, then state what your work does cover. Never argue.",
    "Keep a pen at the lectern; writing a question down buys two seconds of thinking time.",
]:
    bullet(qa, t, 10)

def qa_item(num, section, q, a, caution=None):
    qa.add_heading(f"Q{num}.  {q}", level=3)
    rich(qa, [("Model answer:  ", True, False, TEAL), (a, False, False, None)], size=10.5, space_after=4)
    if caution:
        rich(qa, [("Caution:  ", True, False, AMBER), (caution, False, False, None)], size=9.5, space_after=6)

qa.add_heading("A.  Chemistry and mechanisms", level=1)
qa_item(1, "A", "What exactly are asphaltenes — one molecule?",
    "No — they are a solubility class, defined operationally: insoluble in n-heptane and soluble in toluene, quantified by SARA analysis (ASTM D6560 / IP 143). Chemically they are the heaviest, most polar, most aromatic fraction of the crude, existing as nano-aggregates rather than single molecules.")
qa_item(2, "A", "Why do asphaltenes precipitate when we gas-lift the well?",
    "The injected light gas lowers the liquid phase's solvency and can strip the stabilising resin/aromatic fraction — the colloid loses its protective shell, so aggregates flocculate and drop out. It is a composition change: gas lift makes the oil a poorer solvent for its own heaviest fraction.")
qa_item(3, "A", "Does temperature cause it, like wax?",
    "Not primarily. For wax, temperature falling below the WAT is the trigger; for asphaltenes the trigger is pressure falling below the onset, or a composition change — temperature modifies the envelope rather than driving the event. And because asphaltene deposits are amorphous, they do not remelt, so thermal methods that work on wax fail here.")
qa_item(4, "A", "What is the difference between precipitation and deposition?",
    "Precipitation is thermodynamic — asphaltenes leave solution when the oil crosses the envelope. Deposition is a transport phenomenon on top of it: flocculated particles must reach the wall, adhere, and survive shear removal. Many wells produce precipitated asphaltenes with no deposit — which is exactly why deposition prediction is the weakest link.")
qa_item(5, "A", "What is the Colloidal Instability Index?",
    "A screening index: CII = (Saturates + Asphaltenes) / (Resins + Aromatics). Higher CII means less stable crude; the commonly cited threshold is around 0.9, though the exact boundary is laboratory-specific — which is why I always pair it with a measured onset curve.")
qa_item(6, "A", "What is the de Boer plot and why use it?",
    "It is a cross-plot of in-situ oil density (x) against the degree of gas under-saturation — reservoir pressure minus bubble-point pressure (y) — with three risk zones: severe, moderate and minimal. The 1995 paper concludes that risk is greatest for LIGHT, undersaturated crudes with only a small asphaltene content. It needs nothing but a PVT report, which is why it is still used as a pre-screen.",
    "Caution: do NOT describe it as an AOP-versus-bubble-point cross-plot — that is a common and easily-caught error. If challenged: 'the de Boer plot uses density and under-saturation; the onset-above-bubble-point observation is a separate envelope-based screen, and I keep the two distinct.'")
qa_item(7, "A", "Why is the onset pressure often above the bubble point?",
    "Because asphaltene stability collapses before the lightest fractions come out of solution — the fluid loses solvency as pressure drops and the colloidal shell fails first. The size of that window is fluid-specific: hundreds of bar in many oils, thousands of psi in severe ones. Precipitation peaks as pressure approaches the bubble point, and some asphaltenes re-dissolve below it.",
    "Caution: never quote '100–200 bar' as a general figure — published windows span a few hundred bar to several thousand psi. Say 'fluid-specific — hundreds of bar or more'.")
qa_item(8, "A", "What about CO₂ injection — does it make this worse?",
    "Yes. CO₂ is a well-documented destabiliser: it changes the solvent character of the oil and shifts the onset envelope upward, so CO₂ EOR and CO₂ storage projects on asphaltenic fields must screen for it. My workflow applies directly to those projects — it is one of the reasons the topic is current.")
qa_item(9, "A", "Are there 'lower onset pressures' — a second envelope?",
    "Yes — some fluids show a second, lower onset region at low pressure, which is why the full envelope can be closed rather than monotonic. It sits outside the scope of this seminar but is noted on the backup slide; the thesis will state whether the case fluid shows it.")

qa.add_heading("B.  Field operations and defenses", level=1)
qa_item(10, "B", "What happens if we just choke the well to keep pressure high?",
    "Choking keeps the flowing pressure above the onset but sacrifices rate — so it trades production for prevention, and the economics usually suffer. The more interesting answer is that Hassi Messaoud found the opposite lever as well: producing at low wellhead pressure drastically reduced the number of cleanouts, because it changed the P–T path and flow regime. The honest conclusion is that the right operating window is well-specific — which is exactly what my project quantifies.")
qa_item(11, "B", "Why do the deposits form so deep, not at the wellhead?",
    "Because the pressure gradient along the tubing is roughly steady — there is no special 'fastest drop' point. The deposit appears where the flowing P–T path first crosses the onset envelope, and in deep, hot, high-GOR wells that crossing can sit hundreds of metres or more below the wellhead. The documented Gulf of Mexico case is the proof — the deposit zone lay below the chemical injection mandrel.")
qa_item(12, "B", "Why do solvent washes only give a temporary fix?",
    "A solvent dissolves what is already there but changes nothing about why it formed — the flowing conditions remain below the onset, so deposition restarts as soon as production resumes. Add HSE exposure (BTEX), flammability and cost, and it becomes a maintenance cycle rather than a solution. Prevention means keeping the particles dispersed, not removing them after the fact.")
qa_item(13, "B", "Are inhibitor chemicals just solvents by another name?",
    "No — that is the key distinction. Solvents dissolve existing deposits; inhibitors and dispersants keep asphaltenes suspended so they never deposit. Within the prevention family there is a further distinction worth making: strictly, INHIBITORS shift the flocculation onset, while DISPERSANTS reduce the size of aggregates already formed — an inhibitor can act as a dispersant, but not necessarily the reverse (Kelland, 2009).",
    "Caution: if you call them all 'dispersants', an examiner may correct you. Answer: 'both are prevention chemistry; the difference is onset-shift versus aggregate-size reduction.'")
qa_item(14, "B", "What determines the inhibitor dose?",
    "Fluid-specific lab work plus field monitoring: minimum dosage is found by titration/optical-tube testing and refined with field sampling. In the documented case the well started at 750 ppm and stabilised around 500 ppm. Dosing below the effective threshold usually means deposition continues slowly — which is why the case study used continuous sampling to tune it.")
qa_item(15, "B", "Where should the injection point be placed?",
    "Below the deepest credible onset depth — onset plus a margin sized by the sensitivity study, not by habit. The requirement is to dose the oil before it enters the unstable window. In the documented case the mandrel sat about 300 m ABOVE the predicted onset, so the deepest 3,400 ft had no protection. And where the sensitivity shows the whole tubing going unstable — the gas-lift scenario — placement alone cannot protect, so the dose window does the work.")
qa_item(16, "B", "How do we know it is asphaltene and not wax or scale?",
    "Sample the deposit: asphaltene deposits are dark, hard and solvent-soluble in aromatics; wax is soft and melts; scale is mineral. Production signatures also differ in temperature sensitivity. The well-integrity literature recommends confirming with deposit analysis before choosing treatment — the wrong diagnosis means the wrong chemistry.")
qa_item(17, "B", "What about mechanical cleanouts — are they obsolete?",
    "No — they remain necessary for severe blockages, and coiled-tubing cleanouts were still part of the documented GoM well's surveillance strategy alongside continuous inhibition. But as a replacement for prevention they are uneconomic: each cleanout is NPT and deferred production, and the deposit returns.")

qa.add_heading("C.  Thesis scope, data, tools and validation", level=1)
qa_item(18, "C", "Is one case study enough for a graduation project?",
    "The case study validates the method; the method is the contribution. One fully documented and reproducible prediction workflow, plus a sensitivity analysis around it, is a defensible Bachelor's scope — and I state the boundary honestly rather than pretending broader coverage.")
qa_item(19, "C", "What data will you actually use?",
    "Published fluid and field data: the documented Gulf of Mexico case, Hassi Messaoud data, and PVT/SARA reports from the literature, with NeqSim as an open-source modelling fallback. Week 1 locks the data set and verifies availability before modelling starts — so the project cannot stall on data.")
qa_item(20, "C", "What software will you use — and what if licences don't exist?",
    "The hierarchy is declared in advance: licensed PVT and flow-assurance tools if the department has them (PVTsim/Multiflash-class for onset, OLGA/PIPESIM-class for wellbore), otherwise NeqSim — an open-source CPA equation-of-state toolkit with asphaltene screening — plus correlation-based screening as the floor. That is a designed feature of the project, not a fallback plan.")
qa_item(21, "C", "How will you validate the predictions?",
    "Three layers. Reproduce the documented case's observed behavior — onset below the existing mandrel, the ID-reduction trend, dosage within the field window — within a stated tolerance. Cross-check onset against an independent model or correlation. And check internal consistency: the predicted depth must respond correctly to pressure and GOR changes.")
qa_item(22, "C", "What is genuinely new here — isn't this textbook material?",
    "At Bachelor's level, novelty is documented rigour applied to a documented gap. The specific contribution is a transparent, parameterised workflow that connects SARA screening, an EoS onset model, wellbore crossing-depth prediction, injection-point placement and breakeven economics — applied end to end on a real published case with an auditable trail. That consolidated artifact is not publicly available.")
qa_item(23, "C", "You're not modelling deposition thickness — why not?",
    "Because deposition-transport modelling depends on kinetic and adhesion parameters that must be measured, and my scope is honest about that boundary. I model the thermodynamic onset depth — the quantity that decides injection placement — and treat deposition growth as first-order plus sensitivity. It is declared, not accidental.")
qa_item(24, "C", "What is the main risk to your schedule?",
    "Data and licences — both mitigated in Week 1 (data locked; tool hierarchy declared). The second risk is scope creep into deposition-transport physics, mitigated by the frozen scope statement and the 16-week plan with a weekly deliverable.")
qa_item(25, "C", "Why does this belong in a petroleum engineering degree?",
    "Because flow assurance is core petroleum engineering: it decides whether the well you designed and completed actually produces. This project connects production chemistry, PVT thermodynamics and wellbore hydraulics in one workflow — precisely the skill set the industry hires for.")

qa.add_heading("D.  Sources and the numbers (Q&A armour)", level=1)
qa_item(26, "D", "Two reviews give different countries for the same inhibitor-cost figure. Which is right?",
    "The ranges trace to Cenegy's 2001 worldwide inhibitor survey. Fluid Phase Equilibria (2020) and the Elsevier monograph (2021) assign the low range to the Middle East and the high range to the Gulf of Mexico; the Energy & Fuels review (2021) states the reverse. I quote the majority attribution on the slide and flag the discrepancy — and confirming Cenegy's original wording is a Week-1 library task in my source log.",
    "This answer is stronger than picking a side: it demonstrates source discipline. Keep 04_Source_Log.docx at the lectern.")
qa_item(27, "D", "How do you define asphaltenes — and with which standards?",
    "Operationally, not chemically: insoluble in n-heptane, soluble in toluene — quantified by ASTM D6560 / IP 143, with SARA fractionation by ASTM D2007-class methods and onset titrations by ASTM D6703 (automated Heithaus). I cite the edition years I actually consult.",
    None)
qa_item(28, "D", "Why did the field case dose at 750 ppm and then drop to ~500 ppm?",
    "750 ppm was the high-dose start used to re-stabilise the well after the soak; continuous produced-fluid monitoring showed the stabilised dose could be reduced to about 500 ppm without losing protection. Dosing below the effective threshold lets deposition continue slowly — which is why the thesis treats the dosage window as an output, not an assumption.",
    None)

qa_item(29, "D", "Isn't the de Boer plot just AOP plotted against bubble point?",
    "No — that is a widespread shorthand, and it is wrong. The de Boer plot is in-situ density versus gas under-saturation (P_res minus Pb), with three risk zones. The onset-above-bubble-point rule is a separate, envelope-based screen. I keep both in the workflow precisely because they can disagree, and disagreement is itself information.",
    "Caution: this is the most likely 'gotcha' in the session — the correction is already printed on backup slide 24, so you can point at it.")
qa_item(30, "D", "Steam has been used successfully on asphaltene damage — doesn't that contradict your thermal claim?",
    "It is a fair challenge. Reports exist of high permeability recovery after ~497 °C steam treatment — but note what the same work flags: at that temperature asphaltene does not melt, it decomposes into solid-like material that can itself damage the formation. Thermal energy is not dissolving the deposit the way it melts wax; reported successes come from other mechanisms, and heat can even weaken the resins that keep asphaltenes dispersed. My claim on the slide is narrow and defensible: thermal is not the asphaltene tool the way it is the wax tool.",
    "Caution: do not claim thermal 'never works' — claim it is not a melting mechanism and is not the standard remedy.")
qa_item(31, "D", "If a little inhibitor helps, doesn't more help more?",
    "No — and this is a real operational trap. In heptane-rich (strongly destabilised) systems, over-dosing above the recommended concentration has been reported to INCREASE deposition, and at high concentration inhibitor molecules can self-aggregate and lose adsorption efficiency. That is why the thesis delivers a dosage WINDOW with a placement margin, not a single number.",
    "Caution: if asked for a number, give the documented case only — 750 ppm initial, stabilised near 500 ppm — and label it as that well's outcome.")
qa.add_heading("The three questions to hope for", level=2)
for t in ["Q2 (gas lift) — you have a clean two-sentence physics answer.",
          "Q10 (choking) — you have the standard answer AND the Hassi Messaoud counter-answer; this shows depth.",
          "Q22 (what is new) — 'documented rigour plus a consolidated public workflow', with the case study as proof."]:
    bullet(qa, t, 10)
para(qa, "Read this pack ONCE the evening before and ONCE the morning of the seminar. The rules: two sentences, then stop; never bluff a number; redirect, never argue.", size=10.5, bold=True, color=NAVY)
qa.save("/home/user/Seminar_Asphaltenes/02_Asphaltene_QandA_Pack.docx")
print("saved Q&A pack")
