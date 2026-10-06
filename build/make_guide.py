#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Builds 03_Animation_and_Design_Guide.docx"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

NAVY = RGBColor(0x10,0x2A,0x43); TEAL = RGBColor(0x1F,0x7A,0x84)
GRAY = RGBColor(0x5A,0x66,0x72); RED = RGBColor(0xB4,0x3B,0x2E); AMBER = RGBColor(0xA8,0x6B,0x00)

doc = Document()
st = doc.styles["Normal"]; st.font.name="Calibri"; st.font.size=Pt(10.5)
st.paragraph_format.space_after=Pt(6); st.paragraph_format.line_spacing=1.08
for name,size,color in (("Heading 1",16,NAVY),("Heading 2",13,TEAL),("Heading 3",11.5,NAVY)):
    s=doc.styles[name]; s.font.name="Calibri"; s.font.size=Pt(size); s.font.bold=True
    s.font.color.rgb=color; s.paragraph_format.space_before=Pt(12); s.paragraph_format.space_after=Pt(4)
sec=doc.sections[0]; sec.top_margin=Cm(1.9); sec.bottom_margin=Cm(1.9); sec.left_margin=Cm(2.0); sec.right_margin=Cm(2.0)
p=sec.footer.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
r=p.add_run("Animation & Design Guide — Asphaltenes seminar"); r.font.size=Pt(8); r.font.color.rgb=GRAY

def para(text,size=10.5,bold=False,italic=False,color=None,space_after=6):
    pp=doc.add_paragraph(); pp.paragraph_format.space_after=Pt(space_after)
    rr=pp.add_run(text); rr.font.size=Pt(size); rr.font.bold=bold; rr.font.italic=italic
    if color: rr.font.color.rgb=color
    return pp

def bullet(text,size=10.5):
    pp=doc.add_paragraph(style="List Bullet"); pp.paragraph_format.space_after=Pt(3)
    rr=pp.add_run(text); rr.font.size=Pt(size); return pp

def shade(cell,hexc):
    tcPr=cell._tc.get_or_add_tcPr(); shd=OxmlElement("w:shd")
    shd.set(qn("w:val"),"clear"); shd.set(qn("w:color"),"auto"); shd.set(qn("w:fill"),hexc); tcPr.append(shd)

def tbl(rows,widths=None,fs=9.5):
    t=doc.add_table(rows=len(rows),cols=len(rows[0])); t.style="Table Grid"
    t.alignment=WD_TABLE_ALIGNMENT.CENTER
    for ri,row in enumerate(rows):
        for ci,val in enumerate(row):
            cell=t.cell(ri,ci); cell.text=""
            pp=cell.paragraphs[0]; pp.paragraph_format.space_after=Pt(2)
            rr=pp.add_run(str(val)); rr.font.size=Pt(fs); rr.font.name="Calibri"
            if ri==0: rr.font.bold=True; rr.font.color.rgb=RGBColor(0xFF,0xFF,0xFF); shade(cell,"102A43")
            elif ri%2==0: shade(cell,"F1F5F8")
    if widths:
        total=sum(widths)
        for ci,w in enumerate(widths):
            for ri in range(len(rows)): t.cell(ri,ci).width=Cm(17.0*w/total)
    doc.add_paragraph().paragraph_format.space_after=Pt(2)
    return t

p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
r=p.add_run("ANIMATION MAP & DESIGN GUIDE"); r.font.size=Pt(12); r.font.bold=True; r.font.color.rgb=TEAL
p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
r=p.add_run("Asphaltene deck — boardroom styling, build order and software workflow"); r.font.size=Pt(18); r.font.bold=True; r.font.color.rgb=NAVY

doc.add_heading("1.  About the software (honest answer to your question)", level=1)
para("The deck was generated programmatically as a native PowerPoint file (Python + python-pptx from a build script), not drawn in PowerPoint, Canva, or LaTeX. What that means for you:", size=10.5)
for t in [
    "You receive a standard .pptx — it opens in PowerPoint and is fully editable (every shape, text box, table and picture is a native object).",
    "Why not LaTeX/Beamer: Beamer cannot be edited by your supervisor, and university committees increasingly expect an editable PPTX.",
    "Why not Canva: Canva exports are layout-locked, degrade on projector aspect ratios, and cannot hold native chart objects.",
    "The one honest trade-off: click-build animations cannot be written reliably by the generator (hand-injected timing XML corrupts slides more often than it works). Slide transitions and builds must be applied once, in PowerPoint, using the map in §2 — budget 15 minutes.",
    "The figures (charts, cross-section, colloid schematic) are exported at 200 dpi as standalone PNGs in _build/figs/ — reuse them in the thesis and in the thesis defence.",
]:
    bullet(t)

doc.add_heading("2.  Animation map (apply in PowerPoint: Animations → Appear / Fade, 'On Click')", level=1)
para("Principle: one build per idea you speak. Never animate decoration; animate content the audience should receive with your voice.", size=10, italic=True, color=GRAY)
tbl([
    ["Slide", "What to build", "When to trigger", "Why it matters"],
    ["3 — What are asphaltenes", "Cover the 2nd and 3rd panels of the colloid schematic with white rectangles; reveal panel 2, then panel 3", "On your words 'destabilised' and 'flocculated'", "The highest-value build in the deck: the audience watches the shell strip away as you explain it"],
    ["4 — Cross-section", "Reveal the 'AFTER' half of the drawing (group the deposit ring + its callouts)", "After describing the clean bore", "Before/after is a story; showing both at once kills it"],
    ["9 — Mechanism chain", "Reveal the four step boxes one at a time", "One per step as you speak it", "Forces the jury to follow your pacing, not read ahead"],
    ["10 — Envelope chart", "Reveal in three clicks: (1) axes + AOP orange line, (2) bubble point + shaded zone, (3) flowing path with arrow + the star", "Each layer with its explanation", "The path crossing the star is the moment the talk clicks — let them see it happen"],
    ["12 — Triggers", "Reveal the four trigger boxes first; keep the misconception box hidden until its cue", "Misconception box after the four triggers", "A reveal makes a correction feel like a discovery, not a lecture"],
    ["13 — Case study", "Reveal chart layers; then the amber unprotected-interval band last", "Band last — it is the punchline", "The 3,400 ft band appearing over the star is the visual proof"],
    ["16 — Proactive", "Reveal the placement rule bullet last", "With its sentence", "It links the case study to your project"],
    ["17 — Economics", "Reveal the breakeven box after the table", "When you say 'the missing number'", "Turns data into a research question"],
    ["22 — Three messages", "Fade each message on click", "One per message", "Your close — control the pacing absolutely"],
], widths=[3.4, 6.4, 3.6, 4.6], fs=9)
para("Do NOT animate: titles, kickers, footers, section labels, table rows (except slides 9 and 17), or any transition longer than 0.3 s.", size=10, bold=True, color=RED)

doc.add_heading("3.  Design system already applied (so you can defend it)", level=1)
tbl([
    ["Element", "Specification in this deck"],
    ["Background / text", "White background (lecture-hall friendly); deep navy #102A43 text — never pure black"],
    ["Alert colour", "Safety-orange #E8630A used ONLY for the AOP line, the onset star, the instability zone and the deposit"],
    ["Secondary colour", "Teal #2E9BA6 = the benign/stabilising elements (resins, injection point, stable reservoir)"],
    ["Action titles", "Every slide title is the takeaway sentence, 22–24 pt bold (see §4)"],
    ["Section indicator", "Top-left kicker on every slide ('PART 2 · THE SCIENCE') — placed at the top so the eye catches it on entry, acting as the progress bar"],
    ["Footers", "Slide number, target seconds and cumulative clock — your on-stage timing instrument"],
    ["Charts", "No background gridlines; orange dashed AOP; blue bubble point; navy bold flowing path with arrowhead; depth axis increases DOWNWARD"],
    ["Equations", "Native text runs, colour-coded: orange numerator = destabilising, teal denominator = natural dispersant"],
    ["Body type", "12–14 pt for the audience-facing text, 22–24 pt titles. The 10-foot rule beats an arbitrary point size: this deck passes it"],
], widths=[3.6, 13.4], fs=9.5)

doc.add_heading("4.  The 10-foot test (your pre-flight check)", level=1)
para("Print slides 3, 4, 10, 13 and 17, tape them to the wall, and stand 3 metres away. You must be able to read, from that distance: the action title, the axis labels, and the single number each slide exists for. If you cannot, do not shrink the point size — delete a bullet. Every slide in this deck was designed so that exactly one fact survives the squint test.", size=10.5)

doc.add_heading("5.  If you want to extend the design yourself in PowerPoint", level=1)
for t in [
    "Slide Master (View → Slide Master): edit the kicker or footer text once to change it on all 25 slides.",
    "Colour discipline: add the four brand colours (navy, orange, teal, amber) to your Theme Colours so every new shape matches automatically.",
    "Any new chart you add: orange dashed for AOP/onset, navy bold for the flowing path, no gridlines, depth downward.",
    "Keep the sentence-as-title habit: if you add a slide, its title should be a claim, not a label.",
]:
    bullet(t, 10)
doc.save("/home/user/Seminar_Asphaltenes/03_Animation_and_Design_Guide.docx")
print("saved guide")
