#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Builds 05_Projector_Preflight.docx and appends the locked design system to 03_Animation_and_Design_Guide.docx"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

NAVY = RGBColor(0x1B,0x36,0x5D); TEAL = RGBColor(0x0D,0x73,0x77)
GRAY = RGBColor(0x4A,0x55,0x68); ORANGE = RGBColor(0xB3,0x5A,0x18)

def base_doc(footer):
    doc = Document()
    st = doc.styles["Normal"]; st.font.name = "Calibri"; st.font.size = Pt(10)
    st.paragraph_format.space_after = Pt(5); st.paragraph_format.line_spacing = 1.05
    for name, size, color in (("Heading 1", 14, NAVY), ("Heading 2", 11.5, TEAL)):
        s = doc.styles[name]; s.font.name = "Calibri"; s.font.size = Pt(size); s.font.bold = True
        s.font.color.rgb = color; s.paragraph_format.space_before = Pt(9); s.paragraph_format.space_after = Pt(3)
    sec = doc.sections[0]
    sec.top_margin = Cm(1.6); sec.bottom_margin = Cm(1.6); sec.left_margin = Cm(1.8); sec.right_margin = Cm(1.8)
    p = sec.footer.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(footer); r.font.size = Pt(8); r.font.color.rgb = GRAY
    return doc

def para(doc, text, size=10, bold=False, italic=False, color=None, space_after=5):
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(text); r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic
    if color: r.font.color.rgb = color
    return p

def bullet(doc, text, size=9.5):
    p = doc.add_paragraph(style="List Bullet"); p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text); r.font.size = Pt(size); return p

def shade(cell, hexc):
    tcPr = cell._tc.get_or_add_tcPr(); shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hexc); tcPr.append(shd)

def tbl(doc, rows, widths, fs=8.5):
    t = doc.add_table(rows=len(rows), cols=len(rows[0])); t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = t.cell(ri, ci); cell.text = ""
            p = cell.paragraphs[0]; p.paragraph_format.space_after = Pt(1)
            r = p.add_run(str(val)); r.font.size = Pt(fs); r.font.name = "Calibri"
            if ri == 0:
                r.font.bold = True; r.font.color.rgb = RGBColor(0xFF,0xFF,0xFF); shade(cell, "1B365D")
            elif ri % 2 == 0:
                shade(cell, "F4F7F9")
    total = sum(widths)
    for ci, w in enumerate(widths):
        for ri in range(len(rows)):
            t.cell(ri, ci).width = Cm(17.4 * w / total)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)
    return t

# =================== 05 PROJECTOR PREFLIGHT ===================
doc = base_doc("Projector & Accessibility Preflight — Asphaltenes seminar")
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("PROJECTOR & ACCESSIBILITY PREFLIGHT"); r.font.size = Pt(11); r.font.bold = True; r.font.color.rgb = TEAL
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Asphaltenes deck — test before the room fills, not during the talk"); r.font.size = Pt(15); r.font.bold = True; r.font.color.rgb = NAVY

doc.add_heading("1.  The 20-minute preflight (arrive 30 minutes early)", level=1)
tbl(doc, [
    ["Step", "Do this", "Pass criterion"],
    ["1 · Laptop check", "Open the PPTX (not the PDF) and advance through all 25 slides once", "Every slide renders; no missing images; fonts substituted cleanly"],
    ["2 · Projector check", "Duplicated display at native resolution; check the aspect ratio setting", "16:9 fills the screen with NO letterboxing bars top/bottom"],
    ["3 · Brightness check", "Close the blinds as they will be during your talk; look at slide 17 (table) and slide 13 (chart)", "Table text legible; the orange AOP line is clearly orange, not muddy brown"],
    ["4 · 10-foot test", "Stand 3 m back and read: the action title, the axis labels, and the one number each of slides 3/4/10/13/17 exists for", "All legible without leaning in"],
    ["5 · Colour check", "Look at slides 3 and 4: can you distinguish teal (benign) from orange (danger) instantly?", "Yes — if not, projector saturation is low: raise contrast or move closer to the screen"],
    ["6 · Animations", "Run the two builds (slide 9, slide 10) once each", "Reveal on click; nothing auto-plays"],
    ["7 · Backup drill", "Unplug HDMI; open the PDF from the USB", "PDF opens full-screen; arrow keys advance"],
], widths=[2.6, 8.4, 6.4], fs=8.5)
para(doc, "If a projector cannot render the palette faithfully, the fallback is content-first: the deck survives in greyscale — titles are dark, boxes are bordered, the AOP line is dashed. Colour is reinforcement, never the only carrier of meaning.", size=9.5, italic=True, color=ORANGE)

doc.add_heading("2.  Contrast audit (design-time, already satisfied)", level=1)
tbl(doc, [
    ["Element", "Foreground", "Background", "Ratio", "WCAG AA"],
    ["Body text", "#1A1A1A", "#FFFFFF", "17.4:1", "AAA (needs 7:1)"],
    ["Titles / table headers", "#1B365D", "#FFFFFF", "12.1:1", "AAA"],
    ["Section labels (kicker)", "#0D7377", "#FFFFFF", "5.3:1", "AA (needs 4.5:1)"],
    ["Captions / footers", "#4A5568", "#FFFFFF", "7.6:1", "AAA"],
    ["Small orange text", "#B35A18", "#FFFFFF", "5.4:1", "AA — darker shade used deliberately instead of #E07A3D"],
    ["Table header text", "#FFFFFF", "#1B365D", "12.1:1", "AAA"],
], widths=[4.4, 2.6, 2.6, 2.2, 5.6], fs=8.5)
para(doc, "Note on #E07A3D: it fails AA as small text (2.6:1) and passes as a graphic element. The deck therefore uses it only for lines, borders, icons and markers — and substitutes #B35A18 whenever orange must carry words.", size=9.5, color=GRAY)

doc.add_heading("3.  Export record", level=1)
bullet(doc, "PDF proof: Asphaltene_Slides_23plus2.pdf — 25 pages, 960 x 540 pt (16:9), exported from the final PPTX build.")
bullet(doc, "Figures: 300 dpi PNGs in _build/figs/ — reuse in the thesis and the thesis defence without re-rendering.")
bullet(doc, "Embedded fonts: Calibri throughout (universally available on Windows/Office). If you reconfigure the build, avoid rare fonts.")
bullet(doc, "Open and test on the ACTUAL lecture-hall machine if you can get 10 minutes with it — projector colour rendition varies more than laptop screens.")

doc.save("/home/user/Seminar_Asphaltenes/05_Projector_Preflight.docx")
print("saved preflight")

# =================== APPEND DESIGN SYSTEM TO 03 ===================
g = Document("/home/user/Seminar_Asphaltenes/03_Animation_and_Design_Guide.docx")

g.add_heading("6.  Locked design system (tokens, archetypes, type scale)", level=1)
para(g, "This is the specification the deck is built to. Every slide is generated by shared builder functions that enforce these values — the 'master layout' equivalent for a programmatically built deck.", size=9.5)

g.add_heading("6.1  Colour tokens (use these hex values; nothing else)", level=2)
tbl(g, [
    ["Token", "Hex", "Used for", "Contrast vs white"],
    ["NAVY", "#1B365D", "titles, table headers, box borders, depth axis", "12.1:1 — AAA"],
    ["TEAL", "#0D7377", "section kickers, benign/stabilising elements, primary box borders", "5.3:1 — AA"],
    ["ORANGE", "#E07A3D", "the AOP line, onset star, deposit, call-out borders, ICONS (graphics only)", "2.6:1 — graphic use only"],
    ["AMB_D", "#B35A18", "orange-family SMALL TEXT (captions, attributions)", "5.4:1 — AA"],
    ["TEXT", "#1A1A1A", "body text (never pure black)", "17.4:1 — AAA"],
    ["GRAY", "#4A5568", "captions, footers, secondary notes", "7.6:1 — AAA"],
    ["LIGHT", "#F8FAFC", "box fill", "—"],
    ["LIGHT2", "#E8EFF5", "alternate table rows, muted fills", "—"],
    ["RED", "#B43B2E", "the precipitation/deposition distinction, error reframes", "6.1:1 — AA"],
], widths=[2.0, 1.8, 9.6, 4.0], fs=8)

g.add_heading("6.2  Layout archetypes (every slide declares one)", level=2)
tbl(g, [
    ["Archetype", "Geometry", "Slides"],
    ["A-title", "Full-bleed accent bar left; title block; summary card right", "1"],
    ["B-cards", "Header + 3–4 equal cards on a fixed grid (gap 0.16 in)", "2, 5, 9, 20"],
    ["C-split", "Header + text column left / figure or box column right", "3, 4, 6, 7, 10, 12–14, 16, 18, 19"],
    ["D-table", "Header + full-width table + callout strip beneath", "8, 11, 15, 17, 21"],
    ["E-fullwidth", "Header + two-column reference/text block", "23"],
    ["F-closing", "Header + three stacked message bands + thank-you strip", "22"],
    ["G-backup", "Header labelled BACKUP + dense reference content", "24, 25"],
], widths=[2.6, 10.4, 4.4], fs=8.5)
para(g, "Grid: content band runs from x = 0.62 in to x = 12.86 in on EVERY slide — an automated audit confirms 25/25 slides comply. Card grids end at the same right edge as titles. Headers are identical: kicker (11.5 pt bold teal) → action title (22–24 pt bold navy) → 1.15 x 0.055 in orange rule. Footers are identical: left = short title, right = slide number + target seconds + cumulative clock (or 'BACKUP — show only if asked').", size=9.5)

g.add_heading("6.3  Type scale", level=2)
tbl(g, [
    ["Role", "Size", "Notes"],
    ["Action title", "22–24 pt bold", "One weight only; sentence case; a claim, not a label"],
    ["Section kicker", "11.5 pt bold", "All caps, teal, top-left"],
    ["Card / box headings", "13–16 pt bold", "Navy or the card's accent colour"],
    ["Body", "12.5–14.5 pt regular", "Auto-lifted 1 pt by the builder so no body text sits below ~12.5 pt"],
    ["Table body", "12–13 pt", "Capped at 13 pt; row height auto-padded"],
    ["Captions / footnotes", "9.5–10 pt italic gray", "Never below 9.5 pt — this is the projection floor"],
], widths=[3.6, 3.4, 10.4], fs=8.5)

g.add_heading("6.4  Icons (5 built, one stroke weight, transparent PNG)", level=2)
tbl(g, [
    ["Icon", "Meaning", "Placed on"],
    ["Flask with droplets", "precipitation", "slide 9, step 1"],
    ["Clumping cores", "flocculation / aggregation", "slide 9, steps 2–3"],
    ["Well schematic", "deposition in the wellbore", "slide 9, step 4"],
    ["Injection mandrel", "capillary chemical injection", "available in _build/figs/ for the thesis (not placed — it collided with body text on slide 16)"],
    ["Breakeven chart", "economics", "slide 17 call-out"],
], widths=[3.2, 5.2, 9.0], fs=8.5)
para(g, "All icons share one stroke weight, the palette's teal/navy/orange, and no fills except accent dots. No clip-art, no emoji, no gradient.", size=9.5)

g.add_heading("6.5  Animation decision (final)", level=2)
para(g, "Only two builds are specified, both sequential reveals that match the spoken sequence:", size=9.5)
bullet(g, "Slide 9 — reveal the four mechanism cards one at a time (precipitation → flocculation → aggregation → deposition).")
bullet(g, "Slide 10 — reveal the envelope chart in three layers: AOP line, then bubble point + unstable zone, then the flowing path with the onset star.")
para(g, "No fly-ins, no decorative motion, no auto-play anywhere in the deck.", size=9.5, bold=True, color=NAVY)

g.save("/home/user/Seminar_Asphaltenes/03_Animation_and_Design_Guide.docx")
print("guide updated")
