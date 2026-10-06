# Seminar Package — Asphaltenes: Predicting and Preventing Wellbore Plugging

**Course:** Seminar (BSc Petroleum and Gas Engineering) · 15-minute presentation · 20–25 slides
**Route committed:** graduation-project **Route A (modelling)** — PVT onset modelling + wellbore crossing-depth prediction + breakeven economics, with a declared tool hierarchy so no licence problem can stall the project.
**Reference case:** documented **deepwater Gulf of Mexico well** (declared as such on the slides).

> This package replaces the earlier casing-design package at your request. The casing package remains in `/home/user/Seminar_Casing_Design/` as a ready alternative if your supervisor prefers a drilling-oriented topic.

---

## Files

| File | What it is |
|---|---|
| `Asphaltene_Slides_23plus2.pptx` | The deck: 23 core + 2 labelled backup slides. Speaker notes, target times and cumulative clocks on every slide. Two original charts (stability envelope; onset-depth prediction) built from a consistent model. |
| `Asphaltene_Slides_23plus2.pdf` | PDF export — your projector-failure backup. |
| `00_Asphaltene_Master_Plan.docx` | Full plan: §0.1 lists every change from the original blueprint and why; timing architecture, checkpoints, cut list, 6-week plan, sources/tools, rehearsal ladder, risk register, formula & numbers cheat sheet, cue card. |
| `01_Asphaltene_Speaking_Script.docx` | Word-for-word script for all 23 slides (≈1,835 words ≈ 14:07 of speech), with DO / IF-BEHIND cues. |
| `02_Asphaltene_QandA_Pack.docx` | 25 questions with model answers (chemistry & mechanisms · field operations · thesis scope), including sharpened versions of the blueprint's two originals. |
| `03_Animation_and_Design_Guide.docx` | The click-build animation map (10 slides, ~15 min in PowerPoint), the design-system specification, and the 10-foot legibility check. |
| `_build/` | Scripts that generated everything (deck, figures, documents) — edit and re-run if you want changes. |

## Design-system pass (latest build)

Tokens, archetypes and type scale are now **locked** and enforced by shared builder functions — the "master layout" equivalent for a generated deck:

- **Colour tokens (exact hex):** NAVY `#1B365D` · TEAL `#0D7377` · ORANGE `#E07A3D` (graphics only) · AMB_D `#B35A18` (orange-family small text) · TEXT `#1A1A1A` · GRAY `#4A5568` · LIGHT `#F8FAFC` · RED `#B43B2E`. All text pairings meet **WCAG AA**; most are AAA (audit table in `05_Projector_Preflight.docx`).
- **Seven layout archetypes** (A-title → G-backup), each slide declaring one; identical header, footer and 0.62″–12.86″ content grid. **Automated margin audit: 25/25 slides pass.**
- **Type scale:** titles 22–24 pt, kickers 11.5 pt, body auto-lifted 1 pt so nothing sits below ~12.5 pt, captions floored at 9.5 pt (the projection limit).
- **Five line icons** built to one stroke weight: flask, flocculation clusters, well schematic, injection mandrel, breakeven chart.
- **Whitespace pass** on the dense slides (8, 12, 15, 17): slide 15 trimmed to four rows with the thermal point promoted to a readable callout; slide 17's callout enlarged.
- **Two defects found by the audit and fixed:** a duplicated overlapping text block on slide 15, and a money-icon/text collision on slide 17.
- **Figures re-exported at 300 dpi** with the new tokens and larger caption text.
- **PDF proof:** 25 pages, 16:9, in the package.

## Content-audit pass

Every number on the slides was traced to its source and, where possible, verified verbatim against the publisher text. Outcome:

- **One citation error corrected:** the deepwater GoM case is **Wylde, J.; Punase, A., JPT, 30 April 2020** (Clariant Oil Services) — not "JPT/SPE 2021" as the search metadata implied. Corrected on slides 4, 13, 16 and in the reference list.
- **One attribution conflict found and handled:** the $330–390 k/well/yr inhibitor figure is assigned to the **Gulf of Mexico** by Fluid Phase Equilibria (2020) and the Elsevier monograph (2021), but to the **Middle East** by Energy & Fuels (2021). The deck quotes the majority attribution and flags the discrepancy; Q26 has the defence answer. All ranges now carry basin + year.
- **One weak claim replaced:** the unsourced "AOP often 100–200 bar above Pb" became the properly-sourced UAOP wording (Fuel 2021) plus de Boer (1995) for the screening rule.
- **Journal correction:** the dead-oil study is *Fluid Phase Equilibria* 2020, 514, 112552 — not *Fuel*.
- **All `[VERIFY]` flags resolved** (final scan: clean). Standards now named exactly: ASTM D6560 / IP 143, D6703, D2007-class.
- **Reference slide rebuilt** in SPE style with 14 entries carrying DOIs and SPE paper numbers.
- **`04_Source_Log.docx` added** — every slide number, every claim, its source, the verbatim quote, and a `[V]`/`[L]` status (L = page-number confirmation at your library in Week 1). Print it; hand it to any examiner who challenges a figure.

## Design refresh (latest build)

- **Action titles everywhere** — every slide title states the takeaway (e.g. *"GoM case study: injection placement above the onset left 3,400 ft unprotected"*).
- **Four purpose-built figures**, all at 200 dpi and reusable in the thesis: the P–T envelope (safety-orange dashed AOP, red instability zone, arrowed flowing path with onset star), the depth-prediction chart (**depth increases downward**), a **to-scale engineering cross-section** of the tubing choking from 2.30″ to 1.35″ (replaces the AI rendering), and a three-panel **core-shell colloidal schematic**.
- **Colour-coded CII equation** — destabilising numerator in orange, natural-dispersant denominator in teal.
- **Design discipline:** white background, navy text (never pure black), orange reserved exclusively for "danger/onset/deposit", no chart gridlines, 12–14 pt body / 22–24 pt titles, section kicker on every slide.
- Animation builds are **not** baked in (generator limitation — explained honestly in `03_...`); apply them in PowerPoint following the map.

---

## What I fixed from the pasted blueprint (full list in §0.1 of the plan)

1. **Corrected the stability inversion.** Asphaltenes don't keep oil stable — they're *kept dispersed* by resins and aromatics. The slide now says this correctly; juries test it.
2. **Corrected the deposition location.** "Near the top where pressure drops fastest" is wrong — the gradient is roughly steady; deposits form where the flowing P–T path crosses the onset envelope, often deep. That correction became the **case study**: a documented GoM well with **3,400 ft of tubing below the chemical injection mandrel** (ID choked 2.3″ → 1.35″, fixed with solvent soak + continuous inhibitor 750→500 ppm, +16% production over 4 months).
3. **Fixed the timing.** The blueprint summed to exactly 15:00 with no buffer. This deck speaks to ≈14:25 with four checkpoints (7:05 · 8:00 · 10:40 · 14:15) and a written cut list.
4. **Separated conclusions from thanks** and added references (skippable) + two labelled backup slides, preserving the 23+2 = 25 compliance trick.
5. **Committed one graduation route with a declared tool hierarchy** — licensed PVTsim/Multiflash/OLGA if available → **NeqSim** (open-source, Equinor, Apache-2.0, CPA-EoS asphaltene screening) → correlations. Verified that NeqSim genuinely exists and has asphaltene capability.
6. **Added real numbers**: ≈$70M single-well GoM shut-in, up to $1.2M/day deferred production, inhibitor cost ranges ($31–46k Middle East vs $330–390k GoM per well/yr), Hassi Messaoud (83.4% asphaltene deposit; 20–25% WHP loss in 15–20 days). Every figure carries its source on slide 23, with a caveat that ranges span different fields and years.
7. **Expanded Q&A from 2 to 25 questions**, and sharpened the "choke the well" answer with the Hassi Messaoud counter-intuitive finding (producing at *low* wellhead pressure dramatically reduced cleanouts).

## The two charts (both in the deck)

- **`_build/figs/envelope_pt.png`** — stability envelope: AOP curve, bubble point, flowing path crossing at ≈1,014 bar / 133 °C.
- **`_build/figs/depth_prediction.png`** — flowing pressure vs AOP along the wellbore; predicted onset ≈ 4,752 m (~300 m *below* the existing mandrel), highlighted against the case's 3,400 ft unprotected interval. This chart is the visual proof of your thesis problem statement.
- `_build/figs/plugged_tubing.png` — **AI-generated rendering** of a plugged tubing cutaway. It is labelled on the slide as a rendering, not a field photo. Replace it with a real published photograph if you can obtain permission; use `[VERIFY]` discipline.

## Before your first rehearsal (mandatory)

1. Replace every `[bracketed]` placeholder (name, ID, course code, supervisor, month).
2. Resolve every `[VERIFY]` flag: exact CII paper, ASTM/D6560/D6703 editions, de Boer citation, cost sources, licence availability (Week 1!).
3. Confirm with your supervisor that a modelling-based thesis will be accepted, and which software the department owns.
4. Decide the fabricated-vs-real image question for slide 4.

## Recommended preparation order

Week 1 read: `00_..._Master_Plan.docx` §0–2 and the deck once, cover to cover, without speaking.
Week 2: cut the script to your voice; build the calculation sheet with the case geometry.
Week 3–4: rehearse and polish. The five lines to memorise verbatim are flagged in §5 of the plan.

## Figures pass (step 3) — publication-quality, vector-exported

All five figures are **generated by code** (`_build/make_figs.py`), never screenshotted, using the locked
tokens (navy #1B365D, teal #0D7377, orange #E07A3D, amber-text #B35A18, red #B43B2E). Each figure is
exported twice: **PNG at 300 dpi** for the deck and **SVG (vector)** for PowerPoint / the thesis.
Every figure carries its source line at the bottom (10–11 pt in figure coordinates).

| figure | slide | what it shows | source line |
|---|---|---|---|
| `envelope_pt` | 10 | P–T stability envelope: reservoir 145 °C / 1,150 bar, wellhead 60 °C / 150 bar, dashed AOP, bubble-point line, instability shading, exact crossing **1,014 bar @ 133 °C** | Wylde & Punase, JPT, 30 Apr 2020 |
| `depth_prediction` | 13 | flowing pressure P(d) vs AOP(d) with the twin temperature axis; **predicted onset depth 4,752 m**, existing mandrel 4,450 m, the 3,400 ft unprotected interval, **target injection ≈ 5,000 m** | ditto + illustrative P–T model |
| `pipe_views` | 4 | three panels: clean bore (ID 2.30 in) \| deposit on the wall (ID 1.35 in) drawn **to scale** \| gauge-ring deposit profile (⅔-radius maximum) | Wylde & Punase 2020; profile after SPE-994-PA |
| `core_shell` | 3 | the three colloidal states: stable shell → destabilised → flocculated | textbook summary, see `04_Source_Log.docx` |
| `well_schematic` | 16 | where the chemical must arrive: capillary string, mandrel 4,450 m, onset 4,752 m, unprotected interval, perfs 17,700–18,000 ft | Wylde & Punase 2020 + slide-13 model |

Notes
- The two charts are placed at **8.15 in wide (≈97 % of native size)**, so their labels project at ~9.5–10 pt
  effective — above the ~18 mm-on-screen floor for a lecture room. The well schematic is drawn on a compact
  canvas so its 10.5 pt labels land at ~10 pt when placed at 4.60 in.
- Superseded assets moved to `_build/figs/_superseded/` (`pipe_cross_section.png`, `plugged_tubing.png`).
- Deck now embeds **10 images** (5 figures + 5 icons); margin audit still PASS (content inside 0.62–12.86 in).
- The 300-dpi proof `Asphaltene_Slides_23plus2.pdf` was re-exported after the figure swap (25 pp, 960×540 pt).

## Narrative & density pass

Wording, transitions and projected density tightened; nothing removed from the five-part story.

**Structural changes**
- **Slide 6 (agenda)** now carries the required scope line: *“SCOPE — a flow-assurance review plus a thesis feasibility study; not a completed field intervention.”* (visible box, and spoken in the script).
- **Slide 9** — the `PRECIPITATION ≠ DEPOSITION` box is now the visual anchor: tinted fill (`#FBEAE5`), 2.6 pt red border, centred 17 pt heading, body clearly subordinate to the four step cards.
- **Slide 13** — bullets compressed and cross-referenced to the figure (“all depths labelled on the figure”); a speaker cue was added to the notes to **point at the ≈300 m bracket** (mandrel 4,450 m → onset 4,752 m) while speaking.
- **Slide 17** — the requested sentence is now the callout’s lead line: *“The thesis supplies the missing breakeven frequency for this specific well…”* The attribution caveat shrank to one italic line.
- **Slide 20** — the tool hierarchy is now three visible chips rather than a paragraph: **licensed (if available) → NeqSim open-source fallback (verified) + Python → correlation floor**, with a one-line note that the project cannot stall on software.
- **Slide 22** — notes now instruct rehearsing each numbered message as a standalone 15-second soundbite, practised out of order.

**Transitions added** (teal `NEXT →` lines, one per part except the closing slides): slides 8, 12, 13, 21.

**Trims** (wording only — no fact or citation was dropped): slides 2, 7, 12, 13, 14, 16, 17. Slide 6 rows tightened to make room for the scope box. The wax-vs-asphaltene table (slide 8), the misconception box (slide 12), the reactive-toolkit table (15) and the business-case table (17) were kept at full strength.

**Consistency re-verified after the pass**
- 25 slides · 25 notes · 10 images · margin audit **PASS 25/25**; flag scan CLEAN.
- Speaking script rebuilt: **1,838 spoken words ≈ 14.1 min** (unchanged budget); scope line added to slide 6, the ≈300 m pointing cue added, slides 16/17 wording synced to the deck, trim absorbed the addition so the word count stayed flat.
- 300-dpi proof `Asphaltene_Slides_23plus2.pdf` re-exported (25 pp, 960×540 pt).

## Technical accuracy & risk pass

New deliverable: **`06_Technical_Accuracy_and_Risk_Register.docx`** — Part A verification register (14 claims →
CONFIRMED / REFINED / CORRECTED with the source evidence), Part B change log, Part C 14-row risk register with
likelihood × impact and triggers, Part D the quantified sensitivity study, Part E the HSE/ethics position.

**Claims re-checked against originals (A1–A14).** Confirmed as written: CII formula and thresholds (+ the
laboratory-specific caveat); precipitation grows to a maximum near Pb with re-dissolution below it; ASTM D6560 /
IP 143 for asphaltene content; Yen–Mullins nano-aggregates (~750 g/mol, nanoaggregates < 10, clusters ~5 nm);
thermal methods are not a melting route (no melting point; a 497 °C steam case is openly questioned in the
literature); BTEX exposure as a workplace-safety topic.

**Three things were wrong or imprecise and are now fixed in the deck, script, Q&A and source log:**
1. **The de Boer plot was mis-described.** It is in-situ density vs gas under-saturation (P_res − Pb) with three
   risk zones — *not* an AOP-versus-Pb cross-plot. The onset-above-Pb rule is a separate envelope screen. Fixed on
   slide 10, in the script, in Q6 + new Q29, on backup slide 24, and flagged in the slide-10 speaker notes.
2. **The placement rule was stated backwards in effect.** "Deeper than the *shallowest* predicted onset" → corrected to
   "below the **deepest credible** onset depth — onset plus a margin sized by the sensitivity study."
3. **The thermal caption over-claimed.** "Thermal methods fail here by design" → "not a melting story … heat can even
   weaken the resins that keep them dispersed", with ref [15].

Also refined: inhibitor vs dispersant (onset-shift vs aggregate-size reduction — Kelland 2009, ref [16]); over-dosing
can *increase* deposition (dose window, not dose maximum — new Q31); "100–200 bar above Pb" replaced by "hundreds of
bar — thousands of psi in severe oils"; ASTM list corrected (D6560-22 / IP 143:21, D2007, D4124, D3279, and **D6703 =
Automated Heithaus Titrimetry**, not a general precipitation test).

**Backup slide 24 now carries the sensitivity study as arithmetic, not as a plan**: ±50 bar AOP → ±255 m; ±100 bar →
±511 m; ±10 % gradient → −404 / +487 m; gas-lift (gradient −10 %, AOP +50 bar) → 5,520 m, i.e. at/below the
perforations and the whole tubing unstable — placement alone cannot protect, so the dose does the work.

**Slide 21 limitations box** retitled "LIMITATIONS & ETHICS" and keeps all five technical limits plus the ethics line
(no uncontrolled BTEX handling, no vendor recommended). Q&A grew to **31 questions** (Q29 de Boer definition,
Q30 the steam counter-example, Q31 over-dosing). Speaking script re-balanced to **1,839 words ≈ 14.1 min**.
Proof re-exported; margin audit PASS; 25 slides / 25 notes / 10 images.
