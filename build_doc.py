#!/usr/bin/env python3
"""Build the methodology reference: one source, two outputs.

    python3 build_doc.py            # writes docs/DCR_methodology.pdf and
                                    # injects the print view into the page

The document explains every field and every formula. It is authored once, here,
and rendered twice -- to PDF for circulation, and to an HTML section the page
shows only when printing. Two hand-maintained copies would drift, and a
methodology note that disagrees with the tool is worse than none.

Needs reportlab for the PDF; the HTML injection works without it.
"""
from __future__ import annotations

import os
import re
import sys
from datetime import date

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "dcr-calculator.src.html")
OUT_PDF = os.path.join(HERE, "docs", "DCR_methodology.pdf")
BEGIN, END = "<!-- DOC:BEGIN -->", "<!-- DOC:END -->"

TITLE = "Development Control Regulations — Plot Yield Calculator"
SUBTITLE = "Fields, formulas and calculation methodology"

# ---------------------------------------------------------------------------
# Content. ("h1"|"h2", text) | ("p", text) | ("formula", [lines])
#          | ("table", [rows], [widths]) | ("bullets", [items])
# Inline markup: <b>bold</b>, <i>italic</i>. Keep it to those two.
# ---------------------------------------------------------------------------
DOC = [
 ("p", "This note describes what every field in the calculator means, where its value comes "
       "from, and how the derived figures are computed. Formulas are given exactly as the tool "
       "applies them. Where a value is an assumption rather than a published figure, it is "
       "marked as such — see <b>Assumptions and open items</b> at the end."),

 ("h1", "1. How a calculation flows"),
 ("p", "Four figures follow from the plot area and its land-use designation. Each is a single "
       "multiplication or division; the sequence is what matters."),
 ("formula", [
   "plot area          entered, or taken from the plot register",
   "FAR                from the designation, overridable",
   "coverage %         from the designation, overridable",
   "",
   "Max GFA          = plot area x FAR",
   "Plot coverage    = plot area x coverage %",
   "GLA              = Max GFA x GLA share",
   "Max activities   = floor(GLA / unit area)",
 ]),
 ("p", "<b>Plot coverage is a footprint, not a share of the total GFA.</b> It is the floor area "
       "the Code permits on the ground floor, expressed as a proportion of the plot. What is "
       "left of the plot outside it is the open ground, and that is what decides how much "
       "parking can sit on grade."),
 ("p", "<b>Max activities is a count, not an area.</b> The division is floored, because a "
       "partial tenancy is not a tenancy. It is a ceiling on how many separate activities the "
       "leasable area supports at the minimum unit size — not a target."),

 ("h1", "2. Input fields"),
 ("table", [
   ["Field", "Where the value comes from", "How it is used"],
   ["Plot",
    "The plot register: 66 plots, each with its sector/plot address, district, DevCode and "
    "area. Selecting one fills the plot area.",
    "Convenience only. Typing an area by hand returns this to manual entry, so a figure is "
    "never filed against the wrong plot."],
   ["Plot size",
    "Entered, or filled from the register. In m<sup>2</sup>.",
    "The driver of every other figure."],
   ["Land use / designation",
    "The Code schedule: 59 designations, each publishing a FAR and a Max Plot Coverage.",
    "Supplies FAR and coverage. 15 designations publish no FAR and 12 no coverage; those "
    "need an override."],
   ["FAR override",
    "Typed. A ratio, not a percentage.",
    "Replaces the designation's FAR. Used when the Code governs a plot by note reference "
    "rather than a published figure."],
   ["Max plot coverage",
    "The designation, or typed. 0–100 with at most two decimals.",
    "Sets the ground-floor footprint, and with it the open ground left for parking."],
   ["GLA share",
    "Defaults to 75%. <b>Not a Code citation</b> — the worked example's value.",
    "Converts Max GFA to leasable area."],
   ["ITC location",
    "Abu Dhabi or Al Ain, CBD or not.",
    "Chooses which variant of an ITC class applies, falling back progressively when a class "
    "is not split that finely."],
  ], [92, 200, 200]),

 ("h1", "3. Parameters"),
 ("p", "Each of these has a published or standard value and may be overridden. An override "
       "shows the original beside the field and annotates every figure it changed, so a "
       "number is never quietly different from the Code."),
 ("table", [
   ["Parameter", "Default", "Meaning"],
   ["Unit area per activity", "60 m<sup>2</sup>",
    "The minimum tenancy size. Divides GLA to give the activity ceiling."],
   ["Area per parking space", "32.5 m<sup>2</sup>",
    "From the standard. Includes the manoeuvring share, which is why it exceeds a bay's own "
    "footprint."],
   ["Usable per basement floor", "75%",
    "The share of a basement floor that becomes parking, after ramps, cores and plant."],
   ["Basement floor area", "the plot area",
    "A basement may use the whole plot; unlike a storey above ground it is not held to the "
    "coverage limit. That is why this defaults to the plot and not to the footprint."],
   ["Required spaces", "from the schedule",
    "Override when UPPC or ITC has granted a reduction. Replaces the schedule's total."],
  ], [110, 78, 304]),

 ("h1", "4. The activity schedule"),
 ("p", "Parking is not a property of the plot; it follows from what occupies it. Each row is "
       "one activity, and contributes its own parking demand."),
 ("table", [
   ["Column", "Meaning"],
   ["Activity (DED)",
    "One of 3,892 DED economic activities, with its ISIC class."],
   ["Slots",
    "How many tenancies of this activity. Counts against the activity ceiling."],
   ["Unit area",
    "The floor area one slot occupies. Inherited from the parameter, or entered per row — "
    "shown as <i>inherited</i> or <i>entered</i>."],
   ["ITC category",
    "The ITC class whose rate applies. Mapped automatically from the activity, with a "
    "confidence of high, medium or low, and reassignable by hand."],
   ["Rate / units",
    "The rate for the resolved class. Classes charged per unit rather than by area show a "
    "quantity box here instead, with the per-unit rate beneath it."],
   ["Bays",
    "This row's exact contribution, unrounded."],
  ], [92, 400]),
 ("p", "<b>Rounding happens once, on the total.</b> Rounding each row up and then adding would "
       "charge for part spaces the scheme never needs, and the error compounds with the number "
       "of rows:"),
 ("formula", [
   "2.372 + 5.536 + 5.99  =  13.898  ->  14 spaces      correct",
   "ceil(2.372) + ceil(5.536) + ceil(5.99)  =  15       wrong",
 ]),

 ("h1", "5. ITC rates"),
 ("p", "The matrix holds 141 rate codes across 70 classes, each published for a weekday and a "
       "weekend. The tool uses a single blended rate on a standard week:"),
 ("formula", [
   "combined = weekday x (5/7) + weekend x (2/7)",
   "         = weekday x 71.43% + weekend x 28.57%",
   "",
   "rate_total = employee/resident + visitor + truck/bus",
 ]),
 ("p", "<b>A rate is not always charged against floor area.</b> Each class names its own driver, "
       "and only 63 of the 141 codes are charged against an area at all:"),
 ("table", [
   ["Driver", "Codes", "Quantity the rate is charged against"],
   ["count", "73", "A number the user supplies: units, bedrooms, seats, students, beds, "
                   "doctors, berths, taxi bays, invitees, fuelling positions."],
   ["gfa", "62", "Floor area, per 100 or per 1,000 m<sup>2</sup> of GFA."],
   ["site", "5", "Site area, per 100 or per 1,000 m<sup>2</sup>."],
   ["gla", "1", "Leasable area, per 100 m<sup>2</sup>."],
  ], [70, 45, 377]),
 ("formula", [
   "area-driven   bays = slots x unit area x conversion x rate",
   "count-driven  bays = quantity x conversion x rate",
 ]),
 ("p", "A count-driven row with no quantity entered contributes nothing and raises a flag, "
       "rather than being silently charged by floor area."),

 ("h1", "6. Parking: open ground first, then basement"),
 ("p", "A space needs its own area, so the ground left outside the footprint holds only so "
       "many. The remainder goes to basement or podium, where a floor is only partly usable."),
 ("formula", [
   "spaces required        = ceil( sum of the rows' exact bays )",
   "total parking area     = spaces required x 32.5",
   "open ground area       = plot area - plot coverage",
   "spaces on ground       = min( required, floor(open ground / 32.5) )",
   "offset spaces          = required - spaces on ground",
   "offset parking area    = total parking area - on-ground area",
   "basement floors        = ceil( offset area / (basement floor area x 75%) )",
 ]),
 ("p", "<b>Community Retail is the case that forces basement parking, and it follows from the "
       "Code rather than a special rule.</b> CR is published at 100% plot coverage, so the "
       "footprint is the whole plot, there is no open ground, and every space goes below."),
 ("table", [
   ["", "general plot", "Community Retail"],
   ["Spaces required", "30", "120"],
   ["Total parking area", "975 m<sup>2</sup>", "3,900 m<sup>2</sup>"],
   ["Plot / coverage", "1,000 m<sup>2</sup> / 60%", "1,000 m<sup>2</sup> / <b>100%</b>"],
   ["Open ground", "400 m<sup>2</sup>, holds 12", "<b>none</b>"],
   ["On ground", "12", "0"],
   ["In basement/podium", "18, 585 m<sup>2</sup>", "120, 3,900 m<sup>2</sup>"],
   ["Floors at 75% usable", "1", "<b>6</b>"],
  ], [150, 171, 171]),

 ("h1", "7. Limitations L-3 and L-4"),
 ("p", "The Regulations cap parking demand at two spaces per 100 m<sup>2</sup> of GFA. The "
       "ceiling that puts on a plot is exact:"),
 ("formula", [
   "ceiling = Max GFA x 2 / 100        spaces",
 ]),
 ("bullets", [
   "<b>Below 1,600 m<sup>2</sup></b> — limitation <b>L-4</b> bars a use whose demand "
   "exceeds the ceiling.",
   "<b>At or above 1,600 m<sup>2</sup></b> — <b>L-3</b> permits it, but the excess is the "
   "developer's responsibility and earns no additional GFA.",
 ]),

 ("h1", "8. Restrictions the tool raises"),
 ("table", [
   ["Flag", "Condition", "Why it matters"],
   ["SLOTS", "slots used &gt; activities permitted",
    "The schedule holds more tenancies than the GLA supports."],
   ["GFA", "scheduled floor area &gt; Max GFA",
    "A restriction. The FAR does not permit the floor area, whatever the slot count says."],
   ["GLA", "scheduled floor area &gt; GLA, within GFA",
    "A note, not a restriction. Tenancies are let from leasable area, so this is the binding "
    "limit in practice — but GLA is itself an assumption, so it is not treated as a cap."],
   ["QTY", "a count-driven activity has no quantity",
    "Its parking cannot be counted until the quantity is given."],
   ["L-4 / L-3", "demand exceeds 2 spaces per 100 m<sup>2</sup> GFA",
    "Barred below 1,600 m<sup>2</sup>; above it, the excess is the developer's."],
  ], [66, 156, 270]),
 ("p", "The slot count stops standing in for floor area as soon as unit areas vary: six slots "
       "at a hand-entered 260 m<sup>2</sup> can sit inside the permitted count and still "
       "overrun the plot. The GFA and GLA checks exist because the slot check cannot catch that."),

 ("h1", "9. Assumptions and open items"),
 ("p", "Treat the derived figures as working numbers, not cleared outputs."),
 ("bullets", [
   "<b>The GLA share of 75% is uncited.</b> It is the worked example's value, not a Code "
   "figure. It scales GLA, the activity ceiling, and the GLA restriction.",
   "<b>The DED to ITC crosswalk is heuristic</b> — authored here, not supplied by the "
   "client: 162 activities mapped with high confidence, 2,722 medium, 1,008 low. Every row "
   "can be reassigned by hand, and the confidence is shown on each.",
   "<b>15 designations publish no FAR and 12 no coverage.</b> The Code governs them by note "
   "reference, and those notes are not in the data. They need an override.",
   "<b>An issued DCR report gave AVG. REQUIRED PARKING = 0</b> for a small commercial plot "
   "where the ITC rate implies 3 spaces. Not a rounding artefact, and unresolved. It should "
   "be raised with whoever owns the DCR generation system before any parking figure is "
   "quoted from that source.",
   "<b>Area per space (32.5 m<sup>2</sup>) and floor efficiency (75%) come from the "
   "standard.</b> They replaced earlier stand-ins that had no source behind them.",
 ]),
]

# ---------------------------------------------------------------------------
# HTML rendering: the section the page prints.
# ---------------------------------------------------------------------------
def esc(t):
    """Escape for HTML, keeping the handful of inline tags used above."""
    t = t.replace("&", "&amp;")
    for tag in ("b", "i", "sup"):
        t = t.replace(f"&lt;{tag}&gt;", f"<{tag}>").replace(f"&lt;/{tag}&gt;", f"</{tag}>")
    return t


def to_html():
    out = ['<section id="methodology" aria-hidden="true">',
           '  <header class="doc-hd">',
           f'    <h1>{TITLE}</h1>',
           f'    <p class="doc-sub">{SUBTITLE}</p>',
           f'    <p class="doc-meta">Generated {date.today().isoformat()}'
           ' &#x00B7; regenerate with <code>python3 build_doc.py</code></p>',
           '  </header>']
    for kind, *rest in DOC:
        if kind in ("h1", "h2"):
            out.append(f'  <{kind}>{esc(rest[0])}</{kind}>')
        elif kind == "p":
            out.append(f'  <p>{esc(rest[0])}</p>')
        elif kind == "formula":
            body = "\n".join(esc(l) for l in rest[0])
            out.append(f'  <pre class="doc-formula">{body}</pre>')
        elif kind == "bullets":
            out.append('  <ul>')
            out += [f'    <li>{esc(i)}</li>' for i in rest[0]]
            out.append('  </ul>')
        elif kind == "table":
            rows = rest[0]
            out.append('  <table class="doc-table"><thead><tr>')
            out += [f'    <th>{esc(c)}</th>' for c in rows[0]]
            out.append('  </tr></thead><tbody>')
            for r in rows[1:]:
                out.append('    <tr>' + "".join(f'<td>{esc(c)}</td>' for c in r) + '</tr>')
            out.append('  </tbody></table>')
    out.append('</section>')
    return "\n".join(out)


def inject():
    if not os.path.exists(SRC):
        sys.exit(f"{SRC} not found")
    s = open(SRC, encoding="utf-8").read()
    if BEGIN not in s or END not in s:
        sys.exit(f"markers {BEGIN} / {END} not found in the source")
    head, rest = s.split(BEGIN, 1)
    _, tail = rest.split(END, 1)
    open(SRC, "w", encoding="utf-8").write(head + BEGIN + "\n" + to_html() + "\n" + END + tail)
    return len(to_html().splitlines())


# ---------------------------------------------------------------------------
# PDF rendering.
# ---------------------------------------------------------------------------
def to_pdf():
    try:
        from reportlab.lib import colors
        from reportlab.lib.enums import TA_LEFT
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
        from reportlab.lib.units import mm
        from reportlab.platypus import (BaseDocTemplate, Frame, PageTemplate,
                                        Paragraph, Preformatted, Spacer, Table,
                                        TableStyle, ListFlowable, ListItem)
    except ImportError:
        print("  reportlab not installed; skipped the PDF")
        return None

    INK, INK2, RULE, ACCENT = (colors.HexColor("#1C1C1C"), colors.HexColor("#5C5C61"),
                               colors.HexColor("#C9C9CD"), colors.HexColor("#2C6098"))
    ss = getSampleStyleSheet()
    body = ParagraphStyle("body", parent=ss["BodyText"], fontName="Helvetica",
                          fontSize=9.2, leading=13.4, textColor=INK, spaceAfter=7,
                          alignment=TA_LEFT)
    h1 = ParagraphStyle("h1", parent=body, fontName="Helvetica-Bold", fontSize=11.5,
                        leading=15, spaceBefore=15, spaceAfter=6, textColor=INK)
    mono = ParagraphStyle("mono", parent=body, fontName="Courier", fontSize=8.2,
                          leading=11.6, textColor=INK)
    cell = ParagraphStyle("cell", parent=body, fontSize=8.4, leading=11.4, spaceAfter=0)
    cellh = ParagraphStyle("cellh", parent=cell, fontName="Helvetica-Bold", textColor=INK)

    # reportlab's built-in fonts have no <sup> glyphs for m2, but the <super> tag
    # shifts a normal digit, which does.
    def rl(t):
        return (t.replace("<sup>", "<super>").replace("</sup>", "</super>")
                 .replace("&gt;", ">").replace("&lt;", "<"))

    story = [
        Paragraph(TITLE, ParagraphStyle("t", parent=h1, fontSize=17, leading=21,
                                        spaceBefore=0, spaceAfter=3)),
        Paragraph(SUBTITLE, ParagraphStyle("st", parent=body, fontSize=10.5,
                                           textColor=ACCENT, spaceAfter=2)),
        Paragraph(f"Generated {date.today().isoformat()}",
                  ParagraphStyle("m", parent=body, fontSize=8, textColor=INK2,
                                 spaceAfter=13)),
    ]
    for kind, *rest in DOC:
        if kind in ("h1", "h2"):
            story.append(Paragraph(rl(rest[0]), h1))
        elif kind == "p":
            story.append(Paragraph(rl(rest[0]), body))
        elif kind == "formula":
            story.append(Spacer(1, 2))
            story.append(Preformatted("\n".join(rest[0]), mono))
            story.append(Spacer(1, 8))
        elif kind == "bullets":
            story.append(ListFlowable(
                [ListItem(Paragraph(rl(i), body), leftIndent=12) for i in rest[0]],
                bulletType="bullet", bulletFontSize=6, leftIndent=11, bulletOffsetY=-1))
            story.append(Spacer(1, 4))
        elif kind == "table":
            rows, widths = rest[0], rest[1]
            data = [[Paragraph(rl(c), cellh) for c in rows[0]]]
            data += [[Paragraph(rl(c), cell) for c in r] for r in rows[1:]]
            t = Table(data, colWidths=[w for w in widths], repeatRows=1, hAlign="LEFT")
            t.setStyle(TableStyle([
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LINEBELOW", (0, 0), (-1, 0), 0.9, INK),
                ("LINEBELOW", (0, 1), (-1, -2), 0.35, RULE),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 9),
            ]))
            story += [Spacer(1, 2), t, Spacer(1, 10)]

    os.makedirs(os.path.dirname(OUT_PDF), exist_ok=True)

    def furniture(canv, doc):
        canv.saveState()
        canv.setFont("Helvetica", 7.4)
        canv.setFillColor(INK2)
        canv.drawString(20 * mm, 12 * mm, TITLE)
        canv.drawRightString(A4[0] - 20 * mm, 12 * mm, f"Page {canv.getPageNumber()}")
        canv.setStrokeColor(RULE)
        canv.setLineWidth(0.4)
        canv.line(20 * mm, 15.5 * mm, A4[0] - 20 * mm, 15.5 * mm)
        canv.restoreState()

    doc = BaseDocTemplate(OUT_PDF, pagesize=A4,
                          leftMargin=20 * mm, rightMargin=20 * mm,
                          topMargin=18 * mm, bottomMargin=20 * mm,
                          title="DCR Plot Yield — Methodology",
                          author="Origen", subject=SUBTITLE)
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="f")
    doc.addPageTemplates([PageTemplate(id="p", frames=[frame], onPage=furniture)])
    doc.build(story)
    return OUT_PDF


def main():
    n = inject()
    print(f"  print view injected into the page ({n} lines)")
    p = to_pdf()
    if p:
        print(f"  {p}  ({os.path.getsize(p):,} bytes)")


if __name__ == "__main__":
    main()
