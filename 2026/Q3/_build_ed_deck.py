"""
hūmānus — Q3 2026 ED Operations & Impact Update, designed edition.

Run from the repo root:
    HUMANUS_BRAND=/path/to/humanus-brand python3 2026/Q3/_build_ed_deck.py

Dependencies: python-pptx, Pillow

Content: Eva's own edit of the deck (CAP-525, 13 Oct 2026), transcribed
verbatim into the CONTENT dataclass at the bottom. No words added or removed;
this script only controls layout, type and colour. The previous Q2-cloning
builder lives in git history (commit 84e0986).

Spec
----
Canvas      16:9, 13.333 × 7.5 in. Margins 0.7 in left/right.
Palette     Core #410F37 (titles, body, dark slides), Bold #96FF14 (text on
            Core), Touch #E12D64 (section labels ≥14pt bold, bullet marks).
            White page. No tints, no extra colours.
Type        Fagun throughout, weight-named families (Fagun Black / Bold /
            Medium) plus base "Fagun" for Regular and its italic.
            Titles Bold 30pt; subtitles Regular Italic 14pt; labels Bold
            14pt all-caps Touch; body Regular 14–17pt, lead-ins Bold;
            captions Regular Italic 10pt. Links in Core, underlined.
            Theme fonts also set to Fagun so new text added in PowerPoint
            stays on-brand.
Fonts       Fagun is referenced, not embedded: PowerPoint cannot embed
            CFF-flavoured OTF. Machines without Fagun substitute; install
            Fagun from humanus-brand/fonts/ before presenting or exporting.
Logo        Bold (lime) master on Core cover and closing slides — never
            enlarged beyond native 5179 px; no other treatment.
Pattern     Lime tile from humanus_patterns.png, cropped at build time,
            as a full-height panel on cover and closing slides.
Contrast    Core on white 13.6:1; Bold on Core 5.4:1 (≥14pt here);
            Touch on white 4.4:1 — used only for 14pt bold labels (large-text
            AA) and decorative bullet marks.
"""

from __future__ import annotations

import io
import os
from dataclasses import dataclass, field
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt
from lxml import etree

REPO_ROOT = Path(__file__).resolve().parents[2]
HERE = Path(__file__).resolve().parent
ASSETS = HERE / "assets"
BRAND = Path(os.environ.get("HUMANUS_BRAND", "/humanus-brand"))
OUT = HERE / "ED-Operations-Impact-Q3-2026.pptx"

CORE = RGBColor(0x41, 0x0F, 0x37)
BOLD = RGBColor(0x96, 0xFF, 0x14)
TOUCH = RGBColor(0xE1, 0x2D, 0x64)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

F_BLACK = "Fagun Black"
F_BOLD = "Fagun Bold"
F_MEDIUM = "Fagun Medium"
F_BOOK = "Fagun"                # Regular: the base family resolves on every platform
F_BOOK_IT = "Fagun italic"      # sentinel → Fagun + italic flag (true Regular Italic)

W, H = 13.333, 7.5
ML = 0.7
CW = W - 2 * ML               # content width
TOP = 2.05                    # content top, every content slide
PHOTO_W = 4.2
PANEL_W = 4.0                 # pattern panel on cover / closing


# ---------------------------------------------------------------- helpers

def _check(path: Path) -> Path:
    if not path.exists():
        raise SystemExit(f"Brand asset not found: {path} (set HUMANUS_BRAND)")
    return path


def rect(slide, x, y, w, h, fill=None, line=None, line_w=1.25, name=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(line_w)
    sh.shadow.inherit = False
    style = sh._element.find(qn("p:style"))
    if style is not None:
        sh._element.remove(style)
    if name:
        sh.name = name
    return sh


def textbox(slide, x, y, w, h, anchor=MSO_ANCHOR.TOP, name=None):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    if name:
        tb.name = name
    return tf


def run(p, text, font, size, color, caps=False, link=None):
    italic = font == F_BOOK_IT
    font = F_BOOK if italic else font
    r = p.add_run()
    r.text = text
    r.font.name = font
    r.font.size = Pt(size)
    r.font.color.rgb = color
    r.font.bold = False
    r.font.italic = italic
    if caps:
        r.font._rPr.set("cap", "all")
    if link:
        r.hyperlink.address = link
    # East-Asian / complex-script slots too, so no glyph falls back to Calibri
    rPr = r.font._rPr
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set("typeface", font)
    return r


def para(tf, first=False, align=PP_ALIGN.LEFT, before=0, after=0, line=None):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.space_before = Pt(before)
    p.space_after = Pt(after)
    if line:
        p.line_spacing = line
    return p


def bullet(p, color=TOUCH, indent=0.24):
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(Inches(indent)))
    pPr.set("indent", str(-Inches(indent)))
    for tag in ("a:buClr", "a:buSzPct", "a:buFont", "a:buChar", "a:buNone"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    clr = etree.SubElement(pPr, qn("a:buClr"))
    etree.SubElement(clr, qn("a:srgbClr")).set("val", str(color))
    etree.SubElement(pPr, qn("a:buSzPct")).set("val", "80000")
    etree.SubElement(pPr, qn("a:buFont")).set("typeface", "Arial")
    etree.SubElement(pPr, qn("a:buChar")).set("char", "■")


def hang(p, indent=0.24):
    """Continuation paragraph aligned with bullet text, no mark."""
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(Inches(indent)))
    pPr.set("indent", "0")
    etree.SubElement(pPr, qn("a:buNone"))


def lead_runs(p, text, size, color=CORE, lead_font=F_BOLD, body_font=F_BOOK):
    """'Lead — rest' → lead in Bold, rest in Regular. Text is unchanged."""
    if " — " in text:
        lead, rest = text.split(" — ", 1)
        run(p, lead, lead_font, size, color)
        run(p, " — " + rest, body_font, size, color)
    else:
        run(p, text, body_font, size, color)


def label(tf, text, first=False, before=14, color=TOUCH):
    p = para(tf, first=first, before=0 if first else before, after=6)
    run(p, text, F_BOLD, 14, color, caps=True)
    return p


def add_picture_fit(slide, path, x, y, w):
    with Image.open(path) as im:
        pw, ph = im.size
    h = w * ph / pw
    slide.shapes.add_picture(str(path), Inches(x), Inches(y), Inches(w), Inches(h))
    return h


def logo(slide, variant, x_glyph, y, w):
    """Place a logo so its visible glyphs (not the PNG padding) start at x_glyph."""
    path = _check(BRAND / "logo" / f"humanus_logo_master_{variant}.png")
    with Image.open(path) as im:
        pw, ph = im.size
        l, t, r, b = im.getbbox()
    scale = w / pw
    slide.shapes.add_picture(str(path), Inches(x_glyph - l * scale), Inches(y - t * scale),
                             Inches(w), Inches(w * ph / pw))
    return (b - t) * scale            # visible glyph height


def pattern_panel(slide, x, w):
    """Lime tile (row 1, col 2) of the official pattern sheet, cropped to the panel."""
    path = _check(BRAND / "humanus_patterns.png")
    with Image.open(path) as im:
        tw, th = im.width // 3, im.height // 4
        crop_h = th - 140
        crop_w = int(crop_h * w / H)
        tile = im.crop((tw + 120, 40, tw + 120 + crop_w, 40 + crop_h))
        buf = io.BytesIO()
        tile.save(buf, "PNG")
    buf.seek(0)
    slide.shapes.add_picture(buf, Inches(x), 0, Inches(w), Inches(H))


def background(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def header(slide, title, subtitle):
    tf = textbox(slide, ML, 0.55, CW, 0.6, anchor=MSO_ANCHOR.BOTTOM, name="Title")
    run(para(tf, first=True), title, F_BOLD, 30, CORE)
    tf = textbox(slide, ML, 1.3, CW, 0.4, name="Subtitle")
    run(para(tf, first=True), subtitle, F_BOOK_IT, 14, CORE)


def footer(slide, text, n, total):
    tf = textbox(slide, ML, 7.0, 9.0, 0.25, name="Footer")
    run(para(tf, first=True), text, F_BOOK, 9, CORE)
    tf = textbox(slide, W - ML - 1.5, 7.0, 1.5, 0.25, name="Page")
    run(para(tf, first=True, align=PP_ALIGN.RIGHT), f"{n} / {total}", F_BOOK, 9, CORE)


def caption(slide, x, y, w, text):
    tf = textbox(slide, x, y, w, 0.3, name="Caption")
    run(para(tf, first=True), text, F_BOOK_IT, 10, CORE)


def bullets_block(tf, items, size=14, first=False):
    for i, item in enumerate(items):
        lines = item if isinstance(item, tuple) else (item,)
        p = para(tf, first=first and i == 0, after=0 if len(lines) > 1 else 7)
        bullet(p)
        lead_runs(p, lines[0], size)
        for extra in lines[1:]:
            p = para(tf, after=7)
            hang(p)
            text, url = extra
            run(p, text, F_BOOK, size, CORE, link=url)


def set_theme_fonts(prs):
    """Point the theme's major/minor Latin fonts at Fagun."""
    master = prs.slide_masters[0]
    for rel in master.part.rels.values():
        if rel.reltype.endswith("/theme"):
            part = rel.target_part
            root = etree.fromstring(part.blob)
            ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
            root.find(".//a:majorFont/a:latin", ns).set("typeface", F_BOLD)
            root.find(".//a:minorFont/a:latin", ns).set("typeface", F_BOOK)
            hl = root.find(".//a:hlink", ns)
            for ch in list(hl):
                hl.remove(ch)
            etree.SubElement(hl, qn("a:srgbClr")).set("val", str(CORE))
            part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


# ---------------------------------------------------------------- slides

def slide_cover(prs, c):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s, CORE)
    pattern_panel(s, W - PANEL_W, PANEL_W)
    x = 0.9
    gh = logo(s, "bold", x, 1.35, 6.4)
    y = 1.35 + gh + 0.55
    tf = textbox(s, x, y, W - PANEL_W - x - 0.6, 1.0)
    run(para(tf, first=True, after=4), c.cover_line, F_MEDIUM, 20, BOLD)
    run(para(tf), c.cover_date, F_BOOK, 16, BOLD)
    tf = textbox(s, x, 4.75, W - PANEL_W - x - 0.6, 1.0)
    run(para(tf, first=True, after=6), c.tagline_cover, F_BLACK, 30, TOUCH)
    run(para(tf), c.values, F_MEDIUM, 12, BOLD)
    tf = textbox(s, x, 6.75, W - PANEL_W - x - 0.4, 0.4)
    run(para(tf, first=True), c.sponsor_line, F_BOOK_IT, 10, BOLD)


def slide_goals(prs, c, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header(s, c.goals_title, c.goals_subtitle)
    gap = 0.3
    cw = (CW - 2 * gap) / 3
    for i, text in enumerate(c.goals):
        x = ML + i * (cw + gap)
        rect(s, x, TOP, cw, 3.65, fill=CORE, name=f"Goal {i + 1}")
        tf = textbox(s, x + 0.35, TOP + 0.3, cw - 0.7, 1.2)
        run(para(tf, first=True), str(i + 1), F_BLACK, 72, BOLD)
        tf = textbox(s, x + 0.35, TOP + 1.65, cw - 0.7, 1.8)
        run(para(tf, first=True, line=1.1), text, F_MEDIUM, 18, BOLD)
    tf = textbox(s, ML, TOP + 3.65 + 0.4, CW, 0.5)
    run(para(tf, first=True), c.goals_note, F_BOOK_IT, 14, CORE)
    footer(s, c.footer, n, total)


def _cell(cell, text, font, size, color, fill=None):
    cell.margin_left = cell.margin_right = Inches(0.2)
    cell.margin_top = cell.margin_bottom = Inches(0.08)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf = cell.text_frame
    tf.word_wrap = True
    run(para(tf, first=True), text, font, size, color)
    if fill is None:
        cell.fill.background()
    else:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill
    tcPr = cell._tc.get_or_add_tcPr()
    for side in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        ln = etree.SubElement(tcPr, qn(side))
        if side == "a:lnB":
            ln.set("w", str(Pt(1)))
            sf = etree.SubElement(ln, qn("a:solidFill"))
            etree.SubElement(sf, qn("a:srgbClr")).set("val", str(CORE))
        else:
            ln.set("w", "0")
            etree.SubElement(ln, qn("a:noFill"))
    # schema order: borders must precede the fill element
    for f in tcPr.findall(qn("a:solidFill")) + tcPr.findall(qn("a:noFill")):
        tcPr.remove(f)
        tcPr.append(f)


def slide_where(prs, c, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header(s, c.where_title, c.where_subtitle)
    rows = len(c.where_rows) + 1
    gf = s.shapes.add_table(rows, 2, Inches(ML), Inches(TOP), Inches(CW), Inches(0.6 + 1.05 * (rows - 1)))
    gf.name = "Programme table"
    tbl = gf.table
    tblPr = tbl._tbl.tblPr
    for attr in ("firstRow", "bandRow"):
        tblPr.set(attr, "0")
    style = tblPr.find(qn("a:tableStyleId"))
    if style is not None:
        tblPr.remove(style)
    tbl.columns[0].width = Inches(6.6)
    tbl.columns[1].width = Inches(CW - 6.6)
    tbl.rows[0].height = Inches(0.6)
    _cell(tbl.cell(0, 0), c.where_head[0], F_BOLD, 14, BOLD, fill=CORE)
    _cell(tbl.cell(0, 1), c.where_head[1], F_BOLD, 14, BOLD, fill=CORE)
    for r, (strand, funder) in enumerate(c.where_rows, start=1):
        tbl.rows[r].height = Inches(1.05)
        _cell(tbl.cell(r, 0), strand, F_MEDIUM, 20, CORE)
        _cell(tbl.cell(r, 1), funder, F_BOOK, 16, CORE)
    footer(s, c.footer, n, total)


def slide_text_photo(prs, c, n, total, title, subtitle, blocks, photo, cap, size=14, photo_w=PHOTO_W):
    """blocks: list of (label, paragraph-str | list-of-bullets)."""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header(s, title, subtitle)
    photo_x = W - ML - photo_w
    tf = textbox(s, ML, TOP, photo_x - ML - 0.45, 4.8, name="Body")
    for i, (lab, body) in enumerate(blocks):
        label(tf, lab, first=i == 0, before=size)
        if isinstance(body, str):
            run(para(tf, after=4, line=1.05), body, F_BOOK, size, CORE)
        else:
            bullets_block(tf, body, size=size)
    ph = add_picture_fit(s, ASSETS / photo, photo_x, TOP + 0.05, photo_w)
    caption(s, photo_x, TOP + 0.05 + ph + 0.1, photo_w, cap)
    footer(s, c.footer, n, total)


def slide_dialogue(prs, c, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header(s, c.dialogue_title, c.dialogue_subtitle)
    left_w = 5.5
    tf = textbox(s, ML, TOP, left_w, 4.6)
    label(tf, c.dialogue_label_1, first=True)
    run(para(tf, line=1.1), c.dialogue_body, F_BOOK, 16, CORE)
    px = ML + left_w + 0.45
    pw = W - ML - px
    rect(s, px, TOP, pw, 4.55, fill=CORE, name="Since 29 June panel")
    tf = textbox(s, px + 0.45, TOP + 0.4, pw - 0.9, 3.8)
    label(tf, c.dialogue_label_2, first=True, color=BOLD)
    p = para(tf, before=6, line=1.1)
    lead_runs(p, c.dialogue_since, 20, color=BOLD, lead_font=F_BOLD, body_font=F_MEDIUM)
    footer(s, c.footer, n, total)


def slide_business(prs, c, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header(s, c.biz_title, c.biz_subtitle)
    tf = textbox(s, ML, TOP, CW, 1.2)
    label(tf, c.biz_label_1, first=True)
    run(para(tf, line=1.05), c.biz_framework, F_BOOK, 16, CORE)
    tf = textbox(s, ML, TOP + 1.45, CW, 0.35)
    label(tf, c.biz_label_2, first=True)
    gap = 0.3
    cw = (CW - 2 * gap) / 3
    y = TOP + 1.95
    ch = 6.7 - y
    for i, (head, paras) in enumerate(c.biz_streams):
        x = ML + i * (cw + gap)
        rect(s, x, y, cw, ch, line=CORE, name=head)
        tf = textbox(s, x + 0.3, y + 0.28, cw - 0.6, ch - 0.5)
        run(para(tf, first=True, after=8), head, F_BOLD, 18, CORE)
        for t in paras:
            run(para(tf, after=6, line=1.05), t, F_BOOK, 14, CORE)
    footer(s, c.footer, n, total)


def slide_further(prs, c, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header(s, c.further_title, c.further_subtitle)
    tf = textbox(s, ML, TOP, CW, 0.35)
    label(tf, c.further_label, first=True)
    gap = 0.3
    cw = (CW - gap) / 2
    y = TOP + 0.55
    for i, text in enumerate(c.further_items):
        x = ML + i * (cw + gap)
        rect(s, x, y, cw, 3.4, fill=CORE, name=f"Activity {i + 1}")
        tf = textbox(s, x + 0.45, y + 0.45, cw - 0.9, 2.6)
        lead, rest = text.split(" — ", 1)
        p = para(tf, first=True, line=1.1)
        run(p, lead, F_BOLD, 20, BOLD)
        run(p, " — " + rest, F_MEDIUM, 18, BOLD)
    footer(s, c.footer, n, total)


def slide_closing(prs, c):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s, CORE)
    pattern_panel(s, 0, PANEL_W)
    x = PANEL_W + 0.9
    gh = logo(s, "bold", x, 1.9, 6.4)
    tf = textbox(s, x, 1.9 + gh + 0.5, W - x - ML, 1.2)
    run(para(tf, first=True, after=6), c.tagline_close, F_BLACK, 30, BOLD)
    run(para(tf), c.values, F_MEDIUM, 12, BOLD)
    tf = textbox(s, x, 5.2, W - x - ML, 0.4)
    run(para(tf, first=True), c.contact, F_MEDIUM, 16, BOLD)
    tf = textbox(s, x, 6.75, W - x - ML, 0.4)
    run(para(tf, first=True), c.closing_line, F_BOOK_IT, 10, BOLD)


# ---------------------------------------------------------------- content

@dataclass
class Content:
    footer: str = "hūmānus  ·  Q3 2026 Board Meeting  ·  Operations & Impact Update"

    cover_line: str = "update since 29 June  ·  Q3 2026 Board Meeting"
    cover_date: str = "29 October 2026"
    tagline_cover: str = "Humanising Justice"
    tagline_close: str = "Humanising Justice"
    values: str = "ACCOUNTABILITY  ·  COMMUNITY  ·  EVIDENCE"
    sponsor_line: str = ("Sponsor: Eva Buzo, Executive Director  ·  Agenda item TBC  ·  15 minutes  ·  "
                         "Discussion")

    goals_title: str = "Three goals guide everything we do"
    goals_subtitle: str = "Our 2024–29 strategy frames the year."
    goals: list = field(default_factory=lambda: [
        "Victims secure outcomes in line with their own conception of justice.",
        "Victims are empowered by the process of pursuing justice — not just by its results.",
        "Justice processes are transformed to put victims at the centre.",
    ])
    goals_note: str = ("The September retreat broadened our framing from serious international crimes "
                       "to human rights violations (slide 9).")

    where_title: str = "Where the work has sat since June"
    where_subtitle: str = "29 June – 12 October 2026, by funded programme and other significant strands."
    where_head: tuple = ("Programme / strand", "Funder")
    where_rows: list = field(default_factory=lambda: [
        ("Rohingya in India — legal protection", "Core funding"),
        ("Rohingya dialogue — Bangladesh and Arakan", "FCDO (UK) and Swiss FDFA, via Stimson Center"),
        ("Accountability for business and human rights", "Core funding"),
    ])

    india_title: str = "Rohingya in India — legal protection"
    india_subtitle: str = "Core funded  ·  no project grant  ·  Lead: Arati Ranade, India Programme Manager"
    india_blocks: list = field(default_factory=lambda: [
        ("What the strand is",
         "Direct legal intervention for detained Rohingya refugees in India, building a cross-state network "
         "of lawyers and an evidentiary base on detention and deportation patterns. Funded from core — no "
         "donor restriction on case selection or strategy."),
        ("Since 29 June", [
            "Delhi detention roundtable (5 July) — lawyers from across states convened over three panels to "
            "discuss detention litigation strategy.",
            "Right to Education — strategy note drafted on an education-rights route for detained minors, "
            "alongside the refugee route; raised with African states in Geneva (25 August).",
            "Litigation pipeline — habeas corpus petition in drafting; bail application being finalised for "
            "counsel's review; Right to Information petitioner being identified.",
            "Lawyer network — West Bengal, Uttar Pradesh and Delhi lawyers in place. Case-law report on "
            "Rohingya deportation in preparation.",
        ]),
    ])
    india_photo: str = "delhi-detention-roundtable.jpg"
    india_caption: str = "Delhi, July 2026"

    dialogue_title: str = "Rohingya dialogue — Bangladesh and Arakan"
    dialogue_subtitle: str = ("FCDO (UK) and Swiss FDFA, via Stimson Center as prime  ·  "
                              "hūmānus as implementing partner")
    dialogue_label_1: str = "What the project is"
    dialogue_body: str = ("Sustained dialogue and engagement with Rohingya leadership, building trust among "
                          "Rohingya, Arakanese and Bangladesh stakeholders through bilateral engagement, capacity "
                          "building and facilitated dialogue. FCDO and Swiss FDFA funding support the same "
                          "project; both reimburse in arrears.")
    dialogue_label_2: str = "Since 29 June"
    dialogue_since: str = ("Analysis — ten reports by Sawyadullah on Rohingya relationships with the United "
                           "States, China, Bangladesh, India and Pakistan, to aid the dialogue and support "
                           "Rohingya leadership. Website launching November 2026; a workshop on the content "
                           "will be run in the Rohingya language.")

    biz_title: str = "Accountability for business and human rights"
    biz_subtitle: str = ("Core funded  ·  legal action, regulatory complaint and advocacy — confirmed as a "
                         "core model at the retreat")
    biz_label_1: str = "THE FRAMEWORK"
    biz_framework: str = ("Bringing together our work that holds business actors accountable for contribution "
                          "to, or benefit from, serious international crimes — across litigation, regulatory "
                          "complaint, and investor-facing advocacy.")
    biz_label_2: str = "ACTIVE WORKSTREAMS"
    biz_streams: list = field(default_factory=lambda: [
        ("Meta / Facebook", [
            "Meta has responded to our case, including a request for particulars.",
            "Documentary narrative and video series in production with Amnesty International",
        ]),
        ("Adani", [
            "Coordination and investigation work continuing, including research on Australian super-fund "
            "exposure, fiduciary-duty case law, the OFAC civil probe and Adani Defence.",
        ]),
        ("Against the Tide Campaign", [
            "Supporting activists in Puerto Rico, Jamaica and Seychelles to fight development of luxury "
            "resorts in environmentally sensitive protected areas.",
        ]),
    ])

    public_title: str = "Public engagement"
    public_subtitle: str = "Core funded  ·  Against the Tide, speaking, advocacy and donor relations"
    public_blocks: list = field(default_factory=lambda: [
        ("Speaking and advocacy", [
            ("Berkeley (28 September) — keynote, Human Rights for Victims of Conflict in a Fragmenting World; "
             "session with Berkeley IIS clinic students on conflict monitoring.",
             ("Recording: youtu.be/_d9m1Zl0Pro", "https://youtu.be/_d9m1Zl0Pro?si=I3CIYQGk4J4UphfQ")),
            "Amnesty × hūmānus side event at UNGA (25 September).",
        ]),
        ("Donor relations", [
            "OSF partner retreat, Mexico — Finding Each Other on Shifting Ground, bringing together around "
            "100 OSF partners; good connections made.",
        ]),
    ])
    public_photo: str = "esencia-no-va-panel.jpg"
    public_caption: str = "Esencia NO VA panel"

    further_title: str = "Further activities"
    further_subtitle: str = "Advisory work outside our funded programmes"
    further_label: str = "Since 29 June"
    further_items: list = field(default_factory=lambda: [
        "Australian war crimes, victim participation — advice to two sets of clients in Australia on victim "
        "participation in prosecutions of alleged war crimes committed in Afghanistan.",
        "Caste-based discrimination — advice on a complaint to the Australian Human Rights Commission against "
        "a New South Wales local council, concerning discrimination against Dalit, Bahujan and Adivasi "
        "communities. Instructed September 2026.",
    ])

    retreat_title: str = "Strategy retreat — what we decided"
    retreat_subtitle: str = "Team retreat, September 2026  ·  strategy, impact and organisational focus"
    retreat_blocks: list = field(default_factory=lambda: [
        ("Strategy and impact", [
            "Plan from what clients need — an outcomes ladder linking litigation, mediation and community "
            "engagement to meaningful results.",
            "Broaden the framing from serious international crimes to human rights violations; stay "
            "people-focused and prioritise systemic change.",
            "Corporate accountability — legal action, regulatory complaints and advocacy — adopted as a core "
            "model, including Adani and Facebook.",
            "Rohingya — continue the lawyer-network and coordination approach, and take stock of progress "
            "after six months.",
        ]),
        ("Priorities", [
            "Keep excellence, rigour and clear accountability central; new geographies only where needs-driven.",
            "Tell the organisation's story well and show impact to clients and funders; consider dedicated "
            "fundraising capacity.",
        ]),
    ])
    retreat_photo: str = "team-retreat.jpg"
    retreat_caption: str = "Team retreat, September 2026"

    contact: str = "www.humanus.co  ·  hello@humanus.co"
    closing_line: str = "Q3 2026 Board Meeting  ·  Operations & Impact Update  ·  Sponsor: Eva Buzo (ED)"


def build(c: Content, out: Path = OUT) -> Path:
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)
    set_theme_fonts(prs)
    total = 10

    slide_cover(prs, c)
    slide_goals(prs, c, 2, total)
    slide_where(prs, c, 3, total)
    slide_text_photo(prs, c, 4, total, c.india_title, c.india_subtitle, c.india_blocks,
                     c.india_photo, c.india_caption)
    slide_dialogue(prs, c, 5, total)
    slide_business(prs, c, 6, total)
    slide_text_photo(prs, c, 7, total, c.public_title, c.public_subtitle, c.public_blocks,
                     c.public_photo, c.public_caption, size=17, photo_w=5.2)
    slide_further(prs, c, 8, total)
    slide_text_photo(prs, c, 9, total, c.retreat_title, c.retreat_subtitle, c.retreat_blocks,
                     c.retreat_photo, c.retreat_caption)
    slide_closing(prs, c)

    assert len(prs.slides) == total
    prs.core_properties.title = "ED Operations & Impact Update — Q3 2026 Board Meeting"
    prs.core_properties.author = "hūmānus"
    prs.save(out)
    return out


if __name__ == "__main__":
    print(build(Content()))
