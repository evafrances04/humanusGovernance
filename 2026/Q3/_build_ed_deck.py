"""
Build the Q3 2026 ED Operations & Impact Update (DRAFT) from the Q2 2026 deck.

Run from the repo root:
    python3 2026/Q3/_build_ed_deck.py

Template: 2026/Q2/08-ED-Operations-Impact-Y2D.pptx on branch
agent/el/q2-2026-deck-polish (read via `git show`, so no copy is committed).
Every slide is cloned from a Q2 slide so palette, typography, logo and footer
match the board's last ED paper exactly. Content is laid out top-down by a
small flow helper; `[Eva to confirm ...]` markers are set in Touch pink so the
ED can find every open item at sign-off.

Sources: MA's strand-mapped pre-draft (12 Oct 2026, green-lit by Eva), Eva's
retreat summary, revisions, further activities and photos (CAP-518 comments, 12 Oct 2026), Stacy McGill's Q3 2026 Board
Finance Report (branch agent/stacy-mcgill/01a12f27).
"""

from __future__ import annotations

import copy
import io
import math
import re
import subprocess
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor

REPO_ROOT = Path(__file__).resolve().parents[2]
TEMPLATE_REF = "agent/el/q2-2026-deck-polish:2026/Q2/08-ED-Operations-Impact-Y2D.pptx"
OUT = REPO_ROOT / "2026" / "Q3" / "ED-Operations-Impact-Q3-2026-DRAFT.pptx"

TOUCH = RGBColor(0xE1, 0x2D, 0x64)
MARK = re.compile(r"(\[Eva to confirm[^\]]*\])")

FOOTER = "hūmānus  ·  Q3 2026 Board Meeting  ·  Operations & Impact Update  ·  DRAFT"
CONTENT_TOP = 1874519
CONTENT_BOTTOM = 6400800
EMU_PER_PT = 12700
ASSETS = REPO_ROOT / "2026" / "Q3" / "assets"
BERKELEY_URL = "https://youtu.be/_d9m1Zl0Pro?si=I3CIYQGk4J4UphfQ"
PHOTO_RIGHT = 11612880
PHOTO_W = 3931920          # 4.3 in, single photo beside text
PHOTO_W_SMALL = 3566160    # 3.9 in, two stacked photos
TEXT_W_PHOTO = PHOTO_RIGHT - PHOTO_W - 228600 - 640080
TEXT_W_TWO_PHOTOS = PHOTO_RIGHT - PHOTO_W_SMALL - 228600 - 640080


# ---------------------------------------------------------------- helpers

def load_template() -> Presentation:
    blob = subprocess.run(
        ["git", "show", TEMPLATE_REF], cwd=REPO_ROOT, capture_output=True, check=True
    ).stdout
    return Presentation(io.BytesIO(blob))


def shape(slide, name):
    for sh in slide.shapes:
        if sh.name == name:
            return sh
    raise KeyError(name)


def remove(slide, *names):
    for n in names:
        el = shape(slide, n)._element
        el.getparent().remove(el)


def set_text(sh, text):
    """Replace text, keeping the first run's formatting; colour markers Touch."""
    tf = sh.text_frame
    paras = tf.paragraphs
    p0 = paras[0]._p
    for p in paras[1:]:
        p._p.getparent().remove(p._p)
    proto_r = copy.deepcopy(p0.r_lst[0]) if p0.r_lst else None
    for r in list(p0.r_lst):
        p0.remove(r)
    for br in p0.findall("{http://schemas.openxmlformats.org/drawingml/2006/main}br"):
        p0.remove(br)
    template_p = copy.deepcopy(p0)
    for i, line in enumerate(text.split("\n")):
        p = p0 if i == 0 else copy.deepcopy(template_p)
        if i:
            tf._txBody.append(p)
        end = p.find("{http://schemas.openxmlformats.org/drawingml/2006/main}endParaRPr")
        for seg in filter(None, MARK.split(line)):
            r = copy.deepcopy(proto_r)
            r.find("{http://schemas.openxmlformats.org/drawingml/2006/main}t").text = seg
            if end is not None:
                end.addprevious(r)
            else:
                p.append(r)
        if MARK.search(line):
            for run in sh.text_frame.paragraphs[i].runs:
                if MARK.fullmatch(run.text):
                    run.font.color.rgb = TOUCH


def clone_slide(prs, src):
    new = prs.slides.add_slide(src.slide_layout)
    for ph in list(new.shapes):
        ph._element.getparent().remove(ph._element)
    rid_map = {}
    for rid, rel in src.part.rels.items():
        if rel.reltype.endswith("/image"):
            rid_map[rid] = new.part.relate_to(rel._target, rel.reltype)
    for sh in src.shapes:
        el = copy.deepcopy(sh._element)
        for node in el.iter():
            for attr, val in list(node.attrib.items()):
                if attr.endswith("}embed") and val in rid_map:
                    node.set(attr, rid_map[val])
        new.shapes._spTree.append(el)
    return new


def add_clone(slide, proto_shape, left, top, width, height, text):
    el = copy.deepcopy(proto_shape._element)
    slide.shapes._spTree.append(el)
    sh = slide.shapes[-1]
    sh.left, sh.top, sh.width, sh.height = left, top, width, height
    if text is not None:
        set_text(sh, text)
    return sh


def est_height(text, width_emu, size_pt, char_em=0.55, leading=1.4):
    width_pt = width_emu / EMU_PER_PT
    per_line = max(1, int(width_pt / (char_em * size_pt)))
    lines = sum(max(1, math.ceil(len(MARK.sub(lambda m: m.group(0), ln)) / per_line))
                for ln in text.split("\n"))
    return int(lines * size_pt * leading * EMU_PER_PT)


class Flow:
    """Top-down placement of eyebrows, body paragraphs and bullets."""

    def __init__(self, slide, protos, left=640080, width=10972800, top=CONTENT_TOP,
                 body_pt=12, bullet_pt=13):
        self.s, self.p = slide, protos
        self.left, self.width, self.y = left, width, top
        self.body_pt, self.bullet_pt = body_pt, bullet_pt

    def eyebrow(self, text):
        add_clone(self.s, self.p["eyebrow"], self.left, self.y, self.width, 457200, text)
        self.y += 411480

    def body(self, text, gap=160000):
        h = est_height(text, self.width, self.body_pt)
        add_clone(self.s, self.p["body"], self.left, self.y, self.width, h, text)
        self.y += h + gap

    def bullet(self, text, gap=150000):
        tw = self.width - 274320
        h = est_height(text, tw, self.bullet_pt)
        add_clone(self.s, self.p["marker"], self.left, self.y + 118872, 109728, 109728, None)
        sh = add_clone(self.s, self.p["bullet"], self.left + 274320, self.y, tw, h, text)
        self.y += h + gap
        return sh

    def space(self, emu=120000):
        self.y += emu

    def check(self, label):
        if self.y > CONTENT_BOTTOM:
            print(f"WARNING: {label} overflows by {(self.y - CONTENT_BOTTOM) / 914400:.2f} in")


def photo(slide, protos, path, top, caption, width=PHOTO_W, height=None, crop=None):
    """Right-aligned photo with a small caption; returns the y below the caption."""
    left = PHOTO_RIGHT - width
    pic = slide.shapes.add_picture(str(path), left, top, width=width)
    if crop:
        pic.crop_top, pic.crop_bottom = crop
    if height:
        pic.height = height
    y = top + pic.height + 45720
    add_clone(slide, protos["caption"], left, y, width, 228600, caption)
    return y + 228600 + 91440


def link(sh, needle, url):
    for p in sh.text_frame.paragraphs:
        for r in p.runs:
            if needle in r.text:
                r.hyperlink.address = url
                r.font.color.rgb = TOUCH


def frame(prs, tpl, title, subtitle):
    """Clone the Rohingya-in-India slide down to title, accent, subtitle, footer."""
    s = clone_slide(prs, tpl)
    keep = {"Rectangle 1", "TextBox 2", "Rectangle 3", "TextBox 4", "TextBox 18", "TextBox 19"}
    for sh in list(s.shapes):
        if sh.name not in keep:
            sh._element.getparent().remove(sh._element)
    set_text(shape(s, "TextBox 2"), title)
    set_text(shape(s, "TextBox 4"), subtitle)
    return s


# ---------------------------------------------------------------- content

def build():
    prs = load_template()
    q2 = list(prs.slides)
    n_template = len(q2)
    cover, goals, table, arakan, fcdo, india, wd, frontline, bhr, outlook, close = q2

    protos = {
        "eyebrow": shape(india, "TextBox 5"),
        "body": shape(india, "TextBox 6"),
        "marker": shape(india, "Rectangle 8"),
        "bullet": shape(india, "TextBox 9"),
        "qhead": shape(bhr, "TextBox 8"),
        "qbody": shape(bhr, "TextBox 9"),
        "bar": shape(outlook, "Rectangle 5"),
        "whead": shape(outlook, "TextBox 6"),
        "wbody": shape(outlook, "TextBox 7"),
        "caption": shape(india, "TextBox 18"),
    }
    footers = []  # (footer shape, page shape)

    # 1 — cover
    s = clone_slide(prs, cover)
    set_text(shape(s, "TextBox 4"), "update since 29 June  ·  Q3 2026 Board Meeting")
    set_text(shape(s, "TextBox 5"), "29 October 2026")
    set_text(shape(s, "TextBox 8"),
             "Sponsor: Eva Buzo, Executive Director  ·  Agenda item TBC  ·  "
             "15 minutes  ·  Discussion  ·  DRAFT for ED sign-off")

    # 2 — three goals
    s = clone_slide(prs, goals)
    set_text(shape(s, "TextBox 14"),
             "The September retreat broadened our framing from serious international crimes to "
             "human rights violations (slide 9).")
    footers.append((shape(s, "TextBox 15"), shape(s, "TextBox 16")))

    # 3 — where the work has sat (programme and funder only)
    s = clone_slide(prs, table)
    set_text(shape(s, "TextBox 2"), "Where the work has sat since June")
    set_text(shape(s, "TextBox 4"),
             "29 June – 12 October 2026, by funded programme and other significant strands.")
    rows = [
        ("Rohingya in India — legal protection", "Core funding"),
        ("Rohingya dialogue — Bangladesh and Arakan", "FCDO (UK) and Swiss FDFA, via Stimson Center"),
        ("Accountability for business and human rights", "Core funding"),
        ("Public engagement", "Core funding"),
    ]
    # Drop the status column and the unused rows; widen the two remaining columns.
    remove(s, "TextBox 8", *[f"TextBox {12 + 4 * r}" for r in range(len(rows))])
    for r in range(len(rows), 6):
        remove(s, f"Rectangle {9 + 4 * r}", *[f"TextBox {10 + 4 * r + c}" for c in range(3)])
    remove(s, "TextBox 33")
    prog_w, fund_x = 5486400, 6309360
    shape(s, "TextBox 6").width = prog_w
    shape(s, "TextBox 7").left, shape(s, "TextBox 7").width = fund_x, 5212080
    for name in ("TextBox 6", "TextBox 7"):
        shape(s, name).text_frame.paragraphs[0].runs[0].font.name = "Fagun Bold"
    for r, (prog, funder) in enumerate(rows):
        p, fu = shape(s, f"TextBox {10 + 4 * r}"), shape(s, f"TextBox {11 + 4 * r}")
        p.text_frame.paragraphs[0].runs[0].font.name = "Fagun Bold"
        set_text(p, prog)
        set_text(fu, funder)
        p.width = prog_w
        fu.left, fu.width = fund_x, 5212080
    footers.append((shape(s, "TextBox 34"), shape(s, "TextBox 35")))

    # 4 — Rohingya in India
    s = frame(prs, india, "Rohingya in India — legal protection",
              "Core funded  ·  no project grant  ·  Lead: Arati Ranade, India Programme Manager")
    f = Flow(s, protos, width=TEXT_W_PHOTO, bullet_pt=12)
    f.eyebrow("What the strand is")
    f.body("Direct legal intervention for detained Rohingya refugees in India, building a "
           "cross-state network of lawyers and an evidentiary base on detention and deportation "
           "patterns. Funded from core — no donor restriction on case selection or strategy.", gap=110000)
    f.eyebrow("Since 29 June")
    f.bullet("Delhi detention roundtable (5 July) — lawyers from across states convened over three "
             "panels to discuss detention litigation strategy.", gap=110000)
    f.bullet("Right to Education — strategy note drafted on an education-rights route for detained "
             "minors, alongside the refugee route; raised with African states in Geneva (25 August).", gap=110000)
    f.bullet("Litigation pipeline — habeas corpus petition in drafting; bail application being "
             "finalised for counsel's review; Right to Information petitioner being identified.", gap=110000)
    f.bullet("Lawyer network — Jammu & Kashmir leads in place; Rajasthan, Tripura and Hyderabad "
             "being assessed. Case-law report on Rohingya deportation in preparation.", gap=110000)
    f.check("India")
    photo(s, protos, ASSETS / "delhi-detention-roundtable.jpg", CONTENT_TOP, "Delhi, July 2026")
    footers.append((shape(s, "TextBox 18"), shape(s, "TextBox 19")))

    # 5 — Rohingya dialogue (FCDO and Swiss FDFA fund the same project)
    s = frame(prs, india, "Rohingya dialogue — Bangladesh and Arakan",
              "FCDO (UK) and Swiss FDFA, via Stimson Center as prime  ·  hūmānus as implementing partner")
    f = Flow(s, protos, bullet_pt=12)
    f.eyebrow("What the project is")
    f.body("Sustained dialogue and engagement with Rohingya leadership, building trust among "
           "Rohingya, Arakanese and Bangladesh stakeholders through bilateral engagement, capacity "
           "building and facilitated dialogue. FCDO and Swiss FDFA funding support the same "
           "project; both reimburse in arrears.", gap=110000)
    f.eyebrow("Since 29 June")
    f.bullet("Dialogue work continuing with Rohingya leadership.", gap=110000)
    f.bullet("FCDO — hūmānus inputs being prepared for the Stimson-led interim report (June–October); "
             "June–October claim in GBP and Schedule 2 disbursement amounts being agreed.", gap=110000)
    f.bullet("Swiss FDFA — head grant and approved joint workplan requested from Stimson.", gap=110000)
    f.bullet("Analysis — ten reports by Sawyadullah on Rohingya relationships with the "
             "United States, China, Bangladesh, India and Pakistan, to aid the dialogue and support "
             "Rohingya leadership. Website launching November 2026; a workshop on the content will "
             "be run in the Rohingya language.", gap=110000)
    f.check("Dialogue")
    footers.append((shape(s, "TextBox 18"), shape(s, "TextBox 19")))

    # 6 — Business and human rights
    s = clone_slide(prs, bhr)
    set_text(shape(s, "TextBox 4"),
             "Core funded  ·  legal action, regulatory complaint and advocacy — confirmed as a core model at the retreat")
    set_text(shape(s, "TextBox 8"), "Meta / Facebook")
    set_text(shape(s, "TextBox 9"),
             "Meta has responded to our case, including a request for particulars. "
             "[Eva to confirm: source to be supplied] Documentary narrative and video series in production.")
    set_text(shape(s, "TextBox 10"), "Adani")
    set_text(shape(s, "TextBox 11"),
             "Coordination and investigation work continuing, including research on Australian "
             "super-fund exposure, fiduciary-duty case law, the OFAC civil probe and Adani Defence.")
    set_text(shape(s, "TextBox 12"), "Other accountability work")
    set_text(shape(s, "TextBox 13"),
             "US and UK greenwashing opinions (Reuben brothers / Escénica) in progress. "
             "CERD / ICJ brief concerning Australia in preparation.")
    set_text(shape(s, "TextBox 15"),
             "At Q2 we proposed a dedicated Adani paper for this meeting. Coordination and "
             "investigation work is continuing; the board is updated through this report.")
    footers.append((shape(s, "TextBox 16"), shape(s, "TextBox 17")))

    # 7 — Public engagement
    s = frame(prs, india, "Public engagement",
              "Core funded  ·  Against the Tide, speaking, advocacy and donor relations")
    f = Flow(s, protos, width=TEXT_W_TWO_PHOTOS, bullet_pt=12)
    f.eyebrow("Against the Tide")
    f.bullet("Eva Buzo Swim × Small Islands Coalition (28 August); Islands Swim: Seychelles "
             "(15 September); small-islands press, social and documentary session at UNGA (25 September).",
             gap=100000)
    f.bullet("The swim supported the Esencia NO VA environmental campaign, with a public panel and "
             "press coverage (El Vocero).", gap=100000)
    f.eyebrow("Speaking and advocacy")
    talk = f.bullet("Berkeley (28 September) — keynote, Human Rights for Victims of Conflict in a "
                    "Fragmenting World; session with Berkeley IIS clinic students on conflict "
                    "monitoring.\nRecording: youtu.be/_d9m1Zl0Pro", gap=100000)
    link(talk, "youtu.be", BERKELEY_URL)
    f.bullet("Amnesty × hūmānus side event at UNGA (25 September).", gap=100000)
    f.eyebrow("Donor relations")
    f.bullet("OSF partner retreat, Mexico — Finding Each Other on Shifting Ground, bringing together "
             "around 100 OSF partners; good connections made. Preceded by online pre-gathering "
             "sessions in July.", gap=100000)
    f.check("Public engagement")
    y = photo(s, protos, ASSETS / "esencia-no-va-panel.jpg", CONTENT_TOP, "Esencia NO VA panel",
              width=PHOTO_W_SMALL)
    photo(s, protos, ASSETS / "el-vocero-interview.jpg", y, "El Vocero interview",
          width=PHOTO_W_SMALL, height=1900000, crop=(0.12, 0.332))
    footers.append((shape(s, "TextBox 18"), shape(s, "TextBox 19")))

    # 8 — Further activities
    s = frame(prs, india, "Further activities",
              "Advisory work outside our funded programmes")
    f = Flow(s, protos)
    f.eyebrow("Since 29 June")
    f.bullet("Australian war crimes, victim participation — advice to two sets of clients in "
             "Australia on victim participation in prosecutions of alleged war crimes committed "
             "in Afghanistan.")
    f.bullet("Caste-based discrimination — advice on a complaint to the Australian Human Rights "
             "Commission against a New South Wales local council, concerning discrimination against "
             "Dalit, Bahujan and Adivasi communities. Instructed September 2026.")
    f.check("Further activities")
    footers.append((shape(s, "TextBox 18"), shape(s, "TextBox 19")))

    # 9 — Retreat
    s = frame(prs, india, "Strategy retreat — what we decided",
              "Team retreat, September 2026  ·  strategy, impact and organisational focus")
    f = Flow(s, protos, width=TEXT_W_PHOTO, bullet_pt=12)
    f.eyebrow("Strategy and impact")
    f.bullet("Plan from what clients need — an outcomes ladder linking litigation, mediation and "
             "community engagement to meaningful results.", gap=110000)
    f.bullet("Broaden the framing from serious international crimes to human rights violations; "
             "stay people-focused and prioritise systemic change.", gap=110000)
    f.bullet("Corporate accountability — legal action, regulatory complaints and advocacy — "
             "adopted as a core model, including Adani and Facebook.", gap=110000)
    f.bullet("Rohingya — continue the lawyer-network and coordination approach, and take stock "
             "of progress after six months.", gap=110000)
    f.eyebrow("Priorities")
    f.bullet("Keep excellence, rigour and clear accountability central; new geographies only "
             "where needs-driven.", gap=110000)
    f.bullet("Tell the organisation's story well and show impact to clients and funders; "
             "consider dedicated fundraising capacity.", gap=110000)
    f.check("Retreat")
    photo(s, protos, ASSETS / "team-retreat.jpg", CONTENT_TOP, "Team retreat, September 2026")
    footers.append((shape(s, "TextBox 18"), shape(s, "TextBox 19")))

    # 10 — Strategic outlook and watch-items
    s = clone_slide(prs, outlook)
    set_text(shape(s, "TextBox 4"), "Things to keep in view between now and the February 2027 meeting.")
    watch = [
        ("Cash position",
         "Cash USD 33,679 at 30 September (Q3 management accounts, unaudited); no contracted "
         "receipts until the FCDO Phase II claim (indicatively November) and OSF Year 3 (January 2027)."),
        ("Strategic framing",
         "The retreat broadened the framing of the board-approved 2024–29 strategy from serious "
         "international crimes to human rights violations — brought to the board for discussion."),
        ("Rohingya work",
         "Progress to be taken stock of around March 2027, between board meetings; the ED will "
         "report the outcome to the board."),
        ("Fundraising and post-OSF horizon",
         "Dedicated fundraising capacity under consideration — continues the Q2 item 6 discussion on "
         "matching ambition to fundraising."),
        ("Governance and entity",
         "Legal name change VAI → hūmānus notified to OSF grant contacts; staff contract renewals "
         "in progress."),
    ]
    for i, (head, body) in enumerate(watch):
        set_text(shape(s, f"TextBox {6 + 3 * i}"), head)
        set_text(shape(s, f"TextBox {7 + 3 * i}"), body)
    remove(s, "Rectangle 20", "TextBox 21", "TextBox 22")  # Q2 had six items; five here
    footers.append((shape(s, "TextBox 23"), shape(s, "TextBox 24")))

    # 11 — close
    s = clone_slide(prs, close)
    set_text(shape(s, "TextBox 6"),
             "Q3 2026 Board Meeting  ·  Operations & Impact Update  ·  Sponsor: Eva Buzo (ED)  ·  DRAFT")

    total = len(prs.slides) - n_template
    for i, (foot, page) in enumerate(footers, start=2):
        set_text(foot, FOOTER)
        set_text(page, f"{i} / {total}")

    sld_ids = prs.slides._sldIdLst
    for sld_id in list(sld_ids)[:n_template]:
        prs.part.drop_rel(sld_id.rId)
        sld_ids.remove(sld_id)

    prs.save(OUT)
    print(f"wrote {OUT.relative_to(REPO_ROOT)} ({total} slides)")


if __name__ == "__main__":
    build()
