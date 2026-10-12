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
retreat summary (CAP-518 comment, 12 Oct 2026), Stacy McGill's Q3 2026 Board
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
        add_clone(self.s, self.p["bullet"], self.left + 274320, self.y, tw, h, text)
        self.y += h + gap

    def space(self, emu=120000):
        self.y += emu

    def check(self, label):
        if self.y > CONTENT_BOTTOM:
            print(f"WARNING: {label} overflows by {(self.y - CONTENT_BOTTOM) / 914400:.2f} in")


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
    }
    footers = []  # (footer shape, page shape)

    # 1 — cover
    s = clone_slide(prs, cover)
    set_text(shape(s, "TextBox 4"), "update since 29 June  ·  Q3 2026 Board Meeting")
    set_text(shape(s, "TextBox 5"), "29 October 2026")
    set_text(shape(s, "TextBox 8"),
             "Sponsor: Eva Buzo, Executive Director  ·  Agenda item [Eva to confirm]  ·  "
             "15 minutes  ·  Discussion  ·  DRAFT for ED sign-off")

    # 2 — three goals
    s = clone_slide(prs, goals)
    set_text(shape(s, "TextBox 14"),
             "The September retreat proposed broadening our framing from serious international "
             "crimes to human rights violations (slide 10). [Eva to confirm: Note or Endorse on 29 Oct agenda]")
    footers.append((shape(s, "TextBox 15"), shape(s, "TextBox 16")))

    # 3 — where the work has sat
    s = clone_slide(prs, table)
    set_text(shape(s, "TextBox 2"), "Where the work has sat since June")
    set_text(shape(s, "TextBox 4"),
             "29 June – 12 October 2026, by funded programme and other significant strands.")
    set_text(shape(s, "TextBox 8"), "Status since 29 June")
    rows = [
        ("Rohingya in India — legal protection", "Core funding",
         "Delhi detention roundtable convened lawyers (5 Jul); Right to Education route for "
         "detained minors developed; habeas and bail filings in preparation."),
        ("Bangladesh dialogue", "FCDO (UK)",
         "Phase I delivery continuing; inputs to Stimson-led interim report (Jun–Oct); "
         "Jun–Oct financial claim in preparation."),
        ("Towards a More Peaceful Arakan", "Swiss FDFA via Stimson Center",
         "Rohingya dialogue continuing; RCC roundtable on conditions for return at UNGA (25 Sep)."),
        ("Accountability for business and human rights", "Core funding",
         "Meta response to our case; Amnesty × hūmānus UNGA side event; Adani research dossier."),
        ("Public engagement", "Core funding",
         "Against the Tide swim (Small Islands Coalition, Seychelles, UNGA); Berkeley keynote (28 Sep)."),
        ("Research — Rohingya analysis", "[Eva to confirm]",
         "Sawyeddollah's analysis reports on Rohingya; website in development."),
    ]
    # Q2 mixed Fagun Bold and Inter bold across rows; Inter bold falls back badly
    # where Inter isn't installed, so set row and header labels in Fagun Bold.
    for name in ("TextBox 6", "TextBox 7", "TextBox 8"):
        shape(s, name).text_frame.paragraphs[0].runs[0].font.name = "Fagun Bold"
    for r, (prog, funder, status) in enumerate(rows):
        shape(s, f"TextBox {10 + 4 * r}").text_frame.paragraphs[0].runs[0].font.name = "Fagun Bold"
        set_text(shape(s, f"TextBox {10 + 4 * r}"), prog)
        set_text(shape(s, f"TextBox {11 + 4 * r}"), funder)
        set_text(shape(s, f"TextBox {12 + 4 * r}"), status)
    set_text(shape(s, "TextBox 33"),
             "Women's Dialogue (SDC) closed 30 April 2026; no activity this period. "
             "[Eva to confirm] Finance: see Stacy McGill's Q3 Board Finance Report and Forecast Report.")
    footers.append((shape(s, "TextBox 34"), shape(s, "TextBox 35")))

    # 4 — Rohingya in India
    s = frame(prs, india, "Rohingya in India — legal protection",
              "Core funded  ·  no project grant  ·  Lead: Arati Ranade, India Programme Manager")
    f = Flow(s, protos, bullet_pt=12)
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
    f.eyebrow("Why this matters for the board")
    f.body("The retreat agreed to continue the lawyer-network approach and reassess the Rohingya "
           "project after six months (slide 10). [Eva to confirm: report the review in writing or verbally]")
    f.check("India")
    footers.append((shape(s, "TextBox 18"), shape(s, "TextBox 19")))

    # 5 — Bangladesh dialogue (FCDO)
    s = frame(prs, india, "Bangladesh dialogue programme", "FCDO (UK)  ·  continuing programme")
    f = Flow(s, protos)
    f.eyebrow("What the funding supports")
    f.body("Sustained dialogue and engagement work with Rohingya leadership in Bangladesh, "
           "complementing the Swiss/Stimson Arakan partnership. Operates in a deficit / "
           "reimbursement model — hūmānus spends and is reimbursed in arrears.")
    f.eyebrow("Since 29 June")
    f.bullet("Phase I delivery continuing.")
    f.bullet("hūmānus inputs being prepared for the Stimson-led FCDO interim report covering June–October.")
    f.bullet("June–October financial claim being prepared in GBP; Schedule 2 disbursement amounts "
             "being agreed and countersigned.")
    f.bullet("Cash timing — Stacy McGill's Q3 report shows the FCDO Phase II claim as the next "
             "contracted receipt, indicatively November.")
    f.check("FCDO")
    footers.append((shape(s, "TextBox 18"), shape(s, "TextBox 19")))

    # 6 — Peaceful Arakan
    s = frame(prs, india, "Towards a More Peaceful Arakan",
              "Swiss FDFA, via Stimson Center as prime  ·  hūmānus as implementing partner")
    f = Flow(s, protos)
    f.eyebrow("What the project is")
    f.body("An 18-month partnership (December 2025 – May 2027) to identify, create, and support "
           "trust-building opportunities among Rohingya, Arakanese, and Bangladesh stakeholders, "
           "combining bilateral engagement, capacity building, and facilitated dialogue.")
    f.eyebrow("Since 29 June")
    f.bullet("Rohingya dialogue work continuing with Rohingya leadership.")
    f.bullet("UNGA week — RCC High-Level Roundtable, The Conditions for Return: A Rohingya-Led "
             "Dialogue (25 September).")
    f.bullet("Request made to Stimson for the FDFA head grant and the approved joint workplan.")
    f.bullet("Related peace-network engagement — PHRD Myanmar partner discussion (16 July); "
             "Pre-Gathering online sessions, Finding Each Other on Shifting Ground (13 and 16 July); "
             "Swiss Peace Partners Network, Geneva (6 August). [Eva to confirm placement]")
    f.check("Arakan")
    footers.append((shape(s, "TextBox 18"), shape(s, "TextBox 19")))

    # 7 — Business and human rights
    s = clone_slide(prs, bhr)
    set_text(shape(s, "TextBox 4"),
             "Core funded  ·  legal action, regulatory complaint and advocacy — confirmed as a core model at the retreat")
    set_text(shape(s, "TextBox 8"), "Meta / Facebook")
    set_text(shape(s, "TextBox 9"),
             "Meta has responded to our case, including a request for particulars. "
             "[Eva to confirm: source to be supplied] Amnesty × hūmānus and Meta Inc. side events "
             "at UNGA (25 Sep); documentary narrative and video series in production.")
    set_text(shape(s, "TextBox 10"), "Adani / UniSuper / APRA")
    set_text(shape(s, "TextBox 11"),
             "Team research dossier progressing: UniSuper holdings disclosure, super-fund exposure "
             "and GQG mandates, Cbus deep dive, OFAC civil probe, Australian fiduciary-duty case "
             "law, thermal-coal revenue share, Adani Defence.")
    set_text(shape(s, "TextBox 12"), "Other accountability work")
    set_text(shape(s, "TextBox 13"),
             "US and UK greenwashing opinions (Reuben brothers / Escénica) in progress. "
             "CERD / ICJ brief concerning Australia in preparation.")
    set_text(shape(s, "TextBox 15"),
             "At Q2 we proposed a dedicated Adani / UniSuper paper for this meeting. "
             "[Eva to confirm: delivered, deferred (with reason) or dropped]")
    footers.append((shape(s, "TextBox 16"), shape(s, "TextBox 17")))

    # 8 — Public engagement
    s = frame(prs, india, "Public engagement", "Core funded  ·  Against the Tide, speaking and advocacy")
    f = Flow(s, protos)
    f.eyebrow("Against the Tide")
    f.bullet("Eva Buzo Swim × Small Islands Coalition (28 August); Islands Swim: Seychelles "
             "(15 September); small-islands press, social and documentary session at UNGA (25 September).")
    f.eyebrow("Speaking and advocacy")
    f.bullet("Berkeley (28 September) — keynote, Human Rights for Victims of Conflict in a "
             "Fragmenting World; session with Berkeley IIS clinic students on conflict monitoring.")
    f.bullet("UNGA week (21–25 September) — Afrobarometer–OSF side event Beyond Borders (24 Sep), "
             "OSF reception and network partners dinner, alongside the Amnesty, Meta and RCC events "
             "on slides 6 and 7.")
    f.bullet("9th anniversary of the Rohingya genocide (24 August) — communications moment.")
    f.bullet("Media bank (B-roll and photography) being built for future communications.")
    f.check("Public engagement")
    footers.append((shape(s, "TextBox 18"), shape(s, "TextBox 19")))

    # 9 — Research
    s = frame(prs, india, "Research — Rohingya analysis",
              "M. Sawyeddollah [Eva to confirm spelling]  ·  funding line [Eva to confirm]")
    f = Flow(s, protos)
    f.eyebrow("What the strand is")
    f.body("Analysis reports on the Rohingya situation by M. Sawyeddollah, and a website he is "
           "building to host the work. [Eva to confirm framing]")
    f.eyebrow("Since 29 June")
    f.bullet("Analysis papers discussed with him in New York during UNGA week (23 September).")
    f.bullet("Website in development. [Eva to confirm launch timing]")
    f.bullet("Scope and audience of the reports. [Eva to confirm]")
    f.check("Research")
    footers.append((shape(s, "TextBox 18"), shape(s, "TextBox 19")))

    # 10 — Retreat: strategy
    s = frame(prs, india, "Strategy retreat — what we decided",
              "Team retreat, September 2026  ·  Day 1: strategy, impact and organisational focus")
    f = Flow(s, protos, bullet_pt=12)
    f.eyebrow("Strategy and impact")
    f.bullet("Plan from what clients need — an outcomes ladder linking litigation, mediation and "
             "community engagement to meaningful results.", gap=110000)
    f.bullet("Broaden the framing from serious international crimes to human rights violations; "
             "stay people-focused and prioritise systemic change. "
             "[Eva to confirm: Note or Endorse on 29 Oct agenda]", gap=110000)
    f.bullet("Corporate accountability — legal action, regulatory complaints and advocacy — "
             "adopted as a core model, including Adani and Facebook.", gap=110000)
    f.bullet("Rohingya — continue the lawyer-network and coordination approach; reassess after "
             "six months. [Eva to confirm: written or verbal]", gap=110000)
    f.eyebrow("Organisational choices")
    f.bullet("Be realistic about capacity; close or refocus dormant or ineffective projects "
             "without a credible path to impact. [Eva to confirm: name projects or keep principled]",
             gap=110000)
    f.bullet("Keep excellence, rigour and clear accountability central; new geographies only "
             "where needs-driven.", gap=110000)
    f.bullet("Tell the organisation's story well, show impact to clients and funders; consider "
             "dedicated fundraising capacity. [Eva to confirm: decided or exploratory]", gap=110000)
    f.check("Retreat strategy")
    footers.append((shape(s, "TextBox 18"), shape(s, "TextBox 19")))

    # 11 — Retreat: ways of working
    s = frame(prs, india, "Strategy retreat — how we work",
              "Team retreat, September 2026  ·  Day 2: culture, collaboration and ways of working")
    f = Flow(s, protos, bullet_pt=12)
    f.eyebrow("Ways of working")
    f.bullet("Channels — Motion for tasks and coordination, Signal for quick contact, email for "
             "external and out-of-hours.", gap=110000)
    f.bullet("Rhythm — full-team meetings fortnightly, project meetings on alternate weeks, "
             "15-minute calls for quick clarification.", gap=110000)
    f.bullet("Supportive accountability — flag delays and dependencies early; clear owners, "
             "timelines and expectations; setbacks treated as learning.", gap=110000)
    f.bullet("Audiences — \"humanizing justice\" for clients seeking remedies and funders seeking "
             "measurable impact; policymakers, lawyers and media as multipliers.", gap=110000)
    f.bullet("People — regionally grounded recruitment, internships and fellowships; external "
             "mentors and internal peer support.", gap=110000)
    f.eyebrow("Follow-up themes")
    f.body("Connect activities and resource requests to client outcomes; reassess work without a "
           "realistic route to impact, including the Rohingya project after the review period; "
           "strengthen public storytelling; keep refining strategy and practice as capacity evolves.")
    f.check("Retreat ways of working")
    footers.append((shape(s, "TextBox 18"), shape(s, "TextBox 19")))

    # 12 — Strategic outlook and watch-items
    s = clone_slide(prs, outlook)
    set_text(shape(s, "TextBox 4"), "Things to keep in view between now and the February 2027 meeting.")
    watch = [
        ("Cash position",
         "Cash USD 33,679 at 30 September; no contracted receipts until the FCDO Phase II claim "
         "(indicatively November) and OSF Year 3 (January 2027). Per Stacy McGill's Q3 Board Finance Report."),
        ("Strategic framing",
         "The retreat proposes broadening the board-approved 2024–29 strategy from serious "
         "international crimes to human rights violations. [Eva to confirm: Note or Endorse]"),
        ("Rohingya programme review",
         "Six-month reassessment agreed in September falls around March 2027, between board "
         "meetings. [Eva to confirm: how and when the board hears the outcome]"),
        ("Prioritisation",
         "Dormant or ineffective projects to be closed or refocused. [Eva to confirm: name projects or keep principled]"),
        ("Fundraising and post-OSF horizon",
         "Dedicated fundraising capacity under consideration — continues the Q2 item 6 discussion on "
         "matching ambition to fundraising. [Eva to confirm: decided or exploratory]"),
        ("Governance and entity",
         "Legal name change VAI → hūmānus notified to OSF grant contacts; staff contract renewals "
         "in progress. Finance detail in Stacy McGill's papers."),
    ]
    for i, (head, body) in enumerate(watch):
        set_text(shape(s, f"TextBox {6 + 3 * i}"), head)
        set_text(shape(s, f"TextBox {7 + 3 * i}"), body)
    footers.append((shape(s, "TextBox 23"), shape(s, "TextBox 24")))

    # 13 — close
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
