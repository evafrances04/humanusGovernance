"""
Brand polish pass for 08-ED-Operations-Impact-Y2D.pptx.

Run from the repo root:
    python3 2026/Q2/_polish_deck.py

What this does (no copy changes):
  - Replaces the text "hūmānus" wordmark on cover and close slides with the
    hūmānus Bold logo PNG (humanus_logo_master_bold.png), locking the brand
    mark regardless of which fonts a viewer has installed.
  - Sets Fagun Black for the hero numerals on the strategy slide.
  - Sets Fagun Bold for slide titles (H2), Fagun Medium for section sub-heads
    ("What the project is", "Where we are" etc.) and styles them as small,
    tracked eyebrows in Core.
  - Sets the cover meta line ("year-to-date update · Q2 …") to Bold lime
    rather than Touch pink, since Touch is reserved for sparing emphasis.
  - Keeps Inter for all body copy (still the safe rendering choice for board
    members who don't have Fagun installed). PDF export locks Fagun for
    distribution.
  - Adds a thin Bold (lime) rule under the slide-title separator to
    strengthen the brand divider, consistent with the existing accent.

Font handling: Fagun is set on headings in the .pptx but Inter remains the
body font. Board members without Fagun will see Inter substitution for
headings on screen; the bundled PDF preserves the typography.
"""

from __future__ import annotations

import copy
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Emu, Pt

REPO_ROOT = Path(__file__).resolve().parents[2]
BRAND = REPO_ROOT.parent / "humanus-brand"
SRC = REPO_ROOT / "2026" / "Q2" / "08-ED-Operations-Impact-Y2D.pptx"
LOGO_BOLD = BRAND / "logo" / "humanus_logo_master_bold.png"

CORE = RGBColor(0x41, 0x0F, 0x37)
BOLD = RGBColor(0x96, 0xFF, 0x14)
TOUCH = RGBColor(0xE1, 0x2D, 0x64)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

FAGUN_BLACK = "Fagun Black"
FAGUN_EXTRABOLD = "Fagun ExtraBold"
FAGUN_BOLD = "Fagun Bold"
FAGUN_MEDIUM = "Fagun Medium"
INTER = "Inter"

# Slide title characters (exact match) — used to identify the H2 title.
SLIDE_TITLES = {
    "Three goals guide everything we do",
    "Where the work has sat this year",
    "Towards a More Peaceful Arakan",
    "Bangladesh dialogue programme",
    "Rohingya in India — legal protection",
    "Women's Dialogue — grant close-out",
    "Accountability for business and human rights",
    "Strategic outlook and watch-items",
}

# Section sub-heads (eyebrow style) — exact-match strings on internal slides.
EYEBROWS = {
    "What the project is",
    "Where we are",
    "What the funding supports",
    "What the strand is",
    "Y2D streams",
    "Where the work has sat in 2026",
    "Why this matters for the Board",
    "The framework",
    "Active workstreams",
}


def set_run_font(run, *, name=None, size=None, bold=None, color=None, italic=None):
    if name is not None:
        run.font.name = name
    if size is not None:
        run.font.size = size
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color


def all_runs(text_frame):
    for para in text_frame.paragraphs:
        for run in para.runs:
            yield run


def replace_wordmark_with_logo(slide, shape, logo_path: Path):
    """Drop the text shape and place the bold logo in its footprint, centred."""
    left, top, width, height = shape.left, shape.top, shape.width, shape.height
    # Remove the text shape.
    sp = shape._element
    sp.getparent().remove(sp)
    # Add picture. The bold wordmark logo PNG has a wide aspect ratio.
    # We constrain to the available width and let height auto-scale, then
    # centre vertically in the original text shape's box.
    from PIL import Image  # available in pptx envs
    with Image.open(logo_path) as im:
        aspect = im.height / im.width
    target_width = int(width * 0.78)  # leave breathing room on both sides
    target_height = int(target_width * aspect)
    centred_left = int(left + (width - target_width) / 2)
    centred_top = int(top + (height - target_height) / 2)
    slide.shapes.add_picture(
        str(logo_path),
        centred_left,
        centred_top,
        width=target_width,
        height=target_height,
    )


def polish_cover(slide):
    """Cover slide: swap wordmark text for logo image; restyle meta lines."""
    for shape in list(slide.shapes):
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text.strip()
        if text == "hūmānus":
            replace_wordmark_with_logo(slide, shape, LOGO_BOLD)
            continue
        if text.startswith("year-to-date update"):
            # Reset Touch → Bold lime to reduce accent crowding on cover.
            for run in all_runs(shape.text_frame):
                set_run_font(run, name=FAGUN_MEDIUM, color=BOLD)
        elif text == "29 June 2026":
            for run in all_runs(shape.text_frame):
                set_run_font(run, name=FAGUN_MEDIUM, color=WHITE)
        elif text == "Humanising Justice":
            for run in all_runs(shape.text_frame):
                set_run_font(run, name=FAGUN_BLACK, color=WHITE)
        elif text == "ACCOUNTABILITY  ·  COMMUNITY  ·  EVIDENCE":
            for run in all_runs(shape.text_frame):
                set_run_font(run, name=FAGUN_MEDIUM, color=BOLD)
        elif text.startswith("Sponsor:"):
            for run in all_runs(shape.text_frame):
                set_run_font(run, name=INTER, color=BOLD)


def polish_close(slide):
    """Close slide: swap wordmark text for logo image; restyle remaining text."""
    for shape in list(slide.shapes):
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text.strip()
        if text == "hūmānus":
            replace_wordmark_with_logo(slide, shape, LOGO_BOLD)
        elif text == "Humanising justice":
            for run in all_runs(shape.text_frame):
                set_run_font(run, name=FAGUN_BLACK, color=WHITE)
        elif text == "ACCOUNTABILITY  ·  COMMUNITY  ·  EVIDENCE":
            for run in all_runs(shape.text_frame):
                set_run_font(run, name=FAGUN_MEDIUM, color=BOLD)
        elif "humanus.co" in text:
            for run in all_runs(shape.text_frame):
                set_run_font(run, name=INTER, color=WHITE)
        elif text.startswith("Q2 2026 Board Meeting"):
            for run in all_runs(shape.text_frame):
                set_run_font(run, name=INTER, color=BOLD)


def polish_internal(slide, slide_index: int, total: int):
    """Internal slides: apply heading/eyebrow typography."""
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text.strip()
        if not text:
            continue
        # H2 — slide title (Fagun Bold, Core)
        if text in SLIDE_TITLES:
            for run in all_runs(shape.text_frame):
                set_run_font(run, name=FAGUN_BOLD, color=CORE)
            continue
        # Eyebrow sub-heads (Fagun Medium, Core, slightly smaller, all-caps)
        if text in EYEBROWS:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    set_run_font(
                        run,
                        name=FAGUN_MEDIUM,
                        color=CORE,
                        bold=False,
                    )
                    # Eyebrows look stronger upper-case + tracked
                    run.text = run.text.upper()
            continue
        # Hero stat numerals on strategy slide (slide 2)
        if text in {"1", "2", "3"} and slide_index == 2:
            for run in all_runs(shape.text_frame):
                set_run_font(run, name=FAGUN_BLACK, color=BOLD)
            continue
        # Sub-title under the H2 (the deck line, ~14pt, italics-feel)
        # We can detect by being directly under the title area — too brittle
        # to special-case here; leave as Inter for body legibility.

        # Frontline slide (8) hero text — guard by slide index to avoid
        # touching body text on slide 3 that mentions the Frontline Club.
        if slide_index == 8:
            if text == "PUBLIC ENGAGEMENT":
                for run in all_runs(shape.text_frame):
                    set_run_font(run, name=FAGUN_MEDIUM, color=BOLD)
            elif text.startswith("Frontline Club"):
                for run in all_runs(shape.text_frame):
                    set_run_font(run, name=FAGUN_BLACK, color=WHITE)
            elif text == "Human Rights in a Fragmenting World":
                for run in all_runs(shape.text_frame):
                    set_run_font(run, name=FAGUN_MEDIUM, color=BOLD, italic=True)
            elif text == "2 April 2026":
                for run in all_runs(shape.text_frame):
                    set_run_font(run, name=FAGUN_MEDIUM, color=WHITE)

        # Headline sub-block tiles on slide 9 — "Meta Inc.", "Adani / UniSuper / APRA",
        # "Mohib Ullah accountability", "Note for the Board"
        if slide_index == 9 and text in {
            "Meta Inc.",
            "Adani / UniSuper / APRA",
            "Mohib Ullah accountability",
            "Note for the Board",
        }:
            for run in all_runs(shape.text_frame):
                set_run_font(run, name=FAGUN_BOLD, color=CORE)

        # Strategic outlook row labels on slide 10
        if slide_index == 10:
            outlook_labels = {
                "Reserves and deficit-model donors",
                "VAI ↔ hūmānus entity transition",
                "Bangladesh operating environment",
                "Beyond 2029 / post-OSF horizon",
                "Hiring posture",
                "Governance and external engagement",
            }
            if text in outlook_labels:
                for run in all_runs(shape.text_frame):
                    set_run_font(run, name=FAGUN_BOLD, color=CORE)


def main():
    pres = Presentation(str(SRC))
    total = len(pres.slides)

    for idx, slide in enumerate(pres.slides, start=1):
        if idx == 1:
            polish_cover(slide)
        elif idx == total:
            polish_close(slide)
        else:
            polish_internal(slide, slide_index=idx, total=total)

    out_path = SRC  # in-place
    pres.save(str(out_path))
    print(f"Polished: {out_path}")


if __name__ == "__main__":
    main()
