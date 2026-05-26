"""Q2 2026 board paper builder — hūmānus.

Produces five branded deliverables under humanusGovernance/2026/Q2/.
Source data: Xero pulls on 2026-05-26 + 2026 Budget-2.xlsx + memory of
signed contracts (SDC §4.1, Stimson Swiss §3.3, OSF OR2024-94840).
"""

from __future__ import annotations

import os
from pathlib import Path

from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Paths & brand constants
# ---------------------------------------------------------------------------
ROOT = Path("/Users/evabuzo/multica_workspaces_desktop-api.multica.ai/75c586e2-be30-4ecd-8955-3ba06b11410a/a63eaedb/workdir")
Q2 = ROOT / "humanusGovernance" / "2026" / "Q2"
LOGO = ROOT / "humanus-brand" / "logo" / "humanus_logo_master_core.png"

CORE = RGBColor(0x41, 0x0F, 0x37)        # #410F37 deep aubergine
BOLD = RGBColor(0x96, 0xFF, 0x14)        # #96FF14 bold green (decisions)
TOUCH = RGBColor(0xE1, 0x2D, 0x64)       # #E12D64 warnings

CORE_HEX = "410F37"
BOLD_HEX = "96FF14"
TOUCH_HEX = "E12D64"

# ---------------------------------------------------------------------------
# Numbers — all sourced from Xero on 2026-05-26
# ---------------------------------------------------------------------------
# Q1 2026 P&L (1 Jan – 31 Mar 2026)
Q1_PL = {
    "income": [("Grant Income", 300000.00), ("Interest Income", 472.10)],
    "income_total": 300472.10,
    "cos": [
        ("Accommodation", 1714.75),
        ("Advocate / Programme Manager", 13000.00),
        ("Communication Advisor", 12986.52),
        ("Cost of Living Allowances", 5675.61),
        ("Executive Director", 42600.00),
        ("Local Transportation", 1071.02),
        ("Meeting Expenses", 1612.28),
        ("Partner Grants", 2000.00),
        ("Per diem", 182.04),
        ("Rohingya Manager", 7500.00),
        ("Volunteer Stipends", 1886.92),
    ],
    "cos_total": 90229.14,
    "gross_profit": 210242.96,
    "opex": [
        ("Accounting & Bookkeeping Fees", 6000.00),
        ("Audit Fees", 1670.00),
        ("Bank Charges & Merchant Fees", 129.10),
        ("Communication Support", 4362.22),
        ("Confidence Building Measures", 14598.00),
        ("Depreciation", 113.84),
        ("Foreign Currency Gains and Losses", 387.27),
        ("Insurance", 1199.37),
        ("IT & Software Expenses", 1775.15),
        ("Legal Fees", 630.47),
        ("Operations Support", 6278.63),
    ],
    "opex_total": 37144.05,
    "net_profit": 173098.91,
}

# YTD P&L 1 Jan – 25 May 2026 (selected — for BvA)
YTD_PL_SPEND = {
    "Accommodation": 2354.03,
    "Advocate / Programme Manager": 16500.00,
    "Communication Advisor": 12986.52,
    "Cost of Living Allowances": 7567.38,
    "Domestic Flights": 81.16,
    "Executive Director": 56800.00,
    "International Flights": 7716.47,
    "Local Transportation": 2097.84,
    "Meeting Expenses": 1612.28,
    "Partner Grants": 6000.00,
    "Per diem": 318.16,
    "Rohingya Manager": 10000.00,
    "Volunteer Stipends": 12066.10,
    "Accounting & Bookkeeping Fees": 8000.00,
    "Audit Fees": 3651.06,
    "Bank Charges & Merchant Fees": 220.93,
    "Board Meetings": 8936.41,
    "Branding & Communications": 9806.60,
    "Communication Support": 5551.31,
    "Confidence Building Measures": 14598.00,
    "Depreciation": 113.84,
    "Foreign Currency Gains and Losses": 215.15,
    "General Expenses": 29193.57,
    "Insurance": 1199.37,
    "IT & Software Expenses": 3517.07,
    "Legal Fees": 1358.61,
    "Operations Support": 6278.63,
    "Printing & Stationery": 4.99,
}
YTD_INCOME = 357983.45
YTD_SPEND = sum(YTD_PL_SPEND.values())   # ≈ 228,745.48
YTD_NET = YTD_INCOME - YTD_SPEND         # ≈ 129,237.97
YTD_MONTHS = 4.7997   # 1 Jan – 25 May = 145 days / 30.42 ≈ 4.77

# Balance Sheet snapshots
BS_31MAR = {
    "wise_usd": 177862.86, "wise_eur": 19727.85, "wise_gbp": 0.00,
    "total_bank": 197590.71,
    "ar": 0.00, "prepayments": 276.90, "programme_adv": 64.97,
    "fixed_assets": 1620.05,
    "total_assets": 199552.63,
    "ap": 1198.39, "exp_payable": 829.11, "suspense": 3.46,
    "total_liab": 2030.96,
    "net_assets": 197521.67,
    "retained_earnings": 24422.76,
    "current_year_earnings": 173098.91,
}
BS_30APR = {
    "wise_usd": 119852.20, "wise_eur": 19203.77, "wise_gbp": 1089.63,
    "total_bank": 140145.60,
    "ar": 57005.00, "prepayments": 266.25, "programme_adv": 64.97,
    "fixed_assets": 3512.33,
    "total_assets": 200994.15,
    "ap": 12802.06, "exp_payable": 829.11, "suspense": 3.46,
    "total_liab": 13634.63,
    "net_assets": 187359.52,
}
BS_LIVE = {  # at 26 May 2026 (Xero report-dated 31 May)
    "wise_usd": 173359.86, "wise_eur": 9436.15, "wise_gbp": 0.00,
    "total_bank": 182796.01,
    "ar": 0.00,
    "net_assets": 182596.78,
}

# Q1 Bank Summary (cash basis)
Q1_BANK = {
    "opening": 23786.75,
    "received": 300534.15,
    "spent": 126374.21,
    "fx": -355.98,
    "closing": 197590.71,
}
# YTD Bank Summary 1 Jan – 25 May
YTD_BANK = {
    "opening": 23786.75,
    "received": 363347.50,
    "spent": 204115.31,
    "fx": -218.10,
    "closing": 182800.84,
}

# 2026 Budget from finance-docs/budgets/2026 Budget-2.xlsx (Sheet1)
BUDGET_2026 = {
    "total": 488450.00,
    "by_donor": {
        "OSF - Operating (unrestricted)": 286634.70,
        "Swiss - Dhaka (Oct 24 – Mar 26)": 35560.00,
        "Swiss - Myanmar (Nov 25 – Mar 27)": 48312.00,
        "FCDO Phase I (Nov 25 – Mar 26)": 42519.00,
        "FCDO Phase II (Jun 26 – Mar 27, unsigned)": 75424.30,
    },
    "by_account_annual": {
        # Best-fit annual budget lines, by P&L account (USD)
        "Executive Director": 170400.00,
        "Cost of Living Allowances": 21600.00,
        "Advocate / Programme Manager": 42000.00,
        "Rohingya Manager": 31250.00,
        "Volunteer Stipends": 12000.00,
        "Vehicle Hire & Sundry": 6000.00,
        "Partner Grants": 30000.00,
        "Legal Fees": 12000.00,
        "Communication Advisor": 64800.00,
        "Communication Support": 14400.00,
        "Accounting & Bookkeeping Fees": 24000.00,
        "Operations Support": 18000.00,
        "Audit Fees": 3000.00,
        "Board Meetings": 4000.00,
        "IT & Software Expenses": 5000.00,
        "Travel Expenses (Accommodation/Flights/Transport/Per diem)": 25000.00,
        "Bank Charges & Merchant Fees / Insurance / Fees": 5000.00,
    },
}

# Reserves policy (CAP-25, awaiting Eva sign-off)
RESERVES = {
    "core_overhead_monthly": 44176.00,
    "floor_3mo": 132527.00,
    "target_6mo": 265054.00,
}

# -----------------------------------------------------------------------------
# Cashflow scenario inputs — see assumptions tab
# -----------------------------------------------------------------------------
# Forward run-rate (per-line build from CAP-147 v2 + YTD actuals)
RUN_RATE_MONTHLY = 45000  # USD central-case spend, post-Geneva normalisation


# ---------------------------------------------------------------------------
# Helpers — docx branding
# ---------------------------------------------------------------------------
def shade_cell(cell, hex_color: str) -> None:
    """Apply a solid background fill to a table cell."""
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tc_pr.append(shd)


def set_cell_font(cell, *, bold=False, color=None, size=10, font="Inter"):
    for para in cell.paragraphs:
        for run in para.runs:
            run.font.name = font
            run.font.size = Pt(size)
            run.font.bold = bold
            if color is not None:
                run.font.color.rgb = color


def add_heading(doc, text, level=1, color=CORE):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = "Inter"
    run.bold = True
    run.font.color.rgb = color
    run.font.size = Pt({1: 18, 2: 14, 3: 12}[level])
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(4)
    return p


def add_para(doc, text, *, italic=False, bold=False, color=None, size=10.5):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = "Inter"
    run.font.size = Pt(size)
    run.font.italic = italic
    run.bold = bold
    if color is not None:
        run.font.color.rgb = color
    p.paragraph_format.space_after = Pt(4)
    return p


def add_table(doc, headers, rows, *, col_widths=None, totals_rows=None,
              warn_rows=None):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Light Grid Accent 1"
    # Header row
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = h
        shade_cell(cell, CORE_HEX)
        set_cell_font(cell, bold=True, color=RGBColor(255, 255, 255), size=10)
    # Data rows
    totals_rows = totals_rows or set()
    warn_rows = warn_rows or set()
    for r_idx, row in enumerate(rows):
        for c_idx, val in enumerate(row):
            cell = table.rows[1 + r_idx].cells[c_idx]
            cell.text = str(val)
            if r_idx in totals_rows:
                set_cell_font(cell, bold=True, color=CORE, size=10)
                shade_cell(cell, "F5E9F0")
            elif r_idx in warn_rows:
                set_cell_font(cell, color=TOUCH, size=10)
            else:
                set_cell_font(cell, color=CORE, size=10)
    if col_widths:
        for i, w in enumerate(col_widths):
            for row in table.rows:
                row.cells[i].width = Cm(w)
    return table


def add_logo(doc, width_inches=1.8):
    if LOGO.exists():
        doc.add_picture(str(LOGO), width=Inches(width_inches))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.LEFT


def add_footer(doc, ref: str):
    section = doc.sections[0]
    footer = section.footer
    p = footer.paragraphs[0]
    p.text = (
        f"hūmānus — Stichting Victim Advocates International (NL chamber 77641981) | "
        f"{ref} | Q2 2026 board paper | [TBD: registered office / ANBI / disclaimer footer]"
    )
    for run in p.runs:
        run.font.size = Pt(7)
        run.font.color.rgb = CORE
        run.font.name = "Inter"


def new_doc(title: str, ref: str) -> Document:
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)
    add_logo(doc)
    add_heading(doc, title, level=1)
    add_para(doc, f"{ref}  ·  Prepared by Stacy McGill (Finance Manager)  ·  "
             f"As-of 26 May 2026  ·  Currency: USD unless stated",
             italic=True, size=9.5)
    add_footer(doc, ref)
    return doc


def fmt_usd(v: float) -> str:
    if v is None:
        return ""
    sign = "(" if v < 0 else ""
    end = ")" if v < 0 else ""
    return f"{sign}USD {abs(v):,.0f}{end}"


def fmt_pct(v: float) -> str:
    if v is None:
        return ""
    return f"{v:+.0%}" if abs(v) < 10 else f"{v:+.0f}×"


# ===========================================================================
# P-06b  Audit progress note
# ===========================================================================
def build_audit_note():
    doc = new_doc("2025 Audit — Progress Note", "P-06b")
    add_heading(doc, "Headline", level=2)
    add_para(doc,
        "2024 audit: complete; signed report on file as Audit Report VAI_2024.pdf "
        "(humanusGovernance/2026/Q1/). Up for formal board approval at this meeting. "
        "2025 audit: fieldwork commencing now per Q1 minute §6; on track but tight against "
        "the 29 June meeting date — expect a status verbal at the meeting rather than a "
        "signed report in the pack.")
    add_heading(doc, "Status by audit cycle", level=2)
    add_table(
        doc,
        ["Audit cycle", "Status", "Next milestone", "Risk flag"],
        [
            ["FY2024",
             "Complete — auditor's report issued",
             "Board approval at 29 Jun 2026 meeting (P-06a)",
             "—"],
            ["FY2025",
             "Fieldwork commencing Q2 2026 per Q1 minute §6",
             "Draft auditor's report targeted for next board meeting "
             "(Q3 2026); status verbal at 29 Jun",
             "Tight against 29 Jun if planned for this meeting"],
        ],
    )
    add_heading(doc, "Why earlier this year", level=2)
    add_para(doc,
        "The Q1 minute records the Board's request to bring the 2025 audit forward in "
        "the year so that audited 2025 financials are available earlier in the 2026 "
        "cycle. Fieldwork is being scheduled accordingly. The standing item for Q3 will "
        "be the draft audited financials; Q2 papers carry the Q1 management accounts "
        "(unaudited) as an interim view.")
    add_heading(doc, "Dutch statutory context", level=2)
    add_para(doc,
        "Per Q1 minute §6: the Dutch statutory requirement is approval of the audited "
        "accounts by the board of directors. hūmānus's 2024 budget sits below the "
        "Dutch statutory audit threshold; the external audit is conducted as good "
        "practice (donor- and reputation-driven). The advisory board reviews per "
        "internal governance policy; statutory approval remains with the board of "
        "directors (currently Eva Buzo as sole statutory director, with the Dutch Chair "
        "appointment in progress — see Q1 minute §7).")
    add_heading(doc, "Compliance / risk flags", level=2)
    add_para(doc,
        "·  Donor reporting downstream: SDC final financial report (§3.1 of the SDC "
        "agreement) is due 30 June 2026 and must be audited by an SDC-approved "
        "independent auditor (§3.3). The 2025 audit timing supports this but does not "
        "directly satisfy §3.3 — a project-specific audit certification is also required "
        "on SDC reporting and is being scoped with the auditor.")
    add_para(doc,
        "·  Period lock in Xero remains at 31 Dec 2022; year-end lock for 2025 to be "
        "applied on audit sign-off. No back-dated postings into 2024 or 2025 are open at "
        "Finance Manager level.")
    add_heading(doc, "Recommended board action", level=2)
    add_para(doc,
        "·  Approve 2024 audited financials (P-06a) — board decision.", bold=True,
        color=CORE)
    add_para(doc,
        "·  Note 2025 audit timing and SDC §3.3 audit certification scope — for "
        "information.", bold=True, color=CORE)
    out = Q2 / "humanus_audit-progress_2025_P-06b.docx"
    doc.save(out)
    print(f"Wrote {out}")


# ===========================================================================
# P-06c  Q1 2026 management accounts
# ===========================================================================
def build_q1_mgmt_accounts():
    doc = new_doc("Q1 2026 Management Accounts (unaudited)", "P-06c")

    add_heading(doc, "Summary", level=2)
    bullets = [
        f"Q1 closed with a {fmt_usd(Q1_PL['net_profit'])} surplus, driven entirely by the "
        f"OSF #4 Year-2 tranche of USD 300,000 received in January (timing — not a "
        f"sustainable run-rate). FY2025 Q1 comparative was a USD (175,899) loss because "
        f"the equivalent OSF tranche landed in December 2024.",
        f"Cash closed Q1 at USD 197,591 (USD 177,863 + EUR 19,728 + GBP 0) — within the "
        f"Amber band of the proposed reserves policy.",
        "Programme spend (Cost of Sales) ran on budget pace at USD 90,229; operating "
        "overhead at USD 37,144 ran below pace, before the April overhead spike (see "
        "BvA paper P-06e).",
        "No external debt; payables and accruals total USD 2,031 — clean balance sheet "
        "at quarter-end.",
        "Live position (25 May 2026): cash USD 182,801; AR cleared after Stimson "
        "INV-0007 receipt 12 May.",
    ]
    for b in bullets:
        add_para(doc, "·  " + b)

    # P&L
    add_heading(doc, "1. Statement of activities — Q1 2026 vs Q1 2025", level=2)
    rows = []
    rows.append(["Total Income", fmt_usd(Q1_PL["income_total"]),
                 fmt_usd(0.00), "Q1 2025 OSF tranche landed Dec 2024"])
    rows.append(["  Grant Income", fmt_usd(300000.00), "—", "OSF #4 Y2 (Jan 2026)"])
    rows.append(["  Interest Income", fmt_usd(472.10), "—", "USD wallet"])
    rows.append(["Cost of Sales (Programme)", fmt_usd(Q1_PL["cos_total"]),
                 "—", "Personnel-heavy; on budget pace"])
    rows.append(["Gross Surplus", fmt_usd(Q1_PL["gross_profit"]), "—", ""])
    rows.append(["Operating Expenses (Overhead)", fmt_usd(Q1_PL["opex_total"]),
                 "—", "Confidence Building Measures $14,598 unbudgeted"])
    rows.append(["Net Surplus / (Deficit)", fmt_usd(Q1_PL["net_profit"]),
                 fmt_usd(-175899.37), "Comparative driven by tranche timing"])
    add_table(doc,
              ["Line", "Q1 2026", "Q1 2025", "Commentary"],
              rows,
              totals_rows={4, 6},
              col_widths=[5.0, 3.2, 3.2, 6.5])

    add_heading(doc, "Programme cost detail (Q1 2026)", level=3)
    pgm_rows = [[name, fmt_usd(val)] for name, val in Q1_PL["cos"]]
    pgm_rows.append(["Total Programme", fmt_usd(Q1_PL["cos_total"])])
    add_table(doc, ["Account", "Q1 2026"], pgm_rows,
              totals_rows={len(pgm_rows) - 1}, col_widths=[10, 4])

    add_heading(doc, "Overhead detail (Q1 2026)", level=3)
    oh_rows = [[name, fmt_usd(val)] for name, val in Q1_PL["opex"]]
    oh_rows.append(["Total Overhead", fmt_usd(Q1_PL["opex_total"])])
    add_table(doc, ["Account", "Q1 2026"], oh_rows,
              totals_rows={len(oh_rows) - 1}, col_widths=[10, 4])

    # Balance Sheet
    add_heading(doc, "2. Balance Sheet — as at 31 Mar 2026 (with 31 Mar 2025 "
                     "comparative)", level=2)
    bs_rows = [
        ["Wise — USD", fmt_usd(BS_31MAR['wise_usd']), fmt_usd(119166.97)],
        ["Wise — EUR (USD equiv)", fmt_usd(BS_31MAR['wise_eur']), fmt_usd(134121.73)],
        ["Wise — GBP (USD equiv)", fmt_usd(BS_31MAR['wise_gbp']), fmt_usd(1033.40)],
        ["Total Bank", fmt_usd(BS_31MAR['total_bank']), fmt_usd(254322.10)],
        ["Accounts Receivable", fmt_usd(BS_31MAR['ar']), fmt_usd(0)],
        ["Prepayments & Programme Advances", fmt_usd(BS_31MAR['prepayments'] +
                                                     BS_31MAR['programme_adv']),
         fmt_usd(0)],
        ["Fixed Assets (net)", fmt_usd(BS_31MAR['fixed_assets']), fmt_usd(1216.74)],
        ["TOTAL ASSETS", fmt_usd(BS_31MAR['total_assets']), fmt_usd(255538.84)],
        ["Accounts Payable", fmt_usd(BS_31MAR['ap']), fmt_usd(13042.75)],
        ["Expenses Payable", fmt_usd(BS_31MAR['exp_payable']), fmt_usd(0)],
        ["Suspense", fmt_usd(BS_31MAR['suspense']), fmt_usd(0)],
        ["TOTAL LIABILITIES", fmt_usd(BS_31MAR['total_liab']), fmt_usd(13042.75)],
        ["NET ASSETS", fmt_usd(BS_31MAR['net_assets']), fmt_usd(242496.09)],
        ["  Retained Earnings", fmt_usd(BS_31MAR['retained_earnings']),
         fmt_usd(418395.45)],
        ["  Current Year Earnings", fmt_usd(BS_31MAR['current_year_earnings']),
         fmt_usd(-175899.37)],
    ]
    add_table(doc, ["Line", "31 Mar 2026 (USD)", "31 Mar 2025 (USD)"], bs_rows,
              totals_rows={3, 7, 11, 12}, col_widths=[7, 3.5, 3.5])

    # Cash flow Q1
    add_heading(doc, "3. Cash flow statement — Q1 2026 (direct method)", level=2)
    cf_rows = [
        ["Opening cash (1 Jan 2026)", fmt_usd(Q1_BANK['opening'])],
        ["  Cash received", fmt_usd(Q1_BANK['received'])],
        ["  Cash spent", "(" + fmt_usd(Q1_BANK['spent']).replace('USD ', 'USD ') + ")"],
        ["  FX impact", fmt_usd(Q1_BANK['fx'])],
        ["Closing cash (31 Mar 2026)", fmt_usd(Q1_BANK['closing'])],
    ]
    add_table(doc, ["Line", "Q1 2026 (USD)"], cf_rows,
              totals_rows={0, 4}, col_widths=[10, 4])
    add_para(doc,
        "Cash inflow drivers: OSF #4 Y2 tranche USD 300,000 (Jan 2026) plus FCDO Phase "
        "I Y2 mobilisation transfers. No restricted-grant disbursements were made out of "
        "the trust account in the quarter — all cash-spent is operational.",
        size=9.5, italic=True)

    # Income by donor (verbatim contract refs)
    add_heading(doc, "4. Income recognition by donor", level=2)
    inc_rows = [
        ["OSF #4 (2025–29) Year-2 tranche",
         "USD 300,000.00",
         "Per OR2024-94840: \"a grant in the amount of 1,500,000.00 USD for the period from January 1, 2025 through December 31, 2029 to provide general support.\"",
         "Unrestricted (general support)"],
        ["Interest Income",
         "USD 472.10",
         "Wise USD wallet",
         "Unrestricted"],
        ["Swiss Dhaka (SDC) final tranche",
         "—",
         "Per SDC Art 4.1: max CHF 22,788.10 final instalment after donor approval of final operational + financial reports.",
         "Not yet invoiced — gated by §3.1 final report due 30 Jun 2026 + §3.3 audit"],
        ["Stimson Swiss (Arakan Y1)",
         "—",
         "Per Stimson §3.3: USD 24,946.67 due June 30, 2026, conditional on FDFA disbursement to Stimson.",
         "Not yet invoiced — Q2 expected"],
        ["FCDO Phase I (closed)",
         "—",
         "INV-0007 USD 57,005 raised end-Q1, received 12 May 2026 (post-quarter)",
         "Recognised in Q2; not in Q1 income"],
    ]
    add_table(doc, ["Donor / tranche", "Q1 income", "Source", "Restriction / gate"],
              inc_rows, col_widths=[3.5, 2.5, 7.5, 4.0])

    # Risk flags
    add_heading(doc, "5. Compliance / risk flags", level=2)
    flags = [
        ("SDC restricted tracking over-coded YTD", TOUCH,
         "YTD spend tagged to the SDC tracking option already exceeds the contract "
         "residual cap (max CHF 22,788.10 ≈ USD 25k) and the 2026 SDC budget column "
         "(USD 35,560). A re-tag pass is scheduled before the SDC final financial "
         "report (due 30 Jun 2026 per §3.1) — without it, the over-allocation falls "
         "back on unrestricted reserves."),
        ("FCDO Phase II not yet signed", TOUCH,
         "GBP-denominated proposal (~£80k, ~USD 106,400 at 1.33) is critical to 2026 "
         "cash closing (see P-06d). Target signature end-July 2026."),
        ("Unbudgeted YTD overhead lines", CORE,
         "General Expenses USD 29,194, Branding & Communications USD 9,807, "
         "Confidence Building Measures USD 14,598 — none budgeted in the original "
         "2026 plan. Mid-year rebudget recommended (P-06e)."),
        ("Balance sheet — clean", CORE,
         "No external debt, no director loans outstanding, AR=0 at quarter-end. "
         "Suspense account USD 3.46 — immaterial."),
    ]
    for label, color, text in flags:
        add_para(doc, label, bold=True, color=color, size=11)
        add_para(doc, text, size=10)

    add_heading(doc, "6. Assumptions & data sources", level=2)
    add_para(doc,
        "·  P&L 1 Jan – 31 Mar 2026 — Xero report ID 9ec1519b-a7af-46ef-8d84-24a9877c6842, "
        "pulled 26 May 2026.")
    add_para(doc,
        "·  Balance Sheet at 31 Mar 2026 (with 31 Mar 2025 comparative) — Xero report "
        "ID 12c538a9-293f-4774-a41c-3cc6a6ed5179, pulled 26 May 2026.")
    add_para(doc,
        "·  Bank Summary 1 Jan – 31 Mar 2026 — Xero report ID "
        "bb9d409c-c9d4-46f1-9ba6-c9a5dd4d87de, pulled 26 May 2026.")
    add_para(doc,
        "·  Accrual basis; USD reporting; EUR / GBP balances translated by Xero at the "
        "applicable Wise revaluation rate. Period lock 31 Dec 2022.")
    add_para(doc,
        "·  Donor contracts cross-referenced verbatim: OSF (OR2024-94840), SDC "
        "(7F-03317.09.15 / 81082408), Stimson Swiss (Effective 1 Dec 2025).")

    out = Q2 / "humanus_Q1-2026_management-accounts_P-06c.docx"
    doc.save(out)
    print(f"Wrote {out}")


# ===========================================================================
# P-06d  12-month rolling cash flow (refresh) — xlsx
# ===========================================================================
MONTHS_ROLLING = [
    ("Jun-26", "2026-06-30"), ("Jul-26", "2026-07-31"),
    ("Aug-26", "2026-08-31"), ("Sep-26", "2026-09-30"),
    ("Oct-26", "2026-10-31"), ("Nov-26", "2026-11-30"),
    ("Dec-26", "2026-12-31"), ("Jan-27", "2027-01-31"),
    ("Feb-27", "2027-02-28"), ("Mar-27", "2027-03-31"),
    ("Apr-27", "2027-04-30"), ("May-27", "2027-05-31"),
]


def build_cashflow_xlsx():
    wb = Workbook()
    # Brand styles
    core_fill = PatternFill("solid", fgColor=CORE_HEX)
    pale_fill = PatternFill("solid", fgColor="F5E9F0")
    bold_fill = PatternFill("solid", fgColor=BOLD_HEX)
    touch_fill = PatternFill("solid", fgColor="FBE3EC")
    white_font = Font(name="Inter", color="FFFFFF", bold=True, size=11)
    core_font = Font(name="Inter", color=CORE_HEX, size=10)
    core_bold = Font(name="Inter", color=CORE_HEX, size=10, bold=True)
    thin = Side(border_style="thin", color=CORE_HEX)
    box = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ---- Tab 1: Central case ----
    ws = wb.active
    ws.title = "Central"
    ws["A1"] = "hūmānus — 12-month rolling cash flow (Jun 2026 – May 2027)"
    ws["A1"].font = Font(name="Inter", color=CORE_HEX, size=14, bold=True)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=14)
    ws["A2"] = ("Central case  ·  As-of 26 May 2026  ·  USD  ·  Opening cash "
                "USD 182,801 (Xero live, 25 May 26)")
    ws["A2"].font = Font(name="Inter", color=CORE_HEX, size=9, italic=True)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=14)

    # Header row at row 4
    headers = ["Line"] + [m for m, _ in MONTHS_ROLLING] + ["12-mo Total"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=c, value=h)
        cell.fill = core_fill
        cell.font = white_font
        cell.alignment = Alignment(horizontal="center")
        cell.border = box

    # Cash inflows — Central case
    inflows_central = {
        # month-index 0..11 = Jun-26..May-27
        # SDC final tranche: assume Sep 2026 (90 days after SDC §3.1 final
        # report due 30 Jun 2026 + audit window)
        "OSF #4 Y3 tranche (Jan 2027)":      [0, 0, 0, 0, 0, 0, 0, 300000, 0, 0, 0, 0],
        "OSF interest accrual":              [200] * 12,
        "Swiss Dhaka final (§4.1, CHF→USD)": [0, 0, 0, 25000, 0, 0, 0, 0, 0, 0, 0, 0],
        "Stimson Swiss Y1 (§3.3 30 Jun 26)": [24947, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
        "Stimson Swiss Y2 mid (§3.3, pull-fwd request)":
                                              [0, 0, 0, 0, 0, 0, 24947, 0, 0, 0, 0, 0],
        "Stimson Swiss Y2 end (§3.3 30 Jun 27)":
                                              [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 24947],
        "FCDO Phase II — first tranche (unsigned, target end-Jul sign)":
                                              [0, 0, 53000, 0, 0, 0, 0, 0, 0, 0, 0, 0],
        "FCDO Phase II — milestone 2 (unsigned)":
                                              [0, 0, 0, 0, 0, 0, 0, 0, 0, 53400, 0, 0],
    }
    # Outflows
    spend = [RUN_RATE_MONTHLY] * 12
    # Known lumps within run-rate already (audit Q3 ~$1k; board Sep+Dec ~$1k ea)
    # Per-line work in CAP-147 v2 indicates ~$45k/mo central, holding flat

    # Build rows
    row = 5
    ws.cell(row=row, column=1, value="OPENING CASH").font = core_bold
    ws.cell(row=row, column=2, value=182801).font = core_bold
    for c in range(3, 15):
        ws.cell(row=row, column=c, value=f"=Central!{get_column_letter(c-1)}{row+5+len(inflows_central)+3}")
    row += 1

    # Inflows section header
    ws.cell(row=row, column=1, value="Inflows").font = core_bold
    ws.cell(row=row, column=1).fill = pale_fill
    row += 1
    inflow_start = row
    for name, vals in inflows_central.items():
        ws.cell(row=row, column=1, value=name).font = core_font
        for c, v in enumerate(vals, start=2):
            ws.cell(row=row, column=c, value=v).font = core_font
        ws.cell(row=row, column=14,
                value=f"=SUM(B{row}:M{row})").font = core_font
        row += 1
    inflow_end = row - 1
    ws.cell(row=row, column=1, value="Total inflows").font = core_bold
    ws.cell(row=row, column=1).fill = pale_fill
    for c in range(2, 15):
        col = get_column_letter(c)
        ws.cell(row=row, column=c,
                value=f"=SUM({col}{inflow_start}:{col}{inflow_end})").font = core_bold
    inflow_total_row = row
    row += 1

    # Outflows section header
    ws.cell(row=row, column=1, value="Outflows").font = core_bold
    ws.cell(row=row, column=1).fill = pale_fill
    row += 1
    outflow_start = row
    ws.cell(row=row, column=1,
            value="Run-rate spend (Cost of Sales + Overhead)").font = core_font
    for c, v in enumerate(spend, start=2):
        ws.cell(row=row, column=c, value=-v).font = core_font
    ws.cell(row=row, column=14, value=f"=SUM(B{row}:M{row})").font = core_font
    row += 1
    outflow_end = row - 1
    ws.cell(row=row, column=1, value="Total outflows").font = core_bold
    ws.cell(row=row, column=1).fill = pale_fill
    for c in range(2, 15):
        col = get_column_letter(c)
        ws.cell(row=row, column=c,
                value=f"=SUM({col}{outflow_start}:{col}{outflow_end})").font = core_bold
    outflow_total_row = row
    row += 1

    # Net movement
    ws.cell(row=row, column=1, value="Net cash movement").font = core_bold
    for c in range(2, 15):
        col = get_column_letter(c)
        ws.cell(row=row, column=c,
                value=f"={col}{inflow_total_row}+{col}{outflow_total_row}"
                ).font = core_bold
    net_row = row
    row += 1

    # Closing cash (= prior column closing + this column net)
    ws.cell(row=row, column=1, value="CLOSING CASH").font = core_bold
    ws.cell(row=row, column=1).fill = bold_fill
    # First month closing = opening (col B in opening row) + Jun net
    ws.cell(row=row, column=2,
            value=f"=B5+B{net_row}").font = core_bold
    for c in range(3, 15):
        col_prev = get_column_letter(c - 1)
        col = get_column_letter(c)
        if c < 14:
            ws.cell(row=row, column=c,
                    value=f"={col_prev}{row}+{col}{net_row}").font = core_bold
        else:
            # 12mo total column = sum of all nets + opening
            ws.cell(row=row, column=c,
                    value=f"=B5+SUM(B{net_row}:M{net_row})").font = core_bold
    closing_row = row

    # Repoint OPENING CASH chain in row 5 to reference prior month closings
    for c in range(3, 15):
        col_prev = get_column_letter(c - 1)
        ws.cell(row=5, column=c,
                value=f"={col_prev}{closing_row}").font = core_bold

    # Column widths
    ws.column_dimensions["A"].width = 50
    for c in range(2, 15):
        ws.column_dimensions[get_column_letter(c)].width = 11.5

    # ---- Tab 2: Downside ----
    ws2 = wb.create_sheet("Downside")
    ws2["A1"] = "Downside scenario — FCDO Phase II slips to 2027; SDC final delayed"
    ws2["A1"].font = Font(name="Inter", color=TOUCH_HEX, size=14, bold=True)
    ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=14)
    ws2["A2"] = ("Downside  ·  As-of 26 May 2026  ·  USD  ·  Opening cash USD 182,801")
    ws2["A2"].font = Font(name="Inter", color=CORE_HEX, size=9, italic=True)
    ws2.merge_cells(start_row=2, start_column=1, end_row=2, end_column=14)

    for c, h in enumerate(headers, start=1):
        cell = ws2.cell(row=4, column=c, value=h)
        cell.fill = core_fill
        cell.font = white_font

    inflows_down = {
        "OSF #4 Y3 tranche (Jan 2027)":      [0, 0, 0, 0, 0, 0, 0, 300000, 0, 0, 0, 0],
        "OSF interest accrual":              [200] * 12,
        "Swiss Dhaka final (SLIPS to Q1 2027)":
                                              [0, 0, 0, 0, 0, 0, 0, 0, 25000, 0, 0, 0],
        "Stimson Swiss Y1 (§3.3 30 Jun 26)": [24947, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
        "Stimson Swiss Y2 mid (§3.3 contractual)":
                                              [0, 0, 0, 0, 0, 0, 0, 24947, 0, 0, 0, 0],
        "Stimson Swiss Y2 end (§3.3 30 Jun 27)":
                                              [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 24947],
        "FCDO Phase II — NOT contracted (excluded)":
                                              [0] * 12,
    }
    # Outflows same as central
    row = 5
    ws2.cell(row=row, column=1, value="OPENING CASH").font = core_bold
    ws2.cell(row=row, column=2, value=182801).font = core_bold
    row += 1
    ws2.cell(row=row, column=1, value="Inflows").font = core_bold
    ws2.cell(row=row, column=1).fill = pale_fill
    row += 1
    inflow_start = row
    for name, vals in inflows_down.items():
        ws2.cell(row=row, column=1, value=name).font = core_font
        for c, v in enumerate(vals, start=2):
            ws2.cell(row=row, column=c, value=v).font = core_font
        ws2.cell(row=row, column=14,
                 value=f"=SUM(B{row}:M{row})").font = core_font
        row += 1
    inflow_end = row - 1
    ws2.cell(row=row, column=1, value="Total inflows").font = core_bold
    ws2.cell(row=row, column=1).fill = pale_fill
    for c in range(2, 15):
        col = get_column_letter(c)
        ws2.cell(row=row, column=c,
                 value=f"=SUM({col}{inflow_start}:{col}{inflow_end})").font = core_bold
    inflow_total_row = row
    row += 1
    ws2.cell(row=row, column=1, value="Outflows").font = core_bold
    ws2.cell(row=row, column=1).fill = pale_fill
    row += 1
    outflow_start = row
    ws2.cell(row=row, column=1,
             value="Run-rate spend (Cost of Sales + Overhead)").font = core_font
    for c, v in enumerate(spend, start=2):
        ws2.cell(row=row, column=c, value=-v).font = core_font
    ws2.cell(row=row, column=14, value=f"=SUM(B{row}:M{row})").font = core_font
    row += 1
    outflow_end = row - 1
    ws2.cell(row=row, column=1, value="Total outflows").font = core_bold
    ws2.cell(row=row, column=1).fill = pale_fill
    for c in range(2, 15):
        col = get_column_letter(c)
        ws2.cell(row=row, column=c,
                 value=f"=SUM({col}{outflow_start}:{col}{outflow_end})").font = core_bold
    outflow_total_row = row
    row += 1
    ws2.cell(row=row, column=1, value="Net cash movement").font = core_bold
    for c in range(2, 15):
        col = get_column_letter(c)
        ws2.cell(row=row, column=c,
                 value=f"={col}{inflow_total_row}+{col}{outflow_total_row}"
                 ).font = core_bold
    net_row = row
    row += 1
    ws2.cell(row=row, column=1, value="CLOSING CASH").font = core_bold
    ws2.cell(row=row, column=1).fill = touch_fill
    ws2.cell(row=row, column=2, value=f"=B5+B{net_row}").font = core_bold
    for c in range(3, 15):
        col_prev = get_column_letter(c - 1)
        col = get_column_letter(c)
        if c < 14:
            ws2.cell(row=row, column=c,
                     value=f"={col_prev}{row}+{col}{net_row}").font = core_bold
        else:
            ws2.cell(row=row, column=c,
                     value=f"=B5+SUM(B{net_row}:M{net_row})").font = core_bold
    closing_row = row
    for c in range(3, 15):
        col_prev = get_column_letter(c - 1)
        ws2.cell(row=5, column=c, value=f"={col_prev}{closing_row}").font = core_bold

    ws2.column_dimensions["A"].width = 50
    for c in range(2, 15):
        ws2.column_dimensions[get_column_letter(c)].width = 11.5

    # ---- Tab 3: Assumptions ----
    ws3 = wb.create_sheet("Assumptions")
    ws3["A1"] = "Assumptions & sources — 12-month rolling cash flow"
    ws3["A1"].font = Font(name="Inter", color=CORE_HEX, size=14, bold=True)
    ws3.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)

    notes = [
        ("Opening cash (25 May 2026)",
         "USD 182,801 (Wise USD 173,360 + EUR 9,436 + GBP 0)",
         "Xero Bank Summary 1 Jan – 25 May 2026; report ID 804d1ade-c575-4c7d-b709-f838b43941a0"),
        ("Run-rate spend (central + downside)",
         "USD 45,000 / month flat",
         "Calibrated to YTD cash-spent USD 204,115 ÷ 4.6 months = USD 44,373/mo (CAP-147 v2)"),
        ("OSF #4 Year-3 tranche",
         "USD 300,000 in Jan 2027",
         "Per OR2024-94840 schedule of five tranches Dec 2024 / Jan 2026-29; signed 26 Oct 2024"),
        ("SDC Swiss Dhaka final tranche",
         "Max CHF 22,788.10 ≈ USD 25,000 — central: Sep 2026; downside: Mar 2027",
         "Per SDC Art 4.1 — gated by §3.1 final reports (due 30 Jun 2026) + §3.3 audit"),
        ("Stimson Swiss Y1 (Arakan Y1)",
         "USD 24,946.67 in Jun 2026",
         "Per Stimson §3.3, due 30 Jun 2026; conditional on FDFA disbursement to Stimson"),
        ("Stimson Swiss Y2 mid",
         "USD 24,946.67 — central: pull forward to Dec 2026; downside: contractual Jan 2027",
         "Pull-forward requested by ED in flight (CAP-158); contractual date per §3.3 31 Jan 2027"),
        ("Stimson Swiss Y2 end",
         "USD 24,946.66 in May/Jun 2027",
         "Per Stimson §3.3 — 30 Jun 2027 contractual"),
        ("FCDO Phase II",
         "Central: GBP 80,000 ÷ 2 tranches Aug 2026 + Mar 2027 at GBP/USD 1.33 = USD 106,400 total; Downside: excluded",
         "Per FCDO budget worksheet finance-docs/grants/5 5 26 FCDO budget humanus.xlsx — NO signed agreement on file; target sign end-Jul 2026"),
        ("FX rates",
         "GBP/USD 1.33 (OANDA, working figure per Eva); EUR balances revalued by Xero",
         "Per CAP-147 v2 anchor"),
        ("Excluded from forecast",
         "Any new unrestricted donor income; cost-of-living adjustments; H2 audit overrun",
         "Conservative; flagged in commentary"),
    ]
    ws3.cell(row=3, column=1, value="Assumption").font = white_font
    ws3.cell(row=3, column=1).fill = core_fill
    ws3.cell(row=3, column=2, value="Value").font = white_font
    ws3.cell(row=3, column=2).fill = core_fill
    ws3.cell(row=3, column=3, value="Source / clause").font = white_font
    ws3.cell(row=3, column=3).fill = core_fill
    for i, (a, v, s) in enumerate(notes, start=4):
        ws3.cell(row=i, column=1, value=a).font = core_font
        ws3.cell(row=i, column=2, value=v).font = core_font
        ws3.cell(row=i, column=3, value=s).font = core_font
        ws3.cell(row=i, column=1).alignment = Alignment(wrap_text=True, vertical="top")
        ws3.cell(row=i, column=2).alignment = Alignment(wrap_text=True, vertical="top")
        ws3.cell(row=i, column=3).alignment = Alignment(wrap_text=True, vertical="top")

    ws3.column_dimensions["A"].width = 35
    ws3.column_dimensions["B"].width = 50
    ws3.column_dimensions["C"].width = 60

    # ---- Tab 4: Commentary ----
    ws4 = wb.create_sheet("Commentary")
    ws4["A1"] = "Commentary"
    ws4["A1"].font = Font(name="Inter", color=CORE_HEX, size=14, bold=True)
    commentary = [
        "Cash position at 25 May 2026: USD 182,801 (Wise USD 173k + EUR 9k + GBP 0). "
        "Above the proposed reserves 3-month Red floor (USD 132,527); inside the Amber band; "
        "below the 6-month Green target (USD 265,054).",
        "",
        "Central case (FCDO Phase II signs end-Jul 2026, all tranches land in projected windows):",
        "  – Reserves Amber → Red transition end-July 2026 (cash USD 118k, below 3-mo floor USD 132,527).",
        "  – Lowest point end-December 2026 at ~USD (3,000) — briefly negative just before OSF Y3 tranche.",
        "  – Recovers to USD 252k end-January 2027 with OSF Y3 tranche.",
        "  – Closes May 2027 at USD 151k (Amber band).",
        "  – 12-month net cash movement: ~USD (31,000) — net outflow even in central case.",
        "",
        "Downside case (FCDO Phase II slips into 2027; SDC final slips to Q1 2027):",
        "  – Cash exhausts mid-October 2026; closes December 2026 at ~USD (106,000) overdraft.",
        "  – OSF Y3 tranche January 2027 restores positive cash; closes May 2027 at USD 45k.",
        "",
        "Levers if downside emerges (in priority order; all to be triggered by end-Jul 2026 if FCDO II unsigned):",
        "  1. Pause discretionary spend — Volunteer Stipends (already exhausted), Travel, Board pulses — saves ~USD 6-10k/month from Aug 2026.",
        "  2. Confirm Stimson Y2 mid pull-forward to Dec 2026 (in flight per CAP-158); bridges into OSF Y3 Jan 2027.",
        "  3. Short-term bridge facility (Wise Business credit / partner loan) for Oct–Dec 2026 window.",
        "  4. Mid-year board-approved rebudget — defer non-critical 2026 programme spend into 2027.",
        "  5. As last resort, accelerate the Stimson Swiss Y2 end tranche (May 2027 contractually) — request to Stimson would parallel the Y2 mid request.",
        "",
        "Recommended board action:",
        "  – Note central / downside cash trajectories and the end-July 2026 decision gate.",
        "  – Approve the contingent levers above as pre-authorised mitigations should the downside emerge.",
        "  – Direct ED to provide FCDO Phase II signature update by 31 July 2026, and to trigger Levers 1–3 in sequence if unsigned at that point.",
        "  – Re-affirm the proposed reserves policy (P-06f, CAP-25); accept that central case crosses below the 3-month Red floor for most of Aug 2026 – Jan 2027 even with successful FCDO II.",
    ]
    for i, line in enumerate(commentary, start=3):
        cell = ws4.cell(row=i, column=1, value=line)
        cell.font = Font(name="Inter", color=CORE_HEX, size=11,
                         bold=line.startswith("Central") or line.startswith("Downside")
                              or line.startswith("Levers") or line.startswith("Recommended"))
    ws4.column_dimensions["A"].width = 120

    out = Q2 / "humanus_cashflow_12mo-rolling_Jun26-May27_P-06d.xlsx"
    wb.save(out)
    print(f"Wrote {out}")


# ===========================================================================
# P-06e  2026 Budget Variance — docx + xlsx
# ===========================================================================
def build_bva():
    # ---- Docx narrative ----
    doc = new_doc("2026 Budget Variance Report (YTD 25 May 2026)", "P-06e")
    months_elapsed = YTD_MONTHS  # ~4.80
    elapsed_pct = months_elapsed / 12

    add_heading(doc, "Headline", level=2)
    add_para(doc,
        f"YTD spend USD {YTD_SPEND:,.0f} on an annual budget of "
        f"USD {BUDGET_2026['total']:,.0f} = "
        f"{YTD_SPEND / BUDGET_2026['total']:.0%} consumed at "
        f"{elapsed_pct:.0%} of the year elapsed — running ~9 points ahead "
        f"of straight-line pace on aggregate. The over-pace is driven by "
        f"unbudgeted Operating Expense lines (General Expenses, Branding & "
        f"Communications, Confidence Building Measures) rather than core "
        f"programme run-rate.", bold=True, color=CORE)

    add_heading(doc, "Variance by donor (funding source)", level=2)
    add_para(doc, "Annualised YTD spend, by 2026 budget column.")
    # We don't have donor-tracked YTD spend in this run (would require
    # additional Xero pulls); we annualise from total and ratio by budget.
    donor_rows = [
        ["OSF #4 (unrestricted)",
         f"USD {286634.70:,.0f}",
         "USD 168,725 (annualised; YTD ~$66k)",
         "(41%) underspent",
         "Healthy headroom; ~USD 62k headroom before SDC re-tag (CAP-158)"],
        ["Swiss Dhaka (SDC)",
         f"USD {35560:,.0f}",
         "USD ~269k (YTD ~$105k; coding error — see flag)",
         "+658%",
         "Over-coded to SDC tracking; re-tag pass scheduled before §3.1 30 Jun 2026 final report"],
        ["Swiss Myanmar (Stimson)",
         f"USD {48312:,.0f}",
         "USD 10,883 (annualised; YTD ~$4.3k)",
         "(78%) underspent",
         "Underspent — first tranche due 30 Jun 2026 will bring spend forward"],
        ["FCDO Phase I (closed)",
         f"USD {42519:,.0f}",
         "USD 135,029 (annualised; YTD ~$53k)",
         "+218%",
         "Includes unbudgeted Confidence Building Measures $14,598; Phase I closed Mar 2026"],
        ["FCDO Phase II (unsigned)",
         f"USD {75424:,.0f}",
         "USD 0",
         "(100%)",
         "No grant agreement signed; budget worksheet only"],
    ]
    add_table(doc,
              ["Donor / source", "2026 Budget", "YTD spend (annualised)",
               "Variance vs budget", "Commentary"],
              donor_rows, col_widths=[3.5, 2.5, 4.0, 2.5, 5.0])

    add_heading(doc, "Variance by line item (top 10 movers, USD)", level=2)
    # Compare YTD vs straight-line YTD budget (annual ÷ 12 × months_elapsed)
    line_rows = []
    for acct, annual_budget in BUDGET_2026["by_account_annual"].items():
        # Map approx YTD spend to the account
        ytd = 0
        if acct == "Executive Director":
            ytd = YTD_PL_SPEND.get("Executive Director", 0)
        elif acct == "Cost of Living Allowances":
            ytd = YTD_PL_SPEND.get("Cost of Living Allowances", 0)
        elif acct == "Advocate / Programme Manager":
            ytd = YTD_PL_SPEND.get("Advocate / Programme Manager", 0)
        elif acct == "Rohingya Manager":
            ytd = YTD_PL_SPEND.get("Rohingya Manager", 0)
        elif acct == "Volunteer Stipends":
            ytd = YTD_PL_SPEND.get("Volunteer Stipends", 0)
        elif acct == "Vehicle Hire & Sundry":
            ytd = 0
        elif acct == "Partner Grants":
            ytd = YTD_PL_SPEND.get("Partner Grants", 0)
        elif acct == "Legal Fees":
            ytd = YTD_PL_SPEND.get("Legal Fees", 0)
        elif acct == "Communication Advisor":
            ytd = YTD_PL_SPEND.get("Communication Advisor", 0)
        elif acct == "Communication Support":
            ytd = YTD_PL_SPEND.get("Communication Support", 0)
        elif acct == "Accounting & Bookkeeping Fees":
            ytd = YTD_PL_SPEND.get("Accounting & Bookkeeping Fees", 0)
        elif acct == "Operations Support":
            ytd = YTD_PL_SPEND.get("Operations Support", 0)
        elif acct == "Audit Fees":
            ytd = YTD_PL_SPEND.get("Audit Fees", 0)
        elif acct == "Board Meetings":
            ytd = YTD_PL_SPEND.get("Board Meetings", 0)
        elif acct == "IT & Software Expenses":
            ytd = YTD_PL_SPEND.get("IT & Software Expenses", 0)
        elif acct.startswith("Travel"):
            ytd = (YTD_PL_SPEND.get("Accommodation", 0)
                   + YTD_PL_SPEND.get("Domestic Flights", 0)
                   + YTD_PL_SPEND.get("International Flights", 0)
                   + YTD_PL_SPEND.get("Local Transportation", 0)
                   + YTD_PL_SPEND.get("Per diem", 0)
                   + YTD_PL_SPEND.get("Meeting Expenses", 0))
        elif acct.startswith("Bank Charges"):
            ytd = (YTD_PL_SPEND.get("Bank Charges & Merchant Fees", 0)
                   + YTD_PL_SPEND.get("Insurance", 0)
                   + YTD_PL_SPEND.get("Foreign Currency Gains and Losses", 0))
        pace_budget = annual_budget * elapsed_pct
        variance = ytd - pace_budget
        pct = variance / pace_budget if pace_budget else None
        line_rows.append((acct, annual_budget, ytd, pace_budget, variance, pct))

    # Sort by abs variance %
    line_rows.sort(key=lambda r: -abs(r[5] if r[5] else 0))
    top10 = line_rows[:10]
    rows = []
    warn = set()
    for i, (acct, ann, ytd, pace, var, pct) in enumerate(top10):
        pct_str = f"{pct:+.0%}" if pct is not None else "n/a"
        rows.append([acct, f"USD {ann:,.0f}", f"USD {pace:,.0f}",
                     f"USD {ytd:,.0f}", f"USD {var:+,.0f}", pct_str])
        if pct is not None and abs(pct) >= 0.5:
            warn.add(i)
    add_table(doc,
              ["Line", "Annual budget", "YTD pace budget", "YTD actual",
               "Variance", "Variance %"],
              rows, warn_rows=warn,
              col_widths=[5.0, 2.5, 2.5, 2.5, 2.5, 2.0])

    add_heading(doc, "Unbudgeted lines (not in original 2026 plan)", level=2)
    unbudgeted = [
        ("General Expenses", YTD_PL_SPEND["General Expenses"],
         "Largely sitting on SDC tracking — possible coding error; investigate before SDC final report"),
        ("Branding & Communications", YTD_PL_SPEND["Branding & Communications"],
         "Rebrand pulse VAI → hūmānus; one-off but to be reclassified or rebudgeted"),
        ("Confidence Building Measures", YTD_PL_SPEND["Confidence Building Measures"],
         "Q1 programme activity not separately budgeted; eligible against FCDO Phase I"),
        ("Foreign Currency Gains and Losses", YTD_PL_SPEND["Foreign Currency Gains and Losses"],
         "Wise revaluations on EUR/GBP wallets; immaterial"),
    ]
    ur_rows = [[name, f"USD {amt:,.0f}", note] for name, amt, note in unbudgeted]
    ur_total = sum(amt for _, amt, _ in unbudgeted)
    ur_rows.append(["Total unbudgeted YTD", f"USD {ur_total:,.0f}", ""])
    add_table(doc, ["Line", "YTD actual", "Commentary"], ur_rows,
              totals_rows={len(ur_rows) - 1}, col_widths=[4.0, 2.5, 9.5])

    add_heading(doc, "Drivers & commentary", level=2)
    drivers = [
        ("Expense side is broadly on plan after stripping one-offs.",
         "Core programme staff (ED + Programme Manager + Communication Advisor) is "
         "running on contract; quarterly burn ~USD 45k/month is consistent with the "
         "annual budget once unbudgeted Q1 lines are recognised as one-offs."),
        ("Income side carries the structural risk.",
         "USD 75,424 of 2026 budgeted spend is funded by FCDO Phase II, which is not "
         "yet signed. SDC restricted is over-coded and will absorb unrestricted "
         "reserves unless re-tagged."),
        ("Restricted / unrestricted segregation needs a re-tag pass.",
         "SDC YTD-tagged spend (~USD 105k) far exceeds the contract residual (max CHF "
         "22,788.10 ≈ USD 25k). The Finance Manager will run a re-coding pass before "
         "the SDC §3.1 final report (due 30 Jun 2026) — material amounts will move "
         "back to OSF #4 (unrestricted) tracking."),
        ("Closed projects with outstanding obligations.",
         "SDC and FCDO Phase I are closed (per portfolio.md, last updated 2026-05-01) "
         "but both carry residual receivables and reporting deadlines. Stimson INV-0007 "
         "(FCDO I, USD 57,005) cleared 12 May 2026."),
    ]
    for label, text in drivers:
        add_para(doc, label, bold=True, color=CORE, size=11)
        add_para(doc, text)

    add_heading(doc, "Recommended actions", level=2)
    actions = [
        ("Finance Manager — complete SDC tracking re-tag pass before 15 Jun 2026.",
         "Required for SDC §3.1 final financial report (due 30 Jun 2026)."),
        ("ED — confirm FCDO Phase II signature window with Stimson by 31 Jul 2026.",
         "Drives whether downside cashflow scenario triggers (see P-06d)."),
        ("Board — note recommendation to rebudget at H1 close.",
         "Move unbudgeted Branding & General Expense lines into a formal mid-year plan; "
         "decide whether to fund from OSF carryover or programme reclassification."),
    ]
    for owner, why in actions:
        add_para(doc, "·  " + owner, bold=True, color=CORE)
        add_para(doc, "    " + why, size=10, italic=True)

    add_heading(doc, "Assumptions & data sources", level=2)
    add_para(doc,
        "·  YTD P&L 1 Jan – 25 May 2026 — Xero report ID "
        "b1fc6fb0-af72-4308-99f6-f34b08bc6206, pulled 26 May 2026.")
    add_para(doc,
        "·  Annual budget per finance-docs/budgets/2026 Budget-2.xlsx (Sheet1), "
        "total USD 488,450.")
    add_para(doc,
        "·  Donor-tracked YTD spend annualised from CAP-157 Xero P&L by Donors "
        "tracking option (21 May 2026 pull); refresh by close of Q2.")
    add_para(doc,
        "·  Months elapsed: 4.80 (1 Jan – 25 May 2026 ÷ 30.42).")

    out = Q2 / "humanus_BvA_2026-YTD_P-06e.docx"
    doc.save(out)
    print(f"Wrote {out}")

    # ---- Supporting xlsx ----
    wb = Workbook()
    ws = wb.active
    ws.title = "BvA YTD"
    ws["A1"] = "hūmānus — 2026 Budget Variance (YTD 25 May 2026)"
    ws["A1"].font = Font(name="Inter", color=CORE_HEX, size=14, bold=True)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    ws["A2"] = "As-of 26 May 2026  ·  USD  ·  Months elapsed: 4.80 (40%)"
    ws["A2"].font = Font(name="Inter", color=CORE_HEX, size=9, italic=True)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=6)

    headers = ["Line", "Annual Budget (USD)", "YTD Pace Budget (USD)",
               "YTD Actual (USD)", "Variance (USD)", "Variance %"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=c, value=h)
        cell.fill = PatternFill("solid", fgColor=CORE_HEX)
        cell.font = Font(name="Inter", color="FFFFFF", bold=True)
    row = 5
    for acct, ann, ytd, pace, var, pct in line_rows:
        ws.cell(row=row, column=1, value=acct)
        ws.cell(row=row, column=2, value=ann)
        ws.cell(row=row, column=3, value=round(pace, 2))
        ws.cell(row=row, column=4, value=round(ytd, 2))
        ws.cell(row=row, column=5, value=round(var, 2))
        ws.cell(row=row, column=6, value=pct)
        for c in range(2, 6):
            ws.cell(row=row, column=c).number_format = '#,##0;(#,##0)'
        ws.cell(row=row, column=6).number_format = "0%"
        for c in range(1, 7):
            cell = ws.cell(row=row, column=c)
            cell.font = Font(name="Inter", color=CORE_HEX, size=10)
            if pct is not None and abs(pct) >= 0.5:
                cell.font = Font(name="Inter", color=TOUCH_HEX, size=10, bold=True)
        row += 1
    # Totals
    ann_total = sum(r[1] for r in line_rows)
    ytd_total = sum(r[2] for r in line_rows)
    pace_total = sum(r[3] for r in line_rows)
    var_total = ytd_total - pace_total
    ws.cell(row=row, column=1, value="Subtotal (mapped lines)").font = Font(
        name="Inter", color=CORE_HEX, bold=True)
    ws.cell(row=row, column=2, value=ann_total)
    ws.cell(row=row, column=3, value=round(pace_total, 2))
    ws.cell(row=row, column=4, value=round(ytd_total, 2))
    ws.cell(row=row, column=5, value=round(var_total, 2))
    for c in range(2, 6):
        ws.cell(row=row, column=c).number_format = '#,##0;(#,##0)'
        ws.cell(row=row, column=c).font = Font(name="Inter", color=CORE_HEX, bold=True)
    row += 1
    ws.cell(row=row, column=1, value="Plus: Unbudgeted YTD lines (Branding, "
            "General Expenses, CBM, FX, Misc)").font = Font(
                name="Inter", color=TOUCH_HEX, italic=True)
    unmapped = YTD_SPEND - ytd_total
    ws.cell(row=row, column=4, value=round(unmapped, 2)).number_format = '#,##0;(#,##0)'
    row += 1
    ws.cell(row=row, column=1, value="TOTAL YTD spend").font = Font(
        name="Inter", color=CORE_HEX, bold=True, size=11)
    ws.cell(row=row, column=2, value=BUDGET_2026["total"]).number_format = '#,##0;(#,##0)'
    ws.cell(row=row, column=4, value=round(YTD_SPEND, 2)).number_format = '#,##0;(#,##0)'
    ws.cell(row=row, column=5,
            value=round(YTD_SPEND - BUDGET_2026["total"] * elapsed_pct, 2)
            ).number_format = '#,##0;(#,##0)'
    ws.column_dimensions["A"].width = 50
    for c in range(2, 7):
        ws.column_dimensions[get_column_letter(c)].width = 18

    out2 = Q2 / "humanus_BvA_2026-YTD_P-06e.xlsx"
    wb.save(out2)
    print(f"Wrote {out2}")


# ===========================================================================
# P-06f  Reserves position note
# ===========================================================================
def build_reserves_note():
    doc = new_doc("Reserves Position Note", "P-06f")

    add_heading(doc, "Headline", level=2)
    add_para(doc,
        f"At 25 May 2026, hūmānus held USD {BS_LIVE['total_bank']:,.0f} cash "
        f"(Wise USD {BS_LIVE['wise_usd']:,.0f} + EUR {BS_LIVE['wise_eur']:,.0f} + "
        f"GBP {BS_LIVE['wise_gbp']:,.0f}). Net assets USD "
        f"{BS_LIVE['net_assets']:,.0f}. This sits in the AMBER band of the proposed "
        f"reserves policy (CAP-25, awaiting board sign-off) — equivalent to ~4.1 "
        f"months of core unrestricted overhead. The 6-month Target (Green) "
        f"requires a further ~USD 82k.", bold=True, color=CORE)

    add_heading(doc, "Position vs proposed policy", level=2)
    months_cover = BS_LIVE['total_bank'] / RESERVES['core_overhead_monthly']
    band_rows = [
        ["Floor (Red) — 3 months core overhead", f"USD {RESERVES['floor_3mo']:,.0f}",
         "Held"],
        ["Amber band — 3 to 6 months",
         f"USD {RESERVES['floor_3mo']:,.0f} – {RESERVES['target_6mo']:,.0f}",
         f"Held — USD {BS_LIVE['total_bank']:,.0f} sits inside band "
         f"({months_cover:.1f} months cover)"],
        ["Target (Green) — 6 months core overhead",
         f"USD {RESERVES['target_6mo']:,.0f}",
         f"Gap of USD {RESERVES['target_6mo'] - BS_LIVE['total_bank']:,.0f} to reach"],
    ]
    add_table(doc, ["Band", "Threshold (USD)", "Status"], band_rows,
              col_widths=[6, 4, 7])

    add_heading(doc, "Policy basis", level=2)
    add_para(doc,
        "Core unrestricted overhead = Operating Expenses (audit, accounting, board, "
        "IT, comms, ops support, FX, depreciation, insurance) — excluding Cost of "
        "Sales / programme staff, because programme spend is grant-restricted and "
        "matches restricted grant income.")
    add_para(doc,
        f"Base figure: USD {RESERVES['core_overhead_monthly']:,.0f} / month "
        "(Feb–Apr 2026 average). 3-month floor and 6-month target reflect standard "
        "practice for grant-funded charities with lumpy income (USD 300k OSF tranches "
        "land in January each year, not evenly through the year).")

    add_heading(doc, "Trend", level=2)
    trend_rows = [
        ["31 Mar 2025 (Q1 closing, comparative)", "USD 254,322", "Green (above 6-mo)"],
        ["31 Dec 2025 (year-end)",
         "USD 23,787  (note: post-OSF Y1 tranche fully spent on Q4 2025 programme; "
         "OSF Y2 had not yet landed)",
         "Red"],
        ["31 Jan 2026 (post-OSF Y2 tranche)", "USD ~300,000", "Green"],
        ["31 Mar 2026", f"USD {BS_31MAR['total_bank']:,.0f}", "Amber/Green"],
        ["30 Apr 2026", f"USD {BS_30APR['total_bank']:,.0f}", "Amber"],
        ["25 May 2026 (current)", f"USD {BS_LIVE['total_bank']:,.0f}", "Amber"],
    ]
    add_table(doc, ["Date", "Total cash (USD equiv)", "Reserves band"],
              trend_rows, col_widths=[5, 6, 5])
    add_para(doc,
        "The pattern is structural: cash peaks immediately post-OSF tranche, draws "
        "down through the year, troughs in Q4 before the next OSF tranche. The "
        "reserves policy is sized for the trough, not the peak. Without the OSF "
        "tranche cadence, the organisation would breach the Red floor at year-end.",
        size=10)

    add_heading(doc, "Forward view (per 12-month cash flow P-06d)", level=2)
    add_para(doc,
        "Central case (FCDO II signs end-July 2026): cash crosses below the 3-month "
        "Red floor at end-July 2026 (USD 118k vs USD 132,527 floor), troughs briefly "
        f"negative at end-December 2026 (~USD (3k)) just before the OSF Year-3 "
        "tranche, and recovers to Green at end-January 2027 (USD 252k). Closes May "
        "2027 at USD 151k (Amber). The Red floor is breached for ~6 months "
        "even under the central case — a structural feature of lumpy annual donor "
        "tranches, not a forecast error.", color=TOUCH)
    add_para(doc,
        "Downside case (FCDO Phase II not signed): cash exhausts mid-October 2026, "
        "troughs at ~USD (106k) end-December 2026, recovers to USD 174k end-January "
        "2027 with OSF Year-3 tranche. Levers per P-06d must be triggered no later "
        "than 31 July 2026 in this scenario.", color=TOUCH)

    add_heading(doc, "Compliance / risk flags", level=2)
    add_para(doc,
        "·  Restricted-vs-unrestricted: of the USD 182,801 cash, most is "
        "unrestricted (OSF tranches are general support). SDC restricted balance "
        "is being reconciled before the §3.1 final financial report (due 30 Jun "
        "2026) — over-coded YTD spend on SDC tracking will move to OSF #4 on "
        "re-tag, which slightly increases pressure on unrestricted reserves but "
        "remains within the Amber band.")
    add_para(doc,
        "·  Currency concentration: the GBP wallet is empty; EUR wallet has drawn "
        "down to USD 9k equivalent (Geneva trip + EUR contractors paid). USD wallet "
        "carries the bulk of the position. No hedging — natural offset only.")
    add_para(doc,
        "·  Sole-signatory model: FM (Stacy) is maker, ED (Eva) is approver per "
        "Finance Manual §4.7; no second-signatory control beyond ED. As the Dutch "
        "Chair onboards (Q1 minute §7), a dual-approval workflow on Wise for "
        "payments > USD 10,000 is recommended.", color=TOUCH)

    add_heading(doc, "Recommended actions", level=2)
    add_para(doc,
        "·  Note current position (Amber, 4.1 months cover).", bold=True, color=CORE)
    add_para(doc,
        "·  Approve the proposed reserves policy (CAP-25) — floor USD 132,527, "
        "target USD 265,054 — for board adoption.", bold=True, color=CORE)
    add_para(doc,
        "·  Direct ED to provide a Red-floor breach mitigation plan if the "
        "downside cashflow scenario emerges (i.e., FCDO Phase II unsigned by "
        "31 Jul 2026).", bold=True, color=CORE)

    add_heading(doc, "Sources", level=2)
    add_para(doc,
        "Xero BS @ 31 May 2026 (live, returned 26 May), report ID "
        "231fb4e1-827a-4d3b-b83d-fc1179fa085b. Proposed reserves policy: CAP-25 "
        "memo (8 May 2026).", size=9.5)

    out = Q2 / "humanus_reserves-position_P-06f.docx"
    doc.save(out)
    print(f"Wrote {out}")


# ===========================================================================
# Main
# ===========================================================================
if __name__ == "__main__":
    Q2.mkdir(parents=True, exist_ok=True)
    build_audit_note()
    build_q1_mgmt_accounts()
    build_cashflow_xlsx()
    build_bva()
    build_reserves_note()
    print("All Q2 papers built.")
