"""Build hūmānus Q3 2026 board finance bundle. Run: venv/bin/python build.py <outdir>"""
import sys, os, datetime
sys.path.insert(0, os.path.dirname(__file__))
from data import *
import model
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = sys.argv[1]
WORK = "/Users/evabuzo/multica_workspaces_desktop-api.multica.ai/75c586e2-be30-4ecd-8955-3ba06b11410a/01a12f27/workdir"
LOGO = f"{WORK}/humanus-brand/logo/humanus_logo_master_core.png"
CORE, BOLD, TOUCH = RGBColor(0x41,0x0F,0x37), RGBColor(0x96,0xFF,0x14), RGBColor(0xE1,0x2D,0x64)
FONT = "Fagun"
AS_OF = "12 October 2026"

def f(x, dp=0, paren=True):
    if x is None: return ""
    if abs(x) < 0.5 and dp == 0: return "–"
    s = f"{abs(x):,.{dp}f}"
    return f"({s})" if (x < 0 and paren) else s

def pct(x): return "–" if x is None else f"{x*100:,.0f}%"

# ---------------- docx helpers ----------------
def shade(cell, hex_):
    tcPr = cell._tc.get_or_add_tcPr(); sh = OxmlElement('w:shd')
    sh.set(qn('w:val'), 'clear'); sh.set(qn('w:color'), 'auto'); sh.set(qn('w:fill'), hex_); tcPr.append(sh)

def borders(table, color="D9CFD6"):
    tbl = table._tbl; tblPr = tbl.tblPr
    b = OxmlElement('w:tblBorders')
    for e in ('top','left','bottom','right','insideH','insideV'):
        el = OxmlElement(f'w:{e}'); el.set(qn('w:val'),'single'); el.set(qn('w:sz'),'4'); el.set(qn('w:color'),color); b.append(el)
    tblPr.append(b)

def run(p, text, size=10, bold=False, italic=False, color=CORE, font=FONT):
    r = p.add_run(text); r.font.name = font; r.font.size = Pt(size); r.bold = bold; r.italic = italic
    r.font.color.rgb = color
    r._element.rPr.rFonts.set(qn('w:eastAsia'), font); r._element.rPr.rFonts.set(qn('w:cs'), font)
    return r


def rich(p, t, size, color=CORE, italic=False):
    parts = t.split("**")
    for i, seg in enumerate(parts):
        sub = seg.split("*")
        for k, x in enumerate(sub):
            if x: run(p, x, size, bold=(i % 2 == 1), italic=(italic or k % 2 == 1), color=color)

class Doc:
    def __init__(self, title, subtitle, footer):
        self.d = Document()
        st = self.d.styles['Normal']; st.font.name = FONT; st.font.size = Pt(10); st.font.color.rgb = CORE
        st.element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
        sec = self.d.sections[0]
        sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
        for m in ('left_margin','right_margin'): setattr(sec, m, Cm(2.0))
        sec.top_margin, sec.bottom_margin = Cm(1.8), Cm(1.8)
        fp = sec.footer.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run(fp, footer, 7.5, color=CORE)
        # cover
        p = self.d.add_paragraph(); p.add_run().add_picture(LOGO, width=Cm(5.5))
        self.space(60)
        p = self.d.add_paragraph(); run(p, title, 26, bold=True, font="Fagun Black")
        p = self.d.add_paragraph(); run(p, subtitle, 13)
        self.space(24)
        self.rule()
        self.title = title

    def space(self, pt):
        p = self.d.add_paragraph(); p.paragraph_format.space_after = Pt(pt); return p
    def rule(self, color="410F37"):
        p = self.d.add_paragraph(); pPr = p._p.get_or_add_pPr(); bd = OxmlElement('w:pBdr')
        b = OxmlElement('w:bottom'); b.set(qn('w:val'),'single'); b.set(qn('w:sz'),'12'); b.set(qn('w:color'),color); bd.append(b); pPr.append(bd)
    def meta(self, rows):
        for k, v in rows:
            p = self.d.add_paragraph(); p.paragraph_format.space_after = Pt(2)
            run(p, f"{k}  ", 9, bold=True, font="Fagun ExtraBold"); run(p, v, 10)
    def page(self): self.d.add_page_break()
    def h1(self, t):
        p = self.d.add_paragraph(); p.paragraph_format.space_before = Pt(10); p.paragraph_format.space_after = Pt(4)
        run(p, t, 16, bold=True, font="Fagun ExtraBold")
    def h2(self, t):
        p = self.d.add_paragraph(); p.paragraph_format.space_before = Pt(8); p.paragraph_format.space_after = Pt(2)
        run(p, t, 12, bold=True, font="Fagun Bold")
    def para(self, t, size=10, italic=False, color=CORE, after=4):
        p = self.d.add_paragraph(); p.paragraph_format.space_after = Pt(after)
        # support **bold** segments
        rich(p, t, size, color, italic)
        return p
    def bullets(self, items, size=10):
        for t in items:
            p = self.d.add_paragraph(style='List Bullet'); p.paragraph_format.space_after = Pt(2)
            rich(p, t, size)
    def callout(self, heading, items, fill="96FF14", text=CORE):
        t = self.d.add_table(rows=1, cols=1); t.alignment = WD_TABLE_ALIGNMENT.CENTER
        c = t.cell(0,0); shade(c, fill); c.paragraphs[0].text = ""
        run(c.paragraphs[0], heading, 11, bold=True, color=text, font="Fagun ExtraBold")
        for it in items:
            p = c.add_paragraph(); p.paragraph_format.space_after = Pt(2)
            rich(p, it, 9.5, text)
        self.space(4)
    def table(self, header, rows, widths=None, num_cols=None, total_rows=(), flag_cells=(), size=8.5, note=None):
        t = self.d.add_table(rows=1, cols=len(header)); t.alignment = WD_TABLE_ALIGNMENT.CENTER; borders(t)
        num_cols = set(num_cols if num_cols is not None else range(1, len(header)))
        for j, h in enumerate(header):
            c = t.rows[0].cells[j]; shade(c, "410F37"); c.paragraphs[0].text = ""
            p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if j in num_cols else WD_ALIGN_PARAGRAPH.LEFT
            run(p, h, size, bold=True, color=BOLD)
        for i, r in enumerate(rows):
            cells = t.add_row().cells
            is_tot = i in total_rows
            for j, v in enumerate(r):
                c = cells[j]; c.paragraphs[0].text = ""
                p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if j in num_cols else WD_ALIGN_PARAGRAPH.LEFT
                col = TOUCH if (i, j) in flag_cells else CORE
                run(p, str(v), size, bold=is_tot, color=col)
                if is_tot: shade(c, "EFE8EE")
        if widths:
            t.autofit = False
            grid = t._tbl.tblGrid
            for j, gc in enumerate(grid.findall(qn('w:gridCol'))):
                if j < len(widths): gc.set(qn('w:w'), str(int(widths[j]*567)))
            for row in t.rows:
                for j, w in enumerate(widths): row.cells[j].width = Cm(w)
            tblW = t._tbl.tblPr.find(qn('w:tblW'))
            if tblW is None:
                tblW = OxmlElement('w:tblW'); t._tbl.tblPr.append(tblW)
            tblW.set(qn('w:type'), 'dxa'); tblW.set(qn('w:w'), str(int(sum(widths)*567)))
        if note: self.para(note, 8, italic=True, after=6)
        else: self.space(4)
        return t
    def picture(self, path, width=17):
        p = self.d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.add_run().add_picture(path, width=Cm(width))
    def save(self, path): self.d.save(path)

FOOTER = "hūmānus  ·  Stichting hūmānus international  ·  KvK 77641981  ·  Board finance pack Q3 2026  ·  Confidential — board use only  ·  [TBD: registered-charity footer boilerplate]"

# ---------------- computations ----------------
acct = {a: (q3, q2, ytd) for a, g, q3, q2, ytd in PL}
groups = []
for g in ["Staff & consultants","Programme & travel","Governance & operations","FX revaluation"]:
    q3 = sum(x[2] for x in PL if x[1]==g); q2 = sum(x[3] for x in PL if x[1]==g); y = sum(x[4] for x in PL if x[1]==g)
    groups.append((g,q3,q2,y))
assert abs(sum(x[1] for x in groups)-EXP_TOT["Q3"])<0.05 and abs(sum(x[3] for x in groups)-EXP_TOT["YTD"])<0.05 and abs(sum(x[2] for x in groups)-EXP_TOT["Q2"])<0.05

# Q4 forecast by account (central)
Q4 = {"Executive Director":42600.00,"Cost of Living Allowances":round(3*1400*model.GBPUSD,2),
      "Advocate / Programme Manager":10500.00,"Rohingya Manager":7500.00,"Communication Advisor":3*5383.27,
      "Communication Support":3*649.41,"Accounting & Bookkeeping Fees":4000.00,"Operations Support":round(700*model.GBPUSD,2),
      "Project Consultants":round(2000*model.EURUSD,2),"IT & Software Expenses":3*352.70,"Bank Charges & Merchant Fees":3*55.74,
      "International Flights":4500.0,"Accommodation":3000.0,"Per diem":900.0,"Local Transportation":600.0,
      "Partner Grants":4000.0,"Meeting Expenses":2000.0}
Q4_TOT = sum(Q4.values())
FY_EXP = EXP_TOT["YTD"] + Q4_TOT
FCDO_CLAIM1 = 52800.0
FY_INC = INC_TOT["YTD"] + FCDO_CLAIM1 + 100.0
FY_NET = FY_INC - FY_EXP
NA_DEC = NET_ASSETS[0] + (FY_NET - NET["YTD"])
# accrued but uninvoiced restricted income at 31 Dec (spent, claimable later)
FCDO_NOVDEC = 2*(5680+2380+330+240)
SWMY_ACCR = (27410 + 3*(2130+2500)) - 24946
NA_DEC_ADJ = NA_DEC + FCDO_NOVDEC + SWMY_ACCR
ACCR_SEP = 43847.67 + 2464.00
NA_SEP_ADJ = NET_ASSETS[0] + ACCR_SEP

# BvA
def actual(accts, col=2): return sum(acct[a][col] for a in accts)
bva_rows = []
for name, accts, wb, bb, typ in BVA:
    ytd = actual(accts); fy = ytd + sum(Q4.get(a,0) for a in accts)
    bva_rows.append((name, typ, wb, bb, ytd, fy))
FX_YTD = acct["Foreign Currency Gains and Losses"][2]
assert abs(sum(r[4] for r in bva_rows) + FX_YTD - EXP_TOT["YTD"]) < 0.05

SC = model.scen()
central = SC["A. Central"]

# FY deficit decomposition
OSF_Q4 = Q4_TOT - (8952.33 + 2*8630) - 3*4630
decomp = [
 ("Planned SDC co-funding, final year — hūmānus' matched share (Art. 1.2); grant balanced and audited", 28130.47-73790.56),
 ("Restricted timing — spent in 2026, claimable in 2027 (FCDO II Nov–Dec + Swiss Myanmar)", -(FCDO_NOVDEC+SWMY_ACCR)),
 ("FCDO Phase I — 2026 receipts covering 2025 spend", 8132.90),
 ("FCDO Phase II Jun–Oct — claim matched to spend", FCDO_CLAIM1 - (43847.67+8952.33)),
 ("Core (OSF #4 + interest) — income less core spend", (300023.70+1575.45+100) - (203699.50+OSF_Q4)),
 ("FX revaluation and legacy codes", -(235.04+42.38)),
]
DECOMP_RES = FY_NET - sum(v for _, v in decomp)

if __name__ == "__main__" and "--check" in sys.argv:
    print("Q4", round(Q4_TOT,2), "FY exp", round(FY_EXP,2), "FY inc", round(FY_INC,2), "FY net", round(FY_NET,2))
    print("NA Dec", round(NA_DEC,2), "adj", round(NA_DEC_ADJ,2), "NA Sep adj", round(NA_SEP_ADJ,2))
    for d_ in decomp: print(" ", d_[0], round(d_[1],2))
    print("decomp residual", round(DECOMP_RES,2))
    for r in bva_rows: print(r)
    sys.exit()

# =====================================================================
os.makedirs(OUT, exist_ok=True)
import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
for fp_ in ["/Users/evabuzo/Library/Fonts/Fagun-Regular.otf","/Users/evabuzo/Library/Fonts/Fagun-Bold.otf"]:
    if os.path.exists(fp_): font_manager.fontManager.addfont(fp_)
plt.rcParams["font.family"] = ["Fagun","Helvetica","Arial"]
CH = os.path.join("/tmp/stacy_cap521", "cash_chart.png")
fig, ax = plt.subplots(figsize=(9, 4.2), dpi=200)
xs = ["Sep 26\n(actual)"] + model.MONTHS
def series(k): return [model.OPEN_CASH] + [r["close"] for r in SC[k]]
styles = [("A. Central","#410F37","-",2.6),("B. Freeze (committed only)","#410F37","--",1.6),
          ("C. Q3 run-rate continues","#E12D64","-",1.6),("F. Central + Swiss T2 pulled to Dec 26","#8C6A84",":",2.0)]
for k,c,ls,lw in styles:
    ax.plot(xs, [v/1000 for v in series(k)], color=c, ls=ls, lw=lw, marker="o", ms=3.5, label=k)
ax.axhline(0, color="#E12D64", lw=1)
ax.set_ylabel("Month-end cash, USD thousand", color="#410F37")
ax.set_title("Month-end cash, Oct 2026 – Jun 2027 (USD thousand)", color="#410F37", loc="left", fontsize=12)
ax.tick_params(colors="#410F37", labelsize=8)
for s in ax.spines.values(): s.set_color("#D9CFD6")
ax.grid(axis="y", color="#EFE8EE"); ax.legend(frameon=False, fontsize=8, loc="upper left", labelcolor="#410F37")
ax.annotate("OSF Year-3 tranche\nUSD 300k (Jan 2027)", xy=(4, series("A. Central")[4]/1000), xytext=(5.1, 90),
            fontsize=8, color="#410F37", arrowprops=dict(arrowstyle="->", color="#410F37"))
plt.tight_layout(); plt.savefig(CH); plt.close()

# ---------------------------------------------------------------- PAPER 1
D = Doc("Board Finance Report — Q3 2026",
        "Management accounts to 30 September 2026 (unaudited)", FOOTER)
D.meta([("Board meeting","29 October 2026 (virtual)"),("Period","1 July – 30 September 2026, with year-to-date"),
        ("Data","Xero, pulled live 12 October 2026; signed grant agreements"),
        ("Currency","USD (organisational base currency) unless stated"),("Companion paper","Forecast Report, October 2026 – June 2027")])
D.page()

D.h1("1. Summary")
D.para(f"**Q3 closed with cash of USD {f(33679.30)} and net assets of USD {f(NET_ASSETS[0])}.** Spending is broadly on plan; the pressure is cash timing. "
       "We have no contracted receipts until the FCDO Phase II claim (indicatively November) and the OSF Year-3 tranche (January 2027).")
D.bullets([
 f"**Q3 net result: deficit USD {f(-NET['Q3'])}.** Expected: no grant income fell due in Q3. The SDC final instalment and Swiss Myanmar tranche 1 were invoiced at 30 June and recorded in Q2. Q3 spending was USD {f(EXP_TOT['Q3'])}, down 19% on Q2 (USD {f(EXP_TOT['Q2'])}).",
 f"**Year to date: surplus USD {f(NET['YTD'])}.** Income USD {f(INC_TOT['YTD'])}, of which OSF Year 2 was USD 300,000. Expenditure USD {f(EXP_TOT['YTD'])}: 82% of the working budget (USD 488,450) at 75% of the year, or 64% of the board-approved budget (USD 631,450).",
 "**The overrun against the working budget is in travel, board and meeting costs**, plus Confidence Building Measures (USD 14,598, funded by FCDO Phase I). Staff and consultant costs are on budget.",
 f"**Restricted grants are earned but not yet invoiced.** FCDO Phase II costs of USD 43,848 are claimable from Stimson (first claim due 16 October). Swiss Myanmar costs are USD 2,464 ahead of tranche 1. Counting these, adjusted net assets are USD {f(NA_SEP_ADJ)}.",
 f"**Liquidity is the issue for this meeting.** Cash on 12 October was USD 31,329. The 28 October payment run plus approved bills come to about USD 32,700, so the run cannot be paid in full without action. See the Forecast Report.",
])

D.h1("2. Income and expenditure")
D.para("Accrual basis, per Xero. Quarterly columns are three-month totals; YTD is 1 January – 30 September 2026.", 8.5, italic=True)
rows = [("Grant income", f(0), f(114081.47), f(414081.47)), ("Interest income", f(431.74), f(695.31), f(1599.15)),
        ("Total income", f(INC_TOT['Q3']), f(INC_TOT['Q2']), f(INC_TOT['YTD']))]
for g,q3,q2,y in groups: rows.append((g, f(q3), f(q2), f(y)))
rows += [("Total expenditure", f(EXP_TOT['Q3']), f(EXP_TOT['Q2']), f(EXP_TOT['YTD'])),
         ("Net surplus / (deficit)", f(NET['Q3']), f(NET['Q2']), f(NET['YTD']))]
D.table(["USD","Q3 2026 (Jul–Sep)","Q2 2026 (Apr–Jun)","YTD 2026 (9 months)"], rows, widths=[7,3.3,3.3,3.4],
        total_rows=(2,8,9), flag_cells={(9,1)}, note="Account-level detail in Appendix A. Check: Executive Director USD 42,600 in Q3 = 3 × 14,200; USD 127,800 YTD = 9 × 14,200.")
D.h2("What moved in Q3")
D.bullets([
 "**Staff and consultants USD 94,013** (Q2 USD 94,505), flat as expected. The Q3 figure includes USD 2,294 Project Consultants: the Chair's stipend invoice of 29 July (EUR 2,000).",
 "**Programme and travel USD 25,310** (Q2 USD 42,952). Partner grants and volunteer stipends fell; September included New York accommodation (USD 4,759) and USD 1,786 of flights charged to FCDO Phase II.",
 "**Governance and operations USD 4,140** (Q2 USD 14,586). Q2 carried board meeting costs (USD 8,936) and audit fees (USD 1,981).",
])

D.h1("3. Balance sheet")
bs_rows = [(k, f(v[0]), f(v[1]), f(v[2])) for k, v in BS.items()]
bs_rows.append(("Net assets (total funds)", f(NET_ASSETS[0]), f(NET_ASSETS[1]), f(NET_ASSETS[2])))
bs_rows.append(("Add: grant income earned, not yet invoiced", f(ACCR_SEP), "–", "–"))
bs_rows.append(("Adjusted net assets", f(NA_SEP_ADJ), f(NET_ASSETS[1]), f(NET_ASSETS[2])))
D.table(["USD","30 Sep 2026","30 Jun 2026","31 Dec 2025"], bs_rows, widths=[7.6,3.1,3.1,3.1], total_rows=(10,12),
        note="Liabilities shown in brackets. Earned-not-invoiced = FCDO Phase II costs USD 43,847.67 (claim to Stimson, Schedule 2) + Swiss Myanmar costs ahead of tranche 1 USD 2,464.00. Xero does not record this until invoices are raised.")
D.bullets([
 "**Cash fell USD 65,146 in Q3**, from USD 98,825 to USD 33,679. Receipts of USD 53,076 (SDC final instalment + Swiss Myanmar tranche 1) cleared the 30 June receivables.",
 "**No receivables are invoiced at 30 September.** The FCDO Phase II claim is the next one.",
 "**Staff advance USD 2,065** is being recovered at USD 1,000 a month from the Rohingya Manager's fee and clears in December.",
 "**Accounts payable USD 1,914** at 30 September; two further bills (Chair stipend EUR 2,000, Free Thought GBP 700) are awaiting approval.",
])

D.h1("4. Funds — restricted and unrestricted")
D.para("2026 movements by funding source, Xero Donors tracking. Expenditure excludes FX revaluation, which is shown on its own line. Transfers show hūmānus' planned matched contribution to the SDC project from core funds.", 8.5, italic=True)
fr = []; ti=te=0
for n,t,i_,e_,tr_,note in FUNDS:
    fr.append((n, t, f(i_), f(e_), f(tr_), f(i_-e_+tr_), note)); ti+=i_; te+=e_
fr.append(("Total","",f(ti),f(te),"–",f(ti-te),"= YTD net surplus"))
D.table(["Fund","Type","Income","Expenditure","Transfers","Net","Note"], fr, widths=[4.2,1.8,1.9,2.0,1.9,1.9,3.6], num_cols={2,3,4,5},
        total_rows=(len(fr)-1,), size=7.2)
D.bullets([
 "**Core (OSF #4 general support) is in surplus by USD 50,664 for the year to date, after meeting our planned share of the SDC project.** That surplus has to carry the organisation to January 2027.",
 "**SDC (closed 30 April 2026) balanced over the life of the grant and has been audited.** FDFA paid its full contribution, CHF 182,788.10 (USD 221,668). The agreement set FDFA's share at no more than half the project budget. Per Art. 1.2: *\"The contribution of the FDFA is intended to cover at the maximum 50% of the total budget.\"* hūmānus' matched share was planned from the start: USD 98,371 over the grant's life, of which USD 45,660 fell in 2026. It is shown as a transfer from core funds, not an overspend.",
 "**FCDO Phase II USD (43,848) and Swiss Myanmar USD (2,464) are timing only.** Both grants pay in arrears; the gaps close when the claims are paid.",
])

D.h1("5. Budget versus actual — year to date")
D.para("Monthly lines are compared with 9/12 of the annual budget. One-off and annual envelopes are compared with the full-year amount and are never pro-rated. Both budgets are shown: the board approved USD 631,450 in December 2025; the working budget (USD 488,450) removes the Senior Legal Advisor, half the lawyer fees and some travel.", 8.5, italic=True)
br = []; flags=set()
tot = {"M":[0,0,0],"L":[0,0,0],"U":[0,0,0]}
for i,(n,typ,wb,bb,ytd,fy) in enumerate(bva_rows):
    basis = wb*0.75 if typ=="M" else wb
    var = ytd - basis
    br.append((n, {"M":"Monthly","L":"Envelope","U":"Unbudgeted"}[typ], f(wb), f(bb), f(basis), f(ytd), f(var), pct(ytd/wb) if wb else "n/a"))
    if (wb and ytd/wb > (0.85 if typ=="M" else 1.0)) or typ=="U": flags.add((i,6))
    tot[typ][0]+=wb; tot[typ][1]+=bb; tot[typ][2]+=ytd
br.append(("FX revaluation","—","–","–","–",f(FX_YTD),f(FX_YTD),"n/a"))
wb_all = sum(r[2] for r in bva_rows); bb_all = sum(r[3] for r in bva_rows)
basis_all = sum((r[2]*0.75 if r[1]=="M" else r[2]) for r in bva_rows)
br.append(("Total",  "", f(wb_all), f(bb_all), f(basis_all), f(EXP_TOT['YTD']), f(EXP_TOT['YTD']-basis_all), pct(EXP_TOT['YTD']/wb_all)))
D.table(["Budget line","Type","Working budget","Board budget","YTD basis (working)","Actual YTD","Variance","% of working"], br,
        widths=[4.6,1.6,1.8,1.8,1.9,1.9,1.8,1.6], num_cols={2,3,4,5,6,7}, total_rows=(len(br)-1,), flag_cells=flags, size=7)
mon_ytd = tot["M"][2]; mon_basis = tot["M"][0]*0.75
D.bullets([
 f"**Monthly staff and running costs: USD {f(mon_ytd)} against USD {f(mon_basis)} pro-rata ({pct(mon_ytd/mon_basis)}).** On budget. Communications and finance are slightly under; the Programme Manager line is USD 2,500 over because Q1 carried USD 13,000 against USD 10,500.",
 "**Envelopes over budget: travel USD 39,178 against USD 25,000 working / USD 36,000 board**, mainly international flights (USD 20,519 in Q2). **Board meetings USD 8,936 against USD 4,000.** Volunteer stipends are fully used.",
 "**Envelopes under budget, available to offset: Azadi partner grants (USD 14,000 of 30,000), lawyer fees (USD 4,363 of 12,000), operations support (USD 6,267 of 18,000).**",
 f"**Against the board-approved budget, YTD spending is USD {f(631450*0.75-EXP_TOT['YTD'])} below a straight 9/12 pace (USD {f(631450*0.75)}).** The unfilled Senior Legal Advisor post accounts for USD 90,000 of that pace; without it, spending is about USD {f(EXP_TOT['YTD']-(631450-120000)*0.75)} ahead.",
])

D.h1("6. Grants and contracts")
D.para("Each live agreement reconciled to Xero receipts, in the contract currency.", 8.5, italic=True)
D.h2("Summary")
D.bullets([
 "**(a) Received in 2026:** OSF Year 2 USD 300,000 (Jan); FCDO Phase I USD 61,005 (INV-0007 + INV-0008); SDC final CHF 22,788.10 = USD 28,130.47 (INV-0011, paid 21 Jul); Swiss Myanmar tranche 1 USD 24,946.00 (INV-0010, paid 10 Jul).",
 "**(b) Contracted, not yet received:** FCDO Phase II claim 1 (Jun–Oct, about USD 52,800, gated on the 16 Oct report and on Stimson receiving FCDO funds); OSF Year 3 USD 300,000 (January 2027); Swiss Myanmar tranches 2 and 3 (USD 24,946.67 on 31 Jan 2027; USD 24,946.66 on 30 Jun 2027).",
])
g_rows = [
 ("OSF #4 Y2","Jan 2026","USD 300,000.00","bank 22 Jan 26","300,000","PAID"),
 ("FCDO II claim 1","16 Oct report → ind. Nov 2026","GBP — reported expenses","USD 54,017","—",""),
 ("FCDO II claim 2","30 Apr 2027 → ind. May 2027","GBP — reported expenses","—","—","Within £80,884.70 ceiling (cl. 6.1)"),
 ("Swiss Myanmar T1","30 Jun 2026","USD 24,946.67","INV-0010","24,946.00","PAID 10 Jul"),
 ("Swiss Myanmar T2","31 Jan 2027","USD 24,946.67","—","—","Gated: invoice + FDFA funds; Stimson indicated Dec 2026"),
 ("Swiss Myanmar T3","30 Jun 2027","USD 24,946.66","—","—","Gated as T2"),
 ("SDC (closed) T1–T3","Nov 24 / Aug 25 / after final report","CHF 182,788.10","INV-0005/6/11","221,668.08","FULLY DRAWN"),
]
D.table(["Contracted payment","Due","Amount (contract ccy)","Xero ref","Received (USD)","Status"], g_rows,
        widths=[2.9,3.2,2.9,2.6,2.0,3.9], num_cols={4}, size=7)
D.h2("Burn against grant life (30 September 2026)")
D.table(["Grant","Ceiling","Spent to date","% spent","% of term elapsed","Read"], [
 ("OSF #4 (lifetime)","USD 1,500,000","USD 455,619","30.4%","35.0% (21 of 60 mo)","On track — slightly under"),
 ("FCDO Phase II","GBP 80,884.70","USD 43,848 ≈ GBP 32,360 @1.355","40.0%","40.0% (4 of 10 mo)","On track"),
 ("Swiss Myanmar","USD 74,840","USD 27,410","36.6%","55.6% (10 of 18 mo)","Under-spending — only ED and Rohingya Manager are charged"),
], widths=[3,2.6,3.8,1.6,2.8,3.6], num_cols={3}, size=7.5, flag_cells={(2,5)})

D.h1("7. Cash and reserves")
D.table(["Month-end","Total cash","Wise USD","Wise EUR (USD equiv.)","Wise GBP"],
        [(m, f(t), f(u), f(e), f(g)) for m,t,u,e,g in CASH_MONTHS], widths=[4.5,3,3,3.5,3], size=8)
D.bullets([
 "**Q3 cash movement:** opening USD 98,825; receipts USD 53,508 (SDC USD 28,130, Swiss Myanmar USD 24,946, interest USD 432); spending USD 118,513 net of transfers between wallets; FX USD (141); closing USD 33,679.",
 "**Committed monthly cash floor: about USD 29,500** (consultant roster, Communications Advisor, COLA, IT and bank charges). Everything above this is programme and travel.",
 f"**Months of cover:** Xero net assets USD {f(NET_ASSETS[0])} = 1.3 months of the floor; adjusted net assets USD {f(NA_SEP_ADJ)} = 2.8 months. Both are below the proposed 3-month reserves floor.",
 "**Most of the cash is in EUR** (USD 28,872 of USD 31,329 on 12 October). The USD-denominated payment run needs an EUR→USD conversion before 28 October.",
])

D.h1("8. Risk and compliance")
D.table(["#","Risk / flag","Rating","Action (owner, date)"], [
 ("1","Liquidity: on committed costs alone, cash is short at the 28 Oct payment run and again in December","High","Liquidity plan, Forecast Report §4 (ED + FM, now)"),
 ("2",f"Xero will show negative net assets at 31 Dec 2026 (forecast USD {f(NA_DEC)}) unless earned grant income is accrued","High","Year-end accrual of earned grant income; confirm treatment with auditor (FM, Dec)"),
 ("3","Related party: Chair stipend invoiced EUR 2,000/quarter vs TOR §4 USD 2,000 (≈ USD 250–295/quarter above TOR); coded to a closed grant (OSF #1)","Medium","Board to confirm currency (Chair recused); re-code to OSF #4 core; Legal to confirm ANBI remuneration rules"),
 ("4","Repeating templates for the ED fee and COLA end 31 Dec 2026; the SK Das template still shows USD 2,000 although the fee falls to USD 1,000 (effective date to confirm)","Low","Renew/amend templates (FM, by 25 Nov)"),
 ("5","Swiss Myanmar under-spend (−19pp against term): risk of leaving grant money unclaimed","Medium","Charge eligible programme costs; plan Y2 activity (ED, Q4)"),
], widths=[0.7,7.6,1.6,7], num_cols=set(), size=7.5, flag_cells={(0,2),(1,2)})

D.page()
D.h1("Appendix A — Income and expenditure by account")
ar = [(n, g, f(q3,2), f(q2,2), f(y,2)) for n,g,q3,q2,y in PL]
ar.append(("Total expenditure","",f(EXP_TOT['Q3'],2),f(EXP_TOT['Q2'],2),f(EXP_TOT['YTD'],2)))
D.table(["Account","Group","Q3 2026","Q2 2026","YTD 2026"], ar, widths=[5.2,4,2.6,2.6,2.6], num_cols={2,3,4}, total_rows=(len(ar)-1,), size=7.5)
D.h1("Appendix B — Sources and basis")
D.bullets([
 "Xero (Stichting hūmānus international, base USD), pulled 12 Oct 2026: P&L Q3 `64172c20`; P&L YTD `9aa8308a`; P&L by quarter `16b22d31`; P&L Q2 direct `3aeb792c`; Donors P&L YTD `13b97ef5`, Q3 `7d6a8730`; Balance Sheet 30 Sep / 30 Jun `052baafa`, 31 Dec 2025 `cfaec7b2`; Bank Summary Q3 `6d5984f7`, 1–12 Oct `f55f65a4`.",
 "Budgets: working `finance-docs/budgets/2026 Budget-2.xlsx` (USD 488,450); board-approved `humanusGovernance/2026 Budget.xlsx` (USD 631,450).",
 "Mapping of Xero accounts to budget lines: Travel = international + domestic flights, accommodation, per diem; Vehicle Hire & Sundry = local transportation; Fees = bank charges. FX revaluation is excluded from variance.",
 "Contracts: OSF OR2024-94840 (finance-docs/grants); FCDO Phase II sub-grant 2026.06.24 and Swiss Myanmar Partnership Agreement 2026.02.04 (finance-docs/contracts); SDC 7F-03317.09.15 (scan).",
 "Accounts are unaudited and on an accrual basis. Grant income is recorded when invoiced or received.",
])
P1 = os.path.join(OUT, "humanus_board-finance-report_Q3-2026.docx"); D.save(P1)

# ---------------------------------------------------------------- PAPER 2
F = Doc("Forecast Report — October 2026 to June 2027",
        "2026 full-year outturn and nine-month cash forecast", FOOTER)
F.meta([("Board meeting","29 October 2026 (virtual)"),("Opening position","Xero 30 September 2026; cash checked live 12 October 2026"),
        ("Currency","USD; GBP and EUR converted at Xero rates of late Sep / early Oct 2026"),
        ("Companion paper","Board Finance Report, Q3 2026")])
F.page()
A = central; Bf = SC["B. Freeze (committed only)"]
F.h1("1. Headline")
F.para(f"**The 2026 deficit comes from grant-claim timing and our planned co-funding of the SDC project, not overspending.** Core (OSF) activity runs a surplus of about USD {f(decomp[4][1])}. Even so, **cash runs short twice before the OSF Year-3 tranche arrives in January 2027**: at the 28 October payment run, and in December.")
F.bullets([
 f"**2026 outturn (central):** expenditure USD {f(FY_EXP)}, which is 4.3% over the working budget (USD 488,450) and 19% under the board-approved budget (USD 631,450). Income USD {f(FY_INC)}. Reported deficit USD {f(-FY_NET)}.",
 f"**What drives the deficit:** our planned matched share of the SDC project's final year, USD {f(-decomp[0][1])}; grant income earned in 2026 but claimable only in 2027, USD {f(-decomp[1][1])}. These are partly offset by the core surplus.",
 f"**Cash, central case:** USD {f(A[0]['close'])} at end-October, USD {f(A[1]['close'])} at end-November (after the FCDO claim), USD {f(A[2]['close'])} at end-December, then USD {f(A[3]['close'])} at end-January once OSF pays.",
 f"**Even with a complete spending freeze, cash is short:** USD {f(Bf[0]['close'])} in October and USD {f(Bf[2]['close'])} in December. Cutting costs alone cannot fix this; the receipts have to come earlier.",
 "**What fixes it:** OSF releasing the January tranche in December (it paid in December in 2024) **or** Stimson paying the FCDO claim on receipt and Swiss tranche 2 in December, together with a freeze on discretionary spending.",
])

F.h1("2. 2026 full-year outturn")
F.para("Accrual basis, central case. Q4 is built from committed costs (repeating invoice templates, approved bills, the Communications Advisor's monthly invoice) plus USD 5,000 a month of programme and travel spending.", 8.5, italic=True)
F.table(["USD","Actual YTD (9 mo)","Forecast Q4","Forecast FY 2026","Working budget","Board budget"], [
 ("Income", f(INC_TOT['YTD']), f(FCDO_CLAIM1+100), f(FY_INC), "488,450†", "631,450†"),
 ("Expenditure", f(EXP_TOT['YTD']), f(Q4_TOT), f(FY_EXP), "488,450", "631,450"),
 ("Net surplus / (deficit)", f(NET['YTD']), f(FCDO_CLAIM1+100-Q4_TOT), f(FY_NET), "–", "–"),
 ("Net assets at year-end (Xero basis)", "", "", f(NA_DEC), "", ""),
 ("Add: earned grant income to accrue (FCDO II Nov–Dec, Swiss Myanmar)", "", "", f(FCDO_NOVDEC+SWMY_ACCR), "", ""),
 ("Adjusted net assets at 31 Dec 2026", "", "", f(NA_DEC_ADJ), "", ""),
], widths=[5.6,2.3,2.1,2.4,2.2,2.2], total_rows=(2,5), flag_cells={(2,3),(3,3)},
 note="† Both budgets were set as balanced plans: the funding columns equal planned spend. Income includes FCDO Phase II claim 1 (Jun–Oct, about USD 52,800) on the basis that it is invoiced in October.")
F.h2("Where the deficit comes from")
F.table(["Driver","USD"], [(n, f(v)) for n,v in decomp] + [("Reported 2026 deficit (central)", f(FY_NET))],
        widths=[13.5,3.3], total_rows=(len(decomp),), size=8.5)
F.bullets([
 "**The SDC co-funding was planned and is now complete.** The project closed on 30 April 2026, balanced over the life of the grant, and has been audited. FDFA paid its contribution in full, and nothing further falls in 2027.",
 "**The timing items reverse in 2027.** FCDO Phase II Nov–Dec costs are claimed in the April 2027 report, and the Swiss Myanmar costs are covered by tranches 2 and 3.",
 f"**Core is in surplus.** OSF-funded spending of about USD {f(203699.50+OSF_Q4)} is below the USD 286,635 working-budget allocation and inside the USD 300,000 tranche.",
])
F.h2("Expenditure by budget line — full-year forecast")
fr2=[]; fl2=set()
for i,(n,typ,wb,bb,ytd,fy) in enumerate(bva_rows):
    fr2.append((n, f(wb), f(bb), f(ytd), f(fy), f(fy-wb)))
    if fy-wb > 1000: fl2.add((i,5))
fr2.append(("FX revaluation","–","–",f(FX_YTD),f(FX_YTD),f(FX_YTD)))
fr2.append(("Total", f(wb_all), f(bb_all), f(EXP_TOT['YTD']), f(FY_EXP), f(FY_EXP-wb_all)))
F.table(["Budget line","Working budget","Board budget","Actual YTD","Forecast FY","FY vs working"], fr2,
        widths=[5.8,2.2,2.2,2.2,2.2,2.2], total_rows=(len(fr2)-1,), flag_cells=fl2, size=7)

F.h1("3. Cash forecast — October 2026 to June 2027")
F.picture(CH, 16.5)
F.para("Month-end closing cash, USD. Central case below. Every column is one month of cash in and cash out.", 8.5, italic=True)
hdr = ["USD"] + model.MONTHS
floor_lines = list(model.floor(0).keys())
rows=[]; rows.append(["Opening cash"] + [f(model.OPEN_CASH if i==0 else A[i-1]['close']) for i in range(9)])
for ln in floor_lines: rows.append([ln] + [f(model.floor(i)[ln]) for i in range(9)])
rows.append(["One-off committed bills"] + [f(A[i]['oneoff']) for i in range(9)])
rows.append(["Programme & travel (central)"] + [f(A[i]['variable']) for i in range(9)])
rows.append(["Total cash out"] + [f(A[i]['out']) for i in range(9)])
for mi,l,a in model.INFLOWS["base"]:
    rows.append([l] + [f(a) if i==mi else "–" for i in range(9)])
rows.append(["Closing cash"] + [f(A[i]['close']) for i in range(9)])
neg = {(len(rows)-1, j+1) for j in range(9) if A[j]['close'] < 0}
F.table(hdr, rows, widths=[4.6]+[1.38]*9, total_rows=(0, len(floor_lines)+3, len(rows)-1), flag_cells=neg, size=6.3)
F.h2("Scenarios — month-end closing cash (USD)")
srows=[[k]+[f(r['close']) for r in v] for k,v in SC.items()]
sneg={(i,j+1) for i,(k,v) in enumerate(SC.items()) for j,r in enumerate(v) if r['close']<0}
F.table(["Scenario"]+model.MONTHS, srows, widths=[4.6]+[1.38]*9, flag_cells=sneg, size=6.3,
        note="A: USD 5,000/month programme and travel. B: committed costs only. C: Q3's average variable spend continues (USD 10,636/month). D: as A, but FCDO claim 1 slips to January. E: as A, with OSF paying in early December. F: as A, with Swiss Myanmar tranche 2 pulled into December.")
F.bullets([
 "**Every scenario is short in October.** This is the 28 October payment run; no scenario brings receipts in before it.",
 f"**December is short unless one receipt moves earlier.** Freeze + Swiss tranche 2 in December closes at about USD {f(Bf[2]['close']+24946.67)}; OSF in December closes at about USD {f(SC['E. Central + OSF released early Dec 26'][2]['close'])}.",
 f"**If the FCDO claim slips to January, the low point is USD {f(min(r['close'] for r in SC['D. Central, FCDO claim slips to Jan 27']))} at end-December.** That is why the standby bridge is sized at USD 75,000.",
 f"**After January, cash recovers to about USD {f(A[3]['close'])}** and stays above three months of the floor through June 2027 in cases A and B. In case C it falls to about USD {f(SC['C. Q3 run-rate continues'][8]['close'])} by June, which is the warning for the 2027 budget.",
])

F.h1("4. The October pinch point")
F.table(["Item","Date","USD"], [
 ("Cash at 12 October (Wise USD 2,456 + EUR 28,872)","12 Oct",f(31328.83)),
 ("Chair stipend Q3 — EUR 2,000 (awaiting approval)","due 15 Oct",f(-2000*model.EURUSD)),
 ("Free Thought Ltd — GBP 700, speech support (awaiting approval)","due 15 Oct",f(-700*model.GBPUSD)),
 ("Payment run: ED 14,200; Comms Advisor ~5,383; Programme Mgr 3,500; Finance Mgr 2,000; COLA GBP 1,400; Rohingya Mgr 1,500 (net of advance); comms support, IT, bank","25 → 28 Oct",f(-sum(model.floor(0).values()))),
 ("Position after the run, before any programme or travel spend","28 Oct",f(31328.83-2000*model.EURUSD-700*model.GBPUSD-sum(model.floor(0).values()))),
], widths=[11.5,2.4,2.9], num_cols={2}, total_rows=(4,), flag_cells={(4,2)}, size=8)
F.para("**Options before 28 October, in order of preference:** (1) re-state FCDO claim 1 in GBP from ledger actuals and submit it early, asking Stimson to pay on receipt (cl. 6.5 rules out advances, but claim 1 covers costs already incurred); (2) defer the two non-staff bills (about USD 3,180) to November; (3) if still short, the board agrees which roster payment is delayed and for how long. Any deferral of the ED's fee is a related-party matter for the board, and Legal should advise first.")
F.para("**Correction to the 12 October dashboard note (CAP-520):** that note put end-October cash at about USD 5,100. It left out the Communications Advisor's monthly invoice (about USD 5,383), which is raised by hand and not from a template, and Communication Support (about USD 649). The correct figure is shown above.", 8.5, italic=True)

F.h1("5. Change since the August cash plan")
F.para("The August cash ceiling (CAP-389, 18 August) projected September USD 50,800, October USD 21,650, November USD 29,700 and December USD 550.", 9)
F.table(["Month-end","August plan","Now (freeze case B)","Change","Why"], [
 ("Sep 2026","50,800",f(33679.30),f(33679.30-50800),"Aug closed USD 3,837 below plan. Sep net cash decline was USD 42,433 against a USD 29,150 ceiling: Tavy Cussinel F-2026-002 GBP 4,000 (USD 5,303), New York accommodation USD 4,759 and Pennyhill Park balance GBP 1,888 (USD 2,525); other lines net USD 696 over"),
 ("Oct 2026","21,650",f(Bf[0]['close']),f(Bf[0]['close']-21650),"Sep shortfall carried forward (USD 17,121); Oct bills not in the plan (Chair stipend, Free Thought, SK Das Sep fee paid 1 Oct) USD 5,283; floor USD 347 above ceiling"),
 ("Nov 2026","29,700",f(Bf[1]['close']),f(Bf[1]['close']-29700),"October shortfall carried forward (USD 22,751), partly offset by a larger FCDO claim (+USD 15,600): about USD 52,800, not USD 37,200, because Schedule 2's first period runs June–October, not June–September"),
 ("Dec 2026","550",f(Bf[2]['close']),f(Bf[2]['close']-550),"November shortfall carried forward (USD 6,498); December floor USD 282 above the ceiling"),
], widths=[2,1.9,2.5,1.8,8.6], num_cols={1,2,3}, size=7.5)

F.h1("6. Assumptions")
F.table(["Area","Assumption","Source"], [
 ("Opening cash","USD 33,679.30 at 30 Sep 2026","Xero Balance Sheet `052baafa`"),
 ("Roster","ED USD 14,200; Programme Mgr 3,500; Rohingya Mgr 2,500 (cash 1,500 until the advance clears in Dec); COLA GBP 1,400; SK Das 2,000 in Oct, 1,000 from Nov","Xero repeating templates; CAP-485 fee decision (effective date to confirm)"),
 ("Communications","Advisor USD 5,383/month and Support USD 649/month (Q3 averages); the Advisor's invoice is raised by hand","Xero P&L Q3"),
 ("Overheads","IT USD 353, bank USD 56 per month (Q3 averages)","Xero P&L Q3"),
 ("Programme & travel","Central USD 5,000/month; Freeze nil; Q3 run-rate USD 10,636/month","Q3 spend above the committed floor"),
 ("FCDO Phase II","Claim 1 (Jun–Oct) about USD 52,800 in late Nov; claim 2 (Nov–Mar) about USD 47,000 in May 2027; both within the £80,884.70 ceiling","Sub-grant Sched. 2; ledger USD 43,847.67 to 30 Sep + Oct roster share"),
 ("OSF #4","USD 300,000 on about 22 Jan 2027 (the 2026 tranche landed 22 Jan)","Award letter p.1"),
 ("Swiss Myanmar","USD 24,946.67 on 31 Jan 2027; USD 24,946.66 on 30 Jun 2027","Partnership Agreement §3.3"),
 ("Rohingya Manager","Template and Swiss Myanmar term end May 2027; no cost from June 2027","§17; Xero template"),
 ("Audit","2026 audit about USD 3,000, paid April 2027","Deferred per ED, May 2026"),
 ("Excluded","Any new grant, the Q4 Chair stipend (paid in arrears, Jan 2027), FX movements, interest","Prudence"),
 ("FX","GBP/USD 1.3258; EUR/USD 1.1255","Xero rates on late-Sep / early-Oct transactions"),
], widths=[2.8,9.2,4.8], num_cols=set(), size=7.5)
F.h2("Proposed year-end accrual (draft, for auditor agreement)")
F.para("Assets increase with a debit; income increases with a credit. FCDO Phase II and Swiss Myanmar costs incurred in 2026 have earned the grant income, but it is invoiced in 2027, so we recognise an accrued-income asset and the matching restricted income:", 9)
F.para(f"Dr  Accrued income — grants (current asset)            USD {f(FCDO_NOVDEC+SWMY_ACCR,2)}\n"
       f"      Cr  Grant income — FCDO Phase II (restricted)        USD {f(FCDO_NOVDEC,2)}\n"
       f"      Cr  Grant income — Swiss Myanmar (restricted)        USD {f(SWMY_ACCR,2)}", 8.5)
F.para(f"Debits USD {f(FCDO_NOVDEC+SWMY_ACCR,2)} = credits USD {f(FCDO_NOVDEC+SWMY_ACCR,2)}. The entry is balanced. This would be a draft manual journal at 31 December 2026 and is reversed when the invoices are raised.", 8.5, italic=True)
P2 = os.path.join(OUT, "humanus_forecast-report_Oct26-Jun27.docx"); F.save(P2)
print("papers written")

# ---------------------------------------------------------------- WORKBOOK
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter as L
from openpyxl.chart import LineChart, Reference
wb_ = Workbook()
HF = PatternFill("solid", fgColor="410F37"); TF = PatternFill("solid", fgColor="EFE8EE"); GF = PatternFill("solid", fgColor="96FF14")
FN = "Fagun"
def hdr(ws, r, vals):
    for j,v in enumerate(vals,1):
        c=ws.cell(r,j,v); c.fill=HF; c.font=Font(name=FN,bold=True,color="96FF14"); c.alignment=Alignment(wrap_text=True,vertical="center")
def body(ws, r, vals, bold=False, fill=None, fmt='#,##0;(#,##0);"–"'):
    for j,v in enumerate(vals,1):
        c=ws.cell(r,j,v); c.font=Font(name=FN,bold=bold,color="410F37")
        if isinstance(v,(int,float)) or (isinstance(v,str) and v.startswith("=")): c.number_format=fmt
        if fill: c.fill=fill
def widths(ws, ws_):
    for j,w in enumerate(ws_,1): ws.column_dimensions[L(j)].width=w
def title(ws, t, sub):
    ws["A1"]=t; ws["A1"].font=Font(name=FN,bold=True,size=14,color="410F37")
    ws["A2"]=sub; ws["A2"].font=Font(name=FN,italic=True,size=9,color="410F37")

ws = wb_.active; ws.title="Cover"
ws.add_image(XLImage(LOGO), "A1"); ws.row_dimensions[1].height=60
ws["A6"]="Board finance workbook — Q3 2026"; ws["A6"].font=Font(name=FN,bold=True,size=18,color="410F37")
for i,(k,v) in enumerate([("Board meeting","29 October 2026"),("Period","to 30 September 2026; forecast Oct 2026 – Jun 2027"),
    ("Source","Xero, pulled live 12 Oct 2026 (report IDs on Sources tab)"),("Currency","USD"),
    ("Footer","Stichting hūmānus international · KvK 77641981 · [TBD: registered-charity footer boilerplate] · Confidential — board use only")],8):
    ws.cell(i,1,k).font=Font(name=FN,bold=True,color="410F37"); ws.cell(i,2,v).font=Font(name=FN,color="410F37")
widths(ws,[18,90])

ws = wb_.create_sheet("I&E"); title(ws,"Income & expenditure (USD, accrual)","Quarter columns = three-month totals; YTD = 1 Jan – 30 Sep 2026")
hdr(ws,4,["Account","Group","Q3 2026","Q2 2026","YTD 2026"]); r=5
for n,q3,q2,y in INCOME: body(ws,r,[n,"Income",q3,q2,y]); r+=1
body(ws,r,["Total income","",f"=SUM(C5:C{r-1})",f"=SUM(D5:D{r-1})",f"=SUM(E5:E{r-1})"],bold=True,fill=TF); ri=r; r+=1
s0=r
for n,g,q3,q2,y in PL: body(ws,r,[n,g,q3,q2,y],fmt='#,##0.00;(#,##0.00);"–"'); r+=1
body(ws,r,["Total expenditure","",f"=SUM(C{s0}:C{r-1})",f"=SUM(D{s0}:D{r-1})",f"=SUM(E{s0}:E{r-1})"],bold=True,fill=TF,fmt='#,##0.00;(#,##0.00)'); re_=r; r+=1
body(ws,r,["Net surplus / (deficit)","",f"=C{ri}-C{re_}",f"=D{ri}-D{re_}",f"=E{ri}-E{re_}"],bold=True,fill=TF,fmt='#,##0.00;(#,##0.00)')
widths(ws,[36,24,14,14,14])

ws = wb_.create_sheet("Balance sheet"); title(ws,"Balance sheet (USD)","Liabilities negative")
hdr(ws,4,["Line","30 Sep 2026","30 Jun 2026","31 Dec 2025"]); r=5
for k,v in BS.items(): body(ws,r,[k,*v],fmt='#,##0.00;(#,##0.00);"–"'); r+=1
body(ws,r,["Net assets","=SUM(B5:B%d)"%(r-1),"=SUM(C5:C%d)"%(r-1),"=SUM(D5:D%d)"%(r-1)],bold=True,fill=TF,fmt='#,##0.00;(#,##0.00)'); na=r; r+=1
body(ws,r,["Grant income earned, not yet invoiced (FCDO II 43,847.67 + Swiss Myanmar 2,464.00)",ACCR_SEP,None,None],fmt='#,##0.00'); r+=1
body(ws,r,["Adjusted net assets",f"=B{na}+B{r-1}",None,None],bold=True,fill=TF,fmt='#,##0.00')
widths(ws,[60,16,16,16])

ws = wb_.create_sheet("Funds"); title(ws,"2026 YTD movements by fund (USD)","Donors tracking; expenditure ex-FX, FX on its own line")
hdr(ws,4,["Fund","Type","Income","Expenditure","Transfers","Net","Note"]); r=5
for n,t,i_,e_,tr_,note in FUNDS: body(ws,r,[n,t,i_,e_,tr_,f"=C{r}-D{r}+E{r}",note],fmt='#,##0.00;(#,##0.00);"–"'); r+=1
body(ws,r,["Total","",f"=SUM(C5:C{r-1})",f"=SUM(D5:D{r-1})",f"=SUM(E5:E{r-1})",f"=SUM(F5:F{r-1})",""],bold=True,fill=TF,fmt='#,##0.00;(#,##0.00);"–"')
widths(ws,[44,13,14,14,14,14,70])

ws = wb_.create_sheet("Budget v actual"); title(ws,"Budget vs actual and full-year forecast (USD)","Monthly lines vs 9/12; envelopes vs full year")
hdr(ws,4,["Budget line","Type","Xero accounts","Working budget","Board budget","YTD basis (working)","Actual YTD","Variance","Forecast FY","FY vs working"]); r=5
for (name,accts,wbv,bbv,typ),(n2,t2,_,_,ytd,fy) in zip(BVA,bva_rows):
    body(ws,r,[name,{"M":"Monthly","L":"Envelope","U":"Unbudgeted"}[typ],", ".join(accts),wbv,bbv,
               (f"=D{r}*9/12" if typ=="M" else f"=D{r}"),round(ytd,2),f"=G{r}-F{r}",round(fy,2),f"=I{r}-D{r}"]); r+=1
body(ws,r,["FX revaluation","—","Foreign Currency Gains and Losses",0,0,0,FX_YTD,f"=G{r}-F{r}",FX_YTD,f"=I{r}-D{r}"]); r+=1
body(ws,r,["Total","","",*[f"=SUM({L(c)}5:{L(c)}{r-1})" for c in range(4,11)]],bold=True,fill=TF)
widths(ws,[40,12,44,13,13,14,13,13,13,13])

ws = wb_.create_sheet("Assumptions"); title(ws,"Cash forecast assumptions (USD per month)","Edit the yellow cells; Cash forecast tab recalculates")
hdr(ws,4,["Line"]+model.MONTHS); r=5; arow={}
YF = PatternFill("solid", fgColor="FFF6CC")
for ln in floor_lines:
    body(ws,r,[ln]+[round(model.floor(i)[ln],2) for i in range(9)]); arow[ln]=r
    for j in range(2,11): ws.cell(r,j).fill=YF
    r+=1
body(ws,r,["One-off committed bills"]+[round(sum(model.ONE_OFFS.get(i,{}).values()),2) for i in range(9)]); arow["one"]=r
for j in range(2,11): ws.cell(r,j).fill=YF
r+=2
ws.cell(r,1,"Programme & travel per month — selected scenario").font=Font(name=FN,bold=True,color="410F37"); ws.cell(r,2,5000).fill=YF; ws.cell(r,2).number_format='#,##0'; arow["var"]=r
ws.cell(r,3,"Central 5,000 · Freeze 0 · Q3 run-rate 10,636").font=Font(name=FN,italic=True,color="410F37"); r+=2
hdr(ws,r,["Inflow"]+model.MONTHS); r+=1; inrows=[]
for mi,l,a in model.INFLOWS["base"]:
    body(ws,r,[l]+[a if i==mi else 0 for i in range(9)]); inrows.append(r)
    for j in range(2,11): ws.cell(r,j).fill=YF
    r+=1
ws.cell(r+1,1,"One-offs: Oct = paid 1–12 Oct 2,103.64 + Chair stipend EUR 2,000 + Free Thought GBP 700; Apr 27 = 2026 audit fee ~3,000.").font=Font(name=FN,italic=True,size=9,color="410F37")
widths(ws,[52]+[12]*9)

ws = wb_.create_sheet("Cash forecast"); title(ws,"Monthly cash forecast — selected scenario (USD, month-end)","Linked to Assumptions; opening cash = Xero BS 30 Sep 2026")
hdr(ws,4,["Line"]+model.MONTHS); r=5
body(ws,r,["Opening cash"]+[model.OPEN_CASH if i==0 else f"={L(i+1)}{0}" for i in range(9)],bold=True); op=r; r+=1
fr_=r
for ln in floor_lines: body(ws,r,[ln]+[f"=Assumptions!{L(i+2)}{arow[ln]}" for i in range(9)]); r+=1
body(ws,r,["One-off committed bills"]+[f"=Assumptions!{L(i+2)}{arow['one']}" for i in range(9)]); r+=1
body(ws,r,["Programme & travel"]+[f"=Assumptions!$B${arow['var']}" for i in range(9)]); r+=1
body(ws,r,["Total cash out"]+[f"=SUM({L(i+2)}{fr_}:{L(i+2)}{r-1})" for i in range(9)],bold=True,fill=TF); to=r; r+=1
fi=r
for k,(mi,l,a) in enumerate(model.INFLOWS["base"]): body(ws,r,[l]+[f"=Assumptions!{L(i+2)}{inrows[k]}" for i in range(9)]); r+=1
body(ws,r,["Total cash in"]+[f"=SUM({L(i+2)}{fi}:{L(i+2)}{r-1})" for i in range(9)],bold=True,fill=TF); ti_=r; r+=1
body(ws,r,["Closing cash"]+[f"={L(i+2)}{op}+{L(i+2)}{ti_}-{L(i+2)}{to}" for i in range(9)],bold=True,fill=GF); cl=r
for i in range(1,9): ws.cell(op,i+2).value=f"={L(i+1)}{cl}"
from openpyxl.formatting.rule import CellIsRule
ws.conditional_formatting.add(f"B{cl}:J{cl}", CellIsRule(operator="lessThan", formula=["0"], font=Font(name=FN,bold=True,color="E12D64")))
widths(ws,[52]+[12]*9)
ch = LineChart(); ch.title="Month-end cash (selected scenario)"; ch.height=7; ch.width=18
ch.add_data(Reference(ws,min_col=1,max_col=10,min_row=cl,max_row=cl),from_rows=True,titles_from_data=True)
ch.set_categories(Reference(ws,min_col=2,max_col=10,min_row=4)); ws.add_chart(ch, f"A{cl+3}")

ws = wb_.create_sheet("Scenarios"); title(ws,"Scenario month-end closing cash (USD)","Values from the build model, 12 Oct 2026")
hdr(ws,4,["Scenario"]+model.MONTHS); r=5
for k,v in SC.items(): body(ws,r,[k]+[round(x['close'],0) for x in v]); r+=1
ws.conditional_formatting.add(f"B5:J{r-1}", CellIsRule(operator="lessThan", formula=["0"], font=Font(name=FN,bold=True,color="E12D64")))
widths(ws,[44]+[12]*9)

ws = wb_.create_sheet("Grants"); title(ws,"Contract-to-Xero reconciliation","Contract currency in the Amount column")
hdr(ws,4,["Contracted payment","Due","Amount (contract ccy)","Xero ref","Received (USD)","Status"]); r=5
for row in g_rows: body(ws,r,list(row)); r+=1
widths(ws,[22,30,24,26,14,50])

ws = wb_.create_sheet("Sources"); title(ws,"Sources","")
src = ["Xero P&L Q3 2026 — 64172c20-ed61-46f3-9b5d-77406b0ae5d9","Xero P&L YTD 2026 — 9aa8308a-7cab-41fb-8541-d6b452021fd7",
 "Xero P&L by quarter — 16b22d31-27b6-4f0e-80eb-d39cac34a238","Xero P&L Q2 direct (Donors) — 3aeb792c-3bc8-47d6-be9b-eeba99c00d1e",
 "Xero P&L by Donors YTD — 13b97ef5-027d-4984-af8c-50616727cf97","Xero P&L by Donors Q3 — 7d6a8730-9816-4408-a7c7-a85669c0b742",
 "Xero Balance Sheet 30 Sep / 30 Jun — 052baafa-a222-4e25-9853-eb0c4c9668ef","Xero Balance Sheet 31 Dec 2025 — cfaec7b2-db33-46e2-ae1b-0ddba2fb0304",
 "Xero Bank Summary Q3 — 6d5984f7-e1f8-4989-ad4b-385df2de97fa","Xero Bank Summary 1–12 Oct — f55f65a4-66de-4df1-8172-0d34425088fe",
 "Xero repeating invoices — bd24128f; open ACCPAY — 911256c0; ACCREC — 85561f8e",
 "Budgets: finance-docs/budgets/2026 Budget-2.xlsx (working 488,450); humanusGovernance/2026 Budget.xlsx (board 631,450)",
 "Contracts: OSF OR2024-94840; FCDO II sub-grant 2026.06.24; Swiss Myanmar Partnership Agreement 2026.02.04; SDC 7F-03317.09.15 (scan)"]
for i,sx in enumerate(src,4): ws.cell(i,1,sx).font=Font(name=FN,color="410F37")
widths(ws,[120])
P3 = os.path.join(OUT, "humanus_board-finance-workbook_Q3-2026.xlsx"); wb_.save(P3)
print("workbook written")
