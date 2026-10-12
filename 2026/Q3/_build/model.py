"""hūmānus Q3 2026 board pack — data + forecast model. All Xero figures pulled 12 Oct 2026."""
MONTHS = ["Oct 26","Nov 26","Dec 26","Jan 27","Feb 27","Mar 27","Apr 27","May 27","Jun 27"]
OPEN_CASH = 33679.30          # Xero BS 30 Sep 2026 (052baafa)
GBPUSD = 1.3258               # Xero rate on COLA payment 28 Sep (0.75429)
EURUSD = 1.1255               # Xero rate on Els bill (0.888497)

def floor(i):
    """Committed monthly cash floor (repeating templates + manual Comms Advisor + Q3 avg overheads)."""
    das = 2000 if i == 0 else 1000                       # Eva decision CAP-485: 1,000 from Nov (date to confirm)
    shafi = 1500 if i < 2 else (2500 - 64.97 - 1000 + 1000 if i == 2 else 2500)  # advance 2,064.97 recovers 1,000/mo
    if i == 2: shafi = 2500 - 64.97
    lines = {
        "Executive Director": 14200.0,
        "Communication Advisor": 5383.27,    # Q3 avg 16,149.82/3 (manual invoice)
        "Advocate / Programme Manager": 3500.0,
        "Rohingya Manager (cash)": shafi,
        "Cost of Living Allowance (GBP 1,400)": round(1400*GBPUSD,2),
        "Finance Manager (SK Das)": das,
        "Communication Support": 649.41,     # Q3 avg 1,948.22/3
        "IT & software": 352.70,             # Q3 avg 1,058.11/3
        "Bank charges": 55.74,               # Q3 avg 167.22/3
    }
    if i >= 8: lines["Rohingya Manager (cash)"] = 0.0   # Swiss Myanmar term ends 31 May 2027; template ends May 27
    return lines

ONE_OFFS = {  # cash, USD
    0: {"Paid 1–12 Oct (incl. SK Das Sep fee 2,000)": 2103.64,
        "Els Niehaus bill EUR 2,000 (SUBMITTED, due 15 Oct)": round(2000*EURUSD,2),
        "Free Thought Ltd GBP 700 (SUBMITTED, due 15 Oct)": round(700*GBPUSD,2)},
    6: {"2026 statutory audit fee (deferred to 2027)": 3000.0},
}
VARIABLE = {"Freeze": 0.0, "Central": 5000.0, "Q3 run-rate": 10636.0}   # programme/travel layer per month

# inflows: (month index, label, amount, scenario set)
INFLOWS = {
  "base": [
    (1, "FCDO Phase II claim Jun–Oct 2026 (Schedule 2, ind. Nov 2026)", 52800.0),
    (3, "OSF #4 Year-3 tranche (contract: January 2027)", 300000.0),
    (3, "Swiss Myanmar tranche 2 (§3.3: 31 Jan 2027)", 24946.67),
    (7, "FCDO Phase II final claim Nov 26–Mar 27 (ind. May 2027)", 47000.0),
    (8, "Swiss Myanmar tranche 3 (§3.3: 30 Jun 2027)", 24946.66),
  ],
}

def run(variable, inflows, shift=None):
    shift = shift or {}
    cash = OPEN_CASH; rows = []
    for i, m in enumerate(MONTHS):
        fl = sum(floor(i).values()); one = sum(ONE_OFFS.get(i, {}).values()); var = variable
        inc = sum(a for (mi, l, a) in inflows if shift.get(l, mi) == i)
        out = fl + one + var
        cash = cash + inc - out
        rows.append(dict(month=m, floor=fl, oneoff=one, variable=var, out=out, inflow=inc, close=cash))
    return rows

def scen():
    base = INFLOWS["base"]
    res = {}
    res["A. Central"] = run(VARIABLE["Central"], base)
    res["B. Freeze (committed only)"] = run(VARIABLE["Freeze"], base)
    res["C. Q3 run-rate continues"] = run(VARIABLE["Q3 run-rate"], base)
    lab_fcdo = base[0][1]
    res["D. Central, FCDO claim slips to Jan 27"] = run(VARIABLE["Central"], base, {lab_fcdo: 3})
    res["E. Central + OSF released early Dec 26"] = run(VARIABLE["Central"], base, {base[1][1]: 2})
    res["F. Central + Swiss T2 pulled to Dec 26"] = run(VARIABLE["Central"], base, {base[2][1]: 2})
    return res

if __name__ == "__main__":
    for k, rows in scen().items():
        print(k)
        for r in rows: print(f"  {r['month']}: out {r['out']:10,.0f} in {r['inflow']:10,.0f} close {r['close']:11,.0f}")
