"""Xero actuals pulled 12 Oct 2026 (report IDs in SOURCES)."""
PL = [  # account, group, Q3, Q2, YTD
 ("Executive Director","Staff & consultants",42600.00,42600.00,127800.00),
 ("Communication Advisor","Staff & consultants",16149.82,16085.64,48461.94),
 ("Advocate / Programme Manager","Staff & consultants",10500.00,10500.00,34000.00),
 ("Rohingya Manager","Staff & consultants",7500.00,7500.00,22500.00),
 ("Cost of Living Allowances","Staff & consultants",5630.35,5625.05,16931.01),
 ("Accounting & Bookkeeping Fees","Staff & consultants",6000.00,6000.00,18000.00),
 ("Communication Support","Staff & consultants",1948.22,5556.31,11866.75),
 ("Operations Support","Staff & consultants",1390.72,638.14,6266.90),
 ("Project Consultants","Staff & consultants",2293.57,0.00,2293.57),
 ("Volunteer Stipends","Programme & travel",178.45,10268.63,12334.00),
 ("Partner Grants","Programme & travel",4000.00,8000.00,14000.00),
 ("Confidence Building Measures","Programme & travel",0.00,0.00,14598.00),
 ("International Flights","Programme & travel",4997.80,20519.49,25517.29),
 ("Accommodation","Programme & travel",8893.66,639.28,11247.69),
 ("Domestic Flights","Programme & travel",849.99,274.42,1124.41),
 ("Per diem","Programme & travel",500.21,606.53,1288.78),
 ("Local Transportation","Programme & travel",188.38,1026.82,2286.22),
 ("Meeting Expenses","Programme & travel",5701.24,1616.62,8930.14),
 ("Board Meetings","Governance & operations",0.00,8936.41,8936.41),
 ("Audit Fees","Governance & operations",0.00,1981.06,2822.06),
 ("Legal Fees","Governance & operations",2914.57,818.27,4363.31),
 ("IT & Software Expenses","Governance & operations",1058.11,2636.00,5469.26),
 ("Bank Charges & Merchant Fees","Governance & operations",167.22,214.31,510.63),
 ("Depreciation","Governance & operations",0.00,0.00,113.84),
 ("Foreign Currency Gains and Losses","FX revaluation",79.65,-221.56,235.02),
]
INCOME = [("Grant income",0.00,114081.47,414081.47),("Interest income",431.74,695.31,1599.15)]
EXP_TOT = {"Q3":123541.96,"Q2":151821.42,"YTD":401897.23}
INC_TOT = {"Q3":431.74,"Q2":114776.78,"YTD":415680.62}
NET = {"Q3":-123110.22,"Q2":-37044.64,"YTD":13783.39}

BS = {  # 30 Sep 26, 30 Jun 26, 31 Dec 25
 "Wise – USD":(4404.33,90275.20,292.77),
 "Wise – EUR (USD equiv.)":(29274.97,8550.08,23493.98),
 "Wise – GBP":(0.00,0.00,0.00),
 "Accounts receivable":(0.00,53135.45,0.00),
 "Advance payments (staff advance)":(2064.97,5064.97,64.97),
 "Prepayments":(213.00,244.95,308.85),
 "Fixed assets (net)":(3512.33,3512.33,1733.89),
 "Accounts payable":(-1914.00,-113.77,-1271.76),
 "Expenses payable":(-0.11,-0.11,-829.11),
 "Suspense":(21.50,18.10,0.00),
}
NET_ASSETS = (37576.99,160687.20,23793.59)

# 2026 YTD by fund (ex-FX expenditure), Donors tracking P&L 13b97ef5
FUNDS = [
 # name, type, income, expenditure ex-FX, note
 ("OSF #4 general support (2025–29)","Unrestricted",300023.70,203699.50,"Year-2 tranche received Jan 2026"),
 ("Reserves (interest)","Unrestricted",1575.45,0.00,""),
 ("FX revaluation (all codes)","Unrestricted",0.00,235.04,"non-cash"),
 ("Legacy / unassigned codes","Unrestricted",0.00,42.38,"OSF #1 9.36 + untagged 33.02"),
 ("FCDO Phase I (closed Mar 2026)","Restricted",61005.00,52872.10,"2026 receipts also fund 2025 spend of 8,132.90 — lifetime nil"),
 ("FCDO Phase II (Jun 26 – Mar 27)","Restricted",0.00,43847.67,"reimbursement grant — claim not yet invoiced"),
 ("SDC Rohingya Women's Dialogue (closed Apr 2026)","Restricted",28130.47,73790.56,"final instalment; Art. 1.2 50% co-funding falls on core"),
 ("Swiss Myanmar (Dec 25 – May 27)","Restricted",24946.00,27410.00,"tranche 2 due 31 Jan 2027"),
]

CASH_MONTHS = [  # month-end, total, USD, EUR, GBP
 ("31 Dec 2025",23786.75,292.77,23493.98,0.00),
 ("30 Jun 2026",98825.28,90275.20,8550.08,0.00),
 ("31 Jul 2026",115335.08,80138.31,30656.83,4539.94),
 ("31 Aug 2026",76112.64,45601.90,30510.74,0.00),
 ("30 Sep 2026",33679.30,4404.33,29274.97,0.00),
 ("12 Oct 2026 (live)",31328.83,2456.42,28872.41,0.00),
]

# Budget mapping: budget line -> (xero accounts, working budget, board budget, type)
BVA = [
 ("Executive Director",["Executive Director"],170400,170400,"M"),
 ("COLA",["Cost of Living Allowances"],21600,21600,"M"),
 ("Senior Legal Advisor",[],0,120000,"M"),
 ("Programme Manager",["Advocate / Programme Manager"],42000,42000,"M"),
 ("Project Manager (Rohingya Manager)",["Rohingya Manager"],31250,31250,"M"),
 ("Communications Advisor",["Communication Advisor"],64800,64800,"M"),
 ("Communication Support",["Communication Support"],14400,14400,"M"),
 ("Finance Manager",["Accounting & Bookkeeping Fees"],24000,24000,"M"),
 ("Operations Support",["Operations Support"],18000,18000,"M"),
 ("Vehicle Hire & Sundry",["Local Transportation"],6000,6000,"M"),
 ("Volunteer Stipends",["Volunteer Stipends"],12000,12000,"L"),
 ("Partner Expenses (Azadi)",["Partner Grants"],30000,30000,"L"),
 ("Lawyer Fees",["Legal Fees"],12000,24000,"L"),
 ("Audit",["Audit Fees"],3000,3000,"L"),
 ("Board Meetings",["Board Meetings"],4000,4000,"L"),
 ("Technology & Software",["IT & Software Expenses"],5000,5000,"L"),
 ("Travel Expenses",["International Flights","Domestic Flights","Accommodation","Per diem"],25000,36000,"L"),
 ("Fees (bank)",["Bank Charges & Merchant Fees"],5000,5000,"L"),
 ("Unbudgeted — Confidence Building Measures (FCDO I funded)",["Confidence Building Measures"],0,0,"U"),
 ("Unbudgeted — Meeting Expenses",["Meeting Expenses"],0,0,"U"),
 ("Unbudgeted — Project Consultants",["Project Consultants"],0,0,"U"),
 ("Unbudgeted — Depreciation",["Depreciation"],0,0,"U"),
]
