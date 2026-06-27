# -*- coding: utf-8 -*-
"""
Ather Energy Limited - Integrated 7/8-Year Financial Model
Forecast: FY2026 - FY2033 (8 years) | History: FY2021 - FY2025
Built to replicate the United Spirits modelling template (structure, colour coding, philosophy).
Colour coding: BLUE = hardcoded input | GREEN = link/reference | BLACK = calculation
"""
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string

# ---------- style constants ----------
BLUE  = "FF0070C0"   # hardcoded inputs
GREEN = "FF00B050"   # links / references
BLACK = "FF000000"   # calculations / labels
GREY  = "FF7F7F7F"
WHITE = "FFFFFFFF"
NAVY  = "FF1F3864"
LBLUE = "FFD9E1F2"   # light header fill
LGREEN= "FFE2EFDA"
LGREY = "FFF2F2F2"
LYEL  = "FFFFF2CC"

MONEY = '#,##0_);(#,##0)'
MONEY2= '0.00_);(0.00)'
PCT   = '0.0%'
PCT0  = '0%'
PCT2  = '0.00%'
NUM0  = '#,##0'
NUM2  = '0.00'
DAYS  = '0'
MULT  = '0.0"x"'
YEAR_A= '0"A"'
YEAR_E= '0"E"'
PRICE = '#,##0'

thin = Side(style="thin", color="FFBFBFBF")
BORDER_BOTTOM = Border(bottom=Side(style="thin", color="FF808080"))
BORDER_TOPBOT = Border(top=Side(style="thin", color="FF808080"), bottom=Side(style="thin", color="FF808080"))

HCOLS = ["C","D","E","F","G"]                       # FY2021..FY2025
FCOLS = ["H","I","J","K","L","M","N","O"]           # FY2026..FY2033
ALL   = HCOLS + FCOLS
FIRST_YEAR = 2021

# sheet names
COV="Cover"; SC="Scenarios"; ASm="Assumptions"; RB="Revenue Build"; PLs="PL"; BSs="BS"; CSs="CS"
WCS="Working Capital Schedule"; WS="Working Schedules"; RA="Ratio Analysis"; DCFs="DCF"
EC="Error Checks"; CM="Company & Market Insights"; MN="Modelling Notes"

def ref(sheet, cell):
    return f"'{sheet}'!{cell}" if " " in sheet or "&" in sheet else f"{sheet}!{cell}"

def col_shift(col, n):
    return get_column_letter(column_index_from_string(col) + n)

wb = openpyxl.Workbook()

def put(ws, coord, value, kind="calc", nf=None, bold=False, size=11, italic=False,
        fill=None, align=None, wrap=False, color=None):
    c = ws[coord]
    c.value = value
    col = {"input":BLUE, "link":GREEN, "calc":BLACK, "label":BLACK, "header":BLACK}.get(kind, BLACK)
    if color: col = color
    c.font = Font(name="Calibri", size=size, bold=bold, italic=italic, color=col)
    if nf: c.number_format = nf
    if fill: c.fill = PatternFill("solid", fgColor=fill)
    a = {}
    if align: a["horizontal"] = align
    if wrap: a["wrap_text"] = True
    a["vertical"] = "center"
    c.alignment = Alignment(**a)
    return c

def title(ws, text, coord="B2"):
    put(ws, coord, text, kind="label", bold=True, size=16, color=NAVY)

def section(ws, row, text, col="B", span_to="O"):
    put(ws, f"{col}{row}", text, kind="label", bold=True, size=12, color=NAVY, fill=LBLUE)
    # fill across
    for cc in range(column_index_from_string(col)+1, column_index_from_string(span_to)+1):
        ws[f"{get_column_letter(cc)}{row}"].fill = PatternFill("solid", fgColor=LBLUE)

def year_header(ws, row, label="Particulars", hist=HCOLS, fcst=FCOLS, label_col="B"):
    put(ws, f"{label_col}{row}", label, kind="label", bold=True, fill=LGREY)
    cols = hist + fcst
    for i, cc in enumerate(cols):
        if i == 0:
            put(ws, f"{cc}{row}", FIRST_YEAR, kind="label", bold=True, nf=(YEAR_A if cc in hist else YEAR_E),
                align="center", fill=LGREY)
        else:
            prev = cols[i-1]
            put(ws, f"{cc}{row}", f"={prev}{row}+1", kind="label", bold=True,
                nf=(YEAR_A if cc in hist else YEAR_E), align="center", fill=LGREY)

def colwidths(ws, mapping):
    for c, w in mapping.items():
        ws.column_dimensions[c].width = w

def write_row(ws, row, label, values_by_col, kind="input", nf=MONEY, bold=False, indent=False):
    """values_by_col: dict col->value (number or formula string)."""
    lab = ("   " + label) if indent else label
    put(ws, f"B{row}", lab, kind="label", bold=bold)
    for cc, v in values_by_col.items():
        k = kind
        if isinstance(v, str) and v.startswith("="):
            # decide link vs calc by presence of '!'
            k = "link" if "!" in v else "calc"
        put(ws, f"{cc}{row}", v, kind=k, nf=nf, bold=bold)

# =====================================================================================
#  SCENARIOS
# =====================================================================================
ws = wb.active; ws.title = SC
colwidths(ws, {"A":2.5,"B":40,"C":12})
title(ws, "Ather Energy Ltd — Economic / Operating Scenarios", "B2")
put(ws, "B4", "Scenario Switch  (1 = Base, 2 = Bull, 3 = Bear)", kind="label", bold=True)
put(ws, "F4", 1, kind="input", nf="0", bold=True, fill=LYEL, align="center")
put(ws, "B5", "Currently running:", kind="label", italic=True)
put(ws, "F5", '=CHOOSE($F$4,"BASE","BULL","BEAR")', kind="calc", bold=True, align="center", color=NAVY)

# Volume growth block
def scen_block(ws, top, name, base, best, bear, nf=PCT):
    section(ws, top, name)
    year_header(ws, top+1, label="")
    put(ws, f"B{top+2}", "Selected (driven by switch)", kind="label", bold=True)
    for cc in FCOLS:
        put(ws, f"{cc}{top+2}", f"=CHOOSE($F$4,{cc}{top+4},{cc}{top+5},{cc}{top+6})", kind="calc", nf=nf, bold=True)
    for j,(rowname,arr) in enumerate([("Base",base),("Bull",best),("Bear",bear)]):
        put(ws, f"B{top+4+j}", rowname, kind="label")
        for i,cc in enumerate(FCOLS):
            put(ws, f"{cc}{top+4+j}", arr[i], kind="input", nf=nf)

vol_base=[0.28,0.26,0.24,0.22,0.18,0.15,0.12,0.10]
vol_best=[0.35,0.32,0.28,0.25,0.22,0.18,0.15,0.12]
vol_bear=[0.18,0.16,0.14,0.12,0.10,0.08,0.07,0.06]
scen_block(ws, 7, "Vehicle Unit Volume Growth %", vol_base, vol_best, vol_bear)

gm_base=[0.18,0.20,0.22,0.24,0.255,0.265,0.275,0.28]
gm_best=[0.20,0.225,0.245,0.265,0.28,0.29,0.30,0.305]
gm_bear=[0.155,0.165,0.18,0.19,0.20,0.21,0.22,0.225]
scen_block(ws, 16, "Gross Margin %", gm_base, gm_best, gm_bear)

put(ws, "B26", "Note: Volume growth feeds the Revenue Build; Gross Margin feeds the P&L COGS line. "
                "Switch one cell (F4) to flex the entire model across Base / Bull / Bear cases.",
    kind="label", italic=True, color=GREY)
print("Scenarios done")


# =====================================================================================
#  ASSUMPTIONS
# =====================================================================================
ws = wb.create_sheet(ASm)
colwidths(ws, {"A":2.5,"B":44,"C":9,"D":9,"E":9,"F":9,"G":9,"H":9,"I":9,"J":9,"K":9,"L":9,"M":9,"N":9,"O":9,"P":2,"Q":34,"R":62})
title(ws, "Ather Energy Ltd — Assumptions & Key Drivers", "B2")
put(ws, "B3", "All monetary figures in INR millions unless stated. Forecast FY2026E–FY2033E.", kind="label", italic=True, color=GREY)
year_header(ws, 4)
put(ws, "Q4", "Source", kind="label", bold=True, fill=LGREY)
put(ws, "R4", "Reason / Simple explanation", kind="label", bold=True, fill=LGREY)

# assumption registry: row -> (label, list_of_8_values_or_formula, nf, source, reason, kind)
A = {}  # name -> row
def arow(row, name, label, vals, nf, source, reason, kind="input"):
    A[name] = row
    put(ws, f"B{row}", label, kind="label")
    for i,cc in enumerate(FCOLS):
        v = vals[i] if isinstance(vals, list) else vals
        k = kind
        if isinstance(v,str) and v.startswith("="):
            k = "link" if "!" in v else "calc"
        put(ws, f"{cc}{row}", v, kind=k, nf=nf)
    put(ws, f"Q{row}", source, kind="label", size=9, wrap=True, color=GREY)
    put(ws, f"R{row}", reason, kind="label", size=9, wrap=True, color=GREY)

const = lambda x: [x]*8

section(ws, 5, "A. Macroeconomic & Industry Assumptions")
arow(6,  "gdp",   "India real GDP growth %", [0.065,0.065,0.067,0.068,0.068,0.065,0.065,0.065], PCT, "IMF/RBI projections", "Stable ~6.5% growth supports consumer discretionary demand for two-wheelers.")
arow(7,  "infl",  "CPI inflation %", const(0.045), PCT, "RBI target band (4%±2%)", "Moderate inflation; input-cost pressure largely offset by battery cost decline.")
arow(8,  "repo",  "RBI repo rate %", [0.06,0.0575,0.0575,0.0575,0.06,0.06,0.06,0.06], PCT, "RBI policy outlook", "Gradual easing then stable; lowers cost of borrowings & leases.")
arow(9,  "ind2w", "India 2W industry (m units)", [19.5,20.8,22.0,23.3,24.5,25.7,26.8,27.9], NUM2, "SIAM; ~8% CAGR", "Overall 2W market grows ~8%; base for E2W penetration math.")
arow(10, "pen",   "E2W penetration of 2W %", [0.075,0.095,0.12,0.15,0.18,0.22,0.26,0.30], PCT, "Kearney (30% by 2030); Ather (40% by FY31)", "Electric share of 2W rises sharply on cost parity, charging, policy.")
arow(11, "e2wgr", "E2W industry growth %", [0.25,0.28,0.30,0.32,0.28,0.26,0.22,0.18], PCT, "IMARC/Renub/TechSci (26-28% CAGR)", "Strong structural growth in electric two-wheelers.")
arow(12, "batt",  "Li-ion battery pack (USD/kWh)", [105,99,95,91,88,85,83,81], NUM0, "BloombergNEF Dec-2025 ($108 in 2025)", "Falling cells/lithium prices cut Ather's largest input cost -> margin tailwind.")
arow(13, "share", "Ather E2W market share %", [0.16,0.17,0.18,0.19,0.20,0.205,0.21,0.21], PCT, "Vahan; ETAuto (18.7% Mar-26)", "Share gains from capacity, products (Rizta), 700+ outlets; tapers as market matures.")

section(ws, 15, "B. Revenue Drivers")
arow(16, "volg", "Vehicle unit volume growth %", [f"={ref(SC,cc+'9')}" for cc in FCOLS], PCT,
     "Scenarios sheet (driver)", "Driven by Base/Bull/Bear switch; reflects share gains + market growth, tapering over time.")
arow(17, "aspg", "Realisation / ASP growth %", [-0.01,0.0,0.01,0.01,0.015,0.02,0.02,0.02], PCT,
     "Ather ASP: FY23 155.6k, FY24 143.3k, FY25 ~127.6k", "Near-term dip from cheaper Rizta mix, then modest premiumisation/price recovery.")
arow(18, "acc",  "Accessories revenue % of vehicle rev", const(0.025), PCT, "AR Note 20; mgmt commentary", "Helmets, chargers, spares attach steadily with deliveries.")
arow(19, "soft", "Software & connectivity % of vehicle rev", [0.015,0.018,0.02,0.022,0.025,0.027,0.03,0.03], PCT, "AtherStack/Pro-Pack adoption", "High-margin subscriptions rise with connected installed base.")
arow(20, "chg",  "Charging (Ather Grid) % of vehicle rev", const(0.005), PCT, "Ather Grid network", "Small but growing recurring charging revenue.")
arow(21, "svcg", "Service & after-sales growth %", [0.30,0.28,0.26,0.24,0.22,0.20,0.18,0.16], PCT, "AR Note 20 (service rev 1,451m FY25)", "Grows with cumulative fleet on road; tapers as base enlarges.")
arow(22, "othop","Other operating income (INR m)", const(50), MONEY, "Historical run-rate", "Scrap, incentives and sundry operating income held broadly flat.")

section(ws, 24, "C. Cost & Margin Drivers")
arow(25, "gm",   "Gross margin %", [f"={ref(SC,cc+'18')}" for cc in FCOLS], PCT,
     "Scenarios sheet (driver)", "Expands on battery-cost decline, scale, localisation/PLI and mix; Base 18%->28%.")
arow(26, "emp",  "Employee cost % of sales", [0.15,0.13,0.115,0.105,0.10,0.095,0.09,0.09], PCT,
     "FY25 18.3% of sales", "Operating leverage: headcount grows slower than revenue.")
arow(27, "oopex","Other operating expenses % of sales", [0.21,0.185,0.165,0.15,0.14,0.135,0.13,0.13], PCT,
     "FY25 24.2% of sales (S&M, R&D-opex, admin, freight)", "Marketing/logistics/admin scale sub-linearly; brand now established.")
arow(28, "warr", "Warranty provision % of sales (non-cash)", const(0.027), PCT, "FY25 warranty 602m", "Non-cash provision added back in cash flow.")
arow(29, "sbp",  "Share-based payments % of employee cost", const(0.18), PCT, "FY25 SBP 831m vs emp 4,124m", "Non-cash ESOP charge added back in cash flow.")

section(ws, 31, "D. Capital Expenditure, Depreciation & Amortisation")
arow(32, "ppxc", "PPE capex % of revenue", [0.09,0.07,0.05,0.045,0.04,0.038,0.035,0.035], PCT,
     "Factory-3 (Maharashtra) IPO use-of-proceeds", "Front-loaded capex for capacity tripling, then normalises to maintenance levels.")
arow(33, "life", "PPE useful life (years)", const(8), DAYS, "Auto plant & equipment", "Straight-line depreciation base for PPE.")
arow(34, "intc", "Intangible additions % of revenue", [0.04,0.038,0.035,0.032,0.03,0.028,0.025,0.025], PCT,
     "Capitalised product/software development", "R&D capitalised for new platforms; eases as portfolio matures.")
arow(35, "amrt", "Intangible amortisation % of opening", const(0.18), PCT, "~5-6 yr life", "Tech/software amortised over useful life.")
arow(36, "rouc", "ROU (lease) additions % of revenue", [0.03,0.028,0.025,0.022,0.02,0.018,0.016,0.015], PCT,
     "Retail (700+ outlets) & warehouse leases", "New experience-centre/warehouse leases scale with network, then taper.")
arow(37, "roua", "ROU amortisation %", const(0.22), PCT, "Avg lease ~4-5 yrs", "Straight-line amortisation of right-of-use assets.")

section(ws, 39, "E. Financing & Tax")
arow(40, "bint", "Interest rate on borrowings %", const(0.10), PCT, "FY25 finance cost / debt", "Blended rate on working-capital + term debt.")
arow(41, "bnet", "Net borrowing draw/(repay) (INR m)", [1500,500,1500,2000,1500,500,-1000,-1500], MONEY,
     "Mgmt: working-capital + term debt for capacity build", "Modest debt draws fund the capex/loss trough (FY26-FY30), repaid once free cash flow turns positive (FY32-FY33).")
arow(42, "lint", "Lease interest rate %", const(0.09), PCT, "Ind AS 116 incremental borrowing rate", "Interest unwound on lease liabilities.")
arow(43, "lrep", "Lease principal repayment growth %", const(0.10), PCT, "Historical lease repayment trend", "Lease principal outflow grows with expanding lease base.")
arow(44, "iinc", "Interest income rate on cash & inv %", const(0.06), PCT, "Yield on deposits/MFs", "Earned on opening cash, term deposits and liquid investments.")
arow(45, "eqis", "Equity issuance / IPO (INR m)", [26000,0,0,0,0,0,0,0], MONEY,
     "IPO May-2025 fresh issue ~INR 2,626 cr", "One-off FY26 primary equity raise (Factory-3, R&D, debt, marketing).")
arow(46, "divp", "Dividend payout %", const(0.0), PCT, "Loss-making; reinvestment phase", "No dividends expected over the forecast horizon.")
arow(47, "tax",  "Effective tax rate %", const(0.0), PCT, "Accumulated losses > INR 40,000m", "Carry-forward losses shield taxable profits for the whole forecast; statutory 25.17% (noted).")

section(ws, 49, "F. Working Capital")
arow(50, "rdays","Receivable days", const(4), DAYS, "FY25 AR 118m on rev 22,550m (~2 days)", "Mostly cash/dealer-financed sales -> minimal receivables.")
arow(51, "idays","Inventory days (on COGS)", [48,46,44,42,40,40,40,40], DAYS, "FY25 ~48 days", "Improves with better S&OP and faster throughput at scale.")
arow(52, "pdays","Payable days (on COGS)", [100,98,95,92,90,88,85,85], DAYS, "FY25 ~109 days", "High supplier credit normalises modestly as scale/terms mature.")
arow(53, "ocas", "Other current assets % of sales", const(0.14), PCT, "FY25 15.5%", "GST/balances/advances scale with revenue.")
arow(54, "ocls", "Other current liabilities % of sales", const(0.04), PCT, "FY25 3.9%", "Customer advances/deferred revenue scale with revenue.")
arow(55, "prov", "Current provisions % of sales", const(0.05), PCT, "FY25 5.3% (incl warranty)", "Warranty & employee provisions scale with revenue.")

print("Assumptions done")


# =====================================================================================
#  REVENUE BUILD (driver-based)
# =====================================================================================
ws = wb.create_sheet(RB)
colwidths(ws, {"A":2.5,"B":40,**{c:11 for c in ALL}})
title(ws, "Ather Energy Ltd — Revenue Build (driver-based)", "B2")
put(ws, "B3", "Units in numbers; ASP in INR/vehicle; all revenue in INR millions.", kind="label", italic=True, color=GREY)
year_header(ws, 4)

def hf_row(ws, row, label, hist_vals, fcst_formulas, nf=MONEY, bold=False, hist_kind="input"):
    put(ws, f"B{row}", label, kind="label", bold=bold)
    for i,cc in enumerate(HCOLS):
        put(ws, f"{cc}{row}", hist_vals[i], kind=hist_kind, nf=nf, bold=bold)
    for i,cc in enumerate(FCOLS):
        f = fcst_formulas[i]
        k = "link" if ("!" in f) else "calc"
        put(ws, f"{cc}{row}", f, kind=k, nf=nf, bold=bold)

# volumes
hf_row(ws, 6, "Vehicle volumes (units)",
       [7100,26000,106000,113900,165000],
       [f"={col_shift(cc,-1)}6*(1+{ref(ASm,cc+'16')})" for cc in FCOLS], nf=NUM0)
# ASP
hf_row(ws, 7, "Average Selling Price (INR / vehicle)",
       [98592,138462,154717,143292,127606],
       [f"={col_shift(cc,-1)}7*(1+{ref(ASm,cc+'17')})" for cc in FCOLS], nf=PRICE)
# vehicle revenue
put(ws, "B8", "Vehicle revenue", kind="label", bold=True)
for cc in ALL:
    put(ws, f"{cc}8", f"={cc}6*{cc}7/1000000", kind="calc", nf=MONEY, bold=True)
# accessories
hf_row(ws, 9, "Accessories & spares revenue",
       [18,90,250,0,0],
       [f"={cc}8*{ref(ASm,cc+'18')}" for cc in FCOLS], nf=MONEY)
# software
hf_row(ws,10, "Software & connectivity revenue",
       [5,20,80,0,0],
       [f"={cc}8*{ref(ASm,cc+'19')}" for cc in FCOLS], nf=MONEY)
# charging
hf_row(ws,11, "Charging (Ather Grid) revenue",
       [0,5,20,0,0],
       [f"={cc}8*{ref(ASm,cc+'20')}" for cc in FCOLS], nf=MONEY)
# service
hf_row(ws,12, "Service & after-sales revenue",
       [55,320,1000,1190,1451],
       [f"={col_shift(cc,-1)}12*(1+{ref(ASm,cc+'21')})" for cc in FCOLS], nf=MONEY)
# other operating income
hf_row(ws,13, "Other operating income",
       [20,50,86,27,44],
       [f"={ref(ASm,cc+'22')}" for cc in FCOLS], nf=MONEY)
# total
put(ws, "B14", "Total revenue from operations", kind="label", bold=True)
for cc in ALL:
    put(ws, f"{cc}14", f"=SUM({cc}8:{cc}13)", kind="calc", nf=MONEY, bold=True)
ws[f"B14"].border = BORDER_TOPBOT
for cc in ALL: ws[f"{cc}14"].border = BORDER_TOPBOT

# memo: growth & mix
put(ws,"B16","Memo:",kind="label",italic=True,color=GREY)
put(ws,"B17","Total revenue growth %",kind="label")
for i,cc in enumerate(ALL):
    if i==0:
        put(ws,f"{cc}17","",kind="calc",nf=PCT)
    else:
        put(ws,f"{cc}17",f"={cc}14/{col_shift(cc,-1)}14-1",kind="calc",nf=PCT)
put(ws,"B18","Non-vehicle revenue % of total",kind="label")
for cc in ALL:
    put(ws,f"{cc}18",f"=({cc}14-{cc}8)/{cc}14",kind="calc",nf=PCT)
print("Revenue Build done")


# =====================================================================================
#  P&L  (Income Statement)
# =====================================================================================
ws = wb.create_sheet(PLs)
colwidths(ws, {"A":2.5,"B":42,**{c:11 for c in ALL}})
title(ws, "Ather Energy Ltd — Income Statement (INR millions)", "B2")
put(ws,"R2",'=CONCATENATE("Scenario: ",'+f"{ref(SC,'F5')})", kind="calc", bold=True, color=NAVY)
year_header(ws,4)

def stmt_row(ws,row,label,hist,fcst,nf=MONEY,bold=False,topbot=False,border_b=False):
    put(ws,f"B{row}",label,kind="label",bold=bold)
    for i,cc in enumerate(HCOLS):
        if hist[i] is None: continue
        put(ws,f"{cc}{row}",hist[i],kind="input",nf=nf,bold=bold)
    for i,cc in enumerate(FCOLS):
        f=fcst[i]
        if f is None: continue
        k="link" if "!" in f else "calc"
        put(ws,f"{cc}{row}",f,kind=k,nf=nf,bold=bold)
    if topbot:
        for cc in ["B"]+ALL: ws[f"{cc}{row}"].border=BORDER_TOPBOT
    if border_b:
        for cc in ["B"]+ALL: ws[f"{cc}{row}"].border=BORDER_BOTTOM

stmt_row(ws,5,"Revenue from operations",[798,4085,17836,17538,22550],
         [f"={ref(RB,cc+'14')}" for cc in FCOLS],bold=True)
stmt_row(ws,6,"Total cost of goods sold",[-917,-3916,-18200,-16318,-18768],
         [f"=-{cc}5*(1-{ref(ASm,cc+'25')})" for cc in FCOLS])
stmt_row(ws,7,"Gross Profit",[None]*5,[None]*8,bold=True,topbot=True)
for cc in ALL: put(ws,f"{cc}7",f"={cc}5+{cc}6",kind="calc",nf=MONEY,bold=True)
put(ws,"B8","Operating expenses",kind="label",italic=True)
stmt_row(ws,9,"   Employee benefits expense",[-657,-1148,-2600,-3692,-4124],
         [f"=-{cc}5*{ref(ASm,cc+'26')}" for cc in FCOLS])
stmt_row(ws,10,"   Other operating expenses",[-1015,-1634,-4078,-4375,-5467],
         [f"=-{cc}5*{ref(ASm,cc+'27')}" for cc in FCOLS])
put(ws,"B11","Total operating expenses",kind="label",bold=True)
for cc in ALL: put(ws,f"{cc}11",f"=SUM({cc}9:{cc}10)",kind="calc",nf=MONEY,bold=True)
put(ws,"B12","EBITDA",kind="label",bold=True)
for cc in ALL: put(ws,f"{cc}12",f"={cc}7+{cc}11",kind="calc",nf=MONEY,bold=True)
for cc in ["B"]+ALL: ws[f"{cc}12"].border=BORDER_TOPBOT
stmt_row(ws,13,"Depreciation & amortisation",[-351,-484,-1128,-1467,-1710],
         [f"={ref(WS,cc+'33')}" for cc in FCOLS])
put(ws,"B14","EBIT",kind="label",bold=True)
for cc in ALL: put(ws,f"{cc}14",f"={cc}12+{cc}13",kind="calc",nf=MONEY,bold=True)
stmt_row(ws,15,"Other income",[85,54,225,353,502],["0"]*8)
stmt_row(ws,16,"Exceptional items",[0,0,0,-1746,0],["0"]*8)
stmt_row(ws,17,"Finance costs",[-276,-398,-700,-890,-1106],
         [f"=-{ref(WS,cc+'53')}" for cc in FCOLS])
stmt_row(ws,18,"Interest income",[0,0,0,0,0],
         [f"={ref(BSs,col_shift(cc,-1)+'43')}*{ref(ASm,cc+'44')}" for cc in FCOLS])
put(ws,"B19","Profit / (Loss) before tax",kind="label",bold=True)
for cc in ALL: put(ws,f"{cc}19",f"=SUM({cc}14:{cc}18)",kind="calc",nf=MONEY,bold=True)
for cc in ["B"]+ALL: ws[f"{cc}19"].border=BORDER_TOPBOT
stmt_row(ws,20,"Tax expense",[0,0,0,0,0],[f"=-{ref(WS,cc+'62')}" for cc in FCOLS])
stmt_row(ws,21,"Adjustments / share of associates",[0,0,0,0,0],["0"]*8)
put(ws,"B22","Profit / (Loss) after tax (PAT)",kind="label",bold=True)
for cc in ALL: put(ws,f"{cc}22",f"={cc}19+{cc}20+{cc}21",kind="calc",nf=MONEY,bold=True)
for cc in ["B"]+ALL: ws[f"{cc}22"].border=BORDER_TOPBOT
print("PL done")

# =====================================================================================
#  BALANCE SHEET
# =====================================================================================
ws = wb.create_sheet(BSs)
colwidths(ws, {"A":2.5,"B":42,**{c:11 for c in ALL}})
title(ws,"Ather Energy Ltd — Balance Sheet (INR millions)","B2")
year_header(ws,4)
section(ws,6,"EQUITY & LIABILITIES")
put(ws,"B7","Shareholders' funds",kind="label",italic=True)
# share capital
stmt_row(ws,8,"   Equity share capital & instruments",[4,7,6,8,291],
         ["=G8+81"]+[f"={col_shift(cc,-1)}8" for cc in FCOLS[1:]])
# reserves
stmt_row(ws,9,"   Other equity (reserves & surplus)",[3758,2242,6131,5451,4639],
         [f"={col_shift(cc,-1)}9+{ref(PLs,cc+'22')}+{ref(ASm,cc+'45')}-({cc}8-{col_shift(cc,-1)}8)+(-{ref(PLs,cc+'9')})*{ref(ASm,cc+'29')}" for cc in FCOLS])
put(ws,"B10","Total equity",kind="label",bold=True)
for cc in ALL: put(ws,f"{cc}10",f"={cc}8+{cc}9",kind="calc",nf=MONEY,bold=True)
section(ws,12,"Non-current liabilities") if False else None
put(ws,"B12","Non-current liabilities",kind="label",italic=True)
stmt_row(ws,13,"   Long-term borrowings",[1108,1088,1200,309,1169],
         [f"={ref(WS,cc+'39')}*($G$13/($G$13+$G$21))" for cc in FCOLS])
stmt_row(ws,14,"   Lease liabilities (non-current)",[375,597,1600,1419,1431],
         [f"={ref(WS,cc+'47')}*($G$14/($G$14+$G$22))" for cc in FCOLS])
stmt_row(ws,15,"   Long-term provisions",[124,290,480,702,792],
         [f"=$G$15/{ref(PLs,'$G$5')}*{ref(PLs,cc+'5')}" for cc in FCOLS])
stmt_row(ws,16,"   Other non-current liabilities",[6,19,200,482,952],
         [f"=$G$16/{ref(PLs,'$G$5')}*{ref(PLs,cc+'5')}" for cc in FCOLS])
put(ws,"B17","Total non-current liabilities",kind="label",bold=True)
for cc in ALL: put(ws,f"{cc}17",f"=SUM({cc}13:{cc}16)",kind="calc",nf=MONEY,bold=True)
put(ws,"B19","Current liabilities",kind="label",italic=True)
stmt_row(ws,20,"   Trade payables",[676,1216,2400,4027,5609],
         [f"={ref(WCS,cc+'15')}" for cc in FCOLS])
stmt_row(ws,21,"   Short-term borrowings",[609,1897,3652,2840,3330],
         [f"={ref(WS,cc+'39')}*($G$21/($G$13+$G$21))" for cc in FCOLS])
stmt_row(ws,22,"   Lease liabilities (current)",[97,68,264,209,263],
         [f"={ref(WS,cc+'47')}*($G$22/($G$14+$G$22))" for cc in FCOLS])
stmt_row(ws,23,"   Short-term provisions",[44,105,400,807,1189],
         [f"={ref(ASm,cc+'55')}*{ref(PLs,cc+'5')}" for cc in FCOLS])
stmt_row(ws,24,"   Other current liabilities",[612,657,1400,2881,1341],
         [f"={ref(ASm,cc+'54')}*{ref(PLs,cc+'5')}" for cc in FCOLS])
put(ws,"B25","Total current liabilities",kind="label",bold=True)
for cc in ALL: put(ws,f"{cc}25",f"=SUM({cc}20:{cc}24)",kind="calc",nf=MONEY,bold=True)
put(ws,"B27","TOTAL EQUITY & LIABILITIES",kind="label",bold=True)
for cc in ALL: put(ws,f"{cc}27",f"={cc}10+{cc}17+{cc}25",kind="calc",nf=MONEY,bold=True)
for cc in ["B"]+ALL: ws[f"{cc}27"].border=BORDER_TOPBOT

section(ws,29,"ASSETS")
put(ws,"B30","Non-current assets",kind="label",italic=True)
stmt_row(ws,31,"   Property, plant & equipment",[800,931,1400,1871,2674],
         [f"={ref(WS,cc+'9')}" for cc in FCOLS])
stmt_row(ws,32,"   Capital work-in-progress",[1,4,2,0,57],
         [f"={ref(WS,cc+'27')}" for cc in FCOLS])
stmt_row(ws,33,"   Right-of-use assets",[444,646,1700,1489,2443],
         [f"={ref(WS,cc+'16')}" for cc in FCOLS])
stmt_row(ws,34,"   Intangible assets (incl. under dev.)",[2294,2692,2300,1935,2206],
         [f"={ref(WS,cc+'23')}" for cc in FCOLS])
put(ws,"B35","Total fixed assets",kind="label",bold=True)
for cc in ALL: put(ws,f"{cc}35",f"=SUM({cc}31:{cc}34)",kind="calc",nf=MONEY,bold=True)
stmt_row(ws,36,"   Other non-current assets",[1289,987,1200,1546,2058],
         [f"=$G$36/{ref(PLs,'$G$5')}*{ref(PLs,cc+'5')}" for cc in FCOLS])
put(ws,"B37","Total non-current assets",kind="label",bold=True)
for cc in ALL: put(ws,f"{cc}37",f"={cc}35+{cc}36",kind="calc",nf=MONEY,bold=True)
put(ws,"B39","Current assets",kind="label",italic=True)
stmt_row(ws,40,"   Trade receivables",[0,10,12,16,118],
         [f"={ref(WCS,cc+'9')}" for cc in FCOLS])
stmt_row(ws,41,"   Inventories",[567,607,2574,1167,2446],
         [f"={ref(WCS,cc+'12')}" for cc in FCOLS])
stmt_row(ws,42,"   Other current assets",[541,1049,4819,3711,4890],
         [f"={ref(ASm,cc+'53')}*{ref(PLs,cc+'5')}" for cc in FCOLS])
stmt_row(ws,43,"   Cash, bank & current investments",[1477,1260,3726,7400,4114],
         [f"={ref(CSs,cc+'30')}" for cc in FCOLS])
put(ws,"B44","Total current assets",kind="label",bold=True)
for cc in ALL: put(ws,f"{cc}44",f"=SUM({cc}40:{cc}43)",kind="calc",nf=MONEY,bold=True)
put(ws,"B46","TOTAL ASSETS",kind="label",bold=True)
for cc in ALL: put(ws,f"{cc}46",f"={cc}37+{cc}44",kind="calc",nf=MONEY,bold=True)
for cc in ["B"]+ALL: ws[f"{cc}46"].border=BORDER_TOPBOT
put(ws,"B48","Balance check (Assets − Equity & Liab.)",kind="label",italic=True,color=GREY)
for cc in ALL: put(ws,f"{cc}48",f"={cc}46-{cc}27",kind="calc",nf=MONEY,italic=True)
print("BS done")

# =====================================================================================
#  CASH FLOW STATEMENT  (fully articulated -> balance sheet balances without a plug)
# =====================================================================================
ws = wb.create_sheet(CSs)
colwidths(ws, {"A":2.5,"B":48,**{c:11 for c in ALL}})
title(ws,"Ather Energy Ltd — Cash Flow Statement (INR millions)","B2")
year_header(ws,4)
put(ws,"B5","A. Operating activities",kind="label",italic=True)
stmt_row(ws,6,"Profit / (Loss) after tax",[-2333,-3441,-8645,-10597,-8123],
         [f"={ref(PLs,cc+'22')}" for cc in FCOLS],bold=True)
stmt_row(ws,7,"Add: Depreciation & amortisation",[351,484,1128,1467,1710],
         [f"=-{ref(PLs,cc+'13')}" for cc in FCOLS])
stmt_row(ws,8,"Add: Finance costs",[276,398,700,890,1106],
         [f"=-{ref(PLs,cc+'17')}" for cc in FCOLS])
stmt_row(ws,9,"Less: Interest income",[0,0,0,0,0],
         [f"=-{ref(PLs,cc+'18')}" for cc in FCOLS])
stmt_row(ws,10,"Add: Non-cash items (ESOP/share-based pay)",[227,359,520,500,640],
         [f"=(-{ref(PLs,cc+'9')})*{ref(ASm,cc+'29')}" for cc in FCOLS])
stmt_row(ws,11,"Change in net working capital",[138,-260,-1500,1407,-1278],
         [f"=-{ref(WCS,cc+'21')}" for cc in FCOLS])
stmt_row(ws,12,"Change in other non-current items",[0,0,0,0,0],
         [f"=({ref(BSs,cc+'15')}+{ref(BSs,cc+'16')}-{ref(BSs,col_shift(cc,-1)+'15')}-{ref(BSs,col_shift(cc,-1)+'16')})-({ref(BSs,cc+'36')}-{ref(BSs,col_shift(cc,-1)+'36')})" for cc in FCOLS])
put(ws,"B13","Cash flow from operating activities",kind="label",bold=True)
for cc in ALL: put(ws,f"{cc}13",f"=SUM({cc}6:{cc}12)",kind="calc",nf=MONEY,bold=True)
for cc in ["B"]+ALL: ws[f"{cc}13"].border=BORDER_TOPBOT
put(ws,"B14","B. Investing activities",kind="label",italic=True)
stmt_row(ws,15,"Interest income received",[33,18,40,196,190],
         [f"={ref(PLs,cc+'18')}" for cc in FCOLS])
stmt_row(ws,16,"Capex — PPE & CWIP",[-700,-560,-900,-900,-2600],
         [f"=-{ref(WS,cc+'7')}" for cc in FCOLS])
stmt_row(ws,17,"Capex — intangibles / development",[-170,-152,-300,-259,-790],
         [f"=-{ref(WS,cc+'21')}" for cc in FCOLS])
stmt_row(ws,18,"Other investing (deposits/investments, net)",[-590,638,-1900,-1318,-582],["0"]*8)
put(ws,"B19","Cash flow from investing activities",kind="label",bold=True)
for cc in ALL: put(ws,f"{cc}19",f"=SUM({cc}15:{cc}18)",kind="calc",nf=MONEY,bold=True)
for cc in ["B"]+ALL: ws[f"{cc}19"].border=BORDER_TOPBOT
put(ws,"B20","C. Financing activities",kind="label",italic=True)
stmt_row(ws,21,"Proceeds / (repayment) of borrowings, net",[-98,1268,452,-541,3247],
         [f"={ref(WS,cc+'38')}" for cc in FCOLS])
stmt_row(ws,22,"Lease principal repayment",[-100,-152,-180,-168,-211],
         [f"={ref(WS,cc+'46')}" for cc in FCOLS])
stmt_row(ws,23,"Finance costs paid",[-276,-341,-700,-770,-973],
         [f"={ref(PLs,cc+'17')}" for cc in FCOLS])
stmt_row(ws,24,"Equity issuance (IPO / fund raises)",[3439,1500,9000,9011,866],
         [f"={ref(ASm,cc+'45')}" for cc in FCOLS])
stmt_row(ws,25,"Dividends paid",[0,0,0,0,0],
         [f"=-{ref(ASm,cc+'46')}*MAX({ref(PLs,cc+'22')},0)" for cc in FCOLS])
put(ws,"B26","Cash flow from financing activities",kind="label",bold=True)
for cc in ALL: put(ws,f"{cc}26",f"=SUM({cc}21:{cc}25)",kind="calc",nf=MONEY,bold=True)
for cc in ["B"]+ALL: ws[f"{cc}26"].border=BORDER_TOPBOT
put(ws,"B28","Net change in cash & investments",kind="label",bold=True)
for cc in HCOLS: put(ws,f"{cc}28",f"={cc}30-{cc}29",kind="calc",nf=MONEY,bold=True)
for cc in FCOLS: put(ws,f"{cc}28",f"={cc}13+{cc}19+{cc}26",kind="calc",nf=MONEY,bold=True)
put(ws,"B29","Opening cash & investments",kind="label")
hist_open=[1343,1477,1260,3726,7400]
for i,cc in enumerate(HCOLS): put(ws,f"{cc}29",hist_open[i],kind="input",nf=MONEY)
for cc in FCOLS: put(ws,f"{cc}29",f"={col_shift(cc,-1)}30",kind="link",nf=MONEY)
put(ws,"B30","Closing cash & investments",kind="label",bold=True)
hist_close=[1477,1260,3726,7400,4114]
for i,cc in enumerate(HCOLS): put(ws,f"{cc}30",hist_close[i],kind="input",nf=MONEY,bold=True)
for cc in FCOLS: put(ws,f"{cc}30",f"={cc}28+{cc}29",kind="calc",nf=MONEY,bold=True)
for cc in ["B"]+ALL: ws[f"{cc}30"].border=BORDER_TOPBOT
print("CS done")


# =====================================================================================
#  WORKING CAPITAL SCHEDULE
# =====================================================================================
ws = wb.create_sheet(WCS)
colwidths(ws, {"A":2.5,"B":40,**{c:11 for c in ALL}})
title(ws,"Ather Energy Ltd — Working Capital Schedule (INR millions)","B2")
year_header(ws,4)
put(ws,"B6","Revenue",kind="label")
for cc in ALL: put(ws,f"{cc}6",f"={ref(PLs,cc+'5')}",kind="link",nf=MONEY)
put(ws,"B7","Cost of goods sold",kind="label")
for cc in ALL: put(ws,f"{cc}7",f"=-{ref(PLs,cc+'6')}",kind="link",nf=MONEY)
# Trade receivables
put(ws,"B9","Trade receivables",kind="label")
for cc in HCOLS: put(ws,f"{cc}9",f"={ref(BSs,cc+'40')}",kind="link",nf=MONEY)
for cc in FCOLS: put(ws,f"{cc}9",f"=(2*({cc}6*{ref(ASm,cc+'50')}/365))-{col_shift(cc,-1)}9",kind="calc",nf=MONEY)
put(ws,"B10","   Receivable days",kind="label",italic=True)
for cc in ALL[1:]: put(ws,f"{cc}10",f"=AVERAGE({col_shift(cc,-1)}9:{cc}9)/{cc}6*365",kind="calc",nf=DAYS,italic=True)
# Inventory
put(ws,"B12","Inventories",kind="label")
for cc in HCOLS: put(ws,f"{cc}12",f"={ref(BSs,cc+'41')}",kind="link",nf=MONEY)
for cc in FCOLS: put(ws,f"{cc}12",f"=(2*({cc}7*{ref(ASm,cc+'51')}/365))-{col_shift(cc,-1)}12",kind="calc",nf=MONEY)
put(ws,"B13","   Inventory days",kind="label",italic=True)
for cc in ALL[1:]: put(ws,f"{cc}13",f"=AVERAGE({col_shift(cc,-1)}12:{cc}12)/{cc}7*365",kind="calc",nf=DAYS,italic=True)
# Payables
put(ws,"B15","Trade payables",kind="label")
for cc in HCOLS: put(ws,f"{cc}15",f"={ref(BSs,cc+'20')}",kind="link",nf=MONEY)
for cc in FCOLS: put(ws,f"{cc}15",f"=(2*({cc}7*{ref(ASm,cc+'52')}/365))-{col_shift(cc,-1)}15",kind="calc",nf=MONEY)
put(ws,"B16","   Payable days",kind="label",italic=True)
for cc in ALL[1:]: put(ws,f"{cc}16",f"=AVERAGE({col_shift(cc,-1)}15:{cc}15)/{cc}7*365",kind="calc",nf=DAYS,italic=True)
# NWC
put(ws,"B18","Operating current assets (AR+Inv+Other CA)",kind="label")
for cc in ALL: put(ws,f"{cc}18",f"={ref(BSs,cc+'40')}+{ref(BSs,cc+'41')}+{ref(BSs,cc+'42')}",kind="link",nf=MONEY)
put(ws,"B19","Operating current liabilities (Pay+Prov+Other CL)",kind="label")
for cc in ALL: put(ws,f"{cc}19",f"={ref(BSs,cc+'20')}+{ref(BSs,cc+'23')}+{ref(BSs,cc+'24')}",kind="link",nf=MONEY)
put(ws,"B20","Net working capital",kind="label",bold=True)
for cc in ALL: put(ws,f"{cc}20",f"={cc}18-{cc}19",kind="calc",nf=MONEY,bold=True)
put(ws,"B21","Change in net working capital",kind="label",bold=True)
for cc in ALL[1:]: put(ws,f"{cc}21",f"={cc}20-{col_shift(cc,-1)}20",kind="calc",nf=MONEY,bold=True)
print("WCS done")

# =====================================================================================
#  WORKING SCHEDULES  (PPE, ROU, Intangibles, CWIP, D&A, Debt, Lease, Finance, Tax)
# =====================================================================================
ws = wb.create_sheet(WS)
colwidths(ws, {"A":2.5,"B":40,**{c:11 for c in ALL},"Q":14,"R":10})
title(ws,"Ather Energy Ltd — Supporting Schedules (INR millions)","B2")

def ws_section(row,text): section(ws,row,text)
def ws_yr(row): year_header(ws,row)
def anchorG(row,val): put(ws,f"G{row}",val,kind="input",nf=MONEY,bold=True)

# ---- PPE ----
ws_section(4,"Property, Plant & Equipment Schedule (net block roll-forward)")
ws_yr(5)
put(ws,"B6","Opening net block",kind="label")
for cc in FCOLS: put(ws,f"{cc}6",f"={col_shift(cc,-1)}9",kind="link",nf=MONEY)
put(ws,"B7","Capex (PPE & CWIP)",kind="label")
for cc in FCOLS: put(ws,f"{cc}7",f"={ref(PLs,cc+'5')}*{ref(ASm,cc+'32')}",kind="link",nf=MONEY)
put(ws,"B8","Depreciation",kind="label")
for cc in FCOLS: put(ws,f"{cc}8",f"=-({cc}6+0.5*{cc}7)/{ref(ASm,cc+'33')}",kind="calc",nf=MONEY)
put(ws,"B9","Closing net block",kind="label",bold=True)
anchorG(9,2674)
for cc in FCOLS: put(ws,f"{cc}9",f"={cc}6+{cc}7+{cc}8",kind="calc",nf=MONEY,bold=True)

# ---- ROU ----
ws_section(11,"Right-of-Use (Lease) Assets Schedule")
ws_yr(12)
put(ws,"B13","Opening ROU assets",kind="label")
for cc in FCOLS: put(ws,f"{cc}13",f"={col_shift(cc,-1)}16",kind="link",nf=MONEY)
put(ws,"B14","Additions (new leases)",kind="label")
for cc in FCOLS: put(ws,f"{cc}14",f"={ref(PLs,cc+'5')}*{ref(ASm,cc+'36')}",kind="link",nf=MONEY)
put(ws,"B15","Amortisation",kind="label")
for cc in FCOLS: put(ws,f"{cc}15",f"=-({cc}13+0.5*{cc}14)*{ref(ASm,cc+'37')}",kind="calc",nf=MONEY)
put(ws,"B16","Closing ROU assets",kind="label",bold=True)
anchorG(16,2443)
for cc in FCOLS: put(ws,f"{cc}16",f"=SUM({cc}13:{cc}15)",kind="calc",nf=MONEY,bold=True)

# ---- Intangibles ----
ws_section(18,"Intangible Assets Schedule (incl. assets under development)")
ws_yr(19)
put(ws,"B20","Opening intangibles",kind="label")
for cc in FCOLS: put(ws,f"{cc}20",f"={col_shift(cc,-1)}23",kind="link",nf=MONEY)
put(ws,"B21","Additions (capitalised development)",kind="label")
for cc in FCOLS: put(ws,f"{cc}21",f"={ref(PLs,cc+'5')}*{ref(ASm,cc+'34')}",kind="link",nf=MONEY)
put(ws,"B22","Amortisation",kind="label")
for cc in FCOLS: put(ws,f"{cc}22",f"=-{cc}20*{ref(ASm,cc+'35')}",kind="calc",nf=MONEY)
put(ws,"B23","Closing intangibles",kind="label",bold=True)
anchorG(23,2206)
for cc in FCOLS: put(ws,f"{cc}23",f"=SUM({cc}20:{cc}22)",kind="calc",nf=MONEY,bold=True)

# ---- CWIP ----
ws_section(25,"Capital Work-in-Progress")
ws_yr(26)
put(ws,"B27","Capital work-in-progress (held)",kind="label",bold=True)
anchorG(27,57)
for cc in FCOLS: put(ws,f"{cc}27",f"={col_shift(cc,-1)}27",kind="link",nf=MONEY,bold=True)

# ---- D&A summary ----
ws_section(29,"Total Depreciation & Amortisation")
put(ws,"B30","Depreciation — PPE",kind="label")
for cc in FCOLS: put(ws,f"{cc}30",f"={cc}8",kind="calc",nf=MONEY)
put(ws,"B31","Amortisation — ROU",kind="label")
for cc in FCOLS: put(ws,f"{cc}31",f"={cc}15",kind="calc",nf=MONEY)
put(ws,"B32","Amortisation — intangibles",kind="label")
for cc in FCOLS: put(ws,f"{cc}32",f"={cc}22",kind="calc",nf=MONEY)
put(ws,"B33","Total D&A (to P&L)",kind="label",bold=True)
for cc in FCOLS: put(ws,f"{cc}33",f"=SUM({cc}30:{cc}32)",kind="calc",nf=MONEY,bold=True)

# ---- Borrowings ----
ws_section(35,"Borrowings (Debt) Schedule")
ws_yr(36)
put(ws,"B37","Opening borrowings",kind="label")
for cc in FCOLS: put(ws,f"{cc}37",f"={col_shift(cc,-1)}39",kind="link",nf=MONEY)
put(ws,"B38","Net drawdown / (repayment)",kind="label")
for cc in FCOLS: put(ws,f"{cc}38",f"={ref(ASm,cc+'41')}",kind="link",nf=MONEY)
put(ws,"B39","Closing borrowings",kind="label",bold=True)
anchorG(39,4499)
for cc in FCOLS: put(ws,f"{cc}39",f"={cc}37+{cc}38",kind="calc",nf=MONEY,bold=True)
put(ws,"B40","Interest on borrowings",kind="label")
for cc in FCOLS: put(ws,f"{cc}40",f"={cc}37*{ref(ASm,cc+'40')}",kind="calc",nf=MONEY)

# ---- Lease liabilities ----
ws_section(42,"Lease Liability Schedule")
ws_yr(43)
put(ws,"B44","Opening lease liability",kind="label")
for cc in FCOLS: put(ws,f"{cc}44",f"={col_shift(cc,-1)}47",kind="link",nf=MONEY)
put(ws,"B45","Additions (new leases)",kind="label")
for cc in FCOLS: put(ws,f"{cc}45",f"={cc}14",kind="calc",nf=MONEY)
put(ws,"B46","Principal repayment",kind="label")
put(ws,"G46",-211,kind="input",nf=MONEY)
for cc in FCOLS: put(ws,f"{cc}46",f"={col_shift(cc,-1)}46*(1+{ref(ASm,cc+'43')})",kind="calc",nf=MONEY)
put(ws,"B47","Closing lease liability",kind="label",bold=True)
anchorG(47,1694)
for cc in FCOLS: put(ws,f"{cc}47",f"={cc}44+{cc}45+{cc}46",kind="calc",nf=MONEY,bold=True)
put(ws,"B48","Interest on lease liabilities",kind="label")
for cc in FCOLS: put(ws,f"{cc}48",f"={cc}44*{ref(ASm,cc+'42')}",kind="calc",nf=MONEY)

# ---- Finance cost ----
ws_section(50,"Finance Cost Schedule")
put(ws,"B51","Interest on borrowings",kind="label")
for cc in FCOLS: put(ws,f"{cc}51",f"={cc}40",kind="calc",nf=MONEY)
put(ws,"B52","Interest on leases",kind="label")
for cc in FCOLS: put(ws,f"{cc}52",f"={cc}48",kind="calc",nf=MONEY)
put(ws,"B53","Total finance cost (to P&L)",kind="label",bold=True)
for cc in FCOLS: put(ws,f"{cc}53",f"={cc}51+{cc}52",kind="calc",nf=MONEY,bold=True)

# ---- Tax & carry-forward losses ----
ws_section(55,"Tax & Carry-forward Losses Schedule")
ws_yr(56)
put(ws,"B57","Profit / (Loss) before tax",kind="label")
for cc in FCOLS: put(ws,f"{cc}57",f"={ref(PLs,cc+'19')}",kind="link",nf=MONEY)
put(ws,"B58","Opening carry-forward losses",kind="label")
put(ws,"G60b" if False else "G58",45000,kind="input",nf=MONEY)  # FY25 accumulated unabsorbed losses (approx)
for cc in FCOLS:
    if cc=="H": put(ws,f"{cc}58","=G58",kind="link",nf=MONEY)
    else: put(ws,f"{cc}58",f"={col_shift(cc,-1)}61",kind="link",nf=MONEY)
put(ws,"B59","Add: new losses / (Less: utilised)",kind="label")
for cc in FCOLS: put(ws,f"{cc}59",f"=IF({cc}57<0,-{cc}57,-MIN({cc}57,{cc}58))",kind="calc",nf=MONEY)
put(ws,"B60","Taxable profit after set-off",kind="label")
for cc in FCOLS: put(ws,f"{cc}60",f"=MAX({cc}57-{cc}58,0)",kind="calc",nf=MONEY)
put(ws,"B61","Closing carry-forward losses",kind="label",bold=True)
for cc in FCOLS: put(ws,f"{cc}61",f"={cc}58+{cc}59",kind="calc",nf=MONEY,bold=True)
put(ws,"B62","Current tax (to P&L)",kind="label",bold=True)
for cc in FCOLS: put(ws,f"{cc}62",f"={cc}60*{ref(ASm,cc+'47')}",kind="calc",nf=MONEY,bold=True)
print("WS done")


# =====================================================================================
#  RATIO ANALYSIS
# =====================================================================================
ws = wb.create_sheet(RA)
colwidths(ws, {"A":2.5,"B":40,**{c:11 for c in ALL}})
title(ws,"Ather Energy Ltd — Ratio Analysis","B2")
year_header(ws,4)
def rrow(row,label,formula_fn,nf=PCT,first_blank=True,bold=False,section_hdr=False):
    if section_hdr:
        section(ws,row,label); return
    put(ws,f"B{row}",label,kind="label",bold=bold)
    for i,cc in enumerate(ALL):
        if first_blank and i==0:
            put(ws,f"{cc}{row}","",kind="calc",nf=nf); continue
        prev=ALL[i-1] if i>0 else cc
        f=formula_fn(cc,prev)
        put(ws,f"{cc}{row}",f,kind=("link" if "!" in f else "calc"),nf=nf)

section(ws,5,"Profitability")
rrow(6,"Revenue growth %",lambda c,p:f"={ref(PLs,c+'5')}/{ref(PLs,p+'5')}-1")
rrow(7,"Gross margin %",lambda c,p:f"={ref(PLs,c+'7')}/{ref(PLs,c+'5')}",first_blank=False)
rrow(8,"EBITDA margin %",lambda c,p:f"={ref(PLs,c+'12')}/{ref(PLs,c+'5')}",first_blank=False)
rrow(9,"EBIT margin %",lambda c,p:f"={ref(PLs,c+'14')}/{ref(PLs,c+'5')}",first_blank=False)
rrow(10,"Net profit margin %",lambda c,p:f"={ref(PLs,c+'22')}/{ref(PLs,c+'5')}",first_blank=False)
rrow(11,"Return on equity (ROE) %",lambda c,p:f"={ref(PLs,c+'22')}/AVERAGE({ref(BSs,p+'10')},{ref(BSs,c+'10')})")
rrow(12,"Return on capital employed (ROCE) %",lambda c,p:f"={ref(PLs,c+'14')}/({ref(BSs,c+'46')}-{ref(BSs,c+'25')})",first_blank=False)
rrow(13,"Return on invested capital (ROIC) %",lambda c,p:f"={ref(PLs,c+'14')}/({ref(BSs,c+'10')}+{ref(BSs,c+'13')}+{ref(BSs,c+'21')}+{ref(BSs,c+'14')}+{ref(BSs,c+'22')})",first_blank=False)
section(ws,15,"Liquidity & Solvency")
rrow(16,"Current ratio",lambda c,p:f"={ref(BSs,c+'44')}/{ref(BSs,c+'25')}",nf=MULT,first_blank=False)
rrow(17,"Quick ratio",lambda c,p:f"=({ref(BSs,c+'44')}-{ref(BSs,c+'41')})/{ref(BSs,c+'25')}",nf=MULT,first_blank=False)
rrow(18,"Debt-to-equity (incl. leases)",lambda c,p:f"=({ref(BSs,c+'13')}+{ref(BSs,c+'21')}+{ref(BSs,c+'14')}+{ref(BSs,c+'22')})/{ref(BSs,c+'10')}",nf=MULT,first_blank=False)
rrow(19,"Net debt / EBITDA",lambda c,p:f"=({ref(BSs,c+'13')}+{ref(BSs,c+'21')}+{ref(BSs,c+'14')}+{ref(BSs,c+'22')}-{ref(BSs,c+'43')})/{ref(PLs,c+'12')}",nf=MULT,first_blank=False)
rrow(20,"Interest coverage (EBIT/Interest)",lambda c,p:f"={ref(PLs,c+'14')}/-{ref(PLs,c+'17')}",nf=MULT,first_blank=False)
section(ws,22,"Efficiency / Working Capital")
rrow(23,"Asset turnover (x)",lambda c,p:f"={ref(PLs,c+'5')}/AVERAGE({ref(BSs,p+'46')},{ref(BSs,c+'46')})",nf=MULT)
rrow(24,"Fixed-asset turnover (x)",lambda c,p:f"={ref(PLs,c+'5')}/AVERAGE({ref(BSs,p+'35')},{ref(BSs,c+'35')})",nf=MULT)
rrow(25,"Receivable days",lambda c,p:f"={ref(WCS,c+'10')}",nf=DAYS,first_blank=True)
rrow(26,"Inventory days",lambda c,p:f"={ref(WCS,c+'13')}",nf=DAYS,first_blank=True)
rrow(27,"Payable days",lambda c,p:f"={ref(WCS,c+'16')}",nf=DAYS,first_blank=True)
rrow(28,"Cash conversion cycle (days)",lambda c,p:f"={ref(WCS,c+'10')}+{ref(WCS,c+'13')}-{ref(WCS,c+'16')}",nf=DAYS,first_blank=True)
section(ws,30,"Cash Flow & Per-share")
rrow(31,"CFO / Revenue %",lambda c,p:f"={ref(CSs,c+'13')}/{ref(PLs,c+'5')}",first_blank=False)
rrow(32,"Free cash flow (INR m)",lambda c,p:f"={ref(CSs,c+'13')}+{ref(CSs,c+'16')}+{ref(CSs,c+'17')}",nf=MONEY,first_blank=False)
# shares outstanding & EPS
put(ws,"B34","Shares outstanding (m)",kind="label")
sh=[290.6,290.6,290.6,290.6,290.6]
for i,cc in enumerate(HCOLS): put(ws,f"{cc}34",sh[i],kind="input",nf=NUM2)
for cc in FCOLS: put(ws,f"{cc}34",372,kind="input",nf=NUM2)
put(ws,"B35","EPS (INR)",kind="label")
for cc in ALL: put(ws,f"{cc}35",f"={ref(PLs,cc+'22')}/{cc}34",kind="calc",nf=NUM2)
print("Ratio Analysis done")

# =====================================================================================
#  ERROR CHECKS
# =====================================================================================
ws = wb.create_sheet(EC)
colwidths(ws, {"A":2.5,"B":52,**{c:11 for c in ALL}})
title(ws,"Ather Energy Ltd — Integrity & Error Checks","B2")
put(ws,"B3","Every check should read PASS. Tolerance = 0.5 (INR m).",kind="label",italic=True,color=GREY)
year_header(ws,4)
def check(row,label,cond_fn,cols=ALL):
    put(ws,f"B{row}",label,kind="label")
    for cc in cols:
        f=cond_fn(cc)
        put(ws,f"{cc}{row}",f,kind=("link" if "!" in f else "calc"),align="center")
section(ws,5,"Balance sheet integrity")
check(6,"Balance sheet balances (Assets = Equity + Liab.)",
      lambda c:f'=IF(ABS({ref(BSs,c+"48")})<0.5,"PASS","FAIL")')
check(7,"Total assets reconcile (NCA + CA)",
      lambda c:f'=IF(ABS({ref(BSs,c+"46")}-{ref(BSs,c+"37")}-{ref(BSs,c+"44")})<0.5,"PASS","FAIL")')
section(ws,9,"Cash flow integrity")
check(10,"BS cash = Cash-flow closing cash",
      lambda c:f'=IF(ABS({ref(BSs,c+"43")}-{ref(CSs,c+"30")})<0.5,"PASS","FAIL")')
check(11,"Opening cash (t) = Closing cash (t-1)",
      lambda c:f'=IF(ABS({ref(CSs,c+"29")}-{ref(CSs,col_shift(c,-1)+"30")})<0.5,"PASS","FAIL")',cols=ALL[1:])
check(12,"Cash balance non-negative",
      lambda c:f'=IF({ref(BSs,c+"43")}>=-0.5,"PASS","FAIL")')
section(ws,14,"Roll-forward checks (forecast)")
check(15,"Borrowings roll-forward (close = open + net)",
      lambda c:f'=IF(ABS({ref(WS,c+"39")}-{ref(WS,c+"37")}-{ref(WS,c+"38")})<0.5,"PASS","FAIL")',cols=FCOLS)
check(16,"PPE roll-forward (close = open + capex + dep)",
      lambda c:f'=IF(ABS({ref(WS,c+"9")}-{ref(WS,c+"6")}-{ref(WS,c+"7")}-{ref(WS,c+"8")})<0.5,"PASS","FAIL")',cols=FCOLS)
check(17,"Lease roll-forward (close = open + add + repay)",
      lambda c:f'=IF(ABS({ref(WS,c+"47")}-{ref(WS,c+"44")}-{ref(WS,c+"45")}-{ref(WS,c+"46")})<0.5,"PASS","FAIL")',cols=FCOLS)
check(18,"Reserves roll-forward (PAT flows to equity)",
      lambda c:f'=IF(ABS({ref(BSs,c+"9")}-{ref(BSs,col_shift(c,-1)+"9")}-{ref(PLs,c+"22")}-{ref(ASm,c+"45")}+({ref(BSs,c+"8")}-{ref(BSs,col_shift(c,-1)+"8")})-(-{ref(PLs,c+"9")})*{ref(ASm,c+"29")})<0.5,"PASS","FAIL")',cols=FCOLS)
check(19,"Gross profit = Revenue + COGS",
      lambda c:f'=IF(ABS({ref(PLs,c+"7")}-{ref(PLs,c+"5")}-{ref(PLs,c+"6")})<0.5,"PASS","FAIL")')
section(ws,21,"Overall")
put(ws,"B22","ALL CHECKS PASS?",kind="label",bold=True)
# count fails across the check cells
rng_parts=[]
for r in [6,7,10,11,12,15,16,17,18,19]:
    rng_parts.append(f'COUNTIF({"C" if r in (6,7,10,12,19) else "D"}{r}:O{r},"FAIL")')
put(ws,"C22",f'=IF(SUM({",".join(rng_parts)})=0,"ALL CHECKS PASS","REVIEW REQUIRED")',kind="calc",bold=True,color=NAVY)
print("Error Checks done")


# =====================================================================================
#  DCF VALUATION  (3-stage: explicit FY2026-33, fade FY2034-40, terminal)
# =====================================================================================
ws = wb.create_sheet(DCFs)
DCOLS=["C","D","E","F","G","H","I","J"]      # Stage 1: FY2026..FY2033
S2=["C","D","E","F","G","H","I"]             # Stage 2 (fade): FY2034..FY2040
colwidths(ws, {"A":2.5,"B":42,**{c:11 for c in DCOLS},"K":2,"L":26,"M":11,"N":18})
title(ws,"Ather Energy Ltd — DCF Valuation (3-stage FCFF)","B2")
pairs=list(zip(DCOLS,FCOLS))

# WACC block (CAPM)
put(ws,"L2","WACC (CAPM)",kind="label",bold=True,size=14,color=NAVY)
def wrow(r,lbl,val,nf,src="",kind="input"):
    put(ws,f"L{r}",lbl,kind="label"); put(ws,f"M{r}",val,kind=kind,nf=nf)
    if src: put(ws,f"N{r}",src,kind="label",size=9,color=GREY)
wrow(4,"Risk-free rate (Rf)",0.068,PCT2,"10-yr G-Sec")
wrow(5,"Equity market return (Rm)",0.12,PCT2,"Nifty long-run")
wrow(6,"Equity risk premium",f"=M5-M4",PCT2,"Rm - Rf","calc")
wrow(7,"Beta (levered)",1.15,NUM2,"High-growth EV peer")
wrow(8,"Cost of equity (Ke)",f"=M4+M7*M6",PCT2,"CAPM","calc")
wrow(9,"WACC",f"=M8",PCT2,"Debt small -> ~Ke","calc")

# ---- Stage 1: explicit (linked to detailed model) ----
section(ws,3,"Stage 1 — Explicit forecast (FY2026E–FY2033E), linked to the detailed model")
put(ws,"B4","Particulars",kind="label",bold=True,fill=LGREY)
put(ws,"C4",2026,kind="label",bold=True,nf=YEAR_E,align="center",fill=LGREY)
for i,cc in enumerate(DCOLS[1:],1):
    put(ws,f"{cc}4",f"={DCOLS[i-1]}4+1",kind="label",bold=True,nf=YEAR_E,align="center",fill=LGREY)
put(ws,"B5","EBIT",kind="label")
for dc,sc in pairs: put(ws,f"{dc}5",f"={ref(PLs,sc+'14')}",kind="link",nf=MONEY)
put(ws,"B6","Less: tax on EBIT",kind="label")
for dc,sc in pairs: put(ws,f"{dc}6",f"=-{dc}5*{ref(ASm,sc+'47')}",kind="calc",nf=MONEY)
put(ws,"B7","NOPAT",kind="label",bold=True)
for dc,sc in pairs: put(ws,f"{dc}7",f"={dc}5+{dc}6",kind="calc",nf=MONEY,bold=True)
put(ws,"B8","Add: Depreciation & amortisation",kind="label")
for dc,sc in pairs: put(ws,f"{dc}8",f"=-{ref(PLs,sc+'13')}",kind="link",nf=MONEY)
put(ws,"B9","Less: Capex (PPE + intangibles)",kind="label")
for dc,sc in pairs: put(ws,f"{dc}9",f"=-({ref(WS,sc+'7')}+{ref(WS,sc+'21')})",kind="link",nf=MONEY)
put(ws,"B10","Less: Increase in net working capital",kind="label")
for dc,sc in pairs: put(ws,f"{dc}10",f"=-{ref(WCS,sc+'21')}",kind="link",nf=MONEY)
put(ws,"B11","FCFF",kind="label",bold=True)
for dc,sc in pairs: put(ws,f"{dc}11",f"=SUM({dc}7:{dc}10)",kind="calc",nf=MONEY,bold=True)
for cc in ["B"]+DCOLS: ws[f"{cc}11"].border=BORDER_TOPBOT
put(ws,"B12","Discount period",kind="label")
put(ws,"C12",1,kind="input",nf=NUM2)
for i,cc in enumerate(DCOLS[1:],1): put(ws,f"{cc}12",f"={DCOLS[i-1]}12+1",kind="calc",nf=NUM2)
put(ws,"B13","PV of FCFF",kind="label",bold=True)
for dc,sc in pairs: put(ws,f"{dc}13",f"={dc}11/(1+$M$9)^{dc}12",kind="calc",nf=MONEY,bold=True)

# ---- Stage 2: fade / convergence ----
section(ws,15,"Stage 2 — Fade / convergence (FY2034E–FY2040E): growth & margins converge to steady state")
put(ws,"B16","Particulars",kind="label",bold=True,fill=LGREY)
put(ws,"C16",2034,kind="label",bold=True,nf=YEAR_E,align="center",fill=LGREY)
for i,cc in enumerate(S2[1:],1):
    put(ws,f"{cc}16",f"={S2[i-1]}16+1",kind="label",bold=True,nf=YEAR_E,align="center",fill=LGREY)
fade_g   =[0.11,0.10,0.09,0.08,0.07,0.06,0.055]   # revenue growth fades
fade_m   =[0.04,0.06,0.08,0.10,0.12,0.135,0.15]   # EBIT margin ramps to mature ~15%
fade_tax =[0.05,0.08,0.12,0.16,0.20,0.2517,0.2517]# losses fade, then statutory
fade_rir =[0.55,0.50,0.45,0.40,0.36,0.33,0.30]    # reinvestment rate (g / ROIC ~18-20%)
put(ws,"B17","Revenue growth %",kind="label")
for i,cc in enumerate(S2): put(ws,f"{cc}17",fade_g[i],kind="input",nf=PCT)
put(ws,"B18","Revenue",kind="label")
put(ws,"C18",f"={ref(PLs,'O5')}*(1+C17)",kind="link",nf=MONEY)
for i,cc in enumerate(S2[1:],1): put(ws,f"{cc}18",f"={S2[i-1]}18*(1+{cc}17)",kind="calc",nf=MONEY)
put(ws,"B19","EBIT margin %",kind="label")
for i,cc in enumerate(S2): put(ws,f"{cc}19",fade_m[i],kind="input",nf=PCT)
put(ws,"B20","EBIT",kind="label")
for cc in S2: put(ws,f"{cc}20",f"={cc}18*{cc}19",kind="calc",nf=MONEY)
put(ws,"B21","Effective tax rate %",kind="label")
for i,cc in enumerate(S2): put(ws,f"{cc}21",fade_tax[i],kind="input",nf=PCT)
put(ws,"B22","NOPAT",kind="label",bold=True)
for cc in S2: put(ws,f"{cc}22",f"={cc}20*(1-{cc}21)",kind="calc",nf=MONEY,bold=True)
put(ws,"B23","Reinvestment rate % (capex+NWC, net of D&A)",kind="label")
for i,cc in enumerate(S2): put(ws,f"{cc}23",fade_rir[i],kind="input",nf=PCT)
put(ws,"B24","FCFF",kind="label",bold=True)
for cc in S2: put(ws,f"{cc}24",f"={cc}22*(1-{cc}23)",kind="calc",nf=MONEY,bold=True)
for cc in ["B"]+S2: ws[f"{cc}24"].border=BORDER_TOPBOT
put(ws,"B25","Discount period",kind="label")
put(ws,"C25","=J12+1",kind="calc",nf=NUM2)
for i,cc in enumerate(S2[1:],1): put(ws,f"{cc}25",f"={S2[i-1]}25+1",kind="calc",nf=NUM2)
put(ws,"B26","PV of FCFF",kind="label",bold=True)
for cc in S2: put(ws,f"{cc}26",f"={cc}24/(1+$M$9)^{cc}25",kind="calc",nf=MONEY,bold=True)

# ---- Valuation summary ----
section(ws,28,"Valuation summary")
put(ws,"B29","Terminal growth rate (g)",kind="label"); put(ws,"C29",0.055,kind="input",nf=PCT2)
put(ws,"B30","Sum of PV — Stage 1 (FY26-33)",kind="label"); put(ws,"C30","=SUM(C13:J13)",kind="calc",nf=MONEY)
put(ws,"B31","Sum of PV — Stage 2 (FY34-40)",kind="label"); put(ws,"C31","=SUM(C26:I26)",kind="calc",nf=MONEY)
put(ws,"B32","Terminal value (end FY2040)",kind="label"); put(ws,"C32",f"=I24*(1+C29)/($M$9-C29)",kind="calc",nf=MONEY)
put(ws,"B33","PV of terminal value",kind="label"); put(ws,"C33",f"=C32/(1+$M$9)^I25",kind="calc",nf=MONEY)
put(ws,"B34","Enterprise value (EV)",kind="label",bold=True); put(ws,"C34","=C30+C31+C33",kind="calc",nf=MONEY,bold=True)
put(ws,"B35","Less: net debt (FY2025)",kind="label")
put(ws,"C35",f"=-({ref(BSs,'G13')}+{ref(BSs,'G21')}+{ref(BSs,'G14')}+{ref(BSs,'G22')}-{ref(BSs,'G43')})",kind="link",nf=MONEY)
put(ws,"B36","Equity value",kind="label",bold=True); put(ws,"C36","=C34+C35",kind="calc",nf=MONEY,bold=True)
put(ws,"B37","Shares outstanding (m)",kind="label"); put(ws,"C37",372,kind="input",nf=NUM2)
put(ws,"B38","Value per share (INR)",kind="label",bold=True); put(ws,"C38","=C36/C37",kind="calc",nf=NUM2,bold=True)
for cc in ["B","C"]: ws[f"{cc}38"].border=BORDER_TOPBOT
put(ws,"B40","% of value in terminal value",kind="label",italic=True,color=GREY)
put(ws,"C40","=C33/C34",kind="calc",nf=PCT,italic=True)
put(ws,"B42","Note: 3-stage FCFF. Stage 1 is the detailed model. Stage 2 lets growth fade and EBIT margin ramp to a "
              "mature ~15% before the terminal value is struck on a steady-state business (taxed NOPAT, reinvestment for "
              "growth). This is more rigorous than a single-period terminal off the still-maturing FY2033 — note the "
              "terminal value now correctly reflects tax and reinvestment. Highly sensitive to WACC, g, terminal margin "
              "and the fade path; read with the Relative Valuation and the Scenarios switch.",
    kind="label",italic=True,color=GREY,wrap=True)
ws.row_dimensions[42].height=58
print("DCF done")


# =====================================================================================
#  COMPANY & MARKET INSIGHTS
# =====================================================================================
ws = wb.create_sheet(CM)
colwidths(ws, {"A":2.5,"B":120})
title(ws,"Ather Energy Ltd — Company & Market Insights","B2")
def para(row,text,bold=False,size=11,color=BLACK,italic=False):
    put(ws,f"B{row}",text,kind="label",bold=bold,size=size,color=color,italic=italic,wrap=True)
    ws.row_dimensions[row].height = max(15, 15*(1+len(text)//110))
r=4
items=[
 ("Company snapshot",True,NAVY),
 ("Ather Energy Ltd (Bengaluru; incorporated 2013) designs and manufactures premium smart electric scooters (450 series, Rizta) and runs the AtherStack software and Ather Grid charging ecosystem. It listed on NSE/BSE in May 2025. Ather prepares STANDALONE financial statements only — it has no subsidiaries, so no consolidated statements exist; standalone is therefore the consolidated-equivalent used here.",False,BLACK),
 ("Operating performance",True,NAVY),
 ("Revenue grew from INR 798m (FY21) to INR 22,550m (FY25); FY25 revenue +29% YoY. The company remains loss-making (FY25 PAT INR -8,123m) but gross margin inflected sharply to ~16.8% in FY25 (from ~7% FY24) on scale, the Rizta launch and falling battery costs. Vehicle sales are ~88% of revenue; software, accessories, charging and service make up the balance and are higher-margin, recurring streams.",False,BLACK),
 ("Volumes, ASP & market share",True,NAVY),
 ("FY25 deliveries ~1.65 lakh units; revenue per vehicle ~INR 1.28 lakh (down from ~1.43 lakh FY24 and ~1.56 lakh FY23 as the cheaper Rizta lifted mix). Ather's E2W market share rose from ~13-14% (FY25) to ~18.7% by Mar-2026; retail network expanded from 351 outlets (Mar-25) toward ~700 (FY26). Source: Vahan / ETAuto / Autocar India / Storyboard18.",False,BLACK),
 ("Industry & demand drivers",True,NAVY),
 ("India electric two-wheeler volumes (~1.1-1.2m in FY25) are forecast to grow ~26-28% CAGR to 2034 (IMARC/Renub/TechSci). E2W penetration of the ~19m-unit 2W market is ~6-7% today and is projected at ~30% by 2030 (Kearney) and up to ~40% by FY31 in Ather's optimistic case. PM E-DRIVE (successor to FAME-II), 5% GST on EVs, state EV policies and PLI for cell manufacturing support demand and localisation.",False,BLACK),
 ("Battery cost trend (key margin lever)",True,NAVY),
 ("Lithium-ion pack prices fell ~8% to ~USD 108/kWh in 2025 and are forecast at ~USD 105/kWh in 2026, declining toward ~USD 80-90/kWh by 2030 (BloombergNEF, Dec-2025), aided by cell overcapacity and an ~84% lithium price correction. Batteries are Ather's largest input cost, so this is a structural gross-margin tailwind.",False,BLACK),
 ("Competition",True,NAVY),
 ("Top E2W OEMs (CY2025 / early-2026 Vahan): TVS Motor ~24%, Bajaj Auto ~22%, Ather ~13-18%, with Ola Electric's share collapsing (24.8% Jan-25 to <6% Jan-26), plus Hero/Greaves. TVS and Bajaj are profitable legacy OEMs with large ICE dealer networks; Ola and Ather are pure-play EV challengers still investing for scale. Ather differentiates on software, design and a premium brand.",False,BLACK),
 ("Capital & capacity",True,NAVY),
 ("The IPO raised a fresh issue of ~INR 2,626 cr (~26,260m), used for the new Maharashtra plant (Factory-3, tripling capacity), R&D, debt repayment and marketing. Post-IPO shares outstanding ~372m. Large accumulated tax losses (>INR 40,000m) shield taxable income across the forecast horizon (statutory rate 25.17%).",False,BLACK),
 ("Sources",True,NAVY),
 ("BloombergNEF battery price survey (Dec-2025); SIAM / Vahan registration data; ETAuto, Economic Times, Moneycontrol, Autocar India, Storyboard18, Angel One, Financial Express; IMARC / Renub / TechSci / Kearney industry reports; Ather Energy Annual Reports FY2021-FY2025 (standalone financial statements). Content paraphrased / summarised for licensing compliance.",False,GREY),
]
for txt,b,col in items:
    para(r,txt,bold=b,size=(12 if b else 11),color=col); r+=1
print("Company & Market Insights done")

# =====================================================================================
#  MODELLING NOTES (documentation)
# =====================================================================================
ws = wb.create_sheet(MN)
colwidths(ws, {"A":2.5,"B":120})
title(ws,"Ather Energy Ltd — Modelling Notes, Methodology, Risks & Limitations","B2")
def mpara(row,text,bold=False,size=11,color=BLACK):
    put(ws,f"B{row}",text,kind="label",bold=bold,size=(12 if bold else 11),color=color,wrap=True)
    ws.row_dimensions[row].height = max(15, 15*(1+len(text)//110))
notes=[
 ("1. Purpose & scope",True),
 ("Fully-linked, driver-based operating model and DCF for Ather Energy Ltd. History FY2021-FY2025; forecast FY2026E-FY2033E (8 years). All figures INR millions unless stated. Built to mirror the supplied United Spirits template's structure, colour coding and philosophy. Colour key: BLUE = hardcoded input, GREEN = link/reference to another sheet, BLACK = in-sheet calculation.",False),
 ("2. Data basis (consolidated vs standalone)",True),
 ("Ather has no subsidiaries and files only STANDALONE Ind AS financial statements; no consolidated statements exist, so standalone is used as the consolidated-equivalent. FY2021/22/24/25 are taken directly from the annual reports in the repository. FY2023's annual reports in the repo are image-only (no text layer); FY2023 was therefore reconstructed from reliable public sources (Economic Times/Moneycontrol: revenue INR 17,836m, total expenses INR 26,706m, net loss INR 8,645m, D&A INR 1,128m) and from anchors verified in the FY2025 report (1-Apr-2023 reserves INR 6,131m; opening borrowings/lease; opening cash INR 826m; FY23 inventory INR 2,574m). FY2023 sub-line splits are estimates that reconcile exactly to the verified totals.",False),
 ("3. Revenue methodology",True),
 ("Driver-based: vehicle revenue = unit volumes x ASP. Volume growth is set on the Scenarios sheet (Base/Bull/Bear) reflecting E2W market growth plus Ather share gains, tapering over time. ASP reflects the recent Rizta-led mix dip then modest premiumisation. Accessories, software/connectivity and charging are forecast as a % of vehicle revenue (higher-margin recurring streams broken out prospectively); service grows with the installed fleet. Historically these streams were bundled into product/service revenue per Note 20, so a small step-up appears in FY2026.",False),
 ("4. Cost, margin & operating leverage",True),
 ("Gross margin (Scenarios) expands from ~18% (FY26) to ~28% (FY33) on falling battery costs, scale, localisation/PLI and mix. Employee and other operating expenses are forecast as a declining % of sales, capturing operating leverage as a largely fixed cost base is spread over a fast-growing top line. The model reaches EBITDA breakeven around FY2030 and approaches PAT breakeven by FY2032-FY2033 — a realistic path, neither over-optimistic nor over-pessimistic.",False),
 ("5. Schedules",True),
 ("PPE, ROU (leases) and intangibles use opening + additions - depreciation/amortisation roll-forwards (capex and additions driven off revenue). Depreciation uses a straight-line life with a half-year convention on additions. Borrowings and leases use opening + net movement roll-forwards with interest on opening balances feeding the finance-cost schedule. A tax schedule tracks carry-forward losses (>INR 40,000m) which keep cash tax at nil across the forecast.",False),
 ("6. Integration & balancing",True),
 ("The three statements are fully linked with no circular references (interest income/expense use opening balances). The model is FULLY ARTICULATED: every balance-sheet line is forecast from a driver and cash is the genuine output of the cash-flow statement, so the balance sheet balances by accounting identity (no balancing plug). The Error Checks sheet confirms balance-sheet balancing, cash-flow reconciliation, beginning/ending balances and all roll-forwards return PASS for every year.",False),
 ("7. Valuation",True),
 ("A 3-stage FCFF DCF: Stage 1 is the explicit detailed model (FY2026-FY2033); Stage 2 (FY2034-FY2040) lets revenue growth fade and the EBIT margin ramp to a mature ~15% with taxed NOPAT and reinvestment for growth; the terminal value is then struck on a genuine steady state. Discounted at a CAPM-based WACC (Rf 6.8%, ERP 5.2%, beta 1.15 -> ~12.8%) with ~5.5% terminal growth. This horizon (to FY2040) is appropriate because Ather only reaches EBITDA breakeven around FY2030 and PAT breakeven around FY2032-FY2033 — an 8-year cut-off would strike the terminal value on a still-immature, near-breakeven business. The DCF remains conservative versus the market price; read it alongside the Relative Valuation (EV/Sales) and the Base/Bull/Bear scenarios.",False),
 ("8. Key risks",True),
 ("Demand/competition: aggressive pricing by Ola, TVS, Bajaj could cap volumes/ASP and margins. Margin: slower battery-cost decline or input inflation. Execution: Factory-3 ramp and capex discipline. Funding: the model shows cash thinning by FY2033 — a further capital raise or faster profitability may be required. Policy: changes to PM E-DRIVE/GST incentives. Technology: rare-earth magnet supply and battery-chemistry shifts.",False),
 ("9. Limitations",True),
 ("FY2023 sub-line items are reconstructed (totals verified). Forecasts are scenario-based judgements, not guidance. Other non-current/other current items are modelled as ratios of sales. The DCF is sensitive to WACC and terminal growth (a sensitivity grid can be layered on). This model is for analytical/educational purposes and is not investment advice.",False),
 ("10. How to use",True),
 ("Change the single Scenarios!F4 switch (1=Base, 2=Bull, 3=Bear) to flex volume growth and gross margin across the whole model. All other assumptions live on the Assumptions sheet (each with a Source and a plain-language Reason). Inputs are blue; do not type over green links or black formulas.",False),
]
r=4
for txt,b in notes:
    mpara(r,txt,bold=b); r+=1
print("Modelling Notes done")

# =====================================================================================
#  COVER / INDEX
# =====================================================================================
ws = wb.create_sheet(COV)
colwidths(ws, {"A":2.5,"B":42,"C":60})
put(ws,"B2","ATHER ENERGY LIMITED",kind="label",bold=True,size=22,color=NAVY)
put(ws,"B3","Integrated Financial Model & DCF Valuation",kind="label",bold=True,size=14,color=GREY)
put(ws,"B4","History FY2021-FY2025  |  Forecast FY2026E-FY2033E  |  INR millions  |  Standalone (consolidated-equivalent)",kind="label",italic=True,color=GREY)
put(ws,"B6","Contents",kind="label",bold=True,size=12,color=NAVY)
idx=[("Scenarios","Base / Bull / Bear switch — volume growth & gross margin"),
     ("Assumptions","All macro, industry & company drivers with source + reason"),
     ("Revenue Build","Driver-based revenue (volume x ASP + accessories/software/charging/service)"),
     ("PL","Income statement (history + forecast)"),
     ("BS","Balance sheet (history + forecast)"),
     ("CS","Cash flow statement (fully articulated)"),
     ("Working Capital Schedule","Receivables / inventory / payables & change in NWC"),
     ("Working Schedules","PPE, ROU, intangibles, debt, lease, finance cost, tax/losses, D&A"),
     ("Ratio Analysis","Profitability, liquidity, leverage, efficiency, per-share"),
     ("DCF","FCFF DCF with CAPM WACC and terminal value"),
     ("Error Checks","Automated integrity checks (all PASS)"),
     ("Company & Market Insights","Industry, competition, battery costs, policy, sources"),
     ("Modelling Notes","Methodology, assumptions rationale, risks, limitations")]
r=7
for nm,desc in idx:
    put(ws,f"B{r}",nm,kind="label",bold=True,color=NAVY)
    put(ws,f"C{r}",desc,kind="label",color=GREY)
    r+=1
put(ws,f"B{r+1}","Colour key:",kind="label",bold=True)
put(ws,f"B{r+2}","Blue = hardcoded input",kind="label",color=BLUE)
put(ws,f"B{r+3}","Green = link / reference to another sheet",kind="label",color=GREEN)
put(ws,f"B{r+4}","Black = calculation",kind="label",color=BLACK)
put(ws,f"B{r+6}","Prepared as an equity-research-style model for analytical/educational use. Not investment advice.",kind="label",italic=True,color=GREY)
# headline outputs
put(ws,"E6","Headline outputs (Base case)",kind="label",bold=True,size=12,color=NAVY)
hl=[("FY2025 revenue (INR m)",f"={ref(PLs,'G5')}",MONEY),
    ("FY2033E revenue (INR m)",f"={ref(PLs,'O5')}",MONEY),
    ("FY2025 EBITDA margin",f"={ref(PLs,'G12')}/{ref(PLs,'G5')}",PCT),
    ("FY2033E EBITDA margin",f"={ref(PLs,'O12')}/{ref(PLs,'O5')}",PCT),
    ("FY2033E PAT (INR m)",f"={ref(PLs,'O22')}",MONEY),
    ("DCF value per share (INR)",f"={ref(DCFs,'C38')}",NUM2),
    ("Integrity checks",f"={ref(EC,'C22')}","General")]
rr=7
for lbl,f,nf in hl:
    put(ws,f"E{rr}",lbl,kind="label")
    put(ws,f"H{rr}",f,kind=("link" if "!" in f else "calc"),nf=(None if nf=="General" else nf),bold=True)
    rr+=1
colwidths(ws,{"E":30,"F":4,"G":4,"H":16})
# move cover to first
wb.move_sheet(COV, -(len(wb.sheetnames)-1))
print("Cover done")


# =====================================================================================
#  POLISH: freeze panes, tab colours, gridlines, zoom
# =====================================================================================
freeze_at = {SC:"C7", ASm:"C5", RB:"C5", PLs:"C5", BSs:"C5", CSs:"C5", WCS:"C5",
             WS:"C5", RA:"C5", EC:"C5", DCFs:"C5"}
for nm,cell in freeze_at.items():
    if nm in wb.sheetnames:
        wb[nm].freeze_panes = cell
tabcolors = {COV:"1F3864", SC:"2E75B6", ASm:"2E75B6", RB:"2E75B6", PLs:"548235", BSs:"548235",
             CSs:"548235", WCS:"BF8F00", WS:"BF8F00", RA:"7030A0", DCFs:"C00000",
             EC:"C00000", CM:"808080", MN:"808080"}
for nm,col in tabcolors.items():
    if nm in wb.sheetnames: wb[nm].sheet_properties.tabColor = col
for nm in wb.sheetnames:
    ws=wb[nm]
    ws.sheet_view.showGridLines=False
    ws.sheet_view.zoomScale=90
wb[COV].sheet_view.zoomScale=110
print("Polish done")


# =====================================================================================
#  RELATIVE VALUATION (Comparable Company Analysis)
# =====================================================================================
RV="Relative Valuation"
ws = wb.create_sheet(RV)
colwidths(ws, {"A":2.5,"B":24,"C":13,"D":12,"E":13,"F":12,"G":12,"H":11,"I":11,"J":12,"K":10,"L":11})
title(ws,"Ather Energy Ltd — Relative Valuation (Comparable Companies)","B2")
put(ws,"B3","Peer market data indicative, as of ~Jun-2026 (public sources: NSE/BSE, stockanalysis, dhan, companies-marketcap). Figures in INR crore unless stated. 1 cr = 10 INR m.",
    kind="label",italic=True,color=GREY,wrap=True)
ws.row_dimensions[3].height=28
hdr=["Company","Mkt cap (cr)","Net debt (cr)","EV (cr)","Revenue (cr)","EBITDA (cr)","PAT (cr)","EV/Sales","EV/EBITDA","P/E","EBITDA %"]
hcols=["B","C","D","E","F","G","H","I","J","K","L"]
for cc,h in zip(hcols,hdr):
    put(ws,f"{cc}5",h,kind="label",bold=True,fill=LGREY,align="center",wrap=True)
ws.row_dimensions[5].height=28
# peers: name, mktcap, netdebt, revenue, ebitda, pat   (INR crore, FY25/TTM indicative)
peers=[
 ("TVS Motor",      169592,  2000, 36000, 4320, 2200),
 ("Bajaj Auto",     265000,-15000, 49000, 9800, 7300),
 ("Hero MotoCorp",   90000, -8000, 40000, 5600, 4500),
 ("Eicher Motors",  208424,-16000, 18000, 4500, 4700),
 ("Ola Electric",    19000, -2000,  4514,-1500,-2000),
 ("Ather Energy (mkt)",37920,  150,  2255, -581, -812),
]
r=6
for nm,mc,nd,rev,eb,pat in peers:
    put(ws,f"B{r}",nm,kind="label",bold=(nm.startswith("Ather")))
    put(ws,f"C{r}",mc,kind="input",nf=NUM0)
    put(ws,f"D{r}",nd,kind="input",nf=NUM0)
    put(ws,f"E{r}",f"=C{r}+D{r}",kind="calc",nf=NUM0)
    put(ws,f"F{r}",rev,kind="input",nf=NUM0)
    put(ws,f"G{r}",eb,kind="input",nf=NUM0)
    put(ws,f"H{r}",pat,kind="input",nf=NUM0)
    put(ws,f"I{r}",f"=E{r}/F{r}",kind="calc",nf=MULT)
    put(ws,f"J{r}",f'=IF(G{r}>0,E{r}/G{r},"NM")',kind="calc",nf=MULT)
    put(ws,f"K{r}",f'=IF(H{r}>0,C{r}/H{r},"NM")',kind="calc",nf=MULT)
    put(ws,f"L{r}",f"=G{r}/F{r}",kind="calc",nf=PCT)
    r+=1
# medians
put(ws,"B13","Legacy 2W OEM median (TVS/Bajaj/Hero/Eicher)",kind="label",bold=True)
put(ws,"I13","=MEDIAN(I6:I9)",kind="calc",nf=MULT,bold=True)
put(ws,"J13","=MEDIAN(J6:J9)",kind="calc",nf=MULT,bold=True)
put(ws,"K13","=MEDIAN(K6:K9)",kind="calc",nf=MULT,bold=True)
put(ws,"B14","EV pure-play reference (Ola Electric)",kind="label",bold=True)
put(ws,"I14","=I10",kind="calc",nf=MULT,bold=True)

# implied valuation of Ather
section(ws,16,"Implied valuation of Ather Energy (EV/Sales — primary, as company is pre-profit)")
put(ws,"B17","Ather FY2027E revenue (INR cr)",kind="label")
put(ws,"C17",f"={ref(PLs,'I5')}/10",kind="link",nf=NUM0)
put(ws,"B18","Ather net debt — FY2025 (INR cr)",kind="label")
put(ws,"C18",f"=({ref(BSs,'G13')}+{ref(BSs,'G21')}+{ref(BSs,'G14')}+{ref(BSs,'G22')}-{ref(BSs,'G43')})/10",kind="link",nf=NUM0)
put(ws,"B19","Shares outstanding (m)",kind="label"); put(ws,"C19",383,kind="input",nf=NUM0)
put(ws,"B21","Scenario",kind="label",bold=True,fill=LGREY)
put(ws,"C21","EV/Sales (x)",kind="label",bold=True,fill=LGREY,align="center")
put(ws,"D21","Implied EV (cr)",kind="label",bold=True,fill=LGREY,align="center")
put(ws,"E21","Implied equity (cr)",kind="label",bold=True,fill=LGREY,align="center")
put(ws,"F21","Implied price (INR)",kind="label",bold=True,fill=LGREY,align="center")
mults=[("Conservative (3.0x)",3.0),("Base (4.5x)",4.5),("Bull (6.0x)",6.0)]
rr=22
for lbl,m in mults:
    put(ws,f"B{rr}",lbl,kind="label")
    put(ws,f"C{rr}",m,kind="input",nf=MULT)
    put(ws,f"D{rr}",f"=C{rr}*$C$17",kind="calc",nf=NUM0)
    put(ws,f"E{rr}",f"=D{rr}-$C$18",kind="calc",nf=NUM0)
    put(ws,f"F{rr}",f"=E{rr}*10/$C$19",kind="calc",nf=NUM0,bold=True)
    rr+=1

# football field / summary
section(ws,27,"Valuation summary (football field)")
put(ws,"B28","DCF — Base case (INR/share)",kind="label"); put(ws,"C28",f"={ref(DCFs,'C38')}",kind="link",nf=NUM0,bold=True)
put(ws,"B29","Relative — conservative to bull (INR/share)",kind="label")
put(ws,"C29","=F22",kind="calc",nf=NUM0,bold=True); put(ws,"D29",'="to "&TEXT(F24,"#,##0")',kind="calc")
put(ws,"B30","Current market price (INR/share)",kind="label"); put(ws,"C30",990,kind="input",nf=NUM0,bold=True)
put(ws,"B31","Implied market EV/Sales on FY25 revenue (x)",kind="label"); put(ws,"C31","=I11",kind="link",nf=MULT,bold=True)
put(ws,"B33","Read-through: Ather trades at a steep premium to legacy 2W OEMs and to Ola Electric on EV/Sales, so the "
              "market is already pricing in strong volume growth and a sharp margin ramp. On a conservative DCF and on "
              "peer-based EV/Sales, the stock screens expensive unless one underwrites the bull-case volumes and "
              "terminal margins. Use this alongside the DCF and the Scenarios switch. Indicative, not investment advice.",
    kind="label",italic=True,color=GREY,wrap=True)
ws.row_dimensions[33].height=58
# polish this sheet
ws.freeze_panes="C6"; ws.sheet_view.showGridLines=False; ws.sheet_view.zoomScale=90
ws.sheet_properties.tabColor="7030A0"
# position right after DCF
wb.move_sheet(RV, -(len(wb.sheetnames)-1-wb.sheetnames.index(DCFs)-1))
print("Relative Valuation done; order:", wb.sheetnames)
