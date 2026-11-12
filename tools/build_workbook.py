"""Builds model/Tech_Spend_Renewal_Analytics.xlsx with live Excel formulas.

Contracts, Spend, Inputs and the RFx bids are the only hard-coded inputs (blue font).
Every other figure is a formula, so changing a contract, a month of spend or a threshold
flows through utilization, benchmarking, renewals, savings and the dashboard.
Run from tools/:  ../.venv/bin/python build_workbook.py
"""
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter as col
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.chart import BarChart, Reference

from assumptions import (COMPANY, FISCAL_YEAR, AS_OF, MONTHS_CLOSED, MONTHS, INPUTS, CATEGORIES,
                         DEPARTMENTS, RFX_CONTRACT, RFX_CRITERIA, RFX_BIDS, contracts, spend_rows, acv)

OUT = Path(__file__).resolve().parents[1] / "model" / "Tech_Spend_Renewal_Analytics.xlsx"

NAVY = "1F3864"
HDR, HDR_FILL = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor=NAVY)
INPUT, BOLD = Font(color="0000FF"), Font(bold=True)
TITLE = Font(bold=True, size=14, color=NAVY)
K = '#,##0.0;(#,##0.0);"-"'      # $000s
USD = '#,##0;(#,##0);"-"'        # whole dollars / units
PCT = '0.0%;(0.0%);"-"'
DATE = "yyyy-mm-dd"
TOP = Border(top=Side(style="thin"))
RED = PatternFill("solid", fgColor="F8CBAD")
AMBER = PatternFill("solid", fgColor="FFE699")
GREEN = PatternFill("solid", fgColor="C6EFCE")

CT = contracts()
SP = spend_rows()
C1, CN = 4, 4 + len(CT) - 1          # Contracts data rows
S1, SN = 4, 4 + len(SP) - 1          # Spend data rows
SEAT = [(i, c) for i, c in enumerate(CT) if c["model"] == "Per seat"]
COMMIT = [(i, c) for i, c in enumerate(CT) if c["model"] == "Commit"]


def rng(sheet, c, r1, r2):
    return f"'{sheet}'!${c}${r1}:${c}${r2}"


CTR = lambda c: rng("Contracts", c, C1, CN)
SPR = lambda c: rng("Spend", c, S1, SN)


def header(ws, row, labels, start=1):
    for i, text in enumerate(labels, start=start):
        c = ws.cell(row, i, text)
        c.font, c.fill = HDR, HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 30


def title(ws, text, sub):
    ws["A1"], ws["A2"] = text, sub
    ws["A1"].font, ws["A2"].font = TITLE, Font(italic=True, color="666666")


def widths(ws, w):
    for i, x in enumerate(w, start=1):
        ws.column_dimensions[col(i)].width = x


def put(ws, r, c, v, fmt=None, font=None):
    cell = ws.cell(r, c, v)
    if fmt:
        cell.number_format = fmt
    if font:
        cell.font = font
    return cell


def total_row(ws, r, label, sums, fmt=K):
    put(ws, r, 1, label, font=BOLD)
    for c, formula in sums.items():
        put(ws, r, c, formula, fmt if not isinstance(fmt, dict) else fmt.get(c, K), BOLD)
    for c in range(1, ws.max_column + 1):
        ws.cell(r, c).border = TOP


def name(wb, nm, ref):
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)


# ---------------------------------------------------------------- sheets
def cover(wb):
    ws = wb.active
    ws.title = "Cover"
    title(ws, f"{COMPANY}: Technology Spend, License Utilization & Renewal Analytics",
          "Illustrative portfolio model. Company and data are fictional. Amounts in $ thousands unless noted.")
    lines = [
        ("Purpose", "Gives Procurement, Finance and Technology one view of the technology supplier portfolio: "
                    "where the money goes, which licenses and commits are under-used, which prices are above market, "
                    "which renewals need action now, and how much can be saved."),
        ("How to update", "1) Paste the latest AP spend into Spend (one row per supplier). 2) Update active users and "
                          "contract terms on Contracts. 3) Change AsOfDate and MonthsClosed on Inputs. Everything else recalculates."),
        ("Conventions", "Blue font = input. Black = formula. Savings are annual and in $000s. Unit prices are $ per unit per year."),
        ("", ""),
        ("Sheet", "What it does"),
        ("Dashboard", "Headline KPIs and charts: spend by category, renewal pipeline by quarter, savings by lever"),
        ("Inputs", "As-of date, closed months and analysis thresholds (named ranges)"),
        ("Contracts", "28 technology contracts: terms, notice periods, seats, unit and benchmark prices, commits, InfoSec status"),
        ("Spend", "FY2026 AP spend by supplier and month, including 5 off-contract suppliers"),
        ("Spend Cube", "Spend by category and department, contract coverage, top-10 supplier concentration"),
        ("Utilization", "Seat utilization, under-used licenses, right-sizing and unused license cost"),
        ("Benchmark", "Unit price vs. market benchmark and renegotiation savings on right-sized volumes"),
        ("Consumption", "Cloud, data and AI commits: projected full-year consumption vs. commit (shortfall / overage)"),
        ("Renewals", "Renewal calendar, notice deadlines, risk score, priority, next action and pipeline by quarter"),
        ("RFx", "Weighted evaluation of the observability platform RFx: 3-year TCO, net present cost, scoring"),
        ("Savings", "Savings pipeline by lever, hard savings vs. cost avoidance"),
        ("Checks", "12 integrity checks rolled up to one ALL CHECKS PASS cell"),
    ]
    for r, (a, b) in enumerate(lines, start=4):
        put(ws, r, 1, a, font=BOLD)
        ws.cell(r, 2, b).alignment = Alignment(wrap_text=True, vertical="top")
    for c in (1, 2):
        ws.cell(8, c).font, ws.cell(8, c).fill = HDR, HDR_FILL
    widths(ws, [16, 110])


def inputs(wb):
    ws = wb.create_sheet("Inputs")
    title(ws, "Inputs", "Change these cells only. Each is a named range used across the model.")
    header(ws, 3, ["Name", "Value", "Description"])
    rows = [("AsOfDate", AS_OF, "Analysis date (last closed month end)", DATE),
            ("MonthsClosed", MONTHS_CLOSED, "Closed months of FY spend", "0"),
            ("RFxContract", RFX_CONTRACT, "Contract covered by the RFx (its uplift is counted on the RFx sheet)", None)]
    rows += [(k, v, d, PCT if t == "pct" else "0") for k, (v, d, t) in INPUTS.items()]
    for r, (nm, v, d, fmt) in enumerate(rows, start=4):
        put(ws, r, 1, nm, font=BOLD)
        put(ws, r, 2, v, fmt, INPUT)
        put(ws, r, 3, d)
        name(wb, nm, f"Inputs!$B${r}")
    widths(ws, [18, 13, 75])


def contracts_sheet(wb):
    ws = wb.create_sheet("Contracts")
    title(ws, f"Technology contract register (as of {AS_OF:%b %d, %Y})",
          "ACV = seats x unit price for per-seat deals, the annual commit for consumption deals, or the fixed fee.")
    header(ws, 3, ["ID", "Supplier", "Category", "Product / service", "Owner dept", "Pricing model", "Billing",
                   "Start", "End (renewal)", "Auto-renew", "Notice (days)", "Units", "Unit price ($/yr)",
                   "Benchmark ($/yr)", "Active units", "Commit ($000)", "ACV ($000)", "Proposed uplift", "InfoSec review"])
    for r, c in enumerate(CT, start=C1):
        vals = [c["id"], c["supplier"], c["category"], c["product"], c["dept"], c["model"], c["billing"],
                c["start"], c["end"], c["auto"], c["notice"], c["units"], c["price"], c["bench"], c["active"],
                c["commit"]]
        fmts = [None] * 7 + [DATE, DATE, None, "0", USD, USD, USD, USD, K]
        for j, (v, f) in enumerate(zip(vals, fmts), start=1):
            if v is not None:
                put(ws, r, j, v, f, INPUT)
        if c["model"] == "Per seat":
            put(ws, r, 17, f"=L{r}*M{r}/1000", K)
        elif c["model"] == "Commit":
            put(ws, r, 17, f"=P{r}", K)
        else:
            put(ws, r, 17, c["fixed_acv"], K, INPUT)
        put(ws, r, 18, c["uplift"], PCT, INPUT)
        put(ws, r, 19, c["infosec"], None, INPUT)
    total_row(ws, CN + 1, "Total", {17: f"=SUM(Q{C1}:Q{CN})"})
    widths(ws, [6, 22, 21, 27, 12, 11, 9, 11, 11, 9, 8, 8, 10, 10, 9, 10, 10, 9, 10])
    ws.freeze_panes = "C4"
    ws.conditional_formatting.add(f"S{C1}:S{CN}", CellIsRule(operator="notEqual", formula=['"Current"'], fill=AMBER))


def spend_sheet(wb):
    ws = wb.create_sheet("Spend")
    title(ws, f"FY{FISCAL_YEAR} technology AP spend by supplier ($000s)",
          'Contract ID "None" = off-contract spend. Open months are blank.')
    header(ws, 3, ["Supplier", "Contract ID", "Category", "Owner dept"] + MONTHS + ["YTD", "Last-3-mo avg", "Rank"])
    for r, s in enumerate(SP, start=S1):
        for j, v in enumerate([s["supplier"], s["contract_id"], s["category"], s["dept"]], start=1):
            put(ws, r, j, v, font=INPUT)
        for i, v in enumerate(s["months"]):
            if v is not None:
                put(ws, r, 5 + i, v, K, INPUT)
        put(ws, r, 17, f"=SUM(E{r}:P{r})", K, BOLD)
        put(ws, r, 18, f"=AVERAGE(OFFSET(E{r},0,MonthsClosed-3,1,3))", K)
        put(ws, r, 19, f"=RANK(Q{r},$Q${S1}:$Q${SN})", "0")
    total_row(ws, SN + 1, "Total", {c: f"=SUM({col(c)}{S1}:{col(c)}{SN})" for c in range(5, 18)})
    widths(ws, [22, 10, 21, 12] + [8] * 12 + [10, 10, 6])
    ws.freeze_panes = "E4"
    ws.conditional_formatting.add(f"A{S1}:B{SN}", FormulaRule(formula=[f'$B{S1}="None"'], fill=AMBER))


def spend_cube(wb):
    ws = wb.create_sheet("Spend Cube")
    title(ws, "Spend cube: where the technology money goes", "YTD AP spend by category and department ($000s)")
    header(ws, 4, ["Category", "YTD spend", "% of total", "Suppliers", "Contracted ACV"])
    r1, rn = 5, 5 + len(CATEGORIES) - 1
    for r, cat in enumerate(CATEGORIES, start=r1):
        put(ws, r, 1, cat)
        put(ws, r, 2, f'=SUMIFS({SPR("Q")},{SPR("C")},A{r})', K)
        put(ws, r, 3, f"=B{r}/$B${rn + 1}", PCT)
        put(ws, r, 4, f'=COUNTIFS({SPR("C")},A{r})', "0")
        put(ws, r, 5, f'=SUMIFS({CTR("Q")},{CTR("C")},A{r})', K)
    total_row(ws, rn + 1, "Total", {2: f"=SUM(B{r1}:B{rn})", 3: f"=SUM(C{r1}:C{rn})", 4: f"=SUM(D{r1}:D{rn})",
                                    5: f"=SUM(E{r1}:E{rn})"}, {2: K, 3: PCT, 4: "0", 5: K})
    d0 = rn + 4
    header(ws, d0, ["Owner dept", "YTD spend", "% of total"])
    d1, dn = d0 + 1, d0 + len(DEPARTMENTS)
    for r, d in enumerate(DEPARTMENTS, start=d1):
        put(ws, r, 1, d)
        put(ws, r, 2, f'=SUMIFS({SPR("Q")},{SPR("D")},A{r})', K)
        put(ws, r, 3, f"=B{r}/$B${dn + 1}", PCT)
    total_row(ws, dn + 1, "Total", {2: f"=SUM(B{d1}:B{dn})", 3: f"=SUM(C{d1}:C{dn})"}, {2: K, 3: PCT})
    header(ws, 4, ["Portfolio summary", "Value"], start=7)
    summ = [("Total YTD spend", f'=SUM({SPR("Q")})', K),
            ("Contracted spend", f'=SUMIFS({SPR("Q")},{SPR("B")},"<>None")', K),
            ("Off-contract spend", f'=SUMIFS({SPR("Q")},{SPR("B")},"None")', K),
            ("% of spend under contract", "=H6/H5", PCT),
            ("Top-10 supplier spend", f'=SUMIFS({SPR("Q")},{SPR("S")},"<=10")', K),
            ("Top-10 supplier share", "=H9/H5", PCT),
            ("Active suppliers (YTD spend > 0)", f'=COUNTIFS({SPR("Q")},">0")', "0"),
            ("Annualized run-rate", "=H5/MonthsClosed*12", K)]
    for r, (lab, f, fmt) in enumerate(summ, start=5):
        put(ws, r, 7, lab)
        put(ws, r, 8, f, fmt, BOLD)
    widths(ws, [24, 12, 11, 10, 14, 3, 32, 12])
    return dict(cat=(r1, rn), dept=(d1, dn))


def utilization(wb):
    ws = wb.create_sheet("Utilization")
    title(ws, "License utilization and right-sizing (per-seat contracts)",
          "Under-used = active / purchased below UtilThreshold. Right-size to active users + RightsizeBuffer at renewal.")
    header(ws, 3, ["ID", "Supplier", "Category", "Units purchased", "Active units", "Utilization", "Flag",
                   "Right-sized units", "Units removed", "Unit price ($/yr)", "Annual savings ($000)",
                   "Unused license cost ($000)"])
    r1 = 4
    for k, (i, c) in enumerate(SEAT):
        r, cr = r1 + k, C1 + i
        for j, src in enumerate(["A", "B", "C", "L", "O"], start=1):
            put(ws, r, j, f"=Contracts!{src}{cr}", USD if j > 3 else None)
        put(ws, r, 6, f"=E{r}/D{r}", PCT)
        put(ws, r, 7, f'=IF(F{r}<UtilThreshold,"Under-used","")')
        put(ws, r, 8, f'=IF(G{r}="Under-used",MIN(D{r},ROUNDUP(E{r}*(1+RightsizeBuffer),0)),D{r})', USD)
        put(ws, r, 9, f"=D{r}-H{r}", USD)
        put(ws, r, 10, f"=Contracts!M{cr}", USD)
        put(ws, r, 11, f"=I{r}*J{r}/1000", K)
        put(ws, r, 12, f"=(D{r}-E{r})*J{r}/1000", K)
    rn = r1 + len(SEAT) - 1
    total_row(ws, rn + 1, "Total", {4: f"=SUM(D{r1}:D{rn})", 5: f"=SUM(E{r1}:E{rn})", 6: f"=E{rn + 1}/D{rn + 1}",
                                    7: f'=COUNTIF(G{r1}:G{rn},"Under-used")&" flagged"', 9: f"=SUM(I{r1}:I{rn})",
                                    11: f"=SUM(K{r1}:K{rn})", 12: f"=SUM(L{r1}:L{rn})"},
              {4: USD, 5: USD, 6: PCT, 7: None, 9: USD, 11: K, 12: K})
    ws.conditional_formatting.add(f"F{r1}:G{rn}", FormulaRule(formula=[f'$G{r1}="Under-used"'], fill=RED))
    widths(ws, [6, 18, 18, 11, 10, 11, 11, 11, 10, 11, 12, 13])
    ws.freeze_panes = "C4"
    return r1, rn


def benchmark(wb, u):
    ws = wb.create_sheet("Benchmark")
    title(ws, "Unit price vs. market benchmark (per-seat contracts)",
          "Benchmark = median price paid by peer companies of similar size. Savings are on right-sized volumes.")
    header(ws, 3, ["ID", "Supplier", "Unit price ($/yr)", "Benchmark ($/yr)", "Variance vs. benchmark", "Flag",
                   "Right-sized units", "Renegotiation savings ($000)"])
    r1, rn = u
    for k, (i, c) in enumerate(SEAT):
        r, cr = r1 + k, C1 + i
        put(ws, r, 1, f"=Contracts!A{cr}")
        put(ws, r, 2, f"=Contracts!B{cr}")
        put(ws, r, 3, f"=Contracts!M{cr}", USD)
        put(ws, r, 4, f"=Contracts!N{cr}", USD)
        put(ws, r, 5, f"=C{r}/D{r}-1", PCT)
        put(ws, r, 6, f'=IF(E{r}>BenchTolerance,"Above market","")')
        put(ws, r, 7, f"=Utilization!H{r}", USD)
        put(ws, r, 8, f'=IF(F{r}="Above market",(C{r}-D{r})*G{r}/1000,0)', K)
    total_row(ws, rn + 1, "Total", {6: f'=COUNTIF(F{r1}:F{rn},"Above market")&" flagged"', 8: f"=SUM(H{r1}:H{rn})"},
              {6: None, 8: K})
    ws.conditional_formatting.add(f"E{r1}:F{rn}", FormulaRule(formula=[f'$F{r1}="Above market"'], fill=AMBER))
    widths(ws, [6, 18, 12, 12, 13, 13, 11, 15])
    ws.freeze_panes = "C4"


def consumption(wb):
    ws = wb.create_sheet("Consumption")
    title(ws, "Consumption commits: projected full-year usage vs. commit ($000s)",
          "Projection = YTD + last-3-month average x open months, so ramping usage (e.g. AI) is not understated.")
    header(ws, 3, ["ID", "Supplier", "Annual commit", "YTD consumption", "Last-3-mo avg", "Open months",
                   "Projected FY", "% of commit", "Shortfall (unused commit)", "Overage (on-demand)", "Status",
                   "Overage savings if committed"])
    r1 = 4
    for k, (i, c) in enumerate(COMMIT):
        r, cr = r1 + k, C1 + i
        put(ws, r, 1, f"=Contracts!A{cr}")
        put(ws, r, 2, f"=Contracts!B{cr}")
        put(ws, r, 3, f"=Contracts!P{cr}", K)
        put(ws, r, 4, f'=SUMIFS({SPR("Q")},{SPR("B")},A{r})', K)
        put(ws, r, 5, f'=SUMIFS({SPR("R")},{SPR("B")},A{r})', K)
        put(ws, r, 6, "=12-MonthsClosed", "0")
        put(ws, r, 7, f"=D{r}+E{r}*F{r}", K)
        put(ws, r, 8, f"=G{r}/C{r}", PCT)
        put(ws, r, 9, f"=MAX(0,C{r}-G{r})", K)
        put(ws, r, 10, f"=MAX(0,G{r}-C{r})", K)
        put(ws, r, 11, f'=IF(I{r}>0,"Shortfall risk",IF(J{r}>0,"Overage","On track"))')
        put(ws, r, 12, f"=J{r}*OverageDiscount", K)
    rn = r1 + len(COMMIT) - 1
    total_row(ws, rn + 1, "Total", {c: f"=SUM({col(c)}{r1}:{col(c)}{rn})" for c in (3, 4, 7, 9, 10, 12)})
    ws.conditional_formatting.add(f"K{r1}:K{rn}", CellIsRule(operator="equal", formula=['"Shortfall risk"'], fill=RED))
    ws.conditional_formatting.add(f"K{r1}:K{rn}", CellIsRule(operator="equal", formula=['"Overage"'], fill=AMBER))
    widths(ws, [6, 22, 10, 11, 10, 8, 11, 9, 12, 12, 13, 13])
    return r1, rn


def renewals(wb, u, cons):
    ws = wb.create_sheet("Renewals")
    title(ws, "Renewal calendar, notice deadlines and risk",
          "Risk score: auto-renew +1, notice deadline within ActNowDays +2, under-used +1, above benchmark +1, "
          "InfoSec review not current +1, commit shortfall +1.")
    header(ws, 3, ["ID", "Supplier", "Category", "ACV ($000)", "Renewal date", "Days to renewal", "Auto-renew",
                   "Notice (days)", "Notice deadline", "Days to notice", "In window", "Proposed uplift",
                   "Uplift exposure ($000)", "Avoidable uplift ($000)", "InfoSec review", "Risk score", "Priority",
                   "Next action"])
    ur1, urn = u
    cr1, crn = cons
    for k, c in enumerate(CT):
        r = C1 + k
        put(ws, r, 1, f"=Contracts!A{r}")
        put(ws, r, 2, f"=Contracts!B{r}")
        put(ws, r, 3, f"=Contracts!C{r}")
        put(ws, r, 4, f"=Contracts!Q{r}", K)
        put(ws, r, 5, f"=Contracts!I{r}", DATE)
        put(ws, r, 6, f"=E{r}-AsOfDate", "0")
        put(ws, r, 7, f"=Contracts!J{r}")
        put(ws, r, 8, f"=Contracts!K{r}", "0")
        put(ws, r, 9, f"=E{r}-H{r}", DATE)
        put(ws, r, 10, f"=I{r}-AsOfDate", "0")
        put(ws, r, 11, f'=IF(F{r}<=RenewalWindow,"Yes","No")')
        put(ws, r, 12, f"=Contracts!R{r}", PCT)
        put(ws, r, 13, f"=D{r}*L{r}", K)
        put(ws, r, 14, f'=IF(AND(K{r}="Yes",A{r}<>RFxContract),D{r}*MAX(0,L{r}-TargetUplift),0)', K)
        put(ws, r, 15, f"=Contracts!S{r}")
        put(ws, r, 16, f'=(G{r}="Y")*1+AND(J{r}>=0,J{r}<=ActNowDays)*2'
                       f'+COUNTIFS(Utilization!$A${ur1}:$A${urn},A{r},Utilization!$G${ur1}:$G${urn},"Under-used")'
                       f'+COUNTIFS(Benchmark!$A${ur1}:$A${urn},A{r},Benchmark!$F${ur1}:$F${urn},"Above market")'
                       f'+(O{r}<>"Current")*1'
                       f'+COUNTIFS(Consumption!$A${cr1}:$A${crn},A{r},Consumption!$I${cr1}:$I${crn},">0")', "0")
        put(ws, r, 17, f'=IF(K{r}="Yes",IF(P{r}>=4,"High",IF(P{r}>=2,"Medium","Low")),"Monitor")')
        put(ws, r, 18, f'=IF(K{r}="No","Monitor",IF(J{r}<0,"Notice date passed: negotiate terms",'
                       f'IF(J{r}<=ActNowDays,"Act now: send notice / open negotiation","Prepare sourcing strategy")))')
    total_row(ws, CN + 1, "Total", {4: f"=SUM(D{C1}:D{CN})", 13: f"=SUM(M{C1}:M{CN})", 14: f"=SUM(N{C1}:N{CN})"})
    ws.conditional_formatting.add(f"Q{C1}:Q{CN}", CellIsRule(operator="equal", formula=['"High"'], fill=RED))
    ws.conditional_formatting.add(f"Q{C1}:Q{CN}", CellIsRule(operator="equal", formula=['"Medium"'], fill=AMBER))
    ws.conditional_formatting.add(f"J{C1}:J{CN}", FormulaRule(formula=[f"AND(J{C1}>=0,J{C1}<=ActNowDays)"], fill=RED))

    # Pipeline by quarter and summary (to the right)
    header(ws, 3, ["Quarter", "Start", "End", "Renewals", "ACV ($000)"], start=20)
    for k in range(4):
        r = 4 + k
        start = "=EOMONTH(AsOfDate,0)+1" if k == 0 else f"=V{r - 1}+1"
        put(ws, r, 21, start, DATE)
        put(ws, r, 22, f"=EDATE(U{r},3)-1", DATE)
        put(ws, r, 20, f'="Q"&ROUNDUP(MONTH(U{r})/3,0)&" "&YEAR(U{r})')
        put(ws, r, 23, f'=COUNTIFS($E${C1}:$E${CN},">="&U{r},$E${C1}:$E${CN},"<="&V{r})', "0")
        put(ws, r, 24, f'=SUMIFS($D${C1}:$D${CN},$E${C1}:$E${CN},">="&U{r},$E${C1}:$E${CN},"<="&V{r})', K)
    header(ws, 10, ["Renewal summary", "", "", "Value"], start=20)
    summ = [("Renewals in window (count)", f'=COUNTIF(K{C1}:K{CN},"Yes")', "0"),
            ("ACV renewing in window", f'=SUMIFS(D{C1}:D{CN},K{C1}:K{CN},"Yes")', K),
            ("Notice deadlines within ActNowDays", f'=COUNTIFS(K{C1}:K{CN},"Yes",J{C1}:J{CN},">=0",J{C1}:J{CN},"<="&ActNowDays)', "0"),
            ("ACV behind act-now deadlines", f'=SUMIFS(D{C1}:D{CN},K{C1}:K{CN},"Yes",J{C1}:J{CN},">=0",J{C1}:J{CN},"<="&ActNowDays)', K),
            ("High-priority renewals", f'=COUNTIF(Q{C1}:Q{CN},"High")', "0"),
            ("InfoSec reviews not current", f'=COUNTIF(O{C1}:O{CN},"<>Current")', "0")]
    for k, (lab, f, fmt) in enumerate(summ):
        put(ws, 11 + k, 20, lab)
        put(ws, 11 + k, 23, f, fmt, BOLD)
    widths(ws, [6, 20, 18, 9, 11, 8, 7, 7, 11, 8, 7, 8, 9, 9, 9, 6, 9, 34, 2, 10, 11, 11, 9, 10])
    ws.freeze_panes = "C4"


def rfx(wb):
    ws = wb.create_sheet("RFx")
    title(ws, "Observability platform RFx: weighted evaluation (Sentinel renewal, contract C19)",
          "Price score = 5 x lowest TCO / bidder TCO. Net present cost discounts annual fees at DiscountRate.")
    header(ws, 4, ["Criterion", "Weight"])
    for k, (crit, w) in enumerate(RFX_CRITERIA):
        put(ws, 5 + k, 1, crit)
        put(ws, 5 + k, 2, w, PCT, INPUT)
    wn = 5 + len(RFX_CRITERIA) - 1
    total_row(ws, wn + 1, "Total", {2: f"=SUM(B5:B{wn})"}, PCT)
    b0 = wn + 3
    header(ws, b0, ["Bidder", "Year 1 ($000)", "Year 2", "Year 3", "One-time (migration)", "3-yr TCO",
                    "Net present cost", "Functionality", "Security & compliance", "Integration effort",
                    "Support & SLA", "Price score", "Weighted score", "Rank"])
    b1 = b0 + 1
    bn = b1 + len(RFX_BIDS) - 1
    for k, bid in enumerate(RFX_BIDS):
        r = b1 + k
        put(ws, r, 1, bid[0], font=INPUT)
        for j, v in enumerate(bid[1:5], start=2):
            put(ws, r, j, v, K, INPUT)
        for j, v in enumerate(bid[5:], start=8):
            put(ws, r, j, v, "0.0", INPUT)
        put(ws, r, 6, f"=SUM(B{r}:E{r})", K)
        put(ws, r, 7, f"=E{r}+B{r}/(1+DiscountRate)+C{r}/(1+DiscountRate)^2+D{r}/(1+DiscountRate)^3", K)
        put(ws, r, 12, f"=5*MIN($F${b1}:$F${bn})/F{r}", "0.00")
        put(ws, r, 13, f"=H{r}*$B$5+I{r}*$B$6+J{r}*$B$7+K{r}*$B$8+L{r}*$B$9", "0.00", BOLD)
        put(ws, r, 14, f"=RANK(M{r},$M${b1}:$M${bn})", "0")
    ws.conditional_formatting.add(f"A{b1}:N{bn}", FormulaRule(formula=[f"$N{b1}=1"], fill=GREEN))
    s0 = bn + 2
    rows = [("Recommended bidder", f"=INDEX(A{b1}:A{bn},MATCH(MAX(M{b1}:M{bn}),M{b1}:M{bn},0))", None),
            ("3-yr TCO, renewal quote", f"=F{b1}", K),
            ("3-yr TCO, recommended", f"=INDEX(F{b1}:F{bn},MATCH(B{s0 + 1},A{b1}:A{bn},0))", K),
            ("3-yr savings vs. renewal quote", f"=B{s0 + 2}-B{s0 + 3}", K),
            ("Annualized savings", f"=B{s0 + 4}/3", K)]
    for k, (lab, f, fmt) in enumerate(rows):
        put(ws, s0 + 1 + k, 1, lab, font=BOLD)
        put(ws, s0 + 1 + k, 2, f, fmt, BOLD)
    widths(ws, [28, 13, 10, 10, 11, 10, 11, 11, 11, 11, 10, 9, 10, 6])
    return dict(weights=(5, wn), bids=(b1, bn), rec=f"RFx!$B${s0 + 1}", annual=f"RFx!$B${s0 + 5}",
                total_w=f"RFx!$B${wn + 1}")


def savings(wb, u, cons, rf, cube_off):
    ws = wb.create_sheet("Savings")
    title(ws, "Savings pipeline (annual, $000s)",
          "Hard savings reduce spend vs. today; cost avoidance prevents an increase. Uplift avoidance is on current ACV.")
    header(ws, 3, ["Initiative", "Lever", "Type", "Annual savings ($000)", "Source"])
    ur1, urn = u
    cr1, crn = cons
    items = [
        ("License right-sizing", "Reduce seats on under-used licenses at renewal", "Hard savings", f"=Utilization!K{urn + 1}", "Utilization"),
        ("Price-to-benchmark renegotiation", "Bring unit prices to market benchmark", "Hard savings", f"=Benchmark!H{urn + 1}", "Benchmark"),
        ("Renewal uplift avoidance", "Cap renewal increases at TargetUplift", "Cost avoidance", f"=Renewals!N{CN + 1}", "Renewals"),
        ("Observability RFx (annualized)", "Competitive RFx vs. incumbent renewal quote", "Cost avoidance", f"={rf['annual']}", "RFx"),
        ("Overage converted to commit", "Commit projected overage at discounted rates", "Hard savings", f"=Consumption!L{crn + 1}", "Consumption"),
        ("Commit right-sizing at renewal", "Lower commits that will not be consumed", "Hard savings", f"=Consumption!I{crn + 1}", "Consumption"),
        ("Off-contract spend consolidation", "Move tail spend onto contracted suppliers", "Hard savings", f"={cube_off}/MonthsClosed*12*TailSavingsRate", "Spend Cube"),
    ]
    for k, (a, b, t, f, s) in enumerate(items):
        r = 4 + k
        put(ws, r, 1, a)
        put(ws, r, 2, b)
        put(ws, r, 3, t)
        put(ws, r, 4, f, K)
        put(ws, r, 5, s)
    rn = 4 + len(items) - 1
    total_row(ws, rn + 1, "Total identified savings", {4: f"=SUM(D4:D{rn})"})
    put(ws, rn + 2, 1, "Hard savings")
    put(ws, rn + 2, 4, f'=SUMIFS(D4:D{rn},C4:C{rn},"Hard savings")', K)
    put(ws, rn + 3, 1, "Cost avoidance")
    put(ws, rn + 3, 4, f'=SUMIFS(D4:D{rn},C4:C{rn},"Cost avoidance")', K)
    put(ws, rn + 4, 1, "% of contracted ACV", font=BOLD)
    put(ws, rn + 4, 4, f"=D{rn + 1}/SUM({CTR('Q')})", PCT, BOLD)
    widths(ws, [32, 46, 15, 14, 13])
    return dict(items=(4, rn), total=f"Savings!$D${rn + 1}", pct=f"Savings!$D${rn + 4}")


def dashboard(wb, cube, sv, rf):
    ws = wb.create_sheet("Dashboard", 1)
    title(ws, f"{COMPANY}: technology procurement dashboard",
          f"FY{FISCAL_YEAR} YTD through MonthsClosed; renewals measured from AsOfDate. $000s unless noted.")
    header(ws, 4, ["KPI", "Value"], start=2)
    kpis = [("YTD technology spend", "='Spend Cube'!H5", K),
            ("Annualized run-rate", "='Spend Cube'!H12", K),
            ("Contracted ACV", f"=SUM({CTR('Q')})", K),
            ("Active technology suppliers", "='Spend Cube'!H11", "0"),
            ("% of spend under contract", "='Spend Cube'!H8", PCT),
            ("Top-10 supplier share of spend", "='Spend Cube'!H10", PCT),
            ("Weighted license utilization", f"=Utilization!F{4 + len(SEAT)}", PCT),
            ("Unused license cost (annual)", f"=Utilization!L{4 + len(SEAT)}", K),
            ("Renewals in next RenewalWindow days", "=Renewals!W11", "0"),
            ("ACV renewing in window", "=Renewals!W12", K),
            ("Notice deadlines needing action now", "=Renewals!W13", "0"),
            ("Identified annual savings", f"={sv['total']}", K),
            ("Savings as % of contracted ACV", f"={sv['pct']}", PCT),
            ("Observability RFx recommendation", f"={rf['rec']}", None)]
    for k, (lab, f, fmt) in enumerate(kpis):
        put(ws, 5 + k, 2, lab)
        put(ws, 5 + k, 3, f, fmt, Font(bold=True, size=12, color=NAVY))
    widths(ws, [2, 38, 26])

    def bar(title_, data_ref, cats_ref, anchor, horizontal=True):
        ch = BarChart()
        ch.type = "bar" if horizontal else "col"
        ch.title, ch.legend, ch.height, ch.width = title_, None, 7.5, 16
        ch.add_data(data_ref, titles_from_data=False)
        ch.set_categories(cats_ref)
        ch.y_axis.numFmt = "#,##0"
        ch.y_axis.majorGridlines = None
        ws.add_chart(ch, anchor)

    c1, cn = cube["cat"]
    sc = wb["Spend Cube"]
    bar("YTD spend by category ($000s)", Reference(sc, min_col=2, min_row=c1, max_row=cn),
        Reference(sc, min_col=1, min_row=c1, max_row=cn), "E4")
    rn = wb["Renewals"]
    bar("Renewal pipeline: ACV by quarter ($000s)", Reference(rn, min_col=24, min_row=4, max_row=7),
        Reference(rn, min_col=20, min_row=4, max_row=7), "E20", horizontal=False)
    s1, sn = sv["items"]
    svs = wb["Savings"]
    bar("Identified annual savings by lever ($000s)", Reference(svs, min_col=4, min_row=s1, max_row=sn),
        Reference(svs, min_col=1, min_row=s1, max_row=sn), "N4")


def checks(wb, cube, u, cons, rf, sv):
    ws = wb.create_sheet("Checks")
    title(ws, "Integrity checks", "Each check returns TRUE when it passes.")
    header(ws, 3, ["Check", "Pass?"])
    c1, cn = cube["cat"]
    d1, dn = cube["dept"]
    ur1, urn = u
    cr1, crn = cons
    w1, wn = rf["weights"]
    s1, sn = sv["items"]
    tests = [
        ("Category spend sums to total spend", f"=ABS('Spend Cube'!B{cn + 1}-'Spend Cube'!H5)<0.01"),
        ("Department spend sums to total spend", f"=ABS('Spend Cube'!B{dn + 1}-'Spend Cube'!H5)<0.01"),
        ("Contracted + off-contract = total spend", "=ABS('Spend Cube'!H6+'Spend Cube'!H7-'Spend Cube'!H5)<0.01"),
        ("Category ACV sums to contract register ACV", f"=ABS('Spend Cube'!E{cn + 1}-SUM({CTR('Q')}))<0.01"),
        ("Every contract ID on Spend exists on Contracts", f'=SUMPRODUCT(({SPR("B")}<>"None")*(COUNTIF({CTR("A")},{SPR("B")})=0))=0'),
        ("Every contract has a Spend row", f'=SUMPRODUCT(--(COUNTIF({SPR("B")},{CTR("A")})=0))=0'),
        ("Spend entered for every closed month and no open month",
         f"=COUNT(Spend!$E${S1}:$P${SN})=ROWS(Spend!$E${S1}:$E${SN})*MonthsClosed"),
        ("Utilization between 0% and 100%", f"=AND(MIN(Utilization!F{ur1}:F{urn})>=0,MAX(Utilization!F{ur1}:F{urn})<=1)"),
        ("Right-sized seats never exceed purchased seats", f"=SUMPRODUCT(--(Utilization!H{ur1}:H{urn}>Utilization!D{ur1}:D{urn}))=0"),
        ("No contract expired before the as-of date", f'=COUNTIF({CTR("I")},"<"&AsOfDate)=0'),
        ("RFx criteria weights sum to 100%", f"=ABS({rf['total_w']}-1)<0.0001"),
        ("Hard savings + cost avoidance = total savings",
         f"=ABS(Savings!D{sn + 2}+Savings!D{sn + 3}-Savings!D{sn + 1})<0.01"),
    ]
    for k, (lab, f) in enumerate(tests):
        put(ws, 4 + k, 1, lab)
        put(ws, 4 + k, 2, f)
    last = 4 + len(tests) - 1
    put(ws, last + 2, 1, "OVERALL", font=BOLD)
    put(ws, last + 2, 2, f'=IF(AND(B4:B{last}),"ALL CHECKS PASS","CHECK FAILED")', font=BOLD)
    ws.conditional_formatting.add(f"B4:B{last}", CellIsRule(operator="equal", formula=["FALSE"], fill=RED))
    ws.conditional_formatting.add(f"B{last + 2}", CellIsRule(operator="equal", formula=['"ALL CHECKS PASS"'], fill=GREEN))
    widths(ws, [52, 18])


def main():
    wb = Workbook()
    cover(wb)
    inputs(wb)
    contracts_sheet(wb)
    spend_sheet(wb)
    cube = spend_cube(wb)
    u = utilization(wb)
    benchmark(wb, u)
    cons = consumption(wb)
    renewals(wb, u, cons)
    rf = rfx(wb)
    sv = savings(wb, u, cons, rf, "'Spend Cube'!$H$7")
    dashboard(wb, cube, sv, rf)
    checks(wb, cube, u, cons, rf, sv)
    wb.calculation.fullCalcOnLoad = True
    OUT.parent.mkdir(exist_ok=True)
    wb.save(OUT)
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    main()
