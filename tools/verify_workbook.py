"""Compares a recalculated workbook against the Python reference model.

Usage (from tools/):  ../.venv/bin/python verify_workbook.py <recalculated.xlsx>
"""
import sys
from openpyxl import load_workbook

from assumptions import CATEGORIES, DEPARTMENTS
from reference_model import run

TOL = 0.05  # $000s


def main(path):
    wb = load_workbook(path, data_only=True)
    o = run()
    fails = n = 0

    def check(label, got, want, tol=TOL):
        nonlocal fails, n
        n += 1
        if isinstance(want, str) or isinstance(want, bool):
            ok = got == want
        else:
            ok = got is not None and not isinstance(got, str) and abs(got - want) <= tol
        if not ok:
            fails += 1
            print(f"MISMATCH {label}: workbook {got!r} reference {want!r}")

    cube = wb["Spend Cube"]
    for k, row in o["cat"].iterrows():
        check(f"cat {row.category} ytd", cube.cell(5 + k, 2).value, row.ytd)
        check(f"cat {row.category} acv", cube.cell(5 + k, 5).value, row.acv)
        check(f"cat {row.category} suppliers", cube.cell(5 + k, 4).value, row.suppliers)
    d1 = 5 + len(CATEGORIES) + 4
    for k, row in o["dept"].iterrows():
        check(f"dept {row.dept}", cube.cell(d1 + k, 2).value, row.ytd)
    s = o["summ"]
    for r, key in zip(range(5, 13), ["total", "contracted", "off_contract", "pct_contracted", "top10",
                                     "top10_share", "suppliers", None]):
        if key:
            check(f"summary {key}", cube.cell(r, 8).value, s[key], 0.0005 if "pct" in key or "share" in key else TOL)

    ut = wb["Utilization"]
    for k, row in o["u"].iterrows():
        r = 4 + k
        check(f"util {row.id}", ut.cell(r, 6).value, row.util, 1e-6)
        check(f"util flag {row.id}", ut.cell(r, 7).value or "", "Under-used" if row.flag else "")
        check(f"rightsized {row.id}", ut.cell(r, 8).value, row.rightsized, 1e-9)
        check(f"util savings {row.id}", ut.cell(r, 11).value, row.savings)
        check(f"unused {row.id}", ut.cell(r, 12).value, row.unused_cost)

    bm = wb["Benchmark"]
    for k, row in o["b"].iterrows():
        check(f"bench var {row.id}", bm.cell(4 + k, 5).value, row["var"], 1e-6)
        check(f"bench savings {row.id}", bm.cell(4 + k, 8).value, row.savings)

    cs = wb["Consumption"]
    for k, row in o["c"].iterrows():
        r = 4 + k
        for c, key in [(4, "ytd"), (5, "l3m"), (7, "projected"), (9, "shortfall"), (10, "overage"), (12, "overage_savings")]:
            check(f"commit {row.id} {key}", cs.cell(r, c).value, row[key])
        check(f"commit {row.id} status", cs.cell(r, 11).value, row.status)

    rn = wb["Renewals"]
    for k, row in o["r"].iterrows():
        r = 4 + k
        check(f"ren {row.id} days", rn.cell(r, 6).value, row.days_to_renewal, 1e-9)
        check(f"ren {row.id} notice days", rn.cell(r, 10).value, row.days_to_notice, 1e-9)
        check(f"ren {row.id} window", rn.cell(r, 11).value, "Yes" if row.in_window else "No")
        check(f"ren {row.id} avoidable", rn.cell(r, 14).value, row.avoidable_uplift)
        check(f"ren {row.id} score", rn.cell(r, 16).value, row.score, 1e-9)
        check(f"ren {row.id} priority", rn.cell(r, 17).value, row.priority)
    for k, row in o["q"].iterrows():
        check(f"quarter {k} count", rn.cell(4 + k, 23).value, row["count"], 1e-9)
        check(f"quarter {k} acv", rn.cell(4 + k, 24).value, row.acv)

    rf = wb["RFx"]
    b1 = 5 + 5 + 3  # weights rows 5-9, total row 10, header row 12
    for k, row in o["rfx"].iterrows():
        check(f"rfx {row.bidder} tco", rf.cell(b1 + k, 6).value, row.tco)
        check(f"rfx {row.bidder} npc", rf.cell(b1 + k, 7).value, row.npc)
        check(f"rfx {row.bidder} score", rf.cell(b1 + k, 13).value, row.weighted, 1e-6)
    s0 = b1 + len(o["rfx"]) + 1
    check("rfx recommendation", rf.cell(s0 + 1, 2).value, o["rfx_rec"])
    check("rfx 3yr savings", rf.cell(s0 + 4, 2).value, o["rfx_sav"])

    sv = wb["Savings"]
    for k, (lab, v) in enumerate(o["savings"].items()):
        check(f"savings {lab}", sv.cell(4 + k, 4).value, v)
    total = sum(o["savings"].values())
    check("savings total", sv.cell(4 + len(o["savings"]), 4).value, total)

    ck = wb["Checks"]
    overall = [ck.cell(r, 2).value for r in range(1, ck.max_row + 1) if ck.cell(r, 1).value == "OVERALL"][0]
    print(f"Values compared: {n}  Checks sheet: {overall}")
    print("ALL MATCH" if fails == 0 and overall == "ALL CHECKS PASS" else f"{fails} MISMATCHES")
    return 1 if fails or overall != "ALL CHECKS PASS" else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))
