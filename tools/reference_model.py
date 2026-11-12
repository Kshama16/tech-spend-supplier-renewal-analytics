"""Independent pandas implementation of every calculation in the workbook.

verify_workbook.py compares the recalculated Excel file against these results.
Run from tools/:  ../.venv/bin/python reference_model.py
"""
import math
import pandas as pd

from assumptions import (INPUTS, AS_OF, MONTHS_CLOSED, CATEGORIES, DEPARTMENTS, RFX_BIDS,
                         RFX_CRITERIA, RFX_CONTRACT, contracts, spend_rows, acv)

P = {k: v[0] for k, v in INPUTS.items()}


def contract_table():
    df = pd.DataFrame(contracts())
    df["acv"] = [acv(c) for c in contracts()]
    return df


def spend_table():
    rows = spend_rows()
    df = pd.DataFrame([{k: v for k, v in r.items() if k != "months"} for r in rows])
    closed = [[x for x in r["months"][:MONTHS_CLOSED]] for r in rows]
    df["ytd"] = [sum(m) for m in closed]
    df["l3m"] = [sum(m[-3:]) / 3 for m in closed]
    df["rank"] = df["ytd"].rank(ascending=False, method="min").astype(int)
    return df


def spend_cube(sp, ct):
    total = sp["ytd"].sum()
    cat = pd.DataFrame({"category": CATEGORIES})
    cat["ytd"] = [sp.loc[sp.category == c, "ytd"].sum() for c in CATEGORIES]
    cat["share"] = cat["ytd"] / total
    cat["suppliers"] = [int((sp.category == c).sum()) for c in CATEGORIES]
    cat["acv"] = [ct.loc[ct.category == c, "acv"].sum() for c in CATEGORIES]
    dept = pd.DataFrame({"dept": DEPARTMENTS})
    dept["ytd"] = [sp.loc[sp.dept == d, "ytd"].sum() for d in DEPARTMENTS]
    dept["share"] = dept["ytd"] / total
    off = sp.loc[sp.contract_id == "None", "ytd"].sum()
    summary = dict(total=total, contracted=total - off, off_contract=off,
                   pct_contracted=(total - off) / total,
                   top10=sp.loc[sp["rank"] <= 10, "ytd"].sum(),
                   top10_share=sp.loc[sp["rank"] <= 10, "ytd"].sum() / total,
                   suppliers=int((sp.ytd > 0).sum()))
    return cat, dept, summary


def utilization(ct):
    u = ct[ct.model == "Per seat"].copy()
    u["util"] = u.active / u.units
    u["flag"] = u.util < P["UtilThreshold"]
    # round() first so float noise (820 * 1.1 = 902.0000000000001) doesn't add a seat
    u["rightsized"] = [min(r.units, math.ceil(round(r.active * (1 + P["RightsizeBuffer"]), 6))) if r.flag else r.units
                       for r in u.itertuples()]
    u["removed"] = u.units - u.rightsized
    u["savings"] = u.removed * u.price / 1000
    u["unused_cost"] = (u.units - u.active) * u.price / 1000
    return u.reset_index(drop=True)


def benchmark(u):
    b = u[["id", "supplier", "price", "bench", "rightsized"]].copy()
    b["var"] = b.price / b.bench - 1
    b["flag"] = b["var"] > P["BenchTolerance"]
    b["savings"] = [(r.price - r.bench) * r.rightsized / 1000 if r.flag else 0.0 for r in b.itertuples()]
    return b


def consumption(ct, sp):
    c = ct[ct.model == "Commit"][["id", "supplier", "commit"]].copy().reset_index(drop=True)
    c["ytd"] = [sp.loc[sp.contract_id == i, "ytd"].sum() for i in c.id]
    c["l3m"] = [sp.loc[sp.contract_id == i, "l3m"].sum() for i in c.id]
    c["projected"] = c.ytd + c.l3m * (12 - MONTHS_CLOSED)
    c["pct"] = c.projected / c.commit
    c["shortfall"] = (c.commit - c.projected).clip(lower=0)
    c["overage"] = (c.projected - c.commit).clip(lower=0)
    c["status"] = ["Shortfall risk" if s > 0 else ("Overage" if o > 0 else "On track")
                   for s, o in zip(c.shortfall, c.overage)]
    c["overage_savings"] = c.overage * P["OverageDiscount"]
    return c


def renewals(ct, u, b, c):
    r = ct[["id", "supplier", "category", "acv", "end", "auto", "notice", "uplift", "infosec"]].copy()
    asof = pd.Timestamp(AS_OF)
    r["days_to_renewal"] = [(pd.Timestamp(e) - asof).days for e in r.end]
    r["notice_deadline"] = [pd.Timestamp(e) - pd.Timedelta(days=n) for e, n in zip(r.end, r.notice)]
    r["days_to_notice"] = [(d - asof).days for d in r.notice_deadline]
    r["in_window"] = r.days_to_renewal <= P["RenewalWindow"]
    r["uplift_exposure"] = r.acv * r.uplift
    r["avoidable_uplift"] = [a * max(0.0, up - P["TargetUplift"]) if w and i != RFX_CONTRACT else 0.0
                             for a, up, w, i in zip(r.acv, r.uplift, r.in_window, r.id)]
    util_flag = set(u.loc[u.flag, "id"])
    bench_flag = set(b.loc[b.flag, "id"])
    short_flag = set(c.loc[c.shortfall > 0, "id"])
    r["score"] = [(a == "Y") + 2 * (0 <= dn <= P["ActNowDays"]) + (i in util_flag) + (i in bench_flag)
                  + (s != "Current") + (i in short_flag)
                  for a, dn, i, s in zip(r.auto, r.days_to_notice, r.id, r.infosec)]
    r["priority"] = ["High" if w and s >= 4 else "Medium" if w and s >= 2 else "Low" if w else "Monitor"
                     for w, s in zip(r.in_window, r.score)]
    return r


def pipeline_by_quarter(r):
    starts = [pd.Timestamp(2026, 10, 1), pd.Timestamp(2027, 1, 1), pd.Timestamp(2027, 4, 1), pd.Timestamp(2027, 7, 1)]
    out = []
    for s in starts:
        e = s + pd.DateOffset(months=3) - pd.Timedelta(days=1)
        m = (pd.to_datetime(r.end) >= s) & (pd.to_datetime(r.end) <= e)
        out.append(dict(start=s, count=int(m.sum()), acv=r.loc[m, "acv"].sum()))
    return pd.DataFrame(out)


def rfx():
    rows = []
    r = P["DiscountRate"]
    for bidder, y1, y2, y3, one, *scores in RFX_BIDS:
        tco = y1 + y2 + y3 + one
        npc = one + y1 / (1 + r) + y2 / (1 + r) ** 2 + y3 / (1 + r) ** 3
        rows.append(dict(bidder=bidder, tco=tco, npc=npc, scores=scores))
    df = pd.DataFrame(rows)
    min_tco = df.tco.min()
    df["price_score"] = 5 * min_tco / df.tco
    w = [x[1] for x in RFX_CRITERIA]
    df["weighted"] = [sum(s * wi for s, wi in zip(sc + [ps], w)) for sc, ps in zip(df.scores, df.price_score)]
    best = df.loc[df.weighted.idxmax()]
    savings_3yr = df.loc[0, "tco"] - best.tco
    return df, best.bidder, savings_3yr


def savings_pipeline(u, b, r, c, sp_summary, rfx_savings):
    tail_annual = sp_summary["off_contract"] / MONTHS_CLOSED * 12
    return {
        "License right-sizing": u.savings.sum(),
        "Price-to-benchmark renegotiation": b.savings.sum(),
        "Renewal uplift avoidance": r.avoidable_uplift.sum(),
        "Observability RFx (annualized)": rfx_savings / 3,
        "Overage converted to commit": c.overage_savings.sum(),
        "Commit right-sizing at renewal": c.shortfall.sum(),
        "Off-contract spend consolidation": tail_annual * P["TailSavingsRate"],
    }


def run():
    ct, sp = contract_table(), spend_table()
    cat, dept, summ = spend_cube(sp, ct)
    u = utilization(ct)
    b = benchmark(u)
    c = consumption(ct, sp)
    r = renewals(ct, u, b, c)
    q = pipeline_by_quarter(r)
    rf, rec, rfx_sav = rfx()
    sv = savings_pipeline(u, b, r, c, summ, rfx_sav)
    return dict(ct=ct, sp=sp, cat=cat, dept=dept, summ=summ, u=u, b=b, c=c, r=r, q=q,
                rfx=rf, rfx_rec=rec, rfx_sav=rfx_sav, savings=sv)


if __name__ == "__main__":
    pd.set_option("display.width", 200, "display.max_columns", 30)
    o = run()
    print("SPEND SUMMARY", {k: round(v, 3) for k, v in o["summ"].items()})
    print(o["cat"].round(3)); print(o["dept"].round(3))
    print(o["u"][["id", "supplier", "units", "active", "util", "flag", "rightsized", "savings", "unused_cost"]].round(3))
    print(o["b"].round(3))
    print(o["c"].round(1))
    print(o["r"][["id", "supplier", "acv", "days_to_renewal", "days_to_notice", "in_window", "avoidable_uplift", "score", "priority"]].round(1))
    print(o["q"])
    print(o["rfx"][["bidder", "tco", "npc", "price_score", "weighted"]].round(3), "->", o["rfx_rec"], round(o["rfx_sav"], 1))
    tot = sum(o["savings"].values())
    for k, v in o["savings"].items():
        print(f"  {k:38s} {v:8.1f}")
    print("TOTAL SAVINGS", round(tot, 1), "ACV", round(o["ct"].acv.sum(), 1), "pct", round(tot / o["ct"].acv.sum(), 4))
