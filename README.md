# Technology Spend, License Utilization & Renewal Analytics

**Portfolio project | Procurement analytics / technology sourcing**

An Excel model of the technology supplier portfolio at **Northpeak Interactive**, a fictional digital entertainment company with about 1,500 employees. It answers the questions a technology procurement analyst works on every month:

- **Where does the technology money go?** A spend cube by category and department, contract coverage, and supplier concentration.
- **What are we paying for but not using?** License utilization, right-sizing at renewal, and unused license cost.
- **Are we paying market rates?** Unit price compared with benchmarks, and the savings from renegotiating.
- **Will we use what we committed to?** Cloud, data and AI consumption commits projected to year-end, flagging shortfalls and overages.
- **Which renewals need action now?** A renewal calendar with notice deadlines, a risk score, a priority, the next action, and the pipeline by quarter.
- **Which vendor should win the RFx?** A weighted evaluation of an observability platform RFx, using 3-year TCO and net present cost.
- **How much can we save?** A savings pipeline by lever, split into hard savings and cost avoidance.

Every number is a live Excel formula driven by four input sheets: Contracts, Spend, Inputs and the RFx bids. A separate Python implementation recalculates every output, and a script checks that the two agree.

> **Illustrative data.** Northpeak Interactive, its suppliers and all figures are fictional. The data is generated from a fixed seed (`tools/assumptions.py`), with realistic contract terms, utilization patterns and consumption trends built in. Amounts are in $ thousands, and unit prices are $ per unit per year.

📥 **[Download the workbook](model/Tech_Spend_Renewal_Analytics.xlsx)** (open in Excel; it recalculates on open)

---

## Findings (as of September 30, 2026)

| Portfolio | |
|---|---|
| YTD technology spend (9 months) | $27.4M (run-rate $36.6M) |
| Contracted annual contract value (ACV) | $34.3M across 28 contracts |
| Active technology suppliers | 33 |
| Spend under contract | 97.2% (off-contract spend: $773K across 5 suppliers) |
| Top-10 supplier share of spend | 81.3% |
| Largest categories | Cloud & infrastructure 33%, Technology services 19%, Data 18% |

**Licenses.** Weighted utilization is 86%, but 5 of 13 per-seat contracts are under 80%. Voxline is at 48%, TaskFlow at 60%, Docuvault at 62%, Pulse Analytics at 65% and TicketHub at 75%. Unused licenses cost **$491K a year**. Right-sizing those five contracts at renewal, while keeping 10% headroom, removes 894 seats and saves **$310K a year**.

**Pricing.** 10 of 13 per-seat contracts are more than 5% above the benchmark. The biggest gaps are Voxline (+29%), Pulse Analytics (+20%) and TaskFlow (+19%). Renegotiating to benchmark, on right-sized volumes, saves **$189K**.

**Consumption commits.** Stratus Cloud is projected at 110% of its $9.6M commit ($962K overage at on-demand rates), and Quarry Data at 121%. Nimbus Edge and BuildBay will leave $403K and $150K of their commits unused. Cortex AI looks under-used on a simple YTD run-rate, but on the last three months it is on track, so the projection method matters for ramping AI workloads.

**Renewals.** 14 contracts worth **$10.9M** renew in the next 180 days. Seven of them, worth $3.9M, have **notice deadlines within 60 days**. TaskFlow's deadline is tomorrow and IdentiKey's is in 2 days, and both auto-renew. Seven renewals are rated high priority, and 5 contracts have InfoSec reviews that are due or expired. Those include the two AI tools, which need a security and privacy review before renewal.

**Observability RFx.** The incumbent's renewal quote included a 12% increase. With competitive bids from two alternatives, the incumbent's best-and-final offer scores highest (4.49 of 5). It keeps the lowest integration risk and saves **$1.18M over 3 years** (**$394K a year**) against the original quote. Beacon Logs is cheapest but scores lowest on functionality and security.

**Savings pipeline: $1.92M a year, 5.6% of contracted ACV**

| Lever | Type | Annual savings ($000s) |
|---|---|---|
| Commit right-sizing at renewal | Hard savings | 566 |
| Observability RFx (annualized) | Cost avoidance | 394 |
| License right-sizing | Hard savings | 310 |
| Overage converted to commit | Hard savings | 253 |
| Price-to-benchmark renegotiation | Hard savings | 189 |
| Renewal uplift avoidance (cap at 3%) | Cost avoidance | 103 |
| Off-contract spend consolidation | Hard savings | 103 |
| **Total** | | **1,919** |

## Workbook structure

| Sheet | What it does |
|---|---|
| **Cover** | Purpose, how to update, conventions, sheet guide |
| **Dashboard** | 14 KPIs and charts: spend by category, renewal ACV by quarter, savings by lever |
| **Inputs** | As-of date, closed months, and thresholds as named ranges (utilization, benchmark tolerance, renewal window, notice window, target uplift, discount rate) |
| **Contracts** | Contract register: category, owner, pricing model, dates, auto-renew, notice period, seats, unit and benchmark prices, commits, ACV, proposed uplift, InfoSec status |
| **Spend** | FY2026 AP spend by supplier and month, including off-contract suppliers |
| **Spend Cube** | Spend by category and department, contracted vs. off-contract spend, top-10 concentration (SUMIFS, COUNTIFS, RANK) |
| **Utilization** | Utilization %, under-used flag, right-sized seats, savings, unused license cost |
| **Benchmark** | Variance vs. benchmark, above-market flag, renegotiation savings |
| **Consumption** | Projected full-year consumption vs. commit, shortfall and overage, status |
| **Renewals** | Days to renewal and notice deadline, in-window flag, uplift exposure, risk score, priority, next action, pipeline by quarter |
| **RFx** | Criteria weights, 3-year TCO, net present cost, price score, weighted score, rank, recommendation |
| **Savings** | Savings by lever, hard savings vs. cost avoidance, % of ACV |
| **Checks** | 12 integrity checks rolled up to one ALL CHECKS PASS cell |

## Methodology

- **Right-sizing:** a contract under the utilization threshold (80%) is right-sized to active users plus 10% headroom, never above current seats. Savings = seats removed × unit price.
- **Benchmarking:** savings are calculated on right-sized volumes, so they are not double-counted with right-sizing.
- **Consumption projection:** YTD plus the last-3-month average × open months. For ramping workloads this is more accurate than YTD ÷ months closed.
- **Risk score:** auto-renew +1, notice deadline within 60 days +2, under-used +1, above benchmark +1, InfoSec review not current +1, commit shortfall +1. In-window renewals scoring 4 or more are high priority.
- **Renewal uplift avoidance:** in-window renewals with a proposed increase above 3% are capped at 3%. The RFx contract is excluded because the RFx sheet already counts its savings. Uplift avoidance is calculated on current ACV and slightly overlaps right-sizing for the three contracts that have both.
- **RFx scoring:** functionality 30%, security & compliance 20%, integration effort 15%, support & SLA 10%, price 25%. Price score = 5 × lowest TCO ÷ bidder TCO. Net present cost discounts annual fees at 9%.
- **Integrity checks:** category and department spend tie to the total; contracted plus off-contract equals the total; category ACV ties to the register; every spend row maps to a contract and every contract has spend; every closed month is filled in and no open month is; utilization stays between 0% and 100%; right-sized seats never exceed purchased seats; no expired contracts; RFx weights sum to 100%; and savings types sum to the total.

## Verification

`tools/reference_model.py` is an independent pandas implementation of every calculation. `tools/recalc_and_verify.sh` recalculates the workbook headless in LibreOffice and compares 372 values: spend cube, utilization, benchmarks, consumption, every renewal's dates, score and priority, the quarterly pipeline, RFx scores and the recommendation, the savings levers, and the Checks sheet.

```
Values compared: 372  Checks sheet: ALL CHECKS PASS
ALL MATCH
```

## Reproduce

```bash
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
(cd tools && ../.venv/bin/python build_workbook.py)     # writes the workbook
(cd tools && ../.venv/bin/python reference_model.py)    # Python results
tools/recalc_and_verify.sh                              # needs LibreOffice
```

---

Built by [Kshama Upadhyay](https://github.com/Kshama16).
