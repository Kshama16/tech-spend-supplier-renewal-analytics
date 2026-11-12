"""Illustrative data for the technology spend and renewal model.

Northpeak Interactive is a fictional digital entertainment company with about 1,500 employees.
Its technology supplier portfolio (28 contracts plus 5 off-contract suppliers) and FY2026
accounts payable spend are simulated from a fixed seed. All amounts are in $ thousands
except unit prices, which are $ per unit per year.

This file feeds both the Excel workbook (build_workbook.py) and the Python reference model
(reference_model.py), so the two always start from identical inputs.
"""
from datetime import date
import numpy as np

COMPANY = "Northpeak Interactive"
FISCAL_YEAR = 2026
AS_OF = date(2026, 9, 30)
MONTHS_CLOSED = 9
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# Analysis thresholds (Inputs sheet, named ranges)
INPUTS = {
    "UtilThreshold": (0.80, "License utilization below this is flagged as under-used", "pct"),
    "RightsizeBuffer": (0.10, "Headroom kept above active users when right-sizing", "pct"),
    "BenchTolerance": (0.05, "Unit price more than this above benchmark is flagged", "pct"),
    "RenewalWindow": (180, "Renewals within this many days are in the sourcing pipeline", "int"),
    "ActNowDays": (60, "Notice deadlines within this many days need action now", "int"),
    "TargetUplift": (0.03, "Maximum renewal price increase procurement will accept", "pct"),
    "OverageDiscount": (0.15, "Commit discount vs. on-demand rates for consumption overages", "pct"),
    "TailSavingsRate": (0.10, "Savings from consolidating off-contract spend onto contracts", "pct"),
    "DiscountRate": (0.09, "Discount rate for RFx net present cost", "pct"),
}

CATEGORIES = ["SaaS", "Cloud & infrastructure", "Data", "AI", "Developer tools",
              "Cybersecurity", "Telecom", "Hardware", "Technology services"]
DEPARTMENTS = ["Engineering", "Product", "Data", "IT", "Security", "Finance", "Corporate"]

# id, supplier, category, product, dept, model, billing, start, end, auto_renew, notice_days,
# units, unit_price, benchmark, active_units, commit, fixed_acv, proposed_uplift, infosec
D = date
CONTRACTS = [
    ("C01", "Collabry", "SaaS", "Collaboration suite", "IT", "Per seat", "Monthly", D(2024, 1, 1), D(2026, 12, 31), "Y", 60, 1500, 264, 240, 1395, None, None, 0.07, "Current"),
    ("C02", "TaskFlow", "SaaS", "Work management", "Product", "Per seat", "Monthly", D(2025, 11, 1), D(2026, 10, 31), "Y", 30, 900, 300, 252, 540, None, None, 0.08, "Current"),
    ("C03", "Docuvault", "SaaS", "E-signature", "Corporate", "Per seat", "Annual", D(2025, 3, 1), D(2027, 2, 28), "Y", 60, 120, 480, 420, 74, None, None, 0.05, "Due"),
    ("C04", "PeopleGrid", "SaaS", "HR information system", "Corporate", "Per seat", "Annual", D(2024, 7, 1), D(2027, 6, 30), "N", 90, 1500, 180, 186, 1420, None, None, 0.04, "Current"),
    ("C05", "Chatterly", "SaaS", "Team messaging", "IT", "Per seat", "Monthly", D(2025, 4, 1), D(2027, 3, 31), "Y", 30, 1600, 105, 96, 1510, None, None, 0.06, "Current"),
    ("C06", "Pulse Analytics", "Data", "Product analytics", "Product", "Per seat", "Monthly", D(2025, 12, 1), D(2026, 11, 30), "Y", 45, 250, 1440, 1200, 162, None, None, 0.09, "Current"),
    ("C07", "PairCode AI", "AI", "AI coding assistant", "Engineering", "Per seat", "Monthly", D(2026, 1, 1), D(2026, 12, 31), "Y", 60, 700, 228, 228, 610, None, None, 0.00, "Due"),
    ("C08", "GitForge", "Developer tools", "Source control", "Engineering", "Per seat", "Annual", D(2025, 2, 1), D(2027, 1, 31), "Y", 60, 750, 252, 231, 702, None, None, 0.05, "Current"),
    ("C09", "TicketHub", "Developer tools", "Issue tracking", "Engineering", "Per seat", "Annual", D(2026, 4, 1), D(2027, 3, 31), "Y", 60, 1100, 186, 174, 820, None, None, 0.06, "Current"),
    ("C10", "Shieldpoint", "Cybersecurity", "Endpoint protection", "Security", "Per seat", "Annual", D(2025, 6, 1), D(2027, 5, 31), "N", 90, 1800, 66, 60, 1710, None, None, 0.04, "Current"),
    ("C11", "IdentiKey", "Cybersecurity", "SSO and MFA", "Security", "Per seat", "Annual", D(2025, 1, 1), D(2026, 12, 31), "Y", 90, 1600, 72, 66, 1480, None, None, 0.07, "Current"),
    ("C12", "Voxline", "Telecom", "Unified communications", "IT", "Per seat", "Monthly", D(2024, 5, 1), D(2027, 4, 30), "Y", 60, 600, 216, 168, 290, None, None, 0.05, "Expired"),
    ("C13", "LedgerLoop", "SaaS", "Spend and AP automation", "Finance", "Per seat", "Annual", D(2026, 2, 1), D(2028, 1, 31), "N", 90, 80, 1800, 1740, 71, None, None, 0.03, "Current"),
    ("C14", "Stratus Cloud", "Cloud & infrastructure", "IaaS compute and storage", "Engineering", "Commit", "Monthly", D(2025, 1, 1), D(2027, 12, 31), "N", 90, None, None, None, None, 9600, None, 0.00, "Current"),
    ("C15", "Nimbus Edge", "Cloud & infrastructure", "CDN and edge delivery", "Engineering", "Commit", "Monthly", D(2026, 1, 1), D(2026, 12, 31), "Y", 60, None, None, None, None, 1800, None, 0.04, "Current"),
    ("C16", "Quarry Data", "Data", "Cloud data warehouse", "Data", "Commit", "Monthly", D(2025, 7, 1), D(2027, 6, 30), "N", 90, None, None, None, None, 2400, None, 0.00, "Current"),
    ("C17", "Cortex AI", "AI", "LLM API platform", "Product", "Commit", "Monthly", D(2026, 1, 1), D(2026, 12, 31), "Y", 30, None, None, None, None, 1200, None, 0.00, "Due"),
    ("C18", "BuildBay", "Developer tools", "CI/CD build minutes", "Engineering", "Commit", "Monthly", D(2026, 1, 1), D(2026, 12, 31), "Y", 30, None, None, None, None, 600, None, 0.05, "Current"),
    ("C19", "Sentinel Observability", "Developer tools", "Observability and logging", "Engineering", "Commit", "Monthly", D(2025, 3, 1), D(2027, 2, 28), "Y", 90, None, None, None, None, 2100, None, 0.12, "Current"),
    ("C20", "OddsFeed", "Data", "Sports data feed", "Product", "Fixed", "Monthly", D(2024, 1, 1), D(2027, 12, 31), "N", 180, None, None, None, None, None, 3200, 0.05, "Current"),
    ("C21", "Metrolink", "Telecom", "WAN and internet circuits", "IT", "Fixed", "Monthly", D(2023, 7, 1), D(2026, 12, 31), "Y", 90, None, None, None, None, None, 840, 0.03, "Current"),
    ("C22", "Apex Devices", "Hardware", "Laptop leasing", "IT", "Fixed", "Monthly", D(2025, 1, 1), D(2027, 12, 31), "N", 90, None, None, None, None, None, 1560, 0.00, "Current"),
    ("C23", "Rackline", "Hardware", "Network hardware maintenance", "IT", "Fixed", "Annual", D(2026, 3, 1), D(2027, 2, 28), "Y", 60, None, None, None, None, None, 420, 0.06, "Current"),
    ("C24", "VaultGuard", "Cybersecurity", "Secrets management", "Security", "Fixed", "Annual", D(2026, 5, 1), D(2027, 4, 30), "Y", 60, None, None, None, None, None, 360, 0.05, "Current"),
    ("C25", "Bastion", "Cybersecurity", "WAF and DDoS protection", "Security", "Fixed", "Monthly", D(2025, 9, 1), D(2027, 8, 31), "N", 90, None, None, None, None, None, 720, 0.04, "Current"),
    ("C26", "Keystone Consulting", "Technology services", "Platform engineering SOW", "Engineering", "Fixed", "Monthly", D(2026, 1, 1), D(2026, 12, 31), "N", 30, None, None, None, None, None, 2400, 0.00, "Current"),
    ("C27", "Northgate Talent", "Technology services", "Contract developers", "Engineering", "Fixed", "Monthly", D(2025, 7, 1), D(2027, 6, 30), "N", 60, None, None, None, None, None, 3600, 0.03, "Due"),
    ("C28", "Prism QA Labs", "Technology services", "Test automation services", "Product", "Fixed", "Monthly", D(2026, 4, 1), D(2027, 3, 31), "Y", 60, None, None, None, None, None, 900, 0.05, "Current"),
]
CONTRACT_COLS = ["id", "supplier", "category", "product", "dept", "model", "billing", "start", "end",
                 "auto", "notice", "units", "price", "bench", "active", "commit", "fixed_acv",
                 "uplift", "infosec"]

# Monthly consumption profile for commit contracts, Jan-Sep ($000s, before noise)
CONSUMPTION = {
    "C14": [830, 845, 860, 870, 880, 890, 900, 915, 930],   # growing past its commit
    "C15": [122, 118, 116, 120, 117, 115, 119, 116, 118],   # flat, well under commit
    "C16": [215, 220, 228, 232, 236, 240, 244, 248, 252],   # growing past its commit
    "C17": [40, 50, 60, 75, 85, 100, 115, 130, 145],        # AI usage ramping up
    "C18": [39, 38, 37, 38, 39, 37, 38, 38, 37],            # flat, under commit
    "C19": [180, 184, 187, 190, 192, 195, 198, 200, 203],   # slightly over commit
}

# Off-contract (tail) suppliers: name, category, dept, typical monthly spend ($000s)
TAIL = [
    ("Snapframe", "SaaS", "Product", 14),
    ("Cloudnote AI", "AI", "Corporate", 9),
    ("QuickPoll Pro", "SaaS", "Product", 6),
    ("Ridgeway IT Services", "Technology services", "IT", 38),
    ("Lumenstack Monitoring", "Developer tools", "Engineering", 17),
]

# RFx for the Sentinel Observability renewal (C19): 3-year subscription ($000s per year),
# one-time implementation / migration cost, and 1-5 evaluation scores
RFX_CONTRACT = "C19"
RFX_CRITERIA = [("Functionality", 0.30), ("Security & compliance", 0.20),
                ("Integration effort", 0.15), ("Support & SLA", 0.10), ("Price (3-yr TCO)", 0.25)]
RFX_BIDS = [
    # bidder, y1, y2, y3, one-time, functionality, security, integration, support
    ("Sentinel (renewal quote)", 2352.0, 2469.6, 2593.1, 0.0, 4.5, 4.5, 5.0, 4.0),
    ("Sentinel (best and final)", 2016.0, 2076.5, 2138.8, 0.0, 4.5, 4.5, 5.0, 4.0),
    ("Lumen Trace", 1850.0, 1905.5, 1962.7, 260.0, 4.0, 4.5, 3.5, 4.5),
    ("Beacon Logs", 1620.0, 1700.0, 1785.0, 340.0, 3.5, 3.0, 3.0, 3.5),
]

SEED = 2026


def acv(c):
    """Annual contract value ($000s): seats x unit price, the commit, or the fixed fee."""
    if c["model"] == "Per seat":
        return c["units"] * c["price"] / 1000
    if c["model"] == "Commit":
        return float(c["commit"])
    return float(c["fixed_acv"])


def contracts():
    return [dict(zip(CONTRACT_COLS, row)) for row in CONTRACTS]


def spend_rows():
    """FY2026 AP spend by supplier and month ($000s). Open months (Oct-Dec) are None.

    Returns a list of dicts: supplier, contract_id ("None" if off-contract), category, dept, months[12].
    """
    rng = np.random.default_rng(SEED)
    rows = []
    for c in contracts():
        m = [None] * 12
        a = acv(c)
        for i in range(MONTHS_CLOSED):
            month_start = date(FISCAL_YEAR, i + 1, 1)
            if c["model"] == "Commit":
                v = CONSUMPTION[c["id"]][i] * (1 + rng.normal(0, 0.02))
            elif month_start < date(c["start"].year, c["start"].month, 1) and c["start"].year == FISCAL_YEAR:
                v = 0.0  # supplier not engaged yet
            elif c["billing"] == "Annual":
                v = a if (i + 1) == c["start"].month else 0.0
            else:
                noise = 0.08 if c["category"] == "Technology services" else 0.015
                v = a / 12 * (1 + rng.normal(0, noise))
            m[i] = round(v, 1)
        rows.append(dict(supplier=c["supplier"], contract_id=c["id"], category=c["category"],
                         dept=c["dept"], months=m))
    for name, cat, dept, typical in TAIL:
        m = [None] * 12
        for i in range(MONTHS_CLOSED):
            m[i] = round(max(0.0, typical * (1 + rng.normal(0, 0.35))), 1)
        rows.append(dict(supplier=name, contract_id="None", category=cat, dept=dept, months=m))
    return rows
