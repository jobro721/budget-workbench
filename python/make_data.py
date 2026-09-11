#!/usr/bin/env python3
"""Generate ALL synthetic sample data for the budget workbench (seeded, reproducible).

One seed -> every artifact in the repo is regenerable:
  ../data/synthetic/ledgers.csv         two "fund ledger" files (A and B) for reconciliation
  ../data/synthetic/budget.csv          budget (line x object class)
  ../data/synthetic/actuals.csv         actuals (same keys)
  ../data/synthetic/filed_sample.csv    File D-shaped sample (account x object class x period)

The agency is fictional ("Federal Health Administration"). Magnitudes are
realistic; every value is generated — no real data, no real TAS codes.

Usage:  python3 make_data.py            (writes ../data/synthetic/)
"""
import csv
import random
from pathlib import Path

SEED = 20260911
OUT = Path(__file__).resolve().parent.parent / "data" / "synthetic"

OBJECT_CLASSES = [
    ("11.1", "Compensation of employees"),
    ("11.2", "Other personnel expenses"),
    ("12.1", "Travel and transportation of persons"),
    ("13.1", "Equipment and other capital assets"),
    ("25.0", "Research and development"),
    ("31.0", "Rental of buildings, machinery, and equipment"),
    ("33.0", "Grants, subsidies, and contributions"),
    ("95.0", "Other miscellaneous expenses"),
]
LINES = [
    ("L-100", "Clinical Services"),
    ("L-200", "Outpatient Care"),
    ("L-300", "Inpatient Care"),
    ("L-400", "Pharmacy Operations"),
    ("L-500", "Patient Support"),
    ("L-600", "Administrative & General"),
]
FUND_POINTS = [f"FCP-{i:03d}" for i in range(1, 13)]  # 12 synthetic fund control points


def money(rng: random.Random, lo: float, hi: float) -> float:
    return round(rng.uniform(lo, hi) / 1000, 0)  # whole thousands — obviously synthetic


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    rng = random.Random(SEED)

    # --- 1. two ledgers for reconciliation (A = "status of allowances" style, B = "fund balance" style)
    a_rows, b_rows = [], []
    for fp in FUND_POINTS:
        for _ in range(rng.randint(2, 4)):
            oc = rng.choice(OBJECT_CLASSES)
            base = money(rng, 50_000, 900_000)
            # 85% match exactly, 10% within tolerance (< 5k), 5% real exceptions (> 5k or missing)
            kind = rng.random()
            if kind < 0.85:
                amt_b = base
            elif kind < 0.95:
                amt_b = base + round(rng.uniform(100, 4_900), 0)
            else:
                amt_b = base + round(rng.choice([-1, 1]) * rng.uniform(6_000, 400_000), 0)
            a_rows.append({"fund_control_point": fp, "object_class": oc[0], "description": oc[1],
                           "period": f"FY2025-P{rng.randint(1,12):02d}", "amount": base})
            if rng.random() > 0.02:  # 2% missing from B
                b_rows.append({"fund_control_point": fp, "object_class": oc[0], "description": oc[1],
                               "period": "", "balance": amt_b})
    with (OUT / "ledgers.csv").open("w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["file", "fund_control_point", "object_class", "description", "period", "amount"])
        for r in a_rows:
            w.writerow(["A", r["fund_control_point"], r["object_class"], r["description"], r["period"], r["amount"]])
        for r in b_rows:
            w.writerow(["B", r["fund_control_point"], r["object_class"], r["description"], r["period"], r["balance"]])
    print(f"ledgers.csv: {len(a_rows)} rows in A, {len(b_rows)} rows in B")

    # --- 2. budget vs actuals (line x object class)
    budget, actuals = [], []
    for line in LINES:
        for oc in OBJECT_CLASSES:
            b = money(rng, 200_000, 5_000_000)
            # actuals: 70% on-plan, 20% under, 10% over (variance flags)
            drift = rng.choice([1.0, 1.0, 1.0, 1.0, 1.0, 0.85, 0.7, 1.15, 1.3])
            a = round(b * drift, 0)
            budget.append({"line": line[0], "line_name": line[1], "object_class": oc[0],
                           "oc_name": oc[1], "budget": b})
            actuals.append({"line": line[0], "line_name": line[1], "object_class": oc[0],
                            "oc_name": oc[1], "actual": a})
    for name, rows in (("budget.csv", budget), ("actuals.csv", actuals)):
        with (OUT / name).open("w", newline="") as f:
            w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
            w.writeheader()
            w.writerows(rows)
    print(f"budget.csv / actuals.csv: {len(budget)} rows each")

    # --- 3. File D-shaped sample (account x object class x period, obligated/outlay/UOB)
    rows = []
    for period in range(1, 13):
        for acc in range(1, 5):
            for oc in rng.sample(OBJECT_CLASSES, 4):
                ob = money(rng, 50_000, 2_000_000)
                ol = round(ob * rng.uniform(0.6, 0.95), 0)
                rows.append({"federal_account_code": f"099-{acc:04d}", "object_class_code": oc[0],
                             "fiscal_period": period, "obligated_amt": ob, "outlay_amt": ol,
                             "unliquidated_balance_amt": round(ob - ol, 0)})
    with (OUT / "filed_sample.csv").open("w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)
    print(f"filed_sample.csv: {len(rows)} rows")
    print(f"OK -> {OUT}")


if __name__ == "__main__":
    main()
