#!/usr/bin/env python3
"""Variance twin of the VBA macro BuildVarianceReport.

Reads ../data/synthetic/{budget,actuals}.csv and produces
../workbooks/output/variance_report.xlsx:
  - Variance tab: line x object class with budget, actual, $ variance, % variance,
    threshold flags (|var %| > 10% amber, > 25% red)
  - Summary tab: totals + top-5 adverse variances
  - Waterfall: embedded matplotlib chart (budget -> actual, by top drivers)

The math mirrors vba/BuildVarianceReport.bas one-for-one.

Usage:  python3 variance.py
"""
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "workbooks" / "output" / "variance_report.xlsx"
AMBER_PCT, RED_PCT = 10.0, 25.0
GOLD = "A68A5B"
AMBER_FILL = PatternFill("solid", fgColor="F6EEDD")
RED_FILL = PatternFill("solid", fgColor="F3E1E4")
HEAD_FONT = Font(name="Georgia", size=11, bold=True, color="1E4B38")
THIN = Border(bottom=Side(style="thin", color=GOLD))


def main() -> None:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    b = pd.read_csv(ROOT / "data" / "synthetic" / "budget.csv")
    a = pd.read_csv(ROOT / "data" / "synthetic" / "actuals.csv")
    v = b.merge(a, on=["line", "line_name", "object_class", "oc_name"], how="outer")
    v["variance"] = v["actual"] - v["budget"]
    v["var_pct"] = v["variance"] / v["budget"] * 100
    v["flag"] = v["var_pct"].abs().apply(
        lambda p: "RED" if p > RED_PCT else ("AMBER" if p > AMBER_PCT else "OK")
    )

    wb = Workbook()
    ws = wb.active
    ws.title = "Variance"
    ws.append(["Line", "Line Name", "Object Class", "OC Name", "Budget", "Actual",
               "Variance $", "Variance %", "Flag"])
    for c in range(1, 10):
        cell = ws.cell(row=1, column=c)
        cell.font = HEAD_FONT
        cell.border = THIN
    v = v.sort_values(["flag", "var_pct"], ascending=[True, False])
    for _, r in v.iterrows():
        ws.append([r["line"], r["line_name"], r["object_class"], r["oc_name"],
                   r["budget"], r["actual"], r["variance"], round(r["var_pct"], 1), r["flag"]])
        row = ws.max_row
        if r["flag"] == "AMBER":
            for c in range(1, 10):
                ws.cell(row=row, column=c).fill = AMBER_FILL
        elif r["flag"] == "RED":
            for c in range(1, 10):
                ws.cell(row=row, column=c).fill = RED_FILL
        for c in (5, 6, 7):
            ws.cell(row=row, column=c).number_format = "#,##0"
        ws.cell(row=row, column=8).number_format = "0.0"
    for i, w in enumerate([8, 22, 12, 30, 13, 13, 13, 11, 8], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    s = wb.create_sheet("Summary")
    s.append(["Variance Summary — synthetic FY (seed 20260911)"])
    s["A1"].font = Font(name="Georgia", size=14, bold=True, color="1E4B38")
    s.append([])
    s.append(["Total budget", float(v.budget.sum())])
    s.append(["Total actual", float(v.actual.sum())])
    s.append(["Total variance $", float(v.variance.sum())])
    s.append(["Total variance %", round(float(v.variance.sum() / v.budget.sum() * 100), 1)])
    for row in s.iter_rows(min_row=3, max_row=6, max_col=2):
        row[1].number_format = "#,##0"
    s.append([])
    s.append(["Top 5 adverse variances"])
    s.cell(row=s.max_row, column=1).font = HEAD_FONT
    top = v.nsmallest(5, "variance")
    for _, r in top.iterrows():
        s.append([f"{r['line']} {r['oc_name']}", round(float(r.var_pct), 1), float(r.variance)])
    s.column_dimensions["A"].width = 44
    s.column_dimensions["B"].width = 14
    s.column_dimensions["C"].width = 14

    # waterfall: start budget, top +/- drivers, end actual
    by_line = v.groupby("line_name")["variance"].sum().sort_values()
    drivers = list(by_line.head(3).index) + list(by_line.tail(3).index)  # 3 most adverse + 3 most favorable
    labels = ["Budget"] + [f"{n}" for n in drivers] + ["Actual"]
    vals = [float(v.budget.sum())] + [float(by_line[n]) for n in drivers] + [float(v.actual.sum())]
    fig, ax = plt.subplots(figsize=(9, 4.5), dpi=110)
    fig.patch.set_facecolor("#FAF7F0")
    bottoms, heights = [], []
    running = vals[0]
    for i, val in enumerate(vals):
        if i == 0 or i == len(vals) - 1:
            bottoms.append(0)
            heights.append(val)
            color = "#1E4B38" if i == 0 else "#7C3B4B"
        else:
            color = "#7C3B4B" if val < 0 else "#A68A5B"
            bottoms.append(min(running, running + val))
            heights.append(abs(val))
            running += val
        ax.bar(i, heights[i], bottom=bottoms[i], color=color, width=0.6)
    ax.set_xticks(range(len(labels)))
    ax.set_xticklabels(labels, rotation=20, ha="right", fontsize=8)
    ax.set_ylabel("$ (thousands)", fontsize=9)
    ax.set_title("Budget → Actual: top variance drivers (synthetic data)", fontsize=11, color="#2B241D", family="serif")
    ax.spines[["top", "right"]].set_visible(False)
    for s_ in ax.spines.values():
        s_.set_color("#" + GOLD)
    ax.tick_params(colors="#5F564A", labelsize=8)
    ax.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda x, _: f"{x:,.0f}"))
    fig.tight_layout()
    img = ROOT / "workbooks" / "output" / "waterfall.png"
    fig.savefig(img)
    from openpyxl.drawing.image import Image
    ws2 = wb.create_sheet("Waterfall")
    ws2["A1"] = "Budget → Actual (synthetic data)"
    ws2["A1"].font = Font(name="Georgia", size=12, bold=True, color="1E4B38")
    ws2.add_image(Image(str(img)), "A3")

    wb.save(OUT)
    print(f"OK -> {OUT} ({len(v)} rows; {int((v.flag != 'OK').sum())} flagged)")


if __name__ == "__main__":
    main()
