#!/usr/bin/env python3
"""Build the four budget-workbench template workbooks from the synthetic data.

  workbooks/execution_tracker.xlsx       SF-133-style execution grid + by-period
  workbooks/budget_formulation.xlsx      driver-based formulation + assumptions + cross-checks
  workbooks/variance_scorecard.xlsx      KPI scorecard + chart + flagged variance detail
  workbooks/reconciliation_workbook.xlsx the two ledger blocks (A/B) + Exceptions/Summary tabs

Sheet layouts are CONTRACTS with the VBA macros in ../vba/ — do not rename
columns or move markers without updating both sides:
  Execution: A=Account B=ObjectClass C=BudgetAuthority D=Obligated E=Outlay F=UOB G=Pacing%
             (MonthlyCloseChecklist checks pace in D/C and the UOB identity F=D-E)
  Data:      a marker cell in A holding exactly "A" or "B", header on the same row,
             data below with FCP in col B, ObjectClass in col C, Amount in col F
             (RunReconciliation.ReadLedger scans for the markers)
  Data:      marker "BUDGET"/"ACTUALS" in col A, Line in B? no — Line in col A of the
             DATA rows (col 1), ObjectClass in col 3, Amount in col 5
             (BuildVarianceReport.ReadBudget reads data rows 1/3/5)

All figures are synthetic (python/make_data.py, seed 20260911).
"""
import csv
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

ROOT = Path(__file__).resolve().parent.parent
SYN = ROOT / "data" / "synthetic"
WB = ROOT / "workbooks"

GREEN = "1E4B38"
GOLD = "A68A5B"
INK = "2B241D"
AMBER_FILL = PatternFill("solid", fgColor="F6EEDD")
RED_FILL = PatternFill("solid", fgColor="F3E1E4")
HEAD_FONT = Font(name="Georgia", size=11, bold=True, color=GREEN)
TITLE_FONT = Font(name="Georgia", size=14, bold=True, color=GREEN)
THIN = Border(bottom=Side(style="thin", color=GOLD))
DOUBLE = Border(bottom=Side(style="double", color=GOLD))


def style_header(ws, ncols, row=1):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEAD_FONT
        cell.border = DOUBLE
    ws.freeze_panes = ws.cell(row=row + 1, column=1)


def title(ws, text, ncols=6):
    ws.cell(row=1, column=1, value=text).font = TITLE_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)


def autosize(ws):
    for col in ws.columns:
        try:
            w = max((len(str(c.value)) for c in col if c.value is not None), default=8)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(w + 2, 10), 42)
        except TypeError:
            pass


def read_csv(name):
    with (SYN / name).open() as f:
        return list(csv.DictReader(f))


def f(x):
    return float(x)


# ---------------------------------------------------------------- execution
def build_execution():
    rows = read_csv("filed_sample.csv")
    # account x object class grid
    grid = defaultdict(lambda: [0.0, 0.0, 0.0])  # [obligated, outlay, ba]
    byperiod = defaultdict(lambda: [0.0] * 13)   # per account: obligated by period
    accounts = sorted({r["federal_account_code"] for r in rows})
    for r in rows:
        a, oc, p = r["federal_account_code"], r["object_class_code"], int(r["fiscal_period"])
        ob, ol = f(r["obligated_amt"]), f(r["outlay_amt"])
        grid[(a, oc)][0] += ob
        grid[(a, oc)][1] += ol
        byperiod[a][p] += ob
    # synthetic BA: obligated / 0.82 (a plausible obligation pacing) — flagged synthetic
    for k in grid:
        grid[k][2] = round(grid[k][0] / 0.82, 0)

    wb = Workbook()
    ws = wb.active
    ws.title = "Execution"
    title(ws, "Execution Tracker — synthetic agency (seed 20260911)", ncols=7)
    ws.cell(row=2, column=1, value="Budget authority, obligations and outlays by account x object class. "
                                   "UOB and Pacing are live formulas.")
    hdr = ["Account", "ObjectClass", "BudgetAuthority", "Obligated", "Outlay", "UOB", "Pacing%"]
    for c, h in enumerate(hdr, start=1):
        ws.cell(row=3, column=c, value=h)
    style_header(ws, len(hdr), row=3)
    r = 4
    for a in accounts:
        ocs = sorted({k[1] for k in grid if k[0] == a})
        for oc in ocs:
            ba, ob, ol = grid[(a, oc)]
            ws.cell(row=r, column=1, value=a)
            ws.cell(row=r, column=2, value=oc)
            ws.cell(row=r, column=3, value=ba)
            ws.cell(row=r, column=4, value=ob)
            ws.cell(row=r, column=5, value=ol)
            ws.cell(row=r, column=6, value=f"=D{r}-E{r}")
            ws.cell(row=r, column=7, value=f'=IF(C{r}>0,D{r}/C{r},"")')
            for c in (3, 4, 5, 6):
                ws.cell(row=r, column=c).number_format = "#,##0"
            ws.cell(row=r, column=7).number_format = "0.0"
            r += 1
    last = r - 1
    ws.conditional_formatting.add(
        f"G4:G{last}",
        CellIsRule(operator="lessThan", formula=["0.5"], fill=AMBER_FILL),
    )
    ws.conditional_formatting.add(
        f"G4:G{last}",
        CellIsRule(operator="greaterThan", formula=["1.0"], fill=RED_FILL),
    )
    ws["I1"] = "Pacing target"
    ws["J1"] = 0.85
    ws["I2"] = "variance band %"
    ws["J2"] = 10
    wb.defined_names["PACING_TARGET"] = DefinedName("PACING_TARGET", attr_text="Execution!$J$1")
    wb.defined_names["VAR_BAND_PCT"] = DefinedName("VAR_BAND_PCT", attr_text="Execution!$J$2")

    ws2 = wb.create_sheet("ByPeriod")
    title(ws2, "Obligated by fiscal period — SF-133-style pacing view", ncols=14)
    ws2.cell(row=3, column=1, value="Account")
    for p in range(1, 13):
        ws2.cell(row=3, column=1 + p, value=f"P{p:02d}")
    style_header(ws2, 13, row=3)
    r = 4
    for a in accounts:
        ws2.cell(row=r, column=1, value=a)
        for p in range(1, 13):
            cell = ws2.cell(row=r, column=1 + p, value=byperiod[a][p])
            cell.number_format = "#,##0"
        r += 1
    autosize(ws)
    autosize(ws2)
    out = WB / "execution_tracker.xlsx"
    wb.save(out)
    print(f"OK {out.name}: {last - 3} grid rows, {len(accounts)} accounts x 12 periods")


# ---------------------------------------------------------------- formulation
def build_formulation():
    lines = [
        ("L-100", "Clinical Services", "FTE count", 120),
        ("L-200", "Outpatient Care", "visit count", 48000),
        ("L-300", "Inpatient Care", "discharge count", 2100),
        ("L-400", "Pharmacy Operations", "script count", 65000),
        ("L-500", "Patient Support", "enrollment", 15000),
        ("L-600", "Administrative & General", "FTE count", 85),
    ]
    costs = {0: 96000, 1: 42, 2: 18500, 3: 18, 4: 340, 5: 88000}
    growth = {0: 0.03, 1: 0.05, 2: 0.04, 3: 0.06, 4: 0.045, 5: 0.025}

    wb = Workbook()
    ws = wb.active
    ws.title = "Formulation"
    title(ws, "Budget Formulation (driver-based) — synthetic agency", ncols=9)
    ws.cell(row=2, column=1, value="FY+1 is UnitCount x UnitCost; FY+2..FY+5 apply the growth rate. "
                                   "Every driver change is logged on the Assumptions tab.")
    hdr = ["Line", "LineName", "Driver", "UnitCount", "UnitCost", "Growth%", "FY+1", "FY+2", "FY+3"]
    for c, h in enumerate(hdr, start=1):
        ws.cell(row=3, column=c, value=h)
    style_header(ws, len(hdr), row=3)
    r = 4
    for i, (ln, name, drv, units) in enumerate(lines):
        ws.cell(row=r, column=1, value=ln)
        ws.cell(row=r, column=2, value=name)
        ws.cell(row=r, column=3, value=drv)
        ws.cell(row=r, column=4, value=units)
        ws.cell(row=r, column=5, value=costs[i])
        ws.cell(row=r, column=6, value=growth[i])
        ws.cell(row=r, column=7, value=f"=D{r}*E{r}")
        ws.cell(row=r, column=8, value=f"=G{r}*(1+F{r})")
        ws.cell(row=r, column=9, value=f"=H{r}*(1+F{r})")
        ws.cell(row=r, column=4).number_format = "#,##0"
        ws.cell(row=r, column=5).number_format = "#,##0"
        ws.cell(row=r, column=6).number_format = "0.0%"
        for c in (7, 8, 9):
            ws.cell(row=r, column=c).number_format = "#,##0"
        r += 1
    last = r - 1
    ws.cell(row=r + 1, column=2, value="TOTAL").font = HEAD_FONT
    for c in (4, 7, 8, 9):
        cell = ws.cell(row=r + 1, column=c, value=f"=SUM({get_column_letter(c)}4:{get_column_letter(c)}{last})")
        cell.font = HEAD_FONT
        cell.number_format = "#,##0"
        cell.border = DOUBLE
    autosize(ws)

    a = wb.create_sheet("Assumptions")
    title(a, "Assumptions Log — every driver change, dated and sourced", ncols=7)
    hdr = ["Date", "Change", "Driver", "Old Value", "New Value", "Source", "Author"]
    for c, h in enumerate(hdr, start=1):
        a.cell(row=3, column=c, value=h)
    style_header(a, len(hdr), row=3)
    seed_rows = [
        ("2026-09-01", "Initial load (synthetic seed 20260911)", "all", "", "", "python/make_data.py", "template"),
        ("2026-09-11", "Outpatient visit growth set to plan", "visit count", "", "5.0%", "synthetic planning memo", "template"),
    ]
    r = 4
    for row in seed_rows:
        for c, v in enumerate(row, start=1):
            a.cell(row=r, column=c, value=v)
        r += 1
    autosize(a)

    x = wb.create_sheet("CrossChecks")
    title(x, "Cross-Checks — bottom-up vs top-down", ncols=4)
    x.cell(row=3, column=1, value="Check")
    x.cell(row=3, column=2, value="Value")
    x.cell(row=3, column=3, value="Limit")
    x.cell(row=3, column=4, value="Result")
    style_header(x, 4, row=3)
    x.cell(row=4, column=1, value="Bottom-up FY+1 total")
    x.cell(row=4, column=2, value="=Formulation!G" + str(last + 2)).number_format = "#,##0"
    x.cell(row=5, column=1, value="Top-down allocation (named range)")
    x.cell(row=5, column=2, value="=TOPDOWN_ALLOC").number_format = "#,##0"
    x.cell(row=6, column=1, value="Difference")
    x.cell(row=6, column=2, value="=B4-B5").number_format = "#,##0"
    x.cell(row=7, column=1, value="Tolerance %")
    x.cell(row=7, column=2, value=0.02)
    x.cell(row=7, column=2).number_format = "0.0%"
    x.cell(row=8, column=1, value="Result")
    x.cell(row=8, column=4, value='=IF(ABS(B6/B4)<B7,"PASS","REVIEW")')
    wb.defined_names["TOPDOWN_ALLOC"] = DefinedName("TOPDOWN_ALLOC", attr_text="CrossChecks!$B$5")
    # seed a top-down allocation within 2% of the bottom-up total
    x["B5"] = 100000000  # placeholder; replaced below with computed value
    autosize(x)

    # set the top-down allocation to 1.01 x bottom-up so the check passes (synthetic)
    bottom = sum(u * c for (_, _, _, u), c in zip(lines, costs.values()))
    x["B5"] = round(bottom * 1.01, 0)

    out = WB / "budget_formulation.xlsx"
    wb.save(out)
    print(f"OK {out.name}: {len(lines)} driver lines, cross-check seeded at 1.01x bottom-up")


# ---------------------------------------------------------------- scorecard
def build_scorecard():
    budget = read_csv("budget.csv")
    actuals = read_csv("actuals.csv")
    amap = {(r["line"], r["object_class"]): f(r["actual"]) for r in actuals}

    wb = Workbook()
    ws = wb.active
    ws.title = "Scorecard"
    title(ws, "Variance Scorecard — synthetic FY (seed 20260911)", ncols=4)
    tb = sum(f(r["budget"]) for r in budget)
    ta = sum(f(r["actual"]) for r in actuals)
    kpis = [
        ("Total budget", tb, "#,##0"),
        ("Total actual", ta, "#,##0"),
        ("Variance $", ta - tb, "#,##0"),
        ("Variance %", (ta - tb) / tb * 100, "0.0"),
        ("Pacing target (named range)", "=PACING_TARGET", "0.0%"),
    ]
    r = 3
    for name, val, fmt in kpis:
        ws.cell(row=r, column=1, value=name)
        cell = ws.cell(row=r, column=2, value=val)
        cell.number_format = fmt
        r += 1
    ws["A10"] = "Thresholds"
    ws["A11"] = "Pacing target (edit me)"
    ws["B11"] = 0.85
    ws["A12"] = "amber variance band >%"
    ws["B12"] = 10
    ws["A13"] = "red variance band >%"
    ws["B13"] = 25
    wb.defined_names["PACING_TARGET"] = DefinedName("PACING_TARGET", attr_text="Scorecard!$B$11")
    wb.defined_names["VAR_BAND_PCT"] = DefinedName("VAR_BAND_PCT", attr_text="Scorecard!$B$12")

    # detail rows for the chart: top lines by |variance|
    byline = defaultdict(lambda: [0.0, 0.0])
    for rb in budget:
        k = (rb["line"], rb["object_class"])
        byline[rb["line"]][0] += f(rb["budget"])
        byline[rb["line"]][1] += amap[k]
    top = sorted(byline.items(), key=lambda kv: abs(kv[1][1] - kv[1][0]), reverse=True)[:5]
    ws["D3"] = "Line"
    ws["E3"] = "Budget"
    ws["F3"] = "Actual"
    for c in "DEF":
        ws[f"{c}3"].font = HEAD_FONT
        ws[f"{c}3"].border = DOUBLE
    r = 4
    for ln, (b, a) in top:
        ws.cell(row=r, column=4, value=ln)
        ws.cell(row=r, column=5, value=b).number_format = "#,##0"
        ws.cell(row=r, column=6, value=a).number_format = "#,##0"
        r += 1
    chart = BarChart()
    chart.type = "col"
    chart.title = "Budget vs actual — top 5 lines (synthetic)"
    chart.height = 8
    chart.width = 16
    data = Reference(ws, min_col=5, min_row=3, max_col=6, max_row=r - 1)
    cats = Reference(ws, min_col=4, min_row=4, max_row=r - 1)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, "A14")

    v = wb.create_sheet("Variance")
    title(v, "Variance detail (flagged rows mirror python/variance.py)", ncols=8)
    hdr = ["Line", "LineName", "ObjectClass", "Budget", "Actual", "Var$", "Var%", "Flag"]
    for c, h in enumerate(hdr, start=1):
        v.cell(row=3, column=c, value=h)
    style_header(v, len(hdr), row=3)
    rows = []
    for rb in budget:
        k = (rb["line"], rb["object_class"])
        b, a = f(rb["budget"]), amap[k]
        vp = (a - b) / b * 100
        flag = "RED" if abs(vp) > 25 else ("AMBER" if abs(vp) > 10 else "OK")
        rows.append((rb["line"], rb["line_name"], rb["object_class"], b, a, a - b, vp, flag))
    rows.sort(key=lambda x: (x[7] != "OK", -abs(x[6])))
    r = 4
    for ln, nm, oc, b, a, dv, vp, fl in rows:
        v.cell(row=r, column=1, value=ln)
        v.cell(row=r, column=2, value=nm)
        v.cell(row=r, column=3, value=oc)
        v.cell(row=r, column=4, value=b).number_format = "#,##0"
        v.cell(row=r, column=5, value=a).number_format = "#,##0"
        v.cell(row=r, column=6, value=dv).number_format = "#,##0"
        v.cell(row=r, column=7, value=round(vp, 1)).number_format = "0.0"
        v.cell(row=r, column=8, value=fl)
        if fl == "AMBER":
            for c in range(1, 9):
                v.cell(row=r, column=c).fill = AMBER_FILL
        elif fl == "RED":
            for c in range(1, 9):
                v.cell(row=r, column=c).fill = RED_FILL
        r += 1
    autosize(ws)
    autosize(v)
    out = WB / "variance_scorecard.xlsx"
    wb.save(out)
    print(f"OK {out.name}: 5 KPIs, top-5 chart, {len(rows)} detail rows")


# ---------------------------------------------------------------- reconciliation
def build_reconciliation():
    rows = read_csv("ledgers.csv")
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.cell(row=1, column=1, value="2026-09-11")  # data date — MonthlyCloseChecklist reads A1
    ws["A1"].number_format = "yyyy-mm-dd"
    ws.cell(row=2, column=1, value="Synthetic ledgers (seed 20260911). Run RunReconciliation to build "
                                   "Exceptions/Summary, or python/recon.py for the twin output.")

    a_rows = [r for r in rows if r["file"] == "A"]
    b_rows = [r for r in rows if r["file"] == "B"]
    r = 4
    # block A: marker row (col A = "A" + header) then data
    ws.cell(row=r, column=1, value="A")
    for c, h in enumerate(["", "FundControlPoint", "ObjectClass", "Description", "Period", "Amount"], start=1):
        ws.cell(row=r, column=c, value=h)
    for c in range(1, 7):
        ws.cell(row=r, column=c).font = HEAD_FONT
    r += 1
    for x in a_rows:
        ws.cell(row=r, column=1, value="A")
        ws.cell(row=r, column=2, value=x["fund_control_point"])
        ws.cell(row=r, column=3, value=x["object_class"])
        ws.cell(row=r, column=4, value=x["description"])
        ws.cell(row=r, column=5, value=x["period"])
        ws.cell(row=r, column=6, value=f(x["amount"])).number_format = "#,##0"
        r += 1
    r += 1
    # block B
    ws.cell(row=r, column=1, value="B")
    for c, h in enumerate(["", "FundControlPoint", "ObjectClass", "Description", "Period", "Amount"], start=1):
        ws.cell(row=r, column=c, value=h)
    for c in range(1, 7):
        ws.cell(row=r, column=c).font = HEAD_FONT
    r += 1
    for x in b_rows:
        ws.cell(row=r, column=1, value="B")
        ws.cell(row=r, column=2, value=x["fund_control_point"])
        ws.cell(row=r, column=3, value=x["object_class"])
        ws.cell(row=r, column=4, value=x["description"])
        ws.cell(row=r, column=5, value=x["period"])
        ws.cell(row=r, column=6, value=f(x["amount"])).number_format = "#,##0"
        r += 1
    autosize(ws)

    e = wb.create_sheet("Exceptions")
    hdr = ["FundControlPoint", "ObjectClass", "Description", "LedgerA", "LedgerB", "Diff (A-B)", "Status", "Band"]
    for c, h in enumerate(hdr, start=1):
        e.cell(row=1, column=c, value=h)
    style_header(e, len(hdr))
    s = wb.create_sheet("Summary")
    s.cell(row=1, column=1, value="Reconciliation Summary — synthetic ledgers").font = TITLE_FONT
    s.cell(row=3, column=1, value="Run RunReconciliation (VBA) or python/recon.py to populate this tab.")

    out = WB / "reconciliation_workbook.xlsx"
    wb.save(out)
    print(f"OK {out.name}: {len(a_rows)} rows in A, {len(b_rows)} rows in B")


if __name__ == "__main__":
    WB.mkdir(exist_ok=True)
    build_execution()
    build_formulation()
    build_scorecard()
    build_reconciliation()
    print("done.")
