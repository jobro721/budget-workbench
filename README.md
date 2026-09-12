# budget-workbench — an Excel budget toolkit: 4 template workbooks + 6 VBA macros (with Python twins)

**The monthly-close workflow as software.** Four styled template workbooks
(budget formulation, execution tracking, variance scorecard, two-way
reconciliation), six VBA macros that run the analysis inside them, and
**Python twins of the two computational macros** so every number is
reproducible headlessly — `recon.py` and `variance.py` mirror
`RunReconciliation.bas` and `BuildVarianceReport.bas` one-for-one, so a
reviewer can diff the two implementations.

**Every sample figure is synthetic.** `python/make_data.py` generates all of it
with one seed (**20260911**) for a fictional agency ("Federal Health
Administration") — realistic magnitudes, no real data, no real TAS codes, no
CUI. One seed → every artifact in the repo is regenerable.

## Two ways to run it

**Headless (Python — no Excel needed):**

```sh
# requirements: pandas, openpyxl, matplotlib
python3 python/make_data.py          # regenerate all synthetic data (seeded)
python3 python/build_templates.py    # rebuild the four template workbooks
python3 python/variance.py           # → workbooks/output/variance_report.xlsx (+ waterfall.png)
python3 python/recon.py              # → workbooks/output/reconciliation_report.xlsx
```

`workbooks/output/` already contains the reference outputs of both twins
(variance report with the embedded waterfall, reconciliation report) — tracked
in the repo so you can see exactly what the macros produce.

**In Excel (VBA):** import the `.bas` files (VBA Editor → *File → Import
File*, one at a time, or copy each into a standard module), open the
workbook, **Alt+F8** → run.

## The six macros

| Macro | What it does |
|---|---|
| `RunReconciliation` | Two-way ledger reconciliation: outer-joins the A and B ledger blocks on (FundControlPoint, ObjectClass), bands the difference against the `TOLERANCE` constant, writes `Exceptions` + `Summary`. **Twin: `python/recon.py`.** |
| `BuildVarianceReport` | Budget-vs-actuals variance by line × object class, `$` and `%` variance with threshold flags: **\|var %\| > 10% amber, > 25% red**. **Twin: `python/variance.py`.** |
| `MonthlyCloseChecklist` | Runs the five named close checks against the workbook and appends a dated pass/fail entry to the `CloseLog` tab (newest first). |
| `EnforceLedgerFormat` | Applies the house format to a sheet: Georgia bold ledger-green headers on a double gold rule, tabular `#,##0` currency, auto-fit, frozen top row. |
| `ExportScorecardImage` | Exports the Scorecard chart as a PNG next to the workbook (for PPT / email / the close package). |
| `LoadFileD_CSV` | Ingests a USAspending File D (account balances) CSV into the workbook — the bridge from the public pipeline (`federal-budget-data`) into Excel. |

The five close checks (each a named constant, so the checklist is auditable):

1. `DATA_DATE` present on `Data!A1` and not older than 35 days
2. no `MATERIAL` or `MISSING` rows open on `Exceptions`
3. no unflagged \|var %\| > 25% rows on `Variance` (flag column says OK)
4. execution pacing = obligated / BA within 0–1.2 (sanity)
5. UOB identity: unliquidated balance = obligated − outlayed within $1

## The four workbooks

| Workbook | Sheets | Role |
|---|---|---|
| `budget_formulation.xlsx` | Formulation · Assumptions · CrossChecks | driver-based formulation (lines × object classes) with stated assumptions and cross-check totals |
| `execution_tracker.xlsx` | Execution · ByPeriod | SF-133-style execution grid: Account, ObjectClass, BudgetAuthority, Obligated, Outlay, UOB, Pacing% (the checklist paces D/C and checks the F = D − E identity) |
| `variance_scorecard.xlsx` | Scorecard · Variance | KPI scorecard + chart, flagged variance detail (amber/red bands) |
| `reconciliation_workbook.xlsx` | Data · Exceptions · Summary | the two ledger blocks (markers `A` / `B`) + reconciliation output tabs |

Sheet layouts are **contracts with the VBA macros** — the `Data` sheet holds
marker cells (`A`, `B`; `BUDGET`, `ACTUALS`) with headers on the same row and
data below; `RunReconciliation.ReadLedger` and `BuildVarianceReport.ReadBudget`
locate the blocks by those markers. Don't rename columns or move markers
without updating both sides (the contract is documented in the
`build_templates.py` docstring).

## The synthetic data

`python/make_data.py` (seed 20260911) writes four CSVs to `data/synthetic/`:

| File | Feeds |
|---|---|
| `ledgers.csv` | the two fund-ledger files (A and B) for reconciliation — includes deliberate matches, sub-tolerance noise, material breaks, and one-sided rows |
| `budget.csv` / `actuals.csv` | line × object class (standard federal object classes 11.1…95.0) for the variance work |
| `filed_sample.csv` | File D-shaped sample (account × object class × period) for `LoadFileD_CSV` |

Tolerance: recon banding is `TOLERANCE = $5,000` (green = match, amber =
within tolerance, red = material or missing on one side).

## Verification

- The Python twins are the executable spec: run `recon.py` / `variance.py` and
  compare against the VBA output — the banding/threshold logic is identical by
  design, and the committed `workbooks/output/*` are the twins' reference
  artifacts.
- `make_data.py` is deterministic: re-running it regenerates byte-identical
  CSVs from the seed, so every workbook and report in the repo traces back to
  one reproducible dataset.

## Layout

```
vba/                        the six macros (.bas, import into the workbook)
python/
  make_data.py              seeded synthetic data generator (the single source)
  build_templates.py        builds the four template workbooks (style + contracts)
  recon.py                  twin of RunReconciliation
  variance.py               twin of BuildVarianceReport (+ matplotlib waterfall)
workbooks/                  the four template workbooks
  output/                   reference outputs from the Python twins (tracked)
data/synthetic/             the four seeded CSVs
```

All sample data is fictional and generated locally — no agency-internal data,
no real TAS codes, no CUI.

A personal project by a federal budget analyst — not affiliated with my
employer.
