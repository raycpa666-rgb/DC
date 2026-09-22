"""Builds the Budget & Debt Payoff Planner spreadsheet (the product being sold)."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule, DataBarRule
from openpyxl.worksheet.datavalidation import DataValidation

OUT = "product/Budget-and-Debt-Payoff-Planner.xlsx"
F = "Arial"
INPUT = Font(name=F, color="0000FF")
BOLD = Font(name=F, bold=True)
BASE = Font(name=F)
TITLE = Font(name=F, bold=True, size=16, color="1F3A5F")
HEAD = Font(name=F, bold=True, color="FFFFFF")
HEAD_FILL = PatternFill("solid", fgColor="1F3A5F")
IN_FILL = PatternFill("solid", fgColor="FFF7D6")
TOT_FILL = PatternFill("solid", fgColor="E8EEF5")
thin = Side(style="thin", color="C9D3DF")
BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
MONEY = '$#,##0.00;($#,##0.00);"-"'
PCT = '0.0%;(0.0%);"-"'

wb = Workbook()


def header(ws, row, labels, col=1):
    for i, text in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=text)
        c.font, c.fill, c.border = HEAD, HEAD_FILL, BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def inp(ws, ref, value=None, fmt=None):
    c = ws[ref]
    if value is not None:
        c.value = value
    c.font, c.fill, c.border = INPUT, IN_FILL, BOX
    if fmt:
        c.number_format = fmt


def calc(ws, ref, formula, fmt=None, bold=False):
    c = ws[ref]
    c.value = formula
    c.font = BOLD if bold else BASE
    c.border = BOX
    if fmt:
        c.number_format = fmt


def widths(ws, w):
    for col, width in w.items():
        ws.column_dimensions[col].width = width


# ---------------------------------------------------------------- Start Here
ws = wb.active
ws.title = "Start Here"
ws["A1"] = "Budget & Debt Payoff Planner"
ws["A1"].font = TITLE
rows = [
    ("How to use this planner", True),
    ("1. Yellow cells with blue text are yours to edit. Everything else calculates automatically.", False),
    ("2. The first row of each table holds example values - overwrite or delete them.", False),
    ("3. Budget tab: enter monthly income and planned vs. actual spending per category.", False),
    ("4. Debt Payoff tab: list each debt, its APR and minimum payment, then try extra payments", False),
    ("   to see how many months and how much interest you save.", False),
    ("5. Savings Goals tab: set a target and monthly contribution to get a finish date.", False),
    ("", False),
    ("Colour key", True),
    ("Yellow fill / blue text = input", False),
    ("Black text = formula (do not overwrite)", False),
    ("", False),
    ("Notes", True),
    ("Debt payoff math assumes a fixed APR, interest compounded monthly, and the same payment every month.", False),
    ("This is a planning tool, not financial advice.", False),
]
for i, (text, is_head) in enumerate(rows, start=3):
    ws.cell(row=i, column=1, value=text).font = BOLD if is_head else BASE
inp(ws, "A12")
ws["A12"].value = "Yellow fill / blue text = input"
widths(ws, {"A": 100})

# ---------------------------------------------------------------- Budget
ws = wb.create_sheet("Budget")
ws["A1"] = "Monthly Budget"
ws["A1"].font = TITLE
ws["A2"] = "Month:"
ws["A2"].font = BOLD
inp(ws, "B2", "January 2027")

header(ws, 4, ["Income source", "Amount ($/month)"])
income = [("Salary (example)", 3800), ("Side income (example)", 400)]
INC_FIRST, INC_LAST = 5, 12
for r in range(INC_FIRST, INC_LAST + 1):
    inp(ws, f"A{r}")
    inp(ws, f"B{r}", fmt=MONEY)
for i, (name, amt) in enumerate(income):
    ws[f"A{INC_FIRST+i}"] = name
    ws[f"B{INC_FIRST+i}"] = amt
ws[f"A{INC_LAST+1}"] = "Total income"
ws[f"A{INC_LAST+1}"].font = BOLD
calc(ws, f"B{INC_LAST+1}", f"=SUM(B{INC_FIRST}:B{INC_LAST})", MONEY, True)
ws[f"B{INC_LAST+1}"].fill = TOT_FILL

EXP_HEAD = INC_LAST + 3
header(ws, EXP_HEAD, ["Expense category", "Budgeted ($)", "Actual ($)", "Difference ($)", "% of income"])
expenses = [
    ("Rent / mortgage", 1400, 1400), ("Utilities", 180, 205), ("Groceries", 450, 510),
    ("Transport", 220, 190), ("Insurance", 160, 160), ("Phone & internet", 110, 110),
    ("Debt payments", 600, 600), ("Eating out", 150, 240), ("Subscriptions", 45, 45),
    ("Savings", 500, 400), ("Personal / fun", 200, 260), ("Other", 100, 80),
]
EXP_FIRST = EXP_HEAD + 1
EXP_LAST = EXP_FIRST + 19
TOTAL_INC = f"$B${INC_LAST+1}"
for r in range(EXP_FIRST, EXP_LAST + 1):
    inp(ws, f"A{r}")
    inp(ws, f"B{r}", fmt=MONEY)
    inp(ws, f"C{r}", fmt=MONEY)
    calc(ws, f"D{r}", f'=IF(AND(B{r}="",C{r}=""),"",B{r}-C{r})', MONEY)
    calc(ws, f"E{r}", f'=IF(OR(C{r}="",{TOTAL_INC}=0),"",C{r}/{TOTAL_INC})', PCT)
for i, (cat, b, a) in enumerate(expenses):
    r = EXP_FIRST + i
    ws[f"A{r}"], ws[f"B{r}"], ws[f"C{r}"] = cat, b, a
T = EXP_LAST + 1
ws[f"A{T}"] = "Total expenses"
ws[f"A{T}"].font = BOLD
for col in "BCD":
    calc(ws, f"{col}{T}", f"=SUM({col}{EXP_FIRST}:{col}{EXP_LAST})", MONEY, True)
    ws[f"{col}{T}"].fill = TOT_FILL
calc(ws, f"E{T}", f"=IF({TOTAL_INC}=0,\"\",C{T}/{TOTAL_INC})", PCT, True)
ws.conditional_formatting.add(
    f"D{EXP_FIRST}:D{EXP_LAST}",
    CellIsRule(operator="lessThan", formula=["0"], font=Font(name=F, color="C00000")),
)

# Summary block
header(ws, 4, ["Summary", "Planned", "Actual"], col=7)
summary = [
    ("Income", f"={TOTAL_INC}", f"={TOTAL_INC}"),
    ("Expenses", f"=B{T}", f"=C{T}"),
    ("Left over", "=H5-H6", "=I5-I6"),
    ("Savings rate", f"=IF(H5=0,\"\",SUMIF($A${EXP_FIRST}:$A${EXP_LAST},\"Savings\",$B${EXP_FIRST}:$B${EXP_LAST})/H5)",
     f"=IF(I5=0,\"\",SUMIF($A${EXP_FIRST}:$A${EXP_LAST},\"Savings\",$C${EXP_FIRST}:$C${EXP_LAST})/I5)"),
]
for i, (label, p, a) in enumerate(summary, start=5):
    ws[f"G{i}"] = label
    ws[f"G{i}"].font = BOLD
    fmt = PCT if label == "Savings rate" else MONEY
    calc(ws, f"H{i}", p, fmt)
    calc(ws, f"I{i}", a, fmt)
ws["G9"] = 'Savings rate counts the row named exactly "Savings".'
ws["G9"].font = Font(name=F, italic=True, size=9, color="666666")
ws.conditional_formatting.add(
    "H7:I7", CellIsRule(operator="lessThan", formula=["0"], font=Font(name=F, bold=True, color="C00000"))
)
widths(ws, {"A": 26, "B": 16, "C": 14, "D": 16, "E": 13, "F": 3, "G": 16, "H": 14, "I": 14})
ws.freeze_panes = "A3"

# ---------------------------------------------------------------- Debt Payoff
ws = wb.create_sheet("Debt Payoff")
ws["A1"] = "Debt Payoff Calculator"
ws["A1"].font = TITLE
ws["A2"] = "Enter each debt. Add an extra monthly payment to see months and interest saved."
ws["A2"].font = Font(name=F, italic=True, color="666666")
header(ws, 4, ["Debt", "Balance ($)", "APR", "Minimum payment ($)", "Extra payment ($)",
               "Months (minimum only)", "Interest (minimum only)", "Months (with extra)",
               "Interest (with extra)", "Interest saved ($)", "Months saved"])
ws.row_dimensions[4].height = 32
debts = [("Credit card (example)", 4200, 0.2299, 120, 100),
         ("Car loan (example)", 9500, 0.069, 290, 0),
         ("Student loan (example)", 18000, 0.055, 200, 50)]
D1, DN = 5, 14
LOW = '"Payment too low"'
for r in range(D1, DN + 1):
    inp(ws, f"A{r}")
    inp(ws, f"B{r}", fmt=MONEY)
    inp(ws, f"C{r}", fmt="0.00%")
    inp(ws, f"D{r}", fmt=MONEY)
    inp(ws, f"E{r}", fmt=MONEY)
    # months to payoff via NPER; blank when row empty; message when payment never covers interest
    nper_min = f"NPER(C{r}/12,-D{r},B{r})"
    nper_ext = f"NPER(C{r}/12,-(D{r}+N(E{r})),B{r})"
    guard_min = f"D{r}<=B{r}*C{r}/12"
    guard_ext = f"(D{r}+N(E{r}))<=B{r}*C{r}/12"
    calc(ws, f"F{r}", f'=IF(B{r}="","",IF({guard_min},{LOW},IF(C{r}=0,CEILING(B{r}/D{r},1),CEILING({nper_min},1))))', "0")
    calc(ws, f"G{r}", f'=IF(OR(B{r}="",ISTEXT(F{r})),"",IF(C{r}=0,0,{nper_min}*D{r}-B{r}))', MONEY)
    calc(ws, f"H{r}", f'=IF(B{r}="","",IF({guard_ext},{LOW},IF(C{r}=0,CEILING(B{r}/(D{r}+N(E{r})),1),CEILING({nper_ext},1))))', "0")
    calc(ws, f"I{r}", f'=IF(OR(B{r}="",ISTEXT(H{r})),"",IF(C{r}=0,0,{nper_ext}*(D{r}+N(E{r}))-B{r}))', MONEY)
    calc(ws, f"J{r}", f'=IF(OR(G{r}="",I{r}=""),"",G{r}-I{r})', MONEY)
    calc(ws, f"K{r}", f'=IF(OR(ISTEXT(F{r}),ISTEXT(H{r}),F{r}="",H{r}=""),"",F{r}-H{r})', "0")
for i, row in enumerate(debts):
    for col, v in zip("ABCDE", row):
        ws[f"{col}{D1+i}"] = v
T = DN + 1
ws[f"A{T}"] = "Totals"
ws[f"A{T}"].font = BOLD
for col in "BDEGIJ":
    calc(ws, f"{col}{T}", f"=SUM({col}{D1}:{col}{DN})", MONEY, True)
    ws[f"{col}{T}"].fill = TOT_FILL
calc(ws, f"F{T}", f"=MAX(F{D1}:F{DN})", "0", True)
calc(ws, f"H{T}", f"=MAX(H{D1}:H{DN})", "0", True)
ws[f"A{T+2}"] = "Totals for months show the longest payoff. Interest figures are exact for a fixed payment and fixed APR, compounded monthly."
ws[f"A{T+2}"].font = Font(name=F, italic=True, size=9, color="666666")
ws.conditional_formatting.add(f"J{D1}:J{DN}", DataBarRule(start_type="min", end_type="max", color="4CAF50"))
dv = DataValidation(type="decimal", operator="between", formula1="0", formula2="1",
                    error="Enter APR as a percentage, e.g. 19.99%", showErrorMessage=True)
ws.add_data_validation(dv)
dv.add(f"C{D1}:C{DN}")
widths(ws, {"A": 26, "B": 14, "C": 10, "D": 14, "E": 14, "F": 14, "G": 16, "H": 14, "I": 16, "J": 16, "K": 12})
ws.freeze_panes = "B5"

# ---------------------------------------------------------------- Savings Goals
ws = wb.create_sheet("Savings Goals")
ws["A1"] = "Savings Goals"
ws["A1"].font = TITLE
header(ws, 3, ["Goal", "Target ($)", "Saved so far ($)", "Monthly contribution ($)",
               "Remaining ($)", "% complete", "Months to go", "Estimated finish"])
ws.row_dimensions[3].height = 32
goals = [("Emergency fund (example)", 10000, 2500, 400),
         ("Vacation (example)", 3000, 600, 150)]
G1, GN = 4, 13
for r in range(G1, GN + 1):
    inp(ws, f"A{r}")
    inp(ws, f"B{r}", fmt=MONEY)
    inp(ws, f"C{r}", fmt=MONEY)
    inp(ws, f"D{r}", fmt=MONEY)
    calc(ws, f"E{r}", f'=IF(B{r}="","",MAX(0,B{r}-N(C{r})))', MONEY)
    calc(ws, f"F{r}", f'=IF(OR(B{r}="",B{r}=0),"",MIN(1,N(C{r})/B{r}))', "0%")
    calc(ws, f"G{r}", f'=IF(E{r}="","",IF(E{r}=0,0,IF(N(D{r})<=0,"Add a contribution",CEILING(E{r}/D{r},1))))', "0")
    calc(ws, f"H{r}", f'=IF(OR(G{r}="",ISTEXT(G{r})),"",EDATE(TODAY(),G{r}))', "mmm yyyy")
for i, row in enumerate(goals):
    for col, v in zip("ABCD", row):
        ws[f"{col}{G1+i}"] = v
ws.conditional_formatting.add(f"F{G1}:F{GN}", DataBarRule(start_type="num", start_value=0,
                                                          end_type="num", end_value=1, color="4C8BF5"))
widths(ws, {"A": 26, "B": 13, "C": 15, "D": 16, "E": 14, "F": 12, "G": 13, "H": 16})
ws.freeze_panes = "A4"

for sheet in wb.worksheets:
    sheet.sheet_view.showGridLines = False
wb.save(OUT)
print("saved", OUT)
