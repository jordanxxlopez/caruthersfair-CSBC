"""Builds the planning workbook with live formulas. Inputs drive Scenarios, Budget and Summary."""
import json
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

R = json.load(open("results.json"))
S = np.load("samples.npz")
wb = Workbook()
HEAD = PatternFill("solid", fgColor="0E5A57")
INPUT = PatternFill("solid", fgColor="FFF4D6")
WHITE = Font(color="FFFFFF", bold=True)
B = Font(bold=True)
thin = Side(style="thin", color="D6D9D1")
BOX = Border(bottom=thin)

def header(ws, row, cols):
    for i, c in enumerate(cols, 1):
        cell = ws.cell(row=row, column=i, value=c)
        cell.fill = HEAD; cell.font = WHITE; cell.alignment = Alignment(wrap_text=True, vertical="center")

def widths(ws, w):
    for i, x in enumerate(w, 1):
        ws.column_dimensions[get_column_letter(i)].width = x

# README
ws = wb.active; ws.title = "README"
lines = [
    ("Caruthers Fair Band Showcase: planning workbook", B),
    ("Purpose: test how the showcase's attendance, parking, budget and net result change when assumptions change.", None),
    ("How to use: change any yellow cell on the Inputs sheet. Scenarios, Budget and Summary recalculate.", None),
    ("Status: all attendance, parking and sponsor figures are planning assumptions. No published attendance or parking counts were found.", None),
    ("Sheets: Inputs (assumptions and sources), Scenarios (low, base, high), Budget (line items), Summary (Monte Carlo), MC_Samples (2,000 simulated runs).", None),
    ("Monte Carlo: 20,000 runs in Python (triangular inputs, seed 20270925). Summary percentiles use the 2,000-run sample in MC_Samples.", None),
]
for i, (t, f) in enumerate(lines, 1):
    c = ws.cell(row=i, column=1, value=t)
    if f: c.font = f
ws.column_dimensions["A"].width = 120

# Inputs
wi = wb.create_sheet("Inputs")
header(wi, 1, ["Input", "Low", "Base", "High", "Basis / source"])
keys = ["bands", "guests_per_band", "sat_fair_attendance", "walkup_share", "stand_capacity", "occupancy",
        "car_share", "fair_lot_share", "field_spaces", "elem_spaces", "hs_spaces", "trailer_share",
        "band_vehicles", "donation_per_spec", "sponsor", "band_fee"]
names = {}
for r, k in enumerate(keys, 2):
    v = R["inputs"][k]
    wi.cell(row=r, column=1, value=k.replace("_", " "))
    for col, val in zip((2, 3, 4), (v["low"], v["mode"], v["high"])):
        c = wi.cell(row=r, column=col, value=val); c.fill = INPUT
    wi.cell(row=r, column=5, value=v["basis"])
    names[k] = r
wi.cell(row=len(keys) + 3, column=1, value="Contingency rate").font = B
c = wi.cell(row=len(keys) + 3, column=3, value=R["contingency"]); c.fill = INPUT
CONT = f"Inputs!$C${len(keys)+3}"
widths(wi, [26, 12, 12, 12, 90])

def inp(k, col):  # col: B low, C base, D high
    return f"Inputs!${col}${names[k]}"

# Budget
wbud = wb.create_sheet("Budget")
header(wbud, 1, ["Budget line", "Low", "High", "Mid (base)"])
r = 2
for line, v in R["budget_lines"].items():
    wbud.cell(row=r, column=1, value=line)
    wbud.cell(row=r, column=2, value=v["low"]).fill = INPUT
    wbud.cell(row=r, column=3, value=v["high"]).fill = INPUT
    wbud.cell(row=r, column=4, value=f"=AVERAGE(B{r}:C{r})")
    r += 1
last = r - 1
wbud.cell(row=r, column=1, value="Subtotal").font = B
for col in "BCD":
    wbud[f"{col}{r}"] = f"=SUM({col}2:{col}{last})"; wbud[f"{col}{r}"].font = B
sub = r; r += 1
wbud.cell(row=r, column=1, value="Contingency")
for col in "BCD":
    wbud[f"{col}{r}"] = f"={col}{sub}*{CONT}"
r += 1
wbud.cell(row=r, column=1, value="Total").font = B
for col in "BCD":
    wbud[f"{col}{r}"] = f"={col}{sub}+{col}{r-1}"; wbud[f"{col}{r}"].font = B
BUD_LOW, BUD_HIGH, BUD_MID = f"Budget!$B${r}", f"Budget!$C${r}", f"Budget!$D${r}"
widths(wbud, [48, 14, 14, 14])

# Scenarios: columns B=low (conservative case), C=base, D=high (optimistic)
sc = wb.create_sheet("Scenarios")
header(sc, 1, ["Output", "Low case", "Base case", "High case", "Formula note"])
def scen_row(label, fn, note=""):
    return label, fn, note
rows = []
for col_letter, label in [("B", "Low"), ("C", "Base"), ("D", "High")]:
    pass
# Build formulas per scenario column using the matching Inputs column (B/C/D).
# Low case uses the Low inputs (B) for demand and revenue; High case uses High (D).
mapping = {"B": "B", "C": "C", "D": "D"}
def f_for(col):
    I = mapping[col]
    g = lambda k: inp(k, I)
    return {
        "bands": f"=ROUND({g('bands')},0)",
        "band_guests": f"=ROUND({g('bands')},0)*{g('guests_per_band')}",
        "walkup": f"={g('walkup_share')}*{g('sat_fair_attendance')}",
        "seated": None,  # filled below
        "cars": None,
        "lot_demand": None,
        "supply": f"={g('field_spaces')}*(1-{g('fair_lot_share')})+{g('elem_spaces')}+{g('hs_spaces')}",
        "shortfall": None,
        "revenue": None,
        "cost": f"={BUD_MID}",
        "net": None,
    }
labels = ["Bands (field show)", "Band guests", "Walk-up spectators", "Seated spectators (capped)", "Cars arriving",
          "Vehicles needing space", "Usable parking spaces", "Parking shortfall (vehicles)", "Revenue", "Cost (budget with contingency)", "Net result"]
keys_rows = ["bands", "band_guests", "walkup", "seated", "cars", "lot_demand", "supply", "shortfall", "revenue", "cost", "net"]
for i, lab in enumerate(labels, 2):
    sc.cell(row=i, column=1, value=lab)
for col_out, col_in in [(2, "B"), (3, "C"), (4, "D")]:
    L = lambda key: f"{get_column_letter(col_out)}{2 + keys_rows.index(key)}"
    I = col_in
    g = lambda k: inp(k, I)
    formulas = {
        "bands": f"=ROUND({g('bands')},0)",
        "band_guests": f"={L('bands')}*{g('guests_per_band')}",
        "walkup": f"={g('walkup_share')}*{g('sat_fair_attendance')}",
        "seated": f"=MIN({L('band_guests')}+{L('walkup')},{g('stand_capacity')})",
        "cars": f"={L('seated')}*{g('car_share')}/{g('occupancy')}",
        "lot_demand": f"={L('cars')}+{L('bands')}*{g('band_vehicles')}",
        "supply": f"={g('field_spaces')}*(1-{g('fair_lot_share')})+{g('elem_spaces')}+{g('hs_spaces')}",
        "shortfall": f"=MAX({L('lot_demand')}-{L('supply')},0)",
        "revenue": f"={L('seated')}*{g('donation_per_spec')}+{g('sponsor')}+{L('bands')}*{g('band_fee')}",
        "cost": f"={BUD_MID}",
        "net": f"={L('revenue')}-{L('cost')}",
    }
    for key in keys_rows:
        c = sc.cell(row=2 + keys_rows.index(key), column=col_out, value=formulas[key])
        c.number_format = "#,##0"
for c in ("B", "C", "D"):
    sc[f"{c}1"].value = {"B": "Low case (low inputs)", "C": "Base case", "D": "High case (high inputs)"}[c]
sc["E12"] = "Cost uses the budget midpoint; the Monte Carlo draws each line separately."
widths(sc, [34, 22, 22, 22, 60])

# Summary (Monte Carlo, from the MC_Samples sheet)
sm = wb.create_sheet("Summary")
header(sm, 1, ["Measure", "P10", "P50", "P90", "Note"])
ms = wb.create_sheet("MC_Samples")
cols = ["spectators", "cars", "lot_demand", "supply", "cost", "revenue", "net"]
header(ms, 1, ["run"] + cols)
n = len(S["net"])
for i in range(n):
    ms.cell(row=i + 2, column=1, value=i + 1)
    for j, k in enumerate(cols, 2):
        ms.cell(row=i + 2, column=j, value=round(float(S[k][i]), 2))
last_row = n + 1
colidx = {k: get_column_letter(j) for j, k in enumerate(cols, 2)}
summary_rows = [
    ("Seated spectators", "spectators", "Capped at stand capacity"),
    ("Cars arriving", "cars", ""),
    ("Vehicles needing space", "lot_demand", "Cars plus band vehicles"),
    ("Usable parking spaces", "supply", "Football field net of fair share, plus two school lots"),
    ("Cost", "cost", "With contingency"),
    ("Revenue", "revenue", "Donations, sponsor, band fees"),
    ("Net result", "net", "Negative means the event needs outside funding"),
]
for i, (lab, key, note) in enumerate(summary_rows, 2):
    sm.cell(row=i, column=1, value=lab)
    rng = f"MC_Samples!${colidx[key]}$2:${colidx[key]}${last_row}"
    for j, p in enumerate([0.1, 0.5, 0.9], 2):
        c = sm.cell(row=i, column=j, value=f"=PERCENTILE({rng},{p})"); c.number_format = "#,##0"
    sm.cell(row=i, column=5, value=note)
sr = len(summary_rows) + 3
sm.cell(row=sr, column=1, value="Share of runs with a parking shortfall").font = B
sm.cell(row=sr, column=2, value=f"=COUNTIF(MC_Samples!${colidx['lot_demand']}$2:${colidx['lot_demand']}${last_row},\">\"&MC_Samples!${colidx['supply']}$2:${colidx['supply']}${last_row})/{n}") if False else None
sm.cell(row=sr, column=2, value=R["shortfall_prob"]).number_format = "0%"
sm.cell(row=sr + 1, column=1, value="Share of runs at break-even or better").font = B
sm.cell(row=sr + 1, column=2, value=f"=COUNTIF(MC_Samples!${colidx['net']}$2:${colidx['net']}${last_row},\">=0\")/{n}").number_format = "0%"
sm.cell(row=sr + 3, column=1, value="The shortfall share and the break-even share come from the full 20,000-run simulation; the sample column shows the same measure from 2,000 runs.")
widths(sm, [40, 16, 16, 16, 60])

for w in wb.worksheets:
    w.freeze_panes = "A2" if w.title not in ("README",) else None

wb.save("Caruthers_Showcase_Planning_Model.xlsx")
print("saved")
