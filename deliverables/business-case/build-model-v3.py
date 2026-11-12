# Builds v3-business-case-model.xlsx: calibrated OEE / predictive maintenance / AI & IoT business case.
# Run:  python3 build-model-v3.py   (then open in Excel; LibreOffice recalc is used to cache values)
# Style: IBM Carbon light (White theme) colours; Arial as the font (Carbon's fallback for IBM Plex Sans).
# All numbers are live formulas driven by the input cells on tab 2.

import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from openpyxl.comments import Comment

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "v3-business-case-model.xlsx")

# ---------- IBM Carbon light tokens ----------
INK, SEC, HELP = "161616", "525252", "6F6F6F"
BLUE, BLUE_BG, BLUE_DK = "0F62FE", "EDF5FF", "0043CE"
GREEN_LINK = "198038"
BORDER, BORDER_STRONG, LAYER, HDR = "E0E0E0", "8D8D8D", "F4F4F4", "E0E0E0"
KEY_BG = "FCF4D6"
OK, OK_BG, ERR, ERR_BG = "24A148", "DEFBE6", "DA1E28", "FFF1F1"
FONT = "Arial"

S1, S2, S3, S4 = "1 Explainer", "2 Assumptions & Refs", "3 P&L TCO NPV", "4 Sensitivity & Stress"
q = lambda s: "'" + s + "'"

EUR = '#,##0.0;(#,##0.0);"–"'
EUR2 = '#,##0.00;(#,##0.00);"–"'
EUR3 = '#,##0.000;(#,##0.000);"–"'
PCT = '0.0%;(0.0%);"–"'
PCT0 = '0%;(0%);"–"'
NUM = '#,##0;(#,##0);"–"'
YRS = '0.0" yrs"'

thin = Side(style="thin", color=BORDER)
strong = Side(style="thin", color=BORDER_STRONG)

wb = openpyxl.Workbook()
ws1 = wb.active; ws1.title = S1
ws2 = wb.create_sheet(S2); ws3 = wb.create_sheet(S3); ws4 = wb.create_sheet(S4)
for ws in (ws1, ws2, ws3, ws4):
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = BLUE if ws is ws1 else "393939"


def font(color=INK, size=10, bold=False, italic=False):
    return Font(name=FONT, size=size, bold=bold, italic=italic, color=color)


def put(ws, ref, value, color=INK, size=10, bold=False, fmt=None, fill=None, wrap=False,
        align=None, italic=False, border=True):
    c = ws[ref]
    c.value = value
    c.font = font(color, size, bold, italic)
    if fmt: c.number_format = fmt
    if fill: c.fill = PatternFill("solid", fgColor=fill)
    c.alignment = Alignment(wrap_text=wrap, vertical="top", horizontal=align)
    if border: c.border = Border(bottom=thin)
    return c


def title(ws, eyebrow, text, sub):
    put(ws, "B2", eyebrow.upper(), SEC, 9, border=False)
    put(ws, "B3", text, INK, 20, border=False)
    put(ws, "B4", sub, SEC, 10, border=False)
    ws.row_dimensions[3].height = 30


def section(ws, row, text, last_col="K", note=None):
    put(ws, f"B{row}", text, INK, 13, bold=True, border=False)
    ws.row_dimensions[row].height = 22
    for col in range(2, openpyxl.utils.column_index_from_string(last_col) + 1):
        ws.cell(row=row, column=col).border = Border(bottom=Side(style="medium", color=BLUE))
    if note:
        put(ws, f"B{row+1}", note, SEC, 9, border=False, italic=True)


RIGHT_HDR = {"Low", "Base", "High", "Active", "3-yr", "5-yr", "Value", "5 years", "7 years", "Run-rate", "Weight (sel.)",
             "Baseline", "After", "Change", "Value €m/yr", "In option?", "Quantity", "Unit cost", "Excl. cont.", "Incl. cont.",
             "Option A", "Option B", "Option C", "Base (model)", "CFO downside", "Severe stress", "Upside", "Concept", "Calibrated",
             "L1 PdM", "L2 OEE", "L3 Vision", "L4 Energy"}


def header(ws, row, values, start_col="B"):
    c0 = openpyxl.utils.column_index_from_string(start_col)
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=c0 + i, value=v)
        c.font = font(INK, 10, True)
        c.fill = PatternFill("solid", fgColor=HDR)
        c.alignment = Alignment(wrap_text=True, vertical="center",
                                horizontal="right" if (i > 0 and isinstance(v, str) and (v[:2] in ("20", "Y1", "Y2", "Y3", "Y4") or v in RIGHT_HDR)) else "left")
        c.border = Border(bottom=strong)
    ws.row_dimensions[row].height = 28


def name(n, sheet, ref):
    col = "".join(ch for ch in ref if ch.isalpha()); row = "".join(ch for ch in ref if ch.isdigit())
    wb.defined_names[n] = DefinedName(n, attr_text=f"{q(sheet)}!${col}${row}")


def widths(ws, d):
    for k, v in d.items():
        ws.column_dimensions[k].width = v


# =====================================================================================
# TAB 2 – ASSUMPTIONS & INDUSTRY REFERENCES
# =====================================================================================
widths(ws2, {"A": 2, "B": 9, "C": 52, "D": 11, "E": 10, "F": 10, "G": 10, "H": 11, "I": 11, "J": 9, "K": 100})
title(ws2, "Tab 2 · Assumptions & industry references",
      "Every input in one place, labelled and referenced",
      "Blue text = input you can change. Yellow = selector. Black = formula. Label: case · group · source · assumption · calc. "
      "Source codes refer to the reference table at the bottom of this tab.")

put(ws2, "B6", "Selectors", INK, 11, True, border=False)
put(ws2, "C7", "Scenario for all Low / Base / High inputs", border=True)
put(ws2, "F7", "Base", BLUE, 10, True, fill=KEY_BG, align="right")
put(ws2, "G7", "index", SEC, 9, border=True); put(ws2, "H7", '=MATCH(F7,{"Low","Base","High"},0)', fmt="0")
put(ws2, "K7", "Low = cautious on every lever, High = optimistic. Base is the board proposal.", SEC, 9)
put(ws2, "C8", "Option shown on tab 3 (P&L, TCO, cash flow)", border=True)
put(ws2, "F8", "C", BLUE, 10, True, fill=KEY_BG, align="right")
put(ws2, "G8", "index", SEC, 9); put(ws2, "H8", '=MATCH(F8,{"A","B","C"},0)', fmt="0")
put(ws2, "K8", "A = predictive maintenance only · B = connected OEE + predictive maintenance · C = full AI & IoT. All three are compared on tab 3 §3.6.", SEC, 9)
for r in (7, 8):
    for col in "BDE": ws2[f"{col}{r}"].border = Border(bottom=thin)
name("x_scen", S2, "F7"); name("x_scenIdx", S2, "H7"); name("x_opt", S2, "F8"); name("x_optIdx", S2, "H8")
dv = DataValidation(type="list", formula1='"Low,Base,High"', allow_blank=False); ws2.add_data_validation(dv); dv.add("F7")
dv2 = DataValidation(type="list", formula1='"A,B,C"', allow_blank=False); ws2.add_data_validation(dv2); dv2.add("F8")

row = 10
# ---- scalar input sections ----
# (id, label, unit, low, base, high, tag, src, rationale, fmt)
SECTIONS = [
    ("A. Company baseline, calibrated to listed peers", [
        ("rev0", "Revenue 2026, all plants", "€m", None, 2000, None, "group", "BC P1 P2 P3",
         "Group concept §1. Peer check: Ariston €2.7 bn (32 plants), BDR Thermea €1.9 bn (14), Bosch Home Comfort ≈ €4.4 bn (17). €2 bn sits inside the peer range.", EUR),
        ("plants", "Plants", "#", None, 10, None, "group", "BC",
         "Ten European capitals (fictional), Amsterdam = HQ and plant. €200 m per plant vs peer median ≈ €136 m: fewer, larger plants, so more value per standard roll-out and more at stake per OT incident.", NUM),
        ("lines", "Production lines", "#", None, 120, None, "group", "BC", "10 plants × 12 lines; 72 accu production, 48 assembly.", NUM),
        ("legacy", "of which legacy lines", "#", None, 18, None, "group", "BC", "15% of lines. Need protocol conversion, isolation and longer test windows, so they connect last and cost more.", NUM),
        ("growth", "Organic market growth (baseline only)", "%/yr", None, 0.03, None, "source", "P1 P2",
         "Ariston +3% organic FY2025, guidance 1–4% for 2026; Bosch HC ≈ +3% FX-adjusted. Grows the baseline P&L only; programme benefits are NOT indexed to growth (conservative).", PCT),
        ("mat", "Material cost, % of revenue", "%", None, 0.453, None, "source", "P1", "Ariston FY2025: (purchases + change in inventories) / revenue.", PCT),
        ("serv", "Services cost, % of revenue", "%", None, 0.182, None, "source", "P1", "Ariston FY2025 services / revenue (logistics, energy, external maintenance).", PCT),
        ("pers", "Personnel cost, % of revenue", "%", None, 0.246, None, "source", "P1", "Ariston FY2025 personnel / revenue.", PCT),
        ("ebitda", "EBITDA margin baseline", "%", None, 0.097, None, "source", "P1 P3",
         "Peer median adj. EBITDA margin (Ariston 11.5%, BDR Thermea 8.0%). Other operating cost is the residual.", PCT),
        ("da", "D&A baseline, % of revenue", "%", None, 0.051, None, "source", "P1", "Ariston FY2025 D&A / revenue.", PCT),
        ("tax", "Corporate income tax rate", "%", None, 0.258, None, "source", "T1", "Dutch CIT top rate (group HQ Amsterdam). Assumes group profits absorb programme losses in early years.", PCT),
        ("wacc", "Discount rate (WACC)", "%", None, 0.08, None, "source", "K1", "KPMG Cost of Capital Study 2026: average WACC 7.9%. Finance to confirm the hurdle rate; tab 4 tests 6–12%.", PCT),
    ]),
    ("B. Operations baseline (replace with pilot measurements)", [
        ("A0", "Availability (OEE component)", "%", None, 0.80, None, "assumption", "O1 O2",
         "No pilot baseline in the case. Typical discrete plants run OEE ≈ 60–67% (O1); world class ≈ 85% = 90% × 95% × 99.9% (O2).", PCT),
        ("P0", "Performance (speed loss, micro-stops)", "%", None, 0.84, None, "assumption", "O1", "Gives baseline OEE ≈ 65%, the middle of the typical range.", PCT),
        ("Q0", "Quality (first-pass yield)", "%", None, 0.97, None, "assumption", "O1", "3% of output is scrapped or reworked.", PCT),
        ("unpl", "Unplanned equipment failures, % of planned time", "%", None, 0.08, None, "assumption", "D1",
         "8 of the 20 availability points lost; the rest is changeovers, material and staffing. Siemens: unplanned downtime costs the 500 largest firms ≈ 11% of revenue.", PCT),
        ("maint", "Maintenance spend, % of revenue", "%", None, 0.02, None, "assumption", "–",
         "≈ €40 m (labour, parts, contractors). Peers do not disclose a maintenance line; v2 used 1.5%.", PCT),
        ("scrap", "Scrap and rework cost, % of revenue", "%", None, 0.02, None, "assumption", "–", "≈ €40 m; consistent with 97% quality.", PCT),
        ("energy", "Plant energy cost, % of revenue", "%", None, 0.015, None, "assumption", "P1", "≈ €30 m; sits inside the 18% services line.", PCT),
        ("spares", "Spare-parts (MRO) inventory", "€m", None, 25, None, "assumption", "–", "≈ 1.25% of revenue. Cash item (working capital), not P&L.", EUR),
        ("sell", "Sellable share of extra capacity", "%", 0.25, 0.50, 0.75, "assumption", "BC P1",
         "Concept assumed 100%, then 75% as an example. Peers grow ≈ 3%/yr, so 50% (≈ +1–2% extra growth) is the base. The unsold part frees overtime and temporary labour.", PCT0),
        ("varserv", "Variable share of services on extra volume", "%", None, 0.40, None, "assumption", "CAL", "Logistics, energy, consumables scale with volume (Ariston calibration workbook).", PCT0),
        ("flex", "Flexible share of personnel cost", "%", None, 0.15, None, "assumption", "–", "Overtime and temporary staff that can be cut when extra capacity is not sold.", PCT0),
        ("fp", "Cost of false / missed AI alerts, % of maintenance spend", "%", 0.04, 0.02, 0.01, "source", "B1",
         "Unneeded interventions and missed failures. Reliamag: ≈ 10% false positives can wipe out a programme's savings. Low scenario = worse.", PCT),
    ]),
    ("C. AI & IoT improvement levers (Low / Base / High)", [
        ("redDown", "L1 Predictive maintenance: fewer failure stops", "%", 0.15, 0.25, 0.40, "source", "B3 B4",
         "IoT condition sensors + AI anomaly / remaining-life models. Deloitte: uptime +10–20%; McKinsey: downtime up to −50% (mature). Base 25% = −2 OEE points of failure downtime.", PCT0),
        ("redMaint", "L1 Predictive maintenance: lower maintenance cost", "%", 0.05, 0.08, 0.12, "source", "B3", "Deloitte: maintenance cost −5–10%, planning time −20–50%.", PCT0),
        ("redSpares", "L1 Lower spare-parts inventory", "%", 0.05, 0.10, 0.20, "assumption", "–", "Fewer emergency spares when failures are predicted. One-off cash release.", PCT0),
        ("redPerf", "L2 Connected OEE: less speed loss & micro-stops", "%", 0.05, 0.10, 0.20, "assumption", "W1",
         "IoT line data + MES + AI root-cause analysis. WEF Lighthouses report far more (best in class); we take 10% of the 16% speed loss.", PCT0),
        ("redScrap", "L3 AI vision: less scrap and rework", "%", 0.08, 0.15, 0.25, "source", "Q1 W1",
         "P&G: scrap −10–20% with AI visual inspection; WEF Lighthouses: defects −41% (best in class).", PCT0),
        ("redEnergy", "L4 AI energy optimisation: lower energy cost", "%", 0.03, 0.06, 0.10, "source", "E1 W1",
         "IoT sub-metering + AI load scheduling. IEA: AI process optimisation 3–10% of energy cost; WEF Lighthouses: −28% (best in class).", PCT0),
    ]),
]
TAGCOL = {"source": BLUE_DK, "assumption": "8A3800", "group": SEC, "case": SEC, "calc": SEC}

for sec_title, items in SECTIONS:
    section(ws2, row, sec_title); row += 1
    header(ws2, row, ["ID", "Assumption", "Unit", "Low", "Base", "High", "Active", "Label", "Source", "Rationale and industry reference"]); row += 1
    for (i, lab, unit, lo, ba, hi, tag, src, rat, fmt) in items:
        put(ws2, f"B{row}", i, SEC, 9); put(ws2, f"C{row}", lab); put(ws2, f"D{row}", unit, SEC, 9)
        for col, v in (("E", lo), ("F", ba), ("G", hi)):
            if v is not None:
                put(ws2, f"{col}{row}", v, BLUE, fmt=fmt, fill=BLUE_BG)
            else:
                put(ws2, f"{col}{row}", None)
        put(ws2, f"H{row}", f"=IF(COUNT(E{row},G{row})=2,CHOOSE(x_scenIdx,E{row},F{row},G{row}),F{row})", INK, bold=True, fmt=fmt)
        put(ws2, f"I{row}", tag, TAGCOL.get(tag, SEC), 9); put(ws2, f"J{row}", src, SEC, 9)
        put(ws2, f"K{row}", rat, SEC, 9, wrap=True)
        name("x_" + i, S2, f"H{row}")
        row += 1
    if sec_title.startswith("B."):
        # calc rows
        for (i, lab, f, fmt, rat) in [
            ("OEE0", "Baseline OEE = A × P × Q", "=x_A0*x_P0*x_Q0", PCT, "Typical 60–67% (O1); world class ≈ 85% (O2)."),
            ("CM", "Contribution margin on extra sales", "=1-x_mat-x_varserv*x_serv", PCT, "1 − material − variable services. v2 used 30%; calibration with peer cost structure gives ≈ 47%."),
        ]:
            put(ws2, f"B{row}", i, SEC, 9); put(ws2, f"C{row}", lab, INK, italic=True)
            put(ws2, f"H{row}", f, INK, bold=True, fmt=fmt); put(ws2, f"I{row}", "calc", SEC, 9); put(ws2, f"K{row}", rat, SEC, 9, wrap=True)
            for col in "DEFGJ": ws2[f"{col}{row}"].border = Border(bottom=thin)
            name("x_" + i, S2, f"H{row}"); row += 1
    row += 1

# ---- roll-out series ----
section(ws2, row, "D. Roll-out plan and ramp-up (Year 1 = 2027)",
        note="Year 4+ applies to 2030 onwards. Lines are cumulative at year end; benefits use the average of start and end of year."); row += 2
header(ws2, row, ["ID", "Series", "Unit", "Y1 2027", "Y2 2028", "Y3 2029", "Y4+", "", "Label", "", "Rationale"]); row += 1
SERIES = [
    ("cumStd", "Standard lines connected (cumulative)", "#", [30, 72, 102, 102], NUM, "group",
     "Y1: pilot plant (12 lines) plus two more plants, partly. All 102 standard lines by end of Y3. Gate before each wave."),
    ("cumLeg", "Legacy lines connected (cumulative)", "#", [0, 4, 18, 18], NUM, "assumption", "Legacy last: isolate first, retrofit in Y2–Y3 after the standard pattern is proven."),
    ("adopt", "Adoption / model maturity on connected lines", "%", [0.6, 0.8, 0.95, 1.0], PCT0, "assumption",
     "Share of possible value captured: operators trusting alerts, process change, model learning. 60% in the first year is typical for a first wave."),
    ("phF", "Foundation investment phasing", "%", [0.5, 0.35, 0.15, 0], PCT0, "assumption", "Platform, OT segmentation, security and integration front-loaded: the foundation comes before scale."),
    ("phD", "Use-case development phasing", "%", [0.5, 0.3, 0.2, 0], PCT0, "assumption", "Model and dashboard development."),
    ("ramp", "Foundation run-cost ramp", "%", [0.6, 0.9, 1.0, 1.0], PCT0, "assumption", "Cloud and licences grow with use."),
]
SROW = {}
for (i, lab, unit, vals, fmt, tag, rat) in SERIES:
    put(ws2, f"B{row}", i, SEC, 9); put(ws2, f"C{row}", lab); put(ws2, f"D{row}", unit, SEC, 9)
    for col, v in zip("EFGH", vals):
        put(ws2, f"{col}{row}", v, BLUE, fmt=fmt, fill=BLUE_BG)
    put(ws2, f"I{row}", None); put(ws2, f"J{row}", tag, TAGCOL[tag], 9); put(ws2, f"K{row}", rat, SEC, 9, wrap=True)
    SROW[i] = row; row += 1
put(ws2, f"C{row}", "Check: total lines connected by Y3 = all lines", SEC, 9, italic=True)
put(ws2, f"H{row}", f'=IF(G{SROW["cumStd"]}+G{SROW["cumLeg"]}=x_lines,"OK","Check")', OK, 9, True)
row += 2

# ---- costs ----
COSTS = [
    ("E. Implementation cost, one-off (suggested, end-to-end)", [
        ("cPlat", "Central data & AI platform set-up (IoT hub, lakehouse, MLOps)", "€m", 4.0, "F", "Cloud-first, hybrid: analytics central, control stays local. Capitalised."),
        ("cOTseg", "OT segmentation, OT-DMZ and edge gateways", "€m/plant", 0.6, "F", "IEC 62443 zones & conduits; local buffering so plants run if the link to Amsterdam fails. v2 used €0.5 m."),
        ("cOTsec", "OT security: assessment, OT monitoring, privileged remote access", "€m", 2.5, "F", "Covers the case's remote-access and supplier-access risks (COBIT DSS05, APO10)."),
        ("cInteg", "ERP / MES integration and common data model", "€m", 3.0, "F", "The group standard that every plant reuses."),
        ("cPMO", "Programme, SAFe ART set-up, change management", "€m", 3.0, "F", "Expensed (not capitalised)."),
        ("cTrain", "Training and IT/OT academy", "€m", 2.0, "F", "≈ 8,000 staff (peer revenue/employee); operators, maintenance, engineers. Expensed."),
        ("cL1std", "L1 Sensors, gateways, install – per standard line", "€m/line", 0.06, "L1", "≈ 40 monitoring points × ≈ €0.5k installed wireless (C1: $200–500; wired > $2k) plus engineering and commissioning."),
        ("cL1leg", "L1 Retrofit – per legacy line", "€m/line", 0.20, "L1", "Protocol conversion, wired sensors, isolation, longer tests (v2: €200k)."),
        ("cL1dev", "L1 Failure models and CMMS integration", "€m", 1.5, "L1", "Work orders created from AI alerts, human approves."),
        ("cL2std", "L2 Line connectivity and MES connector – per standard line", "€m/line", 0.04, "L2", "PLC/SCADA read-only data via edge; no write-back to control."),
        ("cL2leg", "L2 Line connectivity – per legacy line", "€m/line", 0.12, "L2", "Extra gateways / protocol converters."),
        ("cL2dev", "L2 OEE dashboards and loss analytics", "€m", 1.0, "L2", "One standard OEE definition (ISO 22400) for all plants."),
        ("cL3st", "L3 AI vision station (cameras, lighting, edge GPU, integration)", "€m/station", 0.08, "L3", "Assumption; vendor quotes needed."),
        ("nL3", "L3 Number of quality-critical stations", "#", 60, "L3", "About half the lines get one station."),
        ("cL3dev", "L3 Vision model training and labelling", "€m", 1.0, "L3", ""),
        ("cL4pl", "L4 Sub-metering and AI energy optimisation – per plant", "€m/plant", 0.25, "L4", ""),
        ("cont", "Contingency on all one-off costs", "%", 0.15, "all", "OT projects overrun; tab 4 tests up to +100%."),
        ("life", "Depreciation period, capitalised items", "yrs", 5, "all", "Straight line, starting in the year of spend."),
    ]),
    ("F. Recurring run cost (per year at full scale)", [
        ("rPlat", "Platform cloud, licences and support", "€m/yr", 2.5, "F", ""),
        ("rFTE", "Loaded cost per data / AI / OT specialist", "€m/FTE", 0.11, "all", "Peer average personnel cost ≈ €62k per employee (P1); specialists cost more."),
        ("fteF", "Core platform & data team", "FTE", 8, "F", "Group IT, Amsterdam: platform, data engineering, MLOps, architecture."),
        ("rSOC", "OT security monitoring (SOC service)", "€m/yr", 0.8, "F", "24/7 monitoring of OT zones."),
        ("rChamp", "Plant digital champion (1 per plant)", "€m/plant", 0.09, "F", "Local OT/data role; owns go/no-go input and adoption."),
        ("rTrain", "Ongoing training", "€m/yr", 0.3, "F", ""),
        ("fteL1", "L1 reliability / data-science FTE", "FTE", 3, "L1", ""),
        ("rL1", "L1 PdM software and connectivity per line", "€m/line", 0.008, "L1", "C1: subscriptions ≈ $200–800 per asset per year."),
        ("fteL2", "L2 OEE / MES analysts", "FTE", 2, "L2", ""),
        ("rL2", "L2 MES / OEE licences per line", "€m/line", 0.005, "L2", ""),
        ("fteL3", "L3 vision engineers", "FTE", 2, "L3", ""),
        ("rL3", "L3 support per vision station", "€m/station", 0.006, "L3", ""),
        ("fteL4", "L4 energy analyst", "FTE", 1, "L4", ""),
        ("rL4", "L4 energy software per plant", "€m/plant", 0.03, "L4", ""),
        ("upkeep", "Hardware and software upkeep, % of cumulative capex", "%/yr", 0.04, "all", "Sensor replacement, firmware, model retraining."),
    ]),
]
for sec_title, items in COSTS:
    section(ws2, row, sec_title); row += 1
    header(ws2, row, ["ID", "Cost item", "Unit", "", "Base", "", "Active", "Block", "Label", "Rationale"]); row += 1
    for (i, lab, unit, v, blk, rat) in items:
        fmt = PCT0 if unit.startswith("%") else (NUM if unit in ("#", "FTE", "yrs") else (EUR3 if v < 0.1 else EUR2))
        put(ws2, f"B{row}", i, SEC, 9); put(ws2, f"C{row}", lab); put(ws2, f"D{row}", unit, SEC, 9)
        put(ws2, f"E{row}", None); put(ws2, f"G{row}", None)
        put(ws2, f"F{row}", v, BLUE, fmt=fmt, fill=BLUE_BG)
        put(ws2, f"H{row}", f"=F{row}", INK, bold=True, fmt=fmt)
        put(ws2, f"I{row}", blk, SEC, 9); put(ws2, f"J{row}", "assumption", TAGCOL["assumption"], 9)
        put(ws2, f"K{row}", rat, SEC, 9, wrap=True)
        name("x_" + i, S2, f"H{row}"); row += 1
    row += 1

# ---- options ----
section(ws2, row, "G. Options: which levers each option switches on (1 = in, 0 = out)"); row += 1
header(ws2, row, ["Option", "Description", "", "L1 PdM", "L2 OEE", "L3 Vision", "L4 Energy", "", "", "What it means"]); row += 1
OPTS = [("A", "Predictive maintenance only", [1, 0, 0, 0], "The pilot scaled as-is. Carries the full foundation cost on one lever."),
        ("B", "Connected OEE + predictive maintenance", [1, 1, 0, 0], "Adds line data and loss analytics: the OEE story of the group concept."),
        ("C", "Full AI & IoT (recommended)", [1, 1, 1, 1], "Same foundation, four use cases: the foundation is paid once and reused. This is the group-standard argument.")]
OPT_ROW0 = row
for (o, d, flags, rat) in OPTS:
    put(ws2, f"B{row}", o, INK, bold=True); put(ws2, f"C{row}", d); put(ws2, f"D{row}", None)
    for col, v in zip("EFGH", flags): put(ws2, f"{col}{row}", v, BLUE, fmt="0", fill=BLUE_BG, align="right")
    put(ws2, f"I{row}", None); put(ws2, f"J{row}", None); put(ws2, f"K{row}", rat, SEC, 9, wrap=True)
    row += 1
put(ws2, f"B{row}", "Sel.", SEC, 9); put(ws2, f"C{row}", "Selected option (from selector above)", INK, bold=True); put(ws2, f"D{row}", None)
for k, col in enumerate("EFGH"):
    put(ws2, f"{col}{row}", f"=INDEX({col}{OPT_ROW0}:{col}{OPT_ROW0+2},x_optIdx)", INK, bold=True, fmt="0", align="right")
    name(f"x_f{k+1}", S2, f"{col}{row}")
put(ws2, f"K{row}", '=INDEX(C' + str(OPT_ROW0) + ':C' + str(OPT_ROW0 + 2) + ',x_optIdx)', SEC, 9)
OPT_ROWS = {o: OPT_ROW0 + k for k, o in enumerate("ABC")}
row += 2

# ---- calibration check ----
CAL_ROW = row
section(ws2, row, "H. Calibration: group concept vs calibrated model vs industry"); row += 1
header(ws2, row, ["", "Metric", "", "Concept", "Calibrated", "", "", "", "", "Reading"]); row += 1
CAL_START = row  # filled after tab 3 rows exist (needs references)
row += 7

# ---- industry references ----
section(ws2, row, "I. Industry references (accessed November 2026)"); row += 1
header(ws2, row, ["Code", "What we take from it", "", "", "", "", "", "Type", "", "Reference (APA-style) and link"]); row += 1
REFS = [
    ("BC", "Group concept: 10 plants, 120 lines (18 legacy), €2 bn, 15–25% capacity scenarios", "case",
     "Group 3. (2026). Buderus business case OEE & predictive maintenance – concept [Unpublished group draft]. materials/group-drafts/ (fictional case)."),
    ("CAL", "Ariston-based calibration of Buderus cost structure and growth", "group",
     "Group 3. (2026). Ariston ratios – Buderus calibration [Workbook]. materials/group-drafts/Ariston_Ratios_Buderus_Calibration.xlsx"),
    ("P1", "Ariston FY2025: revenue €2,707 m, 32 plants; material 45.3%, services 18.2%, personnel 24.6%, adj. EBITDA 11.5%, D&A 5.1%, organic growth +3%", "source",
     "Ariston Group. (2026, March 3). FY 2025 results [Press release and analyst presentation]. https://www.aristongroup.com/en/investors"),
    ("P2", "Bosch Home Comfort: sales ≈ €4.4 bn, ≈ +3% FX-adjusted, 17 plants before the JCI-Hitachi acquisition", "source",
     "Bosch. (2026). Business year 2025 of the Bosch Home Comfort Group in figures. https://www.bosch-presse.de/pressportal/de/en/business-year-2025-of-the-bosch-home-comfort-group-in-figures-280903.html"),
    ("P3", "BDR Thermea: net sales ≈ €1.9 bn, 14 plants, normalised EBITDA €151 m (≈ 8%)", "source",
     "BDR Thermea Group. (2026, March 30). Annual review 2025. https://www.bdrthermeagroup.com"),
    ("O1", "OEE: median plant ≈ 60%, top quartile ≈ 75%, world class ≈ 85%; discrete average 66–67% (vendor surveys, indicative)", "source",
     "WorkTrek. (2025). Debunking 7 myths about OEE. https://worktrek.com/blog/overall-equipment-effectiveness/ ; TeepTrak. (2026). OEE benchmark 2026. https://teeptrak.com/en/oee-benchmark-2026/"),
    ("O2", "World-class OEE 85% = 90% availability × 95% performance × 99.9% quality (Nakajima, TPM)", "source",
     "Nakajima, S. (1988). Introduction to TPM: Total productive maintenance. Productivity Press."),
    ("D1", "Unplanned downtime costs the world's 500 largest companies ≈ 11% of revenue ($1.4 tn); automotive ≈ $2.3 m per hour", "source",
     "Siemens. (2024). The true cost of downtime 2024. https://assets.new.siemens.com/siemens/assets/api/uuid:1b43afb5-2d07-47f7-9eb7-893fe7d0bc59/TCOD-2024_original.pdf"),
    ("B1", "Review of PdM ROI claims; ≈ 10% false positives can wipe out savings; many quoted figures weakly sourced", "source",
     "Reliamag. (n.d.). Predictive maintenance ROI benchmarks: What the studies show. https://reliamag.com/guides/predictive-maintenance-roi-benchmarks-what-the-studies-show/"),
    ("B3", "PdM: uptime +10–20%, maintenance cost −5–10%, planning time −20–50%", "source",
     "Deloitte. (2024). Predictive maintenance [Position paper]. https://www.deloitte.com/content/dam/assets-zone2/de/de/docs/about/2024/Deloitte_Predictive-Maintenance_PositionPaper.pdf"),
    ("B4", "Digital maintenance and reliability: availability gains and cost reduction; PdM not always the most economic choice", "source",
     "McKinsey & Company. (n.d.). A smarter way to digitize maintenance and reliability; Prediction at scale. https://www.mckinsey.com/capabilities/operations/our-insights/a-smarter-way-to-digitize-maintenance-and-reliability"),
    ("W1", "Global Lighthouse Network (best in class): defects −41%, energy −28%, labour productivity +40% (Sep 2025 cohort)", "source",
     "World Economic Forum. (2025). Global Lighthouse Network 2025. https://www.weforum.org/press/2025/01/global-lighthouse-network-2025-world-economic-forum-recognizes-companies-transforming-manufacturing-through-innovation/"),
    ("Q1", "AI visual inspection cut scrap 10–20% on some products (P&G with Siemens)", "source",
     "Manufacturing Digital. (2024). Quality: How P&G cut scrap up to 20% with AI inspection. https://manufacturingdigital.com/news/quality-how-p-g-cut-scrap-up-to-20-with-ai-inspection"),
    ("E1", "AI-enabled process optimisation could cut energy cost by 3–10% in energy-intensive industry", "source",
     "International Energy Agency. (2025). Energy and AI. https://www.iea.org/reports/energy-and-ai"),
    ("C1", "Wireless vibration point installed ≈ $200–500 (wired > $2k); subscriptions ≈ $200–800 per asset per year (vendor data)", "source",
     "Tractian. (2025). The true cost of a condition monitoring program. https://tractian.com/en/blog/the-true-cost-of-a-condition-monitoring-program"),
    ("K1", "Average WACC across industries 7.9% (2026)", "source",
     "KPMG. (2026). Cost of capital study 2026. https://kpmg.com/de/en/insights/business-performance-and-resilience/capital-cost-study.html"),
    ("T1", "Dutch corporate income tax 25.8% on profits above €200k", "source",
     "Government of the Netherlands. (2026). Corporate income tax rates. https://www.belastingdienst.nl"),
]
for (c, what, typ, ref) in REFS:
    put(ws2, f"B{row}", c, INK, 9, True); put(ws2, f"C{row}", what, INK, 9, wrap=True)
    ws2.merge_cells(f"C{row}:H{row}")
    put(ws2, f"I{row}", typ, TAGCOL.get(typ, SEC), 9); put(ws2, f"J{row}", None)
    put(ws2, f"K{row}", ref, SEC, 9, wrap=True)
    ws2.row_dimensions[row].height = 26
    row += 1
put(ws2, f"B{row+1}", "Vendor and secondary sources (O1, B1, C1, Q1) are indicative only. Before quoting McKinsey or Deloitte in the paper, cite the original report. "
    "Best-in-class Lighthouse figures (W1) are used as a ceiling, not as the base case.", SEC, 9, italic=True, border=False)
ws2.freeze_panes = "D10"

# =====================================================================================
# TAB 3 – P&L, TCO, NPV
# =====================================================================================
widths(ws3, {"A": 2, "B": 62, "C": 11, "D": 12, "E": 11, "F": 11, "G": 11, "H": 11, "I": 11, "J": 11, "K": 11, "L": 12, "M": 12, "N": 60})
title(ws3, "Tab 3 · 3-year P&L view, TCO and NPV",
      "What AI and IoT do to the P&L, what it costs end-to-end, and what it is worth",
      "€ million unless stated. Selected option and scenario come from the selectors on tab 2. Year 1 = 2027.")
YC = list("EFGHIJK")
S2C = ["E", "F", "G", "H", "H", "H", "H"]  # tab-2 series column for each year
R = {}


def yrow(ws, r, label, fn, unit="€m", fmt=EUR, bold=False, color=INK, tot3=True, tot5=True, note=None, d=None, years=7, top=False):
    """fn(col, prev_col_or_None, idx) -> formula string."""
    put(ws, f"B{r}", label, color if color != INK else INK, bold=bold)
    put(ws, f"C{r}", unit, SEC, 9)
    put(ws, f"D{r}", d, INK, bold=bold, fmt=fmt)
    for k in range(7):
        col = YC[k]
        if k < years:
            put(ws, f"{col}{r}", fn(col, YC[k - 1] if k else None, k), INK, bold=bold, fmt=fmt)
        else:
            put(ws, f"{col}{r}", None)
    put(ws, f"L{r}", f"=SUM(E{r}:G{r})" if tot3 else None, INK, bold=True, fmt=fmt)
    put(ws, f"M{r}", f"=SUM(E{r}:I{r})" if (tot5 and years >= 5) else None, INK, bold=True, fmt=fmt)
    put(ws, f"N{r}", note, SEC, 9, wrap=False)
    if top:
        for col in "BCDEFGHIJKLMN":
            ws[f"{col}{r}"].border = Border(top=strong, bottom=thin)


row = 6
# ---- KPI header (filled with refs later) ----
section(ws3, row, "Headline: selected option and scenario", last_col="N"); KPI_ROW = row + 1
row += 9

# ---- 3.0 timeline ----
section(ws3, row, "3.0 Timeline and roll-out drivers", last_col="N"); row += 1
header(ws3, row, ["Driver", "Unit", "2026 / run-rate", "2027", "2028", "2029", "2030", "2031", "2032", "2033", "3-yr", "5-yr", "Note"]); row += 1
R["t"] = row; yrow(ws3, row, "Year index", lambda c, p, k: str(k + 1), "#", "0", tot3=False, tot5=False); row += 1
for c in YC: ws3[f"{c}{R['t']}"].value = int(ws3[f"{c}{R['t']}"].value)
R["cumStd"] = row; yrow(ws3, row, "Standard lines connected, cumulative", lambda c, p, k: f"={q(S2)}!{S2C[k]}{SROW['cumStd']}", "#", NUM, color=GREEN_LINK, tot3=False, tot5=False); row += 1
R["cumLeg"] = row; yrow(ws3, row, "Legacy lines connected, cumulative", lambda c, p, k: f"={q(S2)}!{S2C[k]}{SROW['cumLeg']}", "#", NUM, tot3=False, tot5=False); row += 1
R["newStd"] = row; yrow(ws3, row, "New standard lines connected in year", lambda c, p, k: f"={c}{R['cumStd']}" + (f"-{p}{R['cumStd']}" if p else ""), "#", NUM); row += 1
R["newLeg"] = row; yrow(ws3, row, "New legacy lines connected in year", lambda c, p, k: f"={c}{R['cumLeg']}" + (f"-{p}{R['cumLeg']}" if p else ""), "#", NUM); row += 1
R["avg"] = row; yrow(ws3, row, "Average lines live in year", lambda c, p, k: f"=(({c}{R['cumStd']}+{c}{R['cumLeg']})" + (f"+({p}{R['cumStd']}+{p}{R['cumLeg']})" if p else "") + ")/2", "#", '#,##0.0', tot3=False, tot5=False); row += 1
R["adopt"] = row; yrow(ws3, row, "Adoption / model maturity", lambda c, p, k: f"={q(S2)}!{S2C[k]}{SROW['adopt']}", "%", PCT0, tot3=False, tot5=False); row += 1
R["phF"] = row; yrow(ws3, row, "Foundation investment phasing", lambda c, p, k: f"={q(S2)}!{S2C[k]}{SROW['phF']}", "%", PCT0, tot5=False); row += 1
R["phD"] = row; yrow(ws3, row, "Use-case development phasing", lambda c, p, k: f"={q(S2)}!{S2C[k]}{SROW['phD']}", "%", PCT0, tot5=False); row += 1
R["ramp"] = row; yrow(ws3, row, "Foundation run-cost ramp", lambda c, p, k: f"={q(S2)}!{S2C[k]}{SROW['ramp']}", "%", PCT0, tot3=False, tot5=False); row += 1
R["real"] = row; yrow(ws3, row, "Benefit realisation (share of run-rate)", lambda c, p, k: f"={c}{R['avg']}/x_lines*{c}{R['adopt']}", "%", PCT, bold=True, tot3=False, tot5=False,
                       note="Average lines live ÷ all lines × adoption. Every benefit follows this curve."); row += 1
R["df"] = row; yrow(ws3, row, "Discount factor (end of year)", lambda c, p, k: f"=1/(1+x_wacc)^{c}{R['t']}", "x", "0.000", tot3=False, tot5=False); row += 1
for k, c in enumerate(YC):
    for key in ("cumStd", "cumLeg", "adopt", "phF", "phD", "ramp"):
        ws3[f"{c}{R[key]}"].font = font(GREEN_LINK)
row += 1

# ---- 3.1 impact on the business ----
section(ws3, row, "3.1 Impact of AI and IoT on the business (full run-rate, selected scenario)", last_col="N",
        note="What each lever changes in operations, and what that is worth per year once all lines are live. In option? = included in the selected option."); row += 2
header(ws3, row, ["Lever and technology", "KPI", "Baseline", "After", "Change", "Value €m/yr", "In option?", "", "", "", "", "", "How the value is created"]); row += 1


def rr(key, label, kpi, base_f, after_f, fmt, value_f, flag, note, chg_fmt=None):
    global row
    put(ws3, f"B{row}", label); put(ws3, f"C{row}", kpi, SEC, 9)
    put(ws3, f"D{row}", base_f, INK, fmt=fmt); put(ws3, f"E{row}", after_f, INK, fmt=fmt)
    put(ws3, f"F{row}", f"=E{row}-D{row}" if base_f is not None else None, INK, fmt=chg_fmt or fmt)
    put(ws3, f"G{row}", value_f, INK, bold=True, fmt=EUR)
    put(ws3, f"H{row}", flag, INK, fmt="0", align="right")
    for col in "IJKLM": put(ws3, f"{col}{row}", None)
    put(ws3, f"N{row}", note, SEC, 9)
    R[key] = row; name("x_rr_" + key, S3, f"G{row}"); row += 1


PTS = '+0.0%;-0.0%;"–"'  # change in a share, shown as percentage points
PTN = '+0.0" pt";-0.0" pt";"–"'  # OEE points as whole numbers (tab 4)
# helpers for capacity value
put(ws3, f"B{row}", "Value of 100% extra capacity (sold at contribution margin, unsold frees flexible labour)", SEC, 9, italic=True)
put(ws3, f"C{row}", "€m", SEC, 9)
put(ws3, f"G{row}", "=x_rev0*(x_sell*x_CM+(1-x_sell)*x_pers*x_flex)", INK, fmt=EUR); name("x_Vcap", S3, f"G{row}")
for col in "DEFHIJKLM": put(ws3, f"{col}{row}", None)
put(ws3, f"N{row}", "Revenue × [sellable × CM + (1 − sellable) × personnel % × flexible %]", SEC, 9); R["Vcap"] = row; row += 1
put(ws3, f"B{row}", "Capacity gain from availability (L1) and performance (L2)", SEC, 9, italic=True)
put(ws3, f"C{row}", "a / p", SEC, 9)
put(ws3, f"D{row}", "=x_unpl*x_redDown/x_A0", INK, fmt=PCT); put(ws3, f"E{row}", "=(1-x_P0)*x_redPerf/x_P0", INK, fmt=PCT)
name("x_a", S3, f"D{row}"); name("x_p", S3, f"E{row}")
for col in "FGHIJKLM": put(ws3, f"{col}{row}", None)
put(ws3, f"N{row}", "a = availability points gained ÷ A0; p = performance points gained ÷ P0", SEC, 9); row += 1

rr("capL1", "L1 Predictive maintenance – IoT condition sensors + AI failure prediction", "Availability",
   "=x_A0", "=x_A0+x_unpl*x_redDown", PCT, "=x_a*x_Vcap", "=x_f1",
   "Failures predicted → repairs planned in breaks/changeovers → fewer stops → more output on the same lines", PTS)
rr("maint", "L1 … maintenance cost", "€m/yr", "=x_rev0*x_maint", "=D{r}*(1-x_redMaint)".replace("{r}", str(row)), EUR, "=x_rev0*x_maint*x_redMaint", "=x_f1",
   "Condition-based instead of calendar-based work; fewer emergency call-outs and rush parts")
rr("fp", "L1 … cost of false / missed alerts (negative)", "€m/yr", None, None, EUR, "=-x_rev0*x_maint*x_fp", "=x_f1",
   "AI risk made explicit: unnecessary interventions and missed failures. Human approves every work order.")
rr("capL2", "L2 Connected OEE – IoT line data + MES + AI loss analytics", "Performance",
   "=x_P0", "=x_P0+(1-x_P0)*x_redPerf", PCT, "=x_p*x_Vcap", "=x_f2",
   "Real-time micro-stop and speed-loss data per line; AI finds root causes; teams fix the top losses weekly", PTS)
rr("X", "L1 × L2 interaction (availability × performance)", "capacity", None, None, PCT, "=x_a*x_p*x_Vcap", "=x_f1*x_f2",
   "OEE multiplies: the two gains compound slightly")
rr("scrap", "L3 AI vision – cameras + edge AI at quality-critical stations", "Scrap €m/yr", "=x_rev0*x_scrap", f"=D{row}*(1-x_redScrap)", EUR,
   "=x_rev0*x_scrap*x_redScrap", "=x_f3", "Defects caught at the station, not at end-of-line or in the field. Valued as cost saving only (no extra capacity, to avoid double counting)")
rr("energy", "L4 AI energy optimisation – IoT sub-metering + AI load scheduling", "Energy €m/yr", "=x_rev0*x_energy", f"=D{row}*(1-x_redEnergy)", EUR,
   "=x_rev0*x_energy*x_redEnergy", "=x_f4", "Idle-time shutdowns, compressed-air leaks, peak shaving, heat recovery; also lowers CO2 (ESG)")
# totals
put(ws3, f"B{row}", "Gross run-rate benefit, selected option", INK, bold=True)
put(ws3, f"G{row}", f"=G{R['capL1']}*H{R['capL1']}+G{R['maint']}*H{R['maint']}+G{R['fp']}*H{R['fp']}+G{R['capL2']}*H{R['capL2']}+G{R['X']}*H{R['X']}+G{R['scrap']}*H{R['scrap']}+G{R['energy']}*H{R['energy']}",
    INK, bold=True, fmt=EUR)
name("x_RR", S3, f"G{row}")
for col in "CDEFHIJKLMN": put(ws3, f"{col}{row}", None)
for col in "BCDEFGHIJKLMN": ws3[f"{col}{row}"].border = Border(top=strong, bottom=thin)
R["RR"] = row; row += 1
put(ws3, f"B{row}", "OEE, selected option", INK, bold=True); put(ws3, f"C{row}", "OEE", SEC, 9)
put(ws3, f"D{row}", "=x_OEE0", INK, bold=True, fmt=PCT)
put(ws3, f"E{row}", "=(x_A0+x_unpl*x_redDown*x_f1)*(x_P0+(1-x_P0)*x_redPerf*x_f2)*(x_Q0+(1-x_Q0)*x_redScrap*x_f3)", INK, bold=True, fmt=PCT)
put(ws3, f"F{row}", f"=E{row}-D{row}", INK, bold=True, fmt=PTS)
for col in "GHIJKLM": put(ws3, f"{col}{row}", None)
put(ws3, f"N{row}", "World class ≈ 85% (O2). Quality gain shown in OEE but valued only as scrap saving.", SEC, 9)
R["OEE"] = row; name("x_OEE1", S3, f"E{row}"); row += 1
put(ws3, f"B{row}", "Capacity gain valued (availability × performance)", INK, bold=True); put(ws3, f"C{row}", "%", SEC, 9)
put(ws3, f"E{row}", "=(1+x_a*x_f1)*(1+x_p*x_f2)-1", INK, bold=True, fmt=PCT)
for col in "DFGHIJKLM": put(ws3, f"{col}{row}", None)
put(ws3, f"N{row}", "Group concept assumed 15–25%. Calibrated: a few percent, inside the peer reality check (§H on tab 2).", SEC, 9)
R["cap"] = row; name("x_capSel", S3, f"E{row}"); row += 2


# ---- engine layout (rows fixed up front, written at the bottom of tab 3) ----
SEC31_END = row
ENGINE_START = 194  # tuned so §3.2–3.6 fit above; asserted below
BLOCKS = ["F", "L1", "L2", "L3", "L4", "X"]
BNAME = {"F": "Foundation: platform, OT segmentation, security, integration, core team",
         "L1": "L1 Predictive maintenance (IoT + AI)", "L2": "L2 Connected OEE (IoT + MES + AI)",
         "L3": "L3 AI vision quality", "L4": "L4 AI energy optimisation", "X": "L1 × L2 interaction"}
WSEL = {"F": 1, "L1": "=x_f1", "L2": "=x_f2", "L3": "=x_f3", "L4": "=x_f4", "X": "=x_f1*x_f2"}
E = {}
r_ = ENGINE_START + 3
for b in BLOCKS:
    kinds = ["w", "ben", "opex"] + (["oneoff"] if b == "F" else []) + ["capex", "da", "wc", "fcf"]
    for kd in kinds:
        E[(b, kd)] = r_; r_ += 1
    r_ += 1
ENGINE_END = r_


def wsum(kind, c, weights=None, blocks=BLOCKS):
    terms = []
    for b in blocks:
        if (b, kind) not in E: continue
        w = weights[b] if weights else f"$D${E[(b, 'w')]}"
        terms.append(f"{w}*{c}{E[(b, kind)]}")
    return "=" + "+".join(terms)


def hdr_years(ws, r, first="Line", third="2026 / base"):
    header(ws, r, [first, "Unit", third, "2027", "2028", "2029", "2030", "2031", "2032", "2033", "3-yr", "5-yr", "Note"])


# ---- 3.2 benefit build ----
row = SEC31_END
section(ws3, row, "3.2 Benefit build-up, selected option (€m per year)", last_col="N",
        note="Run-rate value per lever × option flag × benefit realisation curve (§3.0)."); row += 2
hdr_years(ws3, row, "Benefit", "Run-rate"); row += 1
rl = R["real"]
BEN_LINES = [("capL1", "L1 Extra capacity from fewer failures", "x_f1"), ("maint", "L1 Lower maintenance cost", "x_f1"),
             ("fp", "L1 Cost of false / missed alerts", "x_f1"), ("capL2", "L2 Extra capacity from less speed loss", "x_f2"),
             ("X", "L1 × L2 interaction", "x_f1*x_f2"), ("scrap", "L3 Less scrap and rework", "x_f3"), ("energy", "L4 Lower energy cost", "x_f4")]
b0 = row
for key, lab, fl in BEN_LINES:
    yrow(ws3, row, lab, lambda c, p, k, key=key, fl=fl: f"={c}{rl}*x_rr_{key}*{fl}", d=f"=x_rr_{key}*{fl}"); row += 1
R["benTot"] = row
yrow(ws3, row, "Gross benefits (EBITDA)", lambda c, p, k: f"=SUM({c}{b0}:{c}{row-1})", bold=True, top=True, d=f"=SUM(D{b0}:D{row-1})"); row += 1
yrow(ws3, row, "Check: equals engine (should be 0)", lambda c, p, k: f"=ROUND({c}{R['benTot']}-({wsum('ben', c)[1:]}),6)", color=SEC, fmt=EUR3, tot3=False, tot5=False); R["chkBen"] = row; row += 2

# ---- 3.3 P&L ----
section(ws3, row, "3.3 Three-year P&L view: baseline, programme effect and pro-forma (€m)", last_col="N",
        note="Baseline built from the peer-calibrated cost structure (tab 2 §A). Costs shown negative. Programme effects follow the selected option and scenario."); row += 2
hdr_years(ws3, row, "P&L line", "2026 base"); row += 1
put(ws3, f"B{row}", "Baseline (no programme)", INK, 10, True); row += 1
P = {}


def pl(key, label, fyear, d=None, fmt=EUR, bold=False, top=False, color=INK, tot3=True, note=None):
    global row
    yrow(ws3, row, label, fyear, unit=("%" if "%" in fmt else "€m"), fmt=fmt, bold=bold, top=top, color=color, tot3=tot3, tot5=False, years=3, d=d, note=note)
    P[key] = row; row += 1


t = R["t"]
pl("rev", "Revenue", lambda c, p, k: f"=x_rev0*(1+x_growth)^{c}{t}", d="=x_rev0", note="Grows with organic market growth")
pl("mat", "Material cost", lambda c, p, k: f"=-{c}{P['rev']}*x_mat", d=f"=-D{P['rev']}*x_mat")
pl("serv", "Services", lambda c, p, k: f"=-{c}{P['rev']}*x_serv", d=f"=-D{P['rev']}*x_serv")
pl("pers", "Personnel", lambda c, p, k: f"=-{c}{P['rev']}*x_pers", d=f"=-D{P['rev']}*x_pers")
pl("oth", "Other operating cost", lambda c, p, k: f"=-{c}{P['rev']}*(1-x_mat-x_serv-x_pers-x_ebitda)", d=f"=-D{P['rev']}*(1-x_mat-x_serv-x_pers-x_ebitda)", note="Residual to the peer-median EBITDA margin")
pl("ebitda", "EBITDA", lambda c, p, k: f"=SUM({c}{P['rev']}:{c}{P['oth']})", d=f"=SUM(D{P['rev']}:D{P['oth']})", bold=True, top=True)
pl("da", "Depreciation & amortisation", lambda c, p, k: f"=-{c}{P['rev']}*x_da", d=f"=-D{P['rev']}*x_da")
pl("ebit", "EBIT", lambda c, p, k: f"={c}{P['ebitda']}+{c}{P['da']}", d=f"=D{P['ebitda']}+D{P['da']}", bold=True, top=True)
pl("ebitm", "EBIT margin", lambda c, p, k: f"={c}{P['ebit']}/{c}{P['rev']}", d=f"=D{P['ebit']}/D{P['rev']}", fmt=PCT, tot3=False)
row += 1
put(ws3, f"B{row}", "Programme effect (selected option)", INK, 10, True); row += 1
pl("cap", "Capacity gain realised", lambda c, p, k: f"={c}{rl}*x_capSel", fmt=PCT, tot3=False, color=SEC)
pl("xrev", "Extra revenue (sold share of extra capacity)", lambda c, p, k: f"=x_rev0*{c}{P['cap']}*x_sell", note="IoT + AI raise availability and performance; only the sellable share becomes revenue")
pl("xvar", "Variable cost on extra revenue", lambda c, p, k: f"=-{c}{P['xrev']}*(x_mat+x_varserv*x_serv)", note="Material + variable services")
pl("xlab", "Overtime / temp labour avoided (unsold capacity)", lambda c, p, k: f"=x_rev0*x_pers*x_flex*{c}{P['cap']}*(1-x_sell)")
pl("xmnt", "Maintenance saving, net of false alerts (L1)", lambda c, p, k: f"={c}{rl}*(x_rr_maint+x_rr_fp)*x_f1")
pl("xscr", "Scrap & rework saving (L3)", lambda c, p, k: f"={c}{rl}*x_rr_scrap*x_f3")
pl("xen", "Energy saving (L4)", lambda c, p, k: f"={c}{rl}*x_rr_energy*x_f4")
pl("xrun", "Recurring run cost (platform, licences, team, OT SOC, upkeep)",
   lambda c, p, k: f"=-({wsum('opex', c)[1:]}-{c}{E[('F','oneoff')]})")
pl("xone", "One-off implementation cost expensed (programme, training)", lambda c, p, k: f"=-{c}{E[('F','oneoff')]}")
pl("xebitda", "EBITDA impact", lambda c, p, k: f"=SUM({c}{P['xrev']}:{c}{P['xone']})", bold=True, top=True)
pl("xda", "Extra depreciation (capitalised investment)", lambda c, p, k: f"=-({wsum('da', c)[1:]})")
pl("xebit", "EBIT impact", lambda c, p, k: f"={c}{P['xebitda']}+{c}{P['xda']}", bold=True, top=True)
pl("chk", "Check: EBITDA impact equals engine (should be 0)",
   lambda c, p, k: f"=ROUND({c}{P['xebitda']}-({wsum('ben', c)[1:]}-({wsum('opex', c)[1:]})),6)", fmt=EUR3, color=SEC, tot3=False)
row += 1
put(ws3, f"B{row}", "Pro-forma (baseline + programme)", INK, 10, True); row += 1
pl("frev", "Revenue", lambda c, p, k: f"={c}{P['rev']}+{c}{P['xrev']}", d=f"=D{P['rev']}")
pl("febitda", "EBITDA", lambda c, p, k: f"={c}{P['ebitda']}+{c}{P['xebitda']}", d=f"=D{P['ebitda']}", bold=True)
pl("febitdam", "EBITDA margin", lambda c, p, k: f"={c}{P['febitda']}/{c}{P['frev']}", d=f"=D{P['febitda']}/D{P['frev']}", fmt=PCT, tot3=False)
pl("febit", "EBIT", lambda c, p, k: f"={c}{P['ebit']}+{c}{P['xebit']}", d=f"=D{P['ebit']}", bold=True)
pl("febitm", "EBIT margin", lambda c, p, k: f"={c}{P['febit']}/{c}{P['frev']}", d=f"=D{P['febit']}/D{P['frev']}", fmt=PCT, tot3=False)
pl("dm", "Change in EBIT margin vs baseline", lambda c, p, k: f"={c}{P['febitm']}-{c}{P['ebitm']}", fmt='+0.00%;-0.00%;"–"', tot3=False,
   note="Year 1–2 dip is the J-curve: foundation first, value follows the roll-out")
row += 1

# ---- 3.4 TCO ----
section(ws3, row, "3.4 Total cost of ownership, selected option (€m, undiscounted)", last_col="N",
        note="Everything it costs to implement and run end-to-end: investment, one-off expensed cost and recurring run cost."); row += 2
hdr_years(ws3, row, "Cost line", ""); row += 1
T = {}
t0 = row
for key, lab, f in [
    ("cF", "Investment – foundation (platform, OT segmentation & edge, OT security, integration)", lambda c, p, k: f"={c}{E[('F','capex')]}"),
    ("cL1", "Investment – L1 sensors, retrofit, failure models", lambda c, p, k: f"=$D${E[('L1','w')]}*{c}{E[('L1','capex')]}"),
    ("cL2", "Investment – L2 line connectivity, OEE analytics", lambda c, p, k: f"=$D${E[('L2','w')]}*{c}{E[('L2','capex')]}"),
    ("cL3", "Investment – L3 AI vision stations and models", lambda c, p, k: f"=$D${E[('L3','w')]}*{c}{E[('L3','capex')]}"),
    ("cL4", "Investment – L4 energy metering and optimisation", lambda c, p, k: f"=$D${E[('L4','w')]}*{c}{E[('L4','capex')]}"),
    ("one", "One-off expensed – programme, SAFe set-up, training", lambda c, p, k: f"={c}{E[('F','oneoff')]}"),
    ("runF", "Run – foundation (cloud, core team, OT SOC, champions, upkeep)", lambda c, p, k: f"={c}{E[('F','opex')]}-{c}{E[('F','oneoff')]}"),
    ("runU", "Run – use cases (licences, specialist FTE, upkeep)", lambda c, p, k: wsum("opex", c, blocks=["L1", "L2", "L3", "L4"])),
]:
    yrow(ws3, row, lab, f); T[key] = row; row += 1
T["tco"] = row; yrow(ws3, row, "Total cost of ownership", lambda c, p, k: f"=SUM({c}{t0}:{c}{row-1})", bold=True, top=True); row += 1
T["cum"] = row; yrow(ws3, row, "Cumulative TCO", lambda c, p, k: f"={c}{T['tco']}" + (f"+{p}{T['cum']}" if p else ""), tot3=False, tot5=False, color=SEC); row += 2

# cost build-up table
put(ws3, f"B{row}", "Implementation cost build-up, one-off, end-to-end (€m)", INK, 11, True, border=False); row += 1
header(ws3, row, ["Item", "Quantity", "Unit cost", "Excl. cont.", "Incl. cont.", "Option A", "Option B", "Option C", "", "", "", "", "Note"]); row += 1
fl = lambda o, col: f"{q(S2)}!${col}${OPT_ROWS[o]}"
BU = [("Central data & AI platform", "1", "x_cPlat", None, "Foundation"),
      ("OT segmentation, OT-DMZ, edge gateways", "=x_plants", "x_cOTseg", None, "per plant"),
      ("OT security tooling and remote-access control", "1", "x_cOTsec", None, ""),
      ("ERP / MES integration, common data model", "1", "x_cInteg", None, ""),
      ("Programme, SAFe ART, change (expensed)", "1", "x_cPMO", None, ""),
      ("Training and IT/OT academy (expensed)", "1", "x_cTrain", None, ""),
      ("L1 sensors & install – standard lines", "=x_lines-x_legacy", "x_cL1std", "E", "per line"),
      ("L1 retrofit – legacy lines", "=x_legacy", "x_cL1leg", "E", "per line"),
      ("L1 failure models, CMMS integration", "1", "x_cL1dev", "E", ""),
      ("L2 connectivity – standard lines", "=x_lines-x_legacy", "x_cL2std", "F", "per line"),
      ("L2 connectivity – legacy lines", "=x_legacy", "x_cL2leg", "F", "per line"),
      ("L2 OEE dashboards, loss analytics", "1", "x_cL2dev", "F", ""),
      ("L3 AI vision stations", "=x_nL3", "x_cL3st", "G", "per station"),
      ("L3 vision models, labelling", "1", "x_cL3dev", "G", ""),
      ("L4 sub-metering, energy AI", "=x_plants", "x_cL4pl", "H", "per plant")]
bu0 = row
for (lab, qty, unit, flagcol, note) in BU:
    put(ws3, f"B{row}", lab); put(ws3, f"C{row}", int(qty) if qty == "1" else qty, INK, fmt=NUM)
    put(ws3, f"D{row}", f"={unit}", GREEN_LINK, fmt=EUR3)
    put(ws3, f"E{row}", f"=C{row}*D{row}", fmt=EUR); put(ws3, f"F{row}", f"=E{row}*(1+x_cont)", fmt=EUR)
    for o, col in zip("ABC", "GHI"):
        put(ws3, f"{col}{row}", f"=F{row}" + (f"*{fl(o, flagcol)}" if flagcol else ""), fmt=EUR)
    for col in "JKLM": put(ws3, f"{col}{row}", None)
    put(ws3, f"N{row}", note, SEC, 9); row += 1
put(ws3, f"B{row}", "Total one-off implementation cost", INK, bold=True)
for col in "EFGHI":
    put(ws3, f"{col}{row}", f"=SUM({col}{bu0}:{col}{row-1})", INK, bold=True, fmt=EUR)
for col in "CDJKLMN": put(ws3, f"{col}{row}", None)
for col in "BCDEFGHIJKLMN": ws3[f"{col}{row}"].border = Border(top=strong, bottom=thin)
T["buTot"] = row
put(ws3, f"N{row}", "Peer capex envelope ≈ 5% of revenue ≈ €100 m per year (P1): this programme is a fraction of 1 year", SEC, 9); row += 1
put(ws3, f"B{row}", "Check: selected option equals 7-yr engine total (should be 0)", SEC, 9)
put(ws3, f"F{row}", f"=ROUND(INDEX(G{T['buTot']}:I{T['buTot']},x_optIdx)-SUM(E{T['cF']}:K{T['cL4']})-SUM(E{T['one']}:K{T['one']}),6)", SEC, 9, fmt=EUR3)
T["chk"] = row; row += 2

# ---- 3.5 cash flow & NPV ----
section(ws3, row, "3.5 Cash flow and NPV, selected option (€m)", last_col="N",
        note="Free cash flow after tax = EBITDA impact − tax on EBIT impact − investment + working-capital release. Discounted at WACC, end of year."); row += 2
hdr_years(ws3, row, "Cash-flow line", ""); row += 1
C = {}
for key, lab, f, kw in [
    ("ben", "Gross benefits", lambda c, p, k: wsum("ben", c), {}),
    ("opex", "Run cost incl. one-off expensed", lambda c, p, k: "=-(" + wsum("opex", c)[1:] + ")", {}),
    ("ebitda", "EBITDA impact", lambda c, p, k: f"={c}{row_ben}+{c}{row_opex}", {"bold": True, "top": True}),
    ("da", "Depreciation", lambda c, p, k: "=-(" + wsum("da", c)[1:] + ")", {"color": SEC}),
    ("tax", "Tax on EBIT impact", lambda c, p, k: f"=-x_tax*({c}{row_ebitda}+{c}{row_da})", {}),
    ("capex", "Investment (capex)", lambda c, p, k: "=-(" + wsum("capex", c)[1:] + ")", {}),
    ("wc", "Working-capital release (spare parts)", lambda c, p, k: wsum("wc", c), {}),
]:
    row_ben = C.get("ben"); row_opex = C.get("opex"); row_ebitda = C.get("ebitda"); row_da = C.get("da")
    yrow(ws3, row, lab, f, **kw); C[key] = row; row += 1
C["fcf"] = row
yrow(ws3, row, "Free cash flow after tax", lambda c, p, k: f"={c}{C['ebitda']}+{c}{C['tax']}+{c}{C['capex']}+{c}{C['wc']}", bold=True, top=True); row += 1
C["cum"] = row; yrow(ws3, row, "Cumulative free cash flow", lambda c, p, k: f"={c}{C['fcf']}" + (f"+{p}{C['cum']}" if p else ""), tot3=False, tot5=False); row += 1
C["pv"] = row; yrow(ws3, row, "Present value", lambda c, p, k: f"={c}{C['fcf']}*{c}{R['df']}"); row += 1
C["cpv"] = row; yrow(ws3, row, "Cumulative present value (= NPV to date)", lambda c, p, k: f"={c}{C['pv']}" + (f"+{p}{C['cpv']}" if p else ""), tot3=False, tot5=False, bold=True); row += 1
C["pb"] = row
yrow(ws3, row, "Payback helper", lambda c, p, k: (f'=IF(AND({c}{C["cum"]}>=0,{p}{C["cum"]}<0),{k}+(-{p}{C["cum"]})/{c}{C["fcf"]},"")' if p else f'=IF({c}{C["cum"]}>=0,0,"")'),
     "yrs", '0.0', tot3=False, tot5=False, color=SEC); row += 1
C["chk"] = row
yrow(ws3, row, "Check: FCF equals engine (should be 0)", lambda c, p, k: f"=ROUND({c}{C['fcf']}-({wsum('fcf', c)[1:]}),6)", fmt=EUR3, color=SEC, tot3=False, tot5=False); row += 2

NPVF = lambda f_r, n: f"=SUMPRODUCT($E${f_r}:$K${f_r},$E${R['df']}:$K${R['df']},--($E${R['t']}:$K${R['t']}<={n}))"
KP = [("NPV 3 years", NPVF(C["fcf"], 3), EUR, "x_NPV3"), ("NPV 5 years", NPVF(C["fcf"], 5), EUR, "x_NPV5"),
      ("NPV 7 years", NPVF(C["fcf"], 7), EUR, "x_NPV7"),
      ("IRR 5 years", f'=IFERROR(IRR(E{C["fcf"]}:I{C["fcf"]}),"n/a")', PCT, "x_IRR5"),
      ("IRR 7 years", f'=IFERROR(IRR(E{C["fcf"]}:K{C["fcf"]}),"n/a")', PCT, "x_IRR7"),
      ("Payback (undiscounted)", f'=IF(COUNT(E{C["pb"]}:K{C["pb"]})=0,"> 7 yrs",MIN(E{C["pb"]}:K{C["pb"]}))', YRS, "x_PB"),
      ("Peak funding need (lowest cumulative FCF)", f"=MIN(E{C['cum']}:K{C['cum']})", EUR, "x_peak"),
      ("TCO 3 years", f"=L{T['tco']}", EUR, "x_TCO3"), ("TCO 5 years", f"=M{T['tco']}", EUR, "x_TCO5"),
      ("Run-rate EBITDA impact (2031)", f"=I{P['xebitda']}" if False else f"=I{C['ebitda']}", EUR, "x_RRebitda"),
      ("One-off implementation cost", f"=INDEX(G{T['buTot']}:I{T['buTot']},x_optIdx)", EUR, "x_oneoff")]
header(ws3, row, ["Result", "Unit", "Value"]); row += 1
for (lab, f, fmt, nm) in KP:
    put(ws3, f"B{row}", lab); put(ws3, f"C{row}", "€m" if fmt == EUR else ("%" if fmt == PCT else "yrs"), SEC, 9)
    put(ws3, f"D{row}", f, INK, bold=True, fmt=fmt); name(nm, S3, f"D{row}"); row += 1
row += 1

# ---- 3.6 options comparison ----
section(ws3, row, "3.6 Options compared: NPV, TCO and payback (selected scenario)", last_col="N",
        note="Same foundation in every option. The more use cases run on it, the better the return: the financial case for one group standard."); row += 2
hdr_years(ws3, row, "Option engine", ""); row += 1
O = {}
for o in "ABC":
    fr = OPT_ROWS[o]
    W = {"F": "1", "L1": f"{q(S2)}!$E${fr}", "L2": f"{q(S2)}!$F${fr}", "L3": f"{q(S2)}!$G${fr}", "L4": f"{q(S2)}!$H${fr}",
         "X": f"{q(S2)}!$E${fr}*{q(S2)}!$F${fr}"}
    O[(o, "fcf")] = row; yrow(ws3, row, f"Option {o} – free cash flow", lambda c, p, k, W=W: wsum("fcf", c, W), bold=True); row += 1
    O[(o, "ebitda")] = row; yrow(ws3, row, f"Option {o} – EBITDA impact", lambda c, p, k, W=W: "=" + wsum("ben", c, W)[1:] + "-(" + wsum("opex", c, W)[1:] + ")"); row += 1
    O[(o, "tco")] = row; yrow(ws3, row, f"Option {o} – TCO", lambda c, p, k, W=W: "=" + wsum("capex", c, W)[1:] + "+" + wsum("opex", c, W)[1:]); row += 1
    O[(o, "cum")] = row; yrow(ws3, row, f"Option {o} – cumulative FCF", lambda c, p, k, o=o: f"={c}{O[(o,'fcf')]}" + (f"+{p}{O[(o,'cum')]}" if p else ""), tot3=False, tot5=False, color=SEC); row += 1
    O[(o, "pb")] = row
    yrow(ws3, row, f"Option {o} – payback helper", lambda c, p, k, o=o: (f'=IF(AND({c}{O[(o,"cum")]}>=0,{p}{O[(o,"cum")]}<0),{k}+(-{p}{O[(o,"cum")]})/{c}{O[(o,"fcf")]},"")' if p else f'=IF({c}{O[(o,"cum")]}>=0,0,"")'),
         "yrs", "0.0", tot3=False, tot5=False, color=SEC); row += 1
row += 1
header(ws3, row, ["Comparison", "Unit", "Option A", "Option B", "Option C", "", "", "", "", "", "", "", "Reading"]); row += 1
OPT_CMP = row
for (lab, unit, fx, fmt) in [
    ("Description", "", lambda o: f"={q(S2)}!C{OPT_ROWS[o]}", None),
    ("One-off implementation cost", "€m", lambda o: f"={'GHI'['ABC'.index(o)]}{T['buTot']}", EUR),
    ("TCO 3 years", "€m", lambda o: f"=L{O[(o,'tco')]}", EUR),
    ("TCO 5 years", "€m", lambda o: f"=M{O[(o,'tco')]}", EUR),
    ("Run-rate EBITDA impact (2031)", "€m", lambda o: f"=I{O[(o,'ebitda')]}", EUR),
    ("NPV 3 years", "€m", lambda o: NPVF(O[(o, 'fcf')], 3), EUR),
    ("NPV 5 years", "€m", lambda o: NPVF(O[(o, 'fcf')], 5), EUR),
    ("NPV 7 years", "€m", lambda o: NPVF(O[(o, 'fcf')], 7), EUR),
    ("IRR 7 years", "%", lambda o: f'=IFERROR(IRR(E{O[(o,"fcf")]}:K{O[(o,"fcf")]}),"n/a")', PCT),
    ("Payback (undiscounted)", "yrs", lambda o: f'=IF(COUNT(E{O[(o,"pb")]}:K{O[(o,"pb")]})=0,"> 7 yrs",MIN(E{O[(o,"pb")]}:K{O[(o,"pb")]}))', YRS),
    ("Run-rate EBITDA per € one-off cost", "x", lambda o: f"=IFERROR({'DEF'['ABC'.index(o)]}{OPT_CMP+4}/{'DEF'['ABC'.index(o)]}{OPT_CMP+1},0)", '0.00"x"'),
]:
    put(ws3, f"B{row}", lab, INK, bold=(lab.startswith("NPV 5"))); put(ws3, f"C{row}", unit, SEC, 9)
    for o, col in zip("ABC", "DEF"):
        put(ws3, f"{col}{row}", fx(o), INK, bold=(lab.startswith("NPV")), fmt=fmt, wrap=(fmt is None), align="right" if fmt else None)
    for col in "GHIJKLM": put(ws3, f"{col}{row}", None)
    put(ws3, f"N{row}", "", SEC, 9)
    if lab == "Description": ws3.row_dimensions[row].height = 40
    R["cmp_" + lab] = row; row += 1
put(ws3, f"N{R['cmp_NPV 5 years']}", "Option A carries the whole foundation on one lever; C spreads it over four.", SEC, 9)
row += 1
assert row < ENGINE_START, f"sections end at {row}, engine starts {ENGINE_START}"
print("sections end", row, "engine start", ENGINE_START)

# ---- write engine ----
row = ENGINE_START
section(ws3, row, "3.7 Engine by building block (feeds §3.2–3.6)", last_col="N",
        note="Each block has its own benefits, run cost, investment, depreciation and free cash flow. Options are weighted sums of blocks. Tax = rate × change in EBIT."); row += 2
hdr_years(ws3, row, "Block / line", "Weight (sel.)")
avg, ramp, phF, phD, nS, nL = R["avg"], R["ramp"], R["phF"], R["phD"], R["newStd"], R["newLeg"]
for b in BLOCKS:
    cap_r = E[(b, "capex")]
    upk = lambda c, p: (f"+x_upkeep*SUM($E${cap_r}:{p}{cap_r})" if p else "")
    put(ws3, f"B{E[(b,'w')]}", BNAME[b], INK, 11, True, border=False)
    put(ws3, f"C{E[(b,'w')]}", "weight", SEC, 9, border=False)
    put(ws3, f"D{E[(b,'w')]}", WSEL[b], INK, bold=True, fmt="0", border=False)
    ben = {"F": lambda c, p, k: "=0",
           "L1": lambda c, p, k: f"={c}{rl}*(x_rr_capL1+x_rr_maint+x_rr_fp)",
           "L2": lambda c, p, k: f"={c}{rl}*x_rr_capL2",
           "L3": lambda c, p, k: f"={c}{rl}*x_rr_scrap",
           "L4": lambda c, p, k: f"={c}{rl}*x_rr_energy",
           "X": lambda c, p, k: f"={c}{rl}*x_rr_X"}[b]
    yrow(ws3, E[(b, "ben")], "  Benefits (EBITDA)", ben)
    opx = {"F": lambda c, p, k: f"=(x_rPlat+x_fteF*x_rFTE+x_rSOC+x_rChamp*x_plants+x_rTrain)*{c}{ramp}+{c}{E[('F','oneoff')]}" + upk(c, p),
           "L1": lambda c, p, k: f"=x_fteL1*x_rFTE+x_rL1*{c}{avg}" + upk(c, p),
           "L2": lambda c, p, k: f"=x_fteL2*x_rFTE+x_rL2*{c}{avg}" + upk(c, p),
           "L3": lambda c, p, k: f"=x_fteL3*x_rFTE+x_rL3*x_nL3*{c}{avg}/x_lines" + upk(c, p),
           "L4": lambda c, p, k: f"=x_fteL4*x_rFTE+x_rL4*x_plants*{c}{avg}/x_lines" + upk(c, p),
           "X": lambda c, p, k: "=0"}[b]
    yrow(ws3, E[(b, "opex")], "  Run cost incl. one-off expensed (EBITDA)", opx)
    if b == "F":
        yrow(ws3, E[("F", "oneoff")], "    of which one-off expensed (programme, training)", lambda c, p, k: f"=(x_cPMO+x_cTrain)*(1+x_cont)*{c}{phF}", color=SEC)
    cpx = {"F": lambda c, p, k: f"=(x_cPlat+x_cOTseg*x_plants+x_cOTsec+x_cInteg)*(1+x_cont)*{c}{phF}",
           "L1": lambda c, p, k: f"=(x_cL1std*{c}{nS}+x_cL1leg*{c}{nL}+x_cL1dev*{c}{phD})*(1+x_cont)",
           "L2": lambda c, p, k: f"=(x_cL2std*{c}{nS}+x_cL2leg*{c}{nL}+x_cL2dev*{c}{phD})*(1+x_cont)",
           "L3": lambda c, p, k: f"=(x_cL3st*x_nL3*({c}{nS}+{c}{nL})/x_lines+x_cL3dev*{c}{phD})*(1+x_cont)",
           "L4": lambda c, p, k: f"=x_cL4pl*x_plants*({c}{nS}+{c}{nL})/x_lines*(1+x_cont)",
           "X": lambda c, p, k: "=0"}[b]
    yrow(ws3, cap_r, "  Investment (capex, capitalised)", cpx)
    yrow(ws3, E[(b, "da")], "  Depreciation", lambda c, p, k: f"=SUMPRODUCT(($E${cap_r}:{c}{cap_r})*(({c}${t}-$E${t}:{c}${t})<x_life))/x_life")
    wc = {"L1": lambda c, p, k: f"=x_spares*x_redSpares*({c}{rl}" + (f"-{p}{rl})" if p else ")")}.get(b, lambda c, p, k: "=0")
    yrow(ws3, E[(b, "wc")], "  Working-capital release (spare parts)", wc)
    bb, oo, dd, cc, ww = (E[(b, k)] for k in ("ben", "opex", "da", "capex", "wc"))
    yrow(ws3, E[(b, "fcf")], "  Free cash flow after tax", lambda c, p, k: f"={c}{bb}-{c}{oo}-x_tax*({c}{bb}-{c}{oo}-{c}{dd})-{c}{cc}+{c}{ww}", bold=True, top=True)

# ---- KPI header on tab 3 ----
r0 = KPI_ROW
LEFT = [("Option", "=x_opt&\" – \"&" + f"INDEX({q(S2)}!C{OPT_ROWS['A']}:C{OPT_ROWS['C']},x_optIdx)", None),
        ("Scenario", "=x_scen", None), ("One-off implementation cost", "=x_oneoff", EUR),
        ("TCO 3 / 5 years", '=TEXT(x_TCO3,"0.0")&" / "&TEXT(x_TCO5,"0.0")', None),
        ("Run-rate EBITDA impact (2031)", "=x_RRebitda", EUR),
        ("OEE before → after", '=TEXT(x_OEE0,"0.0%")&" → "&TEXT(x_OEE1,"0.0%")', None)]
RIGHT = [("NPV 3 years", "=x_NPV3", EUR), ("NPV 5 years", "=x_NPV5", EUR), ("NPV 7 years", "=x_NPV7", EUR),
         ("IRR 7 years", "=x_IRR7", PCT), ("Payback (undiscounted)", "=x_PB", YRS), ("Peak funding need", "=x_peak", EUR)]
for k, ((l1, f1, m1), (l2, f2, m2)) in enumerate(zip(LEFT, RIGHT)):
    r = r0 + k
    put(ws3, f"B{r}", l1, SEC); ws3.merge_cells(f"C{r}:G{r}")
    put(ws3, f"C{r}", f1, INK, 11, True, fmt=m1, align="left")
    ws3.merge_cells(f"H{r}:J{r}"); put(ws3, f"H{r}", l2, SEC)
    ws3.merge_cells(f"K{r}:L{r}"); put(ws3, f"K{r}", f2, BLUE_DK if "NPV" in l2 else INK, 11, True, fmt=m2, align="right")
    for col in "DEFGIJLM": ws3[f"{col}{r}"].border = Border(bottom=thin)
put(ws3, f"B{r0+6}", "€m unless stated. Change selectors on tab 2. Checks on this tab should all read 0.", SEC, 9, italic=True, border=False)
ws3.freeze_panes = "C6"

# =====================================================================================
# TAB 4 – SENSITIVITY & STRESS TESTS
# =====================================================================================
widths(ws4, {"A": 2, "B": 44, "C": 12, "D": 12, "E": 12, "F": 12, "G": 12, "H": 12, "I": 12, "J": 4, "K": 70})
title(ws4, "Tab 4 · Sensitivity analysis and stress tests",
      "Where the case breaks, and what still holds",
      "Four sensitivity grids plus four combined stress scenarios around the selected option and scenario. Green = NPV ≥ 0, red = NPV < 0. €m.")
S3R = lambda col, r: f"{q(S3)}!{col}{r}"
rngS3 = lambda r, a="E", b="K": f"{q(S3)}!${a}${r}:${b}${r}"
row = 6
section(ws4, row, "Building blocks of the NPV (selected option)", last_col="K",
        note="The NPV is linear in benefits, run cost and investment, so each grid below is exact, not an approximation."); row += 2
header(ws4, row, ["Component", "5 years", "7 years", "", "", "", "", "", "", "What it is"]); row += 1
# helper rows: t, df, R, R shifted
H = {}
for key, lab, fs in [("t", "Year index", lambda k, c: f"={S3R(c, R['t'])}"),
                     ("df", "Discount factor", lambda k, c: f"={S3R(c, R['df'])}"),
                     ("r0", "Benefit realisation", lambda k, c: f"={S3R(c, R['real'])}"),
                     ("r1", "Benefit realisation delayed 1 year", lambda k, c: "=0" if k == 0 else f"={S3R(YC[k-1], R['real'])}")]:
    H[key] = row
    put(ws4, f"B{row}", lab, SEC, 9)
    for k, c in enumerate("CDEFGHI"):
        put(ws4, f"{c}{row}", fs(k, YC[k]), GREEN_LINK, 9, fmt="0.000")
    put(ws4, f"K{row}", "Helper row, years 2027–2033 in C:I", SEC, 9); row += 1
hr = lambda key: f"$C${H[key]}:$I${H[key]}"
mask = lambda n: f"--({hr('t')}<={n})"
B = {}
COMP = [("PVB", "PV of benefits after tax", lambda n: f"=SUMPRODUCT({rngS3(C['ben'])}*(1-x_tax),{hr('df')},{mask(n)})", "Σ gross benefit × (1 − tax) × DF"),
        ("PVO", "PV of run cost after tax", lambda n: f"=SUMPRODUCT(-{rngS3(C['opex'])}*(1-x_tax),{hr('df')},{mask(n)})", "Σ run cost incl. one-off expensed × (1 − tax) × DF"),
        ("PVC", "PV of investment net of depreciation tax shield", lambda n: f"=SUMPRODUCT((-{rngS3(C['capex'])}+x_tax*{rngS3(C['da'])}),{hr('df')},{mask(n)})", "Σ (capex − tax × depreciation) × DF"),
        ("PVW", "PV of working-capital release", lambda n: f"=SUMPRODUCT({rngS3(C['wc'])},{hr('df')},{mask(n)})", ""),
        ("NPV", "NPV = PVB − PVO − PVC + PVW", lambda n: f"=C{row+0}-C{row+1}-C{row+2}+C{row+3}", ""),
        ("chk", "Check vs tab 3 (should be 0)", None, ""),
        ("PVF0", "PV factor of €1 m run-rate benefit, no delay", lambda n: f"=SUMPRODUCT({hr('r0')},{hr('df')},{mask(n)})", "Σ realisation × DF"),
        ("PVF1", "PV factor of €1 m run-rate benefit, 1-year delay", lambda n: f"=SUMPRODUCT({hr('r1')},{hr('df')},{mask(n)})", "")]
for (key, lab, fn, note) in COMP:
    B[key] = row; row += 1
for (key, lab, fn, note) in COMP:
    r = B[key]
    put(ws4, f"B{r}", lab, INK, bold=(key == "NPV"))
    for col, n in (("C", 5), ("D", 7)):
        if key == "NPV":
            f = f"={col}{B['PVB']}-{col}{B['PVO']}-{col}{B['PVC']}+{col}{B['PVW']}"
        elif key == "chk":
            f = f"=ROUND({col}{B['NPV']}-" + ("x_NPV5" if n == 5 else "x_NPV7") + ",6)"
        else:
            f = fn(n)
        put(ws4, f"{col}{r}", f, INK, bold=(key == "NPV"), fmt=("0.000" if key.startswith("PVF") else (EUR3 if key == "chk" else EUR)))
    for col in "EFGHIJ": put(ws4, f"{col}{r}", None)
    put(ws4, f"K{r}", note, SEC, 9)
r = row
for (key, lab, f, note) in [("RR", "Run-rate gross benefit", "=x_RR", "From tab 3 §3.1"),
                            ("RRcap", "of which capacity value", "=x_capSel*x_Vcap", "Capacity gain × value of capacity"),
                            ("RRoth", "of which cost savings (maintenance, scrap, energy, false alerts)", f"=C{r}-C{r+1}", ""),
                            ("pts", "OEE points from availability + performance", "=x_capSel*x_OEE0*100", "Capacity gain × baseline OEE, in OEE points"),
                            ("chkL", "Check: PVB = RR × (1 − tax) × PVF0 (should be 0)", None, "Confirms the benefits follow one realisation curve")]:
    B[key] = row
    put(ws4, f"B{row}", lab)
    if key == "chkL":
        put(ws4, f"C{row}", f"=ROUND(C{B['PVB']}-C{B['RR']}*(1-x_tax)*C{B['PVF0']},6)", SEC, fmt=EUR3)
    else:
        put(ws4, f"C{row}", f, INK, fmt=(PTN if key == "pts" else EUR))
    for col in "DEFGHIJ": put(ws4, f"{col}{row}", None)
    put(ws4, f"K{row}", note, SEC, 9); row += 1
row += 1
VCAP = lambda s: f"x_rev0*({s}*x_CM+(1-{s})*x_pers*x_flex)"
cf_rules = []


def grid(r0, c0, rows_vals, cols_vals, rfmt, cfmt, formula, rlab, clab):
    """formula(rcell, ccell) -> excel formula (no '=')."""
    cols = "CDEFGHI"
    put(ws4, f"B{r0}", f"{rlab} ↓   {clab} →", SEC, 9, True, fill=HDR)
    for j, v in enumerate(cols_vals):
        put(ws4, f"{cols[j]}{r0}", v, BLUE, 10, True, fmt=cfmt, fill=HDR, align="right")
    for i, rv in enumerate(rows_vals):
        rr_ = r0 + 1 + i
        put(ws4, f"B{rr_}", rv, BLUE, 10, True, fmt=rfmt, align="right")
        for j in range(len(cols_vals)):
            put(ws4, f"{cols[j]}{rr_}", "=" + formula(f"$B{rr_}", f"{cols[j]}${r0}"), INK, fmt=EUR, align="right")
    rng = f"C{r0+1}:{cols[len(cols_vals)-1]}{r0+len(rows_vals)}"
    ws4.conditional_formatting.add(rng, CellIsRule(operator="greaterThanOrEqual", formula=["0"], fill=PatternFill("solid", fgColor=OK_BG), font=Font(name=FONT, color="0E6027")))
    ws4.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor=ERR_BG), font=Font(name=FONT, color="A2191F")))
    return r0 + len(rows_vals) + 1


# S1
section(ws4, row, "S1 · Demand and OEE uplift → 5-year NPV", last_col="K",
        note="Can the extra output be sold, and how many OEE points do availability (L1) and performance (L2) really add? Cost savings unchanged."); row += 2
put(ws4, f"K{row}", "Sellable share is the biggest swing factor: unsold capacity is only worth the flexible labour it frees. Blue row and column values are inputs: edit them to test other points.", SEC, 9, wrap=True)
row = grid(row, "C", [0, 0.25, 0.5, 0.75, 1.0], [1, 2, 3, 4, 6, 8, 12], "0%", '"+"0" pt"',
           lambda rc, cc: f"x_NPV5+(({cc}/100)/x_OEE0*{VCAP(rc)}-$C${B['RRcap']})*(1-x_tax)*$C${B['PVF0']}",
           "Sellable share", "OEE points")
put(ws4, f"B{row}", '="Model point: "&TEXT(x_sell,"0%")&" sellable, +"&TEXT(C' + str(B['pts']) + ',"0.0")&" OEE pts. The concept\'s 15-25% capacity is about +10 to +16 pts."', SEC, 9, italic=True, border=False)
row += 2
# S2
section(ws4, row, "S2 · Cost overrun stress → 5-year NPV", last_col="K",
        note="Investment (capex) overrun down the side, run-cost overrun (incl. one-off expensed) across. Contingency of 15% is already inside the base."); row += 2
row = grid(row, "C", [0, 0.1, 0.25, 0.5, 0.75, 1.0], [0, 0.15, 0.3, 0.5], "+0%", '+0%',
           lambda rc, cc: f"x_NPV5-{rc}*$C${B['PVC']}-{cc}*$C${B['PVO']}", "Capex overrun", "Run-cost overrun")
row += 1
# S3
section(ws4, row, "S3 · Benefit capture and roll-out delay → NPV", last_col="K",
        note="What if adoption is slower or plants block go-live? Capture = share of planned benefits realised; delay shifts all benefits one year (costs unchanged)."); row += 2
put(ws4, f"B{row}", "Benefit capture ↓   Horizon / delay →", SEC, 9, True, fill=HDR)
S3COLS = [("5 yrs, on time", 5, 0), ("5 yrs, 1-yr delay", 5, 1), ("7 yrs, on time", 7, 0), ("7 yrs, 1-yr delay", 7, 1)]
for j, (lab, n, d) in enumerate(S3COLS):
    put(ws4, f"{'CDEF'[j]}{row}", lab, INK, 9, True, fill=HDR, wrap=True, align="right")
ws4.row_dimensions[row].height = 28
g0 = row
for i, k in enumerate([0.5, 0.6, 0.7, 0.8, 0.9, 1.0, 1.1, 1.2]):
    rr_ = g0 + 1 + i
    put(ws4, f"B{rr_}", k, BLUE, 10, True, fmt=PCT0, align="right")
    for j, (lab, n, d) in enumerate(S3COLS):
        col = "C" if n == 5 else "D"
        npv = "x_NPV5" if n == 5 else "x_NPV7"
        pvf = f"${col}${B['PVF0'] if d == 0 else B['PVF1']}"
        put(ws4, f"{'CDEF'[j]}{rr_}", f"={npv}-${col}${B['PVB']}+$B{rr_}*$C${B['RR']}*(1-x_tax)*{pvf}", INK, fmt=EUR, align="right")
rng = f"C{g0+1}:F{g0+8}"
ws4.conditional_formatting.add(rng, CellIsRule(operator="greaterThanOrEqual", formula=["0"], fill=PatternFill("solid", fgColor=OK_BG), font=Font(name=FONT, color="0E6027")))
ws4.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor=ERR_BG), font=Font(name=FONT, color="A2191F")))
row = g0 + 10
# S4
section(ws4, row, "S4 · Discount rate and horizon → NPV", last_col="K",
        note="A critical CFO may ask for a higher hurdle rate or a shorter horizon. Exact recalculation of the selected free cash flow."); row += 2
row = grid(row, "C", [0.06, 0.07, 0.08, 0.09, 0.10, 0.12], [3, 5, 7], PCT0, '0" yrs"',
           lambda rc, cc: f"SUMPRODUCT({rngS3(C['fcf'])}/(1+{rc})^{hr('t')}*({hr('t')}<={cc}))", "Discount rate", "Horizon")
row += 1

# Stress tests
section(ws4, row, "Stress tests: four combined scenarios", last_col="K",
        note="Several things go wrong (or right) at once. Inputs in blue can be changed. Base column follows the model."); row += 2
header(ws4, row, ["Driver", "Base (model)", "CFO downside", "Severe stress", "Upside", "", "", "", "", "Story"]); row += 1
ST = {}
STRESS = [
    ("s", "Sellable share of extra capacity", ["=x_sell", 0.25, 0.10, 0.75], PCT0,
     "CFO: market flat, little extra sold. Severe: heat-pump demand dip, almost nothing sold. Upside: energy-transition demand."),
    ("pts", "OEE points from availability + performance", [f"=C{B['pts']}", f"=C{B['pts']}", 2.0, 6.0], PTN,
     "Severe: only part of the PdM/OEE effect materialises (legacy lines, data quality)."),
    ("k", "Benefit capture", [1, 0.9, 0.7, 1.0], PCT0, "Adoption: operators must trust and act on AI alerts."),
    ("d", "Roll-out delay (0 or 1 year)", [0, 0, 1, 0], "0", "Severe: an OT incident or plant go/no-go stops the roll-out for a year."),
    ("kc", "Investment overrun", [0, 0.25, 0.5, 0], PCT0, "Legacy retrofit and OT segmentation are the usual overruns."),
    ("ko", "Run-cost overrun", [0, 0.15, 0.25, 0], PCT0, "Cloud consumption, licences, scarce data/OT talent."),
]
for (key, lab, vals, fmt, story) in STRESS:
    ST[key] = row
    put(ws4, f"B{row}", lab)
    for j, v in enumerate(vals):
        col = "CDEF"[j]
        if j == 0:
            put(ws4, f"{col}{row}", v, INK, fmt=fmt, align="right")
        else:
            put(ws4, f"{col}{row}", v, BLUE, fmt=fmt, fill=BLUE_BG, align="right")
    for col in "GHIJ": put(ws4, f"{col}{row}", None)
    put(ws4, f"K{row}", story, SEC, 9, wrap=True); row += 1
dvd = DataValidation(type="list", formula1='"0,1"'); ws4.add_data_validation(dvd); dvd.add(f"D{ST['d']}:F{ST['d']}")
OUTS = {}
for key, lab, fmt in [("rr", "Run-rate gross benefit", EUR), ("npv5", "NPV 5 years", EUR), ("npv7", "NPV 7 years", EUR), ("v", "Verdict", None)]:
    OUTS[key] = row
    put(ws4, f"B{row}", lab, INK, bold=True)
    for j in range(4):
        col = "CDEF"[j]
        s_, p_, k_, d_, kc_, ko_ = (f"{col}{ST[x]}" for x in ("s", "pts", "k", "d", "kc", "ko"))
        if key == "rr":
            f = f"={k_}*($C${B['RRoth']}+{p_}/100/x_OEE0*{VCAP(s_)})"
        elif key in ("npv5", "npv7"):
            n = "C" if key == "npv5" else "D"
            f = (f"={col}{OUTS['rr']}*(1-x_tax)*IF({d_}=1,${n}${B['PVF1']},${n}${B['PVF0']})"
                 f"-(1+{ko_})*${n}${B['PVO']}-(1+{kc_})*${n}${B['PVC']}+${n}${B['PVW']}")
        else:
            f = f'=IF({col}{OUTS["npv5"]}>=0,"Pays back within 5 yrs",IF({col}{OUTS["npv7"]}>=0,"Positive within 7 yrs","Negative: do not scale"))'
        put(ws4, f"{col}{row}", f, INK, bold=True, fmt=fmt, align="right", wrap=(fmt is None))
    for col in "GHIJK": put(ws4, f"{col}{row}", None)
    for col in "BCDEFGHIJK": ws4[f"{col}{row}"].border = Border(top=strong if key == "rr" else thin, bottom=thin)
    row += 1
ws4.row_dimensions[OUTS["v"]].height = 28
put(ws4, f"K{OUTS['rr']}", "Check: Base column equals the model's 5- and 7-year NPV.", SEC, 9)
put(ws4, f"G{OUTS['npv5']}", f"=ROUND(C{OUTS['npv5']}-x_NPV5,6)", SEC, 9, fmt=EUR3)
for key in ("npv5", "npv7"):
    rng = f"C{OUTS[key]}:F{OUTS[key]}"
    ws4.conditional_formatting.add(rng, CellIsRule(operator="greaterThanOrEqual", formula=["0"], fill=PatternFill("solid", fgColor=OK_BG), font=Font(name=FONT, bold=True, color="0E6027")))
    ws4.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor=ERR_BG), font=Font(name=FONT, bold=True, color="A2191F")))
name("x_stressCFO5", S4, f"D{OUTS['npv5']}"); name("x_stressCFO7", S4, f"D{OUTS['npv7']}")
name("x_stressSev7", S4, f"E{OUTS['npv7']}"); name("x_stressUp5", S4, f"F{OUTS['npv5']}")
row += 1

section(ws4, row, "Break-even: how far can it fall before NPV = 0?", last_col="K"); row += 2
header(ws4, row, ["Break-even point", "5 years", "7 years", "", "", "", "", "", "", "Reading"]); row += 1
BE = [("Minimum benefit capture", lambda n: f"=IFERROR(({n}{B['PVO']}+{n}{B['PVC']}-{n}{B['PVW']})/{n}{B['PVB']},0)", PCT0,
       "Below this share of planned benefits the programme destroys value"),
      ("Maximum investment overrun", lambda n: f"=IFERROR({'x_NPV5' if n=='C' else 'x_NPV7'}/{n}{B['PVC']},0)", PCT0,
       "Negative = already below zero; positive = headroom before NPV turns negative"),
      ("Minimum sellable share of extra capacity",
       lambda n: (f"=MAX(0,((x_sell*x_CM+(1-x_sell)*x_pers*x_flex)-{'x_NPV5' if n=='C' else 'x_NPV7'}/(x_rev0*MAX(x_capSel,0.000001)*(1-x_tax)*{n}{B['PVF0']})"
                  f"-x_pers*x_flex)/(x_CM-x_pers*x_flex))"), PCT0,
       "0% = NPV stays positive even if no extra output is sold; above 100% = cannot break even on demand alone")]
for (lab, fn, fmt, note) in BE:
    put(ws4, f"B{row}", lab)
    put(ws4, f"C{row}", fn("C"), INK, bold=True, fmt=fmt); put(ws4, f"D{row}", fn("D"), INK, bold=True, fmt=fmt)
    for col in "EFGHIJ": put(ws4, f"{col}{row}", None)
    put(ws4, f"K{row}", note, SEC, 9); row += 1
name("x_beK5", S4, f"C{row-3}"); name("x_beK7", S4, f"D{row-3}"); name("x_beS7", S4, f"D{row-1}")
ws4.freeze_panes = "C6"

# =====================================================================================
# TAB 2 §H – calibration table (needs tab-3 names)
# =====================================================================================
row = CAL_START
CAL = [("Capacity gain (run-rate)", "%", 0.20, "=x_capSel", PCT,
        "Concept: 15–25% (mid 20%). Calibrated from industry PdM/OEE ranges (tab 2 §C)."),
       ("OEE needed / reached", "%", "=x_OEE0*1.2", "=x_OEE0*(1+x_capSel)", PCT,
        "From ≈ 65% the concept needs ≈ 78%, close to world class (85%). Calibrated stays inside the typical range."),
       ("Extra revenue per year (run-rate)", "€m", 300, "=x_rev0*x_capSel*x_sell", EUR,
        "Concept mid: €400 m × 75% sellable = €300 m. Calibrated uses the sellable share on tab 2."),
       ("Implied extra growth on top of the market", "%", 0.15, "=x_capSel*x_sell", PCT,
        "Peers grow ≈ 3% a year (P1, P2). The concept implies 5 'Ariston years' of growth at once."),
       ("EBITDA uplift per year (run-rate, after run cost)", "€m", 142.3, "=x_RRebitda", EUR,
        "Concept value from the calibration workbook (CAL, before programme cost). Calibrated is net of run cost."),
       ("Uplift vs peer margin gap (median → best ≈ €35 m)", "x", "=E{r}/35.3", "=F{r}/35.3", '0.0"x"',
        "Above 1 means claiming more than the whole margin gap between Ariston and BDR Thermea; needs strong pilot proof."),
       ("One-off implementation cost", "€m", "not quantified", "=x_oneoff", EUR,
        "Concept named the cost items but no amounts. Calibrated: suggested end-to-end costs (tab 3 §3.4).")]
for k, (lab, unit, cv, mv, fmt, rd) in enumerate(CAL):
    rr_ = row + k
    if isinstance(cv, str): cv = cv.replace("{r}", str(rr_ - 1))
    if isinstance(mv, str): mv = mv.replace("{r}", str(rr_ - 1))
    put(ws2, f"B{rr_}", None); put(ws2, f"C{rr_}", lab); put(ws2, f"D{rr_}", unit, SEC, 9)
    put(ws2, f"E{rr_}", cv, BLUE if not (isinstance(cv, str) and cv.startswith("=")) else INK, fmt=fmt, align="right")
    put(ws2, f"F{rr_}", mv, INK, bold=True, fmt=fmt, align="right")
    for col in "GHIJ": put(ws2, f"{col}{rr_}", None)
    put(ws2, f"K{rr_}", rd, SEC, 9, wrap=True)

# =====================================================================================
# TAB 1 – EXPLAINER
# =====================================================================================
widths(ws1, {"A": 2, "B": 34, "C": 120})
title(ws1, "Buderus · Group 3 · Business case v3 · November 2026",
      "OEE, predictive maintenance and AI & IoT: calibrated business case",
      "A qualitative and quantitative business case for the board. Fictional company; public peer data used only as a plausibility check.")
row = 6


def eu(x):
    return f'IF({x}<0,"−€","€")&TEXT(ABS({x}),"0.0")&" m"'


def para(lab, text, bold=False, color=INK, h=None, fmt=None):
    global row
    put(ws1, f"B{row}", lab, SEC if not bold else INK, 10, bold)
    put(ws1, f"C{row}", text, color, 10, bold, wrap=True, fmt=fmt)
    if h: ws1.row_dimensions[row].height = h
    row += 1


def head(text):
    global row
    row += 1
    section(ws1, row, text, last_col="C"); row += 1


head("Headline numbers (live from the model)")
para("Option and scenario", '=x_opt&" – "&' + f"INDEX({q(S2)}!C{OPT_ROWS['A']}:C{OPT_ROWS['C']},x_optIdx)" + '&" · scenario "&x_scen', True)
para("Operational effect", '="OEE "&TEXT(x_OEE0,"0.0%")&" → "&TEXT(x_OEE1,"0.0%")&" · capacity +"&TEXT(x_capSel,"0.0%")&" (concept assumed +15–25%)"')
para("Value per year (run-rate)", '="Gross benefit "&' + eu("x_RR") + '&" · EBITDA impact after run cost "&' + eu("x_RRebitda") + '&" (from 2031)"')
para("What it costs", '="One-off "&' + eu("x_oneoff") + '&" end-to-end · TCO "&' + eu("x_TCO3") + '&" over 3 years, "&' + eu("x_TCO5") + '&" over 5 years"')
para("What it is worth", '="NPV "&' + eu("x_NPV3") + '&" (3 yrs) · "&' + eu("x_NPV5") + '&" (5 yrs) · "&' + eu("x_NPV7") + '&" (7 yrs) · payback "&IF(ISNUMBER(x_PB),TEXT(x_PB,"0.0")&" yrs",x_PB)', True)
para("Stress test", '="CFO downside (25% sellable, +25% capex, +15% run cost, 90% capture): NPV "&' + eu("x_stressCFO5") + '&" (5 yrs), "&' + eu("x_stressCFO7") + '&" (7 yrs). Break-even needs "&TEXT(x_beK7,"0%")&" of planned benefits and "&TEXT(x_beS7,"0%")&" sellable share over 7 years."')

head("1. The board choice")
para("Central message", "Standardise the digital foundation at group level; plants apply it within local OT, safety and continuity limits; scale per plant and per use case only when value, technical fit and continuity are proven.", True, h=30)
para("What we ask", "Fund the foundation plus four AI & IoT use cases in three gated waves (2027–2029). Each wave is released only when the previous one meets agreed value, safety and continuity criteria.", h=30)

head("2. What changed compared with the group concept")
for lab, txt in [
    ("Capacity, not revenue", "The concept counted 15–25% extra production as €300–500 m extra revenue. That needs OEE to rise from ≈ 65% to ≈ 78–81%, close to world class, and growth of 4–6× what listed peers achieve. We calibrate the gain from industry ranges for each lever: a few percent of capacity, not 20%."),
    ("Margin, not turnover", "Only the sellable share of extra output becomes revenue, and only the contribution margin (≈ 47% with the peer cost structure) is value. Unsold capacity is valued as overtime and temporary labour avoided."),
    ("Costs made explicit", "The concept named the cost items but gave no amounts. This version suggests end-to-end costs: foundation, sensors and retrofit per line, AI vision, energy, training, change, run cost, contingency."),
    ("Peer calibration", "Cost structure, margin, growth and investment capacity come from Ariston, BDR Thermea and Bosch Home Comfort public figures (tab 2, references P1–P3)."),
    ("Three options", "A: predictive maintenance only. B: connected OEE + predictive maintenance. C: full AI & IoT. All share one foundation."),
]:
    para(lab, txt, h=42)

head("3. How AI and IoT create value in the business")
for lab, txt in [
    ("Foundation (enabler)", "One group data & AI platform, OT segmentation with edge gateways per plant, OT security and a common data model. No value on its own, but every use case reuses it. Control (PLC/SCADA) stays local; AI advises, people decide."),
    ("L1 Predictive maintenance", "IoT vibration and temperature sensors feed AI models that predict failures days ahead. Repairs move into planned windows: fewer stops (availability), lower maintenance cost, fewer emergency spares. Risk: false alerts, which the model counts as a cost."),
    ("L2 Connected OEE", "Line data from PLC/SCADA via the edge into MES and one OEE definition for all plants. AI ranks micro-stops and speed losses by cause; teams fix the top losses each week (performance)."),
    ("L3 AI vision quality", "Cameras with edge AI at quality-critical stations catch defects where they arise instead of at end-of-line or in the field: less scrap and rework, fewer warranty cases (not valued)."),
    ("L4 AI energy optimisation", "Sub-metering per line and AI scheduling cut idle consumption, compressed-air losses and peak load: lower energy cost and CO2."),
    ("Why scale matters", "The foundation is a fixed cost. Each extra line and use case adds value at a small marginal cost. Compare options on tab 3 §3.6: that is the financial argument for one group standard."),
]:
    para(lab, txt, h=42)

head("4. Value the numbers do not capture (qualitative)")
for lab, txt in [
    ("Safety and continuity", "Fewer unplanned breakdowns mean fewer hazardous emergency repairs. Local control keeps plants running if the link to Amsterdam fails."),
    ("Cybersecurity", "Segmentation, controlled remote access and OT monitoring lower the risk of an OT incident; today's ad-hoc connections are the bigger risk."),
    ("Quality and customers", "Fewer field failures of boilers and heat pumps protect the brand and lower warranty cost."),
    ("ESG", "Energy and scrap reduction lower CO2 and material use, which supports reporting duties."),
    ("Talent and scale", "Shared data and AI skills (academy, plant digital champions) let new plants and use cases go live faster and make the group a better employer for scarce OT/IT talent."),
]:
    para(lab, txt, h=30)

head("5. Risks and what the stress tests say")
for lab, txt in [
    ("Demand", "The biggest swing factor. If extra output cannot be sold, capacity is only worth the flexible labour it frees (tab 4, S1)."),
    ("Cost overrun", "Legacy retrofit and OT segmentation are the usual overruns; 15% contingency is in the base, S2 tests up to +100%."),
    ("Adoption and delay", "Value depends on operators acting on AI alerts and on plants allowing go-live. S3 tests capture and a one-year delay."),
    ("Hurdle rate and horizon", "Infrastructure pays back over 5–7 years, not 3. S4 shows NPV for 6–12% and 3/5/7 years."),
    ("Combined stress", "Four stress scenarios (base, CFO downside, severe, upside) and the break-even points are at the bottom of tab 4."),
]:
    para(lab, txt, h=30)

head("6. Gates and data still needed")
para("Gates", "Wave 1 (2027): pilot plant + two plants; prove OEE baseline, alert precision, adoption and OT safety. Wave 2 (2028): standard lines in all plants. Wave 3 (2029): legacy lines and remaining use cases. Plant manager keeps go/no-go on anything that touches safety or continuity.", h=42)
para("Pilot data to replace assumptions", "Measured OEE per line (A, P, Q), unplanned downtime and MTTR, maintenance spend, scrap cost, energy cost, useful / false / missed alerts, sellable share (Sales), real cost per plant type (OT/Engineering).", h=30)

head("How to use this workbook")
for lab, txt in [
    ("Tab 1 Explainer", "This page: the story and the live headline numbers."),
    ("Tab 2 Assumptions & Refs", "All inputs with Low/Base/High values, the label (case, group, source, assumption) and the industry reference. Two selectors at the top: scenario and option. Calibration table and full reference list at the bottom."),
    ("Tab 3 P&L TCO NPV", "Impact of each lever, benefit build-up, 3-year P&L (baseline, programme effect, pro-forma), TCO, cash flow and NPV, options compared, and the calculation engine."),
    ("Tab 4 Sensitivity & Stress", "Four sensitivity grids, four combined stress scenarios and break-even points."),
    ("Colours", "Blue text on light blue = input you can change. Yellow = selector. Black = formula. Green = link from another tab. Check cells should read 0 or OK."),
    ("Facts vs assumptions", "The case gives no pilot figures. Nothing here is a pilot result. Every input is labelled; public sources are listed on tab 2 and in context/sources.md."),
    ("Style", "IBM Carbon light palette; Arial as the font (Carbon's fallback for IBM Plex Sans, so the file looks the same on every PC)."),
]:
    para(lab, txt, h=30)

from openpyxl.worksheet.properties import PageSetupProperties
for ws in (ws1, ws2, ws3, ws4):
    ws.page_setup.orientation = "landscape"; ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print("saved", OUT)
