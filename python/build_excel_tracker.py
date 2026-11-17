"""
build_excel_tracker.py
----------------------
Builds excel/ITCN_IM_Tracker.xlsx from the cleaned outputs of clean_and_reconcile.py.
The workbook is a working information-management tool: input columns are validated with
drop-down lists, every calculated column is a live Excel formula (INDEX/MATCH, SUMIFS,
COUNTIFS, IFERROR), and a summary sheet updates automatically when data are added.
"""
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "outputs"
sites = pd.read_csv(OUT / "clean" / "site_register_clean.csv")
prog = pd.read_csv(OUT / "progress_by_site_component.csv")
rec = pd.read_csv(OUT / "reconciliation_field_vs_contractor.csv")
dq = pd.read_csv(OUT / "data_quality_log.csv")

F = "Arial"
H_INPUT = PatternFill("solid", fgColor="FFF2CC")    # yellow = data-entry column
H_CALC = PatternFill("solid", fgColor="D9E1F2")     # blue-grey = formula column (do not edit)
TITLE = Font(name=F, size=14, bold=True, color="1F3864")
BOLD = Font(name=F, bold=True)
BASE = Font(name=F, size=10)
thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = Workbook()


def style_sheet(ws):
    for row in ws.iter_rows():
        for c in row:
            if c.font == Font() or not c.font.bold:
                c.font = Font(name=F, size=10, bold=c.font.bold, color=c.font.color)


def header(ws, row, cols, input_cols=()):
    for j, name in enumerate(cols, 1):
        c = ws.cell(row=row, column=j, value=name)
        c.font = Font(name=F, bold=True, size=10)
        c.fill = H_INPUT if name in input_cols else H_CALC
        c.alignment = Alignment(wrap_text=True, vertical="center")
        c.border = BOX


def widths(ws, w):
    for j, x in enumerate(w, 1):
        ws.column_dimensions[get_column_letter(j)].width = x


def add_table(ws, ref, name):
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws.add_table(t)


# ------------------------------------------------------------------ README
ws = wb.active
ws.title = "README"
lines = [
    ("ITCN Information-Management Tracker", TITLE),
    ("Portfolio project - all data are SIMULATED (fictional sites, households and workers).", Font(name=F, italic=True, color="C00000")),
    ("", None),
    ("Purpose", BOLD),
    ("Single source of truth for site locations, activity progress against targets, field-vs-contractor reconciliation and data-quality follow-up across programme sites.", BASE),
    ("", None),
    ("How to use", BOLD),
    ("1. Add or edit data only in YELLOW columns. BLUE-GREY columns are formulas - do not type over them.", BASE),
    ("2. Site_Register: one row per site. Coordinates are checked automatically against the Gaza Strip bounding box.", BASE),
    ("3. Activity_Tracker: one row per site x component. Pick Site ID and Component from the drop-down; name, governorate, % complete, status and flags fill in automatically.", BASE),
    ("4. Reconciliation: enter the contractor's figure next to the field engineer's figure; the gap and status are calculated.", BASE),
    ("5. DQ_Log: every data problem found gets one row with dimension, action and status, so each correction is traceable.", BASE),
    ("6. Summary: updates itself (SUMIFS / COUNTIFS). Add a PivotTable from the Activity_Tracker table for ad-hoc views.", BASE),
    ("", None),
    ("Colour legend", BOLD),
    ("Yellow header = data-entry column", BASE),
    ("Blue-grey header = calculated column (formula)", BASE),
    ("", None),
    ("Version control", BOLD),
    ("File naming: ITCN_IM_Tracker_vYYYYMMDD_initials.xlsx - keep previous versions in /archive; never overwrite the shared master.", BASE),
    ("v20261102 - initial build from cleaned datasets (python/clean_and_reconcile.py).", BASE),
]
for i, (txt, font) in enumerate(lines, 1):
    c = ws.cell(row=i, column=1, value=txt)
    c.font = font or BASE
    c.alignment = Alignment(wrap_text=True, vertical="top")
ws.cell(row=16, column=1).fill = H_INPUT
ws.cell(row=17, column=1).fill = H_CALC
ws.column_dimensions["A"].width = 120

# ------------------------------------------------------------------ Lists
wl = wb.create_sheet("Lists")
lists = {
    "Governorate": ["North Gaza", "Gaza", "Deir al-Balah", "Khan Younis", "Rafah"],
    "Site_Type": ["Transitional Neighborhood", "Collective Centre", "Informal Site"],
    "Component": ["Shelter", "WASH", "Electricity", "Roads & Access", "Community Facility", "Solid Waste"],
    "Field_Engineer": [f"FE-{i:02d}" for i in range(1, 7)],
    "Contractor": ["CTR-A", "CTR-B", "CTR-C", "CTR-D"],
    "DQ_Dimension": ["Accuracy", "Completeness", "Consistency", "Validity", "Timeliness", "Uniqueness", "Data protection"],
    "DQ_Status": ["Auto-corrected", "Auto-corrected - confirm with GIS", "Open - field follow-up",
                  "Open - verify with engineer", "Open - joint verification", "Open - case review",
                  "Open - obtain consent", "Open - verify timesheet", "Open - verify worker", "Closed"],
}
for j, (name, vals) in enumerate(lists.items(), 1):
    wl.cell(row=1, column=j, value=name).font = BOLD
    for i, v in enumerate(vals, 2):
        wl.cell(row=i, column=j, value=v).font = BASE
    col = get_column_letter(j)
    wb.defined_names[f"lst_{name}"] = DefinedName(f"lst_{name}", attr_text=f"Lists!${col}$2:${col}${len(vals) + 1}")
    wl.column_dimensions[col].width = 30

# ------------------------------------------------------------------ Data dictionary
wd = wb.create_sheet("Data_Dictionary")
dd = [
    ("Site_Register", "site_id", "Unique site code ITCN-<GOV>-<NNN>", "Text", "Pattern ITCN-XX-000; unique", "ITCN-DB-003"),
    ("Site_Register", "governorate", "Governorate", "List", "lst_Governorate", "Deir al-Balah"),
    ("Site_Register", "site_type", "Type of site", "List", "lst_Site_Type", "Transitional Neighborhood"),
    ("Site_Register", "latitude / longitude", "Site centroid, WGS84 decimal degrees", "Decimal", "Lat 31.21-31.60, Lon 34.20-34.57", "31.448479 / 34.397058"),
    ("Site_Register", "coord_check", "Formula: OK if inside Gaza Strip box, else CHECK", "Formula", "OK / CHECK / MISSING", "OK"),
    ("Site_Register", "households / individuals", "Population reported at site", "Whole number", ">= 0", "65 / 434"),
    ("Activity_Tracker", "component", "Programme component", "List", "lst_Component", "Shelter"),
    ("Activity_Tracker", "target", "Agreed output target for site x component", "Whole number", "> 0", "93"),
    ("Activity_Tracker", "achieved", "Cumulative output verified by field engineer", "Whole number", "0 - target", "70"),
    ("Activity_Tracker", "pct_complete", "Formula: achieved / target", "Formula", "0-100%", "75%"),
    ("Activity_Tracker", "status", "Formula: Not started / Early stage / Advanced / Completed", "Formula", "Derived", "Advanced"),
    ("Activity_Tracker", "flag", "Formula: data-quality warning", "Formula", "Above target / Unknown site / blank", ""),
    ("Reconciliation", "difference", "Formula: contractor figure - field figure", "Formula", "Any integer", "-3"),
    ("Reconciliation", "rec_status", "Formula: Match if |gap| <= 5% of target (min 1); Minor <= 15%; else Major", "Formula", "Match / Minor gap / Major gap", "Match"),
    ("DQ_Log", "dq_dimension", "Data-quality dimension affected", "List", "lst_DQ_Dimension", "Validity"),
    ("DQ_Log", "status", "Follow-up status", "List", "lst_DQ_Status", "Open - field follow-up"),
]
header(wd, 1, ["Sheet", "Field", "Definition", "Type", "Allowed values / rule", "Example"])
for i, r in enumerate(dd, 2):
    for j, v in enumerate(r, 1):
        wd.cell(row=i, column=j, value=v).font = BASE
widths(wd, [18, 24, 62, 14, 34, 26])
wd.freeze_panes = "A2"

# ------------------------------------------------------------------ Site register
wsr = wb.create_sheet("Site_Register")
cols = ["site_id", "site_name", "governorate", "locality", "site_type", "latitude", "longitude",
        "households", "individuals", "field_engineer", "start_date", "coord_check"]
header(wsr, 1, cols, input_cols=cols[:-1])
for i, r in enumerate(sites.itertuples(index=False), 2):
    vals = [r.site_id, r.site_name, r.governorate, r.locality, r.site_type,
            None if pd.isna(r.latitude) else r.latitude, None if pd.isna(r.longitude) else r.longitude,
            r.households, r.individuals, r.field_engineer, pd.Timestamp(r.start_date).to_pydatetime()]
    for j, v in enumerate(vals, 1):
        wsr.cell(row=i, column=j, value=v)
    wsr.cell(row=i, column=11).number_format = "yyyy-mm-dd"
    wsr.cell(row=i, column=12, value=f'=IF(OR(F{i}="",G{i}=""),"MISSING",IF(AND(F{i}>=31.21,F{i}<=31.6,G{i}>=34.2,G{i}<=34.57),"OK","CHECK"))')
n_sites = len(sites) + 1
SITE_LAST = 200  # room for new sites
add_table(wsr, f"A1:L{n_sites}", "tblSites")
widths(wsr, [14, 26, 15, 22, 24, 11, 11, 11, 11, 13, 12, 12])
wsr.freeze_panes = "B2"
wb.defined_names["lst_SiteID"] = DefinedName("lst_SiteID", attr_text=f"Site_Register!$A$2:$A${SITE_LAST}")
for rng, lst in [(f"C2:C{SITE_LAST}", "lst_Governorate"), (f"E2:E{SITE_LAST}", "lst_Site_Type"),
                 (f"J2:J{SITE_LAST}", "lst_Field_Engineer")]:
    dv = DataValidation(type="list", formula1=f"={lst}", allow_blank=True, showErrorMessage=True,
                        errorTitle="Invalid entry", error="Choose a value from the list.")
    dv.add(rng); wsr.add_data_validation(dv)
dv = DataValidation(type="decimal", operator="between", formula1="31.21", formula2="31.6", allow_blank=True,
                    showErrorMessage=True, error="Latitude must be inside the Gaza Strip (31.21 - 31.60).")
dv.add(f"F2:F{SITE_LAST}"); wsr.add_data_validation(dv)
dv = DataValidation(type="decimal", operator="between", formula1="34.2", formula2="34.57", allow_blank=True,
                    showErrorMessage=True, error="Longitude must be inside the Gaza Strip (34.20 - 34.57).")
dv.add(f"G2:G{SITE_LAST}"); wsr.add_data_validation(dv)
dv = DataValidation(type="custom", formula1="=COUNTIF($A$2:$A$200,A2)=1", showErrorMessage=True,
                    error="Site ID already exists - IDs must be unique.")
dv.add(f"A2:A{SITE_LAST}"); wsr.add_data_validation(dv)
red = PatternFill("solid", fgColor="F8CBAD")
wsr.conditional_formatting.add(f"L2:L{SITE_LAST}", CellIsRule(operator="notEqual", formula=['"OK"'], fill=red))

# ------------------------------------------------------------------ Activity tracker
wa = wb.create_sheet("Activity_Tracker")
cols = ["site_id", "component", "indicator", "unit", "target", "achieved", "last_report", "contractor",
        "site_name", "governorate", "pct_complete", "status", "flag"]
inputs = cols[:8]
header(wa, 1, cols, input_cols=inputs)
prog = prog.sort_values(["site_id", "component"])
for i, r in enumerate(prog.itertuples(index=False), 2):
    vals = [r.site_id, r.component, r.indicator, r.unit, int(r.target), int(r.achieved),
            pd.Timestamp(r.last_report).to_pydatetime() if pd.notna(r.last_report) else None, r.contractor]
    for j, v in enumerate(vals, 1):
        wa.cell(row=i, column=j, value=v)
    wa.cell(row=i, column=7).number_format = "yyyy-mm-dd"
    wa.cell(row=i, column=9, value=f'=IFERROR(INDEX(Site_Register!$B:$B,MATCH(A{i},Site_Register!$A:$A,0)),"")')
    wa.cell(row=i, column=10, value=f'=IFERROR(INDEX(Site_Register!$C:$C,MATCH(A{i},Site_Register!$A:$A,0)),"")')
    wa.cell(row=i, column=11, value=f'=IFERROR(F{i}/E{i},0)').number_format = "0%"
    wa.cell(row=i, column=12, value=f'=IF(F{i}="","",IF(F{i}=0,"Not started",IF(K{i}>=1,"Completed",IF(K{i}>=0.5,"Advanced","Early stage"))))')
    wa.cell(row=i, column=13, value=f'=IF(I{i}="","Unknown site",IF(F{i}>E{i},"Above target - verify",""))')
n_act = len(prog) + 1
add_table(wa, f"A1:M{n_act}", "tblActivities")
widths(wa, [14, 19, 34, 9, 8, 10, 12, 11, 26, 15, 12, 13, 20])
wa.freeze_panes = "C2"
ACT_LAST = 1000
for rng, lst in [(f"A2:A{ACT_LAST}", "lst_SiteID"), (f"B2:B{ACT_LAST}", "lst_Component"),
                 (f"H2:H{ACT_LAST}", "lst_Contractor")]:
    dv = DataValidation(type="list", formula1=f"={lst}", allow_blank=True, showErrorMessage=True,
                        error="Choose a value from the list.")
    dv.add(rng); wa.add_data_validation(dv)
dv = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
                    showErrorMessage=True, error="Enter a whole number of 0 or more.")
dv.add(f"E2:F{ACT_LAST}"); wa.add_data_validation(dv)
wa.conditional_formatting.add(f"K2:K{ACT_LAST}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="5B9BD5"))
wa.conditional_formatting.add(f"M2:M{ACT_LAST}", FormulaRule(formula=[f'LEN(M2)>0'], fill=red))

# ------------------------------------------------------------------ Reconciliation
wr = wb.create_sheet("Reconciliation")
cols = ["site_id", "component", "target", "field_reported", "contractor_reported", "contractor",
        "difference", "gap_pct_of_target", "rec_status"]
header(wr, 1, cols, input_cols=cols[:6])
rec = rec.sort_values(["site_id", "component"])
for i, r in enumerate(rec.itertuples(index=False), 2):
    vals = [r.site_id, r.component, int(r.target), int(r.achieved),
            None if pd.isna(r.contractor_reported) else int(r.contractor_reported), r.contractor]
    for j, v in enumerate(vals, 1):
        wr.cell(row=i, column=j, value=v)
    wr.cell(row=i, column=7, value=f'=IF(E{i}="","",E{i}-D{i})')
    wr.cell(row=i, column=8, value=f'=IF(OR(E{i}="",C{i}=0),"",ABS(G{i})/C{i})').number_format = "0.0%"
    wr.cell(row=i, column=9, value=(f'=IF(E{i}="","No contractor report",IF(ABS(G{i})<=MAX(1,0.05*C{i}),"Match",'
                                    f'IF(ABS(G{i})<=MAX(3,0.15*C{i}),"Minor gap","Major gap")))'))
n_rec = len(rec) + 1
add_table(wr, f"A1:I{n_rec}", "tblReconciliation")
widths(wr, [14, 19, 8, 14, 18, 11, 11, 16, 20])
wr.freeze_panes = "C2"
wr.conditional_formatting.add(f"I2:I{n_rec}", CellIsRule(operator="equal", formula=['"Major gap"'], fill=red))
wr.conditional_formatting.add(f"I2:I{n_rec}", CellIsRule(operator="equal", formula=['"Minor gap"'], fill=PatternFill("solid", fgColor="FFE699")))
wr.conditional_formatting.add(f"I2:I{n_rec}", CellIsRule(operator="equal", formula=['"Match"'], fill=PatternFill("solid", fgColor="C6EFCE")))

# ------------------------------------------------------------------ DQ log
wq = wb.create_sheet("DQ_Log")
cols = ["issue_no", "date_logged", "dataset", "record_id", "field", "issue", "dq_dimension",
        "original_value", "action_taken", "corrected_value", "status", "responsible", "date_closed"]
header(wq, 1, cols, input_cols=cols)
resp = {"Site register": "GIS focal point / FE", "Field reports": "Field Engineer",
        "Reconciliation": "FE + Contractor", "Household registry": "Community Mobiliser",
        "CfW attendance": "Site supervisor"}
for i, r in enumerate(dq.itertuples(index=False), 2):
    vals = [r.issue_no, pd.Timestamp(r.date_logged).to_pydatetime(), r.dataset, r.record_id, r.field, r.issue,
            r.dq_dimension, None if pd.isna(r.original_value) else str(r.original_value), r.action_taken,
            None if pd.isna(r.corrected_value) else str(r.corrected_value), r.status, resp.get(r.dataset, ""), None]
    for j, v in enumerate(vals, 1):
        wq.cell(row=i, column=j, value=v)
    wq.cell(row=i, column=2).number_format = "yyyy-mm-dd"
    wq.cell(row=i, column=13).number_format = "yyyy-mm-dd"
n_dq = len(dq) + 1
add_table(wq, f"A1:M{n_dq}", "tblDQLog")
widths(wq, [10, 11, 18, 26, 18, 38, 15, 20, 46, 16, 28, 20, 11])
wq.freeze_panes = "B2"
for rng, lst in [(f"G2:G2000", "lst_DQ_Dimension"), (f"K2:K2000", "lst_DQ_Status")]:
    dv = DataValidation(type="list", formula1=f"={lst}", allow_blank=True, showErrorMessage=True,
                        error="Choose a value from the list.")
    dv.add(rng); wq.add_data_validation(dv)
wq.conditional_formatting.add("K2:K2000", FormulaRule(formula=['LEFT($K2,4)="Open"'], fill=PatternFill("solid", fgColor="FCE4D6")))

# ------------------------------------------------------------------ Summary
s = wb.create_sheet("Summary", 1)
s["A1"] = "ITCN Programme Summary"; s["A1"].font = TITLE
s["A2"] = "Live formulas - updates when tracker sheets change. Simulated data."; s["A2"].font = Font(name=F, italic=True, color="7F7F7F")
AT, SR, RC, DL = "Activity_Tracker", "Site_Register", "Reconciliation", "DQ_Log"
kpis = [
    ("Sites in register", f"=COUNTA({SR}!A2:A{SITE_LAST})", "0"),
    ("Sites with coordinate issues", f'=COUNTIF({SR}!L2:L{SITE_LAST},"CHECK")+COUNTIF({SR}!L2:L{SITE_LAST},"MISSING")', "0"),
    ("Households (site register)", f"=SUM({SR}!H2:H{SITE_LAST})", "#,##0"),
    ("Relief Housing Units installed", f'=SUMIFS({AT}!F2:F{ACT_LAST},{AT}!B2:B{ACT_LAST},"Shelter")', "#,##0"),
    ("RHU target", f'=SUMIFS({AT}!E2:E{ACT_LAST},{AT}!B2:B{ACT_LAST},"Shelter")', "#,##0"),
    ("RHU % complete", "=IFERROR(B7/B8,0)", "0%"),
    ("Activities completed", f'=COUNTIF({AT}!L2:L{ACT_LAST},"Completed")', "0"),
    ("Reconciliation - major gaps", f'=COUNTIF({RC}!I2:I{n_rec + 200},"Major gap")', "0"),
    ("Open data-quality issues", f'=COUNTIF({DL}!K2:K2000,"Open*")', "#,##0"),
]
s["A3"] = "Key figures"; s["A3"].font = BOLD
for k, (lbl, f, fmt) in enumerate(kpis, 4):
    s.cell(row=k, column=1, value=lbl).font = BASE
    c = s.cell(row=k, column=2, value=f); c.number_format = fmt; c.font = Font(name=F, bold=True)
    c.border = BOX

r0 = 15
s.cell(row=r0, column=1, value="Progress by component").font = BOLD
header(s, r0 + 1, ["Component", "Target", "Achieved", "% complete", "Activities", "Completed"])
for k, comp in enumerate(lists["Component"], r0 + 2):
    s.cell(row=k, column=1, value=comp)
    s.cell(row=k, column=2, value=f'=SUMIFS({AT}!$E$2:$E${ACT_LAST},{AT}!$B$2:$B${ACT_LAST},$A{k})').number_format = "#,##0"
    s.cell(row=k, column=3, value=f'=SUMIFS({AT}!$F$2:$F${ACT_LAST},{AT}!$B$2:$B${ACT_LAST},$A{k})').number_format = "#,##0"
    s.cell(row=k, column=4, value=f'=IFERROR(C{k}/B{k},0)').number_format = "0%"
    s.cell(row=k, column=5, value=f'=COUNTIF({AT}!$B$2:$B${ACT_LAST},$A{k})')
    s.cell(row=k, column=6, value=f'=COUNTIFS({AT}!$B$2:$B${ACT_LAST},$A{k},{AT}!$L$2:$L${ACT_LAST},"Completed")')
comp_end = r0 + 1 + len(lists["Component"])

g0 = comp_end + 3
s.cell(row=g0, column=1, value="Progress by governorate (all components)").font = BOLD
header(s, g0 + 1, ["Governorate", "Sites", "Households", "Avg % complete", "RHU installed", "Major gaps"])
for k, gov in enumerate(lists["Governorate"], g0 + 2):
    s.cell(row=k, column=1, value=gov)
    s.cell(row=k, column=2, value=f'=COUNTIF({SR}!$C$2:$C${SITE_LAST},$A{k})')
    s.cell(row=k, column=3, value=f'=SUMIFS({SR}!$H$2:$H${SITE_LAST},{SR}!$C$2:$C${SITE_LAST},$A{k})').number_format = "#,##0"
    s.cell(row=k, column=4, value=f'=IFERROR(AVERAGEIFS({AT}!$K$2:$K${ACT_LAST},{AT}!$J$2:$J${ACT_LAST},$A{k}),0)').number_format = "0%"
    s.cell(row=k, column=5, value=f'=SUMIFS({AT}!$F$2:$F${ACT_LAST},{AT}!$J$2:$J${ACT_LAST},$A{k},{AT}!$B$2:$B${ACT_LAST},"Shelter")').number_format = "#,##0"
    s.cell(row=k, column=6, value=(f'=SUMPRODUCT(({RC}!$I$2:$I${n_rec}="Major gap")*'
                                   f'(IFERROR(INDEX({SR}!$C$2:$C${n_sites},MATCH({RC}!$A$2:$A${n_rec},{SR}!$A$2:$A${n_sites},0)),"")=$A{k}))'))
gov_end = g0 + 1 + len(lists["Governorate"])

q0 = gov_end + 3
s.cell(row=q0, column=1, value="Data-quality issues by dimension").font = BOLD
header(s, q0 + 1, ["Dimension", "Logged", "Open"])
for k, d in enumerate(lists["DQ_Dimension"], q0 + 2):
    s.cell(row=k, column=1, value=d)
    s.cell(row=k, column=2, value=f'=COUNTIF({DL}!$G$2:$G$2000,$A{k})')
    s.cell(row=k, column=3, value=f'=COUNTIFS({DL}!$G$2:$G$2000,$A{k},{DL}!$K$2:$K$2000,"Open*")')
widths(s, [34, 14, 14, 15, 14, 13])

ch = BarChart(); ch.type = "bar"; ch.style = 10
ch.title = "% of target achieved by component"
ch.add_data(Reference(s, min_col=4, min_row=r0 + 1, max_row=comp_end), titles_from_data=True)
ch.set_categories(Reference(s, min_col=1, min_row=r0 + 2, max_row=comp_end))
ch.y_axis.number_format = "0%"; ch.y_axis.scaling.min = 0; ch.y_axis.scaling.max = 1
ch.legend = None; ch.height = 7; ch.width = 14
ch.roundedCorners = False; ch.title.overlay = False
ch.x_axis.delete = False; ch.y_axis.delete = False
s.add_chart(ch, "H3")

ch2 = BarChart(); ch2.type = "col"; ch2.style = 10
ch2.title = "Open data-quality issues by dimension"
ch2.add_data(Reference(s, min_col=3, min_row=q0 + 1, max_row=q0 + 1 + len(lists["DQ_Dimension"])), titles_from_data=True)
ch2.set_categories(Reference(s, min_col=1, min_row=q0 + 2, max_row=q0 + 1 + len(lists["DQ_Dimension"])))
ch2.legend = None; ch2.height = 7; ch2.width = 14
ch2.roundedCorners = False; ch2.title.overlay = False
ch2.x_axis.delete = False; ch2.y_axis.delete = False
s.add_chart(ch2, "H19")

# fonts everywhere
for sh in wb.worksheets:
    for row in sh.iter_rows():
        for c in row:
            if c.value is not None and c.font.name != F:
                c.font = Font(name=F, size=c.font.size or 10, bold=c.font.bold, italic=c.font.italic, color=c.font.color)

dest = ROOT / "excel" / "ITCN_IM_Tracker.xlsx"
dest.parent.mkdir(exist_ok=True)
wb.save(dest)
print("saved", dest)
