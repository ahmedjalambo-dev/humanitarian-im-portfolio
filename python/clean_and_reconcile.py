"""
clean_and_reconcile.py
----------------------
Data-quality and reconciliation pipeline for a (simulated) area-based shelter programme.

What it does
  1. Standardises and validates the site register (names, codes, coordinates).
  2. Cleans weekly field-engineer progress reports (ID typos, duplicates, impossible values,
     future dates, cumulative figures that go backwards, achievements above target).
  3. Reconciles field-engineer figures against contractor figures and classifies gaps.
  4. Cleans the household registry (gender coding, ID check digits, phone normalisation,
     cross-site duplicates, consent) and produces a PSEUDONYMISED copy for dashboards.
  5. Cleans cash-for-work attendance and calculates workdays generated.
  6. Writes every finding to a single data-quality log (one row per issue) so each
     correction is traceable, plus clean datasets, a Power BI model, GIS layers and charts.

Usage:  python clean_and_reconcile.py
"""
import json
import re
from datetime import date
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
RAW, REF = ROOT / "data" / "raw", ROOT / "data" / "reference"
OUT, CLEAN, CHARTS = ROOT / "outputs", ROOT / "outputs" / "clean", ROOT / "outputs" / "charts"
GIS, PBI = ROOT / "gis", ROOT / "powerbi"
for p in (OUT, CLEAN, CHARTS, GIS, PBI):
    p.mkdir(parents=True, exist_ok=True)

REPORTING_DATE = date(2026, 11, 2)          # last valid reporting week
GAZA_BBOX = dict(lat=(31.21, 31.60), lon=(34.20, 34.57))
GOV_BY_CODE = {"NG": "North Gaza", "GZ": "Gaza", "DB": "Deir al-Balah",
               "KY": "Khan Younis", "RF": "Rafah"}

dq_log = []


def log(dataset, record, field, issue, dimension, original, action, corrected="", status="Auto-corrected"):
    dq_log.append(dict(dataset=dataset, record_id=record, field=field, issue=issue,
                       dq_dimension=dimension, original_value=original, action_taken=action,
                       corrected_value=corrected, status=status))


def in_gaza(lat, lon):
    return (pd.notna(lat) and pd.notna(lon)
            and GAZA_BBOX["lat"][0] <= lat <= GAZA_BBOX["lat"][1]
            and GAZA_BBOX["lon"][0] <= lon <= GAZA_BBOX["lon"][1])


def normalise_site_id(raw):
    """'itcn db 4' / 'ITCN-DB-4' / ' ITCN-DB-004 ' -> 'ITCN-DB-004'"""
    m = re.match(r"^\s*ITCN[\s\-_]*([A-Z]{2})[\s\-_]*(\d{1,3})\s*$", str(raw), flags=re.I)
    return f"ITCN-{m.group(1).upper()}-{int(m.group(2)):03d}" if m else None


# =====================================================================================
# 1. SITE REGISTER
# =====================================================================================
sites = pd.read_csv(RAW / "site_register_submitted.csv")
for i, r in sites.iterrows():
    sid = r.site_id
    # governorate must match the code embedded in the site ID
    expected = GOV_BY_CODE[r.governorate_code]
    if r.governorate != expected:
        log("Site register", sid, "governorate", "Non-standard governorate name", "Consistency",
            r.governorate, "Replaced with standard name from coding list", expected)
        sites.at[i, "governorate"] = expected
    st = str(r.site_type).strip().title().replace("Centre", "Centre")
    if st != r.site_type:
        log("Site register", sid, "site_type", "Inconsistent case/whitespace", "Consistency",
            r.site_type, "Standardised to reference list", st)
        sites.at[i, "site_type"] = st
    lat, lon = r.latitude, r.longitude
    if not in_gaza(lat, lon):
        if in_gaza(lon, lat):
            log("Site register", sid, "lat/lon", "Latitude and longitude swapped", "Accuracy",
                f"{lat}, {lon}", "Swapped back; flagged to GIS focal point for confirmation",
                f"{lon}, {lat}", "Auto-corrected - confirm with GIS")
            sites.at[i, "latitude"], sites.at[i, "longitude"] = lon, lat
        else:
            issue = "Missing coordinate" if pd.isna(lat) or pd.isna(lon) else "Coordinate outside Gaza Strip"
            log("Site register", sid, "lat/lon", issue,
                "Completeness" if "Missing" in issue else "Validity",
                f"{lat}, {lon}", "Excluded from map; field engineer asked to re-capture GPS",
                "", "Open - field follow-up")
            sites.at[i, "latitude"], sites.at[i, "longitude"] = None, None
sites["coords_valid"] = sites.apply(lambda r: in_gaza(r.latitude, r.longitude), axis=1)
valid_ids = set(sites.site_id)

# =====================================================================================
# 2. FIELD REPORTS
# =====================================================================================
targets = pd.read_csv(REF / "targets.csv")
fr = pd.read_csv(RAW / "field_reports.csv")
n_raw_fr = len(fr)

for i, r in fr.iterrows():
    fixed = normalise_site_id(r.site_id)
    if fixed != r.site_id:
        if fixed in valid_ids:
            log("Field reports", r.report_id, "site_id", "Site ID format error", "Validity",
                r.site_id, "Normalised to ITCN-XX-000 pattern", fixed)
            fr.at[i, "site_id"] = fixed
        else:
            log("Field reports", r.report_id, "site_id", "Unknown site ID", "Validity",
                r.site_id, "Record held back from reporting", "", "Open - field follow-up")
            fr.at[i, "site_id"] = None

# exact duplicate submissions (same content, different report_id)
key = ["week_ending", "site_id", "component", "achieved_cumulative", "reported_by"]
dup_mask = fr.duplicated(subset=key, keep="first")
for _, r in fr[dup_mask].iterrows():
    log("Field reports", r.report_id, "(row)", "Duplicate submission", "Uniqueness", "",
        "Removed duplicate; first submission kept", "")
fr = fr[~dup_mask]

# future dates
fr["week_ending"] = pd.to_datetime(fr.week_ending)
fut = fr.week_ending.dt.date > REPORTING_DATE
for _, r in fr[fut].iterrows():
    log("Field reports", r.report_id, "week_ending", "Date after reporting period", "Timeliness",
        r.week_ending.date(), "Excluded; engineer asked to correct date", "", "Open - field follow-up")
fr = fr[~fut]

# negative values
neg = fr.achieved_cumulative < 0
for _, r in fr[neg].iterrows():
    log("Field reports", r.report_id, "achieved_cumulative", "Negative value", "Validity",
        r.achieved_cumulative, "Excluded; engineer asked to resubmit", "", "Open - field follow-up")
fr = fr[~neg & fr.site_id.notna()]

# above target
fr = fr.merge(targets[["site_id", "component", "target", "contractor"]], on=["site_id", "component"], how="left")
over = fr.achieved_cumulative > fr.target
for _, r in fr[over].iterrows():
    log("Field reports", r.report_id, "achieved_cumulative", "Achievement exceeds target", "Accuracy",
        f"{r.achieved_cumulative} vs target {r.target}",
        "Excluded pending verification against site records", "", "Open - verify with engineer")
fr = fr[~over]

# cumulative must not go backwards week to week
fr = fr.sort_values(["site_id", "component", "week_ending"])
fr["prev"] = fr.groupby(["site_id", "component"]).achieved_cumulative.shift()
back = fr.prev.notna() & (fr.achieved_cumulative < fr.prev)
for _, r in fr[back].iterrows():
    log("Field reports", r.report_id, "achieved_cumulative", "Cumulative value decreased", "Consistency",
        f"{int(r.prev)} -> {r.achieved_cumulative}", "Flagged for engineer review (kept)", "",
        "Open - verify with engineer")
fr = fr.drop(columns="prev")
fr.to_csv(CLEAN / "field_reports_clean.csv", index=False)

# latest progress per site/component
latest = (fr.sort_values("week_ending").groupby(["site_id", "component"]).tail(1)
          .rename(columns={"achieved_cumulative": "achieved", "week_ending": "last_report"}))
progress = targets.merge(latest[["site_id", "component", "achieved", "last_report"]],
                         on=["site_id", "component"], how="left")
progress["achieved"] = progress.achieved.fillna(0).astype(int)
progress["pct_complete"] = (progress.achieved / progress.target).round(3)
progress["status"] = pd.cut(progress.pct_complete, [-0.01, 0.0, 0.5, 0.999, 10],
                            labels=["Not started", "Early stage", "Advanced", "Completed"])
progress = progress.merge(sites[["site_id", "site_name", "governorate", "locality"]], on="site_id")

# =====================================================================================
# 3. RECONCILIATION: field engineers vs contractors (latest reporting week)
# =====================================================================================
cr = pd.read_csv(RAW / "contractor_reports.csv", parse_dates=["week_ending"])
last_week = pd.Timestamp(REPORTING_DATE)
rec = (progress[["site_id", "site_name", "component", "target", "achieved", "contractor"]]
       .merge(cr[cr.week_ending == last_week][["site_id", "component", "achieved_cumulative"]]
              .rename(columns={"achieved_cumulative": "contractor_reported"}),
              on=["site_id", "component"], how="left"))
rec["difference"] = rec.contractor_reported - rec.achieved
tol = (rec.target * 0.05).clip(lower=1)
rec["reconciliation_status"] = "Match"
rec.loc[rec.difference.abs() > tol, "reconciliation_status"] = "Minor gap"
rec.loc[rec.difference.abs() > tol * 3, "reconciliation_status"] = "Major gap"
rec.loc[rec.contractor_reported.isna(), "reconciliation_status"] = "No contractor report"
for _, r in rec[rec.reconciliation_status.isin(["Major gap", "No contractor report"])].iterrows():
    log("Reconciliation", f"{r.site_id} / {r.component}", "achieved",
        f"Field vs contractor: {r.reconciliation_status}", "Consistency",
        f"field={r.achieved}, contractor={r.contractor_reported}",
        "Joint verification visit requested (engineer + contractor)", "", "Open - joint verification")
rec.to_csv(OUT / "reconciliation_field_vs_contractor.csv", index=False)

# =====================================================================================
# 4. HOUSEHOLD REGISTRY
# =====================================================================================
hh = pd.read_csv(RAW / "beneficiary_registry.csv", dtype={"id_number": str, "phone": str})
GENDER = {"m": "M", "male": "M", "f": "F", "female": "F"}


def luhn_ok(idn):
    if not re.fullmatch(r"\d{9}", str(idn)):
        return False
    total = 0
    for i, ch in enumerate(idn[:8]):
        d = int(ch) * (1 if i % 2 == 0 else 2)
        total += d if d < 10 else d - 9
    return (10 - total % 10) % 10 == int(idn[8])


def norm_phone(p):
    digits = re.sub(r"\D", "", str(p))
    for pre in ("00970", "00972", "970", "972"):
        if digits.startswith(pre):
            digits = digits[len(pre):]
            break
    digits = digits.lstrip("0")
    return f"+970{digits}" if re.fullmatch(r"5[69]\d{7}", digits) else None


for i, r in hh.iterrows():
    g = str(r.head_gender).strip().lower() if pd.notna(r.head_gender) else ""
    if g in GENDER and GENDER[g] != r.head_gender:
        hh.at[i, "head_gender"] = GENDER[g]
        log("Household registry", r.hh_id, "head_gender", "Non-standard gender code", "Consistency",
            r.head_gender, "Recoded to M/F", GENDER[g])
    elif g == "":
        log("Household registry", r.hh_id, "head_gender", "Missing gender", "Completeness", "",
            "Community mobiliser to complete at next visit", "", "Open - field follow-up")
    ph = norm_phone(r.phone)
    if ph is None:
        log("Household registry", r.hh_id, "phone", "Invalid phone number", "Validity", "(masked)",
            "Flagged for update", "", "Open - field follow-up")
    hh.at[i, "phone"] = ph
    if not luhn_ok(r.id_number):
        log("Household registry", r.hh_id, "id_number", "ID number fails check digit", "Validity",
            "(masked)", "Verify against ID document", "", "Open - field follow-up")
    if r.hh_size == 0:
        log("Household registry", r.hh_id, "hh_size", "Household size is zero", "Validity", 0,
            "Verify with household", "", "Open - field follow-up")
    if r.consent_given != "yes":
        log("Household registry", r.hh_id, "consent_given", "No recorded consent", "Data protection",
            r.consent_given if pd.notna(r.consent_given) else "", "Excluded from shared datasets", "",
            "Open - obtain consent")

hh["id_valid"] = hh.id_number.apply(luhn_ok)
dup_ids = hh[hh.duplicated("id_number", keep=False)].sort_values("id_number")
for idn, grp in dup_ids.groupby("id_number"):
    log("Household registry", " / ".join(grp.hh_id), "id_number",
        f"Same head-of-household ID registered in {grp.site_id.nunique()} sites", "Uniqueness",
        "(masked)", "Held for case review by Area Coordinator", "", "Open - case review")
hh["possible_duplicate"] = hh.id_number.isin(dup_ids.id_number)

# Pseudonymised copy: no names / ID numbers / phones leave the secure registry
hh_pseudo = hh[hh.consent_given == "yes"][["hh_id", "site_id", "head_gender", "hh_size",
                                           "members_with_disability", "registration_date",
                                           "possible_duplicate"]]
hh_pseudo.to_csv(CLEAN / "households_pseudonymised.csv", index=False)

# =====================================================================================
# 5. CASH-FOR-WORK
# =====================================================================================
att = pd.read_csv(RAW / "cfw_attendance.csv")
workers = pd.read_csv(REF / "cfw_workers.csv", dtype={"id_number": str})
d = att.duplicated(["worker_id", "date"], keep="first")
log_n = int(d.sum())
for _, r in att[d].iterrows():
    log("CfW attendance", f"{r.worker_id} {r.date}", "(row)", "Same worker recorded twice on one day",
        "Uniqueness", "", "Duplicate removed")
att = att[~d]
unknown = ~att.worker_id.isin(workers.worker_id)
for _, r in att[unknown].iterrows():
    log("CfW attendance", f"{r.worker_id} {r.date}", "worker_id", "Worker not in worker register",
        "Validity", r.worker_id, "Excluded from payment sheet", "", "Open - verify worker")
att = att[~unknown]
bad_h = att.hours > 8
for _, r in att[bad_h].iterrows():
    log("CfW attendance", f"{r.worker_id} {r.date}", "hours", "Hours above daily maximum (8)",
        "Validity", r.hours, "Capped at 0 pending timesheet check", 0, "Open - verify timesheet")
att.loc[bad_h, "hours"] = 0
att["workdays"] = att.hours.map(lambda h: 1.0 if h >= 6 else 0.5 if h >= 4 else 0.0)
att = att.merge(workers[["worker_id", "gender"]], on="worker_id")
att.to_csv(CLEAN / "cfw_attendance_clean.csv", index=False)
cfw_site = (att.groupby("site_id").agg(workers=("worker_id", "nunique"),
                                       workdays=("workdays", "sum"),
                                       female_workers=("worker_id", lambda s: s[att.loc[s.index, "gender"] == "F"].nunique()))
            .reset_index())

# =====================================================================================
# 6. OUTPUTS
# =====================================================================================
dq = pd.DataFrame(dq_log)
dq.insert(0, "issue_no", [f"DQ-{k:04d}" for k in range(1, len(dq) + 1)])
dq.insert(1, "date_logged", REPORTING_DATE.isoformat())
dq.to_csv(OUT / "data_quality_log.csv", index=False)
sites.to_csv(CLEAN / "site_register_clean.csv", index=False)
progress.to_csv(OUT / "progress_by_site_component.csv", index=False)

# site-level summary for dashboard + GIS
hh_site = hh_pseudo.groupby("site_id").agg(households_registered=("hh_id", "count"),
                                           individuals_registered=("hh_size", "sum"),
                                           female_headed=("head_gender", lambda s: (s == "F").sum())).reset_index()
rhu = progress[progress.component == "Shelter"][["site_id", "target", "achieved"]].rename(
    columns={"target": "rhu_target", "achieved": "rhu_installed"})
site_summary = (sites.merge(hh_site, on="site_id", how="left").merge(rhu, on="site_id", how="left")
                .merge(cfw_site, on="site_id", how="left")
                .merge(progress.groupby("site_id").pct_complete.mean().round(3).rename("avg_progress"),
                       on="site_id", how="left"))
open_issues = dq[dq.status.str.startswith("Open")]
site_summary.to_csv(OUT / "site_summary.csv", index=False)

# GIS layers
pts = site_summary[site_summary.coords_valid]
geojson = {"type": "FeatureCollection", "features": [
    {"type": "Feature",
     "geometry": {"type": "Point", "coordinates": [r.longitude, r.latitude]},
     "properties": {k: (None if pd.isna(v) else (v.item() if hasattr(v, "item") else v))
                    for k, v in r.drop(["latitude", "longitude"]).items()}}
    for _, r in pts.iterrows()]}
(GIS / "itcn_sites.geojson").write_text(json.dumps(geojson, indent=1, default=str))
pts.to_csv(GIS / "itcn_sites_points.csv", index=False)

# Power BI model (one workbook, one table per sheet)
weekly = (fr.groupby(["week_ending", "component"]).achieved_cumulative.sum().reset_index())
with pd.ExcelWriter(PBI / "ITCN_PowerBI_Data.xlsx") as xw:
    site_summary.to_excel(xw, "Sites", index=False)
    progress.to_excel(xw, "Progress", index=False)
    weekly.to_excel(xw, "Weekly_Trend", index=False)
    rec.to_excel(xw, "Reconciliation", index=False)
    att.groupby(["site_id", "date", "gender"]).workdays.sum().reset_index().to_excel(xw, "CfW_Workdays", index=False)
    dq.to_excel(xw, "DQ_Log", index=False)

# Charts
plt.rcParams.update({"font.size": 9, "axes.spines.top": False, "axes.spines.right": False})
comp = progress.groupby("component")[["target", "achieved"]].sum()
comp["pct"] = comp.achieved / comp.target * 100
ax = comp.pct.sort_values().plot.barh(color="#1f6fb2", figsize=(6, 3.2))
ax.set_xlabel("% of target achieved"); ax.set_ylabel(""); ax.set_xlim(0, 100)
ax.set_title(f"Progress by component (as of {REPORTING_DATE})", loc="left")
plt.tight_layout(); plt.savefig(CHARTS / "progress_by_component.png", dpi=160); plt.close()

dim = dq.groupby("dq_dimension").size().sort_values()
ax = dim.plot.barh(color="#c0504d", figsize=(6, 3))
ax.set_xlabel("Issues logged"); ax.set_ylabel(""); ax.set_title("Data-quality issues by dimension", loc="left")
plt.tight_layout(); plt.savefig(CHARTS / "dq_issues_by_dimension.png", dpi=160); plt.close()

rs = rec.reconciliation_status.value_counts()
ax = rs.plot.bar(color=["#4f9d69", "#e0a030", "#c0504d", "#888888"][:len(rs)], figsize=(5, 3), rot=0)
ax.set_ylabel("Site-component records"); ax.set_title("Field vs contractor reconciliation", loc="left")
plt.tight_layout(); plt.savefig(CHARTS / "reconciliation_status.png", dpi=160); plt.close()

# Markdown report
summary = f"""# Data-Quality & Reconciliation Report (simulated data)
Reporting period ending **{REPORTING_DATE}** - generated by `python/clean_and_reconcile.py`

## Headline figures
| Item | Value |
|---|---|
| Sites in register | {len(sites)} ({int(sites.coords_valid.sum())} with valid coordinates) |
| Field reports received / accepted | {n_raw_fr} / {len(fr)} |
| Households registered (consented, pseudonymised) | {len(hh_pseudo):,} |
| Relief Housing Units installed / target | {int(rhu.rhu_installed.sum()):,} / {int(rhu.rhu_target.sum()):,} |
| Cash-for-work workers / workdays generated | {att.worker_id.nunique()} / {att.workdays.sum():,.1f} |
| Data-quality issues logged | {len(dq)} ({len(open_issues)} open, {len(dq) - len(open_issues)} auto-corrected) |

## Issues by type
{dq.groupby(['dataset', 'issue']).size().rename('count').reset_index().to_markdown(index=False)}

## Field vs contractor reconciliation (latest week)
{rec.reconciliation_status.value_counts().rename('records').to_markdown()}

## Open items requiring field follow-up
{open_issues.groupby('status').size().rename('count').to_markdown()}

> All records are fictional and generated for portfolio purposes. Personal fields (names,
> ID numbers, phone numbers) never leave the secure registry; shared outputs are pseudonymised.
"""
(OUT / "data_quality_report.md").write_text(summary)
print(summary)
