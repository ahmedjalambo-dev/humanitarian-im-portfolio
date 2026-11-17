"""
generate_sample_data.py
-----------------------
Generates a SIMULATED (fictional) dataset that mimics the information flows of an
area-based transitional-shelter programme in the Gaza Strip:

  * site register (locations + coordinates)
  * weekly field-engineer progress reports
  * contractor progress reports (second source, for reconciliation)
  * household beneficiary registry
  * cash-for-work (temporary employment) attendance

All names, ID numbers and phone numbers are randomly generated. No real person or
real beneficiary record is included. Data-quality problems (duplicates, typos,
swapped coordinates, impossible values, inconsistent sources) are injected on
purpose so the cleaning / validation pipeline has something realistic to catch.

Usage:  python generate_sample_data.py   (writes to ../data/raw and ../data/reference)
"""
import random
from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pandas as pd

SEED = 2026
random.seed(SEED)
np.random.seed(SEED)

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw"
REF = ROOT / "data" / "reference"
RAW.mkdir(parents=True, exist_ok=True)
REF.mkdir(parents=True, exist_ok=True)

# --------------------------------------------------------------------------------------
# Reference lists
# --------------------------------------------------------------------------------------
# (governorate code, governorate, locality, approx. centre lat, lon)
LOCALITIES = [
    ("NG", "North Gaza", "Jabalia", 31.530, 34.483),
    ("GZ", "Gaza", "Gaza City - Rimal", 31.522, 34.445),
    ("GZ", "Gaza", "Gaza City - Sheikh Radwan", 31.533, 34.462),
    ("DB", "Deir al-Balah", "Deir al-Balah", 31.418, 34.351),
    ("DB", "Deir al-Balah", "Nuseirat", 31.449, 34.393),
    ("DB", "Deir al-Balah", "Bureij", 31.440, 34.403),
    ("DB", "Deir al-Balah", "Maghazi", 31.421, 34.384),
    ("KY", "Khan Younis", "Khan Younis", 31.344, 34.306),
    ("KY", "Khan Younis", "Al-Mawasi", 31.360, 34.270),
    ("RF", "Rafah", "Rafah", 31.297, 34.250),
]
SITE_TYPES = ["Transitional Neighborhood", "Collective Centre", "Informal Site"]
COMPONENTS = {
    # component: (indicator, unit, typical target range)
    "Shelter": ("Relief Housing Units installed", "RHU", (20, 120)),
    "WASH": ("Water points rehabilitated", "point", (2, 12)),
    "Electricity": ("Solar lighting units installed", "unit", (10, 60)),
    "Roads & Access": ("Access road rehabilitated", "metre", (150, 900)),
    "Community Facility": ("Community facilities rehabilitated", "facility", (1, 3)),
    "Solid Waste": ("Waste collection points established", "point", (2, 10)),
}
FIRST_M = ["Ahmed", "Mohammed", "Mahmoud", "Omar", "Khaled", "Yousef", "Ibrahim", "Hassan",
           "Ali", "Mustafa", "Sami", "Rami", "Nidal", "Iyad", "Hani", "Fadi", "Tamer", "Wael"]
FIRST_F = ["Fatima", "Aisha", "Mariam", "Huda", "Asmaa", "Noor", "Samar", "Reem", "Amal",
           "Hanan", "Lina", "Rana", "Dina", "Iman", "Sara", "Nisreen"]
FAMILY = ["Abu Shaaban", "Al-Masri", "Al-Najjar", "Abu Hasira", "Al-Agha", "Al-Farra",
          "Shurrab", "Abu Taha", "Al-Astal", "Al-Kurd", "Hamad", "Abu Amra", "Al-Hindi",
          "Barakat", "Saqqa", "Abu Jarad", "Al-Batsh", "Jaber", "Qudaih", "Radwan"]
FIELD_ENGINEERS = [f"FE-{i:02d}" for i in range(1, 7)]
CONTRACTORS = ["CTR-A", "CTR-B", "CTR-C", "CTR-D"]
CFW_TASKS = ["Rubble removal", "Site cleaning", "Drainage works", "RHU assembly support",
             "Solid waste collection", "Community facility works"]


def luhn_check_digit(body: str) -> int:
    """Check digit for a 9-digit ID (Luhn-style, as used by the Palestinian ID format)."""
    total = 0
    for i, ch in enumerate(body):
        d = int(ch) * (1 if i % 2 == 0 else 2)
        total += d if d < 10 else d - 9
    return (10 - total % 10) % 10


def fake_id(valid=True) -> str:
    body = f"{random.choice([4, 8, 9])}{random.randint(0, 9999999):07d}"
    chk = luhn_check_digit(body)
    if not valid:
        chk = (chk + random.randint(1, 9)) % 10
    return body + str(chk)


def fake_phone() -> str:
    num = f"5{random.choice([6, 9])}{random.randint(0, 9999999):07d}"
    style = random.random()
    if style < 0.45:
        return "0" + num                     # 0599xxxxxx
    if style < 0.70:
        return "+970" + num                  # +970599xxxxxx
    if style < 0.85:
        return "00970 " + num[:2] + " " + num[2:]   # spaced
    if style < 0.985:
        return "972" + num                   # 972 prefix
    return num[:6]                           # truncated / invalid


# --------------------------------------------------------------------------------------
# 1. Site register
# --------------------------------------------------------------------------------------
sites = []
counter, loc_counter = {}, {}
for n in range(30):
    code, gov, loc, lat, lon = LOCALITIES[n % len(LOCALITIES)]
    counter[code] = counter.get(code, 0) + 1
    loc_counter[loc] = loc_counter.get(loc, 0) + 1
    site_id = f"ITCN-{code}-{counter[code]:03d}"
    hh = random.randint(60, 420)
    sites.append({
        "site_id": site_id,
        "site_name": f"{loc} Site {loc_counter[loc]}",
        "governorate_code": code,
        "governorate": gov,
        "locality": loc,
        "site_type": random.choices(SITE_TYPES, weights=[5, 3, 2])[0],
        "latitude": round(lat + np.random.normal(0, 0.006), 6),
        "longitude": round(lon + np.random.normal(0, 0.006), 6),
        "households": hh,
        "individuals": int(hh * np.random.uniform(5.2, 6.8)),
        "field_engineer": random.choice(FIELD_ENGINEERS),
        "start_date": (date(2026, 8, 3) + timedelta(days=7 * random.randint(0, 3))).isoformat(),
    })
sites_df = pd.DataFrame(sites)
# keep a clean reference copy; inject coordinate errors into the "as submitted" copy
sites_raw = sites_df.copy()
i = 4
lat_i, lon_i = sites_raw.at[i, "latitude"], sites_raw.at[i, "longitude"]
sites_raw.at[i, "latitude"], sites_raw.at[i, "longitude"] = lon_i, lat_i   # swapped lat/lon
sites_raw.loc[11, "latitude"] = 3.144521        # missing digit
sites_raw.loc[17, "longitude"] = None            # missing
sites_raw.loc[22, "governorate"] = "Dier Al Balah"   # inconsistent spelling
sites_raw.loc[25, "site_type"] = "collective centre "  # case / whitespace
sites_df.to_csv(REF / "site_register_master.csv", index=False)
sites_raw.to_csv(RAW / "site_register_submitted.csv", index=False)

# --------------------------------------------------------------------------------------
# 2. Targets + weekly field-engineer reports (cumulative achieved)
# --------------------------------------------------------------------------------------
weeks = [date(2026, 8, 31) + timedelta(days=7 * k) for k in range(10)]  # 10 reporting weeks
targets, field_rows, contractor_rows = [], [], []
for s in sites:
    comps = random.sample(list(COMPONENTS), k=random.randint(3, 5))
    if "Shelter" not in comps:
        comps.append("Shelter")
    for comp in comps:
        ind, unit, (lo, hi) = COMPONENTS[comp]
        target = random.randint(lo, hi)
        targets.append({"site_id": s["site_id"], "component": comp, "indicator": ind,
                        "unit": unit, "target": target,
                        "contractor": random.choice(CONTRACTORS)})
        progress = 0.0
        speed = np.random.uniform(0.04, 0.16)
        for wk in weeks:
            progress = min(1.0, progress + max(0, np.random.normal(speed, 0.04)))
            achieved = int(round(target * progress))
            field_rows.append({
                "report_id": None, "week_ending": wk.isoformat(), "site_id": s["site_id"],
                "component": comp, "indicator": ind, "unit": unit,
                "achieved_cumulative": achieved, "reported_by": s["field_engineer"],
            })
            # contractor reports the same thing, usually close, sometimes not
            noise = 0
            r = random.random()
            if r < 0.10:
                noise = random.randint(2, max(3, int(target * 0.2)))   # over-reporting
            elif r < 0.15:
                noise = -random.randint(1, max(2, int(target * 0.1)))
            contractor_rows.append({
                "week_ending": wk.isoformat(), "site_id": s["site_id"], "component": comp,
                "achieved_cumulative": max(0, achieved + noise),
                "contractor": targets[-1]["contractor"],
            })

targets_df = pd.DataFrame(targets)
field_df = pd.DataFrame(field_rows)
contractor_df = pd.DataFrame(contractor_rows)

# inject field-report problems
idx = field_df.sample(12, random_state=1).index
field_df.loc[idx[:4], "site_id"] = field_df.loc[idx[:4], "site_id"].str.replace("ITCN-", "itcn ", regex=False)  # format typos
field_df.loc[idx[4:7], "site_id"] = field_df.loc[idx[4:7], "site_id"].str.replace("-0", "-", regex=False)        # missing zero
for j in idx[7:10]:                                                                                  # above target
    t = targets_df[(targets_df.site_id == field_df.at[j, "site_id"]) &
                   (targets_df.component == field_df.at[j, "component"])]["target"]
    if len(t):
        field_df.at[j, "achieved_cumulative"] = int(t.iloc[0] * 1.6)
field_df.loc[idx[10], "achieved_cumulative"] = -5                                                   # negative
field_df.loc[idx[11], "week_ending"] = "2027-03-01"                                                 # future date
dups = field_df.sample(9, random_state=2)                                                           # duplicate submissions
field_df = pd.concat([field_df, dups], ignore_index=True)
field_df = field_df.sample(frac=1, random_state=3).reset_index(drop=True)
field_df["report_id"] = [f"FR-{k:05d}" for k in range(1, len(field_df) + 1)]

targets_df.to_csv(REF / "targets.csv", index=False)
field_df.to_csv(RAW / "field_reports.csv", index=False)
contractor_df.to_csv(RAW / "contractor_reports.csv", index=False)

# --------------------------------------------------------------------------------------
# 3. Household beneficiary registry
# --------------------------------------------------------------------------------------
hh_rows = []
for s in sites:
    n = int(s["households"] * np.random.uniform(0.85, 1.0))
    for k in range(n):
        female = random.random() < 0.27
        first = random.choice(FIRST_F if female else FIRST_M)
        hh_rows.append({
            "hh_id": f"{s['site_id']}-HH{k + 1:04d}",
            "site_id": s["site_id"],
            "head_name": f"{first} {random.choice(FIRST_M)} {random.choice(FAMILY)}",
            "head_gender": "F" if female else "M",
            "id_number": fake_id(valid=random.random() > 0.008),
            "phone": fake_phone(),
            "hh_size": max(1, int(np.random.poisson(5.8))),
            "members_with_disability": int(np.random.binomial(2, 0.12)),
            "registration_date": (date(2026, 8, 1) + timedelta(days=random.randint(0, 90))).isoformat(),
            "consent_given": random.choices(["yes", "no", ""], weights=[985, 8, 7])[0],
        })
hh_df = pd.DataFrame(hh_rows)
# inject: same person registered in two sites, missing gender, zero hh size, gender coding mess
dup_src = hh_df.sample(25, random_state=4)
dup_rows = []
for _, r in dup_src.iterrows():
    other = random.choice([s["site_id"] for s in sites if s["site_id"] != r.site_id])
    rr = r.copy()
    rr["site_id"] = other
    rr["hh_id"] = f"{other}-HH9{random.randint(100, 999)}"
    dup_rows.append(rr)
hh_df = pd.concat([hh_df, pd.DataFrame(dup_rows)], ignore_index=True)
hh_df.loc[hh_df.sample(15, random_state=5).index, "head_gender"] = ""
hh_df.loc[hh_df.sample(20, random_state=6).index, "head_gender"] = random.choice(["female", "Male", "m"])
hh_df.loc[hh_df.sample(6, random_state=7).index, "hh_size"] = 0
hh_df = hh_df.sample(frac=1, random_state=8).reset_index(drop=True)
hh_df.to_csv(RAW / "beneficiary_registry.csv", index=False)

# --------------------------------------------------------------------------------------
# 4. Cash-for-work attendance
# --------------------------------------------------------------------------------------
workers = []
for w in range(1, 241):
    s = random.choice(sites)
    workers.append({"worker_id": f"CFW-{w:04d}", "site_id": s["site_id"],
                    "id_number": fake_id(), "gender": random.choices(["M", "F"], [80, 20])[0],
                    "task": random.choice(CFW_TASKS)})
workers_df = pd.DataFrame(workers)
att = []
for wk in workers:
    start = date(2026, 9, 1) + timedelta(days=random.randint(0, 20))
    for d in range(random.randint(10, 40)):
        day = start + timedelta(days=d)
        if day.weekday() == 4:   # Friday off
            continue
        att.append({"worker_id": wk["worker_id"], "site_id": wk["site_id"], "date": day.isoformat(),
                    "hours": random.choices([6, 4, 0], [88, 9, 3])[0], "task": wk["task"]})
att_df = pd.DataFrame(att)
att_df = pd.concat([att_df, att_df.sample(40, random_state=9)], ignore_index=True)    # double entry
att_df.loc[att_df.sample(8, random_state=10).index, "hours"] = 14                   # impossible hours
att_df.loc[att_df.sample(10, random_state=11).index, "worker_id"] = "CFW-9" + pd.Series(range(10)).astype(str).str.zfill(3).values  # unknown workers
att_df = att_df.sample(frac=1, random_state=12).reset_index(drop=True)
workers_df.to_csv(REF / "cfw_workers.csv", index=False)
att_df.to_csv(RAW / "cfw_attendance.csv", index=False)

print("Sites:", len(sites_df), "| Field reports:", len(field_df), "| Contractor reports:",
      len(contractor_df), "| Households:", len(hh_df), "| CfW workers:", len(workers_df),
      "| Attendance rows:", len(att_df))
