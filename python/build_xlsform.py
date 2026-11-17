"""
build_xlsform.py
----------------
Builds kobo/ITCN_Site_Monitoring_XLSForm.xlsx - a bilingual (English/Arabic) KoboToolbox /
ODK form for weekly field-engineer site-monitoring visits. Features: cascading
governorate -> site selection, GPS constrained to the Gaza Strip, skip logic per component,
cross-field validation (installed <= target, female workers <= total), calculations,
and a consent gate before photos.
"""
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

ROOT = Path(__file__).resolve().parents[1]
sites = pd.read_csv(ROOT / "outputs" / "clean" / "site_register_clean.csv")

EN, AR = "label::English (en)", "label::Arabic (ar)"
survey_cols = ["type", "name", EN, AR, "hint::English (en)", "required", "relevant", "constraint",
               "constraint_message::English (en)", "constraint_message::Arabic (ar)", "calculation",
               "appearance", "choice_filter"]


def q(type_, name, en="", ar="", **kw):
    row = {"type": type_, "name": name, EN: en, AR: ar}
    row.update({("hint::English (en)" if k == "hint" else
                 "constraint_message::English (en)" if k == "msg_en" else
                 "constraint_message::Arabic (ar)" if k == "msg_ar" else k): v for k, v in kw.items()})
    return row


GAZA = ("selected-at(., 0) >= 31.21 and selected-at(., 0) <= 31.60 and "
        "selected-at(., 1) >= 34.20 and selected-at(., 1) <= 34.57")
survey = [
    q("start", "start"), q("end", "end"), q("today", "today"), q("deviceid", "deviceid"),
    q("note", "intro", "**ITCN weekly site-monitoring visit.** Complete one form per site per week.",
      "**زيارة المتابعة الأسبوعية للموقع.** عبّئ استمارة واحدة لكل موقع أسبوعياً."),
    q("select_one engineer", "engineer", "Field engineer code", "رمز المهندس الميداني", required="yes"),
    q("date", "visit_date", "Date of visit", "تاريخ الزيارة", required="yes",
      constraint=". <= today()", msg_en="Visit date cannot be in the future.", msg_ar="لا يمكن أن يكون تاريخ الزيارة في المستقبل."),
    q("select_one governorate", "governorate", "Governorate", "المحافظة", required="yes"),
    q("select_one site", "site_id", "Site", "الموقع", required="yes", choice_filter="governorate=${governorate}",
      appearance="minimal"),
    q("geopoint", "gps", "Record GPS at the site entrance", "سجّل الإحداثيات عند مدخل الموقع", required="yes",
      hint="Wait until accuracy is below 10 m.", constraint=GAZA,
      msg_en="GPS point is outside the Gaza Strip - re-capture.", msg_ar="النقطة خارج قطاع غزة - أعد التسجيل."),
    q("select_multiple component", "components", "Components with activity this week",
      "المكونات التي فيها نشاط هذا الأسبوع", required="yes"),

    q("begin_group", "grp_shelter", "Shelter - Relief Housing Units", "الإيواء - الوحدات السكنية المؤقتة",
      relevant="selected(${components}, 'shelter')"),
    q("integer", "rhu_target", "RHU target for this site", "العدد المستهدف من الوحدات", required="yes",
      constraint=". > 0 and . <= 500", msg_en="Target must be 1-500.", msg_ar="الهدف بين 1 و 500."),
    q("integer", "rhu_installed", "RHUs installed (cumulative)", "الوحدات المركّبة (تراكمي)", required="yes",
      constraint=". >= 0 and . <= ${rhu_target}", msg_en="Installed cannot exceed target.",
      msg_ar="لا يمكن أن يتجاوز المركّب العدد المستهدف."),
    q("integer", "rhu_occupied", "RHUs occupied by households", "الوحدات المسكونة", required="yes",
      constraint=". >= 0 and . <= ${rhu_installed}", msg_en="Occupied cannot exceed installed.",
      msg_ar="لا يمكن أن يتجاوز المسكون عدد المركّب."),
    q("calculate", "rhu_pct", calculation="round(${rhu_installed} div ${rhu_target} * 100, 1)"),
    q("note", "rhu_pct_note", "Progress: ${rhu_pct}% of target", "نسبة الإنجاز: ${rhu_pct}% من الهدف"),
    q("end_group", "grp_shelter_end"),

    q("begin_group", "grp_wash", "Water & sanitation", "المياه والصرف الصحي",
      relevant="selected(${components}, 'wash')"),
    q("integer", "water_points_functional", "Functional water points", "نقاط المياه العاملة", required="yes",
      constraint=". >= 0 and . <= 50"),
    q("select_one quality", "water_quality", "Observed water-point condition", "حالة نقاط المياه", required="yes"),
    q("end_group", "grp_wash_end"),

    q("begin_group", "grp_elec", "Electricity", "الكهرباء", relevant="selected(${components}, 'electricity')"),
    q("integer", "solar_units", "Solar lighting units installed (cumulative)", "وحدات الإنارة الشمسية المركّبة",
      required="yes", constraint=". >= 0 and . <= 300"),
    q("end_group", "grp_elec_end"),

    q("begin_group", "grp_cfw", "Cash-for-work", "النقد مقابل العمل", relevant="selected(${components}, 'cfw')"),
    q("integer", "cfw_workers", "Workers present today", "عدد العمال الحاضرين اليوم", required="yes",
      constraint=". >= 0 and . <= 200"),
    q("integer", "cfw_female", "of whom women", "منهم نساء", required="yes",
      constraint=". >= 0 and . <= ${cfw_workers}", msg_en="Cannot exceed total workers.",
      msg_ar="لا يمكن أن يتجاوز إجمالي العمال."),
    q("calculate", "cfw_female_pct", calculation="if(${cfw_workers} > 0, round(${cfw_female} div ${cfw_workers} * 100, 0), 0)"),
    q("end_group", "grp_cfw_end"),

    q("select_multiple issue", "issues", "Implementation issues observed", "مشاكل التنفيذ الملاحظة", required="yes",
      constraint="not(selected(., 'none') and count-selected(.) > 1)",
      msg_en="'None' cannot be combined with other answers.", msg_ar="لا يمكن اختيار 'لا يوجد' مع خيارات أخرى."),
    q("text", "issue_details", "Describe the issues and agreed follow-up", "صف المشاكل والمتابعة المتفق عليها",
      required="yes", relevant="not(selected(${issues}, 'none'))", appearance="multiline"),
    q("select_one yesno", "photo_consent", "Do people appear in the photo, and have they given consent?",
      "هل يظهر أشخاص في الصورة، وهل وافقوا على التصوير؟", hint="Choose 'No people / consent given' only if true.",
      required="yes"),
    q("image", "site_photo", "Photo of works", "صورة للأعمال", relevant="${photo_consent} = 'yes'",
      appearance="new"),
    q("note", "end_note", "Thank you. Submit the form before leaving the site.", "شكراً. أرسل الاستمارة قبل مغادرة الموقع."),
]

choices = [("engineer", f"FE-{i:02d}", f"FE-{i:02d}", f"FE-{i:02d}", "") for i in range(1, 7)]
govs = {"NG": ("North Gaza", "شمال غزة"), "GZ": ("Gaza", "غزة"), "DB": ("Deir al-Balah", "دير البلح"),
        "KY": ("Khan Younis", "خانيونس"), "RF": ("Rafah", "رفح")}
choices += [("governorate", c.lower(), en, ar, "") for c, (en, ar) in govs.items()]
choices += [("site", r.site_id.lower().replace("-", "_"), f"{r.site_id} - {r.site_name}",
             f"{r.site_id} - {r.site_name}", r.governorate_code.lower()) for r in sites.itertuples()]
choices += [("component", n, en, ar, "") for n, en, ar in [
    ("shelter", "Shelter / RHUs", "الإيواء / الوحدات"), ("wash", "Water & sanitation", "المياه والصرف الصحي"),
    ("electricity", "Electricity", "الكهرباء"), ("roads", "Roads & access", "الطرق والوصول"),
    ("facility", "Community facility", "المرافق المجتمعية"), ("waste", "Solid waste", "النفايات الصلبة"),
    ("cfw", "Cash-for-work", "النقد مقابل العمل")]]
choices += [("quality", n, en, ar, "") for n, en, ar in [
    ("good", "Good", "جيدة"), ("fair", "Fair - minor repairs", "مقبولة - إصلاحات بسيطة"),
    ("poor", "Poor - not usable", "سيئة - غير صالحة")]]
choices += [("issue", n, en, ar, "") for n, en, ar in [
    ("delay", "Contractor delay", "تأخير المقاول"), ("materials", "Material shortage", "نقص المواد"),
    ("access", "Access constraints", "صعوبات الوصول"), ("complaint", "Community complaint", "شكوى مجتمعية"),
    ("quality", "Quality of works", "جودة الأعمال"), ("none", "None", "لا يوجد")]]
choices += [("yesno", "yes", "No people in photo / consent given", "لا أشخاص في الصورة / تم أخذ الموافقة", ""),
            ("yesno", "no", "People present without consent - skip photo", "يوجد أشخاص بدون موافقة - لا تصوّر", "")]

wb = Workbook()
ws = wb.active; ws.title = "survey"
ws.append(survey_cols)
for r in survey:
    ws.append([r.get(c, "") for c in survey_cols])
wc = wb.create_sheet("choices")
wc.append(["list_name", "name", EN, AR, "governorate"])
for r in choices:
    wc.append(list(r))
wst = wb.create_sheet("settings")
wst.append(["form_title", "form_id", "version", "default_language", "style"])
wst.append(["ITCN Weekly Site Monitoring", "itcn_site_monitoring", "2026111601", "English (en)", "pages"])
for sh in wb.worksheets:
    for c in sh[1]:
        c.font = Font(name="Arial", bold=True); c.fill = PatternFill("solid", fgColor="D9E1F2")
    for col in sh.columns:
        sh.column_dimensions[col[0].column_letter].width = 28
dest = ROOT / "kobo" / "ITCN_Site_Monitoring_XLSForm.xlsx"
dest.parent.mkdir(exist_ok=True)
wb.save(dest)
print("saved", dest, "| questions:", len(survey), "| choices:", len(choices))
