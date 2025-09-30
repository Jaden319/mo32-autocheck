# v7.6 – Fix DuplicateWidgetID, keep Weather/Shift + Loose Gear
import io, os, sys, importlib
from datetime import date, datetime, timedelta
import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Inches
from fpdf import FPDF

APP_VER = "v7.6"
st.set_page_config(page_title="MO32 Crane Compliance - Auto Check", layout="wide")
st.title("MO32 Crane Compliance - Auto Check")
st.caption("For Stevedores made by a stevedore, Example but not complete")

TODAY = date.today()
DATE_FORMATS = ("%Y-%m-%d","%d/%m/%Y","%d-%m-%Y")

def asciiize(s):
    if s is None: return ""
    trans = {"–":"-","—":"-","‑":"-","•":"*","·":"*","“":'"',"”":'"',"‘":"'", "’":"'", "…":"...", "°":" deg ","×":"x","✓":"OK","\u00a0":" "}
    out = str(s)
    for k,v in trans.items(): out = out.replace(k,v)
    try:
        out.encode("latin-1"); return out
    except UnicodeEncodeError:
        return out.encode("latin-1","ignore").decode("latin-1")

CHECK_COLUMNS = [
    "Crane #","Vessel Name","IMO","Crane Make/Model","Serial Number","SWL (t)",
    "Install/Commission Date","Last 5-year Proof Test Date","Last Annual Thorough Exam Date",
    "Annual Exam By (Competent/Responsible Person)","Certificate of Test # (AMSA 365/642/etc)",
    "Certificate Current? (Y/N)","Register of MHE Onboard? (Y/N)","Pre-use Visual Exam OK? (Y/N)",
    "Rigging Plan/Drawings Onboard? (Y/N)","Controls layout labelled & accessible? (Y/N)",
    "Limit switches operational? (Y/N)","Brakes operational? (Y/N)","Operator visibility adequate? (Y/N)",
    "Visibility: Shift (Day/Evening/Night)","Visibility: Weather conditions",
    "Weather protection at winch/controls? (Y/N)","Access/escape to cabin compliant? (Y/N)","Notes / Defects",
    "Loose Gear: Hook/Block Serial Number","Loose Gear: Hook SWL (t)","Loose Gear: Certificate Number",
    "Loose Gear: Last Inspection/Proof Date","Loose Gear: Notes"
]

GUIDANCE = [
    ("5-year proof load test interval", "MO32 Sch.3 2(2)(a)"),
    ("12-month thorough exam interval", "MO32 Sch.3 2(2)(b), 2(5)"),
    ("Certificates current / approved forms", "MO32 s.23"),
    ("Register of MHE kept onboard (maintenance & repair log)", "MO32 s.25"),
    ("Pre-use visual exam before operation", "MO32 s.22(2)(c)"),
    ("Rigging plan/drawings onboard", "MO32 Sch.6 Div.2 cl.1"),
    ("Controls & brakes (incl. limit switches)", "MO32 Sch.6 Div.3; Sch.3 cl.4(3)"),
    ("Weather protection at controls", "MO32 Sch.6 cl.19"),
    ("Loose Gear inspection & traceability", "Good practice; tie to crane SWL"),
]

RISK_KEYWORDS = {
  "Controls layout labelled & accessible? (Y/N)": ["sloppy","sticky","not labelled","label missing","hard to reach","unresponsive","stiff"],
  "Brakes operational? (Y/N)": ["drifts","creeps","won't hold","wont hold","weak brake","fade","brake fade","slips"],
  "Limit switches operational? (Y/N)": ["overtravel","limit not working","no cutout","failsafe","upper limit","lower limit"],
  "Operator visibility adequate? (Y/N)": ["blind spot","obstructed","poor visibility","camera not working","wiper","dirty screen"],
  "Weather protection at winch/controls? (Y/N)": ["no canopy","exposed","water ingress","leaks","rain on controls"],
  "Access/escape to cabin compliant? (Y/N)": ["ladder loose","handrail","blocked","trip hazard","missing step"],
  "Rigging Plan/Drawings Onboard? (Y/N)": ["no plan","drawing missing","outdated plan","no rigging plan"],
  "Certificate Current? (Y/N)": ["expired","out of date","no certificate","missing cert"],
  "Register of MHE Onboard? (Y/N)": ["no register","not onboard","out of date","not updated"],
}
LOOSE_GEAR_RISK = ["bent","worn","crack","cracked","deformed","latch missing","latch bent","grooved sheave","scored","corroded","pitted"]
WEATHER_BAD = ["rain","raining","storm","storming","hail","fog","mist","spray","squall","gust","windy","dust","smoke","night","dark","low light","glare"]

def safe_text(v):
    if v is None: return ""
    try:
        import math
        if isinstance(v, float) and (v!=v): return ""
    except Exception: pass
    s = str(v)
    return "" if s.strip().lower() in ("nan","none") else s

def parse_date(v):
    if not v: return None
    if isinstance(v,(pd.Timestamp,datetime)):
        return v.date() if hasattr(v,"date") else v
    s = str(v).strip()
    for fmt in DATE_FORMATS:
        try: return datetime.strptime(s,fmt).date()
        except Exception: pass
    try:
        return pd.to_datetime(float(s), unit="D", origin="1899-12-30").date()
    except Exception:
        return None

def to_float(v):
    try: return float(str(v).strip())
    except Exception: return None

def yn(v): return str(v).strip().upper()=="Y"
def days_since(d): return None if not d else (TODAY-d).days
def days_left_since(d, interval):
    if not d: return None
    from datetime import timedelta
    due = d + timedelta(days=interval)
    return (due - TODAY).days

def contradiction_notes_check(row):
    notes = (safe_text(row.get("Notes / Defects"))+" "+safe_text(row.get("Loose Gear: Notes"))+" "+safe_text(row.get("Visibility: Weather conditions"))).lower()
    findings = []
    for field, words in RISK_KEYWORDS.items():
        tick = str(row.get(field,"")).strip().upper()
        if not notes or tick not in ("Y","N"): continue
        if tick=="Y" and any(w in notes for w in words):
            hits = [w for w in words if w in notes][:3]
            findings.append(f"{field}: Y but notes mention {', '.join(hits)}")
        if tick=="N" and any(x in notes for x in ["ok now","temporary fix","workaround","still usable"]):
            findings.append(f"{field}: N but notes imply workaround/operation")
    vis = str(row.get("Operator visibility adequate? (Y/N)","")).strip().upper()
    shift = safe_text(row.get("Visibility: Shift (Day/Evening/Night)")).lower()
    if vis=="Y" and (shift=="night" or any(w in notes for w in WEATHER_BAD)):
        findings.append("Visibility = Y but conditions indicate low visibility (night/adverse weather).")
    return findings

def evidence_prompts(row):
    prompts = []
    if yn(row.get("Certificate Current? (Y/N)")) and not safe_text(row.get("Certificate of Test # (AMSA 365/642/etc)")).strip():
        prompts.append("Certificate marked current but certificate number is blank - add ID/photo.")
    if yn(row.get("Register of MHE Onboard? (Y/N)")) and not safe_text(row.get("Annual Exam By (Competent/Responsible Person)")).strip():
        prompts.append("Register marked onboard - add last entry details/competent or responsible person.")
    if yn(row.get("Rigging Plan/Drawings Onboard? (Y/N)")):
        notes = safe_text(row.get("Notes / Defects")).lower()
        if "plan" in notes and "rev" not in notes:
            prompts.append("Rigging plan onboard - add drawing ID/revision/date in notes.")
    if yn(row.get("Certificate Current? (Y/N)")) and not safe_text(row.get("Loose Gear: Certificate Number")).strip():
        prompts.append("Main certificate current but loose gear cert # blank - add accessory cert reference.")
    return prompts

def evaluate_row(row):
    issues, attention, due_soon = [], [], []
    d5 = parse_date(row.get("Last 5-year Proof Test Date"))
    if not d5 or days_since(d5) > 5*365:
        issues.append("Overdue/missing 5-year proof test (MO32 Sch.3 2(2)(a)).")
    else:
        left = days_left_since(d5, 5*365)
        if left is not None and left <= 90:
            due_soon.append(f"5-year proof test due in {left} days.")
    d12 = parse_date(row.get("Last Annual Thorough Exam Date"))
    if not d12 or days_since(d12) > 365+31:
        issues.append("Overdue/missing annual thorough exam (MO32 Sch.3 2(2)(b), 2(5)).")
    else:
        left = days_left_since(d12, 365)
        if left is not None and left <= 30:
            due_soon.append(f"Annual thorough exam due in {left} days.")
    for field, ref in [
        ("Certificate Current? (Y/N)", "s.23"),
        ("Register of MHE Onboard? (Y/N)", "s.25"),
        ("Pre-use Visual Exam OK? (Y/N)", "s.22(2)(c)"),
        ("Rigging Plan/Drawings Onboard? (Y/N)", "Sch.6 Div.2 cl.1"),
        ("Limit switches operational? (Y/N)", "Sch.3 cl.4(3)"),
        ("Brakes operational? (Y/N)", "Sch.6 Div.3"),
        ("Controls layout labelled & accessible? (Y/N)", "Sch.6 Div.3"),
        ("Operator visibility adequate? (Y/N)", "Sch.6 vis."),
        ("Weather protection at winch/controls? (Y/N)", "Sch.6 cl.19"),
        ("Access/escape to cabin compliant? (Y/N)", "Sch.6 access"),
    ]:
        tick = str(row.get(field,"")).strip().upper()
        if tick not in ("Y","N"):
            attention.append(f"{field} not answered (tick Y or N).")
        elif tick != "Y":
            issues.append(f"{field.replace(' (Y/N)','')}: NOT OK (MO32 {ref}).")
    shift = safe_text(row.get("Visibility: Shift (Day/Evening/Night)")).strip()
    weather = safe_text(row.get("Visibility: Weather conditions")).strip()
    if not shift:
        attention.append("Shift/Lighting not selected (Day/Evening/Night).")
    if not weather:
        attention.append("Weather conditions box is empty (e.g., Raining / Clear / Fog).")
    crane_swl = to_float(row.get("SWL (t)"))
    hook_swl = to_float(row.get("Loose Gear: Hook SWL (t)"))
    if hook_swl is not None and crane_swl is not None and hook_swl > crane_swl:
        issues.append(f"Loose gear SWL ({hook_swl} t) exceeds crane SWL ({crane_swl} t) - mismatch.")
    lg_date = parse_date(row.get("Loose Gear: Last Inspection/Proof Date"))
    if not lg_date or days_since(lg_date) > 365:
        issues.append("Loose gear last inspection/proof >12 months or missing - not compliant.")
    else:
        left = days_left_since(lg_date, 365)
        if left is not None and left <= 30:
            due_soon.append(f"Loose gear inspection due in {left} days.")
    if yn(row.get("Certificate Current? (Y/N)")) and not safe_text(row.get("Loose Gear: Certificate Number")).strip():
        attention.append("Loose gear certificate number blank while main cert is current - add accessory cert ref.")
    contradictions = contradiction_notes_check(row)
    if contradictions:
        attention.extend([f"Notes contradict ticks: {c}" for c in contradictions])
    prompts = evidence_prompts(row)
    if prompts:
        attention.extend(prompts)
    status = "PASS"
    if issues:
        status = "FAIL"
    elif attention or due_soon:
        status = "ATTENTION"
    return status, issues, attention, due_soon

def evaluate(df):
    missing = [c for c in CHECK_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError("Form columns missing: " + ", ".join(missing))
    rows = []
    for _, row in df.iterrows():
        status, issues, attention, due_soon = evaluate_row(row)
        rows.append({
            "Crane #": row.get("Crane #"),
            "Vessel Name": row.get("Vessel Name"),
            "IMO": row.get("IMO"),
            "Serial Number": row.get("Serial Number"),
            "SWL (t)": row.get("SWL (t)"),
            "Shift": row.get("Visibility: Shift (Day/Evening/Night)"),
            "Weather": row.get("Visibility: Weather conditions"),
            "Loose Gear Serial": row.get("Loose Gear: Hook/Block Serial Number"),
            "Loose Gear SWL (t)": row.get("Loose Gear: Hook SWL (t)"),
            "Status": status,
            "Issues (FAIL)": "; ".join(issues) if issues else "",
            "Attention (notes/evidence)": "; ".join(attention) if attention else "",
            "Due soon": "; ".join(due_soon) if due_soon else "",
        })
    return pd.DataFrame(rows)

with st.sidebar:
    st.header("Options")
    want_pdf = st.toggle("Create PDF output", value=True, key="opt_pdf")
    st.divider()
    st.subheader("CSV (optional)")
    if st.button("Download blank CSV", key="btn_blankcsv"):
        df_blank = pd.DataFrame([[i+1] + [""]*(len(CHECK_COLUMNS)-1) for i in range(4)], columns=CHECK_COLUMNS)
        st.download_button("Save blank CSV", df_blank.to_csv(index=False).encode("utf-8"), file_name="Crane_Compliance_MO32_Blank.csv", key="dl_blankcsv")
    csv_up = st.file_uploader("Import CSV (optional)", type=["csv"], key="u_csv")

st.markdown("### Job/Vessel info")
colv1, colv2, colv3 = st.columns(3)
with colv1: vessel = st.text_input("Vessel Name","", key="vessel")
with colv2: imo = st.text_input("IMO","", key="imo")
with colv3: operator = st.text_input("Your name / role","", key="operator")

st.markdown("### Crane checks")
crane_data = []; photos_map = {}; photos_loose_map = {}
for n in [1,2,3,4]:
    with st.expander(f"Crane {n}", expanded=(n==1)):
        c1,c2,c3 = st.columns(3)
        with c1:
            make_model = st.text_input(f"Crane {n} Make/Model", key=f"mm{n}")
            serial = st.text_input(f"Crane {n} Serial Number", key=f"sn{n}")
        with c2:
            swl = st.text_input(f"Crane {n} SWL (t)", key=f"swl{n}")
            install = st.text_input(f"Install/Commission Date (YYYY-MM-DD)", key=f"inst{n}")
        with c3:
            proof = st.text_input("Last 5-year Proof Test Date", key=f"p5{n}")
            annual = st.text_input("Last Annual Thorough Exam Date", key=f"a12{n}")
        c4,c5 = st.columns(2)
        with c4:
            annual_by = st.text_input("Annual Exam By (Competent/Responsible Person)", key=f"by{n}")
            cert_no = st.text_input("Certificate of Test # (AMSA 365/642/etc)", key=f"cert{n}")
        with c5:
            st.markdown("**Y/N items** (tick Y if compliant)")
            y_cert = st.selectbox("Certificate Current? (AMSA 365/642 form)", ["","Y","N"], key=f"yc{n}")
            y_reg  = st.selectbox("Register of MHE Onboard? (Maintenance & repair log)", ["","Y","N"], key=f"yr{n}")
            y_pre  = st.selectbox("Pre-use Visual Exam OK? (before operation)", ["","Y","N"], key=f"yp{n}")
            y_plan = st.selectbox("Rigging Plan/Drawings Onboard? (latest revision available)", ["","Y","N"], key=f"ypl{n}")
            y_ctrl = st.selectbox("Controls labelled & accessible? (labels present, reachable)", ["","Y","N"], key=f"yct{n}")
            y_lim  = st.selectbox("Limit switches operational?", ["","Y","N"], key=f"yl{n}")
            y_brk  = st.selectbox("Brakes operational?", ["","Y","N"], key=f"yb{n}")
            y_vis  = st.selectbox("Operator visibility adequate? (consider lighting & weather)", ["","Y","N"], key=f"yv{n}")
            y_wth  = st.selectbox("Weather protection at controls? (canopy/cover, no ingress)", ["","Y","N"], key=f"yw{n}")
            y_acc  = st.selectbox("Access/escape compliant? (ladder/handrails clear)", ["","Y","N"], key=f"ya{n}")
        w1,w2 = st.columns([1,2])
        with w1:
            shift = st.selectbox("Shift/Lighting", ["","Day","Evening","Night"], key=f"shift{n}")
        with w2:
            wx = st.text_input("Weather conditions (e.g., Raining, Storming, Fog, Clear)", key=f"wx{n}")
        notes = st.text_area("Notes / Defects", key=f"notes{n}", height=100)
        photos = st.file_uploader(f"Crane {n} photos (JPG/PNG; up to 8)", type=["jpg","jpeg","png"], accept_multiple_files=True, key=f"photos{n}")
        photos_map[n] = photos or []
        st.markdown("#### Loose Gear (hook/block)")
        lg1, lg2, lg3 = st.columns(3)
        with lg1:
            lg_serial = st.text_input("Hook/Block Serial Number", key=f"lgsn{n}")
            lg_cert   = st.text_input("Certificate Number", key=f"lgcert{n}")
        with lg2:
            lg_swl = st.text_input("Hook SWL (t)", key=f"lgswl{n}")
            lg_date = st.text_input("Last Inspection/Proof Date (YYYY-MM-DD)", key=f"lgdate{n}")
        with lg3:
            lg_notes = st.text_area("Loose Gear Notes", key=f"lgnotes{n}", height=80)
        photos_loose = st.file_uploader(f"Crane {n} loose gear photos (JPG/PNG; up to 6)", type=["jpg","jpeg","png"], accept_multiple_files=True, key=f"photos_loose{n}")
        photos_loose_map[n] = photos_loose or []
        crane_data.append({
            "Crane #": n, "Vessel Name": vessel, "IMO": imo,
            "Crane Make/Model": make_model, "Serial Number": serial, "SWL (t)": swl,
            "Install/Commission Date": install, "Last 5-year Proof Test Date": proof, "Last Annual Thorough Exam Date": annual,
            "Annual Exam By (Competent/Responsible Person)": annual_by, "Certificate of Test # (AMSA 365/642/etc)": cert_no,
            "Certificate Current? (Y/N)": y_cert, "Register of MHE Onboard? (Y/N)": y_reg, "Pre-use Visual Exam OK? (Y/N)": y_pre,
            "Rigging Plan/Drawings Onboard? (Y/N)": y_plan, "Controls layout labelled & accessible? (Y/N)": y_ctrl,
            "Limit switches operational? (Y/N)": y_lim, "Brakes operational? (Y/N)": y_brk, "Operator visibility adequate? (Y/N)": y_vis,
            "Visibility: Shift (Day/Evening/Night)": shift, "Visibility: Weather conditions": wx,
            "Weather protection at winch/controls? (Y/N)": y_wth, "Access/escape to cabin compliant? (Y/N)": y_acc, "Notes / Defects": notes,
            "Loose Gear: Hook/Block Serial Number": lg_serial, "Loose Gear: Hook SWL (t)": lg_swl, "Loose Gear: Certificate Number": lg_cert,
            "Loose Gear: Last Inspection/Proof Date": lg_date, "Loose Gear: Notes": lg_notes
        })

st.divider()
# SINGLE button row with UNIQUE KEYS
left, mid, right = st.columns([2,1,1])
eval_clicked = left.button("Evaluate & Generate Report", type="primary", use_container_width=True, key="btn_eval")
csv_clicked  = mid.button("Download current inputs as CSV", use_container_width=True, key="btn_csv")
demo_clicked = right.button("Try demo (run sample data)", use_container_width=True, key="btn_demo")

if csv_clicked:
    df_now = pd.DataFrame(crane_data, columns=CHECK_COLUMNS)
    st.download_button("Save this CSV now", df_now.to_csv(index=False).encode("utf-8"), file_name="Crane_Compliance_MO32_Current.csv", key="dl_currentcsv")

def build_docx(results_df, df_original, photos_map, photos_loose_map, out_path=None):
    doc = Document()
    doc.add_heading("Crane Compliance Check - MO32", level=1)
    doc.add_paragraph(f"Date: {TODAY.isoformat()}")
    doc.add_paragraph(f"Vessel: {asciiize(df_original['Vessel Name'].iloc[0])}   IMO: {asciiize(df_original['IMO'].iloc[0])}")
    table = doc.add_table(rows=1, cols=12)
    hdr = ["Crane #","Vessel","IMO","Serial #","SWL (t)","Shift","Weather","LG Serial","LG SWL","Status","Issues","Attention/Due"]
    for i,h in enumerate(hdr): table.rows[0].cells[i].text = h
    for _, r in results_df.iterrows():
        c = table.add_row().cells
        c[0].text = asciiize(safe_text(r.get("Crane #"))); c[1].text = asciiize(safe_text(r.get("Vessel Name")))
        c[2].text = asciiize(safe_text(r.get("IMO"))); c[3].text = asciiize(safe_text(r.get("Serial Number")))
        c[4].text = asciiize(safe_text(r.get("SWL (t)"))); c[5].text = asciiize(safe_text(r.get("Shift","")))
        c[6].text = asciiize(safe_text(r.get("Weather",""))); c[7].text = asciiize(safe_text(r.get("Loose Gear Serial","")))
        c[8].text = asciiize(safe_text(r.get("Loose Gear SWL (t)",""))); c[9].text = asciiize(safe_text(r.get("Status")))
        c[10].text= asciiize(safe_text(r.get("Issues (FAIL)"))); c[11].text= asciiize("; ".join([p for p in [safe_text(r.get("Attention (notes/evidence)","")), safe_text(r.get("Due soon",""))] if p]))
    for crane_no in [1,2,3,4]:
        sub = df_original[df_original["Crane #"]==crane_no]
        imgs = photos_map.get(crane_no) or []; imgs_lg = photos_loose_map.get(crane_no) or []
        if sub.empty and not imgs and not imgs_lg: continue
        doc.add_page_break(); doc.add_heading(f"Crane {crane_no} - Detail", level=2)
        if not sub.empty:
            row = sub.iloc[0]
            doc.add_paragraph(f"Make/Model: {safe_text(row.get('Crane Make/Model'))}  |  Serial: {safe_text(row.get('Serial Number'))}  |  SWL: {safe_text(row.get('SWL (t)'))} t")
            doc.add_paragraph(f"Proof test: {safe_text(row.get('Last 5-year Proof Test Date'))}  |  Annual exam: {safe_text(row.get('Last Annual Thorough Exam Date'))}")
            doc.add_paragraph(f"Visibility: Shift={safe_text(row.get('Visibility: Shift (Day/Evening/Night)'))}, Weather='{safe_text(row.get('Visibility: Weather conditions'))}'")
            doc.add_paragraph(f"Loose Gear: Serial {safe_text(row.get('Loose Gear: Hook/Block Serial Number'))}, Hook SWL {safe_text(row.get('Loose Gear: Hook SWL (t)'))} t, Cert {safe_text(row.get('Loose Gear: Certificate Number'))}, Last insp {safe_text(row.get('Loose Gear: Last Inspection/Proof Date'))}")
            notes = safe_text(row.get('Notes / Defects')); lg_notes = safe_text(row.get('Loose Gear: Notes'))
            if notes or lg_notes:
                doc.add_paragraph("Notes/Defects:")
                if notes: doc.add_paragraph(notes)
                if lg_notes: doc.add_paragraph(lg_notes)
        if imgs:
            doc.add_paragraph("Crane Photos:")
            for f in imgs[:8]:
                try:
                    bio = io.BytesIO(f.read()); bio.seek(0); doc.add_picture(bio, width=Inches(2.6))
                except Exception as e:
                    doc.add_paragraph(f"(Could not embed image: {e})")
        if imgs_lg:
            doc.add_paragraph("Loose Gear Photos:")
            for f in imgs_lg[:6]:
                try:
                    bio = io.BytesIO(f.read()); bio.seek(0); doc.add_picture(bio, width=Inches(2.6))
                except Exception as e:
                    doc.add_paragraph(f"(Could not embed image: {e})")
    if out_path: doc.save(out_path); return None
    bio = io.BytesIO(); doc.save(bio); bio.seek(0); return bio

def build_pdf(results_df, df_original, photos_map, photos_loose_map, out_path=None):
    pdf = FPDF(format="A4", unit="mm"); pdf.set_auto_page_break(auto=True, margin=12); pdf.add_page()
    pdf.set_font("Helvetica","B",16); pdf.cell(0,10, asciiize("Crane Compliance Check - MO32"), ln=1)
    pdf.set_font("Helvetica","",11); pdf.cell(0,6, asciiize(f"Date: {TODAY.isoformat()}"), ln=1)
    pdf.cell(0,6, asciiize(f"Vessel: {safe_text(df_original['Vessel Name'].iloc[0])}   IMO: {safe_text(df_original['IMO'].iloc[0])}"), ln=1); pdf.ln(2)
    headers = ["Crane #","Vessel","IMO","Serial #","SWL (t)","Shift","Weather","LG Serial","LG SWL","Status","Issues","Attention/Due"]
    widths  = [12,26,22,26,14,18,28,26,16,16,40,40]
    pdf.set_font("Helvetica","B",7)
    for h,w in zip(headers,widths): pdf.cell(w,6, asciiize(h), border=1)
    pdf.ln(6); pdf.set_font("Helvetica","",7)
    for _, r in results_df.iterrows():
        vals = [safe_text(r["Crane #"]), safe_text(r["Vessel Name"]), safe_text(r["IMO"]), safe_text(r["Serial Number"]), safe_text(r["SWL (t)"]),
                safe_text(r.get("Shift","")), safe_text(r.get("Weather","")), safe_text(r.get("Loose Gear Serial","")), safe_text(r.get("Loose Gear SWL (t)","")),
                safe_text(r["Status"]), asciiize(safe_text(r["Issues (FAIL)"])), asciiize("; ".join([p for p in [safe_text(r["Attention (notes/evidence)"]), safe_text(r["Due soon"])] if p]))]
        y_max = pdf.get_y()
        for v,w in zip(vals,widths):
            x0,y0 = pdf.get_x(), pdf.get_y()
            pdf.multi_cell(w,4, asciiize(v), border=1)
            y_max = max(y_max, pdf.get_y())
            pdf.set_xy(x0+w, y0)
        pdf.set_y(y_max)
    pdf.add_page(); pdf.set_font("Helvetica","B",13); pdf.cell(0,7, asciiize("Checks & Clauses"), ln=1)
    pdf.set_font("Helvetica","",10)
    for text,ref in GUIDANCE: pdf.multi_cell(0,5, asciiize(f"- {text} — {ref}"))
    for crane_no in [1,2,3,4]:
        sub = df_original[df_original["Crane #"]==crane_no]
        imgs = photos_map.get(crane_no) or []; imgs_lg = photos_loose_map.get(crane_no) or []
        if sub.empty and not imgs and not imgs_lg: continue
        pdf.add_page(); pdf.set_font("Helvetica","B",12); pdf.cell(0,7, asciiize(f"Crane {crane_no} - Detail"), ln=1)
        if not sub.empty:
            row = sub.iloc[0]; pdf.set_font("Helvetica","",10)
            pdf.multi_cell(0,5, asciiize(f"Make/Model: {safe_text(row.get('Crane Make/Model'))}  |  Serial: {safe_text(row.get('Serial Number'))}  |  SWL: {safe_text(row.get('SWL (t)'))} t"))
            pdf.multi_cell(0,5, asciiize(f"Proof test: {safe_text(row.get('Last 5-year Proof Test Date'))}  |  Annual exam: {safe_text(row.get('Last Annual Thorough Exam Date'))}"))
            pdf.multi_cell(0,5, asciiize(f"Visibility: Shift={safe_text(row.get('Visibility: Shift (Day/Evening/Night)'))}, Weather='{safe_text(row.get('Visibility: Weather conditions'))}'"))
            pdf.multi_cell(0,5, asciiize(f"Loose Gear: Serial {safe_text(row.get('Loose Gear: Hook/Block Serial Number'))}, Hook SWL {safe_text(row.get('Loose Gear: Hook SWL (t)'))} t, Cert {safe_text(row.get('Loose Gear: Certificate Number'))}, Last insp {safe_text(row.get('Loose Gear: Last Inspection/Proof Date'))}"))
            notes = safe_text(row.get("Notes / Defects")); lg_notes = safe_text(row.get("Loose Gear: Notes"))
            if notes or lg_notes:
                pdf.set_font("Helvetica","B",10); pdf.cell(0,5, asciiize("Notes/Defects:"), ln=1)
                pdf.set_font("Helvetica","",10)
                if notes: pdf.multi_cell(0,5, asciiize(notes))
                if lg_notes: pdf.multi_cell(0,5, asciiize(lg_notes))
        def add_images(imgs, label, prefix):
            if not imgs: return
            pdf.set_font("Helvetica","B",10); pdf.cell(0,5, asciiize(label), ln=1)
            pdf.set_font("Helvetica","",10)
            x0 = pdf.get_x(); y0 = pdf.get_y(); x = x0; y = y0; max_h = 0
            for i, f in enumerate(imgs[:8]):
                try:
                    data = f.read(); fn = f"tmp_{prefix}_{crane_no}_{i}.jpg"
                    with open(fn,"wb") as imgf: imgf.write(data)
                    pdf.image(fn, x=x, y=y, w=65); new_h = 48; max_h = max(max_h, new_h)
                    if (i % 2)==1: y += max_h + 3; x = x0; max_h = 0
                    else: x += 70
                except Exception as e:
                    pdf.multi_cell(0,5, asciiize(f"(Could not embed image: {e})"))
            pdf.ln(5)
        add_images(imgs, "Crane Photos:", "cr")
        add_images(imgs_lg, "Loose Gear Photos:", "lg")
    if out_path: pdf.output(out_path); return None
    out = pdf.output(dest="S").encode("latin1", errors="ignore"); return io.BytesIO(out)

def save_case(results_df, df_original, photos_map, photos_loose_map):
    base = "mo32_cases"; os.makedirs(base, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S"); case_dir = os.path.join(base, f"case_{stamp}"); os.makedirs(case_dir, exist_ok=True)
    df_original.to_csv(os.path.join(case_dir,"inputs.csv"), index=False)
    results_df.to_csv(os.path.join(case_dir,"results.csv"), index=False)
    for label, mapping in [("crane",photos_map), ("loose",photos_loose_map)]:
        for k, files in mapping.items():
            if not files: continue
            pdir = os.path.join(case_dir, f"crane_{k}_{label}_photos"); os.makedirs(pdir, exist_ok=True)
            for i,f in enumerate(files):
                try:
                    data = f.read()
                    with open(os.path.join(pdir, f"photo_{i+1}.jpg"), "wb") as imgf: imgf.write(data)
                except Exception: pass
    build_docx(results_df, df_original, photos_map, photos_loose_map, out_path=os.path.join(case_dir,"MO32_Crane_Compliance_Report.docx"))
    build_pdf(results_df, df_original, photos_map, photos_loose_map, out_path=os.path.join(case_dir,"MO32_Crane_Compliance_Report.pdf"))
    return case_dir

# Actions
if demo_clicked:
    demo_df = pd.DataFrame([
        [1,"MV Example","9526722","NMF DKII","CRN-123","45","2019-05-10","2022-05-08","2025-06-01","Joe Bloggs (Comp)","AMSA365-111","Y","Y","Y","Y","Y","Y","Y","Y","Night","Raining",
         "Y","Y","Controls slightly sloppy at low speed",
         "LG-001","40","LGCERT-789","2025-06-15","hook latch slightly bent"]
    ], columns=CHECK_COLUMNS)
    try:
        out_df = evaluate(demo_df)
        st.subheader("Results (PASS/ATTENTION/FAIL) - Demo")
        st.dataframe(out_df, use_container_width=True)
        docx_io = build_docx(out_df, demo_df, {1:[],2:[],3:[],4:[]}, {1:[],2:[],3:[],4:[]})
        st.download_button("Download Word report (.docx) - Demo", docx_io.getvalue(), file_name="MO32_Crane_Compliance_Report_DEMO.docx", key="dl_docx_demo")
        pdf_io = build_pdf(out_df, demo_df, {1:[],2:[],3:[],4:[]}, {1:[],2:[],3:[],4:[]})
        st.download_button("Download PDF report (.pdf) - Demo", pdf_io.getvalue(), file_name="MO32_Crane_Compliance_Report_DEMO.pdf", key="dl_pdf_demo")
        st.info("Demo used sample data for Crane 1 only (includes Weather/Shift and Loose Gear).")
    except Exception as e:
        st.error(f"Error during demo evaluation: {e}")

if eval_clicked:
    df_input = pd.DataFrame(crane_data, columns=CHECK_COLUMNS)
    try:
        out_df = evaluate(df_input)
        st.subheader("Results (PASS/ATTENTION/FAIL)")
        st.dataframe(out_df, use_container_width=True)
        st.success("Evaluation complete. Download your reports below.")
        docx_io = build_docx(out_df, df_input, photos_map, photos_loose_map)
        st.download_button("Download Word report (.docx)", docx_io.getvalue(), file_name="MO32_Crane_Compliance_Report.docx", key="dl_docx_real")
        pdf_io = build_pdf(out_df, df_input, photos_map, photos_loose_map)
        st.download_button("Download PDF report (.pdf)", pdf_io.getvalue(), file_name="MO32_Crane_Compliance_Report.pdf", key="dl_pdf_real")
        case_dir = save_case(out_df, df_input, photos_map, photos_loose_map)
        st.info(f"Saved a copy of this submission to: {case_dir}")
    except Exception as e:
        st.error(f"Error during evaluation: {e}")
