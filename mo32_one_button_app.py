# mo32_one_button_app.py (v7.3 Auto Check – ASCII-safe PDF, demo fix)
import io, os, json, importlib, sys
from datetime import date, datetime, timedelta
import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Inches
from fpdf import FPDF

APP_VER = "v7.3 (Auto Check, ASCII-safe PDF, demo fix)"
st.set_page_config(page_title="MO32 Crane Compliance - Auto Check", layout="wide")
st.title("MO32 Crane Compliance - Auto Check")
st.caption("Built-in web form for stevedores. Photos, contradictions, due-soon, PDF/DOCX, and auto-save. PDF text is ASCII-sanitised to avoid font issues. Demo no longer edits session_state.")

TODAY = date.today()
DATE_FORMATS = ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y")

def asciiize(s: str) -> str:
    if s is None: return ""
    trans = {
        "–": "-", "—": "-", "‑": "-", "•": "*", "·": "*",
        "“": '"', "”": '"', "‘": "'", "’": "'", "…": "...",
        "°": " deg ", "×": "x", "✓": "OK", "\u00a0": " ",
    }
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
    "Weather protection at winch/controls? (Y/N)","Access/escape to cabin compliant? (Y/N)","Notes / Defects",
]

GUIDANCE = [
    ("5-year proof load test interval", "MO32 Sch.3 2(2)(a)"),
    ("12-month thorough exam interval (+alignment)", "MO32 Sch.3 2(2)(b), 2(5)"),
    ("Certificates current / approved forms", "MO32 s.23"),
    ("Register of MHE kept onboard", "MO32 s.25"),
    ("Pre-use visual exam before operation", "MO32 s.22(2)(c)"),
    ("Rigging plan/drawings onboard (cranes)", "MO32 Sch.6 Div.2 cl.1"),
    ("Controls & brakes (incl. limit switches)", "MO32 Sch.6 Div.3; Sch.3 cl.4(3)"),
    ("Weather protection at winch/controls", "MO32 Sch.6 cl.19"),
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

def safe_text(val):
    if val is None:
        return ""
    try:
        if isinstance(val, float) and pd.isna(val):
            return ""
    except Exception:
        pass
    s = str(val)
    if s.strip().lower() in ("nan","none"):
        return ""
    return s

def parse_date(val):
    if not val: return None
    if isinstance(val, (pd.Timestamp, datetime)):
        return val.date() if hasattr(val, "date") else val
    s = str(val).strip()
    for fmt in DATE_FORMATS:
        try: return datetime.strptime(s, fmt).date()
        except Exception: pass
    try:
        return pd.to_datetime(float(s), unit="D", origin="1899-12-30").date()
    except Exception:
        return None

def yn(val): return str(val).strip().upper() == "Y"
def days_since(d): return None if not d else (TODAY - d).days
def days_left_since(d, interval):
    if not d: return None
    due = d + timedelta(days=interval)
    return (due - TODAY).days

def contradiction_notes_check(row):
    notes = safe_text(row.get("Notes / Defects")).lower()
    findings = []
    for field, words in RISK_KEYWORDS.items():
        yn_val = str(row.get(field,"")).strip().upper()
        if not notes or yn_val not in ("Y","N"): continue
        if yn_val == "Y" and any(w in notes for w in words):
            hits = [w for w in words if w in notes]
            findings.append(f"{field}: ticked Y but notes mention {', '.join(hits[:3])}")
        if yn_val == "N" and any(x in notes for x in ["ok now","temporary fix","workaround","still usable"]):
            findings.append(f"{field}: ticked N but notes imply operation/workaround (not acceptable)")
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
    if not d12 or days_since(d12) > 365 + 31:
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
        if not yn(row.get(field)):
            issues.append(f"{field.replace(' (Y/N)','')}: NOT OK (MO32 {ref}).")
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
    results = []
    for _, row in df.iterrows():
        status, issues, attention, due_soon = evaluate_row(row)
        results.append({
            "Crane #": row.get("Crane #"),
            "Vessel Name": row.get("Vessel Name"),
            "IMO": row.get("IMO"),
            "Serial Number": row.get("Serial Number"),
            "SWL (t)": row.get("SWL (t)"),
            "Status": status,
            "Issues (FAIL)": "; ".join(issues) if issues else "",
            "Attention (notes/evidence)": "; ".join(attention) if attention else "",
            "Due soon": "; ".join(due_soon) if due_soon else "",
        })
    return pd.DataFrame(results)

with st.sidebar:
    st.header("Options")
    st.write("Built-in form. You can still export/import a CSV below if you want.")
    want_pdf = st.toggle("Create PDF output", value=True)
    st.divider()
    st.subheader("CSV (optional)")
    if st.button("Download blank CSV"):
        df_blank = pd.DataFrame([[i+1] + [""]*21 for i in range(4)], columns=CHECK_COLUMNS)
        st.download_button("Save blank CSV", df_blank.to_csv(index=False).encode("utf-8"), file_name="Crane_Compliance_MO32_Blank.csv")
    csv_up = st.file_uploader("Import CSV (optional)", type=["csv"], key="csvup")

st.markdown("### Job/Vessel info")
colv1, colv2, colv3 = st.columns(3)
with colv1:
    vessel = st.text_input("Vessel Name", "")
with colv2:
    imo = st.text_input("IMO", "")
with colv3:
    operator = st.text_input("Your name / role", "")

st.markdown("### Crane checks")
crane_data = []
photos_map = {}

for n in [1,2,3,4]:
    with st.expander(f"Crane {n}", expanded=(n==1)):
        c1, c2, c3 = st.columns(3)
        with c1:
            make_model = st.text_input(f"Crane {n} Make/Model", key=f"mm{n}")
            serial = st.text_input(f"Crane {n} Serial Number", key=f"sn{n}")
        with c2:
            swl = st.text_input(f"Crane {n} SWL (t)", key=f"swl{n}")
            install = st.text_input(f"Install/Commission Date (YYYY-MM-DD)", key=f"inst{n}")
        with c3:
            proof = st.text_input("Last 5-year Proof Test Date", key=f"p5{n}")
            annual = st.text_input("Last Annual Thorough Exam Date", key=f"a12{n}")
        c4, c5 = st.columns(2)
        with c4:
            annual_by = st.text_input("Annual Exam By (Competent/Responsible Person)", key=f"by{n}")
            cert_no = st.text_input("Certificate of Test # (AMSA 365/642/etc)", key=f"cert{n}")
        with c5:
            st.markdown("**Y/N items** (tick Y if compliant)")
            y_cert = st.selectbox("Certificate Current?", ["", "Y", "N"], key=f"yc{n}")
            y_reg = st.selectbox("Register of MHE Onboard?", ["", "Y", "N"], key=f"yr{n}")
            y_pre = st.selectbox("Pre-use Visual Exam OK?", ["", "Y", "N"], key=f"yp{n}")
            y_plan = st.selectbox("Rigging Plan/Drawings Onboard?", ["", "Y", "N"], key=f"ypl{n}")
            y_ctrl = st.selectbox("Controls labelled & accessible?", ["", "Y", "N"], key=f"yct{n}")
            y_lim = st.selectbox("Limit switches operational?", ["", "Y", "N"], key=f"yl{n}")
            y_brk = st.selectbox("Brakes operational?", ["", "Y", "N"], key=f"yb{n}")
            y_vis = st.selectbox("Operator visibility adequate?", ["", "Y", "N"], key=f"yv{n}")
            y_wth = st.selectbox("Weather protection at controls?", ["", "Y", "N"], key=f"yw{n}")
            y_acc = st.selectbox("Access/escape compliant?", ["", "Y", "N"], key=f"ya{n}")
        notes = st.text_area("Notes / Defects", key=f"notes{n}", height=100, placeholder="e.g., controls slightly sloppy at low speed; upper limit switch intermittent")
        photos = st.file_uploader(f"Crane {n} photos (JPG/PNG; up to 8)", type=["jpg","jpeg","png"], accept_multiple_files=True, key=f"photos{n}")
        photos_map[n] = photos or []
        crane_data.append({
            "Crane #": n, "Vessel Name": vessel, "IMO": imo,
            "Crane Make/Model": make_model, "Serial Number": serial, "SWL (t)": swl,
            "Install/Commission Date": install, "Last 5-year Proof Test Date": proof, "Last Annual Thorough Exam Date": annual,
            "Annual Exam By (Competent/Responsible Person)": annual_by, "Certificate of Test # (AMSA 365/642/etc)": cert_no,
            "Certificate Current? (Y/N)": y_cert, "Register of MHE Onboard? (Y/N)": y_reg, "Pre-use Visual Exam OK? (Y/N)": y_pre,
            "Rigging Plan/Drawings Onboard? (Y/N)": y_plan, "Controls layout labelled & accessible? (Y/N)": y_ctrl,
            "Limit switches operational? (Y/N)": y_lim, "Brakes operational? (Y/N)": y_brk, "Operator visibility adequate? (Y/N)": y_vis,
            "Weather protection at winch/controls? (Y/N)": y_wth, "Access/escape to cabin compliant? (Y/N)": y_acc,
            "Notes / Defects": notes,
        })

st.divider()
left, mid, right = st.columns([2,1,1])
with left:
    run_btn = st.button("Evaluate & Generate Report", type="primary", use_container_width=True)
with mid:
    dl_csv = st.button("Download current inputs as CSV", use_container_width=True)
with right:
    demo_btn = st.button("Try demo (run sample data)", use_container_width=True)

if csv_up is not None:
    try:
        df_csv = pd.read_csv(csv_up)
        st.success("CSV loaded into session. You can evaluate it directly even if fields above aren't auto-filled.")
        st.session_state["_csv_loaded_df"] = df_csv
    except Exception as e:
        st.error(f"Could not load CSV: {e}")

def df_from_form():
    return pd.DataFrame(crane_data, columns=CHECK_COLUMNS)

def ensure_case_dir():
    base = "mo32_cases"
    os.makedirs(base, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    case_dir = os.path.join(base, f"case_{stamp}")
    os.makedirs(case_dir, exist_ok=True)
    return case_dir

def build_docx(results_df, df_original, photos_map, out_path=None):
    doc = Document()
    doc.add_heading("Crane Compliance Check - MO32", level=1)
    doc.add_paragraph(f"Date: {TODAY.isoformat()}")
    doc.add_paragraph(f"Vessel: {safe_text(df_original['Vessel Name'].iloc[0])}   IMO: {safe_text(df_original['IMO'].iloc[0])}")
    table = doc.add_table(rows=1, cols=8)
    hdr = ["Crane #","Vessel","IMO","Serial #","SWL (t)","Status","Issues","Attention/Due"]
    for i, h in enumerate(hdr):
        table.rows[0].cells[i].text = h
    for _, r in results_df.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = safe_text(r["Crane #"])
        row_cells[1].text = safe_text(r["Vessel Name"])
        row_cells[2].text = safe_text(r["IMO"])
        row_cells[3].text = safe_text(r["Serial Number"])
        row_cells[4].text = safe_text(r["SWL (t)"])
        row_cells[5].text = safe_text(r["Status"])
        row_cells[6].text = safe_text(r["Issues (FAIL)"])
        row_cells[7].text = "; ".join([p for p in [safe_text(r["Attention (notes/evidence)"]), safe_text(r["Due soon"])] if p])

    for crane_no in [1,2,3,4]:
        sub = df_original[df_original["Crane #"]==crane_no]
        imgs = photos_map.get(crane_no) or []
        if sub.empty and not imgs:
            continue
        doc.add_page_break()
        doc.add_heading(f"Crane {crane_no} - Detail", level=2)
        if not sub.empty:
            row = sub.iloc[0]
            doc.add_paragraph(f"Make/Model: {safe_text(row.get('Crane Make/Model'))}  |  Serial: {safe_text(row.get('Serial Number'))}  |  SWL: {safe_text(row.get('SWL (t)'))} t")
            doc.add_paragraph(f"Proof test: {safe_text(row.get('Last 5-year Proof Test Date'))}  |  Annual exam: {safe_text(row.get('Last Annual Thorough Exam Date'))}")
            notes = safe_text(row.get("Notes / Defects"))
            if notes:
                doc.add_paragraph("Notes/Defects:"); doc.add_paragraph(notes)
        if imgs:
            doc.add_paragraph("Photos:")
            for i, f in enumerate(imgs[:8]):
                try:
                    bio = io.BytesIO(f.read()); bio.seek(0)
                    doc.add_picture(bio, width=Inches(2.6))
                except Exception as e:
                    doc.add_paragraph(f"(Could not embed image: {e})")

    if out_path:
        doc.save(out_path); return None
    else:
        bio = io.BytesIO(); doc.save(bio); bio.seek(0); return bio

def build_pdf(results_df, df_original, photos_map, out_path=None):
    pdf = FPDF(format="A4", unit="mm")
    pdf.set_auto_page_break(auto=True, margin=12)
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 16); pdf.cell(0, 10, asciiize("Crane Compliance Check - MO32"), ln=1)
    pdf.set_font("Helvetica", "", 11); pdf.cell(0, 6, asciiize(f"Date: {TODAY.isoformat()}"), ln=1)
    pdf.cell(0, 6, asciiize(f"Vessel: {safe_text(df_original['Vessel Name'].iloc[0])}   IMO: {safe_text(df_original['IMO'].iloc[0])}"), ln=1)
    pdf.ln(2)

    headers = ["Crane #","Vessel","IMO","Serial #","SWL (t)","Status","Issues","Attention/Due"]
    widths = [18, 35, 22, 35, 16, 18, 48, 48]
    pdf.set_font("Helvetica", "B", 9)
    for h, w in zip(headers, widths):
        pdf.cell(w, 7, asciiize(h), border=1)
    pdf.ln(7); pdf.set_font("Helvetica", "", 8)

    for _, r in results_df.iterrows():
        issues = asciiize(safe_text(r["Issues (FAIL)"]))
        attn  = asciiize("; ".join([p for p in [safe_text(r["Attention (notes/evidence)"]), safe_text(r["Due soon"])] if p]))
        vals = [safe_text(r["Crane #"]), safe_text(r["Vessel Name"]), safe_text(r["IMO"]), safe_text(r["Serial Number"]), safe_text(r["SWL (t)"]), safe_text(r["Status"]), issues, attn]
        x = pdf.get_x(); y = pdf.get_y()
        for v, w in zip(vals, widths):
            start_x = pdf.get_x(); start_y = pdf.get_y()
            pdf.multi_cell(w, 5, asciiize(v), border=1)
            end_y = pdf.get_y()
            pdf.set_xy(start_x + w, start_y)
            y = max(y, end_y)
        pdf.set_y(y)

    pdf.add_page()
    pdf.set_font("Helvetica", "B", 14); pdf.cell(0, 8, asciiize("Checks & Clauses"), ln=1)
    pdf.set_font("Helvetica", "", 10)
    for item, ref in GUIDANCE:
        pdf.multi_cell(0, 5, asciiize(f"- {item} — {ref}"))

    for crane_no in [1,2,3,4]:
        sub = df_original[df_original["Crane #"]==crane_no]
        imgs = photos_map.get(crane_no) or []
        if sub.empty and not imgs: continue
        pdf.add_page()
        pdf.set_font("Helvetica", "B", 12); pdf.cell(0, 7, asciiize(f"Crane {crane_no} - Detail"), ln=1)
        if not sub.empty:
            row = sub.iloc[0]
            pdf.set_font("Helvetica", "", 10)
            pdf.multi_cell(0, 5, asciiize(f"Make/Model: {safe_text(row.get('Crane Make/Model'))}  |  Serial: {safe_text(row.get('Serial Number'))}  |  SWL: {safe_text(row.get('SWL (t)'))} t"))
            pdf.multi_cell(0, 5, asciiize(f"Proof test: {safe_text(row.get('Last 5-year Proof Test Date'))}  |  Annual exam: {safe_text(row.get('Last Annual Thorough Exam Date'))}"))
            notes = safe_text(row.get("Notes / Defects"))
            if notes:
                pdf.set_font("Helvetica", "B", 10); pdf.cell(0, 5, asciiize("Notes/Defects:"), ln=1)
                pdf.set_font("Helvetica", "", 10); pdf.multi_cell(0, 5, asciiize(notes))
        if imgs:
            pdf.set_font("Helvetica", "B", 10); pdf.cell(0, 5, asciiize("Photos:"), ln=1)
            pdf.set_font("Helvetica", "", 10)
            x0 = pdf.get_x(); y0 = pdf.get_y(); x = x0; y = y0; max_h = 0
            for i, f in enumerate(imgs[:8]):
                try:
                    data = f.read()
                    fn = f"tmp_{crane_no}_{i}.jpg"
                    with open(fn, "wb") as imgf: imgf.write(data)
                    pdf.image(fn, x=x, y=y, w=65)
                    new_h = 48; max_h = max(max_h, new_h)
                    if (i % 2) == 1:
                        y += max_h + 3; x = x0; max_h = 0
                    else:
                        x += 70
                except Exception as e:
                    pdf.multi_cell(0, 5, asciiize(f"(Could not embed image: {e})"))
            pdf.ln(5)

    if out_path:
        pdf.output(out_path); return None
    else:
        out = pdf.output(dest="S").encode("latin1", errors="ignore")
        return io.BytesIO(out)

def save_case(results_df, df_original, photos_map):
    base = "mo32_cases"
    os.makedirs(base, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    case_dir = os.path.join(base, f"case_{stamp}")
    os.makedirs(case_dir, exist_ok=True)
    df_original.to_csv(os.path.join(case_dir, "inputs.csv"), index=False)
    results_df.to_csv(os.path.join(case_dir, "results.csv"), index=False)
    for k, files in photos_map.items():
        if not files: continue
        pdir = os.path.join(case_dir, f"crane_{k}_photos"); os.makedirs(pdir, exist_ok=True)
        for i, f in enumerate(files):
            try:
                data = f.read()
                with open(os.path.join(pdir, f"photo_{i+1}.jpg"), "wb") as imgf: imgf.write(data)
            except Exception: pass
    build_docx(results_df, df_original, photos_map, out_path=os.path.join(case_dir, "MO32_Crane_Compliance_Report.docx"))
    build_pdf(results_df, df_original, photos_map, out_path=os.path.join(case_dir, "MO32_Crane_Compliance_Report.pdf"))
    return case_dir

# Actions
if dl_csv:
    df_now = pd.DataFrame(crane_data, columns=CHECK_COLUMNS)
    st.download_button("Save this CSV now", df_now.to_csv(index=False).encode("utf-8"), file_name="Crane_Compliance_MO32_Current.csv")

# Demo path: generate a DataFrame directly (no session_state edits)
if demo_btn:
    demo_df = pd.DataFrame([
        [1,"MV Example","9526722","NMF DKII","CRN-123","45","2019-05-10","2022-05-08","2025-06-01","Joe Bloggs (Comp)","AMSA365-111","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","controls slightly sloppy at low speed"]
    ], columns=CHECK_COLUMNS)
    try:
        out_df = evaluate(demo_df)
        st.subheader("Results (PASS/ATTENTION/FAIL) - Demo")
        st.dataframe(out_df, use_container_width=True)
        docx_io = build_docx(out_df, demo_df, {1:[],2:[],3:[],4:[]})
        st.download_button("Download Word report (.docx) - Demo", docx_io.getvalue(), file_name="MO32_Crane_Compliance_Report_DEMO.docx")
        if want_pdf:
            pdf_io = build_pdf(out_df, demo_df, {1:[],2:[],3:[],4:[]})
            st.download_button("Download PDF report (.pdf) - Demo", pdf_io.getvalue(), file_name="MO32_Crane_Compliance_Report_DEMO.pdf")
        st.info("Demo used sample data for Crane 1 only.")
    except Exception as e:
        st.error(f"Error during demo evaluation: {e}")

if run_btn:
    df_input = pd.DataFrame(crane_data, columns=CHECK_COLUMNS)
    try:
        out_df = evaluate(df_input)
        st.subheader("Results (PASS/ATTENTION/FAIL)")
        st.dataframe(out_df, use_container_width=True)
        st.success("Evaluation complete. Download your reports below.")
        docx_io = build_docx(out_df, df_input, photos_map)
        st.download_button("Download Word report (.docx)", docx_io.getvalue(), file_name="MO32_Crane_Compliance_Report.docx")
        if want_pdf:
            pdf_io = build_pdf(out_df, df_input, photos_map)
            st.download_button("Download PDF report (.pdf)", pdf_io.getvalue(), file_name="MO32_Crane_Compliance_Report.pdf")
        case_dir = save_case(out_df, df_input, photos_map)
        st.info(f"Saved a copy of this submission to: {case_dir}")
    except Exception as e:
        st.error(f"Error during evaluation: {e}")

with st.expander("Diagnostics"):
    import platform
    st.code({
        "app_version": APP_VER,
        "python": sys.version,
        "platform": platform.platform(),
        "pandas": importlib_import_module("pandas").__version__ if True else "2.2.2",
        "streamlit": importlib_import_module("streamlit").__version__ if True else "1.33.0",
        "fpdf2": importlib_import_module("fpdf").__version__ if True else "2.7.9",
    }, language="json")
    st.caption("Each submission is saved under ./mo32_cases/<timestamp> with CSV, results, photos, and DOCX/PDF.")
