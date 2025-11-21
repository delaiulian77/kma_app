import json, os, ast, smtplib, bcrypt
from datetime import datetime, timedelta
import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from io import BytesIO
from email.message import EmailMessage
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import Table, TableStyle, Paragraph, SimpleDocTemplate, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_RIGHT
import gspread
from google.oauth2.service_account import Credentials


import gspread
from google.oauth2.service_account import Credentials

DB_SHEETS = {
    "Users": ["FullName","PasswordHash","Email","IsActive"],
    "Equipment": ["Type","Brand","Model","Serial","Notes"],
    "Templates": ["Template","Type","Brand","Model"],
    "TemplateItems": ["Template","Item","Instruction"],
    "Inspections": ["Timestamp","User","Action","Type","Brand","Model","Serial","ResultsJSON","Comment","NextDate","PdfPath","Recipients"],
    "Logins": ["Timestamp","User","Action","Equipment","NextDate"],
}

def _gcp_creds():
    gcp_info = dict(st.secrets["gcp"])   # same as in your test app
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.readonly",  # drive not used, but ok
    ]
    return Credentials.from_service_account_info(gcp_info, scopes=scopes)

@st.cache_resource
def _gsheet_client():
    creds = _gcp_creds()
    gc = gspread.authorize(creds)
    sheet_id = st.secrets["app"]["spreadsheet_id"]
    return gc.open_by_key(sheet_id)

def read_sheet(name: str) -> pd.DataFrame:
    """Read a tab from the KMA_DB Google Sheet into a DataFrame."""
    sh = _gsheet_client()
    ws = sh.worksheet(name)
    records = ws.get_all_records()
    df = pd.DataFrame(records)
    # ensure all expected columns exist, even for empty sheets
    for col in DB_SHEETS[name]:
        if col not in df.columns:
            df[col] = ""
    return df[DB_SHEETS[name]]

def write_sheet(name: str, df: pd.DataFrame) -> None:
    """Overwrite a tab in the KMA_DB Google Sheet from a DataFrame."""
    sh = _gsheet_client()
    ws = sh.worksheet(name)
    df = df.copy()
    df = df.fillna("")
    values = [df.columns.tolist()] + df.astype(str).values.tolist()
    ws.clear()
    ws.update(values)


def add_user(fullname, password, email=""):
    users = read_sheet("Users")
    if (users["FullName"].str.lower().str.strip() == fullname.lower().strip()).any():
        return False, "User already exists."
    salt = bcrypt.gensalt()
    pw_hash = bcrypt.hashpw(password.encode("utf-8"), salt).decode("utf-8")
    new = {"FullName": fullname.strip(), "PasswordHash": pw_hash, "Email": email, "IsActive": True}
    users = pd.concat([users, pd.DataFrame([new])], ignore_index=True)
    write_sheet("Users", users)
    return True, "User created."

def authenticate(fullname, password):
    users = read_sheet("Users")
    row = users[users["FullName"].str.lower().str.strip() == fullname.lower().strip()]
    if row.empty:
        return False, "USER_NOT_FOUND"
    if not bool(row.iloc[0]["IsActive"]):
        return False, "INACTIVE"
    ok = bcrypt.checkpw(password.encode("utf-8"), row.iloc[0]["PasswordHash"].encode("utf-8"))
    return (True, "") if ok else (False, "BAD_PASSWORD")

def upsert_equipment(rec):
    eq = read_sheet("Equipment")

    def norm(x):  # normalize a single value
        return str(x or "").strip().lower()

    def sser(s):  # normalize a pandas Series (handles NaN / numbers)
        return s.fillna("").astype(str).str.strip().str.lower()

    # Ensure the 4 columns exist (in case of a fresh/empty sheet)
    for col in ["Type", "Brand", "Model", "Serial", "Notes"]:
        if col not in eq.columns:
            eq[col] = ""

    # Build a normalized mask (case/space insensitive, robust to numeric cells)
    mask = (
        (sser(eq["Type"])   == norm(rec.get("Type")))   &
        (sser(eq["Brand"])  == norm(rec.get("Brand")))  &
        (sser(eq["Model"])  == norm(rec.get("Model")))  &
        (sser(eq["Serial"]) == norm(rec.get("Serial")))
    )

    if mask.any():
        # Update notes only (no duplicate row)
        idx = eq[mask].index[0]
        eq.loc[idx, "Notes"] = rec.get("Notes", "")
    else:
        # Always write as strings to preserve leading zeros in Serial
        new = {
            "Type":   str(rec.get("Type", "")),
            "Brand":  str(rec.get("Brand", "")),
            "Model":  str(rec.get("Model", "")),
            "Serial": str(rec.get("Serial", "")),
            "Notes":  str(rec.get("Notes", "")),
        }
        eq = pd.concat([eq, pd.DataFrame([new])], ignore_index=True)

    write_sheet("Equipment", eq)


def get_checklist(t, b, m):
    tpls = read_sheet("Templates")
    items = read_sheet("TemplateItems")
    def sser(s):  # safe string-normalization for Series
        return s.fillna("").astype(str).str.strip().str.lower()

    tt, bb, mm = t.strip().lower(), b.strip().lower(), m.strip().lower()
    hit = tpls[(sser(tpls["Type"]) == tt) &
               (sser(tpls["Brand"]) == bb) &
               (sser(tpls["Model"]) == mm)]
    if hit.empty:
        return None, []
    tpl = hit.iloc[0]["Template"]
    rows = items[items["Template"]==tpl][["Item","Instruction"]].to_dict(orient="records")
    return tpl, rows


def gen_pdf(report_dict) -> bytes:
    buf = BytesIO()

    # Document
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=18*mm, rightMargin=18*mm,
        topMargin=16*mm, bottomMargin=14*mm
    )

    # Styles
    h1  = ParagraphStyle('h1', fontName='Helvetica-Bold', fontSize=16, leading=18, spaceAfter=6)
    h2  = ParagraphStyle('h2', fontName='Helvetica-Bold', fontSize=12, leading=14, spaceAfter=6)
    txt = ParagraphStyle('txt', fontName='Helvetica', fontSize=10, leading=12)
    small = ParagraphStyle('small', fontName='Helvetica', fontSize=9, leading=11)
    right = ParagraphStyle('right', parent=small, alignment=TA_RIGHT)

    story = []

    # Header
    story.append(Paragraph("Nordic Maskin & Rail.", h2))
    story.append(Paragraph("Krumtappen 5, 6580 Vamdrup", txt))
    story.append(Paragraph("CVR. 36078405", txt))
    story.append(Spacer(1, 6))

    title = "Kalibreringscertifikat" if report_dict["Action"].lower().startswith("kalibr") else "Service inspektionsrapport"
    story.append(Paragraph(title, h1))
    story.append(Spacer(1, 6))

    # Equipment fields
    fields = [
        ("Udstyr.",  f"{report_dict['Type']}"),
        ("Fabrikat.",f"{report_dict['Brand']}"),
        ("Serie nr.",f"{report_dict['Model']}"),  # swap to Serial if you prefer
        ("Bemærkning", report_dict.get("Comment","")),
    ]
    t = Table(
        [[Paragraph(f"<b>{k}</b>", txt), Paragraph(v or "", txt)] for k,v in fields],
        colWidths=[30*mm, 140*mm], hAlign='LEFT'
    )
    t.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP'), ('BOTTOMPADDING',(0,0),(-1,-1),4)]))
    story.append(t)
    story.append(Spacer(1, 4))

    # 3-column line
    cal_to = report_dict.get("CalibratedTo","")
    ordno  = report_dict.get("OrderNo","")
    kdate  = report_dict["Timestamp"].split(" ")[0]
    tri = [
        [Paragraph("<b>Kalibreret til.</b>", txt),
         Paragraph("<b>Ordre nr.</b>", txt),
         Paragraph("<b>Kalibreret Dato</b>", txt)],
        [Paragraph(cal_to or "", txt),
         Paragraph(ordno or "", txt),
         Paragraph(kdate, txt)]
    ]
    t3 = Table(tri, colWidths=[50*mm, 40*mm, 40*mm])
    t3.setStyle(TableStyle([
        ('BOX',(0,0),(-1,-1),0.3,colors.black),
        ('INNERGRID',(0,0),(-1,-1),0.3,colors.black),
        ('BACKGROUND',(0,0),(-1,0), colors.whitesmoke),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('ALIGN',(0,0),(-1,0),'CENTER'),
    ]))
    story.append(t3)
    story.append(Spacer(1, 8))

    # Checklist
    def map_status(s):
        return {"green":"OK", "yellow":"ATTENTION", "red":"NOT OK"}.get((s or "").lower(), "-")

    cl_rows = []
    for r in report_dict["Results"]:
        left = f"- {r['item']}"
        note = r.get('note') or ""
        if note: left += f" — {note}"
        cl_rows.append([Paragraph(left, txt), Paragraph(map_status(r.get('status')), right)])

    if cl_rows:
        cl = Table(cl_rows, colWidths=[130*mm, 40*mm], hAlign='LEFT')
        cl.setStyle(TableStyle([
            ('VALIGN',(0,0),(-1,-1),'TOP'),
            ('ALIGN',(1,0),(1,-1),'RIGHT'),
            ('BOTTOMPADDING',(0,0),(-1,-1),3),
        ]))
        story.append(cl)
        story.append(Spacer(1, 6))

    # Banedanmark note
    bd_text = ("Udstysr kalibreres if. GAB-Banedanmark anlæg & fornyelse. "
               "General arbejdsbeskrivelse for sporarbejde.(GAB spor) udgave 14 af. "
               "D.4-4-2016 pct. 2.6.1.1")
    story.append(Paragraph(bd_text, small))
    story.append(Spacer(1, 8))

    # Next date + Kontrolleret af
    story.append(Paragraph(f"Næste kontrol dato: {report_dict['NextDate']}", h2))
    story.append(Paragraph(f"Kontrolleret af: {report_dict['User']}", h2))
    story.append(Spacer(1, 14))

    # Footer
    story.append(Paragraph("Revision 03-03-2022  Udarbejdet: SH    Kontrolleret: TJ    Godkendt: DCS", small))

    doc.build(story)
    return buf.getvalue()



def send_email(recipients, subject, body, pdf_bytes, filename):
    """
    Send a PDF as e-mail attachment using SMTP settings from st.secrets['smtp'].
    """
    smtp_cfg = st.secrets["smtp"]  # will raise a clear error if [smtp] is missing

    host         = smtp_cfg["host"]
    port         = int(smtp_cfg.get("port", 587))
    user         = smtp_cfg["user"]
    password     = smtp_cfg["password"]
    sender_email = smtp_cfg.get("sender_email", user)
    sender_name  = smtp_cfg.get("sender_name", "KMA App")

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = f"{sender_name} <{sender_email}>"
    msg["To"] = ", ".join(recipients)
    msg.set_content(body)

    # Attach PDF
    msg.add_attachment(
        pdf_bytes,
        maintype="application",
        subtype="pdf",
        filename=filename,
    )

    # Send
    with smtplib.SMTP(host, port) as s:
        s.starttls()
        s.login(user, password)
        s.send_message(msg)


# ----------------- UI -----------------
st.set_page_config(page_title="KMA — Kalibrering / Service", page_icon="🧰", layout="centered")


if "step" not in st.session_state:
    st.session_state.step = 1
if "user" not in st.session_state:
    st.session_state.user = None
if "action" not in st.session_state:
    st.session_state.action = None
if "selection" not in st.session_state:
    st.session_state.selection = {}
if "results" not in st.session_state:
    st.session_state.results = []

st.title("KMA — Kalibrering / Service")

# ---- Step 1: Login / Sign up ----
if st.session_state.step == 1:
    st.subheader("Login")
    name = st.text_input("Full name")
    pw = st.text_input("Password", type="password")
    c1, c2 = st.columns(2)
    if c1.button("Login", use_container_width=True):
        ok, reason = authenticate(name, pw)
        if ok:
            st.session_state.user = name.strip()
            st.session_state.step = 2
        else:
            st.error({"USER_NOT_FOUND":"User not found","INACTIVE":"User is inactive","BAD_PASSWORD":"Wrong password"}.get(reason, reason))
    if c2.button("Close app", use_container_width=True):
        st.stop()

    st.markdown("---")
    st.subheader("New user")
    new_name = st.text_input("New full name", key="nu_name")
    new_pw = st.text_input("New password", type="password", key="nu_pw")
    new_email = st.text_input("Email (optional)", key="nu_email")
    if st.button("Create account"):
        if not new_name or not new_pw:
            st.warning("Name and password are required.")
        else:
            ok, msg = add_user(new_name, new_pw, new_email)
            (st.success if ok else st.error)(msg)

# ---- Step 2: Choose action ----
elif st.session_state.step == 2:
    st.success(f"Logged in as: {st.session_state.user}")
    st.subheader("Vælg handling")
    c1, c2 = st.columns(2)
    if c1.button("Kalibrering", use_container_width=True):
        st.session_state.action = "Kalibrering"; st.session_state.step = 3
    if c2.button("Service inspektion", use_container_width=True):
        st.session_state.action = "Service inspektion"; st.session_state.step = 3

# ---- Step 3: Equipment selection / add new ----
elif st.session_state.step == 3:
    eq = read_sheet("Equipment")
    types = sorted(eq["Type"].dropna().unique().tolist())

    st.subheader("Udstyr")
    c1, c2 = st.columns([3,1])
    with c1:
        t = st.selectbox("Type", types, index=0 if types else None, key="sel_type")
        brands = sorted(eq[eq["Type"]==t]["Brand"].dropna().unique().tolist()) if t else []
        b = st.selectbox("Mærke", brands, index=0 if brands else None, key="sel_brand")
        models = sorted(eq[(eq["Type"]==t)&(eq["Brand"]==b)]["Model"].dropna().unique().tolist()) if b else []
        m = st.selectbox("Model", models, index=0 if models else None, key="sel_model")
        serials = sorted(eq[(eq["Type"]==t)&(eq["Brand"]==b)&(eq["Model"]==m)]["Serial"].dropna().unique().tolist()) if m else []
        s = st.selectbox("Serienr.", serials, index=0 if serials else None, key="sel_serial")
    with c2:
        st.markdown("**Ny udstyr**")
        nt = st.text_input("Type", key="ne_type")
        nb = st.text_input("Mærke", key="ne_brand")
        nm = st.text_input("Model", key="ne_model")
        ns = st.text_input("Serienr.", key="ne_serial")
        nn = st.text_input("Noter", key="ne_notes")
        if st.button("Gem og vælg"):
            if not nt or not nb or not nm or not ns:
                st.error("Udfyld alle felter.")
            else:
                upsert_equipment({"Type":nt,"Brand":nb,"Model":nm,"Serial":ns,"Notes":nn})
                for k in ("sel_type", "sel_brand", "sel_model", "sel_serial"):
                    st.session_state.pop(k, None)
                st.session_state.selection = {"Type":nt,"Brand":nb,"Model":nm,"Serial":ns}
                st.session_state.step = 4
                st.rerun()

    c3, c4 = st.columns(2)
    if c3.button("Tilbage"):
        st.session_state.step = 2
    if c4.button("Fortsæt"):
        if not (st.session_state.get("sel_type") and st.session_state.get("sel_brand") and st.session_state.get("sel_model") and st.session_state.get("sel_serial")):
            st.warning("Vælg Type, Mærke, Model, Serienr.")
        else:
            st.session_state.selection = {"Type":st.session_state.sel_type,"Brand":st.session_state.sel_brand,"Model":st.session_state.sel_model,"Serial":st.session_state.sel_serial}
            st.session_state.step = 4

# ---- Step 4: Checklist ----
elif st.session_state.step == 4:
    sel = st.session_state.selection
    st.subheader(f"Tjekliste — {sel['Type']} / {sel['Brand']} / {sel['Model']} / {sel['Serial']}")
    tpl, rows = get_checklist(sel["Type"], sel["Brand"], sel["Model"])
    if not rows:
        st.error("Ingen tjekliste for denne kombination. Tilføj i Templates/TemplateItems.")
        if st.button("Tilbage"): st.session_state.step = 3
    else:
        # build checklist widgets
        results = []
        for i, r in enumerate(rows):
            st.write(f"**{r['Item']}**  \n_{r.get('Instruction','')}_")
            col1, col2, col3, col4 = st.columns([1,1,1,3])
            status = col1.radio("Status", ["green","yellow","red"], horizontal=True, key=f"st_{i}")
            note = col4.text_input("Note", key=f"nt_{i}")
            results.append({"item": r["Item"], "instruction": r.get("Instruction",""), "status": status, "note": note})
            st.divider()
        
        cal_to = st.text_input("Kalibreret til", placeholder="f.eks. 150Nm")
        order_no = st.text_input("Ordre nr.", placeholder="f.eks. 1/11")

        comment = st.text_area("Kommentar")
        default_next = (datetime.today() + timedelta(days=365)).date()
        next_date = st.date_input("Næste dato", value=default_next)

                # --- actions ---
   
        c1, c2, c3 = st.columns(3)

        if c1.button("Tilbage"):
            st.session_state.step = 3

        if c2.button("Bekræft og generér rapport"):
            # --- 1) Timestamp & recipients from secrets ---
            ts = datetime.now().strftime("%Y-%m-%d %H:%M")

            app_cfg = st.secrets.get("app", {})
            recipients = app_cfg.get("default_recipients", [])

            if not recipients:
                st.warning(
                    "Ingen modtagere er konfigureret. "
                    "Tilføj [app].default_recipients i Streamlit secrets."
                )

            # --- 2) Build report dict ---
            sel = st.session_state.selection

            report = {
                "Timestamp": ts,
                "User": st.session_state.user,
                "Action": st.session_state.action,
                "Type": sel["Type"],
                "Brand": sel["Brand"],
                "Model": sel["Model"],
                "Serial": sel["Serial"],
                "Results": results,
                "Comment": comment,
                "NextDate": str(next_date),
                "CalibratedTo": cal_to,
                "OrderNo": order_no,
            }

            # --- 3) Generate PDF (always) ---
            pdf_bytes = gen_pdf(report)
            base_name = (
                "Kalibrering"
                if st.session_state.action.lower().startswith("kalibr")
                else "Service"
            )
            pdf_name = (
                f"{base_name}_{sel['Type']}_{sel['Brand']}_"
                f"{sel['Model']}_{sel['Serial']}_"
                f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
            )

            # Optional: still save a copy on disk (useful for local runs)
            pdf_path = os.path.join(os.getcwd(), pdf_name)
            try:
                with open(pdf_path, "wb") as f:
                    f.write(pdf_bytes)
                st.info(f"PDF gemt som: **{pdf_path}**")
            except Exception as e:
                st.warning(f"Kunne ikke gemme PDF lokalt: {e}")
                pdf_path = ""  # avoid logging a wrong path

            # --- 4) Send e-mail automatically (if recipients exist) ---
            if recipients:
                try:
                    send_email(
                        recipients=recipients,
                        subject=(
                            f"Rapport: {st.session_state.action} — "
                            f"{sel['Type']}/{sel['Brand']}/"
                            f"{sel['Model']}/{sel['Serial']}"
                        ),
                        body=(
                            f"Se vedhæftet PDF.\n\n"
                            f"Bruger: {st.session_state.user}\n"
                            f"Dato: {ts}"
                        ),
                        pdf_bytes=pdf_bytes,
                        filename=pdf_name,
                    )
                    st.success("PDF sendt på e-mail.")
                except Exception as e:
                    st.warning(f"E-mail kunne ikke sendes: {e}")

            # --- 5) Log til Inspections + Logins (samme som før) ---
            insp = read_sheet("Inspections")
            row = {
                "Timestamp": ts,
                "User": st.session_state.user,
                "Action": st.session_state.action,
                "Type": sel["Type"],
                "Brand": sel["Brand"],
                "Model": sel["Model"],
                "Serial": sel["Serial"],
                "ResultsJSON": str(results),
                "Comment": comment,
                "NextDate": str(next_date),
                "PdfPath": pdf_path,
                "Recipients": ", ".join(recipients),
            }
            insp = pd.concat([insp, pd.DataFrame([row])], ignore_index=True)
            write_sheet("Inspections", insp)

            log = read_sheet("Logins")
            log_row = {
                "Timestamp": ts,
                "User": st.session_state.user,
                "Action": st.session_state.action,
                "Equipment": (
                    f"{sel['Type']}/{sel['Brand']}/{sel['Model']}/{sel['Serial']}"
                ),
                "NextDate": str(next_date),
            }
            log = pd.concat([log, pd.DataFrame([log_row])], ignore_index=True)
            write_sheet("Logins", log)

            st.session_state.results = results
            st.session_state.step = 5

        if c3.button("Close app"):
            st.stop()
