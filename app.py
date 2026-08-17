import uuid
from datetime import datetime, date, time, timedelta
from pathlib import Path
import pandas as pd
import streamlit as st

try:
    import gspread
    from google.oauth2.service_account import Credentials
except Exception:
    gspread = None
    Credentials = None

APP_DIR = Path(__file__).parent
LOCAL_XLSX = APP_DIR / "data" / "vehicle_gate_pass_data.xlsx"
st.set_page_config(page_title="Vehicle Gate Pass", page_icon="🚐", layout="wide")

SHEETS = {
    "Departments": ["department","manager_username","manager_name","manager_password"],
    "Drivers": ["driver_username","driver_name","driver_password","active"],
    "Vehicles": ["vehicle_number","vehicle_type","active"],
    "GatePasses": [
        "request_id","created_at","requisitioner","department","contact","companions",
        "destination","purpose","travel_date","departure_time","return_date","return_time",
        "vehicle_type","vehicle_number","driver_username","driver_name","status",
        "manager_username","manager_name","manager_approved_at","hr_approved_at",
        "security_released_at","security_name","start_mileage","end_mileage",
        "driver_started_at","driver_completed_at","security_verified_at",
        "security_notes","rejection_reason"],
    "ApprovalAudit": ["audit_id","request_id","timestamp","actor_username","actor_name","actor_role","action","remarks"],
}

DEFAULT_DEPARTMENTS = [
    ["Administration","admin_manager","Administration Manager","Admin@123"],
    ["Finance","finance_manager","Finance Manager","Finance@123"],
    ["HR","hr_manager","HR Manager","HR@123"],
    ["Sales","sales_manager","Sales Manager","Sales@123"],
    ["Operations","operations_manager","Operations Manager","Operations@123"],
    ["IT","it_manager","IT Manager","IT@123"],
]
DEFAULT_DRIVERS = [
    ["driver01","Driver One","Driver@123","Yes"],
    ["driver02","Driver Two","Driver@123","Yes"],
    ["driver03","Driver Three","Driver@123","Yes"],
]
DEFAULT_VEHICLES = [["VAN-01","Van","Yes"],["VAN-02","Van","Yes"],["VAN-03","Van","Yes"]]

def blank_data():
    return {name: pd.DataFrame(columns=cols) for name, cols in SHEETS.items()}

def initialize_workbook():
    if LOCAL_XLSX.exists():
        return
    data = blank_data()
    data["Departments"] = pd.DataFrame(DEFAULT_DEPARTMENTS, columns=SHEETS["Departments"])
    data["Drivers"] = pd.DataFrame(DEFAULT_DRIVERS, columns=SHEETS["Drivers"])
    data["Vehicles"] = pd.DataFrame(DEFAULT_VEHICLES, columns=SHEETS["Vehicles"])
    with pd.ExcelWriter(LOCAL_XLSX, engine="openpyxl") as writer:
        for name, df in data.items():
            df.to_excel(writer, sheet_name=name, index=False)

def load_data():
    initialize_workbook()
    raw = pd.read_excel(LOCAL_XLSX, sheet_name=None, dtype=str)
    data = blank_data()
    for name, cols in SHEETS.items():
        df = raw.get(name, pd.DataFrame(columns=cols)).fillna("")
        for c in cols:
            if c not in df.columns:
                df[c] = ""
        data[name] = df[cols].astype(str)
    return data

def save_data(data):
    with pd.ExcelWriter(LOCAL_XLSX, engine="openpyxl") as writer:
        for name, df in data.items():
            df.to_excel(writer, sheet_name=name, index=False)

def append_record(sheet, record):
    data = load_data()
    data[sheet] = pd.concat([
        data[sheet],
        pd.DataFrame([{c: str(record.get(c, "")) for c in SHEETS[sheet]}])
    ], ignore_index=True)
    save_data(data)

def update_record(sheet, key_col, key_value, updates):
    data = load_data()
    idxs = data[sheet].index[data[sheet][key_col].astype(str) == str(key_value)].tolist()
    if not idxs:
        return False
    for k, v in updates.items():
        data[sheet].loc[idxs[0], k] = str(v)
    save_data(data)
    return True

def google_configured():
    try:
        return bool(st.secrets.get("google_sheet_name")) and bool(st.secrets.get("google_service_account"))
    except Exception:
        return False

def sync_to_google():
    if not google_configured() or gspread is None or Credentials is None:
        return False, "Google Sheets is not configured."
    info = dict(st.secrets["google_service_account"])
    scopes = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    client = gspread.authorize(creds)
    ss = client.open(st.secrets["google_sheet_name"])
    data = load_data()
    for name, df in data.items():
        try:
            ws = ss.worksheet(name)
        except Exception:
            ws = ss.add_worksheet(title=name, rows=1000, cols=max(10, len(df.columns)+2))
        ws.clear()
        ws.update([df.columns.tolist()] + df.fillna("").astype(str).values.tolist())
    return True, "All workbook tabs synchronized to one Google Spreadsheet."

def parse_dt(d, t):
    return datetime.combine(pd.to_datetime(d).date(), pd.to_datetime(t).time())

def overlap(a_start, a_end, b_start, b_end):
    return a_start < b_end and a_end > b_start

def vehicle_available(data, vehicle, td, dep, rd, ret):
    start, end = parse_dt(td, dep), parse_dt(rd, ret)
    if end <= start:
        return False
    for _, r in data["GatePasses"].iterrows():
        if r["vehicle_number"] != vehicle or r["status"] in {"Rejected by Manager","Rejected by HR","Cancelled"}:
            continue
        if not all(r[x] for x in ["travel_date","departure_time","return_date","return_time"]):
            continue
        try:
            if overlap(start, end, parse_dt(r["travel_date"],r["departure_time"]), parse_dt(r["return_date"],r["return_time"])):
                return False
        except Exception:
            pass
    return True

def available_vehicles(data, vtype, td, dep, rd, ret):
    v = data["Vehicles"]
    v = v[(v["vehicle_type"] == vtype) & (v["active"].str.lower() == "yes")]
    return [r["vehicle_number"] for _, r in v.iterrows()
            if vehicle_available(data,r["vehicle_number"],td,dep,rd,ret)]

def slots():
    out=[]; cur=datetime.combine(date.today(),time(5,0))
    end=datetime.combine(date.today(),time(23,30))
    while cur<=end:
        out.append(cur.strftime("%H:%M")); cur += timedelta(minutes=30)
    return out

def audit(rid, user, action, remarks=""):
    append_record("ApprovalAudit", {
        "audit_id":"AUD-"+uuid.uuid4().hex[:8].upper(),"request_id":rid,
        "timestamp":datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "actor_username":user.get("username",""),"actor_name":user.get("name",""),
        "actor_role":user.get("role",""),"action":action,"remarks":remarks})

def show_request(r):
    a,b=st.columns(2)
    with a:
        st.write(f"**Requisitioner:** {r['requisitioner']}")
        st.write(f"**Department:** {r['department']}")
        st.write(f"**Destination:** {r['destination']}")
        st.write(f"**Purpose:** {r['purpose']}")
        st.write(f"**Companions:** {r['companions']}")
    with b:
        st.write(f"**Vehicle:** {r['vehicle_number']} ({r['vehicle_type']})")
        st.write(f"**Trip:** {r['travel_date']} {r['departure_time']} → {r['return_date']} {r['return_time']}")
        st.write(f"**Driver:** {r['driver_name'] or 'Not assigned'}")
        st.write(f"**Status:** `{r['status']}`")

def employee_portal():
    st.title("🚐 Vehicle Request")
    st.caption("No password required for requisitioners.")
    data=load_data()
    depts=data["Departments"]["department"].tolist()
    types=sorted(data["Vehicles"]["vehicle_type"].unique().tolist()) or ["Van"]
    with st.form("request"):
        c1,c2=st.columns(2)
        name=c1.text_input("Your name *"); dept=c2.selectbox("Department *",depts)
        contact=c1.text_input("Contact number"); companions=c2.text_input("Person(s) travelling with you")
        destination=st.text_input("Destination *"); purpose=st.text_area("Purpose *")
        c1,c2=st.columns(2)
        td=c1.date_input("Travel date",min_value=date.today()); rd=c2.date_input("Return date",min_value=date.today())
        c1,c2=st.columns(2)
        vt=c1.selectbox("Vehicle type",types); dep=c2.selectbox("Departure time",slots())
        depdt=datetime.combine(td,datetime.strptime(dep,"%H:%M").time())
        ropts=[x for x in slots() if datetime.combine(rd,datetime.strptime(x,"%H:%M").time())>depdt]
        ret=c1.selectbox("Expected return time",ropts or slots())
        av=available_vehicles(data,vt,td,dep,rd,ret)
        if av:
            vehicle=c2.selectbox("Available vehicle",av)
            st.success(f"{len(av)} vehicle(s) available for this exact period.")
        else:
            vehicle=""; c2.error("No vehicle is available for this selected date/time.")
        submit=st.form_submit_button("Submit Request",type="primary")
        if submit:
            if not name or not destination or not purpose:
                st.error("Complete all required fields.")
            elif not av or vehicle not in av:
                st.error("The selected vehicle/time is no longer available.")
            else:
                manager=data["Departments"][data["Departments"]["department"]==dept].iloc[0]
                rid="VGP-"+datetime.now().strftime("%Y%m%d")+"-"+uuid.uuid4().hex[:6].upper()
                append_record("GatePasses",{
                    "request_id":rid,"created_at":datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "requisitioner":name,"department":dept,"contact":contact,"companions":companions,
                    "destination":destination,"purpose":purpose,"travel_date":td,"departure_time":dep,
                    "return_date":rd,"return_time":ret,"vehicle_type":vt,"vehicle_number":vehicle,
                    "status":"Pending Department Manager","manager_username":manager["manager_username"],
                    "manager_name":manager["manager_name"]})
                audit(rid,{"username":"self-service","name":name,"role":"Requisitioner"},"REQUEST_SUBMITTED")
                st.success(f"{rid} submitted to {manager['manager_name']}.")
                st.rerun()
    st.subheader("Check request status")
    rid=st.text_input("Request ID")
    if rid:
        m=data["GatePasses"][data["GatePasses"]["request_id"]==rid]
        if m.empty: st.warning("Request not found.")
        else:
            r=m.iloc[0]; st.info(f"{r['request_id']} — {r['status']}")

def manager_login():
    d=load_data()["Departments"]; st.subheader("Department Manager Login")
    dept=st.selectbox("Department",d["department"].tolist())
    u=st.text_input("Manager username"); p=st.text_input("Manager password",type="password")
    if st.button("Login",type="primary"):
        m=d[(d.department==dept)&(d.manager_username==u)&(d.manager_password==p)]
        if not m.empty:
            r=m.iloc[0]; st.session_state.user={"role":"Department Manager","username":r.manager_username,"name":r.manager_name,"department":r.department}; st.rerun()
        st.error("Invalid manager credentials.")

def driver_login():
    d=load_data()["Drivers"]; st.subheader("Driver Login")
    u=st.text_input("Driver username"); p=st.text_input("Driver password",type="password")
    if st.button("Login",type="primary"):
        m=d[(d.driver_username==u)&(d.driver_password==p)&(d.active.str.lower()=="yes")]
        if not m.empty:
            r=m.iloc[0]; st.session_state.user={"role":"Driver","username":r.driver_username,"name":r.driver_name}; st.rerun()
        st.error("Invalid or inactive driver account.")

def simple_login(role,label,key,pdefault):
    st.subheader(label); p=st.text_input("Password",type="password")
    if st.button("Login",type="primary"):
        if p==st.secrets.get(key,pdefault):
            st.session_state.user={"role":role,"username":key,"name":label}; st.rerun()
        st.error("Invalid password.")

def manager_portal():
    st.title("Department Manager Approval"); d=load_data()
    p=d["GatePasses"][(d["GatePasses"].status=="Pending Department Manager")&(d["GatePasses"].department==st.session_state.user["department"])]
    if p.empty: st.info("No pending requests."); return
    rid=st.selectbox("Request",p.request_id.tolist()); r=p[p.request_id==rid].iloc[0]; show_request(r)
    if st.button("Approve",type="primary"):
        d=load_data()
        if not vehicle_available(d,r.vehicle_number,r.travel_date,r.departure_time,r.return_date,r.return_time):
            st.error("Vehicle is no longer available."); return
        update_record("GatePasses","request_id",rid,{"status":"Pending HR","manager_approved_at":datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
        audit(rid,st.session_state.user,"MANAGER_APPROVED"); st.rerun()
    reason=st.text_input("Rejection reason")
    if st.button("Reject") and reason:
        update_record("GatePasses","request_id",rid,{"status":"Rejected by Manager","rejection_reason":reason})
        audit(rid,st.session_state.user,"MANAGER_REJECTED",reason); st.rerun()

def hr_portal():
    st.title("HR Approval"); d=load_data(); p=d["GatePasses"][d["GatePasses"].status=="Pending HR"]
    if p.empty: st.info("No pending HR approvals."); return
    rid=st.selectbox("Request",p.request_id.tolist()); r=p[p.request_id==rid].iloc[0]; show_request(r)
    if st.button("Approve HR",type="primary"):
        update_record("GatePasses","request_id",rid,{"status":"Pending Security","hr_approved_at":datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
        audit(rid,st.session_state.user,"HR_APPROVED"); st.rerun()
    reason=st.text_input("Rejection reason")
    if st.button("Reject HR") and reason:
        update_record("GatePasses","request_id",rid,{"status":"Rejected by HR","rejection_reason":reason})
        audit(rid,st.session_state.user,"HR_REJECTED",reason); st.rerun()

def driver_portal():
    st.title("Driver Mileage"); d=load_data()
    p=d["GatePasses"][(d["GatePasses"].driver_username==st.session_state.user["username"])&(d["GatePasses"].status.isin(["Vehicle Released","Trip In Progress"]))]
    if p.empty: st.info("No assigned vehicle trips."); return
    rid=st.selectbox("Trip",p.request_id.tolist()); r=p[p.request_id==rid].iloc[0]; show_request(r)
    start=st.number_input("Start mileage",min_value=0.0,step=0.1,value=float(r.start_mileage or 0))
    if r.status=="Vehicle Released" and st.button("Record Start Mileage",type="primary"):
        update_record("GatePasses","request_id",rid,{"start_mileage":start,"driver_started_at":datetime.now().strftime("%Y-%m-%d %H:%M:%S"),"status":"Trip In Progress"})
        audit(rid,st.session_state.user,"DRIVER_STARTED",f"Start mileage {start}"); st.rerun()
    if r.status=="Trip In Progress":
        end=st.number_input("End mileage",min_value=0.0,step=0.1,value=float(r.end_mileage or 0))
        if st.button("Record End Mileage",type="primary"):
            if end<start: st.error("End mileage cannot be lower than start mileage.")
            else:
                update_record("GatePasses","request_id",rid,{"end_mileage":end,"driver_completed_at":datetime.now().strftime("%Y-%m-%d %H:%M:%S"),"status":"Pending Security Verification"})
                audit(rid,st.session_state.user,"DRIVER_COMPLETED",f"End mileage {end}"); st.rerun()

def security_portal():
    st.title("Security Verification & Gate Control"); d=load_data()
    p=d["GatePasses"][d["GatePasses"].status=="Pending Security"]
    if not p.empty:
        st.subheader("HR-approved requests")
        rid=st.selectbox("Request to release",p.request_id.tolist(),key="release"); r=p[p.request_id==rid].iloc[0]; show_request(r)
        dr=d["Drivers"]; dr=dr[dr.active.str.lower()=="yes"]
        driver=st.selectbox("Assign driver",dr.driver_username.tolist(),format_func=lambda x:dr[dr.driver_username==x].iloc[0].driver_name)
        notes=st.text_area("Security release notes")
        if st.button("Verify & Release Vehicle",type="primary"):
            x=dr[dr.driver_username==driver].iloc[0]
            update_record("GatePasses","request_id",rid,{"driver_username":x.driver_username,"driver_name":x.driver_name,"security_name":st.session_state.user["name"],"security_released_at":datetime.now().strftime("%Y-%m-%d %H:%M:%S"),"security_notes":notes,"status":"Vehicle Released"})
            audit(rid,st.session_state.user,"SECURITY_RELEASED",f"Driver {x.driver_name}"); st.rerun()
    else: st.info("No HR-approved requests waiting for release.")
    st.divider(); st.subheader("Completed trips awaiting final verification")
    p=d["GatePasses"][d["GatePasses"].status=="Pending Security Verification"]
    if p.empty: st.info("No completed trips awaiting verification."); return
    rid=st.selectbox("Completed trip",p.request_id.tolist(),key="verify"); r=p[p.request_id==rid].iloc[0]; show_request(r)
    st.write(f"Start mileage: {r.start_mileage}"); st.write(f"End mileage: {r.end_mileage}")
    notes=st.text_area("Verification notes",key="vn")
    if st.button("Verify & Approve Trip",type="primary"):
        update_record("GatePasses","request_id",rid,{"security_verified_at":datetime.now().strftime("%Y-%m-%d %H:%M:%S"),"security_name":st.session_state.user["name"],"security_notes":notes,"status":"Trip Completed - Security Approved"})
        audit(rid,st.session_state.user,"SECURITY_FINAL_APPROVAL",notes); st.rerun()

def admin_portal():
    st.title("Administration"); d=load_data()
    st.dataframe(d["GatePasses"],use_container_width=True,hide_index=True)
    st.download_button("Download one Excel workbook",LOCAL_XLSX.read_bytes(),"vehicle_gate_pass_data.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if google_configured() and st.button("Sync entire workbook to Google Sheets"):
        ok,msg=sync_to_google(); (st.success if ok else st.error)(msg)
    for x in ["Departments","Drivers","Vehicles"]: 
        st.subheader(x); st.dataframe(d[x],use_container_width=True,hide_index=True)

if "user" not in st.session_state:
    st.title("🚐 Company Vehicle Gate Pass")
    role=st.radio("Continue as",["Employee / Requisitioner","Department Manager","HR Manager","Security","Driver","Administration"])
    if role=="Employee / Requisitioner": employee_portal()
    elif role=="Department Manager": manager_login()
    elif role=="HR Manager": simple_login("HR Manager","HR Manager","hr_password","HR@123")
    elif role=="Security": simple_login("Security","Security","security_password","Security@123")
    elif role=="Driver": driver_login()
    else: simple_login("Administration","Administration","admin_password","Admin@123")
    st.stop()

u=st.session_state.user
with st.sidebar:
    st.write(f"**{u['name']}**")
    st.write(f"Role: **{u['role']}**")
    if u.get("department"): st.write(f"Department: **{u['department']}**")
    if st.button("Logout"): del st.session_state["user"]; st.rerun()

if u["role"]=="Department Manager": manager_portal()
elif u["role"]=="HR Manager": hr_portal()
elif u["role"]=="Security": security_portal()
elif u["role"]=="Driver": driver_portal()
else: admin_portal()
