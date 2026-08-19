"""Vehicle Gate Pass Management System.

Google Sheets is the only application data store.
Google service-account credentials are read only from Streamlit Secrets.
No Excel/local data file is created by this application.
"""

import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, date, time, timedelta
import pandas as pd
import uuid
import time as time_module

st.set_page_config(page_title="Vehicle Gate Pass", page_icon="🚐", layout="wide")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SHEET_NAME = "vehicle_gate_pass_data"

SHEET_HEADERS = {
    "Users": ["username", "name", "role", "password", "active"],
    "Departments": [
        "department", "manager_username", "manager_name",
        "manager_password", "active"
    ],
    "Drivers": ["driver_username", "driver_name", "password", "active"],
    "Vehicles": [
        "vehicle_number", "vehicle_type", "driver_username",
        "driver_name", "active"
    ],
    "GatePasses": [
        "request_id", "created_at", "requisitioner_name", "department",
        "manager_username", "companions", "duration_minutes", "destination",
        "purpose", "travel_date", "start_time", "end_time", "vehicle_number",
        "driver_username", "driver_name", "status", "manager_decision",
        "manager_approved_by", "manager_approved_at", "hr_decision",
        "hr_approved_by", "hr_approved_at", "security_released_by",
        "security_released_at", "start_mileage", "driver_started_at",
        "end_mileage", "distance_km", "driver_completed_at",
        "security_verified_by", "security_verified_at", "rejection_reason",
    ],
    "ApprovalAudit": [
        "timestamp", "request_id", "username", "role", "action", "remarks"
    ],
}

ACTIVE_STATUSES = {
    "Pending Department Manager",
    "Pending Vehicle Allocation",
    "Pending HR",
    "Pending Security",
    "Vehicle Released",
    "Trip In Progress",
    "Pending Security Verification",
}

ROLES = [
    "Employee / Requisitioner",
    "Department Manager",
    "Vehicle Allocator",
    "HR Manager",
    "Security",
    "Driver",
    "Administration",
]


# ============================================================
# GOOGLE SHEETS
# ============================================================

@st.cache_resource
def get_google_client():
    required = [
        "type", "project_id", "private_key_id", "private_key",
        "client_email", "client_id", "auth_uri", "token_uri",
        "auth_provider_x509_cert_url", "client_x509_cert_url",
    ]
    missing = [key for key in required if key not in st.secrets]
    if missing:
        raise RuntimeError(
            "Missing Google service-account secrets: " + ", ".join(missing)
        )

    info = {key: st.secrets[key] for key in required}
    info["client_id"] = str(info["client_id"])

    if "universe_domain" in st.secrets:
        info["universe_domain"] = st.secrets["universe_domain"]

    credentials = Credentials.from_service_account_info(
        info, scopes=SCOPES
    )
    return gspread.authorize(credentials)


@st.cache_resource
def get_spreadsheet():
    client = get_google_client()
    sheet_id = str(st.secrets.get("google_sheet_id", "")).strip()

    if sheet_id:
        return client.open_by_key(sheet_id)

    sheet_name = (
        str(st.secrets.get("google_sheet_name", SHEET_NAME)).strip()
        or SHEET_NAME
    )
    return client.open(sheet_name)


def _api_call_with_backoff(operation, max_attempts=5):
    for attempt in range(max_attempts):
        try:
            return operation()
        except gspread.exceptions.APIError as exc:
            status = getattr(
                getattr(exc, "response", None), "status_code", None
            )
            if status != 429 or attempt == max_attempts - 1:
                raise
            time_module.sleep(min(2 ** attempt, 16))


@st.cache_resource
def get_or_create_worksheet(name):
    spreadsheet = get_spreadsheet()
    try:
        worksheet = spreadsheet.worksheet(name)
    except gspread.WorksheetNotFound:
        headers = SHEET_HEADERS[name]
        worksheet = spreadsheet.add_worksheet(
            title=name,
            rows=1000,
            cols=max(20, len(headers) + 2),
        )
        worksheet.append_row(headers, value_input_option="USER_ENTERED")
        return worksheet

    try:
        current_headers = worksheet.row_values(1)
    except Exception:
        current_headers = []

    if not current_headers:
        worksheet.append_row(
            SHEET_HEADERS[name], value_input_option="USER_ENTERED"
        )
    return worksheet


READ_CACHE_TTL = 60


@st.cache_data(ttl=READ_CACHE_TTL, show_spinner=False)
def read_sheet(name):
    worksheet = get_or_create_worksheet(name)
    return pd.DataFrame(_api_call_with_backoff(worksheet.get_all_records))


def invalidate_data_cache():
    read_sheet.clear()


def append_row(name, row_dict):
    worksheet = get_or_create_worksheet(name)
    headers = SHEET_HEADERS[name]

    _api_call_with_backoff(
        lambda: worksheet.append_row(
            [row_dict.get(header, "") for header in headers],
            value_input_option="USER_ENTERED",
        )
    )
    invalidate_data_cache()


def update_request(request_id, updates):
    worksheet = get_or_create_worksheet("GatePasses")
    headers = SHEET_HEADERS["GatePasses"]
    records = read_sheet("GatePasses")
    target_row = None

    if not records.empty and "request_id" in records.columns:
        matches = records.index[
            records["request_id"].astype(str).str.strip()
            == str(request_id).strip()
        ].tolist()
        if matches:
            target_row = matches[0] + 2

    if target_row is None:
        read_sheet.clear()
        for row_number, record in enumerate(
            _api_call_with_backoff(worksheet.get_all_records), start=2
        ):
            if str(record.get("request_id", "")).strip() == str(request_id).strip():
                target_row = row_number
                break

    if target_row is None:
        raise ValueError(f"Request {request_id} was not found.")

    payload = []
    for key, value in updates.items():
        if key not in headers:
            continue
        col = headers.index(key) + 1
        payload.append({
            "range": gspread.utils.rowcol_to_a1(target_row, col),
            "values": [["" if value is None else str(value)]],
        })

    if payload:
        _api_call_with_backoff(
            lambda: worksheet.batch_update(payload, raw=False)
        )
    invalidate_data_cache()


def audit(request_id, username, role, action, remarks=""):
    append_row(
        "ApprovalAudit",
        {
            "timestamp": now_str(),
            "request_id": request_id,
            "username": username,
            "role": role,
            "action": action,
            "remarks": remarks,
        },
    )


# ============================================================
# HELPERS
# ============================================================

def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def normalize_bool(value):
    return str(value).strip().lower() in {
        "yes", "true", "1", "active", "y"
    }


def parse_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def parse_time(value):
    if isinstance(value, time):
        return value

    text = str(value).strip()
    for fmt in ("%H:%M", "%H:%M:%S", "%I:%M %p"):
        try:
            return datetime.strptime(text, fmt).time()
        except ValueError:
            pass
    return None


def generate_request_id():
    return (
        "VGP-"
        + datetime.now().strftime("%Y%m%d")
        + "-"
        + uuid.uuid4().hex[:6].upper()
    )


def active_records(dataframe):
    if dataframe.empty or "active" not in dataframe.columns:
        return dataframe
    return dataframe[dataframe["active"].apply(normalize_bool)].copy()


def get_users():
    return active_records(read_sheet("Users"))


def get_departments():
    return active_records(read_sheet("Departments"))


def get_drivers():
    return active_records(read_sheet("Drivers"))


def get_vehicles():
    return active_records(read_sheet("Vehicles"))


def get_gatepasses():
    return read_sheet("GatePasses")


# ============================================================
# AUTHENTICATION
# ============================================================

def authenticate_user(username, password, role):
    users = get_users()
    if users.empty:
        return None

    username = str(username).strip()
    password = str(password).strip()
    role = str(role).strip().lower()

    for _, row in users.iterrows():
        if (
            str(row.get("username", "")).strip() == username
            and str(row.get("password", "")).strip() == password
            and str(row.get("role", "")).strip().lower() == role
            and normalize_bool(row.get("active", "Yes"))
        ):
            return row.to_dict()
    return None


def authenticate_manager(username, password):
    departments = get_departments()
    if departments.empty:
        return None

    username = str(username).strip()
    password = str(password).strip()

    for _, row in departments.iterrows():
        if (
            str(row.get("manager_username", "")).strip() == username
            and str(row.get("manager_password", "")).strip() == password
            and normalize_bool(row.get("active", "Yes"))
        ):
            return row.to_dict()
    return None


def authenticate_driver(username, password):
    drivers = get_drivers()
    if drivers.empty:
        return None

    username = str(username).strip()
    password = str(password).strip()

    for _, row in drivers.iterrows():
        if (
            str(row.get("driver_username", "")).strip() == username
            and str(row.get("password", "")).strip() == password
            and normalize_bool(row.get("active", "Yes"))
        ):
            return row.to_dict()
    return None


# ============================================================
# VEHICLE / FIXED DRIVER LOGIC
# ============================================================

def vehicle_driver(vehicle_number):
    vehicles = get_vehicles()
    drivers = get_drivers()

    for _, vehicle in vehicles.iterrows():
        if str(vehicle.get("vehicle_number", "")).strip() == str(vehicle_number).strip():
            driver_username = str(vehicle.get("driver_username", "")).strip()
            driver_name = str(vehicle.get("driver_name", "")).strip()

            if driver_username and not driver_name:
                for _, driver in drivers.iterrows():
                    if str(driver.get("driver_username", "")).strip() == driver_username:
                        driver_name = str(driver.get("driver_name", "")).strip()
                        break

            return {
                "driver_username": driver_username,
                "driver_name": driver_name,
            }
    return None


def request_interval(row):
    travel_date = parse_date(row.get("travel_date"))
    start_time = parse_time(row.get("start_time"))
    end_time = parse_time(row.get("end_time"))

    if not travel_date or not start_time or not end_time:
        return None, None

    return (
        datetime.combine(travel_date, start_time),
        datetime.combine(travel_date, end_time),
    )


def vehicle_is_available(
    vehicle_number,
    travel_date,
    start_dt,
    end_dt,
    exclude_request_id=None,
):
    """Return True when the vehicle does not overlap another active booking."""
    gatepasses = get_gatepasses()
    if gatepasses.empty:
        return True

    for _, row in gatepasses.iterrows():
        if str(row.get("vehicle_number", "")).strip() != str(vehicle_number).strip():
            continue

        if (
            exclude_request_id
            and str(row.get("request_id", "")).strip()
            == str(exclude_request_id).strip()
        ):
            continue

        if str(row.get("status", "")).strip() not in ACTIVE_STATUSES:
            continue

        row_date = parse_date(row.get("travel_date"))
        if row_date != travel_date:
            continue

        old_start, old_end = request_interval(row)
        if old_start is None or old_end is None:
            continue

        if start_dt < old_end and end_dt > old_start:
            return False

    return True


def available_vehicles_for_interval(
    travel_date,
    start_time,
    duration_minutes,
):
    """Find vehicles available for the exact time requested by the employee."""
    start_dt = datetime.combine(travel_date, start_time)
    end_dt = start_dt + timedelta(minutes=duration_minutes)

    if end_dt.date() != travel_date:
        return []

    result = []

    for _, vehicle in get_vehicles().iterrows():
        vehicle_number = str(vehicle.get("vehicle_number", "")).strip()
        if not vehicle_number:
            continue

        if vehicle_is_available(
            vehicle_number,
            travel_date,
            start_dt,
            end_dt,
        ):
            driver_info = vehicle_driver(vehicle_number) or {}
            result.append({
                "vehicle_number": vehicle_number,
                "vehicle_type": str(
                    vehicle.get("vehicle_type", "")
                ).strip(),
                "driver_username": driver_info.get(
                    "driver_username",
                    str(vehicle.get("driver_username", "")).strip(),
                ),
                "driver_name": driver_info.get(
                    "driver_name",
                    str(vehicle.get("driver_name", "")).strip(),
                ),
            })

    return result


# ============================================================
# SESSION / LOGIN
# ============================================================

def logout():
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()


def login_portal(role):
    st.subheader(f"{role} Login")

    if role == "Department Manager":
        username = st.text_input("Manager Username", key="manager_username_login")
        password = st.text_input(
            "Manager Password",
            type="password",
            key="manager_password_login",
        )

        if st.button("Login", key="manager_login_button"):
            user = authenticate_manager(username, password)
            if user:
                st.session_state.logged_in = True
                st.session_state.role = role
                st.session_state.username = username.strip()
                st.session_state.user = user
                st.rerun()
            else:
                st.error("Invalid manager username or password.")

    elif role == "Driver":
        username = st.text_input("Driver Username", key="driver_username_login")
        password = st.text_input(
            "Driver Password",
            type="password",
            key="driver_password_login",
        )

        if st.button("Login", key="driver_login_button"):
            user = authenticate_driver(username, password)
            if user:
                st.session_state.logged_in = True
                st.session_state.role = role
                st.session_state.username = username.strip()
                st.session_state.user = user
                st.rerun()
            else:
                st.error("Invalid driver username or password.")

    else:
        username = st.text_input(
            "Username",
            key=f"{role}_username_login",
        )
        password = st.text_input(
            "Password",
            type="password",
            key=f"{role}_password_login",
        )

        if st.button("Login", key=f"{role}_login_button"):
            user = authenticate_user(username, password, role)
            if user:
                st.session_state.logged_in = True
                st.session_state.role = role
                st.session_state.username = username.strip()
                st.session_state.user = user
                st.rerun()
            else:
                st.error("Invalid username or password.")


# ============================================================
# EMPLOYEE / REQUISITIONER
# ============================================================

def _store_pending_request(draft):
    st.session_state["pending_vehicle_request"] = draft
    st.session_state["transfer_reminder_snooze_until"] = None


def _clear_pending_request():
    st.session_state.pop("pending_vehicle_request", None)
    st.session_state.pop("transfer_reminder_snooze_until", None)


@st.dialog("⚠️ Vehicle Request Not Yet Transferred", width="medium")
def transfer_reminder_dialog():
    draft = st.session_state.get("pending_vehicle_request")
    if not draft:
        return

    st.warning(
        "You have completed a vehicle request, but it has NOT been "
        "transferred to the company Google Sheet yet."
    )

    st.write(
        f"**Request:** {draft['request_id']}\n\n"
        f"**Date:** {draft['travel_date']}\n\n"
        f"**Required Time:** {draft['start_time']} - {draft['end_time']}\n\n"
        f"**Duration:** {draft['duration_minutes']} minutes"
    )

    st.info(
        "The request is currently stored only in this browser session. "
        "Managers and HR will not see it until you transfer it."
    )

    c1, c2 = st.columns(2)

    with c1:
        if st.button("Transfer Data Now", type="primary", use_container_width=True):
            st.session_state["transfer_from_reminder"] = True
            st.rerun()

    with c2:
        if st.button("Continue Editing", use_container_width=True):
            st.session_state["transfer_reminder_snooze_until"] = (
                datetime.now() + timedelta(seconds=30)
            ).isoformat()
            st.rerun()


def maybe_show_transfer_reminder():
    draft = st.session_state.get("pending_vehicle_request")
    if not draft:
        return

    snooze = st.session_state.get("transfer_reminder_snooze_until")
    if snooze:
        try:
            if datetime.now() < datetime.fromisoformat(snooze):
                return
        except ValueError:
            pass

    transfer_reminder_dialog()


def _transfer_pending_request():
    draft = st.session_state.get("pending_vehicle_request")
    if not draft:
        return False

    if st.session_state.get("_vehicle_transfer_in_progress", False):
        st.warning("The request is already being transferred. Please wait.")
        return False

    st.session_state["_vehicle_transfer_in_progress"] = True

    try:
        read_sheet.clear()
        append_row("GatePasses", draft["row"])

        audit(
            draft["request_id"],
            "employee",
            "Employee / Requisitioner",
            "Request Submitted",
            "Vehicle will be allocated after Department Manager approval.",
        )

        _clear_pending_request()
        st.session_state.pop("transfer_from_reminder", None)
        st.session_state.pop("transfer_reminder_snooze_until", None)
        return True
    finally:
        st.session_state["_vehicle_transfer_in_progress"] = False


def employee_portal():
    st.header("🚐 Vehicle Requisition")

    st.caption(
        "Enter the date and exact time you need the vehicle. "
        "The Vehicle Allocator will select an available vehicle for "
        "your requested time."
    )

    if st.session_state.get("transfer_from_reminder"):
        try:
            with st.spinner("Transferring request..."):
                if _transfer_pending_request():
                    st.success("The request has been transferred to Google Sheets successfully.")
                    st.rerun()
        except Exception as error:
            st.error("The request could not be transferred to Google Sheets.")
            st.exception(error)

    maybe_show_transfer_reminder()

    departments = get_departments()

    if departments.empty or "department" not in departments.columns:
        st.error("No departments are configured in the Departments sheet.")
        return

    department_names = sorted([
        str(v).strip()
        for v in departments["department"].tolist()
        if str(v).strip()
    ])

    with st.form("vehicle_request_form"):
        col1, col2 = st.columns(2)

        with col1:
            requisitioner = st.text_input("Requisitioner Name *")

            department = st.selectbox(
                "Department *",
                department_names,
            )

            companions = st.text_area(
                "Person(s) travelling with you",
                placeholder="Enter names separated by commas",
            )

            destination = st.text_input("Where are you going? *")

            purpose = st.text_area("Purpose of travel *")

        with col2:
            travel_date = st.date_input(
                "Travel Date *",
                min_value=date.today(),
            )

            requested_start_time = st.time_input(
                "Vehicle Required From *",
                value=time(8, 0),
                step=timedelta(minutes=30),
            )

            duration_options = {
                "30 minutes": 30,
                "1 hour": 60,
                "1.5 hours": 90,
                "2 hours": 120,
                "3 hours": 180,
                "4 hours": 240,
                "6 hours": 360,
                "8 hours": 480,
                "10 hours": 600,
                "12 hours": 720,
            }

            duration_label = st.selectbox(
                "Expected duration *",
                list(duration_options.keys()),
            )
            duration_minutes = duration_options[duration_label]

            requested_start_dt = datetime.combine(
                travel_date,
                requested_start_time,
            )
            requested_end_dt = requested_start_dt + timedelta(
                minutes=duration_minutes
            )

            if requested_end_dt.date() == travel_date:
                st.info(
                    f"Vehicle required from **{requested_start_dt.strftime('%H:%M')}** "
                    f"to **{requested_end_dt.strftime('%H:%M')}**."
                )
            else:
                st.error(
                    "The requested duration extends into the next day. "
                    "Please choose an earlier start time or shorter duration."
                )

        st.divider()

        st.info(
            "Your requested date and time will be sent for approval. "
            "After approval, the Vehicle Allocator will select a vehicle "
            "and its permanently assigned driver that are available during "
            "this exact period."
        )

        prepared = st.form_submit_button(
            "Review Request & Prepare Transfer",
            type="primary",
        )

    if prepared:
        if not requisitioner.strip() or not destination.strip() or not purpose.strip():
            st.error("Please complete all required fields.")
            return

        if requested_end_dt.date() != travel_date:
            st.error(
                "The requested vehicle time cannot extend into the next day."
            )
            return

        manager = None
        for _, row in departments.iterrows():
            if str(row.get("department", "")).strip() == department:
                manager = row.to_dict()
                break

        if not manager:
            st.error("No manager is configured for this department.")
            return

        request_id = generate_request_id()

        row = {
            "request_id": request_id,
            "created_at": now_str(),
            "requisitioner_name": requisitioner.strip(),
            "department": department,
            "manager_username": str(
                manager.get("manager_username", "")
            ).strip(),
            "companions": companions.strip(),
            "duration_minutes": duration_minutes,
            "destination": destination.strip(),
            "purpose": purpose.strip(),
            "travel_date": travel_date.isoformat(),
            "start_time": requested_start_time.strftime("%H:%M"),
            "end_time": requested_end_dt.strftime("%H:%M"),
            "vehicle_number": "",
            "driver_username": "",
            "driver_name": "",
            "status": "Pending Department Manager",
            "manager_decision": "Pending",
            "manager_approved_by": "",
            "manager_approved_at": "",
            "hr_decision": "Pending",
            "hr_approved_by": "",
            "hr_approved_at": "",
            "security_released_by": "",
            "security_released_at": "",
            "start_mileage": "",
            "driver_started_at": "",
            "end_mileage": "",
            "distance_km": "",
            "driver_completed_at": "",
            "security_verified_by": "",
            "security_verified_at": "",
            "rejection_reason": "",
        }

        _store_pending_request({
            "request_id": request_id,
            "travel_date": travel_date.isoformat(),
            "start_time": requested_start_time.strftime("%H:%M"),
            "end_time": requested_end_dt.strftime("%H:%M"),
            "duration_minutes": duration_minutes,
            "row": row,
            "manager_name": manager.get(
                "manager_name",
                manager.get("manager_username", "Manager"),
            ),
        })

        st.rerun()

    draft = st.session_state.get("pending_vehicle_request")
    if not draft:
        return

    st.divider()
    st.subheader("📋 Final Transfer Point")

    st.warning(
        "Your request is ready, but it has not been sent to Google Sheets. "
        "Click the button below to officially submit it."
    )

    review = pd.DataFrame([{
        "Request ID": draft["request_id"],
        "Requisitioner": draft["row"]["requisitioner_name"],
        "Department": draft["row"]["department"],
        "Destination": draft["row"]["destination"],
        "Purpose": draft["row"]["purpose"],
        "Date": draft["travel_date"],
        "Required Time": f"{draft['start_time']} - {draft['end_time']}",
        "Duration": f"{draft['duration_minutes']} minutes",
        "Vehicle": "To be allocated",
        "Driver": "To be allocated",
    }])

    st.dataframe(review, use_container_width=True, hide_index=True)

    if st.button(
        "🚀 TRANSFER DATA TO GOOGLE SHEET",
        type="primary",
        use_container_width=True,
        key="transfer_pending_request",
    ):
        try:
            with st.spinner("Transferring request..."):
                if _transfer_pending_request():
                    st.success(
                        f"Request {draft['request_id']} has been transferred successfully."
                    )
                    st.info(
                        "Approval route: "
                        f"{draft['manager_name']} → Vehicle Allocator → "
                        "HR Manager → Security"
                    )
                    st.rerun()
        except Exception as error:
            st.error("The request could not be transferred to Google Sheets.")
            st.exception(error)


# ============================================================
# DEPARTMENT MANAGER
# ============================================================

def manager_portal():
    user = st.session_state.get("user", {})
    username = st.session_state.get("username", "")

    st.header("👤 Department Manager")
    st.caption(
        f"Department: {user.get('department', '')} | "
        f"Manager: {user.get('manager_name', username)}"
    )

    dataframe = get_gatepasses()
    if dataframe.empty:
        st.info("No vehicle requests found.")
        return

    department = str(user.get("department", "")).strip()
    manager_username = str(user.get("manager_username", username)).strip()

    pending = dataframe[
        (dataframe["status"].astype(str) == "Pending Department Manager")
        & (
            (dataframe["manager_username"].astype(str).str.strip() == manager_username)
            | (dataframe["department"].astype(str).str.strip() == department)
        )
    ]

    if pending.empty:
        st.success("No pending requests for your department.")
        return

    for _, row in pending.iterrows():
        request_id = str(row.get("request_id", ""))

        with st.expander(
            f"{request_id} — {row.get('requisitioner_name', '')} — {row.get('travel_date', '')}"
        ):
            col1, col2 = st.columns(2)

            with col1:
                st.write(f"**Department:** {row.get('department', '')}")
                st.write(f"**Destination:** {row.get('destination', '')}")
                st.write(f"**Purpose:** {row.get('purpose', '')}")
                st.write(f"**Passengers:** {row.get('companions', '')}")

            with col2:
                st.write(f"**Date:** {row.get('travel_date', '')}")
                st.write(
                    f"**Requested Vehicle Time:** "
                    f"{row.get('start_time', '')} - {row.get('end_time', '')}"
                )
                st.write(
                    f"**Expected Duration:** {row.get('duration_minutes', '')} minutes"
                )
                st.info(
                    "The Vehicle Allocator will select an available vehicle "
                    "for this requested time."
                )

            reason = st.text_input(
                "Remarks / rejection reason",
                key=f"manager_reason_{request_id}",
            )

            col_a, col_b = st.columns(2)

            with col_a:
                if st.button("Approve", key=f"manager_approve_{request_id}"):
                    update_request(
                        request_id,
                        {
                            "status": "Pending Vehicle Allocation",
                            "manager_decision": "Approved",
                            "manager_approved_by": username,
                            "manager_approved_at": now_str(),
                        },
                    )
                    audit(
                        request_id,
                        username,
                        "Department Manager",
                        "Manager Approved",
                        reason,
                    )
                    st.success(
                        "Request approved and sent to the Vehicle Allocator."
                    )
                    st.rerun()

            with col_b:
                if st.button("Reject", key=f"manager_reject_{request_id}"):
                    update_request(
                        request_id,
                        {
                            "status": "Rejected by Department Manager",
                            "manager_decision": "Rejected",
                            "manager_approved_by": username,
                            "manager_approved_at": now_str(),
                            "rejection_reason": reason,
                        },
                    )
                    audit(
                        request_id,
                        username,
                        "Department Manager",
                        "Manager Rejected",
                        reason,
                    )
                    st.warning("Request rejected.")
                    st.rerun()


# ============================================================
# VEHICLE ALLOCATOR
# ============================================================

def vehicle_allocator_portal():
    username = st.session_state.get("username", "")
    user = st.session_state.get("user", {})

    st.header("🚐 Vehicle Allocator")
    st.caption(f"Vehicle Allocator: {user.get('name', username)}")

    dataframe = get_gatepasses()
    if dataframe.empty:
        st.info("No vehicle requests found.")
        return

    pending = dataframe[
        dataframe["status"].astype(str) == "Pending Vehicle Allocation"
    ]

    if pending.empty:
        st.success("No requests are awaiting vehicle allocation.")
        return

    for _, row in pending.iterrows():
        request_id = str(row.get("request_id", ""))
        travel_date = parse_date(row.get("travel_date"))

        try:
            duration_minutes = int(float(row.get("duration_minutes", 0)))
        except (ValueError, TypeError):
            duration_minutes = 0

        requested_start = parse_time(row.get("start_time"))
        requested_end = parse_time(row.get("end_time"))

        with st.expander(
            f"{request_id} — {row.get('requisitioner_name', '')} — {row.get('travel_date', '')}"
        ):
            col1, col2 = st.columns(2)

            with col1:
                st.write(
                    f"**Requisitioner:** {row.get('requisitioner_name', '')}"
                )
                st.write(f"**Department:** {row.get('department', '')}")
                st.write(f"**Destination:** {row.get('destination', '')}")
                st.write(f"**Purpose:** {row.get('purpose', '')}")
                st.write(f"**Passengers:** {row.get('companions', '')}")

            with col2:
                st.write(f"**Travel Date:** {row.get('travel_date', '')}")
                st.write(
                    f"**Requested Vehicle Time:** "
                    f"{row.get('start_time', '')} - {row.get('end_time', '')}"
                )
                st.write(f"**Duration:** {duration_minutes} minutes")
                st.success("Department Manager: Approved")

            if not travel_date or not requested_start or duration_minutes <= 0:
                st.error("The travel date, requested start time, or duration is invalid.")
                continue

            requested_start_dt = datetime.combine(travel_date, requested_start)
            calculated_end_dt = requested_start_dt + timedelta(
                minutes=duration_minutes
            )
            allocator_start_time = st.time_input(
                "Departure Time",
                value=requested_start_dt.time(),
                key=f"allocator_start_time_{request_id}",
            )
            
            allocator_start_dt = datetime.combine(
                travel_date,
                allocator_start_time,
            )
            
            allocator_end_dt = allocator_start_dt + timedelta(
                minutes=duration_minutes
            )
            
            st.info(
                f"**New Confirmed Time:** "
                f"{allocator_start_dt.strftime('%H:%M')} - "
                f"{allocator_end_dt.strftime('%H:%M')}"
            )

            if calculated_end_dt.date() != travel_date:
                st.error("The requested trip extends into the next day.")
                continue

            if requested_end:
                stored_end_dt = datetime.combine(travel_date, requested_end)
                if stored_end_dt != calculated_end_dt:
                    st.warning(
                        "The stored end time does not match the requested duration. "
                        "The calculated duration will be used."
                    )

            vehicles = available_vehicles_for_interval(
                travel_date,
                requested_start,
                duration_minutes,
            )

            if not vehicles:
                st.warning(
                    "No vehicles are currently available for the requested "
                    f"time {requested_start_dt.strftime('%H:%M')} - "
                    f"{calculated_end_dt.strftime('%H:%M')}."
                )
                continue

            st.info(
                f"Employee requested: **{requested_start_dt.strftime('%H:%M')} - "
                f"{calculated_end_dt.strftime('%H:%M')}**"
            )

            vehicle_labels = [
                (
                    f"{v['vehicle_number']} — {v['vehicle_type']} — "
                    f"Fixed Driver: {v['driver_name'] or v['driver_username']}"
                )
                for v in vehicles
            ]

            selected_label = st.selectbox(
                "Allocate Vehicle *",
                vehicle_labels,
                key=f"allocator_vehicle_{request_id}",
            )

            selected_vehicle = vehicles[vehicle_labels.index(selected_label)]
            selected_vehicle_number = selected_vehicle["vehicle_number"]

            selected_driver = vehicle_driver(selected_vehicle_number) or {}

            st.info(
                f"**Vehicle:** {selected_vehicle_number}\n\n"
                f"**Fixed Driver:** "
                f"{selected_driver.get('driver_name', '') or selected_driver.get('driver_username', '')}\n\n"
                f"**Confirmed Time:** "
                f"{requested_start_dt.strftime('%H:%M')} - "
                f"{calculated_end_dt.strftime('%H:%M')}"
            )

            remarks = st.text_input(
                "Allocation remarks",
                key=f"allocator_remarks_{request_id}",
            )

            if st.button(
                "Allocate Vehicle & Send to HR",
                type="primary",
                key=f"allocate_vehicle_{request_id}",
            ):
                read_sheet.clear()

                if not vehicle_is_available(
                    selected_vehicle_number,
                    travel_date,
                    allocator_start_dt,
                    allocator_end_dt,
                    exclude_request_id=request_id,
                ):
                    st.error(
                        "This vehicle has just been allocated to another request. "
                        "Please select another vehicle."
                    )
                    st.rerun()

                driver_info = vehicle_driver(selected_vehicle_number) or {}

                update_request(
                    request_id,
                    {
                        "status": "Pending HR",
                        "start_time": (
                            allocator_start_dt.strftime(
                                "%H:%M"
                            )
                        ),
                        "end_time": (
                            allocator_end_dt.strftime(
                                "%H:%M"
                            )
                        ),
                        "vehicle_number": selected_vehicle_number,
                        "driver_username": driver_info.get("driver_username", ""),
                        "driver_name": driver_info.get("driver_name", ""),
                    },
                )

                audit(
                    request_id,
                    username,
                    "Vehicle Allocator",
                    "Vehicle Allocated",
                    (
                        f"Vehicle {selected_vehicle_number}; "
                        f"Fixed driver "
                        f"{driver_info.get('driver_name', '')}; "
                        f"Time {allocator_start_dt.strftime('%H:%M')} - "
                        f"{allocator_end_dt.strftime('%H:%M')}; "
                        f"{remarks}"
                    ),
                )

                st.success(
                    "Vehicle allocated successfully and request sent to HR."
                )
                st.rerun()


# ============================================================
# HR
# ============================================================

def hr_portal():
    username = st.session_state.get("username", "")
    st.header("👩‍💼 HR Manager")

    dataframe = get_gatepasses()
    if dataframe.empty:
        st.info("No vehicle requests found.")
        return

    pending = dataframe[
        dataframe["status"].astype(str) == "Pending HR"
    ]

    if pending.empty:
        st.success("No pending HR approvals.")
    else:
        for _, row in pending.iterrows():
            request_id = str(row.get("request_id", ""))

            with st.expander(
                f"{request_id} — {row.get('requisitioner_name', '')} — {row.get('department', '')}"
            ):
                col1, col2 = st.columns(2)

                with col1:
                    st.write(f"**Destination:** {row.get('destination', '')}")
                    st.write(f"**Purpose:** {row.get('purpose', '')}")
                    st.write(f"**Date:** {row.get('travel_date', '')}")
                    st.write(
                        f"**Requested/Allocated Time:** "
                        f"{row.get('start_time', '')} - {row.get('end_time', '')}"
                    )

                with col2:
                    st.write(f"**Vehicle:** {row.get('vehicle_number', '')}")
                    st.write(f"**Fixed Driver:** {row.get('driver_name', '')}")
                    st.success("Department Manager: Approved")
                    st.success("Vehicle Allocated")

                reason = st.text_input(
                    "HR remarks / rejection reason",
                    key=f"hr_reason_{request_id}",
                )

                col_a, col_b = st.columns(2)

                with col_a:
                    if st.button("HR Approve", key=f"hr_approve_{request_id}"):
                        update_request(
                            request_id,
                            {
                                "status": "Pending Security",
                                "hr_decision": "Approved",
                                "hr_approved_by": username,
                                "hr_approved_at": now_str(),
                            },
                        )
                        audit(
                            request_id,
                            username,
                            "HR Manager",
                            "HR Approved",
                            reason,
                        )
                        st.success(
                            "Request approved and sent to Security."
                        )
                        st.rerun()

                with col_b:
                    if st.button("HR Reject", key=f"hr_reject_{request_id}"):
                        update_request(
                            request_id,
                            {
                                "status": "Rejected by HR",
                                "hr_decision": "Rejected",
                                "hr_approved_by": username,
                                "hr_approved_at": now_str(),
                                "rejection_reason": reason,
                            },
                        )
                        audit(
                            request_id,
                            username,
                            "HR Manager",
                            "HR Rejected",
                            reason,
                        )
                        st.warning("Request rejected.")
                        st.rerun()

    st.divider()
    st.subheader("Recent Vehicle Requests")
    st.dataframe(dataframe.tail(30), use_container_width=True)


# ============================================================
# SECURITY
# ============================================================

def security_portal():
    username = st.session_state.get("username", "")
    st.header("🛡️ Security")

    dataframe = get_gatepasses()
    if dataframe.empty:
        st.info("No vehicle requests found.")
        return

    st.subheader("Approved Requests Awaiting Vehicle Release")

    pending_release = dataframe[
        dataframe["status"].astype(str) == "Pending Security"
    ]

    if pending_release.empty:
        st.info("No approved requests are awaiting release.")
    else:
        for _, row in pending_release.iterrows():
            request_id = str(row.get("request_id", ""))

            with st.expander(
                f"{request_id} — {row.get('requisitioner_name', '')} — {row.get('vehicle_number', '')}"
            ):
                st.write(f"**Employee:** {row.get('requisitioner_name', '')}")
                st.write(f"**Department:** {row.get('department', '')}")
                st.write(f"**Destination:** {row.get('destination', '')}")
                st.write(f"**Purpose:** {row.get('purpose', '')}")
                st.write(f"**Date:** {row.get('travel_date', '')}")
                st.write(
                    f"**Time:** {row.get('start_time', '')} - {row.get('end_time', '')}"
                )
                st.write(f"**Vehicle:** {row.get('vehicle_number', '')}")
                st.write(f"**Fixed Driver:** {row.get('driver_name', '')}")

                st.success("Department Manager: Approved")
                st.success("Vehicle Allocated")
                st.success("HR: Approved")

                st.info(
                    "Security does not assign the vehicle or driver. "
                    "Both are already defined by the approved request."
                )

                if st.button(
                    "Verify & Release Vehicle",
                    key=f"security_release_{request_id}",
                ):
                    update_request(
                        request_id,
                        {
                            "status": "Vehicle Released",
                            "security_released_by": username,
                            "security_released_at": now_str(),
                        },
                    )
                    audit(
                        request_id,
                        username,
                        "Security",
                        "Vehicle Released",
                        (
                            f"Vehicle {row.get('vehicle_number', '')}; "
                            f"fixed driver {row.get('driver_name', '')}"
                        ),
                    )
                    st.success(
                        "Vehicle released. The assigned driver can now start the trip."
                    )
                    st.rerun()

    st.divider()
    st.subheader("Trips Awaiting Final Security Verification")

    completed_by_driver = dataframe[
        dataframe["status"].astype(str) == "Pending Security Verification"
    ]

    if completed_by_driver.empty:
        st.info("No trips are awaiting final verification.")
    else:
        for _, row in completed_by_driver.iterrows():
            request_id = str(row.get("request_id", ""))

            with st.expander(
                f"{request_id} — {row.get('vehicle_number', '')} — {row.get('driver_name', '')}"
            ):
                st.write(f"**Vehicle:** {row.get('vehicle_number', '')}")
                st.write(f"**Fixed Driver:** {row.get('driver_name', '')}")
                st.write(f"**Start Mileage:** {row.get('start_mileage', '')} km")
                st.write(f"**End Mileage:** {row.get('end_mileage', '')} km")
                st.write(f"**Distance:** {row.get('distance_km', '')} km")

                remarks = st.text_input(
                    "Verification remarks",
                    key=f"security_verify_reason_{request_id}",
                )

                if st.button(
                    "Verify & Close Trip",
                    key=f"security_verify_{request_id}",
                ):
                    update_request(
                        request_id,
                        {
                            "status": "Completed",
                            "security_verified_by": username,
                            "security_verified_at": now_str(),
                        },
                    )
                    audit(
                        request_id,
                        username,
                        "Security",
                        "Trip Verified and Closed",
                        remarks,
                    )
                    st.success("Trip verified and closed.")
                    st.rerun()


# ============================================================
# DRIVER
# ============================================================

def driver_portal():
    username = st.session_state.get("username", "")
    user = st.session_state.get("user", {})

    st.header("🚐 Driver Portal")
    st.caption(f"Driver: {user.get('driver_name', username)}")

    dataframe = get_gatepasses()
    if dataframe.empty:
        st.info("No assigned trips.")
        return

    assigned = dataframe[
        (
            dataframe["driver_username"].astype(str).str.strip()
            == username
        )
        & dataframe["status"].astype(str).isin([
            "Vehicle Released",
            "Trip In Progress",
            "Pending Security Verification",
        ])
    ]

    if assigned.empty:
        st.success("No active assigned trips.")
        return

    for _, row in assigned.iterrows():
        request_id = str(row.get("request_id", ""))
        status = str(row.get("status", ""))

        with st.expander(
            f"{request_id} — {row.get('vehicle_number', '')} — {row.get('travel_date', '')}"
        ):
            st.write(f"**Vehicle:** {row.get('vehicle_number', '')}")
            st.write(f"**Destination:** {row.get('destination', '')}")
            st.write(f"**Purpose:** {row.get('purpose', '')}")
            st.write(
                f"**Departure:** {row.get('start_time', '')}"
            )
            st.write(
                f"**Return:** {row.get('end_time', '')}"
            )
            st.write(f"**Status:** {status}")

            if status == "Vehicle Released":
                start_mileage = st.number_input(
                    "Starting mileage (km)",
                    min_value=0.0,
                    step=1.0,
                    key=f"start_mileage_{request_id}",
                )

                if st.button("Start Trip", key=f"start_trip_{request_id}"):
                    if start_mileage <= 0:
                        st.error("Please enter the starting mileage.")
                    else:
                        update_request(
                            request_id,
                            {
                                "status": "Trip In Progress",
                                "start_mileage": start_mileage,
                                "driver_started_at": now_str(),
                            },
                        )
                        audit(
                            request_id,
                            username,
                            "Driver",
                            "Trip Started",
                            f"Start mileage: {start_mileage} km",
                        )
                        st.success("Trip started.")
                        st.rerun()

            elif status == "Trip In Progress":
                try:
                    start_value = float(row.get("start_mileage", 0))
                except (ValueError, TypeError):
                    start_value = 0.0

                st.info(f"Recorded starting mileage: {start_value:g} km")

                end_mileage = st.number_input(
                    "Ending mileage (km)",
                    min_value=start_value,
                    step=1.0,
                    key=f"end_mileage_{request_id}",
                )

                if st.button(
                    "Complete Trip",
                    key=f"complete_trip_{request_id}",
                ):
                    if end_mileage < start_value:
                        st.error(
                            "Ending mileage cannot be less than starting mileage."
                        )
                    else:
                        distance = end_mileage - start_value

                        update_request(
                            request_id,
                            {
                                "status": "Pending Security Verification",
                                "end_mileage": end_mileage,
                                "distance_km": distance,
                                "driver_completed_at": now_str(),
                            },
                        )
                        audit(
                            request_id,
                            username,
                            "Driver",
                            "Trip Completed",
                            (
                                f"End mileage: {end_mileage} km; "
                                f"Distance: {distance:g} km"
                            ),
                        )
                        st.success(
                            "Trip completed and sent to Security for verification."
                        )
                        st.rerun()

            else:
                st.success(
                    "Trip completed by driver and is waiting for Security verification."
                )
                st.write(
                    f"Start mileage: {row.get('start_mileage', '')} km"
                )
                st.write(
                    f"End mileage: {row.get('end_mileage', '')} km"
                )
                st.write(
                    f"Distance: {row.get('distance_km', '')} km"
                )


# ============================================================
# ADMINISTRATION
# ============================================================

def admin_portal():
    st.header("⚙️ Administration")

    dataframe = get_gatepasses()
    if dataframe.empty:
        st.info("No vehicle gate-pass records yet.")
        return

    st.metric("Total Requests", len(dataframe))

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric(
            "Pending Manager",
            int((dataframe["status"] == "Pending Department Manager").sum()),
        )

    with col2:
        st.metric(
            "Pending Vehicle Allocation",
            int((dataframe["status"] == "Pending Vehicle Allocation").sum()),
        )

    with col3:
        st.metric(
            "Pending HR",
            int((dataframe["status"] == "Pending HR").sum()),
        )

    with col4:
        st.metric(
            "Completed",
            int((dataframe["status"] == "Completed").sum()),
        )

    st.subheader("Gate Pass History Download")

    travel_dates = sorted([
        parsed_date
        for parsed_date in (
            parse_date(value)
            for value in dataframe["travel_date"].tolist()
        )
        if parsed_date is not None
    ])

    if travel_dates:
        selected_history_date = st.date_input(
            "Travel Date *",
            value=travel_dates[-1],
            key="admin_history_travel_date",
        )

        selected_records = dataframe[
            dataframe["travel_date"].apply(parse_date)
            == selected_history_date
        ].copy()

        st.write(
            f"Records for Travel Date **{selected_history_date.isoformat()}**: "
            f"**{len(selected_records)}**"
        )

        if selected_records.empty:
            st.info("No GatePass records found for the selected Travel Date.")
        else:
            st.dataframe(
                selected_records,
                use_container_width=True,
                hide_index=True,
            )

            excel_buffer = pd.io.common.BytesIO()

            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                selected_records.to_excel(
                    writer,
                    index=False,
                    sheet_name="GatePasses",
                )

            excel_buffer.seek(0)

            st.download_button(
                label="📥 Download Gate Pass History as Excel",
                data=excel_buffer.getvalue(),
                file_name=(
                    f"GatePass_History_{selected_history_date.isoformat()}.xlsx"
                ),
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True,
            )

    st.divider()

    st.subheader("All Gate Pass Records")
    st.dataframe(dataframe, use_container_width=True, height=500)

    st.subheader("Vehicle Master")
    st.dataframe(get_vehicles(), use_container_width=True)

    st.subheader("Driver Master")
    drivers_view = get_drivers().copy()
    if "password" in drivers_view.columns:
        drivers_view = drivers_view.drop(columns=["password"])
    st.dataframe(drivers_view, use_container_width=True)

    st.subheader("Department Managers")
    departments_view = get_departments().copy()
    if "manager_password" in departments_view.columns:
        departments_view = departments_view.drop(columns=["manager_password"])
    st.dataframe(departments_view, use_container_width=True)

    st.subheader("System Users")
    users_view = get_users().copy()
    if "password" in users_view.columns:
        users_view = users_view.drop(columns=["password"])
    st.dataframe(users_view, use_container_width=True)


# ============================================================
# MAIN
# ============================================================

def main():
    st.title("🚐 Vehicle Gate Pass Management System")

    st.caption(
        "Digital vehicle requisition, multi-level approval, "
        "employee-requested vehicle time, vehicle allocation, "
        "fixed vehicle-driver assignment, mileage recording "
        "and security verification."
    )

    if st.session_state.get("logged_in"):
        role = st.session_state.get("role", "")

        with st.sidebar:
            st.success(
                f"Logged in: {st.session_state.get('username', '')}"
            )
            st.write(f"Role: **{role}**")

            if st.button(
                "🔄 Refresh Google Sheet Data",
                use_container_width=True,
            ):
                invalidate_data_cache()
                st.success("Google Sheet cache refreshed.")
                st.rerun()

            if st.button("Logout", use_container_width=True):
                logout()

        if role == "Employee / Requisitioner":
            employee_portal()
        elif role == "Department Manager":
            manager_portal()
        elif role == "Vehicle Allocator":
            vehicle_allocator_portal()
        elif role == "HR Manager":
            hr_portal()
        elif role == "Security":
            security_portal()
        elif role == "Driver":
            driver_portal()
        elif role == "Administration":
            admin_portal()
        else:
            st.error("Unknown role.")
        return

    role = st.radio("Select Portal", ROLES, index=0)

    if role == "Employee / Requisitioner":
        employee_portal()
    else:
        login_portal(role)


if __name__ == "__main__":
    main()
