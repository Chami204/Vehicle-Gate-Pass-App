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

st.set_page_config(page_title="Vehicle Gate Pass", page_icon="🚐", layout="wide")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SHEET_NAME = "vehicle_gate_pass_data"

SHEET_HEADERS = {
    "Users": ["username", "name", "role", "password", "active"],
    "Departments": ["department", "manager_username", "manager_name", "manager_password", "active"],
    "Drivers": ["driver_username", "driver_name", "password", "active"],
    "Vehicles": ["vehicle_number", "vehicle_type", "driver_username", "driver_name", "active"],
    "GatePasses": [
        "request_id", "created_at", "requisitioner_name", "department",
        "manager_username", "companions", "duration_minutes", "destination",
        "purpose", "travel_date", "start_time", "end_time", "vehicle_number",
        "driver_username", "driver_name", "status",
        "manager_decision", "manager_approved_by", "manager_approved_at",
        "hr_decision", "hr_approved_by", "hr_approved_at",
        "security_released_by", "security_released_at",
        "start_mileage", "driver_started_at", "end_mileage", "distance_km",
        "driver_completed_at", "security_verified_by", "security_verified_at",
        "rejection_reason",
    ],
    "ApprovalAudit": [
        "timestamp", "request_id", "username", "role", "action", "remarks"
    ],
}

ACTIVE_STATUSES = {
    "Pending Department Manager",
    "Pending HR",
    "Pending Security",
    "Vehicle Released",
    "Trip In Progress",
    "Pending Security Verification",
}

ROLES = [
    "Employee / Requisitioner",
    "Department Manager",
    "HR Manager",
    "Security",
    "Driver",
    "Administration",
]


# ============================================================
# GOOGLE SHEETS CONNECTION
# ============================================================

@st.cache_resource
def get_google_client():
    """Read the service-account fields from Streamlit Secrets."""
    required = [
        "type",
        "project_id",
        "private_key_id",
        "private_key",
        "client_email",
        "client_id",
        "auth_uri",
        "token_uri",
        "auth_provider_x509_cert_url",
        "client_x509_cert_url",
    ]

    missing = [key for key in required if key not in st.secrets]

    if missing:
        raise RuntimeError(
            "Missing Google service-account secrets: "
            + ", ".join(missing)
        )

    info = {
        "type": st.secrets["type"],
        "project_id": st.secrets["project_id"],
        "private_key_id": st.secrets["private_key_id"],
        "private_key": st.secrets["private_key"],
        "client_email": st.secrets["client_email"],
        "client_id": str(st.secrets["client_id"]),
        "auth_uri": st.secrets["auth_uri"],
        "token_uri": st.secrets["token_uri"],
        "auth_provider_x509_cert_url": st.secrets["auth_provider_x509_cert_url"],
        "client_x509_cert_url": st.secrets["client_x509_cert_url"],
    }

    if "universe_domain" in st.secrets:
        info["universe_domain"] = st.secrets["universe_domain"]

    credentials = Credentials.from_service_account_info(
        info,
        scopes=SCOPES,
    )

    return gspread.authorize(credentials)


@st.cache_resource
def get_spreadsheet():
    """Open the Google Sheet by ID. Name is only a fallback."""
    client = get_google_client()

    sheet_id = str(
        st.secrets.get("google_sheet_id", "")
    ).strip()

    if sheet_id:
        return client.open_by_key(sheet_id)

    sheet_name = str(
        st.secrets.get("google_sheet_name", SHEET_NAME)
    ).strip() or SHEET_NAME

    return client.open(sheet_name)


def get_or_create_worksheet(name):
    """Create a missing worksheet tab with the expected headers."""
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

        worksheet.append_row(
            headers,
            value_input_option="USER_ENTERED",
        )
    else:
        headers = SHEET_HEADERS[name]
        current_headers = worksheet.row_values(1)

        if not current_headers:
            worksheet.append_row(
                headers,
                value_input_option="USER_ENTERED",
            )

    return worksheet


@st.cache_data(ttl=10)
def read_sheet(name):
    worksheet = get_or_create_worksheet(name)
    records = worksheet.get_all_records()
    return pd.DataFrame(records)


def invalidate_data_cache():
    read_sheet.clear()


def append_row(name, row_dict):
    worksheet = get_or_create_worksheet(name)
    headers = SHEET_HEADERS[name]

    worksheet.append_row(
        [row_dict.get(header, "") for header in headers],
        value_input_option="USER_ENTERED",
    )

    invalidate_data_cache()


def update_request(request_id, updates):
    worksheet = get_or_create_worksheet("GatePasses")
    headers = worksheet.row_values(1)
    records = worksheet.get_all_records()

    target_row = None

    for row_number, record in enumerate(records, start=2):
        if str(record.get("request_id", "")).strip() == str(request_id).strip():
            target_row = row_number
            break

    if target_row is None:
        raise ValueError(f"Request {request_id} was not found.")

    for key, value in updates.items():
        if key not in headers:
            continue

        column_number = headers.index(key) + 1

        worksheet.update_cell(
            target_row,
            column_number,
            "" if value is None else str(value),
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
# GENERAL HELPERS
# ============================================================

def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def normalize_bool(value):
    return str(value).strip().lower() in {
        "yes",
        "true",
        "1",
        "active",
        "y",
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


def dt_from_parts(d, t):
    return datetime.combine(d, t)


def fmt_time(value):
    return value.strftime("%H:%M")


def generate_request_id():
    return (
        "VGP-"
        + datetime.now().strftime("%Y%m%d")
        + "-"
        + uuid.uuid4().hex[:6].upper()
    )


def active_records(dataframe):
    if dataframe.empty:
        return dataframe

    if "active" not in dataframe.columns:
        return dataframe

    return dataframe[
        dataframe["active"].apply(normalize_bool)
    ].copy()


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
    """HR, Security and Administration login from Users sheet."""
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
    """Department Manager login from Departments sheet."""
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
    """Driver login from Drivers sheet."""
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
    """
    Return the permanent driver assigned to a vehicle.

    Security never assigns a driver.
    The vehicle master determines the driver.
    """
    vehicles = get_vehicles()
    drivers = get_drivers()

    for _, vehicle in vehicles.iterrows():
        if (
            str(vehicle.get("vehicle_number", "")).strip()
            == str(vehicle_number).strip()
        ):
            driver_username = str(
                vehicle.get("driver_username", "")
            ).strip()

            driver_name = str(
                vehicle.get("driver_name", "")
            ).strip()

            if (
                driver_username
                and not driver_name
            ):
                for _, driver in drivers.iterrows():
                    if (
                        str(driver.get("driver_username", "")).strip()
                        == driver_username
                    ):
                        driver_name = str(
                            driver.get("driver_name", "")
                        ).strip()
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
        dt_from_parts(travel_date, start_time),
        dt_from_parts(travel_date, end_time),
    )


def vehicle_is_available(
    vehicle_number,
    travel_date,
    start_dt,
    end_dt,
    exclude_request_id=None,
):
    """
    A vehicle is unavailable when another active request overlaps it.
    Rejected and completed requests do not block future bookings.
    """
    gatepasses = get_gatepasses()

    if gatepasses.empty:
        return True

    for _, row in gatepasses.iterrows():
        if (
            str(row.get("vehicle_number", "")).strip()
            != str(vehicle_number).strip()
        ):
            continue

        if (
            exclude_request_id
            and str(row.get("request_id", "")).strip()
            == str(exclude_request_id).strip()
        ):
            continue

        status = str(row.get("status", "")).strip()

        if status not in ACTIVE_STATUSES:
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


def available_start_times(
    vehicle_number,
    travel_date,
    duration_minutes,
):
    """
    Returns 30-minute departure slots between 06:00 and 22:00
    for which the complete requested duration is available.
    """
    slots = []

    current = datetime.combine(
        travel_date,
        time(6, 0),
    )

    latest_start = (
        datetime.combine(
            travel_date,
            time(22, 0),
        )
        - timedelta(minutes=duration_minutes)
    )

    while current <= latest_start:
        end_dt = current + timedelta(
            minutes=duration_minutes
        )

        if vehicle_is_available(
            vehicle_number,
            travel_date,
            current,
            end_dt,
        ):
            slots.append(current.time())

        current += timedelta(minutes=30)

    return slots


def available_vehicles(travel_date, duration_minutes):
    result = []

    for _, vehicle in get_vehicles().iterrows():
        vehicle_number = str(
            vehicle.get("vehicle_number", "")
        ).strip()

        vehicle_type = str(
            vehicle.get("vehicle_type", "")
        ).strip()

        if not vehicle_number:
            continue

        times = available_start_times(
            vehicle_number,
            travel_date,
            duration_minutes,
        )

        if times:
            driver_info = vehicle_driver(
                vehicle_number
            ) or {}

            result.append(
                {
                    "vehicle_number": vehicle_number,
                    "vehicle_type": vehicle_type,
                    "driver_username": driver_info.get(
                        "driver_username",
                        str(vehicle.get("driver_username", "")).strip(),
                    ),
                    "driver_name": driver_info.get(
                        "driver_name",
                        str(vehicle.get("driver_name", "")).strip(),
                    ),
                }
            )

    return result


# ============================================================
# SESSION / LOGOUT
# ============================================================

def logout():
    for key in list(st.session_state.keys()):
        del st.session_state[key]

    st.rerun()


# ============================================================
# LOGIN UI
# ============================================================

def login_portal(role):
    st.subheader(f"{role} Login")

    if role == "Department Manager":
        username = st.text_input(
            "Manager Username",
            key="manager_username_login",
        )

        password = st.text_input(
            "Manager Password",
            type="password",
            key="manager_password_login",
        )

        if st.button(
            "Login",
            key="manager_login_button",
        ):
            user = authenticate_manager(
                username,
                password,
            )

            if user:
                st.session_state.logged_in = True
                st.session_state.role = role
                st.session_state.username = username.strip()
                st.session_state.user = user
                st.rerun()
            else:
                st.error(
                    "Invalid manager username or password."
                )

    elif role == "Driver":
        username = st.text_input(
            "Driver Username",
            key="driver_username_login",
        )

        password = st.text_input(
            "Driver Password",
            type="password",
            key="driver_password_login",
        )

        if st.button(
            "Login",
            key="driver_login_button",
        ):
            user = authenticate_driver(
                username,
                password,
            )

            if user:
                st.session_state.logged_in = True
                st.session_state.role = role
                st.session_state.username = username.strip()
                st.session_state.user = user
                st.rerun()
            else:
                st.error(
                    "Invalid driver username or password."
                )

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

        if st.button(
            "Login",
            key=f"{role}_login_button",
        ):
            user = authenticate_user(
                username,
                password,
                role,
            )

            if user:
                st.session_state.logged_in = True
                st.session_state.role = role
                st.session_state.username = username.strip()
                st.session_state.user = user
                st.rerun()
            else:
                st.error(
                    "Invalid username or password."
                )


# ============================================================
# EMPLOYEE / REQUISITIONER
# ============================================================

def employee_portal():
    st.header("🚐 Vehicle Requisition")

    departments = get_departments()

    if (
        departments.empty
        or "department" not in departments.columns
    ):
        st.error(
            "No departments are configured in the Departments sheet."
        )
        return

    department_names = sorted(
        [
            str(value).strip()
            for value in departments["department"].tolist()
            if str(value).strip()
        ]
    )

    with st.form("vehicle_request_form"):
        col1, col2 = st.columns(2)

        with col1:
            requisitioner = st.text_input(
                "Requisitioner Name *"
            )

            department = st.selectbox(
                "Department *",
                department_names,
            )

            companions = st.text_area(
                "Person(s) travelling with you",
                placeholder="Enter names separated by commas",
            )

            destination = st.text_input(
                "Where are you going? *"
            )

            purpose = st.text_area(
                "Purpose of travel *"
            )

        with col2:
            travel_date = st.date_input(
                "Travel Date *",
                min_value=date.today(),
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

            duration_minutes = duration_options[
                duration_label
            ]

        st.divider()

        vehicles = available_vehicles(
            travel_date,
            duration_minutes,
        )

        if not vehicles:
            st.warning(
                "No vehicles have an available time slot "
                "for the selected date and duration."
            )

            selected_vehicle_number = None
            selected_start = None

        else:
            vehicle_labels = [
                (
                    f"{vehicle['vehicle_number']} — "
                    f"{vehicle['vehicle_type']} — "
                    f"Fixed Driver: "
                    f"{vehicle['driver_name'] or vehicle['driver_username']}"
                )
                for vehicle in vehicles
            ]

            selected_label = st.selectbox(
                "Available Vehicle *",
                vehicle_labels,
            )

            selected_index = vehicle_labels.index(
                selected_label
            )

            selected_vehicle = vehicles[
                selected_index
            ]

            selected_vehicle_number = (
                selected_vehicle["vehicle_number"]
            )

            times = available_start_times(
                selected_vehicle_number,
                travel_date,
                duration_minutes,
            )

            if not times:
                st.error(
                    "No available time is currently "
                    "offered for this vehicle."
                )
                selected_start = None

            else:
                selected_start = st.selectbox(
                    "Available Departure Time *",
                    times,
                    format_func=fmt_time,
                )

                st.info(
                    f"**Vehicle:** {selected_vehicle_number}\n\n"
                    f"**Fixed Driver:** "
                    f"{selected_vehicle['driver_name'] or selected_vehicle['driver_username']}\n\n"
                    "The requisitioner cannot change the vehicle's assigned driver."
                )

        submitted = st.form_submit_button(
            "Submit Vehicle Request",
            disabled=(
                not vehicles
                or selected_start is None
            ),
        )

    if not submitted:
        return

    if (
        not requisitioner.strip()
        or not destination.strip()
        or not purpose.strip()
    ):
        st.error(
            "Please complete all required fields."
        )
        return

    manager = None

    for _, row in departments.iterrows():
        if (
            str(row.get("department", "")).strip()
            == department
        ):
            manager = row.to_dict()
            break

    if not manager:
        st.error(
            "No manager is configured for this department."
        )
        return

    start_dt = datetime.combine(
        travel_date,
        selected_start,
    )

    end_dt = start_dt + timedelta(
        minutes=duration_minutes
    )

    # Final availability check immediately before saving.
    if not vehicle_is_available(
        selected_vehicle_number,
        travel_date,
        start_dt,
        end_dt,
    ):
        st.error(
            "This vehicle/time was just booked. "
            "Please select another available time."
        )
        return

    driver_info = (
        vehicle_driver(
            selected_vehicle_number
        )
        or {}
    )

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
        "start_time": start_dt.strftime("%H:%M"),
        "end_time": end_dt.strftime("%H:%M"),
        "vehicle_number": selected_vehicle_number,
        "driver_username": driver_info.get(
            "driver_username",
            "",
        ),
        "driver_name": driver_info.get(
            "driver_name",
            "",
        ),
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

    append_row(
        "GatePasses",
        row,
    )

    audit(
        request_id,
        "employee",
        "Employee / Requisitioner",
        "Request Submitted",
        (
            f"Vehicle {selected_vehicle_number}; "
            f"fixed driver {driver_info.get('driver_name', '')}"
        ),
    )

    st.success(
        f"Request {request_id} submitted successfully."
    )

    st.info(
        "Approval route: "
        f"{manager.get('manager_name', manager.get('manager_username', 'Manager'))}"
        " → HR Manager → Security"
    )


# ============================================================
# DEPARTMENT MANAGER
# ============================================================

def manager_portal():
    user = st.session_state.get(
        "user",
        {},
    )

    username = st.session_state.get(
        "username",
        "",
    )

    st.header("👤 Department Manager")

    st.caption(
        f"Department: {user.get('department', '')} | "
        f"Manager: {user.get('manager_name', username)}"
    )

    dataframe = get_gatepasses()

    if dataframe.empty:
        st.info(
            "No vehicle requests found."
        )
        return

    department = str(
        user.get("department", "")
    ).strip()

    manager_username = str(
        user.get("manager_username", username)
    ).strip()

    pending = dataframe[
        (
            dataframe["status"].astype(str)
            == "Pending Department Manager"
        )
        & (
            (
                dataframe["manager_username"]
                .astype(str)
                .str.strip()
                == manager_username
            )
            | (
                dataframe["department"]
                .astype(str)
                .str.strip()
                == department
            )
        )
    ]

    if pending.empty:
        st.success(
            "No pending requests for your department."
        )
        return

    for _, row in pending.iterrows():
        request_id = str(
            row.get("request_id", "")
        )

        with st.expander(
            (
                f"{request_id} — "
                f"{row.get('requisitioner_name', '')} — "
                f"{row.get('travel_date', '')} "
                f"{row.get('start_time', '')}"
            )
        ):
            col1, col2 = st.columns(2)

            with col1:
                st.write(
                    f"**Department:** {row.get('department', '')}"
                )
                st.write(
                    f"**Destination:** {row.get('destination', '')}"
                )
                st.write(
                    f"**Purpose:** {row.get('purpose', '')}"
                )
                st.write(
                    f"**Passengers:** {row.get('companions', '')}"
                )

            with col2:
                st.write(
                    f"**Vehicle:** {row.get('vehicle_number', '')}"
                )
                st.write(
                    f"**Fixed Driver:** {row.get('driver_name', '')}"
                )
                st.write(
                    f"**Date:** {row.get('travel_date', '')}"
                )
                st.write(
                    f"**Time:** {row.get('start_time', '')} - "
                    f"{row.get('end_time', '')}"
                )

            reason = st.text_input(
                "Remarks / rejection reason",
                key=f"manager_reason_{request_id}",
            )

            col_a, col_b = st.columns(2)

            with col_a:
                if st.button(
                    "Approve",
                    key=f"manager_approve_{request_id}",
                ):
                    update_request(
                        request_id,
                        {
                            "status": "Pending HR",
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
                        "Request approved and sent to HR."
                    )
                    st.rerun()

            with col_b:
                if st.button(
                    "Reject",
                    key=f"manager_reject_{request_id}",
                ):
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

                    st.warning(
                        "Request rejected."
                    )
                    st.rerun()


# ============================================================
# HR MANAGER
# ============================================================

def hr_portal():
    username = st.session_state.get(
        "username",
        "",
    )

    st.header("👩‍💼 HR Manager")

    dataframe = get_gatepasses()

    if dataframe.empty:
        st.info(
            "No vehicle requests found."
        )
        return

    pending = dataframe[
        dataframe["status"].astype(str)
        == "Pending HR"
    ]

    if pending.empty:
        st.success(
            "No pending HR approvals."
        )

    else:
        for _, row in pending.iterrows():
            request_id = str(
                row.get("request_id", "")
            )

            with st.expander(
                (
                    f"{request_id} — "
                    f"{row.get('requisitioner_name', '')} — "
                    f"{row.get('department', '')}"
                )
            ):
                col1, col2 = st.columns(2)

                with col1:
                    st.write(
                        f"**Destination:** {row.get('destination', '')}"
                    )
                    st.write(
                        f"**Purpose:** {row.get('purpose', '')}"
                    )
                    st.write(
                        f"**Date:** {row.get('travel_date', '')}"
                    )
                    st.write(
                        f"**Time:** {row.get('start_time', '')} - "
                        f"{row.get('end_time', '')}"
                    )

                with col2:
                    st.write(
                        f"**Vehicle:** {row.get('vehicle_number', '')}"
                    )
                    st.write(
                        f"**Fixed Driver:** {row.get('driver_name', '')}"
                    )
                    st.success(
                        "Department Manager: Approved"
                    )

                reason = st.text_input(
                    "HR remarks / rejection reason",
                    key=f"hr_reason_{request_id}",
                )

                col_a, col_b = st.columns(2)

                with col_a:
                    if st.button(
                        "HR Approve",
                        key=f"hr_approve_{request_id}",
                    ):
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
                    if st.button(
                        "HR Reject",
                        key=f"hr_reject_{request_id}",
                    ):
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

                        st.warning(
                            "Request rejected."
                        )
                        st.rerun()

    st.divider()

    st.subheader(
        "Recent Vehicle Requests"
    )

    st.dataframe(
        dataframe.tail(30),
        use_container_width=True,
    )


# ============================================================
# SECURITY
# ============================================================

def security_portal():
    username = st.session_state.get(
        "username",
        "",
    )

    st.header("🛡️ Security")

    dataframe = get_gatepasses()

    if dataframe.empty:
        st.info(
            "No vehicle requests found."
        )
        return

    st.subheader(
        "Approved Requests Awaiting Vehicle Release"
    )

    pending_release = dataframe[
        dataframe["status"].astype(str)
        == "Pending Security"
    ]

    if pending_release.empty:
        st.info(
            "No approved requests are awaiting release."
        )

    else:
        for _, row in pending_release.iterrows():
            request_id = str(
                row.get("request_id", "")
            )

            with st.expander(
                (
                    f"{request_id} — "
                    f"{row.get('requisitioner_name', '')} — "
                    f"{row.get('vehicle_number', '')}"
                )
            ):
                st.write(
                    f"**Employee:** {row.get('requisitioner_name', '')}"
                )
                st.write(
                    f"**Department:** {row.get('department', '')}"
                )
                st.write(
                    f"**Destination:** {row.get('destination', '')}"
                )
                st.write(
                    f"**Purpose:** {row.get('purpose', '')}"
                )
                st.write(
                    f"**Date:** {row.get('travel_date', '')}"
                )
                st.write(
                    f"**Time:** {row.get('start_time', '')} - "
                    f"{row.get('end_time', '')}"
                )
                st.write(
                    f"**Vehicle:** {row.get('vehicle_number', '')}"
                )
                st.write(
                    f"**Fixed Driver:** {row.get('driver_name', '')}"
                )

                st.success(
                    "Department Manager: Approved"
                )
                st.success(
                    "HR: Approved"
                )

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
                        "Vehicle released. "
                        "The assigned driver can now start the trip."
                    )
                    st.rerun()

    st.divider()

    st.subheader(
        "Trips Awaiting Final Security Verification"
    )

    completed_by_driver = dataframe[
        dataframe["status"].astype(str)
        == "Pending Security Verification"
    ]

    if completed_by_driver.empty:
        st.info(
            "No trips are awaiting final verification."
        )

    else:
        for _, row in completed_by_driver.iterrows():
            request_id = str(
                row.get("request_id", "")
            )

            with st.expander(
                (
                    f"{request_id} — "
                    f"{row.get('vehicle_number', '')} — "
                    f"{row.get('driver_name', '')}"
                )
            ):
                st.write(
                    f"**Vehicle:** {row.get('vehicle_number', '')}"
                )
                st.write(
                    f"**Fixed Driver:** {row.get('driver_name', '')}"
                )
                st.write(
                    f"**Start Mileage:** {row.get('start_mileage', '')} km"
                )
                st.write(
                    f"**End Mileage:** {row.get('end_mileage', '')} km"
                )
                st.write(
                    f"**Distance:** {row.get('distance_km', '')} km"
                )

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

                    st.success(
                        "Trip verified and closed."
                    )
                    st.rerun()


# ============================================================
# DRIVER
# ============================================================

def driver_portal():
    username = st.session_state.get(
        "username",
        "",
    )

    user = st.session_state.get(
        "user",
        {},
    )

    st.header("🚐 Driver Portal")

    st.caption(
        f"Driver: {user.get('driver_name', username)}"
    )

    dataframe = get_gatepasses()

    if dataframe.empty:
        st.info(
            "No assigned trips."
        )
        return

    assigned = dataframe[
        (
            dataframe["driver_username"]
            .astype(str)
            .str.strip()
            == username
        )
        & (
            dataframe["status"]
            .astype(str)
            .isin(
                [
                    "Vehicle Released",
                    "Trip In Progress",
                    "Pending Security Verification",
                ]
            )
        )
    ]

    if assigned.empty:
        st.success(
            "No active assigned trips."
        )
        return

    for _, row in assigned.iterrows():
        request_id = str(
            row.get("request_id", "")
        )

        status = str(
            row.get("status", "")
        )

        with st.expander(
            (
                f"{request_id} — "
                f"{row.get('vehicle_number', '')} — "
                f"{row.get('travel_date', '')}"
            )
        ):
            st.write(
                f"**Vehicle:** {row.get('vehicle_number', '')}"
            )
            st.write(
                f"**Destination:** {row.get('destination', '')}"
            )
            st.write(
                f"**Purpose:** {row.get('purpose', '')}"
            )
            st.write(
                f"**Departure:** {row.get('start_time', '')}"
            )
            st.write(
                f"**Return:** {row.get('end_time', '')}"
            )
            st.write(
                f"**Status:** {status}"
            )

            if status == "Vehicle Released":
                start_mileage = st.number_input(
                    "Starting mileage (km)",
                    min_value=0.0,
                    step=1.0,
                    key=f"start_mileage_{request_id}",
                )

                if st.button(
                    "Start Trip",
                    key=f"start_trip_{request_id}",
                ):
                    if start_mileage <= 0:
                        st.error(
                            "Please enter the starting mileage."
                        )
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

                        st.success(
                            "Trip started."
                        )
                        st.rerun()

            elif status == "Trip In Progress":
                try:
                    start_value = float(
                        row.get(
                            "start_mileage",
                            0,
                        )
                    )
                except (ValueError, TypeError):
                    start_value = 0.0

                st.info(
                    f"Recorded starting mileage: "
                    f"{start_value:g} km"
                )

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
                            "Ending mileage cannot be less "
                            "than starting mileage."
                        )
                    else:
                        distance = (
                            end_mileage
                            - start_value
                        )

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
                            "Trip completed and sent to Security "
                            "for verification."
                        )
                        st.rerun()

            else:
                st.success(
                    "Trip completed by driver and is "
                    "waiting for Security verification."
                )

                st.write(
                    f"Start mileage: "
                    f"{row.get('start_mileage', '')} km"
                )
                st.write(
                    f"End mileage: "
                    f"{row.get('end_mileage', '')} km"
                )
                st.write(
                    f"Distance: "
                    f"{row.get('distance_km', '')} km"
                )


# ============================================================
# ADMINISTRATION
# ============================================================

def admin_portal():
    st.header("⚙️ Administration")

    dataframe = get_gatepasses()

    if dataframe.empty:
        st.info(
            "No vehicle gate-pass records yet."
        )
        return

    st.metric(
        "Total Requests",
        len(dataframe),
    )

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric(
            "Pending Manager",
            int(
                (
                    dataframe["status"]
                    == "Pending Department Manager"
                ).sum()
            ),
        )

    with col2:
        st.metric(
            "Pending HR",
            int(
                (
                    dataframe["status"]
                    == "Pending HR"
                ).sum()
            ),
        )

    with col3:
        st.metric(
            "Pending Security",
            int(
                (
                    dataframe["status"]
                    == "Pending Security"
                ).sum()
            ),
        )

    with col4:
        st.metric(
            "Completed",
            int(
                (
                    dataframe["status"]
                    == "Completed"
                ).sum()
            ),
        )

    st.subheader(
        "All Gate Pass Records"
    )

    st.dataframe(
        dataframe,
        use_container_width=True,
        height=500,
    )

    st.subheader(
        "Vehicle Master"
    )

    st.dataframe(
        get_vehicles(),
        use_container_width=True,
    )

    st.subheader(
        "Driver Master"
    )

    drivers_view = get_drivers().copy()

    if "password" in drivers_view.columns:
        drivers_view = drivers_view.drop(
            columns=["password"]
        )

    st.dataframe(
        drivers_view,
        use_container_width=True,
    )

    st.subheader(
        "Department Managers"
    )

    departments_view = get_departments().copy()

    if "manager_password" in departments_view.columns:
        departments_view = departments_view.drop(
            columns=["manager_password"]
        )

    st.dataframe(
        departments_view,
        use_container_width=True,
    )

    st.subheader(
        "System Users"
    )

    users_view = get_users().copy()

    if "password" in users_view.columns:
        users_view = users_view.drop(
            columns=["password"]
        )

    st.dataframe(
        users_view,
        use_container_width=True,
    )


# ============================================================
# GOOGLE SHEET INITIALIZATION
# ============================================================

def ensure_sheets():
    """
    Creates missing worksheet tabs only.
    It never creates a local Excel file.
    """
    for name in SHEET_HEADERS:
        get_or_create_worksheet(name)


# ============================================================
# MAIN
# ============================================================

def main():
    st.title(
        "🚐 Vehicle Gate Pass Management System"
    )

    st.caption(
        "Digital vehicle requisition, multi-level approval, "
        "fixed vehicle-driver assignment, mileage recording "
        "and security verification."
    )

    try:
        ensure_sheets()

    except Exception as error:
        st.error(
            "Unable to connect to Google Sheets."
        )
        st.exception(error)
        st.stop()

    if st.session_state.get("logged_in"):
        role = st.session_state.get(
            "role",
            "",
        )

        with st.sidebar:
            st.success(
                f"Logged in: "
                f"{st.session_state.get('username', '')}"
            )

            st.write(
                f"Role: **{role}**"
            )

            if st.button("Logout"):
                logout()

        if role == "Employee / Requisitioner":
            employee_portal()

        elif role == "Department Manager":
            manager_portal()

        elif role == "HR Manager":
            hr_portal()

        elif role == "Security":
            security_portal()

        elif role == "Driver":
            driver_portal()

        elif role == "Administration":
            admin_portal()

        else:
            st.error(
                "Unknown role."
            )

        return

    role = st.radio(
        "Select Portal",
        ROLES,
        index=0,
    )

    if role == "Employee / Requisitioner":
        employee_portal()
    else:
        login_portal(role)


if __name__ == "__main__":
    main()
