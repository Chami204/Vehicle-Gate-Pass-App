import uuid
from datetime import datetime, date, time, timedelta

import pandas as pd
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Vehicle Gate Pass System",
    page_icon="🚐",
    layout="wide",
)


# ============================================================
# CONFIGURATION
# ============================================================

GOOGLE_SHEET_NAME = st.secrets.get(
    "google_sheet_name",
    "vehicle_gate_pass_data"
)


# ============================================================
# GOOGLE SHEETS STRUCTURE
# ============================================================

SHEET_COLUMNS = {

    "Departments": [
        "department",
        "manager_username",
        "manager_name",
        "manager_password",
    ],

    "Drivers": [
        "driver_username",
        "driver_name",
        "driver_password",
        "active",
    ],

    "Vehicles": [
        "vehicle_number",
        "vehicle_type",
        "active",
    ],

    "GatePasses": [
        "request_id",
        "created_at",
        "requisitioner",
        "department",
        "contact",
        "companions",
        "destination",
        "purpose",
        "travel_date",
        "departure_time",
        "return_date",
        "return_time",
        "vehicle_type",
        "vehicle_number",
        "driver_username",
        "driver_name",
        "status",
        "manager_username",
        "manager_name",
        "manager_approved_at",
        "hr_approved_at",
        "security_released_at",
        "security_name",
        "start_mileage",
        "end_mileage",
        "driver_started_at",
        "driver_completed_at",
        "security_verified_at",
        "security_notes",
        "rejection_reason",
    ],

    "ApprovalAudit": [
        "audit_id",
        "request_id",
        "timestamp",
        "actor_username",
        "actor_name",
        "actor_role",
        "action",
        "remarks",
    ],
}


# ============================================================
# GOOGLE AUTHENTICATION
# ============================================================

@st.cache_resource
def get_google_credentials():

    # IMPORTANT:
    # Your Streamlit secrets are at the TOP LEVEL.
    #
    # Therefore we use:
    # st.secrets["project_id"]
    #
    # NOT:
    # st.secrets["google_service_account"]

    service_account_info = {
        "type": st.secrets["type"],
        "project_id": st.secrets["project_id"],
        "private_key_id": st.secrets["private_key_id"],
        "private_key": st.secrets["private_key"],
        "client_email": st.secrets["client_email"],
        "client_id": st.secrets["client_id"],
        "auth_uri": st.secrets["auth_uri"],
        "token_uri": st.secrets["token_uri"],
        "auth_provider_x509_cert_url": st.secrets[
            "auth_provider_x509_cert_url"
        ],
        "client_x509_cert_url": st.secrets[
            "client_x509_cert_url"
        ],
        "universe_domain": st.secrets.get(
            "universe_domain",
            "googleapis.com"
        ),
    }

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    credentials = Credentials.from_service_account_info(
        service_account_info,
        scopes=scopes,
    )

    return credentials


@st.cache_resource
def get_google_client():

    credentials = get_google_credentials()

    return gspread.authorize(credentials)


@st.cache_resource
def get_spreadsheet():

    client = get_google_client()

    return client.open(GOOGLE_SHEET_NAME)


# ============================================================
# GOOGLE SHEET FUNCTIONS
# ============================================================

def get_worksheet(sheet_name):

    spreadsheet = get_spreadsheet()

    return spreadsheet.worksheet(sheet_name)


@st.cache_data(ttl=10)
def read_sheet(sheet_name):

    worksheet = get_worksheet(sheet_name)

    values = worksheet.get_all_values()

    expected_columns = SHEET_COLUMNS[sheet_name]

    if not values:

        return pd.DataFrame(
            columns=expected_columns
        )

    headers = values[0]

    rows = values[1:]

    normalized_rows = []

    for row in rows:

        row = list(row)

        if len(row) < len(headers):

            row.extend(
                [""] * (
                    len(headers) - len(row)
                )
            )

        elif len(row) > len(headers):

            row = row[:len(headers)]

        normalized_rows.append(row)

    df = pd.DataFrame(
        normalized_rows,
        columns=headers,
    )

    for column in expected_columns:

        if column not in df.columns:

            df[column] = ""

    df = df[expected_columns]

    return df.fillna("").astype(str)


def load_data():

    return {
        name: read_sheet(name)
        for name in SHEET_COLUMNS
    }


def clear_data_cache():

    read_sheet.clear()


# ============================================================
# WRITE DATA
# ============================================================

def append_record(
    sheet_name,
    record,
):

    worksheet = get_worksheet(sheet_name)

    columns = SHEET_COLUMNS[sheet_name]

    row = [
        str(record.get(column, ""))
        for column in columns
    ]

    worksheet.append_row(
        row,
        value_input_option="USER_ENTERED",
    )

    clear_data_cache()


def update_record(
    sheet_name,
    key_column,
    key_value,
    changes,
):

    worksheet = get_worksheet(sheet_name)

    values = worksheet.get_all_values()

    if not values:

        return False

    headers = values[0]

    if key_column not in headers:

        return False

    key_index = headers.index(key_column)

    target_row = None

    for row_number, row in enumerate(
        values[1:],
        start=2,
    ):

        if len(row) <= key_index:

            continue

        if str(row[key_index]).strip() == str(key_value).strip():

            target_row = row_number
            break

    if target_row is None:

        return False

    for column, value in changes.items():

        if column not in headers:

            continue

        column_number = headers.index(column) + 1

        worksheet.update_cell(
            target_row,
            column_number,
            str(value),
        )

    clear_data_cache()

    return True


# ============================================================
# AUDIT
# ============================================================

def add_audit(
    request_id,
    user,
    action,
    remarks="",
):

    append_record(
        "ApprovalAudit",
        {
            "audit_id":
                "AUD-"
                + uuid.uuid4().hex[:8].upper(),

            "request_id":
                request_id,

            "timestamp":
                datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                ),

            "actor_username":
                user.get("username", ""),

            "actor_name":
                user.get("name", ""),

            "actor_role":
                user.get("role", ""),

            "action":
                action,

            "remarks":
                remarks,
        },
    )


# ============================================================
# TIME FUNCTIONS
# ============================================================

def generate_time_slots():

    slots = []

    current = datetime.combine(
        date.today(),
        time(5, 0),
    )

    end = datetime.combine(
        date.today(),
        time(23, 30),
    )

    while current <= end:

        slots.append(
            current.strftime("%H:%M")
        )

        current += timedelta(
            minutes=30
        )

    return slots


def make_datetime(
    date_value,
    time_value,
):

    if isinstance(date_value, date):

        actual_date = date_value

    else:

        actual_date = pd.to_datetime(
            date_value
        ).date()

    actual_time = datetime.strptime(
        str(time_value),
        "%H:%M",
    ).time()

    return datetime.combine(
        actual_date,
        actual_time,
    )


# ============================================================
# VEHICLE AVAILABILITY
# ============================================================

def vehicle_is_available(
    data,
    vehicle_number,
    travel_date,
    departure_time,
    return_date,
    return_time,
    ignore_request_id=None,
):

    try:

        requested_start = make_datetime(
            travel_date,
            departure_time,
        )

        requested_end = make_datetime(
            return_date,
            return_time,
        )

    except Exception:

        return False

    if requested_end <= requested_start:

        return False

    bookings = data["GatePasses"]

    ignored_statuses = {
        "Rejected by Manager",
        "Rejected by HR",
        "Cancelled",
        "Trip Completed - Security Approved",
    }

    for _, booking in bookings.iterrows():

        if (
            booking["vehicle_number"]
            != vehicle_number
        ):

            continue

        if (
            ignore_request_id
            and booking["request_id"]
            == ignore_request_id
        ):

            continue

        if booking["status"] in ignored_statuses:

            continue

        try:

            existing_start = make_datetime(
                booking["travel_date"],
                booking["departure_time"],
            )

            existing_end = make_datetime(
                booking["return_date"],
                booking["return_time"],
            )

        except Exception:

            continue

        if (
            requested_start < existing_end
            and requested_end > existing_start
        ):

            return False

    return True


def get_available_vehicles(
    data,
    vehicle_type,
    travel_date,
    departure_time,
    return_date,
    return_time,
):

    vehicles = data["Vehicles"].copy()

    vehicles = vehicles[
        vehicles["vehicle_type"].str.strip()
        == str(vehicle_type).strip()
    ]

    vehicles = vehicles[
        vehicles["active"]
        .str.strip()
        .str.lower()
        .isin(["yes", "true", "1", "active"])
    ]

    available = []

    for _, vehicle in vehicles.iterrows():

        vehicle_number = vehicle[
            "vehicle_number"
        ]

        if vehicle_is_available(
            data,
            vehicle_number,
            travel_date,
            departure_time,
            return_date,
            return_time,
        ):

            available.append(
                vehicle_number
            )

    return available


def get_available_departure_times(
    data,
    vehicle_type,
    travel_date,
    return_date,
    return_time,
):

    available_times = []

    for departure_time in generate_time_slots():

        try:

            departure_dt = make_datetime(
                travel_date,
                departure_time,
            )

            return_dt = make_datetime(
                return_date,
                return_time,
            )

            if return_dt <= departure_dt:

                continue

        except Exception:

            continue

        vehicles = get_available_vehicles(
            data,
            vehicle_type,
            travel_date,
            departure_time,
            return_date,
            return_time,
        )

        if vehicles:

            available_times.append(
                departure_time
            )

    return available_times


# ============================================================
# REQUEST DISPLAY
# ============================================================

def display_request(request):

    col1, col2 = st.columns(2)

    with col1:

        st.write(
            "**Request ID:**",
            request["request_id"],
        )

        st.write(
            "**Requisitioner:**",
            request["requisitioner"],
        )

        st.write(
            "**Department:**",
            request["department"],
        )

        st.write(
            "**Contact:**",
            request["contact"],
        )

        st.write(
            "**Travelling With:**",
            request["companions"],
        )

        st.write(
            "**Destination:**",
            request["destination"],
        )

        st.write(
            "**Purpose:**",
            request["purpose"],
        )

    with col2:

        st.write(
            "**Vehicle Type:**",
            request["vehicle_type"],
        )

        st.write(
            "**Vehicle:**",
            request["vehicle_number"]
            or "Not assigned",
        )

        st.write(
            "**Driver:**",
            request["driver_name"]
            or "Not assigned",
        )

        st.write(
            "**Departure:**",
            f'{request["travel_date"]} '
            f'{request["departure_time"]}',
        )

        st.write(
            "**Return:**",
            f'{request["return_date"]} '
            f'{request["return_time"]}',
        )

        st.write(
            "**Status:**",
            request["status"],
        )


# ============================================================
# EMPLOYEE PORTAL
# ============================================================

def employee_portal():

    st.title(
        "🚐 Vehicle Gate Pass Request"
    )

    st.caption(
        "No password is required for vehicle requests."
    )

    data = load_data()

    departments = sorted(
        [
            x for x in
            data["Departments"]["department"].tolist()
            if x.strip()
        ]
    )

    vehicle_types = sorted(
        [
            x for x in
            data["Vehicles"]["vehicle_type"].tolist()
            if x.strip()
        ]
    )

    if not departments:

        st.error(
            "No departments found in the Departments sheet."
        )

        return

    if not vehicle_types:

        st.error(
            "No vehicle types found in the Vehicles sheet."
        )

        return

    st.subheader(
        "Request a Vehicle"
    )

    with st.form("vehicle_request_form"):

        col1, col2 = st.columns(2)

        requisitioner = col1.text_input(
            "Requisitioner Name *"
        )

        department = col2.selectbox(
            "Department *",
            departments,
        )

        contact = col1.text_input(
            "Contact Number"
        )

        companions = col2.text_input(
            "Person(s) Travelling With"
        )

        destination = st.text_input(
            "Destination *"
        )

        purpose = st.text_area(
            "Purpose *"
        )

        st.markdown(
            "### Trip Details"
        )

        col1, col2 = st.columns(2)

        travel_date = col1.date_input(
            "Travel Date *",
            min_value=date.today(),
        )

        return_date = col2.date_input(
            "Return Date *",
            min_value=date.today(),
        )

        vehicle_type = col1.selectbox(
            "Vehicle Type *",
            vehicle_types,
        )

        return_times = generate_time_slots()

        return_time = col2.selectbox(
            "Return Time *",
            return_times,
            index=(
                return_times.index("17:00")
                if "17:00" in return_times
                else 0
            ),
        )

        # ----------------------------------------------------
        # ONLY SHOW DEPARTURE TIMES WHERE A VEHICLE EXISTS
        # ----------------------------------------------------

        available_departure_times = (
            get_available_departure_times(
                data,
                vehicle_type,
                travel_date,
                return_date,
                return_time,
            )
        )

        if available_departure_times:

            departure_time = st.selectbox(
                "Available Departure Time *",
                available_departure_times,
            )

            available_vehicles = (
                get_available_vehicles(
                    data,
                    vehicle_type,
                    travel_date,
                    departure_time,
                    return_date,
                    return_time,
                )
            )

            if available_vehicles:

                vehicle_number = st.selectbox(
                    "Available Vehicle *",
                    available_vehicles,
                )

                st.success(
                    f"{len(available_vehicles)} "
                    "vehicle(s) available."
                )

            else:

                vehicle_number = ""

        else:

            departure_time = ""

            vehicle_number = ""

            st.warning(
                "No vehicles are available for "
                "the selected date, vehicle type "
                "and return time."
            )

        submitted = st.form_submit_button(
            "Submit Vehicle Request",
            type="primary",
        )

    if submitted:

        if not requisitioner.strip():

            st.error(
                "Please enter the requisitioner's name."
            )

            return

        if not destination.strip():

            st.error(
                "Please enter the destination."
            )

            return

        if not purpose.strip():

            st.error(
                "Please enter the purpose."
            )

            return

        if not departure_time:

            st.error(
                "No available departure time."
            )

            return

        if not vehicle_number:

            st.error(
                "No available vehicle."
            )

            return

        # ----------------------------------------------------
        # RECHECK AVAILABILITY BEFORE SAVING
        # ----------------------------------------------------

        fresh_data = load_data()

        if not vehicle_is_available(
            fresh_data,
            vehicle_number,
            travel_date,
            departure_time,
            return_date,
            return_time,
        ):

            st.error(
                "This vehicle was just booked. "
                "Please select another available time."
            )

            st.rerun()

        # ----------------------------------------------------
        # FIND DEPARTMENT MANAGER
        # ----------------------------------------------------

        department_rows = fresh_data[
            "Departments"
        ]

        manager_rows = department_rows[
            department_rows["department"].str.strip()
            == department.strip()
        ]

        if manager_rows.empty:

            st.error(
                "No manager is configured for "
                "this department."
            )

            return

        manager = manager_rows.iloc[0]

        request_id = (
            "VGP-"
            + datetime.now().strftime("%Y%m%d")
            + "-"
            + uuid.uuid4().hex[:6].upper()
        )

        append_record(
            "GatePasses",
            {
                "request_id":
                    request_id,

                "created_at":
                    datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    ),

                "requisitioner":
                    requisitioner.strip(),

                "department":
                    department,

                "contact":
                    contact.strip(),

                "companions":
                    companions.strip(),

                "destination":
                    destination.strip(),

                "purpose":
                    purpose.strip(),

                "travel_date":
                    travel_date.strftime(
                        "%Y-%m-%d"
                    ),

                "departure_time":
                    departure_time,

                "return_date":
                    return_date.strftime(
                        "%Y-%m-%d"
                    ),

                "return_time":
                    return_time,

                "vehicle_type":
                    vehicle_type,

                "vehicle_number":
                    vehicle_number,

                "driver_username":
                    "",

                "driver_name":
                    "",

                "status":
                    "Pending Department Manager",

                "manager_username":
                    manager[
                        "manager_username"
                    ],

                "manager_name":
                    manager[
                        "manager_name"
                    ],
            },
        )

        add_audit(
            request_id,
            {
                "username":
                    "self-service",

                "name":
                    requisitioner,

                "role":
                    "Requisitioner",
            },
            "REQUEST_SUBMITTED",
        )

        st.success(
            "Vehicle request submitted successfully."
        )

        st.markdown(
            "### Your Request ID"
        )

        st.code(
            request_id
        )

        st.info(
            "Keep this Request ID for tracking."
        )

    # ========================================================
    # REQUEST STATUS
    # ========================================================

    st.divider()

    st.subheader(
        "Check Request Status"
    )

    request_lookup = st.text_input(
        "Request ID",
        key="request_status_lookup",
    )

    if request_lookup.strip():

        current_data = load_data()

        matches = current_data[
            "GatePasses"
        ][
            current_data[
                "GatePasses"
            ]["request_id"].str.strip()
            == request_lookup.strip()
        ]

        if matches.empty:

            st.warning(
                "Request not found."
            )

        else:

            request = matches.iloc[0]

            display_request(
                request
            )


# ============================================================
# MANAGER LOGIN
# ============================================================

def manager_login():

    st.title(
        "👔 Department Manager Login"
    )

    data = load_data()

    departments = data[
        "Departments"
    ]

    if departments.empty:

        st.error(
            "No departments configured."
        )

        return

    department_names = sorted(
        [
            x for x in
            departments["department"].tolist()
            if x.strip()
        ]
    )

    department = st.selectbox(
        "Department",
        department_names,
    )

    username = st.text_input(
        "Manager Username"
    )

    password = st.text_input(
        "Manager Password",
        type="password",
    )

    if st.button(
        "Login",
        type="primary",
    ):

        matches = departments[
            (
                departments["department"].str.strip()
                == department.strip()
            )
            &
            (
                departments["manager_username"].str.strip()
                == username.strip()
            )
        ]

        if matches.empty:

            st.error(
                "Invalid manager username."
            )

            return

        manager = matches.iloc[0]

        if password != manager[
            "manager_password"
        ]:

            st.error(
                "Invalid manager password."
            )

            return

        st.session_state.user = {
            "role":
                "Department Manager",

            "username":
                manager["manager_username"],

            "name":
                manager["manager_name"],

            "department":
                manager["department"],
        }

        st.rerun()


# ============================================================
# MANAGER PORTAL
# ============================================================

def manager_portal():

    st.title(
        "👔 Department Manager Portal"
    )

    user = st.session_state.user

    st.write(
        f"Department: **{user['department']}**"
    )

    data = load_data()

    pending = data[
        "GatePasses"
    ][
        (
            data["GatePasses"]["status"]
            == "Pending Department Manager"
        )
        &
        (
            data["GatePasses"]["manager_username"]
            == user["username"]
        )
    ]

    if pending.empty:

        st.success(
            "No pending vehicle requests."
        )

        return

    request_id = st.selectbox(
        "Select Request",
        pending["request_id"].tolist(),
    )

    request = pending[
        pending["request_id"] == request_id
    ].iloc[0]

    display_request(request)

    col1, col2 = st.columns(2)

    with col1:

        if st.button(
            "✅ Approve Request",
            type="primary",
            use_container_width=True,
        ):

            fresh_data = load_data()

            if not vehicle_is_available(
                fresh_data,
                request["vehicle_number"],
                request["travel_date"],
                request["departure_time"],
                request["return_date"],
                request["return_time"],
                ignore_request_id=request_id,
            ):

                st.error(
                    "The selected vehicle is no longer "
                    "available."
                )

                return

            update_record(
                "GatePasses",
                "request_id",
                request_id,
                {
                    "status":
                        "Pending HR",

                    "manager_approved_at":
                        datetime.now().strftime(
                            "%Y-%m-%d %H:%M:%S"
                        ),
                },
            )

            add_audit(
                request_id,
                user,
                "MANAGER_APPROVED",
            )

            st.success(
                "Approved. Sent to HR."
            )

            st.rerun()

    with col2:

        if st.button(
            "❌ Reject Request",
            use_container_width=True,
        ):

            st.session_state[
                "manager_reject"
            ] = True

    if st.session_state.get(
        "manager_reject",
        False,
    ):

        reason = st.text_area(
            "Reason for rejection"
        )

        if st.button(
            "Confirm Rejection",
            type="primary",
        ):

            if not reason.strip():

                st.warning(
                    "Enter a reason."
                )

                return

            update_record(
                "GatePasses",
                "request_id",
                request_id,
                {
                    "status":
                        "Rejected by Manager",

                    "rejection_reason":
                        reason,
                },
            )

            add_audit(
                request_id,
                user,
                "MANAGER_REJECTED",
                reason,
            )

            st.success(
                "Request rejected."
            )

            st.session_state[
                "manager_reject"
            ] = False

            st.rerun()


# ============================================================
# HR LOGIN
# ============================================================

def hr_login():

    st.title(
        "👩‍💼 HR Manager Login"
    )

    password = st.text_input(
        "HR Password",
        type="password",
    )

    if st.button(
        "Login",
        type="primary",
    ):

        expected = st.secrets.get(
            "hr_password",
            "",
        )

        if password == expected:

            st.session_state.user = {
                "role":
                    "HR Manager",

                "username":
                    "hr_manager",

                "name":
                    "HR Manager",
            }

            st.rerun()

        else:

            st.error(
                "Invalid HR password."
            )


# ============================================================
# HR PORTAL
# ============================================================

def hr_portal():

    st.title(
        "👩‍💼 HR Approval"
    )

    user = st.session_state.user

    data = load_data()

    pending = data[
        "GatePasses"
    ][
        data["GatePasses"]["status"]
        == "Pending HR"
    ]

    if pending.empty:

        st.success(
            "No requests awaiting HR approval."
        )

        return

    request_id = st.selectbox(
        "Select Request",
        pending["request_id"].tolist(),
    )

    request = pending[
        pending["request_id"] == request_id
    ].iloc[0]

    display_request(request)

    col1, col2 = st.columns(2)

    with col1:

        if st.button(
            "✅ Approve HR",
            type="primary",
            use_container_width=True,
        ):

            update_record(
                "GatePasses",
                "request_id",
                request_id,
                {
                    "status":
                        "Pending Security",

                    "hr_approved_at":
                        datetime.now().strftime(
                            "%Y-%m-%d %H:%M:%S"
                        ),
                },
            )

            add_audit(
                request_id,
                user,
                "HR_APPROVED",
            )

            st.success(
                "Approved. Sent to Security."
            )

            st.rerun()

    with col2:

        if st.button(
            "❌ Reject HR",
            use_container_width=True,
        ):

            st.session_state[
                "hr_reject"
            ] = True

    if st.session_state.get(
        "hr_reject",
        False,
    ):

        reason = st.text_area(
            "HR Rejection Reason"
        )

        if st.button(
            "Confirm HR Rejection",
            type="primary",
        ):

            if not reason.strip():

                st.warning(
                    "Enter a reason."
                )

                return

            update_record(
                "GatePasses",
                "request_id",
                request_id,
                {
                    "status":
                        "Rejected by HR",

                    "rejection_reason":
                        reason,
                },
            )

            add_audit(
                request_id,
                user,
                "HR_REJECTED",
                reason,
            )

            st.success(
                "Request rejected."
            )

            st.session_state[
                "hr_reject"
            ] = False

            st.rerun()


# ============================================================
# SECURITY LOGIN
# ============================================================

def security_login():

    st.title(
        "🛡️ Security Login"
    )

    password = st.text_input(
        "Security Password",
        type="password",
    )

    if st.button(
        "Login",
        type="primary",
    ):

        expected = st.secrets.get(
            "security_password",
            "",
        )

        if password == expected:

            st.session_state.user = {
                "role":
                    "Security",

                "username":
                    "security",

                "name":
                    "Security",
            }

            st.rerun()

        else:

            st.error(
                "Invalid security password."
            )


# ============================================================
# SECURITY PORTAL
# ============================================================

def security_portal():

    st.title(
        "🛡️ Security Portal"
    )

    user = st.session_state.user

    data = load_data()

    # ========================================================
    # RELEASE VEHICLE
    # ========================================================

    st.subheader(
        "Approved Requests"
    )

    pending = data[
        "GatePasses"
    ][
        data["GatePasses"]["status"]
        == "Pending Security"
    ]

    if pending.empty:

        st.info(
            "No requests waiting for Security."
        )

    else:

        request_id = st.selectbox(
            "Select Request",
            pending["request_id"].tolist(),
            key="security_release",
        )

        request = pending[
            pending["request_id"] == request_id
        ].iloc[0]

        display_request(request)

        drivers = data["Drivers"]

        drivers = drivers[
            drivers["active"]
            .str.strip()
            .str.lower()
            .isin(
                ["yes", "true", "1", "active"]
            )
        ]

        if drivers.empty:

            st.error(
                "No active drivers found."
            )

        else:

            driver_username = st.selectbox(
                "Assign Driver",
                drivers[
                    "driver_username"
                ].tolist(),
                format_func=lambda x:
                    drivers[
                        drivers[
                            "driver_username"
                        ] == x
                    ].iloc[0][
                        "driver_name"
                    ],
            )

            if st.button(
                "🚐 Verify & Release Vehicle",
                type="primary",
            ):

                driver = drivers[
                    drivers["driver_username"]
                    == driver_username
                ].iloc[0]

                update_record(
                    "GatePasses",
                    "request_id",
                    request_id,
                    {
                        "driver_username":
                            driver[
                                "driver_username"
                            ],

                        "driver_name":
                            driver[
                                "driver_name"
                            ],

                        "security_name":
                            user["name"],

                        "security_released_at":
                            datetime.now().strftime(
                                "%Y-%m-%d %H:%M:%S"
                            ),

                        "status":
                            "Vehicle Released",
                    },
                )

                add_audit(
                    request_id,
                    user,
                    "SECURITY_RELEASED",
                    driver["driver_name"],
                )

                st.success(
                    "Vehicle released."
                )

                st.rerun()

    # ========================================================
    # FINAL TRIP VERIFICATION
    # ========================================================

    st.divider()

    st.subheader(
        "Completed Trips Awaiting Verification"
    )

    completed = data[
        "GatePasses"
    ][
        data["GatePasses"]["status"]
        == "Pending Security Verification"
    ]

    if completed.empty:

        st.info(
            "No completed trips waiting for verification."
        )

        return

    request_id = st.selectbox(
        "Select Completed Trip",
        completed["request_id"].tolist(),
        key="security_completed",
    )

    request = completed[
        completed["request_id"] == request_id
    ].iloc[0]

    display_request(request)

    st.write(
        f"**Start Mileage:** "
        f"{request['start_mileage']}"
    )

    st.write(
        f"**End Mileage:** "
        f"{request['end_mileage']}"
    )

    try:

        start = float(
            request["start_mileage"]
        )

        end = float(
            request["end_mileage"]
        )

        if end >= start:

            st.metric(
                "Distance Travelled",
                f"{end - start:.1f}",
            )

    except Exception:

        pass

    notes = st.text_area(
        "Security Verification Notes"
    )

    if st.button(
        "✅ Verify & Approve Trip",
        type="primary",
    ):

        update_record(
            "GatePasses",
            "request_id",
            request_id,
            {
                "security_verified_at":
                    datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    ),

                "security_name":
                    user["name"],

                "security_notes":
                    notes,

                "status":
                    "Trip Completed - Security Approved",
            },
        )

        add_audit(
            request_id,
            user,
            "SECURITY_FINAL_APPROVAL",
            notes,
        )

        st.success(
            "Trip verified and approved."
        )

        st.rerun()


# ============================================================
# DRIVER LOGIN
# ============================================================

def driver_login():

    st.title(
        "🚐 Driver Login"
    )

    data = load_data()

    username = st.text_input(
        "Driver Username"
    )

    password = st.text_input(
        "Driver Password",
        type="password",
    )

    if st.button(
        "Login",
        type="primary",
    ):

        drivers = data["Drivers"]

        matches = drivers[
            (
                drivers["driver_username"].str.strip()
                == username.strip()
            )
            &
            (
                drivers["active"]
                .str.strip()
                .str.lower()
                .isin(
                    ["yes", "true", "1", "active"]
                )
            )
        ]

        if matches.empty:

            st.error(
                "Driver not found or inactive."
            )

            return

        driver = matches.iloc[0]

        if password != driver[
            "driver_password"
        ]:

            st.error(
                "Invalid driver password."
            )

            return

        st.session_state.user = {
            "role":
                "Driver",

            "username":
                driver["driver_username"],

            "name":
                driver["driver_name"],
        }

        st.rerun()


# ============================================================
# DRIVER PORTAL
# ============================================================

def driver_portal():

    st.title(
        "🚐 Driver Portal"
    )

    user = st.session_state.user

    st.write(
        f"Driver: **{user['name']}**"
    )

    data = load_data()

    assigned = data[
        "GatePasses"
    ][
        (
            data["GatePasses"]["driver_username"]
            == user["username"]
        )
        &
        (
            data["GatePasses"]["status"].isin(
                [
                    "Vehicle Released",
                    "Trip In Progress",
                ]
            )
        )
    ]

    if assigned.empty:

        st.success(
            "No active trips assigned to you."
        )

        return

    request_id = st.selectbox(
        "Assigned Trip",
        assigned["request_id"].tolist(),
    )

    request = assigned[
        assigned["request_id"] == request_id
    ].iloc[0]

    display_request(request)

    # ========================================================
    # START TRIP
    # ========================================================

    if request["status"] == "Vehicle Released":

        st.subheader(
            "Start Trip"
        )

        start_mileage = st.number_input(
            "Starting Mileage",
            min_value=0.0,
            step=0.1,
        )

        if st.button(
            "▶️ Record Starting Mileage",
            type="primary",
        ):

            update_record(
                "GatePasses",
                "request_id",
                request_id,
                {
                    "start_mileage":
                        start_mileage,

                    "driver_started_at":
                        datetime.now().strftime(
                            "%Y-%m-%d %H:%M:%S"
                        ),

                    "status":
                        "Trip In Progress",
                },
            )

            add_audit(
                request_id,
                user,
                "DRIVER_STARTED",
                str(start_mileage),
            )

            st.success(
                "Starting mileage recorded."
            )

            st.rerun()

    # ========================================================
    # END TRIP
    # ========================================================

    else:

        st.subheader(
            "Complete Trip"
        )

        try:

            start_mileage = float(
                request["start_mileage"]
            )

        except Exception:

            start_mileage = 0.0

        end_mileage = st.number_input(
            "Ending Mileage",
            min_value=start_mileage,
            value=start_mileage,
            step=0.1,
        )

        if st.button(
            "⏹️ Record Ending Mileage",
            type="primary",
        ):

            if end_mileage < start_mileage:

                st.error(
                    "Ending mileage cannot be "
                    "less than starting mileage."
                )

                return

            update_record(
                "GatePasses",
                "request_id",
                request_id,
                {
                    "end_mileage":
                        end_mileage,

                    "driver_completed_at":
                        datetime.now().strftime(
                            "%Y-%m-%d %H:%M:%S"
                        ),

                    "status":
                        "Pending Security Verification",
                },
            )

            add_audit(
                request_id,
                user,
                "DRIVER_COMPLETED",
                str(end_mileage),
            )

            st.success(
                "Trip completed. Waiting for Security verification."
            )

            st.rerun()


# ============================================================
# ADMIN LOGIN
# ============================================================

def admin_login():

    st.title(
        "⚙️ Administration Login"
    )

    password = st.text_input(
        "Administration Password",
        type="password",
    )

    if st.button(
        "Login",
        type="primary",
    ):

        expected = st.secrets.get(
            "admin_password",
            "",
        )

        if password == expected:

            st.session_state.user = {
                "role":
                    "Administration",

                "username":
                    "admin",

                "name":
                    "Administrator",
            }

            st.rerun()

        else:

            st.error(
                "Invalid administration password."
            )


# ============================================================
# ADMIN DASHBOARD
# ============================================================

def admin_portal():

    st.title(
        "⚙️ Administration Dashboard"
    )

    data = load_data()

    requests = data["GatePasses"]

    col1, col2, col3, col4 = st.columns(4)

    col1.metric(
        "Total Requests",
        len(requests),
    )

    col2.metric(
        "Pending Manager",
        len(
            requests[
                requests["status"]
                == "Pending Department Manager"
            ]
        ),
    )

    col3.metric(
        "Pending HR",
        len(
            requests[
                requests["status"]
                == "Pending HR"
            ]
        ),
    )

    col4.metric(
        "Pending Security",
        len(
            requests[
                requests["status"]
                == "Pending Security"
            ]
        ),
    )

    st.divider()

    st.subheader(
        "Vehicle Requests"
    )

    if requests.empty:

        st.info(
            "No requests yet."
        )

    else:

        st.dataframe(
            requests,
            use_container_width=True,
            hide_index=True,
        )

    st.divider()

    st.subheader(
        "Departments"
    )

    st.dataframe(
        data["Departments"],
        use_container_width=True,
        hide_index=True,
    )

    st.subheader(
        "Drivers"
    )

    st.dataframe(
        data["Drivers"],
        use_container_width=True,
        hide_index=True,
    )

    st.subheader(
        "Vehicles"
    )

    st.dataframe(
        data["Vehicles"],
        use_container_width=True,
        hide_index=True,
    )

    st.subheader(
        "Approval Audit"
    )

    st.dataframe(
        data["ApprovalAudit"],
        use_container_width=True,
        hide_index=True,
    )


# ============================================================
# SESSION STATE
# ============================================================

if "user" not in st.session_state:

    st.session_state.user = None


# ============================================================
# TEST GOOGLE CONNECTION
# ============================================================

try:

    get_spreadsheet()

except Exception as error:

    st.error(
        "❌ Google Sheets connection failed."
    )

    st.warning(
        "Check your Streamlit Secrets and make sure "
        "the Google Sheet is shared with the "
        "service-account email."
    )

    with st.expander(
        "Technical Error"
    ):

        st.exception(error)

    st.stop()


# ============================================================
# LOGGED-IN USER
# ============================================================

if st.session_state.user:

    user = st.session_state.user

    st.sidebar.success(
        f"👤 {user['name']}\n\n"
        f"Role: {user['role']}"
    )

    if st.sidebar.button(
        "🚪 Logout"
    ):

        st.session_state.user = None

        st.rerun()

    if user["role"] == "Department Manager":

        manager_portal()

    elif user["role"] == "HR Manager":

        hr_portal()

    elif user["role"] == "Security":

        security_portal()

    elif user["role"] == "Driver":

        driver_portal()

    elif user["role"] == "Administration":

        admin_portal()


# ============================================================
# LOGIN / PORTAL SELECTION
# ============================================================

else:

    st.sidebar.title(
        "🚐 Vehicle Gate Pass"
    )

    portal = st.sidebar.radio(
        "Select Portal",
        [
            "Employee / Requisitioner",
            "Department Manager",
            "HR Manager",
            "Security",
            "Driver",
            "Administration",
        ],
    )

    if portal == "Employee / Requisitioner":

        employee_portal()

    elif portal == "Department Manager":

        manager_login()

    elif portal == "HR Manager":

        hr_login()

    elif portal == "Security":

        security_login()

    elif portal == "Driver":

        driver_login()

    elif portal == "Administration":

        admin_login()
