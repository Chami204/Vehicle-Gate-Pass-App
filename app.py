import uuid
from datetime import datetime, date, time, timedelta

import pandas as pd
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials


# ============================================================
# CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Vehicle Gate Pass",
    page_icon="🚐",
    layout="wide",
)

# Your Google Spreadsheet name
GOOGLE_SHEET_NAME = "vehicle_gate_pass_data"


# ============================================================
# GOOGLE SHEETS CONNECTION
# ============================================================

@st.cache_resource
def get_google_credentials():
    service_account_info = dict(st.secrets["google_service_account"])

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    return Credentials.from_service_account_info(
        service_account_info,
        scopes=scopes,
    )


@st.cache_resource
def get_google_client():
    credentials = get_google_credentials()
    return gspread.authorize(credentials)


@st.cache_resource
def get_spreadsheet():
    client = get_google_client()
    return client.open(GOOGLE_SHEET_NAME)


# ============================================================
# GOOGLE SHEET STRUCTURE
# ============================================================

SHEET_COLUMNS = {
    "Departments": [
        "department",
        "manager_username",
        "manager_name",
    ],

    "Drivers": [
        "driver_username",
        "driver_name",
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
# DEFAULT DATA
# Only inserted if the respective sheet is empty.
# ============================================================

DEFAULT_DEPARTMENTS = [
    ["Administration", "admin_manager", "Administration Manager"],
    ["Finance", "finance_manager", "Finance Manager"],
    ["HR", "hr_manager", "HR Manager"],
    ["Sales", "sales_manager", "Sales Manager"],
    ["Operations", "operations_manager", "Operations Manager"],
    ["IT", "it_manager", "IT Manager"],
]

DEFAULT_DRIVERS = [
    ["driver01", "Driver One", "Yes"],
    ["driver02", "Driver Two", "Yes"],
    ["driver03", "Driver Three", "Yes"],
]

DEFAULT_VEHICLES = [
    ["VAN-01", "Van", "Yes"],
    ["VAN-02", "Van", "Yes"],
    ["VAN-03", "Van", "Yes"],
]


# ============================================================
# GOOGLE SHEET HELPERS
# ============================================================

def get_worksheet(sheet_name):
    return get_spreadsheet().worksheet(sheet_name)


def create_required_sheets():
    spreadsheet = get_spreadsheet()

    existing_sheets = {
        worksheet.title
        for worksheet in spreadsheet.worksheets()
    }

    for sheet_name, columns in SHEET_COLUMNS.items():

        if sheet_name not in existing_sheets:

            worksheet = spreadsheet.add_worksheet(
                title=sheet_name,
                rows=2000,
                cols=max(20, len(columns) + 2),
            )

            worksheet.update(
                "A1",
                [columns],
            )

    # Add initial data only if sheets are empty.

    departments = get_worksheet("Departments")

    if len(departments.get_all_values()) <= 1:
        departments.append_rows(
            DEFAULT_DEPARTMENTS,
            value_input_option="USER_ENTERED",
        )

    drivers = get_worksheet("Drivers")

    if len(drivers.get_all_values()) <= 1:
        drivers.append_rows(
            DEFAULT_DRIVERS,
            value_input_option="USER_ENTERED",
        )

    vehicles = get_worksheet("Vehicles")

    if len(vehicles.get_all_values()) <= 1:
        vehicles.append_rows(
            DEFAULT_VEHICLES,
            value_input_option="USER_ENTERED",
        )


@st.cache_data(ttl=10)
def read_sheet(sheet_name):

    worksheet = get_worksheet(sheet_name)

    records = worksheet.get_all_records()

    return pd.DataFrame(
        records,
        columns=SHEET_COLUMNS[sheet_name],
    ).fillna("").astype(str)


def load_all_data():

    create_required_sheets()

    return {
        sheet_name: read_sheet(sheet_name)
        for sheet_name in SHEET_COLUMNS
    }


def refresh_data():

    read_sheet.clear()


def append_record(sheet_name, record):

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

    refresh_data()


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

    row_number = None

    for row_index, row in enumerate(values[1:], start=2):

        if (
            len(row) > key_index
            and str(row[key_index]) == str(key_value)
        ):
            row_number = row_index
            break

    if row_number is None:
        return False

    for column, value in changes.items():

        if column in headers:

            column_index = headers.index(column) + 1

            worksheet.update_cell(
                row_number,
                column_index,
                str(value),
            )

    refresh_data()

    return True


# ============================================================
# SECRETS
# ============================================================

def get_secret_password(name):

    return st.secrets.get(name)


def get_nested_password(section, username):

    try:

        passwords = dict(
            st.secrets[section]
        )

        return passwords.get(username)

    except Exception:

        return None


# ============================================================
# AUDIT
# ============================================================

def create_audit(
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
# DATE / TIME
# ============================================================

def convert_datetime(
    date_value,
    time_value,
):

    if isinstance(date_value, date):
        converted_date = date_value
    else:
        converted_date = pd.to_datetime(
            date_value
        ).date()

    converted_time = datetime.strptime(
        str(time_value),
        "%H:%M",
    ).time()

    return datetime.combine(
        converted_date,
        converted_time,
    )


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
):

    try:

        requested_start = convert_datetime(
            travel_date,
            departure_time,
        )

        requested_end = convert_datetime(
            return_date,
            return_time,
        )

    except Exception:

        return False

    if requested_end <= requested_start:
        return False

    ignored_statuses = {
        "Rejected by Manager",
        "Rejected by HR",
        "Cancelled",
    }

    bookings = data["GatePasses"]

    for _, booking in bookings.iterrows():

        if (
            booking["vehicle_number"]
            != vehicle_number
        ):
            continue

        if booking["status"] in ignored_statuses:
            continue

        required_fields = [
            "travel_date",
            "departure_time",
            "return_date",
            "return_time",
        ]

        if not all(
            booking[field]
            for field in required_fields
        ):
            continue

        try:

            existing_start = convert_datetime(
                booking["travel_date"],
                booking["departure_time"],
            )

            existing_end = convert_datetime(
                booking["return_date"],
                booking["return_time"],
            )

            # Overlap test
            if (
                requested_start < existing_end
                and requested_end > existing_start
            ):
                return False

        except Exception:
            continue

    return True


def get_available_vehicles(
    data,
    vehicle_type,
    travel_date,
    departure_time,
    return_date,
    return_time,
):

    vehicles = data["Vehicles"]

    vehicles = vehicles[
        (
            vehicles["vehicle_type"]
            == vehicle_type
        )
        &
        (
            vehicles["active"]
            .str.lower()
            == "yes"
        )
    ]

    available = []

    for _, vehicle in vehicles.iterrows():

        if vehicle_is_available(
            data,
            vehicle["vehicle_number"],
            travel_date,
            departure_time,
            return_date,
            return_time,
        ):

            available.append(
                vehicle["vehicle_number"]
            )

    return available


# ============================================================
# DISPLAY REQUEST
# ============================================================

def display_request(request):

    left, right = st.columns(2)

    with left:

        st.write(
            "**Requisitioner:**",
            request["requisitioner"],
        )

        st.write(
            "**Department:**",
            request["department"],
        )

        st.write(
            "**Destination:**",
            request["destination"],
        )

        st.write(
            "**Purpose:**",
            request["purpose"],
        )

        st.write(
            "**Travelling With:**",
            request["companions"],
        )

    with right:

        st.write(
            "**Vehicle:**",
            request["vehicle_number"],
        )

        st.write(
            "**Trip:**",
            f'{request["travel_date"]} '
            f'{request["departure_time"]} → '
            f'{request["return_date"]} '
            f'{request["return_time"]}',
        )

        st.write(
            "**Driver:**",
            request["driver_name"]
            or "Not assigned",
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
        "Employees do not need a password to request a vehicle."
    )

    data = load_all_data()

    departments = (
        data["Departments"]
        ["department"]
        .tolist()
    )

    vehicle_types = sorted(
        data["Vehicles"]
        ["vehicle_type"]
        .unique()
        .tolist()
    )

    if not departments:

        st.error(
            "No departments are configured."
        )

        return

    if not vehicle_types:

        st.error(
            "No vehicles are configured."
        )

        return

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
            "Purpose of Trip *"
        )

        col1, col2 = st.columns(2)

        travel_date = col1.date_input(
            "Travel Date",
            min_value=date.today(),
        )

        return_date = col2.date_input(
            "Return Date",
            min_value=date.today(),
        )

        col1, col2 = st.columns(2)

        vehicle_type = col1.selectbox(
            "Vehicle Type",
            vehicle_types,
        )

        departure_time = col2.selectbox(
            "Departure Time",
            generate_time_slots(),
        )

        departure_datetime = datetime.combine(
            travel_date,
            datetime.strptime(
                departure_time,
                "%H:%M",
            ).time(),
        )

        return_time_options = [
            slot
            for slot in generate_time_slots()
            if datetime.combine(
                return_date,
                datetime.strptime(
                    slot,
                    "%H:%M",
                ).time(),
            ) > departure_datetime
        ]

        return_time = col1.selectbox(
            "Return Time",
            return_time_options
            or generate_time_slots(),
        )

        available = get_available_vehicles(
            data,
            vehicle_type,
            travel_date,
            departure_time,
            return_date,
            return_time,
        )

        if available:

            vehicle_number = col2.selectbox(
                "Available Vehicle",
                available,
            )

            st.success(
                f"{len(available)} vehicle(s) "
                "available for this period."
            )

        else:

            vehicle_number = ""

            st.error(
                "No vehicle is available "
                "for the selected date/time."
            )

        submitted = st.form_submit_button(
            "Submit Vehicle Request",
            type="primary",
        )

    if submitted:

        if (
            not requisitioner
            or not destination
            or not purpose
        ):

            st.error(
                "Please complete all required fields."
            )

            return

        if not available:

            st.error(
                "No vehicle is available."
            )

            return

        # Re-check availability immediately
        # before creating the booking.

        fresh_data = load_all_data()

        if not vehicle_is_available(
            fresh_data,
            vehicle_number,
            travel_date,
            departure_time,
            return_date,
            return_time,
        ):

            st.error(
                "The vehicle was just reserved. "
                "Please select another time."
            )

            return

        manager = fresh_data[
            "Departments"
        ][
            fresh_data["Departments"]
            ["department"]
            == department
        ]

        if manager.empty:

            st.error(
                "No manager is configured "
                "for this department."
            )

            return

        manager = manager.iloc[0]

        request_id = (
            "VGP-"
            + datetime.now().strftime(
                "%Y%m%d"
            )
            + "-"
            + uuid.uuid4()
            .hex[:6]
            .upper()
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
                    requisitioner,

                "department":
                    department,

                "contact":
                    contact,

                "companions":
                    companions,

                "destination":
                    destination,

                "purpose":
                    purpose,

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

        create_audit(
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
            f"Request submitted successfully: "
            f"{request_id}"
        )

        st.info(
            f"Sent to {manager['manager_name']}."
        )

        st.rerun()

    st.divider()

    st.subheader(
        "Check Request Status"
    )

    request_id_check = st.text_input(
        "Enter Request ID"
    )

    if request_id_check:

        matches = data[
            "GatePasses"
        ][
            data["GatePasses"]
            ["request_id"]
            == request_id_check
        ]

        if matches.empty:

            st.warning(
                "Request not found."
            )

        else:

            request = matches.iloc[0]

            st.info(
                f"{request_id_check} — "
                f"{request['status']}"
            )


# ============================================================
# MANAGER LOGIN
# ============================================================

def manager_login():

    data = load_all_data()

    st.subheader(
        "Department Manager Login"
    )

    departments = data[
        "Departments"
    ]

    department = st.selectbox(
        "Department",
        departments[
            "department"
        ].tolist(),
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

        match = departments[
            (
                departments[
                    "department"
                ]
                == department
            )
            &
            (
                departments[
                    "manager_username"
                ]
                == username
            )
        ]

        expected_password = (
            get_nested_password(
                "manager_passwords",
                username,
            )
        )

        if (
            not match.empty
            and expected_password
            and password == expected_password
        ):

            manager = match.iloc[0]

            st.session_state.user = {
                "role":
                    "Department Manager",

                "username":
                    username,

                "name":
                    manager[
                        "manager_name"
                    ],

                "department":
                    manager[
                        "department"
                    ],
            }

            st.rerun()

        st.error(
            "Invalid manager credentials."
        )


# ============================================================
# MANAGER PORTAL
# ============================================================

def manager_portal():

    data = load_all_data()

    user = st.session_state.user

    pending = data[
        "GatePasses"
    ][
        (
            data["GatePasses"]
            ["status"]
            == "Pending Department Manager"
        )
        &
        (
            data["GatePasses"]
            ["department"]
            == user["department"]
        )
    ]

    st.subheader(
        "Department Manager Approval"
    )

    if pending.empty:

        st.success(
            "No pending requests."
        )

        return

    request_id = st.selectbox(
        "Select Request",
        pending[
            "request_id"
        ].tolist(),
    )

    request = pending[
        pending["request_id"]
        == request_id
    ].iloc[0]

    display_request(request)

    if st.button(
        "Approve Request",
        type="primary",
    ):

        fresh = load_all_data()

        if not vehicle_is_available(
            fresh,
            request[
                "vehicle_number"
            ],
            request[
                "travel_date"
            ],
            request[
                "departure_time"
            ],
            request[
                "return_date"
            ],
            request[
                "return_time"
            ],
        ):

            st.error(
                "The vehicle is no longer available."
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

        create_audit(
            request_id,
            user,
            "MANAGER_APPROVED",
        )

        st.success(
            "Approved and sent to HR."
        )

        st.rerun()

    rejection_reason = st.text_area(
        "Rejection Reason"
    )

    if st.button(
        "Reject Request"
    ):

        if not rejection_reason.strip():

            st.warning(
                "Please enter a rejection reason."
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
                    rejection_reason,
            },
        )

        create_audit(
            request_id,
            user,
            "MANAGER_REJECTED",
            rejection_reason,
        )

        st.success(
            "Request rejected."
        )

        st.rerun()


# ============================================================
# HR LOGIN
# ============================================================

def hr_login():

    st.subheader(
        "HR Manager Login"
    )

    password = st.text_input(
        "HR Password",
        type="password",
    )

    if st.button(
        "Login",
        type="primary",
    ):

        if password == get_secret_password(
            "hr_password"
        ):

            st.session_state.user = {
                "role":
                    "HR Manager",

                "username":
                    "hr_manager",

                "name":
                    "HR Manager",
            }

            st.rerun()

        st.error(
            "Invalid HR password."
        )


# ============================================================
# HR PORTAL
# ============================================================

def hr_portal():

    data = load_all_data()

    pending = data[
        "GatePasses"
    ][
        data["GatePasses"]
        ["status"]
        == "Pending HR"
    ]

    st.subheader(
        "HR Approval"
    )

    if pending.empty:

        st.success(
            "No pending HR requests."
        )

        return

    request_id = st.selectbox(
        "Select Request",
        pending[
            "request_id"
        ].tolist(),
    )

    request = pending[
        pending["request_id"]
        == request_id
    ].iloc[0]

    display_request(request)

    if st.button(
        "Approve HR",
        type="primary",
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

        create_audit(
            request_id,
            st.session_state.user,
            "HR_APPROVED",
        )

        st.success(
            "HR approval completed."
        )

        st.rerun()

    rejection_reason = st.text_area(
        "Rejection Reason"
    )

    if st.button(
        "Reject HR"
    ):

        if not rejection_reason.strip():

            st.warning(
                "Please enter a rejection reason."
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
                    rejection_reason,
            },
        )

        create_audit(
            request_id,
            st.session_state.user,
            "HR_REJECTED",
            rejection_reason,
        )

        st.success(
            "Request rejected."
        )

        st.rerun()


# ============================================================
# SECURITY LOGIN
# ============================================================

def security_login():

    st.subheader(
        "Security Login"
    )

    password = st.text_input(
        "Security Password",
        type="password",
    )

    if st.button(
        "Login",
        type="primary",
    ):

        if password == get_secret_password(
            "security_password"
        ):

            st.session_state.user = {
                "role":
                    "Security",

                "username":
                    "security",

                "name":
                    "Security",
            }

            st.rerun()

        st.error(
            "Invalid Security password."
        )


# ============================================================
# SECURITY PORTAL
# ============================================================

def security_portal():

    data = load_all_data()

    user = st.session_state.user

    st.subheader(
        "Security"
    )

    pending = data[
        "GatePasses"
    ][
        data["GatePasses"]
        ["status"]
        == "Pending Security"
    ]

    if not pending.empty:

        request_id = st.selectbox(
            "Request to Release",
            pending[
                "request_id"
            ].tolist(),
        )

        request = pending[
            pending["request_id"]
            == request_id
        ].iloc[0]

        display_request(request)

        drivers = data["Drivers"]

        drivers = drivers[
            drivers["active"]
            .str.lower()
            == "yes"
        ]

        if not drivers.empty:

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
                "Verify & Release Vehicle",
                type="primary",
            ):

                driver = drivers[
                    drivers[
                        "driver_username"
                    ]
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

                create_audit(
                    request_id,
                    user,
                    "SECURITY_RELEASED",
                    driver[
                        "driver_name"
                    ],
                )

                st.success(
                    "Vehicle released."
                )

                st.rerun()

    else:

        st.info(
            "No HR-approved requests waiting for release."
        )

    st.divider()

    st.subheader(
        "Completed Trips Awaiting Security Verification"
    )

    completed = data[
        "GatePasses"
    ][
        data["GatePasses"]
        ["status"]
        == "Pending Security Verification"
    ]

    if completed.empty:

        st.info(
            "No completed trips waiting for verification."
        )

        return

    request_id = st.selectbox(
        "Completed Trip",
        completed[
            "request_id"
        ].tolist(),
        key="completed_trip",
    )

    request = completed[
        completed["request_id"]
        == request_id
    ].iloc[0]

    display_request(request)

    st.write(
        "**Start Mileage:**",
        request["start_mileage"],
    )

    st.write(
        "**End Mileage:**",
        request["end_mileage"],
    )

    notes = st.text_area(
        "Security Verification Notes"
    )

    if st.button(
        "Verify & Approve Trip",
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

        create_audit(
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

    data = load_all_data()

    st.subheader(
        "Driver Login"
    )

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

        match = drivers[
            (
                drivers["driver_username"]
                == username
            )
            &
            (
                drivers["active"]
                .str.lower()
                == "yes"
            )
        ]

        expected_password = (
            get_nested_password(
                "driver_passwords",
                username,
            )
        )

        if (
            not match.empty
            and expected_password
            and password == expected_password
        ):

            driver = match.iloc[0]

            st.session_state.user = {
                "role":
                    "Driver",

                "username":
                    username,

                "name":
                    driver["driver_name"],
            }

            st.rerun()

        st.error(
            "Invalid driver credentials."
        )


# ============================================================
# DRIVER PORTAL
# ============================================================

def driver_portal():

    data = load_all_data()

    user = st.session_state.user

    assigned = data[
        "GatePasses"
    ][
        (
            data["GatePasses"]
            ["driver_username"]
            == user["username"]
        )
        &
        (
            data["GatePasses"]
            ["status"]
            .isin(
                [
                    "Vehicle Released",
                    "Trip In Progress",
                ]
            )
        )
    ]

    st.subheader(
        "Driver Trip Management"
    )

    if assigned.empty:

        st.info(
            "You have no assigned trips."
        )

        return

    request_id = st.selectbox(
        "Assigned Trip",
        assigned[
            "request_id"
        ].tolist(),
    )

    request = assigned[
        assigned["request_id"]
        == request_id
    ].iloc[0]

    display_request(request)

    if request["status"] == "Vehicle Released":

        start_mileage = st.number_input(
            "Start Mileage",
            min_value=0.0,
            step=0.1,
        )

        if st.button(
            "Record Start Mileage",
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

            create_audit(
                request_id,
                user,
                "DRIVER_STARTED",
                str(start_mileage),
            )

            st.success(
                "Start mileage recorded."
            )

            st.rerun()

    elif request["status"] == "Trip In Progress":

        try:
            start_value = float(
                request["start_mileage"]
            )
        except Exception:
            start_value = 0.0

        end_mileage = st.number_input(
            "End Mileage",
            min_value=start_value,
            value=start_value,
            step=0.1,
        )

        if st.button(
            "Record End Mileage",
            type="primary",
        ):

            if end_mileage < start_value:

                st.error(
                    "End mileage cannot be "
                    "less than start mileage."
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

            create_audit(
                request_id,
                user,
                "DRIVER_COMPLETED",
                str(end_mileage),
            )

            st.success(
                "Trip completed. "
                "Security can now verify it."
            )

            st.rerun()


# ============================================================
# ADMINISTRATION LOGIN
# ============================================================

def admin_login():

    st.subheader(
        "Administration Login"
    )

    password = st.text_input(
        "Admin Password",
        type="password",
    )

    if st.button(
        "Login",
        type="primary",
    ):

        if password == get_secret_password(
            "admin_password"
        ):

            st.session_state.user = {
                "role":
                    "Administration",

                "username":
                    "admin",

                "name":
                    "Administrator",
            }

            st.rerun()

        st.error(
            "Invalid admin password."
        )


# ============================================================
# ADMIN DASHBOARD
# ============================================================

def admin_portal():

    data = load_all_data()

    requests = data["GatePasses"]

    st.subheader(
        "Administration Dashboard"
    )

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
        "All Vehicle Requests"
    )

    if requests.empty:

        st.info(
            "No vehicle requests yet."
        )

    else:

        st.dataframe(
            requests,
            use_container_width=True,
            hide_index=True,
        )


# ============================================================
# MAIN APPLICATION
# ============================================================

if "user" not in st.session_state:

    st.session_state.user = None


# Test Google connection

try:

    create_required_sheets()

except Exception as error:

    st.error(
        "Unable to connect to the Google Spreadsheet."
    )

    st.error(
        "Check Streamlit Secrets, Google APIs, "
        "and the spreadsheet sharing permission."
    )

    st.exception(error)

    st.stop()


# ============================================================
# LOGGED-IN USER
# ============================================================

if st.session_state.user:

    user = st.session_state.user

    st.sidebar.success(
        f"{user['name']}\n\n"
        f"{user['role']}"
    )

    if st.sidebar.button(
        "Logout"
    ):

        st.session_state.clear()

        st.rerun()

    role = user["role"]

    if role == "Department Manager":

        manager_portal()

    elif role == "HR Manager":

        hr_portal()

    elif role == "Security":

        security_portal()

    elif role == "Driver":

        driver_portal()

    elif role == "Administration":

        admin_portal()


# ============================================================
# NOT LOGGED IN
# ============================================================

else:

    st.sidebar.title(
        "🚐 Vehicle Gate Pass"
    )

    selected_portal = st.sidebar.radio(
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

    if selected_portal == "Employee / Requisitioner":

        employee_portal()

    elif selected_portal == "Department Manager":

        manager_login()

    elif selected_portal == "HR Manager":

        hr_login()

    elif selected_portal == "Security":

        security_login()

    elif selected_portal == "Driver":

        driver_login()

    elif selected_portal == "Administration":

        admin_login()
