"""Vehicle Gate Pass Management System.

Google Sheets is the only application data store.
Google service-account credentials are read only from Streamlit Secrets.
No Excel/local data file is created by this application.

Email workflow:
    Employee
        -> Department Manager
        -> Vehicle Allocator
        -> HR Manager
        -> Security
        -> Driver
        -> Security
        -> Driver
        -> Security
        -> Employee

All workflow emails are handled through the centralized notification
functions below. Individual workflow stages do not need separate SMTP code.
"""

import streamlit as st
import gspread

from google.oauth2.service_account import Credentials

from datetime import datetime, date, time, timedelta
from zoneinfo import ZoneInfo

import pandas as pd
import uuid
import time as time_module

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


# ============================================================
# STREAMLIT CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Vehicle Gate Pass",
    page_icon="🚐",
    layout="wide",
)


# ============================================================
# CONFIGURATION
# ============================================================

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SHEET_NAME = "vehicle_gate_pass_data"

TIMEZONE = "Asia/Colombo"

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

READ_CACHE_TTL = 60


# ============================================================
# SHEET HEADERS
# ============================================================

SHEET_HEADERS = {
    "Users": [
        "username",
        "name",
        "role",
        "password",
        "active",
        "email",
    ],

    "Departments": [
        "department",
        "manager_username",
        "manager_name",
        "manager_password",
        "active",
    ],

    "Drivers": [
        "driver_username",
        "driver_name",
        "password",
        "active",
    ],

    "Vehicles": [
        "vehicle_number",
        "vehicle_type",
        "driver_username",
        "driver_name",
        "active",
    ],

    "GatePasses": [
        "request_id",
        "created_at",
        "requisitioner_name",
        "department",
        "manager_username",
        "companions",
        "duration_minutes",
        "destination",
        "purpose",
        "travel_date",
        "start_time",
        "end_time",
        "vehicle_number",
        "driver_username",
        "driver_name",
        "status",
        "manager_decision",
        "manager_approved_by",
        "manager_approved_at",
        "hr_decision",
        "hr_approved_by",
        "hr_approved_at",
        "security_released_by",
        "security_released_at",
        "start_mileage",
        "driver_started_at",
        "end_mileage",
        "distance_km",
        "driver_completed_at",
        "security_verified_by",
        "security_verified_at",
        "rejection_reason",
    ],

    "ApprovalAudit": [
        "timestamp",
        "request_id",
        "username",
        "role",
        "action",
        "remarks",
    ],
}


# ============================================================
# STATUS / ROLES
# ============================================================

ACTIVE_STATUSES = {
    "Pending Department Manager",
    "Pending Vehicle Allocation",
    "Pending HR",
    "Pending Security",
    "Vehicle Released",
    "Pending Security Start Mileage Verification",
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

    missing = [
        key for key in required
        if key not in st.secrets
    ]

    if missing:
        raise RuntimeError(
            "Missing Google service-account secrets: "
            + ", ".join(missing)
        )

    info = {
        key: st.secrets[key]
        for key in required
    }

    info["client_id"] = str(info["client_id"])

    if "universe_domain" in st.secrets:
        info["universe_domain"] = st.secrets["universe_domain"]

    credentials = Credentials.from_service_account_info(
        info,
        scopes=SCOPES,
    )

    return gspread.authorize(credentials)


@st.cache_resource
def get_spreadsheet():
    client = get_google_client()

    sheet_id = str(
        st.secrets.get("google_sheet_id", "")
    ).strip()

    if sheet_id:
        return client.open_by_key(sheet_id)

    sheet_name = (
        str(
            st.secrets.get(
                "google_sheet_name",
                SHEET_NAME,
            )
        ).strip()
        or SHEET_NAME
    )

    return client.open(sheet_name)


def _api_call_with_backoff(
    operation,
    max_attempts=5,
):
    for attempt in range(max_attempts):
        try:
            return operation()

        except gspread.exceptions.APIError as exc:
            status = getattr(
                getattr(exc, "response", None),
                "status_code",
                None,
            )

            if (
                status != 429
                or attempt == max_attempts - 1
            ):
                raise

            time_module.sleep(
                min(2 ** attempt, 16)
            )

    return None


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
            cols=max(
                20,
                len(headers) + 2,
            ),
        )

        worksheet.append_row(
            headers,
            value_input_option="USER_ENTERED",
        )

        return worksheet

    try:
        current_headers = worksheet.row_values(1)

    except Exception:
        current_headers = []

    if not current_headers:
        worksheet.append_row(
            SHEET_HEADERS[name],
            value_input_option="USER_ENTERED",
        )

    return worksheet


@st.cache_data(
    ttl=READ_CACHE_TTL,
    show_spinner=False,
)
def read_sheet(name):
    worksheet = get_or_create_worksheet(name)

    records = _api_call_with_backoff(
        worksheet.get_all_records
    )

    return pd.DataFrame(records)


def invalidate_data_cache():
    read_sheet.clear()


def append_row(name, row_dict):
    worksheet = get_or_create_worksheet(name)

    headers = SHEET_HEADERS[name]

    values = [
        row_dict.get(header, "")
        for header in headers
    ]

    _api_call_with_backoff(
        lambda: worksheet.append_row(
            values,
            value_input_option="USER_ENTERED",
        )
    )

    invalidate_data_cache()


def update_request(
    request_id,
    updates,
):
    worksheet = get_or_create_worksheet(
        "GatePasses"
    )

    headers = SHEET_HEADERS["GatePasses"]

    records = read_sheet("GatePasses")

    target_row = None

    if (
        not records.empty
        and "request_id" in records.columns
    ):
        matches = records.index[
            records["request_id"]
            .astype(str)
            .str.strip()
            == str(request_id).strip()
        ].tolist()

        if matches:
            target_row = matches[0] + 2

    if target_row is None:
        read_sheet.clear()

        raw_records = _api_call_with_backoff(
            worksheet.get_all_records
        )

        for row_number, record in enumerate(
            raw_records,
            start=2,
        ):
            if (
                str(
                    record.get(
                        "request_id",
                        "",
                    )
                ).strip()
                == str(request_id).strip()
            ):
                target_row = row_number
                break

    if target_row is None:
        raise ValueError(
            f"Request {request_id} was not found."
        )

    payload = []

    for key, value in updates.items():

        if key not in headers:
            continue

        column_number = (
            headers.index(key) + 1
        )

        payload.append(
            {
                "range": gspread.utils.rowcol_to_a1(
                    target_row,
                    column_number,
                ),
                "values": [
                    [
                        ""
                        if value is None
                        else str(value)
                    ]
                ],
            }
        )

    if payload:
        _api_call_with_backoff(
            lambda: worksheet.batch_update(
                payload,
                raw=False,
            )
        )

    invalidate_data_cache()


# ============================================================
# GENERAL HELPERS
# ============================================================

def now_str():
    return datetime.now(
        ZoneInfo(TIMEZONE)
    ).strftime(
        "%Y-%m-%d %H:%M:%S"
    )


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

    for fmt in (
        "%Y-%m-%d",
        "%d/%m/%Y",
        "%m/%d/%Y",
    ):
        try:
            return datetime.strptime(
                text,
                fmt,
            ).date()
        except ValueError:
            pass

    return None


def parse_time(value):
    if isinstance(value, time):
        return value

    text = str(value).strip()

    for fmt in (
        "%H:%M",
        "%H:%M:%S",
        "%I:%M %p",
    ):
        try:
            return datetime.strptime(
                text,
                fmt,
            ).time()
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
    if (
        dataframe.empty
        or "active" not in dataframe.columns
    ):
        return dataframe

    return dataframe[
        dataframe["active"].apply(
            normalize_bool
        )
    ].copy()


def get_users():
    return active_records(
        read_sheet("Users")
    )


def get_departments():
    return active_records(
        read_sheet("Departments")
    )


def get_drivers():
    return active_records(
        read_sheet("Drivers")
    )


def get_vehicles():
    return active_records(
        read_sheet("Vehicles")
    )


def get_gatepasses():
    return read_sheet("GatePasses")


# ============================================================
# AUDIT
# ============================================================

def audit(
    request_id,
    username,
    role,
    action,
    remarks="",
):
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
# EMAIL SYSTEM
# ============================================================

def email_enabled():
    sender = str(
        st.secrets.get(
            "gmail_sender",
            "",
        )
    ).strip()

    password = str(
        st.secrets.get(
            "gmail_password",
            "",
        )
    ).strip()

    return bool(
        sender and password
    )


def send_email(
    subject,
    body,
    recipient,
):
    """
    Central email sender.

    Returns:
        True  = email sent
        False = email failed
    """

    recipient = str(
        recipient or ""
    ).strip()

    if not recipient:
        return False

    sender = str(
        st.secrets.get(
            "gmail_sender",
            "",
        )
    ).strip()

    password = str(
        st.secrets.get(
            "gmail_password",
            "",
        )
    ).strip()

    if not sender or not password:
        print(
            "Email not configured. "
            "Missing gmail_sender or gmail_password."
        )
        return False

    try:
        message = MIMEMultipart()

        message["From"] = sender
        message["To"] = recipient
        message["Subject"] = subject

        message.attach(
            MIMEText(
                body,
                "plain",
            )
        )

        with smtplib.SMTP(
            SMTP_SERVER,
            SMTP_PORT,
        ) as server:

            server.starttls()

            server.login(
                sender,
                password,
            )

            server.send_message(
                message
            )

        return True

    except Exception as error:
        print(
            f"Email Error: {error}"
        )
        return False


def get_user_email_by_username(username):
    """
    Find an active user's email from Users sheet.
    """

    username = str(
        username or ""
    ).strip()

    if not username:
        return None

    users = get_users()

    if users.empty:
        return None

    matches = users[
        users["username"]
        .astype(str)
        .str.strip()
        == username
    ]

    if matches.empty:
        return None

    email = str(
        matches.iloc[0].get(
            "email",
            "",
        )
    ).strip()

    return email or None


def get_user_email_by_name(name):
    """
    Find an active user's email using their name.

    This is useful for the employee/requisitioner because the
    current employee portal does not require an employee login.
    """

    name = str(
        name or ""
    ).strip().lower()

    if not name:
        return None

    users = get_users()

    if users.empty:
        return None

    if "name" not in users.columns:
        return None

    matches = users[
        users["name"]
        .astype(str)
        .str.strip()
        .str.lower()
        == name
    ]

    if matches.empty:
        return None

    email = str(
        matches.iloc[0].get(
            "email",
            "",
        )
    ).strip()

    return email or None


def get_emails_by_role(role):
    """
    Return all active users having the specified role.
    """

    role = str(
        role or ""
    ).strip().lower()

    users = get_users()

    if users.empty:
        return []

    matches = users[
        users["role"]
        .astype(str)
        .str.strip()
        .str.lower()
        == role
    ]

    emails = []

    for _, row in matches.iterrows():
        email = str(
            row.get(
                "email",
                "",
            )
        ).strip()

        if email and email not in emails:
            emails.append(email)

    return emails


def get_allocator_emails():
    return get_emails_by_role(
        "Vehicle Allocator"
    )


def get_hr_emails():
    return get_emails_by_role(
        "HR Manager"
    )


def get_security_emails():
    return get_emails_by_role(
        "Security"
    )


def get_driver_email(driver_username):
    return get_user_email_by_username(
        driver_username
    )


def get_employee_email(row):
    """
    Try to identify the requisitioner's email.

    First try username if available.
    Then try matching requisitioner name.
    """

    username = str(
        row.get(
            "requisitioner_username",
            "",
        )
    ).strip()

    if username:
        email = get_user_email_by_username(
            username
        )

        if email:
            return email

    return get_user_email_by_name(
        row.get(
            "requisitioner_name",
            "",
        )
    )


def get_department_manager_email(
    department,
    manager_username=None,
):
    """
    Find the manager's email.

    Manager credentials are stored in Departments,
    while manager email is stored in Users.
    """

    if manager_username:
        email = get_user_email_by_username(
            manager_username
        )

        if email:
            return email

    departments = get_departments()

    if departments.empty:
        return None

    department = str(
        department or ""
    ).strip().lower()

    matches = departments[
        departments["department"]
        .astype(str)
        .str.strip()
        .str.lower()
        == department
    ]

    if matches.empty:
        return None

    manager_username = str(
        matches.iloc[0].get(
            "manager_username",
            "",
        )
    ).strip()

    if not manager_username:
        return None

    return get_user_email_by_username(
        manager_username
    )


def request_email_details(row):
    """
    Build common information used by workflow emails.
    """

    return f"""
Request ID: {row.get('request_id', '')}
Employee: {row.get('requisitioner_name', '')}
Department: {row.get('department', '')}
Destination: {row.get('destination', '')}
Purpose: {row.get('purpose', '')}
Passengers: {row.get('companions', '')}

Travel Date: {row.get('travel_date', '')}
Departure Time: {row.get('start_time', '')}
Return Time: {row.get('end_time', '')}

Vehicle: {row.get('vehicle_number', '') or 'Not allocated'}
Driver: {row.get('driver_name', '') or 'Not assigned'}
""".strip()


def notify_manager_new_request(row):
    """
    Employee -> Department Manager.
    """

    recipient = get_department_manager_email(
        row.get("department"),
        row.get("manager_username"),
    )

    if not recipient:
        return False

    subject = (
        "Vehicle Gate Pass Approval Required - "
        f"{row.get('request_id', '')}"
    )

    body = f"""
A new vehicle gate pass request is waiting for your approval.

{request_email_details(row)}

Current Status:
Pending Department Manager

Please log in to the Vehicle Gate Pass Management System
to approve or reject this request.
"""

    return send_email(
        subject,
        body,
        recipient,
    )


def notify_vehicle_allocator(row):
    """
    Department Manager -> Vehicle Allocator.
    """

    recipients = get_allocator_emails()

    if not recipients:
        return False

    subject = (
        "Vehicle Allocation Required - "
        f"{row.get('request_id', '')}"
    )

    body = f"""
The Department Manager has approved a vehicle request.

{request_email_details(row)}

Current Status:
Pending Vehicle Allocation

Please log in to the Vehicle Gate Pass Management System
and allocate an available vehicle and its fixed driver.
"""

    success = False

    for recipient in recipients:
        if send_email(
            subject,
            body,
            recipient,
        ):
            success = True

    return success


def notify_hr(row):
    """
    Vehicle Allocator -> HR Manager.
    """

    recipients = get_hr_emails()

    if not recipients:
        return False

    subject = (
        "HR Approval Required - "
        f"{row.get('request_id', '')}"
    )

    body = f"""
A vehicle has been allocated and the request is now
waiting for HR approval.

{request_email_details(row)}

Current Status:
Pending HR

Please log in to the Vehicle Gate Pass Management System
and approve or reject this vehicle gate pass.
"""

    success = False

    for recipient in recipients:
        if send_email(
            subject,
            body,
            recipient,
        ):
            success = True

    return success


def notify_security(row):
    """
    HR -> Security.
    """

    recipients = get_security_emails()

    if not recipients:
        return False

    subject = (
        "Vehicle Gate Pass Approved - Security Action Required - "
        f"{row.get('request_id', '')}"
    )

    body = f"""
HR has approved the vehicle gate pass.

{request_email_details(row)}

HR Approved By:
{row.get('hr_approved_by', '')}

HR Approved At:
{row.get('hr_approved_at', '')}

Current Status:
Pending Security

Please log in to the Vehicle Gate Pass Management System
and verify/release the vehicle when appropriate.
"""

    success = False

    for recipient in recipients:
        if send_email(
            subject,
            body,
            recipient,
        ):
            success = True

    return success


def notify_driver_vehicle_released(row):
    """
    Security -> Driver.
    """

    driver_username = str(
        row.get(
            "driver_username",
            "",
        )
    ).strip()

    recipient = get_driver_email(
        driver_username
    )

    if not recipient:
        return False

    subject = (
        "Vehicle Released - Trip Assigned - "
        f"{row.get('request_id', '')}"
    )

    body = f"""
Security has released the vehicle for your assigned trip.

{request_email_details(row)}

Current Status:
Vehicle Released

Please log in to the Driver Portal when you are
ready to enter the starting mileage.
"""

    return send_email(
        subject,
        body,
        recipient,
    )


def notify_security_start_mileage(row):
    """
    Driver -> Security.
    """

    recipients = get_security_emails()

    if not recipients:
        return False

    subject = (
        "Starting Mileage Verification Required - "
        f"{row.get('request_id', '')}"
    )

    body = f"""
The assigned driver has submitted the starting mileage.

{request_email_details(row)}

Starting Mileage:
{row.get('start_mileage', '')} km

Driver:
{row.get('driver_name', '')}

Current Status:
Pending Security Start Mileage Verification

Please log in to the system and verify the starting mileage.
"""

    success = False

    for recipient in recipients:
        if send_email(
            subject,
            body,
            recipient,
        ):
            success = True

    return success


def notify_driver_trip_started(row):
    """
    Security -> Driver.
    """

    driver_username = str(
        row.get(
            "driver_username",
            "",
        )
    ).strip()

    recipient = get_driver_email(
        driver_username
    )

    if not recipient:
        return False

    subject = (
        "Starting Mileage Verified - Trip May Start - "
        f"{row.get('request_id', '')}"
    )

    body = f"""
Security has verified the starting mileage.

{request_email_details(row)}

Starting Mileage:
{row.get('start_mileage', '')} km

Current Status:
Trip In Progress

You may now proceed with the trip.
"""

    return send_email(
        subject,
        body,
        recipient,
    )


def notify_security_trip_completed(row):
    """
    Driver -> Security.
    """

    recipients = get_security_emails()

    if not recipients:
        return False

    subject = (
        "Trip Completed - Security Verification Required - "
        f"{row.get('request_id', '')}"
    )

    body = f"""
The driver has completed the trip.

{request_email_details(row)}

Starting Mileage:
{row.get('start_mileage', '')} km

Ending Mileage:
{row.get('end_mileage', '')} km

Distance:
{row.get('distance_km', '')} km

Current Status:
Pending Security Verification

Please log in to the system and verify/close the trip.
"""

    success = False

    for recipient in recipients:
        if send_email(
            subject,
            body,
            recipient,
        ):
            success = True

    return success


def notify_employee_completed(row):
    """
    Security -> Employee.

    Since the employee portal currently does not require login,
    this attempts to find the employee by name in Users.
    """

    recipient = get_employee_email(
        row
    )

    if not recipient:
        return False

    subject = (
        "Vehicle Gate Pass Completed - "
        f"{row.get('request_id', '')}"
    )

    body = f"""
Your vehicle gate pass process has been completed.

{request_email_details(row)}

Starting Mileage:
{row.get('start_mileage', '')} km

Ending Mileage:
{row.get('end_mileage', '')} km

Distance:
{row.get('distance_km', '')} km

Current Status:
Completed

Thank you.
"""

    return send_email(
        subject,
        body,
        recipient,
    )


def notify_employee_rejected(
    row,
    rejected_by,
    reason,
):
    """
    Manager / HR -> Employee.
    """

    recipient = get_employee_email(
        row
    )

    if not recipient:
        return False

    subject = (
        "Vehicle Gate Pass Rejected - "
        f"{row.get('request_id', '')}"
    )

    body = f"""
Your vehicle gate pass request has been rejected.

{request_email_details(row)}

Rejected By:
{rejected_by}

Reason:
{reason or 'No reason provided.'}

Please contact the relevant manager or HR department
if you need further information.
"""

    return send_email(
        subject,
        body,
        recipient,
    )


# ============================================================
# AUTHENTICATION
# ============================================================

def authenticate_user(
    username,
    password,
    role,
):
    users = get_users()

    if users.empty:
        return None

    username = str(
        username
    ).strip()

    password = str(
        password
    ).strip()

    role = str(
        role
    ).strip().lower()

    for _, row in users.iterrows():

        if (
            str(
                row.get(
                    "username",
                    "",
                )
            ).strip()
            == username

            and str(
                row.get(
                    "password",
                    "",
                )
            ).strip()
            == password

            and str(
                row.get(
                    "role",
                    "",
                )
            ).strip().lower()
            == role

            and normalize_bool(
                row.get(
                    "active",
                    "Yes",
                )
            )
        ):
            return row.to_dict()

    return None


def authenticate_manager(
    username,
    password,
):
    departments = get_departments()

    if departments.empty:
        return None

    username = str(
        username
    ).strip()

    password = str(
        password
    ).strip()

    for _, row in departments.iterrows():

        if (
            str(
                row.get(
                    "manager_username",
                    "",
                )
            ).strip()
            == username

            and str(
                row.get(
                    "manager_password",
                    "",
                )
            ).strip()
            == password

            and normalize_bool(
                row.get(
                    "active",
                    "Yes",
                )
            )
        ):
            return row.to_dict()

    return None


def authenticate_driver(
    username,
    password,
):
    drivers = get_drivers()

    if drivers.empty:
        return None

    username = str(
        username
    ).strip()

    password = str(
        password
    ).strip()

    for _, row in drivers.iterrows():

        if (
            str(
                row.get(
                    "driver_username",
                    "",
                )
            ).strip()
            == username

            and str(
                row.get(
                    "password",
                    "",
                )
            ).strip()
            == password

            and normalize_bool(
                row.get(
                    "active",
                    "Yes",
                )
            )
        ):
            return row.to_dict()

    return None


# ============================================================
# VEHICLE / DRIVER
# ============================================================

def vehicle_driver(
    vehicle_number,
):
    vehicles = get_vehicles()
    drivers = get_drivers()

    vehicle_number = str(
        vehicle_number
    ).strip()

    for _, vehicle in vehicles.iterrows():

        if (
            str(
                vehicle.get(
                    "vehicle_number",
                    "",
                )
            ).strip()
            == vehicle_number
        ):

            driver_username = str(
                vehicle.get(
                    "driver_username",
                    "",
                )
            ).strip()

            driver_name = str(
                vehicle.get(
                    "driver_name",
                    "",
                )
            ).strip()

            if (
                driver_username
                and not driver_name
            ):
                for _, driver in drivers.iterrows():

                    if (
                        str(
                            driver.get(
                                "driver_username",
                                "",
                            )
                        ).strip()
                        == driver_username
                    ):
                        driver_name = str(
                            driver.get(
                                "driver_name",
                                "",
                            )
                        ).strip()
                        break

            return {
                "driver_username": driver_username,
                "driver_name": driver_name,
            }

    return None


def request_interval(row):
    travel_date = parse_date(
        row.get("travel_date")
    )

    start_time = parse_time(
        row.get("start_time")
    )

    end_time = parse_time(
        row.get("end_time")
    )

    if (
        not travel_date
        or not start_time
        or not end_time
    ):
        return None, None

    return (
        datetime.combine(
            travel_date,
            start_time,
        ),
        datetime.combine(
            travel_date,
            end_time,
        ),
    )


def vehicle_is_available(
    vehicle_number,
    travel_date,
    start_dt,
    end_dt,
    exclude_request_id=None,
):
    gatepasses = get_gatepasses()

    if gatepasses.empty:
        return True

    for _, row in gatepasses.iterrows():

        if (
            str(
                row.get(
                    "vehicle_number",
                    "",
                )
            ).strip()
            != str(
                vehicle_number
            ).strip()
        ):
            continue

        if (
            exclude_request_id
            and str(
                row.get(
                    "request_id",
                    "",
                )
            ).strip()
            == str(
                exclude_request_id
            ).strip()
        ):
            continue

        if (
            str(
                row.get(
                    "status",
                    "",
                )
            ).strip()
            not in ACTIVE_STATUSES
        ):
            continue

        row_date = parse_date(
            row.get("travel_date")
        )

        if row_date != travel_date:
            continue

        old_start, old_end = request_interval(
            row
        )

        if (
            old_start is None
            or old_end is None
        ):
            continue

        if (
            start_dt < old_end
            and end_dt > old_start
        ):
            return False

    return True


def available_vehicles_for_interval(
    travel_date,
    start_time,
    duration_minutes,
):
    start_dt = datetime.combine(
        travel_date,
        start_time,
    )

    end_dt = (
        start_dt
        + timedelta(
            minutes=duration_minutes
        )
    )

    if end_dt.date() != travel_date:
        return []

    result = []

    for _, vehicle in get_vehicles().iterrows():

        vehicle_number = str(
            vehicle.get(
                "vehicle_number",
                "",
            )
        ).strip()

        if not vehicle_number:
            continue

        if vehicle_is_available(
            vehicle_number,
            travel_date,
            start_dt,
            end_dt,
        ):
            driver_info = (
                vehicle_driver(
                    vehicle_number
                )
                or {}
            )

            result.append(
                {
                    "vehicle_number":
                        vehicle_number,

                    "vehicle_type":
                        str(
                            vehicle.get(
                                "vehicle_type",
                                "",
                            )
                        ).strip(),

                    "driver_username":
                        driver_info.get(
                            "driver_username",
                            str(
                                vehicle.get(
                                    "driver_username",
                                    "",
                                )
                            ).strip(),
                        ),

                    "driver_name":
                        driver_info.get(
                            "driver_name",
                            str(
                                vehicle.get(
                                    "driver_name",
                                    "",
                                )
                            ).strip(),
                        ),
                }
            )

    return result


# ============================================================
# SESSION / LOGIN
# ============================================================

def logout():
    for key in list(
        st.session_state.keys()
    ):
        del st.session_state[key]

    st.rerun()


def login_portal(role):
    st.subheader(
        f"{role} Login"
    )

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
# EMPLOYEE
# ============================================================

def _store_pending_request(
    draft,
):
    st.session_state[
        "pending_vehicle_request"
    ] = draft

    st.session_state[
        "transfer_reminder_snooze_until"
    ] = None


def _clear_pending_request():
    st.session_state.pop(
        "pending_vehicle_request",
        None,
    )

    st.session_state.pop(
        "transfer_reminder_snooze_until",
        None,
    )


@st.dialog(
    "⚠️ Vehicle Request Not Yet Transferred",
    width="medium",
)
def transfer_reminder_dialog():

    draft = st.session_state.get(
        "pending_vehicle_request"
    )

    if not draft:
        return

    st.warning(
        "You have completed a vehicle request, "
        "but it has NOT been transferred to Google Sheets yet."
    )

    st.write(
        f"**Request:** {draft['request_id']}\n\n"
        f"**Date:** {draft['travel_date']}\n\n"
        f"**Required Time:** "
        f"{draft['start_time']} - "
        f"{draft['end_time']}\n\n"
        f"**Duration:** "
        f"{draft['duration_minutes']} minutes"
    )

    st.info(
        "The request is currently stored only in this browser "
        "session. Managers and HR will not see it until you transfer it."
    )

    c1, c2 = st.columns(2)

    with c1:

        if st.button(
            "Transfer Data Now",
            type="primary",
            use_container_width=True,
        ):
            st.session_state[
                "transfer_from_reminder"
            ] = True

            st.rerun()

    with c2:

        if st.button(
            "Continue Editing",
            use_container_width=True,
        ):
            st.session_state[
                "transfer_reminder_snooze_until"
            ] = (
                datetime.now()
                + timedelta(seconds=30)
            ).isoformat()

            st.rerun()


def maybe_show_transfer_reminder():

    draft = st.session_state.get(
        "pending_vehicle_request"
    )

    if not draft:
        return

    snooze = st.session_state.get(
        "transfer_reminder_snooze_until"
    )

    if snooze:
        try:
            if (
                datetime.now()
                < datetime.fromisoformat(
                    snooze
                )
            ):
                return
        except ValueError:
            pass

    transfer_reminder_dialog()


def _transfer_pending_request():

    draft = st.session_state.get(
        "pending_vehicle_request"
    )

    if not draft:
        return False

    if st.session_state.get(
        "_vehicle_transfer_in_progress",
        False,
    ):
        st.warning(
            "The request is already being transferred. Please wait."
        )
        return False

    st.session_state[
        "_vehicle_transfer_in_progress"
    ] = True

    try:

        read_sheet.clear()

        append_row(
            "GatePasses",
            draft["row"],
        )

        audit(
            draft["request_id"],
            "employee",
            "Employee / Requisitioner",
            "Request Submitted",
            (
                "Vehicle request submitted. "
                "Waiting for Department Manager approval."
            ),
        )

        # ----------------------------------------------------
        # EMAIL DEPARTMENT MANAGER
        # ----------------------------------------------------

        email_sent = notify_manager_new_request(
            draft["row"]
        )

        if not email_sent:
            print(
                "Warning: Department Manager email "
                "could not be sent."
            )

        _clear_pending_request()

        st.session_state.pop(
            "transfer_from_reminder",
            None,
        )

        st.session_state.pop(
            "transfer_reminder_snooze_until",
            None,
        )

        return True

    finally:
        st.session_state[
            "_vehicle_transfer_in_progress"
        ] = False


def requisitioner_status_checker():

    st.subheader(
        "🔎 Check Gate Pass Status"
    )

    st.write(
        "Enter the Request ID you received after submitting your vehicle request."
    )

    request_id = st.text_input(
        "Request ID",
        placeholder="Example: VGP-20260820-ABC123",
        key="requisitioner_status_request_id",
    ).strip()

    if st.button(
        "Check Status",
        type="primary",
        key="check_requisitioner_status",
    ):

        if not request_id:
            st.warning(
                "Please enter your Request ID."
            )
            return

        dataframe = get_gatepasses()

        if dataframe.empty:
            st.error(
                "No vehicle requests are currently available."
            )
            return

        matches = dataframe[
            dataframe["request_id"]
            .astype(str)
            .str.strip()
            .str.upper()
            == request_id.upper()
        ]

        if matches.empty:
            st.error(
                "Request ID not found. "
                "Please check the Request ID and try again."
            )
            return

        row = matches.iloc[0]

        status = str(
            row.get(
                "status",
                "",
            )
        ).strip()

        st.success(
            f"Request found: **{request_id}**"
        )

        st.markdown(
            f"### Current Status: **{status}**"
        )

        col1, col2 = st.columns(2)

        with col1:

            st.write(
                f"**Requisitioner:** "
                f"{row.get('requisitioner_name', '')}"
            )

            st.write(
                f"**Department:** "
                f"{row.get('department', '')}"
            )

            st.write(
                f"**Travel Date:** "
                f"{row.get('travel_date', '')}"
            )

            st.write(
                f"**Destination:** "
                f"{row.get('destination', '')}"
            )

        with col2:

            st.write(
                f"**Vehicle:** "
                f"{row.get('vehicle_number', '') or 'Not allocated yet'}"
            )

            st.write(
                f"**Driver:** "
                f"{row.get('driver_name', '') or 'Not allocated yet'}"
            )

            st.write(
                f"**Vehicle Time:** "
                f"{row.get('start_time', '')} - "
                f"{row.get('end_time', '')}"
            )

        st.divider()

        st.subheader(
            "Approval Progress"
        )

        manager_decision = str(
            row.get(
                "manager_decision",
                "",
            )
        ).strip()

        hr_decision = str(
            row.get(
                "hr_decision",
                "",
            )
        ).strip()

        if manager_decision == "Approved":
            st.success(
                "✅ Department Manager: Approved"
            )

        elif manager_decision == "Rejected":
            st.error(
                "❌ Department Manager: Rejected"
            )

        else:
            st.info(
                "⏳ Department Manager: Pending"
            )

        if row.get(
            "vehicle_number",
            "",
        ):
            st.success(
                "✅ Vehicle Allocation: Completed"
            )
        else:
            st.info(
                "⏳ Vehicle Allocation: Pending"
            )

        if hr_decision == "Approved":
            st.success(
                "✅ HR Manager: Approved"
            )

        elif hr_decision == "Rejected":
            st.error(
                "❌ HR Manager: Rejected"
            )

        else:
            st.info(
                "⏳ HR Manager: Pending"
            )

        if status in [
            "Pending Security",
            "Vehicle Released",
            "Trip In Progress",
            "Pending Security Verification",
            "Completed",
        ]:
            st.success(
                "✅ Security Stage: Reached"
            )

        if status == "Completed":
            st.success(
                "🎉 Gate Pass process completed."
            )

        rejection_reason = str(
            row.get(
                "rejection_reason",
                "",
            )
        ).strip()

        if rejection_reason:
            st.error(
                f"**Reason:** {rejection_reason}"
            )


def employee_portal():

    st.header(
        "🚐 Vehicle Requisition"
    )

    st.caption(
        "Enter the date and exact time you need the vehicle. "
        "The Vehicle Allocator will select an available vehicle."
    )

    st.divider()

    requisitioner_status_checker()

    st.divider()

    st.subheader(
        "📝 Submit a New Vehicle Request"
    )

    if st.session_state.get(
        "transfer_from_reminder"
    ):

        try:

            with st.spinner(
                "Transferring request..."
            ):

                if _transfer_pending_request():

                    st.success(
                        "The request has been transferred to Google Sheets successfully."
                    )

                    st.rerun()

        except Exception as error:

            st.error(
                "The request could not be transferred to Google Sheets."
            )

            st.exception(error)

    maybe_show_transfer_reminder()

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
            str(v).strip()
            for v in departments["department"].tolist()
            if str(v).strip()
        ]
    )

    with st.form(
        "vehicle_request_form"
    ):

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

            requested_start_time = st.time_input(
                "Vehicle Required From *",
                value=time(8, 0),
                step=timedelta(
                    minutes=30
                ),
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
                list(
                    duration_options.keys()
                ),
            )

            duration_minutes = (
                duration_options[
                    duration_label
                ]
            )

            requested_start_dt = datetime.combine(
                travel_date,
                requested_start_time,
            )

            requested_end_dt = (
                requested_start_dt
                + timedelta(
                    minutes=duration_minutes
                )
            )

            if (
                requested_end_dt.date()
                == travel_date
            ):

                st.info(
                    f"Vehicle required from "
                    f"**{requested_start_dt.strftime('%H:%M')}** "
                    f"to "
                    f"**{requested_end_dt.strftime('%H:%M')}**."
                )

            else:

                st.error(
                    "The requested duration extends into the next day."
                )

        st.divider()

        st.info(
            "Your request will first go to the Department Manager. "
            "After approval it will go to Vehicle Allocation, then HR, "
            "then Security."
        )

        prepared = st.form_submit_button(
            "Review Request & Prepare Transfer",
            type="primary",
        )

    if prepared:

        if (
            not requisitioner.strip()
            or not destination.strip()
            or not purpose.strip()
        ):
            st.error(
                "Please complete all required fields."
            )
            return

        if requested_end_dt.date() != travel_date:
            st.error(
                "The requested vehicle time cannot extend into the next day."
            )
            return

        manager = None

        for _, row in departments.iterrows():

            if (
                str(
                    row.get(
                        "department",
                        "",
                    )
                ).strip()
                == department
            ):
                manager = row.to_dict()
                break

        if not manager:
            st.error(
                "No manager is configured for this department."
            )
            return

        request_id = generate_request_id()

        row = {
            "request_id": request_id,
            "created_at": now_str(),
            "requisitioner_name": requisitioner.strip(),
            "department": department,

            "manager_username": str(
                manager.get(
                    "manager_username",
                    "",
                )
            ).strip(),

            "companions": companions.strip(),

            "duration_minutes": duration_minutes,

            "destination": destination.strip(),
            "purpose": purpose.strip(),

            "travel_date": travel_date.isoformat(),

            "start_time": requested_start_time.strftime(
                "%H:%M"
            ),

            "end_time": requested_end_dt.strftime(
                "%H:%M"
            ),

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

        _store_pending_request(
            {
                "request_id": request_id,
                "travel_date": travel_date.isoformat(),
                "start_time": requested_start_time.strftime(
                    "%H:%M"
                ),
                "end_time": requested_end_dt.strftime(
                    "%H:%M"
                ),
                "duration_minutes": duration_minutes,
                "row": row,
                "manager_name": manager.get(
                    "manager_name",
                    manager.get(
                        "manager_username",
                        "Manager",
                    ),
                ),
            }
        )

        st.rerun()

    draft = st.session_state.get(
        "pending_vehicle_request"
    )

    if not draft:
        return

    st.divider()

    st.subheader(
        "📋 Final Transfer Point"
    )

    st.warning(
        "Your request is ready, but it has not been sent to Google Sheets."
    )

    review = pd.DataFrame(
        [
            {
                "Request ID":
                    draft["request_id"],

                "Requisitioner":
                    draft["row"]["requisitioner_name"],

                "Department":
                    draft["row"]["department"],

                "Destination":
                    draft["row"]["destination"],

                "Purpose":
                    draft["row"]["purpose"],

                "Date":
                    draft["travel_date"],

                "Required Time":
                    f"{draft['start_time']} - "
                    f"{draft['end_time']}",

                "Duration":
                    f"{draft['duration_minutes']} minutes",

                "Vehicle":
                    "To be allocated",

                "Driver":
                    "To be allocated",
            }
        ]
    )

    st.dataframe(
        review,
        use_container_width=True,
        hide_index=True,
    )

    if st.button(
        "🚀 TRANSFER DATA TO GOOGLE SHEET",
        type="primary",
        use_container_width=True,
        key="transfer_pending_request",
    ):

        try:

            with st.spinner(
                "Transferring request..."
            ):

                if _transfer_pending_request():

                    st.success(
                        f"Request {draft['request_id']} "
                        "has been transferred successfully."
                    )

                    st.info(
                        "The Department Manager has been notified by email."
                    )

                    st.rerun()

        except Exception as error:

            st.error(
                "The request could not be transferred to Google Sheets."
            )

            st.exception(error)


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

    st.header(
        "👤 Department Manager"
    )

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
        user.get(
            "department",
            "",
        )
    ).strip()

    manager_username = str(
        user.get(
            "manager_username",
            username,
        )
    ).strip()

    pending = dataframe[
        (
            dataframe["status"]
            .astype(str)
            == "Pending Department Manager"
        )
        &
        (
            (
                dataframe[
                    "manager_username"
                ]
                .astype(str)
                .str.strip()
                == manager_username
            )
            |
            (
                dataframe[
                    "department"
                ]
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
            row.get(
                "request_id",
                "",
            )
        )

        with st.expander(
            f"{request_id} — "
            f"{row.get('requisitioner_name', '')} — "
            f"{row.get('travel_date', '')}"
        ):

            col1, col2 = st.columns(2)

            with col1:

                st.write(
                    f"**Department:** "
                    f"{row.get('department', '')}"
                )

                st.write(
                    f"**Destination:** "
                    f"{row.get('destination', '')}"
                )

                st.write(
                    f"**Purpose:** "
                    f"{row.get('purpose', '')}"
                )

                st.write(
                    f"**Passengers:** "
                    f"{row.get('companions', '')}"
                )

            with col2:

                st.write(
                    f"**Date:** "
                    f"{row.get('travel_date', '')}"
                )

                st.write(
                    f"**Requested Vehicle Time:** "
                    f"{row.get('start_time', '')} - "
                    f"{row.get('end_time', '')}"
                )

                st.write(
                    f"**Expected Duration:** "
                    f"{row.get('duration_minutes', '')} minutes"
                )

                st.info(
                    "The Vehicle Allocator will select an available vehicle."
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

                    approved_at = now_str()

                    update_request(
                        request_id,
                        {
                            "status":
                                "Pending Vehicle Allocation",

                            "manager_decision":
                                "Approved",

                            "manager_approved_by":
                                username,

                            "manager_approved_at":
                                approved_at,
                        },
                    )

                    audit(
                        request_id,
                        username,
                        "Department Manager",
                        "Manager Approved",
                        reason,
                    )

                    invalidate_data_cache()

                    updated_row = (
                        get_gatepasses()
                    )

                    matches = updated_row[
                        updated_row["request_id"]
                        .astype(str)
                        .str.strip()
                        == request_id
                    ]

                    if not matches.empty:
                        email_row = matches.iloc[0]
                    else:
                        email_row = row

                    email_sent = notify_vehicle_allocator(
                        email_row
                    )

                    if email_sent:
                        st.success(
                            "Request approved. "
                            "Vehicle Allocator has been notified by email."
                        )
                    else:
                        st.warning(
                            "Request approved, but the Vehicle Allocator "
                            "email could not be sent."
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
                            "status":
                                "Rejected by Department Manager",

                            "manager_decision":
                                "Rejected",

                            "manager_approved_by":
                                username,

                            "manager_approved_at":
                                now_str(),

                            "rejection_reason":
                                reason,
                        },
                    )

                    audit(
                        request_id,
                        username,
                        "Department Manager",
                        "Manager Rejected",
                        reason,
                    )

                    employee_email_sent = (
                        notify_employee_rejected(
                            row,
                            username,
                            reason,
                        )
                    )

                    if employee_email_sent:
                        st.warning(
                            "Request rejected. "
                            "Employee has been notified by email."
                        )
                    else:
                        st.warning(
                            "Request rejected. "
                            "Employee email could not be found."
                        )

                    st.rerun()


# ============================================================
# VEHICLE ALLOCATOR
# ============================================================

def vehicle_allocator_portal():

    username = st.session_state.get(
        "username",
        "",
    )

    user = st.session_state.get(
        "user",
        {},
    )

    st.header(
        "🚐 Vehicle Allocator"
    )

    st.caption(
        f"Vehicle Allocator: "
        f"{user.get('name', username)}"
    )

    dataframe = get_gatepasses()

    if dataframe.empty:
        st.info(
            "No vehicle requests found."
        )
        return

    pending = dataframe[
        dataframe["status"]
        .astype(str)
        == "Pending Vehicle Allocation"
    ]

    if pending.empty:
        st.success(
            "No requests are awaiting vehicle allocation."
        )
        return

    for _, row in pending.iterrows():

        request_id = str(
            row.get(
                "request_id",
                "",
            )
        )

        travel_date = parse_date(
            row.get("travel_date")
        )

        try:
            duration_minutes = int(
                float(
                    row.get(
                        "duration_minutes",
                        0,
                    )
                )
            )
        except (
            ValueError,
            TypeError,
        ):
            duration_minutes = 0

        requested_start = parse_time(
            row.get("start_time")
        )

        requested_end = parse_time(
            row.get("end_time")
        )

        with st.expander(
            f"{request_id} — "
            f"{row.get('requisitioner_name', '')} — "
            f"{row.get('travel_date', '')}"
        ):

            col1, col2 = st.columns(2)

            with col1:

                st.write(
                    f"**Requisitioner:** "
                    f"{row.get('requisitioner_name', '')}"
                )

                st.write(
                    f"**Department:** "
                    f"{row.get('department', '')}"
                )

                st.write(
                    f"**Destination:** "
                    f"{row.get('destination', '')}"
                )

                st.write(
                    f"**Purpose:** "
                    f"{row.get('purpose', '')}"
                )

                st.write(
                    f"**Passengers:** "
                    f"{row.get('companions', '')}"
                )

            with col2:

                st.write(
                    f"**Travel Date:** "
                    f"{row.get('travel_date', '')}"
                )

                st.write(
                    f"**Requested Vehicle Time:** "
                    f"{row.get('start_time', '')} - "
                    f"{row.get('end_time', '')}"
                )

                st.write(
                    f"**Duration:** "
                    f"{duration_minutes} minutes"
                )

                st.success(
                    "Department Manager: Approved"
                )

            if (
                not travel_date
                or not requested_start
                or duration_minutes <= 0
            ):
                st.error(
                    "The travel date, start time, or duration is invalid."
                )
                continue

            requested_start_dt = datetime.combine(
                travel_date,
                requested_start,
            )

            calculated_end_dt = (
                requested_start_dt
                + timedelta(
                    minutes=duration_minutes
                )
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

            allocator_end_dt = (
                allocator_start_dt
                + timedelta(
                    minutes=duration_minutes
                )
            )

            st.info(
                f"**Confirmed Time:** "
                f"{allocator_start_dt.strftime('%H:%M')} - "
                f"{allocator_end_dt.strftime('%H:%M')}"
            )

            if (
                calculated_end_dt.date()
                != travel_date
            ):
                st.error(
                    "The requested trip extends into the next day."
                )
                continue

            if requested_end:

                stored_end_dt = datetime.combine(
                    travel_date,
                    requested_end,
                )

                if (
                    stored_end_dt
                    != calculated_end_dt
                ):
                    st.warning(
                        "The stored end time does not match "
                        "the requested duration."
                    )

            vehicles = (
                available_vehicles_for_interval(
                    travel_date,
                    allocator_start_time,
                    duration_minutes,
                )
            )

            if not vehicles:

                st.warning(
                    "No vehicles are currently available for "
                    f"{allocator_start_dt.strftime('%H:%M')} - "
                    f"{allocator_end_dt.strftime('%H:%M')}."
                )

                continue

            vehicle_labels = [
                (
                    f"{v['vehicle_number']} — "
                    f"{v['vehicle_type']} — "
                    f"Fixed Driver: "
                    f"{v['driver_name'] or v['driver_username']}"
                )
                for v in vehicles
            ]

            selected_label = st.selectbox(
                "Allocate Vehicle *",
                vehicle_labels,
                key=f"allocator_vehicle_{request_id}",
            )

            selected_vehicle = vehicles[
                vehicle_labels.index(
                    selected_label
                )
            ]

            selected_vehicle_number = (
                selected_vehicle[
                    "vehicle_number"
                ]
            )

            selected_driver = (
                vehicle_driver(
                    selected_vehicle_number
                )
                or {}
            )

            st.info(
                f"**Vehicle:** "
                f"{selected_vehicle_number}\n\n"
                f"**Fixed Driver:** "
                f"{selected_driver.get('driver_name', '') or selected_driver.get('driver_username', '')}\n\n"
                f"**Confirmed Time:** "
                f"{allocator_start_dt.strftime('%H:%M')} - "
                f"{allocator_end_dt.strftime('%H:%M')}"
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

                invalidate_data_cache()

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

                driver_info = (
                    vehicle_driver(
                        selected_vehicle_number
                    )
                    or {}
                )

                update_request(
                    request_id,
                    {
                        "status":
                            "Pending HR",

                        "start_time":
                            allocator_start_dt.strftime(
                                "%H:%M"
                            ),

                        "end_time":
                            allocator_end_dt.strftime(
                                "%H:%M"
                            ),

                        "vehicle_number":
                            selected_vehicle_number,

                        "driver_username":
                            driver_info.get(
                                "driver_username",
                                "",
                            ),

                        "driver_name":
                            driver_info.get(
                                "driver_name",
                                "",
                            ),
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
                        f"Time "
                        f"{allocator_start_dt.strftime('%H:%M')} - "
                        f"{allocator_end_dt.strftime('%H:%M')}; "
                        f"{remarks}"
                    ),
                )

                invalidate_data_cache()

                updated = get_gatepasses()

                matches = updated[
                    updated["request_id"]
                    .astype(str)
                    .str.strip()
                    == request_id
                ]

                if not matches.empty:
                    email_row = matches.iloc[0]
                else:
                    email_row = row

                email_sent = notify_hr(
                    email_row
                )

                if email_sent:
                    st.success(
                        "Vehicle allocated successfully. "
                        "HR has been notified by email."
                    )
                else:
                    st.warning(
                        "Vehicle allocated successfully, "
                        "but HR email could not be sent."
                    )

                st.rerun()


# ============================================================
# HR
# ============================================================

def hr_portal():

    username = st.session_state.get(
        "username",
        "",
    )

    st.header(
        "👩‍💼 HR Manager"
    )

    dataframe = get_gatepasses()

    if dataframe.empty:
        st.info(
            "No vehicle requests found."
        )
        return

    pending = dataframe[
        dataframe["status"]
        .astype(str)
        == "Pending HR"
    ]

    if pending.empty:
        st.success(
            "No pending HR approvals."
        )

    else:

        for _, row in pending.iterrows():

            request_id = str(
                row.get(
                    "request_id",
                    "",
                )
            )

            with st.expander(
                f"{request_id} — "
                f"{row.get('requisitioner_name', '')} — "
                f"{row.get('department', '')}"
            ):

                col1, col2 = st.columns(2)

                with col1:

                    st.write(
                        f"**Destination:** "
                        f"{row.get('destination', '')}"
                    )

                    st.write(
                        f"**Purpose:** "
                        f"{row.get('purpose', '')}"
                    )

                    st.write(
                        f"**Date:** "
                        f"{row.get('travel_date', '')}"
                    )

                    st.write(
                        f"**Time:** "
                        f"{row.get('start_time', '')} - "
                        f"{row.get('end_time', '')}"
                    )

                with col2:

                    st.write(
                        f"**Vehicle:** "
                        f"{row.get('vehicle_number', '')}"
                    )

                    st.write(
                        f"**Fixed Driver:** "
                        f"{row.get('driver_name', '')}"
                    )

                    st.success(
                        "Department Manager: Approved"
                    )

                    st.success(
                        "Vehicle Allocated"
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

                        approved_at = now_str()

                        update_request(
                            request_id,
                            {
                                "status":
                                    "Pending Security",

                                "hr_decision":
                                    "Approved",

                                "hr_approved_by":
                                    username,

                                "hr_approved_at":
                                    approved_at,
                            },
                        )

                        audit(
                            request_id,
                            username,
                            "HR Manager",
                            "HR Approved",
                            reason,
                        )

                        invalidate_data_cache()

                        updated = get_gatepasses()

                        matches = updated[
                            updated["request_id"]
                            .astype(str)
                            .str.strip()
                            == request_id
                        ]

                        if not matches.empty:
                            email_row = matches.iloc[0]
                        else:
                            email_row = row

                        email_sent = notify_security(
                            email_row
                        )

                        if email_sent:
                            st.success(
                                "Request approved by HR. "
                                "Security has been notified by email."
                            )
                        else:
                            st.warning(
                                "HR approval successful, "
                                "but Security email could not be sent."
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
                                "status":
                                    "Rejected by HR",

                                "hr_decision":
                                    "Rejected",

                                "hr_approved_by":
                                    username,

                                "hr_approved_at":
                                    now_str(),

                                "rejection_reason":
                                    reason,
                            },
                        )

                        audit(
                            request_id,
                            username,
                            "HR Manager",
                            "HR Rejected",
                            reason,
                        )

                        employee_email_sent = (
                            notify_employee_rejected(
                                row,
                                username,
                                reason,
                            )
                        )

                        if employee_email_sent:
                            st.warning(
                                "Request rejected. "
                                "Employee has been notified by email."
                            )
                        else:
                            st.warning(
                                "Request rejected. "
                                "Employee email could not be found."
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

    st.header(
        "🛡️ Security"
    )

    dataframe = get_gatepasses()

    if dataframe.empty:
        st.info(
            "No vehicle requests found."
        )
        return

    # --------------------------------------------------------
    # RELEASE VEHICLE
    # --------------------------------------------------------

    st.subheader(
        "Approved Requests Awaiting Vehicle Release"
    )

    pending_release = dataframe[
        dataframe["status"]
        .astype(str)
        == "Pending Security"
    ]

    if pending_release.empty:

        st.info(
            "No approved requests are awaiting release."
        )

    else:

        for _, row in pending_release.iterrows():

            request_id = str(
                row.get(
                    "request_id",
                    "",
                )
            )

            with st.expander(
                f"{request_id} — "
                f"{row.get('requisitioner_name', '')} — "
                f"{row.get('vehicle_number', '')}"
            ):

                st.write(
                    f"**Employee:** "
                    f"{row.get('requisitioner_name', '')}"
                )

                st.write(
                    f"**Department:** "
                    f"{row.get('department', '')}"
                )

                st.write(
                    f"**Destination:** "
                    f"{row.get('destination', '')}"
                )

                st.write(
                    f"**Purpose:** "
                    f"{row.get('purpose', '')}"
                )

                st.write(
                    f"**Date:** "
                    f"{row.get('travel_date', '')}"
                )

                st.write(
                    f"**Time:** "
                    f"{row.get('start_time', '')} - "
                    f"{row.get('end_time', '')}"
                )

                st.write(
                    f"**Vehicle:** "
                    f"{row.get('vehicle_number', '')}"
                )

                st.write(
                    f"**Fixed Driver:** "
                    f"{row.get('driver_name', '')}"
                )

                st.success(
                    "Department Manager: Approved"
                )

                st.success(
                    "Vehicle Allocated"
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

                    released_at = now_str()

                    update_request(
                        request_id,
                        {
                            "status":
                                "Vehicle Released",

                            "security_released_by":
                                username,

                            "security_released_at":
                                released_at,
                        },
                    )

                    audit(
                        request_id,
                        username,
                        "Security",
                        "Vehicle Released",
                        (
                            f"Vehicle "
                            f"{row.get('vehicle_number', '')}; "
                            f"fixed driver "
                            f"{row.get('driver_name', '')}"
                        ),
                    )

                    invalidate_data_cache()

                    updated = get_gatepasses()

                    matches = updated[
                        updated["request_id"]
                        .astype(str)
                        .str.strip()
                        == request_id
                    ]

                    if not matches.empty:
                        email_row = matches.iloc[0]
                    else:
                        email_row = row

                    email_sent = (
                        notify_driver_vehicle_released(
                            email_row
                        )
                    )

                    if email_sent:
                        st.success(
                            "Vehicle released. "
                            "Assigned driver has been notified by email."
                        )
                    else:
                        st.warning(
                            "Vehicle released, "
                            "but the driver's email could not be sent."
                        )

                    st.rerun()

    # --------------------------------------------------------
    # STARTING MILEAGE
    # --------------------------------------------------------

    st.divider()

    st.subheader(
        "Starting Mileage Awaiting Verification"
    )

    dataframe = get_gatepasses()

    pending_start_mileage = dataframe[
        dataframe["status"]
        .astype(str)
        == "Pending Security Start Mileage Verification"
    ]

    if pending_start_mileage.empty:

        st.info(
            "No starting mileage is awaiting verification."
        )

    else:

        for _, row in pending_start_mileage.iterrows():

            request_id = str(
                row.get(
                    "request_id",
                    "",
                )
            )

            with st.expander(
                f"{request_id} — "
                f"{row.get('vehicle_number', '')} — "
                f"{row.get('driver_name', '')}"
            ):

                st.write(
                    f"**Vehicle:** "
                    f"{row.get('vehicle_number', '')}"
                )

                st.write(
                    f"**Driver:** "
                    f"{row.get('driver_name', '')}"
                )

                st.write(
                    f"**Travel Date:** "
                    f"{row.get('travel_date', '')}"
                )

                st.write(
                    f"**Starting Mileage:** "
                    f"{row.get('start_mileage', '')} km"
                )

                mileage_remarks = st.text_input(
                    "Starting mileage verification remarks",
                    key=f"start_mileage_remarks_{request_id}",
                )

                if st.button(
                    "Verify Starting Mileage & Allow Trip",
                    key=f"verify_start_mileage_{request_id}",
                ):

                    update_request(
                        request_id,
