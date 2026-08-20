"""Vehicle Gate Pass Management System.

Google Sheets is the only application data store.
Google service-account credentials are read only from Streamlit Secrets.
No Excel/local data file is created by this application.

Email notifications are centralized. Whenever a request moves to a new
workflow status, the system automatically determines the relevant recipient
and sends the appropriate email.
"""

import streamlit as st
import gspread

from google.oauth2.service_account import Credentials

from datetime import datetime, date, time, timedelta
import pandas as pd
import uuid
import time as time_module
from zoneinfo import ZoneInfo

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Vehicle Gate Pass",
    page_icon="🚐",
    layout="wide",
)


# ============================================================
# CONSTANTS
# ============================================================

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SHEET_NAME = "vehicle_gate_pass_data"

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
# WORKFLOW STATUSES
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
        key
        for key in required
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

    info["client_id"] = str(
        info["client_id"]
    )

    if "universe_domain" in st.secrets:
        info["universe_domain"] = st.secrets[
            "universe_domain"
        ]

    credentials = Credentials.from_service_account_info(
        info,
        scopes=SCOPES,
    )

    return gspread.authorize(credentials)


@st.cache_resource
def get_spreadsheet():

    client = get_google_client()

    sheet_id = str(
        st.secrets.get(
            "google_sheet_id",
            "",
        )
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
                getattr(
                    exc,
                    "response",
                    None,
                ),
                "status_code",
                None,
            )

            if (
                status != 429
                or attempt == max_attempts - 1
            ):
                raise

            time_module.sleep(
                min(
                    2 ** attempt,
                    16,
                )
            )


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


READ_CACHE_TTL = 60


@st.cache_data(
    ttl=READ_CACHE_TTL,
    show_spinner=False,
)
def read_sheet(name):

    worksheet = get_or_create_worksheet(name)

    return pd.DataFrame(
        _api_call_with_backoff(
            worksheet.get_all_records
        )
    )


def invalidate_data_cache():
    read_sheet.clear()


def append_row(
    name,
    row_dict,
):

    worksheet = get_or_create_worksheet(name)

    headers = SHEET_HEADERS[name]

    _api_call_with_backoff(
        lambda: worksheet.append_row(
            [
                row_dict.get(
                    header,
                    "",
                )
                for header in headers
            ],
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

    headers = SHEET_HEADERS[
        "GatePasses"
    ]

    records = read_sheet(
        "GatePasses"
    )

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

        records_from_sheet = _api_call_with_backoff(
            worksheet.get_all_records
        )

        for row_number, record in enumerate(
            records_from_sheet,
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

        col = (
            headers.index(key)
            + 1
        )

        payload.append(
            {
                "range": gspread.utils.rowcol_to_a1(
                    target_row,
                    col,
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
# BASIC HELPERS
# ============================================================

def now_str():

    return datetime.now(
        ZoneInfo("Asia/Colombo")
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
        + datetime.now().strftime(
            "%Y%m%d"
        )
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
    return read_sheet(
        "GatePasses"
    )


# ============================================================
# EMAIL SYSTEM
# ============================================================

def send_email(
    subject,
    body,
    recipient,
):

    recipient = str(
        recipient or ""
    ).strip()

    if not recipient:
        return False

    try:

        sender = str(
            st.secrets[
                "gmail_sender"
            ]
        ).strip()

        password = str(
            st.secrets[
                "gmail_password"
            ]
        ).strip()

        if not sender or not password:
            return False

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
            "smtp.gmail.com",
            587,
            timeout=30,
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


# ============================================================
# EMAIL RECIPIENT LOOKUP
# ============================================================

def get_email_by_username(username):

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


def get_email_by_role(role):

    role = str(
        role or ""
    ).strip().lower()

    if not role:
        return None

    users = get_users()

    if users.empty:
        return None

    matches = users[
        users["role"]
        .astype(str)
        .str.strip()
        .str.lower()
        == role
    ]

    if matches.empty:
        return None

    for _, user in matches.iterrows():

        email = str(
            user.get(
                "email",
                "",
            )
        ).strip()

        if email:
            return email

    return None


def get_manager_email(row):

    manager_username = str(
        row.get(
            "manager_username",
            "",
        )
    ).strip()

    email = get_email_by_username(
        manager_username
    )

    if email:
        return email

    return get_email_by_role(
        "Department Manager"
    )


def get_driver_email(row):

    driver_username = str(
        row.get(
            "driver_username",
            "",
        )
    ).strip()

    if driver_username:

        email = get_email_by_username(
            driver_username
        )

        if email:
            return email

    driver_name = str(
        row.get(
            "driver_name",
            "",
        )
    ).strip()

    if not driver_name:
        return None

    users = get_users()

    if users.empty:
        return None

    matches = users[
        users["name"]
        .astype(str)
        .str.strip()
        .str.lower()
        == driver_name.lower()
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


def get_requisitioner_email(row):

    users = get_users()

    if users.empty:
        return None

    name = str(
        row.get(
            "requisitioner_name",
            "",
        )
    ).strip()

    if not name:
        return None

    matches = users[
        users["name"]
        .astype(str)
        .str.strip()
        .str.lower()
        == name.lower()
    ]

    if not matches.empty:

        email = str(
            matches.iloc[0].get(
                "email",
                "",
            )
        ).strip()

        if email:
            return email

    return None


# ============================================================
# CENTRALIZED WORKFLOW EMAIL NOTIFICATION
# ============================================================

def get_notification_recipient(
    row,
    status,
):

    status = str(
        status or ""
    ).strip()

    if status == "Pending Department Manager":

        return (
            get_manager_email(row),
            "Department Manager",
        )

    if status == "Pending Vehicle Allocation":

        return (
            get_email_by_role(
                "Vehicle Allocator"
            ),
            "Vehicle Allocator",
        )

    if status == "Pending HR":

        return (
            get_email_by_role(
                "HR Manager"
            ),
            "HR Manager",
        )

    if status == "Pending Security":

        return (
            get_email_by_role(
                "Security"
            ),
            "Security",
        )

    if status == "Vehicle Released":

        return (
            get_driver_email(row),
            "Driver",
        )

    if status == "Pending Security Start Mileage Verification":

        return (
            get_email_by_role(
                "Security"
            ),
            "Security",
        )

    if status == "Pending Security Verification":

        return (
            get_email_by_role(
                "Security"
            ),
            "Security",
        )

    if status in {
        "Rejected by Department Manager",
        "Rejected by HR",
        "Completed",
    }:

        return (
            get_requisitioner_email(row),
            "Employee / Requisitioner",
        )

    return None, None


def workflow_email_subject(
    request_id,
    status,
):

    subjects = {

        "Pending Department Manager":
            f"Vehicle Request Requires Manager Approval - {request_id}",

        "Pending Vehicle Allocation":
            f"Vehicle Allocation Required - {request_id}",

        "Pending HR":
            f"HR Approval Required - {request_id}",

        "Pending Security":
            f"Security Verification Required - {request_id}",

        "Vehicle Released":
            f"Vehicle Released - Driver Action Required - {request_id}",

        "Pending Security Start Mileage Verification":
            f"Starting Mileage Verification Required - {request_id}",

        "Pending Security Verification":
            f"Final Vehicle Gate Pass Verification Required - {request_id}",

        "Rejected by Department Manager":
            f"Vehicle Request Rejected by Department Manager - {request_id}",

        "Rejected by HR":
            f"Vehicle Request Rejected by HR - {request_id}",

        "Completed":
            f"Vehicle Gate Pass Completed - {request_id}",
    }

    return subjects.get(
        status,
        f"Vehicle Gate Pass Update - {request_id}",
    )


def build_workflow_email_body(
    row,
    status,
    recipient_role,
):

    request_id = str(
        row.get(
        )      
