# Vehicle Gate Pass Application

Streamlit vehicle gate-pass management system using Google Sheets as the database.

## Database

The application uses this Google Spreadsheet:

vehicle_gate_pass_data

The application automatically creates these tabs:

- Departments
- Drivers
- Vehicles
- GatePasses
- ApprovalAudit

## Workflow

Employee / Requisitioner
        ↓
Department Manager
        ↓
HR Manager
        ↓
Security
        ↓
Driver
        ↓
Security Final Verification

## Employee

Employees do not need a password.

They can enter:

- Name
- Department
- Contact number
- Person travelling with them
- Destination
- Purpose
- Travel date
- Return date
- Departure time
- Return time
- Vehicle type

The application only displays vehicles that are available for the selected period.

## Department Manager

Managers log in using passwords stored in Streamlit Secrets.

Managers can:

- View requests from their department
- Approve requests
- Reject requests
- Add rejection reasons

## HR

HR logs in using the HR password stored in Streamlit Secrets.

HR can:

- View manager-approved requests
- Approve
- Reject

## Security

Security logs in using the Security password.

Security can:

- View HR-approved requests
- Assign a driver
- Release the vehicle
- Verify completed trips
- Approve the final trip

## Driver

Drivers have individual usernames and passwords.

Drivers can:

- See their assigned trips
- Enter starting mileage
- Enter ending mileage
- Complete the trip

## Security Final Verification

After the driver enters the ending mileage:

Driver
↓
Pending Security Verification
↓
Security verifies
↓
Trip Completed - Security Approved

## Security

Passwords and Google credentials are stored only in Streamlit Secrets.

Never commit:

- secrets.toml
- Google service-account JSON
- passwords
- company credentials
- private keys

to GitHub.

## Google Cloud

Enable:

Google Sheets API

Google Drive API

The Google service account must be given Editor access to:

vehicle_gate_pass_data

## Running locally

Install requirements:

pip install -r requirements.txt

Then create:

.streamlit/secrets.toml

with your credentials.

Run:

streamlit run app.py
