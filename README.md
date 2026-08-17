# Vehicle Gate Pass — V2

This version supports multiple departments/managers, password-protected managers and drivers, passwordless employee requests, vehicle time availability, security verification, and one Excel/Google Spreadsheet.

## Workbook tabs
- Departments: department, manager username/name/password
- Drivers: driver username/name/password/active
- Vehicles: vehicle number/type/active
- GatePasses: all requisitions and workflow state
- ApprovalAudit: audit trail

## Workflow
Employee (no password) -> Department Manager -> HR -> Security release/driver assignment -> Driver start mileage -> Driver end mileage -> Security final verification.

## Availability
Employees select date, departure and return times. Only vehicles free for that exact period are shown. The manager approval step re-checks availability before approving.

## Demo manager accounts
admin_manager / Admin@123
finance_manager / Finance@123
hr_manager / HR@123
sales_manager / Sales@123
operations_manager / Operations@123
it_manager / IT@123

## Demo drivers
driver01 / Driver@123
driver02 / Driver@123
driver03 / Driver@123

HR: HR@123
Security: Security@123
Administration: Admin@123

Change all credentials for real use.

## Run
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Google Sheets
Create a Google Cloud service account, enable Sheets and Drive APIs, share one Google Spreadsheet with the service-account email, and put the credentials in Streamlit Secrets. The app synchronizes all five workbook tabs into that one spreadsheet.

For production, use Google Workspace SSO/secure authentication and a real database rather than plaintext passwords and Excel as the transaction database.
