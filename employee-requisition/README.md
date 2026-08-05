# TrainHub — Employee Training Requisition Portal (`employee-requisition`)

The **Employee Training Requisition Portal** is a separate, decoupled Google Apps Script web application dedicated to employees submitting Training Requisition Forms (`AP-HRD-F01-00`).

## Key Capabilities:
1. **Decoupled Deployment Access**: Deployed independently from `participant-portal` and `hod-portal` to support custom execution and domain access policies.
2. **Company Email Restricted**: Accepts requisitions from official company email account holders (`@company.com`).
3. **Automated Requester Lookup & Fill**: Entering Employee ID auto-populates Requester Name, Cost Centre / Department, Position, and Request Date.
4. **Direct HOD Email Notification**: Submitting a requisition auto-emails the HOD/Manager with a direct review link to `hod-portal`.

## Setup & Deployment Instructions:
1. Open Google Apps Script (`script.google.com`) and create a new project named `TrainHub Employee Requisition Portal`.
2. Deploy as Web App:
   - Execute as: `Me` (or `User accessing the web app`)
   - Access: `Anyone with company account`
3. Configure Script Properties:
   - `SPREADSHEET_ID`: Master Database Spreadsheet ID.
   - `ALLOWED_DOMAIN`: e.g. `company.com`
   - `HOD_PORTAL_URL`: Web App URL of `hod-portal`.
