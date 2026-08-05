# TrainHub — HOD Management Portal (`hod-portal`)

The **HOD Management Portal** is a separate Google Apps Script web application dedicated to Heads of Departments (HODs) and Managers.

## Key Capabilities:
1. **Domain & Company Email Restricted**: Accessible only by authenticated company account holders (`@company.com`).
2. **Requisition Form Review**: Full scroll-down view of Training Requisition Requests submitted by employees.
3. **Bottom Approval Section**: Displays HOD Employee ID, Name, and Cost Centre with managerial decision buttons:
   - **Approve**: Approves request, transitions status to Approved, stamps digital timestamp, and emails requester.
   - **Reject**: Rejects request, cancels training, stamps digital timestamp, and emails requester.
   - **Postpone**: Postpones training while retaining history in database for carry-forward/rescheduling.
   - **Reschedule**: Allows selecting a new date and updating the requisition schedule.
4. **3-Month Post Evaluation (Level 3 Review)**:
   - Filtered view displaying pending participants **under the logged-in HOD's Cost Centre only**.
   - After each evaluation submission, dynamically updates and displays the list of remaining un-evaluated participants under that HOD.

## Setup & Deployment Instructions:
1. Open Google Apps Script (`script.google.com`) and create a new project named `TrainHub HOD Portal`.
2. Deploy as Web App:
   - Execute as: `Me`
   - Access: `Anyone` (or restricted to domain users)
3. Set Script Property:
   - `SPREADSHEET_ID`: Master Database Spreadsheet ID.
   - `ALLOWED_DOMAIN`: e.g. `company.com`
4. Copy the Web App Deployment URL and configure it in `admin-system` Script Property `HOD_PORTAL_URL`.
