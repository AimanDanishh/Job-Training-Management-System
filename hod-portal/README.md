# TrainHub — HOD Management Portal (`hod-portal`)

The **HOD Management Portal** is a separate Google Apps Script web application dedicated to Heads of Departments (HODs), C-Suite Executives, and Managers.

## Key Capabilities:
1. **Domain & Company Email Restricted**: Accessible only by authenticated company account holders (`@company.com`).
2. **Auto-Display of Pending Requests**: Automatically loads and displays pending approval requests immediately upon login or session restore without requiring manual clicks.
3. **Requisition Form Review**: Full scroll-down view of Training Requisition Requests submitted by employees.
4. **Approval Actions**: Displays Approver Employee ID, Name, and Cost Centre with decision actions:
   - **Approve**: Approves request, advances to next approval stage (HOD ➔ C-Suite ➔ HOHR), stamps digital timestamp, and notifies relevant parties.
   - **Reject**: Rejects request, stamps digital timestamp, and notifies requester with reason.
   - **Return for Revision**: Sends form back to employee for corrections.
   - **Reschedule**: Allows selecting a new date and updating the requisition schedule.
5. **Cross-Portal "📊 Post Evaluation" Tab**: Navigation tab in the header that seamlessly links directly to the **3-Month Post Evaluation Dashboard** in the Participant Portal (`?page=post&emp=...`), auto-authenticating the supervisor.
6. **3-Month Post Evaluation Console (Legacy Hub)**:
   - Filtered view displaying pending participants under the logged-in HOD's Cost Centre (`?page=posteval`).

## Setup & Deployment Instructions:
1. Open Google Apps Script (`script.google.com`) and create a new project named `TrainHub HOD Portal`.
2. Deploy as Web App:
   - Execute as: `Me`
   - Access: `Anyone` (Identity is verified strictly server-side via Google Workspace Session)
3. Set Script Properties:
   - `SPREADSHEET_ID`: Master Database Spreadsheet ID.
   - `EMPLOYEE_SPREADSHEET_ID`: Employee Master Directory Spreadsheet ID (optional fallback).
   - `EMPLOYEE_PORTAL_URL`: Web App URL of `employee-requisition`.
   - `PARTICIPANT_PORTAL_URL`: Web App URL of `participant-portal`.
   - `ALLOWED_DOMAIN`: e.g. `company.com` or `apollofood.com.my`
4. Copy the Web App Deployment URL and configure it in `admin-system` Script Property `HOD_PORTAL_URL` and `employee-requisition` Script Property `HOD_PORTAL_URL`.
