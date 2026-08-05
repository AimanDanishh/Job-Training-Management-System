# Job Training Management System (TrainHub) — Complete Setup & Deployment Guide

Welcome to the **Job Training Management System (TrainHub)**, an enterprise-grade Google Workspace application built with Google Apps Script, Google Sheets, Google Drive, and HTML5 Web Components.

This repository comprises four decoupled Apps Script projects:
1. **`admin-system`**: Admin Web Application for managing training programmes, participants, QR code generation, session scheduling, evaluation summaries, requisition forms, and automated Google Drive folder workspaces.
2. **`participant-portal`**: Lightweight, public-facing Web Application for QR code attendance check-in and Level 1 participant training evaluations.
3. **`hod-portal`**: Dedicated HOD & Managerial Portal for reviewing employee training requisition forms, digitally stamping approvals/rejections/postponements/reschedules, and conducting 3-month post-training evaluations for participants under their Cost Centre.
4. **`employee-requisition`**: Dedicated Employee Web Application for filling out and submitting Training Requisition Forms (`AP-HRD-F01-00`) with custom deployment access permissions.

---

## 🚀 System Architecture & Feature Overview

```
Job Training Management System
├── Admin System (admin-system)
│   ├── Interactive Admin Web Dashboard
│   ├── Training Programme & Session Creator
│   ├── Requisition Approval Status Tracking & History
│   ├── QR Attendance Restriction (Approved Trainings Only)
│   └── 3-Month Post-Training Notification Engine
│
├── Public Participant Portal (participant-portal)
│   ├── Session Attendance QR Check-In
│   ├── Level 1 Participant Training Evaluation (2-Week Post-Completion Limit)
│   └── Attendance Prerequisite Verification (Disallows Absent Participants)
│
├── Employee Requisition System (employee-requisition) [NEW]
│   ├── Decoupled Deployment Access & Execution Permissions
│   ├── Company Email Account Authentication
│   ├── Employee ID Lookup & Autofill (Name, Cost Centre, Position, Date)
│   └── Automated HOD Review Email Triggering
│
├── HOD System (hod-portal)
│   ├── Company Account Authentication & Access Restriction
│   ├── Requisition Form Review & Scroll-Down Approval Section
│   ├── HOD Profile Stamp (Employee ID, Name, Cost Centre, Timestamp)
│   ├── Actions: Approve / Reject / Postpone / Reschedule + Requester Status Email
│   └── 3-Month Post Evaluation (Lists pending participants under HOD Cost Centre only)
│
└── Shared Database & Drive Hierarchy
    ├── Master Google Spreadsheet
    └── Google Drive Training Workspace Folders
```

### Key Automated Capabilities:
- **Automated Lifecycle Stages**: Automatically transitions training stages:
  - `Created` ➔ `Participants Imported` ➔ `Attendance In Progress` (upon session creation or attendance check-in).
  - `Attendance In Progress` ➔ `Training Completed` & Status `Completed` (automatically when training date passes).
  - `Training Completed` / `Evaluation Completed` ➔ `Waiting for 6-Month Review` (automatically when 6 months / 180 days elapse).
- **6-Month Milestone Alerts**: Displays notification banners and visual badges (`⏰ 6-Month Review Due`) on the dashboard and training management views.
- **Formula-Equipped Drive Sheets**: Each training folder directly contains 5 per-training Google Sheets plus the Requisition Form, complete with live Google Sheets formulas (`=COUNTA(...)`, `=COUNTIF(...)`, `=AVERAGE(...)`) on dedicated **`Summary`** tabs.
- **Single Submission & Required Ratings**: Prevents repeat submissions per participant and enforces scale 1–5 answers.

---

## 🛠️ Step-by-Step Setup Guide

### Step 1: Prepare Master Google Spreadsheet & Google Drive Template

1. Open [Google Sheets](https://sheets.google.com/) and create a new Spreadsheet named `TrainHub Master Database`.
2. Copy the **Spreadsheet ID** from the URL:
   `https://docs.google.com/spreadsheets/d/`**`1a2b3c4d5e6f7g8h9i0j...`**`/edit`
3. Upload your master **Training Requisition Form** Google Sheet (`AP-HRD-F01-00`) to Google Drive and copy its **Template File ID** from the URL.

---

### Step 2: Deploy `admin-system` (Admin Web App)

1. Open [Google Apps Script](https://script.google.com/) and click **+ New project**.
2. Rename the project to `TrainHub Admin System`.
3. If using `clasp` to push files from your local terminal:
   ```bash
   cd "c:\Users\Intern IT\Documents\Job Training System\admin-system"
   clasp push
   ```
   *Or manually copy the `.js` and `.html` files into the Apps Script editor.*

4. Configure **Script Properties**:
   - In Apps Script, click **Project Settings** ⚙️ (gear icon).
   - Scroll down to **Script Properties** and click **Add script property**:

| Property Key | Example / Description |
| :--- | :--- |
| `SPREADSHEET_ID` | **Required.** Your Master Database Spreadsheet ID. |
| `TRAINING_REQUISITION_TEMPLATE_ID` | **Required.** File ID of your master `AP-HRD-F01-00` Google Sheet template. |
| `ROOT_FOLDER_ID` | Optional. Folder ID of your Google Drive root directory (defaults to `Job Training System/Training`). |
| `PUBLIC_PORTAL_URL` | Web App URL of `participant-portal` (set in Step 4). |
| `COMPANY_LOGO_URL` | Optional. Direct image URL to embed inside center of QR codes. |

5. Run Initial Setup Function:
   - In the Apps Script code editor, select `setupSheets` from the function dropdown and click **Run**.
   - This auto-creates and formats the 7 required database sheets (`Employees`, `Trainings`, `TrainingSessions`, `Attendance`, `TrainingEval`, `PostEval`, `TrainingParticipants`).

6. Deploy Admin Web App:
   - Click **Deploy** > **New deployment**.
   - Select type: **Web app**.
   - Description: `TrainHub Admin System v1.0`
   - Execute as: `Me` (*Your Google Account*)
   - Who has access: `Anyone` (or restricted to your domain).
   - Click **Deploy** and copy the **Admin Web App URL**.

---

### Step 3: Deploy `participant-portal` (Public Portal)

1. Open [Google Apps Script](https://script.google.com/) and click **+ New project**.
2. Rename the project to `TrainHub Participant Portal`.
3. Push files via `clasp` or paste code:
   ```bash
   cd "c:\Users\Intern IT\Documents\Job Training System\participant-portal"
   clasp push
   ```
4. Configure **Script Properties**:
   - Click **Project Settings** ⚙️ > **Script Properties** > **Add script property**:

| Property Key | Example / Description |
| :--- | :--- |
| `SPREADSHEET_ID` | **Required.** Same Master Database Spreadsheet ID used in `admin-system`. |
| `APP_TITLE` | `TrainHub — Participant Portal` |

5. Deploy Participant Web App:
   - Click **Deploy** > **New deployment**.
   - Select type: **Web app**.
   - Description: `TrainHub Participant Portal v1.0`
   - Execute as: `Me` (*Your Google Account*)
   - Who has access: **`Anyone`** (*Required so participants & supervisors can access without Google login*).
   - Click **Deploy** and copy the generated **Participant Portal Web App URL**.

---

### Step 4: Link Admin System & Participant Portal

1. Return to the `admin-system` Apps Script project.
2. Go to **Project Settings** ⚙️ > **Script Properties**.
3. Set `PUBLIC_PORTAL_URL` to your **Participant Portal Web App URL** from Step 3:
   `https://script.google.com/macros/s/AKfycb.../exec`
4. Click **Save script properties**.

---

## 🔗 Participant Portal URL Routing Reference

The Public Participant Portal routes requests using query parameters:

| Page / Purpose | URL Query Format |
| :--- | :--- |
| **Attendance Check-In** | `https://script.google.com/macros/s/.../exec?page=attendance&session=SES0001` |
| **Level 1 Training Evaluation** | `https://script.google.com/macros/s/.../exec?page=evaluation&id=TRN-1001` |
| **Level 3 6-Month Review** | `https://script.google.com/macros/s/.../exec?page=post&id=TRN-1001&emp=EMP-1001` |

---

## 📁 Repository Directory Sitemap

```
c:\Users\Intern IT\Documents\Job Training System\
├── README.md                          <-- (This Setup Guide)
├── Job Training Form.md               <-- Project Requirements Log
│
├── admin-system/                      <-- Admin Web Application (Apps Script)
│   ├── .clasp.json
│   ├── appsscript.json
│   ├── Code.js                        <-- Web App Routing & doGet Engine
│   ├── Helper.js                      <-- Sheet Access, Utilities & Script Properties
│   ├── Training.js                    <-- Lifecycle Stage Engine & Participant Sync
│   ├── TrainingService.js             <-- Training Lookup & Summary API
│   ├── SessionService.js              <-- Session Creator & QR Code Generator
│   ├── Attendance.js                  <-- Per-Day Attendance Recording & Sync
│   ├── AttendanceService.js           <-- Session Attendance API & Drive Sync
│   ├── Evaluation.js                  <-- Evaluation Storage & 6-Month Review API
│   ├── DriveManager.js                <-- Workspace Folder & Formula Sheet Generator
│   ├── Employee.js                    <-- Employee Lookup & Directory API
│   ├── QRService.js                   <-- QR Code Renderer & Link Converter
│   ├── ValidationService.js           <-- Data Validation & Security Checks
│   ├── Report.js                      <-- Requisition & Summary Export Engine
│   ├── dashboard.html                 <-- Admin Dashboard UI & 6-Month Alerts
│   ├── training.html                  <-- Training Management & Session QR View
│   ├── session.html                   <-- Session Schedule Management UI
│   ├── attendance.html                <-- Attendance Logs & Manual Entry UI
│   ├── evaluation.html                <-- Evaluation Table & QR Sharing UI
│   ├── report.html                    <-- Requisition Form & Report Hyperlinks
│   ├── employee.html                  <-- Employee Directory View
│   ├── index.html                     <-- Main Container Layout
│   ├── sidebar.html                   <-- Unified Sidebar Navigation
│   ├── style.html                     <-- TrainHub CSS Design Tokens
│   └── script.html                    <-- Shared Frontend Helpers & Badges
│
└── participant-portal/                <-- Public Participant Portal (Apps Script)
    ├── .clasp.json
    ├── appsscript.json
    ├── Code.gs                        <-- Public Web App Routing
    ├── Helper.gs                      <-- Spreadsheet & Config Utilities
    ├── ValidationService.gs           <-- Server-Side Security & Duplicate Protection
    ├── AttendanceService.gs           <-- Attendance Submission Backend
    ├── Evaluation.gs                  <-- Evaluation Submission Backend
    ├── Attendance.html                <-- Public Attendance Check-In UI
    ├── TrainingEvaluation.html        <-- Level 1 Participant Evaluation UI
    ├── PostEvaluation.html            <-- Level 3 6-Month Supervisor Review UI
    ├── Success.html                   <-- Submission Success View
    ├── Error.html                     <-- Rejection & Error View
    ├── style.html                     <-- Mobile-First Styling
    └── script.html                    <-- Public Frontend Client Helpers
```

---

## 🔒 Security & Data Integrity Highlights

1. **Server-Side Validation**: All public submissions pass through [`ValidationService.gs`](file:///c:/Users/Intern%20IT/Documents/Job%20Training%20System/participant-portal/ValidationService.gs) before touching the database.
2. **Duplicate Protection**: Prevents multiple attendance check-ins or duplicate evaluation submissions from the same employee.
3. **Session Auto-Expiration**: Session QR codes auto-expire when current time exceeds session end date and time.
4. **Isolated Execution**: `participant-portal` executes as the project owner, granting secure write access to the spreadsheet without exposing edit permissions to external users.
