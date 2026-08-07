# Job Training Management System (TrainHub) — Complete Setup & Architecture Guide

Welcome to the **Job Training Management System (TrainHub)**, an enterprise-grade Google Workspace application built with Google Apps Script, Google Sheets, Google Drive, and HTML5 Web Components.

This repository comprises four decoupled Apps Script projects:
1. **`admin-system`**: Admin Web Application for managing training programmes, participants, QR code generation, session scheduling, evaluation summaries, requisition forms, and automated Google Drive folder workspaces.
2. **`participant-portal`**: Lightweight, public-facing Web Application for QR code attendance check-in and Level 1 participant training evaluations.
3. **`hod-portal`**: Dedicated HOD & Managerial Portal for reviewing employee training requisition forms, digitally stamping multi-tier approvals (HOD, C-Suite, HOHR), and conducting 3-month post-training evaluations for participants under their Cost Centre.
4. **`employee-requisition`**: Dedicated Employee Web Application for filling out and submitting Training Requisition Forms (`AP-HRD-F01-01`) with custom deployment access permissions.

---

## 🚀 System Architecture & Feature Overview

```
Job Training Management System
├── Admin System (admin-system)
│   ├── Interactive Admin Web Dashboard
│   ├── Training Programme & Session Creator (Approved Requisitions Only)
│   ├── Requisition Approval Status Tracking & History
│   ├── QR Attendance & Evaluation QR Code Generation
│   └── 3-Month Post-Training Notification Engine & Live Countdown
│
├── Public Participant Portal (participant-portal)
│   ├── Session Attendance QR Check-In
│   ├── Level 1 Participant Training Evaluation
│   └── Attendance Prerequisite Verification
│
├── Employee Requisition System (employee-requisition)
│   ├── Employee ID Lookup & Autofill (Name, Cost Centre, Position Title, Date)
│   ├── Complete 15-Field Training Request Form (AP-HRD-F01-01)
│   ├── Dynamic Cost Centre Dropdown & Master Participant Picker
│   └── Multi-Tier Approval Trigger & Auto-Bypass Engine
│
├── HOD & Managerial System (hod-portal)
│   ├── Multi-Tier Approval Review (HOD ➔ C-Suite ➔ HOHR ➔ Arina)
│   ├── Digital Approval Stamp (Employee ID, Name, Cost Centre, Timestamp)
│   ├── Actions: Approve / Reject / Postpone / Reschedule
│   └── 3-Month Post Evaluation (Lists pending participants under HOD Cost Centre)
│
└── Shared Database & Drive Hierarchy
    ├── Master Google Spreadsheet (SPREADSHEET_ID)
    ├── Employee Database (EMPLOYEE_SPREADSHEET_ID)
    └── Google Drive Training Workspace Folders
```

### Key Capabilities & System Rules:

1. **Spreadsheet Structure & Mapping (`EMPLOYEE_SPREADSHEET_ID`)**:
   - **Cost Centre as Department**: `Company` field is completely removed across the system. `Cost Centre` is strictly used as `Department`.
   - **Position Title as Job Title**: `Position Title` is mapped as `Job Title` for all employee records.
   - **Dynamic Cost Centre Dropdown**: Populated directly from the `Cost Centre` tab in `EMPLOYEE_SPREADSHEET_ID`.
   - **Master Participant Picker**: Participant selection and bulk addition search strictly against `EMPLOYEE_SPREADSHEET_ID`.

2. **Standardized Training Categories**:
   - `Behavioral Skills (Soft Skills + Leadership + Customer Service)`
   - `Technical Skills (Technical + Quality + SHE + Digital)`
   - `Compliance Training (Compliance + Safety)`
   - `Business Skills (Finance + Sales & Marketing + Project Management)`
   - `Onboarding`

3. **Automated Lifecycle Stages & 3-Month Countdown**:
   - Sequential progression: `Created` ➔ `Participants Imported` ➔ `Attendance In Progress` ➔ `Training Completed` ➔ `Evaluation Completed` ➔ `Waiting for 3-Month Review` ➔ `Programme Closed`.
   - **3-Month Post Evaluation Countdown**: Displays remaining days (e.g. `24 days remaining`), or remaining hours if < 1 day (e.g. `14 hours remaining`).

4. **Multi-Tier Approval Workflow & Auto-Bypass Logic**:
   - Flow: **Requester** ➔ **HOD Review** ➔ **C-Suite Approval** ➔ **HOHR Approval** ➔ **Arina Acknowledgment** (`arina.ismail@apollofood.com.my`) ➔ **Admin Creation Unlocked**.
   - **Auto-Bypass Rules** (read from `HOD email` tab in `EMPLOYEE_SPREADSHEET_ID`):
     - `Requester Name == HOD Name`: Auto-verifies HOD stage upon submission.
     - `HOD Name == C-Suite Name`: Auto-verifies HOD & auto-approves C-Suite.
     - `HOD Name == C-Suite Name == HOHR Name`: Auto-verifies HOD and auto-approves both C-Suite & HOHR.
   - **Arina Notification & Admin Unlock**: After HOD, C-Suite, and HOHR approve, an automated email notifies Arina (`arina.ismail@apollofood.com.my`) and unlocks session/QR creation in the Admin System.

---

## 📧 Centralized System Email Communication Model ("Middleman")

The system operates using **one central system email account** (the deployment execution email via `MailApp.sendEmail`) acting as the central **Automated Middleman Router**.

```
                           ┌──────────────────────────┐
                           │   EMPLOYEE REQUESTER     │
                           └────────────┬─────────────┘
                                        │ Submits Form
                                        ▼
                           ┌──────────────────────────┐
                           │   CENTRAL SYSTEM EMAIL   │
                           │     ("THE MIDDLEMAN")    │
                           └──────┬────────────┬──────┘
                                  │            │
             Auto-Bypass Trigger  │            │ Approvals Needed
                                  ▼            ▼
             ┌────────────────────────┐    ┌─────────────────────────┐
             │ HOD / C-SUITE / HOHR   │    │ HOD / C-SUITE / HOHR    │
             │   Auto-Bypass Engine   │    │  Email Review Link      │
             └───────────┬────────────┘    └────────────┬────────────┘
                         │                              │
                         └──────────────┬───────────────┘
                                        │ All Approved
                                        ▼
                           ┌──────────────────────────┐
                           │ ARINA (arina.ismail@...) │
                           │  Final Acknowledgment    │
                           └────────────┬─────────────┘
                                        │ Unlocks Admin Access
                                        ▼
                           ┌──────────────────────────┐
                           │       ADMIN SYSTEM       │
                           │ (Create Session & QRs)   │
                           └──────────────────────────┘
```

### How the Central System Email Operates:
1. **Single Sender Identity**: All outbound emails (notifications, approval review links, rejection notices, and post-evaluation alerts) originate from one consistent system email address rather than personal user accounts.
2. **Automated HOD Routing**: When an employee submits a requisition, the system looks up the `HOD email` tab in `EMPLOYEE_SPREADSHEET_ID` and emails the specific HOD with a direct web review link to `hod-portal`.
3. **Step-by-Step Escalation**:
   - If HOD approves ➔ System email automatically routes the review notice to C-Suite.
   - If C-Suite approves ➔ System email automatically routes the review notice to HOHR.
   - If HOHR approves ➔ System email notifies **Arina Ismail** (`arina.ismail@apollofood.com.my`) for final acknowledgment and updates the requisition status to `Approved`.
4. **Requester Status Loopback**: Whenever an approver takes action (Approve / Reject / Postpone / Reschedule), the system email immediately transmits a status update back to the employee requester.

---

## 🛠️ Step-by-Step Setup Guide

### Step 1: Prepare Master Spreadsheets & Google Drive Template
1. Create master database spreadsheet (`SPREADSHEET_ID`).
2. Create employee database spreadsheet (`EMPLOYEE_SPREADSHEET_ID`) containing tabs:
   - `Employees`: Columns `ID`, `Name`, `Position Title`, `Cost Centre`, `Email`, `Phone`, `Status`.
   - `Cost Centre`: Column `Cost Centre` (list of real cost centres).
   - `HOD email`: Columns `Department`, `HODName`, `HODEmail`, `CsuiteName`, `CsuiteEmail`, `HohrName`, `HohrEmail`.
3. Upload `AP-HRD-F01-01` Google Sheet template (`TRAINING_REQUISITION_TEMPLATE_ID`).

### Step 2: Deploy `admin-system` (Admin Web App)
1. Push project files via `clasp push` or copy to Apps Script editor.
2. Configure Script Properties (`SPREADSHEET_ID`, `EMPLOYEE_SPREADSHEET_ID`, `TRAINING_REQUISITION_TEMPLATE_ID`, `PUBLIC_PORTAL_URL`, `HOD_PORTAL_URL`).
3. Run `setupSheets()` from Apps Script editor to initialize database headers (no dummy data added).
4. Deploy Web App as `Me` (Execute as) with access `Anyone`.

### Step 3: Deploy `employee-requisition` (Employee Requisition App)
1. Configure Script Properties (`SPREADSHEET_ID`, `EMPLOYEE_SPREADSHEET_ID`, `HOD_PORTAL_URL`).
2. Deploy Web App for domain or company employees to submit requisition requests.

### Step 4: Deploy `hod-portal` (HOD & Managerial Review App)
1. Configure Script Properties (`SPREADSHEET_ID`, `EMPLOYEE_SPREADSHEET_ID`).
2. Deploy Web App for managers to review requisitions and complete 3-Month Post Evaluations.

### Step 5: Deploy `participant-portal` (Public Participant Portal)
1. Configure Script Properties (`SPREADSHEET_ID`).
2. Deploy Web App with access `Anyone` for participant QR attendance check-in and Level 1 evaluations.

---

## 🔗 Public Portal URL Routing Reference

| Page / Purpose | URL Query Format |
| :--- | :--- |
| **Attendance Check-In** | `https://script.google.com/macros/s/.../exec?page=attendance&session=SES0001` |
| **Level 1 Training Evaluation** | `https://script.google.com/macros/s/.../exec?page=evaluation&id=TRN-1001` |
| **Level 3 3-Month Review** | `https://script.google.com/macros/s/.../exec?page=post&id=TRN-1001&emp=EMP-1001` |
| **HOD Requisition Review** | `https://script.google.com/macros/s/.../exec?page=review&id=TRN-1001` |

---

## 📁 Repository Directory Sitemap

```
c:\Users\Intern IT\Documents\Job Training System\
├── README.md                          <-- System Overview & Setup Guide
├── Job Training Form.md               <-- Project Requirements Log
│
├── admin-system/                      <-- Admin Web Application (Apps Script)
│   ├── Code.js                        <-- Web App Routing & doGet Engine
│   ├── Helper.js                      <-- Sheet Access, Utilities & Cost Centre API
│   ├── Training.js                    <-- Lifecycle Stage Engine & 3-Month Countdown
│   ├── SessionService.js              <-- Session Creator & Approval Check
│   ├── DriveManager.js                <-- Workspace Folder & Template Generator
│   ├── dashboard.html                 <-- Admin Dashboard & 3-Month Alerts
│   ├── training.html                  <-- Training Management UI
│   ├── attendance.html                <-- Attendance Logs UI
│   ├── evaluation.html                <-- Level 1 & Level 3 Evaluation UI
│   └── report.html                    <-- Requisition & Export Engine
│
├── employee-requisition/              <-- Employee Requisition System
│   ├── Code.gs                        <-- Server Routing & Multi-Tier Approval Logic
│   ├── Helper.gs                      <-- Employee Database Connection
│   ├── ValidationService.gs           <-- Employee & Company Email Validator
│   └── Requisition.html               <-- 15-Field Training Request Form
│
├── hod-portal/                        <-- HOD & Managerial Review Portal
│   ├── Code.gs                        <-- HOD Approval & Post Eval Server Engine
│   ├── HodReview.html                 <-- Digital Stamp Review Page
│   └── HodPostEvaluation.html         <-- 3-Month Post Evaluation UI
│
└── participant-portal/                <-- Public Participant Portal
    ├── Code.gs                        <-- Public Web App Router
    ├── Attendance.html                <-- Attendance Check-In UI
    ├── TrainingEvaluation.html        <-- Level 1 Evaluation UI
    └── PostEvaluation.html            <-- 3-Month Supervisor Review UI
```
