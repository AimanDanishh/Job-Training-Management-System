# Job Training Management System (TrainHub) — System Overview & Deployment Guide

Welcome to the **Job Training Management System (TrainHub)**, an enterprise-grade Google Workspace application built with Google Apps Script, Google Sheets, Google Drive, HTML5 Web Components, and Vanilla CSS.

TrainHub automates end-to-end employee training workflows: from initial Training Requisition Form (`AP-HRD-F01-01`) submission, multi-tier managerial approvals (HOD ➔ C-Suite ➔ HOHR ➔ HR Acknowledgment), automated Google Drive workspace generation, QR-code session check-ins, Level 1 participant training evaluations, to 3-month supervisor post-training evaluations and automated annual training plan exports.

---

## 📐 1. System Architecture & Project Breakdown

The system comprises **four decoupled Google Apps Script projects** sharing a single Google Spreadsheet database and Google Drive hierarchy:

```
Job Training Management System (TrainHub)
│
├── 1. Admin System (admin-system)
│   ├── Interactive Admin Web Dashboard & Statistics
│   ├── Training Requisition Tracker & Status Overview
│   ├── Approved Requisition Session Creator & Scheduling
│   ├── Attendance & Level 1 / Level 3 QR Code Generator
│   ├── 3-Month Post-Training Countdown & Email Alert Engine
│   ├── Excel / Sheets Report Generator (Training Hours, Costs, Title Report)
│   ├── Annual Training Plan Export Engine
│   └── Automated Google Drive Workspace & AP-HRD-F01-01 Form File Generator
│
├── 2. Employee Requisition System (employee-requisition)
│   ├── Employee ID Lookup & Auto-fill (Name, Cost Centre, Position Title, Date)
│   ├── 15-Field Training Requisition Form (AP-HRD-F01-01)
│   ├── Dynamic Cost Centre Dropdown & Master Employee Participant Picker
│   ├── "My Requests" Submission History View
│   ├── Resubmission Engine for Returned/Postponed Requisitions
│   └── Multi-Tier Approval Trigger & Auto-Bypass Router
│
├── 3. HOD & Managerial Review Portal (hod-portal)
│   ├── Multi-Tier Approval Review Interface (HOD ➔ C-Suite ➔ HOHR)
│   ├── Digital Approval Stamping Engine (Status, ID, Name, Position, Timestamp)
│   ├── Decision Actions: Approve / Reject / Postpone / Reschedule (with Remarks)
│   └── 3-Month Post Evaluation (Filtered strictly by HOD Cost Centre)
│
├── 4. Public Participant Portal (participant-portal)
│   ├── Public QR Code Attendance Check-In (No Google Login Required)
│   ├── Level 1 Participant Training Evaluation (7-scale Likert + Feedback)
│   ├── 3-Month Supervisor Post-Training Evaluation Engine
│   └── Server-Side Duplicate Check & Session Status Validation
│
└── 5. Shared Database & Drive Hierarchy
    ├── Master Database Spreadsheet (SPREADSHEET_ID)
    ├── Employee Master Database Spreadsheet (EMPLOYEE_SPREADSHEET_ID)
    └── Google Drive Workspace Folders & AP-HRD-F01-01 Template Sheet
```

---

## 🔄 2. Key Business Workflows & Engine Rules

### 2.1 Centralized Email Router Model ("Middleman")
All outbound system emails (approval links, status updates, rejection notices, and 3-month evaluation reminders) originate from **one central system account** (`it@apollofood.com.my`) using Google Apps Script `MailApp.sendEmail`.

```
                  ┌─────────────────────────────────────┐
                  │    EMPLOYEE REQUISITION SUBMISSION  │
                  └──────────────────┬──────────────────┘
                                     │ Submits Request Form
                                     ▼
                  ┌─────────────────────────────────────┐
                  │    CENTRAL SYSTEM EMAIL ACCOUNT     │
                  │   ("AUTOMATED MIDDLEMAN ROUTER")    │
                  └──────────┬──────────────────┬───────┘
                             │                  │
         Auto-Bypass Trigger │                  │ Approval Required
                             ▼                  ▼
              ┌─────────────────────┐    ┌─────────────────────┐
              │ Auto-Bypass Engine  │    │ Email Review Link   │
              │ (HOD/Csuite/Hohr)   │    │ Sent to Approver    │
              └──────────┬──────────┘    └──────────┬──────────┘
                         │                          │
                         └────────────┬─────────────┘
                                      │ All Approvals Granted
                                      ▼
                  ┌─────────────────────────────────────┐
                  │      ARINA (HR ACKNOWLEDGMENT)      │
                  │     (arina.ismail@apollo...)        │
                  └──────────────────┬──────────────────┘
                                     │ Unlocks Admin Creation
                                     ▼
                  ┌─────────────────────────────────────┐
                  │            ADMIN SYSTEM             │
                  │  (Create Sessions, Print QRs, Plan) │
                  └─────────────────────────────────────┘
```

### 2.2 Multi-Tier Approval Workflow & Auto-Bypass Rules
1. **Approval Chain**: Requester ➔ HOD Review ➔ C-Suite Approval ➔ HOHR Approval ➔ HR Acknowledgment (Arina Ismail).
2. **Auto-Bypass Rules** (evaluated dynamically against `HOD email` tab in `EMPLOYEE_SPREADSHEET_ID`):
   - **Requester == HOD**: Automatically verifies HOD stage upon form submission.
   - **HOD == C-Suite**: Auto-verifies HOD and auto-approves C-Suite stage.
   - **HOD == C-Suite == HOHR**: Auto-verifies HOD and auto-approves both C-Suite & HOHR stages.
3. **HR Acknowledgment**: When HOD, C-Suite, and HOHR approve, an automated email alerts HR (`arina.ismail@apollofood.com.my`), updates status to `Approved`, stamps digital signature details on `AP-HRD-F01-01`, and unlocks session scheduling in `admin-system`.

### 2.3 Training Lifecycle Stages
Sequential progression enforced across the system:
1. `Created`: Initial requisition submitted/created.
2. `Participants Imported`: Participants assigned from Employee Database.
3. `Attendance In Progress`: Active session QR codes created and live.
4. `Training Completed`: Training end date reached.
5. `Evaluation Completed`: Level 1 evaluations recorded.
6. `Waiting for 3-Month Review`: Live countdown active for 3-month post evaluation.
7. `Programme Closed`: 3-month post evaluations finalized.

---

## 📊 3. Database Schema & Structure

### 3.1 Master Database Spreadsheet (`SPREADSHEET_ID`)
- **`Trainings`**: Training programmes, lifecycle stage, fees, dates, trainer, provider, requisition ID.
- **`TrainingSessions`**: Individual sessions, date, start/end time, venue, QR code links, session status.
- **`TrainingParticipants`**: Employee ID, Name, Department (Cost Centre), Job Title (Position Title), attendance status.
- **`Attendance`**: Timestamp, Session ID, Employee ID, Name, Department, check-in status, remarks.
- **`TrainingEval`**: Level 1 Likert ratings (Q1–Q7), feedback, timestamp, Employee ID.
- **`PostEval`**: 3-Month Post Evaluation ratings, supervisor Employee ID, competencies, timestamp.
- **`Requisitions`**: 15-field request records, approval statuses (HOD, C-Suite, HOHR, Arina), digital stamps.

### 3.2 Employee Master Database Spreadsheet (`EMPLOYEE_SPREADSHEET_ID`)
- **`Employees`**: Columns `ID`, `Name`, `Position Title` (mapped system-wide as Job Title), `Cost Centre` (mapped system-wide as Department), `Email`, `Phone`, `Status`.
- **`Cost Centre`**: List of valid company cost centres / departments.
- **`HOD email`**: Department mapping containing `Department`, `HODName`, `HODEmail`, `CsuiteName`, `CsuiteEmail`, `HohrName`, `HohrEmail`.

---

## 🛠️ 4. Complete Step-by-Step Deployment Guide

### Prerequisites
1. **Google Account**: Google Workspace administrator or deployment account (recommended: `it@apollofood.com.my`).
2. **Google Apps Script API**: Enabled in [Google Apps Script Settings](https://script.google.com/home/usersettings).
3. **Clasp CLI** (Optional, for command line deployment):
   ```bash
   npm install -g @google/clasp
   clasp login
   ```

---

### Step 1: Prepare Master Spreadsheets & Google Drive Templates

1. **Create Master Database Spreadsheet**:
   - Create a blank Google Sheet named `TrainHub Master Database`.
   - Copy its Spreadsheet ID from the URL (`https://docs.google.com/spreadsheets/d/`**`SPREADSHEET_ID`**`/edit`).

2. **Create Employee Master Spreadsheet**:
   - Create a Google Sheet named `TrainHub Employee Master Database`.
   - Setup tabs:
     - **`Employees`**: Headers: `ID`, `Name`, `Position Title`, `Cost Centre`, `Email`, `Phone`, `Status`.
     - **`Cost Centre`**: Header: `Cost Centre` (list of official cost centres).
     - **`HOD email`**: Headers: `Department`, `HODName`, `HODEmail`, `CsuiteName`, `CsuiteEmail`, `HohrName`, `HohrEmail`.
   - Copy its Spreadsheet ID from the URL (`EMPLOYEE_SPREADSHEET_ID`).

3. **Upload `AP-HRD-F01-01` Training Requisition Form Template**:
   - Upload the master Google Sheet template of `AP-HRD-F01-01` to Google Drive.
   - Copy the File ID from its URL (`TRAINING_REQUISITION_TEMPLATE_ID`).

4. **Create Root Drive Folder**:
   - Create a Google Drive folder named `TrainHub Workspaces`.
   - Copy the Folder ID from its URL (`ROOT_FOLDER_ID`).
   - The application creates and uses `TrainHub Workspaces/Training/<training>` for every training. Do not set `ROOT_FOLDER_ID` to an individual training folder or to the `Training` child folder.

---

### Step 2: Deploy `admin-system` (Admin Web Dashboard)

1. Navigate to directory:
   ```bash
   cd "c:\Users\Intern IT\Documents\Job Training System\admin-system"
   ```
2. Push code using clasp or copy `.js` and `.html` files into a new project at [script.google.com](https://script.google.com):
   ```bash
   clasp push
   ```
3. In the Apps Script Editor, open **Project Settings** ⚙️ ➔ **Script Properties** and add:

   | Key | Description | Example / Required Value |
   | :--- | :--- | :--- |
   | `SPREADSHEET_ID` | Master Database ID | `1a2b3c...` |
   | `EMPLOYEE_SPREADSHEET_ID` | Employee Database ID | `1x2y3z...` |
   | `TRAINING_REQUISITION_TEMPLATE_ID` | `AP-HRD-F01-01` Sheet Template ID | `1t2u3v...` |
   | `ROOT_FOLDER_ID` | Google Drive Workspace Folder ID | `1f2g3h...` |
   | `ALLOWED_DOMAIN` | Allowed company domain | `apollofood.com.my` |
   | `ADMIN_EMAILS` | Authorized Admin Emails | `it@apollofood.com.my,arina.ismail@apollofood.com.my` |
   | `PUBLIC_PORTAL_URL` | Deployed Participant Portal URL | *(Populate in Step 6)* |
   | `HOD_PORTAL_URL` | Deployed HOD Portal URL | *(Populate in Step 6)* |

4. Run Initializer Function:
   - Select `initDefaultScriptProperties` from the function dropdown and click **Run**.
   - Select `setupSheets` from the function dropdown and click **Run** (this creates required sheet headers in the Master Database without inserting dummy data).
5. Deploy Web App:
   - Click **Deploy** ➔ **New deployment**.
   - Select type: **Web app**.
   - Description: `Admin System v1.0`
   - Execute as: **Me** (`it@apollofood.com.my`)
   - Who has access: **Anyone** (or *Anyone within domain*)
   - Click **Deploy** and save the **Web App URL**.

---

### Step 3: Deploy `employee-requisition` (Employee Requisition Portal)

1. Navigate to directory:
   ```bash
   cd "c:\Users\Intern IT\Documents\Job Training System\employee-requisition"
   ```
2. Push code via clasp or paste files to Apps Script:
   ```bash
   clasp push
   ```
3. Configure **Script Properties**:

   | Key | Description | Example Value |
   | :--- | :--- | :--- |
   | `SPREADSHEET_ID` | Master Database ID | `1a2b3c...` |
   | `EMPLOYEE_SPREADSHEET_ID` | Employee Database ID | `1x2y3z...` |
   | `HOD_PORTAL_URL` | Deployed HOD Portal URL | *(Populate in Step 6)* |
   | `ALLOWED_DOMAIN` | Company Domain | `apollofood.com.my` |
   | `ROOT_FOLDER_ID` | Same TrainHub Workspaces root used by Admin | `1f2g3h...` |

4. Deploy Web App:
   - Click **Deploy** ➔ **New deployment** ➔ **Web app**.
   - Execute as: **Me**
   - Who has access: **Anyone** (or *Anyone with company account*)
   - Save the **Web App URL**.

---

### Step 4: Deploy `hod-portal` (HOD & Managerial Review Portal)

1. Navigate to directory:
   ```bash
   cd "c:\Users\Intern IT\Documents\Job Training System\hod-portal"
   ```
2. Push code via clasp or paste files to Apps Script:
   ```bash
   clasp push
   ```
3. Configure **Script Properties**:

   | Key | Description | Example Value |
   | :--- | :--- | :--- |
   | `SPREADSHEET_ID` | Master Database ID | `1a2b3c...` |
   | `EMPLOYEE_SPREADSHEET_ID` | Employee Database ID | `1x2y3z...` |
   | `EMPLOYEE_PORTAL_URL` | Deployed Employee Requisition App URL | `https://script.google.com/macros/s/.../exec` |
   | `ALLOWED_DOMAIN` | Company Domain | `apollofood.com.my` |

4. Deploy Web App:
   - Click **Deploy** ➔ **New deployment** ➔ **Web app**.
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Save the **Web App URL**.

---

### Step 5: Deploy `participant-portal` (Public Participant Portal)

1. Navigate to directory:
   ```bash
   cd "c:\Users\Intern IT\Documents\Job Training System\participant-portal"
   ```
2. Push code via clasp or paste files to Apps Script:
   ```bash
   clasp push
   ```
3. Configure **Script Properties**:

   | Key | Description | Example Value |
   | :--- | :--- | :--- |
   | `SPREADSHEET_ID` | Master Database ID | `1a2b3c...` |
   | `APP_TITLE` | Application Title | `TrainHub — Participant Portal` |

4. Deploy Web App:
   - Click **Deploy** ➔ **New deployment** ➔ **Web app**.
   - Execute as: **Me**
   - Who has access: **Anyone** *(Must be set to Anyone to allow non-Google participant QR check-ins)*
   - Save the **Web App URL**.

---

### Step 6: Inter-App Link Wireup & Final Configuration

Once all four Apps Script web apps are deployed, link them together by updating their Script Properties:

> All four projects have separate Script Properties. `SPREADSHEET_ID` must be the identical master database ID in `admin-system`, `employee-requisition`, `hod-portal`, and `participant-portal`; otherwise a portal can continue to display records from an older database.

For workspaces that were created before this rule, set `ROOT_FOLDER_ID` in the Admin project and run `migrateTrainingWorkspacesToConfiguredRoot` once from the Apps Script editor. It moves only folders referenced by the `Trainings` sheet into `ROOT_FOLDER_ID/Training`.

1. **In `admin-system` Script Properties**:
   - Set `PUBLIC_PORTAL_URL` = Web App URL of `participant-portal`.
   - Set `HOD_PORTAL_URL` = Web App URL of `hod-portal`.
2. **In `employee-requisition` Script Properties**:
   - Set `HOD_PORTAL_URL` = Web App URL of `hod-portal`.
3. **In `hod-portal` Script Properties**:
   - Set `EMPLOYEE_PORTAL_URL` = Web App URL of `employee-requisition`.

---

## 🔗 5. Public Portal & Deep Link Routing Reference

| Application / Page | Query Parameter Format | Description |
| :--- | :--- | :--- |
| **Attendance Check-In** | `/exec?page=attendance&session=SES0001` | Scanned by participants from session QR codes. |
| **Level 1 Evaluation** | `/exec?page=evaluation&id=TRN-1001` | Participant post-training evaluation form. |
| **Level 3 3-Month Post Evaluation** | `/exec?page=post&id=TRN-1001&emp=EMP-1001` | Direct supervisor evaluation link. |
| **HOD Requisition Review** | `/exec?page=review&id=TRN-1001` | Direct review link sent to HOD/C-Suite/HOHR. |
| **My Requisitions List** | `/exec?page=my_requests` | Employee requisition history & resubmission portal. |

---

## 🔧 6. Clasp Command Quick Reference

Each sub-directory contains its pre-configured `.clasp.json` pointing to its Google Apps Script project ID:

```bash
# Admin System
cd admin-system && clasp push

# Employee Requisition Portal
cd employee-requisition && clasp push

# HOD Review Portal
cd hod-portal && clasp push

# Public Participant Portal
cd participant-portal && clasp push
```

> ⚠️ **Important Note on Deployment Updates**:  
> Running `clasp push` updates the project code. To make changes active on the Web App URLs, go to **Deploy** ➔ **Manage deployments** ➔ Edit active deployment ➔ Select **New version** ➔ **Save**.

---

## 📁 7. Comprehensive Directory Sitemap

```
Job Training Management System/
├── README.md                          <-- Master System Architecture & Deployment Guide
├── Job Training Form.md               <-- Requirements Log & Change History
│
├── admin-system/                      <-- Admin Web Application
│   ├── .clasp.json                    <-- Clasp Project ID Config
│   ├── appsscript.json                <-- Manifest File
│   ├── Code.js                        <-- Web App Router & doGet Engine
│   ├── Helper.js                      <-- Sheet Access, Script Properties & Setup Engine
│   ├── Training.js                    <-- Lifecycle Stage Engine & 3-Month Countdown
│   ├── TrainingService.js             <-- Programme CRUD Services
│   ├── SessionService.js              <-- Session Creator & Scheduling Logic
│   ├── AttendanceService.js           <-- Attendance Records Engine
│   ├── Evaluation.js                  <-- Level 1 & Level 3 Evaluation Engine
│   ├── Employee.js                    <-- Employee Database Lookup & Utilities
│   ├── QRService.js                   <-- QR Code Generation Engine
│   ├── DriveManager.js                <-- Workspace Folder & AP-HRD-F01-01 Generator
│   ├── Report.js                      <-- Excel/Sheets Export & Annual Training Plan Generator
│   ├── ValidationService.js           <-- Server-Side Input Validator
│   ├── index.html                     <-- Main Admin Container UI
│   ├── dashboard.html                 <-- Admin Dashboard UI
│   ├── training.html                  <-- Training Management UI
│   ├── session.html                   <-- Session Scheduling UI
│   ├── attendance.html                <-- Attendance Log UI
│   ├── evaluation.html                <-- Evaluation Summary UI
│   ├── employee.html                  <-- Employee Lookup UI
│   ├── report.html                    <-- Reports & Annual Training Plan Export UI
│   ├── sidebar.html                   <-- Navigation Sidebar UI
│   ├── style.html                     <-- CSS Design System
│   └── script.html                    <-- Frontend JS Controllers
│
├── employee-requisition/              <-- Employee Requisition Portal
│   ├── .clasp.json                    <-- Clasp Project ID Config
│   ├── appsscript.json                <-- Manifest File
│   ├── README.md                      <-- Portal Specific Readme
│   ├── Code.gs                        <-- Server Routing & Multi-Tier Approval Engine
│   ├── Helper.gs                      <-- Employee DB Connection & Requisition Logic
│   ├── ValidationService.gs           <-- Form & Email Validator
│   ├── Requisition.html               <-- 15-Field Training Request Form UI
│   ├── Success.html                   <-- Submission Confirmation Page
│   ├── Error.html                     <-- Error Boundary Page
│   ├── style.html                     <-- Portal CSS Styling
│   └── script.html                    <-- Frontend JS Form Handler
│
├── hod-portal/                        <-- HOD & Managerial Review Portal
│   ├── .clasp.json                    <-- Clasp Project ID Config
│   ├── appsscript.json                <-- Manifest File
│   ├── README.md                      <-- Portal Specific Readme
│   ├── Code.gs                        <-- Approval Routing & Post Eval Engine
│   ├── Helper.gs                      <-- Spreadsheet & HOD Cost Centre Service
│   ├── ValidationService.gs           <-- Managerial Access Validator
│   ├── HodReview.html                 <-- Digital Stamp Requisition Review UI
│   ├── HodPostEvaluation.html         <-- 3-Month Post Evaluation UI (Cost Centre Filtered)
│   ├── Success.html                   <-- Approval Confirmation UI
│   ├── Error.html                     <-- Access Denied / Error UI
│   ├── style.html                     <-- HOD Portal CSS Styling
│   └── script.html                    <-- Frontend Interactions JS
│
└── participant-portal/                <-- Public Participant Portal
    ├── .clasp.json                    <-- Clasp Project ID Config
    ├── appsscript.json                <-- Manifest File
    ├── README.md                      <-- Participant Portal Setup Guide
    ├── Code.gs                        <-- Public Web App Router
    ├── Helper.gs                      <-- Database Connection Service
    ├── AttendanceService.gs           <-- QR Check-In Backend Handler
    ├── Evaluation.gs                  <-- Level 1 & Level 3 Submission Handler
    ├── ValidationService.gs           <-- Server-Side Access & Duplicate Validator
    ├── Attendance.html                <-- Public QR Attendance Check-In UI
    ├── TrainingEvaluation.html        <-- Level 1 Evaluation UI
    ├── PostEvaluation.html            <-- 3-Month Supervisor Review UI
    ├── Success.html                   <-- Submission Success Page
    ├── Error.html                     <-- Error Page UI
    ├── style.html                     <-- Participant Portal Styling
    └── script.html                    <-- Form Interactive JS
```
