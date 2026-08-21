# Job Training Management System (TrainHub) — Enterprise Documentation & Deployment Guide

[![Platform](https://img.shields.io/badge/Google%20Workspace-Apps%20Script-4285F4?logo=google&logoColor=white)](https://developers.google.com/apps-script)
[![Database](https://img.shields.io/badge/Database-Google%20Sheets-34A853?logo=googlesheets&logoColor=white)](https://www.google.com/sheets/about/)
[![Storage](https://img.shields.io/badge/Storage-Google%20Drive-FBBC05?logo=googledrive&logoColor=white)](https://www.google.com/drive/)
[![Frontend](https://img.shields.io/badge/Frontend-HTML5%20%7C%20CSS3%20%7C%20Vanilla%20JS-E34F26?logo=html5&logoColor=white)](https://developer.mozilla.org/en-US/docs/Web)
[![Compliance](https://img.shields.io/badge/Compliance-AP--HRD--F01--01-8B5CF6)](https://apollofood.com.my)
[![License](https://img.shields.io/badge/License-Proprietary-gray)](#)

---

## 📖 Executive Summary

The **Job Training Management System (TrainHub)** is an enterprise-grade, serverless training governance and tracking suite engineered for Google Workspace. Built with **Google Apps Script**, **Google Sheets**, **Google Drive**, and modern responsive web interfaces, TrainHub orchestrates the entire corporate training lifecycle without requiring external cloud databases or third-party servers.

### Core Automated Capabilities:
- 📝 **Digital Training Requisition (`AP-HRD-F01-01`)**: 15-field employee requisition portal with dynamic Cost Centre selection, employee directory lookup, brochure file attachments, and automated requisition sheet generation.
- ⚡ **Multi-Tier Approval & Auto-Bypass Engine**: Dynamically routes requests through HOD review, C-Suite approval, Head of HR (HOHR) sign-off, and final HR acknowledgment with intelligent multi-role auto-bypass logic.
- 📁 **Automated Drive Workspace Generator**: Automatically provisions isolated Google Drive training folders, copies and pre-fills official `AP-HRD-F01-01` form templates, and generates individual training data spreadsheets.
- 📱 **Public QR Attendance Check-In**: Non-Google login QR check-in portal for participants with session expiry enforcement and duplicate check-in prevention.
- ⭐ **Two-Stage Training Evaluation**:
  - **Level 1 (Reaction)**: 7-criteria Likert feedback form submitted immediately upon training completion.
  - **Level 3 (Behavior / 3-Month Post-Evaluation)**: Automated 90-day countdown timer notifying supervisors to evaluate workplace competency improvements.
- 📊 **Enterprise Analytics & Excel Export Engine**: Live dashboards and automated Google Sheet / Microsoft Excel (`.xlsx`) generation for Annual Training Plans (ATP), Training Hours by Cost Centre, Cost Breakdown Reports, and Employee Training Records.

---

## 📐 1. System Architecture & Micro-Project Breakdown

TrainHub is designed as a **four-tier decoupled micro-application architecture** sharing a centralized Google Sheets database and Google Drive hierarchy:

```
Job Training Management System (TrainHub)
│
├── 1. Admin System [admin-system]
│   ├── Interactive Master Dashboard (KPIs, Active Trainings, Approval Pipeline)
│   ├── Training Programme & Session Manager (Lifecycle Stage Automations)
│   ├── Session Attendance Monitor & Manual Attendance Adjustments
│   ├── Branded QR Code Generator (Attendance, Level 1 & Level 3 Evaluations)
│   ├── 3-Month Post-Evaluation Countdown & Supervisor Assignment Engine
│   ├── Unified Report Generator (Hours, Cost, Title, Employee, ATP)
│   └── Direct-to-Drive Excel (.xlsx) Export Engine
│
├── 2. Employee Requisition Portal [employee-requisition]
│   ├── Employee ID Directory Lookup (Auto-fill Name, Cost Centre, Position)
│   ├── 15-Field Training Request Form (AP-HRD-F01-01 Standard)
│   ├── Multi-Participant Picker & Dynamic Cost Centre Dropdowns
│   ├── Brochure / Document Upload Handler
│   ├── "My Requests" Submission History & Resubmission Engine for Returned Forms
│   └── Central Email Router Trigger (Originating from System Service Account)
│
├── 3. HOD & Managerial Review Portal [hod-portal]
│   ├── Server-Side Identity Resolution (Restricted to Authorized Approvers)
│   ├── Multi-Tier Approval Review Interface (HOD ➔ C-Suite ➔ HOHR)
│   ├── Decision Actions: Approve / Reject / Return / Reschedule (with Remarks)
│   ├── Automated Digital Approval Stamping on Google Sheet Templates
│   └── Cost-Centre-Filtered 3-Month Post-Evaluation Review Console
│
├── 4. Public Participant Portal [participant-portal]
│   ├── Public Mobile-Optimized QR Attendance Check-In (No Google Login Required)
│   ├── Level 1 Participant Training Evaluation Form (7-Point Likert + Feedback)
│   ├── Level 3 Direct Supervisor Post-Training Review Form
│   └── Server-Side Enrollment, Expiry, and Duplicate Submission Validations
│
└── 5. Shared Infrastructure & Storage [Google Drive & Google Sheets]
    ├── Master Database Spreadsheet (SPREADSHEET_ID)
    ├── Employee Master Directory Spreadsheet (EMPLOYEE_SPREADSHEET_ID)
    ├── Root Workspaces Directory (ROOT_FOLDER_ID)
    └── AP-HRD-F01-01 Master Sheet Template (TRAINING_REQUISITION_TEMPLATE_ID)
```

---

## 🔄 2. Key Business Workflows & Engine Rules

### 2.1 Centralized Email Router Model ("Middleman")
To guarantee 100% email delivery without exposing personal employee Gmail inboxes or hitting individual Apps Script email quotas, all notifications (approval requests, status updates, rejection notices, and post-evaluation alerts) are routed through a **single centralized system account** (`it@apollofood.com.my`) using `MailApp.sendEmail`.

```
                  ┌─────────────────────────────────────┐
                  │    EMPLOYEE REQUISITION SUBMISSION  │
                  └──────────────────┬──────────────────┘
                                     │ Submits Request Form
                                     ▼
                  ┌─────────────────────────────────────┐
                  │    CENTRAL SYSTEM EMAIL ACCOUNT     │
                  │   ("AUTOMATED MIDDLEMAN ROUTER")    │
                  │       (it@apollofood.com.my)        │
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

---

### 2.2 Multi-Tier Approval Chain & Auto-Bypass Rules

Approval routing is dynamically resolved using the **`For IT`** and **`HOD email`** sheets in `EMPLOYEE_SPREADSHEET_ID`:

| Stage | Reviewer Role | Description |
| :--- | :--- | :--- |
| **Stage 1** | **Head of Department (HOD)** | Verified by the direct department head or manager mapped to the requester. |
| **Stage 2** | **C-Suite Executive** | Approved by the executive overseeing the operational division. |
| **Stage 3** | **Head of Human Resources (HOHR)** | Approved by Head of HR for compliance and budget authorization. |
| **Stage 4** | **HR Department Acknowledgment** | Acknowledged by HR Administrator (`arina.ismail@apollofood.com.my`) to unlock admin scheduling. |

#### ⚡ Dynamic Auto-Bypass Rules:
1. **Requester is HOD**: When the requester's name matches the assigned HOD, Stage 1 is marked `Auto-Verified` immediately upon submission.
2. **HOD is C-Suite**: When the HOD name matches the C-Suite Executive, Stage 1 is verified and Stage 2 is marked `Auto-Approved`.
3. **HOD is C-Suite is HOHR**: When one person holds HOD, C-Suite, and HOHR roles, Stages 1, 2, and 3 are automatically verified and approved simultaneously.
4. **Form Return & Resubmission**: If an approver marks a request as **Returned**, the status becomes `Pending Revision`. The employee can edit and resubmit via the Requisition Portal, which resets downstream approvals and re-triggers the chain.

---

### 2.3 Training Programme Lifecycle Stages

TrainHub strictly enforces sequential lifecycle stage transitions across the administration backend:

```
[Created] ➔ [Participants Imported] ➔ [Attendance In Progress] ➔ [Training Completed] ➔ [Evaluation Completed] ➔ [Waiting for 3-Month Review] ➔ [Programme Closed]
```

1. **`Created`**: Initial training requisition approved or manual draft created.
2. **`Participants Imported`**: Participants added from Employee Directory.
3. **`Attendance In Progress`**: Active session created and QR code generated.
4. **`Training Completed`**: End date and time elapsed.
5. **`Evaluation Completed`**: Level 1 participant evaluation threshold reached.
6. **`Waiting for 3-Month Review`**: System enters a 90-day countdown before supervisor evaluation becomes due.
7. **`Programme Closed`**: Post-training evaluations completed and archived.

---

## 📁 3. Google Drive Workspace & Document Automation

When a training requisition is submitted, TrainHub automatically builds an organized Google Drive folder hierarchy:

```
ROOT_FOLDER_ID (TrainHub Workspaces Root)
│
├── Training Folder/
│   └── TRN-2026-001 Leadership Excellence Workshop/
│       ├── TRN-2026-001 Training Requisition Form     <-- Master AP-HRD-F01-01 Sheet
│       ├── Training Data                              <-- Isolated Sheet (Participants, Attendance, Evals)
│       └── Brochure/                                  <-- Uploaded Course Brochure / PDF attachments
│
├── Reports/                                           <-- Automated Master Analytics Exports
│   ├── Master_Training_Reports_2026.xlsx
│   ├── Annual_Training_Plan_2026.xlsx
│   └── Training_Hours_Report_2026.xlsx
│
└── Namelist/                                          <-- Master Directory Backups
```

### 3.1 `AP-HRD-F01-01` Form Automation
TrainHub makes a copy of the official template and dynamically injects data into specific cell ranges:
- **Training Details**:
  - `C5:I5`: Training Title
  - `C6:E6`: Course Fee (RM)
  - `G6:I6`: Training Date (`From` - `Until`)
  - `C7:E7`: Duration (`X day(s) (Y hours)`)
  - `G7:I7`: Training Venue
  - `C8:I8`: Training Provider / Trainer
  - `A11:I12`: Objectives & Reasons for Training
- **Participant Table (`A15:I38`)**: Auto-fills Employee No (`A:B`), Full Name (`C`), Department / Cost Centre (`D:G`), and Job Position (`H:I`) for up to 24 participants per form.
- **Digital Signatures & Approvals (`Rows 42:45`)**:
  - `A42:B45`: **Requested By** (Requester Name, ID, Position, Date)
  - `C42:C45`: **Verified by Head of Department** (Status, ID, Name, Position, Date)
  - `D42:E45`: **Approved by C-Suite** (Status, ID, Name, Position, Date)
  - `F42:G45`: **Approved by Head of HR (HOHR)** (Status, ID, Name, Position, Date)
  - `H42:I45`: **HR Department Acknowledgment** (Status, ID, Name, Position, Date)

---

## 📊 4. Database Schema & Data Dictionary

### 4.1 Master Database Spreadsheet (`SPREADSHEET_ID`)

#### Sheet: `Trainings`
| Column Name | Data Type | Description |
| :--- | :--- | :--- |
| `ID` | String | Unique Training identifier (`TRN-XXXXXX`) |
| `Code` | String | Formatted Training Code (e.g. `TRN-2026-001`) |
| `Name` | String | Title / Programme Name |
| `Category` | String | Category: *Behavioral Skills*, *Technical Skills*, *Compliance Training*, *Business Skills*, *Onboarding* |
| `Trainer` | String | Trainer / Instructor Name |
| `TrainingProvider` | String | Provider / Vendor Company Name |
| `Venue` | String | Physical Location or Virtual Platform |
| `StartDate` | String / Date | Start Date (`YYYY-MM-DD`) |
| `EndDate` | String / Date | End Date (`YYYY-MM-DD`) |
| `Duration` | Number | Total training duration in days |
| `TotalHours` | Number | Total training duration in hours |
| `CourseFee` | Number / String | Total training investment (RM) |
| `Department` | String | Target Department / Cost Centre |
| `Status` | String | Lifecycle Status (*Draft*, *Upcoming*, *Ongoing*, *Completed*, *Cancelled*, *Pending Revision*) |
| `Stage` | String | Sequential stage in the training lifecycle pipeline |
| `ApprovalStatus` | String | Requisition status (*Pending HOD Approval*, *Pending C-Suite Approval*, *Pending HOHR Approval*, *Approved*, *Rejected*, *Returned*) |
| `RequestedBy` | String | Requester Employee ID |
| `RequestedByName`| String | Requester Full Name |
| `RequestedByEmail`| String | Requester Company Email |
| `HOD` / `HODStatus`| String | Assigned HOD Name & Stage Status |
| `Csuite` / `CsuiteStatus` | String | Assigned C-Suite Name & Stage Status |
| `HOHR` / `HOHRStatus` | String | Assigned HOHR Name & Stage Status |
| `FolderID` | String | Google Drive workspace folder ID |
| `RequisitionFormFileID` | String | Google Drive file ID of the generated `AP-HRD-F01-01` sheet |
| `BrochureURL` | String | Direct URL to uploaded brochure/attachment |

---

### 4.2 Employee Master Database (`EMPLOYEE_SPREADSHEET_ID`)

- **`Employees`**: Complete company roster.
  - Columns: `ID`, `Name`, `Position Title` *(mapped system-wide as Job Title)*, `Cost Centre` *(mapped system-wide as Department)*, `Email`, `Phone`, `Status`.
- **`Cost Centre`**: Master list of active company Cost Centres / Departments.
- **`HOD email`**: Management routing matrix.
  - Columns: `Department` (or `Cost Centre`), `HODName`, `HODEmail`, `CsuiteName`, `CsuiteEmail`, `HohrName`, `HOHREmail`.
- **`For IT`**: Detailed reporting hierarchy mapping each employee to their direct supervisor/HOD.

---

### 4.3 Training Workspace Spreadsheet (`Training Data`)

Each training workspace possesses its own standalone `Training Data` spreadsheet:
- **`Participants`**: `ID`, `TrainingID`, `EmployeeID`, `EmployeeName`, `Department`, `Position`, `AddedAt`, `SupervisorID`, `SupervisorEmail`, `SupervisorName`.
- **`Sessions`**: `SessionID`, `TrainingID`, `SessionName`, `SessionDate`, `StartTime`, `EndTime`, `AttendanceURL`, `QRCodeURL`, `QRStatus`, `CreatedDate`.
- **`Attendance`**: `AttendanceID`, `SessionID`, `TrainingID`, `EmployeeNo`, `EmployeeName`, `Department`, `ScanTime`, `Status`, `TrainingCode`, `Day`, `Date`, `Hours`, `Remarks`, `EditedBy`, `EditedAt`.
- **`Evaluation`**: `ID`, `TrainingID`, `EmployeeID`, `EmployeeName`, `Q1` to `Q7`, `SectionB1` to `SectionB3`, `AvgScore`, `SubmittedAt`.
- **`Post Evaluation`**: `ID`, `TrainingID`, `EmployeeID`, `EvaluatorName`, `EvaluatorID`, `CompetencyBefore`, `CompetencyAfter`, `Improvement`, `CanApply`, `FurtherTraining`, `Comments`, `SubmittedAt`.
- **`Summary`**: Live spreadsheet formulas (`=COUNTA`, `=COUNTIF`, `=AVERAGE`) for instant session health metrics.

---

## 📈 5. Reporting, Analytics & Excel Export Engine

TrainHub includes a server-side report generator capable of live UI rendering and generating multi-tab Microsoft Excel (`.xlsx`) workbooks saved directly to Google Drive (`Reports/` folder):

| Report Name | Key Metrics & Data Columns Included |
| :--- | :--- |
| **Annual Training Plan (ATP)** | `No`, `Training Title`, `Training Category`, `TNA Source`, `Training Mode`, `Duration (hrs)`, `Trainer`, `Department`, `Position / Employee List (Hyperlinked)`, `Total Pax`, `Planned Date`, `Actual Date (From/To)`, `Status`, `Remarks` |
| **Training Hours Report** | Matrix of **Cost Centre vs. Month** (Jan–Dec) training hours delivered + Year-to-Date totals. |
| **Training Cost Breakdown** | Itemized costs per training: *Training Fees*, *Meal*, *Subsistence Allowance*, *Hotel Fees*, *Mileage Claim*, *Taxi Fees*, *Toll Fees*, *Flight*, *Total Cost*, *HRDF Grant Claimable*. |
| **Training Title Report** | Master training log by course name, trainer, hours, cost centre, and certificate expiry dates. |
| **Employee Training Record** | Individual employee transcript showing completed courses, cumulative training hours, and evaluation scores. |

---

## 📱 6. QR Code Check-In & Evaluation Engine

### 6.1 Attendance QR Code System
- **Dynamic QR Generation**: QR codes are generated with high error correction and embedded with the official company logo.
- **Expiry Protection**: Sessions automatically deactivate QR check-ins once the scheduled end date/time has passed.
- **Public Check-In Flow**:
  1. Participant scans QR code on their smartphone (redirects to `participant-portal?page=attendance&session=SESXXXX`).
  2. Participant enters their 5-digit Employee ID (e.g. `00123`).
  3. System validates employee enrollment against `TrainingParticipants` in real time, displays their verified name, and records attendance with server timestamp.
  4. Duplicate check-ins for the same session are blocked.

### 6.2 Level 1 & Level 3 Post-Evaluations
- **Level 1 (Reaction)**: Evaluates Course Content, Instructor Effectiveness, Facility/Platform Quality, Time Allocation, and Learning Objectives (Scale 1–5 + Qualitative text areas).
- **Level 3 (3-Month Behavior Evaluation)**:
  - 90 days after training completion, supervisors access `hod-portal?page=posteval` or receive direct links.
  - Supervisors evaluate competency levels before vs. after training, job applicability, and further training requirements.
  - Features a live countdown timer showing remaining months, days, and hours until the evaluation window unlocks.

---

## 🛠️ 7. Complete Step-by-Step Deployment Guide

### Prerequisites
1. **Google Account**: Google Workspace administrator or dedicated deployment account (e.g., `it@apollofood.com.my`).
2. **Google Apps Script API**: Enabled under [Google Apps Script User Settings](https://script.google.com/home/usersettings).
3. **Node.js & Clasp CLI** (Recommended for local command-line deployment):
   ```bash
   npm install -g @google/clasp
   clasp login
   ```

---

### Step 1: Prepare Master Spreadsheets & Drive Hierarchy

1. **Create Master Database Spreadsheet**:
   - Create a blank Google Spreadsheet named `TrainHub Master Database`.
   - Copy the Spreadsheet ID from the URL (`https://docs.google.com/spreadsheets/d/`**`SPREADSHEET_ID`**`/edit`).

2. **Create Employee Master Spreadsheet**:
   - Create a Google Spreadsheet named `TrainHub Employee Master Database`.
   - Create the following tabs:
     - **`Employees`**: Headers: `ID`, `Name`, `Position Title`, `Cost Centre`, `Email`, `Phone`, `Status`.
     - **`Cost Centre`**: Header: `Cost Centre` (list of official company departments).
     - **`HOD email`**: Headers: `Department`, `HODName`, `HODEmail`, `CsuiteName`, `CsuiteEmail`, `HohrName`, `HOHREmail`.
     - **`For IT`**: Headers: `ID`, `Name`, `HODName`, `Cost Centre`.
   - Copy the Spreadsheet ID from the URL (`EMPLOYEE_SPREADSHEET_ID`).

3. **Upload `AP-HRD-F01-01` Form Template**:
   - Upload the master Google Sheet template of `AP-HRD-F01-01` to Google Drive.
   - Copy its File ID (`TRAINING_REQUISITION_TEMPLATE_ID`).

4. **Create Root Drive Folder**:
   - Create a master Google Drive folder named `TrainHub Workspaces`.
   - Copy its Folder ID (`ROOT_FOLDER_ID`).
   - Create child folders inside it named `Training Folder`, `Reports`, and `Namelist`.

---

### Step 2: Deploy `admin-system` (Admin Web Application)

1. Open terminal and navigate to directory:
   ```bash
   cd "admin-system"
   clasp push
   ```
2. In the Apps Script Editor, go to **Project Settings** ⚙️ ➔ **Script Properties** and configure:

   | Key | Description | Example / Value |
   | :--- | :--- | :--- |
   | `SPREADSHEET_ID` | Master Database ID | `1a2b3c...` |
   | `EMPLOYEE_SPREADSHEET_ID` | Employee Master Database ID | `1x2y3z...` |
   | `TRAINING_REQUISITION_TEMPLATE_ID` | `AP-HRD-F01-01` Template Sheet ID | `1t2u3v...` |
   | `ROOT_FOLDER_ID` | Google Drive Root Workspace Folder ID | `1f2g3h...` |
   | `ALLOWED_DOMAIN` | Allowed company email domain | `apollofood.com.my` |
   | `ADMIN_EMAILS` | Authorized Admin Emails | `it@apollofood.com.my,arina.ismail@apollofood.com.my` |
   | `PUBLIC_PORTAL_URL` | Web App URL of `participant-portal` | *(Set in Step 6)* |
   | `HOD_PORTAL_URL` | Web App URL of `hod-portal` | *(Set in Step 6)* |

3. **Run Initializers**:
   - In the Apps Script editor toolbar, select `initDefaultScriptProperties` and click **Run**.
   - Select `setupSheets` and click **Run** (creates all necessary sheet tabs and headers in the Master Database).
4. **Deploy Web App**:
   - Click **Deploy** ➔ **New deployment** ➔ Select type: **Web app**.
   - Execute as: **Me** (`it@apollofood.com.my`)
   - Who has access: **Anyone within domain** (or *Anyone*)
   - Save the generated **Admin Web App URL**.

---

### Step 3: Deploy `employee-requisition` (Requisition Portal)

1. Navigate to directory:
   ```bash
   cd "employee-requisition"
   clasp push
   ```
2. Configure **Script Properties**:

   | Key | Description | Value |
   | :--- | :--- | :--- |
   | `SPREADSHEET_ID` | Master Database ID | `1a2b3c...` |
   | `EMPLOYEE_SPREADSHEET_ID` | Employee Master Database ID | `1x2y3z...` |
   | `TRAINING_REQUISITION_TEMPLATE_ID` | Form Template Sheet ID | `1t2u3v...` |
   | `ROOT_FOLDER_ID` | Master Drive Root Folder ID | `1f2g3h...` |
   | `HOD_PORTAL_URL` | Web App URL of `hod-portal` | *(Set in Step 6)* |
   | `ALLOWED_DOMAIN` | Company domain | `apollofood.com.my` |

3. **Deploy Web App**:
   - Execute as: **Me**
   - Access: **Anyone with company account**
   - Save the **Employee Requisition Web App URL**.

---

### Step 4: Deploy `hod-portal` (Managerial Review Portal)

1. Navigate to directory:
   ```bash
   cd "hod-portal"
   clasp push
   ```
2. Configure **Script Properties**:

   | Key | Description | Value |
   | :--- | :--- | :--- |
   | `SPREADSHEET_ID` | Master Database ID | `1a2b3c...` |
   | `EMPLOYEE_SPREADSHEET_ID` | Employee Master Database ID | `1x2y3z...` |
   | `EMPLOYEE_PORTAL_URL` | Web App URL of `employee-requisition` | `https://script.google.com/macros/s/.../exec` |
   | `ALLOWED_DOMAIN` | Company domain | `apollofood.com.my` |

3. **Deploy Web App**:
   - Execute as: **Me**
   - Access: **Anyone** *(Approver identity is strictly verified server-side via Google Workspace Session)*
   - Save the **HOD Portal Web App URL**.

---

### Step 5: Deploy `participant-portal` (Public Participant Portal)

1. Navigate to directory:
   ```bash
   cd "participant-portal"
   clasp push
   ```
2. Configure **Script Properties**:

   | Key | Description | Value |
   | :--- | :--- | :--- |
   | `SPREADSHEET_ID` | Master Database ID | `1a2b3c...` |
   | `APP_TITLE` | Display Title | `TrainHub — Participant Portal` |

3. **Deploy Web App**:
   - Execute as: **Me**
   - Access: **Anyone** *(Must be set to "Anyone" so participants scanning QR codes can check in without Google login)*
   - Save the **Public Participant Web App URL**.

---

### Step 6: Inter-Project Link Wireup & Property Sync

After obtaining all four Web App URLs, complete the circular link routing across project properties:

```
┌─────────────────────────┐        PUBLIC_PORTAL_URL        ┌─────────────────────────┐
│      admin-system       ├────────────────────────────────►│   participant-portal    │
└────────────┬────────────┘                                 └─────────────────────────┘
             │                                                           ▲
             │ HOD_PORTAL_URL                             HOD_PORTAL_URL │
             ▼                                                           │
┌─────────────────────────┐       EMPLOYEE_PORTAL_URL       ┌────────────┴────────────┐
│       hod-portal        │◄────────────────────────────────┤  employee-requisition   │
└─────────────────────────┘                                 └─────────────────────────┘
```

1. **In `admin-system`**:
   - Set `PUBLIC_PORTAL_URL` = Web App URL of `participant-portal`.
   - Set `HOD_PORTAL_URL` = Web App URL of `hod-portal`.
2. **In `employee-requisition`**:
   - Set `HOD_PORTAL_URL` = Web App URL of `hod-portal`.
3. **In `hod-portal`**:
   - Set `EMPLOYEE_PORTAL_URL` = Web App URL of `employee-requisition`.

> 💡 **Legacy Workspace Migration**: If migrating from previous test folders, run `migrateTrainingWorkspacesToConfiguredRoot` from `admin-system` to safely relocate all existing training workspaces into `ROOT_FOLDER_ID/Training Folder/`.

---

## 🔗 8. URL Routing & Deep Link Reference Table

| Portal & Action | URL Query Parameter Pattern | Target User / Access Policy |
| :--- | :--- | :--- |
| **Attendance QR Check-In** | `/exec?page=attendance&session=SES0001` | Public / Mobile Browser (No login required) |
| **Level 1 Training Evaluation** | `/exec?page=evaluation&id=TRN-1001` | Participants upon course completion |
| **Level 3 Supervisor Post-Review**| `/exec?page=post&id=TRN-1001&emp=00123` | Direct supervisor review link |
| **HOD Requisition Review** | `/exec?page=review&id=TRN-1001` | Authenticated HOD / C-Suite / HOHR |
| **HOD 3-Month Evaluation Hub** | `/exec?page=posteval` | Authenticated HOD (Filtered by Cost Centre) |
| **Employee "My Requests"** | `/exec?page=my_requests` | Authenticated Employee |

---

## 🛡️ 9. Security, Authentication & Access Control

1. **Server-Side Session Validation**:
   - Sensitive portals (`admin-system`, `hod-portal`, `employee-requisition`) resolve identity via `Session.getActiveUser().getEmail()` server-side. Client-side email tampering or profile spoofing is strictly prevented.
2. **Domain Enforcement**:
   - Operations verify that logged-in accounts belong to `@apollofood.com.my` (or the configured `ALLOWED_DOMAIN`).
3. **Stage-Aware Approver Guard**:
   - An approver cannot approve or reject a requisition unless that requisition is in their specific active stage (`HOD`, `CSuite`, or `HOHR`) and their email matches the directory mapping.
4. **Duplicate Protection & Tamper Resistance**:
   - All attendance check-ins, Level 1 evaluations, and Level 3 reviews execute server-side duplicate verification against the training's dedicated spreadsheet.
5. **No Google Login Barrier for Attendees**:
   - The Public Participant Portal executes as the service account (`Execute as: Me`), enabling blue-collar, field, or contract personnel to check in via QR code using any smartphone browser without requiring corporate Google credentials.

---

## 🧰 10. Clasp CLI Command Reference

Each subfolder contains its own `.clasp.json` linked to its Google Apps Script project:

```bash
# 1. Admin System
cd admin-system
clasp push

# 2. Employee Requisition Portal
cd ../employee-requisition
clasp push

# 3. HOD & Managerial Review Portal
cd ../hod-portal
clasp push

# 4. Public Participant Portal
cd ../participant-portal
clasp push
```

> ⚠️ **Applying Live Updates**:  
> Running `clasp push` updates the backend script files. To reflect updates on production Web App URLs, go to **Deploy** ➔ **Manage deployments** ➔ Edit active deployment ➔ Select **New version** ➔ **Deploy**.

---

## 🗂️ 11. Complete Repository File Structure

```
Job Training Management System/
├── README.md                                  <-- Master System Architecture & Deployment Guide
├── Job Training Form.md                       <-- Requirements Log & Change History
│
├── admin-system/                              <-- Admin Web Application
│   ├── .clasp.json                            <-- Clasp Project Config
│   ├── appsscript.json                        <-- Manifest File
│   ├── Code.js                                <-- Web App Router & doGet Engine
│   ├── Helper.js                              <-- Sheet Access, Script Properties & Setup Engine
│   ├── Training.js                            <-- Lifecycle Stage Engine & 3-Month Countdown
│   ├── TrainingService.js                     <-- Programme CRUD Services
│   ├── SessionService.js                      <-- Session Creator & Scheduling Logic
│   ├── AttendanceService.js                   <-- Attendance Records Engine
│   ├── Evaluation.js                          <-- Level 1 & Level 3 Evaluation Engine
│   ├── Employee.js                            <-- Employee Database Lookup & Utilities
│   ├── QRService.js                           <-- QR Code Generation Engine
│   ├── DriveManager.js                        <-- Workspace Folder & AP-HRD-F01-01 Generator
│   ├── Report.js                              <-- Excel/Sheets Export & Annual Training Plan Generator
│   ├── ValidationService.js                   <-- Server-Side Input Validator
│   ├── index.html                             <-- Main Admin Container UI
│   ├── dashboard.html                         <-- Admin Dashboard UI
│   ├── training.html                          <-- Training Management UI
│   ├── session.html                           <-- Session Scheduling UI
│   ├── attendance.html                        <-- Attendance Log UI
│   ├── evaluation.html                        <-- Evaluation Summary UI
│   ├── employee.html                          <-- Employee Lookup UI
│   ├── report.html                            <-- Reports & Annual Training Plan Export UI
│   ├── settings.html                          <-- System Settings UI
│   ├── sidebar.html                           <-- Navigation Sidebar UI
│   ├── style.html                             <-- CSS Design System
│   └── script.html                            <-- Frontend JS Controllers
│
├── employee-requisition/                      <-- Employee Requisition Portal
│   ├── .clasp.json                            <-- Clasp Project Config
│   ├── appsscript.json                        <-- Manifest File
│   ├── README.md                              <-- Portal Specific Setup Guide
│   ├── Code.gs                                <-- Server Routing & Multi-Tier Approval Engine
│   ├── Helper.gs                              <-- Employee DB Connection & Requisition Logic
│   ├── ValidationService.gs                   <-- Form & Email Validator
│   ├── Requisition.html                       <-- 15-Field Training Request Form UI
│   ├── Success.html                           <-- Submission Confirmation Page
│   ├── Error.html                             <-- Error Boundary Page
│   ├── style.html                             <-- Portal CSS Styling
│   └── script.html                            <-- Frontend JS Form Handler
│
├── hod-portal/                                <-- HOD & Managerial Review Portal
│   ├── .clasp.json                            <-- Clasp Project Config
│   ├── appsscript.json                        <-- Manifest File
│   ├── README.md                              <-- Portal Specific Setup Guide
│   ├── Code.gs                                <-- Approval Routing & Post Eval Engine
│   ├── Helper.gs                              <-- Spreadsheet & HOD Cost Centre Service
│   ├── ValidationService.gs                   <-- Managerial Access Validator
│   ├── HodReview.html                         <-- Digital Stamp Requisition Review UI
│   ├── HodPostEvaluation.html                 <-- 3-Month Post Evaluation UI (Cost Centre Filtered)
│   ├── Success.html                           <-- Approval Confirmation UI
│   ├── Error.html                             <-- Access Denied / Error UI
│   ├── style.html                             <-- HOD Portal CSS Styling
│   └── script.html                            <-- Frontend Interactions JS
│
└── participant-portal/                        <-- Public Participant Portal
    ├── .clasp.json                            <-- Clasp Project Config
    ├── appsscript.json                        <-- Manifest File
    ├── README.md                              <-- Participant Portal Setup Guide
    ├── Code.gs                                <-- Public Web App Router
    ├── Helper.gs                              <-- Database Connection Service
    ├── AttendanceService.gs                   <-- QR Check-In Backend Handler
    ├── Evaluation.gs                          <-- Level 1 & Level 3 Submission Handler
    ├── ValidationService.gs                   <-- Server-Side Access & Duplicate Validator
    ├── Attendance.html                        <-- Public QR Attendance Check-In UI
    ├── TrainingEvaluation.html                <-- Level 1 Evaluation UI
    ├── PostEvaluation.html                    <-- 3-Month Supervisor Review UI
    ├── Success.html                           <-- Submission Success Page
    ├── Error.html                             <-- Error Page UI
    ├── style.html                             <-- Participant Portal Styling
    └── script.html                            <-- Form Interactive JS
```

---

## ❓ 12. Troubleshooting & Frequently Asked Questions (FAQ)

### Q1: The Admin System dashboard displays an old database or no records.
- **Cause**: Different `SPREADSHEET_ID` values across projects or cached Apps Script session data.
- **Fix**: Verify that `SPREADSHEET_ID` in `admin-system` Script Properties matches the exact ID of your active Master Database spreadsheet. Then run `initDefaultScriptProperties` and reload.

### Q2: Participants get "Check-in rejected: Not enrolled" during QR check-in.
- **Cause**: The participant's Employee ID is not present in the training's `Participants` sheet tab.
- **Fix**: Ensure that the employee was included during requisition submission or added via the Admin System's *Training Programme ➔ Participants* manager before scanning.

### Q3: How do I change the sender address for automated notification emails?
- **Cause**: Google Apps Script `MailApp.sendEmail` sends emails using the account that authorized the Web App deployment.
- **Fix**: Deploy all four Web Apps using the designated central system account (`it@apollofood.com.my`) with `Execute as: Me`.

### Q4: An approver sees "Access Denied: Unauthorized Account" on the HOD Portal.
- **Cause**: The approver's logged-in Google email does not match any entry in the `HOD email` or `For IT` tabs in `EMPLOYEE_SPREADSHEET_ID`.
- **Fix**: Ensure the manager's email address is listed in the `HOD email` tab under `HODEmail`, `CsuiteEmail`, or `HohrEmail`.

---

## 📜 Maintenance & Operational Notes
- **Author**: IT Department / TrainHub Core Engineering Team
- **Version**: `2.0.0 (Enterprise Release)`
- **Compliance Standard**: ISO 9001 / HRD Corp Compliant Training Formats (`AP-HRD-F01-01`)
