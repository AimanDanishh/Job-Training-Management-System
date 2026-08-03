# TrainHub — Job Training & Management System

> A comprehensive Google Apps Script (GAS) Web Application for employee training management, attendance tracking, evaluations, requisition form automation, and reporting backed by Google Sheets.

---

## 📋 Table of Contents
- [Overview](#-overview)
- [Key Features](#-key-features)
- [System Architecture](#-system-architecture)
- [Prerequisites](#-prerequisites)
- [Step-by-Step Setup & Running Locally](#-step-by-step-setup--running-locally)
  - [1. Google Sheet Setup](#1-google-sheet-setup)
  - [2. Clone or Sync Code via Clasp](#2-clone-or-sync-code-via-clasp)
  - [3. Configure Script Properties](#3-configure-script-properties)
  - [4. Initialize Database Sheets](#4-initialize-database-sheets)
- [Google Drive Workspace & Requisition Form Automation](#-google-drive-workspace--requisition-form-automation)
- [Deployment Guide](#-deployment-guide)
  - [Deploying as Web App](#deploying-as-web-app)
  - [Updating an Existing Deployment](#updating-an-existing-deployment)
- [File Structure](#-file-structure)
- [Troubleshooting & FAQs](#-troubleshooting--faqs)

---

## 🌟 Overview

**TrainHub** is designed for organizations to streamline end-to-end training administration:
- **Employee Directory**: Centralized management of employee profiles, departments, positions, NRICs, and contact details.
- **Training Programmes & Lifecycle**: Manage training courses through structured lifecycle stages (Created, Participants Imported, Attendance In Progress, Training Completed, Evaluation Completed, Waiting for 6-Month Review, Programme Closed).
- **Attendance Logging & QR Integration**: Track per-day check-in/check-out times, calculate training hours, edit remarks, and generate QR codes for fast mobile check-ins.
- **Automated Requisition Forms**: Automatically duplicate and populate master Google Sheets requisition forms (`AP-HRD-F01-00`) for each training program and keep participant rosters synchronized.
- **Training & Post-Training Evaluations**: Collect instant participant feedback and conduct post-training competency evaluations.
- **Analytics & PDF Export**: Real-time management dashboards, department metrics, summary reports, and PDF export support.

---

## 🚀 Key Features

| Feature | Description |
| :--- | :--- |
| **Employee Management** | Full CRUD for staff records with department filtering and active/inactive status tracking. |
| **Training Lifecycle Control** | End-to-end tracking of training programmes from requisition through 6-month evaluation. |
| **Automated Drive Workspaces** | Automatically generates organized Drive folders and subfolders (`Attendance`, `Evaluation`, `Certificates`, `Materials`, `Photos`, `Reports`, `Trainer Notes`) for every new training. |
| **Form Population Engine** | Auto-fills requisition form headers and updates participant lists directly inside Google Sheets. |
| **QR Code Attendance** | Integrated QR generator for rapid mobile attendance logging and check-in/out timestamping. |
| **Evaluations & Competency** | Form handling for immediate training evaluation and scheduled 6-month post-training effectiveness reviews. |
| **Executive Reports & PDF Export** | Multi-dimensional analytics, department participation breakdowns, and downloadable PDF reports. |

---

## 🏗️ System Architecture

- **Frontend**: HTML5, Vanilla JavaScript, CSS3 with responsive UI layouts, custom fonts (`aphakind`), and smooth animations.
- **Backend Router**: Google Apps Script (`Code.gs`) serving HTML templates via `HtmlService`.
- **Drive & Form Engine**: `DriveManager.gs` managing folder hierarchies, template copies, and spreadsheet form data injection.
- **Database Layer**: Google Sheets operating as a structured tabular database (`Employees`, `Trainings`, `Attendance`, `TrainingEval`, `PostEval`).
- **Development Tooling**: Google Clasp (`@google/clasp`) for local sync, watch mode, and automated deployment.

---

## ⚡ Prerequisites

Before setting up TrainHub, ensure you have:
1. **Google Account**: Access to Google Drive, Google Sheets, and Google Apps Script.
2. **Node.js & npm** *(Optional, recommended for local development)*: [Download Node.js](https://nodejs.org/).
3. **Clasp CLI** *(Optional, for syncing code locally)*: Install `@google/clasp` globally via npm.

---

## 🚀 Step-by-Step Setup & Running Locally

### 1. Google Sheet Setup
1. Open [Google Sheets](https://sheets.google.com) and create a **Blank Spreadsheet**.
2. Name your spreadsheet (e.g., `TrainHub Database`).
3. Copy the **Spreadsheet ID** from the browser URL:
   ```text
   https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID_HERE/edit
   ```
4. Open `Helper.gs` and set the `SPREADSHEET_ID` variable or configure it in Script Properties:
   ```javascript
   const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';
   ```
   *(Note: If you bind this Apps Script directly inside the Google Sheet via **Extensions > Apps Script**, you can leave `SPREADSHEET_ID = ''`)*.

---

### 2. Clone or Sync Code via Clasp

#### Option A: Using Clasp (Recommended for Developers)
1. Install Clasp globally:
   ```bash
   npm install -g @google/clasp
   ```
2. Enable the **Google Apps Script API** in your Google account settings at:
   👉 [https://script.google.com/home/usersettings](https://script.google.com/home/usersettings)

3. Log in to Clasp:
   ```bash
   clasp login
   ```
4. In your project root directory, verify or create `.clasp.json`:
   ```json
   {
     "scriptId": "YOUR_APPS_SCRIPT_ID",
     "rootDir": ""
   }
   ```
5. Push local code to Google Apps Script:
   ```bash
   clasp push
   ```

#### Option B: Manual Copy-Paste (Alternative)
1. Open [Google Apps Script Editor](https://script.google.com/).
2. Create a new project.
3. Create files matching the project structure (`Code.gs`, `Helper.gs`, `Employee.gs`, `DriveManager.gs`, `index.html`, etc.) and paste their respective code.

---

### 💡 Local Development & Live Testing Workflow

Since Google Apps Script applications run on Google's cloud infrastructure, the recommended local development workflow uses **Clasp Auto-Sync** + **Development Test URL (`/dev`)**:

1. **Start Live Auto-Sync** in your terminal:
   ```bash
   clasp push --watch
   ```
   *(This watches your local files and uploads changes automatically whenever you save).*

2. **Open the Test URL (`/dev`)**:
   - In Google Apps Script Editor, click **Deploy** > **Test deployments**.
   - Copy the Web App test URL (ends with `/dev`).
   - Open it in your browser.

3. **Develop & Refresh**:
   - Edit your HTML, CSS, JavaScript, or `.gs` code locally in your IDE.
   - `clasp` automatically pushes changes to Apps Script.
   - Refresh your browser tab to test live changes instantly!

---

### 3. Configure Script Properties (Apps Script Project Settings)

System IDs, credentials, sheet names, and application parameters are configured in Google Apps Script **Project Settings -> Script Properties**:

1. In the Google Apps Script editor, click **Project Settings** (⚙️ gear icon on the left sidebar).
2. Scroll down to **Script Properties**.
3. Click **Add script property** (or run `initDefaultScriptProperties()` from `Helper.gs` to set defaults automatically).

#### 🛠️ Available Script Properties:

| Property Key | Description | Default / Example |
| :--- | :--- | :--- |
| `ALLOWED_DOMAIN` | Restrict login to company Google Workspace domain | `company.com` |
| `ADMIN_EMAILS` | Comma-separated list of authorized Admin Google emails | `admin@company.com, manager@company.com` |
| `ADMIN_USER` | Fallback administrative username | `admin` |
| `ADMIN_PASS` | Fallback administrative password | `admin123` |
| `APP_TITLE` | Application title displayed in browser tab | `TrainHub — Training Management System` |
| `SPREADSHEET_ID` | ID of your Google Spreadsheet database | `1Y-pvLmTOllx84Vw...` *(or blank if bound)* |
| `ROOT_FOLDER_ID` | Root Google Drive parent folder ID for training workspaces | `1abc...` *(defaults to `Job Training System/Training`)* |
| `TEMPLATE_FOLDER_ID` | Master Template folder ID | `1xyz...` |
| `ATTENDANCE_TEMPLATE_ID` | Template file ID for Attendance Records | `1att...` |
| `EVALUATION_TEMPLATE_ID` | Template file ID for Evaluation Forms | `1eval...` |
| `CERTIFICATE_TEMPLATE_ID` | Template file ID for Training Certificates | `1cert...` |
| `REPORT_TEMPLATE_ID` | Template file ID for Summary Reports | `1rep...` |
| `TRAINING_REQUISITION_TEMPLATE_ID` | Google Sheets ID of master `AP-HRD-F01-00` requisition form. Automatically duplicated and filled for each training. | `1form...` |
| `SHEET_EMPLOYEES` | Custom tab name for Employees | `Employees` |
| `SHEET_TRAININGS` | Custom tab name for Trainings | `Trainings` |
| `SHEET_ATTENDANCE` | Custom tab name for Attendance | `Attendance` |
| `SHEET_TRAINING_EVAL` | Custom tab name for Training Evaluations | `TrainingEval` |
| `SHEET_POST_EVAL` | Custom tab name for Post-Training Evaluations | `PostEval` |

---

### 📂 Google Drive Workspace & Requisition Form Automation

When a new training programme is created, TrainHub automatically generates a dedicated Google Drive folder structure and initializes template files:

#### 1. Folder Hierarchy
```text
Job Training System/
└── Training/
    └── TR-2026-001 Safety Induction/
        ├── Attendance/
        ├── Evaluation/
        ├── Certificates/
        ├── Materials/
        ├── Photos/
        ├── Reports/
        └── Trainer Notes/
```

#### 2. Requisition Form Population (`AP-HRD-F01-00`)
- Setting `TRAINING_REQUISITION_TEMPLATE_ID` automatically copies the master `AP-HRD-F01-00 (Training Requisition Form)` into the programme workspace.
- Basic fields (Course Name, Course Fee, Training Dates, Duration/Hours, Venue, Trainer, Objectives) are populated automatically.
- As attendance records are saved, `syncTrainingRequisitionParticipants()` automatically updates the participant grid (Employee ID, Name, Department, NRIC, Position) in the requisition sheet.

---

### 4. Initialize Database Sheets
To automatically create all required sheet tabs with pre-styled headers (`Employees`, `Trainings`, `Attendance`, `TrainingEval`, `PostEval`):

1. In the Apps Script Editor, open `Helper.gs`.
2. Select `setupSheets` from the function dropdown menu at the top.
3. Click **Run**.
4. Grant the required permissions when prompted by Google OAuth.
5. Check your Google Sheet — all required tabs will be generated automatically!

---

## 🌐 Deployment Guide

### Deploying as Web App

1. In the Google Apps Script Editor, click **Deploy** > **New deployment** (top right corner).
2. Click the gear icon (**Select type**) and choose **Web app**.
3. Fill in the deployment configuration:
   - **Description**: `TrainHub Web App Production`
   - **Execute as**: `Me (your email address)`
   - **Who has access**: `Anyone` (or `Anyone within your organization`)
4. Click **Deploy**.
5. Click **Authorize access** and choose your Google Account.
   - *If prompted with "Google hasn't verified this app", click **Advanced** > **Go to TrainHub (unsafe)** to grant permissions.*
6. Copy the generated **Web App URL** (ends with `/exec`). This is the shareable link for accessing TrainHub!

---

### Updating an Existing Deployment

Whenever you push code changes via `clasp push` or modify code in the Apps Script editor:

1. Click **Deploy** > **Manage deployments**.
2. Select your active Web App deployment from the left panel.
3. Click the **Edit** (pencil icon) at the top right.
4. Set **Version** to **New version**.
5. Click **Deploy**.

---

## 📂 File Structure

```text
├── Code.gs             # Main entry point, page routing (doGet), template inclusion & user auth
├── DriveManager.gs     # Drive workspace creation, template copying & requisition form population
├── Helper.gs           # Database connector, sheet initialization, date/JSON formatters
├── Employee.gs         # Employee CRUD logic & sheet handlers
├── Training.gs         # Training programme CRUD & lifecycle stage tracking
├── Attendance.gs       # Attendance record handling, hour calculations & day grouping
├── Evaluation.gs       # Training & 6-month post-training evaluation handling
├── Report.gs           # Report generation, analytics data & spreadsheet/PDF exports
├── Auth.gs             # Authentication and authorization helpers
├── index.html          # Main application layout, sidebar & view router
├── dashboard.html      # Overview dashboard UI with metrics & charts
├── employee.html       # Employee directory UI
├── training.html       # Training course creation & lifecycle management UI
├── attendance.html     # Attendance tracking & QR check-in UI
├── evaluation.html     # Participant feedback & post-training evaluation UI
├── report.html         # Analytics reports & export preview UI
├── style.html          # Modular CSS styles (Glassmorphism, animations, UI tokens, aphakind font)
├── script.html         # Client-side JavaScript controllers & API integrations
├── Job Training Form.md # Reference notes on form formatting & features
├── appsscript.json     # Apps Script runtime configuration (V8, Timezone)
└── .clasp.json         # Clasp CLI project configuration
```

---

## 🔍 Troubleshooting & FAQs

#### Q1: "Error loading page: Script function not found"
- **Solution**: Ensure all `.gs` and `.html` files are pushed to Google Apps Script using `clasp push`.

#### Q2: Spreadsheet data is not saving
- **Solution**: Verify `SPREADSHEET_ID` in Script Properties or `Helper.gs`. Ensure the spreadsheet is shared with write permissions for the Google account executing the script.

#### Q3: Requisition forms are not generated in Drive
- **Solution**: Ensure `TRAINING_REQUISITION_TEMPLATE_ID` is set in Script Properties pointing to a valid Google Sheets file format (not `.xlsx`).

#### Q4: Changes to HTML/CSS/JS are not reflecting on the Web App URL
- **Solution**: Ensure you published a **New version** under **Deploy > Manage deployments**. The test URL (ending in `/dev`) updates immediately, while the production URL (ending in `/exec`) requires creating a new deployment version.
