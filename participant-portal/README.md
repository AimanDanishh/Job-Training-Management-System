# Public Participant Portal — Google Apps Script Setup Guide

This Google Apps Script project serves as the **Public Participant Portal** for the Job Training Management System. It allows participants to check in via QR code, submit training evaluations, and allows supervisors to submit 3-month post-training evaluations without requiring Google sign-in.

---

## 📌 Features

- **Attendance QR Check-In**: Participants scan session QR codes and enter their Employee ID.
- **Training Evaluation**: 7-statement Likert rating scale + open feedback text areas.
- **6-Month Post Evaluation**: Supervisor competency assessment before/after training.
- **Server-Side Validation**: Validates employee existence, enrollment, active session status, and prevents duplicate submissions.
- **Shared Spreadsheet Database**: Reads and writes to the main Admin Google Spreadsheet without schema changes.

---

## 🛠️ Step 1: Apps Script Project Creation

1. Open [Google Apps Script](https://script.google.com/).
2. Click **+ New project**.
3. Rename the project to `Public Participant Portal`.

---

## ⚙️ Step 2: Configure Script Properties (Project Settings)

The Public Portal connects to your existing Google Spreadsheet database via **Script Properties**.

1. In the Apps Script editor sidebar, click **Project Settings** ⚙️ (gear icon).
2. Scroll down to **Script Properties**.
3. Click **Add script property** and add the following key-value pairs:

| Property Key | Example Value | Description |
| :--- | :--- | :--- |
| `SPREADSHEET_ID` | `1a2b3c4d5e6f7g8h9i0j...` | **Required.** The ID of your existing Master Google Spreadsheet (from its URL). |
| `APP_TITLE` | `TrainHub — Participant Portal` | Optional. Display title for browser tabs. |

> 💡 **Where to find `SPREADSHEET_ID`?**  
> Open your existing Google Spreadsheet. Copy the long string of letters and numbers from the URL between `/d/` and `/edit`:  
> `https://docs.google.com/spreadsheets/d/`**`1a2b3c4d5e6f7g8h9i0j`**`/edit`

---

## 🚀 Step 3: Push Code via Clasp

If you are using the `clasp` command line tool to deploy:

1. Copy your new Apps Script Project ID from **Project Settings** > **IDs** > **Script ID**.
2. Open `.clasp.json` in the `participant-portal/` directory:
   ```json
   {
     "scriptId": "YOUR_ACTUAL_SCRIPT_ID_HERE"
   }
   ```
3. In your terminal, run:
   ```bash
   cd "c:\Users\Intern IT\Documents\Job Training System\participant-portal"
   clasp push
   ```

---

## 🌐 Step 4: Web App Deployment Settings

To make the portal accessible to participants and supervisors without Google sign-in:

1. In the Apps Script Editor, click **Deploy** > **New deployment**.
2. Click the gear icon ⚙️ next to *Select type* and choose **Web app**.
3. Configure the following fields:
   - **Description**: `Public Participant Portal v1.0`
   - **Execute as**: `Me` (*Your Google Account*)
   - **Who has access**: `Anyone`
4. Click **Deploy**.
5. Copy the generated **Web App URL** (e.g., `https://script.google.com/macros/s/AKfycb.../exec`).

---

## 🔗 URL Parameters & Links

The Public Portal routes requests based on URL query parameters:

### 1. Attendance Check-In (QR Code Landing)
```
https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec?page=attendance&session=SES0001
```

### 2. Training Evaluation Page
```
https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec?page=evaluation&id=TRN-1001
```

### 3. 6-Month Post Evaluation Page (Supervisor Link)
```
https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec?page=post&id=TRN-1001&emp=EMP-1001
```

---

## 🛡️ Security & Validation Summary

All data submitted from HTML pages is validated on the Apps Script backend ([`ValidationService.gs`](file:///c:/Users/Intern%20IT/Documents/Job%20Training%20System/participant-portal/ValidationService.gs)):

1. **Training Existence**: Rejects requests if `TrainingID` / `SessionID` does not exist.
2. **Employee Registration**: Rejects check-in if Employee ID is missing from `Employees`.
3. **Participant Enrollment**: Rejects submission if Employee ID is not registered in `TrainingParticipants`.
4. **Duplicate Protection**: Blocks multiple attendance or evaluation submissions from the same employee.
5. **Session Expiry**: Blocks check-in if the session status is set to `Expired` or `Inactive`.
