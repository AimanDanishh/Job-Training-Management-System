**Job Training Form**

aphakind  - font

monkey qr - qr generator



edit, remark attendance - hours



print session:

download attendance result, training form - pdf



**Improvement:**

1. select employee for each training (each training have different participant)

~~2. remove add employee button (both UI and backend)~~

~~3. quick action button does not directly work, just redirect to pages~~

~~4. ONLY Dashboard, Training Programmes, Attendance, Evaluation, Reports and Logout buttons on sidebar~~

~~5. The clicked sidebar, highlight the same button, need to highlight when clicked~~

~~6. Remove Employees tab completely, remove add, edit, delete functions~~

~~7. Make the view Training full instead of pop up style~~

~~8. Categorize the department by Cost Centre~~

~~9. Training UI list the participant, click attendance to directly view attendance instead of directed to attendance tab and need to select training~~

~~10. Just enter employee ID for attendance, then display name~~

~~11. Separate the App Script for attendance (Security concern)~~

~~1. Add delete QR option or edit QR~~

~~2. Auto deactivate QR when exceeding time~~

~~3. Simplify the link to copy link button~~

~~6. Bulk add participant, paste text, or docs or sheets or excel~~

~~7. Auto change programme status when arriving the date or pass the date~~

~~11. Auto deactivate the QR when pass the date~~

~~5. Change view attendance directly to view attendance (manual)~~

~~12. Enhance the UI so that easy to use or navigate~~



Additional Improvement:

4\. Form programme start and end date is not stored, need to enter again when edit, it displays dd/mm/yyyy

8\. Auto add/remove participant list into the Training Requisition Form when add/remove the participant from the training form

9\. Add generate QR and link for evaluation form (participant \& post training for supervisor)

10\. Fix the QR Sessions UI (link overflow) make it consistent like the training programme UI

13\. For attendance, add edit button for each session and save button, make the remark box bigger

14\. Enhance the manual attendance, drop down to select session (display session name and date only)

15\. For each training, auto generate the evaluation form QR, show the QR in the Evaluation tab under training and under training programme, for the participant to fill in, when participant enter the employee ID, their name is auto display then they can fill the form

16\. For 6 month evaluation form, each training will list the participant, admin can select the participant to be email to their supervisor manually (supervisor need to enter their employee ID, display supervisor name then, the form will have the participant name to be evaluate personally), then the selected participant will have the status like evaluated, pending or no action yet

17\. Display real admin acc instead of dummy

18\. Add loading screen or animation when the loading is slow

19\. change reports to hyperlink to the Training Requisition Form



**Workflow**

1. Admin create new Job Training, enter details
2. Admin add participant to the training (by name or id)
3. Admin select create QR to automatically generate QR code attendance for participants, and evaluation form
4. Admin can check participant attendance and add remark to it
5. Admin can check Training Evaluation and 6-Month Post Evaluation
6. Admin select some participant (bulk) and email the post evaluation to supervisor (manual type the email)
7. For Reports, admin can select format, PDF or Excel to export



**Improvement 3**

1. Make Evaluation form is required for scale 1-5
2. Remove the attendance records table
3. Add edit and save button besides the participant to edit and save the remarks/status for each participant
4. Report change to hyperlink to Training Requisition Form
5. Auto fill the employee no, name, department and job title in the Training Requisition Form
6. Create new session should be in the Training Programme tab, no need separate anymore, just directly show new session form
7. Update the dashboard



**Improvement 4**

Overall UI

1. When create the session attendance, auto change the lifecycle stage to attendance in progress, then when the training is passed the date, change to training completed, (automate lifecycle stage)
2. Show indicator or noti for training that have reach 6 months
3. Auto fill the employee no, name, department and job title in the Training Requisition Form (the template)
4. Allow the evaluation to be filled once for each participant
5. Make the evaluation form required for the question that have scale 1 until 5
6. Remove submit evaluation button on the admin system
7. Remove unnecessary folder from the drive, certificates, evaluation, materials, photos, reports, trainer notes
8. Add attendance sheet to each folder training, 1 attendance sheet per training, not dump everything in the database including the training evaluation and post evaluation



1. Participant added to the backend but not frontend
2. Bulk add participant not working, it should retrieve the participant from the employee database
3. The database still have the other tabs besides the Training tab
4. Use single sheets for attendance, post evaluation, training evaluation, participant, training session (just different tab)
5. Training Requisition From sheets keep it separated (individual sheets)
6. Fix all the routing and functions to the new structure



1. Employee fill \& submit the Training Requisition Form
2. System send email to HOD/Manager (the system link)
3. HOD/Manager approve/reject the request through the system
4. System marks the form as approve/reject digitally with timestamp
5. Admin can view the request status, email



**New Backend Structure**

1. All training data stored in Training Data sheets only, contains participants, sessions, attendance, evaluation and post evaluation
2. The main sheets (SPREADSHEET\_ID) contains only the training information (please update the table header to latest database structure and remove the other sheets tab inside this sheets)
3. Ensure all the frontend is connected to the latest backend structure



**Sheets table update**

1. By referring to the EMPLOYEE\_SPREADSHEET\_ID, do not use Company as Department, please use Cost Centre as Department and Position Title as Job Title (update this across all the system and remove the old one)
2. Remove all the dummy data when first setup the Sheets, leave it empty if no data inserted
3. Remove the lifecycle stage in the New Training Form
4. Change the Department/ Cost Centre drop down in New Training Form to real Cost Centre from the EMPLOYEE\_SPREADSHEET\_ID under Cost Centre tab
5. Change the Training Category drop down to Behavioral Skills (Soft Skills + Leadership + Customer Service | Technical Skills (Technical + Quality + SHE + Digital) | Compliance Training (Compliance + Safety) | Business Skills (Finance + Sales \& Marketing + Project Management) | Onboarding
6. The programme lifecycle stage is not aligned, ensure each step is done before updating the lifecycle stage, after participant added then participant imported, then after session created then attendance is in progress, when pass the date then the training is completed and so on
7. Update the UI from 6-Month to 3-Month Post Evaluation with countdown by day and if less then 1 day shown countdown by hours
8. The course fee in the UI display the date instead of Course Fee
9. When add bulk participant, please get the employee details from the EMPLOYEE\_SPREADSHEET\_ID only, not adding directly from the pasted text



**Training Form Requirement (The Training Request Form)**

1. Employee details (user just need to enter their Employee ID), the detail will be use to fill the REQUEST BY column in the Training Requisition Form in the template (employee no., Name, Job position, date)
2. Training Title
3. Training Fee
4. Training Date (From \& Until)
5. Training Duration (hrs, 0.5 accuracy)
6. Training Venue
7. Department (Cost Centre) (All department or multiple selection)
8. Trainer's name (if applicable)
9. Training Provider (Text box) (if applicable)
10. Training Type (In-House Training | Public)
11. Certificate Expiry Date (if applicable)
12. Total Pax (no. of participant)
13. Participant list (can search and select or bulk add, from EMPLOYEE\_SPREADSHEET\_ID only)
14. Reason for training (optional)
15. Brochure upload (to support the application) (optional)



**Training Requisition Form Flow (After request)**

1. HOD approve, reject, return (remarks)
2. Fill the Verified By Head of Department column, (request status, employee no., name, job position, date)
3. Approve By (c-suite \& HOHR) (after c-suite approve, HOHR approve)
4. Acknowledge by Arina



**Notes**

1. The system will use one email for all the process of emailing or automation (not use the user email)
2. All request need to be approved by HOD first (depends on the employee) (email the HOD using the email from HOD email tab in EMPLOYEE\_SPREADSHEET\_ID
3. If the Requester name == HOD name, directly fill request and verify
4. If the HOD name == Csuite name, directly verify \& approve both
5. If the HOD name == Csuite name == HOHR name, directly verify \& approve both (Csuite \& HOHR)
6. Lastly after all HOD, Csuite and HOHR approve, notify Arina (arina.ismail@apollofood.com.my) and allow the admin to create session \& everything



**Improvement/Problem (Training Request Form)**

1. Get the user email through what email they use to access the website
2. Save the email to database so that they could get notify if the training status have change (HOD approved, etc.)



**Improvement/Problem (HOD system)**

1. Work out all the UI again, make sure to use real data from the EMPLOYEE\_SPREADSHEET\_ID, no dummy, no fake approve, all real data only



**Improvement/Problem (Admin system)**

1. Training keeps loading forever, unable to see "No training list" or create new training



**Latest Database Update/Improvement**

1. For each training request, it should automatically create the Training Requisition Form and autofill
2. After each approval HOD, Csuite and HOHR it should also update the Training Requisition Form
3. The details in the request signature must be: (employee no., Name, Job position, date)
4. The details in the signature must be: (request status, employee no., name, job position, date)



**JOB TRAINING REQUISITION FORM**

1. Training Title: C5 until I5

2\. Course Fee: C6 until E6

3\. Date: G6 until I6

4\. Duration: C7 until E7

5\. Venue: G7 until I7

6\. Training Provider: C8 until I8

7\. Reasons for Training: A11 \& A12 until I11 and I12

8\. List of participant:

9\. Employee No: A15 \& B15 until A39 and B39

10\. Name: C15 until C39

11\. Department: D,E,F,G 15 until D,E,F,G 39

12\. Job Title: H15 \& I15 until H39 and I39

13\. Request By: A,B 42 until 45

14\. Verified by HOD: C42 until C45

15\. Approved By:

16\. Csuite: D,E 42 until 45

17\. HOHR: F, G 42 until 45

18\. HR Department (Arina): H, I 42 until 45



10/8/2026

**TO DO (HOD portal):**

1. Auto email for HOD portal
2. Link real HOD details to HOD portal
3. Approve or Reject or Return (with remarks) option only
4. If status == Reject or Return, email the requester to notify the status
5. If status == Approve, email the requester the status and email the next approval, for example (C-Suite), then after C-Suite approval, HOHR approval
6. For now just save all the email to draft (development mode)
7. In the signature part in the Training Requisition Form, the approval skip the Head of Department signature and directly to C-Suite, please fix it and fill the information correctly: Status: Employee No: Name: Job Position: Date:



1. What if there is many requester, can HOD see all request under their supervision (the employee under them only)
2. When log in using their email, retrieve the HOD/employee information from the database and display the employee/HOD information, i dont want anything like this:
3. HOD Employee ID
4. EMP-IT.INTERN
5. HOD Name
6. It Intern
7. HOD Cost Centre
8. All Departments / Cost Centre



**Improvement:**

1. If the number participant > 24, add new form and add the remaining participant

2\. Just notify the HR Department (arina) via email but auto sign it in the Training Requisition Form when she open the training in admin system



**11/8/2026**

**Improvement (Admin System):**

1. Admin Dashboard: add training status (HOD approval, Csuite, HOHR)

2\. Report

3\. Annual Training Plan

4\. Have filter selection first, for example (by Training Title, by month, by year), then generate the excel/sheets (export)



**Report Requirement:**

1. Training Hours (by year) (Cost Centre vs Month)

Header:

A1-A2: (Cost Centre/Month)

B-M: row 1 (Training Hours); row 2 (Jan-26, Feb-26, ...)

N1-N2: Total 2026

O1-O2: Total Training Hours 2026



2\. Training Cost (by year) (leave the cost empty, header \& Tranining Title, Date, Total Participant, Training Cost only)

Header:

A1-A2: Training Title

B1-B2: Training Date (From)

C1-C2: Training Date (To)

D1-D2: "Total Participant"

E1-L1:  Training Cost (RM)

E2:  Training Fees

F2:  Meal

G2:  Subsistance Allowance

H2:  Hotel Fees

I2:  Mileage Claim

J2:  Taxi Fees

K2:  Toll Fees

L2:  Flight

M1-M2:  Total  Cost 2026

N1-N2:  Total HRDF Grant (RM)



3\. Training Title
Header:

Name | Emp. No | Training Title | (EE)/Cost Centre Description | Training Type | Date (From) | Date (Until) | Month | Total Hours | Trainer | Training Provide | Expiry Date (if applicable)



4\. Employee

5\. Filter by month



**Annual Training Plan Requirement:**

1. Header:

No | Training Title | Training Category | TNA Source | Training Mode | Training Duration (hrs) | Trainer | Department | Position / Employee No | Total Pax | Planned Date | Actual Date (From) | Actual Date (To) | Training Status | Remarks



**Improvement (Email):**

1. Create image to inform user (plain text not nice)
2. Approval status for HOD, Csuite and HOHR does not need to inform requester, just inform if Reject, Return or complete approval only



**Improvement (Training Request Form)**

1. Edit Training (if the training request is return by HOD, Csuite, HOHR)
2. Employee can see previous training request that they have submit (according to their employee ID)



**Cleanup (for Deployment)**

1. Deploy using it@apollofood.com.my
2. overall UI enhancement
3. Admin login (no option to change email)
4. HOD portal (remove profile switcher), lock status after submit
5. ~~Participant UI: employee ID only~~
6. ~~Only allow view \& fill attendance and evaluation, after training got full approval only~~
7. Goldenscoop access to all website
8. ~~Auto update sheets when employee request training~~



**Annual Training Plan (correct data):**

1. TNA Source: Training Requisition Form

2\. Training Mode: In-house or Public

3\. Position: Name list (hyperlink to list of participant in another sheets just for training participant for the specific training)

4\. Date format: 01st Jan 2026

5\. Training status: Completed, On planning, On going, Pending



**UI Cleanup:**
TNA Source: "Training Requisition Form" and "Training Procedure"

Training Type: add "On-Job Training"

**Training Request Form (edit):**
1. missing department/ cost centre when select multiple (not saved/shown all)

2\. missing training provider 

3\. missing cert expiry date
4. missing participants list



**Admin:**
1. Returned is count as fully approve in admin dashboard

2\. Training Programmes detail is not detailed enough

3\. QR attendance does not have Apollo logo in the middle

4\. All date and time should remove timezone, keep it simple



**Attendance:**

1. Date and time is incorrect, too messy
2. Just Date and time (no time zone)
3. Enter employee ID should display name first then can click continue
4. When click continue, it says check in rejected eventhough the employee already in the participant list
5. Employee ID example should be "00000"



**Post evaluation:**
1. Admin assign/select employee for post evaluation to supervisor or person in charge

2\. System check supervisor email or employee ID

3\. System display pending evaluation (list of employee to be evaluate under the supervisor, separate by training)

