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

1. No multiple selection for Cost Centre
2. Add file button
3. Participant name displayed wrongly contains "()"
4. When add participant B, participant A are added
5. Bulk add participant not working
6. Missing No. of pax question
7. When trying to send form, it says "⚠️ Error: Official company email is required to submit a training requisition."
8. FOR TESTING/DEVELOPEMENT ADD THE EMAIL TO DRAFT FIRST



**Improvement/Problem (HOD system)**

1. Only allow HOD (match the HOD email) can access the system
2. If the user don't have access, give a proper UI access denied



**Improvement/Problem (Admin system)**

1. Training keeps loading forever, unable to see "No training list" or create new training

