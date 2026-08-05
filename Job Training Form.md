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





**NEW FLOW**

1. Employee fill \& submit the Training Requisition Form, employee ID requester (only those with company email can fill the form)
2. The system autofill the Training Requisition Form including the Request By (employee name, cost centre and date)
3. System send email to HOD/Manager (the system link)
4. HOD/Manager approve/reject/postpone/reschedule the request through the system
5. System marks the form as approve/reject digitally with timestamp
6. The status of the request will be emailed to the person who request
7. Postponed or cancelled training still recorded in history (may be reschedule/carry forward)
8. Admin can view the request status
9. For the approved training, admin can generate QR attendance
10. When participant finish the training, the participant need to fill the training evaluation form
11. Participant need to submit the training form within 2 weeks after the course completion
12. The system disallow training evaluation for participant who absent the training
13. The system auto email to correspond HOD after 3 months to evaluate the participant
14. HOD can click the link in the email to answer the post evaluation form for each participant under the HOD
15. The post evaluation form have list of participant to be evaluated (UNDER THE HOD ONLY) and each submission will display the list of participant that have not been evaluated yet
16. The system will create report based on the report format (will give later/ next time)



**NEW REQUIREMENT**

1. New separate App Script for HOD, the system is to display the training request for HOD to review (add new separate folder for this system) (can be access by company acc only)
2. On the bottom of the training request (scroll down), have a section to approve (show the HOD employee ID, name and Cost Centre)
3. The post evaluation form are also here after 3 month of the programme

