# Timesheet_Scanner
4 x VBA subs that scans a staffs timesheet and 1) checks if they have the timesheet for the current week 2) ensures that the project number they use on their timesheet is a real project number 3) ensures that they are allowed to add to a specific project and 4) the rest of their timesheet is error free


The timesheet checking sequence is;
1) weekCheck
2) activeStaff
3) activeProject
4) PLACTCheck

Pending passing all of the above then add the timesheet info to the correct project cost control file via:
1) pullTimesheetInfo sub
