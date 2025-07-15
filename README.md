# NAMS CAMPUS TRANSITION NOTIFICATION PROJECT 2025-2026
## Academic Technology & Northside Alternative Middle School (NAMS)
<b>Project Lead:</b> Reggie Ollendieck, Associate Principal

<b>Script Developer:</b> Alvaro Gomez, Academic Technology Coach

<b>Purpose:</b> This project supports NAMS' requirement from the Texas Education Agency to coordinate the transition from a Disciplinary Alternative Education (DAEP) to a regular classroom [§37.023](https://statutes.capitol.texas.gov/docs/ed/htm/ed.37.htm).

This is a Google Sheet bound Google Apps Script that sends an email to administrators of receiving campuses with links to the student's Personalized Transition Plan, the student's AEP Transition Plan, and a link to the Campus's folder with all of the transition plans for the year.

<b>Procedure:</b>
1. Teachers fill out the Google Sheet with the necessary information.
2. Administrator reviews the input provided by the teachers, makes revisions, then adds a return date in column F.
3. Administrator runs two Autocrat jobs: <i>Home Campus Transition Plan</i> and <i>Student Transition</i>. These two jobs are placed dynamically (with Autocrat) in the receiving campus' folder. The autocrat jobs are set with dynamic folder references in order to put them into the right campus folder. Each middle school has their own folder.
4. Administrator sends the notification emails by clicking on " 📬 Notify Campus ". This will send emails to the administrators hardcoded in the script.

<b>Beginning of Year Requirements:</b>
1. Create a main folder that will contain each middle school's individual folders.
2. Inside of the main folder, create a folder for each of the middle schools. Keep track of each folder's id because you will need to add it to the script and into the "CampusReferenceInfo" sheet. Also, share the folder with the administrators who need access to the letters.
3. Go to the sheet named "CampusReferenceInfo" and add each campus' folder id in column B.
4. Go to the script and update each campus' administrators information and folder id in the script.
5. Create the two Autocrat templates (AEP Transition Plan and Personalized Transition Plan).
6. Create the two Autocrat print jobs (one for each of the Autocrat templates). On step 6 of creating the Autocrat print jobs, add the dynamic folder reference (use <<Campus folder ID>>).

<b>Throughout the Year</b>:

Return to the script to update administrators. For example new administrators need to be included in the emails or administrators have moved to other roles and don't need the email notifications anymore.
Also, if administrators are added or removed, remember to update the share privileges in their campus folders as well.