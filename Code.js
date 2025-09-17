/**
 * @title nams-campus-transition-notification-project-25-26
 * 
 * @description This script is designed to be run from a Google Sheet. It will send an email
 * with attachments and links to a list of recipients based on the value in the
 * "Campus" column of the sheet.
 * 
 * Procedure: The user will first create a transition letter using Autocrat. The letter
 * will be saved in a campus specific shared Google Drive folder.
 * When the user is ready to send the email, he clicks on 'Notify Campuses'
 * in the menubar followed by the option that is provided in the dropdown.
 * An email will be sent to the campus administrators with a link to the folder and the
 * transition letter.
 * 
 * @projectLead Reggie Ollendieck, Associate Principal, NAMS
 * @author Alvaro Gomez, Academic Technology Coach, 1-210-397-9408, alvaro.gomez@nisd.net
 * @lastUpdated 08/11/25
 */

/**
 * The name of the column used to track the return date for sent emails.
 * @constant {string}
 */
const RETURN_DATE = "Anticipated Return date";

/**
 * The name of the column used to track the date when the email was sent to campuses.
 * @constant {string}
 */
const DATE_SENT_COL = "Date when the email was sent to campuses";

/**
 * The name of the column used to store the campus folder ID.
 * @constant {string}
 */
const CAMPUS_FOLDER_COL = "Campus folder ID";

/**
 * Adds a custom menu to the Google Sheets UI when the spreadsheet is opened.
 * The menu allows users to preview emails or send emails to campuses based
 * on specific criteria.
 * @function
 * @returns {void}
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📬 Notify Campuses")
    .addItem(
      "Preview emails (don't send)",
      "previewEmails"
    )
    .addItem(
      "Send emails to campuses",
      "sendEmails"
    )
    .addToUi();
}

/**
 * Preview emails without sending them.
 * Shows a summary of what would be sent.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} [sheet] - The sheet to process.
 * @returns {void}
 */
function previewEmails(
  sheet = SpreadsheetApp.openById(
    "1p0KwjAMLnI4KyPG1ErNq0oRmGpQ8GWGDSnYVa7MYnP0"
  ).getSheetByName("Teacher Notes")
) {
  processEmails(sheet, true);
}

/**
 * Sends emails to campuses based on the data in the sheet.
 * Only sends emails for rows with an anticipated return date, null for date when the
 * email was sent to campuses, and has a campus folder ID.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} [sheet] - The sheet to process. Defaults to the 'Teacher Notes' sheet.
 * @returns {void}
 */
function sendEmails(
  sheet = SpreadsheetApp.openById(
    "1p0KwjAMLnI4KyPG1ErNq0oRmGpQ8GWGDSnYVa7MYnP0"
  ).getSheetByName("Teacher Notes")
) {
  processEmails(sheet, false);
}

/**
 * Main function to process emails with preview or send mode.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to process.
 * @param {boolean} previewMode - If true, shows preview without sending.
 * @returns {void}
 */
function processEmails(sheet, previewMode = false) {
  const startTime = new Date();
  Logger.log(`${previewMode ? 'Preview' : 'Send'} operation started at ${startTime}`);
  
  // Validation: Check sheet
  const sheetName = sheet.getName();
  const targetSheetname = "Teacher Notes";

  if (sheetName !== targetSheetname) {
    const errorMsg = `Function can only be run from the "${targetSheetname}" sheet. Current sheet: "${sheetName}"`;
    Logger.log(`ERROR: ${errorMsg}`);
    SpreadsheetApp.getUi().alert(errorMsg);
    return;
  }

  // Validation: Check if sheet has data
  if (sheet.getLastRow() < 2) {
    const errorMsg = "No data found in sheet. Sheet must have at least a header row and one data row.";
    Logger.log(`ERROR: ${errorMsg}`);
    SpreadsheetApp.getUi().alert(errorMsg);
    return;
  }

  const dataRange = sheet.getRange(
    2,
    1,
    sheet.getLastRow() - 1,
    sheet.getLastColumn()
  );
  const data = dataRange.getDisplayValues();
  const heads = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Validation: Check required columns
  const requiredColumns = [RETURN_DATE, DATE_SENT_COL, CAMPUS_FOLDER_COL, "Campus", "Name"];
  const missingColumns = requiredColumns.filter(col => heads.indexOf(col) === -1);
  
  if (missingColumns.length > 0) {
    const errorMsg = `Missing required columns: ${missingColumns.join(", ")}`;
    Logger.log(`ERROR: ${errorMsg}`);
    SpreadsheetApp.getUi().alert(errorMsg);
    return;
  }

  const emailSentColIdx = heads.indexOf(DATE_SENT_COL);

  const obj = data.map((r) =>
    heads.reduce((o, k, i) => ((o[k] = r[i] || ""), o), {})
  );

  // Group students by campus
  const campusGroups = {};
  const validRows = [];

  obj.forEach(function (row, rowIdx) {
    if (
      row[RETURN_DATE].trim() !== "" &&
      row[DATE_SENT_COL] === "" &&
      row[CAMPUS_FOLDER_COL] !== ""
    ) {
      const campus = row["Campus"];
      if (!campus || campus.trim() === "") {
        Logger.log(`WARNING: Row ${rowIdx + 2} has empty campus field`);
        return;
      }
      
      if (!campusGroups[campus]) {
        campusGroups[campus] = [];
      }
      campusGroups[campus].push({ row, rowIdx });
      validRows.push({ row, rowIdx });
    }
  });

  const campusCount = Object.keys(campusGroups).length;
  const totalStudents = validRows.length;

  Logger.log(`Found ${totalStudents} students across ${campusCount} campuses`);

  if (totalStudents === 0) {
    const msg = "No students found that meet the criteria (have return date, no sent date, and campus folder ID).";
    Logger.log(`INFO: ${msg}`);
    SpreadsheetApp.getUi().alert(msg);
    return;
  }

  // Preview mode - show summary and exit
  if (previewMode) {
    let previewMessage = `PREVIEW MODE - No emails will be sent\n\n`;
    previewMessage += `Found ${totalStudents} students across ${campusCount} campuses:\n\n`;
    
    Object.keys(campusGroups).forEach(function (campus) {
      const students = campusGroups[campus];
      previewMessage += `${campus}: ${students.length} student${students.length === 1 ? '' : 's'}\n`;
      students.forEach(function(student) {
        previewMessage += `  • ${student.row["Name"]} (returns: ${student.row[RETURN_DATE]})\n`;
      });
      previewMessage += `\n`;
    });
    
    SpreadsheetApp.getUi().alert(previewMessage);
    Logger.log(`Preview completed. ${totalStudents} students ready to process.`);
    return;
  }

  // Confirmation dialog for sending
  const confirmMessage = `Ready to send emails to ${campusCount} campuses for ${totalStudents} students.\n\nProceed with sending emails?`;
  const response = SpreadsheetApp.getUi().alert(
    "Confirm Email Send",
    confirmMessage,
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );

  if (response !== SpreadsheetApp.getUi().Button.YES) {
    Logger.log("Email send cancelled by user");
    return;
  }

  // Initialize counters for success and error counts
  let successCount = 0;
  let errorCount = 0;
  let errorMessages = [];

  // Send one email per campus with all students
  const campusList = Object.keys(campusGroups);
  
  Logger.log(`Starting to send emails to ${campusList.length} campuses`);
  
  campusList.forEach(function (campus, index) {
    Logger.log(`Processing campus ${index + 1}/${campusList.length}: ${campus}`);
    
    try {
      const students = campusGroups[campus];
      const campusInfo = getInfoByCampus(campus);
      
      // Validation: Check if campus info exists
      if (!campusInfo.recipients || !campusInfo.driveLink) {
        throw new Error(`No configuration found for campus: ${campus}`);
      }
      
      let recipients = campusInfo.recipients;
      // Ensure recipients is a comma-separated string
      if (Array.isArray(recipients)) {
        recipients = recipients.join(",");
      }
      
      // Validation: Check recipients
      if (!recipients || recipients.trim() === "") {
        throw new Error(`No recipients configured for campus: ${campus}`);
      }
      
      const driveLink = campusInfo.driveLink;
      
      Logger.log(`  Sending to ${recipients} for ${students.length} students`);
      
      // Create grouped email template
      const emailTemplate = getGroupedGmailTemplate_(students, driveLink, campus);
      const msgObj = fillInTemplateFromObject_(
        emailTemplate.message,
        { Campus: campus },
        driveLink
      );

      GmailApp.sendEmail(recipients, msgObj.subject, msgObj.text, {
        htmlBody: msgObj.html,
        replyTo: "reggie.ollendieck@nisd.net",
        // cc: "reggie.ollendieck@nisd.net,zina.gonzales@nisd.net",
      });
      
      successCount++;
      Logger.log(`  ✓ Email sent successfully to ${campus}`);
      
      // Mark all rows in this campus as sent using batch operation
      const updates = students.map(function(student) {
        return [student.rowIdx + 2, emailSentColIdx + 1, new Date()];
      });
      
      updates.forEach(function(update) {
        sheet.getRange(update[0], update[1]).setValue(update[2]);
      });
      
      Logger.log(`  ✓ Marked ${students.length} students as sent`);
      
    } catch (e) {
      const errorMsg = `Campus ${campus}: ${e.message}`;
      Logger.log(`  ✗ ERROR: ${errorMsg}`);
      console.error(`Detailed error for ${campus}:`, e);
      errorCount++;
      errorMessages.push(errorMsg);
    }
  });

  const endTime = new Date();
  const duration = Math.round((endTime - startTime) / 1000);
  const resultMessage = `Operation completed in ${duration} seconds\n\nEmails Sent: ${successCount}\nErrors: ${errorCount}${errorMessages.length > 0 ? '\n\nErrors:\n' + errorMessages.join('\n') : ''}`;
  
  Logger.log(`Final results: ${successCount} successful, ${errorCount} errors`);
  SpreadsheetApp.getUi().alert(resultMessage);

  /**
   * Returns campus-specific information including recipients and drive folder link.
   * @param {string} campusValue - The campus name.
   * @returns {{recipients: string[]|string, driveLink: string}} Campus info object.
   */
  function getInfoByCampus(campusValue) {
    switch (campusValue.toLowerCase()) {
      case "bernal":
        return {
          recipients: [
            "david.laboy@nisd.net",
            "Marla.Reynolds@nisd.net",
            "sally.maher@nisd.net",
            "monica.flores@nisd.net",
          ],
          driveLink: "1QlavZvp-8tvqiPF3zYQanZ9SQ7UhiJaE",
        };
      case "briscoe":
        return {
          recipients: [
            "joe.bishop@nisd.net",
            "francesca.parker@nisd.net",
            "brigitte.rauschuber@nisd.net",
            "xavier.aguirre@nisd.net",
          ],
          driveLink: "1JgHL75iSr5F1lgHLieZl0y_psuc7gp-N",
        };
      case "connally":
        return {
          recipients: [
            "erica.robles@nisd.net",
            "monica.ramirez@nisd.net"
          ],
          driveLink: "1cWPX_nOXb9yldONekm9Ba3Rsksl9yqKe",
        };
      case "folks":
        return {
          recipients: [
            "yvette.lopez@nisd.net",
            "miguel.trevino@nisd.net",
            "terry.precie@nisd.net",
            "angelica.perez@nisd.net",
            "norma.esparza@nisd.net",
            "james-1.garza@nisd.net",
            "keli.hall@nisd.net",
            "ann.devlin@nisd.net"
          ],
          driveLink: "1MZ9MCh1DmI9cJFWbr5BA-v1jJh5Dd1E9",
        };
      case "garcia":
        return {
          recipients: [
            "mateo.macias@nisd.net",
            "anna.lopez@nisd.net",
            "julie.minnis@nisd.net",
            "mark.lopez@nisd.net",
            "lori.persyn@nisd.net",
          ],
          driveLink: "1D8_8q9fcB6tn3fXX8xnlkGE8aro8fbDm",
        };
      case "hobby":
        return {
          recipients: [
            "gregory.dylla@nisd.net",
            "marian.johnson@nisd.net",
            "lawrence.carranco@nisd.net",
            "jose.texidor@nisd.net",
            "victoria.denton@nisd.net"
          ],
          driveLink: "1u76TEIq5BbCNCG-i8Ta73VaY7NTOsK4x",
        };
      case "hobby magnet":
        return {
          recipients: ["jaime.heye@nisd.net"],
          driveLink: "1YrvPC7m0eX128C0dR3u4RGNfU82MXOg2",
        };
      case "holmgreen":
        return {
          recipients: ["cheryl.parra@nisd.net", "frank.johnson@nisd.net"],
          driveLink: "1c8ufu7MvAsNwQAwx0IOwi4O1dNU1-7yo",
        };
      case "jefferson":
        return {
          recipients: [
            "monica.cabico@nisd.net",
            "Nicole.Gomez@nisd.net",
            "tiffany.watkins@nisd.net",
            "catherine.villela@nisd.net",
          ],
          driveLink: "1LlQxIJBeCxV440nalwS5fjzu8_1dYvlx",
        };
      case "jones":
        return {
          recipients: [
            "rudolph.arzola@nisd.net",
            "nicole.mcevoy@nisd.net",
            "erica.lashley@nisd.net",
            "javier.lazo@nisd.net",
            "aaron.logan@nisd.net",
          ],
          driveLink: "1jBxe9OFTTcones277XnehgvW4EnM4YQk",
        };
      case "jones magnet":
        return {
          recipients: ["david.johnston@nisd.net"],
          driveLink: "1oeMqPEr_cpstSWRa0LV-uodQwteQKRWn",
        };
      case "jordan":
        return {
          recipients: [
            "Shannon.Zavala@nisd.net",
            "juaquin.zavala@nisd.net",
            "erica.parra@nisd.net",
            "laurel.graham@nisd.net",
            "robert.ruiz@nisd.net",
            "anabel.romero@nisd.net",
          ],
          driveLink: "1T90JGPgUu7DhfBBytxRrfgqkMBrkzOAN",
        };
      case "jordan magnet":
        return {
          recipients: ["jessica.marcha@nisd.net"],
          driveLink: "1BVECP6fsaqGXap1uEOchy5SW-6PsADM9",
        };
      case "luna":
        return {
          recipients: [
            "leti.chapa@nisd.net",
            "jennifer.cipollone@nisd.net",
            "amanda.king@nisd.net",
            "lisa.richard@nisd.net",
          ],
          driveLink: "10DdpBdHwp7bH5ph-pKvfbvG23tsi9ZMc",
        };
      case "neff":
        return {
          recipients: [
            "yvonne.correa@nisd.net",
            "theresa.heim@nisd.net",
            "laura-i.sanroman@nisd.net",
            "joseph.castellanos@nisd.net",
            "mackenzie.fulton@nisd.net",
            "adriana.aguero@nisd.net",
            "priscilla.vela@nisd.net",
            "sarah.tennery@nisd.net",
            "jessica.montalvo@nisd.net",
            "hayley.giorgio@nisd.net",
          ],
          driveLink: "1Sd7DrcgHjnAcuqR79DVmISnVn3Bzzfyn",
        };
      case "pease":
        return {
          recipients: [
            "Lynda.Desutter@nisd.net",
            "jessica-1.barrera@nisd.net",                      
            "guadalupe.brister@nisd.net",
          ],
          driveLink: "1tOusVf1SxNckZC5ro-dwBlk8-YpKUMuh",
        };
      case "rawlinson":
        return {
          recipients: [
            "jesus.villela@nisd.net",
            "david.rojas@nisd.net",
            "nicole.buentello@nisd.net",
          ],
          driveLink: "1p9IXf40oikwOrxSSmonwnBJ-dOJB7Ui6",
        };
      case "rayburn":
        return {
          recipients: [
            "robert.alvarado@nisd.net",
            "maricela.garza@nisd.net",
            "carol.zule@nisd.net",
            "micaela.welsh@nisd.net",
          ],
          driveLink: "1DcP6LUpcT8wT9PEYgc4_dwk2mclWI9bp",
        };
      case "ross":
        return {
          recipients: [
            "christina.lozano@nisd.net",            
            "priscilla.sigala@nisd.net",
            "dolores.cardenas@nisd.net",
            "katherine.vela@nisd.net",
            "roxanne.romo@nisd.net",
            "cristina.castillo@nisd.net",
          ],
          driveLink: "1KGyYAJF5Qf-Gt0oWvRRjH2Vw_DTARMcg",
        };
      case "rudder":
        return {
          recipients: [
            "catelyn.vasquez@nisd.net",
            "jeanette.navarro@nisd.net",
            "jason.padron@nisd.net",
            "adrian.hysten@nisd.net",
          ],
          driveLink: "1a3PiBLrTtsMJkR6lthx86gUqd2_qeF-D",
        };
      case "stevenson":
        return {
          recipients: [
            "chaeleen.garcia@nisd.net",
            "anthony.allen01@nisd.net",
            "hilary.pilaczynski@nisd.net",
            "johanna.davenport@nisd.net",
          ],
          driveLink: "1Y43jZCtjKFbF-I09lBS6wnr-Vbl0p_Cm",
        };
      case "stinson":
        return {
          recipients: [
            "lourdes.medina@nisd.net",
            "louis.villarreal@nisd.net",
            "jeannette.rainey@nisd.net",
            "linda.boyett@nisd.net",
            "elda.garza@nisd.net",
            "maranda.luna@nisd.net",
            "alexis.lopez@nisd.net",
          ],
          driveLink: "1pG4NIUveTfv46DvQoq0TD4uuwaXCLwkS",
        };
      case "straus":
        return {
          recipients: [
            "araceli.farias@nisd.net",
            "jose.gonzalez02@nisd.net",
            "leigh.davis@nisd.net",
          ],
          driveLink: "15p9xqZoyikuVRk4ZVi7sUYYgoxL1PSew",
        };
      case "vale":
        return {
          recipients: [
            "jenna.bloom@nisd.net",
            "brenda.rayburg@nisd.net",
            "daniel.novosad@nisd.net",
            "mary.harrington@nisd.net",
          ],
          driveLink: "1QoRQNEt7_gWT3PC_DPnDFjlT1XKVQ3sN",
        };
      case "zachry":
        return {
          recipients: [
            "Richard.DeLaGarza@nisd.net",
            "randolph.neuenfeldt@nisd.net",
            "jennifer-a.garcia@nisd.net",
            "jimann.caliva@nisd.net",
            "veronica.poblano@nisd.net",
          ],
          driveLink: "1UZDcETdHG5eN9DSCdfKcV0wt72cVY7Ek",
        };
      case "zachry magnet":
        return {
          recipients: ["matthew.patty@nisd.net"],
          driveLink: "1wMjhAx6wGOw5j-tu7-4wTNnH7V80pfPe",
        };
      case "test":
        return {
          recipients: [
            "reggie.ollendieck@nisd.net",
            "zina.gonzales@nisd.net",
            // "alvaro.gomez@nisd.net"
          ],
          driveLink: "1nMJAEcGIh_QnhfS5gjCkKd6CtoA3r5cf",
        };
      default:
        return {
          recipients: "",
          driveLink: "",
        };
    }
  }

  /**
   * Generates the Gmail message template for multiple students grouped by campus.
   * @param {Array} students - Array of student objects with row and rowIdx properties.
   * @param {string} driveLink - The campus drive folder link.
   * @param {string} campus - The campus name.
   * @returns {{message: {subject: string, html: string}}} Gmail message template object.
   */
  function getGroupedGmailTemplate_(students, driveLink, campus) {
    const studentCount = students.length;
    const studentWord = studentCount === 1 ? "student" : "students";
    
    // Build student list with links
    let studentList = "";
    students.forEach(function(student, index) {
      const row = student.row;
      studentList += `
        <li>
          <strong>${row["Name"]}</strong> - Expected return: ${row[RETURN_DATE]}<br>
        </li>`;
    });

    return {
      message: {
        subject: `DAEP Placement Transition Plan Notification`,
        html: `Dear ${campus},<br><br>
        You have ${studentCount} ${studentWord} who ${
          studentCount === 1 ? "has" : "have"
        } nearly completed their assigned placement at NAMS and should be returning to ${campus} soon.<br><br>   
              On their last day of placement, they will be given withdrawal documents and the parents/guardians will have been called and told to contact ${campus} to set up an appointment to re-enroll and meet with an administrator/counselor.<br><br>  
              Below is a list of the returning ${studentWord}. You can find their DAEP Transition Plans (with grades and notes from their teachers at NAMS) linked below in your campus folder.<br><br>

              <h3>${studentWord} returning to ${campus}:</h3>
              <ul>${studentList}</ul>

              <h3>Important Links:</h3>
              <ul>
                <li><a href="https://drive.google.com/drive/folders/${driveLink}">${campus} Here is a link to your campus folder for the year.</a></li>
                <li><a href="https://drive.google.com/file/d/1qnyQ8cCxLVM9D6rg4wkyBp6KrXIELfNx/view?usp=sharing">Updates in Special Education</a></li>
              </ul>
              
              Please let me know if you have any questions or concerns.<br><br>
              Thank you for all you do,<br>
              Reggie Ollendieck<br>
              Associate Principal<br><br>`,
      },
    };
  }

  /**
   * Fills in a template string with data from the row and drive link.
   * @param {string|Object} template - The template string or object.
   * @param {Object} row - The row data object.
   * @param {string} driveLink - The campus drive folder link.
   * @returns {Object} The filled-in template object.
   */
  function fillInTemplateFromObject_(template, row, driveLink) {
    let template_string = JSON.stringify(template);
    template_string = template_string.replace(/{{[^{}]+}}/g, (key) => {
      if (key === "${driveLink}") {
        return escapeData_(driveLink);
      }
      return escapeData_(row[key.replace(/[{}]+/g, "")] || "", driveLink);
    });
    return JSON.parse(template_string);
  }

  /**
   * Escapes special characters in a string for safe insertion into templates.
   * @param {string} str - The string to escape.
   * @returns {string} The escaped string.
   */
  function escapeData_(str) {
    return str
      .replace(/[\\]/g, "\\\\")
      .replace(/[\"]/g, '\\"')
      .replace(/[\/]/g, "\\/")
      .replace(/[\b]/g, "\\b")
      .replace(/[\f]/g, "\\f")
      .replace(/[\n]/g, "\\n")
      .replace(/[\r]/g, "\\r")
      .replace(/[\t]/g, "\\t");
  }
}
