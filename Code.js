/**
 * @title nams-campus-transition-notification-project-25-26
 * @description This script is designed to be run from a Google Sheet. It will send an email with attachments and links to a list of recipients based on the value in the "Campus" column of the sheet. The administrator will first create two letters using Autocrat. The letters will be saved in a campus specific shared Google Drive folder. When the adminstrator is ready to send the email, he clicks on the 'Notify Campuses' user menu and then the option that is provided by the dropdown. An email will be sent to the campus administrators with a link to the folder and the two letters.
 * @projectLead Reggie Ollendieck, Associate Principal, NAMS
 * @author Alvaro Gomez, Academic Technology Coach, 1-210-397-9408, alvaro.gomez@nisd.net
 * @lastUpdated 07/15/25
 */

const EMAIL_SENT_COL = "Return date";
const DATE_SENT_COL = "Date when the email was sent to campuses";
const CAMPUS_FOLDER_COL = "Campus folder ID";

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📬 Notify Campuses")
    .addItem(
      "Send emails to campuses with a date in column F and a blank in column BD",
      "sendEmails"
    )
    .addToUi();
}

function sendEmails(
  sheet = SpreadsheetApp.openById(
    "1p0KwjAMLnI4KyPG1ErNq0oRmGpQ8GWGDSnYVa7MYnP0"
  ).getSheetByName("Teacher Notes")
) {
  const sheetName = sheet.getName();
  const targetSheetname = "Teacher Notes";

  if (sheetName !== targetSheetname) {
    SpreadsheetApp.getUi().alert(
      'This function can only be run from the "' + targetSheetname + '" sheet.'
    );
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

  const emailSentColIdx = heads.indexOf(DATE_SENT_COL);

  if (emailSentColIdx === -1) {
    SpreadsheetApp.getUi().alert("Required column is missing.");
  }

  const obj = data.map((r) =>
    heads.reduce((o, k, i) => ((o[k] = r[i] || ""), o), {})
  );

  // Initialize counters for success and error counts for dialog display
  let successCount = 0;
  let errorCount = 0;
  let errorMessages = []; // Stores the error messages for dialog display

  obj.forEach(function (row, rowIdx) {
    if (
      row[EMAIL_SENT_COL].trim() !== "" &&
      row[DATE_SENT_COL] === "" &&
      row[CAMPUS_FOLDER_COL] !== ""
    ) {
      try {
        const campusInfo = getInfoByCampus(row["Campus"]);
        const recipients = campusInfo.recipients;
        const driveLink = campusInfo.driveLink;
        const emailTemplate = getGmailTemplateFromDrafts_(row, driveLink);
        const msgObj = fillInTemplateFromObject_(
          emailTemplate.message,
          row,
          driveLink
        );

        GmailApp.sendEmail(recipients, msgObj.subject, msgObj.text, {
          htmlBody: msgObj.html,
          replyTo: "reggie.ollendieck@nisd.net",
          cc: "reggie.ollendieck@nisd.net",
        });

        successCount++;
        sheet.getRange(rowIdx + 2, emailSentColIdx + 1).setValue(new Date());
      } catch (e) {
        Logger.log(`Error on row ${rowIdx + 2}: ${e.message}`);
        errorCount++;
        errorMessages.push(`Row ${rowIdx + 2}: ${e.message}`);
      }
    }
  });

  SpreadsheetApp.getUi().alert(
    `Emails Sent: ${successCount}\nErrors: ${errorCount}\n${errorMessages.join(
      "\n"
    )}`
  );

  function getInfoByCampus(campusValue) {
    switch (campusValue.toLowerCase()) {
      case "bernal":
        return {
          recipients: [
            "david.laboy@nisd.net",
            "jose.mendez@nisd.net",
            "monica.flores@nisd.net",
            "sally.maher@nisd.net",
          ],
          driveLink: "1QlavZvp-8tvqiPF3zYQanZ9SQ7UhiJaE",
        };
      case "briscoe":
        return {
          recipients: [
            "Joe.bishop@nisd.net",
            "nereida.ollendieck@nisd.net",
            "brigitte.rauschuber@nisd.net",
            "veronica.martinezl@nisd.net",
          ],
          driveLink: "1JgHL75iSr5F1lgHLieZl0y_psuc7gp-N",
        };
      case "connally":
        return {
          recipients: ["erica.robles@nisd.net", "monica.ramirez@nisd.net"],
          driveLink: "1cWPX_nOXb9yldONekm9Ba3Rsksl9yqKe",
        };
      case "folks":
        return {
          recipients: [
            "yvette.lopez@nisd.net",
            "miguel.trevino@nisd.net",
            "terry.precie@nisd.net",
            "angie.perez@nisd.net",
            "norma.esparza@nisd.net",
            "james.garza@nisd.net",
            "keli.hall@nisd.net",
            "ann.devlin@nisd.net",
            "kendra.delavega@nisd.net",
          ],
          driveLink: "1MZ9MCh1DmI9cJFWbr5BA-v1jJh5Dd1E9",
        };
      case "garcia":
        return {
          recipients: [
            "mark.lopez@nisd.net",
            "lori.persyn@nisd.net",
            "julie.minnis@nisd.net",
            "mateo.macias@nisd.net",
            "anna.lopez@nisd.net",
          ],
          driveLink: "1D8_8q9fcB6tn3fXX8xnlkGE8aro8fbDm",
        };
      case "hobby":
        return {
          recipients: [
            "gregory.dylla@nisd.net",
            "marian.johnson@nisd.net",
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
            "Nicole.aguirreGomez@nisd.net",
            "monica.cabico@nisd.net",
            "maria-1.martinez@nisd.net",
            "tiffany.watkins@nisd.net",
            "catherine.villela@nisd.net",
          ],
          driveLink: "1LlQxIJBeCxV440nalwS5fjzu8_1dYvlx",
        };
      case "jones":
        return {
          recipients: [
            "rudolph.arzola@nisd.net",
            "javier.lazo@nisd.net",
            "erica.lashley@nisd.net",
            "nicole.mcevoy@nisd.net",
            "brian.pfeiffer@nisd.net",
          ],
          driveLink: "1jBxe9OFTTcones277XnehgvW4EnM4YQk",
        };
      case "jones magnet":
        return {
          recipients: ["xavier.maldonado@nisd.net"],
          driveLink: "1oeMqPEr_cpstSWRa0LV-uodQwteQKRWn",
        };
      case "jordan":
        return {
          recipients: [
            "anabel.romero@nisd.net",
            "Shannon.Zavala@nisd.net",
            "laurel.graham@nisd.net",
            "robert.ruiz@nisd.net",
            "ryanne.barecky@nisd.net",
            "adrian.hysten@nisd.net",
            "patti.vlieger@nisd.net",
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
            "joseph.castellanos@nisd.net",
            "adriana.aguero@nisd.net",
            "priscilla.vela@nisd.net",
            "michele.adkins@nisd.net",
            "laura-i.sanroman@nisd.net",
            "jessica.montalvo@nisd.net",
          ],
          driveLink: "1Sd7DrcgHjnAcuqR79DVmISnVn3Bzzfyn",
        };
      case "pease":
        return {
          recipients: [
            "Lynda.Desutter@nisd.net",
            "tamara.campbell-babin@nisd.net",
            "tiffany.flores@nisd.net",
            "tanya.alanis@nisd.net",
            "brian.pfeiffer@nisd.net",
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
            "elizabeth.smith@nisd.net",
          ],
          driveLink: "1p9IXf40oikwOrxSSmonwnBJ-dOJB7Ui6",
        };
      case "rayburn":
        return {
          recipients: [
            "Robert.Alvarado@nisd.net",
            "maricela.garza@nisd.net",
            "aissa.zambrano@nisd.net",
            "brandon.masters@nisd.net",
            "micaela.welsh@nisd.net",
          ],
          driveLink: "1DcP6LUpcT8wT9PEYgc4_dwk2mclWI9bp",
        };
      case "ross":
        return {
          recipients: [
            "mahntie.reeves@nisd.net",
            "christina.lozano@nisd.net",
            "claudia.salazar@nisd.net",
            "jason.padron@nisd.net",
            "katherine.vela@nisd.net",
            "dolores.cardenas@nisd.net",
            "roxanne.romo@nisd.net",
            "rose.vincent@nisd.net",
          ],
          driveLink: "1KGyYAJF5Qf-Gt0oWvRRjH2Vw_DTARMcg",
        };
      case "rudder":
        return {
          recipients: [
            "kevin.vanlanham@nisd.net",
            "catelyn.vasquez@nisd.net",
            "barbra.bloomingdale@nisd.net",
            "susana.duran@nisd.net",
            "jesus.alonzo@nisd.net",
            "grissel.gandaria@nisd.net",
            "amandam.gonzalez@nisd.net",
            "jeanette.navarro@nisd.net",
            "paul.ramirez@nisd.net",
          ],
          driveLink: "1a3PiBLrTtsMJkR6lthx86gUqd2_qeF-D",
        };
      case "stevenson":
        return {
          recipients: [
            "anthony.allen01@nisd.net",
            "hilary.pilaczynski@nisd.net",
            "johanna.davenport@nisd.net",
            "amanda.cardenas@nisd.net",
            "chaeleen.garcia@nisd.net",
          ],
          driveLink: "1Y43jZCtjKFbF-I09lBS6wnr-Vbl0p_Cm",
        };
      case "stinson":
        return {
          recipients: [
            "louis.villarreal@nisd.net",
            "lourdes.medina@nisd.net",
            "jeannette.rainey@nisd.net",
            "rick.lane@nisd.net",
            "linda.boyett@nisd.net",
            "elda.garza@nisd.net",
            "maria.figueroa@nisd.net",
            "maranda.luna@nisd.net",
            "alexis.lopez@nisd.net",
          ],
          driveLink: "1pG4NIUveTfv46DvQoq0TD4uuwaXCLwkS",
        };
      case "straus":
        return {
          recipients: ["araceli.farias@nisd.net", "brandy.bergeron@nisd.net"],
          driveLink: "15p9xqZoyikuVRk4ZVi7sUYYgoxL1PSew",
        };
      case "vale":
        return {
          recipients: [
            "jenna.bloom@nisd.net",
            "brenda.rayburg@nisd.net",
            "daniel.novosad@nisd.net",
            "mary.harrington@nisd.net",
            "janet.medina@nisd.net",
          ],
          driveLink: "1QoRQNEt7_gWT3PC_DPnDFjlT1XKVQ3sN",
        };
      case "zachry":
        return {
          recipients: [
            "Richard.DeLaGarza@nisd.net",
            "juliana.molina@nisd.net",
            "randolph.neuenfeldt@nisd.net",
            "jaclyn.galvan@nisd.net",
            "veronica.poblano@nisd.net",
            "monica.perez@nisd.net",
            "chris.benavidez@nisd.net",
          ],
          driveLink: "1UZDcETdHG5eN9DSCdfKcV0wt72cVY7Ek",
        };
      case "zachry magnet":
        return {
          recipients: ["mattew.patty@nisd.net"],
          driveLink: "1wMjhAx6wGOw5j-tu7-4wTNnH7V80pfPe",
        };
      case "test":
        return {
          recipients: ["alvaro.gomez@nisd.net"], //, "reggie.ollendieck@nisd.net"],
          driveLink: "1nMJAEcGIh_QnhfS5gjCkKd6CtoA3r5cf",
        };
      default:
        return {
          recipients: "",
          driveLink: "",
        };
    }
  }

  function getGmailTemplateFromDrafts_(row, driveLink) {
    return {
      message: {
        subject: "AEP Placement Transition Plan",
        html: `${row["Name"]} has nearly completed their assigned placement at NAMS and should be returning to ${row["Campus"]} on or around ${row["Return date"]}.<br><br>   
              On their last day of placement, they will be given withdrawal documents and the parents/guardians will have been called and told to contact ${row["Campus"]} to set up an appointment to re-enroll and meet with an administrator/counselor.<br><br>  
              Below are links and attachments to a Personalized Transition Plan (with notes from NAMS' assigned social worker), the student's AEP Transition Plan (with grades and notes from their teachers at NAMS), and a link to ${row["Campus"]}'s folder with all of the transition plans for this year.<br><br>
              Please let me know if you have any questions or concerns.<br><br>
              Thank you for all you do,<br>
              JD<br><br>
              <ul>
                <li><a href="${row["Merged Doc URL - Home Campus Transition Plan"]}">Personalized Transition Plan</a></li>
                <li><a href="${row["Merged Doc URL - Student Transition "]}">AEP Transition Plan</a></li>
                <li><a href="https://drive.google.com/drive/folders/${driveLink}">Drive Folder</a></li>
                <li><a href="https://drive.google.com/file/d/1qnyQ8cCxLVM9D6rg4wkyBp6KrXIELfNx/view?usp=sharing">Updates in Special Education</a></li>
              </ul>`,
      },
    };
  }

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
