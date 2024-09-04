// Docs: https://github.com/google/clasp

function main() {
  const constants = getConstants();
  const newFileData = copyAndStoreFile();
  const newFileUrl = newFileData.newFileUrl;
  const ui = SpreadsheetApp.getUi();

  const projectDirectorEmail = "marlon.nunez@newglobe.education";
  const currentUserEmail = Session.getActiveUser().getEmail();
  const emailParts = currentUserEmail.split("@")[0];
  const nameParts = emailParts.split(".");
  const formattedName = nameParts.map((part) => part.charAt(0).toUpperCase() + part.slice(1).toLowerCase()).join(" ");

  const recipientsList = [currentUserEmail, projectDirectorEmail];

  const startTime = new Date().getTime();
  const reportDatabaseTab = constants.spreadsheet.getSheetByName("Report_Database");

  const defaultOutputData = [
    newFileData.uniqueReportId,
    constants.programName,
    constants.newFileName,
    "N/A",
    "N/A",
    "N/A",
    "N/A",
    "N/A",
  ];

  let outputData = [...defaultOutputData];
  let executionTime;
  let currentDate;

  try {
    // Get the current date and time in a readable format
    currentDate = new Date().toLocaleString("en-US", {
      weekday: "short",
      year: "numeric",
      month: "long",
      day: "numeric",
      timeZoneName: "short",
      hour12: false,
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
    });

    outputData[3] = currentDate;
    outputData[5] = formattedName;
    outputData[6] = currentUserEmail;
    outputData[7] = `=HYPERLINK("${newFileUrl}", "Link")`;

    formatDocument(newFileData);

    ui.alert(
      `Report Creation Successful!\n\nExecution Time: ${executionTime} milliseconds.\n\nSee additional details in the "Report_Database" Tab`,
      ui.ButtonSet.OK
    );
  } catch (err) {
    outputData[3] = "Report Creation Failed";
    outputData[7] = "N/A";

    const errorEmailSubject = `Report Creation Unsuccessful`;
    const errorEmailBody = `
      An error occurred when creating an FLND Report.\n
      Program Name: ${constants.programName}.
      Executed By: ${formattedName}
      Report Name: ${constants.newFileName}
      Report ID: ${newFileData.uniqueReportId}
      Error Message: ${err}.`;

    sendEmail(recipientsList, errorEmailSubject, errorEmailBody);
    ui.alert(`Report Creation Unsuccessful!\n\nSee additional details in email sent to your inbox.`, ui.ButtonSet.OK);
  } finally {
    executionTime = new Date().getTime() - startTime;
    outputData[4] = executionTime;

    const nextRow = reportDatabaseTab.getLastRow() + 1;
    reportDatabaseTab.getRange(nextRow, 2, 1, outputData.length).setValues([outputData]);

    // Reset fields (optional)
    // resetFields(constants.sheet, constants.row, [6, 7, 8]);
  }
}
