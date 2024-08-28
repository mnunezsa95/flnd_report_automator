// Docs: https://github.com/google/clasp

function main() {
  const constants = getConstants();
  const newFileData = copyAndStoreFile();
  const newFileUrl = newFileData.newFileUrl;

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

  let outputData = [...defaultOutputData]; // Initialize with default values
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

    // Extract the email from the session
    const currentUserEmail = Session.getActiveUser().getEmail();

    const emailParts = currentUserEmail.split("@")[0]; // Get the part before '@'
    const nameParts = emailParts.split("."); // Split by '.'
    const formattedName = nameParts.map((part) => part.charAt(0).toUpperCase() + part.slice(1).toLowerCase()).join(" ");

    outputData[3] = currentDate;
    outputData[5] = formattedName;
    outputData[6] = currentUserEmail;

    // Set the hyperlink formula
    outputData[7] = `=HYPERLINK("${newFileUrl}", "Link")`;

    // Call formatDocument and ensure it doesn't throw an error
    formatDocument(newFileData);
  } catch (err) {
    console.log(err);
    outputData[3] = "Report Creation Failed";
    outputData[7] = "N/A";
  } finally {
    executionTime = new Date().getTime() - startTime;
    outputData[4] = executionTime;

    const nextRow = reportDatabaseTab.getLastRow() + 1;
    reportDatabaseTab.getRange(nextRow, 2, 1, outputData.length).setValues([outputData]);

    // Reset fields (optional)
    // resetFields(constants.sheet, constants.row, [6, 7, 8]);

    // Display message to the user
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      `Report Executed Successfully!\n\nExecution Time: ${executionTime} milliseconds.\n\nSee additional details in the "Report_Database" Tab`,
      ui.ButtonSet.OK
    );
  }
}
