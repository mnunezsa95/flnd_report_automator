const getConstants = () => {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = SpreadsheetApp.getActiveSheet();
  const cell = sheet.getActiveCell();
  const row = cell.getRow();
  const programName = sheet.getRange(row, 2).getValue();
  const targetFolder = sheet.getRange(row, 3).getValue();
  const reportTemplateURL = sheet.getRange(row, 4).getValue();
  const courseLevelsURL = sheet.getRange(row, 5).getValue();
  const resultsFolderURL = sheet.getRange(row, 6).getValue();
  const newFileName = sheet.getRange(row, 7).getValue();
  const reportYear = sheet.getRange(row, 8).getValue();

  return {
    spreadsheet: spreadsheet,
    sheet: sheet,
    row: row,
    programName: programName,
    reportTemplateURL: reportTemplateURL,
    newFileName: newFileName,
    reportYear: reportYear,
    resultsFolderURL: resultsFolderURL,
    courseLevelsURL: courseLevelsURL,
    targetFolder,
    targetFolder,
  };
};

function getBrandingTemplate() {
  const constants = getConstants();
  const reportTemplateArray = constants.reportTemplateURL.split("/");
  const reportTemplateID = reportTemplateArray[reportTemplateArray.length - 2];
  SpreadsheetApp.getUi().alert(`Report Template ID for ${constants.programName} (row ${constants.row}):\n ${reportTemplateID}`);

  return {
    reportTemplateID: reportTemplateID,
  };
}

function getResultsFolder() {
  const constants = getConstants();
  const resultsFolderIDs = constants.resultsFolderURL.split("/").filter((item) => item.startsWith("1"));

  try {
    const formattedStr = resultsFolderIDs.map((item, i) => `${item}\n`).join(" ");
    SpreadsheetApp.getUi().alert(`M&E Results Folder IDs for ${constants.programName} (row ${constants.row}):\n ${formattedStr}`);
  } catch {
    SpreadsheetApp.getUi().alert(`Error: Could not extract folder ID from URL:\n ${constants.resultsFolderURL}`);
  }

  return resultsFolderIDs;
}

function getTableFile() {
  const constants = getConstants();
  const courseLevelsFileArray = constants.courseLevelsURL.split("/");
  let courseLevelsFileID = null;

  try {
    courseLevelsFileID = courseLevelsFileArray[courseLevelsFileArray.length - 2];
    SpreadsheetApp.getUi().alert(`Course Levels File ID for ${constants.programName} (row ${constants.row}):\n ${courseLevelsFileID}`);
  } catch {
    SpreadsheetApp.getUi().alert(`Error: Could not extract folder ID from URL:\n ${constants.courseLevelsURL}`);
  }

  return courseLevelsFileID;
}

function getImageFiles() {
  const resultsFolderIDs = getResultsFolder();
  const fileIdMap = {};
  let numberOfFiles;

  if (!resultsFolderIDs || resultsFolderIDs.length === 0) {
    SpreadsheetApp.getUi().alert("No valid M&E folder IDs found.");
    return {};
  }

  resultsFolderIDs.forEach((folderID) => {
    try {
      const folder = DriveApp.getFolderById(folderID);
      const folderName = folder.getName();
      const files = folder.getFiles();

      while (files.hasNext()) {
        const file = files.next();
        if (file.getMimeType() === MimeType.PNG) {
          const fullName = file.getName();
          const nameParts = fullName.split(" - ");

          if (nameParts.length > 1) {
            let key;
            if (folderName.includes("Progressive")) {
              key = `<${nameParts[0]} - Progressive>`;
            } else if (folderName.includes("JSS")) {
              key = `<${nameParts[0]} - JSS>`;
            } else if (folderName.includes("Primary")) {
              key = `<${nameParts[0]} - Primary>`;
            } else {
              key = `<${nameParts[0]}>`;
            }
            const fileId = file.getId();
            fileIdMap[key] = fileId;
          }
        }
      }
    } catch (e) {
      Logger.log("Error processing folder ID " + folderID + ": " + e.message);
    }
  });

  numberOfFiles = Object.keys(fileIdMap).length;
  if (Object.keys(fileIdMap).length > 0) {
    let alertMessage = `Found ${numberOfFiles} .png files:\n\n`;
    for (const key in fileIdMap) {
      alertMessage += `${key}: ${fileIdMap[key]}\n`;
    }
    SpreadsheetApp.getUi().alert(alertMessage);
  } else {
    SpreadsheetApp.getUi().alert("No .png files found in the folders.");
  }

  return fileIdMap;
}
