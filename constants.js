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
