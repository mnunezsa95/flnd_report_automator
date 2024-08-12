function getResultsFolder() {
  const constants = getConstants();
  const resultsFolderArray = constants.resultsFolderURL.split("/");

  let resultsFolderID = null
  try {
    resultsFolderID = resultsFolderArray[resultsFolderArray.length - 2];
    SpreadsheetApp.getUi().alert(`M&E Results Folder ID for ${constants.programName} (row ${constants.row}):\n ${resultsFolderID}`);
  } catch {
    SpreadsheetApp.getUi().alert(`Error: Could not extract folder ID from URL:\n ${constants.resultsFolderURL}`);
  }

  return resultsFolderID;
}