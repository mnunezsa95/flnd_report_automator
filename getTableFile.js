function getTableFile() {
  const constants = getConstants()
  const courseLevelsFileArray = constants.courseLevelsURL.split("/")
  
  let courseLevelsFileID = null

  try {
    courseLevelsFileID = courseLevelsFileArray[courseLevelsFileArray.length - 2]
    SpreadsheetApp.getUi().alert(`Course Levels File ID for ${constants.programName} (row ${constants.row}):\n ${courseLevelsFileID}`);
  } catch {
    SpreadsheetApp.getUi().alert(`Error: Could not extract folder ID from URL:\n ${constants.courseLevelsURL}`)
  }

  return courseLevelsFileID
}
