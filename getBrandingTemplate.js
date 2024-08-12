function getBrandingTemplate() {
  const constants = getConstants();
  const reportTemplateArray = constants.reportTemplateURL.split("/");
  const reportTemplateID = reportTemplateArray[reportTemplateArray.length - 2];
  
  SpreadsheetApp.getUi().alert(`Report Template ID for ${constants.programName} (row ${constants.row}):\n ${reportTemplateID}`);

  return {
    reportTemplateID: reportTemplateID,
  };
}