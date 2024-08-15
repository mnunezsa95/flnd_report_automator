function copyAndStoreFile() {
  const templateData = getBrandingTemplate();
  const constants = getConstants();

  try {
    const file = DriveApp.getFileById(templateData.reportTemplateID);

    const newFile = file.makeCopy(constants.newFileName);
    const newFileID = newFile.getId();
    const newFileUrl = newFile.getUrl();

    const targetFolderArray = constants.targetFolder.split("/");
    const targetFolderID = targetFolderArray[targetFolderArray.length - 2];

    const folder = DriveApp.getFolderById(targetFolderID);
    folder.addFile(newFile);

    DriveApp.getRootFolder().removeFile(newFile);

    const uniqueReportId = generateUniqueID();

    SpreadsheetApp.getUi().alert(`File copied and stored successfully as ${constants.newFileName}.\nReport ID: ${uniqueReportId}.`);

    return {
      newFileID: newFileID,
      newFileUrl: newFileUrl,
      uniqueReportId: uniqueReportId,
    };
  } catch (e) {
    SpreadsheetApp.getUi().alert(`Error: ${e.message}`);
  }
}
