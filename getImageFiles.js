function getImageFiles() {
  const resultsFolderID = getResultsFolder(); 
  
  if (!resultsFolderID) {
    SpreadsheetApp.getUi().alert('No valid M&E folder ID found.');
    return {};
  }
  
  const folder = DriveApp.getFolderById(resultsFolderID);
  const files = folder.getFiles();
  const fileIdMap = {};
  
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() === MimeType.PNG) {
      const fullName = file.getName();
      const nameParts = fullName.split(' - ');

      if (nameParts.length > 1) {
        const key = `<${nameParts[0]}>`; // Add < and > around the key
        const fileId = file.getId();
        fileIdMap[key] = fileId;
      }
    }
  }
  
  // Optionally, alert the results
  if (Object.keys(fileIdMap).length > 0) {
    let alertMessage = 'Found .png files:\n';
    for (const key in fileIdMap) {
      alertMessage += `${key}: ${fileIdMap[key]}\n`;
    }
    SpreadsheetApp.getUi().alert(alertMessage);
  } else {
    SpreadsheetApp.getUi().alert('No .png files found in the folder.');
  }
  
  return fileIdMap;
}

