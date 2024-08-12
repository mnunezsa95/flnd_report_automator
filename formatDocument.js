function formatDocument(newFileData) {
  const constants = getConstants();
  const fileIdMap = getImageFiles(); 
  const { numeracyData, literacyData } = extractLevelsInfo(); 
  const docTitle = constants.newFileName;
  const reportForYear = constants.reportYear;
  const resultsURL = constants.resultsFolderURL;
  const newFileID = newFileData.newFileID;
  let doc; // Declare `doc` outside the try block

  if (!newFileID) {
    SpreadsheetApp.getUi().alert('Error: Document ID is not available.');
    throw new Error('Document ID is not available.');
  }

  try {
    doc = DocumentApp.openById(newFileID); 
    const body = doc.getBody();

    body.replaceText('<ReportTitle>', docTitle);
    body.replaceText('<YYYY>', reportForYear);
    body.replaceText('<results>', 'results');

    // Find and hyperlink only the text 'results'
    let searchResult = body.findText('results');
    while (searchResult) {
      const rangeElement = searchResult.getElement();
      const start = searchResult.getStartOffset();
      const end = searchResult.getEndOffsetInclusive();
      rangeElement.asText().setLinkUrl(start, end, resultsURL);
      searchResult = body.findText('results', searchResult);
    }

    // Replace image placeholders with actual images
    for (const key in fileIdMap) {
      if (fileIdMap.hasOwnProperty(key)) {
        const placeholder = key;
        const imageId = fileIdMap[key];
        
        let imageResult = body.findText(placeholder);
        while (imageResult) {
          const element = imageResult.getElement();
          const start = imageResult.getStartOffset();
          const end = imageResult.getEndOffsetInclusive();
          
          const imageBlob = DriveApp.getFileById(imageId).getBlob();
          const parent = element.getParent();
          const image = body.insertImage(body.getChildIndex(parent) + 1, imageBlob);
          
          // Resize image based on content or page width
          image.setWidth(500); // Adjust as needed
          image.setHeight(400); // Adjust as needed

          element.asText().deleteText(start, end);
          imageResult = body.findText(placeholder, imageResult);
        }
      }
    }

    // Insert tables at placeholders
    createTableAtPlaceholder(body, numeracyData, ['Grade', 'Current Years\' Levels', 'On-level Avg', '2025 Numeracy Level', 'Repeating'], true, '<Numeracy Table>');
    createTableAtPlaceholder(body, literacyData, ['Grade', 'Current Years\' Levels', 'On-level Avg', '2025 Lit Level', '2025 Lang', 'Repeating'], false, '<Literacy Table>');

  } catch (e) {
    SpreadsheetApp.getUi().alert(`Error: ${e.message}`);
    throw e;
    
  } finally {
    if (doc) {
      doc.saveAndClose(); 
    }
  }
}

