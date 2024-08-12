function replacePlaceholders() {
  const newFileData = copyAndStoreFile();
  const constants = getConstants();
  const fileIdMap = getImageFiles(); 

  const docTitle = constants.newFileName;
  const reportForYear = constants.reportYear;
  const resultsURL = constants.resultsFolderURL;
  const newFileUrl = newFileData.newFileUrl;
  const newFileID = newFileData.newFileID;

  const startTime = new Date().getTime();
  const reportDatabaseTab = constants.spreadsheet.getSheetByName('Report_Database');


  const defaultOutputData = [
    newFileData.uniqueReportId, 
    constants.programName, 'N/A', 'N/A', 'N/A', 'N/A', 'N/A'];

  let outputData = [...defaultOutputData]; // Initialize with default values
  let executionTime;

  try {
    // Get the current date and time in a readable format
    const currentDate = new Date().toLocaleString('en-US', {
      weekday: 'short',
      year: 'numeric',
      month: 'long',
      day: 'numeric',
      timeZoneName: 'short',
      hour12: false, 
      hour: '2-digit', 
      minute: '2-digit', 
      second: '2-digit'
    });

    // Extract the email from the session
    const currentUserEmail = Session.getActiveUser().getEmail();
    
    const emailParts = currentUserEmail.split('@')[0]; // Get the part before '@'
    const nameParts = emailParts.split('.'); // Split by '.'
    const formattedName = nameParts.map(part => part.charAt(0).toUpperCase() + part.slice(1).toLowerCase()).join(' '); 

    outputData[2] = currentDate; 
    outputData[4] = formattedName; 
    outputData[5] = currentUserEmail;
    
    // Set the hyperlink formula
    outputData[6] = `=HYPERLINK("${newFileUrl}", "Link")`;

    const { numeracyData, literacyData } = extractLevelsInfo(); 

    if (newFileID) {
      try {
        const doc = DocumentApp.openById(newFileID);
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
            
            // Find each placeholder and replace with the image
            let imageResult = body.findText(placeholder);
            while (imageResult) {
              const element = imageResult.getElement();
              const start = imageResult.getStartOffset();
              const end = imageResult.getEndOffsetInclusive();
              
              // Fetch the image from Drive and resize
              const imageBlob = DriveApp.getFileById(imageId).getBlob();
                          
              // Insert the image below the placeholder text
              const parent = element.getParent();
              const image = body.insertImage(body.getChildIndex(parent) + 1, imageBlob);
              
              // Resize the image to fit within the page
              image.setWidth(500); // Adjust width as needed
              image.setHeight(400); // Adjust height as needed

              // Remove the placeholder text
              element.asText().deleteText(start, end);
              
              // Move to the next occurrence of the placeholder
              imageResult = body.findText(placeholder, imageResult);
            }
          }
        }

        // Insert tables at placeholders
        createTableAtPlaceholder(body, numeracyData, ['Grade', 'Current Years\' Levels', 'On-level Avg', '2025 Numeracy Level', 'Repeating'], true, '<Numeracy Table>');
        createTableAtPlaceholder(body, literacyData, ['Grade', 'Current Years\' Levels', 'On-level Avg', '2025 Lit Level', '2025 Lang', 'Repeating'], false, '<Literacy Table>');

        // Save and close the document
        doc.saveAndClose();
        
      } catch (e) {
        SpreadsheetApp.getUi().alert(`Error: ${e.message}`);
        throw e; // Re-throw to ensure the finally block executes
      }
    } else {
      SpreadsheetApp.getUi().alert('Error: Document ID is not available.');
      throw new Error('Document ID is not available.');
    }

  } catch (err) {
    console.log(err);
    outputData[2] = "Report Creation Failed"; 
    outputData[6] = 'N/A'; 
  } finally {
    executionTime = new Date().getTime() - startTime;
    outputData[3] = executionTime; 

    const nextRow = reportDatabaseTab.getLastRow() + 1;
    reportDatabaseTab.getRange(nextRow, 2, 1, outputData.length).setValues([outputData]);
    
    resetFields(constants.sheet, constants.row, [6, 7,8])
  }
}

