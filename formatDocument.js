function formatDocument(newFileData) {
  const constants = getConstants();
  const fileIdMap = getImageFiles();
  const { numeracyData, literacyData } = extractLevelsInfo();
  const docTitle = constants.newFileName;
  const reportForYear = constants.reportYear;
  const resultsURL = constants.resultsFolderURL;
  const newFileID = newFileData.newFileID;
  let doc;

  if (!newFileID) {
    SpreadsheetApp.getUi().alert("Error: Document ID is not available.");
    throw new Error("Document ID is not available.");
  }

  try {
    doc = DocumentApp.openById(newFileID);
    const body = doc.getBody();

    body.replaceText("<ReportTitle>", docTitle);
    body.replaceText("<YYYY>", reportForYear);
    body.replaceText("<results>", "results");

    // Find and hyperlink only the text 'results'
    let searchResult = body.findText("results");
    while (searchResult) {
      const rangeElement = searchResult.getElement();
      const start = searchResult.getStartOffset();
      const end = searchResult.getEndOffsetInclusive();
      rangeElement.asText().setLinkUrl(start, end, resultsURL);
      searchResult = body.findText("results", searchResult);
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
          image.setWidth(400); // Adjust as needed
          image.setHeight(300); // Adjust as needed

          element.asText().deleteText(start, end);
          imageResult = body.findText(placeholder, imageResult);
        }
      }
    }

    // Iterate over keys of the data object to create tables
    for (const key in numeracyData) {
      if (numeracyData.hasOwnProperty(key)) {
        const data = numeracyData[key];
        createTable(
          body,
          data,
          ["Rising Grade", "Last Year’s Levels", "On-level Item Avg (1)", "On-level Item Avg (2)", "Level to Remediate", "Proposed Level", "Custom"],
          true,
          `<Numeracy Table - ${key}>`
        );
      }
    }

    for (const key in literacyData) {
      if (literacyData.hasOwnProperty(key)) {
        const data = literacyData[key];
        createTable(
          body,
          data,
          [
            "Rising Grade",
            "Last Year’s Reading Levels",
            "Last Year’s Lang Levels",
            "On-level Item Avg (1)",
            "On-level Item Avg (2)",
            "Level to Remediate",
            "Proposed Reading Level",
            "Proposed Lang Level",
            "Custom",
          ],
          false,
          `<Literacy Table - ${key}>`
        );
      }
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert(`Error: ${e.message}`);
    throw e;
  } finally {
    if (doc) {
      doc.saveAndClose();
    }
  }
}
