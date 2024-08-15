function creatTable(body, data, headers, isNumeracy, placeholder) {
  // Define number of columns based on whether it's numeracy or literacy
  const numColumns = isNumeracy ? 7 : 9;

  // Find the placeholder text
  let searchResult = body.findText(placeholder);

  while (searchResult) {
    const element = searchResult.getElement();
    const start = searchResult.getStartOffset();
    const end = searchResult.getEndOffsetInclusive();

    // Insert the table below the placeholder text
    const parent = element.getParent();
    const index = body.getChildIndex(parent);

    // Insert the table after the placeholder
    const table = body.insertTable(index + 1);

    // Add header row
    const headerRow = table.appendTableRow();
    headers.forEach((header) => {
      const cell = headerRow.appendTableCell(header);
      cell.setFontSize(9);
      cell.setFontFamily("Arial");
      cell.setBackgroundColor("#efefef");
      cell.setBold(true);
    });

    const tempData = {};

    // First pass: Accumulate values
    data.forEach((item) => {
      if (isNumeracy) {
        if (!tempData[item.Grade]) {
          tempData[item.Grade] = {
            Grade: item.Grade,
            Level: item.Level || "",
            ItemAvg1: "", // Placeholder
            ItemAvg2: "", // Placeholder
            RemediateLevel: "", // Placeholder
            ProposedLevel: "", // Placeholder
            Custom: "", // Placeholder
          };
        }
      } else {
        // Initialize or update temporary data
        if (!tempData[item.Grade]) {
          tempData[item.Grade] = {
            Grade: item.Grade,
            LastYearReadingLevel: "",
            LastYearLangLevel: "",
            ItemAvg1: "", // Placeholder
            ItemAvg2: "", // Placeholder
            RemediateLevel: "", // Placeholder
            ProposedReadingLevel: "", // Placeholder
            ProposedLangLevel: "", // Placeholder
            Custom: "", // Placeholder
          };
        }

        if (
          item.Subject === "English Studies - Reading 1 & 2" ||
          item.Subject === "Literacy Revision 1 & 2" ||
          item.Subject === "Literacy 1 & 2" ||
          item.Subject === "Supplementary English 1 & 2" ||
          item.Subject === "English Literacy Revision 1 & 2"
        ) {
          tempData[item.Grade].LastYearReadingLevel = item.Level;
        }

        if (item.Subject === "English Studies - Language" || item.Subject === "Supplemental Language") {
          tempData[item.Grade].LastYearLangLevel = item.Level;
        }
      }
    });

    // Convert tempData to an array
    const processedData = Object.values(tempData);

    // Debugging: log processed data to verify correctness
    console.log("Processed Data:", processedData);

    // Add data rows
    processedData.forEach((item) => {
      const row = table.appendTableRow();
      row.appendTableCell(item.Grade || ""); // Populate "Grade" column

      if (isNumeracy) {
        row.appendTableCell(item.Level || ""); // Populate "Last Year’s Levels" column
        row.appendTableCell(item.ItemAvg1); // Populate "On-level Item Avg (1)"
        row.appendTableCell(item.ItemAvg2); // Populate "On-level Item Avg (2)"
        row.appendTableCell(item.RemediateLevel); // Populate "Level to Remediate"
        row.appendTableCell(item.ProposedLevel); // Populate "Proposed Level"
      } else {
        row.appendTableCell(item.LastYearReadingLevel); // Populate "Last Year’s Reading Levels"
        row.appendTableCell(item.LastYearLangLevel); // Populate "Last Year’s Lang Levels"
        row.appendTableCell(item.ItemAvg1); // Populate "On-level Item Avg (1)"
        row.appendTableCell(item.ItemAvg2); // Populate "On-level Item Avg (2)"
        row.appendTableCell(item.RemediateLevel); // Populate "Level to Remediate"
        row.appendTableCell(item.ProposedReadingLevel); // Populate "Proposed Reading Level"
        row.appendTableCell(item.ProposedLangLevel); // Populate "Proposed Lang Level"
      }

      row.appendTableCell(item.Custom || ""); // Populate "Custom" column
    });

    const rows = table.getNumRows();
    for (let i = 1; i < rows; i++) {
      const row = table.getRow(i);
      const cells = row.getNumCells();
      for (let j = 0; j < cells; j++) {
        const cell = row.getCell(j);
        cell.setFontSize(8);
        cell.setFontFamily("Arial");
        cell.setBold(false);
        cell.setForegroundColor("#666666");
      }
    }

    // Remove the placeholder text
    element.asText().deleteText(start, end);

    // Move to the next occurrence of the placeholder
    searchResult = body.findText(placeholder, searchResult);
  }
}
