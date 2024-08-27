function createTable(body, data, headers, isNumeracy, placeholder) {
  // Define number of columns based on whether it's numeracy or literacy
  const numColumns = isNumeracy ? 7 : 9;

  // Define constants for styling
  const GREEN = "#d9ead3";
  const LIGHT_YELLOW = "#fff2cc";
  const YELLOW = "#ffd966";
  const ORANGE = "#f6b26b";
  const RED = "#ea9999";
  const FONT_SIZE = 8;
  const FONT_FAMILY = "Arial";
  const HEADER_BG_COLOR = "#efefef";
  const FONT_COLOR = "#666666";

  // Find the placeholder text in the document
  let searchResult = body.findText(placeholder);

  while (searchResult) {
    const element = searchResult.getElement();
    const start = searchResult.getStartOffset();
    const end = searchResult.getEndOffsetInclusive();
    const parent = element.getParent();
    const index = body.getChildIndex(parent);

    // Insert the table after the placeholder text
    const table = body.insertTable(index + 1);

    // Add header row to the table
    const headerRow = table.appendTableRow();
    headers.forEach((header) => {
      const cell = headerRow.appendTableCell(header);
      cell.setFontSize(FONT_SIZE).setFontFamily(FONT_FAMILY).setBackgroundColor(HEADER_BG_COLOR).setBold(true);
    });

    // Accumulate data by grade
    const tempData = {};

    data.forEach((item) => {
      if (!tempData[item.Grade]) {
        tempData[item.Grade] = {
          Grade: item.Grade,
          Level: item.Level || "",
          ItemAvg1: "", // Placeholder
          ItemAvg2: "", // Placeholder
          RemediateLevel: "", // Placeholder
          ProposedLevel: "", // Placeholder
          Custom: "", // Placeholder
          LastYearReadingLevel: "",
          LastYearLangLevel: "",
          ProposedReadingLevel: "",
          ProposedLangLevel: "",
        };
      }

      if (!isNumeracy) {
        // Literacy specific data accumulation
        if (
          [
            "English Studies - Reading 1 & 2",
            "Literacy Revision 1 & 2",
            "Literacy 1 & 2",
            "Supplementary English 1 & 2",
            "English Literacy Revision 1 & 2",
          ].includes(item.Subject)
        ) {
          tempData[item.Grade].LastYearReadingLevel = item.Level;
        }
        if (["English Studies - Language", "Supplemental Language"].includes(item.Subject)) {
          tempData[item.Grade].LastYearLangLevel = item.Level;
        }
      }
    });

    // Convert accumulated data into an array
    const processedData = Object.values(tempData);

    // Add data rows to the table
    processedData.forEach((item) => {
      const row = table.appendTableRow();
      row.appendTableCell(item.Grade || "");

      if (isNumeracy) {
        row.appendTableCell(item.Level || "");
        row.appendTableCell(item.ItemAvg1);
        row.appendTableCell(item.ItemAvg2);
        row.appendTableCell(item.RemediateLevel);
        row.appendTableCell(item.ProposedLevel);
      } else {
        const lastYearReadingCell = row.appendTableCell(item.LastYearReadingLevel);
        row.appendTableCell(item.LastYearLangLevel);
        row.appendTableCell(item.ItemAvg1);
        row.appendTableCell(item.ItemAvg2);
        row.appendTableCell(item.RemediateLevel);
        row.appendTableCell(item.ProposedReadingLevel);
        row.appendTableCell(item.ProposedLangLevel);

        // Apply conditional formatting for reading levels
        const readingColors = {
          2: {
            A: LIGHT_YELLOW,
            AA: LIGHT_YELLOW,
            AB: GREEN,
            B: GREEN,
            BB: GREEN,
          },
          3: {
            A: YELLOW,
            AA: YELLOW,
            AB: LIGHT_YELLOW,
            B: LIGHT_YELLOW,
            BB: LIGHT_YELLOW,
            BC: GREEN,
            C: GREEN,
            CC: GREEN,
          },
          4: {
            A: ORANGE,
            AA: ORANGE,
            AB: YELLOW,
            B: YELLOW,
            BB: YELLOW,
            BC: LIGHT_YELLOW,
            C: LIGHT_YELLOW,
            CC: LIGHT_YELLOW,
            CD: GREEN,
            D: GREEN,
            DD: GREEN,
          },
          5: {
            A: ORANGE,
            AA: ORANGE,
            AB: YELLOW,
            B: YELLOW,
            BB: YELLOW,
            BC: LIGHT_YELLOW,
            C: LIGHT_YELLOW,
            CC: LIGHT_YELLOW,
            CD: GREEN,
            D: GREEN,
            DD: GREEN,
          },
          6: {
            A: ORANGE,
            AA: ORANGE,
            AB: YELLOW,
            B: YELLOW,
            BB: YELLOW,
            BC: LIGHT_YELLOW,
            C: LIGHT_YELLOW,
            CC: LIGHT_YELLOW,
            CD: GREEN,
            D: GREEN,
            DD: GREEN,
          },
        };

        const gradeSuffix = item.Grade.slice(-1);
        const readingLevel = item.LastYearReadingLevel;

        if (readingColors[gradeSuffix]) {
          const color =
            readingColors[gradeSuffix][readingLevel] ||
            (readingLevel.includes("A") && readingLevel.includes("B") && readingColors[gradeSuffix]["AB"]) ||
            (readingLevel.includes("B") && readingLevel.includes("C") && readingColors[gradeSuffix]["BC"]) ||
            (readingLevel.includes("C") && readingLevel.includes("D") && readingColors[gradeSuffix]["CD"]) ||
            (readingLevel.includes("A") && readingColors[gradeSuffix]["AA"]) ||
            (readingLevel.includes("B") && readingColors[gradeSuffix]["BB"]) ||
            (readingLevel.includes("C") && readingColors[gradeSuffix]["CC"]) ||
            (readingLevel.includes("D") && readingColors[gradeSuffix]["DD"]);
          if (color) lastYearReadingCell.setBackgroundColor(color);
        }
      }

      row.appendTableCell(item.Custom || "");
    });

    // Apply consistent styling to all table cells
    const rows = table.getNumRows();
    for (let i = 1; i < rows; i++) {
      const row = table.getRow(i);
      const cells = row.getNumCells();
      for (let j = 0; j < cells; j++) {
        const cell = row.getCell(j);
        cell.setFontSize(FONT_SIZE).setFontFamily(FONT_FAMILY).setBold(false).setForegroundColor(FONT_COLOR);
      }
    }

    // Remove the placeholder text
    element.asText().deleteText(start, end);

    // Move to the next occurrence of the placeholder
    searchResult = body.findText(placeholder, searchResult);
  }
}
