function createTable(body, data, headers, isNumeracy, placeholder) {
  const numColumns = isNumeracy ? 7 : 9;

  const GREEN = "#d9ead3";
  const LIGHT_YELLOW = "#fff2cc";
  const YELLOW = "#ffd966";
  const ORANGE = "#f6b26b";
  const RED = "#ea9999";
  const FONT_SIZE = 8;
  const FONT_FAMILY = "Arial";
  const HEADER_BG_COLOR = "#efefef";
  const DEFAULT_FONT_COLOR = "#666666";
  const LOW_COLOR = "#990000"; // 0-39%
  const MEDIUM_COLOR = "#e69138"; // 40-69%
  const HIGH_COLOR = "#6aa84f"; // 70-100%

  let searchResult = body.findText(placeholder);

  while (searchResult) {
    const element = searchResult.getElement();
    const start = searchResult.getStartOffset();
    const end = searchResult.getEndOffsetInclusive();
    const parent = element.getParent();
    const index = body.getChildIndex(parent);

    const table = body.insertTable(index + 1);

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
          ItemAvg1: item.avgOne ? item.avgOne + "%" : "",
          ItemAvg2: item.avgTwo ? item.avgTwo + "%" : "",
          RemediateLevel: "",
          ProposedLevel: "",
          Custom: "",
          LastYearReadingLevel: "",
          LastYearLangLevel: "",
          ProposedReadingLevel: "",
          ProposedLangLevel: "",
        };
      }

      if (item.avgOne && !tempData[item.Grade].ItemAvg1) {
        tempData[item.Grade].ItemAvg1 = item.avgOne + "%";
      }

      if (item.avgTwo && !tempData[item.Grade].ItemAvg2) {
        tempData[item.Grade].ItemAvg2 = item.avgTwo + "%";
      }

      if (!isNumeracy) {
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

    const processedData = Object.values(tempData);

    processedData.forEach((item) => {
      const row = table.appendTableRow();
      row.appendTableCell(item.Grade || "");

      if (isNumeracy) {
        const lastYearMathCell = row.appendTableCell(item.Level);
        const itemAvg1Cell = row.appendTableCell(item.ItemAvg1);
        const itemAvg2Cell = row.appendTableCell(item.ItemAvg2);
        row.appendTableCell(item.RemediateLevel);
        row.appendTableCell(item.ProposedLevel);

        applyPercentageColor(itemAvg1Cell, item.ItemAvg1, [LOW_COLOR, MEDIUM_COLOR, HIGH_COLOR]);
        applyPercentageColor(itemAvg2Cell, item.ItemAvg2, [LOW_COLOR, MEDIUM_COLOR, HIGH_COLOR]);

        // Apply conditional formatting for math levels
        const mathColors = {
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
            A: RED,
            AA: RED,
            AB: ORANGE,
            B: ORANGE,
            BB: ORANGE,
            BC: YELLOW,
            C: YELLOW,
            CC: YELLOW,
            CD: LIGHT_YELLOW,
            D: LIGHT_YELLOW,
            DD: LIGHT_YELLOW,
            DE: GREEN,
            E: GREEN,
            EE: GREEN,
          },
          6: {
            A: RED,
            AA: RED,
            AB: ORANGE,
            B: ORANGE,
            BB: ORANGE,
            BC: YELLOW,
            C: YELLOW,
            CC: YELLOW,
            CD: LIGHT_YELLOW,
            D: LIGHT_YELLOW,
            DD: LIGHT_YELLOW,
            DE: GREEN,
            E: GREEN,
            EE: GREEN,
          },
          12: {
            A: LIGHT_YELLOW,
            AA: LIGHT_YELLOW,
            AB: GREEN,
            B: GREEN,
            BB: GREEN,
          },
          34: {
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
          56: {
            A: RED,
            AA: RED,
            AB: ORANGE,
            B: ORANGE,
            BB: ORANGE,
            BC: YELLOW,
            C: YELLOW,
            CC: YELLOW,
            CD: LIGHT_YELLOW,
            D: LIGHT_YELLOW,
            DD: LIGHT_YELLOW,
            DE: GREEN,
            E: GREEN,
            EE: GREEN,
          },
        };

        const gradeSuffixMatch = item.Grade.match(/\d+/);
        const gradeSuffix = gradeSuffixMatch ? gradeSuffixMatch[0] : null;
        const mathLevel = item.Level;

        if (gradeSuffix) {
          if (mathColors[gradeSuffix]) {
            const color =
              mathColors[gradeSuffix][mathLevel] ||
              (mathLevel.includes("A") && mathLevel.includes("B") && mathColors[gradeSuffix]["AB"]) ||
              (mathLevel.includes("B") && mathLevel.includes("C") && mathColors[gradeSuffix]["BC"]) ||
              (mathLevel.includes("C") && mathLevel.includes("D") && mathColors[gradeSuffix]["CD"]) ||
              (mathLevel.includes("D") && mathLevel.includes("E") && mathColors[gradeSuffix]["DE"]) ||
              (mathLevel.includes("A") && mathColors[gradeSuffix]["AA"]) ||
              (mathLevel.includes("B") && mathColors[gradeSuffix]["BB"]) ||
              (mathLevel.includes("C") && mathColors[gradeSuffix]["CC"]) ||
              (mathLevel.includes("D") && mathColors[gradeSuffix]["DD"]) ||
              (mathLevel.includes("E") && mathColors[gradeSuffix]["EE"]);
            if (color) lastYearMathCell.setBackgroundColor(color);
          }
        }
      } else {
        const lastYearReadingCell = row.appendTableCell(item.LastYearReadingLevel);
        const lastYearLangCell = row.appendTableCell(item.LastYearLangLevel);
        const itemAvg1Cell = row.appendTableCell(item.ItemAvg1);
        const itemAvg2Cell = row.appendTableCell(item.ItemAvg2);
        row.appendTableCell(item.RemediateLevel);
        row.appendTableCell(item.ProposedReadingLevel);
        row.appendTableCell(item.ProposedLangLevel);

        applyPercentageColor(itemAvg1Cell, item.ItemAvg1, [LOW_COLOR, MEDIUM_COLOR, HIGH_COLOR]);
        applyPercentageColor(itemAvg2Cell, item.ItemAvg2, [LOW_COLOR, MEDIUM_COLOR, HIGH_COLOR]);

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
            A: RED,
            AA: RED,
            AB: ORANGE,
            B: ORANGE,
            BB: ORANGE,
            BC: YELLOW,
            C: YELLOW,
            CC: YELLOW,
            CD: LIGHT_YELLOW,
            D: LIGHT_YELLOW,
            DD: LIGHT_YELLOW,
            DE: GREEN,
            E: GREEN,
            EE: GREEN,
          },
          6: {
            A: RED,
            AA: RED,
            AB: ORANGE,
            B: ORANGE,
            BB: ORANGE,
            BC: YELLOW,
            C: YELLOW,
            CC: YELLOW,
            CD: LIGHT_YELLOW,
            D: LIGHT_YELLOW,
            DD: LIGHT_YELLOW,
            DE: GREEN,
            E: GREEN,
            EE: GREEN,
          },
          12: {
            A: LIGHT_YELLOW,
            AA: LIGHT_YELLOW,
            AB: GREEN,
            B: GREEN,
            BB: GREEN,
          },
          34: {
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
          56: {
            A: RED,
            AA: RED,
            AB: ORANGE,
            B: ORANGE,
            BB: ORANGE,
            BC: YELLOW,
            C: YELLOW,
            CC: YELLOW,
            CD: LIGHT_YELLOW,
            D: LIGHT_YELLOW,
            DD: LIGHT_YELLOW,
            DE: GREEN,
            E: GREEN,
            EE: GREEN,
          },
        };

        const gradeSuffixMatch = item.Grade.match(/\d+/);
        const gradeSuffix = gradeSuffixMatch ? gradeSuffixMatch[0] : null;
        const readingLevel = item.LastYearReadingLevel;

        if (gradeSuffix) {
          if (readingColors[gradeSuffix]) {
            const color =
              readingColors[gradeSuffix][readingLevel] ||
              (readingLevel.includes("A") && readingLevel.includes("B") && readingColors[gradeSuffix]["AB"]) ||
              (readingLevel.includes("B") && readingLevel.includes("C") && readingColors[gradeSuffix]["BC"]) ||
              (readingLevel.includes("C") && readingLevel.includes("D") && readingColors[gradeSuffix]["CD"]) ||
              (readingLevel.includes("D") && readingLevel.includes("E") && readingColors[gradeSuffix]["DE"]) ||
              (readingLevel.includes("A") && readingColors[gradeSuffix]["AA"]) ||
              (readingLevel.includes("B") && readingColors[gradeSuffix]["BB"]) ||
              (readingLevel.includes("C") && readingColors[gradeSuffix]["CC"]) ||
              (readingLevel.includes("D") && readingColors[gradeSuffix]["DD"]) ||
              (readingLevel.includes("E") && readingColors[gradeSuffix]["EE"]);
            if (color) {
              lastYearReadingCell.setBackgroundColor(color);
              lastYearLangCell.setBackgroundColor(color);
            }
          }
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
        const cellText = cell.getText();

        cell.setFontSize(FONT_SIZE).setFontFamily(FONT_FAMILY).setBold(true);

        // Check if the cell text does NOT contain a '%'
        if (!cellText.includes("%")) {
          cell.setBold(false).setForegroundColor(DEFAULT_FONT_COLOR); // Set color only for cells without '%'
        }
      }
    }

    // Remove the placeholder text
    element.asText().deleteText(start, end);

    // Move to the next occurrence of the placeholder
    searchResult = body.findText(placeholder, searchResult);
  }
}
