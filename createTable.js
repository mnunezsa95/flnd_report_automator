// Function to create a table at a placeholder location in the document
function createTableAtPlaceholder(body, data, headers, isNumeracy, placeholder) {
  // Define number of columns based on whether it's numeracy or literacy
  const numColumns = isNumeracy ? 5 : 6;

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
    headers.forEach(header => {
      const cell = headerRow.appendTableCell(header);
      cell.setFontSize(9);
      cell.setFontFamily('Arial');
      cell.setBackgroundColor('#efefef');
      cell.setBold(true);
    });

    // Add data rows
    data.forEach(item => {
      const row = table.appendTableRow();
      row.appendTableCell(item.Grade || '');         // Populate "Grade" column
      row.appendTableCell(item.Level || '');         // Populate "Previous Levels" column
      
      // Add empty cells for other columns
      row.appendTableCell('');  // on-level Avg
      if (isNumeracy) {
        row.appendTableCell(''); // 2025 Numeracy Level
      } else {
        row.appendTableCell(''); // 2025 Lit Level
        row.appendTableCell(''); // 2025 Lang
      }
      row.appendTableCell('');  // Repeating?
    });

    const rows = table.getNumRows();
    for (let i = 1; i < rows; i++) {
      const row = table.getRow(i);
      const cells = row.getNumCells();
      for (let j = 0; j < cells; j++) {
        const cell = row.getCell(j);
        cell.setFontSize(9);
        cell.setFontFamily('Arial');
        cell.setBold(false);
        cell.setForegroundColor('#666666');
      }
    }

    // Apply specific colors to rows
    const colorPatternsOne = [
      '#c9daf8', '#d9ead3', '#f4cccc', '#d9d2e9', '#fce5cd', '#ffd966', '#f3f3f3', '#ffffff'
    ];
    
    // Apply colors to the first 3 columns
    for (let i = 0; i < 8; i++) {
      if (i < rows - 1) {
        const row = table.getRow(i + 1); // Skip header row
        const color = colorPatternsOne[i];
        for (let j = 0; j < 3; j++) {
          row.getCell(j).setBackgroundColor(color);
        }
      }
    }

    const colorPatternsTwo = [
      '#ffffff', '#c9daf8', '#d9ead3', '#f4cccc', '#d9d2e9', '#fce5cd', '#ffd966', '#f3f3f3',
    ];

    // Apply colors to columns 4 to end
    for (let i = 0; i < 8; i++) {
      if (i < rows - 1) {
        const row = table.getRow(i + 1); // Skip header row
        const color = colorPatternsTwo[i];
        for (let j = 3; j < numColumns; j++) {
          row.getCell(j).setBackgroundColor(color);
        }
      }
    }

    // Remove the placeholder text
    element.asText().deleteText(start, end);
    
    // Move to the next occurrence of the placeholder
    searchResult = body.findText(placeholder, searchResult);
  }
}

