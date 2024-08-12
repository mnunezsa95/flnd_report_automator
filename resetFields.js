function resetFields(sheet, row, locs) {
  // Iterate through the locations array
  for (let i = 0; i < locs.length; i++) {
    sheet.getRange(row, locs[i]).setValue(""); // Set the value of the specified cell to an empty string
  }
}