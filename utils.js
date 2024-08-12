function generateUniqueID() {
  const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let uniqueID = '';
  for (let i = 0; i < 8; i++) {
    const randomIndex = Math.floor(Math.random() * characters.length);
    uniqueID += characters.charAt(randomIndex);
  }
  return uniqueID;
}


function resetFields(sheet, row, locs) {
  for (let i = 0; i < locs.length; i++) {
    sheet.getRange(row, locs[i]).setValue(""); // Set the value of the specified cell to an empty string
  }
}


