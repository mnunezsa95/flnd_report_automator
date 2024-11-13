// Updated on 11/13/24

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Report Helper").addItem("Create Report", "main").addToUi();
}
