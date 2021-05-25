function onOpen() {
  var ui = SpreadsheetApp.getUi()
  // Or DocumentApp or FormApp.
  ui.createMenu('Scripts')
    .addItem('Rename Sheet', 'renameSheet')
    .addItem('Sort Sheets', 'sortSheets')
    .addItem('Auto Resize all Columns', 'autoSizeAllColumns')
    .addToUi()
}

function renameSheet() {
  const sheet = SpreadsheetApp.getActiveSheet()
  const title = getUserNames(sheet).join(', ')
  sheet.setName(title)
}

function autoSizeAllColumns() {
  const sheet = SpreadsheetApp.getActiveSheet()
  sheet.autoResizeColumns(1, sheet.getMaxColumns())
}

function sortSheets() {
  const activeSheet = SpreadsheetApp.getActiveSheet()
  sortSheetsByName()
  activeSheet.activate()
}

// start function on edit
function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit) {
  onEditTransformCell(e)
}
